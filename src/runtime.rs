use std::borrow::Cow;
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use std::rc::Rc;

use deno_ast::{
    DecoratorsTranspileOption, EmitOptions, ImportsNotUsedAsValues, JsxAutomaticOptions,
    JsxRuntime, MediaType, ParseParams, SourceMapOption, TranspileModuleOptions, TranspileOptions,
};
use deno_core::error::ModuleLoaderError;
use deno_core::{
    JsRuntime, ModuleLoadOptions, ModuleLoadReferrer, ModuleLoadResponse, ModuleLoader,
    ModuleResolveResponse, ModuleSource, ModuleSourceCode, ModuleSpecifier, ModuleType,
    PollEventLoopOptions, ResolutionKind, RuntimeOptions, resolve_import, v8,
};
use deno_error::JsErrorBox;
use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::error::{Error, Result};
use crate::ir::IrEnvelope;

const JSX_RUNTIME_SOURCE: &str = include_str!("../runtime/jsx-runtime.js");
const MAIN_MODULE_SOURCE: &str = include_str!("../runtime/main.js");
const EXTENSIONS: [&str; 5] = ["ts", "tsx", "js", "jsx", "json"];
const TRANSPILE_CACHE_FORMAT_VERSION: u32 = 1;

#[derive(Debug, Default)]
struct XlsxModuleLoader {
    cache_dir: Option<PathBuf>,
}

impl ModuleLoader for XlsxModuleLoader {
    fn resolve(
        &self,
        specifier: &str,
        referrer: &str,
        _kind: ResolutionKind,
    ) -> ModuleResolveResponse {
        match specifier {
            "xlsx-jsx" => return parse_builtin("xlsx-jsx:main"),
            "xlsx-jsx:entry" => return parse_builtin("xlsx-jsx:entry"),
            "xlsx-jsx/jsx-runtime" | "xlsx-jsx/jsx-dev-runtime" => {
                return parse_builtin("xlsx-jsx:jsx-runtime");
            }
            _ => {}
        }
        if is_bare_specifier(specifier) {
            return Err(JsErrorBox::type_error(format!(
                "Bare module `{specifier}` is not supported; only xlsx-jsx and local files are allowed"
            )));
        }
        let resolved = resolve_import(specifier, referrer).map_err(JsErrorBox::from_err)?;
        resolve_local_extension(&resolved)
    }

    fn load(
        &self,
        module_specifier: &ModuleSpecifier,
        _maybe_referrer: Option<&ModuleLoadReferrer>,
        _options: ModuleLoadOptions,
    ) -> ModuleLoadResponse {
        let result = load_module(module_specifier, self.cache_dir.as_deref());
        ModuleLoadResponse::Sync(result)
    }

    fn get_source_map(&self, _specifier: &str) -> Option<Cow<'_, [u8]>> {
        None
    }
}

fn parse_builtin(specifier: &str) -> ModuleResolveResponse {
    ModuleSpecifier::parse(specifier).map_err(JsErrorBox::from_err)
}

fn is_bare_specifier(specifier: &str) -> bool {
    !specifier.starts_with('.')
        && !specifier.starts_with('/')
        && !specifier.starts_with("file:")
        && !specifier.contains(':')
}

fn resolve_local_extension(specifier: &ModuleSpecifier) -> ModuleResolveResponse {
    if specifier.scheme() != "file" {
        return Err(JsErrorBox::type_error(format!(
            "Unsupported module scheme `{}`",
            specifier.scheme()
        )));
    }
    let path = specifier
        .to_file_path()
        .map_err(|()| JsErrorBox::type_error("Invalid file module URL"))?;
    let resolved = resolve_candidate(&path)
        .ok_or_else(|| JsErrorBox::type_error(format!("Module not found: {}", path.display())))?;
    ModuleSpecifier::from_file_path(&resolved).map_err(|()| {
        JsErrorBox::type_error(format!("Invalid module path: {}", resolved.display()))
    })
}

fn resolve_candidate(path: &Path) -> Option<PathBuf> {
    if path.is_file() {
        return Some(path.to_path_buf());
    }
    if path.extension().is_none() {
        for extension in EXTENSIONS {
            let candidate = path.with_extension(extension);
            if candidate.is_file() {
                return Some(candidate);
            }
        }
    }
    if path.is_dir() {
        for extension in EXTENSIONS {
            let candidate = path.join(format!("index.{extension}"));
            if candidate.is_file() {
                return Some(candidate);
            }
        }
    }
    None
}

fn load_module(
    specifier: &ModuleSpecifier,
    cache_dir: Option<&Path>,
) -> std::result::Result<ModuleSource, ModuleLoaderError> {
    let source = match specifier.as_str() {
        "xlsx-jsx:main" => MAIN_MODULE_SOURCE.to_owned(),
        "xlsx-jsx:jsx-runtime" => JSX_RUNTIME_SOURCE.to_owned(),
        _ => load_file_module(specifier, cache_dir)?,
    };
    Ok(ModuleSource::new(
        ModuleType::JavaScript,
        ModuleSourceCode::String(source.into()),
        specifier,
        None,
    ))
}

fn load_file_module(
    specifier: &ModuleSpecifier,
    cache_dir: Option<&Path>,
) -> std::result::Result<String, ModuleLoaderError> {
    let path = specifier
        .to_file_path()
        .map_err(|()| JsErrorBox::type_error("Only local file modules are supported"))?;
    let source = std::fs::read_to_string(&path).map_err(JsErrorBox::from_err)?;
    if path.extension().and_then(|value| value.to_str()) == Some("json") {
        serde_json::from_str::<Value>(&source).map_err(JsErrorBox::from_err)?;
        return Ok(format!("export default {source};"));
    }
    let media_type = MediaType::from_path(&path);
    if matches!(media_type, MediaType::JavaScript | MediaType::Mjs) {
        return Ok(source);
    }
    transpile_with_cache(specifier, &source, media_type, cache_dir)
        .map(|(transpiled, _)| transpiled)
}

#[derive(Debug, Deserialize, Serialize)]
struct TranspileCacheEntry {
    format_version: u32,
    package_version: String,
    specifier: String,
    media_type: String,
    source: String,
    transpiled: String,
}

fn transpile_with_cache(
    specifier: &ModuleSpecifier,
    source: &str,
    media_type: MediaType,
    cache_dir: Option<&Path>,
) -> std::result::Result<(String, bool), JsErrorBox> {
    let Some(cache_dir) = cache_dir else {
        return transpile(specifier, source.to_owned(), media_type)
            .map(|transpiled| (transpiled, false))
            .map_err(|error| JsErrorBox::generic(error.to_string()));
    };

    let cache_path = transpile_cache_path(cache_dir, specifier);
    if let Some(transpiled) = read_transpile_cache(&cache_path, specifier, source, media_type)? {
        return Ok((transpiled, true));
    }

    let transpiled = transpile(specifier, source.to_owned(), media_type)
        .map_err(|error| JsErrorBox::generic(error.to_string()))?;
    let entry = TranspileCacheEntry {
        format_version: TRANSPILE_CACHE_FORMAT_VERSION,
        package_version: env!("CARGO_PKG_VERSION").to_owned(),
        specifier: specifier.as_str().to_owned(),
        media_type: format!("{media_type:?}"),
        source: source.to_owned(),
        transpiled: transpiled.clone(),
    };
    let serialized = serde_json::to_vec(&entry).map_err(|error| {
        JsErrorBox::generic(format!("cannot encode transpilation cache: {error}"))
    })?;
    std::fs::create_dir_all(cache_dir).map_err(|error| {
        JsErrorBox::generic(format!(
            "cannot create transpilation cache directory {}: {error}; ensure `--cache-dir` names a writable directory",
            cache_dir.display()
        ))
    })?;
    std::fs::write(&cache_path, serialized).map_err(|error| {
        JsErrorBox::generic(format!(
            "cannot write transpilation cache {}: {error}; ensure `--cache-dir` names a writable directory",
            cache_path.display()
        ))
    })?;
    Ok((transpiled, false))
}

fn transpile_cache_path(cache_dir: &Path, specifier: &ModuleSpecifier) -> PathBuf {
    let mut hasher = DefaultHasher::new();
    TRANSPILE_CACHE_FORMAT_VERSION.hash(&mut hasher);
    env!("CARGO_PKG_VERSION").hash(&mut hasher);
    specifier.as_str().hash(&mut hasher);
    cache_dir.join(format!("{:016x}.json", hasher.finish()))
}

fn read_transpile_cache(
    cache_path: &Path,
    specifier: &ModuleSpecifier,
    source: &str,
    media_type: MediaType,
) -> std::result::Result<Option<String>, JsErrorBox> {
    let serialized = match std::fs::read_to_string(cache_path) {
        Ok(serialized) => serialized,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(error) => {
            return Err(JsErrorBox::generic(format!(
                "cannot read transpilation cache {}: {error}; ensure `--cache-dir` names an accessible directory",
                cache_path.display()
            )));
        }
    };
    let Ok(entry) = serde_json::from_str::<TranspileCacheEntry>(&serialized) else {
        return Ok(None);
    };
    let matches_current_input = entry.format_version == TRANSPILE_CACHE_FORMAT_VERSION
        && entry.package_version == env!("CARGO_PKG_VERSION")
        && entry.specifier == specifier.as_str()
        && entry.media_type == format!("{media_type:?}")
        && entry.source == source;
    Ok(matches_current_input.then_some(entry.transpiled))
}

fn transpile(specifier: &ModuleSpecifier, source: String, media_type: MediaType) -> Result<String> {
    let parsed = deno_ast::parse_module(ParseParams {
        specifier: specifier.clone(),
        text: source.into(),
        media_type,
        capture_tokens: false,
        scope_analysis: false,
        maybe_syntax: None,
    })
    .map_err(|error| Error::Transpile(error.to_string()))?;
    let emitted = parsed
        .transpile(
            &TranspileOptions {
                imports_not_used_as_values: ImportsNotUsedAsValues::Remove,
                decorators: DecoratorsTranspileOption::Ecma,
                jsx: Some(JsxRuntime::Automatic(JsxAutomaticOptions {
                    development: false,
                    import_source: Some("xlsx-jsx".to_owned()),
                })),
                ..Default::default()
            },
            &TranspileModuleOptions::default(),
            &EmitOptions {
                source_map: SourceMapOption::Inline,
                inline_sources: true,
                ..Default::default()
            },
        )
        .map_err(|error| Error::Transpile(error.to_string()))?
        .into_source();
    Ok(emitted.text)
}

/// Evaluates an entry JSX/TSX module and returns normalized IR v1.
///
/// # Errors
///
/// Returns an error when the module graph cannot be loaded or transpiled,
/// JavaScript evaluation fails, or the default export cannot be decoded as IR.
pub async fn evaluate_entry(entry: &Path, data: Option<&Value>) -> Result<IrEnvelope> {
    evaluate_entry_with_cache(entry, data, None).await
}

/// Evaluates a JSX/TSX module and optionally persists per-module transpilation results.
///
/// The module graph is evaluated for every invocation; only unchanged TypeScript/JSX
/// transpilation output is reused from `cache_dir`.
///
/// # Errors
///
/// Returns an error when the module graph cannot be loaded or transpiled, the cache
/// cannot be read or written, JavaScript evaluation fails, or the default export is
/// not a valid workbook IR.
pub async fn evaluate_entry_with_cache(
    entry: &Path,
    data: Option<&Value>,
    cache_dir: Option<&Path>,
) -> Result<IrEnvelope> {
    let canonical_entry = entry
        .canonicalize()
        .map_err(|error| Error::Input(format!("cannot open {}: {error}", entry.display())))?;
    let entry_url = ModuleSpecifier::from_file_path(&canonical_entry)
        .map_err(|()| Error::Input(format!("invalid entry path: {}", canonical_entry.display())))?;
    let wrapper_url = ModuleSpecifier::parse("xlsx-jsx:entry")
        .map_err(|error| Error::Runtime(format!("cannot create wrapper module URL: {error}")))?;
    let data_source = data.map_or_else(|| "undefined".to_owned(), Value::to_string);
    let entry_source = serde_json::to_string(entry_url.as_str())
        .map_err(|error| Error::Runtime(error.to_string()))?;
    let wrapper = format!(
        "import entry from {entry_source};\nimport {{ finalize }} from \"xlsx-jsx/jsx-runtime\";\nexport default await finalize(entry, {data_source});"
    );
    let mut runtime = JsRuntime::new(RuntimeOptions {
        module_loader: Some(Rc::new(XlsxModuleLoader {
            cache_dir: cache_dir.map(Path::to_path_buf),
        })),
        ..Default::default()
    });
    runtime
        .execute_script(
            "xlsx-jsx:bootstrap",
            r#"const __xlsxFormat = values => values.map(value => {
  if (typeof value === "string") return value;
  try { return JSON.stringify(value); } catch { return String(value); }
}).join(" ");
globalThis.console ??= {
  log(...values) { Deno.core.print(__xlsxFormat(values) + "\n", true); },
  info(...values) { Deno.core.print(__xlsxFormat(values) + "\n", true); },
  warn(...values) { Deno.core.print(__xlsxFormat(values) + "\n", true); },
  error(...values) { Deno.core.print(__xlsxFormat(values) + "\n", true); }
};"#,
        )
        .map_err(|error| Error::Runtime(error.to_string()))?;
    let module_id = runtime
        .load_main_es_module_from_code(&wrapper_url, wrapper)
        .await
        .map_err(|error| Error::Module(error.to_string()))?;
    let evaluation = runtime.mod_evaluate(module_id);
    runtime
        .run_event_loop(PollEventLoopOptions::default())
        .await
        .map_err(|error| Error::Runtime(error.to_string()))?;
    evaluation
        .await
        .map_err(|error| Error::Runtime(error.to_string()))?;
    let namespace = runtime
        .get_module_namespace(module_id)
        .map_err(|error| Error::Runtime(error.to_string()))?;
    let value = {
        deno_core::scope!(scope, runtime);
        let namespace = v8::Local::<v8::Object>::new(scope, namespace);
        let key = v8::String::new(scope, "default")
            .ok_or_else(|| Error::Runtime("cannot allocate export key".to_owned()))?;
        let export = namespace
            .get(scope, key.into())
            .ok_or_else(|| Error::Runtime("entry wrapper has no default export".to_owned()))?;
        deno_core::serde_v8::from_v8::<Value>(scope, export)
            .map_err(|error| Error::Runtime(format!("cannot decode JSX result: {error}")))?
    };
    let ir: IrEnvelope = serde_json::from_value(value)
        .map_err(|error| Error::Runtime(format!("invalid IR from JSX runtime: {error}")))?;
    ir.validate()?;
    Ok(ir)
}

#[cfg(test)]
mod tests {
    use std::fs;

    use tempfile::tempdir;

    use super::*;

    #[tokio::test(flavor = "current_thread")]
    async fn evaluate_should_execute_tsx_components_and_data() {
        let directory = tempdir().expect("tempdir should work");
        let entry = directory.path().join("report.tsx");
        fs::write(
            &entry,
            r#"import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";
type Data = { names: string[] };
const Item = ({ name }: { name: string }) => <Row><Cell>{name}</Cell></Row>;
export default (data: Data) => <Workbook><Worksheet name="Names">{data.names.map(name => <Item name={name} />)}</Worksheet></Workbook>;"#,
        )
        .expect("fixture should write");
        let data = serde_json::json!({"names": ["Ada", "Linus"]});
        let ir = evaluate_entry(&entry, Some(&data))
            .await
            .expect("evaluation should work");
        assert_eq!(ir.workbook.children.len(), 1);
    }

    #[tokio::test(flavor = "current_thread")]
    async fn evaluate_should_load_extensionless_local_module() {
        let directory = tempdir().expect("tempdir should work");
        fs::write(
            directory.path().join("component.tsx"),
            r#"import { Row, Cell } from "xlsx-jsx"; export const Greeting = () => <Row><Cell>Hello</Cell></Row>;"#,
        )
        .expect("component should write");
        let entry = directory.path().join("report.tsx");
        fs::write(
            &entry,
            r#"import { Workbook, Worksheet } from "xlsx-jsx"; import { Greeting } from "./component"; export default <Workbook><Worksheet name="Sheet"><Greeting /></Worksheet></Workbook>;"#,
        )
        .expect("entry should write");
        assert!(evaluate_entry(&entry, None).await.is_ok());
    }

    #[test]
    fn transpilation_cache_hits_for_unchanged_source() {
        let directory = tempdir().expect("tempdir should work");
        let module = directory.path().join("component.tsx");
        let specifier = ModuleSpecifier::from_file_path(&module).expect("module URL is valid");
        let cache = directory.path().join("cache");
        let source = "export const label: string = 'first';";

        let (_, cache_hit) = transpile_with_cache(&specifier, source, MediaType::Tsx, Some(&cache))
            .expect("first transpilation succeeds");
        assert!(!cache_hit, "first compile should miss");

        let (_, cache_hit) = transpile_with_cache(&specifier, source, MediaType::Tsx, Some(&cache))
            .expect("cached transpilation succeeds");
        assert!(cache_hit, "unchanged source should hit");
    }

    #[test]
    fn transpilation_cache_rebuilds_when_a_module_changes() {
        let directory = tempdir().expect("tempdir should work");
        let module = directory.path().join("component.ts");
        let specifier = ModuleSpecifier::from_file_path(&module).expect("module URL is valid");
        let cache = directory.path().join("cache");
        let media_type = MediaType::TypeScript;
        let _ = transpile_with_cache(
            &specifier,
            "export const value = 1;",
            media_type,
            Some(&cache),
        )
        .expect("initial transpilation succeeds");

        let (transpiled, cache_hit) = transpile_with_cache(
            &specifier,
            "export const value = 2;",
            media_type,
            Some(&cache),
        )
        .expect("changed module transpiles");
        assert!(!cache_hit, "changed source must miss the cache");
        assert!(transpiled.contains("value = 2"), "{transpiled}");
    }

    #[test]
    fn transpilation_cache_recovers_from_corrupt_entries() {
        let directory = tempdir().expect("tempdir should work");
        let module = directory.path().join("component.ts");
        let specifier = ModuleSpecifier::from_file_path(&module).expect("module URL is valid");
        let cache = directory.path().join("cache");
        std::fs::create_dir_all(&cache).expect("cache directory creates");
        let cache_path = transpile_cache_path(&cache, &specifier);
        std::fs::write(&cache_path, "not valid cache JSON").expect("bad cache writes");

        let (_, cache_hit) = transpile_with_cache(
            &specifier,
            "export const value = 1;",
            MediaType::TypeScript,
            Some(&cache),
        )
        .expect("corrupt cache is regenerated");
        assert!(!cache_hit, "corrupt entry should be a miss");
    }

    #[test]
    fn transpilation_cache_reports_unwritable_cache_directory() {
        let directory = tempdir().expect("tempdir should work");
        let module = directory.path().join("component.ts");
        let specifier = ModuleSpecifier::from_file_path(&module).expect("module URL is valid");
        let cache_file = directory.path().join("not-a-directory");
        std::fs::write(&cache_file, "occupied").expect("cache path blocker writes");

        let error = transpile_with_cache(
            &specifier,
            "export const value = 1;",
            MediaType::TypeScript,
            Some(&cache_file),
        )
        .expect_err("cache path must be a directory");
        assert!(error.to_string().contains("--cache-dir"), "{error}");
    }
}
