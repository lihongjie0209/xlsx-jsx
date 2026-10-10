use std::path::PathBuf;

pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("input: {0}")]
    Input(String),
    #[error("transpile: {0}")]
    Transpile(String),
    #[error("module: {0}")]
    Module(String),
    #[error("runtime: {0}")]
    Runtime(String),
    #[error("validation at {path}: {message}")]
    Validation { path: String, message: String },
    #[error("resource {path}: {source}")]
    Resource {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    #[error("compile: {0}")]
    Compile(String),
    #[error("reverse: {0}")]
    Reverse(String),
    #[error("output {path}: {source}")]
    Output {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
}

impl Error {
    #[must_use]
    pub fn render_diagnostic(&self) -> String {
        let (explanation, suggestion) = match self {
            Self::Input(_) => (
                "The command input or argument is invalid.",
                "Check the path and run `xlsx-jsx --help` for accepted arguments.",
            ),
            Self::Transpile(_) => (
                "The JSX/TSX source could not be converted to JavaScript.",
                "Fix the syntax error at the reported source location.",
            ),
            Self::Module(_) => (
                "A JavaScript module could not be resolved or evaluated.",
                "Use local imports or the built-in `xlsx-jsx` runtime.",
            ),
            Self::Runtime(_) => (
                "The JSX module failed while producing a workbook tree.",
                "Ensure the default export returns one `<Workbook>` tree.",
            ),
            Self::Validation { .. } | Self::Compile(_) => (
                "The workbook tree violates the supported DSL or writer contract.",
                "Check the component and property rules with `xlsx-jsx spec`.",
            ),
            Self::Resource { .. } => (
                "A referenced local file could not be read.",
                "Check the file path and permissions.",
            ),
            Self::Reverse(message) if message.contains("cell formatting") => (
                "The source workbook has styles that the current lossless DSL cannot represent.",
                "Remove cell/row formatting before reverse, or add style components with both writer and external-reader support.",
            ),
            Self::Reverse(message) if message.contains("package part") => (
                "The workbook contains an extra feature part that would be lost during reverse.",
                "Remove the named feature from the workbook or extend the JSX DSL and reverse reader together.",
            ),
            Self::Reverse(message) if message.contains("hidden worksheets") => (
                "Worksheet visibility is not represented by the v1 DSL.",
                "Make every worksheet visible before reversing, or add the visibility property with round-trip tests.",
            ),
            Self::Reverse(message) if message.contains("named ranges") => (
                "Named ranges and print areas are not represented by the v1 DSL.",
                "Remove the defined names or add a matching workbook-level JSX component and reverse support.",
            ),
            Self::Reverse(message) if message.contains("date/time") => (
                "The source cell is a spreadsheet date or time, not an ordinary number.",
                "Convert it to ISO 8601 text in the source workbook or add a typed date component with round-trip support.",
            ),
            Self::Reverse(message) if message.contains("duration") => (
                "The source cell is an ISO duration that the current DSL cannot express.",
                "Convert it to plain text or add a duration component with matching writer and reader behavior.",
            ),
            Self::Reverse(message) if message.contains("Excel error") => (
                "The source cell contains an Excel error value rather than a scalar result.",
                "Fix the formula/error or add an explicit error-cell type before reversing.",
            ),
            Self::Reverse(message) if message.contains("rich-text") => (
                "The source workbook contains rich-string features outside the lossless TextRun subset.",
                "Remove phonetic annotations or unsupported run font metadata, or extend TextRun and its reverse parser together; supported rich-text runs already round-trip.",
            ),
            Self::Reverse(_) => (
                "The XLSX file uses content outside the lossless v1 reverse subset.",
                "Remove unsupported workbook features or convert values to supported scalar cells.",
            ),
            Self::Output { .. } => (
                "The requested output could not be written.",
                "Check the destination directory and use `--force` only when replacement is intended.",
            ),
        };
        format!(
            "error: {self}\nexplanation: {explanation}\nsuggestion: {suggestion}\nspec: run `xlsx-jsx spec` for the complete contract"
        )
    }
}
