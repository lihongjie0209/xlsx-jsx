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
            Self::Reverse(message) if message.contains("not a readable XLSX ZIP package") => (
                "The input is not a readable XLSX package, so workbook features could not be inspected.",
                "Save or export the source as a valid `.xlsx` file, then retry; do not rename a CSV or other format to `.xlsx`.",
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
            Self::Reverse(message) if message.contains("unsupported Excel error") => (
                "The source cell uses an Excel error token outside the lossless v1 subset.",
                "Replace it with one of the supported error tokens shown above, or repair the source cell before reversing.",
            ),
            Self::Reverse(message) if message.contains("Excel error") => (
                "The source cell contains an Excel error value rather than a scalar result.",
                "Fix the formula/error or add an explicit error-cell type before reversing.",
            ),
            Self::Reverse(message) if message.contains("rich-text") => (
                "The source workbook contains rich-string features outside the lossless TextRun subset.",
                "Remove phonetic annotations or unsupported run font metadata, or extend TextRun and its reverse parser together; supported rich-text runs already round-trip.",
            ),
            Self::Reverse(message) if message.contains("array formula cached result") => (
                "Legacy array-formula cached values cannot be represented independently from their formula range.",
                "Keep only writer-default zero caches for legacy array formulas; scalar `formulaResult` applies to ordinary formulas, not legacy array caches.",
            ),
            Self::Reverse(message) if message.contains("shared formula") => (
                "The external workbook's shared-formula group is malformed or incomplete; shared formulas can only be expanded when their group metadata is consistent.",
                "Restore one anchor with a valid `ref`, give every member the same numeric `si`, and include a formula cell for every cell in the declared range.",
            ),
            Self::Reverse(message) if message.contains("data-table") => (
                "The source workbook contains an Excel data-table formula, which the current DSL cannot represent.",
                "Replace the data-table formula with ordinary formulas before reversing, or keep and use the original workbook unchanged.",
            ),
            Self::Reverse(message) if message.contains("defaultColWidth") => (
                "The source worksheet's default column width is outside the exact pixel widths supported by the pinned writer.",
                "Set the default to 56, 64, 72, 80, 96, 104, or 120 pixels, or restore Excel's 8.43-character default; the reverse reader does not round widths.",
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

#[cfg(test)]
mod tests {
    use super::Error;

    #[test]
    fn unsupported_excel_error_diagnostic_lists_supported_tokens() {
        let diagnostic = Error::Reverse(
            "cell `A1` on Worksheet `Sheet1` contains unsupported Excel error `#CALC!`; use one of #NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A or repair the source cell".into(),
        )
        .render_diagnostic();

        assert!(diagnostic.contains("#CALC!"));
        assert!(diagnostic.contains("#N/A"));
        assert!(diagnostic.contains("repair the source cell"));
        assert!(!diagnostic.contains("add an explicit error-cell type"));
    }

    #[test]
    fn malformed_shared_formula_diagnostic_explains_group_repair() {
        let diagnostic = Error::Reverse(
            "worksheet `xl/worksheets/sheet1.xml` shared formula group `si=0` covers 1 formula cells but its range requires 2; restore every formula cell in the group".into(),
        )
        .render_diagnostic();

        assert!(diagnostic.contains("shared-formula group is malformed or incomplete"));
        assert!(diagnostic.contains("Restore one anchor with a valid `ref`"));
        assert!(!diagnostic.contains("Remove unsupported workbook features"));
    }

    #[test]
    fn data_table_formula_diagnostic_names_the_unsupported_feature() {
        let diagnostic = Error::Reverse(
            "worksheet `xl/worksheets/sheet1.xml` contains a data-table or unsupported formula type; convert it to an ordinary formula or a legacy array formula".into(),
        )
        .render_diagnostic();

        assert!(diagnostic.contains("Excel data-table formula"));
        assert!(diagnostic.contains("Replace the data-table formula with ordinary formulas"));
        assert!(!diagnostic.contains("Remove unsupported workbook features"));
    }

    #[test]
    fn unsupported_default_column_width_diagnostic_lists_exact_choices() {
        let diagnostic = Error::Reverse(
            "worksheet `xl/worksheets/sheet1.xml` defaultColWidth `12.25` is not exactly representable by the supported 56/64/72/80/96/104/120-pixel defaults; use one of those widths or restore Excel's 8.43-character default".into(),
        )
        .render_diagnostic();

        assert!(diagnostic.contains("outside the exact pixel widths supported"));
        assert!(diagnostic.contains("56, 64, 72, 80, 96, 104, or 120 pixels"));
        assert!(!diagnostic.contains("Remove unsupported workbook features"));
    }
}
