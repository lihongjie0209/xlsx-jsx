use std::collections::{BTreeMap, HashMap};
use std::fmt::Write as FmtWrite;
use std::path::PathBuf;
use std::process::ExitCode;

use base64::{Engine, engine::general_purpose::STANDARD as BASE64};
use calamine::{Data, DataType, Reader, open_workbook_auto};
use clap::{Parser, Subcommand};
use quick_xml::{
    Reader as XmlReader, Writer as XmlWriter, XmlVersion,
    events::{BytesStart, Event},
};
use rust_xlsxwriter::{
    Chart, ChartLegendPosition, ChartType, Color, ConditionalFormat2ColorScale,
    ConditionalFormat3ColorScale, ConditionalFormatAverage, ConditionalFormatAverageRule,
    ConditionalFormatBlank, ConditionalFormatCell, ConditionalFormatCellRule,
    ConditionalFormatCustomIcon, ConditionalFormatDataBar, ConditionalFormatDataBarAxisPosition,
    ConditionalFormatDataBarDirection, ConditionalFormatDate, ConditionalFormatDateRule,
    ConditionalFormatDuplicate, ConditionalFormatError, ConditionalFormatFormula,
    ConditionalFormatIconSet, ConditionalFormatIconType, ConditionalFormatText,
    ConditionalFormatTextRule, ConditionalFormatTop, ConditionalFormatTopRule,
    ConditionalFormatType, ConditionalFormatValue, DataValidation, DataValidationErrorStyle,
    DataValidationRule, DocProperties, FilterCondition, FilterCriteria, Format, FormatAlign,
    FormatBorder, FormatDiagonalBorder, FormatPattern, FormatScript, FormatUnderline, Formula,
    IgnoreError, Image as XlsxImage, Note, ObjectMovement, Table, TableColumn, TableStyle, Url,
    Workbook,
};
use serde_json::Value;
use xlsx_jsx::{
    Error, IrEnvelope, Node, OutlineProperties, PrintRange, collapsed_outline_ranges,
    format_a1_cell_reference, format_a1_range, format_print_range, parse_a1_cell_reference,
    parse_a1_range, parse_print_range, runtime::evaluate_entry,
};
use zip::ZipArchive;

const FORMAT_PROPERTY_NAMES: &[&str] = &[
    "bold",
    "italic",
    "horizontalAlign",
    "verticalAlign",
    "indent",
    "wrapText",
    "shrinkToFit",
    "textRotation",
    "readingDirection",
    "quotePrefix",
    "locked",
    "formulaHidden",
    "numberFormat",
    "fillColor",
    "fillPattern",
    "fillForegroundColor",
    "fillBackgroundColor",
    "fontColor",
    "fontName",
    "fontSize",
    "fontFamily",
    "fontCharset",
    "underline",
    "strikethrough",
    "script",
    "borderTop",
    "borderRight",
    "borderBottom",
    "borderLeft",
    "borderTopColor",
    "borderRightColor",
    "borderBottomColor",
    "borderLeftColor",
    "borderDiagonal",
    "borderDiagonalColor",
    "borderDiagonalType",
];

#[derive(Debug, Parser)]
#[command(
    name = "xlsx-jsx",
    version,
    about = "Compile JSX/TSX workbooks to XLSX and reverse supported XLSX files"
)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Debug, Subcommand)]
enum Command {
    /// Compile an executable JSX/TSX workbook.
    Build {
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
        #[arg(long)]
        data: Option<PathBuf>,
        #[arg(long)]
        force: bool,
    },
    /// Validate JSX/TSX without writing an XLSX file.
    Validate {
        input: PathBuf,
        #[arg(long)]
        data: Option<PathBuf>,
    },
    /// Reverse a supported XLSX file to JSX.
    Reverse {
        input: PathBuf,
        #[arg(short, long)]
        output: Option<PathBuf>,
        #[arg(long)]
        force: bool,
    },
    /// Print the machine-readable feature specification.
    Spec {
        #[arg(long, default_value = "markdown")]
        format: String,
    },
}

#[tokio::main(flavor = "current_thread")]
async fn main() -> ExitCode {
    match run(Cli::parse()).await {
        Ok(Some(path)) => {
            println!("{}", path.display());
            ExitCode::SUCCESS
        }
        Ok(None) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("{}", error.render_diagnostic());
            ExitCode::FAILURE
        }
    }
}

async fn run(cli: Cli) -> xlsx_jsx::Result<Option<PathBuf>> {
    match cli.command {
        Command::Spec { format } => {
            let spec = match format.as_str() {
                "markdown" => include_str!("../docs/spec.md"),
                "json" => include_str!("../spec/ir-v1.schema.json"),
                _ => return Err(Error::Input(format!("unknown spec format `{format}`"))),
            };
            print!("{spec}");
            Ok(None)
        }
        Command::Validate { input, data } => {
            evaluate(&input, data.as_deref()).await?;
            println!("valid");
            Ok(None)
        }
        Command::Build {
            input,
            output,
            data,
            force,
        } => {
            require_extension(&output, "xlsx")?;
            ensure_output(&output, force)?;
            let ir = evaluate(&input, data.as_deref()).await?;
            compile(&ir, &output)?;
            Ok(Some(output))
        }
        Command::Reverse {
            input,
            output,
            force,
        } => {
            require_extension(&input, "xlsx")?;
            let output = output.unwrap_or_else(|| input.with_extension("tsx"));
            ensure_output(&output, force)?;
            let source = reverse(&input)?;
            std::fs::write(&output, source).map_err(|source| Error::Output {
                path: output.clone(),
                source,
            })?;
            Ok(Some(output))
        }
    }
}

async fn evaluate(
    input: &std::path::Path,
    data: Option<&std::path::Path>,
) -> xlsx_jsx::Result<IrEnvelope> {
    let value = data
        .map(|path| {
            std::fs::read_to_string(path).map_err(|source| Error::Resource {
                path: path.to_path_buf(),
                source,
            })
        })
        .transpose()?
        .map(|text| {
            serde_json::from_str::<Value>(&text)
                .map_err(|error| Error::Input(format!("invalid JSON data: {error}")))
        })
        .transpose()?;
    evaluate_entry(input, value.as_ref()).await
}

fn ensure_output(output: &std::path::Path, force: bool) -> xlsx_jsx::Result<()> {
    if output.exists() && !force {
        return Err(Error::Input(format!(
            "output {} already exists",
            output.display()
        )));
    }
    Ok(())
}

fn require_extension(path: &std::path::Path, expected: &str) -> xlsx_jsx::Result<()> {
    if !path
        .extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case(expected))
    {
        return Err(Error::Input(format!(
            "expected a .{expected} path, got {}",
            path.display()
        )));
    }
    Ok(())
}

#[allow(clippy::too_many_lines)] // Worksheet serialization keeps row and column ordering together.
fn compile(ir: &IrEnvelope, path: &std::path::Path) -> xlsx_jsx::Result<()> {
    let mut workbook = Workbook::new();
    let mut properties = DocProperties::new();
    if let Some(value) = ir.workbook.props.get("title").and_then(Value::as_str) {
        properties = properties.set_title(value);
    }
    if let Some(value) = ir.workbook.props.get("subject").and_then(Value::as_str) {
        properties = properties.set_subject(value);
    }
    if let Some(value) = ir.workbook.props.get("author").and_then(Value::as_str) {
        properties = properties.set_author(value);
    }
    if let Some(value) = ir.workbook.props.get("keywords").and_then(Value::as_str) {
        properties = properties.set_keywords(value);
    }
    if let Some(value) = ir.workbook.props.get("description").and_then(Value::as_str) {
        properties = properties.set_comment(value);
    }
    workbook.set_properties(&properties);
    let mut local_defined_names = Vec::new();
    for value in &ir.workbook.children {
        let node: Node = serde_json::from_value(value.clone())
            .map_err(|error| Error::Compile(error.to_string()))?;
        if node.r#type == "DefinedName" {
            let name = node.props["name"].as_str().expect("validated name");
            let formula = node.props["formula"].as_str().expect("validated formula");
            workbook
                .define_name(name, formula.strip_prefix('=').unwrap_or(formula))
                .map_err(|error| {
                    Error::Compile(format!(
                        "DefinedName `{name}`: {error}; use a valid Excel name and formula"
                    ))
                })?;
        }
    }
    for sheet_value in &ir.workbook.children {
        let sheet: Node = serde_json::from_value(sheet_value.clone())
            .map_err(|error| Error::Compile(error.to_string()))?;
        if sheet.r#type == "DefinedName" {
            continue;
        }
        let name = sheet
            .props
            .get("name")
            .and_then(Value::as_str)
            .ok_or_else(|| Error::Compile("worksheet missing name".into()))?;
        for child_value in &sheet.children {
            let child: Node = serde_json::from_value(child_value.clone())
                .map_err(|error| Error::Compile(error.to_string()))?;
            if child.r#type == "DefinedName" {
                local_defined_names.push((
                    name.to_owned(),
                    child.props["name"]
                        .as_str()
                        .expect("validated name")
                        .to_owned(),
                    child.props["formula"]
                        .as_str()
                        .expect("validated formula")
                        .strip_prefix('=')
                        .unwrap_or(child.props["formula"].as_str().expect("validated formula"))
                        .to_owned(),
                ));
            }
        }
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name(name)
            .map_err(|error| Error::Compile(error.to_string()))?;
        if sheet.props.get("active").and_then(Value::as_bool) == Some(true) {
            worksheet.set_active(true);
        }
        if let Some(selected) = sheet.props.get("selected").and_then(Value::as_bool) {
            worksheet.set_selected(selected);
        }
        if sheet.props.get("firstTab").and_then(Value::as_bool) == Some(true) {
            worksheet.set_first_tab(true);
        }
        apply_worksheet_visibility(worksheet, &sheet)?;
        apply_worksheet_tab_color(worksheet, &sheet)?;
        apply_default_row_height(worksheet, &sheet)?;
        apply_hide_unused_rows(worksheet, &sheet)?;
        apply_show_gridlines(worksheet, &sheet)?;
        apply_worksheet_reading_direction(worksheet, &sheet);
        apply_print_gridlines(worksheet, &sheet)?;
        apply_worksheet_view(worksheet, &sheet);
        apply_worksheet_zoom(worksheet, &sheet)?;
        apply_worksheet_page_breaks(worksheet, &sheet)?;
        apply_worksheet_page_margins(worksheet, &sheet)?;
        apply_worksheet_print_settings(worksheet, &sheet)?;
        apply_worksheet_header_footer(worksheet, &sheet);
        apply_worksheet_protection(worksheet, &sheet);
        apply_outline_summary_directions(worksheet, &sheet);
        apply_freeze_panes(worksheet, &sheet)?;
        apply_worksheet_selection(worksheet, &sheet)?;
        apply_auto_filter(worksheet, &sheet)?;
        let mut row_index = 0_usize;
        let mut row_outline = Vec::new();
        let mut column_outline = BTreeMap::new();
        let mut merged_ranges: Vec<(u32, u16, u32, u16)> = Vec::new();
        let mut tables = Vec::new();
        for child_value in &sheet.children {
            let child: Node = serde_json::from_value(child_value.clone())
                .map_err(|error| Error::Compile(error.to_string()))?;
            if child.r#type == "DefinedName" {
                continue;
            }
            if child.r#type == "Note" {
                compile_note(worksheet, &child)?;
                continue;
            }
            if child.r#type == "Chart" {
                compile_chart(worksheet, &child)?;
                continue;
            }
            if child.r#type == "Image" {
                compile_image(worksheet, &child)?;
                continue;
            }
            if child.r#type == "Column" {
                let min = u16::try_from(child.props["min"].as_u64().unwrap_or(1))
                    .map_err(|error| Error::Compile(format!("invalid Column.min: {error}")))?;
                let max = u16::try_from(child.props["max"].as_u64().unwrap_or(u64::from(min)))
                    .map_err(|error| Error::Compile(format!("invalid Column.max: {error}")))?;
                let width = child.props.get("widthPixels").and_then(Value::as_u64);
                let width_pixels = width.map(u32::try_from).transpose().map_err(|error| {
                    Error::Compile(format!("invalid Column.widthPixels: {error}"))
                })?;
                let hidden = child.props.get("hidden").and_then(Value::as_bool) == Some(true);
                let outline_level = child
                    .props
                    .get("outlineLevel")
                    .and_then(Value::as_u64)
                    .unwrap_or_default();
                for column_number in min..=max {
                    let col = column_number - 1;
                    if let Some(width_pixels) = width_pixels {
                        worksheet
                            .set_column_width_pixels(col, width_pixels)
                            .map_err(|error| Error::Compile(error.to_string()))?;
                    }
                    if hidden {
                        worksheet
                            .set_column_hidden(col)
                            .map_err(|error| Error::Compile(error.to_string()))?;
                    }
                    if has_format_properties(&child) {
                        let format = build_cell_format(&child)?;
                        worksheet
                            .set_column_format(col, &format)
                            .map_err(|error| Error::Compile(error.to_string()))?;
                    }
                    column_outline.insert(
                        col,
                        ColumnProperties {
                            width_pixels,
                            hidden,
                            outline_level: u8::try_from(outline_level).unwrap_or_default(),
                            collapsed: child.props.get("collapsed").and_then(Value::as_bool)
                                == Some(true),
                            style_index: None,
                        },
                    );
                }
                continue;
            }
            if child.r#type == "DataValidation" {
                compile_data_validation(worksheet, &child)?;
                continue;
            }
            if child.r#type == "ConditionalFormat" {
                compile_conditional_format(worksheet, &child)?;
                continue;
            }
            if child.r#type == "IgnoreError" {
                compile_ignored_error(worksheet, &child)?;
                continue;
            }
            if child.r#type == "Table" {
                tables.push(child);
                continue;
            }
            compile_row(worksheet, &child, row_index, &mut merged_ranges)?;
            row_outline.push(OutlineProperties {
                hidden: child.props.get("hidden").and_then(Value::as_bool) == Some(true),
                level: child
                    .props
                    .get("outlineLevel")
                    .and_then(Value::as_u64)
                    .and_then(|level| u8::try_from(level).ok())
                    .unwrap_or_default(),
                collapsed: child.props.get("collapsed").and_then(Value::as_bool) == Some(true),
            });
            row_index += 1;
        }
        for table in &tables {
            compile_table(worksheet, &sheet, table)?;
        }
        let row_outline = row_outline
            .into_iter()
            .enumerate()
            .map(|(row, properties)| {
                let row = u32::try_from(row)
                    .map_err(|error| Error::Compile(format!("invalid row index: {error}")))?;
                Ok((row, properties))
            })
            .collect::<xlsx_jsx::Result<BTreeMap<_, _>>>()?;
        let row_groups = collapsed_outline_ranges(&row_outline, "row").map_err(Error::Compile)?;
        for &(first_row, last_row) in &row_groups {
            worksheet
                .group_rows_collapsed(first_row, last_row)
                .map_err(|error| Error::Compile(error.to_string()))?;
        }
        for (&row, properties) in &row_outline {
            let collapsed_depth = row_groups
                .iter()
                .filter(|(first, last)| *first <= row && row <= *last)
                .count();
            let expanded_depth = usize::from(properties.level)
                .checked_sub(collapsed_depth)
                .ok_or_else(|| Error::Compile(format!(
                    "row {} belongs to more collapsed groups than its outline level; raise outlineLevel or simplify the group nesting",
                    row + 1
                )))?;
            for _ in 0..expanded_depth {
                worksheet
                    .group_rows(row, row)
                    .map_err(|error| Error::Compile(error.to_string()))?;
            }
        }
        let column_outline_for_groups = column_outline
            .iter()
            .map(|(&column, properties)| {
                (
                    u32::from(column),
                    OutlineProperties {
                        hidden: properties.hidden,
                        level: properties.outline_level,
                        collapsed: properties.collapsed,
                    },
                )
            })
            .collect::<BTreeMap<_, _>>();
        let column_groups = collapsed_outline_ranges(&column_outline_for_groups, "column")
            .map_err(Error::Compile)?;
        for &(first_column, last_column) in &column_groups {
            let first_column = u16::try_from(first_column)
                .map_err(|error| Error::Compile(format!("invalid collapsed column: {error}")))?;
            let last_column = u16::try_from(last_column)
                .map_err(|error| Error::Compile(format!("invalid collapsed column: {error}")))?;
            worksheet
                .group_columns_collapsed(first_column, last_column)
                .map_err(|error| Error::Compile(error.to_string()))?;
        }
        for (&column, properties) in &column_outline {
            let coordinate = u32::from(column);
            let collapsed_depth = column_groups
                .iter()
                .filter(|(first, last)| *first <= coordinate && coordinate <= *last)
                .count();
            let expanded_depth = usize::from(properties.outline_level)
                .checked_sub(collapsed_depth)
                .ok_or_else(|| Error::Compile(format!(
                    "column {} belongs to more collapsed groups than its outline level; raise outlineLevel or simplify the group nesting",
                    coordinate + 1
                )))?;
            for _ in 0..expanded_depth {
                worksheet
                    .group_columns(column, column)
                    .map_err(|error| Error::Compile(error.to_string()))?;
            }
        }
    }
    for (sheet_name, name, formula) in local_defined_names {
        let quoted_sheet = sheet_name.replace('\'', "''");
        workbook
            .define_name(format!("'{quoted_sheet}'!{name}"), &formula)
            .map_err(|error| Error::Compile(format!("worksheet-local DefinedName `{name}` on `{sheet_name}`: {error}; use a valid Excel name and formula")))?;
    }
    workbook.save(path).map_err(|error| Error::Output {
        path: path.to_path_buf(),
        source: std::io::Error::other(error.to_string()),
    })
}

fn compile_table(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
    node: &Node,
) -> xlsx_jsx::Result<()> {
    let range_text = node.props["range"]
        .as_str()
        .ok_or_else(|| Error::Compile("Table.range must be a string".into()))?;
    let range = parse_a1_range(range_text)
        .ok_or_else(|| Error::Compile(format!("invalid Table.range `{range_text}`")))?;
    let columns = build_table_columns(sheet, node, range)?;
    let name = node.props["name"]
        .as_str()
        .ok_or_else(|| Error::Compile("Table.name must be a string".into()))?;
    let mut table = Table::new().set_name(name).set_columns(&columns);
    if node.props.get("totalsRow").and_then(Value::as_bool) == Some(true) {
        table = table.set_total_row(true);
    }
    table = table.set_style(table_style(
        node.props
            .get("style")
            .and_then(Value::as_str)
            .unwrap_or("TableStyleMedium9"),
    )?);
    table = apply_table_accessibility_text(table, &node.props);
    for (key, setter) in [
        ("autoFilter", 0),
        ("bandedRows", 1),
        ("bandedColumns", 2),
        ("firstColumn", 3),
        ("lastColumn", 4),
    ] {
        if let Some(value) = node.props.get(key).and_then(Value::as_bool) {
            table = match setter {
                0 => table.set_autofilter(value),
                1 => table.set_banded_rows(value),
                2 => table.set_banded_columns(value),
                3 => table.set_first_column(value),
                _ => table.set_last_column(value),
            };
        }
    }
    worksheet
        .add_table(range.0, range.1, range.2, range.3, &table)
        .map_err(|error| Error::Compile(format!("cannot create Table `{name}`: {error}")))?;
    Ok(())
}

fn build_table_columns(
    sheet: &Node,
    node: &Node,
    range: (u32, u16, u32, u16),
) -> xlsx_jsx::Result<Vec<TableColumn>> {
    let row_value = sheet
        .children
        .iter()
        .filter(|value| value.get("type").and_then(Value::as_str) == Some("Row"))
        .nth(range.0 as usize)
        .ok_or_else(|| Error::Compile("Table header row is missing".into()))?;
    let row: Node = serde_json::from_value(row_value.clone())
        .map_err(|error| Error::Compile(error.to_string()))?;
    let mut columns = Vec::new();
    for col in range.1..=range.3 {
        let cell: Node = serde_json::from_value(row.children[usize::from(col)].clone())
            .map_err(|error| Error::Compile(error.to_string()))?;
        let header = cell
            .props
            .get("value")
            .and_then(Value::as_str)
            .or_else(|| cell.children.first().and_then(Value::as_str))
            .ok_or_else(|| Error::Compile("table header must be a string".into()))?;
        let column_node = node.children.get(usize::from(col - range.1));
        let mut column = TableColumn::new().set_header(header);
        if let Some(column_node) = column_node {
            let props = column_node.get("props").and_then(Value::as_object);
            if let Some(label) = props
                .and_then(|props| props.get("totalLabel"))
                .and_then(Value::as_str)
            {
                column = column.set_total_label(label);
            }
            if let Some(format) = props.and_then(|props| props.get("format")) {
                column = column.set_format(build_table_column_format(format)?);
            }
            if let Some(formula) = props
                .and_then(|props| props.get("formula"))
                .and_then(Value::as_str)
            {
                column = column.set_formula(formula.trim_start_matches('='));
            }
            if let Some(formula) = props
                .and_then(|props| props.get("totalFormula"))
                .and_then(Value::as_str)
            {
                column = column.set_total_function(rust_xlsxwriter::TableFunction::Custom(
                    Formula::new(formula.trim_start_matches('=')),
                ));
            } else if let Some(function) = props
                .and_then(|props| props.get("totalFunction"))
                .and_then(Value::as_str)
            {
                column = column.set_total_function(table_function(function)?);
            }
        }
        let header_format = column_node
            .and_then(|column| column.get("props"))
            .and_then(|props| props.get("headerFormat"));
        if let Some(format) = header_format {
            column = column.set_header_format(build_table_column_format(format)?);
        } else if has_format_properties(&cell) {
            column = column.set_header_format(build_cell_format(&cell)?);
        }
        columns.push(column);
    }
    Ok(columns)
}

fn apply_table_accessibility_text(
    mut table: Table,
    props: &serde_json::Map<String, Value>,
) -> Table {
    if let Some(text) = props.get("altText").and_then(Value::as_str) {
        table = table.set_alt_text(text);
    }
    if let Some(title) = props.get("altTextTitle").and_then(Value::as_str) {
        table = table.set_alt_text_title(title);
    }
    table
}

fn compile_note(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    component: &Node,
) -> xlsx_jsx::Result<()> {
    let cell = component.props["cell"]
        .as_str()
        .expect("validated Note.cell");
    let (row, column) = parse_a1_cell_reference(cell).expect("validated Note.cell");
    let text = component.props["text"]
        .as_str()
        .expect("validated Note.text");
    let mut compiled_note = Note::new(text);
    if let Some(author) = component.props.get("author").and_then(Value::as_str) {
        compiled_note = compiled_note.set_author(author);
    }
    if let Some(value) = component.props.get("authorPrefix").and_then(Value::as_bool) {
        compiled_note = compiled_note.add_author_prefix(value);
    }
    if let Some(value) = component.props.get("width").and_then(Value::as_u64) {
        compiled_note =
            compiled_note.set_width(u32::try_from(value).expect("validated Note.width"));
    }
    if let Some(value) = component.props.get("height").and_then(Value::as_u64) {
        compiled_note =
            compiled_note.set_height(u32::try_from(value).expect("validated Note.height"));
    }
    if let Some(value) = component.props.get("visible").and_then(Value::as_bool) {
        compiled_note = compiled_note.set_visible(value);
    }
    if let Some(value) = component
        .props
        .get("backgroundColor")
        .and_then(Value::as_str)
    {
        compiled_note = compiled_note.set_background_color(value);
    }
    if let Some(value) = component.props.get("fontName").and_then(Value::as_str) {
        compiled_note = compiled_note.set_font_name(value);
    }
    if let Some(value) = component.props.get("fontSize").and_then(Value::as_f64) {
        compiled_note = compiled_note.set_font_size(value);
    }
    if let Some(value) = component.props.get("fontFamily").and_then(Value::as_u64) {
        compiled_note =
            compiled_note.set_font_family(u8::try_from(value).expect("validated Note.fontFamily"));
    }
    if let Some(value) = component.props.get("altText").and_then(Value::as_str) {
        compiled_note = compiled_note.set_alt_text(value);
    }
    worksheet
        .insert_note(row, column, &compiled_note)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("cannot add Note to {cell}: {error}")))
}

fn compile_chart(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    component: &Node,
) -> xlsx_jsx::Result<()> {
    let chart_type = chart_type_from_name(
        component.props["type"]
            .as_str()
            .expect("validated Chart.type"),
    );
    let cell = component.props["cell"]
        .as_str()
        .expect("validated Chart.cell");
    let (row, column) = parse_a1_cell_reference(cell).expect("validated Chart.cell");
    let mut chart = Chart::new(chart_type);
    if let Some(hole_size) = component.props.get("holeSize").and_then(Value::as_u64) {
        chart.set_hole_size(u8::try_from(hole_size).expect("validated Chart.holeSize"));
    }
    if let Some(angle) = component
        .props
        .get("firstSliceAngle")
        .and_then(Value::as_u64)
    {
        chart.set_rotation(u16::try_from(angle).expect("validated Chart.firstSliceAngle"));
    }
    if let Some(name) = component.props.get("name").and_then(Value::as_str) {
        chart.set_name(name);
    }
    if let Some(alt_text) = component.props.get("altText").and_then(Value::as_str) {
        chart.set_alt_text(alt_text);
    }
    if let Some(title) = component.props.get("title").and_then(Value::as_str) {
        chart.title().set_name(title);
    }
    if let Some(title) = component.props.get("xAxisTitle").and_then(Value::as_str) {
        chart.x_axis().set_name(title);
    }
    if let Some(title) = component.props.get("yAxisTitle").and_then(Value::as_str) {
        chart.y_axis().set_name(title);
    }
    if component.props.get("showLegend").and_then(Value::as_bool) == Some(false) {
        chart.legend().set_hidden();
    }
    if let Some(position) = component
        .props
        .get("legendPosition")
        .and_then(Value::as_str)
    {
        let position = match position {
            "right" => ChartLegendPosition::Right,
            "left" => ChartLegendPosition::Left,
            "top" => ChartLegendPosition::Top,
            "bottom" => ChartLegendPosition::Bottom,
            "topRight" => ChartLegendPosition::TopRight,
            _ => unreachable!("validated Chart.legendPosition"),
        };
        chart.legend().set_position(position);
    }
    if let Some(width) = component.props.get("width").and_then(Value::as_u64) {
        chart.set_width(u32::try_from(width).expect("validated Chart.width"));
    }
    if let Some(height) = component.props.get("height").and_then(Value::as_u64) {
        chart.set_height(u32::try_from(height).expect("validated Chart.height"));
    }
    for child in &component.children {
        let series: Node = serde_json::from_value(child.clone())
            .map_err(|error| Error::Compile(error.to_string()))?;
        let chart_series = chart.add_series();
        chart_series.set_values(
            series.props["values"]
                .as_str()
                .expect("validated Series.values"),
        );
        if let Some(categories) = series.props.get("categories").and_then(Value::as_str) {
            chart_series.set_categories(categories);
        }
        if let Some(name) = series.props.get("name").and_then(Value::as_str) {
            chart_series.set_name(name);
        }
        if let Some(title) = series.props.get("title").and_then(Value::as_str) {
            chart_series.set_name(title);
        }
    }
    worksheet
        .insert_chart(row, column, &chart)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("cannot insert Chart at {cell}: {error}")))
}

fn chart_type_from_name(chart_type: &str) -> ChartType {
    match chart_type {
        "area" => ChartType::Area,
        "areaStacked" => ChartType::AreaStacked,
        "areaPercentStacked" => ChartType::AreaPercentStacked,
        "column" => ChartType::Column,
        "columnStacked" => ChartType::ColumnStacked,
        "columnPercentStacked" => ChartType::ColumnPercentStacked,
        "bar" => ChartType::Bar,
        "barStacked" => ChartType::BarStacked,
        "barPercentStacked" => ChartType::BarPercentStacked,
        "line" => ChartType::Line,
        "lineStacked" => ChartType::LineStacked,
        "linePercentStacked" => ChartType::LinePercentStacked,
        "radar" => ChartType::Radar,
        "radarWithMarkers" => ChartType::RadarWithMarkers,
        "radarFilled" => ChartType::RadarFilled,
        "scatter" => ChartType::Scatter,
        "scatterWithLines" => ChartType::ScatterStraight,
        "scatterWithLinesAndMarkers" => ChartType::ScatterStraightWithMarkers,
        "pie" => ChartType::Pie,
        "doughnut" => ChartType::Doughnut,
        _ => unreachable!("validated Chart.type"),
    }
}

fn compile_image(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    component: &Node,
) -> xlsx_jsx::Result<()> {
    let cell = component.props["cell"]
        .as_str()
        .expect("validated Image.cell");
    let (row, column) = parse_a1_cell_reference(cell).expect("validated Image.cell");
    let data_uri = component.props["data"]
        .as_str()
        .expect("validated Image.data");
    let (_, encoded) = data_uri
        .strip_prefix("data:image/")
        .and_then(|value| value.split_once(";base64,"))
        .expect("validated Image.data URI");
    let bytes = BASE64.decode(encoded).expect("validated Image.data base64");
    let mut image = XlsxImage::new_from_buffer(&bytes)
        .map_err(|error| Error::Compile(format!("cannot decode Image at {cell}: {error}")))?
        .set_width(
            u32::try_from(
                component.props["width"]
                    .as_u64()
                    .expect("validated Image.width"),
            )
            .expect("validated Image.width"),
        )
        .set_height(
            u32::try_from(
                component.props["height"]
                    .as_u64()
                    .expect("validated Image.height"),
            )
            .expect("validated Image.height"),
        );
    let offset_x = u32::try_from(
        component
            .props
            .get("offsetX")
            .and_then(Value::as_u64)
            .unwrap_or_default(),
    )
    .expect("validated Image.offsetX");
    let offset_y = u32::try_from(
        component
            .props
            .get("offsetY")
            .and_then(Value::as_u64)
            .unwrap_or_default(),
    )
    .expect("validated Image.offsetY");
    if let Some(movement) = component.props.get("movement").and_then(Value::as_str) {
        let movement = match movement {
            "moveAndSize" => ObjectMovement::MoveAndSizeWithCells,
            "moveOnly" => ObjectMovement::MoveButDontSizeWithCells,
            "fixed" => ObjectMovement::DontMoveOrSizeWithCells,
            _ => unreachable!("validated Image.movement"),
        };
        image = image.set_object_movement(movement);
    }
    if let Some(alt_text) = component.props.get("altText").and_then(Value::as_str) {
        image = image.set_alt_text(alt_text);
    }
    if component.props.get("decorative").and_then(Value::as_bool) == Some(true) {
        image = image.set_decorative(true);
    }
    worksheet
        .insert_image_with_offset(row, column, &image, offset_x, offset_y)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("cannot add Image at {cell}: {error}")))
}

fn build_table_column_format(value: &Value) -> xlsx_jsx::Result<Format> {
    let props = value.as_object().cloned().ok_or_else(|| {
        Error::Compile(
            "TableColumn format must be an object; use bold, italic, strikethrough, underline, fontColor, or fillColor"
                .into(),
        )
    })?;
    let node = xlsx_jsx::Node {
        r#type: "Cell".into(),
        props,
        children: Vec::new(),
    };
    build_cell_format(&node)
}

fn table_function(value: &str) -> xlsx_jsx::Result<rust_xlsxwriter::TableFunction> {
    use rust_xlsxwriter::TableFunction;
    Ok(match value {
        "none" => TableFunction::None,
        "average" => TableFunction::Average,
        "count" => TableFunction::Count,
        "countNumbers" => TableFunction::CountNumbers,
        "max" => TableFunction::Max,
        "min" => TableFunction::Min,
        "sum" => TableFunction::Sum,
        "stdDev" => TableFunction::StdDev,
        "var" => TableFunction::Var,
        _ => {
            return Err(Error::Compile(format!(
                "unsupported table total function `{value}`"
            )));
        }
    })
}

fn table_style(style: &str) -> xlsx_jsx::Result<TableStyle> {
    macro_rules! styles { ($value:ident; $($variant:ident),+ $(,)?) => { match $value { $(stringify!($variant) => Ok(TableStyle::$variant),)+ _ => Err(Error::Compile(format!("unsupported table style `{style}`"))) } }; }
    let normalized = style.strip_prefix("TableStyle").unwrap_or(style);
    styles!(normalized;
        None, Light1, Light2, Light3, Light4, Light5, Light6, Light7, Light8, Light9, Light10,
        Light11, Light12, Light13, Light14, Light15, Light16, Light17, Light18, Light19, Light20,
        Light21, Medium1, Medium2, Medium3, Medium4, Medium5, Medium6, Medium7, Medium8, Medium9,
        Medium10, Medium11, Medium12, Medium13, Medium14, Medium15, Medium16, Medium17, Medium18,
        Medium19, Medium20, Medium21, Medium22, Medium23, Medium24, Medium25, Medium26, Medium27,
        Medium28, Dark1, Dark2, Dark3, Dark4, Dark5, Dark6, Dark7, Dark8, Dark9, Dark10, Dark11
    )
}

fn apply_auto_filter(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    let Some(auto_filter) = sheet.props.get("autoFilter").and_then(Value::as_str) else {
        return Ok(());
    };
    let range = parse_a1_range(auto_filter).ok_or_else(|| {
        Error::Compile(format!(
            "Worksheet.autoFilter `{auto_filter}` is not a valid A1 range; use a range such as `A1:D20`"
        ))
    })?;
    worksheet
        .autofilter(range.0, range.1, range.2, range.3)
        .map_err(|error| Error::Compile(format!("invalid Worksheet.autoFilter: {error}")))?;
    let Some(filters) = sheet
        .props
        .get("autoFilterColumns")
        .and_then(Value::as_array)
    else {
        return Ok(());
    };
    for filter in filters {
        let column = filter
            .get("column")
            .and_then(Value::as_u64)
            .ok_or_else(|| Error::Compile("invalid autoFilter column index".into()))?;
        let column = u16::try_from(u64::from(range.1) + column)
            .map_err(|error| Error::Compile(format!("invalid autoFilter column index: {error}")))?;
        let mut condition = if filter
            .get("customFilters")
            .and_then(Value::as_object)
            .is_some()
        {
            compile_custom_filter(filter)?
        } else {
            let mut condition = FilterCondition::new();
            if filter.get("includeBlanks").and_then(Value::as_bool) == Some(true) {
                condition = condition.add_list_blanks_filter();
            }
            if let Some(values) = filter.get("values").and_then(Value::as_array) {
                for value in values {
                    let value = value.as_str().ok_or_else(|| {
                        Error::Compile("autoFilter values must be strings".into())
                    })?;
                    condition = condition.add_list_filter(value);
                }
            }
            condition
        };
        if filter
            .get("customFilters")
            .and_then(|custom| custom.get("logic"))
            .and_then(Value::as_str)
            == Some("or")
        {
            condition = condition.add_custom_boolean_or();
        }
        worksheet
            .filter_column(column, &condition)
            .map_err(|error| Error::Compile(format!("invalid autoFilter criteria: {error}")))?;
    }
    Ok(())
}

fn apply_worksheet_visibility(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    match sheet
        .props
        .get("visibility")
        .and_then(Value::as_str)
        .unwrap_or("visible")
    {
        "visible" => Ok(()),
        "hidden" => {
            worksheet.set_hidden(true);
            Ok(())
        }
        "veryHidden" => {
            worksheet.set_very_hidden(true);
            Ok(())
        }
        _ => Err(Error::Compile(
            "invalid Worksheet.visibility; run xlsx-jsx validate for guidance".into(),
        )),
    }
}

fn apply_worksheet_tab_color(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    let Some(color) = sheet.props.get("tabColor").and_then(Value::as_str) else {
        return Ok(());
    };
    let rgb = u32::from_str_radix(&color[1..], 16).map_err(|error| {
        Error::Compile(format!("Worksheet.tabColor `{color}` is invalid: {error}"))
    })?;
    worksheet.set_tab_color(Color::RGB(rgb));
    Ok(())
}

fn apply_default_row_height(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    let Some(height) = sheet
        .props
        .get("defaultRowHeightPixels")
        .and_then(Value::as_u64)
    else {
        return Ok(());
    };
    let height = u32::try_from(height).map_err(|error| {
        Error::Compile(format!("invalid Worksheet.defaultRowHeightPixels: {error}"))
    })?;
    worksheet.set_default_row_height_pixels(height);
    Ok(())
}

fn apply_hide_unused_rows(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    if let Some(enabled) = sheet.props.get("hideUnusedRows").and_then(Value::as_bool) {
        worksheet.hide_unused_rows(enabled);
        return Ok(());
    }
    if sheet.props.contains_key("hideUnusedRows") {
        return Err(Error::Compile(
            "invalid Worksheet.hideUnusedRows; run xlsx-jsx validate for guidance".into(),
        ));
    }
    Ok(())
}

fn apply_show_gridlines(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    if let Some(show) = sheet.props.get("showGridlines").and_then(Value::as_bool) {
        worksheet.set_screen_gridlines(show);
        return Ok(());
    }
    if sheet.props.contains_key("showGridlines") {
        return Err(Error::Compile(
            "invalid Worksheet.showGridlines; run xlsx-jsx validate for guidance".into(),
        ));
    }
    Ok(())
}

fn apply_worksheet_reading_direction(worksheet: &mut rust_xlsxwriter::Worksheet, sheet: &Node) {
    if sheet.props.get("rightToLeft").and_then(Value::as_bool) == Some(true) {
        worksheet.set_right_to_left(true);
    }
}

fn apply_print_gridlines(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    if let Some(enabled) = sheet.props.get("printGridlines").and_then(Value::as_bool) {
        worksheet.set_print_gridlines(enabled);
        return Ok(());
    }
    if sheet.props.contains_key("printGridlines") {
        return Err(Error::Compile(
            "invalid Worksheet.printGridlines; run xlsx-jsx validate for guidance".into(),
        ));
    }
    Ok(())
}

fn apply_worksheet_zoom(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    if let Some(zoom) = sheet.props.get("zoom").and_then(Value::as_u64) {
        let zoom = u16::try_from(zoom).map_err(|_| {
            Error::Compile("Worksheet.zoom exceeds the writer's integer range".into())
        })?;
        worksheet.set_zoom(zoom);
        return Ok(());
    }
    if sheet.props.contains_key("zoom") {
        return Err(Error::Compile(
            "invalid Worksheet.zoom; run xlsx-jsx validate for guidance".into(),
        ));
    }
    Ok(())
}

fn apply_worksheet_view(worksheet: &mut rust_xlsxwriter::Worksheet, sheet: &Node) {
    match sheet.props.get("view").and_then(Value::as_str) {
        Some("pageLayout") => worksheet.set_view_page_layout(),
        Some("pageBreakPreview") => worksheet.set_view_page_break_preview(),
        _ => worksheet.set_view_normal(),
    };
}

fn apply_worksheet_page_breaks(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    if let Some(values) = sheet
        .props
        .get("horizontalPageBreaks")
        .and_then(Value::as_array)
    {
        let breaks = values
            .iter()
            .map(|value| {
                value.as_u64().and_then(|value| u32::try_from(value).ok()).ok_or_else(|| {
                    Error::Compile("invalid Worksheet.horizontalPageBreaks; run xlsx-jsx validate for guidance".into())
                })
            })
            .collect::<xlsx_jsx::Result<Vec<_>>>()?;
        worksheet
            .set_page_breaks(&breaks)
            .map_err(|error| Error::Compile(format!("invalid horizontal page breaks: {error}")))?;
    }
    if let Some(values) = sheet
        .props
        .get("verticalPageBreaks")
        .and_then(Value::as_array)
    {
        let breaks = values
            .iter()
            .map(|value| {
                value.as_u64().and_then(|value| u32::try_from(value).ok()).ok_or_else(|| {
                    Error::Compile("invalid Worksheet.verticalPageBreaks; run xlsx-jsx validate for guidance".into())
                })
            })
            .collect::<xlsx_jsx::Result<Vec<_>>>()?;
        worksheet
            .set_vertical_page_breaks(&breaks)
            .map_err(|error| Error::Compile(format!("invalid vertical page breaks: {error}")))?;
    }
    Ok(())
}

fn apply_worksheet_page_margins(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    let Some(margins) = sheet.props.get("pageMargins") else {
        return Ok(());
    };
    let Some(margins) = margins.as_object() else {
        return Err(Error::Compile(
            "invalid Worksheet.pageMargins; run xlsx-jsx validate for guidance".into(),
        ));
    };
    let value =
        |key: &str, default: f64| margins.get(key).and_then(Value::as_f64).unwrap_or(default);
    worksheet.set_margins(
        value("left", 0.7),
        value("right", 0.7),
        value("top", 0.75),
        value("bottom", 0.75),
        value("header", 0.3),
        value("footer", 0.3),
    );
    Ok(())
}

fn apply_worksheet_print_settings(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    match sheet.props.get("orientation").and_then(Value::as_str) {
        Some("landscape") => {
            worksheet.set_landscape();
        }
        Some("portrait") => {
            worksheet.set_portrait();
        }
        _ => {}
    }
    if let Some(size) = sheet.props.get("paperSize").and_then(Value::as_u64) {
        let size = u8::try_from(size).map_err(|_| {
            Error::Compile("Worksheet.paperSize exceeds the writer's integer range".into())
        })?;
        worksheet.set_paper_size(size);
    }
    if sheet
        .props
        .get("centerHorizontally")
        .and_then(Value::as_bool)
        == Some(true)
    {
        worksheet.set_print_center_horizontally(true);
    }
    if sheet.props.get("centerVertically").and_then(Value::as_bool) == Some(true) {
        worksheet.set_print_center_vertically(true);
    }
    if sheet
        .props
        .get("printBlackAndWhite")
        .and_then(Value::as_bool)
        == Some(true)
    {
        worksheet.set_print_black_and_white(true);
    }
    if sheet.props.get("printDraft").and_then(Value::as_bool) == Some(true) {
        worksheet.set_print_draft(true);
    }
    if sheet.props.get("printHeadings").and_then(Value::as_bool) == Some(true) {
        worksheet.set_print_headings(true);
    }
    if sheet.props.get("pageOrder").and_then(Value::as_str) == Some("overThenDown") {
        worksheet.set_page_order(false);
    }
    apply_worksheet_print_ranges(worksheet, sheet)?;
    if let Some(scale) = sheet.props.get("printScale").and_then(Value::as_u64) {
        let scale = u16::try_from(scale).map_err(|_| {
            Error::Compile("Worksheet.printScale exceeds the writer's integer range".into())
        })?;
        worksheet.set_print_scale(scale);
    }
    if let Some(fit) = sheet.props.get("fitToPages") {
        let Some(fit) = fit.as_object() else {
            return Err(Error::Compile(
                "invalid Worksheet.fitToPages; run xlsx-jsx validate for guidance".into(),
            ));
        };
        let width = fit
            .get("width")
            .and_then(Value::as_u64)
            .and_then(|value| u16::try_from(value).ok())
            .ok_or_else(|| Error::Compile("invalid Worksheet.fitToPages.width".into()))?;
        let height = fit
            .get("height")
            .and_then(Value::as_u64)
            .and_then(|value| u16::try_from(value).ok())
            .ok_or_else(|| Error::Compile("invalid Worksheet.fitToPages.height".into()))?;
        worksheet.set_print_fit_to_pages(width, height);
    }
    Ok(())
}

fn apply_worksheet_header_footer(worksheet: &mut rust_xlsxwriter::Worksheet, sheet: &Node) {
    if let Some(header) = sheet.props.get("header").and_then(Value::as_str) {
        worksheet.set_header(header);
    }
    if let Some(footer) = sheet.props.get("footer").and_then(Value::as_str) {
        worksheet.set_footer(footer);
    }
    if let Some(enabled) = sheet
        .props
        .get("headerFooterScaleWithDoc")
        .and_then(Value::as_bool)
    {
        worksheet.set_header_footer_scale_with_doc(enabled);
    }
    if let Some(enabled) = sheet
        .props
        .get("headerFooterAlignWithPage")
        .and_then(Value::as_bool)
    {
        worksheet.set_header_footer_align_with_page(enabled);
    }
}

fn apply_worksheet_protection(worksheet: &mut rust_xlsxwriter::Worksheet, sheet: &Node) {
    let Some(values) = sheet.props.get("protection").and_then(Value::as_object) else {
        return;
    };
    let mut options = rust_xlsxwriter::ProtectionOptions::default();
    macro_rules! assign_option {
        ($field:ident, $property:literal) => {
            if let Some(value) = values.get($property).and_then(Value::as_bool) {
                options.$field = value;
            }
        };
    }
    assign_option!(select_locked_cells, "selectLockedCells");
    assign_option!(select_unlocked_cells, "selectUnlockedCells");
    assign_option!(format_cells, "formatCells");
    assign_option!(format_columns, "formatColumns");
    assign_option!(format_rows, "formatRows");
    assign_option!(insert_columns, "insertColumns");
    assign_option!(insert_rows, "insertRows");
    assign_option!(insert_links, "insertLinks");
    assign_option!(delete_columns, "deleteColumns");
    assign_option!(delete_rows, "deleteRows");
    assign_option!(sort, "sort");
    assign_option!(use_autofilter, "useAutoFilter");
    assign_option!(use_pivot_tables, "usePivotTables");
    assign_option!(edit_scenarios, "editScenarios");
    assign_option!(edit_objects, "editObjects");
    worksheet.protect_with_options(&options);
}

fn apply_worksheet_print_ranges(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    if let Some(reference) = sheet.props.get("printArea").and_then(Value::as_str) {
        match parse_print_range(reference).ok_or_else(|| {
            Error::Compile("invalid Worksheet.printArea; run xlsx-jsx validate for guidance".into())
        })? {
            PrintRange::Cells {
                first_row,
                first_column,
                last_row,
                last_column,
            } => worksheet.set_print_area(first_row, first_column, last_row, last_column),
            PrintRange::Rows { first, last } => worksheet.set_print_area(first, 0, last, 16_383),
            PrintRange::Columns { first, last } => {
                worksheet.set_print_area(0, first, 1_048_575, last)
            }
        }
        .map_err(|error| Error::Compile(format!("invalid Worksheet.printArea: {error}")))?;
    }
    if let Some(reference) = sheet.props.get("repeatRows").and_then(Value::as_str) {
        let Some(PrintRange::Rows { first, last }) = parse_print_range(reference) else {
            return Err(Error::Compile(
                "invalid Worksheet.repeatRows; use an inclusive row range such as `1:3`".into(),
            ));
        };
        worksheet
            .set_repeat_rows(first, last)
            .map_err(|error| Error::Compile(format!("invalid Worksheet.repeatRows: {error}")))?;
    }
    if let Some(reference) = sheet.props.get("repeatColumns").and_then(Value::as_str) {
        let Some(PrintRange::Columns { first, last }) = parse_print_range(reference) else {
            return Err(Error::Compile(
                "invalid Worksheet.repeatColumns; use an inclusive column range such as `A:C`"
                    .into(),
            ));
        };
        worksheet
            .set_repeat_columns(first, last)
            .map_err(|error| Error::Compile(format!("invalid Worksheet.repeatColumns: {error}")))?;
    }
    Ok(())
}

fn apply_outline_summary_directions(worksheet: &mut rust_xlsxwriter::Worksheet, sheet: &Node) {
    if sheet.props.get("summaryBelow").and_then(Value::as_bool) == Some(false) {
        worksheet.group_symbols_above(true);
    }
    if sheet.props.get("summaryRight").and_then(Value::as_bool) == Some(false) {
        worksheet.group_symbols_to_left(true);
    }
}

fn compile_custom_filter(filter: &Value) -> xlsx_jsx::Result<FilterCondition> {
    let custom = filter
        .get("customFilters")
        .ok_or_else(|| Error::Compile("missing customFilters".into()))?;
    let conditions = custom
        .get("conditions")
        .and_then(Value::as_array)
        .ok_or_else(|| Error::Compile("invalid customFilters.conditions".into()))?;
    let mut compiled = FilterCondition::new();
    for condition in conditions {
        let value = condition
            .get("value")
            .and_then(Value::as_str)
            .ok_or_else(|| Error::Compile("invalid custom filter value".into()))?;
        let operator = match condition.get("operator").and_then(Value::as_str) {
            None => FilterCriteria::EqualTo,
            Some("lessThan") => FilterCriteria::LessThan,
            Some("notEqual") => FilterCriteria::NotEqualTo,
            Some("greaterThan") => FilterCriteria::GreaterThan,
            Some("lessThanOrEqual") => FilterCriteria::LessThanOrEqualTo,
            Some("greaterThanOrEqual") => FilterCriteria::GreaterThanOrEqualTo,
            Some(other) => {
                return Err(Error::Compile(format!(
                    "unsupported custom filter operator `{other}`"
                )));
            }
        };
        compiled = compiled.add_custom_filter(operator, value);
    }
    Ok(compiled)
}

fn compile_data_validation(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
) -> xlsx_jsx::Result<()> {
    let (first_row, first_column, last_row, last_column) = node
        .props
        .get("range")
        .and_then(Value::as_str)
        .and_then(parse_a1_range)
        .ok_or_else(|| Error::Compile("invalid DataValidation.range".into()))?;
    let mut validation = compile_data_validation_rule(node)?;
    if let Some(value) = node.props.get("allowBlank").and_then(Value::as_bool) {
        validation = validation.ignore_blank(value);
    }
    if let Some(value) = node.props.get("showDropDown").and_then(Value::as_bool) {
        validation = validation.show_dropdown(value);
    }
    if let Some(value) = node.props.get("showInputMessage").and_then(Value::as_bool) {
        validation = validation.show_input_message(value);
    }
    if let Some(value) = node.props.get("showErrorMessage").and_then(Value::as_bool) {
        validation = validation.show_error_message(value);
    }
    if let Some(value) = node.props.get("inputTitle").and_then(Value::as_str) {
        validation = validation.set_input_title(value).map_err(|error| {
            Error::Compile(format!("invalid DataValidation.inputTitle: {error}"))
        })?;
    }
    if let Some(value) = node.props.get("inputMessage").and_then(Value::as_str) {
        validation = validation.set_input_message(value).map_err(|error| {
            Error::Compile(format!("invalid DataValidation.inputMessage: {error}"))
        })?;
    }
    if let Some(value) = node.props.get("errorTitle").and_then(Value::as_str) {
        validation = validation.set_error_title(value).map_err(|error| {
            Error::Compile(format!("invalid DataValidation.errorTitle: {error}"))
        })?;
    }
    if let Some(value) = node.props.get("errorMessage").and_then(Value::as_str) {
        validation = validation.set_error_message(value).map_err(|error| {
            Error::Compile(format!("invalid DataValidation.errorMessage: {error}"))
        })?;
    }
    if let Some(value) = node.props.get("errorStyle").and_then(Value::as_str) {
        let error_style = match value {
            "warning" => DataValidationErrorStyle::Warning,
            "information" => DataValidationErrorStyle::Information,
            _ => DataValidationErrorStyle::Stop,
        };
        validation = validation.set_error_style(error_style);
    }
    worksheet
        .add_data_validation(first_row, first_column, last_row, last_column, &validation)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("invalid DataValidation: {error}")))
}

#[expect(
    clippy::too_many_lines,
    reason = "Conditional-format writer dispatch keeps family routing and shared range defaults together."
)]
fn compile_conditional_format(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
) -> xlsx_jsx::Result<()> {
    let range = node
        .props
        .get("range")
        .and_then(Value::as_str)
        .and_then(parse_a1_range)
        .ok_or_else(|| Error::Compile("invalid ConditionalFormat.range".into()))?;
    let scale_type = node
        .props
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::Compile("ConditionalFormat.type is required".into()))?;
    let minimum_color = node
        .props
        .get("minColor")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim_start_matches('#');
    let maximum_color = node
        .props
        .get("maxColor")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim_start_matches('#');
    let stop_if_true = node
        .props
        .get("stopIfTrue")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    match scale_type {
        "average" | "date" | "blank" | "error" | "duplicate" => {
            compile_simple_conditional_format(worksheet, node, range, scale_type, stop_if_true)
        }
        "cellIs" => compile_cell_is_conditional_format(worksheet, node, range, stop_if_true),
        "formula" => compile_formula_conditional_format(worksheet, node, range, stop_if_true),
        "text" => compile_text_conditional_format(worksheet, node, range, stop_if_true),
        "dataBar" => compile_data_bar_conditional_format(worksheet, node, range),
        "iconSet" => compile_icon_set_conditional_format(worksheet, node, range, stop_if_true),
        "top10" => compile_top10_conditional_format(worksheet, node, range, stop_if_true),
        "colorScale2" => {
            let mut rule = ConditionalFormat2ColorScale::new()
                .set_minimum_color(minimum_color)
                .set_maximum_color(maximum_color)
                .set_stop_if_true(stop_if_true);
            if let Some(rule_type) = node.props.get("minType").and_then(Value::as_str) {
                rule = rule.set_minimum(
                    conditional_format_threshold_type(rule_type)?,
                    conditional_format_threshold_value(node, "min", rule_type)?,
                );
            }
            if let Some(rule_type) = node.props.get("maxType").and_then(Value::as_str) {
                rule = rule.set_maximum(
                    conditional_format_threshold_type(rule_type)?,
                    conditional_format_threshold_value(node, "max", rule_type)?,
                );
            }
            worksheet
                .add_conditional_format(range.0, range.1, range.2, range.3, &rule)
                .map(|_| ())
                .map_err(|error| {
                    Error::Compile(format!("invalid ConditionalFormat colorScale2: {error}"))
                })
        }
        "colorScale3" => {
            let midpoint_color = node
                .props
                .get("midColor")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .trim_start_matches('#');
            let mut rule = ConditionalFormat3ColorScale::new()
                .set_minimum_color(minimum_color)
                .set_midpoint_color(midpoint_color)
                .set_maximum_color(maximum_color)
                .set_stop_if_true(stop_if_true);
            if let Some(rule_type) = node.props.get("minType").and_then(Value::as_str) {
                rule = rule.set_minimum(
                    conditional_format_threshold_type(rule_type)?,
                    conditional_format_threshold_value(node, "min", rule_type)?,
                );
            }
            if let Some(rule_type) = node.props.get("midType").and_then(Value::as_str) {
                rule = rule.set_midpoint(
                    conditional_format_threshold_type(rule_type)?,
                    conditional_format_threshold_value(node, "mid", rule_type)?,
                );
            }
            if let Some(rule_type) = node.props.get("maxType").and_then(Value::as_str) {
                rule = rule.set_maximum(
                    conditional_format_threshold_type(rule_type)?,
                    conditional_format_threshold_value(node, "max", rule_type)?,
                );
            }
            worksheet
                .add_conditional_format(range.0, range.1, range.2, range.3, &rule)
                .map(|_| ())
                .map_err(|error| {
                    Error::Compile(format!("invalid ConditionalFormat colorScale3: {error}"))
                })
        }
        _ => Err(Error::Compile(format!(
            "unsupported ConditionalFormat.type `{scale_type}`"
        ))),
    }
}

#[expect(
    clippy::too_many_lines,
    reason = "The five related simple conditional-format writers share rule, style, stop-if-true, and range semantics."
)]
fn compile_simple_conditional_format(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
    range: (u32, u16, u32, u16),
    scale_type: &str,
    stop_if_true: bool,
) -> xlsx_jsx::Result<()> {
    let default_rule = match scale_type {
        "average" => "aboveAverage",
        "date" => "yesterday",
        "blank" => "blank",
        "error" => "error",
        "duplicate" => "duplicate",
        _ => {
            return Err(Error::Compile(format!(
                "unsupported simple ConditionalFormat type `{scale_type}`"
            )));
        }
    };
    let rule = node
        .props
        .get("rule")
        .and_then(Value::as_str)
        .unwrap_or(default_rule);
    let format = compile_conditional_format_style(node)?;
    let has_format = [
        "bold",
        "italic",
        "strikethrough",
        "underline",
        "script",
        "fontColor",
        "fillColor",
        "fillPattern",
        "fillForegroundColor",
        "fillBackgroundColor",
        "numberFormat",
        "borderTop",
        "borderRight",
        "borderBottom",
        "borderLeft",
        "borderTopColor",
        "borderRightColor",
        "borderBottomColor",
        "borderLeftColor",
    ]
    .iter()
    .any(|property| node.props.contains_key(*property));
    match scale_type {
        "average" => {
            let rule = match rule {
                "aboveAverage" => ConditionalFormatAverageRule::AboveAverage,
                "belowAverage" => ConditionalFormatAverageRule::BelowAverage,
                "equalOrAboveAverage" => ConditionalFormatAverageRule::EqualOrAboveAverage,
                "equalOrBelowAverage" => ConditionalFormatAverageRule::EqualOrBelowAverage,
                "oneStdDevAbove" => ConditionalFormatAverageRule::OneStandardDeviationAbove,
                "oneStdDevBelow" => ConditionalFormatAverageRule::OneStandardDeviationBelow,
                "twoStdDevAbove" => ConditionalFormatAverageRule::TwoStandardDeviationsAbove,
                "twoStdDevBelow" => ConditionalFormatAverageRule::TwoStandardDeviationsBelow,
                "threeStdDevAbove" => ConditionalFormatAverageRule::ThreeStandardDeviationsAbove,
                "threeStdDevBelow" => ConditionalFormatAverageRule::ThreeStandardDeviationsBelow,
                _ => return Err(Error::Compile(format!("unsupported average rule `{rule}`"))),
            };
            let mut conditional = ConditionalFormatAverage::new()
                .set_rule(rule)
                .set_stop_if_true(stop_if_true);
            if has_format {
                conditional = conditional.set_format(format);
            }
            worksheet
                .add_conditional_format(range.0, range.1, range.2, range.3, &conditional)
                .map(|_| ())
                .map_err(|error| {
                    Error::Compile(format!("invalid ConditionalFormat average: {error}"))
                })
        }
        "date" => {
            let rule = match rule {
                "yesterday" => ConditionalFormatDateRule::Yesterday,
                "today" => ConditionalFormatDateRule::Today,
                "tomorrow" => ConditionalFormatDateRule::Tomorrow,
                "last7Days" => ConditionalFormatDateRule::Last7Days,
                "lastWeek" => ConditionalFormatDateRule::LastWeek,
                "thisWeek" => ConditionalFormatDateRule::ThisWeek,
                "nextWeek" => ConditionalFormatDateRule::NextWeek,
                "lastMonth" => ConditionalFormatDateRule::LastMonth,
                "thisMonth" => ConditionalFormatDateRule::ThisMonth,
                "nextMonth" => ConditionalFormatDateRule::NextMonth,
                _ => return Err(Error::Compile(format!("unsupported date rule `{rule}`"))),
            };
            let mut conditional = ConditionalFormatDate::new()
                .set_rule(rule)
                .set_stop_if_true(stop_if_true);
            if has_format {
                conditional = conditional.set_format(format);
            }
            worksheet
                .add_conditional_format(range.0, range.1, range.2, range.3, &conditional)
                .map(|_| ())
                .map_err(|error| Error::Compile(format!("invalid ConditionalFormat date: {error}")))
        }
        "blank" => {
            let conditional = ConditionalFormatBlank::new();
            let mut conditional = if rule == "nonBlank" {
                conditional.invert()
            } else {
                conditional
            }
            .set_stop_if_true(stop_if_true);
            if has_format {
                conditional = conditional.set_format(format);
            }
            worksheet
                .add_conditional_format(range.0, range.1, range.2, range.3, &conditional)
                .map(|_| ())
                .map_err(|error| {
                    Error::Compile(format!("invalid ConditionalFormat blank: {error}"))
                })
        }
        "error" => {
            let conditional = ConditionalFormatError::new();
            let mut conditional = if rule == "nonError" {
                conditional.invert()
            } else {
                conditional
            }
            .set_stop_if_true(stop_if_true);
            if has_format {
                conditional = conditional.set_format(format);
            }
            worksheet
                .add_conditional_format(range.0, range.1, range.2, range.3, &conditional)
                .map(|_| ())
                .map_err(|error| {
                    Error::Compile(format!("invalid ConditionalFormat error: {error}"))
                })
        }
        "duplicate" => {
            let conditional = ConditionalFormatDuplicate::new();
            let mut conditional = if rule == "unique" {
                conditional.invert()
            } else {
                conditional
            }
            .set_stop_if_true(stop_if_true);
            if has_format {
                conditional = conditional.set_format(format);
            }
            worksheet
                .add_conditional_format(range.0, range.1, range.2, range.3, &conditional)
                .map(|_| ())
                .map_err(|error| {
                    Error::Compile(format!("invalid ConditionalFormat duplicate: {error}"))
                })
        }
        _ => Err(Error::Compile(format!(
            "unsupported ConditionalFormat type `{scale_type}`"
        ))),
    }
}

fn compile_data_bar_conditional_format(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
    range: (u32, u16, u32, u16),
) -> xlsx_jsx::Result<()> {
    let fill_color = node
        .props
        .get("fillColor")
        .and_then(Value::as_str)
        .unwrap_or("#638EC6");
    let mut rule =
        ConditionalFormatDataBar::new().set_fill_color(fill_color.trim_start_matches('#'));
    for endpoint in ["min", "max"] {
        let type_property = format!("{endpoint}Type");
        if let Some(kind) = node.props.get(&type_property).and_then(Value::as_str) {
            let threshold_type = match kind {
                "automatic" => ConditionalFormatType::Automatic,
                _ => conditional_format_threshold_type(kind)?,
            };
            let value = if kind == "automatic" || kind == "min" || kind == "max" {
                0.0.into()
            } else {
                conditional_format_threshold_value(node, endpoint, kind)?
            };
            rule = if endpoint == "min" {
                rule.set_minimum(threshold_type, value)
            } else {
                rule.set_maximum(threshold_type, value)
            };
        }
    }
    for (property, setter) in [
        ("borderColor", 0_u8),
        ("negativeFillColor", 1),
        ("negativeBorderColor", 2),
        ("axisColor", 3),
    ] {
        if let Some(color) = node.props.get(property).and_then(Value::as_str) {
            let color = color.trim_start_matches('#');
            rule = match setter {
                0 => rule.set_border_color(color),
                1 => rule.set_negative_fill_color(color),
                2 => rule.set_negative_border_color(color),
                _ => rule.set_axis_color(color),
            };
        }
    }
    if let Some(value) = node.props.get("solidFill").and_then(Value::as_bool) {
        rule = rule.set_solid_fill(value);
    }
    if let Some(value) = node.props.get("border").and_then(Value::as_bool) {
        rule = rule.set_border_off(!value);
    }
    if let Some(value) = node.props.get("barOnly").and_then(Value::as_bool) {
        rule = rule.set_bar_only(value);
    }
    if let Some(direction) = node.props.get("direction").and_then(Value::as_str) {
        rule = rule.set_direction(match direction {
            "context" => ConditionalFormatDataBarDirection::Context,
            "leftToRight" => ConditionalFormatDataBarDirection::LeftToRight,
            "rightToLeft" => ConditionalFormatDataBarDirection::RightToLeft,
            _ => {
                return Err(Error::Compile(format!(
                    "unsupported ConditionalFormat.direction `{direction}`"
                )));
            }
        });
    }
    if let Some(position) = node.props.get("axisPosition").and_then(Value::as_str) {
        rule = rule.set_axis_position(match position {
            "automatic" => ConditionalFormatDataBarAxisPosition::Automatic,
            "midpoint" => ConditionalFormatDataBarAxisPosition::Midpoint,
            "none" => ConditionalFormatDataBarAxisPosition::None,
            _ => {
                return Err(Error::Compile(format!(
                    "unsupported ConditionalFormat.axisPosition `{position}`"
                )));
            }
        });
    }
    worksheet
        .add_conditional_format(range.0, range.1, range.2, range.3, &rule)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("invalid ConditionalFormat dataBar: {error}")))
}

fn compile_icon_set_conditional_format(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
    range: (u32, u16, u32, u16),
    stop_if_true: bool,
) -> xlsx_jsx::Result<()> {
    let icon_type = node
        .props
        .get("iconType")
        .and_then(Value::as_str)
        .unwrap_or("threeTrafficLights");
    let mut rule = ConditionalFormatIconSet::new()
        .set_icon_type(conditional_format_icon_type(icon_type)?)
        .reverse_icons(node.props.get("reverse").and_then(Value::as_bool) == Some(true))
        .show_icons_only(node.props.get("showIconsOnly").and_then(Value::as_bool) == Some(true))
        .set_stop_if_true(stop_if_true);
    if let Some(icon_values) = node.props.get("icons").and_then(Value::as_array) {
        let mut icons = Vec::with_capacity(icon_values.len());
        for (index, value) in icon_values.iter().enumerate() {
            let properties = value.as_object().ok_or_else(|| {
                Error::Compile(format!(
                    "ConditionalFormat.icons[{index}] must be an object"
                ))
            })?;
            let kind = properties
                .get("type")
                .and_then(Value::as_str)
                .unwrap_or_default();
            let threshold_type = match kind {
                "number" => ConditionalFormatType::Number,
                "percent" => ConditionalFormatType::Percent,
                "percentile" => ConditionalFormatType::Percentile,
                "formula" => ConditionalFormatType::Formula,
                _ => {
                    return Err(Error::Compile(format!(
                        "ConditionalFormat.icons[{index}].type is unsupported"
                    )));
                }
            };
            let threshold_value: ConditionalFormatValue = if kind == "formula" {
                let formula = properties
                    .get("value")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::Compile(format!(
                            "ConditionalFormat.icons[{index}].value must be a formula string"
                        ))
                    })?;
                Formula::new(formula).into()
            } else {
                properties
                    .get("value")
                    .and_then(Value::as_f64)
                    .ok_or_else(|| {
                        Error::Compile(format!(
                            "ConditionalFormat.icons[{index}].value must be numeric"
                        ))
                    })?
                    .into()
            };
            let mut icon =
                ConditionalFormatCustomIcon::new().set_rule(threshold_type, threshold_value);
            if properties.get("greaterThan").and_then(Value::as_bool) == Some(true) {
                icon = icon.set_greater_than(true);
            }
            if properties.get("noIcon").and_then(Value::as_bool) == Some(true) {
                icon = icon.set_no_icon(true);
            } else if let Some(custom_type) = properties.get("iconType").and_then(Value::as_str) {
                let icon_index = properties.get("iconIndex").and_then(Value::as_u64).and_then(|value| u8::try_from(value).ok()).ok_or_else(|| Error::Compile(format!("ConditionalFormat.icons[{index}].iconIndex must be an integer from 0 to 4")))?;
                icon = icon.set_icon_type(conditional_format_icon_type(custom_type)?, icon_index);
            }
            icons.push(icon);
        }
        rule = rule.set_icons(&icons);
    }
    worksheet
        .add_conditional_format(range.0, range.1, range.2, range.3, &rule)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("invalid ConditionalFormat iconSet: {error}")))
}

fn compile_top10_conditional_format(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
    range: (u32, u16, u32, u16),
    stop_if_true: bool,
) -> xlsx_jsx::Result<()> {
    let rule_name = node
        .props
        .get("rule")
        .and_then(Value::as_str)
        .unwrap_or("top");
    let rank = node
        .props
        .get("rank")
        .and_then(Value::as_u64)
        .map(u16::try_from)
        .transpose()
        .map_err(|_| Error::Compile("ConditionalFormat.rank exceeds u16".into()))?
        .unwrap_or(10);
    let rule = match rule_name {
        "top" => ConditionalFormatTopRule::Top(rank),
        "bottom" => ConditionalFormatTopRule::Bottom(rank),
        "topPercent" => ConditionalFormatTopRule::TopPercent(rank),
        "bottomPercent" => ConditionalFormatTopRule::BottomPercent(rank),
        _ => {
            return Err(Error::Compile(format!(
                "unsupported ConditionalFormat.rule `{rule_name}`"
            )));
        }
    };
    let has_style = [
        "bold",
        "italic",
        "strikethrough",
        "underline",
        "script",
        "fontColor",
        "fillColor",
        "numberFormat",
        "borderTop",
        "borderRight",
        "borderBottom",
        "borderLeft",
        "borderTopColor",
        "borderRightColor",
        "borderBottomColor",
        "borderLeftColor",
    ]
    .iter()
    .any(|property| node.props.contains_key(*property));
    let mut conditional = ConditionalFormatTop::new()
        .set_rule(rule)
        .set_stop_if_true(stop_if_true);
    if has_style {
        conditional = conditional.set_format(compile_conditional_format_style(node)?);
    }
    worksheet
        .add_conditional_format(range.0, range.1, range.2, range.3, &conditional)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("invalid ConditionalFormat top10: {error}")))
}

fn conditional_format_icon_type(value: &str) -> xlsx_jsx::Result<ConditionalFormatIconType> {
    match value {
        "threeArrows" => Ok(ConditionalFormatIconType::ThreeArrows),
        "threeArrowsGray" => Ok(ConditionalFormatIconType::ThreeArrowsGray),
        "threeFlags" => Ok(ConditionalFormatIconType::ThreeFlags),
        "threeTrafficLights" => Ok(ConditionalFormatIconType::ThreeTrafficLights),
        "threeTrafficLightsWithRim" => Ok(ConditionalFormatIconType::ThreeTrafficLightsWithRim),
        "threeSigns" => Ok(ConditionalFormatIconType::ThreeSigns),
        "threeSymbolsCircled" => Ok(ConditionalFormatIconType::ThreeSymbolsCircled),
        "threeSymbols" => Ok(ConditionalFormatIconType::ThreeSymbols),
        "threeStars" => Ok(ConditionalFormatIconType::ThreeStars),
        "threeTriangles" => Ok(ConditionalFormatIconType::ThreeTriangles),
        "fourArrows" => Ok(ConditionalFormatIconType::FourArrows),
        "fourArrowsGray" => Ok(ConditionalFormatIconType::FourArrowsGray),
        "fourRedToBlack" => Ok(ConditionalFormatIconType::FourRedToBlack),
        "fourHistograms" => Ok(ConditionalFormatIconType::FourHistograms),
        "fourTrafficLights" => Ok(ConditionalFormatIconType::FourTrafficLights),
        "fiveArrows" => Ok(ConditionalFormatIconType::FiveArrows),
        "fiveArrowsGray" => Ok(ConditionalFormatIconType::FiveArrowsGray),
        "fiveHistograms" => Ok(ConditionalFormatIconType::FiveHistograms),
        "fiveQuadrants" => Ok(ConditionalFormatIconType::FiveQuadrants),
        "fiveBoxes" => Ok(ConditionalFormatIconType::FiveBoxes),
        _ => Err(Error::Compile(format!(
            "ConditionalFormat.iconType `{value}` is unsupported"
        ))),
    }
}

fn conditional_icon_type_from_xml(value: &str, part: &str) -> xlsx_jsx::Result<String> {
    let icon_type = match value {
        "3Arrows" => "threeArrows",
        "3ArrowsGray" => "threeArrowsGray",
        "3Flags" => "threeFlags",
        "3TrafficLights1" => "threeTrafficLights",
        "3TrafficLights2" => "threeTrafficLightsWithRim",
        "3Signs" => "threeSigns",
        "3Symbols" => "threeSymbolsCircled",
        "3Symbols2" => "threeSymbols",
        "3Stars" => "threeStars",
        "3Triangles" => "threeTriangles",
        "4Arrows" => "fourArrows",
        "4ArrowsGray" => "fourArrowsGray",
        "4RedToBlack" => "fourRedToBlack",
        "4Rating" => "fourHistograms",
        "4TrafficLights" => "fourTrafficLights",
        "5Arrows" => "fiveArrows",
        "5ArrowsGray" => "fiveArrowsGray",
        "5Rating" => "fiveHistograms",
        "5Quarters" => "fiveQuadrants",
        "5Boxes" => "fiveBoxes",
        unsupported => {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` iconSet type `{unsupported}` is unsupported; use a documented built-in icon set"
            )));
        }
    };
    Ok(icon_type.to_owned())
}

fn compile_formula_conditional_format(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
    range: (u32, u16, u32, u16),
    stop_if_true: bool,
) -> xlsx_jsx::Result<()> {
    let formula = node
        .props
        .get("formula1")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::Compile("ConditionalFormat.formula1 is required".into()))?;
    let conditional = ConditionalFormatFormula::new()
        .set_rule(Formula::new(formula))
        .set_format(compile_conditional_format_style(node)?)
        .set_stop_if_true(stop_if_true);
    worksheet
        .add_conditional_format(range.0, range.1, range.2, range.3, &conditional)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("invalid ConditionalFormat formula: {error}")))
}

fn compile_text_conditional_format(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
    range: (u32, u16, u32, u16),
    stop_if_true: bool,
) -> xlsx_jsx::Result<()> {
    let operator = node
        .props
        .get("operator")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::Compile("ConditionalFormat.operator is required".into()))?;
    let text = node
        .props
        .get("text")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::Compile("ConditionalFormat.text is required".into()))?;
    let conditional = ConditionalFormatText::new()
        .set_rule(conditional_text_rule(operator, text)?)
        .set_format(compile_conditional_format_style(node)?)
        .set_stop_if_true(stop_if_true);
    worksheet
        .add_conditional_format(range.0, range.1, range.2, range.3, &conditional)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("invalid ConditionalFormat text: {error}")))
}

fn conditional_text_rule(
    operator: &str,
    text: &str,
) -> xlsx_jsx::Result<ConditionalFormatTextRule> {
    match operator {
        "containsText" => Ok(ConditionalFormatTextRule::Contains(text.to_owned())),
        "notContains" => Ok(ConditionalFormatTextRule::DoesNotContain(text.to_owned())),
        "beginsWith" => Ok(ConditionalFormatTextRule::BeginsWith(text.to_owned())),
        "endsWith" => Ok(ConditionalFormatTextRule::EndsWith(text.to_owned())),
        _ => Err(Error::Compile(format!(
            "ConditionalFormat.operator `{operator}` is unsupported; use containsText, notContains, beginsWith, or endsWith"
        ))),
    }
}

fn compile_conditional_format_style(node: &Node) -> xlsx_jsx::Result<Format> {
    let mut format = Format::new();
    if node.props.get("bold").and_then(Value::as_bool) == Some(true) {
        format = format.set_bold();
    }
    if node.props.get("italic").and_then(Value::as_bool) == Some(true) {
        format = format.set_italic();
    }
    if node.props.get("strikethrough").and_then(Value::as_bool) == Some(true) {
        format = format.set_font_strikethrough();
    }
    if let Some(underline) = node.props.get("underline").and_then(Value::as_str) {
        format = format.set_underline(match underline {
            "single" => FormatUnderline::Single,
            "double" => FormatUnderline::Double,
            "singleAccounting" => FormatUnderline::SingleAccounting,
            "doubleAccounting" => FormatUnderline::DoubleAccounting,
            _ => FormatUnderline::None,
        });
    }
    if let Some(script) = node.props.get("script").and_then(Value::as_str) {
        format = format.set_font_script(match script {
            "superscript" => FormatScript::Superscript,
            "subscript" => FormatScript::Subscript,
            _ => FormatScript::None,
        });
    }
    if let Some(color) = node.props.get("fontColor").and_then(Value::as_str) {
        format = format.set_font_color(color.trim_start_matches('#'));
    }
    if let Some(number_format) = node.props.get("numberFormat").and_then(Value::as_str) {
        format = format.set_num_format(number_format);
    }
    format = apply_cell_borders(format, node);
    apply_cell_fill(format, node)
}

fn compile_ignored_error(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
) -> xlsx_jsx::Result<()> {
    let (first_row, first_column, last_row, last_column) = node
        .props
        .get("range")
        .and_then(Value::as_str)
        .and_then(parse_a1_range)
        .ok_or_else(|| Error::Compile("invalid IgnoreError.range".into()))?;
    let error_type = match node
        .props
        .get("errorType")
        .and_then(Value::as_str)
        .unwrap_or_default()
    {
        "numberStoredAsText" => IgnoreError::NumberStoredAsText,
        "evalError" => IgnoreError::FormulaError,
        "formula" => IgnoreError::FormulaDiffers,
        "emptyCellReference" => IgnoreError::FormulaRefersToEmptyCells,
        "formulaRange" => IgnoreError::FormulaOmitsCells,
        "listDataValidation" => IgnoreError::DataValidationError,
        "TwoDigitTextYear" => IgnoreError::TwoDigitTextYear,
        "unlockedFormula" => IgnoreError::UnlockedCellsWithFormula,
        "calculatedColumn" => IgnoreError::InconsistentColumnFormula,
        _ => return Err(Error::Compile("invalid IgnoreError.errorType".into())),
    };
    worksheet
        .ignore_error_range(first_row, first_column, last_row, last_column, error_type)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("invalid IgnoreError: {error}")))
}

fn compile_cell_is_conditional_format(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    node: &Node,
    range: (u32, u16, u32, u16),
    stop_if_true: bool,
) -> xlsx_jsx::Result<()> {
    let operator = node
        .props
        .get("operator")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::Compile("missing ConditionalFormat.operator".into()))?;
    let formula1 = node
        .props
        .get("formula1")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::Compile("missing ConditionalFormat.formula1".into()))?;
    let first = Formula::new(formula1);
    let second = node
        .props
        .get("formula2")
        .and_then(Value::as_str)
        .map(Formula::new);
    let rule = match operator {
        "equal" => ConditionalFormatCellRule::EqualTo(first),
        "notEqual" => ConditionalFormatCellRule::NotEqualTo(first),
        "greaterThan" => ConditionalFormatCellRule::GreaterThan(first),
        "greaterThanOrEqual" => ConditionalFormatCellRule::GreaterThanOrEqualTo(first),
        "lessThan" => ConditionalFormatCellRule::LessThan(first),
        "lessThanOrEqual" => ConditionalFormatCellRule::LessThanOrEqualTo(first),
        "between" => ConditionalFormatCellRule::Between(
            first,
            second.ok_or_else(|| Error::Compile("missing ConditionalFormat.formula2".into()))?,
        ),
        "notBetween" => ConditionalFormatCellRule::NotBetween(
            first,
            second.ok_or_else(|| Error::Compile("missing ConditionalFormat.formula2".into()))?,
        ),
        _ => {
            return Err(Error::Compile(format!(
                "unsupported ConditionalFormat.operator `{operator}`"
            )));
        }
    };
    let conditional = ConditionalFormatCell::new()
        .set_rule(rule)
        .set_format(compile_conditional_format_style(node)?)
        .set_stop_if_true(stop_if_true);
    worksheet
        .add_conditional_format(range.0, range.1, range.2, range.3, &conditional)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("invalid ConditionalFormat cellIs: {error}")))
}

fn conditional_format_threshold_type(value: &str) -> xlsx_jsx::Result<ConditionalFormatType> {
    match value {
        "min" => Ok(ConditionalFormatType::Lowest),
        "max" => Ok(ConditionalFormatType::Highest),
        "number" => Ok(ConditionalFormatType::Number),
        "percent" => Ok(ConditionalFormatType::Percent),
        "percentile" => Ok(ConditionalFormatType::Percentile),
        "formula" => Ok(ConditionalFormatType::Formula),
        _ => Err(Error::Compile(format!(
            "unsupported conditional format threshold type `{value}`"
        ))),
    }
}

fn conditional_format_threshold_value(
    node: &Node,
    stop: &str,
    rule_type: &str,
) -> xlsx_jsx::Result<rust_xlsxwriter::ConditionalFormatValue> {
    let property = format!("{stop}Value");
    if rule_type == "formula" {
        let formula = node
            .props
            .get(&property)
            .and_then(Value::as_str)
            .ok_or_else(|| Error::Compile(format!("missing ConditionalFormat.{property}")))?;
        Ok(Formula::new(formula).into())
    } else {
        let number = node
            .props
            .get(&property)
            .and_then(Value::as_f64)
            .ok_or_else(|| {
                Error::Compile(format!("missing numeric ConditionalFormat.{property}"))
            })?;
        Ok(number.into())
    }
}

fn compile_data_validation_rule(node: &Node) -> xlsx_jsx::Result<DataValidation> {
    let validation_type = node
        .props
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::Compile("invalid DataValidation.type".into()))?;
    if validation_type == "any" {
        return Ok(DataValidation::new().allow_any_value());
    }
    if validation_type == "list" {
        let source = node
            .props
            .get("source")
            .and_then(Value::as_str)
            .ok_or_else(|| Error::Compile("invalid DataValidation.source".into()))?;
        return Ok(DataValidation::new().allow_list_formula(source.into()));
    }
    let formula1 = node
        .props
        .get("formula1")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::Compile("invalid DataValidation.formula1".into()))?;
    if validation_type == "custom" {
        return Ok(DataValidation::new().allow_custom(Formula::new(formula1)));
    }
    let formula2 = node.props.get("formula2").and_then(Value::as_str);
    let operator = node.props.get("operator").and_then(Value::as_str);
    let rule = compile_data_validation_formula_rule(operator, formula1, formula2)?;
    let validation = match validation_type {
        "whole" => DataValidation::new().allow_whole_number_formula(rule),
        "decimal" => DataValidation::new().allow_decimal_number_formula(rule),
        "date" => DataValidation::new().allow_date_formula(rule),
        "time" => DataValidation::new().allow_time_formula(rule),
        "textLength" => DataValidation::new().allow_text_length_formula(rule),
        _ => {
            return Err(Error::Compile(format!(
                "unsupported DataValidation.type `{validation_type}`"
            )));
        }
    };
    Ok(validation)
}

fn compile_data_validation_formula_rule(
    operator: Option<&str>,
    formula1: &str,
    formula2: Option<&str>,
) -> xlsx_jsx::Result<DataValidationRule<Formula>> {
    let first = Formula::new(formula1);
    let second = formula2.map(Formula::new);
    match (operator, second) {
        (Some("notBetween"), Some(second)) => Ok(DataValidationRule::NotBetween(first, second)),
        (Some("notBetween"), None) => Err(Error::Compile(
            "DataValidation operator `notBetween` requires formula2".into(),
        )),
        (Some("between") | None, Some(second)) => Ok(DataValidationRule::Between(first, second)),
        (Some("between"), None) => Err(Error::Compile(
            "DataValidation operator `between` requires formula2".into(),
        )),
        (Some("equal"), None) => Ok(DataValidationRule::EqualTo(first)),
        (Some("notEqual"), None) => Ok(DataValidationRule::NotEqualTo(first)),
        (Some("greaterThan"), None) => Ok(DataValidationRule::GreaterThan(first)),
        (Some("greaterThanOrEqual"), None) => Ok(DataValidationRule::GreaterThanOrEqualTo(first)),
        (Some("lessThan"), None) => Ok(DataValidationRule::LessThan(first)),
        (Some("lessThanOrEqual"), None) => Ok(DataValidationRule::LessThanOrEqualTo(first)),
        (Some(operator), _) => Err(Error::Compile(format!(
            "DataValidation.operator `{operator}` cannot be combined with formula2"
        ))),
        (None, None) => Err(Error::Compile(
            "DataValidation requires operator or formula2 for a between rule".into(),
        )),
    }
}

fn compile_row(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: &Node,
    row_index: usize,
    merged_ranges: &mut Vec<(u32, u16, u32, u16)>,
) -> xlsx_jsx::Result<()> {
    let row_number = u32::try_from(row_index)
        .map_err(|_| Error::Compile("worksheet exceeds row limit".into()))?;
    if let Some(height) = row.props.get("height").and_then(Value::as_f64) {
        worksheet
            .set_row_height(row_number, height)
            .map_err(|error| Error::Compile(error.to_string()))?;
    }
    if row.props.get("hidden").and_then(Value::as_bool) == Some(true) {
        worksheet
            .set_row_hidden(row_number)
            .map_err(|error| Error::Compile(error.to_string()))?;
    }
    if has_format_properties(row) {
        let format = build_cell_format(row)?;
        worksheet
            .set_row_format(row_number, &format)
            .map_err(|error| Error::Compile(error.to_string()))?;
    }
    let mut column_index = 0_u16;
    for cell_value in &row.children {
        let cell: Node = serde_json::from_value(cell_value.clone())
            .map_err(|error| Error::Compile(error.to_string()))?;
        while merged_ranges
            .iter()
            .any(|(first_row, first_col, last_row, last_col)| {
                *first_row <= row_number
                    && row_number <= *last_row
                    && *first_col <= column_index
                    && column_index <= *last_col
            })
        {
            column_index = column_index
                .checked_add(1)
                .ok_or_else(|| Error::Compile("worksheet exceeds column limit".into()))?;
        }
        let row_span = u32::try_from(
            cell.props
                .get("rowSpan")
                .and_then(Value::as_u64)
                .unwrap_or(1),
        )
        .map_err(|error| Error::Compile(format!("invalid Cell.rowSpan: {error}")))?;
        let column_span = u16::try_from(
            cell.props
                .get("columnSpan")
                .and_then(Value::as_u64)
                .unwrap_or(1),
        )
        .map_err(|error| Error::Compile(format!("invalid Cell.columnSpan: {error}")))?;
        let format = build_cell_format(&cell)?;
        if row_span > 1 || column_span > 1 {
            let last_row = row_number + row_span - 1;
            let last_column = column_index + column_span - 1;
            worksheet
                .merge_range(row_number, column_index, last_row, last_column, "", &format)
                .map_err(|error| Error::Compile(format!("cannot merge cell range: {error}")))?;
            merged_ranges.push((row_number, column_index, last_row, last_column));
        }
        write_cell(worksheet, row_number, column_index, &cell, &format)?;
        if let Some(range) = cell
            .props
            .get("arrayFormulaRange")
            .and_then(Value::as_str)
            .and_then(xlsx_jsx::parse_a1_range)
        {
            merged_ranges.push((range.0, range.1, range.2, range.3));
        }
        column_index = column_index
            .checked_add(column_span)
            .ok_or_else(|| Error::Compile("worksheet exceeds column limit".into()))?;
    }
    Ok(())
}

fn apply_freeze_panes(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    let rows = sheet
        .props
        .get("freezeRows")
        .and_then(Value::as_u64)
        .unwrap_or_default();
    let columns = sheet
        .props
        .get("freezeColumns")
        .and_then(Value::as_u64)
        .unwrap_or_default();
    if rows == 0 && columns == 0 {
        return Ok(());
    }
    worksheet
        .set_freeze_panes(
            u32::try_from(rows).map_err(|error| Error::Compile(error.to_string()))?,
            u16::try_from(columns).map_err(|error| Error::Compile(error.to_string()))?,
        )
        .map_err(|error| Error::Compile(error.to_string()))?;
    if let Some(reference) = sheet.props.get("freezeTopLeftCell").and_then(Value::as_str) {
        let (row, column) = parse_a1_cell_reference(reference).ok_or_else(|| {
            Error::Compile(format!(
                "invalid Worksheet.freezeTopLeftCell `{reference}`; use an A1 reference such as `C4`"
            ))
        })?;
        worksheet
            .set_freeze_panes_top_cell(row, column)
            .map_err(|error| Error::Compile(error.to_string()))?;
    }
    Ok(())
}

fn apply_worksheet_selection(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &Node,
) -> xlsx_jsx::Result<()> {
    if let Some(reference) = sheet.props.get("topLeftCell").and_then(Value::as_str) {
        let (row, column) = parse_a1_cell_reference(reference).ok_or_else(|| {
            Error::Compile(format!(
                "invalid Worksheet.topLeftCell `{reference}`; use an A1 cell such as `H6`"
            ))
        })?;
        worksheet
            .set_top_left_cell(row, column)
            .map_err(|error| Error::Compile(format!("invalid Worksheet.topLeftCell: {error}")))?;
    }
    let Some(selection) = sheet.props.get("selection") else {
        return Ok(());
    };
    let active = selection
        .get("activeCell")
        .and_then(Value::as_str)
        .and_then(parse_a1_cell_reference)
        .ok_or_else(|| Error::Compile("invalid Worksheet.selection.activeCell".into()))?;
    let range_text = selection
        .get("range")
        .and_then(Value::as_str)
        .unwrap_or_default();
    let range = parse_a1_range(range_text)
        .or_else(|| {
            parse_a1_cell_reference(range_text).map(|(row, column)| (row, column, row, column))
        })
        .ok_or_else(|| Error::Compile("invalid Worksheet.selection.range".into()))?;
    let last_row = if active.0 == range.0 {
        range.2
    } else {
        range.0
    };
    let last_column = if active.1 == range.1 {
        range.3
    } else {
        range.1
    };
    worksheet
        .set_selection(active.0, active.1, last_row, last_column)
        .map(|_| ())
        .map_err(|error| Error::Compile(format!("invalid Worksheet.selection: {error}")))
}

#[expect(
    clippy::too_many_lines,
    reason = "Cell values, formulas, dynamic/legacy arrays, and links share one mutually exclusive writer dispatch."
)]
fn write_cell(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    cell: &Node,
    format: &Format,
) -> xlsx_jsx::Result<()> {
    if cell.children.iter().any(|child| {
        child
            .get("type")
            .and_then(Value::as_str)
            .is_some_and(|kind| kind == "TextRun")
    }) {
        return write_rich_text_cell(worksheet, row, col, cell, format);
    }
    if let Some(target) = cell.props.get("hyperlink").and_then(Value::as_str) {
        let label = cell
            .props
            .get("value")
            .or_else(|| cell.children.first())
            .and_then(Value::as_str)
            .ok_or_else(|| {
                Error::Compile("Cell.hyperlink requires a string value or child label".into())
            })?;
        return worksheet
            .write_url_with_format(
                row,
                col,
                Url::new(target)
                    .set_text(label)
                    .set_tip(
                        cell.props
                            .get("hyperlinkTooltip")
                            .and_then(Value::as_str)
                            .unwrap_or_default(),
                    ),
                format,
            )
            .map(|_| ())
            .map_err(|error| {
                Error::Compile(format!(
                    "cannot write Cell.hyperlink `{target}`; use a valid external URI or internal: workbook location and a text label: {error}"
                ))
            });
    }
    if let Some(date) = cell.props.get("date").and_then(Value::as_str) {
        let datetime = rust_xlsxwriter::ExcelDateTime::parse_from_str(date)
            .map_err(|error| Error::Compile(format!("invalid date value `{date}`: {error}")))?;
        worksheet
            .write_datetime_with_format(row, col, &datetime, format)
            .map_err(|error| Error::Compile(error.to_string()))?;
        return Ok(());
    }
    if let Some(formula) = cell.props.get("formula").and_then(Value::as_str) {
        let formula = if formula.starts_with('=') {
            formula.to_owned()
        } else {
            format!("={formula}")
        };
        if cell.props.get("dynamicArray").and_then(Value::as_bool) == Some(true) {
            if let Some(range) = cell
                .props
                .get("arrayFormulaRange")
                .and_then(Value::as_str)
                .and_then(xlsx_jsx::parse_a1_range)
            {
                worksheet
                    .write_dynamic_array_formula_with_format(
                        range.0,
                        range.1,
                        range.2,
                        range.3,
                        formula.as_str(),
                        format,
                    )
                    .map_err(|error| {
                        Error::Compile(format!("cannot write dynamic array formula: {error}"))
                    })?;
            } else {
                worksheet
                    .write_dynamic_formula_with_format(row, col, formula.as_str(), format)
                    .map_err(|error| {
                        Error::Compile(format!("cannot write dynamic formula: {error}"))
                    })?;
            }
            return Ok(());
        }
        if let Some(range) = cell
            .props
            .get("arrayFormulaRange")
            .and_then(Value::as_str)
            .and_then(xlsx_jsx::parse_a1_range)
        {
            worksheet
                .write_array_formula_with_format(
                    range.0,
                    range.1,
                    range.2,
                    range.3,
                    formula.as_str(),
                    format,
                )
                .map_err(|error| Error::Compile(format!("cannot write array formula: {error}")))?;
            return Ok(());
        }
        worksheet
            .write_formula_with_format(row, col, formula.as_str(), format)
            .map_err(|error| Error::Compile(error.to_string()))?;
        write_formula_result(worksheet, row, col, cell)?;
        return Ok(());
    }

    let value = cell
        .props
        .get("value")
        .or_else(|| cell.children.first())
        .unwrap_or(&Value::Null);
    match value {
        Value::String(text) => worksheet.write_string_with_format(row, col, text, format),
        Value::Number(number) => {
            let number = number
                .as_f64()
                .ok_or_else(|| Error::Compile("invalid numeric cell".into()))?;
            worksheet.write_number_with_format(row, col, number, format)
        }
        Value::Bool(boolean) => worksheet.write_boolean_with_format(row, col, *boolean, format),
        Value::Null => worksheet.write_blank(row, col, format),
        _ => return Err(Error::Compile("cell values must be scalar".into())),
    }
    .map(|_| ())
    .map_err(|error| Error::Compile(error.to_string()))
}

fn write_formula_result(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    cell: &Node,
) -> xlsx_jsx::Result<()> {
    let Some(result) = cell.props.get("formulaResult") else {
        return Ok(());
    };
    let result = match result {
        Value::Null => String::new(),
        Value::Bool(value) => value.to_string().to_ascii_uppercase(),
        Value::Number(value) => value.to_string(),
        Value::String(value) => value.clone(),
        Value::Array(_) | Value::Object(_) => {
            return Err(Error::Compile(
                "Cell.formulaResult must be a scalar cached result".into(),
            ));
        }
    };
    worksheet.set_formula_result(row, col, result);
    Ok(())
}

fn write_rich_text_cell(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    cell: &Node,
    cell_format: &Format,
) -> xlsx_jsx::Result<()> {
    let runs = cell
        .children
        .iter()
        .map(|value| {
            serde_json::from_value::<Node>(value.clone())
                .map_err(|error| Error::Compile(format!("invalid TextRun: {error}")))
        })
        .collect::<xlsx_jsx::Result<Vec<_>>>()?;
    let formats = runs
        .iter()
        .map(build_cell_format)
        .collect::<xlsx_jsx::Result<Vec<_>>>()?;
    let segments = runs
        .iter()
        .zip(&formats)
        .map(|(run, format)| {
            let text = run
                .children
                .first()
                .and_then(Value::as_str)
                .unwrap_or_default();
            (format, text)
        })
        .collect::<Vec<_>>();
    worksheet
        .write_rich_string_with_format(row, col, &segments, cell_format)
        .map(|_| ())
        .map_err(|error| {
            Error::Compile(format!(
                "cannot write rich text at row {}, column {}: {error}; ensure runs are non-empty and their combined text is at most 32,767 characters",
                row + 1,
                u32::from(col) + 1
            ))
        })
}

fn build_cell_format(cell: &Node) -> xlsx_jsx::Result<Format> {
    let mut format = rust_xlsxwriter::Format::new();
    if cell.props.get("checkbox").and_then(Value::as_bool) == Some(true) {
        format = format.set_checkbox();
    }
    if cell.props.get("bold").and_then(Value::as_bool) == Some(true) {
        format = format.set_bold();
    }
    if cell.props.get("italic").and_then(Value::as_bool) == Some(true) {
        format = format.set_italic();
    }
    if cell.props.get("quotePrefix").and_then(Value::as_bool) == Some(true) {
        format = format.set_quote_prefix();
    }
    if cell.props.get("locked").and_then(Value::as_bool) == Some(false) {
        format = format.set_unlocked();
    }
    if cell.props.get("formulaHidden").and_then(Value::as_bool) == Some(true) {
        format = format.set_hidden();
    }
    format = apply_cell_font_decorations(format, cell);
    format = apply_cell_fill(format, cell)?;
    if let Some(font_color) = cell.props.get("fontColor").and_then(Value::as_str) {
        format = format.set_font_color(font_color);
    }
    if let Some(font_name) = cell.props.get("fontName").and_then(Value::as_str) {
        format = format.set_font_name(font_name);
    }
    if let Some(font_size) = cell.props.get("fontSize").and_then(Value::as_f64) {
        format = format.set_font_size(font_size);
    }
    if let Some(font_family) = cell.props.get("fontFamily").and_then(Value::as_u64) {
        let font_family = u8::try_from(font_family)
            .map_err(|error| Error::Compile(format!("invalid Cell.fontFamily: {error}")))?;
        format = format.set_font_family(font_family);
    }
    if let Some(font_charset) = cell.props.get("fontCharset").and_then(Value::as_u64) {
        let font_charset = u8::try_from(font_charset)
            .map_err(|error| Error::Compile(format!("invalid Cell.fontCharset: {error}")))?;
        format = format.set_font_charset(font_charset);
    }
    format = apply_font_scheme(format, cell);
    format = apply_cell_borders(format, cell);
    let explicit_number_format = cell.props.get("numberFormat").and_then(Value::as_str);
    if let Some(horizontal_align) = cell.props.get("horizontalAlign").and_then(Value::as_str) {
        let align = match horizontal_align {
            "left" => FormatAlign::Left,
            "center" => FormatAlign::Center,
            "right" => FormatAlign::Right,
            "distributed" => FormatAlign::Distributed,
            _ => FormatAlign::General,
        };
        format = format.set_align(align);
    }
    if let Some(vertical_align) = cell.props.get("verticalAlign").and_then(Value::as_str) {
        let align = match vertical_align {
            "top" => FormatAlign::Top,
            "center" => FormatAlign::VerticalCenter,
            "bottom" => FormatAlign::Bottom,
            _ => FormatAlign::General,
        };
        format = format.set_align(align);
    }
    if let Some(reading_direction) = cell.props.get("readingDirection").and_then(Value::as_str) {
        let direction = match reading_direction {
            "leftToRight" => 1,
            "rightToLeft" => 2,
            _ => 0,
        };
        format = format.set_reading_direction(direction);
    }
    if cell.props.get("wrapText").and_then(Value::as_bool) == Some(true) {
        format = format.set_text_wrap();
    }
    if cell.props.get("shrinkToFit").and_then(Value::as_bool) == Some(true) {
        format = format.set_shrink();
    }
    if let Some(rotation) = cell.props.get("textRotation").and_then(Value::as_i64) {
        let rotation = i16::try_from(rotation)
            .map_err(|error| Error::Compile(format!("invalid Cell.textRotation: {error}")))?;
        format = format.set_rotation(rotation);
    }
    format = apply_cell_indent(format, cell)?;
    if let Some(date) = cell.props.get("date").and_then(Value::as_str) {
        let default_number_format = if date.contains('T') {
            if date.contains('.') {
                "yyyy-mm-dd hh:mm:ss.000"
            } else {
                "yyyy-mm-dd hh:mm:ss"
            }
        } else {
            "yyyy-mm-dd"
        };
        format = format.set_num_format(explicit_number_format.unwrap_or(default_number_format));
    } else if let Some(number_format) = explicit_number_format {
        format = format.set_num_format(number_format);
    }
    Ok(format)
}

fn apply_font_scheme(mut format: Format, cell: &Node) -> Format {
    let Some(font_scheme) = cell.props.get("fontScheme").and_then(Value::as_str) else {
        return format;
    };
    format = match font_scheme {
        "none" => format.set_font_scheme(rust_xlsxwriter::FontScheme::None),
        "body" => format.set_font_scheme(rust_xlsxwriter::FontScheme::Body),
        "headings" => format.set_font_scheme(rust_xlsxwriter::FontScheme::Headings),
        _ => format,
    };
    format
}

fn apply_cell_fill(mut format: Format, cell: &Node) -> xlsx_jsx::Result<Format> {
    if let Some(color) = cell.props.get("fillColor").and_then(Value::as_str) {
        return Ok(format.set_background_color(color));
    }
    let Some(pattern_name) = cell.props.get("fillPattern").and_then(Value::as_str) else {
        return Ok(format);
    };
    let pattern = match pattern_name {
        "none" => FormatPattern::None,
        "solid" => FormatPattern::Solid,
        "mediumGray" => FormatPattern::MediumGray,
        "darkGray" => FormatPattern::DarkGray,
        "lightGray" => FormatPattern::LightGray,
        "darkHorizontal" => FormatPattern::DarkHorizontal,
        "darkVertical" => FormatPattern::DarkVertical,
        "darkDown" => FormatPattern::DarkDown,
        "darkUp" => FormatPattern::DarkUp,
        "darkGrid" => FormatPattern::DarkGrid,
        "darkTrellis" => FormatPattern::DarkTrellis,
        "lightHorizontal" => FormatPattern::LightHorizontal,
        "lightVertical" => FormatPattern::LightVertical,
        "lightDown" => FormatPattern::LightDown,
        "lightUp" => FormatPattern::LightUp,
        "lightGrid" => FormatPattern::LightGrid,
        "lightTrellis" => FormatPattern::LightTrellis,
        "gray125" => FormatPattern::Gray125,
        "gray0625" => FormatPattern::Gray0625,
        _ => {
            return Err(Error::Compile(format!(
                "unsupported fillPattern `{pattern_name}`; use a supported xlsx-jsx spec pattern"
            )));
        }
    };
    format = format.set_pattern(pattern);
    if let Some(color) = cell
        .props
        .get("fillForegroundColor")
        .and_then(Value::as_str)
    {
        format = format.set_foreground_color(color);
    }
    if let Some(color) = cell
        .props
        .get("fillBackgroundColor")
        .and_then(Value::as_str)
    {
        format = format.set_background_color(color);
    }
    Ok(format)
}

fn has_format_properties(node: &Node) -> bool {
    node.props
        .keys()
        .any(|property| FORMAT_PROPERTY_NAMES.contains(&property.as_str()))
}

fn apply_cell_indent(
    format: rust_xlsxwriter::Format,
    cell: &Node,
) -> xlsx_jsx::Result<rust_xlsxwriter::Format> {
    let Some(indent) = cell.props.get("indent").and_then(Value::as_u64) else {
        return Ok(format);
    };
    let indent = u8::try_from(indent).map_err(|_| {
        Error::Compile("Cell.indent exceeds the XLSX writer's supported range".into())
    })?;
    Ok(format.set_indent(indent))
}

fn apply_cell_font_decorations(mut format: Format, cell: &Node) -> Format {
    if cell.props.get("strikethrough").and_then(Value::as_bool) == Some(true) {
        format = format.set_font_strikethrough();
    }
    if let Some(underline) = cell.props.get("underline").and_then(Value::as_str) {
        let underline = match underline {
            "single" => FormatUnderline::Single,
            "double" => FormatUnderline::Double,
            "singleAccounting" => FormatUnderline::SingleAccounting,
            "doubleAccounting" => FormatUnderline::DoubleAccounting,
            _ => FormatUnderline::None,
        };
        format = format.set_underline(underline);
    }
    if let Some(script) = cell.props.get("script").and_then(Value::as_str) {
        let script = match script {
            "superscript" => FormatScript::Superscript,
            "subscript" => FormatScript::Subscript,
            _ => FormatScript::None,
        };
        format = format.set_font_script(script);
    }
    format
}

fn parse_ignored_errors(xml: &str, part: &str) -> xlsx_jsx::Result<Vec<WorksheetIgnoredError>> {
    let mut reader = XmlReader::from_str(xml);
    let mut results = Vec::new();
    let mut ranges = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event))
                if event.local_name().as_ref() == b"ignoredErrors" =>
            {
                if event.attributes().with_checks(false).next().is_some() {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` ignoredErrors container has unsupported attributes; remove them or add explicit JSX support"
                    )));
                }
            }
            Ok(Event::Start(event) | Event::Empty(event))
                if event.local_name().as_ref() == b"ignoredError" =>
            {
                let mut range = None;
                let mut found = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| {
                        Error::Reverse(format!(
                            "worksheet `{part}` has invalid ignoredError metadata: {error}"
                        ))
                    })?;
                    let key =
                        String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    if key == "sqref" {
                        if range.is_some() {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` ignoredError repeats sqref; keep exactly one sqref attribute"
                            )));
                        }
                        range = Some(value);
                        continue;
                    }
                    let error_type = match key.as_str() {
                        "numberStoredAsText" => "numberStoredAsText",
                        "evalError" => "evalError",
                        "formula" => "formula",
                        "emptyCellReference" => "emptyCellReference",
                        "formulaRange" => "formulaRange",
                        "listDataValidation" => "listDataValidation",
                        "TwoDigitTextYear" => "TwoDigitTextYear",
                        "unlockedFormula" => "unlockedFormula",
                        "calculatedColumn" => "calculatedColumn",
                        _ => {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` ignoredError has unsupported `{key}` metadata; add matching JSX support or remove it"
                            )));
                        }
                    };
                    if value != "1" || found.replace(error_type).is_some() {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` ignoredError must contain exactly one true error flag; repair its flags"
                        )));
                    }
                }
                let ranges_value = range.ok_or_else(|| {
                    Error::Reverse(format!(
                        "worksheet `{part}` ignoredError is missing sqref; add one or more A1 ranges"
                    ))
                })?;
                let error_type = found.ok_or_else(|| {
                    Error::Reverse(format!(
                        "worksheet `{part}` ignoredError has no recognized error flag"
                    ))
                })?;
                for range in ranges_value.split_whitespace() {
                    let parsed = parse_a1_range(range)
                        .or_else(|| {
                            parse_a1_cell_reference(range)
                                .map(|(row, column)| (row, column, row, column))
                        })
                        .ok_or_else(|| Error::Reverse(format!("worksheet `{part}` ignoredError sqref `{range}` is invalid; repair it to an ordered A1 rectangle")))?;
                    if ranges.iter().any(|prior| ranges_overlap(parsed, *prior)) {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` has overlapping ignored-error ranges; separate the ranges because one cell cannot have competing ignore flags"
                        )));
                    }
                    ranges.push(parsed);
                    let range = if range.contains(':') {
                        range.to_owned()
                    } else {
                        format!("{range}:{range}")
                    };
                    results.push(WorksheetIgnoredError { range, error_type });
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid XML in `{part}`: {error}"))),
        }
    }
    Ok(results)
}

#[expect(
    clippy::too_many_lines,
    reason = "Reverse rendering keeps workbook metadata and worksheet ordering in one pass."
)]
fn reverse(path: &std::path::Path) -> xlsx_jsx::Result<String> {
    let (metadata_by_sheet, workbook_properties, formats) = audit_reverse_subset(path)?;
    let mut workbook =
        open_workbook_auto(path).map_err(|error| Error::Reverse(error.to_string()))?;
    let sheet_names = workbook.sheet_names().clone();
    let mut output = String::from(
        "import { Workbook, Worksheet, DefinedName, Column, DataValidation, ConditionalFormat, IgnoreError, Table, TableColumn, Note, Image, Chart, Series, Row, Cell, TextRun } from \"xlsx-jsx\";\n\nexport default <Workbook",
    );
    for (name, value) in [
        ("title", workbook_properties.title.as_deref()),
        ("subject", workbook_properties.subject.as_deref()),
        ("author", workbook_properties.author.as_deref()),
        ("keywords", workbook_properties.keywords.as_deref()),
        ("description", workbook_properties.description.as_deref()),
    ] {
        if let Some(value) = value {
            let _ = write!(output, " {name}={value:?}");
        }
    }
    output.push_str(">\n");
    for defined_name in &workbook_properties.defined_names {
        let _ = writeln!(
            output,
            "  <DefinedName name={{{:?}}} formula={{{:?}}} />",
            defined_name.name, defined_name.formula
        );
    }
    for name in sheet_names {
        let sheet_metadata = metadata_by_sheet.get(&name);
        let _ = write!(output, "  <Worksheet name={{{name:?}}}");
        if sheet_metadata.is_some_and(|metadata| metadata.active_worksheet.is_some()) {
            output.push_str(" active={true}");
        }
        if let Some(selected) = sheet_metadata.and_then(|metadata| metadata.selected) {
            let _ = write!(output, " selected={{{selected}}}");
        }
        if sheet_metadata.is_some_and(|metadata| metadata.first_tab.is_some()) {
            output.push_str(" firstTab={true}");
        }
        if let Some(visibility) =
            sheet_metadata.and_then(|metadata| metadata.visibility.jsx_value())
        {
            let _ = write!(output, " visibility={visibility:?}");
        }
        if let Some(color) = sheet_metadata.and_then(|metadata| metadata.tab_color.as_deref()) {
            let _ = write!(output, " tabColor={color:?}");
        }
        if let Some(height) = sheet_metadata.and_then(|metadata| metadata.default_row_height_pixels)
        {
            let _ = write!(output, " defaultRowHeightPixels={{{height}}}");
        }
        if sheet_metadata.is_some_and(|metadata| metadata.hide_unused_rows) {
            output.push_str(" hideUnusedRows={true}");
        }
        if sheet_metadata.is_some_and(|metadata| metadata.hide_screen_gridlines) {
            output.push_str(" showGridlines={false}");
        }
        if sheet_metadata.is_some_and(|metadata| metadata.right_to_left.is_some()) {
            output.push_str(" rightToLeft={true}");
        }
        if let Some(view) = sheet_metadata.and_then(|metadata| metadata.view.jsx_value()) {
            let _ = write!(output, " view=\"{view}\"");
        }
        append_print_option_props(&mut output, sheet_metadata);
        append_page_margins_props(&mut output, sheet_metadata);
        append_print_scaling_props(&mut output, sheet_metadata);
        append_print_range_props(&mut output, sheet_metadata);
        append_header_footer_props(&mut output, sheet_metadata);
        append_worksheet_protection_props(&mut output, sheet_metadata)?;
        if let Some(zoom) = sheet_metadata
            .and_then(|metadata| metadata.zoom)
            .filter(|zoom| *zoom != 100)
        {
            let _ = write!(output, " zoom={{{zoom}}}");
        }
        if let Some(top_left_cell) =
            sheet_metadata.and_then(|metadata| metadata.top_left_cell.as_deref())
        {
            let _ = write!(output, " topLeftCell={top_left_cell:?}");
        }
        if let Some(selection) = sheet_metadata.and_then(|metadata| metadata.selection.as_ref()) {
            let _ = write!(
                output,
                " selection={{{{activeCell: {:?}, range: {:?}}}}}",
                selection.active_cell, selection.range
            );
        }
        append_page_break_props(&mut output, sheet_metadata);
        if sheet_metadata.is_some_and(|metadata| metadata.summary_below == Some(false)) {
            output.push_str(" summaryBelow={false}");
        }
        if sheet_metadata.is_some_and(|metadata| metadata.summary_right == Some(false)) {
            output.push_str(" summaryRight={false}");
        }
        if let Some(auto_filter) =
            sheet_metadata.and_then(|metadata| metadata.auto_filter.as_deref())
        {
            let _ = write!(output, " autoFilter={{{auto_filter:?}}}");
        }
        if let Some(columns) = sheet_metadata
            .map(|metadata| &metadata.auto_filter_columns)
            .filter(|columns| !columns.is_empty())
        {
            let serialized = serde_json::to_string(columns).map_err(|error| {
                Error::Reverse(format!("cannot serialize filter criteria: {error}"))
            })?;
            let _ = write!(output, " autoFilterColumns={{{serialized}}}");
        }
        if let Some(panes) = sheet_metadata.and_then(|metadata| metadata.freeze_panes) {
            if panes.rows > 0 {
                let _ = write!(output, " freezeRows={{{}}}", panes.rows);
            }
            if panes.columns > 0 {
                let _ = write!(output, " freezeColumns={{{}}}", panes.columns);
            }
            let default_top_left = format_a1_cell_reference(panes.rows, panes.columns);
            let top_left = format_a1_cell_reference(panes.top_left_row, panes.top_left_column);
            if top_left != default_top_left {
                let _ = write!(output, " freezeTopLeftCell={{{top_left:?}}}");
            }
        }
        output.push_str(">\n");
        if let Some(defined_names) = sheet_metadata.map(|metadata| &metadata.defined_names) {
            for defined_name in defined_names {
                let _ = writeln!(
                    output,
                    "    <DefinedName name={{{:?}}} formula={{{:?}}} />",
                    defined_name.name, defined_name.formula
                );
            }
        }
        if let Some(columns) = sheet_metadata.map(|metadata| &metadata.columns) {
            render_columns(&mut output, columns, &formats);
        }
        if let Some(validations) = sheet_metadata.map(|metadata| &metadata.data_validations) {
            render_data_validations(&mut output, validations);
        }
        if let Some(formats) = sheet_metadata.map(|metadata| &metadata.conditional_formats) {
            render_conditional_formats(&mut output, formats);
        }
        if let Some(ignored_errors) = sheet_metadata.map(|metadata| &metadata.ignored_errors) {
            render_ignored_errors(&mut output, ignored_errors);
        }
        if let Some(tables) = sheet_metadata.map(|metadata| &metadata.tables) {
            render_tables(&mut output, tables);
        }
        if let Some(notes) = sheet_metadata.map(|metadata| &metadata.notes) {
            render_notes(&mut output, notes);
        }
        if let Some(metadata) = sheet_metadata {
            render_drawings(&mut output, metadata);
        }
        render_worksheet_rows(&mut output, &mut workbook, &name, sheet_metadata, &formats)?;
        output.push_str("  </Worksheet>\n");
    }
    output.push_str("</Workbook>;\n");
    Ok(output)
}

fn append_page_break_props(output: &mut String, metadata: Option<&WorksheetMetadata>) {
    let Some(page_breaks) = metadata.map(|metadata| &metadata.page_breaks) else {
        return;
    };
    if !page_breaks.horizontal.is_empty() {
        let _ = write!(
            output,
            " horizontalPageBreaks={{{:?}}}",
            page_breaks.horizontal
        );
    }
    if !page_breaks.vertical.is_empty() {
        let _ = write!(output, " verticalPageBreaks={{{:?}}}", page_breaks.vertical);
    }
}

fn append_page_margins_props(output: &mut String, metadata: Option<&WorksheetMetadata>) {
    let Some(margins) = metadata.map(|metadata| metadata.print.page_margins) else {
        return;
    };
    let values = [
        ("left", margins.left, 0.7_f64),
        ("right", margins.right, 0.7),
        ("top", margins.top, 0.75),
        ("bottom", margins.bottom, 0.75),
        ("header", margins.header, 0.3),
        ("footer", margins.footer, 0.3),
    ];
    let mut wrote_property = false;
    for (name, value, default) in values {
        if value.to_bits() == default.to_bits() {
            continue;
        }
        if wrote_property {
            output.push_str(", ");
        } else {
            output.push_str(" pageMargins={{");
            wrote_property = true;
        }
        let _ = write!(output, "{name}: {value:?}");
    }
    if wrote_property {
        output.push_str("}}");
    }
}

fn append_print_scaling_props(output: &mut String, metadata: Option<&WorksheetMetadata>) {
    let Some(print) = metadata.map(|metadata| &metadata.print) else {
        return;
    };
    if let Some(orientation) = print.orientation {
        let _ = write!(output, " orientation=\"{orientation}\"");
    }
    if let Some(paper_size) = print.paper_size.filter(|paper_size| *paper_size != 0) {
        let _ = write!(output, " paperSize={{{paper_size}}}");
    }
    if let Some(scale) = print.print_scale.filter(|scale| *scale != 100) {
        let _ = write!(output, " printScale={{{scale}}}");
    }
    if print.page_order == Some("overThenDown") {
        output.push_str(" pageOrder=\"overThenDown\"");
    }
    if let Some(fit) = print.fit_to_pages {
        let _ = write!(
            output,
            " fitToPages={{{{width: {}, height: {}}}}}",
            fit.width, fit.height
        );
    }
}

fn append_print_range_props(output: &mut String, metadata: Option<&WorksheetMetadata>) {
    let Some(metadata) = metadata else {
        return;
    };
    if let Some(print_area) = &metadata.print_area {
        let _ = write!(output, " printArea={print_area:?}");
    }
    if let Some(repeat_rows) = &metadata.repeat_rows {
        let _ = write!(output, " repeatRows={repeat_rows:?}");
    }
    if let Some(repeat_columns) = &metadata.repeat_columns {
        let _ = write!(output, " repeatColumns={repeat_columns:?}");
    }
}

fn append_header_footer_props(output: &mut String, metadata: Option<&WorksheetMetadata>) {
    let Some(header_footer) = metadata.map(|metadata| &metadata.header_footer) else {
        return;
    };
    if let Some(header) = &header_footer.header {
        let _ = write!(output, " header={header:?}");
    }
    if let Some(footer) = &header_footer.footer {
        let _ = write!(output, " footer={footer:?}");
    }
    if header_footer.scale_with_doc == Some(false) {
        output.push_str(" headerFooterScaleWithDoc={false}");
    }
    if header_footer.align_with_page == Some(false) {
        output.push_str(" headerFooterAlignWithPage={false}");
    }
}

fn append_worksheet_protection_props(
    output: &mut String,
    metadata: Option<&WorksheetMetadata>,
) -> xlsx_jsx::Result<()> {
    let Some(protection) = metadata.and_then(|metadata| metadata.protection.as_ref()) else {
        return Ok(());
    };
    let serialized = serde_json::to_string(protection).map_err(|error| {
        Error::Reverse(format!("cannot serialize worksheet protection: {error}"))
    })?;
    let _ = write!(output, " protection={{{serialized}}}");
    Ok(())
}

fn append_print_option_props(output: &mut String, metadata: Option<&WorksheetMetadata>) {
    let Some(print) = metadata.map(|metadata| &metadata.print) else {
        return;
    };
    for (property, enabled) in [
        ("centerHorizontally", print.options.centering.horizontally),
        ("centerVertically", print.options.centering.vertically),
        ("printBlackAndWhite", print.quality.black_and_white),
        ("printDraft", print.quality.draft),
    ] {
        if enabled {
            let _ = write!(output, " {property}={{true}}");
        }
    }
    if print.options.gridlines {
        output.push_str(" printGridlines={true}");
    }
    if print.options.headings {
        output.push_str(" printHeadings={true}");
    }
}

fn render_worksheet_rows(
    output: &mut String,
    workbook: &mut calamine::Sheets<impl std::io::Read + std::io::Seek>,
    sheet_name: &str,
    metadata: Option<&WorksheetMetadata>,
    formats: &[CellFontStyle],
) -> xlsx_jsx::Result<()> {
    let range = workbook
        .worksheet_range(sheet_name)
        .map_err(|error| Error::Reverse(error.to_string()))?;
    let formula_range = workbook
        .worksheet_formula(sheet_name)
        .map_err(|error| Error::Reverse(error.to_string()))?;
    let merged_ranges = metadata.map_or(&[][..], |data| data.merged_ranges.as_slice());
    validate_array_formula_cached_results(&range, metadata, sheet_name)?;
    validate_merged_values(&range, &formula_range, merged_ranges)?;
    if let Some(metadata) = metadata {
        for table in &metadata.tables {
            let coordinates = parse_a1_range(&table.range).ok_or_else(|| {
                Error::Reverse(format!(
                    "table `{}` has invalid range `{}`",
                    table.name, table.range
                ))
            })?;
            let expected_columns = usize::from(coordinates.3 - coordinates.1 + 1);
            if table.headers.len() != expected_columns {
                return Err(Error::Reverse(format!(
                    "table `{}` column count does not match range {}; repair tableColumns",
                    table.name, table.range
                )));
            }
            for (offset, header) in table.headers.iter().enumerate() {
                let col = u32::from(coordinates.1) + u32::try_from(offset).unwrap_or_default();
                if formula_range
                    .get_value((coordinates.0, col))
                    .is_some_and(|formula| !formula.is_empty())
                {
                    return Err(Error::Reverse(format!(
                        "table `{}` header at {} is a formula; use a literal string header because formula headers cannot be represented losslessly",
                        table.name,
                        cell_reference(coordinates.0, col)
                    )));
                }
                let actual = range.get_value((coordinates.0, col));
                if !matches!(actual, Some(Data::String(value)) if value == header) {
                    return Err(Error::Reverse(format!(
                        "table `{}` header at {} does not match its worksheet cell; align the table column name and cell value",
                        table.name,
                        cell_reference(coordinates.0, col)
                    )));
                }
            }
        }
    }
    let bounds = range.start().zip(range.end());
    let last_row = bounds
        .map(|(_, end)| end.0)
        .into_iter()
        .chain(merged_ranges.iter().map(|merged| merged.last_row))
        .chain(metadata.into_iter().flat_map(|data| {
            data.array_formula_ranges
                .values()
                .map(|formula| formula.coordinates.2)
        }))
        .chain(metadata.and_then(|data| data.rows.keys().max().copied()))
        .max();
    let Some(last_row) = last_row else {
        return Ok(());
    };
    let context = WorksheetRenderContext {
        range: &range,
        formula_range: &formula_range,
        merged_ranges,
        metadata,
        sheet_name,
    };
    for row_index in 0..=last_row {
        let row_properties = metadata.and_then(|data| data.rows.get(&row_index));
        render_row_open(output, row_properties, formats);
        context.render_row(output, row_index, bounds)?;
        output.push_str("</Row>\n");
    }
    Ok(())
}

fn render_tables(output: &mut String, tables: &[WorksheetTable]) {
    for table in tables {
        let _ = write!(
            output,
            "    <Table range={{{:?}}} name={{{:?}}}",
            table.range, table.name
        );
        if table.style != "TableStyleMedium9" {
            let _ = write!(output, " style={{{:?}}}", table.style);
        }
        if !table.auto_filter {
            output.push_str(" autoFilter={false}");
        }
        if !table.banded_rows {
            output.push_str(" bandedRows={false}");
        }
        if table.banded_columns {
            output.push_str(" bandedColumns={true}");
        }
        if table.first_column {
            output.push_str(" firstColumn={true}");
        }
        if table.last_column {
            output.push_str(" lastColumn={true}");
        }
        if table.totals_row {
            output.push_str(" totalsRow={true}");
        }
        if let Some(value) = &table.alt_text {
            let _ = write!(output, " altText={{{value:?}}}");
        }
        if let Some(value) = &table.alt_text_title {
            let _ = write!(output, " altTextTitle={{{value:?}}}");
        }
        if table.columns.iter().all(WorksheetTableColumn::is_empty) {
            output.push_str(" />\n");
        } else {
            output.push_str(">\n");
            for column in &table.columns {
                output.push_str("      <TableColumn");
                if let Some(value) = &column.total_label {
                    let _ = write!(output, " totalLabel={{{value:?}}}");
                }
                if let Some(value) = &column.total_function {
                    let _ = write!(output, " totalFunction={{{value:?}}}");
                }
                if let Some(value) = &column.total_formula {
                    let _ = write!(output, " totalFormula={{{value:?}}}");
                }
                if let Some(value) = &column.formula {
                    let _ = write!(output, " formula={{{value:?}}}");
                }
                if let Some(style) = &column.format {
                    let serialized = serde_json::to_string(&table_column_format_value(style))
                        .expect("table differential format is JSON-serializable");
                    output.push_str(" format={");
                    output.push_str(&serialized);
                    output.push('}');
                }
                output.push_str(" />\n");
            }
            output.push_str("    </Table>\n");
        }
    }
}

fn render_notes(output: &mut String, notes: &[WorksheetNote]) {
    for note in notes {
        let _ = write!(
            output,
            "    <Note cell={{{:?}}} text={{{:?}}}",
            note.cell, note.text
        );
        if note.author != "Author" {
            let _ = write!(output, " author={{{:?}}}", note.author);
        }
        if !note.author_prefix {
            output.push_str(" authorPrefix={false}");
        }
        if note.visible {
            output.push_str(" visible={true}");
        }
        if note.width != 128 {
            let _ = write!(output, " width={{{}}}", note.width);
        }
        if note.height != 74 {
            let _ = write!(output, " height={{{}}}", note.height);
        }
        if note.background_color != "#FFFFE1" {
            let _ = write!(output, " backgroundColor={{{:?}}}", note.background_color);
        }
        if note.font_name != "Tahoma" {
            let _ = write!(output, " fontName={{{:?}}}", note.font_name);
        }
        if (note.font_size - 8.0).abs() > f64::EPSILON {
            let _ = write!(output, " fontSize={{{}}}", note.font_size);
        }
        if note.font_family != 0 {
            let _ = write!(output, " fontFamily={{{}}}", note.font_family);
        }
        if !note.alt_text.is_empty() {
            let _ = write!(output, " altText={{{:?}}}", note.alt_text);
        }
        output.push_str(" />\n");
    }
}

fn render_drawings(output: &mut String, metadata: &WorksheetMetadata) {
    for drawing in &metadata.drawing_order {
        match drawing {
            WorksheetDrawingOrder::Chart(index) => render_chart(output, &metadata.charts[*index]),
            WorksheetDrawingOrder::Image(index) => render_image(output, &metadata.images[*index]),
        }
    }
}

fn render_chart(output: &mut String, chart: &WorksheetChart) {
    let _ = writeln!(
        output,
        "    <Chart type={:?} cell={{{:?}}} width={{{}}} height={{{}}} name={{{:?}}}{}{}{}{}{}{}{}{}>",
        chart.chart_type,
        chart.cell,
        chart.width,
        chart.height,
        chart.name,
        chart
            .alt_text
            .as_ref()
            .map(|value| format!(" altText={{{value:?}}}"))
            .unwrap_or_default(),
        chart
            .title
            .as_ref()
            .map(|value| format!(" title={{{value:?}}}"))
            .unwrap_or_default(),
        chart
            .x_axis_title
            .as_ref()
            .map(|value| format!(" xAxisTitle={{{value:?}}}"))
            .unwrap_or_default(),
        chart
            .y_axis_title
            .as_ref()
            .map(|value| format!(" yAxisTitle={{{value:?}}}"))
            .unwrap_or_default(),
        if chart.show_legend {
            ""
        } else {
            " showLegend={false}"
        },
        chart
            .legend_position
            .map(|value| format!(" legendPosition=\"{value}\""))
            .unwrap_or_default(),
        chart
            .hole_size
            .map(|value| format!(" holeSize={{{value}}}"))
            .unwrap_or_default(),
        chart
            .first_slice_angle
            .map(|value| format!(" firstSliceAngle={{{value}}}"))
            .unwrap_or_default()
    );
    for series in &chart.series {
        output.push_str("      <Series");
        let _ = write!(output, " values={{{:?}}}", series.values);
        if let Some(categories) = &series.categories {
            let _ = write!(output, " categories={{{categories:?}}}");
        }
        if let Some(name) = &series.name {
            let _ = write!(output, " name={{{name:?}}}");
        }
        if let Some(title) = &series.title {
            let _ = write!(output, " title={{{title:?}}}");
        }
        output.push_str(" />\n");
    }
    output.push_str("    </Chart>\n");
}

fn render_image(output: &mut String, image: &WorksheetImage) {
    let _ = write!(
        output,
        "    <Image cell={{{:?}}} width={{{}}} height={{{}}} data={{{:?}}}",
        image.cell, image.width, image.height, image.data_uri
    );
    if image.offset_x > 0 {
        let _ = write!(output, " offsetX={{{}}}", image.offset_x);
    }
    if image.offset_y > 0 {
        let _ = write!(output, " offsetY={{{}}}", image.offset_y);
    }
    if let Some(movement) = image.movement {
        let _ = write!(output, " movement=\"{movement}\"");
    }
    if let Some(alt_text) = &image.alt_text {
        let _ = write!(output, " altText={{{alt_text:?}}}");
    }
    if image.decorative {
        output.push_str(" decorative={true}");
    }
    output.push_str(" />\n");
}

fn table_column_format_value(style: &CellFontStyle) -> Value {
    let mut props = serde_json::Map::new();
    if style.bold {
        props.insert("bold".into(), Value::Bool(true));
    }
    if style.italic {
        props.insert("italic".into(), Value::Bool(true));
    }
    if style.strikethrough == Some(true) {
        props.insert("strikethrough".into(), Value::Bool(true));
    }
    if let Some(underline) = &style.underline {
        props.insert("underline".into(), Value::String(underline.clone()));
    }
    if let Some(script) = &style.script {
        props.insert("script".into(), Value::String(script.clone()));
    }
    if let Some(color) = &style.font_color {
        props.insert("fontColor".into(), Value::String(color.clone()));
    }
    if let Some(color) = &style.fill_color {
        props.insert("fillColor".into(), Value::String(color.clone()));
    }
    if let Some(pattern) = &style.fill_pattern {
        props.insert("fillPattern".into(), Value::String(pattern.clone()));
    }
    if let Some(color) = &style.fill_foreground_color {
        props.insert("fillForegroundColor".into(), Value::String(color.clone()));
    }
    if let Some(color) = &style.fill_background_color {
        props.insert("fillBackgroundColor".into(), Value::String(color.clone()));
    }
    if let Some(number_format) = &style.number_format {
        props.insert("numberFormat".into(), Value::String(number_format.clone()));
    }
    for side in BorderSide::ALL {
        let edge = &style.borders.edges[side.index()];
        if let Some(border_style) = edge.style {
            props.insert(
                format!("border{}", side.suffix()),
                Value::String(border_style_name(border_style).into()),
            );
        }
        if let Some(color) = &edge.color {
            props.insert(
                format!("border{}Color", side.suffix()),
                Value::String(color.clone()),
            );
        }
    }
    if let Some(direction) = &style.borders.diagonal_type {
        props.insert(
            "borderDiagonalType".into(),
            Value::String(direction.clone()),
        );
    }
    Value::Object(props)
}

fn validate_array_formula_cached_results(
    range: &calamine::Range<Data>,
    metadata: Option<&WorksheetMetadata>,
    sheet_name: &str,
) -> xlsx_jsx::Result<()> {
    let Some(metadata) = metadata else {
        return Ok(());
    };
    for array_formula in metadata.array_formula_ranges.values() {
        for row in array_formula.coordinates.0..=array_formula.coordinates.2 {
            for column in array_formula.coordinates.1..=array_formula.coordinates.3 {
                let has_default_cache = match range.get_value((row, u32::from(column))) {
                    Some(Data::Int(0)) => true,
                    Some(Data::Float(value)) if *value == 0.0 => true,
                    _ => false,
                };
                if !has_default_cache {
                    let reference = cell_reference(row, u32::from(column));
                    return Err(Error::Reverse(format!(
                        "array formula cached result in cell `{reference}` on Worksheet `{sheet_name}` cannot be preserved by the v1 DSL; only writer-default zero caches are representable, so recalculate/clear the cached result or add array formula result support"
                    )));
                }
            }
        }
    }
    Ok(())
}

fn validate_merged_values(
    values: &calamine::Range<Data>,
    formulas: &calamine::Range<String>,
    merged_ranges: &[MergeRange],
) -> xlsx_jsx::Result<()> {
    for merged in merged_ranges {
        for row in merged.first_row..=merged.last_row {
            for column in u32::from(merged.first_column)..=u32::from(merged.last_column) {
                if row == merged.first_row && column == u32::from(merged.first_column) {
                    continue;
                }
                let coordinate = (row, column);
                if values
                    .get_value(coordinate)
                    .is_some_and(|value| !value.is_empty())
                    || formulas
                        .get_value(coordinate)
                        .is_some_and(|formula| !formula.is_empty())
                {
                    return Err(Error::Reverse(format!(
                        "merged range {}:{} contains a value or formula in an interior cell; move it to the top-left cell or remove the merge before reversing",
                        cell_reference(merged.first_row, u32::from(merged.first_column)),
                        cell_reference(merged.last_row, u32::from(merged.last_column))
                    )));
                }
            }
        }
    }
    Ok(())
}

struct WorksheetRenderContext<'a> {
    range: &'a calamine::Range<Data>,
    formula_range: &'a calamine::Range<String>,
    merged_ranges: &'a [MergeRange],
    metadata: Option<&'a WorksheetMetadata>,
    sheet_name: &'a str,
}

impl WorksheetRenderContext<'_> {
    #[expect(
        clippy::too_many_lines,
        reason = "Row rendering resolves cell values, merged cells, and array-formula groups in one ordered coordinate pass."
    )]
    fn render_row(
        &self,
        output: &mut String,
        row_index: u32,
        bounds: Option<((u32, u32), (u32, u32))>,
    ) -> xlsx_jsx::Result<()> {
        let row_bounds = bounds.filter(|(start, end)| (start.0..=end.0).contains(&row_index));
        let first_column = row_bounds
            .map(|(start, _)| start.1)
            .into_iter()
            .chain(
                self.merged_ranges
                    .iter()
                    .filter(|merged| (merged.first_row..=merged.last_row).contains(&row_index))
                    .map(|merged| u32::from(merged.first_column)),
            )
            .chain(self.metadata.into_iter().flat_map(|data| {
                data.array_formula_ranges
                    .values()
                    .filter(|formula| {
                        (formula.coordinates.0..=formula.coordinates.2).contains(&row_index)
                    })
                    .map(|formula| u32::from(formula.coordinates.1))
            }))
            .min();
        let last_column = row_bounds
            .map(|(_, end)| end.1)
            .into_iter()
            .chain(
                self.merged_ranges
                    .iter()
                    .filter(|merged| (merged.first_row..=merged.last_row).contains(&row_index))
                    .map(|merged| u32::from(merged.last_column)),
            )
            .chain(self.metadata.into_iter().flat_map(|data| {
                data.array_formula_ranges
                    .values()
                    .filter(|formula| {
                        (formula.coordinates.0..=formula.coordinates.2).contains(&row_index)
                    })
                    .map(|formula| u32::from(formula.coordinates.3))
            }))
            .max();
        let (Some(first_column), Some(last_column)) = (first_column, last_column) else {
            return Ok(());
        };
        let empty = Data::Empty;
        let mut column_index = first_column;
        while column_index <= last_column {
            let interior = self.merged_ranges.iter().find(|merged| {
                (merged.first_row..=merged.last_row).contains(&row_index)
                    && column_index >= u32::from(merged.first_column)
                    && (row_index != merged.first_row
                        || column_index > u32::from(merged.first_column))
                    && column_index <= u32::from(merged.last_column)
            });
            if let Some(merged) = interior {
                column_index = u32::from(merged.last_column).saturating_add(1);
                continue;
            }
            let array_interior = self.metadata.and_then(|data| {
                data.array_formula_ranges.values().find_map(|formula| {
                    let range = formula.coordinates;
                    ((range.0..=range.2).contains(&row_index)
                        && column_index >= u32::from(range.1)
                        && column_index <= u32::from(range.3)
                        && (row_index != range.0 || column_index != u32::from(range.1)))
                    .then_some(range)
                })
            });
            if let Some(range) = array_interior {
                column_index = u32::from(range.3).saturating_add(1);
                continue;
            }
            let coordinate = (row_index, column_index);
            let reference = cell_reference(row_index, column_index);
            let merged = self.merged_ranges.iter().find(|merged| {
                merged.first_row == row_index && u32::from(merged.first_column) == column_index
            });
            let style = self
                .metadata
                .and_then(|data| data.cells.get(&reference))
                .cloned()
                .unwrap_or_default();
            let formula = self
                .formula_range
                .get_value(coordinate)
                .filter(|text| !text.is_empty());
            let array_formula_range = self
                .metadata
                .and_then(|data| data.array_formula_ranges.get(&reference));
            let mut cell_output = String::new();
            render_cell(
                &mut cell_output,
                self.range.get_value(coordinate).unwrap_or(&empty),
                formula.map(String::as_str),
                array_formula_range.is_some(),
                self.metadata
                    .and_then(|data| data.hyperlinks.get(&reference)),
                self.metadata
                    .and_then(|data| data.rich_text_cells.get(&reference))
                    .map(Vec::as_slice),
                &style,
                self.sheet_name,
                &reference,
            )?;
            if let Some(formula_range) = array_formula_range {
                if formula.is_none() {
                    return Err(Error::Reverse(format!(
                        "cell `{reference}` on Worksheet `{}` declares an array range but has no recoverable formula; repair the formula cell",
                        self.sheet_name
                    )));
                }
                let mut formula_props = String::new();
                if formula_range.is_dynamic {
                    formula_props.push_str(" dynamicArray={true}");
                }
                if !formula_range.is_dynamic || formula_range.has_explicit_range {
                    let _ = write!(
                        formula_props,
                        " arrayFormulaRange={{{:?}}}",
                        formula_range.reference
                    );
                }
                cell_output = cell_output.replacen("<Cell", &format!("<Cell{formula_props}"), 1);
            }
            if let Some(merged) = merged {
                let mut span_props = String::new();
                let row_span = merged.last_row - merged.first_row + 1;
                let column_span = u32::from(merged.last_column - merged.first_column + 1);
                if row_span > 1 {
                    let _ = write!(span_props, " rowSpan={{{row_span}}}");
                }
                if column_span > 1 {
                    let _ = write!(span_props, " columnSpan={{{column_span}}}");
                }
                cell_output = cell_output.replacen("<Cell", &format!("<Cell{span_props}"), 1);
                column_index = u32::from(merged.last_column);
            }
            output.push_str(&cell_output);
            column_index = column_index.saturating_add(1);
        }
        Ok(())
    }
}

fn render_row_open(
    output: &mut String,
    properties: Option<&RowProperties>,
    formats: &[CellFontStyle],
) {
    output.push_str("    <Row");
    if let Some(height) = properties.and_then(|properties| properties.height) {
        let _ = write!(output, " height={{{height}}}");
    }
    if properties.is_some_and(|properties| properties.hidden) {
        output.push_str(" hidden={true}");
    }
    if properties.is_some_and(|properties| properties.collapsed) {
        output.push_str(" collapsed={true}");
    }
    if let Some(level) = properties.map(|properties| properties.outline_level)
        && level > 0
    {
        let _ = write!(output, " outlineLevel={{{level}}}");
    }
    if let Some(style) = properties
        .and_then(|properties| properties.style_index)
        .and_then(|index| formats.get(index))
    {
        write_style_props(output, style);
    }
    output.push('>');
}

fn render_columns(
    output: &mut String,
    columns: &BTreeMap<u16, ColumnProperties>,
    formats: &[CellFontStyle],
) {
    let mut range: Option<(u16, u16, ColumnProperties)> = None;
    for (&column, &properties) in columns {
        match range {
            Some((first, last, previous))
                if last.checked_add(1) == Some(column) && previous == properties =>
            {
                range = Some((first, column, previous));
            }
            Some((first, last, previous)) => {
                write_column(output, first, last, previous, formats);
                range = Some((column, column, properties));
            }
            None => range = Some((column, column, properties)),
        }
    }
    if let Some((first, last, properties)) = range {
        write_column(output, first, last, properties, formats);
    }
}

fn render_data_validations(output: &mut String, validations: &[WorksheetDataValidation]) {
    for validation in validations {
        let _ = write!(
            output,
            "    <DataValidation range={{{:?}}}",
            validation.range
        );
        if validation.validation_type == "list" {
            let _ = write!(
                output,
                " type=\"list\" source={{{:?}}}",
                validation.formula1.as_deref().unwrap_or_default()
            );
        } else {
            let _ = write!(output, " type={:?}", validation.validation_type);
            if let Some(operator) = validation.operator.as_deref() {
                let _ = write!(output, " operator={operator:?}");
            }
            if let Some(formula1) = validation.formula1.as_deref() {
                let _ = write!(output, " formula1={formula1:?}");
            }
            if let Some(formula2) = validation.formula2.as_deref() {
                let _ = write!(output, " formula2={formula2:?}");
            }
        }
        if !validation.allow_blank {
            output.push_str(" allowBlank={false}");
        }
        if !validation.show_dropdown {
            output.push_str(" showDropDown={false}");
        }
        if !validation.show_input_message {
            output.push_str(" showInputMessage={false}");
        }
        if !validation.show_error_message {
            output.push_str(" showErrorMessage={false}");
        }
        for (property, value) in [
            ("inputTitle", validation.input_title.as_deref()),
            ("inputMessage", validation.input_message.as_deref()),
            ("errorTitle", validation.error_title.as_deref()),
            ("errorMessage", validation.error_message.as_deref()),
            ("errorStyle", validation.error_style.as_deref()),
        ] {
            if let Some(value) = value {
                let _ = write!(output, " {property}={value:?}");
            }
        }
        output.push_str(" />\n");
    }
}

#[expect(
    clippy::too_many_lines,
    reason = "Conditional-format JSX rendering shares one stable component shape across rule families."
)]
fn render_conditional_formats(output: &mut String, formats: &[WorksheetConditionalFormat]) {
    for format in formats {
        let _ = write!(
            output,
            "    <ConditionalFormat range={{{:?}}} type={:?}",
            format.range, format.scale_type
        );
        if format.stop_if_true {
            output.push_str(" stopIfTrue={true}");
        }
        if format.scale_type == "dataBar" {
            if let Some(bar) = &format.data_bar {
                if let Some(color) = bar
                    .fill_color
                    .as_deref()
                    .filter(|color| *color != "#638EC6")
                {
                    let _ = write!(output, " fillColor={color:?}");
                }
                for (name, threshold) in [("min", &bar.minimum), ("max", &bar.maximum)] {
                    if threshold.kind != "auto" {
                        let _ = write!(output, " {name}Type={:?}", threshold.kind);
                        if let Some(value) = threshold.value.as_deref() {
                            if threshold.kind == "formula" {
                                let _ = write!(output, " {name}Value={value:?}");
                            } else if let Ok(value) = value.parse::<f64>() {
                                let _ = write!(output, " {name}Value={{{value}}}");
                            }
                        }
                    }
                }
                for (name, color) in [
                    (
                        "borderColor",
                        bar.border_color
                            .as_deref()
                            .filter(|color| Some(*color) != bar.fill_color.as_deref()),
                    ),
                    (
                        "negativeFillColor",
                        bar.negative_fill_color
                            .as_deref()
                            .filter(|color| *color != "#FF0000"),
                    ),
                    (
                        "negativeBorderColor",
                        bar.negative_border_color
                            .as_deref()
                            .filter(|color| Some(*color) != bar.negative_fill_color.as_deref()),
                    ),
                    (
                        "axisColor",
                        bar.axis_color
                            .as_deref()
                            .filter(|color| *color != "#000000"),
                    ),
                ] {
                    if let Some(color) = color {
                        let _ = write!(output, " {name}={color:?}");
                    }
                }
                for (name, value) in [
                    ("barOnly", bar.bar_only),
                    ("solidFill", bar.solid_fill),
                    ("border", bar.border),
                ] {
                    if let Some(value) = value.filter(|value| match name {
                        "barOnly" | "solidFill" => *value,
                        "border" => !*value,
                        _ => false,
                    }) {
                        let _ = write!(output, " {name}={{{value}}}");
                    }
                }
                for (name, value) in [
                    ("direction", bar.direction.as_deref()),
                    ("axisPosition", bar.axis_position.as_deref()),
                ] {
                    if let Some(value) = value.filter(|value| match name {
                        "direction" => *value != "context",
                        "axisPosition" => *value != "automatic",
                        _ => false,
                    }) {
                        let _ = write!(output, " {name}={value:?}");
                    }
                }
            }
            output.push_str(" />\n");
            continue;
        }
        if format.scale_type == "iconSet" {
            if let Some(icon_set) = &format.icon_set {
                if icon_set.icon_type != "threeTrafficLights" {
                    let _ = write!(output, " iconType={:?}", icon_set.icon_type);
                }
                if icon_set.reverse {
                    output.push_str(" reverse={true}");
                }
                if icon_set.show_icons_only {
                    output.push_str(" showIconsOnly={true}");
                }
                if !icon_set.icons.is_empty() {
                    output.push_str(" icons={[ ");
                    for (index, icon) in icon_set.icons.iter().enumerate() {
                        if index > 0 {
                            output.push_str(", ");
                        }
                        let _ = write!(output, "{{ type: {:?}, value: ", icon.kind);
                        if icon.kind == "formula" {
                            let _ =
                                write!(output, "{:?}", icon.value.as_deref().unwrap_or_default());
                        } else if let Some(value) = icon.value.as_deref() {
                            output.push_str(value);
                        } else {
                            output.push('0');
                        }
                        if icon.greater_than {
                            output.push_str(", greaterThan: true");
                        }
                        if let (Some(icon_type), Some(icon_index)) =
                            (icon.icon_type.as_deref(), icon.icon_index)
                        {
                            let _ = write!(
                                output,
                                ", iconType: {icon_type:?}, iconIndex: {icon_index}"
                            );
                        }
                        if icon.no_icon {
                            output.push_str(", noIcon: true");
                        }
                        output.push_str(" }");
                    }
                    output.push_str(" ]}");
                }
            }
            output.push_str(" />\n");
            continue;
        }
        if ["average", "date", "blank", "error", "duplicate"].contains(&format.scale_type.as_str())
        {
            let default_rule = match format.scale_type.as_str() {
                "average" => "aboveAverage",
                "date" => "yesterday",
                "blank" => "blank",
                "error" => "error",
                "duplicate" => "duplicate",
                _ => "",
            };
            if let Some(rule) = format
                .top_rule
                .as_deref()
                .filter(|rule| *rule != default_rule)
            {
                let _ = write!(output, " rule={rule:?}");
            }
            write_conditional_style_props(output, &format.format);
            output.push_str(" />\n");
            continue;
        }
        if format.scale_type == "top10" {
            if let Some(rule) = format.top_rule.as_deref().filter(|rule| *rule != "top") {
                let _ = write!(output, " rule={rule:?}");
            }
            if let Some(rank) = format.rank.filter(|rank| *rank != 10) {
                let _ = write!(output, " rank={{{rank}}}");
            }
            write_conditional_style_props(output, &format.format);
            output.push_str(" />\n");
            continue;
        }
        if matches!(format.scale_type.as_str(), "cellIs" | "formula") {
            if format.scale_type == "cellIs"
                && let Some(operator) = format.operator.as_deref()
            {
                let _ = write!(output, " operator={operator:?}");
            }
            if let Some(formula) = format.formula1.as_deref() {
                let _ = write!(output, " formula1={formula:?}");
            }
            if let Some(formula) = format.formula2.as_deref() {
                let _ = write!(output, " formula2={formula:?}");
            }
            write_conditional_style_props(output, &format.format);
            output.push_str(" />\n");
            continue;
        }
        if format.scale_type == "text" {
            if let Some(operator) = format.operator.as_deref() {
                let _ = write!(output, " operator={operator:?}");
            }
            if let Some(text) = format.text.as_deref() {
                let _ = write!(output, " text={text:?}");
            }
            write_conditional_style_props(output, &format.format);
            output.push_str(" />\n");
            continue;
        }
        for (name, threshold) in [
            ("min", &format.minimum),
            ("mid", format.midpoint.as_ref().unwrap_or(&format.minimum)),
            ("max", &format.maximum),
        ] {
            if name == "mid" && format.midpoint.is_none() {
                continue;
            }
            let _ = write!(output, " {name}Color={:?}", threshold.color);
            if threshold.kind != "min" && threshold.kind != "max" {
                let _ = write!(output, " {name}Type={:?}", threshold.kind);
                if let Some(value) = threshold.value.as_deref() {
                    if threshold.kind == "formula" {
                        let _ = write!(output, " {name}Value={value:?}");
                    } else {
                        let _ = write!(output, " {name}Value={{{value}}}");
                    }
                }
            }
        }
        output.push_str(" />\n");
    }
}

fn write_conditional_style_props(output: &mut String, style: &CellFontStyle) {
    if style.bold {
        output.push_str(" bold={true}");
    }
    if style.italic {
        output.push_str(" italic={true}");
    }
    if style.strikethrough == Some(true) {
        output.push_str(" strikethrough={true}");
    }
    if let Some(underline) = style.underline.as_deref() {
        let _ = write!(output, " underline={underline:?}");
    }
    if let Some(script) = style.script.as_deref() {
        let _ = write!(output, " script={script:?}");
    }
    if let Some(color) = style.font_color.as_deref() {
        let _ = write!(output, " fontColor={color:?}");
    }
    if let Some(color) = style.fill_color.as_deref() {
        let _ = write!(output, " fillColor={color:?}");
    }
    if let Some(pattern) = style.fill_pattern.as_deref() {
        let _ = write!(output, " fillPattern=\"{pattern}\"");
    }
    if let Some(color) = style.fill_foreground_color.as_deref() {
        let _ = write!(output, " fillForegroundColor=\"{color}\"");
    }
    if let Some(color) = style.fill_background_color.as_deref() {
        let _ = write!(output, " fillBackgroundColor=\"{color}\"");
    }
    if let Some(number_format) = style.number_format.as_deref() {
        let _ = write!(output, " numberFormat={number_format:?}");
    }
    for side in BorderSide::ALL {
        let edge = &style.borders.edges[side.index()];
        if let Some(border_style) = edge.style {
            let _ = write!(
                output,
                " border{}={:?}",
                side.suffix(),
                border_style_name(border_style)
            );
        }
        if let Some(color) = edge.color.as_deref() {
            let _ = write!(output, " border{}Color={color:?}", side.suffix());
        }
    }
}

fn render_ignored_errors(output: &mut String, ignored_errors: &[WorksheetIgnoredError]) {
    for ignored_error in ignored_errors {
        let _ = writeln!(
            output,
            "    <IgnoreError range={:?} errorType={:?} />",
            ignored_error.range, ignored_error.error_type
        );
    }
}

fn write_column(
    output: &mut String,
    first: u16,
    last: u16,
    properties: ColumnProperties,
    formats: &[CellFontStyle],
) {
    let _ = write!(
        output,
        "    <Column min={{{}}} max={{{}}}",
        u32::from(first) + 1,
        u32::from(last) + 1
    );
    if let Some(width_pixels) = properties.width_pixels {
        let _ = write!(output, " widthPixels={{{width_pixels}}}");
    }
    if let Some(style) = properties.style_index.and_then(|index| formats.get(index)) {
        write_style_props(output, style);
    }
    if properties.hidden {
        output.push_str(" hidden={true}");
    }
    if properties.collapsed {
        output.push_str(" collapsed={true}");
    }
    if properties.outline_level > 0 {
        let _ = write!(output, " outlineLevel={{{}}}", properties.outline_level);
    }
    output.push_str(" />\n");
}

#[expect(
    clippy::too_many_arguments,
    reason = "Cell content, cell metadata, and worksheet context are independently sourced and diagnosed together."
)]
fn render_cell(
    output: &mut String,
    value: &Data,
    formula: Option<&str>,
    array_formula: bool,
    hyperlink: Option<&WorksheetHyperlink>,
    rich_text: Option<&[RichTextRun]>,
    style: &CellFontStyle,
    sheet_name: &str,
    reference: &str,
) -> xlsx_jsx::Result<()> {
    if style.checkbox && (formula.is_some() || !matches!(value, Data::Bool(_))) {
        return Err(Error::Reverse(format!(
            "cell `{reference}` on Worksheet `{sheet_name}` has checkbox formatting without a scalar boolean value; use a boolean Cell value or remove the checkbox style"
        )));
    }
    if hyperlink.is_some() && (formula.is_some() || !matches!(value, Data::String(_))) {
        return Err(Error::Reverse(format!(
            "cell `{reference}` on Worksheet `{sheet_name}` has a hyperlink without plain string text; use a text cell or add matching formula/type support"
        )));
    }
    if let (Some(hyperlink), Data::String(text)) = (hyperlink, value)
        && hyperlink
            .display
            .as_deref()
            .is_some_and(|display| display != text)
    {
        return Err(Error::Reverse(format!(
            "cell `{reference}` on Worksheet `{sheet_name}` has hyperlink display text that differs from its cell value; make both strings identical or remove the hyperlink display metadata"
        )));
    }
    if let Some(formula) = formula {
        output.push_str("<Cell");
        write_style_props(output, style);
        let formula = format!("={formula}");
        let _ = write!(output, " formula={{{formula:?}}}");
        if !array_formula {
            write_formula_result_prop(output, value, sheet_name, reference)?;
        }
        output.push_str(" />");
        return Ok(());
    }

    match value {
        Data::String(text) => {
            if let Some(runs) = rich_text {
                let reconstructed = runs.iter().map(|run| run.text.as_str()).collect::<String>();
                if reconstructed != *text {
                    return Err(Error::Reverse(format!(
                        "cell `{reference}` on Worksheet `{sheet_name}` has rich-text runs that do not match its displayed value; repair the shared or inline string before reversing"
                    )));
                }
                output.push_str("<Cell");
                write_style_props(output, style);
                output.push('>');
                for run in runs {
                    output.push_str("<TextRun");
                    write_text_run_props(output, &run.format);
                    let _ = write!(output, ">{{{:?}}}</TextRun>", run.text);
                }
                output.push_str("</Cell>");
                return Ok(());
            }
            output.push_str("<Cell");
            if let Some(target) = hyperlink {
                let _ = write!(output, " hyperlink={:?}", target.target);
                if let Some(tooltip) = &target.tooltip {
                    let _ = write!(output, " hyperlinkTooltip={tooltip:?}");
                }
            }
            write_style_props(output, style);
            let _ = write!(output, ">{{{text:?}}}</Cell>");
        }
        Data::Float(number) => write_numeric_cell(output, style, number),
        Data::Int(number) => write_numeric_cell(output, style, number),
        Data::Bool(boolean) => {
            output.push_str("<Cell");
            write_style_props(output, style);
            let _ = write!(output, " value={{{boolean}}}/>");
        }
        Data::Empty => {
            output.push_str("<Cell");
            write_style_props(output, style);
            output.push_str(" />");
        }
        Data::DateTime(datetime) => render_datetime_cell(
            output,
            value,
            datetime.as_f64(),
            datetime.is_duration(),
            style,
            sheet_name,
            reference,
        )?,
        Data::DateTimeIso(_) => {
            let date = reverse_date_value(value, DateNumberFormat::General, sheet_name, reference)?;
            output.push_str("<Cell");
            write_style_props(output, style);
            let _ = write!(output, " date={{{date:?}}}/>");
        }
        Data::DurationIso(_) => {
            return Err(Error::Reverse(format!(
                "cell `{reference}` on Worksheet `{sheet_name}` contains a duration that v1 cannot preserve; convert it to text or add a duration-aware Cell type"
            )));
        }
        Data::Error(error) => {
            return Err(Error::Reverse(format!(
                "cell `{reference}` on Worksheet `{sheet_name}` contains Excel error `{error}` that v1 cannot preserve; fix the formula/error or add an explicit error-cell type"
            )));
        }
    }
    Ok(())
}

fn write_formula_result_prop(
    output: &mut String,
    value: &Data,
    sheet_name: &str,
    reference: &str,
) -> xlsx_jsx::Result<()> {
    match value {
        Data::Int(number) => {
            let _ = write!(output, " formulaResult={{{number}}}");
        }
        Data::Float(number) if number.is_finite() => {
            let _ = write!(output, " formulaResult={{{number}}}");
        }
        Data::Float(_) => {
            return Err(Error::Reverse(format!(
                "formula cached result in cell `{reference}` on Worksheet `{sheet_name}` is not finite; repair the cached result before reversing"
            )));
        }
        Data::String(text) if text.is_empty() => output.push_str(" formulaResult={null}"),
        Data::String(text) => {
            if xlsx_jsx::formula_result_text_is_ambiguous(text) {
                return Err(Error::Reverse(format!(
                    "formula cached text result in cell `{reference}` on Worksheet `{sheet_name}` is `{text}`, which the writer would serialize as a different value type; change the result to a true string or add typed formula-result support"
                )));
            }
            let _ = write!(output, " formulaResult={text:?}");
        }
        Data::Bool(value) => {
            let _ = write!(output, " formulaResult={{{value}}}");
        }
        Data::Error(error) => {
            let _ = write!(output, " formulaResult={:?}", error.to_string());
        }
        Data::Empty => output.push_str(" formulaResult={null}"),
        Data::DateTime(datetime) => {
            let _ = write!(output, " formulaResult={{{}}}", datetime.as_f64());
        }
        Data::DateTimeIso(_) | Data::DurationIso(_) => {
            return Err(Error::Reverse(format!(
                "formula cached result in cell `{reference}` on Worksheet `{sheet_name}` uses a date/duration cell type that the writer cannot reproduce; preserve a numeric formula cache or add typed formula-result support"
            )));
        }
    }
    Ok(())
}

fn write_text_run_props(output: &mut String, style: &CellFontStyle) {
    if style.bold {
        output.push_str(" bold={true}");
    }
    if style.italic {
        output.push_str(" italic={true}");
    }
    if style.strikethrough == Some(true) {
        output.push_str(" strikethrough={true}");
    }
    if let Some(underline) = &style.underline {
        let _ = write!(output, " underline={{{underline:?}}}");
    }
    if let Some(script) = &style.script {
        let _ = write!(output, " script={{{script:?}}}");
    }
    if let Some(color) = &style.font_color {
        let _ = write!(output, " fontColor={{{color:?}}}");
    }
    if let Some(name) = &style.font_name {
        let _ = write!(output, " fontName={{{name:?}}}");
    }
    if let Some(size) = style.font_size {
        let _ = write!(output, " fontSize={{{size}}}");
    }
    if let Some(family) = style.font_family {
        let _ = write!(output, " fontFamily={{{family}}}");
    }
    if let Some(charset) = style.font_charset {
        let _ = write!(output, " fontCharset={{{charset}}}");
    }
    if style.font_scheme.as_deref() == Some("major") {
        output.push_str(" fontScheme=\"headings\"");
    } else if style.font_scheme.as_deref() == Some("none") {
        output.push_str(" fontScheme=\"none\"");
    }
}

fn write_numeric_cell(output: &mut String, style: &CellFontStyle, number: &impl std::fmt::Display) {
    output.push_str("<Cell");
    write_style_props(output, style);
    let _ = write!(output, " value={{{number}}}/>");
}

#[expect(
    clippy::too_many_lines,
    reason = "The canonical JSX style-property order is emitted together for stable reverse output."
)]
fn write_style_props(output: &mut String, style: &CellFontStyle) {
    if style.checkbox {
        output.push_str(" checkbox={true}");
    }
    if style.bold {
        output.push_str(" bold={true}");
    }
    if style.italic {
        output.push_str(" italic={true}");
    }
    if style.quote_prefix.is_some() {
        output.push_str(" quotePrefix={true}");
    }
    if style.locked == Some(false) {
        output.push_str(" locked={false}");
    }
    if style.formula_hidden {
        output.push_str(" formulaHidden={true}");
    }
    if style.strikethrough == Some(true) {
        output.push_str(" strikethrough={true}");
    }
    if let Some(underline) = &style.underline {
        let _ = write!(output, " underline={{{underline:?}}}");
    }
    if let Some(script) = &style.script {
        let _ = write!(output, " script={{{script:?}}}");
    }
    if style.wrap_text == Some(true) {
        output.push_str(" wrapText={true}");
    }
    if style.shrink_to_fit == Some(true) {
        output.push_str(" shrinkToFit={true}");
    }
    if let Some(indent) = style.indent {
        let _ = write!(output, " indent={{{indent}}}");
    }
    if let Some(alignment) = style.horizontal_align {
        let _ = write!(output, " horizontalAlign={{{:?}}}", alignment.as_str());
    }
    if let Some(alignment) = style.vertical_align {
        let _ = write!(output, " verticalAlign={{{:?}}}", alignment.as_str());
    }
    if let Some(rotation) = style.text_rotation {
        let _ = write!(output, " textRotation={{{rotation}}}");
    }
    if let Some(direction) = style.reading_direction {
        let _ = write!(output, " readingDirection=\"{}\"", direction.jsx_name());
    }
    if let Some(number_format) = &style.number_format {
        let _ = write!(output, " numberFormat={{{number_format:?}}}");
    }
    if let Some(fill_color) = &style.fill_color {
        let _ = write!(output, " fillColor={{{fill_color:?}}}");
    }
    if let Some(pattern) = &style.fill_pattern {
        let _ = write!(output, " fillPattern=\"{pattern}\"");
    }
    if let Some(color) = &style.fill_foreground_color {
        let _ = write!(output, " fillForegroundColor=\"{color}\"");
    }
    if let Some(color) = &style.fill_background_color {
        let _ = write!(output, " fillBackgroundColor=\"{color}\"");
    }
    if let Some(font_color) = &style.font_color {
        let _ = write!(output, " fontColor={{{font_color:?}}}");
    }
    if let Some(font_name) = &style.font_name {
        let _ = write!(output, " fontName={{{font_name:?}}}");
    }
    if let Some(font_size) = style.font_size {
        let _ = write!(output, " fontSize={{{font_size}}}");
    }
    if let Some(font_family) = style.font_family {
        let _ = write!(output, " fontFamily={{{font_family}}}");
    }
    if let Some(font_charset) = style.font_charset {
        let _ = write!(output, " fontCharset={{{font_charset}}}");
    }
    if style.font_scheme.as_deref() == Some("major") {
        output.push_str(" fontScheme=\"headings\"");
    } else if style.font_scheme.as_deref() == Some("none") {
        output.push_str(" fontScheme=\"none\"");
    }
    for side in BorderSide::ALL {
        let edge = &style.borders.edges[side.index()];
        if let Some(border_style) = edge.style {
            let _ = write!(
                output,
                " border{}={{{:?}}}",
                side.suffix(),
                border_style_name(border_style)
            );
        }
        if let Some(color) = &edge.color {
            let _ = write!(output, " border{}Color={{{color:?}}}", side.suffix());
        }
    }
    if let Some(border_style) = style.borders.diagonal.style {
        let _ = write!(
            output,
            " borderDiagonal={{{:?}}}",
            border_style_name(border_style)
        );
    }
    if let Some(color) = &style.borders.diagonal.color {
        let _ = write!(output, " borderDiagonalColor={{{color:?}}}");
    }
    if let Some(direction) = &style.borders.diagonal_type {
        let _ = write!(output, " borderDiagonalType={{{direction:?}}}");
    }
}

fn apply_cell_borders(mut format: Format, cell: &Node) -> Format {
    for side in BorderSide::ALL {
        let style_property = format!("border{}", side.suffix());
        if let Some(style) = cell.props.get(&style_property).and_then(Value::as_str)
            && let Some(border_style) = parse_border_style(style)
        {
            format = match side {
                BorderSide::Top => format.set_border_top(border_style),
                BorderSide::Right => format.set_border_right(border_style),
                BorderSide::Bottom => format.set_border_bottom(border_style),
                BorderSide::Left => format.set_border_left(border_style),
            };
        }
        let color_property = format!("border{}Color", side.suffix());
        if let Some(color) = cell.props.get(&color_property).and_then(Value::as_str) {
            format = match side {
                BorderSide::Top => format.set_border_top_color(color),
                BorderSide::Right => format.set_border_right_color(color),
                BorderSide::Bottom => format.set_border_bottom_color(color),
                BorderSide::Left => format.set_border_left_color(color),
            };
        }
    }
    if let Some(style) = cell.props.get("borderDiagonal").and_then(Value::as_str)
        && let Some(border_style) = parse_border_style(style)
    {
        format = format.set_border_diagonal(border_style);
    }
    if let Some(color) = cell
        .props
        .get("borderDiagonalColor")
        .and_then(Value::as_str)
    {
        format = format.set_border_diagonal_color(color);
    }
    if let Some(direction) = cell.props.get("borderDiagonalType").and_then(Value::as_str) {
        let direction = match direction {
            "up" => FormatDiagonalBorder::BorderUp,
            "down" => FormatDiagonalBorder::BorderDown,
            "upDown" => FormatDiagonalBorder::BorderUpDown,
            _ => FormatDiagonalBorder::None,
        };
        format = format.set_border_diagonal_type(direction);
    }
    format
}

fn parse_border_style(value: &str) -> Option<FormatBorder> {
    Some(match value {
        "none" => FormatBorder::None,
        "thin" => FormatBorder::Thin,
        "medium" => FormatBorder::Medium,
        "dashed" => FormatBorder::Dashed,
        "dotted" => FormatBorder::Dotted,
        "thick" => FormatBorder::Thick,
        "double" => FormatBorder::Double,
        "hair" => FormatBorder::Hair,
        "mediumDashed" => FormatBorder::MediumDashed,
        "dashDot" => FormatBorder::DashDot,
        "mediumDashDot" => FormatBorder::MediumDashDot,
        "dashDotDot" => FormatBorder::DashDotDot,
        "mediumDashDotDot" => FormatBorder::MediumDashDotDot,
        "slantDashDot" => FormatBorder::SlantDashDot,
        _ => return None,
    })
}

fn border_style_name(value: FormatBorder) -> &'static str {
    match value {
        FormatBorder::None => "none",
        FormatBorder::Thin => "thin",
        FormatBorder::Medium => "medium",
        FormatBorder::Dashed => "dashed",
        FormatBorder::Dotted => "dotted",
        FormatBorder::Thick => "thick",
        FormatBorder::Double => "double",
        FormatBorder::Hair => "hair",
        FormatBorder::MediumDashed => "mediumDashed",
        FormatBorder::DashDot => "dashDot",
        FormatBorder::MediumDashDot => "mediumDashDot",
        FormatBorder::DashDotDot => "dashDotDot",
        FormatBorder::MediumDashDotDot => "mediumDashDotDot",
        FormatBorder::SlantDashDot => "slantDashDot",
    }
}

fn render_datetime_cell(
    output: &mut String,
    value: &Data,
    serial: f64,
    is_duration: bool,
    style: &CellFontStyle,
    sheet_name: &str,
    reference: &str,
) -> xlsx_jsx::Result<()> {
    let Some(number_format) = style.number_format.as_deref() else {
        return Err(Error::Reverse(format!(
            "cell `{reference}` on Worksheet `{sheet_name}` is date/time typed but has no recoverable number format; repair the style table or add explicit date-system support"
        )));
    };
    let date_format = classify_date_format(number_format);
    let can_render_as_iso = !is_duration
        && serial_has_millisecond_precision(serial)
        && match date_format {
            DateNumberFormat::Date => serial.fract() == 0.0,
            DateNumberFormat::DateTime => value
                .as_datetime()
                .is_some_and(|datetime| datetime.and_utc().timestamp_subsec_millis() == 0),
            DateNumberFormat::DateTimeMillis => true,
            DateNumberFormat::General | DateNumberFormat::Unsupported => false,
        }
        && !(60.0..61.0).contains(&serial);
    if can_render_as_iso {
        let date = reverse_date_value(value, date_format, sheet_name, reference)?;
        output.push_str("<Cell");
        write_style_props(output, style);
        let _ = write!(output, " date={{{date:?}}}/>");
    } else {
        write_numeric_cell(output, style, &serial);
    }
    Ok(())
}

fn serial_has_millisecond_precision(serial: f64) -> bool {
    let milliseconds = serial * 86_400_000.0;
    milliseconds.is_finite() && (milliseconds - milliseconds.round()).abs() < 1e-5
}

fn reverse_date_value(
    value: &Data,
    number_format: DateNumberFormat,
    sheet_name: &str,
    reference: &str,
) -> xlsx_jsx::Result<String> {
    if let Data::DateTimeIso(text) = value {
        if !xlsx_jsx::is_supported_iso_date(text) {
            return Err(Error::Reverse(format!(
                "cell `{reference}` on Worksheet `{sheet_name}` contains a time-zone-bearing or unsupported ISO date/time; use a local ISO date/time without a zone because Excel serial dates do not retain time zones"
            )));
        }
        return Ok(text.clone());
    }

    let Some(datetime) = value.as_datetime() else {
        return Err(Error::Reverse(format!(
            "cell `{reference}` on Worksheet `{sheet_name}` has an invalid date serial; repair the cell value or date number format"
        )));
    };
    let date = match number_format {
        DateNumberFormat::Date => datetime.date().format("%Y-%m-%d").to_string(),
        DateNumberFormat::DateTime => datetime.format("%Y-%m-%dT%H:%M:%S").to_string(),
        DateNumberFormat::DateTimeMillis => format!(
            "{}.{:03}",
            datetime.format("%Y-%m-%dT%H:%M:%S"),
            datetime.and_utc().timestamp_subsec_millis()
        ),
        DateNumberFormat::General | DateNumberFormat::Unsupported => {
            return Err(Error::Reverse(format!(
                "cell `{reference}` on Worksheet `{sheet_name}` is date-typed without a supported date number format; use one of the DSL date formats or add explicit number-format support"
            )));
        }
    };
    Ok(date)
}

#[expect(
    clippy::too_many_lines,
    reason = "Package-level relationship, part, and worksheet audits share one archive transaction."
)]
fn audit_reverse_subset(
    path: &std::path::Path,
) -> xlsx_jsx::Result<(
    HashMap<String, WorksheetMetadata>,
    WorkbookProperties,
    Vec<CellFontStyle>,
)> {
    let file = std::fs::File::open(path).map_err(|source| Error::Resource {
        path: path.to_path_buf(),
        source,
    })?;
    let mut archive = ZipArchive::new(file).map_err(|error| Error::Reverse(error.to_string()))?;
    let (formats, differential_formats) = parse_style_tables(&mut archive)?;
    let workbook_xml = read_zip_part(&mut archive, "xl/workbook.xml")?;
    let relationships_xml = read_zip_part(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let has_dynamic_metadata = archive.file_names().any(|name| name == "xl/metadata.xml");
    let has_checkbox_property_bag = archive
        .file_names()
        .any(|name| name == "xl/featurePropertyBag/featurePropertyBag.xml");
    validate_dynamic_metadata_relationship(&relationships_xml, has_dynamic_metadata)?;
    validate_checkbox_feature_relationship(&relationships_xml, has_checkbox_property_bag)?;
    let has_checkbox_style = formats.iter().any(|style| style.checkbox);
    if has_checkbox_style != has_checkbox_property_bag {
        return Err(Error::Reverse(
            "checkbox cell styles and xl/featurePropertyBag/featurePropertyBag.xml are inconsistent; repair the checkbox feature metadata or remove checkbox styles"
                .into(),
        ));
    }
    if has_checkbox_property_bag {
        let feature_bag =
            read_zip_part(&mut archive, "xl/featurePropertyBag/featurePropertyBag.xml")?;
        audit_checkbox_feature_property_bag(&feature_bag)?;
    }
    if has_dynamic_metadata {
        let metadata_xml = read_zip_part(&mut archive, "xl/metadata.xml")?;
        audit_dynamic_array_metadata(&metadata_xml)?;
    }
    let sheet_targets = parse_sheet_targets(&workbook_xml, &relationships_xml)?;
    let sheet_order = parse_sheet_relation_ids(&workbook_xml)?
        .into_iter()
        .map(|(name, _)| name)
        .collect::<Vec<_>>();
    let worksheet_names = sheet_order
        .iter()
        .map(|name| name.to_lowercase())
        .collect::<std::collections::HashSet<_>>();
    let visibility_by_sheet = parse_worksheet_visibility(&workbook_xml, &sheet_order)?;
    let shared_strings = if archive
        .file_names()
        .any(|name| name == "xl/sharedStrings.xml")
    {
        parse_shared_strings(
            &read_zip_part(&mut archive, "xl/sharedStrings.xml")?,
            "xl/sharedStrings.xml",
        )?
    } else {
        Vec::new()
    };
    let names_by_target: HashMap<_, _> = sheet_targets
        .iter()
        .map(|(name, target)| (target.as_str(), name.as_str()))
        .collect();
    let relationship_parts = archive
        .file_names()
        .filter(|name| name.starts_with("xl/worksheets/_rels/") && name.ends_with(".xml.rels"))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    let mut hyperlink_relationships_by_sheet = HashMap::new();
    let mut table_relationships_by_sheet = HashMap::new();
    let mut drawing_relationships_by_sheet = HashMap::new();
    let mut note_relationships_by_sheet = HashMap::new();
    let mut referenced_note_parts = std::collections::HashSet::new();
    for relationship_part in relationship_parts {
        let worksheet_part = worksheet_part_from_relationship_part(&relationship_part)
            .ok_or_else(|| Error::Reverse("invalid worksheet relationship part name".into()))?;
        let relationships = parse_worksheet_relationships(
            &read_zip_part(&mut archive, &relationship_part)?,
            &relationship_part,
        )?;
        hyperlink_relationships_by_sheet.insert(worksheet_part.clone(), relationships.hyperlinks);
        if relationships.drawings.len() > 1 {
            return Err(Error::Reverse(format!(
                "worksheet `{worksheet_part}` has multiple drawing relationships; combine the charts into the single worksheet drawing used by xlsx-jsx"
            )));
        }
        drawing_relationships_by_sheet.insert(worksheet_part.clone(), relationships.drawings);
        if relationships.comments.len() != relationships.vml_drawings.len()
            || relationships.comments.len() > 1
        {
            return Err(Error::Reverse(format!(
                "worksheet `{worksheet_part}` must have exactly one comments and one VML relationship for legacy notes; repair its relationship part"
            )));
        }
        if (!relationships.comments.is_empty() || !relationships.vml_drawings.is_empty())
            && !names_by_target.contains_key(worksheet_part.as_str())
        {
            return Err(Error::Reverse(format!(
                "legacy note relationships belong to unlinked worksheet `{worksheet_part}`; repair workbook.xml relationships"
            )));
        }
        if let (
            Some((_comments_relationship_id, comments_target)),
            Some((vml_relationship_id, vml_target)),
        ) = (
            relationships.comments.iter().next(),
            relationships.vml_drawings.iter().next(),
        ) {
            let comments_path = resolve_note_target(comments_target, &worksheet_part, true)?;
            let vml_path = resolve_note_target(vml_target, &worksheet_part, false)?;
            if !referenced_note_parts.insert(comments_path.clone())
                || !referenced_note_parts.insert(vml_path.clone())
            {
                return Err(Error::Reverse(format!(
                    "note part linked from `{worksheet_part}` is shared by multiple worksheets; use one comments/VML pair per sheet"
                )));
            }
            let comments_xml = read_zip_part(&mut archive, &comments_path)?;
            let vml_xml = read_zip_part(&mut archive, &vml_path)?;
            let comments = parse_comments_xml(&comments_xml, &comments_path)?;
            let shapes = parse_vml_notes(&vml_xml, &vml_path)?;
            let notes = match_comments_and_vml(comments, shapes, &worksheet_part)?;
            note_relationships_by_sheet.insert(
                worksheet_part.clone(),
                WorksheetNoteParts {
                    vml_relationship_id: vml_relationship_id.clone(),
                    notes,
                },
            );
        }
        table_relationships_by_sheet.insert(worksheet_part, relationships.tables);
    }
    let mut styles_by_sheet = HashMap::new();
    let mut workbook_properties = WorkbookProperties::default();
    for index in 0..archive.len() {
        let mut part = archive
            .by_index(index)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        if part.is_dir() {
            continue;
        }
        let name = part.name().to_owned();
        if (name.starts_with("xl/comments") || name.starts_with("xl/drawings/vmlDrawing"))
            && !referenced_note_parts.contains(&name)
        {
            return Err(Error::Reverse(format!(
                "package part `{name}` is not referenced as a worksheet legacy note; remove it or add explicit parsing for that drawing/comment type"
            )));
        }
        if has_unsupported_part(&name) {
            return Err(Error::Reverse(format!(
                "package part `{name}` is outside the lossless v1 subset; remove that feature or extend the DSL and reverse reader together"
            )));
        }
        if !is_supported_part(&name) {
            return Err(Error::Reverse(format!(
                "package part `{name}` is not recognized by the v1 reader; remove the custom feature or add explicit parsing and reverse support"
            )));
        }
        if name.starts_with("xl/worksheets/")
            && std::path::Path::new(&name)
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        {
            let mut xml = String::new();
            std::io::Read::read_to_string(&mut part, &mut xml)
                .map_err(|error| Error::Reverse(error.to_string()))?;
            let no_hyperlink_relationships = HashMap::new();
            let hyperlink_relationships = hyperlink_relationships_by_sheet
                .get(&name)
                .unwrap_or(&no_hyperlink_relationships);
            let metadata = audit_worksheet_xml(
                &xml,
                &name,
                &formats,
                &differential_formats,
                hyperlink_relationships,
                &shared_strings,
                has_dynamic_metadata,
            )?;
            let mut metadata = metadata;
            if let Some(note_parts) = note_relationships_by_sheet.get(&name) {
                if metadata.legacy_drawing_relationship_id.as_deref()
                    != Some(note_parts.vml_relationship_id.as_str())
                {
                    return Err(Error::Reverse(format!(
                        "worksheet `{name}` legacyDrawing relationship does not match its VML note relationship; repair r:id `{}`",
                        note_parts.vml_relationship_id
                    )));
                }
                metadata.notes.clone_from(&note_parts.notes);
            } else if metadata.legacy_drawing_relationship_id.is_some() {
                return Err(Error::Reverse(format!(
                    "worksheet `{name}` references legacyDrawing without one comments/VML note relationship pair; repair the worksheet relationships"
                )));
            }
            if !metadata.notes.is_empty() {
                validate_recompiled_note_geometry(&metadata, &name)?;
            }
            let sheet_name = names_by_target.get(name.as_str()).ok_or_else(|| {
                Error::Reverse(format!(
                    "worksheet part `{name}` is not linked from workbook.xml; repair the workbook relationships"
                ))
            })?;
            styles_by_sheet.insert((*sheet_name).to_owned(), metadata);
        }
        if name == "docProps/core.xml" {
            let mut xml = String::new();
            std::io::Read::read_to_string(&mut part, &mut xml)
                .map_err(|error| Error::Reverse(error.to_string()))?;
            workbook_properties = audit_core_properties(&xml)?;
        }
    }
    let mut referenced_drawing_parts = std::collections::HashSet::new();
    let mut referenced_chart_parts = std::collections::HashSet::new();
    let mut referenced_media_parts = std::collections::HashSet::new();
    let mut processed_drawing_sheets = std::collections::HashSet::new();
    for (worksheet_part, relationships) in drawing_relationships_by_sheet {
        processed_drawing_sheets.insert(worksheet_part.clone());
        let sheet_name = names_by_target
            .get(worksheet_part.as_str())
            .ok_or_else(|| Error::Reverse(format!(
                "chart drawing relationship belongs to unlinked worksheet `{worksheet_part}`; repair workbook relationships"
            )))?;
        let metadata = styles_by_sheet.get_mut(*sheet_name).ok_or_else(|| {
            Error::Reverse(format!("missing metadata for worksheet `{sheet_name}`"))
        })?;
        let Some(relationship_id) = metadata.drawing_relationship_id.as_deref() else {
            if relationships.is_empty() {
                continue;
            }
            return Err(Error::Reverse(format!(
                "worksheet `{worksheet_part}` has drawing relationships but no `<drawing r:id>` reference; repair the worksheet XML"
            )));
        };
        let target = relationships.get(relationship_id).ok_or_else(|| Error::Reverse(format!(
            "worksheet `{worksheet_part}` drawing relationship `{relationship_id}` is missing; repair its .rels file"
        )))?;
        let drawing_path = resolve_opc_target(&worksheet_part, target)?;
        if !referenced_drawing_parts.insert(drawing_path.clone()) {
            return Err(Error::Reverse(format!(
                "drawing part `{drawing_path}` is shared by multiple worksheets; use one drawing per worksheet"
            )));
        }
        let drawing_xml = read_zip_part(&mut archive, &drawing_path)?;
        let drawing_relationship_part = opc_relationship_part(&drawing_path);
        let drawing_relationship_xml = read_zip_part(&mut archive, &drawing_relationship_part)?;
        let drawing_relationships =
            parse_generic_relationships(&drawing_relationship_xml, &drawing_relationship_part)?;
        let drawing_anchors = parse_chart_drawing(&drawing_xml, &drawing_path)?;
        if drawing_anchors.is_empty() {
            return Err(Error::Reverse(format!(
                "drawing `{drawing_path}` contains no chart or picture anchors; repair its drawing XML"
            )));
        }
        let mut used_drawing_relationships = std::collections::HashSet::new();
        for anchor in drawing_anchors {
            let drawing_target = drawing_relationships.get(&anchor.relationship_id).ok_or_else(|| Error::Reverse(format!(
                "drawing `{drawing_path}` object references missing relationship `{}`; repair its drawing relationships",
                anchor.relationship_id
            )))?;
            if !used_drawing_relationships.insert(anchor.relationship_id.clone())
                && anchor.kind != DrawingObjectKind::Image
            {
                return Err(Error::Reverse(format!(
                    "drawing `{drawing_path}` reuses chart relationship `{}`; give every chart its own relationship",
                    anchor.relationship_id
                )));
            }
            match anchor.kind {
                DrawingObjectKind::Chart => {
                    let chart_path = resolve_opc_target(&drawing_path, drawing_target)?;
                    if !referenced_chart_parts.insert(chart_path.clone()) {
                        return Err(Error::Reverse(format!(
                            "chart part `{chart_path}` is referenced more than once; give each chart one drawing anchor"
                        )));
                    }
                    let chart_xml = read_zip_part(&mut archive, &chart_path)?;
                    let index = metadata.charts.len();
                    metadata.charts.push(parse_chart_xml(
                        &chart_xml,
                        &chart_path,
                        anchor,
                        metadata,
                        &worksheet_names,
                    )?);
                    metadata
                        .drawing_order
                        .push(WorksheetDrawingOrder::Chart(index));
                }
                DrawingObjectKind::Image => {
                    let image_path = resolve_opc_target(&drawing_path, drawing_target)?;
                    referenced_media_parts.insert(image_path.clone());
                    let image_bytes = read_zip_bytes(&mut archive, &image_path)?;
                    let image = parse_drawing_image(&image_bytes, &image_path, anchor, metadata)?;
                    let index = metadata.images.len();
                    metadata.images.push(image);
                    metadata
                        .drawing_order
                        .push(WorksheetDrawingOrder::Image(index));
                }
                DrawingObjectKind::Unknown => {
                    return Err(Error::Reverse(format!(
                        "drawing `{drawing_path}` contains an unrecognized object; use a chart or raster picture"
                    )));
                }
            }
        }
        if used_drawing_relationships.len() != drawing_relationships.len() {
            return Err(Error::Reverse(format!(
                "drawing `{drawing_path}` contains an unused chart/image relationship; remove it or add its corresponding drawing anchor"
            )));
        }
    }
    for (sheet_name, metadata) in &styles_by_sheet {
        if metadata.drawing_relationship_id.is_some()
            && !names_by_target.iter().any(|(part, name)| {
                *name == sheet_name.as_str() && processed_drawing_sheets.contains(*part)
            })
        {
            return Err(Error::Reverse(format!(
                "worksheet `{sheet_name}` contains `<drawing>` but has no worksheet relationship part; restore the matching .rels file"
            )));
        }
    }
    for part in archive.file_names() {
        if part.starts_with("xl/drawings/drawing")
            && std::path::Path::new(part)
                .extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case("xml"))
            && !referenced_drawing_parts.contains(part)
        {
            return Err(Error::Reverse(format!(
                "drawing part `{part}` is not referenced by a worksheet; remove the orphan part or repair its worksheet relationship"
            )));
        }
        if part.starts_with("xl/charts/chart")
            && std::path::Path::new(part)
                .extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case("xml"))
            && !referenced_chart_parts.contains(part)
        {
            return Err(Error::Reverse(format!(
                "chart part `{part}` is not referenced by a drawing; remove the orphan part or repair its drawing relationship"
            )));
        }
        if part.starts_with("xl/media/image") && !referenced_media_parts.contains(part) {
            return Err(Error::Reverse(format!(
                "image part `{part}` is not referenced by a worksheet picture; remove the orphan media or repair its drawing relationship"
            )));
        }
    }
    let mut referenced_table_parts = std::collections::HashSet::new();
    for (worksheet_part, relationships) in table_relationships_by_sheet {
        let sheet_name = names_by_target
            .get(worksheet_part.as_str())
            .ok_or_else(|| {
                Error::Reverse(format!(
                    "table relationships belong to unlinked worksheet `{worksheet_part}`"
                ))
            })?;
        let metadata = styles_by_sheet.get_mut(*sheet_name).ok_or_else(|| {
            Error::Reverse(format!("missing metadata for worksheet `{sheet_name}`"))
        })?;
        let referenced: std::collections::HashSet<_> =
            metadata.table_relationship_ids.iter().collect();
        if referenced.len() != metadata.table_relationship_ids.len()
            || referenced.len() != relationships.len()
        {
            return Err(Error::Reverse(format!(
                "worksheet `{worksheet_part}` tableParts and table relationships do not match; repair the relationship IDs and tableParts count"
            )));
        }
        for relation_id in &metadata.table_relationship_ids {
            let target = relationships.get(relation_id).ok_or_else(|| Error::Reverse(format!("tablePart relationship `{relation_id}` in `{worksheet_part}` has no matching table relationship")))?;
            let table_path = resolve_table_target(target, &worksheet_part)?;
            if !referenced_table_parts.insert(table_path.clone()) {
                return Err(Error::Reverse(format!(
                    "table part `{table_path}` is referenced more than once; repair worksheet relationships"
                )));
            }
            let xml = read_zip_part(&mut archive, &table_path)?;
            metadata
                .tables
                .push(parse_table_xml(&xml, &table_path, &differential_formats)?);
        }
    }
    let mut table_names = std::collections::HashSet::new();
    for metadata in styles_by_sheet.values() {
        for table in &metadata.tables {
            if !table_names.insert(table.name.to_lowercase()) {
                return Err(Error::Reverse(format!(
                    "table name `{}` is duplicated across the workbook; rename tables to unique identifiers",
                    table.name
                )));
            }
            let range = parse_a1_range(&table.range).ok_or_else(|| {
                Error::Reverse(format!(
                    "table `{}` has an invalid range `{}`",
                    table.name, table.range
                ))
            })?;
            if metadata.tables.iter().any(|other| {
                other.name != table.name
                    && parse_a1_range(&other.range).is_some_and(|other_range| {
                        range.0 <= other_range.2
                            && range.2 >= other_range.0
                            && range.1 <= other_range.3
                            && range.3 >= other_range.1
                    })
            }) {
                return Err(Error::Reverse(format!(
                    "table `{}` overlaps another table; make table ranges disjoint",
                    table.name
                )));
            }
            if metadata.merged_ranges.iter().any(|merged| {
                range.0 <= merged.last_row
                    && range.2 >= merged.first_row
                    && range.1 <= merged.last_column
                    && range.3 >= merged.first_column
            }) {
                return Err(Error::Reverse(format!(
                    "table `{}` intersects merged cells; remove the merge or move the table",
                    table.name
                )));
            }
            if metadata
                .auto_filter
                .as_deref()
                .and_then(parse_a1_range)
                .is_some_and(|filter| {
                    range.0 <= filter.2
                        && range.2 >= filter.0
                        && range.1 <= filter.3
                        && range.3 >= filter.1
                })
            {
                return Err(Error::Reverse(format!(
                    "table `{}` overlaps a worksheet autoFilter; remove the separate filter or move the table",
                    table.name
                )));
            }
        }
    }
    let packaged_table_parts: std::collections::HashSet<_> = archive
        .file_names()
        .filter(|name| {
            name.starts_with("xl/tables/table")
                && std::path::Path::new(name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        })
        .map(str::to_owned)
        .collect();
    if packaged_table_parts != referenced_table_parts {
        return Err(Error::Reverse("table package parts and worksheet relationships do not match; repair or remove orphan table parts".into()));
    }
    let has_dynamic_formula = styles_by_sheet.values().any(|metadata| {
        metadata
            .array_formula_ranges
            .values()
            .any(|formula| formula.is_dynamic)
    });
    if has_dynamic_metadata != has_dynamic_formula {
        return Err(Error::Reverse(
            "workbook dynamic-array metadata and worksheet cm markers do not match; repair the metadata part, workbook relationship, or dynamic formula cells".into(),
        ));
    }
    let (active_sheet_index, first_tab_index) = audit_workbook_xml(
        &workbook_xml,
        &sheet_order,
        &mut styles_by_sheet,
        &mut workbook_properties,
    )?;
    let active_sheet = sheet_order.get(active_sheet_index).ok_or_else(|| {
        Error::Reverse(format!(
            "workbook activeTab {active_sheet_index} references a missing worksheet; repair workbook.xml"
        ))
    })?;
    if visibility_by_sheet.get(active_sheet) != Some(&WorksheetVisibility::Visible) {
        return Err(Error::Reverse(format!(
            "workbook active worksheet `{active_sheet}` is hidden; choose a visible active worksheet or repair workbook.xml"
        )));
    }
    if active_sheet_index > 0
        && let Some(metadata) = styles_by_sheet.get_mut(active_sheet)
    {
        metadata.active_worksheet = Some(());
    }
    apply_selected_worksheet_state(&mut styles_by_sheet, &visibility_by_sheet, active_sheet)?;
    let first_tab = sheet_order.get(first_tab_index).ok_or_else(|| {
        Error::Reverse(format!(
            "workbook firstSheet {} references a missing worksheet; repair workbook.xml",
            first_tab_index.saturating_add(1)
        ))
    })?;
    if visibility_by_sheet.get(first_tab) != Some(&WorksheetVisibility::Visible) {
        return Err(Error::Reverse(format!(
            "workbook first visible tab `{first_tab}` is hidden; choose a visible worksheet or repair workbook.xml"
        )));
    }
    if first_tab_index > 0
        && let Some(metadata) = styles_by_sheet.get_mut(first_tab)
    {
        metadata.first_tab = Some(());
    }
    for (name, visibility) in visibility_by_sheet {
        if let Some(metadata) = styles_by_sheet.get_mut(&name) {
            metadata.visibility = visibility;
        }
    }
    Ok((styles_by_sheet, workbook_properties, formats))
}

fn apply_selected_worksheet_state(
    styles_by_sheet: &mut HashMap<String, WorksheetMetadata>,
    visibility_by_sheet: &HashMap<String, WorksheetVisibility>,
    active_sheet: &str,
) -> xlsx_jsx::Result<()> {
    for (name, metadata) in styles_by_sheet {
        let selected = metadata.tab_selected.unwrap_or(false);
        if selected && visibility_by_sheet.get(name) != Some(&WorksheetVisibility::Visible) {
            return Err(Error::Reverse(format!(
                "worksheet `{name}` is selected but hidden; make it visible or clear sheetView@tabSelected"
            )));
        }
        let is_active = name == active_sheet;
        if is_active && !selected {
            return Err(Error::Reverse(format!(
                "worksheet `{name}` is active but `sheetView@tabSelected` is false; rust_xlsxwriter always selects the active worksheet, so this state cannot be recompiled losslessly. Set `tabSelected=true` or remove the contradictory active/unselected state"
            )));
        }
        if selected != is_active {
            metadata.selected = Some(selected);
        }
    }
    Ok(())
}

#[derive(Clone, Copy, Default, PartialEq, Eq)]
enum WorksheetVisibility {
    #[default]
    Visible,
    Hidden,
    VeryHidden,
}

#[derive(Clone, Copy, Default, PartialEq, Eq)]
enum WorksheetView {
    #[default]
    Normal,
    PageLayout,
    PageBreakPreview,
}

impl WorksheetView {
    fn jsx_value(self) -> Option<&'static str> {
        match self {
            Self::Normal => None,
            Self::PageLayout => Some("pageLayout"),
            Self::PageBreakPreview => Some("pageBreakPreview"),
        }
    }
}

impl WorksheetVisibility {
    fn jsx_value(self) -> Option<&'static str> {
        match self {
            Self::Visible => None,
            Self::Hidden => Some("hidden"),
            Self::VeryHidden => Some("veryHidden"),
        }
    }
}

fn parse_worksheet_visibility(
    xml: &str,
    sheet_order: &[String],
) -> xlsx_jsx::Result<HashMap<String, WorksheetVisibility>> {
    let mut reader = XmlReader::from_str(xml);
    let mut visibility_by_sheet = HashMap::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event))
                if event.local_name().as_ref() == b"sheet" =>
            {
                let mut name = None;
                let mut visibility = WorksheetVisibility::Visible;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?;
                    match attribute.key.local_name().as_ref() {
                        b"name" => name = Some(value.into_owned()),
                        b"state" => {
                            visibility = match value.as_ref() {
                                "visible" => WorksheetVisibility::Visible,
                                "hidden" => WorksheetVisibility::Hidden,
                                "veryHidden" => WorksheetVisibility::VeryHidden,
                                _ => {
                                    return Err(Error::Reverse(format!(
                                        "workbook worksheet state {value} is not supported; use visible, hidden, or veryHidden"
                                    )));
                                }
                            };
                        }
                        _ => {}
                    }
                }
                let name =
                    name.ok_or_else(|| Error::Reverse("worksheet is missing its name".into()))?;
                if visibility_by_sheet
                    .insert(name.clone(), visibility)
                    .is_some()
                {
                    return Err(Error::Reverse(format!(
                        "workbook repeats worksheet {name}; repair workbook.xml"
                    )));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid workbook.xml: {error}"))),
        }
    }
    let first_sheet = sheet_order
        .first()
        .ok_or_else(|| Error::Reverse("workbook contains no worksheets".into()))?;
    if visibility_by_sheet
        .get(first_sheet)
        .copied()
        .unwrap_or_default()
        != WorksheetVisibility::Visible
    {
        return Err(Error::Reverse(
            "the first worksheet must be visible; move a visible worksheet first because this DSL preserves the default active sheet"
                .into(),
        ));
    }
    for sheet in sheet_order {
        if !visibility_by_sheet.contains_key(sheet) {
            return Err(Error::Reverse(format!(
                "workbook worksheet {sheet} is missing its visibility record; repair workbook.xml"
            )));
        }
    }
    Ok(visibility_by_sheet)
}

fn parse_style_tables(
    archive: &mut ZipArchive<std::fs::File>,
) -> xlsx_jsx::Result<(Vec<CellFontStyle>, Vec<CellFontStyle>)> {
    if archive.file_names().any(|name| name == "xl/styles.xml") {
        let styles_xml = read_zip_part(archive, "xl/styles.xml")?;
        Ok((
            parse_styles_xml(&styles_xml)?,
            parse_differential_styles_xml(&styles_xml)?,
        ))
    } else {
        Ok((
            vec![CellFontStyle {
                support: StyleSupport::Supported,
                ..Default::default()
            }],
            Vec::new(),
        ))
    }
}

fn read_zip_part<R: std::io::Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> xlsx_jsx::Result<String> {
    let mut part = archive
        .by_name(name)
        .map_err(|error| Error::Reverse(format!("missing package part `{name}`: {error}")))?;
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut part, &mut xml)
        .map_err(|error| Error::Reverse(error.to_string()))?;
    Ok(xml)
}

fn read_zip_bytes<R: std::io::Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> xlsx_jsx::Result<Vec<u8>> {
    let mut part = archive
        .by_name(name)
        .map_err(|error| Error::Reverse(format!("missing package part `{name}`: {error}")))?;
    let mut bytes = Vec::new();
    std::io::Read::read_to_end(&mut part, &mut bytes)
        .map_err(|error| Error::Reverse(format!("cannot read image part `{name}`: {error}")))?;
    Ok(bytes)
}

fn resolve_table_target(target: &str, worksheet_part: &str) -> xlsx_jsx::Result<String> {
    if target.starts_with('/') || target.contains(':') || target.contains('\\') {
        return Err(Error::Reverse(format!(
            "table target `{target}` from `{worksheet_part}` is not a safe internal relationship; use a relative `../tables/tableN.xml` target"
        )));
    }
    let resolved = match target.strip_prefix("../") {
        Some(rest) if rest.starts_with("tables/table") => format!("xl/{rest}"),
        _ => {
            return Err(Error::Reverse(format!(
                "table target `{target}` from `{worksheet_part}` is outside the supported `../tables/tableN.xml` location"
            )));
        }
    };
    let file = resolved
        .strip_prefix("xl/tables/table")
        .and_then(|v| v.strip_suffix(".xml"));
    if !file.is_some_and(|v| !v.is_empty() && v.parse::<u32>().is_ok()) {
        return Err(Error::Reverse(format!(
            "table target `{target}` is not a canonical table part name"
        )));
    }
    Ok(resolved)
}

#[expect(
    clippy::too_many_lines,
    reason = "Strict parsing of a table part preserves root, columns, filter, and style semantics together."
)]
fn parse_table_xml(
    xml: &str,
    part: &str,
    differential_formats: &[CellFontStyle],
) -> xlsx_jsx::Result<WorksheetTable> {
    let mut reader = XmlReader::from_str(xml);
    let mut table = WorksheetTable {
        style: "TableStyleMedium9".into(),
        ..Default::default()
    };
    let mut in_columns = false;
    let mut saw_root = false;
    let mut saw_style_info = false;
    let mut declared_column_count = None;
    let mut current_column: Option<usize> = None;
    let mut formula_element: Option<&'static str> = None;
    let mut in_extension_list = false;
    let mut in_extension = false;
    let mut extension_has_table = false;
    let mut saw_accessibility_extension = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event)) => {
                let tag = event.local_name();
                let tag = tag.as_ref();
                let mut attrs = HashMap::new();
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let key =
                        String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    attrs.insert(key, value);
                }
                match tag {
                    b"extLst" if !in_extension_list && !in_extension => {
                        if !attrs.is_empty() {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` extLst has unsupported attributes; preserve only the writer's accessibility extension"
                            )));
                        }
                        in_extension_list = true;
                    }
                    b"ext" if in_extension_list && !in_extension => {
                        if attrs.get("uri").map(String::as_str)
                            != Some("{504A1905-F514-4f6f-8877-14C23A59335A}")
                            || attrs.get("x14").map(String::as_str)
                                != Some(
                                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                                )
                            || attrs.len() != 2
                        {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` contains an unsupported extension; retain only the pinned writer's x14 accessibility extension"
                            )));
                        }
                        in_extension = true;
                        extension_has_table = false;
                    }
                    b"table" if in_extension => {
                        if event.name().as_ref() != b"x14:table" {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` extension uses an unexpected table namespace; retain the x14 accessibility element"
                            )));
                        }
                        if saw_accessibility_extension {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` contains duplicate x14 table accessibility metadata"
                            )));
                        }
                        if attrs
                            .keys()
                            .any(|key| key != "altText" && key != "altTextSummary")
                        {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` x14 table has unsupported accessibility attributes; use only altText and altTextSummary"
                            )));
                        }
                        table.alt_text_title = attrs.get("altText").cloned();
                        table.alt_text = attrs.get("altTextSummary").cloned();
                        saw_accessibility_extension = true;
                        extension_has_table = true;
                    }
                    b"table" => {
                        if in_extension_list || in_extension {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` contains an unsupported extension element; repair the x14 accessibility metadata"
                            )));
                        }
                        saw_root = true;
                        let totals = attrs
                            .get("totalsRowCount")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(0);
                        if totals > 1 {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` declares more than one totals row; only one is representable"
                            )));
                        }
                        table.totals_row = totals == 1
                            || attrs
                                .get("totalsRowShown")
                                .is_some_and(|v| v == "1" || v == "true");
                        if attrs.get("headerRowCount").is_some_and(|v| v != "1") {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` is headerless or has multiple header rows; use exactly one header row"
                            )));
                        }
                        table.name = attrs
                            .get("displayName")
                            .or_else(|| attrs.get("name"))
                            .cloned()
                            .ok_or_else(|| {
                                Error::Reverse(format!("table part `{part}` is missing its name"))
                            })?;
                        table.range = attrs.get("ref").cloned().ok_or_else(|| {
                            Error::Reverse(format!("table part `{part}` is missing its ref range"))
                        })?;
                        if attrs.keys().any(|key| {
                            ![
                                "xmlns",
                                "id",
                                "name",
                                "displayName",
                                "ref",
                                "headerRowCount",
                                "totalsRowCount",
                                "totalsRowShown",
                                "insertRow",
                                "insertRowShift",
                                "published",
                                "defaultTableStyle",
                                "defaultPivotStyle",
                            ]
                            .contains(&key.as_str())
                        }) {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` contains unsupported root metadata; remove custom table metadata or extend the DSL and reverse reader"
                            )));
                        }
                    }
                    b"tableColumns" => {
                        in_columns = true;
                        declared_column_count = attrs
                            .get("count")
                            .and_then(|value| value.parse::<usize>().ok());
                        if attrs.keys().any(|key| key != "count") {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` has unsupported tableColumns metadata"
                            )));
                        }
                    }
                    b"tableColumn" if in_columns => {
                        if attrs.keys().any(|key| {
                            ![
                                "id",
                                "name",
                                "totalsRowFunction",
                                "totalsRowLabel",
                                "dataDxfId",
                            ]
                            .contains(&key.as_str())
                        }) || !attrs.contains_key("name")
                        {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` contains unsupported column metadata; remove it or add explicit reverse support"
                            )));
                        }
                        table.headers.push(attrs["name"].clone());
                        let format = attrs
                            .get("dataDxfId")
                            .map(|value| {
                                let index = value.parse::<usize>().map_err(|error| {
                                    Error::Reverse(format!(
                                        "table part `{part}` has invalid dataDxfId `{value}`: {error}"
                                    ))
                                })?;
                                let style = differential_formats.get(index).ok_or_else(|| {
                                    Error::Reverse(format!(
                                        "table part `{part}` references missing differential style {index}; repair styles.xml"
                                    ))
                                })?;
                                if style.support == StyleSupport::Unsupported {
                                    return Err(Error::Reverse(format!(
                                        "table part `{part}` column uses a differential format outside the supported font emphasis/decoration/script/color, fill, number-format, and edge-border subset; simplify the table format or add explicit DSL and reverse support"
                                    )));
                                }
                                if style.font_scheme.is_some() {
                                    return Err(Error::Reverse(format!(
                                        "table part `{part}` column uses a differential font scheme that TableColumn.format cannot represent; remove the theme scheme or use a regular cell/row/column style"
                                    )));
                                }
                                Ok(style.clone())
                            })
                            .transpose()?;
                        table.columns.push(WorksheetTableColumn {
                            total_label: attrs.get("totalsRowLabel").cloned(),
                            total_function: attrs.get("totalsRowFunction").map(|v| {
                                match v.as_str() {
                                    "countNums" => "countNumbers",
                                    "stdDev" => "stdDev",
                                    _ => v,
                                }
                                .to_owned()
                            }),
                            format,
                            ..Default::default()
                        });
                        current_column = Some(table.columns.len() - 1);
                    }
                    b"calculatedColumnFormula" | b"totalsRowFormula" => {
                        if !attrs.is_empty() {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` uses formula attributes that TableColumn cannot represent; remove those attributes or extend the DSL"
                            )));
                        }
                        formula_element = Some(if tag == b"calculatedColumnFormula" {
                            "calculated"
                        } else {
                            "total"
                        });
                    }
                    b"autoFilter" => {
                        if attrs.keys().any(|key| key != "ref") {
                            return Err(Error::Reverse(format!(
                                "table `{}` has filter criteria that are not supported; remove the table filter criteria or add explicit DSL support",
                                table.name
                            )));
                        }
                        if attrs.get("ref").is_some_and(|filter_range| {
                            let expected = if table.totals_row {
                                parse_a1_range(&table.range).map_or_else(
                                    || table.range.clone(),
                                    |(r0, c0, r1, c1)| {
                                        format!(
                                            "{}:{}",
                                            format_a1_cell_reference(r0, c0),
                                            format_a1_cell_reference(r1.saturating_sub(1), c1)
                                        )
                                    },
                                )
                            } else {
                                table.range.clone()
                            };
                            filter_range != &expected
                        }) {
                            return Err(Error::Reverse(format!(
                                "table `{}` autoFilter range does not match the table range; repair the table XML",
                                table.name
                            )));
                        }
                        table.auto_filter = true;
                    }
                    b"tableStyleInfo" => {
                        saw_style_info = true;
                        if attrs.keys().any(|key| {
                            ![
                                "name",
                                "showFirstColumn",
                                "showLastColumn",
                                "showRowStripes",
                                "showColumnStripes",
                            ]
                            .contains(&key.as_str())
                        }) {
                            return Err(Error::Reverse(format!(
                                "table part `{part}` has unmodeled table style metadata"
                            )));
                        }
                        table.style = attrs.get("name").cloned().ok_or_else(|| Error::Reverse(format!("table part `{part}` tableStyleInfo is missing name; the effective style cannot be preserved losslessly")))?;
                        table.first_column = attrs
                            .get("showFirstColumn")
                            .is_some_and(|v| v == "1" || v == "true");
                        table.last_column = attrs
                            .get("showLastColumn")
                            .is_some_and(|v| v == "1" || v == "true");
                        table.banded_rows = attrs
                            .get("showRowStripes")
                            .is_none_or(|v| v == "1" || v == "true");
                        table.banded_columns = attrs
                            .get("showColumnStripes")
                            .is_some_and(|v| v == "1" || v == "true");
                    }
                    _ => {
                        return Err(Error::Reverse(format!(
                            "table part `{part}` contains unsupported `<{}>` metadata; remove it or add explicit reverse support",
                            String::from_utf8_lossy(tag)
                        )));
                    }
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"ext" => {
                if !in_extension {
                    return Err(Error::Reverse(format!(
                        "table part `{part}` has an unmatched extension end tag"
                    )));
                }
                if !extension_has_table {
                    return Err(Error::Reverse(format!(
                        "table part `{part}` contains an empty extension; remove it or add supported x14 accessibility metadata"
                    )));
                }
                in_extension = false;
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"extLst" => {
                if !in_extension_list || in_extension {
                    return Err(Error::Reverse(format!(
                        "table part `{part}` has malformed extension nesting"
                    )));
                }
                in_extension_list = false;
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"tableColumns" => {
                in_columns = false;
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"tableColumn" => {
                current_column = None;
            }
            Ok(Event::End(event))
                if event.local_name().as_ref() == b"calculatedColumnFormula"
                    || event.local_name().as_ref() == b"totalsRowFormula" =>
            {
                formula_element = None;
            }
            Ok(Event::Text(event)) if formula_element.is_some() => {
                let value = event
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?
                    .into_owned();
                let index = current_column.ok_or_else(|| {
                    Error::Reverse(format!(
                        "table part `{part}` has a formula outside a table column"
                    ))
                })?;
                if formula_element == Some("calculated") {
                    table.columns[index].formula = Some(value);
                } else {
                    table.columns[index].total_formula = Some(value);
                    table.columns[index].total_function = None;
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid table XML in `{part}`: {error}"
                )));
            }
        }
    }
    if in_extension_list || in_extension {
        return Err(Error::Reverse(format!(
            "table part `{part}` has incomplete extension nesting; repair the x14 accessibility metadata"
        )));
    }
    if !saw_root || table.headers.is_empty() {
        return Err(Error::Reverse(format!(
            "table part `{part}` is missing its root or columns"
        )));
    }
    if declared_column_count.is_some_and(|count| count != table.headers.len()) {
        return Err(Error::Reverse(format!(
            "table part `{part}` tableColumns count does not match its column entries; repair the count"
        )));
    }
    if !saw_style_info {
        return Err(Error::Reverse(format!(
            "table part `{part}` has no tableStyleInfo; the effective workbook style cannot be represented losslessly"
        )));
    }
    if !xlsx_jsx::is_supported_table_style(&table.style) {
        return Err(Error::Reverse(format!(
            "table `{}` uses unsupported style `{}`; convert it to a built-in TableStyle or add explicit style support",
            table.name, table.style
        )));
    }
    let unique_headers: std::collections::HashSet<_> = table
        .headers
        .iter()
        .map(|header| header.to_lowercase())
        .collect();
    if unique_headers.len() != table.headers.len()
        || table.headers.iter().any(|header| header.trim().is_empty())
    {
        return Err(Error::Reverse(format!(
            "table `{}` has empty or duplicate column names; repair tableColumns and worksheet header cells",
            table.name
        )));
    }
    Ok(table)
}

fn worksheet_part_from_relationship_part(relationship_part: &str) -> Option<String> {
    let file_name = relationship_part
        .strip_prefix("xl/worksheets/_rels/")?
        .strip_suffix(".rels")?;
    Some(format!("xl/worksheets/{file_name}"))
}

fn resolve_opc_target(source_part: &str, target: &str) -> xlsx_jsx::Result<String> {
    if target.is_empty() || target.contains('\\') || target.starts_with('/') {
        return Err(Error::Reverse(format!(
            "relationship target `{target}` from `{source_part}` is not a supported package-relative path; repair it to a relative OPC target"
        )));
    }
    let mut parts = source_part
        .split('/')
        .take(source_part.split('/').count().saturating_sub(1))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                if parts.pop().is_none() {
                    return Err(Error::Reverse(format!(
                        "relationship target `{target}` escapes the package root from `{source_part}`"
                    )));
                }
            }
            segment if segment.contains(':') => {
                return Err(Error::Reverse(format!(
                    "relationship target `{target}` from `{source_part}` has an invalid URI segment; use an internal package part"
                )));
            }
            segment => parts.push(segment.to_owned()),
        }
    }
    Ok(parts.join("/"))
}

fn opc_relationship_part(source_part: &str) -> String {
    let mut parts = source_part.rsplitn(2, '/');
    let filename = parts.next().unwrap_or_default();
    let parent = parts.next().unwrap_or_default();
    if parent.is_empty() {
        format!("_rels/{filename}.rels")
    } else {
        format!("{parent}/_rels/{filename}.rels")
    }
}

fn parse_generic_relationships(xml: &str, part: &str) -> xlsx_jsx::Result<HashMap<String, String>> {
    let mut reader = XmlReader::from_str(xml);
    let mut relationships = HashMap::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event))
                if event.local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                let mut relation_type = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"Id" => id = Some(value),
                        b"Target" => target = Some(value),
                        b"Type" => relation_type = Some(value),
                        b"TargetMode" => {
                            return Err(Error::Reverse(format!(
                                "relationship in `{part}` uses TargetMode; chart drawings must target internal package parts"
                            )));
                        }
                        key => {
                            return Err(Error::Reverse(format!(
                                "relationship in `{part}` has unsupported attribute `{}`; remove it or add explicit support",
                                String::from_utf8_lossy(key)
                            )));
                        }
                    }
                }
                let id = id.ok_or_else(|| {
                    Error::Reverse(format!("relationship in `{part}` is missing Id"))
                })?;
                let target = target.ok_or_else(|| {
                    Error::Reverse(format!("relationship `{id}` in `{part}` is missing Target"))
                })?;
                if !relation_type
                    .as_deref()
                    .is_some_and(|value| value.ends_with("/chart") || value.ends_with("/image"))
                {
                    return Err(Error::Reverse(format!(
                        "drawing relationship `{id}` in `{part}` is not a chart or image relationship; only charts and embedded raster images are supported"
                    )));
                }
                if relationships.insert(id.clone(), target).is_some() {
                    return Err(Error::Reverse(format!(
                        "relationship part `{part}` repeats Id `{id}`; repair its relationship IDs"
                    )));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid relationships in `{part}`: {error}"
                )));
            }
        }
    }
    Ok(relationships)
}

#[derive(Clone, Copy, Default, PartialEq, Eq)]
enum DrawingObjectKind {
    #[default]
    Unknown,
    Chart,
    Image,
}

#[derive(Default)]
struct ChartAnchor {
    kind: DrawingObjectKind,
    edit_as: Option<String>,
    relationship_id: String,
    name: String,
    alt_text: Option<String>,
    object_id: Option<u32>,
    decorative: bool,
    absolute_x: Option<u64>,
    absolute_y: Option<u64>,
    transform_width: Option<u64>,
    transform_height: Option<u64>,
    from_row: Option<u32>,
    from_column: Option<u32>,
    to_row: Option<u32>,
    to_column: Option<u32>,
    to_column_offset: Option<u32>,
    to_row_offset: Option<u32>,
    from_column_offset: Option<u32>,
    from_row_offset: Option<u32>,
    from_column_offset_seen: bool,
    from_row_offset_seen: bool,
}

#[expect(
    clippy::too_many_lines,
    reason = "Chart XML parsing validates the complete supported chart structure in one pass."
)]
fn parse_chart_xml(
    xml: &str,
    part: &str,
    anchor: ChartAnchor,
    metadata: &WorksheetMetadata,
    worksheet_names: &std::collections::HashSet<String>,
) -> xlsx_jsx::Result<WorksheetChart> {
    const PX_EMU: u32 = 9_525;
    const DEFAULT_COLUMN_PIXELS: u32 = 64;
    const DEFAULT_ROW_PIXELS: u32 = 20;
    if anchor.edit_as.is_some() {
        return Err(Error::Reverse(format!(
            "chart drawing in `{part}` has a non-default editAs anchor mode; remove custom movement semantics"
        )));
    }
    let from_row = anchor
        .from_row
        .ok_or_else(|| Error::Reverse(format!("chart drawing in `{part}` is missing from/row")))?;
    let from_column = anchor
        .from_column
        .ok_or_else(|| Error::Reverse(format!("chart drawing in `{part}` is missing from/col")))?;
    let to_row = anchor
        .to_row
        .ok_or_else(|| Error::Reverse(format!("chart drawing in `{part}` is missing to/row")))?;
    let to_column = anchor
        .to_column
        .ok_or_else(|| Error::Reverse(format!("chart drawing in `{part}` is missing to/col")))?;
    if !anchor.from_column_offset_seen
        || !anchor.from_row_offset_seen
        || anchor.from_column_offset != Some(0)
        || anchor.from_row_offset != Some(0)
    {
        return Err(Error::Reverse(format!(
            "chart drawing in `{part}` is missing from/colOff or from/rowOff; repair the two-cell anchor"
        )));
    }
    if to_row < from_row || to_column < from_column {
        return Err(Error::Reverse(format!(
            "chart drawing in `{part}` has reversed anchor coordinates; repair from/to positions"
        )));
    }
    if to_row >= 1_048_576 || to_column >= 16_384 {
        return Err(Error::Reverse(format!(
            "chart drawing in `{part}` extends beyond Excel's worksheet grid; repair its anchor coordinates"
        )));
    }
    let column_offset = anchor.to_column_offset.ok_or_else(|| {
        Error::Reverse(format!(
            "chart drawing in `{part}` is missing to/colOff; repair the two-cell anchor"
        ))
    })?;
    let row_offset = anchor.to_row_offset.ok_or_else(|| {
        Error::Reverse(format!(
            "chart drawing in `{part}` is missing to/rowOff; repair the two-cell anchor"
        ))
    })?;
    if column_offset % PX_EMU != 0 || row_offset % PX_EMU != 0 {
        return Err(Error::Reverse(format!(
            "chart drawing in `{part}` ends at fractional-pixel offsets; use whole-pixel chart dimensions"
        )));
    }
    if column_offset / PX_EMU >= DEFAULT_COLUMN_PIXELS || row_offset / PX_EMU >= DEFAULT_ROW_PIXELS
    {
        return Err(Error::Reverse(format!(
            "chart drawing in `{part}` has end offsets outside the default worksheet grid; preserve standard row/column dimensions or extend chart geometry support"
        )));
    }
    if metadata
        .default_row_height_pixels
        .is_some_and(|height| height != DEFAULT_ROW_PIXELS)
        || metadata.rows.iter().any(|(index, row)| {
            *index >= from_row
                && *index <= to_row
                && (row
                    .height
                    .is_some_and(|height| (height - 15.0).abs() > f64::EPSILON)
                    || row.hidden)
        })
    {
        return Err(Error::Reverse(format!(
            "chart drawing in `{part}` crosses custom row heights, so its pixel height is ambiguous; remove those row heights or add explicit chart extents to the DSL"
        )));
    }
    if metadata
        .columns
        .range(
            u16::try_from(from_column).expect("chart column bounds were validated")
                ..=u16::try_from(to_column).expect("chart column bounds were validated"),
        )
        .any(|(_, column)| {
            column
                .width_pixels
                .is_some_and(|width| width != DEFAULT_COLUMN_PIXELS)
                || column.hidden
        })
    {
        return Err(Error::Reverse(format!(
            "chart drawing in `{part}` crosses custom column widths, so its pixel width is ambiguous; remove those widths or add explicit chart extents to the DSL"
        )));
    }
    let width = (to_column - from_column)
        .checked_mul(DEFAULT_COLUMN_PIXELS)
        .and_then(|value| value.checked_add(column_offset / PX_EMU))
        .ok_or_else(|| {
            Error::Reverse(format!(
                "chart width in `{part}` overflows supported pixel units"
            ))
        })?;
    let height = (to_row - from_row)
        .checked_mul(DEFAULT_ROW_PIXELS)
        .and_then(|value| value.checked_add(row_offset / PX_EMU))
        .ok_or_else(|| {
            Error::Reverse(format!(
                "chart height in `{part}` overflows supported pixel units"
            ))
        })?;
    if !(1..=1790).contains(&width) || !(1..=1790).contains(&height) {
        return Err(Error::Reverse(format!(
            "chart in `{part}` measures {width}×{height} pixels outside the supported 1–1790 range; resize it or extend Chart dimensions"
        )));
    }

    let allowed_tags = [
        "chartSpace",
        "lang",
        "chart",
        "title",
        "tx",
        "rich",
        "bodyPr",
        "lstStyle",
        "p",
        "pPr",
        "defRPr",
        "r",
        "rPr",
        "t",
        "endParaRPr",
        "plotArea",
        "layout",
        "areaChart",
        "barChart",
        "barDir",
        "grouping",
        "overlap",
        "lineChart",
        "radarChart",
        "radarStyle",
        "scatterChart",
        "scatterStyle",
        "pieChart",
        "ser",
        "idx",
        "order",
        "tx",
        "strRef",
        "strCache",
        "f",
        "ptCount",
        "pt",
        "v",
        "cat",
        "xVal",
        "yVal",
        "numRef",
        "numCache",
        "formatCode",
        "val",
        "axId",
        "catAx",
        "scaling",
        "orientation",
        "axPos",
        "tickLblPos",
        "crossAx",
        "crosses",
        "auto",
        "lblAlgn",
        "lblOffset",
        "crossBetween",
        "majorTickMark",
        "valAx",
        "majorGridlines",
        "numFmt",
        "legend",
        "legendPos",
        "delete",
        "printSettings",
        "headerFooter",
        "pageMargins",
        "pageSetup",
        "doughnutChart",
        "holeSize",
        "plotVisOnly",
        "marker",
        "symbol",
        "varyColors",
        "firstSliceAng",
        "txPr",
    ];
    let mut reader = XmlReader::from_str(xml);
    let mut stack = Vec::<String>::new();
    let mut chart_type: Option<&'static str> = None;
    let mut chart_grouping_seen = false;
    let mut series = Vec::<WorksheetChartSeries>::new();
    let mut formula_target: Option<&'static str> = None;
    let mut title_source_open = false;
    let mut category_source_open = false;
    let mut chart_title_open = false;
    let mut chart_title_target: Option<&'static str> = None;
    let mut series_literal_title_open = false;
    let mut chart_title: Option<String> = None;
    let mut x_axis_title: Option<String> = None;
    let mut y_axis_title: Option<String> = None;
    let mut legend_seen = false;
    let mut legend_deleted = false;
    let mut legend_position: Option<&'static str> = None;
    let mut hole_size = None;
    let mut hole_size_seen = false;
    let mut first_slice_angle = None;
    let mut first_slice_angle_seen = false;
    let mut radar_style = None;
    let mut radar_style_seen = false;
    let mut scatter_style = None;
    let mut scatter_style_seen = false;
    let mut chart_title_text = String::new();
    let mut title_paragraphs = 0_usize;
    let mut title_runs = 0_usize;
    let mut title_text_nodes = 0_usize;
    loop {
        match reader.read_event() {
            Ok(event @ (Event::Start(_) | Event::Empty(_))) => {
                let is_empty = matches!(event, Event::Empty(_));
                let (Event::Start(event) | Event::Empty(event)) = event else {
                    unreachable!()
                };
                let tag = String::from_utf8_lossy(event.local_name().as_ref()).into_owned();
                if stack.iter().any(|ancestor| ancestor == "title")
                    && ![
                        "tx",
                        "rich",
                        "bodyPr",
                        "lstStyle",
                        "p",
                        "pPr",
                        "defRPr",
                        "r",
                        "rPr",
                        "t",
                        "endParaRPr",
                        "layout",
                    ]
                    .contains(&tag.as_str())
                {
                    return Err(Error::Reverse(format!(
                        "chart part `{part}` has a formula or unsupported structure in the chart title; use one static unformatted title string"
                    )));
                }
                if !allowed_tags.contains(&tag.as_str()) {
                    return Err(Error::Reverse(format!(
                        "chart part `{part}` contains unsupported `<{tag}>` semantics; remove chart titles/formatting or add matching Chart properties and reverse support"
                    )));
                }
                audit_chart_attributes(&event, &tag, part, reader.decoder(), chart_type)?;
                match tag.as_str() {
                    "legend" => {
                        if legend_seen {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` repeats its legend; retain at most one legend"
                            )));
                        }
                        legend_seen = true;
                    }
                    "legendPos" => {
                        if !stack.iter().any(|ancestor| ancestor == "legend")
                            || legend_position.is_some()
                        {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` has an out-of-place or repeated legendPos; keep one position inside legend"
                            )));
                        }
                        for attribute in event.attributes().with_checks(false) {
                            let attribute =
                                attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                            if attribute.key.local_name().as_ref() == b"val" {
                                let value = attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?;
                                legend_position = Some(match value.as_ref() {
                                    "r" => "right",
                                    "l" => "left",
                                    "t" => "top",
                                    "b" => "bottom",
                                    "tr" => "topRight",
                                    _ => unreachable!("legendPos value was audited"),
                                });
                            }
                        }
                    }
                    "delete" => {
                        if stack.last().is_none_or(|ancestor| ancestor != "legend")
                            || legend_deleted
                        {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` has a delete element outside legend or repeats it; remove it or extend chart deletion support"
                            )));
                        }
                        legend_deleted = true;
                    }
                    _ => {}
                }
                if tag == "title" {
                    if chart_type == Some("scatter")
                        && stack.iter().any(|ancestor| ancestor == "valAx")
                    {
                        return Err(Error::Reverse(format!(
                            "chart part `{part}` has a title on a scatter value axis; axis titles are not represented for scatter charts, so remove them or add explicit X/Y value-axis title support"
                        )));
                    }
                    let target = if stack.iter().any(|ancestor| ancestor == "catAx") {
                        "xAxisTitle"
                    } else if stack.iter().any(|ancestor| ancestor == "valAx") {
                        "yAxisTitle"
                    } else if stack.iter().any(|ancestor| ancestor == "chart") {
                        "title"
                    } else {
                        return Err(Error::Reverse(format!(
                            "chart part `{part}` has a title outside Chart or its primary axes; remove it or extend chart-title support"
                        )));
                    };
                    let already_present = match target {
                        "title" => chart_title.is_some(),
                        "xAxisTitle" => x_axis_title.is_some(),
                        "yAxisTitle" => y_axis_title.is_some(),
                        _ => unreachable!("chart title target was selected above"),
                    };
                    if chart_title_open || already_present {
                        return Err(Error::Reverse(format!(
                            "chart part `{part}` repeats `{target}`; retain at most one title for each chart object or primary axis"
                        )));
                    }
                    chart_title_open = true;
                    chart_title_target = Some(target);
                    chart_title_text.clear();
                    title_paragraphs = 0;
                    title_runs = 0;
                    title_text_nodes = 0;
                    if is_empty {
                        return Err(Error::Reverse(format!(
                            "chart part `{part}` has an empty title; use one static title string or remove the title element"
                        )));
                    }
                }
                if chart_title_open {
                    match tag.as_str() {
                        "p" => title_paragraphs += 1,
                        "r" => title_runs += 1,
                        "t" => title_text_nodes += 1,
                        _ => {}
                    }
                }
                match tag.as_str() {
                    "areaChart" => chart_type = Some(set_chart_type(chart_type, "area", part)?),
                    "barChart" => chart_type = Some(set_chart_type(chart_type, "bar", part)?),
                    "lineChart" => chart_type = Some(set_chart_type(chart_type, "line", part)?),
                    "radarChart" => chart_type = Some(set_chart_type(chart_type, "radar", part)?),
                    "scatterChart" => {
                        chart_type = Some(set_chart_type(chart_type, "scatter", part)?);
                    }
                    "scatterStyle" => {
                        if scatter_style_seen {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` repeats scatterStyle; retain one scatter style"
                            )));
                        }
                        scatter_style_seen = true;
                        for attribute in event.attributes().with_checks(false) {
                            let attribute =
                                attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                            if attribute.key.local_name().as_ref() == b"val" {
                                let value = attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?;
                                scatter_style = Some(if value.as_ref() == "lineMarker" {
                                    "lineMarker"
                                } else {
                                    "unsupported"
                                });
                            }
                        }
                    }
                    "pieChart" => chart_type = Some(set_chart_type(chart_type, "pie", part)?),
                    "radarStyle" => {
                        if radar_style_seen {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` repeats radarStyle; retain one radar style"
                            )));
                        }
                        radar_style_seen = true;
                        for attribute in event.attributes().with_checks(false) {
                            let attribute =
                                attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                            if attribute.key.local_name().as_ref() == b"val" {
                                let value = attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?;
                                radar_style = match value.as_ref() {
                                    "standard" => Some("standard"),
                                    "marker" => Some("marker"),
                                    "filled" => Some("filled"),
                                    _ => None,
                                };
                            }
                        }
                    }
                    "grouping" if chart_type.is_some_and(is_groupable_chart_type) => {
                        if chart_grouping_seen {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` repeats chart grouping; retain one chart group"
                            )));
                        }
                        chart_grouping_seen = true;
                        for attribute in event.attributes().with_checks(false) {
                            let attribute =
                                attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                            if attribute.key.local_name().as_ref() == b"val" {
                                let value = attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?;
                                chart_type = Some(
                                    chart_type_for_grouping(
                                        chart_type.expect("chart family was checked"),
                                        &value,
                                    )
                                    .expect("grouping was validated above"),
                                );
                            }
                        }
                    }
                    "doughnutChart" => {
                        chart_type = Some(set_chart_type(chart_type, "doughnut", part)?);
                    }
                    "holeSize" => {
                        if hole_size_seen {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` repeats holeSize; retain one doughnut hole size"
                            )));
                        }
                        hole_size_seen = true;
                        for attribute in event.attributes().with_checks(false) {
                            let attribute =
                                attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                            if attribute.key.local_name().as_ref() == b"val" {
                                let value = attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?;
                                hole_size = value.parse::<u8>().ok();
                            }
                        }
                    }
                    "firstSliceAng" => {
                        if first_slice_angle_seen {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` repeats firstSliceAng; retain one slice angle"
                            )));
                        }
                        first_slice_angle_seen = true;
                        for attribute in event.attributes().with_checks(false) {
                            let attribute =
                                attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                            if attribute.key.local_name().as_ref() == b"val" {
                                let value = attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?;
                                first_slice_angle = value.parse::<u16>().ok();
                            }
                        }
                    }
                    "ser" => series.push(WorksheetChartSeries {
                        values: String::new(),
                        categories: None,
                        name: None,
                        title: None,
                        marker_seen: false,
                        marker_none: false,
                    }),
                    "marker" if stack.iter().any(|ancestor| ancestor == "ser") => {
                        if let Some(current) = series.last_mut() {
                            current.marker_seen = true;
                        }
                    }
                    "symbol"
                        if stack.iter().any(|ancestor| ancestor == "marker")
                            && stack.iter().any(|ancestor| ancestor == "ser") =>
                    {
                        if let Some(current) = series.last_mut() {
                            current.marker_none = true;
                        }
                    }
                    "tx" if !series.is_empty() => title_source_open = true,
                    "v" if !series.is_empty()
                        && stack.last().is_some_and(|ancestor| ancestor == "tx") =>
                    {
                        let current = series.last_mut().expect("series exists");
                        if current.title.is_some() {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` repeats a literal series title; retain one title per series"
                            )));
                        }
                        current.title = Some(String::new());
                        series_literal_title_open = !is_empty;
                    }
                    "cat" | "xVal" if !series.is_empty() => category_source_open = true,
                    "f" if !series.is_empty() => {
                        formula_target = if stack.iter().any(|ancestor| ancestor == "tx") {
                            Some("name")
                        } else if stack
                            .iter()
                            .any(|ancestor| matches!(ancestor.as_str(), "cat" | "xVal"))
                        {
                            Some("categories")
                        } else if stack
                            .iter()
                            .any(|ancestor| matches!(ancestor.as_str(), "yVal" | "val"))
                        {
                            Some("values")
                        } else {
                            None
                        };
                        if formula_target.is_none() {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` has a formula outside supported series name/category/value ranges; remove it or add matching DSL support"
                            )));
                        }
                    }
                    "barDir" => {
                        for attribute in event.attributes().with_checks(false) {
                            let attribute =
                                attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                            if attribute.key.local_name().as_ref() == b"val" {
                                let value = attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?;
                                let value = if value == "col" {
                                    "column"
                                } else if value == "bar" {
                                    "bar"
                                } else {
                                    return Err(Error::Reverse(format!(
                                        "chart part `{part}` has unsupported barDir `{value}`; use column or bar"
                                    )));
                                };
                                chart_type = Some(value);
                            }
                        }
                    }
                    _ => {}
                }
                if !is_empty {
                    stack.push(tag);
                }
            }
            Ok(Event::Text(text)) => {
                if series_literal_title_open {
                    let title_text = text
                        .xml_content(XmlVersion::Implicit1_0)
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    if let Some(current) = series.last_mut() {
                        current.title.get_or_insert_default().push_str(&title_text);
                    }
                }
                if chart_title_open && stack.last().is_some_and(|tag| tag == "t") {
                    let title_text = text
                        .xml_content(XmlVersion::Implicit1_0)
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    chart_title_text.push_str(&title_text);
                }
                if let Some(target) = formula_target.take() {
                    let formula = text
                        .decode()
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    let current = series.last_mut().ok_or_else(|| {
                        Error::Reverse(format!(
                            "chart part `{part}` contains a source formula outside a series"
                        ))
                    })?;
                    match target {
                        "values" if current.values.is_empty() && !formula.is_empty() => {
                            current.values = formula;
                        }
                        "categories" if current.categories.is_none() && !formula.is_empty() => {
                            current.categories = Some(formula);
                        }
                        "name" if current.name.is_none() && !formula.is_empty() => {
                            current.name = Some(formula);
                        }
                        _ => {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` repeats or has an empty series source formula; retain one nonempty source formula per series field"
                            )));
                        }
                    }
                }
            }
            Ok(Event::GeneralRef(reference)) if chart_title_open || series_literal_title_open => {
                let entity = reference
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                let resolved = reference
                    .resolve_char_ref()
                    .map_err(|error| Error::Reverse(error.to_string()))?
                    .map_or_else(
                        || quick_xml::escape::resolve_predefined_entity(&entity).map(str::to_owned),
                        |character| Some(character.to_string()),
                    )
                    .ok_or_else(|| {
                        Error::Reverse(format!(
                            "chart part `{part}` title contains unsupported entity `&{entity};`; use standard XML character references"
                        ))
                    })?;
                if chart_title_open {
                    chart_title_text.push_str(&resolved);
                } else if let Some(current) = series.last_mut() {
                    current.title.get_or_insert_default().push_str(&resolved);
                }
            }
            Ok(Event::End(event)) => {
                let tag = String::from_utf8_lossy(event.local_name().as_ref()).into_owned();
                if tag == "f" {
                    formula_target = None;
                }
                if tag == "v" && series_literal_title_open {
                    series_literal_title_open = false;
                }
                if tag == "title" && chart_title_open {
                    chart_title_open = false;
                    if title_paragraphs != 1 || title_runs != 1 || title_text_nodes != 1 {
                        return Err(Error::Reverse(format!(
                            "chart part `{part}` has a multi-paragraph or multi-run title; use one unformatted title string"
                        )));
                    }
                    if chart_title_text.is_empty() {
                        return Err(Error::Reverse(format!(
                            "chart part `{part}` has an empty or non-literal title; use one static title string"
                        )));
                    }
                    let target = chart_title_target.take().expect("title target is set");
                    let title = std::mem::take(&mut chart_title_text);
                    if title.chars().count() > 255 || !title.chars().all(is_valid_xml_character) {
                        return Err(Error::Reverse(format!(
                            "chart part `{part}` has `{target}` outside the 1–255 valid XML character subset; shorten or repair the title"
                        )));
                    }
                    if chart_formula_sheet_name(&title).is_some()
                        && chart_formula_has_valid_reference(&title, false)
                    {
                        return Err(Error::Reverse(format!(
                            "chart part `{part}` has literal `{target}` `{title}` that the writer would reinterpret as a cell reference; change the title text or use a source reference"
                        )));
                    }
                    match target {
                        "title" => chart_title = Some(title),
                        "xAxisTitle" => x_axis_title = Some(title),
                        "yAxisTitle" => y_axis_title = Some(title),
                        _ => unreachable!("chart title target was selected above"),
                    }
                }
                if tag == "tx" && std::mem::take(&mut title_source_open) {
                    let current = series.last().ok_or_else(|| {
                        Error::Reverse(format!(
                            "chart part `{part}` contains a series title outside a data series"
                        ))
                    })?;
                    if current.name.is_none() {
                        let Some(title) =
                            current.title.as_deref().filter(|title| !title.is_empty())
                        else {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` has a series title without a source or literal value; add Series.name as a cell reference or Series.title as static text"
                            )));
                        };
                        if title.chars().count() > 255 || !title.chars().all(is_valid_xml_character)
                        {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` has a literal series title outside the 1–255 valid XML character subset; shorten or repair it"
                            )));
                        }
                        if chart_formula_sheet_name(title).is_some()
                            && chart_formula_has_valid_reference(title, false)
                        {
                            return Err(Error::Reverse(format!(
                                "chart part `{part}` has a literal series title `{title}` that the writer would reinterpret as a cell reference; change the title text or use a source reference"
                            )));
                        }
                    }
                }
                if matches!(tag.as_str(), "cat" | "xVal")
                    && std::mem::take(&mut category_source_open)
                    && series.last().is_some_and(|item| item.categories.is_none())
                {
                    return Err(Error::Reverse(format!(
                        "chart part `{part}` has literal categories instead of a source-cell range; use a range reference or extend `Series.categories`"
                    )));
                }
                if tag == "ser" && series.last().is_some_and(|item| item.values.is_empty()) {
                    return Err(Error::Reverse(format!(
                        "chart part `{part}` has a series without a value range; add a valid c:val/c:numRef/c:f source"
                    )));
                }
                stack.pop();
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid chart XML in `{part}`: {error}"
                )));
            }
        }
    }
    let chart_type = chart_type.ok_or_else(|| {
        Error::Reverse(format!(
            "chart part `{part}` has no supported chart type; use an area, column, bar, line, pie, or doughnut chart"
        ))
    })?;
    let chart_type = resolve_radar_chart_type(chart_type, radar_style, &series, part)?;
    let chart_type = resolve_scatter_chart_type(chart_type, scatter_style, &series, part)?;
    if series.is_empty() {
        return Err(Error::Reverse(format!(
            "chart part `{part}` contains no data series; add at least one series"
        )));
    }
    if matches!(chart_type, "pie" | "doughnut") && series.len() != 1 {
        return Err(Error::Reverse(format!(
            "chart part `{part}` has {} series in a pie or doughnut chart; use exactly one data series or split it into separate charts",
            series.len()
        )));
    }
    if matches!(chart_type, "pie" | "doughnut")
        && (x_axis_title.is_some() || y_axis_title.is_some())
    {
        return Err(Error::Reverse(format!(
            "chart part `{part}` is a pie or doughnut chart with axis titles; remove the titles or use a chart type with X/Y axes"
        )));
    }
    if chart_type == "doughnut" && hole_size_seen && hole_size.is_none() {
        return Err(Error::Reverse(format!(
            "chart part `{part}` has a doughnut holeSize without a valid value; set val to an integer from 10 through 90"
        )));
    }
    if first_slice_angle_seen && first_slice_angle.is_none() {
        return Err(Error::Reverse(format!(
            "chart part `{part}` has firstSliceAng without a valid value; set val to an integer from 0 through 360"
        )));
    }
    if chart_type == "doughnut" && hole_size.is_none() {
        // OOXML's absent holeSize default is 10; the pinned writer emits 50.
        hole_size = Some(10);
    }
    let show_legend = legend_seen && !legend_deleted;
    if !show_legend && legend_position.is_some_and(|position| position != "right") {
        return Err(Error::Reverse(format!(
            "chart part `{part}` has a hidden legend with a non-default position; remove legendPos because it cannot be reproduced while hidden"
        )));
    }
    for item in &series {
        for (property, formula, require_range) in [
            ("values", Some(item.values.as_str()), true),
            ("categories", item.categories.as_deref(), true),
            ("name", item.name.as_deref(), false),
        ] {
            let Some(formula) = formula else { continue };
            let sheet_name = chart_formula_sheet_name(formula).ok_or_else(|| Error::Reverse(format!(
                "chart part `{part}` Series.{property} `{formula}` is not a sheet-qualified A1 reference; repair it to a cell or range such as `Sheet1!$A$1:$A$5`"
            )))?;
            if !chart_formula_has_valid_reference(formula, require_range) {
                return Err(Error::Reverse(format!(
                    "chart part `{part}` Series.{property} `{formula}` is not a supported A1 reference; repair it to a cell or range"
                )));
            }
            if !worksheet_names.contains(&sheet_name.to_lowercase()) {
                return Err(Error::Reverse(format!(
                    "chart part `{part}` Series.{property} references missing worksheet `{sheet_name}`; add the referenced worksheet or repair the chart formula"
                )));
            }
        }
        if chart_type.starts_with("scatter") {
            let category_range = item
                .categories
                .as_deref()
                .and_then(|formula| formula.split_once('!'))
                .and_then(|(_, range)| parse_a1_range(range));
            let value_range = item
                .values
                .split_once('!')
                .and_then(|(_, range)| parse_a1_range(range));
            let ranges_match = value_range.zip(category_range).is_some_and(
                |((vr0, vc0, vr1, vc1), (cr0, cc0, cr1, cc1))| {
                    (vr0 == vr1 || vc0 == vc1)
                        && (cr0 == cr1 || cc0 == cc1)
                        && vr1 - vr0 == cr1 - cr0
                        && vc1 - vc0 == cc1 - cc0
                },
            );
            if !ranges_match {
                return Err(Error::Reverse(format!(
                    "chart part `{part}` scatter series must use one-dimensional X/Y ranges with equal point counts; repair categories/values or use another chart type"
                )));
            }
        }
    }
    Ok(WorksheetChart {
        chart_type,
        cell: cell_reference(from_row, from_column),
        width,
        height,
        name: anchor.name,
        title: chart_title,
        x_axis_title,
        y_axis_title,
        show_legend,
        legend_position: legend_position.filter(|position| *position != "right"),
        hole_size: hole_size.filter(|size| *size != 50),
        first_slice_angle: first_slice_angle.filter(|angle| *angle != 0),
        alt_text: anchor.alt_text,
        series,
    })
}

fn parse_drawing_image(
    bytes: &[u8],
    part: &str,
    anchor: ChartAnchor,
    metadata: &WorksheetMetadata,
) -> xlsx_jsx::Result<WorksheetImage> {
    let movement = match anchor.edit_as.as_deref() {
        None => Some("moveAndSize"),
        Some("oneCell") => None,
        Some("absolute") => Some("fixed"),
        Some(value) => {
            return Err(Error::Reverse(format!(
                "picture in `{part}` uses unsupported editAs `{value}`; use moveAndSize, moveOnly, or fixed movement"
            )));
        }
    };
    let extension = part
        .rsplit_once('.')
        .map(|(_, extension)| extension.to_ascii_lowercase())
        .ok_or_else(|| Error::Reverse(format!("image part `{part}` has no file extension")))?;
    let media_type = match extension.as_str() {
        "png" => "png",
        "jpeg" | "jpg" => "jpeg",
        "gif" => "gif",
        "bmp" => "bmp",
        _ => {
            return Err(Error::Reverse(format!(
                "image part `{part}` uses unsupported media type; use PNG, JPEG, GIF, or BMP"
            )));
        }
    };
    XlsxImage::new_from_buffer(bytes).map_err(|error| {
        Error::Reverse(format!(
            "image part `{part}` is not a valid supported raster image ({error}); repair the embedded media bytes"
        ))
    })?;
    let media_matches_bytes = match media_type {
        "png" => bytes.starts_with(b"\x89PNG\r\n\x1a\n"),
        "jpeg" => bytes.starts_with(&[0xff, 0xd8, 0xff]),
        "gif" => bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a"),
        "bmp" => bytes.starts_with(b"BM"),
        _ => false,
    };
    if !media_matches_bytes {
        return Err(Error::Reverse(format!(
            "image part `{part}` extension does not match its raster bytes; repair the package content type or embedded media"
        )));
    }
    let (cell, width, height, offset_x, offset_y) = validate_image_anchor(&anchor, metadata, part)?;
    let object_id = anchor.object_id.ok_or_else(|| {
        Error::Reverse(format!(
            "picture in `{part}` is missing cNvPr@id; repair picture metadata"
        ))
    })?;
    let expected_name = format!("Picture {}", object_id.saturating_sub(1));
    if object_id == 0 || anchor.name != expected_name {
        return Err(Error::Reverse(format!(
            "picture in `{part}` has custom object name `{}` that the pinned writer cannot reproduce; rename it to `{expected_name}` or extend Image naming support",
            anchor.name
        )));
    }
    let data_uri = format!("data:image/{media_type};base64,{}", BASE64.encode(bytes));
    Ok(WorksheetImage {
        cell,
        width,
        height,
        data_uri,
        offset_x,
        offset_y,
        movement,
        alt_text: anchor.alt_text.filter(|text| !text.is_empty()),
        decorative: anchor.decorative,
    })
}

#[expect(
    clippy::too_many_lines,
    reason = "Image anchor geometry must be checked as a single lossless OOXML invariant."
)]
fn validate_image_anchor(
    anchor: &ChartAnchor,
    metadata: &WorksheetMetadata,
    part: &str,
) -> xlsx_jsx::Result<(String, u32, u32, u32, u32)> {
    const PX_EMU: u64 = 9_525;
    const COLUMN_PIXELS: u32 = 64;
    const ROW_PIXELS: u32 = 20;
    let from_row = anchor.from_row.ok_or_else(|| {
        Error::Reverse(format!(
            "picture in `{part}` is missing from/row; repair its anchor"
        ))
    })?;
    let from_column = anchor.from_column.ok_or_else(|| {
        Error::Reverse(format!(
            "picture in `{part}` is missing from/col; repair its anchor"
        ))
    })?;
    let to_row = anchor.to_row.ok_or_else(|| {
        Error::Reverse(format!(
            "picture in `{part}` is missing to/row; repair its anchor"
        ))
    })?;
    let to_column = anchor.to_column.ok_or_else(|| {
        Error::Reverse(format!(
            "picture in `{part}` is missing to/col; repair its anchor"
        ))
    })?;
    if !anchor.from_column_offset_seen || !anchor.from_row_offset_seen {
        return Err(Error::Reverse(format!(
            "picture in `{part}` is missing from/colOff or from/rowOff; repair its anchor"
        )));
    }
    if to_row < from_row || to_column < from_column {
        return Err(Error::Reverse(format!(
            "picture in `{part}` has reversed anchor coordinates; repair its from/to positions"
        )));
    }
    if to_row >= 1_048_576 || to_column >= 16_384 {
        return Err(Error::Reverse(format!(
            "picture in `{part}` extends beyond Excel's worksheet grid; move or resize it"
        )));
    }
    let from_column_offset = u64::from(anchor.from_column_offset.ok_or_else(|| {
        Error::Reverse(format!(
            "picture in `{part}` is missing from/colOff; repair its anchor"
        ))
    })?);
    let from_row_offset = u64::from(anchor.from_row_offset.ok_or_else(|| {
        Error::Reverse(format!(
            "picture in `{part}` is missing from/rowOff; repair its anchor"
        ))
    })?);
    if from_column_offset % PX_EMU != 0
        || from_row_offset % PX_EMU != 0
        || from_column_offset / PX_EMU >= u64::from(COLUMN_PIXELS)
        || from_row_offset / PX_EMU >= u64::from(ROW_PIXELS)
    {
        return Err(Error::Reverse(format!(
            "picture in `{part}` has fractional-pixel or out-of-cell start offsets; use whole-pixel offsets within a 64×20-pixel cell"
        )));
    }
    let to_column_offset = u64::from(anchor.to_column_offset.ok_or_else(|| {
        Error::Reverse(format!(
            "picture in `{part}` is missing to/colOff; repair its anchor"
        ))
    })?);
    let to_row_offset = u64::from(anchor.to_row_offset.ok_or_else(|| {
        Error::Reverse(format!(
            "picture in `{part}` is missing to/rowOff; repair its anchor"
        ))
    })?);
    if to_column_offset % PX_EMU != 0 || to_row_offset % PX_EMU != 0 {
        return Err(Error::Reverse(format!(
            "picture in `{part}` ends at fractional-pixel offsets; use whole-pixel image dimensions"
        )));
    }
    if to_column_offset / PX_EMU >= u64::from(COLUMN_PIXELS)
        || to_row_offset / PX_EMU >= u64::from(ROW_PIXELS)
    {
        return Err(Error::Reverse(format!(
            "picture in `{part}` has an unnormalized end offset; use a standard two-cell anchor"
        )));
    }
    let width = u64::from(to_column - from_column)
        .checked_mul(u64::from(COLUMN_PIXELS))
        .and_then(|pixels| pixels.checked_add(to_column_offset / PX_EMU))
        .and_then(|pixels| pixels.checked_sub(from_column_offset / PX_EMU))
        .ok_or_else(|| Error::Reverse(format!("picture width in `{part}` overflows pixels")))?;
    let height = u64::from(to_row - from_row)
        .checked_mul(u64::from(ROW_PIXELS))
        .and_then(|pixels| pixels.checked_add(to_row_offset / PX_EMU))
        .and_then(|pixels| pixels.checked_sub(from_row_offset / PX_EMU))
        .ok_or_else(|| Error::Reverse(format!("picture height in `{part}` overflows pixels")))?;
    if !(1..=1790).contains(&width) || !(1..=1790).contains(&height) {
        return Err(Error::Reverse(format!(
            "picture in `{part}` measures {width}×{height} pixels outside the supported 1–1790 range; resize it or extend Image dimensions"
        )));
    }
    if metadata
        .default_row_height_pixels
        .is_some_and(|value| value != ROW_PIXELS)
        || metadata.rows.iter().any(|(row, properties)| {
            *row >= from_row
                && *row <= to_row
                && (properties
                    .height
                    .is_some_and(|value| (value - 15.0).abs() > f64::EPSILON)
                    || properties.hidden)
        })
    {
        return Err(Error::Reverse(format!(
            "picture in `{part}` crosses custom row heights, so its pixel size is ambiguous; use default row heights or extend image geometry support"
        )));
    }
    if metadata
        .columns
        .range(
            u16::try_from(from_column).expect("validated Excel column bounds")
                ..=u16::try_from(to_column).expect("validated Excel column bounds"),
        )
        .any(|(_, properties)| {
            properties
                .width_pixels
                .is_some_and(|value| value != COLUMN_PIXELS)
                || properties.hidden
        })
    {
        return Err(Error::Reverse(format!(
            "picture in `{part}` crosses custom column widths, so its pixel size is ambiguous; use default column widths or extend image geometry support"
        )));
    }
    let expected_x =
        (u64::from(from_column) * u64::from(COLUMN_PIXELS) + from_column_offset / PX_EMU) * PX_EMU;
    let expected_y =
        (u64::from(from_row) * u64::from(ROW_PIXELS) + from_row_offset / PX_EMU) * PX_EMU;
    if anchor.absolute_x != Some(expected_x)
        || anchor.absolute_y != Some(expected_y)
        || anchor.transform_width != Some(width * PX_EMU)
        || anchor.transform_height != Some(height * PX_EMU)
    {
        return Err(Error::Reverse(format!(
            "picture in `{part}` has transform geometry that disagrees with its anchor; use an unrotated, uncropped rectangle aligned to whole-pixel cell geometry"
        )));
    }
    let width = u32::try_from(width).map_err(|error| {
        Error::Reverse(format!(
            "picture width in `{part}` exceeds 32-bit pixels: {error}"
        ))
    })?;
    let height = u32::try_from(height).map_err(|error| {
        Error::Reverse(format!(
            "picture height in `{part}` exceeds 32-bit pixels: {error}"
        ))
    })?;
    let offset_x = u32::try_from(from_column_offset / PX_EMU).map_err(|error| {
        Error::Reverse(format!(
            "picture x offset in `{part}` exceeds pixels: {error}"
        ))
    })?;
    let offset_y = u32::try_from(from_row_offset / PX_EMU).map_err(|error| {
        Error::Reverse(format!(
            "picture y offset in `{part}` exceeds pixels: {error}"
        ))
    })?;
    Ok((
        cell_reference(from_row, from_column),
        width,
        height,
        offset_x,
        offset_y,
    ))
}

fn chart_formula_sheet_name(formula: &str) -> Option<String> {
    if formula.starts_with('=') || formula.matches('!').count() != 1 {
        return None;
    }
    let (sheet, _) = formula.split_once('!')?;
    if let Some(quoted) = sheet
        .strip_prefix('\'')
        .and_then(|value| value.strip_suffix('\''))
    {
        if quoted.is_empty()
            || quoted.replace("''", "").contains('\'')
            || quoted.contains('[')
            || quoted.contains(']')
        {
            return None;
        }
        Some(quoted.replace("''", "'"))
    } else if !sheet.is_empty()
        && sheet
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || matches!(character, '_' | '.'))
    {
        Some(sheet.to_owned())
    } else {
        None
    }
}

fn chart_formula_has_valid_reference(formula: &str, require_range: bool) -> bool {
    let Some((_, reference)) = formula.split_once('!') else {
        return false;
    };
    if let Some((first, last)) = reference.split_once(':') {
        parse_a1_cell_reference(first).is_some()
            && parse_a1_cell_reference(last).is_some()
            && parse_a1_range(reference).is_some()
    } else {
        !require_range && parse_a1_cell_reference(reference).is_some()
    }
}

fn set_chart_type(
    current: Option<&'static str>,
    candidate: &'static str,
    part: &str,
) -> xlsx_jsx::Result<&'static str> {
    if current.is_some_and(|value| value != candidate) {
        return Err(Error::Reverse(format!(
            "chart part `{part}` combines chart types; split it into individual charts because combo charts are not represented by Chart"
        )));
    }
    Ok(candidate)
}

fn is_groupable_chart_type(chart_type: &str) -> bool {
    matches!(chart_type, "area" | "bar" | "column" | "line")
}

fn is_percent_stacked_chart_type(chart_type: &str) -> bool {
    matches!(
        chart_type,
        "areaPercentStacked" | "barPercentStacked" | "columnPercentStacked" | "linePercentStacked"
    )
}

fn chart_type_for_grouping(base: &'static str, grouping: &str) -> Option<&'static str> {
    match (base, grouping) {
        ("area", "standard") => Some("area"),
        ("area", "stacked") => Some("areaStacked"),
        ("area", "percentStacked") => Some("areaPercentStacked"),
        ("column", "clustered") => Some("column"),
        ("column", "stacked") => Some("columnStacked"),
        ("column", "percentStacked") => Some("columnPercentStacked"),
        ("bar", "clustered") => Some("bar"),
        ("bar", "stacked") => Some("barStacked"),
        ("bar", "percentStacked") => Some("barPercentStacked"),
        ("line", "standard") => Some("line"),
        ("line", "stacked") => Some("lineStacked"),
        ("line", "percentStacked") => Some("linePercentStacked"),
        _ => None,
    }
}

fn resolve_radar_chart_type(
    chart_type: &'static str,
    radar_style: Option<&'static str>,
    series: &[WorksheetChartSeries],
    part: &str,
) -> xlsx_jsx::Result<&'static str> {
    if chart_type != "radar" {
        return Ok(chart_type);
    }
    let marker_count = series.iter().filter(|item| item.marker_seen).count();
    if marker_count != 0 && marker_count != series.len() {
        return Err(Error::Reverse(format!(
            "chart part `{part}` has mixed radar marker overrides; use the same marker behavior for every series"
        )));
    }
    let disabled_marker_count = series.iter().filter(|item| item.marker_none).count();
    if disabled_marker_count != 0 && disabled_marker_count != marker_count && marker_count != 0 {
        return Err(Error::Reverse(format!(
            "chart part `{part}` has mixed radar marker visibility; use the same marker behavior for every series"
        )));
    }
    let markers_disabled = (marker_count > 0 && series.iter().all(|item| item.marker_none))
        || radar_style == Some("standard");
    let markers_enabled = (marker_count > 0 && series.iter().all(|item| !item.marker_none))
        || (marker_count == 0 && radar_style == Some("marker"));
    let markers_overridden = marker_count > 0;
    match radar_style {
        Some("filled") if !markers_overridden => Ok("radarFilled"),
        Some("filled") => Err(Error::Reverse(format!(
            "chart part `{part}` combines filled radar with per-series markers; remove marker overrides or use a standard radar chart"
        ))),
        Some("standard") if !markers_enabled => Ok("radar"),
        Some("standard") => Err(Error::Reverse(format!(
            "chart part `{part}` radarStyle=standard conflicts with enabled series markers; remove marker overrides or use radarStyle=marker"
        ))),
        Some("marker") | None if markers_disabled => Ok("radar"),
        Some("marker") | None if markers_enabled => Ok("radarWithMarkers"),
        None => Ok("radar"),
        Some(_) => Err(Error::Reverse(format!(
            "chart part `{part}` has an unsupported radar style; use standard, marker, or filled"
        ))),
    }
}

fn resolve_scatter_chart_type(
    chart_type: &'static str,
    scatter_style: Option<&'static str>,
    series: &[WorksheetChartSeries],
    part: &str,
) -> xlsx_jsx::Result<&'static str> {
    if chart_type != "scatter" {
        return Ok(chart_type);
    }
    if scatter_style != Some("lineMarker") {
        return Err(Error::Reverse(format!(
            "chart part `{part}` has unsupported scatter style; use straight lines, or add matching scatter style support"
        )));
    }
    let marker_overrides = series.iter().filter(|item| item.marker_seen).count();
    if marker_overrides != 0 && marker_overrides != series.len() {
        return Err(Error::Reverse(format!(
            "chart part `{part}` has mixed scatter marker overrides; use the same marker behavior for every series"
        )));
    }
    if series
        .iter()
        .any(|item| item.marker_seen && !item.marker_none)
    {
        return Err(Error::Reverse(format!(
            "chart part `{part}` has unsupported visible scatter marker overrides; remove per-series markers or add marker style support"
        )));
    }
    if marker_overrides == series.len() {
        Ok("scatterWithLines")
    } else {
        Ok("scatterWithLinesAndMarkers")
    }
}

fn audit_chart_attributes(
    event: &BytesStart<'_>,
    tag: &str,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    chart_type: Option<&'static str>,
) -> xlsx_jsx::Result<()> {
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.as_ref().starts_with(b"xmlns") {
            continue;
        }
        let key = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?
            .into_owned();
        let accepted = match (tag, key.as_str()) {
            ("lang", "val") | ("endParaRPr" | "rPr", "lang") => value == "en-US",
            ("barDir", "val") => matches!(value.as_str(), "col" | "bar"),
            ("radarStyle", "val") => {
                chart_type == Some("radar")
                    && matches!(value.as_str(), "standard" | "marker" | "filled")
            }
            ("scatterStyle", "val") => chart_type == Some("scatter") && value == "lineMarker",
            ("grouping", "val") => match chart_type {
                Some("area") => {
                    matches!(value.as_str(), "standard" | "stacked" | "percentStacked")
                }
                Some("line") => {
                    matches!(value.as_str(), "standard" | "stacked" | "percentStacked")
                }
                Some("bar" | "column") => {
                    matches!(value.as_str(), "clustered" | "stacked" | "percentStacked")
                }
                _ => false,
            },
            ("overlap", "val") => {
                matches!(
                    chart_type,
                    Some(
                        "columnStacked"
                            | "columnPercentStacked"
                            | "barStacked"
                            | "barPercentStacked"
                    )
                ) && value == "100"
            }
            ("idx" | "order" | "ptCount" | "axId" | "crossAx", "val") | ("pt", "idx") => {
                value.parse::<u32>().is_ok()
            }
            ("orientation", "val") => value == "minMax",
            ("axPos", "val") => matches!(value.as_str(), "b" | "l"),
            ("tickLblPos", "val") => value == "nextTo",
            ("majorTickMark", "val") => value == "cross",
            ("crosses", "val") => value == "autoZero",
            ("auto" | "plotVisOnly" | "marker" | "varyColors", "val")
            | ("numFmt", "sourceLinked") => value == "1",
            ("symbol", "val") => value == "none",
            ("lblAlgn", "val") => value == "ctr",
            ("lblOffset", "val") => value == "100",
            ("crossBetween", "val") => match chart_type {
                Some("area" | "areaStacked" | "areaPercentStacked" | "scatter") => {
                    value == "midCat"
                }
                _ => value == "between",
            },
            ("numFmt", "formatCode") => {
                value == "General"
                    || (chart_type.is_some_and(is_percent_stacked_chart_type) && value == "0%")
            }
            ("legendPos", "val") => matches!(value.as_str(), "r" | "l" | "t" | "b" | "tr"),
            ("holeSize", "val") => {
                chart_type == Some("doughnut")
                    && value
                        .parse::<u8>()
                        .is_ok_and(|size| (10..=90).contains(&size))
            }
            ("firstSliceAng", "val") => {
                matches!(chart_type, Some("pie" | "doughnut"))
                    && value.parse::<u16>().is_ok_and(|angle| angle <= 360)
            }
            ("pPr", "rtl") => value == "0",
            ("pageMargins", "b" | "t") => value == "0.75",
            ("pageMargins", "l" | "r") => value == "0.7",
            ("pageMargins", "header" | "footer") => value == "0.3",
            _ => false,
        };
        if !accepted {
            return Err(Error::Reverse(format!(
                "chart part `{part}` has unsupported `{tag}@{key}={value}`; remove custom chart/axis/print formatting or add matching DSL and reverse support"
            )));
        }
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Drawing XML parsing pairs each supported chart/picture object with its anchor and relationship."
)]
fn parse_chart_drawing(xml: &str, part: &str) -> xlsx_jsx::Result<Vec<ChartAnchor>> {
    let mut reader = XmlReader::from_str(xml);
    let mut anchors = Vec::new();
    let mut object_ids = std::collections::HashSet::new();
    let mut current: Option<ChartAnchor> = None;
    let mut location = None;
    let mut coordinate = None;
    let mut in_transform = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event)) => {
                let tag = event.local_name();
                let tag_name = String::from_utf8_lossy(tag.as_ref());
                audit_chart_drawing_element(&event, &tag_name, part, reader.decoder())?;
                match tag.as_ref() {
                    b"twoCellAnchor" => {
                        if current.replace(ChartAnchor::default()).is_some() {
                            return Err(Error::Reverse(format!(
                                "drawing `{part}` has nested anchors; repair drawing structure"
                            )));
                        }
                        if let Some(anchor) = current.as_mut() {
                            for attribute in event.attributes().with_checks(false) {
                                let attribute =
                                    attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                                if attribute.key.as_ref().starts_with(b"xmlns") {
                                    continue;
                                }
                                if attribute.key.local_name().as_ref() == b"editAs" {
                                    anchor.edit_as = Some(
                                        attribute
                                            .decoded_and_normalized_value(
                                                XmlVersion::Implicit1_0,
                                                reader.decoder(),
                                            )
                                            .map_err(|error| Error::Reverse(error.to_string()))?
                                            .into_owned(),
                                    );
                                }
                            }
                        }
                    }
                    b"graphicFrame" => {
                        if let Some(anchor) = current.as_mut() {
                            anchor.kind = DrawingObjectKind::Chart;
                        }
                    }
                    b"pic" => {
                        if let Some(anchor) = current.as_mut() {
                            anchor.kind = DrawingObjectKind::Image;
                        }
                    }
                    b"oneCellAnchor" | b"absoluteAnchor" => {
                        return Err(Error::Reverse(format!(
                            "drawing `{part}` uses a non-two-cell chart anchor; use a two-cell anchor with integer-pixel geometry or extend chart positioning support"
                        )));
                    }
                    b"from" => location = Some("from"),
                    b"to" => location = Some("to"),
                    b"col" | b"row" | b"colOff" | b"rowOff" => {
                        coordinate = Some(String::from_utf8_lossy(tag.as_ref()).into_owned());
                    }
                    b"chart" => {
                        if let Some(anchor) = current.as_mut() {
                            for attribute in event.attributes().with_checks(false) {
                                let attribute =
                                    attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                                if attribute.key.as_ref().starts_with(b"xmlns") {
                                    continue;
                                }
                                if attribute.key.local_name().as_ref() == b"id" {
                                    anchor.relationship_id = attribute
                                        .decoded_and_normalized_value(
                                            XmlVersion::Implicit1_0,
                                            reader.decoder(),
                                        )
                                        .map_err(|error| Error::Reverse(error.to_string()))?
                                        .into_owned();
                                } else {
                                    return Err(Error::Reverse(format!(
                                        "chart frame in `{part}` has unsupported attribute `{}`",
                                        String::from_utf8_lossy(attribute.key.as_ref())
                                    )));
                                }
                            }
                        }
                    }
                    b"blip" => {
                        if let Some(anchor) = current.as_mut() {
                            for attribute in event.attributes().with_checks(false) {
                                let attribute =
                                    attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                                if attribute.key.as_ref().starts_with(b"xmlns") {
                                    continue;
                                }
                                if attribute.key.local_name().as_ref() == b"embed" {
                                    anchor.relationship_id = attribute
                                        .decoded_and_normalized_value(
                                            XmlVersion::Implicit1_0,
                                            reader.decoder(),
                                        )
                                        .map_err(|error| Error::Reverse(error.to_string()))?
                                        .into_owned();
                                } else {
                                    return Err(Error::Reverse(format!(
                                        "picture in `{part}` has unsupported blip metadata; remove links or unsupported embedded-image effects"
                                    )));
                                }
                            }
                        }
                    }
                    b"cNvPr" => {
                        if let Some(anchor) = current.as_mut() {
                            for attribute in event.attributes().with_checks(false) {
                                let attribute =
                                    attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                                if attribute.key.as_ref().starts_with(b"xmlns") {
                                    continue;
                                }
                                let value = attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?
                                    .into_owned();
                                match attribute.key.local_name().as_ref() {
                                    b"id" if anchor.object_id.is_none() => {
                                        anchor.object_id = value.parse().ok();
                                    }
                                    b"name" if anchor.name.is_empty() => anchor.name = value,
                                    b"descr" if anchor.alt_text.is_none() => {
                                        anchor.alt_text = Some(value);
                                    }
                                    _ => {
                                        return Err(Error::Reverse(format!(
                                            "drawing `{part}` has duplicate or unsupported chart object metadata; preserve one name and optional alt text"
                                        )));
                                    }
                                }
                            }
                        }
                    }
                    b"xfrm" => in_transform = true,
                    b"off" | b"ext" if in_transform => {
                        if let Some(anchor) = current.as_mut() {
                            let mut x = None;
                            let mut y = None;
                            let mut cx = None;
                            let mut cy = None;
                            for attribute in event.attributes().with_checks(false) {
                                let attribute =
                                    attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                                if attribute.key.as_ref().starts_with(b"xmlns") {
                                    continue;
                                }
                                let value = attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                                    .map_err(|error| Error::Reverse(error.to_string()))?.parse::<u64>().map_err(|_| Error::Reverse(format!("drawing `{part}` contains a non-integer picture transform; use whole-pixel geometry")))?;
                                match (tag.as_ref(), attribute.key.local_name().as_ref()) {
                                    (b"off", b"x") => x = Some(value),
                                    (b"off", b"y") => y = Some(value),
                                    (b"ext", b"cx") => cx = Some(value),
                                    (b"ext", b"cy") => cy = Some(value),
                                    _ => {
                                        return Err(Error::Reverse(format!(
                                            "drawing `{part}` has unsupported transform attributes; retain only x/y and cx/cy"
                                        )));
                                    }
                                }
                            }
                            if tag.as_ref() == b"off" {
                                anchor.absolute_x = x;
                                anchor.absolute_y = y;
                            } else {
                                anchor.transform_width = cx;
                                anchor.transform_height = cy;
                            }
                        }
                    }
                    b"decorative" => {
                        if let Some(anchor) = current.as_mut() {
                            anchor.decorative =
                                event.attributes().with_checks(false).any(|attribute| {
                                    attribute.is_ok_and(|attribute| {
                                        attribute.key.local_name().as_ref() == b"val"
                                            && attribute.value.as_ref() == b"1"
                                    })
                                });
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(text)) => {
                let (Some(anchor), Some(field)) = (current.as_mut(), coordinate.take()) else {
                    continue;
                };
                let raw = text
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                let value = raw.parse::<u32>().map_err(|_| Error::Reverse(format!("drawing `{part}` has invalid anchor coordinate `{raw}`; repair col/row values to nonnegative integers")))?;
                match (location, field.as_str()) {
                    (Some("from"), "row") => anchor.from_row = Some(value),
                    (Some("from"), "col") => anchor.from_column = Some(value),
                    (Some("to"), "row") => anchor.to_row = Some(value),
                    (Some("to"), "col") => anchor.to_column = Some(value),
                    (Some("to"), "colOff") => anchor.to_column_offset = Some(value),
                    (Some("to"), "rowOff") => anchor.to_row_offset = Some(value),
                    (Some("from"), "colOff") if value % 9_525 == 0 => {
                        anchor.from_column_offset_seen = true;
                        anchor.from_column_offset = Some(value);
                    }
                    (Some("from"), "rowOff") if value % 9_525 == 0 => {
                        anchor.from_row_offset_seen = true;
                        anchor.from_row_offset = Some(value);
                    }
                    _ => {
                        return Err(Error::Reverse(format!(
                            "drawing `{part}` has unsupported start-anchor offsets or coordinates; use zero start offsets and standard end offsets"
                        )));
                    }
                }
            }
            Ok(Event::End(event)) => match event.local_name().as_ref() {
                b"from" | b"to" => location = None,
                b"xfrm" => in_transform = false,
                b"twoCellAnchor" => {
                    let anchor = current.take().ok_or_else(|| {
                        Error::Reverse(format!(
                            "drawing `{part}` closes an anchor that was not opened"
                        ))
                    })?;
                    if anchor.relationship_id.is_empty() {
                        return Err(Error::Reverse(format!(
                            "drawing `{part}` contains an object without a chart/image relationship; remove unsupported shapes or add matching DSL support"
                        )));
                    }
                    if anchor.name.is_empty() {
                        return Err(Error::Reverse(format!(
                            "drawing `{part}` chart is missing cNvPr@name; repair its object metadata"
                        )));
                    }
                    let object_id = anchor.object_id.ok_or_else(|| {
                        Error::Reverse(format!(
                            "drawing `{part}` object is missing cNvPr@id; repair its non-visual properties"
                        ))
                    })?;
                    if object_id == 0 || !object_ids.insert(object_id) {
                        return Err(Error::Reverse(format!(
                            "drawing `{part}` repeats or has invalid cNvPr id `{object_id}`; assign each object a unique positive ID"
                        )));
                    }
                    if anchor.kind == DrawingObjectKind::Unknown {
                        return Err(Error::Reverse(format!(
                            "drawing `{part}` contains an unrecognized object; use a supported chart or picture"
                        )));
                    }
                    anchors.push(anchor);
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid drawing XML in `{part}`: {error}"
                )));
            }
        }
    }
    if current.is_some() {
        return Err(Error::Reverse(format!(
            "drawing `{part}` has an unterminated twoCellAnchor"
        )));
    }
    Ok(anchors)
}

fn audit_chart_drawing_element(
    event: &BytesStart<'_>,
    tag: &str,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    const ALLOWED: &[&str] = &[
        "wsDr",
        "twoCellAnchor",
        "from",
        "to",
        "col",
        "row",
        "colOff",
        "rowOff",
        "graphicFrame",
        "nvGraphicFramePr",
        "cNvPr",
        "cNvGraphicFramePr",
        "xfrm",
        "off",
        "ext",
        "graphic",
        "graphicData",
        "chart",
        "clientData",
        "pic",
        "nvPicPr",
        "cNvPicPr",
        "picLocks",
        "blipFill",
        "blip",
        "stretch",
        "fillRect",
        "spPr",
        "prstGeom",
        "avLst",
        "extLst",
        "creationId",
        "decorative",
    ];
    if !ALLOWED.contains(&tag) {
        return Err(Error::Reverse(format!(
            "drawing `{part}` contains unsupported `<{tag}>`; only charts and raster pictures with supported two-cell anchors are represented by this DSL"
        )));
    }
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.as_ref().starts_with(b"xmlns") {
            continue;
        }
        let key = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?
            .into_owned();
        let accepted = match (tag, key.as_str()) {
            ("graphicFrame", "macro") => value.is_empty(),
            ("cNvPr", "id") => value.parse::<u32>().is_ok(),
            ("cNvPr", "name" | "descr") => !value.is_empty() && value.chars().count() <= 255,
            ("twoCellAnchor", "editAs") => matches!(value.as_str(), "oneCell" | "absolute"),
            ("off", "x" | "y") | ("ext", "cx" | "cy") => value.parse::<u64>().is_ok(),
            ("picLocks", "noChangeAspect") => value == "1",
            ("prstGeom", "prst") => value == "rect",
            ("decorative", "val") => matches!(value.as_str(), "0" | "1"),
            ("ext", "uri") => matches!(
                value.as_str(),
                "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" | "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
            ),
            ("creationId", "id") => value == "{00000000-0008-0000-0000-000002000000}",
            ("blip", "embed") | ("chart", "id") => !value.is_empty(),
            ("graphicData", "uri") => {
                value == "http://schemas.openxmlformats.org/drawingml/2006/chart"
            }
            _ => false,
        };
        if !accepted {
            return Err(Error::Reverse(format!(
                "drawing `{part}` has unsupported `{tag}@{key}={value}` metadata; remove custom drawing names/geometry or add matching Chart properties"
            )));
        }
    }
    Ok(())
}

#[derive(Default)]
struct WorksheetRelationships {
    hyperlinks: HashMap<String, String>,
    tables: HashMap<String, String>,
    comments: HashMap<String, String>,
    vml_drawings: HashMap<String, String>,
    drawings: HashMap<String, String>,
}

fn parse_worksheet_relationships(
    xml: &str,
    part: &str,
) -> xlsx_jsx::Result<WorksheetRelationships> {
    let mut reader = XmlReader::from_str(xml);
    let mut relationships = WorksheetRelationships::default();
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event))
                if event.local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut relationship_type = None;
                let mut target = None;
                let mut target_mode = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let key = attribute.key.local_name();
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    match key.as_ref() {
                        b"Id" => id = Some(value),
                        b"Type" => relationship_type = Some(value),
                        b"Target" => target = Some(value),
                        b"TargetMode" => target_mode = Some(value),
                        _ => {
                            return Err(Error::Reverse(format!(
                                "worksheet relationship part `{part}` contains unsupported metadata `{}`; remove it or add explicit support",
                                String::from_utf8_lossy(key.as_ref())
                            )));
                        }
                    }
                }
                let id = id.ok_or_else(|| {
                    Error::Reverse(format!(
                        "worksheet relationship part `{part}` has a relationship without Id"
                    ))
                })?;
                let target = target.ok_or_else(|| {
                    Error::Reverse(format!(
                        "worksheet hyperlink relationship `{id}` in `{part}` is missing Target"
                    ))
                })?;
                let relationship_type = relationship_type.as_deref().unwrap_or_default();
                let relationship_map = if relationship_type.ends_with("/hyperlink")
                    && target_mode.as_deref() == Some("External")
                {
                    &mut relationships.hyperlinks
                } else if relationship_type.ends_with("/table") && target_mode.is_none() {
                    &mut relationships.tables
                } else if relationship_type.ends_with("/comments") && target_mode.is_none() {
                    &mut relationships.comments
                } else if relationship_type.ends_with("/vmlDrawing") && target_mode.is_none() {
                    &mut relationships.vml_drawings
                } else if relationship_type.ends_with("/drawing") && target_mode.is_none() {
                    &mut relationships.drawings
                } else {
                    return Err(Error::Reverse(format!(
                        "worksheet relationship `{id}` in `{part}` has unsupported type `{relationship_type}`; supported types are external hyperlinks, tables, legacy comments, VML note drawings, and chart drawings"
                    )));
                };
                if relationship_map.insert(id.clone(), target).is_some() {
                    return Err(Error::Reverse(format!(
                        "worksheet relationship part `{part}` repeats Id `{id}`; repair the relationship IDs"
                    )));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid worksheet relationships in `{part}`: {error}"
                )));
            }
        }
    }
    Ok(relationships)
}

fn resolve_note_target(
    target: &str,
    worksheet_part: &str,
    comments: bool,
) -> xlsx_jsx::Result<String> {
    let (prefix, suffix) = if comments {
        ("../comments", ".xml")
    } else {
        ("../drawings/vmlDrawing", ".vml")
    };
    let Some(index) = target
        .strip_prefix(prefix)
        .and_then(|value| value.strip_suffix(suffix))
        .filter(|value| !value.is_empty())
    else {
        return Err(Error::Reverse(format!(
            "worksheet `{worksheet_part}` note relationship target `{target}` is not a safe relative {kind} part; use the standard ../commentsN.xml or ../drawings/vmlDrawingN.vml target",
            kind = if comments { "comments" } else { "VML" }
        )));
    };
    if index.parse::<u32>().is_err() {
        return Err(Error::Reverse(format!(
            "worksheet `{worksheet_part}` note relationship target `{target}` has an invalid part number"
        )));
    }
    Ok(if comments {
        format!("xl/comments{index}.xml")
    } else {
        format!("xl/drawings/vmlDrawing{index}.vml")
    })
}

struct ParsedCommentRun {
    text: String,
    bold: bool,
    font_name: Option<String>,
    font_size: Option<f64>,
    font_family: Option<u8>,
    indexed_color: Option<u8>,
}

struct ParsedComment {
    cell: String,
    author_id: usize,
    runs: Vec<ParsedCommentRun>,
}

#[derive(Clone)]
struct ParsedVmlNote {
    cell: String,
    visible: bool,
    width: u32,
    height: u32,
    background_color: String,
    alt_text: String,
    margin_left_points: f64,
    margin_top_points: f64,
    z_index: u32,
    anchor: Vec<i32>,
}

struct VmlNoteBuilder {
    part: String,
    width_points: f64,
    height_points: f64,
    background_color: String,
    alt_text: String,
    visibility: String,
    margin_left_points: f64,
    margin_top_points: f64,
    z_index: u32,
    row: Option<u32>,
    column: Option<u16>,
    anchor: String,
    autofill: String,
    has_client_data: bool,
    auto_fill_count: usize,
    anchor_count: usize,
    row_count: usize,
    column_count: usize,
    visible_element: bool,
    move_with_cells: usize,
    size_with_cells: usize,
    has_fill: bool,
    shadow_count: usize,
    shape_path_count: usize,
    textbox_count: usize,
    div_count: usize,
}

#[expect(
    clippy::too_many_lines,
    reason = "Comment author, rich-text run, and note-value checks share one streaming XML state machine."
)]
fn parse_comments_xml(xml: &str, part: &str) -> xlsx_jsx::Result<Vec<WorksheetNote>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut authors = Vec::new();
    let mut comments = Vec::new();
    let mut comment: Option<ParsedComment> = None;
    let mut run: Option<ParsedCommentRun> = None;
    let mut in_author = false;
    let mut author_text = String::new();
    let mut in_text = false;
    let mut saw_root = false;
    loop {
        match reader.read_event() {
            Ok(event @ (Event::Start(_) | Event::Empty(_))) => {
                let is_empty = matches!(&event, Event::Empty(_));
                let (Event::Start(event) | Event::Empty(event)) = event else {
                    unreachable!("event matched start or empty")
                };
                let tag = String::from_utf8_lossy(event.local_name().as_ref()).into_owned();
                match tag.as_str() {
                    "comments" => {
                        let attributes = note_attributes(&event, &reader, part)?;
                        if attributes.len() != 1
                            || attributes.get("xmlns").map(String::as_str)
                                != Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                        {
                            return Err(note_loss_error(
                                part,
                                "comments root has unsupported namespace metadata",
                            ));
                        }
                    }
                    "authors" | "commentList" | "text" | "rPr" => {
                        if !note_attributes(&event, &reader, part)?.is_empty() {
                            return Err(note_loss_error(
                                part,
                                &format!("`<{tag}>` has unsupported attributes"),
                            ));
                        }
                    }
                    "author" => {
                        if !note_attributes(&event, &reader, part)?.is_empty() {
                            return Err(note_loss_error(
                                part,
                                "author element has unsupported attributes",
                            ));
                        }
                        in_author = true;
                        author_text.clear();
                    }
                    "comment" => {
                        if comment.is_some() {
                            return Err(Error::Reverse(format!(
                                "comments part `{part}` nests comment elements; repair commentList structure"
                            )));
                        }
                        let attributes = note_attributes(&event, &reader, part)?;
                        let cell = attributes.get("ref").cloned().ok_or_else(|| {
                            Error::Reverse(format!(
                                "comments part `{part}` has a comment without ref; repair its cell reference"
                            ))
                        })?;
                        let author_id = attributes
                            .get("authorId")
                            .and_then(|value| value.parse::<usize>().ok())
                            .ok_or_else(|| {
                                Error::Reverse(format!(
                                    "comment `{cell}` in `{part}` has an invalid authorId; repair the authors list"
                                ))
                            })?;
                        if attributes.len() != 2 {
                            return Err(Error::Reverse(format!(
                                "comment `{cell}` in `{part}` has unsupported attributes; remove extensions or add DSL support"
                            )));
                        }
                        comment = Some(ParsedComment {
                            cell,
                            author_id,
                            runs: Vec::new(),
                        });
                        saw_root = true;
                    }
                    "r" => {
                        if !note_attributes(&event, &reader, part)?.is_empty() {
                            return Err(note_loss_error(
                                part,
                                "rich-text run has unsupported attributes",
                            ));
                        }
                        if comment.is_none() || run.is_some() {
                            return Err(Error::Reverse(format!(
                                "comments part `{part}` has an invalid rich-text run nesting"
                            )));
                        }
                        run = Some(ParsedCommentRun {
                            text: String::new(),
                            bold: false,
                            font_name: None,
                            font_size: None,
                            font_family: None,
                            indexed_color: None,
                        });
                    }
                    "t" => {
                        let attributes = note_attributes(&event, &reader, part)?;
                        if attributes
                            .iter()
                            .any(|(key, value)| key != "xml:space" || value != "preserve")
                        {
                            return Err(note_loss_error(
                                part,
                                "comment text has unsupported whitespace metadata",
                            ));
                        }
                        in_text = true;
                    }
                    "b" | "sz" | "color" | "rFont" | "family" => {
                        let attributes = note_attributes(&event, &reader, part)?;
                        let current = run.as_mut().ok_or_else(|| {
                            Error::Reverse(format!(
                                "comments part `{part}` has run formatting outside a rich-text run"
                            ))
                        })?;
                        match tag.as_str() {
                            "b" => {
                                if attributes
                                    .get("val")
                                    .is_some_and(|value| value != "1" && value != "true")
                                {
                                    return Err(note_loss_error(
                                        part,
                                        "non-bold or complex bold formatting",
                                    ));
                                }
                                current.bold = true;
                            }
                            "sz" => {
                                current.font_size = attributes
                                    .get("val")
                                    .and_then(|value| value.parse::<f64>().ok());
                                if current.font_size.is_none() || attributes.len() != 1 {
                                    return Err(note_loss_error(part, "invalid note font size"));
                                }
                            }
                            "color" => {
                                current.indexed_color = attributes
                                    .get("indexed")
                                    .and_then(|value| value.parse::<u8>().ok());
                                if current.indexed_color != Some(81) || attributes.len() != 1 {
                                    return Err(note_loss_error(
                                        part,
                                        "nonstandard note text color",
                                    ));
                                }
                            }
                            "rFont" => {
                                current.font_name = attributes.get("val").cloned();
                                if current.font_name.is_none() || attributes.len() != 1 {
                                    return Err(note_loss_error(part, "invalid note font name"));
                                }
                            }
                            "family" => {
                                current.font_family = attributes
                                    .get("val")
                                    .and_then(|value| value.parse::<u8>().ok());
                                if current.font_family.is_none() || attributes.len() != 1 {
                                    return Err(note_loss_error(part, "invalid note font family"));
                                }
                            }
                            _ => unreachable!(),
                        }
                    }
                    _ => {
                        return Err(note_loss_error(
                            part,
                            &format!("unsupported comments XML element `<{tag}>`"),
                        ));
                    }
                }
                if is_empty {
                    match tag.as_str() {
                        "author" => {
                            in_author = false;
                            authors.push(std::mem::take(&mut author_text));
                        }
                        "r" => {
                            if let (Some(current), Some(parsed)) = (run.take(), comment.as_mut()) {
                                parsed.runs.push(current);
                            }
                        }
                        "comment" => {
                            if let Some(parsed) = comment.take() {
                                comments.push(parsed);
                            }
                        }
                        "t" => in_text = false,
                        _ => {}
                    }
                }
            }
            Ok(Event::Text(event)) => {
                let decoded = event
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                let decoded = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                if in_author {
                    author_text.push_str(&decoded);
                } else if in_text {
                    run.as_mut()
                        .ok_or_else(|| {
                            note_loss_error(part, "comment text appears outside a rich-text run")
                        })?
                        .text
                        .push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(note_loss_error(
                        part,
                        "unexpected text outside note content",
                    ));
                }
            }
            Ok(Event::End(event)) => {
                match event.local_name().as_ref() {
                    b"author" => {
                        in_author = false;
                        authors.push(std::mem::take(&mut author_text));
                    }
                    b"t" => in_text = false,
                    b"r" => {
                        let parsed = comment.as_mut().ok_or_else(|| {
                            note_loss_error(part, "rich-text run appears outside a comment")
                        })?;
                        parsed.runs.push(run.take().ok_or_else(|| {
                            note_loss_error(part, "rich-text run has no matching start")
                        })?);
                    }
                    b"comment" => {
                        comments.push(comment.take().ok_or_else(|| {
                            note_loss_error(part, "comment has no matching start")
                        })?);
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid comments XML in `{part}`: {error}"
                )));
            }
        }
    }
    if !saw_root {
        return Err(Error::Reverse(format!(
            "comments part `{part}` has no comments; remove its relationships or add support for empty note parts"
        )));
    }
    let mut notes = Vec::with_capacity(comments.len());
    let mut seen_cells = std::collections::HashSet::new();
    for parsed in comments {
        let (row, column) = parse_a1_cell_reference(&parsed.cell).ok_or_else(|| {
            note_loss_error(
                part,
                &format!("invalid comment reference `{}`", parsed.cell),
            )
        })?;
        let cell = format_a1_cell_reference(row, column);
        if !seen_cells.insert(cell.clone()) {
            return Err(note_loss_error(
                part,
                &format!("duplicate note for `{cell}`"),
            ));
        }
        let author = authors.get(parsed.author_id).cloned().ok_or_else(|| {
            note_loss_error(
                part,
                &format!("comment `{cell}` references a missing author"),
            )
        })?;
        let (text, author_prefix, font_name, font_size, font_family) =
            parse_note_text_runs(&parsed.runs, &author, part, &cell)?;
        if author.chars().count() > 52 {
            return Err(note_loss_error(
                part,
                &format!("note `{cell}` author exceeds the 52-character Excel limit"),
            ));
        }
        if text.chars().count() > 32_713 || font_name.is_empty() || font_name.chars().count() > 255
        {
            return Err(note_loss_error(
                part,
                &format!("note `{cell}` text or font name exceeds the supported XLSX limits"),
            ));
        }
        notes.push(WorksheetNote {
            cell,
            text,
            author,
            author_prefix,
            visible: false,
            width: 128,
            height: 74,
            background_color: "#FFFFE1".into(),
            font_name,
            font_size,
            font_family,
            alt_text: String::new(),
            vml_margin_left_points: None,
            vml_margin_top_points: None,
            vml_z_index: None,
            vml_anchor: None,
        });
    }
    Ok(notes)
}

fn parse_note_text_runs(
    runs: &[ParsedCommentRun],
    author: &str,
    part: &str,
    cell: &str,
) -> xlsx_jsx::Result<(String, bool, String, f64, u8)> {
    let (prefix, text_run) = match runs {
        [single] if !single.bold => (false, single),
        [author_run, text_run]
            if author_run.bold && author_run.text == format!("{author}:") && !text_run.bold =>
        {
            if author_run.font_name != text_run.font_name
                || author_run.font_size != text_run.font_size
                || author_run.font_family != text_run.font_family
                || author_run.indexed_color != text_run.indexed_color
            {
                return Err(note_loss_error(
                    part,
                    &format!("note `{cell}` author prefix and body use different font styles"),
                ));
            }
            let Some(text) = text_run.text.strip_prefix('\n') else {
                return Err(note_loss_error(
                    part,
                    &format!("note `{cell}` author prefix lacks its required newline"),
                ));
            };
            return note_run_style(text_run, text, true, part, cell);
        }
        _ => {
            return Err(note_loss_error(
                part,
                &format!(
                    "note `{cell}` uses rich text that cannot be represented by one Note.text value"
                ),
            ));
        }
    };
    note_run_style(text_run, &text_run.text, prefix, part, cell)
}

fn note_run_style(
    run: &ParsedCommentRun,
    text: &str,
    prefix: bool,
    part: &str,
    cell: &str,
) -> xlsx_jsx::Result<(String, bool, String, f64, u8)> {
    if run.indexed_color != Some(81) {
        return Err(note_loss_error(
            part,
            &format!("note `{cell}` has unsupported text color"),
        ));
    }
    let (Some(font_name), Some(font_size), Some(font_family)) =
        (&run.font_name, run.font_size, run.font_family)
    else {
        return Err(note_loss_error(
            part,
            &format!("note `{cell}` is missing font metadata"),
        ));
    };
    if !font_size.is_finite() || !(1.0..=409.0).contains(&font_size) {
        return Err(note_loss_error(
            part,
            &format!("note `{cell}` has invalid font size"),
        ));
    }
    if font_name.is_empty() || font_name.chars().count() > 255 {
        return Err(note_loss_error(
            part,
            &format!("note `{cell}` has an invalid font name"),
        ));
    }
    Ok((
        text.to_owned(),
        prefix,
        font_name.clone(),
        font_size,
        font_family,
    ))
}

fn note_loss_error(part: &str, reason: &str) -> Error {
    Error::Reverse(format!(
        "legacy note in `{part}` cannot be represented losslessly: {reason}; keep a single plain-text run with the standard author prefix and Note font, or extend Note and reverse support together"
    ))
}

fn note_attributes(
    event: &BytesStart<'_>,
    reader: &XmlReader<&[u8]>,
    part: &str,
) -> xlsx_jsx::Result<HashMap<String, String>> {
    let mut attributes = HashMap::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::Reverse(error.to_string()))?
            .into_owned();
        if attributes.insert(key.clone(), value).is_some() {
            return Err(note_loss_error(
                part,
                &format!("duplicate `{key}` attribute"),
            ));
        }
    }
    Ok(attributes)
}

#[expect(
    clippy::too_many_lines,
    reason = "VML note shapes require one streaming state machine to validate parent-child ordering and positional metadata."
)]
fn parse_vml_notes(xml: &str, part: &str) -> xlsx_jsx::Result<Vec<ParsedVmlNote>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut notes = Vec::new();
    let mut current: Option<VmlNoteBuilder> = None;
    let mut active_text: Option<String> = None;
    let mut saw_root = false;
    loop {
        match reader.read_event() {
            Ok(event @ (Event::Start(_) | Event::Empty(_))) => {
                let is_empty = matches!(&event, Event::Empty(_));
                let (Event::Start(event) | Event::Empty(event)) = event else {
                    unreachable!("event matched start or empty")
                };
                let tag = String::from_utf8_lossy(event.local_name().as_ref()).into_owned();
                match tag.as_str() {
                    "xml" | "shapelayout" | "idmap" | "shapetype" | "stroke" | "path"
                    | "shadow" | "textbox" | "div" => {
                        let attributes = note_attributes(&event, &reader, part)?;
                        validate_vml_decorative_element(&tag, &attributes, part)?;
                        if let Some(shape) = current.as_mut() {
                            match tag.as_str() {
                                "shadow" => shape.shadow_count += 1,
                                "path" => shape.shape_path_count += 1,
                                "textbox" => shape.textbox_count += 1,
                                "div" => shape.div_count += 1,
                                _ => {}
                            }
                        }
                    }
                    "shape" => {
                        if current.is_some() {
                            return Err(note_loss_error(part, "VML note shapes are nested"));
                        }
                        let attributes = note_attributes(&event, &reader, part)?;
                        if attributes.get("type").map(String::as_str) != Some("#_x0000_t202") {
                            return Err(note_loss_error(
                                part,
                                "VML drawing contains a non-note shape; remove drawings/buttons or add their DSL components",
                            ));
                        }
                        let style = attributes.get("style").ok_or_else(|| {
                            note_loss_error(part, "VML note shape is missing its style")
                        })?;
                        let style = parse_vml_style(style, part)?;
                        let width_points = parse_vml_points(style.get("width"), part, "width")?;
                        let height_points = parse_vml_points(style.get("height"), part, "height")?;
                        let visibility = style.get("visibility").cloned().ok_or_else(|| {
                            note_loss_error(part, "VML note shape is missing visibility")
                        })?;
                        let margin_left_points = parse_vml_nonnegative_points(
                            style.get("margin-left"),
                            part,
                            "margin-left",
                        )?;
                        let margin_top_points = parse_vml_nonnegative_points(
                            style.get("margin-top"),
                            part,
                            "margin-top",
                        )?;
                        let z_index = style
                            .get("z-index")
                            .and_then(|value| value.parse::<u32>().ok())
                            .filter(|value| *value > 0)
                            .ok_or_else(|| note_loss_error(part, "VML note z-index is invalid"))?;
                        if style.get("position").map(String::as_str) != Some("absolute")
                            || !style.contains_key("margin-left")
                            || !style.contains_key("margin-top")
                            || !style.contains_key("z-index")
                        {
                            return Err(note_loss_error(
                                part,
                                "VML note shape has unsupported positioning semantics",
                            ));
                        }
                        let background_color = attributes
                            .get("fillcolor")
                            .map(|value| canonical_rgb(value, part))
                            .transpose()?
                            .ok_or_else(|| {
                                note_loss_error(part, "VML note shape has no fillcolor")
                            })?;
                        let alt_text = attributes.get("alt").cloned().unwrap_or_default();
                        for key in attributes.keys() {
                            if !["id", "type", "style", "fillcolor", "alt", "o:insetmode"]
                                .contains(&key.as_str())
                            {
                                return Err(note_loss_error(
                                    part,
                                    &format!("unsupported VML shape attribute `{key}`"),
                                ));
                            }
                        }
                        if attributes
                            .get("o:insetmode")
                            .is_some_and(|value| value != "auto")
                        {
                            return Err(note_loss_error(part, "unsupported VML inset mode"));
                        }
                        if is_empty {
                            return Err(note_loss_error(part, "VML note shape has no ClientData"));
                        }
                        current = Some(VmlNoteBuilder {
                            part: part.to_owned(),
                            width_points,
                            height_points,
                            background_color,
                            alt_text,
                            visibility,
                            margin_left_points,
                            margin_top_points,
                            z_index,
                            row: None,
                            column: None,
                            anchor: String::new(),
                            autofill: String::new(),
                            has_client_data: false,
                            auto_fill_count: 0,
                            anchor_count: 0,
                            row_count: 0,
                            column_count: 0,
                            visible_element: false,
                            move_with_cells: 0,
                            size_with_cells: 0,
                            has_fill: false,
                            shadow_count: 0,
                            shape_path_count: 0,
                            textbox_count: 0,
                            div_count: 0,
                        });
                    }
                    "ClientData" => {
                        let shape = current.as_mut().ok_or_else(|| {
                            note_loss_error(part, "ClientData appears outside a VML note shape")
                        })?;
                        let attributes = note_attributes(&event, &reader, part)?;
                        if attributes.len() != 1
                            || attributes.get("ObjectType").map(String::as_str) != Some("Note")
                        {
                            return Err(note_loss_error(
                                part,
                                "VML ClientData is not a legacy note",
                            ));
                        }
                        if shape.has_client_data {
                            return Err(note_loss_error(
                                part,
                                "VML note has multiple ClientData elements",
                            ));
                        }
                        shape.has_client_data = true;
                    }
                    "MoveWithCells" => {
                        let shape = current.as_mut().ok_or_else(|| {
                            note_loss_error(part, "MoveWithCells appears outside a note shape")
                        })?;
                        shape.move_with_cells += 1;
                    }
                    "SizeWithCells" => {
                        let shape = current.as_mut().ok_or_else(|| {
                            note_loss_error(part, "SizeWithCells appears outside a note shape")
                        })?;
                        shape.size_with_cells += 1;
                    }
                    "Anchor" | "AutoFill" | "Row" | "Column" | "Visible" => {
                        let shape = current.as_mut().ok_or_else(|| {
                            note_loss_error(
                                part,
                                "VML ClientData child appears outside a note shape",
                            )
                        })?;
                        if event.attributes().with_checks(false).next().is_some() {
                            return Err(note_loss_error(
                                part,
                                "VML ClientData child has unsupported attributes",
                            ));
                        }
                        match tag.as_str() {
                            "Anchor" => {
                                shape.anchor_count += 1;
                                active_text = Some("anchor".into());
                            }
                            "AutoFill" => {
                                shape.auto_fill_count += 1;
                                active_text = Some("autofill".into());
                            }
                            "Row" => {
                                shape.row_count += 1;
                                active_text = Some("row".into());
                            }
                            "Column" => {
                                shape.column_count += 1;
                                active_text = Some("column".into());
                            }
                            "Visible" => {
                                if !is_empty || shape.visible_element {
                                    return Err(note_loss_error(
                                        part,
                                        "VML Visible must be a single empty element",
                                    ));
                                }
                                shape.visible_element = true;
                                if shape.visibility != "visible" {
                                    return Err(note_loss_error(
                                        part,
                                        "VML Visible element conflicts with hidden style",
                                    ));
                                }
                            }
                            _ => unreachable!(),
                        }
                    }
                    "fill" => {
                        let shape = current.as_mut().ok_or_else(|| {
                            note_loss_error(part, "fill appears outside a VML note shape")
                        })?;
                        let attributes = note_attributes(&event, &reader, part)?;
                        if attributes.len() != 1
                            || attributes
                                .get("color2")
                                .map(|value| value.to_ascii_lowercase())
                                != Some("#ffffe1".into())
                        {
                            return Err(note_loss_error(
                                part,
                                "VML note fill uses an unsupported secondary color",
                            ));
                        }
                        shape.has_fill = true;
                    }
                    _ => {
                        return Err(note_loss_error(
                            part,
                            &format!("unsupported VML element `<{tag}>`"),
                        ));
                    }
                }
                if tag == "xml" {
                    saw_root = true;
                }
                if is_empty {
                    match tag.as_str() {
                        "MoveWithCells" | "SizeWithCells" | "fill" | "stroke" | "path"
                        | "shadow" | "Visible" => {
                            active_text = None;
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::Text(event)) => {
                let decoded = event
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                let target = active_text.as_deref().ok_or_else(|| {
                    if decoded.trim().is_empty() {
                        note_loss_error(part, "unexpected VML text")
                    } else {
                        note_loss_error(part, "text appears outside VML ClientData values")
                    }
                });
                if decoded.trim().is_empty() && active_text.is_none() {
                    continue;
                }
                let target = target?;
                let shape = current.as_mut().ok_or_else(|| {
                    note_loss_error(part, "VML text appears outside a note shape")
                })?;
                match target {
                    "anchor" => shape.anchor.push_str(&decoded),
                    "autofill" => shape.autofill.push_str(&decoded),
                    "row" => shape.row = decoded.parse().ok(),
                    "column" => shape.column = decoded.parse().ok(),
                    _ => {}
                }
            }
            Ok(Event::End(event)) => {
                let tag = event.local_name();
                if matches!(
                    tag.as_ref(),
                    b"Anchor" | b"AutoFill" | b"Row" | b"Column" | b"Visible"
                ) {
                    active_text = None;
                } else if tag.as_ref() == b"shape" {
                    let shape = current.take().ok_or_else(|| {
                        note_loss_error(part, "VML note shape has no start element")
                    })?;
                    notes.push(finish_vml_note(shape)?);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid VML XML in `{part}`: {error}"
                )));
            }
        }
    }
    if !saw_root || current.is_some() || notes.is_empty() {
        return Err(note_loss_error(
            part,
            "VML file has no complete legacy note shapes",
        ));
    }
    Ok(notes)
}

fn parse_vml_style(style: &str, part: &str) -> xlsx_jsx::Result<HashMap<String, String>> {
    let mut properties = HashMap::new();
    for item in style.split(';').filter(|item| !item.trim().is_empty()) {
        let Some((key, value)) = item.split_once(':') else {
            return Err(note_loss_error(part, "malformed VML shape style"));
        };
        let key = key.trim().to_ascii_lowercase();
        let value = value.trim().to_ascii_lowercase();
        if properties.insert(key.clone(), value).is_some() {
            return Err(note_loss_error(
                part,
                &format!("duplicate VML style `{key}`"),
            ));
        }
    }
    if properties.keys().any(|key| {
        ![
            "position",
            "margin-left",
            "margin-top",
            "width",
            "height",
            "z-index",
            "visibility",
        ]
        .contains(&key.as_str())
    }) {
        return Err(note_loss_error(
            part,
            "VML note has unsupported style properties",
        ));
    }
    Ok(properties)
}

fn validate_vml_decorative_element(
    tag: &str,
    attributes: &HashMap<String, String>,
    part: &str,
) -> xlsx_jsx::Result<()> {
    let expected = match tag {
        "xml" => [
            ("xmlns:v", "urn:schemas-microsoft-com:vml"),
            ("xmlns:o", "urn:schemas-microsoft-com:office:office"),
            ("xmlns:x", "urn:schemas-microsoft-com:office:excel"),
        ]
        .into_iter()
        .map(|(key, value)| (key.to_owned(), value.to_owned()))
        .collect::<HashMap<_, _>>(),
        "shapelayout" => HashMap::from([("v:ext".to_owned(), "edit".to_owned())]),
        "shapetype" => HashMap::from([
            ("id".to_owned(), "_x0000_t202".to_owned()),
            ("coordsize".to_owned(), "21600,21600".to_owned()),
            ("o:spt".to_owned(), "202".to_owned()),
            ("path".to_owned(), "m,l,21600r21600,l21600,xe".to_owned()),
        ]),
        "stroke" => HashMap::from([("joinstyle".to_owned(), "miter".to_owned())]),
        "shadow" => HashMap::from([
            ("on".to_owned(), "t".to_owned()),
            ("color".to_owned(), "black".to_owned()),
            ("obscured".to_owned(), "t".to_owned()),
        ]),
        "textbox" => HashMap::from([("style".to_owned(), "mso-direction-alt:auto".to_owned())]),
        "div" => HashMap::from([("style".to_owned(), "text-align:left".to_owned())]),
        "idmap" => {
            if attributes.get("v:ext").map(String::as_str) == Some("edit")
                && attributes
                    .get("data")
                    .is_some_and(|value| value.parse::<u32>().is_ok())
                && attributes.len() == 2
            {
                return Ok(());
            }
            return Err(note_loss_error(
                part,
                "VML idmap has unsupported identifiers",
            ));
        }
        "path" => {
            let shape_type_path = HashMap::from([
                ("gradientshapeok".to_owned(), "t".to_owned()),
                ("o:connecttype".to_owned(), "rect".to_owned()),
            ]);
            let shape_path = HashMap::from([("o:connecttype".to_owned(), "none".to_owned())]);
            if attributes == &shape_type_path || attributes == &shape_path {
                return Ok(());
            }
            return Err(note_loss_error(
                part,
                "VML path uses unsupported shape geometry",
            ));
        }
        _ => return Err(note_loss_error(part, "unknown VML decoration")),
    };
    if attributes == &expected {
        Ok(())
    } else {
        Err(note_loss_error(
            part,
            &format!("VML `<{tag}>` uses unsupported visual metadata"),
        ))
    }
}

fn parse_vml_points(value: Option<&String>, part: &str, property: &str) -> xlsx_jsx::Result<f64> {
    let Some(value) = value.and_then(|value| value.strip_suffix("pt")) else {
        return Err(note_loss_error(
            part,
            &format!("VML note `{property}` is missing point units"),
        ));
    };
    let number = value
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite() && *value > 0.0)
        .ok_or_else(|| note_loss_error(part, &format!("VML note `{property}` is invalid")))?;
    Ok(number)
}

fn parse_vml_nonnegative_points(
    value: Option<&String>,
    part: &str,
    property: &str,
) -> xlsx_jsx::Result<f64> {
    let Some(value) = value.and_then(|value| value.strip_suffix("pt")) else {
        return Err(note_loss_error(
            part,
            &format!("VML note `{property}` is missing point units"),
        ));
    };
    value
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite() && *value >= 0.0)
        .ok_or_else(|| note_loss_error(part, &format!("VML note `{property}` is invalid")))
}

fn canonical_rgb(value: &str, part: &str) -> xlsx_jsx::Result<String> {
    let value = value
        .strip_prefix('#')
        .ok_or_else(|| note_loss_error(part, "VML note color is not RGB"))?;
    if value.len() != 6 || !value.chars().all(|character| character.is_ascii_hexdigit()) {
        return Err(note_loss_error(
            part,
            "VML note color is not a six-digit RGB value",
        ));
    }
    Ok(format!("#{}", value.to_ascii_uppercase()))
}

fn finish_vml_note(shape: VmlNoteBuilder) -> xlsx_jsx::Result<ParsedVmlNote> {
    if shape.alt_text.chars().count() > 255 {
        return Err(note_loss_error(
            &shape.part,
            "VML note alt text exceeds Excel's 255-character limit",
        ));
    }
    if !shape.has_client_data
        || shape.move_with_cells != 1
        || shape.size_with_cells != 1
        || !shape.has_fill
        || shape.auto_fill_count != 1
        || shape.anchor_count != 1
        || shape.row_count != 1
        || shape.column_count != 1
        || shape.visible_element != (shape.visibility == "visible")
        || shape.shadow_count != 1
        || shape.shape_path_count != 1
        || shape.textbox_count != 1
        || shape.div_count != 1
        || !matches!(shape.visibility.as_str(), "visible" | "hidden")
        || !shape.autofill.contains("False")
    {
        return Err(note_loss_error(
            &shape.part,
            "VML note shape has unsupported client-data, fill, or visibility semantics",
        ));
    }
    let row = shape
        .row
        .ok_or_else(|| note_loss_error(&shape.part, "VML note is missing its Row"))?;
    let column = shape
        .column
        .ok_or_else(|| note_loss_error(&shape.part, "VML note is missing its Column"))?;
    if row >= 1_048_576 || column >= 16_384 {
        return Err(note_loss_error(
            &shape.part,
            "VML note Row/Column is outside the XLSX worksheet limits",
        ));
    }
    let anchor = shape
        .anchor
        .split(',')
        .map(|value| value.trim().parse::<i32>())
        .collect::<std::result::Result<Vec<_>, _>>()
        .ok()
        .filter(|values| values.len() == 8)
        .ok_or_else(|| {
            note_loss_error(
                &shape.part,
                "VML note anchor must contain eight integer coordinates",
            )
        })?;
    let width = reverse_vml_dimension(shape.width_points, &shape.part, "width")?;
    let height = reverse_vml_dimension(shape.height_points, &shape.part, "height")?;
    Ok(ParsedVmlNote {
        cell: format_a1_cell_reference(row, column),
        visible: shape.visibility == "visible",
        width,
        height,
        background_color: shape.background_color,
        alt_text: shape.alt_text,
        margin_left_points: shape.margin_left_points,
        margin_top_points: shape.margin_top_points,
        z_index: shape.z_index,
        anchor,
    })
}

fn reverse_vml_dimension(points: f64, part: &str, property: &str) -> xlsx_jsx::Result<u32> {
    let rounded_pixels = (points * 4.0 / 3.0).round();
    if !rounded_pixels.is_finite() || rounded_pixels < 1.0 || rounded_pixels > f64::from(u32::MAX) {
        return Err(note_loss_error(
            part,
            &format!("VML note {property} is outside the supported pixel range"),
        ));
    }
    #[expect(
        clippy::cast_possible_truncation,
        reason = "rounded_pixels is checked to be a positive value no greater than u32::MAX."
    )]
    #[expect(
        clippy::cast_sign_loss,
        reason = "rounded_pixels is checked to be at least one before conversion."
    )]
    let pixels = rounded_pixels as u32;
    let serialized_points = (f64::from(pixels) + 0.25).floor() * 72.0 / 96.0;
    if (serialized_points - points).abs() > 1e-8 {
        return Err(note_loss_error(
            part,
            &format!("VML note {property} does not map uniquely to an integer pixel value"),
        ));
    }
    Ok(pixels)
}

fn match_comments_and_vml(
    comments: Vec<WorksheetNote>,
    shapes: Vec<ParsedVmlNote>,
    worksheet_part: &str,
) -> xlsx_jsx::Result<Vec<WorksheetNote>> {
    let mut shapes_by_cell = shapes
        .into_iter()
        .map(|shape| (shape.cell.clone(), shape))
        .collect::<HashMap<_, _>>();
    if shapes_by_cell.len() != comments.len() {
        return Err(Error::Reverse(format!(
            "worksheet `{worksheet_part}` comments and VML note shapes are not one-to-one; repair orphaned or duplicate comments/shape entries"
        )));
    }
    comments.into_iter().map(|mut note| {
        let shape = shapes_by_cell.remove(&note.cell).ok_or_else(|| Error::Reverse(format!(
            "worksheet `{worksheet_part}` note `{}` has no matching VML shape; repair its comments/VML parts", note.cell
        )))?;
        note.visible = shape.visible;
        note.width = shape.width;
        note.height = shape.height;
        note.background_color = shape.background_color;
        note.alt_text = shape.alt_text;
        note.vml_margin_left_points = Some(shape.margin_left_points);
        note.vml_margin_top_points = Some(shape.margin_top_points);
        note.vml_z_index = Some(shape.z_index);
        note.vml_anchor = Some(shape.anchor);
        Ok(note)
    }).collect()
}

fn validate_recompiled_note_geometry(
    metadata: &WorksheetMetadata,
    worksheet_part: &str,
) -> xlsx_jsx::Result<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    if let Some(height) = metadata.default_row_height_pixels {
        worksheet.set_default_row_height_pixels(height);
    }
    for (&column, properties) in &metadata.columns {
        if let Some(width) = properties.width_pixels {
            worksheet
                .set_column_width_pixels(column, width)
                .map_err(|error| Error::Reverse(error.to_string()))?;
        }
        if properties.hidden {
            worksheet
                .set_column_hidden(column)
                .map_err(|error| Error::Reverse(error.to_string()))?;
        }
    }
    for (&row, properties) in &metadata.rows {
        if let Some(height) = properties.height {
            worksheet
                .set_row_height(row, height)
                .map_err(|error| Error::Reverse(error.to_string()))?;
        }
        if properties.hidden {
            worksheet
                .set_row_hidden(row)
                .map_err(|error| Error::Reverse(error.to_string()))?;
        }
    }
    for note in &metadata.notes {
        let (row, column) = parse_a1_cell_reference(&note.cell).ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{worksheet_part}` note has invalid cell `{}`",
                note.cell
            ))
        })?;
        let writer_note = Note::new(&note.text)
            .set_author(note.author.clone())
            .add_author_prefix(note.author_prefix)
            .set_width(note.width)
            .set_height(note.height)
            .set_visible(note.visible)
            .set_background_color(note.background_color.as_str())
            .set_font_name(note.font_name.clone())
            .set_font_size(note.font_size)
            .set_font_family(note.font_family)
            .set_alt_text(note.alt_text.clone());
        worksheet
            .insert_note(row, column, &writer_note)
            .map_err(|error| Error::Reverse(error.to_string()))?;
    }
    let bytes = workbook
        .save_to_buffer()
        .map_err(|error| Error::Reverse(format!("cannot verify note geometry: {error}")))?;
    let mut archive = ZipArchive::new(std::io::Cursor::new(bytes)).map_err(|error| {
        Error::Reverse(format!("cannot inspect generated note geometry: {error}"))
    })?;
    let vml_xml = read_zip_part(&mut archive, "xl/drawings/vmlDrawing1.vml")?;
    let generated = parse_vml_notes(&vml_xml, "generated VML note drawing")?;
    let mut generated_by_cell = generated
        .into_iter()
        .map(|shape| (shape.cell.clone(), shape))
        .collect::<HashMap<_, _>>();
    for note in &metadata.notes {
        let shape = generated_by_cell.remove(&note.cell).ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{worksheet_part}` note `{}` could not be reproduced by rust_xlsxwriter; retain standard note geometry or add matching DSL support",
                note.cell
            ))
        })?;
        let left_matches = note
            .vml_margin_left_points
            .is_some_and(|expected| (expected - shape.margin_left_points).abs() < 1e-8);
        let top_matches = note
            .vml_margin_top_points
            .is_some_and(|expected| (expected - shape.margin_top_points).abs() < 1e-8);
        if !left_matches
            || !top_matches
            || note.vml_z_index != Some(shape.z_index)
            || note.vml_anchor.as_deref() != Some(shape.anchor.as_slice())
        {
            return Err(Error::Reverse(format!(
                "worksheet `{worksheet_part}` note `{}` has VML positioning that the DSL cannot reproduce from its row/column layout; preserve the original row/column geometry or extend Note with matching positioning support",
                note.cell
            )));
        }
    }
    Ok(())
}

fn parse_sheet_targets(
    workbook_xml: &str,
    relationships_xml: &str,
) -> xlsx_jsx::Result<HashMap<String, String>> {
    let relation_ids = parse_sheet_relation_ids(workbook_xml)?;
    let targets = parse_relationship_targets(relationships_xml)?;
    relation_ids
        .into_iter()
        .map(|(name, relation_id)| {
            targets
                .get(&relation_id)
                .cloned()
                .map(|target| (name, target))
                .ok_or_else(|| {
                    Error::Reverse(format!(
                        "worksheet relationship `{relation_id}` has no target"
                    ))
                })
        })
        .collect()
}

fn parse_sheet_relation_ids(xml: &str) -> xlsx_jsx::Result<Vec<(String, String)>> {
    let mut reader = XmlReader::from_str(xml);
    let mut sheets = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event))
                if event.local_name().as_ref() == b"sheet" =>
            {
                let mut name = None;
                let mut relation_id = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"name" => name = Some(value),
                        b"id" => relation_id = Some(value),
                        _ => {}
                    }
                }
                sheets.push((
                    name.ok_or_else(|| Error::Reverse("worksheet is missing its name".into()))?,
                    relation_id.ok_or_else(|| {
                        Error::Reverse("worksheet is missing its relationship id".into())
                    })?,
                ));
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid workbook.xml: {error}"))),
        }
    }
    Ok(sheets)
}

fn parse_relationship_targets(xml: &str) -> xlsx_jsx::Result<HashMap<String, String>> {
    let mut reader = XmlReader::from_str(xml);
    let mut targets = HashMap::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event))
                if event.local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"Id" => id = Some(value),
                        b"Target" => target = Some(value),
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    let target = target.trim_start_matches('/');
                    let target = if target.starts_with("xl/") {
                        target.to_owned()
                    } else {
                        format!("xl/{target}")
                    };
                    targets.insert(id, target);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid workbook relationships: {error}"
                )));
            }
        }
    }
    Ok(targets)
}

fn has_unsupported_part(name: &str) -> bool {
    [
        "xl/threadedComments/",
        "xl/pivot",
        "xl/externalLinks/",
        "xl/queryTables/",
        "xl/embeddings/",
        "xl/ctrlProps/",
        "xl/activeX/",
        "xl/vbaProject",
    ]
    .iter()
    .any(|prefix| name.starts_with(prefix))
}

fn is_supported_part(name: &str) -> bool {
    matches!(
        name,
        "[Content_Types].xml"
            | "_rels/.rels"
            | "docProps/core.xml"
            | "docProps/app.xml"
            | "xl/_rels/workbook.xml.rels"
            | "xl/workbook.xml"
            | "xl/metadata.xml"
            | "xl/featurePropertyBag/featurePropertyBag.xml"
            | "xl/styles.xml"
            | "xl/theme/theme1.xml"
            | "xl/sharedStrings.xml"
    ) || name
        .strip_prefix("xl/media/image")
        .and_then(|suffix| suffix.split_once('.'))
        .is_some_and(|(number, extension)| {
            number.parse::<u32>().is_ok()
                && matches!(extension, "png" | "jpeg" | "jpg" | "gif" | "bmp")
        })
        || name
            .strip_prefix("xl/tables/table")
            .and_then(|name| name.strip_suffix(".xml"))
            .is_some_and(|number| number.parse::<u32>().is_ok())
        || name
            .strip_prefix("xl/worksheets/sheet")
            .and_then(|name| name.strip_suffix(".xml"))
            .is_some_and(|number| number.parse::<u32>().is_ok())
        || name
            .strip_prefix("xl/worksheets/_rels/sheet")
            .and_then(|name| name.strip_suffix(".xml.rels"))
            .is_some_and(|number| number.parse::<u32>().is_ok())
        || name
            .strip_prefix("xl/comments")
            .and_then(|name| name.strip_suffix(".xml"))
            .is_some_and(|number| number.parse::<u32>().is_ok())
        || name
            .strip_prefix("xl/drawings/vmlDrawing")
            .and_then(|name| name.strip_suffix(".vml"))
            .is_some_and(|number| number.parse::<u32>().is_ok())
        || name
            .strip_prefix("xl/charts/chart")
            .and_then(|name| name.strip_suffix(".xml"))
            .is_some_and(|number| number.parse::<u32>().is_ok())
        || name
            .strip_prefix("xl/drawings/drawing")
            .and_then(|name| name.strip_suffix(".xml"))
            .is_some_and(|number| number.parse::<u32>().is_ok())
        || name
            .strip_prefix("xl/drawings/_rels/drawing")
            .and_then(|name| name.strip_suffix(".xml.rels"))
            .is_some_and(|number| number.parse::<u32>().is_ok())
}

fn validate_dynamic_metadata_relationship(
    xml: &str,
    metadata_part_present: bool,
) -> xlsx_jsx::Result<()> {
    let mut reader = XmlReader::from_str(xml);
    let mut relationship_count = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event))
                if event.local_name().as_ref() == b"Relationship" =>
            {
                let mut kind = None;
                let mut target = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"Type" => kind = Some(value),
                        b"Target" => target = Some(value),
                        _ => {}
                    }
                }
                let is_metadata_type = kind
                    .as_deref()
                    .is_some_and(|value| value.ends_with("/sheetMetadata"));
                let is_metadata_target = target.as_deref().is_some_and(|value| {
                    value == "metadata.xml"
                        || value == "/xl/metadata.xml"
                        || value == "xl/metadata.xml"
                });
                if is_metadata_type || is_metadata_target {
                    if !is_metadata_type || !is_metadata_target {
                        return Err(Error::Reverse(
                            "workbook sheetMetadata relationship does not target xl/metadata.xml; repair the relationship target and type".into(),
                        ));
                    }
                    relationship_count += 1;
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid workbook relationships: {error}"
                )));
            }
        }
    }
    if relationship_count > 1 || (relationship_count == 1) != metadata_part_present {
        return Err(Error::Reverse(
            "workbook sheetMetadata relationship and xl/metadata.xml part are missing or duplicated; repair the relationship/part pair".into(),
        ));
    }
    Ok(())
}

fn validate_checkbox_feature_relationship(
    xml: &str,
    property_bag_part_present: bool,
) -> xlsx_jsx::Result<()> {
    let mut reader = XmlReader::from_str(xml);
    let mut relationship_count = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event))
                if event.local_name().as_ref() == b"Relationship" =>
            {
                let mut kind = None;
                let mut target = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"Type" => kind = Some(value),
                        b"Target" => target = Some(value),
                        _ => {}
                    }
                }
                let is_feature_type = kind.as_deref().is_some_and(|value| {
                    value.ends_with("/2022/11/relationships/FeaturePropertyBag")
                });
                let is_feature_target = target.as_deref().is_some_and(|value| {
                    value == "featurePropertyBag/featurePropertyBag.xml"
                        || value == "/xl/featurePropertyBag/featurePropertyBag.xml"
                });
                if is_feature_type || is_feature_target {
                    if !is_feature_type || !is_feature_target {
                        return Err(Error::Reverse(
                            "workbook FeaturePropertyBag relationship has an unexpected type or target; repair the relationship to xl/featurePropertyBag/featurePropertyBag.xml"
                                .into(),
                        ));
                    }
                    relationship_count += 1;
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid workbook relationships: {error}"
                )));
            }
        }
    }
    if relationship_count > 1 || (relationship_count == 1) != property_bag_part_present {
        return Err(Error::Reverse(
            "workbook FeaturePropertyBag relationship and featurePropertyBag part are missing or duplicated; repair the relationship/part pair"
                .into(),
        ));
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "This strict parser validates the pinned writer's complete nested feature-bag record in one pass."
)]
fn audit_checkbox_feature_property_bag(xml: &str) -> xlsx_jsx::Result<()> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut root_seen = false;
    let mut bag_types = std::collections::HashSet::new();
    let mut bag_ids = Vec::new();
    let mut unsupported = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event)) => {
                let tag = event.local_name();
                let mut attributes = HashMap::new();
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let key =
                        String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    if attributes.insert(key.clone(), value).is_some() {
                        return Err(Error::Reverse(format!(
                            "featurePropertyBag `{}` repeats attribute `{key}`; repair the metadata",
                            String::from_utf8_lossy(tag.as_ref())
                        )));
                    }
                }
                match tag.as_ref() {
                    b"FeaturePropertyBags" => {
                        if root_seen {
                            unsupported = Some("duplicate FeaturePropertyBags root".into());
                        }
                        root_seen = attributes.get("xmlns").is_some_and(|value| {
                            value == "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"
                        }) && attributes.len() == 1;
                    }
                    b"bag" => {
                        let kind = attributes
                            .get("type")
                            .map(String::as_str)
                            .unwrap_or_default();
                        let expected = match kind {
                            "Checkbox" | "XFControls" | "XFComplement" => attributes.len() == 1,
                            "XFComplements" => {
                                attributes.len() == 2
                                    && attributes
                                        .get("extRef")
                                        .is_some_and(|value| value == "XFComplementsMapperExtRef")
                            }
                            _ => false,
                        };
                        if !expected || !bag_types.insert(kind.to_owned()) {
                            unsupported = Some(format!("unsupported or duplicate bag `{kind}`"));
                        }
                    }
                    b"bagId" => {
                        let key = attributes.get("k").cloned();
                        if attributes.len() != usize::from(key.is_some()) {
                            unsupported = Some("unexpected bagId attributes".into());
                        }
                        bag_ids.push((key, String::new()));
                    }
                    b"a" => {
                        if attributes.len() != 1
                            || attributes
                                .get("k")
                                .is_none_or(|value| value != "MappedFeaturePropertyBags")
                        {
                            unsupported = Some("unexpected feature-property mapping".into());
                        }
                    }
                    _ => {
                        unsupported = Some(format!(
                            "unsupported element `{}`",
                            String::from_utf8_lossy(tag.as_ref())
                        ));
                    }
                }
            }
            Ok(Event::Text(text)) => {
                if let Some((_, value)) = bag_ids.last_mut() {
                    *value = text
                        .decode()
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid featurePropertyBag: {error}"
                )));
            }
        }
    }
    if let Some(reason) = unsupported {
        return Err(Error::Reverse(format!(
            "xl/featurePropertyBag/featurePropertyBag.xml contains {reason}; only the pinned writer's checkbox metadata is supported"
        )));
    }
    if !root_seen
        || bag_types
            != ["Checkbox", "XFControls", "XFComplement", "XFComplements"]
                .into_iter()
                .map(str::to_owned)
                .collect()
        || bag_ids.len() != 3
        || bag_ids
            != [
                (Some("CellControl".to_owned()), "0".to_owned()),
                (Some("XFControls".to_owned()), "1".to_owned()),
                (None, "2".to_owned()),
            ]
    {
        return Err(Error::Reverse(
            "xl/featurePropertyBag/featurePropertyBag.xml has incomplete checkbox feature bags; repair the package metadata"
                .into(),
        ));
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "The metadata part is validated as one constrained OOXML record set with cross-record counts."
)]
fn audit_dynamic_array_metadata(xml: &str) -> xlsx_jsx::Result<()> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut counts = HashMap::<String, usize>::new();
    let mut types = None;
    let mut future_metadata = None;
    let mut extension_uri = None;
    let mut dynamic_properties = None;
    let mut cell_metadata = None;
    let mut cell_metadata_reference = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event)) => {
                let tag = String::from_utf8_lossy(event.local_name().as_ref()).into_owned();
                *counts.entry(tag.clone()).or_default() += 1;
                if ![
                    "metadata",
                    "metadataTypes",
                    "metadataType",
                    "futureMetadata",
                    "bk",
                    "extLst",
                    "ext",
                    "dynamicArrayProperties",
                    "cellMetadata",
                    "rc",
                ]
                .contains(&tag.as_str())
                {
                    return Err(Error::Reverse(format!(
                        "xl/metadata.xml contains unsupported element `{tag}`; remove other metadata features or add explicit reverse support"
                    )));
                }
                let mut attributes = HashMap::new();
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let key =
                        String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    if attributes.insert(key.clone(), value).is_some() {
                        return Err(Error::Reverse(format!(
                            "xl/metadata.xml `{tag}` repeats attribute `{key}`; repair the metadata"
                        )));
                    }
                }
                let values_are = |expected: &[(&str, &str)]| {
                    expected.len() == attributes.len()
                        && expected.iter().all(|(key, value)| {
                            attributes.get(*key).is_some_and(|actual| actual == value)
                        })
                };
                let valid = match tag.as_str() {
                    "metadata" => attributes.get("xmlns").is_some_and(|value| {
                        value == "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    }) && attributes.get("xda").is_some_and(|value| {
                        value
                            == "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
                    }) && attributes.len() == 2,
                    "metadataTypes" => values_are(&[("count", "1")]),
                    "metadataType" => {
                        types = attributes.get("name").cloned();
                        let expected = [
                            ("name", "XLDAPR"),
                            ("minSupportedVersion", "120000"),
                            ("copy", "1"),
                            ("pasteAll", "1"),
                            ("pasteValues", "1"),
                            ("merge", "1"),
                            ("splitFirst", "1"),
                            ("rowColShift", "1"),
                            ("clearFormats", "1"),
                            ("clearComments", "1"),
                            ("assign", "1"),
                            ("coerce", "1"),
                            ("cellMeta", "1"),
                        ];
                        values_are(&expected)
                    }
                    "futureMetadata" => {
                        future_metadata = attributes.get("name").cloned();
                        values_are(&[("name", "XLDAPR"), ("count", "1")])
                    }
                    "ext" => {
                        extension_uri = attributes.get("uri").cloned();
                        values_are(&[("uri", "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}")])
                    }
                    "dynamicArrayProperties" => {
                        dynamic_properties = attributes.get("fDynamic").cloned();
                        values_are(&[("fDynamic", "1"), ("fCollapsed", "0")])
                    }
                    "cellMetadata" => {
                        cell_metadata = attributes.get("count").cloned();
                        values_are(&[("count", "1")])
                    }
                    "rc" => {
                        cell_metadata_reference = attributes.get("v").cloned();
                        values_are(&[("t", "1"), ("v", "0")])
                    }
                    _ => attributes.is_empty(),
                };
                if !valid {
                    return Err(Error::Reverse(format!(
                        "xl/metadata.xml `{tag}` contains unsupported or contradictory attributes; only the pinned writer's dynamic-array metadata can be preserved"
                    )));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid xl/metadata.xml: {error}"))),
        }
    }
    for (tag, expected) in [
        ("metadata", 1),
        ("metadataTypes", 1),
        ("metadataType", 1),
        ("futureMetadata", 1),
        ("bk", 2),
        ("extLst", 1),
        ("ext", 1),
        ("dynamicArrayProperties", 1),
        ("cellMetadata", 1),
        ("rc", 1),
    ] {
        if counts.get(tag).copied().unwrap_or_default() != expected {
            return Err(Error::Reverse(format!(
                "xl/metadata.xml has unexpected `{tag}` count; repair the dynamic-array metadata structure"
            )));
        }
    }
    if types.as_deref() != Some("XLDAPR")
        || future_metadata.as_deref() != Some("XLDAPR")
        || extension_uri.as_deref() != Some("{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}")
        || dynamic_properties.as_deref() != Some("1")
        || cell_metadata.as_deref() != Some("1")
        || cell_metadata_reference.as_deref() != Some("0")
    {
        return Err(Error::Reverse(
            "xl/metadata.xml dynamic-array records are incomplete or inconsistent; repair the metadata indexes".into(),
        ));
    }
    Ok(())
}

#[expect(
    clippy::struct_excessive_bools,
    reason = "This is a parsed style record whose flags map directly to independent OOXML attributes."
)]
#[derive(Clone, Default)]
struct CellFontStyle {
    bold: bool,
    italic: bool,
    checkbox: bool,
    quote_prefix: Option<()>,
    locked: Option<bool>,
    formula_hidden: bool,
    strikethrough: Option<bool>,
    underline: Option<String>,
    script: Option<String>,
    wrap_text: Option<bool>,
    shrink_to_fit: Option<bool>,
    indent: Option<u8>,
    fill_color: Option<String>,
    fill_pattern: Option<String>,
    fill_foreground_color: Option<String>,
    fill_background_color: Option<String>,
    font_color: Option<String>,
    font_name: Option<String>,
    font_size: Option<f64>,
    font_family: Option<u8>,
    font_charset: Option<u8>,
    font_scheme: Option<String>,
    borders: CellBorders,
    support: StyleSupport,
    number_format: Option<String>,
    horizontal_align: Option<HorizontalAlignment>,
    vertical_align: Option<VerticalAlignment>,
    text_rotation: Option<i16>,
    reading_direction: Option<CellReadingDirection>,
    apply_alignment: bool,
    apply_protection: bool,
}

impl CellFontStyle {
    fn has_conditional_dxf_style(&self) -> bool {
        self.bold
            || self.italic
            || self.font_scheme.is_some()
            || self.strikethrough == Some(true)
            || self.underline.is_some()
            || self.script.is_some()
            || self.font_color.is_some()
            || self.fill_color.is_some()
            || self.fill_pattern.is_some()
            || self.fill_foreground_color.is_some()
            || self.fill_background_color.is_some()
            || self.number_format.is_some()
            || self.borders.has_any()
    }

    fn supports_conditional_dxf(&self) -> bool {
        self.support == StyleSupport::Supported
            && self.font_scheme.is_none()
            && (self.script.is_none()
                || self.bold
                || self.italic
                || self.strikethrough == Some(true)
                || self.underline.is_some()
                || self.font_color.is_some()
                || self.fill_color.is_some()
                || self.fill_pattern.is_some()
                || self.fill_foreground_color.is_some()
                || self.fill_background_color.is_some())
    }
}

#[derive(Clone, Copy)]
enum CellReadingDirection {
    LeftToRight,
    RightToLeft,
}

impl CellReadingDirection {
    const fn jsx_name(self) -> &'static str {
        match self {
            Self::LeftToRight => "leftToRight",
            Self::RightToLeft => "rightToLeft",
        }
    }
}

#[derive(Clone, Copy)]
enum BorderSide {
    Top,
    Right,
    Bottom,
    Left,
}

impl BorderSide {
    const ALL: [Self; 4] = [Self::Top, Self::Right, Self::Bottom, Self::Left];

    const fn index(self) -> usize {
        match self {
            Self::Top => 0,
            Self::Right => 1,
            Self::Bottom => 2,
            Self::Left => 3,
        }
    }

    const fn suffix(self) -> &'static str {
        match self {
            Self::Top => "Top",
            Self::Right => "Right",
            Self::Bottom => "Bottom",
            Self::Left => "Left",
        }
    }
}

#[derive(Clone, Default)]
struct BorderEdge {
    style: Option<FormatBorder>,
    color: Option<String>,
}

#[derive(Clone, Default)]
struct CellBorders {
    edges: [BorderEdge; 4],
    diagonal: BorderEdge,
    diagonal_type: Option<String>,
}

impl CellBorders {
    fn has_any(&self) -> bool {
        self.edges
            .iter()
            .any(|edge| edge.style.is_some() || edge.color.is_some())
            || self.diagonal.style.is_some()
            || self.diagonal.color.is_some()
            || self.diagonal_type.is_some()
    }
}

#[derive(Clone, Copy, Default, PartialEq, Eq)]
enum StyleSupport {
    #[default]
    Unsupported,
    Supported,
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum HorizontalAlignment {
    Left,
    Center,
    Right,
    Distributed,
}

impl HorizontalAlignment {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Distributed => "distributed",
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum VerticalAlignment {
    Top,
    Center,
}

impl VerticalAlignment {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
        }
    }
}

#[derive(Default)]
struct WorksheetMetadata {
    active_worksheet: Option<()>,
    selected: Option<bool>,
    tab_selected: Option<bool>,
    first_tab: Option<()>,
    visibility: WorksheetVisibility,
    tab_color: Option<String>,
    default_row_height_pixels: Option<u32>,
    hide_unused_rows: bool,
    hide_screen_gridlines: bool,
    right_to_left: Option<()>,
    top_left_cell: Option<String>,
    selection: Option<WorksheetSelection>,
    zoom: Option<u16>,
    view: WorksheetView,
    page_breaks: WorksheetPageBreaks,
    summary_below: Option<bool>,
    summary_right: Option<bool>,
    sheet_view_seen: bool,
    print: WorksheetPrintMetadata,
    header_footer: WorksheetHeaderFooter,
    protection: Option<BTreeMap<String, bool>>,
    print_area: Option<String>,
    defined_names: Vec<ReverseDefinedName>,
    repeat_rows: Option<String>,
    repeat_columns: Option<String>,
    cells: HashMap<String, CellFontStyle>,
    rich_text_cells: HashMap<String, Vec<RichTextRun>>,
    array_formula_ranges: HashMap<String, ArrayFormulaRange>,
    cell_style_indices: HashMap<String, usize>,
    hyperlinks: HashMap<String, WorksheetHyperlink>,
    used_hyperlink_relationships: std::collections::HashSet<String>,
    pane_selections: Vec<PaneSelection>,
    rows: HashMap<u32, RowProperties>,
    columns: BTreeMap<u16, ColumnProperties>,
    merged_ranges: Vec<MergeRange>,
    data_validations: Vec<WorksheetDataValidation>,
    conditional_formats: Vec<WorksheetConditionalFormat>,
    ignored_errors: Vec<WorksheetIgnoredError>,
    notes: Vec<WorksheetNote>,
    legacy_drawing_relationship_id: Option<String>,
    data_validation_count: Option<usize>,
    auto_filter: Option<String>,
    auto_filter_columns: Vec<WorksheetFilterColumn>,
    filter_mode: Option<bool>,
    freeze_panes: Option<FreezePanes>,
    table_relationship_ids: Vec<String>,
    tables: Vec<WorksheetTable>,
    drawing_relationship_id: Option<String>,
    charts: Vec<WorksheetChart>,
    images: Vec<WorksheetImage>,
    drawing_order: Vec<WorksheetDrawingOrder>,
}

enum WorksheetDrawingOrder {
    Chart(usize),
    Image(usize),
}

#[derive(Clone)]
struct WorksheetChart {
    chart_type: &'static str,
    cell: String,
    width: u32,
    height: u32,
    name: String,
    title: Option<String>,
    x_axis_title: Option<String>,
    y_axis_title: Option<String>,
    show_legend: bool,
    legend_position: Option<&'static str>,
    hole_size: Option<u8>,
    first_slice_angle: Option<u16>,
    alt_text: Option<String>,
    series: Vec<WorksheetChartSeries>,
}

#[derive(Clone)]
struct WorksheetImage {
    cell: String,
    width: u32,
    height: u32,
    data_uri: String,
    offset_x: u32,
    offset_y: u32,
    movement: Option<&'static str>,
    alt_text: Option<String>,
    decorative: bool,
}

#[derive(Clone)]
struct WorksheetChartSeries {
    values: String,
    categories: Option<String>,
    name: Option<String>,
    title: Option<String>,
    marker_seen: bool,
    marker_none: bool,
}

#[derive(Clone)]
struct WorksheetNote {
    cell: String,
    text: String,
    author: String,
    author_prefix: bool,
    visible: bool,
    width: u32,
    height: u32,
    background_color: String,
    font_name: String,
    font_size: f64,
    font_family: u8,
    alt_text: String,
    vml_margin_left_points: Option<f64>,
    vml_margin_top_points: Option<f64>,
    vml_z_index: Option<u32>,
    vml_anchor: Option<Vec<i32>>,
}

struct WorksheetNoteParts {
    vml_relationship_id: String,
    notes: Vec<WorksheetNote>,
}

#[derive(Clone)]
struct WorksheetIgnoredError {
    range: String,
    error_type: &'static str,
}

#[derive(Default)]
struct WorksheetHeaderFooter {
    header: Option<String>,
    footer: Option<String>,
    scale_with_doc: Option<bool>,
    align_with_page: Option<bool>,
}

#[derive(Default)]
struct WorksheetPrintMetadata {
    options: WorksheetPrintOptions,
    quality: WorksheetPrintQuality,
    options_seen: Option<()>,
    page_setup_seen: Option<()>,
    page_margins_seen: Option<()>,
    page_margins: WorksheetPageMargins,
    page_setup_properties_seen: Option<()>,
    fit_to_page: Option<bool>,
    fit_width: Option<u16>,
    fit_height: Option<u16>,
    fit_to_pages: Option<FitToPages>,
    print_scale: Option<u16>,
    orientation: Option<&'static str>,
    paper_size: Option<u8>,
    page_order: Option<&'static str>,
}

#[derive(Default)]
struct WorksheetPrintOptions {
    gridlines: bool,
    headings: bool,
    centering: WorksheetPrintCentering,
}

#[derive(Default)]
struct WorksheetPrintCentering {
    horizontally: bool,
    vertically: bool,
}

#[derive(Default)]
struct WorksheetPrintQuality {
    black_and_white: bool,
    draft: bool,
}

#[derive(Clone, Copy)]
struct FitToPages {
    width: u16,
    height: u16,
}

#[derive(Clone, Copy)]
struct WorksheetPageMargins {
    left: f64,
    right: f64,
    top: f64,
    bottom: f64,
    header: f64,
    footer: f64,
}

impl Default for WorksheetPageMargins {
    fn default() -> Self {
        Self {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3,
        }
    }
}

#[derive(Default)]
struct WorksheetPageBreaks {
    horizontal: Vec<u32>,
    vertical: Vec<u32>,
}

#[derive(Clone)]
struct ArrayFormulaRange {
    reference: String,
    coordinates: (u32, u16, u32, u16),
    is_dynamic: bool,
    has_explicit_range: bool,
}

#[derive(Clone, Default)]
struct RichTextRun {
    text: String,
    format: CellFontStyle,
}

#[derive(Clone, Default)]
#[expect(
    clippy::struct_excessive_bools,
    reason = "Each bool is an independent OOXML tableStyleInfo/autoFilter flag and must survive reverse separately."
)]
struct WorksheetTable {
    range: String,
    name: String,
    style: String,
    auto_filter: bool,
    banded_rows: bool,
    banded_columns: bool,
    first_column: bool,
    last_column: bool,
    headers: Vec<String>,
    totals_row: bool,
    alt_text: Option<String>,
    alt_text_title: Option<String>,
    columns: Vec<WorksheetTableColumn>,
}

#[derive(Clone, Default)]
struct WorksheetTableColumn {
    total_label: Option<String>,
    total_function: Option<String>,
    total_formula: Option<String>,
    formula: Option<String>,
    format: Option<CellFontStyle>,
}

impl WorksheetTableColumn {
    fn is_empty(&self) -> bool {
        self.total_label.is_none()
            && self.total_function.is_none()
            && self.total_formula.is_none()
            && self.formula.is_none()
            && self.format.is_none()
    }
}

#[derive(Clone, Default)]
struct WorksheetHyperlink {
    target: String,
    display: Option<String>,
    tooltip: Option<String>,
}

#[derive(Clone, Default)]
struct WorksheetConditionalFormat {
    range: String,
    scale_type: String,
    priority: Option<u32>,
    top_rule: Option<String>,
    rank: Option<u16>,
    stop_if_true: bool,
    guid: Option<String>,
    operator: Option<String>,
    text: Option<String>,
    formula1: Option<String>,
    formula2: Option<String>,
    format: CellFontStyle,
    minimum: ColorScaleThreshold,
    midpoint: Option<ColorScaleThreshold>,
    maximum: ColorScaleThreshold,
    data_bar: Option<DataBarProperties>,
    icon_set: Option<IconSetProperties>,
}

#[derive(Clone, Default)]
struct DataBarProperties {
    minimum: ColorScaleThreshold,
    maximum: ColorScaleThreshold,
    fill_color: Option<String>,
    border_color: Option<String>,
    negative_fill_color: Option<String>,
    negative_border_color: Option<String>,
    axis_color: Option<String>,
    bar_only: Option<bool>,
    solid_fill: Option<bool>,
    border: Option<bool>,
    direction: Option<String>,
    axis_position: Option<String>,
}

#[derive(Clone, Default)]
struct IconSetProperties {
    icon_type: String,
    reverse: bool,
    show_icons_only: bool,
    icons: Vec<IconThresholdProperties>,
}

#[derive(Clone, Default)]
struct IconThresholdProperties {
    kind: String,
    value: Option<String>,
    greater_than: bool,
    icon_type: Option<String>,
    icon_index: Option<u8>,
    no_icon: bool,
}

#[derive(Clone, Default)]
struct ColorScaleThreshold {
    kind: String,
    value: Option<String>,
    color: String,
}

#[derive(Default, serde::Serialize)]
struct WorksheetFilterColumn {
    column: u16,
    #[serde(skip_serializing_if = "Option::is_none")]
    values: Option<Vec<String>>,
    #[serde(rename = "includeBlanks", skip_serializing_if = "Option::is_none")]
    include_blanks: Option<bool>,
    #[serde(rename = "customFilters", skip_serializing_if = "Option::is_none")]
    custom_filters: Option<WorksheetCustomFilters>,
}

#[derive(Default, serde::Serialize)]
struct WorksheetCustomFilters {
    logic: String,
    conditions: Vec<WorksheetCustomFilterCondition>,
}

#[derive(serde::Serialize)]
struct WorksheetCustomFilterCondition {
    #[serde(skip_serializing_if = "Option::is_none")]
    operator: Option<String>,
    value: String,
}

#[derive(Default)]
struct AutoFilterAuditState {
    current: Option<WorksheetFilterColumn>,
    in_filters: bool,
    in_custom_filters: bool,
    filter_kind_started: bool,
}

impl AutoFilterAuditState {
    fn start(
        &mut self,
        event: &BytesStart<'_>,
        tag: &str,
        part: &str,
        decoder: quick_xml::encoding::Decoder,
    ) -> xlsx_jsx::Result<()> {
        match tag {
            "filterColumn" => self.start_filter_column(event, part, decoder),
            "filters" => self.start_filters(event, part, decoder),
            "filter" => self.start_filter(event, part, decoder),
            "customFilters" => self.start_custom_filters(event, part, decoder),
            "customFilter" => self.start_custom_filter(event, part, decoder),
            _ => Ok(()),
        }
    }

    fn start_filter_column(
        &mut self,
        event: &BytesStart<'_>,
        part: &str,
        decoder: quick_xml::encoding::Decoder,
    ) -> xlsx_jsx::Result<()> {
        if self.current.is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has nested filterColumn elements; repair the autoFilter structure"
            )));
        }
        let mut column = None;
        for attribute in event.attributes().with_checks(false) {
            let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?;
            if attribute.key.local_name().as_ref() != b"colId" || column.is_some() {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` filterColumn has unsupported metadata; remove it or add matching DSL support"
                )));
            }
            column = Some(value.parse::<u16>().map_err(|error| {
                Error::Reverse(format!(
                    "worksheet `{part}` filterColumn has invalid colId `{value}`: {error}"
                ))
            })?);
        }
        let column = column.ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{part}` filterColumn is missing colId; repair the filter metadata"
            ))
        })?;
        self.current = Some(WorksheetFilterColumn {
            column,
            ..Default::default()
        });
        self.filter_kind_started = false;
        Ok(())
    }

    fn start_filters(
        &mut self,
        event: &BytesStart<'_>,
        part: &str,
        decoder: quick_xml::encoding::Decoder,
    ) -> xlsx_jsx::Result<()> {
        if self.current.is_none()
            || self.in_filters
            || self.in_custom_filters
            || self.filter_kind_started
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has misplaced or nested filters; repair the autoFilter structure"
            )));
        }
        for attribute in event.attributes().with_checks(false) {
            let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
            if attribute.key.local_name().as_ref() != b"blank" {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` has unsupported filters attribute `{}`; only blank is modeled",
                    String::from_utf8_lossy(attribute.key.local_name().as_ref())
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?;
            if parse_xml_bool(&value, part, "filters.blank")? {
                self.current
                    .as_mut()
                    .ok_or_else(|| Error::Reverse("filterColumn state is missing".into()))?
                    .include_blanks = Some(true);
            }
        }
        self.filter_kind_started = true;
        self.in_filters = true;
        Ok(())
    }

    fn start_filter(
        &mut self,
        event: &BytesStart<'_>,
        part: &str,
        decoder: quick_xml::encoding::Decoder,
    ) -> xlsx_jsx::Result<()> {
        if !self.in_filters || self.in_custom_filters {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has a filter outside filters; repair the autoFilter structure"
            )));
        }
        let mut value = None;
        for attribute in event.attributes().with_checks(false) {
            let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
            if attribute.key.local_name().as_ref() != b"val" || value.is_some() {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` has an unsupported filter attribute; only string val is modeled"
                )));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| Error::Reverse(error.to_string()))?
                    .into_owned(),
            );
        }
        let value = value.ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{part}` has a missing filter val; add the exact list value"
            ))
        })?;
        let current = self
            .current
            .as_mut()
            .ok_or_else(|| Error::Reverse("filterColumn state is missing".into()))?;
        current.values.get_or_insert_with(Vec::new).push(value);
        Ok(())
    }

    fn start_custom_filters(
        &mut self,
        event: &BytesStart<'_>,
        part: &str,
        decoder: quick_xml::encoding::Decoder,
    ) -> xlsx_jsx::Result<()> {
        if self.current.is_none()
            || self.in_filters
            || self.in_custom_filters
            || self.filter_kind_started
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has misplaced or duplicate customFilters; repair the filterColumn structure"
            )));
        }
        let mut logic = "or";
        for attribute in event.attributes().with_checks(false) {
            let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
            if attribute.key.local_name().as_ref() != b"and" {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` customFilters has unsupported attribute `{}`; only and is modeled",
                    String::from_utf8_lossy(attribute.key.local_name().as_ref())
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?;
            if parse_xml_bool(&value, part, "customFilters.and")? {
                logic = "and";
            }
        }
        self.current
            .as_mut()
            .ok_or_else(|| Error::Reverse("filterColumn state is missing".into()))?
            .custom_filters = Some(WorksheetCustomFilters {
            logic: logic.to_string(),
            conditions: vec![],
        });
        self.filter_kind_started = true;
        self.in_custom_filters = true;
        Ok(())
    }

    fn start_custom_filter(
        &mut self,
        event: &BytesStart<'_>,
        part: &str,
        decoder: quick_xml::encoding::Decoder,
    ) -> xlsx_jsx::Result<()> {
        if !self.in_custom_filters || self.in_filters {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has a customFilter outside customFilters; repair the filter structure"
            )));
        }
        let mut operator = None;
        let mut value = None;
        for attribute in event.attributes().with_checks(false) {
            let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
            let attribute_value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?
                .into_owned();
            match attribute.key.local_name().as_ref() {
                b"operator" if operator.is_none() => {
                    if ![
                        "lessThan",
                        "notEqual",
                        "greaterThan",
                        "lessThanOrEqual",
                        "greaterThanOrEqual",
                    ]
                    .contains(&attribute_value.as_str())
                    {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` customFilter operator `{attribute_value}` is unsupported; repair it to an OOXML comparison operator"
                        )));
                    }
                    operator = Some(attribute_value);
                }
                b"val" if value.is_none() => value = Some(attribute_value),
                _ => {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` customFilter has unsupported or duplicate metadata; only operator and val are modeled"
                    )));
                }
            }
        }
        let value = value.ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{part}` customFilter is missing val; repair the custom condition"
            ))
        })?;
        let filters = self
            .current
            .as_mut()
            .and_then(|column| column.custom_filters.as_mut())
            .ok_or_else(|| Error::Reverse("customFilters state is missing".into()))?;
        filters
            .conditions
            .push(WorksheetCustomFilterCondition { operator, value });
        Ok(())
    }

    fn end(
        &mut self,
        event: &quick_xml::events::BytesEnd<'_>,
        part: &str,
        metadata: &mut WorksheetMetadata,
    ) -> xlsx_jsx::Result<()> {
        match event.local_name().as_ref() {
            b"filters" => self.in_filters = false,
            b"customFilters" => self.in_custom_filters = false,
            b"filterColumn" => {
                self.in_filters = false;
                let filter = self.current.take().ok_or_else(|| {
                    Error::Reverse(format!(
                        "worksheet `{part}` has an unmatched filterColumn end tag"
                    ))
                })?;
                let has_list_filter = filter
                    .values
                    .as_ref()
                    .is_some_and(|values| !values.is_empty())
                    || filter.include_blanks == Some(true);
                let has_custom_filter = filter.custom_filters.as_ref().is_some_and(|custom| {
                    (1..=2).contains(&custom.conditions.len())
                        && !(custom.conditions.len() == 1 && custom.logic == "and")
                });
                if has_list_filter == has_custom_filter {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` filterColumn must contain exactly one supported list or custom filter; repair its criteria"
                    )));
                }
                metadata.auto_filter_columns.push(filter);
            }
            _ => {}
        }
        Ok(())
    }

    fn finish(&self, part: &str, metadata: &WorksheetMetadata) -> xlsx_jsx::Result<()> {
        if self.current.is_some() || self.in_filters || self.in_custom_filters {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has incomplete autoFilter criteria; repair filterColumn/filters nesting"
            )));
        }
        let Some(reference) = metadata.auto_filter.as_deref() else {
            if metadata.auto_filter_columns.is_empty() {
                if metadata.filter_mode == Some(true) {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` declares filterMode=true without filter criteria; remove filterMode or add an autoFilter with criteria"
                    )));
                }
                return Ok(());
            }
            return Err(Error::Reverse(format!(
                "worksheet `{part}` contains filter criteria without an autoFilter range; repair the worksheet metadata"
            )));
        };
        let (_, first_column, _, last_column) = parse_a1_range(reference).ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{part}` has an invalid autoFilter range"
            ))
        })?;
        let width = last_column - first_column + 1;
        let mut seen = std::collections::HashSet::new();
        for filter in &metadata.auto_filter_columns {
            if filter.column >= width {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` filterColumn colId {} is outside autoFilter width {width}; repair the range or column index",
                    filter.column
                )));
            }
            if !seen.insert(filter.column) {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` contains duplicate filterColumn colId {}; combine its values",
                    filter.column
                )));
            }
        }
        if metadata.filter_mode == Some(false) && !metadata.auto_filter_columns.is_empty() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` declares filterMode=false but contains active filter criteria; repair the sheetPr/filterColumn metadata"
            )));
        }
        if metadata.filter_mode == Some(true) && metadata.auto_filter_columns.is_empty() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` declares filterMode=true without filter criteria; remove filterMode or add filterColumn criteria"
            )));
        }
        Ok(())
    }
}

#[derive(Default)]
#[expect(clippy::struct_excessive_bools)] // These flags are independent SpreadsheetML attributes.
struct WorksheetDataValidation {
    range: String,
    validation_type: String,
    formula1: Option<String>,
    formula2: Option<String>,
    operator: Option<String>,
    allow_blank: bool,
    show_dropdown: bool,
    show_input_message: bool,
    show_error_message: bool,
    input_title: Option<String>,
    input_message: Option<String>,
    error_title: Option<String>,
    error_message: Option<String>,
    error_style: Option<String>,
}

#[derive(Default)]
struct DataValidationAuditState {
    current: Option<WorksheetDataValidation>,
    reading_formula: Option<String>,
}

impl DataValidationAuditState {
    fn empty(
        &mut self,
        event: &BytesStart<'_>,
        part: &str,
        decoder: quick_xml::encoding::Decoder,
        metadata: &mut WorksheetMetadata,
    ) -> xlsx_jsx::Result<()> {
        if self.current.is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` contains nested dataValidation elements; repair the worksheet XML"
            )));
        }
        let validation = parse_data_validation_attributes(event, part, decoder)?;
        validate_external_data_validation(&validation, part)?;
        metadata.data_validations.push(validation);
        Ok(())
    }

    fn start(
        &mut self,
        event: &BytesStart<'_>,
        tag: &str,
        part: &str,
        decoder: quick_xml::encoding::Decoder,
        metadata: &mut WorksheetMetadata,
    ) -> xlsx_jsx::Result<()> {
        match tag {
            "dataValidations" => audit_data_validation_container(
                event,
                part,
                decoder,
                &mut metadata.data_validation_count,
            ),
            "dataValidation" => {
                if self.current.is_some() {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` contains nested dataValidation elements; repair the worksheet XML"
                    )));
                }
                self.current = Some(parse_data_validation_attributes(event, part, decoder)?);
                Ok(())
            }
            "formula1" | "formula2" => {
                let Some(validation) = self.current.as_ref() else {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` contains formula1 outside a dataValidation; repair the validation XML"
                    )));
                };
                if self.reading_formula.is_some()
                    || if tag == "formula1" {
                        validation.formula1.is_some()
                    } else {
                        validation.formula2.is_some()
                    }
                {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` contains duplicate/nested {tag}; keep one formula element of each kind"
                    )));
                }
                self.reading_formula = Some(tag.to_string());
                Ok(())
            }
            _ => Ok(()),
        }
    }

    fn text(&mut self, text: &quick_xml::events::BytesText<'_>) -> xlsx_jsx::Result<()> {
        if let Some(formula_tag) = self.reading_formula.as_deref() {
            let decoded = text
                .decode()
                .map_err(|error| Error::Reverse(error.to_string()))?;
            let unescaped = quick_xml::escape::unescape(&decoded)
                .map_err(|error| Error::Reverse(error.to_string()))?;
            if let Some(validation) = &mut self.current {
                let formula = if formula_tag == "formula1" {
                    &mut validation.formula1
                } else {
                    &mut validation.formula2
                };
                formula.get_or_insert_with(String::new).push_str(&unescaped);
            }
        }
        Ok(())
    }

    fn cdata(&mut self, text: &quick_xml::events::BytesCData<'_>) -> xlsx_jsx::Result<()> {
        if let Some(formula_tag) = self.reading_formula.as_deref() {
            let decoded = text
                .decode()
                .map_err(|error| Error::Reverse(error.to_string()))?;
            if let Some(validation) = &mut self.current {
                let formula = if formula_tag == "formula1" {
                    &mut validation.formula1
                } else {
                    &mut validation.formula2
                };
                formula.get_or_insert_with(String::new).push_str(&decoded);
            }
        }
        Ok(())
    }

    fn end(
        &mut self,
        event: &quick_xml::events::BytesEnd<'_>,
        part: &str,
        metadata: &mut WorksheetMetadata,
    ) -> xlsx_jsx::Result<()> {
        match event.local_name().as_ref() {
            b"formula1" | b"formula2" => self.reading_formula = None,
            b"dataValidation" => {
                let validation = self.current.take().ok_or_else(|| {
                    Error::Reverse(format!(
                        "worksheet `{part}` has an unmatched dataValidation end tag"
                    ))
                })?;
                validate_external_data_validation(&validation, part)?;
                metadata.data_validations.push(validation);
            }
            _ => {}
        }
        Ok(())
    }

    fn finish(&self, part: &str, metadata: &WorksheetMetadata) -> xlsx_jsx::Result<()> {
        if self.current.is_some() || self.reading_formula.is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has an incomplete data validation; repair its formula1/dataValidation elements"
            )));
        }
        if metadata
            .data_validation_count
            .is_some_and(|count| count != metadata.data_validations.len())
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` dataValidations count does not match the number of validation rules; repair the container count"
            )));
        }
        validate_data_validation_ranges(part, &metadata.data_validations)
    }
}

#[derive(Clone, Copy)]
struct MergeRange {
    first_row: u32,
    first_column: u16,
    last_row: u32,
    last_column: u16,
}

#[derive(Clone, Copy, PartialEq, Eq)]
struct FreezePanes {
    rows: u32,
    columns: u16,
    top_left_row: u32,
    top_left_column: u16,
}

#[derive(Clone, PartialEq, Eq)]
struct PaneSelection {
    pane: Option<String>,
    active_cell: Option<String>,
    sqref: Option<String>,
}

#[derive(Clone)]
struct WorksheetSelection {
    active_cell: String,
    range: String,
}

#[derive(Clone, Copy, Default, PartialEq, Eq)]
struct ColumnProperties {
    width_pixels: Option<u32>,
    hidden: bool,
    outline_level: u8,
    collapsed: bool,
    style_index: Option<usize>,
}

#[derive(Clone, Copy, Default)]
struct RowProperties {
    height: Option<f64>,
    hidden: bool,
    outline_level: u8,
    collapsed: bool,
    style_index: Option<usize>,
}

#[derive(Clone, Copy, Default, PartialEq, Eq)]
enum DateNumberFormat {
    #[default]
    General,
    Date,
    DateTime,
    DateTimeMillis,
    Unsupported,
}

fn cell_reference(row: u32, column: u32) -> String {
    let mut column_number = column + 1;
    let mut letters = Vec::new();
    while column_number > 0 {
        let remainder = (column_number - 1) % 26;
        let letter = char::from_u32(u32::from(b'A') + remainder).unwrap_or('A');
        letters.push(letter);
        column_number = (column_number - 1) / 26;
    }
    let column_name: String = letters.into_iter().rev().collect();
    format!("{column_name}{}", row + 1)
}

const SUPPORTED_WORKSHEET_TAGS: &[&str] = &[
    "worksheet",
    "sheetPr",
    "pageSetUpPr",
    "tabColor",
    "outlinePr",
    "dimension",
    "sheetViews",
    "sheetView",
    "selection",
    "sheetFormatPr",
    "cols",
    "col",
    "sheetData",
    "row",
    "c",
    "f",
    "v",
    "is",
    "r",
    "rPr",
    "b",
    "i",
    "strike",
    "u",
    "vertAlign",
    "rFont",
    "sz",
    "family",
    "charset",
    "scheme",
    "t",
    "pageMargins",
    "pageSetup",
    "sheetProtection",
    "headerFooter",
    "oddHeader",
    "oddFooter",
    "printOptions",
    "rowBreaks",
    "colBreaks",
    "brk",
    "autoFilter",
    "filterColumn",
    "filters",
    "filter",
    "customFilters",
    "customFilter",
    "pane",
    "mergeCells",
    "mergeCell",
    "dataValidations",
    "dataValidation",
    "formula1",
    "formula2",
    "formula",
    "hyperlinks",
    "hyperlink",
    "drawing",
    "legacyDrawing",
    "tableParts",
    "tablePart",
    "conditionalFormatting",
    "ignoredErrors",
    "ignoredError",
    "cfRule",
    "colorScale",
    "dataBar",
    "iconSet",
    "cfvo",
    "color",
];

const UNSUPPORTED_WORKSHEET_TAGS: &[&str] = &[
    "protectedRanges",
    "protectedRange",
    "evenHeader",
    "evenFooter",
    "firstHeader",
    "firstFooter",
    "legacyDrawingHF",
    "extLst",
    "rPh",
    "phoneticPr",
];

#[expect(
    clippy::too_many_lines,
    reason = "Worksheet XML auditing shares ordered cell, metadata, and feature checks in one streaming pass."
)]
fn audit_worksheet_xml(
    xml: &str,
    part: &str,
    formats: &[CellFontStyle],
    differential_formats: &[CellFontStyle],
    hyperlink_relationships: &HashMap<String, String>,
    shared_strings: &[Option<Vec<RichTextRun>>],
    has_dynamic_metadata: bool,
) -> xlsx_jsx::Result<WorksheetMetadata> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut metadata = WorksheetMetadata {
        rich_text_cells: parse_worksheet_rich_text(xml, part, shared_strings)?,
        ignored_errors: parse_ignored_errors(xml, part)?,
        array_formula_ranges: parse_array_formula_ranges(xml, part, has_dynamic_metadata)?,
        page_breaks: parse_worksheet_page_breaks(xml, part)?,
        header_footer: parse_worksheet_header_footer(xml, part)?,
        protection: parse_worksheet_protection(xml, part)?,
        ..Default::default()
    };
    let mut next_row_index = 0_u32;
    let mut seen_rows = std::collections::HashSet::new();
    let mut validation_audit = DataValidationAuditState::default();
    let mut auto_filter_audit = AutoFilterAuditState::default();
    let mut in_inline_string = false;
    loop {
        match reader.read_event() {
            Ok(event @ (Event::Start(_) | Event::Empty(_))) => {
                let is_empty = matches!(&event, Event::Empty(_));
                let (Event::Start(event) | Event::Empty(event)) = &event else {
                    unreachable!("the event was matched as a start or empty element");
                };
                let local_name = event.local_name();
                let tag = String::from_utf8_lossy(local_name.as_ref());
                if tag == "is" {
                    in_inline_string = true;
                }
                if matches!(
                    tag.as_ref(),
                    "r" | "rPr"
                        | "b"
                        | "i"
                        | "strike"
                        | "u"
                        | "vertAlign"
                        | "rFont"
                        | "sz"
                        | "family"
                        | "charset"
                        | "scheme"
                        | "t"
                ) && !in_inline_string
                {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` contains `<{tag}>` outside an inline string; repair the cell XML"
                    )));
                }
                if tag == "extLst" {
                    if !is_empty {
                        let mut depth = 1_u32;
                        while depth > 0 {
                            match reader.read_event() {
                                Ok(Event::Start(_)) => depth += 1,
                                Ok(Event::End(_)) => depth -= 1,
                                Ok(Event::Eof) => {
                                    return Err(Error::Reverse(format!(
                                        "worksheet `{part}` has an unterminated extLst; repair the worksheet XML"
                                    )));
                                }
                                Ok(_) => {}
                                Err(error) => {
                                    return Err(Error::Reverse(format!(
                                        "invalid extension XML in `{part}`: {error}"
                                    )));
                                }
                            }
                        }
                    }
                    continue;
                }
                if UNSUPPORTED_WORKSHEET_TAGS.contains(&tag.as_ref()) {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` contains unsupported `<{tag}>` semantics; remove the feature or add matching JSX and reverse support"
                    )));
                }
                if !SUPPORTED_WORKSHEET_TAGS.contains(&tag.as_ref()) {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` contains unrecognized `<{tag}>` XML that v1 cannot preserve; remove the feature or add explicit JSX and reverse support"
                    )));
                }
                if is_empty && tag == "dataValidation" {
                    validation_audit.empty(event, part, reader.decoder(), &mut metadata)?;
                } else {
                    validation_audit.start(
                        event,
                        tag.as_ref(),
                        part,
                        reader.decoder(),
                        &mut metadata,
                    )?;
                }
                auto_filter_audit.start(event, tag.as_ref(), part, reader.decoder())?;
                audit_worksheet_element_metadata(
                    event,
                    tag.as_ref(),
                    part,
                    reader.decoder(),
                    formats,
                    hyperlink_relationships,
                    &mut metadata,
                )?;
                if tag == "legacyDrawing" {
                    let mut relationship_id = None;
                    for attribute in event.attributes().with_checks(false) {
                        let attribute =
                            attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                        if attribute.key.as_ref() == b"r:id" {
                            relationship_id = Some(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?
                                    .into_owned(),
                            );
                        } else {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` legacyDrawing has unsupported attribute `{}`; only r:id is losslessly supported",
                                String::from_utf8_lossy(attribute.key.as_ref())
                            )));
                        }
                    }
                    let Some(relationship_id) = relationship_id else {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` legacyDrawing is missing r:id; repair the VML relationship"
                        )));
                    };
                    if metadata
                        .legacy_drawing_relationship_id
                        .replace(relationship_id)
                        .is_some()
                    {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` contains multiple legacyDrawing elements; retain one VML note drawing relationship"
                        )));
                    }
                }
                if tag == "drawing" {
                    let mut relationship_id = None;
                    for attribute in event.attributes().with_checks(false) {
                        let attribute =
                            attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                        if attribute.key.as_ref() == b"r:id" {
                            relationship_id = Some(
                                attribute
                                    .decoded_and_normalized_value(
                                        XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Reverse(error.to_string()))?
                                    .into_owned(),
                            );
                        } else {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` drawing has unsupported attribute `{}`; only r:id is supported",
                                String::from_utf8_lossy(attribute.key.as_ref())
                            )));
                        }
                    }
                    let Some(relationship_id) = relationship_id else {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` drawing is missing r:id; repair its drawing relationship"
                        )));
                    };
                    if metadata
                        .drawing_relationship_id
                        .replace(relationship_id)
                        .is_some()
                    {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` contains multiple drawing elements; use one worksheet drawing part"
                        )));
                    }
                }
                if tag == "row" {
                    audit_worksheet_row(
                        event,
                        part,
                        reader.decoder(),
                        formats,
                        &mut next_row_index,
                        &mut seen_rows,
                        &mut metadata.rows,
                    )?;
                }
            }
            Ok(Event::Text(text)) => validation_audit.text(&text)?,
            Ok(Event::CData(text)) => validation_audit.cdata(&text)?,
            Ok(Event::End(event)) => {
                if event.local_name().as_ref() == b"is" || event.local_name().as_ref() == b"c" {
                    in_inline_string = false;
                }
                validation_audit.end(&event, part, &mut metadata)?;
                auto_filter_audit.end(&event, part, &mut metadata)?;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid XML in `{part}`: {error}"))),
        }
    }
    validation_audit.finish(part, &metadata)?;
    auto_filter_audit.finish(part, &metadata)?;
    validate_page_setup_consistency(&mut metadata, part)?;
    let row_outline = metadata
        .rows
        .iter()
        .map(|(&row, properties)| {
            (
                row,
                OutlineProperties {
                    hidden: properties.hidden,
                    level: properties.outline_level,
                    collapsed: properties.collapsed,
                },
            )
        })
        .collect::<BTreeMap<_, _>>();
    collapsed_outline_ranges(&row_outline, "row")
        .map_err(|error| Error::Reverse(format!("worksheet `{part}` {error}")))?;
    let column_outline = metadata
        .columns
        .iter()
        .map(|(&column, properties)| {
            (
                u32::from(column),
                OutlineProperties {
                    hidden: properties.hidden,
                    level: properties.outline_level,
                    collapsed: properties.collapsed,
                },
            )
        })
        .collect::<BTreeMap<_, _>>();
    collapsed_outline_ranges(&column_outline, "column")
        .map_err(|error| Error::Reverse(format!("worksheet `{part}` {error}")))?;
    validate_merged_ranges(part, &metadata)?;
    metadata.selection = validate_pane_selections(&metadata, &metadata.pane_selections, part)?;
    metadata.conditional_formats = parse_conditional_formats(xml, part, differential_formats)?;
    merge_x14_data_bars(xml, part, &mut metadata.conditional_formats)?;
    merge_x14_icon_sets(xml, part, &mut metadata.conditional_formats)?;
    if metadata.used_hyperlink_relationships.len() != hyperlink_relationships.len() {
        let unused = hyperlink_relationships
            .keys()
            .find(|id| !metadata.used_hyperlink_relationships.contains(*id))
            .map_or("unknown", String::as_str);
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has unreferenced hyperlink relationship `{unused}`; repair or remove orphan relationships"
        )));
    }
    Ok(metadata)
}

fn conditional_text_operator_from_xml(
    rule_type: &str,
    operator: &str,
    part: &str,
) -> xlsx_jsx::Result<&'static str> {
    match (rule_type, operator) {
        ("containsText", "containsText") => Ok("containsText"),
        ("notContainsText", "notContains") => Ok("notContains"),
        ("beginsWith", "beginsWith") => Ok("beginsWith"),
        ("endsWith", "endsWith") => Ok("endsWith"),
        _ => Err(Error::Reverse(format!(
            "worksheet `{part}` text conditional rule has incompatible type `{rule_type}` and operator `{operator}`; repair the pair"
        ))),
    }
}

fn validate_external_text_condition(text: &str, part: &str) -> xlsx_jsx::Result<()> {
    if text.chars().count() > 255 || !text.chars().all(is_valid_xml_character) {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` text conditional rule text exceeds 255 XML-valid characters; shorten it or repair invalid characters"
        )));
    }
    if text.contains('"') {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` text conditional rule contains a double quote that the pinned writer cannot safely interpolate into its formula; remove the quote or use an explicit formula rule"
        )));
    }
    Ok(())
}

fn conditional_text_formula(range: &str, operator: &str, text: &str) -> Result<String, String> {
    let Some((first_row, first_column, _, _)) = parse_a1_range(range) else {
        return Err(format!(
            "text conditional range `{range}` cannot be converted to a canonical formula anchor"
        ));
    };
    let anchor = format_a1_cell_reference(first_row, first_column);
    let formula = match operator {
        "containsText" => format!(r#"NOT(ISERROR(SEARCH("{text}",{anchor})))"#),
        "notContains" => format!(r#"ISERROR(SEARCH("{text}",{anchor}))"#),
        "beginsWith" => format!(r#"LEFT({anchor},{})="{text}""#, text.len()),
        "endsWith" => format!(r#"RIGHT({anchor},{})="{text}""#, text.len()),
        _ => {
            return Err(format!(
                "unsupported text conditional operator `{operator}`"
            ));
        }
    };
    Ok(formula)
}

fn conditional_date_formula(range: &str, period: &str) -> Result<String, String> {
    let Some((first_row, first_column, _, _)) = parse_a1_range(range) else {
        return Err(format!(
            "date conditional range `{range}` cannot be converted to a canonical formula anchor"
        ));
    };
    let anchor = format_a1_cell_reference(first_row, first_column);
    let formula = match period {
        "yesterday" => format!("FLOOR({anchor},1)=TODAY()-1"),
        "today" => format!("FLOOR({anchor},1)=TODAY()"),
        "tomorrow" => format!("FLOOR({anchor},1)=TODAY()+1"),
        "last7Days" => format!("AND(TODAY()-FLOOR({anchor},1)<=6,FLOOR({anchor},1)<=TODAY())"),
        "lastWeek" => format!(
            "AND(TODAY()-ROUNDDOWN({anchor},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({anchor},0)<(WEEKDAY(TODAY())+7))"
        ),
        "thisWeek" => format!(
            "AND(TODAY()-ROUNDDOWN({anchor},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({anchor},0)-TODAY()<=7-WEEKDAY(TODAY()))"
        ),
        "nextWeek" => format!(
            "AND(ROUNDDOWN({anchor},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({anchor},0)-TODAY()<(15-WEEKDAY(TODAY())))"
        ),
        "lastMonth" => format!(
            "AND(MONTH({anchor})=MONTH(TODAY())-1,OR(YEAR({anchor})=YEAR(TODAY()),AND(MONTH({anchor})=1,YEAR({anchor})=YEAR(TODAY())-1)))"
        ),
        "thisMonth" => format!("AND(MONTH({anchor})=MONTH(TODAY()),YEAR({anchor})=YEAR(TODAY()))"),
        "nextMonth" => format!(
            "AND(MONTH({anchor})=MONTH(TODAY())+1,OR(YEAR({anchor})=YEAR(TODAY()),AND(MONTH({anchor})=12,YEAR({anchor})=YEAR(TODAY())+1)))"
        ),
        _ => return Err(format!("unsupported date period `{period}`")),
    };
    Ok(formula)
}

fn range_anchor(range: &str, part: &str) -> xlsx_jsx::Result<String> {
    let Some((first_row, first_column, _, _)) = parse_a1_range(range) else {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` conditional range `{range}` cannot be converted to a formula anchor"
        )));
    };
    Ok(format_a1_cell_reference(first_row, first_column))
}

fn parse_optional_conditional_dxf(
    rule: &mut WorksheetConditionalFormat,
    attributes: &HashMap<String, String>,
    differential_formats: &[CellFontStyle],
    part: &str,
    family: &str,
) -> xlsx_jsx::Result<()> {
    let Some(dxf_id) = attributes.get("dxfId") else {
        return Ok(());
    };
    let dxf_id = dxf_id.parse::<usize>().map_err(|error| {
        Error::Reverse(format!(
            "worksheet `{part}` {family} rule has invalid dxfId: {error}"
        ))
    })?;
    rule.format = differential_formats.get(dxf_id).cloned().ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` {family} rule references missing differential style {dxf_id}; repair styles.xml"
        ))
    })?;
    if !rule.format.supports_conditional_dxf() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` {family} rule uses a differential style outside supported font emphasis/decoration/script/RGB font/fill; simplify the dxf or add matching JSX support"
        )));
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Conditional formatting rules share ordered OOXML state and validation."
)]
fn parse_conditional_formats(
    xml: &str,
    part: &str,
    differential_formats: &[CellFontStyle],
) -> xlsx_jsx::Result<Vec<WorksheetConditionalFormat>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut range: Option<String> = None;
    let mut current: Option<WorksheetConditionalFormat> = None;
    let mut thresholds = Vec::new();
    let mut colors = Vec::new();
    let mut formulas: Vec<String> = Vec::new();
    let mut reading_formula = false;
    let mut formula_text = String::new();
    let mut reading_guid = false;
    let mut guid_text = String::new();
    let mut rules = Vec::new();
    loop {
        match reader.read_event() {
            Ok(event @ (Event::Start(_) | Event::Empty(_))) => {
                let is_empty = matches!(&event, Event::Empty(_));
                let (Event::Start(event) | Event::Empty(event)) = event else {
                    continue;
                };
                let tag = event.local_name();
                match tag.as_ref() {
                    b"conditionalFormatting"
                        if event.name().as_ref() == b"conditionalFormatting" =>
                    {
                        let attributes =
                            conditional_attributes(&event, &["sqref"], part, reader.decoder())?;
                        let reference = attributes.get("sqref").ok_or_else(|| {
                            Error::Reverse(format!(
                                "worksheet `{part}` conditionalFormatting is missing sqref"
                            ))
                        })?;
                        if reference.split_whitespace().count() != 1
                            || parse_a1_range(reference).is_none()
                        {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` conditional formatting range `{reference}` is not one supported A1 range; split it into separate rules or use a single range"
                            )));
                        }
                        range = Some(reference.clone());
                    }
                    b"cfRule" if event.name().as_ref() == b"cfRule" => {
                        let attributes = conditional_attributes(
                            &event,
                            &[
                                "type",
                                "priority",
                                "stopIfTrue",
                                "operator",
                                "dxfId",
                                "text",
                                "rank",
                                "percent",
                                "bottom",
                                "aboveAverage",
                                "equalAverage",
                                "stdDev",
                                "timePeriod",
                            ],
                            part,
                            reader.decoder(),
                        )?;
                        let rule_type = attributes
                            .get("type")
                            .map(String::as_str)
                            .unwrap_or_default();
                        if rule_type != "top10"
                            && ["rank", "percent", "bottom"]
                                .iter()
                                .any(|attribute| attributes.contains_key(*attribute))
                        {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` non-top10 conditional rule contains top/bottom metadata; remove rank/percent/bottom or repair its type"
                            )));
                        }
                        if ![
                            "colorScale",
                            "cellIs",
                            "expression",
                            "containsText",
                            "notContainsText",
                            "beginsWith",
                            "endsWith",
                            "dataBar",
                            "iconSet",
                            "top10",
                            "aboveAverage",
                            "timePeriod",
                            "containsBlanks",
                            "notContainsBlanks",
                            "containsErrors",
                            "notContainsErrors",
                            "duplicateValues",
                            "uniqueValues",
                        ]
                        .contains(&rule_type)
                        {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` conditional formatting rule `{rule_type}` is not supported; use color scale, cell comparison, formula/text, top/bottom, average/date, blank/error, duplicate, data bar, or icon-set rules"
                            )));
                        }
                        let priority = attributes.get("priority").ok_or_else(|| {
                            Error::Reverse(format!(
                                "worksheet `{part}` colorScale rule is missing priority"
                            ))
                        })?;
                        let priority = priority.parse::<u32>().ok().filter(|value| *value > 0).ok_or_else(|| Error::Reverse(format!("worksheet `{part}` colorScale priority must be a positive integer")))?;
                        let stop_if_true = attributes
                            .get("stopIfTrue")
                            .map(|value| parse_xml_bool(value, part, "cfRule.stopIfTrue"))
                            .transpose()?
                            .unwrap_or(false);
                        let mut rule = WorksheetConditionalFormat {
                            range: range.clone().unwrap_or_default(),
                            priority: Some(priority),
                            scale_type: match rule_type {
                                "cellIs" => "cellIs",
                                "expression" => "formula",
                                "containsText" | "notContainsText" | "beginsWith" | "endsWith" => {
                                    "text"
                                }
                                "dataBar" => "dataBar",
                                "iconSet" => "iconSet",
                                "top10" => "top10",
                                "aboveAverage" => "average",
                                "timePeriod" => "date",
                                "containsBlanks" | "notContainsBlanks" => "blank",
                                "containsErrors" | "notContainsErrors" => "error",
                                "duplicateValues" | "uniqueValues" => "duplicate",
                                _ => "",
                            }
                            .to_owned(),
                            stop_if_true,
                            ..WorksheetConditionalFormat::default()
                        };
                        let forbidden_metadata: &[&str] = match rule_type {
                            "aboveAverage" => &["timePeriod"],
                            "timePeriod" => &["aboveAverage", "equalAverage", "stdDev"],
                            "containsBlanks" | "notContainsBlanks" | "containsErrors"
                            | "notContainsErrors" | "duplicateValues" | "uniqueValues" => {
                                &["aboveAverage", "equalAverage", "stdDev", "timePeriod"]
                            }
                            _ => &[],
                        };
                        if let Some(attribute) = forbidden_metadata
                            .iter()
                            .find(|attribute| attributes.contains_key(**attribute))
                        {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` `{rule_type}` rule contains incompatible `{attribute}` metadata; remove it or repair the conditional rule"
                            )));
                        }
                        if rule_type == "dataBar" {
                            rule.data_bar = Some(DataBarProperties {
                                minimum: ColorScaleThreshold {
                                    kind: "auto".to_owned(),
                                    ..Default::default()
                                },
                                maximum: ColorScaleThreshold {
                                    kind: "auto".to_owned(),
                                    ..Default::default()
                                },
                                border: Some(true),
                                direction: Some("context".to_owned()),
                                axis_position: Some("automatic".to_owned()),
                                ..Default::default()
                            });
                        } else if rule_type == "iconSet" {
                            rule.icon_set = Some(IconSetProperties {
                                icon_type: "threeTrafficLights".to_owned(),
                                ..Default::default()
                            });
                        } else if rule_type == "top10" {
                            if attributes.contains_key("operator")
                                || attributes.contains_key("text")
                            {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` top10 rule contains text/cell comparison attributes; remove them or repair the rule"
                                )));
                            }
                            let rank = attributes
                                .get("rank")
                                .and_then(|value| value.parse::<u16>().ok())
                                .filter(|rank| *rank > 0)
                                .ok_or_else(|| Error::Reverse(format!(
                                    "worksheet `{part}` top10 rule rank must be an integer from 1 through 1000"
                                )))?;
                            let is_percent = attributes
                                .get("percent")
                                .map(|value| parse_xml_bool(value, part, "top10.percent"))
                                .transpose()?
                                .unwrap_or(false);
                            let is_bottom = attributes
                                .get("bottom")
                                .map(|value| parse_xml_bool(value, part, "top10.bottom"))
                                .transpose()?
                                .unwrap_or(false);
                            let rule_name = match (is_percent, is_bottom) {
                                (false, false) => "top",
                                (false, true) => "bottom",
                                (true, false) => "topPercent",
                                (true, true) => "bottomPercent",
                            };
                            if rank > if is_percent { 100 } else { 1000 } {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` top10 rule rank {rank} is outside the valid 1–{} range; repair percent/rank metadata",
                                    if is_percent { 100 } else { 1000 }
                                )));
                            }
                            rule.top_rule = Some(rule_name.to_owned());
                            rule.rank = Some(rank);
                            if let Some(dxf_id) = attributes.get("dxfId") {
                                let dxf_id = dxf_id.parse::<usize>().map_err(|error| {
                                    Error::Reverse(format!(
                                        "worksheet `{part}` top10 rule has invalid dxfId: {error}"
                                    ))
                                })?;
                                rule.format = differential_formats.get(dxf_id).cloned().ok_or_else(|| Error::Reverse(format!(
                                    "worksheet `{part}` top10 rule references missing differential style {dxf_id}; repair styles.xml"
                                )))?;
                                if !rule.format.supports_conditional_dxf() {
                                    return Err(Error::Reverse(format!(
                                        "worksheet `{part}` top10 rule uses a differential style outside bold/italic/RGB font/fill; simplify the dxf or add matching JSX support"
                                    )));
                                }
                            }
                        } else if rule_type == "aboveAverage" {
                            let above = attributes
                                .get("aboveAverage")
                                .map(|value| parse_xml_bool(value, part, "cfRule.aboveAverage"))
                                .transpose()?
                                .unwrap_or(true);
                            let equal = attributes
                                .get("equalAverage")
                                .map(|value| parse_xml_bool(value, part, "cfRule.equalAverage"))
                                .transpose()?
                                .unwrap_or(false);
                            let standard_deviations = attributes
                                .get("stdDev")
                                .map(|value| value.parse::<u8>())
                                .transpose()
                                .map_err(|error| Error::Reverse(format!(
                                    "worksheet `{part}` average rule stdDev must be 1, 2, or 3: {error}"
                                )))?
                                .unwrap_or(0);
                            let average_rule = match (above, equal, standard_deviations) {
                                (true, false, 0) => "aboveAverage",
                                (false, false, 0) => "belowAverage",
                                (true, true, 0) => "equalOrAboveAverage",
                                (false, true, 0) => "equalOrBelowAverage",
                                (true, false, 1) => "oneStdDevAbove",
                                (false, false, 1) => "oneStdDevBelow",
                                (true, false, 2) => "twoStdDevAbove",
                                (false, false, 2) => "twoStdDevBelow",
                                (true, false, 3) => "threeStdDevAbove",
                                (false, false, 3) => "threeStdDevBelow",
                                _ => {
                                    return Err(Error::Reverse(format!(
                                        "worksheet `{part}` average rule has contradictory aboveAverage/equalAverage/stdDev attributes; use a canonical average rule"
                                    )));
                                }
                            };
                            rule.top_rule = Some(average_rule.to_owned());
                            parse_optional_conditional_dxf(
                                &mut rule,
                                &attributes,
                                differential_formats,
                                part,
                                "average",
                            )?;
                        } else if matches!(
                            rule_type,
                            "timePeriod"
                                | "containsBlanks"
                                | "notContainsBlanks"
                                | "containsErrors"
                                | "notContainsErrors"
                                | "duplicateValues"
                                | "uniqueValues"
                        ) {
                            let semantic_rule = match rule_type {
                                "timePeriod" => attributes.get("timePeriod").cloned().ok_or_else(|| Error::Reverse(format!(
                                    "worksheet `{part}` date rule is missing timePeriod; repair its metadata"
                                )))?,
                                "containsBlanks" => "blank".to_owned(),
                                "notContainsBlanks" => "nonBlank".to_owned(),
                                "containsErrors" => "error".to_owned(),
                                "notContainsErrors" => "nonError".to_owned(),
                                "duplicateValues" => "duplicate".to_owned(),
                                "uniqueValues" => "unique".to_owned(),
                                _ => {
                                    return Err(Error::Reverse(format!(
                                        "worksheet `{part}` has invalid conditional rule type `{rule_type}`"
                                    )));
                                }
                            };
                            if rule.scale_type == "date"
                                && ![
                                    "yesterday",
                                    "today",
                                    "tomorrow",
                                    "last7Days",
                                    "lastWeek",
                                    "thisWeek",
                                    "nextWeek",
                                    "lastMonth",
                                    "thisMonth",
                                    "nextMonth",
                                ]
                                .contains(&semantic_rule.as_str())
                            {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` timePeriod `{semantic_rule}` is unsupported; use a standard Excel date period"
                                )));
                            }
                            rule.top_rule = Some(semantic_rule);
                            let family = rule.scale_type.clone();
                            parse_optional_conditional_dxf(
                                &mut rule,
                                &attributes,
                                differential_formats,
                                part,
                                &family,
                            )?;
                        } else if rule_type == "cellIs" {
                            let operator =
                                attributes.get("operator").cloned().ok_or_else(|| {
                                    Error::Reverse(format!(
                                        "worksheet `{part}` cellIs rule is missing operator"
                                    ))
                                })?;
                            if ![
                                "equal",
                                "notEqual",
                                "greaterThan",
                                "greaterThanOrEqual",
                                "lessThan",
                                "lessThanOrEqual",
                                "between",
                                "notBetween",
                            ]
                            .contains(&operator.as_str())
                            {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` cellIs operator `{operator}` is unsupported; use a standard equality or comparison operator"
                                )));
                            }
                            rule.operator = Some(operator);
                            let dxf_id = attributes.get("dxfId").ok_or_else(|| Error::Reverse(format!(
                                "worksheet `{part}` cellIs rule has no dxf style; only style-bearing comparison rules can be reversed"
                            )))?.parse::<usize>().map_err(|error| Error::Reverse(format!(
                                "worksheet `{part}` cellIs rule has invalid dxfId: {error}"
                            )))?;
                            rule.format = differential_formats.get(dxf_id).cloned().ok_or_else(|| Error::Reverse(format!(
                                "worksheet `{part}` cellIs rule references missing differential style {dxf_id}; repair styles.xml"
                            )))?;
                            if !rule.format.supports_conditional_dxf() {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` cellIs rule uses a differential style outside bold/italic/RGB font/fill (fontColor={:?}, fillColor={:?}, bold={}, italic={}); simplify the dxf or add matching JSX support",
                                    rule.format.font_color,
                                    rule.format.fill_color,
                                    rule.format.bold,
                                    rule.format.italic
                                )));
                            }
                        } else if rule_type == "expression" {
                            if attributes.contains_key("operator") {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` formula conditional format cannot contain operator; remove it from cfRule"
                                )));
                            }
                            let dxf_id = attributes.get("dxfId").ok_or_else(|| Error::Reverse(format!(
                                "worksheet `{part}` formula rule has no dxf style; only style-bearing formula rules can be reversed"
                            )))?.parse::<usize>().map_err(|error| Error::Reverse(format!(
                                "worksheet `{part}` formula rule has invalid dxfId: {error}"
                            )))?;
                            rule.format = differential_formats.get(dxf_id).cloned().ok_or_else(|| Error::Reverse(format!(
                                "worksheet `{part}` formula rule references missing differential style {dxf_id}; repair styles.xml"
                            )))?;
                            if !rule.format.supports_conditional_dxf() {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` formula rule uses a differential style outside bold/italic/RGB font/fill (fontColor={:?}, fillColor={:?}, bold={}, italic={}); simplify the dxf or add matching JSX support",
                                    rule.format.font_color,
                                    rule.format.fill_color,
                                    rule.format.bold,
                                    rule.format.italic
                                )));
                            }
                        } else if rule_type != "colorScale" {
                            let ooxml_operator = attributes
                                .get("operator")
                                .map(String::as_str)
                                .unwrap_or_default();
                            let operator = conditional_text_operator_from_xml(
                                rule_type,
                                ooxml_operator,
                                part,
                            )?;
                            let text = attributes.get("text").cloned().ok_or_else(|| {
                                Error::Reverse(format!(
                                    "worksheet `{part}` text conditional rule `{rule_type}` is missing its text attribute; repair the rule"
                                ))
                            })?;
                            validate_external_text_condition(&text, part)?;
                            rule.operator = Some(operator.to_owned());
                            rule.text = Some(text);
                            let dxf_id = attributes.get("dxfId").ok_or_else(|| Error::Reverse(format!(
                                "worksheet `{part}` text rule has no dxf style; only style-bearing text rules can be reversed"
                            )))?.parse::<usize>().map_err(|error| Error::Reverse(format!(
                                "worksheet `{part}` text rule has invalid dxfId: {error}"
                            )))?;
                            rule.format = differential_formats.get(dxf_id).cloned().ok_or_else(|| Error::Reverse(format!(
                                "worksheet `{part}` text rule references missing differential style {dxf_id}; repair styles.xml"
                            )))?;
                            if !rule.format.supports_conditional_dxf() {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` text rule uses a differential style outside bold/italic/RGB font/fill; simplify the dxf or add matching JSX support"
                                )));
                            }
                        }
                        if range.is_none() || current.is_some() {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` has malformed nested conditional formatting rules"
                            )));
                        }
                        current = Some(rule);
                        thresholds.clear();
                        colors.clear();
                        formulas.clear();
                        rules.push((priority, current.clone().unwrap_or_default()));
                        if is_empty
                            && matches!(
                                rule_type,
                                "top10" | "aboveAverage" | "duplicateValues" | "uniqueValues"
                            )
                        {
                            current = None;
                        }
                    }
                    b"cfvo" if event.name().as_ref() == b"cfvo" => {
                        let attributes = conditional_attributes(
                            &event,
                            &["type", "val", "gte"],
                            part,
                            reader.decoder(),
                        )?;
                        let kind = match attributes.get("type").map(String::as_str).unwrap_or_default() {
                            "min" => "min", "max" => "max", "num" => "number", "percent" => "percent", "percentile" => "percentile", "formula" => "formula",
                            unsupported => return Err(Error::Reverse(format!("worksheet `{part}` color scale threshold type `{unsupported}` cannot be represented; use min/max/num/percent/percentile/formula"))),
                        }.to_owned();
                        let value = attributes
                            .get("val")
                            .cloned()
                            .filter(|_| kind != "min" && kind != "max");
                        if kind != "min" && kind != "max" && value.is_none() {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` color scale `{kind}` threshold is missing val; repair the rule"
                            )));
                        }
                        thresholds.push((kind, value));
                        if let Some(rule) =
                            current.as_mut().filter(|rule| rule.scale_type == "iconSet")
                        {
                            let greater_than =
                                attributes.get("gte").is_some_and(|value| value == "0");
                            rule.icon_set
                                .as_mut()
                                .unwrap()
                                .icons
                                .push(IconThresholdProperties {
                                    kind: thresholds.last().unwrap().0.clone(),
                                    value: thresholds.last().unwrap().1.clone(),
                                    greater_than,
                                    ..Default::default()
                                });
                        }
                    }
                    b"iconSet" if event.name().as_ref() == b"iconSet" => {
                        let attributes = conditional_attributes(
                            &event,
                            &["iconSet", "showValue", "reverse"],
                            part,
                            reader.decoder(),
                        )?;
                        let icon_set = current
                            .as_mut()
                            .and_then(|rule| rule.icon_set.as_mut())
                            .ok_or_else(|| {
                                Error::Reverse(format!(
                                    "worksheet `{part}` iconSet element is outside an iconSet rule"
                                ))
                            })?;
                        if let Some(value) = attributes.get("iconSet") {
                            icon_set.icon_type = conditional_icon_type_from_xml(value, part)?;
                        }
                        if let Some(value) = attributes.get("showValue") {
                            icon_set.show_icons_only =
                                !parse_xml_bool(value, part, "iconSet.showValue")?;
                        }
                        if let Some(value) = attributes.get("reverse") {
                            icon_set.reverse = parse_xml_bool(value, part, "iconSet.reverse")?;
                        }
                    }
                    b"color" if event.name().as_ref() == b"color" => {
                        let attributes =
                            conditional_attributes(&event, &["rgb"], part, reader.decoder())?;
                        let rgb = attributes.get("rgb").ok_or_else(|| {
                            Error::Reverse(format!(
                                "worksheet `{part}` color scale color is missing rgb"
                            ))
                        })?;
                        if rgb.len() != 8
                            || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit())
                            || !rgb[..2].eq_ignore_ascii_case("FF")
                        {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` color scale RGB `{rgb}` is not an opaque ARGB color; use alpha FF"
                            )));
                        }
                        colors.push(format!("#{}", &rgb[2..]).to_ascii_uppercase());
                    }
                    b"dataBar" if event.name().as_ref() == b"dataBar" => {
                        let attributes = conditional_attributes(
                            &event,
                            &["minLength", "maxLength", "showValue"],
                            part,
                            reader.decoder(),
                        )?;
                        let show_value = attributes
                            .get("showValue")
                            .map(|value| parse_xml_bool(value, part, "dataBar.showValue"))
                            .transpose()?
                            .unwrap_or(true);
                        if let Some(bar) = current.as_mut().and_then(|rule| rule.data_bar.as_mut())
                        {
                            bar.bar_only = Some(!show_value);
                        }
                    }
                    b"id" if event.name().as_ref() == b"x14:id" => {
                        reading_guid = true;
                        guid_text.clear();
                    }
                    b"formula"
                        if current.as_ref().is_some_and(|rule| {
                            matches!(
                                rule.scale_type.as_str(),
                                "cellIs" | "formula" | "text" | "date" | "blank" | "error"
                            )
                        }) =>
                    {
                        conditional_attributes(&event, &[], part, reader.decoder())?;
                        reading_formula = true;
                        formula_text.clear();
                    }
                    _ if current.as_ref().is_some_and(|rule| {
                        ["average", "date", "blank", "error", "duplicate"]
                            .contains(&rule.scale_type.as_str())
                    }) =>
                    {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` conditional rule contains unsupported child `{}`; remove it or add explicit reader support",
                            String::from_utf8_lossy(tag.as_ref())
                        )));
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(text)) if reading_formula => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                formula_text.push_str(&unescaped);
            }
            Ok(Event::GeneralRef(reference)) if reading_formula => {
                let value = if let Some(character) =
                    reference.resolve_char_ref().map_err(|error| {
                        Error::Reverse(format!("invalid formula entity in `{part}`: {error}"))
                    })? {
                    character.to_string()
                } else {
                    let entity = reference.decode().map_err(|error| {
                        Error::Reverse(format!("invalid formula entity in `{part}`: {error}"))
                    })?;
                    quick_xml::escape::resolve_predefined_entity(&entity)
                        .ok_or_else(|| Error::Reverse(format!(
                            "worksheet `{part}` conditional formula contains unsupported XML entity `&{entity};`"
                        )))?
                        .to_owned()
                };
                formula_text.push_str(&value);
            }
            Ok(Event::Text(text)) if reading_guid => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                guid_text.push_str(
                    &quick_xml::escape::unescape(&decoded)
                        .map_err(|error| Error::Reverse(error.to_string()))?,
                );
            }
            Ok(Event::End(event)) if event.name().as_ref() == b"cfRule" => {
                let Some(mut rule) = current.take() else {
                    continue;
                };
                if rule.scale_type == "dataBar" {
                    let bar = rule.data_bar.as_mut().ok_or_else(|| {
                        Error::Reverse(format!(
                            "worksheet `{part}` dataBar rule is missing its body"
                        ))
                    })?;
                    if thresholds.len() != 2 || colors.len() != 1 {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` dataBar must contain two endpoint thresholds and one RGB color"
                        )));
                    }
                    for ((kind, value), is_minimum) in thresholds.drain(..).zip([true, false]) {
                        let threshold = ColorScaleThreshold {
                            kind,
                            value,
                            color: String::new(),
                        };
                        if is_minimum {
                            bar.minimum = threshold;
                        } else {
                            bar.maximum = threshold;
                        }
                    }
                    bar.fill_color = colors.pop();
                    if let Some(entry) = rules.last_mut() {
                        entry.1 = rule;
                    }
                    continue;
                }
                if rule.scale_type == "iconSet" {
                    if rule
                        .icon_set
                        .as_ref()
                        .is_none_or(|icons| icons.icons.is_empty())
                    {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` iconSet has no thresholds; repair the rule"
                        )));
                    }
                    let priority = rules.last().map_or(0, |entry| entry.0);
                    if let Some(entry) = rules.last_mut() {
                        entry.1 = rule;
                        entry.0 = priority;
                    }
                    thresholds.clear();
                    continue;
                }
                if rule.scale_type == "top10" {
                    if !thresholds.is_empty() || !colors.is_empty() || !formulas.is_empty() {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` top10 rule contains unexpected child elements; retain only a self-closing cfRule"
                        )));
                    }
                    let priority = rules.last().map_or(0, |entry| entry.0);
                    if let Some(entry) = rules.last_mut() {
                        entry.1 = rule;
                        entry.0 = priority;
                    }
                    continue;
                }
                if ["average", "date", "blank", "error", "duplicate"]
                    .contains(&rule.scale_type.as_str())
                {
                    if !thresholds.is_empty() || !colors.is_empty() {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` {} conditional rule contains unexpected color-scale children; remove them",
                            rule.scale_type
                        )));
                    }
                    let period = rule.top_rule.as_deref().unwrap_or_default();
                    let expected_formula = match rule.scale_type.as_str() {
                        "date" => Some(conditional_date_formula(&rule.range, period).map_err(
                            |error| Error::Reverse(format!("worksheet `{part}` {error}")),
                        )?),
                        "blank" => {
                            let anchor = range_anchor(&rule.range, part)?;
                            Some(if period == "nonBlank" {
                                format!("LEN(TRIM({anchor}))>0")
                            } else {
                                format!("LEN(TRIM({anchor}))=0")
                            })
                        }
                        "error" => {
                            let anchor = range_anchor(&rule.range, part)?;
                            Some(if period == "nonError" {
                                format!("NOT(ISERROR({anchor}))")
                            } else {
                                format!("ISERROR({anchor})")
                            })
                        }
                        _ => None,
                    };
                    if let Some(expected_formula) = expected_formula {
                        if formulas.len() != 1 || formulas[0] != expected_formula {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` {} conditional rule formula {:?} does not match expected formula {:?} for its rule and range; restore the canonical writer formula or repair the rule",
                                rule.scale_type, formulas, expected_formula,
                            )));
                        }
                    } else if !formulas.is_empty() {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` {} conditional rule must not contain formula children; remove them",
                            rule.scale_type
                        )));
                    }
                    if let Some(entry) = rules.last_mut() {
                        entry.1 = rule;
                    }
                    formulas.clear();
                    continue;
                }
                if rule.scale_type == "text" {
                    if formulas.len() != 1 || formulas[0].is_empty() {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` text conditional rule must contain exactly one non-empty formula element; repair its formula"
                        )));
                    }
                    let operator = rule.operator.as_deref().unwrap_or_default();
                    let text = rule.text.as_deref().unwrap_or_default();
                    let expected_formula = conditional_text_formula(&rule.range, operator, text)
                        .map_err(|error| Error::Reverse(format!("worksheet `{part}` {error}")))?;
                    if formulas[0] != expected_formula {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` text conditional rule formula does not match its operator/text/range; restore the writer-generated formula or use a supported formula conditional rule"
                        )));
                    }
                    if !rule.format.has_conditional_dxf_style() {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` text rule has no representable differential style; add bold/italic/RGB font/fill or remove the rule"
                        )));
                    }
                    if let Some(entry) = rules.last_mut() {
                        entry.1 = rule;
                    }
                    continue;
                }
                if rule.scale_type == "formula" {
                    if formulas.len() != 1 || formulas[0].is_empty() {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` formula conditional rule must contain exactly one non-empty formula element; repair its formula"
                        )));
                    }
                    if !rule.format.has_conditional_dxf_style() {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` formula rule has no representable differential style; add bold/italic/RGB font/fill or remove the rule"
                        )));
                    }
                    rule.formula1 = formulas.first().cloned();
                    if let Some(entry) = rules.last_mut() {
                        entry.1 = rule;
                    }
                    continue;
                }
                if rule.scale_type == "cellIs" {
                    let operator = rule.operator.as_deref().unwrap_or_default();
                    if !rule.format.has_conditional_dxf_style() {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` cellIs rule has no representable differential style; add bold/italic/RGB font/fill or remove the rule"
                        )));
                    }
                    let expected_count = if ["between", "notBetween"].contains(&operator) {
                        2
                    } else {
                        1
                    };
                    if formulas.len() != expected_count || formulas.iter().any(String::is_empty) {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` cellIs rule `{operator}` must contain {expected_count} non-empty formula element(s); repair the comparison bounds"
                        )));
                    }
                    rule.formula1 = formulas.first().cloned();
                    rule.formula2 = formulas.get(1).cloned();
                    let priority = rules.last().map_or(0, |entry| entry.0);
                    if let Some(entry) = rules.last_mut() {
                        entry.1 = rule;
                        entry.0 = priority;
                    }
                    continue;
                }
                if !(thresholds.len() == 2 || thresholds.len() == 3)
                    || thresholds.len() != colors.len()
                {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` colorScale must contain matching sets of two or three thresholds and colors"
                    )));
                }
                if thresholds.len() == 2 {
                    "colorScale2"
                } else {
                    "colorScale3"
                }
                .clone_into(&mut rule.scale_type);
                let mut entries = thresholds
                    .drain(..)
                    .zip(colors.drain(..))
                    .map(|((kind, value), color)| ColorScaleThreshold { kind, value, color });
                rule.minimum = entries.next().unwrap_or_default();
                if rule.scale_type == "colorScale3" {
                    rule.midpoint = entries.next();
                }
                rule.maximum = entries.next().unwrap_or_default();
                let priority = rules.last().map_or(0, |entry| entry.0);
                if let Some(entry) = rules.last_mut() {
                    entry.1 = rule;
                    entry.0 = priority;
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"formula" => {
                if reading_formula {
                    reading_formula = false;
                    formulas.push(std::mem::take(&mut formula_text));
                }
            }
            Ok(Event::End(event)) if event.name().as_ref() == b"x14:id" => {
                reading_guid = false;
                if let Some(rule) = current.as_mut() {
                    rule.guid = Some(std::mem::take(&mut guid_text));
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"conditionalFormatting" => {
                range = None;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid XML in `{part}` conditional formatting: {error}"
                )));
            }
        }
    }
    rules.sort_by_key(|entry| entry.0);
    Ok(rules.into_iter().map(|(_, rule)| rule).collect())
}

#[expect(
    clippy::too_many_lines,
    reason = "The x14 reader validates each writer-emitted element and GUID pairing in one ordered pass."
)]
fn merge_x14_data_bars(
    xml: &str,
    part: &str,
    rules: &mut [WorksheetConditionalFormat],
) -> xlsx_jsx::Result<()> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut sqref = String::new();
    let mut reading_sqref = false;
    let mut reading_formula = false;
    let mut formula = String::new();
    let mut extension_uri: Option<String> = None;
    let mut current: Option<(String, DataBarProperties, Vec<ColorScaleThreshold>)> = None;
    let mut extended = HashMap::<String, (String, DataBarProperties)>::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event)) => {
                let name = event.name();
                match name.as_ref() {
                    b"ext" => {
                        for attribute in event.attributes() {
                            let attribute = attribute.map_err(|error| {
                                Error::Reverse(format!(
                                    "worksheet `{part}` has invalid extension metadata: {error}"
                                ))
                            })?;
                            if ![
                                b"uri".as_slice(),
                                b"xmlns:x14".as_slice(),
                                b"xmlns:xm".as_slice(),
                            ]
                            .contains(&attribute.key.as_ref())
                            {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` conditional-format extension has unsupported `{}` metadata",
                                    String::from_utf8_lossy(attribute.key.as_ref())
                                )));
                            }
                            let value = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| Error::Reverse(error.to_string()))?
                                .into_owned();
                            match attribute.key.as_ref() {
                                b"uri" => {
                                    if ![
                                        "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}",
                                        "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
                                    ]
                                    .contains(&value.as_str())
                                    {
                                        return Err(Error::Reverse(format!(
                                            "worksheet `{part}` uses unsupported conditional-format extension `{value}`; remove the extension or add explicit parsing and reverse support"
                                        )));
                                    }
                                    extension_uri = Some(value);
                                }
                                b"xmlns:x14"
                                    if value
                                        == "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" =>
                                    {}
                                b"xmlns:xm"
                                    if value
                                        == "http://schemas.microsoft.com/office/excel/2006/main" => {
                                }
                                b"xmlns:x14" | b"xmlns:xm" => {
                                    return Err(Error::Reverse(format!(
                                        "worksheet `{part}` dataBar extension uses an unexpected namespace `{value}`"
                                    )));
                                }
                                _ => {}
                            }
                        }
                    }
                    b"x14:conditionalFormatting" => {
                        if extension_uri.as_deref()
                            != Some("{78C0D931-6437-407d-A8EE-F0AAD7539E65}")
                        {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` x14 conditional-format extension has an unknown URI; remove it or use a supported writer"
                            )));
                        }
                        let attributes =
                            conditional_attributes(&event, &["xm"], part, reader.decoder())?;
                        if attributes.get("xm").is_some_and(|namespace| {
                            namespace != "http://schemas.microsoft.com/office/excel/2006/main"
                        }) {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` x14 dataBar uses an unexpected xm namespace"
                            )));
                        }
                        sqref.clear();
                    }
                    b"x14:cfRule" => {
                        let attributes = conditional_attributes(
                            &event,
                            &["type", "id", "priority"],
                            part,
                            reader.decoder(),
                        )?;
                        if attributes.get("type").map(String::as_str) == Some("iconSet") {
                            current = None;
                            continue;
                        }
                        if attributes.get("type").map(String::as_str) != Some("dataBar") {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` x14 conditional rule is not a dataBar; add matching JSX and reader support"
                            )));
                        }
                        let id = attributes
                            .get("id")
                            .filter(|id| !id.is_empty())
                            .cloned()
                            .ok_or_else(|| {
                                Error::Reverse(format!(
                                    "worksheet `{part}` x14 dataBar rule is missing its pairing id"
                                ))
                            })?;
                        current = Some((
                            id,
                            DataBarProperties {
                                minimum: ColorScaleThreshold {
                                    kind: "auto".to_owned(),
                                    ..Default::default()
                                },
                                maximum: ColorScaleThreshold {
                                    kind: "auto".to_owned(),
                                    ..Default::default()
                                },
                                border: Some(true),
                                direction: Some("context".to_owned()),
                                axis_position: Some("automatic".to_owned()),
                                ..Default::default()
                            },
                            Vec::new(),
                        ));
                    }
                    b"x14:dataBar" => {
                        let attributes = conditional_attributes(
                            &event,
                            &[
                                "minLength",
                                "maxLength",
                                "border",
                                "gradient",
                                "direction",
                                "negativeBarColorSameAsPositive",
                                "negativeBarBorderColorSameAsPositive",
                                "axisPosition",
                            ],
                            part,
                            reader.decoder(),
                        )?;
                        if attributes
                            .get("minLength")
                            .is_some_and(|value| value != "0")
                            || attributes
                                .get("maxLength")
                                .is_some_and(|value| value != "100")
                        {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` x14 dataBar has non-default minLength/maxLength values not exposed by the DSL; restore 0/100 or extend the JSX API"
                            )));
                        }
                        if let Some((_, bar, _)) = current.as_mut() {
                            bar.border = Some(
                                attributes
                                    .get("border")
                                    .map(|value| parse_xml_bool(value, part, "x14:dataBar.border"))
                                    .transpose()?
                                    .unwrap_or(true),
                            );
                            bar.solid_fill = Some(
                                attributes
                                    .get("gradient")
                                    .map(|value| {
                                        parse_xml_bool(value, part, "x14:dataBar.gradient")
                                    })
                                    .transpose()?
                                    .is_some_and(|gradient| !gradient),
                            );
                            bar.direction = Some(match attributes.get("direction").map_or("context", String::as_str) {
                                "context" => "context", "leftToRight" => "leftToRight", "rightToLeft" => "rightToLeft",
                                value => return Err(Error::Reverse(format!("worksheet `{part}` x14 dataBar direction `{value}` is unsupported"))),
                            }.to_owned());
                            bar.axis_position = Some(match attributes.get("axisPosition").map_or("automatic", String::as_str) {
                                "middle" => "midpoint", "none" => "none", "automatic" => "automatic",
                                value => return Err(Error::Reverse(format!("worksheet `{part}` x14 dataBar axisPosition `{value}` is unsupported"))),
                            }.to_owned());
                            if attributes
                                .get("negativeBarColorSameAsPositive")
                                .map(|value| {
                                    parse_xml_bool(
                                        value,
                                        part,
                                        "x14:dataBar.negativeBarColorSameAsPositive",
                                    )
                                })
                                .transpose()?
                                .unwrap_or(false)
                            {
                                bar.negative_fill_color = bar.fill_color.clone();
                            }
                            if let Some(value) =
                                attributes.get("negativeBarBorderColorSameAsPositive")
                            {
                                let same_as_positive = parse_xml_bool(
                                    value,
                                    part,
                                    "x14:dataBar.negativeBarBorderColorSameAsPositive",
                                )?;
                                if same_as_positive {
                                    bar.negative_border_color = bar.border_color.clone();
                                }
                            }
                        }
                    }
                    b"x14:cfvo" => {
                        let attributes =
                            conditional_attributes(&event, &["type"], part, reader.decoder())?;
                        let kind = match attributes
                            .get("type")
                            .map(String::as_str)
                            .unwrap_or_default()
                        {
                            "autoMin" | "autoMax" => "auto",
                            "min" => "min",
                            "max" => "max",
                            "num" => "number",
                            "percent" => "percent",
                            "percentile" => "percentile",
                            "formula" => "formula",
                            value => {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` x14 dataBar threshold type `{value}` cannot be represented; repair the threshold"
                                )));
                            }
                        };
                        if let Some((_, _, thresholds)) = current.as_mut() {
                            thresholds.push(ColorScaleThreshold {
                                kind: kind.to_owned(),
                                value: None,
                                color: String::new(),
                            });
                        }
                    }
                    b"xm:f" => {
                        reading_formula = true;
                        formula.clear();
                    }
                    b"xm:sqref" => {
                        reading_sqref = true;
                        sqref.clear();
                    }
                    b"x14:fillColor"
                    | b"x14:borderColor"
                    | b"x14:negativeFillColor"
                    | b"x14:negativeBorderColor"
                    | b"x14:axisColor" => {
                        let tag = String::from_utf8_lossy(name.as_ref()).to_string();
                        let attributes =
                            conditional_attributes(&event, &["rgb"], part, reader.decoder())?;
                        let color = attributes.get("rgb").ok_or_else(|| {
                            Error::Reverse(format!("worksheet `{part}` `{tag}` is missing rgb"))
                        })?;
                        let parsed = parse_conditional_rgb(color, part)?;
                        if let Some((_, bar, _)) = current.as_mut() {
                            match tag.as_str() {
                                "x14:fillColor" => bar.fill_color = Some(parsed),
                                "x14:borderColor" => bar.border_color = Some(parsed),
                                "x14:negativeFillColor" => bar.negative_fill_color = Some(parsed),
                                "x14:negativeBorderColor" => {
                                    bar.negative_border_color = Some(parsed);
                                }
                                _ => bar.axis_color = Some(parsed),
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(text)) if reading_formula => {
                formula.push_str(
                    &text
                        .decode()
                        .map_err(|error| Error::Reverse(error.to_string()))?,
                );
            }
            Ok(Event::Text(text)) if reading_sqref => {
                sqref.push_str(
                    &text
                        .decode()
                        .map_err(|error| Error::Reverse(error.to_string()))?,
                );
            }
            Ok(Event::End(event)) => match event.name().as_ref() {
                b"xm:f" => {
                    reading_formula = false;
                    let Some((_, _, thresholds)) = current.as_mut() else {
                        continue;
                    };
                    let threshold = thresholds.last_mut().ok_or_else(|| {
                        Error::Reverse(format!("worksheet `{part}` x14 formula is outside cfvo"))
                    })?;
                    threshold.value = Some(std::mem::take(&mut formula));
                }
                b"xm:sqref" => {
                    reading_sqref = false;
                }
                b"x14:cfRule" => {
                    let Some((id, mut bar, thresholds)) = current.take() else {
                        continue;
                    };
                    if thresholds.len() != 2 {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` x14 dataBar `{id}` must have exactly two thresholds"
                        )));
                    }
                    bar.minimum = thresholds[0].clone();
                    bar.maximum = thresholds[1].clone();
                    if let Some((_, prior)) = extended.insert(id.clone(), (String::new(), bar)) {
                        let _ = prior;
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` repeats x14 dataBar pairing id `{id}`; repair duplicate IDs"
                        )));
                    }
                }
                b"x14:conditionalFormatting" => {
                    let Some(parsed_range) = parse_a1_range(sqref.trim()) else {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` x14 dataBar sqref `{}` is not one supported A1 range",
                            sqref.trim()
                        )));
                    };
                    let canonical_range = format_a1_range(parsed_range);
                    let ids: Vec<_> = extended
                        .iter()
                        .filter(|(_, (stored_range, _))| stored_range.is_empty())
                        .map(|(id, _)| id.clone())
                        .collect();
                    for id in ids {
                        if let Some((stored_range, _)) = extended.get_mut(&id) {
                            stored_range.clone_from(&canonical_range);
                        }
                    }
                }
                b"ext" => {
                    extension_uri = None;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid x14 conditional-format XML in `{part}`: {error}"
                )));
            }
        }
    }
    let data_bar_count = rules
        .iter()
        .filter(|rule| rule.scale_type == "dataBar")
        .count();
    let mut paired_ids = std::collections::HashSet::new();
    for rule in rules.iter_mut().filter(|rule| rule.scale_type == "dataBar") {
        let Some(id) = rule.guid.as_deref() else {
            if extended.is_empty() {
                continue;
            }
            return Err(Error::Reverse(format!(
                "worksheet `{part}` classic dataBar at `{}` has no x14 pairing id; repair its extension metadata",
                rule.range
            )));
        };
        let Some((extension_range, extension_bar)) = extended.get(id) else {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` classic dataBar id `{id}` has no matching x14 rule; restore its x14 extension"
            )));
        };
        if extension_range != &rule.range {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` classic and x14 dataBar id `{id}` ranges differ (`{}` versus `{extension_range}`); repair the paired rule",
                rule.range
            )));
        }
        let classic_bar = rule.data_bar.as_ref().ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{part}` classic dataBar `{id}` has no parsed style body; repair the rule"
            ))
        })?;
        if !data_bar_thresholds_match(&classic_bar.minimum, &extension_bar.minimum, "min")
            || !data_bar_thresholds_match(&classic_bar.maximum, &extension_bar.maximum, "max")
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` classic and x14 dataBar id `{id}` thresholds disagree; restore matching endpoint types and values"
            )));
        }
        if !paired_ids.insert(id.to_owned()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has duplicate classic dataBar id `{id}`; repair duplicate pairing IDs"
            )));
        }
        let classic_bar = rule.data_bar.take().unwrap_or_default();
        let mut merged_bar = extension_bar.clone();
        if merged_bar.fill_color.is_none() {
            merged_bar.fill_color = classic_bar.fill_color;
        }
        if merged_bar.bar_only.is_none() {
            merged_bar.bar_only = classic_bar.bar_only;
        }
        rule.data_bar = Some(merged_bar);
    }
    if paired_ids.len() != extended.len()
        || data_bar_count != paired_ids.len() && !extended.is_empty()
    {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` contains orphan x14 dataBar rules or IDs; restore one-to-one classic/x14 rule pairs"
        )));
    }
    Ok(())
}

#[derive(Default)]
struct X14IconSetRule {
    id: String,
    priority: u32,
    range: String,
    icon_type: String,
    custom: bool,
    thresholds: Vec<IconThresholdProperties>,
    icons: Vec<IconThresholdProperties>,
}

#[expect(
    clippy::too_many_lines,
    reason = "x14 icon rules require ordered parsing, pairing, and strict semantic cross-checking."
)]
fn merge_x14_icon_sets(
    xml: &str,
    part: &str,
    rules: &mut Vec<WorksheetConditionalFormat>,
) -> xlsx_jsx::Result<()> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut range_text = String::new();
    let mut reading_range = false;
    let mut reading_formula = false;
    let mut formula = String::new();
    let mut current: Option<X14IconSetRule> = None;
    let mut pending = Vec::<X14IconSetRule>::new();
    let mut ids = std::collections::HashSet::new();
    let mut in_supported_extension = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event)) => match event.name().as_ref() {
                b"ext" => {
                    let attributes = conditional_attributes(
                        &event,
                        &["uri", "x14", "xm"],
                        part,
                        reader.decoder(),
                    )?;
                    in_supported_extension = attributes
                        .get("uri")
                        .is_some_and(|uri| uri == "{78C0D931-6437-407d-A8EE-F0AAD7539E65}");
                    for (prefix, expected) in [
                        (
                            "x14",
                            "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                        ),
                        ("xm", "http://schemas.microsoft.com/office/excel/2006/main"),
                    ] {
                        if attributes
                            .get(prefix)
                            .is_some_and(|namespace| namespace != expected)
                        {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` conditional-format extension uses an unexpected `{prefix}` namespace"
                            )));
                        }
                    }
                }
                b"x14:conditionalFormatting" if in_supported_extension => {
                    let attributes =
                        conditional_attributes(&event, &["xm"], part, reader.decoder())?;
                    if attributes.get("xm").is_some_and(|namespace| {
                        namespace != "http://schemas.microsoft.com/office/excel/2006/main"
                    }) {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` x14 iconSet uses an unexpected xm namespace"
                        )));
                    }
                    range_text.clear();
                }
                b"x14:cfRule" if in_supported_extension => {
                    let attributes = conditional_attributes(
                        &event,
                        &["type", "priority", "id"],
                        part,
                        reader.decoder(),
                    )?;
                    match attributes.get("type").map(String::as_str) {
                        Some("dataBar") => current = None,
                        Some("iconSet") => {
                            if current.is_some() {
                                return Err(Error::Reverse(format!(
                                    "worksheet `{part}` has nested x14 conditional-format rules"
                                )));
                            }
                            let id = attributes.get("id").filter(|id| !id.is_empty()).cloned().ok_or_else(|| {
                                Error::Reverse(format!("worksheet `{part}` x14 iconSet rule is missing its pairing id"))
                            })?;
                            let priority = attributes.get("priority").and_then(|value| value.parse::<u32>().ok()).filter(|value| *value > 0).ok_or_else(|| {
                                Error::Reverse(format!("worksheet `{part}` x14 iconSet priority must be a positive integer"))
                            })?;
                            current = Some(X14IconSetRule {
                                id,
                                priority,
                                ..Default::default()
                            });
                        }
                        Some(other) => {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` x14 conditional rule `{other}` is unsupported; remove it or add matching reverse support"
                            )));
                        }
                        None => {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` x14 conditional rule is missing its type"
                            )));
                        }
                    }
                }
                b"x14:iconSet" if in_supported_extension && current.is_some() => {
                    let attributes = conditional_attributes(
                        &event,
                        &["iconSet", "custom"],
                        part,
                        reader.decoder(),
                    )?;
                    let icon_type = attributes
                        .get("iconSet")
                        .map_or("3TrafficLights1", String::as_str);
                    let custom = attributes
                        .get("custom")
                        .map(|value| parse_xml_bool(value, part, "x14:iconSet.custom"))
                        .transpose()?
                        .unwrap_or(false);
                    if let Some(rule) = current.as_mut() {
                        rule.icon_type = conditional_icon_type_from_xml(icon_type, part)?;
                        rule.custom = custom;
                    }
                }
                b"x14:cfvo" if in_supported_extension && current.is_some() => {
                    let attributes =
                        conditional_attributes(&event, &["type"], part, reader.decoder())?;
                    let kind = match attributes
                        .get("type")
                        .map(String::as_str)
                        .unwrap_or_default()
                    {
                        "num" => "number",
                        "percent" => "percent",
                        "percentile" => "percentile",
                        "formula" => "formula",
                        other => {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` x14 iconSet threshold type `{other}` is unsupported"
                            )));
                        }
                    };
                    if let Some(rule) = current.as_mut() {
                        rule.thresholds.push(IconThresholdProperties {
                            kind: kind.to_owned(),
                            ..Default::default()
                        });
                    }
                }
                b"xm:f" if in_supported_extension && current.is_some() => {
                    reading_formula = true;
                    formula.clear();
                }
                b"x14:cfIcon" if in_supported_extension && current.is_some() => {
                    let attributes = conditional_attributes(
                        &event,
                        &["iconSet", "iconId"],
                        part,
                        reader.decoder(),
                    )?;
                    let icon_set = attributes
                        .get("iconSet")
                        .map(String::as_str)
                        .unwrap_or_default();
                    let icon_id = attributes
                        .get("iconId")
                        .and_then(|value| value.parse::<u8>().ok())
                        .ok_or_else(|| {
                            Error::Reverse(format!(
                                "worksheet `{part}` x14 cfIcon has an invalid iconId"
                            ))
                        })?;
                    let icon = if icon_set == "NoIcons" {
                        if icon_id != 0 {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` x14 NoIcons entry has non-default iconId {icon_id}; restore iconId=0 to match the pinned writer"
                            )));
                        }
                        IconThresholdProperties {
                            no_icon: true,
                            ..Default::default()
                        }
                    } else {
                        IconThresholdProperties {
                            icon_type: Some(conditional_icon_type_from_xml(icon_set, part)?),
                            icon_index: Some(icon_id),
                            ..Default::default()
                        }
                    };
                    if let Some(rule) = current.as_mut() {
                        rule.icons.push(icon);
                    }
                }
                b"xm:sqref" if in_supported_extension => {
                    reading_range = true;
                    range_text.clear();
                }
                _ if in_supported_extension && current.is_some() => {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` x14 iconSet contains unsupported child `{}`; preserve only cfvo, cfIcon, and xm:f content",
                        String::from_utf8_lossy(event.name().as_ref())
                    )));
                }
                _ => {}
            },
            Ok(Event::Text(text)) if reading_formula => {
                formula.push_str(
                    &quick_xml::escape::unescape(
                        &text
                            .decode()
                            .map_err(|error| Error::Reverse(error.to_string()))?,
                    )
                    .map_err(|error| Error::Reverse(error.to_string()))?,
                );
            }
            Ok(Event::Text(text)) if reading_range => {
                range_text.push_str(
                    &text
                        .decode()
                        .map_err(|error| Error::Reverse(error.to_string()))?,
                );
            }
            Ok(Event::End(event)) => match event.name().as_ref() {
                b"xm:f" if reading_formula => {
                    reading_formula = false;
                    let threshold = current
                        .as_mut()
                        .and_then(|rule| rule.thresholds.last_mut())
                        .ok_or_else(|| {
                            Error::Reverse(format!(
                                "worksheet `{part}` x14 formula is outside an icon threshold"
                            ))
                        })?;
                    threshold.value = Some(std::mem::take(&mut formula));
                }
                b"xm:sqref" => reading_range = false,
                b"x14:cfRule" => {
                    if let Some(rule) = current.take() {
                        pending.push(rule);
                    }
                }
                b"x14:conditionalFormatting" if in_supported_extension => {
                    let parsed = parse_a1_range(range_text.trim()).ok_or_else(|| Error::Reverse(format!("worksheet `{part}` x14 iconSet sqref `{}` is not one supported A1 range", range_text.trim())))?;
                    let canonical_range = format_a1_range(parsed);
                    for rule in &mut pending {
                        if rule.range.is_empty() {
                            rule.range.clone_from(&canonical_range);
                        }
                    }
                }
                b"ext" => in_supported_extension = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid x14 icon-set XML in `{part}`: {error}"
                )));
            }
        }
    }

    for extension in pending {
        if !ids.insert(extension.id.clone()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` repeats x14 iconSet pairing id `{}`",
                extension.id
            )));
        }
        if extension.icon_type.is_empty() || extension.range.is_empty() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` x14 iconSet `{}` is missing its icon set or sqref",
                extension.id
            )));
        }
        let count = conditional_icon_count(&extension.icon_type);
        if extension.thresholds.len() != count {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` x14 iconSet `{}` has {} thresholds, expected {count}",
                extension.id,
                extension.thresholds.len()
            )));
        }
        if extension.custom && extension.icons.len() != count
            || !extension.custom && !extension.icons.is_empty()
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` x14 iconSet `{}` custom flag and cfIcon count disagree; repair the extension",
                extension.id
            )));
        }
        let matching = rules.iter().position(|rule| {
            rule.priority == Some(extension.priority) && rule.range == extension.range
        });
        let Some(index) = matching else {
            let icon_set = build_x14_icon_set(&extension);
            rules.push(WorksheetConditionalFormat {
                range: extension.range,
                scale_type: "iconSet".to_owned(),
                priority: Some(extension.priority),
                guid: Some(extension.id),
                icon_set: Some(icon_set),
                ..Default::default()
            });
            continue;
        };
        let classic = &mut rules[index];
        if classic.scale_type != "iconSet" {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` classic and x14 iconSet priority {} do not refer to the same rule and range; repair the extension pairing",
                extension.priority
            )));
        }
        let Some(classic_icons) = classic.icon_set.as_ref() else {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` classic iconSet priority {} has no parsed body",
                extension.priority
            )));
        };
        if classic_icons.icon_type != extension.icon_type
            || !icon_thresholds_match(&classic_icons.icons, &extension.thresholds)
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` classic and x14 iconSet `{}` types or thresholds disagree; restore matching rules",
                extension.id
            )));
        }
        if extension.custom {
            let mut merged = extension.thresholds.clone();
            for ((threshold, classic_threshold), custom) in merged
                .iter_mut()
                .zip(&classic_icons.icons)
                .zip(&extension.icons)
            {
                threshold.greater_than = classic_threshold.greater_than;
                threshold.icon_type.clone_from(&custom.icon_type);
                threshold.icon_index = custom.icon_index;
                threshold.no_icon = custom.no_icon;
            }
            if let Some(icon_set) = classic.icon_set.as_mut() {
                icon_set.icons = merged;
            }
        }
        classic.guid = Some(extension.id);
    }
    rules.sort_by_key(|rule| rule.priority.unwrap_or(u32::MAX));
    Ok(())
}

fn conditional_icon_count(icon_type: &str) -> usize {
    if icon_type.starts_with("three") {
        3
    } else if icon_type.starts_with("four") {
        4
    } else {
        5
    }
}

fn build_x14_icon_set(extension: &X14IconSetRule) -> IconSetProperties {
    let mut icons = extension.thresholds.clone();
    if extension.custom {
        for (threshold, custom) in icons.iter_mut().zip(&extension.icons) {
            threshold.icon_type.clone_from(&custom.icon_type);
            threshold.icon_index = custom.icon_index;
            threshold.no_icon = custom.no_icon;
        }
    }
    IconSetProperties {
        icon_type: extension.icon_type.clone(),
        icons,
        ..Default::default()
    }
}

fn icon_thresholds_match(
    classic: &[IconThresholdProperties],
    extended: &[IconThresholdProperties],
) -> bool {
    classic.len() == extended.len()
        && classic
            .iter()
            .zip(extended)
            .all(|(left, right)| left.kind == right.kind && left.value == right.value)
}

fn data_bar_thresholds_match(
    classic: &ColorScaleThreshold,
    extended: &ColorScaleThreshold,
    endpoint: &str,
) -> bool {
    let extended_kind = if extended.kind == "auto" {
        endpoint
    } else {
        extended.kind.as_str()
    };
    if classic.kind != extended_kind {
        return false;
    }
    match (&classic.value, &extended.value) {
        (None, None) => true,
        (Some(left), Some(right)) if classic.kind == "formula" => left == right,
        (Some(left), Some(right)) => match (left.parse::<f64>(), right.parse::<f64>()) {
            (Ok(left), Ok(right)) => left
                .partial_cmp(&right)
                .is_some_and(std::cmp::Ordering::is_eq),
            _ => left == right,
        },
        _ => false,
    }
}

fn parse_conditional_rgb(value: &str, part: &str) -> xlsx_jsx::Result<String> {
    if value.len() != 8
        || !value.bytes().all(|byte| byte.is_ascii_hexdigit())
        || !value[..2].eq_ignore_ascii_case("FF")
    {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` conditional-format color `{value}` is not opaque ARGB; use alpha FF"
        )));
    }
    Ok(format!("#{}", &value[2..]).to_ascii_uppercase())
}

fn conditional_attributes(
    event: &BytesStart<'_>,
    allowed: &[&str],
    part: &str,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<HashMap<String, String>> {
    let mut values = HashMap::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| {
            Error::Reverse(format!(
                "worksheet `{part}` has invalid conditional-format attribute: {error}"
            ))
        })?;
        let key = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        if !allowed.contains(&key.as_str()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` conditional formatting has unsupported `{key}` metadata; remove it or add explicit JSX support"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?
            .into_owned();
        if values.insert(key.clone(), value).is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` conditional formatting repeats `{key}`"
            )));
        }
    }
    Ok(values)
}

fn audit_worksheet_hyperlink_event(
    event: &BytesStart<'_>,
    tag: &str,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    relationships: &HashMap<String, String>,
    used_relationships: &mut std::collections::HashSet<String>,
    hyperlinks: &mut HashMap<String, WorksheetHyperlink>,
) -> xlsx_jsx::Result<()> {
    match tag {
        "hyperlinks" if event.attributes().next().is_some() => Err(Error::Reverse(format!(
            "worksheet `{part}` `<hyperlinks>` has unsupported container metadata; remove it or add explicit support"
        ))),
        "hyperlink" => audit_hyperlink_element(
            event,
            part,
            decoder,
            relationships,
            used_relationships,
            hyperlinks,
        ),
        _ => Ok(()),
    }
}

fn audit_worksheet_element_metadata(
    event: &BytesStart<'_>,
    tag: &str,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    formats: &[CellFontStyle],
    hyperlink_relationships: &HashMap<String, String>,
    metadata: &mut WorksheetMetadata,
) -> xlsx_jsx::Result<()> {
    if tag == "c" {
        audit_cell_style(
            event,
            part,
            decoder,
            formats,
            &mut metadata.cells,
            &mut metadata.cell_style_indices,
        )?;
    }
    if tag == "tabColor" {
        audit_tab_color(event, part, decoder, &mut metadata.tab_color)?;
    }
    if tag == "printOptions" {
        audit_print_options(event, part, decoder, metadata)?;
    }
    if tag == "pageSetup" {
        audit_empty_page_setup(event, part, decoder, metadata)?;
    }
    if tag == "mergeCell" {
        audit_merge_cell(event, part, decoder, &mut metadata.merged_ranges)?;
    }
    audit_worksheet_hyperlink_event(
        event,
        tag,
        part,
        decoder,
        hyperlink_relationships,
        &mut metadata.used_hyperlink_relationships,
        &mut metadata.hyperlinks,
    )?;
    if tag == "mergeCells" {
        audit_merge_cells_container(event, part)?;
    }
    if tag == "col" {
        audit_column_attributes(event, part, decoder, formats, &mut metadata.columns)?;
    }
    if tag == "autoFilter" {
        audit_auto_filter(event, part, decoder, &mut metadata.auto_filter)?;
    }
    if tag == "pane" {
        audit_freeze_pane(event, part, decoder, &mut metadata.freeze_panes)?;
    }
    audit_column_container(event, tag, part)?;
    audit_page_layout_metadata(
        event,
        tag,
        part,
        decoder,
        &mut metadata.default_row_height_pixels,
        &mut metadata.hide_unused_rows,
        &mut metadata.print,
    )?;
    audit_sheet_properties_if_present(event, tag, part, decoder, &mut *metadata)?;
    if tag == "sheetView" {
        audit_sheet_view(event, part, decoder, metadata)?;
    }
    if tag == "selection" {
        metadata
            .pane_selections
            .push(parse_pane_selection(event, part, decoder)?);
    }
    if tag == "tablePart" {
        let mut relation_id = None;
        for attribute in event.attributes().with_checks(false) {
            let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
            if attribute.key.local_name().as_ref() == b"id" {
                relation_id = Some(
                    attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned(),
                );
            }
        }
        let relation_id = relation_id.ok_or_else(|| {
            Error::Reverse(format!("worksheet `{part}` has a tablePart without r:id"))
        })?;
        metadata.table_relationship_ids.push(relation_id);
    }
    Ok(())
}

fn audit_hyperlink_element(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    relationships: &HashMap<String, String>,
    used_relationships: &mut std::collections::HashSet<String>,
    hyperlinks: &mut HashMap<String, WorksheetHyperlink>,
) -> xlsx_jsx::Result<()> {
    let mut reference = None;
    let mut relationship_id = None;
    let mut location = None;
    let mut display = None;
    let mut tooltip = None;
    let mut unsupported_metadata = None;
    let mut seen_attributes = std::collections::HashSet::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?
            .into_owned();
        if !seen_attributes.insert(key.as_ref().to_vec()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` hyperlink repeats `{}`; repair the hyperlink element",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        match key.as_ref() {
            b"ref" => reference = Some(value),
            b"id" => relationship_id = Some(value),
            b"location" => location = Some(value),
            b"display" => display = Some(value),
            b"tooltip" => tooltip = Some(value),
            _ => {
                unsupported_metadata = Some(String::from_utf8_lossy(key.as_ref()).into_owned());
            }
        }
    }
    if let Some(attribute) = unsupported_metadata {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` hyperlink uses unsupported `{attribute}` metadata; remove it or add matching JSX support"
        )));
    }
    let reference = reference.ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` hyperlink is missing its cell `ref`; repair the hyperlink element"
        ))
    })?;
    let (row, column) = parse_a1_cell_reference(&reference).ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` hyperlink range `{reference}` is not one A1 cell; use a single-cell hyperlink"
        ))
    })?;
    let target = if let Some(relationship_id) = relationship_id {
        let mut target = relationships
            .get(&relationship_id)
            .cloned()
            .ok_or_else(|| {
                Error::Reverse(format!(
                    "worksheet `{part}` hyperlink on `{reference}` references missing relationship `{relationship_id}`; repair the relationship part"
                ))
            })?;
        if !used_relationships.insert(relationship_id.clone()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` reuses hyperlink relationship `{relationship_id}`; use one relationship per hyperlink or add relationship sharing support"
            )));
        }
        if let Some(location) = location {
            target.push('#');
            target.push_str(&location);
        }
        target
    } else {
        let location = location.filter(|location| !location.is_empty()).ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{part}` hyperlink on `{reference}` has neither an external relationship nor a non-empty internal location; repair the target"
            ))
        })?;
        format!("internal:{location}")
    };
    let reference = format_a1_cell_reference(row, column);
    if hyperlinks
        .insert(
            reference.clone(),
            WorksheetHyperlink {
                target,
                display,
                tooltip,
            },
        )
        .is_some()
    {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` contains more than one hyperlink on `{reference}`; combine them into one link"
        )));
    }
    Ok(())
}

fn audit_worksheet_row(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    formats: &[CellFontStyle],
    next_row_index: &mut u32,
    seen_rows: &mut std::collections::HashSet<u32>,
    rows: &mut HashMap<u32, RowProperties>,
) -> xlsx_jsx::Result<()> {
    let (row_index, properties) =
        audit_row_attributes(event, part, decoder, formats, *next_row_index)?;
    if !seen_rows.insert(row_index) {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` contains duplicate row index {}; repair the row records",
            row_index + 1
        )));
    }
    if let Some(properties) = properties {
        rows.insert(row_index, properties);
    }
    *next_row_index = row_index.saturating_add(1);
    Ok(())
}

fn audit_column_container(event: &BytesStart<'_>, tag: &str, part: &str) -> xlsx_jsx::Result<()> {
    if tag == "cols" && event.attributes().next().is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` `<cols>` has unsupported container attributes; remove the extra metadata or add explicit column-container support"
        )));
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Array-formula anchors, dynamic cell metadata, and range overlap share one worksheet XML state machine."
)]
fn parse_array_formula_ranges(
    xml: &str,
    part: &str,
    has_dynamic_metadata: bool,
) -> xlsx_jsx::Result<HashMap<String, ArrayFormulaRange>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut current_cell = None;
    let mut current_cell_metadata = None;
    let mut current_cell_has_dynamic_formula = false;
    let mut formulas = HashMap::new();
    loop {
        match reader.read_event() {
            Ok(event @ (Event::Start(_) | Event::Empty(_))) => {
                let is_empty = matches!(&event, Event::Empty(_));
                let (Event::Start(event) | Event::Empty(event)) = &event else {
                    unreachable!("matched an XML start or empty event")
                };
                let tag = event.local_name();
                if tag.as_ref() == b"c" {
                    current_cell = optional_xml_attribute(event, "r", part, reader.decoder())?;
                    current_cell_metadata =
                        optional_xml_attribute(event, "cm", part, reader.decoder())?;
                    current_cell_has_dynamic_formula = false;
                    if let Some(index) = &current_cell_metadata {
                        if index != "1" {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` cell has unsupported metadata index `{index}`; only dynamic-array metadata index 1 is representable"
                            )));
                        }
                        if !has_dynamic_metadata {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` cell uses cm=1 without the workbook dynamic-array metadata part; repair the XLSX metadata relationship"
                            )));
                        }
                        if is_empty {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` has cm=1 on an empty cell; attach dynamic-array metadata only to a dynamic array formula"
                            )));
                        }
                    }
                } else if tag.as_ref() == b"f" && event.attributes().next().is_some() {
                    let mut attributes = HashMap::new();
                    for attribute in event.attributes().with_checks(false) {
                        let attribute =
                            attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|error| Error::Reverse(error.to_string()))?
                            .into_owned();
                        let name = String::from_utf8_lossy(attribute.key.local_name().as_ref())
                            .into_owned();
                        if attributes.insert(name, value).is_some() {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` array formula has duplicate attributes; repair its formula metadata"
                            )));
                        }
                    }
                    let is_dynamic = current_cell_metadata.as_deref() == Some("1");
                    if attributes.get("t").map(String::as_str) != Some("array") {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` contains a shared, data-table, or unsupported formula type; convert it to an ordinary formula or a legacy array formula"
                        )));
                    }
                    if attributes.len() != 2 {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` array formula has unsupported attributes; remove attributes outside `t=array` and `ref`"
                        )));
                    }
                    let reference = attributes.get("ref").ok_or_else(|| Error::Reverse(format!(
                        "worksheet `{part}` array formula is missing `ref`; repair the array range metadata"
                    )))?;
                    let range = parse_array_formula_reference(reference).ok_or_else(|| Error::Reverse(format!(
                        "worksheet `{part}` array formula has invalid or unordered range `{reference}`; repair it to an ordered A1 range"
                    )))?;
                    let anchor = current_cell.as_deref().ok_or_else(|| Error::Reverse(format!(
                        "worksheet `{part}` array formula has no containing cell; repair the cell XML"
                    )))?;
                    let anchor_coordinates = xlsx_jsx::parse_a1_cell_reference(anchor).ok_or_else(|| Error::Reverse(format!(
                        "worksheet `{part}` array formula anchor `{anchor}` is not a valid cell reference; repair the cell XML"
                    )))?;
                    if anchor_coordinates != (range.0, range.1) {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` array formula anchor `{anchor}` is not the top-left of `{reference}`; repair the anchor or range"
                        )));
                    }
                    let anchor = xlsx_jsx::format_a1_cell_reference(range.0, range.1);
                    let canonical_range = xlsx_jsx::format_a1_range(range);
                    current_cell_has_dynamic_formula = is_dynamic;
                    if formulas
                        .insert(
                            anchor.clone(),
                            ArrayFormulaRange {
                                reference: canonical_range,
                                coordinates: range,
                                is_dynamic,
                                has_explicit_range: reference.contains(':'),
                            },
                        )
                        .is_some()
                    {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` contains duplicate array-formula anchors at `{anchor}`; repair the formula groups"
                        )));
                    }
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"c" => {
                if current_cell_metadata.is_some() && !current_cell_has_dynamic_formula {
                    let reference = current_cell.as_deref().unwrap_or("unknown");
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` cell `{reference}` has cm=1 but no dynamic array formula; remove the orphan cell metadata"
                    )));
                }
                current_cell = None;
                current_cell_metadata = None;
                current_cell_has_dynamic_formula = false;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid XML in `{part}`: {error}"))),
        }
    }
    let values = formulas
        .values()
        .map(|formula| formula.coordinates)
        .collect::<Vec<_>>();
    for (index, range) in values.iter().enumerate() {
        if values[..index]
            .iter()
            .any(|prior| ranges_overlap(*range, *prior))
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` contains overlapping array-formula groups; repair the formula ranges so they are disjoint"
            )));
        }
    }
    Ok(formulas)
}

fn optional_xml_attribute(
    event: &BytesStart<'_>,
    name: &str,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<Option<String>> {
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() == name.as_bytes() {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| {
                    Error::Reverse(format!("invalid `{name}` in `{part}`: {error}"))
                })?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

fn parse_array_formula_reference(reference: &str) -> Option<(u32, u16, u32, u16)> {
    xlsx_jsx::parse_a1_range(reference).or_else(|| {
        xlsx_jsx::parse_a1_cell_reference(reference).map(|(row, column)| (row, column, row, column))
    })
}

fn ranges_overlap(a: (u32, u16, u32, u16), b: (u32, u16, u32, u16)) -> bool {
    a.0 <= b.2 && a.2 >= b.0 && a.1 <= b.3 && a.3 >= b.1
}

fn audit_merge_cells_container(event: &BytesStart<'_>, part: &str) -> xlsx_jsx::Result<()> {
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"count" {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` mergeCells container has unsupported metadata; remove it or add explicit merge-container support"
            )));
        }
    }
    Ok(())
}

fn audit_data_validation_container(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    count: &mut Option<usize>,
) -> xlsx_jsx::Result<()> {
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"count" {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` dataValidations container has unsupported metadata; remove it or add matching DSL support"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        *count = Some(value.parse::<usize>().map_err(|error| {
            Error::Reverse(format!(
                "worksheet `{part}` dataValidations has invalid count `{value}`: {error}"
            ))
        })?);
    }
    Ok(())
}

fn parse_data_validation_attributes(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<WorksheetDataValidation> {
    let mut validation = WorksheetDataValidation {
        show_dropdown: true,
        ..Default::default()
    };
    let mut validation_type = None;
    let mut reference = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match key.as_ref() {
            b"type" => {
                validation_type = Some(value.to_string());
                validation.validation_type = value.into_owned();
            }
            b"sqref" => reference = Some(value.into_owned()),
            b"allowBlank" => validation.allow_blank = parse_xml_bool(&value, part, "allowBlank")?,
            b"showDropDown" => {
                validation.show_dropdown = !parse_xml_bool(&value, part, "showDropDown")?;
            }
            b"showInputMessage" => {
                validation.show_input_message = parse_xml_bool(&value, part, "showInputMessage")?;
            }
            b"showErrorMessage" => {
                validation.show_error_message = parse_xml_bool(&value, part, "showErrorMessage")?;
            }
            b"inputTitle" | b"promptTitle" => validation.input_title = Some(value.into_owned()),
            b"prompt" => validation.input_message = Some(value.into_owned()),
            b"errorTitle" => validation.error_title = Some(value.into_owned()),
            b"error" => validation.error_message = Some(value.into_owned()),
            b"errorStyle" => {
                let value = value.into_owned();
                if !["stop", "warning", "information"].contains(&value.as_str()) {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` data validation uses unsupported errorStyle `{value}`; use stop, warning, or information"
                    )));
                }
                if value != "stop" {
                    validation.error_style = Some(value);
                }
            }
            b"operator" => validation.operator = Some(value.into_owned()),
            key => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` data validation has unsupported attribute `{}`; remove it or add matching DSL support",
                    String::from_utf8_lossy(key)
                )));
            }
        }
    }
    let validation_type = validation_type
        .filter(|validation_type| !validation_type.is_empty())
        .unwrap_or_else(|| "any".to_owned());
    validation.validation_type.clone_from(&validation_type);
    if ![
        "any",
        "list",
        "whole",
        "decimal",
        "date",
        "time",
        "textLength",
        "custom",
    ]
    .contains(&validation_type.as_str())
    {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` uses unsupported data validation type `{validation_type}`; use a supported rule or add matching DSL support"
        )));
    }
    let reference = reference.ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` data validation is missing sqref; repair the validation range"
        ))
    })?;
    let parsed = parse_data_validation_range(&reference).ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` data validation range `{reference}` is invalid or contains multiple ranges; use one ordered A1 range"
        ))
    })?;
    validation.range = format_a1_range(parsed);
    Ok(validation)
}

fn validate_external_data_validation(
    validation: &WorksheetDataValidation,
    part: &str,
) -> xlsx_jsx::Result<()> {
    validate_external_validation_rule(validation, part)?;
    for (property, value, maximum) in [
        ("inputTitle", validation.input_title.as_deref(), 32),
        ("errorTitle", validation.error_title.as_deref(), 32),
        ("inputMessage", validation.input_message.as_deref(), 255),
        ("errorMessage", validation.error_message.as_deref(), 255),
    ] {
        if value.is_some_and(|value| value.chars().count() > maximum) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` data validation `{}` has `{property}` longer than {maximum} characters; shorten it before reversing",
                validation.range
            )));
        }
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Validation rules share ordered cross-field checks for a single OOXML element."
)]
fn validate_external_validation_rule(
    validation: &WorksheetDataValidation,
    part: &str,
) -> xlsx_jsx::Result<()> {
    if validation.validation_type == "any" {
        if validation.formula1.is_some()
            || validation.formula2.is_some()
            || validation.operator.is_some()
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` any-value validation on `{}` contains rule formulas or an operator; remove them because type any has no rule payload",
                validation.range
            )));
        }
        if validation.error_style.is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` any-value validation on `{}` has a non-default errorStyle; rust_xlsxwriter omits this setting for type any, so restore `stop` or remove the style",
                validation.range
            )));
        }
        return Ok(());
    }
    let formula1 = validation
        .formula1
        .as_deref()
        .filter(|value| !value.is_empty());
    if validation.validation_type == "list" {
        if validation.operator.is_some() || validation.formula2.is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` list validation on `{}` contains operator/formula2 fields; remove unsupported list-rule metadata",
                validation.range
            )));
        }
        let formula1 = formula1.ok_or_else(|| Error::Reverse(format!(
            "worksheet `{part}` list validation on `{}` is missing formula1; add a list source formula"
        , validation.range)))?;
        if formula1.chars().count() > 255 {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` list validation on `{}` exceeds the 255-character source limit; shorten the source formula",
                validation.range
            )));
        }
        return Ok(());
    }
    let formula1 = formula1.ok_or_else(|| Error::Reverse(format!(
        "worksheet `{part}` {} validation on `{}` is missing formula1; add the first bound or formula",
        validation.validation_type, validation.range
    )))?;
    if formula1.chars().count() > 8192 {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` data validation on `{}` has formula1 longer than 8192 characters; shorten it",
            validation.range
        )));
    }
    if let Some(formula2) = validation.formula2.as_deref()
        && (formula2.is_empty() || formula2.chars().count() > 8192)
    {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` data validation on `{}` has an empty or overlong formula2; repair the second bound",
            validation.range
        )));
    }
    match validation.validation_type.as_str() {
        "custom" => {
            if validation.operator.is_some() || validation.formula2.is_some() {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` custom validation on `{}` supports formula1 only; remove operator/formula2",
                    validation.range
                )));
            }
        }
        "whole" | "decimal" | "date" | "time" | "textLength" => {
            if let Some(operator) = validation.operator.as_deref()
                && ![
                    "between",
                    "notBetween",
                    "equal",
                    "notEqual",
                    "greaterThan",
                    "greaterThanOrEqual",
                    "lessThan",
                    "lessThanOrEqual",
                ]
                .contains(&operator)
            {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` data validation operator `{operator}` is unsupported; use a standard comparison operator"
                )));
            }
            let has_two_bounds = validation.formula2.is_some();
            let is_two_bound_operator = matches!(
                validation.operator.as_deref(),
                Some("between" | "notBetween")
            );
            if is_two_bound_operator && !has_two_bounds {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` data validation on `{}` uses a two-bound operator without formula2; repair the second bound",
                    validation.range
                )));
            }
            if has_two_bounds
                && validation
                    .operator
                    .as_deref()
                    .is_some_and(|operator| !["between", "notBetween"].contains(&operator))
            {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` data validation on `{}` has formula2 with a single-bound operator; repair operator or remove formula2",
                    validation.range
                )));
            }
            if !has_two_bounds && validation.operator.is_none() {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` data validation on `{}` has neither a comparison operator nor a second bound; repair its rule",
                    validation.range
                )));
            }
        }
        _ => {}
    }
    Ok(())
}

fn validate_data_validation_ranges(
    part: &str,
    validations: &[WorksheetDataValidation],
) -> xlsx_jsx::Result<()> {
    for (index, validation) in validations.iter().enumerate() {
        let range = parse_data_validation_range(&validation.range).ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{part}` has an invalid data validation range; repair its sqref"
            ))
        })?;
        if validations[..index].iter().any(|prior| {
            parse_data_validation_range(&prior.range).is_some_and(|prior_range| {
                range.0 <= prior_range.2
                    && range.2 >= prior_range.0
                    && range.1 <= prior_range.3
                    && range.3 >= prior_range.1
            })
        }) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` contains overlapping data validation ranges; make them disjoint or combine them into one rule"
            )));
        }
    }
    Ok(())
}

fn parse_data_validation_range(reference: &str) -> Option<(u32, u16, u32, u16)> {
    parse_a1_range(reference).or_else(|| {
        let (row, column) = parse_a1_cell_reference(reference)?;
        Some((row, column, row, column))
    })
}

fn audit_merge_cell(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    ranges: &mut Vec<MergeRange>,
) -> xlsx_jsx::Result<()> {
    let mut reference = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"ref" {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` mergeCell has unsupported metadata; remove it or add matching JSX support"
            )));
        }
        reference = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?
                .into_owned(),
        );
    }
    let reference = reference.ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` mergeCell is missing its `ref`; repair the merge range"
        ))
    })?;
    let (first_row, first_column, last_row, last_column) =
        parse_a1_range(&reference).ok_or_else(|| {
            Error::Reverse(format!(
                "worksheet `{part}` merged range `{reference}` is invalid or outside Excel limits; repair the range"
            ))
        })?;
    if first_row == last_row && first_column == last_column {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` merged range `{reference}` covers only one cell; remove the invalid merge record"
        )));
    }
    ranges.push(MergeRange {
        first_row,
        first_column,
        last_row,
        last_column,
    });
    Ok(())
}

fn validate_merged_ranges(part: &str, metadata: &WorksheetMetadata) -> xlsx_jsx::Result<()> {
    for formula in metadata.array_formula_ranges.values() {
        if metadata.merged_ranges.iter().any(|merged| {
            ranges_overlap(
                formula.coordinates,
                (
                    merged.first_row,
                    merged.first_column,
                    merged.last_row,
                    merged.last_column,
                ),
            )
        }) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` contains an array formula overlapping a merged range; make the formula and merge ranges disjoint"
            )));
        }
    }
    for (index, range) in metadata.merged_ranges.iter().enumerate() {
        if metadata.merged_ranges[..index].iter().any(|prior| {
            range.first_row <= prior.last_row
                && range.last_row >= prior.first_row
                && range.first_column <= prior.last_column
                && range.last_column >= prior.first_column
        }) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` contains overlapping merged ranges; repair the mergeCell refs so each cell belongs to at most one merge"
            )));
        }
        let base_reference = cell_reference(range.first_row, u32::from(range.first_column));
        let base_style = metadata
            .cell_style_indices
            .get(&base_reference)
            .copied()
            .unwrap_or_default();
        let area = u64::from(range.last_row - range.first_row + 1)
            * u64::from(range.last_column - range.first_column + 1);
        let mut matching_cells = 0_u64;
        for (reference, style_index) in &metadata.cell_style_indices {
            let Some((row, column)) = parse_a1_cell_reference(reference) else {
                continue;
            };
            if row >= range.first_row
                && row <= range.last_row
                && column >= range.first_column
                && column <= range.last_column
            {
                if *style_index != base_style {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` merged range {}:{} has different cell styles; normalize the merge to one style because JSX applies the top-left style to the entire range",
                        base_reference,
                        cell_reference(range.last_row, u32::from(range.last_column))
                    )));
                }
                matching_cells += 1;
            }
        }
        if base_style != 0 && matching_cells != area {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` merged range {}:{} has unstyled or missing interior cells; make every merged cell use the top-left style before reversing",
                base_reference,
                cell_reference(range.last_row, u32::from(range.last_column))
            )));
        }
    }
    Ok(())
}

#[allow(clippy::too_many_lines)] // Attribute validation is deliberately kept adjacent to its whitelist.
fn audit_column_attributes(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    formats: &[CellFontStyle],
    columns: &mut BTreeMap<u16, ColumnProperties>,
) -> xlsx_jsx::Result<()> {
    let mut min = None;
    let mut max = None;
    let mut width = None;
    let mut hidden = false;
    let mut outline_level = 0_u8;
    let mut collapsed = false;
    let mut collapsed_seen = false;
    let mut style_index = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match key.as_ref() {
            b"min" => min = Some(parse_column_endpoint(&value, "min", part)?),
            b"max" => max = Some(parse_column_endpoint(&value, "max", part)?),
            b"width" => {
                width = Some(value.parse::<f64>().map_err(|error| {
                    Error::Reverse(format!(
                        "worksheet `{part}` has invalid column width `{value}`: {error}"
                    ))
                })?);
            }
            b"hidden" => hidden = parse_xml_bool(&value, part, "hidden")?,
            b"collapsed" => {
                if collapsed_seen {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` column range repeats `collapsed`; remove the duplicate attribute"
                    )));
                }
                collapsed_seen = true;
                collapsed = parse_xml_bool(&value, part, "collapsed")?;
            }
            b"outlineLevel" => {
                outline_level = value.parse::<u8>().map_err(|error| {
                    Error::Reverse(format!(
                        "worksheet `{part}` has invalid column outline level `{value}`: {error}"
                    ))
                })?;
                if outline_level > 7 {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` column range {min:?}–{max:?} has outline level {outline_level} above Excel's maximum of 7; lower the level or repair the column metadata"
                    )));
                }
            }
            b"customWidth" => {
                let _ = parse_xml_bool(&value, part, "customWidth")?;
            }
            b"style" => {
                style_index = Some(value.parse::<usize>().map_err(|error| {
                    Error::Reverse(format!(
                        "worksheet `{part}` has invalid column style index `{value}`: {error}"
                    ))
                })?);
            }
            _ => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` column range contains unsupported `{}` metadata; remove best-fit or other unsupported column settings, or add matching DSL support",
                    String::from_utf8_lossy(key.as_ref())
                )));
            }
        }
    }
    let min = min.ok_or_else(|| {
        Error::Reverse(format!("worksheet `{part}` column range is missing `min`"))
    })?;
    let max = max.ok_or_else(|| {
        Error::Reverse(format!("worksheet `{part}` column range is missing `max`"))
    })?;
    if min > max {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` column range min {min} exceeds max {max}; repair the `<col>` range"
        )));
    }
    let style_index = validate_default_style_index(
        style_index.unwrap_or_default(),
        None,
        formats,
        part,
        &format!("column range {min}–{max}"),
    )?;
    if collapsed && min != max {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` collapsed column marker spans columns {min}–{max}; split it into one summary column or add multi-marker range support"
        )));
    }
    let width_pixels = match width {
        None if hidden => None,
        None => None,
        Some(width) if !width.is_finite() || !(0.0..=256.0).contains(&width) => {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` column range {min}–{max} has width `{width}` outside the supported writer range (0, 256]"
            )));
        }
        Some(0.0) => {
            hidden = true;
            None
        }
        Some(width) => {
            let pixels = (width * 7.0).round();
            let serialized = (pixels * 256.0 / 7.0).floor() / 256.0;
            if !(1.0..=1790.0).contains(&pixels) || (serialized - width).abs() > 1e-9 {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` column range {min}–{max} has width `{width}` that cannot be represented exactly as integer pixels by the writer; normalize the width to a writer-generated pixel width or add exact OOXML width support"
                )));
            }
            // The range check above bounds this rounded positive value to 1..=1790.
            #[expect(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
            let pixels = pixels as u32;
            Some(pixels)
        }
    };
    let properties = ColumnProperties {
        width_pixels,
        hidden,
        outline_level,
        collapsed,
        style_index,
    };
    if properties == ColumnProperties::default() {
        return Ok(());
    }
    for column in min..=max {
        let index = u16::try_from(column - 1)
            .map_err(|error| Error::Reverse(format!("invalid column index: {error}")))?;
        if columns.insert(index, properties).is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` contains overlapping `<col>` ranges at column {column}; merge or separate the column ranges"
            )));
        }
    }
    Ok(())
}

fn audit_auto_filter(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    auto_filter: &mut Option<String>,
) -> xlsx_jsx::Result<()> {
    if auto_filter.is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` contains more than one autoFilter; retain a single plain filter range"
        )));
    }
    let mut reference = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match attribute.key.local_name().as_ref() {
            b"ref" => reference = Some(value.into_owned()),
            _ => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` autoFilter has unsupported `{}` metadata; remove filter settings or add matching DSL support",
                    String::from_utf8_lossy(attribute.key.local_name().as_ref())
                )));
            }
        }
    }
    let reference = reference.ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` autoFilter is missing its `ref` range; repair the worksheet metadata"
        ))
    })?;
    let parsed = parse_a1_range(&reference).ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` autoFilter range `{reference}` is invalid or outside Excel limits; repair the range"
        ))
    })?;
    *auto_filter = Some(format_a1_range(parsed));
    Ok(())
}

fn parse_column_endpoint(value: &str, attribute: &str, part: &str) -> xlsx_jsx::Result<u32> {
    let endpoint = value.parse::<u32>().map_err(|error| {
        Error::Reverse(format!(
            "worksheet `{part}` has invalid column `{attribute}` value `{value}`: {error}"
        ))
    })?;
    if !(1..=16_384).contains(&endpoint) {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` column `{attribute}` value `{endpoint}` is outside 1–16384"
        )));
    }
    Ok(endpoint)
}

fn validate_default_style_index(
    style_index: usize,
    custom_format: Option<bool>,
    formats: &[CellFontStyle],
    part: &str,
    location: &str,
) -> xlsx_jsx::Result<Option<usize>> {
    if custom_format == Some(false) && style_index != 0 {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` {location} has a style index but customFormat is false; repair the contradictory style metadata"
        )));
    }
    if style_index == 0 {
        return Ok(None);
    }
    let style = formats.get(style_index).ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` {location} references missing style index {style_index}; repair the XLSX style table"
        ))
    })?;
    if style.support != StyleSupport::Supported {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` {location} uses style index {style_index} with formatting outside the lossless cell-format subset; simplify that style or add matching row/column format support"
        )));
    }
    Ok(Some(style_index))
}

fn audit_cell_style(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    formats: &[CellFontStyle],
    cell_styles: &mut HashMap<String, CellFontStyle>,
    cell_style_indices: &mut HashMap<String, usize>,
) -> xlsx_jsx::Result<()> {
    let (reference, style_index) = audit_cell_attributes(event, part, decoder)?;
    cell_style_indices.insert(reference.clone(), style_index);
    let style = formats.get(style_index).cloned().ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` cell `{reference}` references missing style index {style_index}; repair the XLSX style table"
        ))
    })?;
    if style.support != StyleSupport::Supported {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` cell `{reference}` uses a style outside the bold/italic/underline/strikethrough/script/font-name/size/family/charset/RGB-font-color/four-edge/diagonal-border/solid-fill/alignment/indent/wrapText/shrinkToFit/number-format subset; unsupported metadata may include font sizes outside the OOXML range, patterned or gradient fill, unsupported built-in formats, theme/indexed font colors, theme/indexed border colors, or unmodeled border metadata. Remove the unsupported feature or add matching DSL support"
        )));
    }
    if style.bold
        || style.italic
        || style.quote_prefix.is_some()
        || style.strikethrough == Some(true)
        || style.underline.is_some()
        || style.script.is_some()
        || style.wrap_text == Some(true)
        || style.shrink_to_fit == Some(true)
        || style.indent.is_some()
        || style.fill_color.is_some()
        || style.fill_pattern.is_some()
        || style.fill_foreground_color.is_some()
        || style.fill_background_color.is_some()
        || style.font_color.is_some()
        || style.font_name.is_some()
        || style.font_size.is_some()
        || style.font_family.is_some()
        || style.font_charset.is_some()
        || style.borders.has_any()
        || style.number_format.is_some()
        || style.horizontal_align.is_some()
        || style.vertical_align.is_some()
        || style.text_rotation.is_some()
        || style.reading_direction.is_some()
        || style.locked == Some(false)
        || style.formula_hidden
    {
        cell_styles.insert(reference, style);
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Worksheet view auditing validates coupled zoom attributes as one lossless state."
)]
fn audit_sheet_view(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    metadata: &mut WorksheetMetadata,
) -> xlsx_jsx::Result<()> {
    if metadata.sheet_view_seen {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has multiple sheetView elements; combine the view settings or add multi-view support"
        )));
    }
    metadata.sheet_view_seen = true;
    let mut gridlines_seen = false;
    let mut right_to_left_seen = false;
    let mut seen = std::collections::HashSet::new();
    let mut normal_zoom = None;
    let mut normal_zoom_scale = None;
    let mut page_layout_zoom = None;
    let mut page_break_zoom = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        let allowed = match key.as_ref() {
            b"workbookViewId" => Some("0"),
            b"tabSelected" => {
                if metadata.tab_selected.is_some() {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` sheetView repeats `tabSelected`; remove the duplicate attribute"
                    )));
                }
                metadata.tab_selected =
                    Some(parse_xml_bool(&value, part, "sheetView.tabSelected")?);
                None
            }
            b"showGridLines" => {
                if gridlines_seen {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` sheetView repeats `showGridLines`; remove the duplicate attribute"
                    )));
                }
                gridlines_seen = true;
                let gridlines = parse_xml_bool(&value, part, "sheetView.showGridLines")?;
                metadata.hide_screen_gridlines = !gridlines;
                None
            }
            b"rightToLeft" => {
                metadata.right_to_left =
                    parse_sheet_view_right_to_left(&value, part, &mut right_to_left_seen)?
                        .then_some(());
                None
            }
            b"topLeftCell" => {
                let (row, column) = parse_a1_cell_reference(&value).ok_or_else(|| Error::Reverse(format!(
                    "worksheet `{part}` sheetView topLeftCell `{value}` is invalid; repair it to one Excel cell reference"
                )))?;
                let cell = format_a1_cell_reference(row, column);
                if cell != "A1" {
                    metadata.top_left_cell = Some(cell);
                }
                None
            }
            b"view" => {
                metadata.view = match value.as_ref() {
                    "normal" => WorksheetView::Normal,
                    "pageLayout" => WorksheetView::PageLayout,
                    "pageBreakPreview" => WorksheetView::PageBreakPreview,
                    _ => {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` has unsupported sheetView@view `{value}`; use normal, pageLayout, or pageBreakPreview"
                        )));
                    }
                };
                None
            }
            b"zoomScale" => {
                normal_zoom = Some(parse_sheet_view_zoom(&value, part, "zoomScale")?);
                None
            }
            b"zoomScaleNormal" => {
                normal_zoom_scale = Some(parse_sheet_view_zoom(&value, part, "zoomScaleNormal")?);
                None
            }
            b"zoomScalePageLayoutView" => {
                page_layout_zoom = Some(parse_sheet_view_zoom(
                    &value,
                    part,
                    "zoomScalePageLayoutView",
                )?);
                None
            }
            b"zoomScaleSheetLayoutView" => {
                page_break_zoom = Some(parse_sheet_view_zoom(
                    &value,
                    part,
                    "zoomScaleSheetLayoutView",
                )?);
                None
            }
            _ => None,
        };
        if matches!(
            key.as_ref(),
            b"view"
                | b"tabSelected"
                | b"topLeftCell"
                | b"zoomScale"
                | b"zoomScaleNormal"
                | b"zoomScalePageLayoutView"
                | b"zoomScaleSheetLayoutView"
        ) && !seen.insert(key.as_ref().to_vec())
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` sheetView repeats `{}`; remove the duplicate attribute",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        if !matches!(
            key.as_ref(),
            b"showGridLines"
                | b"rightToLeft"
                | b"tabSelected"
                | b"topLeftCell"
                | b"view"
                | b"zoomScale"
                | b"zoomScaleNormal"
                | b"zoomScalePageLayoutView"
                | b"zoomScaleSheetLayoutView"
        ) && allowed != Some(value.as_ref())
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has non-default worksheet view state `{}`; restore default view settings or add worksheet-view support",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
    }
    if let Some(normal) = normal_zoom_scale
        && normal_zoom != Some(normal)
        && !(normal_zoom.is_none() && normal == 100)
    {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has `zoomScaleNormal` value {normal} that differs from `zoomScale` {normal_zoom:?}; this DSL preserves a single zoom value, so make them equal or set both to the same percentage"
        )));
    }
    let active_zoom = match metadata.view {
        WorksheetView::Normal => {
            reject_inactive_zoom(
                page_layout_zoom,
                part,
                "zoomScalePageLayoutView",
                "pageLayout",
            )?;
            reject_inactive_zoom(
                page_break_zoom,
                part,
                "zoomScaleSheetLayoutView",
                "pageBreakPreview",
            )?;
            normal_zoom
        }
        WorksheetView::PageLayout => {
            reject_inactive_zoom(normal_zoom_scale, part, "zoomScaleNormal", "normal")?;
            reject_inactive_zoom(
                page_break_zoom,
                part,
                "zoomScaleSheetLayoutView",
                "pageBreakPreview",
            )?;
            require_matching_view_zoom(normal_zoom, page_layout_zoom, part, "pageLayout")?;
            page_layout_zoom.or(normal_zoom)
        }
        WorksheetView::PageBreakPreview => {
            reject_inactive_zoom(normal_zoom_scale, part, "zoomScaleNormal", "normal")?;
            reject_inactive_zoom(
                page_layout_zoom,
                part,
                "zoomScalePageLayoutView",
                "pageLayout",
            )?;
            require_matching_view_zoom(normal_zoom, page_break_zoom, part, "pageBreakPreview")?;
            page_break_zoom.or(normal_zoom)
        }
    };
    metadata.zoom = active_zoom.filter(|zoom| *zoom != 100);
    Ok(())
}

#[derive(Clone, Copy)]
enum PageBreakAxis {
    Horizontal,
    Vertical,
}

#[derive(Default)]
struct PageBreakAudit {
    seen: bool,
    expected: Option<(usize, usize)>,
    values: Vec<u32>,
}

fn parse_worksheet_protection(
    xml: &str,
    part: &str,
) -> xlsx_jsx::Result<Option<BTreeMap<String, bool>>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut result = None;
    let mut in_protection = false;
    let mut sheet_protection_enabled = false;
    let mut attributes_seen = std::collections::HashSet::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if event.local_name().as_ref() == b"sheetProtection" => {
                if result.is_some() {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` has duplicate sheetProtection elements; keep exactly one"
                    )));
                }
                result = Some(BTreeMap::new());
                in_protection = true;
                parse_worksheet_protection_attributes(
                    &event,
                    part,
                    reader.decoder(),
                    result
                        .as_mut()
                        .ok_or_else(|| Error::Reverse("missing sheetProtection state".into()))?,
                    &mut sheet_protection_enabled,
                    &mut attributes_seen,
                )?;
            }
            Ok(Event::Empty(event)) if event.local_name().as_ref() == b"sheetProtection" => {
                if result.is_some() {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` has duplicate sheetProtection elements; keep exactly one"
                    )));
                }
                result = Some(BTreeMap::new());
                parse_worksheet_protection_attributes(
                    &event,
                    part,
                    reader.decoder(),
                    result
                        .as_mut()
                        .ok_or_else(|| Error::Reverse("missing sheetProtection state".into()))?,
                    &mut sheet_protection_enabled,
                    &mut attributes_seen,
                )?;
            }
            Ok(Event::Start(event)) if in_protection => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` sheetProtection has nested `{}` content; remove it or add matching protection support",
                    String::from_utf8_lossy(event.local_name().as_ref())
                )));
            }
            Ok(Event::Empty(event)) if in_protection => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` sheetProtection has nested `{}` content; remove it or add matching protection support",
                    String::from_utf8_lossy(event.local_name().as_ref())
                )));
            }
            Ok(Event::Text(event)) if in_protection => {
                let value = event.decode().map_err(|error| {
                    Error::Reverse(format!("invalid sheetProtection text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` sheetProtection contains text content; remove it to preserve the protection settings"
                    )));
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"sheetProtection" => {
                in_protection = false;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid worksheet XML in `{part}` while reading protection: {error}"
                )));
            }
            _ => {}
        }
    }
    if result.is_some() && !sheet_protection_enabled {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` sheetProtection does not enable sheet protection; only passwordless `sheet=1` protection can be represented, so remove it or add explicit disabled-protection support"
        )));
    }
    Ok(result)
}

fn parse_worksheet_protection_attributes(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    result: &mut BTreeMap<String, bool>,
    sheet_enabled: &mut bool,
    seen: &mut std::collections::HashSet<String>,
) -> xlsx_jsx::Result<()> {
    const PERMISSIONS: &[(&[u8], &str, bool)] = &[
        (b"objects", "editObjects", false),
        (b"scenarios", "editScenarios", false),
        (b"formatCells", "formatCells", false),
        (b"formatColumns", "formatColumns", false),
        (b"formatRows", "formatRows", false),
        (b"insertColumns", "insertColumns", false),
        (b"insertRows", "insertRows", false),
        (b"insertHyperlinks", "insertLinks", false),
        (b"deleteColumns", "deleteColumns", false),
        (b"deleteRows", "deleteRows", false),
        (b"selectLockedCells", "selectLockedCells", true),
        (b"sort", "sort", false),
        (b"autoFilter", "useAutoFilter", false),
        (b"pivotTables", "usePivotTables", false),
        (b"selectUnlockedCells", "selectUnlockedCells", true),
    ];
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let name = attribute.key.local_name();
        let name_string = String::from_utf8_lossy(name.as_ref()).into_owned();
        if !seen.insert(name_string.clone()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` sheetProtection repeats `{name_string}`; remove the duplicate"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        if name.as_ref() == b"password"
            || name.as_ref() == b"algorithmName"
            || name.as_ref() == b"hashValue"
            || name.as_ref() == b"saltValue"
            || name.as_ref() == b"spinCount"
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` uses password-based sheet protection (`{name_string}`); the original password cannot be recovered, so remove the password or add an explicit password-hash representation"
            )));
        }
        if name.as_ref() == b"sheet" {
            *sheet_enabled = parse_xml_bool(&value, part, "sheetProtection.sheet")?;
            continue;
        }
        let Some((_, property, default)) = PERMISSIONS
            .iter()
            .find(|(attribute_name, _, _)| *attribute_name == name.as_ref())
        else {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` sheetProtection has unsupported attribute `{name_string}`; remove it or add matching worksheet-protection support"
            )));
        };
        let allowed = !parse_xml_bool(&value, part, &format!("sheetProtection.{name_string}"))?;
        if allowed != *default {
            result.insert((*property).to_owned(), allowed);
        }
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Header/footer text, entity, and unsupported-variant checks share one ordered XML pass."
)]
fn parse_worksheet_header_footer(xml: &str, part: &str) -> xlsx_jsx::Result<WorksheetHeaderFooter> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut result = WorksheetHeaderFooter::default();
    let mut parent_seen = false;
    let mut in_parent = false;
    let mut active_child: Option<&'static str> = None;
    let mut header_seen = false;
    let mut footer_seen = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if event.local_name().as_ref() == b"headerFooter" => {
                if parent_seen {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` has duplicate headerFooter elements; keep exactly one"
                    )));
                }
                parent_seen = true;
                in_parent = true;
                parse_header_footer_attributes(&event, part, reader.decoder(), &mut result)?;
            }
            Ok(Event::Empty(event)) if event.local_name().as_ref() == b"headerFooter" => {
                if parent_seen {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` has duplicate headerFooter elements; keep exactly one"
                    )));
                }
                parent_seen = true;
                parse_header_footer_attributes(&event, part, reader.decoder(), &mut result)?;
            }
            Ok(Event::Start(event)) if in_parent => {
                let child = match event.local_name().as_ref() {
                    b"oddHeader" => {
                        if header_seen {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` has duplicate oddHeader elements; keep one header"
                            )));
                        }
                        header_seen = true;
                        "header"
                    }
                    b"oddFooter" => {
                        if footer_seen {
                            return Err(Error::Reverse(format!(
                                "worksheet `{part}` has duplicate oddFooter elements; keep one footer"
                            )));
                        }
                        footer_seen = true;
                        "footer"
                    }
                    _ => return Err(unsupported_header_footer_element(part, &event)),
                };
                ensure_header_footer_child_has_no_attributes(&event, part)?;
                active_child = Some(child);
            }
            Ok(Event::Empty(event)) if in_parent => match event.local_name().as_ref() {
                b"oddHeader" => {
                    ensure_header_footer_child_has_no_attributes(&event, part)?;
                    if header_seen {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` has duplicate oddHeader elements; keep one header"
                        )));
                    }
                    header_seen = true;
                }
                b"oddFooter" => {
                    ensure_header_footer_child_has_no_attributes(&event, part)?;
                    if footer_seen {
                        return Err(Error::Reverse(format!(
                            "worksheet `{part}` has duplicate oddFooter elements; keep one footer"
                        )));
                    }
                    footer_seen = true;
                }
                _ => return Err(unsupported_header_footer_element(part, &event)),
            },
            Ok(Event::Text(event)) if in_parent && active_child.is_some() => {
                let decoded = event.decode().map_err(|error| {
                    Error::Reverse(format!("invalid header/footer text in `{part}`: {error}"))
                })?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                    Error::Reverse(format!(
                        "invalid escaped header/footer text in `{part}`: {error}"
                    ))
                })?;
                let target = if active_child == Some("header") {
                    &mut result.header
                } else {
                    &mut result.footer
                };
                target.get_or_insert_with(String::new).push_str(&decoded);
            }
            Ok(Event::CData(event)) if in_parent && active_child.is_some() => {
                let decoded = event.decode().map_err(|error| {
                    Error::Reverse(format!("invalid header/footer CDATA in `{part}`: {error}"))
                })?;
                let target = if active_child == Some("header") {
                    &mut result.header
                } else {
                    &mut result.footer
                };
                target.get_or_insert_with(String::new).push_str(&decoded);
            }
            Ok(Event::GeneralRef(event)) if in_parent && active_child.is_some() => {
                let value = if let Some(character) = event.resolve_char_ref().map_err(|error| {
                    Error::Reverse(format!(
                        "invalid header/footer character reference in `{part}`: {error}"
                    ))
                })? {
                    character.to_string()
                } else {
                    let entity = event.decode().map_err(|error| {
                        Error::Reverse(format!("invalid header/footer entity in `{part}`: {error}"))
                    })?;
                    quick_xml::escape::resolve_predefined_entity(&entity)
                        .ok_or_else(|| Error::Reverse(format!(
                            "worksheet `{part}` header/footer contains unsupported XML entity `&{entity};`; replace it with literal text"
                        )))?
                        .to_owned()
                };
                let target = if active_child == Some("header") {
                    &mut result.header
                } else {
                    &mut result.footer
                };
                target.get_or_insert_with(String::new).push_str(&value);
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"oddHeader" => {
                active_child = None;
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"oddFooter" => {
                active_child = None;
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"headerFooter" => {
                in_parent = false;
                active_child = None;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid worksheet XML in `{part}` while reading header/footer: {error}"
                )));
            }
            _ => {}
        }
    }
    for (property, value) in [("header", &result.header), ("footer", &result.footer)] {
        if let Some(value) = value {
            if value.chars().count() > 255 {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` {property} exceeds Excel's 255-character limit; shorten the source header/footer before reversing"
                )));
            }
            if contains_header_footer_picture(value) {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` {property} contains a picture placeholder but header/footer images are not supported; remove the image or add image relationship support"
                )));
            }
        }
    }
    Ok(result)
}

fn parse_header_footer_attributes(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    result: &mut WorksheetHeaderFooter,
) -> xlsx_jsx::Result<()> {
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let name = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match name.as_ref() {
            b"scaleWithDoc" => {
                result.scale_with_doc = Some(parse_xml_bool(&value, part, "scaleWithDoc")?);
            }
            b"alignWithMargins" => {
                result.align_with_page = Some(parse_xml_bool(&value, part, "alignWithMargins")?);
            }
            _ => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` headerFooter has unsupported attribute `{}`; remove it or add matching DSL support",
                    String::from_utf8_lossy(name.as_ref())
                )));
            }
        }
    }
    Ok(())
}

fn unsupported_header_footer_element(part: &str, event: &BytesStart<'_>) -> xlsx_jsx::Error {
    Error::Reverse(format!(
        "worksheet `{part}` headerFooter contains unsupported `{}`; only oddHeader and oddFooter text are supported; remove even/first-page or other metadata",
        String::from_utf8_lossy(event.local_name().as_ref())
    ))
}

fn ensure_header_footer_child_has_no_attributes(
    event: &BytesStart<'_>,
    part: &str,
) -> xlsx_jsx::Result<()> {
    if event.attributes().with_checks(false).next().is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` {} has unsupported attributes; remove them to preserve header/footer text exactly",
            String::from_utf8_lossy(event.local_name().as_ref())
        )));
    }
    Ok(())
}

fn contains_header_footer_picture(text: &str) -> bool {
    let mut remaining = text;
    while let Some(index) = remaining.find('&') {
        remaining = &remaining[index..];
        if remaining.starts_with("&&") {
            remaining = &remaining[2..];
        } else if remaining.starts_with("&G") {
            return true;
        } else {
            remaining = &remaining[1..];
        }
    }
    false
}

fn parse_worksheet_page_breaks(xml: &str, part: &str) -> xlsx_jsx::Result<WorksheetPageBreaks> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut horizontal = PageBreakAudit::default();
    let mut vertical = PageBreakAudit::default();
    let mut active = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) => match event.local_name().as_ref() {
                b"rowBreaks" => {
                    audit_page_break_parent(
                        &event,
                        part,
                        reader.decoder(),
                        PageBreakAxis::Horizontal,
                        &mut horizontal,
                        &mut active,
                    )?;
                }
                b"colBreaks" => {
                    audit_page_break_parent(
                        &event,
                        part,
                        reader.decoder(),
                        PageBreakAxis::Vertical,
                        &mut vertical,
                        &mut active,
                    )?;
                }
                b"brk" => {
                    let axis = active.ok_or_else(|| Error::Reverse(format!("worksheet `{part}` has a page break outside its rowBreaks/colBreaks list; move it into the matching list")))?;
                    let id = parse_page_break(&event, part, reader.decoder(), axis)?;
                    page_break_values_mut(axis, &mut horizontal, &mut vertical)
                        .values
                        .push(id);
                }
                _ => {}
            },
            Ok(Event::Empty(event)) => match event.local_name().as_ref() {
                b"rowBreaks" => audit_empty_page_break_parent(
                    &event,
                    part,
                    reader.decoder(),
                    PageBreakAxis::Horizontal,
                    &mut horizontal,
                )?,
                b"colBreaks" => audit_empty_page_break_parent(
                    &event,
                    part,
                    reader.decoder(),
                    PageBreakAxis::Vertical,
                    &mut vertical,
                )?,
                b"brk" => {
                    let axis = active.ok_or_else(|| Error::Reverse(format!("worksheet `{part}` has a page break outside its rowBreaks/colBreaks list; move it into the matching list")))?;
                    let id = parse_page_break(&event, part, reader.decoder(), axis)?;
                    page_break_values_mut(axis, &mut horizontal, &mut vertical)
                        .values
                        .push(id);
                }
                _ => {}
            },
            Ok(Event::End(event)) if event.local_name().as_ref() == b"rowBreaks" => {
                active = None;
                finish_page_break_parent(&horizontal, part, "rowBreaks")?;
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"colBreaks" => {
                active = None;
                finish_page_break_parent(&vertical, part, "colBreaks")?;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` has malformed page-break XML: {error}"
                )));
            }
            _ => {}
        }
    }
    horizontal.values.sort_unstable();
    vertical.values.sort_unstable();
    Ok(WorksheetPageBreaks {
        horizontal: horizontal.values,
        vertical: vertical.values,
    })
}

fn page_break_values_mut<'a>(
    axis: PageBreakAxis,
    horizontal: &'a mut PageBreakAudit,
    vertical: &'a mut PageBreakAudit,
) -> &'a mut PageBreakAudit {
    match axis {
        PageBreakAxis::Horizontal => horizontal,
        PageBreakAxis::Vertical => vertical,
    }
}

fn page_break_tag(axis: PageBreakAxis) -> &'static str {
    match axis {
        PageBreakAxis::Horizontal => "rowBreaks",
        PageBreakAxis::Vertical => "colBreaks",
    }
}

fn audit_page_break_parent(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    axis: PageBreakAxis,
    audit: &mut PageBreakAudit,
    active: &mut Option<PageBreakAxis>,
) -> xlsx_jsx::Result<()> {
    let tag = page_break_tag(axis);
    if audit.seen || active.is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` repeats or nests `{tag}`; keep one page-break list per axis"
        )));
    }
    audit.seen = true;
    let mut attributes = BTreeMap::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name().as_ref().to_vec();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?
            .into_owned();
        if attributes.insert(key.clone(), value).is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` `{tag}` repeats `{}`; remove the duplicate attribute",
                String::from_utf8_lossy(&key)
            )));
        }
    }
    let count = attributes.remove(b"count".as_slice()).ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` `{tag}` is missing `count`; restore the break count"
        ))
    })?;
    let manual = attributes.remove(b"manualBreakCount".as_slice()).ok_or_else(|| Error::Reverse(format!("worksheet `{part}` `{tag}` is missing `manualBreakCount`; restore the manual break count")))?;
    if let Some(key) = attributes.keys().next() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` `{tag}` has unsupported `{}` metadata; remove it or add explicit page-break support",
            String::from_utf8_lossy(key)
        )));
    }
    let count = count.parse::<usize>().map_err(|_| Error::Reverse(format!("worksheet `{part}` `{tag}` has invalid count; use a nonnegative integer matching the breaks")))?;
    let manual = manual.parse::<usize>().map_err(|_| Error::Reverse(format!("worksheet `{part}` `{tag}` has invalid manualBreakCount; use a nonnegative integer matching the manual breaks")))?;
    audit.expected = Some((count, manual));
    *active = Some(axis);
    Ok(())
}

fn audit_empty_page_break_parent(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    axis: PageBreakAxis,
    audit: &mut PageBreakAudit,
) -> xlsx_jsx::Result<()> {
    let mut active = None;
    audit_page_break_parent(event, part, decoder, axis, audit, &mut active)?;
    finish_page_break_parent(audit, part, page_break_tag(axis))
}

fn parse_page_break(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    axis: PageBreakAxis,
) -> xlsx_jsx::Result<u32> {
    let (tag, id_limit, expected_max, max_text) = match axis {
        PageBreakAxis::Horizontal => ("rowBreaks", 1_048_576_u32, 16_383_u32, "16383"),
        PageBreakAxis::Vertical => ("colBreaks", 16_384_u32, 1_048_575_u32, "1048575"),
    };
    let mut id = None;
    let mut max = None;
    let mut manual = None;
    let mut minimum = None;
    let mut point = None;
    let mut seen = std::collections::HashSet::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        if !seen.insert(key.as_ref().to_vec()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` page break repeats `{}`; remove the duplicate attribute",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match key.as_ref() {
            b"id" => id = Some(value.parse::<u32>().map_err(|_| Error::Reverse(format!("worksheet `{part}` `{tag}` has invalid brk.id `{value}`; use a valid row/column index")))?),
            b"max" => max = Some(value.parse::<u32>().map_err(|_| Error::Reverse(format!("worksheet `{part}` `{tag}` has invalid brk.max `{value}`; use max={max_text} for a full-axis break")))?),
            b"man" => manual = Some(parse_xml_bool(&value, part, "brk.man")?),
            b"min" => minimum = Some(value.parse::<u32>().map_err(|_| Error::Reverse(format!("worksheet `{part}` `{tag}` has invalid brk.min `{value}`; use zero or omit it")))?),
            b"pt" => point = Some(parse_xml_bool(&value, part, "brk.pt")?),
            _ => return Err(Error::Reverse(format!("worksheet `{part}` page break has unsupported `{}` metadata; remove it or add explicit page-break support", String::from_utf8_lossy(key.as_ref())))),
        }
    }
    let id = id.ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` `{tag}` contains a break without `id`; restore its row/column index"
        ))
    })?;
    if id == 0
        || id >= id_limit
        || max != Some(expected_max)
        || minimum.is_some_and(|value| value != 0)
        || point == Some(true)
    {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has a partial page break; use id 1..{} and max={max_text} with default min/pt bounds",
            id_limit - 1
        )));
    }
    if manual != Some(true) {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has an automatic page break in `{tag}`; only manual breaks can be represented, so set man=1 or remove the automatic break"
        )));
    }
    Ok(id)
}

fn finish_page_break_parent(audit: &PageBreakAudit, part: &str, tag: &str) -> xlsx_jsx::Result<()> {
    let (count, manual_count) = audit.expected.ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` `{tag}` has no declared counts; restore count and manualBreakCount"
        ))
    })?;
    if count != audit.values.len() || manual_count != audit.values.len() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` `{tag}` declares count={count} and manualBreakCount={manual_count} but contains {} breaks; make both counts match the listed breaks",
            audit.values.len()
        )));
    }
    if audit.values.len() > 1023 {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` `{tag}` contains more than 1023 breaks; reduce the list"
        )));
    }
    let mut unique = std::collections::HashSet::new();
    if audit.values.iter().any(|id| !unique.insert(*id)) {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` `{tag}` repeats a page-break id; remove duplicate breaks"
        )));
    }
    Ok(())
}

fn require_matching_view_zoom(
    common: Option<u16>,
    specific: Option<u16>,
    part: &str,
    view: &str,
) -> xlsx_jsx::Result<()> {
    if common.is_some() && specific.is_some() && common != specific {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has conflicting common and {view} zoom values ({common:?} and {specific:?}); make them equal to preserve one zoom value"
        )));
    }
    Ok(())
}

fn parse_sheet_view_zoom(value: &str, part: &str, attribute: &str) -> xlsx_jsx::Result<u16> {
    let zoom = value.parse::<u16>().map_err(|_| {
        Error::Reverse(format!(
            "worksheet `{part}` has invalid `sheetView.{attribute}` value `{value}`; use an integer from 10 through 400"
        ))
    })?;
    if !(10..=400).contains(&zoom) {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has unsupported `sheetView.{attribute}` value `{zoom}`; use an integer from 10 through 400"
        )));
    }
    Ok(zoom)
}

fn reject_inactive_zoom(
    zoom: Option<u16>,
    part: &str,
    attribute: &str,
    view: &str,
) -> xlsx_jsx::Result<()> {
    if zoom.is_some_and(|value| value != 100) {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has non-default `{attribute}` for inactive worksheet view `{view}`; remove that zoom attribute or set view=\"{view}\" to preserve it"
        )));
    }
    Ok(())
}

fn parse_sheet_view_right_to_left(
    value: &str,
    part: &str,
    seen: &mut bool,
) -> xlsx_jsx::Result<bool> {
    if *seen {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` sheetView repeats `rightToLeft`; remove the duplicate attribute"
        )));
    }
    *seen = true;
    parse_xml_bool(value, part, "sheetView.rightToLeft")
}

fn audit_print_options(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    metadata: &mut WorksheetMetadata,
) -> xlsx_jsx::Result<()> {
    if metadata.print.options_seen.is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` contains multiple printOptions elements; combine their attributes into one element"
        )));
    }
    metadata.print.options_seen = Some(());
    let mut seen = std::collections::HashSet::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        if !matches!(
            key.as_ref(),
            b"gridLines" | b"horizontalCentered" | b"verticalCentered" | b"headings"
        ) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` printOptions has unsupported `{}` metadata; remove that option or add matching DSL support",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        if !seen.insert(key.as_ref().to_vec()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` printOptions repeats `{}`; remove the duplicate attribute",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        let enabled = parse_xml_bool(
            &value,
            part,
            &format!("printOptions.{}", String::from_utf8_lossy(key.as_ref())),
        )?;
        match key.as_ref() {
            b"gridLines" => metadata.print.options.gridlines = enabled,
            b"horizontalCentered" => metadata.print.options.centering.horizontally = enabled,
            b"verticalCentered" => metadata.print.options.centering.vertically = enabled,
            b"headings" => metadata.print.options.headings = enabled,
            _ => unreachable!("printOptions keys are allowlisted above"),
        }
    }
    Ok(())
}

fn audit_empty_page_setup(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    metadata: &mut WorksheetMetadata,
) -> xlsx_jsx::Result<()> {
    if metadata.print.page_setup_seen.is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` contains multiple pageSetup elements; keep one page setup or add explicit page-setup support"
        )));
    }
    metadata.print.page_setup_seen = Some(());
    let mut seen = std::collections::HashSet::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        if !seen.insert(key.as_ref().to_vec()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` pageSetup repeats `{}`; remove the duplicate attribute",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        match key.as_ref() {
            b"orientation" if value == "portrait" || value == "landscape" => {
                if value == "landscape" {
                    metadata.print.orientation = Some("landscape");
                }
            }
            b"pageOrder" if value == "overThenDown" => {
                metadata.print.page_order = Some("overThenDown");
            }
            b"pageOrder" if value == "downThenOver" => {}
            b"paperSize" => {
                let paper_size = value.parse::<u8>().map_err(|_| {
                    Error::Reverse(format!(
                        "worksheet `{part}` pageSetup has invalid `paperSize` value `{value}`; use an integer from 0 through 255"
                    ))
                })?;
                metadata.print.paper_size = Some(paper_size);
            }
            b"blackAndWhite" => {
                metadata.print.quality.black_and_white =
                    parse_xml_bool(&value, part, "pageSetup.blackAndWhite")?;
            }
            b"draft" => {
                metadata.print.quality.draft = parse_xml_bool(&value, part, "pageSetup.draft")?;
            }
            b"horizontalDpi" | b"verticalDpi" if value == "200" => {}
            b"scale" => {
                let scale = parse_page_setup_u16(&value, part, "scale")?;
                if !(10..=400).contains(&scale) {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` pageSetup scale `{scale}` is outside 10 through 400; choose a valid printed-page percentage"
                    )));
                }
                metadata.print.print_scale = Some(scale);
            }
            b"fitToWidth" => {
                metadata.print.fit_width = Some(parse_page_setup_u16(&value, part, "fitToWidth")?);
            }
            b"fitToHeight" => {
                metadata.print.fit_height =
                    Some(parse_page_setup_u16(&value, part, "fitToHeight")?);
            }
            _ => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` pageSetup has unsupported `{}` value `{value}`; restore writer defaults or add matching page-setup DSL support",
                    String::from_utf8_lossy(key.as_ref())
                )));
            }
        }
    }
    Ok(())
}

fn parse_page_setup_u16(value: &str, part: &str, attribute: &str) -> xlsx_jsx::Result<u16> {
    value.parse::<u16>().map_err(|_| {
        Error::Reverse(format!(
            "worksheet `{part}` pageSetup has invalid `{attribute}` value `{value}`; use an integer from 0 through 65535"
        ))
    })
}

fn audit_page_setup_properties(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    print: &mut WorksheetPrintMetadata,
) -> xlsx_jsx::Result<()> {
    if print.page_setup_properties_seen.is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` contains multiple pageSetUpPr elements; keep one fit-to-page setting"
        )));
    }
    print.page_setup_properties_seen = Some(());
    let mut seen = std::collections::HashSet::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        if !seen.insert(key.as_ref().to_vec()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` pageSetUpPr repeats `{}`; remove the duplicate attribute",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match key.as_ref() {
            b"fitToPage" => {
                print.fit_to_page = Some(parse_xml_bool(&value, part, "pageSetUpPr.fitToPage")?);
            }
            b"autoPageBreaks" => {
                if !parse_xml_bool(&value, part, "pageSetUpPr.autoPageBreaks")? {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` has autoPageBreaks=false, which v1 cannot preserve; restore the default automatic page-break behavior"
                    )));
                }
            }
            _ => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` pageSetUpPr has unsupported `{}` value `{value}`; use fitToPage or the default autoPageBreaks value",
                    String::from_utf8_lossy(key.as_ref())
                )));
            }
        }
    }
    Ok(())
}

fn validate_page_setup_consistency(
    metadata: &mut WorksheetMetadata,
    part: &str,
) -> xlsx_jsx::Result<()> {
    let fit_enabled = metadata.print.fit_to_page == Some(true);
    let has_fit_dimensions =
        metadata.print.fit_width.is_some() || metadata.print.fit_height.is_some();
    if fit_enabled {
        if metadata.print.print_scale.is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has both pageSetup scale and fitToPage enabled; these print modes are mutually exclusive, so keep only one"
            )));
        }
        if !metadata.page_breaks.horizontal.is_empty() || !metadata.page_breaks.vertical.is_empty()
        {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` combines fitToPage with manual page breaks, which fit-to-page overrides; remove one print setting"
            )));
        }
        metadata.print.fit_to_pages = Some(FitToPages {
            width: metadata.print.fit_width.unwrap_or(1),
            height: metadata.print.fit_height.unwrap_or(1),
        });
    } else if has_fit_dimensions
        && (metadata.print.fit_width.is_some_and(|value| value != 1)
            || metadata.print.fit_height.is_some_and(|value| value != 1))
    {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has non-default fitToWidth/fitToHeight without fitToPage=1; enable fit-to-page or remove the orphaned fit dimensions"
        )));
    }
    Ok(())
}

fn audit_freeze_pane(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    freeze_panes: &mut Option<FreezePanes>,
) -> xlsx_jsx::Result<()> {
    if freeze_panes.is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` contains multiple panes; keep one frozen pane or add split-pane support"
        )));
    }
    let mut rows = 0_u32;
    let mut columns = 0_u16;
    let mut top_left = None;
    let mut active_pane = None;
    let mut state = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match attribute.key.local_name().as_ref() {
            b"xSplit" => {
                columns = value.parse::<u16>().map_err(|error| {
                    Error::Reverse(format!(
                        "worksheet `{part}` pane has an invalid xSplit: {error}"
                    ))
                })?;
            }
            b"ySplit" => {
                rows = value.parse::<u32>().map_err(|error| {
                    Error::Reverse(format!(
                        "worksheet `{part}` pane has an invalid ySplit: {error}"
                    ))
                })?;
            }
            b"topLeftCell" => top_left = Some(value.into_owned()),
            b"activePane" => active_pane = Some(value.into_owned()),
            b"state" => state = Some(value.into_owned()),
            key => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` pane has unsupported `{}` metadata; remove the pane option or add matching DSL support",
                    String::from_utf8_lossy(key)
                )));
            }
        }
    }
    if rows == 0 && columns == 0 {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` pane has no frozen rows or columns; remove the empty pane metadata"
        )));
    }
    if rows > 1_048_575 || columns > 16_383 {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` pane split exceeds Excel limits; repair xSplit/ySplit"
        )));
    }
    let top_left = top_left.ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` pane is missing topLeftCell; repair the pane metadata"
        ))
    })?;
    let (top_left_row, top_left_column) = parse_a1_cell_reference(&top_left).ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` pane topLeftCell `{top_left}` is invalid; repair it to an Excel A1 cell"
        ))
    })?;
    if top_left_row < rows || u32::from(top_left_column) < u32::from(columns) {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` pane topLeftCell `{top_left}` is inside the frozen area; move it beyond the frozen rows and columns"
        )));
    }
    let expected_active_pane = match (rows > 0, columns > 0) {
        (true, true) => "bottomRight",
        (true, false) => "bottomLeft",
        (false, true) => "topRight",
        (false, false) => "",
    };
    if active_pane.as_deref() != Some(expected_active_pane) || state.as_deref() != Some("frozen") {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` uses a split or non-canonical pane state; use frozen panes with activePane `{expected_active_pane}`"
        )));
    }
    *freeze_panes = Some(FreezePanes {
        rows,
        columns,
        top_left_row,
        top_left_column,
    });
    Ok(())
}

fn parse_pane_selection(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<PaneSelection> {
    let mut selection = PaneSelection {
        pane: None,
        active_cell: None,
        sqref: None,
    };
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match attribute.key.local_name().as_ref() {
            b"pane" => selection.pane = Some(value.into_owned()),
            b"activeCell" => {
                let (row, column) = parse_a1_cell_reference(&value).ok_or_else(|| {
                    Error::Reverse(format!(
                        "worksheet `{part}` selection activeCell `{value}` is invalid; restore the default pane selection"
                    ))
                })?;
                selection.active_cell = Some(format_a1_cell_reference(row, column));
            }
            b"sqref" => {
                let canonical = if let Some(range) = parse_a1_range(&value) {
                    format_a1_range(range)
                } else if let Some((row, column)) = parse_a1_cell_reference(&value) {
                    format_a1_cell_reference(row, column)
                } else {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` selection sqref `{value}` is unsupported; restore the default pane selection"
                    )));
                };
                selection.sqref = Some(canonical);
            }
            key => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` selection has unsupported `{}` metadata; restore the default selection or add selection support",
                    String::from_utf8_lossy(key)
                )));
            }
        }
    }
    Ok(selection)
}

fn validate_pane_selections(
    metadata: &WorksheetMetadata,
    actual: &[PaneSelection],
    part: &str,
) -> xlsx_jsx::Result<Option<WorksheetSelection>> {
    let Some(panes) = metadata.freeze_panes else {
        if actual.is_empty() {
            return Ok(None);
        }
        if actual.len() != 1 || actual[0].pane.is_some() {
            return Err(selection_error(part));
        }
        return selection_from_pane_selection(&actual[0], part);
    };
    let mut expected = Vec::new();
    let mut add_selection = |pane: &str, cell: Option<String>| {
        expected.push(PaneSelection {
            pane: Some(pane.to_owned()),
            sqref: cell.clone(),
            active_cell: cell,
        });
    };
    match (panes.rows > 0, panes.columns > 0) {
        (true, true) => {
            add_selection("topRight", Some(format_a1_cell_reference(0, panes.columns)));
            add_selection("bottomLeft", Some(format_a1_cell_reference(panes.rows, 0)));
            add_selection("bottomRight", None);
        }
        (true, false) => add_selection("bottomLeft", None),
        (false, true) => add_selection("topRight", None),
        (false, false) => {}
    }
    if actual.is_empty() {
        return Ok(None);
    }
    let active_pane = match (panes.rows > 0, panes.columns > 0) {
        (true, true) => "bottomRight",
        (true, false) => "bottomLeft",
        (false, true) => "topRight",
        (false, false) => return Err(selection_error(part)),
    };
    if actual.len() != expected.len() {
        return Err(selection_error(part));
    }
    let mut active_selection = None;
    for expected_item in &expected {
        let Some(actual_item) = actual.iter().find(|item| item.pane == expected_item.pane) else {
            return Err(selection_error(part));
        };
        if expected_item.pane.as_deref() == Some(active_pane) {
            active_selection = selection_from_pane_selection(actual_item, part)?;
        } else if actual_item != expected_item {
            return Err(selection_error(part));
        }
    }
    Ok(active_selection)
}

fn selection_error(part: &str) -> xlsx_jsx::Error {
    Error::Reverse(format!(
        "worksheet `{part}` contains a non-canonical or multi-range pane selection; use one contiguous range with a supported active cell, or restore the default pane layout"
    ))
}

fn selection_from_pane_selection(
    selection: &PaneSelection,
    part: &str,
) -> xlsx_jsx::Result<Option<WorksheetSelection>> {
    match (&selection.active_cell, &selection.sqref) {
        (None, None) => Ok(None),
        (Some(active_cell), Some(range)) => {
            let active =
                parse_a1_cell_reference(active_cell).ok_or_else(|| selection_error(part))?;
            let parsed = parse_a1_range(range)
                .or_else(|| {
                    parse_a1_cell_reference(range).map(|(row, column)| (row, column, row, column))
                })
                .ok_or_else(|| selection_error(part))?;
            if active.0 < parsed.0
                || active.0 > parsed.2
                || active.1 < parsed.1
                || active.1 > parsed.3
            {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` selection activeCell `{active_cell}` is outside sqref `{range}`; move it inside the selected range"
                )));
            }
            if (active.0 != parsed.0 && active.0 != parsed.2)
                || (active.1 != parsed.1 && active.1 != parsed.3)
            {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` selection activeCell `{active_cell}` is inside sqref `{range}` but not at a range corner; rust_xlsxwriter can only emit a corner active cell, so move it to a corner"
                )));
            }
            let range = if parsed.0 == parsed.2 && parsed.1 == parsed.3 {
                format_a1_cell_reference(parsed.0, parsed.1)
            } else {
                format_a1_range(parsed)
            };
            Ok(Some(WorksheetSelection {
                active_cell: format_a1_cell_reference(active.0, active.1),
                range,
            }))
        }
        _ => Err(selection_error(part)),
    }
}

fn audit_cell_attributes(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<(String, usize)> {
    let mut reference = None;
    let mut style_index = 0;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        if key.as_ref() == b"r" {
            reference = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| Error::Reverse(error.to_string()))?
                    .into_owned(),
            );
        } else if key.as_ref() == b"s" {
            style_index = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?
                .parse::<usize>()
                .map_err(|error| Error::Reverse(format!("invalid cell style index: {error}")))?;
        }
        if ![b"r".as_slice(), b"t", b"s", b"cm", b"vm", b"ph"].contains(&key.as_ref()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` cell contains unsupported attribute `{}`; add explicit cell metadata support before reversing",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
    }
    let reference = reference.ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` contains a cell without a coordinate"
        ))
    })?;
    Ok((reference, style_index))
}

#[allow(clippy::too_many_lines)] // Attribute validation is deliberately kept adjacent to its whitelist.
fn audit_row_attributes(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    formats: &[CellFontStyle],
    fallback_index: u32,
) -> xlsx_jsx::Result<(u32, Option<RowProperties>)> {
    let mut row_index = None;
    let mut height = None;
    let mut custom_height = false;
    let mut hidden = false;
    let mut outline_level = 0_u8;
    let mut collapsed = false;
    let mut collapsed_seen = false;
    let mut style_index = None;
    let mut custom_format = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match key.as_ref() {
            b"r" => {
                let one_based = value.parse::<u32>().map_err(|error| {
                    Error::Reverse(format!(
                        "worksheet `{part}` has an invalid row index: {error}"
                    ))
                })?;
                row_index = Some(one_based.checked_sub(1).ok_or_else(|| {
                    Error::Reverse(format!("worksheet `{part}` contains row index 0"))
                })?);
                if one_based > 1_048_576 {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` row index {one_based} exceeds the XLSX row limit; repair the worksheet"
                    )));
                }
            }
            b"spans" => {}
            b"dyDescent" => {
                let is_default = attribute.key.as_ref() == b"x14ac:dyDescent"
                    && value
                        .parse::<f64>()
                        .is_ok_and(|number| number.to_bits() == 0.25_f64.to_bits());
                if !is_default {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` row uses unsupported dyDescent metadata; only the ignorable x14ac default 0.25 is representable"
                    )));
                }
            }
            b"s" => {
                style_index = Some(value.parse::<usize>().map_err(|error| {
                    Error::Reverse(format!(
                        "worksheet `{part}` has invalid row style index `{value}`: {error}"
                    ))
                })?);
            }
            b"customFormat" => {
                custom_format = Some(parse_xml_bool(&value, part, "customFormat")?);
            }
            b"ht" => {
                height = Some(value.parse::<f64>().map_err(|error| {
                    Error::Reverse(format!(
                        "worksheet `{part}` has an invalid row height: {error}"
                    ))
                })?);
            }
            b"customHeight" => custom_height = parse_xml_bool(&value, part, "customHeight")?,
            b"hidden" => hidden = parse_xml_bool(&value, part, "hidden")?,
            b"collapsed" => {
                if collapsed_seen {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` row repeats `collapsed`; remove the duplicate attribute"
                    )));
                }
                collapsed_seen = true;
                collapsed = parse_xml_bool(&value, part, "collapsed")?;
            }
            b"outlineLevel" => {
                outline_level = value.parse::<u8>().map_err(|error| {
                    Error::Reverse(format!(
                        "worksheet `{part}` has invalid row outline level `{value}`: {error}"
                    ))
                })?;
                if outline_level > 7 {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` row {} has outline level {outline_level} above Excel's maximum of 7; lower the level or repair the row metadata",
                        row_index.unwrap_or(fallback_index) + 1
                    )));
                }
            }
            _ => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` contains unsupported row attribute `{}`; remove row style metadata or add matching DSL support",
                    String::from_utf8_lossy(key.as_ref())
                )));
            }
        }
    }
    let height = match (height, custom_height) {
        (Some(height), true) => {
            let pixels = height * 4.0 / 3.0;
            if !height.is_finite()
                || !(0.75..=409.5).contains(&height)
                || (pixels - pixels.round()).abs() > 1e-9
            {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` row {} has height {height} points that cannot round-trip at 0.75-point precision; normalize the row height or add higher-precision height support",
                    row_index.unwrap_or(fallback_index) + 1
                )));
            }
            Some(height)
        }
        (Some(_), false) => {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` row {} has `ht` without `customHeight=1`; repair the row height metadata",
                row_index.unwrap_or(fallback_index) + 1
            )));
        }
        (None, true) => {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` row {} has `customHeight=1` without `ht`; repair the row height metadata",
                row_index.unwrap_or(fallback_index) + 1
            )));
        }
        (None, false) => None,
    };
    let properties = RowProperties {
        height,
        hidden,
        outline_level,
        collapsed,
        style_index: validate_default_style_index(
            style_index.unwrap_or_default(),
            custom_format,
            formats,
            part,
            &format!("row {}", row_index.unwrap_or(fallback_index) + 1),
        )?,
    };
    let row_index = row_index.unwrap_or(fallback_index);
    Ok((
        row_index,
        (height.is_some()
            || hidden
            || outline_level > 0
            || collapsed
            || properties.style_index.is_some())
        .then_some(properties),
    ))
}

fn parse_xml_bool(value: &str, part: &str, attribute: &str) -> xlsx_jsx::Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(Error::Reverse(format!(
            "worksheet `{part}` has invalid `{attribute}` value `{value}`; expected 0/1 or false/true"
        ))),
    }
}

fn audit_page_layout_metadata(
    event: &BytesStart<'_>,
    tag: &str,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    default_row_height_pixels: &mut Option<u32>,
    hide_unused_rows: &mut bool,
    print: &mut WorksheetPrintMetadata,
) -> xlsx_jsx::Result<()> {
    match tag {
        "sheetFormatPr" => audit_sheet_format(
            event,
            part,
            decoder,
            default_row_height_pixels,
            hide_unused_rows,
        ),
        "pageMargins" => audit_page_margins(
            event,
            part,
            decoder,
            &mut print.page_margins,
            &mut print.page_margins_seen,
        ),
        _ => Ok(()),
    }
}

fn audit_sheet_properties_if_present(
    event: &BytesStart<'_>,
    tag: &str,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    metadata: &mut WorksheetMetadata,
) -> xlsx_jsx::Result<()> {
    if tag == "sheetPr" {
        audit_sheet_properties(event, part, decoder, &mut metadata.filter_mode)?;
    } else if tag == "pageSetUpPr" {
        audit_page_setup_properties(event, part, decoder, &mut metadata.print)?;
    } else if tag == "outlinePr" {
        audit_outline_properties(
            event,
            part,
            decoder,
            &mut metadata.summary_below,
            &mut metadata.summary_right,
        )?;
    }
    Ok(())
}

fn audit_outline_properties(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    summary_below: &mut Option<bool>,
    summary_right: &mut Option<bool>,
) -> xlsx_jsx::Result<()> {
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        let target = match key.as_ref() {
            b"summaryBelow" => &mut *summary_below,
            b"summaryRight" => &mut *summary_right,
            _ => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` outlinePr contains unsupported `{}` setting; remove it or add explicit outline support",
                    String::from_utf8_lossy(key.as_ref())
                )));
            }
        };
        if target.is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` outlinePr repeats `{}`; remove the duplicate setting",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        let attribute_name = format!("outlinePr.{}", String::from_utf8_lossy(key.as_ref()));
        *target = Some(parse_xml_bool(&value, part, &attribute_name)?);
    }
    Ok(())
}

fn audit_tab_color(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    tab_color: &mut Option<String>,
) -> xlsx_jsx::Result<()> {
    if tab_color.is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` repeats tabColor; keep one color or add support for duplicate metadata"
        )));
    }
    let mut attributes = HashMap::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| {
            Error::Reverse(format!(
                "worksheet `{part}` has invalid tabColor attribute: {error}"
            ))
        })?;
        let key = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        if key != "rgb" {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` tabColor has unsupported `{key}` metadata; use an explicit opaque RGB color or add support for theme/indexed colors"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?
            .into_owned();
        if attributes.insert(key.clone(), value).is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` tabColor repeats `{key}`; remove the duplicate attribute"
            )));
        }
    }
    let rgb = attributes.get("rgb").ok_or_else(|| {
        Error::Reverse(format!(
            "worksheet `{part}` tabColor is missing an explicit rgb value; use opaque RGB or add support for theme/indexed colors"
        ))
    })?;
    if rgb.len() != 8
        || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit())
        || !rgb[..2].eq_ignore_ascii_case("FF")
    {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` tabColor `{rgb}` is not an opaque ARGB RGB value; use alpha FF or add explicit alpha handling"
        )));
    }
    *tab_color = Some(format!("#{}", &rgb[2..]).to_ascii_uppercase());
    Ok(())
}

fn audit_sheet_properties(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    filter_mode: &mut Option<bool>,
) -> xlsx_jsx::Result<()> {
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"filterMode" || filter_mode.is_some() {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` sheetPr has unsupported metadata; only derived filterMode is modeled"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        *filter_mode = Some(parse_xml_bool(&value, part, "sheetPr.filterMode")?);
    }
    Ok(())
}

fn audit_sheet_format(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    default_row_height_pixels: &mut Option<u32>,
    hide_unused_rows: &mut bool,
) -> xlsx_jsx::Result<()> {
    let mut custom_height = None;
    let mut zero_height_seen = false;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        let accepted_default = match key.as_ref() {
            b"dyDescent" => {
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                if attribute.key.as_ref() != b"x14ac:dyDescent"
                    || !value
                        .parse::<f64>()
                        .is_ok_and(|number| number.to_bits() == 0.25_f64.to_bits())
                {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` has unsupported sheetFormatPr dyDescent metadata; only the ignorable x14ac default 0.25 is representable"
                    )));
                }
                None
            }
            b"defaultRowHeight" => {
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                let height = value.parse::<f64>().map_err(|_| {
                    Error::Reverse(format!("worksheet `{part}` has invalid defaultRowHeight `{value}`; use a positive numeric OOXML height"))
                })?;
                let pixels = height / 0.75;
                let rounded = pixels.round();
                if !height.is_finite()
                    || height <= 0.0
                    || (pixels - rounded).abs() > 1e-9
                    || !(1.0..=545.0).contains(&rounded)
                {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` defaultRowHeight `{value}` cannot be represented exactly as an integer pixel height; adjust it to a multiple of 0.75 points within 1–545 pixels (Excel's 409-point limit)"
                    )));
                }
                // The preceding range check proves this is a positive integer <= 545.
                #[expect(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
                let pixels = rounded as u32;
                *default_row_height_pixels = (pixels != 20).then_some(pixels);
                None
            }
            b"defaultColWidth" => Some("8.43"),
            b"customHeight" => {
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                custom_height = Some(parse_xml_bool(&value, part, "sheetFormatPr.customHeight")?);
                None
            }
            b"zeroHeight" => {
                if zero_height_seen {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` sheetFormatPr repeats `zeroHeight`; remove the duplicate attribute"
                    )));
                }
                zero_height_seen = true;
                let value = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                *hide_unused_rows = parse_xml_bool(&value, part, "sheetFormatPr.zeroHeight")?;
                None
            }
            _ => None,
        };
        if accepted_default.is_some_and(|default| {
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .is_ok_and(|value| value != default)
        }) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` has a non-default row/column layout setting that v1 cannot preserve; restore defaults or add worksheet layout support"
            )));
        }
    }
    if custom_height == Some(false) {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` has customHeight=false, which conflicts with explicitly stored row-height metadata; remove the contradictory attribute"
        )));
    }
    Ok(())
}

fn audit_page_margins(
    event: &BytesStart<'_>,
    part: &str,
    decoder: quick_xml::encoding::Decoder,
    margins: &mut WorksheetPageMargins,
    seen_margins: &mut Option<()>,
) -> xlsx_jsx::Result<()> {
    if seen_margins.is_some() {
        return Err(Error::Reverse(format!(
            "worksheet `{part}` contains multiple pageMargins elements; keep one margin set"
        )));
    }
    *seen_margins = Some(());
    let mut seen = std::collections::HashSet::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        if !seen.insert(key.as_ref().to_vec()) {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` pageMargins repeats `{}`; remove the duplicate attribute",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        let margin = value.parse::<f64>().map_err(|_| Error::Reverse(format!(
            "worksheet `{part}` pageMargins has invalid `{}` value `{value}`; use a finite nonnegative number of inches",
            String::from_utf8_lossy(key.as_ref())
        )))?;
        if !margin.is_finite() || margin < 0.0 {
            return Err(Error::Reverse(format!(
                "worksheet `{part}` pageMargins `{}` value `{margin}` must be a finite nonnegative number of inches",
                String::from_utf8_lossy(key.as_ref())
            )));
        }
        match key.as_ref() {
            b"left" => margins.left = margin,
            b"right" => margins.right = margin,
            b"top" => margins.top = margin,
            b"bottom" => margins.bottom = margin,
            b"header" => margins.header = margin,
            b"footer" => margins.footer = margin,
            _ => {
                return Err(Error::Reverse(format!(
                    "worksheet `{part}` pageMargins has unsupported `{}` metadata; use left/right/top/bottom/header/footer or add explicit margin support",
                    String::from_utf8_lossy(key.as_ref())
                )));
            }
        }
    }
    Ok(())
}

fn parse_shared_strings(xml: &str, part: &str) -> xlsx_jsx::Result<Vec<Option<Vec<RichTextRun>>>> {
    parse_rich_text_entries(xml, b"si", part)
}

#[expect(
    clippy::too_many_lines,
    reason = "Rich-string parsing must preserve the ordered boundaries among t, r, rPr, and phonetic nodes."
)]
fn parse_rich_text_entries(
    xml: &str,
    entry_tag: &[u8],
    part: &str,
) -> xlsx_jsx::Result<Vec<Option<Vec<RichTextRun>>>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut entries = Vec::new();
    let mut in_entry = false;
    let mut in_run_properties = false;
    let mut in_text = false;
    let mut current_run: Option<RichTextRun> = None;
    let mut current_run_text_seen = false;
    let mut run_properties_seen = false;
    let mut seen_run_properties = std::collections::HashSet::new();
    let mut text = String::new();
    let mut runs = Vec::new();
    let mut has_rich_runs = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) => {
                let tag = event.local_name();
                if tag.as_ref() == entry_tag {
                    if in_entry {
                        return Err(Error::Reverse(format!(
                            "`{part}` contains nested rich-text entries; repair the string XML"
                        )));
                    }
                    if event.attributes().next().is_some() {
                        return Err(Error::Reverse(format!(
                            "`{part}` contains unsupported rich-text entry attributes; remove the extra metadata"
                        )));
                    }
                    in_entry = true;
                    runs.clear();
                    has_rich_runs = false;
                } else if in_entry {
                    match tag.as_ref() {
                        b"r" => {
                            if current_run.is_some() || event.attributes().next().is_some() {
                                return Err(Error::Reverse(format!(
                                    "`{part}` contains nested runs or unsupported rich-text run attributes; repair the string XML"
                                )));
                            }
                            current_run = Some(RichTextRun {
                                format: CellFontStyle {
                                    support: StyleSupport::Supported,
                                    ..Default::default()
                                },
                                ..Default::default()
                            });
                            has_rich_runs = true;
                            current_run_text_seen = false;
                            run_properties_seen = false;
                            seen_run_properties.clear();
                        }
                        b"rPr" => {
                            if current_run.is_none()
                                || run_properties_seen
                                || current_run_text_seen
                                || event.attributes().next().is_some()
                            {
                                return Err(Error::Reverse(format!(
                                    "`{part}` contains unsupported rich-text run properties; remove run metadata not represented by TextRun"
                                )));
                            }
                            run_properties_seen = true;
                            in_run_properties = true;
                        }
                        b"t" => {
                            validate_rich_text_node_attributes(&event, part)?;
                            if current_run.is_some() && current_run_text_seen {
                                return Err(Error::Reverse(format!(
                                    "`{part}` contains multiple text elements in one rich-text run; merge their text into one `<t>` element"
                                )));
                            }
                            if current_run.is_some() {
                                current_run_text_seen = true;
                            }
                            in_text = true;
                            text.clear();
                        }
                        b"rPh" | b"phoneticPr" => return Err(unsupported_phonetics(part)),
                        _ if in_run_properties => {
                            if !seen_run_properties.insert(tag.as_ref().to_vec()) {
                                return Err(unsupported_rich_run_format(part));
                            }
                            parse_rich_run_property(
                                &event,
                                tag.as_ref(),
                                &mut current_run,
                                reader.decoder(),
                                part,
                            )?;
                        }
                        _ => {
                            return Err(Error::Reverse(format!(
                                "`{part}` contains unsupported rich-text element `{}`; remove it or add matching TextRun support",
                                String::from_utf8_lossy(tag.as_ref())
                            )));
                        }
                    }
                }
            }
            Ok(Event::Empty(event)) => {
                let tag = event.local_name();
                if in_entry {
                    match tag.as_ref() {
                        b"r" => {
                            return Err(Error::Reverse(format!(
                                "`{part}` contains an empty rich-text run; add non-empty text to every run"
                            )));
                        }
                        b"rPr" => {
                            if current_run.is_none()
                                || run_properties_seen
                                || current_run_text_seen
                                || event.attributes().next().is_some()
                            {
                                return Err(Error::Reverse(format!(
                                    "`{part}` contains unsupported rich-text run properties; remove run metadata not represented by TextRun"
                                )));
                            }
                            run_properties_seen = true;
                        }
                        b"t" => {
                            validate_rich_text_node_attributes(&event, part)?;
                            if current_run.is_some() {
                                return Err(Error::Reverse(format!(
                                    "`{part}` contains an empty rich-text `<t>` element; add text to the run"
                                )));
                            }
                        }
                        b"rPh" | b"phoneticPr" => return Err(unsupported_phonetics(part)),
                        _ if in_run_properties => {
                            parse_rich_run_property(
                                &event,
                                tag.as_ref(),
                                &mut current_run,
                                reader.decoder(),
                                part,
                            )?;
                        }
                        _ => {
                            return Err(Error::Reverse(format!(
                                "`{part}` contains unsupported rich-text element `{}`; remove it or add matching TextRun support",
                                String::from_utf8_lossy(tag.as_ref())
                            )));
                        }
                    }
                }
            }
            Ok(Event::Text(value)) if in_text => {
                let decoded = value.decode().map_err(|error| {
                    Error::Reverse(format!("invalid text in `{part}`: {error}"))
                })?;
                text.push_str(&quick_xml::escape::unescape(&decoded).map_err(|error| {
                    Error::Reverse(format!("invalid escaped text in `{part}`: {error}"))
                })?);
            }
            Ok(Event::CData(value)) if in_text => {
                text.push_str(&value.decode().map_err(|error| {
                    Error::Reverse(format!("invalid CDATA in `{part}`: {error}"))
                })?);
            }
            Ok(Event::GeneralRef(reference)) if in_text => {
                if let Some(character) = reference.resolve_char_ref().map_err(|error| {
                    Error::Reverse(format!("invalid character reference in `{part}`: {error}"))
                })? {
                    text.push(character);
                } else {
                    let entity = reference.decode().map_err(|error| {
                        Error::Reverse(format!("invalid entity in `{part}`: {error}"))
                    })?;
                    let resolved = quick_xml::escape::resolve_predefined_entity(&entity)
                        .ok_or_else(|| Error::Reverse(format!(
                            "`{part}` contains an unsupported XML entity `&{entity};` in rich text; replace it with literal text"
                        )))?;
                    text.push_str(resolved);
                }
            }
            Ok(Event::End(event)) => match event.local_name().as_ref() {
                b"t" if in_text => {
                    in_text = false;
                    if let Some(run) = current_run.as_mut() {
                        run.text.push_str(&text);
                    } else if !text.is_empty() {
                        runs.push(RichTextRun {
                            text: std::mem::take(&mut text),
                            format: CellFontStyle {
                                support: StyleSupport::Supported,
                                ..Default::default()
                            },
                        });
                    }
                }
                b"rPr" if in_run_properties => in_run_properties = false,
                b"r" => {
                    let Some(mut run) = current_run.take() else {
                        return Err(Error::Reverse(format!(
                            "`{part}` contains an unmatched rich-text run end; repair the string XML"
                        )));
                    };
                    if run.text.is_empty() {
                        return Err(Error::Reverse(format!(
                            "`{part}` contains an empty rich-text run; add non-empty text to every run"
                        )));
                    }
                    run.format = finalize_font_style(run.format);
                    if run.format.support != StyleSupport::Supported {
                        return Err(unsupported_rich_run_format(part));
                    }
                    runs.push(run);
                }
                tag if tag == entry_tag => {
                    if current_run.is_some() || in_text || in_run_properties {
                        return Err(Error::Reverse(format!(
                            "`{part}` contains an incomplete rich-text entry; repair the string XML"
                        )));
                    }
                    if has_rich_runs {
                        validate_reverse_rich_run_text(part, &runs)?;
                    }
                    entries.push(has_rich_runs.then(|| std::mem::take(&mut runs)));
                    in_entry = false;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid XML in `{part}`: {error}"))),
        }
    }
    if in_entry {
        return Err(Error::Reverse(format!(
            "`{part}` contains an unfinished string entry"
        )));
    }
    Ok(entries)
}

fn validate_reverse_rich_run_text(part: &str, runs: &[RichTextRun]) -> xlsx_jsx::Result<()> {
    let mut total_utf16_units = 0_usize;
    for run in runs {
        if run.text.is_empty() || !run.text.chars().all(is_valid_xml_character) {
            return Err(Error::Reverse(format!(
                "`{part}` contains an empty run or XML-invalid rich-text character; repair the text before reversing"
            )));
        }
        total_utf16_units = total_utf16_units
            .checked_add(run.text.encode_utf16().count())
            .ok_or_else(|| Error::Reverse(format!("`{part}` rich-text length overflow")))?;
        if total_utf16_units > 32_767 {
            return Err(Error::Reverse(format!(
                "`{part}` contains a rich-text string longer than Excel's 32,767 UTF-16-unit limit; shorten its combined run text"
            )));
        }
    }
    Ok(())
}

fn is_valid_xml_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || ('\u{20}'..='\u{d7ff}').contains(&character)
        || ('\u{e000}'..='\u{fffd}').contains(&character)
        || ('\u{10000}'..='\u{10ffff}').contains(&character)
}

fn validate_rich_text_node_attributes(event: &BytesStart<'_>, part: &str) -> xlsx_jsx::Result<()> {
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"space" {
            return Err(Error::Reverse(format!(
                "`{part}` contains unsupported rich-text `<t>` metadata; preserve only text and xml:space"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, event.decoder())
            .map_err(|error| Error::Reverse(error.to_string()))?;
        if !matches!(value.as_ref(), "preserve" | "default") {
            return Err(Error::Reverse(format!(
                "`{part}` has invalid xml:space value `{value}` in rich text; use `preserve` or `default`"
            )));
        }
    }
    Ok(())
}

fn parse_rich_run_property(
    event: &BytesStart<'_>,
    tag: &[u8],
    current_run: &mut Option<RichTextRun>,
    decoder: quick_xml::encoding::Decoder,
    part: &str,
) -> xlsx_jsx::Result<()> {
    let Some(run) = current_run.as_mut() else {
        return Err(Error::Reverse(format!(
            "`{part}` contains rich-text font properties outside a run"
        )));
    };
    let style = &mut run.format;
    match tag {
        b"b" | b"i" | b"strike" | b"u" | b"vertAlign" | b"color" => {
            parse_font_component(event, tag, style, decoder)?;
        }
        b"rFont" => parse_font_name_or_size(event, b"name", style, decoder, false)?,
        b"sz" => parse_font_name_or_size(event, tag, style, decoder, false)?,
        b"family" | b"charset" => parse_font_family_or_charset(event, tag, style, decoder)?,
        b"scheme" => parse_font_scheme(event, style, decoder, false)?,
        _ => return Err(unsupported_rich_run_format(part)),
    }
    if style.support != StyleSupport::Supported {
        return Err(unsupported_rich_run_format(part));
    }
    Ok(())
}

fn unsupported_phonetics(part: &str) -> Error {
    Error::Reverse(format!(
        "`{part}` contains rich-text phonetic annotations that TextRun cannot represent; remove phonetic metadata or extend the DSL and reverse reader"
    ))
}

fn unsupported_rich_run_format(part: &str) -> Error {
    Error::Reverse(format!(
        "`{part}` contains unsupported rich-text font properties; preserve only bold, italic, color, font name/size/family/charset, underline, strikethrough, and superscript/subscript or add matching TextRun support"
    ))
}

#[expect(
    clippy::too_many_lines,
    reason = "Worksheet rich-string references must be associated with shared and inline payloads in a single streaming pass."
)]
fn parse_worksheet_rich_text(
    xml: &str,
    part: &str,
    shared_strings: &[Option<Vec<RichTextRun>>],
) -> xlsx_jsx::Result<HashMap<String, Vec<RichTextRun>>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut rich_cells = HashMap::new();
    let mut cell_reference = None;
    let mut cell_type = None;
    let mut shared_index = String::new();
    let mut in_shared_index = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if event.local_name().as_ref() == b"c" => {
                let mut reference = None;
                let mut cell_type_value = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::Reverse(error.to_string()))?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"r" => reference = Some(value),
                        b"t" => cell_type_value = Some(value),
                        _ => {}
                    }
                }
                cell_reference = reference;
                cell_type = cell_type_value;
                shared_index.clear();
            }
            Ok(Event::Start(event)) if event.local_name().as_ref() == b"is" => {
                if cell_type.as_deref() != Some("inlineStr") {
                    return Err(Error::Reverse(format!(
                        "worksheet `{part}` contains an inline string outside a t=inlineStr cell"
                    )));
                }
                let mut buffer = Vec::new();
                let mut writer = XmlWriter::new(&mut buffer);
                writer
                    .write_event(Event::Start(event.into_owned()))
                    .map_err(|error| {
                        Error::Reverse(format!("invalid inline string in `{part}`: {error}"))
                    })?;
                let mut depth = 1_usize;
                while depth > 0 {
                    let event = reader.read_event().map_err(|error| {
                        Error::Reverse(format!("invalid inline string in `{part}`: {error}"))
                    })?;
                    match &event {
                        Event::Start(start) if start.local_name().as_ref() == b"is" => depth += 1,
                        Event::End(end) if end.local_name().as_ref() == b"is" => depth -= 1,
                        _ => {}
                    }
                    writer.write_event(event.into_owned()).map_err(|error| {
                        Error::Reverse(format!("invalid inline string in `{part}`: {error}"))
                    })?;
                }
                let fragment = String::from_utf8(buffer).map_err(|error| {
                    Error::Reverse(format!("invalid inline string UTF-8: {error}"))
                })?;
                if let Some(runs) = parse_rich_text_entries(&fragment, b"is", part)?
                    .into_iter()
                    .next()
                    .flatten()
                {
                    let reference = cell_reference.as_deref().ok_or_else(|| {
                        Error::Reverse(format!(
                            "worksheet `{part}` inline string has no cell reference"
                        ))
                    })?;
                    rich_cells.insert(reference.to_owned(), runs);
                }
            }
            Ok(Event::Start(event))
                if event.local_name().as_ref() == b"v" && cell_type.as_deref() == Some("s") =>
            {
                in_shared_index = true;
            }
            Ok(Event::Text(event)) if in_shared_index => {
                let decoded = event
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                shared_index.push_str(
                    &quick_xml::escape::unescape(&decoded)
                        .map_err(|error| Error::Reverse(error.to_string()))?,
                );
            }
            Ok(Event::GeneralRef(event)) if in_shared_index => {
                if let Some(character) = event
                    .resolve_char_ref()
                    .map_err(|error| Error::Reverse(error.to_string()))?
                {
                    shared_index.push(character);
                } else {
                    let entity = event
                        .decode()
                        .map_err(|error| Error::Reverse(error.to_string()))?;
                    let value = quick_xml::escape::resolve_predefined_entity(&entity)
                        .ok_or_else(|| Error::Reverse(format!("unsupported entity in `{part}`")))?;
                    shared_index.push_str(value);
                }
            }
            Ok(Event::End(event)) => match event.local_name().as_ref() {
                b"v" => in_shared_index = false,
                b"c" => {
                    if cell_type.as_deref() == Some("s") && !shared_index.is_empty() {
                        let index = shared_index.parse::<usize>().map_err(|error| Error::Reverse(format!(
                            "worksheet `{part}` has an invalid shared-string index `{shared_index}`: {error}"
                        )))?;
                        let entry = shared_strings.get(index).ok_or_else(|| Error::Reverse(format!(
                            "worksheet `{part}` references missing shared-string index {index}; repair sharedStrings.xml"
                        )))?;
                        if let Some(runs) = entry {
                            let reference = cell_reference.as_deref().ok_or_else(|| {
                                Error::Reverse(format!(
                                    "worksheet `{part}` rich-text cell has no reference"
                                ))
                            })?;
                            rich_cells.insert(reference.to_owned(), runs.clone());
                        }
                    }
                    cell_reference = None;
                    cell_type = None;
                    shared_index.clear();
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid XML in `{part}`: {error}"))),
        }
    }
    Ok(rich_cells)
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum WorkbookDefinedNameKind {
    FilterDatabase,
    PrintArea,
    PrintTitles,
}

struct WorkbookDefinedName {
    kind: WorkbookDefinedNameKind,
    sheet_index: usize,
    formula: String,
}

struct FilterDatabaseDefinition {
    sheet_index: usize,
    formula: String,
}

#[expect(
    clippy::too_many_lines,
    reason = "Workbook metadata and defined-name entries are parsed in one OOXML state machine."
)]
fn audit_workbook_xml(
    xml: &str,
    sheet_order: &[String],
    worksheet_metadata: &mut HashMap<String, WorksheetMetadata>,
    workbook_properties: &mut WorkbookProperties,
) -> xlsx_jsx::Result<(usize, usize)> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut defined_names = Vec::new();
    let mut current_defined_name: Option<WorkbookDefinedName> = None;
    let mut current_user_defined_name: Option<ReverseDefinedName> = None;
    let mut seen_user_names = std::collections::HashSet::new();
    let mut active_sheet_index = None;
    let mut first_tab_index = None;
    loop {
        match reader.read_event() {
            Ok(Event::Empty(event)) if event.local_name().as_ref() == b"definedName" => {
                return Err(Error::Reverse(
                    "workbook contains an empty defined name; repair or remove the workbook metadata".into(),
                ));
            }
            Ok(Event::Start(event) | Event::Empty(event)) => {
                let tag = event.local_name();
                if tag.as_ref() == b"definedName" {
                    if let Some(user_name) =
                        parse_user_defined_name(&event, reader.decoder(), sheet_order.len())?
                    {
                        current_user_defined_name = Some(user_name);
                    } else {
                        current_defined_name =
                            Some(parse_workbook_defined_name(&event, reader.decoder())?);
                    }
                }
                if tag.as_ref() == b"workbookPr" {
                    for attribute in event.attributes().with_checks(false) {
                        let attribute =
                            attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                        if attribute.key.local_name().as_ref() == b"date1904"
                            && attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .is_ok_and(|value| value == "1" || value == "true")
                        {
                            return Err(Error::Reverse("workbook uses the 1904 date system, which v1 cannot reproduce because the XLSX writer only supports the 1900 date system; convert dates to the 1900 system before reverse or add date-system support".into()));
                        }
                    }
                }
                if tag.as_ref() == b"workbookView" {
                    for attribute in event.attributes().with_checks(false) {
                        let attribute =
                            attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                        if attribute.key.local_name().as_ref() == b"activeTab" {
                            if active_sheet_index.is_some() {
                                return Err(Error::Reverse(
                                    "workbook repeats workbookView.activeTab; remove the duplicate active-sheet metadata"
                                        .into(),
                                ));
                            }
                            let value = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| Error::Reverse(error.to_string()))?;
                            active_sheet_index = Some(value.parse::<usize>().map_err(|error| {
                                Error::Reverse(format!(
                                    "workbook has invalid activeTab `{value}`: {error}"
                                ))
                            })?);
                        }
                        if attribute.key.local_name().as_ref() == b"firstSheet" {
                            if first_tab_index.is_some() {
                                return Err(Error::Reverse(
                                    "workbook repeats workbookView.firstSheet; remove the duplicate first-tab metadata"
                                        .into(),
                                ));
                            }
                            let value = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| Error::Reverse(error.to_string()))?;
                            let one_based = value.parse::<usize>().map_err(|error| {
                                Error::Reverse(format!(
                                    "workbook has invalid firstSheet `{value}`: {error}"
                                ))
                            })?;
                            first_tab_index = Some(one_based.checked_sub(1).ok_or_else(|| {
                                Error::Reverse(
                                    "workbook firstSheet must be a one-based worksheet index; set firstSheet to 1 or greater".into(),
                                )
                            })?);
                        }
                    }
                }
            }
            Ok(Event::Text(text))
                if current_defined_name.is_some() || current_user_defined_name.is_some() =>
            {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                if let Some(definition) = &mut current_defined_name {
                    definition.formula.push_str(&unescaped);
                }
                if let Some(definition) = &mut current_user_defined_name {
                    definition.formula.push_str(&unescaped);
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"definedName" => {
                if let Some(definition) = current_user_defined_name.take() {
                    if definition.hidden {
                        return Err(Error::Reverse(format!(
                            "defined name `{}` is hidden, which v1 cannot preserve; remove the hidden flag or extend DefinedName metadata support",
                            definition.name
                        )));
                    }
                    if definition.formula.trim().is_empty() {
                        return Err(Error::Reverse(format!(
                            "defined name `{}` has an empty formula; provide a non-empty formula or remove the defined name",
                            definition.name
                        )));
                    }
                    if !seen_user_names
                        .insert((definition.sheet_index, definition.name.to_lowercase()))
                    {
                        return Err(Error::Reverse(format!(
                            "workbook repeats defined name `{}` in the same scope; rename or remove the duplicate",
                            definition.name
                        )));
                    }
                    if let Some(index) = definition.sheet_index {
                        let sheet_name = sheet_order.get(index).ok_or_else(|| Error::Reverse(format!("defined name `{}` references missing worksheet index {index}; repair localSheetId", definition.name)))?;
                        worksheet_metadata
                            .get_mut(sheet_name)
                            .ok_or_else(|| {
                                Error::Reverse(format!(
                                    "missing worksheet metadata for `{sheet_name}`"
                                ))
                            })?
                            .defined_names
                            .push(definition);
                    } else {
                        workbook_properties.defined_names.push(definition);
                    }
                } else {
                    let definition = current_defined_name.take().ok_or_else(|| {
                        Error::Reverse("workbook has a malformed defined-name entry".into())
                    })?;
                    defined_names.push(definition);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid workbook.xml: {error}"))),
        }
    }
    validate_workbook_defined_names(&defined_names, sheet_order, worksheet_metadata)?;
    Ok((
        active_sheet_index.unwrap_or_default(),
        first_tab_index.unwrap_or_default(),
    ))
}

#[derive(Debug, Default)]
struct ReverseDefinedName {
    name: String,
    formula: String,
    sheet_index: Option<usize>,
    hidden: bool,
}

fn parse_user_defined_name(
    event: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    sheet_count: usize,
) -> xlsx_jsx::Result<Option<ReverseDefinedName>> {
    let mut name = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() == b"name" {
            name = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| Error::Reverse(error.to_string()))?
                    .into_owned(),
            );
            break;
        }
    }
    let Some(name) = name else {
        return Ok(None);
    };
    if name.starts_with("_xlnm.") {
        return Ok(None);
    }
    if name.is_empty() || rust_xlsxwriter::utility::check_name(&name).is_err() {
        return Err(Error::Reverse(format!(
            "defined name `{name}` is not a valid Excel name that this DSL can rebuild; rename it to a valid name or remove it"
        )));
    }
    let mut local_sheet_id = None;
    let mut hidden = false;
    let mut seen = std::collections::HashSet::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        if !seen.insert(key.as_ref().to_vec()) {
            return Err(Error::Reverse(format!(
                "defined name `{name}` repeats an attribute; remove duplicate metadata"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match key.as_ref() {
            b"name" => {}
            b"localSheetId" => {
                local_sheet_id = Some(value.parse::<usize>().map_err(|error| {
                    Error::Reverse(format!(
                        "defined name `{name}` has invalid localSheetId: {error}"
                    ))
                })?);
            }
            b"hidden" => hidden = parse_xml_bool(&value, "xl/workbook.xml", "definedName.hidden")?,
            key => {
                return Err(Error::Reverse(format!(
                    "defined name `{name}` has unsupported attribute `{}`; remove it or add matching DSL metadata support",
                    String::from_utf8_lossy(key)
                )));
            }
        }
    }
    if local_sheet_id.is_some_and(|index| index >= sheet_count) {
        return Err(Error::Reverse(format!(
            "defined name `{name}` has localSheetId outside the worksheet list; repair localSheetId"
        )));
    }
    Ok(Some(ReverseDefinedName {
        name,
        formula: String::new(),
        sheet_index: local_sheet_id,
        hidden,
    }))
}

fn parse_workbook_defined_name(
    event: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<WorkbookDefinedName> {
    let mut name = None;
    let mut sheet_index = None;
    let mut hidden = false;
    let mut seen_attributes = std::collections::HashSet::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if !seen_attributes.insert(attribute.key.local_name().as_ref().to_vec()) {
            return Err(Error::Reverse(
                "workbook defined name repeats an attribute; remove the duplicate metadata".into(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match attribute.key.local_name().as_ref() {
            b"name" => name = Some(value.into_owned()),
            b"localSheetId" => {
                sheet_index = Some(value.parse::<usize>().map_err(|error| {
                    Error::Reverse(format!(
                        "workbook has invalid defined-name localSheetId: {error}"
                    ))
                })?);
            }
            b"hidden" => hidden = parse_xml_bool(&value, "xl/workbook.xml", "hidden")?,
            key => {
                return Err(Error::Reverse(format!(
                    "workbook contains unsupported defined-name attribute `{}`; remove the defined name or add matching DSL support",
                    String::from_utf8_lossy(key)
                )));
            }
        }
    }
    let kind = match name.as_deref() {
        Some("_xlnm._FilterDatabase") => WorkbookDefinedNameKind::FilterDatabase,
        Some("_xlnm.Print_Area") => WorkbookDefinedNameKind::PrintArea,
        Some("_xlnm.Print_Titles") => WorkbookDefinedNameKind::PrintTitles,
        _ => return Err(Error::Reverse(
            "workbook contains an unsupported defined name; remove it or add matching DSL support"
                .into(),
        )),
    };
    let sheet_index = sheet_index.ok_or_else(|| {
        Error::Reverse("workbook contains a global defined name that v1 cannot preserve; supported auto-filter and print names must be worksheet-local".into())
    })?;
    if kind == WorkbookDefinedNameKind::FilterDatabase && !hidden {
        return Err(Error::Reverse("workbook auto-filter defined name must be hidden; repair `_xlnm._FilterDatabase` metadata".into()));
    }
    if kind != WorkbookDefinedNameKind::FilterDatabase && hidden {
        return Err(Error::Reverse(
            "workbook print-area/title defined names cannot be hidden in the supported subset; make them visible or remove the hidden flag".into(),
        ));
    }
    Ok(WorkbookDefinedName {
        kind,
        sheet_index,
        formula: String::new(),
    })
}

fn validate_workbook_defined_names(
    definitions: &[WorkbookDefinedName],
    sheet_order: &[String],
    worksheet_metadata: &mut HashMap<String, WorksheetMetadata>,
) -> xlsx_jsx::Result<()> {
    let mut seen = std::collections::HashSet::new();
    let mut filter_definitions = Vec::new();
    for definition in definitions {
        let sheet_name = sheet_order.get(definition.sheet_index).ok_or_else(|| {
            Error::Reverse(format!(
                "workbook defined name references missing worksheet index {}; repair localSheetId",
                definition.sheet_index
            ))
        })?;
        if !seen.insert((definition.sheet_index, definition.kind)) {
            return Err(Error::Reverse(format!(
                "workbook repeats a worksheet-local defined name for `{sheet_name}`; keep only one print area, print-title set, or auto-filter name per worksheet"
            )));
        }
        if definition.formula.trim().is_empty() {
            return Err(Error::Reverse(format!(
                "workbook defined name for `{sheet_name}` has an empty formula; repair or remove it"
            )));
        }
        match definition.kind {
            WorkbookDefinedNameKind::FilterDatabase => {
                filter_definitions.push(FilterDatabaseDefinition {
                    sheet_index: definition.sheet_index,
                    formula: definition.formula.clone(),
                });
            }
            WorkbookDefinedNameKind::PrintArea => {
                let (formula_sheet, reference) = parse_sheet_range_reference(&definition.formula)?;
                if formula_sheet != *sheet_name {
                    return Err(Error::Reverse(format!(
                        "workbook print area for `{sheet_name}` points at `{formula_sheet}`; align the range sheet with localSheetId"
                    )));
                }
                let range = parse_print_range(reference).ok_or_else(|| {
                    Error::Reverse(format!(
                        "workbook print area `{}` is outside the supported A1 cell/row/column range forms; repair the formula",
                        definition.formula
                    ))
                })?;
                let metadata = worksheet_metadata.get_mut(sheet_name).ok_or_else(|| {
                    Error::Reverse(format!("missing worksheet metadata for `{sheet_name}`"))
                })?;
                metadata.print_area = Some(format_print_range(range));
            }
            WorkbookDefinedNameKind::PrintTitles => {
                let mut parsed_rows = None;
                let mut parsed_columns = None;
                for reference in split_defined_name_formula(&definition.formula)? {
                    let (formula_sheet, range_text) = parse_sheet_range_reference(reference)?;
                    if formula_sheet != *sheet_name {
                        return Err(Error::Reverse(format!(
                            "workbook print titles for `{sheet_name}` point at `{formula_sheet}`; align every range sheet with localSheetId"
                        )));
                    }
                    match parse_print_range(range_text) {
                        Some(PrintRange::Rows { .. }) if parsed_rows.is_none() => {
                            parsed_rows = parse_print_range(range_text);
                        }
                        Some(PrintRange::Columns { .. }) if parsed_columns.is_none() => {
                            parsed_columns = parse_print_range(range_text);
                        }
                        _ => {
                            return Err(Error::Reverse(format!(
                                "workbook print-title range `{range_text}` must be one unique whole-row range and/or one unique whole-column range"
                            )));
                        }
                    }
                }
                let metadata = worksheet_metadata.get_mut(sheet_name).ok_or_else(|| {
                    Error::Reverse(format!("missing worksheet metadata for `{sheet_name}`"))
                })?;
                if let Some(PrintRange::Rows { first, last }) = parsed_rows {
                    metadata.repeat_rows =
                        Some(format_print_range(PrintRange::Rows { first, last }));
                }
                if let Some(PrintRange::Columns { first, last }) = parsed_columns {
                    metadata.repeat_columns =
                        Some(format_print_range(PrintRange::Columns { first, last }));
                }
            }
        }
    }
    validate_filter_database_definitions(&filter_definitions, sheet_order, worksheet_metadata)
}

fn split_defined_name_formula(formula: &str) -> xlsx_jsx::Result<Vec<&str>> {
    let bytes = formula.as_bytes();
    let mut ranges = Vec::new();
    let mut start = 0;
    let mut index = 0;
    let mut quoted = false;
    while index < bytes.len() {
        match bytes[index] {
            b'\'' if quoted && bytes.get(index + 1) == Some(&b'\'') => index += 2,
            b'\'' => {
                quoted = !quoted;
                index += 1;
            }
            b',' if !quoted => {
                ranges.push(formula[start..index].trim());
                start = index + 1;
                index += 1;
            }
            _ => index += 1,
        }
    }
    if quoted {
        return Err(Error::Reverse(
            "workbook defined-name formula contains an unterminated quoted worksheet name; repair the formula".into(),
        ));
    }
    ranges.push(formula[start..].trim());
    if ranges.is_empty() || ranges.iter().any(|range| range.is_empty()) || ranges.len() > 2 {
        return Err(Error::Reverse(
            "workbook print-title formula must contain one row range and/or one column range; repair the formula".into(),
        ));
    }
    Ok(ranges)
}

fn parse_sheet_range_reference(formula: &str) -> xlsx_jsx::Result<(String, &str)> {
    let (sheet, range) = formula.rsplit_once('!').ok_or_else(|| {
        Error::Reverse(format!(
            "workbook defined-name formula `{formula}` must qualify its A1 range with a worksheet name"
        ))
    })?;
    if sheet.starts_with('\'') {
        let unquoted = sheet
            .strip_prefix('\'')
            .and_then(|value| value.strip_suffix('\''))
            .ok_or_else(|| {
                Error::Reverse(format!(
                    "workbook defined-name formula has malformed worksheet quoting `{sheet}`"
                ))
            })?;
        Ok((unquoted.replace("''", "'"), range))
    } else if sheet.contains('\'') || sheet.is_empty() {
        Err(Error::Reverse(format!(
            "workbook defined-name formula has invalid worksheet name `{sheet}`"
        )))
    } else {
        Ok((sheet.to_owned(), range))
    }
}

fn validate_filter_database_definitions(
    definitions: &[FilterDatabaseDefinition],
    sheet_order: &[String],
    worksheet_metadata: &HashMap<String, WorksheetMetadata>,
) -> xlsx_jsx::Result<()> {
    let mut seen = std::collections::HashSet::new();
    for definition in definitions {
        let Some(sheet_name) = sheet_order.get(definition.sheet_index) else {
            return Err(Error::Reverse(
                "workbook auto-filter defined name references a missing worksheet; repair localSheetId".into(),
            ));
        };
        let expected = worksheet_metadata
            .get(sheet_name)
            .and_then(|metadata| metadata.auto_filter.as_deref())
            .ok_or_else(|| {
                Error::Reverse(format!(
                    "workbook contains an auto-filter defined name for `{sheet_name}` without a matching worksheet filter; repair the workbook metadata"
                ))
            })?;
        let (formula_sheet, formula_range) =
            definition.formula.rsplit_once('!').ok_or_else(|| {
                Error::Reverse(
                    "workbook auto-filter defined name has an invalid range formula; repair it"
                        .into(),
                )
            })?;
        let formula_sheet = formula_sheet
            .strip_prefix('\'')
            .and_then(|value| value.strip_suffix('\''))
            .map_or_else(
                || formula_sheet.to_owned(),
                |value| value.replace("''", "'"),
            );
        let parsed_range = parse_a1_range(formula_range).ok_or_else(|| {
            Error::Reverse(
                "workbook auto-filter defined name has an invalid A1 range; repair it".into(),
            )
        })?;
        if formula_sheet != *sheet_name
            || format_a1_range(parsed_range) != expected
            || !seen.insert(definition.sheet_index)
        {
            return Err(Error::Reverse(format!(
                "workbook auto-filter defined name does not match Worksheet `{sheet_name}` autoFilter `{expected}`; repair the range, sheet reference, or localSheetId"
            )));
        }
    }
    Ok(())
}

#[derive(Debug, Default)]
struct WorkbookProperties {
    title: Option<String>,
    subject: Option<String>,
    author: Option<String>,
    keywords: Option<String>,
    description: Option<String>,
    defined_names: Vec<ReverseDefinedName>,
}

fn audit_core_properties(xml: &str) -> xlsx_jsx::Result<WorkbookProperties> {
    let mut reader = XmlReader::from_str(xml);
    let mut current_property = None;
    let mut properties = WorkbookProperties::default();
    let mut last_modified_by = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) => {
                current_property =
                    Some(String::from_utf8_lossy(event.local_name().as_ref()).into_owned());
            }
            Ok(Event::Text(text)) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Reverse(error.to_string()))?;
                let value = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Reverse(error.to_string()))?
                    .into_owned();
                if !value.is_empty() {
                    let field = match current_property.as_deref() {
                        Some("title") => Some(&mut properties.title),
                        Some("subject") => Some(&mut properties.subject),
                        Some("creator") => Some(&mut properties.author),
                        Some("keywords") => Some(&mut properties.keywords),
                        Some("description") => Some(&mut properties.description),
                        Some("lastModifiedBy") => Some(&mut last_modified_by),
                        Some(name) if is_authored_core_property(name) => {
                            return Err(Error::Reverse(format!(
                                "workbook contains authored document metadata `{name}` that v1 cannot preserve; remove that property or add matching Workbook metadata support"
                            )));
                        }
                        _ => None,
                    };
                    if let Some(field) = field {
                        if let Some(existing) = field.as_mut() {
                            existing.push_str(&value);
                        } else {
                            *field = Some(value);
                        }
                    }
                }
            }
            Ok(Event::GeneralRef(reference)) => {
                let value = reference
                    .resolve_char_ref()
                    .map_err(|error| Error::Reverse(error.to_string()))?
                    .or_else(|| match &reference[..] {
                        b"amp" => Some('&'),
                        b"lt" => Some('<'),
                        b"gt" => Some('>'),
                        b"quot" => Some('\"'),
                        b"apos" => Some('\''),
                        _ => None,
                    })
                    .ok_or_else(|| {
                        Error::Reverse(format!(
                            "docProps/core.xml contains unsupported entity `&{};`",
                            String::from_utf8_lossy(reference.as_ref())
                        ))
                    })?;
                match current_property.as_deref() {
                    Some("title") => {
                        properties.title.get_or_insert_default().push(value);
                    }
                    Some("subject") => {
                        properties.subject.get_or_insert_default().push(value);
                    }
                    Some("creator") => {
                        properties.author.get_or_insert_default().push(value);
                    }
                    Some("keywords") => {
                        properties.keywords.get_or_insert_default().push(value);
                    }
                    Some("description") => {
                        properties.description.get_or_insert_default().push(value);
                    }
                    Some("lastModifiedBy") => {
                        last_modified_by.get_or_insert_default().push(value);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(_)) => current_property = None,
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid docProps/core.xml: {error}"
                )));
            }
        }
    }
    if properties.author.as_ref() != last_modified_by.as_ref()
        && (properties.author.is_some() || last_modified_by.is_some())
    {
        return Err(Error::Reverse(
            "docProps/core.xml creator and lastModifiedBy differ, which Workbook.author cannot represent; make them identical or add separate author properties to the DSL".into(),
        ));
    }
    Ok(properties)
}

fn is_authored_core_property(name: &str) -> bool {
    matches!(
        name,
        "title"
            | "subject"
            | "creator"
            | "description"
            | "keywords"
            | "category"
            | "contentStatus"
            | "identifier"
            | "language"
            | "lastModifiedBy"
            | "revision"
            | "version"
    )
}

#[derive(Default)]
#[expect(
    clippy::struct_excessive_bools,
    reason = "These flags are independent levels in the streaming style XML parser state machine."
)]
struct StyleParser {
    fill_formats: Vec<FillFormat>,
    border_formats: Vec<Option<CellBorders>>,
    fonts: Vec<CellFontStyle>,
    current_font: Option<CellFontStyle>,
    in_cell_xfs: bool,
    in_xf_extension_list: bool,
    in_xf_extension: bool,
    checkbox_complement_seen: bool,
    formats: Vec<CellFontStyle>,
    number_formats: HashMap<usize, String>,
}

impl StyleParser {
    #[expect(
        clippy::too_many_lines,
        reason = "The parser dispatches each nested style element to its strict, context-aware handler."
    )]
    fn parse_start(
        &mut self,
        event: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
    ) -> xlsx_jsx::Result<()> {
        let tag = event.local_name();
        if self.in_cell_xfs
            && tag.as_ref() != b"xf"
            && tag.as_ref() != b"cellXfs"
            && tag.as_ref() != b"alignment"
            && tag.as_ref() != b"protection"
            && tag.as_ref() != b"extLst"
            && !(self.in_xf_extension_list
                && (tag.as_ref() == b"ext" || tag.as_ref() == b"xfComplement"))
            && let Some(style) = self.formats.last_mut()
        {
            style.support = StyleSupport::Unsupported;
        }
        match tag.as_ref() {
            b"font" => {
                self.current_font = Some(CellFontStyle {
                    support: StyleSupport::Supported,
                    ..Default::default()
                });
            }
            b"cellXfs" => self.in_cell_xfs = true,
            b"extLst" if self.in_cell_xfs => {
                self.in_xf_extension_list = true;
            }
            b"ext" if self.in_cell_xfs && self.in_xf_extension_list => {
                let mut uri = None;
                let mut namespace = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    let key =
                        String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                        .map_err(|error| Error::Reverse(error.to_string()))?;
                    match key.as_str() {
                        "uri" => uri = Some(value.into_owned()),
                        "xfpb" => namespace = Some(value.into_owned()),
                        _ => {
                            if let Some(style) = self.formats.last_mut() {
                                style.support = StyleSupport::Unsupported;
                            }
                        }
                    }
                }
                if (uri.as_deref() != Some("{C7286773-470A-42A8-94C5-96B5CB345126}")
                    || namespace.as_deref()
                        != Some(
                            "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag",
                        ))
                    && let Some(style) = self.formats.last_mut()
                {
                    style.support = StyleSupport::Unsupported;
                }
                self.in_xf_extension = true;
            }
            b"xfComplement" if self.in_cell_xfs && self.in_xf_extension => {
                let mut index = None;
                for attribute in event.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
                    if attribute.key.local_name().as_ref() == b"i" {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                            .map_err(|error| Error::Reverse(error.to_string()))?;
                        index = Some(value.into_owned());
                    } else if let Some(style) = self.formats.last_mut() {
                        style.support = StyleSupport::Unsupported;
                    }
                }
                if index.as_deref() == Some("0") {
                    if let Some(style) = self.formats.last_mut() {
                        if self.checkbox_complement_seen {
                            style.support = StyleSupport::Unsupported;
                        }
                        style.checkbox = true;
                    }
                    self.checkbox_complement_seen = true;
                } else if let Some(style) = self.formats.last_mut() {
                    style.support = StyleSupport::Unsupported;
                }
            }
            b"numFmt" if self.current_font.is_none() => {
                if let Some((id, code)) = parse_custom_number_format(event, decoder)? {
                    self.number_formats.insert(id, code);
                }
            }
            b"b" | b"i" | b"strike" | b"u" | b"vertAlign" | b"color"
                if self.current_font.is_some() =>
            {
                if let Some(style) = &mut self.current_font {
                    parse_font_component(event, tag.as_ref(), style, decoder)?;
                }
            }
            b"sz" | b"name" if self.current_font.is_some() => {
                if let Some(style) = &mut self.current_font {
                    parse_font_name_or_size(
                        event,
                        tag.as_ref(),
                        style,
                        decoder,
                        self.fonts.is_empty(),
                    )?;
                }
            }
            b"scheme" if self.current_font.is_some() => {
                if let Some(style) = &mut self.current_font {
                    parse_font_scheme(event, style, decoder, self.fonts.is_empty())?;
                }
            }
            b"family" | b"charset" if self.current_font.is_some() => {
                if let Some(style) = &mut self.current_font {
                    parse_font_family_or_charset(event, tag.as_ref(), style, decoder)?;
                }
            }
            b"fonts" | b"styleSheet" | b"numFmts" | b"numFmt" | b"fills" | b"fill"
            | b"patternFill" | b"borders" | b"border" | b"left" | b"right" | b"top" | b"bottom"
            | b"diagonal" | b"cellStyleXfs" | b"cellStyles" | b"cellStyle" | b"dxfs"
            | b"tableStyles"
                if self.current_font.is_some() =>
            {
                if let Some(style) = &mut self.current_font {
                    audit_default_font_component(event, tag.as_ref(), style, decoder)?;
                }
            }
            b"xf" if self.in_cell_xfs => {
                self.checkbox_complement_seen = false;
                self.formats.push(parse_cell_format(
                    event,
                    &self.fonts,
                    &self.number_formats,
                    &self.fill_formats,
                    &self.border_formats,
                    decoder,
                )?);
            }
            b"alignment" if self.in_cell_xfs => {
                if let Some(style) = self.formats.last_mut() {
                    parse_cell_alignment(event, style, decoder)?;
                }
            }
            b"protection" if self.in_cell_xfs => {
                if let Some(style) = self.formats.last_mut() {
                    parse_cell_protection(event, style, decoder)?;
                }
            }
            b"fonts" | b"numFmts" | b"numFmt" | b"fills" | b"fill" | b"patternFill"
            | b"borders" | b"border" | b"left" | b"right" | b"top" | b"bottom" | b"diagonal"
            | b"cellStyleXfs" | b"xf" | b"cellStyles" | b"cellStyle" | b"dxfs" | b"tableStyles"
            | b"sz" | b"color" | b"name" | b"family" | b"charset" | b"scheme" | b"alignment"
            | b"protection" => {}
            _ => {
                if let Some(style) = &mut self.current_font {
                    style.support = StyleSupport::Unsupported;
                }
            }
        }
        Ok(())
    }

    fn parse_end(&mut self, tag: &[u8]) {
        match tag {
            b"font" => {
                if let Some(mut style) = self.current_font.take() {
                    if !self.fonts.is_empty() && style.font_scheme.is_none() {
                        style.font_scheme = Some("none".to_owned());
                    }
                    self.fonts.push(finalize_font_style(style));
                }
            }
            b"cellXfs" => self.in_cell_xfs = false,
            b"ext" => self.in_xf_extension = false,
            b"extLst" => self.in_xf_extension_list = false,
            _ => {}
        }
    }
}

fn parse_styles_xml(xml: &str) -> xlsx_jsx::Result<Vec<CellFontStyle>> {
    let mut parser = StyleParser {
        fill_formats: parse_fill_formats(xml)?,
        border_formats: parse_border_formats(xml)?,
        ..Default::default()
    };
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event)) => {
                parser.parse_start(&event, reader.decoder())?;
            }
            Ok(Event::End(event)) => parser.parse_end(event.local_name().as_ref()),
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid styles.xml: {error}"))),
        }
    }
    if parser.formats.is_empty() {
        return Err(Error::Reverse("styles.xml has no cell formats".into()));
    }
    Ok(parser.formats)
}

#[expect(
    clippy::too_many_lines,
    reason = "The dxf reader validates nested font/fill children as one XML state machine."
)]
fn parse_differential_styles_xml(xml: &str) -> xlsx_jsx::Result<Vec<CellFontStyle>> {
    let differential_borders = parse_differential_border_formats(xml)?;
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut in_dxfs = false;
    let mut in_dxf = false;
    let mut in_font = false;
    let mut in_fill = false;
    let mut font_seen = false;
    let mut fill_seen = false;
    let mut pattern_seen = false;
    let mut foreground_seen = false;
    let mut background_seen = false;
    let mut number_format_seen = false;
    let mut in_border = false;
    let mut formats = Vec::new();
    let mut current: Option<CellFontStyle> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event) | Event::Empty(event)) => {
                let tag = event.local_name();
                if tag.as_ref() == b"dxfs" {
                    in_dxfs = true;
                    continue;
                }
                if !in_dxfs {
                    continue;
                }
                if tag.as_ref() == b"dxf" {
                    if in_dxf {
                        return Err(Error::Reverse(
                            "styles.xml contains nested differential styles".into(),
                        ));
                    }
                    in_dxf = true;
                    font_seen = false;
                    fill_seen = false;
                    pattern_seen = false;
                    foreground_seen = false;
                    background_seen = false;
                    number_format_seen = false;
                    in_border = false;
                    current = Some(CellFontStyle {
                        support: StyleSupport::Supported,
                        ..Default::default()
                    });
                    continue;
                }
                if !in_dxf {
                    continue;
                }
                let Some(style) = current.as_mut() else {
                    continue;
                };
                match tag.as_ref() {
                    b"border" if in_dxf => in_border = true,
                    b"top" | b"right" | b"bottom" | b"left" | b"diagonal" | b"vertical"
                    | b"horizontal" | b"color"
                        if in_border => {}
                    b"font" => {
                        if font_seen || event.attributes().next().is_some() {
                            style.support = StyleSupport::Unsupported;
                        }
                        font_seen = true;
                        in_font = true;
                    }
                    b"fill" => {
                        if fill_seen || event.attributes().next().is_some() {
                            style.support = StyleSupport::Unsupported;
                        }
                        fill_seen = true;
                        in_fill = true;
                    }
                    b"b" | b"i" if in_font => {
                        parse_font_component(&event, tag.as_ref(), style, reader.decoder())?;
                        if (tag.as_ref() == b"b" && !style.bold)
                            || (tag.as_ref() == b"i" && !style.italic)
                        {
                            style.support = StyleSupport::Unsupported;
                        }
                    }
                    b"strike" | b"u" if in_font => {
                        parse_font_component(&event, tag.as_ref(), style, reader.decoder())?;
                        if tag.as_ref() == b"strike" && style.strikethrough != Some(true) {
                            style.support = StyleSupport::Unsupported;
                        }
                        if tag.as_ref() == b"u" && style.underline.is_none() {
                            style.support = StyleSupport::Unsupported;
                        }
                    }
                    b"vertAlign" if in_font => {
                        parse_font_component(&event, tag.as_ref(), style, reader.decoder())?;
                        if style.script.is_none() {
                            style.support = StyleSupport::Unsupported;
                        }
                    }
                    b"color" if in_font => {
                        parse_font_color(&event, style, reader.decoder())?;
                    }
                    b"numFmt" if in_dxf => {
                        if number_format_seen {
                            style.support = StyleSupport::Unsupported;
                        }
                        number_format_seen = true;
                        style.number_format =
                            parse_differential_number_format(&event, reader.decoder())?;
                    }
                    b"patternFill" if in_fill => {
                        if pattern_seen {
                            style.support = StyleSupport::Unsupported;
                        }
                        pattern_seen = true;
                        match parse_differential_pattern_type(&event, reader.decoder()) {
                            Some(pattern) => style.fill_pattern = Some(pattern),
                            None => style.support = StyleSupport::Unsupported,
                        }
                    }
                    b"fgColor" if in_fill => {
                        if foreground_seen {
                            style.support = StyleSupport::Unsupported;
                        }
                        foreground_seen = true;
                        match parse_fill_rgb(&event, reader.decoder()) {
                            Some(color) => style.fill_foreground_color = Some(color),
                            None => style.support = StyleSupport::Unsupported,
                        }
                    }
                    b"bgColor" if in_fill => {
                        if background_seen {
                            style.support = StyleSupport::Unsupported;
                        }
                        background_seen = true;
                        if !is_default_fill_background(&event, reader.decoder()) {
                            match parse_fill_rgb(&event, reader.decoder()) {
                                Some(color) => style.fill_background_color = Some(color),
                                _ => style.support = StyleSupport::Unsupported,
                            }
                        }
                    }
                    b"dxfs" | b"dxf" => {}
                    _ => style.support = StyleSupport::Unsupported,
                }
            }
            Ok(Event::End(event)) => match event.local_name().as_ref() {
                b"font" => in_font = false,
                b"fill" => in_fill = false,
                b"border" => in_border = false,
                b"dxf" => {
                    let Some(mut style) = current.take() else {
                        continue;
                    };
                    if let Some(Some(borders)) = differential_borders.get(formats.len()) {
                        style.borders = borders.clone();
                        if !style.borders.has_any()
                            || style.borders.diagonal.style.is_some()
                            || style.borders.diagonal.color.is_some()
                        {
                            style.support = StyleSupport::Unsupported;
                        }
                    }
                    let has_supported_style = style.bold
                        || style.italic
                        || style.strikethrough == Some(true)
                        || style.underline.is_some()
                        || style.script.is_some()
                        || style.font_color.is_some()
                        || style.fill_color.is_some()
                        || style.fill_pattern.is_some()
                        || style.fill_foreground_color.is_some()
                        || style.fill_background_color.is_some()
                        || style.number_format.is_some()
                        || style.borders.has_any();
                    if !has_supported_style {
                        style.support = StyleSupport::Unsupported;
                    }
                    if fill_seen && (!pattern_seen || (!foreground_seen && !background_seen)) {
                        style.support = StyleSupport::Unsupported;
                    }
                    if style.fill_pattern.as_deref() == Some("solid") {
                        let canonical_color = match (
                            style.fill_foreground_color.take(),
                            style.fill_background_color.take(),
                        ) {
                            (Some(color), None) | (None, Some(color)) => Some(color),
                            (foreground, background) => {
                                style.fill_foreground_color = foreground;
                                style.fill_background_color = background;
                                None
                            }
                        };
                        if let Some(color) = canonical_color {
                            style.fill_color = Some(color);
                            style.fill_pattern = None;
                        }
                    }
                    formats.push(style);
                    in_dxf = false;
                    in_font = false;
                    in_fill = false;
                    in_border = false;
                }
                b"dxfs" => in_dxfs = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid styles.xml dxf section: {error}"
                )));
            }
        }
    }
    if in_dxf || current.is_some() {
        return Err(Error::Reverse(
            "styles.xml contains an unfinished differential style".into(),
        ));
    }
    Ok(formats)
}

#[derive(Clone)]
enum FillFormat {
    None,
    Solid(String),
    Pattern {
        pattern: String,
        foreground: Option<String>,
        background: Option<String>,
    },
    Unsupported,
}

#[derive(Default)]
#[expect(
    clippy::struct_excessive_bools,
    reason = "The parser tracks independent seen/valid state for four edges plus two OOXML diagonal flags."
)]
struct BorderBuilder {
    borders: CellBorders,
    valid: bool,
    current_edge: Option<BorderXmlEdge>,
    edge_seen: [bool; 4],
    color_seen: [bool; 4],
    diagonal_seen: bool,
    diagonal_color_seen: bool,
    diagonal_up: bool,
    diagonal_down: bool,
}

#[derive(Clone, Copy)]
enum BorderXmlEdge {
    Side(BorderSide),
    Diagonal,
}

impl BorderBuilder {
    fn new() -> Self {
        Self {
            valid: true,
            ..Default::default()
        }
    }

    fn parse_border_attributes(
        &mut self,
        event: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
    ) {
        let mut up_seen = false;
        let mut down_seen = false;
        for attribute in event.attributes().with_checks(false) {
            let Ok(attribute) = attribute else {
                self.valid = false;
                continue;
            };
            let key = attribute.key.local_name();
            let Ok(value) =
                attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            else {
                self.valid = false;
                continue;
            };
            match key.as_ref() {
                b"diagonalUp" if !up_seen => {
                    up_seen = true;
                    if let Some(enabled) = parse_ooxml_bool(&value) {
                        self.diagonal_up = enabled;
                    } else {
                        self.valid = false;
                    }
                }
                b"diagonalDown" if !down_seen => {
                    down_seen = true;
                    if let Some(enabled) = parse_ooxml_bool(&value) {
                        self.diagonal_down = enabled;
                    } else {
                        self.valid = false;
                    }
                }
                b"outline" if value == "1" => {}
                _ => self.valid = false,
            }
        }
    }

    fn parse_differential_border_attributes(
        &mut self,
        event: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
    ) {
        for attribute in event.attributes().with_checks(false) {
            let Ok(attribute) = attribute else {
                self.valid = false;
                continue;
            };
            if ![b"diagonalUp".as_slice(), b"diagonalDown".as_slice()]
                .contains(&attribute.key.local_name().as_ref())
            {
                self.valid = false;
            }
        }
        self.parse_border_attributes(event, decoder);
    }

    fn add_edge(
        &mut self,
        event: &BytesStart<'_>,
        tag: &[u8],
        decoder: quick_xml::encoding::Decoder,
    ) {
        let edge = match tag {
            b"top" => Some(BorderXmlEdge::Side(BorderSide::Top)),
            b"right" => Some(BorderXmlEdge::Side(BorderSide::Right)),
            b"bottom" => Some(BorderXmlEdge::Side(BorderSide::Bottom)),
            b"left" => Some(BorderXmlEdge::Side(BorderSide::Left)),
            b"diagonal" => Some(BorderXmlEdge::Diagonal),
            _ => None,
        };
        let Some(edge) = edge else {
            self.valid = false;
            return;
        };
        if self.current_edge.is_some() {
            self.valid = false;
        }
        match edge {
            BorderXmlEdge::Side(side) => {
                let index = side.index();
                if self.edge_seen[index] {
                    self.valid = false;
                }
                self.edge_seen[index] = true;
                self.current_edge = Some(edge);
                self.parse_edge_attributes(event, side, decoder);
            }
            BorderXmlEdge::Diagonal => {
                if self.diagonal_seen {
                    self.valid = false;
                }
                self.diagonal_seen = true;
                self.current_edge = Some(edge);
                let mut style_seen = false;
                for attribute in event.attributes().with_checks(false) {
                    let Ok(attribute) = attribute else {
                        self.valid = false;
                        continue;
                    };
                    if attribute.key.local_name().as_ref() != b"style" || style_seen {
                        self.valid = false;
                        continue;
                    }
                    style_seen = true;
                    let Ok(value) =
                        attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    else {
                        self.valid = false;
                        continue;
                    };
                    match parse_border_style(&value) {
                        Some(FormatBorder::None) => self.borders.diagonal.style = None,
                        Some(style) => self.borders.diagonal.style = Some(style),
                        None => self.valid = false,
                    }
                }
            }
        }
    }

    fn parse_edge_attributes(
        &mut self,
        event: &BytesStart<'_>,
        side: BorderSide,
        decoder: quick_xml::encoding::Decoder,
    ) {
        let edge = &mut self.borders.edges[side.index()];
        let mut style_seen = false;
        for attribute in event.attributes().with_checks(false) {
            let Ok(attribute) = attribute else {
                self.valid = false;
                continue;
            };
            if attribute.key.local_name().as_ref() != b"style" || style_seen {
                self.valid = false;
                continue;
            }
            style_seen = true;
            let Ok(value) =
                attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            else {
                self.valid = false;
                continue;
            };
            match parse_border_style(&value) {
                Some(FormatBorder::None) => edge.style = None,
                Some(style) => edge.style = Some(style),
                None => self.valid = false,
            }
        }
    }

    fn add_color(&mut self, event: &BytesStart<'_>, decoder: quick_xml::encoding::Decoder) {
        let edge = match self.current_edge {
            Some(BorderXmlEdge::Side(side)) => {
                if self.color_seen[side.index()] {
                    self.valid = false;
                    return;
                }
                self.color_seen[side.index()] = true;
                &mut self.borders.edges[side.index()]
            }
            Some(BorderXmlEdge::Diagonal) => {
                if self.diagonal_color_seen {
                    self.valid = false;
                    return;
                }
                self.diagonal_color_seen = true;
                &mut self.borders.diagonal
            }
            None => {
                self.valid = false;
                return;
            }
        };
        match parse_border_color(event, decoder) {
            Some(BorderColor::Rgb(color)) => edge.color = Some(color),
            Some(BorderColor::Automatic) => {}
            None => self.valid = false,
        }
    }

    fn finish(mut self) -> Option<CellBorders> {
        if self.current_edge.is_some() {
            self.valid = false;
        }
        if self
            .borders
            .edges
            .iter()
            .any(|edge| edge.color.is_some() && edge.style.is_none())
        {
            self.valid = false;
        }
        if self.borders.diagonal.color.is_some() && self.borders.diagonal.style.is_none() {
            self.valid = false;
        }
        self.borders.diagonal_type = match (self.diagonal_up, self.diagonal_down) {
            (true, false) => Some("up".into()),
            (false, true) => Some("down".into()),
            (true, true) => Some("upDown".into()),
            (false, false) => None,
        };
        self.valid.then_some(self.borders)
    }
}

#[derive(Default)]
struct DifferentialBorderParser {
    in_dxfs: bool,
    in_dxf: bool,
    border_seen: bool,
    builder: Option<BorderBuilder>,
    border: Option<CellBorders>,
    formats: Vec<Option<CellBorders>>,
}

impl DifferentialBorderParser {
    fn start(
        &mut self,
        event: &BytesStart<'_>,
        empty: bool,
        decoder: quick_xml::encoding::Decoder,
    ) -> xlsx_jsx::Result<()> {
        let tag = event.local_name();
        match tag.as_ref() {
            b"dxfs" => self.in_dxfs = true,
            b"dxf" if self.in_dxfs => {
                if self.in_dxf {
                    return Err(Error::Reverse(
                        "styles.xml contains nested differential styles".into(),
                    ));
                }
                self.in_dxf = true;
                self.border_seen = false;
                self.border = None;
                if empty {
                    self.formats.push(None);
                    self.in_dxf = false;
                }
            }
            b"border" if self.in_dxf => {
                if self.border_seen {
                    return Err(Error::Reverse(
                        "styles.xml differential style contains multiple borders; keep one border per dxf"
                            .into(),
                    ));
                }
                self.border_seen = true;
                let mut builder = BorderBuilder::new();
                builder.parse_differential_border_attributes(event, decoder);
                if empty {
                    self.border = Some(Self::finish_border(builder)?);
                } else {
                    self.builder = Some(builder);
                }
            }
            b"top" | b"right" | b"bottom" | b"left" | b"diagonal" if self.in_dxf => {
                let Some(builder) = self.builder.as_mut() else {
                    return Err(Error::Reverse(
                        "styles.xml differential border edge is outside a border; repair styles.xml"
                            .into(),
                    ));
                };
                builder.add_edge(event, tag.as_ref(), decoder);
                if empty {
                    builder.current_edge = None;
                }
            }
            b"color" if self.in_dxf && self.builder.is_some() => {
                if let Some(builder) = self.builder.as_mut() {
                    builder.add_color(event, decoder);
                }
            }
            b"vertical" | b"horizontal" if self.in_dxf => {
                if !empty || event.attributes().next().is_some() {
                    return Err(Error::Reverse(
                        "styles.xml differential border vertical/horizontal edges must be empty; remove unsupported edge metadata"
                            .into(),
                    ));
                }
            }
            _ if self.builder.is_some() => {
                return Err(Error::Reverse(format!(
                    "styles.xml differential border contains unsupported `{}` metadata; use only four edge borders and RGB colors",
                    String::from_utf8_lossy(tag.as_ref())
                )));
            }
            _ => {}
        }
        Ok(())
    }

    fn end(&mut self, tag: &[u8]) -> xlsx_jsx::Result<()> {
        match tag {
            b"top" | b"right" | b"bottom" | b"left" | b"diagonal" => {
                if let Some(builder) = self.builder.as_mut() {
                    builder.current_edge = None;
                }
            }
            b"border" if self.in_dxf => {
                let builder = self.builder.take().ok_or_else(|| {
                    Error::Reverse(
                        "styles.xml differential border has no matching opening element".into(),
                    )
                })?;
                self.border = Some(Self::finish_border(builder)?);
            }
            b"dxf" if self.in_dxf => {
                if self.builder.is_some() {
                    return Err(Error::Reverse(
                        "styles.xml has an unfinished differential border; repair styles.xml"
                            .into(),
                    ));
                }
                self.formats.push(self.border.take());
                self.in_dxf = false;
            }
            b"dxfs" => self.in_dxfs = false,
            _ => {}
        }
        Ok(())
    }

    fn finish_border(builder: BorderBuilder) -> xlsx_jsx::Result<CellBorders> {
        builder.finish().ok_or_else(|| {
            Error::Reverse(
                "styles.xml differential border contains unsupported edge attributes or colors; preserve only OOXML border styles and opaque RGB colors"
                    .into(),
            )
        })
    }

    fn finish(self) -> xlsx_jsx::Result<Vec<Option<CellBorders>>> {
        if self.in_dxf || self.builder.is_some() {
            return Err(Error::Reverse(
                "styles.xml contains an unfinished differential border section".into(),
            ));
        }
        Ok(self.formats)
    }
}

fn parse_differential_border_formats(xml: &str) -> xlsx_jsx::Result<Vec<Option<CellBorders>>> {
    let mut parser = DifferentialBorderParser::default();
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) => parser.start(&event, false, reader.decoder())?,
            Ok(Event::Empty(event)) => parser.start(&event, true, reader.decoder())?,
            Ok(Event::End(event)) => parser.end(event.local_name().as_ref())?,
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid styles.xml differential border section: {error}"
                )));
            }
        }
    }
    parser.finish()
}

fn parse_ooxml_bool(value: &str) -> Option<bool> {
    match value {
        "1" | "true" => Some(true),
        "0" | "false" => Some(false),
        _ => None,
    }
}

fn parse_border_formats(xml: &str) -> xlsx_jsx::Result<Vec<Option<CellBorders>>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut borders = Vec::new();
    let mut current: Option<BorderBuilder> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) => parse_border_start_event(
                &event,
                false,
                reader.decoder(),
                &mut current,
                &mut borders,
            ),
            Ok(Event::Empty(event)) => {
                parse_border_start_event(
                    &event,
                    true,
                    reader.decoder(),
                    &mut current,
                    &mut borders,
                );
            }
            Ok(Event::End(event)) => match event.local_name().as_ref() {
                b"top" | b"right" | b"bottom" | b"left" | b"diagonal" => {
                    if let Some(border) = &mut current {
                        border.current_edge = None;
                    }
                }
                b"border" => {
                    if let Some(border) = current.take() {
                        borders.push(border.finish());
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid styles.xml: {error}"))),
        }
    }
    if current.is_some() {
        return Err(Error::Reverse(
            "styles.xml has an unfinished border definition".into(),
        ));
    }
    Ok(borders)
}

fn parse_border_start_event(
    event: &BytesStart<'_>,
    is_empty: bool,
    decoder: quick_xml::encoding::Decoder,
    current: &mut Option<BorderBuilder>,
    borders: &mut Vec<Option<CellBorders>>,
) {
    let tag = event.local_name();
    match tag.as_ref() {
        b"border" => {
            if current.is_some() {
                borders.push(None);
                *current = None;
            }
            let mut builder = BorderBuilder::new();
            builder.parse_border_attributes(event, decoder);
            if is_empty {
                borders.push(builder.finish());
            } else {
                *current = Some(builder);
            }
        }
        b"top" | b"right" | b"bottom" | b"left" | b"diagonal" => {
            if let Some(border) = current {
                border.add_edge(event, tag.as_ref(), decoder);
                if is_empty {
                    border.current_edge = None;
                }
            }
        }
        b"color" => {
            if let Some(border) = current {
                border.add_color(event, decoder);
            }
        }
        _ => {
            if let Some(border) = current {
                border.valid = false;
            }
        }
    }
}

enum BorderColor {
    Rgb(String),
    Automatic,
}

fn parse_border_color(
    event: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Option<BorderColor> {
    let mut color = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.ok()?;
        if color.is_some() {
            return None;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .ok()?;
        color = Some(match attribute.key.local_name().as_ref() {
            b"rgb"
                if value.len() == 8
                    && value
                        .get(..2)
                        .is_some_and(|alpha| alpha.eq_ignore_ascii_case("FF"))
                    && value
                        .get(2..)
                        .is_some_and(|rgb| rgb.bytes().all(|b| b.is_ascii_hexdigit())) =>
            {
                BorderColor::Rgb(format!("#{}", value.get(2..)?.to_ascii_uppercase()))
            }
            b"auto" if value == "1" => BorderColor::Automatic,
            b"indexed" if value == "64" => BorderColor::Automatic,
            _ => return None,
        });
    }
    color
}

#[derive(Default)]
struct FillBuilder {
    pattern: Option<String>,
    foreground: Option<String>,
    background: Option<String>,
    background_seen: bool,
    valid: bool,
}

impl FillBuilder {
    fn finish(self) -> FillFormat {
        if !self.valid {
            return FillFormat::Unsupported;
        }
        match self.pattern.as_deref() {
            Some("none") if self.foreground.is_none() && self.background.is_none() => {
                FillFormat::None
            }
            Some("solid") if self.background.is_none() => self
                .foreground
                .map_or(FillFormat::Unsupported, FillFormat::Solid),
            Some(pattern) => FillFormat::Pattern {
                pattern: pattern.into(),
                foreground: self.foreground,
                background: self.background,
            },
            None => FillFormat::Unsupported,
        }
    }
}

fn parse_fill_formats(xml: &str) -> xlsx_jsx::Result<Vec<FillFormat>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut fills = Vec::new();
    let mut current_fill: Option<FillBuilder> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if event.local_name().as_ref() == b"fill" => {
                if current_fill.is_some() {
                    return Err(Error::Reverse(
                        "styles.xml contains nested fill elements".into(),
                    ));
                }
                current_fill = Some(FillBuilder {
                    valid: true,
                    ..Default::default()
                });
            }
            Ok(Event::Empty(event)) if event.local_name().as_ref() == b"fill" => {
                fills.push(FillFormat::None);
            }
            Ok(Event::Start(event) | Event::Empty(event)) => {
                let tag = event.local_name();
                match tag.as_ref() {
                    b"patternFill" => {
                        if let Some(fill) = &mut current_fill {
                            if fill.pattern.is_some() {
                                fill.valid = false;
                            }
                            fill.pattern = parse_pattern_type(&event, reader.decoder());
                            fill.valid &= fill.pattern.is_some();
                        }
                    }
                    b"fgColor" => {
                        if let Some(fill) = &mut current_fill {
                            if fill.foreground.is_some() {
                                fill.valid = false;
                            }
                            match parse_fill_rgb(&event, reader.decoder()) {
                                Some(color) => fill.foreground = Some(color),
                                None => fill.valid = false,
                            }
                        }
                    }
                    b"bgColor" => {
                        if let Some(fill) = &mut current_fill {
                            if fill.background_seen {
                                fill.valid = false;
                            } else if is_default_fill_background(&event, reader.decoder()) {
                                // OOXML's default indexed background is equivalent to no explicit color.
                            } else if let Some(color) = parse_fill_rgb(&event, reader.decoder()) {
                                fill.background = Some(color);
                            } else {
                                fill.valid = false;
                            }
                            fill.background_seen = true;
                        }
                    }
                    _ => {
                        if let Some(fill) = &mut current_fill {
                            fill.valid = false;
                        }
                    }
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"fill" => {
                if let Some(fill) = current_fill.take() {
                    fills.push(fill.finish());
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(Error::Reverse(format!("invalid styles.xml: {error}"))),
        }
    }
    Ok(fills)
}

fn parse_pattern_type(
    event: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Option<String> {
    let mut pattern = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.ok()?;
        if attribute.key.local_name().as_ref() != b"patternType" || pattern.is_some() {
            return None;
        }
        pattern = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .ok()?
                .into_owned(),
        );
    }
    pattern.filter(|pattern| is_supported_fill_pattern(pattern))
}

fn parse_differential_pattern_type(
    event: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Option<String> {
    if event.attributes().next().is_none() {
        return Some("solid".into());
    }
    parse_pattern_type(event, decoder)
}

fn is_supported_fill_pattern(pattern: &str) -> bool {
    matches!(
        pattern,
        "none"
            | "solid"
            | "mediumGray"
            | "darkGray"
            | "lightGray"
            | "darkHorizontal"
            | "darkVertical"
            | "darkDown"
            | "darkUp"
            | "darkGrid"
            | "darkTrellis"
            | "lightHorizontal"
            | "lightVertical"
            | "lightDown"
            | "lightUp"
            | "lightGrid"
            | "lightTrellis"
            | "gray125"
            | "gray0625"
    )
}

fn parse_fill_rgb(event: &BytesStart<'_>, decoder: quick_xml::encoding::Decoder) -> Option<String> {
    let mut rgb = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.ok()?;
        if attribute.key.local_name().as_ref() != b"rgb" || rgb.is_some() {
            return None;
        }
        rgb = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .ok()?
                .into_owned(),
        );
    }
    let rgb = rgb?;
    if rgb.len() != 8
        || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit())
        || !rgb[..2].eq_ignore_ascii_case("FF")
    {
        return None;
    }
    Some(format!("#{}", &rgb[2..]).to_ascii_uppercase())
}

fn is_default_fill_background(
    event: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> bool {
    let mut indexed = None;
    for attribute in event.attributes().with_checks(false) {
        let Ok(attribute) = attribute else {
            return false;
        };
        if attribute.key.local_name().as_ref() != b"indexed" || indexed.is_some() {
            return false;
        }
        let Ok(value) = attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
        else {
            return false;
        };
        indexed = Some(value.into_owned());
    }
    indexed.as_deref() == Some("64")
}

fn parse_font_emphasis(
    event: &BytesStart<'_>,
    tag: &[u8],
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    let mut enabled = true;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" {
            style.support = StyleSupport::Unsupported;
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        enabled = !matches!(value.as_ref(), "0" | "false" | "off");
    }
    if tag == b"b" {
        style.bold = enabled;
    } else {
        style.italic = enabled;
    }
    Ok(())
}

fn parse_font_component(
    event: &BytesStart<'_>,
    tag: &[u8],
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    match tag {
        b"b" | b"i" => parse_font_emphasis(event, tag, style, decoder),
        b"strike" => parse_font_strikethrough(event, style, decoder),
        b"u" => parse_font_underline(event, style, decoder),
        b"vertAlign" => parse_font_script(event, style, decoder),
        b"color" => parse_font_color(event, style, decoder),
        _ => Ok(()),
    }
}

fn parse_font_strikethrough(
    event: &BytesStart<'_>,
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    let mut enabled = true;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" {
            style.support = StyleSupport::Unsupported;
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        enabled = !matches!(value.as_ref(), "0" | "false" | "off");
    }
    style.strikethrough = Some(enabled);
    Ok(())
}

fn parse_font_underline(
    event: &BytesStart<'_>,
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    let mut value = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" || value.is_some() {
            style.support = StyleSupport::Unsupported;
            continue;
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?
                .into_owned(),
        );
    }
    match value.as_deref().unwrap_or("single") {
        "none" => style.underline = None,
        "single" => style.underline = Some("single".to_owned()),
        "double" => style.underline = Some("double".to_owned()),
        "singleAccounting" => style.underline = Some("singleAccounting".to_owned()),
        "doubleAccounting" => style.underline = Some("doubleAccounting".to_owned()),
        _ => style.support = StyleSupport::Unsupported,
    }
    Ok(())
}

fn parse_font_script(
    event: &BytesStart<'_>,
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    let mut value = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" || value.is_some() {
            style.support = StyleSupport::Unsupported;
            continue;
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?
                .into_owned(),
        );
    }
    match value.as_deref() {
        Some("superscript") => style.script = Some("superscript".to_owned()),
        Some("subscript") => style.script = Some("subscript".to_owned()),
        _ => style.support = StyleSupport::Unsupported,
    }
    Ok(())
}

fn parse_font_color(
    event: &BytesStart<'_>,
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    let mut color = None;
    let mut default_theme = false;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match attribute.key.local_name().as_ref() {
            b"rgb"
                if value.len() == 8
                    && value
                        .get(..2)
                        .is_some_and(|alpha| alpha.eq_ignore_ascii_case("FF"))
                    && value
                        .get(2..)
                        .is_some_and(|rgb| rgb.bytes().all(|byte| byte.is_ascii_hexdigit())) =>
            {
                color = Some(format!("#{}", value[2..].to_ascii_uppercase()));
            }
            b"theme" if value == "1" => default_theme = true,
            _ => style.support = StyleSupport::Unsupported,
        }
    }
    if color.is_some() && default_theme {
        style.support = StyleSupport::Unsupported;
    }
    style.font_color = color;
    if style.font_color.is_none() && !default_theme {
        style.support = StyleSupport::Unsupported;
    }
    Ok(())
}

fn parse_font_name_or_size(
    event: &BytesStart<'_>,
    tag: &[u8],
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
    default_font: bool,
) -> xlsx_jsx::Result<()> {
    if default_font {
        return audit_default_font_component(event, tag, style, decoder);
    }
    let mut value = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" || value.is_some() {
            style.support = StyleSupport::Unsupported;
            continue;
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?
                .into_owned(),
        );
    }
    let Some(value) = value else {
        style.support = StyleSupport::Unsupported;
        return Ok(());
    };
    match tag {
        b"name" if !value.is_empty() && value.chars().count() <= 255 => {
            style.font_name = Some(value);
        }
        b"sz" => match value.parse::<f64>() {
            Ok(size) if (1.0..=409.0).contains(&size) => style.font_size = Some(size),
            _ => style.support = StyleSupport::Unsupported,
        },
        _ => style.support = StyleSupport::Unsupported,
    }
    Ok(())
}

fn parse_font_family_or_charset(
    event: &BytesStart<'_>,
    tag: &[u8],
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    let mut value = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" || value.is_some() {
            style.support = StyleSupport::Unsupported;
            continue;
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?
                .into_owned(),
        );
    }
    let Some(value) = value else {
        style.support = StyleSupport::Unsupported;
        return Ok(());
    };
    let Ok(value) = value.parse::<u8>() else {
        style.support = StyleSupport::Unsupported;
        return Ok(());
    };
    match tag {
        b"family" if value <= 5 => {
            if style.font_family.replace(value).is_some() {
                style.support = StyleSupport::Unsupported;
            }
        }
        b"charset" => {
            if style.font_charset.replace(value).is_some() {
                style.support = StyleSupport::Unsupported;
            }
        }
        _ => style.support = StyleSupport::Unsupported,
    }
    Ok(())
}

fn parse_font_scheme(
    event: &BytesStart<'_>,
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
    default_font: bool,
) -> xlsx_jsx::Result<()> {
    if default_font {
        return audit_default_font_component(event, b"scheme", style, decoder);
    }
    let mut scheme = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" || scheme.is_some() {
            style.support = StyleSupport::Unsupported;
            continue;
        }
        scheme = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Reverse(error.to_string()))?
                .into_owned(),
        );
    }
    match scheme.as_deref() {
        Some("none") => style.font_scheme = Some("none".to_owned()),
        Some("major") => style.font_scheme = Some("major".to_owned()),
        Some("minor") => style.font_scheme = Some("minor".to_owned()),
        _ => style.support = StyleSupport::Unsupported,
    }
    Ok(())
}

fn finalize_font_style(mut style: CellFontStyle) -> CellFontStyle {
    if style.font_scheme.as_deref() == Some("minor") {
        if style.font_name.as_deref() == Some("Calibri") {
            style.font_name = None;
        }
        // `Format::new()` already selects the minor/body theme font.
        style.font_scheme = None;
    }
    if matches!(style.font_family, Some(0 | 2)) {
        style.font_family = None;
    }
    if style.font_charset == Some(0) {
        style.font_charset = None;
    }
    style
}

fn audit_default_font_component(
    event: &BytesStart<'_>,
    tag: &[u8],
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    let expected = match tag {
        b"sz" => Some((b"val".as_slice(), "11")),
        b"color" => Some((b"theme".as_slice(), "1")),
        b"name" => Some((b"val".as_slice(), "Calibri")),
        b"family" => Some((b"val".as_slice(), "2")),
        b"scheme" => Some((b"val".as_slice(), "minor")),
        _ => None,
    };
    let mut matched = false;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        if expected.is_some_and(|(name, expected_value)| {
            attribute.key.local_name().as_ref() == name && value == expected_value
        }) {
            matched = true;
        } else {
            style.support = StyleSupport::Unsupported;
        }
    }
    if expected.is_some() && !matched {
        style.support = StyleSupport::Unsupported;
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "A cell XF is decoded as one coupled style record so unsupported combinations cannot be partially accepted."
)]
fn parse_cell_format(
    event: &BytesStart<'_>,
    fonts: &[CellFontStyle],
    number_formats: &HashMap<usize, String>,
    fill_formats: &[FillFormat],
    border_formats: &[Option<CellBorders>],
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<CellFontStyle> {
    let mut font_id = 0;
    let mut number_format_id = 0;
    let mut fill_id = 0;
    let mut border_id = 0;
    let mut supported = true;
    let mut apply_font = true;
    let mut apply_number_format = true;
    let mut apply_fill = true;
    let mut apply_border = true;
    let mut apply_alignment = false;
    let mut apply_protection = false;
    let mut quote_prefix = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let key = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match key.as_ref() {
            b"fontId" => {
                font_id = value
                    .parse::<usize>()
                    .map_err(|error| Error::Reverse(format!("invalid font index: {error}")))?;
            }
            b"numFmtId" => {
                number_format_id = value.parse::<usize>().map_err(|error| {
                    Error::Reverse(format!("invalid number format index: {error}"))
                })?;
            }
            b"fillId" => {
                fill_id = value
                    .parse::<usize>()
                    .map_err(|error| Error::Reverse(format!("invalid fill index: {error}")))?;
            }
            b"borderId" => {
                border_id = value
                    .parse::<usize>()
                    .map_err(|error| Error::Reverse(format!("invalid border index: {error}")))?;
            }
            b"xfId" => {
                supported &= value == "0";
            }
            b"applyFont" => apply_font = value != "0",
            b"applyNumberFormat" => apply_number_format = value != "0",
            b"applyFill" => apply_fill = value != "0",
            b"applyBorder" => apply_border = value != "0",
            b"applyAlignment" => apply_alignment = value != "0",
            b"applyProtection" => {
                apply_protection = parse_xml_bool(&value, "styles.xml", "xf.applyProtection")?;
            }
            b"quotePrefix" => quote_prefix = parse_quote_prefix(&value)?,
            _ => supported = false,
        }
    }
    let mut style = fonts.get(font_id).cloned().unwrap_or_default();
    style.quote_prefix = quote_prefix;
    style.apply_protection = apply_protection;
    if fill_id == 0 || !apply_fill {
        style.fill_color = None;
        style.fill_pattern = None;
        style.fill_foreground_color = None;
        style.fill_background_color = None;
    } else {
        match fill_formats.get(fill_id) {
            Some(FillFormat::Solid(color)) => style.fill_color = Some(color.clone()),
            Some(FillFormat::None) => style.fill_color = None,
            Some(FillFormat::Pattern {
                pattern,
                foreground,
                background,
            }) => {
                style.fill_pattern = Some(pattern.clone());
                style.fill_foreground_color.clone_from(foreground);
                style.fill_background_color.clone_from(background);
            }
            Some(FillFormat::Unsupported) | None => supported = false,
        }
    }
    if apply_border {
        match border_formats.get(border_id) {
            Some(Some(borders)) if border_id == 0 && !borders.has_any() => {}
            Some(Some(borders)) if border_id != 0 => style.borders = borders.clone(),
            None if border_id == 0 => {}
            _ => supported = false,
        }
    }
    if number_format_id == 0 {
        style.number_format = None;
    } else if !apply_number_format {
        supported = false;
    } else if let Some(code) = number_formats.get(&number_format_id) {
        style.number_format = Some(code.clone());
    } else if let Some(code) = builtin_number_format(number_format_id) {
        style.number_format = Some(code.to_owned());
    } else {
        supported = false;
    }
    if !supported {
        style.support = StyleSupport::Unsupported;
    }
    if !apply_font {
        style.bold = false;
        style.italic = false;
        style.font_color = None;
        style.font_name = None;
        style.font_size = None;
        style.font_family = None;
        style.font_charset = None;
    }
    style.apply_alignment = apply_alignment;
    Ok(style)
}

fn parse_quote_prefix(value: &str) -> xlsx_jsx::Result<Option<()>> {
    let enabled = parse_xml_bool(value, "styles.xml", "xf.quotePrefix").map_err(|_| {
        Error::Reverse(format!(
            "styles.xml cell format has invalid quotePrefix value `{value}`; this must be a boolean (0/1 or false/true), so repair it before reversing"
        ))
    })?;
    Ok(enabled.then_some(()))
}

fn parse_cell_protection(
    event: &BytesStart<'_>,
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    let mut locked = None;
    let mut hidden = false;
    let mut locked_seen = false;
    let mut hidden_seen = false;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        let enabled = parse_xml_bool(&value, "styles.xml", "protection flag")?;
        match attribute.key.local_name().as_ref() {
            b"locked" if !locked_seen => {
                locked_seen = true;
                if !enabled {
                    locked = Some(false);
                }
            }
            b"hidden" if !hidden_seen => {
                hidden_seen = true;
                hidden = enabled;
            }
            _ => {
                return Err(Error::Reverse(format!(
                    "styles.xml cell protection has duplicate or unsupported attribute `{}`; only locked and hidden are supported",
                    String::from_utf8_lossy(attribute.key.local_name().as_ref())
                )));
            }
        }
    }
    if !style.apply_protection && (locked.is_some() || hidden) {
        style.support = StyleSupport::Unsupported;
    } else {
        style.locked = locked;
        style.formula_hidden = hidden;
    }
    Ok(())
}

fn parse_cell_alignment(
    event: &BytesStart<'_>,
    style: &mut CellFontStyle,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<()> {
    if !style.apply_alignment {
        return Ok(());
    }
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?;
        match attribute.key.local_name().as_ref() {
            b"horizontal" => {
                style.horizontal_align = match value.as_ref() {
                    "left" => Some(HorizontalAlignment::Left),
                    "center" => Some(HorizontalAlignment::Center),
                    "right" => Some(HorizontalAlignment::Right),
                    "distributed" => Some(HorizontalAlignment::Distributed),
                    "general" => None,
                    _ => {
                        style.support = StyleSupport::Unsupported;
                        None
                    }
                };
            }
            b"vertical" => {
                style.vertical_align = match value.as_ref() {
                    "top" => Some(VerticalAlignment::Top),
                    "center" => Some(VerticalAlignment::Center),
                    "bottom" => None,
                    _ => {
                        style.support = StyleSupport::Unsupported;
                        None
                    }
                };
            }
            b"wrapText" => match value.as_ref() {
                "1" | "true" => style.wrap_text = Some(true),
                "0" | "false" => style.wrap_text = Some(false),
                _ => style.support = StyleSupport::Unsupported,
            },
            b"shrinkToFit" => match value.as_ref() {
                "1" | "true" => style.shrink_to_fit = Some(true),
                "0" | "false" => style.shrink_to_fit = Some(false),
                _ => style.support = StyleSupport::Unsupported,
            },
            b"indent" => match value.parse::<u8>() {
                Ok(0) => style.indent = None,
                Ok(indent) => style.indent = Some(indent),
                _ => {
                    return Err(Error::Reverse(
                        "cell alignment indent must be an integer from 0 through 255; repair the style metadata before reversing".into(),
                    ));
                }
            },
            b"textRotation" => style.text_rotation = parse_text_rotation(&value)?,
            b"readingOrder" => {
                style.reading_direction = match value.parse::<u8>() {
                    Ok(0) => None,
                    Ok(1) => Some(CellReadingDirection::LeftToRight),
                    Ok(2) => Some(CellReadingDirection::RightToLeft),
                    _ => {
                        return Err(Error::Reverse(format!(
                            "cell alignment readingOrder `{value}` is unsupported; use OOXML value 1 (left-to-right), 2 (right-to-left), or the contextual default 0"
                        )));
                    }
                };
            }
            b"relativeIndent" | b"justifyLastLine" => {
                style.support = StyleSupport::Unsupported;
            }
            _ => style.support = StyleSupport::Unsupported,
        }
    }
    if style.wrap_text == Some(true) && style.shrink_to_fit == Some(true) {
        return Err(Error::Reverse(
            "cell alignment in styles.xml combines wrapText and shrinkToFit, but rust_xlsxwriter suppresses shrinkToFit when wrapText is enabled; remove one flag to allow lossless JSX output".into(),
        ));
    }
    if style.indent.is_some()
        && !matches!(
            style.horizontal_align,
            Some(
                HorizontalAlignment::Left
                    | HorizontalAlignment::Right
                    | HorizontalAlignment::Distributed
            )
        )
    {
        return Err(Error::Reverse(
            "cell alignment indent requires explicit left, right, or distributed horizontal alignment; add a supported horizontal alignment or remove indent to avoid changing layout".into(),
        ));
    }
    if style.horizontal_align == Some(HorizontalAlignment::Distributed)
        && style.shrink_to_fit == Some(true)
    {
        return Err(Error::Reverse(
            "distributed horizontal alignment cannot be combined with shrinkToFit because rust_xlsxwriter suppresses shrink-to-fit; remove one setting to preserve the source behavior".into(),
        ));
    }
    Ok(())
}

fn parse_text_rotation(value: &str) -> xlsx_jsx::Result<Option<i16>> {
    let angle = value.parse::<u16>().map_err(|_| {
        Error::Reverse(
            "cell alignment textRotation must be an OOXML angle from 0 through 180 or 255 for stacked text; repair the alignment before reversing".into(),
        )
    })?;
    match angle {
        0 => Ok(None),
        1..=90 => Ok(Some(i16::try_from(angle).map_err(|error| {
            Error::Reverse(format!("invalid textRotation in styles.xml: {error}"))
        })?)),
        91..=180 => {
            let angle = i16::try_from(angle).map_err(|error| {
                Error::Reverse(format!("invalid textRotation in styles.xml: {error}"))
            })?;
            Ok(Some(90 - angle))
        }
        255 => Ok(Some(270)),
        _ => Err(Error::Reverse(
            "cell alignment textRotation must be an OOXML angle from 0 through 180 or 255 for stacked text; repair the alignment before reversing".into(),
        )),
    }
}

fn parse_custom_number_format(
    event: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<Option<(usize, String)>> {
    let mut id = None;
    let mut code = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| Error::Reverse(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?
            .into_owned();
        match attribute.key.local_name().as_ref() {
            b"numFmtId" => id = value.parse::<usize>().ok(),
            b"formatCode" => code = Some(value),
            _ => {}
        }
    }
    Ok(id.zip(code))
}

fn parse_differential_number_format(
    event: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> xlsx_jsx::Result<Option<String>> {
    let mut id = None;
    let mut code = None;
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| {
            Error::Reverse(format!("invalid differential number format: {error}"))
        })?;
        let key = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Reverse(error.to_string()))?
            .into_owned();
        match key.as_ref() {
            b"numFmtId" if id.is_none() => {
                id = Some(value.parse::<usize>().map_err(|error| {
                    Error::Reverse(format!("invalid differential number format ID: {error}"))
                })?);
            }
            b"formatCode" if code.is_none() => code = Some(value),
            _ => {
                return Err(Error::Reverse(format!(
                    "styles.xml differential numFmt contains duplicate or unsupported attribute `{}`; preserve only numFmtId and formatCode",
                    String::from_utf8_lossy(key.as_ref())
                )));
            }
        }
    }
    let id = id.ok_or_else(|| {
        Error::Reverse("styles.xml differential numFmt is missing numFmtId".into())
    })?;
    let code = code.unwrap_or_default();
    let format = if code.is_empty() {
        builtin_number_format(id).ok_or_else(|| {
            Error::Reverse(format!(
                "styles.xml differential numFmt ID {id} has no formatCode and is not a supported built-in; use a standard built-in number format or an explicit format code"
            ))
        })?
        .to_owned()
    } else {
        code
    };
    if format.chars().count() > 255 || !format.chars().all(is_valid_xml_character) {
        return Err(Error::Reverse(
            "styles.xml differential number format is longer than 255 characters or contains an XML-invalid character; repair formatCode".into(),
        ));
    }
    Ok(Some(format))
}

fn classify_date_format(code: &str) -> DateNumberFormat {
    if code.eq_ignore_ascii_case("yyyy-mm-dd") {
        DateNumberFormat::Date
    } else if code.eq_ignore_ascii_case("yyyy-mm-dd hh:mm:ss") {
        DateNumberFormat::DateTime
    } else if code.eq_ignore_ascii_case("yyyy-mm-dd hh:mm:ss.000") {
        DateNumberFormat::DateTimeMillis
    } else {
        DateNumberFormat::Unsupported
    }
}

fn builtin_number_format(id: usize) -> Option<&'static str> {
    match id {
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        5 => Some("($#,##0_);($#,##0)"),
        6 => Some("($#,##0_);[Red]($#,##0)"),
        7 => Some("($#,##0.00_);($#,##0.00)"),
        8 => Some("($#,##0.00_);[Red]($#,##0.00)"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("m/d/yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("(#,##0_);(#,##0)"),
        38 => Some("(#,##0_);[Red](#,##0)"),
        39 => Some("(#,##0.00_);(#,##0.00)"),
        40 => Some("(#,##0.00_);[Red](#,##0.00)"),
        41 => Some("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(_)"),
        42 => Some("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(_)"),
        43 => Some("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(_)"),
        44 => Some("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(_)"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mm:ss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dynamic_array_metadata_xml() -> String {
        r#"<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>"#.to_owned()
    }

    #[test]
    fn dynamic_array_metadata_accepts_writer_structure_and_rejects_unknown_semantics() {
        let xml = dynamic_array_metadata_xml();
        assert!(audit_dynamic_array_metadata(&xml).is_ok());

        let invalid = xml.replace("fCollapsed=\"0\"", "fCollapsed=\"1\"");
        let error =
            audit_dynamic_array_metadata(&invalid).expect_err("collapsed metadata is unsupported");
        assert!(error.to_string().contains("dynamic-array metadata"));
    }

    #[test]
    fn dynamic_array_metadata_relationship_must_match_part_presence() {
        let relationships = r#"<Relationships><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>"#;
        assert!(validate_dynamic_metadata_relationship(relationships, true).is_ok());
        let error = validate_dynamic_metadata_relationship(relationships, false)
            .expect_err("orphan sheetMetadata relationship is invalid");
        assert!(
            error
                .to_string()
                .contains("relationship and xl/metadata.xml part")
        );
    }

    #[test]
    fn core_properties_preserve_entities_and_significant_whitespace() {
        let xml = r#"<cp:coreProperties xmlns:cp="urn:cp" xmlns:dc="urn:dc"><dc:title> A &amp; B </dc:title><dc:creator>Analyst</dc:creator><cp:lastModifiedBy>Analyst</cp:lastModifiedBy></cp:coreProperties>"#;

        let properties = audit_core_properties(xml).expect("supported properties parse");

        assert_eq!(properties.title.as_deref(), Some(" A & B "));
        assert_eq!(properties.author.as_deref(), Some("Analyst"));
    }

    #[test]
    fn core_properties_reject_unmodeled_authored_fields() {
        let xml = r#"<cp:coreProperties xmlns:cp="urn:cp"><cp:category>Finance</cp:category></cp:coreProperties>"#;

        let error = audit_core_properties(xml).expect_err("category is not modeled");

        assert!(error.to_string().contains("category"));
        assert!(
            error
                .to_string()
                .contains("matching Workbook metadata support")
        );
    }

    #[test]
    fn core_properties_reject_different_creator_and_last_modifier() {
        let xml = r#"<cp:coreProperties xmlns:cp="urn:cp" xmlns:dc="urn:dc"><dc:creator>A</dc:creator><cp:lastModifiedBy>B</cp:lastModifiedBy></cp:coreProperties>"#;

        let error = audit_core_properties(xml).expect_err("author fields differ");

        assert!(
            error
                .to_string()
                .contains("creator and lastModifiedBy differ")
        );
    }
}
