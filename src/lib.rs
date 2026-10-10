use base64::{Engine, engine::general_purpose::STANDARD as BASE64};
use rust_xlsxwriter::Image as XlsxImage;
use serde::{Deserialize, Serialize};
use serde_json::Value;
pub mod error;
pub mod runtime;
pub use error::Error;
pub type Result<T> = std::result::Result<T, Error>;

pub mod ir {
    pub use crate::{IrEnvelope, Node};
}

/// Parses an Excel A1 cell range into zero-based `(first_row, first_column,
/// last_row, last_column)` coordinates.
///
/// Absolute-reference markers and lowercase column letters are accepted.
///
/// # Examples
///
/// ```
/// assert_eq!(xlsx_jsx::parse_a1_range("$B$2:D10"), Some((1, 1, 9, 3)));
/// ```
#[must_use]
pub fn parse_a1_range(reference: &str) -> Option<(u32, u16, u32, u16)> {
    let (first, last) = reference.split_once(':')?;
    if last.contains(':') {
        return None;
    }
    let (first_row, first_column) = parse_a1_cell_reference(first)?;
    let (last_row, last_column) = parse_a1_cell_reference(last)?;
    if first_row > last_row || first_column > last_column {
        return None;
    }
    Some((first_row, first_column, last_row, last_column))
}

/// Formats zero-based worksheet coordinates as a canonical uppercase A1 range.
#[must_use]
pub fn format_a1_range(range: (u32, u16, u32, u16)) -> String {
    format!(
        "{}:{}",
        format_a1_cell(range.0, range.1),
        format_a1_cell(range.2, range.3)
    )
}

/// Inclusive range used by worksheet print areas and repeating print titles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PrintRange {
    Cells {
        first_row: u32,
        first_column: u16,
        last_row: u32,
        last_column: u16,
    },
    Rows {
        first: u32,
        last: u32,
    },
    Columns {
        first: u16,
        last: u16,
    },
}

/// Parses a cell, whole-row, or whole-column print range using Excel A1 notation.
#[must_use]
pub fn parse_print_range(reference: &str) -> Option<PrintRange> {
    let (first, last) = reference.split_once(':')?;
    if last.contains(':') {
        return None;
    }
    if let (Some((first_row, first_column)), Some((last_row, last_column))) = (
        parse_a1_cell_reference(first),
        parse_a1_cell_reference(last),
    ) {
        if first_row <= last_row && first_column <= last_column {
            return Some(PrintRange::Cells {
                first_row,
                first_column,
                last_row,
                last_column,
            });
        }
        return None;
    }
    if let (Some(first), Some(last)) = (parse_print_row(first), parse_print_row(last)) {
        return (first <= last).then_some(PrintRange::Rows { first, last });
    }
    if let (Some(first), Some(last)) = (parse_print_column(first), parse_print_column(last)) {
        return (first <= last).then_some(PrintRange::Columns { first, last });
    }
    None
}

/// Formats a print range as canonical uppercase A1 notation.
#[must_use]
pub fn format_print_range(range: PrintRange) -> String {
    match range {
        PrintRange::Cells {
            first_row,
            first_column,
            last_row,
            last_column,
        } => format_a1_range((first_row, first_column, last_row, last_column)),
        PrintRange::Rows { first, last } => format!("{}:{}", first + 1, last + 1),
        PrintRange::Columns { first, last } => {
            format!("{}:{}", print_column_name(first), print_column_name(last))
        }
    }
}

fn parse_print_row(reference: &str) -> Option<u32> {
    let digits = reference.strip_prefix('$').unwrap_or(reference);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    let row = digits.parse::<u32>().ok()?;
    (1..=1_048_576).contains(&row).then_some(row - 1)
}

fn parse_print_column(reference: &str) -> Option<u16> {
    let letters = reference.strip_prefix('$').unwrap_or(reference);
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return None;
    }
    let mut column = 0_u32;
    for byte in letters.bytes() {
        column = column
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    if !(1..=16_384).contains(&column) {
        return None;
    }
    u16::try_from(column - 1).ok()
}

fn print_column_name(column: u16) -> String {
    format_a1_cell(0, column)
        .chars()
        .take_while(char::is_ascii_alphabetic)
        .collect()
}

/// Parses an Excel A1 cell reference into zero-based `(row, column)` coordinates.
#[must_use]
pub fn parse_a1_cell_reference(reference: &str) -> Option<(u32, u16)> {
    parse_a1_cell(reference)
}

/// Formats zero-based worksheet coordinates as a canonical uppercase A1 cell reference.
#[must_use]
pub fn format_a1_cell_reference(row: u32, column: u16) -> String {
    format_a1_cell(row, column)
}

fn parse_a1_cell(reference: &str) -> Option<(u32, u16)> {
    let bytes = reference.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'$'));
    let column_start = index;
    let mut column = 0_u32;
    while let Some(byte) = bytes.get(index).copied() {
        if !byte.is_ascii_alphabetic() {
            break;
        }
        column = column
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))?;
        index += 1;
    }
    if index == column_start || column > 16_384 {
        return None;
    }
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let row_start = index;
    if bytes.get(index) == Some(&b'0') {
        return None;
    }
    let mut row = 0_u32;
    while let Some(byte) = bytes.get(index).copied() {
        if !byte.is_ascii_digit() {
            return None;
        }
        row = row.checked_mul(10)?.checked_add(u32::from(byte - b'0'))?;
        index += 1;
    }
    if row_start == index || row == 0 || row > 1_048_576 {
        return None;
    }
    Some((row - 1, u16::try_from(column - 1).ok()?))
}

fn format_a1_cell(row: u32, column: u16) -> String {
    let mut column_number = u32::from(column) + 1;
    let mut letters = Vec::new();
    while column_number > 0 {
        let remainder = (column_number - 1) % 26;
        letters.push(char::from_u32(u32::from(b'A') + remainder).unwrap_or('A'));
        column_number = (column_number - 1) / 26;
    }
    let column_name: String = letters.into_iter().rev().collect();
    format!("{column_name}{}", row + 1)
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Node {
    pub r#type: String,
    #[serde(default)]
    pub props: serde_json::Map<String, Value>,
    #[serde(default)]
    pub children: Vec<Value>,
}

const FORMAT_PROPERTIES: &[&str] = &[
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
    "fontScheme",
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

const FILL_PATTERN_NAMES: &[&str] = &[
    "none",
    "solid",
    "mediumGray",
    "darkGray",
    "lightGray",
    "darkHorizontal",
    "darkVertical",
    "darkDown",
    "darkUp",
    "darkGrid",
    "darkTrellis",
    "lightHorizontal",
    "lightVertical",
    "lightDown",
    "lightUp",
    "lightGrid",
    "lightTrellis",
    "gray125",
    "gray0625",
];

/// Outline state for one worksheet row or column coordinate.
#[derive(Debug, Clone, Copy, Default)]
pub struct OutlineProperties {
    /// Whether the row or column is hidden.
    pub hidden: bool,
    /// Its full outline nesting level.
    pub level: u8,
    /// Whether it is the summary coordinate after a collapsed group.
    pub collapsed: bool,
}

/// Derives collapsed outline member ranges from summary markers and hidden coordinates.
///
/// # Errors
///
/// Returns a diagnostic when a summary has no preceding contiguous hidden,
/// outlined members, or when nested collapsed ranges exceed a coordinate's
/// declared outline level.
pub fn collapsed_outline_ranges(
    entries: &std::collections::BTreeMap<u32, OutlineProperties>,
    axis: &str,
) -> std::result::Result<Vec<(u32, u32)>, String> {
    let mut ranges = Vec::new();
    for (&summary, properties) in entries {
        if !properties.collapsed {
            continue;
        }
        let Some(mut member) = summary.checked_sub(1) else {
            return Err(format!(
                "collapsed {axis} summary at index 0 must follow hidden outline {axis}s"
            ));
        };
        let mut first_member = None;
        while let Some(properties) = entries.get(&member) {
            if !properties.hidden {
                break;
            }
            if properties.level == 0 {
                return Err(format!(
                    "collapsed {axis} summary at index {summary} follows a hidden {axis} without an outline level; add `outlineLevel` to the hidden members"
                ));
            }
            first_member = Some(member);
            let Some(previous) = member.checked_sub(1) else {
                break;
            };
            member = previous;
        }
        let first_member = first_member.ok_or_else(|| {
            format!(
                "collapsed {axis} summary at index {summary} must follow at least one contiguous hidden {axis} with a positive outline level"
            )
        })?;
        ranges.push((first_member, summary - 1));
    }
    for (&coordinate, properties) in entries {
        let collapsed_depth = ranges
            .iter()
            .filter(|(first, last)| *first <= coordinate && coordinate <= *last)
            .count();
        if collapsed_depth > usize::from(properties.level) {
            return Err(format!(
                "{axis} at index {coordinate} has outline level {} but belongs to {collapsed_depth} nested collapsed groups; raise `outlineLevel` or simplify the group nesting",
                properties.level
            ));
        }
    }
    Ok(ranges)
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IrEnvelope {
    pub version: u8,
    pub workbook: Node,
}

impl IrEnvelope {
    /// Validates component nesting, property names, worksheet names, and cell values.
    ///
    /// # Errors
    ///
    /// Returns a diagnostic with the offending component and a repair suggestion
    /// when the workbook violates the v1 DSL contract.
    pub fn validate(&self) -> Result<()> {
        if self.version != 1 || self.workbook.r#type != "Workbook" {
            return Err(Error::Compile("expected IR v1 Workbook root".into()));
        }
        if self.workbook.children.is_empty() {
            return Err(Error::Compile(
                "Workbook must contain at least one Worksheet".into(),
            ));
        }
        reject_unknown_props(
            &self.workbook,
            &["title", "subject", "author", "keywords", "description"],
            "Workbook",
        )?;
        validate_workbook_properties(&self.workbook)?;
        let mut names = std::collections::HashSet::new();
        let mut table_names = std::collections::HashSet::new();
        let mut global_defined_names = std::collections::HashSet::new();
        let mut worksheet_index = 0;
        validate_active_worksheet_uniqueness(&self.workbook.children)?;
        validate_default_active_selection(&self.workbook.children)?;
        validate_first_tab_uniqueness(&self.workbook.children)?;
        for (child_index, sheet_value) in self.workbook.children.iter().enumerate() {
            let sheet: Node = serde_json::from_value(sheet_value.clone()).map_err(|error| {
                Error::Compile(format!(
                    "Workbook child {child_index} must be Worksheet or DefinedName: {error}"
                ))
            })?;
            if sheet.r#type == "DefinedName" {
                validate_defined_name(&sheet, "Workbook", &mut global_defined_names)?;
                continue;
            }
            if sheet.r#type != "Worksheet" {
                return Err(Error::Compile(format!(
                    "Workbook child {child_index} is `{}`; expected Worksheet or DefinedName",
                    sheet.r#type
                )));
            }
            let sheet_index = worksheet_index;
            worksheet_index += 1;
            validate_worksheet_properties(&sheet)?;
            let name = sheet
                .props
                .get("name")
                .and_then(Value::as_str)
                .ok_or_else(|| {
                    Error::Compile(format!("Worksheet {sheet_index} requires name: string"))
                })?;
            validate_worksheet_visibility_and_active(&sheet, sheet_index, name)?;
            validate_first_tab(&sheet, name)?;
            validate_worksheet_tab_color(&sheet, name)?;
            validate_default_row_height(&sheet, name)?;
            validate_hide_unused_rows(&sheet, name)?;
            validate_show_gridlines(&sheet, name)?;
            validate_worksheet_zoom(&sheet, name)?;
            validate_worksheet_view_and_page_breaks(&sheet, name)?;
            validate_worksheet_reading_and_outline(&sheet, name)?;
            let units = name.encode_utf16().count();
            if !(1..=31).contains(&units)
                || name.chars().any(|ch| "[]:*?/\\".contains(ch))
                || name.starts_with('\'')
                || name.ends_with('\'')
            {
                return Err(Error::Compile(format!(
                    "Worksheet name `{name}` is invalid; use 1–31 characters, avoid []:*?/\\, and do not start/end with an apostrophe"
                )));
            }
            if let Some(auto_filter) = sheet.props.get("autoFilter") {
                let Some(auto_filter) = auto_filter.as_str() else {
                    return Err(Error::Compile(format!(
                        "Worksheet.autoFilter on `{name}` must be an A1 range string such as `A1:D20`; use a quoted range"
                    )));
                };
                if parse_a1_range(auto_filter).is_none() {
                    return Err(Error::Compile(format!(
                        "Worksheet.autoFilter `{auto_filter}` on `{name}` is not a valid ordered A1 range within Excel limits; use a range such as `A1:D20`"
                    )));
                }
            }
            validate_auto_filter_columns(&sheet, name)?;
            validate_freeze_panes(&sheet, name)?;
            if !names.insert(name.to_lowercase()) {
                return Err(Error::Compile(format!(
                    "Worksheet name `{name}` duplicates another sheet ignoring case; choose a unique name"
                )));
            }
            validate_worksheet_children(&sheet, name, &mut table_names)?;
        }
        if names.is_empty() {
            return Err(Error::Compile(
                "Workbook must contain at least one Worksheet; add a worksheet alongside any DefinedName declarations".into(),
            ));
        }
        validate_chart_sheet_references(&self.workbook.children, &names)?;
        Ok(())
    }
}

fn validate_workbook_properties(workbook: &Node) -> Result<()> {
    for property in ["title", "subject", "author", "keywords", "description"] {
        if let Some(value) = workbook.props.get(property)
            && !value.is_string()
        {
            return Err(Error::Compile(format!(
                "Workbook.{property} must be a string; use a quoted string or remove the property"
            )));
        }
    }
    Ok(())
}

fn validate_defined_name(
    component: &Node,
    scope: &str,
    names: &mut std::collections::HashSet<String>,
) -> Result<()> {
    reject_unknown_props(component, &["name", "formula"], "DefinedName")?;
    if !component.children.is_empty() {
        return Err(Error::Compile(format!(
            "DefinedName in {scope} cannot have children; remove them"
        )));
    }
    let name = component
        .props
        .get("name")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(format!(
                "DefinedName in {scope} requires name: string; add a valid Excel name"
            ))
        })?;
    let formula = component.props.get("formula").and_then(Value::as_str).ok_or_else(|| {
        Error::Compile(format!("DefinedName `{name}` in {scope} requires formula: string; provide an Excel formula or reference"))
    })?;
    if name.is_empty() || name.starts_with("_xlnm.") {
        return Err(Error::Compile(format!(
            "DefinedName `{name}` in {scope} is empty or reserved; choose a non-reserved Excel name"
        )));
    }
    if let Err(error) = rust_xlsxwriter::utility::check_name(name) {
        return Err(Error::Compile(format!(
            "DefinedName `{name}` in {scope} is not a valid Excel name: {error}; use 1–255 characters without spaces or Excel-reserved symbols"
        )));
    }
    if formula.trim().is_empty() || formula.chars().any(char::is_control) {
        return Err(Error::Compile(format!(
            "DefinedName `{name}` formula must be non-empty text without control characters; provide an Excel formula or reference"
        )));
    }
    if !names.insert(name.to_lowercase()) {
        return Err(Error::Compile(format!(
            "DefinedName `{name}` duplicates another name in {scope} ignoring case; choose a unique name"
        )));
    }
    Ok(())
}

fn validate_active_worksheet_uniqueness(sheets: &[Value]) -> Result<()> {
    let mut active_name = None;
    for sheet in sheets {
        if sheet
            .get("props")
            .and_then(|props| props.get("active"))
            .and_then(Value::as_bool)
            == Some(true)
        {
            let name = sheet
                .get("props")
                .and_then(|props| props.get("name"))
                .and_then(Value::as_str)
                .unwrap_or("<unnamed>");
            if let Some(previous) = active_name {
                return Err(Error::Compile(format!(
                    "worksheets `{previous}` and `{name}` both have active=true; keep it on exactly one visible worksheet"
                )));
            }
            active_name = Some(name);
        }
    }
    Ok(())
}

fn validate_default_active_selection(sheets: &[Value]) -> Result<()> {
    let active = sheets
        .iter()
        .find(|sheet| {
            sheet.get("type").and_then(Value::as_str) == Some("Worksheet")
                && sheet
                    .get("props")
                    .and_then(|props| props.get("active"))
                    .and_then(Value::as_bool)
                    == Some(true)
        })
        .or_else(|| {
            sheets
                .iter()
                .find(|sheet| sheet.get("type").and_then(Value::as_str) == Some("Worksheet"))
        });
    let Some(active) = active else {
        return Ok(());
    };
    if active.get("props").and_then(|props| props.get("selected")) == Some(&Value::Bool(false)) {
        let name = active
            .get("props")
            .and_then(|props| props.get("name"))
            .and_then(Value::as_str)
            .unwrap_or("<unnamed>");
        return Err(Error::Compile(format!(
            "Worksheet `{name}` is the active worksheet but selected=false; rust_xlsxwriter always selects the active worksheet, so remove selected={{false}} or make another worksheet active"
        )));
    }
    Ok(())
}

fn validate_first_tab_uniqueness(sheets: &[Value]) -> Result<()> {
    let mut first_tab_name = None;
    for sheet in sheets {
        if sheet
            .get("props")
            .and_then(|props| props.get("firstTab"))
            .and_then(Value::as_bool)
            == Some(true)
        {
            let name = sheet
                .get("props")
                .and_then(|props| props.get("name"))
                .and_then(Value::as_str)
                .unwrap_or("<unnamed>");
            if let Some(previous) = first_tab_name {
                return Err(Error::Compile(format!(
                    "worksheets `{previous}` and `{name}` both have firstTab=true; select exactly one visible worksheet"
                )));
            }
            first_tab_name = Some(name);
        }
    }
    Ok(())
}

fn validate_first_tab(sheet: &Node, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("firstTab") else {
        return Ok(());
    };
    let Some(first_tab) = value.as_bool() else {
        return Err(Error::Compile(format!(
            "Worksheet.firstTab on `{name}` must be a boolean; use firstTab={{true}} on one visible worksheet"
        )));
    };
    if first_tab
        && sheet
            .props
            .get("visibility")
            .and_then(Value::as_str)
            .is_some_and(|visibility| visibility != "visible")
    {
        return Err(Error::Compile(format!(
            "firstTab Worksheet `{name}` must be visible; set visibility=\"visible\" or remove firstTab={{true}}"
        )));
    }
    Ok(())
}

fn validate_worksheet_visibility(sheet: &Node, sheet_index: usize, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("visibility") else {
        return Ok(());
    };
    let Some(visibility) = value.as_str() else {
        return Err(Error::Compile(format!(
            "Worksheet.visibility on {name} must be visible, hidden, or veryHidden; use a supported visibility string"
        )));
    };
    if !["visible", "hidden", "veryHidden"].contains(&visibility) {
        return Err(Error::Compile(format!(
            "Worksheet.visibility {visibility} on {name} is unsupported; use visible, hidden, or veryHidden"
        )));
    }
    if sheet_index == 0 && visibility != "visible" {
        return Err(Error::Compile(
            "the first Worksheet must have visibility visible so the workbook has a deterministic active sheet; move the visible sheet first"
                .into(),
        ));
    }
    Ok(())
}

fn validate_active_worksheet(sheet: &Node, name: &str) -> Result<bool> {
    let Some(value) = sheet.props.get("active") else {
        return Ok(false);
    };
    let Some(active) = value.as_bool() else {
        return Err(Error::Compile(format!(
            "Worksheet.active on `{name}` must be a boolean; use active={{true}} on one visible worksheet"
        )));
    };
    if active
        && sheet
            .props
            .get("visibility")
            .and_then(Value::as_str)
            .is_some_and(|visibility| visibility != "visible")
    {
        return Err(Error::Compile(format!(
            "active Worksheet `{name}` must be visible; set visibility=\"visible\" or remove active={{true}}"
        )));
    }
    Ok(active)
}

fn validate_worksheet_tab_color(sheet: &Node, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("tabColor") else {
        return Ok(());
    };
    let Some(color) = value.as_str() else {
        return Err(Error::Compile(format!(
            "Worksheet.tabColor on `{name}` must be an opaque RGB string such as `#4472C4`; use exactly six hexadecimal digits"
        )));
    };
    if color.len() != 7
        || !color.starts_with('#')
        || !color[1..].bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        return Err(Error::Compile(format!(
            "Worksheet.tabColor `{color}` on `{name}` must match `#RRGGBB`; use exactly six hexadecimal digits after `#`"
        )));
    }
    Ok(())
}

fn validate_default_row_height(sheet: &Node, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("defaultRowHeightPixels") else {
        return Ok(());
    };
    let Some(pixels) = value.as_u64() else {
        return Err(Error::Compile(format!(
            "Worksheet.defaultRowHeightPixels on `{name}` must be a positive integer; use a whole pixel value such as 24"
        )));
    };
    if !(1..=545).contains(&pixels) {
        return Err(Error::Compile(format!(
            "Worksheet.defaultRowHeightPixels on `{name}` must be between 1 and 545; choose a positive pixel height no greater than Excel's 409-point limit"
        )));
    }
    Ok(())
}

fn validate_hide_unused_rows(sheet: &Node, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("hideUnusedRows") else {
        return Ok(());
    };
    if !value.is_boolean() {
        return Err(Error::Compile(format!(
            "Worksheet.hideUnusedRows on `{name}` must be a boolean; use true to hide unused rows or false to keep them visible"
        )));
    }
    Ok(())
}

fn validate_show_gridlines(sheet: &Node, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("showGridlines") else {
        return Ok(());
    };
    if !value.is_boolean() {
        return Err(Error::Compile(format!(
            "Worksheet.showGridlines on `{name}` must be a boolean; use false to hide screen gridlines or true to show them"
        )));
    }
    Ok(())
}

fn validate_print_gridlines(sheet: &Node, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("printGridlines") else {
        return Ok(());
    };
    if !value.is_boolean() {
        return Err(Error::Compile(format!(
            "Worksheet.printGridlines on `{name}` must be a boolean; use true to print gridlines or false to omit them"
        )));
    }
    Ok(())
}

fn validate_worksheet_zoom(sheet: &Node, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("zoom") else {
        return Ok(());
    };
    let Some(zoom) = value.as_u64() else {
        return Err(Error::Compile(format!(
            "Worksheet.zoom on `{name}` must be an integer from 10 through 400; use a value such as 100 or 125"
        )));
    };
    if !(10..=400).contains(&zoom) {
        return Err(Error::Compile(format!(
            "Worksheet.zoom on `{name}` must be from 10 through 400; choose a valid screen zoom percentage"
        )));
    }
    Ok(())
}

fn validate_worksheet_view(sheet: &Node, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("view") else {
        return Ok(());
    };
    if !matches!(
        value.as_str(),
        Some("normal" | "pageLayout" | "pageBreakPreview")
    ) {
        return Err(Error::Compile(format!(
            "Worksheet.view on `{name}` must be `normal`, `pageLayout`, or `pageBreakPreview`; choose one supported worksheet view"
        )));
    }
    Ok(())
}

fn validate_worksheet_selection(sheet: &Node, name: &str) -> Result<()> {
    if let Some(value) = sheet.props.get("topLeftCell") {
        let reference = value.as_str().ok_or_else(|| Error::Compile(format!(
            "Worksheet.topLeftCell on `{name}` must be an A1 cell reference; use a value such as `A1` or `H6`"
        )))?;
        if parse_a1_cell_reference(reference).is_none() {
            return Err(Error::Compile(format!(
                "Worksheet.topLeftCell `{reference}` on `{name}` is invalid; use one cell reference such as `H6`"
            )));
        }
    }
    let Some(selection) = sheet.props.get("selection") else {
        return Ok(());
    };
    let Some(object) = selection.as_object() else {
        return Err(Error::Compile(format!(
            "Worksheet.selection on `{name}` must be an object with activeCell and range; provide `{{ activeCell: \"D3\", range: \"D3:F5\" }}`"
        )));
    };
    if object
        .keys()
        .any(|key| !matches!(key.as_str(), "activeCell" | "range"))
    {
        return Err(Error::Compile(format!(
            "Worksheet.selection on `{name}` only supports `activeCell` and `range`; remove other fields"
        )));
    }
    let active_cell = selection.get("activeCell").and_then(Value::as_str).ok_or_else(|| Error::Compile(format!(
        "Worksheet.selection.activeCell on `{name}` is required; provide a cell such as `D3`"
    )))?;
    let (active_row, active_column) = parse_a1_cell_reference(active_cell).ok_or_else(|| Error::Compile(format!(
        "Worksheet.selection.activeCell `{active_cell}` on `{name}` is invalid; use one Excel cell reference such as `D3`"
    )))?;
    let range = selection.get("range").and_then(Value::as_str).ok_or_else(|| Error::Compile(format!(
        "Worksheet.selection.range on `{name}` is required; provide one cell or contiguous range such as `D3:F5`"
    )))?;
    let parsed = parse_a1_range(range)
        .or_else(|| parse_a1_cell_reference(range).map(|(row, column)| (row, column, row, column)))
        .ok_or_else(|| Error::Compile(format!(
            "Worksheet.selection.range `{range}` on `{name}` is invalid; use one ordered cell/range such as `D3:F5`"
        )))?;
    if active_row < parsed.0
        || active_row > parsed.2
        || active_column < parsed.1
        || active_column > parsed.3
    {
        return Err(Error::Compile(format!(
            "Worksheet.selection.activeCell `{active_cell}` must be inside the selection range `{range}` on `{name}`; move the active cell into the selected cells"
        )));
    }
    if (active_row != parsed.0 && active_row != parsed.2)
        || (active_column != parsed.1 && active_column != parsed.3)
    {
        return Err(Error::Compile(format!(
            "Worksheet.selection.activeCell `{active_cell}` must be at a corner of range `{range}` because rust_xlsxwriter only exposes corner-based selection; choose one of the four range corners"
        )));
    }
    Ok(())
}

fn validate_worksheet_view_and_page_breaks(sheet: &Node, name: &str) -> Result<()> {
    validate_worksheet_view(sheet, name)?;
    validate_worksheet_selection(sheet, name)?;
    validate_worksheet_page_breaks(sheet, name)?;
    validate_print_gridlines(sheet, name)?;
    validate_worksheet_page_margins(sheet, name)?;
    validate_worksheet_protection(sheet, name)?;
    validate_worksheet_page_setup(sheet, name)?;
    validate_worksheet_print_scale_and_fit(sheet, name)
}

fn validate_worksheet_properties(sheet: &Node) -> Result<()> {
    reject_unknown_props(
        sheet,
        &[
            "name",
            "active",
            "selected",
            "firstTab",
            "tabColor",
            "visibility",
            "defaultRowHeightPixels",
            "hideUnusedRows",
            "showGridlines",
            "printGridlines",
            "zoom",
            "view",
            "topLeftCell",
            "selection",
            "horizontalPageBreaks",
            "verticalPageBreaks",
            "pageMargins",
            "printScale",
            "fitToPages",
            "orientation",
            "paperSize",
            "printArea",
            "repeatRows",
            "repeatColumns",
            "header",
            "footer",
            "headerFooterScaleWithDoc",
            "headerFooterAlignWithPage",
            "protection",
            "centerHorizontally",
            "centerVertically",
            "printBlackAndWhite",
            "printDraft",
            "printHeadings",
            "pageOrder",
            "rightToLeft",
            "summaryBelow",
            "summaryRight",
            "autoFilter",
            "autoFilterColumns",
            "freezeRows",
            "freezeColumns",
            "freezeTopLeftCell",
        ],
        "Worksheet",
    )
}

fn validate_worksheet_protection(sheet: &Node, name: &str) -> Result<()> {
    const OPTIONS: &[&str] = &[
        "selectLockedCells",
        "selectUnlockedCells",
        "formatCells",
        "formatColumns",
        "formatRows",
        "insertColumns",
        "insertRows",
        "insertLinks",
        "deleteColumns",
        "deleteRows",
        "sort",
        "useAutoFilter",
        "usePivotTables",
        "editScenarios",
        "editObjects",
    ];
    let Some(value) = sheet.props.get("protection") else {
        return Ok(());
    };
    let Some(options) = value.as_object() else {
        return Err(Error::Compile(format!(
            "Worksheet.protection on `{name}` must be an object of boolean permissions; use `protection={{}}` for default protection"
        )));
    };
    for (property, value) in options {
        if !OPTIONS.contains(&property.as_str()) {
            return Err(Error::Compile(format!(
                "Worksheet.protection on `{name}` has unsupported option `{property}`; use one of {}",
                OPTIONS.join(", ")
            )));
        }
        if !value.is_boolean() {
            return Err(Error::Compile(format!(
                "Worksheet.protection.{property} on `{name}` must be a boolean; set it to true/false or omit it to use the writer default"
            )));
        }
    }
    Ok(())
}

fn validate_worksheet_page_setup(sheet: &Node, name: &str) -> Result<()> {
    validate_worksheet_header_footer(sheet, name)?;
    if let Some(value) = sheet.props.get("orientation")
        && !matches!(value.as_str(), Some("portrait" | "landscape"))
    {
        return Err(Error::Compile(format!(
            "Worksheet.orientation on `{name}` must be `portrait` or `landscape`; choose one of the supported printed-page directions"
        )));
    }
    if let Some(value) = sheet.props.get("paperSize") {
        let Some(size) = value.as_u64() else {
            return Err(Error::Compile(format!(
                "Worksheet.paperSize on `{name}` must be an integer from 0 through 255; use an OOXML paper-size code such as 9 for A4"
            )));
        };
        if size > u64::from(u8::MAX) {
            return Err(Error::Compile(format!(
                "Worksheet.paperSize on `{name}` must be from 0 through 255; choose a paper-size code accepted by rust_xlsxwriter"
            )));
        }
    }
    for (property, expected_shape) in [
        ("printArea", None),
        ("repeatRows", Some("rows")),
        ("repeatColumns", Some("columns")),
    ] {
        let Some(value) = sheet.props.get(property) else {
            continue;
        };
        let Some(reference) = value.as_str() else {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` must be a print range string; use an A1 range such as `A1:D20`"
            )));
        };
        let Some(range) = parse_print_range(reference) else {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` has invalid or reversed range `{reference}`; use a valid inclusive A1 range"
            )));
        };
        let shape_matches = matches!(
            (expected_shape, range),
            (
                None,
                PrintRange::Cells { .. } | PrintRange::Rows { .. } | PrintRange::Columns { .. }
            ) | (Some("rows"), PrintRange::Rows { .. })
                | (Some("columns"), PrintRange::Columns { .. })
        );
        if !shape_matches {
            let shape = expected_shape.unwrap_or("cell, whole-row, or whole-column");
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` requires a {shape} range; use `1:3` for rows or `A:C` for columns"
            )));
        }
    }
    for property in [
        "centerHorizontally",
        "centerVertically",
        "printBlackAndWhite",
        "printDraft",
        "printHeadings",
    ] {
        if sheet
            .props
            .get(property)
            .is_some_and(|value| !value.is_boolean())
        {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` must be a boolean; use `true` to enable the print setting or `false` to leave it disabled"
            )));
        }
    }
    if let Some(value) = sheet.props.get("pageOrder")
        && !matches!(value.as_str(), Some("downThenOver" | "overThenDown"))
    {
        return Err(Error::Compile(format!(
            "Worksheet.pageOrder on `{name}` must be `downThenOver` or `overThenDown`; choose a supported printed-page traversal order"
        )));
    }
    Ok(())
}

fn validate_worksheet_header_footer(sheet: &Node, name: &str) -> Result<()> {
    for property in ["header", "footer"] {
        let Some(value) = sheet.props.get(property) else {
            continue;
        };
        let Some(text) = value.as_str() else {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` must be a string; use Excel header/footer text with optional `&L`, `&C`, or `&R` sections"
            )));
        };
        if expand_header_footer_aliases(text).chars().count() > 255 {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` exceeds Excel's 255-character limit after placeholder expansion; shorten the text or use a shorter placeholder"
            )));
        }
        if contains_header_footer_picture(text) {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` uses a picture placeholder, but header/footer images are not supported; remove `&G`/`&[Picture]` or provide text only"
            )));
        }
    }
    for property in ["headerFooterScaleWithDoc", "headerFooterAlignWithPage"] {
        if sheet
            .props
            .get(property)
            .is_some_and(|value| !value.is_boolean())
        {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` must be a boolean; use `false` to disable the Excel default or omit it to keep the default"
            )));
        }
    }
    Ok(())
}

fn expand_header_footer_aliases(text: &str) -> String {
    text.replace("&[Tab]", "&A")
        .replace("&[Date]", "&D")
        .replace("&[File]", "&F")
        .replace("&[Page]", "&P")
        .replace("&[Path]", "&Z")
        .replace("&[Time]", "&T")
        .replace("&[Pages]", "&N")
        .replace("&[Picture]", "&G")
}

fn contains_header_footer_picture(text: &str) -> bool {
    let mut remaining = text;
    while let Some(index) = remaining.find('&') {
        remaining = &remaining[index..];
        if remaining.starts_with("&&") {
            remaining = &remaining[2..];
        } else if remaining.starts_with("&G") || remaining.starts_with("&[Picture]") {
            return true;
        } else {
            remaining = &remaining[1..];
        }
    }
    false
}

fn validate_worksheet_reading_and_outline(sheet: &Node, name: &str) -> Result<()> {
    validate_worksheet_reading_direction(sheet, name)?;
    validate_outline_summary_directions(sheet, name)
}

fn validate_worksheet_visibility_and_active(sheet: &Node, index: usize, name: &str) -> Result<()> {
    validate_worksheet_visibility(sheet, index, name)?;
    let active = validate_active_worksheet(sheet, name)?;
    if let Some(value) = sheet.props.get("selected") {
        let Some(selected) = value.as_bool() else {
            return Err(Error::Compile(format!(
                "Worksheet.selected on `{name}` must be a boolean; use selected={{true}} or selected={{false}}"
            )));
        };
        if selected
            && sheet
                .props
                .get("visibility")
                .and_then(Value::as_str)
                .is_some_and(|visibility| visibility != "visible")
        {
            return Err(Error::Compile(format!(
                "selected Worksheet `{name}` must be visible; set visibility=\"visible\" or remove selected={{true}}"
            )));
        }
        if active && !selected {
            return Err(Error::Compile(format!(
                "Worksheet `{name}` cannot combine active={{true}} with selected={{false}} because rust_xlsxwriter selects activated worksheets; remove selected={{false}} or omit active={{true}}"
            )));
        }
    }
    Ok(())
}

fn validate_worksheet_print_scale_and_fit(sheet: &Node, name: &str) -> Result<()> {
    let print_scale = sheet.props.get("printScale");
    if let Some(value) = print_scale {
        let Some(scale) = value.as_u64() else {
            return Err(Error::Compile(format!(
                "Worksheet.printScale on `{name}` must be an integer from 10 through 400; use a value such as 80 or 100"
            )));
        };
        if !(10..=400).contains(&scale) {
            return Err(Error::Compile(format!(
                "Worksheet.printScale on `{name}` must be from 10 through 400; choose a valid printed-page percentage"
            )));
        }
    }
    let fit_to_pages = sheet.props.get("fitToPages");
    if let Some(value) = fit_to_pages {
        let Some(fit) = value.as_object() else {
            return Err(Error::Compile(format!(
                "Worksheet.fitToPages on `{name}` must be an object with integer `width` and `height`; use `{{width: 1, height: 0}}`"
            )));
        };
        for field in fit.keys() {
            if !["width", "height"].contains(&field.as_str()) {
                return Err(Error::Compile(format!(
                    "Worksheet.fitToPages on `{name}` has unknown field `{field}`; use only `width` and `height`"
                )));
            }
        }
        for field in ["width", "height"] {
            let Some(pages) = fit.get(field).and_then(Value::as_u64) else {
                return Err(Error::Compile(format!(
                    "Worksheet.fitToPages.{field} on `{name}` is required and must be an integer from 0 through 65535; use 0 for unlimited"
                )));
            };
            if pages > u64::from(u16::MAX) {
                return Err(Error::Compile(format!(
                    "Worksheet.fitToPages.{field} on `{name}` must be from 0 through 65535; use 0 for unlimited"
                )));
            }
        }
    }
    if print_scale.is_some() && fit_to_pages.is_some() {
        return Err(Error::Compile(format!(
            "Worksheet.printScale and Worksheet.fitToPages on `{name}` are mutually exclusive; keep one print scaling mode"
        )));
    }
    if fit_to_pages.is_some()
        && ["horizontalPageBreaks", "verticalPageBreaks"]
            .iter()
            .any(|key| {
                sheet
                    .props
                    .get(*key)
                    .and_then(Value::as_array)
                    .is_some_and(|breaks| !breaks.is_empty())
            })
    {
        return Err(Error::Compile(format!(
            "Worksheet.fitToPages on `{name}` overrides manual page breaks; remove fitToPages or the manual page breaks"
        )));
    }
    Ok(())
}

fn validate_worksheet_page_margins(sheet: &Node, name: &str) -> Result<()> {
    let Some(value) = sheet.props.get("pageMargins") else {
        return Ok(());
    };
    let Some(margins) = value.as_object() else {
        return Err(Error::Compile(format!(
            "Worksheet.pageMargins on `{name}` must be an object with one or more inch values; use `{{left: 0.5, top: 1}}`"
        )));
    };
    if margins.is_empty() {
        return Err(Error::Compile(format!(
            "Worksheet.pageMargins on `{name}` must set at least one margin; use one or more of left, right, top, bottom, header, footer"
        )));
    }
    for (key, margin) in margins {
        if !["left", "right", "top", "bottom", "header", "footer"].contains(&key.as_str()) {
            return Err(Error::Compile(format!(
                "Worksheet.pageMargins on `{name}` has unknown field `{key}`; use left, right, top, bottom, header, or footer"
            )));
        }
        if !margin
            .as_f64()
            .is_some_and(|value| value.is_finite() && value >= 0.0)
        {
            return Err(Error::Compile(format!(
                "Worksheet.pageMargins.{key} on `{name}` must be a finite nonnegative number of inches; correct the value or omit this margin"
            )));
        }
    }
    Ok(())
}

fn validate_worksheet_page_breaks(sheet: &Node, name: &str) -> Result<()> {
    for (property, limit) in [
        ("horizontalPageBreaks", 1_048_576_u64),
        ("verticalPageBreaks", 16_384_u64),
    ] {
        let Some(value) = sheet.props.get(property) else {
            continue;
        };
        let Some(indices) = value.as_array() else {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` must be an array of unique zero-based indices; use a value such as `[20, 40]`"
            )));
        };
        if indices.len() > 1023 {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` supports at most 1023 breaks; reduce the array length"
            )));
        }
        let mut seen = std::collections::HashSet::new();
        for index in indices {
            let Some(index) = index.as_u64() else {
                return Err(Error::Compile(format!(
                    "Worksheet.{property} on `{name}` must contain integers; use unique row/column indices"
                )));
            };
            if index == 0 || index >= limit || !seen.insert(index) {
                return Err(Error::Compile(format!(
                    "Worksheet.{property} on `{name}` must contain unique indices greater than 0 and below {limit}; remove duplicates or out-of-range values"
                )));
            }
        }
    }
    Ok(())
}

fn validate_outline_summary_directions(sheet: &Node, name: &str) -> Result<()> {
    for property in ["summaryBelow", "summaryRight"] {
        if sheet
            .props
            .get(property)
            .is_some_and(|value| !value.is_boolean())
        {
            return Err(Error::Compile(format!(
                "Worksheet.{property} on `{name}` must be a boolean; use false to place outline summaries above/left or true to use the default below/right position"
            )));
        }
    }
    Ok(())
}

fn validate_worksheet_reading_direction(sheet: &Node, name: &str) -> Result<()> {
    if sheet
        .props
        .get("rightToLeft")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(Error::Compile(format!(
            "Worksheet.rightToLeft on `{name}` must be a boolean; use true or false to choose the sheet reading direction"
        )));
    }
    Ok(())
}

fn validate_freeze_panes(sheet: &Node, sheet_name: &str) -> Result<()> {
    let freeze_rows = parse_worksheet_split(sheet, "freezeRows", sheet_name, 1_048_575)?;
    let freeze_columns = parse_worksheet_split(sheet, "freezeColumns", sheet_name, 16_383)?;
    let freeze_top_left = sheet.props.get("freezeTopLeftCell");
    let is_configured = sheet.props.contains_key("freezeRows")
        || sheet.props.contains_key("freezeColumns")
        || freeze_top_left.is_some();
    if is_configured && freeze_rows == 0 && freeze_columns == 0 {
        return Err(Error::Compile(format!(
            "Worksheet freezeRows/freezeColumns on `{sheet_name}` must freeze at least one row or column; set a positive split or remove the freeze properties"
        )));
    }
    if let Some(freeze_top_left) = freeze_top_left {
        let Some(reference) = freeze_top_left.as_str() else {
            return Err(Error::Compile(format!(
                "Worksheet.freezeTopLeftCell on `{sheet_name}` must be an A1 cell reference; use a value such as `A2`"
            )));
        };
        let Some((row, column)) = parse_a1_cell_reference(reference) else {
            return Err(Error::Compile(format!(
                "Worksheet.freezeTopLeftCell `{reference}` on `{sheet_name}` is not a valid A1 cell within Excel limits; use a reference such as `C4`"
            )));
        };
        if row < freeze_rows || u32::from(column) < freeze_columns {
            return Err(Error::Compile(format!(
                "Worksheet.freezeTopLeftCell `{reference}` on `{sheet_name}` must be outside the frozen {freeze_rows} rows and {freeze_columns} columns; move the scroll origin beyond the frozen area"
            )));
        }
    }
    Ok(())
}

fn validate_auto_filter_columns(sheet: &Node, sheet_name: &str) -> Result<()> {
    let Some(columns) = sheet.props.get("autoFilterColumns") else {
        return Ok(());
    };
    let Some(reference) = sheet.props.get("autoFilter").and_then(Value::as_str) else {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns on `{sheet_name}` requires Worksheet.autoFilter; add autoFilter with the complete range"
        )));
    };
    let (_, first_column, _, last_column) = parse_a1_range(reference).ok_or_else(|| {
        Error::Compile(format!(
            "Worksheet.autoFilterColumns on `{sheet_name}` cannot use invalid Worksheet.autoFilter `{reference}`; repair the range first"
        ))
    })?;
    let Value::Array(columns) = columns else {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns on `{sheet_name}` must be an array of filter objects; provide a JSON-compatible list"
        )));
    };
    if columns.is_empty() {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns on `{sheet_name}` must contain at least one column; remove the property or add criteria"
        )));
    }
    let mut seen_columns = std::collections::HashSet::new();
    let width = u64::from(last_column - first_column + 1);
    for (index, filter) in columns.iter().enumerate() {
        validate_auto_filter_column(filter, index, width, &mut seen_columns, sheet_name)?;
    }
    Ok(())
}

fn validate_auto_filter_column(
    filter: &Value,
    index: usize,
    width: u64,
    seen_columns: &mut std::collections::HashSet<u64>,
    sheet_name: &str,
) -> Result<()> {
    let Some(object) = filter.as_object() else {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}] on `{sheet_name}` must be an object with column and filter criteria; repair this entry"
        )));
    };
    if object
        .keys()
        .any(|key| !["column", "values", "includeBlanks", "customFilters"].contains(&key.as_str()))
    {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}] on `{sheet_name}` has an unsupported property; use column, values, includeBlanks, and/or customFilters"
        )));
    }
    let column = object.get("column").and_then(Value::as_u64).ok_or_else(|| {
        Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].column on `{sheet_name}` must be a zero-based integer; use a relative index such as 0"
        ))
    })?;
    if column >= width {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].column {column} is outside autoFilter width {width}; use an index from 0 to {}",
            width - 1
        )));
    }
    if !seen_columns.insert(column) {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns column {column} is repeated; combine criteria in one entry"
        )));
    }
    if let Some(include_blanks) = object.get("includeBlanks")
        && include_blanks.as_bool().is_none()
    {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].includeBlanks must be boolean; use true when blank cells should match"
        )));
    }
    let has_custom = object.contains_key("customFilters");
    let has_list = object.contains_key("values") || object.contains_key("includeBlanks");
    if has_custom && has_list {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}] cannot combine customFilters with values or includeBlanks; use exactly one filter kind"
        )));
    }
    if let Some(values) = object.get("values") {
        validate_auto_filter_values(values, index)?;
    }
    if let Some(custom) = object.get("customFilters") {
        validate_custom_filter_spec(custom, index)?;
    }
    let has_list_criteria = object.contains_key("values")
        || object.get("includeBlanks").and_then(Value::as_bool) == Some(true);
    if !has_list_criteria && !has_custom {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}] must have list criteria or customFilters; provide selected values, includeBlanks=true, or custom conditions"
        )));
    }
    Ok(())
}

fn validate_auto_filter_values(values: &Value, index: usize) -> Result<()> {
    let Some(values) = values.as_array() else {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].values must be a non-empty array of strings; provide selected values"
        )));
    };
    if values.is_empty() {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].values must not be empty; provide at least one selected string"
        )));
    }
    for value in values {
        let Some(value) = value.as_str() else {
            return Err(Error::Compile(format!(
                "Worksheet.autoFilterColumns[{index}].values must contain only strings; convert each filter value to text"
            )));
        };
        if !value.chars().all(is_valid_xml_character) {
            return Err(Error::Compile(format!(
                "Worksheet.autoFilterColumns[{index}].values contains an XML-invalid string; remove invalid XML control characters"
            )));
        }
    }
    Ok(())
}

fn validate_custom_filter_spec(custom: &Value, index: usize) -> Result<()> {
    let Some(custom) = custom.as_object() else {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters must be an object; provide logic and conditions"
        )));
    };
    if custom
        .keys()
        .any(|key| key != "logic" && key != "conditions")
    {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters has an unsupported property; use only logic and conditions"
        )));
    }
    let logic = custom.get("logic").and_then(Value::as_str).ok_or_else(|| {
        Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters.logic must be `and` or `or`; choose how two conditions combine"
        ))
    })?;
    if !["and", "or"].contains(&logic) {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters.logic `{logic}` is unsupported; use `and` or `or`"
        )));
    }
    let conditions = custom.get("conditions").and_then(Value::as_array).ok_or_else(|| {
        Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters.conditions must contain one or two conditions"
        ))
    })?;
    if !(1..=2).contains(&conditions.len()) {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters supports at most two custom conditions; remove extra conditions"
        )));
    }
    if conditions.len() == 1 && logic == "and" {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters.logic cannot be `and` with one condition; use `or` or add a second condition"
        )));
    }
    for condition in conditions {
        validate_custom_filter_condition(condition, index)?;
    }
    Ok(())
}

fn validate_custom_filter_condition(condition: &Value, index: usize) -> Result<()> {
    let Some(condition) = condition.as_object() else {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters condition must be an object; provide operator and value"
        )));
    };
    if condition
        .keys()
        .any(|key| key != "operator" && key != "value")
    {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters condition has an unsupported property; use operator and value"
        )));
    }
    let Some(value) = condition.get("value").and_then(Value::as_str) else {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters condition.value must be a string; provide the exact OOXML filter value"
        )));
    };
    if !value.chars().all(is_valid_xml_character) {
        return Err(Error::Compile(format!(
            "Worksheet.autoFilterColumns[{index}].customFilters condition.value contains invalid XML characters; remove XML control characters"
        )));
    }
    if let Some(operator) = condition.get("operator") {
        let Some(operator) = operator.as_str() else {
            return Err(Error::Compile(format!(
                "Worksheet.autoFilterColumns[{index}].customFilters operator must be a supported string"
            )));
        };
        if ![
            "lessThan",
            "notEqual",
            "greaterThan",
            "lessThanOrEqual",
            "greaterThanOrEqual",
        ]
        .contains(&operator)
        {
            return Err(Error::Compile(format!(
                "Worksheet.autoFilterColumns custom operator `{operator}` is unsupported; use lessThan, notEqual, greaterThan, lessThanOrEqual, greaterThanOrEqual, or omit operator for equality"
            )));
        }
    }
    Ok(())
}

fn parse_worksheet_split(
    sheet: &Node,
    property: &str,
    sheet_name: &str,
    maximum: u64,
) -> Result<u32> {
    let Some(value) = sheet.props.get(property) else {
        return Ok(0);
    };
    let Some(value) = value.as_u64() else {
        return Err(Error::Compile(format!(
            "Worksheet.{property} on `{sheet_name}` must be a non-negative integer; use a whole-number split count"
        )));
    };
    if value > maximum {
        return Err(Error::Compile(format!(
            "Worksheet.{property} on `{sheet_name}` exceeds the XLSX limit {maximum}; use a smaller freeze split"
        )));
    }
    u32::try_from(value).map_err(|error| Error::Compile(error.to_string()))
}

#[expect(
    clippy::too_many_lines,
    reason = "Worksheet child declarations share ordering and cross-component range validation state."
)]
fn validate_worksheet_children(
    sheet: &Node,
    sheet_name: &str,
    table_names: &mut std::collections::HashSet<String>,
) -> Result<()> {
    let mut local_name_set = std::collections::HashSet::new();
    let mut column_ranges = Vec::new();
    let mut column_outline = std::collections::BTreeMap::new();
    let mut row_outline = std::collections::BTreeMap::new();
    let mut merged_ranges: Vec<(usize, usize, usize, usize)> = Vec::new();
    let mut validation_ranges = Vec::new();
    let mut ignored_error_ranges = Vec::new();
    let mut table_ranges = Vec::new();
    let mut note_cells = std::collections::HashSet::new();
    let mut chart_nodes = Vec::new();
    let mut image_nodes = Vec::new();
    let mut chart_custom_width_ranges = Vec::new();
    let mut chart_custom_height_rows = std::collections::HashSet::new();
    let mut saw_row = false;
    let mut row_index = 0;
    for (child_index, child_value) in sheet.children.iter().enumerate() {
        let child: Node = serde_json::from_value(child_value.clone()).map_err(|error| {
            Error::Compile(format!(
                "Worksheet `{sheet_name}` child {child_index} must be Column, DataValidation, ConditionalFormat, IgnoreError, Table, or Row: {error}"
            ))
        })?;
        if child.r#type == "DefinedName" {
            if saw_row {
                return Err(Error::Compile(format!(
                    "Worksheet `{sheet_name}` DefinedName at child {} appears after a Row; move defined names before rows",
                    child_index + 1
                )));
            }
            validate_defined_name(
                &child,
                &format!("Worksheet `{sheet_name}`"),
                &mut local_name_set,
            )?;
            continue;
        }
        if child.r#type == "Note" {
            if saw_row {
                return Err(Error::Compile(format!(
                    "Worksheet `{sheet_name}` Note at child {} appears after a Row; move notes before rows",
                    child_index + 1
                )));
            }
            validate_note(&child, child_index, &mut note_cells)?;
            continue;
        }
        if child.r#type == "Chart" {
            if saw_row {
                return Err(Error::Compile(format!(
                    "Worksheet `{sheet_name}` Chart at child {} appears after a Row; move charts before rows",
                    child_index + 1
                )));
            }
            validate_chart(&child, child_index)?;
            chart_nodes.push(child);
            continue;
        }
        if child.r#type == "Image" {
            if saw_row {
                return Err(Error::Compile(format!(
                    "Worksheet `{sheet_name}` Image at child {} appears after a Row; move images before rows",
                    child_index + 1
                )));
            }
            validate_image(&child, child_index)?;
            image_nodes.push(child);
            continue;
        }
        if child.r#type == "Column" {
            if saw_row {
                return Err(Error::Compile(format!(
                    "Worksheet `{sheet_name}` Column at child {} appears after a Row; move all Column declarations before rows",
                    child_index + 1
                )));
            }
            validate_column(&child, child_index)?;
            let min = child.props.get("min").and_then(Value::as_u64).unwrap_or(1);
            let max = child
                .props
                .get("max")
                .and_then(Value::as_u64)
                .unwrap_or(min);
            let outline = OutlineProperties {
                hidden: child.props.get("hidden").and_then(Value::as_bool) == Some(true),
                level: child
                    .props
                    .get("outlineLevel")
                    .and_then(Value::as_u64)
                    .and_then(|level| u8::try_from(level).ok())
                    .unwrap_or_default(),
                collapsed: child.props.get("collapsed").and_then(Value::as_bool) == Some(true),
            };
            for column in min..=max {
                let coordinate = u32::try_from(column - 1).map_err(|error| {
                    Error::Compile(format!("Column outline coordinate is invalid: {error}"))
                })?;
                column_outline.insert(coordinate, outline);
            }
            if column_ranges
                .iter()
                .any(|(start, end): &(u64, u64)| min <= *end && max >= *start)
            {
                return Err(Error::Compile(format!(
                    "Column range {min}–{max} overlaps another Column declaration; combine the range or make column ranges disjoint"
                )));
            }
            column_ranges.push((min, max));
            if child
                .props
                .get("widthPixels")
                .and_then(Value::as_u64)
                .is_some_and(|width| width != 64)
                || child.props.get("hidden").and_then(Value::as_bool) == Some(true)
            {
                chart_custom_width_ranges.push((min, max));
            }
            continue;
        }
        if child.r#type == "DataValidation" {
            if saw_row {
                return Err(Error::Compile(format!(
                    "Worksheet `{sheet_name}` DataValidation at child {} appears after a Row; move worksheet metadata before rows",
                    child_index + 1
                )));
            }
            let range = validate_data_validation(&child, child_index)?;
            if validation_ranges
                .iter()
                .any(|prior: &(u32, u16, u32, u16)| {
                    range.0 <= prior.2
                        && range.2 >= prior.0
                        && range.1 <= prior.3
                        && range.3 >= prior.1
                })
            {
                return Err(Error::Compile(format!(
                    "DataValidation range `{}` overlaps another validation; combine the ranges or make them disjoint",
                    format_a1_range(range)
                )));
            }
            validation_ranges.push(range);
            continue;
        }
        if child.r#type == "ConditionalFormat" {
            if saw_row {
                return Err(Error::Compile(format!(
                    "Worksheet `{sheet_name}` ConditionalFormat at child {} appears after a Row; move conditional formatting before rows",
                    child_index + 1
                )));
            }
            validate_conditional_format(&child, child_index)?;
            continue;
        }
        if child.r#type == "IgnoreError" {
            if saw_row {
                return Err(Error::Compile(format!(
                    "Worksheet `{sheet_name}` IgnoreError at child {} appears after a Row; move worksheet metadata before rows",
                    child_index + 1
                )));
            }
            let range = validate_ignore_error(&child, child_index)?;
            if ignored_error_ranges
                .iter()
                .any(|prior| ranges_overlap(range, *prior))
            {
                return Err(Error::Compile(format!(
                    "IgnoreError range `{}` overlaps another ignored-error range; make ranges disjoint",
                    format_a1_range(range)
                )));
            }
            ignored_error_ranges.push(range);
            continue;
        }
        if child.r#type == "Table" {
            if saw_row {
                return Err(Error::Compile(format!(
                    "Worksheet `{sheet_name}` Table at child {} appears after a Row; move all Table declarations before rows",
                    child_index + 1
                )));
            }
            let range = validate_table(&child, child_index, table_names)?;
            if let Some(auto_filter) = sheet
                .props
                .get("autoFilter")
                .and_then(Value::as_str)
                .and_then(parse_a1_range)
                && ranges_overlap(range, auto_filter)
            {
                return Err(Error::Compile(format!(
                    "Table range `{}` overlaps Worksheet.autoFilter; remove the separate worksheet filter or move the table",
                    format_a1_range(range)
                )));
            }
            if table_ranges
                .iter()
                .any(|prior: &(u32, u16, u32, u16)| ranges_overlap(range, *prior))
            {
                return Err(Error::Compile(format!(
                    "Table range `{}` overlaps another table; make table ranges disjoint",
                    format_a1_range(range)
                )));
            }
            table_ranges.push(range);
            continue;
        }
        saw_row = true;
        let current_row_index = row_index;
        row_index += 1;
        if child.r#type != "Row" {
            return Err(Error::Compile(format!(
                "Worksheet `{sheet_name}` child {child_index} is `{}`; expected Row",
                child.r#type
            )));
        }
        validate_row_properties(&child, current_row_index)?;
        if child
            .props
            .get("height")
            .and_then(Value::as_f64)
            .is_some_and(|height| (height - 15.0).abs() > f64::EPSILON)
            || child.props.get("hidden").and_then(Value::as_bool) == Some(true)
        {
            chart_custom_height_rows.insert(current_row_index);
        }
        let coordinate = u32::try_from(current_row_index)
            .map_err(|error| Error::Compile(format!("invalid row outline coordinate: {error}")))?;
        row_outline.insert(
            coordinate,
            OutlineProperties {
                hidden: child.props.get("hidden").and_then(Value::as_bool) == Some(true),
                level: child
                    .props
                    .get("outlineLevel")
                    .and_then(Value::as_u64)
                    .and_then(|level| u8::try_from(level).ok())
                    .unwrap_or_default(),
                collapsed: child.props.get("collapsed").and_then(Value::as_bool) == Some(true),
            },
        );
        validate_empty_row_placement(sheet, &child, current_row_index, sheet_name, &merged_ranges)?;
        validate_row_cells(&child, current_row_index, &mut merged_ranges)?;
    }
    collapsed_outline_ranges(&row_outline, "row").map_err(Error::Compile)?;
    collapsed_outline_ranges(&column_outline, "column").map_err(Error::Compile)?;
    validate_chart_layouts(
        &chart_nodes,
        &chart_custom_width_ranges,
        &chart_custom_height_rows,
        sheet
            .props
            .get("defaultRowHeightPixels")
            .and_then(Value::as_u64)
            .is_some_and(|height| height != 20),
        sheet_name,
    )?;
    validate_image_layouts(
        &image_nodes,
        &chart_custom_width_ranges,
        &chart_custom_height_rows,
        sheet
            .props
            .get("defaultRowHeightPixels")
            .and_then(Value::as_u64)
            .is_some_and(|height| height != 20),
        sheet_name,
    )?;
    for range in table_ranges {
        if range.2 as usize >= row_index {
            return Err(Error::Compile(format!(
                "Table range `{}` extends beyond the declared rows; add the missing header/data rows",
                format_a1_range(range)
            )));
        }
        let headers = worksheet_table_headers(sheet, range)?;
        let unique: std::collections::HashSet<_> =
            headers.iter().map(|header| header.to_lowercase()).collect();
        if unique.len() != headers.len() {
            return Err(Error::Compile(format!(
                "Table range `{}` has duplicate header names; use a unique non-empty string for each header",
                format_a1_range(range)
            )));
        }
        if merged_ranges.iter().any(|(r0, c0, r1, c1)| {
            range.0 as usize <= *r1
                && range.2 as usize >= *r0
                && usize::from(range.1) <= *c1
                && usize::from(range.3) >= *c0
        }) {
            return Err(Error::Compile(format!(
                "Table range `{}` intersects merged cells; remove the merge or move the table",
                format_a1_range(range)
            )));
        }
    }
    Ok(())
}

fn validate_chart(chart: &Node, index: usize) -> Result<()> {
    reject_unknown_props(
        chart,
        &[
            "type", "cell", "width", "height", "name", "altText", "title",
        ],
        "Chart",
    )?;
    let chart_type = chart
        .props
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(format!(
                "Chart at declaration {} requires type: column, bar, line, or pie",
                index + 1
            ))
        })?;
    if !["column", "bar", "line", "pie"].contains(&chart_type) {
        return Err(Error::Compile(format!(
            "Chart.type `{chart_type}` is unsupported; use column, bar, line, or pie"
        )));
    }
    let cell = chart
        .props
        .get("cell")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile("Chart.cell is required; use an A1 anchor such as D2".into())
        })?;
    if parse_a1_cell_reference(cell).is_none() {
        return Err(Error::Compile(format!(
            "Chart.cell `{cell}` is not a valid A1 anchor; use a cell such as D2"
        )));
    }
    for dimension in ["width", "height"] {
        if let Some(value) = chart.props.get(dimension)
            && !value
                .as_u64()
                .is_some_and(|value| (1..=1790).contains(&value))
        {
            return Err(Error::Compile(format!(
                "Chart.{dimension} must be an integer from 1 to 1790 pixels; use a positive pixel size"
            )));
        }
    }
    for property in ["name", "altText", "title"] {
        if let Some(value) = chart.props.get(property) {
            let Some(value) = value.as_str() else {
                return Err(Error::Compile(format!(
                    "Chart.{property} must be a string; use quoted text or remove the property"
                )));
            };
            let maximum = 255;
            if value.is_empty()
                || value.chars().count() > maximum
                || !value.chars().all(is_valid_xml_character)
            {
                return Err(Error::Compile(format!(
                    "Chart.{property} must contain 1–{maximum} valid XML characters; shorten it or remove control characters"
                )));
            }
        }
    }
    validate_chart_series(chart)?;
    if chart_type == "pie" && chart.children.len() != 1 {
        return Err(Error::Compile(
            "pie Chart accepts exactly one Series; split the data into separate charts or choose column, bar, or line".into(),
        ));
    }
    Ok(())
}

fn validate_chart_series(chart: &Node) -> Result<()> {
    if chart.children.is_empty() {
        return Err(Error::Compile(
            "Chart requires at least one Series child".into(),
        ));
    }
    for (series_index, value) in chart.children.iter().enumerate() {
        let series: Node = serde_json::from_value(value.clone())
            .map_err(|error| Error::Compile(error.to_string()))?;
        if series.r#type != "Series" || !series.children.is_empty() {
            return Err(Error::Compile(format!(
                "Chart child {} must be an empty Series component",
                series_index + 1
            )));
        }
        reject_unknown_props(&series, &["values", "categories", "name"], "Series")?;
        let values = series
            .props
            .get("values")
            .and_then(Value::as_str)
            .filter(|value| !value.trim().is_empty())
            .ok_or_else(|| {
                Error::Compile(format!(
                    "Chart Series {} requires values as an A1 range formula",
                    series_index + 1
                ))
            })?;
        if !valid_chart_source_reference(values, true) {
            return Err(Error::Compile(format!(
                "Series.values `{values}` must be a sheet-qualified A1 range without a leading `=`; for example `Data!$B$2:$B$5`"
            )));
        }
        for key in ["categories", "name"] {
            if let Some(formula) = series.props.get(key).and_then(Value::as_str)
                && !valid_chart_source_reference(formula, key == "categories")
            {
                return Err(Error::Compile(format!(
                    "Series.{key} `{formula}` must be a sheet-qualified A1 reference without a leading `=`; use a cell/range such as `Data!$A$1` or `Data!$A$1:$A$5`"
                )));
            }
        }
    }
    Ok(())
}

fn valid_chart_source_reference(formula: &str, require_range: bool) -> bool {
    if formula.starts_with('=') || formula.matches('!').count() != 1 {
        return false;
    }
    let Some((_, reference)) = formula.split_once('!') else {
        return false;
    };
    if chart_source_sheet_name(formula).is_none() || reference.is_empty() {
        return false;
    }
    match reference.split_once(':') {
        Some((first, last)) => {
            parse_a1_cell_reference(first).is_some()
                && parse_a1_cell_reference(last).is_some()
                && parse_a1_range(reference).is_some()
        }
        None => !require_range && parse_a1_cell_reference(reference).is_some(),
    }
}

fn chart_source_sheet_name(formula: &str) -> Option<String> {
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
    } else {
        (!sheet.is_empty()
            && sheet.chars().all(|character| {
                character.is_ascii_alphanumeric() || matches!(character, '_' | '.')
            }))
        .then(|| sheet.to_owned())
    }
}

fn validate_chart_sheet_references(
    workbook_children: &[Value],
    worksheet_names: &std::collections::HashSet<String>,
) -> Result<()> {
    for (sheet_index, sheet_value) in workbook_children.iter().enumerate() {
        let sheet: Node = serde_json::from_value(sheet_value.clone()).map_err(|error| {
            Error::Compile(format!("Workbook child {sheet_index} is invalid: {error}"))
        })?;
        if sheet.r#type != "Worksheet" {
            continue;
        }
        for child_value in &sheet.children {
            let chart: Node = serde_json::from_value(child_value.clone())
                .map_err(|error| Error::Compile(error.to_string()))?;
            if chart.r#type != "Chart" {
                continue;
            }
            for series_value in &chart.children {
                let series: Node = serde_json::from_value(series_value.clone())
                    .map_err(|error| Error::Compile(error.to_string()))?;
                for property in ["values", "categories", "name"] {
                    let Some(formula) = series.props.get(property).and_then(Value::as_str) else {
                        continue;
                    };
                    let Some(referenced_sheet) = chart_source_sheet_name(formula) else {
                        continue;
                    };
                    if !worksheet_names.contains(&referenced_sheet.to_lowercase()) {
                        return Err(Error::Compile(format!(
                            "Chart Series.{property} `{formula}` on Worksheet `{}` refers to missing worksheet `{referenced_sheet}`; add that sheet or correct the reference",
                            sheet
                                .props
                                .get("name")
                                .and_then(Value::as_str)
                                .unwrap_or("?")
                        )));
                    }
                }
            }
        }
    }
    Ok(())
}

fn validate_chart_layouts(
    charts: &[Node],
    custom_width_ranges: &[(u64, u64)],
    custom_height_rows: &std::collections::HashSet<usize>,
    has_custom_default_height: bool,
    sheet_name: &str,
) -> Result<()> {
    for chart in charts {
        let cell = chart.props["cell"].as_str().expect("validated Chart.cell");
        let (first_row, first_column) =
            parse_a1_cell_reference(cell).expect("validated Chart.cell");
        let width = chart
            .props
            .get("width")
            .and_then(Value::as_u64)
            .unwrap_or(480);
        let height = chart
            .props
            .get("height")
            .and_then(Value::as_u64)
            .unwrap_or(288);
        let last_column = u64::from(first_column) + (width - 1) / 64;
        let last_row = usize::try_from(first_row).unwrap_or_default()
            + usize::try_from((height - 1) / 20).unwrap_or_default();
        if last_column >= 16_384 || last_row >= 1_048_576 {
            return Err(Error::Compile(format!(
                "Chart at {cell} on Worksheet `{sheet_name}` extends beyond Excel's worksheet grid; move it up/left or reduce its pixel dimensions"
            )));
        }
        let crosses_custom_width = custom_width_ranges
            .iter()
            .any(|(first, last)| *first <= last_column + 1 && *last > u64::from(first_column));
        let crosses_custom_height = custom_height_rows
            .iter()
            .any(|row| *row >= first_row as usize && *row <= last_row);
        if crosses_custom_width || crosses_custom_height || has_custom_default_height {
            return Err(Error::Compile(format!(
                "Chart at {cell} on Worksheet `{sheet_name}` crosses custom row/column dimensions; external chart reverse requires standard 64×20-pixel cells in the chart footprint. Move the chart or remove the intersecting custom widths/heights"
            )));
        }
    }
    Ok(())
}

fn validate_image(image: &Node, index: usize) -> Result<()> {
    if !image.children.is_empty() {
        return Err(Error::Compile(format!(
            "Image at declaration {} cannot have children; provide image bytes through data",
            index + 1
        )));
    }
    reject_unknown_props(
        image,
        &[
            "cell",
            "width",
            "height",
            "data",
            "altText",
            "decorative",
            "offsetX",
            "offsetY",
            "movement",
        ],
        "Image",
    )?;
    let cell = image
        .props
        .get("cell")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile("Image.cell is required; use an A1 cell such as A1".into())
        })?;
    if parse_a1_cell_reference(cell).is_none() {
        return Err(Error::Compile(format!(
            "Image.cell `{cell}` is not a valid A1 anchor; use a cell such as B2"
        )));
    }
    for property in ["width", "height"] {
        if !image
            .props
            .get(property)
            .and_then(Value::as_u64)
            .is_some_and(|value| (1..=1790).contains(&value))
        {
            return Err(Error::Compile(format!(
                "Image.{property} is required and must be an integer from 1 to 1790 pixels"
            )));
        }
    }
    for (property, maximum) in [("offsetX", 63), ("offsetY", 19)] {
        if let Some(value) = image.props.get(property)
            && value.as_u64().is_none_or(|value| value > maximum)
        {
            return Err(Error::Compile(format!(
                "Image.{property} must be an integer from 0 through {maximum} pixels within the anchored cell"
            )));
        }
    }
    if let Some(value) = image.props.get("movement") {
        let Some(value) = value.as_str() else {
            return Err(Error::Compile(
                "Image.movement must be a string; use moveAndSize, moveOnly, or fixed".into(),
            ));
        };
        if !["moveAndSize", "moveOnly", "fixed"].contains(&value) {
            return Err(Error::Compile(format!(
                "Image.movement `{value}` is unsupported; use moveAndSize, moveOnly, or fixed"
            )));
        }
    }
    validate_image_data(image)?;
    if let Some(value) = image.props.get("altText") {
        let Some(value) = value.as_str() else {
            return Err(Error::Compile(
                "Image.altText must be a string; use descriptive text or remove the property"
                    .into(),
            ));
        };
        if value.chars().count() > 255 || !value.chars().all(is_valid_xml_character) {
            return Err(Error::Compile(
                "Image.altText must contain at most 255 valid XML characters; shorten it or remove control characters".into(),
            ));
        }
    }
    if image
        .props
        .get("decorative")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(Error::Compile(
            "Image.decorative must be a boolean; use true or false".into(),
        ));
    }
    Ok(())
}

fn validate_image_data(image: &Node) -> Result<()> {
    let data_uri = image
        .props
        .get("data")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile("Image.data is required; use a valid base64 data URI".into())
        })?;
    let (media_type, encoded) = data_uri
        .strip_prefix("data:image/")
        .and_then(|value| value.split_once(";base64,"))
        .ok_or_else(|| {
            Error::Compile(
                "Image.data must be a valid base64 data URI; use data:image/png;base64,..., data:image/jpeg;base64,..., data:image/gif;base64,..., or data:image/bmp;base64,...".into(),
            )
        })?;
    let bytes = BASE64.decode(encoded).map_err(|error| {
        Error::Compile(format!(
            "Image.data is not valid base64 ({error}); encode the image bytes with standard base64"
        ))
    })?;
    let matches_type = match media_type {
        "png" => bytes.starts_with(b"\x89PNG\r\n\x1a\n"),
        "jpeg" => bytes.starts_with(&[0xff, 0xd8, 0xff]),
        "gif" => bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a"),
        "bmp" => bytes.starts_with(b"BM"),
        _ => false,
    };
    if !matches_type || bytes.is_empty() {
        return Err(Error::Compile(format!(
            "Image.data media type `{media_type}` does not match a supported image payload; use matching PNG, JPEG, GIF, or BMP bytes"
        )));
    }
    XlsxImage::new_from_buffer(&bytes).map_err(|error| {
        Error::Compile(format!(
            "Image.data is not a valid `{media_type}` image ({error}); use a complete supported image file encoded as base64"
        ))
    })?;
    Ok(())
}

fn validate_image_layouts(
    images: &[Node],
    custom_width_ranges: &[(u64, u64)],
    custom_height_rows: &std::collections::HashSet<usize>,
    has_custom_default_height: bool,
    sheet_name: &str,
) -> Result<()> {
    for image in images {
        let cell = image.props["cell"].as_str().expect("validated Image.cell");
        let (first_row, first_column) =
            parse_a1_cell_reference(cell).expect("validated Image.cell");
        let width = image.props["width"]
            .as_u64()
            .expect("validated Image.width");
        let height = image.props["height"]
            .as_u64()
            .expect("validated Image.height");
        let width_with_offset = image
            .props
            .get("offsetX")
            .and_then(Value::as_u64)
            .unwrap_or_default()
            + width;
        let height_with_offset = image
            .props
            .get("offsetY")
            .and_then(Value::as_u64)
            .unwrap_or_default()
            + height;
        let last_column = u64::from(first_column) + (width_with_offset - 1) / 64;
        let last_row = usize::try_from(first_row).unwrap_or_default()
            + usize::try_from((height_with_offset - 1) / 20).unwrap_or_default();
        if last_column >= 16_384 || last_row >= 1_048_576 {
            return Err(Error::Compile(format!(
                "Image at {cell} on Worksheet `{sheet_name}` extends beyond Excel's worksheet grid; move it up/left or reduce its pixel dimensions"
            )));
        }
        let crosses_custom_width = custom_width_ranges
            .iter()
            .any(|(first, last)| *first <= last_column + 1 && *last > u64::from(first_column));
        let crosses_custom_height = custom_height_rows
            .iter()
            .any(|row| *row >= first_row as usize && *row <= last_row);
        if crosses_custom_width || crosses_custom_height || has_custom_default_height {
            return Err(Error::Compile(format!(
                "Image at {cell} on Worksheet `{sheet_name}` crosses custom row/column dimensions; external image reverse requires standard 64×20-pixel cells in the image footprint. Move it or remove the intersecting custom widths/heights"
            )));
        }
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Note validation applies a single component-specific property whitelist and range audit."
)]
fn validate_note(
    note: &Node,
    declaration_index: usize,
    cells: &mut std::collections::HashSet<String>,
) -> Result<()> {
    if !note.children.is_empty() {
        return Err(Error::Compile(format!(
            "Note at declaration {} cannot have children; use its text property",
            declaration_index + 1
        )));
    }
    reject_unknown_props(
        note,
        &[
            "cell",
            "text",
            "author",
            "authorPrefix",
            "visible",
            "width",
            "height",
            "backgroundColor",
            "fontName",
            "fontSize",
            "fontFamily",
            "altText",
        ],
        "Note",
    )?;
    let cell = note
        .props
        .get("cell")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(format!(
                "Note at declaration {} requires cell, for example cell=\"B3\"",
                declaration_index + 1
            ))
        })?;
    let Some((row, column)) = parse_a1_cell_reference(cell) else {
        return Err(Error::Compile(format!(
            "Note.cell `{cell}` is not a valid worksheet cell; use an absolute A1 coordinate such as B3"
        )));
    };
    let canonical_cell = format_a1_cell(row, column);
    if !cells.insert(canonical_cell.clone()) {
        return Err(Error::Compile(format!(
            "more than one Note targets {canonical_cell}; keep one note per cell"
        )));
    }
    let text = note
        .props
        .get("text")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(format!(
                "Note at declaration {} requires text: string",
                declaration_index + 1
            ))
        })?;
    if text.chars().count() > 32_713
        || text
            .chars()
            .any(|character| !is_valid_xml_character(character))
    {
        return Err(Error::Compile(format!(
            "Note.text for {canonical_cell} exceeds the Excel note text limit or contains an invalid XML character; shorten it to 32,713 characters and use valid XML text"
        )));
    }
    if note.props.get("author").is_some_and(|value| {
        value
            .as_str()
            .is_none_or(|author| author.chars().count() > 52)
    }) {
        return Err(Error::Compile(format!(
            "Note.author for {canonical_cell} must be a string of at most 52 characters"
        )));
    }
    for property in ["authorPrefix", "visible"] {
        if note
            .props
            .get(property)
            .is_some_and(|value| !value.is_boolean())
        {
            return Err(Error::Compile(format!(
                "Note.{property} for {canonical_cell} must be a boolean"
            )));
        }
    }
    for property in ["width", "height"] {
        if note.props.get(property).is_some_and(|value| {
            value
                .as_u64()
                .is_none_or(|number| number == 0 || number > u64::from(u32::MAX))
        }) {
            return Err(Error::Compile(format!(
                "Note.{property} for {canonical_cell} must be a positive 32-bit integer number of pixels"
            )));
        }
    }
    if let Some(color) = note.props.get("backgroundColor") {
        validate_rgb_color(color, "backgroundColor", row as usize, usize::from(column))?;
    }
    if note.props.get("fontName").is_some_and(|value| {
        value
            .as_str()
            .is_none_or(|name| name.is_empty() || name.chars().count() > 255)
    }) {
        return Err(Error::Compile(format!(
            "Note.fontName for {canonical_cell} must contain 1–255 characters"
        )));
    }
    if note.props.get("fontSize").is_some_and(|value| {
        !value
            .as_f64()
            .is_some_and(|size| size.is_finite() && (1.0..=409.0).contains(&size))
    }) {
        return Err(Error::Compile(format!(
            "Note.fontSize for {canonical_cell} must be a finite number from 1 through 409 points"
        )));
    }
    if note.props.get("fontFamily").is_some_and(|value| {
        value
            .as_u64()
            .is_none_or(|family| family > u64::from(u8::MAX))
    }) {
        return Err(Error::Compile(format!(
            "Note.fontFamily for {canonical_cell} must be an integer from 0 through 255"
        )));
    }
    if note
        .props
        .get("altText")
        .is_some_and(|value| value.as_str().is_none_or(|text| text.chars().count() > 255))
    {
        return Err(Error::Compile(format!(
            "Note.altText for {canonical_cell} must be a string of at most 255 characters"
        )));
    }
    Ok(())
}

fn ranges_overlap(a: (u32, u16, u32, u16), b: (u32, u16, u32, u16)) -> bool {
    a.0 <= b.2 && a.2 >= b.0 && a.1 <= b.3 && a.3 >= b.1
}

fn validate_table(
    table: &Node,
    declaration_index: usize,
    names: &mut std::collections::HashSet<String>,
) -> Result<(u32, u16, u32, u16)> {
    for (index, value) in table.children.iter().enumerate() {
        let column: Node = serde_json::from_value(value.clone()).map_err(|error| {
            Error::Compile(format!("invalid TableColumn {}: {error}", index + 1))
        })?;
        validate_table_column(&column, index + 1)?;
    }
    reject_unknown_props(
        table,
        &[
            "range",
            "name",
            "style",
            "autoFilter",
            "bandedRows",
            "bandedColumns",
            "firstColumn",
            "lastColumn",
            "totalsRow",
            "altText",
            "altTextTitle",
        ],
        "Table",
    )?;
    let range_text = table
        .props
        .get("range")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(format!(
                "Table at declaration {} requires range, for example `A1:B4`",
                declaration_index + 1
            ))
        })?;
    let range = parse_a1_range(range_text).filter(|r| r.0 < r.2).ok_or_else(|| Error::Compile(format!("Table.range `{range_text}` must be a valid A1 range with a header and at least one data row; use a range such as `A1:B4`")))?;
    if table.props.get("totalsRow").and_then(Value::as_bool) == Some(true) && range.2 < range.0 + 2
    {
        return Err(Error::Compile(format!(
            "Table.range `{range_text}` must include a header, at least one data row, and the totals row; extend the range by one row"
        )));
    }
    let name = table
        .props
        .get("name")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile("Table.name is required; choose a valid Excel table name".into())
        })?;
    let valid_name = !name.is_empty()
        && name
            .chars()
            .next()
            .is_some_and(|c| c.is_ascii_alphabetic() || c == '_')
        && name
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.')
        && !name.contains('.')
        && parse_a1_cell_reference(name).is_none();
    if !valid_name || !names.insert(name.to_lowercase()) {
        return Err(Error::Compile(format!(
            "Table.name `{name}` is invalid or duplicated; use a unique Excel name starting with a letter or underscore, without spaces or cell-reference syntax"
        )));
    }
    if let Some(style) = table.props.get("style").and_then(Value::as_str)
        && !is_supported_table_style(style)
    {
        return Err(Error::Compile(format!(
            "Table.style `{style}` is unsupported; use TableStyleNone, TableStyleLight1–21, TableStyleMedium1–28, or TableStyleDark1–11"
        )));
    }
    for key in [
        "autoFilter",
        "bandedRows",
        "bandedColumns",
        "firstColumn",
        "lastColumn",
        "totalsRow",
    ] {
        if table.props.get(key).is_some_and(|v| v.as_bool().is_none()) {
            return Err(Error::Compile(format!(
                "Table.{key} must be a boolean; use `{key}={{true}}` or `{key}={{false}}`"
            )));
        }
    }
    validate_table_accessibility_text(table)?;
    let width = usize::from(range.3 - range.1 + 1);
    if !table.children.is_empty() && table.children.len() != width {
        return Err(Error::Compile(format!(
            "Table has {} TableColumn children but range `{range_text}` has {width} columns; provide one TableColumn per column or omit all TableColumn children",
            table.children.len()
        )));
    }
    Ok(range)
}

fn validate_table_accessibility_text(table: &Node) -> Result<()> {
    for key in ["altText", "altTextTitle"] {
        if let Some(value) = table.props.get(key) {
            let Some(text) = value.as_str() else {
                return Err(Error::Compile(format!(
                    "Table.{key} must be a string; provide accessibility text or omit the property"
                )));
            };
            if !text.chars().all(is_valid_xml_character) {
                return Err(Error::Compile(format!(
                    "Table.{key} contains an invalid XML character; remove control characters and use valid text"
                )));
            }
        }
    }
    Ok(())
}

fn validate_table_column(column: &Node, index: usize) -> Result<()> {
    if column.r#type != "TableColumn" || !column.children.is_empty() {
        return Err(Error::Compile(format!(
            "Table child {index} must be an empty TableColumn; remove other children"
        )));
    }
    reject_unknown_props(
        column,
        &[
            "totalLabel",
            "totalFunction",
            "totalFormula",
            "formula",
            "format",
            "headerFormat",
        ],
        "TableColumn",
    )?;
    if column
        .props
        .get("totalLabel")
        .is_some_and(|v| !v.is_string())
        || column
            .props
            .get("formula")
            .is_some_and(|v| v.as_str().is_none_or(str::is_empty))
        || column
            .props
            .get("totalFormula")
            .is_some_and(|v| v.as_str().is_none_or(str::is_empty))
    {
        return Err(Error::Compile(format!(
            "TableColumn {index} has an invalid label or formula; use non-empty strings"
        )));
    }
    if let Some(value) = column.props.get("totalFunction").and_then(Value::as_str)
        && ![
            "none",
            "average",
            "count",
            "countNumbers",
            "max",
            "min",
            "sum",
            "stdDev",
            "var",
        ]
        .contains(&value)
    {
        return Err(Error::Compile(format!(
            "TableColumn.totalFunction `{value}` is unsupported; use none, average, count, countNumbers, max, min, sum, stdDev, or var"
        )));
    }
    if column.props.contains_key("totalFormula") && column.props.contains_key("totalFunction") {
        return Err(Error::Compile(
            "TableColumn.totalFormula and totalFunction are mutually exclusive; specify only one"
                .into(),
        ));
    }
    if column.props.contains_key("totalLabel")
        && (column.props.contains_key("totalFormula") || column.props.contains_key("totalFunction"))
    {
        return Err(Error::Compile(
            "TableColumn.totalLabel cannot be combined with totalFunction or totalFormula; specify a label or a calculation"
                .into(),
        ));
    }
    for property in ["format", "headerFormat"] {
        if let Some(value) = column.props.get(property) {
            validate_table_column_format(value, property, index)?;
        }
    }
    Ok(())
}

fn validate_table_column_format(value: &Value, property: &str, index: usize) -> Result<()> {
    let Some(format) = value.as_object() else {
        return Err(Error::Compile(format!(
            "TableColumn.{property} must be an object; use font emphasis, RGB colors, numberFormat, or supported border edges"
        )));
    };
    if format.is_empty() {
        return Err(Error::Compile(format!(
            "TableColumn.{property} cannot be empty; add a supported style property"
        )));
    }
    for (key, value) in format {
        validate_table_column_format_property(key, value, property, index)?;
    }
    let has_pattern = format.contains_key("fillPattern");
    let has_color = format.contains_key("fillColor");
    let has_pattern_colors =
        format.contains_key("fillForegroundColor") || format.contains_key("fillBackgroundColor");
    if ((has_pattern || has_pattern_colors) && (!has_pattern || has_color))
        || (format.get("fillPattern").and_then(Value::as_str) == Some("none") && has_pattern_colors)
    {
        return Err(Error::Compile(format!(
            "TableColumn.{property} fill properties conflict: fillColor is the solid-fill shorthand; foreground/background colors require a non-none fillPattern. Choose one fill representation"
        )));
    }
    if property == "format" && format.contains_key("script") && format.len() == 1 {
        return Err(Error::Compile(format!(
            "TableColumn.{property}.script requires another TableColumn differential-format property; the pinned writer drops script-only dxf styles, so add a supported property such as bold or fontColor"
        )));
    }
    for suffix in ["Top", "Right", "Bottom", "Left"] {
        let style_key = format!("border{suffix}");
        let color_key = format!("border{suffix}Color");
        if format.contains_key(&color_key)
            && format
                .get(&style_key)
                .and_then(Value::as_str)
                .is_none_or(|style| style == "none")
        {
            return Err(Error::Compile(format!(
                "TableColumn.{property}.{color_key} requires a matching non-`none` {style_key}; add the edge style or remove the color"
            )));
        }
    }
    Ok(())
}

fn validate_table_column_format_property(
    key: &str,
    value: &Value,
    property: &str,
    index: usize,
) -> Result<()> {
    match key {
        "bold" | "italic" | "strikethrough" if value.as_bool() == Some(true) => Ok(()),
        "underline"
            if value.as_str().is_some_and(|underline| {
                ["single", "double", "singleAccounting", "doubleAccounting"].contains(&underline)
            }) =>
        {
            Ok(())
        }
        "script"
            if value
                .as_str()
                .is_some_and(|script| ["superscript", "subscript"].contains(&script)) =>
        {
            Ok(())
        }
        "numberFormat" => validate_table_column_number_format(value, property),
        "fontName" | "fontSize" | "fontFamily" | "fontCharset" => {
            validate_table_column_font_property(key, value, property)
        }
        "fillPattern"
            if value
                .as_str()
                .is_some_and(|pattern| FILL_PATTERN_NAMES.contains(&pattern)) =>
        {
            Ok(())
        }
        "fillPattern" => Err(Error::Compile(format!(
            "TableColumn.{property}.fillPattern must be one of {}; choose a supported OOXML pattern",
            FILL_PATTERN_NAMES.join(", ")
        ))),
        "fontColor"
        | "fillColor"
        | "borderTopColor"
        | "borderRightColor"
        | "borderBottomColor"
        | "borderLeftColor"
        | "fillForegroundColor"
        | "fillBackgroundColor" => validate_rgb_color(value, key, index, 0),
        "borderTop" | "borderRight" | "borderBottom" | "borderLeft" => {
            let Some(style) = value.as_str() else {
                return Err(Error::Compile(format!(
                    "TableColumn.{property}.{key} must be a quoted border style"
                )));
            };
            if ![
                "none",
                "thin",
                "medium",
                "dashed",
                "dotted",
                "thick",
                "double",
                "hair",
                "mediumDashed",
                "dashDot",
                "mediumDashDot",
                "dashDotDot",
                "mediumDashDotDot",
                "slantDashDot",
            ]
            .contains(&style)
            {
                return Err(Error::Compile(format!(
                    "TableColumn.{property}.{key} has unsupported border style `{style}`; use an OOXML edge style"
                )));
            }
            Ok(())
        }
        "borderDiagonalType"
            if value
                .as_str()
                .is_some_and(|direction| ["up", "down", "upDown"].contains(&direction)) =>
        {
            Ok(())
        }
        "script" => Err(Error::Compile(format!(
            "TableColumn.{property}.script must be `superscript` or `subscript`; choose one of the OOXML vertAlign values supported by the writer"
        ))),
        "borderDiagonal" | "borderDiagonalColor" => Err(Error::Compile(format!(
            "TableColumn.{property}.{key} is not serializable by the pinned writer in a differential format; use borderDiagonalType only or move the style to a worksheet Cell"
        ))),
        "bold" | "italic" | "strikethrough" => Err(Error::Compile(format!(
            "TableColumn.{property}.{key} must be true; table differential formats can add emphasis but cannot encode a false override"
        ))),
        _ => Err(Error::Compile(format!(
            "TableColumn.{property}.{key} is not supported; use font name/size/family/charset, bold/italic/strikethrough, underline, script, RGB colors, numberFormat, or supported edge borders"
        ))),
    }
}

fn validate_table_column_font_property(key: &str, value: &Value, property: &str) -> Result<()> {
    if property != "headerFormat" {
        return Err(Error::Compile(format!(
            "TableColumn.format.{key} is not serialized by the pinned writer; move this font metadata to headerFormat or the header Cell"
        )));
    }
    match key {
        "fontName" => {
            let Some(font_name) = value.as_str() else {
                return Err(Error::Compile(
                    "TableColumn.headerFormat.fontName must be a string; use a literal font family name".into(),
                ));
            };
            if font_name.is_empty()
                || font_name.chars().count() > 255
                || !font_name.chars().all(is_valid_xml_character)
            {
                return Err(Error::Compile(
                    "TableColumn.headerFormat.fontName must contain 1 to 255 valid XML characters; shorten it or remove control characters".into(),
                ));
            }
        }
        "fontSize" => {
            let Some(font_size) = value.as_f64() else {
                return Err(Error::Compile(
                    "TableColumn.headerFormat.fontSize must be a number in points from 1 through 409".into(),
                ));
            };
            if !(1.0..=409.0).contains(&font_size) {
                return Err(Error::Compile(
                    "TableColumn.headerFormat.fontSize must be from 1 through 409 points; choose an OOXML-compatible font size".into(),
                ));
            }
        }
        "fontFamily" => {
            let Some(font_family) = value.as_u64() else {
                return Err(Error::Compile(
                    "TableColumn.headerFormat.fontFamily must be an integer from 0 through 5; use an OOXML font family value".into(),
                ));
            };
            if font_family > 5 {
                return Err(Error::Compile(
                    "TableColumn.headerFormat.fontFamily must be from 0 through 5; use 0 (default), 1 (roman), 2 (swiss), 3 (modern), 4 (script), or 5 (decorative)".into(),
                ));
            }
        }
        "fontCharset" => {
            if value
                .as_u64()
                .is_none_or(|charset| u8::try_from(charset).is_err())
            {
                return Err(Error::Compile(
                    "TableColumn.headerFormat.fontCharset must be an integer byte from 0 through 255; choose an OOXML charset value".into(),
                ));
            }
        }
        _ => {
            return Err(Error::Compile(format!(
                "TableColumn.headerFormat.{key} is not a supported font metadata property"
            )));
        }
    }
    Ok(())
}

fn validate_table_column_number_format(value: &Value, property: &str) -> Result<()> {
    let Some(number_format) = value.as_str() else {
        return Err(Error::Compile(format!(
            "TableColumn.{property}.numberFormat must be a string; use an Excel format code such as `#,##0.00`"
        )));
    };
    if number_format.is_empty() || number_format.chars().count() > 255 {
        return Err(Error::Compile(format!(
            "TableColumn.{property}.numberFormat must contain 1 to 255 characters; shorten or provide an Excel format code"
        )));
    }
    if !number_format.chars().all(is_valid_xml_character) {
        return Err(Error::Compile(format!(
            "TableColumn.{property}.numberFormat contains a character that is invalid in XML; remove control characters and use a valid Excel format code"
        )));
    }
    Ok(())
}

#[must_use]
pub fn is_supported_table_style(style: &str) -> bool {
    if style == "TableStyleNone" {
        return true;
    }
    [
        ("TableStyleLight", 21),
        ("TableStyleMedium", 28),
        ("TableStyleDark", 11),
    ]
    .iter()
    .any(|(prefix, max)| {
        style
            .strip_prefix(prefix)
            .and_then(|n| n.parse::<u8>().ok())
            .is_some_and(|n| (1..=*max).contains(&n))
    })
}

fn worksheet_table_headers(sheet: &Node, range: (u32, u16, u32, u16)) -> Result<Vec<String>> {
    let row_value = sheet
        .children
        .iter()
        .filter(|value| value.get("type").and_then(Value::as_str) == Some("Row"))
        .nth(range.0 as usize)
        .ok_or_else(|| Error::Compile("Table header row is missing".into()))?;
    let row: Node = serde_json::from_value(row_value.clone())
        .map_err(|error| Error::Compile(error.to_string()))?;
    let mut headers = Vec::new();
    for column in range.1..=range.3 {
        let cell_value = row.children.get(usize::from(column)).ok_or_else(|| {
            Error::Compile(format!(
                "Table header cell {} is missing; provide every header cell in the first table row",
                format_a1_cell(range.0, column)
            ))
        })?;
        let cell: Node = serde_json::from_value(cell_value.clone())
            .map_err(|error| Error::Compile(error.to_string()))?;
        if cell.r#type != "Cell" {
            return Err(Error::Compile(
                "Table headers must be Cell components".into(),
            ));
        }
        let header = cell
            .props
            .get("value")
            .and_then(Value::as_str)
            .or_else(|| cell.children.first().and_then(Value::as_str))
            .ok_or_else(|| {
                Error::Compile(format!(
                    "Table header cell {} must contain a non-empty string",
                    format_a1_cell(range.0, column)
                ))
            })?;
        if header.trim().is_empty() {
            return Err(Error::Compile(format!(
                "Table header cell {} is empty; provide a unique non-empty string",
                format_a1_cell(range.0, column)
            )));
        }
        if cell
            .props
            .keys()
            .any(|key| key != "value" && !FORMAT_PROPERTIES.contains(&key.as_str()))
        {
            return Err(Error::Compile(format!(
                "Table header cell {} has an unsupported property; use value and supported Cell format properties",
                format_a1_cell(range.0, column)
            )));
        }
        headers.push(header.to_owned());
    }
    Ok(headers)
}

const IGNORE_ERROR_TYPES: &[&str] = &[
    "numberStoredAsText",
    "evalError",
    "formula",
    "emptyCellReference",
    "formulaRange",
    "listDataValidation",
    "TwoDigitTextYear",
    "unlockedFormula",
    "calculatedColumn",
];

fn validate_ignore_error(
    component: &Node,
    declaration_index: usize,
) -> Result<(u32, u16, u32, u16)> {
    if !component.children.is_empty() {
        return Err(Error::Compile(format!(
            "IgnoreError at declaration {} cannot have children; remove its children",
            declaration_index + 1
        )));
    }
    reject_unknown_props(component, &["range", "errorType"], "IgnoreError")?;
    let range = component
        .props
        .get("range")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile("IgnoreError.range is required; use an A1 range such as `A1:B3`".into())
        })?;
    let parsed = parse_a1_range(range).ok_or_else(|| {
        Error::Compile(format!(
            "IgnoreError.range `{range}` is invalid; use one ordered A1 rectangle such as `A1:B3`"
        ))
    })?;
    let error_type = component
        .props
        .get("errorType")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(
                "IgnoreError.errorType is required; choose one supported Excel error-checking flag"
                    .into(),
            )
        })?;
    if !IGNORE_ERROR_TYPES.contains(&error_type) {
        return Err(Error::Compile(format!(
            "IgnoreError.errorType `{error_type}` is unsupported; use one of {}",
            IGNORE_ERROR_TYPES.join(", ")
        )));
    }
    Ok(parsed)
}

fn validate_data_validation(
    validation: &Node,
    declaration_index: usize,
) -> Result<(u32, u16, u32, u16)> {
    if !validation.children.is_empty() {
        return Err(Error::Compile(format!(
            "DataValidation at declaration {} cannot have children; put the validation on a worksheet metadata component",
            declaration_index + 1
        )));
    }
    reject_unknown_props(
        validation,
        &[
            "range",
            "type",
            "source",
            "formula1",
            "formula2",
            "operator",
            "allowBlank",
            "showDropDown",
            "showInputMessage",
            "showErrorMessage",
            "inputTitle",
            "inputMessage",
            "errorTitle",
            "errorMessage",
            "errorStyle",
        ],
        "DataValidation",
    )?;
    let range = validation
        .props
        .get("range")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(format!(
                "DataValidation.range at declaration {} must be an A1 range such as `B2:B20`",
                declaration_index + 1
            ))
        })?;
    let range = parse_a1_range(range).ok_or_else(|| {
        Error::Compile(format!(
            "DataValidation.range `{range}` is invalid or outside worksheet limits; use an ordered A1 range such as `B2:B20`"
        ))
    })?;
    validate_data_validation_rule(validation, declaration_index)?;
    for property in [
        "allowBlank",
        "showDropDown",
        "showInputMessage",
        "showErrorMessage",
    ] {
        if validation
            .props
            .get(property)
            .is_some_and(|value| value.as_bool().is_none())
        {
            return Err(Error::Compile(format!(
                "DataValidation.{property} at declaration {} must be a boolean; use `{property}={{true}}` or `{property}={{false}}`",
                declaration_index + 1
            )));
        }
    }
    for property in ["inputTitle", "inputMessage", "errorTitle", "errorMessage"] {
        if let Some(value) = validation.props.get(property) {
            let Some(value) = value.as_str() else {
                return Err(Error::Compile(format!(
                    "DataValidation.{property} at declaration {} must be a string; provide plain text",
                    declaration_index + 1
                )));
            };
            let max_length = if property.ends_with("Title") { 32 } else { 255 };
            validate_data_validation_text(property, value, declaration_index, max_length)?;
        }
    }
    if let Some(error_style) = validation.props.get("errorStyle") {
        let Some(error_style) = error_style.as_str() else {
            return Err(Error::Compile(format!(
                "DataValidation.errorStyle at declaration {} must be one of `stop`, `warning`, or `information`",
                declaration_index + 1
            )));
        };
        if !["stop", "warning", "information"].contains(&error_style) {
            return Err(Error::Compile(format!(
                "DataValidation.errorStyle `{error_style}` is unsupported; use `stop`, `warning`, or `information`"
            )));
        }
    }
    Ok(range)
}

fn validate_data_validation_rule(validation: &Node, declaration_index: usize) -> Result<()> {
    let validation_type = validation
        .props
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(format!(
                "DataValidation.type at declaration {} must name a supported validation type",
                declaration_index + 1
            ))
        })?;
    if validation_type == "list" {
        return validate_data_validation_list_rule(validation, declaration_index);
    }
    if validation_type == "any" {
        for property in ["source", "formula1", "formula2", "operator"] {
            if validation.props.contains_key(property) {
                return Err(Error::Compile(format!(
                    "DataValidation type `any` does not accept {property}; remove the formula fields and use prompt/message properties for help text"
                )));
            }
        }
        if validation
            .props
            .get("errorStyle")
            .and_then(Value::as_str)
            .is_some_and(|style| style != "stop")
        {
            return Err(Error::Compile(
                "DataValidation type `any` only supports the default errorStyle `stop`; choose `stop` or remove errorStyle because the writer omits non-default styles for any-value rules".into()
            ));
        }
        return Ok(());
    }
    if !["whole", "decimal", "date", "time", "textLength", "custom"].contains(&validation_type) {
        return Err(Error::Compile(format!(
            "DataValidation.type `{validation_type}` is unsupported; use list, whole, decimal, date, time, textLength, or custom"
        )));
    }
    if validation.props.contains_key("source") {
        return Err(Error::Compile(format!(
            "DataValidation.source is only valid for type=\"list\"; use formula1 for `{validation_type}`"
        )));
    }
    let formula1 = validation
        .props
        .get("formula1")
        .and_then(Value::as_str)
        .filter(|formula| !formula.is_empty())
        .ok_or_else(|| {
            Error::Compile(format!(
                "DataValidation.formula1 at declaration {} is required for type `{validation_type}`; provide the first bound or formula",
                declaration_index + 1
            ))
        })?;
    validate_data_validation_text("formula1", formula1, declaration_index, 8192)?;
    let formula2 = validation.props.get("formula2");
    if let Some(formula2) = formula2 {
        let Some(formula2) = formula2.as_str().filter(|value| !value.is_empty()) else {
            return Err(Error::Compile(format!(
                "DataValidation.formula2 at declaration {} must be a non-empty string; provide the second bound",
                declaration_index + 1
            )));
        };
        validate_data_validation_text("formula2", formula2, declaration_index, 8192)?;
    }
    let operator = validate_data_validation_operator(validation, declaration_index)?;
    if validation_type == "custom" {
        if operator.is_some() || formula2.is_some() {
            return Err(Error::Compile(
                "DataValidation type `custom` accepts formula1 only; remove operator and formula2"
                    .into(),
            ));
        }
        return Ok(());
    }
    let is_two_bound = matches!(operator, Some("between" | "notBetween")) || formula2.is_some();
    if is_two_bound && formula2.is_none() {
        return Err(Error::Compile(format!(
            "DataValidation operator `{}` requires formula2; add the upper bound",
            operator.unwrap_or("between")
        )));
    }
    if formula2.is_some()
        && operator.is_some_and(|value| !["between", "notBetween"].contains(&value))
    {
        return Err(Error::Compile(
            "DataValidation.formula2 requires operator between/notBetween or an omitted operator; remove formula2 or choose a two-bound rule".into(),
        ));
    }
    if operator.is_none() && formula2.is_none() {
        return Err(Error::Compile(format!(
            "DataValidation `{validation_type}` requires operator for a single bound; use an explicit comparison or provide formula2 for between"
        )));
    }
    Ok(())
}

fn validate_data_validation_operator(
    validation: &Node,
    declaration_index: usize,
) -> Result<Option<&str>> {
    let Some(operator) = validation.props.get("operator") else {
        return Ok(None);
    };
    let Some(operator) = operator.as_str() else {
        return Err(Error::Compile(format!(
            "DataValidation.operator at declaration {} must be a supported comparison string",
            declaration_index + 1
        )));
    };
    if ![
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
        return Err(Error::Compile(format!(
            "DataValidation.operator `{operator}` is unsupported; use between, notBetween, equal, notEqual, greaterThan, greaterThanOrEqual, lessThan, or lessThanOrEqual"
        )));
    }
    Ok(Some(operator))
}

fn validate_data_validation_list_rule(validation: &Node, declaration_index: usize) -> Result<()> {
    if validation.props.contains_key("formula1")
        || validation.props.contains_key("formula2")
        || validation.props.contains_key("operator")
    {
        return Err(Error::Compile(
            "DataValidation list rules use source, not formula1, formula2, or operator; remove those properties".into(),
        ));
    }
    let source = validation
        .props
        .get("source")
        .and_then(Value::as_str)
        .filter(|source| !source.is_empty())
        .ok_or_else(|| {
            Error::Compile(format!(
                "DataValidation.source at declaration {} must be a non-empty list source formula; use a quoted list such as `\"Open,Closed\"` or an A1 reference",
                declaration_index + 1
            ))
        })?;
    validate_data_validation_text("source", source, declaration_index, 255)
}

#[expect(
    clippy::too_many_lines,
    reason = "Conditional format family dispatch and shared input validation are kept together."
)]
fn validate_conditional_format(format: &Node, declaration_index: usize) -> Result<()> {
    if !format.children.is_empty() {
        return Err(Error::Compile(format!(
            "ConditionalFormat at declaration {} cannot have children; put it on a Worksheet",
            declaration_index + 1
        )));
    }
    reject_unknown_props(
        format,
        &[
            "range",
            "type",
            "stopIfTrue",
            "minColor",
            "midColor",
            "maxColor",
            "minType",
            "midType",
            "maxType",
            "minValue",
            "midValue",
            "maxValue",
            "operator",
            "formula1",
            "formula2",
            "text",
            "rule",
            "rank",
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
            "barOnly",
            "solidFill",
            "border",
            "direction",
            "axisPosition",
            "borderColor",
            "negativeFillColor",
            "negativeBorderColor",
            "axisColor",
            "iconType",
            "reverse",
            "showIconsOnly",
            "icons",
        ],
        "ConditionalFormat",
    )?;
    let range = format
        .props
        .get("range")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(format!(
                "ConditionalFormat.range at declaration {} must be an A1 range such as `B2:B20`",
                declaration_index + 1
            ))
        })?;
    if parse_a1_range(range).is_none() {
        return Err(Error::Compile(format!(
            "ConditionalFormat.range `{range}` is invalid or outside worksheet limits; use one ordered A1 range such as `B2:B20`"
        )));
    }
    let scale_type = format
        .props
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(
                "ConditionalFormat.type must be colorScale2, colorScale3, cellIs, formula, text, dataBar, iconSet, top10, average, date, blank, error, or duplicate"
                    .into(),
            )
        })?;
    let rule_types = ["top10", "average", "date", "blank", "error", "duplicate"];
    if !rule_types.contains(&scale_type) && format.props.contains_key("rule") {
        return Err(Error::Compile(
            "ConditionalFormat.rule is only valid with a rule-based conditional format type; remove it or change the type".into(),
        ));
    }
    if scale_type != "top10" && format.props.contains_key("rank") {
        return Err(Error::Compile(
            "ConditionalFormat.rank is only valid with type top10; remove it or change the type"
                .into(),
        ));
    }
    if format.props.contains_key("stopIfTrue")
        && format
            .props
            .get("stopIfTrue")
            .and_then(Value::as_bool)
            .is_none()
    {
        return Err(Error::Compile(
            "ConditionalFormat.stopIfTrue must be boolean; use stopIfTrue={true} or {false}".into(),
        ));
    }
    if scale_type == "cellIs" {
        return validate_cell_is_conditional_format(format);
    }
    if scale_type == "formula" {
        return validate_formula_conditional_format(format);
    }
    if scale_type == "text" {
        return validate_text_conditional_format(format);
    }
    if scale_type == "dataBar" {
        return validate_data_bar_conditional_format(format);
    }
    if scale_type == "iconSet" {
        return validate_icon_set_conditional_format(format);
    }
    if scale_type == "top10" {
        return validate_top10_conditional_format(format);
    }
    if ["average", "date", "blank", "error", "duplicate"].contains(&scale_type) {
        return validate_simple_conditional_format(format, scale_type);
    }
    if !["colorScale2", "colorScale3"].contains(&scale_type) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.type `{scale_type}` is unsupported; use colorScale2, colorScale3, cellIs, formula, text, dataBar, iconSet, top10, average, date, blank, error, or duplicate"
        )));
    }
    validate_color_scale_conditional_format(format, declaration_index, scale_type)
}

fn validate_simple_conditional_format(format: &Node, scale_type: &str) -> Result<()> {
    const INCOMPATIBLE: &[&str] = &[
        "minColor",
        "midColor",
        "maxColor",
        "minType",
        "midType",
        "maxType",
        "minValue",
        "midValue",
        "maxValue",
        "operator",
        "formula1",
        "formula2",
        "text",
        "rank",
        "barOnly",
        "solidFill",
        "border",
        "direction",
        "axisPosition",
        "borderColor",
        "negativeFillColor",
        "negativeBorderColor",
        "axisColor",
        "iconType",
        "reverse",
        "showIconsOnly",
        "icons",
    ];
    if let Some(property) = INCOMPATIBLE
        .iter()
        .find(|property| format.props.contains_key(**property))
    {
        return Err(Error::Compile(format!(
            "ConditionalFormat type `{scale_type}` accepts rule, stopIfTrue, and differential-style properties only; remove `{property}`"
        )));
    }
    let default_rule = match scale_type {
        "average" => "aboveAverage",
        "date" => "yesterday",
        "blank" => "blank",
        "error" => "error",
        "duplicate" => "duplicate",
        _ => "",
    };
    let rule = format
        .props
        .get("rule")
        .map(|value| {
            value
                .as_str()
                .ok_or_else(|| Error::Compile("ConditionalFormat.rule must be a string".into()))
        })
        .transpose()?
        .unwrap_or(default_rule);
    let choices: &[&str] = match scale_type {
        "average" => &[
            "aboveAverage",
            "belowAverage",
            "equalOrAboveAverage",
            "equalOrBelowAverage",
            "oneStdDevAbove",
            "oneStdDevBelow",
            "twoStdDevAbove",
            "twoStdDevBelow",
            "threeStdDevAbove",
            "threeStdDevBelow",
        ],
        "date" => &[
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
        ],
        "blank" => &["blank", "nonBlank"],
        "error" => &["error", "nonError"],
        "duplicate" => &["duplicate", "unique"],
        _ => &[],
    };
    if !choices.contains(&rule) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.rule `{rule}` is invalid for type `{scale_type}`; use {}",
            choices.join(", ")
        )));
    }
    validate_conditional_differential_style(format, false)?;
    Ok(())
}

fn validate_top10_conditional_format(format: &Node) -> Result<()> {
    const INCOMPATIBLE: &[&str] = &[
        "minColor",
        "midColor",
        "maxColor",
        "minType",
        "midType",
        "maxType",
        "minValue",
        "midValue",
        "maxValue",
        "operator",
        "formula1",
        "formula2",
        "text",
        "barOnly",
        "solidFill",
        "border",
        "direction",
        "axisPosition",
        "borderColor",
        "negativeFillColor",
        "negativeBorderColor",
        "axisColor",
        "iconType",
        "reverse",
        "showIconsOnly",
        "icons",
    ];
    if let Some(property) = INCOMPATIBLE
        .iter()
        .find(|property| format.props.contains_key(**property))
    {
        return Err(Error::Compile(format!(
            "top10 accepts rule, rank, stopIfTrue, and differential-style properties only; remove `{property}`"
        )));
    }
    let rule = format
        .props
        .get("rule")
        .and_then(Value::as_str)
        .unwrap_or("top");
    if !["top", "bottom", "topPercent", "bottomPercent"].contains(&rule) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.rule `{rule}` is unsupported; use top, bottom, topPercent, or bottomPercent"
        )));
    }
    let rank = match format.props.get("rank") {
        Some(value) => value
            .as_u64()
            .ok_or_else(|| Error::Compile("ConditionalFormat.rank must be an integer".into()))?,
        None => 10,
    };
    let maximum = if rule.ends_with("Percent") { 100 } else { 1000 };
    if !(1..=maximum).contains(&rank) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.rank must be from 1 through {maximum} for rule `{rule}`"
        )));
    }
    validate_conditional_differential_style(format, false)?;
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Data-bar validation keeps its property whitelist and interdependent endpoint constraints together."
)]
fn validate_data_bar_conditional_format(format: &Node) -> Result<()> {
    let incompatible = [
        "midColor",
        "midType",
        "midValue",
        "maxColor",
        "minColor",
        "operator",
        "formula1",
        "formula2",
        "text",
        "bold",
        "italic",
        "fontColor",
        "strikethrough",
        "underline",
        "script",
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
    ];
    if incompatible
        .iter()
        .any(|property| format.props.contains_key(*property))
    {
        return Err(Error::Compile(
            "dataBar accepts endpoint, bar, direction, axis, and RGB color properties only; remove comparison, text, color-scale, and differential-style properties".into(),
        ));
    }
    if format.props.contains_key("stopIfTrue") {
        return Err(Error::Compile(
            "ConditionalFormat.stopIfTrue is not supported by rust_xlsxwriter data bars; remove it"
                .into(),
        ));
    }
    if let Some(fill_color) = format.props.get("fillColor") {
        validate_rgb_color(fill_color, "fillColor", 0, 0)?;
    }
    for property in [
        "borderColor",
        "negativeFillColor",
        "negativeBorderColor",
        "axisColor",
    ] {
        if let Some(color) = format.props.get(property) {
            validate_rgb_color(color, property, 0, 0)?;
        }
    }
    for property in ["barOnly", "solidFill", "border"] {
        if format
            .props
            .get(property)
            .is_some_and(|value| value.as_bool().is_none())
        {
            return Err(Error::Compile(format!(
                "ConditionalFormat.{property} must be boolean"
            )));
        }
    }
    if let Some(value) = format.props.get("direction") {
        let Some(direction) = value.as_str() else {
            return Err(Error::Compile(
                "ConditionalFormat.direction must be a string; use context, leftToRight, or rightToLeft".into(),
            ));
        };
        if !["context", "leftToRight", "rightToLeft"].contains(&direction) {
            return Err(Error::Compile(format!(
                "ConditionalFormat.direction `{direction}` is unsupported; use context, leftToRight, or rightToLeft"
            )));
        }
    }
    if let Some(value) = format.props.get("axisPosition") {
        let Some(position) = value.as_str() else {
            return Err(Error::Compile(
                "ConditionalFormat.axisPosition must be a string; use automatic, midpoint, or none"
                    .into(),
            ));
        };
        if !["automatic", "midpoint", "none"].contains(&position) {
            return Err(Error::Compile(format!(
                "ConditionalFormat.axisPosition `{position}` is unsupported; use automatic, midpoint, or none"
            )));
        }
    }
    for endpoint in ["min", "max"] {
        let type_property = format!("{endpoint}Type");
        let value_property = format!("{endpoint}Value");
        let Some(kind_value) = format.props.get(&type_property) else {
            if format.props.contains_key(&value_property) {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{value_property} requires {type_property}; add the endpoint type"
                )));
            }
            continue;
        };
        let Some(kind) = kind_value.as_str() else {
            return Err(Error::Compile(format!(
                "ConditionalFormat.{type_property} must be a threshold type string"
            )));
        };
        let valid = match endpoint {
            "min" => [
                "automatic",
                "min",
                "number",
                "percent",
                "percentile",
                "formula",
            ]
            .contains(&kind),
            _ => [
                "automatic",
                "max",
                "number",
                "percent",
                "percentile",
                "formula",
            ]
            .contains(&kind),
        };
        if !valid {
            return Err(Error::Compile(format!(
                "ConditionalFormat.{type_property} `{kind}` is not valid for the {endpoint} endpoint"
            )));
        }
        if kind == "automatic" || kind == endpoint {
            if format.props.contains_key(&value_property) {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{value_property} is not used with endpoint type `{kind}`; remove it"
                )));
            }
        } else {
            validate_color_scale_threshold(format, endpoint)?;
        }
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Icon-set threshold validation reports precise property paths and actionable repairs."
)]
fn validate_icon_set_conditional_format(format: &Node) -> Result<()> {
    const ICON_TYPES: &[&str] = &[
        "threeArrows",
        "threeArrowsGray",
        "threeFlags",
        "threeTrafficLights",
        "threeTrafficLightsWithRim",
        "threeSigns",
        "threeSymbolsCircled",
        "threeSymbols",
        "threeStars",
        "threeTriangles",
        "fourArrows",
        "fourArrowsGray",
        "fourRedToBlack",
        "fourHistograms",
        "fourTrafficLights",
        "fiveArrows",
        "fiveArrowsGray",
        "fiveHistograms",
        "fiveQuadrants",
        "fiveBoxes",
    ];
    const ICON_PROPERTIES: &[&str] = &[
        "type",
        "value",
        "greaterThan",
        "iconType",
        "iconIndex",
        "noIcon",
    ];
    let incompatible = [
        "minColor",
        "midColor",
        "maxColor",
        "minType",
        "midType",
        "maxType",
        "minValue",
        "midValue",
        "maxValue",
        "operator",
        "formula1",
        "formula2",
        "text",
        "bold",
        "italic",
        "fontColor",
        "strikethrough",
        "underline",
        "script",
        "numberFormat",
        "borderTop",
        "borderRight",
        "borderBottom",
        "borderLeft",
        "borderTopColor",
        "borderRightColor",
        "borderBottomColor",
        "borderLeftColor",
        "fillColor",
        "fillPattern",
        "fillForegroundColor",
        "fillBackgroundColor",
        "barOnly",
        "solidFill",
        "border",
        "direction",
        "axisPosition",
        "borderColor",
        "negativeFillColor",
        "negativeBorderColor",
        "axisColor",
    ];
    if incompatible
        .iter()
        .any(|property| format.props.contains_key(*property))
    {
        return Err(Error::Compile(
            "iconSet accepts iconType, icons, reverse, showIconsOnly, and stopIfTrue only; remove properties from other conditional-format families".into(),
        ));
    }
    let icon_type = match format.props.get("iconType") {
        Some(value) => value.as_str().ok_or_else(|| Error::Compile(
            "ConditionalFormat.iconType must be a string; choose a supported 3-, 4-, or 5-icon set".into(),
        ))?,
        None => "threeTrafficLights",
    };
    if !ICON_TYPES.contains(&icon_type) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.iconType `{icon_type}` is unsupported; choose one of the documented built-in icon sets"
        )));
    }
    let icon_count = if icon_type.starts_with("three") {
        3
    } else if icon_type.starts_with("four") {
        4
    } else {
        5
    };
    let custom_icons = format.props.get("icons").and_then(Value::as_array);
    let x14_only = ["threeStars", "threeTriangles", "fiveBoxes"].contains(&icon_type)
        || custom_icons.is_some_and(|icons| {
            icons.iter().any(|icon| {
                icon.as_object().is_some_and(|properties| {
                    properties.contains_key("iconType") || properties.contains_key("noIcon")
                })
            })
        });
    for property in ["reverse", "showIconsOnly"] {
        if let Some(value) = format.props.get(property) {
            if value.as_bool().is_none() {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{property} must be boolean"
                )));
            }
            if x14_only && value.as_bool() == Some(true) {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{property} cannot be true with iconType `{icon_type}` because the pinned writer omits it for x14-only icon sets; choose a classic icon set or remove the flag"
                )));
            }
        }
    }
    if x14_only && format.props.get("stopIfTrue").and_then(Value::as_bool) == Some(true) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.stopIfTrue cannot be true with iconType `{icon_type}` because the pinned writer omits it for x14-only icon sets; choose a classic icon set or remove the flag"
        )));
    }
    let Some(icons) = format.props.get("icons") else {
        return Ok(());
    };
    let Some(icons) = icons.as_array() else {
        return Err(Error::Compile(
            "ConditionalFormat.icons must be an array of icon threshold objects".into(),
        ));
    };
    if icons.len() != icon_count {
        return Err(Error::Compile(format!(
            "ConditionalFormat.icons has {} entries but iconType `{icon_type}` requires {icon_count}; match the number of thresholds to the icon count",
            icons.len()
        )));
    }
    for (index, icon) in icons.iter().enumerate() {
        let Some(properties) = icon.as_object() else {
            return Err(Error::Compile(format!(
                "ConditionalFormat.icons[{index}] must be an object"
            )));
        };
        if let Some(name) = properties
            .keys()
            .find(|name| !ICON_PROPERTIES.contains(&name.as_str()))
        {
            return Err(Error::Compile(format!(
                "ConditionalFormat.icons[{index}] has unsupported property `{name}`; use type, value, greaterThan, iconType/iconIndex, or noIcon"
            )));
        }
        let kind = properties.get("type").and_then(Value::as_str).ok_or_else(|| Error::Compile(format!("ConditionalFormat.icons[{index}].type is required; use number, percent, percentile, or formula")))?;
        if !["number", "percent", "percentile", "formula"].contains(&kind) {
            return Err(Error::Compile(format!(
                "ConditionalFormat.icons[{index}].type `{kind}` is unsupported; use number, percent, percentile, or formula"
            )));
        }
        let value = properties.get("value").ok_or_else(|| {
            Error::Compile(format!(
                "ConditionalFormat.icons[{index}].value is required"
            ))
        })?;
        if kind == "formula" {
            let Some(formula) = value.as_str().filter(|formula| {
                !formula.is_empty()
                    && formula.chars().count() <= 8192
                    && formula.chars().all(is_valid_xml_character)
            }) else {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].value must be a non-empty XML-valid formula string of at most 8,192 characters"
                )));
            };
            let _ = formula;
        } else {
            let Some(number) = value.as_f64().filter(|number| number.is_finite()) else {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].value must be a finite number for `{kind}` thresholds"
                )));
            };
            if ["percent", "percentile"].contains(&kind) && !(0.0..=100.0).contains(&number) {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].value must be between 0 and 100 for `{kind}` thresholds"
                )));
            }
            if index == 0 && (kind != "percent" || number != 0.0) {
                return Err(Error::Compile("ConditionalFormat.icons[0] is fixed by rust_xlsxwriter to percent 0; specify type=\"percent\" and value={0}".into()));
            }
        }
        if let Some(value) = properties.get("greaterThan") {
            if value.as_bool().is_none() {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].greaterThan must be boolean"
                )));
            }
            if index == 0 && value.as_bool() == Some(true) {
                return Err(Error::Compile("ConditionalFormat.icons[0].greaterThan cannot be true because the writer forces the first threshold to percent 0 with the default inclusive boundary".into()));
            }
            if x14_only && value.as_bool() == Some(true) {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].greaterThan cannot be true when iconType `{icon_type}` or custom icon choices require an x14-only rule; the pinned writer does not serialize greaterThan in x14 cfvo elements, so remove the flag or use a classic icon set without custom icons"
                )));
            }
        }
        let custom_icon_type = properties.get("iconType");
        let icon_index = properties.get("iconIndex");
        if custom_icon_type.is_some() != icon_index.is_some() {
            return Err(Error::Compile(format!(
                "ConditionalFormat.icons[{index}] must provide iconType and iconIndex together"
            )));
        }
        if let Some(custom_type) = custom_icon_type {
            let Some(custom_type) = custom_type.as_str() else {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].iconType must name a built-in icon set"
                )));
            };
            if !ICON_TYPES.contains(&custom_type) {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].iconType `{custom_type}` is unsupported"
                )));
            }
            let custom_count = if custom_type.starts_with("three") {
                3
            } else if custom_type.starts_with("four") {
                4
            } else {
                5
            };
            if properties
                .get("iconIndex")
                .and_then(Value::as_u64)
                .is_none_or(|icon_index| icon_index >= custom_count)
            {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].iconIndex is outside the 0..{} range for `{custom_type}`",
                    custom_count - 1
                )));
            }
        }
        if let Some(no_icon) = properties.get("noIcon") {
            if no_icon.as_bool() != Some(true) {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].noIcon can only be true when present; omit it to show an icon"
                )));
            }
            if custom_icon_type.is_some() || icon_index.is_some() {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.icons[{index}].noIcon cannot be combined with iconType or iconIndex"
                )));
            }
        }
    }
    Ok(())
}

fn validate_text_conditional_format(format: &Node) -> Result<()> {
    if [
        "formula1", "formula2", "minColor", "midColor", "maxColor", "minType", "midType",
        "maxType", "minValue", "midValue", "maxValue",
    ]
    .iter()
    .any(|property| format.props.contains_key(*property))
    {
        return Err(Error::Compile(
            "text conditional formats accept operator, text, and differential-style properties only; remove formula and color-scale properties".into(),
        ));
    }
    let operator = format
        .props
        .get("operator")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(
                "ConditionalFormat.operator is required for type text; choose containsText, notContains, beginsWith, or endsWith".into(),
            )
        })?;
    if !["containsText", "notContains", "beginsWith", "endsWith"].contains(&operator) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.operator `{operator}` is unsupported for type text; use containsText, notContains, beginsWith, or endsWith"
        )));
    }
    let text = format
        .props
        .get("text")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(
                "ConditionalFormat.text is required for type text; provide the matching text"
                    .into(),
            )
        })?;
    if text.chars().count() > 255 || !text.chars().all(is_valid_xml_character) {
        return Err(Error::Compile(
            "ConditionalFormat.text must contain at most 255 XML-valid characters; shorten it or remove invalid characters".into(),
        ));
    }
    if text.contains('"') {
        return Err(Error::Compile(
            "ConditionalFormat.text cannot contain a double quote because the pinned writer inserts it unescaped into an Excel formula; remove the quote or use type=\"formula\" with an explicitly escaped formula".into(),
        ));
    }
    validate_conditional_differential_style(format, true)
}

fn validate_formula_conditional_format(format: &Node) -> Result<()> {
    if [
        "operator", "formula2", "minColor", "midColor", "maxColor", "minType", "midType",
        "maxType", "minValue", "midValue", "maxValue",
    ]
    .iter()
    .any(|property| format.props.contains_key(*property))
    {
        return Err(Error::Compile(
            "formula conditional formats accept formula1 and differential-style properties only; remove operator, formula2, and color-scale properties".into(),
        ));
    }
    let has_valid_formula = format
        .props
        .get("formula1")
        .and_then(Value::as_str)
        .is_some_and(|value| {
            !value.is_empty()
                && value.chars().count() <= 8192
                && value.chars().all(is_valid_xml_character)
        });
    if !has_valid_formula {
        return Err(Error::Compile(
            "ConditionalFormat.formula1 must contain 1 to 8,192 XML-valid characters; provide the boolean Excel formula".into(),
        ));
    }
    validate_conditional_differential_style(format, true)
}

fn validate_conditional_differential_style(format: &Node, require_style: bool) -> Result<()> {
    let mut has_style = false;
    for property in ["bold", "italic", "strikethrough"] {
        if let Some(value) = format.props.get(property) {
            if value.as_bool() != Some(true) {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{property} must be true when present; remove it or enable the differential style"
                )));
            }
            has_style = true;
        }
    }
    if let Some(value) = format.props.get("underline") {
        if !value.as_str().is_some_and(|underline| {
            ["single", "double", "singleAccounting", "doubleAccounting"].contains(&underline)
        }) {
            return Err(Error::Compile("ConditionalFormat.underline must be single, double, singleAccounting, or doubleAccounting; choose a supported OOXML underline value".into()));
        }
        has_style = true;
    }
    let has_fill_style = validate_conditional_fill_properties(format)?;
    has_style |= has_fill_style;
    if let Some(value) = format.props.get("script") {
        let Some(script) = value.as_str() else {
            return Err(Error::Compile("ConditionalFormat.script must be superscript or subscript; choose a supported OOXML vertAlign value".into()));
        };
        if !["superscript", "subscript"].contains(&script) {
            return Err(Error::Compile("ConditionalFormat.script must be superscript or subscript; choose a supported OOXML vertAlign value".into()));
        }
        if !has_style && !format.props.contains_key("fontColor") && !has_fill_style {
            return Err(Error::Compile("ConditionalFormat.script requires another differential-style property because the pinned writer omits script-only dxf styles; add bold, fontColor, or fillColor".into()));
        }
        has_style = true;
    }
    if let Some(value) = format.props.get("numberFormat") {
        validate_number_format(value, "ConditionalFormat")?;
        has_style = true;
    }
    if let Some(color) = format.props.get("fontColor") {
        validate_rgb_color(color, "fontColor", 0, 0)?;
        has_style = true;
    }
    for suffix in ["Top", "Right", "Bottom", "Left"] {
        let style_key = format!("border{suffix}");
        let color_key = format!("border{suffix}Color");
        if let Some(value) = format.props.get(&style_key) {
            let Some(style) = value.as_str() else {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{style_key} must be a quoted OOXML border style"
                )));
            };
            if !is_supported_conditional_border_style(style) {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{style_key} `{style}` is unsupported; use a supported edge border style"
                )));
            }
            has_style = true;
        }
        if let Some(value) = format.props.get(&color_key) {
            validate_rgb_color(value, &color_key, 0, 0)?;
            if format
                .props
                .get(&style_key)
                .and_then(Value::as_str)
                .is_none_or(|style| style == "none")
            {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{color_key} requires a matching non-`none` {style_key}; add the edge style or remove the color"
                )));
            }
            has_style = true;
        }
    }
    if require_style && !has_style {
        return Err(Error::Compile(
            "ConditionalFormat requires at least one differential style; add a font style, numberFormat, or edge border".into(),
        ));
    }
    Ok(())
}

fn validate_conditional_fill_properties(format: &Node) -> Result<bool> {
    for property in ["fillColor", "fillForegroundColor", "fillBackgroundColor"] {
        if let Some(color) = format.props.get(property) {
            validate_rgb_color(color, property, 0, 0)?;
        }
    }
    if let Some(value) = format.props.get("fillPattern")
        && !value
            .as_str()
            .is_some_and(|pattern| FILL_PATTERN_NAMES.contains(&pattern))
    {
        return Err(Error::Compile(format!(
            "ConditionalFormat.fillPattern must be one of {}; use a supported OOXML pattern name",
            FILL_PATTERN_NAMES.join(", ")
        )));
    }
    let has_pattern = format.props.contains_key("fillPattern");
    let has_fill_color = format.props.contains_key("fillColor");
    let has_pattern_colors = format.props.contains_key("fillForegroundColor")
        || format.props.contains_key("fillBackgroundColor");
    if ((has_pattern || has_pattern_colors) && (!has_pattern || has_fill_color))
        || (format.props.get("fillPattern").and_then(Value::as_str) == Some("none")
            && has_pattern_colors)
    {
        return Err(Error::Compile(
            "ConditionalFormat fill properties conflict: fillColor is the solid-fill shorthand; foreground/background colors require a non-none fillPattern. Choose one fill representation".into(),
        ));
    }
    Ok(has_pattern || has_fill_color || has_pattern_colors)
}

fn validate_number_format(value: &Value, context: &str) -> Result<()> {
    let Some(number_format) = value.as_str() else {
        return Err(Error::Compile(format!(
            "{context}.numberFormat must be a string; use an Excel format code such as `#,##0.00`"
        )));
    };
    if number_format.is_empty() || number_format.chars().count() > 255 {
        return Err(Error::Compile(format!(
            "{context}.numberFormat must contain 1 to 255 characters; shorten or provide an Excel format code"
        )));
    }
    if !number_format.chars().all(is_valid_xml_character) {
        return Err(Error::Compile(format!(
            "{context}.numberFormat contains a character that is invalid in XML; remove control characters and use a valid Excel format code"
        )));
    }
    Ok(())
}

fn is_supported_conditional_border_style(style: &str) -> bool {
    [
        "thin",
        "medium",
        "dashed",
        "dotted",
        "thick",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "mediumDashDotDot",
        "slantDashDot",
    ]
    .contains(&style)
}

fn validate_color_scale_conditional_format(
    format: &Node,
    declaration_index: usize,
    scale_type: &str,
) -> Result<()> {
    if [
        "operator",
        "formula1",
        "formula2",
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
    .any(|property| format.props.contains_key(*property))
    {
        return Err(Error::Compile(
            "colorScale conditional formats do not accept cellIs operator, formula, or differential-style properties; remove those properties".into(),
        ));
    }
    for property in ["minColor", "maxColor"] {
        let color = format.props.get(property).ok_or_else(|| {
            Error::Compile(format!(
                "ConditionalFormat.{property} is required; provide an opaque RGB color such as `#63BE7B`"
            ))
        })?;
        validate_rgb_color(color, property, declaration_index, 0)?;
    }
    if scale_type == "colorScale3" {
        let color = format.props.get("midColor").ok_or_else(|| {
            Error::Compile(
                "ConditionalFormat.midColor is required for colorScale3; provide an opaque RGB color such as `#FFEB84`".into(),
            )
        })?;
        validate_rgb_color(color, "midColor", declaration_index, 0)?;
    } else if format.props.contains_key("midColor")
        || format.props.contains_key("midType")
        || format.props.contains_key("midValue")
    {
        return Err(Error::Compile(
            "colorScale2 does not accept midpoint properties; remove midColor, midType, and midValue".into(),
        ));
    }
    for threshold in ["min", "mid", "max"] {
        if threshold == "mid" && scale_type != "colorScale3" {
            continue;
        }
        validate_color_scale_threshold(format, threshold)?;
    }
    Ok(())
}

fn validate_cell_is_conditional_format(format: &Node) -> Result<()> {
    if [
        "minColor", "midColor", "maxColor", "minType", "midType", "maxType", "minValue",
        "midValue", "maxValue",
    ]
    .iter()
    .any(|property| format.props.contains_key(*property))
    {
        return Err(Error::Compile(
            "cellIs conditional formats do not accept color-scale threshold or color properties; remove them".into(),
        ));
    }
    let operator = format
        .props
        .get("operator")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            Error::Compile(
                "ConditionalFormat.operator is required for cellIs; select a comparison operator"
                    .into(),
            )
        })?;
    let allowed = [
        "equal",
        "notEqual",
        "greaterThan",
        "greaterThanOrEqual",
        "lessThan",
        "lessThanOrEqual",
        "between",
        "notBetween",
    ];
    if !allowed.contains(&operator) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.operator `{operator}` is unsupported; use {}",
            allowed.join(", ")
        )));
    }
    for property in ["formula1", "formula2"] {
        if let Some(value) = format.props.get(property) {
            let Some(value) = value.as_str() else {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{property} must be a string containing exact Excel formula text"
                )));
            };
            if value.is_empty()
                || value.chars().count() > 8192
                || !value.chars().all(is_valid_xml_character)
            {
                return Err(Error::Compile(format!(
                    "ConditionalFormat.{property} must contain 1 to 8,192 XML-valid characters; repair or shorten the formula"
                )));
            }
        }
    }
    if !format.props.contains_key("formula1") {
        return Err(Error::Compile(
            "ConditionalFormat.formula1 is required for cellIs; provide the comparison value or formula".into(),
        ));
    }
    let two_bound = ["between", "notBetween"].contains(&operator);
    if two_bound != format.props.contains_key("formula2") {
        return Err(Error::Compile(if two_bound {
            "ConditionalFormat.formula2 is required for between/notBetween; provide the upper bound"
                .into()
        } else {
            "ConditionalFormat.formula2 is only valid for between/notBetween; remove it or change operator".into()
        }));
    }
    validate_conditional_differential_style(format, true)
}

fn validate_color_scale_threshold(format: &Node, threshold: &str) -> Result<()> {
    let type_property = format!("{threshold}Type");
    let value_property = format!("{threshold}Value");
    let Some(type_value) = format.props.get(&type_property) else {
        if format.props.contains_key(&value_property) {
            return Err(Error::Compile(format!(
                "ConditionalFormat.{value_property} requires {type_property}; add its threshold type"
            )));
        }
        return Ok(());
    };
    let Some(type_value) = type_value.as_str() else {
        return Err(Error::Compile(format!(
            "ConditionalFormat.{type_property} must be a threshold type; use min, max, number, percent, percentile, or formula"
        )));
    };
    let allowed = ["min", "max", "number", "percent", "percentile", "formula"];
    if !allowed.contains(&type_value) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.{type_property} `{type_value}` is unsupported; use {}",
            allowed.join(", ")
        )));
    }
    if (threshold == "min" && type_value == "max")
        || (threshold == "max" && type_value == "min")
        || (threshold == "mid" && ["min", "max"].contains(&type_value))
    {
        return Err(Error::Compile(format!(
            "ConditionalFormat.{type_property} `{type_value}` cannot be used for the {threshold} stop; choose a matching endpoint or a numeric threshold"
        )));
    }
    if ["min", "max"].contains(&type_value) {
        if format.props.contains_key(&value_property) {
            return Err(Error::Compile(format!(
                "ConditionalFormat.{value_property} is not used with threshold type `{type_value}`; remove it"
            )));
        }
        return Ok(());
    }
    let value = format.props.get(&value_property).ok_or_else(|| {
        Error::Compile(format!(
            "ConditionalFormat.{value_property} is required when {type_property} is `{type_value}`"
        ))
    })?;
    if type_value == "formula" {
        let Some(value) = value.as_str().filter(|value| !value.trim().is_empty()) else {
            return Err(Error::Compile(format!(
                "ConditionalFormat.{value_property} must be a non-empty formula string when {type_property} is `formula`"
            )));
        };
        if value.chars().count() > 8192 || !value.chars().all(is_valid_xml_character) {
            return Err(Error::Compile(format!(
                "ConditionalFormat.{value_property} must be at most 8,192 XML-valid formula characters"
            )));
        }
        return Ok(());
    }
    let Some(value) = value.as_f64().filter(|value| value.is_finite()) else {
        return Err(Error::Compile(format!(
            "ConditionalFormat.{value_property} must be a finite number when {type_property} is `{type_value}`"
        )));
    };
    if ["percent", "percentile"].contains(&type_value) && !(0.0..=100.0).contains(&value) {
        return Err(Error::Compile(format!(
            "ConditionalFormat.{value_property} must be in the inclusive range 0–100 for `{type_value}`"
        )));
    }
    Ok(())
}

fn validate_data_validation_text(
    property: &str,
    value: &str,
    declaration_index: usize,
    maximum_length: usize,
) -> Result<()> {
    if value.chars().count() > maximum_length || !value.chars().all(is_valid_xml_character) {
        return Err(Error::Compile(format!(
            "DataValidation.{property} at declaration {} must contain at most {maximum_length} valid XML characters; shorten it or remove control characters",
            declaration_index + 1
        )));
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "Row cell validation keeps positional spans, merge constraints, and the per-cell property whitelist in one ordered pass."
)]
fn validate_row_cells(
    row: &Node,
    current_row_index: usize,
    merged_ranges: &mut Vec<(usize, usize, usize, usize)>,
) -> Result<()> {
    let mut column_index = 0_usize;
    for (cell_index, cell_value) in row.children.iter().enumerate() {
        while merged_ranges
            .iter()
            .any(|(first_row, first_column, last_row, last_column)| {
                *first_row <= current_row_index
                    && current_row_index <= *last_row
                    && *first_column <= column_index
                    && column_index <= *last_column
            })
        {
            column_index += 1;
        }
        let cell: Node = serde_json::from_value(cell_value.clone()).map_err(|error| {
            Error::Compile(format!(
                "Row {current_row_index} cell {cell_index} must be Cell: {error}"
            ))
        })?;
        if cell.r#type != "Cell" {
            return Err(Error::Compile(format!(
                "Row {current_row_index} child {cell_index} is `{}`; expected Cell",
                cell.r#type
            )));
        }
        let mut allowed_cell_props = vec![
            "value",
            "formula",
            "formulaResult",
            "arrayFormulaRange",
            "dynamicArray",
            "date",
            "hyperlink",
            "hyperlinkTooltip",
            "rowSpan",
            "columnSpan",
            "checkbox",
        ];
        allowed_cell_props.extend_from_slice(FORMAT_PROPERTIES);
        reject_unknown_props(&cell, &allowed_cell_props, "Cell")?;
        let row_span = validate_cell_span(&cell, "rowSpan", 1, current_row_index, column_index)?;
        let column_span =
            validate_cell_span(&cell, "columnSpan", 1, current_row_index, column_index)?;
        let array_formula_range = cell
            .props
            .get("arrayFormulaRange")
            .and_then(Value::as_str)
            .and_then(crate::parse_a1_range);
        if let Some(range) = array_formula_range {
            if (range.0 as usize, usize::from(range.1)) != (current_row_index, column_index) {
                return Err(Error::Compile(format!(
                    "Cell.arrayFormulaRange at row {}, column {} must start at this Cell's top-left coordinate; move the formula anchor or adjust the range",
                    current_row_index + 1,
                    column_index + 1
                )));
            }
            if merged_ranges
                .iter()
                .any(|(first_row, first_column, last_row, last_column)| {
                    range.0 as usize <= *last_row
                        && range.2 as usize >= *first_row
                        && usize::from(range.1) <= *last_column
                        && usize::from(range.3) >= *first_column
                })
            {
                return Err(Error::Compile(format!(
                    "Cell.arrayFormulaRange `{}` overlaps a prior array formula or merged range; make formula and merged ranges disjoint",
                    crate::format_a1_range(range)
                )));
            }
            merged_ranges.push((
                range.0 as usize,
                usize::from(range.1),
                range.2 as usize,
                usize::from(range.3),
            ));
        }
        let last_row = current_row_index.checked_add(row_span - 1);
        let last_column = column_index.checked_add(column_span - 1);
        if last_row.is_none_or(|last| last >= 1_048_576)
            || last_column.is_none_or(|last| last >= 16_384)
        {
            return Err(Error::Compile(format!(
                "Cell span at row {}, column {} exceeds Excel's worksheet limits; reduce rowSpan/columnSpan or move the cell earlier",
                current_row_index + 1,
                column_index + 1
            )));
        }
        if row_span > 1 || column_span > 1 {
            let (last_row, last_column) = (
                last_row.unwrap_or_default(),
                last_column.unwrap_or_default(),
            );
            if merged_ranges
                .iter()
                .any(|(first_row, first_column, end_row, end_column)| {
                    current_row_index <= *end_row
                        && last_row >= *first_row
                        && column_index <= *end_column
                        && last_column >= *first_column
                })
            {
                return Err(Error::Compile(format!(
                    "Cell span at row {}, column {} overlaps another merged range; adjust the span or place the cell outside that range",
                    current_row_index + 1,
                    column_index + 1
                )));
            }
            merged_ranges.push((current_row_index, column_index, last_row, last_column));
        }
        validate_cell_style(&cell, current_row_index, column_index)?;
        validate_cell(&cell, current_row_index, column_index)?;
        column_index += column_span;
    }
    Ok(())
}

fn validate_cell_span(
    cell: &Node,
    property: &str,
    default: usize,
    row: usize,
    column: usize,
) -> Result<usize> {
    let Some(value) = cell.props.get(property) else {
        return Ok(default);
    };
    let Some(value) = value.as_u64().and_then(|value| usize::try_from(value).ok()) else {
        return Err(Error::Compile(format!(
            "Cell.{property} at row {}, column {} must be a positive integer; use a span count such as `{property}={{2}}`",
            row + 1,
            column + 1
        )));
    };
    if value == 0 {
        return Err(Error::Compile(format!(
            "Cell.{property} at row {}, column {} must be at least 1; use `{property}={{1}}` for an unmerged axis or a larger count to merge",
            row + 1,
            column + 1
        )));
    }
    Ok(value)
}

fn validate_column(column: &Node, column_index: usize) -> Result<()> {
    if column.r#type != "Column" {
        return Err(Error::Compile(format!(
            "Worksheet child {} is `{}`; expected Column or Row",
            column_index + 1,
            column.r#type
        )));
    }
    let mut allowed_column_props = vec![
        "min",
        "max",
        "widthPixels",
        "hidden",
        "outlineLevel",
        "collapsed",
    ];
    allowed_column_props.extend_from_slice(FORMAT_PROPERTIES);
    reject_unknown_props(column, &allowed_column_props, "Column")?;
    for property in ["min", "max"] {
        let Some(value) = column.props.get(property).and_then(Value::as_u64) else {
            return Err(Error::Compile(format!(
                "Column.{property} at declaration {} must be an integer from 1 to 16384; use 1-based Excel column numbers",
                column_index + 1
            )));
        };
        if !(1..=16_384).contains(&value) {
            return Err(Error::Compile(format!(
                "Column.{property} at declaration {} must be from 1 to 16384; use a valid Excel column number",
                column_index + 1
            )));
        }
    }
    let min = column.props["min"].as_u64().unwrap_or_default();
    let max = column.props["max"].as_u64().unwrap_or_default();
    if min > max {
        return Err(Error::Compile(format!(
            "Column.min must not exceed Column.max at declaration {}; reorder the range endpoints",
            column_index + 1
        )));
    }
    if let Some(width) = column.props.get("widthPixels") {
        let Some(width) = width.as_u64() else {
            return Err(Error::Compile(format!(
                "Column.widthPixels at declaration {} must be an integer from 1 to 1790; use pixel units",
                column_index + 1
            )));
        };
        if !(1..=1790).contains(&width) {
            return Err(Error::Compile(format!(
                "Column.widthPixels at declaration {} must be from 1 to 1790; use a representable pixel width",
                column_index + 1
            )));
        }
    }
    if let Some(hidden) = column.props.get("hidden")
        && hidden.as_bool().is_none()
    {
        return Err(Error::Compile(format!(
            "Column.hidden must be a boolean at declaration {}; use `hidden={{true}}` or `hidden={{false}}`",
            column_index + 1
        )));
    }
    if let Some(collapsed) = column.props.get("collapsed") {
        let Some(collapsed) = collapsed.as_bool() else {
            return Err(Error::Compile(format!(
                "Column.collapsed must be a boolean at declaration {}; use `collapsed={{true}}` on a single summary column",
                column_index + 1
            )));
        };
        if collapsed && min != max {
            return Err(Error::Compile(format!(
                "Column.collapsed at declaration {} requires min=max because each collapsed marker belongs to one summary column; split the declaration into single-column ranges",
                column_index + 1
            )));
        }
    }
    validate_outline_level(
        &column.props,
        "Column",
        &format!("declaration {}", column_index + 1),
    )?;
    validate_format_component(
        column,
        "Column",
        column_index,
        0,
        &format!("declaration {}", column_index + 1),
    )?;
    if !column.children.is_empty() {
        return Err(Error::Compile(format!(
            "Column at declaration {} cannot have children; move cell data into Row components",
            column_index + 1
        )));
    }
    Ok(())
}

fn reject_unknown_props(node: &Node, allowed: &[&str], component: &str) -> Result<()> {
    if let Some(property) = node
        .props
        .keys()
        .find(|property| !allowed.contains(&property.as_str()))
    {
        return Err(Error::Compile(format!(
            "{component} has unsupported property `{property}`; remove it or check `xlsx-jsx spec` for supported properties"
        )));
    }
    Ok(())
}

fn validate_row_properties(row: &Node, row_index: usize) -> Result<()> {
    if let Some(height) = row.props.get("height") {
        let Some(height) = height.as_f64() else {
            return Err(Error::Compile(format!(
                "Row.height at row {} must be a finite number in 0.75-point steps; use a value such as `height={{22.5}}`",
                row_index + 1
            )));
        };
        let pixels = height * 4.0 / 3.0;
        if !height.is_finite()
            || !(0.75..=409.5).contains(&height)
            || (pixels - pixels.round()).abs() > 1e-9
        {
            return Err(Error::Compile(format!(
                "Row.height at row {} must be between 0.75 and 409.5 points in 0.75-point steps; choose a representable height such as `height={{22.5}}`",
                row_index + 1
            )));
        }
    }
    if let Some(hidden) = row.props.get("hidden")
        && hidden.as_bool().is_none()
    {
        return Err(Error::Compile(format!(
            "Row.hidden must be a boolean at row {}; use `hidden={{true}}` or `hidden={{false}}`",
            row_index + 1
        )));
    }
    if let Some(collapsed) = row.props.get("collapsed")
        && collapsed.as_bool().is_none()
    {
        return Err(Error::Compile(format!(
            "Row.collapsed must be a boolean at row {}; use `collapsed={{true}}` for a summary row",
            row_index + 1
        )));
    }
    let mut allowed_row_props = vec!["height", "hidden", "collapsed", "outlineLevel"];
    allowed_row_props.extend_from_slice(FORMAT_PROPERTIES);
    reject_unknown_props(row, &allowed_row_props, "Row")?;
    validate_format_component(row, "Row", row_index, 0, &format!("row {}", row_index + 1))?;
    validate_outline_level(&row.props, "Row", &format!("row {}", row_index + 1))?;
    Ok(())
}

fn validate_format_component(
    node: &Node,
    component: &str,
    row: usize,
    column: usize,
    location: &str,
) -> Result<()> {
    let cell = Node {
        r#type: "Cell".into(),
        props: node.props.clone(),
        children: Vec::new(),
    };
    validate_cell_style(&cell, row, column).map_err(|error| match error {
        Error::Compile(message) => {
            Error::Compile(message.replace("Cell.", &format!("{component}.")).replace(
                &format!("at row {}, column {}", row + 1, column + 1),
                &format!("at {location}"),
            ))
        }
        other => other,
    })
}

fn validate_outline_level(
    props: &serde_json::Map<String, Value>,
    component: &str,
    location: &str,
) -> Result<()> {
    let Some(value) = props.get("outlineLevel") else {
        return Ok(());
    };
    let Some(level) = value.as_u64() else {
        return Err(Error::Compile(format!(
            "{component}.outlineLevel at {location} must be an integer from 1 through 7; use an Excel outline level such as `outlineLevel={{1}}`"
        )));
    };
    if !(1..=7).contains(&level) {
        return Err(Error::Compile(format!(
            "{component}.outlineLevel at {location} must be from 1 through 7; use a supported Excel nesting level"
        )));
    }
    Ok(())
}

fn validate_empty_row_placement(
    sheet: &Node,
    row: &Node,
    row_index: usize,
    sheet_name: &str,
    merged_ranges: &[(usize, usize, usize, usize)],
) -> Result<()> {
    let row_has_metadata = row.props.contains_key("height")
        || row.props.get("hidden").and_then(Value::as_bool) == Some(true)
        || row.props.contains_key("outlineLevel")
        || row.props.get("collapsed").and_then(Value::as_bool) == Some(true);
    let has_later_observable_row = sheet
        .children
        .iter()
        .filter_map(|value| serde_json::from_value::<Node>(value.clone()).ok())
        .filter(|later| later.r#type == "Row")
        .skip(row_index + 1)
        .any(|later| {
            !later.children.is_empty()
                || later.props.contains_key("height")
                || later.props.get("hidden").and_then(Value::as_bool) == Some(true)
                || later.props.contains_key("outlineLevel")
                || later.props.get("collapsed").and_then(Value::as_bool) == Some(true)
        });
    let row_is_covered_by_merge = merged_ranges
        .iter()
        .any(|(first_row, _, last_row, _)| *first_row <= row_index && row_index <= *last_row);
    if row.children.is_empty()
        && !row_has_metadata
        && !has_later_observable_row
        && !row_is_covered_by_merge
    {
        return Err(Error::Compile(format!(
            "trailing empty Row {} in Worksheet `{sheet_name}` has no XLSX representation; add a Cell or row metadata, or remove the row",
            row_index + 1
        )));
    }
    Ok(())
}

fn validate_dynamic_array_properties(
    cell: &Node,
    formula: Option<&Value>,
    value: Option<&Value>,
    date: Option<&Value>,
    row: usize,
    column: usize,
) -> Result<()> {
    let Some(dynamic_array) = cell.props.get("dynamicArray") else {
        return Ok(());
    };
    if !dynamic_array.is_boolean() {
        return Err(Error::Compile(format!(
            "Cell.dynamicArray at row {}, column {} must be a boolean; use dynamicArray={{true}} for Excel 365 dynamic formulas",
            row + 1,
            column + 1
        )));
    }
    if dynamic_array.as_bool() == Some(true)
        && (formula.is_none()
            || value.is_some()
            || date.is_some()
            || !cell.children.is_empty()
            || ["hyperlink", "formulaResult", "rowSpan", "columnSpan"]
                .iter()
                .any(|property| cell.props.contains_key(*property)))
    {
        return Err(Error::Compile(format!(
            "Cell.dynamicArray at row {}, column {} requires formula and cannot combine with value/date/hyperlink/formulaResult/children/merged-cell spans; add a formula or remove conflicting properties",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_cell(cell: &Node, row: usize, column: usize) -> Result<()> {
    if cell.children.iter().any(|child| {
        child
            .get("type")
            .and_then(Value::as_str)
            .is_some_and(|kind| kind == "TextRun")
    }) {
        return validate_rich_text_cell(cell, row, column);
    }
    let formula = cell.props.get("formula");
    let array_formula_range = cell.props.get("arrayFormulaRange");
    let value = cell.props.get("value");
    let date = cell.props.get("date");
    validate_checkbox_cell(cell, row, column)?;
    validate_dynamic_array_properties(cell, formula, value, date, row, column)?;
    if formula.is_some() && (value.is_some() || date.is_some() || !cell.children.is_empty()) {
        return Err(Error::Compile(format!(
            "Cell at row {}, column {} cannot combine formula with value/date/children; use one representation",
            row + 1,
            column + 1
        )));
    }
    validate_array_formula_cell(cell, row, column, formula, array_formula_range)?;
    validate_formula_result(cell, row, column, formula, array_formula_range)?;
    if date.is_some() && (value.is_some() || !cell.children.is_empty()) {
        return Err(Error::Compile(format!(
            "Cell at row {}, column {} cannot combine date with value/children; use one representation",
            row + 1,
            column + 1
        )));
    }
    if value.is_some() && !cell.children.is_empty() {
        return Err(Error::Compile(format!(
            "Cell at row {}, column {} cannot combine the value property with children; use one representation",
            row + 1,
            column + 1
        )));
    }
    if let Some(formula) = formula
        && formula.as_str().is_none_or(|text| text.trim().is_empty())
    {
        return Err(Error::Compile(format!(
            "Cell formula at row {}, column {} must be a non-empty string; provide an Excel formula such as `=SUM(A1:A3)`",
            row + 1,
            column + 1
        )));
    }
    if let Some(date) = date {
        let valid = date.as_str().is_some_and(is_supported_iso_date);
        if !valid {
            return Err(Error::Compile(format!(
                "Cell date at row {}, column {} must be a valid local ISO date/time (`YYYY-MM-DD` or `YYYY-MM-DDTHH:mm:ss[.fff]`); remove the time-zone suffix because Excel serial dates do not retain a time zone",
                row + 1,
                column + 1
            )));
        }
    }
    validate_cell_hyperlink(cell, row, column)?;
    let scalars = cell
        .children
        .iter()
        .filter(|value| !value.is_null())
        .count();
    if scalars > 1 {
        return Err(Error::Compile(format!(
            "Cell at row {}, column {} has multiple values; provide one scalar or use the value property",
            row + 1,
            column + 1
        )));
    }
    let value = value.or_else(|| cell.children.first());
    if value.is_some_and(|value| {
        !(value.is_null()
            || value.is_string()
            || value.is_boolean()
            || value.as_f64().is_some_and(f64::is_finite))
    }) {
        return Err(Error::Compile(format!(
            "Cell at row {}, column {} must contain a string, finite number, boolean, or null; convert dates to ISO text in v1",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_checkbox_cell(cell: &Node, row: usize, column: usize) -> Result<()> {
    let Some(value) = cell.props.get("checkbox") else {
        return Ok(());
    };
    let Some(enabled) = value.as_bool() else {
        return Err(Error::Compile(format!(
            "Cell.checkbox at row {}, column {} must be a boolean; use checkbox={{true}} for a checkbox cell",
            row + 1,
            column + 1
        )));
    };
    if !enabled {
        return Ok(());
    }
    let cell_value = cell.props.get("value").or_else(|| cell.children.first());
    if !cell_value.is_some_and(Value::is_boolean)
        || cell.props.contains_key("formula")
        || cell.props.contains_key("date")
        || cell.props.contains_key("hyperlink")
        || cell
            .props
            .get("rowSpan")
            .and_then(Value::as_u64)
            .unwrap_or(1)
            != 1
        || cell
            .props
            .get("columnSpan")
            .and_then(Value::as_u64)
            .unwrap_or(1)
            != 1
        || cell.children.iter().any(|child| {
            child
                .get("type")
                .and_then(Value::as_str)
                .is_some_and(|kind| kind == "TextRun")
        })
    {
        return Err(Error::Compile(format!(
            "Cell.checkbox at row {}, column {} requires a scalar boolean value and cannot be combined with formula, date, hyperlink, rich text, or merged spans; use value={{true}} or value={{false}} on a plain Cell",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_formula_result(
    cell: &Node,
    row: usize,
    column: usize,
    formula: Option<&Value>,
    array_formula_range: Option<&Value>,
) -> Result<()> {
    let Some(result) = cell.props.get("formulaResult") else {
        return Ok(());
    };
    if formula.is_none() {
        return Err(Error::Compile(format!(
            "Cell.formulaResult at row {}, column {} requires formula; add Cell.formula or remove formulaResult",
            row + 1,
            column + 1
        )));
    }
    if array_formula_range.is_some() {
        return Err(Error::Compile(format!(
            "Cell.formulaResult at row {}, column {} cannot combine formulaResult with arrayFormulaRange because array cached results are not modeled; remove formulaResult",
            row + 1,
            column + 1
        )));
    }
    if cell.props.get("dynamicArray").and_then(Value::as_bool) == Some(true) {
        return Err(Error::Compile(format!(
            "Cell.formulaResult at row {}, column {} cannot combine with dynamicArray because dynamic array caches are not modeled; remove formulaResult",
            row + 1,
            column + 1
        )));
    }
    match result {
        Value::Null | Value::Bool(_) => Ok(()),
        Value::Number(number) if number.as_f64().is_some_and(f64::is_finite) => Ok(()),
        Value::String(text) => {
            if !text.chars().all(is_valid_xml_character) {
                return Err(Error::Compile(format!(
                    "Cell.formulaResult at row {}, column {} contains an invalid XML character; remove that character",
                    row + 1,
                    column + 1
                )));
            }
            if formula_result_text_is_ambiguous(text) {
                return Err(Error::Compile(format!(
                    "Cell.formulaResult `{text}` at row {}, column {} looks numeric or boolean and would be serialized with a different result type; use a number or boolean value instead",
                    row + 1,
                    column + 1
                )));
            }
            Ok(())
        }
        _ => Err(Error::Compile(format!(
            "Cell.formulaResult at row {}, column {} must be a string, finite number, boolean, or null; use a scalar cached result",
            row + 1,
            column + 1
        ))),
    }
}

/// Returns whether a cached formula string would be serialized as a number or boolean.
#[must_use]
pub fn formula_result_text_is_ambiguous(text: &str) -> bool {
    text.parse::<f64>().is_ok() || matches!(text, "TRUE" | "FALSE")
}

fn validate_array_formula_cell(
    cell: &Node,
    row: usize,
    column: usize,
    formula: Option<&Value>,
    array_formula_range: Option<&Value>,
) -> Result<()> {
    let Some(array_formula_range) = array_formula_range else {
        return Ok(());
    };
    if formula.is_none() {
        return Err(Error::Compile(format!(
            "Cell.arrayFormulaRange at row {}, column {} requires formula; add the array formula or remove arrayFormulaRange",
            row + 1,
            column + 1
        )));
    }
    if cell.props.contains_key("hyperlink")
        || cell.props.contains_key("formulaResult")
        || cell.props.contains_key("rowSpan")
        || cell.props.contains_key("columnSpan")
    {
        return Err(Error::Compile(format!(
            "array formula at row {}, column {} cannot combine arrayFormulaRange with hyperlink/formulaResult/rowSpan/columnSpan; use only the array formula range",
            row + 1,
            column + 1
        )));
    }
    let Some(range_text) = array_formula_range.as_str() else {
        return Err(Error::Compile(format!(
            "Cell.arrayFormulaRange at row {}, column {} must be an A1 range such as `B1:B2`; use a quoted range",
            row + 1,
            column + 1
        )));
    };
    if crate::parse_a1_range(range_text).is_none() {
        return Err(Error::Compile(format!(
            "Cell.arrayFormulaRange `{range_text}` at row {}, column {} is invalid or unordered; use an ordered A1 range such as `B1:B2`",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_rich_text_cell(cell: &Node, row: usize, column: usize) -> Result<()> {
    if cell.children.is_empty() {
        return Err(Error::Compile(format!(
            "Cell rich text at row {}, column {} requires at least one TextRun child; add a non-empty run or use a scalar Cell value",
            row + 1,
            column + 1
        )));
    }
    if [
        "value",
        "date",
        "formula",
        "formulaResult",
        "hyperlink",
        "hyperlinkTooltip",
        "arrayFormulaRange",
    ]
    .iter()
    .any(|property| cell.props.contains_key(*property))
    {
        return Err(Error::Compile(format!(
            "Cell rich text at row {}, column {} cannot combine TextRun children with value/date/formula/hyperlink; use TextRun children as the only content representation",
            row + 1,
            column + 1
        )));
    }
    let mut total_utf16_units = 0_usize;
    for (index, child) in cell.children.iter().enumerate() {
        let run: Node = serde_json::from_value(child.clone()).map_err(|error| {
            Error::Compile(format!(
                "Cell rich text child {} at row {}, column {} must be a TextRun: {error}",
                index + 1,
                row + 1,
                column + 1
            ))
        })?;
        if run.r#type != "TextRun" {
            return Err(Error::Compile(format!(
                "Cell rich text child {} at row {}, column {} is `{}`; use only TextRun children or remove all TextRun children",
                index + 1,
                row + 1,
                column + 1,
                run.r#type
            )));
        }
        reject_unknown_props(
            &run,
            &[
                "bold",
                "italic",
                "fontColor",
                "fontName",
                "fontSize",
                "fontFamily",
                "fontCharset",
                "fontScheme",
                "underline",
                "strikethrough",
                "script",
            ],
            "TextRun",
        )?;
        validate_text_run_font_scheme(&run, row, column)?;
        if run.children.len() != 1
            || run
                .children
                .first()
                .is_none_or(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(Error::Compile(format!(
                "TextRun {} at row {}, column {} requires exactly one non-empty string child; provide text or use a plain Cell value",
                index + 1,
                row + 1,
                column + 1
            )));
        }
        let text = run.children[0].as_str().unwrap_or_default();
        if !text.chars().all(is_valid_xml_character) {
            return Err(Error::Compile(format!(
                "TextRun {} at row {}, column {} contains a character that is invalid in XML; remove control characters and use valid text",
                index + 1,
                row + 1,
                column + 1
            )));
        }
        total_utf16_units = total_utf16_units
            .checked_add(text.encode_utf16().count())
            .ok_or_else(|| Error::Compile("rich-text length overflow".into()))?;
        if total_utf16_units > 32_767 {
            return Err(Error::Compile(format!(
                "Cell rich text at row {}, column {} exceeds Excel's 32,767 UTF-16-unit text limit; shorten the combined TextRun content",
                row + 1,
                column + 1
            )));
        }
        validate_cell_style(&run, row, column)?;
    }
    Ok(())
}

fn validate_text_run_font_scheme(run: &Node, row: usize, column: usize) -> Result<()> {
    let Some(font_scheme) = run.props.get("fontScheme") else {
        return Ok(());
    };
    if font_scheme
        .as_str()
        .is_some_and(|scheme| matches!(scheme, "none" | "body" | "headings"))
    {
        return Ok(());
    }
    Err(Error::Compile(format!(
        "TextRun.fontScheme at row {}, column {} must be `none`, `body`, or `headings`; use a supported font scheme",
        row + 1,
        column + 1
    )))
}

fn validate_cell_hyperlink(cell: &Node, row: usize, column: usize) -> Result<()> {
    let target = cell.props.get("hyperlink");
    if let Some(tooltip) = cell.props.get("hyperlinkTooltip") {
        let Some(tooltip) = tooltip.as_str() else {
            return Err(Error::Compile(format!(
                "Cell.hyperlinkTooltip at row {}, column {} must be a string; use tooltip text or remove the property",
                row + 1,
                column + 1
            )));
        };
        if !tooltip.chars().all(is_valid_xml_character) {
            return Err(Error::Compile(format!(
                "Cell.hyperlinkTooltip at row {}, column {} contains an invalid XML character; remove that character",
                row + 1,
                column + 1
            )));
        }
        if target.is_none() {
            return Err(Error::Compile(format!(
                "Cell.hyperlinkTooltip at row {}, column {} requires hyperlink; add Cell.hyperlink or remove the tooltip",
                row + 1,
                column + 1
            )));
        }
    }
    let Some(target) = target else { return Ok(()) };
    let Some(target) = target.as_str() else {
        return Err(Error::Compile(format!(
            "Cell.hyperlink at row {}, column {} must be a string URI; provide an external URL",
            row + 1,
            column + 1
        )));
    };
    let internal_anchor = target.strip_prefix("internal:");
    let has_supported_external_scheme =
        target.starts_with("mailto:") || target.starts_with("file://") || target.contains("://");
    if target.is_empty()
        || target.chars().count() > 2080
        || !target.chars().all(is_valid_xml_character)
        || internal_anchor.is_some_and(str::is_empty)
        || (internal_anchor.is_none() && !has_supported_external_scheme)
    {
        return Err(Error::Compile(format!(
            "Cell.hyperlink `{target}` at row {}, column {} must be a non-empty `internal:` workbook location or a supported external URI (http(s)://, mailto:, file://, or another URI containing ://), up to 2,080 XML-valid characters",
            row + 1,
            column + 1
        )));
    }
    if cell.props.contains_key("formula") || cell.props.contains_key("date") {
        return Err(Error::Compile(format!(
            "Cell.hyperlink at row {}, column {} cannot be combined with formula or date; use a text cell",
            row + 1,
            column + 1
        )));
    }
    let value = cell.props.get("value").or_else(|| cell.children.first());
    let Some(label) = value.and_then(Value::as_str) else {
        return Err(Error::Compile(format!(
            "Cell.hyperlink at row {}, column {} requires string text; use value or a string child as the visible label",
            row + 1,
            column + 1
        )));
    };
    if label.chars().count() > 32_767 || !label.chars().all(is_valid_xml_character) {
        return Err(Error::Compile(format!(
            "Cell hyperlink text at row {}, column {} exceeds 32,767 characters or contains an invalid XML character; shorten the visible label",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

/// Returns whether a string is a supported, timezone-free XLSX date or datetime.
#[must_use]
pub fn is_supported_iso_date(value: &str) -> bool {
    if let Ok(date) = chrono::NaiveDate::parse_from_str(value, "%Y-%m-%d") {
        return date.format("%Y-%m-%d").to_string() == value;
    }
    if value.contains('Z') || value.contains('+') || value.matches('-').count() > 2 {
        return false;
    }
    let Some((_, time)) = value.split_once('T') else {
        return false;
    };
    if time.split_once('.').is_some_and(|(_, fraction)| {
        fraction.is_empty()
            || fraction.len() > 3
            || !fraction.bytes().all(|byte| byte.is_ascii_digit())
    }) {
        return false;
    }
    chrono::NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
}

#[expect(
    clippy::too_many_lines,
    reason = "Cell style validation keeps the shared property whitelist and per-cell diagnostics together."
)]
fn validate_cell_style(cell: &Node, row: usize, column: usize) -> Result<()> {
    for property in ["bold", "italic"] {
        if cell
            .props
            .get(property)
            .is_some_and(|value| !value.is_boolean())
        {
            return Err(Error::Compile(format!(
                "Cell `{property}` at row {}, column {} must be a boolean; use {property}={{true}} or {property}={{false}}",
                row + 1,
                column + 1
            )));
        }
    }
    if let Some(value) = cell.props.get("strikethrough")
        && !value.is_boolean()
    {
        return Err(Error::Compile(format!(
            "Cell.strikethrough at row {}, column {} must be a boolean; use strikethrough={{true}} or strikethrough={{false}}",
            row + 1,
            column + 1
        )));
    }
    if let Some(value) = cell.props.get("underline") {
        const UNDERLINES: &[&str] = &["single", "double", "singleAccounting", "doubleAccounting"];
        if !value
            .as_str()
            .is_some_and(|value| UNDERLINES.contains(&value))
        {
            return Err(Error::Compile(format!(
                "Cell.underline at row {}, column {} must be one of {}; use a quoted underline value",
                row + 1,
                column + 1,
                UNDERLINES.join(", ")
            )));
        }
    }
    if let Some(value) = cell.props.get("script") {
        const SCRIPTS: &[&str] = &["superscript", "subscript"];
        if !value.as_str().is_some_and(|value| SCRIPTS.contains(&value)) {
            return Err(Error::Compile(format!(
                "Cell.script at row {}, column {} must be one of {}; use a quoted script value",
                row + 1,
                column + 1,
                SCRIPTS.join(", ")
            )));
        }
    }
    validate_wrap_text(cell, row, column)?;
    validate_cell_text_rotation(cell, row, column)?;
    validate_cell_reading_direction(cell, row, column)?;
    validate_cell_boolean_property(cell, "quotePrefix", row, column)?;
    validate_cell_boolean_property(cell, "locked", row, column)?;
    validate_cell_boolean_property(cell, "formulaHidden", row, column)?;
    for (property, allowed) in [
        (
            "horizontalAlign",
            &["left", "center", "right", "distributed"][..],
        ),
        ("verticalAlign", &["top", "center", "bottom"][..]),
    ] {
        if let Some(value) = cell.props.get(property) {
            let Some(value) = value.as_str() else {
                return Err(Error::Compile(format!(
                    "Cell.{property} at row {}, column {} must be one of {}; use a quoted alignment value",
                    row + 1,
                    column + 1,
                    allowed.join(", ")
                )));
            };
            if !allowed.contains(&value) {
                return Err(Error::Compile(format!(
                    "Cell.{property} at row {}, column {} must be one of {}; use a quoted alignment value",
                    row + 1,
                    column + 1,
                    allowed.join(", ")
                )));
            }
        }
    }
    validate_cell_indent_and_alignment(cell, row, column)?;
    if let Some(number_format) = cell.props.get("numberFormat") {
        let Some(number_format) = number_format.as_str() else {
            return Err(Error::Compile(format!(
                "Cell.numberFormat at row {}, column {} must be a string; use an Excel format code such as `#,##0.00`",
                row + 1,
                column + 1
            )));
        };
        if number_format.is_empty() || number_format.chars().count() > 255 {
            return Err(Error::Compile(format!(
                "Cell.numberFormat at row {}, column {} must contain 1 to 255 characters; shorten or provide an Excel format code",
                row + 1,
                column + 1
            )));
        }
        if !number_format.chars().all(is_valid_xml_character) {
            return Err(Error::Compile(format!(
                "Cell.numberFormat at row {}, column {} contains a character that is invalid in XML; remove control characters and use a valid Excel format code",
                row + 1,
                column + 1
            )));
        }
    }
    validate_cell_font_properties(cell, row, column)?;
    Ok(())
}

fn validate_cell_text_rotation(cell: &Node, row: usize, column: usize) -> Result<()> {
    let Some(rotation) = cell.props.get("textRotation") else {
        return Ok(());
    };
    let valid = rotation
        .as_i64()
        .is_some_and(|angle| (-90..=90).contains(&angle) || angle == 270);
    if !valid {
        return Err(Error::Compile(format!(
            "Cell.textRotation at row {}, column {} must be an integer from -90 through 90 or 270 for stacked text; use a supported text angle",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_cell_reading_direction(cell: &Node, row: usize, column: usize) -> Result<()> {
    let Some(value) = cell.props.get("readingDirection") else {
        return Ok(());
    };
    if !matches!(value.as_str(), Some("leftToRight" | "rightToLeft")) {
        return Err(Error::Compile(format!(
            "Cell.readingDirection at row {}, column {} must be `leftToRight` or `rightToLeft`; use a supported cell reading order",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_cell_boolean_property(
    cell: &Node,
    property: &str,
    row: usize,
    column: usize,
) -> Result<()> {
    if cell
        .props
        .get(property)
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(Error::Compile(format!(
            "Cell.{property} at row {}, column {} must be a boolean; use `true` to enable the format flag or `false` to leave it disabled",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_cell_indent_and_alignment(cell: &Node, row: usize, column: usize) -> Result<()> {
    if let Some(indent) = cell.props.get("indent") {
        let Some(indent) = indent.as_u64() else {
            return Err(Error::Compile(format!(
                "Cell.indent at row {}, column {} must be an integer from 1 through 255; use a positive Excel indentation level",
                row + 1,
                column + 1
            )));
        };
        if !(1..=255).contains(&indent) {
            return Err(Error::Compile(format!(
                "Cell.indent at row {}, column {} must be from 1 through 255; use a supported Excel indentation level",
                row + 1,
                column + 1
            )));
        }
        let horizontal = cell.props.get("horizontalAlign").and_then(Value::as_str);
        if !matches!(horizontal, Some("left" | "right" | "distributed")) {
            return Err(Error::Compile(format!(
                "Cell.indent at row {}, column {} requires horizontalAlign to be explicitly left, right, or distributed; choose one so the writer does not normalize the alignment",
                row + 1,
                column + 1
            )));
        }
    }
    if cell.props.get("horizontalAlign").and_then(Value::as_str) == Some("distributed")
        && cell.props.get("shrinkToFit").and_then(Value::as_bool) == Some(true)
    {
        return Err(Error::Compile(format!(
            "Cell.horizontalAlign=distributed and Cell.shrinkToFit at row {}, column {} cannot be combined because the XLSX writer suppresses shrink-to-fit for distributed alignment; choose one behavior",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_wrap_text(cell: &Node, row: usize, column: usize) -> Result<()> {
    for property in ["wrapText", "shrinkToFit"] {
        if let Some(value) = cell.props.get(property)
            && !value.is_boolean()
        {
            return Err(Error::Compile(format!(
                "Cell.{property} at row {}, column {} must be a boolean; use {property}={{true}} or {property}={{false}}",
                row + 1,
                column + 1
            )));
        }
    }
    if cell.props.get("wrapText").and_then(Value::as_bool) == Some(true)
        && cell.props.get("shrinkToFit").and_then(Value::as_bool) == Some(true)
    {
        return Err(Error::Compile(format!(
            "Cell.wrapText and Cell.shrinkToFit at row {}, column {} cannot both be true because the XLSX writer suppresses shrink-to-fit when wrapping; choose one behavior",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_cell_font_properties(cell: &Node, row: usize, column: usize) -> Result<()> {
    for property in [
        "fillColor",
        "fillForegroundColor",
        "fillBackgroundColor",
        "fontColor",
    ] {
        if let Some(color) = cell.props.get(property) {
            validate_rgb_color(color, property, row, column)?;
        }
    }
    validate_cell_fill_properties(cell, row, column)?;
    if let Some(font_name) = cell.props.get("fontName") {
        let Some(font_name) = font_name.as_str() else {
            return Err(Error::Compile(format!(
                "Cell.fontName at row {}, column {} must be a string; use a literal font family name",
                row + 1,
                column + 1
            )));
        };
        if font_name.is_empty()
            || font_name.chars().count() > 255
            || !font_name.chars().all(is_valid_xml_character)
        {
            return Err(Error::Compile(format!(
                "Cell.fontName at row {}, column {} must contain 1 to 255 valid XML characters; shorten it or remove control characters",
                row + 1,
                column + 1
            )));
        }
    }
    if let Some(font_size) = cell.props.get("fontSize") {
        let Some(font_size) = font_size.as_f64() else {
            return Err(Error::Compile(format!(
                "Cell.fontSize at row {}, column {} must be a number in points from 1 through 409",
                row + 1,
                column + 1
            )));
        };
        if !(1.0..=409.0).contains(&font_size) {
            return Err(Error::Compile(format!(
                "Cell.fontSize at row {}, column {} must be from 1 through 409 points; choose an OOXML-compatible font size",
                row + 1,
                column + 1
            )));
        }
    }
    if let Some(font_family) = cell.props.get("fontFamily") {
        let Some(font_family) = font_family.as_u64() else {
            return Err(Error::Compile(format!(
                "Cell.fontFamily at row {}, column {} must be an integer from 0 through 5; use an OOXML font family value",
                row + 1,
                column + 1
            )));
        };
        if font_family > 5 {
            return Err(Error::Compile(format!(
                "Cell.fontFamily at row {}, column {} must be from 0 through 5; use 0 (default), 1 (roman), 2 (swiss), 3 (modern), 4 (script), or 5 (decorative)",
                row + 1,
                column + 1
            )));
        }
    }
    if let Some(font_charset) = cell.props.get("fontCharset")
        && font_charset
            .as_u64()
            .is_none_or(|charset| u8::try_from(charset).is_err())
    {
        return Err(Error::Compile(format!(
            "Cell.fontCharset at row {}, column {} must be an integer byte from 0 through 255; choose an OOXML charset value",
            row + 1,
            column + 1
        )));
    }
    if let Some(font_scheme) = cell.props.get("fontScheme")
        && !font_scheme
            .as_str()
            .is_some_and(|scheme| matches!(scheme, "none" | "body" | "headings"))
    {
        return Err(Error::Compile(format!(
            "Cell.fontScheme at row {}, column {} must be `none`, `body`, or `headings`; use a supported font scheme",
            row + 1,
            column + 1
        )));
    }
    validate_cell_border_properties(cell, row, column)?;
    Ok(())
}

fn validate_cell_fill_properties(cell: &Node, row: usize, column: usize) -> Result<()> {
    if let Some(value) = cell.props.get("fillPattern")
        && !value
            .as_str()
            .is_some_and(|pattern| FILL_PATTERN_NAMES.contains(&pattern))
    {
        return Err(Error::Compile(format!(
            "Cell.fillPattern at row {}, column {} must be one of {}; use a supported OOXML pattern name",
            row + 1,
            column + 1,
            FILL_PATTERN_NAMES.join(", ")
        )));
    }
    let has_pattern = cell.props.contains_key("fillPattern");
    let has_color = cell.props.contains_key("fillColor");
    let has_foreground = cell.props.contains_key("fillForegroundColor");
    let has_background = cell.props.contains_key("fillBackgroundColor");
    let has_pattern_colors = has_foreground || has_background;
    let invalid_combination = (has_pattern || has_pattern_colors) && (!has_pattern || has_color);
    let colors_on_none =
        cell.props.get("fillPattern").and_then(Value::as_str) == Some("none") && has_pattern_colors;
    if invalid_combination || colors_on_none {
        return Err(Error::Compile(format!(
            "Cell fill properties at row {}, column {} conflict: fillColor is the solid-fill shorthand; foreground/background colors require a non-none fillPattern. Choose one fill representation",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_cell_border_properties(cell: &Node, row: usize, column: usize) -> Result<()> {
    const SIDES: [&str; 4] = ["Top", "Right", "Bottom", "Left"];
    const STYLES: &[&str] = &[
        "none",
        "thin",
        "medium",
        "dashed",
        "dotted",
        "thick",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "mediumDashDotDot",
        "slantDashDot",
    ];
    for suffix in SIDES {
        let style_property = format!("border{suffix}");
        let color_property = format!("border{suffix}Color");
        let border_style = if let Some(value) = cell.props.get(&style_property) {
            let Some(value) = value.as_str() else {
                return Err(Error::Compile(format!(
                    "Cell.{style_property} at row {}, column {} must be one of {}; use a quoted border style",
                    row + 1,
                    column + 1,
                    STYLES.join(", ")
                )));
            };
            if !STYLES.contains(&value) {
                return Err(Error::Compile(format!(
                    "Cell.{style_property} at row {}, column {} must be one of {}; use a quoted border style",
                    row + 1,
                    column + 1,
                    STYLES.join(", ")
                )));
            }
            Some(value)
        } else {
            None
        };
        if let Some(color) = cell.props.get(&color_property) {
            if border_style.is_none_or(|style| style == "none") {
                return Err(Error::Compile(format!(
                    "Cell.{color_property} at row {}, column {} requires a matching non-`none` {style_property}; add the edge style or remove the color",
                    row + 1,
                    column + 1
                )));
            }
            validate_rgb_color(color, &color_property, row, column)?;
        }
    }
    let diagonal_style = cell.props.get("borderDiagonal");
    if let Some(value) = diagonal_style {
        let Some(value) = value.as_str() else {
            return Err(Error::Compile(format!(
                "Cell.borderDiagonal at row {}, column {} must be one of {}; use a quoted border style",
                row + 1,
                column + 1,
                STYLES.join(", ")
            )));
        };
        if !STYLES.contains(&value) {
            return Err(Error::Compile(format!(
                "Cell.borderDiagonal at row {}, column {} must be one of {}; use a quoted border style",
                row + 1,
                column + 1,
                STYLES.join(", ")
            )));
        }
    }
    if let Some(color) = cell.props.get("borderDiagonalColor") {
        if diagonal_style
            .and_then(Value::as_str)
            .is_none_or(|style| style == "none")
        {
            return Err(Error::Compile(format!(
                "Cell.borderDiagonalColor at row {}, column {} requires a matching non-`none` borderDiagonal; add the diagonal style or remove the color",
                row + 1,
                column + 1
            )));
        }
        validate_rgb_color(color, "borderDiagonalColor", row, column)?;
    }
    if let Some(direction) = cell.props.get("borderDiagonalType")
        && !direction
            .as_str()
            .is_some_and(|value| ["up", "down", "upDown"].contains(&value))
    {
        return Err(Error::Compile(format!(
            "Cell.borderDiagonalType at row {}, column {} must be `up`, `down`, or `upDown`; choose a supported direction",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn validate_rgb_color(value: &Value, property: &str, row: usize, column: usize) -> Result<()> {
    let Some(color) = value.as_str() else {
        return Err(Error::Compile(format!(
            "Cell.{property} at row {}, column {} must be a `#RRGGBB` RGB string; use six hexadecimal digits",
            row + 1,
            column + 1
        )));
    };
    if color.len() != 7
        || !color.starts_with('#')
        || !color[1..].bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        return Err(Error::Compile(format!(
            "Cell.{property} at row {}, column {} must match `#RRGGBB`; use exactly six hexadecimal digits after `#`",
            row + 1,
            column + 1
        )));
    }
    Ok(())
}

fn is_valid_xml_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || ('\u{20}'..='\u{d7ff}').contains(&character)
        || ('\u{e000}'..='\u{fffd}').contains(&character)
        || ('\u{10000}'..='\u{10ffff}').contains(&character)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn workbook(sheets: Vec<Value>) -> IrEnvelope {
        IrEnvelope {
            version: 1,
            workbook: Node {
                r#type: "Workbook".into(),
                props: serde_json::Map::default(),
                children: sheets,
            },
        }
    }

    fn worksheet(name: &str, rows: Vec<Value>) -> Value {
        let mut worksheet = serde_json::json!({"type":"Worksheet","props":{"name":name}});
        worksheet["children"] = Value::Array(rows);
        worksheet
    }

    #[test]
    fn accepts_minimal_workbook() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![
                serde_json::json!({"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["ok"]}]}),
            ],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_note_component_with_supported_properties() {
        let input = workbook(vec![worksheet(
            "Notes",
            vec![
                serde_json::json!({
                    "type":"Note",
                    "props":{"cell":"B3","text":"Review","author":"Ada","authorPrefix":false,"visible":true,"width":200,"height":100,"backgroundColor":"#AABBCC","fontName":"Arial","fontSize":9.5,"fontFamily":2,"altText":"Review note"},
                    "children":[]
                }),
                serde_json::json!({"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["value"]}]}),
            ],
        )]);
        let validation = input.validate();
        assert!(validation.is_ok(), "{validation:?}");
    }

    #[test]
    fn rejects_zero_pixel_note_dimensions_with_repair() {
        let input = workbook(vec![worksheet(
            "Notes",
            vec![
                serde_json::json!({"type":"Note","props":{"cell":"A1","text":"first","width":0},"children":[]}),
            ],
        )]);
        let error = input.validate().expect_err("zero dimensions are invalid");
        assert!(error.to_string().contains("Note.width"));
        assert!(error.to_string().contains("positive 32-bit integer"));
    }

    #[test]
    fn rejects_duplicate_note_cells_with_repair() {
        let input = workbook(vec![worksheet(
            "Notes",
            vec![
                serde_json::json!({"type":"Note","props":{"cell":"A1","text":"first"},"children":[]}),
                serde_json::json!({"type":"Note","props":{"cell":"a1","text":"second"},"children":[]}),
            ],
        )]);
        let error = input
            .validate()
            .expect_err("only one note per cell is supported");
        assert!(error.to_string().contains("more than one Note targets A1"));
    }

    #[test]
    fn rejects_notes_declared_after_rows_with_repair() {
        let input = workbook(vec![worksheet(
            "Notes",
            vec![
                serde_json::json!({"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["value"]}]}),
                serde_json::json!({"type":"Note","props":{"cell":"A1","text":"late"},"children":[]}),
            ],
        )]);
        let error = input.validate().expect_err("notes must precede row data");
        assert!(error.to_string().contains("move notes before rows"));
    }

    #[test]
    fn accepts_global_and_worksheet_scoped_defined_names() {
        let mut sheet = worksheet("Data", vec![]);
        sheet["children"].as_array_mut().unwrap().push(serde_json::json!({
            "type":"DefinedName","props":{"name":"LocalRate","formula":"Data!$A$1"},"children":[]
        }));
        let input = workbook(vec![
            serde_json::json!({"type":"DefinedName","props":{"name":"TaxRate","formula":"0.05"},"children":[]}),
            sheet,
        ]);
        let validation = input.validate();
        assert!(validation.is_ok(), "{validation:?}");
    }

    #[test]
    fn accepts_dynamic_array_formulas_with_and_without_a_known_range() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"formula":"=SEQUENCE(3)","dynamicArray":true,"arrayFormulaRange":"A1:A3"},"children":[]},
                    {"type":"Cell","props":{"formula":"=FILTER(A1:A3,A1:A3<>\"\")","dynamicArray":true},"children":[]}
                ]
            })],
        )]);
        let validation = input.validate();
        assert!(validation.is_ok(), "{validation:?}");
    }

    #[test]
    fn rejects_dynamic_array_without_formula_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"dynamicArray":true},"children":[]}
                ]
            })],
        )]);
        let error = input.validate().unwrap_err().to_string();
        assert!(error.contains("dynamicArray"), "{error}");
        assert!(error.contains("formula"), "{error}");
    }

    #[test]
    fn rejects_reserved_defined_name_with_actionable_diagnostic() {
        let input = workbook(vec![
            serde_json::json!({"type":"DefinedName","props":{"name":"_xlnm.Print_Area","formula":"Data!$A$1"},"children":[]}),
            worksheet("Data", vec![]),
        ]);
        let error = input.validate().unwrap_err().to_string();
        assert!(error.contains("reserved"), "{error}");
        assert!(error.contains("_xlnm"), "{error}");
    }

    #[test]
    fn rejects_invalid_defined_name_syntax_with_repair_guidance() {
        let input = workbook(vec![
            serde_json::json!({"type":"DefinedName","props":{"name":"Bad Name","formula":"1"},"children":[]}),
            worksheet("Data", vec![]),
        ]);
        let error = input.validate().unwrap_err().to_string();
        assert!(error.contains("valid Excel name"), "{error}");
        assert!(error.contains("without spaces"), "{error}");
    }

    #[test]
    fn accepts_supported_core_document_properties() {
        let mut input = workbook(vec![worksheet("Data", vec![])]);
        input.workbook.props = serde_json::from_value(serde_json::json!({
            "title": "Quarterly report",
            "subject": "Revenue",
            "author": "Analyst",
            "keywords": "quarterly,finance",
            "description": "Reviewed & approved"
        }))
        .expect("properties are a JSON object");

        let validation = input.validate();
        assert!(validation.is_ok(), "{validation:?}");
    }

    #[test]
    fn accepts_row_and_column_default_format_properties() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Data"},"children":[
                {"type":"Column","props":{"min":1,"max":2,"fontColor":"#123456","numberFormat":"0.00"},"children":[]},
                {"type":"Row","props":{"bold":true,"fillColor":"#ABCDEF"},"children":[
                    {"type":"Cell","props":{},"children":["inherited"]}
                ]}
            ]
        })]);

        let validation = input.validate();
        assert!(validation.is_ok(), "{validation:?}");
    }

    #[test]
    fn rejects_conflicting_row_default_format_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Data"},"children":[
                {"type":"Row","props":{"wrapText":true,"shrinkToFit":true},"children":[]}
            ]
        })]);

        let error = input
            .validate()
            .expect_err("these flags conflict in the writer");

        assert!(
            error
                .to_string()
                .contains("Row.wrapText and Row.shrinkToFit")
        );
        assert!(error.to_string().contains("choose one behavior"));
    }

    #[test]
    fn rejects_non_string_workbook_document_properties_with_guidance() {
        let mut input = workbook(vec![worksheet("Data", vec![])]);
        input.workbook.props = serde_json::from_value(serde_json::json!({"title": 42}))
            .expect("properties are a JSON object");

        let error = input
            .validate()
            .expect_err("document properties must be strings");
        assert!(
            error
                .to_string()
                .contains("Workbook.title must be a string")
        );
        assert!(error.to_string().contains("quoted string"));
    }

    #[test]
    fn accepts_hidden_and_very_hidden_secondary_worksheets() {
        let input = workbook(vec![
            serde_json::json!({
                "type":"Worksheet","props":{"name":"Visible"},"children":[]
            }),
            serde_json::json!({
                "type":"Worksheet","props":{"name":"Hidden","visibility":"hidden"},"children":[]
            }),
            serde_json::json!({
                "type":"Worksheet","props":{"name":"VeryHidden","visibility":"veryHidden"},"children":[]
            }),
        ]);

        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_selected_hidden_worksheets_and_non_boolean_selection() {
        let hidden = workbook(vec![
            serde_json::json!({"type":"Worksheet","props":{"name":"Visible"},"children":[]}),
            serde_json::json!({"type":"Worksheet","props":{"name":"Hidden","visibility":"hidden","selected":true},"children":[]}),
        ]);
        let error = hidden
            .validate()
            .expect_err("selected hidden worksheets are not representable");
        assert!(
            error
                .to_string()
                .contains("selected Worksheet `Hidden` must be visible")
        );

        let invalid = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Invalid","selected":"yes"},"children":[]
        })]);
        let error = invalid.validate().expect_err("selected must be boolean");
        assert!(error.to_string().contains("Worksheet.selected"));
        assert!(error.to_string().contains("must be a boolean"));

        let contradictory = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Active","active":true,"selected":false},"children":[]
        })]);
        let error = contradictory
            .validate()
            .expect_err("writer cannot activate an unselected worksheet");
        assert!(
            error
                .to_string()
                .contains("is the active worksheet but selected=false")
        );

        let default_active = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"First","selected":false},"children":[]
        })]);
        let error = default_active
            .validate()
            .expect_err("default first active worksheet cannot be unselected");
        assert!(
            error
                .to_string()
                .contains("is the active worksheet but selected=false")
        );
    }

    #[test]
    fn rejects_hidden_first_worksheet_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Hidden","visibility":"hidden"},"children":[]
        })]);
        let error = input
            .validate()
            .expect_err("the first worksheet must remain visible");

        assert!(
            error
                .to_string()
                .contains("first Worksheet must have visibility")
        );
        assert!(error.to_string().contains("move the visible sheet first"));
    }

    #[test]
    fn accepts_positive_default_row_height_pixels() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Rows","defaultRowHeightPixels":24},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_non_positive_or_fractional_default_row_height_with_guidance() {
        for value in [
            serde_json::json!(0),
            serde_json::json!(24.5),
            serde_json::json!(546),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":{"name":"Rows","defaultRowHeightPixels":value},"children":[]
            })]);
            let error = input
                .validate()
                .expect_err("invalid row height must be rejected");
            assert!(error.to_string().contains("defaultRowHeightPixels"));
            assert!(error.to_string().contains("use") || error.to_string().contains("choose"));
        }
    }

    #[test]
    fn accepts_boolean_hide_unused_rows_property() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Hidden unused","hideUnusedRows":true},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_non_boolean_hide_unused_rows_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Hidden unused","hideUnusedRows":"yes"},"children":[]
        })]);
        let error = input.validate().expect_err("property must be boolean");
        assert!(error.to_string().contains("hideUnusedRows"));
        assert!(error.to_string().contains("boolean"));
    }

    #[test]
    fn accepts_boolean_show_gridlines_property() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Grid","showGridlines":false},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_non_boolean_show_gridlines_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Grid","showGridlines":"no"},"children":[]
        })]);
        let error = input.validate().expect_err("property must be boolean");
        assert!(error.to_string().contains("showGridlines"));
        assert!(error.to_string().contains("boolean"));
    }

    #[test]
    fn accepts_boolean_print_gridlines_property() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Printed","printGridlines":true},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_non_boolean_print_gridlines_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Printed","printGridlines":"yes"},"children":[]
        })]);
        let error = input.validate().expect_err("property must be boolean");
        assert!(error.to_string().contains("printGridlines"));
        assert!(error.to_string().contains("boolean"));
    }

    #[test]
    fn accepts_worksheet_zoom_in_excel_range() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Zoomed","zoom":125},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_supported_worksheet_view_modes() {
        for view in ["normal", "pageLayout", "pageBreakPreview"] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":{"name":"View","view":view},"children":[]
            })]);
            assert!(input.validate().is_ok(), "view `{view}` should validate");
        }
    }

    #[test]
    fn accepts_manual_page_breaks_within_sheet_limits() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","horizontalPageBreaks":[20, 40],"verticalPageBreaks":[8]},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_invalid_manual_page_breaks_with_repair_guidance() {
        for breaks in [
            serde_json::json!([0]),
            serde_json::json!([4, 4]),
            serde_json::json!([1_048_576]),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":{"name":"Print","horizontalPageBreaks":breaks},"children":[]
            })]);
            let error = input
                .validate()
                .expect_err("page-break index is unsupported");
            assert!(error.to_string().contains("horizontalPageBreaks"));
            assert!(error.to_string().contains("unique"));
        }
    }

    #[test]
    fn accepts_partial_page_margins_in_inches() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","pageMargins":{"left":0.5,"top":1.0}},"children":[]
        })]);
        assert!(
            input.validate().is_ok(),
            "validation error: {:?}",
            input.validate()
        );
    }

    #[test]
    fn rejects_invalid_page_margins_with_repair_guidance() {
        for margins in [
            serde_json::json!({"left":-0.1}),
            serde_json::json!({"inside":0.5}),
            serde_json::json!({}),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":{"name":"Print","pageMargins":margins},"children":[]
            })]);
            let error = input
                .validate()
                .expect_err("invalid page margins must be rejected");
            assert!(error.to_string().contains("pageMargins"));
        }
    }

    #[test]
    fn accepts_print_scale_and_fit_to_pages() {
        for props in [
            serde_json::json!({"name":"Print","printScale":80}),
            serde_json::json!({"name":"Print","fitToPages":{"width":1,"height":0}}),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":props,"children":[]
            })]);
            assert!(
                input.validate().is_ok(),
                "validation error: {:?}",
                input.validate()
            );
        }
    }

    #[test]
    fn accepts_page_orientation_and_paper_size() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","orientation":"landscape","paperSize":9},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_supported_print_layout_and_quality_flags() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","centerHorizontally":true,"centerVertically":true,"printBlackAndWhite":true,"printDraft":true},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_non_boolean_print_layout_and_quality_flags() {
        for (property, value) in [
            ("centerHorizontally", serde_json::json!("yes")),
            ("centerVertically", serde_json::json!(1)),
            ("printBlackAndWhite", serde_json::json!("true")),
            ("printDraft", serde_json::json!(0)),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":{"name":"Print",property:value},"children":[]
            })]);
            let error = input.validate().expect_err("print option must be boolean");
            assert!(error.to_string().contains(property));
            assert!(error.to_string().contains("boolean"));
        }
    }

    #[test]
    fn accepts_print_headings_and_supported_page_order() {
        for props in [
            serde_json::json!({"name":"Print","printHeadings":true}),
            serde_json::json!({"name":"Print","pageOrder":"overThenDown"}),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":props,"children":[]
            })]);
            assert!(
                input.validate().is_ok(),
                "validation failed: {:?}",
                input.validate()
            );
        }
    }

    #[test]
    fn accepts_print_area_and_repeating_titles() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","printArea":"A1:D20","repeatRows":"1:3","repeatColumns":"A:B"},"children":[]
        })]);
        assert!(
            input.validate().is_ok(),
            "validation failed: {:?}",
            input.validate()
        );
    }

    #[test]
    fn accepts_header_footer_strings_and_behavior_flags() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","header":"&CPage &[Page] of &[Pages]","footer":"&R&D","headerFooterScaleWithDoc":false,"headerFooterAlignWithPage":false},"children":[]
        })]);
        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn rejects_unrepresentable_or_oversized_headers_and_invalid_flags() {
        for (property, value) in [
            ("header", serde_json::json!("&C&G")),
            ("footer", serde_json::json!("x".repeat(256))),
            ("headerFooterScaleWithDoc", serde_json::json!("false")),
            ("headerFooterAlignWithPage", serde_json::json!(0)),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":{"name":"Print",property:value},"children":[]
            })]);
            let error = input
                .validate()
                .expect_err("invalid header/footer property");
            assert!(error.to_string().contains(property), "{error}");
            assert!(
                error.to_string().contains("use") || error.to_string().contains("must"),
                "{error}"
            );
        }
    }

    #[test]
    fn rejects_invalid_print_area_and_repeating_title_ranges() {
        for (property, value) in [
            ("printArea", "D1:A4"),
            ("printArea", "A1:B2,C3:D4"),
            ("repeatRows", "1:3,5:6"),
            ("repeatColumns", "A:C,E:F"),
            ("repeatRows", "A:C"),
            ("repeatColumns", "1:3"),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":{"name":"Print",property:value},"children":[]
            })]);
            let error = input.validate().expect_err("invalid print range");
            assert!(error.to_string().contains(property), "{error}");
        }
    }

    #[test]
    fn rejects_invalid_print_headings_and_page_order() {
        for (property, value) in [
            ("printHeadings", serde_json::json!("yes")),
            ("pageOrder", serde_json::json!("across")),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":{"name":"Print",property:value},"children":[]
            })]);
            let error = input.validate().expect_err("invalid print property");
            assert!(error.to_string().contains(property));
        }
    }

    #[test]
    fn rejects_invalid_page_orientation_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","orientation":"sideways"},"children":[]
        })]);
        let error = input
            .validate()
            .expect_err("unknown orientation is rejected");
        assert!(error.to_string().contains("Worksheet.orientation"));
        assert!(error.to_string().contains("portrait"));
    }

    #[test]
    fn rejects_paper_size_outside_unsigned_byte_range() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","paperSize":256},"children":[]
        })]);
        let error = input
            .validate()
            .expect_err("paper code exceeds writer range");
        assert!(error.to_string().contains("Worksheet.paperSize"));
        assert!(error.to_string().contains("0 through 255"));
    }

    #[test]
    fn rejects_invalid_print_scale_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","printScale":401},"children":[]
        })]);
        let error = input
            .validate()
            .expect_err("scale exceeds the supported range");
        assert!(error.to_string().contains("printScale"));
        assert!(error.to_string().contains("10 through 400"));
    }

    #[test]
    fn rejects_conflicting_print_settings_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","printScale":80,"fitToPages":{"width":1,"height":0}},"children":[]
        })]);
        let error = input
            .validate()
            .expect_err("print scale and fit-to-pages are exclusive");
        assert!(error.to_string().contains("mutually exclusive"));
        assert!(error.to_string().contains("printScale"));
    }

    #[test]
    fn rejects_fit_to_pages_with_manual_breaks() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Print","fitToPages":{"width":1,"height":0},"horizontalPageBreaks":[20]},"children":[]
        })]);
        let error = input
            .validate()
            .expect_err("fit-to-pages overrides manual breaks");
        assert!(error.to_string().contains("fitToPages"));
        assert!(error.to_string().contains("manual page breaks"));
    }

    #[test]
    fn rejects_unknown_worksheet_view_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"View","view":"print"},"children":[]
        })]);
        let error = input.validate().expect_err("unknown view is not supported");
        assert!(error.to_string().contains("Worksheet.view"));
        assert!(error.to_string().contains("pageBreakPreview"));
    }

    #[test]
    fn accepts_boolean_worksheet_reading_direction() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Arabic","rightToLeft":true},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_non_boolean_worksheet_reading_direction_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Arabic","rightToLeft":"rtl"},"children":[]
        })]);
        let error = input.validate().expect_err("rightToLeft must be a boolean");
        assert!(error.to_string().contains("Worksheet.rightToLeft"));
        assert!(error.to_string().contains("boolean"));
        assert!(error.to_string().contains("use true or false"));
    }

    #[test]
    fn accepts_boolean_outline_summary_directions() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Outline","summaryBelow":false,"summaryRight":false},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_opaque_rgb_worksheet_tab_color() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Colored","tabColor":"#4472C4"},"children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_one_visible_active_worksheet() {
        let input = workbook(vec![
            serde_json::json!({"type":"Worksheet","props":{"name":"First"},"children":[]}),
            serde_json::json!({"type":"Worksheet","props":{"name":"Second","active":true},"children":[]}),
        ]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_one_visible_first_tab_worksheet() {
        let input = workbook(vec![
            serde_json::json!({"type":"Worksheet","props":{"name":"First"},"children":[]}),
            serde_json::json!({"type":"Worksheet","props":{"name":"Second","firstTab":true},"children":[]}),
        ]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_non_boolean_first_tab_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Second","firstTab":"yes"},"children":[]
        })]);
        let error = input.validate().expect_err("firstTab must be boolean");
        assert!(error.to_string().contains("Worksheet.firstTab"));
        assert!(error.to_string().contains("must be a boolean"));
        assert!(error.to_string().contains("firstTab={true}"));
    }

    #[test]
    fn rejects_multiple_first_tab_worksheets_with_repair_guidance() {
        let input = workbook(vec![
            serde_json::json!({"type":"Worksheet","props":{"name":"First","firstTab":true},"children":[]}),
            serde_json::json!({"type":"Worksheet","props":{"name":"Second","firstTab":true},"children":[]}),
        ]);
        let error = input
            .validate()
            .expect_err("only one worksheet can be first visible tab");
        assert!(error.to_string().contains("both have firstTab=true"));
    }

    #[test]
    fn rejects_hidden_first_tab_worksheet_with_repair_guidance() {
        let input = workbook(vec![
            serde_json::json!({"type":"Worksheet","props":{"name":"First"},"children":[]}),
            serde_json::json!({"type":"Worksheet","props":{"name":"Hidden","visibility":"hidden","firstTab":true},"children":[]}),
        ]);
        let error = input
            .validate()
            .expect_err("a hidden worksheet cannot be the first visible tab");
        assert!(
            error
                .to_string()
                .contains("firstTab Worksheet `Hidden` must be visible")
        );
    }

    #[test]
    fn rejects_multiple_active_worksheets_with_repair_guidance() {
        let input = workbook(vec![
            serde_json::json!({"type":"Worksheet","props":{"name":"First","active":true},"children":[]}),
            serde_json::json!({"type":"Worksheet","props":{"name":"Second","active":true},"children":[]}),
        ]);
        let error = input
            .validate()
            .expect_err("only one active worksheet is allowed");
        assert!(error.to_string().contains("both have active=true"));
    }

    #[test]
    fn rejects_hidden_active_worksheet_with_repair_guidance() {
        let input = workbook(vec![
            serde_json::json!({"type":"Worksheet","props":{"name":"First"},"children":[]}),
            serde_json::json!({"type":"Worksheet","props":{"name":"Hidden","visibility":"hidden","active":true},"children":[]}),
        ]);
        let error = input
            .validate()
            .expect_err("hidden worksheet cannot be active");
        assert!(
            error
                .to_string()
                .contains("active Worksheet `Hidden` must be visible")
        );
    }

    #[test]
    fn rejects_invalid_worksheet_tab_color_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Colored","tabColor":"#12FG"},"children":[]
        })]);
        let error = input.validate().expect_err("tab color must be #RRGGBB");
        assert!(error.to_string().contains("Worksheet.tabColor"));
        assert!(error.to_string().contains("#RRGGBB"));
        assert!(error.to_string().contains("six hexadecimal digits"));
    }

    #[test]
    fn rejects_non_boolean_outline_summary_direction_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Outline","summaryBelow":"above"},"children":[]
        })]);
        let error = input.validate().expect_err("summaryBelow must be boolean");
        assert!(error.to_string().contains("summaryBelow"));
        assert!(error.to_string().contains("boolean"));
        assert!(error.to_string().contains("use false"));
    }

    #[test]
    fn rejects_invalid_worksheet_zoom_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Zoomed","zoom":401},"children":[]
        })]);
        let error = input
            .validate()
            .expect_err("zoom must be within Excel's range");
        assert!(error.to_string().contains("zoom"));
        assert!(error.to_string().contains("10 through 400"));
        assert!(
            error
                .to_string()
                .contains("choose a valid screen zoom percentage")
        );
    }

    #[test]
    fn accepts_collapsed_row_after_hidden_outlined_rows() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Rows"},"children":[
                {"type":"Row","props":{"hidden":true,"outlineLevel":1},"children":[]},
                {"type":"Row","props":{"collapsed":true},"children":[]}
            ]
        })]);
        assert!(
            input.validate().is_ok(),
            "validation result: {:?}",
            input.validate()
        );
    }

    #[test]
    fn rejects_collapsed_row_without_hidden_outline_members() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Rows"},"children":[
                {"type":"Row","props":{"collapsed":true},"children":[]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("collapsed row needs grouped rows before it");
        assert!(error.to_string().contains("collapsed row"));
        assert!(error.to_string().contains("hidden outline rows"));
    }

    #[test]
    fn rejects_non_boolean_row_collapsed_marker_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Rows"},"children":[
                {"type":"Row","props":{"collapsed":"yes"},"children":[
                    {"type":"Cell","props":{},"children":["summary"]}
                ]}
            ]
        })]);
        let error = input.validate().expect_err("collapsed must be a boolean");
        assert!(error.to_string().contains("Row.collapsed"));
        assert!(error.to_string().contains("boolean"));
    }

    #[test]
    fn accepts_collapsed_column_after_hidden_outlined_columns() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Columns"},"children":[
                {"type":"Column","props":{"min":1,"max":2,"hidden":true,"outlineLevel":1},"children":[]},
                {"type":"Column","props":{"min":3,"max":3,"collapsed":true},"children":[]}
            ]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_collapsed_column_range_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Columns"},"children":[
                {"type":"Column","props":{"min":1,"max":2,"collapsed":true},"children":[]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("collapsed marker must name one column");
        assert!(error.to_string().contains("Column.collapsed"));
        assert!(error.to_string().contains("min=max"));
    }

    #[test]
    fn rejects_non_boolean_column_collapsed_marker_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Columns"},"children":[
                {"type":"Column","props":{"min":1,"max":1,"collapsed":"yes"},"children":[]}
            ]
        })]);
        let error = input.validate().expect_err("collapsed must be a boolean");
        assert!(error.to_string().contains("Column.collapsed"));
        assert!(error.to_string().contains("boolean"));
    }

    #[test]
    fn derives_nested_collapsed_outline_ranges() {
        let entries = std::collections::BTreeMap::from([
            (
                1,
                OutlineProperties {
                    hidden: true,
                    level: 2,
                    collapsed: false,
                },
            ),
            (
                2,
                OutlineProperties {
                    hidden: true,
                    level: 2,
                    collapsed: false,
                },
            ),
            (
                3,
                OutlineProperties {
                    hidden: true,
                    level: 1,
                    collapsed: true,
                },
            ),
            (
                4,
                OutlineProperties {
                    hidden: true,
                    level: 1,
                    collapsed: false,
                },
            ),
            (
                5,
                OutlineProperties {
                    hidden: false,
                    level: 0,
                    collapsed: true,
                },
            ),
        ]);
        assert_eq!(
            collapsed_outline_ranges(&entries, "row"),
            Ok(vec![(1, 2), (1, 4)])
        );
    }

    #[test]
    fn accepts_anchored_array_formula_range() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Array"},"children":[
                {"type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"formula":"=A1:A2*2","arrayFormulaRange":"A1:B2"},"children":[]}
                ]},
                {"type":"Row","props":{},"children":[]}
            ]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_unanchored_array_formula_range_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet","props":{"name":"Array"},"children":[
                {"type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"formula":"=A1:A2*2","arrayFormulaRange":"B1:C2"},"children":[]}
                ]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("anchor must be top-left of array range");
        assert!(error.to_string().contains("arrayFormulaRange"));
        assert!(error.to_string().contains("top-left"));
    }

    #[test]
    fn accepts_supported_formula_cached_result_types() {
        let input = workbook(vec![worksheet(
            "Formula results",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"formula":"=1+1","formulaResult":2},"children":[]},
                    {"type":"Cell","props":{"formula":"=\"ready\"","formulaResult":"ready"},"children":[]},
                    {"type":"Cell","props":{"formula":"=TRUE","formulaResult":true},"children":[]},
                    {"type":"Cell","props":{"formula":"=NA()","formulaResult":"#N/A"},"children":[]},
                    {"type":"Cell","props":{"formula":"=A1","formulaResult":null},"children":[]}
                ]
            })],
        )]);

        input
            .validate()
            .expect("formula results support scalar and empty cached values");
    }

    #[test]
    fn rejects_formula_result_without_formula_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Formula results",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"formulaResult":2},"children":[]}
                ]
            })],
        )]);

        let error = input
            .validate()
            .expect_err("formulaResult requires a formula");
        assert!(
            error
                .to_string()
                .contains("requires formula; add Cell.formula")
        );
        assert!(error.to_string().contains("add Cell.formula"));
    }

    #[test]
    fn rejects_array_formula_result_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Array",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"formula":"=A1:A2*2","arrayFormulaRange":"A1:A2","formulaResult":4},"children":[]}
                ]
            })],
        )]);

        let error = input
            .validate()
            .expect_err("array formula caches are outside the scalar result contract");
        assert!(
            error
                .to_string()
                .contains("arrayFormulaRange with hyperlink/formulaResult"),
            "{error}"
        );
        assert!(
            error
                .to_string()
                .contains("use only the array formula range")
        );
    }

    #[test]
    fn rejects_ambiguous_text_formula_result_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Formula results",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"formula":"=A1","formulaResult":"123"},"children":[]}
                ]
            })],
        )]);

        let error = input
            .validate()
            .expect_err("numeric-looking text would be serialized as a number");
        assert!(error.to_string().contains("looks numeric"));
        assert!(error.to_string().contains("use a number"));
    }

    #[test]
    fn rejects_boolean_looking_text_formula_result_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Formula results",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"formula":"=A1","formulaResult":"TRUE"},"children":[]}
                ]
            })],
        )]);

        let error = input
            .validate()
            .expect_err("boolean-looking text is serialized as a boolean");
        assert!(error.to_string().contains("looks numeric or boolean"));
    }

    #[test]
    fn accepts_external_hyperlink_with_text_label() {
        let input = workbook(vec![worksheet(
            "Links",
            vec![serde_json::json!({
                "type":"Row",
                "props":{},
                "children":[{"type":"Cell","props":{"hyperlink":"https://example.com"},"children":["Open"]}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_hyperlink_tooltip_with_link() {
        let input = workbook(vec![worksheet(
            "Links",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"hyperlink":"https://example.com","hyperlinkTooltip":"Open report"},"children":["Open"]}
                ]
            })],
        )]);

        input
            .validate()
            .expect("a hyperlink tooltip with a target is valid");
    }

    #[test]
    fn rejects_hyperlink_tooltip_without_link_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Links",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"hyperlinkTooltip":"Open report"},"children":["Open"]}
                ]
            })],
        )]);

        let error = input
            .validate()
            .expect_err("a tooltip requires a hyperlink");
        assert!(error.to_string().contains("requires hyperlink"));
        assert!(error.to_string().contains("add Cell.hyperlink"));
    }

    #[test]
    fn rejects_non_string_hyperlink_tooltip_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Links",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"hyperlink":"https://example.com","hyperlinkTooltip":12},"children":["Open"]}
                ]
            })],
        )]);

        let error = input
            .validate()
            .expect_err("a hyperlink tooltip must be text");
        assert!(error.to_string().contains("must be a string"));
    }

    #[test]
    fn rejects_invalid_xml_character_in_hyperlink_tooltip() {
        let input = workbook(vec![worksheet(
            "Links",
            vec![serde_json::json!({
                "type":"Row","props":{},"children":[
                    {"type":"Cell","props":{"hyperlink":"https://example.com","hyperlinkTooltip":"bad\u{1}text"},"children":["Open"]}
                ]
            })],
        )]);

        let error = input
            .validate()
            .expect_err("XML control characters are invalid");
        assert!(error.to_string().contains("invalid XML character"));
    }

    #[test]
    fn rejects_hyperlink_on_non_text_cell_with_repair() {
        let input = workbook(vec![worksheet(
            "Links",
            vec![serde_json::json!({
                "type":"Row",
                "props":{},
                "children":[{"type":"Cell","props":{"hyperlink":"https://example.com","value":42},"children":[]}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("hyperlinks require a visible string label");
        assert!(error.to_string().contains("requires string text"));
        assert!(error.to_string().contains("use value or a string child"));
    }

    #[test]
    fn accepts_internal_hyperlink_target_with_visible_text() {
        let input = workbook(vec![worksheet(
            "Links",
            vec![serde_json::json!({
                "type":"Row",
                "props":{},
                "children":[{"type":"Cell","props":{"hyperlink":"internal:Sheet1!A1"},"children":["Jump"]}]
            })],
        )]);
        input
            .validate()
            .expect("internal hyperlinks retain their anchor in the Cell property");
    }

    #[test]
    fn rejects_empty_internal_hyperlink_location_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Links",
            vec![serde_json::json!({
                "type":"Row",
                "props":{},
                "children":[{"type":"Cell","props":{"hyperlink":"internal:"},"children":["Jump"]}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("internal targets require a location after the prefix");
        assert!(
            error
                .to_string()
                .contains("must be a non-empty `internal:`")
        );
    }

    #[test]
    fn accepts_list_data_validation_component() {
        let input = workbook(vec![worksheet(
            "Validation",
            vec![serde_json::json!({
                "type":"DataValidation",
                "props":{"range":"B2:B20","type":"list","source":"\"Open,Closed\""},
                "children":[]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_whole_number_between_validation() {
        let input = workbook(vec![worksheet(
            "Validation",
            vec![serde_json::json!({
                "type":"DataValidation",
                "props":{"range":"B2:B20","type":"whole","operator":"between","formula1":"1","formula2":"10"},
                "children":[]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_two_color_scale_conditional_format() {
        let input = workbook(vec![worksheet(
            "Heatmap",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"colorScale2","minColor":"#FFF2CC","maxColor":"#70AD47"},
                "children":[]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_greater_than_when_icon_set_is_x14_only() {
        let input = workbook(vec![worksheet(
            "Icons",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{
                    "range":"A1:A3",
                    "type":"iconSet",
                    "iconType":"threeArrows",
                    "icons":[
                        {"type":"percent","value":0},
                        {"type":"percentile","value":50,"iconType":"fourHistograms","iconIndex":2,"greaterThan":true},
                        {"type":"formula","value":"MAX(A1:A3)"}
                    ]
                },
                "children":[]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("x14 cfvo output omits greaterThan");
        assert!(
            error
                .to_string()
                .contains("does not serialize greaterThan in x14 cfvo")
        );
    }

    #[test]
    fn rejects_top_percent_rank_over_one_hundred_with_repair() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"top10","rule":"topPercent","rank":101},
                "children":[]
            })],
        )]);
        let error = input.validate().expect_err("percent rank is capped at 100");
        assert!(error.to_string().contains("from 1 through 100"));
    }

    #[test]
    fn rejects_top_bottom_properties_on_other_conditional_formats() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"cellIs","operator":"greaterThan","formula1":"10","rank":10},
                "children":[]
            })],
        )]);
        let error = input.validate().expect_err("rank belongs only to top10");
        assert!(error.to_string().contains("only valid with type top10"));
    }

    #[test]
    fn rejects_date_rule_outside_the_lossless_writer_subset() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"date","rule":"thisQuarter"},
                "children":[]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("the pinned writer does not expose this date period");
        assert!(error.to_string().contains("is invalid for type `date`"));
    }

    #[test]
    fn rejects_average_rule_with_data_bar_properties() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"average","fillColor":"#00FF00","barOnly":true},
                "children":[]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("data bar properties cannot be reinterpreted as a differential fill");
        assert!(error.to_string().contains("remove `barOnly`"));
    }

    #[test]
    fn accepts_cell_comparison_conditional_format_with_differential_style() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"cellIs","operator":"between","formula1":"10","formula2":"20","bold":true,"fontColor":"#9C0006","fillColor":"#FFC7CE"},
                "children":[]
            })],
        )]);
        input
            .validate()
            .expect("cell comparisons and differential styles are representable");
    }

    #[test]
    fn accepts_conditional_pattern_fill_as_a_differential_style() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"cellIs","operator":"equal","formula1":"10","fillPattern":"darkTrellis","fillForegroundColor":"#FF0000","fillBackgroundColor":"#FFFF00"},
                "children":[]
            })],
        )]);
        input
            .validate()
            .expect("patterned dxf fill is representable");
    }

    #[test]
    fn rejects_invalid_conditional_pattern_fills_with_repair() {
        for fill in [
            serde_json::json!({"fillColor":"#112233","fillPattern":"solid"}),
            serde_json::json!({"fillForegroundColor":"#112233"}),
            serde_json::json!({"fillPattern":"none","fillBackgroundColor":"#112233"}),
            serde_json::json!({"fillPattern":"unknown"}),
        ] {
            let mut props = serde_json::Map::from_iter([
                ("range".into(), serde_json::json!("A1:A10")),
                ("type".into(), serde_json::json!("cellIs")),
                ("operator".into(), serde_json::json!("equal")),
                ("formula1".into(), serde_json::json!("1")),
            ]);
            props.extend(fill.as_object().expect("fill props object").clone());
            let input = workbook(vec![worksheet(
                "Rules",
                vec![serde_json::json!({"type":"ConditionalFormat","props":props,"children":[]})],
            )]);
            let error = input
                .validate()
                .expect_err("invalid conditional fill is rejected");
            assert!(error.to_string().contains("fill"), "{error}");
            assert!(
                error.to_string().contains("Choose") || error.to_string().contains("must be one"),
                "{error}"
            );
        }
    }

    #[test]
    fn accepts_conditional_font_decoration_styles() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"formula","formula1":"B2>0","bold":true,"strikethrough":true,"underline":"double","script":"superscript"},
                "children":[]
            })],
        )]);
        input
            .validate()
            .expect("supported dxf font styles validate");
    }

    #[test]
    fn accepts_conditional_number_format_and_edge_border_styles() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"formula","formula1":"B2>0","numberFormat":"#,##0.00","borderTop":"thin","borderTopColor":"#123456"},
                "children":[]
            })],
        )]);
        input
            .validate()
            .expect("supported differential number format and border validate");
    }

    #[test]
    fn rejects_conditional_border_color_without_matching_edge() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"formula","formula1":"B2>0","borderTopColor":"#123456"},
                "children":[]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("border color requires its matching edge style");
        assert!(
            error
                .to_string()
                .contains("requires a matching non-`none` borderTop")
        );
    }

    #[test]
    fn rejects_conditional_unsupported_border_with_repair() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"formula","formula1":"B2>0","borderTop":"wavy"},
                "children":[]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("unsupported dxf border styles cannot be emitted");
        assert!(error.to_string().contains("supported edge border style"));
    }

    #[test]
    fn rejects_conditional_script_only_style_with_repair() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"formula","formula1":"B2>0","script":"subscript"},
                "children":[]
            })],
        )]);
        let error = input.validate().expect_err("writer drops script-only dxf");
        assert!(
            error
                .to_string()
                .contains("requires another differential-style property")
        );
    }

    #[test]
    fn rejects_cell_comparison_without_differential_style() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"cellIs","operator":"greaterThan","formula1":"10"},
                "children":[]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("a classic comparison without a style has no visible effect");
        assert!(
            error
                .to_string()
                .contains("requires at least one differential style")
        );
    }

    #[test]
    fn rejects_second_formula_for_single_bound_cell_comparison() {
        let input = workbook(vec![worksheet(
            "Rules",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"cellIs","operator":"greaterThan","formula1":"10","formula2":"20","bold":true},
                "children":[]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("single-bound operators must not carry formula2");
        assert!(
            error
                .to_string()
                .contains("formula2 is only valid for between/notBetween")
        );
    }

    #[test]
    fn rejects_percentile_threshold_outside_range() {
        let input = workbook(vec![worksheet(
            "Heatmap",
            vec![serde_json::json!({
                "type":"ConditionalFormat",
                "props":{"range":"B2:B20","type":"colorScale2","minColor":"#FFF2CC","maxColor":"#70AD47","minType":"percentile","minValue":101},
                "children":[]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("percentile thresholds are limited to 0 through 100");
        assert!(error.to_string().contains("minValue"));
        assert!(error.to_string().contains("0–100"));
    }

    #[test]
    fn rejects_between_validation_without_formula2_with_repair() {
        let input = workbook(vec![worksheet(
            "Validation",
            vec![serde_json::json!({
                "type":"DataValidation",
                "props":{"range":"B2:B20","type":"decimal","operator":"between","formula1":"1"},
                "children":[]
            })],
        )]);
        let error = input.validate().expect_err("between requires both bounds");
        assert!(error.to_string().contains("requires formula2"));
        assert!(error.to_string().contains("add the upper bound"));
    }

    #[test]
    fn rejects_data_validation_without_source_with_repair() {
        let input = workbook(vec![worksheet(
            "Validation",
            vec![serde_json::json!({
                "type":"DataValidation",
                "props":{"range":"B2:B20","type":"list"},
                "children":[]
            })],
        )]);
        let error = input.validate().expect_err("source formula is required");
        assert!(error.to_string().contains("DataValidation.source"));
        assert!(error.to_string().contains("list source formula"));
    }

    #[test]
    fn rejects_overlapping_data_validation_ranges_with_repair() {
        let input = workbook(vec![worksheet(
            "Validation",
            vec![
                serde_json::json!({
                    "type":"DataValidation",
                    "props":{"range":"A1:B2","type":"list","source":"\"A,B\""},
                    "children":[]
                }),
                serde_json::json!({
                    "type":"DataValidation",
                    "props":{"range":"B2:C3","type":"list","source":"\"C,D\""},
                    "children":[]
                }),
            ],
        )]);
        let error = input
            .validate()
            .expect_err("validation ranges cannot overlap");
        assert!(error.to_string().contains("overlaps another validation"));
        assert!(error.to_string().contains("combine the ranges"));
    }

    #[test]
    fn accepts_merged_cell_spans_and_skips_occupied_coordinates() {
        let input = workbook(vec![worksheet(
            "Merged",
            vec![
                serde_json::json!({"type":"Row","children":[
                    {"type":"Cell","props":{"rowSpan":2,"columnSpan":2},"children":["title"]},
                    {"type":"Cell","children":["next"]}
                ]}),
                serde_json::json!({"type":"Row","children":[
                    {"type":"Cell","children":["below"]}
                ]}),
            ],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_trailing_row_represented_by_a_vertical_merge() {
        let input = workbook(vec![worksheet(
            "Merged",
            vec![
                serde_json::json!({"type":"Row","children":[
                    {"type":"Cell","props":{"rowSpan":2},"children":["merged"]}
                ]}),
                serde_json::json!({"type":"Row","children":[]}),
            ],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_zero_cell_span_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Merged",
            vec![serde_json::json!({"type":"Row","children":[
                {"type":"Cell","props":{"columnSpan":0},"children":["title"]}
            ]})],
        )]);
        let error = input.validate().expect_err("span count must be positive");
        assert!(error.to_string().contains("Cell.columnSpan"));
        assert!(error.to_string().contains("must be at least 1"));
        assert!(error.to_string().contains("use `columnSpan={1}`"));
    }

    #[test]
    fn accepts_worksheet_autofilter_range() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Filter","autoFilter":"$A$1:$D$20"},
            "children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_string_autofilter_column_criteria_within_range() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Filter","autoFilter":"B1:D20","autoFilterColumns":[
                {"column":0,"values":["East","West"]},
                {"column":2,"values":["Open"]}
            ]},
            "children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_two_custom_autofilter_conditions_with_explicit_logic() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Filter","autoFilter":"A1:C20","autoFilterColumns":[
                {"column":1,"customFilters":{"logic":"and","conditions":[
                    {"operator":"greaterThanOrEqual","value":"10"},
                    {"operator":"lessThan","value":"20"}
                ]}}
            ]},
            "children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_list_filter_that_selects_blanks() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Filter","autoFilter":"A1:B20","autoFilterColumns":[
                {"column":0,"includeBlanks":true}
            ]},
            "children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_more_than_two_custom_autofilter_conditions_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Filter","autoFilter":"A1:B20","autoFilterColumns":[
                {"column":0,"customFilters":{"logic":"or","conditions":[
                    {"value":"a"},{"value":"b"},{"value":"c"}
                ]}}
            ]},
            "children":[]
        })]);
        let error = input
            .validate()
            .expect_err("Excel supports at most two custom conditions");
        assert!(error.to_string().contains("at most two custom conditions"));
        assert!(error.to_string().contains("remove extra conditions"));
    }

    #[test]
    fn rejects_autofilter_criteria_without_range_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Filter","autoFilterColumns":[{"column":0,"values":["East"]}]},
            "children":[]
        })]);
        let error = input
            .validate()
            .expect_err("criteria require an autofilter range");
        assert!(error.to_string().contains("requires Worksheet.autoFilter"));
        assert!(error.to_string().contains("add autoFilter"));
    }

    #[test]
    fn rejects_duplicate_autofilter_column_criteria_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Filter","autoFilter":"A1:C10","autoFilterColumns":[
                {"column":0,"values":["East"]},{"column":0,"values":["West"]}
            ]},
            "children":[]
        })]);
        let error = input.validate().expect_err("each filter column is unique");
        assert!(error.to_string().contains("column 0 is repeated"));
        assert!(error.to_string().contains("combine criteria"));
    }

    #[test]
    fn rejects_autofilter_column_outside_range_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Filter","autoFilter":"B1:C10","autoFilterColumns":[
                {"column":2,"values":["East"]}
            ]},
            "children":[]
        })]);
        let error = input
            .validate()
            .expect_err("column index is beyond filter width");
        assert!(error.to_string().contains("outside autoFilter width 2"));
        assert!(error.to_string().contains("use an index from 0 to 1"));
    }

    #[test]
    fn accepts_empty_string_as_an_exact_autofilter_value() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Filter","autoFilter":"A1:B10","autoFilterColumns":[
                {"column":0,"values":[""]}
            ]},
            "children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_freeze_pane_properties_on_worksheet() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Frozen","freezeRows":1,"freezeColumns":2,"freezeTopLeftCell":"E4"},
            "children":[]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_freeze_origin_inside_frozen_area_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Frozen","freezeRows":1,"freezeColumns":2,"freezeTopLeftCell":"B1"},
            "children":[]
        })]);
        let error = input
            .validate()
            .expect_err("scroll origin cannot be inside frozen rows or columns");
        assert!(
            error
                .to_string()
                .contains("must be outside the frozen 1 rows and 2 columns")
        );
    }

    #[test]
    fn rejects_empty_freeze_splits_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Frozen","freezeRows":0,"freezeColumns":0},
            "children":[]
        })]);
        let error = input
            .validate()
            .expect_err("a configured freeze must split at least one axis");
        assert!(
            error
                .to_string()
                .contains("must freeze at least one row or column")
        );
    }

    #[test]
    fn parses_and_formats_a1_ranges_with_absolute_markers() {
        let parsed = parse_a1_range("$b$2:d10");
        assert_eq!(parsed, Some((1, 1, 9, 3)));
        assert_eq!(format_a1_range(parsed.unwrap()), "B2:D10");
    }

    #[test]
    fn parses_and_formats_cell_row_and_column_print_ranges() {
        for (source, canonical) in [("$A$1:$D$20", "A1:D20"), ("$3:$6", "3:6"), ("$c:$e", "C:E")] {
            let range = parse_print_range(source).expect("print range parses");
            assert_eq!(format_print_range(range), canonical);
        }
    }

    #[test]
    fn rejects_invalid_worksheet_autofilter_range_with_repair() {
        for reference in ["A0:B2", "B2:A1", "XFE1:XFE2", "A1048577:B1048578"] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet",
                "props":{"name":"Filter","autoFilter":reference},
                "children":[]
            })]);
            let error = input
                .validate()
                .expect_err("range is outside valid A1 bounds");
            assert!(error.to_string().contains("Worksheet.autoFilter"));
        }
    }

    #[test]
    fn accepts_losslessly_representable_row_height_and_hidden_properties() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![serde_json::json!({
                "type":"Row",
                "props":{"height":22.5,"hidden":true},
                "children":[{"type":"Cell","props":{},"children":["row"]}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_expanded_row_outline_levels() {
        let input = workbook(vec![worksheet(
            "Outline",
            vec![serde_json::json!({
                "type":"Row",
                "props":{"outlineLevel":2},
                "children":[{"type":"Cell","props":{},"children":["detail"]}]
            })],
        )]);

        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_a_named_rectangular_table_before_rows() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Sales"},
            "children":[
                {"type":"Table","props":{"range":"A1:B3","name":"SalesTable","style":"TableStyleMedium4","autoFilter":false},"children":[]},
                {"type":"Row","props":{},"children":[
                    {"type":"Cell","props":{},"children":["Product"]},
                    {"type":"Cell","props":{},"children":["Amount"]}
                ]},
                {"type":"Row","props":{},"children":[
                    {"type":"Cell","props":{},"children":["Tea"]},
                    {"type":"Cell","props":{},"children":[12]}
                ]},
                {"type":"Row","props":{},"children":[
                    {"type":"Cell","props":{},"children":["Coffee"]},
                    {"type":"Cell","props":{},"children":[18]}
                ]}
            ]
        })]);

        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn accepts_table_accessibility_text() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:B3","name":"SalesTable","altText":"Accessible summary","altTextTitle":"Sales title"},"children":[]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Product"]},{"type":"Cell","props":{},"children":["Amount"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]},{"type":"Cell","props":{},"children":[12]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Coffee"]},{"type":"Cell","props":{},"children":[18]}]}
            ]
        })]);

        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn accepts_table_column_superscript_and_subscript_formats() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:B3","name":"SalesTable"},"children":[
                    {"type":"TableColumn","props":{"format":{"bold":true,"script":"superscript"}},"children":[]},
                    {"type":"TableColumn","props":{"format":{"bold":true,"script":"subscript"}},"children":[]}
                ]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["High"]},{"type":"Cell","props":{},"children":["Low"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]},{"type":"Cell","props":{},"children":["Coffee"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Water"]},{"type":"Cell","props":{},"children":["Juice"]}]}
            ]
        })]);

        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn accepts_table_column_font_metadata_formats() {
        let table = |column| {
            serde_json::json!({
                "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                    {"type":"Table","props":{"range":"A1:A2","name":"SalesTable"},"children":[column]},
                    {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Amount"]}]},
                    {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":[12]}]}
                ]
            })
        };
        let input = workbook(vec![table(serde_json::json!({
            "type":"TableColumn",
            "props":{"headerFormat":{"fontName":"Aptos","fontSize":12.5,"fontFamily":3,"fontCharset":128}},
            "children":[]
        }))]);

        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn rejects_invalid_table_column_font_metadata_with_repair_guidance() {
        let table = |column| {
            serde_json::json!({
                "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                    {"type":"Table","props":{"range":"A1:A2","name":"SalesTable"},"children":[column]},
                    {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Amount"]}]},
                    {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":[12]}]}
                ]
            })
        };
        for (property, value, expected) in [
            (
                "fontName",
                serde_json::json!(""),
                "1 to 255 valid XML characters",
            ),
            ("fontSize", serde_json::json!(410), "1 through 409 points"),
            ("fontFamily", serde_json::json!(6), "0 through 5"),
            (
                "fontCharset",
                serde_json::json!(256),
                "integer byte from 0 through 255",
            ),
        ] {
            let input = workbook(vec![table(serde_json::json!({
                "type":"TableColumn",
                "props":{"headerFormat":{property:value}},
                "children":[]
            }))]);
            let error = input
                .validate()
                .expect_err("out-of-range font metadata must be rejected");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn rejects_table_data_font_metadata_that_writer_drops_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:A2","name":"SalesTable"},"children":[
                    {"type":"TableColumn","props":{"format":{"fontName":"Aptos"}},"children":[]}
                ]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Amount"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":[12]}]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("pinned writer omits font metadata from data differential styles");

        assert!(
            error
                .to_string()
                .contains("not serialized by the pinned writer")
        );
        assert!(
            error
                .to_string()
                .contains("move this font metadata to headerFormat")
        );
    }

    #[test]
    fn rejects_unsupported_table_column_font_script_with_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:A2","name":"SalesTable"},"children":[
                    {"type":"TableColumn","props":{"format":{"script":"baseline"}},"children":[]}
                ]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Value"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]}]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("baseline is not an authored dxf script value");

        assert!(error.to_string().contains("superscript` or `subscript"));
    }

    #[test]
    fn rejects_script_only_table_differential_format_that_writer_drops() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:A2","name":"SalesTable"},"children":[
                    {"type":"TableColumn","props":{"format":{"script":"superscript"}},"children":[]}
                ]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Value"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]}]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("the pinned writer omits script-only differential formats");

        assert!(
            error
                .to_string()
                .contains("script requires another TableColumn")
        );
    }

    #[test]
    fn rejects_conflicting_table_column_fill_shorthand_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:A2","name":"SalesTable"},"children":[
                    {"type":"TableColumn","props":{"format":{"fillColor":"#112233","fillPattern":"solid"}},"children":[]}
                ]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Value"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]}]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("TableColumn fill shorthand cannot combine with pattern fills");
        assert!(
            error
                .to_string()
                .contains("TableColumn.format fill properties conflict")
        );
        assert!(error.to_string().contains("Choose one fill representation"));
    }

    #[test]
    fn accepts_script_only_header_format_because_it_uses_cell_style_xml() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:A2","name":"SalesTable"},"children":[
                    {"type":"TableColumn","props":{"headerFormat":{"script":"superscript"}},"children":[]}
                ]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Value"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]}]}
            ]
        })]);

        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn rejects_non_string_table_accessibility_text_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:B3","name":"SalesTable","altText":42},"children":[]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Product"]},{"type":"Cell","props":{},"children":["Amount"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]},{"type":"Cell","props":{},"children":[12]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Coffee"]},{"type":"Cell","props":{},"children":[18]}]}
            ]
        })]);
        let error = input.validate().expect_err("altText must be a string");

        assert!(error.to_string().contains("Table.altText must be a string"));
        assert!(error.to_string().contains("omit the property"));
    }

    #[test]
    fn validates_table_column_count_and_mutually_exclusive_totals() {
        let table = |children| {
            serde_json::json!({
                "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                    {"type":"Table","props":{"range":"A1:B3","name":"SalesTable","totalsRow":true},"children":children},
                    {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Product"]},{"type":"Cell","props":{},"children":["Amount"]}]},
                    {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]},{"type":"Cell","props":{},"children":[12]}]}
                ]
            })
        };
        let mismatch = workbook(vec![table(vec![
            serde_json::json!({"type":"TableColumn","props":{},"children":[]}),
        ])]);
        let error = mismatch
            .validate()
            .expect_err("column count mismatch is invalid");
        assert!(error.to_string().contains("one TableColumn per column"));
        let conflict = workbook(vec![table(vec![
            serde_json::json!({"type":"TableColumn","props":{"totalFormula":"SUM([Amount])","totalFunction":"sum"},"children":[]}),
            serde_json::json!({"type":"TableColumn","props":{},"children":[]}),
        ])]);
        let error = conflict
            .validate()
            .expect_err("custom and built-in total functions conflict");
        assert!(error.to_string().contains("mutually exclusive"));
        let label_conflict = workbook(vec![table(vec![
            serde_json::json!({"type":"TableColumn","props":{"totalLabel":"Total","totalFunction":"sum"},"children":[]}),
            serde_json::json!({"type":"TableColumn","props":{},"children":[]}),
        ])]);
        let error = label_conflict
            .validate()
            .expect_err("a total label and calculation cannot coexist");
        assert!(error.to_string().contains("label or a calculation"));
        let unsupported_format = workbook(vec![table(vec![
            serde_json::json!({"type":"TableColumn","props":{"format":{"horizontalAlign":"left"}},"children":[]}),
            serde_json::json!({"type":"TableColumn","props":{},"children":[]}),
        ])]);
        let error = unsupported_format
            .validate()
            .expect_err("table dxf formats only support the declared style subset");
        assert!(error.to_string().contains("is not supported"));
        let empty_number_format = workbook(vec![table(vec![
            serde_json::json!({"type":"TableColumn","props":{"format":{"numberFormat":""}},"children":[]}),
            serde_json::json!({"type":"TableColumn","props":{},"children":[]}),
        ])]);
        let error = empty_number_format
            .validate()
            .expect_err("empty differential number formats are invalid");
        assert!(error.to_string().contains("1 to 255 characters"));
        let unsupported_underline = workbook(vec![table(vec![
            serde_json::json!({"type":"TableColumn","props":{"format":{"underline":"dotted"}},"children":[]}),
            serde_json::json!({"type":"TableColumn","props":{},"children":[]}),
        ])]);
        let error = unsupported_underline
            .validate()
            .expect_err("unsupported table underline variants cannot be compiled");
        assert!(error.to_string().contains("is not supported"));
        let border_color_without_style = workbook(vec![table(vec![
            serde_json::json!({"type":"TableColumn","props":{"format":{"borderTopColor":"#123456"}},"children":[]}),
            serde_json::json!({"type":"TableColumn","props":{},"children":[]}),
        ])]);
        let error = border_color_without_style
            .validate()
            .expect_err("a table dxf border color requires its edge style");
        assert!(
            error
                .to_string()
                .contains("requires a matching non-`none` borderTop")
        );
        let diagonal_style = workbook(vec![table(vec![
            serde_json::json!({"type":"TableColumn","props":{"format":{"borderDiagonal":"thin"}},"children":[]}),
            serde_json::json!({"type":"TableColumn","props":{},"children":[]}),
        ])]);
        let error = diagonal_style
            .validate()
            .expect_err("differential writer does not serialize diagonal line styles");
        assert!(
            error
                .to_string()
                .contains("not serializable by the pinned writer")
        );
    }

    #[test]
    fn totals_row_range_requires_a_data_row_before_the_total() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:B2","name":"SalesTable","totalsRow":true},"children":[]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Product"]},{"type":"Cell","props":{},"children":["Amount"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Total"]},{"type":"Cell","props":{},"children":[0]}]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("a totals row cannot replace all data rows");
        assert!(error.to_string().contains("at least one data row"));
    }

    #[test]
    fn rejects_a_table_without_a_data_row_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Sales"},
            "children":[
                {"type":"Table","props":{"range":"A1:B1","name":"SalesTable"},"children":[]},
                {"type":"Row","props":{},"children":[
                    {"type":"Cell","props":{},"children":["Product"]},
                    {"type":"Cell","props":{},"children":["Amount"]}
                ]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("tables require a header row and at least one data row");

        assert!(error.to_string().contains("Table.range"));
        assert!(error.to_string().contains("at least one data row"));
    }

    #[test]
    fn rejects_table_ranges_overlapping_worksheet_filters_with_repair_guidance() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Sales", "autoFilter":"A1:B3"},
            "children":[
                {"type":"Table","props":{"range":"A1:B3","name":"SalesTable"},"children":[]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Product"]},{"type":"Cell","props":{},"children":["Amount"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]},{"type":"Cell","props":{},"children":[12]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Coffee"]},{"type":"Cell","props":{},"children":[18]}]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("a table cannot overlap a worksheet autoFilter");
        assert!(error.to_string().contains("overlaps Worksheet.autoFilter"));
        assert!(
            error
                .to_string()
                .contains("remove the separate worksheet filter")
        );
    }

    #[test]
    fn rejects_case_insensitive_duplicate_table_headers() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet", "props":{"name":"Sales"}, "children":[
                {"type":"Table","props":{"range":"A1:B2","name":"SalesTable"},"children":[]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Name"]},{"type":"Cell","props":{},"children":["name"]}]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["Tea"]},{"type":"Cell","props":{},"children":["Tea"]}]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("table column names are case-insensitively unique");
        assert!(error.to_string().contains("duplicate header names"));
        assert!(error.to_string().contains("unique non-empty string"));
    }

    #[test]
    fn rejects_row_outline_level_above_excel_limit_with_repair() {
        let input = workbook(vec![worksheet(
            "Outline",
            vec![serde_json::json!({
                "type":"Row",
                "props":{"outlineLevel":8},
                "children":[{"type":"Cell","props":{},"children":["detail"]}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("outline levels are limited to 7");

        assert!(error.to_string().contains("Row.outlineLevel"));
        assert!(error.to_string().contains("1 through 7"));
    }

    #[test]
    fn accepts_disjoint_column_ranges_before_rows() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Columns"},
            "children":[
                {"type":"Column","props":{"min":1,"max":2,"widthPixels":160,"outlineLevel":2},"children":[]},
                {"type":"Column","props":{"min":4,"max":4,"hidden":true},"children":[]},
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["value"]}]}
            ]
        })]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_overlapping_column_ranges_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Columns"},
            "children":[
                {"type":"Column","props":{"min":1,"max":3},"children":[]},
                {"type":"Column","props":{"min":3,"max":4},"children":[]}
            ]
        })]);
        let error = input
            .validate()
            .expect_err("overlapping columns are ambiguous");
        assert!(error.to_string().contains("overlaps another Column"));
    }

    #[test]
    fn rejects_column_width_outside_integer_pixel_range_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Columns"},
            "children":[
                {"type":"Column","props":{"min":1,"max":1,"widthPixels":1791},"children":[]}
            ]
        })]);
        let error = input.validate().expect_err("width exceeds writer's limit");
        assert!(error.to_string().contains("widthPixels"));
        assert!(error.to_string().contains("1 to 1790"));
    }

    #[test]
    fn rejects_column_declarations_after_rows_with_repair() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Columns"},
            "children":[
                {"type":"Row","props":{},"children":[{"type":"Cell","props":{},"children":["value"]}]},
                {"type":"Column","props":{"min":1,"max":1},"children":[]}
            ]
        })]);
        let error = input.validate().expect_err("columns must precede rows");
        assert!(
            error
                .to_string()
                .contains("move all Column declarations before rows")
        );
    }

    #[test]
    fn rejects_unrepresentable_row_height_with_repair() {
        for height in [0.0, 22.3, 410.25] {
            let input = workbook(vec![worksheet(
                "Data",
                vec![serde_json::json!({
                    "type":"Row",
                    "props":{"height":height},
                    "children":[{"type":"Cell","props":{},"children":["row"]}]
                })],
            )]);
            let error = input.validate().expect_err("invalid row height");
            assert!(error.to_string().contains("0.75-point step"));
            assert!(error.to_string().contains("409.5"));
        }
    }

    #[test]
    fn rejects_non_boolean_row_hidden_property_with_repair() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![serde_json::json!({
                "type":"Row",
                "props":{"hidden":"yes"},
                "children":[{"type":"Cell","props":{},"children":["row"]}]
            })],
        )]);
        let error = input.validate().expect_err("hidden must be boolean");
        assert!(error.to_string().contains("Row.hidden must be a boolean"));
        assert!(error.to_string().contains("hidden={true}"));
    }

    #[test]
    fn rejects_case_insensitive_duplicate_sheet_names_with_repair() {
        let input = workbook(vec![worksheet("Data", vec![]), worksheet("data", vec![])]);
        let error = input
            .validate()
            .expect_err("duplicate sheet names must fail");
        assert!(error.to_string().contains("duplicates another sheet"));
        assert!(error.to_string().contains("choose a unique name"));
    }

    #[test]
    fn rejects_formula_and_value_conflict_with_repair() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![
                serde_json::json!({"type":"Row","children":[{"type":"Cell","props":{"formula":"=1+1","value":2}}]}),
            ],
        )]);
        let error = input.validate().expect_err("formula and value conflict");
        assert!(error.to_string().contains("cannot combine formula"));
        assert!(error.to_string().contains("use one representation"));
    }

    #[test]
    fn rejects_value_property_and_child_conflict_with_repair() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![
                serde_json::json!({"type":"Row","children":[{"type":"Cell","props":{"value":2},"children":["different"]}]}),
            ],
        )]);
        let error = input.validate().expect_err("value and child conflict");
        assert!(
            error
                .to_string()
                .contains("cannot combine the value property with children")
        );
        assert!(error.to_string().contains("use one representation"));
    }

    #[test]
    fn accepts_positional_empty_rows_and_rejects_unobservable_trailing_rows() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![
                serde_json::json!({"type":"Row","children":[]}),
                serde_json::json!({"type":"Row","children":[{"type":"Cell","props":{},"children":["later"]}]}),
            ],
        )]);
        assert!(input.validate().is_ok());

        let input = workbook(vec![worksheet(
            "Data",
            vec![serde_json::json!({"type":"Row","children":[]})],
        )]);
        let error = input
            .validate()
            .expect_err("trailing empty row has no representation");
        assert!(error.to_string().contains("trailing empty Row"));
        assert!(error.to_string().contains("add a Cell or row metadata"));
    }

    #[test]
    fn rejects_unicode_case_insensitive_duplicate_sheet_names() {
        let input = workbook(vec![
            worksheet("Straße", vec![]),
            worksheet("STRAẞE", vec![]),
        ]);
        let error = input
            .validate()
            .expect_err("Unicode case variants cannot duplicate sheet names");
        assert!(error.to_string().contains("duplicates another sheet"));
    }

    #[test]
    fn accepts_boolean_font_emphasis_properties() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![
                serde_json::json!({"type":"Row","children":[{"type":"Cell","props":{"bold":true,"italic":false},"children":["emphasis"]}]}),
            ],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_checkbox_only_for_boolean_cell_values() {
        let input = workbook(vec![worksheet(
            "Checklist",
            vec![serde_json::json!({"type":"Row","children":[
                {"type":"Cell","props":{"value":true,"checkbox":true},"children":[]}
            ]})],
        )]);
        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn rejects_checkbox_on_non_boolean_cell_values_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Checklist",
            vec![serde_json::json!({"type":"Row","children":[
                {"type":"Cell","props":{"value":"yes","checkbox":true},"children":[]}
            ]})],
        )]);
        let error = input
            .validate()
            .expect_err("checkbox requires boolean value");
        assert!(error.to_string().contains("Cell.checkbox"));
        assert!(error.to_string().contains("boolean value"));
        assert!(error.to_string().contains("value={true}"));
    }

    #[test]
    fn accepts_horizontal_and_vertical_alignment_properties() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![serde_json::json!({
                "type":"Row",
                "children":[
                    {"type":"Cell","props":{"horizontalAlign":"center","verticalAlign":"top"},"children":["aligned"]}
                ]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_supported_cell_reading_directions() {
        let input = workbook(vec![worksheet(
            "Direction",
            vec![serde_json::json!({"type":"Row","children":[
                {"type":"Cell","props":{"readingDirection":"leftToRight"},"children":["ltr"]},
                {"type":"Cell","props":{"readingDirection":"rightToLeft"},"children":["rtl"]}
            ]})],
        )]);
        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn accepts_boolean_quote_prefix_property() {
        let input = workbook(vec![worksheet(
            "QuotePrefix",
            vec![serde_json::json!({"type":"Row","children":[
                {"type":"Cell","props":{"quotePrefix":true},"children":["=not a formula"]}
            ]})],
        )]);
        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn accepts_passwordless_worksheet_protection_options() {
        let input = workbook(vec![serde_json::json!({
            "type":"Worksheet",
            "props":{"name":"Protected","protection":{"formatCells":true,"sort":true,"selectLockedCells":false}},
            "children":[]
        })]);
        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn rejects_invalid_worksheet_protection_options_with_guidance() {
        for protection in [
            serde_json::json!("yes"),
            serde_json::json!({"formatCells":"yes"}),
            serde_json::json!({"unknownOption":true}),
        ] {
            let input = workbook(vec![serde_json::json!({
                "type":"Worksheet","props":{"name":"Protected","protection":protection},"children":[]
            })]);
            let error = input.validate().expect_err("invalid protection options");
            assert!(error.to_string().contains("protection"), "{error}");
        }
    }

    #[test]
    fn accepts_cell_lock_and_formula_visibility_flags() {
        let input = workbook(vec![worksheet(
            "Protected",
            vec![serde_json::json!({"type":"Row","children":[
                {"type":"Cell","props":{"locked":false},"children":["editable"]},
                {"type":"Cell","props":{"formula":"=1+1","formulaHidden":true},"children":[]}
            ]})],
        )]);
        assert!(input.validate().is_ok(), "{:?}", input.validate());
    }

    #[test]
    fn rejects_non_boolean_quote_prefix_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "QuotePrefix",
            vec![serde_json::json!({"type":"Row","children":[
                {"type":"Cell","props":{"quotePrefix":"yes"},"children":["=not a formula"]}
            ]})],
        )]);
        let error = input.validate().expect_err("quotePrefix must be boolean");
        assert!(error.to_string().contains("quotePrefix"));
        assert!(error.to_string().contains("boolean"));
    }

    #[test]
    fn rejects_unknown_cell_reading_direction_with_repair_guidance() {
        let input = workbook(vec![worksheet(
            "Direction",
            vec![serde_json::json!({"type":"Row","children":[
                {"type":"Cell","props":{"readingDirection":"contextual"},"children":["bad"]}
            ]})],
        )]);
        let error = input.validate().expect_err("unknown reading direction");
        assert!(error.to_string().contains("readingDirection"));
        assert!(error.to_string().contains("leftToRight"));
    }

    #[test]
    fn validates_text_rotation_range_and_type() {
        let valid = workbook(vec![worksheet(
            "Rotation",
            vec![serde_json::json!({"type":"Row","children":[
                {"type":"Cell","props":{"textRotation":-45},"children":["negative"]},
                {"type":"Cell","props":{"textRotation":270},"children":["stacked"]}
            ]})],
        )]);
        assert!(valid.validate().is_ok());

        for value in [
            serde_json::json!(91),
            serde_json::json!(-91),
            serde_json::json!(255),
            serde_json::json!("45"),
        ] {
            let invalid = workbook(vec![worksheet(
                "Rotation",
                vec![
                    serde_json::json!({"type":"Row","children":[{"type":"Cell","props":{"textRotation":value},"children":["bad"]}]}),
                ],
            )]);
            let error = invalid
                .validate()
                .expect_err("invalid text rotation must fail");
            assert!(error.to_string().contains("textRotation"));
        }
    }

    #[test]
    fn validates_rich_text_runs_and_rejects_conflicting_cell_values() {
        let valid = workbook(vec![worksheet(
            "RichText",
            vec![
                serde_json::json!({"type":"Row","children":[{"type":"Cell","children":[
                {"type":"TextRun","props":{},"children":["Status: "]},
                {"type":"TextRun","props":{"bold":true,"fontColor":"#008000"},"children":["Ready"]}
            ]}]} ),
            ],
        )]);
        assert!(valid.validate().is_ok());

        let conflicting = workbook(vec![worksheet(
            "RichText",
            vec![
                serde_json::json!({"type":"Row","children":[{"type":"Cell","props":{"value":"ignored"},"children":[
                {"type":"TextRun","props":{},"children":["first"]},
                {"type":"TextRun","props":{},"children":["second"]}
            ]}]} ),
            ],
        )]);
        let error = conflicting
            .validate()
            .expect_err("rich runs cannot combine with value");
        assert!(error.to_string().contains("TextRun"));

        let invalid_run = workbook(vec![worksheet(
            "RichText",
            vec![
                serde_json::json!({"type":"Row","children":[{"type":"Cell","children":[
                {"type":"TextRun","props":{"fillColor":"#FFFFFF"},"children":["not a run font property"]}
            ]}]} ),
            ],
        )]);
        let error = invalid_run
            .validate()
            .expect_err("TextRun must reject cell-only properties");
        assert!(error.to_string().contains("TextRun"));

        let empty_run = workbook(vec![worksheet(
            "RichText",
            vec![
                serde_json::json!({"type":"Row","children":[{"type":"Cell","children":[
                {"type":"TextRun","props":{},"children":[""]}
            ]}]} ),
            ],
        )]);
        let error = empty_run
            .validate()
            .expect_err("TextRun text cannot be empty");
        assert!(error.to_string().contains("non-empty string child"));
    }

    #[test]
    fn accepts_indent_with_supported_horizontal_alignments() {
        let input = workbook(vec![worksheet(
            "Indented",
            vec![serde_json::json!({
                "type":"Row",
                "children":[
                    {"type":"Cell","props":{"value":"left","horizontalAlign":"left","indent":1}},
                    {"type":"Cell","props":{"value":"right","horizontalAlign":"right","indent":255}},
                    {"type":"Cell","props":{"value":"distributed","horizontalAlign":"distributed","indent":2}}
                ]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_indent_without_explicit_supported_horizontal_alignment() {
        for props in [
            serde_json::json!({"value":"text","indent":1}),
            serde_json::json!({"value":"text","horizontalAlign":"center","indent":1}),
        ] {
            let input = workbook(vec![worksheet(
                "Indented",
                vec![serde_json::json!({
                    "type":"Row", "children":[{"type":"Cell","props":props}]
                })],
            )]);
            let error = input
                .validate()
                .expect_err("indent must use a supported explicit horizontal alignment");
            assert!(
                error.to_string().contains("left, right, distributed")
                    || error.to_string().contains("horizontalAlign")
            );
        }
    }

    #[test]
    fn rejects_shrink_to_fit_with_distributed_alignment() {
        let input = workbook(vec![worksheet(
            "Indented",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":"text","horizontalAlign":"distributed","shrinkToFit":true}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("the writer suppresses shrink-to-fit with distributed alignment");
        assert!(error.to_string().contains("cannot be combined"));
    }

    #[test]
    fn rejects_indent_outside_u8_range() {
        let input = workbook(vec![worksheet(
            "Indented",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":"text","horizontalAlign":"left","indent":256}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("indent must fit writer's u8 API");
        assert!(error.to_string().contains("1 through 255"));
    }

    #[test]
    fn accepts_text_wrapping_with_alignment() {
        let input = workbook(vec![worksheet(
            "Wrapped",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"wrapped","wrapText":true,"horizontalAlign":"center"}}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_non_boolean_text_wrapping_with_repair() {
        let input = workbook(vec![worksheet(
            "Wrapped",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":"wrapped","wrapText":"yes"}}]
            })],
        )]);
        let error = input.validate().expect_err("wrapText must be boolean");
        assert!(error.to_string().contains("must be a boolean"));
    }

    #[test]
    fn accepts_shrink_to_fit_with_horizontal_alignment() {
        let input = workbook(vec![worksheet(
            "Shrink",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"content","shrinkToFit":true,"horizontalAlign":"right"}}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_combined_text_wrapping_and_shrink_to_fit_with_repair() {
        let input = workbook(vec![worksheet(
            "Shrink",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":"content","shrinkToFit":true,"wrapText":true}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("the writer drops shrink-to-fit when text wrapping is enabled");
        assert!(error.to_string().contains("cannot both be true"));
    }

    #[test]
    fn rejects_non_boolean_shrink_to_fit_with_repair() {
        let input = workbook(vec![worksheet(
            "Shrink",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":"content","shrinkToFit":"yes"}}]
            })],
        )]);
        let error = input.validate().expect_err("shrinkToFit must be boolean");
        assert!(error.to_string().contains("must be a boolean"));
    }

    #[test]
    fn accepts_number_format_codes_with_scalar_formula_and_date_cells() {
        let input = workbook(vec![worksheet(
            "Formats",
            vec![serde_json::json!({
                "type":"Row",
                "children":[
                    {"type":"Cell","props":{"value":1234.5,"numberFormat":"#,##0.00"}},
                    {"type":"Cell","props":{"formula":"=A1*2","numberFormat":"$#,##0.00"}},
                    {"type":"Cell","props":{"date":"2026-10-10","numberFormat":"dd/mm/yyyy"}}
                ]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_rgb_fill_color_with_a_number_format() {
        let input = workbook(vec![worksheet(
            "Fill",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":12,"fillColor":"#a1B2c3","numberFormat":"#,##0"}}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_rgb_font_color_with_other_styles() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":"text","fontColor":"#a1B2c3","bold":true,"fillColor":"#FFFFFF"}}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_invalid_font_color_with_repair() {
        for color in ["red", "#12345", "#12345678", "#12GG56"] {
            let input = workbook(vec![worksheet(
                "Font",
                vec![serde_json::json!({
                    "type":"Row", "children":[{"type":"Cell","props":{"value":1,"fontColor":color}}]
                })],
            )]);
            let error = input.validate().expect_err("font color must be opaque RGB");
            assert!(error.to_string().contains("fontColor"));
            assert!(error.to_string().contains("must match `#RRGGBB`"));
        }
    }

    #[test]
    fn accepts_literal_font_name_and_fractional_font_size() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"text","fontName":"Aptos","fontSize":12.5}}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_font_family_and_charset_metadata() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"text","fontFamily":3,"fontCharset":128}}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_supported_font_schemes_and_rejects_unknown_values() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"text","fontScheme":"headings"}}]
            })],
        )]);
        assert!(input.validate().is_ok());

        let invalid = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"text","fontScheme":"display"}}]
            })],
        )]);
        let error = invalid
            .validate()
            .expect_err("unknown font scheme is rejected");
        assert!(error.to_string().contains("fontScheme"));
        assert!(error.to_string().contains("none`, `body`, or `headings"));
    }

    #[test]
    fn rejects_font_family_outside_the_ooxml_enum() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"text","fontFamily":6}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("family 6 is not a supported OOXML family");
        assert!(error.to_string().contains("fontFamily"));
        assert!(error.to_string().contains("0 through 5"));
    }

    #[test]
    fn rejects_font_charset_outside_byte_range() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"text","fontCharset":256}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("charset must fit in an OOXML byte");
        assert!(error.to_string().contains("fontCharset"));
        assert!(error.to_string().contains("0 through 255"));
    }

    #[test]
    fn accepts_font_underline_and_strikethrough() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"text","underline":"singleAccounting","strikethrough":true}}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_superscript_and_subscript_font_scripts() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row",
                "children":[
                    {"type":"Cell","props":{"value":"x2","script":"subscript"}},
                    {"type":"Cell","props":{"value":"x2","script":"superscript"}}
                ]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_unsupported_font_script_with_repair() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":1,"script":"baseline"}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("unsupported scripts are invalid");
        assert!(error.to_string().contains("superscript, subscript"));
    }

    #[test]
    fn rejects_invalid_font_underline_with_repair() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":1,"underline":"wavy"}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("unknown underline styles are invalid");
        assert!(error.to_string().contains("singleAccounting"));
    }

    #[test]
    fn rejects_non_boolean_strikethrough_with_repair() {
        let input = workbook(vec![worksheet(
            "Font",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":1,"strikethrough":"yes"}}]
            })],
        )]);
        let error = input.validate().expect_err("strikethrough must be boolean");
        assert!(error.to_string().contains("must be a boolean"));
    }

    #[test]
    fn rejects_invalid_font_names_with_repair() {
        for name in ["", "bad\u{1}font", &"x".repeat(256)] {
            let input = workbook(vec![worksheet(
                "Font",
                vec![serde_json::json!({
                    "type":"Row", "children":[{"type":"Cell","props":{"value":1,"fontName":name}}]
                })],
            )]);
            let error = input
                .validate()
                .expect_err("font name must be valid XML text");
            assert!(error.to_string().contains("fontName"));
        }
    }

    #[test]
    fn rejects_font_sizes_outside_supported_range_or_non_numeric_values() {
        for size in [
            serde_json::json!(0.5),
            serde_json::json!(409.5),
            serde_json::json!("12"),
        ] {
            let input = workbook(vec![worksheet(
                "Font",
                vec![serde_json::json!({
                    "type":"Row", "children":[{"type":"Cell","props":{"value":1,"fontSize":size}}]
                })],
            )]);
            let error = input
                .validate()
                .expect_err("font size must be in points from 1 to 409");
            assert!(error.to_string().contains("fontSize"));
        }
    }

    #[test]
    fn accepts_independent_cell_borders_and_rgb_colors() {
        let input = workbook(vec![worksheet(
            "Borders",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"bordered","borderTop":"thin","borderTopColor":"#123456","borderRight":"dotted","borderRightColor":"#abcdef","borderBottom":"dashed","borderLeft":"double"}}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn accepts_diagonal_border_style_color_and_direction() {
        let input = workbook(vec![worksheet(
            "Borders",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"diagonal","borderDiagonal":"thin","borderDiagonalColor":"#123456","borderDiagonalType":"upDown"}}]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_diagonal_color_without_matching_style() {
        let input = workbook(vec![worksheet(
            "Borders",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"diagonal","borderDiagonalColor":"#123456"}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("diagonal color requires a style");
        assert!(error.to_string().contains("borderDiagonalColor"));
        assert!(
            error
                .to_string()
                .contains("matching non-`none` borderDiagonal")
        );
    }

    #[test]
    fn rejects_unsupported_diagonal_direction_with_repair() {
        let input = workbook(vec![worksheet(
            "Borders",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":"diagonal","borderDiagonalType":"sideways"}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("direction must map to one of the writer enums");
        assert!(error.to_string().contains("borderDiagonalType"));
        assert!(error.to_string().contains("`up`, `down`, or `upDown`"));
    }

    #[test]
    fn rejects_border_colors_without_edge_style_with_repair() {
        let input = workbook(vec![worksheet(
            "Borders",
            vec![serde_json::json!({
                "type":"Row", "children":[{"type":"Cell","props":{"value":1,"borderTopColor":"#123456"}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("edge color requires a corresponding edge style");
        assert!(error.to_string().contains("borderTopColor"));
    }

    #[test]
    fn rejects_unsupported_border_style_or_invalid_color_with_repair() {
        for (property, value) in [("borderLeft", "wavy"), ("borderBottomColor", "red")] {
            let input = workbook(vec![worksheet(
                "Borders",
                vec![serde_json::json!({
                    "type":"Row", "children":[{"type":"Cell","props":{"value":1,"borderBottom":"thin",(property):value}}]
                })],
            )]);
            let error = input
                .validate()
                .expect_err("border style or RGB color is outside supported subset");
            assert!(error.to_string().contains(property));
        }
    }

    #[test]
    fn rejects_invalid_fill_color_with_repair() {
        for fill_color in ["red", "#12345", "#12345678", "#12GG56"] {
            let input = workbook(vec![worksheet(
                "Fill",
                vec![serde_json::json!({
                    "type":"Row",
                    "children":[{"type":"Cell","props":{"value":1,"fillColor":fill_color}}]
                })],
            )]);
            let error = input
                .validate()
                .expect_err("only six-digit RGB fill colors are valid");
            assert!(error.to_string().contains("fillColor"));
        }
    }

    #[test]
    fn accepts_pattern_fill_on_cells_rows_and_columns() {
        let input = workbook(vec![worksheet(
            "Patterns",
            vec![
                serde_json::json!({"type":"Column","props":{"min":1,"max":1,"fillPattern":"darkDown","fillForegroundColor":"#FF0000","fillBackgroundColor":"#FFFF00"}}),
                serde_json::json!({"type":"Row","props":{"fillPattern":"gray125","fillForegroundColor":"#112233"},"children":[{"type":"Cell","props":{"value":"pattern","fillPattern":"darkDown","fillForegroundColor":"#FF0000","fillBackgroundColor":"#FFFF00"}}]}),
            ],
        )]);
        input.validate().expect("pattern fills are valid");
    }

    #[test]
    fn rejects_invalid_pattern_fill_combinations_with_repair() {
        for props in [
            serde_json::json!({"fillColor":"#112233","fillPattern":"solid"}),
            serde_json::json!({"fillForegroundColor":"#112233"}),
            serde_json::json!({"fillPattern":"none","fillBackgroundColor":"#112233"}),
            serde_json::json!({"fillPattern":"badPattern"}),
        ] {
            let mut cell_props =
                serde_json::Map::from_iter([("value".into(), serde_json::json!(1))]);
            cell_props.extend(props.as_object().expect("props object").clone());
            let input = workbook(vec![worksheet(
                "Patterns",
                vec![
                    serde_json::json!({"type":"Row","children":[{"type":"Cell","props":cell_props}]}),
                ],
            )]);
            let error = input
                .validate()
                .expect_err("invalid pattern fill is rejected");
            assert!(error.to_string().contains("fill"), "{error}");
            assert!(
                error.to_string().contains("use")
                    || error.to_string().contains("requires")
                    || error.to_string().contains("Choose"),
                "{error}"
            );
        }
    }

    #[test]
    fn rejects_empty_or_xml_invalid_number_format_with_repair() {
        for format in ["", "bad\u{1}format"] {
            let input = workbook(vec![worksheet(
                "Formats",
                vec![serde_json::json!({
                    "type":"Row",
                    "children":[{"type":"Cell","props":{"value":1,"numberFormat":format}}]
                })],
            )]);
            let error = input
                .validate()
                .expect_err("invalid Excel number format must be rejected");
            assert!(error.to_string().contains("numberFormat"));
            if format.is_empty() {
                assert!(error.to_string().contains("1 to 255 characters"));
            } else {
                assert!(error.to_string().contains("invalid in XML"), "{error}");
            }
        }
    }

    #[test]
    fn rejects_overlong_number_format_with_repair() {
        let format = "0".repeat(256);
        let input = workbook(vec![worksheet(
            "Formats",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"value":1,"numberFormat":format}}]
            })],
        )]);
        let error = input
            .validate()
            .expect_err("number format exceeds Excel's limit");
        assert!(error.to_string().contains("1 to 255 characters"));
    }

    #[test]
    fn rejects_alignment_values_outside_the_lossless_subset() {
        for property in ["horizontalAlign", "verticalAlign"] {
            let input = workbook(vec![worksheet(
                "Data",
                vec![serde_json::json!({
                    "type":"Row",
                    "children":[{"type":"Cell","props":{(property):"justify"},"children":["aligned"]}]
                })],
            )]);
            let error = input
                .validate()
                .expect_err("alignment must be from the documented enum");
            assert!(error.to_string().contains("must be one of"));
        }
    }

    #[test]
    fn rejects_non_boolean_font_emphasis_with_repair() {
        let input = workbook(vec![worksheet(
            "Data",
            vec![
                serde_json::json!({"type":"Row","children":[{"type":"Cell","props":{"bold":"yes"},"children":["emphasis"]}]}),
            ],
        )]);
        let error = input.validate().expect_err("font emphasis is boolean");
        assert!(error.to_string().contains("must be a boolean"));
        assert!(error.to_string().contains("bold={true}"));
    }

    #[test]
    fn accepts_iso_dates_and_local_datetimes() {
        let input = workbook(vec![worksheet(
            "Dates",
            vec![serde_json::json!({
                "type":"Row",
                "children":[
                    {"type":"Cell","props":{"date":"2026-10-10"}},
                    {"type":"Cell","props":{"date":"2026-10-10T14:30:12.125"}}
                ]
            })],
        )]);
        assert!(input.validate().is_ok());
    }

    #[test]
    fn rejects_invalid_or_timezone_bearing_dates_with_repair() {
        for date in [
            "2026-02-30",
            "2026-10-10T14:30:12Z",
            "2026-10-10T14:30:12+08:00",
        ] {
            let input = workbook(vec![worksheet(
                "Dates",
                vec![serde_json::json!({
                    "type":"Row",
                    "children":[{"type":"Cell","props":{"date":date}}]
                })],
            )]);
            let error = input.validate().expect_err("date value is not lossless");
            assert!(error.to_string().contains("local ISO date/time"));
            assert!(error.to_string().contains("remove the time-zone suffix"));
        }
    }

    #[test]
    fn rejects_date_combined_with_a_scalar_value() {
        let input = workbook(vec![worksheet(
            "Dates",
            vec![serde_json::json!({
                "type":"Row",
                "children":[{"type":"Cell","props":{"date":"2026-10-10","value":"not a date"}}]
            })],
        )]);
        let error = input.validate().expect_err("date and scalar are exclusive");
        assert!(error.to_string().contains("cannot combine date"));
        assert!(error.to_string().contains("one representation"));
    }
}
