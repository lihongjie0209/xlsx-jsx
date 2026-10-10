use assert_cmd::Command;
use calamine::Reader;
use predicates::prelude::*;
use std::io::{Read, Write};
use tempfile::tempdir;
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

#[test]
fn spec_is_agent_readable_and_validate_does_not_write_output() {
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["spec", "--format", "json"]);
    command
        .assert()
        .success()
        .stdout(predicate::str::contains("xlsx-jsx IR v1"))
        .stdout(predicate::str::contains("fillColor"))
        .stdout(predicate::str::contains("fillPattern"))
        .stdout(predicate::str::contains("fillForegroundColor"))
        .stdout(predicate::str::contains("fillBackgroundColor"))
        .stdout(predicate::str::contains("hyperlinkTooltip"))
        .stdout(predicate::str::contains("formulaResult"))
        .stdout(predicate::str::contains("fontColor"))
        .stdout(predicate::str::contains("fontName"))
        .stdout(predicate::str::contains("fontSize"))
        .stdout(predicate::str::contains("fontFamily"))
        .stdout(predicate::str::contains("altText"))
        .stdout(predicate::str::contains("\"const\": \"Note\""))
        .stdout(predicate::str::contains("underline"))
        .stdout(predicate::str::contains("strikethrough"))
        .stdout(predicate::str::contains("script"))
        .stdout(predicate::str::contains("wrapText"))
        .stdout(predicate::str::contains("shrinkToFit"))
        .stdout(predicate::str::contains("rowSpan"))
        .stdout(predicate::str::contains("DataValidation"))
        .stdout(predicate::str::contains("autoFilterColumns"))
        .stdout(predicate::str::contains("customFilters"))
        .stdout(predicate::str::contains("arrayFormulaRange"))
        .stdout(predicate::str::contains("dynamicArray"))
        .stdout(predicate::str::contains("hideUnusedRows"))
        .stdout(predicate::str::contains("showGridlines"))
        .stdout(predicate::str::contains("printGridlines"))
        .stdout(predicate::str::contains("horizontalPageBreaks"))
        .stdout(predicate::str::contains("verticalPageBreaks"))
        .stdout(predicate::str::contains("pageMargins"))
        .stdout(predicate::str::contains("printScale"))
        .stdout(predicate::str::contains("fitToPages"))
        .stdout(predicate::str::contains("orientation"))
        .stdout(predicate::str::contains("paperSize"))
        .stdout(predicate::str::contains("centerHorizontally"))
        .stdout(predicate::str::contains("printBlackAndWhite"))
        .stdout(predicate::str::contains("printHeadings"))
        .stdout(predicate::str::contains("pageOrder"))
        .stdout(predicate::str::contains("printArea"))
        .stdout(predicate::str::contains("repeatRows"))
        .stdout(predicate::str::contains("quotePrefix"))
        .stdout(predicate::str::contains("protection"))
        .stdout(predicate::str::contains("formulaHidden"))
        .stdout(predicate::str::contains("headerFooterScaleWithDoc"))
        .stdout(predicate::str::contains("headerFooterAlignWithPage"))
        .stdout(predicate::str::contains("repeatColumns"))
        .stdout(predicate::str::contains("zoom"))
        .stdout(predicate::str::contains("includeBlanks"))
        .stdout(predicate::str::contains("outlineLevel"))
        .stdout(predicate::str::contains("visibility"))
        .stdout(predicate::str::contains("collapsed"))
        .stdout(predicate::str::contains("Table"))
        .stdout(predicate::str::contains("dataBar"))
        .stdout(predicate::str::contains("negativeFillColor"))
        .stdout(predicate::str::contains("axisPosition"))
        .stdout(predicate::str::contains("aboveAverage"))
        .stdout(predicate::str::contains("last7Days"))
        .stdout(predicate::str::contains("nonBlank"))
        .stdout(predicate::str::contains("numberFormat"))
        .stdout(predicate::str::contains("borderTop"))
        .stdout(predicate::str::contains("borderDiagonalType"))
        .stdout(predicate::str::contains("altTextTitle"))
        .stdout(predicate::str::contains("unique"))
        .stdout(predicate::str::contains("checkbox"));
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["spec", "--format", "json"]);
    command
        .assert()
        .success()
        .stdout(predicate::str::contains("DefinedName"));
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["spec", "--format", "json"]);
    command
        .assert()
        .success()
        .stdout(predicate::str::contains("TextRun"));
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["spec", "--format", "json"]);
    command
        .assert()
        .success()
        .stdout(predicate::str::contains("description"));
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["spec", "--format", "json"]);
    command
        .assert()
        .success()
        .stdout(predicate::str::contains("topLeftCell"))
        .stdout(predicate::str::contains("containsText"))
        .stdout(predicate::str::contains("selection"))
        .stdout(predicate::str::contains("activeCell"))
        .stdout(predicate::str::contains("firstTab"));

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["validate", "examples/basic.tsx"]);
    command.assert().success().stdout("valid\n");
}

#[test]
fn spec_exposes_chart_components_and_properties() {
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["spec", "--format", "json"])
        .assert()
        .success()
        .stdout(predicate::str::contains("\"const\": \"Chart\""))
        .stdout(predicate::str::contains("chartSeries"))
        .stdout(predicate::str::contains("categories"))
        .stdout(predicate::str::contains("\"title\""))
        .stdout(predicate::str::contains("xAxisTitle"))
        .stdout(predicate::str::contains("yAxisTitle"))
        .stdout(predicate::str::contains("legendPosition"))
        .stdout(predicate::str::contains("showLegend"));
}

#[test]
fn spec_exposes_embedded_image_component_properties() {
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["spec", "--format", "json"])
        .assert()
        .success()
        .stdout(predicate::str::contains("\"const\": \"Image\""))
        .stdout(predicate::str::contains("data:image/(png|jpeg|gif|bmp)"))
        .stdout(predicate::str::contains("decorative"))
        .stdout(predicate::str::contains("offsetX"))
        .stdout(predicate::str::contains("moveAndSize"));
}

#[test]
fn validates_image_data_uri_and_rejects_invalid_data_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-image.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Image } from "xlsx-jsx";
export default <Workbook><Worksheet name="Images"><Image cell="A1" width={1} height={1} data="data:image/png;base64,not_base64!" /></Worksheet></Workbook>;"#,
    )
    .expect("source writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("Image.data"))
        .stderr(predicate::str::contains("not valid base64"))
        .stderr(predicate::str::contains("standard base64"));
}

#[test]
fn validate_rejects_truncated_image_payload_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("truncated-image.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Image } from "xlsx-jsx";
export default <Workbook><Worksheet name="Images"><Image cell="A1" width={1} height={1} data="data:image/png;base64,iVBORw0KGgo=" /></Worksheet></Workbook>;"#,
    )
    .expect("source writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(&input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("not a valid `png` image"))
        .stderr(predicate::str::contains("complete supported image file"));
}

#[test]
fn validates_image_offsets_inside_the_default_cell() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("out-of-cell-image-offset.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Image } from "xlsx-jsx";
export default <Workbook><Worksheet name="Images"><Image cell="A1" width={1} height={1} offsetX={64} data="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg==" /></Worksheet></Workbook>;"#,
    )
    .expect("source writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("Image.offsetX"))
        .stderr(predicate::str::contains("0 through 63 pixels"));
}

#[test]
#[expect(
    clippy::too_many_lines,
    reason = "This end-to-end test checks media bytes and picture/chart semantics across reverse, rebuild, and a second reverse."
)]
fn builds_reverses_and_rebuilds_embedded_image_without_manifest() {
    const IMAGE_DATA_URI: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNgAAAAAgABSK+kcQAAAABJRU5ErkJggg==";
    let directory = tempdir().expect("temporary directory is available");
    let input = std::path::Path::new("examples/images.tsx");
    let output = directory.path().join("image.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(input)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
    let mut archive = ZipArchive::new(std::fs::File::open(&output).expect("output exists"))
        .expect("xlsx is a ZIP package");
    assert!(
        archive
            .file_names()
            .any(|name| name == "xl/media/image1.png")
    );
    let reversed = directory.path().join("image-reversed.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reversed JSX exists");
    assert!(source.contains(IMAGE_DATA_URI), "{source}");
    assert!(
        source.contains("altText={\"Tiny sample logo\"}"),
        "{source}"
    );
    assert!(source.contains("decorative={true}"), "{source}");
    assert!(source.contains("offsetX={7}"), "{source}");
    assert!(source.contains("offsetY={3}"), "{source}");
    assert!(source.contains("movement=\"fixed\""), "{source}");
    let first_image = source
        .find("<Image cell={\"A1\"}")
        .expect("first image is rendered");
    let chart = source
        .find("<Chart type=\"column\"")
        .expect("chart is rendered");
    let second_image = source
        .find("<Image cell={\"M1\"}")
        .expect("second image is rendered");
    assert!(
        first_image < second_image && second_image < chart,
        "{source}"
    );
    let rebuilt = directory.path().join("image-rebuilt.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let mut rebuilt_archive =
        ZipArchive::new(std::fs::File::open(&rebuilt).expect("rebuilt workbook exists"))
            .expect("rebuilt XLSX is a ZIP package");
    assert!(
        rebuilt_archive
            .file_names()
            .any(|name| name == "xl/media/image1.png")
    );
    let mut original_image = Vec::new();
    archive
        .by_name("xl/media/image1.png")
        .expect("source image exists")
        .read_to_end(&mut original_image)
        .expect("source image reads");
    let mut rebuilt_image = Vec::new();
    rebuilt_archive
        .by_name("xl/media/image1.png")
        .expect("rebuilt image exists")
        .read_to_end(&mut rebuilt_image)
        .expect("rebuilt image reads");
    assert_eq!(original_image, rebuilt_image);
    let twice_reversed = directory.path().join("image-roundtrip.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&twice_reversed)
        .assert()
        .success();
    let second_source =
        std::fs::read_to_string(twice_reversed).expect("second reverse output exists");
    assert!(
        second_source.contains("<Image cell={\"A1\"}"),
        "{second_source}"
    );
    assert!(
        second_source.contains("<Image cell={\"M1\"}"),
        "{second_source}"
    );
    assert!(
        second_source.contains("<Chart type=\"column\""),
        "{second_source}"
    );
    assert!(second_source.contains("offsetX={7}"), "{second_source}");
    assert!(
        second_source.contains("movement=\"fixed\""),
        "{second_source}"
    );
    assert!(second_source.contains(IMAGE_DATA_URI), "{second_source}");
}

#[test]
fn validates_chart_series_ranges_and_actionable_errors() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-chart.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Chart, Series } from "xlsx-jsx";
export default <Workbook><Worksheet name="Data"><Chart type="column" cell="D2"><Series values="B2:B4" /></Chart></Worksheet></Workbook>;"#,
    ).expect("source writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(&input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("sheet-qualified A1 range"))
        .stderr(predicate::str::contains("Data!$B$2:$B$5"));
}

#[test]
fn validate_rejects_chart_geometry_crossing_custom_column_widths() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("chart-custom-width.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Column, Row, Cell, Chart, Series } from "xlsx-jsx";
export default <Workbook><Worksheet name="Data"><Chart type="column" cell="D2"><Series values="Data!$B$2:$B$3" /></Chart><Column min={8} max={8} widthPixels={120} /><Row><Cell>Label</Cell><Cell>Value</Cell></Row><Row><Cell>A</Cell><Cell value={2} /></Row><Row><Cell>B</Cell><Cell value={3} /></Row></Worksheet></Workbook>;"#,
    )
    .expect("source writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(&input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("standard 64×20-pixel cells"))
        .stderr(predicate::str::contains(
            "remove the intersecting custom widths/heights",
        ));
}

#[test]
fn validate_reports_chart_references_to_missing_worksheets() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("chart-missing-sheet.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Chart, Series } from "xlsx-jsx";
export default <Workbook><Worksheet name="Data"><Chart type="column" cell="D2"><Series values="Missing!$B$2:$B$4" /></Chart></Worksheet></Workbook>;"#,
    )
    .expect("source writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(&input)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "refers to missing worksheet `Missing`",
        ))
        .stderr(predicate::str::contains(
            "add that sheet or correct the reference",
        ));
}

#[test]
fn builds_reverses_and_rebuilds_basic_chart_parts() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("charts.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build", "examples/charts.tsx", "-o"])
        .arg(&output)
        .assert()
        .success();
    let archive = ZipArchive::new(std::fs::File::open(&output).expect("output exists"))
        .expect("xlsx is a ZIP package");
    assert!(
        archive
            .file_names()
            .any(|name| name == "xl/charts/chart1.xml")
    );
    assert!(
        archive
            .file_names()
            .any(|name| name == "xl/drawings/drawing1.xml")
    );
    assert!(
        archive
            .file_names()
            .any(|name| name == "xl/charts/chart4.xml")
    );

    let reversed = directory.path().join("charts.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reversed JSX exists");
    assert!(
        source.contains("<Chart type=\"column\" cell={\"D2\"}"),
        "{source}"
    );
    assert!(source.contains("values={\"Data!$B$2:$B$4\"}"), "{source}");
    assert!(source.contains("values={\"Data!$C$2:$C$4\"}"), "{source}");
    assert!(source.contains("<Chart type=\"line\""), "{source}");
    assert!(source.contains("<Chart type=\"bar\""), "{source}");
    assert!(source.contains("<Chart type=\"pie\""), "{source}");
    assert!(
        source.contains("name={\"Revenue and expenses\"}"),
        "{source}"
    );
    assert_chart_title_round_trip(&source);
    assert!(
        source.contains("altText={\"Monthly revenue compared with expenses\"}"),
        "{source}"
    );

    let rebuilt = directory.path().join("charts-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let rebuilt_archive = ZipArchive::new(std::fs::File::open(&rebuilt).expect("rebuilt exists"))
        .expect("rebuilt XLSX is a ZIP package");
    assert!(
        rebuilt_archive
            .file_names()
            .any(|name| name == "xl/charts/chart1.xml")
    );
    assert!(
        rebuilt_archive
            .file_names()
            .any(|name| name == "xl/charts/chart4.xml")
    );
    drop(rebuilt_archive);
    let twice_reversed = directory.path().join("charts-roundtrip.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&twice_reversed)
        .assert()
        .success();
    let second_source =
        std::fs::read_to_string(twice_reversed).expect("second reverse output exists");
    assert!(
        second_source.contains("<Chart type=\"pie\""),
        "{second_source}"
    );
    assert!(
        second_source.contains("values={\"Data!$C$2:$C$4\"}"),
        "{second_source}"
    );
    assert_chart_title_round_trip(&second_source);
}

fn assert_chart_title_round_trip(source: &str) {
    assert!(
        source.contains("title={\"Monthly revenue & expenses — 2026\"}"),
        "{source}"
    );
    assert!(source.contains("title={\"Expenses & travel\"}"), "{source}");
    assert!(source.contains("xAxisTitle={\"Month\"}"), "{source}");
    assert!(source.contains("yAxisTitle={\"Revenue ($)\"}"), "{source}");
    assert!(source.contains("legendPosition=\"bottom\""), "{source}");
    assert!(source.contains("showLegend={false}"), "{source}");
}

#[test]
fn validates_chart_title_length_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("long-chart-title.tsx");
    let title = "x".repeat(256);
    std::fs::write(
        &input,
        format!(r#"import {{ Workbook, Worksheet, Chart, Series }} from "xlsx-jsx"; export default <Workbook><Worksheet name="Data"><Chart type="column" cell="D2" title="{title}"><Series values="A1:A2" /></Chart></Worksheet></Workbook>;"#),
    )
    .expect("invalid fixture is written");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("Chart.title"))
        .stderr(predicate::str::contains("shorten it"));
}

#[test]
fn rejects_axis_titles_on_pie_charts_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("pie-axis-title.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Chart, Series } from "xlsx-jsx"; export default <Workbook><Worksheet name="Data"><Chart type="pie" cell="D2" xAxisTitle="Category"><Series values="Data!A1:A2" /></Chart></Worksheet></Workbook>;"#,
    )
    .expect("invalid fixture is written");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("pie Chart has no X/Y axes"))
        .stderr(predicate::str::contains("remove xAxisTitle and yAxisTitle"));
}

#[test]
fn rejects_hidden_chart_legend_with_position() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("hidden-positioned-legend.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Chart, Series } from "xlsx-jsx"; export default <Workbook><Worksheet name="Data"><Chart type="line" cell="D2" showLegend={false} legendPosition="top"><Series values="Data!A1:A2" /></Chart></Worksheet></Workbook>;"#,
    )
    .expect("invalid fixture is written");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("cannot set legendPosition"))
        .stderr(predicate::str::contains("showLegend={false}"));
}

#[test]
fn validates_chart_legend_position_with_supported_values() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-legend-position.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Chart, Series } from "xlsx-jsx"; export default <Workbook><Worksheet name="Data"><Chart type="line" cell="D2" legendPosition="bottoms"><Series values="Data!A1:A2" /></Chart></Worksheet></Workbook>;"#,
    )
    .expect("invalid fixture is written");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "legendPosition `bottoms` is unsupported",
        ))
        .stderr(predicate::str::contains(
            "right, left, top, bottom, or topRight",
        ));
}

#[test]
fn rejects_reference_shaped_chart_titles_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("reference-shaped-chart-title.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Chart, Series } from "xlsx-jsx"; export default <Workbook><Worksheet name="Data"><Chart type="column" cell="D2" title="Data!$A$1"><Series values="Data!A1:A2" /></Chart></Worksheet></Workbook>;"#,
    )
    .expect("invalid fixture is written");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("Chart.title"))
        .stderr(predicate::str::contains("interprets as a formula"));
}

#[test]
fn rejects_conflicting_series_reference_and_literal_title() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("conflicting-series-title.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Chart, Series } from "xlsx-jsx"; export default <Workbook><Worksheet name="Data"><Chart type="column" cell="D2"><Series values="Data!A1:A2" name="Data!B1" title="Revenue" /></Chart></Worksheet></Workbook>;"#,
    )
    .expect("invalid fixture is written");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("cannot set both name and title"))
        .stderr(predicate::str::contains(
            "use name for a source cell reference",
        ));
}

#[test]
fn rejects_literal_series_title_ambiguous_with_a_cell_reference() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("ambiguous-series-title.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, Chart, Series } from "xlsx-jsx"; export default <Workbook><Worksheet name="Data"><Chart type="column" cell="D2"><Series values="Data!A1:A2" title="Data!$B$1" /></Chart></Worksheet></Workbook>;"#,
    )
    .expect("invalid fixture is written");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "looks like a sheet-qualified A1 reference",
        ))
        .stderr(predicate::str::contains("rephrase it as literal text"));
}

#[test]
fn reverse_rejects_empty_chart_titles_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let external = directory.path().join("external-chart.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build", "examples/charts.tsx", "-o"])
        .arg(&external)
        .assert()
        .success();
    rewrite_part(&external, "xl/charts/chart2.xml", |xml| {
        xml.replace("<c:plotArea>", "<c:title/><c:plotArea>")
    });
    let output = directory.path().join("external-chart.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(external)
        .args(["-o"])
        .arg(output)
        .assert()
        .failure()
        .stderr(predicate::str::contains("empty title"))
        .stderr(predicate::str::contains("remove the title element"));
}

#[test]
fn reverse_rejects_cached_series_titles_without_a_source_or_literal_title() {
    let directory = tempdir().expect("temporary directory is available");
    let external = directory.path().join("literal-chart-title.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build", "examples/charts.tsx", "-o"])
        .arg(&external)
        .assert()
        .success();
    rewrite_part(&external, "xl/charts/chart1.xml", |xml| {
        xml.replace("<c:f>Data!$B$1</c:f>", "")
    });
    let output = directory.path().join("literal-chart-title.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(external)
        .args(["-o"])
        .arg(output)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "series title without a source or literal value",
        ))
        .stderr(predicate::str::contains("Series.title as static text"));
}

#[test]
fn reverse_rejects_literal_series_titles_that_writer_would_reinterpret() {
    let directory = tempdir().expect("temporary directory is available");
    let external = directory
        .path()
        .join("ambiguous-external-series-title.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build", "examples/charts.tsx", "-o"])
        .arg(&external)
        .assert()
        .success();
    rewrite_part(&external, "xl/charts/chart1.xml", |xml| {
        xml.replace("<c:v>Expenses &amp; travel</c:v>", "<c:v>Data!$C$1</c:v>")
    });
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(external)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "writer would reinterpret as a cell reference",
        ))
        .stderr(predicate::str::contains("change the title text"));
}

#[test]
fn reverse_rejects_formatted_chart_titles_instead_of_dropping_formatting() {
    let directory = tempdir().expect("temporary directory is available");
    let external = directory.path().join("formatted-chart-title.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build", "examples/charts.tsx", "-o"])
        .arg(&external)
        .assert()
        .success();
    rewrite_part(&external, "xl/charts/chart1.xml", |xml| {
        xml.replace(
            "<a:rPr lang=\"en-US\"/>",
            "<a:rPr lang=\"en-US\" sz=\"1400\"/>",
        )
    });
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(external)
        .assert()
        .failure()
        .stderr(predicate::str::contains("unsupported `rPr@sz=1400`"))
        .stderr(predicate::str::contains(
            "remove custom chart/axis/print formatting",
        ));
}

#[test]
fn reverse_rejects_reference_shaped_axis_titles_writer_would_reinterpret() {
    let directory = tempdir().expect("temporary directory is available");
    let external = directory.path().join("reference-shaped-axis-title.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build", "examples/charts.tsx", "-o"])
        .arg(&external)
        .assert()
        .success();
    rewrite_part(&external, "xl/charts/chart1.xml", |xml| {
        xml.replace("<a:t>Month</a:t>", "<a:t>Data!$A$1</a:t>")
    });
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(external)
        .assert()
        .failure()
        .stderr(predicate::str::contains("xAxisTitle"))
        .stderr(predicate::str::contains(
            "would reinterpret as a cell reference",
        ));
}

#[test]
fn reverse_rejects_multi_run_chart_titles_instead_of_flattening_runs() {
    let directory = tempdir().expect("temporary directory is available");
    let external = directory.path().join("multi-run-chart-title.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build", "examples/charts.tsx", "-o"])
        .arg(&external)
        .assert()
        .success();
    rewrite_part(&external, "xl/charts/chart1.xml", |xml| {
        xml.replace(
            "</a:p>",
            "<a:r><a:rPr lang=\"en-US\"/><a:t>suffix</a:t></a:r></a:p>",
        )
    });
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(external)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "multi-paragraph or multi-run title",
        ))
        .stderr(predicate::str::contains("one unformatted title string"));
}

#[test]
fn external_global_and_local_defined_names_reverse_and_rebuild() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("external-names.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("Data")
        .expect("sheet name is valid");
    workbook
        .define_name("TaxRate", "=0.05")
        .expect("global name is valid");
    workbook
        .define_name("'Data'!LocalRate", "=Data!$A$1")
        .expect("local name is valid");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("defined-names.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(
        source.contains("<DefinedName name={\"TaxRate\"} formula={\"0.05\"} />"),
        "{source}"
    );
    assert!(
        source.contains("<DefinedName name={\"LocalRate\"} formula={\"Data!$A$1\"} />"),
        "{source}"
    );

    let rebuilt = directory.path().join("rebuilt.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let mut archive = ZipArchive::new(std::fs::File::open(rebuilt).expect("rebuilt file opens"))
        .expect("rebuilt workbook is a ZIP package");
    let mut workbook_xml = String::new();
    archive
        .by_name("xl/workbook.xml")
        .expect("workbook part exists")
        .read_to_string(&mut workbook_xml)
        .expect("workbook XML reads");
    assert!(workbook_xml.contains("name=\"TaxRate\""), "{workbook_xml}");
    assert!(
        workbook_xml.contains("name=\"LocalRate\""),
        "{workbook_xml}"
    );
    assert!(workbook_xml.contains("0.05"), "{workbook_xml}");
    assert!(workbook_xml.contains("Data!$A$1"), "{workbook_xml}");
}

#[test]
fn external_dynamic_array_formulas_reverse_and_rebuild() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("dynamic-arrays.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_dynamic_array_formula(0, 1, 2, 1, "=LEN(A1:A3)")
        .expect("bounded dynamic formula is valid");
    worksheet
        .write_dynamic_formula(0, 3, "=FILTER(A1:A3,A1:A3<>\"\")")
        .expect("single-cell dynamic formula is valid");
    worksheet.write_string(0, 0, "Foo").expect("cell writes");
    worksheet.write_string(1, 0, "Food").expect("cell writes");
    worksheet.write_string(2, 0, "Frood").expect("cell writes");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("dynamic-arrays.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("dynamicArray={true}"), "{source}");
    assert!(source.contains("arrayFormulaRange={\"B1:B3\"}"), "{source}");
    assert!(source.contains("FILTER(A1:A3"), "{source}");

    let rebuilt = directory.path().join("dynamic-arrays-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let mut archive = ZipArchive::new(std::fs::File::open(rebuilt).expect("rebuilt file opens"))
        .expect("rebuilt workbook is a ZIP package");
    assert!(archive.file_names().any(|name| name == "xl/metadata.xml"));
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet part exists")
        .read_to_string(&mut sheet_xml)
        .expect("sheet XML reads");
    assert!(sheet_xml.contains("cm=\"1\""), "{sheet_xml}");
    assert!(sheet_xml.contains("t=\"array\""), "{sheet_xml}");
}

#[test]
fn builds_and_reverses_a_supported_workbook() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("basic.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/basic.tsx", "-o"])
        .arg(&output);
    command.assert().success();
    assert!(output.is_file());

    let reversed = directory.path().join("roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("name={\"Sales \\\"Q\\\"\"}"));
    assert!(source.contains("Tea"));
    assert!(source.contains("formula={\"=IF(A2=\\\"Tea\\\",8,0)\"}"));

    let regenerated = directory.path().join("regenerated.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    assert!(regenerated.is_file());
}

#[test]
fn ignored_error_annotations_round_trip_external_workbooks() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("ignored-errors.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .ignore_error_range(0, 0, 2, 1, rust_xlsxwriter::IgnoreError::NumberStoredAsText)
        .expect("valid ignored-error range");
    worksheet
        .ignore_error(4, 3, rust_xlsxwriter::IgnoreError::FormulaError)
        .expect("valid single-cell annotation");
    worksheet
        .ignore_error(6, 0, rust_xlsxwriter::IgnoreError::NumberStoredAsText)
        .expect("valid second range for the same annotation");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("ignored-errors.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("<IgnoreError range=\"A1:B3\" errorType=\"numberStoredAsText\" />"));
    assert!(source.contains("<IgnoreError range=\"D5:D5\" errorType=\"evalError\" />"));
    assert!(source.contains("<IgnoreError range=\"A7:A7\" errorType=\"numberStoredAsText\" />"));

    let regenerated = directory.path().join("regenerated.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated)
        .assert()
        .success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(regenerated).expect("rebuilt workbook opens"))
            .expect("rebuilt workbook is zipped");
    let mut sheet = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet part exists")
        .read_to_string(&mut sheet)
        .expect("worksheet XML reads");
    assert!(sheet.contains("<ignoredError sqref=\"A1:B3 A7\" numberStoredAsText=\"1\""));
    assert!(sheet.contains("<ignoredError sqref=\"D5\" evalError=\"1\""));
}

#[test]
fn overlapping_ignored_error_ranges_fail_validation_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let source = directory.path().join("overlap.tsx");
    std::fs::write(&source, "import { Workbook, Worksheet, IgnoreError } from \"xlsx-jsx\"; export default <Workbook><Worksheet name=\"S\"><IgnoreError range=\"A1:B2\" errorType=\"numberStoredAsText\" /><IgnoreError range=\"B2:C3\" errorType=\"evalError\" /></Worksheet></Workbook>;").expect("source writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(&source)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "overlaps another ignored-error range",
        ))
        .stderr(predicate::str::contains("make ranges disjoint"));
}

#[test]
fn external_selection_and_scroll_origin_reverse_and_rebuild() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("selection.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_top_left_cell(5, 7)
        .expect("valid scroll origin");
    worksheet
        .set_selection(2, 3, 4, 5)
        .expect("valid selection");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("selection.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("topLeftCell=\"H6\""));
    assert!(source.contains("selection={{activeCell: \"D3\", range: \"D3:F5\"}}"));

    let regenerated = directory.path().join("selection-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated)
        .assert()
        .success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("round-trip opens"))
        .expect("round-trip workbook is zipped");
    let mut sheet = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet part exists")
        .read_to_string(&mut sheet)
        .expect("worksheet XML reads");
    assert!(sheet.contains("topLeftCell=\"H6\""));
    assert!(sheet.contains("<selection activeCell=\"D3\" sqref=\"D3:F5\""));
}

#[test]
fn invalid_selection_reports_a_repair_suggestion() {
    let directory = tempdir().expect("temporary directory is available");
    let source = directory.path().join("invalid-selection.tsx");
    std::fs::write(&source, "import { Workbook, Worksheet } from \"xlsx-jsx\"; export default <Workbook><Worksheet name=\"S\" selection={{activeCell: \"C3\", range: \"A1:B2\"}} /></Workbook>;").expect("source writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(&source)
        .assert()
        .failure()
        .stderr(predicate::str::contains("activeCell"))
        .stderr(predicate::str::contains("inside the selection range"));
}

#[test]
fn reverse_rejects_interior_active_cell_with_writer_limitation_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("interior-selection.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_selection(2, 3, 4, 5)
        .expect("selection is configured");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("activeCell=\"D3\"", "activeCell=\"E4\"")
    });

    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("not at a range corner"))
        .stderr(predicate::str::contains("move it to a corner"));
}

#[test]
fn reverses_and_rebuilds_authored_core_properties() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("metadata.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let properties = rust_xlsxwriter::DocProperties::new()
        .set_title("Quarterly report")
        .set_subject("Revenue")
        .set_author("Analyst")
        .set_keywords("quarterly,finance")
        .set_comment("Reviewed & approved");
    workbook.set_properties(&properties);
    workbook
        .add_worksheet()
        .set_name("Data")
        .expect("valid name");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("metadata.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    for property in [
        "title=\"Quarterly report\"",
        "subject=\"Revenue\"",
        "author=\"Analyst\"",
        "keywords=\"quarterly,finance\"",
        "description=\"Reviewed & approved\"",
    ] {
        assert!(source.contains(property), "missing {property} in {source}");
    }

    let regenerated = directory.path().join("metadata-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(regenerated).expect("generated workbook opens"))
            .expect("generated workbook is a ZIP package");
    let mut core = String::new();
    archive
        .by_name("docProps/core.xml")
        .expect("core properties are present")
        .read_to_string(&mut core)
        .expect("core properties are valid UTF-8");
    for property in [
        "Quarterly report",
        "Revenue",
        "Analyst",
        "quarterly,finance",
        "Reviewed &amp; approved",
    ] {
        assert!(core.contains(property), "missing `{property}` in {core}");
    }
}

#[test]
fn row_and_column_default_formats_reverse_and_rebuild() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("row-column-formats.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let row_format = rust_xlsxwriter::Format::new()
        .set_bold()
        .set_background_color("ABCDEF")
        .set_font_scheme(rust_xlsxwriter::FontScheme::Headings);
    let column_format = rust_xlsxwriter::Format::new()
        .set_font_color("123456")
        .set_num_format("0.00")
        .set_font_scheme(rust_xlsxwriter::FontScheme::None);
    worksheet
        .set_row_format(0, &row_format)
        .expect("row format applies");
    worksheet
        .set_column_format(0, &column_format)
        .expect("column format applies");
    worksheet.write_string(0, 0, "value").expect("cell writes");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("row-column-formats.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("fontColor={\"#123456\"}"), "{source}");
    assert!(source.contains("numberFormat={\"0.00\"}"), "{source}");
    assert!(source.contains("<Row bold={true}"), "{source}");
    assert!(source.contains("fontScheme=\"headings\""), "{source}");
    assert!(source.contains("fontScheme=\"none\""), "{source}");
    assert!(source.contains("fillColor={\"#ABCDEF\"}"), "{source}");

    let regenerated = directory.path().join("row-column-formats-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(regenerated).expect("regenerated workbook opens"))
            .expect("regenerated workbook is a ZIP package");
    let mut sheet = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut sheet)
        .expect("worksheet XML is UTF-8");
    assert!(sheet.contains("s=\"2\" customFormat=\"1\""), "{sheet}");
    assert!(
        sheet.contains("<col min=\"1\" max=\"1\" width=\""),
        "{sheet}"
    );
}

#[test]
fn external_default_row_height_round_trips_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("external-height.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("Rows")
        .expect("worksheet name is valid")
        .set_default_row_height_pixels(24);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("external-height.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("defaultRowHeightPixels={24}"));

    let regenerated = directory.path().join("external-height-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("defaultRowHeight=\"18\""));
}

#[test]
fn external_hide_unused_rows_round_trips_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("hide-unused-rows.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().hide_unused_rows(true);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("hide-unused-rows.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("hideUnusedRows={true}"));

    let regenerated = directory.path().join("hide-unused-rows-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("zeroHeight=\"1\""));
}

#[test]
fn reverse_rejects_invalid_hide_unused_rows_boolean_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-zero-height.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().hide_unused_rows(true);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("zeroHeight=\"1\"", "zeroHeight=\"maybe\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "invalid `sheetFormatPr.zeroHeight`",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_screen_gridlines_round_trip_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("screen-gridlines.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_screen_gridlines(false);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("screen-gridlines.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("showGridlines={false}"));

    let regenerated = directory.path().join("screen-gridlines-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("showGridLines=\"0\""));
}

#[test]
fn external_print_gridlines_round_trip_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("print-gridlines.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_gridlines(true);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("print-gridlines.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("printGridlines={true}"));

    let regenerated = directory.path().join("print-gridlines-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("<printOptions gridLines=\"1\"/>"));
}

#[test]
fn external_worksheet_zoom_round_trips_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("worksheet-zoom.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_zoom(125);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("worksheet-zoom.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("zoom={125}"),
        "actual reverse output: {source}"
    );

    let regenerated = directory.path().join("worksheet-zoom-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("zoomScale=\"125\""));
}

#[test]
fn external_page_views_round_trip_with_view_specific_zoom() {
    for (view, zoom, expected_view, zoom_attr) in [
        ("pageLayout", 125, "pageLayout", "zoomScalePageLayoutView"),
        (
            "pageBreakPreview",
            90,
            "pageBreakPreview",
            "zoomScaleSheetLayoutView",
        ),
    ] {
        let directory = tempdir().expect("temporary directory is available");
        let input = directory.path().join("page-view.xlsx");
        let mut workbook = rust_xlsxwriter::Workbook::new();
        let worksheet = workbook.add_worksheet();
        if view == "pageLayout" {
            worksheet.set_view_page_layout();
        } else {
            worksheet.set_view_page_break_preview();
        }
        worksheet.set_zoom(zoom);
        workbook.save(&input).expect("external workbook saves");

        let reversed = directory.path().join("page-view.tsx");
        let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
        command
            .args(["reverse"])
            .arg(&input)
            .args(["-o"])
            .arg(&reversed);
        command.assert().success();
        let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
        assert!(
            source.contains(&format!("view=\"{expected_view}\"")),
            "{source}"
        );
        assert!(source.contains(&format!("zoom={{{zoom}}}")), "{source}");

        let regenerated = directory.path().join("page-view-roundtrip.xlsx");
        let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
        command
            .args(["build"])
            .arg(&reversed)
            .args(["-o"])
            .arg(&regenerated);
        command.assert().success();
        let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
            .expect("generated XLSX is a ZIP package");
        let mut xml = String::new();
        archive
            .by_name("xl/worksheets/sheet1.xml")
            .expect("sheet exists")
            .read_to_string(&mut xml)
            .expect("sheet XML reads");
        assert!(xml.contains(&format!("view=\"{expected_view}\"")), "{xml}");
        assert!(xml.contains(&format!("{zoom_attr}=\"{zoom}\"")), "{xml}");
    }
}

#[test]
fn external_manual_page_breaks_round_trip_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("page-breaks.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_page_breaks(&[40, 20])
        .expect("horizontal page breaks are valid");
    worksheet
        .set_vertical_page_breaks(&[8, 4])
        .expect("vertical page breaks are valid");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("page-breaks.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("horizontalPageBreaks={[20, 40]}"),
        "{source}"
    );
    assert!(source.contains("verticalPageBreaks={[4, 8]}"), "{source}");

    let regenerated = directory.path().join("page-breaks-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("<rowBreaks count=\"2\" manualBreakCount=\"2\">"));
    assert!(xml.contains("<brk id=\"20\" max=\"16383\" man=\"1\"/>"));
    assert!(xml.contains("<colBreaks count=\"2\" manualBreakCount=\"2\">"));
    assert!(xml.contains("<brk id=\"4\" max=\"1048575\" man=\"1\"/>"));
}

#[test]
fn external_page_margins_round_trip_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("page-margins.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_margins(0.5, 0.6, 1.0, 1.1, 0.2, 0.25);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("page-margins.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("pageMargins={{left: 0.5, right: 0.6, top: 1.0, bottom: 1.1, header: 0.2, footer: 0.25}}"), "{source}");

    let regenerated = directory.path().join("page-margins-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("<pageMargins left=\"0.5\" right=\"0.6\" top=\"1\" bottom=\"1.1\" header=\"0.2\" footer=\"0.25\"/>"));
}

#[test]
fn public_page_margins_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("page-margins.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/page-margins.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("page-margins.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse source exists");
    assert!(
        source.contains("pageMargins={{left: 0.5, top: 1.0}}"),
        "{source}"
    );
}

#[test]
fn reverse_rejects_negative_page_margin_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("negative-page-margin.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_margins(0.5, 0.6, 1.0, 1.1, 0.2, 0.25);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("left=\"0.5\"", "left=\"-0.5\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("pageMargins `left`"))
        .stderr(predicate::str::contains("finite nonnegative"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_print_scale_round_trips_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("print-scale.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_scale(80);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("print-scale.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("printScale={80}"), "{source}");

    let regenerated = directory.path().join("print-scale-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("<pageSetup scale=\"80\" orientation=\"portrait\" horizontalDpi=\"200\" verticalDpi=\"200\"/>"));
}

#[test]
fn upstream_first_page_number_serializes_a_number_into_boolean_ooxml_attribute() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("first-page-number.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_first_page_number(2);
    workbook.save(&input).expect("workbook writes");

    let mut archive = ZipArchive::new(std::fs::File::open(input).expect("workbook opens"))
        .expect("workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(
        worksheet_xml.contains("useFirstPageNumber=\"2\""),
        "pinned writer serialization changed: {worksheet_xml}"
    );
    assert!(
        !worksheet_xml.contains("firstPageNumber="),
        "pinned writer currently does not serialize the page number in the schema field: {worksheet_xml}"
    );
}

#[test]
fn external_page_orientation_and_paper_size_round_trip_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("page-setup.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_landscape().set_paper_size(9);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("page-setup.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("orientation=\"landscape\""), "{source}");
    assert!(source.contains("paperSize={9}"), "{source}");

    let regenerated = directory.path().join("page-setup-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("paperSize=\"9\""));
    assert!(xml.contains("orientation=\"landscape\""));
}

#[test]
fn external_print_layout_and_quality_options_round_trip_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("print-options.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_print_center_horizontally(true)
        .set_print_center_vertically(true)
        .set_print_black_and_white(true)
        .set_print_draft(true);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("print-options.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    for property in [
        "centerHorizontally={true}",
        "centerVertically={true}",
        "printBlackAndWhite={true}",
        "printDraft={true}",
    ] {
        assert!(source.contains(property), "missing {property} in {source}");
    }

    let regenerated = directory.path().join("print-options-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("horizontalCentered=\"1\""));
    assert!(xml.contains("verticalCentered=\"1\""));
    assert!(xml.contains("blackAndWhite=\"1\""));
    assert!(xml.contains("draft=\"1\""));
}

#[test]
fn external_print_headings_and_page_order_round_trip_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("print-headings.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_print_headings(true).set_page_order(false);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("print-headings.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("printHeadings={true}"), "{source}");
    assert!(source.contains("pageOrder=\"overThenDown\""), "{source}");

    let regenerated = directory.path().join("print-headings-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("headings=\"1\""));
    assert!(xml.contains("pageOrder=\"overThenDown\""));
}

#[test]
fn external_print_area_and_both_repeating_axes_round_trip_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("print-titles.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name("O'Brien, West")
        .expect("worksheet name is valid");
    worksheet
        .set_print_area(0, 0, 19, 3)
        .expect("print area is valid");
    worksheet
        .set_repeat_rows(0, 1)
        .expect("repeated rows are valid");
    worksheet
        .set_repeat_columns(0, 1)
        .expect("repeated columns are valid");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("print-titles.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("printArea=\"A1:D20\""), "{source}");
    assert!(source.contains("repeatRows=\"1:2\""), "{source}");
    assert!(source.contains("repeatColumns=\"A:B\""), "{source}");

    let regenerated = directory.path().join("print-titles-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/workbook.xml")
        .expect("workbook part exists")
        .read_to_string(&mut xml)
        .expect("workbook XML reads");
    assert!(xml.contains("name=\"_xlnm.Print_Area\" localSheetId=\"0\""));
    assert!(xml.contains("name=\"_xlnm.Print_Titles\" localSheetId=\"0\""));
    assert!(xml.contains("'O''Brien, West'!$A$1:$D$20"));
    assert!(xml.contains("'O''Brien, West'!$A:$B,'O''Brien, West'!$1:$2"));
}

#[test]
fn external_full_row_and_full_column_print_areas_round_trip() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("whole-axis-print-areas.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("Rows")
        .expect("worksheet name is valid")
        .set_print_area(2, 0, 5, 16_383)
        .expect("full-row print area is valid");
    workbook
        .add_worksheet()
        .set_name("Columns")
        .expect("worksheet name is valid")
        .set_print_area(0, 2, 1_048_575, 4)
        .expect("full-column print area is valid");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("whole-axis-print-areas.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("printArea=\"3:6\""), "{source}");
    assert!(source.contains("printArea=\"C:E\""), "{source}");

    let regenerated = directory
        .path()
        .join("whole-axis-print-areas-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
}

#[test]
fn public_print_ranges_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("print-ranges.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/print-ranges.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("print-ranges.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse source exists");
    assert!(source.contains("printArea=\"A1:D20\""), "{source}");
    assert!(source.contains("repeatRows=\"1:2\""), "{source}");
    assert!(source.contains("repeatColumns=\"A:A\""), "{source}");
}

#[test]
fn reverse_rejects_invalid_print_quality_boolean_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("print-draft-invalid.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_draft(true);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("draft=\"1\"", "draft=\"perhaps\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("pageSetup.draft"))
        .stderr(predicate::str::contains("0/1 or false/true"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn public_print_options_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("print-options.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/print-options.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("print-options.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("centerHorizontally={true}"), "{source}");
    assert!(source.contains("centerVertically={true}"), "{source}");
    assert!(source.contains("printBlackAndWhite={true}"), "{source}");
    assert!(source.contains("printDraft={true}"), "{source}");
    assert!(source.contains("printHeadings={true}"), "{source}");
    assert!(source.contains("pageOrder=\"overThenDown\""), "{source}");
}

#[test]
fn external_fit_to_pages_round_trips_with_page_setup_defaults() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("fit-to-pages.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_fit_to_pages(1, 0);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("fit-to-pages.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("fitToPages={{width: 1, height: 0}}"),
        "{source}"
    );

    let regenerated = directory.path().join("fit-to-pages-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("<pageSetUpPr fitToPage=\"1\"/>"));
    assert!(xml.contains("fitToHeight=\"0\""));
}

#[test]
fn reverse_preserves_fit_to_one_by_one_pages_when_dimensions_are_omitted() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("fit-to-one-page.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_fit_to_pages(1, 1);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("fit-to-one-page.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(
        source.contains("fitToPages={{width: 1, height: 1}}"),
        "{source}"
    );
}

#[test]
fn public_print_scaling_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("print-scaling.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/print-scaling.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("print-scaling.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse source exists");
    assert!(source.contains("printScale={80}"), "{source}");
    assert!(
        source.contains("fitToPages={{width: 1, height: 0}}"),
        "{source}"
    );
}

#[test]
fn reverse_omits_default_print_scale() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("default-print-scale.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_scale(100);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("default-print-scale.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(!source.contains("printScale="), "{source}");
}

#[test]
fn reverse_rejects_conflicting_scale_and_fit_to_pages() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("conflicting-print-settings.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_fit_to_pages(1, 0);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "pageSetup fitToHeight=\"0\"",
            "pageSetup scale=\"80\" fitToHeight=\"0\"",
        )
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "both pageSetup scale and fitToPage enabled",
        ))
        .stderr(predicate::str::contains("mutually exclusive"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_orphan_fit_dimensions_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("orphan-fit-width.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_scale(80);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("scale=\"80\"", "scale=\"80\" fitToWidth=\"2\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "fitToWidth/fitToHeight without fitToPage=1",
        ))
        .stderr(predicate::str::contains(
            "remove the orphaned fit dimensions",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_fit_to_pages_that_overrides_manual_breaks() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("fit-overrides-breaks.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_print_fit_to_pages(1, 0);
    worksheet
        .set_page_breaks(&[20])
        .expect("manual break is valid");
    workbook.save(&input).expect("external workbook saves");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "fitToPage with manual page breaks",
        ))
        .stderr(predicate::str::contains("overrides"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_partial_manual_page_breaks_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("partial-page-break.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_page_breaks(&[20])
        .expect("horizontal page break is valid");
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("max=\"16383\"", "max=\"20\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("partial page break"))
        .stderr(predicate::str::contains("max=16383"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_automatic_page_breaks_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("automatic-page-break.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_page_breaks(&[20])
        .expect("horizontal page break is valid");
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("man=\"1\"", "man=\"0\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("automatic page break"))
        .stderr(predicate::str::contains("only manual breaks"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_page_break_count_mismatches_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("page-break-count.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_page_breaks(&[20])
        .expect("horizontal page break is valid");
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("rowBreaks count=\"1\"", "rowBreaks count=\"2\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("declares count=2"))
        .stderr(predicate::str::contains("match the listed breaks"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_non_default_zoom_for_inactive_worksheet_view() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("inactive-view-zoom.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_zoom(125);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "zoomScale=\"125\"",
            "zoomScale=\"125\" zoomScalePageLayoutView=\"130\"",
        )
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("inactive worksheet view"))
        .stderr(predicate::str::contains("view=\"pageLayout\""))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_right_to_left_worksheet_round_trips_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("right-to-left.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_right_to_left(true)
        .write_string(0, 0, "مرحبا")
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("right-to-left.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("rightToLeft={true}"));

    let regenerated = directory.path().join("right-to-left-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let xml = read_zip_part_for_test(&regenerated, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("rightToLeft=\"1\""));
    let second_reverse = directory.path().join("right-to-left-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn public_right_to_left_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("right-to-left.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/right-to-left.tsx", "-o"])
        .arg(&output);
    command.assert().success();
    let xml = read_zip_part_for_test(&output, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("rightToLeft=\"1\""));

    let reversed = directory.path().join("right-to-left.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    assert!(
        std::fs::read_to_string(reversed)
            .expect("reverse output exists")
            .contains("rightToLeft={true}")
    );
}

#[test]
fn reverse_rejects_invalid_right_to_left_boolean_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-right-to-left.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_right_to_left(true);
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("rightToLeft=\"1\"", "rightToLeft=\"yes\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("rightToLeft"))
        .stderr(predicate::str::contains("expected 0/1 or false/true"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_outline_summary_directions_round_trip_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("outline-directions.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.group_symbols_above(true);
    worksheet.group_symbols_to_left(true);
    worksheet.group_rows(1, 2).expect("row outline is added");
    worksheet
        .group_columns(1, 2)
        .expect("column outline is added");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("outline-directions.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("summaryBelow={false}"));
    assert!(source.contains("summaryRight={false}"));

    let regenerated = directory.path().join("outline-directions-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let xml = read_zip_part_for_test(&regenerated, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("<outlinePr summaryBelow=\"0\" summaryRight=\"0\""));

    let second_reverse = directory.path().join("outline-directions-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn reverse_rejects_unmodeled_outline_properties_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("unsupported-outline-property.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().group_symbols_above(true);
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "<outlinePr summaryBelow=\"0\"",
            "<outlinePr applyStyles=\"1\" summaryBelow=\"0\"",
        )
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("outlinePr contains unsupported"))
        .stderr(predicate::str::contains("applyStyles"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_omits_default_worksheet_zoom() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("worksheet-default-zoom.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_zoom(100);
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("worksheet-default-zoom.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(!source.contains("zoom="));
}

#[test]
fn reverse_rejects_non_default_worksheet_zoom_normal() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("worksheet-zoom-normal.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_zoom(125);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("zoomScaleNormal=\"125\"", "zoomScaleNormal=\"120\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("differs from `zoomScale`"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_worksheet_zoom_outside_supported_range() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("worksheet-zoom-invalid.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_zoom(125);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("zoomScale=\"125\"", "zoomScale=\"401\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("zoomScale"))
        .stderr(predicate::str::contains("10 through 400"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_preserves_page_orientation_with_print_gridlines() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("print-gridlines-landscape.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_print_gridlines(true);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("orientation=\"portrait\"", "orientation=\"landscape\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    let reversed = directory.path().join("print-gridlines-landscape.tsx");
    command
        .args(["reverse"])
        .arg(input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("orientation=\"landscape\""), "{source}");
}

#[test]
fn reverse_rejects_paper_size_outside_writer_range_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("paper-size-invalid.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet().set_paper_size(9);
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("paperSize=\"9\"", "paperSize=\"256\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("paperSize"))
        .stderr(predicate::str::contains("0 through 255"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_array_formula_round_trips_through_jsx() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("array-formula.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(0, 0, 2).expect("input value writes");
    worksheet.write_number(1, 0, 3).expect("input value writes");
    worksheet
        .write_array_formula(0, 1, 1, 1, "A1:A2*2")
        .expect("array formula writes");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("array-formula.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("formula={\"=A1:A2*2\"}"));
    assert!(source.contains("arrayFormulaRange={\"B1:B2\"}"));

    let regenerated = directory.path().join("array-formula-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("file opens"))
        .expect("generated XLSX is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet exists")
        .read_to_string(&mut xml)
        .expect("sheet XML reads");
    assert!(xml.contains("<f t=\"array\" ref=\"B1:B2\">A1:A2*2</f>"));
}

#[test]
fn reverse_rejects_array_formula_cached_results_it_cannot_preserve() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("array-formula-result.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_array_formula(
            0,
            0,
            0,
            0,
            rust_xlsxwriter::Formula::new("1+1").set_result("2"),
        )
        .expect("array formula writes");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("array formula cached result"))
        .stderr(predicate::str::contains("only writer-default zero caches"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_formula_cached_results_round_trip_by_type() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("formula-results.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_formula(0, 0, rust_xlsxwriter::Formula::new("1+1").set_result("2"))
        .expect("numeric formula writes");
    worksheet
        .write_formula(
            0,
            1,
            rust_xlsxwriter::Formula::new("\"ready\"").set_result("ready"),
        )
        .expect("string formula writes");
    worksheet
        .write_formula(
            0,
            2,
            rust_xlsxwriter::Formula::new("TRUE()").set_result("TRUE"),
        )
        .expect("boolean formula writes");
    worksheet
        .write_formula(
            0,
            3,
            rust_xlsxwriter::Formula::new("NA()").set_result("#N/A"),
        )
        .expect("error formula writes");
    worksheet
        .write_formula(0, 4, "1+1")
        .expect("blank-cache formula writes");
    worksheet.set_formula_result(0, 4, "");
    workbook
        .save(&input)
        .expect("external formula workbook writes");

    let reversed = directory.path().join("formula-results.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("formulaResult={2}"));
    assert!(source.contains("formulaResult=\"ready\""), "{source}");
    assert!(source.contains("formulaResult={true}"));
    assert!(source.contains("formulaResult=\"#N/A\""));
    assert!(source.contains("formulaResult={null}"));

    let output = directory.path().join("formula-results-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output);
    command.assert().success();

    let second_reverse = directory.path().join("formula-results-second-reverse.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(output)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn public_formula_results_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("formula-results.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/formula-results.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("formula-results.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("formulaResult={2}"));
    assert!(source.contains("formulaResult=\"ready\""));
    assert!(source.contains("formulaResult={true}"));
    assert!(source.contains("formulaResult=\"#N/A\""));
}

#[test]
fn reverse_rejects_ambiguous_string_formula_cache_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("ambiguous-formula-result.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_formula(
            0,
            0,
            rust_xlsxwriter::Formula::new("\"123\"").set_result("123"),
        )
        .expect("formula cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("<c r=\"A1\"", "<c r=\"A1\" t=\"str\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(&input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("cached text result"))
        .stderr(predicate::str::contains("different value type"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_array_formula_with_mismatched_anchor() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("bad-array-anchor.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_array_formula(0, 0, 1, 0, "1+1")
        .expect("array formula writes");
    workbook.save(&input).expect("external workbook saves");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("ref=\"A1:A2\"", "ref=\"B1:B2\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("not the top-left"))
        .stderr(predicate::str::contains("repair"));
}

#[test]
fn jsx_rich_text_builds_and_round_trips_through_external_xlsx() {
    let directory = tempdir().expect("temporary directory is available");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["validate", "examples/rich-text.tsx"]);
    command.assert().success().stdout("valid\n");

    let output = directory.path().join("rich-text.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/rich-text.tsx", "-o"])
        .arg(&output);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(&output).expect("XLSX opens"))
        .expect("XLSX is a ZIP package");
    let mut shared_strings = String::new();
    archive
        .by_name("xl/sharedStrings.xml")
        .expect("rich strings use the shared string table")
        .read_to_string(&mut shared_strings)
        .expect("shared strings read");
    assert!(shared_strings.contains("<rPr><b/>"), "{shared_strings}");
    assert!(shared_strings.contains("FF008000"), "{shared_strings}");
    assert!(
        shared_strings.contains("<u val=\"double\"/>"),
        "{shared_strings}"
    );
    assert!(
        shared_strings.contains("<vertAlign val=\"superscript\"/>")
            && shared_strings.contains("<charset val=\"128\"/>"),
        "{shared_strings}"
    );

    let reversed = directory.path().join("rich-text.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("<TextRun>{\"Status: \"}</TextRun>"),
        "{source}"
    );
    assert!(
        source.contains(
            "<TextRun bold={true} fontColor={\"#008000\"} fontSize={11}>{\"Ready\"}</TextRun>"
        ),
        "{source}"
    );
    assert!(
        source.contains("<TextRun italic={true} fontSize={11}>{\" — verified\"}</TextRun>"),
        "{source}"
    );
    assert!(
        source.contains("strikethrough={true}")
            && source.contains("underline={\"double\"}")
            && source.contains("script={\"superscript\"}")
            && source.contains("fontName={\"Aptos\"}")
            && source.contains("fontSize={12.5}")
            && source.contains("fontFamily={3}")
            && source.contains("fontCharset={128}"),
        "{source}"
    );

    let regenerated = directory.path().join("rich-text-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("rich-text-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn reverse_preserves_inline_rich_text_runs() {
    let directory = tempdir().expect("temporary directory is available");
    let shared = directory.path().join("shared.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let plain = rust_xlsxwriter::Format::new();
    let bold = rust_xlsxwriter::Format::new()
        .set_bold()
        .set_font_color("008000");
    worksheet
        .write_rich_string_with_format(
            0,
            0,
            &[(&plain, "prefix "), (&bold, "green")],
            &rust_xlsxwriter::Format::new().set_align(rust_xlsxwriter::FormatAlign::Center),
        )
        .expect("rich string writes");
    workbook.save(&shared).expect("workbook writes");

    let mut archive = ZipArchive::new(std::fs::File::open(&shared).expect("workbook opens"))
        .expect("XLSX is a ZIP package");
    let mut strings = String::new();
    archive
        .by_name("xl/sharedStrings.xml")
        .expect("shared strings exist")
        .read_to_string(&mut strings)
        .expect("shared strings read");
    let start = strings.find("<si>").expect("shared string entry starts") + 4;
    let end = strings.find("</si>").expect("shared string entry ends");
    let inline_content = strings[start..end].to_owned();
    let inline = directory.path().join("inline.xlsx");
    std::fs::copy(&shared, &inline).expect("workbook copied");
    rewrite_part(&inline, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "<c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c>",
            &format!("<c r=\"A1\" s=\"1\" t=\"inlineStr\"><is>{inline_content}</is></c>"),
        )
    });

    let reversed = directory.path().join("inline.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&inline)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("<TextRun>{\"prefix \"}</TextRun>"),
        "{source}"
    );
    assert!(
        source.contains("<TextRun bold={true} fontColor={\"#008000\"}"),
        "{source}"
    );
    assert!(source.contains("{\"green\"}</TextRun>"), "{source}");

    let regenerated = directory.path().join("inline-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let second_reverse = directory.path().join("inline-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    let second_source = std::fs::read_to_string(second_reverse).expect("second reverse exists");
    assert!(
        second_source.contains("<TextRun>{\"prefix \"}</TextRun>"),
        "{second_source}"
    );
    assert!(
        second_source.contains("{\"green\"}</TextRun>"),
        "{second_source}"
    );
}

#[test]
fn reverse_rejects_unmodeled_rich_text_and_phonetic_metadata() {
    let directory = tempdir().expect("temporary directory is available");
    for (name, rich_entry, expected) in [
        (
            "unsupported-run",
            "<si><r><rPr><outline/></rPr><t>rich</t></r></si>",
            "unsupported rich-text font properties",
        ),
        (
            "phonetics",
            "<si><t>plain</t><phoneticPr fontId=\"0\" type=\"noConversion\"/></si>",
            "phonetic annotations",
        ),
    ] {
        let input = directory.path().join(format!("{name}.xlsx"));
        let mut workbook = rust_xlsxwriter::Workbook::new();
        workbook
            .add_worksheet()
            .write_string(0, 0, "plain")
            .expect("cell writes");
        workbook.save(&input).expect("workbook writes");
        rewrite_part(&input, "xl/sharedStrings.xml", |xml| {
            xml.replace("<si><t>plain</t></si>", rich_entry)
        });
        let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
        command
            .args(["reverse"])
            .arg(&input)
            .assert()
            .failure()
            .stderr(predicate::str::contains(expected))
            .stderr(predicate::str::contains("lossless TextRun subset"))
            .stderr(predicate::str::contains(
                "supported rich-text runs already round-trip",
            ))
            .stderr(predicate::str::contains("suggestion:"));
    }
}

#[test]
fn public_worksheet_visibility_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("visibility.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/worksheet-visibility.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("visibility.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("visibility=\"hidden\""));
    assert!(source.contains("visibility=\"veryHidden\""));
}

#[test]
fn public_table_formats_example_reverses_header_font_metadata() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("table-formats.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/table-formats.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("table-formats.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("fontName={\"Aptos\"}"), "{source}");
    assert!(source.contains("fontSize={12.5}"), "{source}");
    assert!(source.contains("fontFamily={3}"), "{source}");
    assert!(source.contains("fontCharset={128}"), "{source}");
}

#[test]
fn external_tables_survive_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Product")
        .expect("header writes");
    worksheet
        .write_string(0, 1, "Amount")
        .expect("header writes");
    worksheet.write_string(1, 0, "Tea").expect("cell writes");
    worksheet.write_number(1, 1, 12).expect("cell writes");
    worksheet.write_string(2, 0, "Coffee").expect("cell writes");
    worksheet.write_number(2, 1, 18).expect("cell writes");
    let columns = [
        rust_xlsxwriter::TableColumn::new().set_header("Product"),
        rust_xlsxwriter::TableColumn::new().set_header("Amount"),
    ];
    let table = rust_xlsxwriter::Table::new()
        .set_name("SalesTable")
        .set_columns(&columns)
        .set_style(rust_xlsxwriter::TableStyle::Medium4)
        .set_autofilter(false)
        .set_banded_rows(false)
        .set_first_column(true)
        .set_alt_text("Accessible sales data")
        .set_alt_text_title("Sales summary");
    worksheet
        .add_table(0, 0, 2, 1, &table)
        .expect("table is added");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("table.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains(
        "<Table range={\"A1:B3\"} name={\"SalesTable\"} style={\"TableStyleMedium4\"} autoFilter={false} bandedRows={false} firstColumn={true} altText={\"Accessible sales data\"} altTextTitle={\"Sales summary\"} />"
    ), "{source}");

    let regenerated = directory.path().join("table-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    let second_reverse = directory.path().join("table-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_table_accessibility_extension_rejects_unmodeled_attributes() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table-extension.xlsx");
    let mutated = directory.path().join("table-extension-unsupported.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Name").expect("header writes");
    worksheet.write_string(1, 0, "Tea").expect("cell writes");
    let table = rust_xlsxwriter::Table::new()
        .set_name("Products")
        .set_alt_text("Accessible product list");
    worksheet.add_table(0, 0, 1, 0, &table).expect("table adds");
    workbook.save(&input).expect("workbook writes");

    let mut archive = ZipArchive::new(std::fs::File::open(&input).expect("workbook opens"))
        .expect("workbook is a ZIP package");
    let mut writer = ZipWriter::new(std::fs::File::create(&mutated).expect("mutated file opens"));
    let mut changed = false;
    for index in 0..archive.len() {
        let mut part = archive.by_index(index).expect("ZIP part exists");
        let name = part.name().to_owned();
        let mut bytes = Vec::new();
        part.read_to_end(&mut bytes).expect("ZIP part reads");
        if name == "xl/tables/table1.xml" {
            let xml = String::from_utf8(bytes).expect("table XML is UTF-8");
            let marker = "<x14:table";
            assert!(
                xml.contains(marker),
                "writer output has the x14 extension: {xml}"
            );
            bytes = xml
                .replacen(marker, "<x14:table unsupported=\"true\"", 1)
                .into_bytes();
            changed = true;
        }
        writer
            .start_file(name, SimpleFileOptions::default())
            .expect("mutated ZIP part starts");
        writer.write_all(&bytes).expect("mutated ZIP part writes");
    }
    assert!(changed, "the table part was found");
    writer.finish().expect("mutated workbook finishes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&mutated)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "unsupported accessibility attributes",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_table_column_font_scripts_round_trip_without_manifest() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table-scripts.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    for (column, value) in [(0, "High"), (1, "Low")] {
        worksheet
            .write_string(0, column, value)
            .expect("header writes");
        worksheet
            .write_string(1, column, "Value")
            .expect("cell writes");
    }
    let columns = [
        rust_xlsxwriter::TableColumn::new()
            .set_header("High")
            .set_format(
                rust_xlsxwriter::Format::new()
                    .set_bold()
                    .set_font_script(rust_xlsxwriter::FormatScript::Superscript),
            ),
        rust_xlsxwriter::TableColumn::new()
            .set_header("Low")
            .set_format(
                rust_xlsxwriter::Format::new()
                    .set_bold()
                    .set_font_script(rust_xlsxwriter::FormatScript::Subscript),
            ),
    ];
    let table = rust_xlsxwriter::Table::new()
        .set_name("Measurements")
        .set_columns(&columns);
    worksheet.add_table(0, 0, 1, 1, &table).expect("table adds");
    workbook.save(&input).expect("workbook writes");
    let reversed = directory.path().join("table-scripts.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("\"script\":\"superscript\"")
            && source.contains("\"script\":\"subscript\""),
        "{source}"
    );

    let rebuilt = directory.path().join("table-scripts-rebuilt.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();

    let second_reverse = directory.path().join("table-scripts-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(std::fs::read_to_string(second_reverse).unwrap(), source);
}

#[test]
fn external_legacy_note_survives_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("note.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "item").expect("cell writes");
    let note = rust_xlsxwriter::Note::new("Review this value")
        .set_author("Rust")
        .add_author_prefix(false)
        .set_width(200)
        .set_height(100)
        .set_visible(true)
        .set_background_color("#AABBCC")
        .set_font_name("Arial")
        .set_font_size(9.5)
        .set_font_family(2)
        .set_alt_text("Review note");
    worksheet.insert_note(0, 0, &note).expect("note inserts");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("note.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("<Note cell={\"A1\"} text={\"Review this value\"}"),
        "{source}"
    );
    assert!(
        source.contains("author={\"Rust\"} authorPrefix={false}"),
        "{source}"
    );
    assert!(source.contains("backgroundColor={\"#AABBCC\"}"), "{source}");

    let rebuilt = directory.path().join("note-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let second = directory.path().join("note-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second).expect("second reverse exists"),
        source
    );
}

#[test]
fn external_note_with_unmodeled_rich_text_is_rejected_with_repair() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("plain-note.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let note = rust_xlsxwriter::Note::new("Review this value");
    worksheet.insert_note(0, 0, &note).expect("note inserts");
    workbook.save(&input).expect("workbook writes");

    let mutated = directory.path().join("italic-note.xlsx");
    let mut archive = ZipArchive::new(std::fs::File::open(&input).expect("workbook opens"))
        .expect("workbook is a ZIP package");
    let mut writer = ZipWriter::new(std::fs::File::create(&mutated).expect("mutated file opens"));
    for index in 0..archive.len() {
        let mut part = archive.by_index(index).expect("ZIP part exists");
        let name = part.name().to_owned();
        let mut bytes = Vec::new();
        part.read_to_end(&mut bytes).expect("ZIP part reads");
        if name == "xl/comments1.xml" {
            let xml = String::from_utf8(bytes).expect("comments XML is UTF-8");
            bytes = xml.replace("<rPr>", "<rPr><i/>").into_bytes();
        }
        writer
            .start_file(name, SimpleFileOptions::default())
            .expect("mutated ZIP part starts");
        writer.write_all(&bytes).expect("mutated ZIP part writes");
    }
    writer.finish().expect("mutated workbook finishes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&mutated)
        .assert()
        .failure()
        .stderr(predicate::str::contains("legacy note"))
        .stderr(predicate::str::contains("cannot be represented losslessly"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_note_with_unrepresentable_hidden_row_position_is_rejected() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("moved-note.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_row_hidden(0).expect("row hides");
    let note = rust_xlsxwriter::Note::new("Hidden-row note")
        .set_object_movement(rust_xlsxwriter::ObjectMovement::MoveAndSizeWithCellsAfter);
    worksheet.insert_note(1, 1, &note).expect("note inserts");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("VML positioning"))
        .stderr(predicate::str::contains("row/column layout"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_table_totals_and_calculated_columns_survive_reverse_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table-totals.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Product")
        .expect("header writes");
    worksheet.write_string(1, 0, "Tea").expect("cell writes");
    worksheet
        .write_string(0, 1, "Amount")
        .expect("header writes");
    worksheet.write_number(1, 1, 12).expect("cell writes");
    worksheet.write_string(0, 2, "Tax").expect("header writes");
    worksheet.write_number(1, 2, 1.2).expect("cell writes");
    worksheet
        .write_string(0, 3, "CustomTotal")
        .expect("header writes");
    worksheet.write_number(1, 3, 12).expect("cell writes");
    let columns = [
        rust_xlsxwriter::TableColumn::new()
            .set_header("Product")
            .set_total_label("Total")
            .set_header_format(rust_xlsxwriter::Format::new().set_bold()),
        rust_xlsxwriter::TableColumn::new()
            .set_header("Amount")
            .set_total_function(rust_xlsxwriter::TableFunction::Sum)
            .set_format(
                rust_xlsxwriter::Format::new()
                    .set_font_color("#FF0000")
                    .set_background_color("#FFFF00"),
            ),
        rust_xlsxwriter::TableColumn::new()
            .set_header("Tax")
            .set_formula("[@Amount]*0.1"),
        rust_xlsxwriter::TableColumn::new()
            .set_header("CustomTotal")
            .set_total_function(rust_xlsxwriter::TableFunction::Custom(
                rust_xlsxwriter::Formula::new("SUM([Amount])"),
            )),
    ];
    let table = rust_xlsxwriter::Table::new()
        .set_name("SalesTable")
        .set_total_row(true)
        .set_columns(&columns);
    worksheet
        .add_table(0, 0, 2, 3, &table)
        .expect("table with totals row is added");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("table-totals.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("totalsRow={true}"), "{source}");
    assert!(
        source.contains("<TableColumn totalLabel={\"Total\"} />"),
        "{source}"
    );
    assert!(
        source.contains(
            "totalFunction={\"sum\"} format={{\"fontColor\":\"#FF0000\",\"fillColor\":\"#FFFF00\"}}"
        ),
        "{source}"
    );
    assert!(
        source.contains("<Cell bold={true} fontSize={11}>{\"Product\"}"),
        "{source}"
    );
    assert!(
        source.contains("<TableColumn formula={\"[[#This Row],Amount]*0.1\"} />"),
        "{source}"
    );
    assert!(
        source.contains("<TableColumn totalFormula={\"SUM([Amount])\"} />"),
        "{source}"
    );

    let rebuilt = directory.path().join("table-totals-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let second = directory.path().join("table-totals-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second);
    command.assert().success();
    assert_eq!(std::fs::read_to_string(second).unwrap(), source);
}

#[test]
fn jsx_tables_compile_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let built = directory.path().join("tables.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/tables.tsx", "-o"])
        .arg(&built);
    command.assert().success();

    let reversed = directory.path().join("tables.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&built)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains(
        "<Table range={\"A1:B3\"} name={\"SalesTable\"} style={\"TableStyleMedium4\"} />"
    ));

    let rebuilt = directory.path().join("tables-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let second_reverse = directory.path().join("tables-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn jsx_table_column_header_format_compiles_and_reverses_as_cell_format() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("table-header-format.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/table-totals.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("table-header-format.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reversed JSX exists");
    assert!(
        source.contains("<Cell bold={true} fontSize={11}>{\"Product\"}"),
        "{source}"
    );
}

#[test]
fn external_table_column_strikethrough_and_underline_round_trip() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table-decorations.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Product")
        .expect("header writes");
    worksheet.write_string(1, 0, "Tea").expect("cell writes");
    let format = rust_xlsxwriter::Format::new()
        .set_font_strikethrough()
        .set_underline(rust_xlsxwriter::FormatUnderline::Double);
    let columns = [rust_xlsxwriter::TableColumn::new()
        .set_header("Product")
        .set_format(format)];
    let table = rust_xlsxwriter::Table::new()
        .set_name("Products")
        .set_columns(&columns);
    worksheet.add_table(0, 0, 1, 0, &table).expect("table adds");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("table-decorations.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("format={{\"strikethrough\":true,\"underline\":\"double\"}}"),
        "{source}"
    );

    let rebuilt = directory.path().join("table-decorations-rebuilt.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let second_reverse = directory.path().join("table-decorations-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_table_column_number_format_round_trips() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table-number-format.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Amount")
        .expect("header writes");
    worksheet.write_number(1, 0, 12.5).expect("cell writes");
    let columns = [rust_xlsxwriter::TableColumn::new()
        .set_header("Amount")
        .set_format(rust_xlsxwriter::Format::new().set_num_format("#,##0.000"))];
    let table = rust_xlsxwriter::Table::new()
        .set_name("Amounts")
        .set_columns(&columns);
    worksheet.add_table(0, 0, 1, 0, &table).expect("table adds");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("table-number-format.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("format={{\"numberFormat\":\"#,##0.000\"}}"),
        "{source}"
    );
    let rebuilt = directory.path().join("table-number-format-rebuilt.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let second_reverse = directory.path().join("table-number-format-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_table_header_font_metadata_round_trips_without_manifest() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table-header-font-metadata.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Amount")
        .expect("header writes");
    worksheet.write_number(1, 0, 12.5).expect("cell writes");
    let columns = [rust_xlsxwriter::TableColumn::new()
        .set_header("Amount")
        .set_header_format(
            rust_xlsxwriter::Format::new()
                .set_font_name("Aptos")
                .set_font_size(12.5)
                .set_font_family(3)
                .set_font_charset(128),
        )];
    let table = rust_xlsxwriter::Table::new()
        .set_name("Amounts")
        .set_columns(&columns);
    worksheet.add_table(0, 0, 1, 0, &table).expect("table adds");
    workbook.save(&input).expect("workbook writes");
    let reversed = directory.path().join("table-header-font-metadata.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("fontName={\"Aptos\"}"), "{source}");
    assert!(source.contains("fontSize={12.5}"), "{source}");
    assert!(source.contains("fontFamily={3}"), "{source}");
    assert!(source.contains("fontCharset={128}"), "{source}");

    let rebuilt = directory
        .path()
        .join("table-header-font-metadata-rebuilt.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();

    let second_reverse = directory
        .path()
        .join("table-header-font-metadata-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_table_column_pattern_fill_round_trips() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table-pattern-fill.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Amount")
        .expect("header writes");
    worksheet.write_number(1, 0, 12.5).expect("cell writes");
    let columns = [rust_xlsxwriter::TableColumn::new()
        .set_header("Amount")
        .set_format(
            rust_xlsxwriter::Format::new()
                .set_pattern(rust_xlsxwriter::FormatPattern::DarkDown)
                .set_foreground_color("#FF0000")
                .set_background_color("#FFFF00"),
        )];
    let table = rust_xlsxwriter::Table::new()
        .set_name("Amounts")
        .set_columns(&columns);
    worksheet.add_table(0, 0, 1, 0, &table).expect("table adds");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("table-pattern-fill.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains(r#""fillPattern":"darkDown""#), "{source}");
    assert!(
        source.contains(r##""fillForegroundColor":"#FF0000""##),
        "{source}"
    );
    assert!(
        source.contains(r##""fillBackgroundColor":"#FFFF00""##),
        "{source}"
    );

    let rebuilt = directory.path().join("table-pattern-fill-rebuilt.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let second = directory.path().join("table-pattern-fill-second.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second)
        .assert()
        .success();
    assert_eq!(
        std::fs::read_to_string(second).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_conditional_pattern_fill_round_trips() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("conditional-pattern-fill.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_number(0, 0, 15)
        .expect("sample number writes");
    let rule = rust_xlsxwriter::ConditionalFormatCell::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatCellRule::EqualTo(15))
        .set_format(
            rust_xlsxwriter::Format::new()
                .set_pattern(rust_xlsxwriter::FormatPattern::DarkTrellis)
                .set_foreground_color("#FF0000")
                .set_background_color("#FFFF00"),
        );
    worksheet
        .add_conditional_format(0, 0, 9, 0, &rule)
        .expect("conditional format writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("conditional-pattern-fill.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("fillPattern=\"darkTrellis\""), "{source}");
    assert!(
        source.contains("fillForegroundColor=\"#FF0000\""),
        "{source}"
    );
    assert!(
        source.contains("fillBackgroundColor=\"#FFFF00\""),
        "{source}"
    );

    let rebuilt = directory
        .path()
        .join("conditional-pattern-fill-rebuilt.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let second = directory.path().join("conditional-pattern-fill-second.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second)
        .assert()
        .success();
    assert_eq!(
        std::fs::read_to_string(second).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_table_column_borders_round_trip() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table-borders.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Amount")
        .expect("header writes");
    worksheet.write_number(1, 0, 12.5).expect("cell writes");
    let format = rust_xlsxwriter::Format::new()
        .set_border_top(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_top_color("#123456")
        .set_border_bottom(rust_xlsxwriter::FormatBorder::Dashed)
        .set_border_diagonal_type(rust_xlsxwriter::FormatDiagonalBorder::BorderUpDown);
    let columns = [rust_xlsxwriter::TableColumn::new()
        .set_header("Amount")
        .set_format(format)];
    let table = rust_xlsxwriter::Table::new()
        .set_name("Amounts")
        .set_columns(&columns);
    worksheet.add_table(0, 0, 1, 0, &table).expect("table adds");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("table-borders.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains(
            "format={{\"borderTop\":\"thin\",\"borderTopColor\":\"#123456\",\"borderBottom\":\"dashed\",\"borderDiagonalType\":\"upDown\"}}"
        ),
        "{source}"
    );
    let rebuilt = directory.path().join("table-borders-rebuilt.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let second_reverse = directory.path().join("table-borders-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_table_column_border_metadata_outside_supported_subset_is_rejected() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("table-border-metadata.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Amount")
        .expect("header writes");
    worksheet.write_number(1, 0, 12.5).expect("cell writes");
    let columns = [rust_xlsxwriter::TableColumn::new()
        .set_header("Amount")
        .set_format(
            rust_xlsxwriter::Format::new().set_border_top(rust_xlsxwriter::FormatBorder::Thin),
        )];
    let table = rust_xlsxwriter::Table::new()
        .set_name("Amounts")
        .set_columns(&columns);
    worksheet.add_table(0, 0, 1, 0, &table).expect("table adds");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/styles.xml", |xml| {
        let dxfs_start = xml.find("<dxfs").expect("differential styles exist");
        let dxfs_end = xml.find("</dxfs>").expect("differential styles close") + "</dxfs>".len();
        let dxfs = xml[dxfs_start..dxfs_end].replace("<border>", "<border outline=\"1\">");
        format!("{}{}{}", &xml[..dxfs_start], dxfs, &xml[dxfs_end..])
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(&input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "differential border contains unsupported edge attributes",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_row_height_and_hidden_state_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("rows.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "tall").expect("cell writes");
    worksheet.set_row_height(0, 22.5).expect("row height sets");
    worksheet.write_string(1, 0, "hidden").expect("cell writes");
    worksheet.set_row_hidden(1).expect("row visibility changes");
    worksheet
        .set_row_height(3, 15.75)
        .expect("blank row height sets");
    worksheet
        .set_row_hidden(4)
        .expect("trailing blank row hides");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("rows.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("<Row height={22.5}>"));
    assert!(source.contains("<Row hidden={true}>"));
    assert!(source.contains("<Row height={15.75}>"));
    assert!(source.contains("<Row hidden={true}></Row>"));

    let regenerated = directory.path().join("rows-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    let file = std::fs::File::open(regenerated).expect("regenerated workbook opens");
    let mut archive = ZipArchive::new(file).expect("regenerated workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains("ht=\"22.5\""));
    assert!(worksheet_xml.contains("ht=\"15.75\""));
    assert!(worksheet_xml.contains("hidden=\"1\""));
}

#[test]
fn external_expanded_row_and_column_outlines_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("outlines.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "summary")
        .expect("cell writes");
    worksheet.write_string(1, 0, "detail").expect("cell writes");
    worksheet
        .write_string(2, 0, "subtotal")
        .expect("cell writes");
    worksheet
        .group_rows(1, 1)
        .expect("expanded row outline is added");
    worksheet
        .group_rows(1, 2)
        .expect("outer expanded row outline is added");
    worksheet
        .group_columns(1, 2)
        .expect("expanded column outline is added");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("outlines.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("<Row outlineLevel={1}>"));
    assert!(
        source.contains("<Column min={2} max={3} widthPixels={64} outlineLevel={1} />"),
        "unexpected reverse output:\n{source}"
    );

    let regenerated = directory.path().join("outlines-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    let regenerated_file = std::fs::File::open(&regenerated).expect("workbook exists");
    let mut archive = ZipArchive::new(regenerated_file).expect("workbook is a ZIP package");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains("outlineLevel=\"1\""));
    assert!(worksheet_xml.contains("outlineLevel=\"2\""));

    let second_reverse = directory.path().join("outlines-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn public_outline_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("outlines.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/outlines.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("outlines.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("outlineLevel={1}"));
    assert!(source.contains("outlineLevel={2}"));
}

#[test]
fn public_collapsed_outline_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("collapsed-outlines.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/collapsed-outlines.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("collapsed-outlines.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("collapsed={true}"));
    assert!(source.contains("hidden={true} outlineLevel={1}"));
}

#[test]
fn public_outline_direction_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("outline-directions.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/outline-directions.tsx", "-o"])
        .arg(&output);
    command.assert().success();
    let xml = read_zip_part_for_test(&output, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("summaryBelow=\"0\""));
    assert!(xml.contains("summaryRight=\"0\""));

    let reversed = directory.path().join("outline-directions.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("summaryBelow={false}"));
    assert!(source.contains("summaryRight={false}"));
}

#[test]
fn public_worksheet_tab_color_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("worksheet-tab-color.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/worksheet-tab-color.tsx", "-o"])
        .arg(&output);
    command.assert().success();
    let xml = read_zip_part_for_test(&output, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("<tabColor rgb=\"FF4472C4\""));

    let reversed = directory.path().join("worksheet-tab-color.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("tabColor=\"#4472C4\""));
}

#[test]
fn public_active_worksheet_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("active-worksheet.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/active-worksheet.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("active-worksheet.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(
        source.contains("name={\"Overview\"} selected={true}"),
        "{source}"
    );
    assert!(source.contains("<Worksheet name={\"Details\"} active={true}"));
}

#[test]
fn public_checkbox_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("checkboxes.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/checkboxes.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let styles = read_zip_part_for_test(&output, "xl/styles.xml");
    assert!(styles.contains("xfComplement"), "{styles}");
    let feature_bag =
        read_zip_part_for_test(&output, "xl/featurePropertyBag/featurePropertyBag.xml");
    assert!(feature_bag.contains("type=\"Checkbox\""), "{feature_bag}");

    let reversed = directory.path().join("checkboxes.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("checkbox={true}"), "{source}");
}

#[test]
fn public_worksheet_protection_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("worksheet-protection.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/worksheet-protection.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("worksheet-protection.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("protection={{"));
    assert!(source.contains("locked={false}"));
    assert!(source.contains("formulaHidden={true}"));
}

#[test]
fn external_rgb_tab_color_round_trips_and_theme_colors_are_rejected() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("external-tab-color.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_tab_color(rust_xlsxwriter::Color::RGB(0x0012_ABEF));
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("external-tab-color.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("tabColor=\"#12ABEF\""));

    let theme_input = directory.path().join("theme-tab-color.xlsx");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        let modified = xml.replace("<tabColor rgb=\"FF12ABEF\"", "<tabColor theme=\"4\"");
        assert_ne!(modified, xml, "test must replace the RGB tab color");
        modified
    });
    std::fs::copy(&input, &theme_input).expect("theme test workbook copies");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(theme_input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("tabColor has unsupported"))
        .stderr(predicate::str::contains("theme"));
}

#[test]
fn external_hidden_and_very_hidden_worksheets_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("worksheet-visibility.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("Visible")
        .expect("visible sheet is named")
        .write_string(0, 0, "visible")
        .expect("cell writes");
    let hidden = workbook.add_worksheet();
    hidden.set_name("Hidden").expect("hidden sheet is named");
    hidden.set_hidden(true);
    hidden.write_string(0, 0, "hidden").expect("cell writes");
    let very_hidden = workbook.add_worksheet();
    very_hidden
        .set_name("VeryHidden")
        .expect("very-hidden sheet is named");
    very_hidden.set_very_hidden(true);
    very_hidden
        .write_string(0, 0, "very hidden")
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("worksheet-visibility.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("name={\"Hidden\"} visibility=\"hidden\""));
    assert!(source.contains("name={\"VeryHidden\"} visibility=\"veryHidden\""));

    let regenerated = directory.path().join("worksheet-visibility-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    let regenerated_file = std::fs::File::open(&regenerated).expect("workbook exists");
    let mut archive = ZipArchive::new(regenerated_file).expect("workbook is a ZIP package");
    let mut workbook_xml = String::new();
    archive
        .by_name("xl/workbook.xml")
        .expect("workbook XML exists")
        .read_to_string(&mut workbook_xml)
        .expect("workbook XML reads");
    assert!(workbook_xml.contains("state=\"hidden\""));
    assert!(workbook_xml.contains("state=\"veryHidden\""));

    let second_reverse = directory.path().join("worksheet-visibility-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_non_first_active_worksheet_round_trips_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("active-worksheet.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("First")
        .expect("first sheet names");
    workbook
        .add_worksheet()
        .set_name("Active")
        .expect("second sheet names")
        .set_active(true);
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("active-worksheet.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("<Worksheet name={\"Active\"} active={true}"));

    let regenerated = directory.path().join("active-worksheet-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let xml = read_zip_part_for_test(&regenerated, "xl/workbook.xml");
    assert!(xml.contains("activeTab=\"1\""));
    let second_reverse = directory.path().join("active-worksheet-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_selected_worksheet_tabs_round_trip_independently_from_active() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("selected-worksheets.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("First")
        .expect("first sheet names")
        .set_selected(true);
    workbook
        .add_worksheet()
        .set_name("Active but unselected")
        .expect("active sheet names")
        .set_active(true);
    workbook
        .add_worksheet()
        .set_name("Grouped")
        .expect("grouped sheet names")
        .set_selected(true);
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("selected-worksheets.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("name={\"Active but unselected\"} active={true}"),
        "{source}"
    );
    assert!(
        source.contains("name={\"First\"} selected={true}"),
        "{source}"
    );
    assert!(
        source.contains("name={\"Grouped\"} selected={true}"),
        "{source}"
    );

    let regenerated = directory.path().join("selected-worksheets-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    assert!(
        read_zip_part_for_test(&regenerated, "xl/workbook.xml").contains("activeTab=\"1\""),
        "{}",
        read_zip_part_for_test(&regenerated, "xl/workbook.xml")
    );
    let second_reverse = directory.path().join("selected-worksheets-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn reverse_rejects_active_but_unselected_non_first_worksheet() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("active-unselected.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("First")
        .expect("first sheet names");
    workbook
        .add_worksheet()
        .set_name("Active")
        .expect("second sheet names")
        .set_active(true);
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet2.xml", |xml| {
        xml.replace("tabSelected=\"1\"", "tabSelected=\"0\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "active but `sheetView@tabSelected` is false",
        ))
        .stderr(predicate::str::contains("cannot be recompiled losslessly"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_non_first_visible_tab_round_trips_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("first-tab.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("First")
        .expect("first sheet names");
    workbook
        .add_worksheet()
        .set_name("Leftmost")
        .expect("second sheet names")
        .set_first_tab(true);
    workbook
        .add_worksheet()
        .set_name("Last")
        .expect("third sheet names");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("first-tab.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("<Worksheet name={\"Leftmost\"} firstTab={true}"),
        "{source}"
    );

    let rebuilt = directory.path().join("first-tab-rebuilt.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let xml = read_zip_part_for_test(&rebuilt, "xl/workbook.xml");
    assert!(xml.contains("firstSheet=\"2\""), "{xml}");

    let second_reverse = directory.path().join("first-tab-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_checkbox_cells_reverse_rebuild_without_losing_feature_metadata() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("checkboxes.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .insert_checkbox(0, 0, true)
        .expect("checkbox writes");
    worksheet
        .insert_checkbox_with_format(
            1,
            0,
            false,
            &rust_xlsxwriter::Format::new().set_background_color("#C6EFCE"),
        )
        .expect("formatted checkbox writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("checkboxes.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("checkbox={true}"), "{source}");

    let rebuilt = directory.path().join("checkboxes-rebuilt.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let feature_bag =
        read_zip_part_for_test(&rebuilt, "xl/featurePropertyBag/featurePropertyBag.xml");
    assert!(feature_bag.contains("type=\"Checkbox\""), "{feature_bag}");

    let second_reverse = directory.path().join("checkboxes-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn reverse_rejects_unknown_checkbox_feature_property_bags_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-checkbox-bag.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .insert_checkbox(0, 0, true)
        .expect("checkbox writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(
        &input,
        "xl/featurePropertyBag/featurePropertyBag.xml",
        |xml| {
            let modified = xml.replace("type=\"Checkbox\"", "type=\"UnknownBag\"");
            assert_ne!(modified, xml, "test must change the checkbox bag type");
            modified
        },
    );

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("unsupported or duplicate bag"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_active_tab_outside_worksheet_order_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-active-tab.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("First")
        .expect("first sheet names");
    workbook
        .add_worksheet()
        .set_name("Second")
        .expect("second sheet names")
        .set_active(true);
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/workbook.xml", |xml| {
        let modified = xml.replace("activeTab=\"1\"", "activeTab=\"9\"");
        assert_ne!(modified, xml, "test must modify workbook activeTab");
        modified
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "activeTab 9 references a missing worksheet",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_first_sheet_outside_worksheet_order_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-first-sheet.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("First")
        .expect("first sheet names");
    workbook
        .add_worksheet()
        .set_name("Leftmost")
        .expect("second sheet names")
        .set_first_tab(true);
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/workbook.xml", |xml| {
        let modified = xml.replace("firstSheet=\"2\"", "firstSheet=\"9\"");
        assert_ne!(modified, xml, "test must modify workbook firstSheet");
        modified
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "firstSheet 9 references a missing worksheet",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_hidden_first_worksheet_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("hidden-first-worksheet.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let first = workbook.add_worksheet();
    first.write_string(0, 0, "first").expect("cell writes");
    workbook
        .add_worksheet()
        .write_string(0, 0, "second")
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/workbook.xml", |xml| {
        xml.replacen(
            "<sheet name=\"Sheet1\"",
            "<sheet state=\"hidden\" name=\"Sheet1\"",
            1,
        )
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("first worksheet must be visible"))
        .stderr(predicate::str::contains("move a visible worksheet first"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_collapsed_row_outline_round_trips_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("collapsed-outline.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "summary")
        .expect("cell writes");
    worksheet.write_string(1, 0, "detail").expect("cell writes");
    worksheet
        .group_rows_collapsed(1, 1)
        .expect("collapsed row outline is added");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("collapsed-outline.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("hidden={true} outlineLevel={1}"));
    assert!(source.contains("collapsed={true}"));

    let regenerated = directory.path().join("collapsed-outline-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let xml = read_zip_part_for_test(&regenerated, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("collapsed=\"1\""));
    assert!(xml.contains("hidden=\"1\" outlineLevel=\"1\""));
    let second_reverse = directory
        .path()
        .join("collapsed-outline-second-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn reverse_rejects_collapsed_row_marker_without_hidden_members() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-collapsed-row.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "summary")
        .expect("cell writes");
    worksheet.write_string(1, 0, "detail").expect("cell writes");
    worksheet
        .group_rows_collapsed(1, 1)
        .expect("collapsed row outline is added");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("hidden=\"1\"", "hidden=\"0\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("collapsed row summary"))
        .stderr(predicate::str::contains("hidden row"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_nested_collapsed_row_and_column_groups_round_trip() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("nested-collapsed-groups.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    for row in 0..6 {
        for column in 0..6 {
            worksheet
                .write_number(row, column, f64::from(row) * 10.0 + f64::from(column))
                .expect("cell writes");
        }
    }
    worksheet
        .group_rows_collapsed(1, 4)
        .expect("outer row group collapses");
    worksheet
        .group_rows_collapsed(1, 2)
        .expect("inner row group collapses");
    worksheet
        .group_columns_collapsed(1, 4)
        .expect("outer column group collapses");
    worksheet
        .group_columns_collapsed(1, 2)
        .expect("inner column group collapses");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("nested-collapsed-groups.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("outlineLevel={2}"));
    assert!(source.matches("collapsed={true}").count() >= 4);

    let regenerated = directory
        .path()
        .join("nested-collapsed-groups-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    let second_reverse = directory.path().join("nested-collapsed-groups-second.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_collapsed_column_outline_round_trips_without_loss() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("collapsed-columns.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "summary")
        .expect("cell writes");
    worksheet
        .write_string(0, 1, "detail 1")
        .expect("cell writes");
    worksheet
        .write_string(0, 2, "detail 2")
        .expect("cell writes");
    worksheet
        .group_columns_collapsed(1, 2)
        .expect("collapsed column outline is added");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("collapsed-columns.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("min={2} max={3} hidden={true} outlineLevel={1}"));
    assert!(
        source.contains("min={4} max={4} widthPixels={64} collapsed={true}"),
        "actual reverse output: {source}"
    );

    let regenerated = directory.path().join("collapsed-columns-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let xml = read_zip_part_for_test(&regenerated, "xl/worksheets/sheet1.xml");
    assert!(xml.contains("collapsed=\"1\""));
    assert!(xml.contains("hidden=\"1\""));
    assert!(xml.contains("outlineLevel=\"1\""));
    let second_reverse = directory
        .path()
        .join("collapsed-columns-second-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn reverse_rejects_collapsed_column_marker_spanning_multiple_columns() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-collapsed-column-range.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .group_columns_collapsed(1, 2)
        .expect("collapsed column outline is added");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        let collapsed_marker = xml
            .split("<col")
            .find(|column| column.contains("collapsed=\"1\""))
            .expect("writer emits a collapsed summary column");
        let min = collapsed_marker
            .split("min=\"")
            .nth(1)
            .and_then(|tail| tail.split('\"').next())
            .expect("collapsed summary has min");
        let max = collapsed_marker
            .split("max=\"")
            .nth(1)
            .and_then(|tail| tail.split('\"').next())
            .expect("collapsed summary has max");
        assert_eq!(min, max, "writer collapsed marker is a single column");
        let old = format!("max=\"{max}\"");
        let replacement = format!("max=\"{}\"", max.parse::<u16>().unwrap() + 1);
        let modified = collapsed_marker.replacen(&old, &replacement, 1);
        assert_ne!(
            modified, collapsed_marker,
            "test must alter collapsed range"
        );
        xml.replacen(collapsed_marker, &modified, 1)
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("collapsed column marker spans"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_column_width_and_hidden_state_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("columns.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_column_width_pixels(0, 140)
        .expect("first column width sets");
    worksheet
        .set_column_width_pixels(1, 160)
        .expect("second column width sets");
    worksheet
        .set_column_width_pixels(2, 160)
        .expect("third column width sets");
    worksheet.set_column_hidden(3).expect("fourth column hides");
    worksheet
        .write_string(0, 0, "wide")
        .expect("cell data writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("columns.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("<Column min={1} max={1} widthPixels={140} />"));
    assert!(source.contains("<Column min={2} max={3} widthPixels={160} />"));
    assert!(source.contains("<Column min={4} max={4} hidden={true} />"));

    let regenerated = directory.path().join("columns-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    let file = std::fs::File::open(&regenerated).expect("regenerated workbook opens");
    let mut archive = ZipArchive::new(file).expect("regenerated workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains("<col min=\"1\" max=\"1\" width=\"20\""));
    assert!(worksheet_xml.contains("<col min=\"2\" max=\"3\" width=\"22.85546875\""));
    assert!(worksheet_xml.contains("<col min=\"4\" max=\"4\" width=\"0\" hidden=\"1\""));

    let roundtrip = directory.path().join("columns-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn reverse_rejects_column_metadata_outside_the_lossless_subset() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("unsupported-column.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_column_width_pixels(0, 60)
        .expect("column width sets");
    worksheet
        .write_string(0, 0, "value")
        .expect("cell data writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("customWidth=\"1\"", "customWidth=\"1\" bestFit=\"1\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("unsupported `bestFit` metadata"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_column_width_that_cannot_be_represented_exactly() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("unrepresentable-column.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_column_width_pixels(0, 60)
        .expect("column width sets");
    worksheet
        .write_string(0, 0, "value")
        .expect("cell data writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("width=\"8.5703125\"", "width=\"8.43\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "cannot be represented exactly as integer pixels",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_preserves_leading_empty_row_positions() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("leading-rows.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string(2, 0, "third row")
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("leading-rows.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("<Row></Row>\n    <Row></Row>\n    <Row><Cell>"));

    let regenerated = directory.path().join("leading-rows-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut reread: calamine::Xlsx<_> =
        calamine::open_workbook(&regenerated).expect("recompiled workbook opens");
    let range = reread
        .worksheet_range("Sheet1")
        .expect("worksheet is readable");
    assert_eq!(
        range
            .get_value((2, 0))
            .and_then(calamine::DataType::get_string),
        Some("third row")
    );
}

#[test]
fn external_merged_ranges_reverse_and_recompile_without_coordinate_drift() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("external-merged.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let format = rust_xlsxwriter::Format::new()
        .set_bold()
        .set_align(rust_xlsxwriter::FormatAlign::Center);
    worksheet
        .merge_range(0, 0, 1, 1, "Merged", &format)
        .expect("merge range writes");
    worksheet.write_string(0, 3, "after").expect("cell writes");
    worksheet.write_string(1, 2, "below").expect("cell writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("external-merged.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("rowSpan={2} columnSpan={2}"));
    assert!(source.contains("after"));
    assert!(source.contains("below"));

    let regenerated = directory.path().join("external-merged-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive = ZipArchive::new(std::fs::File::open(&regenerated).expect("XLSX opens"))
        .expect("XLSX package opens");
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet part exists")
        .read_to_string(&mut sheet_xml)
        .expect("worksheet XML reads");
    assert!(sheet_xml.contains("<mergeCell ref=\"A1:B2\""));

    let roundtrip = directory.path().join("external-merged-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("roundtrip reverse output exists"),
        source
    );
}

#[test]
fn external_list_validation_reverses_and_recompiles_without_losing_options() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("list-validation.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let validation = rust_xlsxwriter::DataValidation::new()
        .allow_list_strings(&["Open", "Closed"])
        .expect("list validation is valid")
        .ignore_blank(false)
        .show_dropdown(false)
        .show_input_message(true)
        .show_error_message(true)
        .set_input_title("Choose status")
        .expect("input title is valid")
        .set_input_message("Select Open or Closed")
        .expect("input message is valid")
        .set_error_title("Invalid status")
        .expect("error title is valid")
        .set_error_message("Choose one listed value")
        .expect("error message is valid");
    worksheet
        .add_data_validation(1, 1, 9, 1, &validation)
        .expect("validation is applied");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("list-validation.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("type=\"list\""));
    assert!(source.contains("range={\"B2:B10\"}"));
    assert!(source.contains("allowBlank={false}"));
    assert!(source.contains("showDropDown={false}"));
    assert!(source.contains("inputTitle=\"Choose status\""));

    let regenerated = directory.path().join("list-validation-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("list-validation-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("roundtrip source exists"),
        source
    );
}

#[test]
fn external_any_validation_preserves_prompt_metadata_without_type_attribute() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("any-validation.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let validation = rust_xlsxwriter::DataValidation::new()
        .allow_any_value()
        .set_input_title("Optional note")
        .expect("valid prompt title")
        .set_input_message("Enter any value")
        .expect("valid prompt message");
    workbook
        .add_worksheet()
        .add_data_validation(0, 0, 2, 1, &validation)
        .expect("any validation writes");
    workbook.save(&input).expect("external workbook saves");

    let reversed = directory.path().join("any-validation.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("<DataValidation range={\"A1:B3\"} type=\"any\""));
    assert!(source.contains("inputTitle=\"Optional note\""));
    assert!(source.contains("inputMessage=\"Enter any value\""));

    let regenerated = directory.path().join("any-validation-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated)
        .assert()
        .success();
    let mut archive = ZipArchive::new(std::fs::File::open(regenerated).expect("round-trip opens"))
        .expect("round-trip workbook is zipped");
    let mut sheet = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet part exists")
        .read_to_string(&mut sheet)
        .expect("worksheet XML reads");
    assert!(sheet.contains("<dataValidation"));
    assert!(sheet.contains("promptTitle=\"Optional note\""));
    assert!(sheet.contains("prompt=\"Enter any value\""));
    assert!(!sheet.contains("<dataValidation type="));
}

#[test]
fn any_validation_rejects_formula_rules_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let source = directory.path().join("invalid-any-validation.tsx");
    std::fs::write(&source, "import { Workbook, Worksheet, DataValidation } from \"xlsx-jsx\"; export default <Workbook><Worksheet name=\"S\"><DataValidation range=\"A1:A3\" type=\"any\" formula1=\"1\" /></Worksheet></Workbook>;").expect("source writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(&source)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "type `any` does not accept formula1",
        ))
        .stderr(predicate::str::contains("remove the formula fields"));
}

#[test]
fn external_formula_validations_survive_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("whole-number-validation.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let validation = rust_xlsxwriter::DataValidation::new()
        .allow_whole_number(rust_xlsxwriter::DataValidationRule::Between(1, 5));
    let worksheet = workbook.add_worksheet();
    worksheet
        .add_data_validation(0, 0, 0, 0, &validation)
        .expect("whole-number validation is written");
    let decimal = rust_xlsxwriter::DataValidation::new().allow_decimal_number_formula(
        rust_xlsxwriter::DataValidationRule::GreaterThanOrEqualTo("=A1".into()),
    );
    worksheet
        .add_data_validation(0, 1, 0, 1, &decimal)
        .expect("decimal validation is written");
    let custom = rust_xlsxwriter::DataValidation::new().allow_custom("=MOD(A1,2)=0".into());
    worksheet
        .add_data_validation(0, 2, 0, 2, &custom)
        .expect("custom formula validation is written");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(&input);
    command.assert().success();
    let reversed = directory.path().join("whole-number-validation.tsx");
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("type=\"whole\" formula1=\"1\" formula2=\"5\""));
    assert!(source.contains("type=\"decimal\""), "{source}");
    assert!(
        source.contains("operator=\"greaterThanOrEqual\""),
        "{source}"
    );
    assert!(source.contains("formula1=\"A1\""), "{source}");
    assert!(source.contains("type=\"custom\" formula1=\"MOD(A1,2)=0\""));

    let output = directory.path().join("formula-validation-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(output).expect("roundtrip workbook opens"))
            .expect("roundtrip workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains("type=\"whole\""));
    assert!(worksheet_xml.contains("sqref=\"A1\""));
    assert!(worksheet_xml.contains("<formula1>1</formula1><formula2>5</formula2>"));
    assert!(worksheet_xml.contains("type=\"decimal\" operator=\"greaterThanOrEqual\""));
    assert!(worksheet_xml.contains("<formula1>MOD(A1,2)=0</formula1>"));
}

#[test]
fn public_list_validation_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("list-validation.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/list-validation.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("list-validation.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("type=\"list\""));
    assert!(source.contains("source={\"\\\"Open,Closed\\\"\"}"));
}

#[test]
fn public_formula_validation_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("formula-validation.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/formula-validation.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("formula-validation.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    for expected in [
        "type=\"whole\"",
        "type=\"decimal\"",
        "type=\"date\"",
        "type=\"time\"",
        "type=\"textLength\"",
        "type=\"custom\"",
    ] {
        assert!(source.contains(expected), "missing {expected} in {source}");
    }
}

#[test]
fn external_url_hyperlinks_survive_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("external-hyperlink.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_url_with_format(
            0,
            0,
            rust_xlsxwriter::Url::new("https://example.com/reports?q=x%20y#summary")
                .set_text("Open report"),
            &rust_xlsxwriter::Format::new(),
        )
        .expect("external hyperlink is written");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("external-hyperlink.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("hyperlink=\"https://example.com/reports?q=x%20y#summary\""));
    assert!(source.contains("Open report"));

    let output = directory.path().join("hyperlink-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(output).expect("roundtrip workbook opens"))
            .expect("roundtrip workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains("<hyperlink ref=\"A1\" r:id=\"rId1\""));
    let mut relationships_xml = String::new();
    archive
        .by_name("xl/worksheets/_rels/sheet1.xml.rels")
        .expect("worksheet hyperlink relationships exist")
        .read_to_string(&mut relationships_xml)
        .expect("relationship XML reads");
    assert!(relationships_xml.contains("https://example.com/reports?q=x%20y"));
}

#[test]
fn public_hyperlink_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("hyperlinks.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/hyperlinks.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("hyperlinks.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("hyperlink=\"https://example.com/reports?q=weekly\""));
    assert!(source.contains("hyperlink=\"mailto:reports@example.com\""));
    assert!(source.contains("hyperlink=\"internal:Reports!A1\""));
    assert!(source.contains("hyperlinkTooltip=\"View the latest report\""));
    assert!(source.contains("hyperlinkTooltip=\"Jump to the report sheet\""));
}

#[test]
fn internal_hyperlinks_reverse_and_recompile_without_relationships() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("internal-hyperlink.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_name("Sales Data")
        .expect("worksheet is named");
    workbook
        .add_worksheet()
        .set_name("Index")
        .expect("worksheet is named");
    workbook
        .worksheet_from_name("Index")
        .expect("Index worksheet exists")
        .write_url_with_format(
            0,
            0,
            rust_xlsxwriter::Url::new("internal:'Sales Data'!A1").set_text("Jump"),
            &rust_xlsxwriter::Format::new(),
        )
        .expect("internal hyperlink is written");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("internal-hyperlink.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reversed source exists");
    assert!(
        source.contains("hyperlink=\"internal:'Sales Data'!A1\""),
        "{source}"
    );

    let output = directory.path().join("internal-hyperlink-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output);
    command.assert().success();
    let second_reverse = directory
        .path()
        .join("internal-hyperlink-second-reverse.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse exists"),
        source
    );
}

#[test]
fn external_hyperlink_tooltips_survive_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("hyperlink-tooltip.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_url_with_format(
            0,
            0,
            rust_xlsxwriter::Url::new("https://example.com/reports")
                .set_text("Open report")
                .set_tip("View the weekly report"),
            &rust_xlsxwriter::Format::new(),
        )
        .expect("external tooltip hyperlink writes");
    worksheet
        .write_url_with_format(
            0,
            1,
            rust_xlsxwriter::Url::new("internal:Sheet1!A1")
                .set_text("Jump back")
                .set_tip("Return to the report link"),
            &rust_xlsxwriter::Format::new(),
        )
        .expect("internal tooltip hyperlink writes");
    workbook.save(&input).expect("external workbook writes");

    let reversed = directory.path().join("hyperlink-tooltip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("hyperlinkTooltip=\"View the weekly report\""));
    assert!(source.contains("hyperlinkTooltip=\"Return to the report link\""));

    let output = directory.path().join("hyperlink-tooltip-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(output).expect("round-trip workbook exists"))
            .expect("round-trip workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains("tooltip=\"View the weekly report\""));
    assert!(worksheet_xml.contains("tooltip=\"Return to the report link\""));
}

#[test]
fn reverse_rejects_internal_hyperlink_display_that_conflicts_with_cell_text() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("mismatched-internal-display.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_url_with_format(
            0,
            0,
            rust_xlsxwriter::Url::new("internal:Sheet1!A1").set_text("Cell value"),
            &rust_xlsxwriter::Format::new(),
        )
        .expect("internal hyperlink writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("display=\"Cell value\"", "display=\"Different display\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "hyperlink display text that differs",
        ))
        .stderr(predicate::str::contains("make both strings identical"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_unmodeled_hyperlink_metadata_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("hyperlink-tooltip.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_url_with_format(
            0,
            0,
            rust_xlsxwriter::Url::new("https://example.com").set_text("Open"),
            &rust_xlsxwriter::Format::new(),
        )
        .expect("tooltip hyperlink is written");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("ref=\"A1\"", "ref=\"A1\" customAttr=\"x\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(&input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "unsupported `customAttr` metadata",
        ))
        .stderr(predicate::str::contains("add matching JSX support"));
}

#[test]
fn external_color_scale_rules_survive_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("color-scales.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let two_color = rust_xlsxwriter::ConditionalFormat2ColorScale::new()
        .set_minimum(rust_xlsxwriter::ConditionalFormatType::Number, 10)
        .set_maximum(rust_xlsxwriter::ConditionalFormatType::Number, 90)
        .set_minimum_color("FFF2CC")
        .set_maximum_color("70AD47")
        .set_stop_if_true(true);
    worksheet
        .add_conditional_format(0, 0, 9, 0, &two_color)
        .expect("two-color scale is written");
    let three_color = rust_xlsxwriter::ConditionalFormat3ColorScale::new()
        .set_minimum_color("F8696B")
        .set_midpoint_color("FFEB84")
        .set_maximum_color("63BE7B")
        .set_midpoint(rust_xlsxwriter::ConditionalFormatType::Percentile, 60);
    worksheet
        .add_conditional_format(0, 1, 9, 1, &three_color)
        .expect("three-color scale is written");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("color-scales.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("type=\"colorScale2\""), "{source}");
    assert!(source.contains("minValue={10}"), "{source}");
    assert!(source.contains("maxValue={90}"), "{source}");
    assert!(source.contains("type=\"colorScale3\""), "{source}");
    assert!(source.contains("midType=\"percentile\""), "{source}");
    assert!(source.contains("midValue={60}"), "{source}");
    assert!(source.contains("stopIfTrue={true}"), "{source}");

    let output = directory.path().join("color-scales-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output);
    command.assert().success();
    let second_reverse = directory.path().join("color-scales-second-reverse.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse source exists"),
        source,
        "conditional formats should be stable across reverse/build/reverse"
    );
    let mut archive =
        ZipArchive::new(std::fs::File::open(output).expect("roundtrip workbook opens"))
            .expect("roundtrip workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains("type=\"colorScale\" priority=\"1\""));
    assert!(worksheet_xml.contains("<cfvo type=\"num\" val=\"10\""));
    assert!(worksheet_xml.contains("<cfvo type=\"percentile\" val=\"60\""));
}

#[test]
fn external_data_bar_rules_round_trip_with_paired_x14_extension() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("data-bars.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let data_bar = rust_xlsxwriter::ConditionalFormatDataBar::new()
        .set_minimum(rust_xlsxwriter::ConditionalFormatType::Number, 0)
        .set_maximum(rust_xlsxwriter::ConditionalFormatType::Number, 100)
        .set_fill_color("4472C4")
        .set_negative_fill_color("C00000")
        .set_border_color("305496")
        .set_negative_border_color("7030A0")
        .set_axis_color("222222")
        .set_direction(rust_xlsxwriter::ConditionalFormatDataBarDirection::RightToLeft)
        .set_axis_position(rust_xlsxwriter::ConditionalFormatDataBarAxisPosition::Midpoint)
        .set_solid_fill(true)
        .set_bar_only(true);
    worksheet
        .add_conditional_format(0, 0, 9, 0, &data_bar)
        .expect("data bar writes");
    worksheet
        .add_conditional_format(
            0,
            1,
            9,
            1,
            &rust_xlsxwriter::ConditionalFormatDataBar::new()
                .set_minimum(rust_xlsxwriter::ConditionalFormatType::Percentile, 10)
                .set_maximum(
                    rust_xlsxwriter::ConditionalFormatType::Formula,
                    rust_xlsxwriter::Formula::new("MAX(B1:B10)"),
                )
                .set_fill_color("70AD47"),
        )
        .expect("threshold data bar writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("data-bars.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("type=\"dataBar\""), "{source}");
    assert!(source.contains("fillColor=\"#4472C4\""), "{source}");
    assert!(source.contains("negativeFillColor=\"#C00000\""), "{source}");
    assert!(source.contains("borderColor=\"#305496\""), "{source}");
    assert!(
        source.contains("negativeBorderColor=\"#7030A0\""),
        "{source}"
    );
    assert!(source.contains("axisColor=\"#222222\""), "{source}");
    assert!(source.contains("direction=\"rightToLeft\""), "{source}");
    assert!(source.contains("axisPosition=\"midpoint\""), "{source}");
    assert!(
        source.contains("minType=\"percentile\" minValue={10}"),
        "{source}"
    );
    assert!(
        source.contains("maxType=\"formula\" maxValue=\"MAX(B1:B10)\""),
        "{source}"
    );
    assert!(source.contains("barOnly={true}"), "{source}");
    assert!(source.contains("solidFill={true}"), "{source}");

    let output = directory.path().join("data-bars-round-trip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(output).expect("round-trip workbook exists"))
            .expect("round-trip workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains("x14:cfRule"), "{worksheet_xml}");
    assert!(
        worksheet_xml.contains("negativeFillColor"),
        "{worksheet_xml}"
    );
}

#[test]
fn data_bar_rejects_unknown_direction_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("data-bar-invalid.tsx");
    std::fs::write(
        &input,
        r##"import { Workbook, Worksheet, ConditionalFormat } from "xlsx-jsx";
export default <Workbook><Worksheet name="Sheet1">
  <ConditionalFormat range="A1:A10" type="dataBar" fillColor="#4472C4" direction="diagonal" />
</Worksheet></Workbook>;
"##,
    )
    .expect("invalid fixture writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "direction `diagonal` is unsupported",
        ))
        .stderr(predicate::str::contains(
            "context, leftToRight, or rightToLeft",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn data_bar_uses_writer_default_color_when_fill_color_is_omitted() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("data-bar-default-color.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, ConditionalFormat } from "xlsx-jsx";
export default <Workbook><Worksheet name="Sheet1">
  <ConditionalFormat range="A1:A10" type="dataBar" />
</Worksheet></Workbook>;
"#,
    )
    .expect("data-bar fixture writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(&input)
        .assert()
        .success();
    let output = directory.path().join("data-bar-default-color.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&input)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
    let mut archive = ZipArchive::new(std::fs::File::open(output).expect("workbook exists"))
        .expect("workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(
        worksheet_xml.contains("rgb=\"FF638EC6\""),
        "{worksheet_xml}"
    );
}

#[test]
fn classic_data_bar_without_x14_extension_reverses_to_recompilable_jsx() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("classic-data-bar.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .add_conditional_format(
            0,
            0,
            9,
            0,
            &rust_xlsxwriter::ConditionalFormatDataBar::new()
                .set_fill_color("70AD47")
                .use_classic_style(),
        )
        .expect("classic data bar writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("classic-data-bar.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("type=\"dataBar\""), "{source}");
    assert!(source.contains("fillColor=\"#70AD47\""), "{source}");
    assert!(!source.contains("x14"), "{source}");

    let output = directory.path().join("classic-data-bar-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
}

#[test]
fn reverse_rejects_data_bar_with_mismatched_classic_and_x14_ids() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("mismatched-data-bar-id.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .add_conditional_format(
            0,
            0,
            9,
            0,
            &rust_xlsxwriter::ConditionalFormatDataBar::new().set_fill_color("4472C4"),
        )
        .expect("data bar writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |mut xml| {
        let marker = "x14:cfRule type=\"dataBar\" id=\"";
        let start = xml.find(marker).expect("x14 data bar rule exists") + marker.len();
        let end = xml[start..].find('"').expect("x14 rule id closes") + start;
        xml.replace_range(start..end, "{MISMATCHED-DATA-BAR-ID}");
        xml
    });

    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("has no matching x14 rule"))
        .stderr(predicate::str::contains("restore its x14 extension"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_cell_comparison_conditional_format_round_trips_with_differential_style() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("cell-rule.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_number(0, 0, 15)
        .expect("sample number writes");
    let rule = rust_xlsxwriter::ConditionalFormatCell::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatCellRule::Between(10, 20))
        .set_format(
            rust_xlsxwriter::Format::new()
                .set_bold()
                .set_font_color("9C0006")
                .set_background_color("FFC7CE"),
        )
        .set_stop_if_true(true);
    worksheet
        .add_conditional_format(0, 0, 9, 0, &rule)
        .expect("cell comparison conditional format writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("cell-rule.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("type=\"cellIs\""), "{source}");
    assert!(source.contains("operator=\"between\""), "{source}");
    assert!(source.contains("formula1=\"10\""), "{source}");
    assert!(source.contains("formula2=\"20\""), "{source}");
    assert!(source.contains("bold={true}"), "{source}");
    assert!(source.contains("fontColor=\"#9C0006\""), "{source}");
    assert!(source.contains("fillColor=\"#FFC7CE\""), "{source}");

    let output = directory.path().join("cell-rule-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output);
    command.assert().success();
    let second_reverse = directory.path().join("cell-rule-second-reverse.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse source exists"),
        source
    );
}

#[test]
fn external_formula_conditional_format_round_trips_with_differential_style() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("formula-rule.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let rule = rust_xlsxwriter::ConditionalFormatFormula::new()
        .set_rule("=MOD(A1,2)=0")
        .set_format(
            rust_xlsxwriter::Format::new()
                .set_italic()
                .set_font_strikethrough()
                .set_underline(rust_xlsxwriter::FormatUnderline::Double)
                .set_font_script(rust_xlsxwriter::FormatScript::Superscript)
                .set_background_color("DDEBF7"),
        )
        .set_stop_if_true(true);
    workbook
        .add_worksheet()
        .add_conditional_format(0, 0, 9, 0, &rule)
        .expect("formula conditional format writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("formula-rule.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("type=\"formula\""), "{source}");
    assert!(source.contains("formula1=\"MOD(A1,2)=0\""), "{source}");
    assert!(source.contains("italic={true}"), "{source}");
    assert!(source.contains("strikethrough={true}"), "{source}");
    assert!(source.contains("underline=\"double\""), "{source}");
    assert!(source.contains("script=\"superscript\""), "{source}");
    assert!(source.contains("fillColor=\"#DDEBF7\""), "{source}");
    assert!(source.contains("stopIfTrue={true}"), "{source}");

    let output = directory.path().join("formula-rule-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
    let second_reverse = directory.path().join("formula-rule-second-reverse.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&second_reverse)
        .assert()
        .success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse source exists"),
        source,
        "formula conditional format should remain stable across reverse/build/reverse"
    );
    let mut archive = ZipArchive::new(std::fs::File::open(output).expect("workbook opens"))
        .expect("workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(
        worksheet_xml.contains("type=\"expression\"") && worksheet_xml.contains("priority=\"1\"")
    );
    assert!(worksheet_xml.contains("<formula>MOD(A1,2)=0</formula>"));
}

#[test]
fn external_conditional_number_format_and_borders_round_trip() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("conditional-number-border.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let rule = rust_xlsxwriter::ConditionalFormatFormula::new()
        .set_rule("=A1>0")
        .set_format(
            rust_xlsxwriter::Format::new()
                .set_num_format("#,##0.00")
                .set_border_top(rust_xlsxwriter::FormatBorder::Thin)
                .set_border_top_color("123456")
                .set_border_left(rust_xlsxwriter::FormatBorder::Dashed),
        );
    workbook
        .add_worksheet()
        .add_conditional_format(0, 0, 9, 0, &rule)
        .expect("conditional format writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("conditional-number-border.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("numberFormat=\"#,##0.00\""), "{source}");
    assert!(source.contains("borderTop=\"thin\""), "{source}");
    assert!(source.contains("borderTopColor=\"#123456\""), "{source}");
    assert!(source.contains("borderLeft=\"dashed\""), "{source}");

    let output = directory
        .path()
        .join("conditional-number-border-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
    let mut archive = ZipArchive::new(std::fs::File::open(&output).expect("workbook opens"))
        .expect("workbook is zipped");
    let mut styles_xml = String::new();
    archive
        .by_name("xl/styles.xml")
        .expect("styles XML exists")
        .read_to_string(&mut styles_xml)
        .expect("styles XML reads");
    assert!(
        styles_xml.contains("formatCode=\"#,##0.00\""),
        "{styles_xml}"
    );
    assert!(styles_xml.contains("<border"), "{styles_xml}");
    let second_reverse = directory
        .path()
        .join("conditional-number-border-second.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&second_reverse)
        .assert()
        .success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse source exists"),
        source,
        "conditional number format and edge borders should remain stable across reverse/build/reverse"
    );
}

#[test]
fn formula_conditional_format_requires_supported_differential_style() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("formula-rule-invalid.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, ConditionalFormat } from "xlsx-jsx";
export default <Workbook><Worksheet name="Sheet1">
  <ConditionalFormat range="A1:A10" type="formula" formula1="=A1>0" />
</Worksheet></Workbook>;
"#,
    )
    .expect("invalid fixture writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "requires at least one differential style",
        ))
        .stderr(predicate::str::contains(
            "add a font style, numberFormat, or edge border",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_text_conditional_formats_round_trip_with_formula_verification() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("text-rules.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let text_rules = [
        (
            rust_xlsxwriter::ConditionalFormatTextRule::Contains("foo".to_owned()),
            rust_xlsxwriter::Format::new().set_background_color("DDEBF7"),
        ),
        (
            rust_xlsxwriter::ConditionalFormatTextRule::DoesNotContain("bar".to_owned()),
            rust_xlsxwriter::Format::new().set_bold(),
        ),
        (
            rust_xlsxwriter::ConditionalFormatTextRule::BeginsWith("a".to_owned()),
            rust_xlsxwriter::Format::new().set_italic(),
        ),
        (
            rust_xlsxwriter::ConditionalFormatTextRule::EndsWith("z".to_owned()),
            rust_xlsxwriter::Format::new().set_font_color("C00000"),
        ),
    ];
    for (rule, format) in &text_rules {
        worksheet
            .add_conditional_format(
                0,
                0,
                9,
                0,
                &rust_xlsxwriter::ConditionalFormatText::new()
                    .set_rule(rule.clone())
                    .set_format(format.clone()),
            )
            .expect("text conditional rule writes");
    }
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("text-rules.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(
        source.contains("type=\"text\" operator=\"containsText\" text=\"foo\""),
        "{source}"
    );
    assert!(
        source.contains("type=\"text\" operator=\"notContains\" text=\"bar\""),
        "{source}"
    );
    assert!(
        source.contains("type=\"text\" operator=\"beginsWith\" text=\"a\""),
        "{source}"
    );
    assert!(
        source.contains("type=\"text\" operator=\"endsWith\" text=\"z\""),
        "{source}"
    );

    let output = directory.path().join("text-rules-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
    let second_reverse = directory.path().join("text-rules-second-reverse.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&second_reverse)
        .assert()
        .success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse source exists"),
        source,
        "text conditional rules should remain stable across reverse/build/reverse"
    );
}

#[test]
fn text_conditional_format_rejects_unescaped_quote_text_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("text-rule-quote-invalid.tsx");
    std::fs::write(
        &input,
        r#"import { Workbook, Worksheet, ConditionalFormat } from "xlsx-jsx";
export default <Workbook><Worksheet name="Sheet1">
  <ConditionalFormat range="A1:A10" type="text" operator="containsText" text={'a"b'} bold={true} />
</Worksheet></Workbook>;
"#,
    )
    .expect("invalid fixture writes");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["validate"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "text cannot contain a double quote",
        ))
        .stderr(predicate::str::contains("remove the quote"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_text_conditional_formula_that_conflicts_with_rule_metadata() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("text-rule-mismatched-formula.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let rule = rust_xlsxwriter::ConditionalFormatText::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatTextRule::Contains(
            "urgent".to_owned(),
        ))
        .set_format(rust_xlsxwriter::Format::new().set_bold());
    workbook
        .add_worksheet()
        .add_conditional_format(0, 0, 9, 0, &rule)
        .expect("text conditional rule writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("NOT(ISERROR(SEARCH(\"urgent\",A1)))", "FALSE()")
    });

    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "formula does not match its operator/text/range",
        ))
        .stderr(predicate::str::contains(
            "restore the writer-generated formula",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn public_conditional_cell_rule_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("conditional-cell-rules.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/conditional-cell-rules.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("conditional-cell-rules.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse source exists");
    assert!(source.contains("type=\"cellIs\""), "{source}");
    assert!(source.contains("operator=\"between\""), "{source}");
    assert!(source.contains("fillColor=\"#C6EFCE\""), "{source}");
}

#[test]
fn public_conditional_formula_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("conditional-formula-rules.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build", "examples/conditional-formula-rules.tsx", "-o"])
        .arg(&output)
        .assert()
        .success();
    let reversed = directory.path().join("conditional-formula-rules.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(reversed).expect("reverse source exists");
    assert!(source.contains("type=\"formula\""), "{source}");
    assert!(source.contains("formula1=\"MOD(A2,2)=0\""), "{source}");
}

#[test]
fn public_conditional_text_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("conditional-text-rules.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build", "examples/conditional-text-rules.tsx", "-o"])
        .arg(&output)
        .assert()
        .success();
    let reversed = directory.path().join("conditional-text-rules.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(reversed).expect("reverse source exists");
    assert!(
        source.contains("type=\"text\" operator=\"containsText\" text=\"urgent\""),
        "{source}"
    );
    assert!(
        source.contains("type=\"text\" operator=\"beginsWith\" text=\"Done\""),
        "{source}"
    );
}

#[test]
fn public_page_break_example_builds_and_reverses() {
    let directory = tempdir().expect("temporary directory is available");
    let output = directory.path().join("page-breaks.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build", "examples/page-breaks.tsx", "-o"])
        .arg(&output);
    command.assert().success();

    let reversed = directory.path().join("page-breaks.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&output)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(reversed).expect("reverse source exists");
    assert!(
        source.contains("horizontalPageBreaks={[20, 40]}"),
        "{source}"
    );
    assert!(source.contains("verticalPageBreaks={[8]}"), "{source}");
}

#[test]
fn reverse_rejects_unsupported_cell_rule_differential_style_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("unsupported-cell-rule-style.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let rule = rust_xlsxwriter::ConditionalFormatCell::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatCellRule::GreaterThan(10))
        .set_format(
            rust_xlsxwriter::Format::new().set_border_diagonal(rust_xlsxwriter::FormatBorder::Thin),
        );
    workbook
        .add_worksheet()
        .add_conditional_format(0, 0, 4, 0, &rule)
        .expect("cell comparison writes");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("differential style outside"))
        .stderr(predicate::str::contains("simplify the dxf"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_malformed_data_bar_thresholds_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("unsupported-conditional-rule.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let rule = rust_xlsxwriter::ConditionalFormat2ColorScale::new();
    workbook
        .add_worksheet()
        .add_conditional_format(0, 0, 4, 0, &rule)
        .expect("color scale writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("type=\"colorScale\"", "type=\"dataBar\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "dataBar must contain two endpoint thresholds and one RGB color",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_icon_set_thresholds_round_trip_without_manifest() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("icon-set.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let icons = [
        rust_xlsxwriter::ConditionalFormatCustomIcon::new()
            .set_rule(rust_xlsxwriter::ConditionalFormatType::Percent, 0),
        rust_xlsxwriter::ConditionalFormatCustomIcon::new()
            .set_rule(rust_xlsxwriter::ConditionalFormatType::Number, 50)
            .set_greater_than(true),
        rust_xlsxwriter::ConditionalFormatCustomIcon::new()
            .set_rule(rust_xlsxwriter::ConditionalFormatType::Percentile, 90),
        rust_xlsxwriter::ConditionalFormatCustomIcon::new()
            .set_rule(rust_xlsxwriter::ConditionalFormatType::Number, 75),
        rust_xlsxwriter::ConditionalFormatCustomIcon::new()
            .set_rule(rust_xlsxwriter::ConditionalFormatType::Percent, 100),
    ];
    let rule = rust_xlsxwriter::ConditionalFormatIconSet::new()
        .set_icon_type(rust_xlsxwriter::ConditionalFormatIconType::FiveArrows)
        .reverse_icons(true)
        .show_icons_only(true)
        .set_stop_if_true(true)
        .set_icons(&icons);
    worksheet
        .add_conditional_format(0, 0, 9, 0, &rule)
        .expect("icon set writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("icon-set.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("type=\"iconSet\""), "{source}");
    assert!(source.contains("iconType=\"fiveArrows\""), "{source}");
    assert!(source.contains("reverse={true}"), "{source}");
    assert!(source.contains("showIconsOnly={true}"), "{source}");
    assert!(source.contains("stopIfTrue={true}"), "{source}");
    assert!(source.contains("value: 75"), "{source}");

    let output = directory.path().join("icon-set-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
}

#[test]
fn external_x14_icon_sets_round_trip_without_manifest() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("x14-icon-set.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let custom_icons = [
        rust_xlsxwriter::ConditionalFormatCustomIcon::new()
            .set_rule(rust_xlsxwriter::ConditionalFormatType::Percent, 0),
        rust_xlsxwriter::ConditionalFormatCustomIcon::new()
            .set_rule(rust_xlsxwriter::ConditionalFormatType::Percentile, 50)
            .set_icon_type(
                rust_xlsxwriter::ConditionalFormatIconType::FourHistograms,
                2,
            ),
        rust_xlsxwriter::ConditionalFormatCustomIcon::new()
            .set_rule(
                rust_xlsxwriter::ConditionalFormatType::Formula,
                rust_xlsxwriter::Formula::new("MAX(A1:A10)"),
            )
            .set_no_icon(true),
    ];
    let custom_rule = rust_xlsxwriter::ConditionalFormatIconSet::new()
        .set_icon_type(rust_xlsxwriter::ConditionalFormatIconType::ThreeArrows)
        .set_icons(&custom_icons);
    worksheet
        .add_conditional_format(0, 0, 9, 0, &custom_rule)
        .expect("custom icon set writes");
    let x14_only_rule = rust_xlsxwriter::ConditionalFormatIconSet::new()
        .set_icon_type(rust_xlsxwriter::ConditionalFormatIconType::ThreeStars);
    worksheet
        .add_conditional_format(0, 1, 9, 1, &x14_only_rule)
        .expect("x14-only icon set writes");
    workbook.save(&input).expect("workbook writes");
    let worksheet_xml = read_zip_part_for_test(&input, "xl/worksheets/sheet1.xml");
    assert!(worksheet_xml.contains("<x14:cfIcon iconSet=\"4Rating\" iconId=\"2\"/>"));

    let reversed = directory.path().join("x14-icon-set.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(
        source.contains("iconType: \"fourHistograms\", iconIndex: 2"),
        "{source}"
    );
    assert!(source.contains("noIcon: true"), "{source}");
    assert!(source.contains("MAX(A1:A10)"), "{source}");
    assert!(source.contains("iconType=\"threeStars\""), "{source}");

    let output = directory.path().join("x14-icon-set-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();

    let invalid_no_icon = directory.path().join("x14-icon-set-invalid-no-icon.xlsx");
    std::fs::copy(&input, &invalid_no_icon).expect("noIcon test workbook copies");
    rewrite_part(&invalid_no_icon, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "iconSet=\"NoIcons\" iconId=\"0\"",
            "iconSet=\"NoIcons\" iconId=\"1\"",
        )
    });
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&invalid_no_icon)
        .args(["-o"])
        .arg(directory.path().join("x14-icon-set-invalid-no-icon.tsx"))
        .assert()
        .failure()
        .stderr(predicate::str::contains("non-default iconId 1"));

    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "id=\"{DA7ABA51-AAAA-BBBB-0001-000000000002}\"",
            "id=\"{DA7ABA51-AAAA-BBBB-0001-000000000001}\"",
        )
    });
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(directory.path().join("x14-icon-set-invalid.tsx"))
        .assert()
        .failure()
        .stderr(predicate::str::contains("repeats x14 iconSet pairing id"));
}

#[test]
fn external_top_bottom_conditional_formats_round_trip_without_manifest() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("top-bottom.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let format = rust_xlsxwriter::Format::new()
        .set_font_color("9C0006")
        .set_background_color("FFC7CE");
    let rule = rust_xlsxwriter::ConditionalFormatTop::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatTopRule::BottomPercent(25))
        .set_format(&format)
        .set_stop_if_true(true);
    worksheet
        .add_conditional_format(0, 0, 9, 0, &rule)
        .expect("top/bottom rule writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("top-bottom.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("type=\"top10\""), "{source}");
    assert!(source.contains("rule=\"bottomPercent\""), "{source}");
    assert!(source.contains("rank={25}"), "{source}");
    assert!(source.contains("stopIfTrue={true}"), "{source}");
    assert!(source.contains("fontColor=\"#9C0006\""), "{source}");
    assert!(source.contains("fillColor=\"#FFC7CE\""), "{source}");

    let output = directory.path().join("top-bottom-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
}

#[test]
fn external_average_date_blank_error_and_duplicate_formats_round_trip_without_manifest() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("conditional-families.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let format = rust_xlsxwriter::Format::new()
        .set_font_color("9C0006")
        .set_background_color("FFC7CE");
    let average = rust_xlsxwriter::ConditionalFormatAverage::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatAverageRule::TwoStandardDeviationsBelow)
        .set_format(&format)
        .set_stop_if_true(true);
    worksheet
        .add_conditional_format(0, 0, 9, 0, &average)
        .expect("average rule writes");
    let date = rust_xlsxwriter::ConditionalFormatDate::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatDateRule::NextMonth);
    worksheet
        .add_conditional_format(0, 1, 9, 1, &date)
        .expect("date rule writes");
    let blank = rust_xlsxwriter::ConditionalFormatBlank::new()
        .invert()
        .set_format(&format);
    worksheet
        .add_conditional_format(0, 2, 9, 2, &blank)
        .expect("blank rule writes");
    let error = rust_xlsxwriter::ConditionalFormatError::new().set_format(&format);
    worksheet
        .add_conditional_format(0, 3, 9, 3, &error)
        .expect("error rule writes");
    let duplicate = rust_xlsxwriter::ConditionalFormatDuplicate::new().invert();
    worksheet
        .add_conditional_format(0, 4, 9, 4, &duplicate)
        .expect("duplicate rule writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("conditional-families.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    for expected in [
        "type=\"average\"",
        "rule=\"twoStdDevBelow\"",
        "type=\"date\"",
        "rule=\"nextMonth\"",
        "type=\"blank\"",
        "rule=\"nonBlank\"",
        "type=\"error\"",
        "type=\"duplicate\"",
        "rule=\"unique\"",
        "fillColor=\"#FFC7CE\"",
    ] {
        assert!(source.contains(expected), "missing {expected} in {source}");
    }

    let output = directory.path().join("conditional-families-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
    let original_xml = read_zip_part_for_test(&input, "xl/worksheets/sheet1.xml");
    let xml = read_zip_part_for_test(&output, "xl/worksheets/sheet1.xml");
    for expected in [
        "type=\"aboveAverage\"",
        "type=\"timePeriod\"",
        "type=\"notContainsBlanks\"",
        "type=\"containsErrors\"",
        "type=\"uniqueValues\"",
        "<formula>AND(MONTH(B1)=MONTH(TODAY())+1,OR(YEAR(B1)=YEAR(TODAY()),AND(MONTH(B1)=12,YEAR(B1)=YEAR(TODAY())+1)))</formula>",
    ] {
        assert!(xml.contains(expected), "missing {expected} in {xml}");
    }
    assert_eq!(formula_elements(&original_xml), formula_elements(&xml));
}

#[test]
#[expect(
    clippy::too_many_lines,
    reason = "This integration test exhaustively checks every public average and date rule against an external XLSX writer."
)]
fn all_external_average_and_date_rule_variants_preserve_their_meaning() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("average-date-rules.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let averages = [
        (
            "aboveAverage",
            rust_xlsxwriter::ConditionalFormatAverageRule::AboveAverage,
        ),
        (
            "belowAverage",
            rust_xlsxwriter::ConditionalFormatAverageRule::BelowAverage,
        ),
        (
            "equalOrAboveAverage",
            rust_xlsxwriter::ConditionalFormatAverageRule::EqualOrAboveAverage,
        ),
        (
            "equalOrBelowAverage",
            rust_xlsxwriter::ConditionalFormatAverageRule::EqualOrBelowAverage,
        ),
        (
            "oneStdDevAbove",
            rust_xlsxwriter::ConditionalFormatAverageRule::OneStandardDeviationAbove,
        ),
        (
            "oneStdDevBelow",
            rust_xlsxwriter::ConditionalFormatAverageRule::OneStandardDeviationBelow,
        ),
        (
            "twoStdDevAbove",
            rust_xlsxwriter::ConditionalFormatAverageRule::TwoStandardDeviationsAbove,
        ),
        (
            "twoStdDevBelow",
            rust_xlsxwriter::ConditionalFormatAverageRule::TwoStandardDeviationsBelow,
        ),
        (
            "threeStdDevAbove",
            rust_xlsxwriter::ConditionalFormatAverageRule::ThreeStandardDeviationsAbove,
        ),
        (
            "threeStdDevBelow",
            rust_xlsxwriter::ConditionalFormatAverageRule::ThreeStandardDeviationsBelow,
        ),
    ];
    for (index, (_, rule)) in averages.iter().enumerate() {
        let rule = rust_xlsxwriter::ConditionalFormatAverage::new().set_rule(rule.clone());
        let row = u32::try_from(index).expect("test range fits");
        worksheet
            .add_conditional_format(row, 0, row + 1, 0, &rule)
            .expect("average variant writes");
    }
    let dates = [
        (
            "yesterday",
            rust_xlsxwriter::ConditionalFormatDateRule::Yesterday,
        ),
        ("today", rust_xlsxwriter::ConditionalFormatDateRule::Today),
        (
            "tomorrow",
            rust_xlsxwriter::ConditionalFormatDateRule::Tomorrow,
        ),
        (
            "last7Days",
            rust_xlsxwriter::ConditionalFormatDateRule::Last7Days,
        ),
        (
            "lastWeek",
            rust_xlsxwriter::ConditionalFormatDateRule::LastWeek,
        ),
        (
            "thisWeek",
            rust_xlsxwriter::ConditionalFormatDateRule::ThisWeek,
        ),
        (
            "nextWeek",
            rust_xlsxwriter::ConditionalFormatDateRule::NextWeek,
        ),
        (
            "lastMonth",
            rust_xlsxwriter::ConditionalFormatDateRule::LastMonth,
        ),
        (
            "thisMonth",
            rust_xlsxwriter::ConditionalFormatDateRule::ThisMonth,
        ),
        (
            "nextMonth",
            rust_xlsxwriter::ConditionalFormatDateRule::NextMonth,
        ),
    ];
    for (index, (_, rule)) in dates.iter().enumerate() {
        let rule = rust_xlsxwriter::ConditionalFormatDate::new().set_rule(rule.clone());
        let row = u32::try_from(index).expect("test range fits");
        worksheet
            .add_conditional_format(row, 1, row + 1, 1, &rule)
            .expect("date variant writes");
    }
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("average-date-rules.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    for (rule, _) in &averages {
        if *rule != "aboveAverage" {
            assert!(source.contains(&format!("rule=\"{rule}\"")), "{source}");
        }
    }
    for (rule, _) in &dates {
        if *rule != "yesterday" {
            assert!(source.contains(&format!("rule=\"{rule}\"")), "{source}");
        }
    }
    let output = directory.path().join("average-date-rules-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output)
        .assert()
        .success();
    let original_xml = read_zip_part_for_test(&input, "xl/worksheets/sheet1.xml");
    let regenerated_xml = read_zip_part_for_test(&output, "xl/worksheets/sheet1.xml");
    assert_eq!(
        formula_elements(&original_xml),
        formula_elements(&regenerated_xml)
    );
}

#[test]
fn reverse_rejects_date_formula_that_disagrees_with_its_period() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("mismatched-date.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let rule = rust_xlsxwriter::ConditionalFormatDate::new()
        .set_rule(rust_xlsxwriter::ConditionalFormatDateRule::Yesterday);
    worksheet
        .add_conditional_format(0, 0, 9, 0, &rule)
        .expect("date rule writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("FLOOR(A1,1)=TODAY()-1", "FLOOR(A1,1)=TODAY()")
    });

    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("does not match expected formula"))
        .stderr(predicate::str::contains(
            "restore the canonical writer formula",
        ));
}

#[test]
fn reverse_rejects_row_outline_above_excel_limit() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("grouped-row.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string(0, 0, "grouped")
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("<row r=\"1\"", "<row r=\"1\" outlineLevel=\"8\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "outline level 8 above Excel's maximum of 7",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_unmodeled_border_metadata_instead_of_silently_dropping_it() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("formatted.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "bordered",
            &rust_xlsxwriter::Format::new().set_border(rust_xlsxwriter::FormatBorder::Thin),
        )
        .expect("formatted workbook is valid");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/styles.xml", |xml| {
        assert!(xml.contains("<border>"), "{xml}");
        xml.replace("<border>", "<border outline=\"0\">")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "outside the bold/italic/underline/strikethrough/script/font-name/size/family/charset/RGB-font-color/four-edge/diagonal-border/solid-fill/alignment/indent/wrapText/shrinkToFit/number-format subset",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_per_edge_borders_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("edge-borders.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "bordered",
            &rust_xlsxwriter::Format::new()
                .set_border_top(rust_xlsxwriter::FormatBorder::Thin)
                .set_border_top_color("#12ABEF")
                .set_border_right(rust_xlsxwriter::FormatBorder::Dotted)
                .set_border_right_color("#AABBCC")
                .set_border_bottom(rust_xlsxwriter::FormatBorder::Dashed)
                .set_border_left(rust_xlsxwriter::FormatBorder::Double),
        )
        .expect("edge borders write");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("edge-borders.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("borderTop={\"thin\"}"), "{source}");
    assert!(source.contains("borderTopColor={\"#12ABEF\"}"));
    assert!(source.contains("borderRight={\"dotted\"}"));
    assert!(source.contains("borderRightColor={\"#AABBCC\"}"));
    assert!(source.contains("borderBottom={\"dashed\"}"));
    assert!(source.contains("borderLeft={\"double\"}"));

    let regenerated = directory.path().join("edge-borders-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("edge-borders-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_diagonal_borders_survive_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("diagonal-borders.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    for (column, (label, direction)) in [
        (
            "upDown",
            rust_xlsxwriter::FormatDiagonalBorder::BorderUpDown,
        ),
        ("up", rust_xlsxwriter::FormatDiagonalBorder::BorderUp),
        ("down", rust_xlsxwriter::FormatDiagonalBorder::BorderDown),
    ]
    .into_iter()
    .enumerate()
    {
        worksheet
            .write_string_with_format(
                0,
                u16::try_from(column).expect("small column index"),
                label,
                &rust_xlsxwriter::Format::new()
                    .set_border_diagonal(rust_xlsxwriter::FormatBorder::Thin)
                    .set_border_diagonal_color("#123456")
                    .set_border_diagonal_type(direction),
            )
            .expect("diagonal border writes");
    }
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("diagonal-borders.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("borderDiagonal={\"thin\"}"), "{source}");
    assert!(
        source.contains("borderDiagonalColor={\"#123456\"}"),
        "{source}"
    );
    assert!(
        source.contains("borderDiagonalType={\"upDown\"}"),
        "{source}"
    );
    assert!(source.contains("borderDiagonalType={\"up\"}"), "{source}");
    assert!(source.contains("borderDiagonalType={\"down\"}"), "{source}");

    let rebuilt = directory.path().join("diagonal-borders-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let second_reverse = directory.path().join("diagonal-borders-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse exists"),
        source
    );
}

#[test]
fn reverse_rejects_theme_border_colors_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("theme-border-color.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "theme border",
            &rust_xlsxwriter::Format::new()
                .set_border_top(rust_xlsxwriter::FormatBorder::Thin)
                .set_border_top_color(rust_xlsxwriter::Color::Theme(4, 0)),
        )
        .expect("theme border color writes");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("theme/indexed border colors"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_accepts_writer_normalized_left_indentation() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("wrapped.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "indented",
            &rust_xlsxwriter::Format::new().set_indent(1),
        )
        .expect("indented cell writes");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    let output = command.assert().success().get_output().stdout.clone();
    let source_path = String::from_utf8(output).expect("CLI prints a UTF-8 output path");
    let source = std::fs::read_to_string(source_path.trim()).expect("reversed JSX exists");
    assert!(source.contains("indent={1}"));
    assert!(source.contains("horizontalAlign={\"left\"}"));
}

#[test]
fn external_cell_alignment_survives_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("alignment.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string_with_format(
            0,
            0,
            "center top",
            &rust_xlsxwriter::Format::new()
                .set_align(rust_xlsxwriter::FormatAlign::Center)
                .set_align(rust_xlsxwriter::FormatAlign::Top),
        )
        .expect("centered cell writes");
    worksheet
        .write_string_with_format(
            0,
            1,
            "left middle",
            &rust_xlsxwriter::Format::new()
                .set_align(rust_xlsxwriter::FormatAlign::Left)
                .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter),
        )
        .expect("left-aligned cell writes");
    worksheet
        .write_string_with_format(
            0,
            2,
            "right default vertical",
            &rust_xlsxwriter::Format::new().set_align(rust_xlsxwriter::FormatAlign::Right),
        )
        .expect("right-aligned cell writes");
    worksheet
        .write_formula_with_format(
            0,
            3,
            "=1+1",
            &rust_xlsxwriter::Format::new()
                .set_align(rust_xlsxwriter::FormatAlign::Center)
                .set_align(rust_xlsxwriter::FormatAlign::Bottom),
        )
        .expect("aligned formula writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("alignment.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("horizontalAlign={\"center\"} verticalAlign={\"top\"}"));
    assert!(source.contains("horizontalAlign={\"left\"} verticalAlign={\"center\"}"));
    assert!(source.contains("horizontalAlign={\"right\"}"));
    assert!(source.contains("formula={\"=1+1\"}"));

    let regenerated = directory.path().join("alignment-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    let roundtrip = directory.path().join("alignment-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_text_rotation_survives_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("rotated.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    for (column, angle) in [-45, 90, 270].into_iter().enumerate() {
        worksheet
            .write_string_with_format(
                0,
                u16::try_from(column).expect("small column"),
                "rotated",
                &rust_xlsxwriter::Format::new().set_rotation(angle),
            )
            .expect("rotated cell writes");
    }
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("rotated.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse source exists");
    assert!(source.contains("textRotation={-45}"));
    assert!(source.contains("textRotation={90}"));
    assert!(source.contains("textRotation={270}"));

    let regenerated = directory.path().join("rotated-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("rotated-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    let roundtrip_source = std::fs::read_to_string(roundtrip).expect("roundtrip source exists");
    assert!(roundtrip_source.contains("textRotation={-45}"));
    assert!(roundtrip_source.contains("textRotation={90}"));
    assert!(roundtrip_source.contains("textRotation={270}"));
}

#[test]
fn reverse_rejects_unrepresentable_text_rotation_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-rotation.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "invalid",
            &rust_xlsxwriter::Format::new().set_rotation(30),
        )
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/styles.xml", |xml| {
        xml.replace("textRotation=\"30\"", "textRotation=\"181\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "textRotation must be an OOXML angle",
        ))
        .stderr(predicate::str::contains("repair the alignment"));
}

#[test]
fn external_indentation_and_distributed_alignment_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("indent.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string_with_format(
            0,
            0,
            "indented left",
            &rust_xlsxwriter::Format::new()
                .set_align(rust_xlsxwriter::FormatAlign::Left)
                .set_indent(2),
        )
        .expect("left-indented cell writes");
    worksheet
        .write_string_with_format(
            1,
            0,
            "indented distributed",
            &rust_xlsxwriter::Format::new()
                .set_align(rust_xlsxwriter::FormatAlign::Distributed)
                .set_indent(3),
        )
        .expect("distributed-indented cell writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("indent.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("indent={2} horizontalAlign={\"left\"}"));
    assert!(source.contains("indent={3} horizontalAlign={\"distributed\"}"));

    let regenerated = directory.path().join("indent-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("indent-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("roundtrip exists"),
        source
    );
}

#[test]
fn external_autofilter_range_survives_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("autofilter.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_row(0, 0, ["Name", "Count"])
        .expect("headers write");
    worksheet
        .write_row(1, 0, ["Apples", "3"])
        .expect("data row writes");
    worksheet.autofilter(0, 0, 1, 1).expect("autofilter writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("autofilter.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("Worksheet name={\"Sheet1\"} autoFilter={\"A1:B2\"}"),
        "{source}"
    );

    let regenerated = directory.path().join("autofilter-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("autofilter-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse exists"),
        source
    );
}

#[test]
fn external_frozen_panes_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("frozen-panes.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Frozen header")
        .expect("header writes");
    worksheet
        .set_freeze_panes(1, 2)
        .expect("freeze panes are configured");
    worksheet
        .set_freeze_panes_top_cell(3, 4)
        .expect("scroll origin is configured");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("frozen-panes.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("freezeRows={1}"));
    assert!(source.contains("freezeColumns={2}"));
    assert!(source.contains("freezeTopLeftCell={\"E4\"}"));

    let regenerated = directory.path().join("frozen-panes-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("frozen-panes-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse exists"),
        source
    );
}

#[test]
fn external_single_axis_frozen_panes_survive_reverse() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("single-axis-freeze.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .set_freeze_panes(2, 0)
        .expect("row freeze is configured");
    workbook
        .add_worksheet()
        .set_freeze_panes(0, 3)
        .expect("column freeze is configured");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(&input);
    let output = command.assert().success().get_output().stdout.clone();
    let reversed = String::from_utf8(output).expect("CLI prints a UTF-8 path");
    let source = std::fs::read_to_string(reversed.trim()).expect("reverse output exists");
    assert!(source.contains("freezeRows={2}"));
    assert!(source.contains("freezeColumns={3}"));
}

#[test]
fn external_non_default_frozen_pane_selection_survives_reverse() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("selected-frozen-pane.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_freeze_panes(1, 1)
        .expect("freeze panes are configured");
    worksheet
        .set_selection(3, 3, 3, 3)
        .expect("selection is configured");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("selected-frozen-pane.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("selection={{activeCell: \"D4\", range: \"D4\"}}"));

    let regenerated = directory.path().join("selected-frozen-pane-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated)
        .assert()
        .success();
}

#[test]
fn external_autofilter_without_implementation_defined_name_survives_reverse() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("autofilter-without-name.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.autofilter(0, 0, 2, 1).expect("autofilter writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/workbook.xml", |mut xml| {
        if let Some(start) = xml.find("<definedNames>")
            && let Some(end) = xml.find("</definedNames>")
        {
            xml.replace_range(start..end + "</definedNames>".len(), "");
        }
        xml
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(&input);
    command.assert().success();
    let reversed = directory.path().join("autofilter-without-name.tsx");
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("autoFilter={\"A1:B3\"}"));
}

#[test]
fn reverse_preserves_external_string_autofilter_criteria() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("filtered-criteria.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_row(0, 0, ["Product", "Count"])
        .expect("headers write");
    worksheet
        .write_row(1, 0, ["Apples", "3"])
        .expect("data row writes");
    worksheet.autofilter(0, 0, 1, 1).expect("autofilter writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "<autoFilter ref=\"A1:B2\"/>",
            "<autoFilter ref=\"A1:B2\"><filterColumn colId=\"0\"><filters><filter val=\"Apples\"/></filters></filterColumn></autoFilter>",
        )
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command.assert().success();
    let reversed = directory.path().join("filtered-criteria.tsx");
    let source = std::fs::read_to_string(reversed).expect("reverse output exists");
    assert!(source.contains("autoFilterColumns={[{\"column\":0,\"values\":[\"Apples\"]}]}"));

    let output = directory.path().join("filtered-criteria-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(directory.path().join("filtered-criteria.tsx"));
    command.args(["-o"]).arg(&output);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(output).expect("roundtrip workbook opens"))
            .expect("roundtrip workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains(
        "<filterColumn colId=\"0\"><filters><filter val=\"Apples\"/></filters></filterColumn>"
    ));
}

#[test]
fn external_blank_and_custom_autofilters_survive_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("custom-filters.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_row(0, 0, ["Product", "Count"])
        .expect("headers write");
    worksheet
        .write_row(1, 0, ["Apples", "12"])
        .expect("data row writes");
    worksheet.autofilter(0, 0, 1, 1).expect("autofilter writes");
    let blank_filter = rust_xlsxwriter::FilterCondition::new()
        .add_list_filter("")
        .add_list_blanks_filter();
    worksheet
        .filter_column(0, &blank_filter)
        .expect("blank filter writes");
    let custom_filter = rust_xlsxwriter::FilterCondition::new()
        .add_custom_filter(rust_xlsxwriter::FilterCriteria::GreaterThanOrEqualTo, "10")
        .add_custom_filter(rust_xlsxwriter::FilterCriteria::LessThan, "20");
    worksheet
        .filter_column(1, &custom_filter)
        .expect("custom filter writes");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(&input);
    command.assert().success();
    let reversed = directory.path().join("custom-filters.tsx");
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains(
        "autoFilterColumns={[{\"column\":0,\"values\":[\"\"],\"includeBlanks\":true},{\"column\":1,\"customFilters\":{\"logic\":\"and\",\"conditions\":[{\"operator\":\"greaterThanOrEqual\",\"value\":\"10\"},{\"operator\":\"lessThan\",\"value\":\"20\"}]}}]}"
    ));

    let output = directory.path().join("custom-filters-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&output);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(output).expect("roundtrip workbook opens"))
            .expect("roundtrip workbook is zipped");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains(
        "<filters blank=\"1\"><filter val=\"\"/></filters></filterColumn><filterColumn colId=\"1\"><customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"10\"/><customFilter operator=\"lessThan\" val=\"20\"/></customFilters>"
    ));
}

#[test]
fn reverse_rejects_indent_with_incompatible_external_alignment() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("incompatible-indent.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "indented",
            &rust_xlsxwriter::Format::new()
                .set_align(rust_xlsxwriter::FormatAlign::Left)
                .set_indent(1),
        )
        .expect("indented cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/styles.xml", |xml| {
        xml.replace("horizontal=\"left\"", "horizontal=\"center\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "indent requires explicit left, right, or distributed",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
#[allow(clippy::too_many_lines)] // One external fixture exercises several number-format families and the full round trip.
fn external_custom_and_builtin_number_formats_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("number-formats.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_number_with_format(
            0,
            0,
            1234.5,
            &rust_xlsxwriter::Format::new().set_num_format("#,##0.00"),
        )
        .expect("decimal format writes");
    worksheet
        .write_number_with_format(
            0,
            1,
            0.125,
            &rust_xlsxwriter::Format::new().set_num_format("0.00%"),
        )
        .expect("percent format writes");
    worksheet
        .write_number_with_format(
            0,
            2,
            44927.5,
            &rust_xlsxwriter::Format::new().set_num_format("dd/mm/yyyy hh:mm"),
        )
        .expect("custom date format writes");
    worksheet
        .write_number_with_format(
            0,
            3,
            1.5,
            &rust_xlsxwriter::Format::new().set_num_format("[h]:mm:ss"),
        )
        .expect("elapsed time format writes");
    worksheet
        .write_formula_with_format(
            0,
            4,
            "=A1*2",
            &rust_xlsxwriter::Format::new().set_num_format("$#,##0.00"),
        )
        .expect("formatted formula writes");
    worksheet
        .write_number_with_format(
            0,
            5,
            12.0,
            &rust_xlsxwriter::Format::new().set_num_format_index(4),
        )
        .expect("built-in format writes");
    worksheet
        .write_number_with_format(
            0,
            6,
            44927.0,
            &rust_xlsxwriter::Format::new().set_num_format_index(14),
        )
        .expect("built-in date format writes");
    worksheet
        .write_number_with_format(
            0,
            7,
            1.25,
            &rust_xlsxwriter::Format::new().set_num_format_index(46),
        )
        .expect("built-in duration format writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("number-formats.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("numberFormat={\"#,##0.00\"}"));
    assert!(source.contains("numberFormat={\"0.00%\"}"));
    assert!(source.contains("numberFormat={\"dd/mm/yyyy hh:mm\"}"));
    assert!(source.contains("numberFormat={\"[h]:mm:ss\"}"));
    assert!(source.contains("numberFormat={\"m/d/yy\"}"));
    assert!(source.contains("formula={\"=A1*2\"}"));

    let regenerated = directory.path().join("number-formats-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    let roundtrip = directory.path().join("number-formats-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_solid_rgb_fills_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("solid-fills.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string_with_format(
            0,
            0,
            "filled",
            &rust_xlsxwriter::Format::new().set_background_color("#12ABEF"),
        )
        .expect("solid fill writes");
    worksheet
        .write_blank(
            0,
            1,
            &rust_xlsxwriter::Format::new()
                .set_background_color("#AABBCC")
                .set_bold(),
        )
        .expect("formatted blank writes");
    worksheet
        .write_formula_with_format(
            0,
            2,
            "=1+1",
            &rust_xlsxwriter::Format::new()
                .set_background_color("#FFEEDD")
                .set_num_format("0.00"),
        )
        .expect("formatted formula writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("solid-fills.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("fillColor={\"#12ABEF\"}"));
    assert!(source.contains("fillColor={\"#AABBCC\"}"));
    assert!(source.contains("bold={true}"));
    assert!(source.contains("fillColor={\"#FFEEDD\"}"));
    assert!(source.contains("numberFormat={\"0.00\"}"));

    let regenerated = directory.path().join("solid-fills-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("solid-fills-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_pattern_fills_survive_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("pattern-fill.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let format = rust_xlsxwriter::Format::new()
        .set_pattern(rust_xlsxwriter::FormatPattern::DarkDown)
        .set_foreground_color("#FF0000")
        .set_background_color("#FFFF00");
    worksheet
        .set_row_format(
            1,
            &rust_xlsxwriter::Format::new()
                .set_pattern(rust_xlsxwriter::FormatPattern::LightGrid)
                .set_foreground_color("#112233"),
        )
        .expect("row pattern format writes");
    worksheet
        .set_column_format(
            1,
            &rust_xlsxwriter::Format::new()
                .set_pattern(rust_xlsxwriter::FormatPattern::Gray125)
                .set_foreground_color("#ABCDEF"),
        )
        .expect("column pattern format writes");
    worksheet
        .write_string_with_format(0, 0, "pattern", &format)
        .expect("pattern fill writes");
    worksheet
        .write_string(1, 0, "row style")
        .expect("styled row content writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("pattern-fill.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("fillPattern=\"darkDown\""), "{source}");
    assert!(
        source.contains("fillForegroundColor=\"#FF0000\""),
        "{source}"
    );
    assert!(
        source.contains("fillBackgroundColor=\"#FFFF00\""),
        "{source}"
    );
    assert!(source.contains("fillPattern=\"lightGrid\""), "{source}");
    assert!(source.contains("fillPattern=\"gray125\""), "{source}");

    let regenerated = directory.path().join("pattern-fill-roundtrip.xlsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated)
        .assert()
        .success();
    let roundtrip = directory.path().join("pattern-fill-roundtrip.tsx");
    Command::cargo_bin("xlsx-jsx")
        .expect("binary is built")
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip)
        .assert()
        .success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_rgb_font_colors_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("font-colors.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string_with_format(
            0,
            0,
            "colored",
            &rust_xlsxwriter::Format::new().set_font_color("#12ABEF"),
        )
        .expect("font color writes");
    worksheet
        .write_string_with_format(
            0,
            1,
            "bold",
            &rust_xlsxwriter::Format::new()
                .set_font_color("#AABBCC")
                .set_bold(),
        )
        .expect("formatted blank writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("font-colors.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("fontColor={\"#12ABEF\"}"));
    assert!(source.contains("fontColor={\"#AABBCC\"}"));

    let regenerated = directory.path().join("font-colors-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("font-colors-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn reverse_rejects_theme_font_colors_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("theme-font-color.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "theme",
            &rust_xlsxwriter::Format::new().set_font_color(rust_xlsxwriter::Color::Theme(4, 0)),
        )
        .expect("theme font color writes");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("theme/indexed font colors"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_font_names_and_sizes_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("font-properties.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "font",
            &rust_xlsxwriter::Format::new()
                .set_font_name("Aptos Display")
                .set_font_size(12.5),
        )
        .expect("font name and size write");
    workbook.worksheets_mut()[0]
        .write_string_with_format(
            1,
            0,
            "size only",
            &rust_xlsxwriter::Format::new().set_font_size(14),
        )
        .expect("font size-only style writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("font-properties.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("fontName={\"Aptos Display\"}"));
    assert!(source.contains("fontSize={12.5}"));
    assert!(source.contains("fontSize={14}"));
    assert!(!source.contains("fontName={\"Calibri\"}"));

    let regenerated = directory.path().join("font-properties-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("font-properties-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_theme_font_schemes_survive_reverse_and_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("font-schemes.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string_with_format(
            0,
            0,
            "heading font",
            &rust_xlsxwriter::Format::new().set_font_scheme(rust_xlsxwriter::FontScheme::Headings),
        )
        .expect("heading font cell writes");
    worksheet
        .write_string_with_format(
            1,
            0,
            "body font",
            &rust_xlsxwriter::Format::new()
                .set_font_size(12)
                .set_font_scheme(rust_xlsxwriter::FontScheme::Body),
        )
        .expect("body font cell writes");
    worksheet
        .write_string_with_format(
            0,
            1,
            "non-theme font",
            &rust_xlsxwriter::Format::new()
                .set_font_size(13)
                .set_font_scheme(rust_xlsxwriter::FontScheme::None),
        )
        .expect("non-theme font writes");
    let plain_run = rust_xlsxwriter::Format::new();
    let heading_run =
        rust_xlsxwriter::Format::new().set_font_scheme(rust_xlsxwriter::FontScheme::Headings);
    worksheet
        .write_rich_string(0, 2, &[(&plain_run, "plain "), (&heading_run, "heading")])
        .expect("rich text scheme writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("font-schemes.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("fontScheme=\"headings\""), "{source}");
    assert!(source.contains("fontSize={12}"), "{source}");
    assert!(source.contains("fontScheme=\"none\""), "{source}");

    let regenerated = directory.path().join("font-schemes-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("font-schemes-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("second reverse output exists"),
        source
    );
}

#[test]
fn external_font_family_and_charset_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("font-metadata.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string_with_format(
            0,
            0,
            "font metadata",
            &rust_xlsxwriter::Format::new()
                .set_font_family(3)
                .set_font_charset(128),
        )
        .expect("formatted cell writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("font-metadata.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("fontFamily={3} fontCharset={128}"),
        "{source}"
    );

    let rebuilt = directory.path().join("font-metadata-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt);
    command.assert().success();
    let second_reverse = directory.path().join("font-metadata-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&rebuilt)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(second_reverse).expect("second reverse output exists"),
        source
    );
}

#[test]
fn reverse_rejects_font_sizes_outside_ooxml_range_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-font-size.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "invalid size",
            &rust_xlsxwriter::Format::new().set_font_size(0.5),
        )
        .expect("writer permits serializing font size for reverse validation");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "outside the bold/italic/underline/strikethrough/script/font-name/size",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_font_decorations_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("font-decorations.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    for (row, (label, underline)) in (0_u32..).zip([
        ("single", rust_xlsxwriter::FormatUnderline::Single),
        ("double", rust_xlsxwriter::FormatUnderline::Double),
        (
            "singleAccounting",
            rust_xlsxwriter::FormatUnderline::SingleAccounting,
        ),
        (
            "doubleAccounting",
            rust_xlsxwriter::FormatUnderline::DoubleAccounting,
        ),
    ]) {
        worksheet
            .write_string_with_format(
                row,
                0,
                label,
                &rust_xlsxwriter::Format::new().set_underline(underline),
            )
            .expect("underline writes");
    }
    worksheet
        .write_string_with_format(
            4,
            0,
            "strike",
            &rust_xlsxwriter::Format::new().set_font_strikethrough(),
        )
        .expect("strikethrough writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("font-decorations.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("underline={\"single\"}"));
    assert!(source.contains("underline={\"double\"}"));
    assert!(source.contains("underline={\"singleAccounting\"}"));
    assert!(source.contains("underline={\"doubleAccounting\"}"));
    assert!(source.contains("strikethrough={true}"));

    let regenerated = directory.path().join("font-decorations-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("font-decorations-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("roundtrip exists"),
        source
    );
}

#[test]
fn external_font_scripts_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("font-scripts.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string_with_format(
            0,
            0,
            "superscript",
            &rust_xlsxwriter::Format::new()
                .set_font_script(rust_xlsxwriter::FormatScript::Superscript),
        )
        .expect("superscript writes");
    worksheet
        .write_string_with_format(
            1,
            0,
            "subscript",
            &rust_xlsxwriter::Format::new()
                .set_font_script(rust_xlsxwriter::FormatScript::Subscript),
        )
        .expect("subscript writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("font-scripts.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("script={\"superscript\"}"));
    assert!(source.contains("script={\"subscript\"}"));

    let regenerated = directory.path().join("font-scripts-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("font-scripts-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("roundtrip exists"),
        source
    );
}

#[test]
fn external_text_wrapping_survives_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("wrapped.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "wrapped text",
            &rust_xlsxwriter::Format::new()
                .set_text_wrap()
                .set_align(rust_xlsxwriter::FormatAlign::Center),
        )
        .expect("wrapped cell writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("wrapped.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("wrapText={true}"));
    assert!(source.contains("horizontalAlign={\"center\"}"));

    let regenerated = directory.path().join("wrapped-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("wrapped-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("roundtrip exists"),
        source
    );
}

#[test]
fn external_shrink_to_fit_survives_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("shrink-to-fit.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "small text",
            &rust_xlsxwriter::Format::new()
                .set_shrink()
                .set_align(rust_xlsxwriter::FormatAlign::Right),
        )
        .expect("shrink-to-fit cell writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("shrink-to-fit.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("shrinkToFit={true}"));
    assert!(source.contains("horizontalAlign={\"right\"}"));

    let regenerated = directory.path().join("shrink-to-fit-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("shrink-to-fit-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    assert_eq!(
        std::fs::read_to_string(roundtrip).expect("roundtrip exists"),
        source
    );
}

#[test]
fn reverse_rejects_external_conflicting_wrap_and_shrink_flags() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("conflicting-alignment.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "wrapped",
            &rust_xlsxwriter::Format::new().set_text_wrap(),
        )
        .expect("wrapped cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/styles.xml", |xml| {
        xml.replace("wrapText=\"1\"", "wrapText=\"1\" shrinkToFit=\"1\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "combines wrapText and shrinkToFit",
        ))
        .stderr(predicate::str::contains("suppresses shrinkToFit"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_preserves_patterned_cell_fills() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("pattern-fill.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string_with_format(
            0,
            0,
            "pattern",
            &rust_xlsxwriter::Format::new()
                .set_background_color("#12ABEF")
                .set_pattern(rust_xlsxwriter::FormatPattern::DarkHorizontal),
        )
        .expect("pattern fill writes");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    let output = directory.path().join("pattern-fill.tsx");
    command
        .args(["reverse"])
        .arg(input)
        .args(["-o"])
        .arg(&output);
    command.assert().success();
    let source = std::fs::read_to_string(output).expect("reverse output exists");
    assert!(
        source.contains("fillPattern=\"darkHorizontal\""),
        "{source}"
    );
    assert!(
        source.contains("fillBackgroundColor=\"#12ABEF\""),
        "{source}"
    );
}

#[test]
fn reverse_rejects_unsupported_builtin_number_format_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("unknown-number-format.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_number_with_format(
            0,
            0,
            12.0,
            &rust_xlsxwriter::Format::new().set_num_format_index(23),
        )
        .expect("formatted value writes");
    workbook.save(&input).expect("workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("unsupported built-in formats"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_rejects_1904_date_system_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("date-system-1904.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_number(0, 0, 1)
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/workbook.xml", |xml| {
        xml.replace("<sheets>", "<workbookPr date1904=\"1\"/><sheets>")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("1904 date system"))
        .stderr(predicate::str::contains("convert dates to the 1900 system"))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn external_bold_and_italic_styles_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("font-emphasis.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string_with_format(0, 0, "bold", &rust_xlsxwriter::Format::new().set_bold())
        .expect("bold cell writes");
    worksheet
        .write_string_with_format(0, 1, "italic", &rust_xlsxwriter::Format::new().set_italic())
        .expect("italic cell writes");
    worksheet
        .write_number_with_format(0, 2, 42, &rust_xlsxwriter::Format::new().set_bold())
        .expect("formatted number writes");
    worksheet
        .write_formula_with_format(0, 3, "=C1+1", &rust_xlsxwriter::Format::new().set_italic())
        .expect("formatted formula writes");
    workbook.save(&input).expect("workbook writes");

    let reversed = directory.path().join("font-emphasis.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("bold={true}"));
    assert!(source.contains("italic={true}"));

    let regenerated = directory.path().join("font-emphasis-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();

    let roundtrip = directory.path().join("font-emphasis-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    let roundtrip_source =
        std::fs::read_to_string(roundtrip).expect("second reverse output exists");
    assert!(roundtrip_source.contains("bold={true}"));
    assert!(roundtrip_source.contains("italic={true}"));
}

#[test]
fn external_blank_cells_and_special_strings_survive_recompile() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("external.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "quote \" brace {} angle <&")
        .expect("write string");
    worksheet
        .write_blank(0, 1, &rust_xlsxwriter::Format::new())
        .expect("write blank");
    worksheet
        .write_string(0, 2, "line one\n中文🌊")
        .expect("write multiline string");
    workbook.save(&input).expect("external workbook writes");

    let reversed = directory.path().join("external.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("<Cell />"));
    assert!(source.contains("quote \\\" brace {} angle <&"));

    let regenerated = directory.path().join("external-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut reread: calamine::Xlsx<_> =
        calamine::open_workbook(&regenerated).expect("recompiled XLSX opens");
    let range = reread
        .worksheet_range("Sheet1")
        .expect("worksheet is readable");
    assert_eq!(
        range
            .get_value((0, 0))
            .and_then(calamine::DataType::get_string),
        Some("quote \" brace {} angle <&")
    );
    assert!(matches!(
        range.get_value((0, 1)),
        Some(calamine::Data::Empty)
    ));
    assert_eq!(
        range
            .get_value((0, 2))
            .and_then(calamine::DataType::get_string),
        Some("line one\n中文🌊")
    );
}

#[test]
fn external_header_footer_round_trip_without_manifest() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("external-header-footer.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_header("&CPage &[Page] of &[Pages] && report");
    worksheet.set_footer("&R&D");
    worksheet.set_header_footer_scale_with_doc(false);
    worksheet.set_header_footer_align_with_page(false);
    workbook.save(&input).expect("external workbook writes");

    let reversed = directory.path().join("header-footer.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("header=\"&CPage &P of &N && report\""),
        "{source}"
    );
    assert!(source.contains("footer=\"&R&D\""));
    assert!(source.contains("headerFooterScaleWithDoc={false}"));
    assert!(source.contains("headerFooterAlignWithPage={false}"));

    let regenerated = directory.path().join("header-footer-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(regenerated).expect("round-trip workbook exists"))
            .expect("round-trip workbook is a ZIP package");
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut sheet_xml)
        .expect("worksheet XML reads");
    assert!(sheet_xml.contains("scaleWithDoc=\"0\""));
    assert!(sheet_xml.contains("alignWithMargins=\"0\""));
    assert!(
        sheet_xml.contains("<oddHeader>&amp;CPage &amp;P of &amp;N &amp;&amp; report</oddHeader>")
    );
    assert!(sheet_xml.contains("<oddFooter>&amp;R&amp;D</oddFooter>"));
}

#[test]
fn external_cell_reading_direction_round_trips_and_recompiles() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("reading-direction.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let left_to_right = rust_xlsxwriter::Format::new().set_reading_direction(1);
    let right_to_left = rust_xlsxwriter::Format::new().set_reading_direction(2);
    worksheet
        .write_string_with_format(0, 0, "LTR", &left_to_right)
        .expect("LTR cell writes");
    worksheet
        .write_string_with_format(0, 1, "RTL", &right_to_left)
        .expect("RTL cell writes");
    workbook.save(&input).expect("external workbook writes");

    let reversed = directory.path().join("reading-direction.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(
        source.contains("readingDirection=\"leftToRight\""),
        "{source}"
    );
    assert!(
        source.contains("readingDirection=\"rightToLeft\""),
        "{source}"
    );

    let regenerated = directory.path().join("reading-direction-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(regenerated).expect("round-trip workbook exists"))
            .expect("round-trip workbook is a ZIP package");
    let mut styles_xml = String::new();
    archive
        .by_name("xl/styles.xml")
        .expect("styles XML exists")
        .read_to_string(&mut styles_xml)
        .expect("styles XML reads");
    assert!(styles_xml.contains("readingOrder=\"1\""));
    assert!(styles_xml.contains("readingOrder=\"2\""));
}

#[test]
fn external_quote_prefix_round_trips_and_recompiles() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("external-quote-prefix.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let quote_prefix = rust_xlsxwriter::Format::new().set_quote_prefix();
    worksheet
        .write_string_with_format(0, 0, "=SUM(1, 2)", &quote_prefix)
        .expect("literal formula-like text writes");
    workbook.save(&input).expect("external workbook writes");

    let reversed = directory.path().join("quote-prefix.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("quotePrefix={true}"), "{source}");
    assert!(source.contains("=SUM(1, 2)"), "{source}");

    let regenerated = directory.path().join("quote-prefix-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(regenerated).expect("round-trip workbook exists"))
            .expect("round-trip workbook is a ZIP package");
    let mut styles_xml = String::new();
    archive
        .by_name("xl/styles.xml")
        .expect("styles XML exists")
        .read_to_string(&mut styles_xml)
        .expect("styles XML reads");
    assert!(styles_xml.contains("quotePrefix=\"1\""));
}

#[test]
fn reverse_rejects_invalid_quote_prefix_boolean_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-quote-prefix.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let quote_prefix = rust_xlsxwriter::Format::new().set_quote_prefix();
    worksheet
        .write_string_with_format(0, 0, "=SUM(1, 2)", &quote_prefix)
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/styles.xml", |xml| {
        xml.replace("quotePrefix=\"1\"", "quotePrefix=\"maybe\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("quotePrefix"))
        .stderr(predicate::str::contains("boolean"));
}

#[test]
fn external_worksheet_protection_and_cell_flags_round_trip() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("external-protection.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let options = rust_xlsxwriter::ProtectionOptions {
        select_locked_cells: false,
        select_unlocked_cells: false,
        format_cells: true,
        format_columns: true,
        format_rows: true,
        insert_columns: true,
        insert_rows: true,
        insert_links: true,
        delete_columns: true,
        delete_rows: true,
        sort: true,
        use_autofilter: true,
        use_pivot_tables: true,
        edit_scenarios: true,
        edit_objects: true,
        ..Default::default()
    };
    worksheet.protect_with_options(&options);
    let unlocked = rust_xlsxwriter::Format::new().set_unlocked();
    let hidden_formula = rust_xlsxwriter::Format::new().set_hidden();
    worksheet
        .write_string_with_format(0, 0, "editable", &unlocked)
        .expect("unlocked cell writes");
    worksheet
        .write_formula_with_format(0, 1, "=1+1", &hidden_formula)
        .expect("hidden formula writes");
    workbook.save(&input).expect("external workbook writes");

    let reversed = directory.path().join("protection.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    for option in [
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
    ] {
        assert!(source.contains(option), "missing {option} in {source}");
    }
    assert!(source.contains("locked={false}"), "{source}");
    assert!(source.contains("formulaHidden={true}"), "{source}");

    let regenerated = directory.path().join("protection-round-trip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let mut archive =
        ZipArchive::new(std::fs::File::open(regenerated).expect("round-trip workbook exists"))
            .expect("round-trip workbook is a ZIP package");
    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet XML exists")
        .read_to_string(&mut worksheet_xml)
        .expect("worksheet XML reads");
    assert!(worksheet_xml.contains("sheet=\"1\""));
    assert!(worksheet_xml.contains("formatCells=\"0\""));
    assert!(worksheet_xml.contains("sort=\"0\""));

    let mut styles_xml = String::new();
    archive
        .by_name("xl/styles.xml")
        .expect("styles XML exists")
        .read_to_string(&mut styles_xml)
        .expect("styles XML reads");
    assert!(styles_xml.contains("<protection locked=\"0\"/>"));
    assert!(styles_xml.contains("<protection hidden=\"1\"/>"));
}

#[test]
fn reverse_rejects_password_protected_worksheet_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("password-protected.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .protect_with_password("not-a-secret");
    workbook.save(&input).expect("protected workbook writes");

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("password-based sheet protection"))
        .stderr(predicate::str::contains(
            "original password cannot be recovered",
        ));
}

#[test]
fn reverse_rejects_invalid_cell_reading_order_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid-reading-order.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    let format = rust_xlsxwriter::Format::new().set_reading_direction(1);
    worksheet
        .write_string_with_format(0, 0, "LTR", &format)
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/styles.xml", |xml| {
        xml.replace("readingOrder=\"1\"", "readingOrder=\"3\"")
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("readingOrder `3`"))
        .stderr(predicate::str::contains("left-to-right"));
}

#[test]
fn reverse_rejects_even_page_header_with_repair_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("even-header.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook.add_worksheet();
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "</worksheet>",
            "<headerFooter><evenHeader>alternate</evenHeader></headerFooter></worksheet>",
        )
    });
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("evenHeader"))
        .stderr(predicate::str::contains("only oddHeader and oddFooter"));
}

#[test]
fn reverse_reports_excel_error_cell_with_coordinate_and_repair() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("error-cell.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_number(0, 0, 1)
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "<c r=\"A1\"><v>1</v></c>",
            "<c r=\"A1\" t=\"e\"><v>#DIV/0!</v></c>",
        )
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("cell `A1`"))
        .stderr(predicate::str::contains("#DIV/0!"))
        .stderr(predicate::str::contains("fix the formula/error"));
}

#[test]
fn reverse_rejects_timezone_bearing_iso_date_with_guidance() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("date-cell.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_number(0, 0, 1)
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "<c r=\"A1\"><v>1</v></c>",
            "<c r=\"A1\" t=\"d\"><v>2026-10-10T00:00:00Z</v></c>",
        )
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("cell `A1`"))
        .stderr(predicate::str::contains("time-zone"))
        .stderr(predicate::str::contains("local ISO date/time"));
}

#[test]
fn reverse_preserves_timezone_free_iso_date_cells() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("iso-date-cell.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string(0, 0, "placeholder")
        .expect("cell writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "<c r=\"A1\" t=\"s\"><v>0</v></c>",
            "<c r=\"A1\" t=\"d\"><v>2026-10-10</v></c>",
        )
    });

    let reversed = directory.path().join("iso-date-cell.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("date={\"2026-10-10\"}"));

    let regenerated = directory.path().join("iso-date-cell-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("iso-date-cell-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    let source = std::fs::read_to_string(roundtrip).expect("roundtrip source exists");
    assert!(source.contains("date={\"2026-10-10\"}"));
}

#[test]
fn external_date_and_datetime_cells_round_trip() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("dates.xlsx");
    let date = rust_xlsxwriter::ExcelDateTime::parse_from_str("2026-10-10").expect("date is valid");
    let datetime = rust_xlsxwriter::ExcelDateTime::parse_from_str("2026-10-10T14:30:12.125")
        .expect("datetime is valid");
    let datetime_seconds = rust_xlsxwriter::ExcelDateTime::parse_from_str("2026-10-10T14:30:12")
        .expect("second precision datetime is valid");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_datetime_with_format(
            0,
            0,
            &date,
            &rust_xlsxwriter::Format::new().set_num_format("yyyy-mm-dd"),
        )
        .expect("date cell writes");
    worksheet
        .write_datetime_with_format(
            0,
            1,
            &datetime,
            &rust_xlsxwriter::Format::new().set_num_format("yyyy-mm-dd hh:mm:ss.000"),
        )
        .expect("datetime cell writes");
    worksheet
        .write_datetime_with_format(
            0,
            2,
            &datetime_seconds,
            &rust_xlsxwriter::Format::new().set_num_format("yyyy-mm-dd hh:mm:ss"),
        )
        .expect("second precision datetime cell writes");
    workbook.save(&input).expect("workbook writes");
    let reversed = directory.path().join("dates.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&input)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("reverse output exists");
    assert!(source.contains("date={\"2026-10-10\"}"));
    assert!(source.contains("date={\"2026-10-10T14:30:12.125\"}"));
    assert!(source.contains("date={\"2026-10-10T14:30:12\"}"));

    let regenerated = directory.path().join("dates-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let roundtrip = directory.path().join("dates-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&roundtrip);
    command.assert().success();
    let roundtrip_source =
        std::fs::read_to_string(roundtrip).expect("second reverse output exists");
    assert!(roundtrip_source.contains("date={\"2026-10-10\"}"));
    assert!(roundtrip_source.contains("date={\"2026-10-10T14:30:12.125\"}"));
    assert!(roundtrip_source.contains("date={\"2026-10-10T14:30:12\"}"));
}

#[test]
fn reverse_preserves_rich_text_and_rejects_unknown_package_parts() {
    let directory = tempdir().expect("temporary directory is available");
    let rich_text = directory.path().join("rich-text.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_string(0, 0, "plain")
        .expect("cell writes");
    workbook.save(&rich_text).expect("workbook writes");
    rewrite_part(&rich_text, "xl/sharedStrings.xml", |xml| {
        xml.replace(
            "<si><t>plain</t></si>",
            "<si><r><rPr><b/></rPr><t>rich</t></r></si>",
        )
    });
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    let reversed = directory.path().join("rich-text.tsx");
    command
        .args(["reverse"])
        .arg(&rich_text)
        .args(["-o"])
        .arg(&reversed);
    command.assert().success();
    let source = std::fs::read_to_string(&reversed).expect("rich text reverse output exists");
    assert!(
        source.contains("<TextRun bold={true}>{\"rich\"}</TextRun>"),
        "{source}"
    );
    let regenerated = directory.path().join("rich-text-roundtrip.xlsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["build"])
        .arg(&reversed)
        .args(["-o"])
        .arg(&regenerated);
    command.assert().success();
    let second_reverse = directory.path().join("rich-text-roundtrip.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(&regenerated)
        .args(["-o"])
        .arg(&second_reverse);
    command.assert().success();
    let second_source =
        std::fs::read_to_string(second_reverse).expect("second rich text reverse exists");
    assert!(
        second_source.contains("<TextRun bold={true} fontSize={11}>{\"rich\"}</TextRun>"),
        "{second_source}"
    );

    let unknown_part = directory.path().join("unknown-part.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_number(0, 0, 1)
        .expect("cell writes");
    workbook.save(&unknown_part).expect("workbook writes");
    add_part(&unknown_part, "xl/customData/custom.xml", "<customData/>");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(&unknown_part);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("not recognized by the v1 reader"));
}

#[test]
fn reverse_rejects_shared_formula_structures() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("shared-formula.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_formula(0, 0, "=SUM(1,2)")
        .expect("formula writes");
    workbook.save(&input).expect("workbook writes");
    rewrite_part(&input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace(
            "<f>SUM(1,2)</f>",
            "<f t=\"shared\" si=\"0\" ref=\"A1:A2\">SUM(1,2)</f>",
        )
    });

    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["reverse"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains(
            "shared, data-table, or unsupported formula type",
        ))
        .stderr(predicate::str::contains("suggestion:"));
}

#[test]
fn reverse_preserves_zoom_and_supported_authored_metadata() {
    let directory = tempdir().expect("temporary directory is available");
    let view_input = directory.path().join("zoom.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_number(0, 0, 1)
        .expect("cell writes");
    workbook.save(&view_input).expect("workbook writes");
    rewrite_part(&view_input, "xl/worksheets/sheet1.xml", |xml| {
        xml.replace("tabSelected=\"1\"", "tabSelected=\"1\" zoomScale=\"120\"")
    });

    let view_output = directory.path().join("zoom.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(view_input)
        .args(["-o"])
        .arg(&view_output);
    command.assert().success();
    let source = std::fs::read_to_string(view_output).expect("reverse output exists");
    assert!(source.contains("zoom={120}"));

    let metadata_input = directory.path().join("metadata.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    workbook
        .add_worksheet()
        .write_number(0, 0, 1)
        .expect("cell writes");
    workbook.save(&metadata_input).expect("workbook writes");
    rewrite_part(&metadata_input, "docProps/core.xml", |xml| {
        xml.replace(
            "</cp:coreProperties>",
            "<dc:title>Quarterly report</dc:title></cp:coreProperties>",
        )
    });

    let metadata_output = directory.path().join("metadata.tsx");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command
        .args(["reverse"])
        .arg(metadata_input)
        .args(["-o"])
        .arg(&metadata_output);
    command.assert().success();
    let source = std::fs::read_to_string(metadata_output).expect("reverse output exists");
    assert!(source.contains("title=\"Quarterly report\""));
}

fn rewrite_part(path: &std::path::Path, part_name: &str, transform: impl Fn(String) -> String) {
    let rewritten = path.with_extension("rewritten.xlsx");
    let input = std::fs::File::open(path).expect("source package opens");
    let mut archive = ZipArchive::new(input).expect("source package is ZIP");
    let output = std::fs::File::create(&rewritten).expect("rewrite package can be created");
    let mut writer = ZipWriter::new(output);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let mut found = false;
    for index in 0..archive.len() {
        let mut part = archive.by_index(index).expect("package part is readable");
        let name = part.name().to_owned();
        let mut bytes = Vec::new();
        part.read_to_end(&mut bytes).expect("package part reads");
        if name == part_name {
            let text = String::from_utf8(bytes).expect("XML part is UTF-8");
            bytes = transform(text).into_bytes();
            found = true;
        }
        writer
            .start_file(name, options)
            .expect("package part starts");
        writer.write_all(&bytes).expect("package part writes");
    }
    assert!(found, "the requested package part existed");
    writer.finish().expect("rewritten package closes");
    std::fs::rename(rewritten, path).expect("rewritten package replaces test fixture");
}

fn read_zip_part_for_test(path: &std::path::Path, part_name: &str) -> String {
    let mut archive = ZipArchive::new(std::fs::File::open(path).expect("workbook opens"))
        .expect("workbook is a ZIP package");
    let mut xml = String::new();
    archive
        .by_name(part_name)
        .expect("requested package part exists")
        .read_to_string(&mut xml)
        .expect("package part reads");
    xml
}

fn formula_elements(xml: &str) -> Vec<&str> {
    xml.split("<formula>")
        .skip(1)
        .filter_map(|remainder| {
            remainder
                .split_once("</formula>")
                .map(|(formula, _)| formula)
        })
        .collect()
}

fn add_part(path: &std::path::Path, part_name: &str, contents: &str) {
    let rewritten = path.with_extension("added.xlsx");
    let input = std::fs::File::open(path).expect("source package opens");
    let mut archive = ZipArchive::new(input).expect("source package is ZIP");
    let output = std::fs::File::create(&rewritten).expect("package can be created");
    let mut writer = ZipWriter::new(output);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for index in 0..archive.len() {
        let mut part = archive.by_index(index).expect("package part is readable");
        let name = part.name().to_owned();
        let mut bytes = Vec::new();
        part.read_to_end(&mut bytes).expect("package part reads");
        writer
            .start_file(name, options)
            .expect("package part starts");
        writer.write_all(&bytes).expect("package part writes");
    }
    writer
        .start_file(part_name, options)
        .expect("custom package part starts");
    writer
        .write_all(contents.as_bytes())
        .expect("custom part writes");
    writer.finish().expect("rewritten package closes");
    std::fs::rename(rewritten, path).expect("rewritten package replaces test fixture");
}

#[test]
fn invalid_dsl_has_reason_and_repair_suggestion() {
    let directory = tempdir().expect("temporary directory is available");
    let input = directory.path().join("invalid.tsx");
    std::fs::write(&input, r#"import { Workbook, Worksheet } from "xlsx-jsx"; export default <Workbook><Worksheet name="Bad/Name" /></Workbook>;"#).expect("fixture writes");
    let mut command = Command::cargo_bin("xlsx-jsx").expect("binary is built");
    command.args(["validate"]).arg(input);
    command
        .assert()
        .failure()
        .stderr(predicate::str::contains("explanation:"))
        .stderr(predicate::str::contains("suggestion:"))
        .stderr(predicate::str::contains("Worksheet name"));
}
