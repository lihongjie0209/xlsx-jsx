# xlsx-jsx

Rust CLI for compiling executable JSX/TSX workbook descriptions to `.xlsx` and
reversing the documented lossless subset of external spreadsheets back to JSX.
The frontend runs in embedded V8; no Node.js runtime is required.

```tsx
import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Summary">
    <Row><Cell>Product</Cell><Cell>Count</Cell></Row>
    <Row><Cell>Apples</Cell><Cell value={12} /></Row>
    <Row><Cell formula="=SUM(B2:B2)" /></Row>
  </Worksheet>
</Workbook>;
```

```sh
xlsx-jsx validate report.tsx
xlsx-jsx build report.tsx -o report.xlsx
xlsx-jsx reverse report.xlsx -o report.roundtrip.tsx
xlsx-jsx spec --format json
```

## Feature matrix

| Feature | Build | Reverse external XLSX | Status |
|---|---:|---:|---|
| Worksheets and positional rows/cells | yes | yes | v1 |
| Rectangular tables, totals rows, column totals, calculated columns, and formats | yes | yes | `Table` supports headered and headerless ranges (`headerRow={false}`); headerless `TableColumn` declarations preserve explicit OOXML column names. Columns also support totals/formulas and bold/italic/strikethrough/underline/script/RGB/solid and patterned fills/number-format/four-edge-border differential data formats; script needs another dxf property because of an upstream writer omission; `headerFormat` supports the full Cell-format subset and reverses as header `Cell` formatting; unsupported dxf properties are rejected |
| Worksheet visibility | yes | yes | `visibility="hidden"` or `"veryHidden"`; first worksheet stays visible |
| Active worksheet | yes | yes | `active={true}` selects one visible worksheet when opened; default is the first worksheet |
| First visible tab | yes | yes | `firstTab={true}` selects one visible worksheet as the leftmost tab; default is the first worksheet |
| In-cell checkboxes | yes | yes | `Cell value={true} checkbox={true}` writes Excel 2024 checkbox metadata; boolean values stay usable in formulas and older Excel versions display TRUE/FALSE |
| Worksheet tab color | yes | yes | `tabColor="#4472C4"` uses opaque RGB; theme/indexed colors are rejected |
| Default row height | yes | yes | `defaultRowHeightPixels` uses integer pixels; non-representable external point values are rejected |
| Hide unused worksheet rows | yes | yes | `hideUnusedRows` hides rows without explicit content |
| Screen gridlines | yes | yes | `showGridlines={false}` hides on-screen worksheet gridlines; print settings are separate |
| Worksheet view mode | yes | yes | `view="pageLayout"` or `view="pageBreakPreview"`; `zoom` follows the active view |
| Worksheet active selection and scroll origin | yes | yes | `selection={{activeCell: "D3", range: "D3:F5"}}` and `topLeftCell="H6"`; frozen-pane selection applies to its active pane |
| Manual page breaks | yes | yes | `horizontalPageBreaks={[20, 40]}` and `verticalPageBreaks={[8]}` preserve full-axis manual print breaks |
| Worksheet print margins | yes | yes | `pageMargins={{left: 0.5, top: 1}}` sets margins in inches; omitted sides use Excel writer defaults |
| Print scaling and fit-to-pages | yes | yes | `printScale={80}` or `fitToPages={{width: 1, height: 0}}`; mutually exclusive and fit-to-pages cannot override manual breaks |
| Page orientation and paper size | yes | yes | `orientation="landscape"` and `paperSize={9}` preserve printed page setup; first printed-page number remains unsupported because upstream emits invalid OOXML |
| Printed-page centering and quality | yes | yes | `centerHorizontally`, `centerVertically`, `printBlackAndWhite`, and `printDraft` control print layout/output |
| Printed headings and page order | yes | yes | `printHeadings={true}` prints row/column headings; `pageOrder="overThenDown"` chooses traversal order |
| Print area and repeated titles | yes | yes | `printArea="A1:D20"`, `repeatRows="1:2"`, and `repeatColumns="A:B"` preserve print ranges and repeating headings |
| Worksheet header and footer | yes | yes | `header="&CPage &P of &N"`, `footer="&R&D"`, plus scale/alignment flags; image and even/first-page variants are rejected |
| Worksheet protection and cell lock/formula visibility | yes | yes | `protection={{formatCells: true}}`, `locked={false}`, and `formulaHidden={true}`; password hashes and protected ranges are rejected |
| Print gridlines | yes | yes | `printGridlines={true}` enables gridlines on printed pages |
| Worksheet zoom | yes | yes | `zoom={125}` preserves screen zoom percentage (10–400; default 100) |
| Selected worksheet tabs | yes | yes | `selected={true}` preserves grouped/highlighted worksheet tabs independently of `active={true}` |
| Worksheet reading direction | yes | yes | `rightToLeft={true}` enables right-to-left sheet layout |
| Outline summary direction | yes | yes | `summaryBelow={false}` / `summaryRight={false}` place row/column summary markers above/left; defaults are true |
| Excel tables | yes | yes | Rectangular named tables with built-in styles, filter/banding flags, totals rows, totals functions/labels, calculated-column formulas, supported differential formats, and `altText` / `altTextTitle` accessibility metadata |
| String, number, boolean, blank values | yes | yes | v1 |
| Formula text, cached result and dynamic arrays | yes | yes | Ordinary formulas, scalar caches, legacy array groups and Excel 365 dynamic arrays round-trip; dynamic arrays use `dynamicArray={true}` with optional `arrayFormulaRange`; formulas are not calculated, non-default array caches and shared/data-table formulas are rejected |
| Bold and italic cell font | yes | yes | applied as boolean `Cell` props |
| Cell, row and column fills | yes | yes | `fillColor` is the solid-fill shorthand; `fillPattern` plus optional opaque RGB foreground/background colors preserves OOXML patterns; gradients and theme/indexed colors are rejected |
| Font color | yes | yes | `fontColor` uses opaque RGB `#RRGGBB`; theme/indexed colors are rejected |
| Font name, size, family, charset, theme scheme, underline, strikethrough and script | yes | yes | Literal `fontName`; size 1–409 pt; family 0–5; charset byte; `fontScheme` none/body/headings; four underline styles; boolean strike; superscript/subscript |
| Rich text cell runs | yes | yes | `<TextRun>` preserves run boundaries and supported font properties in shared or inline strings; phonetics and unmodeled run properties are rejected |
| Cell borders | yes | yes | Four independent edge styles, diagonal style/direction, and opaque RGB colors; theme/indexed colors remain rejected |
| Date and local datetime | yes | yes | ISO `date` property; time zones and 1904-epoch workbooks are rejected |
| Row height, hidden state, outline levels and default format | yes | yes | `outlineLevel` supports levels 1–7; `collapsed={true}` marks a summary row; row format properties provide defaults for cells without explicit format |
| Column width, hidden state, outline levels and default format | yes | yes | `Column` ranges use 1-based indices and pixel widths; `collapsed={true}` marks a one-column summary; format properties provide defaults for cells without explicit format; best-fit flags are rejected |
| Alignment, rotation, reading direction, quote prefix, wrapping, indentation and shrink-to-fit | yes | yes | `horizontalAlign`, `verticalAlign`, `textRotation`, `readingDirection`, `quotePrefix`, `wrapText`, `indent`, `shrinkToFit`; unsupported combinations are rejected |
| Auto-filter range, list, blank and custom criteria | yes | yes | `autoFilterColumns` uses zero-based relative columns; unsupported date-group filters are rejected |
| Frozen panes | yes | yes | `freezeRows`, `freezeColumns`, optional `freezeTopLeftCell`, and supported active-pane selection; custom split panes remain rejected |
| Merged cells | yes | yes | `Cell rowSpan={2} columnSpan={3}`; populated or differently styled merged interiors are rejected |
| Data validation | yes | yes | Any-value rules with input prompts, list, whole-number, decimal, date, time, text-length, and custom formula rules; `formula1`/`formula2` and operators round-trip |
| Ignored Excel errors | yes | yes | Worksheet-level `<IgnoreError range="A1:B2" errorType="numberStoredAsText" />` preserves Excel's error-checking exceptions; overlapping annotations are rejected |
| Hyperlinks | yes | yes | External URL/URI and internal `internal:Sheet!A1` links on text cells preserve target, display text, and `hyperlinkTooltip` screen tips |
| Legacy cell notes | yes | yes (supported subset) | `<Note cell="B3" text="Review this value" movement="moveAndSizeAfter" />` preserves author/prefix, integer-pixel dimensions, visibility, solid RGB fill, font name/size/family, alt text, and the writer's hidden-row/column anchor mode; unsupported VML shapes, rich text, or positions that cannot be regenerated from sheet geometry are rejected |
| Sparklines | yes | yes (supported subset) | `<Sparkline>` and `<SparklineGroup>` support line/column/win-loss, seven RGB color roles, display flags, empty-cell handling, line weight, individual/group/manual axis bounds, reversible date ranges, and all 36 built-in style palettes |
| Basic charts | yes | yes (supported subset) | `<Chart type="line" cell="D2" style={7} xAxisCrossing="max" yAxisCrossing="min" title="Quarterly sales" xAxisTitle="Quarter" yAxisTitle="Revenue" yAxisNumberFormat="$#,##0.00" yAxisMin={0} yAxisMax={100000} yAxisMajorUnit={10000} yAxisMinorUnit={2000} yAxisLogBase={10} valueAxisReversed={true} valueAxisDisplayUnit="millions" valueAxisDisplayUnitLabel={false} dropLines={true} dropLinesColor="#4472C4" dropLinesWidth={1.5} dropLinesDash="dashDot" dropLinesTransparency={25} categoryAxisLabelInterval={2} categoryAxisTickInterval={2} categoryAxisLabelAlignment="left" categoryAxisOnTickMarks={true} horizontalAxisHidden={true} horizontalAxisMajorTickType="none" verticalAxisMajorTickType="outside" horizontalAxisLabelPosition="high" verticalAxisLabelPosition="low" verticalAxisMajorGridlines={false} horizontalAxisMinorGridlines={true} legendPosition="bottom" legendLineColor="#123456" legendLineWidth={1.25} legendLineDash="dashDot" legendLineTransparency={25} legendLayout={{x: 0.2, y: 0.3, width: 0.4, height: 0.5}}><Series values="Sales!$B$2:$B$4" categories="Sales!$A$2:$A$4" title="Revenue" showValue={true} showCategoryName={true} labelNumberFormat="$#,##0.00" labelFontSize={12} labelFontBold={true} labelFontColor="#336699" /></Chart>` writes and reverses chart style presets (IDs 1–48, default 2), X/Y axis crossing positions (automatic/min/max), per-series data-label contents, supported positions, unlinked number formats, and the supported series-level data-label font subset, static chart/axis/series titles, X/Y axis number formats, value-axis min/max bounds, major/minor units, logarithmic base, reversed orientation, and built-in display units with unit-label visibility; drop lines and supported line styling round-trip on area, line, and stock charts; category-axis label/tick intervals, label alignment and on-tick position, horizontal/vertical axis hidden state, major/minor tick types, gridline visibility, legend visibility/position/layout and supported shape formatting, standard/stacked/percent-stacked area, column, bar, and line charts, all three radar styles, plain, straight-, and smooth-line scatter with/without markers, 3-series HLC and 4-series OHLC stock charts, and pie/doughnut charts, doughnut hole size and first-slice angle, series formulas, object name, alt text, pixel size, and two-cell anchor; linked label formats, formula-backed custom label text, unsupported point-label structures, custom display-unit labels/layouts, unsupported drop-line effects/fills, and other custom chart/axis styling, non-default geometry, and other drawing shapes are rejected |
| Secondary chart axes | yes | yes (supported subset) | Set `secondaryAxis={true}` on one or more Cartesian `<Series>` while retaining at least one primary series; reverse links series to the secondary pair via OOXML axis IDs. Distinct category ranges, category-axis label/tick intervals, static titles, secondary tick-label fonts, explicit number formats, major/minor gridlines and styling, tick-mark types, label positions, hidden/visible axis states, and secondary value-axis bounds, intervals, log base, reversal, and built-in display units with label visibility round-trip; scatter charts support independent `x2AxisReversed` / `y2AxisReversed` and display units; horizontal bar scaling settings are rejected because the pinned writer puts them on a category axis |
| Chart data-label positions | yes | yes (supported subset) | `Series.labelPosition="insideEnd"` and `Series.pointLabels={{0: {position: "outsideEnd"}}}` round-trip supported `dLblPos` placements; valid options depend on chart type |
| Chart accessibility | yes | yes | `Chart.decorative={true}` marks a chart as decorative for assistive technology; it is mutually exclusive with `altText` |
| Chart anchor movement | yes | yes (supported subset) | `Chart.movement="moveOnly"` or `"fixed"` preserves the chart's behavior when worksheet cells move or resize; the default `"moveAndSize"` is omitted |
| OHLC stock bar styles | yes | yes (supported subset) | Four-series `type="stock"` charts accept separate fill transparency and `ChartLine` outlines for up/down bars |
| Chart data-label extras | yes | yes (supported subset) | `showXValue` / `showYValue` name scatter-plot coordinates (exact aliases of `showCategoryName` / `showValue`); `showLegendKey`, `showLeaderLines`, and `labelSeparator="semicolon"` preserve series-level keys, leader-line behavior, and supported separators; `Series.customLabels` and sparse `Series.pointLabels` preserve literal text, gaps, per-point display settings, and supported formatting; formula-backed labels remain rejected |
| Chart data-label number formats | yes | yes (supported subset) | `Series.labelNumberFormat="0.00%"` and `Series.pointLabels={{0: {numberFormat: "0.0"}}}` preserve unlinked series- and point-level formats; linked formats remain rejected |
| Chart data-label fonts | yes | yes (supported subset) | Series-level `labelFont*` properties and point-level `pointLabels[index].font*` properties round-trip; only opaque RGB colors, schema-valid values, and single/none underline are supported |
| Chart data-label shape formatting | yes | yes (supported subset) | Series-level `labelFill*` / `labelLine*` properties and point-level `pointLabels[index].fill*` / `line*` properties round-trip; patterns, gradients, and effects remain rejected |
| Chart error bars | yes | yes (supported subset) | `<ErrorBars axis="y" type="custom" plusRange="Data!$D$2:$D$4" minusRange="Data!$E$2:$E$4" lineColor="#336699" lineWidth={1.5} lineDash="dash" />` supports scalar and custom range error amounts plus reversible RGB/width/dash/transparency line styles; other line formatting remains rejected |
| Chart data tables | yes | yes (supported subset) | `<ChartDataTable showHorizontalBorders={false} showLegendKeys={true} fontSize={11} fontBold={true} fontColor="#336699" fontRightToLeft={true} fillPattern="pct20" fillForegroundColor="#F0E0D0" fillBackgroundColor="#FFFFFF" lineColor="#4472C4" lineWidth={1.25} lineDash="dash" lineTransparency={20} />` preserves display flags, supported font properties, RGB solid/pattern/no-fill representations, and a `ChartLine` outline subset; gradients, effects, and extensions remain rejected |
| Chart data display | yes | yes (supported subset) | `showHiddenData`, `emptyCellsAs="gap|zero|span"`, and `showNaAsEmptyCell` round-trip hidden source data, empty-cell behavior, and the pinned c16r3 N/A-as-blank extension |
| Automatic chart title | yes | yes | `autoTitleDeleted={true}` suppresses Excel's generated single-series title; it cannot be combined with an explicit `title` |
| Chart title overlay | yes | yes (supported subset) | `titleOverlay={true}` draws an explicit chart title over the plot area; it requires `title` |
| Chart titles | yes | yes (supported subset) | `title` writes static text; `titleFormula="Data!$A$1"` uses a sheet-qualified single-cell reference and reverses external formula titles. Formula titles cannot be combined with static titles, title styling/layout/overlay, or `autoTitleDeleted` |
| Chart axis titles | yes | yes (supported subset) | `xAxisTitleFormula="Data!$A$1"` / `yAxisTitleFormula="Data!$B$1"` link the primary axis titles to worksheet cells and reverse external formula-backed axes; each excludes its static title and matching title formatting/layout |
| Chart title layout | yes | yes (supported subset) | `titleLayout={{x: 0.25, y: 0.15}}` manually positions a static chart title with reversible fractional offsets (0–1, including 0); it requires `title` |
| Chart title styling | yes | yes (supported subset) | `titleFont*` preserves supported static-title run styling; `titleFillColor` with optional `titleFillTransparency` and `titleLineColor` / `titleLineWidth` / `titleLineDash` / `titleLineTransparency` preserve solid fill alpha and line styling; fill transparency requires fill color; line transparency requires line color; all title shape properties require `title` |
| Chart and plot area fills, outlines, and layout | yes | yes (supported subset) | `chartAreaFillColor="#FFFFFF"` with optional `chartAreaFillTransparency={25}`, `chartAreaLineColor="#123456"`, `plotAreaLineDash="dash"`, and `plotAreaLayout={{x: 0.2, y: 0.3, width: 0.7, height: 0.5}}` round-trip solid fill alpha, supported outline lines, and exact inner/edge manual geometry; patterns, gradients, unsupported effects, and other layout modes are rejected |
| Date category axes | yes | yes (supported subset) | `categoryAxisType="date"` supports day/month/year time units and ISO `categoryAxisMinDate` / `categoryAxisMaxDate` bounds; scatter, pie, and doughnut have no category axis |
| Scatter axis bounds | yes | yes | `xAxisMin` / `xAxisMax` and `yAxisMin` / `yAxisMax` preserve the horizontal and vertical value-axis scales independently |
| Scatter axis intervals | yes | yes | `xAxisMajorUnit` / `xAxisMinorUnit` and `yAxisMajorUnit` / `yAxisMinorUnit` preserve major and minor tick spacing on both value axes |
| Scatter axis number formats | yes | yes | `xAxisNumberFormat` and `yAxisNumberFormat` apply independently to the horizontal and vertical value axes |
| Scatter logarithmic axes | yes | yes | `xAxisLogBase` / `yAxisLogBase` independently apply valid base-2–1000 logarithmic scales to scatter X/Y axes |
| Chart axis crossing | yes | yes (supported subset) | `xAxisCrossing` / `yAxisCrossing` and `x2AxisCrossing` / `y2AxisCrossing` preserve automatic/min/max plus numeric `crossesAt` values, accounting for bar-chart axis swaps and secondary chart-group IDs; category indexes are positive integers, value-axis numbers are finite; horizontal-bar `x2AxisCrossing` overrides, numeric date-axis and ambiguous crossings are rejected |
| Chart axis title styling | yes | yes (supported subset) | `xAxisTitleFont*` / `yAxisTitleFont*` preserve supported title run styling; `xAxisTitleLayout={{x, y}}` / `yAxisTitleLayout={{x, y}}` position static axis titles; matching `*FillColor` and `*FillTransparency` preserve RGB solid-fill alpha; `xAxisTitleLine*` / `yAxisTitleLine*` add supported outlines; layouts and formatting require matching titles, fill alpha requires fill color, and unsupported effects are rejected |
| Chart axis label fonts | yes | yes (supported subset) | `xAxisLabelFont*` and `yAxisLabelFont*` preserve tick-label size, emphasis, underline, opaque RGB, Latin typeface metadata, and rotation independently from axis-title styling; X/Y axes are resolved by category/value type or scatter axis position |
| Scatter axis titles | yes | yes | `xAxisTitle` / `yAxisTitle` preserve static axis titles, including when crossings move the axes to the top/right |
| Scatter axis direction | yes | yes | `xAxisReversed={true}` / `yAxisReversed={true}` and `x2AxisReversed={true}` / `y2AxisReversed={true}` reverse primary and secondary horizontal/vertical value scales independently |
| Scatter axis display units | yes | yes | `xAxisDisplayUnit` / `yAxisDisplayUnit` and `x2AxisDisplayUnit` / `y2AxisDisplayUnit` independently scale primary and secondary value axes; corresponding `*Label={false}` hides that unit label |
| Chart legend layout and overlay | yes | yes (supported subset) | `legendLayout={{x: 0.2, y: 0.3, width: 0.4, height: 0.5}}` positions a visible legend with reversible fractional geometry; `legendOverlay={true}` overlays it on the plot area; neither works with `showLegend={false}` |
| Chart legend shape fill | yes | yes (supported subset) | `legendFillColor="#F2F2F2"` with optional `legendFillTransparency={25}` applies an RGB solid fill with whole-percent alpha to the visible legend; patterns, gradients, and borders are rejected; the pinned writer's overlay + shape-format ordering is schema-invalid, so that combination is rejected |
| Chart legend fonts | yes | yes (supported subset) | `legendFontSize`, bold/italic/underline, opaque RGB color, Latin typeface/pitch-family/charset and rotation format legend text; the font properties require a visible legend and reverse only uniform, supported DrawingML text formatting |
| Chart legend entry visibility | yes | yes | `legendDeletedEntries={[1, 3]}` hides zero-based legend entries (pie/doughnut point indices); duplicate indices and hidden legends are rejected |
| Chart series and trendline legend entries | yes | yes | `Series` / `Trendline deleteFromLegend={true}` hides the corresponding entry; chart-level indices reverse to these properties when unambiguous |
| Chart series markers | yes | yes (supported subset) | `Series markerType="circle" markerSize={9} markerFillColor="#4472C4" markerFillTransparency={25} markerLineColor="#336699" markerLineWidth={1.5} markerLineDash="dashDot" markerLineTransparency={25}` preserves supported symbols, size, RGB fill/alpha and `ChartLine` outlines on line, radar, and scatter charts; pattern/gradient fills and line effects remain rejected |
| Chart series line styles | yes | yes (supported subset) | `Series lineColor="#336699" lineWidth={1.5} lineDash="dashDot"` preserves RGB, quarter-point widths, supported dash patterns, transparency, and hidden lines; unsupported line fills/effects are rejected |
| Chart data-point colors | yes | yes (supported subset) | `pointColors={{0: "#FF0000", 3: "#336699"}} pointColorTransparencies={{3: 25}}` assigns zero-based solid RGB fills with whole-percent transparency and preserves sparse indices; custom point borders/fills are rejected |
| Chart trendlines | yes | yes (supported subset) | `<Trendline type="linear" name="Forecast" forward={2} displayEquation={true} labelFontSize={12} labelFillColor="#FFF2CC" labelFillTransparency={20} labelLineColor="#336699" />` preserves supported models, legend names, equation/R² labels, supported font and shape formatting, and trendline line color/width/dash/transparency/visibility; patterns, gradients, effects, other label formatting, and layouts remain rejected |
| Bar/column gap and overlap | yes | yes (supported subset) | `Series gap={70} overlap={-25}` preserves group-level spacing; conflicting per-series declarations are rejected and reverse canonicalizes the value onto the first Series |
| Chart axis gridline styling | yes | yes (supported subset) | Primary and `x2Axis`/`y2Axis` gridline properties accept a visibility boolean or `{color, width, dash, transparency}`; supported RGB line styles round-trip, while fills/effects are rejected |
| Chart series smoothing | yes | yes (supported subset) | `Series smooth={true}` or `{false}` explicitly selects smoothed/straight segments on line and scatter charts; omission follows chart-type defaults |
| Invert negative chart values | yes | yes (supported subset) | `Series invertIfNegative={true}` enables inverse fills on bar/column charts; custom inverse RGB color extensions are rejected |
| Raster images | yes | yes (supported subset) | `<Image cell="A1" width={96} height={48} offsetX={8} data="data:image/png;base64,..." hyperlink="https://example.com" />` embeds PNG/JPEG/GIF/BMP bytes and reverses alt text/decorative metadata, supported URI/internal workbook hyperlinks, pixel offsets, move/resize behavior, and two-cell anchors; unsupported transforms, crops, or custom geometry are rejected |
| Worksheet background images | yes | yes | `<BackgroundImage data="data:image/png;base64,..." />` tiles PNG/JPEG/GIF/BMP bytes from the worksheet origin and reverses external image relationships without a manifest |
| Worksheet header/footer images | yes | yes (unscaled subset) | `<HeaderImage position="left" data="data:image/png;base64,..." />` / `<FooterImage ... />` round-trip external VML image relationships without a manifest; require a matching `&G` marker in `header`/`footer`; independently scaled and even/first-page images are rejected |
| Worksheet VBA code name | yes | yes | `codeName="DataSheet"` preserves `sheetPr@codeName` without requiring a VBA project; names follow Excel's identifier rules and must be unique |
| Workbook VBA code name | yes | yes | Workbook `codeName="ReportBook"` preserves `workbookPr@codeName` without requiring a VBA project |
| Conditional formatting | yes | partial | Color scales, cell comparisons, formula/text, top/bottom, average, date, blank/error/duplicate rules, data bars, classic and x14 icon sets including custom icon assignments; dxf styles include font decoration, number formats, four edge borders, solid and patterned RGB fills; x14 icon thresholds preserve `greaterThan`, reverse, icon visibility, and stop-if-true despite pinned writer omissions |
| Number formats | yes | yes | `numberFormat` preserves custom codes and common built-in formats |
| Blank cells and special text | yes | yes | quoted, multiline, Unicode, and interior blank cells round-trip |
| Time-only and duration display | yes | yes | numeric serial plus `numberFormat`, including elapsed-time formats |
| Excel error cells | no | rejected with repair diagnostics | fix the error or add an explicit error-cell type |
| Workbook document properties | yes | yes | Core: `title`, `subject`, `author`, `keywords`, `description`, `category`, `status`; extended: `manager`, `company`, `hyperlinkBase`; generated metadata is ignored, unsupported authored metadata is rejected |
| Custom workbook document properties | yes | yes | `<CustomProperty name="Department" type="string" value="Finance" />`; supports string, integer, real, boolean and dateTime values |
| Workbook and worksheet-scoped defined names | yes | yes | `<DefinedName name="TaxRate" formula="0.05" />` at workbook scope or inside a worksheet; names and formula text round-trip, while hidden/comment metadata is rejected |
| Phonetics, custom package parts | no | rejected | unsupported semantics are not silently flattened or discarded |
| Shared/data-table formulas, other custom view state, unsupported authored metadata | no | rejected with diagnostics | ordinary/legacy array formulas, screen gridline visibility, and default workbook view remain supported |
| Other font/border metadata, images | no | rejected when present | font family/charset and diagonal borders are supported; start/end/outline border metadata and unrepresented package features are rejected |

See [DSL specification](docs/spec.md) and [writer coverage audit](docs/xlsxwriter-coverage.md).

## Development checks

```sh
cargo fmt -- --check
cargo test --all-targets --all-features --locked
cargo clippy --all-targets --all-features --locked -- -D warnings
./scripts/openxml-test.sh
```

The Open XML integration script validates generated workbooks with
`DocumentFormat.OpenXml` for Microsoft 365.
