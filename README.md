# xlsx-jsx：用 JSX 编写 Excel 工作簿

使用 JSX/TSX 描述 Excel 工作簿，构建 `.xlsx` 文件；也可将支持无损表示的外部工作簿
逆向为 JSX。项目使用 Rust CLI 和内嵌 V8 运行前端，无需安装 Node.js。

[最新版本](https://github.com/lihongjie0209/xlsx-jsx/releases/latest) ·
[DSL 规范](docs/spec.md) · [开发路线图](docs/roadmap.md) ·
[Writer 覆盖审计](docs/xlsxwriter-coverage.md)

## 快速开始

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
xlsx-jsx build report.tsx -o report.xlsx --cache-dir target/xlsx-jsx-cache
xlsx-jsx reverse report.xlsx -o report.roundtrip.tsx
xlsx-jsx spec --format json
```

`build` 和 `validate` 可通过 `--cache-dir` 在多次运行间复用 TS/TSX/JSX 转译缓存。
变更的模块会单独重新转译；工作簿代码每次都会重新执行，`--data` 不会缓存。不传该选项时，
不启用持久化缓存。

错误信息会包含原因、修复建议和对应 DSL 规范。逆向诊断会区分不支持的公式缓存情况：例如，
标量 `formulaResult` 适用于普通公式，不适用于旧式数组公式的缓存值。无法读取的 XLSX ZIP
包会报告为文件格式错误，并建议重新保存或导出为 `.xlsx`，而不会误报成工作簿功能不支持。
缺失或无法读取必需部件时，会报告工作簿包不完整/格式错误，并给出修复建议。

## 功能矩阵

优先级较低、暂缓实现的功能见[开发路线图](docs/roadmap.md)。矩阵中的 JSX 属性名和
OOXML 标识保留英文，以便直接对应代码和文件格式。

| 功能领域 | 构建 | 外部 XLSX 逆向 | 状态 |
|---|---:|---:|---|
| Worksheets and positional rows/cells | 支持 | 支持 | v1 |
| Incremental TS/TSX/JSX transpilation cache | 支持 | 不适用 | `build` and `validate --cache-dir DIR` persist per-module transpilation; source/compiler changes invalidate entries, while workbook evaluation and data inputs always run fresh |
| Rectangular tables, totals rows, column totals, calculated columns, and formats | 支持 | 支持 | `Table` supports headered/headerless ranges, list-value, custom comparison and text-match auto-filter criteria, totals/formulas, accessibility metadata, differential formats, and header formatting; unsupported filter kinds and dxf properties are rejected |
| Worksheet visibility | 支持 | 支持 | `visibility="hidden"` or `"veryHidden"`; first worksheet stays visible |
| Active worksheet | 支持 | 支持 | `active={true}` selects one visible worksheet when opened; default is the first worksheet |
| Workbook theme | 支持 | 支持（指定预设） | `theme="excel2023"` selects the Excel 2023 Office/Aptos theme; unknown custom theme XML is rejected rather than silently replaced |
| First visible tab | 支持 | 支持 | `firstTab={true}` selects one visible worksheet as the leftmost tab; default is the first worksheet |
| Chartsheets | 支持 | 支持（子集） | `<Chartsheet name="Sales"><Chart type="column">…</Chart></Chartsheet>` uses one full-sheet chart with data series referencing ordinary worksheets; worksheet anchors/size, active or grouped Chartsheet tabs, portrait canvas geometry, and extra drawing objects are rejected |
| In-cell checkboxes | 支持 | 支持 | `Cell value={true} checkbox={true}` writes Excel 2024 checkbox metadata; boolean values stay usable in formulas and older Excel versions display TRUE/FALSE |
| Worksheet tab color | 支持 | 支持 | `tabColor="#4472C4"` or `tabColor={{theme: 4, shade: 2}}` preserves opaque RGB or a writer-supported theme shade; indexed/automatic colors and arbitrary tints are rejected |
| Default row/column dimensions | 支持 | 支持 | `defaultRowHeightPixels` uses integer pixels; `defaultColumnWidthPixels` supports writer widths 56, 64, 72, 80, 96, 104, and 120; unsupported external widths are rejected instead of rounded |
| Autofit worksheet columns | 支持 | 规范化为显式列宽 | `Worksheet autoFit={true}` calculates approximate final widths at build time using the pinned writer's Calibri 11 metrics; the intent is not stored in OOXML. Per-column `Column.autoFitWidthPixels` preserves explicit width plus `bestFit="1"` metadata |
| Hide unused worksheet rows | 支持 | 支持 | `hideUnusedRows` hides rows without explicit content |
| Screen gridlines | 支持 | 支持 | `showGridlines={false}` hides on-screen worksheet gridlines; print settings are separate |
| Worksheet row/column headings | 支持 | 支持 | `showRowColHeaders={false}` hides Excel's numbered row and lettered column headings; independent of screen gridlines and printed headings |
| Worksheet view mode | 支持 | 支持 | `view="pageLayout"` or `view="pageBreakPreview"`; `zoom` follows the active view |
| Worksheet active selection and scroll origin | 支持 | 支持 | `selection={{activeCell: "D3", range: "D3:F5"}}` and `topLeftCell="H6"`; frozen-pane selection applies to its active pane |
| Manual page breaks | 支持 | 支持 | `horizontalPageBreaks={[20, 40]}` and `verticalPageBreaks={[8]}` preserve full-axis manual print breaks |
| Worksheet print margins | 支持 | 支持 | `pageMargins={{left: 0.5, top: 1}}` sets margins in inches; omitted sides use Excel writer defaults |
| Print scaling and fit-to-pages | 支持 | 支持 | `printScale={80}` or `fitToPages={{width: 1, height: 0}}`; mutually exclusive and fit-to-pages cannot override manual breaks |
| Page orientation and paper size | 支持 | 支持 | `orientation="landscape"` and `paperSize={9}` preserve printed page setup; first printed-page number remains unsupported because upstream emits invalid OOXML |
| Printed-page centering and quality | 支持 | 支持 | `centerHorizontally`, `centerVertically`, `printBlackAndWhite`, and `printDraft` control print layout/output |
| Printed headings and page order | 支持 | 支持 | `printHeadings={true}` prints row/column headings; `pageOrder="overThenDown"` chooses traversal order |
| Print area and repeated titles | 支持 | 支持 | `printArea="A1:D20"`, `repeatRows="1:2"`, and `repeatColumns="A:B"` preserve print ranges and repeating headings |
| Worksheet header and footer | 支持 | 支持 | `header="&CPage &P of &N"`, `footer="&R&D"`, plus scale/alignment flags; image and even/first-page variants are rejected |
| Worksheet protection and cell lock/formula visibility | 支持 | 支持 | `protection={{formatCells: true}}`, `locked={false}`, and `formulaHidden={true}`; password hashes and protected ranges are rejected |
| Print gridlines | 支持 | 支持 | `printGridlines={true}` enables gridlines on printed pages |
| Worksheet zoom | 支持 | 支持 | `zoom={125}` preserves screen zoom percentage (10–400; default 100) |
| Selected worksheet tabs | 支持 | 支持 | `selected={true}` preserves grouped/highlighted worksheet tabs independently of `active={true}` |
| Worksheet reading direction | 支持 | 支持 | `rightToLeft={true}` enables right-to-left sheet layout |
| Outline summary direction | 支持 | 支持 | `summaryBelow={false}` / `summaryRight={false}` place row/column summary markers above/left; defaults are true |
| Excel tables | 支持 | 支持 | Rectangular named tables with built-in styles, filter/banding flags, totals rows, totals functions/labels, calculated-column formulas, supported differential formats, and `altText` / `altTextTitle` accessibility metadata |
| String, number, boolean, blank values | 支持 | 支持 | v1 |
| Formula text, cached result and dynamic arrays | 支持 | 支持 | Ordinary formulas, scalar caches, legacy array groups and Excel 365 dynamic arrays round-trip; external shared formulas expand to ordinary per-cell formulas; dynamic arrays use `dynamicArray={true}` with optional `arrayFormulaRange`; formulas are not calculated, non-default array caches and data-table formulas are rejected |
| Bold and italic cell font | 支持 | 支持 | applied as boolean `Cell` props |
| Cell, row and column fills | 支持 | 支持 | `fillColor` is the solid-fill shorthand; `fillPattern` plus optional RGB/theme foreground/background colors preserves OOXML patterns; arbitrary tint, indexed/automatic colors and gradients are rejected |
| Font color | 支持 | 支持 | `fontColor` accepts RGB `#RRGGBB` or `{theme: 0..9, shade: 0..5}` and reverses the writer's exact theme tint variants |
| Font name, size, family, charset, theme scheme, underline, strikethrough and script | 支持 | 支持 | Literal `fontName`; size 1–409 pt; family 0–5; charset byte; `fontScheme` none/body/headings; four underline styles; boolean strike; superscript/subscript |
| Rich text cell runs | 支持 | 支持 | `<TextRun>` preserves run boundaries and supported RGB/theme font colors in shared or inline strings; phonetics and unmodeled run properties are rejected |
| Cell borders | 支持 | 支持 | Four independent edge styles, diagonal style/direction, and RGB/theme colors; indexed/automatic colors and non-writer tint values remain rejected |
| Date and local datetime | 支持 | 支持 | ISO `date` property; time zones and 1904-epoch workbooks are rejected |
| Row height, hidden state, outline levels and default format | 支持 | 支持 | `height={18.5}` preserves point precision, `outlineLevel` supports levels 1–7, `collapsed={true}` marks a summary row; row format properties provide defaults for cells without explicit format |
| Column width, hidden state, outline levels and default format | 支持 | 支持 | `Worksheet.defaultColumnWidthPixels={80}` sets a supported sheet-wide default; `Column` ranges use 1-based indices; `width={12.5}` uses Excel character units, `widthPixels={144}` uses pixels, and `autoFitWidthPixels={144}` preserves an explicit pixel width with OOXML `bestFit="1"`; `collapsed={true}` marks a one-column summary; format properties provide defaults for cells without explicit format |
| Alignment, rotation, reading direction, quote prefix, wrapping, indentation and shrink-to-fit | 支持 | 支持 | `horizontalAlign` supports left/center/centerAcross/right/justify/distributed; `verticalAlign` supports top/center/bottom/justify/distributed; `centerAcross` maps to OOXML `centerContinuous` and must be applied to each participating cell; `textRotation`, `readingDirection`, `quotePrefix`, `wrapText`, `indent`, `shrinkToFit`; reverse diagnostics name unsupported OOXML alignment attributes and suggest a repair |
| Auto-filter range, list, blank, comparison and text-match criteria | 支持 | 支持 | `autoFilterColumns` uses zero-based relative columns; custom filters support comparisons and begins/ends/contains (plus negated forms), while date-group filters are rejected |
| Frozen panes | 支持 | 支持 | `freezeRows`, `freezeColumns`, optional `freezeTopLeftCell`, and supported active-pane selection; custom split panes remain rejected |
| Merged cells | 支持 | 支持 | `Cell rowSpan={2} columnSpan={3}`; populated or differently styled merged interiors are rejected |
| Data validation | 支持 | 支持 | Any-value rules with input prompts, list, whole-number, decimal, date, time, text-length, and custom formula rules; `formula1`/`formula2`, operators, and disjoint whitespace-separated ranges round-trip |
| Ignored Excel errors | 支持 | 支持 | Worksheet-level `<IgnoreError range="A1:B2" errorType="numberStoredAsText" />` preserves Excel's error-checking exceptions; overlapping annotations are rejected |
| Hyperlinks | 支持 | 支持 | External URL/URI and internal `internal:Sheet!A1` links on text cells preserve target, display text, and `hyperlinkTooltip` screen tips |
| Legacy cell notes | 支持 | 支持（子集） | `<Note cell="B3" text="Review this value" movement="moveAndSizeAfter" />` preserves author/prefix, integer-pixel dimensions, visibility, solid RGB fill, font name/size/family, alt text, and the writer's hidden-row/column anchor mode; unsupported VML shapes, rich text, or positions that cannot be regenerated from sheet geometry are rejected |
| Sparklines | 支持 | 支持（子集） | `<Sparkline>` and `<SparklineGroup>` support line/column/win-loss, seven RGB color roles, display flags, empty-cell handling, line weight, individual/group/manual axis bounds, reversible date ranges, all 36 built-in style palettes, and RGB overrides over uniquely identifiable theme presets |
| Basic charts | 支持 | 支持（子集） | `<Chart type="line" cell="D2" style={7} xAxisCrossing="max" yAxisCrossing="min" title="Quarterly sales" xAxisTitle="Quarter" yAxisTitle="Revenue" yAxisNumberFormat="$#,##0.00" yAxisMin={0} yAxisMax={100000} yAxisMajorUnit={10000} yAxisMinorUnit={2000} yAxisLogBase={10} valueAxisReversed={true} valueAxisDisplayUnit="millions" valueAxisDisplayUnitLabel={false} dropLines={true} dropLinesColor="#4472C4" dropLinesWidth={1.5} dropLinesDash="dashDot" dropLinesTransparency={25} categoryAxisLabelInterval={2} categoryAxisTickInterval={2} categoryAxisLabelAlignment="left" categoryAxisOnTickMarks={true} horizontalAxisHidden={true} horizontalAxisMajorTickType="none" verticalAxisMajorTickType="outside" horizontalAxisLabelPosition="high" verticalAxisLabelPosition="low" verticalAxisMajorGridlines={false} horizontalAxisMinorGridlines={true} legendPosition="bottom" legendLineColor="#123456" legendLineWidth={1.25} legendLineDash="dashDot" legendLineTransparency={25} legendLayout={{x: 0.2, y: 0.3, width: 0.4, height: 0.5}}><Series values="Sales!$B$2:$B$4" categories="Sales!$A$2:$A$4" title="Revenue" showValue={true} showCategoryName={true} labelNumberFormat="$#,##0.00" labelFontSize={12} labelFontBold={true} labelFontColor="#336699" /></Chart>` writes and reverses chart style presets (IDs 1–48, default 2), X/Y axis crossing positions (automatic/min/max), per-series data-label contents, supported positions, unlinked number formats, and the supported series-level data-label font subset, static chart/axis/series titles, X/Y axis number formats, value-axis min/max bounds, major/minor units, logarithmic base, reversed orientation, and built-in display units with unit-label visibility; drop lines and supported line styling round-trip on area, line, and stock charts; category-axis label/tick intervals, label alignment and on-tick position, horizontal/vertical axis hidden state, major/minor tick types, gridline visibility, legend visibility/position/layout and supported shape formatting, standard/stacked/percent-stacked area, column, bar, and line charts, all three radar styles, plain, straight-, and smooth-line scatter with/without markers, 3-series HLC and 4-series OHLC stock charts, and pie/doughnut charts, doughnut hole size and first-slice angle, series formulas, object name, alt text, pixel size, and two-cell anchor; linked label formats, formula-backed custom label text, unsupported point-label structures, custom display-unit labels/layouts, unsupported drop-line effects/fills, and other custom chart/axis styling, non-default geometry, and other drawing shapes are rejected |
| Secondary chart axes | 支持 | 支持（子集） | Set `secondaryAxis={true}` on one or more Cartesian `<Series>` while retaining at least one primary series; reverse links series to the secondary pair via OOXML axis IDs. Distinct category ranges, category-axis label/tick intervals, static titles, secondary tick-label fonts, explicit number formats, major/minor gridlines and styling, tick-mark types, label positions, hidden/visible axis states, and secondary value-axis bounds, intervals, log base, reversal, and built-in display units with label visibility round-trip; scatter charts support independent `x2AxisReversed` / `y2AxisReversed` and display units; horizontal bar scaling settings are rejected because the pinned writer puts them on a category axis |
| Chart data-label positions | 支持 | 支持（子集） | `Series.labelPosition="insideEnd"` and `Series.pointLabels={{0: {position: "outsideEnd"}}}` round-trip supported `dLblPos` placements; valid options depend on chart type |
| Chart accessibility | 支持 | 支持 | `Chart.decorative={true}` marks a chart as decorative for assistive technology; it is mutually exclusive with `altText` |
| Chart anchor movement | 支持 | 支持（子集） | `Chart.movement="moveOnly"` or `"fixed"` preserves the chart's behavior when worksheet cells move or resize; the default `"moveAndSize"` is omitted |
| OHLC stock bar styles | 支持 | 支持（子集） | Four-series `type="stock"` charts accept separate fill transparency and `ChartLine` outlines for up/down bars |
| Chart data-label extras | 支持 | 支持（子集） | `showXValue` / `showYValue` name scatter-plot coordinates (exact aliases of `showCategoryName` / `showValue`); `showLegendKey`, `showLeaderLines`, and `labelSeparator="semicolon"` preserve series-level keys, leader-line behavior, and supported separators; `Series.customLabels` and sparse `Series.pointLabels` preserve literal text, gaps, per-point display settings, and supported formatting; formula-backed labels remain rejected |
| Chart data-label number formats | 支持 | 支持（子集） | `Series.labelNumberFormat="0.00%"` and `Series.pointLabels={{0: {numberFormat: "0.0"}}}` preserve unlinked series- and point-level formats; linked formats remain rejected |
| Chart data-label fonts | 支持 | 支持（子集） | Series-level `labelFont*` properties and point-level `pointLabels[index].font*` properties round-trip; only opaque RGB colors, schema-valid values, and single/none underline are supported |
| Chart data-label shape formatting | 支持 | 支持（子集） | Series-level `labelFill*` / `labelLine*` properties and point-level `pointLabels[index].fill*` / `line*` properties round-trip; patterns, gradients, and effects remain rejected |
| Chart error bars | 支持 | 支持（子集） | `<ErrorBars axis="y" type="custom" plusRange="Data!$D$2:$D$4" minusRange="Data!$E$2:$E$4" lineColor="#336699" lineWidth={1.5} lineDash="dash" />` supports scalar and custom range error amounts plus reversible RGB/width/dash/transparency line styles; other line formatting remains rejected |
| Chart data tables | 支持 | 支持（子集） | `<ChartDataTable showHorizontalBorders={false} showLegendKeys={true} fontSize={11} fontBold={true} fontColor="#336699" fontRightToLeft={true} fillPattern="pct20" fillForegroundColor="#F0E0D0" fillBackgroundColor="#FFFFFF" lineColor="#4472C4" lineWidth={1.25} lineDash="dash" lineTransparency={20} />` preserves display flags, supported font properties, RGB solid/pattern/no-fill representations, and a `ChartLine` outline subset; gradients, effects, and extensions remain rejected |
| Chart data display | 支持 | 支持（子集） | `showHiddenData`, `emptyCellsAs="gap|zero|span"`, and `showNaAsEmptyCell` round-trip hidden source data, empty-cell behavior, and the pinned c16r3 N/A-as-blank extension |
| Automatic chart title | 支持 | 支持 | `autoTitleDeleted={true}` suppresses Excel's generated single-series title; it cannot be combined with an explicit `title` |
| Chart title overlay | 支持 | 支持（子集） | `titleOverlay={true}` draws an explicit chart title over the plot area; it requires `title` |
| Chart titles | 支持 | 支持（子集） | `title` writes static text; `titleFormula="Data!$A$1"` uses a sheet-qualified single-cell reference and reverses external formula titles. Formula titles cannot be combined with static titles, title styling/layout/overlay, or `autoTitleDeleted` |
| Chart axis titles | 支持 | 支持（子集） | `xAxisTitleFormula="Data!$A$1"` / `yAxisTitleFormula="Data!$B$1"` link the primary axis titles to worksheet cells and reverse external formula-backed axes; each excludes its static title and matching title formatting/layout |
| Chart title layout | 支持 | 支持（子集） | `titleLayout={{x: 0.25, y: 0.15}}` manually positions a static chart title with reversible fractional offsets (0–1, including 0); it requires `title` |
| Chart title styling | 支持 | 支持（子集） | `titleFont*` preserves supported static-title run styling; `titleFillColor` with optional `titleFillTransparency` and `titleLineColor` / `titleLineWidth` / `titleLineDash` / `titleLineTransparency` preserve solid fill alpha and line styling; fill transparency requires fill color; line transparency requires line color; all title shape properties require `title` |
| Chart and plot area fills, outlines, and layout | 支持 | 支持（子集） | `chartAreaFillColor="#FFFFFF"` with optional `chartAreaFillTransparency={25}`, `chartAreaLineColor="#123456"`, `plotAreaLineDash="dash"`, and `plotAreaLayout={{x: 0.2, y: 0.3, width: 0.7, height: 0.5}}` round-trip solid fill alpha, supported outline lines, and exact inner/edge manual geometry; patterns, gradients, unsupported effects, and other layout modes are rejected |
| Date category axes | 支持 | 支持（子集） | `categoryAxisType="date"` supports day/month/year time units and ISO `categoryAxisMinDate` / `categoryAxisMaxDate` bounds; scatter, pie, and doughnut have no category axis |
| Scatter axis bounds | 支持 | 支持 | `xAxisMin` / `xAxisMax` and `yAxisMin` / `yAxisMax` preserve the horizontal and vertical value-axis scales independently |
| Scatter axis intervals | 支持 | 支持 | `xAxisMajorUnit` / `xAxisMinorUnit` and `yAxisMajorUnit` / `yAxisMinorUnit` preserve major and minor tick spacing on both value axes |
| Scatter axis number formats | 支持 | 支持 | `xAxisNumberFormat` and `yAxisNumberFormat` apply independently to the horizontal and vertical value axes |
| Scatter logarithmic axes | 支持 | 支持 | `xAxisLogBase` / `yAxisLogBase` independently apply valid base-2–1000 logarithmic scales to scatter X/Y axes |
| Chart axis crossing | 支持 | 支持（子集） | `xAxisCrossing` / `yAxisCrossing` and `x2AxisCrossing` / `y2AxisCrossing` preserve automatic/min/max plus numeric `crossesAt` values, accounting for bar-chart axis swaps and secondary chart-group IDs; category indexes are positive integers, value-axis numbers are finite; horizontal-bar `x2AxisCrossing` overrides, numeric date-axis and ambiguous crossings are rejected |
| Chart axis title styling | 支持 | 支持（子集） | `xAxisTitleFont*` / `yAxisTitleFont*` preserve supported title run styling; `xAxisTitleLayout={{x, y}}` / `yAxisTitleLayout={{x, y}}` position static axis titles; matching `*FillColor` and `*FillTransparency` preserve RGB solid-fill alpha; `xAxisTitleLine*` / `yAxisTitleLine*` add supported outlines; layouts and formatting require matching titles, fill alpha requires fill color, and unsupported effects are rejected |
| Chart axis label fonts | 支持 | 支持（子集） | `xAxisLabelFont*` and `yAxisLabelFont*` preserve tick-label size, emphasis, underline, opaque RGB, Latin typeface metadata, and rotation independently from axis-title styling; X/Y axes are resolved by category/value type or scatter axis position |
| Scatter axis titles | 支持 | 支持 | `xAxisTitle` / `yAxisTitle` preserve static axis titles, including when crossings move the axes to the top/right |
| Scatter axis direction | 支持 | 支持 | `xAxisReversed={true}` / `yAxisReversed={true}` and `x2AxisReversed={true}` / `y2AxisReversed={true}` reverse primary and secondary horizontal/vertical value scales independently |
| Scatter axis display units | 支持 | 支持 | `xAxisDisplayUnit` / `yAxisDisplayUnit` and `x2AxisDisplayUnit` / `y2AxisDisplayUnit` independently scale primary and secondary value axes; corresponding `*Label={false}` hides that unit label |
| Chart legend layout and overlay | 支持 | 支持（子集） | `legendLayout={{x: 0.2, y: 0.3, width: 0.4, height: 0.5}}` positions a visible legend with reversible fractional geometry; `legendOverlay={true}` overlays it on the plot area; neither works with `showLegend={false}` |
| Chart legend shape fill | 支持 | 支持（子集） | `legendFillColor="#F2F2F2"` with optional `legendFillTransparency={25}` applies an RGB solid fill with whole-percent alpha to the visible legend; patterns, gradients, and borders are rejected; the pinned writer's overlay + shape-format ordering is schema-invalid, so that combination is rejected |
| Chart legend fonts | 支持 | 支持（子集） | `legendFontSize`, bold/italic/underline, opaque RGB color, Latin typeface/pitch-family/charset and rotation format legend text; the font properties require a visible legend and reverse only uniform, supported DrawingML text formatting |
| Chart legend entry visibility | 支持 | 支持 | `legendDeletedEntries={[1, 3]}` hides zero-based legend entries (pie/doughnut point indices); duplicate indices and hidden legends are rejected |
| Chart series and trendline legend entries | 支持 | 支持 | `Series` / `Trendline deleteFromLegend={true}` hides the corresponding entry; chart-level indices reverse to these properties when unambiguous |
| Chart series markers | 支持 | 支持（子集） | `Series markerType="circle" markerSize={9} markerFillPattern={{pattern: "diagBrick", foregroundColor: "#C00000", backgroundColor: "#FFF2CC"}} markerLineColor="#336699" markerLineWidth={1.5} markerLineDash="dashDot"` preserves supported symbols, size, opaque RGB solid/pattern/gradient fills and `ChartLine` outlines on line, radar, and scatter charts; gradient alpha/theme colors and line effects remain rejected |
| Chart series fills and outlines | 支持 | 支持（子集） | `Series fillColor="#4472C4" fillTransparency={20} lineColor="#17365D" lineWidth={1.25}` or `fillPattern={{pattern: "diagBrick", foregroundColor: "#C00000", backgroundColor: "#FFF2CC"}}` preserves series solid/pattern/gradient/no-fill plus supported line outlines; one fill representation per series, while gradient alpha/theme colors and effects are rejected |
| Chart series line styles | 支持 | 支持（子集） | `Series lineColor="#336699" lineWidth={1.5} lineDash="dashDot"` preserves RGB, quarter-point widths, supported dash patterns, transparency, and hidden lines; unsupported line fills/effects are rejected |
| Chart data-point formatting | 支持 | 支持（子集） | `pointColors={{0: "#FF0000"}} pointPatterns={{1: {pattern: "pct20", foregroundColor: "#336699"}}} pointGradients={{2: {type: "linear", angle: 45, stops: [{color: "#336699", position: 0}, {color: "#FFFFFF", position: 100}]}}} pointFillHidden={[3]} pointLines={{3: {color: "#336699", width: 1.25, dash: "dash"}}}` assigns sparse solid/pattern/gradient/no-fill values and `ChartLine` outlines by zero-based index; gradient stops are opaque RGB whole-percent values, while effects and theme/alpha gradient colors are rejected |
| Chart trendlines | 支持 | 支持（子集） | `<Trendline type="linear" name="Forecast" forward={2} displayEquation={true} labelFontSize={12} labelFillColor="#FFF2CC" labelFillTransparency={20} labelLineColor="#336699" />` preserves supported models, legend names, equation/R² labels, supported font and shape formatting, and trendline line color/width/dash/transparency/visibility; patterns, gradients, effects, other label formatting, and layouts remain rejected |
| Bar/column gap and overlap | 支持 | 支持（子集） | `Series gap={70} overlap={-25}` preserves group-level spacing; conflicting per-series declarations are rejected and reverse canonicalizes the value onto the first Series |
| Chart axis gridline styling | 支持 | 支持（子集） | Primary and `x2Axis`/`y2Axis` gridline properties accept a visibility boolean or `{color, width, dash, transparency}`; supported RGB line styles round-trip, while fills/effects are rejected |
| Chart series smoothing | 支持 | 支持（子集） | `Series smooth={true}` or `{false}` explicitly selects smoothed/straight segments on line and scatter charts; omission follows chart-type defaults |
| Invert negative chart values | 支持 | 支持（子集） | `Series invertIfNegative={true}` enables inverse fills, or `invertIfNegativeColor="#FF6666"` selects the custom opaque RGB negative-value fill on bar/column charts; theme colors and effects are rejected |
| Raster images | 支持 | 支持（子集） | `<Image cell="A1" fit="containCenter" data="data:image/png;base64,..." />` scales an image to its target cell (`stretch`, `contain`, `containCenter`), including custom row heights and column widths; explicit pixel geometry can span mixed custom/default cells. Reverse canonicalizes fitted images to width/height/offsets because OOXML stores geometry, not the fit intent. Alt text/decorative metadata, supported URI/internal workbook hyperlinks, movement, and two-cell anchors round-trip; hidden-cell targets, non-pixel sizes, unsupported transforms, or crops are rejected |
| Worksheet background images | 支持 | 支持 | `<BackgroundImage data="data:image/png;base64,..." />` tiles PNG/JPEG/GIF/BMP bytes from the worksheet origin and reverses external image relationships without a manifest |
| Worksheet header/footer images | 支持 | 支持（unscaled subset） | `<HeaderImage position="left" data="data:image/png;base64,..." />` / `<FooterImage ... />` round-trip external VML image relationships without a manifest; require a matching `&G` marker in `header`/`footer`; independently scaled and even/first-page images are rejected |
| Worksheet VBA code name | 支持 | 支持 | `codeName="DataSheet"` preserves `sheetPr@codeName` without requiring a VBA project; names follow Excel's identifier rules and must be unique |
| Workbook VBA code name | 支持 | 支持 | Workbook `codeName="ReportBook"` preserves `workbookPr@codeName` without requiring a VBA project |
| Conditional formatting | 支持 | 部分支持 | Color scales, cell comparisons, formula/text, top/bottom, average, date, blank/error/duplicate rules, data bars, classic and x14 icon sets including custom icon assignments; dxf styles include font decoration, RGB/theme colors, number formats, four edge borders, solid and patterned fills; x14 icon thresholds preserve `greaterThan`, reverse, icon visibility, and stop-if-true despite pinned writer omissions |
| Number formats | 支持 | 支持 | `numberFormat` preserves custom codes and common built-in formats |
| Blank cells and special text | 支持 | 支持 | quoted, multiline, Unicode, and interior blank cells round-trip |
| Time-only and duration display | 支持 | 支持 | numeric serial plus `numberFormat`, including elapsed-time formats |
| Excel error cells | 支持 | 支持 | `<Cell error="#DIV/0!" />` preserves the seven supported OOXML error tokens as literal error cells without formulas |
| Workbook document properties | 支持 | 支持 | Core: `title`, `subject`, `author`, `keywords`, `description`, `category`, `status`; extended: `manager`, `company`, `hyperlinkBase`; generated metadata is ignored, unsupported authored metadata is rejected |
| Custom workbook document properties | 支持 | 支持 | `<CustomProperty name="Department" type="string" value="Finance" />`; supports string, integer, real, boolean and dateTime values |
| Workbook and worksheet-scoped defined names | 支持 | 支持 | `<DefinedName name="TaxRate" formula="0.05" />` at workbook scope or inside a worksheet; names and formula text round-trip, while hidden/comment metadata is rejected |
| Phonetics, custom package parts | 不支持 | 拒绝 | unsupported semantics are not silently flattened or discarded |
| Data-table formulas, other custom view state, unsupported authored metadata | 不支持 | 拒绝并给出诊断 | ordinary/shared-expanded/legacy array formulas, screen gridline visibility, and default workbook view remain supported |
| Other font/border metadata, images | 不支持 | 检测到时拒绝 | font family/charset and diagonal borders are supported; start/end/outline border metadata and unrepresented package features are rejected |

完整约定见 [DSL 规范](docs/spec.md) 和 [writer 覆盖审计](docs/xlsxwriter-coverage.md)。

## 开发校验

```sh
cargo fmt -- --check
cargo test --all-targets --all-features --locked
cargo clippy --all-targets --all-features --locked -- -D warnings
./scripts/openxml-test.sh
```

Open XML 集成脚本使用 `DocumentFormat.OpenXml` 校验面向 Microsoft 365 的生成工作簿。
脚本支持 `CARGO_TARGET_DIR`，可在独立目录构建和校验，不改动默认的 `target/`。
