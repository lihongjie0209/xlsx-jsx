# xlsx-jsx DSL Specification

Version: 1.62

`xlsx-jsx` is a Rust CLI. It evaluates local JSX/TSX modules, normalizes the
result into a JSON-compatible workbook IR, writes `.xlsx` files with
`rust_xlsxwriter`, and reverses the supported subset of external workbooks with
`calamine`. Node.js is not a runtime dependency.

## Component tree

```tsx
<Workbook>
  <Worksheet name="Summary">
    <Column min={2} max={2} widthPixels={160} />
    <ConditionalFormat range="B2:B20" type="colorScale2" minColor="#FFF2CC" maxColor="#70AD47" />
    <Row><Cell fillColor="#DDEEFF">Name</Cell><Cell>Count</Cell></Row>
    <Row><Cell>Apples</Cell><Cell value={12} /></Row>
    <Row><Cell formula="=SUM(B2:B2)" /></Row>
  </Worksheet>
</Workbook>
```

### Image

`Image` embeds raster image bytes in a worksheet and reverses external image
relationships and drawing anchors without relying on a private manifest.

- `cell`, `width`, `height`, and `data` are required. `cell` is an A1 anchor;
  width and height are positive integer pixels from 1 through 1790. `data` is a
  base64 data URI with one of `image/png`, `image/jpeg`, `image/gif`, or
  `image/bmp`; the encoded bytes are preserved exactly.
- `altText?: string` is limited to 255 valid XML characters.
  `decorative?: boolean` marks an image as decorative. Both are preserved in
  the drawing's non-visual picture properties.
- `offsetX?: integer` and `offsetY?: integer` position the image from the
  top-left of `cell`, in pixels. They default to zero and must be less than
  the default cell dimensions (64×20 pixels) so the anchor remains
  unambiguous.
- `movement?: "moveAndSize" | "moveOnly" | "fixed"` controls how the picture
  responds to cell edits. The default is `moveOnly`, matching
  `Image::new_from_buffer()`. These map to the OOXML two-cell anchor's default,
  `editAs="oneCell"`, and `editAs="absolute"` modes respectively.
- Images are direct worksheet children before rows. Reverse supports the
  writer's two-cell anchors on the default 64×20-pixel row/column grid, zero
  start offsets, whole-pixel dimensions, and default rectangular stretch
  geometry. Unsupported crops, hyperlinks, custom transforms/effects, drawing
  shapes, or geometry that cannot be represented exactly fail with guidance.
- Data URIs make reverse output self-contained and recompilable; authored file
  paths and remote fetches are not part of this interface.

### Chart

`Chart` models area, column, bar, line, and pie charts with series source formulas
and a worksheet anchor. Chart components must appear before worksheet rows.

- `type` is one of `area`, `areaStacked`, `areaPercentStacked`, `column`,
  `columnStacked`, `columnPercentStacked`, `bar`, `barStacked`,
  `barPercentStacked`, `line`, `lineStacked`, `linePercentStacked`, `radar`,
  `radarWithMarkers`, `radarFilled`, `scatter`, `scatterWithLines`,
  `scatterWithLinesAndMarkers`, `scatterWithSmoothLines`,
  `scatterWithSmoothLinesAndMarkers`, `stock`, `pie`, or `doughnut`.
  The `*Stacked` and `*PercentStacked` variants map to the matching OOXML chart
  grouping; plain `area`, `column`, `bar`, and `line` use the writer's standard
  grouping (clustered for bar/column). Reverse reconstructs the DSL type from
  the chart element, bar direction, and grouping; unknown groupings fail with
  repair guidance.
- Radar charts use `Chart.type="radar"` for no markers,
  `"radarWithMarkers"` for point markers, and `"radarFilled"` for filled
  areas. Reverse identifies marker behavior from the radar style and each
  series marker; mixed marker behavior is rejected rather than normalized.
- Scatter charts use `categories` as the X-value range and `values` as the
  Y-value range (the cell contents should be numeric). Both must be explicit,
  sheet-qualified one-dimensional cell ranges of equal
  length. `scatter` renders markers with no connecting line;
  `scatterWithLines` connects points without markers; and
  `scatterWithLinesAndMarkers` connects points and displays markers.
  `scatterWithSmoothLines` and `scatterWithSmoothLinesAndMarkers` use smoothed
  connections, with markers absent/present respectively. Reverse infers the
  line visibility, smoothing, and marker mode from scatter style, per-series
  line properties, smoothing, and marker overrides. Mixed per-series settings, axis titles,
  and category-axis scatter charts fail with repair guidance.
- Stock charts support the standard high-low-close (3 series) and
  open-high-low-close (4 series) forms. `Series` order is significant and must
  be High/Low/Close or Open/High/Low/Close respectively; every series must have
  the same explicit category range. The compiler emits the stock chart's
  required high-low lines, and emits up/down bars only for the 4-series form.
  Series line/marker defaults are normalized to the writer's canonical stock
  representation. Reverse detects these OOXML structures without relying on
  the embedded manifest; charts with missing stock bars, inconsistent
  categories, or unsupported per-series formatting fail with repair guidance.
- Pie and doughnut charts require exactly one `Series` and have no X/Y axes.
  `Chart.holeSize?: number` is valid only for `doughnut`, uses integer
  percentages from 10 through 90, and defaults to the writer's 50% hole. The
  upstream setter accepts 0 but would serialize a schema-invalid value, so the
  DSL excludes it. If an external doughnut omits `holeSize`, reverse emits the
  OOXML default of 10 explicitly. `Chart.firstSliceAngle?: number` is valid
  only for pie and doughnut charts, accepts integer degrees from 0 through 360,
  and defaults to 0; reverse preserves non-default angles.
- `Series.markerType?: "automatic" | "square" | "diamond" | "triangle" | "x" | "star" |
  "shortDash" | "longDash" | "circle" | "plusSign" | "none"` selects an
  OOXML marker symbol; `Series.markerSize?: number` accepts integer sizes from
  2 through 72. Omission preserves each chart type's writer default. Marker
  shape and size overrides are only valid for line, radar, and scatter charts;
  marker shape formatting is currently rejected during reverse rather than
  discarded.
- `Series.lineColor?: "#RRGGBB"`, `lineWidth?: number`, `lineDash?: "solid" |
  "roundDot" | "squareDot" | "dash" | "dashDot" | "longDash" |
  "longDashDot" | "longDashDotDot"`, `lineTransparency?: number`, and
  `lineHidden?: boolean` format series lines on area, bar/column, line, radar,
  scatter, and stock charts. Colors are opaque RGB. Width is in points, from
  0.25 through 1584, in exact 0.25-point increments. Transparency is an
  integer percentage from 0 through 100 and requires `lineColor`. A hidden
  line may include `lineWidth`, but cannot include color, non-solid dash, or
  transparency because the pinned writer emits no-fill and drops those values.
  Omission preserves the chart-type default. Reverse maps only exact writer
  representations; unsupported fills, theme colors, effects, and line XML
  structures fail with repair guidance.
- `ChartDataTable` is an optional child of `Chart`, at most one per chart.
  `showHorizontalBorders`, `showVerticalBorders`, `showOutlineBorders`, and
  `showLegendKeys` are booleans. Omission uses the pinned writer defaults
  (horizontal/vertical/outline on, legend keys off). Reverse preserves the
  four OOXML display flags, treating an omitted flag element as false. Custom
  data-table shape formatting, fonts, and extensions are rejected with repair
  guidance rather than dropped.
- `cell` is an A1 cell anchor. `width` and `height` are positive integer pixels
  (1–1790); omission uses the writer defaults of 480×288 pixels. `Series` has
  required `values` and optional `categories` and `name` sheet-qualified A1
  references. `values` and `categories` must be ranges; `name` may be a cell or
  range. Formula text omits a leading `=` and optional `$` markers are retained.
  Alternatively, `Series.title?: string` sets a static series title (1–255
  valid XML characters); it is mutually exclusive with `name`. Because the
  pinned writer interprets sheet-qualified A1-looking strings as formulas,
  `Series.title` rejects strings matching that syntax. Reverse preserves literal
  titles separately from cell-reference names.
- `Chart.title?: string` is a static chart title (1–255 valid XML characters).
  External titles are reversible only when they are one unformatted rich-text
  run; formulas, multiple runs, and custom title formatting fail with guidance.
- `Chart.xAxisTitle?: string` and `Chart.yAxisTitle?: string` set static X- and
  Y-axis titles (1–255 valid XML characters) for non-pie charts. Reverse maps
  the OOXML category axis to X and value axis to Y; only a single unformatted
  rich-text run is supported. Formula titles and custom formatting are rejected.
- `Chart.xAxisNumberFormat?: string` and `Chart.yAxisNumberFormat?: string`
  set Excel number-format codes on the corresponding category/date and value
  axes. Each is a non-empty, XML-safe format code of at most 255 characters;
  omission keeps the writer's source-linked default. The properties are
  mutually exclusive with pie/doughnut charts and scatter charts (whose two
  axes are both value axes and cannot be unambiguously mapped to X/Y by this
  writer subset). Reverse preserves explicit `sourceLinked=false` codes and
  rejects non-default source-linked codes rather than dropping them.
- `Chart.yAxisMin?: number` and `Chart.yAxisMax?: number` set the numeric bounds
  of the chart's value axis and default to automatic scaling when omitted. Each
  value must be finite; if both are supplied, the minimum must be less than the
  maximum. These properties are not valid for scatter, pie, or doughnut charts.
  Reverse preserves `<valAx><scaling><min/max>` values and rejects bounds on a
  category/date axis because this DSL currently exposes only value-axis bounds.
  Horizontal bar charts use the same value-axis meaning even though the writer
  swaps its internal axis objects.
- `Chart.yAxisMajorUnit?: number` and `Chart.yAxisMinorUnit?: number` set the
  major and minor tick intervals on the numeric value axis. Each provided
  interval must be finite and greater than zero; omission leaves Excel's
  automatic interval. These properties are not valid for scatter, pie, or
  doughnut charts. Reverse preserves `<valAx><majorUnit>/<minorUnit>` values and
  rejects intervals on category/date axes instead of dropping them.
- `Chart.yAxisLogBase?: integer` enables a logarithmic value axis with a base
  from 2 through 1000. Omission leaves the value axis linear. This property is
  available only when the chart has one unambiguous value axis; scatter, pie,
  and doughnut charts reject it. On a logarithmic axis any explicit `yAxisMin`
  or `yAxisMax` must be greater than zero, and series data must contain positive
  values for Excel to plot it. Reverse preserves `<valAx><scaling><logBase>` and
  rejects category/date-axis or otherwise ambiguous log scales.
- `Chart.valueAxisReversed?: boolean` reverses the primary value-axis scale
  (`<c:scaling><c:orientation val="maxMin"/>`). Omission or `false` keeps the
  normal minimum-to-maximum direction. It applies to area, column, bar, line,
  radar, and stock charts; horizontal bar charts still refer to their numeric
  value axis. Scatter charts (two value axes) and pie/doughnut charts (no axes)
  reject it. Reverse preserves explicit reversed value-axis orientation and
  rejects reversed category/date axes rather than silently normalizing them.
- `Chart.categoryAxisLabelInterval?: integer` and
  `Chart.categoryAxisTickInterval?: integer` set the interval between visible
  category/date-axis labels and tick marks, respectively. Each accepts 1–65535;
  omission or `1` uses the writer/Excel default. These settings are unavailable
  on scatter, pie, and doughnut charts because their axes are value-only or
  absent. Reverse preserves `tickLblSkip` and `tickMarkSkip` only on a primary
  category/date axis, omits the semantically default value `1`, and rejects
  invalid values or skips on value axes.
- `Chart.horizontalAxisMajorTickType?: "cross" | "inside" | "outside" | "none"`,
  `Chart.horizontalAxisMinorTickType?: "cross" | "inside" | "outside" | "none"`,
  `Chart.verticalAxisMajorTickType?: "cross" | "inside" | "outside" | "none"`,
  and `Chart.verticalAxisMinorTickType?: "cross" | "inside" | "outside" | "none"`
  configure the tick marks on the primary horizontal/vertical axes. Omission
  retains the writer default. The names follow OOXML axis position, so they
  retain their meaning for bars, reversed axes, and scatter charts. Pie/doughnut charts
  reject these properties because they have no axes. Reverse preserves the
  four OOXML `majorTickMark` / `minorTickMark` values and rejects secondary-axis
  ticks or unsupported tick semantics.
- `Chart.horizontalAxisLabelPosition?: "nextTo" | "high" | "low" | "none"`
  and `Chart.verticalAxisLabelPosition?: "nextTo" | "high" | "low" | "none"`
  position labels on the primary axes. The names follow physical OOXML axis
  position (`b`/`t` versus `l`/`r`) for bar, reversed, and scatter charts.
  Omission and `nextTo` use the writer default; reverse folds explicit
  `nextTo` to omission and preserves non-default values. Pie/doughnut reject
  these properties. Secondary-axis and unsupported label-position semantics
  are rejected rather than discarded.
- `Chart.categoryAxisLabelAlignment?: "center" | "left" | "right"` aligns
  category-axis labels relative to tick marks. The writer default `center` is
  omitted on reverse; `left` and `right` round-trip as OOXML `lblAlgn` values
  `l` and `r`. This property applies to category-axis charts, including bars;
  scatter and pie/doughnut charts reject it because they have no category axis.
  Reverse rejects label alignment on unsupported axis structures.
- `Chart.categoryAxisOnTickMarks?: boolean` sets a category axis to cross at
  category tick marks when `true`; omission or `false` keeps the writer's
  default “between tick marks” position. It applies only to column, bar, and
  line charts. Reverse reads `crossBetween=midCat` from the primary value axis
  and maps it back to the logical category axis (including the swapped axes in
  bar charts); writer-default `crossBetween=between` is omitted. Other values
  or ambiguous/secondary axis structures are rejected rather than discarded.
- `Chart.style?: integer` selects the built-in Excel chart style written as
  `<c:style val="…">`; valid values are 1–48. Omission uses
  `rust_xlsxwriter` style 2, and reverse omits style 2 while preserving other
  style IDs. This is a preset selector, not a guarantee of identical rendering
  across spreadsheet applications; custom series/area/plot styling remains a
  separate unsupported surface.
- `Chart.xAxisCrossing?: "automatic" | "min" | "max"` and
  `Chart.yAxisCrossing?: "automatic" | "min" | "max"` set where the opposite
  axis crosses each `ChartAxis` handle. They map to OOXML `crosses` values
  `autoZero`, `min`, and `max`; omission and `automatic` use the default and
  are folded on reverse. The names follow rust_xlsxwriter's x/y axis handles,
  including its bar-chart axis swap. Numeric `crossesAt` values and crossing
  metadata on secondary, date, or ambiguous axes are rejected with repair guidance.
- `Series.showValue?: boolean`, `Series.showCategoryName?: boolean`,
  `Series.showSeriesName?: boolean`, and
  `Series.showPercentage?: boolean` configure per-series data-label contents
  through `ChartDataLabel`. `showPercentage` is valid only for pie/doughnut
  charts. `Series.labelPosition?: string` configures the label position using
  `center`, `right`, `left`, `above`, `below`, `insideBase`, `insideEnd`,
  `outsideEnd`, or `bestFit`. Line/scatter charts accept the first five;
  column/bar charts accept `center`, `insideBase`, `insideEnd`, and `outsideEnd`;
  pie/doughnut accept `center`, `insideEnd`, `outsideEnd`, and `bestFit`;
  area/radar accept only `center`. Stock charts do not support this property.
  The writer's chart-specific default positions are omitted from reverse output.
  `Series.showLegendKey?: boolean` adds the series key to each label, and
  `Series.showLeaderLines?: boolean` enables leader lines (rendered by Excel
  when labels need them). `Series.labelSeparator?: "comma" | "semicolon" |
  "period" | "space" | "newline"` selects the separator between displayed
  fields; omission uses the writer's comma default. Reverse strips the one
  trailing space the pinned writer serializes after the separator character.
  Explicit false flags and the default comma separator fold to omission.
  `Series.labelNumberFormat?: string` sets the label number format using
  `ChartDataLabel::set_num_format`; it must contain 1–255 valid XML characters.
  Reverse preserves a series `<c:numFmt>` only when `sourceLinked="0"` (the
  writer's explicit, unlinked format) is present; linked or implicit formats
  are rejected because they cannot be reproduced exactly. When no category or
  percentage is displayed, the pinned writer serializes value labels even if
  legend keys, leader lines, position, a separator, or a number format are also
  requested. Therefore `showValue={false}` cannot be combined with those
  options unless `showCategoryName` or `showPercentage` is true. Reverse rejects
  external combinations the writer would normalize by adding values. Default
  false flags and the comma separator are omitted. Custom label text, fonts,
  and point-specific labels are not represented and are rejected with repair
  guidance.
- `ErrorBars` is an optional child of `Series`, with at most one each for the
  `x` and `y` axes. `axis` is required and selects horizontal (`x`) or vertical
  (`y`) error bars; X error bars are valid only for scatter and bar charts,
  while Y error bars are unavailable on bar and pie/doughnut charts because
  the pinned writer either omits the requested axis or has no matching axis.
  For column and bar charts, the writer omits `errDir`; reverse infers its
  unambiguous implicit Y or X axis from the chart type. `type` is `standardError`,
  `fixedValue`, `percentage`, or `standardDeviation`. The latter three require
  a finite `value`; fixed values must be greater than zero, while percentage
  and standard-deviation multipliers may be zero. `standardError` forbids
  `value`. `direction?: "both" | "minus" | "plus"` defaults to `both`, and
  `endCap?: boolean` defaults to true. Reverse preserves the writer's
  `errDir`, `errBarType`, `errValType`, optional `val`, and `noEndCap` values;
  for writer-defaulted bar/column directions it restores the implicit axis.
  Custom plus/minus ranges and custom line formatting are not yet represented;
  external instances using either are rejected instead of being flattened.
- `Chart.horizontalAxisHidden?: boolean` and `Chart.verticalAxisHidden?: boolean`
  hide the primary axis by physical orientation. For bar charts the horizontal
  axis is the value axis and the vertical axis is the category axis; reversed
  axes retain their orientation-based names. Omission/`false` is the default;
  reverse preserves `delete=true` as `true`. Pie/doughnut reject these fields,
  and reverse rejects hidden declarations that cannot be mapped to one primary
  horizontal or vertical axis.
- `Chart.horizontalAxisMajorGridlines?: boolean`,
  `Chart.horizontalAxisMinorGridlines?: boolean`,
  `Chart.verticalAxisMajorGridlines?: boolean`, and
  `Chart.verticalAxisMinorGridlines?: boolean` control primary axis gridline
  visibility. Horizontal/vertical refer to OOXML axis position (`b`/`l`), so they
  retain their meaning for bars and scatter charts. Major-gridline defaults are
  chart-specific (horizontal for bar/radar, vertical for area/column/line/radar/
  scatter/stock); minor gridlines default off. Reverse preserves visibility and
  rejects custom gridline line formatting or secondary-axis gridlines because
  those are not represented by these properties. Pie/doughnut charts have no
  axes and reject all four properties.
- `Chart.showLegend?: boolean` defaults to `true`; `Chart.legendPosition?:
  "right" | "left" | "top" | "bottom" | "topRight"` defaults to `"right"`.
  `legendPosition` cannot be combined with `showLegend={false}` because the
  writer omits the hidden legend and its position.
- `Chart.name?: string` preserves the workbook chart object name and
  `Chart.altText?: string` preserves its accessibility description. Both are
  limited to 255 valid XML characters. Reverse emits an explicit name because
  the writer otherwise assigns one automatically.
- Pie charts require exactly one `Series`; other chart types support one or more
  series. Series order, chart type, anchor, pixel size, and formulas survive
  reverse/recompile.
- Reverse supports two-cell chart anchors with zero start offsets, whole-pixel
  sizes, standard 64×20-pixel worksheet geometry across the chart footprint,
  writer-default chart/axis presentation apart from supported static chart and
  axis titles, X/Y number formats, and value-axis bounds. Other custom chart
  formatting, axis settings, drawing shapes, custom dimensions in the anchor
  footprint, and other chart kinds fail with repair guidance.

### Workbook

- Children: one or more `Worksheet` components and optional global `DefinedName`
  declarations. Names may appear before or after worksheets; source ordering is
  retained when reversing within each scope.
- v1 has no implicit sheet; sheets are emitted in source order.
- Optional `title`, `subject`, `author`, `keywords`, and `description` string
  properties map to OOXML core document properties. Empty strings are omitted.
  Reverse preserves these values; `creator` and `lastModifiedBy` must agree and
  are represented as `author`. Writer-generated creation/modification timestamps
  are intentionally ignored. Other authored core metadata (category, status,
  identifiers, custom properties, and independent last-modifier values) is
  rejected because it has no lossless Workbook property in v1.

### DefinedName

- `name: string` and `formula: string` are required. A `DefinedName` directly
  under `Workbook` is workbook-scoped; directly under a `Worksheet` it is local
  to that sheet. Names use Excel's defined-name syntax, are unique ignoring
  case within a scope, and may not use reserved `_xlnm.*` names. Formulas are
  stored as Excel formula/reference text without a leading `=`; the compiler
  also accepts and removes one leading `=` for authoring convenience.
- Children are not allowed. Reverse preserves global/local scope, name spelling,
  and formula text. `hidden`, comments, and other defined-name metadata are not
  modeled and cause reverse to fail with a repair suggestion rather than being
  dropped. Built-in print-area, print-title, and autofilter names remain
  represented by their worksheet properties, not `DefinedName` components.

### Worksheet

- Required `name: string`; unique ignoring Unicode lowercase comparison.
- `active?: boolean` selects this worksheet when the workbook opens. At most one
  worksheet may set it to `true`; the active worksheet must be visible. When
  omitted, the writer's default active worksheet is used. Reverse emits the
  property only when the active worksheet is not the first worksheet.
- `selected?: boolean` controls whether this worksheet tab is highlighted and
  participates in grouped-sheet operations; it is independent of `active`.
  Multiple worksheets may be selected, but selected sheets must be visible.
  Omission defaults to false except that explicitly setting `active={true}` in
  the compiler also selects that worksheet in the writer. Therefore
  `active={true}` and `selected={false}` are invalid together; when no
  worksheet explicitly sets `active={true}`, the first worksheet is the
  writer's active sheet and likewise cannot set `selected={false}`. Reverse preserves
  `sheetView@tabSelected` independently from `activeTab`, omits values already
  implied by the writer's active-sheet behavior, and rejects an external
  active worksheet that is explicitly unselected because the pinned writer
  always selects the active worksheet during serialization.
- `firstTab?: boolean` selects this worksheet as the leftmost visible tab. At
  most one worksheet may set it to `true`, and it must be visible. When omitted,
  the first worksheet is the leftmost visible tab. Reverse preserves a
  non-first selection using `workbookView@firstSheet` and omits the default.
- `tabColor?: string` sets the worksheet tab's opaque RGB color using exactly
  `#RRGGBB`; reverse preserves explicit RGB values in uppercase. Theme, indexed,
  automatic, and alpha-bearing colors are rejected rather than approximated.
- Children: `Column`, `DataValidation`, `ConditionalFormat`, `IgnoreError`,
  `Table`, `Note`, and `Chart`
  components followed by zero or more `Row` components. Metadata must appear
  before rows;
  reverse canonicalizes it before row data.
- Name is 1–31 UTF-16 code units, excludes `[]:*?/\\`, and cannot begin or
  end with an apostrophe.
- `visibility?: "visible" | "hidden" | "veryHidden"` selects worksheet
  visibility; `visible` is the default and is omitted by reverse. The first
  worksheet must remain visible so the workbook retains a deterministic active
  sheet. Hidden and very-hidden worksheets after the first preserve their
  values and supported worksheet features.
- `defaultRowHeightPixels?: integer` sets the default height for all rows in
  pixels. It must be a positive integer; the XLSX default is 20 pixels and is
  omitted by reverse. Reverse accepts external `defaultRowHeight` values only
  when they map exactly to an integer pixel height (OOXML points ÷ 0.75) within
  Excel's 409-point limit, and rejects contradictory or fractional-pixel
  layouts rather than rounding.
- `hideUnusedRows?: boolean` hides rows without explicit row data using
  `sheetFormatPr@zeroHeight`; it defaults to false and is omitted by reverse.
  It does not replace per-row `hidden` metadata on rows that contain data.
- `showGridlines?: boolean` controls worksheet screen gridlines through
  `sheetView@showGridLines`. It defaults to true and is omitted by reverse;
  print-gridline settings remain separate.
- `view?: "normal" | "pageLayout" | "pageBreakPreview"` selects the worksheet
  view and defaults to `normal`, which is omitted by reverse. `zoom?` is the
  zoom for the selected view: `zoomScale` stores the active zoom; normal view
  also uses matching `zoomScaleNormal`, page layout mirrors it to
  `zoomScalePageLayoutView`, and page-break preview mirrors it to
  `zoomScaleSheetLayoutView`. Reverse rejects conflicting zoom values and
  non-default zoom metadata for inactive views because this DSL models one
  active view at a time.
- `printGridlines?: boolean` enables gridlines when printing through
  `printOptions@gridLines`. It defaults to false and is omitted by reverse;
  other print options not listed here remain unsupported. Writer-default `pageSetup`
  attributes emitted with this setting are accepted; supported print scaling
  settings are described separately below; remaining unsupported print and
  page-setup settings are rejected.
- `horizontalPageBreaks?: integer[]` and `verticalPageBreaks?: integer[]` set
  manual breaks after zero-based worksheet row/column indices. Each list must
  contain unique indices greater than zero and below the worksheet row/column
  limit; at most 1,023 breaks are allowed per axis. The writer sorts indices
  before emitting them. Reverse preserves only manual full-row/full-column
  breaks with matching OOXML counts and bounds; automatic breaks, partial
  breaks, and other `brk` metadata are rejected with repair guidance.
- `pageMargins?: object` overrides any subset of `left`, `right`, `top`,
  `bottom`, `header`, and `footer` margins, in inches. Each value must be a
  finite nonnegative number. Unspecified fields use writer defaults: left/right
  0.7, top/bottom 0.75, header/footer 0.3. Reverse omits default-valued fields
  and preserves all non-default values without unit conversion.
- `printScale?: integer` sets printed-page scale from 10 through 400 (default
  100, omitted by reverse). `fitToPages?: {width: integer, height: integer}`
  instead fits output to the requested page count on each axis; zero means
  unlimited on that axis and one is the writer default. Both fields are
  mutually exclusive. `fitToPages` also cannot be combined with manual page
  breaks because the writer makes fit-to-page override those breaks. Reverse
  validates `sheetPr/pageSetUpPr@fitToPage` against `pageSetup@scale`,
  `fitToWidth`, and `fitToHeight` and rejects contradictory or orphaned values.
- `orientation?: "portrait" | "landscape"` selects printed page orientation
  (default `portrait`, omitted by reverse). `paperSize?: integer` selects the
  OOXML paper-size code accepted by `rust_xlsxwriter` as an unsigned byte
  (0–255; 0 is the writer default). Reverse preserves explicit codes, including
  codes not named by the DSL, and rejects non-numeric or out-of-range values.
  Page order is modeled separately. `Worksheet::set_print_first_page_number()`
  is intentionally not exposed: the pinned writer serializes the numeric page
  number into `pageSetup@useFirstPageNumber` (a boolean attribute) instead of
  `firstPageNumber`, producing schema-invalid OOXML.
- `centerHorizontally?: boolean` and `centerVertically?: boolean` center the
  worksheet data between page margins when printed. `printBlackAndWhite?:
  boolean` forces monochrome output and `printDraft?: boolean` requests draft
  print quality. All default to `false` and reverse omits false/default values.
  They map to `printOptions@horizontalCentered` / `verticalCentered` and
  `pageSetup@blackAndWhite` / `draft`; only true/false OOXML values are
  accepted. These settings combine independently with margins and scaling.
- `printHeadings?: boolean` prints row and column headings, defaulting to
  `false` and omitted by reverse. `pageOrder?: "downThenOver" |
  "overThenDown"` selects print traversal order, defaulting to
  `downThenOver`; the non-default maps to `pageSetup@pageOrder="overThenDown"`.
  The corresponding `printOptions@headings` and page-order values are reversed
  without dropping other supported print settings.
- `zoom?: integer` sets the worksheet screen zoom percentage through
  `sheetView@zoomScale`. It must be an integer from 10 through 400; the default
  is 100 and is omitted by reverse. Print scaling is independent and uses
  `printScale` or `fitToPages`. The writer mirrors this value to the OOXML zoom
  property for the selected view; reverse requires any emitted mirror to match
  `zoomScale`.
- `topLeftCell?: string` sets the first visible worksheet cell, independent of
  frozen-pane `freezeTopLeftCell`; it defaults to `A1` and reverse omits `A1`.
  `selection?: { activeCell: string, range: string }` chooses one contiguous
  selected A1 cell/range and the active cell at one of its corners, matching
  `rust_xlsxwriter::Worksheet::set_selection()`; an interior active cell cannot
  be emitted by the pinned writer and is rejected. Both are worksheet view
  metadata, independent of stored values. When frozen panes are enabled, the
  selection applies to the writer's active pane and other panes retain their
  canonical writer defaults. Reverse rejects multi-range selections, unknown
  pane layouts, and active cells outside the selection.
- `rightToLeft?: boolean` controls worksheet reading direction via
  `sheetView@rightToLeft`. It defaults to `false` and is omitted by reverse;
  `true` is preserved. Cell-level reading-direction formatting remains a
  separate style property.
- `summaryBelow?: boolean` controls whether row outline summary rows appear
  below their detail rows; `true` is the default and is omitted by reverse.
  `false` maps to `sheetPr/outlinePr@summaryBelow="0"` and reverses back to
  `false`. `summaryRight?: boolean` controls whether summary columns appear to
  the right of detail columns; `true` is the default and is omitted by reverse.
  Other outline settings remain unsupported.
- `autoFilter?: string` adds a filter dropdown to one A1 cell range such as
  `A1:D20`. It must be a valid, ordered, absolute worksheet range (optional `$`
  markers are accepted).
- `protection?: object` enables worksheet protection without a password. Its
  optional boolean permissions are `selectLockedCells`, `selectUnlockedCells`,
  `formatCells`, `formatColumns`, `formatRows`, `insertColumns`, `insertRows`,
  `insertLinks`, `deleteColumns`, `deleteRows`, `sort`, `useAutoFilter`,
  `usePivotTables`, `editScenarios`, and `editObjects`. Omitted values use
  `ProtectionOptions` defaults; the object's presence enables protection even
  when empty. Password hashes, chartsheet-only `contents`, and protected cell
  ranges are rejected by reverse because their semantics cannot be recovered
  or represented by this worksheet DSL.
- `Table` describes one Excel table over an existing rectangular cell range.
  The range includes a string header row and at least one data row. `name` is
  an Excel table name; `style?` accepts `TableStyleNone`, `TableStyleLight1`
  through `TableStyleLight21`, `TableStyleMedium1` through
  `TableStyleMedium28`, or `TableStyleDark1` through `TableStyleDark11`;
  the default is `TableStyleMedium9`. `autoFilter?`, `bandedRows?`,
  `bandedColumns?`, `firstColumn?`, and `lastColumn?` control the matching
  table flags and default to true, true, false, false, and false. Table ranges
  must not overlap each other, worksheet auto-filters, or merged cells.
  Header cells must be unique, non-empty strings; explicit header-cell formats
  are preserved as ordinary `Cell` format properties independently of table style.
  `totalsRow?` enables the final row as a totals row (default false). Optional
  `TableColumn` children describe each column in order; if present, their count
  must equal the range width. A column may set `totalLabel?` or
  `totalFunction?` (`none`, `average`, `count`, `countNumbers`, `max`, `min`,
  `sum`, `stdDev`, or `var`) for a totals row, `totalFormula?` for a custom
  totals formula, and/or `formula?` for a calculated column. `totalLabel`,
  `totalFunction`, and `totalFormula` are mutually exclusive. A totals-row range
  includes its final totals row;
  worksheet cell rows must still include all table rows. Reverse preserves
  these attributes and rejects unsupported column metadata or formula shapes.
  `format?` accepts `bold={true}`, `italic={true}`, `strikethrough={true}`,
  `underline` (`single`, `double`, `singleAccounting`, or `doubleAccounting`),
  `script` (`superscript` or `subscript`), and `fontColor`,
  `fillColor`, or differential/ordinary pattern-fill properties `fillPattern`,
  `fillForegroundColor`, and `fillBackgroundColor`, `numberFormat`, the four
  edge borders and their RGB colors, and `borderDiagonalType`; `format` applies a differential style to data
  cells and is preserved as table-column metadata. `headerFormat` styles the
  worksheet header cell and reverses canonically as ordinary `Cell` formatting.
  `headerFormat?` additionally accepts `fontName`, `fontSize` (points, 1–409),
  `fontFamily` (OOXML values 0–5), and `fontCharset` (byte, 0–255); it is
  serialized as ordinary header-cell formatting and reverses canonically as
  `Cell` formatting. The pinned writer does not serialize those four font
  metadata fields in data-cell differential formats, so `format` rejects them
  with a suggestion to move them to `headerFormat` or the header `Cell`.
  Other table-specific format properties (including theme `fontScheme`) are
  rejected until the differential-style reader can preserve them losslessly.
  Differential styles cannot encode
  false overrides, so emphasis properties are additive and must be true.
  `fillColor` is the solid-fill shorthand and is mutually exclusive with pattern
  properties. Pattern names and opaque RGB colors follow the Cell/Row/Column
  fill contract; foreground/background colors require an explicit pattern.
  Reverse preserves supported pattern types and colors from table differential
  formats and ordinary header-cell styles.
  `numberFormat` is an Excel format code of 1–255 XML-valid characters and is
  preserved as the table column's differential number format.
  For data `format`, `script` is written as the differential font's `vertAlign`
  property and is preserved when reversing external tables. Because the pinned
  writer omits a script-only differential format from `<dxfs>`, data `format`
  `script` must be combined with another supported property such as `bold` or
  `fontColor`. `headerFormat` uses ordinary cell style XML and can use `script`
  by itself.
  Border colors require the corresponding edge style. Diagonal direction
  (`up`, `down`, or `upDown`) is preserved; diagonal line style/color is not
  serialized by the pinned writer in differential formats and remains rejected.
  `altText?` and `altTextTitle?` preserve the accessibility summary and title
  written by `Table::set_alt_text()` and `Table::set_alt_text_title()`. Values
  must be strings; empty strings are equivalent to omission. Reverse accepts
  only the pinned writer's x14 table accessibility extension and rejects
  unrelated extension metadata. Headerless tables and other unmodeled table
  metadata remain unsupported.
- `Note` adds one legacy cell note and must appear before worksheet `Row`
  components. Required `cell` is an absolute A1 coordinate and `text` is plain
  text. `author?` defaults to `Author`, and `authorPrefix?` defaults to true;
  `visible?` defaults to false. `width?` and `height?` are positive integer
  pixels and default to 128 and 74. `backgroundColor?` is opaque RGB and defaults
  to `#FFFFE1`; `fontName?`, `fontSize?`, and `fontFamily?` default to `Tahoma`,
  8 points, and 0. `altText?` defaults to empty. Notes are unique by cell.
  Reverse preserves text, author/prefix, font metadata, fill, visibility, alt
  text, and integer-pixel dimensions. It verifies the VML anchor and offsets by
  regenerating the note with the worksheet's recovered row/column layout;
  positioning that cannot be reproduced is rejected with repair guidance.
  Rich-text note runs beyond the standard author-prefix form, non-note VML
  shapes, and `Note::set_object_movement()` cases that change the serialized
  position are not silently flattened.
- `autoFilterColumns?: { column: integer, values?: string[], includeBlanks?:
  boolean, customFilters?: { logic: "and" | "or", conditions: { operator?:
  string, value: string }[] } }[]` adds criteria to columns within that range.
  `column` is zero-based relative to the first column in `autoFilter`; indices
  must be unique and inside the range width. List criteria use `values` as an
  ORed string list and may include blanks with `includeBlanks`. Custom criteria
  support one or two conditions and OOXML operators `lessThan`, `notEqual`,
  `greaterThan`, `lessThanOrEqual`, and `greaterThanOrEqual`; omitted `operator`
  means equality. `value` preserves the exact OOXML text, including wildcard
  patterns emitted by string-match filters. Custom filters require explicit
  `logic`. A single condition uses `or`, while two conditions preserve either
  `and` or `or`. A filter entry
  must select exactly one of list criteria or `customFilters`. Filter criteria
  are metadata only; they do not hide rows in the generated worksheet.
  Reverse preserves supported `<filters>` and `<customFilters>` values,
  validates a matching hidden `_xlnm._FilterDatabase` name when present, and
  rejects date-group filters and unrelated defined names with repair guidance.
- `printArea?: string` sets the worksheet print area using an inclusive A1
  range (`A1:D20`), whole-row range (`1:20`), or whole-column range (`A:C`).
  References are one-based as in Excel. `repeatRows?: string` accepts an
  inclusive row range such as `1:3`; `repeatColumns?: string` accepts an
  inclusive column range such as `A:C`. Repeated rows and columns can be used
  together. Reverse resolves worksheet-local `_xlnm.Print_Area` and
  `_xlnm.Print_Titles` defined names, validates their formulas against
  `localSheetId`, and rejects duplicates, unsupported range shapes, or other
  defined-name metadata rather than dropping it.
- `header?: string` and `footer?: string` set the worksheet's odd-page header
  and footer using Excel's header/footer control-string syntax (`&L`, `&C`,
  `&R`, font directives, and page/date/file placeholders). Reverse canonicalizes
  long placeholders such as `&[Page]` to their equivalent short codes (`&P`)
  because the writer does so in OOXML. Text is limited to 255 characters after
  placeholder expansion. `headerFooterScaleWithDoc?: boolean` and
  `headerFooterAlignWithPage?: boolean` preserve the corresponding
  `headerFooter` attributes; their default is `true`, so reverse omits them.
  Images (`&G`/`&[Picture]`) and even/first-page header/footer variants are not
  representable and are rejected with repair guidance.
- `freezeRows?: integer` and `freezeColumns?: integer` are counts of leading
  rows/columns to freeze (default 0; maximum 1,048,575 rows and 16,383 columns).
  If either property is specified, at least one count must be positive.
  `freezeTopLeftCell?: string` optionally selects the first scrolling cell; it
  must be a valid A1 reference outside the frozen area. If omitted, the first
  non-frozen cell is used. Reverse preserves frozen panes and custom scroll
  origins. Screen gridline visibility is modeled separately by
  `showGridlines`; print-gridline settings, custom non-frozen split panes,
  noncanonical frozen-pane selections, and other unmodeled view state remain
  rejected rather than discarded. The `selection` worksheet property models
  the supported active selection separately.
- Unsupported worksheet metadata must be rejected rather than discarded.

### DataValidation

- A worksheet-level child that applies a list validation to one absolute A1
  `range`, such as `B2:B20`. It must appear before the first `Row`.
- `type` is one of `any`, `list`, `whole`, `decimal`, `date`, `time`,
  `textLength`, or `custom`. `any` leaves values unrestricted and is useful
  with input help messages; it does not accept `source`, formulas, or an
  operator. The pinned writer omits the OOXML `type` attribute for `any`, so
  reverse maps an absent type to `any`. Only the default `errorStyle="stop"`
  is representable for `any`; other error styles are rejected because the
  writer omits them. List validations use required `source: string`, the exact list
  formula (for example `source={'"Open,Closed"'}` or
  `source="Options!$A$1:$A$3"`). Other types use required `formula1: string`;
  `formula2?: string` is used only for a two-bound rule. `operator?: string`
  supports `between`, `notBetween`, `equal`, `notEqual`, `greaterThan`,
  `greaterThanOrEqual`, `lessThan`, and `lessThanOrEqual`. For custom rules,
  only `formula1` is allowed. Formulas are preserved as the exact OOXML text.
  Reverse preserves these supported types and rejects multiple ranges,
  extension metadata, and invalid formula/operator combinations with detailed
  errors.
- `allowBlank?: boolean` defaults to `true`; `showDropDown?: boolean`,
  `showInputMessage?: boolean`, and `showErrorMessage?: boolean` default to
  `true`. The OOXML `showDropDown` flag is normalized to user-facing behavior.
- Optional `inputTitle` and `errorTitle` are at most 32 XML-valid characters;
  `inputMessage` and `errorMessage` are at most 255 XML-valid characters.
  `errorStyle?: "stop" |
  "warning" | "information"` defaults to `"stop"`.
- Validation ranges may not overlap. Formula-based rules use Excel formulas or
  literal values as expected by the selected rule type; `formula2` is required
  for `between`/`notBetween` and forbidden for single-bound/custom/list rules.
  If `formula2` is supplied without an operator, the rule means `between`,
  matching the writer's canonical OOXML representation.

### ConditionalFormat

- A worksheet-level conditional rule applied to one contiguous A1 `range`; it
  must appear before the first `Row`. `type` is `colorScale2` or
  `colorScale3`, representing the public `ConditionalFormat2ColorScale` and
  `ConditionalFormat3ColorScale` writer APIs. Colors are opaque RGB `#RRGGBB`;
  two-color rules require `minColor` and `maxColor`, while three-color rules
  also require `midColor`. `stopIfTrue?: boolean` defaults to false.
- Threshold types are `min`, `max`, `number`, `percent`, `percentile`, and
  `formula`. A threshold type other than `min`/`max` requires its matching
  `minValue`, `midValue`, or `maxValue`. Numeric/percent/percentile values are
  finite numbers; formula values are Excel formula strings and remain exact.
  Percent and percentile thresholds must be in 0–100. Defaults match the
  writer: min/max endpoints for two-color scales and min/50th-percentile/max
  for three-color scales. `type="cellIs"` adds a classic comparison rule with
  `operator` set to `equal`, `notEqual`, `greaterThan`,
  `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`, `between`, or
  `notBetween`. `formula1` is required; `formula2` is required only for the
  two-bound operators. Formula text is emitted exactly. A cellIs rule must set
  at least one supported differential-format style. Conditional rule styles
  support `bold={true}`, `italic={true}`, `strikethrough={true}`,
  `underline="single|double|singleAccounting|doubleAccounting"`,
  `script="superscript|subscript"`, `fontColor="#RRGGBB"`, and
  `fillColor="#RRGGBB"` or `fillPattern` with optional opaque-RGB
  `fillForegroundColor` / `fillBackgroundColor`, `numberFormat`, and `borderTop`, `borderRight`,
  `borderBottom`, `borderLeft` with matching optional `border*Color` RGB values.
  Number formats use the same 1–255 XML-valid character contract as cell and
  table formats. Border values use supported OOXML edge styles (not `none`);
  an edge color requires a matching non-`none` edge style. The pinned writer
  omits a script-only differential font, so `script` must be paired with a
  font property such as emphasis, underline, or font color.
- `type="formula"` applies a boolean Excel formula rule. It requires exact
  `formula1` text and one or more of the same supported differential-style
  properties; `formula2`, operator,
  and color-scale thresholds are forbidden. Reverse preserves the formula text,
  style, stop flag, and rule priority. The writer removes one optional leading
  `=` because OOXML formula elements store formula text without it. Reverse
  rejects unsupported differential properties, extensions, and multi-range
  conditional formatting.
- `type="text"` applies a text comparison using `operator="containsText"`,
  `"notContains"`, `"beginsWith"`, or `"endsWith"`, with required `text` and
  one or more supported differential styles. Text is limited to 255 XML-valid
  characters and cannot contain a double quote because the pinned writer
  interpolates it into a formula string without escaping. Reverse verifies that
  the `formula` element is exactly the formula implied by `operator`, `text`,
  and the range's top-left cell; altered formulas are rejected rather than
  silently replaced.
- `type="dataBar"` applies the pinned `ConditionalFormatDataBar` writer API.
  `fillColor?: string` is opaque RGB `#RRGGBB` and defaults to `#638EC6`.
  Optional `minType` / `maxType`
  and `minValue` / `maxValue` define endpoints; types are `automatic`, `min`,
  `max`, `number`, `percent`, `percentile`, or `formula`. `barOnly`, `solidFill`,
  `border`, `direction`, `axisPosition`, `borderColor`, `negativeFillColor`,
  `negativeBorderColor`, and `axisColor` preserve x14 writer settings. Color and
  direction defaults match `rust_xlsxwriter`; reverse omits defaults. The x14
  rule is paired to its classic dataBar rule by GUID and range. Both halves
  must agree; orphaned, duplicate, or contradictory IDs and unsupported
  extension content are rejected with repair guidance. Classic dataBar rules
  without x14 extensions are reversed using canonical defaults; unknown
  attributes or children are rejected rather than discarded.
- `type="iconSet"` applies `ConditionalFormatIconSet`. `iconType` selects one
  of the public three-, four-, or five-icon sets (`threeArrows`,
  `threeArrowsGray`, `threeFlags`, `threeTrafficLights`,
  `threeTrafficLightsWithRim`, `threeSigns`, `threeSymbolsCircled`,
  `threeSymbols`, `threeStars`, `threeTriangles`, `fourArrows`,
  `fourArrowsGray`, `fourRedToBlack`, `fourHistograms`, `fourTrafficLights`,
  `fiveArrows`, `fiveArrowsGray`, `fiveHistograms`, `fiveQuadrants`, or
  `fiveBoxes`). It defaults to `threeTrafficLights`. `reverse?: boolean`,
  `showIconsOnly?: boolean`, and `stopIfTrue?: boolean` default to false.
  `icons?: IconThreshold[]` defines per-icon thresholds; the array length must
  match the icon set (3, 4, or 5). Each item has `type` of `number`, `percent`,
  `percentile`, or `formula`, a matching numeric or formula `value`, and
  optional `greaterThan` (default false for classic icon sets). The first
  threshold is fixed by the writer at percent 0 and must be specified as such.
  When an x14-only type or custom icon choice is used, `greaterThan: true` is
  invalid because the pinned writer omits it from x14 `cfvo`; its IR validator
  reports this before generation. Items may also set
  `iconType` and `iconIndex` to choose an icon from a built-in set, or `noIcon`
  to suppress the icon; those choices are mutually exclusive. Classic icon-set
  rules are read directly; custom icon assignments and `noIcon` are preserved
  from x14 `cfIcon` entries. x14-only rules are reconstructed from their
  extension range and priority. If both representations refer to the same
  rule, their thresholds and icon type are cross-checked. Orphaned, duplicate,
  or contradictory classic/x14 rules are rejected.
  Icon-set types that are x14-only are supported, but the writer does not
  serialize `reverse`, `showIconsOnly`, or `stopIfTrue` for
  those types, so setting any of those flags to true is invalid. Reverse keeps
  threshold order and rejects unrecognized extension metadata.
- `type="top10"` applies `ConditionalFormatTop`. `rule` is `top`, `bottom`,
  `topPercent`, or `bottomPercent` and defaults to `top`; `rank` defaults to 10
  and is an integer from 1–1000 for top/bottom counts or 1–100 for percent
  rules. `bold`, `italic`, `fontColor`, and `fillColor` optionally set a
  differential format; `stopIfTrue` defaults to false. Reverse preserves the
  writer's `top10` rank/percent/bottom attributes and supported differential
  styles, and rejects malformed ranks or unsupported style properties rather
  than guessing.
- `type="average"` applies `ConditionalFormatAverage`; `rule` defaults to
  `aboveAverage` and accepts `belowAverage`, `equalOrAboveAverage`,
  `equalOrBelowAverage`, and `one`, `two`, or `three` standard deviations
  above/below. Optional differential styles use the supported bold, italic,
  RGB font, and RGB fill subset. Reverse validates the exact `aboveAverage`,
  `equalAverage`, and `stdDev` attribute combination.
- `type="date"` applies `ConditionalFormatDate`; `rule` defaults to
  `yesterday` and accepts `today`, `tomorrow`, `last7Days`, `lastWeek`,
  `thisWeek`, `nextWeek`, `lastMonth`, `thisMonth`, and `nextMonth`. Reverse
  cross-checks `timePeriod` and the canonical writer formula against the range
  anchor; mismatches are errors. Optional differential styles are preserved.
- `type="blank"`, `type="error"`, and `type="duplicate"` apply the matching
  blank, error, and duplicate-value writers. Their `rule` values are
  `blank`/`nonBlank`, `error`/`nonError`, and `duplicate`/`unique`
  respectively, defaulting to the first value. Optional differential styles
  are preserved. Reverse checks canonical formulas and rejects malformed or
  contradictory type/formula combinations.
- Reverse preserves rule priority and rejects other conditional-format
  families, unsupported differential-style properties, unrecognized
  extensions, and
  multi-range conditional formatting.

### IgnoreError

- A worksheet-level error-checking exception with required `range` (one ordered
  A1 rectangle, such as `A1:B3`) and `errorType`; it must appear before the
  first `Row`. `errorType` is one of `numberStoredAsText`, `evalError`,
  `formula`, `emptyCellReference`, `formulaRange`, `listDataValidation`,
  `TwoDigitTextYear`, `unlockedFormula`, or `calculatedColumn`, matching the
  corresponding `ignoredError` OOXML flag and `rust_xlsxwriter::IgnoreError`
  variant. Ranges may not overlap, even when their flags differ. Reverse emits
  one component per `ignoredError` range and rejects unknown/multiple flags,
  non-true flag values, and malformed ranges instead of discarding them.

### Column

- Children: none. `min` and `max` are required 1-based Excel column numbers in
  the inclusive range 1–16,384, with `min <= max`.
- `widthPixels?: integer` sets a column's width in pixels, from 1 through 1,790.
  Pixel units match `rust_xlsxwriter::Worksheet::set_column_width_pixels()`;
  Excel's character-width serialization is quantized by the writer.
- `hidden?: boolean` hides the columns in the inclusive range. `false` is the
  default and is omitted by reverse. A hidden column with no explicit width
  uses the worksheet's default width when unhidden.
- `outlineLevel?: integer` sets the outline nesting level from 1 through 7.
  It composes with width and hidden state. Expanded row/column outline levels
  are preserved. `collapsed?: boolean` marks a one-column summary immediately
  after a contiguous run of hidden columns with positive outline levels; it
  defaults to false. A collapsed marker must use a single-column `min`/`max`
  range. The writer creates the corresponding group and hidden members.
- `Column` accepts the same format properties as `Cell` directly (font, fill,
  border, alignment, protection, and number format). They apply as the default
  format for unformatted cells in the inclusive column range;
  explicit cell formats override them. Reverse preserves supported OOXML
  `<col style>` values as these attributes. Column ranges must not overlap.
- Best-fit metadata is rejected: `rust_xlsxwriter` computes widths for autofit
  rather than serializing an equivalent OOXML `bestFit` setting.
- External `<col>` ranges are expanded and re-coalesced only when adjacent
  columns have identical supported properties. Widths that cannot be mapped
  exactly to the writer's integer-pixel representation and unsupported
  attributes are rejected rather than rounded.

### Row

- Children: `Cell` components only, in column order. Empty rows are positional
  placeholders; a trailing empty row is only meaningful when it has row metadata.
- `height?: number` sets row height in points. Values must be finite, from 0.75
  through 409.5, and a multiple of 0.75 points so the value survives the
  writer's pixel quantization without changing on reverse.
- `hidden?: boolean` hides the row when true. False is equivalent to the default
  visible state and is omitted by reverse.
- `outlineLevel?: integer` sets the expanded outline nesting level from 1
  through 7. It composes with height and hidden state. `collapsed?: boolean`
  marks a summary row immediately after the contiguous run of hidden rows with
  positive outline levels that it collapses; it defaults to false. The DSL
  requires the marker row to follow at least one such row. The writer creates
  the group, hides its members, and preserves nested outline levels.
- `Row` accepts the same format properties as `Cell` directly; they provide a
  default format for unformatted cells in that row. Explicit cell formats
  override them. Reverse preserves supported OOXML row `s` styles; unsupported
  style semantics still fail with repair guidance.
- Rows and cells are positional. Blank cells inside the used range and empty row
  placeholders before later data are preserved. A trailing empty row without
  row metadata is rejected because XLSX has no evidence that it existed.

### Cell

- `checkbox?: boolean` displays a boolean cell as an in-cell checkbox when true.
  It requires a boolean `value` (or boolean child), cannot be combined with a
  formula/date/hyperlink/rich text/merge, and defaults to false. Supported cell
  formatting composes with the checkbox. Reverse detects the checkbox style in
  external XLSX files, validates its Office feature-property metadata, and emits
  the property only for checkbox cells. Excel 2024 or later is required to
  render the checkbox; older versions retain the underlying TRUE/FALSE value.
- `formula?: string` describes an Excel formula. Pair it with
  `arrayFormulaRange?: string` to describe an array-formula group. The range is
  an ordered A1 range whose first cell must be this `Cell`'s coordinate;
  `formula` is required and may be prefixed with `=`. Omit the range only for a
  single-cell dynamic formula. `dynamicArray?: boolean` selects the Excel 365
  dynamic-array writer APIs; it defaults to false and is omitted by reverse.
  Dynamic arrays with a known spill range use `arrayFormulaRange`, while a
  range-less dynamic formula preserves a single-cell spill anchor. Array-formula
  cells cannot also contain `value`, `date`, `hyperlink`, rich-text children,
  `formulaResult`, or merged-cell spans. Known array ranges cannot overlap other
  array formulas or merged cells. Reverse distinguishes dynamic formulas using
  worksheet cell metadata and the workbook metadata part, preserving dynamic
  status and formula text; legacy arrays remain distinct. Shared and data-table
  formulas remain rejected.

- `rowSpan?: integer` and `columnSpan?: integer` define a merged rectangle
  beginning at the cell's positional coordinate. Both default to 1; each must
  be positive and remain within Excel's row/column limits. At least one must be
  greater than 1 to create a merge. Following cells in the same row and later
  rows skip coordinates occupied by the rectangle. The top-left cell owns the
  value and style; merged interior values or style differences from external
  workbooks are rejected because the DSL/writer cannot preserve them exactly.
  Reverse emits span properties only on the top-left cell and omits interior
  placeholders.
- Either scalar children, rich `TextRun` children, `value`, or a `formula`, never
  more than one representation. Rich text uses one or more `<TextRun>` children,
  each with exactly one non-empty string child. `TextRun` supports only font
  properties: `bold`, `italic`, `fontColor`, `fontName`, `fontSize`,
  `fontFamily`, `fontCharset`, `underline`, `strikethrough`, and `script`.
  Cell-level formatting remains independent and applies to the whole cell.
  Rich text cannot be combined with `value`, `date`, `formula`, or `hyperlink`.
  Empty runs, phonetic annotations, and unsupported run-level font properties
  are rejected during compile or reverse with a repair suggestion.
- `date?: string` is an alternative to `value`, formula, or children. It accepts
  `YYYY-MM-DD` and `YYYY-MM-DDTHH:mm:ss[.fraction]` (1–3 fractional digits)
  as a local date/time and compiles to an Excel date serial with the matching
  standard display format.
  Time zone suffixes are rejected because Excel serial dates do not retain a
  time zone. Reverse emits this property for date-formatted numeric cells and
  OOXML ISO date cells.
- `bold?: boolean` and `italic?: boolean` apply font emphasis to the cell and
  work with scalar values, dates, formulas, and blanks. The reverse reader emits
  only explicitly applied `true` flags; default/false values use the unstyled form.
- `horizontalAlign?: "left" | "center" | "right" | "distributed"` and
  `verticalAlign?: "top" | "center" | "bottom"` control the independent axes.
  Default/General horizontal and Bottom vertical alignment are omitted when
  reversing because they have the same effective behavior. Fill, justify,
  center-across and other unmodeled alignment metadata are
  rejected on reverse. `textRotation?: integer` accepts -90 through 90 degrees
  or 270 for stacked vertical text; ordinary negative angles are preserved
  using Excel's 91–180 OOXML encoding; the default is 0 and is omitted. The OOXML value 255 (stacked vertical
  text) reverses to 270.
- `readingDirection?: "leftToRight" | "rightToLeft"` sets the cell's reading
  order through `Format::set_reading_direction()` (`readingOrder=1` or `2`). The
  default contextual order (`readingOrder=0` or omitted) is omitted by reverse.
  Other reading-order values are rejected rather than normalized.
- `quotePrefix?: boolean` sets Excel's `quotePrefix` cell-format flag. `true`
  keeps text that looks like a formula treated as text when edited; `false` is
  the default and is omitted by reverse. It composes with other cell formats.
- `locked?: boolean` controls the cell's protection-format flag (default
  `true`; only `false` needs to be emitted), and `formulaHidden?: boolean`
  hides a formula's expression while preserving its result (default `false`).
  These flags only restrict editing or hide formulas when the worksheet also
  has `protection` enabled; reverse preserves them independently of that flag.
- `indent?: integer` sets the Excel cell indentation level from 1 through 255.
  It requires `horizontalAlign` to be explicitly `left`, `right`, or
  `distributed`; centered or implicit/general alignment would be changed by the
  writer. `shrinkToFit={true}` cannot be combined with distributed alignment
  because the writer suppresses shrink-to-fit in that mode. Reverse preserves
  the supported alignment and indent attributes and rejects incompatible or
  unrepresentable combinations.
- `wrapText?: boolean` toggles cell text wrapping. `false` is the default and
  is omitted by reverse; `true` composes with horizontal and vertical alignment.
  Reverse preserves `wrapText` and rejects other unsupported alignment metadata.
- `shrinkToFit?: boolean` toggles shrink-to-fit for cell content. `false` is the
  default and is omitted by reverse. It composes with horizontal/vertical
  alignment, but cannot be combined with `wrapText={true}`: the locked writer
  suppresses shrink-to-fit whenever wrapping is enabled. DSL validation and
  reverse reject that combination rather than silently dropping one setting.
- `numberFormat?: string` sets an Excel number-format code for values, formulas,
  blanks, and dates. It must contain 1–255 characters and valid XML
  characters. It can be combined with `date`; when present it overrides the
  canonical display format selected by `date`. Common examples include
  `"#,##0.00"`, `"0.00%"`, and `"$"#,##0.00`.
- Reverse preserves custom format codes verbatim and maps built-in IDs 1–22 and
  37–49 to their standard codes. Other nonzero built-in IDs are rejected because
  some are locale-dependent or reserved. Number-formatted dates, time-only
  values, and durations are represented as their numeric serial plus
  `numberFormat`; `date` is used for ISO-typed dates and the three canonical DSL
  date formats when that conversion preserves millisecond precision.
- Cell, Row, and Column accept fill properties. `fillColor?: string` remains the shorthand for a solid fill using an
  HTML-style RGB color `#RRGGBB`. For pattern fills, `fillPattern` is one of
  `none`, `solid`, `mediumGray`, `darkGray`, `lightGray`, `darkHorizontal`,
  `darkVertical`, `darkDown`, `darkUp`, `darkGrid`, `darkTrellis`,
  `lightHorizontal`, `lightVertical`, `lightDown`, `lightUp`, `lightGrid`,
  `lightTrellis`, `gray125`, or `gray0625`, matching OOXML `patternType` and
  `FormatPattern`. Optional `fillForegroundColor` and `fillBackgroundColor`
  provide opaque RGB `#RRGGBB` colors. The legacy `fillColor` shorthand is
  mutually exclusive with `fillPattern` and the foreground/background props;
  foreground/background colors require an explicit `fillPattern`. Reverse
  preserves supported pattern names and both colors exactly; theme, indexed,
  automatic custom colors and gradient fills remain rejected.
- `fontColor?: string` sets the cell font color as HTML-style RGB `#RRGGBB`.
  It composes with other supported cell styles. Reverse emits uppercase RGB;
  theme, indexed, automatic, and non-opaque alpha colors are rejected.
- `fontName?: string` selects a literal font name (not a theme font). It must
  contain 1–255 valid XML characters. Reverse preserves explicit font names;
  theme-linked font names and font metadata outside family/charset remain unsupported.
- `fontFamily?: integer` selects SpreadsheetML font family 0–5 (not applicable,
  roman, swiss, modern, script, decorative); 2 is the writer's default and may
  be omitted. `fontCharset?: integer` selects the OOXML charset byte 0–255; 0 is
  the default and may be omitted. Both are composed with the other font props,
  written with `Format::set_font_family()` / `set_font_charset()`, and reversed
  from explicit font records. Invalid family values and non-byte charsets are
  rejected; unsupported font metadata remains an error.
- `fontScheme?: "none" | "body" | "headings"` selects a non-theme font, or
  links the font to the workbook theme's minor (body) or major (heading)
  typeface. `Format::new()` defaults to `body`; reverse folds `body` to omission
  because it recompiles to the same writer default. `none` is emitted explicitly
  when needed to preserve a non-theme font without a literal `fontName`.
  `fontName` independently selects a literal name and upstream also clears the
  scheme when it is set; an explicit `fontScheme` takes precedence if both are
  supplied. Applies to cell, row/column, and rich-text run formats; differential
  table/conditional formats do not expose it in v1.
- `fontSize?: number` sets the font size in points, from 1 through 409 inclusive.
  Fractional sizes are preserved. Reverse emits the numeric size and rejects
  values outside this range. Size-only styles retain the standard Calibri theme
  fallback without emitting a redundant `fontName` property.
- `underline?: "single" | "double" | "singleAccounting" | "doubleAccounting"`
  selects the Excel font underline style. `strikethrough?: boolean` enables
  strike-through when true; false is equivalent to the default and is omitted
  by reverse. Both compose with other supported font properties. Reverse
  preserves all four underline variants and rejects unsupported font
  decorations or attributes.
- `script?: "superscript" | "subscript"` changes the cell font's vertical
  script using `Format::set_font_script()`. Reverse preserves both values from
  external font records. It is independent of underline, strikethrough, and
  other supported font properties; the same supported property is available
  per `TextRun`.
- `borderTop`, `borderRight`, `borderBottom`, and `borderLeft` independently
  set the corresponding edge. Values are `none`, `thin`, `medium`, `dashed`,
  `dotted`, `thick`, `double`, `hair`, `mediumDashed`, `dashDot`,
  `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, or `slantDashDot`.
  Optional `borderTopColor`, `borderRightColor`, `borderBottomColor`, and
  `borderLeftColor` set opaque RGB `#RRGGBB` colors and require the matching
  edge style. Reverse preserves each edge independently; theme/indexed/automatic
  custom colors and additional border metadata are rejected, while implicit
  automatic edge colors use the default and are omitted.
- `borderDiagonal?: borderStyle`, `borderDiagonalColor?: "#RRGGBB"`, and
  `borderDiagonalType?: "up" | "down" | "upDown"` represent the diagonal
  edge, its opaque RGB color, and direction. The color requires a non-`none`
  diagonal style. Reverse preserves the direction flags and diagonal edge;
  theme/indexed/automatic custom colors remain rejected.
- Workbooks using the 1904 date system are rejected: the writer currently only
  emits the 1900 system, and silently changing the epoch would shift date values.
- Font metadata other than the literal name, size, family, charset, emphasis,
  underline, strikethrough, script, and RGB color is rejected on reverse.
  Border metadata other than four edges and the diagonal edge/direction is also
  rejected rather than discarded.
- Scalar values: string, finite number, boolean, or `null` (blank). Child booleans
  must be wrapped as `{true}`/`{false}`; boolean JSX children outside cells are
  ignored like ordinary JSX conditional output.
- `formula` is an Excel formula string, with or without leading `=`; emitted
  formulas include `=`. `formulaResult?: string | number | boolean | null`
  preserves the ordinary formula's cached result; `null` represents an empty
  result, while omitting the property uses the writer's default result. The
  writer does not calculate formulas. String results that look numeric or are
  `TRUE`/`FALSE` are rejected because the writer serializes those as numeric or
  boolean results; use a number/boolean result instead. Cached date-typed
  formula values and formula results combined with `arrayFormulaRange` are not
  represented; reverse rejects external array formulas whose cached results are
  not the writer's default zero values rather than silently discarding them.
  Shared, data-table, and otherwise attributed formula structures are rejected.
- `hyperlink?: string` writes an external URL/URI relationship or an internal
  workbook location on a text cell. External targets include HTTP(S), `mailto:`,
  `file://`, and other writer-supported URI schemes. Internal targets use the
  `internal:` pseudo-URI, for example `hyperlink="internal:Summary!A1"` or
  `hyperlink={"internal:'Sales Data'!B2"}`. The anchor after `internal:` must be
  non-empty and XML-valid; it may name a cell, range, or workbook location.
  The visible cell text comes from `value` or its single string child. The target
  is limited to 2,080 characters and the display text to Excel's 32,767-character
  cell limit. `hyperlinkTooltip?: string` preserves the hyperlink screen tip;
  it requires `hyperlink`, must contain XML-valid characters, and an empty value
  is equivalent to omitting the tooltip. A hyperlink cannot be combined with
  `formula` or `date`. Reverse preserves internal locations, external hyperlink
  targets/fragments, and tooltip text. It rejects a hyperlink `display` value that conflicts with the cell
  text, unrepresented hyperlink metadata, and non-hyperlink worksheet
  relationships rather than dropping them.
- No additional cell properties are accepted.
- Apart from bold/italic, font name/size/color, edge borders, alignment, number
  formats, fill color, rich `TextRun` font formatting, the `date` property, and
  external/internal hyperlinks, v1 does not expose Excel error values.

### TextRun

- `TextRun` is valid only as a direct child of `Cell`. It has no nested JSX
  components and requires exactly one non-empty string child.
- `bold`, `italic`, `fontColor`, `fontName`, `fontSize`, `fontFamily`,
  `fontCharset`, `underline`, `strikethrough`, and `script` mirror the matching
  cell font properties and affect only that run. Fill, borders, alignment,
  number formats, and phonetic annotations are not run-level properties.
- `Cell` may contain one or more `TextRun` children, but cannot combine them
  with `value`, `date`, `formula`, or `hyperlink`. Cell-level styles still apply
  independently to the cell. Reverse preserves run order and boundaries for
  shared strings and inline strings; unsupported run font metadata fails with
  a repair suggestion.

## CLI contract

- `xlsx-jsx build <input.tsx> -o <output.xlsx> [--data <json>]`
- `xlsx-jsx reverse <input.xlsx> [-o <output.tsx>]`
- `xlsx-jsx validate <input.tsx> [--data <json>]` evaluates and validates DSL,
  but does not write an XLSX file.
- `xlsx-jsx spec [--format json|markdown]` emits the machine-readable contract.
- Errors identify the failing component/property and suggest a concrete repair.

## Reverse contract

Reverse must work on external XLSX files without a private embedded manifest.
The v1 lossless subset is worksheets, any-value, list and formula-based data
validations (whole, decimal, date, time, textLength, and custom), positional rows/cells, worksheet visibility,
active tab and reading direction, row height/hidden/outline state, collapsed
groups, summary direction, RGB worksheet tab colors, one active worksheet
selection and scroll origin, strings,
numbers, booleans, checkbox booleans, blanks, date/time serials and formats, literal font name/size/color, underline/strikethrough,
superscript/subscript, rich-text run font properties, edge borders, merged
cells, and formula text. If a workbook contains unsupported
semantics that cannot be represented, reverse must report them instead of
silently flattening or discarding them. The reader rejects common extra package
parts, named ranges, unsupported outline settings, multi-range or non-corner active-cell selections,
row/column styles,
and non-default cell styles outside the bold/italic/underline/strikethrough/script/font-name/size/color/edge-border/solid-fill/alignment/wrapText/shrinkToFit/number-format subset.
It also rejects Excel error cells, phonetic annotations, unsupported rich-text
run properties, unsupported row/column
metadata, unsupported built-in format IDs, the 1904 date system, and formula
structures instead of coercing or flattening them.
Non-default worksheet view state outside the listed subset and unmodeled authored core document
properties are rejected because v1 has no matching fields. Writer-generated
timestamps and workbook-derived app properties are package metadata, not
workbook content. Unknown package parts are rejected, including unknown parts
outside the `xl/` tree.
