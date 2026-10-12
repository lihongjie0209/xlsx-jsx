# xlsx-jsx Development Roadmap

Priorities favor common spreadsheet authoring workflows that can be represented
losslessly in both directions. Each delivered batch must follow `AGENTS.md`:
coverage audit, spec-first, red/green tests, external-file reverse coverage,
Open XML SDK validation, and its own commit. After push, verify GitHub Actions CI;
after CI succeeds, create and verify a GitHub Release and its assets. Continue
work on the next common feature while CI/release pipelines run, but report their
actual state accurately. Low-frequency, upstream-blocked, or lossy features stay
in this roadmap instead of blocking common workflows.

## Next priorities

1. Improve diagnostics and integration fixtures where current feature coverage
   exposes ambiguous or lossy external workbook structures.
2. Continue auditing public `rust_xlsxwriter` writer APIs, prioritizing common
   cell/table behavior; the recent audit found remaining table-format gaps are
   lower-frequency or missing an upstream serialization path.
3. Re-audit public writer APIs as batches land; full upstream API coverage is
   not claimed.

## Release prerequisite

The local checkout currently has no Git remote configured. Do not guess a remote
or create a public repository. Before the next push/release, the user must provide
or configure the intended GitHub repository remote; until then local batches can
be committed and fully tested, but GitHub CI and Release cannot be truthfully
reported as complete.

## Recently delivered

- Cell, row, and column formats now preserve horizontal justify alignment in
  both directions. Table-column justify DXFs remain deferred because the pinned
  writer can change explicitly formatted worksheet cells during compilation.
- Excel table autofilter criteria now support list-value filters and one/two
  custom comparisons with `and`/`or`; table-relative column indices are
  validated and reverse without a manifest. Date grouping, top-10, dynamic,
  color/icon, and other filter kinds remain deferred.
- The Excel 2023 Office/Aptos workbook theme now has an explicit preset and
  external reverse recognition; arbitrary custom theme payloads remain deferred
  because they also need matching default-format metadata.
- Reverse diagnostics now explain the legacy array-formula cache limitation and
  distinguish it from scalar cached results on ordinary formulas.
- Reverse diagnostics for unsupported Excel error tokens now direct users to the
  seven supported literal error values instead of suggesting that the already
  implemented `Cell.error` type be added.
- Image fit modes: `stretch`, `contain`, and `containCenter` use the effective
  dimensions of the target cell. Reverse emits canonical explicit pixel
  dimensions/offsets because OOXML records picture geometry, not the author's
  fit method.
- Worksheet `autoFit={true}` calculates final column widths from cell contents;
  the compiler removes the writer's transient `bestFit` marker and external
  reverse canonicalizes the resulting widths as explicit `Column` values.
  External best-fit markers without supported explicit widths remain
  unsupported, and `autoFit` cannot be combined with `Image.fit` because
  their geometry calculations conflict.
- Per-column `Column.autoFitWidthPixels` preserves an explicit width together
  with external `bestFit="1"`; it is distinct from content-measuring autofit.
  Reverse accepts only positive widths exactly representable as integer pixels.
- External `bestFit` now recognizes all XML boolean spellings and folds false
  into an ordinary explicit width without changing column behavior.
- External picture reverse now handles anchors crossing mixed custom/default
  row heights and column widths; exact whole-pixel anchor/transform agreement
  is still required.
- Data validation supports disjoint multi-range `sqref` values; range order is
  preserved and overlaps within/across rules are rejected with repair guidance.

## Deferred lower-frequency table differential formats

- Table-column `horizontalAlign="justify"`. The writer also applies table
  differential formatting to worksheet cells, so an external workbook with a
  different direct cell alignment cannot be rebuilt without changing that
  explicit cell style; external justify DXFs are rejected with repair guidance.
- Table-column `textRotation`, `readingDirection`, `indent`, and `shrinkToFit`.
  They are less common than alignment/wrapping and require more constraints for
  lossless dxf handling.
- False table-column differential overrides such as `wrapText={false}`. The
  pinned writer does not serialize these; support waits for an upstream writer
  path or a narrowly verified package repair with external reverse parity.

These items remain explicitly unsupported and must continue to produce a clear
validation/reverse error with repair guidance until implemented as a complete
forward/reverse batch.

## Audited but deferred workbook behavior

- Chartsheet `zoomToFit`: the pinned `Worksheet::set_zoom_to_fit()` API only
  affects chartsheets, so this is lower priority than worksheet/table authoring.
- External column `bestFit=true` metadata without a positive, exactly representable
  width remains rejected with a repair diagnostic.
