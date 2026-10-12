# xlsx-jsx Development Roadmap

Priorities favor common spreadsheet authoring workflows that can be represented
losslessly in both directions. Each delivered batch must follow `AGENTS.md`:
coverage audit, spec-first, red/green tests, external-file reverse coverage,
Open XML SDK validation, and its own commit. CI and GitHub Release follow a
completed batch; they do not block work on the next batch.

## Next priorities

1. Improve diagnostics and integration fixtures where current feature coverage
   exposes ambiguous or lossy external workbook structures.
2. Continue auditing public `rust_xlsxwriter` writer APIs, prioritizing common
   cell/table behavior; the recent audit found remaining table-format gaps are
   lower-frequency or missing an upstream serialization path.
3. Re-audit public writer APIs as batches land; full upstream API coverage is
   not claimed.

## Recently delivered

- Excel table autofilter criteria now support list-value filters and one/two
  custom comparisons with `and`/`or`; table-relative column indices are
  validated and reverse without a manifest. Date grouping, top-10, dynamic,
  color/icon, and other filter kinds remain deferred.
- The Excel 2023 Office/Aptos workbook theme now has an explicit preset and
  external reverse recognition; arbitrary custom theme payloads remain deferred
  because they also need matching default-format metadata.
- Reverse diagnostics now explain the legacy array-formula cache limitation and
  distinguish it from scalar cached results on ordinary formulas.
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
- External picture reverse now handles anchors crossing mixed custom/default
  row heights and column widths; exact whole-pixel anchor/transform agreement
  is still required.
- Data validation supports disjoint multi-range `sqref` values; range order is
  preserved and overlaps within/across rules are rejected with repair guidance.

## Deferred lower-frequency table differential formats

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
- External column `bestFit` metadata without a positive, exactly representable
  width remains rejected with a repair diagnostic.
