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

- Reverse diagnostics now explain the legacy array-formula cache limitation and
  distinguish it from scalar cached results on ordinary formulas.
- Image fit modes: `stretch`, `contain`, and `containCenter` use the effective
  dimensions of the target cell. Reverse emits canonical explicit pixel
  dimensions/offsets because OOXML records picture geometry, not the author's
  fit method.
- Worksheet `autoFit={true}` calculates final column widths from cell contents;
  the compiler removes the writer's transient `bestFit` marker and external
  reverse canonicalizes the resulting widths as explicit `Column` values.
  Arbitrary external `bestFit` flags remain unsupported, and `autoFit` cannot
  be combined with `Image.fit` because their geometry calculations conflict.
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
- External column `bestFit` metadata remains rejected: replaying it may produce
  different widths. The supported one-shot `Worksheet.autoFit` operation is
  documented above and reverses to explicit final widths.
