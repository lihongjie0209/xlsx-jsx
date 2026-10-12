# xlsx-jsx Development Roadmap

Priorities favor common spreadsheet authoring workflows that can be represented
losslessly in both directions. Each delivered batch must follow `AGENTS.md`:
coverage audit, spec-first, red/green tests, external-file reverse coverage,
Open XML SDK validation, and its own commit. CI and GitHub Release follow a
completed batch; they do not block work on the next batch.

## Next priorities

1. Audit remaining commonly used cell/table formatting gaps against the pinned
   `rust_xlsxwriter` serializer and reverse reader, then implement the next
   common lossless subset.
2. Extend image fit-to-cell geometry to custom row heights/column widths only
   after reverse can preserve those actual dimensions exactly.
3. Improve diagnostics and integration fixtures where current feature coverage
   exposes ambiguous or lossy external workbook structures.
4. Re-audit public writer APIs as batches land; full upstream API coverage is
   not claimed.

## Recently delivered

- Image fit modes: `stretch`, `contain`, and `containCenter` for default-size
  cells. Reverse emits canonical explicit pixel dimensions/offsets because
  OOXML records picture geometry, not the author's fit method. Custom-sized
  target cells remain rejected until exact external reverse support lands.

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
- Column best-fit/autofit behavior: `rust_xlsxwriter::autofit()` computes
  widths but does not serialize OOXML `bestFit`; the resulting widths are
  covered, while external `bestFit` metadata remains rejected because it
  cannot be reproduced as an editable auto-fit behavior.
