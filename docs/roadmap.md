# xlsx-jsx Development Roadmap

Priorities favor common spreadsheet authoring workflows that can be represented
losslessly in both directions. Each delivered batch must follow `AGENTS.md`:
coverage audit, spec-first, red/green tests, external-file reverse coverage,
Open XML SDK validation, and its own commit. CI and GitHub Release follow a
completed batch; they do not block work on the next batch.

## Next priorities

1. Review remaining commonly used cell/table formatting gaps against the pinned
   `rust_xlsxwriter` serializer and reverse reader. Only add properties when
   their serialized semantics can be reproduced without an IR manifest.
2. Improve diagnostics and integration fixtures where current feature coverage
   exposes ambiguous or lossy external workbook structures.
3. Re-audit public writer APIs and update the coverage table as each batch is
   completed; the project is not yet claiming full upstream API coverage.

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
