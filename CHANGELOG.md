# Changelog

## 0.1.0 - 2026-10-13

Initial public release of `xlsx-jsx`, a Rust CLI for compiling JSX/TSX
workbooks to XLSX and reversing the supported external XLSX subset to JSX.

- Build, validate, reverse, and agent-readable spec commands.
- Lossless round trips for supported worksheets, cells, styles, formulas,
  validations, tables, images, charts, and workbook metadata.
- Actionable diagnostics for unsupported/lossy workbook structures and
  malformed XLSX packages.
- Incremental JSX/TSX transpilation cache and Microsoft Open XML SDK validation.
- Cross-platform release targets for Linux, macOS, and Windows.
