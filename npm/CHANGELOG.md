# Changelog

All notable changes to the `excel-community` npm package are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.6.0] - 2026-10-07
### Fixed
- Fixed `appendRow` getting slower as the sheet grows.
- Fixed tables with `showTotalsRow` overwriting the last row of data.
- Fixed files that Excel could not open after re-saving (page breaks, phonetic settings, ignored errors).
- Fixed headers and footers with different odd, even or first pages.
- Fixed pie chart colors changing on every save and failing with more than 20 slices.
- Fixed pivot table cells being empty until the workbook is saved.
- Fixed errors without a name; they are now `ExcelArgumentError`, `ExcelStateError` and `ExcelFormatError`.
- Fixed `cell.style` returning values that `setStyle` does not take (`'FF1F4E78'`, `'Center'`).

### Improved
- Improved save, read and `sheet.cell()` speed.

### Added
- `sheet.appendRows()` to load many rows in one call.
- `sheet.refreshPivotTables()` to recompute pivot tables after changing their data.
- Benchmark against exceljs (`npm run bench`) and a Limitations section in the README.

### Changed
- A `Cell` points to a position, so it no longer follows its data when rows are inserted or removed.

## [2.5.3] - 2026-10-06
### Fixed
- Fixed the package in browsers and bundlers (Vite, webpack, esbuild).
- Fixed `sheet.unmerge('A1')` doing nothing.
- Fixed duplicated charts, images and styles when saving the same workbook twice.
- Fixed line and area charts that Excel could not open.
- Fixed pivot tables with several values that Excel could not open.
- Fixed pivot tables showing up empty outside Excel (Protected View, LibreOffice, Google Sheets).
- Fixed header and footer text being escaped twice.
- Fixed `Sheet1` being renamed and losing changes when another sheet is added.

### Added
- Browser build for `<script>` tags and an example page.
- Conditional formatting, data validation, printing, pivot tables, charts, tables, export/import and the rest of the API.

### Changed
- `cell.setStyle()` keeps the options it is not given, and `sheet.protect()` options mean "allow".

## [2.5.2] - 2026-10-04
### Added
- Freeze panes (`frozenRows`, `frozenColumns`) and hidden rows and columns.

## [2.5.1] - 2026-10-04
### Added
- First release: the excel_community engine compiled to JavaScript for Node.js and browsers, with TypeScript types.
