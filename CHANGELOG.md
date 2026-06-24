# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.2] - 2026-06-13

### Fixed
- **Wasm**: Fix save excel file.
- **Lint issues**: Fix lint issues.
- **XML parsing pipeline**: Refactored to use a streaming-based approach with iterative DOM building instead of full recursive parsing.
- **Performance**: Optimized cell coordinate parsing and XML value extraction for better performance.
- **Memory usage**: Reduced memory usage during XML parsing by avoiding full recursive parsing.

## [2.0.1] - 2026-06-11

### Fixed
- **Cell Style**: Fix underline and strikethrough preservation bug in CellStyle constructor and parsing logic.
- **Cell Style**: Add strikethrough text style support.
- **Cell Style**: Add comprehensive tests for underline and strikethrough preservation.
- **Cell Style**: Update Flutter example with underline and strikethrough demos.

## [2.0.0+2] - 2026-06-11

### Fixed
- **Wasm**: Fix save excel file.
- **Lint issues**: Fix lint issues.

## [2.0.0+1] - 2026-06-11

### Improved
- **XML parsing pipeline**: Refactored to use a streaming-based approach with iterative DOM building instead of full recursive parsing.
- **Performance**: Optimized cell coordinate parsing and XML value extraction for better performance.
- **Memory usage**: Reduced memory usage during XML parsing by avoiding full recursive parsing.

### Fixed
- **XML parsing: NullPointerException**: Fixed NullPointerException in `_XfCache._readXfs` method when processing shared strings with rich text formatting.



## [1.2.0] - 2026-06-04

### Fixed
- **Chart XML — series name (`c:tx`)**: The series name inside `<c:tx>` was incorrectly wrapped in `<c:strLit>`, which is not a valid child according to the OOXML `CT_SerTx` schema (§21.2.2.174). It now correctly uses `<c:v>` directly, fixing broken/missing series names when opening generated files in Microsoft Excel.
- **Chart XML — element order in `CT_Chart`**: `<c:autoTitleDeleted>` was emitted before `<c:title>`, violating the required sequence defined in OOXML §21.2.2.29. The order is now `<c:title>` → `<c:autoTitleDeleted>`, fixing chart validation errors in strict Excel readers.
- **`standard_21` format code**: Was `'h:mm:dd'` (invalid), corrected to `'h:mm:ss'` per ECMA-376 §18.8.30.
- **`standard_40` format code**: Was `'#,##0.00;[Red](#,#)'` (truncated), corrected to `'#,##0.00;[Red](#,##0.00)'`.

### Added
- **Complete OOXML standard number format table** (ECMA-376 §18.8.30): added previously missing built-in format IDs:
  - Currency formats `standard_5`–`standard_8` (`$#,##0`, `$#,##0.00` with normal/red negative variants).
  - Reserved/fallback formats `standard_23`–`standard_26` (mapped to `General`).
  - CJK locale date/time formats `standard_27`–`standard_36`.
  - Accounting with fill-character formats `standard_41`–`standard_44` (`_(* …)` / `_($* …)` patterns).
- **Parser refactor**: split monolithic `parse.dart` into `styles_parser.dart` (style/XF resolution) and `worksheet_parser.dart` (row/cell parsing) for improved maintainability.

## [1.1.4] - 2026-05-11

### Added
- Image embedding support: embed PNG, JPEG, GIF, BMP, TIFF, WMF, EMF, SVG, WebP, and ICO images into worksheets.
- `ExcelImage` model class with `fromFile` factory for automatic type detection from file extension.
- `ExcelImageType` enum with `fromExtension` factory covering 10 image formats.
- `ImageAnchor` class with `fromPixels` factory (96 DPI) and direct EMU constructor.
- `SheetImages` extension on `Sheet` exposing `addImage()` and `images` getter.
- `_ImageManager` save-pipeline component that writes OOXML drawing, relationship, and media files.
- New example `example/excel_images.dart` demonstrating PNG and SVG embedding.
- README "Images" section with full usage examples and supported-formats table.



## [1.1.3] - 2026-05-07
- Fix Excel Community logo.

## [1.1.2] - 2026-05-07
- Added logo for the package in the README.md.

## [1.1.1] - 2026-05-07
- Add Excel Community logo to pub.dev.

## [1.1.0] - 2026-05-06
- Update logo to open source version.
- Improve package README.md and add a preview image.
- Update package metadata for pub.dev.
- Fix chart x-axis labels for scatter charts.

## [1.0.10] - 2026-05-02
- Fix Losing Cell styling when setting a new value to a cell.

## [1.0.9] - 2026-04-18
- Fix rich-text bold/italic/underline parsing per ECMA-376

## [1.0.8] - 2026-04-07
- add funding text
## [1.0.7] - 2026-04-07

- Fix underline preservation bug in CellStyle constructor and parsing logic.
- Add strikethrough text style support.
- Add comprehensive tests for underline and strikethrough preservation.
- Update Flutter example with underline and strikethrough demos.
## [1.0.6] - 2026-03-17
- Fix color accessibility.
## [1.0.5] - 2026-03-17
- Apply clean code fix minors bugs.
## [1.0.4] - 2026-03-15
- Apply Clean Architecture in Save class.
- Refactor Sheet class.
- Refactor Color class.
- Refactor Border class.
- Refactor Formula class.
- Refactor NumberFormat class.

## [1.0.3] - 2026-03-15

- Renamed main example to `example/main.dart` for better visibility on pub.dev.
- Significantly increased documentation coverage (>28%) across the public API.
- Added detailed documentation for all chart types and cell value classes.

## [1.0.2] - 2026-03-15

## [1.0.1] - 2026-03-14

- Reorganized tests: moved manual debug scripts to `test/manual/`.
- Fixed static analysis errors in `parse.dart` and `save_file.dart`.
- Improved package metadata for pub.dev.

## [1.0.0] - 2026-03-14

- Initial Release of the community fork.
- Forked from [excel](https://github.com/justkawal/excel)
- Updated minimum Dart SDK to 3.6.0
- Migrate from `dart:html` to `package:web`
- Update packages to latest versions
- Added support for charts

