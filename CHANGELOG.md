# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

