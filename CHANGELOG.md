# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.5.0] - 2026-10-03
### Added
- **Page Setup & Print Configuration (`<pageSetup>`)**: OpenXML/SpreadsheetML support for worksheet print settings.
  - **`PageSetup` Model**: Orientation (`PageOrientation`), paper size (`PaperSize`), custom `paperWidth`/`paperHeight`, `scale`, fit-to-page (`fitToPage`, `fitToWidth`, `fitToHeight`), `firstPageNumber`/`useFirstPageNumber`, `pageOrder`, `blackAndWhite`, `draft`, `cellComments`, `errors`, `horizontalDpi`/`verticalDpi`, `copies` and `usePrinterDefaults`. Only explicitly set attributes are written.
  - **`PaperSize`**: Named constants for common sizes (`letter`, `legal`, `tabloid`, `a3`, `a4`, `a5`, `b4`, `b5`, envelopes, ...) with millimetre dimensions; any other code is kept via `PaperSize.fromCode`.
  - **`PageMargins`**: `<pageMargins>` in inches with `normal`, `wide` and `narrow` presets, `PageMargins.all()` and `PageMargins.fromCentimeters()`.
  - **`PrintOptions`**: `<printOptions>` for gridlines, row/column headings and horizontal/vertical centering.
  - **Sheet API**: `sheet.pageSetup`, `sheet.setPageOrientation()`, `sheet.setPaperSize()`, `sheet.setPrintScale()`, `sheet.fitToPages(width:, height:)`, `sheet.clearPageSetup()` / `removePageSetup()`, `sheet.hasPageSetup`, `sheet.pageMargins`, `sheet.clearPageMargins()`, `sheet.printOptions`, `sheet.setPrintGridLines()`, `sheet.setPrintHeadings()`, `sheet.setPrintCentered()`, `sheet.clearPrintOptions()`, `sheet.hasPrintOptions`.
  - **Interactive Flutter Example**: Added a "Page Setup & Printing" wiki to `excel_flutter_example`: every paper size drawn to scale with a portrait/landscape toggle and search, plus Orientation & Scaling, Margins and Print Options tabs. Each card shows a live page preview and its own snippet. The generated workbook contains an overview plus one worksheet per paper size.
  - **Parsing & Preservation**: The SAX worksheet parser reads `<pageSetup>`, `<pageMargins>`, `<printOptions>` and `<sheetPr><pageSetUpPr fitToPage>`. On save, the printer settings `r:id` and the `autoPageBreaks` flag are preserved, and elements are written in CT_Worksheet / CT_SheetPr schema order. Settings are kept when copying or renaming sheets.
- **Data Export & Transformation (`rowsAsMaps`, `toJson`, `toCsv`)**: Convert worksheet data to Dart collections, JSON or CSV, and import rows from maps.
  - **`sheet.rowsAsMaps()`**: One map per row keyed by the header row (`headerRow`, default `0`). Empty headers are named by column letter and repeated headers get a numeric suffix (`Name_2`); empty rows are skipped by default (`skipEmptyRows`).
  - **`ExportValueMode`**: `typed` returns native Dart values (`String`, `int`, `double`, `bool`, `DateTime` (UTC) for dates, `Duration` for times, the cached value for formulas); `displayText` returns the text Excel shows with the cell's number format.
  - **`sheet.rowsAsValues()`**: The whole grid as a list of value lists, header included.
  - **`sheet.toJson()`** / **`excel.toJson()`** / **`excel.toMaps()`**: JSON array of objects per sheet (or an object keyed by sheet name for the workbook), with dates as `YYYY-MM-DD`, date-times as ISO 8601 and times as `HH:MM:SS`; optional `indent` for pretty printing.
  - **`sheet.toCsv()`**: RFC 4180 CSV with configurable `separator` and `lineTerminator`; cells are written as displayed by default.
  - **`sheet.appendRowsFromMaps()`**: Writes maps as rows, adding the header row on an empty sheet, matching existing headers and appending new columns for unknown keys. `DateTime` values without a time become date cells and `Duration` values become time cells.
- **Data Validation & Dropdowns (`<dataValidation>`)**: Read, write and edit validation rules.
  - **`DataValidation` Model**: `DataValidationType` (list, whole number, decimal, date, time, text length, custom, any), `DataValidationOperator` (between, notBetween, equal, notEqual, lessThan(OrEqual), greaterThan(OrEqual)), input message (`promptTitle`/`prompt`), error message (`errorTitle`/`error`) and `DataValidationErrorStyle` (stop, warning, information), plus `allowBlank` and `showDropdown`.
  - **Factories**: `DataValidation.list([...])` (validated against Excel's comma/quote and 255-character limits), `listFromRange('A1:A10', sheetName:)` (absolute, quoted references), `wholeNumber`, `decimal`, `date` (Excel serial dates), `time`, `textLength`, `custom('formula')` and `inputMessage`; builders `withPrompt()` and `withError()`.
  - **`accepts(CellValue)`**: Evaluates a rule in Dart (case-insensitive list matching, number/date/time/length comparisons, blanks); returns `null` for formula or range based rules.
  - **Sheet API**: `sheet.addDataValidation(range, rule)` (one or several space-separated ranges), `setDataValidation(cell, rule)`, `getDataValidation`, `removeDataValidation(range)`, `clearDataValidations`, `dataValidations`, `hasDataValidations`, and `cell.dataValidation` getter/setter. Since a cell has one validation, a new rule's area is subtracted from overlapping rules (split into rectangles).
  - **Parsing & Preservation**: `<dataValidations>` are read with the SAX parser and written in schema order; Excel 2010 `x14:dataValidation` rules inside `<extLst>` are kept untouched. Rules move with their cells on row/column inserts and removals.
- **Cell Hyperlinks (`<hyperlinks>`)**: Read, write and edit hyperlinks on cells and ranges.
  - **`Hyperlink` Model**: External `url` (web pages, `mailto:` addresses, files) and/or internal `location` (cells, ranges, defined names) with `tooltip` and `display`. Factories `Hyperlink.url()`, `Hyperlink.email(subject:)`, `Hyperlink.cell(sheet, ref)` (quotes sheet names when needed) and `Hyperlink.location()`.
  - **Sheet API**: `sheet.setHyperlink(cell, link, text:, styled:)`, `sheet.setHyperlinkRange('A1:C1', link)`, `sheet.getHyperlink()`, `sheet.removeHyperlink()`, `sheet.clearHyperlinks()`, `sheet.hyperlinks`, `sheet.hasHyperlinks`, and `cell.hyperlink` getter/setter. Empty cells get the link text, and new links use Excel's hyperlink look (blue, underlined) unless `styled: false`.
  - **Relationships**: External targets are written to the worksheet `.rels` with `TargetMode="External"`; on re-save, hyperlink relationships are rebuilt while every other relationship is kept.
  - **Row/column edits**: Links move with their cells on `insertRow`, `removeRow`, `insertColumn` and `removeColumn` (links on a removed row/column are dropped, ranges shrink or grow).
- **Example App - Interactive Wikis**: The Fonts & Styles, Number Formats and Page Setup sections of `excel_flutter_example` share one wiki layout (`widgets/wiki/wiki_components.dart`): header with the generate button, tabs, search and a grid of cards, each with a live preview and its own copyable code snippet.
  - **Number Formats Wiki**: All built-in number formats (IDs 0-49, except the reserved 23-26) plus common custom codes (phone template, thousands/millions scaling, currency and unit suffixes, custom dates), grouped into Numbers, Currency & Accounting, Percent & Scientific, Date & Time, CJK Locale and Custom tabs, with search across categories. Previews are rendered by `NumFormat.format()` and show `[Red]`/`[Blue]` section colors. The demo workbook has one worksheet per category with the formatted cell next to the output Excel shows.
  - **Fonts & Styles Wiki**: Migrated to the shared layout with no visual changes.
  - **Data Validation Wiki**: Dropdown, number/date/time and text/custom examples rendering the cell's dropdown and input message, with sample values checked live by `accepts()`. The demo workbook has a "Tasks" sheet with 11 validated columns and a "Lookup Lists" source sheet.
  - **Hyperlinks Wiki**: External, internal and management examples; each card runs its snippet, saves and reopens the workbook, and shows the resulting cell and links. The demo workbook links between a "Links" sheet and a "Q1 Sales" sheet.
  - **Data Export Wiki**: Export and Import tabs that run `rowsAsMaps`, `rowsAsValues`, `toJson`, `toCsv` and `appendRowsFromMaps` on a sample workbook and show the live output next to each snippet. The demo workbook includes the JSON and CSV exports and a sheet imported from JSON.

### Fixed
- **Worksheet relationships of reopened files**: Saving a file read from disk no longer rewrites a worksheet's `.rels` from scratch. Previously, adding a comment, image, chart or pivot table to a reopened file dropped the sheet's existing relationships (e.g. its drawing), leaving `<drawing r:id>` pointing at the wrong part, which Excel reports as a corrupted file. Existing `.rels` and drawing parts are now loaded and extended, new relationship ids never collide (`rIdN` = highest + 1), and new drawings, charts (`xl/charts/chartN.xml`) and media (`xl/media/imageN.*`) are numbered after the ones in the original file instead of overwriting them.
- **Default sheet renaming**: A `Sheet1` that only contains images, hyperlinks or data validations is no longer renamed when another sheet is first accessed.
- **Boolean display text**: `Data.displayText` / `NumFormat.format()` now render booleans as `TRUE` / `FALSE` (as Excel does) instead of `1` / `0`.
- **`NumFormat.format()` / `Data.displayText` number format rendering** now matches Excel for codes that were previously rendered incorrectly:
  - **Fractions** (`# ?/?`, `# ??/??`, `?/?`, `# ?/8`): closest fraction for the denominator digits (or the fixed denominator), Excel-style padding, blank fraction for whole numbers. Before: `1235 ?/?`; now: `1234 4/7`.
  - **Engineering notation** (`##0.0E+0`): exponent in multiples of three (`123.5E+6`). Scientific notation no longer renders `10.00E+00` when rounding carries (`9.999` -> `1.00E+01`).
  - **Digit templates and scaling**: literals between placeholders keep their position (`000-000-0000` -> `555-123-4567`), trailing commas scale by thousands (`#,##0,"K"` -> `1,235K`, `0.0,,"M"` -> `12.3M`), and `?` placeholders pad with spaces, so accounting zero sections (`"-"??`) render a dash instead of `-??`.
  - **CJK dates and times**: `e` renders the era year for the locale tag (`[$-404]` Republic of China year, `[$-411]` Japanese era year; otherwise the Gregorian year). `上午/下午` and `午前/午後` work as AM/PM designators, and `AM/PM` is localized for zh, ja and ko locale tags (`115/10/3 2:30 下午`, `下午2時30分`).
  - **Fractions of a second** (`ss.0`, `ss.00`, `ss.000`) render the milliseconds, and elapsed-time codes (`[h]`, `[mm]`, `[ss]`) count from Excel's day zero for date values and include the hours in `[mm]`/`[ss]`.
- **Example App - Sidebar**: Menu items no longer throw "ListTile background color or ink splashes may be invisible" on every build. Each item now paints its selected background and tap ripple on its own `Material`.

## [2.4.2] - 2026-10-01
### Fixed
- **`Sheet.getColumnWidth()` / `Sheet.getRowHeight()` no longer throw** `Null check operator used on a null value` on sheets without a default width/height (e.g. sheets created with `excel['New sheet']` or decoded from files without `<sheetFormatPr>` defaults). They now fall back to the sheet default and then to Excel's defaults (column width `8.43`, row height `15.0`), the same values already used when saving.

## [2.4.1] - 2026-09-25
### Added
- **Sheet Tab Colors (`<tabColor>`)**: Full OpenXML/SpreadsheetML support for colored worksheet tabs.
  - **`TabColor` Domain Model**: Encapsulates `<tabColor>` element with support for 8-digit ARGB hex (`#RRGGBB`, `RRGGBB`, `#AARRGGBB`, `AARRGGBB`, and 3-char shorthand `#RGB`), `ExcelColor` presets, Office theme color indices (`theme`), tints (`tint`), indexed colors (`indexed`), and auto colors (`auto`).
  - **Sheet API**:
    - `sheet.tabColor` (getter and setter)
    - `sheet.setTabColor(ExcelColor color)`
    - `sheet.setTabColorHex(String hex)`
    - `sheet.clearTabColor()` and alias `sheet.removeTabColor()`
    - `sheet.hasTabColor` (boolean getter)
  - **SAX Streaming Parser**: Event-based streaming parser in `_WorksheetParser` extracts `<tabColor>` attributes (`rgb`, `theme`, `tint`, `indexed`, `auto`) directly with zero memory overhead.
  - **Schema Compliance**: Serialized inside `<sheetPr>` as the very first child element preceding `<outlinePr>` and `<pageSetUpPr>` in strict compliance with ECMA-376 Part 4 CT_SheetPr sequence rules.
  - **Attribute & Element Preservation**: Preserves existing `<sheetPr>` attributes (`codeName`, `filterMode`, etc.) and non-color child elements (`<outlinePr>`, `<pageSetUpPr>`) across read/write cycles.
  - **Multi-Sheet Management**: Setting tab color marks a sheet as customized, preventing it from being inadvertently renamed or discarded when adding additional sheets.
  - **Interactive Flutter Example**: Added "Sheet Tab Colors (`<tabColor>`)" section to `excel_flutter_example` with real-time UI preview showing colored tabs, source code snippets, and live XLSX file generation.

- **AutoFilter (`<autoFilter>`)**: Full OpenXML/SpreadsheetML support for worksheet auto-filters.
  - **`AutoFilter` Model**: Represents `<autoFilter ref="...">` with normalized coordinates, start/end cells (`startCell`, `endCell`), dimensions (`rowCount`, `columnCount`), and containment checks (`containsCell`, `containsCellId`).
  - **`FilterColumn` & `CustomFilterRule`**: Column criteria configuration (`<filterColumn>`), button visibility flags (`hiddenButton`, `showButton`), matching values (`<filters><filter val="..."/></filters>`), blank filtering, and custom comparison rules (`<customFilters>`).
  - **Sheet API**:
    - `sheet.setAutoFilter(CellIndex start, CellIndex end, {List<FilterColumn>? filterColumns})`
    - `sheet.setAutoFilterByString(String range, {List<FilterColumn>? filterColumns})`
    - `sheet.clearAutoFilter()` and alias `sheet.removeAutoFilter()`
    - `sheet.hasAutoFilter` (boolean getter)
    - `sheet.addFilterColumn(FilterColumn filterColumn)`
  - **SAX Streaming Parser**: Event-based parsing in `_WorksheetParser` supporting self-closing `<autoFilter ref="..."/>` and child elements with zero memory overhead.
  - **Schema Compliance**: Serialized in exact ECMA-376 schema order (immediately following `sheetProtection` and preceding `sortState` and `mergeCells`).
  - **Round-Trip Preservation**: Full round-trip preservation of filter ranges and custom/advanced filter XML when opening, modifying, and saving existing spreadsheets.
  - **Flutter Example**: Added new "AutoFilter (`<autoFilter>`)" showcase section in `excel_flutter_example` with live code, interactive preview, and downloadable XLSX file.

## [2.4.0] - 2026-09-09
### Added
- **`FormulaCellValue.cachedValue`**: formula cells now retain the pre-calculated `<v>` result found in the source file (instead of discarding it), so the last value Excel computed can be read without recalculating the formula.
- **`Data.displayText`** and **`NumFormat.format(CellValue?)`**: render a cell's value as the display text a spreadsheet application would show, based on its assigned number format (e.g. `1234.5` with `NumFormat.standard_7` → `"$1,234.50"`). Best-effort coverage of the ~50 built-in standard formats plus common custom currency/percentage/date/time patterns.
- **Flutter Example**: new "Formulas & Display Text" showcase section demonstrating `FormulaCellValue.cachedValue` and `Data.displayText` against a bundled fixture with real cached formula results.

## [2.3.0] - 2026-08-22
### Added
- **New Standard Chart Types**:
  - **`BubbleChart` (`<c:bubbleChart>`)**: 3-dimensional data comparison plotting X, Y, and Bubble Size with configurable `bubbleScale` and `showNegativeBubbles`.
  - **`StockChart` (`<c:stockChart>`)**: Financial stock price fluctuations (High-Low-Close, Open-High-Low-Close) with `showHighLowLines` and `showUpDownBars`.
  - **`OfPieChart` (`<c:ofPieChart>`)**: Pie-of-Pie and Bar-of-Pie secondary breakdown charts with configurable `ofPieType` (`pie`/`bar`), `splitType` (`position`/`value`/`percent`), and `secondPieSize`.
- **Chart Grouping & Subtypes (`ChartGrouping`)**:
  - Added `grouping` parameter (`clustered`, `stacked`, `percentStacked`) to `ColumnChart`, `BarChart`, `AreaChart`, and `LineChart`.
  - Added `showMarkers` and `smooth` (spline curves) to `LineChart`.
  - Added `showLines`, `showMarkers`, and `smooth` (spline curves) to `ScatterChart`.
- **Chart Data Labels (`ChartDataLabels`)**: Every chart type now accepts an optional `dataLabels` parameter that renders labels on each data point. Supported components: `value`, `categoryName`, `seriesName`, and `percentage` (pie / doughnut). Custom `separator` and `labelPosition` are also supported. The XML round-trip is preserved (labels written to `<c:dLbls>` are re-read via `ChartXmlWriter.parseDataLabelsFromXml`).
- **Chart Color & Style Customization (`ChartSeriesStyle`, `ChartFillType`)**: Per-series visual customization for all chart types via `ChartSeries.style`.
  - **Fills**: Support for solid fills (`ChartFillType.solid`), customizable transparency with `fillAlpha` (`ChartFillType.transparent`), and border-only no-fill (`ChartFillType.none`).
  - **Borders**: Customizable `borderColor`, `borderAlpha`, and `borderWidth` in EMUs.
- **Flutter Example & Documentation**:
  - Added interactive sections in `excel_flutter_example` for **Bubble, Stock & Stacked Charts**, **Chart Data Labels**, and **Chart Color Customization**.
  - Updated `README.md` with complete documentation, parameter tables, and worked examples for all 11 chart types.

## [2.2.1] - 2026-08-09
### Fixed
- **Merge Cells**: Fix merge cells parsing and saving.
## [2.2.0] - 2026-07-20
### Added
- **Conditional Formatting (`<conditionalFormatting>`)**: Native OpenXML/SpreadsheetML support for Conditional Formatting rules and Differential Styles (`<dxfs>`).
  - Support for numeric ranges (`cellIs`), text matching (`containsText`, `notContains`, `beginsWith`, `endsWith`), custom formula expressions (`expression`), and duplicate/unique values (`duplicateValues`, `uniqueValues`).
  - Multi-sheet support with worksheet-wide unique rule priority sequencing.
  - Round-trip preservation when reading, saving, and editing `.xlsx` files.
- **Example & Flutter App Integration**: Added standalone CLI example (`example/excel_conditional_formatting.dart`), documentation snippet, and a new interactive section in the Flutter demo app (`excel_flutter_example`).

## [2.1.9] - 2026-07-19
### Added
- **Cell Comments**: Added support and usage instructions for Cell Comments.
- **Flutter example app**: New `Cell Comments` section displaying a sheet with formatted comments and hover preview tooltips.

## [2.1.8] - 2026-07-12
### Added
- **Pivot Table** Add Official support and usage instructions for Pivot Tables.

## [2.1.7] - 2026-07-09
### Added
 - **Show Icon** show excel community logo.
 - **Chage Tags** update tags to excel_community.
## [2.1.6] - 2026-07-06
### Fixed
- **SST parsing**: Fixed the SST parsing logic to correctly read the shared strings from the SST stream.

## [2.1.5] - 2026-07-06

### Added
- **Legacy XLS File Support**: Added read-only support for parsing legacy `.xls` (Excel 97-2003) workbooks using a custom, clean-room OLE2 and BIFF8 parser implemented from scratch in pure Dart. It decodes sheet names, row/column dimensions, and cell values (`TextCellValue`, `IntCellValue`, `DoubleCellValue`), with transparent fallback inside `Excel.decodeBytes` and `Excel.decodeBuffer` using signature magic bytes detection.

## [2.1.4] - 2026-07-03
- **Improve documentation**

## [2.1.3] - 2026-07-02

### Added
- **Hidden Columns & Rows**: New `Sheet.setColumnHidden(index, hidden)`, `Sheet.isColumnHidden(index)`, `Sheet.setRowHidden(index, hidden)`, and `Sheet.isRowHidden(index)` methods to toggle and query visibility of columns/rows. Full serialization to OOXML `<col hidden="1">` and `<row hidden="1">` on save, and parsing from XLSX on decode.
- **Flutter example app**: New `Hidden Columns & Rows` sidebar demo displaying a multi-sheet spreadsheet where some sheets contain hidden elements and others are fully visible.

### Fixed
- **Cleaned Template**: Removed stale empty drawing (`xl/drawings/drawing1.xml`) and its worksheet relationship from the base sheet template base64 string inside `constants.dart`. This ensures newly created worksheets (such as `Sheet2`) that have charts do not bleed/render their charts onto the first sheet (`Sheet1`) of the workbook, whilst preserving pre-existing charts when editing loaded files.

## [2.1.2] - 2026-06-29

### Fixed
- **Multi-Sheet Freeze Panes**: Added hidden columns to the Staff Salaries and Executive Perks sheets, with corresponding data removed. Now only the Public Directory sheet is visible, while the other two remain hidden (but fully functional) in the generated XLSX file.

- **Multi-Sheet Charts**: Fixed chart issue in multi-sheet workbooks.

## [2.1.1] - 2026-06-29

### Added
- **Freeze Panes**: New `Sheet.frozenRows` and `Sheet.frozenColumns` setters/getters (nullable `int?`). Each sheet in a workbook can carry its own independent freeze combination (rows+cols, rows only, cols only, or none). Values are serialized as OOXML-compliant `<pane>` elements on save and fully restored on decode.
- **Examples**: `example/excel_freeze_panes.dart` (single-sheet samples) and `example/excel_freeze_panes_multisheet.dart` (multi-sheet workbook with per-sheet freezes and round-trip assertions).
- **Flutter example app**: New `Freeze Panes` and `Multi-Sheet Freeze Panes` sidebar sections, with matching helpers and code snippets.
- **Tests**: Re-enabled `test/freeze_panes_test.dart` covering the four freeze configurations (4/4 passing).
- **Documentation**: README "Features" and "Usage" sections now document the freeze-pane API.

### Fixed
- **`<sheetView>` writer**: Only the first sheet in the workbook carries `tabSelected="1"` (Excel requires exactly one active tab). `<selection>` entries and `activePane` are now chosen dynamically based on the actual split configuration, so Excel no longer shows the "Reparaciones / Recover as much as we can" dialog on multi-sheet workbooks.

## [2.1.0] - 2026-06-28

### Added
- **Sheet Protection**: Added full programmatic support for worksheet protection, including password hashing (legacy 16-bit XOR algorithm) and custom protection settings (objects, scenarios, format cells, select locked/unlocked cells, etc.).
- **Cell Style Locking**: Exposed `locked` and `hidden` boolean properties on `CellStyle` to control cell editability/visibility under sheet protection.
- **Worksheet Protection Parsing & Serialization**: Added parser and serializer rules to correctly process and serialize worksheet protection states and cell locking XML elements, fully compatible with Microsoft Excel and Google Sheets.

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

