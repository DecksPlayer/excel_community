# Project Structure & Class Reference

This document provides a comprehensive tree view of the **`excel_community`** project along with detailed descriptions in English of each class, enum, extension, and component across the codebase.

---

## 1. Project Directory Tree

```
excel_community/
├── .github/                               # GitHub Actions workflows and CI/CD pipelines
├── assets/                                # Asset files and media for documentation
├── benchmark/                             # Performance and benchmarking scripts
│   └── benchmark.dart                     # Speed and memory benchmarking tool
├── doc/                                   # Documentation assets, guides, and diagrams
├── example/                               # Standalone Dart CLI sample scripts
│   ├── excel_border.dart                  # Cell border styles & colors example
│   ├── excel_charts.dart                  # Chart generation example (all 9 types)
│   ├── excel_conditional_formatting.dart  # Conditional formatting rules example
│   ├── excel_custom_size.dart             # Row height & column width example
│   ├── excel_example_style.dart           # Font styling, colors & alignments example
│   ├── excel_freeze_panes.dart            # Freeze panes demo for single sheet
│   ├── excel_freeze_panes_multisheet.dart # Multi-sheet freeze panes demo
│   ├── excel_images.dart                  # Embedding images (PNG, SVG, etc.) example
│   ├── excel_save.dart                    # Workbook creation and file saving demo
│   ├── excel_time_consuming.dart          # High-volume stress test and benchmark demo
│   ├── generate_underline_example.dart    # Font underline styles showcase
│   └── main.dart                          # Quickstart example script
├── excel_flutter_example/                 # Interactive Flutter Web/Desktop showcase app
│   ├── lib/
│   │   ├── data/
│   │   │   ├── code_snippets.dart         # Snippet lookup registry for UI view
│   │   │   ├── data_export_samples.dart   # Sample workbook + JSON for the Data Export wiki
│   │   │   ├── data_validation_samples.dart # Data validation wiki rules and sample values
│   │   │   ├── hyperlink_samples.dart     # Hyperlink wiki examples (run, saved and reopened)
│   │   │   ├── number_format_catalog.dart # Number formats + expected Excel output (wiki & demo workbook)
│   │   │   ├── section_details.dart       # Showcase section catalog & descriptions
│   │   │   └── snippets/                  # Code sample snippets for each feature
│   │   │       ├── basic.dart
│   │   │       ├── cell_comments.dart
│   │   │       ├── charts.dart
│   │   │       ├── conditional_formatting.dart
│   │   │       ├── fonts_styles.dart
│   │   │       ├── formulas_display_text.dart
│   │   │       ├── hidden_columns.dart
│   │   │       ├── images.dart
│   │   │       ├── merged_cells.dart
│   │   │       ├── styling.dart
│   │   │       ├── autofilter.dart
│   │   │       ├── tab_color.dart
│   │   │       ├── page_setup.dart
│   │   │       ├── data_export.dart
│   │   │       ├── hyperlinks.dart
│   │   │       └── data_validation.dart
│   │   ├── models/
│   │   │   └── section_detail.dart        # Section detail data model
│   │   ├── services/
│   │   │   ├── excel_generator.dart       # Main Excel generation service orchestrator
│   │   │   └── helpers/                   # Specialized scenario builder helpers
│   │   │       ├── autofilter_helper.dart
│   │   │       ├── cell_comments_helper.dart
│   │   │       ├── chart_helper.dart
│   │   │       ├── conditional_formatting_helper.dart
│   │   │       ├── data_export_helper.dart
│   │   │       ├── data_validation_helper.dart
│   │   │       ├── hyperlinks_helper.dart
│   │   │       ├── formulas_display_text_helper.dart
│   │   │       ├── full_demo_helper.dart
│   │   │       ├── hidden_columns_helper.dart
│   │   │       ├── image_helper.dart
│   │   │       ├── merged_cells_helper.dart
│   │   │       ├── multi_page_charts_helper.dart
│   │   │       ├── protection_helper.dart
│   │   │       ├── sheets_helper.dart
│   │   │       ├── simple_helper.dart
│   │   │       ├── page_setup_helper.dart
│   │   │       ├── styles_helper.dart
│   │   │       └── tab_color_helper.dart
│   │   ├── widgets/                       # Flutter UI presentation widgets
│   │   │   ├── about_view.dart
│   │   │   ├── code_view_card.dart
│   │   │   ├── data_export_view.dart
│   │   │   ├── data_validation_view.dart
│   │   │   ├── fonts_styles_view.dart
│   │   │   ├── header_card.dart
│   │   │   ├── hyperlinks_view.dart
│   │   │   ├── number_formats_view.dart
│   │   │   ├── page_setup_view.dart
│   │   │   ├── preview_card.dart
│   │   │   ├── sidebar.dart
│   │   │   ├── spreadsheet_preview.dart
│   │   │   └── wiki/
│   │   │       └── wiki_components.dart   # Shared wiki layout: header, tabs, search, grid, cards
│   │   └── main.dart                      # Flutter app entry point
│   ├── test/                              # Flutter widget tests
│   └── pubspec.yaml                       # Flutter demo dependencies
├── lib/                                   # Core excel_community library source
│   ├── excel_community.dart               # Library entry point & part definitions
│   └── src/
│       ├── excel.dart                     # Core Excel workbook class & lifecycle
│       ├── chart/                         # Chart domain models
│       │   ├── chart_base.dart            # Chart, ChartSeries, ChartAnchor base models
│       │   └── chart_types.dart           # Concrete chart types (Column, Bar, Line, Pie, etc.)
│       ├── number_format/                 # Number format processing engine
│       │   ├── num_format.dart            # NumFormat base, StandardNumFormat, CustomNumFormat
│       │   ├── format_renderer.dart       # Best-effort numeric/date display-text renderer (NumFormat.format)
│       │   └── formats/numbers/
│       │       ├── datetime_format.dart   # Date and time numeric format handlers
│       │       ├── numeric_format.dart    # Decimal, currency, and scientific format handlers
│       │       ├── standard_formats.dart  # ECMA-376 built-in standard format lookup
│       │       └── time_format.dart       # Time duration format handlers
│       ├── parser/                        # XLSX ZIP & XML parsing subsystem
│       │   ├── parse.dart                 # Parser orchestrator (ZIP bootstrap, relationships)
│       │   ├── styles_parser.dart         # Dedicated styles.xml parser
│       │   └── worksheet_parser.dart      # High-performance SAX worksheet parser
│       ├── pivot_table/                   # Pivot Table definitions
│       │   └── pivot_table.dart           # PivotTable, PivotTableValue, PivotValueFunction
│       ├── save/                          # XLSX serialization & archive generation
│       │   ├── save_file.dart             # Save orchestrator & archive bundler
│       │   ├── self_correct_span.dart     # Merged cell span boundary correction helper
│       │   ├── charts/
│       │   │   └── chart_manager.dart     # Chart XML & drawing relationship generator
│       │   ├── comments/
│       │   │   └── comment_manager.dart   # VML drawing & comment XML serializer
│       │   ├── hyperlinks/
│       │   │   └── hyperlink_manager.dart # External hyperlink relationships (TargetMode=External)
│       │   ├── images/
│       │   │   └── image_manager.dart     # Drawing XML image anchor & media manager
│       │   ├── pivot_table/
│       │   │   └── pivot_table_manager.dart # Pivot definition & cache record serializer
│       │   ├── styles/
│       │   │   ├── style_manager.dart     # styles.xml orchestrator
│       │   │   ├── style_resource_collector.dart # Font, fill, border & DXF deduplicator
│       │   │   └── style_xml_builders.dart # XML element builder utilities
│       │   ├── workbook/
│       │   │   └── workbook_manager.dart  # workbook.xml & sharedStrings.xml generator
│       │   └── worksheet/
│       │       └── worksheet_manager.dart # worksheet XML serializer & row/cell builder
│       ├── sharedStrings/                 # Shared strings table subsystem
│       │   └── shared_strings.dart        # _SharedStringsMaintainer, SharedString, TextSpan
│       ├── sheet/                         # Worksheet domain model & extensions
│       │   ├── border_style.dart          # Border, _BorderSet, BorderStyle
│       │   ├── cell_index.dart            # CellIndex (coordinates & cell IDs)
│       │   ├── cell_style.dart            # CellStyle (fonts, colors, alignments, borders)
│       │   ├── cell_value/                # Typed cell value hierarchy (sealed classes)
│       │   │   ├── bool_cell_value.dart   # BoolCellValue
│       │   │   ├── cell_value.dart        # CellValue (sealed base class)
│       │   │   ├── date_cell_value.dart   # DateCellValue
│       │   │   ├── datetime_cell_value.dart # DateTimeCellValue
│       │   │   ├── double_cell_value.dart # DoubleCellValue
│       │   │   ├── formula_cell_value.dart # FormulaCellValue
│       │   │   ├── int_cell_value.dart    # IntCellValue
│       │   │   ├── text_cell_value.dart   # TextCellValue
│       │   │   └── time_cell_value.dart   # TimeCellValue
│       │   ├── conditional_formatting.dart # Conditional formatting models & DXF styles
│       │   ├── data_model.dart            # Data (individual cell model)
│       │   ├── excel_image.dart           # ExcelImage, ImageAnchor, ExcelImageType
│       │   ├── font_family.dart           # FontFamily constants/enum
│       │   ├── font_style.dart            # _FontStyle internal container
│       │   ├── header_footer.dart         # HeaderFooter configuration & BoolParsing
│       │   ├── sheet.dart                 # Sheet core class
│       │   ├── sheet_charts.dart          # SheetCharts extension (addChart, charts API)
│       │   ├── sheet_data_ext.dart        # SheetDataExt extension (rows, cols, mutations)
│       │   ├── sheet_dimensions.dart      # SheetDimensions extension (sizes, autofit, hiding)
│       │   ├── sheet_images.dart          # SheetImages extension (addImage, images API)
│       │   ├── auto_filter.dart           # AutoFilter, FilterColumn, & CustomFilterRule models
│       │   ├── tab_color.dart             # TabColor OpenXML <tabColor> domain model
│       │   ├── page_setup.dart            # PageSetup, PaperSize, PageMargins & PrintOptions print models
│       │   ├── sheet_export.dart          # SheetExport/ExcelExport: rowsAsMaps, toJson, toCsv, appendRowsFromMaps
│       │   ├── hyperlink.dart             # Hyperlink model, SheetHyperlinks & DataHyperlink APIs
│       │   ├── data_validation.dart       # DataValidation model, SheetDataValidations & DataDataValidation APIs
│       │   ├── sheet_protection.dart      # SheetProtection & password hashing
│       │   └── sheet_spans.dart           # SheetSpans extension (merge, unMerge)
│       ├── utilities/                     # Shared helpers, colors, enums, builders
│       │   ├── archive.dart               # Archive copying and compression utilities
│       │   ├── chart_builders/            # Clean architecture chart style builders
│       │   │   ├── area_chart_builder.dart
│       │   │   ├── chart_style_builder.dart # Base ChartStyleBuilder interface
│       │   │   ├── chart_style_builder_factory.dart # Factory pattern dispatcher
│       │   │   ├── column_bar_chart_builder.dart
│       │   │   ├── line_chart_builder.dart
│       │   │   ├── pie_chart_builder.dart
│       │   │   ├── radar_chart_builder.dart
│       │   │   └── scatter_chart_builder.dart
│       │   ├── chart_color_config.dart    # Centralized color palettes for chart types
│       │   ├── chart_xml_writer.dart      # OpenXML chart & drawing XML orchestrator
│       │   ├── colors/                    # Color constants & hex conversions
│       │   │   ├── accents.dart           # AccentColors palette
│       │   │   ├── base.dart              # BaseColors (black, white, greys)
│       │   │   ├── blue.dart              # Blue & LightBlue palettes
│       │   │   ├── excel_color.dart       # ExcelColor class, ColorType enum, StringExt
│       │   │   ├── green.dart             # Green & LightGreen palettes
│       │   │   ├── others.dart            # Purple, Indigo, Cyan, Teal, Lime, Brown, BlueGrey
│       │   │   ├── red.dart               # Red & Pink palettes
│       │   │   └── yellow_orange.dart     # Yellow, Amber, Orange, DeepOrange palettes
│       │   ├── constants.dart             # OpenXML namespaces, XML templates, MIME types
│       │   ├── enum.dart                  # TextWrapping, VerticalAlign, HorizontalAlign, Underline, FontScheme
│       │   ├── fast_list.dart             # FastList<K> optimized collection
│       │   ├── cell_rect.dart             # _CellRect: range parsing, containment, subtraction & row/column shifts
│       │   ├── span.dart                  # _Span merged cell boundary model
│       │   └── utility.dart               # Coordinate calculations & helper functions
│       ├── web_helper/                    # Platform-specific file downloading
│       │   ├── client_save_excel.dart     # Web/IO stub implementation
│       │   └── web_save_excel_browser.dart # Web browser Blob/Anchor saving implementation
│       └── xls_parser/                    # Legacy BIFF8 (.xls) parser subsystem
│           └── excel_xls.dart             # CfbFile, CfbDirectoryEntry, BiffParser, _BiffSheet, _Record, _SstStream
├── test/                                  # Automated unit and integration test suites
│   ├── excel_border_test.dart
│   ├── excel_charts_test.dart
│   ├── excel_conditional_formatting_test.dart
│   ├── excel_custom_size_test.dart
│   ├── excel_datetime_test.dart
│   ├── excel_drawing_cleanup_test.dart
│   ├── excel_formula_test.dart
│   ├── excel_freeze_panes_multisheet_test.dart
│   ├── excel_freeze_panes_test.dart
│   ├── excel_image_test.dart
│   ├── excel_merged_cell_styles_test.dart
│   ├── excel_pivot_table_test.dart
│   ├── excel_protection_test.dart
│   ├── excel_reading_test.dart
│   ├── excel_rows_and_columns_test.dart
│   ├── excel_save_test.dart
│   ├── excel_test.dart
│   ├── excel_xls_test.dart
│   └── helper.dart
├── analysis_options.yaml                  # Dart analyzer and linter configuration
├── CHANGELOG.md                           # Version history and release notes
├── CHART_COLORS_GUIDE.md                  # Comprehensive guide to chart color schemes
├── CLEAN_CODE_ARCHITECTURE.md             # Architectural documentation for chart builders
├── LICENSE                                # Project license (MIT)
├── pubspec.yaml                           # Package manifest, metadata & dependencies
└── README.md                              # Main package documentation & examples
```

---

## 2. Class Descriptions by Module

### 2.1 Core Workbook & Lifecycle (`lib/src/`)

#### [`lib/src/excel.dart`](lib/src/excel.dart)
- **`Excel`**: The primary entry point for creating, reading, modifying, and saving Microsoft Excel workbooks. Manages worksheets (`Sheet`), global shared strings (`_SharedStringsMaintainer`), cell styles (`CellStyle`), differential styles (`DifferentialStyle`), number formats (`NumFormatMaintainer`), and the underlying OpenXML ZIP archive. Supports decoding from `.xlsx` (OpenXML) and legacy `.xls` (BIFF8 binary format).

---

### 2.2 Sheet & Cell Domain Models (`lib/src/sheet/`)

#### [`lib/src/sheet/sheet.dart`](lib/src/sheet/sheet.dart)
- **`Sheet`**: Represents an individual Excel worksheet. Encapsulates cell data indexed by row and column, dimensions, frozen pane settings, merged cell spans, sheet protection, charts, embedded images, pivot tables, and conditional formatting rules.

#### [`lib/src/sheet/data_model.dart`](lib/src/sheet/data_model.dart)
- **`Data`**: Represents an individual cell within a worksheet. Holds the cell's typed value (`CellValue`), formatting style (`CellStyle`), 0-based row and column coordinates, parent sheet reference, and optional cell comment. Exposes `displayText`, which renders the value as formatted text via `CellStyle.numberFormat.format()`.

#### [`lib/src/sheet/cell_index.dart`](lib/src/sheet/cell_index.dart)
- **`CellIndex`**: Immutable value object representing a 2D coordinate on a worksheet. Supports conversions between 0-based integer indexes (`columnIndex`, `rowIndex`) and standard Excel alphanumeric references (e.g., `"A1"`, `"BC42"`).

#### [`lib/src/sheet/cell_style.dart`](lib/src/sheet/cell_style.dart)
- **`CellStyle`**: Comprehensive styling configuration for cells. Controls font attributes (color, family, size, bold, italic, underline, strikethrough, font scheme), background fill colors, text alignments (horizontal and vertical), text rotation (-90° to 90°), text wrapping, cell borders (left, right, top, bottom, diagonal), cell protection locks, and number format associations.

#### [`lib/src/sheet/font_style.dart`](lib/src/sheet/font_style.dart)
- **`_FontStyle`**: Internal model encapsulating font appearance properties used for deduplicating fonts during XML style generation.

#### [`lib/src/sheet/border_style.dart`](lib/src/sheet/border_style.dart)
- **`Border`**: Configures a single border line with a specific line style (`BorderStyle`) and color (`ExcelColor`).
- **`_BorderSet`**: Internal data structure grouping all five borders of a cell (left, right, top, bottom, diagonal) plus diagonal orientation flags.
- **`BorderStyle`** (`enum`): OpenXML border line types (e.g., `none`, `thin`, `medium`, `thick`, `dashed`, `dotted`, `double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, `slantDashDot`).

#### [`lib/src/sheet/font_family.dart`](lib/src/sheet/font_family.dart)
- **`FontFamily`** (`enum`): Standard font family constants including Arial, Calibri, Comic Sans MS, Courier New, Georgia, Impact, Times New Roman, Trebuchet MS, and Verdana.

#### [`lib/src/sheet/header_footer.dart`](lib/src/sheet/header_footer.dart)
- **`HeaderFooter`**: Defines printing and page layout header and footer configurations (left, center, right sections for normal, odd, even, and first pages).
- **`BoolParsing`** (`extension`): Utility extension for parsing boolean attributes in header/footer XML nodes.

#### [`lib/src/sheet/auto_filter.dart`](lib/src/sheet/auto_filter.dart)
- **`AutoFilter`**: Represents an OpenXML `<autoFilter>` element containing the cell range reference (`ref`), upper-left and lower-right coordinates, column filters, and methods to test cell inclusion (`containsCell`, `containsCellId`).
- **`FilterColumn`**: Represents column filter configuration within an AutoFilter (`<filterColumn>`), including button visibility, matching filter values (`<filters>`), blank filtering, and custom comparison rules (`<customFilters>`).
- **`CustomFilterRule`**: Custom filter criteria mapping an operator (`equal`, `greaterThan`, `lessThan`, etc.) and comparison value.
- **`FilterOperator`** (`enum`): OpenXML comparison operators for custom filter conditions.

#### [`lib/src/sheet/tab_color.dart`](lib/src/sheet/tab_color.dart)
- **`TabColor`**: Represents an OpenXML `<tabColor>` element for worksheet tab coloring. Encapsulates ARGB hex strings, `ExcelColor` references, Office theme indices (`theme`), tint modifiers (`tint`), indexed colors (`indexed`), and auto colors (`auto`). Provides normalization from `#RRGGBB`, `#AARRGGBB`, and shorthand `#RGB`, plus schema-compliant serialization inside `<sheetPr>`.

#### [`lib/src/sheet/page_setup.dart`](lib/src/sheet/page_setup.dart)
- **`PageSetup`**: Represents the OpenXML `<pageSetup>` element (orientation, paper size, scale, fit-to-width/height, first page number, page order, black & white, draft, comments/errors printing, DPI, copies) plus the `fitToPage` flag stored in `<sheetPr><pageSetUpPr>`. Only explicitly set attributes are serialized; an existing printer settings `r:id` is preserved on save.
- **`PaperSize`**: SpreadsheetML paper size code with named constants (`letter`, `legal`, `a3`, `a4`, `a5`, envelopes, ...) and millimetre dimensions; unknown codes round-trip via `PaperSize.fromCode`.
- **`PageMargins`**: `<pageMargins>` in inches with Excel's `normal`, `wide` and `narrow` presets and a `fromCentimeters` factory.
- **`PrintOptions`**: `<printOptions>` flags for gridlines, headings and horizontal/vertical centering.
- **`PageOrientation`**, **`PageOrder`**, **`PrintCellComments`**, **`PrintErrors`** (`enum`s): OpenXML enumerations for the corresponding `<pageSetup>` attributes.

#### [`lib/src/sheet/data_validation.dart`](lib/src/sheet/data_validation.dart)
- **`DataValidation`**: Represents an OpenXML `<dataValidation>`: `type` (`DataValidationType`), `operator` (`DataValidationOperator`), `formula1`/`formula2`, blank/dropdown/message flags, prompt and error texts and `errorStyle` (`DataValidationErrorStyle`). Factories `list`, `listFromRange`, `wholeNumber`, `decimal`, `date`, `time`, `textLength`, `custom` and `inputMessage`; builders `withPrompt`/`withError`; `accepts(CellValue)` evaluates the rule in Dart (`null` when it depends on formulas or ranges).
- **`SheetDataValidations`** (`extension` on `Sheet`): `addDataValidation(range, rule)` (subtracts the area from overlapping rules), `setDataValidation`, `getDataValidation`, `removeDataValidation`, `clearDataValidations`, `dataValidations`, `hasDataValidations`. Rules shift with row/column inserts and removals.
- **`DataDataValidation`** (`extension` on `Data`): `cell.dataValidation` getter/setter.

#### [`lib/src/sheet/hyperlink.dart`](lib/src/sheet/hyperlink.dart)
- **`Hyperlink`**: Represents an OpenXML `<hyperlink>`: external `url` (web, `mailto:`, file) and/or internal `location` (cell, range or defined name), plus `tooltip` and `display`. Factories `Hyperlink.url`, `Hyperlink.email`, `Hyperlink.cell` (quotes sheet names when needed) and `Hyperlink.location`.
- **`SheetHyperlinks`** (`extension` on `Sheet`): `setHyperlink`, `setHyperlinkRange`, `getHyperlink`, `removeHyperlink`, `clearHyperlinks`, `hyperlinks`, `hasHyperlinks`. Links are keyed by cell or range reference and shift with `insertRow`/`removeRow`/`insertColumn`/`removeColumn`.
- **`DataHyperlink`** (`extension` on `Data`): `cell.hyperlink` getter/setter.

#### [`lib/src/sheet/sheet_export.dart`](lib/src/sheet/sheet_export.dart)
- **`SheetExport`** (`extension` on `Sheet`): `rowsAsMaps` (header row as keys; empty headers named by column letter, duplicates suffixed), `rowsAsValues`, `toJson`, `toCsv` (RFC 4180) and the inverse `appendRowsFromMaps`.
- **`ExcelExport`** (`extension` on `Excel`): `toMaps` and `toJson` for every sheet of the workbook.
- **`ExportValueMode`** (`enum`): `typed` (native Dart values: `DateTime` for dates, `Duration` for times, cached values for formulas) or `displayText` (formatted text as shown by Excel).

#### [`lib/src/sheet/sheet_protection.dart`](lib/src/sheet/sheet_protection.dart)
- **`SheetProtection`**: Manages worksheet protection flags (locking cells, formatting, inserting/deleting rows/columns, sorting, filtering) and implements the standard Excel 16-bit password hashing algorithm.

#### [`lib/src/sheet/sheet_spans.dart`](lib/src/sheet/sheet_spans.dart)
- **`SheetSpans`** (`extension`): Extends `Sheet` with methods for merging cell ranges (`merge()`), removing merges (`unMerge()`), and inspecting active spans.

#### [`lib/src/sheet/sheet_dimensions.dart`](lib/src/sheet/sheet_dimensions.dart)
- **`SheetDimensions`** (`extension`): Extends `Sheet` with dimension management: setting explicit column widths, row heights, enabling column auto-fitting, and hiding/unhiding specific rows or columns.

#### [`lib/src/sheet/sheet_data_ext.dart`](lib/src/sheet/sheet_data_ext.dart)
- **`SheetDataExt`** (`extension`): Extends `Sheet` with row and column iteration and structure mutation APIs (e.g., `rows`, `row()`, `insertRowIterables()`, `insertRow()`, `deleteRow()`, `insertColumn()`, `deleteColumn()`, `findAndReplace()`).

#### [`lib/src/sheet/sheet_charts.dart`](lib/src/sheet/sheet_charts.dart)
- **`SheetCharts`** (`extension`): Extends `Sheet` with chart capabilities: `addChart()`, `removeChart()`, `clearCharts()`, and accessing the list of attached charts.

#### [`lib/src/sheet/sheet_images.dart`](lib/src/sheet/sheet_images.dart)
- **`SheetImages`** (`extension`): Extends `Sheet` with image capabilities: `addImage()`, `removeImage()`, `clearImages()`, and accessing embedded images.

---

### 2.3 Typed Cell Values (`lib/src/sheet/cell_value/`)

#### [`lib/src/sheet/cell_value/cell_value.dart`](lib/src/sheet/cell_value/cell_value.dart)
- **`CellValue`** (`sealed class`): Abstract sealed base class for all strongly-typed cell value representations. Enforces a default number format and a `write(NumFormat?)` string serialization method.

#### [`lib/src/sheet/cell_value/text_cell_value.dart`](lib/src/sheet/cell_value/text_cell_value.dart)
- **`TextCellValue`**: Represents plain text or rich formatted text (via `TextSpan` list) stored in a cell.

#### [`lib/src/sheet/cell_value/int_cell_value.dart`](lib/src/sheet/cell_value/int_cell_value.dart)
- **`IntCellValue`**: Represents a 64-bit integer numeric value in a cell.

#### [`lib/src/sheet/cell_value/double_cell_value.dart`](lib/src/sheet/cell_value/double_cell_value.dart)
- **`DoubleCellValue`**: Represents a floating-point double precision numeric value in a cell.

#### [`lib/src/sheet/cell_value/bool_cell_value.dart`](lib/src/sheet/cell_value/bool_cell_value.dart)
- **`BoolCellValue`**: Represents a boolean (`TRUE`/`FALSE`) value in a cell.

#### [`lib/src/sheet/cell_value/date_cell_value.dart`](lib/src/sheet/cell_value/date_cell_value.dart)
- **`DateCellValue`**: Represents a calendar date (year, month, day) formatted as an Excel serial date number.

#### [`lib/src/sheet/cell_value/time_cell_value.dart`](lib/src/sheet/cell_value/time_cell_value.dart)
- **`TimeCellValue`**: Represents a time of day (hours, minutes, seconds, milliseconds) formatted as a fractional day serial value.

#### [`lib/src/sheet/cell_value/datetime_cell_value.dart`](lib/src/sheet/cell_value/datetime_cell_value.dart)
- **`DateTimeCellValue`**: Represents a combined date and time timestamp with convenience accessors for local and UTC `DateTime` instances.

#### [`lib/src/sheet/cell_value/formula_cell_value.dart`](lib/src/sheet/cell_value/formula_cell_value.dart)
- **`FormulaCellValue`**: Represents an Excel calculation formula string (e.g., `"=SUM(A1:A10)"`, `"=VLOOKUP(...)"`). Carries an optional `cachedValue` — the pre-calculated `<v>` result read from the source file, if it had one (not recalculated or rewritten by this library).

---

### 2.4 Embedded Images (`lib/src/sheet/excel_image.dart`)

#### [`lib/src/sheet/excel_image.dart`](lib/src/sheet/excel_image.dart)
- **`ExcelImage`**: Model representing an image embedded inside a worksheet. Holds raw image binary bytes, image format (`ExcelImageType`), and sheet position/dimension anchor (`ImageAnchor`). Supports automatic type inference from file extensions.
- **`ImageAnchor`**: Defines the placement and dimensions of an embedded image using English Metric Units (EMUs) or pixels (96 DPI conversion: 1 pixel = 9,525 EMUs).
- **`ExcelImageType`** (`enum`): Supported image formats: PNG, JPEG, GIF, BMP, TIFF, WMF, EMF, SVG, WebP, ICO.

---

### 2.5 Conditional Formatting (`lib/src/sheet/conditional_formatting.dart`)

#### [`lib/src/sheet/conditional_formatting.dart`](lib/src/sheet/conditional_formatting.dart)
- **`ConditionalFormattingGroup`**: Represents a collection of formatting rules applied over a specific cell range (`sqref`, e.g., `"A1:D50"`).
- **`ConditionalFormattingRule`**: Individual conditional formatting rule definition specifying the rule type (`ConditionalFormattingType`), comparison operator (`ConditionalFormattingOperator`), target formula/value, priority, and applied differential style (`DifferentialStyle`).
- **`DifferentialStyle`**: Represents OpenXML `<dxf>` styles containing differential formatting overrides (font color/weight, cell background pattern fill, custom border highlights).
- **`ConditionalFormattingType`** (`enum`): OpenXML rule types: `cellIs`, `expression`, `colorScale`, `dataBar`, `iconSet`, `top10`, `uniqueValues`, `duplicateValues`, `containsText`, `notContainsText`, `beginsWith`, `endsWith`, `containsBlanks`, `notContainsBlanks`, `containsErrors`, `notContainsErrors`, `timePeriod`.
- **`ConditionalFormattingOperator`** (`enum`): Comparison operators: `lessThan`, `lessThanOrEqual`, `equal`, `notEqual`, `greaterThanOrEqual`, `greaterThan`, `between`, `notBetween`, `containsText`, `notContains`, `beginsWith`, `endsWith`.

---

### 2.6 Chart Domain Models (`lib/src/chart/`)

#### [`lib/src/chart/chart_base.dart`](lib/src/chart/chart_base.dart)
- **`Chart`** (`abstract class`): Abstract base class for all Excel charts. Contains chart title, position anchor (`ChartAnchor`), data series collection (`ChartSeries`), 3D view settings, and legend configuration.
- **`ChartSeries`**: Represents a single data series in a chart, defining category cell ranges (X-axis labels), value cell ranges (Y-axis values), optional series name, and custom series styling.
- **`ChartAnchor`**: Defines the bounding box and position of a chart on the worksheet grid using column/row start/end offsets and pixel dimensions.

#### [`lib/src/chart/chart_types.dart`](lib/src/chart/chart_types.dart)
- **`ColumnChart`**: Represents vertical clustered, stacked, or 100% stacked column charts.
- **`BarChart`**: Represents horizontal bar charts.
- **`LineChart`**: Represents line charts with customizable line thickness, marker styles, and smooth curves.
- **`PieChart`**: Represents 2D pie charts with slice colors and custom first-slice angles.
- **`DoughnutChart`**: Represents doughnut charts with configurable center hole size percentage.
- **`AreaChart`**: Represents area charts with semi-transparent filled series.
- **`ScatterChart`**: Represents XY scatter plots with point markers and trendlines.
- **`RadarChart`**: Represents radar/spider charts supporting both wireframe lines and filled polygon styles.

---

### 2.7 Clean Architecture Chart Builders (`lib/src/utilities/chart_builders/`)

#### [`lib/src/utilities/chart_builders/chart_style_builder.dart`](lib/src/utilities/chart_builders/chart_style_builder.dart)
- **`ChartStyleBuilder`** (`abstract class`): Base builder interface adhering to SOLID principles. Defines contracts for building chart-level XML properties (`buildProperties`) and series-level visual styling (`buildSeriesStyle`).

#### [`lib/src/utilities/chart_builders/chart_style_builder_factory.dart`](lib/src/utilities/chart_builders/chart_style_builder_factory.dart)
- **`ChartStyleBuilderFactory`**: Factory class that returns the concrete `ChartStyleBuilder` implementation corresponding to a given `Chart` runtime type.

#### Concrete Chart Builders:
- **`ColumnBarChartBuilder`**: Builds XML styling for column and bar charts with solid fills and borders.
- **`LineChartBuilder`**: Builds XML styling for line charts with stroke widths and circular markers.
- **`AreaChartBuilder`**: Builds XML styling for area charts with alpha transparency fills (50% fill opacity, 90% line opacity).
- **`PieChartBuilder`**: Builds XML styling for pie and doughnut charts, handling random color assignment per slice and hole sizes.
- **`ScatterChartBuilder`**: Builds XML styling for scatter charts with distinct marker shapes and point colors.
- **`RadarChartBuilder`**: Builds XML styling for radar charts with filled polygons or wireframe lines.

#### [`lib/src/utilities/chart_color_config.dart`](lib/src/utilities/chart_color_config.dart)
- **`ChartColorConfig`**: Centralized configuration providing curated color palettes (series palette, radar palette, pie slice palette) and styling constants for all chart builders.

#### [`lib/src/utilities/chart_xml_writer.dart`](lib/src/utilities/chart_xml_writer.dart)
- **`ChartXmlWriter`**: Orchestrator that generates OpenXML `drawing*.xml` and `chart*.xml` files, delegating chart-specific styling to builder classes.

---

### 2.8 Number Format Engine (`lib/src/number_format/`)

#### [`lib/src/number_format/num_format.dart`](lib/src/number_format/num_format.dart)
- **`NumFormat`** (`sealed class`): Base class for number formats. Provides factory methods for standard ECMA-376 formats (e.g., `standard_0`, `standard_14`, currency, dates, percentages) and custom format strings. Declares `format(CellValue?)`, implemented per format family to render a cell's display text.

#### [`lib/src/number_format/format_renderer.dart`](lib/src/number_format/format_renderer.dart)
- Best-effort renderer (`renderNumericValue`, `renderDateTimeValue`) that converts a raw `CellValue` into the display text a spreadsheet application would show for a given ECMA-376 §18.8.30 format code — thousands separators, fixed decimals, percentages, scientific notation, currency/date/time literals. Backs `NumFormat.format()` and `Data.displayText`. Not a full implementation of the Excel formatting language (exotic locale codes and precise fraction rendering fall back to a simplified output).
- **`StandardNumFormat`** (`sealed class`): Represents a built-in OpenXML standard format with a fixed `numFmtId` (0–163).
- **`CustomNumFormat`** (`sealed class`): Represents a user-defined custom format string with a dynamic format ID (≥ 164).
- **`NumFormatMaintainer`**: Tracks, allocates, and maps standard and custom number format IDs to avoid duplication in `styles.xml`.

#### [`lib/src/number_format/formats/numbers/numeric_format.dart`](lib/src/number_format/formats/numbers/numeric_format.dart)
- **`NumericNumFormat`**: Base class for decimal, integer, currency, and scientific number formatting.
- **`StandardNumericNumFormat`**: Standard numeric format implementation.
- **`CustomNumericNumFormat`**: Custom numeric format string implementation.

#### [`lib/src/number_format/formats/numbers/datetime_format.dart`](lib/src/number_format/formats/numbers/datetime_format.dart)
- **`DateTimeNumFormat`**: Base class for date and timestamp formats.
- **`StandardDateTimeNumFormat`**: Standard ECMA date/time format implementation.
- **`CustomDateTimeNumFormat`**: Custom date/time format string implementation.

#### [`lib/src/number_format/formats/numbers/time_format.dart`](lib/src/number_format/formats/numbers/time_format.dart)
- **`TimeNumFormat`**: Base class for time duration and clock formats.
- **`StandardTimeNumFormat`**: Standard time format implementation.
- **`CustomTimeNumFormat`**: Custom time format string implementation.

#### [`lib/src/number_format/formats/numbers/standard_formats.dart`](lib/src/number_format/formats/numbers/standard_formats.dart)
- **`StandardFormats`**: Static dictionary of all 50+ built-in ECMA-376 standard number formats (§18.8.30).

---

### 2.9 XLSX Parser Subsystem (`lib/src/parser/`)

#### [`lib/src/parser/parse.dart`](lib/src/parser/parse.dart)
- **`Parser`**: Orchestrates the complete decoding of an `.xlsx` archive: unzipping files, reading `[Content_Types].xml`, resolving package and sheet relationships (`.rels`), loading shared strings, and coordinating styles and worksheet parsers.

#### [`lib/src/parser/styles_parser.dart`](lib/src/parser/styles_parser.dart)
- **`_StylesParser`**: Dedicated parser for `xl/styles.xml`. Decodes font pools (`<fonts>`), fill pools (`<fills>`), border pools (`<borders>`), custom number formats (`<numFmts>`), cell style records (`<cellXfs>`), and differential styles (`<dxfs>`).

#### [`lib/src/parser/worksheet_parser.dart`](lib/src/parser/worksheet_parser.dart)
- **`_WorksheetParser`**: High-performance streaming SAX/event-based parser for worksheet XML. Decodes row and cell structures, resolves cell coordinates and shared string indices, parses formulas, dimensions, and merged cells with minimal memory allocation.

---

### 2.10 XLSX Save & Serialization Subsystem (`lib/src/save/`)

#### [`lib/src/save/save_file.dart`](lib/src/save/save_file.dart)
- **`Save`**: Coordinates the entire workbook serialization process. Invokes manager classes to build XML parts, packages media and drawing assets, updates `[Content_Types].xml`, and encodes the final ZIP archive byte stream. Shared helpers load parts from the original file on demand (`_loadXmlPart`, `_relationshipsPart`), assign unused relationship ids (`_nextRelationshipId`) and continue part numbering after the original file (`_highestPartIndex`).

#### Specialized Save Managers:
- **`_ChartManager`** (`lib/src/save/charts/chart_manager.dart`): Serializes chart objects into OpenXML chart parts (`xl/charts/chart*.xml`), creates drawing parts (`xl/drawings/drawing*.xml`), and writes relationship files.
- **`_ImageManager`** (`lib/src/save/images/image_manager.dart`): Embeds image binaries into `xl/media/`, generates drawing XML image anchors, and merges image anchors into existing chart drawings when both exist on the same worksheet.
- **`_PivotTableManager`** (`lib/src/save/pivot_table/pivot_table_manager.dart`): Serializes Pivot Table definitions, Pivot Cache definitions, and Pivot Cache records into their respective XML parts.
- **`_CommentManager`** (`lib/src/save/comments/comment_manager.dart`): Generates cell comment XML parts (`xl/comments*.xml`), legacy VML drawing shapes (`xl/drawings/vmlDrawing*.vml`), and relationship links.
- **`_HyperlinkManager`** (`lib/src/save/hyperlinks/hyperlink_manager.dart`): Rebuilds each worksheet's external hyperlink relationships (`TargetMode="External"`) from the `Hyperlink` model, keeping every other relationship of the sheet.
- **`_StyleManager`** (`lib/src/save/styles/style_manager.dart`): Coordinates generation of the unified `xl/styles.xml` file.
- **`_StyleResourceCollector`** (`lib/src/save/styles/style_resource_collector.dart`): Traverses all cells and conditional formatting rules to collect, deduplicate, and index all fonts, fills, borders, number formats, and differential styles.
- **`_StyleResources`** (`lib/src/save/styles/style_resource_collector.dart`): Data container holding deduplicated style resource collections and index maps.
- **`_StyleXmlBuilders`** (`lib/src/save/styles/style_xml_builders.dart`): Utility providing static builder methods to construct XML strings for individual fonts, fills, borders, and number formats.
- **`_WorkbookManager`** (`lib/src/save/workbook/workbook_manager.dart`): Updates `xl/workbook.xml` (sheet catalog, active sheet, views) and generates `xl/sharedStrings.xml`.
- **`_WorksheetManager`** (`lib/src/save/worksheet/worksheet_manager.dart`): Generates complete worksheet XML files (`xl/worksheets/sheet*.xml`), writing rows, cells, formulas, dimensions, column widths, row heights, frozen panes, merged cells, and protection tags.
- **`_MergedCellEntry`** (`lib/src/save/worksheet/worksheet_manager.dart`): Represents a cell entry (data or style-only) for non-origin merged cells when generating worksheet XML.

---

### 2.11 Pivot Table Subsystem (`lib/src/pivot_table/`)

#### [`lib/src/pivot_table/pivot_table.dart`](lib/src/pivot_table/pivot_table.dart)
- **`PivotTable`**: Data model representing programmatic Pivot Table configuration, specifying source data ranges, destination cell anchors, row fields, column fields, filter fields, and aggregated value fields.
- **`PivotTableValue`**: Represents a single aggregated data field in a Pivot Table, binding a source column to an aggregation function.
- **`PivotValueFunction`** (`enum`): Supported aggregation functions: `sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevp`, `variance`, `varp`.

---

### 2.12 Shared Strings Subsystem (`lib/src/sharedStrings/`)

#### [`lib/src/sharedStrings/shared_strings.dart`](lib/src/sharedStrings/shared_strings.dart)
- **`_SharedStringsMaintainer`**: Manages the workbook-wide Shared String Table (SST), providing O(1) string deduplication, integer index allocation, and XML serialization.
- **`SharedString`**: Model representing an entry in the shared string table, supporting plain string text or rich formatted text composed of multiple `TextSpan` objects.
- **`TextSpan`**: Represents a styled segment of text within a rich-text string, containing specific font formatting overrides.

---

### 2.13 Legacy BIFF8 (.xls) Parser Subsystem (`lib/src/xls_parser/`)

#### [`lib/src/xls_parser/excel_xls.dart`](lib/src/xls_parser/excel_xls.dart)
- **`CfbFile`**: Compound File Binary (CFB / OLE2) container parser for reading structured storage in legacy Excel 97–2003 `.xls` files.
- **`CfbDirectoryEntry`**: Represents a stream or storage directory entry inside a CFB container (e.g., the `Workbook` stream).
- **`BiffParser`**: Binary Interchange File Format (BIFF8) parser that decodes raw record streams into worksheets, cell values, formulas, styles, and dimensions.
- **`_BiffSheet`**: Internal model holding sheet metadata and stream offsets parsed from the BIFF `BOUNDSHEET` record.
- **`_Record`**: Represents a single BIFF record containing record ID, byte length, and binary payload.
- **`_SstStream`**: Parser for the BIFF8 Shared String Table stream.

---

### 2.14 Utilities, Color Palettes & Enums (`lib/src/utilities/`)

#### [`lib/src/utilities/colors/excel_color.dart`](lib/src/utilities/colors/excel_color.dart)
- **`ExcelColor`**: Robust color model with hex conversion, integer conversion, and 100+ predefined Material and standard color constants. Supports 6-character hex output (`RRGGBB`) for OpenXML chart compatibility.
- **`ColorType`** (`enum`): Categorizes color instances (`color`, `material`, `materialAccent`).
- **`StringExt`** (`extension`): Extension on `String` enabling direct conversion of hex strings to `ExcelColor` instances.

#### Color Palette Classes:
- **`AccentColors`** (`lib/src/utilities/colors/accents.dart`): Material accent colors.
- **`BaseColors`** (`lib/src/utilities/colors/base.dart`): Base colors (black, white, greys).
- **`BlueColors`** (`lib/src/utilities/colors/blue.dart`): Blue and light blue color ranges.
- **`GreenColors`** (`lib/src/utilities/colors/green.dart`): Green and light green color ranges.
- **`RedColors`** (`lib/src/utilities/colors/red.dart`): Red and pink color ranges.
- **`YellowOrangeColors`** (`lib/src/utilities/colors/yellow_orange.dart`): Yellow, amber, orange, and deep orange color ranges.
- **`OtherColors`** (`lib/src/utilities/colors/others.dart`): Purple, deep purple, indigo, cyan, teal, lime, brown, and blue-grey color ranges.

#### Enums & Collections:
- **`TextWrapping`** (`lib/src/utilities/enum.dart`): Text wrapping modes (`WrapText`, `Clip`).
- **`VerticalAlign`** (`lib/src/utilities/enum.dart`): Vertical cell alignment options (`Top`, `Center`, `Bottom`, `Justify`, `Distributed`).
- **`HorizontalAlign`** (`lib/src/utilities/enum.dart`): Horizontal cell alignment options (`Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterContinuous`, `Distributed`).
- **`Underline`** (`lib/src/utilities/enum.dart`): Font underline styles (`None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting`).
- **`FontScheme`** (`lib/src/utilities/enum.dart`): OpenXML font schemes (`Unset`, `Major`, `Minor`).
- **`FastList<K>`** (`lib/src/utilities/fast_list.dart`): High-performance collection pairing list ordering with hash map indexing for O(1) membership checks.
- **`_Span`** (`lib/src/utilities/span.dart`): Represents rectangular boundary coordinates for merged cell regions.

---

### 2.15 Cross-Platform Web Saving (`lib/src/web_helper/`)

#### [`lib/src/web_helper/client_save_excel.dart`](lib/src/web_helper/client_save_excel.dart) & [`lib/src/web_helper/web_save_excel_browser.dart`](lib/src/web_helper/web_save_excel_browser.dart)
- **`SavingHelper`**: Conditional compilation abstraction for saving generated Excel files in browser web environments via `dart:html` Blob URLs and simulated anchor clicks.

---

### 2.16 Flutter Showcase Application (`excel_flutter_example/`)

- **`ExcelDemoApp` / `ExcelCommunityDemoApp`** (`main.dart`): Root widget initializing the interactive Flutter demo application.
- **`SectionDetail`** (`models/section_detail.dart`): Model representing demo feature sections with title, description, icons, and code snippets.
- **`ExcelGeneratorService`** (`services/excel_generator.dart`): Main service coordinating workbook generation and file downloading across demo scenarios.
- **Demo Helper Services** (`services/helpers/`): Modular services generating sample workbooks for specific Excel capabilities:
  - `CellCommentsHelper`: Cell comments with author and text styling.
  - `ChartHelper`: Full suite of 9 chart types.
  - `ConditionalFormattingHelper`: Highlight rules, data bars, and gradient styling.
  - `FullDemoHelper`: Comprehensive multi-sheet showcase workbook.
  - `HiddenColumnsHelper`: Hidden row and column configurations.
  - `ImageHelper`: Embedding PNG and SVG images with custom anchors.
  - `MergedCellsHelper`: Merged ranges with preserved cell borders and styles.
  - `MultiPageChartsHelper`: Charts organized across multiple worksheet tabs.
  - `ProtectionHelper`: Sheet protection and password locking demonstrations.
  - `SheetsHelper`: Multi-sheet operations (renaming, reordering, RTL layout).
  - `SimpleHelper`: Quickstart basic table creation.
  - `StylesHelper`: Typography, colors, patterns, and alignment showcases.
  - `FormulasDisplayTextHelper`: Reads a bundled fixture with real cached `<v>` formula results and demonstrates `FormulaCellValue.cachedValue` and `Data.displayText` across currency, percentage, date, custom, and time formats.
- **UI Widgets** (`widgets/`):
  - `SpreadsheetPreview`: Interactive grid previewing generated Excel sheets.
  - `Sidebar`: Navigation drawer for selecting demo categories.
  - `HeaderCard`: Top banner displaying section titles and export buttons.
  - `CodeViewCard`: Syntax-highlighted Dart code snippet viewer.
  - `PreviewCard`: Card container for table previews.
  - `AboutView`: Library features, architecture, and documentation summary view.
  - `FontsStylesView`: Interactive font and typography preview widget.
  - `wiki/wiki_components.dart`: Shared building blocks for the interactive wiki sections (`WikiPage`, `WikiTab`, `WikiSearchField`, `WikiGrid`, `WikiCard`, `WikiPreviewBox`). Each `WikiCard` has a live preview and its own copyable code snippet.
  - `DataExportView`: Data export wiki: runs `rowsAsMaps`, `rowsAsValues`, `toJson`, `toCsv` and `appendRowsFromMaps` on a sample workbook and shows the library's live output next to each snippet (Export and Import tabs).
  - `DataValidationView`: Data validation wiki (Dropdown Lists, Numbers/Dates/Times, Text/Custom/Checks): each card renders the cell with its dropdown and input message and checks sample values with `DataValidation.accepts()`.
  - `HyperlinksView`: Hyperlinks wiki (External, Inside the Workbook, Read/Style/Remove tabs): each card applies a link, shows the cell as Excel renders it and the links read back from the saved file.
  - `NumberFormatsView`: Number formats wiki: every built-in `NumFormat` (IDs 23–26 reserved) and common custom codes, grouped by category with cross-category search. Previews are rendered by `NumFormat.format()`; data comes from `data/number_format_catalog.dart`, which also drives the demo workbook.
  - `PageSetupView`: Interactive page setup & print wiki: every `PaperSize` drawn to scale with a portrait/landscape toggle, plus orientation & scaling, margin presets and print options, each with a live page preview and copyable code.

---

## 3. Summary Reference Table

| Category | Class / Type | File Path | Primary Responsibility |
|---|---|---|---|
| **Core** | `Excel` | [`lib/src/excel.dart`](lib/src/excel.dart) | Main workbook manager for creating, loading, and saving Excel files. |
| **Sheet** | `Sheet` | [`lib/src/sheet/sheet.dart`](lib/src/sheet/sheet.dart) | Represents an Excel worksheet containing cells, dimensions, charts, images, etc. |
| **Sheet** | `Data` | [`lib/src/sheet/data_model.dart`](lib/src/sheet/data_model.dart) | Individual cell holding value, style, coordinates, and comment. |
| **Sheet** | `CellIndex` | [`lib/src/sheet/cell_index.dart`](lib/src/sheet/cell_index.dart) | 2D cell coordinate supporting "A1" string syntax and 0-based indexes. |
| **Sheet** | `CellStyle` | [`lib/src/sheet/cell_style.dart`](lib/src/sheet/cell_style.dart) | Cell styling definition (fonts, fills, borders, alignments, number format). |
| **Sheet** | `Border` | [`lib/src/sheet/border_style.dart`](lib/src/sheet/border_style.dart) | Individual border line configuration (style and color). |
| **Sheet** | `SheetProtection` | [`lib/src/sheet/sheet_protection.dart`](lib/src/sheet/sheet_protection.dart) | Worksheet protection settings and 16-bit password hashing. |
| **Sheet** | `AutoFilter` | [`lib/src/sheet/auto_filter.dart`](lib/src/sheet/auto_filter.dart) | Worksheet AutoFilter range and column filter definitions. |
| **Sheet** | `FilterColumn` | [`lib/src/sheet/auto_filter.dart`](lib/src/sheet/auto_filter.dart) | Filter criteria and button configuration per column. |
| **Values** | `CellValue` | [`lib/src/sheet/cell_value/cell_value.dart`](lib/src/sheet/cell_value/cell_value.dart) | Sealed base class for typed cell values (`Text`, `Int`, `Double`, `Date`, etc.). |
| **Images** | `ExcelImage` | [`lib/src/sheet/excel_image.dart`](lib/src/sheet/excel_image.dart) | Image embedding model with binary payload and anchor positioning. |
| **Images** | `ImageAnchor` | [`lib/src/sheet/excel_image.dart`](lib/src/sheet/excel_image.dart) | Image position and size anchor using EMUs or pixels. |
| **Formatting** | `ConditionalFormattingRule` | [`lib/src/sheet/conditional_formatting.dart`](lib/src/sheet/conditional_formatting.dart) | Rule definition for OpenXML conditional formatting with DXF styles. |
| **Formatting** | `DifferentialStyle` | [`lib/src/sheet/conditional_formatting.dart`](lib/src/sheet/conditional_formatting.dart) | Differential style override applied by conditional formatting rules. |
| **Charts** | `Chart` | [`lib/src/chart/chart_base.dart`](lib/src/chart/chart_base.dart) | Abstract base class for all 9 chart types. |
| **Charts** | `ChartSeries` | [`lib/src/chart/chart_base.dart`](lib/src/chart/chart_base.dart) | Data series binding categories and value ranges. |
| **Charts** | `ChartStyleBuilder` | [`lib/src/utilities/chart_builders/chart_style_builder.dart`](lib/src/utilities/chart_builders/chart_style_builder.dart) | Base interface for SOLID clean architecture chart style builders. |
| **Charts** | `ChartStyleBuilderFactory` | [`lib/src/utilities/chart_builders/chart_style_builder_factory.dart`](lib/src/utilities/chart_builders/chart_style_builder_factory.dart) | Factory pattern resolver creating specific chart builders. |
| **Charts** | `ChartXmlWriter` | [`lib/src/utilities/chart_xml_writer.dart`](lib/src/utilities/chart_xml_writer.dart) | High-level OpenXML chart and drawing XML generator. |
| **Charts** | `ChartColorConfig` | [`lib/src/utilities/chart_color_config.dart`](lib/src/utilities/chart_color_config.dart) | Centralized color palette manager for chart series, pie, and radar charts. |
| **Numbers** | `NumFormat` | [`lib/src/number_format/num_format.dart`](lib/src/number_format/num_format.dart) | Sealed base class for built-in standard and custom number formats. |
| **Numbers** | `NumFormatMaintainer` | [`lib/src/number_format/num_format.dart`](lib/src/number_format/num_format.dart) | Deduplicator and allocator for number format IDs in `styles.xml`. |
| **Parser** | `Parser` | [`lib/src/parser/parse.dart`](lib/src/parser/parse.dart) | Main XLSX archive parsing coordinator. |
| **Parser** | `_StylesParser` | [`lib/src/parser/styles_parser.dart`](lib/src/parser/styles_parser.dart) | Dedicated parser for `styles.xml` fonts, fills, borders, and XFs. |
| **Parser** | `_WorksheetParser` | [`lib/src/parser/worksheet_parser.dart`](lib/src/parser/worksheet_parser.dart) | High-performance streaming SAX worksheet parser. |
| **Save** | `Save` | [`lib/src/save/save_file.dart`](lib/src/save/save_file.dart) | Main XLSX save orchestrator and ZIP encoder. |
| **Save** | `_StyleManager` | [`lib/src/save/styles/style_manager.dart`](lib/src/save/styles/style_manager.dart) | Serializes complete `styles.xml` file. |
| **Save** | `_WorksheetManager` | [`lib/src/save/worksheet/worksheet_manager.dart`](lib/src/save/worksheet/worksheet_manager.dart) | Serializes rows, cells, spans, and dimensions into worksheet XML. |
| **Pivot** | `PivotTable` | [`lib/src/pivot_table/pivot_table.dart`](lib/src/pivot_table/pivot_table.dart) | Configuration model for programmatic Pivot Table generation. |
| **Strings** | `_SharedStringsMaintainer` | [`lib/src/sharedStrings/shared_strings.dart`](lib/src/sharedStrings/shared_strings.dart) | O(1) deduplication and indexing for the Shared String Table. |
| **Colors** | `ExcelColor` | [`lib/src/utilities/colors/excel_color.dart`](lib/src/utilities/colors/excel_color.dart) | Color model supporting hex, ARGB, RGB, and predefined palettes. |
| **Legacy XLS** | `BiffParser` | [`lib/src/xls_parser/excel_xls.dart`](lib/src/xls_parser/excel_xls.dart) | Binary stream parser for legacy Excel 97–2003 (.xls) files. |
