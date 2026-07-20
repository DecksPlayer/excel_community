![Logo](https://raw.githubusercontent.com/Decksplayer/excel_community/main/assets/logo.png)

# Excel Community

> [!NOTE]
> This is a community-maintained fork of the [excel](https://github.com/justkawal/excel) library.

<a href='https://flutter.io'>  
  <img src='https://img.shields.io/badge/Platform-Flutter-yellow.svg'  
    alt='Platform' />  
</a>
<a href='https://pub.dev/packages/excel_community'>  
  <img src='https://img.shields.io/pub/v/excel_community.svg'  
    alt='Pub Package' />
</a>
<!-- Official Documentation -->
<a href='https://excelcommunity.decksplayer.com/'>  
  <img src='https://img.shields.io/badge/Documentation-Official-blue.svg'  
    alt='Official Documentation' />
</a>
<a href='https://opensource.org/licenses/MIT'>  
  <img src='https://img.shields.io/badge/License-MIT-red.svg'  
    alt='License: MIT' />
</a>
<a href='https://github.com/decksplayer/excel_community/issues'>  
  <img src='https://img.shields.io/github/issues/decksplayer/excel_community'  
    alt='Issues' />  
</a>
<a href='https://github.com/decksplayer/excel_community/network'>  
  <img src='https://img.shields.io/github/forks/decksplayer/excel_community'  
    alt='Forks' />  
</a>
<a href='https://github.com/decksplayer/excel_community/stargazers'>  
  <img src='https://img.shields.io/github/stars/decksplayer/excel_community'  
    alt='Stars' />  
</a>


### [Excel Community](https://www.pub.dev/packages/excel_community) is a community-maintained fork of the [excel](https://github.com/justkawal/excel) library. It provides a robust and high-performance way to read, create, and update XLSX files, as well as parse legacy XLS files.

## Features

- ✅ **Read & Write XLSX**: Full support for reading and writing Excel files in modern `.xlsx` format
- ✅ **Read Legacy XLS**: Read-only support for old `.xls` (Excel 97-2003) workbooks using a custom, zero-dependency parser
- ✅ **Multiple Data Types**: Text, Numbers, Formulas, Dates, Times, Booleans
- ✅ **Cell Styling**: Fonts (Bold, Italic, Underline, Strikethrough), Colors, Borders, Alignment, Number Formats
- ✅ **Conditional Formatting**: Dynamic cell highlighting (numeric ranges, text match, custom formulas, duplicate values) with custom fills & fonts
- ✅ **Charts**: Column, Bar, Line, Area, Pie, Doughnut, Scatter, and Radar charts
- ✅ **Images**: Embed PNG, JPEG, BMP, GIF, TIFF, WMF, EMF, SVG, WebP and ICO images
- ✅ **Pivot Tables**: Create dynamic pivot tables programmatically with row/column fields and various aggregation functions (Sum, Count, Average, Max, Min, etc.)
- ✅ **Cell Comments**: Attach rich descriptions or review notes to specific cells, displaying red triangle markers in Excel (read & write)
- ✅ **Cell Operations**: Merge cells, insert/delete rows and columns
- ✅ **Sheet Management**: Create, copy, rename, delete sheets
- ✅ **Freeze Panes**: Lock rows and/or columns so headers and key columns stay visible while scrolling (single sheet or multi-sheet workbooks)
- ✅ **Cross-platform**: Works on Flutter Web, Android, iOS, Desktop

## Road-map:
 - ➕ Formulas and Calculations
 - 💾 Support Multiple Data type efficiently
 - ✅ Charts (Implemented!)
 - ✅ Images (Implemented!)
 - 📰 Create Tables and style
 - 🔐 Encrypt and Decrypt excel on the go.
 - Many more **features**

<!-- Breaking changes moved to the end -->


## 📖 Documentation & Guides

For more details on how to use `excel_community`, see the following detailed guides:

- 🚀 **[Getting Started Guide](doc/getting_started.md)**: Installation, reading/writing files, and saving across different platforms.
- 🎨 **[Cell Styling & Operations Guide](doc/styling_and_operations.md)**: How to update cells, set custom fonts, borders, alignments, and number formats.
- 📊 **[Working with Charts Guide](doc/charts.md)**: Inserting column, bar, line, area, pie, scatter, and radar charts.
- 🌄 **[Working with Images Guide](doc/images.md)**: Embedding PNG, JPEG, SVG, WebP, and other image formats.
- 🎨 **[Chart Color Strategy Guide](doc/chart_colors.md)**: Detailed description of colors, opacities, and accessibility options used in charts.
- 🏗️ **[Architecture & Design Guide](doc/architecture.md)**: Overview of the clean code architecture implemented in the chart components.
- 🌐 **[Official Documentation Site](https://excelcommunity.decksplayer.com/)**: Browse the official online docs for `excel_community`.

## If you find this tool useful, please drop a ⭐️

## Performance & Benchmarks
<details open>

`excel_community` is highly optimized for large-scale operations. Below is a cold-start scaling comparison (measuring Build + Encode time on a fresh Dart VM) and an isolated active-process benchmark (1,000,000 cells) against the original `excel` package (v4.0.6) and `excel_plus` (v2.7.2).

### 1. Cold-Start Scaling Benchmark (Build + Encode)

| Workload | Library | Build | Encode | Total | File Size | Speedup vs Original | Speedup vs Plus | Speedup vs Community |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **5,000,000 cells** <br>*(500k rows × 10 cols)* | **excel_community** | **5.76 s** | **28.71 s** | **34.48 s** | 34.7 MB | **>3.48x** | **1.33x** | **1.00x** |
| | excel_plus | 14.42 s | 31.30 s | 45.72 s | 34.7 MB | >2.62x | 1.00x | 0.75x |
| | excel_original | Timeout | Timeout | Timeout (>2m) | — | 1.00x | — | — |
| | | | | | | | | |
| **100,000 cells** <br>*(10k rows × 10 cols)* | **excel_community** | **426 ms** | 757 ms | **1,183 ms** | 661.6 KB | **2.70x** | **1.06x** | **1.00x** |
| | excel_plus | 516 ms | **741 ms** | 1,257 ms | 660.9 KB | 2.54x | 1.00x | 0.94x |
| | excel_original | 548 ms | 2,643 ms | 3,191 ms | 695.2 KB | 1.00x | 0.39x | 0.37x |
| | | | | | | | | |
| **10,000 cells** <br>*(1k rows × 10 cols)* | **excel_community** | 349 ms | 284 ms | 633 ms | 61.8 KB | **1.98x** | 0.86x | **1.00x** |
| | excel_plus | **323 ms** | **221 ms** | **544 ms** | 61.7 KB | 2.30x | **1.00x** | 1.16x |
| | excel_original | 452 ms | 800 ms | 1,252 ms | 70.5 KB | 1.00x | 0.43x | 0.51x |

### 2. Isolated Benchmark (1,000,000 Cells)
*20,000 rows × 50 columns workload in an active process (measuring all phases and Peak RSS memory):*

| Library | Create | Encode | Decode | Total Time | Peak RSS | File Size | Speedup vs Original | Speedup vs Plus | Speedup vs Community |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **excel_community** (Ours) | **959 ms** | **2,167 ms** | 24,263 ms | 27,388 ms | 1,695 MB | 7.08 MB | **1.96x** | **0.65x** | **1.00x** |
| excel_plus | 2,350 ms | 5,762 ms | **9,657 ms** | **17,768 ms** | **726 MB** | 7.08 MB | 3.02x | 1.00x | 1.54x |
| excel_original (v4.0.6) | 1,907 ms | 23,404 ms | 28,340 ms | 53,650 ms | 2,552 MB | 7.04 MB | 1.00x | 0.33x | 0.51x |

> [!TIP]
> **Eager vs. Lazy Parsing**: `excel_community` eagerly parses sheets on load to guarantee direct $O(1)$ cell updates and stable identities, whereas `excel_plus` loads cells lazily. While lazy loading makes the initial decode faster, `excel_community` delivers unmatched performance for workloads requiring heavy read/write cell manipulations and ultra-fast encoding.
</details>

<details open><summary><h2>📖 Usage</h2></summary>

<details open>
<summary><h3>📄 Read XLSX / XLS File</h3></summary>

> [!NOTE]
> The same `Excel.decodeBytes(bytes)` and `Excel.decodeBuffer(stream)` methods automatically detect the format using magic signature bytes and parse both modern `.xlsx` and legacy `.xls` (Excel 97-2003) files transparently.

```dart
var file = 'Path_to_pre_existing_Excel_File/excel_file.xlsx';
var bytes = File(file).readAsBytesSync();
var excel = Excel.decodeBytes(bytes);
for (var table in excel.tables.keys) {
  print(table); //sheet Name
  print(excel.tables[table].maxColumns);
  print(excel.tables[table].maxRows);
  for (var row in excel.tables[table].rows) {
    for (var cell in row) {
      print('cell ${cell.rowIndex}/${cell.columnIndex}');
      final value = cell.value;
      final numFormat = cell.cellStyle?.numberFormat ?? NumFormat.standard_0;
      switch(value){
        case null:
          print('  empty cell');
          print('  format: ${numFormat}');
        case TextCellValue():
          print('  text: ${value.value}');
        case FormulaCellValue():
          print('  formula: ${value.formula}');
          print('  format: ${numFormat}');
        case IntCellValue():
          print('  int: ${value.value}');
          print('  format: ${numFormat}');
        case BoolCellValue():
          print('  bool: ${value.value ? 'YES!!' : 'NO..' }');
          print('  format: ${numFormat}');
        case DoubleCellValue():
          print('  double: ${value.value}');
          print('  format: ${numFormat}');
        case DateCellValue():
          print('  date: ${value.year} ${value.month} ${value.day} (${value.asDateTimeLocal()})');
        case TimeCellValue():
          print('  time: ${value.hour} ${value.minute} ... (${value.asDuration()})');
        case DateTimeCellValue():
          print('  date with time: ${value.year} ${value.month} ${value.day} ${value.hour} ... (${value.asDateTimeLocal()})');
      }

      print('$row');
    }
  }
}
```

</details>

<details>
<summary><h3>🌐 Read XLSX / XLS in Flutter Web</h3></summary>

Use `FilePicker` to pick files in Flutter Web. [FilePicker](https://pub.dev/packages/file_picker.git)

```dart
/// Use FilePicker to pick files in Flutter Web

FilePickerResult pickedFile = await FilePicker.platform.pickFiles(
  type: FileType.custom,
  allowedExtensions: ['xlsx'],
  allowMultiple: false,
);

/// file might be picked

if (pickedFile != null) {
  var bytes = pickedFile.files.single.bytes;
  var excel = Excel.decodeBytes(bytes);
  for (var table in excel.tables.keys) {
    print(table); //sheet Name
    print(excel.tables[table].maxColumns);
    print(excel.tables[table].maxRows);
    for (var row in excel.tables[table].rows) {
      print('$row');
    }
  }
}
```

</details>

<details>
<summary><h3>📂 Read XLSX / XLS from Flutter's Asset Folder</h3></summary>

```dart
import 'package:flutter/services.dart' show ByteData, rootBundle;

/* Your ......other important..... code here */

ByteData data = await rootBundle.load('assets/existing_excel_file.xlsx');
var bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
var excel = Excel.decodeBytes(bytes);

for (var table in excel.tables.keys) {
  print(table); //sheet Name
  print(excel.tables[table].maxColumns);
  print(excel.tables[table].maxRows);
  for (var row in excel.tables[table].rows) {
    print('$row');
  }
}
```

</details>

<details open>
<summary><h3>➕ Create New XLSX File</h3></summary>

```dart
// automatically creates 1 empty sheet: Sheet1
var excel = Excel.createExcel();
```

</details>

<details open>
<summary><h3>✏️ Update Cell Values</h3></summary>

```dart
/*
 * sheetObject.updateCell(cell, value, { CellStyle (Optional)});
 * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
 * cell can be identified with Cell Address or by 2D array having row and column Index;
 * Cell Style options are optional
 */

Sheet sheetObject = excel['SheetName'];

CellStyle cellStyle = CellStyle(backgroundColorHex: '#1AFF1A', fontFamily :getFontFamily(FontFamily.Calibri));

cellStyle.underline = Underline.Single; // or Underline.Double
cellStyle.isStrikethrough = true; // enable strikethrough text


var cell = sheetObject.cell(CellIndex.indexByString('A1'));
cell.value = null; // removing any value
cell.value = TextCellValue('Some Text');
cell.value = IntCellValue(8);
cell.value = BoolCellValue(true);
cell.value = DoubleCellValue(13.37);
cell.value = DateCellValue(year: 2023, month: 4, day: 20);
cell.value = TimeCellValue(hour: 20, minute: 15, second: 5, millisecond: ...);
cell.value = DateTimeCellValue(year: 2023, month: 4, day: 20, hour: 15, ...);
cell.cellStyle = cellStyle;

// setting the number style
cell.cellStyle = (cell.cellStyle ?? CellStyle()).copyWith(

  /// for IntCellValue, DoubleCellValue and BoolCellValue use; 
  numberFormat: CustomNumericNumFormat('#,##0.00 \\m\\²'),

  /// for DateCellValue and DateTimeCellValue use:
  numberFormat: CustomDateTimeNumFormat('m/d/yy h:mm'),

  /// for TimeCellValue use:
  numberFormat: CustomDateTimeNumFormat('mm:ss'),

  /// a builtin format for dates
  numberFormat: NumFormat.standard_14,
  
  /// a builtin format that uses a red text color for negative numbers
  numberFormat: NumFormat.standard_38,

  /// currency formats (v1.2.0+)
  numberFormat: NumFormat.standard_5,  // $#,##0_);($#,##0)
  numberFormat: NumFormat.standard_6,  // $#,##0_);[Red]($#,##0)
  numberFormat: NumFormat.standard_7,  // $#,##0.00_);($#,##0.00)
  numberFormat: NumFormat.standard_8,  // $#,##0.00_);[Red]($#,##0.00)

  /// accounting with fill character (v1.2.0+)
  numberFormat: NumFormat.standard_41, // _(* #,##0_);...
  numberFormat: NumFormat.standard_44, // _($* #,##0.00_);...

  // The numberFormat changes automatially if you set a CellValue that 
  // does not work with the numberFormat set previously. So in case you
  // want to set a new value, e.g. from a date to a decimal number, 
  // make sure you set the new value first and then your custom
  // numberFormat).
);


// printing cell-type
print('CellType: ' + switch(cell.value) {
  null => 'empty cell',
  TextCellValue() => 'text',
  FormulaCellValue() => 'formula',
  IntCellValue() => 'int',
  BoolCellValue() => 'bool',
  DoubleCellValue() => 'double',
  DateCellValue() => 'date',
  TimeCellValue() => 'time',
  DateTimeCellValue() => 'date with time',
});

///
/// Inserting and removing column and rows

// insert column at index = 8
sheetObject.insertColumn(8);

// remove column at index = 18
sheetObject.removeColumn(18);

// insert row at index = 82
sheetObject.insertRow(82);

// remove row at index = 80
sheetObject.removeRow(80);
```

</details>

<details>
<summary><h3>🎨 Cell-Style Options</h3></summary>

| key                | description                                                                                                                             |
|--------------------|-----------------------------------------------------------------------------------------------------------------------------------------|
| fontFamily         | eg. getFontFamily(`FontFamily.Arial`) or getFontFamily(`FontFamily.Comic_Sans_MS`) `There is total 182 Font Families available for now` |
| fontSize           | specify the font-size as integer eg. fontSize = 15                                                                                      |
| bold               | makes text bold - when set to `true`, by-default it is set to `false`                                                                   |
| italic             | makes text italic - when set to `true`, by-default it is set to `false`                                                                 |
| underline          | Gives underline to text `enum Underline { None, Single, Double }` eg. Underline.Single, by-default it is set to Underline.None          |
| strikethrough      | makes text strikethrough (struck through) - when set to `true`, by-default it is set to `false`                                         |
| fontColorHex       | Font Color eg. '#0000FF'                                                                                                                |
| rotation (degree)  | rotation of text eg. 50, rotation varies from `-90 to 90`, with including `90` and `-90`                                                |
| backgroundColorHex | Background color of cell eg. '#faf487'                                                                                                  |
| wrap               | Text wrapping `enum TextWrapping { WrapText, Clip }` eg. TextWrapping.Clip                                                              |
| verticalAlign      | align text vertically `enum VerticalAlign { Top, Center, Bottom }` eg. VerticalAlign.Top                                                |
| horizontalAlign    | align text horizontally `enum HorizontalAlign { Left, Center, Right }` eg. HorizontalAlign.Right                                        |
| leftBorder         | the left border of the cell (see below)                                                                                                 |
| rightBorder        | the right border of the cell                                                                                                            |
| topBorder          | the top border of the cell                                                                                                              |
| bottomBorder       | the bottom border of the cell                                                                                                           |
| diagonalBorder     | the diagonal "border" of the cell                                                                                                       |
| diagonalBorderUp   | boolean value indicating if the diagonal "border" should be displayed on the up diagonal                                                |
| diagonalBorderDown | boolean value indicating if the diagonal "border" should be displayed on the down diagonal                                              |
| locked             | boolean value indicating if the cell should be locked (read-only) when sheet protection is active, default is `true`                     |
| hidden             | boolean value indicating if the cell formulas should be hidden when sheet protection is active, default is `false`                        |
| numberFormat       | a subtype of ```NumFormat``` to style the CellValue displayed, use default formats such as ```NumFormat.standard_34``` or create your own using ```CustomNumericNumFormat('#,##0.00 \\m\\²')``` ```CustomDateTimeNumFormat('m/d/yy h:mm')```  ```CustomTimeNumFormat('mm:ss')``` |

</details>

<details>
<summary><h3>🎨 Conditional Formatting</h3></summary>

Automatically highlight cells dynamically based on numeric ranges, text matching, or custom formulas (with support for multiple worksheets):

```dart
final excel = Excel.createExcel();

// Sheet 1: Semester 1 (Numeric Range & Text Match Rules)
excel.rename('Sheet1', 'Semester 1');
final sheet1 = excel['Semester 1'];

// Rule 1: Pass (> 70) -> Solid Green Background (#2E7D32) & White Text (#FFFFFF)
final passRule = ConditionalFormattingRule.cellIs(
  operator: ConditionalFormattingOperator.greaterThan,
  formulae: ['70'],
  backgroundColor: ExcelColor.fromHexString('2E7D32'),
  fontColor: ExcelColor.fromHexString('FFFFFF'),
  bold: true,
);

// Rule 2: Fail (<= 50) -> Solid Red Background (#D32F2F) & White Text (#FFFFFF)
final failRule = ConditionalFormattingRule.cellIs(
  operator: ConditionalFormattingOperator.lessThanOrEqual,
  formulae: ['50'],
  backgroundColor: ExcelColor.fromHexString('D32F2F'),
  fontColor: ExcelColor.fromHexString('FFFFFF'),
  bold: true,
);

sheet1.addConditionalFormatting('B2:B10', [passRule, failRule]);

// Rule 3: Text Rule -> Highlight cells containing 'Passed' with Blue Background (#1976D2)
final passedRule = ConditionalFormattingRule.containsText(
  text: 'Passed',
  backgroundColor: ExcelColor.fromHexString('1976D2'),
  fontColor: ExcelColor.fromHexString('FFFFFF'),
  bold: true,
);
sheet1.addConditionalFormattingRule('C2:C10', passedRule);

// Sheet 2: Semester 2 (Duplicate Values & Custom Expression Rules)
final sheet2 = excel['Semester 2'];

// Rule 1: Highlight duplicate subject names in Orange (#E65100)
final duplicateRule = ConditionalFormattingRule.duplicateValues(
  backgroundColor: ExcelColor.fromHexString('E65100'),
  fontColor: ExcelColor.fromHexString('FFFFFF'),
  bold: true,
);
sheet2.addConditionalFormattingRule('A2:A10', duplicateRule);

// Rule 2: Custom Expression formula (Score > 90) -> Gold Background (#FBC02D)
final topScoreRule = ConditionalFormattingRule.expression(
  formula: 'B2>90',
  backgroundColor: ExcelColor.fromHexString('FBC02D'),
  fontColor: ExcelColor.fromHexString('000000'),
  bold: true,
);
sheet2.addConditionalFormattingRule('B2:B10', topScoreRule);
```

</details>

<details>
<summary><h3>🔲 Borders</h3></summary>
Borders are defined for each side (left, right, top, and bottom) of the cell. Both diagonals (up and down) share the
same settings. A boolean value `true` must be set to either `diagonalBorderUp` or `diagonalBorderDown` (or both) to
display the desired diagonal.

Each border must be a `Border` object. This object accepts two parameters : `borderStyle` to select one of the different
supported styles and `borderColorHex` to change the border color.

The `borderStyle` must be a value from the enumeration`BorderStyle`:

- `BorderStyle.None`
- `BorderStyle.DashDot`
- `BorderStyle.DashDotDot`
- `BorderStyle.Dashed`
- `BorderStyle.Dotted`
- `BorderStyle.Double`
- `BorderStyle.Hair`
- `BorderStyle.Medium`
- `BorderStyle.MediumDashDot`
- `BorderStyle.MediumDashDotDot`
- `BorderStyle.MediumDashed`
- `BorderStyle.SlantDashDot`
- `BorderStyle.Thick`
- `BorderStyle.Thin`

```dart
/*
 *
 * Defines thin borders on the left and right of the cell, red thin border on the top
 * and blue medium border on the bottom.
 *
 */

CellStyle cellStyle = CellStyle(
  leftBorder: Border(borderStyle: BorderStyle.Thin),
  rightBorder: Border(borderStyle: BorderStyle.Thin),
  topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: 'FFFF0000'),
  bottomBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: 'FF0000FF'),
);
```

</details>

<details>
<summary><h3>🔒 Sheet Protection & Cell Locking</h3></summary>

Excel sheet protection lets you lock down specific cells to make them read-only, while keeping other cells editable for final users.

* **Sheet-level Protection**: Activate protection on a sheet using a raw string password (which is automatically legacy 16-bit hashed).
* **Cell-level Locking/Hiding**: In Excel, all cells are locked by default under a protected sheet. To allow users to edit specific cells, you must set `locked: false` on their `CellStyle`.

```dart
var excel = Excel.createExcel();
var sheet = excel['Protected Report'];

// 1. Protect the sheet programmatically with a password
sheet.protect('password');

// 2. Configure CellStyles for locked (default) and unlocked (editable) cells
final headerStyle = CellStyle(
  bold: true,
  locked: true, // locked cells are read-only under protection
);

final editableStyle = CellStyle(
  locked: false, // unlocked cells remain editable by users
);

// 3. Apply styles to cells
sheet.updateCell(
  CellIndex.indexByString("A1"),
  TextCellValue("Locked Header"),
  cellStyle: headerStyle,
);

sheet.updateCell(
  CellIndex.indexByString("B1"),
  IntCellValue(100),
  cellStyle: editableStyle, // Users can edit this cell in Excel!
);
```

You can also customize advanced protection flags on `sheet.sheetProtection`:
```dart
sheet.sheetProtection
  ..sheet = true
  ..objects = true
  ..scenarios = true
  ..formatCells = false // allows formatting cells even when protected
  ..selectLockedCells = true;
```

</details>

<details>
<summary><h3>🧊 Freeze Panes</h3></summary>

Freeze Panes lets you keep one or more rows (typically a header) and/or one or more columns (typically a key/identifier column) **always visible** while the user scrolls through the rest of the sheet. `excel_community` exposes this through two simple, nullable properties on `Sheet`:

- `sheet.frozenRows` — number of rows frozen at the **top**.
- `sheet.frozenColumns` — number of columns frozen on the **left**.

Assign `null` (or `0`) to either property to disable that part of the freeze. The values are serialized to OOXML on save and fully restored on decode, so a freeze written by Excel and one written by `excel_community` are interchangeable.

#### Basic usage

```dart
var excel = Excel.createExcel();
var sheet = excel['Sales Report'];

// Lock the header row (row 1) and the product column (column A).
sheet.frozenRows = 1;
sheet.frozenColumns = 1;

// You can also freeze only rows, only columns, or nothing:
sheet.frozenRows = 2;          // freeze the first 2 rows
sheet.frozenColumns = null;    // no column freeze

// sheet.frozenRows = null;     // remove the freeze entirely
// sheet.frozenColumns = 3;     // freeze only the first 3 columns

excel.save(fileName: 'sales_report.xlsx');
```

Under the hood, the writer emits a `<sheetView>` block that is strictly OOXML-compliant — including a `tabSelected="1"` on a single active sheet and `<selection>` entries that match the active panes (so Excel does not show the "recover as much as we can" dialog):

```xml
<sheetViews>
  <sheetView tabSelected="1" workbookViewId="0">
    <pane xSplit="1" ySplit="1" topLeftCell="B2"
          activePane="bottomRight" state="frozen"/>
    <selection pane="topRight"    activeCell="B2" sqref="B2"/>
    <selection pane="bottomLeft"  activeCell="B2" sqref="B2"/>
    <selection pane="bottomRight" activeCell="B2" sqref="B2"/>
  </sheetView>
</sheetViews>
```

For split-only modes the writer automatically picks the right `activePane`:

| `frozenRows` | `frozenColumns` | `activePane` emitted | Selections emitted |
| :--- | :--- | :--- | :--- |
| `> 0` | `> 0` | `bottomRight` | `topRight`, `bottomLeft`, `bottomRight` |
| `> 0` | `null` / `0` | `bottomLeft` | `bottomLeft` |
| `null` / `0` | `> 0` | `topRight` | `topRight` |
| `null` / `0` | `null` / `0` | _(no `<pane>` emitted)_ | _(self-closing `<sheetView>`)_ |

#### Multi-sheet workbooks (different freezes per sheet)

A common requirement is to apply a **different freeze configuration to each sheet** of the same workbook. Just assign the properties per sheet:

```dart
var excel = Excel.createExcel();
excel.delete('Sheet1');

final sales     = excel['Sales'];
final inventory = excel['Inventory'];
final customers = excel['Customers'];
final logs      = excel['Logs'];

// Populate each sheet with its data...
// (omitted here for brevity)

// Then assign a different freeze to each one.
sales.frozenRows     = 1; sales.frozenColumns     = 1;     // rows + cols
inventory.frozenRows = 2; inventory.frozenColumns = null; // rows only
customers.frozenRows = null; customers.frozenColumns = 2;  // cols only
logs.frozenRows = null; logs.frozenColumns = null;         // no freeze

excel.save(fileName: 'multi_freeze_panes.xlsx');

// Round-trip check:
final reopened = Excel.decodeBytes(excel.encode()!);
assert(reopened['Sales'].frozenRows == 1);
assert(reopened['Inventory'].frozenColumns == null);
assert(reopened['Customers'].frozenColumns == 2);
assert(reopened['Logs'].frozenRows == null);
```

Run the full examples to see this in action:

- `example/excel_freeze_panes.dart` — basic freeze (rows+cols, rows only, cols only, none).
- `example/excel_freeze_panes_multisheet.dart` — multi-sheet workbook with one freeze per sheet and a round-trip assertion.

The Flutter example app (`excel_flutter_example/`) exposes the same behavior under the **Freeze Panes** and **Multi-Sheet Freeze Panes** sidebar entries.

</details>

<details>
<summary><h3>🙈 Hidden Columns & Rows</h3></summary>

You can hide specific columns and rows in a worksheet to keep formula inputs clean, protect internal identifiers, or mask confidential details (like salaries or private allowances). `excel_community` provides simple methods to toggle and query visibility:

- `sheet.setColumnHidden(int columnIndex, bool hidden)` — hide or show a column by index (0-based).
- `sheet.isColumnHidden(int columnIndex)` — check if a column is hidden.
- `sheet.setRowHidden(int rowIndex, bool hidden)` — hide or show a row by index (0-based).
- `sheet.isRowHidden(int rowIndex)` — check if a row is hidden.

Hidden settings are correctly saved to OOXML and fully preserved during decoding (round-trip verified).

```dart
var excel = Excel.createExcel();
var sheet = excel['Payroll'];

sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Name"));
sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Salary (Confidential)"));
sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Role"));

// Hide Column B (index 1)
sheet.setColumnHidden(1, true);
assert(sheet.isColumnHidden(1) == true);

// Hide Row 3 (index 2)
sheet.setRowHidden(2, true);
assert(sheet.isRowHidden(2) == true);

excel.save(fileName: 'hidden_columns_demo.xlsx');
```

</details>

<details>
<summary><h3>💬 Cell Comments</h3></summary>

Add descriptive text notes or review comments to any cell in a worksheet. These comments are compatible with Microsoft Excel, Google Sheets, and other modern spreadsheet readers, appearing as a red indicator triangle in the corner of the cell.

* **Set a comment**: Set the comment text using `cell.comment = 'Your comment'`.
* **Read a comment**: Get the comment text using `cell.comment`.

```dart
var excel = Excel.createExcel();
var sheet = excel['Sheet1'];

var cell = sheet.cell(CellIndex.indexByString("B2"));
cell.value = TextCellValue("Widget A");

// Set a comment on cell B2
cell.comment = "This product has a 10% discount this month.";

// Read comment back
String? text = cell.comment;
print("Cell B2 comment: $text");

excel.save(fileName: 'cell_comments.xlsx');
```

</details>

### Make sheet RTL

```dart
/*
 * set rtl to true for making sheet to right-to-left
 * default value of rtl = false ( which means the fresh or default sheet is ltr )
 *
 */

var sheetObject = excel['SheetName'];
sheetObject.rtl = true;
```

### Copy sheet contents to another sheet

```dart
/*
 * excel.copy(String 'existingSheetName', String 'anotherSheetName');
 * existingSheetName should exist in excel.tables.keys in order to successfully copy
 * if anotherSheetName does not exist then it will be automatically created.
 *
 */

excel.copy('existingSheetName', 'anotherSheetName');
```

### Rename sheet

```dart
/*
 * excel.rename(String 'existingSheetName', String 'newSheetName');
 * existingSheetName should exist in excel.tables.keys in order to successfully rename
 *
 */

excel.rename('existingSheetName', 'newSheetName');
```

### Delete sheet

```dart
/*
 * excel.delete(String 'existingSheetName');
 * (existingSheetName should exist in excel.tables.keys) and (excel.tables.keys.length >= 2), in order to successfully delete.
 *
 */

excel.delete('existingSheetName');
```

### Link sheet

```dart
/*
 * excel.link(String 'sheetName', Sheet sheetObject);
 *
 * Any operations performed on (object of 'sheetName') or sheetObject then the operation is performed on both.
 * if 'sheetName' does not exist then it will be automatically created and linked with the sheetObject's operation.
 *
 */

excel.link('sheetName', sheetObject);
```

### Un-Link sheet

```dart
/*
 * excel.unLink(String 'sheetName');
 * In order to successfully unLink the 'sheetName' then it must exist in excel.tables.keys
 *
 */

excel.unLink('sheetName');

// After calling the above function be sure to re-make a new reference of this.

Sheet unlinked_sheetObject = excel['sheetName'];
```

### Merge Cells

```dart
/*
 * sheetObject.merge(CellIndex starting_cell, CellIndex ending_cell, TextCellValue('customValue'));
 * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
 * starting_cell and ending_cell can be identified with Cell Address or by 2D array having row and column Index;
 * customValue is optional
 */

sheetObject.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('E4'), customValue: TextCellValue('Put this text after merge'));
```

### Get Merged Cells List

```dart
// Check which cells are merged

sheetObject.spannedItems.forEach((cells) {
  print('Merged:' + cells.toString());
});
```

### Un-Merge Cells

```dart
/*
 * sheetObject.unMerge(cell);
 * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
 * cell should be identified with string only with an example as 'A1:E4'.
 * to check if 'A1:E4' is un-merged or not
 * call the method excel.getMergedCells(sheet); and verify that it is not present in it.
 */

sheetObject.unMerge('A1:E4');
```

### Find and Replace

```dart
/*
 * int replacedCount = sheetObject.findAndReplace(source, target);
 * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
 * source is the string or ( User's Custom Pattern Matching RegExp )
 * target is the string which is put in cells in place of source
 *
 * it returns the number of replacements made
 */

int replacedCount = sheetObject.findAndReplace('Flutter', 'Google');
```

### Insert Row Iterables

```dart
/*
 * sheetObject.insertRowIterables(list-iterables, rowIndex, iterable-options?);
 * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
 * list-iterables === list of iterables which has to be put in specific row
 * rowIndex === the row in which the iterables has to be put
 * Iterable options are optional
 */

/// It will put the list-iterables in the 8th index row
List<CellValue> dataList = [TextCellValue('Google'), TextCellValue('loves'), TextCellValue('Flutter'), TextCellValue('and'), TextCellValue('Flutter'), TextCellValue('loves'), TextCellValue('Excel')];

sheetObject.insertRowIterables(dataList, 8);
```

### Iterable Options

| key                  | description                                                                                                                       |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| startingColumn       | starting column index from which list-iterables should be started                                                                 |
| overwriteMergedCells | overwriteMergedCells is by-defalut set to `true`, when set to `false` it will stop over-write and will write only in unique cells |

### Append Row

```dart
/*
 * sheetObject.appendRow(list-iterables);
 * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
 * list-iterables === list of iterables
 */

sheetObject.appendRow([TextCellValue('Flutter'), TextCellValue('till'), TextCellValue('Eternity')]);
```

### Get Default Opening Sheet

```dart
/*
 * method which returns the name of the default sheet
 * excel.getDefaultSheet();
 */

var defaultSheet = excel.getDefaultSheet();
print('Default Sheet:' + defaultSheet.toString());
```

### Set Default Opening Sheet

```dart
/*
 * method which sets the name of the default sheet
 * returns bool if successful then true else false
 * excel.setDefaultSheet(sheet);
 * sheet = 'SheetName'
 */

var isSet = excel.setDefaultSheet(sheet);
if (isSet) {
  print('$sheet is set to default sheet.');
} else {
  print('Unable to set $sheet to default sheet.');
}
```

</details>

<details open>
<summary><h2>🌄 Images</h2></summary>

### Embed an Image

Use `sheet.addImage(ExcelImage(...))` to embed an image into a worksheet.
The image floats over the cells at the position defined by its `ImageAnchor`.

```dart
import 'dart:io';
import 'package:excel_community/excel_community.dart';

final sheet = excel['Sheet1'];

// Option A — explicit bytes + type
final pngBytes = File('assets/logo.png').readAsBytesSync();
sheet.addImage(ExcelImage(
  imageBytes: pngBytes,
  imageType: ExcelImageType.png,
  anchor: ImageAnchor.fromPixels(
    column: 0, row: 2,        // top-left corner: column A, row 3
    widthPixels: 200,
    heightPixels: 80,
  ),
));
```

### Auto-detect type from file

`ExcelImage.fromFile` reads the bytes and infers `ExcelImageType` from the
file extension automatically:

```dart
sheet.addImage(ExcelImage.fromFile(
  File('assets/logo.png'),
  anchor: ImageAnchor.fromPixels(
    column: 0, row: 2, widthPixels: 200, heightPixels: 80,
  ),
));

// SVG — visible in Excel 2016+ / Microsoft 365
sheet.addImage(ExcelImage.fromFile(
  File('assets/logo.svg'),
  anchor: ImageAnchor.fromPixels(
    column: 3, row: 2, widthPixels: 200, heightPixels: 80,
  ),
));
```

### Infer type from extension string

```dart
final bytes = File('assets/photo.jpg').readAsBytesSync();
sheet.addImage(ExcelImage(
  imageBytes: bytes,
  imageType: ExcelImageType.fromExtension('photo.jpg'), // → ExcelImageType.jpeg
  anchor: ImageAnchor.fromPixels(
    column: 0, row: 5, widthPixels: 300, heightPixels: 200,
  ),
));
```

### ImageAnchor — pixels vs EMUs

```dart
// fromPixels — assumes 96 DPI (1 px = 9 525 EMUs)
ImageAnchor.fromPixels(
  column: 1, row: 3,
  widthPixels: 200, heightPixels: 100,
  colOffsetPixels: 5,   // optional sub-column offset
  rowOffsetPixels: 5,   // optional sub-row offset
)

// Direct EMU values — full precision
// 1 cm ≈ 360 000 EMUs
ImageAnchor(
  fromColumn: 1, fromRow: 3,
  colOffset: 47_625,    // 5 px
  rowOffset: 47_625,    // 5 px
  widthEmu:  1_905_000, // ≈ 5.29 cm
  heightEmu:   952_500, // ≈ 2.65 cm
)
```

### Supported formats

| `ExcelImageType` | Extensions | Notes |
|---|---|---|
| `png` | `.png` | |
| `jpeg` | `.jpg` `.jpeg` `.jfif` | |
| `gif` | `.gif` | |
| `bmp` | `.bmp` `.dib` | |
| `tiff` | `.tif` `.tiff` | |
| `wmf` | `.wmf` | |
| `emf` | `.emf` | |
| `svg` | `.svg` `.svgz` | Excel 2016+ / 365 |
| `webp` | `.webp` | Excel 365 |
| `ico` | `.ico` | |

</details>

<details open>
<summary><h2>📊 Charts</h2></summary>

Excel Community supports creating various types of charts in your Excel files. Each chart type comes with automatic color schemes for professional-looking visualizations.

### Available Chart Types

- 📈 **ColumnChart** - Vertical bar charts
- 📉 **BarChart** - Horizontal bar charts  
- 📊 **LineChart** - Line charts with markers
- 🏞️ **AreaChart** - Area charts with fill
- 🍰 **PieChart** - Pie charts
- 🍩 **DoughnutChart** - Doughnut charts (pie with center hole)
- ✨ **ScatterChart** - Scatter/XY charts
- 🕸️ **RadarChart** - Radar/spider charts (with filled option)

<details open>
<summary><h3>📝 Basic Chart Example</h3></summary>

```dart
// Create an Excel file and add data
var excel = Excel.createExcel();
var sheet = excel['Sheet1'];

// Add headers
sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Month"));
sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Sales"));
sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Expenses"));

// Add data
final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
final sales = [45, 55, 42, 60, 58, 65];
final expenses = [35, 48, 52, 45, 62, 55];

for (var i = 0; i < months.length; i++) {
  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
    TextCellValue(months[i]),
  );
  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
    IntCellValue(sales[i]),
  );
  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1),
    IntCellValue(expenses[i]),
  );
}

// Create a chart
var chart = ColumnChart(
  title: "Monthly Sales vs Expenses",
  series: [
    ChartSeries(
      name: "Sales",
      categoriesRange: r"Sheet1!$A$2:$A$7",
      valuesRange: r"Sheet1!$B$2:$B$7",
    ),
    ChartSeries(
      name: "Expenses",
      categoriesRange: r"Sheet1!$A$2:$A$7",
      valuesRange: r"Sheet1!$C$2:$C$7",
    ),
  ],
  anchor: ChartAnchor.at(column: 4, row: 1, width: 12, height: 16),
  showLegend: true,
);

// Add chart to sheet
sheet.addChart(chart);

// Save the file
var bytes = excel.save();
```

</details>

<details>
<summary><h3>📈 Column Chart</h3></summary>

Vertical bar chart ideal for comparing values across categories.

```dart
var chart = ColumnChart(
  title: "Sales by Quarter",
  series: [
    ChartSeries(
      name: "Q1",
      categoriesRange: r"Sheet1!$A$2:$A$5",
      valuesRange: r"Sheet1!$B$2:$B$5",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  showLegend: true,
);

sheet.addChart(chart);
```

</details>

<details>
<summary><h3>📉 Bar Chart</h3></summary>

Horizontal bar chart, better for longer category labels.

```dart
var chart = BarChart(
  title: "Revenue by Product",
  series: [
    ChartSeries(
      name: "Revenue",
      categoriesRange: r"Sheet1!$A$2:$A$10",
      valuesRange: r"Sheet1!$B$2:$B$10",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1),
);

sheet.addChart(chart);
```

</details>

<details>
<summary><h3>📊 Line Chart</h3></summary>

Line chart with markers, ideal for showing trends over time.

```dart
var chart = LineChart(
  title: "Website Traffic Trends",
  series: [
    ChartSeries(
      name: "Visitors",
      categoriesRange: r"Sheet1!$A$2:$A$13",
      valuesRange: r"Sheet1!$B$2:$B$13",
    ),
    ChartSeries(
      name: "Page Views",
      categoriesRange: r"Sheet1!$A$2:$A$13",
      valuesRange: r"Sheet1!$C$2:$C$13",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1, width: 12, height: 15),
);

sheet.addChart(chart);
```

</details>

<details>
<summary><h3>🏞️ Area Chart</h3></summary>

Area chart with colored fill, useful for showing cumulative values.

```dart
var chart = AreaChart(
  title: "Cumulative Growth",
  series: [
    ChartSeries(
      name: "Product A",
      categoriesRange: r"Sheet1!$A$2:$A$7",
      valuesRange: r"Sheet1!$B$2:$B$7",
    ),
    ChartSeries(
      name: "Product B",
      categoriesRange: r"Sheet1!$A$2:$A$7",
      valuesRange: r"Sheet1!$C$2:$C$7",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1),
);

sheet.addChart(chart);
```

</details>

<details>
<summary><h3>🍰 Pie Chart</h3></summary>

Pie chart for showing proportions. Uses only one series.

```dart
var chart = PieChart(
  title: "Market Share",
  series: [
    ChartSeries(
      name: "Share",
      categoriesRange: r"Sheet1!$A$2:$A$6",
      valuesRange: r"Sheet1!$B$2:$B$6",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  showLegend: true,
);

sheet.addChart(chart);
```

</details>

<details>
<summary><h3>🍩 Doughnut Chart</h3></summary>

Similar to pie chart but with a center hole.

```dart
var chart = DoughnutChart(
  title: "Department Budget",
  series: [
    ChartSeries(
      name: "Budget",
      categoriesRange: r"Sheet1!$A$2:$A$8",
      valuesRange: r"Sheet1!$B$2:$B$8",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1),
);

sheet.addChart(chart);
```

</details>

<details>
<summary><h3>✨ Scatter Chart (XY)</h3></summary>

Scatter charts (XY) are ideal for showing the relationship between two numeric variables. For a proper representation, ensure your data consists of numbers in both the X and Y columns.

```dart
var chart = ScatterChart(
  title: "Height vs Weight Correlation",
  series: [
    ChartSeries(
      name: "Sample Data",
      // For Scatter charts, categoriesRange acts as the X-axis (numeric)
      categoriesRange: r"Sheet1!$A$2:$A$20",  
      // valuesRange acts as the Y-axis (numeric)
      valuesRange: r"Sheet1!$B$2:$B$20",      
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1, width: 12, height: 15),
  showLegend: true,
);

sheet.addChart(chart);
```

> [!TIP]
> Use `DoubleCellValue` or `IntCellValue` in your cells for the best visual representation in Scatter Charts.

</details>

<details>
<summary><h3>🕸️ Radar Chart</h3></summary>

Radar/spider chart for multivariate data. Can be filled or lines-only.

```dart
// Filled radar chart
var chart = RadarChart(
  title: "Performance Metrics",
  series: [
    ChartSeries(
      name: "Employee A",
      categoriesRange: r"Sheet1!$A$2:$A$7",
      valuesRange: r"Sheet1!$B$2:$B$7",
    ),
    ChartSeries(
      name: "Employee B",
      categoriesRange: r"Sheet1!$A$2:$A$7",
      valuesRange: r"Sheet1!$C$2:$C$7",
    ),
  ],
  anchor: ChartAnchor.at(column: 5, row: 1),
  filled: true,  // Set to false for lines only
);

sheet.addChart(chart);
```

</details>

### Chart Positioning with ChartAnchor

The `ChartAnchor` defines where and how large your chart appears on the worksheet.

```dart
// Simple positioning with automatic sizing
ChartAnchor.at(
  column: 5,      // Starting column (0-indexed)
  row: 1,         // Starting row (0-indexed)
  width: 10,      // Width in columns (default: 8)
  height: 15,     // Height in rows (default: 15)
)

// Or specify exact positions
ChartAnchor(
  fromColumn: 5,
  fromRow: 1,
  toColumn: 15,   // End column
  toRow: 16,      // End row
)
```

### Multiple Charts

You can add multiple charts to the same sheet or different sheets.

```dart
var sheet1 = excel['Sales'];
var sheet2 = excel['Analysis'];

// Add first chart to Sales sheet
sheet1.addChart(ColumnChart(
  title: "Monthly Sales",
  series: [...],
  anchor: ChartAnchor.at(column: 5, row: 1),
));

// Add second chart to Sales sheet
sheet1.addChart(LineChart(
  title: "Sales Trend",
  series: [...],
  anchor: ChartAnchor.at(column: 5, row: 18),
));

// Add chart to Analysis sheet
sheet2.addChart(PieChart(
  title: "Distribution",
  series: [...],
  anchor: ChartAnchor.at(column: 2, row: 2),
));
```

### Chart Colors

Charts automatically use a professional 12-color palette:

1. #4472C4 - Blue
2. #ED7D31 - Orange  
3. #70AD47 - Green
4. #FFC000 - Gold
5. #5B9BD5 - Light Blue
6. #C5504B - Red
7. #8064A2 - Purple
8. #4BACC6 - Cyan
9. #9BBB59 - Olive
10. #F79646 - Light Orange
11. #17B897 - Teal
12. #E83352 - Crimson

Colors are automatically assigned to each series. If you have more than 12 series, colors will repeat.

For more details about chart colors and customization, see the [CHART_COLORS_GUIDE.md](CHART_COLORS_GUIDE.md) file.

### Important Notes

- **Range Format**: Use Excel's absolute reference format for ranges (e.g., `r"Sheet1!$A$2:$A$10"`)
- **Pie/Doughnut Charts**: These typically use only one series
- **Multiple Series**: Column, Bar, Line, Area, Scatter, and Radar charts support multiple series
- **Chart Title**: The title parameter is optional but recommended for clarity
- **Legend**: Set `showLegend: false` to hide the chart legend
- **Microsoft Excel Compatibility** *(v1.2.0+)*: Chart XML is fully compliant with the OOXML spec. Series names use the correct `<c:tx><c:v>` structure and element ordering inside `<c:chart>` follows the required CT_Chart sequence, ensuring charts open without errors in Microsoft Excel and Excel 365.

</details>

<details open>
<summary><h2>🔄 Pivot Tables (Tablas Dinámicas)</h2></summary>

Excel Community supports programmatically creating dynamic Pivot Tables in your worksheets. Pivot Tables allow you to summarize, analyze, explore, and present summary data from a larger dataset.

### Basic Pivot Table Example

```dart
// Create an Excel file and add source data
var excel = Excel.createExcel();
var sheet = excel['Sales Data'];
excel.delete('Sheet1');

// Add headers
sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Region"));
sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Product"));
sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Amount"));

// Add data rows
final data = [
  ['North', 'Apples', 100],
  ['North', 'Oranges', 150],
  ['South', 'Apples', 200],
  ['South', 'Oranges', 120],
  ['North', 'Apples', 80],
];

for (var i = 0; i < data.length; i++) {
  sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(data[i][0] as String));
  sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), TextCellValue(data[i][1] as String));
  sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), IntCellValue(data[i][2] as int));
}

// Create a new sheet for the pivot report
var reportSheet = excel['Pivot Report'];

// Define the pivot table
final pivotTable = PivotTable(
  name: 'PivotTable1',
  sourceSheet: 'Sales Data',
  sourceRange: 'A1:C6', // Source data range (including headers)
  targetCell: CellIndex.indexByString('A3'), // Top-left position on the target sheet
  rows: ['Region'], // Row fields
  columns: ['Product'], // Column fields
  values: [
    PivotTableValue(
      field: 'Amount',
      function: PivotValueFunction.sum,
      customName: 'Total Sales', // Optional custom header name
    ),
  ],
);

// Add pivot table to sheet
reportSheet.addPivotTable(pivotTable);

// Save the file
var bytes = excel.save();
```

### Supported Aggregation Functions

The `PivotValueFunction` enum provides the following data aggregation functions:

| Function | Description |
|---|---|
| `PivotValueFunction.sum` | Sums all numeric values in the field (Default) |
| `PivotValueFunction.count` | Counts all non-blank cells (both numeric and text) |
| `PivotValueFunction.average` | Calculates the average of numeric values |
| `PivotValueFunction.max` | Finds the maximum value |
| `PivotValueFunction.min` | Finds the minimum value |
| `PivotValueFunction.product` | Multiplies all numeric values |
| `PivotValueFunction.countNums` | Counts only cells containing numeric values |
| `PivotValueFunction.stdDev` | Estimates standard deviation of a sample |
| `PivotValueFunction.stdDevp` | Calculates standard deviation of an entire population |
| `PivotValueFunction.varVal` | Estimates variance of a sample |
| `PivotValueFunction.varp` | Calculates variance of an entire population |

### Key Details & Customization

- **Fields**: Row and column names specified in `rows` and `columns` must match the header names in your source sheet exactly (e.g. `'Region'`, `'Product'`).
- **Source Range**: Ensure `sourceRange` fully covers the headers and all data rows (e.g. `'A1:C6'`).
- **Values**: You can define multiple values in the `values` list to calculate multiple summaries (e.g. both `sum` and `average` for a field, or summaries of different fields).
- **Target Cell**: The `targetCell` indicates where the top-left of the pivot table will be rendered on the target sheet. Make sure you don't overwrite existing data in that area.

</details>

<details open>
<summary><h2>💾 Saving</h2></summary>

### On Flutter Web

```dart
// when you are in flutter web then save() downloads the excel file.

// Call function save() to download the file
var fileBytes = excel.save(fileName: 'My_Excel_File_Name.xlsx');
```

### On Android / iOS

For getting saving directory on Android or iOS, Use: [path_provider](https://pub.dev/packages/path_provider)

```dart
var fileBytes = excel.save();
var directory = await getApplicationDocumentsDirectory();

File(join('$directory/output_file_name.xlsx'))
  ..createSync(recursive: true)
  ..writeAsBytesSync(fileBytes);
```

</details>
</details>

## Credits

This project is based on the [excel](https://github.com/justkawal/excel) package created by **Kawaljeet Singh (justkawal)**. 

We are grateful for the original work and continue to maintain and enhance this library for the community.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

**Original Author:** Kawaljeet Singh (justkawal)  
**Original Project:** https://github.com/justkawal/excel

**Current Maintainer:** DecksPlayer  
**Community Fork:** https://github.com/DecksPlayer/excel_community

## Donation
If you find this package helpful and want to support its development, consider making a donation. Thank you for your support!

[![PayPal](https://img.shields.io/badge/Donate-PayPal-ff3f59.svg?style=for-the-badge&logo=paypal)](https://www.paypal.com/paypalme/gonojuarez)
<details>
<summary><h2>📜 History & Legacy Breaking Changes</h2></summary>

### Breaking changes from 4.x.x to 5.x.x (Original excel library)

- Updated minimum Dart SDK to 3.6.0

### Breaking changes from 3.x.x to 4.x.x (Original excel library)

- Renamed `Formula` to `FormulaCellValue`
- Cells value now represented by the sealed class `CellValue` instead of `dynamic`. Subtypes are `TextCellValue` `FormulaCellValue`, `IntCellValue`, `DoubleCellValue`, `DateCellValue`, `TextCellValue`, `BoolCellValue`, `TimeCellValue`, `DateTimeCellValue` and they allow for exhaustive switch.

### Breaking changes from 2.x.x to 3.x.x (Original excel library)

- Renamed `getColAutoFits()` to `getColumnAutoFits()`, and changed return type to `Map<int, bool>` in `Sheet`
- Renamed `getColWidths()` to `getColumnWidths()`, and changed return type to `Map<int, double>` in `Sheet`
- Renamed `setColWidth()` to `setColumnWidth()` in `Sheet`

</details>
