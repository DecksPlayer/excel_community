# 🎨 Cell Styling & Operations

This guide covers updating cell values, setting styles and borders, adding custom number formatting, inserting/deleting columns or rows, and performing sheet-level actions like copy, rename, merge, and link.

---

## ✏️ Updating Cell Values

To modify a cell, get a reference to a sheet and access the cell via `CellIndex`. 

```dart
Sheet sheetObject = excel['Sheet1'];

// Access cells using sheetObject.cell()
var cell = sheetObject.cell(CellIndex.indexByString('A1'));

// Assigning values (using the sealed class CellValue)
cell.value = null; // Removing any value
cell.value = TextCellValue('Some Text');
cell.value = IntCellValue(8);
cell.value = BoolCellValue(true);
cell.value = DoubleCellValue(13.37);
cell.value = DateCellValue(year: 2023, month: 4, day: 20);
cell.value = TimeCellValue(hour: 20, minute: 15, second: 5);
cell.value = DateTimeCellValue(year: 2023, month: 4, day: 20, hour: 15, minute: 30);
```

### Checking Cell Types

You can use Dart's pattern matching to inspect the cell value:

```dart
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
```

---

## 🎨 Cell Styling Options

Cell styles are applied by creating a `CellStyle` object and assigning it to `cell.cellStyle`.

```dart
CellStyle cellStyle = CellStyle(
  backgroundColorHex: '#1AFF1A', 
  fontFamily: getFontFamily(FontFamily.Calibri),
);

cellStyle.underline = Underline.Single; // or Underline.Double
cellStyle.isStrikethrough = true;       // enable strikethrough text

cell.cellStyle = cellStyle;
```

Here are all the fields supported by `CellStyle`:

| Property | Description |
|---|---|
| `fontFamily` | Font name, e.g., via `getFontFamily(FontFamily.Arial)`. Over 182 font families supported. |
| `fontSize` | Font size as an integer (e.g. `fontSize: 15`). |
| `bold` | Bold text (`true` or `false`). |
| `italic` | Italic text (`true` or `false`). |
| `underline` | Underline enum style: `Underline.None`, `Underline.Single`, `Underline.Double`. |
| `isStrikethrough` | Strikethrough text (`true` or `false`). |
| `fontColorHex` | Font color in hex format (e.g. `'#0000FF'`). |
| `rotation` | Rotation of text in degrees, ranging from `-90` to `90`. |
| `backgroundColorHex` | Background color of cell in hex format (e.g. `'#faf487'`). |
| `wrap` | Text wrapping: `TextWrapping.WrapText` or `TextWrapping.Clip`. |
| `verticalAlign` | Vertical alignment: `VerticalAlign.Top`, `VerticalAlign.Center`, `VerticalAlign.Bottom`. |
| `horizontalAlign` | Horizontal alignment: `HorizontalAlign.Left`, `HorizontalAlign.Center`, `HorizontalAlign.Right`. |
| `leftBorder` | The left border configuration (`Border` object). |
| `rightBorder` | The right border configuration (`Border` object). |
| `topBorder` | The top border configuration (`Border` object). |
| `bottomBorder` | The bottom border configuration (`Border` object). |
| `diagonalBorder` | The diagonal border configuration (`Border` object). |
| `diagonalBorderUp` | Whether to display the upward diagonal border (`true` or `false`). |
| `diagonalBorderDown` | Whether to display the downward diagonal border (`true` or `false`). |
| `numberFormat` | Formatting style for numeric, date, time or custom display. |

---

## 🔲 Borders

Borders are configured using the `Border` class, specifying a `borderStyle` and `borderColorHex`.

```dart
CellStyle cellStyle = CellStyle(
  leftBorder: Border(borderStyle: BorderStyle.Thin),
  rightBorder: Border(borderStyle: BorderStyle.Thin),
  topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: 'FFFF0000'), // red
  bottomBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: 'FF0000FF'), // blue
);
```

Available `BorderStyle` values:
- `BorderStyle.None`
- `BorderStyle.Thin`
- `BorderStyle.Medium`
- `BorderStyle.Thick`
- `BorderStyle.Dashed`
- `BorderStyle.Dotted`
- `BorderStyle.Double`
- `BorderStyle.Hair`
- `BorderStyle.MediumDashed`
- `BorderStyle.DashDot`
- `BorderStyle.MediumDashDot`
- `BorderStyle.DashDotDot`
- `BorderStyle.MediumDashDotDot`
- `BorderStyle.SlantDashDot`

---

## 🧮 Number and Date Formatting

You can apply standard or custom formatting for numbers, dates, times, and financial amounts:

```dart
cell.cellStyle = (cell.cellStyle ?? CellStyle()).copyWith(
  // Custom numeric formatting
  numberFormat: CustomNumericNumFormat('#,##0.00 \\m\\²'),

  // Custom date/time formatting
  numberFormat: CustomDateTimeNumFormat('m/d/yy h:mm'),

  // Built-in standard formats
  numberFormat: NumFormat.standard_14, // standard date format (m/d/yy)
  numberFormat: NumFormat.standard_5,  // currency: $#,##0_);($#,##0)
  numberFormat: NumFormat.standard_41, // accounting: _(* #,##0_);...
);
```

> [!TIP]
> The `numberFormat` changes automatically if you assign a `CellValue` type that does not support the format set previously. Always write the `CellValue` first, then assign the `numberFormat`.

---

## 📐 Cell Operations (Merging, Finding & Replacing)

### Merging Cells
Merge a range of cells into one.
```dart
// Merges from A1 to E4 and puts custom text in it
sheetObject.merge(
  CellIndex.indexByString('A1'), 
  CellIndex.indexByString('E4'), 
  customValue: TextCellValue('Put this text after merge'),
);
```

### Listing Merged Ranges
Retrieve all currently merged cell ranges:
```dart
sheetObject.spannedItems.forEach((cells) {
  print('Merged Range: $cells');
});
```

### Un-merging Cells
```dart
sheetObject.unMerge('A1:E4');
```

### Find and Replace
Find matching cells and replace their content. Supports RegExp patterns.
```dart
// Returns count of replacements made
int replacedCount = sheetObject.findAndReplace('Flutter', 'Google');
```

---

## 🧱 Row & Column Operations

### Inserting and Removing Rows/Columns

```dart
// Insert a column at index 8
sheetObject.insertColumn(8);

// Remove a column at index 18
sheetObject.removeColumn(18);

// Insert a row at index 82
sheetObject.insertRow(82);

// Remove a row at index 80
sheetObject.removeRow(80);
```

### Row Iterables & Appending

#### Inserting Row Iterables
Insert a list of cell values starting at a specific row:
```dart
List<CellValue> dataList = [
  TextCellValue('Google'), 
  TextCellValue('loves'), 
  TextCellValue('Flutter'),
];

// Puts dataList in the 8th row (index 8)
sheetObject.insertRowIterables(dataList, 8);
```

Options for `insertRowIterables`:
- `startingColumn`: Column index to start inserting values from.
- `overwriteMergedCells`: Defaults to `true`. Set to `false` to avoid overwriting values in merged cells.

#### Appending Rows
Add a new row at the very end of the sheet:
```dart
sheetObject.appendRow([
  TextCellValue('Flutter'), 
  TextCellValue('till'), 
  TextCellValue('Eternity'),
]);
```

---

## 🗂️ Sheet Management

### Make Sheet Right-to-Left (RTL)
```dart
var sheetObject = excel['Sheet1'];
sheetObject.rtl = true; // Default is false (LTR)
```

### Copying Sheet Contents
```dart
// Copies sheet contents from existingSheetName to anotherSheetName (creates if it doesn't exist)
excel.copy('existingSheetName', 'anotherSheetName');
```

### Renaming a Sheet
```dart
excel.rename('existingSheetName', 'newSheetName');
```

### Deleting a Sheet
Note: The sheet must exist and there must be at least 2 sheets in the workbook to perform a deletion.
```dart
excel.delete('existingSheetName');
```

### Linking Worksheets
Two names can refer to the exact same worksheet object. Modifying either modifies both.
```dart
excel.link('sheetName', sheetObject);
```

### Un-linking Worksheets
```dart
excel.unLink('sheetName');
// After unlinking, create a new reference if you wish to modify it separately
Sheet unlinkedSheet = excel['sheetName'];
```

### Default Opening Sheet

Get the name of the sheet that Excel displays first when opened:
```dart
var defaultSheet = excel.getDefaultSheet();
print('Default Sheet: $defaultSheet');
```

Set the default sheet to open:
```dart
var isSet = excel.setDefaultSheet('SheetName');
if (isSet) {
  print('Successfully changed default sheet.');
}
```
