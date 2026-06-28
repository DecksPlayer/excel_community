# 🚀 Getting Started with Excel Community

This guide details how to install the package, read Excel spreadsheets, create new workbooks, and save files across multiple platforms (Android, iOS, Web, and Desktop).

---

## 📦 Installation

Add `excel_community` to your `pubspec.yaml` dependencies:

```yaml
dependencies:
  excel_community: ^2.0.2
```

Then, run:
```bash
flutter pub get
```

---

## 📄 Reading XLSX Files

Excel files are read by passing a list of bytes to `Excel.decodeBytes()`.

### 1. Read a File from Local Storage (Mobile / Desktop)

```dart
import 'dart:io';
import 'package:excel_community/excel_community.dart';

void readExcel() {
  var file = 'path/to/your/excel_file.xlsx';
  var bytes = File(file).readAsBytesSync();
  var excel = Excel.decodeBytes(bytes);

  for (var table in excel.tables.keys) {
    print(table); // Sheet Name
    var sheet = excel.tables[table]!;
    print('Max Columns: ${sheet.maxColumns}');
    print('Max Rows: ${sheet.maxRows}');

    for (var row in sheet.rows) {
      for (var cell in row) {
        if (cell == null) continue;
        print('Cell ${cell.rowIndex}/${cell.columnIndex}');
        final value = cell.value;
        final numFormat = cell.cellStyle?.numberFormat ?? NumFormat.standard_0;

        switch (value) {
          case null:
            print('  empty cell (format: $numFormat)');
          case TextCellValue():
            print('  text: ${value.value}');
          case FormulaCellValue():
            print('  formula: ${value.formula} (format: $numFormat)');
          case IntCellValue():
            print('  int: ${value.value} (format: $numFormat)');
          case BoolCellValue():
            print('  bool: ${value.value ? 'YES!!' : 'NO..'} (format: $numFormat)');
          case DoubleCellValue():
            print('  double: ${value.value} (format: $numFormat)');
          case DateCellValue():
            print('  date: ${value.year}-${value.month}-${value.day} (${value.asDateTimeLocal()})');
          case TimeCellValue():
            print('  time: ${value.hour}:${value.minute} (${value.asDuration()})');
          case DateTimeCellValue():
            print('  date with time: ${value.year}-${value.month}-${value.day} ${value.hour}:${value.minute} (${value.asDateTimeLocal()})');
        }
      }
    }
  }
}
```

### 2. Read an XLSX File in Flutter Web

On the web, users upload files through their browser. We can use the [`file_picker`](https://pub.dev/packages/file_picker) package to read the bytes directly.

```dart
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<void> readExcelWeb() async {
  FilePickerResult? pickedFile = await FilePicker.platform.pickFiles(
    type: FileType.custom,
    allowedExtensions: ['xlsx'],
    allowMultiple: false,
  );

  if (pickedFile != null) {
    var bytes = pickedFile.files.single.bytes;
    if (bytes != null) {
      var excel = Excel.decodeBytes(bytes);
      for (var table in excel.tables.keys) {
        print(table); // Sheet Name
        var sheet = excel.tables[table]!;
        for (var row in sheet.rows) {
          print('$row');
        }
      }
    }
  }
}
```

### 3. Read an XLSX File from Flutter Assets

To read a file packaged within your Flutter assets, register it first in `pubspec.yaml`:

```yaml
flutter:
  assets:
    - assets/existing_excel_file.xlsx
```

Then, load and parse the asset:

```dart
import 'package:flutter/services.dart' show rootBundle;
import 'package:excel_community/excel_community.dart';

Future<void> readExcelAsset() async {
  final data = await rootBundle.load('assets/existing_excel_file.xlsx');
  var bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
  var excel = Excel.decodeBytes(bytes);

  for (var table in excel.tables.keys) {
    var sheet = excel.tables[table]!;
    for (var row in sheet.rows) {
      print('$row');
    }
  }
}
```

---

## ➕ Creating a New XLSX File

Creating a new spreadsheet is simple and automatically creates one empty sheet named `Sheet1`.

```dart
import 'package:excel_community/excel_community.dart';

void createNewExcel() {
  var excel = Excel.createExcel();
  // Automatically creates a sheet named 'Sheet1'
  Sheet sheetObject = excel['Sheet1'];
  
  // Now you can start updating cells or adding features
}
```

---

## 💾 Saving XLSX Files

The library returns a `List<int>` representing the zipped XLSX file bytes when you call `excel.save()`.

### 1. Saving in Flutter Web (Triggers File Download)

When running on the web, calling `excel.save(fileName: ...)` automatically downloads the file.

```dart
// Under Flutter Web, this will trigger a browser download
var fileBytes = excel.save(fileName: 'My_Excel_File_Name.xlsx');
```

### 2. Saving on Mobile & Desktop

On Android, iOS, or Desktop, you must write the returned bytes to a physical file on the device's storage. You can use the [`path_provider`](https://pub.dev/packages/path_provider) package to find standard storage directories.

```dart
import 'dart:io';
import 'package:path/to/path.dart'; // for join()
import 'package:path_provider/path_provider.dart';
import 'package:excel_community/excel_community.dart';

Future<void> saveExcelFile(Excel excel) async {
  var fileBytes = excel.save();
  var directory = await getApplicationDocumentsDirectory();

  final file = File(join('${directory.path}/output_file_name.xlsx'));
  await file.create(recursive: true);
  await file.writeAsBytes(fileBytes!);
}
```
