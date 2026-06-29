// filepath: example/excel_freeze_panes.dart
//
// Example: how to freeze (lock) rows and columns in a worksheet.
//
// Run with: `dart run example/excel_freeze_panes.dart`
// The script writes four small XLSX files next to itself so the resulting
// `<pane>` / `<selection>` blocks can be inspected directly inside the
// `xl/worksheets/sheet1.xml` part.

import 'dart:io';

import 'package:excel_community/excel_community.dart';

void main() {
  final outDir = Directory.current.path;

  _writeBasicFreeze(outDir);
  _writeOnlyRows(outDir);
  _writeOnlyColumns(outDir);
  _writeNoFreeze(outDir);

  print('Done. Wrote 4 freeze-pane samples to: $outDir');
}

/// 1) Freeze first row + first column.
void _writeBasicFreeze(String outDir) {
  final excel = Excel.createExcel();
  final sheet = excel['Sales Report'];
  excel.delete('Sheet1');

  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Product Name'),
    cellStyle: CellStyle(bold: true),
  );
  sheet.updateCell(
    CellIndex.indexByString('B1'),
    TextCellValue('Sales Revenue'),
    cellStyle: CellStyle(bold: true),
  );

  for (var i = 2; i <= 30; i++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
      TextCellValue('Product $i'),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
      IntCellValue(150 * i),
    );
  }

  // 👇 The freeze-pane API: simply assign to these two setters.
  sheet.frozenRows = 1;
  sheet.frozenColumns = 1;

  _save(excel, 'sales_report_frozen_basic.xlsx', outDir);
}

/// 2) Freeze only the first 2 rows (no column lock).
void _writeOnlyRows(String outDir) {
  final excel = Excel.createExcel();
  final sheet = excel['Only Rows'];
  excel.delete('Sheet1');

  for (var r = 0; r < 20; r++) {
    for (var c = 0; c < 4; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        TextCellValue('R${r + 1}C${c + 1}'),
      );
    }
  }

  sheet.frozenRows = 2;
  sheet.frozenColumns = null; // explicit: do not freeze columns.

  _save(excel, 'sales_report_frozen_rows.xlsx', outDir);
}

/// 3) Freeze only the first column (no row lock).
void _writeOnlyColumns(String outDir) {
  final excel = Excel.createExcel();
  final sheet = excel['Only Columns'];
  excel.delete('Sheet1');

  for (var r = 0; r < 20; r++) {
    for (var c = 0; c < 4; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        TextCellValue('R${r + 1}C${c + 1}'),
      );
    }
  }

  sheet.frozenRows = null;
  sheet.frozenColumns = 1;

  _save(excel, 'sales_report_frozen_columns.xlsx', outDir);
}

/// 4) No freeze panes — clean sheetView, no <pane>.
void _writeNoFreeze(String outDir) {
  final excel = Excel.createExcel();
  final sheet = excel['No Freeze'];
  excel.delete('Sheet1');

  for (var r = 0; r < 5; r++) {
    for (var c = 0; c < 4; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        TextCellValue('R${r + 1}C${c + 1}'),
      );
    }
  }

  sheet.frozenRows = null;
  sheet.frozenColumns = null;

  _save(excel, 'sales_report_no_freeze.xlsx', outDir);
}

void _save(Excel excel, String fileName, String outDir) {
  final bytes = excel.encode();
  if (bytes == null) {
    throw StateError('Failed to encode "$fileName".');
  }
  final file = File('$outDir${Platform.pathSeparator}$fileName');
  file.writeAsBytesSync(bytes);
  print('  wrote $fileName '
      '(${bytes.length} bytes, '
      'frozenRows=${excel[excel.getDefaultSheet()!].frozenRows}, '
      'frozenColumns=${excel[excel.getDefaultSheet()!].frozenColumns})');
}
