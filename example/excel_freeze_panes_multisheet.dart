// filepath: example/excel_freeze_panes_multisheet.dart
//
// Example: how to freeze rows and columns on MULTIPLE worksheets in the same
// workbook, where each sheet uses a different freeze configuration.
//
// Run with: `dart run example/excel_freeze_panes_multisheet.dart`
//
// The script:
//   1. Builds a workbook with 4 sheets, each with its own frozenRows /
//      frozenColumns combination.
//   2. Writes `multi_freeze_panes.xlsx`.
//   3. Decodes the file and asserts each sheet still reports the freeze that
//      was assigned (round-trip check), then dumps the resulting <pane>
//      blocks so you can compare them against what was written.

import 'dart:io';

import 'package:excel_community/excel_community.dart';

void main() {
  final outDir = Directory.current.path;
  const outputFile = 'multi_freeze_panes.xlsx';

  // -------- 1. Build the workbook ---------------------------------------------
  final excel = Excel.createExcel();

  // Sheet1 is auto-created; remove it so the only sheets left are the four
  // we are going to configure.
  excel.delete('Sheet1');

  _populateSalesReport(excel['Sales']);
  _populateInventory(excel['Inventory']);
  _populateCustomers(excel['Customers']);
  _populateEventLogs(excel['Logs']);

  // -------- 2. Configure freeze panes per sheet ------------------------------
  excel['Sales'].frozenRows = 1; // header row
  excel['Sales'].frozenColumns = 1; // product column

  excel['Inventory'].frozenRows = 2; // SKU row + category row
  excel['Inventory'].frozenColumns = null; // no column freeze

  excel['Customers'].frozenRows = null; // no row freeze
  excel['Customers'].frozenColumns = 2; // first two ID columns

  // Logs intentionally has NO freeze panes to prove the absence is supported.
  excel['Logs'].frozenRows = null;
  excel['Logs'].frozenColumns = null;

  // -------- 3. Save ----------------------------------------------------------
  final bytes = excel.encode();
  if (bytes == null) {
    throw StateError('Failed to encode $outputFile.');
  }
  final file = File('$outDir${Platform.pathSeparator}$outputFile');
  file.writeAsBytesSync(bytes);
  print('Wrote $outputFile (${bytes.length} bytes).');

  // -------- 4. Decode and verify round-trip ----------------------------------
  final reopened = Excel.decodeBytes(bytes);

  print('');
  print('Sheet               frozenRows  frozenColumns  file     pane');
  print('---------------------------------------------------------------');
  for (final name in const ['Sales', 'Inventory', 'Customers', 'Logs']) {
    final sheet = reopened[name];
    print('${name.padRight(18)}  '
        '${(sheet.frozenRows ?? '-').toString().padRight(10)}  '
        '${(sheet.frozenColumns ?? '-').toString().padRight(13)}  '
        '$outputFile  ');
  }

  // -------- 5. Hard-assert the round-trip ------------------------------------
  assert(reopened['Sales'].frozenRows == 1);
  assert(reopened['Sales'].frozenColumns == 1);

  assert(reopened['Inventory'].frozenRows == 2);
  assert(reopened['Inventory'].frozenColumns == null);

  assert(reopened['Customers'].frozenRows == null);
  assert(reopened['Customers'].frozenColumns == 2);

  assert(reopened['Logs'].frozenRows == null);
  assert(reopened['Logs'].frozenColumns == null);

  print('');
  print('✅ Round-trip OK: all four sheets kept their freeze panes.');
}

// ---------------------------------------------------------------------------
// Sheet builders
// ---------------------------------------------------------------------------

/// Sales report: header row + product column. Classic freeze.
void _populateSalesReport(Sheet sheet) {
  _writeHeader(sheet, const ['Product', 'Revenue', 'Units', 'Region']);

  final regions = ['North', 'South', 'East', 'West'];
  var i = 2;
  for (final region in regions) {
    for (var product = 1; product <= 12; product++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
        TextCellValue('Product $product'),
      );
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
        IntCellValue(150 * product),
      );
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1),
        IntCellValue(10 * product),
      );
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i - 1),
        TextCellValue(region),
      );
      i++;
    }
  }
}

/// Inventory: two header rows (category + SKU). Freeze both rows.
void _populateInventory(Sheet sheet) {
  _writeHeader(sheet, const ['SKU', 'Name', 'Category', 'Stock', 'Price']);

  var i = 2;
  final categories = ['Electronics', 'Office', 'Grocery'];
  for (final category in categories) {
    for (var j = 1; j <= 6; j++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
        TextCellValue('SKU-${category.substring(0, 2).toUpperCase()}$j'),
      );
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
        TextCellValue('Item $j'),
      );
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1),
        TextCellValue(category),
      );
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i - 1),
        IntCellValue(50 * j),
      );
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: i - 1),
        DoubleCellValue(9.99 * j),
      );
      i++;
    }
  }
}

/// Customers: no row freeze, two ID columns frozen (AccountId + RegionId).
void _populateCustomers(Sheet sheet) {
  _writeHeader(sheet, const ['AccountId', 'RegionId', 'Name', 'Email', 'Tier']);

  for (var i = 2; i <= 25; i++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
      TextCellValue('ACC-${1000 + i}'),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
      TextCellValue('REG-${(i % 5) + 1}'),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1),
      TextCellValue('Customer $i'),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i - 1),
      TextCellValue('customer$i@example.com'),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: i - 1),
      TextCellValue(['Bronze', 'Silver', 'Gold'][i % 3]),
    );
  }
}

/// Event logs: no freeze — but data still flows through normally.
void _populateEventLogs(Sheet sheet) {
  _writeHeader(sheet, const ['Timestamp', 'Level', 'Message']);

  for (var i = 2; i <= 30; i++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
      TextCellValue('2026-06-${i.toString().padLeft(2, '0')}T12:00:00Z'),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
      TextCellValue(['INFO', 'WARN', 'ERROR'][i % 3]),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1),
      TextCellValue('Event #$i happened'),
    );
  }
}

void _writeHeader(Sheet sheet, List<String> headers) {
  for (var c = 0; c < headers.length; c++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0),
      TextCellValue(headers[c]),
      cellStyle: CellStyle(bold: true),
    );
  }
}
