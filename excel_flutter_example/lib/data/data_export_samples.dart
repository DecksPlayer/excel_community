import 'package:excel_community/excel_community.dart';

/// Sample workbook used by the Data Export wiki and its demo workbook, so
/// the outputs shown in the app are the ones the library really produces.
Excel buildDataExportSampleWorkbook() {
  final excel = Excel.createExcel();

  final customers = excel['Customers'];
  customers.appendRow([
    TextCellValue('Name'),
    TextCellValue('Age'),
    TextCellValue('Active'),
    TextCellValue('Joined'),
    TextCellValue('Score'),
  ]);
  customers.appendRow([
    TextCellValue('Ana'),
    IntCellValue(31),
    BoolCellValue(true),
    DateCellValue(year: 2024, month: 3, day: 5),
    DoubleCellValue(0.95),
  ]);
  customers.appendRow([
    TextCellValue('Smith, Luis'),
    IntCellValue(28),
    BoolCellValue(false),
    DateCellValue(year: 2025, month: 1, day: 20),
    DoubleCellValue(0.725),
  ]);
  for (var row = 1; row <= 2; row++) {
    customers.cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: row)).cellStyle =
        CellStyle(numberFormat: NumFormat.custom(formatCode: 'dd/mm/yyyy'));
    customers.cell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: row)).cellStyle =
        CellStyle(numberFormat: NumFormat.custom(formatCode: '0.0%'));
  }

  final orders = excel['Orders'];
  orders.appendRow([TextCellValue('Order'), TextCellValue('Customer'), TextCellValue('Total')]);
  orders.appendRow([TextCellValue('A-100'), TextCellValue('Ana'), DoubleCellValue(120.5)]);

  return excel;
}

/// JSON used by the import examples.
const String dataExportImportJson = '[\n'
    '  {"Name": "Eva", "Age": 40, "City": "Lima"},\n'
    '  {"Name": "Tom", "Age": 35, "Joined": "2026-01-15"}\n'
    ']';

/// Readable one-line-per-row rendering of export results.
String describeRows(List<Object?> rows) => rows.map(_describe).join('\n');

String _describe(Object? value) => switch (value) {
      null => 'null',
      String s => "'$s'",
      DateTime d => 'DateTime.utc(${d.year}, ${d.month}, ${d.day})',
      Duration d => 'Duration(${d.toString().split('.').first})',
      Map m => '{${m.entries.map((e) => "'${e.key}': ${_describe(e.value)}").join(', ')}}',
      List l => '[${l.map(_describe).join(', ')}]',
      _ => value.toString(),
    };
