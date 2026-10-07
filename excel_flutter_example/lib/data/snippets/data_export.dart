const String dataExportSnippet = r'''
import 'dart:convert';
import 'dart:io';
import 'package:excel_community/excel_community.dart';

/// Generates the complete Data Export & Transformation demo workbook:
/// - Base sheets with structured data (Customers, Orders)
/// - "JSON Export": formatted sheet containing sheet.toJson() output
/// - "CSV Export": sheet containing sheet.toCsv() output
/// - "Imported from JSON": recreates a table from JSON maps via appendRowsFromMaps
void generateDataExportWorkbook() {
  final excel = Excel.createExcel();

  // 1. Setup 'Customers' sheet with typed columns
  final customers = excel['Customers'];
  customers.appendRow([
    TextCellValue('id'),
    TextCellValue('name'),
    TextCellValue('balance'),
    TextCellValue('registered'),
    TextCellValue('active'),
  ]);

  final sampleData = [
    (101, 'Alice Smith', 1240.50, DateTime(2023, 4, 15), true),
    (102, 'Bob Jones', 0.00, DateTime(2024, 1, 10), false),
    (103, 'Charlie Brown', -45.20, DateTime(2022, 11, 28), true),
    (104, 'Diana Prince', 3500.00, DateTime(2021, 6, 5), true),
  ];

  for (final (id, name, balance, reg, active) in sampleData) {
    customers.appendRow([
      IntCellValue(id),
      TextCellValue(name),
      DoubleCellValue(balance),
      DateCellValue(year: reg.year, month: reg.month, day: reg.day),
      BoolCellValue(active),
    ]);
  }

  // 2. Export Customers to JSON and CSV text
  final customersJson = customers.toJson(indent: '  ');
  final customersCsv = customers.toCsv(lineTerminator: '\n');

  // Helper to write multiline text to a worksheet
  void writeLines(String sheetName, String title, String text) {
    final sheet = excel[sheetName];
    sheet.updateCell(
      CellIndex.indexByString('A1'),
      TextCellValue(title),
      cellStyle: CellStyle(bold: true, fontSize: 14),
    );
    final lines = const LineSplitter().convert(text);
    for (var i = 0; i < lines.length; i++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 2),
        TextCellValue(lines[i]),
      );
    }
    sheet.setColumnWidth(0, 70);
  }

  writeLines('JSON Export', "Customers sheet exported to JSON:", customersJson);
  writeLines('CSV Export', "Customers sheet exported to CSV:", customersCsv);

  // 3. Import JSON maps back into a new worksheet
  const rawJson = '[\n'
      '  {"id": 201, "name": "Elena Rostova", "country": "ES", "vip": true},\n'
      '  {"id": 202, "name": "Marcus Aurelius", "country": "IT", "vip": false},\n'
      '  {"id": 203, "name": "Sophie Martin", "country": "FR", "vip": true}\n'
      ']';
  final maps = (jsonDecode(rawJson) as List).cast<Map<String, Object?>>();
  excel['Imported from JSON'].appendRowsFromMaps(maps);

  for (final name in ['Customers', 'Imported from JSON']) {
    for (var col = 0; col < excel[name].maxColumns; col++) {
      excel[name].setColumnWidth(col, 18);
    }
  }

  // 4. Save workbook
  final bytes = excel.save(fileName: 'data_export_example.xlsx');
  // Or: File('data_export_example.xlsx').writeAsBytesSync(excel.encode()!);
}
''';
