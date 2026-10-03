const String dataExportSnippet = r'''
import 'dart:convert';
import 'package:excel_community/excel_community.dart';

void exportCustomers(Excel excel) {
  final sheet = excel['Customers'];

  // Rows as maps with native Dart values
  final rows = sheet.rowsAsMaps();

  // ...or as the text Excel displays
  final shown = sheet.rowsAsMaps(mode: ExportValueMode.displayText);

  // JSON and CSV
  final json = sheet.toJson(indent: '  ');
  final workbookJson = excel.toJson();
  final csv = sheet.toCsv(separator: ';');

  // Import: JSON -> maps -> rows
  final maps = (jsonDecode(json) as List).cast<Map<String, Object?>>();
  excel['Copy'].appendRowsFromMaps(maps);
}
''';
