import 'dart:convert';
import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

import '../../data/data_export_samples.dart';

/// Demo workbook: the sample sheets, their JSON and CSV exports written into
/// worksheets, and a sheet imported from JSON with `appendRowsFromMaps`.
Excel buildDataExportWorkbook() {
  final excel = buildDataExportSampleWorkbook();
  final customers = excel['Customers'];

  void writeLines(String sheetName, String title, String text) {
    final sheet = excel[sheetName];
    sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue(title),
        cellStyle: CellStyle(bold: true, fontSize: 14));
    final lines = const LineSplitter().convert(text);
    for (var i = 0; i < lines.length; i++) {
      sheet.updateCell(
          CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 2), TextCellValue(lines[i]));
    }
    sheet.setColumnWidth(0, 70);
  }

  writeLines('JSON Export', "excel['Customers'].toJson(indent: '  ')", customers.toJson(indent: '  '));
  writeLines('CSV Export', "excel['Customers'].toCsv()", customers.toCsv(lineTerminator: '\n'));

  final maps = (jsonDecode(dataExportImportJson) as List).cast<Map<String, Object?>>();
  excel['Imported from JSON'].appendRowsFromMaps(maps);

  for (final name in ['Customers', 'Orders', 'Imported from JSON']) {
    for (var col = 0; col < excel[name].maxColumns; col++) {
      excel[name].setColumnWidth(col, 16);
    }
  }
  return excel;
}

Future<String> generateDataExportHelper() async {
  final excel = buildDataExportWorkbook();
  const summary = 'Customers, Orders, JSON Export, CSV Export, Imported from JSON';

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'data_export_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Data Export workbook generated successfully!\n'
          'Worksheets: $summary\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 File: data_export_example.xlsx';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Data Export Example',
      fileName: 'data_export_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Data Export workbook saved successfully!\n'
          'Location: $outputFile\n'
          'Worksheets: $summary\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return 'Save cancelled.';
  }
}
