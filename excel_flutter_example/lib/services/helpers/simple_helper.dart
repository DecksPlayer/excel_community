import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateSimpleExcelHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Name"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Age"));
  sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Alice"));
  sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(30));
  sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Bob"));
  sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(25));

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'simple_no_chart.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Simple Excel generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '\n📥 Check your Downloads folder\n'
          '📌 File: simple_no_chart.xlsx';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Simple Excel File',
      fileName: 'simple_no_chart.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Simple Excel saved successfully!\n'
          'Location: $outputFile\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
          '\n📌 Open with Excel to verify it works';
    }
    return 'Save cancelled.';
  }
}
