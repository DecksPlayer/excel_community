import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateLockedCellsReportHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Protected Report'];
  excel.delete('Sheet1');

  // 1. Protect the sheet programmatically using a password
  sheet.protect('password');

  // Excel locks all cells by default under protected sheets.
  // Explicitly set locked: true for read-only cells, and locked: false for editable cells.
  final headerStyle = CellStyle(
    bold: true,
    locked: true,
    backgroundColorHex: ExcelColor.fromHexString('#D9D9D9'),
  );
  final editableStyle = CellStyle(
    locked: false,
    backgroundColorHex: ExcelColor.fromHexString(
      '#E2EFDA',
    ), // light green to indicate editability
  );
  final readOnlyStyle = CellStyle(locked: true);

  // 2. Setup the headers (Read-Only)
  sheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Read-Only Header"),
    cellStyle: headerStyle,
  );
  sheet.updateCell(
    CellIndex.indexByString("B1"),
    TextCellValue("Editable Values (Type Here!)"),
    cellStyle: headerStyle,
  );

  // 3. Setup data rows - column A is locked, column B is unlocked
  sheet.updateCell(
    CellIndex.indexByString("A2"),
    TextCellValue("North Sales"),
    cellStyle: readOnlyStyle,
  );
  sheet.updateCell(
    CellIndex.indexByString("B2"),
    IntCellValue(8500),
    cellStyle: editableStyle,
  );

  sheet.updateCell(
    CellIndex.indexByString("A3"),
    TextCellValue("South Sales"),
    cellStyle: readOnlyStyle,
  );
  sheet.updateCell(
    CellIndex.indexByString("B3"),
    IntCellValue(6400),
    cellStyle: editableStyle,
  );

  sheet.setColumnWidth(0, 25.0);
  sheet.setColumnWidth(1, 30.0);

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'protected_sales_report.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Programmatic Protected Sheet generated successfully!\n'
          'Password: password\n'
          'Column A is Locked, Column B is Unlocked (Editable).\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Protected Report Excel',
      fileName: 'protected_sales_report.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      return '✅ Protected Report saved successfully!\n'
          'Password: password\n'
          'Location: $outputFile';
    }
    return 'Save cancelled.';
  }
}
