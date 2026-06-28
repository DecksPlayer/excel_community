import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateMultiSheetsHelper() async {
  var excel = Excel.createExcel();
  
  // Sheet 1: Summary Sheet
  var summarySheet = excel['Summary'];
  excel.delete('Sheet1'); // Remove default

  summarySheet.updateCell(
    CellIndex.indexByString("A1"), 
    TextCellValue("Monthly Financial Summary"),
    cellStyle: CellStyle(bold: true, fontSize: 14, fontColorHex: ExcelColor.blue),
  );

  // Sheet 2: Revenues Detail
  var revenueSheet = excel['Revenues'];
  revenueSheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Source"), cellStyle: CellStyle(bold: true));
  revenueSheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Amount"), cellStyle: CellStyle(bold: true));
  revenueSheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("SaaS Subscriptions"));
  revenueSheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(15400));
  revenueSheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Consulting"));
  revenueSheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(4200));

  // Sheet 3: Expenses Detail
  var expenseSheet = excel['Expenses'];
  expenseSheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"), cellStyle: CellStyle(bold: true));
  expenseSheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Cost"), cellStyle: CellStyle(bold: true));
  expenseSheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Hosting & Cloud"));
  expenseSheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(3100));
  expenseSheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Marketing"));
  expenseSheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(1200));

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'multisheet_financials.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Multi-sheet Excel generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 Includes: Summary, Revenues, and Expenses sheets.';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Multi-sheet Excel',
      fileName: 'multisheet_financials.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      return '✅ Multi-sheet Excel saved successfully!\n'
          'Location: $outputFile';
    }
    return 'Save cancelled.';
  }
}

Future<String> generatePivotTemplateHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Sales Data'];
  excel.delete('Sheet1');

  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Region"), cellStyle: CellStyle(bold: true));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Amount"), cellStyle: CellStyle(bold: true));

  sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("North Region"));
  sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(45000));

  sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("South Region"));
  sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(32000));

  sheet.updateCell(CellIndex.indexByString("A4"), TextCellValue("East Region"));
  sheet.updateCell(CellIndex.indexByString("B4"), IntCellValue(51000));

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'sales_report_pivot.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Template data updated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'Download started. (In a real app, loading a template with pre-configured Pivot Tables preserves them entirely)';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Sales Report Excel',
      fileName: 'sales_report_pivot.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      return '✅ Template data saved successfully!\n'
          'Location: $outputFile';
    }
    return 'Save cancelled.';
  }
}

Future<String> generateFreezePanesHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Sales Report'];
  excel.delete('Sheet1');

  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Product Name"), cellStyle: CellStyle(bold: true));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Sales Revenue"), cellStyle: CellStyle(bold: true));

  for (int i = 2; i <= 50; i++) {
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1), TextCellValue("Product $i"));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1), IntCellValue(150 * i));
  }

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'sales_report_frozen.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Frozen Panes template data updated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'Download started. (In a real app, loading a template with pre-configured Freeze Panes preserves them.)';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Sales Report Excel',
      fileName: 'sales_report_frozen.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      return '✅ Sales Report saved successfully!\n'
          'Location: $outputFile';
    }
    return 'Save cancelled.';
  }
}
