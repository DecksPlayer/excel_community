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
    cellStyle: CellStyle(
      bold: true,
      fontSize: 14,
      fontColorHex: ExcelColor.blue,
    ),
  );

  // Sheet 2: Revenues Detail
  var revenueSheet = excel['Revenues'];
  revenueSheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Source"),
    cellStyle: CellStyle(bold: true),
  );
  revenueSheet.updateCell(
    CellIndex.indexByString("B1"),
    TextCellValue("Amount"),
    cellStyle: CellStyle(bold: true),
  );
  revenueSheet.updateCell(
    CellIndex.indexByString("A2"),
    TextCellValue("SaaS Subscriptions"),
  );
  revenueSheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(15400));
  revenueSheet.updateCell(
    CellIndex.indexByString("A3"),
    TextCellValue("Consulting"),
  );
  revenueSheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(4200));

  // Sheet 3: Expenses Detail
  var expenseSheet = excel['Expenses'];
  expenseSheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Category"),
    cellStyle: CellStyle(bold: true),
  );
  expenseSheet.updateCell(
    CellIndex.indexByString("B1"),
    TextCellValue("Cost"),
    cellStyle: CellStyle(bold: true),
  );
  expenseSheet.updateCell(
    CellIndex.indexByString("A2"),
    TextCellValue("Hosting & Cloud"),
  );
  expenseSheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(3100));
  expenseSheet.updateCell(
    CellIndex.indexByString("A3"),
    TextCellValue("Marketing"),
  );
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

Future<String> generateMultiFreezePanesHelper() async {
  final excel = Excel.createExcel();
  excel.delete('Sheet1'); // remove default

  // -------- Sheet 1: Sales — header row + product column frozen --------------
  final salesSheet = excel['Sales'];
  _writeHeader(salesSheet, const ['Product', 'Revenue', 'Units', 'Region']);
  for (var i = 2; i <= 25; i++) {
    salesSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
      TextCellValue('Product $i'),
    );
    salesSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
      IntCellValue(150 * i),
    );
    salesSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1),
      IntCellValue(10 * i),
    );
    salesSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i - 1),
      TextCellValue(['North', 'South', 'East', 'West'][i % 4]),
    );
  }
  salesSheet.frozenRows = 1;
  salesSheet.frozenColumns = 1;

  // -------- Sheet 2: Inventory — two header rows frozen ----------------------
  final inventorySheet = excel['Inventory'];
  _writeHeader(inventorySheet, const [
    'SKU',
    'Name',
    'Category',
    'Stock',
    'Price',
  ]);
  for (var i = 2; i <= 18; i++) {
    inventorySheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
      TextCellValue('SKU-$i'),
    );
    inventorySheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
      TextCellValue('Item $i'),
    );
    inventorySheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1),
      TextCellValue(['Electronics', 'Office', 'Grocery'][i % 3]),
    );
    inventorySheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i - 1),
      IntCellValue(50 * i),
    );
    inventorySheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: i - 1),
      DoubleCellValue(9.99 * i),
    );
  }
  inventorySheet.frozenRows = 2;
  inventorySheet.frozenColumns = null; // explicit: only rows

  // -------- Sheet 3: Customers — only the first 2 columns frozen -------------
  final customersSheet = excel['Customers'];
  _writeHeader(customersSheet, const [
    'AccountId',
    'RegionId',
    'Name',
    'Email',
    'Tier',
  ]);
  for (var i = 2; i <= 20; i++) {
    customersSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
      TextCellValue('ACC-${1000 + i}'),
    );
    customersSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
      TextCellValue('REG-${(i % 5) + 1}'),
    );
    customersSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1),
      TextCellValue('Customer $i'),
    );
    customersSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i - 1),
      TextCellValue('customer$i@example.com'),
    );
    customersSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: i - 1),
      TextCellValue(['Bronze', 'Silver', 'Gold'][i % 3]),
    );
  }
  customersSheet.frozenRows = null; // explicit: only columns
  customersSheet.frozenColumns = 2;

  // -------- Sheet 4: Logs — no freeze panes at all ---------------------------
  final logsSheet = excel['Logs'];
  _writeHeader(logsSheet, const ['Timestamp', 'Level', 'Message']);
  for (var i = 2; i <= 18; i++) {
    logsSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
      TextCellValue('2026-06-${i.toString().padLeft(2, '0')}T12:00:00Z'),
    );
    logsSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
      TextCellValue(['INFO', 'WARN', 'ERROR'][i % 3]),
    );
    logsSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i - 1),
      TextCellValue('Event #$i happened'),
    );
  }
  logsSheet.frozenRows = null;
  logsSheet.frozenColumns = null;

  // -------- Save ------------------------------------------------------------
  if (kIsWeb) {
    final bytes = excel.save(fileName: 'multi_freeze_panes.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Multi-sheet Freeze Panes generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          '📌 Includes: Sales, Inventory, Customers and Logs sheets, '
          'each with a different freeze configuration.';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    final String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Multi-sheet Freeze Panes Excel',
      fileName: 'multi_freeze_panes.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      return '✅ Multi-sheet Freeze Panes saved successfully!\n'
          'Location: $outputFile';
    }
    return 'Save cancelled.';
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

Future<String> generatePivotTemplateHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Sales Data'];
  excel.delete('Sheet1');

  sheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Region"),
    cellStyle: CellStyle(bold: true),
  );
  sheet.updateCell(
    CellIndex.indexByString("B1"),
    TextCellValue("Amount"),
    cellStyle: CellStyle(bold: true),
  );

  sheet.updateCell(
    CellIndex.indexByString("A2"),
    TextCellValue("North Region"),
  );
  sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(45000));

  sheet.updateCell(
    CellIndex.indexByString("A3"),
    TextCellValue("South Region"),
  );
  sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(32000));

  sheet.updateCell(CellIndex.indexByString("A4"), TextCellValue("East Region"));
  sheet.updateCell(CellIndex.indexByString("B4"), IntCellValue(51000));

  // Create target sheet for Pivot Table
  var pivotSheet = excel['Pivot Report'];

  // Create Pivot Table programmatically
  final pivotTable = PivotTable(
    name: 'PivotTable1',
    sourceSheet: 'Sales Data',
    sourceRange: 'A1:B4',
    targetCell: CellIndex.indexByString('A3'),
    rows: ['Region'],
    values: [
      PivotTableValue(
        field: 'Amount',
        function: PivotValueFunction.sum,
      ),
    ],
  );

  pivotSheet.addPivotTable(pivotTable);

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'sales_report_pivot.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Pivot Table programmatically generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'Download started.';
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
      return '✅ Pivot Table programmatically saved successfully!\n'
          'Location: $outputFile';
    }
    return 'Save cancelled.';
  }
}

Future<String> generateFreezePanesHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Sales Report'];
  excel.delete('Sheet1');

  sheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Product Name"),
    cellStyle: CellStyle(bold: true),
  );
  sheet.updateCell(
    CellIndex.indexByString("B1"),
    TextCellValue("Sales Revenue"),
    cellStyle: CellStyle(bold: true),
  );

  for (int i = 2; i <= 50; i++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1),
      TextCellValue("Product $i"),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1),
      IntCellValue(150 * i),
    );
  }

  // 👇 Freeze the header row and the first column.
  // Setting these properties writes a <pane> element to xl/worksheets/sheet1.xml
  // and is preserved when the file is re-decoded.
  sheet.frozenRows = 1;
  sheet.frozenColumns = 1;

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
