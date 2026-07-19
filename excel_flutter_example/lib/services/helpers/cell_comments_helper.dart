import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateCellCommentsHelper() async {
  final excel = Excel.createExcel();
  final sheet = excel['Cell Comments'];

  if (excel.sheets.containsKey('Sheet1')) {
    excel.delete('Sheet1');
  }

  // 1. Title Block
  sheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Sales Inventory with Cell Comments"),
    cellStyle: CellStyle(bold: true, fontSize: 14, fontColorHex: ExcelColor.teal),
  );
  sheet.updateCell(
    CellIndex.indexByString("A2"),
    TextCellValue("Hover or select cells with the red corner triangle in Excel to read comments."),
    cellStyle: CellStyle(italic: true, fontSize: 10),
  );

  // 2. Headers
  final headers = ["Product Code", "Product Name", "Units In Stock", "Unit Price", "Status"];
  final headerStyle = CellStyle(
    bold: true,
    fontSize: 11,
    fontColorHex: ExcelColor.white,
    backgroundColorHex: ExcelColor.teal,
    horizontalAlign: HorizontalAlign.Center,
  );

  for (var i = 0; i < headers.length; i++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 3),
      TextCellValue(headers[i]),
      cellStyle: headerStyle,
    );
  }

  // 3. Data Rows
  final products = [
    ["P1001", "Widget A", 120, 15.50, "Active"],
    ["P1002", "Gadget B", 45, 89.99, "Restocking"],
    ["P1003", "Device C", 8, 249.00, "Low Stock"],
    ["P1004", "Widget D", 200, 5.25, "Discontinuing"],
  ];

  for (var r = 0; r < products.length; r++) {
    final rowIdx = 4 + r;
    final product = products[r];
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: rowIdx), TextCellValue(product[0] as String));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: rowIdx), TextCellValue(product[1] as String));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: rowIdx), IntCellValue(product[2] as int));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: rowIdx), DoubleCellValue(product[3] as double));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: rowIdx), TextCellValue(product[4] as String));
  }

  // 4. Attach comments to specific cells
  // Comment on Product Name of Widget A (B5)
  sheet.cell(CellIndex.indexByString("B5")).comment = "Top selling item this quarter. Restocked weekly.";

  // Comment on Restocking status of Gadget B (E6)
  sheet.cell(CellIndex.indexByString("E6")).comment = "Shipment delayed due to customs. Expected delivery next Friday.";

  // Comment on Low Stock quantity of Device C (C7)
  sheet.cell(CellIndex.indexByString("C7")).comment = "Reorder point reached! Minimum threshold is 10 units.";

  // Comment on Discontinuing status of Widget D (E8)
  sheet.cell(CellIndex.indexByString("E8")).comment = "Replaced by Widget E starting next fiscal period.";

  // Set column widths
  sheet.setColumnWidth(0, 15);
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 16);
  sheet.setColumnWidth(3, 14);
  sheet.setColumnWidth(4, 16);

  // 5. Save
  const fileName = 'cell_comments_report.xlsx';

  if (kIsWeb) {
    final bytes = excel.save(fileName: fileName);
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Cell Comments Excel generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'Open this file in Excel to view comments on cells B5, E6, C7 and E8.';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) throw Exception('Failed to encode Excel file.');

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Cell Comments Excel',
      fileName: fileName,
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      await File(outputFile).writeAsBytes(bytes);
      final size = await File(outputFile).length();
      return '✅ Cell Comments Excel saved successfully!\n'
          'Location: $outputFile\n'
          'Size: ${(size / 1024).toStringAsFixed(2)} KB\n'
          'Open this file in Excel to view comments on cells B5, E6, C7 and E8.';
    }
    return 'Save cancelled.';
  }
}
