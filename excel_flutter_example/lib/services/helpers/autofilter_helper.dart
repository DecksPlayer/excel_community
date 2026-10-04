import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateAutoFilterHelper() async {
  final excel = Excel.createExcel();
  final sheet = excel['Inventory & Orders'];

  // 1. Setup Styled Table Headers
  final headers = [
    'Order ID',
    'Product Name',
    'Category',
    'Region',
    'Units Sold',
    'Unit Price',
    'Total Revenue',
  ];

  final headerStyle = CellStyle(
    bold: true,
    fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
    backgroundColorHex: ExcelColor.fromHexString('#1E3A8A'), // Deep Indigo
    horizontalAlign: HorizontalAlign.Center,
    verticalAlign: VerticalAlign.Center,
  );

  for (int col = 0; col < headers.length; col++) {
    final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0));
    cell.value = TextCellValue(headers[col]);
    cell.cellStyle = headerStyle;
  }

  // 2. Data Rows
  final records = [
    [1001, 'MacBook Pro 16"', 'Electronics', 'North America', 18, 2499.0, 44982.0],
    [1002, 'Ergonomic Standing Desk', 'Furniture', 'Europe', 45, 680.0, 30600.0],
    [1003, '4K UltraSharp Monitor', 'Electronics', 'Asia-Pacific', 60, 520.0, 31200.0],
    [1004, 'Executive Mesh Chair', 'Furniture', 'North America', 85, 340.0, 28900.0],
    [1005, 'Wireless Noise-Canceling Headphones', 'Electronics', 'Latin America', 120, 199.0, 23880.0],
    [1006, 'Solid Oak Conference Table', 'Furniture', 'Europe', 12, 1250.0, 15000.0],
    [1007, 'Mechanical RGB Keyboard', 'Electronics', 'Asia-Pacific', 150, 129.0, 19350.0],
    [1008, 'Steel Filing Cabinet', 'Furniture', 'North America', 35, 210.0, 7350.0],
  ];

  for (int r = 0; r < records.length; r++) {
    final row = records[r];
    final rowIndex = r + 1;

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: rowIndex)).value =
        IntCellValue(row[0] as int);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: rowIndex)).value =
        TextCellValue(row[1] as String);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: rowIndex)).value =
        TextCellValue(row[2] as String);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: rowIndex)).value =
        TextCellValue(row[3] as String);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: rowIndex)).value =
        IntCellValue(row[4] as int);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: rowIndex)).value =
        DoubleCellValue(row[5] as double);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: rowIndex)).value =
        DoubleCellValue(row[6] as double);
  }

  // Set explicit column widths for readability
  sheet.setColumnWidth(0, 14.0);
  sheet.setColumnWidth(1, 32.0);
  sheet.setColumnWidth(2, 18.0);
  sheet.setColumnWidth(3, 20.0);
  sheet.setColumnWidth(4, 14.0);
  sheet.setColumnWidth(5, 16.0);
  sheet.setColumnWidth(6, 18.0);

  // 3. Enable AutoFilter over headers and data (A1 to G9)
  sheet.setAutoFilter(
    CellIndex.indexByString('A1'),
    CellIndex.indexByString('G9'),
  );

  // 4. Configure column filter on Category (colId: 2, 0-based from column A)
  sheet.addFilterColumn(
    const FilterColumn(
      colId: 2,
      filterValues: ['Electronics'],
    ),
  );

  final summaryInfo = 'Range: ${sheet.autoFilter?.ref} (${sheet.autoFilter?.columnCount} columns × ${sheet.autoFilter?.rowCount} rows)';

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'autofilter_showcase.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ AutoFilter spreadsheet generated successfully!\n'
          'Filter configuration: $summaryInfo\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '\n📥 Check your Downloads folder\n'
          '📌 File: autofilter_showcase.xlsx\n'
          '💡 When opened in Excel, dropdown filter arrows appear on Row 1';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save AutoFilter Showcase',
      fileName: 'autofilter_showcase.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ AutoFilter spreadsheet saved successfully!\n'
          'Location: $outputFile\n'
          'Filter configuration: $summaryInfo\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
          '\n📌 Open with Excel or Google Sheets to test interactive dropdown filters!';
    }
    return 'Save cancelled.';
  }
}
