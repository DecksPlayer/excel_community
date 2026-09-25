import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateTabColorHelper() async {
  final excel = Excel.createExcel();

  // 1. Sheet 1: Executive Dashboard (Royal Blue)
  final execSheet = excel['Executive Dashboard'];
  execSheet.setTabColorHex('#2563EB');
  _buildSheetContent(
    execSheet,
    title: 'Executive KPI Dashboard',
    themeColor: '#2563EB',
    headers: ['KPI Metric', 'Target', 'Actual', 'Variance', 'Status'],
    rows: [
      ['Gross ARR', '\$12.5M', '\$13.2M', '+\$700K', 'Exceeded'],
      ['Net Retention Rate', '115%', '118%', '+3%', 'Healthy'],
      ['CAC Payback', '12 mos', '10.5 mos', '-1.5 mos', 'Optimized'],
      ['Monthly Active Users', '850K', '920K', '+70K', 'Surpassed'],
      ['Customer Satisfaction (CSAT)', '92%', '94.5%', '+2.5%', 'Top Tier'],
    ],
  );

  // 2. Sheet 2: Marketing & Growth (Emerald Green)
  final mktSheet = excel['Marketing & Growth'];
  mktSheet.setTabColorHex('#10B981');
  _buildSheetContent(
    mktSheet,
    title: 'Growth & Acquisition Channels',
    themeColor: '#10B981',
    headers: ['Channel', 'Spend', 'Leads Generated', 'MQL-to-SQL', 'CPA'],
    rows: [
      ['Organic Search (SEO)', '\$15,000', '4,200', '28%', '\$3.57'],
      ['Paid Search (SEM)', '\$45,000', '3,100', '22%', '\$14.51'],
      ['Social Campaigns', '\$25,000', '2,800', '19%', '\$8.92'],
      ['Developer Community', '\$8,000', '5,600', '35%', '\$1.43'],
      ['Events & Webinars', '\$12,000', '1,400', '42%', '\$8.57'],
    ],
  );

  // 3. Sheet 3: Financial Performance (Amber Gold)
  final finSheet = excel['Financial Performance'];
  finSheet.setTabColorHex('#F59E0B');
  _buildSheetContent(
    finSheet,
    title: 'Financial Statements Summary (Q3)',
    themeColor: '#F59E0B',
    headers: ['Category', 'Q1 Budget', 'Q2 Actual', 'Q3 Projected', 'YoY Growth'],
    rows: [
      ['Operating Revenue', '\$3,200,000', '\$3,450,000', '\$3,800,000', '+22%'],
      ['Cost of Goods Sold (COGS)', '\$640,000', '\$680,000', '\$720,000', '+12%'],
      ['Gross Profit', '\$2,560,000', '\$2,770,000', '\$3,080,000', '+25%'],
      ['Operating Expenses (OPEX)', '\$1,450,000', '\$1,520,000', '\$1,610,000', '+11%'],
      ['Operating Income (EBITDA)', '\$1,110,000', '\$1,250,000', '\$1,470,000', '+32%'],
    ],
  );

  // 4. Sheet 4: Operations & Logistics (Pink / Magenta)
  final opsSheet = excel['Operations & Logistics'];
  opsSheet.setTabColorHex('#EC4899');
  _buildSheetContent(
    opsSheet,
    title: 'Global Supply Chain & Fulfillment',
    themeColor: '#EC4899',
    headers: ['Hub Location', 'Inventory Capacity', 'Inbound Rate', 'On-Time Delivery', 'Incidents'],
    rows: [
      ['North America (Chicago)', '94%', '12,400 / day', '99.2%', '0'],
      ['Europe (Frankfurt)', '88%', '8,900 / day', '98.7%', '1'],
      ['Asia-Pacific (Singapore)', '91%', '15,200 / day', '99.5%', '0'],
      ['Latin America (São Paulo)', '79%', '4,500 / day', '97.8%', '2'],
    ],
  );

  // 5. Sheet 5: Standard Tab (Default color)
  final notesSheet = excel['Notes & Documentation'];
  _buildSheetContent(
    notesSheet,
    title: 'Workbook Information & Release Notes',
    themeColor: '#475569',
    headers: ['Section', 'Details', 'Author', 'Last Updated'],
    rows: [
      ['Spreadsheet Standard', 'OpenXML SpreadsheetML ECMA-376', 'excel_community team', '2026-09-25'],
      ['Tab Color Feature', '<sheetPr><tabColor rgb="AARRGGBB"/></sheetPr>', 'Engineering', 'Version 2.4.2'],
      ['Compatibility', 'Excel, LibreOffice Calc, Apple Numbers, Google Sheets', 'QA Team', 'Verified'],
    ],
  );

  final summaryInfo = '5 Sheets (Executive: #2563EB, Marketing: #10B981, Finance: #F59E0B, Ops: #EC4899)';

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'colored_sheets_portfolio.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Tab Colors spreadsheet generated successfully!\n'
          'Worksheets: $summaryInfo\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '\n📥 Check your Downloads folder\n'
          '📌 File: colored_sheets_portfolio.xlsx\n'
          '💡 Notice the colorful tab strips at the bottom of Excel!';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Tab Colors Showcase',
      fileName: 'colored_sheets_portfolio.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Tab Colors spreadsheet saved successfully!\n'
          'Location: $outputFile\n'
          'Worksheets: $summaryInfo\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
          '\n📌 Open with Microsoft Excel to admire the custom sheet tab colors!';
    }
    return 'Save cancelled.';
  }
}

void _buildSheetContent(
  Sheet sheet, {
  required String title,
  required String themeColor,
  required List<String> headers,
  required List<List<String>> rows,
}) {
  // Title row
  sheet.cell(CellIndex.indexByString('A1')).value = TextCellValue(title);
  sheet.cell(CellIndex.indexByString('A1')).cellStyle = CellStyle(
    bold: true,
    fontSize: 14,
    fontColorHex: ExcelColor.fromHexString(themeColor),
  );

  // Header row
  for (int col = 0; col < headers.length; col++) {
    final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 2));
    cell.value = TextCellValue(headers[col]);
    cell.cellStyle = CellStyle(
      bold: true,
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
      backgroundColorHex: ExcelColor.fromHexString(themeColor),
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    );
    sheet.setColumnWidth(col, 22.0);
  }

  // Data rows
  for (int r = 0; r < rows.length; r++) {
    final row = rows[r];
    for (int c = 0; c < row.length; c++) {
      final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 3));
      cell.value = TextCellValue(row[c]);
      cell.cellStyle = CellStyle(
        horizontalAlign: HorizontalAlign.Left,
      );
    }
  }
}
