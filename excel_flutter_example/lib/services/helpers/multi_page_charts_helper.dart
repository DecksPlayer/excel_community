import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

/// Generates an Excel workbook with five sheets, each containing a different
/// chart type. This demonstrates that charts render correctly on non-default
/// sheets.
Future<String> generateMultiPageChartsHelper() async {
  final excel = Excel.createExcel();

  // ── Sheet 1: Sales — Column Chart ────────────────────────────────────────
  final salesSheet = excel['Sales'];
  _writeHeaders(salesSheet, ['Month', 'Revenue', 'Expenses']);
  final salesData = [
    ['Jan', 42000, 28000],
    ['Feb', 55000, 31000],
    ['Mar', 48000, 27000],
    ['Apr', 63000, 35000],
    ['May', 71000, 38000],
    ['Jun', 59000, 33000],
  ];
  _writeRows(salesSheet, salesData);

  salesSheet.addChart(ColumnChart(
    title: 'Monthly Revenue vs Expenses',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"Sales!$A$2:$A$7",
        valuesRange: r"Sales!$B$2:$B$7",
      ),
      ChartSeries(
        name: 'Expenses',
        categoriesRange: r"Sales!$A$2:$A$7",
        valuesRange: r"Sales!$C$2:$C$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  ));

  // ── Sheet 2: Market Share — Pie Chart ────────────────────────────────────
  final marketSheet = excel['Market Share'];
  _writeHeaders(marketSheet, ['Region', 'Share (%)']);
  final marketData = [
    ['North America', 38],
    ['Europe', 27],
    ['Asia Pacific', 21],
    ['Latin America', 9],
    ['Rest of World', 5],
  ];
  _writeRows(marketSheet, marketData);

  marketSheet.addChart(PieChart(
    title: 'Global Market Share Distribution',
    series: [
      ChartSeries(
        name: 'Market Share',
        categoriesRange: r"'Market Share'!$A$2:$A$6",
        valuesRange: r"'Market Share'!$B$2:$B$6",
      ),
    ],
    anchor: ChartAnchor.at(column: 4, row: 1, width: 10, height: 16),
  ));

  // ── Sheet 3: Trends — Line Chart ─────────────────────────────────────────
  final trendsSheet = excel['Trends'];
  _writeHeaders(trendsSheet, ['Quarter', 'Product A', 'Product B', 'Product C']);
  final trendsData = [
    ['Q1 2023', 120, 95, 60],
    ['Q2 2023', 145, 110, 78],
    ['Q3 2023', 132, 125, 95],
    ['Q4 2023', 178, 140, 112],
    ['Q1 2024', 195, 158, 130],
    ['Q2 2024', 220, 172, 148],
  ];
  _writeRows(trendsSheet, trendsData);

  trendsSheet.addChart(LineChart(
    title: 'Product Sales Trend by Quarter',
    series: [
      ChartSeries(
        name: 'Product A',
        categoriesRange: r"Trends!$A$2:$A$7",
        valuesRange: r"Trends!$B$2:$B$7",
      ),
      ChartSeries(
        name: 'Product B',
        categoriesRange: r"Trends!$A$2:$A$7",
        valuesRange: r"Trends!$C$2:$C$7",
      ),
      ChartSeries(
        name: 'Product C',
        categoriesRange: r"Trends!$A$2:$A$7",
        valuesRange: r"Trends!$D$2:$D$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 6, row: 1, width: 11, height: 15),
  ));

  // ── Sheet 4: Performance — Radar Chart ───────────────────────────────────
  final perfSheet = excel['Performance'];
  _writeHeaders(perfSheet, ['Metric', 'Team Alpha', 'Team Beta']);
  final perfData = [
    ['Speed', 85, 72],
    ['Quality', 90, 88],
    ['Efficiency', 78, 82],
    ['Innovation', 95, 70],
    ['Support', 80, 91],
    ['Reliability', 88, 85],
  ];
  _writeRows(perfSheet, perfData);

  perfSheet.addChart(RadarChart(
    title: 'Team Performance Comparison',
    filled: true,
    series: [
      ChartSeries(
        name: 'Team Alpha',
        categoriesRange: r"Performance!$A$2:$A$7",
        valuesRange: r"Performance!$B$2:$B$7",
      ),
      ChartSeries(
        name: 'Team Beta',
        categoriesRange: r"Performance!$A$2:$A$7",
        valuesRange: r"Performance!$C$2:$C$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 4, row: 1, width: 10, height: 15),
  ));

  // ── Sheet 5: Correlation — Scatter Chart ─────────────────────────────────
  final corrSheet = excel['Correlation'];
  _writeHeaders(corrSheet, ['Ad Spend (k)', 'Revenue (k)', 'Leads']);
  final corrData = [
    [1.0, 12.0, 8.0],
    [2.5, 22.5, 15.0],
    [4.0, 31.0, 24.0],
    [6.0, 44.0, 35.0],
    [8.0, 58.0, 47.0],
    [10.0, 75.0, 62.0],
    [12.5, 95.0, 78.0],
  ];
  for (var i = 0; i < corrData.length; i++) {
    corrSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
        DoubleCellValue(corrData[i][0]));
    corrSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
        DoubleCellValue(corrData[i][1]));
    corrSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1),
        DoubleCellValue(corrData[i][2]));
  }

  corrSheet.addChart(ScatterChart(
    title: 'Ad Spend vs Revenue Correlation',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"Correlation!$A$2:$A$8",
        valuesRange: r"Correlation!$B$2:$B$8",
      ),
      ChartSeries(
        name: 'Leads',
        categoriesRange: r"Correlation!$A$2:$A$8",
        valuesRange: r"Correlation!$C$2:$C$8",
      ),
    ],
    anchor: ChartAnchor.at(column: 4, row: 1, width: 11, height: 15),
  ));

  // Delete the default empty Sheet1 that Excel.createExcel() creates when no
  // sheets named 'Sheet1' were explicitly used above.
  if (excel.sheets.containsKey('Sheet1') &&
      !['Sales', 'Market Share', 'Trends', 'Performance', 'Correlation']
          .contains('Sheet1')) {
    excel.delete('Sheet1');
  }

  // ── Save ─────────────────────────────────────────────────────────────────
  const fileName = 'multi_page_charts.xlsx';

  if (kIsWeb) {
    final bytes = excel.save(fileName: fileName);
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Multi-page charts workbook generated!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '\n📊 5 sheets: Sales (Column), Market Share (Pie),\n'
          '   Trends (Line), Performance (Radar), Correlation (Scatter)';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) throw Exception('Failed to encode Excel file.');

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Multi-Page Charts Excel',
      fileName: fileName,
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      await File(outputFile).writeAsBytes(bytes);
      final size = await File(outputFile).length();
      return '✅ Multi-page charts workbook saved!\n'
          'Location: $outputFile\n'
          'Size: ${(size / 1024).toStringAsFixed(2)} KB\n'
          '\n📊 5 sheets: Sales (Column), Market Share (Pie),\n'
          '   Trends (Line), Performance (Radar), Correlation (Scatter)';
    }
    return 'Save cancelled.';
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

void _writeHeaders(Sheet sheet, List<String> headers) {
  for (var i = 0; i < headers.length; i++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0),
      TextCellValue(headers[i]),
    );
  }
}

void _writeRows(Sheet sheet, List<List<dynamic>> rows) {
  for (var r = 0; r < rows.length; r++) {
    for (var c = 0; c < rows[r].length; c++) {
      final val = rows[r][c];
      CellValue cellValue;
      if (val is int) {
        cellValue = IntCellValue(val);
      } else if (val is double) {
        cellValue = DoubleCellValue(val);
      } else {
        cellValue = TextCellValue(val.toString());
      }
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1),
        cellValue,
      );
    }
  }
}
