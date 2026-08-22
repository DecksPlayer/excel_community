import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateNewChartsHelper() async {
  final excel = Excel.createExcel();

  // ── 1. Bubble Chart ───────────────────────────────────────────────────────
  final bubbleSheet = excel['Bubble Chart'];
  bubbleSheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Risk (X)'));
  bubbleSheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Return (Y)'));
  bubbleSheet.updateCell(CellIndex.indexByString('C1'), TextCellValue('Market Cap (Size)'));

  final bubbleData = [
    [2.5, 12.0, 45.0],
    [4.0, 18.0, 70.0],
    [5.5, 25.0, 30.0],
    [7.0, 30.0, 95.0],
    [8.5, 45.0, 60.0],
  ];
  for (var i = 0; i < bubbleData.length; i++) {
    bubbleSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), DoubleCellValue(bubbleData[i][0]));
    bubbleSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), DoubleCellValue(bubbleData[i][1]));
    bubbleSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), DoubleCellValue(bubbleData[i][2]));
  }

  bubbleSheet.addChart(BubbleChart(
    title: 'Risk vs Return vs Market Cap',
    series: [
      ChartSeries(
        name: 'Equities',
        categoriesRange: r"'Bubble Chart'!$A$2:$A$6",
        valuesRange: r"'Bubble Chart'!$B$2:$B$6",
        bubbleSizeRange: r"'Bubble Chart'!$C$2:$C$6",
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
    bubbleScale: 100,
  ));

  // ── 2. Stock Chart ────────────────────────────────────────────────────────
  final stockSheet = excel['Stock Chart'];
  stockSheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Day'));
  stockSheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('High'));
  stockSheet.updateCell(CellIndex.indexByString('C1'), TextCellValue('Low'));
  stockSheet.updateCell(CellIndex.indexByString('D1'), TextCellValue('Close'));

  final stockData = [
    ['Mon', 152.5, 148.0, 151.2],
    ['Tue', 155.0, 150.5, 153.8],
    ['Wed', 154.0, 147.2, 148.5],
    ['Thu', 158.0, 152.0, 157.0],
    ['Fri', 162.5, 156.0, 161.0],
  ];
  for (var i = 0; i < stockData.length; i++) {
    stockSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(stockData[i][0] as String));
    stockSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), DoubleCellValue(stockData[i][1] as double));
    stockSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), DoubleCellValue(stockData[i][2] as double));
    stockSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i + 1), DoubleCellValue(stockData[i][3] as double));
  }

  stockSheet.addChart(StockChart(
    title: 'Weekly Stock Performance (HLC)',
    series: [
      ChartSeries(name: 'High', categoriesRange: r"'Stock Chart'!$A$2:$A$6", valuesRange: r"'Stock Chart'!$B$2:$B$6"),
      ChartSeries(name: 'Low', categoriesRange: r"'Stock Chart'!$A$2:$A$6", valuesRange: r"'Stock Chart'!$C$2:$C$6"),
      ChartSeries(name: 'Close', categoriesRange: r"'Stock Chart'!$A$2:$A$6", valuesRange: r"'Stock Chart'!$D$2:$D$6"),
    ],
    anchor: ChartAnchor.at(column: 6, row: 1, width: 11, height: 15),
    showHighLowLines: true,
    showUpDownBars: true,
  ));

  // ── 3. OfPie Chart ────────────────────────────────────────────────────────
  final ofPieSheet = excel['Pie of Pie'];
  final pieCategories = ['Product A', 'Product B', 'Product C', 'Minor X', 'Minor Y'];
  final pieValues = [45, 30, 15, 6, 4];
  ofPieSheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Product'));
  ofPieSheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Sales'));
  for (var i = 0; i < pieCategories.length; i++) {
    ofPieSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(pieCategories[i]));
    ofPieSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), IntCellValue(pieValues[i]));
  }

  ofPieSheet.addChart(OfPieChart(
    title: 'Sales Breakdown (Pie of Pie)',
    series: [
      ChartSeries(name: 'Sales', categoriesRange: r"'Pie of Pie'!$A$2:$A$6", valuesRange: r"'Pie of Pie'!$B$2:$B$6"),
    ],
    anchor: ChartAnchor.at(column: 4, row: 1, width: 11, height: 15),
    ofPieType: OfPieType.pie,
    splitPosition: 2,
  ));

  // ── 4. Stacked Column Chart ───────────────────────────────────────────────
  final stackedSheet = excel['Stacked Columns'];
  stackedSheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Month'));
  stackedSheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Product 1'));
  stackedSheet.updateCell(CellIndex.indexByString('C1'), TextCellValue('Product 2'));

  final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  final p1 = [40, 55, 35, 70, 60, 80];
  final p2 = [25, 30, 20, 35, 40, 45];
  for (var i = 0; i < months.length; i++) {
    stackedSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(months[i]));
    stackedSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), IntCellValue(p1[i]));
    stackedSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), IntCellValue(p2[i]));
  }

  stackedSheet.addChart(ColumnChart(
    title: 'Stacked Monthly Product Sales',
    series: [
      ChartSeries(name: 'Product 1', categoriesRange: r"'Stacked Columns'!$A$2:$A$7", valuesRange: r"'Stacked Columns'!$B$2:$B$7"),
      ChartSeries(name: 'Product 2', categoriesRange: r"'Stacked Columns'!$A$2:$A$7", valuesRange: r"'Stacked Columns'!$C$2:$C$7"),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
    grouping: ChartGrouping.stacked,
  ));

  excel.delete('Sheet1');

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'advanced_charts_demo.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Advanced Charts generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          '\n4 sheets:\n'
          '• Bubble Chart (3D X/Y/Size)\n'
          '• Stock Chart (High/Low/Close)\n'
          '• Pie of Pie (Sub-breakdown)\n'
          '• Stacked Columns (Cumulative totals)';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) throw Exception('Failed to encode Excel file.');

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Advanced Charts Demo',
      fileName: 'advanced_charts_demo.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      await File(outputFile).writeAsBytes(bytes);
      return '✅ Saved successfully!\n'
          'Location: $outputFile\n'
          '\n4 sheets:\n'
          '• Bubble Chart (3D X/Y/Size)\n'
          '• Stock Chart (High/Low/Close)\n'
          '• Pie of Pie (Sub-breakdown)\n'
          '• Stacked Columns (Cumulative totals)';
    }
    return 'Save cancelled.';
  }
}
