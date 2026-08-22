// Advanced Chart types: Bubble, Stock, OfPie, and Stacked variations
library;

const String newChartsSnippet = r'''
import 'package:excel_community/excel_community.dart';

/// Demonstrates new chart types and grouping options:
/// 1. BubbleChart (3D: X, Y, Size)
/// 2. StockChart (High-Low-Close with High-Low lines & Up-Down bars)
/// 3. OfPieChart (Pie-of-Pie secondary breakdown)
/// 4. Stacked Column Chart (ChartGrouping.stacked & percentStacked)
void generateAdvancedCharts() {
  final excel = Excel.createExcel();

  // ── 1. Bubble Chart (Risk vs Return vs Market Cap) ────────────────────────
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

  // ── 2. Stock Chart (High - Low - Close) ───────────────────────────────────
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

  // ── 3. OfPie Chart (Pie of Pie) ──────────────────────────────────────────
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
    splitPosition: 2, // Last 2 items split to secondary pie
  ));

  // ── 4. Stacked Column Chart ──────────────────────────────────────────────
  final stackedSheet = excel['Stacked Columns'];
  // ... populate series data ...
  stackedSheet.addChart(ColumnChart(
    title: 'Stacked Revenue & Cost',
    series: [
      ChartSeries(name: 'Revenue', categoriesRange: r"'Stacked Columns'!$A$2:$A$6", valuesRange: r"'Stacked Columns'!$B$2:$B$6"),
      ChartSeries(name: 'Cost', categoriesRange: r"'Stacked Columns'!$A$2:$A$6", valuesRange: r"'Stacked Columns'!$C$2:$C$6"),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
    grouping: ChartGrouping.stacked,
  ));

  excel.delete('Sheet1');
  excel.save(fileName: 'advanced_charts_demo.xlsx');
}
''';
