import 'dart:convert';
import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  String extractChartXml(List<int> bytes, [int chartIndex = 1]) {
    final archive = ZipDecoder().decodeBytes(bytes);
    final chartFile = archive.files
        .firstWhere((f) => f.name == 'xl/charts/chart$chartIndex.xml');
    chartFile.decompress();
    return utf8.decode(chartFile.content);
  }

  group('Chart Grouping (Stacked & 100% Stacked)', () {
    test('ColumnChart generates clustered, stacked, and percentStacked XML', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Cat'));
      sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(10));
      sheet.updateCell(CellIndex.indexByString('C1'), IntCellValue(20));

      final series = [
        ChartSeries(
            name: 'S1',
            categoriesRange: r"Sheet1!$A$1:$A$1",
            valuesRange: r"Sheet1!$B$1:$B$1"),
        ChartSeries(
            name: 'S2',
            categoriesRange: r"Sheet1!$A$1:$A$1",
            valuesRange: r"Sheet1!$C$1:$C$1"),
      ];

      // Stacked
      sheet.addChart(ColumnChart(
        title: 'Stacked Col',
        series: series,
        anchor: ChartAnchor.at(column: 2, row: 2),
        grouping: ChartGrouping.stacked,
      ));

      // 100% Stacked
      sheet.addChart(ColumnChart(
        title: '100% Stacked Col',
        series: series,
        anchor: ChartAnchor.at(column: 10, row: 2),
        grouping: ChartGrouping.percentStacked,
      ));

      final bytes = excel.save()!;
      final xml1 = extractChartXml(bytes, 1);
      final xml2 = extractChartXml(bytes, 2);

      expect(xml1.contains('val="stacked"'), isTrue);
      expect(xml1.contains('<c:overlap val="100"/>'), isTrue);

      expect(xml2.contains('val="percentStacked"'), isTrue);
      expect(xml2.contains('<c:overlap val="100"/>'), isTrue);
    });

    test('BarChart generates stacked grouping and overlap', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Cat'));
      sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(10));

      sheet.addChart(BarChart(
        title: 'Stacked Bar',
        series: [
          ChartSeries(
              name: 'S1',
              categoriesRange: r"Sheet1!$A$1:$A$1",
              valuesRange: r"Sheet1!$B$1:$B$1")
        ],
        anchor: ChartAnchor.at(column: 2, row: 2),
        grouping: ChartGrouping.stacked,
      ));

      final bytes = excel.save()!;
      final xml = extractChartXml(bytes, 1);
      expect(xml.contains('val="stacked"'), isTrue);
      expect(xml.contains('<c:overlap val="100"/>'), isTrue);
    });

    test('AreaChart generates stacked grouping', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Cat'));
      sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(10));

      sheet.addChart(AreaChart(
        title: 'Stacked Area',
        series: [
          ChartSeries(
              name: 'S1',
              categoriesRange: r"Sheet1!$A$1:$A$1",
              valuesRange: r"Sheet1!$B$1:$B$1")
        ],
        anchor: ChartAnchor.at(column: 2, row: 2),
        grouping: ChartGrouping.percentStacked,
      ));

      final bytes = excel.save()!;
      final xml = extractChartXml(bytes, 1);
      expect(xml.contains('val="percentStacked"'), isTrue);
    });
  });

  group('Line & Scatter Variations (Smooth, Lines, Markers)', () {
    test('LineChart with smooth curves and no markers', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Cat'));
      sheet.updateCell(CellIndex.indexByString('B1'), IntCellValue(10));

      sheet.addChart(LineChart(
        title: 'Smooth Line No Markers',
        series: [
          ChartSeries(
              name: 'S1',
              categoriesRange: r"Sheet1!$A$1:$A$1",
              valuesRange: r"Sheet1!$B$1:$B$1")
        ],
        anchor: ChartAnchor.at(column: 2, row: 2),
        smooth: true,
        showMarkers: false,
      ));

      final bytes = excel.save()!;
      final xml = extractChartXml(bytes, 1);

      expect(xml.contains('<c:smooth val="1"/>'), isTrue);
      expect(xml.contains('<c:symbol val="none"/>'), isTrue);
    });

    test('ScatterChart with connecting lines and smooth curves', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.updateCell(CellIndex.indexByString('A1'), DoubleCellValue(1.0));
      sheet.updateCell(CellIndex.indexByString('B1'), DoubleCellValue(5.0));

      sheet.addChart(ScatterChart(
        title: 'Smooth Scatter with Lines',
        series: [
          ChartSeries(
              name: 'S1',
              categoriesRange: r"Sheet1!$A$1:$A$1",
              valuesRange: r"Sheet1!$B$1:$B$1")
        ],
        anchor: ChartAnchor.at(column: 2, row: 2),
        showLines: true,
        smooth: true,
        showMarkers: true,
      ));

      final bytes = excel.save()!;
      final xml = extractChartXml(bytes, 1);

      expect(xml.contains('scatterStyle val="lineMarker"'), isTrue);
      expect(xml.contains('<c:smooth val="1"/>'), isTrue);
      expect(xml.contains('<c:symbol val="circle"/>'), isTrue);
    });
  });

  group('BubbleChart', () {
    test('creates BubbleChart with 3D data (X, Y, Size) and valid XML', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      // Col A: X, Col B: Y, Col C: Size
      sheet.updateCell(CellIndex.indexByString('A1'), DoubleCellValue(1.5));
      sheet.updateCell(CellIndex.indexByString('B1'), DoubleCellValue(10.0));
      sheet.updateCell(CellIndex.indexByString('C1'), DoubleCellValue(25.0));

      sheet.addChart(BubbleChart(
        title: 'Risk vs Return',
        series: [
          ChartSeries(
            name: 'Project A',
            categoriesRange: r"Sheet1!$A$1:$A$1",
            valuesRange: r"Sheet1!$B$1:$B$1",
            bubbleSizeRange: r"Sheet1!$C$1:$C$1",
          ),
        ],
        anchor: ChartAnchor.at(column: 4, row: 1),
        bubbleScale: 120,
      ));

      final bytes = excel.save()!;
      expect(bytes, isNotNull);

      final xml = extractChartXml(bytes, 1);
      expect(xml.contains('<c:bubbleChart>'), isTrue);
      expect(xml.contains('<c:bubbleScale val="120"/>'), isTrue);
      expect(xml.contains('<c:xVal>'), isTrue);
      expect(xml.contains('<c:yVal>'), isTrue);
      expect(xml.contains('<c:bubbleSize>'), isTrue);

      // Verify decode
      final decoded = Excel.decodeBytes(bytes);
      expect(decoded.sheets.containsKey('Sheet1'), isTrue);
    });
  });

  group('StockChart', () {
    test('creates StockChart with High-Low lines and Up-Down bars', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Day 1'));
      sheet.updateCell(CellIndex.indexByString('B1'), DoubleCellValue(105.0)); // High
      sheet.updateCell(CellIndex.indexByString('C1'), DoubleCellValue(95.0));  // Low
      sheet.updateCell(CellIndex.indexByString('D1'), DoubleCellValue(100.0)); // Close

      sheet.addChart(StockChart(
        title: 'Stock HLC',
        series: [
          ChartSeries(name: 'High', categoriesRange: r"Sheet1!$A$1:$A$1", valuesRange: r"Sheet1!$B$1:$B$1"),
          ChartSeries(name: 'Low', categoriesRange: r"Sheet1!$A$1:$A$1", valuesRange: r"Sheet1!$C$1:$C$1"),
          ChartSeries(name: 'Close', categoriesRange: r"Sheet1!$A$1:$A$1", valuesRange: r"Sheet1!$D$1:$D$1"),
        ],
        anchor: ChartAnchor.at(column: 5, row: 1),
        showHighLowLines: true,
        showUpDownBars: true,
      ));

      final bytes = excel.save()!;
      final xml = extractChartXml(bytes, 1);

      expect(xml.contains('<c:stockChart>'), isTrue);
      expect(xml.contains('<c:hiLowLines/>'), isTrue);
      expect(xml.contains('<c:upDownBars>'), isTrue);
    });
  });

  group('OfPieChart', () {
    test('creates Pie-of-Pie and Bar-of-Pie with secondary breakdown', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];

      final categories = ['Cat 1', 'Cat 2', 'Cat 3', 'Cat 4'];
      final values = [50, 30, 12, 8];
      for (var i = 0; i < categories.length; i++) {
        sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i), TextCellValue(categories[i]));
        sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i), IntCellValue(values[i]));
      }

      // 1. Pie-of-Pie
      sheet.addChart(OfPieChart(
        title: 'Pie of Pie',
        series: [
          ChartSeries(name: 'Share', categoriesRange: r"Sheet1!$A$1:$A$4", valuesRange: r"Sheet1!$B$1:$B$4"),
        ],
        anchor: ChartAnchor.at(column: 4, row: 1),
        ofPieType: OfPieType.pie,
        splitPosition: 2,
      ));

      // 2. Bar-of-Pie
      sheet.addChart(OfPieChart(
        title: 'Bar of Pie',
        series: [
          ChartSeries(name: 'Share', categoriesRange: r"Sheet1!$A$1:$A$4", valuesRange: r"Sheet1!$B$1:$B$4"),
        ],
        anchor: ChartAnchor.at(column: 12, row: 1),
        ofPieType: OfPieType.bar,
        splitPosition: 2,
      ));

      final bytes = excel.save()!;
      final xml1 = extractChartXml(bytes, 1);
      final xml2 = extractChartXml(bytes, 2);

      expect(xml1.contains('<c:ofPieChart>'), isTrue);
      expect(xml1.contains('ofPieType val="pie"'), isTrue);
      expect(xml1.contains('<c:splitPos val="2"/>'), isTrue);

      expect(xml2.contains('<c:ofPieChart>'), isTrue);
      expect(xml2.contains('ofPieType val="bar"'), isTrue);
    });
  });

  test('line and area charts write "standard" instead of "clustered" grouping', () {
    String groupingOf(Chart chart) {
      final excel = Excel.createExcel();
      excel['Sheet1'].addChart(chart);
      final file = ZipDecoder().decodeBytes(excel.encode()!).findFile('xl/charts/chart1.xml')!;
      return RegExp(r'<c:grouping val="(\w+)"/>').firstMatch(utf8.decode(file.content))!.group(1)!;
    }

    final series = [
      ChartSeries(name: 'S', categoriesRange: r'Sheet1!$A$1:$A$2', valuesRange: r'Sheet1!$B$1:$B$2'),
    ];
    final anchor = ChartAnchor.at(column: 3, row: 1);
    expect(groupingOf(LineChart(title: 'L', series: series, anchor: anchor)), 'standard');
    expect(groupingOf(AreaChart(title: 'A', series: series, anchor: anchor)), 'standard');
    expect(
        groupingOf(LineChart(title: 'L', series: series, anchor: anchor, grouping: ChartGrouping.stacked)),
        'stacked');
    expect(groupingOf(ColumnChart(title: 'C', series: series, anchor: anchor)), 'clustered');
  });
}
