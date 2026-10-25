import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

/// ECMA-376 child sequences of the chart elements the library writes.
const _chartSequences = {
  'barChart': ['barDir', 'grouping', 'varyColors', 'ser', 'dLbls', 'gapWidth', 'overlap', 'serLines', 'axId'],
  'lineChart': ['grouping', 'varyColors', 'ser', 'dLbls', 'dropLines', 'hiLowLines', 'upDownBars', 'marker', 'smooth', 'axId'],
  'areaChart': ['grouping', 'varyColors', 'ser', 'dLbls', 'dropLines', 'axId'],
  'pieChart': ['varyColors', 'ser', 'dLbls', 'firstSliceAng'],
  'doughnutChart': ['varyColors', 'ser', 'dLbls', 'firstSliceAng', 'holeSize'],
  'ofPieChart': ['ofPieType', 'varyColors', 'ser', 'dLbls', 'gapWidth', 'splitType', 'splitPos', 'custSplit', 'secondPieSize', 'serLines'],
  'scatterChart': ['scatterStyle', 'varyColors', 'ser', 'dLbls', 'axId'],
  'bubbleChart': ['varyColors', 'ser', 'dLbls', 'bubble3D', 'bubbleScale', 'showNegBubbles', 'sizeRepresents', 'axId'],
  'radarChart': ['radarStyle', 'varyColors', 'ser', 'dLbls', 'axId'],
  'stockChart': ['ser', 'dLbls', 'dropLines', 'hiLowLines', 'upDownBars', 'axId'],
  'dLbls': ['dLbl', 'delete', 'numFmt', 'spPr', 'txPr', 'dLblPos', 'showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize', 'separator', 'showLeaderLines', 'leaderLines'],
};

/// CT_*Ser sequences, by the chart element that holds the series.
const _seriesSequences = {
  'barChart': ['idx', 'order', 'tx', 'spPr', 'invertIfNegative', 'pictureOptions', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'val', 'shape'],
  'lineChart': ['idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'val', 'smooth'],
  'stockChart': ['idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'val', 'smooth'],
  'areaChart': ['idx', 'order', 'tx', 'spPr', 'pictureOptions', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'val'],
  'pieChart': ['idx', 'order', 'tx', 'spPr', 'explosion', 'dPt', 'dLbls', 'cat', 'val'],
  'doughnutChart': ['idx', 'order', 'tx', 'spPr', 'explosion', 'dPt', 'dLbls', 'cat', 'val'],
  'ofPieChart': ['idx', 'order', 'tx', 'spPr', 'explosion', 'dPt', 'dLbls', 'cat', 'val'],
  'scatterChart': ['idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'trendline', 'errBars', 'xVal', 'yVal', 'smooth'],
  'bubbleChart': ['idx', 'order', 'tx', 'spPr', 'invertIfNegative', 'dPt', 'dLbls', 'trendline', 'errBars', 'xVal', 'yVal', 'bubbleSize', 'bubble3D'],
  'radarChart': ['idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'cat', 'val'],
};

/// Problems with the order of [element]'s children, for [sequence].
List<String> _orderProblems(XmlElement element, List<String> sequence) {
  final problems = <String>[];
  var last = -1;
  for (final child in element.childElements) {
    final i = sequence.indexOf(child.name.local);
    if (i < 0) {
      problems.add('<${element.name.local}> has unexpected <${child.name.local}>');
    } else if (i < last) {
      problems.add('<${child.name.local}> is out of order in <${element.name.local}>');
    } else {
      last = i;
    }
  }
  return problems;
}

void main() {
  test('every chart type is written in ECMA-376 element order', () {
    final excel = Excel.createExcel();
    final sheet = excel['Sheet1'];
    for (var r = 0; r < 4; r++) {
      sheet.appendRow([
        TextCellValue('C$r'),
        IntCellValue(10 + r),
        IntCellValue(5 + r),
        IntCellValue(8 + r),
      ]);
    }
    ChartSeries series(String name, String column) => ChartSeries(
        name: name,
        categoriesRange: r'Sheet1!$A$1:$A$4',
        valuesRange: 'Sheet1!\$$column\$1:\$$column\$4');
    final two = [series('S1', 'B'), series('S2', 'C')];
    const labels = ChartDataLabels(value: true, separator: '; ', labelPosition: 'ctr');
    var row = 0;
    ChartAnchor anchor() => ChartAnchor.at(column: 6, row: (row++) * 16);

    final charts = <Chart>[
      ColumnChart(title: 'Stacked', series: two, anchor: anchor(), grouping: ChartGrouping.stacked, dataLabels: labels),
      BarChart(title: 'Bar', series: two, anchor: anchor(), grouping: ChartGrouping.percentStacked, dataLabels: labels),
      LineChart(title: 'Line', series: two, anchor: anchor(), smooth: true, dataLabels: labels),
      AreaChart(title: 'Area', series: two, anchor: anchor(), dataLabels: labels),
      PieChart(title: 'Pie', series: [series('S1', 'B')], anchor: anchor(), dataLabels: labels),
      DoughnutChart(title: 'Doughnut', series: [series('S1', 'B')], anchor: anchor(), dataLabels: labels),
      OfPieChart(title: 'Of pie', series: [series('S1', 'B')], anchor: anchor(), dataLabels: labels),
      ScatterChart(title: 'Scatter', series: two, anchor: anchor(), smooth: true, dataLabels: labels),
      BubbleChart(title: 'Bubble', series: [series('S1', 'B')], anchor: anchor(), dataLabels: labels),
      RadarChart(title: 'Radar', series: two, anchor: anchor(), dataLabels: labels),
      StockChart(
          title: 'Stock',
          series: [series('High', 'B'), series('Low', 'C'), series('Close', 'D')],
          anchor: anchor(),
          dataLabels: labels),
    ];
    charts.forEach(sheet.addChart);

    final archive = ZipDecoder().decodeBytes(excel.encode()!);
    final seen = <String>{};
    for (var i = 1; i <= charts.length; i++) {
      final xml = XmlDocument.parse(
          utf8.decode(archive.findFile('xl/charts/chart$i.xml')!.content));
      final plotArea = xml.findAllElements('c:plotArea').single;
      for (final chart in plotArea.childElements
          .where((e) => _chartSequences.containsKey(e.name.local))) {
        final type = chart.name.local;
        seen.add(type);
        final problems = [
          ..._orderProblems(chart, _chartSequences[type]!),
          for (final ser in chart.findElements('c:ser'))
            ..._orderProblems(ser, _seriesSequences[type]!),
          for (final dLbls in chart.findElements('c:dLbls'))
            ..._orderProblems(dLbls, _chartSequences['dLbls']!),
        ];
        expect(problems, isEmpty, reason: 'chart$i.xml ($type)');
      }
    }
    expect(seen, containsAll(_seriesSequences.keys));
  });
}
