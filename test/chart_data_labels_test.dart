import 'dart:convert';

import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';
import 'package:archive/archive.dart';

void main() {
  // ---------------------------------------------------------------------------
  // ChartDataLabels model
  // ---------------------------------------------------------------------------

  group('ChartDataLabels model', () {
    test('default instance has all flags false', () {
      const labels = ChartDataLabels();
      expect(labels.value, isFalse);
      expect(labels.categoryName, isFalse);
      expect(labels.seriesName, isFalse);
      expect(labels.percentage, isFalse);
      expect(labels.separator, equals(', '));
      expect(labels.labelPosition, isNull);
    });

    test('isEnabled returns false when all flags are false', () {
      expect(const ChartDataLabels().isEnabled, isFalse);
    });

    test('isEnabled returns true when value is true', () {
      expect(const ChartDataLabels(value: true).isEnabled, isTrue);
    });

    test('isEnabled returns true when percentage is true', () {
      expect(const ChartDataLabels(percentage: true).isEnabled, isTrue);
    });

    test('equality holds', () {
      const a = ChartDataLabels(value: true, separator: '; ');
      const b = ChartDataLabels(value: true, separator: '; ');
      expect(a, equals(b));
      expect(a.hashCode, equals(b.hashCode));
    });

    test('inequality on different flags', () {
      const a = ChartDataLabels(value: true);
      const b = ChartDataLabels(categoryName: true);
      expect(a, isNot(equals(b)));
    });
  });

  // ---------------------------------------------------------------------------
  // Chart constructors accept dataLabels
  // ---------------------------------------------------------------------------

  group('Chart constructors', () {
    final series = [
      ChartSeries(
        name: 'S1',
        categoriesRange: r"Sheet1!$A$2:$A$3",
        valuesRange: r"Sheet1!$B$2:$B$3",
      ),
    ];
    final anchor = ChartAnchor.at(column: 2, row: 2);
    const labels = ChartDataLabels(value: true, categoryName: true);

    test('ColumnChart accepts dataLabels', () {
      final c = ColumnChart(
          title: 'T', series: series, anchor: anchor, dataLabels: labels);
      expect(c.dataLabels, equals(labels));
    });

    test('LineChart accepts dataLabels', () {
      final c = LineChart(
          title: 'T', series: series, anchor: anchor, dataLabels: labels);
      expect(c.dataLabels, equals(labels));
    });

    test('PieChart accepts dataLabels with percentage', () {
      const pieLbls = ChartDataLabels(value: true, percentage: true);
      final c = PieChart(
          title: 'T', series: series, anchor: anchor, dataLabels: pieLbls);
      expect(c.dataLabels!.percentage, isTrue);
    });

    test('DoughnutChart accepts dataLabels', () {
      final c = DoughnutChart(
          title: 'T', series: series, anchor: anchor, dataLabels: labels);
      expect(c.dataLabels, equals(labels));
    });

    test('BarChart accepts dataLabels', () {
      final c =
          BarChart(title: 'T', series: series, anchor: anchor, dataLabels: labels);
      expect(c.dataLabels, equals(labels));
    });

    test('AreaChart accepts dataLabels', () {
      final c = AreaChart(
          title: 'T', series: series, anchor: anchor, dataLabels: labels);
      expect(c.dataLabels, equals(labels));
    });

    test('ScatterChart accepts dataLabels', () {
      final c = ScatterChart(
          title: 'T', series: series, anchor: anchor, dataLabels: labels);
      expect(c.dataLabels, equals(labels));
    });

    test('RadarChart accepts dataLabels', () {
      final c = RadarChart(
          title: 'T', series: series, anchor: anchor, dataLabels: labels);
      expect(c.dataLabels, equals(labels));
    });

    test('chart without dataLabels has null field', () {
      final c = ColumnChart(title: 'T', series: series, anchor: anchor);
      expect(c.dataLabels, isNull);
    });
  });

  // ---------------------------------------------------------------------------
  // XML generation
  // ---------------------------------------------------------------------------

  group('ChartXmlWriter data labels XML generation', () {
    ChartSeries makeSeries() => ChartSeries(
          name: 'Sales',
          categoriesRange: r"Sheet1!$A$2:$A$4",
          valuesRange: r"Sheet1!$B$2:$B$4",
        );

    String chartXmlFor(Chart chart) {
      final bytes = _makeExcelBytes(chart);
      final archive = ZipDecoder().decodeBytes(bytes);
      final chartFile =
          archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
      chartFile.decompress();
      return utf8.decode(chartFile.content);
    }

    test('no dLbls element when dataLabels is null', () {
      final chart = ColumnChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('dLbls'), isFalse);
    });

    test('no dLbls element when dataLabels.isEnabled is false', () {
      final chart = ColumnChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('dLbls'), isFalse);
    });

    test('showVal=1 when value:true', () {
      final chart = ColumnChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(value: true),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('showVal'), isTrue);
      expect(xml.contains('showVal val="1"'), isTrue);
    });

    test('showCatName=1 when categoryName:true', () {
      final chart = LineChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(categoryName: true),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('showCatName val="1"'), isTrue);
    });

    test('showSerName=1 when seriesName:true', () {
      final chart = BarChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(seriesName: true),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('showSerName val="1"'), isTrue);
    });

    test('showPercent=1 on PieChart with percentage:true', () {
      final chart = PieChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(percentage: true),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('showPercent val="1"'), isTrue);
    });

    test('custom separator is written when not default', () {
      final chart = ColumnChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(value: true, separator: '\n'),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('<c:separator>'), isTrue);
    });

    test('default separator is NOT written', () {
      final chart = ColumnChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(value: true),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('<c:separator>'), isFalse);
    });

    test('dLblPos is written when labelPosition is set', () {
      final chart = ColumnChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(value: true, labelPosition: 'outEnd'),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('dLblPos'), isTrue);
      expect(xml.contains('val="outEnd"'), isTrue);
    });

    test('pie chart gets showLeaderLines=1', () {
      final chart = PieChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(value: true),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('showLeaderLines val="1"'), isTrue);
    });

    test('column chart gets showLeaderLines=0', () {
      final chart = ColumnChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels: const ChartDataLabels(value: true),
      );
      final xml = chartXmlFor(chart);
      expect(xml.contains('showLeaderLines val="0"'), isTrue);
    });

    test('generated file can be decoded back without error', () {
      final chart = ColumnChart(
        title: 'T',
        series: [makeSeries()],
        anchor: ChartAnchor.at(column: 2, row: 2),
        dataLabels:
            const ChartDataLabels(value: true, categoryName: true, seriesName: true),
      );
      final bytes = _makeExcelBytes(chart);
      expect(() => Excel.decodeBytes(bytes), returnsNormally);
    });
  });

  // ---------------------------------------------------------------------------
  // Round-trip: parseDataLabelsFromXml
  // ---------------------------------------------------------------------------

  group('ChartXmlWriter.parseDataLabelsFromXml', () {
    XmlElement? nodeFromXml(String xml) {
      try {
        return XmlDocument.parse(xml).rootElement;
      } catch (_) {
        return null;
      }
    }

    test('returns null for null input', () {
      expect(ChartXmlWriter.parseDataLabelsFromXml(null), isNull);
    });

    test('returns null when all show flags are false', () {
      final node = nodeFromXml('''
        <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
        </c:dLbls>
      ''');
      expect(ChartXmlWriter.parseDataLabelsFromXml(node), isNull);
    });

    test('parses showVal correctly', () {
      final node = nodeFromXml('''
        <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:showVal val="1"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
        </c:dLbls>
      ''');
      final result = ChartXmlWriter.parseDataLabelsFromXml(node);
      expect(result, isNotNull);
      expect(result!.value, isTrue);
      expect(result.categoryName, isFalse);
    });

    test('parses showPercent correctly', () {
      final node = nodeFromXml('''
        <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="1"/>
        </c:dLbls>
      ''');
      final result = ChartXmlWriter.parseDataLabelsFromXml(node);
      expect(result!.percentage, isTrue);
    });

    test('parses custom separator', () {
      final node = nodeFromXml('''
        <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:separator>\\n</c:separator>
          <c:showVal val="1"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
        </c:dLbls>
      ''');
      final result = ChartXmlWriter.parseDataLabelsFromXml(node);
      expect(result!.separator, equals('\\n'));
    });

    test('parses dLblPos', () {
      final node = nodeFromXml('''
        <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:showVal val="1"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:dLblPos val="outEnd"/>
        </c:dLbls>
      ''');
      final result = ChartXmlWriter.parseDataLabelsFromXml(node);
      expect(result!.labelPosition, equals('outEnd'));
    });
  });
}

// ---------------------------------------------------------------------------
// Helper
// ---------------------------------------------------------------------------

List<int> _makeExcelBytes(Chart chart) {
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  sheet.updateCell(CellIndex.indexByString('A2'), TextCellValue('Jan'));
  sheet.updateCell(CellIndex.indexByString('A3'), TextCellValue('Feb'));
  sheet.updateCell(CellIndex.indexByString('A4'), TextCellValue('Mar'));
  sheet.updateCell(CellIndex.indexByString('B2'), IntCellValue(10));
  sheet.updateCell(CellIndex.indexByString('B3'), IntCellValue(20));
  sheet.updateCell(CellIndex.indexByString('B4'), IntCellValue(30));
  sheet.addChart(chart);
  return excel.save()!;
}
