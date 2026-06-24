import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

void main() {
  // ---------------------------------------------------------------------------
  // Helper: create a sheet with sample data for chart tests
  // ---------------------------------------------------------------------------
  Excel _createExcelWithData() {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];
    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Value"));
    sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Alpha"));
    sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(10));
    sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Beta"));
    sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(20));
    sheet.updateCell(CellIndex.indexByString("A4"), TextCellValue("Gamma"));
    sheet.updateCell(CellIndex.indexByString("B4"), IntCellValue(30));
    return excel;
  }

  ChartSeries _defaultSeries() => ChartSeries(
        name: "S1",
        categoriesRange: r"Sheet1!$A$2:$A$4",
        valuesRange: r"Sheet1!$B$2:$B$4",
      );

  // ---------------------------------------------------------------------------
  // 1. ZIP archive is valid and can be decoded without errors
  // ---------------------------------------------------------------------------
  test('ZIP archive produced by save() is valid', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "T",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    expect(bytes, isNotNull);

    expect(
      () => ZipDecoder().decodeBytes(bytes!),
      returnsNormally,
      reason: 'The ZIP archive should decode without throwing',
    );
  });

  // ---------------------------------------------------------------------------
  // 2. ColumnChart generates all required parts
  // ---------------------------------------------------------------------------
  test('ColumnChart generates drawing, chart, rels, and content type entries',
      () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Column",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);
    var names = archive.files.map((f) => f.name).toList();

    expect(names, contains('xl/drawings/drawing1.xml'));
    expect(names, contains('xl/charts/chart1.xml'));
    expect(names, contains('xl/drawings/_rels/drawing1.xml.rels'));
    expect(names, contains('[Content_Types].xml'));

    var ct = String.fromCharCodes(archive.files
        .firstWhere((f) => f.name == '[Content_Types].xml')
        .content);
    expect(ct, contains('/xl/charts/chart1.xml'));
    expect(ct, contains('/xl/drawings/drawing1.xml'));
  });

  // ---------------------------------------------------------------------------
  // 3. LineChart generates required parts and well-formed XML
  // ---------------------------------------------------------------------------
  test('LineChart produces valid chart XML', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(LineChart(
      title: "Line",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));
    var chartSpace = xml.findAllElements('c:chartSpace');
    expect(chartSpace.isNotEmpty, isTrue,
        reason: 'chartSpace root element must exist');

    var lineChart = xml.findAllElements('c:lineChart');
    expect(lineChart.isNotEmpty, isTrue,
        reason: 'lineChart element must exist');
  });

  // ---------------------------------------------------------------------------
  // 4. PieChart produces valid chart XML (no axes)
  // ---------------------------------------------------------------------------
  test('PieChart produces chart XML without axes', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(PieChart(
      title: "Pie",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));

    expect(xml.findAllElements('c:pieChart').isNotEmpty, isTrue);
    expect(xml.findAllElements('c:catAx').isEmpty, isTrue,
        reason: 'PieChart must not have category axis');
    expect(xml.findAllElements('c:valAx').isEmpty, isTrue,
        reason: 'PieChart must not have value axis');
  });

  // ---------------------------------------------------------------------------
  // 5. BarChart produces valid chart XML
  // ---------------------------------------------------------------------------
  test('BarChart produces valid chart XML', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(BarChart(
      title: "Bar",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));
    expect(xml.findAllElements('c:barChart').isNotEmpty, isTrue);
  });

  // ---------------------------------------------------------------------------
  // 6. AreaChart produces valid chart XML
  // ---------------------------------------------------------------------------
  test('AreaChart produces valid chart XML', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(AreaChart(
      title: "Area",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));
    expect(xml.findAllElements('c:areaChart').isNotEmpty, isTrue);
  });

  // ---------------------------------------------------------------------------
  // 7. ScatterChart produces valid chart XML with xVal/yVal
  // ---------------------------------------------------------------------------
  test('ScatterChart uses xVal/yVal instead of cat/val', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ScatterChart(
      title: "Scatter",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));

    expect(xml.findAllElements('c:scatterChart').isNotEmpty, isTrue);
    expect(xml.findAllElements('c:xVal').isNotEmpty, isTrue);
    expect(xml.findAllElements('c:yVal').isNotEmpty, isTrue);
  });

  // ---------------------------------------------------------------------------
  // 8. DoughnutChart produces valid chart XML (no axes)
  // ---------------------------------------------------------------------------
  test('DoughnutChart produces chart XML without axes', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(DoughnutChart(
      title: "Doughnut",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));

    expect(xml.findAllElements('c:doughnutChart').isNotEmpty, isTrue);
    expect(xml.findAllElements('c:catAx').isEmpty, isTrue);
    expect(xml.findAllElements('c:valAx').isEmpty, isTrue);
  });

  // ---------------------------------------------------------------------------
  // 9. RadarChart produces valid chart XML
  // ---------------------------------------------------------------------------
  test('RadarChart produces valid chart XML', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(RadarChart(
      title: "Radar",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
      filled: true,
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));
    expect(xml.findAllElements('c:radarChart').isNotEmpty, isTrue);
  });

  // ---------------------------------------------------------------------------
  // 10. Drawing XML has correct namespaces
  // ---------------------------------------------------------------------------
  test('Drawing XML declares required namespaces', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "NS Test",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var drawingFile =
        archive.files.firstWhere((f) => f.name == 'xl/drawings/drawing1.xml');
    var content = String.fromCharCodes(drawingFile.content);

    expect(
        content,
        contains(
            'schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'));
    expect(content, contains('schemas.openxmlformats.org/drawingml/2006/main'));
    expect(
        content,
        contains(
            'schemas.openxmlformats.org/officeDocument/2006/relationships'));
  });

  // ---------------------------------------------------------------------------
  // 11. Chart XML has correct namespaces
  // ---------------------------------------------------------------------------
  test('Chart XML declares required namespaces', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "NS Test",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var content = String.fromCharCodes(chartFile.content);

    expect(
        content, contains('schemas.openxmlformats.org/drawingml/2006/chart'));
    expect(content, contains('schemas.openxmlformats.org/drawingml/2006/main'));
    expect(
        content,
        contains(
            'schemas.openxmlformats.org/officeDocument/2006/relationships'));
  });

  // ---------------------------------------------------------------------------
  // 12. Drawing relationship file links chart correctly
  // ---------------------------------------------------------------------------
  test('Drawing rels file contains correct chart relationship', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Rels Test",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var relsFile = archive.files
        .firstWhere((f) => f.name == 'xl/drawings/_rels/drawing1.xml.rels');
    var xml = XmlDocument.parse(String.fromCharCodes(relsFile.content));

    var rels = xml.findAllElements('Relationship');
    expect(rels.isNotEmpty, isTrue);

    var rel = rels.first;
    expect(rel.getAttribute('Target'), contains('chart1.xml'));
    expect(rel.getAttribute('Type'), contains('relationships/chart'));
    expect(rel.getAttribute('Id'), equals('rId1'));
  });

  // ---------------------------------------------------------------------------
  // 13. Sheet rels file links drawing correctly
  // ---------------------------------------------------------------------------
  test('Sheet rels file contains drawing relationship', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Sheet Rels",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var sheetRels = archive.files.where((f) =>
        f.name.startsWith('xl/worksheets/_rels/') && f.name.endsWith('.rels'));
    expect(sheetRels.isNotEmpty, isTrue);

    var content = String.fromCharCodes(sheetRels.first.content);
    expect(content, contains('relationships/drawing'));
    expect(content, contains('drawing1.xml'));
  });

  // ---------------------------------------------------------------------------
  // 14. Worksheet XML contains <drawing> element with r:id
  // ---------------------------------------------------------------------------
  test('Worksheet XML contains drawing element', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "WS Drawing",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var sheetFile = archive.files.firstWhere((f) =>
        f.name.startsWith('xl/worksheets/sheet') && !f.name.contains('_rels'));
    var content = String.fromCharCodes(sheetFile.content);
    expect(content, contains('<drawing'));
    expect(content, contains('r:id'));
  });

  // ---------------------------------------------------------------------------
  // 15. Multiple charts on one sheet produce correct rIds
  // ---------------------------------------------------------------------------
  test('Multiple charts on one sheet have sequential rIds in drawing rels', () {
    var excel = _createExcelWithData();
    var sheet = excel['Sheet1'];
    sheet.addChart(ColumnChart(
      title: "C1",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));
    sheet.addChart(LineChart(
      title: "C2",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 17),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var relsFile = archive.files
        .firstWhere((f) => f.name == 'xl/drawings/_rels/drawing1.xml.rels');
    var xml = XmlDocument.parse(String.fromCharCodes(relsFile.content));

    var rels = xml.findAllElements('Relationship').toList();
    expect(rels.length, equals(2));
    expect(rels[0].getAttribute('Id'), equals('rId1'));
    expect(rels[1].getAttribute('Id'), equals('rId2'));
  });

  // ---------------------------------------------------------------------------
  // 16. Multi-sheet charts produce separate drawings
  // ---------------------------------------------------------------------------
  test('Charts on separate sheets produce separate drawing files', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "S1 Chart",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var sheet2 = excel['Sheet2'];
    sheet2.updateCell(CellIndex.indexByString("A1"), IntCellValue(5));
    sheet2.addChart(ColumnChart(
      title: "S2 Chart",
      series: [
        ChartSeries(
            name: "X",
            categoriesRange: r"Sheet2!$A$1:$A$1",
            valuesRange: r"Sheet2!$A$1:$A$1")
      ],
      anchor: ChartAnchor.at(column: 2, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);
    var names = archive.files.map((f) => f.name).toList();

    expect(names, contains('xl/drawings/drawing1.xml'));
    expect(names, contains('xl/drawings/drawing2.xml'));
    expect(names, contains('xl/charts/chart1.xml'));
    expect(names, contains('xl/charts/chart2.xml'));
  });

  // ---------------------------------------------------------------------------
  // 17. Round-trip decode does not crash for any chart type
  // ---------------------------------------------------------------------------
  test('Round-trip decode succeeds for all chart types', () {
    final chartFactories = <Chart Function()>[
      () => ColumnChart(
          title: "Col",
          series: [_defaultSeries()],
          anchor: ChartAnchor.at(column: 4, row: 1)),
      () => LineChart(
          title: "Lin",
          series: [_defaultSeries()],
          anchor: ChartAnchor.at(column: 4, row: 1)),
      () => PieChart(
          title: "Pie",
          series: [_defaultSeries()],
          anchor: ChartAnchor.at(column: 4, row: 1)),
      () => BarChart(
          title: "Bar",
          series: [_defaultSeries()],
          anchor: ChartAnchor.at(column: 4, row: 1)),
      () => AreaChart(
          title: "Area",
          series: [_defaultSeries()],
          anchor: ChartAnchor.at(column: 4, row: 1)),
      () => ScatterChart(
          title: "Scat",
          series: [_defaultSeries()],
          anchor: ChartAnchor.at(column: 4, row: 1)),
      () => DoughnutChart(
          title: "Dnut",
          series: [_defaultSeries()],
          anchor: ChartAnchor.at(column: 4, row: 1)),
      () => RadarChart(
          title: "Radar",
          series: [_defaultSeries()],
          anchor: ChartAnchor.at(column: 4, row: 1)),
    ];

    for (var factory in chartFactories) {
      var excel = _createExcelWithData();
      excel['Sheet1'].addChart(factory());
      var bytes = excel.save();
      expect(bytes, isNotNull);
      expect(
        () => Excel.decodeBytes(bytes!),
        returnsNormally,
        reason:
            'Round-trip decode should not crash for ${factory().runtimeType}',
      );
    }
  });

  // ---------------------------------------------------------------------------
  // 18. Drawing XML is well-formed (parseable)
  // ---------------------------------------------------------------------------
  test('Drawing XML is well-formed and parseable', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Parse Test",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var drawingFile =
        archive.files.firstWhere((f) => f.name == 'xl/drawings/drawing1.xml');
    expect(
      () => XmlDocument.parse(String.fromCharCodes(drawingFile.content)),
      returnsNormally,
      reason: 'Drawing XML must be well-formed',
    );
  });

  // ---------------------------------------------------------------------------
  // 19. Chart XML is well-formed (parseable)
  // ---------------------------------------------------------------------------
  test('Chart XML is well-formed and parseable', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Parse Test",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    expect(
      () => XmlDocument.parse(String.fromCharCodes(chartFile.content)),
      returnsNormally,
      reason: 'Chart XML must be well-formed',
    );
  });

  // ---------------------------------------------------------------------------
  // 20. Drawing rels XML is well-formed
  // ---------------------------------------------------------------------------
  test('Drawing rels XML is well-formed and parseable', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Rels Parse",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var relsFile = archive.files
        .firstWhere((f) => f.name == 'xl/drawings/_rels/drawing1.xml.rels');
    expect(
      () => XmlDocument.parse(String.fromCharCodes(relsFile.content)),
      returnsNormally,
      reason: 'Drawing rels XML must be well-formed',
    );
  });

  // ---------------------------------------------------------------------------
  // 21. Content_Types.xml is well-formed after chart addition
  // ---------------------------------------------------------------------------
  test('Content_Types.xml is well-formed after adding charts', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "CT Test",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var ctFile =
        archive.files.firstWhere((f) => f.name == '[Content_Types].xml');
    expect(
      () => XmlDocument.parse(String.fromCharCodes(ctFile.content)),
      returnsNormally,
    );
  });

  // ---------------------------------------------------------------------------
  // 22. Drawing XML contains twoCellAnchor with graphicFrame
  // ---------------------------------------------------------------------------
  test('Drawing XML contains twoCellAnchor with graphicFrame', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Anchor Test",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var drawingFile =
        archive.files.firstWhere((f) => f.name == 'xl/drawings/drawing1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(drawingFile.content));

    expect(xml.findAllElements('xdr:twoCellAnchor').isNotEmpty, isTrue);
    expect(xml.findAllElements('xdr:graphicFrame').isNotEmpty, isTrue);
    expect(xml.findAllElements('xdr:clientData').isNotEmpty, isTrue);
  });

  // ---------------------------------------------------------------------------
  // 23. Chart title is preserved in chart XML
  // ---------------------------------------------------------------------------
  test('Chart title is present in generated chart XML', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "My Custom Title",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var content = String.fromCharCodes(chartFile.content);
    expect(content, contains('My Custom Title'));
  });

  // ---------------------------------------------------------------------------
  // 24. Saving an Excel with no charts does not produce chart XML files
  // ---------------------------------------------------------------------------
  test('No chart XML files when no charts are added', () {
    var excel = _createExcelWithData();

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);
    var names = archive.files.map((f) => f.name).toList();

    expect(names.any((n) => n.startsWith('xl/charts/')), isFalse);
  });

  // ---------------------------------------------------------------------------
  // 25. Chart series data references are in the chart XML
  // ---------------------------------------------------------------------------
  test('Chart XML contains series data range references', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Range Test",
      series: [
        ChartSeries(
          name: "MySeries",
          categoriesRange: r"Sheet1!$A$2:$A$4",
          valuesRange: r"Sheet1!$B$2:$B$4",
        ),
      ],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var content = String.fromCharCodes(chartFile.content);

    expect(content, contains(r'Sheet1!$A$2:$A$4'));
    expect(content, contains(r'Sheet1!$B$2:$B$4'));
    expect(content, contains('MySeries'));
  });

  // ---------------------------------------------------------------------------
  // 26. Chart with multiple series produces correct number of c:ser elements
  // ---------------------------------------------------------------------------
  test('Multiple series produce multiple c:ser elements', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Multi Series",
      series: [
        ChartSeries(
            name: "S1",
            categoriesRange: r"Sheet1!$A$2:$A$4",
            valuesRange: r"Sheet1!$B$2:$B$4"),
        ChartSeries(
            name: "S2",
            categoriesRange: r"Sheet1!$A$2:$A$4",
            valuesRange: r"Sheet1!$B$2:$B$4"),
      ],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));

    var serElements = xml.findAllElements('c:ser').toList();
    expect(serElements.length, equals(2));
  });

  // ---------------------------------------------------------------------------
  // 27. Chart with showLegend=false omits legend element
  // ---------------------------------------------------------------------------
  test('showLegend false omits c:legend from chart XML', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "No Legend",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
      showLegend: false,
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));

    expect(xml.findAllElements('c:legend').isEmpty, isTrue,
        reason: 'Legend should be absent when showLegend is false');
  });

  // ---------------------------------------------------------------------------
  // 28. ColumnChart with axes generates catAx and valAx
  // ---------------------------------------------------------------------------
  test('ColumnChart generates category and value axes', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Axes Test",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));

    expect(xml.findAllElements('c:catAx').isNotEmpty, isTrue);
    expect(xml.findAllElements('c:valAx').isNotEmpty, isTrue);
  });

  // ---------------------------------------------------------------------------
  // 29. ScatterChart generates two valAx (no catAx)
  // ---------------------------------------------------------------------------
  test('ScatterChart generates two value axes and no category axis', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ScatterChart(
      title: "Scatter Axes",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(chartFile.content));

    expect(xml.findAllElements('c:valAx').length, equals(2));
    expect(xml.findAllElements('c:catAx').isEmpty, isTrue);
  });

  // ---------------------------------------------------------------------------
  // 30. Drawing anchor position matches chart anchor coordinates
  // ---------------------------------------------------------------------------
  test('Drawing anchor from/to match chart anchor values', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Pos Test",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 5, row: 3, width: 10, height: 12),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var drawingFile =
        archive.files.firstWhere((f) => f.name == 'xl/drawings/drawing1.xml');
    var xml = XmlDocument.parse(String.fromCharCodes(drawingFile.content));

    var from = xml.findAllElements('xdr:from').first;
    var to = xml.findAllElements('xdr:to').first;

    expect(from.findElements('xdr:col').first.innerText, equals('5'));
    expect(from.findElements('xdr:row').first.innerText, equals('3'));
    expect(to.findElements('xdr:col').first.innerText, equals('15'));
    expect(to.findElements('xdr:row').first.innerText, equals('15'));
  });

  // ---------------------------------------------------------------------------
  // 31. Open a saved ColumnChart file and retrieve the chart from the archive
  // ---------------------------------------------------------------------------
  test('Open saved ColumnChart file and retrieve chart data from archive', () {
    var excel = _createExcelWithData();
    var sheet = excel['Sheet1'];
    sheet.addChart(ColumnChart(
      title: "Sales Overview",
      series: [
        ChartSeries(
          name: "Revenue",
          categoriesRange: r"Sheet1!$A$2:$A$4",
          valuesRange: r"Sheet1!$B$2:$B$4",
        ),
      ],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    expect(bytes, isNotNull);

    var reopened = Excel.decodeBytes(bytes!);
    expect(reopened.sheets.containsKey('Sheet1'), isTrue);

    var archive = ZipDecoder().decodeBytes(bytes);
    var names = archive.files.map((f) => f.name).toList();

    expect(names, contains('xl/charts/chart1.xml'));
    expect(names, contains('xl/drawings/drawing1.xml'));
    expect(names, contains('xl/drawings/_rels/drawing1.xml.rels'));

    var chartFile =
        archive.files.firstWhere((f) => f.name == 'xl/charts/chart1.xml');
    var chartXml = XmlDocument.parse(String.fromCharCodes(chartFile.content));

    var barChart = chartXml.findAllElements('c:barChart');
    expect(barChart.isNotEmpty, isTrue,
        reason: 'ColumnChart should produce a barChart element');

    var serElements = chartXml.findAllElements('c:ser').toList();
    expect(serElements.length, equals(1));

    var seriesTx = serElements.first.findAllElements('c:tx').first;
    var seriesName = seriesTx.findAllElements('c:v').first.innerText;
    expect(seriesName, equals('Revenue'));

    var title = chartXml.findAllElements('c:title').first;
    var titleText = title.findAllElements('a:t').first.innerText;
    expect(titleText, equals('Sales Overview'));

    var catRef = chartXml
        .findAllElements('c:cat')
        .first
        .findAllElements('c:f')
        .first
        .innerText;
    expect(catRef, equals(r'Sheet1!$A$2:$A$4'));

    var valRef = chartXml
        .findAllElements('c:val')
        .first
        .findAllElements('c:f')
        .first
        .innerText;
    expect(valRef, equals(r'Sheet1!$B$2:$B$4'));

    var strCache = chartXml.findAllElements('c:strCache');
    expect(strCache.isNotEmpty, isTrue,
        reason: 'Category string cache should be present');
    var cachedCategories =
        strCache.first.findAllElements('c:v').map((e) => e.innerText).toList();
    expect(cachedCategories, equals(['Alpha', 'Beta', 'Gamma']));

    var numCache = chartXml.findAllElements('c:numCache');
    expect(numCache.isNotEmpty, isTrue,
        reason: 'Value number cache should be present');
    var cachedValues =
        numCache.first.findAllElements('c:v').map((e) => e.innerText).toList();
    expect(cachedValues, equals(['10', '20', '30']));
  });

  // ===========================================================================
  // MICROSOFT EXCEL COMPATIBILITY TESTS
  // ===========================================================================

  // ---------------------------------------------------------------------------
  // 32. Sheet rels XML must use xmlns= (not xmlns:=) for default namespace
  // ---------------------------------------------------------------------------
  test(
      'Sheet rels uses xmlns= without colon for default namespace (MSXML strict)',
      () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "NS Compat",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var sheetRels = archive.files.firstWhere((f) =>
        f.name.startsWith('xl/worksheets/_rels/') && f.name.endsWith('.rels'));
    var content = String.fromCharCodes(sheetRels.content);

    expect(content, isNot(contains('xmlns:=')),
        reason:
            'xmlns:= is malformed; Excel rejects it. Must be xmlns= for the default namespace');
    expect(content, contains('xmlns='),
        reason: 'Sheet rels must declare the default namespace with xmlns=');
  });

  // ---------------------------------------------------------------------------
  // 33. Drawing rels XML must use xmlns= (not xmlns:=) for default namespace
  // ---------------------------------------------------------------------------
  test(
      'Drawing rels uses xmlns= without colon for default namespace (MSXML strict)',
      () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "NS Compat",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var drawingRels = archive.files
        .firstWhere((f) => f.name == 'xl/drawings/_rels/drawing1.xml.rels');
    var content = String.fromCharCodes(drawingRels.content);

    expect(content, isNot(contains('xmlns:=')),
        reason:
            'xmlns:= is malformed; Excel rejects it. Must be xmlns= for the default namespace');
    expect(content, contains('xmlns='),
        reason: 'Drawing rels must declare the default namespace with xmlns=');
  });

  // ---------------------------------------------------------------------------
  // 34. All .rels files in archive must not contain xmlns:= anywhere
  // ---------------------------------------------------------------------------
  test(
      'No .rels file in the archive contains the malformed xmlns:= declaration',
      () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Global NS",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);

    var relsFiles = archive.files.where((f) => f.name.endsWith('.rels'));
    for (var f in relsFiles) {
      var content = String.fromCharCodes(f.content);
      expect(content, isNot(contains('xmlns:=')),
          reason: '${f.name} contains malformed xmlns:= declaration');
    }
  });

  // ---------------------------------------------------------------------------
  // 35. Chart relationship file (xl/charts/_rels/chartN.xml.rels) must exist
  // ---------------------------------------------------------------------------
  test('Chart relationship file exists for each chart', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "Chart Rels",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);
    var names = archive.files.map((f) => f.name).toList();

    expect(names, contains('xl/charts/_rels/chart1.xml.rels'),
        reason: 'Excel requires a .rels file for each chart part');
  });

  // ---------------------------------------------------------------------------
  // 36. Multi-chart: each chart must have its own relationship file
  // ---------------------------------------------------------------------------
  test('Multiple charts each have their own rels file', () {
    var excel = _createExcelWithData();
    var sheet = excel['Sheet1'];
    sheet.addChart(ColumnChart(
      title: "C1",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));
    sheet.addChart(LineChart(
      title: "C2",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 17),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);
    var names = archive.files.map((f) => f.name).toList();

    expect(names, contains('xl/charts/_rels/chart1.xml.rels'),
        reason: 'chart1.xml.rels must exist');
    expect(names, contains('xl/charts/_rels/chart2.xml.rels'),
        reason: 'chart2.xml.rels must exist');
  });

  // ---------------------------------------------------------------------------
  // 37. Multi-sheet: chart rels files exist for charts across sheets
  // ---------------------------------------------------------------------------
  test('Charts on separate sheets each have rels files', () {
    var excel = _createExcelWithData();
    excel['Sheet1'].addChart(ColumnChart(
      title: "S1C",
      series: [_defaultSeries()],
      anchor: ChartAnchor.at(column: 4, row: 1),
    ));

    var sheet2 = excel['Sheet2'];
    sheet2.updateCell(CellIndex.indexByString("A1"), IntCellValue(5));
    sheet2.addChart(ColumnChart(
      title: "S2C",
      series: [
        ChartSeries(
            name: "X",
            categoriesRange: r"Sheet2!$A$1:$A$1",
            valuesRange: r"Sheet2!$A$1:$A$1")
      ],
      anchor: ChartAnchor.at(column: 2, row: 1),
    ));

    var bytes = excel.save();
    var archive = ZipDecoder().decodeBytes(bytes!);
    var names = archive.files.map((f) => f.name).toList();

    expect(names, contains('xl/charts/_rels/chart1.xml.rels'));
    expect(names, contains('xl/charts/_rels/chart2.xml.rels'));
  });
}
