part of '../../../excel_community.dart';

class _ChartManager {
  final Excel _excel;
  final Save _save;

  _ChartManager(this._excel, this._save);

  void processCharts() {
    final writer = ChartXmlWriter();
    int chartCount = 0;

    _excel._sheetMap.forEach((sheetName, sheet) {
      if (sheet.charts.isEmpty) return;

      final sheetId = _excel._xmlSheetId[sheetName]!;
      final sheetFileName = sheetId.split('/').last; // e.g. sheet1.xml
      final sheetRelsPath = 'xl/worksheets/_rels/$sheetFileName.rels';

      // --- Resolve or create the drawing for this sheet ---
      String drawingPath;
      String drawingRelsPath;
      bool drawingIsNew;
      int drawingIdx;

      final existing = _findExistingDrawing(sheetRelsPath);
      if (existing != null) {
        drawingPath = existing.$1;
        drawingRelsPath = existing.$2;
        drawingIsNew = false;
        final nameOnly = drawingPath.split('/').last; // drawing2.xml
        final digits = RegExp(r'\d+').stringMatch(nameOnly) ?? '1';
        drawingIdx = int.parse(digits);
      } else {
        final idx = _countExistingDrawings() + 1;
        drawingPath = 'xl/drawings/drawing$idx.xml';
        drawingRelsPath = 'xl/drawings/_rels/drawing$idx.xml.rels';
        drawingIsNew = true;
        drawingIdx = idx;
      }

      // --- Ensure drawing rels XML exists ---
      var drawingRels = _excel._xmlFiles[drawingRelsPath];
      if (drawingRels == null) {
        final relsBuilder = XmlBuilder();
        relsBuilder.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
        relsBuilder.element('Relationships', attributes: {
          'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
        }, nest: () {});
        drawingRels = relsBuilder.buildDocument();
        _excel._xmlFiles[drawingRelsPath] = drawingRels;
      }
      final relsRoot = drawingRels.findAllElements('Relationships').first;
      int nextRId = relsRoot.children.whereType<XmlElement>().length + 1;

      // --- Ensure drawing XML exists ---
      var drawingDoc = _excel._xmlFiles[drawingPath];
      if (drawingDoc == null) {
        drawingDoc = _buildEmptyDrawing();
        _excel._xmlFiles[drawingPath] = drawingDoc;
      }
      final drawingRoot = drawingDoc.rootElement;

      // --- Process each chart ---
      for (int i = 0; i < sheet.charts.length; i++) {
        chartCount++;
        final chart = sheet.charts[i];
        
        // Resolve ranges for each series
        for (var series in chart.series) {
          final catData = _resolveChartRange(series.categoriesRange);
          final valData = _resolveChartRange(series.valuesRange);
          
          series.categories = catData.map((e) => e?.toString() ?? "").toList();
          series.values = valData.map((e) {
            if (e is IntCellValue) return e.value;
            if (e is DoubleCellValue) return e.value;
            if (e is TextCellValue) {
              return num.tryParse(e.toString()) ?? 0;
            }
            return 0;
          }).toList();

          // Fix #4: para ScatterChart resolver xValues (categoriesRange es eje X numérico)
          if (chart is ScatterChart) {
            series.xValues = catData.map((e) {
              if (e is IntCellValue) return e.value as num;
              if (e is DoubleCellValue) return e.value as num;
              if (e is TextCellValue) return num.tryParse(e.toString()) ?? 0;
              return 0 as num;
            }).toList();
          }
        }
        
        final chartPath = 'xl/charts/chart$chartCount.xml';

        // Generate Chart XML
        _excel._xmlFiles[chartPath] = writer.generateChartXml(chart);

        // Generate Chart .rels file (required by Excel for each chart part)
        final chartRelsPath = 'xl/charts/_rels/chart$chartCount.xml.rels';
        final chartRelsBuilder = XmlBuilder();
        chartRelsBuilder.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
        chartRelsBuilder.element('Relationships', attributes: {
          'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
        }, nest: () {});
        _excel._xmlFiles[chartRelsPath] = chartRelsBuilder.buildDocument();

        // Add to Drawing Relationships
        final rId = 'rId$nextRId';
        nextRId++;

        relsRoot.children.add(XmlElement(XmlName('Relationship'), [
          XmlAttribute(XmlName('Id'), rId),
          XmlAttribute(
            XmlName('Type'),
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
          ),
          XmlAttribute(XmlName('Target'), '../charts/chart$chartCount.xml'),
        ]));

        // Add to Drawing XML using the actual relationship ID
        drawingRoot.children.add(
          writer.buildChartAnchorElement(chart, i, drawingIdx, rId),
        );

        // Add Content Type for Chart
        _save._addContentType(
          'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
          '/$chartPath',
        );
      }

      // --- If drawing is new, wire it up to the worksheet ---
      if (drawingIsNew) {
        _save._addContentType(
          'application/vnd.openxmlformats-officedocument.drawing+xml',
          '/$drawingPath',
        );

        // Add drawing relationship to the sheet's rels file
        var sheetRels = _excel._xmlFiles[sheetRelsPath];
        if (sheetRels == null) {
          final relsBuilder = XmlBuilder();
          relsBuilder.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
          relsBuilder.element('Relationships', attributes: {
            'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
          }, nest: () {});
          sheetRels = relsBuilder.buildDocument();
          _excel._xmlFiles[sheetRelsPath] = sheetRels;
        }

        final sheetRelsRoot = sheetRels.findAllElements('Relationships').first;
        final drawingRIdIndex = sheetRelsRoot.children.whereType<XmlElement>().length + 1;
        final drawingRId = 'rId$drawingRIdIndex';
        final drawingFileName = drawingPath.split('/').last;

        sheetRelsRoot.children.add(XmlElement(XmlName('Relationship'), [
          XmlAttribute(XmlName('Id'), drawingRId),
          XmlAttribute(
            XmlName('Type'),
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          ),
          XmlAttribute(XmlName('Target'), '../drawings/$drawingFileName'),
        ]));

        sheet._drawingRId = drawingRId;
      }
    });
  }

  (String, String)? _findExistingDrawing(String sheetRelsPath) {
    final sheetRels = _excel._xmlFiles[sheetRelsPath];
    if (sheetRels == null) return null;

    for (final rel in sheetRels.findAllElements('Relationship')) {
      final type = rel.getAttribute('Type') ?? '';
      if (type.endsWith('/drawing')) {
        final target = rel.getAttribute('Target') ?? '';
        final drawingFileName = target.split('/').last;
        return (
          'xl/drawings/$drawingFileName',
          'xl/drawings/_rels/$drawingFileName.rels',
        );
      }
    }
    return null;
  }

  int _countExistingDrawings() {
    return _excel._xmlFiles.keys
        .where((k) =>
            k.startsWith('xl/drawings/drawing') &&
            k.endsWith('.xml') &&
            !k.contains('/_rels/'))
        .length;
  }

  XmlDocument _buildEmptyDrawing() {
    final b = XmlBuilder();
    b.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    b.element('xdr:wsDr', namespaces: {
      'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing': 'xdr',
      'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
      'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
    }, nest: () {});
    return b.buildDocument();
  }

  List<CellValue?> _resolveChartRange(String range) {
    if (range.isEmpty) return [];

    try {
      final parts = range.split('!');
      if (parts.length != 2) return [];

      final sheetName = parts[0].replaceAll("'", "");
      final cellRange = parts[1].replaceAll("\$", "");

      final sheet = _excel._sheetMap[sheetName];
      if (sheet == null) return [];

      final rangeParts = cellRange.split(':');
      if (rangeParts.length == 1) {
        // Single cell
        final coords = _cellCoordsFromCellId(rangeParts[0]);
        final cell = sheet._sheetData[coords.$1]?[coords.$2];
        return [cell?.value];
      } else if (rangeParts.length == 2) {
        // Range
        final startCoords = _cellCoordsFromCellId(rangeParts[0]);
        final endCoords = _cellCoordsFromCellId(rangeParts[1]);

        final List<CellValue?> values = [];
        for (int row = startCoords.$1; row <= endCoords.$1; row++) {
          for (int col = startCoords.$2; col <= endCoords.$2; col++) {
            final cell = sheet._sheetData[row]?[col];
            values.add(cell?.value);
          }
        }
        return values;
      }
    } catch (e) {
      // Ignore parsing errors
    }
    return [];
  }
}
