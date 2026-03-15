part of '../../../excel_community.dart';

class _ChartManager {
  final Excel _excel;
  final Save _save;

  _ChartManager(this._excel, this._save);

  void processCharts() {
    final writer = ChartXmlWriter();
    int chartCount = 0;
    int drawingCount = 0;

    _excel._sheetMap.forEach((sheetName, sheet) {
      if (sheet.charts.isEmpty) return;

      drawingCount++;
      final drawingPath = 'xl/drawings/drawing$drawingCount.xml';
      final drawingRelsPath = 'xl/drawings/_rels/drawing$drawingCount.xml.rels';
      final sheetId = _excel._xmlSheetId[sheetName]!;
      final sheetRelsPath = 'xl/worksheets/_rels/${sheetId.split("/").last}.rels';

      // 1. Generate Chart XMLs and Drawing Relationships
      final drawingRelsBuilder = XmlBuilder();
      drawingRelsBuilder.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
      drawingRelsBuilder.element('Relationships', namespaces: <String, String>{
        'http://schemas.openxmlformats.org/package/2006/relationships': '',
      }, nest: () {
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
          }
          
          final chartPath = 'xl/charts/chart$chartCount.xml';

          // Generate Chart XML
          _excel._xmlFiles[chartPath] = writer.generateChartXml(chart);

          // Add to Drawing Relationships
          drawingRelsBuilder.element('Relationship', attributes: {
            'Id': 'rId${i + 1}',
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
            'Target': '../charts/chart$chartCount.xml',
          });

          // Add Content Type for Chart
          _save._addContentType(
            'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
            '/$chartPath',
          );
        }
      });
      _excel._xmlFiles[drawingRelsPath] = drawingRelsBuilder.buildDocument();

      // 2. Generate Drawing XML (linking to rId1, rId2... in drawingRels)
      _excel._xmlFiles[drawingPath] = writer.generateDrawingXml(sheet.charts, drawingCount);

      // Add Content Type for Drawing
      _save._addContentType(
        'application/vnd.openxmlformats-officedocument.drawing+xml',
        '/$drawingPath',
      );

      // 3. Update Sheet Relationships
      var sheetRels = _excel._xmlFiles[sheetRelsPath];
      if (sheetRels == null) {
        final relsBuilder = XmlBuilder();
        relsBuilder.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
        relsBuilder.element('Relationships', namespaces: <String, String>{
          'http://schemas.openxmlformats.org/package/2006/relationships': '',
        }, nest: () {});
        sheetRels = relsBuilder.buildDocument();
        _excel._xmlFiles[sheetRelsPath] = sheetRels;
      }

      final relsElement = sheetRels.findAllElements('Relationships').first;
      
      // Look for existing rId to avoid duplicates
      int rIdIndex = relsElement.children.whereType<XmlElement>().length + 1;
      final drawingRId = 'rId$rIdIndex';
      
      relsElement.children.add(XmlElement(XmlName('Relationship'), [
        XmlAttribute(XmlName('Id'), drawingRId),
        XmlAttribute(XmlName('Type'), 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'),
        XmlAttribute(XmlName('Target'), '../drawings/drawing$drawingCount.xml'),
      ]));

      // 4. Update Worksheet XML with <drawing> tag at the correct position
      final worksheet = _excel._xmlFiles[sheetId]!.findAllElements('worksheet').first;
      final existingDrawings = worksheet.findAllElements('drawing').toList();
      if (existingDrawings.isEmpty) {
        final drawingElement = XmlElement(XmlName('drawing'), [
          XmlAttribute(
              XmlName('id', 'r'),
              drawingRId),
        ]);

        int insertIndex = -1;
        final tagsAfterDrawing = [
          'legacyDrawing',
          'picture',
          'oleObjects',
          'drawingHF',
          'extLst'
        ];

        for (int i = 0; i < worksheet.children.length; i++) {
          final child = worksheet.children[i];
          if (child is XmlElement && tagsAfterDrawing.contains(child.name.local)) {
            insertIndex = i;
            break;
          }
        }

        if (insertIndex != -1) {
          worksheet.children.insert(insertIndex, drawingElement);
        } else {
          worksheet.children.add(drawingElement);
        }
      }
    });
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
