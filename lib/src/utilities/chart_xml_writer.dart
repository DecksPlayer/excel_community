part of '../../excel_community.dart';

/// Internal utility to generate XML for Charts and Drawings.
///
/// Organized by responsibility:
/// - Drawing XML: Structure for drawing elements
/// - Chart Orchestration: High-level chart building logic
/// - Chart Type Builders: Specific builders for each chart type
/// - Color Configuration: Centralized color palettes and logic
/// - XML Components: Reusable XML structure helpers
class ChartXmlWriter {
  ChartXmlWriter();

  // ========================================================================
  // PUBLIC API
  // ========================================================================

  /// Generates the Drawing XML (xl/drawings/drawing*.xml)
  XmlDocument generateDrawingXml(List<Chart> charts, int drawingCount) {
    final builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('xdr:wsDr', namespaceUris: _buildDrawingNamespaces(),
        nest: () {
      for (int i = 0; i < charts.length; i++) {
        _buildChartAnchor(builder, charts[i], i, drawingCount, 'rId${i + 1}');
      }
    });
    return builder.buildDocument();
  }

  /// Builds a `<xdr:twoCellAnchor>` element for the given [chart] using [XmlBuilder].
  XmlElement buildChartAnchorElement(
      Chart chart, int index, int drawingCount, String rId) {
    final builder = XmlBuilder();
    builder.element('xdr:twoCellAnchor', namespaceUris: _buildDrawingNamespaces(),
        nest: () {
      _buildAnchorPosition(
          builder, 'xdr:from', chart.anchor.fromColumn, chart.anchor.fromRow);
      _buildAnchorPosition(
          builder, 'xdr:to', chart.anchor.toColumn, chart.anchor.toRow);
      _buildGraphicFrame(builder, index, drawingCount, rId);
      builder.element('xdr:clientData');
    });
    return builder.buildDocument().rootElement.copy();
  }

  // Generates the Chart XML (xl/charts/chart*.xml)
  XmlDocument generateChartXml(Chart chart) {
    final builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('c:chartSpace', namespaceUris: _buildChartNamespaces(),
        nest: () {
      // c:lang is required by Excel as first child of chartSpace
      builder.element('c:lang', attributes: {'val': 'en-US'});
      builder.element('c:chart', nest: () {
        // CT_Chart sequence (OOXML §21.2.2.29): title comes BEFORE autoTitleDeleted
        _buildChartTitle(builder, chart.title);
        builder.element('c:autoTitleDeleted', attributes: {'val': '0'});
        _buildPlotArea(builder, chart);
        // plotVisOnly/dispBlanksAs/showDLblsOverMax must come AFTER legend inside c:chart
        if (chart.showLegend) {
          _buildLegend(builder);
        }
        _buildChartOptions(builder);
      });
    });
    return builder.buildDocument();
  }

  // ========================================================================
  // PRIVATE: Drawing XML Helpers
  // ========================================================================

  Map<String, String> _buildDrawingNamespaces() {
    return {
      'xdr':
          'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    };
  }

  void _buildChartAnchor(XmlBuilder builder, Chart chart, int index,
      int drawingCount, String rId) {
    builder.element('xdr:twoCellAnchor', nest: () {
      _buildAnchorPosition(
          builder, 'xdr:from', chart.anchor.fromColumn, chart.anchor.fromRow);
      _buildAnchorPosition(
          builder, 'xdr:to', chart.anchor.toColumn, chart.anchor.toRow);
      _buildGraphicFrame(builder, index, drawingCount, rId);
      builder.element('xdr:clientData');
    });
  }

  void _buildAnchorPosition(
      XmlBuilder builder, String element, int column, int row) {
    builder.element(element, nest: () {
      builder.element('xdr:col', nest: () => builder.text(column.toString()));
      builder.element('xdr:colOff', nest: () => builder.text('0'));
      builder.element('xdr:row', nest: () => builder.text(row.toString()));
      builder.element('xdr:rowOff', nest: () => builder.text('0'));
    });
  }

  void _buildGraphicFrame(
      XmlBuilder builder, int index, int drawingCount, String rId) {
    builder.element('xdr:graphicFrame', nest: () {
      builder.element('xdr:nvGraphicFramePr', nest: () {
        builder.element('xdr:cNvPr', attributes: {
          'id': '${index + 1 + (drawingCount * 1024)}',
          'name': 'Chart ${index + 1}'
        });
        builder.element('xdr:cNvGraphicFramePr');
      });
      builder.element('xdr:xfrm', nest: () {
        builder.element('a:off', attributes: {'x': '0', 'y': '0'});
        builder.element('a:ext', attributes: {'cx': '0', 'cy': '0'});
      });
      builder.element('a:graphic', nest: () {
        builder.element('a:graphicData', attributes: {
          'uri': 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        }, nest: () {
          builder.element('c:chart', attributes: {'r:id': rId});
        });
      });
    });
  }

  // ========================================================================
  // PRIVATE: Chart XML Helpers
  // ========================================================================

  Map<String, String> _buildChartNamespaces() {
    return {
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    };
  }

  void _buildChartTitle(XmlBuilder builder, String title) {
    builder.element('c:title', nest: () {
      builder.element('c:tx', nest: () {
        builder.element('c:rich', nest: () {
          builder.element('a:bodyPr');
          builder.element('a:lstStyle');
          builder.element('a:p', nest: () {
            builder.element('a:pPr', nest: () {
              builder.element('a:defRPr');
            });
            builder.element('a:r', nest: () {
              builder.element('a:rPr', attributes: {'lang': 'en-US'});
              builder.element('a:t', nest: () => builder.text(title));
            });
          });
        });
      });
      builder.element('c:layout');
      builder.element('c:overlay', attributes: {'val': '0'});
    });
  }

  void _buildPlotArea(XmlBuilder builder, Chart chart) {
    final bool hasAxes = chart is! PieChart && chart is! DoughnutChart;

    builder.element('c:plotArea', nest: () {
      builder.element('c:layout');
      _buildChartElement(builder, chart, hasAxes);
      if (hasAxes) {
        _buildAxes(builder, chart);
      }
    });
  }

  void _buildChartElement(XmlBuilder builder, Chart chart, bool hasAxes) {
    builder.element('c:${chart.chartTagName}', nest: () {
      _buildChartTypeProperties(builder, chart);
      _buildAllSeries(builder, chart);
      if (hasAxes) {
        _buildAxesIds(builder);
      }
    });
  }

  void _buildChartTypeProperties(XmlBuilder builder, Chart chart) {
    // Delegate to chart-specific builder
    final styleBuilder = ChartStyleBuilderFactory.getBuilder(chart);
    styleBuilder.buildProperties(builder, chart);
  }

  void _buildAllSeries(XmlBuilder builder, Chart chart) {
    for (int i = 0; i < chart.series.length; i++) {
      _buildSeries(builder, chart, chart.series[i], i);
    }
  }

  void _buildSeries(
      XmlBuilder builder, Chart chart, ChartSeries series, int index) {
    builder.element('c:ser', nest: () {
      builder.element('c:idx', attributes: {'val': '$index'});
      builder.element('c:order', attributes: {'val': '$index'});
      // CT_SerTx (OOXML spec §21.2.2.174) is a choice between c:strRef and c:v only.
      // c:strLit is NOT valid here — use c:v directly for a literal series name.
      builder.element('c:tx', nest: () {
        builder.element('c:v', nest: () => builder.text(series.name));
      });

      _buildSeriesColors(builder, chart, series, index);
      _buildSeriesData(builder, chart, series);
    });
  }

  void _buildSeriesColors(
      XmlBuilder builder, Chart chart, ChartSeries series, int index) {
    // Delegate to chart-specific builder
    final styleBuilder = ChartStyleBuilderFactory.getBuilder(chart);
    styleBuilder.buildSeriesStyle(builder, chart, series, index);
  }

  void _buildSeriesData(XmlBuilder builder, Chart chart, ChartSeries series) {
    if (chart is ScatterChart) {
      builder.element('c:xVal', nest: () {
        builder.element('c:numRef', nest: () {
          builder.element('c:f',
              nest: () => builder.text(series.categoriesRange));
          // Fix #4: usar xValues reales (resueltos en chart_manager), no ceros
          if (series.xValues != null && series.xValues!.isNotEmpty) {
            _buildNumCache(builder, series.xValues!);
          }
        });
      });
      builder.element('c:yVal', nest: () {
        builder.element('c:numRef', nest: () {
          builder.element('c:f', nest: () => builder.text(series.valuesRange));
          if (series.values != null && series.values!.isNotEmpty) {
            _buildNumCache(builder, series.values!);
          }
        });
      });
    } else {
      builder.element('c:cat', nest: () {
        builder.element('c:strRef', nest: () {
          builder.element('c:f',
              nest: () => builder.text(series.categoriesRange));
          if (series.categories != null && series.categories!.isNotEmpty) {
            _buildStrCache(builder, series.categories!);
          }
        });
      });
      builder.element('c:val', nest: () {
        builder.element('c:numRef', nest: () {
          builder.element('c:f', nest: () => builder.text(series.valuesRange));
          if (series.values != null && series.values!.isNotEmpty) {
            _buildNumCache(builder, series.values!);
          }
        });
      });
    }
  }

  void _buildStrCache(XmlBuilder builder, List<String> categories) {
    builder.element('c:strCache', nest: () {
      builder.element('c:ptCount', attributes: {'val': '${categories.length}'});
      for (int i = 0; i < categories.length; i++) {
        builder.element('c:pt', attributes: {'idx': '$i'}, nest: () {
          builder.element('c:v', nest: () => builder.text(categories[i]));
        });
      }
    });
  }

  void _buildNumCache(XmlBuilder builder, List<num> values) {
    builder.element('c:numCache', nest: () {
      builder.element('c:formatCode', nest: () => builder.text('General'));
      builder.element('c:ptCount', attributes: {'val': '${values.length}'});
      for (int i = 0; i < values.length; i++) {
        builder.element('c:pt', attributes: {'idx': '$i'}, nest: () {
          builder.element('c:v',
              nest: () => builder.text(values[i].toString()));
        });
      }
    });
  }

  void _buildAxesIds(XmlBuilder builder) {
    builder.element('c:axId', attributes: {'val': '10000001'});
    builder.element('c:axId', attributes: {'val': '10000002'});
  }

  void _buildAxes(XmlBuilder builder, Chart chart) {
    if (chart is ScatterChart) {
      _buildValueAxis(builder, id: '10000001', pos: 'b', crossAx: '10000002');
      _buildValueAxis(builder, id: '10000002', pos: 'l', crossAx: '10000001');
    } else {
      _buildCategoryAxis(builder, chart);
      _buildValueAxis(builder, id: '10000002', pos: 'l', crossAx: '10000001');
    }
  }

  void _buildCategoryAxis(XmlBuilder builder, Chart chart) {
    builder.element('c:catAx', nest: () {
      builder.element('c:axId', attributes: {'val': '10000001'});
      builder.element('c:scaling', nest: () {
        builder.element('c:orientation', attributes: {'val': 'minMax'});
      });
      builder.element('c:delete', attributes: {'val': '0'});
      builder.element('c:axPos', attributes: {'val': 'b'});
      // Fix #5: c:numFmt es requerido en catAx según OOXML spec
      builder.element('c:numFmt',
          attributes: {'formatCode': 'General', 'sourceLinked': '1'});
      builder.element('c:majorTickMark', attributes: {'val': 'out'});
      builder.element('c:minorTickMark', attributes: {'val': 'none'});
      builder.element('c:tickLblPos', attributes: {'val': 'nextTo'});
      builder.element('c:crossAx', attributes: {'val': '10000002'});
      builder.element('c:crosses', attributes: {'val': 'autoZero'});
      builder.element('c:auto', attributes: {'val': '1'});
      builder.element('c:lblAlgn', attributes: {'val': 'ctr'});
      builder.element('c:lblOffset', attributes: {'val': '100'});
    });
  }

  void _buildValueAxis(XmlBuilder builder,
      {String id = '10000002', String pos = 'l', String crossAx = '10000001'}) {
    builder.element('c:valAx', nest: () {
      builder.element('c:axId', attributes: {'val': id});
      builder.element('c:scaling', nest: () {
        builder.element('c:orientation', attributes: {'val': 'minMax'});
      });
      builder.element('c:delete', attributes: {'val': '0'});
      builder.element('c:axPos', attributes: {'val': pos});
      builder.element('c:majorGridlines');
      builder.element('c:numFmt',
          attributes: {'formatCode': 'General', 'sourceLinked': '1'});
      builder.element('c:majorTickMark', attributes: {'val': 'out'});
      builder.element('c:minorTickMark', attributes: {'val': 'none'});
      builder.element('c:tickLblPos', attributes: {'val': 'nextTo'});
      builder.element('c:crossAx', attributes: {'val': crossAx});
      builder.element('c:crosses', attributes: {'val': 'autoZero'});
      builder.element('c:crossBetween', attributes: {'val': 'between'});
    });
  }

  void _buildLegend(XmlBuilder builder) {
    builder.element('c:legend', nest: () {
      builder.element('c:legendPos', attributes: {'val': 'r'});
      builder.element('c:layout');
      builder.element('c:overlay', attributes: {'val': '0'});
    });
  }

  void _buildChartOptions(XmlBuilder builder) {
    builder.element('c:plotVisOnly', attributes: {'val': '1'});
    builder.element('c:dispBlanksAs', attributes: {'val': 'gap'});
    builder.element('c:showDLblsOverMax', attributes: {'val': '0'});
  }
}
