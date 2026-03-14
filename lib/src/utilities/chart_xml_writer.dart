part of excel;

/// Internal utility to generate XML for Charts and Drawings.
class ChartXmlWriter {
  ChartXmlWriter();

  /// Generates the Drawing XML (xl/drawings/drawing*.xml)
  XmlDocument generateDrawingXml(Chart chart, int chartIndex) {
    final builder = XmlBuilder();
    builder.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('xdr:wsDr', namespaces: {
      'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing': 'xdr',
      'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
    }, nest: () {
      builder.element('xdr:twoCellAnchor', nest: () {
        // From position
        builder.element('xdr:from', nest: () {
          builder.element('xdr:col', nest: () => builder.text(chart.anchor.fromColumn.toString()));
          builder.element('xdr:colOff', nest: () => builder.text('0'));
          builder.element('xdr:row', nest: () => builder.text(chart.anchor.fromRow.toString()));
          builder.element('xdr:rowOff', nest: () => builder.text('0'));
        });
        // To position
        builder.element('xdr:to', nest: () {
          builder.element('xdr:col', nest: () => builder.text(chart.anchor.toColumn.toString()));
          builder.element('xdr:colOff', nest: () => builder.text('0'));
          builder.element('xdr:row', nest: () => builder.text(chart.anchor.toRow.toString()));
          builder.element('xdr:rowOff', nest: () => builder.text('0'));
        });

        // Graphic Frame
        builder.element('xdr:graphicFrame', attributes: {'macro': ''}, nest: () {
          builder.element('xdr:nvGraphicFramePr', nest: () {
            builder.element('xdr:cNvPr', attributes: {
              'id': '${chartIndex + 1024}',
              'name': 'Chart ${chartIndex + 1}'
            });
            builder.element('xdr:cNvGraphicFramePr');
          });
          builder.element('xdr:xfrm', nest: () {
            builder.element('a:off', attributes: {'x': '0', 'y': '0'});
            builder.element('a:ext', attributes: {'cx': '0', 'cy': '0'});
          });
          builder.element('a:graphic', nest: () {
            builder.element('a:graphicData',
                attributes: {'uri': 'http://schemas.openxmlformats.org/drawingml/2006/chart'},
                nest: () {
              builder.element('c:chart', namespaces: {
                'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
              }, attributes: {
                'r:id': 'rId1' // This will be mapped in drawing*.xml.rels to chart*.xml
              });
            });
          });
        });
        builder.element('xdr:clientData');
      });
    });
    return builder.buildDocument();
  }

  /// Generates the Chart XML (xl/charts/chart*.xml)
  XmlDocument generateChartXml(Chart chart) {
    final builder = XmlBuilder();
    builder.processing('xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('c:chartSpace', namespaces: {
      'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
      'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
    }, nest: () {
      builder.element('c:chart', nest: () {
        builder.element('c:title', nest: () {
          builder.element('c:tx', nest: () {
            builder.element('c:rich', nest: () {
              builder.element('a:bodyPr');
              builder.element('a:lstStyle');
              builder.element('a:p', nest: () {
                builder.element('a:r', nest: () {
                  builder.element('a:t', nest: () => builder.text(chart.title));
                });
              });
            });
          });
          builder.element('c:layout');
        });

        builder.element('c:plotArea', nest: () {
          builder.element('c:layout');
          
          // Chart Type specific element
          builder.element('c:${chart.chartTagName}', nest: () {
            if (chart is ColumnChart) {
              builder.element('c:barDir', attributes: {'val': chart.isVertical ? 'col' : 'bar'});
            }
            
            for (int i = 0; i < chart.series.length; i++) {
              final series = chart.series[i];
              builder.element('c:ser', nest: () {
                builder.element('c:idx', attributes: {'val': '$i'});
                builder.element('c:order', attributes: {'val': '$i'});
                
                // Series Name
                builder.element('c:tx', nest: () {
                  builder.element('c:v', nest: () => builder.text(series.name));
                });

                // Categories (X-Axis)
                builder.element('c:cat', nest: () {
                  builder.element('c:strRef', nest: () {
                    builder.element('c:f', nest: () => builder.text(series.categoriesRange));
                  });
                });

                // Values (Y-Axis)
                builder.element('c:val', nest: () {
                  builder.element('c:numRef', nest: () {
                    builder.element('c:f', nest: () => builder.text(series.valuesRange));
                  });
                });
              });
            }

            // Axes IDs (Dummy IDs for now, usually randomized or sequential)
            builder.element('c:axId', attributes: {'val': '10000001'});
            builder.element('c:axId', attributes: {'val': '10000002'});
          });

          // Category Axis
          builder.element('c:catAx', nest: () {
            builder.element('c:axId', attributes: {'val': '10000001'});
            builder.element('c:scaling', nest: () => builder.element('c:orientation', attributes: {'val': 'minMax'}));
            builder.element('c:axPos', attributes: {'val': 'b'});
            builder.element('c:tickLblPos', attributes: {'val': 'nextTo'});
            builder.element('c:crossAx', attributes: {'val': '10000002'});
            builder.element('c:crosses', attributes: {'val': 'autoZero'});
          });

          // Value Axis
          builder.element('c:valAx', nest: () {
            builder.element('c:axId', attributes: {'val': '10000002'});
            builder.element('c:scaling', nest: () => builder.element('c:orientation', attributes: {'val': 'minMax'}));
            builder.element('c:axPos', attributes: {'val': 'l'});
            builder.element('c:majorGridlines');
            builder.element('c:numFmt', attributes: {'val': 'General', 'sourceLinked': '1'});
            builder.element('c:tickLblPos', attributes: {'val': 'nextTo'});
            builder.element('c:crossAx', attributes: {'val': '10000001'});
            builder.element('c:crosses', attributes: {'val': 'autoZero'});
            builder.element('c:crossBetween', attributes: {'val': 'between'});
          });
        });

        if (chart.showLegend) {
          builder.element('c:legend', nest: () {
            builder.element('c:legendPos', attributes: {'val': 'r'});
            builder.element('c:layout');
          });
        }
        
        builder.element('c:plotVisOnly', attributes: {'val': '1'});
      });
    });
    return builder.buildDocument();
  }
}
