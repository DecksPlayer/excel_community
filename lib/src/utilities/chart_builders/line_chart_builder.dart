part of excel_community;

/// Builder for Line chart styles
class LineChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    builder.element('c:grouping', attributes: {'val': 'standard'});
  }

  @override
  void buildSeriesStyle(XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final color = ChartColorConfig.getSeriesColor(seriesIndex).colorHex6;
    
    // Line properties
    builder.element('c:spPr', nest: () {
      builder.element('a:ln', attributes: {'w': ChartColorConfig.thickLineWidth}, nest: () {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color});
        });
      });
    });
    
    // Marker properties
    builder.element('c:marker', nest: () {
      builder.element('c:symbol', attributes: {'val': 'circle'});
      builder.element('c:size', attributes: {'val': ChartColorConfig.smallMarker});
      builder.element('c:spPr', nest: () {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color});
        });
      });
    });
  }
}
