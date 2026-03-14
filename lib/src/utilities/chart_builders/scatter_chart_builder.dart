part of excel_community;

/// Builder for Scatter chart styles
class ScatterChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    // Scatter charts don't have grouping properties
  }

  @override
  void buildSeriesStyle(XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final color = ChartColorConfig.getSeriesColor(seriesIndex);
    
    // Point fill
    builder.element('c:spPr', nest: () {
      builder.element('a:solidFill', nest: () {
        builder.element('a:srgbClr', attributes: {'val': color});
      });
    });
    
    // Marker with white border
    builder.element('c:marker', nest: () {
      builder.element('c:symbol', attributes: {'val': 'circle'});
      builder.element('c:size', attributes: {'val': ChartColorConfig.largeMarker});
      builder.element('c:spPr', nest: () {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color});
        });
        builder.element('a:ln', attributes: {'w': ChartColorConfig.thinLineWidth}, nest: () {
          builder.element('a:solidFill', nest: () {
            builder.element('a:srgbClr', attributes: {'val': ChartColorConfig.white});
          });
        });
      });
    });
  }
}
