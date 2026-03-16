part of excel_community;

/// Builder for Scatter chart styles
class ScatterChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    // Scatter charts don't have grouping properties
  }

  @override
  void buildSeriesStyle(XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final color = ChartColorConfig.getSeriesColor(seriesIndex).colorHex6;
    
    builder.element('c:spPr', nest: () {
      // No line for scatter points by default
      builder.element('a:ln', nest: () {
        builder.element('a:noFill');
      });
      // Marker fill
      builder.element('a:solidFill', nest: () {
        builder.element('a:srgbClr', attributes: {'val': color});
      });
      // Marker border
      builder.element('a:ln', attributes: {'w': ChartColorConfig.thinLineWidth}, nest: () {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color});
        });
      });
    });
    
    // Marker style
    builder.element('c:marker', nest: () {
      builder.element('c:symbol', attributes: {'val': 'circle'});
      builder.element('c:size', attributes: {'val': ChartColorConfig.smallMarker});
      builder.element('c:spPr', nest: () {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color});
        });
        builder.element('a:ln', attributes: {'w': ChartColorConfig.thinLineWidth}, nest: () {
          builder.element('a:solidFill', nest: () {
            builder.element('a:srgbClr', attributes: {'val': ChartColorConfig.white.colorHex6});
          });
        });
      });
    });
  }
}
