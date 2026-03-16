part of excel_community;

/// Builder for Area chart styles with transparency.
class AreaChartBuilder implements ChartStyleBuilder {
  /// Creates a new AreaChartBuilder.
  AreaChartBuilder();

  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    builder.element('c:grouping', attributes: {'val': 'standard'});
  }

  @override
  void buildSeriesStyle(XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final color = ChartColorConfig.getSeriesColor(seriesIndex).colorHex6;
    
    builder.element('c:spPr', nest: () {
      // Area fill
      builder.element('a:solidFill', nest: () {
        builder.element('a:srgbClr', attributes: {'val': color}, nest: () {
          builder.element('a:alpha', attributes: {'val': ChartColorConfig.opacity50});
        });
      });
      // Border line
      builder.element('a:ln', attributes: {'w': ChartColorConfig.thickLineWidth}, nest: () {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color}, nest: () {
            builder.element('a:alpha', attributes: {'val': ChartColorConfig.opacity90});
          });
        });
      });
    });
  }
}
