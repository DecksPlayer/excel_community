part of excel_community;

/// Builder for Column and Bar chart styles
class ColumnBarChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    if (chart is ColumnChart) {
      builder.element('c:barDir', attributes: {'val': chart.isVertical ? 'col' : 'bar'});
    } else {
      builder.element('c:barDir', attributes: {'val': 'bar'});
    }
    builder.element('c:grouping', attributes: {'val': 'clustered'});
  }

  @override
  void buildSeriesStyle(XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final color = ChartColorConfig.getSeriesColor(seriesIndex).colorHex6;
    
    builder.element('c:spPr', nest: () {
      // Solid fill
      builder.element('a:solidFill', nest: () {
        builder.element('a:srgbClr', attributes: {'val': color});
      });
      // Border with same color
      builder.element('a:ln', attributes: {'w': ChartColorConfig.thinLineWidth}, nest: () {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color});
        });
      });
    });
  }
}
