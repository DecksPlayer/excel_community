part of excel_community;

/// Builder for Radar chart styles with optional fill
class RadarChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    final radarChart = chart as RadarChart;
    builder.element('c:radarStyle', attributes: {'val': radarChart.filled ? 'filled' : 'marker'});
  }

  @override
  void buildSeriesStyle(XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final radarChart = chart as RadarChart;
    final color = ChartColorConfig.getRadarColor(seriesIndex).colorHex6;
    
    builder.element('c:spPr', nest: () {
      // Line formatting
      builder.element('a:ln', attributes: {'w': ChartColorConfig.thickLineWidth}, nest: () {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color});
        });
      });
      // Marker formatting
      builder.element('a:solidFill', nest: () {
        builder.element('a:srgbClr', attributes: {'val': color});
      });
      
      // Fill with 45% opacity (only if filled style)
      if (radarChart.filled) {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color}, nest: () {
            builder.element('a:alpha', attributes: {'val': ChartColorConfig.opacity45});
          });
        });
      }
    });
  }
}
