part of '../../../excel_community.dart';

/// Builder for Radar chart styles with optional fill.
class RadarChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    final radarChart = chart as RadarChart;
    builder.element('c:radarStyle',
        attributes: {'val': radarChart.filled ? 'filled' : 'marker'});
  }

  @override
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final radarChart = chart as RadarChart;
    ChartColorConfig.buildSpPr(
      builder,
      series,
      seriesIndex,
      ChartColorConfig.radarPalette,
      // Filled radar default: 45 % fill, 85 % border.
      // Non-filled radar: no fill (only border line).
      includeFill: radarChart.filled,
      defaultFillAlpha: 45,
      defaultBorderAlpha: 85,
      defaultBorderWidth: ChartColorConfig.thickLineWidth,
    );
  }
}
