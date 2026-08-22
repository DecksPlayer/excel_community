part of '../../../excel_community.dart';

/// Builder for Area chart styles with transparency and grouping.
class AreaChartBuilder implements ChartStyleBuilder {
  AreaChartBuilder();

  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    final areaChart = chart as AreaChart;
    builder.element('c:grouping',
        attributes: {'val': areaChart.grouping.ooxmlValue});
  }

  @override
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    ChartColorConfig.buildSpPr(
      builder,
      series,
      seriesIndex,
      ChartColorConfig.seriesPalette,
      // Area default: 50 % transparent fill, 90 % opaque border
      defaultFillAlpha: 50,
      defaultBorderAlpha: 90,
      defaultBorderWidth: ChartColorConfig.thickLineWidth,
    );
  }
}
