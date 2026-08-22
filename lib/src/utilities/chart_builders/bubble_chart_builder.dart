part of '../../../excel_community.dart';

/// Builder for Bubble chart styles.
class BubbleChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    final bubbleChart = chart as BubbleChart;
    builder.element('c:varyColors', attributes: {'val': '1'});
    builder.element('c:bubbleScale',
        attributes: {'val': bubbleChart.bubbleScale.toString()});
    builder.element('c:showNegBubbles',
        attributes: {'val': bubbleChart.showNegativeBubbles ? '1' : '0'});
  }

  @override
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    ChartColorConfig.buildSpPr(
      builder,
      series,
      seriesIndex,
      ChartColorConfig.seriesPalette,
      defaultFillAlpha: 60,
      defaultBorderAlpha: 100,
      defaultBorderWidth: ChartColorConfig.thinLineWidth,
    );
  }
}
