part of '../../../excel_community.dart';

/// Builder for Stock chart styles (HLC, OHLC, etc.).
class StockChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    final stockChart = chart as StockChart;
    if (stockChart.showHighLowLines) {
      builder.element('c:hiLowLines');
    }
    if (stockChart.showUpDownBars) {
      builder.element('c:upDownBars', nest: () {
        builder.element('c:gapWidth', attributes: {'val': '150'});
        builder.element('c:upBars');
        builder.element('c:downBars');
      });
    }
  }

  @override
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    // Stock charts use line properties without fill for open/high/low/close series
    ChartColorConfig.buildSpPr(
      builder,
      series,
      seriesIndex,
      ChartColorConfig.seriesPalette,
      includeFill: false,
      defaultBorderAlpha: 100,
      defaultBorderWidth: ChartColorConfig.thinLineWidth,
    );
  }
}
