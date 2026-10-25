part of '../../../excel_community.dart';

/// Builder for Column and Bar chart styles
class ColumnBarChartBuilder extends ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    ChartGrouping grouping = ChartGrouping.clustered;
    if (chart is ColumnChart) {
      builder.element('c:barDir',
          attributes: {'val': chart.isVertical ? 'col' : 'bar'});
      grouping = chart.grouping;
    } else if (chart is BarChart) {
      builder.element('c:barDir', attributes: {'val': 'bar'});
      grouping = chart.grouping;
    }

    builder.element('c:grouping', attributes: {'val': grouping.ooxmlValue});
  }

  @override
  void buildPropertiesAfterSeries(XmlBuilder builder, Chart chart) {
    final grouping = switch (chart) {
      ColumnChart() => chart.grouping,
      BarChart() => chart.grouping,
      _ => ChartGrouping.clustered,
    };
    if (grouping != ChartGrouping.clustered) {
      builder.element('c:overlap', attributes: {'val': '100'});
    }
  }

  @override
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    ChartColorConfig.buildSpPr(
      builder,
      series,
      seriesIndex,
      ChartColorConfig.seriesPalette,
      defaultFillAlpha: 100,
      defaultBorderAlpha: 100,
      defaultBorderWidth: ChartColorConfig.thinLineWidth,
    );
  }
}
