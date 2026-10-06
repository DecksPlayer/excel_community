part of '../../../excel_community.dart';

/// Builder for Line chart styles.
class LineChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    final lineChart = chart as LineChart;
    builder.element('c:grouping',
        attributes: {'val': lineChart.grouping.lineAreaOoxmlValue});
  }

  @override
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final lineChart = chart as LineChart;

    // Line body — no area fill, only border (the line itself).
    ChartColorConfig.buildSpPr(
      builder,
      series,
      seriesIndex,
      ChartColorConfig.seriesPalette,
      includeFill: false,
      defaultBorderAlpha: 100,
      defaultBorderWidth: ChartColorConfig.thickLineWidth,
    );

    // Marker
    final markerColor = ChartColorConfig.resolveFillColor(
        series, seriesIndex, ChartColorConfig.seriesPalette);
    final markerAlpha =
        series.style?.fillType == ChartFillType.none ? 0 : 100;

    if (lineChart.showMarkers) {
      builder.element('c:marker', nest: () {
        builder.element('c:symbol', attributes: {'val': 'circle'});
        builder.element('c:size',
            attributes: {'val': ChartColorConfig.smallMarker});
        builder.element('c:spPr', nest: () {
          ChartColorConfig.emitSolidFill(builder, markerColor, markerAlpha);
          builder.element('a:ln',
              attributes: {'w': ChartColorConfig.thinLineWidth}, nest: () {
            ChartColorConfig.emitSolidFill(builder, markerColor, markerAlpha);
          });
        });
      });
    } else {
      builder.element('c:marker', nest: () {
        builder.element('c:symbol', attributes: {'val': 'none'});
      });
    }

    if (lineChart.smooth) {
      builder.element('c:smooth', attributes: {'val': '1'});
    }
  }
}
