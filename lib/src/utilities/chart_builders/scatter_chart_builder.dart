part of '../../../excel_community.dart';

/// Builder for Scatter (XY) chart styles.
class ScatterChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    final scatterChart = chart as ScatterChart;
    builder.element('c:scatterStyle',
        attributes: {'val': scatterChart.showLines ? 'lineMarker' : 'marker'});
  }

  @override
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final scatterChart = chart as ScatterChart;
    final fillColor = ChartColorConfig.resolveFillColor(
        series, seriesIndex, ChartColorConfig.seriesPalette);

    final borderColor =
        series.style?.borderColor ?? ChartColorConfig.white;
    final borderAlpha = series.style?.borderAlpha ?? 100;
    final fillAlpha = series.style?.fillType == ChartFillType.transparent
        ? series.style!.fillAlpha
        : 100;
    final noFill = series.style?.fillType == ChartFillType.none;
    final borderWidth =
        series.style?.borderWidth ?? ChartColorConfig.thinLineWidth;

    // Optional connecting line
    if (scatterChart.showLines) {
      final lineColor = series.style?.borderColor ?? fillColor;
      builder.element('c:spPr', nest: () {
        builder.element('a:ln',
            attributes: {'w': ChartColorConfig.thickLineWidth}, nest: () {
          ChartColorConfig.emitSolidFill(builder, lineColor, borderAlpha);
        });
      });
    }

    // Marker
    if (scatterChart.showMarkers) {
      builder.element('c:marker', nest: () {
        builder.element('c:symbol', attributes: {'val': 'circle'});
        builder.element('c:size',
            attributes: {'val': ChartColorConfig.smallMarker});
        builder.element('c:spPr', nest: () {
          if (noFill) {
            builder.element('a:noFill');
          } else {
            ChartColorConfig.emitSolidFill(builder, fillColor, fillAlpha);
          }
          builder.element('a:ln', attributes: {'w': borderWidth}, nest: () {
            ChartColorConfig.emitSolidFill(builder, borderColor, borderAlpha);
          });
        });
      });
    } else {
      builder.element('c:marker', nest: () {
        builder.element('c:symbol', attributes: {'val': 'none'});
      });
    }

    if (scatterChart.smooth) {
      builder.element('c:smooth', attributes: {'val': '1'});
    }
  }
}
