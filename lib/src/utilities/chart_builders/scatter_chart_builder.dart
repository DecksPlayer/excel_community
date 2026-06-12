part of '../../../excel_community.dart';

/// Builder for Scatter chart styles
class ScatterChartBuilder implements ChartStyleBuilder {
  @override
  void buildProperties(XmlBuilder builder, Chart chart) {
    // scatterStyle is required for Excel to render scatter charts correctly
    builder.element('c:scatterStyle', attributes: {'val': 'marker'});
  }

  @override
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {
    final color = ChartColorConfig.getSeriesColor(seriesIndex).colorHex6;

    // Marker style for scatter series
    builder.element('c:marker', nest: () {
      builder.element('c:symbol', attributes: {'val': 'circle'});
      builder
          .element('c:size', attributes: {'val': ChartColorConfig.smallMarker});
      builder.element('c:spPr', nest: () {
        builder.element('a:solidFill', nest: () {
          builder.element('a:srgbClr', attributes: {'val': color});
        });
        builder.element('a:ln',
            attributes: {'w': ChartColorConfig.thinLineWidth}, nest: () {
          builder.element('a:solidFill', nest: () {
            builder.element('a:srgbClr',
                attributes: {'val': ChartColorConfig.white.colorHex6});
          });
        });
      });
    });
  }
}
