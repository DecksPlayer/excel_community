part of '../../../excel_community.dart';

/// Base interface for chart-specific styling builders.
///
/// Each chart type has its own builder that knows how to render
/// its specific visual properties (colors, markers, transparency, etc.)
abstract class ChartStyleBuilder {
  /// Builds the type-specific properties that come before the series.
  void buildProperties(XmlBuilder builder, Chart chart);

  /// Builds the type-specific properties that ECMA-376 places after the
  /// series and data labels (e.g. `c:gapWidth`, `c:firstSliceAng`).
  void buildPropertiesAfterSeries(XmlBuilder builder, Chart chart) {}

  /// Builds the color styling for a series
  void buildSeriesStyle(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex);

  /// Builds the series elements that come after its data (e.g. `c:smooth`).
  void buildSeriesEnd(
      XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex) {}
}
