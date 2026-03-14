part of excel_community;

/// Base interface for chart-specific styling builders.
/// 
/// Each chart type has its own builder that knows how to render
/// its specific visual properties (colors, markers, transparency, etc.)
abstract class ChartStyleBuilder {
  /// Builds the type-specific properties for the chart element
  void buildProperties(XmlBuilder builder, Chart chart);

  /// Builds the color styling for a series
  void buildSeriesStyle(XmlBuilder builder, Chart chart, ChartSeries series, int seriesIndex);
}
