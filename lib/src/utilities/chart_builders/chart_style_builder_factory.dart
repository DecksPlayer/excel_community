part of '../../../excel_community.dart';

/// Factory to get the appropriate style builder for a chart type
class ChartStyleBuilderFactory {
  /// Returns the appropriate builder for the given chart type
  static ChartStyleBuilder getBuilder(Chart chart) {
    if (chart is ColumnChart || chart is BarChart) {
      return ColumnBarChartBuilder();
    } else if (chart is LineChart) {
      return LineChartBuilder();
    } else if (chart is AreaChart) {
      return AreaChartBuilder();
    } else if (chart is ScatterChart) {
      return ScatterChartBuilder();
    } else if (chart is BubbleChart) {
      return BubbleChartBuilder();
    } else if (chart is StockChart) {
      return StockChartBuilder();
    } else if (chart is OfPieChart) {
      return OfPieChartBuilder();
    } else if (chart is PieChart || chart is DoughnutChart) {
      return PieChartBuilder();
    } else if (chart is RadarChart) {
      return RadarChartBuilder();
    }

    // Default fallback
    return ColumnBarChartBuilder();
  }
}
