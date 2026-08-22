part of '../../excel_community.dart';

/// Base class for all Excel charts.
abstract class Chart {
  final String title;
  final List<ChartSeries> series;
  final ChartAnchor anchor;
  final bool showLegend;

  /// Optional data labels shown on every data point.
  ///
  /// When `null` (the default), no `<c:dLbls>` element is written and Excel
  /// uses its default (no labels). Supply a [ChartDataLabels] instance to
  /// enable one or more label components (value, category name, series name,
  /// percentage).
  final ChartDataLabels? dataLabels;

  Chart({
    required this.title,
    required this.series,
    required this.anchor,
    this.showLegend = true,
    this.dataLabels,
  });

  /// The XML element name for this chart type in ChartML (e.g., 'barChart', 'lineChart').
  String get chartTagName;
}

/// Represents a single data series in a chart.
class ChartSeries {
  final String name;
  final String categoriesRange; // e.g., "Sheet1!$A$2:$A$10"
  final String valuesRange; // e.g., "Sheet1!$B$2:$B$10"

  /// Optional cached data for categories (labels for bar/line/pie) or X values (scatter)
  List<String>? categories;

  /// Optional cached data for values (Y axis)
  List<num>? values;

  /// Optional cached X numeric values for ScatterChart and BubbleChart
  List<num>? xValues;

  /// Optional range reference for bubble size values in BubbleChart (e.g. "Sheet1!$C$2:$C$10")
  final String? bubbleSizeRange;

  /// Optional cached numeric values for bubble sizes in BubbleChart
  List<num>? bubbleSizes;

  /// Optional per-series visual styling (fill color, transparency, border).
  ///
  /// When `null`, the chart builder falls back to the automatic palette color
  /// for this series index. Supply a [ChartSeriesStyle] to override any or
  /// all visual properties.
  final ChartSeriesStyle? style;

  ChartSeries({
    required this.name,
    required this.categoriesRange,
    required this.valuesRange,
    this.categories,
    this.values,
    this.xValues,
    this.bubbleSizeRange,
    this.bubbleSizes,
    this.style,
  });
}

/// Defines the position and size of a chart on the worksheet.
class ChartAnchor {
  final int fromColumn;
  final int fromRow;
  final int toColumn;
  final int toRow;

  ChartAnchor({
    required this.fromColumn,
    required this.fromRow,
    required this.toColumn,
    required this.toRow,
  });

  factory ChartAnchor.at({
    required int column,
    required int row,
    int width = 8,
    int height = 15,
  }) {
    return ChartAnchor(
      fromColumn: column,
      fromRow: row,
      toColumn: column + width,
      toRow: row + height,
    );
  }
}
