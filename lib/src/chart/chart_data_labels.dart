part of '../../excel_community.dart';

/// Configuration for data labels displayed on chart series.
///
/// Data labels annotate each data point with one or more of:
/// - the numeric [value]
/// - the [categoryName] (X-axis label)
/// - the [seriesName]
/// - the [percentage] (pie / doughnut only)
///
/// Example – show values and percentages on a pie chart:
/// ```dart
/// PieChart(
///   title: 'Market Share',
///   dataLabels: ChartDataLabels(value: true, percentage: true),
///   series: [ ... ],
///   anchor: ChartAnchor.at(column: 2, row: 2),
/// )
/// ```
class ChartDataLabels {
  /// Show the numeric value of each data point.
  final bool value;

  /// Show the category name (X-axis label) for each data point.
  final bool categoryName;

  /// Show the series name for each data point.
  final bool seriesName;

  /// Show the percentage share (pie / doughnut charts only).
  final bool percentage;

  /// Separator inserted between multiple label parts (default: ", ").
  final String separator;

  /// Position of the label relative to its data point.
  ///
  /// Common values: `'bestFit'`, `'ctr'`, `'inBase'`, `'inEnd'`,
  /// `'l'`, `'outEnd'`, `'r'`, `'t'`, `'b'`.
  /// When `null`, Excel chooses a default position.
  final String? labelPosition;

  const ChartDataLabels({
    this.value = false,
    this.categoryName = false,
    this.seriesName = false,
    this.percentage = false,
    this.separator = ', ',
    this.labelPosition,
  });

  /// Returns `true` when at least one label component is enabled.
  bool get isEnabled =>
      value || categoryName || seriesName || percentage;

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is ChartDataLabels &&
          other.value == value &&
          other.categoryName == categoryName &&
          other.seriesName == seriesName &&
          other.percentage == percentage &&
          other.separator == separator &&
          other.labelPosition == labelPosition;

  @override
  int get hashCode => Object.hash(
        value,
        categoryName,
        seriesName,
        percentage,
        separator,
        labelPosition,
      );
}
