part of '../../excel_community.dart';

/// Grouping mode for Column, Bar, Area, and Line charts.
enum ChartGrouping {
  /// Clustered / standard side-by-side or unstacked display.
  clustered('clustered'),

  /// Stacked series displaying the cumulative total.
  stacked('stacked'),

  /// 100% Stacked series comparing percentage contribution to the whole.
  percentStacked('percentStacked');

  final String ooxmlValue;
  const ChartGrouping(this.ooxmlValue);
}

/// Subtype for [OfPieChart] (Pie-of-Pie or Bar-of-Pie).
enum OfPieType {
  /// Pie chart with a secondary pie breakdown.
  pie('pie'),

  /// Pie chart with a secondary stacked bar breakdown.
  bar('bar');

  final String ooxmlValue;
  const OfPieType(this.ooxmlValue);
}

/// Method used to split data slices into the secondary pie/bar chart.
enum OfPieSplitType {
  /// Split by position (e.g., last N slices moved to secondary).
  position('pos'),

  /// Split by threshold value (slices with values less than X moved).
  value('val'),

  /// Split by percentage threshold (slices with < X% of total moved).
  percent('percent'),

  /// Split manually by custom slice assignment.
  custom('cust');

  final String ooxmlValue;
  const OfPieSplitType(this.ooxmlValue);
}
