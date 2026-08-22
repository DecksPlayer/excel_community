part of '../../excel_community.dart';

/// Represents an Excel Column/Bar Chart.
class ColumnChart extends Chart {
  /// Whether the bars are vertical (true) or horizontal (false).
  final bool isVertical;

  /// Grouping mode: [ChartGrouping.clustered], [ChartGrouping.stacked], or [ChartGrouping.percentStacked].
  final ChartGrouping grouping;

  /// Creates a new ColumnChart.
  ColumnChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
    this.isVertical = true,
    this.grouping = ChartGrouping.clustered,
  });

  @override
  String get chartTagName => 'barChart';
}

/// Represents an Excel Line Chart.
class LineChart extends Chart {
  /// Grouping mode: [ChartGrouping.clustered], [ChartGrouping.stacked], or [ChartGrouping.percentStacked].
  final ChartGrouping grouping;

  /// Whether to display circular markers on data points (default: true).
  final bool showMarkers;

  /// Whether to render the series line with smooth curves (default: false).
  final bool smooth;

  /// Creates a new LineChart.
  LineChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
    this.grouping = ChartGrouping.clustered,
    this.showMarkers = true,
    this.smooth = false,
  });

  @override
  String get chartTagName => 'lineChart';
}

/// Represents an Excel Pie Chart.
class PieChart extends Chart {
  /// Creates a new PieChart.
  PieChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
  });

  @override
  String get chartTagName => 'pieChart';
}

/// Represents an Excel Scatter (XY) Chart.
class ScatterChart extends Chart {
  /// Whether to draw connecting lines between data points (default: false).
  final bool showLines;

  /// Whether to display data markers (default: true).
  final bool showMarkers;

  /// Whether to draw connecting lines with smooth curves (default: false).
  final bool smooth;

  /// Creates a new ScatterChart.
  ScatterChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
    this.showLines = false,
    this.showMarkers = true,
    this.smooth = false,
  });

  @override
  String get chartTagName => 'scatterChart';
}

/// Represents an Excel Area Chart.
class AreaChart extends Chart {
  /// Grouping mode: [ChartGrouping.clustered], [ChartGrouping.stacked], or [ChartGrouping.percentStacked].
  final ChartGrouping grouping;

  /// Creates a new AreaChart.
  AreaChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
    this.grouping = ChartGrouping.clustered,
  });

  @override
  String get chartTagName => 'areaChart';
}

/// Represents an Excel Doughnut Chart (Pie chart with a hole).
class DoughnutChart extends Chart {
  /// Creates a new DoughnutChart.
  DoughnutChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
  });

  @override
  String get chartTagName => 'doughnutChart';
}

/// Represents an Excel Radar Chart.
class RadarChart extends Chart {
  /// Whether the radar areas are filled.
  final bool filled;

  /// Creates a new RadarChart.
  RadarChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
    this.filled = false,
  });

  @override
  String get chartTagName => 'radarChart';
}

/// Represents an Excel Bar Chart (Horizontal bars).
/// Note: For vertical bars, use ColumnChart with isVertical=true.
class BarChart extends Chart {
  /// Grouping mode: [ChartGrouping.clustered], [ChartGrouping.stacked], or [ChartGrouping.percentStacked].
  final ChartGrouping grouping;

  /// Creates a new BarChart.
  BarChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
    this.grouping = ChartGrouping.clustered,
  });

  @override
  String get chartTagName => 'barChart';
}

/// Represents an Excel Bubble Chart (3D data comparison: X, Y, and Bubble Size).
class BubbleChart extends Chart {
  /// Scale of the bubbles as a percentage (default: 100).
  final int bubbleScale;

  /// Whether negative bubbles are shown (default: false).
  final bool showNegativeBubbles;

  /// Creates a new BubbleChart.
  BubbleChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
    this.bubbleScale = 100,
    this.showNegativeBubbles = false,
  });

  @override
  String get chartTagName => 'bubbleChart';
}

/// Represents an Excel Stock Chart (High-Low-Close, Open-High-Low-Close, etc.).
class StockChart extends Chart {
  /// Whether to display High-Low vertical lines connecting extremes.
  final bool showHighLowLines;

  /// Whether to display Up-Down bars indicating price direction.
  final bool showUpDownBars;

  /// Creates a new StockChart.
  StockChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
    this.showHighLowLines = true,
    this.showUpDownBars = true,
  });

  @override
  String get chartTagName => 'stockChart';
}

/// Represents an Excel Pie-of-Pie or Bar-of-Pie Chart.
class OfPieChart extends Chart {
  /// Subtype: [OfPieType.pie] (Pie of Pie) or [OfPieType.bar] (Bar of Pie).
  final OfPieType ofPieType;

  /// How data points are split into the secondary chart.
  final OfPieSplitType splitType;

  /// Split position threshold (e.g., number of slices or value limit).
  final int? splitPosition;

  /// Relative size of the secondary pie/bar as a percentage (default: 75).
  final int secondPieSize;

  /// Creates a new OfPieChart.
  OfPieChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    super.dataLabels,
    this.ofPieType = OfPieType.pie,
    this.splitType = OfPieSplitType.position,
    this.splitPosition = 2,
    this.secondPieSize = 75,
  });

  @override
  String get chartTagName => 'ofPieChart';
}
