part of excel_community;

/// Represents an Excel Column/Bar Chart.
class ColumnChart extends Chart {
  /// Whether the bars are vertical (true) or horizontal (false).
  final bool isVertical;

  /// Creates a new ColumnChart.
  ColumnChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
    this.isVertical = true,
  });

  @override
  String get chartTagName => 'barChart';
}

/// Represents an Excel Line Chart.
class LineChart extends Chart {
  /// Creates a new LineChart.
  LineChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
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
  });

  @override
  String get chartTagName => 'pieChart';
}

/// Represents an Excel Scatter Chart.
class ScatterChart extends Chart {
  /// Creates a new ScatterChart.
  ScatterChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
  });

  @override
  String get chartTagName => 'scatterChart';
}

/// Represents an Excel Area Chart.
class AreaChart extends Chart {
  /// Creates a new AreaChart.
  AreaChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
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
    this.filled = false,
  });

  @override
  String get chartTagName => 'radarChart';
}

/// Represents an Excel Bar Chart (Horizontal bars).
/// Note: For vertical bars, use ColumnChart with isVertical=true.
class BarChart extends Chart {
  /// Creates a new BarChart.
  BarChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
  });

  @override
  String get chartTagName => 'barChart';
}
