part of excel_community;

/// Represents an Excel Column/Bar Chart.
class ColumnChart extends Chart {
  final bool isVertical;

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
  final bool filled;

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
  BarChart({
    required super.title,
    required super.series,
    required super.anchor,
    super.showLegend,
  });

  @override
  String get chartTagName => 'barChart';
}
