part of excel;

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
