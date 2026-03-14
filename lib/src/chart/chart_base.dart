part of excel;

/// Base class for all Excel charts.
abstract class Chart {
  final String title;
  final List<ChartSeries> series;
  final ChartAnchor anchor;
  final bool showLegend;

  Chart({
    required this.title,
    required this.series,
    required this.anchor,
    this.showLegend = true,
  });

  /// The XML element name for this chart type in ChartML (e.g., 'barChart', 'lineChart').
  String get chartTagName;
}

/// Represents a single data series in a chart.
class ChartSeries {
  final String name;
  final String categoriesRange; // e.g., "Sheet1!$A$2:$A$10"
  final String valuesRange;     // e.g., "Sheet1!$B$2:$B$10"

  ChartSeries({
    required this.name,
    required this.categoriesRange,
    required this.valuesRange,
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
