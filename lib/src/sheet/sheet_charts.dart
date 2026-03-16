part of excel_community;

extension SheetCharts on Sheet {
  /// returns the `charts` list
  List<Chart> get charts => List.unmodifiable(_charts);

  /// Adds a chart to the sheet.
  void addChart(Chart chart) {
    _charts.add(chart);
  }
}
