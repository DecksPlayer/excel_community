// Chart snippets: one snippet per supported chart type, plus the
// "all-charts-grid" composite demo.
library;

// ---------------------------------------------------------------------------
// Column chart
// ---------------------------------------------------------------------------
const String columnChartSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateColumnChart() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Add category headers and numerical series data...
  _populateData(sheet);

  // Define ColumnChart
  var chart = ColumnChart(
    title: "Monthly Data (Column)",
    series: [
      ChartSeries(
        name: "Series 1",
        categoriesRange: r"Sheet1!\$A\$2:\$A\$7",
        valuesRange: r"Sheet1!\$B\$2:\$B\$7",
      ),
      ChartSeries(
        name: "Series 2",
        categoriesRange: r"Sheet1!\$A\$2:\$A\$7",
        valuesRange: r"Sheet1!\$C\$2:\$C\$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  );

  sheet.addChart(chart);
  excel.save(fileName: 'column_chart_example.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Line chart
// ---------------------------------------------------------------------------
const String lineChartSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateLineChart() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Populate data...
  _populateData(sheet);

  var chart = LineChart(
    title: "Monthly Data (Line)",
    series: series,
    anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  );

  sheet.addChart(chart);
  excel.save(fileName: 'line_chart_example.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Pie chart
// ---------------------------------------------------------------------------
const String pieChartSnippet = '''
import 'package:excel_community/excel_community.dart';

void generatePieChart() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Populate data...
  _populateData(sheet);

  var chart = PieChart(
    title: "Monthly Data (Pie)",
    series: [
      ChartSeries(
        name: "Series 1",
        categoriesRange: r"Sheet1!\$A\$2:\$A\$7",
        valuesRange: r"Sheet1!\$B\$2:\$B\$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  );

  sheet.addChart(chart);
  excel.save(fileName: 'pie_chart_example.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Area chart
// ---------------------------------------------------------------------------
const String areaChartSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateAreaChart() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Populate data...
  _populateData(sheet);

  var chart = AreaChart(
    title: "Monthly Data (Area)",
    series: series,
    anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  );

  sheet.addChart(chart);
  excel.save(fileName: 'area_chart_example.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Doughnut chart
// ---------------------------------------------------------------------------
const String doughnutChartSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateDoughnutChart() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Populate data...
  _populateData(sheet);

  var chart = DoughnutChart(
    title: "Monthly Data (Doughnut)",
    series: [
      ChartSeries(
        name: "Series 1",
        categoriesRange: r"Sheet1!\$A\$2:\$A\$7",
        valuesRange: r"Sheet1!\$B\$2:\$B\$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  );

  sheet.addChart(chart);
  excel.save(fileName: 'doughnut_chart_example.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Radar chart
// ---------------------------------------------------------------------------
const String radarChartSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateRadarChart() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Populate data...
  _populateData(sheet);

  var chart = RadarChart(
    title: "Monthly Data (Radar)",
    series: series,
    anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
    filled: true, // true for filled area radar, false for line-only grid
  );

  sheet.addChart(chart);
  excel.save(fileName: 'radar_chart_example.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Bar chart
// ---------------------------------------------------------------------------
const String barChartSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateBarChart() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Populate data...
  _populateData(sheet);

  var chart = BarChart(
    title: "Monthly Data (Bar)",
    series: series,
    anchor: ChartAnchor.at(column: 5, row: 1, width: 10, height: 15),
  );

  sheet.addChart(chart);
  excel.save(fileName: 'bar_chart_example.xlsx');
}
''';

// ---------------------------------------------------------------------------
// Scatter chart
// ---------------------------------------------------------------------------
const String scatterChartSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateScatterChart() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Note: Scatter charts require numeric values on both axes
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("X Values"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Y Values"));
  sheet.updateCell(CellIndex.indexByString("A2"), DoubleCellValue(1.5));
  sheet.updateCell(CellIndex.indexByString("B2"), DoubleCellValue(10.0));
  // ... add more data ...

  var chart = ScatterChart(
    title: "XY Correlation Map",
    series: [
      ChartSeries(
        name: "Scatter Point Map",
        categoriesRange: r"Sheet1!\$A\$2:\$A\$8",
        valuesRange: r"Sheet1!\$B\$2:\$B\$8",
      ),
    ],
    anchor: ChartAnchor.at(column: 4, row: 1, width: 10, height: 15),
  );

  sheet.addChart(chart);
  excel.save(fileName: 'scatter_chart_example.xlsx');
}
''';

// ---------------------------------------------------------------------------
// All charts grid
// ---------------------------------------------------------------------------
const String allChartsSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateAllCharts() {
  var excel = Excel.createExcel();
  var sheet = excel['All Charts Grid'];

  // Populate data...
  _populateData(sheet);

  // Add various charts anchored at different column/row coordinates
  sheet.addChart(ColumnChart(
    title: "Columns",
    series: series,
    anchor: ChartAnchor.at(column: 5, row: 1, width: 8, height: 12),
  ));

  sheet.addChart(BarChart(
    title: "Bars",
    series: series,
    anchor: ChartAnchor.at(column: 14, row: 1, width: 8, height: 12),
  ));

  sheet.addChart(LineChart(
    title: "Lines",
    series: series,
    anchor: ChartAnchor.at(column: 5, row: 14, width: 8, height: 12),
  ));

  sheet.addChart(AreaChart(
    title: "Areas",
    series: series,
    anchor: ChartAnchor.at(column: 14, row: 14, width: 8, height: 12),
  ));

  excel.save(fileName: 'all_charts_grid_demo.xlsx');
}
''';