const String simpleExcelSnippet = '''
import 'package:excel_community/excel_community.dart';

void main() {
  // Create an Excel workbook
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Update cell values using CellIndex
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Name"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Age"));
  sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Alice"));
  sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(30));
  sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Bob"));
  sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(25));

  // Save the workbook
  var bytes = excel.save(fileName: 'simple_no_chart.xlsx');
}
''';

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

const String imageEmbeddingSnippet = '''
import 'dart:convert';
import 'package:excel_community/excel_community.dart';

void embedImage() {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Decode image bytes (e.g. from file, internet, or base64)
  var imageBytes = base64Decode("iVBORw0KGgoAAAANSUhEUgAAADIAAAAy..."); 

  sheet.addImage(ExcelImage(
    imageBytes: imageBytes,
    imageType: ExcelImageType.png,
    anchor: ImageAnchor.fromPixels(
      column: 1, row: 4, // B5 (0-indexed column B = 1, row 5 = 4)
      widthPixels: 150,
      heightPixels: 150,
    ),
  ));

  excel.save(fileName: 'image_example.xlsx');
}
''';

const String textStylesSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateTextStyles() {
  var excel = Excel.createExcel();
  var sheet = excel['Styles Demo'];

  // Define CellStyle configurations
  var cellStyle = CellStyle(
    bold: true,
    italic: true,
    underline: Underline.Double, // Double underline
    strikethrough: true,        // Strikethrough line
    fontSize: 12,
    fontColorHex: ExcelColor.fromHexString('#FF0000'),       // Red text
    backgroundColorHex: ExcelColor.fromHexString('#CCCCCC'), // Grey fill
  );

  var cell = sheet.cell(CellIndex.indexByString("B5"));
  cell.value = TextCellValue("Styled Text Example");
  cell.cellStyle = cellStyle;

  excel.save(fileName: 'text_styles_demo.xlsx');
}
''';

const String numberFormatsSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateNumberFormats() {
  var excel = Excel.createExcel();
  var sheet = excel['Number Formatting'];

  // Apply numberFormat to CellStyle
  var currencyStyle = CellStyle(
    numberFormat: NumFormat.standard_44, // Currency Accounting Indent fmt
    horizontalAlign: HorizontalAlign.Right,
  );

  var cell = sheet.cell(CellIndex.indexByString("C3"));
  cell.value = DoubleCellValue(1234.56);
  cell.cellStyle = currencyStyle;

  // Custom Formats can also be declared:
  var customStyle = CellStyle(
    numberFormat: CustomNumericNumFormat('#,##0.00 \\m\\²'),
  );

  excel.save(fileName: 'number_formats_demo.xlsx');
}
''';

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

const String fullDemoSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateFullReport() {
  var excel = Excel.createExcel();
  var sheetName = 'Financial Report';
  excel.rename('Sheet1', sheetName);
  var sheet = excel[sheetName];

  // Define headers, borders, colors...
  var headerStyle = CellStyle(
    bold: true,
    fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
    backgroundColorHex: ExcelColor.fromHexString('#4472C4'),
  );

  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Month"), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Revenue"), cellStyle: headerStyle);
  
  // Calculate Profit with Excel Formulas:
  sheet.updateCell(CellIndex.indexByString("D2"), FormulaCellValue("B2-C2"));
  
  // Summing Totals:
  sheet.updateCell(CellIndex.indexByString("B8"), FormulaCellValue("SUM(B2:B7)"));

  // Add ColumnChart comparing revenue/profit
  sheet.addChart(ColumnChart(
    title: "Monthly Financial Performance",
    series: [
      ChartSeries(name: "Revenue", categoriesRange: "'Financial Report'!\$A\$2:\$A\$7", valuesRange: "'Financial Report'!\$B\$2:\$B\$7"),
      ChartSeries(name: "Profit", categoriesRange: "'Financial Report'!\$A\$2:\$A\$7", valuesRange: "'Financial Report'!\$D\$2:\$D\$7"),
    ],
    anchor: ChartAnchor.at(column: 6, row: 1, width: 10, height: 15),
  ));

  excel.save(fileName: 'financial_performance_report.xlsx');
}
''';
