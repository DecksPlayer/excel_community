import 'dart:io';
import 'package:excel_community/excel.dart';

/// Example demonstrating how to create charts in Excel files
void main() {
  print('📊 Creating Excel file with various chart types...\n');

  var excel = Excel.createExcel();
  var sheet = excel['ChartsDemo'];

  // Add sample data
  print('Adding sample data...');
  _addSampleData(sheet);

  // Create different chart types
  print('Creating Column Chart...');
  _addColumnChart(sheet);

  print('Creating Line Chart...');
  _addLineChart(sheet);

  print('Creating Pie Chart...');
  _addPieChart(sheet);

  print('Creating Area Chart...');
  _addAreaChart(sheet);

  // Save the file
  print('\nSaving Excel file...');
  var fileBytes = excel.save();
  
  if (fileBytes != null) {
    File('charts_demo.xlsx')
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
    
    print('✅ File saved successfully as charts_demo.xlsx');
    print('\nThe file contains:');
    print('  • Sample data in columns A-D');
    print('  • Column chart showing monthly comparison');
    print('  • Line chart showing trends');
    print('  • Pie chart showing distribution');
    print('  • Area chart showing cumulative values');
  } else {
    print('❌ Failed to save file');
  }
}

/// Add sample data to the sheet
void _addSampleData(Sheet sheet) {
  // Headers
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Month"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Product A"));
  sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Product B"));
  sheet.updateCell(CellIndex.indexByString("D1"), TextCellValue("Product C"));

  // Data
  final months = ['January', 'February', 'March', 'April', 'May', 'June'];
  final productA = [45, 55, 42, 60, 58, 65];
  final productB = [35, 48, 52, 45, 62, 55];
  final productC = [50, 40, 58, 52, 48, 60];

  for (var i = 0; i < months.length; i++) {
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
      TextCellValue(months[i]),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
      IntCellValue(productA[i]),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1),
      IntCellValue(productB[i]),
    );
    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i + 1),
      IntCellValue(productC[i]),
    );
  }
}

/// Create a Column Chart
void _addColumnChart(Sheet sheet) {
  var chart = ColumnChart(
    title: "Monthly Sales Comparison",
    series: [
      ChartSeries(
        name: "Product A",
        categoriesRange: r"ChartsDemo!$A$2:$A$7",
        valuesRange: r"ChartsDemo!$B$2:$B$7",
      ),
      ChartSeries(
        name: "Product B",
        categoriesRange: r"ChartsDemo!$A$2:$A$7",
        valuesRange: r"ChartsDemo!$C$2:$C$7",
      ),
      ChartSeries(
        name: "Product C",
        categoriesRange: r"ChartsDemo!$A$2:$A$7",
        valuesRange: r"ChartsDemo!$D$2:$D$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 0, width: 12, height: 16),
    showLegend: true,
  );

  sheet.addChart(chart);
}

/// Create a Line Chart
void _addLineChart(Sheet sheet) {
  var chart = LineChart(
    title: "Sales Trends Over Time",
    series: [
      ChartSeries(
        name: "Product A",
        categoriesRange: r"ChartsDemo!$A$2:$A$7",
        valuesRange: r"ChartsDemo!$B$2:$B$7",
      ),
      ChartSeries(
        name: "Product B",
        categoriesRange: r"ChartsDemo!$A$2:$A$7",
        valuesRange: r"ChartsDemo!$C$2:$C$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 18, row: 0, width: 12, height: 16),
    showLegend: true,
  );

  sheet.addChart(chart);
}

/// Create a Pie Chart
void _addPieChart(Sheet sheet) {
  var chart = PieChart(
    title: "Average Sales Distribution",
    series: [
      ChartSeries(
        name: "Products",
        categoriesRange: r"ChartsDemo!$B$1:$D$1", // Product names
        valuesRange: r"ChartsDemo!$B$7:$D$7",     // Last month values
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 18, width: 10, height: 15),
    showLegend: true,
  );

  sheet.addChart(chart);
}

/// Create an Area Chart
void _addAreaChart(Sheet sheet) {
  var chart = AreaChart(
    title: "Cumulative Sales",
    series: [
      ChartSeries(
        name: "Product A",
        categoriesRange: r"ChartsDemo!$A$2:$A$7",
        valuesRange: r"ChartsDemo!$B$2:$B$7",
      ),
      ChartSeries(
        name: "Product B",
        categoriesRange: r"ChartsDemo!$A$2:$A$7",
        valuesRange: r"ChartsDemo!$C$2:$C$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 18, row: 18, width: 12, height: 15),
    showLegend: true,
  );

  sheet.addChart(chart);
}
