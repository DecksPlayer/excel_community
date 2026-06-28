import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';
import '../../models/section_detail.dart';

Future<String> generateExcelWithChartHelper(ChartType type) async {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // Populate helper data inline to keep helper self-contained
  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Value 1"));
  sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Value 2"));

  final data = [
    ['Jan', 10, 15],
    ['Feb', 20, 25],
    ['Mar', 15, 30],
    ['Apr', 25, 20],
    ['May', 30, 35],
    ['Jun', 20, 40],
  ];

  for (var i = 0; i < data.length; i++) {
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(data[i][0] as String));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), IntCellValue(data[i][1] as int));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), IntCellValue(data[i][2] as int));
  }

  Chart chart;
  final series = [
    ChartSeries(
      name: "Series 1",
      categoriesRange: r"Sheet1!$A$2:$A$7",
      valuesRange: r"Sheet1!$B$2:$B$7",
    ),
    ChartSeries(
      name: "Series 2",
      categoriesRange: r"Sheet1!$A$2:$A$7",
      valuesRange: r"Sheet1!$C$2:$C$7",
    ),
  ];

  final anchor = ChartAnchor.at(column: 5, row: 1, width: 10, height: 15);

  switch (type) {
    case ChartType.column:
      chart = ColumnChart(title: "Monthly Data (Column)", series: series, anchor: anchor);
    case ChartType.bar:
      chart = BarChart(title: "Monthly Data (Bar)", series: series, anchor: anchor);
    case ChartType.line:
      chart = LineChart(title: "Monthly Data (Line)", series: series, anchor: anchor);
    case ChartType.area:
      chart = AreaChart(title: "Monthly Data (Area)", series: series, anchor: anchor);
    case ChartType.pie:
      chart = PieChart(title: "Pie Chart Example", series: [series[0]], anchor: anchor);
    case ChartType.doughnut:
      chart = DoughnutChart(title: "Doughnut Chart Example", series: [series[0]], anchor: anchor);
    case ChartType.radar:
      chart = RadarChart(title: "Radar Chart Example", series: series, anchor: anchor, filled: true);
    case ChartType.scatter:
      chart = ScatterChart(title: "Scatter Chart Example", series: series, anchor: anchor);
  }

  sheet.addChart(chart);

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'chart_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Excel generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '\n📥 Check your Downloads folder\n'
          '📌 File: chart_example.xlsx';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Excel File',
      fileName: 'chart_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Excel saved successfully!\n'
          'Location: $outputFile\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
          '\n📌 IMPORTANT: Open with Excel, not a text editor!';
    }
    return 'Save cancelled.';
  }
}

Future<String> generateAllChartsHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['All Charts Demo'];
  excel.delete('Sheet1');

  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Series A"));
  sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Series B"));

  final data = [
    ['Q1', 10, 15],
    ['Q2', 25, 20],
    ['Q3', 15, 30],
    ['Q4', 30, 25],
  ];

  for (var i = 0; i < data.length; i++) {
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(data[i][0] as String));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), IntCellValue(data[i][1] as int));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), IntCellValue(data[i][2] as int));
  }

  final seriesMulti = [
    ChartSeries(
      name: "Series A",
      categoriesRange: "'All Charts Demo'!\$A\$2:\$A\$5",
      valuesRange: "'All Charts Demo'!\$B\$2:\$B\$5",
    ),
    ChartSeries(
      name: "Series B",
      categoriesRange: "'All Charts Demo'!\$A\$2:\$A\$5",
      valuesRange: "'All Charts Demo'!\$C\$2:\$C\$5",
    ),
  ];
  final seriesSingle = [seriesMulti[0]];

  sheet.addChart(ColumnChart(title: "Column Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 5, row: 1, width: 8, height: 12)));
  sheet.addChart(BarChart(title: "Bar Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 14, row: 1, width: 8, height: 12)));
  sheet.addChart(LineChart(title: "Line Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 5, row: 14, width: 8, height: 12)));
  sheet.addChart(AreaChart(title: "Area Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 14, row: 14, width: 8, height: 12)));
  sheet.addChart(PieChart(title: "Pie Chart", series: seriesSingle, anchor: ChartAnchor.at(column: 5, row: 27, width: 8, height: 12)));
  sheet.addChart(DoughnutChart(title: "Doughnut Chart", series: seriesSingle, anchor: ChartAnchor.at(column: 14, row: 27, width: 8, height: 12)));
  sheet.addChart(RadarChart(title: "Radar Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 5, row: 40, width: 8, height: 12), filled: true));

  sheet.updateCell(CellIndex.indexByString("E1"), TextCellValue("X Values"));
  sheet.updateCell(CellIndex.indexByString("F1"), TextCellValue("Y Values A"));
  sheet.updateCell(CellIndex.indexByString("G1"), TextCellValue("Y Values B"));

  final scatterData = [
    [1.0, 5.0, 10.0],
    [2.5, 12.0, 8.0],
    [4.0, 18.0, 15.0],
    [5.5, 25.0, 22.0],
    [7.0, 30.0, 28.0],
    [8.5, 45.0, 35.0],
    [10.0, 55.0, 48.0],
  ];

  for (var i = 0; i < scatterData.length; i++) {
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: i + 1), DoubleCellValue(scatterData[i][0]));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: i + 1), DoubleCellValue(scatterData[i][1]));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: i + 1), DoubleCellValue(scatterData[i][2]));
  }

  final scatterSeries = [
    ChartSeries(name: "Trend A", categoriesRange: "'All Charts Demo'!\$E\$2:\$E\$8", valuesRange: "'All Charts Demo'!\$F\$2:\$F\$8"),
    ChartSeries(name: "Trend B", categoriesRange: "'All Charts Demo'!\$E\$2:\$E\$8", valuesRange: "'All Charts Demo'!\$G\$2:\$G\$8"),
  ];

  sheet.addChart(ScatterChart(title: "Scatter Chart (XY Relationship)", series: scatterSeries, anchor: ChartAnchor.at(column: 14, row: 40, width: 8, height: 12)));

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'all_charts_demo.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Excel with ALL charts generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 Includes: Column, Bar, Line, Area, Pie, Doughnut, Radar, Scatter.';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes != null) {
      String? outputFile = await FilePicker.platform.saveFile(
        dialogTitle: 'Save All Charts Excel',
        fileName: 'all_charts_demo.xlsx',
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );
      if (outputFile != null) {
        await File(outputFile).writeAsBytes(bytes);
        return '✅ Excel with ALL charts saved successfully!\n'
            'Location: $outputFile';
      }
    }
    return 'Save cancelled.';
  }
}
