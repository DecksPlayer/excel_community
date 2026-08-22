import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateChartDataLabelsHelper() async {
  final excel = Excel.createExcel();

  // ── Sheet 1: Column chart — value labels ────────────────────────────────
  final salesSheet = excel['Sales'];
  salesSheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Month'));
  salesSheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Revenue'));
  salesSheet.updateCell(CellIndex.indexByString('C1'), TextCellValue('Expenses'));

  final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  final revenue = [42000, 55000, 38000, 71000, 64000, 80000];
  final expenses = [30000, 32000, 28000, 45000, 40000, 52000];
  for (var i = 0; i < months.length; i++) {
    salesSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
        TextCellValue(months[i]));
    salesSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
        IntCellValue(revenue[i]));
    salesSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1),
        IntCellValue(expenses[i]));
  }

  salesSheet.addChart(ColumnChart(
    title: 'Monthly Revenue vs Expenses',
    dataLabels: const ChartDataLabels(
      value: true,
      labelPosition: 'outEnd',
    ),
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"Sales!$A$2:$A$7",
        valuesRange: r"Sales!$B$2:$B$7",
      ),
      ChartSeries(
        name: 'Expenses',
        categoriesRange: r"Sales!$A$2:$A$7",
        valuesRange: r"Sales!$C$2:$C$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
  ));

  // ── Sheet 2: Pie chart — percentage + value labels ───────────────────────
  final marketSheet = excel['Market Share'];
  final companies = ['Alpha', 'Beta', 'Gamma', 'Delta', 'Others'];
  final shares = [34, 27, 18, 12, 9];
  marketSheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Company'));
  marketSheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Share %'));
  for (var i = 0; i < companies.length; i++) {
    marketSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
        TextCellValue(companies[i]));
    marketSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
        IntCellValue(shares[i]));
  }

  marketSheet.addChart(PieChart(
    title: 'Global Market Share',
    dataLabels: const ChartDataLabels(
      value: true,
      percentage: true,
      separator: '\n',
    ),
    series: [
      ChartSeries(
        name: 'Share',
        categoriesRange: r"'Market Share'!$A$2:$A$6",
        valuesRange: r"'Market Share'!$B$2:$B$6",
      ),
    ],
    anchor: ChartAnchor.at(column: 4, row: 1, width: 11, height: 16),
  ));

  // ── Sheet 3: Line chart — category name + value labels ───────────────────
  final trendSheet = excel['Trend'];
  final quarters = ['Q1 2023', 'Q2 2023', 'Q3 2023', 'Q4 2023', 'Q1 2024', 'Q2 2024'];
  final units = [120, 145, 132, 178, 165, 210];
  trendSheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Quarter'));
  trendSheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Units'));
  for (var i = 0; i < quarters.length; i++) {
    trendSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
        TextCellValue(quarters[i]));
    trendSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
        IntCellValue(units[i]));
  }

  trendSheet.addChart(LineChart(
    title: 'Units Sold by Quarter',
    dataLabels: const ChartDataLabels(
      value: true,
      categoryName: true,
      separator: ' — ',
    ),
    series: [
      ChartSeries(
        name: 'Units',
        categoriesRange: r"Trend!$A$2:$A$7",
        valuesRange: r"Trend!$B$2:$B$7",
      ),
    ],
    anchor: ChartAnchor.at(column: 4, row: 1, width: 12, height: 15),
  ));

  excel.delete('Sheet1');

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'charts_with_data_labels.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Charts with Data Labels generated!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          '\n3 sheets:\n'
          '• Sales — Column chart (value labels)\n'
          '• Market Share — Pie chart (% + value)\n'
          '• Trend — Line chart (category + value)';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) throw Exception('Failed to encode Excel file.');

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Charts with Data Labels',
      fileName: 'charts_with_data_labels.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      await File(outputFile).writeAsBytes(bytes);
      return '✅ Saved successfully!\n'
          'Location: $outputFile\n'
          '\n3 sheets:\n'
          '• Sales — Column chart (value labels)\n'
          '• Market Share — Pie chart (% + value)\n'
          '• Trend — Line chart (category + value)';
    }
    return 'Save cancelled.';
  }
}
