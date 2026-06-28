import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateFullDemoHelper() async {
  var excel = Excel.createExcel();
  var sheetName = 'Full Demo';
  excel.rename('Sheet1', sheetName);
  var sheet = excel[sheetName];

  final headerStyle = CellStyle(
    bold: true,
    fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
    backgroundColorHex: ExcelColor.fromHexString('#4472C4'),
    horizontalAlign: HorizontalAlign.Center,
    verticalAlign: VerticalAlign.Center,
    leftBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
    rightBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
    topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
    bottomBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
  );

  final dataStyle = CellStyle(
    leftBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
    rightBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
    topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
    bottomBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
  );

  final formulaStyle = CellStyle(
    bold: true,
    backgroundColorHex: ExcelColor.fromHexString('#E2EFDA'),
    leftBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
    rightBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
    topBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
    bottomBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
  );

  final multiColorStyle = CellStyle(
    bold: true,
    horizontalAlign: HorizontalAlign.Center,
    verticalAlign: VerticalAlign.Center,
    leftBorder: Border(borderStyle: BorderStyle.Thick, borderColorHex: ExcelColor.fromHexString('#00FF00')),
    rightBorder: Border(borderStyle: BorderStyle.Thick, borderColorHex: ExcelColor.fromHexString('#FFFF00')),
    topBorder: Border(borderStyle: BorderStyle.Thick, borderColorHex: ExcelColor.fromHexString('#FF0000')),
    bottomBorder: Border(borderStyle: BorderStyle.Thick, borderColorHex: ExcelColor.fromHexString('#0000FF')),
  );

  sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Month"), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Revenue"), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Expenses"), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString("D1"), TextCellValue("Profit"), cellStyle: headerStyle);

  final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  final revenues = [1000, 1200, 1500, 1300, 1700, 2000];
  final expenses = [800, 900, 1000, 950, 1100, 1200];

  for (var i = 0; i < months.length; i++) {
    var row = i + 2;
    sheet.updateCell(CellIndex.indexByString("A$row"), TextCellValue(months[i]), cellStyle: dataStyle);
    sheet.updateCell(CellIndex.indexByString("B$row"), IntCellValue(revenues[i]), cellStyle: dataStyle);
    sheet.updateCell(CellIndex.indexByString("C$row"), IntCellValue(expenses[i]), cellStyle: dataStyle);
    sheet.updateCell(CellIndex.indexByString("D$row"), FormulaCellValue("B$row-C$row"), cellStyle: dataStyle);
  }

  var totalRow = months.length + 2;
  sheet.updateCell(CellIndex.indexByString("A$totalRow"), TextCellValue("TOTAL"), cellStyle: formulaStyle);
  sheet.updateCell(CellIndex.indexByString("B$totalRow"), FormulaCellValue("SUM(B2:B${totalRow - 1})"), cellStyle: formulaStyle);
  sheet.updateCell(CellIndex.indexByString("C$totalRow"), FormulaCellValue("SUM(C2:C${totalRow - 1})"), cellStyle: formulaStyle);
  sheet.updateCell(CellIndex.indexByString("D$totalRow"), FormulaCellValue("SUM(D2:D${totalRow - 1})"), cellStyle: formulaStyle);

  sheet.updateCell(CellIndex.indexByString("A${totalRow + 2}"), TextCellValue("Multi-colored Borders"), cellStyle: multiColorStyle);
  sheet.setColumnWidth(0, 25.0);

  final series = [
    ChartSeries(
      name: "Revenue",
      categoriesRange: "'Full Demo'!\$A\$2:\$A\$7",
      valuesRange: "'Full Demo'!\$B\$2:\$B\$7",
    ),
    ChartSeries(
      name: "Profit",
      categoriesRange: "'Full Demo'!\$A\$2:\$A\$7",
      valuesRange: "'Full Demo'!\$D\$2:\$D\$7",
    ),
  ];

  final chart = ColumnChart(
    title: "Financial Overview",
    series: series,
    anchor: ChartAnchor.at(column: 6, row: 1, width: 10, height: 15),
  );

  sheet.addChart(chart);

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'full_demo_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Full Demo Excel generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 Includes: Formulas, Headers, Borders, and Charts.';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Full Demo Excel',
      fileName: 'full_demo_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      return '✅ Full Demo Excel saved successfully!\n'
          'Location: $outputFile\n'
          '📌 Includes: Formulas, Headers, Borders, and Charts.';
    }
    return 'Save cancelled.';
  }
}
