import 'dart:io';
import 'package:excel_community/excel_community.dart';

void main() {
  final excel = Excel.createExcel();
  final sheet = excel['Sales'];

  // Add Headers
  final headers = ['Order ID', 'Product', 'Region', 'Category', 'Quantity', 'Total'];
  for (int i = 0; i < headers.length; i++) {
    final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0));
    cell.value = TextCellValue(headers[i]);
    cell.cellStyle = CellStyle(
      bold: true,
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
      backgroundColorHex: ExcelColor.fromHexString('#1F4E79'),
      horizontalAlign: HorizontalAlign.Center,
    );
  }

  // Add Sample Data
  final data = [
    [1001, 'Laptop', 'North', 'Electronics', 5, 4500.0],
    [1002, 'Desk Chair', 'South', 'Furniture', 12, 1800.0],
    [1003, 'Monitor', 'East', 'Electronics', 8, 2400.0],
    [1004, 'Bookshelf', 'West', 'Furniture', 3, 750.0],
    [1005, 'Mouse', 'North', 'Electronics', 25, 625.0],
    [1006, 'Coffee Table', 'South', 'Furniture', 4, 1200.0],
    [1007, 'Keyboard', 'East', 'Electronics', 15, 1125.0],
    [1008, 'Filing Cabinet', 'West', 'Furniture', 2, 500.0],
  ];

  for (int r = 0; r < data.length; r++) {
    final row = data[r];
    final rowIndex = r + 1;

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: rowIndex)).value =
        IntCellValue(row[0] as int);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: rowIndex)).value =
        TextCellValue(row[1] as String);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: rowIndex)).value =
        TextCellValue(row[2] as String);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: rowIndex)).value =
        TextCellValue(row[3] as String);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: rowIndex)).value =
        IntCellValue(row[4] as int);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: rowIndex)).value =
        DoubleCellValue(row[5] as double);
  }

  // Enable AutoFilter covering headers and all data rows (A1 to F9)
  sheet.setAutoFilter(
    CellIndex.indexByString('A1'),
    CellIndex.indexByString('F9'),
  );

  print('AutoFilter enabled on range: ${sheet.autoFilter?.ref}');
  print('Columns covered: ${sheet.autoFilter?.columnCount}');
  print('Rows covered: ${sheet.autoFilter?.rowCount}');

  // Save the spreadsheet
  final bytes = excel.save()!;
  final file = File('example/autofilter_output.xlsx');
  file.writeAsBytesSync(bytes);

  print('Workbook saved successfully to ${file.path}');
}
