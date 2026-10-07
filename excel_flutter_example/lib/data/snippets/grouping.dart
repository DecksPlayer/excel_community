const String groupingSnippet = r'''
import 'dart:io';
import 'package:excel_community/excel_community.dart';

/// Generates the complete Row & Column Grouping demo workbook:
/// - Yearly sales report with months grouped per quarter (rows)
/// - Quarters grouped into half-years (nested 2 levels)
/// - Regional columns grouped with outline settings
void generateGroupingWorkbook() {
  final excel = Excel.createExcel();
  final sheet = excel['Sales 2026'];

  const regions = ['North', 'South', 'East', 'West'];
  const quarters = {
    'Q1': ['Jan', 'Feb', 'Mar'],
    'Q2': ['Apr', 'May', 'Jun'],
    'Q3': ['Jul', 'Aug', 'Sep'],
    'Q4': ['Oct', 'Nov', 'Dec'],
  };

  final header = CellStyle(
    bold: true,
    backgroundColorHex: ExcelColor.fromHexString('#FFEDD5'),
  );
  final total = CellStyle(bold: true);

  sheet.appendRow([
    TextCellValue('Period'),
    for (final r in regions) TextCellValue(r),
    TextCellValue('Total'),
  ]);
  for (var c = 0; c <= regions.length + 1; c++) {
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: 0)).cellStyle = header;
  }

  var row = 1;
  final halfStart = <String, int>{};
  final quarterTotals = <int>[];

  // 1. Build row groups for quarters and months
  for (final entry in quarters.entries) {
    if (entry.key == 'Q1' || entry.key == 'Q3') halfStart[entry.key] = row;
    final first = row;

    for (final month in entry.value) {
      sheet.appendRow([
        TextCellValue(month),
        for (var i = 0; i < regions.length; i++) IntCellValue(100 + (row * 37 + i * 53) % 150),
        FormulaCellValue('SUM(B${row + 1}:E${row + 1})'),
      ]);
      row++;
    }

    // Quarter summary row below its months
    sheet.appendRow([
      TextCellValue('${entry.key} total'),
      for (final col in ['B', 'C', 'D', 'E', 'F'])
        FormulaCellValue('SUM($col${first + 1}:$col$row)'),
    ]);
    for (var c = 0; c <= regions.length + 1; c++) {
      sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: row)).cellStyle = total;
    }

    // Group the months under this quarter (level 1)
    sheet.groupRows(first, row - 1);
    quarterTotals.add(row);
    row++;

    // Half-year summary row below its two quarters
    if (entry.key == 'Q2' || entry.key == 'Q4') {
      final start = halfStart[entry.key == 'Q2' ? 'Q1' : 'Q3']!;
      final totals = quarterTotals.sublist(quarterTotals.length - 2);
      sheet.appendRow([
        TextCellValue(entry.key == 'Q2' ? 'H1 total' : 'H2 total'),
        for (final col in ['B', 'C', 'D', 'E', 'F'])
          FormulaCellValue('$col${totals[0] + 1}+$col${totals[1] + 1}'),
      ]);
      for (var c = 0; c <= regions.length + 1; c++) {
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: row)).cellStyle = total;
      }
      // Group the two quarters under this half-year (level 2)
      sheet.groupRows(start, row - 1);
      row++;
    }
  }

  // 2. Build column groups: group inner detail regions (B-D)
  sheet.groupColumns(1, 3);

  // Set column widths
  sheet.setColumnWidth(0, 16);
  for (var c = 1; c <= regions.length + 1; c++) {
    sheet.setColumnWidth(c, 12);
  }

  // 3. Save the workbook
  final bytes = excel.save(fileName: 'grouping_example.xlsx');
  // Or: File('grouping_example.xlsx').writeAsBytesSync(excel.encode()!);
}
''';
