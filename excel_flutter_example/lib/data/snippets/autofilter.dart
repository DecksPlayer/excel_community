const String autoFilterSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateAutoFilterExcel() {
  final excel = Excel.createExcel();
  final sheet = excel['Inventory'];

  // 1. Setup Table Headers
  final headers = ['ID', 'Product', 'Category', 'Region', 'Stock', 'Price'];
  for (int col = 0; col < headers.length; col++) {
    final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0));
    cell.value = TextCellValue(headers[col]);
    cell.cellStyle = CellStyle(
      bold: true,
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
      backgroundColorHex: ExcelColor.fromHexString('#1E3A8A'), // Dark Blue
      horizontalAlign: HorizontalAlign.Center,
    );
  }

  // 2. Populate Sample Data Rows
  final records = [
    [101, 'MacBook Pro', 'Electronics', 'North', 45, 1999.0],
    [102, 'Ergonomic Chair', 'Furniture', 'South', 120, 350.0],
    [103, '4K Ultra Monitor', 'Electronics', 'East', 75, 499.0],
    [104, 'Standing Desk', 'Furniture', 'West', 30, 650.0],
    [105, 'Wireless Mouse', 'Electronics', 'North', 210, 49.0],
    [106, 'Bookshelf Unit', 'Furniture', 'South', 15, 180.0],
    [107, 'Mechanical Keyboard', 'Electronics', 'East', 95, 129.0],
    [108, 'Filing Cabinet', 'Furniture', 'West', 40, 220.0],
  ];

  for (int r = 0; r < records.length; r++) {
    final row = records[r];
    final rowIndex = r + 1;
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: rowIndex)).value = IntCellValue(row[0] as int);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: rowIndex)).value = TextCellValue(row[1] as String);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: rowIndex)).value = TextCellValue(row[2] as String);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: rowIndex)).value = TextCellValue(row[3] as String);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: rowIndex)).value = IntCellValue(row[4] as int);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: rowIndex)).value = DoubleCellValue(row[5] as double);
  }

  // 3. Enable AutoFilter covering headers and data rows (A1 to F9)
  sheet.setAutoFilter(
    CellIndex.indexByString('A1'),
    CellIndex.indexByString('F9'),
  );

  // 4. (Optional) Configure active column filter criteria
  // colId is 0-based relative to the start of the AutoFilter range:
  sheet.addFilterColumn(
    const FilterColumn(
      colId: 2, // Category column
      filterValues: ['Electronics'],
    ),
  );

  // 5. Inspect or clear AutoFilter programmatically
  if (sheet.hasAutoFilter) {
    print('Filter range: \${sheet.autoFilter?.ref}'); // "A1:F9"
    print('Columns: \${sheet.autoFilter?.columnCount}'); // 6
    print('Rows: \${sheet.autoFilter?.rowCount}'); // 9
  }

  // sheet.clearAutoFilter(); // Call this to remove the filter

  excel.save(fileName: 'autofilter_inventory.xlsx');
}
''';
