import 'package:excel_community/excel_community.dart';

/// Every built-in table style, in Excel's gallery order.
final List<(String, TableStyle)> builtInTableStyles = [
  for (var i = 1; i <= 21; i++) ('Light $i', TableStyle.light(i)),
  for (var i = 1; i <= 28; i++) ('Medium $i', TableStyle.medium(i)),
  for (var i = 1; i <= 11; i++) ('Dark $i', TableStyle.dark(i)),
];

/// Sheet with the sample data used by the table examples (header + 3 rows).
Sheet tableSampleSheet() {
  final sheet = Excel.createExcel()['Sheet1'];
  sheet.appendRow([TextCellValue('Region'), TextCellValue('Units'), TextCellValue('Price')]);
  sheet.appendRow([TextCellValue('North'), IntCellValue(10), DoubleCellValue(2.5)]);
  sheet.appendRow([TextCellValue('South'), IntCellValue(7), DoubleCellValue(3)]);
  sheet.appendRow([TextCellValue('East'), IntCellValue(4), DoubleCellValue(1.5)]);
  return sheet;
}

/// One example of the Options / Data tabs.
class TableSample {
  final String category; // options, data
  final String title;
  final String subtitle;
  final String code;

  /// Builds the sheet and returns the table to draw.
  final ExcelTable Function(Sheet sheet) apply;

  /// Optional text output computed from the sheet.
  final String Function(Sheet sheet)? note;

  const TableSample({
    required this.category,
    required this.title,
    required this.subtitle,
    required this.code,
    required this.apply,
    this.note,
  });
}

final List<TableSample> tableSamples = [
  // --- Options ----------------------------------------------------------------
  TableSample(
    category: 'options',
    title: 'Totals Row',
    subtitle: 'showTotalsRow + TableColumn',
    code: "sheet.addTable('A1:C5', name: 'Sales',\n"
        '    showTotalsRow: true,\n'
        '    columns: const [\n'
        "      TableColumn('Region', totalsLabel: 'Total'),\n"
        "      TableColumn('Units',\n"
        '          totalsFunction: TableTotalsFunction.sum),\n'
        "      TableColumn('Price',\n"
        '          totalsFunction: TableTotalsFunction.average),\n'
        '    ]);',
    apply: (sheet) => sheet.addTable('A1:C5', name: 'Sales', showTotalsRow: true, columns: const [
      TableColumn('Region', totalsLabel: 'Total'),
      TableColumn('Units', totalsFunction: TableTotalsFunction.sum),
      TableColumn('Price', totalsFunction: TableTotalsFunction.average),
    ]),
  ),
  TableSample(
    category: 'options',
    title: 'Banded Columns',
    subtitle: 'showColumnStripes: true',
    code: "sheet.addTable('A1:C4', name: 'Sales',\n"
        '    showRowStripes: false,\n'
        '    showColumnStripes: true);',
    apply: (sheet) =>
        sheet.addTable('A1:C4', name: 'Sales', showRowStripes: false, showColumnStripes: true),
  ),
  TableSample(
    category: 'options',
    title: 'First & Last Column',
    subtitle: 'showFirstColumn / showLastColumn',
    code: "sheet.addTable('A1:C4', name: 'Sales',\n"
        '    style: TableStyle.medium(9),\n'
        '    showFirstColumn: true,\n'
        '    showLastColumn: true);',
    apply: (sheet) => sheet.addTable('A1:C4',
        name: 'Sales', style: TableStyle.medium(9), showFirstColumn: true, showLastColumn: true),
  ),
  TableSample(
    category: 'options',
    title: 'No Filter Buttons',
    subtitle: 'showFilterButtons: false',
    code: "sheet.addTable('A1:C4', name: 'Sales',\n"
        '    showFilterButtons: false);',
    apply: (sheet) => sheet.addTable('A1:C4', name: 'Sales', showFilterButtons: false),
  ),
  TableSample(
    category: 'options',
    title: 'No Header Row',
    subtitle: 'showHeaderRow: false',
    code: "// Data starts on the first row; columns are\n"
        '// named Column1, Column2, ...\n'
        "sheet.addTable('A2:C4', name: 'Sales',\n"
        '    showHeaderRow: false);',
    apply: (sheet) => sheet.addTable('A2:C4', name: 'Sales', showHeaderRow: false),
    note: (sheet) => 'columns → ${sheet.getTable('Sales')!.columns.map((c) => c.name).toList()}',
  ),
  TableSample(
    category: 'options',
    title: 'Without Style',
    subtitle: 'style: null',
    code: "sheet.addTable('A1:C4', name: 'Sales',\n"
        '    style: null);\n'
        '// Still a table: filters and structured refs',
    apply: (sheet) => sheet.addTable('A1:C4', name: 'Sales', style: null),
  ),

  // --- Data -------------------------------------------------------------------
  TableSample(
    category: 'data',
    title: 'Columns from the Header',
    subtitle: 'addTable(range, name:)',
    code: '// Column names are read from row 1\n'
        "final table = sheet.addTable('A1:C4', name: 'Sales');\n"
        'print(table.columns.map((c) => c.name));',
    apply: (sheet) => sheet.addTable('A1:C4', name: 'Sales'),
    note: (sheet) => 'columns → ${sheet.getTable('Sales')!.columns.map((c) => c.name).toList()}',
  ),
  TableSample(
    category: 'data',
    title: 'Read as Maps',
    subtitle: 'tableRowsAsMaps()',
    code: "sheet.addTable('A1:C4', name: 'Sales');\n"
        "final rows = sheet.tableRowsAsMaps('Sales');",
    apply: (sheet) => sheet.addTable('A1:C4', name: 'Sales'),
    note: (sheet) => sheet.tableRowsAsMaps('Sales').map((r) => '$r').join('\n'),
  ),
  TableSample(
    category: 'data',
    title: 'Append a Row',
    subtitle: 'appendTableRow()',
    code: "sheet.addTable('A1:C5', name: 'Sales',\n"
        '    showTotalsRow: true, columns: [...]);\n'
        "sheet.appendTableRow('Sales',\n"
        "    [TextCellValue('West'), IntCellValue(9), DoubleCellValue(2)]);\n"
        '// The totals row moves down',
    apply: (sheet) {
      sheet.addTable('A1:C5', name: 'Sales', showTotalsRow: true, columns: const [
        TableColumn('Region', totalsLabel: 'Total'),
        TableColumn('Units', totalsFunction: TableTotalsFunction.sum),
        TableColumn('Price'),
      ]);
      return sheet.appendTableRow('Sales', [TextCellValue('West'), IntCellValue(9), DoubleCellValue(2)]);
    },
    note: (sheet) => 'ref → ${sheet.getTable('Sales')!.ref}',
  ),
  TableSample(
    category: 'data',
    title: 'Structured References',
    subtitle: 'table.columnReference()',
    code: "final table = sheet.addTable('A1:C4', name: 'Sales');\n"
        "sheet.updateCell(CellIndex.indexByString('E1'),\n"
        "    FormulaCellValue('SUM(\${table.columnReference('Units')})'));",
    apply: (sheet) {
      final table = sheet.addTable('A1:C4', name: 'Sales');
      sheet.updateCell(CellIndex.indexByString('E1'),
          FormulaCellValue('SUM(${table.columnReference('Units')})'));
      return table;
    },
    note: (sheet) => 'E1 → =${(sheet.cell(CellIndex.indexByString('E1')).value as FormulaCellValue).formula}',
  ),
  TableSample(
    category: 'data',
    title: 'Tables Move with Rows',
    subtitle: 'insertRow() / insertColumn()',
    code: "sheet.addTable('A1:C4', name: 'Sales');\n"
        'sheet.insertRow(0);\n'
        '// New column inside the table → "Column4"\n'
        'sheet.insertColumn(1);',
    apply: (sheet) {
      sheet.addTable('A1:C4', name: 'Sales');
      sheet.insertRow(0);
      sheet.insertColumn(1);
      return sheet.getTable('Sales')!;
    },
    note: (sheet) {
      final t = sheet.getTable('Sales')!;
      return 'ref → ${t.ref}\ncolumns → ${t.columns.map((c) => c.name).toList()}';
    },
  ),
];
