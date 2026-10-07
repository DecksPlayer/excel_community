const String tablesSnippet = r'''
import 'dart:io';
import 'package:excel_community/excel_community.dart';

/// Generates the complete Excel Tables demo workbook:
/// - "Sales": formatted table with totals row and structured references
/// - "Style Gallery": sample table demonstrating built-in TableStyles
void generateTablesWorkbook() {
  final excel = Excel.createExcel();

  // -------------------------------------------------------------------------
  // 1. "Sales" sheet: Table with structured references and totals row
  // -------------------------------------------------------------------------
  final sales = excel['Sales'];
  const rows = [
    ('North', 120, 2.5),
    ('South', 95, 3.0),
    ('East', 140, 1.75),
    ('West', 80, 2.25),
    ('Central', 110, 2.0),
  ];

  sales.appendRow([
    TextCellValue('Region'),
    TextCellValue('Units'),
    TextCellValue('Price'),
    TextCellValue('Revenue'),
  ]);

  for (var i = 0; i < rows.length; i++) {
    final (region, units, price) = rows[i];
    sales.appendRow([
      TextCellValue(region),
      IntCellValue(units),
      DoubleCellValue(price),
      FormulaCellValue('B${i + 2}*C${i + 2}'),
    ]);
  }

  // Add the formatted table covering the headers, rows, and totals row
  final table = sales.addTable(
    'A1:D${rows.length + 2}',
    name: 'Sales',
    style: TableStyle.medium(9),
    showTotalsRow: true,
    columns: const [
      TableColumn('Region', totalsLabel: 'Total'),
      TableColumn('Units', totalsFunction: TableTotalsFunction.sum),
      TableColumn('Price', totalsFunction: TableTotalsFunction.average),
      TableColumn('Revenue', totalsFunction: TableTotalsFunction.sum),
    ],
  );

  // Use structured reference formula: [Revenue] column of table
  sales.updateCell(
    CellIndex.indexByString('F1'),
    TextCellValue('Best region revenue:'),
  );
  sales.updateCell(
    CellIndex.indexByString('G1'),
    FormulaCellValue('MAX(${table.columnReference('Revenue')})'),
  );

  for (var c = 0; c < 7; c++) {
    sales.setColumnWidth(c, c == 5 ? 22 : 13);
  }

  // -------------------------------------------------------------------------
  // 2. "Style Gallery" sheet: Demonstrating built-in TableStyles
  // -------------------------------------------------------------------------
  final gallery = excel['Style Gallery'];
  final sampleStyles = [
    ('Light 1', TableStyle.light(1)),
    ('Medium 2', TableStyle.medium(2)),
    ('Medium 9', TableStyle.medium(9)),
    ('Dark 1', TableStyle.dark(1)),
  ];

  for (var i = 0; i < sampleStyles.length; i++) {
    final (label, style) = sampleStyles[i];
    final left = i * 4;
    CellIndex at(int r, int c) =>
        CellIndex.indexByColumnRow(columnIndex: left + c, rowIndex: r);

    gallery.updateCell(at(0, 0), TextCellValue(label),
        cellStyle: CellStyle(bold: true));
    gallery.updateCell(at(1, 0), TextCellValue('Item'));
    gallery.updateCell(at(1, 1), TextCellValue('Qty'));
    gallery.updateCell(at(1, 2), TextCellValue('Price'));

    for (var r = 0; r < 3; r++) {
      gallery.updateCell(at(r + 2, 0), TextCellValue('Product ${r + 1}'));
      gallery.updateCell(at(r + 2, 1), IntCellValue(10 + r * 5));
      gallery.updateCell(at(r + 2, 2), DoubleCellValue(1.5 + r));
    }

    final topCol = getColumnAlphabet(left);
    final botCol = getColumnAlphabet(left + 2);
    gallery.addTable(
      '${topCol}2:${botCol}5',
      name: 'Table_${style.name}',
      style: style,
    );
  }

  // Save output
  final bytes = excel.save(fileName: 'tables_example.xlsx');
  // Or: File('tables_example.xlsx').writeAsBytesSync(excel.encode()!);
}
''';
