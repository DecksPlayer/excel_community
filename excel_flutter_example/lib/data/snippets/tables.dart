const String tablesSnippet = r'''
import 'package:excel_community/excel_community.dart';

void addSalesTable(Excel excel) {
  final sheet = excel['Sales'];

  final table = sheet.addTable('A1:C10',
      name: 'Sales',
      style: TableStyle.medium(9),
      showTotalsRow: true,
      columns: const [
        TableColumn('Region', totalsLabel: 'Total'),
        TableColumn('Units', totalsFunction: TableTotalsFunction.sum),
        TableColumn('Price', totalsFunction: TableTotalsFunction.average),
      ]);

  sheet.updateCell(CellIndex.indexByString('E1'),
      FormulaCellValue('SUM(${table.columnReference('Units')})'));

  final rows = sheet.tableRowsAsMaps('Sales');
}
''';
