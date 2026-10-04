const String tabColorSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateTabColorsExcel() {
  final excel = Excel.createExcel();

  // 1. Executive Summary Sheet - Vibrant Royal Blue
  final execSheet = excel['Executive Summary'];
  execSheet.setTabColorHex('#2563EB'); // Hex ARGB or #RRGGBB
  execSheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Quarterly Executive Review');

  // 2. Marketing Campaigns Sheet - Emerald Green
  final mktSheet = excel['Marketing'];
  mktSheet.setTabColorHex('#10B981');
  mktSheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Q3 Marketing Metrics');

  // 3. Finance & P&L Sheet - Amber / Gold
  final finSheet = excel['Finance'];
  finSheet.setTabColorHex('#F59E0B');
  finSheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('P&L and Cash Flow');

  // 4. Operations Sheet - Pink
  final opsSheet = excel['Operations'];
  opsSheet.setTabColorHex('#EC4899');
  opsSheet.cell(CellIndex.indexByString('A1')).value = TextCellValue('Logistics & Inventory');

  // 5. Using standard ExcelColor enum
  final devSheet = excel['Engineering'];
  devSheet.setTabColor(ExcelColor.purple);

  // 6. Using Office theme color with tint
  final archiveSheet = excel['Archive'];
  archiveSheet.tabColor = TabColor.fromTheme(1, tint: -0.25);

  // 7. Inspect or remove tab colors dynamically
  if (execSheet.hasTabColor) {
    print('Executive Tab Color: \${execSheet.tabColor?.colorHex6}'); // "2563EB"
  }

  // To clear/remove a tab color:
  // archiveSheet.clearTabColor();

  excel.save(fileName: 'colored_sheets_portfolio.xlsx');
}
''';
