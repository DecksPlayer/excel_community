const String hiddenColumnsSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateHiddenColumnsMultiSheet() {
  final excel = Excel.createExcel();

  // ── Sheet 1: Public Directory (All Visible) ──────────────────────
  final publicSheet = excel['Public Directory'];
  publicSheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Name"));
  publicSheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Role"));
  // (No hidden columns or rows in this sheet)

  // ── Sheet 2: Staff Salaries (Confidential - Hidden Col & Row) ────
  final salarySheet = excel['Staff Salaries (Confidential)'];
  salarySheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Name"));
  salarySheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Salary (Confidential)"));
  salarySheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Role"));
  
  // Hide Column B (index 1) - Confidential Salary
  salarySheet.setColumnHidden(1, true);

  // Hide Row 3 (index 2) - Inactive Employee
  salarySheet.setRowHidden(2, true);

  // ── Sheet 3: Executive Perks (Confidential - Hidden Col) ─────────
  final perksSheet = excel['Executive Perks (Confidential)'];
  perksSheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Executive"));
  perksSheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Allowance (Confidential)"));
  
  // Hide Column B (index 1) - Confidential Allowance
  perksSheet.setColumnHidden(1, true);

  excel.save(fileName: 'hidden_elements_multi_sheet.xlsx');
}
''';
