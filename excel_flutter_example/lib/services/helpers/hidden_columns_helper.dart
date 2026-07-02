import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateHiddenColumnsHelper() async {
  final excel = Excel.createExcel();

  // ── Sheet 1: Public Directory (No Hidden Elements) ──────────────────────
  final publicSheet = excel['Public Directory'];
  if (excel.sheets.containsKey('Sheet1')) {
    excel.delete('Sheet1');
  }

  publicSheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Public Employee Directory"),
    cellStyle: CellStyle(bold: true, fontSize: 14, fontColorHex: ExcelColor.blue),
  );
  publicSheet.updateCell(
    CellIndex.indexByString("A2"),
    TextCellValue("All columns and rows are visible in this sheet."),
    cellStyle: CellStyle(italic: true, fontSize: 10),
  );

  final publicHeaders = ["ID", "Full Name", "Role", "Department"];
  for (var i = 0; i < publicHeaders.length; i++) {
    publicSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 3),
      TextCellValue(publicHeaders[i]),
      cellStyle: CellStyle(bold: true, backgroundColorHex: ExcelColor.grey300),
    );
  }

  final publicData = [
    ["EMP01", "Alice Johnson", "Lead Architect", "Engineering"],
    ["EMP02", "Bob Smith", "UI Designer", "Design"],
    ["EMP04", "Diana Prince", "Product Manager", "Product"],
    ["EMP05", "Evan Wright", "QA Specialist", "Engineering"],
  ];

  for (var r = 0; r < publicData.length; r++) {
    final rowIdx = 4 + r;
    for (var c = 0; c < publicData[r].length; c++) {
      publicSheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: rowIdx),
        TextCellValue(publicData[r][c]),
      );
    }
  }

  publicSheet.setColumnWidth(0, 12);
  publicSheet.setColumnWidth(1, 20);
  publicSheet.setColumnWidth(2, 20);
  publicSheet.setColumnWidth(3, 18);


  // ── Sheet 2: Staff Salaries (Has Hidden Column & Row) ───────────────────
  final salarySheet = excel['Staff Salaries (Confidential)'];
  salarySheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Internal Payroll List"),
    cellStyle: CellStyle(bold: true, fontSize: 14, fontColorHex: ExcelColor.red),
  );
  salarySheet.updateCell(
    CellIndex.indexByString("A2"),
    TextCellValue("Column C (Base Salary) and Row 7 (Charlie Brown - Inactive) are HIDDEN."),
    cellStyle: CellStyle(italic: true, fontSize: 10),
  );

  final salaryHeaders = ["ID", "Full Name", "Base Salary (HIDDEN)", "Role", "Department", "Status"];
  for (var i = 0; i < salaryHeaders.length; i++) {
    salarySheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 3),
      TextCellValue(salaryHeaders[i]),
      cellStyle: CellStyle(bold: true, backgroundColorHex: ExcelColor.grey300),
    );
  }

  final salaryData = [
    ["EMP01", "Alice Johnson", 85000, "Lead Architect", "Engineering", "Active"],
    ["EMP02", "Bob Smith", 62000, "UI Designer", "Design", "Active"],
    ["EMP03", "Charlie Brown", 45000, "Junior Developer", "Engineering", "Inactive (Hidden Row)"],
    ["EMP04", "Diana Prince", 95000, "Product Manager", "Product", "Active"],
    ["EMP05", "Evan Wright", 58000, "QA Specialist", "Engineering", "Active"],
  ];

  for (var r = 0; r < salaryData.length; r++) {
    final rowIdx = 4 + r;
    final row = salaryData[r];
    salarySheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: rowIdx), TextCellValue(row[0] as String));
    salarySheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: rowIdx), TextCellValue(row[1] as String));
    salarySheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: rowIdx), IntCellValue(row[2] as int));
    salarySheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: rowIdx), TextCellValue(row[3] as String));
    salarySheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: rowIdx), TextCellValue(row[4] as String));
    salarySheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: rowIdx), TextCellValue(row[5] as String));
  }

  salarySheet.setColumnHidden(2, true); // Hide Column C
  salarySheet.setRowHidden(6, true);    // Hide row index 6 ("Charlie Brown")

  salarySheet.setColumnWidth(0, 12);
  salarySheet.setColumnWidth(1, 20);
  salarySheet.setColumnWidth(3, 20);
  salarySheet.setColumnWidth(4, 18);
  salarySheet.setColumnWidth(5, 15);


  // ── Sheet 3: Executive Perks (Has Hidden Column) ───────────────────────
  final perksSheet = excel['Executive Perks (Confidential)'];
  perksSheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Executive Compensation & Allowances"),
    cellStyle: CellStyle(bold: true, fontSize: 14, fontColorHex: ExcelColor.deepOrange),
  );
  perksSheet.updateCell(
    CellIndex.indexByString("A2"),
    TextCellValue("Column C (Monthly Allowance) is HIDDEN."),
    cellStyle: CellStyle(italic: true, fontSize: 10),
  );

  final perkHeaders = ["Executive", "Role", "Monthly Allowance (HIDDEN)", "Assigned Vehicle"];
  for (var i = 0; i < perkHeaders.length; i++) {
    perksSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 3),
      TextCellValue(perkHeaders[i]),
      cellStyle: CellStyle(bold: true, backgroundColorHex: ExcelColor.grey300),
    );
  }

  final perkData = [
    ["Arthur Dent", "CEO", 12000, "Tesla Model S"],
    ["Ford Prefect", "CTO", 9500, "BMW i4"],
    ["Tricia McMillan", "CFO", 9800, "Audi e-tron"],
  ];

  for (var r = 0; r < perkData.length; r++) {
    final rowIdx = 4 + r;
    final row = perkData[r];
    perksSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: rowIdx), TextCellValue(row[0] as String));
    perksSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: rowIdx), TextCellValue(row[1] as String));
    perksSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: rowIdx), IntCellValue(row[2] as int));
    perksSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: rowIdx), TextCellValue(row[3] as String));
  }

  perksSheet.setColumnHidden(2, true); // Hide Column C

  perksSheet.setColumnWidth(0, 20);
  perksSheet.setColumnWidth(1, 15);
  perksSheet.setColumnWidth(3, 20);


  // ── Save ─────────────────────────────────────────────────────────────────
  const fileName = 'hidden_elements_multi_sheet.xlsx';

  if (kIsWeb) {
    final bytes = excel.save(fileName: fileName);
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Multi-sheet workbook with hidden elements generated!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'Sheets: Public Directory (All visible), Staff Salaries (Hidden Col/Row), Executive Perks (Hidden Col).';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) throw Exception('Failed to encode Excel file.');

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Hidden Columns Multi-Sheet Excel',
      fileName: fileName,
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      await File(outputFile).writeAsBytes(bytes);
      final size = await File(outputFile).length();
      return '✅ Multi-sheet workbook with hidden elements saved!\n'
          'Location: $outputFile\n'
          'Size: ${(size / 1024).toStringAsFixed(2)} KB\n'
          'Sheets: Public Directory (All visible), Staff Salaries (Hidden Col/Row), Executive Perks (Hidden Col).';
    }
    return 'Save cancelled.';
  }
}
