const String dataValidationSnippet = r'''
import 'dart:io';
import 'package:excel_community/excel_community.dart';

/// Generates the complete Data Validation demo workbook with 11 validated
/// columns, dynamic dropdowns, formula validation, and cross-sheet lookup.
void generateDataValidationWorkbook() {
  final excel = Excel.createExcel();
  final tasks = excel['Tasks'];

  // 1. Setup column headers with custom styling
  final columns = [
    'Task', 'Status', 'Priority', 'Hours', 'Due date',
    'Start time', 'Code', 'Billable', 'Grade', 'Discount',
    'Overtime', 'Priority name',
  ];
  final headerStyle = CellStyle(
    bold: true,
    fontColorHex: ExcelColor.white,
    backgroundColorHex: ExcelColor.fromHexString('#0D9488'),
  );
  for (var col = 0; col < columns.length; col++) {
    tasks.updateCell(
      CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0),
      TextCellValue(columns[col]),
      cellStyle: headerStyle,
    );
    tasks.setColumnWidth(col, 16);
  }

  // 2. Add validation rules for rows 2 to 100
  // Column A: Input message only
  tasks.addDataValidation(
    'A2:A100',
    DataValidation.inputMessage('Task', 'Short description of the work'),
  );

  // Column B: Fixed dropdown list with prompt & error messages
  tasks.addDataValidation(
    'B2:B100',
    DataValidation.list(['Open', 'In progress', 'Done'])
        .withPrompt('Status', 'Pick a status')
        .withError('Invalid status', 'Select an option from the list'),
  );

  // Column C: Dropdown list referencing 'Lookup Lists' sheet (A2:A5)
  // Uses listFromRange or OFFSET/COUNTA for dynamic growing lists
  tasks.addDataValidation(
    'C2:C100',
    DataValidation.listFromRange('A2:A5', sheetName: 'Lookup Lists')
        .withPrompt('Priority', 'Pick a code; column L shows its name'),
  );

  // Column D: Whole number between 1 and 40
  tasks.addDataValidation(
    'D2:D100',
    DataValidation.wholeNumber(DataValidationOperator.between, 1, 40)
        .withError('Hours', 'Whole hours from 1 to 40'),
  );

  // Column E: Date in 2026
  tasks.addDataValidation(
    'E2:E100',
    DataValidation.date(
      DataValidationOperator.between,
      DateTime(2026, 1, 1),
      DateTime(2026, 12, 31),
    ),
  );

  // Column F: Time between 09:00 and 18:00
  tasks.addDataValidation(
    'F2:F100',
    DataValidation.time(
      DataValidationOperator.between,
      const Duration(hours: 9),
      const Duration(hours: 18),
    ),
  );

  // Column G: Text length <= 8 characters
  tasks.addDataValidation(
    'G2:G100',
    DataValidation.textLength(DataValidationOperator.lessThanOrEqual, 8)
        .withPrompt('Code', 'Up to 8 characters'),
  );

  // Column H: Required Yes / No (blanks disallowed)
  tasks.addDataValidation(
    'H2:H100',
    DataValidation.list(['Yes', 'No']).copyWith(allowBlank: false),
  );

  // Column I: List without in-cell dropdown arrow
  tasks.addDataValidation(
    'I2:I100',
    DataValidation.list(['A', 'B', 'C']).copyWith(showDropdown: false),
  );

  // Column J: Decimal discount with Warning icon
  tasks.addDataValidation(
    'J2:J100',
    DataValidation.decimal(DataValidationOperator.lessThanOrEqual, 0.3)
        .withError('Large discount', 'Over 30%. Continue?',
            style: DataValidationErrorStyle.warning),
  );

  // Column K: Custom formula (Overtime <= Hours)
  tasks.addDataValidation(
    'K2:K100',
    DataValidation.custom('K2<=D2')
        .withError('Overtime', 'Cannot exceed planned hours'),
  );

  // 3. Append sample valid rows
  const rows = [
    ('Write release notes', 'In progress', 'P3', 6, 10, 15, 9, 30, 'Yes', 'A', 0.10, 4),
    ('Fix login bug', 'Open', 'P4', 12, 10, 8, 10, 0, 'Yes', 'A', 0.00, 3),
    ('Update docs', 'Done', 'P1', 3, 9, 22, 14, 15, 'No', 'C', 0.05, 0),
    ('Code review', 'In progress', 'P2', 4, 10, 20, 11, 0, 'Yes', 'B', 0.15, 1),
    ('Design mockups', 'Open', 'P3', 16, 11, 3, 9, 0, 'Yes', 'A', 0.20, 6),
    ('Database backup', 'Done', 'P2', 2, 8, 30, 17, 45, 'No', 'B', 0.00, 0),
    ('Customer demo', 'Open', 'P4', 8, 12, 1, 15, 30, 'Yes', 'A', 0.25, 2),
    ('Refactor tests', 'In progress', 'P1', 10, 11, 18, 13, 0, 'No', 'C', 0.30, 5),
  ];
  var n = 1;
  for (final r in rows) {
    tasks.appendRow([
      TextCellValue(r.$1),
      TextCellValue(r.$2),
      TextCellValue(r.$3),
      IntCellValue(r.$4),
      DateCellValue(year: 2026, month: r.$5, day: r.$6),
      TimeCellValue(hour: r.$7, minute: r.$8, second: 0),
      TextCellValue('TSK-${(n++).toString().padLeft(3, '0')}'),
      TextCellValue(r.$9),
      TextCellValue(r.$10),
      DoubleCellValue(r.$11),
      IntCellValue(r.$12),
    ]);
  }

  // Column L: VLOOKUP formula linking to 'Lookup Lists'
  for (var row = 2; row <= 100; row++) {
    tasks.updateCell(
      CellIndex.indexByString('L$row'),
      FormulaCellValue(
        "IF(C$row=\"\",\"\",IFERROR(VLOOKUP(C$row,'Lookup Lists'!\$A:\$B,2,FALSE),\"?\"))",
      ),
    );
  }

  // 4. Create 'Lookup Lists' source table
  final lists = excel['Lookup Lists'];
  lists.appendRow([TextCellValue('Code'), TextCellValue('Priority')]);
  const priorityLookup = {'P1': 'Low', 'P2': 'Medium', 'P3': 'High', 'P4': 'Critical'};
  for (final entry in priorityLookup.entries) {
    lists.appendRow([TextCellValue(entry.key), TextCellValue(entry.value)]);
  }

  // 5. Save the workbook
  final bytes = excel.save(fileName: 'data_validation_example.xlsx');
  // Or on Desktop / Server:
  // File('data_validation_example.xlsx').writeAsBytesSync(excel.encode()!);
}
''';
