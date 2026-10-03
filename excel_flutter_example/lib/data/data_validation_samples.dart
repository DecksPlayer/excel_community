import 'package:excel_community/excel_community.dart';

/// One data validation example for the wiki and the demo workbook.
class ValidationSample {
  final String category; // lists, numbers, text
  final String title;
  final String subtitle;
  final String code;
  final DataValidation rule;

  /// Values checked with [DataValidation.accepts] in the preview.
  final List<CellValue?> tries;

  const ValidationSample({
    required this.category,
    required this.title,
    required this.subtitle,
    required this.code,
    required this.rule,
    required this.tries,
  });
}

final List<ValidationSample> validationSamples = [
  // --- Dropdown lists ---------------------------------------------------------
  ValidationSample(
    category: 'lists',
    title: 'Fixed List with Messages',
    subtitle: 'DataValidation.list([...])',
    code: "sheet.addDataValidation('B2:B100',\n"
        "  DataValidation.list(['Open', 'In progress', 'Done'])\n"
        "      .withPrompt('Status', 'Pick a status')\n"
        "      .withError('Invalid status', 'Use the list'),\n"
        ');',
    rule: DataValidation.list(['Open', 'In progress', 'Done'])
        .withPrompt('Status', 'Pick a status')
        .withError('Invalid status', 'Use the list'),
    tries: [TextCellValue('Done'), TextCellValue('done'), TextCellValue('Closed'), null],
  ),
  ValidationSample(
    category: 'lists',
    title: 'List from Another Sheet',
    subtitle: 'DataValidation.listFromRange()',
    code: "sheet.addDataValidation('C2:C100',\n"
        "  DataValidation.listFromRange('A2:A5',\n"
        "      sheetName: 'Lookup Lists'),\n"
        ');\n'
        "// formula1: 'Lookup Lists'!\$A\$2:\$A\$5",
    rule: DataValidation.listFromRange('A2:A5', sheetName: 'Lookup Lists')
        .withPrompt('Priority', 'Values come from Lookup Lists'),
    tries: [TextCellValue('High')],
  ),
  ValidationSample(
    category: 'lists',
    title: 'Required Yes / No',
    subtitle: 'allowBlank: false',
    code: "sheet.addDataValidation('H2:H100',\n"
        "  DataValidation.list(['Yes', 'No'])\n"
        '      .copyWith(allowBlank: false),\n'
        ');',
    rule: DataValidation.list(['Yes', 'No']).copyWith(allowBlank: false),
    tries: [TextCellValue('Yes'), TextCellValue('Maybe'), null],
  ),
  ValidationSample(
    category: 'lists',
    title: 'List without Arrow',
    subtitle: 'showDropdown: false',
    code: "sheet.addDataValidation('I2:I100',\n"
        "  DataValidation.list(['A', 'B', 'C'])\n"
        '      .copyWith(showDropdown: false),\n'
        ');\n'
        '// Values are still checked; no in-cell arrow',
    rule: DataValidation.list(['A', 'B', 'C']).copyWith(showDropdown: false),
    tries: [TextCellValue('B'), TextCellValue('D')],
  ),

  // --- Numbers, dates and times ---------------------------------------------
  ValidationSample(
    category: 'numbers',
    title: 'Whole Number Between',
    subtitle: 'DataValidation.wholeNumber()',
    code: "sheet.addDataValidation('D2:D100',\n"
        '  DataValidation.wholeNumber(\n'
        '      DataValidationOperator.between, 1, 40)\n'
        "      .withError('Hours', 'Whole hours from 1 to 40'),\n"
        ');',
    rule: DataValidation.wholeNumber(DataValidationOperator.between, 1, 40)
        .withError('Hours', 'Whole hours from 1 to 40'),
    tries: [IntCellValue(8), DoubleCellValue(7.5), IntCellValue(41)],
  ),
  ValidationSample(
    category: 'numbers',
    title: 'Decimal with Warning',
    subtitle: 'errorStyle: warning',
    code: "sheet.addDataValidation('J2:J100',\n"
        '  DataValidation.decimal(\n'
        '      DataValidationOperator.lessThanOrEqual, 0.3)\n'
        "      .withError('Large discount', 'Over 30%. Continue?',\n"
        '          style: DataValidationErrorStyle.warning),\n'
        ');',
    rule: DataValidation.decimal(DataValidationOperator.lessThanOrEqual, 0.3).withError(
        'Large discount', 'Over 30%. Continue?',
        style: DataValidationErrorStyle.warning),
    tries: [DoubleCellValue(0.15), DoubleCellValue(0.45)],
  ),
  ValidationSample(
    category: 'numbers',
    title: 'Date in 2026',
    subtitle: 'DataValidation.date()',
    code: "sheet.addDataValidation('E2:E100',\n"
        '  DataValidation.date(DataValidationOperator.between,\n'
        '      DateTime(2026, 1, 1), DateTime(2026, 12, 31)),\n'
        ');',
    rule: DataValidation.date(
        DataValidationOperator.between, DateTime(2026, 1, 1), DateTime(2026, 12, 31)),
    tries: [
      DateCellValue(year: 2026, month: 10, day: 3),
      DateCellValue(year: 2027, month: 1, day: 1),
    ],
  ),
  ValidationSample(
    category: 'numbers',
    title: 'Office Hours',
    subtitle: 'DataValidation.time()',
    code: "sheet.addDataValidation('F2:F100',\n"
        '  DataValidation.time(DataValidationOperator.between,\n'
        '      const Duration(hours: 9), const Duration(hours: 18)),\n'
        ');',
    rule: DataValidation.time(
        DataValidationOperator.between, const Duration(hours: 9), const Duration(hours: 18)),
    tries: [
      TimeCellValue(hour: 9, minute: 30, second: 0),
      TimeCellValue(hour: 19, minute: 0, second: 0),
    ],
  ),

  // --- Text and custom --------------------------------------------------------
  ValidationSample(
    category: 'text',
    title: 'Text Length',
    subtitle: 'DataValidation.textLength()',
    code: "sheet.addDataValidation('G2:G100',\n"
        '  DataValidation.textLength(\n'
        '      DataValidationOperator.lessThanOrEqual, 8)\n'
        "      .withPrompt('Code', 'Up to 8 characters'),\n"
        ');',
    rule: DataValidation.textLength(DataValidationOperator.lessThanOrEqual, 8)
        .withPrompt('Code', 'Up to 8 characters'),
    tries: [TextCellValue('TSK-001'), TextCellValue('TASK-00001')],
  ),
  ValidationSample(
    category: 'text',
    title: 'Custom Formula',
    subtitle: 'DataValidation.custom()',
    code: "sheet.addDataValidation('K2:K100',\n"
        "  DataValidation.custom('K2<=D2'),\n"
        ');\n'
        '// Relative to the first cell of the range',
    rule: DataValidation.custom('K2<=D2').withError('Overtime', 'Cannot exceed the planned hours'),
    tries: [IntCellValue(5)],
  ),
  ValidationSample(
    category: 'text',
    title: 'Input Message Only',
    subtitle: 'DataValidation.inputMessage()',
    code: "sheet.addDataValidation('A2:A100',\n"
        '  DataValidation.inputMessage(\n'
        "      'Task', 'Short description of the work'),\n"
        ');',
    rule: DataValidation.inputMessage('Task', 'Short description of the work'),
    tries: [TextCellValue('Anything goes')],
  ),
  ValidationSample(
    category: 'text',
    title: 'Check in Dart',
    subtitle: 'rule.accepts(value)',
    code: 'final rule = sheet.getDataValidation(\n'
        "    CellIndex.indexByString('B2'));\n"
        "if (rule?.accepts(TextCellValue(input)) == false) {\n"
        "  print('Invalid value');\n"
        '}',
    rule: DataValidation.list(['Open', 'In progress', 'Done']),
    tries: [TextCellValue('In progress'), TextCellValue('Blocked')],
  ),
];
