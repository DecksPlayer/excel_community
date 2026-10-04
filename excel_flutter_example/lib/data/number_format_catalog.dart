import 'package:excel_community/excel_community.dart';

/// Catalog of number formats shown in the Number Formats wiki and written to
/// its demo workbook. Each entry records sample values together with what
/// Excel displays for them, so the wiki and the generated file stay in sync.
///
/// Built-in IDs 23–26 are reserved by the spec (they behave as General) and
/// are therefore not listed.

enum NumFormatCategory {
  numbers('Numbers'),
  currency('Currency & Accounting'),
  percent('Percent & Scientific'),
  dateTime('Date & Time'),
  cjk('CJK Locale'),
  custom('Custom Codes');

  final String label;
  const NumFormatCategory(this.label);
}

class NumFormatSample {
  final CellValue value;

  /// What Excel displays for [value] (alignment padding trimmed).
  final String expected;

  const NumFormatSample(this.value, this.expected);
}

class NumFormatEntry {
  final NumFormatCategory category;
  final String name;
  final NumFormat format;

  /// Name of the `NumFormat` constant (e.g. `standard_4`); `null` for
  /// custom format codes.
  final String? constant;

  final List<NumFormatSample> samples;

  const NumFormatEntry({
    required this.category,
    required this.name,
    required this.format,
    this.constant,
    required this.samples,
  });

  int? get numFmtId => switch (format) {
        StandardNumFormat f => f.numFmtId,
        _ => null,
      };

  String get dartExpression => constant != null
      ? 'NumFormat.$constant'
      : 'NumFormat.custom(formatCode: ${dartStringLiteral(format.formatCode)})';

  /// [sample] rendered by the library (alignment padding trimmed).
  String display(NumFormatSample sample) => format.format(sample.value).trim();

  /// Complete, copyable snippet applying this format to a cell.
  String get snippet {
    final sample = samples.first;
    final comment = constant != null ? ' // ${format.formatCode}' : '';
    return "final cell = sheet.cell(CellIndex.indexByString('A1'));\n"
        'cell.value = ${cellValueCode(sample.value)};\n'
        'cell.cellStyle = CellStyle(\n'
        '  numberFormat: $dartExpression,$comment\n'
        ');\n'
        '// Excel shows: ${sample.expected}';
  }
}

/// Dart source that recreates [value].
String cellValueCode(CellValue value) => switch (value) {
      IntCellValue v => 'IntCellValue(${v.value})',
      DoubleCellValue v => 'DoubleCellValue(${v.value})',
      TextCellValue v => "TextCellValue('$v')",
      DateCellValue v => 'DateCellValue(year: ${v.year}, month: ${v.month}, day: ${v.day})',
      TimeCellValue v => 'TimeCellValue(hour: ${v.hour}, minute: ${v.minute}, second: ${v.second})',
      DateTimeCellValue v => 'DateTimeCellValue(year: ${v.year}, month: ${v.month}, day: ${v.day}, '
          'hour: ${v.hour}, minute: ${v.minute}, second: ${v.second})',
      BoolCellValue v => 'BoolCellValue(${v.value})',
      FormulaCellValue v => "FormulaCellValue('${v.formula}')",
    };

/// Short human readable label for a sample value (e.g. `2026-10-03`).
String sampleLabel(CellValue value) => switch (value) {
      TextCellValue v => '"$v"',
      DateCellValue v => '${v.year}-${_two(v.month)}-${_two(v.day)}',
      TimeCellValue v => '${_two(v.hour)}:${_two(v.minute)}:${_two(v.second)}',
      DateTimeCellValue v =>
        '${v.year}-${_two(v.month)}-${_two(v.day)} ${_two(v.hour)}:${_two(v.minute)}',
      _ => value.toString(),
    };

String _two(int n) => n.toString().padLeft(2, '0');

/// Dart string literal for [text], using a raw string when it contains `$`.
String dartStringLiteral(String text) {
  if (text.contains("'")) return 'r"$text"';
  if (text.contains(r'$') || text.contains(r'\')) return "r'$text'";
  return "'$text'";
}

const _pos = DoubleCellValue(1234.567);
const _neg = DoubleCellValue(-1234.567);
const _zero = IntCellValue(0);
const _date = DateCellValue(year: 2026, month: 10, day: 3);
const _time = TimeCellValue(hour: 14, minute: 30, second: 45);
const _dateTime = DateTimeCellValue(year: 2026, month: 10, day: 3, hour: 14, minute: 30, second: 45);


final List<NumFormatEntry> numFormatCatalog = [
  // --- Numbers -------------------------------------------------------------
  const NumFormatEntry(
    category: NumFormatCategory.numbers,
    name: 'General',
    format: NumFormat.standard_0,
    constant: 'standard_0',
    samples: [NumFormatSample(_pos, '1234.567'), NumFormatSample(_neg, '-1234.567')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.numbers,
    name: 'Integer',
    format: NumFormat.standard_1,
    constant: 'standard_1',
    samples: [NumFormatSample(_pos, '1235'), NumFormatSample(_neg, '-1235')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.numbers,
    name: 'Two Decimals',
    format: NumFormat.standard_2,
    constant: 'standard_2',
    samples: [NumFormatSample(_pos, '1234.57'), NumFormatSample(_neg, '-1234.57')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.numbers,
    name: 'Thousands Separator',
    format: NumFormat.standard_3,
    constant: 'standard_3',
    samples: [NumFormatSample(_pos, '1,235'), NumFormatSample(_neg, '-1,235')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.numbers,
    name: 'Thousands + Two Decimals',
    format: NumFormat.standard_4,
    constant: 'standard_4',
    samples: [NumFormatSample(_pos, '1,234.57'), NumFormatSample(_neg, '-1,234.57')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.numbers,
    name: 'Text (@)',
    format: NumFormat.standard_49,
    constant: 'standard_49',
    samples: [NumFormatSample(TextCellValue('00123'), '00123'), const NumFormatSample(_pos, '1234.567')],
  ),

  // --- Currency & Accounting -----------------------------------------------
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Currency',
    format: NumFormat.standard_5,
    constant: 'standard_5',
    samples: [NumFormatSample(_pos, r'$1,235'), NumFormatSample(_neg, r'($1,235)')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Currency, Red Negatives',
    format: NumFormat.standard_6,
    constant: 'standard_6',
    samples: [NumFormatSample(_pos, r'$1,235'), NumFormatSample(_neg, r'($1,235)')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Currency, Two Decimals',
    format: NumFormat.standard_7,
    constant: 'standard_7',
    samples: [NumFormatSample(_pos, r'$1,234.57'), NumFormatSample(_neg, r'($1,234.57)')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Currency, Two Decimals, Red',
    format: NumFormat.standard_8,
    constant: 'standard_8',
    samples: [NumFormatSample(_pos, r'$1,234.57'), NumFormatSample(_neg, r'($1,234.57)')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Accounting Integer',
    format: NumFormat.standard_37,
    constant: 'standard_37',
    samples: [NumFormatSample(_pos, '1,235'), NumFormatSample(_neg, '(1,235)')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Accounting Integer, Red',
    format: NumFormat.standard_38,
    constant: 'standard_38',
    samples: [NumFormatSample(_pos, '1,235'), NumFormatSample(_neg, '(1,235)')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Accounting Two Decimals',
    format: NumFormat.standard_39,
    constant: 'standard_39',
    samples: [NumFormatSample(_pos, '1,234.57'), NumFormatSample(_neg, '(1,234.57)')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Accounting Two Decimals, Red',
    format: NumFormat.standard_40,
    constant: 'standard_40',
    samples: [NumFormatSample(_pos, '1,234.57'), NumFormatSample(_neg, '(1,234.57)')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Accounting, Dash for Zero',
    format: NumFormat.standard_41,
    constant: 'standard_41',
    samples: [
      NumFormatSample(_pos, '1,235'),
      NumFormatSample(_neg, '(1,235)'),
      NumFormatSample(_zero, '-'),
    ],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: r'Accounting $, Dash for Zero',
    format: NumFormat.standard_42,
    constant: 'standard_42',
    samples: [
      NumFormatSample(_pos, r'$1,235'),
      NumFormatSample(_neg, r'$(1,235)'),
      NumFormatSample(_zero, r'$-'),
    ],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: 'Accounting, Two Decimals',
    format: NumFormat.standard_43,
    constant: 'standard_43',
    samples: [
      NumFormatSample(_pos, '1,234.57'),
      NumFormatSample(_neg, '(1,234.57)'),
      NumFormatSample(_zero, '-'),
    ],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.currency,
    name: r'Accounting $, Two Decimals',
    format: NumFormat.standard_44,
    constant: 'standard_44',
    samples: [
      NumFormatSample(_pos, r'$1,234.57'),
      NumFormatSample(_neg, r'$(1,234.57)'),
      NumFormatSample(_zero, r'$-'),
    ],
  ),

  // --- Percent, Scientific & Fractions -------------------------------------
  const NumFormatEntry(
    category: NumFormatCategory.percent,
    name: 'Percent',
    format: NumFormat.standard_9,
    constant: 'standard_9',
    samples: [
      NumFormatSample(DoubleCellValue(0.756), '76%'),
      NumFormatSample(DoubleCellValue(-0.0425), '-4%'),
    ],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.percent,
    name: 'Percent, Two Decimals',
    format: NumFormat.standard_10,
    constant: 'standard_10',
    samples: [
      NumFormatSample(DoubleCellValue(0.756), '75.60%'),
      NumFormatSample(DoubleCellValue(-0.0425), '-4.25%'),
    ],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.percent,
    name: 'Scientific',
    format: NumFormat.standard_11,
    constant: 'standard_11',
    samples: [
      NumFormatSample(DoubleCellValue(123456789), '1.23E+08'),
      NumFormatSample(DoubleCellValue(0.000123), '1.23E-04'),
    ],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.percent,
    name: 'Engineering',
    format: NumFormat.standard_48,
    constant: 'standard_48',
    samples: [
      NumFormatSample(DoubleCellValue(123456789), '123.5E+6'),
      NumFormatSample(DoubleCellValue(0.000123), '123.0E-6'),
    ],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.percent,
    name: 'Fraction, One Digit',
    format: NumFormat.standard_12,
    constant: 'standard_12',
    samples: [
      NumFormatSample(DoubleCellValue(1.75), '1 3/4'),
      NumFormatSample(DoubleCellValue(0.6), '3/5'),
    ],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.percent,
    name: 'Fraction, Two Digits',
    format: NumFormat.standard_13,
    constant: 'standard_13',
    samples: [
      NumFormatSample(DoubleCellValue(2.3125), '2  5/16'),
      NumFormatSample(DoubleCellValue(1234.567), '1234 55/97'),
    ],
  ),

  // --- Date & Time ----------------------------------------------------------
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Short Date',
    format: NumFormat.standard_14,
    constant: 'standard_14',
    samples: [NumFormatSample(_date, '10-03-26')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Day-Month-Year',
    format: NumFormat.standard_15,
    constant: 'standard_15',
    samples: [NumFormatSample(_date, '3-Oct-26')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Day-Month',
    format: NumFormat.standard_16,
    constant: 'standard_16',
    samples: [NumFormatSample(_date, '3-Oct')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Month-Year',
    format: NumFormat.standard_17,
    constant: 'standard_17',
    samples: [NumFormatSample(_date, 'Oct-26')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Month/Day/Year',
    format: NumFormat.standard_30,
    constant: 'standard_30',
    samples: [NumFormatSample(_date, '10/3/26')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Date and Time',
    format: NumFormat.standard_22,
    constant: 'standard_22',
    samples: [NumFormatSample(_dateTime, '10/3/26 14:30')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Time 12h',
    format: NumFormat.standard_18,
    constant: 'standard_18',
    samples: [NumFormatSample(_time, '2:30 PM')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Time 12h with Seconds',
    format: NumFormat.standard_19,
    constant: 'standard_19',
    samples: [NumFormatSample(_time, '2:30:45 PM')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Time 24h',
    format: NumFormat.standard_20,
    constant: 'standard_20',
    samples: [NumFormatSample(_time, '14:30')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Time 24h with Seconds',
    format: NumFormat.standard_21,
    constant: 'standard_21',
    samples: [NumFormatSample(_time, '14:30:45')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Minutes:Seconds',
    format: NumFormat.standard_45,
    constant: 'standard_45',
    samples: [NumFormatSample(_time, '30:45')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Elapsed Hours',
    format: NumFormat.standard_46,
    constant: 'standard_46',
    samples: [NumFormatSample(_time, '14:30:45')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.dateTime,
    name: 'Minutes:Seconds.Tenths',
    format: NumFormat.standard_47,
    constant: 'standard_47',
    samples: [NumFormatSample(_time, '30:45.0')],
  ),

  // --- CJK locale -------------------------------------------------------------
  const NumFormatEntry(
    category: NumFormatCategory.cjk,
    name: 'Year-Month-Day (Kanji)',
    format: NumFormat.standard_31,
    constant: 'standard_31',
    samples: [NumFormatSample(_date, '2026年10月3日')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.cjk,
    name: 'Hour-Minute (Kanji)',
    format: NumFormat.standard_32,
    constant: 'standard_32',
    samples: [NumFormatSample(_time, '14時30分')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.cjk,
    name: 'Hour-Minute-Second (Kanji)',
    format: NumFormat.standard_33,
    constant: 'standard_33',
    samples: [NumFormatSample(_time, '14時30分45秒')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.cjk,
    name: 'Taiwan Locale Date',
    format: NumFormat.standard_27,
    constant: 'standard_27',
    samples: [NumFormatSample(_date, '2026/10/3')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.cjk,
    name: 'Taiwan Locale Date and Time',
    format: NumFormat.standard_28,
    constant: 'standard_28',
    samples: [NumFormatSample(_dateTime, '2026/10/3 2:30 下午')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.cjk,
    name: 'Taiwan Locale Date (Kanji)',
    format: NumFormat.standard_29,
    constant: 'standard_29',
    samples: [NumFormatSample(_date, '2026年10月3日')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.cjk,
    name: 'Taiwan Locale Year-Month-Day (Kanji)',
    format: NumFormat.standard_36,
    constant: 'standard_36',
    samples: [NumFormatSample(_date, '2026月10日3日')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.cjk,
    name: 'AM/PM Hour-Minute (Kanji)',
    format: NumFormat.standard_34,
    constant: 'standard_34',
    samples: [NumFormatSample(_time, '下午2時30分')],
  ),
  const NumFormatEntry(
    category: NumFormatCategory.cjk,
    name: 'AM/PM Hour-Minute-Second (Kanji)',
    format: NumFormat.standard_35,
    constant: 'standard_35',
    samples: [NumFormatSample(_time, '下午2時30分45秒')],
  ),

  // --- Custom codes -----------------------------------------------------------
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Three Decimals',
    format: NumFormat.custom(formatCode: '#,##0.000'),
    samples: const [NumFormatSample(_pos, '1,234.567')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Currency Code Suffix',
    format: NumFormat.custom(formatCode: '#,##0.00 "USD"'),
    samples: const [NumFormatSample(_pos, '1,234.57 USD')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Euro Symbol',
    format: NumFormat.custom(formatCode: r'[$€-2] #,##0.00'),
    samples: const [NumFormatSample(_pos, '€ 1,234.57'), NumFormatSample(_neg, '-€ 1,234.57')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Unit Suffix',
    format: NumFormat.custom(formatCode: '0.00" kg"'),
    samples: const [NumFormatSample(_pos, '1234.57 kg')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Text Prefix',
    format: NumFormat.custom(formatCode: '"Total: "#,##0'),
    samples: const [NumFormatSample(_pos, 'Total: 1,235')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Leading Zeros',
    format: NumFormat.custom(formatCode: '00000'),
    samples: const [NumFormatSample(IntCellValue(42), '00042')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Phone Number Template',
    format: NumFormat.custom(formatCode: '000-000-0000'),
    samples: const [NumFormatSample(IntCellValue(5551234567), '555-123-4567')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Thousands (K)',
    format: NumFormat.custom(formatCode: '#,##0,"K"'),
    samples: const [NumFormatSample(IntCellValue(1234567), '1,235K')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Millions (M)',
    format: NumFormat.custom(formatCode: '0.0,,"M"'),
    samples: const [NumFormatSample(IntCellValue(12345678), '12.3M')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Fraction in Eighths',
    format: NumFormat.custom(formatCode: '# ?/8'),
    samples: const [NumFormatSample(DoubleCellValue(1.3), '1 2/8')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Percent, One Decimal',
    format: NumFormat.custom(formatCode: '0.0%'),
    samples: const [NumFormatSample(DoubleCellValue(0.756), '75.6%')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Colors per Section',
    format: NumFormat.custom(formatCode: '[Blue]#,##0;[Red]-#,##0;"zero"'),
    samples: const [
      NumFormatSample(_pos, '1,235'),
      NumFormatSample(_neg, '-1,235'),
      NumFormatSample(_zero, 'zero'),
    ],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Day/Month/Year',
    format: NumFormat.custom(formatCode: 'dd/mm/yyyy'),
    samples: const [NumFormatSample(_date, '03/10/2026')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Long Date',
    format: NumFormat.custom(formatCode: 'dddd, mmmm d, yyyy'),
    samples: const [NumFormatSample(_date, 'Saturday, October 3, 2026')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Month Name and Year',
    format: NumFormat.custom(formatCode: 'mmm yyyy'),
    samples: const [NumFormatSample(_date, 'Oct 2026')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'Weekday and Date',
    format: NumFormat.custom(formatCode: 'ddd dd-mmm'),
    samples: const [NumFormatSample(_date, 'Sat 03-Oct')],
  ),
  NumFormatEntry(
    category: NumFormatCategory.custom,
    name: 'ISO Date and Time',
    format: NumFormat.custom(formatCode: 'yyyy-mm-dd hh:mm:ss'),
    samples: const [NumFormatSample(_dateTime, '2026-10-03 14:30:45')],
  ),
];
