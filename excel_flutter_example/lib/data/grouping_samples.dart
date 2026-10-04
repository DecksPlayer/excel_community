import 'package:excel_community/excel_community.dart';

/// One grouping example: builds a small sheet and groups it like the
/// snippet does.
class GroupingSample {
  final String category; // rows, columns, options
  final String title;
  final String subtitle;
  final String code;
  final Sheet Function() build;

  /// Optional text output (e.g. `rowGroups`) shown under the preview.
  final String Function(Sheet sheet)? note;

  const GroupingSample({
    required this.category,
    required this.title,
    required this.subtitle,
    required this.code,
    required this.build,
    this.note,
  });
}

/// Quarter of monthly rows: Jan, Feb, Mar and a "Q1 total" summary row.
Sheet _monthsSheet({bool summaryAbove = false}) {
  final sheet = Excel.createExcel()['Sheet1'];
  sheet.appendRow([TextCellValue('Month'), TextCellValue('Sales')]);
  if (summaryAbove) sheet.appendRow([TextCellValue('Q1 total'), IntCellValue(360)]);
  for (final (month, value) in [('Jan', 100), ('Feb', 120), ('Mar', 140)]) {
    sheet.appendRow([TextCellValue(month), IntCellValue(value)]);
  }
  if (!summaryAbove) sheet.appendRow([TextCellValue('Q1 total'), IntCellValue(360)]);
  sheet.appendRow([TextCellValue('Apr'), IntCellValue(90)]);
  return sheet;
}

/// Two quarters: months grouped per quarter, quarters grouped in a half.
Sheet _halfYearSheet() {
  final sheet = Excel.createExcel()['Sheet1'];
  sheet.appendRow([TextCellValue('Period'), TextCellValue('Sales')]);
  for (final row in [
    ('Jan', 100), ('Feb', 120), ('Mar', 140), ('Q1', 360),
    ('Apr', 90), ('May', 110), ('Jun', 130), ('Q2', 330),
    ('H1', 690),
  ]) {
    sheet.appendRow([TextCellValue(row.$1), IntCellValue(row.$2)]);
  }
  return sheet;
}

/// Columns: Region, Jan, Feb, Mar, Q1.
Sheet _columnsSheet() {
  final sheet = Excel.createExcel()['Sheet1'];
  sheet.appendRow([
    TextCellValue('Region'), TextCellValue('Jan'), TextCellValue('Feb'),
    TextCellValue('Mar'), TextCellValue('Q1'),
  ]);
  sheet.appendRow([TextCellValue('North'), IntCellValue(10), IntCellValue(12), IntCellValue(14), IntCellValue(36)]);
  sheet.appendRow([TextCellValue('South'), IntCellValue(8), IntCellValue(9), IntCellValue(11), IntCellValue(28)]);
  return sheet;
}

String _groups(List<OutlineGroup> groups) =>
    groups.isEmpty ? '[]' : groups.map((g) => '${g.start}..${g.end} L${g.level}${g.collapsed ? ' collapsed' : ''}').join(', ');

final List<GroupingSample> groupingSamples = [
  // --- Rows ---------------------------------------------------------------
  GroupingSample(
    category: 'rows',
    title: 'Group Rows',
    subtitle: 'groupRows(start, end)',
    code: '// Jan-Mar (rows 2-4); "Q1 total" below is the\n'
        '// summary row with the -/+ button\n'
        'sheet.groupRows(1, 3);',
    build: () => _monthsSheet()..groupRows(1, 3),
  ),
  GroupingSample(
    category: 'rows',
    title: 'Start Collapsed',
    subtitle: 'groupRows(..., collapsed: true)',
    code: '// Rows are hidden; click + to expand\n'
        'sheet.groupRows(1, 3, collapsed: true);',
    build: () => _monthsSheet()..groupRows(1, 3, collapsed: true),
  ),
  GroupingSample(
    category: 'rows',
    title: 'Nested Groups',
    subtitle: 'levels 1 and 2',
    code: 'sheet.groupRows(1, 8);           // H1: level 1\n'
        'sheet.groupRows(1, 3);           // Q1 months: level 2\n'
        'sheet.groupRows(5, 7, collapsed: true); // Q2',
    build: () => _halfYearSheet()
      ..groupRows(1, 8)
      ..groupRows(1, 3)
      ..groupRows(5, 7, collapsed: true),
    note: (sheet) => 'rowGroups → ${_groups(sheet.rowGroups)}',
  ),
  GroupingSample(
    category: 'rows',
    title: 'Ungroup One Level',
    subtitle: 'ungroupRows(start, end)',
    code: 'sheet.groupRows(1, 8);\n'
        'sheet.groupRows(1, 3);\n'
        '// Removes the inner level only\n'
        'sheet.ungroupRows(1, 3);',
    build: () => _halfYearSheet()
      ..groupRows(1, 8)
      ..groupRows(1, 3)
      ..ungroupRows(1, 3),
    note: (sheet) => 'getRowOutlineLevel(2) → ${sheet.getRowOutlineLevel(2)}',
  ),

  // --- Columns --------------------------------------------------------------
  GroupingSample(
    category: 'columns',
    title: 'Group Columns',
    subtitle: 'groupColumns(start, end)',
    code: '// Jan-Mar (columns B-D); "Q1" is the summary\n'
        'sheet.groupColumns(1, 3);',
    build: () => _columnsSheet()..groupColumns(1, 3),
  ),
  GroupingSample(
    category: 'columns',
    title: 'Collapsed Columns',
    subtitle: 'collapsed: true',
    code: 'sheet.groupColumns(1, 3, collapsed: true);',
    build: () => _columnsSheet()..groupColumns(1, 3, collapsed: true),
  ),
  GroupingSample(
    category: 'columns',
    title: 'Rows and Columns',
    subtitle: 'both outline bars',
    code: 'sheet.groupColumns(1, 3);\n'
        '// Regions in one group above "Total"\n'
        'sheet.groupRows(1, 2);',
    build: () {
      final sheet = _columnsSheet()..groupColumns(1, 3);
      sheet.appendRow([TextCellValue('Total'), IntCellValue(18), IntCellValue(21), IntCellValue(25), IntCellValue(64)]);
      return sheet..groupRows(1, 2);
    },
  ),

  // --- Options ----------------------------------------------------------------
  GroupingSample(
    category: 'options',
    title: 'Summary Row Above',
    subtitle: 'OutlineSettings(summaryBelow: false)',
    code: 'sheet.outlineSettings =\n'
        '    const OutlineSettings(summaryBelow: false);\n'
        '// "Q1 total" (row 2) is above Jan-Mar\n'
        'sheet.groupRows(2, 4);',
    build: () => _monthsSheet(summaryAbove: true)
      ..outlineSettings = const OutlineSettings(summaryBelow: false)
      ..groupRows(2, 4),
  ),
  GroupingSample(
    category: 'options',
    title: 'Collapse from Code',
    subtitle: 'collapseRowGroup / expandRowGroup',
    code: 'sheet.groupRows(1, 8);\n'
        'sheet.groupRows(1, 3, collapsed: true);\n'
        'sheet.collapseRowGroup(1, 8);\n'
        '// Q1 stays collapsed inside H1\n'
        'sheet.expandRowGroup(1, 8);',
    build: () => _halfYearSheet()
      ..groupRows(1, 8)
      ..groupRows(1, 3, collapsed: true)
      ..collapseRowGroup(1, 8)
      ..expandRowGroup(1, 8),
    note: (sheet) => 'hidden rows → ${sheet.getHiddenRows.toList()..sort()}',
  ),
  GroupingSample(
    category: 'options',
    title: 'Groups Move with Rows',
    subtitle: 'insertRow()',
    code: 'sheet.groupRows(1, 3);\n'
        '// A new first row shifts the group down\n'
        'sheet.insertRow(0);',
    build: () {
      final sheet = _monthsSheet()..groupRows(1, 3);
      sheet.insertRow(0);
      sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('(new row)'));
      return sheet;
    },
    note: (sheet) => 'rowGroups → ${_groups(sheet.rowGroups)}',
  ),
];
