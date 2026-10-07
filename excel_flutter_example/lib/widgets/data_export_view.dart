import 'dart:convert';

import 'package:excel_community/excel_community.dart'
    show CellIndex, Excel, ExcelExport, ExportValueMode, SheetDataExt, SheetExport;
import 'package:flutter/material.dart';

import '../data/data_export_samples.dart';
import '../data/snippets/data_export.dart';
import 'wiki/wiki_components.dart';

const _accent = Color(0xFF7C3AED);

/// Data export & transformation wiki: every card runs one export/import
/// method on a sample workbook and shows the library's real output next to
/// the snippet that produced it.
class DataExportView extends StatefulWidget {
  final bool isGenerating;
  final VoidCallback onGenerate;
  final String status;

  const DataExportView({
    super.key,
    required this.isGenerating,
    required this.onGenerate,
    required this.status,
  });

  @override
  State<DataExportView> createState() => _DataExportViewState();
}

class _ExportExample {
  final String title;
  final String subtitle;
  final String code;
  final String output;

  const _ExportExample(this.title, this.subtitle, this.code, this.output);
}

class _DataExportViewState extends State<DataExportView> {
  String _tab = 'export';
  late final List<_ExportExample> _exportExamples;
  late final List<_ExportExample> _importExamples;
  late final List<List<String>> _sampleGrid;

  @override
  void initState() {
    super.initState();
    final excel = buildDataExportSampleWorkbook();
    final customers = excel['Customers'];
    _sampleGrid = customers.rows
        .map((row) => row.map((cell) => cell?.displayText ?? '').toList())
        .toList();

    _exportExamples = [
      _ExportExample(
        'Rows as Maps',
        'rowsAsMaps()',
        "final rows = excel['Customers'].rowsAsMaps();\n"
            '// Native Dart values keyed by the header row',
        describeRows(customers.rowsAsMaps()),
      ),
      _ExportExample(
        'Rows as Displayed Text',
        'rowsAsMaps(mode: displayText)',
        "final rows = excel['Customers'].rowsAsMaps(\n"
            '  mode: ExportValueMode.displayText,\n'
            ');\n'
            '// Number formats applied, as Excel shows them',
        describeRows(customers.rowsAsMaps(mode: ExportValueMode.displayText)),
      ),
      _ExportExample(
        'Value Grid',
        'rowsAsValues()',
        "final grid = excel['Customers'].rowsAsValues();\n"
            '// One list per row, header included',
        describeRows(customers.rowsAsValues()),
      ),
      _ExportExample(
        'Sheet to JSON',
        'toJson(indent: "  ")',
        "final json = excel['Customers'].toJson(indent: '  ');\n"
            '// Dates as YYYY-MM-DD, date-times as ISO 8601',
        customers.toJson(indent: '  '),
      ),
      _ExportExample(
        'Workbook to JSON',
        'excel.toJson()',
        "final json = excel.toJson(indent: '  ');\n"
            '// {"Customers": [...], "Orders": [...]}',
        excel.toJson(indent: '  '),
      ),
      _ExportExample(
        'Sheet to CSV',
        'toCsv()',
        "final csv = excel['Customers'].toCsv();\n"
            '// RFC 4180: "Smith, Luis" is quoted',
        customers.toCsv(lineTerminator: '\n'),
      ),
      _ExportExample(
        'CSV with Semicolons',
        "toCsv(separator: ';')",
        "final csv = excel['Customers'].toCsv(\n"
            "  separator: ';',\n"
            "  lineTerminator: '\\n',\n"
            ');',
        customers.toCsv(separator: ';', lineTerminator: '\n'),
      ),
    ];

    // Import: JSON -> maps -> rows.
    final imported = Excel.createExcel()['Imported'];
    final maps = (jsonDecode(dataExportImportJson) as List).cast<Map<String, Object?>>();
    imported.appendRowsFromMaps(maps);

    final existing = buildDataExportSampleWorkbook()['Customers'];
    existing.appendRowsFromMaps([
      {'Name': 'Eva', 'Age': 40, 'City': 'Lima'},
    ]);

    final roundTrip = Excel.createExcel()['Copy'];
    roundTrip.appendRowsFromMaps(customers.rowsAsMaps());
    final copyJoined = roundTrip.cell(CellIndex.indexByString('D2')).value;

    _importExamples = [
      _ExportExample(
        'JSON to a New Sheet',
        'appendRowsFromMaps(maps)',
        'final maps = (jsonDecode(json) as List)\n'
            '    .cast<Map<String, Object?>>();\n'
            "excel['Imported'].appendRowsFromMaps(maps);\n"
            '// Header row written from the keys',
        '$dataExportImportJson\n\n→ sheet rows:\n${imported.toCsv(lineTerminator: '\n')}',
      ),
      _ExportExample(
        'Append to Existing Columns',
        'matches the header row',
        "excel['Customers'].appendRowsFromMaps([\n"
            "  {'Name': 'Eva', 'Age': 40, 'City': 'Lima'},\n"
            ']);\n'
            '// "City" is added as a new column',
        existing.toCsv(lineTerminator: '\n'),
      ),
      _ExportExample(
        'Round Trip',
        'rowsAsMaps() → appendRowsFromMaps()',
        "final rows = excel['Customers'].rowsAsMaps();\n"
            "excel['Copy'].appendRowsFromMaps(rows);\n"
            '// DateTime without time → date cell',
        "Copy!D2 = ${copyJoined.runtimeType}($copyJoined)\n\n"
            '${roundTrip.toJson(indent: '  ')}',
      ),
    ];
  }

  @override
  Widget build(BuildContext context) {
    final examples = _tab == 'export' ? _exportExamples : _importExamples;
    return WikiPage(
      icon: Icons.data_object,
      title: 'Data Export & Transformation',
      description: 'Turn worksheets into Dart maps, value grids, JSON or CSV, and write maps back as rows. '
          'Every output below is produced live by the library from the sample sheet.',
      accent: _accent,
      isGenerating: widget.isGenerating,
      onGenerate: widget.onGenerate,
      generateLabel: 'Generate Demo Workbook',
      status: widget.status,
      tabs: [
        WikiTab('export', 'Export (${_exportExamples.length})'),
        WikiTab('import', 'Import (${_importExamples.length})'),
        const WikiTab('code', 'Full Example (Code)'),
      ],
      selectedTab: _tab,
      onTabSelected: (tab) => setState(() => _tab = tab),
      child: _tab == 'code'
          ? const WikiCodeCard(
              title: 'Data Export & Transformation Demo Workbook Code',
              subtitle:
                  'Complete code configuring JSON and CSV exports, sheet writing, and importing maps back into Excel',
              code: dataExportSnippet,
              accent: _accent,
            )
          : Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                _SampleTable(grid: _sampleGrid),
                const SizedBox(height: 16),
          WikiGrid(
            itemCount: examples.length,
            extent: 400,
            itemBuilder: (context, index) {
              final example = examples[index];
              return WikiCard(
                title: example.title,
                subtitle: example.subtitle,
                code: example.code,
                accent: _accent,
                previewLabel: 'Live Output:',
                preview: WikiPreviewBox(
                  alignment: Alignment.topLeft,
                  padding: const EdgeInsets.all(8),
                  child: SingleChildScrollView(
                    child: SelectableText(
                      example.output,
                      style: const TextStyle(
                        fontFamily: 'monospace',
                        fontSize: 9.5,
                        height: 1.35,
                        color: Color(0xFF1E293B),
                      ),
                    ),
                  ),
                ),
              );
            },
          ),
        ],
      ),
    );
  }
}

/// The sample "Customers" sheet as Excel displays it.
class _SampleTable extends StatelessWidget {
  final List<List<String>> grid;

  const _SampleTable({required this.grid});

  @override
  Widget build(BuildContext context) {
    return Container(
      width: double.infinity,
      padding: const EdgeInsets.all(12),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(12),
        border: Border.all(color: const Color(0xFFE2E8F0)),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text(
            "Sample sheet: excel['Customers']",
            style: TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)),
          ),
          const SizedBox(height: 8),
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            child: Table(
              defaultColumnWidth: const IntrinsicColumnWidth(),
              border: TableBorder.all(color: const Color(0xFFE2E8F0)),
              children: [
                for (var r = 0; r < grid.length; r++)
                  TableRow(
                    decoration: BoxDecoration(
                      color: r == 0 ? _accent.withValues(alpha: 0.1) : null,
                    ),
                    children: [
                      for (final value in grid[r])
                        Padding(
                          padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
                          child: Text(
                            value,
                            style: TextStyle(
                              fontSize: 11,
                              fontWeight: r == 0 ? FontWeight.bold : FontWeight.normal,
                              color: const Color(0xFF1E293B),
                            ),
                          ),
                        ),
                    ],
                  ),
              ],
            ),
          ),
        ],
      ),
    );
  }
}
