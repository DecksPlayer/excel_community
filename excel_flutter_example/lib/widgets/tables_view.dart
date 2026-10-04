import 'package:excel_community/excel_community.dart'
    show CellIndex, DoubleCellValue, ExcelTable, FormulaCellValue, IntCellValue, Sheet, SheetTables, TableStyle;
import 'package:flutter/material.dart';

import '../data/table_samples.dart';
import 'wiki/wiki_components.dart';

const _accent = Color(0xFF2563EB);

/// Excel tables wiki: the style gallery, table options and data helpers,
/// each card drawing a table created with `sheet.addTable`.
class TablesView extends StatefulWidget {
  final bool isGenerating;
  final VoidCallback onGenerate;
  final String status;

  const TablesView({
    super.key,
    required this.isGenerating,
    required this.onGenerate,
    required this.status,
  });

  @override
  State<TablesView> createState() => _TablesViewState();
}

class _TablesViewState extends State<TablesView> {
  final TextEditingController _searchController = TextEditingController();
  String _query = '';
  String _tab = 'styles';

  @override
  void dispose() {
    _searchController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return WikiPage(
      icon: Icons.table_chart_outlined,
      title: 'Excel Tables',
      description: '"Format as Table" with every built-in style, totals rows, banded rows or columns and '
          'structured references. Each preview is a table created with sheet.addTable '
          '(style colors are approximated from the Office theme).',
      accent: _accent,
      isGenerating: widget.isGenerating,
      onGenerate: widget.onGenerate,
      generateLabel: 'Generate Demo Workbook',
      status: widget.status,
      tabs: [
        WikiTab('styles', 'Styles (${builtInTableStyles.length})'),
        WikiTab('options', 'Options (${tableSamples.where((s) => s.category == 'options').length})'),
        WikiTab('data', 'Data & Formulas (${tableSamples.where((s) => s.category == 'data').length})'),
      ],
      selectedTab: _tab,
      onTabSelected: (tab) => setState(() => _tab = tab),
      child: _tab == 'styles' ? _buildStyles() : _buildSamples(_tab),
    );
  }

  Widget _buildStyles() {
    final query = _query.toLowerCase();
    final styles = builtInTableStyles
        .where((s) => s.$1.toLowerCase().contains(query) || s.$2.name.toLowerCase().contains(query))
        .toList();
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        WikiSearchField(
          controller: _searchController,
          hint: 'Search styles: light, medium 9, dark...',
          accent: _accent,
          onChanged: (v) => setState(() => _query = v),
        ),
        const SizedBox(height: 16),
        WikiGrid(
          itemCount: styles.length,
          extent: 250,
          itemBuilder: (context, index) {
            final (label, style) = styles[index];
            final kind = label.split(' ').first.toLowerCase();
            final number = label.split(' ').last;
            final sheet = tableSampleSheet();
            final table = sheet.addTable('A1:C4', name: 'Sales', style: style);
            return WikiCard(
              title: label,
              subtitle: style.name,
              accent: _accent,
              code: "sheet.addTable('A1:C4', name: 'Sales',\n"
                  '    style: TableStyle.$kind($number));',
              codeSummary: 'style: TableStyle.$kind($number)',
              previewLabel: 'Style Preview:',
              preview: WikiPreviewBox(child: TablePreview(sheet: sheet, table: table)),
            );
          },
        ),
      ],
    );
  }

  Widget _buildSamples(String category) {
    final samples = tableSamples.where((s) => s.category == category).toList();
    return WikiGrid(
      itemCount: samples.length,
      extent: 400,
      itemBuilder: (context, index) {
        final sample = samples[index];
        final sheet = tableSampleSheet();
        final table = sample.apply(sheet);
        final note = sample.note?.call(sheet);
        return WikiCard(
          title: sample.title,
          subtitle: sample.subtitle,
          code: sample.code,
          accent: _accent,
          previewLabel: 'Live Preview:',
          preview: WikiPreviewBox(
            alignment: Alignment.topLeft,
            padding: const EdgeInsets.all(8),
            child: SingleChildScrollView(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  TablePreview(sheet: sheet, table: table),
                  if (note != null) ...[
                    const SizedBox(height: 8),
                    Text(note,
                        style: const TextStyle(fontFamily: 'monospace', fontSize: 9.5, color: Color(0xFF334155))),
                  ],
                ],
              ),
            ),
          ),
        );
      },
    );
  }
}

/// Approximate colors of a built-in table style (Office theme accents).
class _StyleColors {
  final Color? headerFill;
  final Color headerText;
  final Color? band;
  final Color? bodyFill;
  final Color bodyText;
  final Color line;

  const _StyleColors({
    this.headerFill,
    this.headerText = const Color(0xFF000000),
    this.band,
    this.bodyFill,
    this.bodyText = const Color(0xFF000000),
    required this.line,
  });

  static const _accents = [
    Color(0xFF000000), // style 1, 8, 15, 22: black/gray
    Color(0xFF4472C4),
    Color(0xFFED7D31),
    Color(0xFFA5A5A5),
    Color(0xFFFFC000),
    Color(0xFF5B9BD5),
    Color(0xFF70AD47),
  ];

  static Color _tint(Color c, double amount) => Color.lerp(c, Colors.white, amount)!;

  factory _StyleColors.of(TableStyle? style) {
    if (style == null) return const _StyleColors(line: Color(0xFFE2E8F0));
    final match = RegExp(r'TableStyle(Light|Medium|Dark)(\d+)').firstMatch(style.name);
    if (match == null) return const _StyleColors(line: Color(0xFFCBD5E1));
    final kind = match.group(1)!;
    final n = int.parse(match.group(2)!);
    final accent = _accents[(n - 1) % 7];
    final gray = accent == _accents[0];
    final group = (n - 1) ~/ 7;
    switch (kind) {
      case 'Light':
        return switch (group) {
          0 => _StyleColors(band: _tint(accent, gray ? 0.85 : 0.8), line: accent),
          1 => _StyleColors(headerFill: accent, headerText: Colors.white, line: accent),
          _ => _StyleColors(band: _tint(accent, gray ? 0.85 : 0.8), line: accent),
        };
      case 'Medium':
        return switch (group) {
          0 => _StyleColors(
              headerFill: accent, headerText: Colors.white, band: _tint(accent, 0.8), line: _tint(accent, 0.4)),
          1 => _StyleColors(
              headerFill: accent,
              headerText: Colors.white,
              band: _tint(accent, 0.6),
              bodyFill: _tint(accent, 0.8),
              line: Colors.white),
          2 => _StyleColors(
              headerFill: Colors.black,
              headerText: Colors.white,
              band: _tint(accent, gray ? 0.75 : 0.6),
              bodyFill: _tint(accent, gray ? 0.85 : 0.8),
              line: Colors.white),
          _ => _StyleColors(
              headerFill: _tint(accent, 0.6), band: _tint(accent, 0.6), bodyFill: _tint(accent, 0.8), line: Colors.white),
        };
      default: // Dark
        return group == 0
            ? _StyleColors(
                headerFill: Colors.black,
                headerText: Colors.white,
                band: Color.lerp(accent, Colors.black, gray ? 0.3 : 0.25),
                bodyFill: gray ? const Color(0xFF737373) : accent,
                bodyText: Colors.white,
                line: Colors.black)
            : _StyleColors(
                headerFill: Colors.black,
                headerText: Colors.white,
                band: _tint(_accents[(n - 1) % 4 * 2 % 7], 0.6),
                bodyFill: _tint(_accents[(n - 1) % 4 * 2 % 7], 0.8),
                line: Colors.white);
    }
  }
}

/// Draws [table] with the values of [sheet] as Excel would show it.
class TablePreview extends StatelessWidget {
  final Sheet sheet;
  final ExcelTable table;

  const TablePreview({super.key, required this.sheet, required this.table});

  @override
  Widget build(BuildContext context) {
    final colors = _StyleColors.of(table.style);
    final start = CellIndex.indexByString(table.ref.split(':').first);
    final end = CellIndex.indexByString(table.ref.split(':').last);
    final rows = <TableRow>[];

    for (var r = start.rowIndex; r <= end.rowIndex; r++) {
      final isHeader = table.headerRowIndex == r;
      final isTotals = table.totalsRowIndex == r;
      final dataIndex = r - table.firstDataRow;
      final cells = <Widget>[];
      for (var c = start.columnIndex; c <= end.columnIndex; c++) {
        final value = sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r)).value;
        var text = value is FormulaCellValue ? '=${value.formula}' : value?.toString() ?? '';
        final firstOrLast = (table.showFirstColumn && c == start.columnIndex) ||
            (table.showLastColumn && c == end.columnIndex);
        Color? fill;
        var textColor = colors.bodyText;
        if (isHeader) {
          fill = colors.headerFill;
          textColor = colors.headerText;
        } else if (!isTotals) {
          final banded = (table.showRowStripes && dataIndex.isEven) ||
              (table.showColumnStripes && (c - start.columnIndex).isEven);
          fill = banded ? (colors.band ?? colors.bodyFill) : colors.bodyFill;
        } else {
          fill = colors.bodyFill;
        }
        cells.add(Container(
          height: 20,
          padding: const EdgeInsets.symmetric(horizontal: 4),
          color: fill,
          child: Row(
            children: [
              Expanded(
                child: Text(
                  text,
                  textAlign: value is IntCellValue || value is DoubleCellValue
                      ? TextAlign.right
                      : TextAlign.left,
                  maxLines: 1,
                  overflow: TextOverflow.ellipsis,
                  style: TextStyle(
                    fontSize: 9,
                    color: textColor,
                    fontWeight: isHeader || isTotals || firstOrLast ? FontWeight.bold : FontWeight.normal,
                    fontStyle: text.startsWith('=') ? FontStyle.italic : null,
                  ),
                ),
              ),
              if (isHeader && table.showFilterButtons)
                Icon(Icons.arrow_drop_down, size: 12, color: textColor.withValues(alpha: 0.8)),
            ],
          ),
        ));
      }
      rows.add(TableRow(
        decoration: BoxDecoration(
          border: Border(
            top: isTotals ? BorderSide(color: colors.line, width: 1.5) : BorderSide.none,
            bottom: BorderSide(color: colors.line, width: isHeader ? 1.2 : 0.5),
          ),
        ),
        children: cells,
      ));
    }

    return Table(
      defaultColumnWidth: const FlexColumnWidth(),
      children: rows,
    );
  }
}
