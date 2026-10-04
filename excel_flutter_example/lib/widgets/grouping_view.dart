import 'dart:math' as math;

import 'package:excel_community/excel_community.dart'
    show CellIndex, OutlineGroup, Sheet, SheetDimensions, SheetGrouping, getColumnAlphabet;
import 'package:flutter/material.dart';

import '../data/grouping_samples.dart';
import 'wiki/wiki_components.dart';

const _accent = Color(0xFFEA580C);

/// Row & column grouping wiki. Each card renders its sheet with Excel's
/// outline bars; the +/- buttons call the library's collapse/expand methods.
class GroupingView extends StatefulWidget {
  final bool isGenerating;
  final VoidCallback onGenerate;
  final String status;

  const GroupingView({
    super.key,
    required this.isGenerating,
    required this.onGenerate,
    required this.status,
  });

  @override
  State<GroupingView> createState() => _GroupingViewState();
}

class _GroupingViewState extends State<GroupingView> {
  String _tab = 'rows';

  static const _tabs = {
    'rows': 'Row Groups',
    'columns': 'Column Groups',
    'options': 'Options & Code',
  };

  @override
  Widget build(BuildContext context) {
    final samples = groupingSamples.where((s) => s.category == _tab).toList();
    return WikiPage(
      icon: Icons.account_tree_outlined,
      title: 'Row & Column Grouping',
      description: 'Nested, collapsible groups like Excel\'s Data > Group. '
          'Click the +/- buttons in each preview: they call collapseRowGroup / expandRowGroup on the sheet.',
      accent: _accent,
      isGenerating: widget.isGenerating,
      onGenerate: widget.onGenerate,
      generateLabel: 'Generate Demo Workbook',
      status: widget.status,
      tabs: [
        for (final entry in _tabs.entries)
          WikiTab(entry.key,
              '${entry.value} (${groupingSamples.where((s) => s.category == entry.key).length})'),
      ],
      selectedTab: _tab,
      onTabSelected: (tab) => setState(() => _tab = tab),
      child: WikiGrid(
        itemCount: samples.length,
        extent: 470,
        itemBuilder: (context, index) {
          final sample = samples[index];
          return WikiCard(
            key: ValueKey(sample.title),
            title: sample.title,
            subtitle: sample.subtitle,
            code: sample.code,
            accent: _accent,
            previewLabel: 'Live Preview (click +/-):',
            preview: WikiPreviewBox(
              alignment: Alignment.topLeft,
              padding: const EdgeInsets.all(8),
              child: OutlinePreview(sample: sample),
            ),
          );
        },
      ),
    );
  }
}

/// Renders a sheet with its row/column outline bars. Holds its own [Sheet]
/// so the +/- buttons change the library state and the preview redraws.
class OutlinePreview extends StatefulWidget {
  final GroupingSample sample;

  const OutlinePreview({super.key, required this.sample});

  @override
  State<OutlinePreview> createState() => _OutlinePreviewState();
}

class _OutlinePreviewState extends State<OutlinePreview> {
  late Sheet _sheet = widget.sample.build();

  static const double _cellW = 46;
  static const double _cellH = 18;
  static const double _headerW = 20;
  static const double _level = 12;

  @override
  void didUpdateWidget(covariant OutlinePreview oldWidget) {
    super.didUpdateWidget(oldWidget);
    if (oldWidget.sample != widget.sample) _sheet = widget.sample.build();
  }

  int _maxLevel(Map<int, int> levels) => levels.isEmpty ? 0 : levels.values.reduce(math.max);

  @override
  Widget build(BuildContext context) {
    final sheet = _sheet;
    final rowCount = math.max(sheet.maxRows, 1);
    final colCount = math.max(sheet.maxColumns, 1);
    final rowLevels = {for (var r = 0; r < rowCount; r++) r: sheet.getRowOutlineLevel(r)}
      ..removeWhere((_, l) => l == 0);
    final colLevels = {for (var c = 0; c < colCount; c++) c: sheet.getColumnOutlineLevel(c)}
      ..removeWhere((_, l) => l == 0);
    final maxRowLevel = _maxLevel(rowLevels);
    final maxColLevel = _maxLevel(colLevels);
    final rowBarW = maxRowLevel * _level;

    final visibleRows = [for (var r = 0; r < rowCount; r++) if (!sheet.isRowHidden(r)) r];
    final visibleCols = [for (var c = 0; c < colCount; c++) if (!sheet.isColumnHidden(c)) c];

    // Summary index -> group, per level.
    final settings = sheet.outlineSettings;
    final rowSummaries = {
      for (final g in sheet.rowGroups)
        (settings.summaryBelow ? g.end + 1 : g.start - 1, g.level): g,
    };
    final colSummaries = {
      for (final g in sheet.columnGroups)
        (settings.summaryRight ? g.end + 1 : g.start - 1, g.level): g,
    };

    Widget button(OutlineGroup group, bool rows) => InkWell(
          onTap: () => setState(() {
            if (rows) {
              group.collapsed
                  ? sheet.expandRowGroup(group.start, group.end)
                  : sheet.collapseRowGroup(group.start, group.end);
            } else {
              group.collapsed
                  ? sheet.expandColumnGroup(group.start, group.end)
                  : sheet.collapseColumnGroup(group.start, group.end);
            }
          }),
          child: Container(
            width: 10,
            height: 10,
            alignment: Alignment.center,
            decoration: BoxDecoration(
              color: Colors.white,
              border: Border.all(color: const Color(0xFF475569)),
            ),
            child: Text(group.collapsed ? '+' : '−',
                style: const TextStyle(fontSize: 8, height: 1, fontWeight: FontWeight.bold)),
          ),
        );

    Widget rowBar(int row) => SizedBox(
          width: rowBarW,
          height: _cellH,
          child: Row(
            children: [
              for (var level = 1; level <= maxRowLevel; level++)
                SizedBox(
                  width: _level,
                  child: Center(
                    child: rowSummaries[(row, level)] != null
                        ? button(rowSummaries[(row, level)]!, true)
                        : (rowLevels[row] ?? 0) >= level
                            ? Container(width: 1.5, height: _cellH, color: const Color(0xFF94A3B8))
                            : null,
                  ),
                ),
            ],
          ),
        );

    Widget colBar(int level) => Row(
          children: [
            SizedBox(width: rowBarW + _headerW),
            for (final c in visibleCols)
              SizedBox(
                width: _cellW,
                height: _level,
                child: Center(
                  child: colSummaries[(c, level)] != null
                      ? button(colSummaries[(c, level)]!, false)
                      : (colLevels[c] ?? 0) >= level
                          ? Container(height: 1.5, width: _cellW, color: const Color(0xFF94A3B8))
                          : null,
                ),
              ),
          ],
        );

    Widget box(String text, {bool header = false, double width = _cellW}) => Container(
          width: width,
          height: _cellH,
          alignment: header ? Alignment.center : Alignment.centerLeft,
          padding: const EdgeInsets.symmetric(horizontal: 3),
          decoration: BoxDecoration(
            color: header ? const Color(0xFFF1F5F9) : Colors.white,
            border: Border.all(color: const Color(0xFFE2E8F0), width: 0.5),
          ),
          child: Text(text,
              maxLines: 1,
              overflow: TextOverflow.clip,
              style: TextStyle(fontSize: 9, color: header ? const Color(0xFF64748B) : const Color(0xFF1E293B))),
        );

    final note = widget.sample.note?.call(sheet);
    return SingleChildScrollView(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                for (var level = 1; level <= maxColLevel; level++) colBar(level),
                Row(children: [
                  SizedBox(width: rowBarW),
                  box('', header: true, width: _headerW),
                  for (final c in visibleCols) box(getColumnAlphabet(c), header: true),
                ]),
                for (final r in visibleRows)
                  Row(children: [
                    rowBar(r),
                    box('${r + 1}', header: true, width: _headerW),
                    for (final c in visibleCols)
                      box(sheet
                              .cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r))
                              .value
                              ?.toString() ??
                          ''),
                  ]),
              ],
            ),
          ),
          if (note != null) ...[
            const SizedBox(height: 6),
            Text(note,
                style: const TextStyle(fontFamily: 'monospace', fontSize: 9.5, color: Color(0xFF334155))),
          ],
        ],
      ),
    );
  }
}
