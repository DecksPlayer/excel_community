import 'package:excel_community/excel_community.dart' show CellValue, DoubleCellValue, IntCellValue;
import 'package:flutter/material.dart';

import '../data/number_format_catalog.dart';
import '../data/snippets/number_formats.dart';
import 'wiki/wiki_components.dart';

const _accent = Colors.teal;

/// Interactive wiki of every built-in number format plus common custom
/// format codes. Each card renders sample values with the library's own
/// `NumFormat.format()` and has a copyable snippet for that format.
class NumberFormatsView extends StatefulWidget {
  final bool isGenerating;
  final VoidCallback onGenerate;
  final String status;

  const NumberFormatsView({
    super.key,
    required this.isGenerating,
    required this.onGenerate,
    required this.status,
  });

  @override
  State<NumberFormatsView> createState() => _NumberFormatsViewState();
}

class _NumberFormatsViewState extends State<NumberFormatsView> {
  final TextEditingController _searchController = TextEditingController();
  String _searchQuery = '';
  String _tab = 'numbers';

  NumFormatCategory get _category {
    try {
      return NumFormatCategory.values.byName(_tab);
    } catch (_) {
      return NumFormatCategory.numbers;
    }
  }

  @override
  void dispose() {
    _searchController.dispose();
    super.dispose();
  }

  List<NumFormatEntry> get _visibleEntries {
    final query = _searchQuery.trim().toLowerCase();
    if (query.isEmpty) {
      return numFormatCatalog.where((e) => e.category == _category).toList();
    }
    // A search looks through every category.
    return numFormatCatalog.where((e) {
      return e.name.toLowerCase().contains(query) ||
          e.format.formatCode.toLowerCase().contains(query) ||
          (e.constant?.toLowerCase().contains(query) ?? false) ||
          '${e.numFmtId}' == query;
    }).toList();
  }

  @override
  Widget build(BuildContext context) {
    final entries = _visibleEntries;
    final builtInCount = numFormatCatalog.where((e) => e.numFmtId != null).length;

    return WikiPage(
      icon: Icons.pin,
      title: 'Number Formats Wiki',
      description: 'All $builtInCount built-in Excel number formats (IDs 23-26 are reserved) and '
          'common custom format codes. Every preview is rendered by NumFormat.format(), '
          'and each card has the snippet that applies that format to a cell.',
      accent: _accent,
      isGenerating: widget.isGenerating,
      onGenerate: widget.onGenerate,
      generateLabel: 'Generate Demo Workbook',
      status: widget.status,
      tabs: [
        for (final category in NumFormatCategory.values)
          WikiTab(
            category.name,
            '${category.label} (${numFormatCatalog.where((e) => e.category == category).length})',
          ),
        const WikiTab('code', 'Full Example (Code)'),
      ],
      selectedTab: _tab,
      onTabSelected: (key) => setState(() {
        _tab = key;
        _searchController.clear();
        _searchQuery = '';
      }),
      child: _tab == 'code'
          ? const WikiCodeCard(
              title: 'Number Formats Demo Workbook Code',
              subtitle:
                  'Complete code configuring standard formats, currency, dates, times, and exporting the file',
              code: numberFormatsFullSnippet,
              accent: _accent,
            )
          : Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          WikiSearchField(
            controller: _searchController,
            hint: 'Search all formats by name, code, constant or ID...',
            accent: _accent,
            onChanged: (val) => setState(() => _searchQuery = val),
          ),
          if (_searchQuery.trim().isNotEmpty) ...[
            const SizedBox(height: 8),
            Text(
              '${entries.length} matches across all categories',
              style: TextStyle(fontSize: 11, color: Colors.grey.shade600),
            ),
          ],
          const SizedBox(height: 16),
          WikiGrid(
            itemCount: entries.length,
            extent: 380,
            itemBuilder: (context, index) {
              final entry = entries[index];
              return WikiCard(
                title: entry.name,
                subtitle: entry.constant != null
                    ? 'NumFormat.${entry.constant} · ID ${entry.numFmtId}'
                    : 'NumFormat.custom(...)',
                code: entry.snippet,
                accent: _accent,
                previewLabel: 'Live Format Preview:',
                preview: WikiPreviewBox(
                  alignment: Alignment.topLeft,
                  padding: const EdgeInsets.all(10),
                  child: _FormatPreview(entry: entry),
                ),
              );
            },
          ),
        ],
      ),
    );
  }
}

/// Format code plus one "input → output" row per sample value.
class _FormatPreview extends StatelessWidget {
  final NumFormatEntry entry;

  const _FormatPreview({required this.entry});

  @override
  Widget build(BuildContext context) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Row(
          children: [
            Flexible(
              child: Container(
                padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 3),
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: BorderRadius.circular(4),
                  border: Border.all(color: const Color(0xFFE2E8F0)),
                ),
                child: Text(
                  entry.format.formatCode,
                  maxLines: 1,
                  overflow: TextOverflow.ellipsis,
                  style: const TextStyle(fontFamily: 'monospace', fontSize: 10, color: Color(0xFF334155)),
                ),
              ),
            ),
          ],
        ),
        const SizedBox(height: 10),
        for (final sample in entry.samples)
            Padding(
              padding: const EdgeInsets.only(bottom: 4),
              child: _sampleRow(
                sampleLabel(sample.value),
                entry.display(sample),
                _sectionColor(entry.format.formatCode, sample.value),
              ),
            ),
      ],
    );
  }

  Widget _sampleRow(String input, String output, Color? outputColor) {
    return Row(
      children: [
        Expanded(
          flex: 4,
          child: Text(
            input,
            maxLines: 1,
            overflow: TextOverflow.ellipsis,
            style: const TextStyle(fontFamily: 'monospace', fontSize: 10, color: Color(0xFF64748B)),
          ),
        ),
        const Padding(
          padding: EdgeInsets.symmetric(horizontal: 6),
          child: Icon(Icons.arrow_forward, size: 12, color: Color(0xFF94A3B8)),
        ),
        Expanded(
          flex: 6,
          child: Text(
            output,
            maxLines: 1,
            overflow: TextOverflow.ellipsis,
            textAlign: TextAlign.right,
            style: TextStyle(
              fontSize: 15,
              fontWeight: FontWeight.bold,
              color: outputColor ?? const Color(0xFF0F172A),
            ),
          ),
        ),
      ],
    );
  }

  static const _namedColors = {
    'red': Color(0xFFDC2626),
    'blue': Color(0xFF2563EB),
    'green': Color(0xFF16A34A),
    'magenta': Color(0xFFC026D3),
    'cyan': Color(0xFF0891B2),
    'yellow': Color(0xFFCA8A04),
  };

  /// Color from a `[Red]`-style tag in the format section that applies to
  /// [value] (positive;negative;zero), as Excel would paint it.
  static Color? _sectionColor(String formatCode, CellValue value) {
    final number = switch (value) {
      IntCellValue v => v.value.toDouble(),
      DoubleCellValue v => v.value,
      _ => null,
    };
    if (number == null) return null;
    final sections = formatCode.split(';');
    var section = sections.first;
    if (number < 0 && sections.length > 1) section = sections[1];
    if (number == 0 && sections.length > 2) section = sections[2];
    final match = RegExp(r'\[([A-Za-z]+)\]').firstMatch(section);
    return match == null ? null : _namedColors[match.group(1)!.toLowerCase()];
  }
}
