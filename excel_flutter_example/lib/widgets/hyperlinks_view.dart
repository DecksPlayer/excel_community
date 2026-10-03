import 'package:flutter/material.dart';

import '../data/hyperlink_samples.dart';
import 'wiki/wiki_components.dart';

const _accent = Color(0xFF2563EB);

/// Hyperlinks wiki: each card applies one kind of link, shows the cell as
/// Excel renders it and the links read back from the saved file.
class HyperlinksView extends StatefulWidget {
  final bool isGenerating;
  final VoidCallback onGenerate;
  final String status;

  const HyperlinksView({
    super.key,
    required this.isGenerating,
    required this.onGenerate,
    required this.status,
  });

  @override
  State<HyperlinksView> createState() => _HyperlinksViewState();
}

class _HyperlinksViewState extends State<HyperlinksView> {
  String _tab = 'external';
  late final Map<HyperlinkSample, HyperlinkSampleResult> _results = {
    for (final sample in hyperlinkSamples) sample: runHyperlinkSample(sample),
  };

  static const _tabs = {
    'external': 'External Links',
    'internal': 'Inside the Workbook',
    'manage': 'Read, Style & Remove',
  };

  @override
  Widget build(BuildContext context) {
    final samples = hyperlinkSamples.where((s) => s.category == _tab).toList();
    return WikiPage(
      icon: Icons.link,
      title: 'Cell Hyperlinks',
      description: 'Link cells to web pages, e-mail addresses, files, or other cells and defined names. '
          'Each card runs the snippet, saves the workbook and reads the links back from the file.',
      accent: _accent,
      isGenerating: widget.isGenerating,
      onGenerate: widget.onGenerate,
      generateLabel: 'Generate Demo Workbook',
      status: widget.status,
      tabs: [
        for (final entry in _tabs.entries)
          WikiTab(entry.key,
              '${entry.value} (${hyperlinkSamples.where((s) => s.category == entry.key).length})'),
      ],
      selectedTab: _tab,
      onTabSelected: (tab) => setState(() => _tab = tab),
      child: WikiGrid(
        itemCount: samples.length,
        extent: 340,
        itemBuilder: (context, index) {
          final sample = samples[index];
          return WikiCard(
            title: sample.title,
            subtitle: sample.subtitle,
            code: sample.code,
            accent: _accent,
            previewLabel: 'Live Preview (saved & reopened):',
            preview: WikiPreviewBox(
              alignment: Alignment.topLeft,
              padding: const EdgeInsets.all(10),
              child: _HyperlinkPreview(result: _results[sample]!),
            ),
          );
        },
      ),
    );
  }
}

class _HyperlinkPreview extends StatelessWidget {
  final HyperlinkSampleResult result;

  const _HyperlinkPreview({required this.result});

  @override
  Widget build(BuildContext context) {
    final link = result.link;
    final cellText = Text(
      result.cellText.isEmpty ? '(empty)' : result.cellText,
      maxLines: 1,
      overflow: TextOverflow.ellipsis,
      style: TextStyle(
        fontSize: 13,
        color: result.styled ? const Color(0xFF0563C1) : const Color(0xFF1E293B),
        decoration: result.styled ? TextDecoration.underline : null,
      ),
    );

    return SingleChildScrollView(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          // The cell as Excel shows it.
          Container(
            width: double.infinity,
            padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 6),
            decoration: BoxDecoration(
              color: Colors.white,
              border: Border.all(color: const Color(0xFFCBD5E1)),
            ),
            child: link?.tooltip != null ? Tooltip(message: link!.tooltip!, child: cellText) : cellText,
          ),
          const SizedBox(height: 8),
          if (link != null) ...[
            _line(Icons.public, link.isExternal ? 'url: ${link.url}' : 'internal link'),
            if (link.location != null) _line(Icons.my_location, 'location: ${link.location}'),
            if (link.tooltip != null) _line(Icons.chat_bubble_outline, 'tooltip: ${link.tooltip}'),
          ] else
            _line(Icons.link_off, 'no hyperlink'),
          if (result.note != null) _line(Icons.terminal, result.note!),
          const SizedBox(height: 6),
          Text(
            'Reopened file: ${result.reopened.isEmpty ? 'no links' : result.reopened.keys.join(', ')}',
            style: const TextStyle(fontSize: 10, color: Color(0xFF64748B), fontStyle: FontStyle.italic),
          ),
        ],
      ),
    );
  }

  Widget _line(IconData icon, String text) => Padding(
        padding: const EdgeInsets.only(bottom: 3),
        child: Row(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Icon(icon, size: 12, color: const Color(0xFF64748B)),
            const SizedBox(width: 6),
            Expanded(
              child: Text(
                text,
                style: const TextStyle(fontFamily: 'monospace', fontSize: 10, color: Color(0xFF334155)),
              ),
            ),
          ],
        ),
      );
}
