import 'package:excel_community/excel_community.dart'
    show
        CellValue,
        DataValidation,
        DataValidationErrorStyle,
        DataValidationType;
import 'package:flutter/material.dart';

import '../data/data_validation_samples.dart';
import '../data/number_format_catalog.dart' show sampleLabel;
import '../data/snippets/data_validation.dart';
import 'wiki/wiki_components.dart';

const _accent = Color(0xFF0D9488);

/// Data validation wiki: each card shows the cell as Excel renders the
/// rule (dropdown, input message) and checks sample values with
/// `DataValidation.accepts()`.
class DataValidationView extends StatefulWidget {
  final bool isGenerating;
  final VoidCallback onGenerate;
  final String status;

  const DataValidationView({
    super.key,
    required this.isGenerating,
    required this.onGenerate,
    required this.status,
  });

  @override
  State<DataValidationView> createState() => _DataValidationViewState();
}

class _DataValidationViewState extends State<DataValidationView> {
  String _tab = 'lists';

  static const _tabs = {
    'lists': 'Dropdown Lists',
    'numbers': 'Numbers, Dates & Times',
    'text': 'Text, Custom & Checks',
  };

  @override
  Widget build(BuildContext context) {
    final samples = validationSamples.where((s) => s.category == _tab).toList();
    return WikiPage(
      icon: Icons.fact_check_outlined,
      title: 'Data Validation & Dropdowns',
      description:
          'Restrict what can be typed in a range: dropdown lists, number, date, time and text limits, '
          'or custom formulas, with input and error messages. Sample values are checked live with accepts().',
      accent: _accent,
      isGenerating: widget.isGenerating,
      onGenerate: widget.onGenerate,
      generateLabel: 'Generate Demo Workbook',
      status: widget.status,
      tabs: [
        for (final entry in _tabs.entries)
          WikiTab(
            entry.key,
            '${entry.value} (${validationSamples.where((s) => s.category == entry.key).length})',
          ),
        const WikiTab('code', 'Full Example (Code)'),
      ],
      selectedTab: _tab,
      onTabSelected: (tab) => setState(() => _tab = tab),
      child: _tab == 'code'
          ? const WikiCodeCard(
              title: 'Data Validation Demo Workbook Code',
              subtitle:
                  'Complete code that configures all 11 validated columns, dynamic lookups, sample rows, and exports the file',
              code: dataValidationSnippet,
              accent: _accent,
            )
          : WikiGrid(
              itemCount: samples.length,
              extent: 440,
              itemBuilder: (context, index) {
                final sample = samples[index];
                return WikiCard(
                  title: sample.title,
                  subtitle: sample.subtitle,
                  code: sample.code,
                  accent: _accent,
                  previewLabel: 'Live Preview:',
                  preview: WikiPreviewBox(
                    alignment: Alignment.topLeft,
                    padding: const EdgeInsets.all(10),
                    child: _ValidationPreview(sample: sample),
                  ),
                );
              },
            ),
    );
  }
}

class _ValidationPreview extends StatelessWidget {
  final ValidationSample sample;

  const _ValidationPreview({required this.sample});

  @override
  Widget build(BuildContext context) {
    final rule = sample.rule;
    final items = rule.listItems;
    final isList = rule.type == DataValidationType.list;

    return SingleChildScrollView(
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              // The selected cell, with the dropdown arrow for lists.
              Container(
                width: 120,
                height: 26,
                decoration: BoxDecoration(
                  color: Colors.white,
                  border: Border.all(color: const Color(0xFF16A34A), width: 2),
                ),
                child: Row(
                  children: [
                    const Spacer(),
                    if (isList && rule.showDropdown)
                      Container(
                        width: 18,
                        color: const Color(0xFFF1F5F9),
                        child: const Icon(
                          Icons.arrow_drop_down,
                          size: 16,
                          color: Color(0xFF334155),
                        ),
                      ),
                  ],
                ),
              ),
              const SizedBox(width: 8),
              if (rule.prompt != null) Expanded(child: _PromptBox(rule: rule)),
            ],
          ),
          if (isList) ...[
            const SizedBox(height: 4),
            Container(
              width: 120,
              padding: const EdgeInsets.symmetric(vertical: 2),
              decoration: BoxDecoration(
                color: Colors.white,
                border: Border.all(color: const Color(0xFFCBD5E1)),
                boxShadow: const [
                  BoxShadow(
                    color: Colors.black12,
                    blurRadius: 3,
                    offset: Offset(1, 2),
                  ),
                ],
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  for (final item in items ?? ['= ${rule.formula1}'])
                    Padding(
                      padding: const EdgeInsets.symmetric(
                        horizontal: 6,
                        vertical: 2,
                      ),
                      child: Text(
                        item,
                        maxLines: 1,
                        overflow: TextOverflow.ellipsis,
                        style: const TextStyle(
                          fontSize: 10,
                          color: Color(0xFF1E293B),
                        ),
                      ),
                    ),
                ],
              ),
            ),
          ],
          const SizedBox(height: 10),
          const Text(
            'accepts():',
            style: TextStyle(
              fontSize: 9,
              fontWeight: FontWeight.bold,
              color: Color(0xFF64748B),
            ),
          ),
          const SizedBox(height: 4),
          for (final value in sample.tries) _tryRow(rule, value),
          if (rule.error != null) ...[
            const SizedBox(height: 6),
            Text(
              '${_styleLabel(rule.errorStyle)} "${rule.errorTitle}": ${rule.error}',
              style: const TextStyle(
                fontSize: 10,
                color: Color(0xFF64748B),
                fontStyle: FontStyle.italic,
              ),
            ),
          ],
        ],
      ),
    );
  }

  Widget _tryRow(DataValidation rule, CellValue? value) {
    final result = rule.accepts(value);
    final (icon, color, label) = switch (result) {
      true => (Icons.check_circle, const Color(0xFF16A34A), 'accepted'),
      false => (Icons.cancel, const Color(0xFFDC2626), 'rejected'),
      null => (
        Icons.help_outline,
        const Color(0xFF94A3B8),
        'evaluated by Excel',
      ),
    };
    return Padding(
      padding: const EdgeInsets.only(bottom: 3),
      child: Row(
        children: [
          Icon(icon, size: 13, color: color),
          const SizedBox(width: 6),
          Expanded(
            child: Text(
              value == null ? '(blank)' : sampleLabel(value),
              maxLines: 1,
              overflow: TextOverflow.ellipsis,
              style: const TextStyle(
                fontFamily: 'monospace',
                fontSize: 10,
                color: Color(0xFF334155),
              ),
            ),
          ),
          Text(
            label,
            style: TextStyle(
              fontSize: 10,
              color: color,
              fontWeight: FontWeight.w600,
            ),
          ),
        ],
      ),
    );
  }

  static String _styleLabel(DataValidationErrorStyle style) => switch (style) {
    DataValidationErrorStyle.stop => '⛔ Stop',
    DataValidationErrorStyle.warning => '⚠️ Warning',
    DataValidationErrorStyle.information => 'ℹ️ Information',
  };
}

/// Excel-like input message shown next to the selected cell.
class _PromptBox extends StatelessWidget {
  final DataValidation rule;

  const _PromptBox({required this.rule});

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.all(6),
      decoration: BoxDecoration(
        color: const Color(0xFFFFFBEB),
        border: Border.all(color: const Color(0xFFFCD34D)),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          if (rule.promptTitle != null)
            Text(
              rule.promptTitle!,
              style: const TextStyle(
                fontSize: 10,
                fontWeight: FontWeight.bold,
                color: Color(0xFF1E293B),
              ),
            ),
          Text(
            rule.prompt!,
            style: const TextStyle(fontSize: 10, color: Color(0xFF334155)),
          ),
        ],
      ),
    );
  }
}
