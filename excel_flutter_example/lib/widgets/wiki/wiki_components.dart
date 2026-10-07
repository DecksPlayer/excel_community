import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

/// Building blocks shared by the interactive "wiki" sections
/// (Fonts & Styles, Number Formats, Page Setup): a header with the generate
/// button, a status bar, tabs, a search field and a responsive grid of cards,
/// each with a live preview and its own code snippet.

class WikiTab {
  final String key;
  final String label;

  const WikiTab(this.key, this.label);
}

/// Copies [code] to the clipboard and confirms it with a snackbar.
void copyWikiSnippet(BuildContext context, String code, String label) {
  Clipboard.setData(ClipboardData(text: code));
  ScaffoldMessenger.of(context).showSnackBar(
    SnackBar(
      content: Row(
        children: [
          const Icon(Icons.check_circle, color: Color(0xFF10B981), size: 16),
          const SizedBox(width: 8),
          Expanded(child: Text('Copied $label snippet to clipboard!')),
        ],
      ),
      backgroundColor: const Color(0xFF1E293B),
      behavior: SnackBarBehavior.floating,
      duration: const Duration(seconds: 1),
    ),
  );
}

/// Page layout of a wiki section: header, generation status, tab selector
/// and the content of the selected tab.
class WikiPage extends StatelessWidget {
  final IconData icon;
  final String title;
  final String description;
  final Color accent;
  final bool isGenerating;
  final VoidCallback onGenerate;
  final String generateLabel;
  final String status;
  final List<WikiTab> tabs;
  final String selectedTab;
  final ValueChanged<String> onTabSelected;
  final Widget child;

  const WikiPage({
    super.key,
    required this.icon,
    required this.title,
    required this.description,
    required this.accent,
    required this.isGenerating,
    required this.onGenerate,
    this.generateLabel = 'Generate Demo Sheet',
    required this.status,
    required this.tabs,
    required this.selectedTab,
    required this.onTabSelected,
    required this.child,
  });

  @override
  Widget build(BuildContext context) {
    final isCompact = MediaQuery.of(context).size.width < 700;

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        _buildHeader(isCompact),
        const SizedBox(height: 12),
        Container(
          width: double.infinity,
          padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 10),
          decoration: BoxDecoration(
            color: const Color(0xFFF1F5F9),
            borderRadius: BorderRadius.circular(10),
            border: Border.all(color: const Color(0xFFE2E8F0)),
          ),
          child: Row(
            children: [
              Icon(Icons.info_outline, size: 14, color: accent),
              const SizedBox(width: 8),
              Expanded(
                child: Text(
                  status,
                  style: const TextStyle(fontSize: 11, color: Color(0xFF334155)),
                ),
              ),
            ],
          ),
        ),
        const SizedBox(height: 20),
        Wrap(
          spacing: 12,
          runSpacing: 12,
          children: [for (final tab in tabs) _buildTabButton(tab)],
        ),
        const SizedBox(height: 20),
        child,
      ],
    );
  }

  Widget _buildHeader(bool isCompact) {
    final iconBox = Container(
      padding: const EdgeInsets.all(12),
      decoration: BoxDecoration(
        color: accent.withValues(alpha: 0.1),
        borderRadius: BorderRadius.circular(12),
      ),
      child: Icon(icon, color: accent, size: 28),
    );
    final titleText = Text(
      title,
      style: const TextStyle(fontSize: 20, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)),
    );
    final descriptionText = Text(
      description,
      style: TextStyle(fontSize: 13, color: Colors.grey.shade600),
    );
    final button = ElevatedButton.icon(
      onPressed: isGenerating ? null : onGenerate,
      style: ElevatedButton.styleFrom(
        backgroundColor: accent,
        foregroundColor: Colors.white,
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 14),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
        elevation: 0,
      ),
      icon: isGenerating
          ? const SizedBox(
              width: 16,
              height: 16,
              child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2),
            )
          : const Icon(Icons.file_download_outlined, size: 16),
      label: Text(
        isGenerating ? 'Generating...' : generateLabel,
        style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 13),
      ),
    );

    return Card(
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
        side: const BorderSide(color: Color(0xFFE2E8F0)),
      ),
      color: Colors.white,
      child: Padding(
        padding: const EdgeInsets.all(20.0),
        child: isCompact
            ? Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Row(children: [iconBox, const SizedBox(width: 16), Expanded(child: titleText)]),
                  const SizedBox(height: 16),
                  descriptionText,
                  const SizedBox(height: 20),
                  SizedBox(width: double.infinity, child: button),
                ],
              )
            : Row(
                children: [
                  iconBox,
                  const SizedBox(width: 16),
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [titleText, const SizedBox(height: 4), descriptionText],
                    ),
                  ),
                  const SizedBox(width: 16),
                  button,
                ],
              ),
      ),
    );
  }

  Widget _buildTabButton(WikiTab tab) {
    final isActive = selectedTab == tab.key;
    return Container(
      decoration: BoxDecoration(
        color: isActive ? accent : const Color(0xFFF1F5F9),
        borderRadius: BorderRadius.circular(8),
      ),
      child: InkWell(
        onTap: () => onTabSelected(tab.key),
        borderRadius: BorderRadius.circular(8),
        child: Padding(
          padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 10),
          child: Text(
            tab.label,
            style: TextStyle(
              fontSize: 12,
              fontWeight: FontWeight.bold,
              color: isActive ? Colors.white : const Color(0xFF475569),
            ),
          ),
        ),
      ),
    );
  }
}

/// Search box with a clear button, used above a [WikiGrid].
class WikiSearchField extends StatelessWidget {
  final TextEditingController controller;
  final String hint;
  final Color accent;
  final ValueChanged<String> onChanged;
  final double? maxWidth;

  const WikiSearchField({
    super.key,
    required this.controller,
    required this.hint,
    required this.accent,
    required this.onChanged,
    this.maxWidth,
  });

  @override
  Widget build(BuildContext context) {
    final field = TextField(
      controller: controller,
      onChanged: onChanged,
      decoration: InputDecoration(
        hintText: hint,
        prefixIcon: const Icon(Icons.search, size: 20),
        suffixIcon: controller.text.isNotEmpty
            ? IconButton(
                icon: const Icon(Icons.clear, size: 18),
                onPressed: () {
                  controller.clear();
                  onChanged('');
                },
              )
            : null,
        filled: true,
        fillColor: Colors.white,
        contentPadding: const EdgeInsets.symmetric(vertical: 14),
        border: OutlineInputBorder(
          borderRadius: BorderRadius.circular(12),
          borderSide: const BorderSide(color: Color(0xFFE2E8F0)),
        ),
        enabledBorder: OutlineInputBorder(
          borderRadius: BorderRadius.circular(12),
          borderSide: const BorderSide(color: Color(0xFFE2E8F0)),
        ),
        focusedBorder: OutlineInputBorder(
          borderRadius: BorderRadius.circular(12),
          borderSide: BorderSide(color: accent, width: 1.5),
        ),
      ),
    );
    if (maxWidth == null) return field;
    return ConstrainedBox(constraints: BoxConstraints(maxWidth: maxWidth!), child: field);
  }
}

/// Responsive grid of fixed-height cards: 3 columns on wide screens, 2 on
/// medium and 1 on phones.
class WikiGrid extends StatelessWidget {
  final int itemCount;
  final double extent;
  final IndexedWidgetBuilder itemBuilder;

  const WikiGrid({
    super.key,
    required this.itemCount,
    required this.extent,
    required this.itemBuilder,
  });

  @override
  Widget build(BuildContext context) {
    final width = MediaQuery.of(context).size.width;
    return GridView.builder(
      shrinkWrap: true,
      physics: const NeverScrollableScrollPhysics(),
      gridDelegate: SliverGridDelegateWithFixedCrossAxisCount(
        crossAxisCount: width >= 1100 ? 3 : (width >= 750 ? 2 : 1),
        crossAxisSpacing: 16,
        mainAxisSpacing: 16,
        mainAxisExtent: extent,
      ),
      itemCount: itemCount,
      itemBuilder: itemBuilder,
    );
  }
}

/// A wiki entry: title, subtitle, copy button, live preview and the code
/// snippet for that specific type.
///
/// [code] is what the copy button copies; [codeSummary] (defaults to [code])
/// is what the dark code box shows.
class WikiCard extends StatelessWidget {
  final String title;
  final String subtitle;
  final String code;
  final String? codeSummary;
  final String previewLabel;
  final Widget preview;
  final Color accent;

  const WikiCard({
    super.key,
    required this.title,
    required this.subtitle,
    required this.code,
    this.codeSummary,
    required this.previewLabel,
    required this.preview,
    required this.accent,
  });

  @override
  Widget build(BuildContext context) {
    return Card(
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(12),
        side: const BorderSide(color: Color(0xFFE2E8F0)),
      ),
      color: Colors.white,
      child: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        title,
                        maxLines: 1,
                        overflow: TextOverflow.ellipsis,
                        style: const TextStyle(
                          fontSize: 14,
                          fontWeight: FontWeight.bold,
                          color: Color(0xFF0F172A),
                        ),
                      ),
                      Text(
                        subtitle,
                        maxLines: 1,
                        overflow: TextOverflow.ellipsis,
                        style: TextStyle(
                          fontSize: 10,
                          color: Colors.grey.shade500,
                          fontFamily: 'monospace',
                        ),
                      ),
                    ],
                  ),
                ),
                IconButton(
                  icon: Icon(Icons.copy, size: 16, color: accent),
                  tooltip: 'Copy Code',
                  onPressed: () => copyWikiSnippet(context, code, title),
                ),
              ],
            ),
            const SizedBox(height: 10),
            const Divider(height: 1, color: Color(0xFFF1F5F9)),
            const SizedBox(height: 10),
            Text(
              previewLabel,
              style: const TextStyle(fontSize: 9, fontWeight: FontWeight.bold, color: Color(0xFF64748B)),
            ),
            const SizedBox(height: 4),
            Expanded(child: SizedBox(width: double.infinity, child: preview)),
            const SizedBox(height: 8),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(8),
              decoration: BoxDecoration(
                color: const Color(0xFF0F172A),
                borderRadius: BorderRadius.circular(6),
              ),
              // Lines never wrap, so the box height only depends on the
              // number of lines; long lines scroll horizontally.
              child: SingleChildScrollView(
                scrollDirection: Axis.horizontal,
                child: Text(
                  codeSummary ?? code,
                  softWrap: false,
                  style: const TextStyle(
                    fontFamily: 'monospace',
                    fontSize: 8.5,
                    height: 1.35,
                    color: Color(0xFF38BDF8),
                  ),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}

/// Rounded background for the preview area of a [WikiCard].
class WikiPreviewBox extends StatelessWidget {
  final Widget child;
  final Color color;
  final AlignmentGeometry alignment;
  final EdgeInsetsGeometry padding;
  final BoxBorder? border;

  const WikiPreviewBox({
    super.key,
    required this.child,
    this.color = const Color(0xFFF8FAFC),
    this.alignment = Alignment.center,
    this.padding = const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
    this.border,
  });

  @override
  Widget build(BuildContext context) {
    return Container(
      alignment: alignment,
      padding: padding,
      decoration: BoxDecoration(
        color: color,
        borderRadius: BorderRadius.circular(6),
        border: border,
      ),
      child: child,
    );
  }
}

/// Full example code card shown in the "Full Example" tab of wiki pages.
/// Displays the complete source code used to produce the downloadable .xlsx file,
/// with a copy button and a formatted dark code box.
class WikiCodeCard extends StatelessWidget {
  final String title;
  final String subtitle;
  final String code;
  final Color accent;

  const WikiCodeCard({
    super.key,
    this.title = 'Complete Workbook Code',
    this.subtitle = 'Self-contained Dart code that generates the downloadable .xlsx demo file',
    required this.code,
    required this.accent,
  });

  @override
  Widget build(BuildContext context) {
    return Card(
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(12),
        side: const BorderSide(color: Color(0xFFE2E8F0)),
      ),
      color: Colors.white,
      child: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        title,
                        style: const TextStyle(
                          fontSize: 15,
                          fontWeight: FontWeight.bold,
                          color: Color(0xFF0F172A),
                        ),
                      ),
                      const SizedBox(height: 2),
                      Text(
                        subtitle,
                        style: TextStyle(
                          fontSize: 12,
                          color: Colors.grey.shade600,
                        ),
                      ),
                    ],
                  ),
                ),
                ElevatedButton.icon(
                  onPressed: () => copyWikiSnippet(context, code, title),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: accent,
                    foregroundColor: Colors.white,
                    padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
                    shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                    elevation: 0,
                  ),
                  icon: const Icon(Icons.copy, size: 14),
                  label: const Text(
                    'Copy Code',
                    style: TextStyle(fontSize: 12, fontWeight: FontWeight.bold),
                  ),
                ),
              ],
            ),
            const SizedBox(height: 12),
            const Divider(height: 1, color: Color(0xFFF1F5F9)),
            const SizedBox(height: 12),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(14),
              decoration: BoxDecoration(
                color: const Color(0xFF0F172A),
                borderRadius: BorderRadius.circular(8),
              ),
              child: SingleChildScrollView(
                scrollDirection: Axis.horizontal,
                child: SelectableText(
                  code,
                  style: const TextStyle(
                    fontFamily: 'monospace',
                    fontSize: 12,
                    height: 1.45,
                    color: Color(0xFFF1F5F9),
                  ),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
