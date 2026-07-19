import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

class FontsStylesView extends StatefulWidget {
  final bool isGenerating;
  final VoidCallback onGenerate;
  final String status;

  const FontsStylesView({
    super.key,
    required this.isGenerating,
    required this.onGenerate,
    required this.status,
  });

  @override
  State<FontsStylesView> createState() => _FontsStylesViewState();
}

class _FontsStylesViewState extends State<FontsStylesView> {
  final TextEditingController _searchController = TextEditingController();
  String _searchQuery = '';
  String _selectedTab = 'fonts'; // 'fonts' or 'styles'

  // Exactly 48 supported font family items
  final List<Map<String, String>> _fonts = const [
    {'name': 'Arial', 'enum': 'FontFamily.Arial'},
    {'name': 'Arial Narrow', 'enum': 'FontFamily.Arial_Narrow'},
    {'name': 'Arial Rounded MT Bold', 'enum': 'FontFamily.Arial_Rounded_MT_Bold'},
    {'name': 'Arial Unicode MS', 'enum': 'FontFamily.Arial_Unicode_MS'},
    {'name': 'Avenir Book', 'enum': 'FontFamily.Avenir_Book'},
    {'name': 'Avenir Next Regular', 'enum': 'FontFamily.Avenir_Next_Regular'},
    {'name': 'Baskerville', 'enum': 'FontFamily.Baskerville'},
    {'name': 'Baskerville Old Face', 'enum': 'FontFamily.Baskerville_Old_Face'},
    {'name': 'Bauhaus 93', 'enum': 'FontFamily.Bauhaus_93'},
    {'name': 'Bell MT', 'enum': 'FontFamily.Bell_MT'},
    {'name': 'Bernard MT Condensed', 'enum': 'FontFamily.Bernard_MT_Condensed'},
    {'name': 'Book Antiqua', 'enum': 'FontFamily.Book_Antiqua'},
    {'name': 'Bookman Old Style', 'enum': 'FontFamily.Bookman_Old_Style'},
    {'name': 'Bradley Hand', 'enum': 'FontFamily.Bradley_Hand'},
    {'name': 'Britannic Bold', 'enum': 'FontFamily.Britannic_Bold'},
    {'name': 'Brush Script MT', 'enum': 'FontFamily.Brush_Script_MT'},
    {'name': 'Calibri', 'enum': 'FontFamily.Calibri'},
    {'name': 'Calisto MT', 'enum': 'FontFamily.Calisto_MT'},
    {'name': 'Cambria', 'enum': 'FontFamily.Cambria'},
    {'name': 'Candara', 'enum': 'FontFamily.Candara'},
    {'name': 'Century', 'enum': 'FontFamily.Century'},
    {'name': 'Century Gothic', 'enum': 'FontFamily.Century_Gothic'},
    {'name': 'Century Schoolbook', 'enum': 'FontFamily.Century_Schoolbook'},
    {'name': 'Chalkboard', 'enum': 'FontFamily.Chalkboard'},
    {'name': 'Chalkduster', 'enum': 'FontFamily.Chalkduster'},
    {'name': 'Charter', 'enum': 'FontFamily.Charter'},
    {'name': 'Comic Sans MS', 'enum': 'FontFamily.Comic_Sans_MS'},
    {'name': 'Consolas', 'enum': 'FontFamily.Consolas'},
    {'name': 'Constantia', 'enum': 'FontFamily.Constantia'},
    {'name': 'Cooper Black', 'enum': 'FontFamily.Cooper_Black'},
    {'name': 'Copperplate', 'enum': 'FontFamily.Copperplate'},
    {'name': 'Corbel', 'enum': 'FontFamily.Corbel'},
    {'name': 'Courier', 'enum': 'FontFamily.Courier'},
    {'name': 'Courier New', 'enum': 'FontFamily.Courier_New'},
    {'name': 'Dubai', 'enum': 'FontFamily.Dubai'},
    {'name': 'Eurostile', 'enum': 'FontFamily.Eurostile'},
    {'name': 'Futura', 'enum': 'FontFamily.Futura'},
    {'name': 'Geneva', 'enum': 'FontFamily.Geneva'},
    {'name': 'Georgia', 'enum': 'FontFamily.Georgia'},
    {'name': 'Gill Sans', 'enum': 'FontFamily.Gill_Sans'},
    {'name': 'Helvetica', 'enum': 'FontFamily.Helvetica'},
    {'name': 'Helvetica Neue', 'enum': 'FontFamily.Helvetica_Neue'},
    {'name': 'Impact', 'enum': 'FontFamily.Impact'},
    {'name': 'Lucida Bright', 'enum': 'FontFamily.Lucida_Bright'},
    {'name': 'Lucida Console', 'enum': 'FontFamily.Lucida_Console'},
    {'name': 'Lucida Grande', 'enum': 'FontFamily.Lucida_Grande'},
    {'name': 'Lucida Sans', 'enum': 'FontFamily.Lucida_Sans'},
    {'name': 'Monaco', 'enum': 'FontFamily.Monaco'},
  ];

  // Font style showcase items
  final List<Map<String, dynamic>> _styles = const [
    {
      'name': 'Normal Text',
      'detail': 'Default standard style',
      'code': '''
var cellStyle = CellStyle();
cell.cellStyle = cellStyle;
''',
      'style': TextStyle(fontSize: 12),
    },
    {
      'name': 'Bold Font',
      'detail': 'bold: true',
      'code': '''
var cellStyle = CellStyle(
  bold: true,
);
cell.cellStyle = cellStyle;
''',
      'style': TextStyle(fontSize: 12, fontWeight: FontWeight.bold),
    },
    {
      'name': 'Italic Font',
      'detail': 'italic: true',
      'code': '''
var cellStyle = CellStyle(
  italic: true,
);
cell.cellStyle = cellStyle;
''',
      'style': TextStyle(fontSize: 12, fontStyle: FontStyle.italic),
    },
    {
      'name': 'Single Underline',
      'detail': 'underline: Underline.Single',
      'code': '''
var cellStyle = CellStyle(
  underline: Underline.Single,
);
cell.cellStyle = cellStyle;
''',
      'style': TextStyle(fontSize: 12, decoration: TextDecoration.underline),
    },
    {
      'name': 'Double Underline',
      'detail': 'underline: Underline.Double',
      'code': '''
var cellStyle = CellStyle(
  underline: Underline.Double,
);
cell.cellStyle = cellStyle;
''',
      'style': TextStyle(fontSize: 12, decoration: TextDecoration.underline, decorationStyle: TextDecorationStyle.double),
    },
    {
      'name': 'Strikethrough',
      'detail': 'strikethrough: true',
      'code': '''
var cellStyle = CellStyle(
  strikethrough: true,
);
cell.cellStyle = cellStyle;
''',
      'style': TextStyle(fontSize: 12, decoration: TextDecoration.lineThrough),
    },
    {
      'name': 'Combined Style',
      'detail': 'bold: true, italic: true, underline: Underline.Single, fontColorHex: ExcelColor.blue',
      'code': '''
var cellStyle = CellStyle(
  bold: true,
  italic: true,
  underline: Underline.Single,
  fontColorHex: ExcelColor.blue,
);
cell.cellStyle = cellStyle;
''',
      'style': TextStyle(fontSize: 12, fontWeight: FontWeight.bold, fontStyle: FontStyle.italic, decoration: TextDecoration.underline, color: Colors.blue),
    },
    {
      'name': 'Colored & Background Fill',
      'detail': 'fontColorHex: ExcelColor.white, backgroundColorHex: ExcelColor.indigo',
      'code': '''
var cellStyle = CellStyle(
  fontColorHex: ExcelColor.white,
  backgroundColorHex: ExcelColor.indigo,
  bold: true,
);
cell.cellStyle = cellStyle;
''',
      'style': TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: Colors.white),
      'background': Color(0xFF3F51B5),
    },
  ];

  @override
  void dispose() {
    _searchController.dispose();
    super.dispose();
  }

  Widget _buildTabButton(String label, String tabKey) {
    final isActive = _selectedTab == tabKey;
    return Container(
      decoration: BoxDecoration(
        color: isActive ? Colors.indigo : const Color(0xFFF1F5F9),
        borderRadius: BorderRadius.circular(8),
      ),
      child: InkWell(
        onTap: () {
          setState(() {
            _selectedTab = tabKey;
          });
        },
        borderRadius: BorderRadius.circular(8),
        child: Padding(
          padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 10),
          child: Text(
            label,
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

  @override
  Widget build(BuildContext context) {
    final screenWidth = MediaQuery.of(context).size.width;
    final isCompact = screenWidth < 700;

    final filteredFonts = _fonts.where((font) {
      final name = font['name']!.toLowerCase();
      final enumName = font['enum']!.toLowerCase();
      final query = _searchQuery.toLowerCase();
      return name.contains(query) || enumName.contains(query);
    }).toList();

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        // Title & Generate Header
        Card(
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
                      Row(
                        crossAxisAlignment: CrossAxisAlignment.center,
                        children: [
                          Container(
                            padding: const EdgeInsets.all(12),
                            decoration: BoxDecoration(
                              color: Colors.indigo.withOpacity(0.1),
                              borderRadius: BorderRadius.circular(12),
                            ),
                            child: const Icon(
                              Icons.font_download_outlined,
                              color: Colors.indigo,
                              size: 28,
                            ),
                          ),
                          const SizedBox(width: 16),
                          const Expanded(
                            child: Text(
                              'Fonts & Styles Wiki',
                              style: TextStyle(
                                fontSize: 20,
                                fontWeight: FontWeight.bold,
                                color: Color(0xFF0F172A),
                              ),
                            ),
                          ),
                        ],
                      ),
                      const SizedBox(height: 16),
                      Text(
                        'A comprehensive guide of supported fonts and styles. Select and copy code snippets for font families or text decoration styles below.',
                        style: TextStyle(
                          fontSize: 13,
                          color: Colors.grey.shade600,
                        ),
                      ),
                      const SizedBox(height: 20),
                      SizedBox(
                        width: double.infinity,
                        child: ElevatedButton.icon(
                          onPressed: widget.isGenerating ? null : widget.onGenerate,
                          style: ElevatedButton.styleFrom(
                            backgroundColor: Colors.indigo,
                            foregroundColor: Colors.white,
                            padding: const EdgeInsets.symmetric(
                              horizontal: 18,
                              vertical: 14,
                            ),
                            shape: RoundedRectangleBorder(
                              borderRadius: BorderRadius.circular(8),
                            ),
                            elevation: 0,
                          ),
                          icon: widget.isGenerating
                              ? const SizedBox(
                                  width: 16,
                                  height: 16,
                                  child: CircularProgressIndicator(
                                    color: Colors.white,
                                    strokeWidth: 2,
                                  ),
                                )
                              : const Icon(Icons.file_download_outlined, size: 16),
                          label: Text(
                            widget.isGenerating ? 'Generating...' : 'Generate Demo Sheet',
                            style: const TextStyle(
                              fontWeight: FontWeight.bold,
                              fontSize: 13,
                            ),
                          ),
                        ),
                      ),
                    ],
                  )
                : Row(
                    children: [
                      Container(
                        padding: const EdgeInsets.all(12),
                        decoration: BoxDecoration(
                          color: Colors.indigo.withOpacity(0.1),
                          borderRadius: BorderRadius.circular(12),
                        ),
                        child: const Icon(
                          Icons.font_download_outlined,
                          color: Colors.indigo,
                          size: 28,
                        ),
                      ),
                      const SizedBox(width: 16),
                      Expanded(
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            const Text(
                              'Fonts & Styles Wiki',
                              style: TextStyle(
                                fontSize: 20,
                                fontWeight: FontWeight.bold,
                                color: Color(0xFF0F172A),
                              ),
                            ),
                            const SizedBox(height: 4),
                            Text(
                              'A comprehensive guide of supported fonts and styles. Select and copy code snippets for font families or text decoration styles below.',
                              style: TextStyle(
                                fontSize: 13,
                                color: Colors.grey.shade600,
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(width: 16),
                      ElevatedButton.icon(
                        onPressed: widget.isGenerating ? null : widget.onGenerate,
                        style: ElevatedButton.styleFrom(
                          backgroundColor: Colors.indigo,
                          foregroundColor: Colors.white,
                          padding: const EdgeInsets.symmetric(
                            horizontal: 18,
                            vertical: 14,
                          ),
                          shape: RoundedRectangleBorder(
                            borderRadius: BorderRadius.circular(8),
                          ),
                          elevation: 0,
                        ),
                        icon: widget.isGenerating
                            ? const SizedBox(
                                width: 16,
                                height: 16,
                                child: CircularProgressIndicator(
                                  color: Colors.white,
                                  strokeWidth: 2,
                                ),
                              )
                            : const Icon(Icons.file_download_outlined, size: 16),
                        label: Text(
                          widget.isGenerating ? 'Generating...' : 'Generate Demo Sheet',
                          style: const TextStyle(
                            fontWeight: FontWeight.bold,
                            fontSize: 13,
                          ),
                        ),
                      ),
                    ],
                  ),
          ),
        ),
        const SizedBox(height: 12),
        // Generation Status
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
              const Icon(Icons.info_outline, size: 14, color: Colors.indigo),
              const SizedBox(width: 8),
              Expanded(
                child: Text(
                  widget.status,
                  style: const TextStyle(
                    fontSize: 11,
                    color: Color(0xFF334155),
                  ),
                ),
              ),
            ],
          ),
        ),
        const SizedBox(height: 20),
        // Tab Header Selector
        Row(
          children: [
            _buildTabButton('Font Families (48)', 'fonts'),
            const SizedBox(width: 12),
            _buildTabButton('Styles & Decorations', 'styles'),
          ],
        ),
        const SizedBox(height: 20),

        if (_selectedTab == 'fonts') ...[
          // Search Bar
          TextField(
            controller: _searchController,
            onChanged: (val) {
              setState(() {
                _searchQuery = val;
              });
            },
            decoration: InputDecoration(
              hintText: 'Search fonts by name or enum...',
              prefixIcon: const Icon(Icons.search, size: 20),
              suffixIcon: _searchQuery.isNotEmpty
                  ? IconButton(
                      icon: const Icon(Icons.clear, size: 18),
                      onPressed: () {
                        _searchController.clear();
                        setState(() {
                          _searchQuery = '';
                        });
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
                borderSide: const BorderSide(color: Colors.indigo, width: 1.5),
              ),
            ),
          ),
          const SizedBox(height: 16),
          // Font Grid
          GridView.builder(
            shrinkWrap: true,
            physics: const NeverScrollableScrollPhysics(),
            gridDelegate: SliverGridDelegateWithFixedCrossAxisCount(
              crossAxisCount: MediaQuery.of(context).size.width >= 1100 ? 3 : (MediaQuery.of(context).size.width >= 750 ? 2 : 1),
              crossAxisSpacing: 16,
              mainAxisSpacing: 16,
              mainAxisExtent: 220,
            ),
            itemCount: filteredFonts.length,
            itemBuilder: (context, index) {
              final font = filteredFonts[index];
              final fontName = font['name']!;
              final fontEnum = font['enum']!;

              final codeSnippet = '''
var cellStyle = CellStyle(
  fontFamily: getFontFamily($fontEnum), // Mapped to '$fontName'
  fontSize: 12,
  bold: true,
);
cell.cellStyle = cellStyle;
''';

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
                        mainAxisAlignment: MainAxisAlignment.spaceBetween,
                        children: [
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Text(
                                  fontName,
                                  style: const TextStyle(
                                    fontSize: 14,
                                    fontWeight: FontWeight.bold,
                                    color: Color(0xFF0F172A),
                                  ),
                                ),
                                Text(
                                  fontEnum,
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
                            icon: const Icon(Icons.copy, size: 16, color: Colors.indigo),
                            tooltip: 'Copy CellStyle Code',
                            onPressed: () {
                              Clipboard.setData(ClipboardData(text: codeSnippet));
                              ScaffoldMessenger.of(context).showSnackBar(
                                SnackBar(
                                  content: Row(
                                    children: [
                                      const Icon(Icons.check_circle, color: Color(0xFF10B981), size: 16),
                                      const SizedBox(width: 8),
                                      Text('Copied $fontName snippet to clipboard!'),
                                    ],
                                  ),
                                  backgroundColor: const Color(0xFF1E293B),
                                  behavior: SnackBarBehavior.floating,
                                  duration: const Duration(seconds: 1),
                                ),
                              );
                            },
                          ),
                        ],
                      ),
                      const SizedBox(height: 10),
                      const Divider(height: 1, color: Color(0xFFF1F5F9)),
                      const SizedBox(height: 10),
                      const Text(
                        'Live Font Preview:',
                        style: TextStyle(fontSize: 9, fontWeight: FontWeight.bold, color: Color(0xFF64748B)),
                      ),
                      const SizedBox(height: 4),
                      Expanded(
                        child: Container(
                          width: double.infinity,
                          alignment: Alignment.centerLeft,
                          padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                          decoration: BoxDecoration(
                            color: const Color(0xFFF8FAFC),
                            borderRadius: BorderRadius.circular(6),
                          ),
                          child: Text(
                            'The quick brown fox jumps over the lazy dog.',
                            maxLines: 1,
                            overflow: TextOverflow.ellipsis,
                            style: TextStyle(
                              fontFamily: fontName,
                              fontSize: 12,
                              color: const Color(0xFF1E293B),
                            ),
                          ),
                        ),
                      ),
                      const SizedBox(height: 8),
                      Container(
                        width: double.infinity,
                        padding: const EdgeInsets.all(8),
                        decoration: BoxDecoration(
                          color: const Color(0xFF0F172A),
                          borderRadius: BorderRadius.circular(6),
                        ),
                        child: Text(
                          "CellStyle(fontFamily: getFontFamily($fontEnum))",
                          style: const TextStyle(
                            fontFamily: 'monospace',
                            fontSize: 8.5,
                            color: Color(0xFF38BDF8),
                          ),
                        ),
                      ),
                    ],
                  ),
                ),
              );
            },
          ),
        ] else ...[
          // Styles & Decorations Grid
          GridView.builder(
            shrinkWrap: true,
            physics: const NeverScrollableScrollPhysics(),
            gridDelegate: SliverGridDelegateWithFixedCrossAxisCount(
              crossAxisCount: MediaQuery.of(context).size.width >= 1100 ? 3 : (MediaQuery.of(context).size.width >= 750 ? 2 : 1),
              crossAxisSpacing: 16,
              mainAxisSpacing: 16,
              mainAxisExtent: 220,
            ),
            itemCount: _styles.length,
            itemBuilder: (context, index) {
              final item = _styles[index];
              final styleName = item['name'] as String;
              final styleDetail = item['detail'] as String;
              final styleCode = item['code'] as String;
              final textStyle = item['style'] as TextStyle;
              final background = item['background'] as Color?;

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
                        mainAxisAlignment: MainAxisAlignment.spaceBetween,
                        children: [
                          Expanded(
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.start,
                              children: [
                                Text(
                                  styleName,
                                  style: const TextStyle(
                                    fontSize: 14,
                                    fontWeight: FontWeight.bold,
                                    color: Color(0xFF0F172A),
                                  ),
                                ),
                                Text(
                                  styleDetail,
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
                            icon: const Icon(Icons.copy, size: 16, color: Colors.indigo),
                            tooltip: 'Copy Style Code',
                            onPressed: () {
                              Clipboard.setData(ClipboardData(text: styleCode));
                              ScaffoldMessenger.of(context).showSnackBar(
                                SnackBar(
                                  content: Row(
                                    children: [
                                      const Icon(Icons.check_circle, color: Color(0xFF10B981), size: 16),
                                      const SizedBox(width: 8),
                                      Text('Copied $styleName style snippet!'),
                                    ],
                                  ),
                                  backgroundColor: const Color(0xFF1E293B),
                                  behavior: SnackBarBehavior.floating,
                                  duration: const Duration(seconds: 1),
                                ),
                              );
                            },
                          ),
                        ],
                      ),
                      const SizedBox(height: 10),
                      const Divider(height: 1, color: Color(0xFFF1F5F9)),
                      const SizedBox(height: 10),
                      const Text(
                        'Live Style Preview:',
                        style: TextStyle(fontSize: 9, fontWeight: FontWeight.bold, color: Color(0xFF64748B)),
                      ),
                      const SizedBox(height: 4),
                      Expanded(
                        child: Container(
                          width: double.infinity,
                          alignment: Alignment.center,
                          padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                          decoration: BoxDecoration(
                            color: background ?? const Color(0xFFF8FAFC),
                            borderRadius: BorderRadius.circular(6),
                            border: background == null ? null : Border.all(color: const Color(0xFFE2E8F0)),
                          ),
                          child: Text(
                            'Styled Sample Text',
                            maxLines: 1,
                            overflow: TextOverflow.ellipsis,
                            style: textStyle.copyWith(
                              color: background != null ? Colors.white : (textStyle.color ?? const Color(0xFF1E293B)),
                            ),
                          ),
                        ),
                      ),
                      const SizedBox(height: 8),
                      // Short Preview of Code
                      Container(
                        width: double.infinity,
                        padding: const EdgeInsets.all(8),
                        decoration: BoxDecoration(
                          color: const Color(0xFF0F172A),
                          borderRadius: BorderRadius.circular(6),
                        ),
                        child: Text(
                          "${styleCode.split(';')[0].trim()};",
                          style: const TextStyle(
                            fontFamily: 'monospace',
                            fontSize: 9,
                            color: Color(0xFF38BDF8),
                          ),
                        ),
                      ),
                    ],
                  ),
                ),
              );
            },
          ),
        ],
      ],
    );
  }
}
