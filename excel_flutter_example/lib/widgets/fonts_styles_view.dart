import 'package:flutter/material.dart';

import 'wiki/wiki_components.dart';

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

  @override
  Widget build(BuildContext context) {
    final query = _searchQuery.toLowerCase();
    final filteredFonts = _fonts.where((font) {
      return font['name']!.toLowerCase().contains(query) ||
          font['enum']!.toLowerCase().contains(query);
    }).toList();

    return WikiPage(
      icon: Icons.font_download_outlined,
      title: 'Fonts & Styles Wiki',
      description:
          'A comprehensive guide of supported fonts and styles. Select and copy code snippets for font families or text decoration styles below.',
      accent: Colors.indigo,
      isGenerating: widget.isGenerating,
      onGenerate: widget.onGenerate,
      status: widget.status,
      tabs: const [
        WikiTab('fonts', 'Font Families (48)'),
        WikiTab('styles', 'Styles & Decorations'),
      ],
      selectedTab: _selectedTab,
      onTabSelected: (tab) => setState(() => _selectedTab = tab),
      child: _selectedTab == 'fonts'
          ? Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                WikiSearchField(
                  controller: _searchController,
                  hint: 'Search fonts by name or enum...',
                  accent: Colors.indigo,
                  onChanged: (val) => setState(() => _searchQuery = val),
                ),
                const SizedBox(height: 16),
                WikiGrid(
                  itemCount: filteredFonts.length,
                  extent: 220,
                  itemBuilder: (context, index) {
                    final fontName = filteredFonts[index]['name']!;
                    final fontEnum = filteredFonts[index]['enum']!;
                    return WikiCard(
                      title: fontName,
                      subtitle: fontEnum,
                      accent: Colors.indigo,
                      code: '''
var cellStyle = CellStyle(
  fontFamily: getFontFamily($fontEnum), // Mapped to '$fontName'
  fontSize: 12,
  bold: true,
);
cell.cellStyle = cellStyle;
''',
                      codeSummary: 'CellStyle(fontFamily: getFontFamily($fontEnum))',
                      previewLabel: 'Live Font Preview:',
                      preview: WikiPreviewBox(
                        alignment: Alignment.centerLeft,
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
                    );
                  },
                ),
              ],
            )
          : WikiGrid(
              itemCount: _styles.length,
              extent: 220,
              itemBuilder: (context, index) {
                final item = _styles[index];
                final styleName = item['name'] as String;
                final styleCode = item['code'] as String;
                final textStyle = item['style'] as TextStyle;
                final background = item['background'] as Color?;
                return WikiCard(
                  title: styleName,
                  subtitle: item['detail'] as String,
                  accent: Colors.indigo,
                  code: styleCode,
                  codeSummary: '${styleCode.split(';')[0].trim()};',
                  previewLabel: 'Live Style Preview:',
                  preview: WikiPreviewBox(
                    color: background ?? const Color(0xFFF8FAFC),
                    border: background == null ? null : Border.all(color: const Color(0xFFE2E8F0)),
                    child: Text(
                      'Styled Sample Text',
                      maxLines: 1,
                      overflow: TextOverflow.ellipsis,
                      style: textStyle.copyWith(
                        color: background != null ? Colors.white : (textStyle.color ?? const Color(0xFF1E293B)),
                      ),
                    ),
                  ),
                );
              },
            ),
    );
  }
}
