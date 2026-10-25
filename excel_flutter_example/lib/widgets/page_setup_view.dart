import 'dart:math' as math;

import 'package:excel_community/excel_community.dart' show PaperSize, PageMargins;
import 'package:flutter/material.dart';

import '../data/snippets/page_setup.dart';
import 'wiki/wiki_components.dart';


const _accent = Color(0xFF0EA5E9);

/// Pixels per millimetre used in the "Paper Sizes" tab so every sheet of
/// paper is drawn at the same scale and sizes can be compared visually.
const double _paperScalePpm = 0.26;

/// Interactive wiki for page setup & print configuration: paper sizes drawn
/// to scale, orientation & scaling, margins and print options, each with a
/// live page preview and a copyable code snippet.
class PageSetupView extends StatefulWidget {
  final bool isGenerating;
  final VoidCallback onGenerate;
  final String status;

  const PageSetupView({
    super.key,
    required this.isGenerating,
    required this.onGenerate,
    required this.status,
  });

  @override
  State<PageSetupView> createState() => _PageSetupViewState();
}

class _PageSetupViewState extends State<PageSetupView> {
  final TextEditingController _searchController = TextEditingController();
  String _searchQuery = '';
  String _selectedTab = 'paper'; // 'paper', 'scaling', 'margins', 'print'
  bool _landscape = false;

  /// Dart identifiers of the [PaperSize] constants, used in code snippets.
  static const Map<int, String> _paperIdentifiers = {
    1: 'letter',
    3: 'tabloid',
    4: 'ledger',
    5: 'legal',
    6: 'statement',
    7: 'executive',
    8: 'a3',
    9: 'a4',
    11: 'a5',
    12: 'b4',
    13: 'b5',
    14: 'folio',
    15: 'quarto',
    20: 'envelope10',
    27: 'envelopeDL',
    28: 'envelopeC5',
    66: 'a2',
    70: 'a6',
  };

  static const List<_PageExample> _scalingExamples = [
    _PageExample(
      name: 'Portrait',
      detail: 'PageOrientation.portrait',
      code: r"""final sheet = excel['Sheet1'];
// Taller than wide (the usual default)
sheet.setPageOrientation(PageOrientation.portrait);""",
      look: _PageLook(columns: 8),
    ),
    _PageExample(
      name: 'Landscape',
      detail: 'PageOrientation.landscape',
      code: r"""final sheet = excel['Sheet1'];
// Wider than tall: more columns fit per page
sheet.setPageOrientation(PageOrientation.landscape);""",
      look: _PageLook(landscape: true, columns: 8),
    ),
    _PageExample(
      name: 'No Scaling (100 %)',
      detail: 'Wide table overflows to a 2nd page',
      code: r"""final sheet = excel['Sheet1'];
// Print at real size: wide tables
// continue on another page
sheet.setPrintScale(100);""",
      look: _PageLook(columns: 14),
    ),
    _PageExample(
      name: 'Fit All Columns on One Page',
      detail: 'fitToPages(width: 1, height: 0)',
      code: r"""final sheet = excel['Sheet1'];
// Every column on 1 page wide, as many
// pages tall as needed (0 = automatic)
sheet.fitToPages(width: 1, height: 0);""",
      look: _PageLook(columns: 14, fitWidth: true),
    ),
    _PageExample(
      name: 'Fit Sheet on One Page',
      detail: 'fitToPages(width: 1, height: 1)',
      code: r"""final sheet = excel['Sheet1'];
// Shrink the whole sheet onto a single page
sheet.fitToPages(width: 1, height: 1);""",
      look: _PageLook(columns: 14, rows: 70, fitWidth: true, fitHeight: true),
    ),
    _PageExample(
      name: 'Scale 50 %',
      detail: 'setPrintScale(50)',
      code: r"""final sheet = excel['Sheet1'];
// Shrink to 50 % (allowed range: 10-400)
sheet.setPrintScale(50);""",
      look: _PageLook(scale: 0.5),
    ),
    _PageExample(
      name: 'Scale 150 %',
      detail: 'setPrintScale(150)',
      code: r"""final sheet = excel['Sheet1'];
// Enlarge to 150 % (allowed range: 10-400)
sheet.setPrintScale(150);""",
      look: _PageLook(scale: 1.5),
    ),
  ];

  static const List<_PageExample> _marginExamples = [
    _PageExample(
      name: 'Normal',
      detail: 'L/R 0.7"  T/B 0.75"  H/F 0.3"',
      code: r"""final sheet = excel['Sheet1'];
// Excel's default: left/right 0.7", top/bottom 0.75"
sheet.pageMargins = PageMargins.normal;""",
      look: _PageLook(margins: PageMargins.normal, columns: 5, headerFooter: true),
    ),
    _PageExample(
      name: 'Wide',
      detail: 'L/R/T/B 1"  H/F 0.5"',
      code: r"""final sheet = excel['Sheet1'];
// 1" on every side, header/footer at 0.5"
sheet.pageMargins = PageMargins.wide;""",
      look: _PageLook(margins: PageMargins.wide, columns: 5, headerFooter: true),
    ),
    _PageExample(
      name: 'Narrow',
      detail: 'L/R 0.25"  T/B 0.75"  H/F 0.3"',
      code: r"""final sheet = excel['Sheet1'];
// Left/right 0.25": more room for columns
sheet.pageMargins = PageMargins.narrow;""",
      look: _PageLook(margins: PageMargins.narrow, columns: 5, headerFooter: true),
    ),
    _PageExample(
      name: 'Same on All Sides',
      detail: 'PageMargins.all(0.5)',
      code: r"""final sheet = excel['Sheet1'];
// 0.5" on all four sides
sheet.pageMargins = const PageMargins.all(0.5);""",
      look: _PageLook(margins: PageMargins.all(0.5), columns: 5, headerFooter: true),
    ),
    _PageExample(
      name: 'Custom (inches)',
      detail: 'left: 1.5, right: 0.5, top: 1.2',
      code: r"""final sheet = excel['Sheet1'];
// Values in inches; sides you omit keep
// the Normal defaults
sheet.pageMargins = const PageMargins(
  left: 1.5,
  right: 0.5,
  top: 1.2,
);""",
      look: _PageLook(
        margins: PageMargins(left: 1.5, right: 0.5, top: 1.2),
        columns: 5,
        headerFooter: true,
      ),
    ),
    _PageExample(
      name: 'Centimetres',
      detail: 'fromCentimeters(left: 3, right: 3, ...)',
      code: r"""final sheet = excel['Sheet1'];
// Converted to inches when the file is saved
sheet.pageMargins = PageMargins.fromCentimeters(
  left: 3,
  right: 3,
  top: 2.5,
  bottom: 2.5,
);""",
      look: _PageLook(
        // 3 cm ≈ 1.18", 2.5 cm ≈ 0.98"
        margins: PageMargins(left: 1.18, right: 1.18, top: 0.98, bottom: 0.98),
        columns: 5,
        headerFooter: true,
      ),
    ),
  ];

  static const List<_PageExample> _printExamples = [
    _PageExample(
      name: 'Plain (Default)',
      detail: 'No gridlines, no headings',
      code: r"""final sheet = excel['Sheet1'];
// Remove gridlines, headings and centering
sheet.clearPrintOptions();""",
      look: _PageLook(columns: 5),
    ),
    _PageExample(
      name: 'Print Gridlines',
      detail: 'printOptions gridLines="1"',
      code: r"""final sheet = excel['Sheet1'];
// Print the cell borders you see on screen
sheet.setPrintGridLines(true);""",
      look: _PageLook(columns: 5, gridLines: true),
    ),
    _PageExample(
      name: 'Print Row & Column Headings',
      detail: 'printOptions headings="1"',
      code: r"""final sheet = excel['Sheet1'];
// Print A, B, C... and 1, 2, 3...
sheet.setPrintHeadings(true);
sheet.setPrintGridLines(true);""",
      look: _PageLook(columns: 5, headings: true, gridLines: true),
    ),
    _PageExample(
      name: 'Center Horizontally',
      detail: 'horizontalCentered="1"',
      code: r"""final sheet = excel['Sheet1'];
// Center the content between left/right margins
sheet.setPrintCentered(horizontally: true);""",
      look: _PageLook(columns: 4, rows: 14, centerH: true),
    ),
    _PageExample(
      name: 'Center on Page',
      detail: 'horizontally + vertically',
      code: r"""final sheet = excel['Sheet1'];
// Center the content in the middle of the page
sheet.setPrintCentered(
  horizontally: true,
  vertically: true,
);""",
      look: _PageLook(columns: 4, rows: 14, centerH: true, centerV: true),
    ),
    _PageExample(
      name: 'Black & White',
      detail: 'PageSetup(blackAndWhite: true)',
      code: r"""final sheet = excel['Sheet1'];
// copyWith keeps the rest of the page setup
sheet.pageSetup = (sheet.pageSetup ?? const PageSetup())
    .copyWith(blackAndWhite: true);""",
      look: _PageLook(columns: 5, blackAndWhite: true),
    ),
    _PageExample(
      name: 'Start Numbering at Page 10',
      detail: 'firstPageNumber: 10',
      code: r"""final sheet = excel['Sheet1'];
sheet.pageSetup = (sheet.pageSetup ?? const PageSetup())
    .copyWith(firstPageNumber: 10, useFirstPageNumber: true);
// Show the number in the footer: "Page 10"
sheet.headerFooter = HeaderFooter(oddFooter: '&CPage &P');""",
      look: _PageLook(columns: 5, footerLabel: 'Page 10'),
    ),
  ];

  @override
  void dispose() {
    _searchController.dispose();
    super.dispose();
  }

  Widget _buildOrientationToggle() {
    Widget option(String label, IconData icon, bool landscape) {
      final isActive = _landscape == landscape;
      return InkWell(
        onTap: () => setState(() => _landscape = landscape),
        borderRadius: BorderRadius.circular(8),
        child: Container(
          padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
          decoration: BoxDecoration(
            color: isActive ? _accent.withValues(alpha: 0.12) : Colors.white,
            borderRadius: BorderRadius.circular(8),
            border: Border.all(color: isActive ? _accent : const Color(0xFFE2E8F0)),
          ),
          child: Row(
            mainAxisSize: MainAxisSize.min,
            children: [
              Icon(icon, size: 16, color: isActive ? _accent : const Color(0xFF64748B)),
              const SizedBox(width: 6),
              Text(
                label,
                style: TextStyle(
                  fontSize: 12,
                  fontWeight: FontWeight.bold,
                  color: isActive ? _accent : const Color(0xFF475569),
                ),
              ),
            ],
          ),
        ),
      );
    }

    return Row(
      mainAxisSize: MainAxisSize.min,
      children: [
        option('Portrait', Icons.crop_portrait, false),
        const SizedBox(width: 8),
        option('Landscape', Icons.crop_landscape, true),
      ],
    );
  }

  Widget _buildPaperTab() {
    final query = _searchQuery.toLowerCase();
    final papers = PaperSize.values.where((p) {
      final id = _paperIdentifiers[p.code] ?? '';
      return p.name.toLowerCase().contains(query) ||
          id.toLowerCase().contains(query) ||
          '${p.code}' == query;
    }).toList();

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Wrap(
          spacing: 12,
          runSpacing: 12,
          crossAxisAlignment: WrapCrossAlignment.center,
          children: [
            WikiSearchField(
              controller: _searchController,
              hint: 'Search paper sizes by name, constant or code...',
              accent: _accent,
              maxWidth: 360,
              onChanged: (val) => setState(() => _searchQuery = val),
            ),
            _buildOrientationToggle(),
          ],
        ),
        const SizedBox(height: 8),
        Text(
          'All pages are drawn at the same scale, so you can compare their real sizes.',
          style: TextStyle(fontSize: 11, color: Colors.grey.shade600),
        ),
        const SizedBox(height: 16),
        WikiGrid(
          itemCount: papers.length,
          extent: 420,
          itemBuilder: (context, index) {
            final paper = papers[index];
            final id = _paperIdentifiers[paper.code] ?? 'fromCode(${paper.code})';
            final orientation = _landscape ? 'landscape' : 'portrait';
            final code = "final sheet = excel['Sheet1'];\n"
                '// ${paper.name}: ${_fmt(paper.widthMm!)} × ${_fmt(paper.heightMm!)} mm\n'
                'sheet.setPaperSize(PaperSize.$id);\n'
                'sheet.setPageOrientation(PageOrientation.$orientation);';
            final w = paper.widthMm ?? 210;
            final h = paper.heightMm ?? 297;
            final inches = '${_fmt(w / 25.4)} × ${_fmt(h / 25.4)} in';
            return WikiCard(
              title: paper.name,
              subtitle: 'PaperSize.$id · code ${paper.code}',
              code: code,
              accent: _accent,
              previewLabel: 'Live Page Preview:',
              preview: _previewBox(Stack(
                children: [
                  Align(
                    alignment: Alignment.bottomCenter,
                    child: _PagePreview(
                      widthMm: w,
                      heightMm: h,
                      look: _PageLook(landscape: _landscape, columns: 5, rows: 40),
                      pixelsPerMm: _paperScalePpm,
                    ),
                  ),
                  Positioned(
                    left: 0,
                    top: 0,
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(
                          '${_fmt(w)} × ${_fmt(h)} mm',
                          style: const TextStyle(
                            fontSize: 10,
                            fontWeight: FontWeight.bold,
                            color: Color(0xFF1E293B),
                          ),
                        ),
                        Text(inches, style: const TextStyle(fontSize: 9, color: Color(0xFF64748B))),
                      ],
                    ),
                  ),
                ],
              )),
            );
          },
        ),
      ],
    );
  }

  Widget _buildExampleTab(List<_PageExample> examples) {
    return WikiGrid(
      itemCount: examples.length,
      extent: 420,
      itemBuilder: (context, index) {
        final example = examples[index];
        return WikiCard(
          title: example.name,
          subtitle: example.detail,
          code: example.code,
          accent: _accent,
          previewLabel: 'Live Page Preview:',
          preview: _previewBox(Center(
            child: _PagePreview(
              widthMm: PaperSize.a4.widthMm!,
              heightMm: PaperSize.a4.heightMm!,
              look: example.look,
            ),
          )),
        );
      },
    );
  }

  Widget _previewBox(Widget child) => WikiPreviewBox(
        color: const Color(0xFFF1F5F9),
        padding: const EdgeInsets.all(8),
        child: child,
      );

  @override
  Widget build(BuildContext context) {
    return WikiPage(
      icon: Icons.print_outlined,
      title: 'Page Setup & Print Wiki',
      description: 'Every paper size drawn to scale, plus orientation, scaling, margins and print options. '
          'Each card shows how the sheet prints and the code that configures it.',
      accent: _accent,
      isGenerating: widget.isGenerating,
      onGenerate: widget.onGenerate,
      generateLabel: 'Generate Demo Workbook',
      status: widget.status,
      tabs: [
        WikiTab('paper', 'Paper Sizes (${PaperSize.values.length})'),
        const WikiTab('scaling', 'Orientation & Scaling'),
        const WikiTab('margins', 'Margins'),
        const WikiTab('print', 'Print Options'),
        const WikiTab('code', 'Full Example (Code)'),
      ],
      selectedTab: _selectedTab,
      onTabSelected: (tab) => setState(() => _selectedTab = tab),
      child: switch (_selectedTab) {
        'scaling' => _buildExampleTab(_scalingExamples),
        'margins' => _buildExampleTab(_marginExamples),
        'print' => _buildExampleTab(_printExamples),
        'code' => const WikiCodeCard(
            title: 'Page Setup & Print Demo Workbook Code',
            subtitle:
                'Complete code configuring paper sizes, orientation, scaling, fit-to-pages, margins, and headers/footers',
            code: pageSetupSnippet,
            accent: _accent,
          ),
        _ => _buildPaperTab(),
      },
    );
  }
}

String _fmt(double v) {
  final s = v.toStringAsFixed(1);
  return s.endsWith('.0') ? s.substring(0, s.length - 2) : s;
}

class _PageExample {
  final String name;
  final String detail;
  final String code;
  final _PageLook look;

  const _PageExample({
    required this.name,
    required this.detail,
    required this.code,
    required this.look,
  });
}

/// How a mock printed page looks: orientation, margins, scaling and print
/// options applied to a sample table.
class _PageLook {
  final bool landscape;
  final PageMargins margins;
  final int columns;
  final int rows;
  final double scale;
  final bool fitWidth;
  final bool fitHeight;
  final bool gridLines;
  final bool headings;
  final bool centerH;
  final bool centerV;
  final bool blackAndWhite;
  final bool headerFooter;
  final String? footerLabel;

  const _PageLook({
    this.landscape = false,
    this.margins = PageMargins.normal,
    this.columns = 6,
    this.rows = 40,
    this.scale = 1.0,
    this.fitWidth = false,
    this.fitHeight = false,
    this.gridLines = false,
    this.headings = false,
    this.centerH = false,
    this.centerV = false,
    this.blackAndWhite = false,
    this.headerFooter = false,
    this.footerLabel,
  });
}

/// Draws a sheet of paper with its margins and a sample table laid out the
/// way Excel would print it.
///
/// When [pixelsPerMm] is set the page is drawn at that fixed scale;
/// otherwise it is fitted into the available space.
class _PagePreview extends StatelessWidget {
  final double widthMm;
  final double heightMm;
  final _PageLook look;
  final double? pixelsPerMm;

  const _PagePreview({
    required this.widthMm,
    required this.heightMm,
    required this.look,
    this.pixelsPerMm,
  });

  @override
  Widget build(BuildContext context) {
    final w = look.landscape ? math.max(widthMm, heightMm) : math.min(widthMm, heightMm);
    final h = look.landscape ? math.min(widthMm, heightMm) : math.max(widthMm, heightMm);

    return LayoutBuilder(
      builder: (context, constraints) {
        final ppm = pixelsPerMm ??
            math.min(constraints.maxWidth / w, constraints.maxHeight / h) * 0.95;
        return AnimatedContainer(
          duration: const Duration(milliseconds: 250),
          curve: Curves.easeOut,
          width: w * ppm,
          height: h * ppm,
          decoration: const BoxDecoration(
            boxShadow: [BoxShadow(color: Colors.black12, blurRadius: 4, offset: Offset(1, 2))],
          ),
          child: CustomPaint(painter: _PagePainter(look: look, ppm: ppm)),
        );
      },
    );
  }
}

class _PagePainter extends CustomPainter {
  final _PageLook look;
  final double ppm;

  _PagePainter({required this.look, required this.ppm});

  static const double _mmPerInch = 25.4;
  // Default Excel column (8.43 chars) and row (15 pt) sizes in millimetres.
  static const double _colMm = 17.0;
  static const double _rowMm = 5.3;
  static const double _headingColMm = 7.0;

  @override
  void paint(Canvas canvas, Size size) {
    canvas.drawRect(Offset.zero & size, Paint()..color = Colors.white);
    canvas.drawRect(
      Offset.zero & size,
      Paint()
        ..color = const Color(0xFFCBD5E1)
        ..style = PaintingStyle.stroke
        ..strokeWidth = 1,
    );

    double px(double inches) => inches * _mmPerInch * ppm;
    final m = look.margins;
    final printable = Rect.fromLTRB(
      px(m.left),
      px(m.top),
      size.width - px(m.right),
      size.height - px(m.bottom),
    );
    if (printable.width <= 2 || printable.height <= 2) return;

    final ink = look.blackAndWhite ? const Color(0xFF334155) : _accent;
    _drawDashedRect(canvas, printable, ink.withValues(alpha: 0.45));

    // Header / footer text placeholders.
    if (look.headerFooter) {
      final barPaint = Paint()..color = const Color(0xFF94A3B8);
      final barW = printable.width * 0.3;
      final barH = math.max(1.0, 1.4 * ppm);
      final headerY = px(m.header);
      final footerY = size.height - px(m.footer);
      if (headerY < printable.top) {
        canvas.drawRect(Rect.fromLTWH(printable.center.dx - barW / 2, headerY, barW, barH), barPaint);
      }
      if (footerY > printable.bottom) {
        canvas.drawRect(Rect.fromLTWH(printable.center.dx - barW / 2, footerY - barH, barW, barH), barPaint);
      }
    }
    if (look.footerLabel != null) {
      final tp = TextPainter(
        text: TextSpan(
          text: look.footerLabel,
          style: TextStyle(fontSize: math.max(6, 3 * ppm), color: const Color(0xFF475569), fontWeight: FontWeight.bold),
        ),
        textDirection: TextDirection.ltr,
      )..layout();
      tp.paint(canvas, Offset(printable.center.dx - tp.width / 2, size.height - px(m.footer) - tp.height));
    }

    // Sample table, scaled like Excel does.
    final headColMm = look.headings ? _headingColMm : 0.0;
    final headRowMm = look.headings ? _rowMm : 0.0;
    final naturalW = headColMm + look.columns * _colMm;
    final naturalH = headRowMm + look.rows * _rowMm;
    var s = look.scale;
    if (look.fitWidth || look.fitHeight) {
      final printableWmm = printable.width / ppm;
      final printableHmm = printable.height / ppm;
      final sw = look.fitWidth ? printableWmm / naturalW : double.infinity;
      final sh = look.fitHeight ? printableHmm / naturalH : double.infinity;
      s = math.min(1.0, math.min(sw, sh)); // Fit-to-page only shrinks.
    }

    final cw = _colMm * s * ppm;
    final rh = _rowMm * s * ppm;
    final hc = headColMm * s * ppm;
    final hr = headRowMm * s * ppm;
    final tableW = hc + look.columns * cw;
    final tableH = hr + look.rows * rh;

    var x = printable.left;
    var y = printable.top;
    if (look.centerH && tableW < printable.width) x += (printable.width - tableW) / 2;
    if (look.centerV && tableH < printable.height) y += (printable.height - tableH) / 2;

    canvas.save();
    canvas.clipRect(printable);

    if (look.headings) {
      final headingPaint = Paint()..color = const Color(0xFFE2E8F0);
      canvas.drawRect(Rect.fromLTWH(x, y, tableW, hr), headingPaint);
      canvas.drawRect(Rect.fromLTWH(x, y, hc, tableH), headingPaint);
    }

    final tx = x + hc;
    final ty = y + hr;
    canvas.drawRect(Rect.fromLTWH(tx, ty, look.columns * cw, rh), Paint()..color = ink);

    final textPaint = Paint()..color = const Color(0xFF94A3B8).withValues(alpha: 0.8);
    final barH = math.max(0.6, rh * 0.35);
    for (var r = 1; r < look.rows; r++) {
      final rowTop = ty + r * rh;
      if (rowTop > printable.bottom) break;
      for (var c = 0; c < look.columns; c++) {
        final barW = cw * (c == 0 ? 0.7 : 0.45 + ((r * 7 + c * 3) % 5) * 0.06);
        canvas.drawRect(Rect.fromLTWH(tx + c * cw + cw * 0.12, rowTop + (rh - barH) / 2, barW, barH), textPaint);
      }
    }

    if (look.gridLines) {
      final gridPaint = Paint()
        ..color = const Color(0xFFCBD5E1)
        ..strokeWidth = 0.5;
      for (var c = 0; c <= look.columns; c++) {
        final gx = tx + c * cw;
        canvas.drawLine(Offset(gx, ty), Offset(gx, ty + look.rows * rh), gridPaint);
      }
      for (var r = 0; r <= look.rows; r++) {
        final gy = ty + r * rh;
        if (gy > printable.bottom) break;
        canvas.drawLine(Offset(tx, gy), Offset(tx + look.columns * cw, gy), gridPaint);
      }
    }
    canvas.restore();

    // Signal content that does not fit and spills onto another page.
    if (tableW > printable.width + 0.5) {
      final tp = TextPainter(
        text: const TextSpan(
          text: '→ page 2',
          style: TextStyle(fontSize: 7, color: Color(0xFFEF4444), fontWeight: FontWeight.bold),
        ),
        textDirection: TextDirection.ltr,
      )..layout();
      tp.paint(canvas, Offset(size.width - tp.width - 2, printable.top - tp.height - 1));
    }
  }

  void _drawDashedRect(Canvas canvas, Rect rect, Color color) {
    final paint = Paint()
      ..color = color
      ..strokeWidth = 0.8;
    const dash = 3.0;
    const gap = 2.0;
    void line(Offset a, Offset b) {
      final total = (b - a).distance;
      final dir = (b - a) / total;
      for (double d = 0; d < total; d += dash + gap) {
        canvas.drawLine(a + dir * d, a + dir * math.min(d + dash, total), paint);
      }
    }

    line(rect.topLeft, rect.topRight);
    line(rect.topRight, rect.bottomRight);
    line(rect.bottomRight, rect.bottomLeft);
    line(rect.bottomLeft, rect.topLeft);
  }

  @override
  bool shouldRepaint(covariant _PagePainter oldDelegate) =>
      oldDelegate.look != look || oldDelegate.ppm != ppm;
}
