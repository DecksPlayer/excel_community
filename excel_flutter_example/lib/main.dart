import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart' hide Border, BorderStyle;
import 'package:flutter/painting.dart' as fp;
import 'package:flutter/services.dart';

import 'models/section_detail.dart';
import 'data/code_snippets.dart';
import 'services/excel_generator.dart';
import 'widgets/spreadsheet_preview.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'excel_community examples',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.blue),
        useMaterial3: true,
      ),
      home: const MyHomePage(title: 'excel_community examples'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  String _status = 'Press a button to generate an Excel with a chart.';
  bool _isGenerating = false;
  SelectedSection _selectedSection = SelectedSection.columnChart;

  Future<void> _handleGeneration() async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating Excel...';
    });

    try {
      String resultStatus;
      switch (_selectedSection) {
        case SelectedSection.simpleExcel:
          resultStatus = await ExcelGenerator.generateSimpleExcel();
          break;
        case SelectedSection.imageEmbedding:
          resultStatus = await ExcelGenerator.generateExcelWithImage();
          break;
        case SelectedSection.columnChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(ChartType.column);
          break;
        case SelectedSection.lineChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(ChartType.line);
          break;
        case SelectedSection.pieChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(ChartType.pie);
          break;
        case SelectedSection.areaChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(ChartType.area);
          break;
        case SelectedSection.doughnutChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(ChartType.doughnut);
          break;
        case SelectedSection.radarChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(ChartType.radar);
          break;
        case SelectedSection.barChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(ChartType.bar);
          break;
        case SelectedSection.scatterChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(ChartType.scatter);
          break;
        case SelectedSection.textStyles:
          resultStatus = await ExcelGenerator.generateUnderlineStyles();
          break;
        case SelectedSection.numberFormats:
          resultStatus = await ExcelGenerator.generateNumberFormats();
          break;
        case SelectedSection.multiSheets:
          resultStatus = await ExcelGenerator.generateMultiSheets();
          break;
        case SelectedSection.allCharts:
          resultStatus = await ExcelGenerator.generateAllCharts();
          break;
        case SelectedSection.fullDemo:
          resultStatus = await ExcelGenerator.generateFullDemo();
          break;
      }
      setState(() {
        _status = resultStatus;
      });
    } catch (e, stackTrace) {
      setState(() {
        _status = 'Error: $e';
      });
      if (kDebugMode) {
        print('Error generating Excel: $e');
        print(stackTrace);
      }
    } finally {
      setState(() {
        _isGenerating = false;
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    final isDesktop = MediaQuery.of(context).size.width >= 950;
    final detail = _getSectionDetail(_selectedSection);

    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC), // Slate 50
      appBar: AppBar(
        backgroundColor: const Color(0xFF0F172A), // Slate 900
        title: Row(
          children: [
            const Icon(Icons.table_view, color: Color(0xFF10B981), size: 24),
            const SizedBox(width: 8),
            Text(
              widget.title,
              style: const TextStyle(
                color: Colors.white,
                fontWeight: FontWeight.bold,
                fontSize: 18,
              ),
            ),
          ],
        ),
        actions: [
          if (kIsWeb)
            Container(
              margin: const EdgeInsets.only(right: 16),
              child: Container(
                padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 6),
                decoration: BoxDecoration(
                  color: const Color(0xFF1E293B), // Slate 800
                  borderRadius: BorderRadius.circular(20),
                  border: fp.Border.all(color: Colors.white10),
                ),
                child: const Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    Icon(Icons.language, size: 14, color: Color(0xFF10B981)),
                    SizedBox(width: 6),
                    Text(
                      'Wasm Web Optimized',
                      style: TextStyle(color: Colors.white70, fontSize: 11, fontWeight: FontWeight.bold),
                    ),
                  ],
                ),
              ),
            ),
        ],
        elevation: 0,
        iconTheme: const IconThemeData(color: Colors.white),
      ),
      drawer: isDesktop ? null : Drawer(
        child: _buildSidebar(context),
      ),
      body: Row(
        children: [
          if (isDesktop)
            Container(
              width: 260,
              decoration: const BoxDecoration(
                color: Colors.white,
                border: fp.Border(right: fp.BorderSide(color: Color(0xFFE2E8F0))),
              ),
              child: _buildSidebar(context),
            ),
          Expanded(
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(24.0),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  _buildHeaderCard(context, detail),
                  const SizedBox(height: 24),
                  if (isDesktop)
                    Row(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Expanded(flex: 3, child: _buildCodeViewCard(context, detail)),
                        const SizedBox(width: 24),
                        Expanded(flex: 2, child: _buildPreviewCard(context, detail)),
                      ],
                    )
                  else ...[
                    _buildPreviewCard(context, detail),
                    const SizedBox(height: 24),
                    _buildCodeViewCard(context, detail),
                  ],
                ],
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildSidebar(BuildContext context) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        const Padding(
          padding: EdgeInsets.fromLTRB(20, 20, 20, 8),
          child: Text(
            'EXCEL SPREADSHEETS',
            style: TextStyle(
              fontSize: 10,
              fontWeight: FontWeight.bold,
              color: Color(0xFF94A3B8), // Slate 400
              letterSpacing: 1.2,
            ),
          ),
        ),
        Expanded(
          child: ListView(
            padding: const EdgeInsets.symmetric(horizontal: 8),
            children: [
              _buildSidebarItem(SelectedSection.simpleExcel, 'Quick Start (No Chart)', Icons.bolt, Colors.amber),
              const Padding(
                padding: EdgeInsets.fromLTRB(12, 16, 12, 8),
                child: Text('Chart Types', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 11, color: Color(0xFF64748B))),
              ),
              _buildSidebarItem(SelectedSection.columnChart, 'Column Chart', Icons.bar_chart, Colors.blue),
              _buildSidebarItem(SelectedSection.lineChart, 'Line Chart', Icons.show_chart, Colors.orange),
              _buildSidebarItem(SelectedSection.pieChart, 'Pie Chart', Icons.pie_chart, Colors.red),
              _buildSidebarItem(SelectedSection.areaChart, 'Area Chart', Icons.area_chart, Colors.purple),
              _buildSidebarItem(SelectedSection.doughnutChart, 'Doughnut Chart', Icons.donut_large, Colors.teal),
              _buildSidebarItem(SelectedSection.radarChart, 'Radar Chart', Icons.radar, Colors.indigo),
              _buildSidebarItem(SelectedSection.barChart, 'Bar Chart', Icons.horizontal_split, Colors.pink),
              _buildSidebarItem(SelectedSection.scatterChart, 'Scatter Chart', Icons.scatter_plot, Colors.deepOrange),
              _buildSidebarItem(SelectedSection.imageEmbedding, 'Image Embedding', Icons.image_outlined, Colors.teal),
              const Padding(
                padding: EdgeInsets.fromLTRB(12, 16, 12, 8),
                child: Text('Style & Formatting', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 11, color: Color(0xFF64748B))),
              ),
              _buildSidebarItem(SelectedSection.textStyles, 'Text Underlines & Fills', Icons.format_underlined, Colors.deepPurple),
              _buildSidebarItem(SelectedSection.numberFormats, 'Number Formatting', Icons.pin, Colors.teal),
              _buildSidebarItem(SelectedSection.multiSheets, 'Multi-Worksheets', Icons.layers_outlined, Colors.cyan),
              const Padding(
                padding: EdgeInsets.fromLTRB(12, 16, 12, 8),
                child: Text('Demonstrations', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 11, color: Color(0xFF64748B))),
              ),
              _buildSidebarItem(SelectedSection.allCharts, 'All 8 Charts Grid', Icons.grid_view, Colors.blueGrey),
              _buildSidebarItem(SelectedSection.fullDemo, 'Full Sheet Report', Icons.star, Colors.amber.shade700),
            ],
          ),
        ),
      ],
    );
  }

  Widget _buildSidebarItem(SelectedSection section, String label, IconData icon, Color color) {
    final isSelected = _selectedSection == section;
    return Container(
      margin: const EdgeInsets.only(bottom: 2),
      decoration: BoxDecoration(
        color: isSelected ? color.withOpacity(0.08) : Colors.transparent,
        borderRadius: BorderRadius.circular(8),
      ),
      child: ListTile(
        onTap: () {
          setState(() {
            _selectedSection = section;
          });
          if (Scaffold.of(context).isDrawerOpen) {
            Navigator.pop(context);
          }
        },
        leading: Icon(
          icon,
          color: isSelected ? color : const Color(0xFF64748B),
          size: 18,
        ),
        title: Text(
          label,
          style: TextStyle(
            color: isSelected ? const Color(0xFF0F172A) : const Color(0xFF475569),
            fontWeight: isSelected ? FontWeight.bold : FontWeight.normal,
            fontSize: 13,
          ),
        ),
        dense: true,
        visualDensity: const VisualDensity(vertical: -2),
      ),
    );
  }

  Widget _buildHeaderCard(BuildContext context, SectionDetail detail) {
    return Card(
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
        side: const BorderSide(color: Color(0xFFE2E8F0)),
      ),
      color: Colors.white,
      child: Padding(
        padding: const EdgeInsets.all(20.0),
        child: Row(
          children: [
            Container(
              padding: const EdgeInsets.all(12),
              decoration: BoxDecoration(
                color: detail.themeColor.withOpacity(0.1),
                borderRadius: BorderRadius.circular(12),
              ),
              child: Icon(detail.icon, color: detail.themeColor, size: 28),
            ),
            const SizedBox(width: 16),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    detail.title,
                    style: const TextStyle(
                      fontSize: 20,
                      fontWeight: FontWeight.bold,
                      color: Color(0xFF0F172A),
                    ),
                  ),
                  const SizedBox(height: 4),
                  Text(
                    detail.description,
                    style: const TextStyle(
                      fontSize: 13,
                      color: Color(0xFF64748B),
                    ),
                  ),
                ],
              ),
            ),
            const SizedBox(width: 16),
            ElevatedButton.icon(
              onPressed: _isGenerating ? null : _handleGeneration,
              style: ElevatedButton.styleFrom(
                backgroundColor: detail.themeColor,
                foregroundColor: Colors.white,
                padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 14),
                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                elevation: 0,
              ),
              icon: _isGenerating
                  ? const SizedBox(
                      width: 16,
                      height: 16,
                      child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2),
                    )
                  : const Icon(Icons.file_download_outlined, size: 16),
              label: Text(
                _isGenerating ? 'Generating...' : 'Generate Spreadsheet',
                style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 13),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildCodeViewCard(BuildContext context, SectionDetail detail) {
    return Card(
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
        side: const BorderSide(color: Color(0xFFE2E8F0)),
      ),
      color: const Color(0xFF0F172A), // Slate 900
      child: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                const Row(
                  children: [
                    Icon(Icons.code, color: Color(0xFF38BDF8), size: 18), // Sky 400
                    SizedBox(width: 8),
                    Text(
                      'Dart Source Code',
                      style: TextStyle(
                        color: Color(0xFFE2E8F0),
                        fontWeight: FontWeight.bold,
                        fontSize: 13,
                      ),
                    ),
                  ],
                ),
                TextButton.icon(
                  onPressed: () {
                    Clipboard.setData(ClipboardData(text: detail.codeSnippet));
                    ScaffoldMessenger.of(context).showSnackBar(
                      const SnackBar(
                        content: Row(
                          children: [
                            Icon(Icons.check_circle, color: Color(0xFF10B981)),
                            SizedBox(width: 8),
                            Text('Code copied to clipboard successfully!'),
                          ],
                        ),
                        backgroundColor: Color(0xFF1E293B),
                        behavior: SnackBarBehavior.floating,
                      ),
                    );
                  },
                  icon: const Icon(Icons.copy, size: 13, color: Color(0xFF38BDF8)),
                  label: const Text(
                    'Copy Code',
                    style: TextStyle(color: Color(0xFF38BDF8), fontSize: 12, fontWeight: FontWeight.bold),
                  ),
                ),
              ],
            ),
            const Divider(color: Colors.white10),
            const SizedBox(height: 8),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(12),
              decoration: BoxDecoration(
                color: const Color(0xFF020617), // Slate 950
                borderRadius: BorderRadius.circular(8),
              ),
              child: SelectableText(
                detail.codeSnippet,
                style: const TextStyle(
                  fontFamily: 'monospace',
                  fontSize: 11,
                  color: Color(0xFFF1F5F9),
                  height: 1.4,
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildPreviewCard(BuildContext context, SectionDetail detail) {
    return Card(
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
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
                Icon(Icons.remove_red_eye_outlined, color: detail.themeColor, size: 18),
                const SizedBox(width: 8),
                const Text(
                  'Sheet Preview & Details',
                  style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13, color: Color(0xFF0F172A)),
                ),
              ],
            ),
            const SizedBox(height: 10),
            const Divider(color: Color(0xFFF1F5F9)),
            const SizedBox(height: 6),
            const Text(
              'Spreadsheet highlights:',
              style: TextStyle(fontWeight: FontWeight.bold, fontSize: 12, color: Color(0xFF475569)),
            ),
            const SizedBox(height: 6),
            Column(
              children: detail.highlights.map((h) {
                return Padding(
                  padding: const EdgeInsets.only(bottom: 6),
                  child: Row(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      const Icon(Icons.check_circle_outline, size: 14, color: Color(0xFF10B981)),
                      const SizedBox(width: 6),
                      Expanded(
                        child: Text(
                          h,
                          style: const TextStyle(fontSize: 11, color: Color(0xFF64748B)),
                        ),
                      ),
                    ],
                  ),
                );
              }).toList(),
            ),
            const SizedBox(height: 12),
            const Text(
              'Generation Status:',
              style: TextStyle(fontWeight: FontWeight.bold, fontSize: 12, color: Color(0xFF475569)),
            ),
            const SizedBox(height: 4),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
              decoration: BoxDecoration(
                color: const Color(0xFFF8FAFC),
                borderRadius: BorderRadius.circular(8),
                border: fp.Border.all(color: const Color(0xFFE2E8F0)),
              ),
              child: Text(
                _status,
                style: const TextStyle(fontSize: 11, color: Color(0xFF334155), height: 1.3),
              ),
            ),
            const SizedBox(height: 16),
            const Text(
              'Mock Spreadsheet Anchoring View:',
              style: TextStyle(fontWeight: FontWeight.bold, fontSize: 12, color: Color(0xFF475569)),
            ),
            const SizedBox(height: 10),
            const Divider(color: Color(0xFFF1F5F9)),
            const SizedBox(height: 12),
            SpreadsheetPreview(selectedSection: _selectedSection, detail: detail),
          ],
        ),
      ),
    );
  }

  SectionDetail _getSectionDetail(SelectedSection section) {
    switch (section) {
      case SelectedSection.simpleExcel:
        return SectionDetail(
          title: 'Quick Start Excel',
          description: 'A basic Excel spreadsheet containing formatted texts and integer values without any charts.',
          icon: Icons.bolt,
          themeColor: Colors.amber,
          highlights: [
            'Basic single-sheet workbook creation',
            'Supports Text and Integer cell values',
            'Fast saving using File (Desktop) or auto-downloads (Web)',
          ],
          codeSnippet: simpleExcelSnippet,
        );
      case SelectedSection.columnChart:
        return SectionDetail(
          title: 'Column Chart',
          description: 'Vertical clustered columns for category data comparison.',
          icon: Icons.bar_chart,
          themeColor: Colors.blue,
          highlights: [
            'Professional standard Excel Column charts',
            'Compliance with ECMA-376 OOXML standards',
            'Supports multiple data series overlay',
          ],
          codeSnippet: columnChartSnippet,
        );
      case SelectedSection.lineChart:
        return SectionDetail(
          title: 'Line Chart',
          description: 'A trend display over continuous intervals with circular series markers.',
          icon: Icons.show_chart,
          themeColor: Colors.orange,
          highlights: [
            'Thick line borders (28575 EMUs) for modern aesthetics',
            'Circular data markers matching series line colors',
            'Perfect for time-series and continuous trends',
          ],
          codeSnippet: lineChartSnippet,
        );
      case SelectedSection.pieChart:
        return SectionDetail(
          title: 'Pie Chart',
          description: 'Circular slice distribution demonstrating relative percentages.',
          icon: Icons.pie_chart,
          themeColor: Colors.red,
          highlights: [
            'Supports single series segment visualization',
            'Shuffles standard 20-color palette for unique slices',
            'Ensures no repeated colors within the same pie chart',
          ],
          codeSnippet: pieChartSnippet,
        );
      case SelectedSection.areaChart:
        return SectionDetail(
          title: 'Area Chart',
          description: 'Filled area series overlay showing accumulation over categories.',
          icon: Icons.area_chart,
          themeColor: Colors.purple,
          highlights: [
            'Semi-transparent fills (50% opacity) for overlapping series visibility',
            'Opaque contours (90% opacity, 2.25pt thick) defining series paths',
            'Clean vector output for modern reporting sheets',
          ],
          codeSnippet: areaChartSnippet,
        );
      case SelectedSection.doughnutChart:
        return SectionDetail(
          title: 'Doughnut Chart',
          description: 'A circular chart with a central cutout, displaying part-to-whole data.',
          icon: Icons.donut_large,
          themeColor: Colors.teal,
          highlights: [
            'Built-in hole size set to 50% for optimal readability',
            'Shuffled colors from a professional 20-color palette',
            'Compatible with Microsoft Excel, Excel 365, and Google Sheets',
          ],
          codeSnippet: doughnutChartSnippet,
        );
      case SelectedSection.radarChart:
        return SectionDetail(
          title: 'Radar Chart',
          description: 'Spider chart showing multi-variable comparison in a radial layout.',
          icon: Icons.radar,
          themeColor: Colors.indigo,
          highlights: [
            'Filled radar series support with 45% transparency layers',
            'Bold contour boundary lines with 85% opacity settings',
            'Clean radial axes grid generation compliant with OOXML standards',
          ],
          codeSnippet: radarChartSnippet,
        );
      case SelectedSection.barChart:
        return SectionDetail(
          title: 'Bar Chart',
          description: 'Horizontal clustered bars ideal for displaying descriptive labels.',
          icon: Icons.horizontal_split,
          themeColor: Colors.pink,
          highlights: [
            'Horizontal clustered columns formatting',
            'Consistent 12-color rotating series palette',
            'Thin borders matching series color fills (0.75pt)',
          ],
          codeSnippet: barChartSnippet,
        );
      case SelectedSection.scatterChart:
        return SectionDetail(
          title: 'Scatter Chart',
          description: 'XY plotting representing numerical values mapping correlations.',
          icon: Icons.scatter_plot,
          themeColor: Colors.deepOrange,
          highlights: [
            'Requires both numeric X values and numeric Y values',
            'Solid data points highlighted by a clean white outer halo',
            'Aids readability when dealing with heavy overlapping point maps',
          ],
          codeSnippet: scatterChartSnippet,
        );
      case SelectedSection.imageEmbedding:
        return SectionDetail(
          title: 'Image Embedding',
          description: 'Embed PNG, JPEG, SVG, WebP, and other image formats into your worksheets.',
          icon: Icons.image_outlined,
          themeColor: Colors.green,
          highlights: [
            'Embed images with custom pixel dimensions',
            'Supports positioning using column/row coordinates and EMU offsets',
            'Compatible with PNG, JPEG, BMP, GIF, SVG, and WebP formats',
          ],
          codeSnippet: imageEmbeddingSnippet,
        );
      case SelectedSection.textStyles:
        return SectionDetail(
          title: 'Text Styles & Underlines',
          description: 'Format fonts, sizes, italic/bold, colors, backgrounds, strikethrough, and underlines.',
          icon: Icons.format_underlined,
          themeColor: Colors.deepPurple,
          highlights: [
            'Underline options: Underline.None, Underline.Single, Underline.Double',
            'Font family selection (Arial, Calibri, Comic Sans, etc.)',
            'Text strikethrough styling and color preservation',
          ],
          codeSnippet: textStylesSnippet,
        );
      case SelectedSection.numberFormats:
        return SectionDetail(
          title: 'Number Formatting',
          description: 'Built-in standard and custom number formatting rules.',
          icon: Icons.pin,
          themeColor: Colors.teal,
          highlights: [
            'Integrates built-in ECMA-376 Standard formats (0-44)',
            'Formats float decimals, scientific exponents, percentages, currencies',
            'Accounting currency values with indent alignment support (ID 44)',
          ],
          codeSnippet: numberFormatsSnippet,
        );
      case SelectedSection.multiSheets:
        return SectionDetail(
          title: 'Multi-Worksheets',
          description: 'Create and organize multiple sheets inside a single Excel workbook.',
          icon: Icons.layers_outlined,
          themeColor: Colors.cyan,
          highlights: [
            'Create, rename, and select sheets on the fly',
            'Delete default template sheets to clean up output',
            'Isolate data schemas across dedicated worksheets',
          ],
          codeSnippet: multiSheetsSnippet,
        );
      case SelectedSection.allCharts:
        return SectionDetail(
          title: 'All Charts Grid',
          description: 'A composite grid layout showing all 8 supported charts side-by-side.',
          icon: Icons.grid_view,
          themeColor: Colors.blueGrey,
          highlights: [
            'Grid sheet layout positioning charts dynamically',
            'Leverages distinct custom styles for each chart builder type',
            'Full compliance validation under a single workbook relationship XML',
          ],
          codeSnippet: allChartsSnippet,
        );
      case SelectedSection.fullDemo:
        return SectionDetail(
          title: 'Full Sheet Demo',
          description: 'A complete real-world reporting sheet with headers, custom cell styles, formulas, and charts.',
          icon: Icons.star,
          themeColor: Colors.amber.shade700,
          highlights: [
            'Alternating row styles with customized thin grid borders',
            'Emits Excel formulas: SUM(A1:A5) and arithmetic equations',
            'Anchors column charts automatically inside a styled spreadsheet',
          ],
          codeSnippet: fullDemoSnippet,
        );
    }
  }
}
