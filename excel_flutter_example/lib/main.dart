import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart' hide Border, BorderStyle;
import 'package:flutter/painting.dart' as fp;
import 'package:flutter_svg/flutter_svg.dart';

import 'models/section_detail.dart';
import 'services/excel_generator.dart';
import 'data/section_details.dart';
import 'widgets/sidebar.dart';
import 'widgets/header_card.dart';
import 'widgets/about_view.dart';
import 'widgets/code_view_card.dart';
import 'widgets/preview_card.dart';
import 'widgets/data_export_view.dart';
import 'widgets/data_validation_view.dart';
import 'widgets/fonts_styles_view.dart';
import 'widgets/grouping_view.dart';
import 'widgets/hyperlinks_view.dart';
import 'widgets/number_formats_view.dart';
import 'widgets/page_setup_view.dart';
import 'widgets/tables_view.dart';

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
  SelectedSection _selectedSection = SelectedSection.about;

  Future<void> _handleGeneration() async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating Excel...';
    });

    try {
      String resultStatus;
      switch (_selectedSection) {
        case SelectedSection.about:
          resultStatus = 'Welcome to excel_community!';
          break;
        case SelectedSection.simpleExcel:
          resultStatus = await ExcelGenerator.generateSimpleExcel();
          break;
        case SelectedSection.imageEmbedding:
          resultStatus = await ExcelGenerator.generateExcelWithImage();
          break;
        case SelectedSection.columnChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(
            ChartType.column,
          );
          break;
        case SelectedSection.lineChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(
            ChartType.line,
          );
          break;
        case SelectedSection.pieChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(
            ChartType.pie,
          );
          break;
        case SelectedSection.areaChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(
            ChartType.area,
          );
          break;
        case SelectedSection.doughnutChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(
            ChartType.doughnut,
          );
          break;
        case SelectedSection.radarChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(
            ChartType.radar,
          );
          break;
        case SelectedSection.barChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(
            ChartType.bar,
          );
          break;
        case SelectedSection.scatterChart:
          resultStatus = await ExcelGenerator.generateExcelWithChart(
            ChartType.scatter,
          );
          break;
        case SelectedSection.fontsStyles:
          resultStatus = await ExcelGenerator.generateFontsStyles();
          break;
        case SelectedSection.numberFormats:
          resultStatus = await ExcelGenerator.generateNumberFormats();
          break;
        case SelectedSection.multiSheets:
          resultStatus = await ExcelGenerator.generateMultiSheets();
          break;
        case SelectedSection.pivotTemplate:
          resultStatus = await ExcelGenerator.generatePivotTemplate();
          break;
        case SelectedSection.cellLocking:
          resultStatus = await ExcelGenerator.generateLockedCellsReport();
          break;
        case SelectedSection.freezePanes:
          resultStatus = await ExcelGenerator.generateFreezePanes();
          break;
        case SelectedSection.multiFreezePanes:
          resultStatus = await ExcelGenerator.generateMultiFreezePanes();
          break;
        case SelectedSection.hiddenColumns:
          resultStatus = await ExcelGenerator.generateHiddenColumns();
          break;
        case SelectedSection.mergedCells:
          resultStatus = await ExcelGenerator.generateMergedCells();
          break;
        case SelectedSection.multiPageCharts:
          resultStatus = await ExcelGenerator.generateMultiPageCharts();
          break;
        case SelectedSection.allCharts:
          resultStatus = await ExcelGenerator.generateAllCharts();
          break;
        case SelectedSection.fullDemo:
          resultStatus = await ExcelGenerator.generateFullDemo();
          break;
        case SelectedSection.cellComments:
          resultStatus = await ExcelGenerator.generateCellComments();
          break;
        case SelectedSection.conditionalFormatting:
          resultStatus = await ExcelGenerator.generateConditionalFormatting();
          break;
        case SelectedSection.chartDataLabels:
          resultStatus = await ExcelGenerator.generateChartDataLabels();
          break;
        case SelectedSection.chartColors:
          resultStatus = await ExcelGenerator.generateChartColors();
          break;
        case SelectedSection.newCharts:
          resultStatus = await ExcelGenerator.generateNewCharts();
          break;
        case SelectedSection.readAsset:
          resultStatus = await ExcelGenerator.readAssetExcel();
          break;
        case SelectedSection.formulasDisplayText:
          resultStatus = await ExcelGenerator.generateFormulasDisplayText();
          break;
        case SelectedSection.autoFilter:
          resultStatus = await ExcelGenerator.generateAutoFilter();
          break;
        case SelectedSection.tabColor:
          resultStatus = await ExcelGenerator.generateTabColors();
          break;
        case SelectedSection.pageSetup:
          resultStatus = await ExcelGenerator.generatePageSetup();
          break;
        case SelectedSection.dataExport:
          resultStatus = await ExcelGenerator.generateDataExport();
          break;
        case SelectedSection.hyperlinks:
          resultStatus = await ExcelGenerator.generateHyperlinks();
          break;
        case SelectedSection.dataValidation:
          resultStatus = await ExcelGenerator.generateDataValidation();
          break;
        case SelectedSection.grouping:
          resultStatus = await ExcelGenerator.generateGrouping();
          break;
        case SelectedSection.tables:
          resultStatus = await ExcelGenerator.generateTables();
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
    final detail = getSectionDetail(_selectedSection);

    return Scaffold(
      backgroundColor: const Color(0xFFF8FAFC), // Slate 50
      drawer: isDesktop
          ? null
          : Drawer(
              child: Sidebar(
                selectedSection: _selectedSection,
                onSectionSelected: (section) {
                  setState(() {
                    _selectedSection = section;
                  });
                },
              ),
            ),
      appBar: AppBar(
        backgroundColor: const Color(0xFF0F172A), // Slate 900
        title: Text(
          widget.title,
          style: const TextStyle(
            color: Colors.white,
            fontWeight: FontWeight.bold,
            fontSize: 18,
          ),
        ),
        actions: const [],
        elevation: 0,
        iconTheme: const IconThemeData(color: Colors.white),
        leading: Builder(
          builder: (BuildContext context) {
            return IconButton(
              icon: SvgPicture.asset('assets/logo.svg', width: 28, height: 28),
              onPressed: () {
                showHideDrawer(context);
              },
            );
          },
        ),
      ),
      body: Row(
        children: [
          if (isDesktop)
            Container(
              width: 260,
              decoration: const BoxDecoration(
                color: Colors.white,
                border: fp.Border(
                  right: fp.BorderSide(color: Color(0xFFE2E8F0)),
                ),
              ),
              child: Sidebar(
                selectedSection: _selectedSection,
                onSectionSelected: (section) {
                  setState(() {
                    _selectedSection = section;
                  });
                },
              ),
            ),
          Expanded(
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(24.0),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  if (_selectedSection == SelectedSection.about)
                    const AboutView()
                  else if (_selectedSection == SelectedSection.fontsStyles)
                    FontsStylesView(
                      isGenerating: _isGenerating,
                      onGenerate: _handleGeneration,
                      status: _status,
                    )
                  else if (_selectedSection == SelectedSection.numberFormats)
                    NumberFormatsView(
                      isGenerating: _isGenerating,
                      onGenerate: _handleGeneration,
                      status: _status,
                    )
                  else if (_selectedSection == SelectedSection.tables)
                    TablesView(
                      isGenerating: _isGenerating,
                      onGenerate: _handleGeneration,
                      status: _status,
                    )
                  else if (_selectedSection == SelectedSection.grouping)
                    GroupingView(
                      isGenerating: _isGenerating,
                      onGenerate: _handleGeneration,
                      status: _status,
                    )
                  else if (_selectedSection == SelectedSection.dataValidation)
                    DataValidationView(
                      isGenerating: _isGenerating,
                      onGenerate: _handleGeneration,
                      status: _status,
                    )
                  else if (_selectedSection == SelectedSection.hyperlinks)
                    HyperlinksView(
                      isGenerating: _isGenerating,
                      onGenerate: _handleGeneration,
                      status: _status,
                    )
                  else if (_selectedSection == SelectedSection.dataExport)
                    DataExportView(
                      isGenerating: _isGenerating,
                      onGenerate: _handleGeneration,
                      status: _status,
                    )
                  else if (_selectedSection == SelectedSection.pageSetup)
                    PageSetupView(
                      isGenerating: _isGenerating,
                      onGenerate: _handleGeneration,
                      status: _status,
                    )
                  else ...[
                    HeaderCard(
                      detail: detail,
                      isGenerating: _isGenerating,
                      onGenerate: _handleGeneration,
                    ),
                    const SizedBox(height: 24),
                    if (isDesktop)
                      Row(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Expanded(
                            flex: 3,
                            child: CodeViewCard(detail: detail),
                          ),
                          const SizedBox(width: 24),
                          Expanded(
                            flex: 2,
                            child: PreviewCard(
                              detail: detail,
                              status: _status,
                              selectedSection: _selectedSection,
                            ),
                          ),
                        ],
                      )
                    else ...[
                      PreviewCard(
                        detail: detail,
                        status: _status,
                        selectedSection: _selectedSection,
                      ),
                      const SizedBox(height: 24),
                      CodeViewCard(detail: detail),
                    ],
                  ],
                ],
              ),
            ),
          ),
        ],
      ),
    );
  }

  void showHideDrawer(BuildContext context) {
    if (!Scaffold.hasDrawer(context)) return;

    if (!Scaffold.of(context).isDrawerOpen) {
      Scaffold.of(context).openDrawer();
      return;
    }
    Scaffold.of(context).closeDrawer();
  }
}
