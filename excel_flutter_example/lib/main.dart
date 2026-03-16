import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart' hide Border, BorderStyle;
import 'package:excel_community/excel_community.dart';
import 'package:file_picker/file_picker.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Chart Example',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.blue),
        useMaterial3: true,
      ),
      home: const MyHomePage(title: 'Excel Chart Example'),
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

  Future<void> _generateSimpleExcel() async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating simple Excel (NO chart)...';
    });

    try {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // Add simple data
      sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Name"));
      sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Age"));
      sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Alice"));
      sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(30));
      sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Bob"));
      sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(25));

      if (kIsWeb) {
        if (kDebugMode) {
          print('Generating simple Excel for Web...');
        }
        
        final bytes = excel.save(fileName: 'simple_no_chart.xlsx');
        
        if (bytes != null && bytes.isNotEmpty) {
          setState(() {
            _status = '✅ Simple Excel generated successfully!\n'
                'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
                'The download should start automatically.\n'
                '\n📥 Check your Downloads folder\n'
                '📌 File: simple_no_chart.xlsx';
          });
          
          if (kDebugMode) {
            print('✅ Simple Excel saved for web: ${bytes.length} bytes');
          }
        } else {
          setState(() {
            _status = '❌ Error: Failed to generate Excel file.\n'
                'The file is empty or could not be encoded.';
          });
        }
      } else {
        var bytes = excel.encode();
        
        if (bytes == null) {
          setState(() {
            _status = 'Error: Failed to encode Excel file.';
          });
          return;
        }
        
        String? outputFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Save Simple Excel File',
          fileName: 'simple_no_chart.xlsx',
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
        );

        if (outputFile != null) {
          final file = File(outputFile);
          await file.writeAsBytes(bytes);
          
          final savedFileSize = await file.length();
          
          setState(() {
            _status = '✅ Simple Excel saved successfully!\n'
                'Location: $outputFile\n'
                'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
                '\n📌 Open with Excel to verify it works';
          });
        } else {
          setState(() {
            _status = 'Save cancelled.';
          });
        }
      }
    } catch (e, stackTrace) {
      setState(() {
        _status = 'Error: $e';
      });
      if (kDebugMode) {
        print('Error: $e');
        print('Stack trace: $stackTrace');
      }
    } finally {
      setState(() {
        _isGenerating = false;
      });
    }
  }

  Future<void> _generateExcelWithChart(ChartType type) async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating Excel with chart...';
    });

    try {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // Add Headers
      sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
      sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Value 1"));
      sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Value 2"));

      // Add Data
      final data = [
        ['Jan', 10, 15],
        ['Feb', 20, 25],
        ['Mar', 15, 30],
        ['Apr', 25, 20],
        ['May', 30, 35],
        ['Jun', 20, 40],
      ];

      for (var i = 0; i < data.length; i++) {
        sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(data[i][0] as String));
        sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), IntCellValue(data[i][1] as int));
        sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), IntCellValue(data[i][2] as int));
      }

      Chart chart;
      final series = [
        ChartSeries(
          name: "Series 1",
          categoriesRange: r"Sheet1!$A$2:$A$7",
          valuesRange: r"Sheet1!$B$2:$B$7",
        ),
        ChartSeries(
          name: "Series 2",
          categoriesRange: r"Sheet1!$A$2:$A$7",
          valuesRange: r"Sheet1!$C$2:$C$7",
        ),
      ];

      final anchor = ChartAnchor.at(column: 5, row: 1, width: 10, height: 15);

      switch (type) {
        case ChartType.column:
          chart = ColumnChart(
            title: "Monthly Data (Column)",
            series: series,
            anchor: anchor,
          );
        case ChartType.bar:
          chart = BarChart(
            title: "Monthly Data (Bar)",
            series: series,
            anchor: anchor,
          );
        case ChartType.line:
          chart = LineChart(
            title: "Monthly Data (Line)",
            series: series,
            anchor: anchor,
          );
        case ChartType.area:
          chart = AreaChart(
            title: "Monthly Data (Area)",
            series: series,
            anchor: anchor,
          );
        case ChartType.pie:
          chart = PieChart(
            title: "Pie Chart Example",
            series: [series[0]], // Pie chart usually only takes one series
            anchor: anchor,
          );
        case ChartType.doughnut:
          chart = DoughnutChart(
            title: "Doughnut Chart Example",
            series: [series[0]], // Doughnut chart usually only takes one series
            anchor: anchor,
          );
        case ChartType.radar:
          chart = RadarChart(
            title: "Radar Chart Example",
            series: series,
            anchor: anchor,
            filled: true,
          );
        case ChartType.scatter:
          chart = ScatterChart(
            title: "Scatter Chart Example",
            series: series,
            anchor: anchor,
          );
      }

      sheet.addChart(chart);
      
      if (kDebugMode) {
        print('Chart added to sheet');
      }

      if (kIsWeb) {
        if (kDebugMode) {
          print('Generating Excel for Web...');
        }
        
        final bytes = excel.save(fileName: 'chart_example.xlsx');
        
        if (bytes != null && bytes.isNotEmpty) {
          setState(() {
            _status = '✅ Excel generated successfully!\n'
                'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
                'The download should start automatically.\n'
                '\n📥 Check your Downloads folder\n'
                '📌 File: chart_example.xlsx';
          });
          
          if (kDebugMode) {
            print('✅ Excel saved for web: ${bytes.length} bytes');
          }
        } else {
          setState(() {
            _status = '❌ Error: Failed to generate Excel file for web.\n'
                'The file is empty or could not be encoded.';
          });
          
          if (kDebugMode) {
            print('❌ Error: excel.save() returned null or empty');
          }
        }
      } else {
        var bytes = excel.encode();
        
        if (bytes == null) {
          setState(() {
            _status = 'Error: Failed to encode Excel file.';
          });
          return;
        }
        
        String? outputFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Save Excel File',
          fileName: 'chart_example.xlsx',
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
        );

        if (outputFile != null) {
          final file = File(outputFile);
        
          await file.writeAsBytes(bytes);
          
          // Verify the file was written correctly
          final savedFileSize = await file.length();
          
          setState(() {
            _status = '✅ Excel saved successfully!\n'
                'Location: $outputFile\n'
                'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
                '\n📌 IMPORTANT: Open with Excel, not a text editor!';
          });
          
          if (kDebugMode) {
            print('✅ File saved: $outputFile');
            print('📊 Bytes generated: ${bytes.length}');
            print('💾 File size: $savedFileSize bytes');
          }
        } else {
          setState(() {
            _status = 'Save cancelled.';
          });
        }
      }
    } catch (e, stackTrace) {
      setState(() {
        _status = 'Error: $e';
      });
      if (kDebugMode) {
        print('Error generating Excel with chart: $e');
        print('Stack trace: $stackTrace');
      }
    } finally {
      setState(() {
        _isGenerating = false;
      });
    }
  }

  Future<void> _generateFullExample() async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating Full Demo Excel...';
    });

    try {
      var excel = Excel.createExcel();
      // Rename default sheet
      var sheetName = 'Full Demo';
      excel.rename('Sheet1', sheetName);
      var sheet = excel[sheetName];

      // 1. Defining Styles with Borders and Colors
      final headerStyle = CellStyle(
        bold: true,
        fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
        backgroundColorHex: ExcelColor.fromHexString('#4472C4'),
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
        leftBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
        rightBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
        topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
        bottomBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
      );

      final dataStyle = CellStyle(
        leftBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
        rightBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
        topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
        bottomBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
      );

      final formulaStyle = CellStyle(
        bold: true,
        backgroundColorHex: ExcelColor.fromHexString('#E2EFDA'),
        leftBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
        rightBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
        topBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
        bottomBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
      );

      final multiColorStyle = CellStyle(
        bold: true,
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
        leftBorder: Border(
            borderStyle: BorderStyle.Thick,
            borderColorHex: ExcelColor.fromHexString('#00FF00')), // Green
        rightBorder: Border(
            borderStyle: BorderStyle.Thick,
            borderColorHex: ExcelColor.fromHexString('#FFFF00')), // Yellow
        topBorder: Border(
            borderStyle: BorderStyle.Thick,
            borderColorHex: ExcelColor.fromHexString('#FF0000')), // Red
        bottomBorder: Border(
            borderStyle: BorderStyle.Thick,
            borderColorHex: ExcelColor.fromHexString('#0000FF')), // Blue
      );

      // 2. Headers
      sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Month"), cellStyle: headerStyle);
      sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Revenue"), cellStyle: headerStyle);
      sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Expenses"), cellStyle: headerStyle);
      sheet.updateCell(CellIndex.indexByString("D1"), TextCellValue("Profit"), cellStyle: headerStyle);

      // 3. Data and Formulas
      final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
      final revenues = [1000, 1200, 1500, 1300, 1700, 2000];
      final expenses = [800, 900, 1000, 950, 1100, 1200];

      for (var i = 0; i < months.length; i++) {
        var row = i + 2;
        sheet.updateCell(CellIndex.indexByString("A$row"), TextCellValue(months[i]), cellStyle: dataStyle);
        sheet.updateCell(CellIndex.indexByString("B$row"), IntCellValue(revenues[i]), cellStyle: dataStyle);
        sheet.updateCell(CellIndex.indexByString("C$row"), IntCellValue(expenses[i]), cellStyle: dataStyle);
        // Formula: Revenue - Expenses
        sheet.updateCell(CellIndex.indexByString("D$row"), FormulaCellValue("B$row-C$row"), cellStyle: dataStyle);
      }

      // 4. Totals with Formulas
      var totalRow = months.length + 2;
      sheet.updateCell(CellIndex.indexByString("A$totalRow"), TextCellValue("TOTAL"), cellStyle: formulaStyle);
      sheet.updateCell(CellIndex.indexByString("B$totalRow"), FormulaCellValue("SUM(B2:B${totalRow - 1})"), cellStyle: formulaStyle);
      sheet.updateCell(CellIndex.indexByString("C$totalRow"), FormulaCellValue("SUM(C2:C${totalRow - 1})"), cellStyle: formulaStyle);
      sheet.updateCell(CellIndex.indexByString("D$totalRow"), FormulaCellValue("SUM(D2:D${totalRow - 1})"), cellStyle: formulaStyle);

      // 5. Multi-colored Border Demo
      sheet.updateCell(CellIndex.indexByString("A${totalRow + 2}"), 
          TextCellValue("Multi-colored Borders"), 
          cellStyle: multiColorStyle);
      sheet.setColumnWidth(0, 25.0);

      // 5. Chart using the data
      final series = [
        ChartSeries(
          name: "Revenue",
          categoriesRange: "'Full Demo'!\$A\$2:\$A\$7",
          valuesRange: "'Full Demo'!\$B\$2:\$B\$7",
        ),
        ChartSeries(
          name: "Profit",
          categoriesRange: "'Full Demo'!\$A\$2:\$A\$7",
          valuesRange: "'Full Demo'!\$D\$2:\$D\$7",
        ),
      ];

      final chart = ColumnChart(
        title: "Financial Overview",
        series: series,
        anchor: ChartAnchor.at(column: 6, row: 1, width: 10, height: 15),
      );

      sheet.addChart(chart);

      // 6. Save and Download
      if (kIsWeb) {
        final bytes = excel.save(fileName: 'full_demo_example.xlsx');
        if (bytes != null && bytes.isNotEmpty) {
          setState(() {
            _status = '✅ Full Demo Excel generated successfully!\n'
                'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
                'The download should start automatically.\n'
                '📌 Includes: Formulas, Headers, Borders, and Charts.';
          });
        }
      } else {
        var bytes = excel.encode();
        if (bytes == null) {
          setState(() => _status = 'Error: Failed to encode Excel file.');
          return;
        }

        String? outputFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Save Full Demo Excel',
          fileName: 'full_demo_example.xlsx',
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
        );

        if (outputFile != null) {
          final file = File(outputFile);
          await file.writeAsBytes(bytes);
          setState(() {
            _status = '✅ Full Demo Excel saved successfully!\n'
                'Location: $outputFile\n'
                '📌 Includes: Formulas, Headers, Borders, and Charts.';
          });
        }
      }
    } catch (e, stackTrace) {
      setState(() => _status = 'Error: $e');
      if (kDebugMode) {
        print('Error generating full demo: $e');
        print(stackTrace);
      }
    } finally {
      setState(() => _isGenerating = false);
    }
  }


  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: Text(widget.title),
      ),
      body: Padding(
        padding: const EdgeInsets.all(20.0),
        child: Center(
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: <Widget>[
              const Icon(Icons.table_chart, size: 80, color: Colors.blue),
              const SizedBox(height: 20),
              Text(
                'Excel Charts Demo',
                style: Theme.of(context).textTheme.headlineMedium,
              ),
              const SizedBox(height: 10),
              Text(
                _status,
                textAlign: TextAlign.center,
                style: const TextStyle(fontSize: 16),
              ),
              const SizedBox(height: 40),
              if (_isGenerating)
                const CircularProgressIndicator()
              else ...[
                ElevatedButton.icon(
                  onPressed: _generateSimpleExcel,
                  icon: const Icon(Icons.table_view),
                  label: const Text('Test: Simple Excel (No Chart)'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.green.shade100,
                  ),
                ),
                const SizedBox(height: 20),
                const Text(
                  'Excel with Charts:',
                  style: TextStyle(fontWeight: FontWeight.bold),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.column),
                  icon: const Icon(Icons.bar_chart),
                  label: const Text('Generate Column Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.line),
                  icon: const Icon(Icons.show_chart),
                  label: const Text('Generate Line Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.pie),
                  icon: const Icon(Icons.pie_chart),
                  label: const Text('Generate Pie Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.area),
                  icon: const Icon(Icons.area_chart),
                  label: const Text('Generate Area Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.doughnut),
                  icon: const Icon(Icons.donut_small),
                  label: const Text('Generate Doughnut Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.radar),
                  icon: const Icon(Icons.radar),
                  label: const Text('Generate Radar Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.bar),
                  icon: const Icon(Icons.horizontal_split),
                  label: const Text('Generate Bar Chart'),
                ),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: () => _generateExcelWithChart(ChartType.scatter),
                  icon: const Icon(Icons.scatter_plot),
                  label: const Text('Generate Scatter Chart'),
                ),
                const SizedBox(height: 30),
                const Divider(),
                const SizedBox(height: 10),
                ElevatedButton.icon(
                  onPressed: _generateFullExample,
                  icon: const Icon(Icons.star, color: Colors.amber),
                  label: const Text('GENERATE FULL DEMO (Everything)'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Colors.blue.shade50,
                    padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 15),
                    textStyle: const TextStyle(fontWeight: FontWeight.bold),
                  ),
                ),
              ],
            ],
          ),
        ),
      ),
    );
  }
}

enum ChartType { column, line, pie, area, doughnut, radar, bar, scatter }
