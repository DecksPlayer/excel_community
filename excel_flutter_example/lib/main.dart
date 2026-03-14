import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
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
              ],
            ],
          ),
        ),
      ),
    );
  }
}

enum ChartType { column, line, pie, area, doughnut, radar, bar, scatter }
