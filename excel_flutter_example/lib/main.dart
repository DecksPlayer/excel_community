import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:excel_community/excel.dart';
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

  Future<void> _generateExcelWithChart(ChartType type) async {
    setState(() {
      _isGenerating = true;
      _status = 'Generating Excel...';
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
          categoriesRange: "Sheet1!\$A\$2:\$A\$7",
          valuesRange: "Sheet1!\$B\$2:\$B\$7",
        ),
        ChartSeries(
          name: "Series 2",
          categoriesRange: "Sheet1!\$A\$2:\$A\$7",
          valuesRange: "Sheet1!\$C\$2:\$C\$7",
        ),
      ];

      final anchor = ChartAnchor.at(column: 5, row: 1, width: 10, height: 15);

      if (type == ChartType.column) {
        chart = ColumnChart(
          title: "Monthly Data (Column)",
          series: series,
          anchor: anchor,
        );
      } else if (type == ChartType.line) {
        chart = LineChart(
          title: "Monthly Data (Line)",
          series: series,
          anchor: anchor,
        );
      } else {
        chart = PieChart(
          title: "Pie Chart Example",
          series: [series[0]], // Pie chart usually only takes one series
          anchor: anchor,
        );
      }

      sheet.addChart(chart);

      var bytes = excel.save();

      if (kIsWeb) {
        // Handle web saving if needed, for now just status
        setState(() {
          _status = 'Excel generated (Web saving not implemented in this example).';
        });
      } else {
        String? outputFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Save Excel File',
          fileName: 'chart_example.xlsx',
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
        );

        if (outputFile != null) {
          final file = File(outputFile);
          await file.writeAsBytes(bytes!);
          setState(() {
            _status = 'Excel saved to: $outputFile';
          });
        } else {
          setState(() {
            _status = 'Save cancelled.';
          });
        }
      }
    } catch (e) {
      setState(() {
        _status = 'Error: $e';
      });
      if (kDebugMode) {
        print(e);
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
              ],
            ],
          ),
        ),
      ),
    );
  }
}

enum ChartType { column, line, pie }
