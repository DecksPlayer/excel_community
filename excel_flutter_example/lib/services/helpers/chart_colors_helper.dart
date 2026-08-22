import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateChartColorsHelper() async {
  final excel = Excel.createExcel();

  // ── 1. Column — solid custom colors ─────────────────────────────────────
  final col = excel['Column - Solid'];
  _populate(col);
  col.addChart(ColumnChart(
    title: 'Revenue vs Cost (Solid Colors)',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"'Column - Solid'!$A$2:$A$7",
        valuesRange: r"'Column - Solid'!$B$2:$B$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('2E86AB'),
          fillType: ChartFillType.solid,
        ),
      ),
      ChartSeries(
        name: 'Cost',
        categoriesRange: r"'Column - Solid'!$A$2:$A$7",
        valuesRange: r"'Column - Solid'!$C$2:$C$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('E84855'),
          fillType: ChartFillType.solid,
        ),
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
  ));

  // ── 2. Column — transparent (40 % opacity) ───────────────────────────────
  final colT = excel['Column - Transparent'];
  _populate(colT);
  colT.addChart(ColumnChart(
    title: 'Revenue vs Cost (Transparent Bars)',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"'Column - Transparent'!$A$2:$A$7",
        valuesRange: r"'Column - Transparent'!$B$2:$B$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('2E86AB'),
          fillType: ChartFillType.transparent,
          fillAlpha: 40,
          borderColor: ExcelColor.fromHexString('1A5276'),
        ),
      ),
      ChartSeries(
        name: 'Cost',
        categoriesRange: r"'Column - Transparent'!$A$2:$A$7",
        valuesRange: r"'Column - Transparent'!$C$2:$C$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('E84855'),
          fillType: ChartFillType.transparent,
          fillAlpha: 40,
          borderColor: ExcelColor.fromHexString('922B21'),
        ),
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
  ));

  // ── 3. Area — layered transparency ──────────────────────────────────────
  final area = excel['Area - Transparent'];
  _populate(area);
  area.addChart(AreaChart(
    title: 'Revenue vs Cost (Transparent Area)',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"'Area - Transparent'!$A$2:$A$7",
        valuesRange: r"'Area - Transparent'!$B$2:$B$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('1ABC9C'),
          fillType: ChartFillType.transparent,
          fillAlpha: 55,
          borderColor: ExcelColor.fromHexString('148F77'),
          borderAlpha: 90,
          borderWidth: '28575',
        ),
      ),
      ChartSeries(
        name: 'Cost',
        categoriesRange: r"'Area - Transparent'!$A$2:$A$7",
        valuesRange: r"'Area - Transparent'!$C$2:$C$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('9B59B6'),
          fillType: ChartFillType.transparent,
          fillAlpha: 55,
          borderColor: ExcelColor.fromHexString('7D3C98'),
          borderAlpha: 90,
          borderWidth: '28575',
        ),
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
  ));

  // ── 4. Line — no fill, custom line color ────────────────────────────────
  final line = excel['Line - No Fill'];
  _populate(line);
  line.addChart(LineChart(
    title: 'Revenue vs Cost (Line, No Fill)',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"'Line - No Fill'!$A$2:$A$7",
        valuesRange: r"'Line - No Fill'!$B$2:$B$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('F39C12'),
          fillType: ChartFillType.none,
          borderColor: ExcelColor.fromHexString('F39C12'),
          borderWidth: '28575',
        ),
      ),
      ChartSeries(
        name: 'Cost',
        categoriesRange: r"'Line - No Fill'!$A$2:$A$7",
        valuesRange: r"'Line - No Fill'!$C$2:$C$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('2980B9'),
          fillType: ChartFillType.none,
          borderColor: ExcelColor.fromHexString('2980B9'),
          borderWidth: '28575',
        ),
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
  ));

  excel.delete('Sheet1');

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'charts_custom_colors.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Charts with Custom Colors generated!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          '\n4 sheets:\n'
          '• Column - Solid     → custom solid hex colors\n'
          '• Column - Transparent → 40 % opacity bars\n'
          '• Area - Transparent → layered 55 % opacity areas\n'
          '• Line - No Fill     → line-only, no area fill';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) throw Exception('Failed to encode Excel file.');

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Charts with Custom Colors',
      fileName: 'charts_custom_colors.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      await File(outputFile).writeAsBytes(bytes);
      return '✅ Saved successfully!\n'
          'Location: $outputFile\n'
          '\n4 sheets:\n'
          '• Column - Solid     → custom solid hex colors\n'
          '• Column - Transparent → 40 % opacity bars\n'
          '• Area - Transparent → layered 55 % opacity areas\n'
          '• Line - No Fill     → line-only, no area fill';
    }
    return 'Save cancelled.';
  }
}

void _populate(Sheet sheet) {
  sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Month'));
  sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Revenue'));
  sheet.updateCell(CellIndex.indexByString('C1'), TextCellValue('Cost'));
  final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  final revenue = [42000, 55000, 38000, 71000, 64000, 80000];
  final cost = [30000, 32000, 28000, 45000, 40000, 52000];
  for (var i = 0; i < months.length; i++) {
    sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
        TextCellValue(months[i]));
    sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
        IntCellValue(revenue[i]));
    sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1),
        IntCellValue(cost[i]));
  }
}
