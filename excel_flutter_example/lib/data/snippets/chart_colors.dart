// Chart color customization snippets
library;

const String chartColorsSnippet = r'''
import 'package:excel_community/excel_community.dart';

/// Demonstrates ChartSeriesStyle: solid colors, transparency, and no-fill
/// across Column, Line, Area, and Pie charts.
void generateChartsWithCustomColors() {
  final excel = Excel.createExcel();

  // ── 1. Column chart — custom solid colors ─────────────────────────────────
  // Each series gets its own brand color via ChartSeriesStyle(fillType: solid).
  final colSheet = excel['Column - Solid'];
  _populateMonthly(colSheet);

  colSheet.addChart(ColumnChart(
    title: 'Revenue vs Cost (Custom Colors)',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"'Column - Solid'!$A$2:$A$7",
        valuesRange: r"'Column - Solid'!$B$2:$B$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('2E86AB'), // steel blue
          fillType: ChartFillType.solid,
        ),
      ),
      ChartSeries(
        name: 'Cost',
        categoriesRange: r"'Column - Solid'!$A$2:$A$7",
        valuesRange: r"'Column - Solid'!$C$2:$C$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('E84855'), // crimson
          fillType: ChartFillType.solid,
        ),
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
  ));

  // ── 2. Column chart — transparent fills ───────────────────────────────────
  // Same data but each bar is 40 % opaque so overlapping bars remain readable.
  final colTransSheet = excel['Column - Transparent'];
  _populateMonthly(colTransSheet);

  colTransSheet.addChart(ColumnChart(
    title: 'Revenue vs Cost (Transparent)',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"'Column - Transparent'!$A$2:$A$7",
        valuesRange: r"'Column - Transparent'!$B$2:$B$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('2E86AB'),
          fillType: ChartFillType.transparent,
          fillAlpha: 40,                          // 40 % opaque
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

  // ── 3. Area chart — layered transparency ──────────────────────────────────
  // Semi-transparent fills let overlapping areas show through each other.
  final areaSheet = excel['Area - Transparent'];
  _populateMonthly(areaSheet);

  areaSheet.addChart(AreaChart(
    title: 'Revenue vs Cost (Area Transparent)',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"'Area - Transparent'!$A$2:$A$7",
        valuesRange: r"'Area - Transparent'!$B$2:$B$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('1ABC9C'), // emerald
          fillType: ChartFillType.transparent,
          fillAlpha: 55,
          borderColor: ExcelColor.fromHexString('148F77'),
          borderAlpha: 90,
          borderWidth: '28575', // 2.25 pt
        ),
      ),
      ChartSeries(
        name: 'Cost',
        categoriesRange: r"'Area - Transparent'!$A$2:$A$7",
        valuesRange: r"'Area - Transparent'!$C$2:$C$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('9B59B6'), // amethyst
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

  // ── 4. Line chart — no fill + custom line color ───────────────────────────
  // ChartFillType.none removes any fill; only the line and marker are visible.
  final lineSheet = excel['Line - No Fill'];
  _populateMonthly(lineSheet);

  lineSheet.addChart(LineChart(
    title: 'Revenue vs Cost (Line, No Fill)',
    series: [
      ChartSeries(
        name: 'Revenue',
        categoriesRange: r"'Line - No Fill'!$A$2:$A$7",
        valuesRange: r"'Line - No Fill'!$B$2:$B$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('F39C12'), // orange
          fillType: ChartFillType.none,
          borderColor: ExcelColor.fromHexString('F39C12'),
          borderWidth: '28575', // thick line
        ),
      ),
      ChartSeries(
        name: 'Cost',
        categoriesRange: r"'Line - No Fill'!$A$2:$A$7",
        valuesRange: r"'Line - No Fill'!$C$2:$C$7",
        style: ChartSeriesStyle(
          fillColor: ExcelColor.fromHexString('2980B9'), // peter river
          fillType: ChartFillType.none,
          borderColor: ExcelColor.fromHexString('2980B9'),
          borderWidth: '28575',
        ),
      ),
    ],
    anchor: ChartAnchor.at(column: 5, row: 1, width: 11, height: 15),
  ));

  excel.delete('Sheet1');
  excel.save(fileName: 'charts_custom_colors.xlsx');
}

void _populateMonthly(Sheet sheet) {
  sheet.updateCell(CellIndex.indexByString('A1'), TextCellValue('Month'));
  sheet.updateCell(CellIndex.indexByString('B1'), TextCellValue('Revenue'));
  sheet.updateCell(CellIndex.indexByString('C1'), TextCellValue('Cost'));
  final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  final revenue = [42000, 55000, 38000, 71000, 64000, 80000];
  final cost = [30000, 32000, 28000, 45000, 40000, 52000];
  for (var i = 0; i < months.length; i++) {
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
        TextCellValue(months[i]));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
        IntCellValue(revenue[i]));
    sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1),
        IntCellValue(cost[i]));
  }
}
''';
