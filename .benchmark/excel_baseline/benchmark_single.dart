import 'dart:io';
import 'package:excel/excel.dart';

void main(List<String> arguments) {
  if (arguments.isEmpty) {
    print('Usage: dart run benchmark_single.dart [cell_count]');
    exit(1);
  }

  final cellsCount = int.parse(arguments[0]);
  final cols = 10;
  final rows = cellsCount ~/ cols;

  final stopwatch = Stopwatch()..start();

  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];

  for (int r = 0; r < rows; r++) {
    for (int c = 0; c < cols; c++) {
      final cell = sheet.cell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
      );
      if (r % 4 == 0) {
        cell.value = TextCellValue('Row $r, Col $c');
      } else if (r % 4 == 1) {
        cell.value = IntCellValue(r * c);
      } else if (r % 4 == 2) {
        cell.value = DoubleCellValue(r * c * 1.5);
      } else {
        cell.value = BoolCellValue(r % 2 == 0);
      }
    }
  }
  final buildTime = stopwatch.elapsedMilliseconds;
  
  stopwatch.reset();
  stopwatch.start();
  final bytes = excel.encode()!;
  final encodeTime = stopwatch.elapsedMilliseconds;
  final totalTime = buildTime + encodeTime;

  print('RESULT:build:$buildTime|encode:$encodeTime|total:$totalTime|size:${bytes.length}');
}
