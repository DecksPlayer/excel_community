import 'dart:io';
import 'package:excel_community/excel_community.dart' as ec;
import 'package:excel_plus/excel_plus.dart' as ep;

void main(List<String> arguments) {
  if (arguments.length < 2) {
    print('Usage: dart run run_single.dart [community|plus] [cell_count]');
    exit(1);
  }

  final library = arguments[0];
  final cellsCount = int.parse(arguments[1]);
  final cols = 10;
  final rows = cellsCount ~/ cols;

  final stopwatch = Stopwatch()..start();

  if (library == 'community') {
    final excel = ec.Excel.createExcel();
    final sheet = excel['Sheet1'];

    for (int r = 0; r < rows; r++) {
      for (int c = 0; c < cols; c++) {
        final cell = sheet.cell(
          ec.CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        );
        if (r % 4 == 0) {
          cell.value = ec.TextCellValue('Row $r, Col $c');
        } else if (r % 4 == 1) {
          cell.value = ec.IntCellValue(r * c);
        } else if (r % 4 == 2) {
          cell.value = ec.DoubleCellValue(r * c * 1.5);
        } else {
          cell.value = ec.BoolCellValue(r % 2 == 0);
        }
      }
    }
    final buildTime = stopwatch.elapsedMilliseconds;

    stopwatch.reset();
    stopwatch.start();
    final bytes = excel.encode()!;
    final encodeTime = stopwatch.elapsedMilliseconds;
    final totalTime = buildTime + encodeTime;

    print(
      'RESULT:build:$buildTime|encode:$encodeTime|total:$totalTime|size:${bytes.length}',
    );
  } else if (library == 'plus') {
    final excel = ep.Excel.createExcel();
    final sheet = excel['Sheet1'];

    for (int r = 0; r < rows; r++) {
      for (int c = 0; c < cols; c++) {
        final cell = sheet.cell(
          ep.CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        );
        if (r % 4 == 0) {
          cell.value = ep.TextCellValue('Row $r, Col $c');
        } else if (r % 4 == 1) {
          cell.value = ep.IntCellValue(r * c);
        } else if (r % 4 == 2) {
          cell.value = ep.DoubleCellValue(r * c * 1.5);
        } else {
          cell.value = ep.BoolCellValue(r % 2 == 0);
        }
      }
    }
    final buildTime = stopwatch.elapsedMilliseconds;

    stopwatch.reset();
    stopwatch.start();
    final bytes = excel.encode()!;
    final encodeTime = stopwatch.elapsedMilliseconds;
    final totalTime = buildTime + encodeTime;

    print(
      'RESULT:build:$buildTime|encode:$encodeTime|total:$totalTime|size:${bytes.length}',
    );
  } else {
    print('Unknown library: $library');
    exit(1);
  }
}
