// One measurement for compare.dart, run in a fresh process.
//
//   compare_single <community|plus> cell   <rows>         build with cell().value, then encode
//   compare_single <community|plus> append <rows> [out]   build with appendRow, then encode
//   compare_single <community|plus> read   <file>         decode the file and walk every value
//
// Prints `RESULT:build:<ms>|encode:<ms>|read:<ms>|size:<bytes>|cells:<n>`.
import 'dart:io';

import 'package:excel_community/excel_community.dart' as ec;
import 'package:excel_plus/excel_plus.dart' as ep;

const columns = 10;
const regions = ['North', 'South', 'East', 'West'];

/// Row [i] of the data set: text, ints, doubles and a bool.
List<Object> rowValues(int i) => [
      i, 'Customer $i', regions[i % 4], i % 97, (i % 50) + 0.99,
      (i % 97) * ((i % 50) + 0.99), i % 3 == 0, (i * 7) % 101, 'C-${i % 1000}',
      i % 10 == 0 ? 'priority' : '-',
    ];

ec.CellValue communityValue(Object v) => switch (v) {
      int v => ec.IntCellValue(v),
      double v => ec.DoubleCellValue(v),
      bool v => ec.BoolCellValue(v),
      _ => ec.TextCellValue(v.toString()),
    };

ep.CellValue plusValue(Object v) => switch (v) {
      int v => ep.IntCellValue(v),
      double v => ep.DoubleCellValue(v),
      bool v => ep.BoolCellValue(v),
      _ => ep.TextCellValue(v.toString()),
    };

void main(List<String> args) {
  final library = args[0];
  final mode = args[1];
  final community = library == 'community';
  final sw = Stopwatch()..start();

  if (mode == 'read') {
    final bytes = File(args[2]).readAsBytesSync();
    var cells = 0;
    if (community) {
      for (final row in ec.Excel.decodeBytes(bytes)['Data'].rows) {
        for (final cell in row) {
          if (cell?.value != null) cells++;
        }
      }
    } else {
      for (final row in ep.Excel.decodeBytes(bytes)['Data'].rows) {
        for (final cell in row) {
          if (cell?.value != null) cells++;
        }
      }
    }
    print('RESULT:build:0|encode:0|read:${sw.elapsedMilliseconds}|size:${bytes.length}|cells:$cells');
    return;
  }

  final rows = int.parse(args[2]);
  final List<int> bytes;
  final int build;
  if (community) {
    final excel = ec.Excel.createExcel();
    excel.rename('Sheet1', 'Data');
    final sheet = excel['Data'];
    for (var r = 0; r < rows; r++) {
      final values = rowValues(r).map(communityValue).toList();
      if (mode == 'append') {
        sheet.appendRow(values);
      } else {
        for (var c = 0; c < columns; c++) {
          sheet.cell(ec.CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r)).value = values[c];
        }
      }
    }
    build = sw.elapsedMilliseconds;
    sw.reset();
    bytes = excel.encode()!;
  } else {
    final excel = ep.Excel.createExcel();
    excel.rename('Sheet1', 'Data');
    final sheet = excel['Data'];
    for (var r = 0; r < rows; r++) {
      final values = rowValues(r).map(plusValue).toList();
      if (mode == 'append') {
        sheet.appendRow(values);
      } else {
        for (var c = 0; c < columns; c++) {
          sheet.cell(ep.CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r)).value = values[c];
        }
      }
    }
    build = sw.elapsedMilliseconds;
    sw.reset();
    bytes = excel.encode()!;
  }
  final encode = sw.elapsedMilliseconds;
  if (args.length > 3) File(args[3]).writeAsBytesSync(bytes);
  print('RESULT:build:$build|encode:$encode|read:0|size:${bytes.length}|cells:${rows * columns}');
}
