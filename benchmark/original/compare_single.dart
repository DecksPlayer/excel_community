// compare_single.dart for the original `excel` package (same arguments,
// data and output as ../compare_single.dart).
//
//   compare_single original <cell|append> <rows> [out]
//   compare_single original read <file>
import 'dart:io';

import 'package:excel/excel.dart';

const columns = 10;
const regions = ['North', 'South', 'East', 'West'];

/// Row [i] of the data set: text, ints, doubles and a bool.
List<Object> rowValues(int i) => [
      i, 'Customer $i', regions[i % 4], i % 97, (i % 50) + 0.99,
      (i % 97) * ((i % 50) + 0.99), i % 3 == 0, (i * 7) % 101, 'C-${i % 1000}',
      i % 10 == 0 ? 'priority' : '-',
    ];

CellValue cellValue(Object v) => switch (v) {
      int v => IntCellValue(v),
      double v => DoubleCellValue(v),
      bool v => BoolCellValue(v),
      _ => TextCellValue(v.toString()),
    };

void main(List<String> args) {
  final mode = args[1];
  final sw = Stopwatch()..start();

  if (mode == 'read') {
    final bytes = File(args[2]).readAsBytesSync();
    var cells = 0;
    for (final row in Excel.decodeBytes(bytes)['Data'].rows) {
      for (final cell in row) {
        if (cell?.value != null) cells++;
      }
    }
    print('RESULT:build:0|encode:0|read:${sw.elapsedMilliseconds}|size:${bytes.length}|cells:$cells');
    return;
  }

  final rows = int.parse(args[2]);
  final excel = Excel.createExcel();
  excel.rename('Sheet1', 'Data');
  final sheet = excel['Data'];
  for (var r = 0; r < rows; r++) {
    final values = rowValues(r).map(cellValue).toList();
    if (mode == 'append') {
      sheet.appendRow(values);
    } else {
      for (var c = 0; c < columns; c++) {
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r)).value = values[c];
      }
    }
  }
  final build = sw.elapsedMilliseconds;
  sw.reset();
  final bytes = excel.encode()!;
  final encode = sw.elapsedMilliseconds;
  if (args.length > 3) File(args[3]).writeAsBytesSync(bytes);
  print('RESULT:build:$build|encode:$encode|read:0|size:${bytes.length}|cells:${rows * columns}');
}
