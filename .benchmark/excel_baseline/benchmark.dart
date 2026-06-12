import 'dart:io';
import 'package:excel/excel.dart';

void main(List<String> args) {
  final rows = args.isNotEmpty ? int.parse(args[0]) : 20000;
  final cols = args.length > 1 ? int.parse(args[1]) : 50;
  final cells = rows * cols;

  print('excel_original benchmark — $cells cells ($rows rows x $cols cols)\n');

  final sw = Stopwatch()..start();
  final excel = Excel.createExcel();
  final sheet = excel['Sheet1'];
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r),
        TextCellValue('R${r}C$c'),
      );
    }
  }
  final createUs = sw.elapsedMicroseconds;

  sw.reset();
  final bytes = excel.encode()!;
  final encodeUs = sw.elapsedMicroseconds;
  final fileKb = bytes.length / 1024;

  sw.reset();
  final decoded = Excel.decodeBytes(bytes);
  // Touch a cell to force the lazy per-sheet parse.
  decoded['Sheet1'].cell(CellIndex.indexByString('A1')).value;
  final decodeUs = sw.elapsedMicroseconds;

  final totalUs = createUs + encodeUs + decodeUs;
  final rssMb = ProcessInfo.currentRss / (1024 * 1024);

  print('create:    ${_ms(createUs)}');
  print('encode:    ${_ms(encodeUs)}  (${_size(fileKb)} file)');
  print('decode:    ${_ms(decodeUs)}');
  print('total:     ${_ms(totalUs)}');
  print('peak RSS:  ${rssMb.toStringAsFixed(0)} MB');
}

String _ms(int micros) {
  final ms = micros / 1000;
  return '${ms.toStringAsFixed(ms < 100 ? 2 : 0)} ms';
}

String _size(double kb) => kb >= 1024
    ? '${(kb / 1024).toStringAsFixed(2)} MB'
    : '${kb.toStringAsFixed(1)} KB';
