import 'dart:io';
import 'package:excel_community/excel_community.dart';

void main() {
  final file = './test/test_resources/superscriptExample.xlsx';
  final bytes = File(file).readAsBytesSync();
  try {
    final excel = Excel.decodeBytes(bytes);
    print('Decode succeeded!');
    print('Shared strings count: ${excel.tables["Sheet1"]?.maxRows}');
  } catch (e, stack) {
    print('Decode failed: $e');
    print('Stack: $stack');
  }
}
