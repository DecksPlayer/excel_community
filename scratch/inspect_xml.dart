import 'dart:convert';
import 'dart:io';
import 'package:archive/archive.dart';

void main() {
  final bytes = File('./example/conditional_formatting_output.xlsx').readAsBytesSync();
  final archive = ZipDecoder().decodeBytes(bytes);

  for (final file in archive) {
    if (file.name == 'xl/styles.xml' || file.name == 'xl/worksheets/sheet1.xml') {
      print('=== ${file.name} ===');
      print(utf8.decode(file.content));
      print('\n');
    }
  }
}
