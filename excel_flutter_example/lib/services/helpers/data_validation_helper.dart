import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

import '../../data/data_validation_samples.dart';

/// Demo workbook: a "Tasks" input sheet with one validated column per rule
/// and the "Lookup Lists" sheet used by the range-based dropdown.
Excel buildDataValidationWorkbook() {
  final excel = Excel.createExcel();
  final tasks = excel['Tasks'];

  // Column per sample, in the ranges used by the snippets.
  final columns = <String, String>{
    'Task': 'A',
    'Status': 'B',
    'Priority': 'C',
    'Hours': 'D',
    'Due date': 'E',
    'Start time': 'F',
    'Code': 'G',
    'Billable': 'H',
    'Grade': 'I',
    'Discount': 'J',
    'Overtime': 'K',
  };
  final headerStyle = CellStyle(
    bold: true,
    fontColorHex: ExcelColor.white,
    backgroundColorHex: ExcelColor.fromHexString('#0D9488'),
  );
  var col = 0;
  for (final header in columns.keys) {
    tasks.updateCell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0), TextCellValue(header),
        cellStyle: headerStyle);
    tasks.setColumnWidth(col, 16);
    col++;
  }

  DataValidation rule(String title) => validationSamples.firstWhere((s) => s.title == title).rule;
  tasks.addDataValidation('A2:A100', rule('Input Message Only'));
  tasks.addDataValidation('B2:B100', rule('Fixed List with Messages'));
  tasks.addDataValidation('C2:C100', rule('List from Another Sheet'));
  tasks.addDataValidation('D2:D100', rule('Whole Number Between'));
  tasks.addDataValidation('E2:E100', rule('Date in 2026'));
  tasks.addDataValidation('F2:F100', rule('Office Hours'));
  tasks.addDataValidation('G2:G100', rule('Text Length'));
  tasks.addDataValidation('H2:H100', rule('Required Yes / No'));
  tasks.addDataValidation('I2:I100', rule('List without Arrow'));
  tasks.addDataValidation('J2:J100', rule('Decimal with Warning'));
  tasks.addDataValidation('K2:K100', rule('Custom Formula'));

  tasks.appendRow([
    TextCellValue('Write release notes'),
    TextCellValue('In progress'),
    TextCellValue('High'),
    IntCellValue(6),
    DateCellValue(year: 2026, month: 10, day: 15),
    TimeCellValue(hour: 9, minute: 30, second: 0),
    TextCellValue('TSK-001'),
    TextCellValue('Yes'),
    TextCellValue('A'),
    DoubleCellValue(0.1),
    IntCellValue(4),
  ]);

  final lists = excel['Lookup Lists'];
  lists.appendRow([TextCellValue('Priority')]);
  for (final value in ['Low', 'Medium', 'High', 'Critical']) {
    lists.appendRow([TextCellValue(value)]);
  }
  return excel;
}

Future<String> generateDataValidationHelper() async {
  final excel = buildDataValidationWorkbook();
  const summary = 'Tasks (11 validated columns), Lookup Lists';

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'data_validation_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Data Validation workbook generated successfully!\n'
          'Worksheets: $summary\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 File: data_validation_example.xlsx\n'
          '💡 Select a cell in row 2+ to see its dropdown or input message.';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    final bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    final outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Data Validation Example',
      fileName: 'data_validation_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Data Validation workbook saved successfully!\n'
          'Location: $outputFile\n'
          'Worksheets: $summary\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return 'Save cancelled.';
  }
}
