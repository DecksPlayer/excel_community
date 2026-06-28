import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateUnderlineStylesHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Text Styles Demo'];
  excel.delete('Sheet1');

  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('EJEMPLOS DE ESTILOS DE TEXTO'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 16,
      fontColorHex: ExcelColor.blue,
      horizontalAlign: HorizontalAlign.Center,
    ),
  );
  sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('D1'));

  int row = 3;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('UNDERLINE (SUBRAYADO)'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 14,
      backgroundColorHex: ExcelColor.grey200,
    ),
  );
  sheet.merge(
    CellIndex.indexByString('A$row'),
    CellIndex.indexByString('D$row'),
  );
  row += 2;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Sin subrayado:'),
    cellStyle: CellStyle(bold: true),
  );
  sheet.updateCell(
    CellIndex.indexByString('B$row'),
    TextCellValue('Este texto NO tiene subrayado'),
    cellStyle: CellStyle(underline: Underline.None, fontSize: 12),
  );
  row++;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Subrayado Simple:'),
    cellStyle: CellStyle(bold: true),
  );
  sheet.updateCell(
    CellIndex.indexByString('B$row'),
    TextCellValue('Este texto tiene subrayado SIMPLE'),
    cellStyle: CellStyle(
      underline: Underline.Single,
      fontSize: 12,
      fontColorHex: ExcelColor.blue,
    ),
  );
  row++;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Subrayado Doble:'),
    cellStyle: CellStyle(bold: true),
  );
  sheet.updateCell(
    CellIndex.indexByString('B$row'),
    TextCellValue('Este texto tiene subrayado DOBLE'),
    cellStyle: CellStyle(
      underline: Underline.Double,
      fontSize: 12,
      fontColorHex: ExcelColor.red,
    ),
  );
  row += 2;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('COMBINACIONES DE ESTILOS'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 14,
      backgroundColorHex: ExcelColor.grey200,
    ),
  );
  sheet.merge(
    CellIndex.indexByString('A$row'),
    CellIndex.indexByString('D$row'),
  );
  row += 2;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Negrita + Subrayado Simple'),
    cellStyle: CellStyle(
      bold: true,
      underline: Underline.Single,
      fontSize: 12,
    ),
  );
  row++;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Cursiva + Subrayado Doble'),
    cellStyle: CellStyle(
      italic: true,
      underline: Underline.Double,
      fontSize: 12,
    ),
  );
  row++;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Negrita + Cursiva + Subrayado'),
    cellStyle: CellStyle(
      bold: true,
      italic: true,
      underline: Underline.Single,
      fontSize: 12,
      fontColorHex: ExcelColor.purple,
    ),
  );
  row++;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Texto Tachado (Strikethrough)'),
    cellStyle: CellStyle(
      strikethrough: true,
      fontSize: 12,
      fontColorHex: ExcelColor.red,
    ),
  );
  row++;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Tachado + Subrayado'),
    cellStyle: CellStyle(
      strikethrough: true,
      underline: Underline.Single,
      fontSize: 12,
    ),
  );
  row++;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Negrita + Cursiva + Tachado + Subrayado'),
    cellStyle: CellStyle(
      bold: true,
      italic: true,
      strikethrough: true,
      underline: Underline.Double,
      fontSize: 12,
      fontColorHex: ExcelColor.green,
    ),
  );
  row++;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('Fondo Amarillo + Subrayado'),
    cellStyle: CellStyle(
      underline: Underline.Single,
      fontSize: 12,
      backgroundColorHex: ExcelColor.yellow,
      fontColorHex: ExcelColor.black,
    ),
  );
  row += 2;

  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('TABLA CON BORDES Y ESTILOS'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 14,
      backgroundColorHex: ExcelColor.grey200,
    ),
  );
  sheet.merge(
    CellIndex.indexByString('A$row'),
    CellIndex.indexByString('D$row'),
  );
  row += 2;

  final headerStyle = CellStyle(
    bold: true,
    backgroundColorHex: ExcelColor.blue300,
    fontColorHex: ExcelColor.white,
    horizontalAlign: HorizontalAlign.Center,
    leftBorder: Border(borderStyle: BorderStyle.Thin),
    rightBorder: Border(borderStyle: BorderStyle.Thin),
    topBorder: Border(borderStyle: BorderStyle.Thin),
    bottomBorder: Border(borderStyle: BorderStyle.Thin),
  );

  sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue('Producto'), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString('B$row'), TextCellValue('Cantidad'), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString('C$row'), TextCellValue('Precio'), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString('D$row'), TextCellValue('Total'), cellStyle: headerStyle);
  row++;

  final cellBorderStyle = CellStyle(
    leftBorder: Border(borderStyle: BorderStyle.Thin),
    rightBorder: Border(borderStyle: BorderStyle.Thin),
    topBorder: Border(borderStyle: BorderStyle.Thin),
    bottomBorder: Border(borderStyle: BorderStyle.Thin),
  );

  final data = [
    ['Laptop', 2, 1200.50, 2401.00],
    ['Mouse', 5, 25.99, 129.95],
    ['Teclado', 3, 89.90, 269.70],
  ];

  for (var rowData in data) {
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      TextCellValue(rowData[0] as String),
      cellStyle: cellBorderStyle.copyWith(underlineVal: Underline.Single),
    );
    sheet.updateCell(
      CellIndex.indexByString('B$row'),
      IntCellValue(rowData[1] as int),
      cellStyle: cellBorderStyle.copyWith(horizontalAlignVal: HorizontalAlign.Center),
    );
    sheet.updateCell(
      CellIndex.indexByString('C$row'),
      DoubleCellValue(rowData[2] as double),
      cellStyle: cellBorderStyle.copyWith(
        horizontalAlignVal: HorizontalAlign.Right,
        numberFormat: NumFormat.standard_2,
      ),
    );
    sheet.updateCell(
      CellIndex.indexByString('D$row'),
      DoubleCellValue(rowData[3] as double),
      cellStyle: cellBorderStyle.copyWith(
        horizontalAlignVal: HorizontalAlign.Right,
        numberFormat: NumFormat.standard_2,
        boldVal: true,
        fontColorHexVal: ExcelColor.green,
      ),
    );
    row++;
  }

  sheet.setColumnWidth(0, 25.0);
  sheet.setColumnWidth(1, 15.0);
  sheet.setColumnWidth(2, 15.0);
  sheet.setColumnWidth(3, 15.0);

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'underline_styles_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Underline & Styles Example generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 File: underline_styles_example.xlsx';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Underline & Styles Example',
      fileName: 'underline_styles_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Underline & Styles Example saved successfully!\n'
          'Location: $outputFile\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return 'Save cancelled.';
  }
}

Future<String> generateNumberFormatsHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Number Formats'];
  excel.delete('Sheet1');

  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('DEMO DE FORMATOS DE NÚMERO (BUILT-IN IDs)'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 16,
      fontColorHex: ExcelColor.blue,
      horizontalAlign: HorizontalAlign.Center,
    ),
  );
  sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('D1'));

  int row = 3;

  final headerStyle = CellStyle(
    bold: true,
    backgroundColorHex: ExcelColor.blue300,
    fontColorHex: ExcelColor.white,
    horizontalAlign: HorizontalAlign.Center,
    leftBorder: Border(borderStyle: BorderStyle.Thin),
    rightBorder: Border(borderStyle: BorderStyle.Thin),
    topBorder: Border(borderStyle: BorderStyle.Thin),
    bottomBorder: Border(borderStyle: BorderStyle.Thin),
  );

  sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue('ID de Formato'), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString('B$row'), TextCellValue('Nombre / Descripción'), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString('C$row'), TextCellValue('Valor de Ejemplo'), cellStyle: headerStyle);
  sheet.updateCell(CellIndex.indexByString('D$row'), TextCellValue('Formateado'), cellStyle: headerStyle);
  row++;

  final cellBorderStyle = CellStyle(
    leftBorder: Border(borderStyle: BorderStyle.Thin),
    rightBorder: Border(borderStyle: BorderStyle.Thin),
    topBorder: Border(borderStyle: BorderStyle.Thin),
    bottomBorder: Border(borderStyle: BorderStyle.Thin),
  );

  final formats = [
    {
      'id': 0,
      'desc': 'General (Sin formato)',
      'value': DoubleCellValue(1234.567),
      'format': NumFormat.standard_0,
    },
    {
      'id': 1,
      'desc': 'Integer (0)',
      'value': IntCellValue(1234),
      'format': NumFormat.standard_1,
    },
    {
      'id': 2,
      'desc': 'Float (0.00)',
      'value': DoubleCellValue(1234.567),
      'format': NumFormat.standard_2,
    },
    {
      'id': 3,
      'desc': 'Integer con comas (#,##0)',
      'value': IntCellValue(1234567),
      'format': NumFormat.standard_3,
    },
    {
      'id': 4,
      'desc': 'Float con comas (#,##0.00)',
      'value': DoubleCellValue(1234567.89),
      'format': NumFormat.standard_4,
    },
    {
      'id': 9,
      'desc': 'Porcentaje (0%)',
      'value': DoubleCellValue(0.756),
      'format': NumFormat.standard_9,
    },
    {
      'id': 10,
      'desc': 'Porcentaje con decimales (0.00%)',
      'value': DoubleCellValue(0.7563),
      'format': NumFormat.standard_10,
    },
    {
      'id': 11,
      'desc': 'Científico (0.00E+00)',
      'value': DoubleCellValue(123456789),
      'format': NumFormat.standard_11,
    },
    {
      'id': 14,
      'desc': 'Fecha corta (mm-dd-yy)',
      'value': DateCellValue(year: 2026, month: 5, day: 31),
      'format': NumFormat.standard_14,
    },
    {
      'id': 15,
      'desc': 'Fecha larga (d-mmm-yy)',
      'value': DateCellValue(year: 2026, month: 5, day: 31),
      'format': NumFormat.standard_15,
    },
    {
      'id': 18,
      'desc': 'Hora 12h (h:mm AM/PM)',
      'value': DateTimeCellValue(year: 2026, month: 5, day: 31, hour: 14, minute: 30, second: 0),
      'format': NumFormat.standard_18,
    },
    {
      'id': 20,
      'desc': 'Hora 24h (h:mm)',
      'value': DateTimeCellValue(year: 2026, month: 5, day: 31, hour: 14, minute: 30, second: 0),
      'format': NumFormat.standard_20,
    },
    {
      'id': 22,
      'desc': 'Fecha y Hora (m/d/yy h:mm)',
      'value': DateTimeCellValue(year: 2026, month: 5, day: 31, hour: 14, minute: 30, second: 0),
      'format': NumFormat.standard_22,
    },
    {
      'id': 37,
      'desc': 'Contabilidad entera con parént. (#,##0 ;(#,##0))',
      'value': DoubleCellValue(-1234.5),
      'format': NumFormat.standard_37,
    },
    {
      'id': 38,
      'desc': 'Contabilidad entera en rojo (#,##0 ;[Red](#,##0))',
      'value': DoubleCellValue(-1234.5),
      'format': NumFormat.standard_38,
    },
    {
      'id': 39,
      'desc': 'Contabilidad float con parént. (#,##0.00;(#,##0.00))',
      'value': DoubleCellValue(-1234.56),
      'format': NumFormat.standard_39,
    },
    {
      'id': 40,
      'desc': 'Contabilidad float en rojo (#,##0.00;[Red](#,##0.00))',
      'value': DoubleCellValue(-1234.56),
      'format': NumFormat.standard_40,
    },
    {
      'id': 44,
      'desc': 'Contabilidad (Moneda) con sangría (Accounting ID 44!)',
      'value': DoubleCellValue(1234.56),
      'format': NumFormat.standard_44,
    },
  ];

  for (var f in formats) {
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      IntCellValue(f['id'] as int),
      cellStyle: cellBorderStyle.copyWith(horizontalAlignVal: HorizontalAlign.Center),
    );
    sheet.updateCell(
      CellIndex.indexByString('B$row'),
      TextCellValue(f['desc'] as String),
      cellStyle: cellBorderStyle,
    );
    sheet.updateCell(
      CellIndex.indexByString('C$row'),
      TextCellValue(f['value'].toString()),
      cellStyle: cellBorderStyle,
    );
    sheet.updateCell(
      CellIndex.indexByString('D$row'),
      f['value'] as CellValue,
      cellStyle: cellBorderStyle.copyWith(
        numberFormat: f['format'] as NumFormat,
        horizontalAlignVal: HorizontalAlign.Right,
      ),
    );
    row++;
  }

  sheet.setColumnWidth(0, 15.0);
  sheet.setColumnWidth(1, 45.0);
  sheet.setColumnWidth(2, 25.0);
  sheet.setColumnWidth(3, 25.0);

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'number_formats_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Number Formats Example generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 File: number_formats_example.xlsx';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Number Formats Example',
      fileName: 'number_formats_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Number Formats Example saved successfully!\n'
          'Location: $outputFile\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return 'Save cancelled.';
  }
}
