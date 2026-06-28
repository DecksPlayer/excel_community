import 'dart:convert';
import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';
import '../models/section_detail.dart';

class ExcelGenerator {
  static void _populateData(Sheet sheet) {
    sheet.updateCell(
      CellIndex.indexByString("A1"),
      TextCellValue("Category"),
    );
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Value 1"));
    sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Value 2"));

    final data = [
      ['Jan', 10, 15],
      ['Feb', 20, 25],
      ['Mar', 15, 30],
      ['Apr', 25, 20],
      ['May', 30, 35],
      ['Jun', 20, 40],
    ];

    for (var i = 0; i < data.length; i++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1),
        TextCellValue(data[i][0] as String),
      );
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1),
        IntCellValue(data[i][1] as int),
      );
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1),
        IntCellValue(data[i][2] as int),
      );
    }
  }

  static Future<String> generateExcelWithImage() async {
    var excel = Excel.createExcel();
    var sheet = excel['Images Demo'];
    excel.delete('Sheet1');

    sheet.updateCell(
      CellIndex.indexByString("A1"),
      TextCellValue("Excel Community — Image Embedding Demo"),
      cellStyle: CellStyle(bold: true, fontSize: 14, fontColorHex: ExcelColor.blue),
    );
    sheet.merge(CellIndex.indexByString("A1"), CellIndex.indexByString("F1"));

    sheet.updateCell(
      CellIndex.indexByString("A3"),
      TextCellValue("The image below is embedded dynamically as a PNG into the worksheet cells:"),
    );

    const greenSquareBase64 = 
        "iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAALElEQVR42u3PMQEAAAwCoNm/tDP0gR1IwMCRAgECBAgQIECAwMCRAgECBAgQIEBg4ECBC3sO/8EAAAAASUVORK5CYII=";
    final imageBytes = base64Decode(greenSquareBase64);

    sheet.addImage(ExcelImage(
      imageBytes: imageBytes,
      imageType: ExcelImageType.png,
      anchor: ImageAnchor.fromPixels(
        column: 1, row: 4, // B5
        widthPixels: 150,
        heightPixels: 150,
      ),
    ));

    sheet.updateCell(
      CellIndex.indexByString("B11"),
      TextCellValue("Green square image embedded above at B5"),
      cellStyle: CellStyle(italic: true, fontSize: 10),
    );

    if (kIsWeb) {
      final bytes = excel.save(fileName: 'image_embedding_example.xlsx');
      if (bytes != null && bytes.isNotEmpty) {
        return '✅ Excel with Image generated successfully!\n'
            'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
            'The download should start automatically.';
      }
      throw Exception('Failed to generate Excel file for Web.');
    } else {
      var bytes = excel.encode();
      if (bytes == null) {
        throw Exception('Failed to encode Excel file.');
      }

      String? outputFile = await FilePicker.platform.saveFile(
        dialogTitle: 'Save Excel with Image',
        fileName: 'image_embedding_example.xlsx',
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (outputFile != null) {
        final file = File(outputFile);
        await file.writeAsBytes(bytes);
        return '✅ Excel with Image saved successfully!\n'
            'Location: $outputFile';
      }
      return 'Save cancelled.';
    }
  }

  static Future<String> generateSimpleExcel() async {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];

    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Name"));
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Age"));
    sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Alice"));
    sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(30));
    sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Bob"));
    sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(25));

    if (kIsWeb) {
      final bytes = excel.save(fileName: 'simple_no_chart.xlsx');
      if (bytes != null && bytes.isNotEmpty) {
        return '✅ Simple Excel generated successfully!\n'
            'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
            'The download should start automatically.\n'
            '\n📥 Check your Downloads folder\n'
            '📌 File: simple_no_chart.xlsx';
      }
      throw Exception('Failed to generate Excel file for Web.');
    } else {
      var bytes = excel.encode();
      if (bytes == null) {
        throw Exception('Failed to encode Excel file.');
      }

      String? outputFile = await FilePicker.platform.saveFile(
        dialogTitle: 'Save Simple Excel File',
        fileName: 'simple_no_chart.xlsx',
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (outputFile != null) {
        final file = File(outputFile);
        await file.writeAsBytes(bytes);
        final savedFileSize = await file.length();
        return '✅ Simple Excel saved successfully!\n'
            'Location: $outputFile\n'
            'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
            '\n📌 Open with Excel to verify it works';
      }
      return 'Save cancelled.';
    }
  }

  static Future<String> generateExcelWithChart(ChartType type) async {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];

    _populateData(sheet);

    Chart chart;
    final series = [
      ChartSeries(
        name: "Series 1",
        categoriesRange: r"Sheet1!$A$2:$A$7",
        valuesRange: r"Sheet1!$B$2:$B$7",
      ),
      ChartSeries(
        name: "Series 2",
        categoriesRange: r"Sheet1!$A$2:$A$7",
        valuesRange: r"Sheet1!$C$2:$C$7",
      ),
    ];

    final anchor = ChartAnchor.at(column: 5, row: 1, width: 10, height: 15);

    switch (type) {
      case ChartType.column:
        chart = ColumnChart(
          title: "Monthly Data (Column)",
          series: series,
          anchor: anchor,
        );
      case ChartType.bar:
        chart = BarChart(
          title: "Monthly Data (Bar)",
          series: series,
          anchor: anchor,
        );
      case ChartType.line:
        chart = LineChart(
          title: "Monthly Data (Line)",
          series: series,
          anchor: anchor,
        );
      case ChartType.area:
        chart = AreaChart(
          title: "Monthly Data (Area)",
          series: series,
          anchor: anchor,
        );
      case ChartType.pie:
        chart = PieChart(
          title: "Pie Chart Example",
          series: [series[0]],
          anchor: anchor,
        );
      case ChartType.doughnut:
        chart = DoughnutChart(
          title: "Doughnut Chart Example",
          series: [series[0]],
          anchor: anchor,
        );
      case ChartType.radar:
        chart = RadarChart(
          title: "Radar Chart Example",
          series: series,
          anchor: anchor,
          filled: true,
        );
      case ChartType.scatter:
        chart = ScatterChart(
          title: "Scatter Chart Example",
          series: series,
          anchor: anchor,
        );
    }

    sheet.addChart(chart);

    if (kIsWeb) {
      final bytes = excel.save(fileName: 'chart_example.xlsx');
      if (bytes != null && bytes.isNotEmpty) {
        return '✅ Excel generated successfully!\n'
            'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
            'The download should start automatically.\n'
            '\n📥 Check your Downloads folder\n'
            '📌 File: chart_example.xlsx';
      }
      throw Exception('Failed to generate Excel file for Web.');
    } else {
      var bytes = excel.encode();
      if (bytes == null) {
        throw Exception('Failed to encode Excel file.');
      }

      String? outputFile = await FilePicker.platform.saveFile(
        dialogTitle: 'Save Excel File',
        fileName: 'chart_example.xlsx',
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (outputFile != null) {
        final file = File(outputFile);
        await file.writeAsBytes(bytes);
        final savedFileSize = await file.length();
        return '✅ Excel saved successfully!\n'
            'Location: $outputFile\n'
            'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB\n'
            '\n📌 IMPORTANT: Open with Excel, not a text editor!';
      }
      return 'Save cancelled.';
    }
  }

  static Future<String> generateUnderlineStyles() async {
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

  static Future<String> generateNumberFormats() async {
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

  static Future<String> generateFullDemo() async {
    var excel = Excel.createExcel();
    var sheetName = 'Full Demo';
    excel.rename('Sheet1', sheetName);
    var sheet = excel[sheetName];

    final headerStyle = CellStyle(
      bold: true,
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
      backgroundColorHex: ExcelColor.fromHexString('#4472C4'),
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
      rightBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
      topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
      bottomBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#000000')),
    );

    final dataStyle = CellStyle(
      leftBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
      rightBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
      topBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
      bottomBorder: Border(borderStyle: BorderStyle.Thin, borderColorHex: ExcelColor.fromHexString('#D9D9D9')),
    );

    final formulaStyle = CellStyle(
      bold: true,
      backgroundColorHex: ExcelColor.fromHexString('#E2EFDA'),
      leftBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
      rightBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
      topBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
      bottomBorder: Border(borderStyle: BorderStyle.Medium, borderColorHex: ExcelColor.fromHexString('#000000')),
    );

    final multiColorStyle = CellStyle(
      bold: true,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: Border(borderStyle: BorderStyle.Thick, borderColorHex: ExcelColor.fromHexString('#00FF00')),
      rightBorder: Border(borderStyle: BorderStyle.Thick, borderColorHex: ExcelColor.fromHexString('#FFFF00')),
      topBorder: Border(borderStyle: BorderStyle.Thick, borderColorHex: ExcelColor.fromHexString('#FF0000')),
      bottomBorder: Border(borderStyle: BorderStyle.Thick, borderColorHex: ExcelColor.fromHexString('#0000FF')),
    );

    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Month"), cellStyle: headerStyle);
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Revenue"), cellStyle: headerStyle);
    sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Expenses"), cellStyle: headerStyle);
    sheet.updateCell(CellIndex.indexByString("D1"), TextCellValue("Profit"), cellStyle: headerStyle);

    final months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
    final revenues = [1000, 1200, 1500, 1300, 1700, 2000];
    final expenses = [800, 900, 1000, 950, 1100, 1200];

    for (var i = 0; i < months.length; i++) {
      var row = i + 2;
      sheet.updateCell(CellIndex.indexByString("A$row"), TextCellValue(months[i]), cellStyle: dataStyle);
      sheet.updateCell(CellIndex.indexByString("B$row"), IntCellValue(revenues[i]), cellStyle: dataStyle);
      sheet.updateCell(CellIndex.indexByString("C$row"), IntCellValue(expenses[i]), cellStyle: dataStyle);
      sheet.updateCell(CellIndex.indexByString("D$row"), FormulaCellValue("B$row-C$row"), cellStyle: dataStyle);
    }

    var totalRow = months.length + 2;
    sheet.updateCell(CellIndex.indexByString("A$totalRow"), TextCellValue("TOTAL"), cellStyle: formulaStyle);
    sheet.updateCell(CellIndex.indexByString("B$totalRow"), FormulaCellValue("SUM(B2:B${totalRow - 1})"), cellStyle: formulaStyle);
    sheet.updateCell(CellIndex.indexByString("C$totalRow"), FormulaCellValue("SUM(C2:C${totalRow - 1})"), cellStyle: formulaStyle);
    sheet.updateCell(CellIndex.indexByString("D$totalRow"), FormulaCellValue("SUM(D2:D${totalRow - 1})"), cellStyle: formulaStyle);

    sheet.updateCell(CellIndex.indexByString("A${totalRow + 2}"), TextCellValue("Multi-colored Borders"), cellStyle: multiColorStyle);
    sheet.setColumnWidth(0, 25.0);

    final series = [
      ChartSeries(
        name: "Revenue",
        categoriesRange: "'Full Demo'!\$A\$2:\$A\$7",
        valuesRange: "'Full Demo'!\$B\$2:\$B\$7",
      ),
      ChartSeries(
        name: "Profit",
        categoriesRange: "'Full Demo'!\$A\$2:\$A\$7",
        valuesRange: "'Full Demo'!\$D\$2:\$D\$7",
      ),
    ];

    final chart = ColumnChart(
      title: "Financial Overview",
      series: series,
      anchor: ChartAnchor.at(column: 6, row: 1, width: 10, height: 15),
    );

    sheet.addChart(chart);

    if (kIsWeb) {
      final bytes = excel.save(fileName: 'full_demo_example.xlsx');
      if (bytes != null && bytes.isNotEmpty) {
        return '✅ Full Demo Excel generated successfully!\n'
            'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
            'The download should start automatically.\n'
            '📌 Includes: Formulas, Headers, Borders, and Charts.';
      }
      throw Exception('Failed to generate Excel file for Web.');
    } else {
      var bytes = excel.encode();
      if (bytes == null) {
        throw Exception('Failed to encode Excel file.');
      }

      String? outputFile = await FilePicker.platform.saveFile(
        dialogTitle: 'Save Full Demo Excel',
        fileName: 'full_demo_example.xlsx',
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (outputFile != null) {
        final file = File(outputFile);
        await file.writeAsBytes(bytes);
        return '✅ Full Demo Excel saved successfully!\n'
            'Location: $outputFile\n'
            '📌 Includes: Formulas, Headers, Borders, and Charts.';
      }
      return 'Save cancelled.';
    }
  }

  static Future<String> generateAllCharts() async {
    var excel = Excel.createExcel();
    var sheet = excel['All Charts Demo'];
    excel.delete('Sheet1');

    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"));
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Series A"));
    sheet.updateCell(CellIndex.indexByString("C1"), TextCellValue("Series B"));

    final data = [
      ['Q1', 10, 15],
      ['Q2', 25, 20],
      ['Q3', 15, 30],
      ['Q4', 30, 25],
    ];

    for (var i = 0; i < data.length; i++) {
      sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), TextCellValue(data[i][0] as String));
      sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), IntCellValue(data[i][1] as int));
      sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: i + 1), IntCellValue(data[i][2] as int));
    }

    final seriesMulti = [
      ChartSeries(
        name: "Series A",
        categoriesRange: "'All Charts Demo'!\$A\$2:\$A\$5",
        valuesRange: "'All Charts Demo'!\$B\$2:\$B\$5",
      ),
      ChartSeries(
        name: "Series B",
        categoriesRange: "'All Charts Demo'!\$A\$2:\$A\$5",
        valuesRange: "'All Charts Demo'!\$C\$2:\$C\$5",
      ),
    ];
    final seriesSingle = [seriesMulti[0]];

    sheet.addChart(ColumnChart(title: "Column Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 5, row: 1, width: 8, height: 12)));
    sheet.addChart(BarChart(title: "Bar Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 14, row: 1, width: 8, height: 12)));
    sheet.addChart(LineChart(title: "Line Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 5, row: 14, width: 8, height: 12)));
    sheet.addChart(AreaChart(title: "Area Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 14, row: 14, width: 8, height: 12)));
    sheet.addChart(PieChart(title: "Pie Chart", series: seriesSingle, anchor: ChartAnchor.at(column: 5, row: 27, width: 8, height: 12)));
    sheet.addChart(DoughnutChart(title: "Doughnut Chart", series: seriesSingle, anchor: ChartAnchor.at(column: 14, row: 27, width: 8, height: 12)));
    sheet.addChart(RadarChart(title: "Radar Chart", series: seriesMulti, anchor: ChartAnchor.at(column: 5, row: 40, width: 8, height: 12), filled: true));

    sheet.updateCell(CellIndex.indexByString("E1"), TextCellValue("X Values"));
    sheet.updateCell(CellIndex.indexByString("F1"), TextCellValue("Y Values A"));
    sheet.updateCell(CellIndex.indexByString("G1"), TextCellValue("Y Values B"));

    final scatterData = [
      [1.0, 5.0, 10.0],
      [2.5, 12.0, 8.0],
      [4.0, 18.0, 15.0],
      [5.5, 25.0, 22.0],
      [7.0, 30.0, 28.0],
      [8.5, 45.0, 35.0],
      [10.0, 55.0, 48.0],
    ];

    for (var i = 0; i < scatterData.length; i++) {
      sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: i + 1), DoubleCellValue(scatterData[i][0]));
      sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: i + 1), DoubleCellValue(scatterData[i][1]));
      sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: i + 1), DoubleCellValue(scatterData[i][2]));
    }

    final scatterSeries = [
      ChartSeries(name: "Trend A", categoriesRange: "'All Charts Demo'!\$E\$2:\$E\$8", valuesRange: "'All Charts Demo'!\$F\$2:\$F\$8"),
      ChartSeries(name: "Trend B", categoriesRange: "'All Charts Demo'!\$E\$2:\$E\$8", valuesRange: "'All Charts Demo'!\$G\$2:\$G\$8"),
    ];

    sheet.addChart(ScatterChart(title: "Scatter Chart (XY Relationship)", series: scatterSeries, anchor: ChartAnchor.at(column: 14, row: 40, width: 8, height: 12)));

    if (kIsWeb) {
      final bytes = excel.save(fileName: 'all_charts_demo.xlsx');
      if (bytes != null && bytes.isNotEmpty) {
        return '✅ Excel with ALL charts generated successfully!\n'
            'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
            'The download should start automatically.\n'
            '📌 Includes: Column, Bar, Line, Area, Pie, Doughnut, Radar, Scatter.';
      }
      throw Exception('Failed to generate Excel file for Web.');
    } else {
      var bytes = excel.encode();
      if (bytes != null) {
        String? outputFile = await FilePicker.platform.saveFile(
          dialogTitle: 'Save All Charts Excel',
          fileName: 'all_charts_demo.xlsx',
          type: FileType.custom,
          allowedExtensions: ['xlsx'],
        );
        if (outputFile != null) {
          await File(outputFile).writeAsBytes(bytes);
          return '✅ Excel with ALL charts saved successfully!\n'
              'Location: $outputFile';
        }
      }
      return 'Save cancelled.';
    }
  }

  static Future<String> generateMultiSheets() async {
    var excel = Excel.createExcel();
    
    // Sheet 1: Summary Sheet
    var summarySheet = excel['Summary'];
    excel.delete('Sheet1'); // Remove default

    summarySheet.updateCell(
      CellIndex.indexByString("A1"), 
      TextCellValue("Monthly Financial Summary"),
      cellStyle: CellStyle(bold: true, fontSize: 14, fontColorHex: ExcelColor.blue),
    );

    // Sheet 2: Revenues Detail
    var revenueSheet = excel['Revenues'];
    revenueSheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Source"), cellStyle: CellStyle(bold: true));
    revenueSheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Amount"), cellStyle: CellStyle(bold: true));
    revenueSheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("SaaS Subscriptions"));
    revenueSheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(15400));
    revenueSheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Consulting"));
    revenueSheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(4200));

    // Sheet 3: Expenses Detail
    var expenseSheet = excel['Expenses'];
    expenseSheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Category"), cellStyle: CellStyle(bold: true));
    expenseSheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Cost"), cellStyle: CellStyle(bold: true));
    expenseSheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("Hosting & Cloud"));
    expenseSheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(3100));
    expenseSheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("Marketing"));
    expenseSheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(1200));

    if (kIsWeb) {
      final bytes = excel.save(fileName: 'multisheet_financials.xlsx');
      if (bytes != null && bytes.isNotEmpty) {
        return '✅ Multi-sheet Excel generated successfully!\n'
            'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
            'The download should start automatically.\n'
            '📌 Includes: Summary, Revenues, and Expenses sheets.';
      }
      throw Exception('Failed to generate Excel file for Web.');
    } else {
      var bytes = excel.encode();
      if (bytes == null) {
        throw Exception('Failed to encode Excel file.');
      }

      String? outputFile = await FilePicker.platform.saveFile(
        dialogTitle: 'Save Multi-sheet Excel',
        fileName: 'multisheet_financials.xlsx',
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (outputFile != null) {
        final file = File(outputFile);
        await file.writeAsBytes(bytes);
        return '✅ Multi-sheet Excel saved successfully!\n'
            'Location: $outputFile';
      }
      return 'Save cancelled.';
    }
  }

  static Future<String> generatePivotTemplate() async {
    var excel = Excel.createExcel();
    var sheet = excel['Sales Data'];
    excel.delete('Sheet1');

    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Region"), cellStyle: CellStyle(bold: true));
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Amount"), cellStyle: CellStyle(bold: true));

    sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("North Region"));
    sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(45000));

    sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("South Region"));
    sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(32000));

    sheet.updateCell(CellIndex.indexByString("A4"), TextCellValue("East Region"));
    sheet.updateCell(CellIndex.indexByString("B4"), IntCellValue(51000));

    if (kIsWeb) {
      final bytes = excel.save(fileName: 'sales_report_pivot.xlsx');
      if (bytes != null && bytes.isNotEmpty) {
        return '✅ Template data updated successfully!\n'
            'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
            'Download started. (In a real app, loading a template with pre-configured Pivot Tables preserves them entirely)';
      }
      throw Exception('Failed to generate Excel file for Web.');
    } else {
      var bytes = excel.encode();
      if (bytes == null) {
        throw Exception('Failed to encode Excel file.');
      }

      String? outputFile = await FilePicker.platform.saveFile(
        dialogTitle: 'Save Sales Report Excel',
        fileName: 'sales_report_pivot.xlsx',
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (outputFile != null) {
        final file = File(outputFile);
        await file.writeAsBytes(bytes);
        return '✅ Template data saved successfully!\n'
            'Location: $outputFile';
      }
      return 'Save cancelled.';
    }
  }

  static Future<String> generateLockedCellsReport() async {
    var excel = Excel.createExcel();
    var sheet = excel['Protected Report'];
    excel.delete('Sheet1');

    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Read-Only Header"), cellStyle: CellStyle(bold: true));
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Editable Values"), cellStyle: CellStyle(bold: true));

    sheet.updateCell(CellIndex.indexByString("A2"), TextCellValue("North Sales"));
    sheet.updateCell(CellIndex.indexByString("B2"), IntCellValue(8500));

    sheet.updateCell(CellIndex.indexByString("A3"), TextCellValue("South Sales"));
    sheet.updateCell(CellIndex.indexByString("B3"), IntCellValue(6400));

    if (kIsWeb) {
      final bytes = excel.save(fileName: 'protected_sales_report.xlsx');
      if (bytes != null && bytes.isNotEmpty) {
        return '✅ Locked Cells template data updated successfully!\n'
            'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
            'Download started. (In a real app, loading a template with pre-configured sheet protection preserves cell locking.)';
      }
      throw Exception('Failed to generate Excel file for Web.');
    } else {
      var bytes = excel.encode();
      if (bytes == null) {
        throw Exception('Failed to encode Excel file.');
      }

      String? outputFile = await FilePicker.platform.saveFile(
        dialogTitle: 'Save Protected Report Excel',
        fileName: 'protected_sales_report.xlsx',
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (outputFile != null) {
        final file = File(outputFile);
        await file.writeAsBytes(bytes);
        return '✅ Protected Report saved successfully!\n'
            'Location: $outputFile';
      }
      return 'Save cancelled.';
    }
  }

  static Future<String> generateFreezePanes() async {
    var excel = Excel.createExcel();
    var sheet = excel['Sales Report'];
    excel.delete('Sheet1');

    sheet.updateCell(CellIndex.indexByString("A1"), TextCellValue("Product Name"), cellStyle: CellStyle(bold: true));
    sheet.updateCell(CellIndex.indexByString("B1"), TextCellValue("Sales Revenue"), cellStyle: CellStyle(bold: true));

    for (int i = 2; i <= 50; i++) {
      sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i - 1), TextCellValue("Product $i"));
      sheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i - 1), IntCellValue(150 * i));
    }

    if (kIsWeb) {
      final bytes = excel.save(fileName: 'sales_report_frozen.xlsx');
      if (bytes != null && bytes.isNotEmpty) {
        return '✅ Frozen Panes template data updated successfully!\n'
            'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
            'Download started. (In a real app, loading a template with pre-configured Freeze Panes preserves them.)';
      }
      throw Exception('Failed to generate Excel file for Web.');
    } else {
      var bytes = excel.encode();
      if (bytes == null) {
        throw Exception('Failed to encode Excel file.');
      }

      String? outputFile = await FilePicker.platform.saveFile(
        dialogTitle: 'Save Sales Report Excel',
        fileName: 'sales_report_frozen.xlsx',
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (outputFile != null) {
        final file = File(outputFile);
        await file.writeAsBytes(bytes);
        return '✅ Sales Report saved successfully!\n'
            'Location: $outputFile';
      }
      return 'Save cancelled.';
    }
  }
}
