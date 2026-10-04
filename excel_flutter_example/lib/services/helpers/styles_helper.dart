import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

import '../../data/number_format_catalog.dart';

/// Builds the Number Formats demo workbook from [numFormatCatalog]: one
/// worksheet per category, one row per sample value.
Excel buildNumberFormatsWorkbook() {
  final excel = Excel.createExcel();

  final border = Border(borderStyle: BorderStyle.Thin);
  final cellStyle = CellStyle(
    leftBorder: border,
    rightBorder: border,
    topBorder: border,
    bottomBorder: border,
  );
  final headerStyle = cellStyle.copyWith(
    boldVal: true,
    backgroundColorHexVal: ExcelColor.fromHexString('#0D9488'),
    fontColorHexVal: ExcelColor.white,
    horizontalAlignVal: HorizontalAlign.Center,
  );
  const headers = ['ID', 'Name', 'Format Code', 'Sample', 'Formatted Cell', 'Excel Shows', 'Dart'];

  for (final category in NumFormatCategory.values) {
    final sheet = excel[category.label];
    sheet.updateCell(
      CellIndex.indexByString('A1'),
      TextCellValue('${category.label} - Number Formats'),
      cellStyle: CellStyle(bold: true, fontSize: 16, fontColorHex: ExcelColor.fromHexString('#0D9488')),
    );
    for (var col = 0; col < headers.length; col++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 2),
        TextCellValue(headers[col]),
        cellStyle: headerStyle,
      );
    }

    var row = 3;
    for (final entry in numFormatCatalog.where((e) => e.category == category)) {
      for (final sample in entry.samples) {
        void put(int col, CellValue value, [CellStyle? style]) => sheet.updateCell(
              CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
              value,
              cellStyle: style ?? cellStyle,
            );
        final id = entry.numFmtId;
        put(0, id != null ? IntCellValue(id) : TextCellValue('custom'));
        put(1, TextCellValue(entry.name));
        put(2, TextCellValue(entry.format.formatCode));
        put(3, TextCellValue(sampleLabel(sample.value)));
        put(4, sample.value, cellStyle.copyWith(numberFormat: entry.format, horizontalAlignVal: HorizontalAlign.Right));
        put(5, TextCellValue(sample.expected));
        put(6, TextCellValue(entry.dartExpression));
        row++;
      }
    }

    const widths = [8.0, 30.0, 42.0, 22.0, 26.0, 30.0, 58.0];
    for (var col = 0; col < widths.length; col++) {
      sheet.setColumnWidth(col, widths[col]);
    }
    sheet.frozenRows = 3;
  }
  return excel;
}

Future<String> generateNumberFormatsHelper() async {
  final excel = buildNumberFormatsWorkbook();
  final summary = '${NumFormatCategory.values.length} worksheets, '
      '${numFormatCatalog.length} formats';

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'number_formats_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Number Formats Example generated successfully!\n'
          '$summary\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 File: number_formats_example.xlsx\n'
          '💡 Compare the "Formatted Cell" column with "Excel Shows".';
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
          '$summary\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return 'Save cancelled.';
  }
}

Future<String> generateFontsStylesHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Fonts & Styles Demo'];
  excel.delete('Sheet1');

  // Title block
  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('FONTS & STYLES DEMO'),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 16,
      fontColorHex: ExcelColor.indigo,
      horizontalAlign: HorizontalAlign.Center,
    ),
  );
  sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('E1'));

  int row = 3;

  final groupHeaderStyle = CellStyle(
    bold: true,
    fontSize: 14,
    backgroundColorHex: ExcelColor.grey200,
    fontColorHex: ExcelColor.black,
  );

  // SECTION 1: FONT FAMILIES
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('1. FONT FAMILIES'),
    cellStyle: groupHeaderStyle,
  );
  sheet.merge(CellIndex.indexByString('A$row'), CellIndex.indexByString('E$row'));
  row += 2;

  final tableHeaderStyle = CellStyle(
    bold: true,
    backgroundColorHex: ExcelColor.indigo300,
    fontColorHex: ExcelColor.white,
    horizontalAlign: HorizontalAlign.Center,
    leftBorder: Border(borderStyle: BorderStyle.Thin),
    rightBorder: Border(borderStyle: BorderStyle.Thin),
    topBorder: Border(borderStyle: BorderStyle.Thin),
    bottomBorder: Border(borderStyle: BorderStyle.Thin),
  );

  sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue('Font Family'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('B$row'), TextCellValue('Technical Name'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('C$row'), TextCellValue('Dart Code'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('D$row'), TextCellValue('Sample Text'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('E$row'), TextCellValue('Stylized Sample'), cellStyle: tableHeaderStyle);
  row++;

  final cellBorderStyle = CellStyle(
    leftBorder: Border(borderStyle: BorderStyle.Thin),
    rightBorder: Border(borderStyle: BorderStyle.Thin),
    topBorder: Border(borderStyle: BorderStyle.Thin),
    bottomBorder: Border(borderStyle: BorderStyle.Thin),
  );

  final List<Map<String, dynamic>> fontFamilies = [
    {
      'name': 'Calibri',
      'enum': 'FontFamily.Calibri',
      'code': "fontFamily: 'Calibri'",
      'family': 'Calibri',
    },
    {
      'name': 'Arial',
      'enum': 'FontFamily.Arial',
      'code': "fontFamily: getFontFamily(FontFamily.Arial)",
      'family': getFontFamily(FontFamily.Arial),
    },
    {
      'name': 'Comic Sans MS',
      'enum': 'FontFamily.Comic_Sans_MS',
      'code': "fontFamily: getFontFamily(FontFamily.Comic_Sans_MS)",
      'family': getFontFamily(FontFamily.Comic_Sans_MS),
    },
    {
      'name': 'Consolas',
      'enum': 'FontFamily.Consolas',
      'code': "fontFamily: getFontFamily(FontFamily.Consolas)",
      'family': getFontFamily(FontFamily.Consolas),
    },
    {
      'name': 'Courier New',
      'enum': 'FontFamily.Courier_New',
      'code': "fontFamily: getFontFamily(FontFamily.Courier_New)",
      'family': getFontFamily(FontFamily.Courier_New),
    },
    {
      'name': 'Georgia',
      'enum': 'FontFamily.Georgia',
      'code': "fontFamily: getFontFamily(FontFamily.Georgia)",
      'family': getFontFamily(FontFamily.Georgia),
    },
    {
      'name': 'Impact',
      'enum': 'FontFamily.Impact',
      'code': "fontFamily: getFontFamily(FontFamily.Impact)",
      'family': getFontFamily(FontFamily.Impact),
    },
    {
      'name': 'Lucida Console',
      'enum': 'FontFamily.Lucida_Console',
      'code': "fontFamily: getFontFamily(FontFamily.Lucida_Console)",
      'family': getFontFamily(FontFamily.Lucida_Console),
    },
  ];

  for (var font in fontFamilies) {
    final family = font['family'] as String;
    sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue(font['name'] as String), cellStyle: cellBorderStyle);
    sheet.updateCell(CellIndex.indexByString('B$row'), TextCellValue(font['enum'] as String), cellStyle: cellBorderStyle);
    sheet.updateCell(CellIndex.indexByString('C$row'), TextCellValue(font['code'] as String), cellStyle: cellBorderStyle);
    sheet.updateCell(CellIndex.indexByString('D$row'), TextCellValue('The quick brown fox jumps...'), cellStyle: cellBorderStyle);
    sheet.updateCell(
      CellIndex.indexByString('E$row'),
      TextCellValue('The quick brown fox jumps over the lazy dog.'),
      cellStyle: cellBorderStyle.copyWith(
        fontFamilyVal: family,
        fontSizeVal: 11,
      ),
    );
    row++;
  }
  row += 2;

  // SECTION 2: FONT SIZES
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('2. FONT SIZES'),
    cellStyle: groupHeaderStyle,
  );
  sheet.merge(CellIndex.indexByString('A$row'), CellIndex.indexByString('E$row'));
  row += 2;

  sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue('Size (pt)'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('B$row'), TextCellValue('Dart Code'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('C$row'), TextCellValue('Sample (Arial)'), cellStyle: tableHeaderStyle);
  sheet.merge(CellIndex.indexByString('C$row'), CellIndex.indexByString('E$row'));
  row++;

  final sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24];
  for (var sz in sizes) {
    sheet.updateCell(
      CellIndex.indexByString('A$row'),
      IntCellValue(sz),
      cellStyle: cellBorderStyle.copyWith(horizontalAlignVal: HorizontalAlign.Center),
    );
    sheet.updateCell(
      CellIndex.indexByString('B$row'),
      TextCellValue('fontSize: $sz'),
      cellStyle: cellBorderStyle,
    );
    sheet.updateCell(
      CellIndex.indexByString('C$row'),
      TextCellValue('Sample text with size $sz pt'),
      cellStyle: cellBorderStyle.copyWith(
        fontFamilyVal: getFontFamily(FontFamily.Arial),
        fontSizeVal: sz,
      ),
    );
    sheet.merge(CellIndex.indexByString('C$row'), CellIndex.indexByString('E$row'));
    row++;
  }
  row += 2;

  // SECTION 3: FONT WEIGHTS & STYLES
  sheet.updateCell(
    CellIndex.indexByString('A$row'),
    TextCellValue('3. FONT STYLES & DECORATIONS'),
    cellStyle: groupHeaderStyle,
  );
  sheet.merge(CellIndex.indexByString('A$row'), CellIndex.indexByString('E$row'));
  row += 2;

  sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue('Style'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('B$row'), TextCellValue('Dart Code'), cellStyle: tableHeaderStyle);
  sheet.updateCell(CellIndex.indexByString('C$row'), TextCellValue('Sample'), cellStyle: tableHeaderStyle);
  sheet.merge(CellIndex.indexByString('C$row'), CellIndex.indexByString('E$row'));
  row++;

  final List<Map<String, dynamic>> styleShowcase = [
    {
      'name': 'Normal',
      'code': 'CellStyle()',
      'style': CellStyle(),
    },
    {
      'name': 'Bold',
      'code': 'CellStyle(bold: true)',
      'style': CellStyle(bold: true),
    },
    {
      'name': 'Italic',
      'code': 'CellStyle(italic: true)',
      'style': CellStyle(italic: true),
    },
    {
      'name': 'Single Underline',
      'code': 'CellStyle(underline: Underline.Single)',
      'style': CellStyle(underline: Underline.Single),
    },
    {
      'name': 'Double Underline',
      'code': 'CellStyle(underline: Underline.Double)',
      'style': CellStyle(underline: Underline.Double),
    },
    {
      'name': 'Strikethrough',
      'code': 'CellStyle(strikethrough: true)',
      'style': CellStyle(strikethrough: true),
    },
    {
      'name': 'Combined 1',
      'code': 'CellStyle(bold: true, italic: true, underline: Underline.Single)',
      'style': CellStyle(bold: true, italic: true, underline: Underline.Single, fontColorHex: ExcelColor.blue),
    },
    {
      'name': 'Combined 2',
      'code': 'CellStyle(bold: true, strikethrough: true, fontFamily: ...)',
      'style': CellStyle(bold: true, strikethrough: true, fontFamily: getFontFamily(FontFamily.Courier_New), fontColorHex: ExcelColor.red),
    },
  ];

  for (var item in styleShowcase) {
    sheet.updateCell(CellIndex.indexByString('A$row'), TextCellValue(item['name'] as String), cellStyle: cellBorderStyle);
    sheet.updateCell(CellIndex.indexByString('B$row'), TextCellValue(item['code'] as String), cellStyle: cellBorderStyle);
    
    // Copy the specific style and merge border
    final customStyle = (item['style'] as CellStyle).copyWith(
      leftBorderVal: Border(borderStyle: BorderStyle.Thin),
      rightBorderVal: Border(borderStyle: BorderStyle.Thin),
      topBorderVal: Border(borderStyle: BorderStyle.Thin),
      bottomBorderVal: Border(borderStyle: BorderStyle.Thin),
    );

    sheet.updateCell(
      CellIndex.indexByString('C$row'),
      TextCellValue('This text displays the "${item['name']}" style'),
      cellStyle: customStyle,
    );
    sheet.merge(CellIndex.indexByString('C$row'), CellIndex.indexByString('E$row'));
    row++;
  }

  // Set nice column widths for presentation
  sheet.setColumnWidth(0, 22.0);
  sheet.setColumnWidth(1, 48.0);
  sheet.setColumnWidth(2, 28.0);
  sheet.setColumnWidth(3, 28.0);
  sheet.setColumnWidth(4, 45.0);

  if (kIsWeb) {
    final bytes = excel.save(fileName: 'fonts_styles_example.xlsx');
    if (bytes != null && bytes.isNotEmpty) {
      return '✅ Fonts & Styles Example generated successfully!\n'
          'File size: ${(bytes.length / 1024).toStringAsFixed(2)} KB\n'
          'The download should start automatically.\n'
          '📌 File: fonts_styles_example.xlsx';
    }
    throw Exception('Failed to generate Excel file for Web.');
  } else {
    var bytes = excel.encode();
    if (bytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    String? outputFile = await FilePicker.platform.saveFile(
      dialogTitle: 'Save Fonts & Styles Example',
      fileName: 'fonts_styles_example.xlsx',
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (outputFile != null) {
      final file = File(outputFile);
      await file.writeAsBytes(bytes);
      final savedFileSize = await file.length();
      return '✅ Fonts & Styles Example saved successfully!\n'
          'Location: $outputFile\n'
          'Size: ${(savedFileSize / 1024).toStringAsFixed(2)} KB';
    }
    return 'Save cancelled.';
  }
}
