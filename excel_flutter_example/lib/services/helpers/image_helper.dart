import 'dart:convert';
import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel_community/excel_community.dart';

Future<String> generateExcelWithImageHelper() async {
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
