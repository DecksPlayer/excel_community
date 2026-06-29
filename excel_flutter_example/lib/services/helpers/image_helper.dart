import 'dart:io';
import 'dart:typed_data';
import 'package:flutter/foundation.dart';
import 'package:flutter/services.dart' show rootBundle;
import 'package:file_picker/file_picker.dart';
import 'package:http/http.dart' as http;
import 'package:excel_community/excel_community.dart';

/// Reads a bundled asset as raw bytes, working on every platform
/// (mobile, desktop and web — `dart:io` is not used to stay web-compatible).
Future<Uint8List> _loadAssetBytes(String path) async {
  final data = await rootBundle.load(path);
  return data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
}

/// Downloads an image from a remote URL and returns its raw bytes.
Future<Uint8List> _loadNetworkBytes(String url) async {
  final response = await http.get(Uri.parse(url));
  if (response.statusCode != 200) {
    throw Exception(
      'Failed to download image (HTTP ${response.statusCode}) from $url',
    );
  }
  return response.bodyBytes;
}

Future<String> generateExcelWithImageHelper() async {
  var excel = Excel.createExcel();
  var sheet = excel['Images Demo'];
  excel.delete('Sheet1');

  sheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Excel Community — Image Embedding Demo"),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 14,
      fontColorHex: ExcelColor.blue,
    ),
  );
  sheet.merge(CellIndex.indexByString("A1"), CellIndex.indexByString("F1"));

  sheet.updateCell(
    CellIndex.indexByString("A3"),
    TextCellValue(
      "Two images are embedded: one from a local asset and one fetched from the web.",
    ),
  );

  // 1) Local asset PNG (declared in pubspec.yaml: assets/logo.png)
  final Uint8List assetImageBytes = await _loadAssetBytes('assets/logo.png');
  sheet.addImage(
    ExcelImage(
      imageBytes: assetImageBytes,
      imageType: ExcelImageType.png,
      anchor: ImageAnchor.fromPixels(
        column: 1,
        row: 4, // B5
        widthPixels: 150,
        heightPixels: 150,
      ),
    ),
  );

  // 2) Remote PNG fetched from the web
  final Uint8List networkImageBytes = await _loadNetworkBytes(
    'https://picsum.photos/100/200',
  );
  sheet.addImage(
    ExcelImage(
      imageBytes: networkImageBytes,
      imageType: ExcelImageType.png,
      anchor: ImageAnchor.fromPixels(
        column: 1,
        row: 23, // B24
        widthPixels: 180,
        heightPixels: 120,
      ),
    ),
  );

  sheet.updateCell(
    CellIndex.indexByString("B11"),
    TextCellValue("Local asset image embedded above at B5"),
    cellStyle: CellStyle(italic: true, fontSize: 10),
  );
  sheet.updateCell(
    CellIndex.indexByString("B30"),
    TextCellValue("Web image fetched and embedded above at B24"),
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
