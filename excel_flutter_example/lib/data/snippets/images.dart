// Image-embedding snippet: load a PNG from app assets or download it from
// the web, then embed both into a worksheet.
library;

// ---------------------------------------------------------------------------
// Image embedding (asset + remote URL)
// ---------------------------------------------------------------------------
const String imageEmbeddingSnippet = '''
import 'dart:typed_data';
import 'package:flutter/services.dart' show rootBundle;
import 'package:http/http.dart' as http;
import 'package:excel_community/excel_community.dart';

/// Reads a bundled asset as raw bytes (works on mobile, desktop and web).
Future<Uint8List> _loadAssetBytes(String path) async {
  final data = await rootBundle.load(path);
  return data.buffer.asUint8List(
    data.offsetInBytes,
    data.lengthInBytes,
  );
}

/// Downloads an image from a remote URL and returns its raw bytes.
Future<Uint8List> _loadNetworkBytes(String url) async {
  final response = await http.get(Uri.parse(url));
  if (response.statusCode != 200) {
    throw Exception(
      'Failed to download image (HTTP ' + response.statusCode.toString() + ') from ' + url,
    );
  }
  return response.bodyBytes;
}

Future<void> embedImage() async {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];

  // 1) Image bundled in the app assets (declared in pubspec.yaml)
  final Uint8List assetImageBytes =
      await _loadAssetBytes('assets/logo.png');
  sheet.addImage(ExcelImage(
    imageBytes: assetImageBytes,
    imageType: ExcelImageType.png,
    anchor: ImageAnchor.fromPixels(
      column: 1, row: 4, // B5 (0-indexed column B = 1, row 5 = 4)
      widthPixels: 150,
      heightPixels: 150,
    ),
  ));

  // 2) Image fetched from a remote URL on the web
  final Uint8List networkImageBytes = await _loadNetworkBytes(
    'https://picsum.photos/100/200',
  );
  sheet.addImage(ExcelImage(
    imageBytes: networkImageBytes,
    imageType: ExcelImageType.png,
    anchor: ImageAnchor.fromPixels(
      column: 1, row: 23, // B24
      widthPixels: 180,
      heightPixels: 120,
    ),
  ));

  excel.save(fileName: 'image_example.xlsx');
}
''';
