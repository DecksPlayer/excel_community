part of '../../excel_community.dart';

/// Supported image formats for embedding in Excel worksheets.
///
/// All formats are natively supported by Microsoft Excel.
/// Use [ExcelImageType.fromExtension] to infer the type from a file path.
enum ExcelImageType {
  /// Portable Network Graphics (.png)
  png,

  /// JPEG / JFIF (.jpg / .jpeg)
  jpeg,

  /// Graphics Interchange Format (.gif)
  gif,

  /// Windows Bitmap (.bmp)
  bmp,

  /// Tagged Image File Format (.tif / .tiff)
  tiff,

  /// Windows Metafile (.wmf)
  wmf,

  /// Enhanced Metafile (.emf)
  emf,

  /// Scalable Vector Graphics (.svg) — Excel 2016+ / 365
  svg,

  /// WebP (.webp) — Excel 365 (Windows/Mac)
  webp,

  /// ICON (.ico)
  ico;

  /// Infers the [ExcelImageType] from a file extension or file path.
  ///
  /// ```dart
  /// final type = ExcelImageType.fromExtension('photo.jpg'); // jpeg
  /// ```
  ///
  /// Throws [ArgumentError] if the extension is not recognised.
  factory ExcelImageType.fromExtension(String pathOrExtension) {
    final ext = pathOrExtension.split('.').last.toLowerCase().trim();
    return switch (ext) {
      'png' => ExcelImageType.png,
      'jpg' || 'jpeg' || 'jfif' => ExcelImageType.jpeg,
      'gif' => ExcelImageType.gif,
      'bmp' || 'dib' => ExcelImageType.bmp,
      'tif' || 'tiff' => ExcelImageType.tiff,
      'wmf' => ExcelImageType.wmf,
      'emf' => ExcelImageType.emf,
      'svg' || 'svgz' => ExcelImageType.svg,
      'webp' => ExcelImageType.webp,
      'ico' => ExcelImageType.ico,
      _ => throw ArgumentError.value(
          pathOrExtension,
          'pathOrExtension',
          'Unrecognised image extension "$ext". '
              'Supported: png, jpg/jpeg, gif, bmp, tiff, wmf, emf, svg, webp, ico.',
        ),
    };
  }
}

/// Defines the position and size of an image embedded in a worksheet.
///
/// All offset and dimension values use EMUs (English Metric Units).
/// 1 pixel at 96 DPI = 9525 EMUs.
class ImageAnchor {
  /// 0-indexed column where the image starts.
  final int fromColumn;

  /// 0-indexed row where the image starts.
  final int fromRow;

  /// Column offset in EMUs from the left edge of [fromColumn]. Default 0.
  final int colOffset;

  /// Row offset in EMUs from the top edge of [fromRow]. Default 0.
  final int rowOffset;

  /// Width of the image in EMUs.
  final int widthEmu;

  /// Height of the image in EMUs.
  final int heightEmu;

  const ImageAnchor({
    required this.fromColumn,
    required this.fromRow,
    this.colOffset = 0,
    this.rowOffset = 0,
    required this.widthEmu,
    required this.heightEmu,
  });

  /// Creates an [ImageAnchor] using pixel dimensions (assumes 96 DPI).
  ///
  /// ```dart
  /// ImageAnchor.fromPixels(
  ///   column: 1, row: 2,
  ///   widthPixels: 200, heightPixels: 150,
  /// )
  /// ```
  factory ImageAnchor.fromPixels({
    required int column,
    required int row,
    required int widthPixels,
    required int heightPixels,
    int colOffsetPixels = 0,
    int rowOffsetPixels = 0,
  }) {
    const emusPerPixel = 9525;
    return ImageAnchor(
      fromColumn: column,
      fromRow: row,
      colOffset: colOffsetPixels * emusPerPixel,
      rowOffset: rowOffsetPixels * emusPerPixel,
      widthEmu: widthPixels * emusPerPixel,
      heightEmu: heightPixels * emusPerPixel,
    );
  }
}

/// Represents an image to be embedded in an Excel worksheet.
///
/// ```dart
/// import 'dart:io';
///
/// final imageBytes = File('photo.png').readAsBytesSync();
/// final image = ExcelImage(
///   imageBytes: imageBytes,
///   imageType: ExcelImageType.png,
///   anchor: ImageAnchor.fromPixels(
///     column: 1, row: 2,
///     widthPixels: 200, heightPixels: 150,
///   ),
/// );
/// sheet.addImage(image);
/// ```
///
/// You can also let the library detect the format automatically:
///
/// ```dart
/// ExcelImage.fromFile(File('logo.svg'), anchor: ...)
/// ```
class ExcelImage {
  /// Raw bytes of the image data.
  final List<int> imageBytes;

  /// Image format. Defaults to [ExcelImageType.png].
  final ExcelImageType imageType;

  /// Position and size of the image on the worksheet.
  final ImageAnchor anchor;

  const ExcelImage({
    required this.imageBytes,
    required this.anchor,
    this.imageType = ExcelImageType.png,
  });

  /// Creates an [ExcelImage] from a [File], inferring [imageType] from the
  /// file extension.
  ///
  /// ```dart
  /// final image = ExcelImage.fromFile(
  ///   File('logo.png'),
  ///   anchor: ImageAnchor.fromPixels(column: 0, row: 0, widthPixels: 120, heightPixels: 60),
  /// );
  /// ```
  factory ExcelImage.fromFile(dynamic file, {required ImageAnchor anchor}) {
    // Accept dart:io File or anything with readAsBytesSync + path
    final bytes = (file as dynamic).readAsBytesSync() as List<int>;
    final path = (file as dynamic).path as String;
    final type = ExcelImageType.fromExtension(path);
    return ExcelImage(imageBytes: bytes, imageType: type, anchor: anchor);
  }

  String get _contentType => switch (imageType) {
        ExcelImageType.png => 'image/png',
        ExcelImageType.jpeg => 'image/jpeg',
        ExcelImageType.gif => 'image/gif',
        ExcelImageType.bmp => 'image/bmp',
        ExcelImageType.tiff => 'image/tiff',
        ExcelImageType.wmf => 'image/x-wmf',
        ExcelImageType.emf => 'image/x-emf',
        ExcelImageType.svg => 'image/svg+xml',
        ExcelImageType.webp => 'image/webp',
        ExcelImageType.ico => 'image/x-icon',
      };

  String get _extension => switch (imageType) {
        ExcelImageType.png => 'png',
        ExcelImageType.jpeg => 'jpeg',
        ExcelImageType.gif => 'gif',
        ExcelImageType.bmp => 'bmp',
        ExcelImageType.tiff => 'tiff',
        ExcelImageType.wmf => 'wmf',
        ExcelImageType.emf => 'emf',
        ExcelImageType.svg => 'svg',
        ExcelImageType.webp => 'webp',
        ExcelImageType.ico => 'ico',
      };
}
