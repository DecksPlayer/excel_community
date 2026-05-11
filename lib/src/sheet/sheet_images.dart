part of '../../excel_community.dart';

extension SheetImages on Sheet {
  /// Returns an unmodifiable list of images added to this sheet.
  List<ExcelImage> get images => List.unmodifiable(_images);

  /// Embeds an image in the sheet.
  ///
  /// ```dart
  /// final bytes = File('logo.png').readAsBytesSync();
  /// sheet.addImage(ExcelImage(
  ///   imageBytes: bytes,
  ///   imageType: ExcelImageType.png,
  ///   anchor: ImageAnchor.fromPixels(
  ///     column: 0, row: 0,
  ///     widthPixels: 200, heightPixels: 100,
  ///   ),
  /// ));
  /// ```
  void addImage(ExcelImage image) {
    _images.add(image);
  }
}
