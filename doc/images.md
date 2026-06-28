# 🌄 Working with Images

Excel Community supports embedding PNG, JPEG, SVG, WebP, and other image formats into your worksheets. Embedded images float over cells and can be custom-positioned using column/row coordinates, pixel dimensions, and precise EMU offsets.

---

## 🖼️ Embedding an Image

Images are added to a worksheet using `sheet.addImage()`, passing an `ExcelImage` instance.

### 1. Explicit Bytes & Type

You can manually provide the image's raw bytes and specify its format.

```dart
import 'dart:io';
import 'package:excel_community/excel_community.dart';

void addImage(Excel excel) {
  final sheet = excel['Sheet1'];
  final pngBytes = File('assets/logo.png').readAsBytesSync();

  sheet.addImage(ExcelImage(
    imageBytes: pngBytes,
    imageType: ExcelImageType.png,
    anchor: ImageAnchor.fromPixels(
      column: 0, row: 2,        // Top-left: Column A, Row 3 (0-indexed)
      widthPixels: 200,
      heightPixels: 80,
    ),
  ));
}
```

### 2. Auto-Detecting Type from File

Using `ExcelImage.fromFile()` will automatically read the file's bytes and infer its format from the file extension.

```dart
final file = File('assets/logo.png');

sheet.addImage(ExcelImage.fromFile(
  file,
  anchor: ImageAnchor.fromPixels(
    column: 0, row: 2, 
    widthPixels: 200, heightPixels: 80,
  ),
));
```

### 3. Inferring Type from Extension String

If you already have bytes but want to detect the type using a file name string:

```dart
final bytes = File('assets/photo.jpg').readAsBytesSync();

sheet.addImage(ExcelImage(
  imageBytes: bytes,
  imageType: ExcelImageType.fromExtension('photo.jpg'), // Resolves to ExcelImageType.jpeg
  anchor: ImageAnchor.fromPixels(
    column: 0, row: 5, 
    widthPixels: 300, heightPixels: 200,
  ),
));
```

---

## ⚓ Image Positioning (`ImageAnchor`)

The positioning of an image is configured via `ImageAnchor`. You can position using screen pixels or exact EMUs (English Metric Units).

### A. Pixels (`fromPixels`)

Positions the image using coordinates and pixel dimensions. The DPI is assumed to be 96 (where 1 pixel = 9,525 EMUs).

```dart
ImageAnchor.fromPixels(
  column: 1,            // Column B
  row: 3,               // Row 4
  widthPixels: 200,     // Image width
  heightPixels: 100,    // Image height
  colOffsetPixels: 5,   // (Optional) Sub-column horizontal offset
  rowOffsetPixels: 5,   // (Optional) Sub-row vertical offset
)
```

### B. EMUs (`ImageAnchor` constructor)

For maximum accuracy, specify exact EMUs (1 cm ≈ 360,000 EMUs).

```dart
ImageAnchor(
  fromColumn: 1,
  fromRow: 3,
  colOffset: 47625,     // 5px offset
  rowOffset: 47625,     // 5px offset
  widthEmu: 1905000,    // ~5.29 cm
  heightEmu: 952500,    // ~2.65 cm
)
```

---

## 📋 Supported Formats

| `ExcelImageType` | Extensions | Notes / Compatibility |
|---|---|---|
| `png` | `.png` | Fully supported |
| `jpeg` | `.jpg`, `.jpeg`, `.jfif` | Fully supported |
| `gif` | `.gif` | Fully supported |
| `bmp` | `.bmp`, `.dib` | Fully supported |
| `tiff` | `.tif`, `.tiff` | Fully supported |
| `wmf` | `.wmf` | Vector format |
| `emf` | `.emf` | Vector format |
| `svg` | `.svg`, `.svgz` | Visible in Excel 2016+ and Microsoft 365 |
| `webp` | `.webp` | Visible in Microsoft 365 |
| `ico` | `.ico` | Icon files |
