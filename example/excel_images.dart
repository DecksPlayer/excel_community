import 'dart:io';
import 'package:path/path.dart';
import 'package:excel_community/excel_community.dart';

void main(List<String> args) {
  // -------------------------------------------------------------------------
  // 1. Create workbook and sheet
  // -------------------------------------------------------------------------
  final excel = Excel.createExcel();
  final sheet = excel['Images'];
  excel.delete('Sheet1');

  // -------------------------------------------------------------------------
  // 2. Cell labels
  // -------------------------------------------------------------------------
  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Excel Community — Image Embedding Example'),
    cellStyle: CellStyle(bold: true, fontSize: 14),
  );
  sheet.merge(CellIndex.indexByString('A1'), CellIndex.indexByString('F1'));

  sheet.updateCell(CellIndex.indexByString('A3'),
      TextCellValue('PNG (assets/logo.png):'),
      cellStyle: CellStyle(bold: true));
  sheet.updateCell(CellIndex.indexByString('D3'),
      TextCellValue('SVG — Excel 365 (assets/logo.svg):'),
      cellStyle: CellStyle(bold: true));

  // -------------------------------------------------------------------------
  // 3a. PNG — ExcelImage.fromFile auto-detects ExcelImageType.png
  // -------------------------------------------------------------------------
  sheet.addImage(ExcelImage.fromFile(
    File('assets/logo.png'),
    anchor: ImageAnchor.fromPixels(
      column: 0, row: 4,      // A5
      widthPixels: 200,
      heightPixels: 80,
    ),
  ));

  // -------------------------------------------------------------------------
  // 3b. SVG — ExcelImage.fromFile auto-detects ExcelImageType.svg
  //     Visible in Excel 2016+ and Microsoft 365.
  // -------------------------------------------------------------------------
  sheet.addImage(ExcelImage.fromFile(
    File('assets/logo.svg'),
    anchor: ImageAnchor.fromPixels(
      column: 3, row: 4,      // D5
      widthPixels: 200,
      heightPixels: 80,
    ),
  ));

  // -------------------------------------------------------------------------
  // 3c. Same PNG a second time using an explicit EMU anchor
  //     1 px at 96 DPI = 9 525 EMUs  |  1 cm ≈ 360 000 EMUs
  // -------------------------------------------------------------------------
  sheet.updateCell(CellIndex.indexByString('A12'),
      TextCellValue('PNG via EMU anchor (≈ 5 × 2 cm):'),
      cellStyle: CellStyle(bold: true));

  sheet.addImage(ExcelImage(
    imageBytes: File('assets/logo.png').readAsBytesSync(),
    imageType: ExcelImageType.png,
    anchor: ImageAnchor(
      fromColumn: 0,
      fromRow: 13,
      colOffset: 0,
      rowOffset: 0,
      widthEmu: 1_800_000,   // ≈ 5 cm
      heightEmu: 720_000,    // ≈ 2 cm
    ),
  ));

  // -------------------------------------------------------------------------
  // 4. Column widths
  // -------------------------------------------------------------------------
  sheet.setColumnWidth(0, 30);
  sheet.setColumnWidth(3, 30);

  // -------------------------------------------------------------------------
  // 5. Save
  // -------------------------------------------------------------------------
  final outputFile = join('.', 'example', 'excel_images_output.xlsx');
  final fileBytes = excel.encode();

  if (fileBytes != null) {
    File(outputFile)
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
    print('Saved → $outputFile');
  } else {
    print('Error: could not encode workbook.');
  }
}

