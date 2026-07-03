const String mergedCellsSnippet = '''
import 'package:excel_community/excel_community.dart';

void generateMergedCells() {
  var excel = Excel.createExcel();
  excel.delete('Sheet1'); // Remove default sheet

  // ---------------------------------------------------------
  // SHEET 1: Project Roadmap (Column and Row Merging)
  // ---------------------------------------------------------
  var roadmapSheet = excel['Project Roadmap'];

  // A1:F2 Merged Header
  roadmapSheet.merge(CellIndex.indexByString("A1"), CellIndex.indexByString("F2"));
  roadmapSheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Global Corporate Project Roadmap"),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 16,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: ExcelColor.blue800,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    ),
  );

  // Column Headers in Row 4
  var headers = ["Phase", "Task Name", "Start Date", "End Date", "Owner", "Status"];
  for (var i = 0; i < headers.length; i++) {
    roadmapSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 3),
      TextCellValue(headers[i]),
      cellStyle: CellStyle(
        bold: true,
        backgroundColorHex: ExcelColor.grey300,
        horizontalAlign: HorizontalAlign.Center,
      ),
    );
  }

  // Row Merging in Column A (Phases spanning rows)
  roadmapSheet.merge(CellIndex.indexByString("A5"), CellIndex.indexByString("A7"));
  roadmapSheet.updateCell(
    CellIndex.indexByString("A5"),
    TextCellValue("Phase 1:\\nPlanning"),
    cellStyle: CellStyle(
      bold: true,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      textWrapping: TextWrapping.WrapText,
      backgroundColorHex: ExcelColor.blue50,
    ),
  );

  roadmapSheet.merge(CellIndex.indexByString("A8"), CellIndex.indexByString("A10"));
  roadmapSheet.updateCell(
    CellIndex.indexByString("A8"),
    TextCellValue("Phase 2:\\nExecution"),
    cellStyle: CellStyle(
      bold: true,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      textWrapping: TextWrapping.WrapText,
      backgroundColorHex: ExcelColor.amber50,
    ),
  );

  // Data rows for Phase 1
  roadmapSheet.updateCell(CellIndex.indexByString("B5"), TextCellValue("Requirement Gathering"));
  roadmapSheet.updateCell(CellIndex.indexByString("C5"), TextCellValue("2026-07-01"));
  roadmapSheet.updateCell(CellIndex.indexByString("D5"), TextCellValue("2026-07-05"));
  roadmapSheet.updateCell(CellIndex.indexByString("E5"), TextCellValue("Alice"));
  roadmapSheet.updateCell(CellIndex.indexByString("F5"), TextCellValue("Done"));

  roadmapSheet.updateCell(CellIndex.indexByString("B6"), TextCellValue("Architecture Setup"));
  roadmapSheet.updateCell(CellIndex.indexByString("C6"), TextCellValue("2026-07-06"));
  roadmapSheet.updateCell(CellIndex.indexByString("D6"), TextCellValue("2026-07-10"));
  roadmapSheet.updateCell(CellIndex.indexByString("E6"), TextCellValue("Bob"));
  roadmapSheet.updateCell(CellIndex.indexByString("F6"), TextCellValue("Done"));

  roadmapSheet.updateCell(CellIndex.indexByString("B7"), TextCellValue("Design Specifications"));
  roadmapSheet.updateCell(CellIndex.indexByString("C7"), TextCellValue("2026-07-11"));
  roadmapSheet.updateCell(CellIndex.indexByString("D7"), TextCellValue("2026-07-15"));
  roadmapSheet.updateCell(CellIndex.indexByString("E7"), TextCellValue("Charlie"));
  roadmapSheet.updateCell(CellIndex.indexByString("F7"), TextCellValue("In Progress"));

  // Data rows for Phase 2
  roadmapSheet.updateCell(CellIndex.indexByString("B8"), TextCellValue("Frontend Development"));
  roadmapSheet.updateCell(CellIndex.indexByString("C8"), TextCellValue("2026-07-16"));
  roadmapSheet.updateCell(CellIndex.indexByString("D8"), TextCellValue("2026-08-15"));
  roadmapSheet.updateCell(CellIndex.indexByString("E8"), TextCellValue("David"));
  roadmapSheet.updateCell(CellIndex.indexByString("F8"), TextCellValue("Planned"));

  roadmapSheet.updateCell(CellIndex.indexByString("B9"), TextCellValue("Backend API Development"));
  roadmapSheet.updateCell(CellIndex.indexByString("C9"), TextCellValue("2026-07-16"));
  roadmapSheet.updateCell(CellIndex.indexByString("D9"), TextCellValue("2026-08-20"));
  roadmapSheet.updateCell(CellIndex.indexByString("E9"), TextCellValue("Evan"));
  roadmapSheet.updateCell(CellIndex.indexByString("F9"), TextCellValue("Planned"));

  roadmapSheet.updateCell(CellIndex.indexByString("B10"), TextCellValue("Integration & Testing"));
  roadmapSheet.updateCell(CellIndex.indexByString("C10"), TextCellValue("2026-08-21"));
  roadmapSheet.updateCell(CellIndex.indexByString("D10"), TextCellValue("2026-08-31"));
  roadmapSheet.updateCell(CellIndex.indexByString("E10"), TextCellValue("Alice"));
  roadmapSheet.updateCell(CellIndex.indexByString("F10"), TextCellValue("Planned"));

  // Global Project Duration Summary merged at bottom
  roadmapSheet.merge(CellIndex.indexByString("A12"), CellIndex.indexByString("D12"));
  roadmapSheet.updateCell(
    CellIndex.indexByString("A12"),
    TextCellValue("Total Estimated Project Duration"),
    cellStyle: CellStyle(
      bold: true,
      horizontalAlign: HorizontalAlign.Right,
      backgroundColorHex: ExcelColor.grey200,
    ),
  );
  roadmapSheet.merge(CellIndex.indexByString("E12"), CellIndex.indexByString("F12"));
  roadmapSheet.updateCell(
    CellIndex.indexByString("E12"),
    TextCellValue("60 Business Days"),
    cellStyle: CellStyle(
      bold: true,
      fontColorHex: ExcelColor.blue,
      horizontalAlign: HorizontalAlign.Center,
      backgroundColorHex: ExcelColor.grey200,
    ),
  );


  // ---------------------------------------------------------
  // SHEET 2: Shift Schedule (Column-wise Merge & Grouping)
  // ---------------------------------------------------------
  var shiftSheet = excel['Shift Schedule'];

  // A1:G2 Merged Header
  shiftSheet.merge(CellIndex.indexByString("A1"), CellIndex.indexByString("G2"));
  shiftSheet.updateCell(
    CellIndex.indexByString("A1"),
    TextCellValue("Weekly Resource Shift Schedule"),
    cellStyle: CellStyle(
      bold: true,
      fontSize: 16,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: ExcelColor.teal800,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    ),
  );

  // Column Headers in Row 4
  var shiftHeaders = ["Role", "Resource", "Mon", "Tue", "Wed", "Thu", "Fri"];
  for (var i = 0; i < shiftHeaders.length; i++) {
    shiftSheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 3),
      TextCellValue(shiftHeaders[i]),
      cellStyle: CellStyle(
        bold: true,
        backgroundColorHex: ExcelColor.grey300,
        horizontalAlign: HorizontalAlign.Center,
      ),
    );
  }

  // Row Merging for Role "Support Lead"
  shiftSheet.merge(CellIndex.indexByString("A5"), CellIndex.indexByString("A7"));
  shiftSheet.updateCell(
    CellIndex.indexByString("A5"),
    TextCellValue("Support Lead"),
    cellStyle: CellStyle(
      bold: true,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      backgroundColorHex: ExcelColor.teal50,
    ),
  );

  // Shifts (Column Merges)
  shiftSheet.updateCell(CellIndex.indexByString("B5"), TextCellValue("Alice"));
  shiftSheet.merge(CellIndex.indexByString("C5"), CellIndex.indexByString("E5")); // Mon-Wed
  shiftSheet.updateCell(
    CellIndex.indexByString("C5"),
    TextCellValue("Morning Shift (8 AM - 4 PM)"),
    cellStyle: CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      backgroundColorHex: ExcelColor.green100,
    ),
  );
  shiftSheet.updateCell(CellIndex.indexByString("F5"), TextCellValue("Off"));
  shiftSheet.updateCell(CellIndex.indexByString("G5"), TextCellValue("Off"));

  shiftSheet.updateCell(CellIndex.indexByString("B6"), TextCellValue("Bob"));
  shiftSheet.updateCell(CellIndex.indexByString("C6"), TextCellValue("Off"));
  shiftSheet.updateCell(CellIndex.indexByString("D6"), TextCellValue("Off"));
  shiftSheet.merge(CellIndex.indexByString("E6"), CellIndex.indexByString("G6")); // Wed-Fri
  shiftSheet.updateCell(
    CellIndex.indexByString("E6"),
    TextCellValue("Evening Shift (4 PM - 12 AM)"),
    cellStyle: CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      backgroundColorHex: ExcelColor.orange100,
    ),
  );

  shiftSheet.updateCell(CellIndex.indexByString("B7"), TextCellValue("Charlie"));
  shiftSheet.merge(CellIndex.indexByString("C7"), CellIndex.indexByString("D7")); // Mon-Tue
  shiftSheet.updateCell(
    CellIndex.indexByString("C7"),
    TextCellValue("Night Shift (12 AM - 8 AM)"),
    cellStyle: CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      backgroundColorHex: ExcelColor.purple100,
    ),
  );
  shiftSheet.updateCell(CellIndex.indexByString("E7"), TextCellValue("Off"));
  shiftSheet.merge(CellIndex.indexByString("F7"), CellIndex.indexByString("G7")); // Thu-Fri
  shiftSheet.updateCell(
    CellIndex.indexByString("F7"),
    TextCellValue("Morning Shift (8 AM - 4 PM)"),
    cellStyle: CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      backgroundColorHex: ExcelColor.green100,
    ),
  );

  // Row Merging for Role "DevOps"
  shiftSheet.merge(CellIndex.indexByString("A8"), CellIndex.indexByString("A9"));
  shiftSheet.updateCell(
    CellIndex.indexByString("A8"),
    TextCellValue("DevOps Engineer"),
    cellStyle: CellStyle(
      bold: true,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      backgroundColorHex: ExcelColor.teal100,
    ),
  );

  shiftSheet.updateCell(CellIndex.indexByString("B8"), TextCellValue("David"));
  shiftSheet.merge(CellIndex.indexByString("C8"), CellIndex.indexByString("G8")); // Mon-Fri
  shiftSheet.updateCell(
    CellIndex.indexByString("C8"),
    TextCellValue("On-Call 24/7 Support"),
    cellStyle: CellStyle(
      bold: true,
      horizontalAlign: HorizontalAlign.Center,
      backgroundColorHex: ExcelColor.red100,
      fontColorHex: ExcelColor.red800,
    ),
  );

  shiftSheet.updateCell(CellIndex.indexByString("B9"), TextCellValue("Evan"));
  shiftSheet.merge(CellIndex.indexByString("C9"), CellIndex.indexByString("E9")); // Mon-Wed
  shiftSheet.updateCell(
    CellIndex.indexByString("C9"),
    TextCellValue("Standby duty"),
    cellStyle: CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      backgroundColorHex: ExcelColor.blueGrey100,
    ),
  );
  shiftSheet.merge(CellIndex.indexByString("F9"), CellIndex.indexByString("G9")); // Thu-Fri
  shiftSheet.updateCell(
    CellIndex.indexByString("F9"),
    TextCellValue("Night Shift (12 AM - 8 AM)"),
    cellStyle: CellStyle(
      horizontalAlign: HorizontalAlign.Center,
      backgroundColorHex: ExcelColor.purple100,
    ),
  );

  roadmapSheet.setColumnWidth(0, 16);
  roadmapSheet.setColumnWidth(1, 24);
  roadmapSheet.setColumnWidth(2, 14);
  roadmapSheet.setColumnWidth(3, 14);
  roadmapSheet.setColumnWidth(4, 12);
  roadmapSheet.setColumnWidth(5, 14);

  shiftSheet.setColumnWidth(0, 18);
  shiftSheet.setColumnWidth(1, 14);
  shiftSheet.setColumnWidth(2, 22);
  shiftSheet.setColumnWidth(3, 22);
  shiftSheet.setColumnWidth(4, 22);
  shiftSheet.setColumnWidth(5, 22);
  shiftSheet.setColumnWidth(6, 22);
}
''';
