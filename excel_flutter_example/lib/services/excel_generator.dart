import '../models/section_detail.dart';
import 'helpers/image_helper.dart';
import 'helpers/simple_helper.dart';
import 'helpers/chart_helper.dart';
import 'helpers/styles_helper.dart';
import 'helpers/sheets_helper.dart';
import 'helpers/protection_helper.dart';
import 'helpers/full_demo_helper.dart';
import 'helpers/multi_page_charts_helper.dart';
import 'helpers/hidden_columns_helper.dart';
import 'helpers/merged_cells_helper.dart';
import 'helpers/cell_comments_helper.dart';

class ExcelGenerator {
  static Future<String> generateExcelWithImage() =>
      generateExcelWithImageHelper();

  static Future<String> generateSimpleExcel() => generateSimpleExcelHelper();

  static Future<String> generateExcelWithChart(ChartType type) =>
      generateExcelWithChartHelper(type);

  static Future<String> generateFontsStyles() =>
      generateFontsStylesHelper();

  static Future<String> generateNumberFormats() =>
      generateNumberFormatsHelper();

  static Future<String> generateFullDemo() => generateFullDemoHelper();

  static Future<String> generateAllCharts() => generateAllChartsHelper();

  static Future<String> generateMultiSheets() => generateMultiSheetsHelper();

  static Future<String> generatePivotTemplate() =>
      generatePivotTemplateHelper();

  static Future<String> generateLockedCellsReport() =>
      generateLockedCellsReportHelper();

  static Future<String> generateFreezePanes() => generateFreezePanesHelper();

  static Future<String> generateMultiFreezePanes() =>
      generateMultiFreezePanesHelper();

  static Future<String> generateHiddenColumns() =>
      generateHiddenColumnsHelper();

  static Future<String> generateMultiPageCharts() =>
      generateMultiPageChartsHelper();

  static Future<String> generateMergedCells() =>
      generateMergedCellsHelper();

  static Future<String> generateCellComments() =>
      generateCellCommentsHelper();
}
