import '../models/section_detail.dart';
import 'helpers/image_helper.dart';
import 'helpers/simple_helper.dart';
import 'helpers/chart_helper.dart';
import 'helpers/styles_helper.dart';
import 'helpers/sheets_helper.dart';
import 'helpers/protection_helper.dart';
import 'helpers/full_demo_helper.dart';

class ExcelGenerator {
  static Future<String> generateExcelWithImage() => generateExcelWithImageHelper();

  static Future<String> generateSimpleExcel() => generateSimpleExcelHelper();

  static Future<String> generateExcelWithChart(ChartType type) => generateExcelWithChartHelper(type);

  static Future<String> generateUnderlineStyles() => generateUnderlineStylesHelper();

  static Future<String> generateNumberFormats() => generateNumberFormatsHelper();

  static Future<String> generateFullDemo() => generateFullDemoHelper();

  static Future<String> generateAllCharts() => generateAllChartsHelper();

  static Future<String> generateMultiSheets() => generateMultiSheetsHelper();

  static Future<String> generatePivotTemplate() => generatePivotTemplateHelper();

  static Future<String> generateLockedCellsReport() => generateLockedCellsReportHelper();

  static Future<String> generateFreezePanes() => generateFreezePanesHelper();
}
