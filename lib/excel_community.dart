/// A community-maintained fork of the excel library for reading, creating, editing and updating excel sheets.
///
/// Supports XLSX format with client and server-side compatibility, including chart support.
library excel_community;

import 'dart:convert';
import 'dart:math';
import 'package:archive/archive.dart';
import 'package:collection/collection.dart';
import 'package:equatable/equatable.dart';
import 'package:xml/xml.dart';
import 'src/web_helper/client_save_excel.dart'
    if (dart.library.html) 'src/web_helper/web_save_excel_browser.dart'
    as helper;

/// main directory
part 'src/excel.dart';

/// sharedStrigns
part 'src/sharedStrings/shared_strings.dart';

/// Number Format
part 'src/number_format/num_format.dart';
part 'src/number_format/formats/numbers/standard_formats.dart';
part 'src/number_format/formats/numbers/numeric_format.dart';
part 'src/number_format/formats/numbers/datetime_format.dart';
part 'src/number_format/formats/numbers/time_format.dart';

/// Chart
part 'src/chart/chart_base.dart';
part 'src/chart/chart_types.dart';

/// Utilities
part 'src/utilities/span.dart';
part 'src/utilities/fast_list.dart';
part 'src/utilities/utility.dart';
part 'src/utilities/constants.dart';
part 'src/utilities/enum.dart';
part 'src/utilities/chart_xml_writer.dart';
part 'src/utilities/chart_color_config.dart';
part 'src/utilities/archive.dart';
part 'src/utilities/colors/excel_color.dart';
part 'src/utilities/colors/base.dart';
part 'src/utilities/colors/red.dart';
part 'src/utilities/colors/blue.dart';
part 'src/utilities/colors/green.dart';
part 'src/utilities/colors/yellow_orange.dart';
part 'src/utilities/colors/others.dart';
part 'src/utilities/colors/accents.dart';

/// Chart Builders
part 'src/utilities/chart_builders/chart_style_builder.dart';
part 'src/utilities/chart_builders/column_bar_chart_builder.dart';
part 'src/utilities/chart_builders/line_chart_builder.dart';
part 'src/utilities/chart_builders/area_chart_builder.dart';
part 'src/utilities/chart_builders/scatter_chart_builder.dart';
part 'src/utilities/chart_builders/pie_chart_builder.dart';
part 'src/utilities/chart_builders/radar_chart_builder.dart';
part 'src/utilities/chart_builders/chart_style_builder_factory.dart';

/// Save
part 'src/save/save_file.dart';
part 'src/save/charts/chart_manager.dart';
part 'src/save/styles/style_manager.dart';
part 'src/save/styles/style_resource_collector.dart';
part 'src/save/styles/style_xml_builders.dart';
part 'src/save/worksheet/worksheet_manager.dart';
part 'src/save/workbook/workbook_manager.dart';
part 'src/save/self_correct_span.dart';
part 'src/parser/parse.dart';

/// Sheet
part 'src/sheet/sheet.dart';
part 'src/sheet/sheet_dimensions.dart';
part 'src/sheet/sheet_spans.dart';
part 'src/sheet/sheet_charts.dart';
part 'src/sheet/sheet_data_ext.dart';
part 'src/sheet/font_family.dart';
part 'src/sheet/data_model.dart';
part 'src/sheet/cell_value/cell_value.dart';
part 'src/sheet/cell_value/formula_cell_value.dart';
part 'src/sheet/cell_value/int_cell_value.dart';
part 'src/sheet/cell_value/double_cell_value.dart';
part 'src/sheet/cell_value/date_cell_value.dart';
part 'src/sheet/cell_value/text_cell_value.dart';
part 'src/sheet/cell_value/bool_cell_value.dart';
part 'src/sheet/cell_value/time_cell_value.dart';
part 'src/sheet/cell_value/datetime_cell_value.dart';
part 'src/sheet/cell_index.dart';
part 'src/sheet/cell_style.dart';
part 'src/sheet/font_style.dart';
part 'src/sheet/header_footer.dart';
part 'src/sheet/border_style.dart';
