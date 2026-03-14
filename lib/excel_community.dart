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
part 'src/utilities/colors.dart';

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
part 'src/save/self_correct_span.dart';
part 'src/parser/parse.dart';

/// Sheet
part 'src/sheet/sheet.dart';
part 'src/sheet/font_family.dart';
part 'src/sheet/data_model.dart';
part 'src/sheet/cell_index.dart';
part 'src/sheet/cell_style.dart';
part 'src/sheet/font_style.dart';
part 'src/sheet/header_footer.dart';
part 'src/sheet/border_style.dart';
