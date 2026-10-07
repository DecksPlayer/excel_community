import 'dart:convert';
import 'package:excel_community/excel_community.dart';
import 'js_convert.dart';

/// Builds Dart model objects from the plain JS option objects of the npm API
/// (already converted to maps with `dartify()`), and back.

// ==================== CELL STYLE ====================

BorderStyle _borderStyle(String? name) =>
    name == null ? BorderStyle.Thin : enumByName(BorderStyle.values, name, BorderStyle.Thin);

Border? _border(Object? value) {
  if (value is! Map) return null;
  final m = value.map((k, v) => MapEntry(k.toString(), v));
  final color = m.str('color');
  return Border(
    borderStyle: _borderStyle(m.str('style')),
    borderColorHex: color == null ? null : parseColor(color),
  );
}

NumFormat parseNumFormat(Object value) {
  if (value is num) {
    return NumFormatMaintainer().getByNumFmtId(value.toInt()) ?? NumFormat.standard_0;
  }
  return NumFormat.custom(formatCode: value.toString());
}

/// Returns a copy of [base] with only the options present in [o] changed.
CellStyle applyStyleOptions(CellStyle? base, Map<String, Object?> o) {
  final s = (base ?? CellStyle()).copyWith();
  if (o.flag('bold') case final v?) s.isBold = v;
  if (o.flag('italic') case final v?) s.isItalic = v;
  if (o.flag('strikethrough') case final v?) s.isStrikethrough = v;
  if (o.containsKey('underline')) {
    final u = o['underline'];
    s.underline = u == true || u == 'single'
        ? Underline.Single
        : u == 'double'
            ? Underline.Double
            : Underline.None;
  }
  if (o.integer('fontSize') case final v?) s.fontSize = v;
  if (o.str('fontFamily') case final v?) s.fontFamily = v;
  if (o.str('fontColor') case final v?) s.fontColor = parseColor(v);
  if (o.containsKey('backgroundColor')) {
    final v = o.str('backgroundColor');
    s.backgroundColor = v == null || v == 'none' ? ExcelColor.none : parseColor(v);
  }
  if (o.str('horizontalAlign') case final v?) {
    s.horizontalAlignment = enumByName(HorizontalAlign.values, v, HorizontalAlign.Left);
  }
  if (o.str('verticalAlign') case final v?) {
    s.verticalAlignment = enumByName(VerticalAlign.values, v, VerticalAlign.Bottom);
  }
  if (o.flag('wrapText') case final v?) s.wrap = v ? TextWrapping.WrapText : null;
  if (o.flag('shrinkToFit') case final v?) s.wrap = v ? TextWrapping.Clip : null;
  if (o.integer('rotation') case final v?) s.rotation = v;
  if (o['numberFormat'] case final v?) s.numberFormat = parseNumFormat(v);
  if (o.containsKey('border')) {
    final b = _border(o['border']) ?? Border();
    s
      ..leftBorder = b
      ..rightBorder = b
      ..topBorder = b
      ..bottomBorder = b;
  }
  if (o.containsKey('leftBorder')) s.leftBorder = _border(o['leftBorder']);
  if (o.containsKey('rightBorder')) s.rightBorder = _border(o['rightBorder']);
  if (o.containsKey('topBorder')) s.topBorder = _border(o['topBorder']);
  if (o.containsKey('bottomBorder')) s.bottomBorder = _border(o['bottomBorder']);
  if (o.containsKey('diagonalBorder')) s.diagonalBorder = _border(o['diagonalBorder']);
  if (o.flag('diagonalUp') case final v?) s.diagonalBorderUp = v;
  if (o.flag('diagonalDown') case final v?) s.diagonalBorderDown = v;
  if (o.containsKey('locked')) s.locked = o.flag('locked');
  if (o.containsKey('hidden')) s.hidden = o.flag('hidden');
  return s;
}

Map<String, Object?>? _borderInfo(Border b) => b.borderStyle == null
    ? null
    : {'style': b.borderStyle!.style, 'color': colorInfo(b.borderColorHex)};

Map<String, Object?> styleInfo(CellStyle s) => {
      'bold': s.isBold,
      'italic': s.isItalic,
      'strikethrough': s.isStrikethrough,
      'underline': optionName(s.underline),
      'fontSize': s.fontSize,
      'fontFamily': s.fontFamily,
      'fontColor': colorInfo(s.fontColor.colorHex),
      'backgroundColor': colorInfo(s.backgroundColor.colorHex),
      'horizontalAlign': optionName(s.horizontalAlignment),
      'verticalAlign': optionName(s.verticalAlignment),
      'wrapText': s.wrap == TextWrapping.WrapText,
      'shrinkToFit': s.wrap == TextWrapping.Clip,
      'rotation': s.rotation,
      'numberFormat': s.numberFormat.formatCode,
      'leftBorder': _borderInfo(s.leftBorder),
      'rightBorder': _borderInfo(s.rightBorder),
      'topBorder': _borderInfo(s.topBorder),
      'bottomBorder': _borderInfo(s.bottomBorder),
      'diagonalBorder': _borderInfo(s.diagonalBorder),
      'diagonalUp': s.diagonalBorderUp,
      'diagonalDown': s.diagonalBorderDown,
      'locked': s.locked,
      'hidden': s.hidden,
    };

// ==================== HYPERLINKS ====================

/// `{url}`, `{email, subject}`, `{sheet, cell}` or `{location}`, each with
/// optional `tooltip` and `display`.
Hyperlink parseHyperlink(Map<String, Object?> o) {
  final tooltip = o.str('tooltip');
  final display = o.str('display');
  if (o.str('url') case final url?) return Hyperlink.url(url, tooltip: tooltip, display: display);
  if (o.str('email') case final email?) {
    return Hyperlink.email(email, subject: o.str('subject'), tooltip: tooltip, display: display);
  }
  if (o.str('sheet') case final sheet?) {
    return Hyperlink.cell(sheet, o.str('cell') ?? 'A1', tooltip: tooltip, display: display);
  }
  if (o.str('location') case final location?) {
    return Hyperlink.location(location, tooltip: tooltip, display: display);
  }
  throw ArgumentError('A hyperlink needs url, email, sheet or location');
}

Map<String, Object?> hyperlinkInfo(Hyperlink link) => {
      'url': link.url,
      'location': link.location,
      'tooltip': link.tooltip,
      'display': link.display,
    };

// ==================== CHARTS ====================

ChartGrouping _grouping(String? name) => switch (name?.toLowerCase()) {
      'stacked' => ChartGrouping.stacked,
      'percentstacked' || 'percent' || '100' => ChartGrouping.percentStacked,
      _ => ChartGrouping.clustered,
    };

ChartSeriesStyle? _seriesStyle(Map<String, Object?> s) {
  final style = s.obj('style');
  final fill = style?.str('fillColor') ?? s.str('colorHex') ?? s.str('color');
  if (style == null && fill == null) return null;
  final border = style?.str('borderColor');
  return ChartSeriesStyle(
    fillColor: fill == null ? null : parseColor(fill),
    fillType: enumByName(ChartFillType.values, style?.str('fillType'), ChartFillType.solid),
    fillAlpha: style?.integer('fillAlpha') ?? 50,
    borderColor: border == null ? null : parseColor(border),
    borderAlpha: style?.integer('borderAlpha') ?? 100,
    borderWidth: style?['borderWidth']?.toString(),
  );
}

ChartAnchor _chartAnchor(Map<String, Object?>? a) {
  if (a == null) return ChartAnchor.at(column: 0, row: 0);
  if (a.containsKey('column') || a.containsKey('width')) {
    return ChartAnchor.at(
      column: a.integer('column') ?? 0,
      row: a.integer('row') ?? 0,
      width: a.integer('width') ?? 8,
      height: a.integer('height') ?? 15,
    );
  }
  final fromCol = a.integer('fromCol') ?? a.integer('fromColumn') ?? 0;
  final fromRow = a.integer('fromRow') ?? 0;
  return ChartAnchor(
    fromColumn: fromCol,
    fromRow: fromRow,
    toColumn: a.integer('toCol') ?? a.integer('toColumn') ?? fromCol + 8,
    toRow: a.integer('toRow') ?? fromRow + 15,
  );
}

ChartDataLabels? _dataLabels(Map<String, Object?>? d) => d == null
    ? null
    : ChartDataLabels(
        value: d.flag('value') ?? false,
        categoryName: d.flag('categoryName') ?? false,
        seriesName: d.flag('seriesName') ?? false,
        percentage: d.flag('percentage') ?? false,
        separator: d.str('separator') ?? ', ',
        labelPosition: d.str('labelPosition') ?? d.str('position'),
      );

/// Parses a chart config: `type`, `title`, `series`, `anchor` and the
/// options of each chart type.
Chart parseChart(Map<String, Object?> c) {
  final title = c.str('title') ?? 'Chart';
  final showLegend = c.flag('showLegend') ?? true;
  final anchor = _chartAnchor(c.obj('anchor'));
  final labels = _dataLabels(c.obj('dataLabels'));
  final grouping = _grouping(c.str('grouping'));
  final series = [
    for (final raw in c.list('series') ?? const [])
      if (raw is Map)
        () {
          final s = raw.map((k, v) => MapEntry(k.toString(), v));
          return ChartSeries(
            name: s.str('name') ?? '',
            categoriesRange: s.str('categoriesRange') ?? '',
            valuesRange: s.str('valuesRange') ?? '',
            bubbleSizeRange: s.str('bubbleSizeRange'),
            style: _seriesStyle(s),
          );
        }(),
  ];
  final type = (c.str('type') ?? 'column').toLowerCase().replaceAll(RegExp(r'[-_ ]'), '');
  return switch (type) {
    'bar' => BarChart(
        title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels,
        grouping: grouping),
    'line' => LineChart(
        title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels,
        grouping: grouping, showMarkers: c.flag('showMarkers') ?? true, smooth: c.flag('smooth') ?? false),
    'area' => AreaChart(
        title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels,
        grouping: grouping),
    'pie' => PieChart(title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels),
    'doughnut' =>
      DoughnutChart(title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels),
    'ofpie' || 'pieofpie' || 'barofpie' => OfPieChart(
        title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels,
        ofPieType: type == 'barofpie'
            ? OfPieType.bar
            : enumByName(OfPieType.values, c.str('ofPieType'), OfPieType.pie),
        splitType: enumByName(OfPieSplitType.values, c.str('splitType'), OfPieSplitType.position),
        splitPosition: c.integer('splitPosition') ?? 2,
        secondPieSize: c.integer('secondPieSize') ?? 75),
    'scatter' => ScatterChart(
        title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels,
        showLines: c.flag('showLines') ?? false, showMarkers: c.flag('showMarkers') ?? true,
        smooth: c.flag('smooth') ?? false),
    'bubble' => BubbleChart(
        title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels,
        bubbleScale: c.integer('bubbleScale') ?? 100,
        showNegativeBubbles: c.flag('showNegativeBubbles') ?? false),
    'stock' => StockChart(
        title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels,
        showHighLowLines: c.flag('showHighLowLines') ?? true, showUpDownBars: c.flag('showUpDownBars') ?? true),
    'radar' => RadarChart(
        title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels,
        filled: c.flag('filled') ?? false),
    'column' => ColumnChart(
        title: title, series: series, anchor: anchor, showLegend: showLegend, dataLabels: labels,
        grouping: grouping),
    _ => throw ArgumentError.value(c.str('type'), 'type', 'unknown chart type'),
  };
}

// ==================== CONDITIONAL FORMATTING ====================

/// One rule: `{type, operator, value, value2, formula, text, style, priority}`.
ConditionalFormattingRule parseConditionalRule(Map<String, Object?> r) {
  final style = r.obj('style') ?? r;
  final bg = style.str('backgroundColor');
  final fc = style.str('fontColor');
  final u = style['underline'];
  final differential = DifferentialStyle(
    backgroundColor: bg == null ? null : parseColor(bg),
    fontColor: fc == null ? null : parseColor(fc),
    bold: style.flag('bold'),
    italic: style.flag('italic'),
    strikethrough: style.flag('strikethrough'),
    underline: u == null
        ? null
        : u == 'double'
            ? Underline.Double
            : (u == true || u == 'single')
                ? Underline.Single
                : Underline.None,
  );
  final priority = r.integer('priority') ?? 1;
  String formula(Object? v) => v is String ? (v.startsWith('=') ? v.substring(1) : v) : '$v';
  final formulae = [
    if (r.list('formulae') case final list?) ...list.map(formula),
    if (r['value'] case final v?) formula(v),
    if (r['value2'] case final v?) formula(v),
  ];
  final type = enumByName(ConditionalFormattingType.values, r.str('type'), ConditionalFormattingType.cellIs);
  return switch (type) {
    ConditionalFormattingType.cellIs => ConditionalFormattingRule.cellIs(
        operator: enumByName(ConditionalFormattingOperator.values, r.str('operator'),
            ConditionalFormattingOperator.equal),
        formulae: formulae,
        style: differential,
        priority: priority),
    ConditionalFormattingType.expression => ConditionalFormattingRule.expression(
        formula: r.str('formula') != null ? formula(r.str('formula')) : formulae.first,
        style: differential,
        priority: priority),
    ConditionalFormattingType.containsText => ConditionalFormattingRule.containsText(
        text: r.str('text') ?? '', style: differential, priority: priority),
    ConditionalFormattingType.duplicateValues =>
      ConditionalFormattingRule.duplicateValues(style: differential, priority: priority),
    ConditionalFormattingType.uniqueValues =>
      ConditionalFormattingRule.uniqueValues(style: differential, priority: priority),
    _ => ConditionalFormattingRule(
        type: type,
        operator: ConditionalFormattingOperator.fromValue(type.value == 'notContainsText' ? 'notContains' : type.value),
        text: r.str('text'),
        style: differential,
        priority: priority),
  };
}

Map<String, Object?> conditionalRuleInfo(ConditionalFormattingRule r) => {
      'type': r.type.name,
      'operator': r.operator?.name,
      'formulae': r.formulae,
      'text': r.text,
      'priority': r.priority,
      'style': {
        'backgroundColor': colorInfo(r.style.backgroundColor?.colorHex),
        'fontColor': colorInfo(r.style.fontColor?.colorHex),
        'bold': r.style.bold,
        'italic': r.style.italic,
        'strikethrough': r.style.strikethrough,
        'underline': r.style.underline == null ? null : optionName(r.style.underline!),
      },
    };

// ==================== DATA VALIDATION ====================

DateTime _date(Object? v) => switch (v) {
      DateTime d => d.toLocal(),
      String s => DateTime.parse(s),
      _ => throw ArgumentError.value(v, 'value', 'expected a Date or an ISO date string'),
    };

Duration _time(Object? v) => switch (v) {
      String s => parseTime(s),
      num n => Duration(milliseconds: (n * 86400000).round()),
      _ => throw ArgumentError.value(v, 'value', "expected 'HH:MM[:SS]'"),
    };

/// `{type: 'list', items}`, `{type: 'listFromRange', range, sheet}`,
/// `{type: 'wholeNumber' | 'decimal' | 'date' | 'time' | 'textLength',
/// operator, value, value2}`, `{type: 'custom', formula}` or
/// `{type: 'inputMessage'}`, plus `prompt`, `error`, `allowBlank`,
/// `showDropdown`.
DataValidation parseDataValidation(Map<String, Object?> o) {
  final type = (o.str('type') ?? 'list').toLowerCase();
  final op = enumByName(DataValidationOperator.values, o.str('operator'), DataValidationOperator.between);
  final v1 = o['value'];
  final v2 = o['value2'];
  var rule = switch (type) {
    'list' => DataValidation.list([for (final i in o.list('items') ?? const []) '$i']),
    'listfromrange' => DataValidation.listFromRange(o.str('range') ?? 'A1', sheetName: o.str('sheet')),
    'wholenumber' || 'whole' => DataValidation.wholeNumber(op, (v1 as num).toInt(), (v2 as num?)?.toInt()),
    'decimal' => DataValidation.decimal(op, v1 as num, v2 as num?),
    'date' => DataValidation.date(op, _date(v1), v2 == null ? null : _date(v2)),
    'time' => DataValidation.time(op, _time(v1), v2 == null ? null : _time(v2)),
    'textlength' => DataValidation.textLength(op, (v1 as num).toInt(), (v2 as num?)?.toInt()),
    'custom' => DataValidation.custom(o.str('formula') ?? ''),
    'inputmessage' || 'any' => const DataValidation(type: DataValidationType.any),
    _ => throw ArgumentError.value(o.str('type'), 'type', 'unknown validation type'),
  };
  if (o.obj('prompt') case final p?) {
    rule = rule.withPrompt(p.str('title') ?? '', p.str('message') ?? '');
  }
  if (o.obj('error') case final e?) {
    rule = rule.withError(e.str('title') ?? '', e.str('message') ?? '',
        style: enumByName(DataValidationErrorStyle.values, e.str('style'), DataValidationErrorStyle.stop));
  }
  return rule.copyWith(
    allowBlank: o.flag('allowBlank'),
    showDropdown: o.flag('showDropdown'),
  );
}

Map<String, Object?> dataValidationInfo(DataValidation v) => {
      'type': v.type.name,
      'operator': v.operator.name,
      'formula1': v.formula1,
      'formula2': v.formula2,
      'items': v.listItems,
      'allowBlank': v.allowBlank,
      'showDropdown': v.showDropdown,
      'prompt': v.prompt == null ? null : {'title': v.promptTitle, 'message': v.prompt},
      'error': v.error == null ? null : {'title': v.errorTitle, 'message': v.error, 'style': v.errorStyle.name},
    };

// ==================== PAGE SETUP ====================

PaperSize _paperSize(Object value) {
  if (value is num) return PaperSize.fromCode(value.toInt());
  final name = value.toString().toLowerCase().replaceAll(RegExp(r'[-_ #]'), '');
  for (final p in PaperSize.values) {
    if (p.name.toLowerCase().replaceAll(RegExp(r'[-_ #()]|jis'), '') == name) return p;
  }
  throw ArgumentError.value(value, 'paperSize', 'unknown paper size');
}

/// Merges the options of [o] into [base].
PageSetup parsePageSetup(PageSetup? base, Map<String, Object?> o) {
  final b = base ?? const PageSetup();
  final fitWidth = o.integer('fitToWidth');
  final fitHeight = o.integer('fitToHeight');
  return b.copyWith(
    orientation: o.str('orientation') == null
        ? null
        : enumByName(PageOrientation.values, o.str('orientation'), PageOrientation.portrait),
    paperSize: o['paperSize'] == null ? null : _paperSize(o['paperSize']!),
    scale: o.integer('scale'),
    fitToPage: fitWidth != null || fitHeight != null ? true : o.flag('fitToPage'),
    fitToWidth: fitWidth,
    fitToHeight: fitHeight,
    firstPageNumber: o.integer('firstPageNumber'),
    useFirstPageNumber: o.integer('firstPageNumber') != null ? true : o.flag('useFirstPageNumber'),
    pageOrder: enumByNameOrNull(PageOrder.values, o.str('pageOrder')),
    blackAndWhite: o.flag('blackAndWhite'),
    draft: o.flag('draft'),
    cellComments: enumByNameOrNull(PrintCellComments.values, o.str('cellComments')),
    errors: enumByNameOrNull(PrintErrors.values, o.str('errors')),
    copies: o.integer('copies'),
  );
}

Map<String, Object?> pageSetupInfo(PageSetup p) => {
      'orientation': p.orientation?.name,
      'paperSize': p.paperSize?.name,
      'paperSizeCode': p.paperSize?.code,
      'scale': p.scale,
      'fitToPage': p.fitToPage,
      'fitToWidth': p.fitToWidth,
      'fitToHeight': p.fitToHeight,
      'firstPageNumber': p.firstPageNumber,
      'pageOrder': p.pageOrder?.name,
      'blackAndWhite': p.blackAndWhite,
      'draft': p.draft,
      'cellComments': p.cellComments?.name,
      'errors': p.errors?.name,
      'copies': p.copies,
    };

/// `'normal' | 'wide' | 'narrow'`, or `{left, right, top, bottom, header,
/// footer, unit: 'in' | 'cm'}`.
PageMargins parsePageMargins(Object value) {
  if (value is String) {
    return switch (value.toLowerCase()) {
      'normal' => PageMargins.normal,
      'wide' => PageMargins.wide,
      'narrow' => PageMargins.narrow,
      _ => throw ArgumentError.value(value, 'margins', "expected 'normal', 'wide' or 'narrow'"),
    };
  }
  final m = (value as Map).map((k, v) => MapEntry(k.toString(), v));
  if (m.str('unit') == 'cm') {
    const d = PageMargins.normal;
    return PageMargins.fromCentimeters(
      left: m.number('left') ?? d.left * 2.54,
      right: m.number('right') ?? d.right * 2.54,
      top: m.number('top') ?? d.top * 2.54,
      bottom: m.number('bottom') ?? d.bottom * 2.54,
      header: m.number('header') ?? d.header * 2.54,
      footer: m.number('footer') ?? d.footer * 2.54,
    );
  }
  return PageMargins(
    left: m.number('left') ?? 0.7,
    right: m.number('right') ?? 0.7,
    top: m.number('top') ?? 0.75,
    bottom: m.number('bottom') ?? 0.75,
    header: m.number('header') ?? 0.3,
    footer: m.number('footer') ?? 0.3,
  );
}

Map<String, Object?> pageMarginsInfo(PageMargins m) => {
      'left': m.left,
      'right': m.right,
      'top': m.top,
      'bottom': m.bottom,
      'header': m.header,
      'footer': m.footer,
    };

HeaderFooter parseHeaderFooter(HeaderFooter? base, Map<String, Object?> o) {
  final h = base ?? HeaderFooter();
  return HeaderFooter(
    oddHeader: o.containsKey('header') ? o.str('header') : (o.containsKey('oddHeader') ? o.str('oddHeader') : h.oddHeader),
    oddFooter: o.containsKey('footer') ? o.str('footer') : (o.containsKey('oddFooter') ? o.str('oddFooter') : h.oddFooter),
    evenHeader: o.containsKey('evenHeader') ? o.str('evenHeader') : h.evenHeader,
    evenFooter: o.containsKey('evenFooter') ? o.str('evenFooter') : h.evenFooter,
    firstHeader: o.containsKey('firstHeader') ? o.str('firstHeader') : h.firstHeader,
    firstFooter: o.containsKey('firstFooter') ? o.str('firstFooter') : h.firstFooter,
    differentFirst: o.flag('differentFirst') ?? h.differentFirst,
    differentOddEven: o.flag('differentOddEven') ?? h.differentOddEven,
    scaleWithDoc: o.flag('scaleWithDoc') ?? h.scaleWithDoc,
    alignWithMargins: o.flag('alignWithMargins') ?? h.alignWithMargins,
  );
}

Map<String, Object?> headerFooterInfo(HeaderFooter h) => {
      'oddHeader': h.oddHeader,
      'oddFooter': h.oddFooter,
      'evenHeader': h.evenHeader,
      'evenFooter': h.evenFooter,
      'firstHeader': h.firstHeader,
      'firstFooter': h.firstFooter,
      'differentFirst': h.differentFirst,
      'differentOddEven': h.differentOddEven,
    };

// ==================== TABLES ====================

/// `'TableStyleMedium9'`, `'medium9'`, `'light1'`, `'dark2'` or `null`/`'none'`.
TableStyle? parseTableStyle(String? name) {
  if (name == null || name == 'none') return null;
  final m = RegExp(r'^(?:TableStyle)?(light|medium|dark)(\d+)$', caseSensitive: false).firstMatch(name);
  if (m == null) return TableStyle.named(name);
  final n = int.parse(m.group(2)!);
  return switch (m.group(1)!.toLowerCase()) {
    'light' => TableStyle.light(n),
    'dark' => TableStyle.dark(n),
    _ => TableStyle.medium(n),
  };
}

/// Column names, or objects `{name, totalsFunction, totalsLabel,
/// totalsRowFormula, calculatedColumnFormula}`; also accepts a JSON string.
List<TableColumn>? parseTableColumns(Object? value) {
  final list = value is String ? jsonDecode(value) as List : value as List?;
  if (list == null) return null;
  return [
    for (final c in list)
      if (c is Map)
        () {
          final m = c.map((k, v) => MapEntry(k.toString(), v));
          return TableColumn(
            m.str('name') ?? '',
            totalsFunction: enumByName(TableTotalsFunction.values, m.str('totalsFunction'), TableTotalsFunction.none),
            totalsLabel: m.str('totalsLabel'),
            totalsRowFormula: m.str('totalsRowFormula'),
            calculatedColumnFormula: m.str('calculatedColumnFormula'),
          );
        }()
      else
        TableColumn('$c'),
  ];
}

Map<String, Object?> tableInfo(ExcelTable t) => {
      'name': t.name,
      'ref': t.ref,
      'style': t.style?.name,
      'columns': [
        for (final c in t.columns)
          {
            'name': c.name,
            'totalsFunction': c.totalsFunction.name,
            'totalsLabel': c.totalsLabel,
            'totalsRowFormula': c.totalsRowFormula,
            'calculatedColumnFormula': c.calculatedColumnFormula,
          }
      ],
      'showHeaderRow': t.showHeaderRow,
      'showTotalsRow': t.showTotalsRow,
      'showRowStripes': t.showRowStripes,
      'showColumnStripes': t.showColumnStripes,
      'showFirstColumn': t.showFirstColumn,
      'showLastColumn': t.showLastColumn,
      'showFilterButtons': t.showFilterButtons,
      'dataRowCount': t.dataRowCount,
    };

// ==================== PIVOT TABLES ====================

/// `{name, sourceSheet, sourceRange, targetCell, rows, columns, values:
/// [{field, function, customName}]}`.
PivotTable parsePivotTable(Map<String, Object?> o) => PivotTable(
      name: o.str('name') ?? 'PivotTable1',
      sourceSheet: o.str('sourceSheet') ?? (throw ArgumentError('sourceSheet is required')),
      sourceRange: o.str('sourceRange') ?? (throw ArgumentError('sourceRange is required')),
      targetCell: CellIndex.indexByString(o.str('targetCell') ?? 'A1'),
      rows: [for (final r in o.list('rows') ?? const []) '$r'],
      columns: [for (final c in o.list('columns') ?? const []) '$c'],
      values: [
        for (final v in o.list('values') ?? const [])
          if (v is Map)
            () {
              final m = v.map((k, v) => MapEntry(k.toString(), v));
              return PivotTableValue(
                field: m.str('field') ?? '',
                function: enumByName(
                    PivotValueFunction.values, _pivotFunctionAlias(m.str('function')), PivotValueFunction.sum),
                customName: m.str('customName') ?? m.str('name'),
              );
            }()
          else
            PivotTableValue(field: '$v'),
      ],
    );

String? _pivotFunctionAlias(String? name) => switch (name?.toLowerCase()) {
      'var' || 'variance' => 'varVal',
      'stddevpop' => 'stdDevp',
      'varpop' => 'varp',
      _ => name,
    };

// ==================== AUTOFILTER ====================

/// `{column, values, blank, custom: [{operator, value}], and}`; `column` is
/// the 0-based offset inside the filter range.
FilterColumn parseFilterColumn(Map<String, Object?> o) => FilterColumn(
      colId: o.integer('column') ?? o.integer('colId') ?? 0,
      filterValues: [for (final v in o.list('values') ?? const []) '$v'],
      blank: o.flag('blank') ?? false,
      customFilters: [
        for (final c in o.list('custom') ?? const [])
          if (c is Map)
            CustomFilterRule(
              operator: enumByName(FilterOperator.values, c['operator']?.toString(), FilterOperator.equal),
              val: '${c['value']}',
            ),
      ],
      customFiltersAnd: o.flag('and') ?? false,
    );
