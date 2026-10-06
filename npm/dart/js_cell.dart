import 'dart:js_interop';
import 'dart:js_interop_unsafe';
import 'package:excel_community/excel_community.dart';

@JSExport()
class JsCell {
  final Data _data;
  final Sheet _sheet;
  JsCell(this._data, this._sheet);

  String get cellId => _data.cellIndex.cellId;
  int get row => _data.rowIndex;
  int get col => _data.columnIndex;
  String get displayText => _data.displayText;

  String? get comment => _data.comment;
  set comment(String? val) => _data.comment = val;

  String? get formula {
    final v = _data.value;
    return v is FormulaCellValue ? v.formula : null;
  }

  set formula(String? val) {
    if (val != null) {
      _data.setFormula(val);
    }
  }

  String get type {
    final v = _data.value;
    if (v is TextCellValue) return 'string';
    if (v is IntCellValue) return 'int';
    if (v is DoubleCellValue) return 'double';
    if (v is BoolCellValue) return 'bool';
    if (v is DateCellValue) return 'date';
    if (v is DateTimeCellValue) return 'datetime';
    if (v is TimeCellValue) return 'time';
    if (v is FormulaCellValue) return 'formula';
    return 'null';
  }

  JSAny? get value {
    final v = _data.value;
    if (v is TextCellValue) return (v.value.text ?? v.toString()).toJS;
    if (v is IntCellValue) return v.value.toJS;
    if (v is DoubleCellValue) return v.value.toJS;
    if (v is BoolCellValue) return v.value.toJS;
    if (v is DateCellValue) {
      return '${v.year.toString().padLeft(4, '0')}-${v.month.toString().padLeft(2, '0')}-${v.day.toString().padLeft(2, '0')}'.toJS;
    }
    if (v is DateTimeCellValue) return v.asDateTimeLocal().toIso8601String().toJS;
    if (v is TimeCellValue) return '${v.hour}:${v.minute}:${v.second}'.toJS;
    if (v is FormulaCellValue) return v.formula.toJS;
    return null;
  }

  set value(JSAny? val) {
    if (val == null) {
      _data.value = null;
    } else if (val.isA<JSBoolean>()) {
      _data.value = BoolCellValue((val as JSBoolean).toDart);
    } else if (val.isA<JSNumber>()) {
      final numVal = (val as JSNumber).toDartDouble;
      if (numVal == numVal.roundToDouble() && !numVal.isInfinite && !numVal.isNaN) {
        _data.value = IntCellValue(numVal.toInt());
      } else {
        _data.value = DoubleCellValue(numVal);
      }
    } else if (val.isA<JSString>()) {
      final str = (val as JSString).toDart;
      if (str.startsWith('=')) {
        _data.value = FormulaCellValue(str);
      } else {
        _data.value = TextCellValue(str);
      }
    }
  }

  void setFormula(String formula) {
    _data.setFormula(formula);
  }

  // ==================== STYLING ====================

  static BorderStyle _parseBorderStyle(String? s) {
    if (s == null) return BorderStyle.Thin;
    switch (s.toLowerCase()) {
      case 'none': return BorderStyle.None;
      case 'dashed': return BorderStyle.Dashed;
      case 'dotted': return BorderStyle.Dotted;
      case 'double': return BorderStyle.Double;
      case 'medium': return BorderStyle.Medium;
      case 'thick': return BorderStyle.Thick;
      case 'hair': return BorderStyle.Hair;
      case 'dashdot': return BorderStyle.DashDot;
      case 'dashdotdot': return BorderStyle.DashDotDot;
      default: return BorderStyle.Thin;
    }
  }

  static Border? _parseBorder(JSObject? obj) {
    if (obj == null) return null;
    final styleProp = obj['style'];
    final colorProp = obj['color'];
    final style = styleProp != null && styleProp.isA<JSString>()
        ? _parseBorderStyle((styleProp as JSString).toDart)
        : BorderStyle.Thin;
    final color = colorProp != null && colorProp.isA<JSString>()
        ? ExcelColor.fromHexString((colorProp as JSString).toDart)
        : null;
    return Border(borderStyle: style, borderColorHex: color);
  }

  static NumFormat _parseNumFormat(JSAny? nf) {
    if (nf == null) return NumFormat.standard_0;
    if (nf.isA<JSString>()) {
      return NumFormat.custom(formatCode: (nf as JSString).toDart);
    }
    if (nf.isA<JSNumber>()) {
      final id = (nf as JSNumber).toDartDouble.toInt();
      switch (id) {
        case 0: return NumFormat.standard_0;
        case 1: return NumFormat.standard_1;
        case 2: return NumFormat.standard_2;
        case 3: return NumFormat.standard_3;
        case 4: return NumFormat.standard_4;
        case 9: return NumFormat.standard_9;
        case 10: return NumFormat.standard_10;
        case 11: return NumFormat.standard_11;
        case 14: return NumFormat.standard_14;
        case 15: return NumFormat.standard_15;
        case 16: return NumFormat.standard_16;
        case 17: return NumFormat.standard_17;
        case 18: return NumFormat.standard_18;
        case 19: return NumFormat.standard_19;
        case 20: return NumFormat.standard_20;
        case 21: return NumFormat.standard_21;
        case 22: return NumFormat.standard_22;
        case 37: return NumFormat.standard_37;
        case 38: return NumFormat.standard_38;
        case 41: return NumFormat.standard_41;
        case 42: return NumFormat.standard_42;
        case 44: return NumFormat.standard_44;
        case 49: return NumFormat.standard_49;
        default: return NumFormat.standard_0;
      }
    }
    return NumFormat.standard_0;
  }

  void setStyle(JSObject options) {
    bool bold = options['bold']?.isA<JSBoolean>() == true ? (options['bold'] as JSBoolean).toDart : false;
    bool italic = options['italic']?.isA<JSBoolean>() == true ? (options['italic'] as JSBoolean).toDart : false;
    bool strikethrough = options['strikethrough']?.isA<JSBoolean>() == true ? (options['strikethrough'] as JSBoolean).toDart : false;

    Underline underline = Underline.None;
    final uProp = options['underline'];
    if (uProp != null) {
      if (uProp.isA<JSBoolean>() && (uProp as JSBoolean).toDart) {
        underline = Underline.Single;
      } else if (uProp.isA<JSString>()) {
        final uStr = (uProp as JSString).toDart.toLowerCase();
        if (uStr == 'single') underline = Underline.Single;
        if (uStr == 'double') underline = Underline.Double;
      }
    }

    int? fontSize;
    final fsProp = options['fontSize'];
    if (fsProp != null && fsProp.isA<JSNumber>()) {
      fontSize = (fsProp as JSNumber).toDartDouble.toInt();
    }

    String? fontFamily;
    final ffProp = options['fontFamily'];
    if (ffProp != null && ffProp.isA<JSString>()) {
      fontFamily = (ffProp as JSString).toDart;
    }

    String? fontColor;
    final fcProp = options['fontColor'];
    if (fcProp != null && fcProp.isA<JSString>()) {
      fontColor = (fcProp as JSString).toDart;
    }

    String? bgColor;
    final bgProp = options['backgroundColor'];
    if (bgProp != null && bgProp.isA<JSString>()) {
      bgColor = (bgProp as JSString).toDart;
    }

    HorizontalAlign hAlign = HorizontalAlign.Left;
    final haProp = options['horizontalAlign'];
    if (haProp != null && haProp.isA<JSString>()) {
      final s = (haProp as JSString).toDart.toLowerCase();
      if (s == 'center') hAlign = HorizontalAlign.Center;
      if (s == 'right') hAlign = HorizontalAlign.Right;
    }

    VerticalAlign vAlign = VerticalAlign.Bottom;
    final vaProp = options['verticalAlign'];
    if (vaProp != null && vaProp.isA<JSString>()) {
      final s = (vaProp as JSString).toDart.toLowerCase();
      if (s == 'top') vAlign = VerticalAlign.Top;
      if (s == 'center') vAlign = VerticalAlign.Center;
    }

    bool wrapText = options['wrapText']?.isA<JSBoolean>() == true ? (options['wrapText'] as JSBoolean).toDart : false;

    int rotation = 0;
    final rotProp = options['rotation'];
    if (rotProp != null && rotProp.isA<JSNumber>()) {
      rotation = (rotProp as JSNumber).toDartDouble.toInt();
    }

    Border? allBorder = _parseBorder(options['border'] as JSObject?);
    Border? leftB = _parseBorder(options['leftBorder'] as JSObject?) ?? allBorder;
    Border? rightB = _parseBorder(options['rightBorder'] as JSObject?) ?? allBorder;
    Border? topB = _parseBorder(options['topBorder'] as JSObject?) ?? allBorder;
    Border? bottomB = _parseBorder(options['bottomBorder'] as JSObject?) ?? allBorder;

    NumFormat numFmt = _parseNumFormat(options['numberFormat']);

    _data.cellStyle = CellStyle(
      bold: bold,
      italic: italic,
      strikethrough: strikethrough,
      underline: underline,
      fontSize: fontSize,
      fontFamily: fontFamily,
      fontColorHex: fontColor != null ? ExcelColor.fromHexString(fontColor) : ExcelColor.black,
      backgroundColorHex: bgColor != null ? ExcelColor.fromHexString(bgColor) : ExcelColor.none,
      horizontalAlign: hAlign,
      verticalAlign: vAlign,
      textWrapping: wrapText ? TextWrapping.WrapText : null,
      rotation: rotation,
      leftBorder: leftB,
      rightBorder: rightB,
      topBorder: topB,
      bottomBorder: bottomB,
      numberFormat: numFmt,
    );
  }

  JSObject? get style {
    final s = _data.cellStyle;
    if (s == null) return null;
    final obj = JSObject();
    obj['bold'] = s.isBold.toJS;
    obj['italic'] = s.isItalic.toJS;
    obj['strikethrough'] = s.isStrikethrough.toJS;
    obj['underline'] = s.underline.name.toJS;
    if (s.fontSize != null) obj['fontSize'] = s.fontSize!.toJS;
    if (s.fontFamily != null) obj['fontFamily'] = s.fontFamily!.toJS;
    obj['fontColor'] = s.fontColor.colorHex.toJS;
    obj['backgroundColor'] = s.backgroundColor.colorHex.toJS;
    obj['horizontalAlign'] = s.horizontalAlignment.name.toJS;
    obj['verticalAlign'] = s.verticalAlignment.name.toJS;
    obj['rotation'] = s.rotation.toJS;
    obj['numberFormat'] = s.numberFormat.formatCode.toJS;
    return obj;
  }

  // ==================== HYPERLINKS ====================

  void setHyperlink(String url, [String? tooltip, String? display]) {
    _sheet.setHyperlink(
      _data.cellIndex,
      Hyperlink.url(url, tooltip: tooltip, display: display),
      text: display,
    );
  }

  JSObject? getHyperlink() {
    final link = _sheet.getHyperlink(_data.cellIndex);
    if (link == null) return null;
    final obj = JSObject();
    if (link.url != null) obj['url'] = link.url!.toJS;
    if (link.location != null) obj['location'] = link.location!.toJS;
    if (link.tooltip != null) obj['tooltip'] = link.tooltip!.toJS;
    if (link.display != null) obj['display'] = link.display!.toJS;
    return obj;
  }

  void removeHyperlink() {
    _sheet.removeHyperlink(_data.cellIndex);
  }
}
