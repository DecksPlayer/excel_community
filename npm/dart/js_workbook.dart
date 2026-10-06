import 'dart:convert';
import 'dart:js_interop';
import 'dart:typed_data';
import 'package:excel_community/excel_community.dart';
import 'js_convert.dart';
import 'js_sheet.dart';

@JSExport()
class JsWorkbook {
  final Excel _excel;
  JsWorkbook(this._excel);

  JSArray<JSString> get sheets =>
      plainJsArray(_excel.sheets.keys.map((s) => s.toJS));

  JSObject sheet(String name) {
    return createJSInteropWrapper(JsSheet(_excel[name]));
  }

  /// Adds a sheet; unlike [sheet], it never renames the empty `Sheet1`.
  JSObject createSheet(String name) {
    if (_excel.sheets.containsKey('Sheet1')) _excel['Sheet1'];
    return createJSInteropWrapper(JsSheet(_excel[name]));
  }

  bool deleteSheet(String name) {
    _excel.delete(name);
    return true;
  }

  bool renameSheet(String oldName, String newName) {
    _excel.rename(oldName, newName);
    return true;
  }

  bool copySheet(String fromSheet, String toSheet) {
    _excel.copy(fromSheet, toSheet);
    return true;
  }

  /// Makes [name] share its content with [existingSheet]: changes to either
  /// sheet apply to both.
  void linkSheet(String name, String existingSheet) {
    _excel.link(name, _excel[existingSheet]);
  }

  /// Gives a linked sheet its own copy of the content again.
  void unlinkSheet(String name) {
    _excel.unLink(name);
  }

  String? get defaultSheet => _excel.getDefaultSheet();
  set defaultSheet(String? name) {
    if (name != null) {
      _excel.setDefaultSheet(name);
    }
  }

  ExportValueMode _mode(Map<String, Object?> o) =>
      o.str('mode') == 'displayText' ? ExportValueMode.displayText : ExportValueMode.typed;

  /// Every sheet as `{sheetName: rows as objects}`.
  JSAny? toMaps([JSAny? options]) {
    final o = optionsMap(options);
    return dartToJs(_excel.toMaps(
      headerRow: o.integer('headerRow') ?? 0,
      mode: _mode(o),
      skipEmptyRows: o.flag('skipEmptyRows') ?? true,
    ));
  }

  String toJson([JSAny? options]) {
    final o = optionsMap(options);
    return _excel.toJson(
      headerRow: o.integer('headerRow') ?? 0,
      mode: _mode(o),
      skipEmptyRows: o.flag('skipEmptyRows') ?? true,
      indent: o.str('indent'),
    );
  }

  JSUint8Array? encode() {
    final bytes = _excel.encode();
    return bytes != null ? Uint8List.fromList(bytes).toJS : null;
  }

  String? encodeBase64() {
    final bytes = _excel.encode();
    return bytes != null ? base64Encode(bytes) : null;
  }
}
