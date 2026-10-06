import 'dart:convert';
import 'dart:js_interop';
import 'dart:typed_data';
import 'package:excel_community/excel_community.dart';
import 'js_sheet.dart';

@JSExport()
class JsWorkbook {
  final Excel _excel;
  JsWorkbook(this._excel);

  JSArray<JSString> get sheets =>
      _excel.sheets.keys.map((s) => s.toJS).toList().toJS;

  JSObject sheet(String name) {
    return createJSInteropWrapper(JsSheet(_excel[name]));
  }

  JSObject createSheet(String name) {
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

  String? get defaultSheet => _excel.getDefaultSheet();
  set defaultSheet(String? name) {
    if (name != null) {
      _excel.setDefaultSheet(name);
    }
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
