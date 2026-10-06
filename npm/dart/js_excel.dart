import 'dart:convert';
import 'dart:js_interop';
import 'package:excel_community/excel_community.dart';
import 'js_workbook.dart';

@JSExport()
class JsExcelCommunity {
  JSObject create() {
    return createJSInteropWrapper(JsWorkbook(Excel.createExcel()));
  }

  JSObject read(JSUint8Array bytes) {
    return createJSInteropWrapper(JsWorkbook(Excel.decodeBytes(bytes.toDart)));
  }

  JSObject readBase64(String b64) {
    return createJSInteropWrapper(JsWorkbook(Excel.decodeBytes(base64Decode(b64))));
  }
}
