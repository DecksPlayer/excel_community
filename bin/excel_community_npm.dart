import 'dart:js_interop';
import 'dart:js_interop_unsafe';
import 'package:excel_community/src/npm/js_excel.dart';

void main() {
  globalContext['ExcelCommunity'] = createJSInteropWrapper(JsExcelCommunity());
}
