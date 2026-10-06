import 'dart:js_interop';
import 'dart:js_interop_unsafe';
import 'js_excel.dart';

void main() {
  globalContext['__excelCommunityCore'] = createJSInteropWrapper(JsExcelCommunity());
}
