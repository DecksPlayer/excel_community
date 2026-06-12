import 'dart:js_interop';
import 'dart:typed_data';

import 'package:web/web.dart';

// A wrapper to save the excel file in browser
class SavingHelper {
  static List<int>? saveFile(List<int>? val, String fileName) {
    if (val == null) {
      return null;
    }

    try {
      // Convert List<int> to Uint8List
      final bytes = Uint8List.fromList(val);

      // Create blob with proper MIME type for Excel files
      final blob = Blob(
        [bytes.toJS].toJS,
        BlobPropertyBag(
          type:
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ),
      );

      final url = URL.createObjectURL(blob);
      final anchor = HTMLAnchorElement()
        ..href = url
        ..download = fileName;

      document.body?.append(anchor);

      // Trigger download
      anchor.click();

      // Cleanup
      anchor.remove();
      URL.revokeObjectURL(url);

      return val;
    } catch (e) {
      // In case of error, print to console and return null
      print('Error saving Excel file in browser: $e');
      return null;
    }
  }
}
