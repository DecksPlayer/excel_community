import 'dart:io';
import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';

void main() async {
  print('🔍 DEBUG: Verificando serialización de strikethrough...\n');
  
  // Crear Excel con strikethrough
  var excel = Excel.createExcel();
  var sheet = excel['Test'];
  
  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Strikethrough Text'),
    cellStyle: CellStyle(
      strikethrough: true,
      bold: true,
    ),
  );
  
  sheet.updateCell(
    CellIndex.indexByString('A2'),
    TextCellValue('Normal Text'),
    cellStyle: CellStyle(
      strikethrough: false,
    ),
  );
  
  // Encode
  var bytes = excel.encode()!;
  print('✅ Excel encoded: ${bytes.length} bytes\n');
  
  // Extract and read styles.xml
  var archive = ZipDecoder().decodeBytes(bytes);
  var stylesFile = archive.findFile('xl/styles.xml');
  
  if (stylesFile != null) {
    var stylesXml = String.fromCharCodes(stylesFile.content as List<int>);
    print('📄 styles.xml content:\n');
    
    // Print fonts section
    var fontsStart = stylesXml.indexOf('<fonts');
    var fontsEnd = stylesXml.indexOf('</fonts>') + 8;
    if (fontsStart != -1 && fontsEnd != -1) {
      var fontsSection = stylesXml.substring(fontsStart, fontsEnd);
      print('--- FONTS SECTION ---');
      print(fontsSection);
      print('\n');
      
      // Check for strikethrough elements
      if (fontsSection.contains('<strike')) {
        print('✅ FOUND: <strike> element in XML!');
      } else {
        print('❌ NOT FOUND: <strike> element missing from XML!');
      }
      
      // Count <b> and <i> elements
      var boldCount = '<b'.allMatches(fontsSection).length;
      var italicCount = '<i'.allMatches(fontsSection).length;
      var strikeCount = '<strike'.allMatches(fontsSection).length;
      
      print('\n📊 Element counts:');
      print('   <b> (bold): $boldCount');
      print('   <i> (italic): $italicCount');
      print('   <strike> (strikethrough): $strikeCount');
    }
  } else {
    print('❌ styles.xml not found in archive!');
  }
  
  print('\n🔄 Testing decode...\n');
  
  // Decode and check
  var excelDecoded = Excel.decodeBytes(bytes);
  var sheetDecoded = excelDecoded['Test'];
  
  var cell1 = sheetDecoded.cell(CellIndex.indexByString('A1'));
  var cell2 = sheetDecoded.cell(CellIndex.indexByString('A2'));
  
  print('Cell A1 (should have strikethrough):');
  print('   Value: ${cell1.value}');
  print('   Strikethrough: ${cell1.cellStyle?.isStrikethrough}');
  print('   Bold: ${cell1.cellStyle?.isBold}');
  
  print('\nCell A2 (should NOT have strikethrough):');
  print('   Value: ${cell2.value}');
  print('   Strikethrough: ${cell2.cellStyle?.isStrikethrough}');
  
  if (cell1.cellStyle?.isStrikethrough == true) {
    print('\n✅ SUCCESS: Strikethrough preserved!');
  } else {
    print('\n❌ FAIL: Strikethrough NOT preserved!');
  }
}
