import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

void main() {
  group('ExcelColor Accessibility', () {
    test('Verify common colors', () {
      expect(ExcelColor.black, equals(BaseColors.black));
      expect(ExcelColor.white, equals(BaseColors.white));
      expect(ExcelColor.black.colorHex, equals('FF000000'));
      expect(ExcelColor.white.colorHex, equals('FFFFFFFF'));
    });

    test('Verify red colors', () {
      expect(ExcelColor.red, equals(RedColors.red));
      expect(ExcelColor.red400, equals(RedColors.red400));
      expect(ExcelColor.pinkAccent, equals(AccentColors.pinkAccent));
    });

    test('Verify orange and yellow colors', () {
      expect(ExcelColor.orange400, equals(YellowOrangeColors.orange400));
      expect(ExcelColor.orange400.colorHex, equals('FFFFA726'));
      expect(ExcelColor.amber700, equals(YellowOrangeColors.amber700));
    });

    test('Verify blue and green colors', () {
      expect(ExcelColor.blue, equals(BlueColors.blue));
      expect(ExcelColor.lightBlueAccent, equals(AccentColors.lightBlueAccent));
      expect(ExcelColor.green900, equals(GreenColors.green900));
    });

    test('Verify other colors', () {
      expect(ExcelColor.purple, equals(OtherColors.purple));
      expect(ExcelColor.brown50, equals(OtherColors.brown50));
      expect(ExcelColor.blueGrey800, equals(OtherColors.blueGrey800));
    });
  });
}
