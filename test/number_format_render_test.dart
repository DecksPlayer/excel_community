import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

String fmt(String code, CellValue value) =>
    NumFormat.custom(formatCode: code).format(value);

void main() {
  group('Fractions', () {
    test('closest fraction for the denominator digits', () {
      expect(NumFormat.standard_12.format(DoubleCellValue(1.75)), '1 3/4');
      expect(NumFormat.standard_12.format(DoubleCellValue(1234.567)), '1234 4/7');
      expect(NumFormat.standard_13.format(DoubleCellValue(1234.567)), '1234 55/97');
    });

    test('pads numerator and denominator like Excel', () {
      expect(NumFormat.standard_13.format(DoubleCellValue(2.3125)), '2  5/16');
      expect(NumFormat.standard_13.format(DoubleCellValue(0.25)), '  1/4 ');
    });

    test('whole numbers blank the fraction and carry when it rounds to 1', () {
      expect(NumFormat.standard_12.format(IntCellValue(3)), '3    ');
      expect(NumFormat.standard_12.format(DoubleCellValue(1.99)), '2    ');
      expect(NumFormat.standard_12.format(DoubleCellValue(0.75)), ' 3/4');
    });

    test('fixed denominators, improper fractions and negatives', () {
      expect(fmt('# ?/8', DoubleCellValue(1.3)), '1 2/8');
      expect(fmt('# ??/100', DoubleCellValue(3.14159)), '3 14/100');
      expect(fmt('?/?', DoubleCellValue(1.75)), '7/4');
      expect(NumFormat.standard_12.format(DoubleCellValue(-1.75)), '-1 3/4');
    });
  });

  group('Scientific & engineering notation', () {
    test('engineering exponent is a multiple of three', () {
      expect(NumFormat.standard_48.format(DoubleCellValue(123456789)), '123.5E+6');
      expect(NumFormat.standard_48.format(DoubleCellValue(1234.567)), '1.2E+3');
      expect(NumFormat.standard_48.format(DoubleCellValue(0.000123)), '123.0E-6');
    });

    test('rounding carries into the exponent', () {
      expect(NumFormat.standard_11.format(DoubleCellValue(9.999)), '1.00E+01');
      expect(NumFormat.standard_11.format(DoubleCellValue(123456789)), '1.23E+08');
      expect(NumFormat.standard_11.format(IntCellValue(0)), '0.00E+00');
    });
  });

  group('Digit templates', () {
    test('literals between placeholders keep their position', () {
      expect(fmt('000-000-0000', IntCellValue(5551234567)), '555-123-4567');
      expect(fmt('000-000-0000', DoubleCellValue(1234.567)), '000-000-1235');
      expect(fmt('"Total: "#,##0', DoubleCellValue(1234.567)), 'Total: 1,235');
    });

    test('trailing commas scale by thousands', () {
      expect(fmt('#,##0,"K"', IntCellValue(1234567)), '1,235K');
      expect(fmt('#,##0,"K"', DoubleCellValue(1234.567)), '1K');
      expect(fmt('0.0,,"M"', IntCellValue(12345678)), '12.3M');
    });

    test('optional decimals and ? padding', () {
      expect(fmt('0.0#', DoubleCellValue(1.5)), '1.5');
      expect(fmt('0.0#', DoubleCellValue(1.567)), '1.57');
      expect(fmt('#.##', DoubleCellValue(0.5)), '.5');
      expect(fmt('.00', DoubleCellValue(1.5)), '1.50');
      expect(fmt('?0.0?', DoubleCellValue(1.5)), ' 1.5 ');
    });

    test('accounting zero sections render the dash', () {
      expect(NumFormat.standard_41.format(IntCellValue(0)), '-');
      expect(NumFormat.standard_42.format(IntCellValue(0)), r'$-');
      expect(NumFormat.standard_43.format(IntCellValue(0)), '-  ');
      expect(NumFormat.standard_44.format(IntCellValue(0)), r'$-  ');
    });

    test('booleans show TRUE/FALSE whatever the number format', () {
      expect(NumFormat.standard_0.format(BoolCellValue(true)), 'TRUE');
      expect(NumFormat.standard_2.format(BoolCellValue(false)), 'FALSE');
    });

    test('existing formats keep their output', () {
      expect(NumFormat.standard_4.format(DoubleCellValue(1234567.891)), '1,234,567.89');
      expect(NumFormat.standard_7.format(DoubleCellValue(-1234.567)), r'($1,234.57)');
      expect(NumFormat.standard_10.format(DoubleCellValue(0.756)), '75.60%');
      expect(NumFormat.standard_37.format(DoubleCellValue(1234.567)), '1,235 ');
      expect(fmt(r'[$€-2] #,##0.00', DoubleCellValue(-1234.567)), '-€ 1,234.57');
      expect(fmt('[Blue]#,##0;[Red]-#,##0;"zero"', IntCellValue(0)), 'zero');
      expect(fmt('00000', IntCellValue(42)), '00042');
    });
  });

  group('Date & time', () {
    const date = DateCellValue(year: 2026, month: 10, day: 3);
    const dateTime = DateTimeCellValue(
        year: 2026, month: 10, day: 3, hour: 14, minute: 30, second: 45, millisecond: 678);
    const time = TimeCellValue(hour: 14, minute: 30, second: 45, millisecond: 678);

    test('Taiwan era year with the [\$-404] locale', () {
      expect(NumFormat.standard_27.format(date), '115/10/3');
      expect(NumFormat.standard_29.format(date), '115年10月3日');
      expect(NumFormat.standard_28.format(dateTime), '115/10/3 2:30 下午');
    });

    test('Japanese era year with the [\$-411] locale', () {
      expect(fmt(r'[$-411]e/m/d', date), '8/10/3'); // Reiwa 8
      expect(fmt(r'[$-411]e/m/d', const DateCellValue(year: 2019, month: 4, day: 30)),
          '31/4/30'); // Heisei 31
    });

    test('Chinese AM/PM designator', () {
      expect(NumFormat.standard_34.format(time), '下午2時30分');
      expect(NumFormat.standard_35.format(time), '下午2時30分45秒');
      expect(NumFormat.standard_34.format(const TimeCellValue(hour: 9, minute: 5, second: 0)),
          '上午9時05分');
      expect(fmt(r'[$-412]AM/PM h:mm', dateTime), '오후 2:30');
    });

    test('fractions of a second', () {
      expect(NumFormat.standard_47.format(time), '3045.7');
      expect(fmt('hh:mm:ss.00', dateTime), '14:30:45.68');
    });

    test('elapsed hours on a date count from day zero', () {
      expect(fmt('[h]:mm', dateTime), '1111166:30');
      expect(NumFormat.standard_46.format(time), '14:30:45');
    });
  });
}
