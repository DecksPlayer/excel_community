import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:excel_community/excel_community.dart';
import 'package:test/test.dart';

String _sheetXml(List<int> bytes) => utf8
    .decode(ZipDecoder().decodeBytes(bytes).findFile('xl/worksheets/sheet1.xml')!.content);

CellIndex _at(String ref) => CellIndex.indexByString(ref);

void main() {
  group('DataValidation model', () {
    test('list factories', () {
      final list = DataValidation.list(['Open', 'In progress', 'Done']);
      expect(list.formula1, '"Open,In progress,Done"');
      expect(list.listItems, ['Open', 'In progress', 'Done']);
      expect(() => DataValidation.list(['a,b']), throwsArgumentError);
      expect(() => DataValidation.list([]), throwsArgumentError);
      expect(() => DataValidation.list(List.filled(60, 'item')), throwsArgumentError,
          reason: 'over 255 characters');

      expect(DataValidation.listFromRange('A1:A10').formula1, r'$A$1:$A$10');
      expect(DataValidation.listFromRange('b2:b5', sheetName: 'Lookup Lists').formula1,
          r"'Lookup Lists'!$B$2:$B$5");
      expect(DataValidation.listFromRange('A1:A3').listItems, isNull);
    });

    test('comparison factories encode values', () {
      expect(DataValidation.wholeNumber(DataValidationOperator.between, 1, 10).formula2, '10');
      expect(DataValidation.decimal(DataValidationOperator.greaterThan, 0.5).formula1, '0.5');
      expect(DataValidation.date(DataValidationOperator.greaterThanOrEqual, DateTime(2026, 1, 1)).formula1,
          '46023');
      expect(DataValidation.time(DataValidationOperator.lessThan, const Duration(hours: 18)).formula1,
          '0.75');
      expect(() => DataValidation.textLength(DataValidationOperator.between, 3), throwsArgumentError,
          reason: 'between needs two values');
      expect(DataValidation.custom('=ISNUMBER(A2)').formula1, 'ISNUMBER(A2)');
    });

    test('accepts() evaluates values like Excel', () {
      final status = DataValidation.list(['Open', 'Done']);
      expect(status.accepts(TextCellValue('done')), isTrue, reason: 'case-insensitive');
      expect(status.accepts(TextCellValue('Closed')), isFalse);
      expect(status.accepts(null), isTrue, reason: 'allowBlank');
      expect(status.copyWith(allowBlank: false).accepts(null), isFalse);

      final score = DataValidation.wholeNumber(DataValidationOperator.between, 1, 10);
      expect(score.accepts(IntCellValue(10)), isTrue);
      expect(score.accepts(IntCellValue(11)), isFalse);
      expect(score.accepts(DoubleCellValue(2.5)), isFalse, reason: 'not whole');
      expect(score.accepts(TextCellValue('5')), isFalse);

      expect(DataValidation.decimal(DataValidationOperator.notBetween, 0, 1).accepts(DoubleCellValue(1.5)), isTrue);
      expect(DataValidation.textLength(DataValidationOperator.lessThanOrEqual, 5).accepts(TextCellValue('Hello!')),
          isFalse);
      final from2026 = DataValidation.date(DataValidationOperator.greaterThanOrEqual, DateTime(2026, 1, 1));
      expect(from2026.accepts(DateCellValue(year: 2026, month: 3, day: 1)), isTrue);
      expect(from2026.accepts(DateCellValue(year: 2025, month: 12, day: 31)), isFalse);
      final officeHours = DataValidation.time(
          DataValidationOperator.between, const Duration(hours: 9), const Duration(hours: 18));
      expect(officeHours.accepts(TimeCellValue(hour: 12, minute: 0, second: 0)), isTrue);
      expect(officeHours.accepts(TimeCellValue(hour: 20, minute: 0, second: 0)), isFalse);

      expect(DataValidation.custom('A1>0').accepts(IntCellValue(1)), isNull);
      expect(DataValidation.listFromRange('A1:A3').accepts(TextCellValue('x')), isNull);
    });

    test('builder methods set messages', () {
      final v = DataValidation.list(['A'])
          .withPrompt('Pick', 'Choose a value')
          .withError('Oops', 'Not allowed', style: DataValidationErrorStyle.warning);
      expect([v.promptTitle, v.prompt, v.errorTitle, v.error], ['Pick', 'Choose a value', 'Oops', 'Not allowed']);
      expect(v.errorStyle, DataValidationErrorStyle.warning);
    });
  });

  group('Sheet API', () {
    test('add, get and remove', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.addDataValidation('B2:B10', DataValidation.list(['Yes', 'No']));
      expect(sheet.getDataValidation(_at('B5'))?.listItems, ['Yes', 'No']);
      expect(sheet.getDataValidation(_at('C5')), isNull);
      expect(sheet.cell(_at('B2')).dataValidation, isNotNull);

      sheet.cell(_at('C1')).dataValidation =
          DataValidation.wholeNumber(DataValidationOperator.greaterThan, 0);
      expect(sheet.dataValidations.keys, containsAll(['B2:B10', 'C1']));

      sheet.cell(_at('C1')).dataValidation = null;
      sheet.clearDataValidations();
      expect(sheet.hasDataValidations, isFalse);
    });

    test('a new validation is subtracted from overlapping ones', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.addDataValidation('A1:C3', DataValidation.list(['x']));
      sheet.addDataValidation('B2', DataValidation.list(['y']));
      final x = sheet.dataValidations.entries.firstWhere((e) => e.value.listItems!.first == 'x').key;
      expect(x.split(' ').toSet(), {'A1:C1', 'A3:C3', 'A2', 'C2'});
      expect(sheet.getDataValidation(_at('B2'))?.listItems, ['y']);

      sheet.removeDataValidation('A1:C1');
      expect(sheet.getDataValidation(_at('A1')), isNull);
      expect(sheet.getDataValidation(_at('A3'))?.listItems, ['x']);
    });

    test('multiple ranges and row/column shifts', () {
      final sheet = Excel.createExcel()['Sheet1'];
      sheet.addDataValidation('A2:A5 C2:C5', DataValidation.list(['1', '2']));
      sheet.insertRow(0);
      expect(sheet.dataValidations.keys.single, 'A3:A6 C3:C6');
      sheet.removeColumn(0);
      expect(sheet.dataValidations.keys.single, 'B3:B6');
      sheet.removeRow(2);
      expect(sheet.dataValidations.keys.single, 'B3:B5');
    });

    test('copied with the sheet', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].addDataValidation('A1', DataValidation.list(['a']));
      excel.copy('Sheet1', 'Copy');
      expect(excel['Copy'].getDataValidation(_at('A1')), isNotNull);
    });
  });

  group('XLSX round-trip', () {
    test('writes schema-compliant <dataValidations>', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.cell(_at('A1')).value = TextCellValue('Status');
      sheet.addDataValidation(
          'A2:A10',
          DataValidation.list(['Open', 'Done'])
              .withPrompt('Status', 'Pick <one> & go')
              .withError('Invalid', 'Use the list', style: DataValidationErrorStyle.information));
      sheet.addDataValidation('B2:B10', DataValidation.wholeNumber(DataValidationOperator.between, 1, 5));
      sheet.addDataValidation('C2', DataValidation.list(['a']).copyWith(showDropdown: false));
      sheet.setHyperlink(_at('D1'), Hyperlink.url('https://pub.dev'));
      final xml = _sheetXml(excel.encode()!);

      expect(xml, contains('<dataValidations count="3">'));
      expect(
          xml,
          contains('<dataValidation type="list" errorStyle="information" allowBlank="1" '
              'showInputMessage="1" showErrorMessage="1" errorTitle="Invalid" error="Use the list" '
              'promptTitle="Status" prompt="Pick &lt;one&gt; &amp; go" sqref="A2:A10">'
              '<formula1>&quot;Open,Done&quot;</formula1></dataValidation>'));
      expect(xml, contains('sqref="B2:B10"><formula1>1</formula1><formula2>5</formula2>'));
      expect(xml, contains('showDropDown="1"'));
      // Schema order: conditionalFormatting?, dataValidations, hyperlinks, printOptions, pageMargins
      expect(xml.indexOf('</sheetData>'), lessThan(xml.indexOf('<dataValidations')));
      expect(xml.indexOf('</dataValidations>'), lessThan(xml.indexOf('<hyperlinks>')));
      expect(xml.indexOf('</hyperlinks>'), lessThan(xml.indexOf('<pageMargins')));
    });

    test('decodes what it encodes', () {
      final excel = Excel.createExcel();
      final sheet = excel['Sheet1'];
      sheet.cell(_at('A1')).value = TextCellValue('x');
      sheet.addDataValidation('A2:A10 C2:C10',
          DataValidation.list(['Low', 'Medium', 'High']).withPrompt('Priority', 'Pick one'));
      sheet.addDataValidation('B2:B10', DataValidation.listFromRange('A1:A5', sheetName: 'Lists'));
      sheet.addDataValidation('D2:D10',
          DataValidation.date(DataValidationOperator.between, DateTime(2026, 1, 1), DateTime(2026, 12, 31))
              .withError('Out of range', '2026 only', style: DataValidationErrorStyle.warning));
      sheet.addDataValidation('E2:E10', DataValidation.custom('AND(E2>0,E2<D2)'));
      sheet.addDataValidation('F2', DataValidation.inputMessage('Tip', 'Free text').copyWith(allowBlank: false));

      final decoded = Excel.decodeBytes(excel.encode()!)['Sheet1'];
      expect(decoded.dataValidations, sheet.dataValidations);
    });

    test('Excel 2010 x14 validations in extLst are preserved untouched', () {
      final excel = Excel.createExcel();
      excel['Sheet1'].cell(_at('A1')).value = TextCellValue('x');
      final bytes = excel.encode()!;
      const ext = '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" '
          'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
          '<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
          '<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Lists!\$A\$1:\$A\$3</xm:f>'
          '</x14:formula1><xm:sqref>B2:B5</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>';
      final source = ZipDecoder().decodeBytes(bytes);
      final patched = Archive();
      for (final file in source.files) {
        var content = file.content as List<int>;
        if (file.name == 'xl/worksheets/sheet1.xml') {
          content = utf8.encode(utf8.decode(content).replaceFirst('</worksheet>', '$ext</worksheet>'));
        }
        patched.addFile(ArchiveFile(file.name, content.length, content));
      }

      final reopened = Excel.decodeBytes(ZipEncoder().encode(patched));
      expect(reopened['Sheet1'].hasDataValidations, isFalse);
      final xml = _sheetXml(reopened.encode()!);
      expect(xml, contains('<x14:dataValidation type="list"'));
      expect(xml, isNot(contains('<dataValidations')));
    });
  });
}
