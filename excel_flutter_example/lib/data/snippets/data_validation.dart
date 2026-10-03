const String dataValidationSnippet = r'''
import 'package:excel_community/excel_community.dart';

void addValidations(Excel excel) {
  final sheet = excel['Tasks'];

  sheet.addDataValidation('B2:B100',
      DataValidation.list(['Open', 'In progress', 'Done'])
          .withPrompt('Status', 'Pick a status')
          .withError('Invalid status', 'Use the list'));

  sheet.addDataValidation('C2:C100',
      DataValidation.listFromRange('A2:A5', sheetName: 'Lookup Lists'));

  sheet.addDataValidation('D2:D100',
      DataValidation.wholeNumber(DataValidationOperator.between, 1, 40));

  final rule = sheet.getDataValidation(CellIndex.indexByString('B2'));
  print(rule?.accepts(TextCellValue('Done'))); // true
}
''';
