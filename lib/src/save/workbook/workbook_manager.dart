part of '../../../excel_community.dart';

class _WorkbookManager {
  final Excel _excel;
  final Save _save;

  _WorkbookManager(this._excel, this._save);

  bool setDefaultSheet(String? sheetName) {
    if (sheetName == null || _excel._xmlFiles['xl/workbook.xml'] == null) {
      return false;
    }
    List<XmlElement> sheetList =
        _excel._xmlFiles['xl/workbook.xml']!.findAllElements('sheet').toList();
    XmlElement elementFound = XmlElement(XmlName(''));

    int position = -1;
    for (int i = 0; i < sheetList.length; i++) {
      var _sheetName = sheetList[i].getAttribute('name');
      if (_sheetName != null && _sheetName.toString() == sheetName) {
        elementFound = sheetList[i];
        position = i;
        break;
      }
    }

    if (position == -1) {
      return false;
    }
    if (position == 0) {
      return true;
    }

    _excel._xmlFiles['xl/workbook.xml']!
        .findAllElements('sheets')
        .first
        .children
      ..removeAt(position)
      ..insert(0, elementFound);

    String? expectedSheet = _excel._getDefaultSheet();

    return expectedSheet == sheetName;
  }

  void setSharedStrings() {
    var uniqueCount = 0;
    var count = 0;

    XmlElement shareString = _excel
        ._xmlFiles['xl/${_excel._sharedStringsTarget}']!
        .findAllElements('sst')
        .first;

    shareString.children.clear();

    _excel._sharedStrings._map.forEach((string, ss) {
      uniqueCount += 1;
      count += ss.count;

      shareString.children.add(string.node);
    });

    [
      ['count', '$count'],
      ['uniqueCount', '$uniqueCount']
    ].forEach((value) {
      if (shareString.getAttributeNode(value[0]) == null) {
        shareString.attributes.add(XmlAttribute(XmlName(value[0]), value[1]));
      } else {
        shareString.getAttributeNode(value[0])!.value = value[1];
      }
    });
  }

  void setMerge() {
    _selfCorrectSpanMap(_excel);
    _excel._mergeChangeLook.forEach((s) {
      if (_excel._sheetMap[s] != null &&
          _excel._sheetMap[s]!._spanList.isNotEmpty &&
          _excel._xmlSheetId.containsKey(s) &&
          _excel._xmlFiles.containsKey(_excel._xmlSheetId[s])) {
        Iterable<XmlElement>? iterMergeElement = _excel
            ._xmlFiles[_excel._xmlSheetId[s]]
            ?.findAllElements('mergeCells');
        late XmlElement mergeElement;
        if (iterMergeElement?.isNotEmpty ?? false) {
          mergeElement = iterMergeElement!.first;
        } else {
          if ((_excel._xmlFiles[_excel._xmlSheetId[s]]
                       ?.findAllElements('worksheet')
                       .length ??
                   0) >
               0) {
            int index = _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('worksheet')
                .first
                .children
                .indexOf(_excel._xmlFiles[_excel._xmlSheetId[s]]!
                    .findAllElements("sheetData")
                    .first);
            if (index == -1) {
              _damagedExcel();
            }
            _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('worksheet')
                .first
                .children
                .insert(
                    index + 1,
                    XmlElement(XmlName('mergeCells'),
                        [XmlAttribute(XmlName('count'), '0')]));

            mergeElement = _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('mergeCells')
                .first;
          } else {
            _damagedExcel();
          }
        }

        List<String> _spannedItems =
            List<String>.from(_excel._sheetMap[s]!.spannedItems);

        [
          ['count', _spannedItems.length.toString()],
        ].forEach((value) {
          if (mergeElement.getAttributeNode(value[0]) == null) {
            mergeElement.attributes
                .add(XmlAttribute(XmlName(value[0]), value[1]));
          } else {
            mergeElement.getAttributeNode(value[0])!.value = value[1];
          }
        });

        mergeElement.children.clear();

        _spannedItems.forEach((value) {
          mergeElement.children.add(XmlElement(XmlName('mergeCell'),
              [XmlAttribute(XmlName('ref'), '$value')], []));
        });
      }
    });
  }

  void setRTL() {
    _excel._rtlChangeLook.forEach((s) {
      var sheetObject = _excel._sheetMap[s];
      if (sheetObject != null &&
          _excel._xmlSheetId.containsKey(s) &&
          _excel._xmlFiles.containsKey(_excel._xmlSheetId[s])) {
        var itrSheetViewsRTLElement = _excel._xmlFiles[_excel._xmlSheetId[s]]
            ?.findAllElements('sheetViews');

        if (itrSheetViewsRTLElement?.isNotEmpty ?? false) {
          var itrSheetViewRTLElement = _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('sheetView');

          if (itrSheetViewRTLElement?.isNotEmpty ?? false) {
            _excel._xmlFiles[_excel._xmlSheetId[s]]
                ?.findAllElements('sheetViews')
                .first
                .children
                .clear();
          }

          _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('sheetViews')
              .first
              .children
              .add(XmlElement(
                XmlName('sheetView'),
                [
                  if (sheetObject.isRTL)
                    XmlAttribute(XmlName('rightToLeft'), '1'),
                  XmlAttribute(XmlName('workbookViewId'), '0'),
                ],
              ));
        } else {
          _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('worksheet')
              .first
              .children
              .add(XmlElement(XmlName('sheetViews'), [], [
                XmlElement(
                  XmlName('sheetView'),
                  [
                    if (sheetObject.isRTL)
                      XmlAttribute(XmlName('rightToLeft'), '1'),
                    XmlAttribute(XmlName('workbookViewId'), '0'),
                  ],
                )
              ]));
        }
      }
    });
  }

  void setHeaderFooter(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    if (sheet == null) return;

    final xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
    if (xmlFile == null) return;

    final sheetXmlElement = xmlFile.findAllElements("worksheet").first;

    final results = sheetXmlElement.findAllElements("headerFooter");
    if (results.isNotEmpty) {
      sheetXmlElement.children.remove(results.first);
    }

    if (sheet.headerFooter == null) return;

    sheetXmlElement.children.add(sheet.headerFooter!.toXmlElement());
  }
}
