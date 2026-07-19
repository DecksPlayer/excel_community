part of '../../../excel_community.dart';

class _PivotTableManager {
  final Excel _excel;
  final Save _save;

  _PivotTableManager(this._excel, this._save);

  void processPivotTables() {
    int pivotTableCount = _countExistingPivotTables();
    int pivotCacheCount = _countExistingPivotCaches();

    _excel._sheetMap.forEach((sheetName, sheet) {
      if (sheet.pivotTables.isEmpty) return;

      final sheetId = _excel._xmlSheetId[sheetName]!;
      final sheetFileName = sheetId.split('/').last; // e.g. sheet1.xml
      final sheetRelsPath = 'xl/worksheets/_rels/$sheetFileName.rels';

      sheet._pivotTableRIds.clear();

      // Ensure worksheet rels XML exists
      var sheetRels = _excel._xmlFiles[sheetRelsPath];
      if (sheetRels == null) {
        final relsBuilder = XmlBuilder();
        relsBuilder.processing(
            'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
        relsBuilder.element('Relationships',
            attributes: {
              'xmlns':
                  'http://schemas.openxmlformats.org/package/2006/relationships',
            },
            nest: () {});
        sheetRels = relsBuilder.buildDocument();
        _excel._xmlFiles[sheetRelsPath] = sheetRels;
      }
      final sheetRelsRoot = sheetRels.findAllElements('Relationships').first;

      for (final pt in sheet.pivotTables) {
        pivotTableCount++;
        pivotCacheCount++;

        final sourceSheetObj = _excel._sheetMap[pt.sourceSheet];
        if (sourceSheetObj == null) continue;

        final sourceHeaders = _extractHeaders(sourceSheetObj, pt.sourceRange);
        if (sourceHeaders.isEmpty) continue;

        final pivotTablePath = 'xl/pivotTables/pivotTable$pivotTableCount.xml';
        final pivotTableRelsPath =
            'xl/pivotTables/_rels/pivotTable$pivotTableCount.xml.rels';
        final pivotCacheDefPath =
            'xl/pivotCache/pivotCacheDefinition$pivotCacheCount.xml';
        final pivotCacheDefRelsPath =
            'xl/pivotCache/_rels/pivotCacheDefinition$pivotCacheCount.xml.rels';
        final pivotCacheRecPath =
            'xl/pivotCache/pivotCacheRecords$pivotCacheCount.xml';

        // 1. Generate Pivot Table Definition XML
        _excel._xmlFiles[pivotTablePath] =
            _buildPivotTableXml(pt, pivotCacheCount, sourceHeaders);

        // 2. Generate Pivot Table Rels
        _excel._xmlFiles[pivotTableRelsPath] =
            _buildPivotTableRelsXml(pivotCacheCount);

        // 3. Generate Pivot Cache Definition XML
        _excel._xmlFiles[pivotCacheDefPath] =
            _buildPivotCacheDefXml(pt, sourceHeaders);

        // 4. Generate Pivot Cache Rels
        _excel._xmlFiles[pivotCacheDefRelsPath] =
            _buildPivotCacheDefRelsXml(pivotCacheCount);

        // 5. Generate Pivot Cache Records XML
        _excel._xmlFiles[pivotCacheRecPath] = _buildPivotCacheRecordsXml();

        // Register Content Types
        _save._addContentType(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
          '/$pivotTablePath',
        );
        _save._addContentType(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
          '/$pivotCacheDefPath',
        );
        _save._addContentType(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
          '/$pivotCacheRecPath',
        );

        // Wire to worksheet relationships
        final nextRIdIndex =
            sheetRelsRoot.children.whereType<XmlElement>().length + 1;
        final ptRId = 'rId$nextRIdIndex';

        sheetRelsRoot.children.add(XmlElement(XmlName.parts('Relationship'), [
          XmlAttribute(XmlName.parts('Id'), ptRId),
          XmlAttribute(
            XmlName.parts('Type'),
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
          ),
          XmlAttribute(XmlName.parts('Target'),
              '../pivotTables/pivotTable$pivotTableCount.xml'),
        ]));

        sheet._pivotTableRIds.add(ptRId);

        // Wire to workbook relationships
        var workbookRels = _excel._xmlFiles['xl/_rels/workbook.xml.rels'];
        if (workbookRels != null) {
          final wbRelsRoot = workbookRels.findAllElements('Relationships').first;
          final wbNextRIdNum = _getAvailableWorkbookRId(wbRelsRoot);
          final wbRId = 'rId$wbNextRIdNum';

          wbRelsRoot.children.add(XmlElement(XmlName.parts('Relationship'), [
            XmlAttribute(XmlName.parts('Id'), wbRId),
            XmlAttribute(
              XmlName.parts('Type'),
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
            ),
            XmlAttribute(XmlName.parts('Target'),
                'pivotCache/pivotCacheDefinition$pivotCacheCount.xml'),
          ]));

          // Register cache inside workbook.xml
          _addPivotCacheToWorkbookXml(pivotCacheCount, wbRId);
        }
      }
    });
  }

  int _countExistingPivotTables() {
    final keys = <String>{};
    keys.addAll(_excel._xmlFiles.keys);
    if (_excel._archive != null) {
      for (final f in _excel._archive.files) {
        keys.add(f.name);
      }
    }
    return keys
        .where((k) =>
            k.startsWith('xl/pivotTables/pivotTable') &&
            k.endsWith('.xml') &&
            !k.contains('/_rels/'))
        .length;
  }

  int _countExistingPivotCaches() {
    final keys = <String>{};
    keys.addAll(_excel._xmlFiles.keys);
    if (_excel._archive != null) {
      for (final f in _excel._archive.files) {
        keys.add(f.name);
      }
    }
    return keys
        .where((k) =>
            k.startsWith('xl/pivotCache/pivotCacheDefinition') &&
            k.endsWith('.xml') &&
            !k.contains('/_rels/'))
        .length;
  }

  List<String> _extractHeaders(Sheet sheet, String range) {
    final List<String> headers = [];
    try {
      final rangeParts = range.split(':');
      if (rangeParts.length == 2) {
        final startCoords = _cellCoordsFromCellId(rangeParts[0]);
        final endCoords = _cellCoordsFromCellId(rangeParts[1]);
        final row = startCoords.$1;
        for (int col = startCoords.$2; col <= endCoords.$2; col++) {
          final cell = sheet._sheetData[row]?[col];
          final val = cell?.value?.toString() ?? getCellId(col, row);
          headers.add(val);
        }
      }
    } catch (e) {
      // Ignore parsing errors and return empty list
    }
    return headers;
  }

  XmlDocument _buildPivotTableXml(
      PivotTable pt, int cacheId, List<String> sourceHeaders) {
    final builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');

    final startCell = pt.targetCell.cellId;
    final endColIdx =
        pt.targetCell.columnIndex + max(1, pt.columns.length + pt.values.length);
    final endRowIdx = pt.targetCell.rowIndex + max(2, pt.rows.length + 10);
    final endCell = getCellId(endColIdx, endRowIdx);

    builder.element('pivotTableDefinition',
        attributes: {
          'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          'xmlns:r':
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'name': pt.name,
          'cacheId': cacheId.toString(),
          'dataOnRows': '0',
          'applyWidthHeightFormats': '1',
          'applyNumberFormats': '1',
          'applyBorderFormats': '1',
          'applyFontFormats': '1',
          'applyPatternFormats': '1',
          'applyAlignmentFormats': '1',
          'showHeaders': '1',
          'showGridLines': '1',
          'rowHeaderCaption': 'Row Labels',
          'colHeaderCaption': 'Column Labels',
          'columnGrandTotals': '1',
          'rowGrandTotals': '1',
        }, nest: () {
      builder.element('location', attributes: {
        'ref': '$startCell:$endCell',
        'leftShift': '0',
        'topShift': '0',
        'fitWidth': '1',
      });

      builder.element('pivotFields',
          attributes: {'count': sourceHeaders.length.toString()}, nest: () {
        for (int i = 0; i < sourceHeaders.length; i++) {
          final header = sourceHeaders[i];
          final isRow = pt.rows.contains(header);
          final isCol = pt.columns.contains(header);
          final isVal = pt.values.any((v) => v.field == header);

          if (isRow) {
            builder.element('pivotField', attributes: {
              'axis': 'axisRow',
              'showAllItems': '0',
            }, nest: () {
              builder.element('items', attributes: {'count': '1'}, nest: () {
                builder.element('item', attributes: {'t': 'default'});
              });
            });
          } else if (isCol) {
            builder.element('pivotField', attributes: {
              'axis': 'axisCol',
              'showAllItems': '0',
            }, nest: () {
              builder.element('items', attributes: {'count': '1'}, nest: () {
                builder.element('item', attributes: {'t': 'default'});
              });
            });
          } else if (isVal) {
            builder.element('pivotField', attributes: {
              'dataField': '1',
              'showAllItems': '0',
            });
          } else {
            builder.element('pivotField', attributes: {
              'showAllItems': '0',
            });
          }
        }
      });

      if (pt.rows.isNotEmpty) {
        builder.element('rowFields',
            attributes: {'count': pt.rows.length.toString()}, nest: () {
          for (final rowName in pt.rows) {
            final idx = sourceHeaders.indexOf(rowName);
            if (idx != -1) {
              builder.element('field', attributes: {'x': idx.toString()});
            }
          }
        });

        builder.element('rowItems', attributes: {'count': '1'}, nest: () {
          builder.element('i');
        });
      }

      if (pt.columns.isNotEmpty) {
        builder.element('colFields',
            attributes: {'count': pt.columns.length.toString()}, nest: () {
          for (final colName in pt.columns) {
            final idx = sourceHeaders.indexOf(colName);
            if (idx != -1) {
              builder.element('field', attributes: {'x': idx.toString()});
            }
          }
        });

        builder.element('colItems', attributes: {'count': '1'}, nest: () {
          builder.element('i');
        });
      }

      if (pt.values.isNotEmpty) {
        builder.element('dataFields',
            attributes: {'count': pt.values.length.toString()}, nest: () {
          for (final val in pt.values) {
            final idx = sourceHeaders.indexOf(val.field);
            if (idx != -1) {
              final name = val.customName ??
                  _defaultDataFieldName(val.field, val.function);
              builder.element('dataField', attributes: {
                'name': name,
                'fld': idx.toString(),
                'subtotal': val.function.name,
                'baseField': '0',
                'baseItem': '0',
              });
            }
          }
        });
      }

      builder.element('pivotTableStyleInfo', attributes: {
        'name': 'PivotStyleLight16',
        'showRowHeaders': '1',
        'showColHeaders': '1',
        'showRowStripes': '0',
        'showColStripes': '0',
      });
    });

    return builder.buildDocument();
  }

  String _defaultDataFieldName(String field, PivotValueFunction function) {
    final functionName = function.name;
    final capitalized =
        functionName[0].toUpperCase() + functionName.substring(1);
    return '$capitalized of $field';
  }

  XmlDocument _buildPivotTableRelsXml(int pivotCacheCount) {
    final builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('Relationships', attributes: {
      'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
    }, nest: () {
      builder.element('Relationship', attributes: {
        'Id': 'rId1',
        'Type':
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
        'Target': '../pivotCache/pivotCacheDefinition$pivotCacheCount.xml',
      });
    });
    return builder.buildDocument();
  }

  XmlDocument _buildPivotCacheDefXml(
      PivotTable pt, List<String> sourceHeaders) {
    final builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('pivotCacheDefinition', attributes: {
      'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'r:id': 'rId1',
      'refreshOnLoad': '1',
      'createdVersion': '4',
      'updatedVersion': '4',
      'recordCount': '0',
    }, nest: () {
      builder.element('cacheSource', attributes: {'type': 'worksheet'},
          nest: () {
        builder.element('worksheetSource', attributes: {
          'ref': pt.sourceRange,
          'sheet': pt.sourceSheet,
        });
      });
      builder.element('cacheFields',
          attributes: {'count': sourceHeaders.length.toString()}, nest: () {
        for (final header in sourceHeaders) {
          builder.element('cacheField', attributes: {
            'name': header,
            'numFmtId': '0',
          });
        }
      });
    });
    return builder.buildDocument();
  }

  XmlDocument _buildPivotCacheDefRelsXml(int pivotCacheCount) {
    final builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('Relationships', attributes: {
      'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
    }, nest: () {
      builder.element('Relationship', attributes: {
        'Id': 'rId1',
        'Type':
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
        'Target': 'pivotCacheRecords$pivotCacheCount.xml',
      });
    });
    return builder.buildDocument();
  }

  XmlDocument _buildPivotCacheRecordsXml() {
    final builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('pivotCacheRecords', attributes: {
      'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'count': '0',
    });
    return builder.buildDocument();
  }

  int _getAvailableWorkbookRId(XmlElement relsRoot) {
    int maxId = 0;
    for (final rel in relsRoot.findAllElements('Relationship')) {
      final idAttr = rel.getAttribute('Id');
      if (idAttr != null && idAttr.startsWith('rId')) {
        final numPart = int.tryParse(idAttr.substring(3));
        if (numPart != null && numPart > maxId) {
          maxId = numPart;
        }
      }
    }
    return maxId + 1;
  }

  void _addPivotCacheToWorkbookXml(int cacheId, String rId) {
    final doc = _excel._xmlFiles['xl/workbook.xml'];
    if (doc == null) return;

    final workbook = doc.findAllElements('workbook').first;
    final existingCaches = workbook.findAllElements('pivotCaches');
    XmlElement pivotCaches;

    if (existingCaches.isNotEmpty) {
      pivotCaches = existingCaches.first;
    } else {
      pivotCaches = XmlElement(XmlName.parts('pivotCaches'));

      final children = workbook.children;
      int insertIndex = children.length;

      const order = [
        'fileVersion',
        'workbookPr',
        'workbookProtection',
        'bookViews',
        'sheets',
        'functionGroups',
        'externalReferences',
        'definedNames',
        'calcPr',
        'oleSize',
        'customWorkbookViews',
        'pivotCaches',
        'smartTagPr',
        'smartTagTypes',
        'webPublishing',
        'fileSharing',
        'webPublishObjects',
        'extLst'
      ];

      for (int i = 0; i < children.length; i++) {
        final child = children[i];
        if (child is XmlElement) {
          final childName = child.name.local;
          final orderIdx = order.indexOf(childName);
          final pivotIdx = order.indexOf('pivotCaches');
          if (orderIdx > pivotIdx) {
            insertIndex = i;
            break;
          }
        }
      }

      children.insert(insertIndex, pivotCaches);
    }

    pivotCaches.children.add(XmlElement(XmlName.parts('pivotCache'), [
      XmlAttribute(XmlName.parts('cacheId'), cacheId.toString()),
      XmlAttribute(XmlName.parts('r:id'), rId),
    ]));
  }
}
