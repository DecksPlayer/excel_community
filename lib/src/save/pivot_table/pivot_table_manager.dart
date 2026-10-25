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
      final sheetRels = _save._relationshipsPart(sheetRelsPath);
      final sheetRelsRoot = sheetRels.findAllElements('Relationships').first;

      for (final pt in sheet.pivotTables) {
        pivotTableCount++;
        pivotCacheCount++;

        final sourceSheetObj = _excel._sheetMap[pt.sourceSheet];
        if (sourceSheetObj == null) continue;

        final layout = _PivotLayout.build(pt, sourceSheetObj);
        if (layout == null) continue;

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
            _buildPivotTableXml(layout, pivotCacheCount);

        // 2. Generate Pivot Table Rels
        _excel._xmlFiles[pivotTableRelsPath] =
            _buildPivotTableRelsXml(pivotCacheCount);

        // 3. Generate Pivot Cache Definition XML
        _excel._xmlFiles[pivotCacheDefPath] = _buildPivotCacheDefXml(layout);

        // 4. Generate Pivot Cache Rels
        _excel._xmlFiles[pivotCacheDefRelsPath] =
            _buildPivotCacheDefRelsXml(pivotCacheCount);

        // 5. Generate Pivot Cache Records XML
        _excel._xmlFiles[pivotCacheRecPath] = _buildPivotCacheRecordsXml(layout);

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
        final ptRId = _save._nextRelationshipId(sheetRelsRoot);

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
    for (final f in _excel._archive.files) {
      keys.add(f.name);
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
    for (final f in _excel._archive.files) {
      keys.add(f.name);
    }
    return keys
        .where((k) =>
            k.startsWith('xl/pivotCache/pivotCacheDefinition') &&
            k.endsWith('.xml') &&
            !k.contains('/_rels/'))
        .length;
  }

  /// Writes the computed pivot tables into their target cells, like Excel
  /// does, so the values show even where the table is not refreshed
  /// (Protected View, other spreadsheet apps). Runs before styles are
  /// written.
  void renderPivotTables() {
    _excel._sheetMap.forEach((sheetName, sheet) {
      for (final pt in sheet.pivotTables) {
        render(sheet, pt);
      }
    });
  }

  /// Writes [pt] into [sheet], replacing what an earlier render wrote.
  static void render(Sheet sheet, PivotTable pt) {
    final source = sheet._excel._sheetMap[pt.sourceSheet];
    final layout = source == null ? null : _PivotLayout.build(pt, source);
    final top = pt.targetCell.rowIndex;
    final left = pt.targetCell.columnIndex;

    // Clear what an earlier render wrote for this table.
    for (final (r, c) in _rendered[pt] ?? const <(int, int)>[]) {
      sheet._sheetData[r]?.remove(c);
    }
    if (layout == null) return;

    final written = <(int, int)>[];
    layout.cells().forEach((pos, entry) {
      final (value, general) = entry;
      final index = CellIndex.indexByColumnRow(
          columnIndex: left + pos.$2, rowIndex: top + pos.$1);
      sheet.updateCell(index, value);
      if (general) {
        sheet.cell(index).cellStyle = (sheet.cell(index).cellStyle ?? CellStyle())
            .copyWith(numberFormat: NumFormat.standard_0);
      }
      written.add((index.rowIndex, index.columnIndex));
    });
    _rendered[pt] = written;
  }

  static final _rendered = Expando<List<(int, int)>>('pivotCells');

  XmlDocument _buildPivotTableXml(_PivotLayout layout, int cacheId) {
    final pt = layout.pivot;
    final builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');

    void writeItems(String name, List<_PivotAxisEntry> entries,
        {required bool emptyWhenSingle}) {
      builder.element(name, attributes: {'count': entries.length.toString()},
          nest: () {
        for (final e in entries) {
          if (emptyWhenSingle && e.path.isEmpty) {
            builder.element('i');
            continue;
          }
          final attributes = <String, String>{
            if (e.type != 'data') 't': e.type,
            if (e.repeated > 0) 'r': e.repeated.toString(),
            if (e.dataIndex > 0) 'i': e.dataIndex.toString(),
          };
          builder.element('i', attributes: attributes, nest: () {
            final xs = e.type == 'grand' ? const [0] : e.path.sublist(e.repeated);
            for (final v in xs) {
              builder.element('x', attributes: {if (v != 0) 'v': v.toString()});
            }
          });
        }
      });
    }

    builder.element('pivotTableDefinition',
        attributes: {
          'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          'xmlns:r':
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'name': pt.name,
          'cacheId': cacheId.toString(),
          'applyNumberFormats': '0',
          'applyBorderFormats': '0',
          'applyFontFormats': '0',
          'applyPatternFormats': '0',
          'applyAlignmentFormats': '0',
          'applyWidthHeightFormats': '1',
          'dataCaption': 'Values',
          'updatedVersion': '4',
          'createdVersion': '4',
          'minRefreshableVersion': '3',
          'rowHeaderCaption': 'Row Labels',
          'colHeaderCaption': 'Column Labels',
        }, nest: () {
      builder.element('location', attributes: {
        'ref': layout.locationRef,
        'firstHeaderRow': '1',
        'firstDataRow': layout.headerRows.toString(),
        'firstDataCol': layout.firstDataCol.toString(),
      });

      builder.element('pivotFields',
          attributes: {'count': layout.headers.length.toString()}, nest: () {
        for (int i = 0; i < layout.headers.length; i++) {
          final axis = layout.rowFields.contains(i)
              ? 'axisRow'
              : layout.colFields.contains(i)
                  ? 'axisCol'
                  : null;
          builder.element('pivotField', attributes: {
            if (axis != null) 'axis': axis,
            if (layout.dataFieldIndexes.contains(i)) 'dataField': '1',
            'showAll': '0',
          }, nest: () {
            if (axis == null) return;
            final count = layout.sharedItems[i]!.length;
            builder.element('items',
                attributes: {'count': (count + 1).toString()}, nest: () {
              for (var x = 0; x < count; x++) {
                builder.element('item', attributes: {'x': x.toString()});
              }
              builder.element('item', attributes: {'t': 'default'});
            });
          });
        }
      });

      if (layout.rowFields.isNotEmpty) {
        builder.element('rowFields',
            attributes: {'count': layout.rowFields.length.toString()}, nest: () {
          for (final idx in layout.rowFields) {
            builder.element('field', attributes: {'x': idx.toString()});
          }
        });
      }
      writeItems('rowItems', layout.rowEntries,
          emptyWhenSingle: layout.rowFields.isEmpty);

      // With several data fields Excel needs the virtual "Values" field
      // (x = -2) among the column fields, or it cannot open the file.
      final colFields = [
        ...layout.colFields,
        if (layout.dataValues.length > 1) -2,
      ];
      if (colFields.isNotEmpty) {
        builder.element('colFields',
            attributes: {'count': colFields.length.toString()}, nest: () {
          for (final idx in colFields) {
            builder.element('field', attributes: {'x': idx.toString()});
          }
        });
      }
      writeItems('colItems', layout.colEntries,
          emptyWhenSingle: colFields.isEmpty);

      if (layout.dataValues.isNotEmpty) {
        builder.element('dataFields',
            attributes: {'count': layout.dataValues.length.toString()}, nest: () {
          for (var d = 0; d < layout.dataValues.length; d++) {
            final val = layout.dataValues[d];
            builder.element('dataField', attributes: {
              'name': layout.dataNames[d],
              'fld': layout.dataFieldIndexes[d].toString(),
              if (val.function != PivotValueFunction.sum)
                'subtotal': val.function == PivotValueFunction.varVal
                    ? 'var'
                    : val.function.name,
              'baseField': '0',
              'baseItem': '0',
            });
          }
        });
      }

      builder.element('pivotTableStyleInfo', attributes: {
        'name': 'PivotStyleLight16',
        'showRowHeaders': '1',
        'showColHeaders': '1',
        'showRowStripes': '0',
        'showColStripes': '0',
        'showLastColumn': '1',
      });
    });

    return builder.buildDocument();
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

  XmlDocument _buildPivotCacheDefXml(_PivotLayout layout) {
    final pt = layout.pivot;
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
      'refreshedVersion': '4',
      'minRefreshableVersion': '3',
      'recordCount': layout.records.length.toString(),
    }, nest: () {
      builder.element('cacheSource', attributes: {'type': 'worksheet'},
          nest: () {
        builder.element('worksheetSource', attributes: {
          'ref': pt.sourceRange,
          'sheet': pt.sourceSheet,
        });
      });
      builder.element('cacheFields',
          attributes: {'count': layout.headers.length.toString()}, nest: () {
        for (var i = 0; i < layout.headers.length; i++) {
          final shared = layout.sharedItems[i];
          final values = shared ?? [for (final r in layout.records) r[i]];
          builder.element('cacheField', attributes: {
            'name': layout.headers[i],
            'numFmtId': '0',
          }, nest: () {
            builder.element('sharedItems', attributes: {
              ..._sharedItemsAttributes(values),
              if (shared != null) 'count': shared.length.toString(),
            }, nest: () {
              for (final v in shared ?? const <_PivotCacheValue>[]) {
                _writeCacheValue(builder, v);
              }
            });
          });
        }
      });
    });
    return builder.buildDocument();
  }

  /// Type flags of a cache field, as Excel writes them.
  Map<String, String> _sharedItemsAttributes(List<_PivotCacheValue> values) {
    final kinds = values.map((v) => v.kind).toSet();
    final hasString = kinds.contains('s');
    final hasNumber = kinds.contains('n');
    final hasDate = kinds.contains('d');
    final hasBlank = kinds.contains('m');
    final numbers = [for (final v in values) if (v.kind == 'n') v.number!];
    final dates = [for (final v in values) if (v.kind == 'd') v.text]..sort();
    return {
      if (!hasString && !hasBlank) 'containsSemiMixedTypes': '0',
      if (hasDate && !hasString && !hasNumber && !hasBlank) 'containsNonDate': '0',
      if (hasDate) 'containsDate': '1',
      if (!hasString) 'containsString': '0',
      if (hasBlank) 'containsBlank': '1',
      if ([hasString, hasNumber, hasDate].where((t) => t).length > 1)
        'containsMixedTypes': '1',
      if (hasNumber) 'containsNumber': '1',
      if (hasNumber && numbers.every((n) => n == n.roundToDouble()))
        'containsInteger': '1',
      if (hasNumber) 'minValue': _PivotLayout._formatNumber(numbers.reduce(min)),
      if (hasNumber) 'maxValue': _PivotLayout._formatNumber(numbers.reduce(max)),
      if (hasDate) 'minDate': dates.first,
      if (hasDate) 'maxDate': dates.last,
    };
  }

  void _writeCacheValue(XmlBuilder builder, _PivotCacheValue v) {
    if (v.kind == 'm') {
      builder.element('m');
    } else {
      builder.element(v.kind, attributes: {'v': v.text});
    }
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

  XmlDocument _buildPivotCacheRecordsXml(_PivotLayout layout) {
    final builder = XmlBuilder();
    builder.processing(
        'xml', 'version="1.0" encoding="UTF-8" standalone="yes"');
    builder.element('pivotCacheRecords', attributes: {
      'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'count': layout.records.length.toString(),
    }, nest: () {
      for (var r = 0; r < layout.records.length; r++) {
        builder.element('r', nest: () {
          for (var i = 0; i < layout.headers.length; i++) {
            final items = layout.recordItems[i];
            if (items != null) {
              builder.element('x', attributes: {'v': items[r].toString()});
            } else {
              _writeCacheValue(builder, layout.records[r][i]);
            }
          }
        });
      }
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
