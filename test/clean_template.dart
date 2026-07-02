import 'dart:convert';
import 'dart:io';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

void main() {
  final fileContent = File('lib/src/utilities/constants.dart').readAsStringSync();
  final regExp = RegExp(r"const _newSheet =\s+'([^']+)';");
  final match = regExp.firstMatch(fileContent);
  if (match == null) {
    print('Error: Could not find _newSheet');
    return;
  }
  
  final _newSheet = match.group(1)!.replaceAll('\n', '').replaceAll('\r', '').replaceAll(' ', '');
  final originalBytes = base64Decode(_newSheet);
  final archive = ZipDecoder().decodeBytes(originalBytes);
  final cleanArchive = Archive();

  for (final file in archive.files) {
    if (!file.isFile) continue;

    final name = file.name;
    if (name == 'xl/drawings/drawing1.xml' ||
        name == 'xl/worksheets/_rels/sheet1.xml.rels') {
      continue;
    }

    file.decompress();
    var content = file.content as List<int>;

    if (name == 'xl/worksheets/sheet1.xml') {
      final xmlString = utf8.decode(content);
      final document = XmlDocument.parse(xmlString);
      final drawingNodes = document.findAllElements('drawing').toList();
      for (var node in drawingNodes) {
        node.parent!.children.remove(node);
      }
      content = utf8.encode(document.toString());
    } else if (name == '[Content_Types].xml') {
      final xmlString = utf8.decode(content);
      final document = XmlDocument.parse(xmlString);
      final overrides = document.findAllElements('Override').toList();
      for (var node in overrides) {
        if (node.getAttribute('PartName') == '/xl/drawings/drawing1.xml') {
          node.parent!.children.remove(node);
        }
      }
      content = utf8.encode(document.toString());
    }

    cleanArchive.addFile(ArchiveFile(name, content.length, content));
  }

  final newBytes = ZipEncoder().encode(cleanArchive);
  final newBase64 = base64Encode(newBytes!);

  print('Cleaned length: ${newBase64.length}');
  print('Cleaned base64:');
  print(newBase64);
}
