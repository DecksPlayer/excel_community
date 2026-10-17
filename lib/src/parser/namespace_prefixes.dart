part of '../../excel_community.dart';

/// Namespaces whose elements the parser and the writer look up by their
/// unprefixed name: SpreadsheetML, and the package's content types and
/// relationships.
const _unprefixedNamespaces = {
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://schemas.openxmlformats.org/package/2006/content-types',
  'http://schemas.openxmlformats.org/package/2006/relationships',
};

/// Rewrites in [archive] every XML part written with a namespace prefix, so
/// that the rest of the library reads it as a regular part.
///
/// Some writers (the Open XML SDK, for one) bind the namespace of a part to a
/// prefix and write every element with it:
/// `<x:worksheet xmlns:x="...main">...<x:sheetData>`. Excel opens such files,
/// but the parser and the writer look elements up by their unprefixed name, so
/// no sheet was found and the file was refused as damaged. A part whose root
/// element is prefixed in one of [_unprefixedNamespaces] is replaced by the
/// same document with those elements unprefixed and the namespace declared as
/// the default one. Elements of other namespaces keep their prefixes, and an
/// element that declares another default namespace is kept as it is, with
/// everything below it. Every other part is left untouched.
void _unprefixXmlParts(Archive archive) {
  final rewritten = <ArchiveFile>[];
  for (final file in archive.files) {
    if (!file.isFile) continue;
    final name = file.name.toLowerCase();
    if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
    final document = _prefixedPart(file);
    if (document == null) continue;
    final namespace = document.rootElement.name.namespaceUri!;
    final unprefixed = XmlDocument(
        document.children.map((node) => _unprefixedNode(node, namespace)));
    unprefixed.rootElement.setAttribute('xmlns', namespace);
    rewritten.add(
        ArchiveFile.bytes(file.name, utf8.encode(unprefixed.toXmlString())));
  }
  // Adding a file under the name of an existing one replaces it in place.
  rewritten.forEach(archive.addFile);
}

/// The document in [file] when its root element is prefixed in one of
/// [_unprefixedNamespaces], or `null` (also for a part that is not UTF-8 XML:
/// it is left for the parser to read, or refuse, as it is).
XmlDocument? _prefixedPart(ArchiveFile file) {
  final bytes = file.content;
  // Cheap check before parsing: the opening of the root element, after an
  // optional byte order mark, XML declaration and comments, has a colon.
  final head = utf8.decode(bytes.length > 2048 ? bytes.sublist(0, 2048) : bytes,
      allowMalformed: true);
  final root = RegExp(r'<([A-Za-z_][\w.-]*)(:)?').allMatches(head).where((m) =>
      !head.startsWith('<?', m.start) && !head.startsWith('<!', m.start));
  if (root.isEmpty || root.first.group(2) == null) return null;
  final XmlDocument document;
  try {
    var text = utf8.decode(bytes);
    if (text.startsWith('\uFEFF')) text = text.substring(1);
    document = XmlDocument.parse(text);
  } on FormatException {
    return null;
  } on XmlException {
    return null;
  }
  final name = document.rootElement.name;
  final namespace = name.namespaceUri;
  if (name.prefix == null ||
      namespace == null ||
      !_unprefixedNamespaces.contains(namespace)) {
    return null;
  }
  // The root's own default namespace, if any, would be overwritten.
  final declaredDefault = document.rootElement.getAttribute('xmlns');
  if (declaredDefault != null && declaredDefault != namespace) return null;
  return document;
}

/// A copy of [node] whose elements in [namespace] lose their prefix.
XmlNode _unprefixedNode(XmlNode node, String namespace) {
  if (node is! XmlElement) return node.copy();
  // Below another default namespace an unprefixed name means something else.
  final declaredDefault = node.getAttribute('xmlns');
  if (declaredDefault != null && declaredDefault != namespace) {
    return node.copy();
  }
  final name = node.name;
  return XmlElement(
    name.prefix != null && name.namespaceUri == namespace
        ? XmlName.parts(name.local)
        : XmlName.qualified(name.qualified),
    node.attributes.map((attribute) => attribute.copy()),
    node.children.map((child) => _unprefixedNode(child, namespace)),
    node.isSelfClosing,
  );
}
