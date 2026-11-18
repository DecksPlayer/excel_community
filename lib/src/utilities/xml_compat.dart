part of '../../excel_community.dart';

// Helpers that work on both `xml` 6.5 (the last release that supports Dart
// 3.6/3.7) and `xml` 7.x.
//
// `XmlName.parts`, `XmlName.qualified` and the `namespaceUris` parameter of
// `XmlBuilder.element` only exist in `xml` 7. The APIs used below exist in
// both versions; `xml` 7 marks them as deprecated, hence the `ignore`
// comments. Once the minimum SDK moves past what `xml` 6.5 supports, these
// helpers can be replaced by the `xml` 7 APIs directly.

/// Creates an [XmlName] from a qualified name such as `sheet` or `r:id`,
/// without a namespace URI.
///
/// Same result as `XmlName.parts(name)` / `XmlName.qualified(name)` in
/// `xml` 7.
XmlName _xmlName(String qualified) =>
    // ignore: deprecated_member_use
    XmlName.fromString(qualified);

extension _XmlBuilderCompat on XmlBuilder {
  /// Same as `element(name, namespaceUris: namespaceUris, ...)` in `xml` 7,
  /// where [namespaceUris] maps each prefix to its namespace URI.
  ///
  /// `xml` 6.5 only has the `namespaces` parameter, which maps each URI to its
  /// prefix; `xml` 7 still accepts it, as a deprecated alias.
  void elementWithNamespaceUris(
    String name, {
    required Map<String, String> namespaceUris,
    Map<String, String> attributes = const {},
    bool isSelfClosing = true,
    Object? nest,
  }) {
    final namespaces = <String, String?>{
      for (final entry in namespaceUris.entries) entry.value: entry.key,
    };
    element(
      name,
      // ignore: deprecated_member_use
      namespaces: namespaces,
      attributes: attributes,
      isSelfClosing: isSelfClosing,
      nest: nest,
    );
  }
}
