//! MathML tree serialization back to well-formed XML.
//!
//! The parser expands namespace prefixes and discards `xmlns` declarations,
//! so serialization reconstructs a self-contained document: MathML-namespace
//! elements use the default namespace declared on the subtree root, while
//! foreign namespaces (vendor extensions, content MathML inside
//! `annotation-xml`) receive generated `ns1..nsN` prefixes in first-use
//! order.
//!
//! Entity references that the parser retained verbatim as `&name;` text
//! (non-predefined entities from an unevaluated document type definition) are
//! serialized as the escaped form `&amp;name;`: the result is always
//! well-formed without a DTD, at the cost of no longer spelling the original
//! reference literally.

use super::document::{MATHML_NAMESPACE, MathContent, MathElement};
use std::collections::HashMap;
use std::fmt::Write as _;

/// Serialize a MathML subtree to a well-formed, self-contained XML string.
pub(crate) fn write_mathml(root: &MathElement) -> String {
    let mut namespaces = NamespaceMap::default();
    namespaces.collect(root);
    let mut output = String::new();
    write_element(root, true, &namespaces, &mut output);
    output
}

/// Generated prefixes for foreign namespaces, in first-use document order.
#[derive(Default)]
struct NamespaceMap {
    prefixes: HashMap<String, String>,
}

impl NamespaceMap {
    fn collect(&mut self, element: &MathElement) {
        if let Some(uri) = element.namespace_uri() {
            self.register(uri);
        }
        for attribute in element.attributes() {
            if let Some(uri) = attribute.namespace_uri() {
                self.register(uri);
            }
        }
        for content in element.content() {
            if let MathContent::Element(child) = content {
                self.collect(child);
            }
        }
    }

    fn register(&mut self, uri: &str) {
        if uri == MATHML_NAMESPACE || self.prefixes.contains_key(uri) {
            return;
        }
        let prefix = format!("ns{}", self.prefixes.len() + 1);
        self.prefixes.insert(uri.to_string(), prefix);
    }

    fn qualify<'a>(&'a self, namespace_uri: Option<&'a str>, local_name: &'a str) -> String {
        match namespace_uri {
            Some(uri) if uri != MATHML_NAMESPACE => match self.prefixes.get(uri) {
                Some(prefix) => format!("{prefix}:{local_name}"),
                None => local_name.to_string(),
            },
            _ => local_name.to_string(),
        }
    }
}

fn write_element(
    element: &MathElement,
    root: bool,
    namespaces: &NamespaceMap,
    output: &mut String,
) {
    let name = namespaces.qualify(element.namespace_uri(), element.local_name());
    output.push('<');
    output.push_str(&name);
    if root {
        if element.namespace_uri() == Some(MATHML_NAMESPACE) {
            let _ = write!(output, " xmlns=\"{MATHML_NAMESPACE}\"");
        }
        for (uri, prefix) in &namespaces.prefixes {
            let _ = write!(output, " xmlns:{prefix}=\"");
            escape_attribute(uri, output);
            output.push('"');
        }
    }
    for attribute in element.attributes() {
        let name = namespaces.qualify(attribute.namespace_uri(), attribute.local_name());
        let _ = write!(output, " {name}=\"");
        escape_attribute(attribute.value(), output);
        output.push('"');
    }
    if element.content().is_empty() {
        output.push_str("/>");
        return;
    }
    output.push('>');
    for content in element.content() {
        match content {
            MathContent::Text(text) => escape_text(text, output),
            MathContent::Element(child) => write_element(child, false, namespaces, output),
        }
    }
    output.push_str("</");
    output.push_str(&name);
    output.push('>');
}

fn escape_text(text: &str, output: &mut String) {
    for character in text.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(character),
        }
    }
}

fn escape_attribute(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            // Attribute values are normalized on parse; keep serialization
            // stable by escaping literal whitespace that XML attribute-value
            // normalization would otherwise collapse on re-read.
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
            _ => output.push(character),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::super::document::parse_mathml;
    use super::*;

    #[test]
    fn serializes_parsed_tree_to_equivalent_xml() {
        let xml = r#"<math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><semantics><mrow><mi mathvariant="italic">f</mi><mo>(</mo><mfrac><mn>1</mn><mi>x</mi></mfrac><mtext> a &amp; &lt;b&gt; </mtext></mrow><annotation encoding="StarMath 5.0">f(x)</annotation></semantics></math>"#;
        let tree = parse_mathml(xml).unwrap();
        let serialized = tree.to_xml();
        let reparsed = parse_mathml(&serialized).unwrap();
        assert_eq!(tree, reparsed);
    }

    #[test]
    fn prefixes_foreign_namespaces_in_first_use_order() {
        let xml = r#"<math xmlns="http://www.w3.org/1998/Math/MathML" xmlns:v="urn:vendor:a" xmlns:w="urn:vendor:b"><v:hint v:mode="safe">x</v:hint><w:note/></math>"#;
        let tree = parse_mathml(xml).unwrap();
        let serialized = tree.to_xml();
        assert!(serialized.contains("xmlns:ns1=\"urn:vendor:a\""));
        assert!(serialized.contains("xmlns:ns2=\"urn:vendor:b\""));
        assert!(serialized.contains("<ns1:hint ns1:mode=\"safe\">"));
        assert!(serialized.contains("<ns2:note/>"));
        let reparsed = parse_mathml(&serialized).unwrap();
        assert_eq!(tree, reparsed);
    }

    #[test]
    fn escapes_text_and_attribute_values() {
        let mut root = MathElement::new("math").unwrap();
        let mut identifier = MathElement::new("mi").unwrap();
        identifier.set_attribute(None, "mathvariant", "bo<ld&\"x\"").unwrap();
        identifier.push_text("a<b>&c");
        root.push_child(identifier);
        let serialized = root.to_xml();
        assert!(serialized.contains("mathvariant=\"bo&lt;ld&amp;&quot;x&quot;\""));
        assert!(serialized.contains(">a&lt;b&gt;&amp;c</mi>"));
        let reparsed = parse_mathml(&serialized).unwrap();
        assert_eq!(reparsed, root);
    }
}
