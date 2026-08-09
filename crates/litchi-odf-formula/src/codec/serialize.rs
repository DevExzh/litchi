//! `MathML` tree serialization back to well-formed XML.
//!
//! The parser expands namespace prefixes and discards `xmlns` declarations,
//! so serialization reconstructs a self-contained document: MathML-namespace
//! elements use the default namespace declared on the subtree root, while
//! foreign namespaces (vendor extensions, content `MathML` inside
//! `annotation-xml`) receive generated `ns1..nsN` prefixes in first-use
//! order.
//!
//! Entity references that the parser retained verbatim as `&name;` text
//! (non-predefined entities from an unevaluated document type definition) are
//! serialized as the escaped form `&amp;name;`: the result is always
//! well-formed without a DTD, at the cost of no longer spelling the original
//! reference literally.

use crate::model::{Content, Element, MATHML_NAMESPACE};

/// Generated prefixes for foreign namespaces, in first-use document order.
#[derive(Default)]
struct NamespaceMap {
    prefixes: Vec<(String, String)>,
}

enum Write<'element> {
    Element(&'element Element, bool),
    Text(&'element str),
    End(String),
}

impl NamespaceMap {
    fn collect(&mut self, root: &Element) {
        let mut pending = vec![root];
        while let Some(element) = pending.pop() {
            if let Some(uri) = element.namespace_uri() {
                self.register(uri);
            }
            for attribute in element.attributes() {
                if let Some(uri) = attribute.namespace_uri() {
                    self.register(uri);
                }
            }
            for content in element.content().iter().rev() {
                if let Content::Element(child) = content {
                    pending.push(child);
                }
            }
        }
    }

    fn register(&mut self, uri: &str) {
        if uri == MATHML_NAMESPACE || self.prefixes.iter().any(|(known, _)| known == uri) {
            return;
        }
        let prefix = format!("ns{}", self.prefixes.len() + 1);
        self.prefixes.push((uri.to_string(), prefix));
    }

    fn qualify<'a>(&'a self, namespace_uri: Option<&'a str>, local_name: &'a str) -> String {
        match namespace_uri {
            Some(uri) if uri != MATHML_NAMESPACE => self
                .prefixes
                .iter()
                .find(|(known, _)| known == uri)
                .map_or_else(
                    || local_name.to_string(),
                    |(_, prefix)| format!("{prefix}:{local_name}"),
                ),
            _ => local_name.to_string(),
        }
    }
}

/// Serialize a `MathML` subtree to a well-formed, self-contained XML string.
#[must_use]
pub fn serialize(root: &Element) -> String {
    write_mathml(root)
}

pub(crate) fn write_mathml(root: &Element) -> String {
    let mut namespaces = NamespaceMap::default();
    namespaces.collect(root);
    let mut output = String::new();
    let mut pending = vec![Write::Element(root, true)];
    while let Some(write) = pending.pop() {
        match write {
            Write::Element(element, is_root) => {
                let name = namespaces.qualify(element.namespace_uri(), element.local_name());
                write_start(element, is_root, &name, &namespaces, &mut output);
                if element.content().is_empty() {
                    output.push_str("/>");
                    continue;
                }
                output.push('>');
                pending.push(Write::End(name));
                for content in element.content().iter().rev() {
                    pending.push(match content {
                        Content::Text(text) => Write::Text(text),
                        Content::Element(child) => Write::Element(child, false),
                    });
                }
            },
            Write::Text(value) => escape_text(value, &mut output),
            Write::End(name) => {
                output.push_str("</");
                output.push_str(&name);
                output.push('>');
            },
        }
    }
    output
}

fn write_start(
    element: &Element,
    root: bool,
    name: &str,
    namespaces: &NamespaceMap,
    output: &mut String,
) {
    output.push('<');
    output.push_str(name);
    if root {
        if element.namespace_uri() == Some(MATHML_NAMESPACE) {
            output.push_str(" xmlns=\"");
            output.push_str(MATHML_NAMESPACE);
            output.push('"');
        }
        for (uri, prefix) in &namespaces.prefixes {
            output.push_str(" xmlns:");
            output.push_str(prefix);
            output.push_str("=\"");
            escape_attribute(uri, output);
            output.push('"');
        }
    }
    for attribute in element.attributes() {
        let attribute_name = namespaces.qualify(attribute.namespace_uri(), attribute.local_name());
        output.push(' ');
        output.push_str(&attribute_name);
        output.push_str("=\"");
        escape_attribute(attribute.value(), output);
        output.push('"');
    }
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
