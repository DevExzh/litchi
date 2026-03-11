//! Base Element class for ODF XML elements.
//!
//! This module provides the fundamental Element class that all ODF elements
//! inherit from, providing common functionality for XML manipulation.

use crate::common::{Error, Result};
use crate::odf::elements::namespace::{NamespaceContext, QualifiedName};
use quick_xml::events::Event;
use std::collections::HashMap;

/// Property definition for element attributes
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct PropDef {
    pub name: String,
    pub attr: String,
    pub family: String,
}

#[allow(dead_code)]
impl PropDef {
    pub fn new(name: &str, attr: &str) -> Self {
        Self {
            name: name.to_string(),
            attr: attr.to_string(),
            family: String::new(),
        }
    }

    pub fn with_family(name: &str, attr: &str, family: &str) -> Self {
        Self {
            name: name.to_string(),
            attr: attr.to_string(),
            family: family.to_string(),
        }
    }
}

/// Base trait for all ODF elements
#[allow(dead_code)]
pub trait ElementBase {
    /// Get the tag name of this element
    fn tag_name(&self) -> &str;

    /// Get the attributes of this element
    fn attributes(&self) -> &HashMap<String, String>;

    /// Get a mutable reference to attributes
    fn attributes_mut(&mut self) -> &mut HashMap<String, String>;

    /// Get the text content of this element
    fn text(&self) -> &str;

    /// Set the text content of this element
    fn set_text(&mut self, text: &str);

    /// Get attribute value by name
    fn get_attribute(&self, name: &str) -> Option<&str> {
        self.attributes().get(name).map(|s| s.as_str())
    }

    /// Set attribute value
    fn set_attribute(&mut self, name: &str, value: &str) {
        self.attributes_mut()
            .insert(name.to_string(), value.to_string());
    }

    /// Remove attribute
    fn remove_attribute(&mut self, name: &str) {
        self.attributes_mut().remove(name);
    }

    /// Check if element has attribute
    fn has_attribute(&self, name: &str) -> bool {
        self.attributes().contains_key(name)
    }

    /// Get boolean attribute value
    fn get_bool_attribute(&self, name: &str) -> Option<bool> {
        self.get_attribute(name).and_then(|s| match s {
            "true" | "1" => Some(true),
            "false" | "0" => Some(false),
            _ => None,
        })
    }

    /// Get numeric attribute value
    fn get_numeric_attribute(&self, name: &str) -> Option<f64> {
        self.get_attribute(name).and_then(|s| s.parse().ok())
    }

    /// Get integer attribute value
    fn get_int_attribute(&self, name: &str) -> Option<i64> {
        self.get_attribute(name).and_then(|s| s.parse().ok())
    }
}

/// Concrete Element implementation with namespace support
#[derive(Debug, Clone)]
pub struct Element {
    tag_name: String,
    qualified_name: QualifiedName,
    attributes: HashMap<String, String>,
    namespace_context: NamespaceContext,
    text_content: String,
    pub(crate) children: Vec<Element>,
}

impl Element {
    /// Add a child element (concrete Element type)
    pub fn add_child(&mut self, child: Element) {
        self.children.push(child);
    }

    /// Get children as concrete Elements
    pub fn get_children(&self) -> &[Element] {
        &self.children
    }

    /// Get mutable children as concrete Elements  
    pub fn get_children_mut(&mut self) -> &mut Vec<Element> {
        &mut self.children
    }

    /// Get text recursively from this element and all children
    pub fn get_text_recursive(&self) -> String {
        let mut text = self.text_content.clone();
        for child in &self.children {
            text.push_str(&child.get_text_recursive());
        }
        text
    }
}

impl Element {
    /// Create a new element
    pub fn new(tag_name: &str) -> Self {
        let qualified_name = QualifiedName::from_string(tag_name);
        Self {
            tag_name: tag_name.to_string(),
            qualified_name,
            attributes: HashMap::new(),
            namespace_context: NamespaceContext::default(),
            text_content: String::new(),
            children: Vec::new(),
        }
    }

    /// Create a new element with namespace context
    pub fn new_with_context(tag_name: &str, namespace_context: NamespaceContext) -> Self {
        let qualified_name = namespace_context.parse_qualified_name(tag_name);
        Self {
            tag_name: tag_name.to_string(),
            qualified_name,
            attributes: HashMap::new(),
            namespace_context,
            text_content: String::new(),
            children: Vec::new(),
        }
    }

    /// Get the qualified name
    pub fn qualified_name(&self) -> &QualifiedName {
        &self.qualified_name
    }

    /// Get the namespace URI
    pub fn namespace_uri(&self) -> Option<&str> {
        self.qualified_name.namespace_uri.as_deref()
    }

    /// Get the local name (without namespace prefix)
    pub fn local_name(&self) -> &str {
        &self.qualified_name.local_name
    }

    /// Get the namespace context
    pub fn namespace_context(&self) -> &NamespaceContext {
        &self.namespace_context
    }

    /// Set namespace context
    pub fn set_namespace_context(&mut self, context: NamespaceContext) {
        self.namespace_context = context;
        // Re-parse qualified name with new context
        self.qualified_name = self.namespace_context.parse_qualified_name(&self.tag_name);
    }

    /// Add a namespace declaration
    pub fn add_namespace(&mut self, prefix: &str, uri: &str) {
        self.namespace_context.add_namespace(prefix, uri);
        // Re-parse qualified name with updated context
        self.qualified_name = self.namespace_context.parse_qualified_name(&self.tag_name);
    }

    /// Check if element name matches (namespace-aware)
    pub fn name_matches(&self, name: &str) -> bool {
        self.qualified_name
            .matches_str(name, Some(&self.namespace_context))
    }

    /// Get attribute with namespace-aware lookup
    pub fn get_qualified_attribute(&self, name: &str) -> Option<&str> {
        // First try exact match
        if let Some(value) = self.attributes.get(name) {
            return Some(value);
        }

        // Try namespace-aware match
        let qualified_name = self.namespace_context.parse_qualified_name(name);
        for (key, value) in &self.attributes {
            let key_qualified = self.namespace_context.parse_qualified_name(key);
            if key_qualified.matches(&qualified_name) {
                return Some(value);
            }
        }

        None
    }

    /// Create element from XML bytes
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let mut reader = quick_xml::Reader::from_reader(bytes);
        let mut buf = Vec::new();
        let mut stack = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let tag_name = String::from_utf8(e.name().as_ref().to_vec()).map_err(|_| {
                        Error::InvalidFormat("Invalid UTF-8 in tag name".to_string())
                    })?;

                    let mut namespace_context = NamespaceContext::default();

                    // First pass: collect namespace declarations
                    for attr_result in e.attributes() {
                        let attr = attr_result
                            .map_err(|_| Error::InvalidFormat("Invalid attribute".to_string()))?;
                        let key = String::from_utf8(attr.key.as_ref().to_vec()).map_err(|_| {
                            Error::InvalidFormat("Invalid UTF-8 in attribute key".to_string())
                        })?;
                        let value = String::from_utf8(attr.value.to_vec()).map_err(|_| {
                            Error::InvalidFormat("Invalid UTF-8 in attribute value".to_string())
                        })?;

                        // Check for namespace declarations
                        if key == "xmlns" || key.starts_with("xmlns:") {
                            namespace_context.add_namespace(&key, &value);
                        }
                    }

                    let mut element = Element::new_with_context(&tag_name, namespace_context);

                    // Second pass: set regular attributes
                    for attr_result in e.attributes() {
                        let attr = attr_result
                            .map_err(|_| Error::InvalidFormat("Invalid attribute".to_string()))?;
                        let key = String::from_utf8(attr.key.as_ref().to_vec()).map_err(|_| {
                            Error::InvalidFormat("Invalid UTF-8 in attribute key".to_string())
                        })?;
                        let value = String::from_utf8(attr.value.to_vec()).map_err(|_| {
                            Error::InvalidFormat("Invalid UTF-8 in attribute value".to_string())
                        })?;

                        // Skip namespace declarations - they're already handled
                        if !(key == "xmlns" || key.starts_with("xmlns:")) {
                            element.set_attribute(&key, &value);
                        }
                    }

                    stack.push(element);
                },
                Ok(Event::Text(ref t)) => {
                    if let Some(current) = stack.last_mut() {
                        let text = String::from_utf8(t.to_vec()).map_err(|_| {
                            Error::InvalidFormat("Invalid UTF-8 in text content".to_string())
                        })?;
                        current.text_content.push_str(&text);
                    }
                },
                Ok(Event::End(ref e)) => {
                    let _tag_name = String::from_utf8(e.name().as_ref().to_vec()) // Tag name for debugging - kept for future use
                        .map_err(|_| {
                            Error::InvalidFormat("Invalid UTF-8 in tag name".to_string())
                        })?;

                    if let Some(element) = stack.pop() {
                        if let Some(parent) = stack.last_mut() {
                            parent.children.push(element);
                        } else {
                            // This is the root element
                            return Ok(element);
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::InvalidFormat(format!("XML parsing error: {}", e))),
                _ => {},
            }
            buf.clear();
        }

        Err(Error::InvalidFormat("No root element found".to_string()))
    }

    /// Serialize element to XML string
    pub fn to_xml_string(&self) -> String {
        let mut xml = String::with_capacity(self.estimated_xml_len());
        self.write_xml(&mut xml, 0);
        xml
    }

    fn estimated_xml_len(&self) -> usize {
        let mut len = 0usize;

        len += 1 + self.tag_name.len();

        for (key, value) in &self.attributes {
            len += 1 + key.len();
            len += 2;
            len += value.len() + 8;
            len += 1;
        }

        if self.children.is_empty() && self.text_content.is_empty() {
            len += 3;
            return len;
        }

        len += 1;

        if !self.text_content.is_empty() {
            len += self.text_content.len() + 8;
        }

        for child in &self.children {
            len += child.estimated_xml_len();
        }

        len += 3 + self.tag_name.len() + 1;
        len
    }

    fn write_xml(&self, output: &mut String, indent: usize) {
        let _ = indent;

        // Opening tag
        output.push('<');
        output.push_str(&self.tag_name);

        // Attributes
        for (key, value) in &self.attributes {
            output.push(' ');
            output.push_str(key);
            output.push_str("=\"");
            // Escape quotes in attribute values
            for ch in value.chars() {
                match ch {
                    '"' => output.push_str("&quot;"),
                    '&' => output.push_str("&amp;"),
                    '<' => output.push_str("&lt;"),
                    '>' => output.push_str("&gt;"),
                    _ => output.push(ch),
                }
            }
            output.push('"');
        }

        if self.children.is_empty() && self.text_content.is_empty() {
            // Self-closing tag
            output.push_str(" />");
        } else {
            output.push('>');

            // Text content
            if !self.text_content.is_empty() {
                // Escape text content
                for ch in self.text_content.chars() {
                    match ch {
                        '&' => output.push_str("&amp;"),
                        '<' => output.push_str("&lt;"),
                        '>' => output.push_str("&gt;"),
                        _ => output.push(ch),
                    }
                }
            }

            // Child elements
            for child in &self.children {
                child.write_xml(output, indent + 1);
            }

            // Closing tag
            output.push_str("</");
            output.push_str(&self.tag_name);
            output.push('>');
        }
    }
}

impl ElementBase for Element {
    fn tag_name(&self) -> &str {
        &self.tag_name
    }

    fn attributes(&self) -> &HashMap<String, String> {
        &self.attributes
    }

    fn attributes_mut(&mut self) -> &mut HashMap<String, String> {
        &mut self.attributes
    }

    fn text(&self) -> &str {
        &self.text_content
    }

    fn set_text(&mut self, text: &str) {
        self.text_content = text.to_string();
    }
}

/// Helper for creating elements with specific tag names
#[allow(dead_code)]
pub struct ElementFactory;

#[allow(dead_code)]
impl ElementFactory {
    /// Create a text paragraph element
    pub fn paragraph() -> Element {
        Element::new("text:p")
    }

    /// Create a text span element
    pub fn span() -> Element {
        Element::new("text:span")
    }

    /// Create a heading element
    pub fn heading(level: u8) -> Element {
        let mut element = Element::new("text:h");
        element.set_attribute("text:outline-level", &level.to_string());
        element
    }

    /// Create a table element
    pub fn table() -> Element {
        Element::new("table:table")
    }

    /// Create a table row element
    pub fn table_row() -> Element {
        Element::new("table:table-row")
    }

    /// Create a table cell element
    pub fn table_cell() -> Element {
        Element::new("table:table-cell")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_element_new() {
        let element = Element::new("text:p");
        assert_eq!(element.tag_name(), "text:p");
        assert_eq!(element.local_name(), "p");
        assert_eq!(element.text(), "");
        assert!(element.attributes().is_empty());
    }

    #[test]
    fn test_element_with_attributes() {
        let mut element = Element::new("text:p");
        element.set_attribute("class", "body");
        element.set_attribute("id", "para1");

        assert_eq!(element.get_attribute("class"), Some("body"));
        assert_eq!(element.get_attribute("id"), Some("para1"));
        assert_eq!(element.get_attribute("missing"), None);

        assert!(element.has_attribute("class"));
        assert!(!element.has_attribute("missing"));
    }

    #[test]
    fn test_element_remove_attribute() {
        let mut element = Element::new("text:p");
        element.set_attribute("class", "body");
        assert!(element.has_attribute("class"));

        element.remove_attribute("class");
        assert!(!element.has_attribute("class"));
    }

    #[test]
    fn test_element_text() {
        let mut element = Element::new("text:p");
        element.set_text("Hello World");
        assert_eq!(element.text(), "Hello World");
    }

    #[test]
    fn test_element_children() {
        let mut parent = Element::new("text:p");
        let child1 = Element::new("text:span");
        let child2 = Element::new("text:span");

        parent.add_child(child1);
        parent.add_child(child2);

        let children = parent.get_children();
        assert_eq!(children.len(), 2);
        assert_eq!(children[0].tag_name(), "text:span");
    }

    #[test]
    fn test_element_get_text_recursive() {
        let mut parent = Element::new("text:p");
        parent.set_text("Hello ");

        let mut child = Element::new("text:span");
        child.set_text("World");

        parent.add_child(child);

        assert_eq!(parent.get_text_recursive(), "Hello World");
    }

    #[test]
    fn test_element_bool_attribute() {
        let mut element = Element::new("text:p");
        element.set_attribute("bold", "true");
        element.set_attribute("italic", "1");
        element.set_attribute("underline", "false");

        assert_eq!(element.get_bool_attribute("bold"), Some(true));
        assert_eq!(element.get_bool_attribute("italic"), Some(true));
        assert_eq!(element.get_bool_attribute("underline"), Some(false));
        assert_eq!(element.get_bool_attribute("missing"), None);
    }

    #[test]
    fn test_element_numeric_attribute() {
        let mut element = Element::new("table:table-cell");
        element.set_attribute("table:number-columns-spanned", "3");

        assert_eq!(
            element.get_numeric_attribute("table:number-columns-spanned"),
            Some(3.0)
        );
        assert_eq!(element.get_numeric_attribute("missing"), None);
    }

    #[test]
    fn test_element_int_attribute() {
        let mut element = Element::new("text:h");
        element.set_attribute("text:outline-level", "2");

        assert_eq!(element.get_int_attribute("text:outline-level"), Some(2));
        assert_eq!(element.get_int_attribute("missing"), None);
    }

    #[test]
    fn test_element_from_bytes() {
        let xml = b"<text:p class='body'>Hello World</text:p>";
        let element = Element::from_bytes(xml).unwrap();

        assert_eq!(element.tag_name(), "text:p");
        assert_eq!(element.get_attribute("class"), Some("body"));
        assert_eq!(element.text(), "Hello World");
    }

    #[test]
    fn test_element_from_bytes_with_children() {
        let xml = b"<text:p>Hello <text:span>World</text:span></text:p>";
        let element = Element::from_bytes(xml).unwrap();

        assert_eq!(element.tag_name(), "text:p");
        assert_eq!(element.text(), "Hello ");
        assert_eq!(element.get_children().len(), 1);
        assert_eq!(element.get_children()[0].tag_name(), "text:span");
    }

    #[test]
    fn test_element_to_xml_string() {
        let mut element = Element::new("text:p");
        element.set_attribute("class", "body");
        element.set_text("Hello World");

        let xml = element.to_xml_string();
        assert!(xml.contains("<text:p"));
        assert!(xml.contains('>'));
        assert!(xml.contains("class=\"body\""));
        assert!(xml.contains("Hello World"));
        assert!(xml.contains("</text:p>"));
    }

    #[test]
    fn test_element_to_xml_self_closing() {
        let element = Element::new("text:line-break");
        let xml = element.to_xml_string();
        assert!(xml.contains("/>"));
    }

    #[test]
    fn test_element_xml_escaping() {
        let mut element = Element::new("text:p");
        element.set_text("Hello <World> & \"Friends\"");

        let xml = element.to_xml_string();
        assert!(xml.contains("&lt;World&gt;"));
        assert!(xml.contains("&amp;"));
    }

    #[test]
    fn test_element_namespace_uri() {
        let xml =
            b"<text:p xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0'>Content</text:p>";
        let element = Element::from_bytes(xml).unwrap();
        assert_eq!(
            element.namespace_uri(),
            Some("urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        );
    }

    #[test]
    fn test_element_factory() {
        let para = ElementFactory::paragraph();
        assert_eq!(para.tag_name(), "text:p");

        let span = ElementFactory::span();
        assert_eq!(span.tag_name(), "text:span");

        let heading = ElementFactory::heading(2);
        assert_eq!(heading.tag_name(), "text:h");
        assert_eq!(heading.get_attribute("text:outline-level"), Some("2"));

        let table = ElementFactory::table();
        assert_eq!(table.tag_name(), "table:table");

        let row = ElementFactory::table_row();
        assert_eq!(row.tag_name(), "table:table-row");

        let cell = ElementFactory::table_cell();
        assert_eq!(cell.tag_name(), "table:table-cell");
    }
}
