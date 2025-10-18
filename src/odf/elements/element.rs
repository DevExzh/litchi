//! Base Element class for ODF XML elements.
//!
//! This module provides the fundamental Element class that all ODF elements
//! inherit from, providing common functionality for XML manipulation.

use crate::common::{Error, Result};
use crate::odf::elements::namespace::{QualifiedName, NamespaceContext};
use quick_xml::events::Event;
use std::collections::HashMap;

/// Property definition for element attributes
#[derive(Debug, Clone)]
pub struct PropDef {
    pub name: String,
    pub attr: String,
    pub family: String,
}

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

    /// Get child elements
    fn children(&self) -> &[Box<dyn ElementBase>];

    /// Get mutable child elements
    fn children_mut(&mut self) -> &mut Vec<Box<dyn ElementBase>>;

    /// Get attribute value by name
    fn get_attribute(&self, name: &str) -> Option<&str> {
        self.attributes().get(name).map(|s| s.as_str())
    }

    /// Set attribute value
    fn set_attribute(&mut self, name: &str, value: &str) {
        self.attributes_mut().insert(name.to_string(), value.to_string());
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
        self.get_attribute(name)
            .and_then(|s| match s {
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

    /// Find child elements by tag name
    fn get_elements_by_tag(&self, tag_name: &str) -> Vec<&dyn ElementBase> {
        self.children()
            .iter()
            .filter(|child| child.tag_name() == tag_name)
            .map(|child| child.as_ref())
            .collect()
    }

    /// Find first child element by tag name
    fn get_element_by_tag(&self, tag_name: &str) -> Option<&dyn ElementBase> {
        self.children()
            .iter()
            .find(|child| child.tag_name() == tag_name)
            .map(|child| child.as_ref())
    }

    /// Add child element
    fn add_child(&mut self, child: Box<dyn ElementBase>) {
        self.children_mut().push(child);
    }

    /// Remove child element at index
    fn remove_child(&mut self, index: usize) -> Option<Box<dyn ElementBase>> {
        if index < self.children().len() {
            Some(self.children_mut().remove(index))
        } else {
            None
        }
    }

    /// Get all text content recursively
    fn get_text_recursive(&self) -> String {
        let mut text = self.text().to_string();
        for child in self.children() {
            text.push_str(&child.get_text_recursive());
        }
        text
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
    children: Vec<Element>,
}

impl Element {
    /// Create a new element
    pub fn new(tag_name: &str) -> Self {
        let qualified_name = QualifiedName::from_string(tag_name);
        Self {
            tag_name: tag_name.to_string(),
            qualified_name,
            attributes: HashMap::new(),
            namespace_context: NamespaceContext::new(),
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
        self.qualified_name.matches_str(name, Some(&self.namespace_context))
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
                    let tag_name = String::from_utf8(e.name().as_ref().to_vec())
                        .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in tag name".to_string()))?;

                    let mut namespace_context = NamespaceContext::new();

                    // First pass: collect namespace declarations
                    for attr_result in e.attributes() {
                        let attr = attr_result.map_err(|_| Error::InvalidFormat("Invalid attribute".to_string()))?;
                        let key = String::from_utf8(attr.key.as_ref().to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in attribute key".to_string()))?;
                        let value = String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in attribute value".to_string()))?;

                        // Check for namespace declarations
                        if key == "xmlns" || key.starts_with("xmlns:") {
                            namespace_context.add_namespace(&key, &value);
                        }
                    }

                    let mut element = Element::new_with_context(&tag_name, namespace_context);

                    // Second pass: set regular attributes
                    for attr_result in e.attributes() {
                        let attr = attr_result.map_err(|_| Error::InvalidFormat("Invalid attribute".to_string()))?;
                        let key = String::from_utf8(attr.key.as_ref().to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in attribute key".to_string()))?;
                        let value = String::from_utf8(attr.value.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in attribute value".to_string()))?;

                        // Skip namespace declarations - they're already handled
                        if !(key == "xmlns" || key.starts_with("xmlns:")) {
                            element.set_attribute(&key, &value);
                        }
                    }

                    stack.push(element);
                }
                Ok(Event::Text(ref t)) => {
                    if let Some(current) = stack.last_mut() {
                        let text = String::from_utf8(t.to_vec())
                            .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in text content".to_string()))?;
                        current.text_content.push_str(&text);
                    }
                }
                Ok(Event::End(ref e)) => {
                    let _tag_name = String::from_utf8(e.name().as_ref().to_vec()) // Tag name for debugging - kept for future use
                        .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in tag name".to_string()))?;

                    if let Some(element) = stack.pop() {
                        if let Some(parent) = stack.last_mut() {
                            parent.children.push(element);
                        } else {
                            // This is the root element
                            return Ok(element);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::InvalidFormat(format!("XML parsing error: {}", e))),
                _ => {}
            }
            buf.clear();
        }

        Err(Error::InvalidFormat("No root element found".to_string()))
    }

    /// Serialize element to XML string
    pub fn to_xml_string(&self) -> String {
        let mut xml = String::new();
        self.write_xml(&mut xml, 0);
        xml
    }

    fn write_xml(&self, output: &mut String, indent: usize) {
        let indent_str = "  ".repeat(indent);

        // Opening tag
        output.push_str(&indent_str);
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
                output.push('\n');
                child.write_xml(output, indent + 1);
            }

            // Closing tag
            if !self.children.is_empty() {
                output.push('\n');
                output.push_str(&indent_str);
            }
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

    fn children(&self) -> &[Box<dyn ElementBase>] {
        // This is a workaround since we can't return &[Box<dyn ElementBase>]
        // from a concrete type. In practice, you'd implement this differently
        // with proper trait objects.
        unsafe { std::mem::transmute(&self.children[..]) }
    }

    fn children_mut(&mut self) -> &mut Vec<Box<dyn ElementBase>> {
        unsafe { std::mem::transmute(&mut self.children) }
    }
}

/// Helper for creating elements with specific tag names
pub struct ElementFactory;

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
