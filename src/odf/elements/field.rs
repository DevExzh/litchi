//! Field elements for ODF documents.
//!
//! Fields are dynamic content in ODF documents that can be updated automatically,
//! such as page numbers, dates, cross-references, etc.

use super::element::{Element, ElementBase};
use crate::common::{Error, Result};

/// Represents a text field in the document
#[derive(Debug, Clone)]
pub struct Field {
    element: Element,
}

impl Field {
    /// Create a new field from an element
    pub fn from_element(element: Element) -> Result<Self> {
        let tag = element.tag_name();
        if !Self::is_field_tag(tag) {
            return Err(Error::InvalidFormat(format!(
                "Element {} is not a field",
                tag
            )));
        }
        Ok(Self { element })
    }

    /// Check if a tag name represents a field
    pub fn is_field_tag(tag: &str) -> bool {
        matches!(
            tag,
            "text:page-number"
                | "text:page-count"
                | "text:date"
                | "text:time"
                | "text:file-name"
                | "text:author-name"
                | "text:author-initials"
                | "text:title"
                | "text:subject"
                | "text:keywords"
                | "text:description"
                | "text:user-defined"
                | "text:reference-ref"
                | "text:sequence-ref"
                | "text:bookmark-ref"
                | "text:variable-set"
                | "text:variable-get"
                | "text:user-field-get"
                | "text:expression"
        )
    }

    /// Get the field type
    pub fn field_type(&self) -> &str {
        self.element.tag_name()
    }

    /// Get the field value (text content)
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }

    /// Get the field display format
    pub fn format(&self) -> Option<&str> {
        self.element
            .get_attribute("style:data-style-name")
            .or_else(|| self.element.get_attribute("number:style"))
    }

    /// Get the field name (for named fields like variables or user fields)
    pub fn name(&self) -> Option<&str> {
        self.element
            .get_attribute("text:name")
            .or_else(|| self.element.get_attribute("text:variable-name"))
    }

    /// Get reference target (for reference fields)
    pub fn reference_target(&self) -> Option<&str> {
        self.element
            .get_attribute("text:ref-name")
            .or_else(|| self.element.get_attribute("text:reference-name"))
    }
}

/// Represents a page number field
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct PageNumberField {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl PageNumberField {
    /// Create a new page number field
    pub fn new() -> Self {
        Self {
            element: Element::new("text:page-number"),
        }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:page-number" {
            return Err(Error::InvalidFormat(
                "Element is not a page number field".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the current page number value
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }
}

impl Default for PageNumberField {
    fn default() -> Self {
        Self::new()
    }
}

/// Represents a date field
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct DateField {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl DateField {
    /// Create a new date field
    pub fn new() -> Self {
        Self {
            element: Element::new("text:date"),
        }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:date" {
            return Err(Error::InvalidFormat(
                "Element is not a date field".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the date value
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }

    /// Get the fixed date (if any)
    pub fn fixed_date(&self) -> Option<&str> {
        self.element.get_attribute("text:date-value")
    }

    /// Get whether this date is fixed
    pub fn is_fixed(&self) -> bool {
        self.element
            .get_bool_attribute("text:fixed")
            .unwrap_or(false)
    }
}

impl Default for DateField {
    fn default() -> Self {
        Self::new()
    }
}

/// Represents a reference field
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct ReferenceField {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl ReferenceField {
    /// Create a new reference field
    pub fn new(ref_name: &str) -> Self {
        let mut element = Element::new("text:reference-ref");
        element.set_attribute("text:ref-name", ref_name);
        Self { element }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        let tag = element.tag_name();
        if !matches!(
            tag,
            "text:reference-ref" | "text:bookmark-ref" | "text:sequence-ref"
        ) {
            return Err(Error::InvalidFormat(format!(
                "Element {} is not a reference field",
                tag
            )));
        }
        Ok(Self { element })
    }

    /// Get the reference name
    pub fn ref_name(&self) -> Option<&str> {
        self.element.get_attribute("text:ref-name")
    }

    /// Get the reference format
    pub fn ref_format(&self) -> Option<&str> {
        self.element.get_attribute("text:reference-format")
    }

    /// Get the reference value
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }
}

/// Utilities for parsing fields from documents
pub struct FieldParser;

impl FieldParser {
    /// Parse all fields from XML content
    pub fn parse_fields(xml_content: &str) -> Result<Vec<Field>> {
        let mut reader = quick_xml::Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut fields = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e))
                | Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    if Field::is_field_tag(&tag_name) {
                        let mut element = Element::new(&tag_name);

                        // Parse attributes
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref());
                            let value = String::from_utf8_lossy(&attr.value);
                            element.set_attribute(&key, &value);
                        }

                        if let Ok(field) = Field::from_element(element) {
                            fields.push(field);
                        }
                    }
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(fields)
    }
}
