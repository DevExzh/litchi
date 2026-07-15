//! Field elements for ODF documents.
//!
//! Fields are dynamic content in ODF documents that can be updated automatically,
//! such as page numbers, dates, cross-references, etc.

use super::element::{Element, ElementBase};
use super::xml::{
    TEXT_NAMESPACE, append_checked, append_text_control, copy_canonical_attributes,
    decode_reference, is_bound,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

const MAX_FIELD_DEPTH: usize = 4_096;
const MAX_FIELDS: usize = 1_000_000;

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
                | "text:page-continuation"
                | "text:page-variable-set"
                | "text:page-variable-get"
                | "text:date"
                | "text:time"
                | "text:file-name"
                | "text:template-name"
                | "text:sheet-name"
                | "text:author-name"
                | "text:author-initials"
                | "text:sender-firstname"
                | "text:sender-lastname"
                | "text:sender-initials"
                | "text:sender-title"
                | "text:sender-position"
                | "text:sender-email"
                | "text:sender-phone-private"
                | "text:sender-fax"
                | "text:sender-company"
                | "text:sender-phone-work"
                | "text:sender-street"
                | "text:sender-city"
                | "text:sender-postal-code"
                | "text:sender-country"
                | "text:sender-state-or-province"
                | "text:chapter"
                | "text:title"
                | "text:subject"
                | "text:keywords"
                | "text:description"
                | "text:user-defined"
                | "text:creator"
                | "text:initial-creator"
                | "text:printed-by"
                | "text:creation-date"
                | "text:creation-time"
                | "text:modification-date"
                | "text:modification-time"
                | "text:print-date"
                | "text:print-time"
                | "text:editing-duration"
                | "text:editing-cycles"
                | "text:reference-ref"
                | "text:sequence-ref"
                | "text:bookmark-ref"
                | "text:note-ref"
                | "text:variable-set"
                | "text:variable-get"
                | "text:variable-input"
                | "text:user-field-get"
                | "text:user-field-input"
                | "text:sequence"
                | "text:expression"
                | "text:text-input"
                | "text:placeholder"
                | "text:conditional-text"
                | "text:hidden-text"
                | "text:hidden-paragraph"
                | "text:measure"
                | "text:table-formula"
                | "text:database-display"
                | "text:database-next"
                | "text:database-row-select"
                | "text:database-row-number"
                | "text:database-name"
                | "text:word-count"
                | "text:paragraph-count"
                | "text:character-count"
                | "text:table-count"
                | "text:image-count"
                | "text:object-count"
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
        let mut reader = NsReader::from_str(xml_content);
        let mut buffer = Vec::new();
        let mut document_depth = 0usize;
        let mut active: Vec<ActiveField> = Vec::new();
        let mut fields = Vec::new();
        let mut next_order = 0usize;

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("invalid field XML: {error}")))?;
            let text_element = is_bound(&namespace, TEXT_NAMESPACE);
            match event {
                Event::Start(ref source) => {
                    document_depth = checked_field_depth(document_depth)?;
                    for field in &mut active {
                        field.depth += 1;
                    }
                    if text_element {
                        for field in &mut active {
                            append_text_control(&reader, source, &mut field.text)?;
                        }
                        let tag_name = format!(
                            "text:{}",
                            std::str::from_utf8(source.local_name().as_ref()).map_err(|_| {
                                Error::InvalidFormat("non-UTF-8 field element name".to_string())
                            })?
                        );
                        if Field::is_field_tag(&tag_name) {
                            if next_order >= MAX_FIELDS {
                                return Err(Error::InvalidFormat(format!(
                                    "document exceeds {MAX_FIELDS} fields"
                                )));
                            }
                            let mut element = Element::new(&tag_name);
                            copy_canonical_attributes(&reader, source, &mut element, "field")?;
                            active.push(ActiveField {
                                element,
                                text: String::new(),
                                depth: 1,
                                order: next_order,
                            });
                            next_order += 1;
                        }
                    }
                },
                Event::Empty(ref source) if text_element => {
                    for field in &mut active {
                        append_text_control(&reader, source, &mut field.text)?;
                    }
                    let tag_name = format!(
                        "text:{}",
                        std::str::from_utf8(source.local_name().as_ref()).map_err(|_| {
                            Error::InvalidFormat("non-UTF-8 field element name".to_string())
                        })?
                    );
                    if Field::is_field_tag(&tag_name) {
                        if next_order >= MAX_FIELDS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_FIELDS} fields"
                            )));
                        }
                        let mut element = Element::new(&tag_name);
                        copy_canonical_attributes(&reader, source, &mut element, "field")?;
                        fields.push((next_order, Field::from_element(element)?));
                        next_order += 1;
                    }
                },
                Event::Text(ref value) if !active.is_empty() => {
                    let value = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid field text: {error}"))
                        })?;
                    for field in &mut active {
                        append_checked(&mut field.text, &value)?;
                    }
                },
                Event::CData(ref value) if !active.is_empty() => {
                    let value = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid field CDATA: {error}"))
                        })?;
                    for field in &mut active {
                        append_checked(&mut field.text, &value)?;
                    }
                },
                Event::GeneralRef(ref reference) if !active.is_empty() => {
                    let value = decode_reference(reference, "field")?;
                    for field in &mut active {
                        append_checked(&mut field.text, &value)?;
                    }
                },
                Event::End(_) => {
                    document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("field XML stack underflow".to_string())
                    })?;
                    for field in &mut active {
                        field.depth = field.depth.checked_sub(1).ok_or_else(|| {
                            Error::InvalidFormat("field element stack underflow".to_string())
                        })?;
                    }
                    if active.last().is_some_and(|field| field.depth == 0) {
                        let mut field = active.pop().expect("checked active field");
                        field.element.set_text(&field.text);
                        fields.push((field.order, Field::from_element(field.element)?));
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
        if document_depth != 0 || !active.is_empty() {
            return Err(Error::InvalidFormat(
                "incomplete field XML structure".to_string(),
            ));
        }
        fields.sort_by_key(|(order, _)| *order);
        Ok(fields.into_iter().map(|(_, field)| field).collect())
    }
}

struct ActiveField {
    element: Element,
    text: String,
    depth: usize,
    order: usize,
}

fn checked_field_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("field nesting depth overflow".to_string()))?;
    if depth > MAX_FIELD_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "field nesting exceeds {MAX_FIELD_DEPTH} levels"
        )));
    }
    Ok(depth)
}
