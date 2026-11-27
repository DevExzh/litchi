//! Generic ODF document parser.
//!
//! This module provides a generic parser for ODF document elements that works across
//! all ODF formats (ODT, ODS, ODP). It parses elements (paragraphs, tables, lists, etc.)
//! in the order they appear in the document, preserving the document structure.
//!
//! For format-specific parsing (e.g., ODT track changes, ODP animations), see the
//! format-specific parsers in `odt/parser.rs`, `ods/parser.rs`, etc.

use crate::elements::element::{ElementBase, try_owned_string, try_prefixed_name};
use crate::elements::table::Table;
use crate::elements::text::{Heading, List, Paragraph};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_DOCUMENT_DEPTH: usize = 4_096;

fn try_push<T>(items: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    items
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })?;
    items.push(value);
    Ok(())
}

/// Represents a document element in its original position
#[derive(Debug, Clone)]
pub enum OrderElement {
    /// A paragraph or heading element
    Paragraph(Paragraph),
    /// A `text:numbered-paragraph` element with explicit list numbering
    NumberedParagraph(super::text::NumberedParagraph),
    /// A heading element (for separate access)
    Heading(Heading),
    /// A table element
    Table(Table),
    /// A list element (currently parsed but not exposed in unified API)
    #[allow(dead_code)] // Parsed but not yet exposed in all APIs
    List(List),
}

/// Generic ODF document parser for parsing elements across all ODF formats.
///
/// This parser provides functionality that is common to all ODF document types
/// (text documents, spreadsheets, presentations). It handles the core document
/// structure elements like paragraphs, tables, headings, and lists.
///
/// For format-specific features, use the specialized parsers:
/// - `crate::parser::Parser` for ODT-specific features (track changes, comments, sections)
/// - `OdsParser` for ODS-specific features (cell formulas, named ranges)
/// - `OdpParser` for ODP-specific features (slide transitions, animations)
pub struct Parser;

impl Parser {
    /// Parse all document elements from XML content in document order.
    ///
    /// This function reads through the XML content once and extracts all major
    /// document elements (paragraphs, headings, tables, lists) in the order they appear.
    ///
    /// # Arguments
    ///
    /// * `xml_content` - The XML content to parse (typically from content.xml)
    ///
    /// # Returns
    ///
    /// A vector of `OrderElement` in the order they appear in the document.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use litchi_odt::elements::parser::Parser;
    ///
    /// let xml = r#"<office:text>
    ///     <text:p>First paragraph</text:p>
    ///     <table:table><table:table-row><table:table-cell><text:p>Cell</text:p></table:table-cell></table:table-row></table:table>
    ///     <text:p>Second paragraph</text:p>
    /// </office:text>"#;
    ///
    /// let elements = Parser::parse_elements_in_order(xml).unwrap();
    /// assert_eq!(elements.len(), 3);
    /// ```
    pub fn parse_elements_in_order(xml_content: &str) -> Result<Vec<OrderElement>> {
        let mut reader = NsReader::from_str(xml_content);
        reader.config_mut().expand_empty_elements = true;
        let mut buf = Vec::new();
        let mut elements = Vec::new();

        // Stack to track nested elements
        let mut element_stack: Vec<(String, super::element::Element)> = Vec::new();
        // Depth tracking to avoid parsing nested elements when inside a parent element
        let mut table_depth = 0;
        let mut list_depth = 0;
        let mut document_depth = 0usize;

        loop {
            let (namespace, event) =
                reader.read_resolved_event_into(&mut buf).map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODF document XML: {error}"))
                })?;
            match event {
                Event::Start(ref e) => {
                    document_depth = document_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("ODF document nesting depth overflow".to_string())
                    })?;
                    if document_depth > MAX_DOCUMENT_DEPTH {
                        return Err(Error::InvalidFormat(format!(
                            "ODF document nesting exceeds {MAX_DOCUMENT_DEPTH} levels"
                        )));
                    }
                    let tag_name = canonical_element_name(
                        &namespace,
                        e.local_name().as_ref(),
                        e.name().as_ref(),
                    )?;

                    match tag_name.as_str() {
                        "text:p" if table_depth == 0 && list_depth == 0 => {
                            // Start a paragraph outside of tables and lists
                            let mut element = super::element::Element::try_new(&tag_name)?;

                            copy_attributes(&reader, e, &mut element)?;

                            try_push(
                                &mut element_stack,
                                (tag_name, element),
                                "ODT ordered-element parser stack",
                            )?;
                        },
                        "text:numbered-paragraph" if table_depth == 0 && list_depth == 0 => {
                            // Start a numbered paragraph outside of tables and lists
                            let mut element = super::element::Element::try_new(&tag_name)?;

                            copy_attributes(&reader, e, &mut element)?;

                            try_push(
                                &mut element_stack,
                                (tag_name, element),
                                "ODT ordered-element parser stack",
                            )?;
                        },
                        "text:h" if table_depth == 0 && list_depth == 0 => {
                            // Start a heading outside of tables and lists
                            let mut element = super::element::Element::try_new(&tag_name)?;

                            copy_attributes(&reader, e, &mut element)?;

                            try_push(
                                &mut element_stack,
                                (tag_name, element),
                                "ODT ordered-element parser stack",
                            )?;
                        },
                        "table:table" if table_depth == 0 => {
                            // Start a table
                            table_depth += 1;
                            let mut element = super::element::Element::try_new(&tag_name)?;

                            copy_attributes(&reader, e, &mut element)?;

                            try_push(
                                &mut element_stack,
                                (tag_name, element),
                                "ODT ordered-element parser stack",
                            )?;
                        },
                        "table:table" => {
                            // Nested table
                            table_depth += 1;
                        },
                        "text:list" if list_depth == 0 && table_depth == 0 => {
                            // Start a list outside of tables
                            list_depth += 1;
                            let mut element = super::element::Element::try_new(&tag_name)?;

                            copy_attributes(&reader, e, &mut element)?;

                            try_push(
                                &mut element_stack,
                                (tag_name, element),
                                "ODT ordered-element parser stack",
                            )?;
                        },
                        "text:list" => {
                            // Nested list
                            list_depth += 1;
                        },
                        _ if matches!(
                            element_stack.first().map(|(tag, _)| tag.as_str()),
                            Some("text:p" | "text:h" | "text:numbered-paragraph")
                        ) =>
                        {
                            if let Some((_, element)) = element_stack.last_mut() {
                                append_text_control(&reader, &tag_name, e, element)?;
                            }
                        },
                        // Handle nested elements within tracked elements
                        _ if !element_stack.is_empty() && table_depth <= 1 && list_depth <= 1 => {
                            let mut element = super::element::Element::try_new(&tag_name)?;

                            copy_attributes(&reader, e, &mut element)?;

                            try_push(
                                &mut element_stack,
                                (tag_name, element),
                                "ODT ordered-element parser stack",
                            )?;
                        },
                        _ => {},
                    }
                },
                Event::Text(ref t) => {
                    // Add text content to the current element
                    if let Some((_, element)) = element_stack.last_mut() {
                        let text = t.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                            Error::InvalidFormat(format!("invalid ODF element text: {error}"))
                        })?;
                        element.try_append_text(&text, "ODT ordered-element text")?;
                    }
                },
                Event::CData(ref value) => {
                    if let Some((_, element)) = element_stack.last_mut() {
                        let text = value
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| {
                                Error::InvalidFormat(format!("invalid ODF element CDATA: {error}"))
                            })?;
                        element.try_append_text(&text, "ODT ordered-element CDATA")?;
                    }
                },
                Event::GeneralRef(ref reference) => {
                    if let Some((_, element)) = element_stack.last_mut() {
                        let text = decode_reference(reference)?;
                        element.try_append_text(&text, "ODT ordered-element reference")?;
                    }
                },
                Event::End(ref e) => {
                    document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("ODF document element stack underflow".to_string())
                    })?;
                    let tag_name = canonical_element_name(
                        &namespace,
                        e.local_name().as_ref(),
                        e.name().as_ref(),
                    )?;

                    match tag_name.as_str() {
                        "text:p" if table_depth == 0 && list_depth == 0 => {
                            // Complete a top-level paragraph
                            let (tag, element) = element_stack.pop().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODF paragraph end has no projected start".to_string(),
                                )
                            })?;
                            if tag != "text:p" {
                                return Err(Error::InvalidFormat(
                                    "ODF paragraph projection stack mismatch".to_string(),
                                ));
                            }
                            let para = Paragraph::from_element(element)?;
                            try_push(
                                &mut elements,
                                OrderElement::Paragraph(para),
                                "ODT ordered-element projection",
                            )?;
                        },
                        "text:numbered-paragraph" if table_depth == 0 && list_depth == 0 => {
                            // Complete a top-level numbered paragraph
                            let (tag, element) = element_stack.pop().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODF numbered paragraph end has no projected start".to_string(),
                                )
                            })?;
                            if tag != "text:numbered-paragraph" {
                                return Err(Error::InvalidFormat(
                                    "ODF numbered paragraph projection stack mismatch".to_string(),
                                ));
                            }
                            let para = super::text::NumberedParagraph::from_element(element)?;
                            try_push(
                                &mut elements,
                                OrderElement::NumberedParagraph(para),
                                "ODT ordered-element projection",
                            )?;
                        },
                        "text:h" if table_depth == 0 && list_depth == 0 => {
                            // Complete a top-level heading
                            let (tag, element) = element_stack.pop().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODF heading end has no projected start".to_string(),
                                )
                            })?;
                            if tag != "text:h" {
                                return Err(Error::InvalidFormat(
                                    "ODF heading projection stack mismatch".to_string(),
                                ));
                            }
                            let heading = Heading::from_element(element)?;
                            try_push(
                                &mut elements,
                                OrderElement::Heading(heading),
                                "ODT ordered-element projection",
                            )?;
                        },
                        "table:table" if table_depth == 1 => {
                            // Complete a top-level table
                            table_depth -= 1;
                            let (tag, element) = element_stack.pop().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODF table end has no projected start".to_string(),
                                )
                            })?;
                            if tag != "table:table" {
                                return Err(Error::InvalidFormat(
                                    "ODF table projection stack mismatch".to_string(),
                                ));
                            }
                            let table = Table::from_element(element)?;
                            try_push(
                                &mut elements,
                                OrderElement::Table(table),
                                "ODT ordered-element projection",
                            )?;
                        },
                        "table:table" => {
                            table_depth -= 1;
                        },
                        "text:list" if list_depth == 1 && table_depth == 0 => {
                            // Complete a top-level list
                            list_depth -= 1;
                            let (tag, element) = element_stack.pop().ok_or_else(|| {
                                Error::InvalidFormat(
                                    "ODF list end has no projected start".to_string(),
                                )
                            })?;
                            if tag != "text:list" {
                                return Err(Error::InvalidFormat(
                                    "ODF list projection stack mismatch".to_string(),
                                ));
                            }
                            let list = List::from_element(element)?;
                            try_push(
                                &mut elements,
                                OrderElement::List(list),
                                "ODT ordered-element projection",
                            )?;
                        },
                        "text:list" => {
                            list_depth -= 1;
                        },
                        _ if !element_stack.is_empty() => {
                            // Pop nested element and add to parent
                            if element_stack.len() > 1 {
                                let (_, child_element) = element_stack.pop().ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "ODF element parser lost a nested element".to_string(),
                                    )
                                })?;
                                if let Some((_, parent_element)) = element_stack.last_mut() {
                                    parent_element.try_add_child(
                                        child_element,
                                        "ODT ordered-element child projection",
                                    )?;
                                }
                            } else {
                                // Single element on stack, check if it should be completed
                                if let Some((tag, _)) = element_stack.last()
                                    && tag == &tag_name
                                {
                                    element_stack.pop();
                                }
                            }
                        },
                        _ => {
                            // Ignore end tags when stack is empty or doesn't match
                        },
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            buf.clear();
        }

        if document_depth != 0 || !element_stack.is_empty() || table_depth != 0 || list_depth != 0 {
            return Err(Error::InvalidFormat(
                "incomplete ODF document element structure".to_string(),
            ));
        }
        Ok(elements)
    }

    /// Parse only paragraphs and headings in order.
    ///
    /// This is a convenience method that filters out only text elements.
    #[allow(dead_code)] // Library API for specialized parsing
    pub fn parse_text_elements_in_order(xml_content: &str) -> Result<Vec<Paragraph>> {
        let elements = Self::parse_elements_in_order(xml_content)?;
        let mut paragraphs = Vec::new();

        for element in elements {
            match element {
                OrderElement::Paragraph(para) => {
                    try_push(&mut paragraphs, para, "ODT paragraph projection")?;
                },
                OrderElement::Heading(heading) => {
                    let para = heading.try_into_paragraph()?;
                    try_push(&mut paragraphs, para, "ODT paragraph projection")?;
                },
                _ => {},
            }
        }

        Ok(paragraphs)
    }

    /// Parse only tables in order.
    ///
    /// This is a convenience method that filters out only table elements.
    #[allow(dead_code)] // Library API for specialized parsing
    pub fn parse_tables_in_order(xml_content: &str) -> Result<Vec<Table>> {
        let elements = Self::parse_elements_in_order(xml_content)?;
        let mut tables = Vec::new();

        for element in elements {
            if let OrderElement::Table(table) = element {
                try_push(&mut tables, table, "ODT table projection")?;
            }
        }

        Ok(tables)
    }
}

fn canonical_element_name(
    namespace: &ResolveResult<'_>,
    local_name: &[u8],
    qualified_name: &[u8],
) -> Result<String> {
    let local_name = std::str::from_utf8(local_name)
        .map_err(|_error| Error::InvalidFormat("non-UTF-8 ODF element name".to_string()))?;
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == TEXT_NAMESPACE => {
            try_prefixed_name("text", local_name, "ODT element name")
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE => {
            try_prefixed_name("table", local_name, "ODT element name")
        },
        ResolveResult::Bound(_) | ResolveResult::Unbound => std::str::from_utf8(qualified_name)
            .map_err(|_error| Error::InvalidFormat("non-UTF-8 ODF element name".to_string()))
            .and_then(|name| try_owned_string(name, "ODT element name")),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown ODF element namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn copy_attributes(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
    element: &mut super::element::Element,
) -> Result<()> {
    for attribute in source.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        let local_name = std::str::from_utf8(local_name.as_ref())
            .map_err(|_error| Error::InvalidFormat("non-UTF-8 ODF attribute name".to_string()))?;
        let name = match namespace {
            ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE => {
                try_prefixed_name("text", local_name, "ODT attribute name")?
            },
            ResolveResult::Bound(Namespace(uri)) if uri == TABLE_NAMESPACE => {
                try_prefixed_name("table", local_name, "ODT attribute name")?
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE => {
                try_prefixed_name("xlink", local_name, "ODT attribute name")?
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XML_NAMESPACE => {
                try_prefixed_name("xml", local_name, "ODT attribute name")?
            },
            ResolveResult::Bound(_) | ResolveResult::Unbound => {
                std::str::from_utf8(attribute.key.as_ref())
                    .map_err(|_error| {
                        Error::InvalidFormat("non-UTF-8 ODF attribute name".to_string())
                    })
                    .and_then(|name| try_owned_string(name, "ODT attribute name"))?
            },
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unknown ODF attribute namespace prefix '{}'",
                    String::from_utf8_lossy(&prefix)
                )));
            },
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODF attribute value: {error}"))
            })?;
        if element.has_attribute(&name) {
            return Err(Error::InvalidFormat(format!(
                "duplicate ODF attribute '{name}'"
            )));
        }
        element.try_set_attribute(&name, &value, "ODT ordered-element attribute")?;
    }
    Ok(())
}

fn append_text_control(
    reader: &NsReader<&[u8]>,
    tag_name: &str,
    source: &BytesStart<'_>,
    element: &mut super::element::Element,
) -> Result<()> {
    let value = match tag_name {
        "text:s" => {
            let mut count = 1usize;
            let mut count_seen = false;
            for attribute in source.attributes() {
                let attribute = attribute.map_err(|error| {
                    Error::InvalidFormat(format!("invalid text:s attribute: {error}"))
                })?;
                let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
                if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == TEXT_NAMESPACE)
                    && local_name.as_ref() == b"c"
                {
                    if count_seen {
                        return Err(Error::InvalidFormat(
                            "duplicate expanded text:c attribute".to_string(),
                        ));
                    }
                    let decoded = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid text:c value: {error}"))
                        })?;
                    count = decoded.parse().map_err(|_error| {
                        Error::InvalidFormat("text:c must be a non-negative integer".to_string())
                    })?;
                    count_seen = true;
                }
            }
            if count > 1_000_000 {
                return Err(Error::InvalidFormat(
                    "text:s count exceeds 1000000".to_string(),
                ));
            }
            element.try_append_spaces(count, "ODT ordered-element text")?;
            return Ok(());
        },
        "text:tab" => "\t",
        "text:line-break" => "\n",
        _ => return Ok(()),
    };
    element.try_append_text(value, "ODT ordered-element text")
}

pub(super) fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid ODF character reference: {error}"))
    })? {
        let mut encoded = [0_u8; 4];
        return try_owned_string(
            character.encode_utf8(&mut encoded),
            "ODT character reference",
        );
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODF entity: {error}")))?;
    match name.as_ref() {
        "amp" => try_owned_string("&", "ODT entity reference"),
        "lt" => try_owned_string("<", "ODT entity reference"),
        "gt" => try_owned_string(">", "ODT entity reference"),
        "quot" => try_owned_string("\"", "ODT entity reference"),
        "apos" => try_owned_string("'", "ODT entity reference"),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported ODF entity '&{name};'"
        ))),
    }
}
