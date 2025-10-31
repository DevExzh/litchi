//! Generic ODF document parser.
//!
//! This module provides a generic parser for ODF document elements that works across
//! all ODF formats (ODT, ODS, ODP). It parses elements (paragraphs, tables, lists, etc.)
//! in the order they appear in the document, preserving the document structure.
//!
//! For format-specific parsing (e.g., ODT track changes, ODP animations), see the
//! format-specific parsers in `odt/parser.rs`, `ods/parser.rs`, etc.

use crate::common::Result;
use crate::odf::elements::element::ElementBase;
use crate::odf::elements::table::Table;
use crate::odf::elements::text::{Heading, List, Paragraph};
use quick_xml::Reader;
use quick_xml::events::Event;

/// Represents a document element in its original position
#[derive(Debug, Clone)]
pub enum DocumentOrderElement {
    /// A paragraph or heading element
    Paragraph(Paragraph),
    /// A heading element (for separate access)
    Heading(Heading),
    /// A table element
    Table(Table),
    /// A list element
    List(List),
}

/// Generic ODF document parser for parsing elements across all ODF formats.
///
/// This parser provides functionality that is common to all ODF document types
/// (text documents, spreadsheets, presentations). It handles the core document
/// structure elements like paragraphs, tables, headings, and lists.
///
/// For format-specific features, use the specialized parsers:
/// - `OdtParser` for ODT-specific features (track changes, comments, sections)
/// - `OdsParser` for ODS-specific features (cell formulas, named ranges)
/// - `OdpParser` for ODP-specific features (slide transitions, animations)
pub struct DocumentParser;

impl DocumentParser {
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
    /// A vector of `DocumentOrderElement` in the order they appear in the document.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use litchi::odf::elements::parser::DocumentParser;
    ///
    /// let xml = r#"<office:text>
    ///     <text:p>First paragraph</text:p>
    ///     <table:table><table:table-row><table:table-cell><text:p>Cell</text:p></table:table-cell></table:table-row></table:table>
    ///     <text:p>Second paragraph</text:p>
    /// </office:text>"#;
    ///
    /// let elements = DocumentParser::parse_elements_in_order(xml).unwrap();
    /// assert_eq!(elements.len(), 3);
    /// ```
    pub fn parse_elements_in_order(xml_content: &str) -> Result<Vec<DocumentOrderElement>> {
        let mut reader = Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut elements = Vec::new();

        // Stack to track nested elements
        let mut element_stack: Vec<(String, super::element::Element)> = Vec::new();
        // Depth tracking to avoid parsing nested elements when inside a parent element
        let mut table_depth = 0;
        let mut list_depth = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    match tag_name.as_str() {
                        "text:p" if table_depth == 0 && list_depth == 0 => {
                            // Start a paragraph outside of tables and lists
                            let mut element = super::element::Element::new(&tag_name);

                            // Parse attributes
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                let value = String::from_utf8_lossy(&attr.value);
                                element.set_attribute(&key, &value);
                            }

                            element_stack.push((tag_name, element));
                        },
                        "text:h" if table_depth == 0 && list_depth == 0 => {
                            // Start a heading outside of tables and lists
                            let mut element = super::element::Element::new(&tag_name);

                            // Parse attributes
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                let value = String::from_utf8_lossy(&attr.value);
                                element.set_attribute(&key, &value);
                            }

                            element_stack.push((tag_name, element));
                        },
                        "table:table" if table_depth == 0 => {
                            // Start a table
                            table_depth += 1;
                            let mut element = super::element::Element::new(&tag_name);

                            // Parse attributes
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                let value = String::from_utf8_lossy(&attr.value);
                                element.set_attribute(&key, &value);
                            }

                            element_stack.push((tag_name, element));
                        },
                        "table:table" => {
                            // Nested table
                            table_depth += 1;
                        },
                        "text:list" if list_depth == 0 && table_depth == 0 => {
                            // Start a list outside of tables
                            list_depth += 1;
                            let mut element = super::element::Element::new(&tag_name);

                            // Parse attributes
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                let value = String::from_utf8_lossy(&attr.value);
                                element.set_attribute(&key, &value);
                            }

                            element_stack.push((tag_name, element));
                        },
                        "text:list" => {
                            // Nested list
                            list_depth += 1;
                        },
                        // Handle nested elements within tracked elements
                        _ if !element_stack.is_empty() && table_depth <= 1 && list_depth <= 1 => {
                            let mut element = super::element::Element::new(&tag_name);

                            // Parse attributes
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                let value = String::from_utf8_lossy(&attr.value);
                                element.set_attribute(&key, &value);
                            }

                            element_stack.push((tag_name, element));
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(ref t)) => {
                    // Add text content to the current element
                    if let Some((_, element)) = element_stack.last_mut() {
                        let text = String::from_utf8_lossy(t).to_string();
                        let current_text = element.text().to_string();
                        element.set_text(&format!("{}{}", current_text, text));
                    }
                },
                Ok(Event::End(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    match tag_name.as_str() {
                        "text:p" if table_depth == 0 && list_depth == 0 => {
                            // Complete a top-level paragraph
                            if let Some((tag, element)) = element_stack.pop()
                                && tag == "text:p"
                                && let Ok(para) = Paragraph::from_element(element)
                            {
                                elements.push(DocumentOrderElement::Paragraph(para));
                            }
                        },
                        "text:h" if table_depth == 0 && list_depth == 0 => {
                            // Complete a top-level heading
                            if let Some((tag, element)) = element_stack.pop()
                                && tag == "text:h"
                                && let Ok(heading) = Heading::from_element(element)
                            {
                                elements.push(DocumentOrderElement::Heading(heading));
                            }
                        },
                        "table:table" if table_depth == 1 => {
                            // Complete a top-level table
                            table_depth -= 1;
                            if let Some((tag, element)) = element_stack.pop()
                                && tag == "table:table"
                                && let Ok(table) = Table::from_element(element)
                            {
                                elements.push(DocumentOrderElement::Table(table));
                            }
                        },
                        "table:table" => {
                            table_depth -= 1;
                        },
                        "text:list" if list_depth == 1 && table_depth == 0 => {
                            // Complete a top-level list
                            list_depth -= 1;
                            if let Some((tag, element)) = element_stack.pop()
                                && tag == "text:list"
                                && let Ok(list) = List::from_element(element)
                            {
                                elements.push(DocumentOrderElement::List(list));
                            }
                        },
                        "text:list" => {
                            list_depth -= 1;
                        },
                        _ if !element_stack.is_empty() => {
                            // Pop nested element and add to parent
                            if element_stack.len() > 1 {
                                let (_, child_element) = element_stack.pop().unwrap();
                                if let Some((_, parent_element)) = element_stack.last_mut() {
                                    parent_element.add_child(Box::new(child_element));
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
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
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
                DocumentOrderElement::Paragraph(para) => paragraphs.push(para),
                DocumentOrderElement::Heading(heading) => {
                    // Convert heading to paragraph for unified handling
                    if let Ok(text) = heading.text() {
                        let mut para = Paragraph::new();
                        para.set_text(&text);
                        if let Some(style) = heading.style_name() {
                            para.set_style_name(style);
                        }
                        paragraphs.push(para);
                    }
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
            if let DocumentOrderElement::Table(table) = element {
                tables.push(table);
            }
        }

        Ok(tables)
    }
}
