//! Text-related ODF elements.
//!
//! This module provides classes for text elements like paragraphs, spans,
//! headings, and other text content elements.

use super::element::{Element, ElementBase};
use crate::common::{Error, Result};

/// A text paragraph element
#[derive(Debug, Clone)]
pub struct Paragraph {
    element: Element,
}

impl Default for Paragraph {
    fn default() -> Self {
        Self::new()
    }
}

impl Paragraph {
    /// Create a new paragraph
    pub fn new() -> Self {
        Self {
            element: Element::new("text:p"),
        }
    }

    /// Create paragraph from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:p" {
            return Err(Error::InvalidFormat("Element is not a paragraph".to_string()));
        }
        Ok(Self { element })
    }

    /// Get the text content of the paragraph
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the paragraph
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get all text spans within this paragraph
    pub fn spans(&self) -> Result<Vec<Span>> {
        let mut spans = Vec::new();
        for child in self.element.children() {
            if child.tag_name() == "text:span" {
                // This is a simplified conversion - in practice you'd need proper downcasting
                if let Ok(span) = Span::from_element(unsafe { &*(child as *const _ as *const Element) }.clone()) {
                    spans.push(span);
                }
            }
        }
        Ok(spans)
    }

    /// Add a text span to this paragraph
    pub fn add_span(&mut self, span: Span) {
        self.element.add_child(Box::new(span.element));
    }

    /// Check if this paragraph is a heading
    pub fn is_heading(&self) -> bool {
        false // Paragraphs are not headings
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("text:style-name", name);
    }
}

impl From<Paragraph> for Element {
    fn from(para: Paragraph) -> Element {
        para.element
    }
}

/// A text span element (formatted text within a paragraph)
#[derive(Debug, Clone)]
pub struct Span {
    element: Element,
}

impl Default for Span {
    fn default() -> Self {
        Self::new()
    }
}

impl Span {
    /// Create a new span
    pub fn new() -> Self {
        Self {
            element: Element::new("text:span"),
        }
    }

    /// Create span from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:span" {
            return Err(Error::InvalidFormat("Element is not a span".to_string()));
        }
        Ok(Self { element })
    }

    /// Get the text content of the span
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the span
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("text:style-name", name);
    }
}

impl From<Span> for Element {
    fn from(span: Span) -> Element {
        span.element
    }
}

/// A heading element
#[derive(Debug, Clone)]
pub struct Heading {
    element: Element,
}

impl Heading {
    /// Create a new heading
    pub fn new(level: u8) -> Self {
        let mut element = Element::new("text:h");
        element.set_attribute("text:outline-level", &level.to_string());
        Self { element }
    }

    /// Create heading from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:h" {
            return Err(Error::InvalidFormat("Element is not a heading".to_string()));
        }
        Ok(Self { element })
    }

    /// Get the text content of the heading
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the heading
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get the outline level
    pub fn level(&self) -> Option<u8> {
        self.element.get_int_attribute("text:outline-level").map(|n| n as u8)
    }

    /// Set the outline level
    pub fn set_level(&mut self, level: u8) {
        self.element.set_attribute("text:outline-level", &level.to_string());
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("text:style-name", name);
    }

    /// Check if this is a heading
    pub fn is_heading(&self) -> bool {
        true
    }
}

impl From<Heading> for Element {
    fn from(heading: Heading) -> Element {
        heading.element
    }
}

/// A text list element
#[derive(Debug, Clone)]
pub struct List {
    element: Element,
}

impl Default for List {
    fn default() -> Self {
        Self::new()
    }
}

impl List {
    /// Create a new list
    pub fn new() -> Self {
        Self {
            element: Element::new("text:list"),
        }
    }

    /// Create list from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:list" {
            return Err(Error::InvalidFormat("Element is not a list".to_string()));
        }
        Ok(Self { element })
    }

    /// Get list items
    pub fn items(&self) -> Result<Vec<ListItem>> {
        let mut items = Vec::new();
        for child in self.element.children() {
            if child.tag_name() == "text:list-item"
                && let Ok(item) = ListItem::from_element(unsafe { &*(child as *const _ as *const Element) }.clone()) {
                    items.push(item);
                }
        }
        Ok(items)
    }

    /// Add a list item
    pub fn add_item(&mut self, item: ListItem) {
        self.element.add_child(Box::new(item.element));
    }

    /// Get the style name
    pub fn style_name(&self) -> Option<&str> {
        self.element.get_attribute("text:style-name")
    }

    /// Set the style name
    pub fn set_style_name(&mut self, name: &str) {
        self.element.set_attribute("text:style-name", name);
    }
}

impl From<List> for Element {
    fn from(list: List) -> Element {
        list.element
    }
}

/// A list item element
#[derive(Debug, Clone)]
pub struct ListItem {
    element: Element,
}

impl Default for ListItem {
    fn default() -> Self {
        Self::new()
    }
}

impl ListItem {
    /// Create a new list item
    pub fn new() -> Self {
        Self {
            element: Element::new("text:list-item"),
        }
    }

    /// Create list item from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:list-item" {
            return Err(Error::InvalidFormat("Element is not a list item".to_string()));
        }
        Ok(Self { element })
    }

    /// Get the text content of the list item
    pub fn text(&self) -> Result<String> {
        Ok(self.element.get_text_recursive())
    }

    /// Set the text content of the list item
    pub fn set_text(&mut self, text: &str) {
        self.element.set_text(text);
    }

    /// Get nested paragraphs
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut paragraphs = Vec::new();
        for child in self.element.children() {
            if child.tag_name() == "text:p"
                && let Ok(para) = Paragraph::from_element(unsafe { &*(child as *const _ as *const Element) }.clone()) {
                    paragraphs.push(para);
                }
        }
        Ok(paragraphs)
    }

    /// Add a paragraph to this list item
    pub fn add_paragraph(&mut self, paragraph: Paragraph) {
        self.element.add_child(Box::new(paragraph.element));
    }
}

impl From<ListItem> for Element {
    fn from(item: ListItem) -> Element {
        item.element
    }
}

/// A page break element
#[derive(Debug, Clone)]
pub struct PageBreak {
    element: Element,
}

impl Default for PageBreak {
    fn default() -> Self {
        Self::new()
    }
}

impl PageBreak {
    /// Create a new page break
    pub fn new() -> Self {
        let mut element = Element::new("text:p");
        element.set_attribute("text:style-name", "PageBreak");
        Self { element }
    }

    /// Create page break from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:p" {
            return Err(Error::InvalidFormat("Element is not a page break".to_string()));
        }
        if element.get_attribute("text:style-name") != Some("PageBreak") {
            return Err(Error::InvalidFormat("Element is not a page break".to_string()));
        }
        Ok(Self { element })
    }
}

impl From<PageBreak> for Element {
    fn from(pb: PageBreak) -> Element {
        pb.element
    }
}

/// Collection of text elements for easy parsing
pub struct TextElements;

impl TextElements {
    /// Parse all paragraphs from an XML reader
    pub fn parse_paragraphs(xml_content: &str) -> Result<Vec<Paragraph>> {
        let mut reader = quick_xml::Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut paragraphs = Vec::new();
        let mut current_para: Option<Element> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let tag_name = String::from_utf8(e.name().as_ref().to_vec())
                        .unwrap_or_default();

                    if tag_name == "text:p" || tag_name == "text:h" {
                        if let Some(para) = current_para.take()
                            && let Ok(p) = Paragraph::from_element(para) {
                                paragraphs.push(p);
                            }

                        let mut element = Element::new(&tag_name);

                        // Parse attributes
                        for attr_result in e.attributes() {
                            if let Ok(attr) = attr_result
                                && let (Ok(key), Ok(value)) = (
                                    String::from_utf8(attr.key.as_ref().to_vec()),
                                    String::from_utf8(attr.value.to_vec())
                                ) {
                                    element.set_attribute(&key, &value);
                                }
                        }

                        current_para = Some(element);
                    }
                }
                Ok(quick_xml::events::Event::Text(ref t)) => {
                    if let Some(ref mut para) = current_para
                        && let Ok(text) = String::from_utf8(t.to_vec()) {
                            let current_text = para.text().to_string();
                            para.set_text(&format!("{}{}", current_text, text));
                        }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let tag_name = String::from_utf8(e.name().as_ref().to_vec())
                        .unwrap_or_default();

                    if (tag_name == "text:p" || tag_name == "text:h") && current_para.is_some()
                        && let Some(para) = current_para.take()
                            && let Ok(p) = Paragraph::from_element(para) {
                                paragraphs.push(p);
                            }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        // Handle any remaining paragraph
        if let Some(para) = current_para
            && let Ok(p) = Paragraph::from_element(para) {
                paragraphs.push(p);
            }

        Ok(paragraphs)
    }

    /// Parse all headings from XML content
    #[allow(dead_code)]
    pub fn parse_headings(xml_content: &str) -> Result<Vec<Heading>> {
        let paragraphs = Self::parse_paragraphs(xml_content)?;
        let mut headings = Vec::new();

        for para in paragraphs {
            let element = para.element;
            if element.tag_name() == "text:h"
                && let Ok(heading) = Heading::from_element(element) {
                    headings.push(heading);
                }
        }

        Ok(headings)
    }

    /// Extract all text content from XML with improved handling of nested elements
    pub fn extract_text(xml_content: &str) -> Result<String> {
        let mut reader = quick_xml::Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut text = String::new();
        let mut in_paragraph = false;
        let mut paragraph_text = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e)) => {
                    let tag_name = String::from_utf8(e.name().as_ref().to_vec())
                        .unwrap_or_default();

                    match tag_name.as_str() {
                        "text:p" | "text:h" => {
                            if in_paragraph && !paragraph_text.is_empty() {
                                if !text.is_empty() {
                                    text.push('\n');
                                }
                                text.push_str(&paragraph_text);
                                paragraph_text.clear();
                            }
                            in_paragraph = true;
                        }
                        "text:list" => {
                            // Handle lists - add a newline before list if needed
                            if in_paragraph && !paragraph_text.is_empty() {
                                if !text.is_empty() {
                                    text.push('\n');
                                }
                                text.push_str(&paragraph_text);
                                paragraph_text.clear();
                            }
                            in_paragraph = false;
                        }
                        "text:list-item" => {
                            // Add bullet point
                            if !paragraph_text.is_empty() {
                                paragraph_text.push('\n');
                            }
                            paragraph_text.push_str("• ");
                        }
                        "text:line-break" | "text:tab" => {
                            // Handle line breaks and tabs
                            if in_paragraph {
                                if tag_name == "text:line-break" {
                                    paragraph_text.push('\n');
                                } else {
                                    paragraph_text.push('\t');
                                }
                            }
                        }
                        _ => {} // Ignore other elements
                    }
                }
                Ok(quick_xml::events::Event::Text(ref t)) => {
                    if in_paragraph
                        && let Ok(text_content) = String::from_utf8(t.to_vec()) {
                            paragraph_text.push_str(&text_content);
                        }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let tag_name = String::from_utf8(e.name().as_ref().to_vec())
                        .unwrap_or_default();

                    match tag_name.as_str() {
                        "text:p" | "text:h" => {
                            if in_paragraph && !paragraph_text.is_empty() {
                                if !text.is_empty() {
                                    text.push('\n');
                                }
                                text.push_str(&paragraph_text);
                                paragraph_text.clear();
                            }
                            in_paragraph = false;
                        }
                        "text:list" => {
                            in_paragraph = false;
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => {
                    // Handle any remaining paragraph text
                    if in_paragraph && !paragraph_text.is_empty() {
                        if !text.is_empty() {
                            text.push('\n');
                        }
                        text.push_str(&paragraph_text);
                    }
                    break;
                }
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(text)
    }
}
