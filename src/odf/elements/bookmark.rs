//! Bookmark elements for ODF documents.
//!
//! Bookmarks mark specific locations in a document that can be referenced
//! by cross-references and hyperlinks.

use super::element::{Element, ElementBase};
use crate::common::{Error, Result};

/// Represents a bookmark in the document
#[derive(Debug, Clone)]
pub struct Bookmark {
    element: Element,
}

impl Bookmark {
    /// Create a new bookmark
    pub fn new(name: &str) -> Self {
        let mut element = Element::new("text:bookmark");
        element.set_attribute("text:name", name);
        Self { element }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:bookmark" {
            return Err(Error::InvalidFormat(
                "Element is not a bookmark".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the bookmark name
    pub fn name(&self) -> Option<&str> {
        self.element.get_attribute("text:name")
    }

    /// Set the bookmark name
    pub fn set_name(&mut self, name: &str) {
        self.element.set_attribute("text:name", name);
    }
}

/// Represents a bookmark start marker
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct BookmarkStart {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl BookmarkStart {
    /// Create a new bookmark start marker
    pub fn new(name: &str) -> Self {
        let mut element = Element::new("text:bookmark-start");
        element.set_attribute("text:name", name);
        Self { element }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:bookmark-start" {
            return Err(Error::InvalidFormat(
                "Element is not a bookmark start".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the bookmark name
    pub fn name(&self) -> Option<&str> {
        self.element.get_attribute("text:name")
    }

    /// Set the bookmark name
    pub fn set_name(&mut self, name: &str) {
        self.element.set_attribute("text:name", name);
    }
}

/// Represents a bookmark end marker
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct BookmarkEnd {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl BookmarkEnd {
    /// Create a new bookmark end marker
    pub fn new(name: &str) -> Self {
        let mut element = Element::new("text:bookmark-end");
        element.set_attribute("text:name", name);
        Self { element }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:bookmark-end" {
            return Err(Error::InvalidFormat(
                "Element is not a bookmark end".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the bookmark name
    pub fn name(&self) -> Option<&str> {
        self.element.get_attribute("text:name")
    }

    /// Set the bookmark name
    pub fn set_name(&mut self, name: &str) {
        self.element.set_attribute("text:name", name);
    }
}

/// Represents a bookmark range (start and end)
#[derive(Debug, Clone)]
pub struct BookmarkRange {
    /// Bookmark name
    pub name: String,
    /// Start position (paragraph index, character offset)
    pub start: Option<(usize, usize)>,
    /// End position (paragraph index, character offset)
    pub end: Option<(usize, usize)>,
}

impl BookmarkRange {
    /// Create a new bookmark range
    pub fn new(name: String) -> Self {
        Self {
            name,
            start: None,
            end: None,
        }
    }

    /// Check if the bookmark range is complete (has both start and end)
    pub fn is_complete(&self) -> bool {
        self.start.is_some() && self.end.is_some()
    }
}

/// Utilities for parsing bookmarks from documents
pub struct BookmarkParser;

impl BookmarkParser {
    /// Parse all bookmarks from XML content
    pub fn parse_bookmarks(xml_content: &str) -> Result<Vec<Bookmark>> {
        let mut reader = quick_xml::Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut bookmarks = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    if tag_name == "text:bookmark" {
                        let mut element = Element::new(&tag_name);

                        // Parse attributes
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref());
                            let value = String::from_utf8_lossy(&attr.value);
                            element.set_attribute(&key, &value);
                        }

                        if let Ok(bookmark) = Bookmark::from_element(element) {
                            bookmarks.push(bookmark);
                        }
                    }
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(bookmarks)
    }

    /// Parse bookmark ranges (start/end pairs) from XML content
    pub fn parse_bookmark_ranges(xml_content: &str) -> Result<Vec<BookmarkRange>> {
        let mut reader = quick_xml::Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut ranges: Vec<BookmarkRange> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Empty(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    match tag_name.as_str() {
                        "text:bookmark-start" => {
                            // Extract name
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                if key == "text:name" {
                                    let name = String::from_utf8_lossy(&attr.value).to_string();
                                    ranges.push(BookmarkRange::new(name));
                                }
                            }
                        },
                        "text:bookmark-end" => {
                            // Find matching start
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                if key == "text:name" {
                                    let name = String::from_utf8_lossy(&attr.value);
                                    // Mark the range as complete
                                    if let Some(range) = ranges.iter_mut().find(|r| r.name == name)
                                    {
                                        // In a full implementation, we would track positions
                                        range.end = Some((0, 0));
                                    }
                                }
                            }
                        },
                        _ => {},
                    }
                },
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(ranges)
    }
}
