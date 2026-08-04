//! Bookmark elements for ODF documents.
//!
//! Bookmarks mark specific locations in a document that can be referenced
//! by cross-references and hyperlinks.

use super::element::{Element, ElementBase};
use super::xml::{
    TEXT_NAMESPACE, append_text_control, copy_canonical_attributes, decode_reference, is_bound,
    namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

mod writing;
pub use writing::{
    BookmarkFragments, BookmarkTarget, insert_bookmark_xml, parse_bookmark_targets,
    remove_bookmark_xml, replace_bookmark_xml,
};

const MAX_BOOKMARK_DEPTH: usize = 4_096;
const MAX_BOOKMARKS: usize = 1_000_000;

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
        writing::validate_bookmark_xml(xml_content)?;
        let mut reader = NsReader::from_str(xml_content);
        reader.config_mut().expand_empty_elements = true;
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut bookmark_depth = None;
        let mut bookmarks = Vec::new();

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("invalid bookmark XML: {error}")))?;
            let text_element = is_bound(&namespace, TEXT_NAMESPACE);
            match event {
                Event::Start(ref source) => {
                    depth = checked_depth(depth)?;
                    if bookmark_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "text:bookmark must be empty".to_string(),
                        ));
                    }
                    if text_element && source.local_name().as_ref() == b"bookmark" {
                        if bookmarks.len() >= MAX_BOOKMARKS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_BOOKMARKS} bookmarks"
                            )));
                        }
                        let mut element = Element::new("text:bookmark");
                        copy_canonical_attributes(&reader, source, &mut element, "bookmark")?;
                        if element.get_attribute("text:name").is_none() {
                            return Err(Error::InvalidFormat(
                                "text:bookmark requires text:name".to_string(),
                            ));
                        }
                        bookmarks.push(Bookmark::from_element(element)?);
                        bookmark_depth = Some(depth);
                    }
                },
                Event::End(_) => {
                    if bookmark_depth == Some(depth) {
                        bookmark_depth = None;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("bookmark XML stack underflow".to_string())
                    })?;
                },
                Event::Text(_) | Event::CData(_) | Event::GeneralRef(_)
                    if bookmark_depth.is_some() =>
                {
                    return Err(Error::InvalidFormat(
                        "text:bookmark must be empty".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
        if depth != 0 || bookmark_depth.is_some() {
            return Err(Error::InvalidFormat(
                "incomplete bookmark XML structure".to_string(),
            ));
        }
        Ok(bookmarks)
    }

    /// Parse bookmark ranges (start/end pairs) from XML content.
    ///
    /// Positions are reported as zero-based paragraph/heading and character offsets.
    pub fn parse_bookmark_ranges(xml_content: &str) -> Result<Vec<BookmarkRange>> {
        writing::validate_bookmark_xml(xml_content)?;
        let mut reader = NsReader::from_str(xml_content);
        let mut buffer = Vec::new();
        let mut document_depth = 0usize;
        let mut paragraph_index = 0usize;
        let mut paragraph: Option<(usize, usize, usize)> = None;
        let mut marker_depth = None;
        let mut ranges = Vec::new();

        loop {
            let (namespace, event) =
                reader
                    .read_resolved_event_into(&mut buffer)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid bookmark range XML: {error}"))
                    })?;
            let text_element = is_bound(&namespace, TEXT_NAMESPACE);
            match event {
                Event::Start(ref element) => {
                    document_depth = checked_depth(document_depth)?;
                    if marker_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "bookmark range markers must be empty".to_string(),
                        ));
                    }
                    let range_marker = text_element
                        && matches!(
                            element.local_name().as_ref(),
                            b"bookmark-start" | b"bookmark-end"
                        );
                    if let Some(active) = paragraph.as_mut() {
                        let location = Some((active.0, active.1));
                        record_range_marker(&reader, text_element, element, location, &mut ranges)?;
                        if text_element {
                            add_control_offset(&reader, element, active)?;
                        }
                        active.2 += 1;
                    } else if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                        paragraph = Some((paragraph_index, 0, 1));
                        paragraph_index = paragraph_index.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat("bookmark paragraph count overflow".to_string())
                        })?;
                    } else {
                        record_range_marker(&reader, text_element, element, None, &mut ranges)?;
                    }
                    if range_marker {
                        marker_depth = Some(document_depth);
                    }
                },
                Event::Empty(ref element) => {
                    let location = paragraph.map(|(index, offset, _)| (index, offset));
                    record_range_marker(&reader, text_element, element, location, &mut ranges)?;
                    if text_element && let Some(active) = paragraph.as_mut() {
                        add_control_offset(&reader, element, active)?;
                    }
                },
                Event::Text(_) | Event::CData(_) | Event::GeneralRef(_)
                    if marker_depth.is_some() =>
                {
                    return Err(Error::InvalidFormat(
                        "bookmark range markers must be empty".to_string(),
                    ));
                },
                Event::Text(ref value) if paragraph.is_some() => {
                    let value = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid bookmark range text: {error}"))
                        })?;
                    add_character_offset(paragraph.as_mut().unwrap(), value.chars().count())?;
                },
                Event::CData(ref value) if paragraph.is_some() => {
                    let value = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid bookmark range CDATA: {error}"))
                        })?;
                    add_character_offset(paragraph.as_mut().unwrap(), value.chars().count())?;
                },
                Event::GeneralRef(ref reference) if paragraph.is_some() => {
                    let value = decode_reference(reference, "bookmark range")?;
                    add_character_offset(paragraph.as_mut().unwrap(), value.chars().count())?;
                },
                Event::End(_) => {
                    if marker_depth == Some(document_depth) {
                        marker_depth = None;
                    }
                    document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("bookmark range XML stack underflow".to_string())
                    })?;
                    if let Some((_, _, paragraph_depth)) = paragraph.as_mut() {
                        *paragraph_depth = paragraph_depth.checked_sub(1).ok_or_else(|| {
                            Error::InvalidFormat("bookmark paragraph stack underflow".to_string())
                        })?;
                        if *paragraph_depth == 0 {
                            paragraph = None;
                        }
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
        if document_depth != 0 || paragraph.is_some() || marker_depth.is_some() {
            return Err(Error::InvalidFormat(
                "incomplete bookmark range XML structure".to_string(),
            ));
        }
        Ok(ranges)
    }
}

fn record_range_marker(
    reader: &NsReader<&[u8]>,
    text_element: bool,
    element: &quick_xml::events::BytesStart<'_>,
    location: Option<(usize, usize)>,
    ranges: &mut Vec<BookmarkRange>,
) -> Result<()> {
    if !text_element
        || !matches!(
            element.local_name().as_ref(),
            b"bookmark-start" | b"bookmark-end"
        )
    {
        return Ok(());
    }
    let name = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"name", "bookmark range")?
        .ok_or_else(|| Error::InvalidFormat("bookmark range requires text:name".to_string()))?;
    if element.local_name().as_ref() == b"bookmark-start" {
        if ranges.len() >= MAX_BOOKMARKS {
            return Err(Error::InvalidFormat(format!(
                "document exceeds {MAX_BOOKMARKS} bookmark ranges"
            )));
        }
        let mut range = BookmarkRange::new(name);
        range.start = location;
        ranges.push(range);
    } else if let Some(range) = ranges
        .iter_mut()
        .rev()
        .find(|range| range.name == name && range.end.is_none())
    {
        range.end = location;
    } else {
        if ranges.len() >= MAX_BOOKMARKS {
            return Err(Error::InvalidFormat(format!(
                "document exceeds {MAX_BOOKMARKS} bookmark ranges"
            )));
        }
        ranges.push(BookmarkRange {
            name,
            start: None,
            end: location,
        });
    }
    Ok(())
}

fn add_control_offset(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    paragraph: &mut (usize, usize, usize),
) -> Result<()> {
    let mut text = String::new();
    append_text_control(reader, element, &mut text)?;
    add_character_offset(paragraph, text.chars().count())
}

fn add_character_offset(paragraph: &mut (usize, usize, usize), count: usize) -> Result<()> {
    paragraph.1 = paragraph
        .1
        .checked_add(count)
        .ok_or_else(|| Error::InvalidFormat("bookmark character offset overflow".to_string()))?;
    Ok(())
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("bookmark nesting depth overflow".to_string()))?;
    if depth > MAX_BOOKMARK_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "bookmark nesting exceeds {MAX_BOOKMARK_DEPTH} levels"
        )));
    }
    Ok(depth)
}
