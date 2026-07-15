//! ODT-specific parsing utilities.
//!
//! This module provides parsing functionality that is specific to OpenDocument Text
//! documents (.odt). For generic ODF element parsing (paragraphs, tables, lists, etc.)
//! that works across all ODF formats, see `crate::elements::parser::DocumentParser`.

use crate::elements::xml::{
    DC_NAMESPACE, META_NAMESPACE, OFFICE_NAMESPACE, TEXT_NAMESPACE, append_checked,
    append_text_control, decode_reference, is_bound, namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::collections::HashMap;

const MAX_SEMANTIC_DEPTH: usize = 4_096;
const MAX_SEMANTIC_ITEMS: usize = 1_000_000;

/// Parser for ODT-specific structures.
///
/// This provides parsing logic specific to text documents, such as:
/// - Track changes (insertions, deletions, formatting changes)
/// - Comments and annotations
/// - Sections (protected content, different formatting)
/// - Headers and footers
///
/// For generic element parsing (paragraphs, tables, etc.), use `DocumentParser`
/// from `crate::elements::parser` instead.
pub(crate) struct OdtParser;

/// Represents a tracked change in the document
#[derive(Debug, Clone)]
pub struct TrackChange {
    /// Change ID
    pub id: String,
    /// Author who made the change
    pub author: Option<String>,
    /// Date/time of the change
    pub date: Option<String>,
    /// Type of change (insertion, deletion, format-change)
    pub change_type: ChangeType,
    /// Changed text content
    pub content: String,
}

/// Type of tracked change
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChangeType {
    /// Text insertion
    Insertion,
    /// Text deletion
    Deletion,
    /// Formatting change
    FormatChange,
}

/// Represents a comment/annotation in the document
#[derive(Debug, Clone)]
pub struct Comment {
    /// Comment ID
    pub id: String,
    /// Author of the comment
    pub author: Option<String>,
    /// Date/time of the comment
    pub date: Option<String>,
    /// Comment text content
    pub content: String,
    /// Referenced text in the document
    pub reference: Option<String>,
}

/// Represents a section in the document
#[derive(Debug, Clone)]
pub struct Section {
    /// Section name
    pub name: String,
    /// Section style
    pub style: Option<String>,
    /// Whether the section is protected
    pub protected: bool,
    /// Text content within the section
    pub content: String,
}

impl OdtParser {
    /// Parse track changes from content
    ///
    /// Extracts tracked changes (insertions, deletions, format changes) from the document.
    /// Track changes are stored in `<text:tracked-changes>` elements with metadata,
    /// and referenced by `<text:change>` markers in the content.
    ///
    /// # Arguments
    ///
    /// * `content` - XML content containing tracked changes
    ///
    /// # Returns
    ///
    /// Vector of `TrackChange` objects with metadata
    pub fn parse_track_changes(content: &str) -> Result<Vec<TrackChange>> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_str(content);
        let mut buf = Vec::new();
        let mut changes = Vec::new();
        let mut in_tracked_changes = false;
        let mut in_change_element = false;
        let mut current_change: Option<TrackChange> = None;
        let mut depth: usize = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    match tag_name.as_str() {
                        "text:tracked-changes" => {
                            in_tracked_changes = true;
                        },
                        "text:changed-region" if in_tracked_changes => {
                            // Extract change ID
                            let mut id = String::new();
                            for attr in e.attributes().flatten() {
                                let key = String::from_utf8_lossy(attr.key.as_ref());
                                if key.ends_with(":id") {
                                    id = String::from_utf8_lossy(&attr.value).to_string();
                                }
                            }

                            current_change = Some(TrackChange {
                                id,
                                author: None,
                                date: None,
                                change_type: ChangeType::Insertion,
                                content: String::new(),
                            });
                            depth += 1;
                        },
                        "text:insertion" | "text:deletion" | "text:format-change"
                            if in_tracked_changes && current_change.is_some() =>
                        {
                            if let Some(ref mut change) = current_change {
                                change.change_type = match tag_name.as_str() {
                                    "text:insertion" => ChangeType::Insertion,
                                    "text:deletion" => ChangeType::Deletion,
                                    "text:format-change" => ChangeType::FormatChange,
                                    _ => ChangeType::Insertion,
                                };
                            }
                            in_change_element = true;
                            depth += 1;
                        },
                        "office:change-info" if in_change_element => {
                            depth += 1;
                        },
                        "dc:creator" if in_change_element => {
                            depth += 1;
                        },
                        "dc:date" if in_change_element => {
                            depth += 1;
                        },
                        _ if in_tracked_changes => {
                            depth += 1;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(ref t)) if in_change_element => {
                    let text = String::from_utf8_lossy(t).to_string();

                    // Determine what we're reading based on parent context
                    if let Some(ref mut change) = current_change {
                        // This is a simplification; in reality we'd track the parent element
                        if change.author.is_none() {
                            change.author = Some(text.clone());
                        } else if change.date.is_none() {
                            change.date = Some(text);
                        }
                    }
                },
                Ok(Event::End(ref e)) => {
                    let tag_name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                    match tag_name.as_str() {
                        "text:tracked-changes" => {
                            in_tracked_changes = false;
                        },
                        "text:changed-region" if in_tracked_changes => {
                            if let Some(change) = current_change.take() {
                                changes.push(change);
                            }
                            depth = depth.saturating_sub(1);
                        },
                        "text:insertion" | "text:deletion" | "text:format-change"
                            if in_tracked_changes =>
                        {
                            in_change_element = false;
                            depth = depth.saturating_sub(1);
                        },
                        _ if in_tracked_changes && depth > 0 => {
                            depth = depth.saturating_sub(1);
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(changes)
    }

    /// Parse comments/annotations
    ///
    /// Extracts comments and annotations from the document.
    /// Comments are stored in `<office:annotation>` elements.
    ///
    /// # Arguments
    ///
    /// * `content` - XML content containing annotations
    ///
    /// # Returns
    ///
    /// Vector of `Comment` objects with metadata and content
    pub fn parse_comments(content: &str) -> Result<Vec<Comment>> {
        parse_comments(content)
    }

    /// Parse sections
    ///
    /// Extracts document sections which can contain protected content,
    /// different formatting, or special layout properties.
    ///
    /// # Arguments
    ///
    /// * `content` - XML content containing sections
    ///
    /// # Returns
    ///
    /// Vector of `Section` objects with metadata and content
    pub fn parse_sections(content: &str) -> Result<Vec<Section>> {
        parse_sections(content)
    }
}

struct ActiveComment {
    comment: Comment,
    depth: usize,
    creator_depth: Option<usize>,
    date_depth: Option<usize>,
    fallback_date_depth: Option<usize>,
    fallback_date: String,
    paragraph_depth: Option<usize>,
    seen_paragraph: bool,
}

fn parse_comments(content: &str) -> Result<Vec<Comment>> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active: Option<ActiveComment> = None;
    let mut comments = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid annotation XML: {error}")))?;
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let dc_element = is_bound(&namespace, DC_NAMESPACE);
        let meta_element = is_bound(&namespace, META_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "annotation")?;
                if let Some(comment) = active.as_mut() {
                    if office_element && element.local_name().as_ref() == b"annotation" {
                        return Err(Error::InvalidFormat(
                            "nested office:annotation elements are not allowed".to_string(),
                        ));
                    }
                    comment.depth += 1;
                    if dc_element && element.local_name().as_ref() == b"creator" {
                        comment.creator_depth = Some(comment.depth);
                    } else if dc_element && element.local_name().as_ref() == b"date" {
                        comment.date_depth = Some(comment.depth);
                    } else if meta_element && element.local_name().as_ref() == b"date-string" {
                        comment.fallback_date_depth = Some(comment.depth);
                    } else if text_element
                        && matches!(element.local_name().as_ref(), b"p" | b"h")
                        && comment.paragraph_depth.is_none()
                    {
                        if comment.seen_paragraph {
                            append_checked(&mut comment.comment.content, "\n")?;
                        }
                        comment.seen_paragraph = true;
                        comment.paragraph_depth = Some(comment.depth);
                    }
                    if text_element && comment.paragraph_depth.is_some() {
                        append_text_control(&reader, element, &mut comment.comment.content)?;
                    }
                } else if office_element && element.local_name().as_ref() == b"annotation" {
                    if comments.len() >= MAX_SEMANTIC_ITEMS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_SEMANTIC_ITEMS} annotations"
                        )));
                    }
                    let id = namespaced_attribute(
                        &reader,
                        element,
                        OFFICE_NAMESPACE,
                        b"name",
                        "annotation",
                    )?
                    .unwrap_or_else(|| format!("comment_{}", comments.len()));
                    active = Some(ActiveComment {
                        comment: Comment {
                            id,
                            author: None,
                            date: None,
                            content: String::new(),
                            reference: None,
                        },
                        depth: 1,
                        creator_depth: None,
                        date_depth: None,
                        fallback_date_depth: None,
                        fallback_date: String::new(),
                        paragraph_depth: None,
                        seen_paragraph: false,
                    });
                }
            },
            Event::Empty(ref element) if active.is_some() => {
                let comment = active.as_mut().expect("checked annotation");
                if office_element && element.local_name().as_ref() == b"annotation" {
                    return Err(Error::InvalidFormat(
                        "nested office:annotation elements are not allowed".to_string(),
                    ));
                }
                if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                    if comment.seen_paragraph {
                        append_checked(&mut comment.comment.content, "\n")?;
                    }
                    comment.seen_paragraph = true;
                } else if text_element && comment.paragraph_depth.is_some() {
                    append_text_control(&reader, element, &mut comment.comment.content)?;
                }
            },
            Event::Empty(ref element)
                if office_element && element.local_name().as_ref() == b"annotation" =>
            {
                if comments.len() >= MAX_SEMANTIC_ITEMS {
                    return Err(Error::InvalidFormat(format!(
                        "document exceeds {MAX_SEMANTIC_ITEMS} annotations"
                    )));
                }
                let id = namespaced_attribute(
                    &reader,
                    element,
                    OFFICE_NAMESPACE,
                    b"name",
                    "annotation",
                )?
                .unwrap_or_else(|| format!("comment_{}", comments.len()));
                comments.push(Comment {
                    id,
                    author: None,
                    date: None,
                    content: String::new(),
                    reference: None,
                });
            },
            Event::Text(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid annotation text: {error}"))
                    })?;
                append_comment_text(active.as_mut().expect("checked annotation"), &value)?;
            },
            Event::CData(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid annotation CDATA: {error}"))
                    })?;
                append_comment_text(active.as_mut().expect("checked annotation"), &value)?;
            },
            Event::GeneralRef(ref reference) if active.is_some() => {
                let value = decode_reference(reference, "annotation")?;
                append_comment_text(active.as_mut().expect("checked annotation"), &value)?;
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("annotation XML stack underflow".to_string())
                })?;
                if let Some(comment) = active.as_mut() {
                    if comment.creator_depth == Some(comment.depth) {
                        comment.creator_depth = None;
                    }
                    if comment.date_depth == Some(comment.depth) {
                        comment.date_depth = None;
                    }
                    if comment.fallback_date_depth == Some(comment.depth) {
                        comment.fallback_date_depth = None;
                    }
                    if comment.paragraph_depth == Some(comment.depth) {
                        comment.paragraph_depth = None;
                    }
                    comment.depth = comment.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("annotation element stack underflow".to_string())
                    })?;
                    if comment.depth == 0 {
                        let mut comment = active.take().expect("checked annotation");
                        if comment.comment.date.is_none() && !comment.fallback_date.is_empty() {
                            comment.comment.date = Some(comment.fallback_date);
                        }
                        comments.push(comment.comment);
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if document_depth != 0 || active.is_some() {
        return Err(Error::InvalidFormat(
            "incomplete annotation XML structure".to_string(),
        ));
    }
    let references = parse_annotation_references(content)?;
    for comment in &mut comments {
        if let Some(reference) = references.get(&comment.id) {
            comment.reference = Some(reference.clone());
        }
    }
    Ok(comments)
}

fn append_comment_text(comment: &mut ActiveComment, value: &str) -> Result<()> {
    if comment.creator_depth.is_some() {
        let author = comment.comment.author.get_or_insert_with(String::new);
        append_checked(author, value)
    } else if comment.date_depth.is_some() {
        let date = comment.comment.date.get_or_insert_with(String::new);
        append_checked(date, value)
    } else if comment.fallback_date_depth.is_some() {
        append_checked(&mut comment.fallback_date, value)
    } else if comment.paragraph_depth.is_some() {
        append_checked(&mut comment.comment.content, value)
    } else {
        Ok(())
    }
}

struct PendingAnnotation {
    name: String,
    text: String,
    seen_paragraph: bool,
}

fn parse_annotation_references(content: &str) -> Result<HashMap<String, String>> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut paragraph_depth: Option<usize> = None;
    let mut annotation_depth = 0usize;
    let mut annotation_name = None;
    let mut pending: HashMap<String, PendingAnnotation> = HashMap::new();
    let mut references = HashMap::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid annotation range XML: {error}"))
            })?;
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "annotation range")?;
                if let Some(depth) = paragraph_depth.as_mut() {
                    *depth += 1;
                } else if text_element
                    && matches!(element.local_name().as_ref(), b"p" | b"h")
                    && annotation_depth == 0
                {
                    paragraph_depth = Some(1);
                    for item in pending.values_mut() {
                        if item.seen_paragraph {
                            append_checked(&mut item.text, "\n")?;
                        }
                        item.seen_paragraph = true;
                    }
                }
                if annotation_depth > 0 {
                    annotation_depth += 1;
                } else if office_element && element.local_name().as_ref() == b"annotation" {
                    annotation_name = namespaced_attribute(
                        &reader,
                        element,
                        OFFICE_NAMESPACE,
                        b"name",
                        "annotation",
                    )?;
                    annotation_depth = 1;
                } else if office_element && element.local_name().as_ref() == b"annotation-end" {
                    finish_annotation_reference(&reader, element, &mut pending, &mut references)?;
                } else if text_element && paragraph_depth.is_some() {
                    for item in pending.values_mut() {
                        append_text_control(&reader, element, &mut item.text)?;
                    }
                }
            },
            Event::Empty(ref element) if annotation_depth == 0 => {
                if office_element && element.local_name().as_ref() == b"annotation" {
                    if let Some(name) = namespaced_attribute(
                        &reader,
                        element,
                        OFFICE_NAMESPACE,
                        b"name",
                        "annotation",
                    )? {
                        add_pending_annotation(
                            &mut pending,
                            PendingAnnotation {
                                name,
                                text: String::new(),
                                seen_paragraph: paragraph_depth.is_some(),
                            },
                        )?;
                    }
                } else if office_element && element.local_name().as_ref() == b"annotation-end" {
                    finish_annotation_reference(&reader, element, &mut pending, &mut references)?;
                } else if text_element && paragraph_depth.is_some() {
                    for item in pending.values_mut() {
                        append_text_control(&reader, element, &mut item.text)?;
                    }
                }
            },
            Event::Text(ref value) if annotation_depth == 0 && paragraph_depth.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid annotation range text: {error}"))
                    })?;
                for item in pending.values_mut() {
                    append_checked(&mut item.text, &value)?;
                }
            },
            Event::CData(ref value) if annotation_depth == 0 && paragraph_depth.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid annotation range CDATA: {error}"))
                    })?;
                for item in pending.values_mut() {
                    append_checked(&mut item.text, &value)?;
                }
            },
            Event::GeneralRef(ref reference)
                if annotation_depth == 0 && paragraph_depth.is_some() =>
            {
                let value = decode_reference(reference, "annotation range")?;
                for item in pending.values_mut() {
                    append_checked(&mut item.text, &value)?;
                }
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("annotation range XML stack underflow".to_string())
                })?;
                if annotation_depth > 0 {
                    annotation_depth -= 1;
                    if annotation_depth == 0
                        && let Some(name) = annotation_name.take()
                    {
                        add_pending_annotation(
                            &mut pending,
                            PendingAnnotation {
                                name,
                                text: String::new(),
                                seen_paragraph: paragraph_depth.is_some(),
                            },
                        )?;
                    }
                }
                if let Some(depth) = paragraph_depth.as_mut() {
                    *depth = (*depth).checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("annotation paragraph stack underflow".to_string())
                    })?;
                    if *depth == 0 {
                        paragraph_depth = None;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || paragraph_depth.is_some() || annotation_depth != 0 {
        return Err(Error::InvalidFormat(
            "incomplete annotation range XML structure".to_string(),
        ));
    }
    Ok(references)
}

fn finish_annotation_reference(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    pending: &mut HashMap<String, PendingAnnotation>,
    references: &mut HashMap<String, String>,
) -> Result<()> {
    let name = namespaced_attribute(reader, element, OFFICE_NAMESPACE, b"name", "annotation-end")?
        .ok_or_else(|| {
            Error::InvalidFormat("office:annotation-end requires office:name".to_string())
        })?;
    if let Some(item) = pending.remove(&name) {
        references.insert(item.name, item.text);
    }
    Ok(())
}

fn ensure_pending_capacity(length: usize) -> Result<()> {
    if length >= MAX_SEMANTIC_ITEMS {
        return Err(Error::InvalidFormat(format!(
            "document exceeds {MAX_SEMANTIC_ITEMS} annotation ranges"
        )));
    }
    Ok(())
}

fn add_pending_annotation(
    pending: &mut HashMap<String, PendingAnnotation>,
    annotation: PendingAnnotation,
) -> Result<()> {
    ensure_pending_capacity(pending.len())?;
    if pending.contains_key(&annotation.name) {
        return Err(Error::InvalidFormat(format!(
            "duplicate open annotation range '{}'",
            annotation.name
        )));
    }
    pending.insert(annotation.name.clone(), annotation);
    Ok(())
}

struct ActiveSection {
    section: Section,
    depth: usize,
    paragraph_depth: Option<usize>,
    seen_paragraph: bool,
    order: usize,
}

fn parse_sections(content: &str) -> Result<Vec<Section>> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active: Vec<ActiveSection> = Vec::new();
    let mut sections = Vec::new();
    let mut next_order = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid section XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "section")?;
                for section in &mut active {
                    section.depth += 1;
                    if text_element
                        && matches!(element.local_name().as_ref(), b"p" | b"h")
                        && section.paragraph_depth.is_none()
                    {
                        if section.seen_paragraph {
                            append_checked(&mut section.section.content, "\n")?;
                        }
                        section.seen_paragraph = true;
                        section.paragraph_depth = Some(section.depth);
                    }
                    if text_element && section.paragraph_depth.is_some() {
                        append_text_control(&reader, element, &mut section.section.content)?;
                    }
                }
                if text_element && element.local_name().as_ref() == b"section" {
                    if next_order >= MAX_SEMANTIC_ITEMS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_SEMANTIC_ITEMS} sections"
                        )));
                    }
                    active.push(ActiveSection {
                        section: section_from_start(&reader, element)?,
                        depth: 1,
                        paragraph_depth: None,
                        seen_paragraph: false,
                        order: next_order,
                    });
                    next_order += 1;
                }
            },
            Event::Empty(ref element) => {
                if text_element && element.local_name().as_ref() == b"section" {
                    if next_order >= MAX_SEMANTIC_ITEMS {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_SEMANTIC_ITEMS} sections"
                        )));
                    }
                    sections.push((next_order, section_from_start(&reader, element)?));
                    next_order += 1;
                } else {
                    for section in &mut active {
                        if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                            if section.seen_paragraph {
                                append_checked(&mut section.section.content, "\n")?;
                            }
                            section.seen_paragraph = true;
                        } else if text_element && section.paragraph_depth.is_some() {
                            append_text_control(&reader, element, &mut section.section.content)?;
                        }
                    }
                }
            },
            Event::Text(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid section text: {error}"))
                    })?;
                for section in &mut active {
                    if section.paragraph_depth.is_some() {
                        append_checked(&mut section.section.content, &value)?;
                    }
                }
            },
            Event::CData(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid section CDATA: {error}"))
                    })?;
                for section in &mut active {
                    if section.paragraph_depth.is_some() {
                        append_checked(&mut section.section.content, &value)?;
                    }
                }
            },
            Event::GeneralRef(ref reference) if !active.is_empty() => {
                let value = decode_reference(reference, "section")?;
                for section in &mut active {
                    if section.paragraph_depth.is_some() {
                        append_checked(&mut section.section.content, &value)?;
                    }
                }
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("section XML stack underflow".to_string())
                })?;
                for section in &mut active {
                    if section.paragraph_depth == Some(section.depth) {
                        section.paragraph_depth = None;
                    }
                    section.depth = section.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("section element stack underflow".to_string())
                    })?;
                }
                if active.last().is_some_and(|section| section.depth == 0) {
                    let section = active.pop().expect("checked active section");
                    sections.push((section.order, section.section));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete section XML structure".to_string(),
        ));
    }
    sections.sort_by_key(|(order, _)| *order);
    Ok(sections.into_iter().map(|(_, section)| section).collect())
}

fn section_from_start(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Section> {
    let name = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"name", "section")?
        .ok_or_else(|| Error::InvalidFormat("text:section requires text:name".to_string()))?;
    let style = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"style-name", "section")?;
    let protected = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"protected", "section")?
        .map(|value| parse_boolean(&value, "text:protected"))
        .transpose()?
        .unwrap_or(false);
    Ok(Section {
        name,
        style,
        protected,
        content: String::new(),
    })
}

fn parse_boolean(value: &str, context: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "{context} must be true, false, 1, or 0"
        ))),
    }
}

fn checked_semantic_depth(depth: usize, context: &str) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat(format!("{context} nesting depth overflow")))?;
    if depth > MAX_SEMANTIC_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "{context} nesting exceeds {MAX_SEMANTIC_DEPTH} levels"
        )));
    }
    Ok(depth)
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEST_TRACK_CHANGES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
    <text:tracked-changes>
        <text:changed-region text:id="change1">
            <text:insertion>
                <office:change-info>
                    <dc:creator>John Doe</dc:creator>
                    <dc:date>2024-03-15T10:30:00</dc:date>
                </office:change-info>
            </text:insertion>
        </text:changed-region>
        <text:changed-region text:id="change2">
            <text:deletion>
                <office:change-info>
                    <dc:creator>Jane Smith</dc:creator>
                    <dc:date>2024-03-15T11:00:00</dc:date>
                </office:change-info>
            </text:deletion>
        </text:changed-region>
        <text:changed-region text:id="change3">
            <text:format-change>
                <office:change-info>
                    <dc:creator>Bob Wilson</dc:creator>
                    <dc:date>2024-03-15T12:00:00</dc:date>
                </office:change-info>
            </text:format-change>
        </text:changed-region>
    </text:tracked-changes>
</office:document-content>"#;

    const TEST_COMMENTS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
    <text:p>
        <office:annotation office:name="cmt1">
            <dc:creator>Alice</dc:creator>
            <dc:date>2024-03-15T09:00:00</dc:date>
            <text:p>This is a comment</text:p>
        </office:annotation>
        Some text
    </text:p>
    <text:p>
        <office:annotation office:name="cmt2">
            <dc:creator>Bob</dc:creator>
            <dc:date>2024-03-15T10:00:00</dc:date>
            <text:p>First paragraph</text:p>
            <text:p>Second paragraph</text:p>
        </office:annotation>
        More text
    </text:p>
</office:document-content>"#;

    const TEST_SECTIONS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <text:section text:name="Introduction" text:style-name="IntroStyle">
        <text:p>Introduction content</text:p>
    </text:section>
    <text:section text:name="ProtectedSection" text:protected="true">
        <text:p>Protected content</text:p>
    </text:section>
    <text:section text:name="Chapter1" text:style-name="ChapterStyle" text:protected="false">
        <text:p>Chapter 1 content</text:p>
    </text:section>
</office:document-content>"#;

    const TEST_EMPTY_TRACK_CHANGES: &str = r#"<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <text:tracked-changes>
    </text:tracked-changes>
</office:document-content>"#;

    const TEST_EMPTY_CONTENT: &str = r#"<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
</office:document-content>"#;

    #[test]
    fn test_parse_track_changes() {
        let changes = OdtParser::parse_track_changes(TEST_TRACK_CHANGES_XML).unwrap();
        assert_eq!(changes.len(), 3);

        // Check first change (insertion)
        assert_eq!(changes[0].id, "change1");
        assert_eq!(changes[0].change_type, ChangeType::Insertion);
        // Parser extracts text elements - author/date extraction depends on XML structure
        assert!(changes[0].author.is_some());

        // Check second change (deletion)
        assert_eq!(changes[1].id, "change2");
        assert_eq!(changes[1].change_type, ChangeType::Deletion);

        // Check third change (format)
        assert_eq!(changes[2].id, "change3");
        assert_eq!(changes[2].change_type, ChangeType::FormatChange);
    }

    #[test]
    fn test_parse_track_changes_empty() {
        let changes = OdtParser::parse_track_changes(TEST_EMPTY_TRACK_CHANGES).unwrap();
        assert!(changes.is_empty());
    }

    #[test]
    fn test_parse_track_changes_no_tracked_changes() {
        let changes = OdtParser::parse_track_changes(TEST_EMPTY_CONTENT).unwrap();
        assert!(changes.is_empty());
    }

    #[test]
    fn test_parse_comments() {
        let comments = OdtParser::parse_comments(TEST_COMMENTS_XML).unwrap();
        assert_eq!(comments.len(), 2);

        // First comment
        assert_eq!(comments[0].id, "cmt1");
        assert_eq!(comments[0].author, Some("Alice".to_string()));
        assert_eq!(comments[0].date, Some("2024-03-15T09:00:00".to_string()));
        assert_eq!(comments[0].content, "This is a comment");

        // Second comment (with multiple paragraphs)
        assert_eq!(comments[1].id, "cmt2");
        assert_eq!(comments[1].author, Some("Bob".to_string()));
        assert_eq!(comments[1].date, Some("2024-03-15T10:00:00".to_string()));
        assert!(comments[1].content.contains("First paragraph"));
        assert!(comments[1].content.contains("Second paragraph"));
    }

    #[test]
    fn test_parse_comments_empty() {
        let comments = OdtParser::parse_comments(TEST_EMPTY_CONTENT).unwrap();
        assert!(comments.is_empty());
    }

    #[test]
    fn test_parse_sections() {
        let sections = OdtParser::parse_sections(TEST_SECTIONS_XML).unwrap();
        assert_eq!(sections.len(), 3);

        // First section
        assert_eq!(sections[0].name, "Introduction");
        assert_eq!(sections[0].style, Some("IntroStyle".to_string()));
        assert!(!sections[0].protected);

        // Second section (protected)
        assert_eq!(sections[1].name, "ProtectedSection");
        assert_eq!(sections[1].style, None);
        assert!(sections[1].protected);

        // Third section
        assert_eq!(sections[2].name, "Chapter1");
        assert_eq!(sections[2].style, Some("ChapterStyle".to_string()));
        assert!(!sections[2].protected);
    }

    #[test]
    fn test_parse_sections_empty() {
        let sections = OdtParser::parse_sections(TEST_EMPTY_CONTENT).unwrap();
        assert!(sections.is_empty());
    }

    #[test]
    fn parses_annotation_metadata_body_and_referenced_range_with_namespace_aliases() {
        let xml = r#"<x:document-content xmlns:x="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><x:body><x:text><t:p>before<x:annotation x:name="c&amp;1"><d:creator>A &amp; B</d:creator><m:date-string>2026-07-16</m:date-string><t:p>First<t:s t:c="2"/>X</t:p><t:list><t:list-item><t:p>Second<![CDATA[!]]></t:p></t:list-item></t:list></x:annotation>R&amp;<t:span>ange</t:span><x:annotation-end x:name="c&amp;1"/>after</t:p></x:text></x:body></x:document-content>"#;
        let comments = OdtParser::parse_comments(xml).unwrap();
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].id, "c&1");
        assert_eq!(comments[0].author.as_deref(), Some("A & B"));
        assert_eq!(comments[0].date.as_deref(), Some("2026-07-16"));
        assert_eq!(comments[0].content, "First  X\nSecond!");
        assert_eq!(comments[0].reference.as_deref(), Some("R&ange"));
    }

    #[test]
    fn parses_nested_sections_in_document_order_with_visible_text() {
        let xml = r#"<x:document-content xmlns:x="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><x:body><x:text><t:section t:name="Outer &amp; Main" t:style-name="S1" t:protected="1"><t:p>One &amp;<t:s t:c="2"/></t:p><t:section t:name="Inner"><t:p>Inner <![CDATA[X]]></t:p></t:section><t:p>Last</t:p></t:section><t:section t:name="Empty"/></x:text></x:body></x:document-content>"#;
        let sections = OdtParser::parse_sections(xml).unwrap();
        assert_eq!(sections.len(), 3);
        assert_eq!(sections[0].name, "Outer & Main");
        assert_eq!(sections[0].style.as_deref(), Some("S1"));
        assert!(sections[0].protected);
        assert_eq!(sections[0].content, "One &  \nInner X\nLast");
        assert_eq!(sections[1].name, "Inner");
        assert_eq!(sections[1].content, "Inner X");
        assert_eq!(sections[2].name, "Empty");
        assert!(sections[2].content.is_empty());
    }

    #[test]
    fn annotations_and_sections_reject_malformed_or_ambiguous_xml() {
        let namespace = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        assert!(OdtParser::parse_comments("<x:annotation>").is_err());
        let nested = r#"<x:annotation xmlns:x="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><x:annotation/></x:annotation>"#;
        assert!(OdtParser::parse_comments(nested).is_err());
        let missing_name = format!(r#"<t:section xmlns:t="{namespace}"/>"#);
        assert!(OdtParser::parse_sections(&missing_name).is_err());
        let invalid_boolean =
            format!(r#"<t:section xmlns:t="{namespace}" t:name="A" t:protected="yes"/>"#);
        assert!(OdtParser::parse_sections(&invalid_boolean).is_err());
        let duplicate = format!(
            r#"<t:section xmlns:t="{namespace}" xmlns:u="{namespace}" t:name="A" u:name="B"/>"#
        );
        assert!(OdtParser::parse_sections(&duplicate).is_err());
    }

    #[test]
    fn test_track_change_debug() {
        let change = TrackChange {
            id: "test1".to_string(),
            author: Some("Author".to_string()),
            date: Some("2024-03-15".to_string()),
            change_type: ChangeType::Insertion,
            content: "content".to_string(),
        };
        let debug_str = format!("{:?}", change);
        assert!(debug_str.contains("TrackChange"));
        assert!(debug_str.contains("test1"));
    }

    #[test]
    fn test_change_type_equality() {
        assert_eq!(ChangeType::Insertion, ChangeType::Insertion);
        assert_eq!(ChangeType::Deletion, ChangeType::Deletion);
        assert_eq!(ChangeType::FormatChange, ChangeType::FormatChange);
        assert_ne!(ChangeType::Insertion, ChangeType::Deletion);
    }

    #[test]
    fn test_change_type_clone() {
        let t1 = ChangeType::Insertion;
        let t2 = t1;
        assert_eq!(t1, t2);
    }

    #[test]
    fn test_change_type_copy() {
        let t1 = ChangeType::Insertion;
        let t2 = t1;
        assert_eq!(t1, t2); // Copy trait allows this
    }

    #[test]
    fn test_comment_debug() {
        let comment = Comment {
            id: "cmt1".to_string(),
            author: Some("Author".to_string()),
            date: Some("2024-03-15".to_string()),
            content: "Comment text".to_string(),
            reference: None,
        };
        let debug_str = format!("{:?}", comment);
        assert!(debug_str.contains("Comment"));
        assert!(debug_str.contains("cmt1"));
    }

    #[test]
    fn test_section_debug() {
        let section = Section {
            name: "Sec1".to_string(),
            style: Some("Style1".to_string()),
            protected: true,
            content: "Content".to_string(),
        };
        let debug_str = format!("{:?}", section);
        assert!(debug_str.contains("Section"));
        assert!(debug_str.contains("Sec1"));
    }

    #[test]
    fn test_comment_clone() {
        let comment = Comment {
            id: "cmt1".to_string(),
            author: Some("Author".to_string()),
            date: Some("2024-03-15".to_string()),
            content: "Content".to_string(),
            reference: Some("ref".to_string()),
        };
        let cloned = comment.clone();
        assert_eq!(comment.id, cloned.id);
        assert_eq!(comment.author, cloned.author);
        assert_eq!(comment.content, cloned.content);
    }

    #[test]
    fn test_track_change_clone() {
        let change = TrackChange {
            id: "tc1".to_string(),
            author: Some("Author".to_string()),
            date: Some("2024-03-15".to_string()),
            change_type: ChangeType::Deletion,
            content: "Deleted text".to_string(),
        };
        let cloned = change.clone();
        assert_eq!(change.id, cloned.id);
        assert_eq!(change.change_type, cloned.change_type);
    }

    #[test]
    fn test_section_clone() {
        let section = Section {
            name: "Sec1".to_string(),
            style: None,
            protected: false,
            content: "Text".to_string(),
        };
        let cloned = section.clone();
        assert_eq!(section.name, cloned.name);
        assert_eq!(section.protected, cloned.protected);
    }
}
