//! ODT-specific parsing utilities.
//!
//! This module provides parsing functionality that is specific to OpenDocument Text
//! documents (.odt). For generic ODF element parsing (paragraphs, tables, lists, etc.)
//! that works across all ODF formats, see `crate::elements::parser::DocumentParser`.

use crate::elements::xml::{
    DC_NAMESPACE, META_NAMESPACE, OFFICE_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE, XML_NAMESPACE,
    append_checked, append_text_control, decode_reference, is_bound, namespaced_attribute,
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
    /// Optional XML identifier.
    pub xml_id: Option<String>,
    /// Stored protection-key material; never used to unlock content automatically.
    pub protection_key: Option<String>,
    /// Digest algorithm URI for the protection key.
    pub protection_key_digest_algorithm: Option<String>,
    /// Visibility behavior.
    pub display: SectionDisplay,
    /// Inert condition expression for conditionally displayed sections.
    pub condition: Option<String>,
    /// Optional linked-section source; never fetched.
    pub source: Option<SectionSource>,
    /// Optional DDE source; never activated.
    pub dde_source: Option<SectionDdeSource>,
    /// Text content within the section
    pub content: String,
}

/// Section visibility behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionDisplay {
    /// Display normally.
    Visible,
    /// Do not display.
    Hidden,
    /// Display according to the stored inert condition.
    Condition,
}

/// An inert linked-section source declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionSource {
    /// External or package-local URI. It is never fetched automatically.
    pub href: Option<String>,
    /// Named section within the source document.
    pub section_name: Option<String>,
    /// Producer-specific import filter name.
    pub filter_name: Option<String>,
}

/// An inert Dynamic Data Exchange source declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SectionDdeSource {
    /// DDE source name.
    pub name: Option<String>,
    /// Stored conversion mode.
    pub conversion_mode: Option<String>,
    /// Whether the producer requested automatic updates; no update is performed.
    pub automatic_update: Option<bool>,
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
        let mut changes = parse_change_declarations(content)?;
        correlate_change_ranges(content, &mut changes)?;
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

struct ActiveTrackedChange {
    id: String,
    author: Option<String>,
    date: Option<String>,
    change_type: Option<ChangeType>,
    content: String,
    depth: usize,
    kind_depth: Option<usize>,
    change_info_depth: Option<usize>,
    change_info_seen: bool,
    creator_depth: Option<usize>,
    date_depth: Option<usize>,
    paragraph_depth: Option<usize>,
    seen_paragraph: bool,
}

fn parse_change_declarations(content: &str) -> Result<Vec<TrackChange>> {
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut tracked_depth = 0usize;
    let mut tracked_changes_seen = false;
    let mut active: Option<ActiveTrackedChange> = None;
    let mut changes = Vec::new();
    let mut ids = HashMap::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid tracked-change XML: {error}"))
            })?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        let dc_element = is_bound(&namespace, DC_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "tracked change")?;
                if tracked_depth == 0 {
                    if text_element && element.local_name().as_ref() == b"tracked-changes" {
                        if tracked_changes_seen {
                            return Err(Error::InvalidFormat(
                                "multiple text:tracked-changes elements are not allowed"
                                    .to_string(),
                            ));
                        }
                        tracked_changes_seen = true;
                        tracked_depth = 1;
                    }
                } else {
                    tracked_depth = checked_semantic_depth(tracked_depth, "tracked change")?;
                    if let Some(change) = active.as_mut() {
                        change.depth = checked_semantic_depth(change.depth, "changed region")?;
                        process_change_declaration_start(
                            &reader,
                            element,
                            text_element,
                            office_element,
                            dc_element,
                            change,
                        )?;
                    } else if text_element && element.local_name().as_ref() == b"changed-region" {
                        if tracked_depth != 2 {
                            return Err(Error::InvalidFormat(
                                "text:changed-region must be a direct child of text:tracked-changes"
                                    .to_string(),
                            ));
                        }
                        if changes.len() >= MAX_SEMANTIC_ITEMS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_SEMANTIC_ITEMS} tracked changes"
                            )));
                        }
                        let id = change_region_id(&reader, element)?;
                        if ids.insert(id.clone(), changes.len()).is_some() {
                            return Err(Error::InvalidFormat(format!(
                                "duplicate tracked-change ID '{id}'"
                            )));
                        }
                        active = Some(ActiveTrackedChange {
                            id,
                            author: None,
                            date: None,
                            change_type: None,
                            content: String::new(),
                            depth: 1,
                            kind_depth: None,
                            change_info_depth: None,
                            change_info_seen: false,
                            creator_depth: None,
                            date_depth: None,
                            paragraph_depth: None,
                            seen_paragraph: false,
                        });
                    }
                }
            },
            Event::Empty(ref element) if tracked_depth > 0 => {
                if text_element && element.local_name().as_ref() == b"changed-region" {
                    return Err(Error::InvalidFormat(
                        "text:changed-region requires a change declaration".to_string(),
                    ));
                }
                if let Some(change) = active.as_mut() {
                    process_change_declaration_empty(
                        &reader,
                        element,
                        text_element,
                        office_element,
                        dc_element,
                        change,
                    )?;
                }
            },
            Event::Empty(ref element)
                if text_element && element.local_name().as_ref() == b"tracked-changes" =>
            {
                if tracked_changes_seen {
                    return Err(Error::InvalidFormat(
                        "multiple text:tracked-changes elements are not allowed".to_string(),
                    ));
                }
                tracked_changes_seen = true;
            },
            Event::Text(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid tracked-change text: {error}"))
                    })?;
                append_change_declaration_text(active.as_mut().expect("checked change"), &value)?;
            },
            Event::CData(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid tracked-change CDATA: {error}"))
                    })?;
                append_change_declaration_text(active.as_mut().expect("checked change"), &value)?;
            },
            Event::GeneralRef(ref reference) if active.is_some() => {
                let value = decode_reference(reference, "tracked change")?;
                append_change_declaration_text(active.as_mut().expect("checked change"), &value)?;
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("tracked-change XML stack underflow".to_string())
                })?;
                if tracked_depth > 0 {
                    if let Some(change) = active.as_mut() {
                        if change.creator_depth == Some(change.depth) {
                            change.creator_depth = None;
                        }
                        if change.date_depth == Some(change.depth) {
                            change.date_depth = None;
                        }
                        if change.paragraph_depth == Some(change.depth) {
                            change.paragraph_depth = None;
                        }
                        if change.change_info_depth == Some(change.depth) {
                            change.change_info_depth = None;
                        }
                        if change.kind_depth == Some(change.depth) {
                            change.kind_depth = None;
                        }
                        change.depth = change.depth.checked_sub(1).ok_or_else(|| {
                            Error::InvalidFormat("changed-region stack underflow".to_string())
                        })?;
                        if change.depth == 0 {
                            let change = active.take().expect("checked change");
                            let change_type = change.change_type.ok_or_else(|| {
                                Error::InvalidFormat(format!(
                                    "changed region '{}' has no change declaration",
                                    change.id
                                ))
                            })?;
                            if !change.change_info_seen {
                                return Err(Error::InvalidFormat(format!(
                                    "changed region '{}' has no office:change-info",
                                    change.id
                                )));
                            }
                            changes.push(TrackChange {
                                id: change.id,
                                author: change.author,
                                date: change.date,
                                change_type,
                                content: change.content,
                            });
                        }
                    }
                    tracked_depth = tracked_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("tracked-changes stack underflow".to_string())
                    })?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || tracked_depth != 0 || active.is_some() {
        return Err(Error::InvalidFormat(
            "incomplete tracked-change XML structure".to_string(),
        ));
    }
    Ok(changes)
}

fn change_region_id(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<String> {
    let text_id = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"id", "changed-region")?;
    let xml_id = namespaced_attribute(reader, element, XML_NAMESPACE, b"id", "changed-region")?;
    let id = text_id.or(xml_id).ok_or_else(|| {
        Error::InvalidFormat("text:changed-region requires text:id or xml:id".to_string())
    })?;
    if id.is_empty() {
        return Err(Error::InvalidFormat(
            "tracked-change ID must not be empty".to_string(),
        ));
    }
    Ok(id)
}

fn process_change_declaration_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_element: bool,
    office_element: bool,
    dc_element: bool,
    change: &mut ActiveTrackedChange,
) -> Result<()> {
    if text_element {
        let change_type = match element.local_name().as_ref() {
            b"insertion" => Some(ChangeType::Insertion),
            b"deletion" => Some(ChangeType::Deletion),
            b"format-change" => Some(ChangeType::FormatChange),
            _ => None,
        };
        if let Some(change_type) = change_type {
            if change.change_type.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "changed region '{}' has multiple change declarations",
                    change.id
                )));
            }
            if change.depth != 2 {
                return Err(Error::InvalidFormat(
                    "change declaration must be a direct child of text:changed-region".to_string(),
                ));
            }
            change.change_type = Some(change_type);
            change.kind_depth = Some(change.depth);
            return Ok(());
        }
    }
    if office_element && element.local_name().as_ref() == b"change-info" {
        if change.kind_depth != change.depth.checked_sub(1) || change.change_info_seen {
            return Err(Error::InvalidFormat(format!(
                "invalid office:change-info in changed region '{}'",
                change.id
            )));
        }
        change.change_info_seen = true;
        change.change_info_depth = Some(change.depth);
    } else if dc_element
        && element.local_name().as_ref() == b"creator"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) || change.author.is_some() {
            return Err(Error::InvalidFormat(format!(
                "invalid dc:creator in changed region '{}'",
                change.id
            )));
        }
        change.author = Some(String::new());
        change.creator_depth = Some(change.depth);
    } else if dc_element
        && element.local_name().as_ref() == b"date"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) || change.date.is_some() {
            return Err(Error::InvalidFormat(format!(
                "invalid dc:date in changed region '{}'",
                change.id
            )));
        }
        change.date = Some(String::new());
        change.date_depth = Some(change.depth);
    } else if text_element
        && matches!(element.local_name().as_ref(), b"p" | b"h")
        && change.change_type == Some(ChangeType::Deletion)
        && change.change_info_depth.is_none()
        && change.paragraph_depth.is_none()
    {
        if change.seen_paragraph {
            append_checked(&mut change.content, "\n")?;
        }
        change.seen_paragraph = true;
        change.paragraph_depth = Some(change.depth);
    }
    if text_element && change.paragraph_depth.is_some() {
        append_text_control(reader, element, &mut change.content)?;
    }
    Ok(())
}

fn process_change_declaration_empty(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_element: bool,
    office_element: bool,
    dc_element: bool,
    change: &mut ActiveTrackedChange,
) -> Result<()> {
    if text_element
        && matches!(
            element.local_name().as_ref(),
            b"insertion" | b"deletion" | b"format-change"
        )
    {
        return Err(Error::InvalidFormat(
            "change declaration requires office:change-info".to_string(),
        ));
    }
    if office_element && element.local_name().as_ref() == b"change-info" {
        if change.kind_depth != change.depth.checked_sub(1) || change.change_info_seen {
            return Err(Error::InvalidFormat(format!(
                "invalid office:change-info in changed region '{}'",
                change.id
            )));
        }
        change.change_info_seen = true;
    } else if dc_element
        && element.local_name().as_ref() == b"creator"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) || change.author.is_some() {
            return Err(Error::InvalidFormat(format!(
                "invalid dc:creator in changed region '{}'",
                change.id
            )));
        }
        change.author = Some(String::new());
    } else if dc_element
        && element.local_name().as_ref() == b"date"
        && change.change_info_depth.is_some()
    {
        if change.change_info_depth != change.depth.checked_sub(1) || change.date.is_some() {
            return Err(Error::InvalidFormat(format!(
                "invalid dc:date in changed region '{}'",
                change.id
            )));
        }
        change.date = Some(String::new());
    } else if text_element
        && matches!(element.local_name().as_ref(), b"p" | b"h")
        && change.change_type == Some(ChangeType::Deletion)
        && change.change_info_depth.is_none()
    {
        if change.seen_paragraph {
            append_checked(&mut change.content, "\n")?;
        }
        change.seen_paragraph = true;
    } else if text_element && change.paragraph_depth.is_some() {
        append_text_control(reader, element, &mut change.content)?;
    }
    Ok(())
}

fn append_change_declaration_text(change: &mut ActiveTrackedChange, value: &str) -> Result<()> {
    if change.creator_depth.is_some() {
        append_checked(change.author.as_mut().expect("creator initialized"), value)
    } else if change.date_depth.is_some() {
        append_checked(change.date.as_mut().expect("date initialized"), value)
    } else if change.paragraph_depth.is_some() {
        append_checked(&mut change.content, value)
    } else {
        Ok(())
    }
}

struct PendingChangeRange {
    text: String,
    seen_paragraph: bool,
}

#[derive(Default)]
struct ChangeRangeState {
    pending: HashMap<String, PendingChangeRange>,
    completed: HashMap<String, Vec<String>>,
    completed_count: usize,
}

fn correlate_change_ranges(content: &str, changes: &mut [TrackChange]) -> Result<()> {
    let change_types: HashMap<String, ChangeType> = changes
        .iter()
        .map(|change| (change.id.clone(), change.change_type))
        .collect();
    let mut reader = NsReader::from_str(content);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut tracked_depth = 0usize;
    let mut paragraph_depth: Option<usize> = None;
    let mut annotation_depth = 0usize;
    let mut ranges = ChangeRangeState::default();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid change-range XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "change range")?;
                if tracked_depth > 0 {
                    tracked_depth = checked_semantic_depth(tracked_depth, "change range")?;
                } else if text_element && element.local_name().as_ref() == b"tracked-changes" {
                    tracked_depth = 1;
                } else if annotation_depth > 0 {
                    annotation_depth = checked_semantic_depth(annotation_depth, "annotation")?;
                    if let Some(depth) = paragraph_depth.as_mut() {
                        *depth = checked_semantic_depth(*depth, "change-range paragraph")?;
                    }
                } else {
                    enter_change_paragraph(
                        text_element,
                        element,
                        &mut paragraph_depth,
                        &mut ranges.pending,
                    )?;
                    if office_element && element.local_name().as_ref() == b"annotation" {
                        annotation_depth = 1;
                    } else {
                        process_change_marker(
                            &reader,
                            element,
                            text_element,
                            paragraph_depth.is_some(),
                            &change_types,
                            &mut ranges,
                        )?;
                        if text_element && paragraph_depth.is_some() {
                            for range in ranges.pending.values_mut() {
                                append_text_control(&reader, element, &mut range.text)?;
                            }
                        }
                    }
                }
            },
            Event::Empty(ref element) if tracked_depth == 0 => {
                if annotation_depth > 0
                    || office_element && element.local_name().as_ref() == b"annotation"
                {
                    // Annotation definitions are metadata rather than visible range text.
                } else if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                    for range in ranges.pending.values_mut() {
                        if range.seen_paragraph {
                            append_checked(&mut range.text, "\n")?;
                        }
                        range.seen_paragraph = true;
                    }
                } else {
                    process_change_marker(
                        &reader,
                        element,
                        text_element,
                        paragraph_depth.is_some(),
                        &change_types,
                        &mut ranges,
                    )?;
                    if text_element && paragraph_depth.is_some() {
                        for range in ranges.pending.values_mut() {
                            append_text_control(&reader, element, &mut range.text)?;
                        }
                    }
                }
            },
            Event::Text(ref value)
                if tracked_depth == 0 && annotation_depth == 0 && paragraph_depth.is_some() =>
            {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid change-range text: {error}"))
                    })?;
                for range in ranges.pending.values_mut() {
                    append_checked(&mut range.text, &value)?;
                }
            },
            Event::CData(ref value)
                if tracked_depth == 0 && annotation_depth == 0 && paragraph_depth.is_some() =>
            {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid change-range CDATA: {error}"))
                    })?;
                for range in ranges.pending.values_mut() {
                    append_checked(&mut range.text, &value)?;
                }
            },
            Event::GeneralRef(ref reference)
                if tracked_depth == 0 && annotation_depth == 0 && paragraph_depth.is_some() =>
            {
                let value = decode_reference(reference, "change range")?;
                for range in ranges.pending.values_mut() {
                    append_checked(&mut range.text, &value)?;
                }
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("change-range XML stack underflow".to_string())
                })?;
                if tracked_depth > 0 {
                    tracked_depth = tracked_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("tracked-change range stack underflow".to_string())
                    })?;
                } else {
                    annotation_depth = annotation_depth.saturating_sub(1);
                    if let Some(depth) = paragraph_depth.as_mut() {
                        *depth = depth.checked_sub(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "change-range paragraph stack underflow".to_string(),
                            )
                        })?;
                        if *depth == 0 {
                            paragraph_depth = None;
                        }
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0
        || tracked_depth != 0
        || paragraph_depth.is_some()
        || annotation_depth != 0
    {
        return Err(Error::InvalidFormat(
            "incomplete change-range XML structure".to_string(),
        ));
    }
    if let Some(id) = ranges.pending.keys().next() {
        return Err(Error::InvalidFormat(format!(
            "unclosed text:change-start for '{id}'"
        )));
    }
    for change in changes {
        if change.change_type != ChangeType::Deletion
            && let Some(completed) = ranges.completed.remove(&change.id)
        {
            change.content.clear();
            for (index, range) in completed.iter().enumerate() {
                if index > 0 {
                    append_checked(&mut change.content, "\n")?;
                }
                append_checked(&mut change.content, range)?;
            }
        }
    }
    Ok(())
}

fn enter_change_paragraph(
    text_element: bool,
    element: &BytesStart<'_>,
    paragraph_depth: &mut Option<usize>,
    pending: &mut HashMap<String, PendingChangeRange>,
) -> Result<()> {
    if let Some(depth) = paragraph_depth.as_mut() {
        *depth = checked_semantic_depth(*depth, "change-range paragraph")?;
    } else if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
        *paragraph_depth = Some(1);
        for range in pending.values_mut() {
            if range.seen_paragraph {
                append_checked(&mut range.text, "\n")?;
            }
            range.seen_paragraph = true;
        }
    }
    Ok(())
}

fn process_change_marker(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_element: bool,
    in_paragraph: bool,
    change_types: &HashMap<String, ChangeType>,
    ranges: &mut ChangeRangeState,
) -> Result<()> {
    if !text_element {
        return Ok(());
    }
    let marker = element.local_name();
    if !matches!(marker.as_ref(), b"change" | b"change-start" | b"change-end") {
        return Ok(());
    }
    let id = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"change-id",
        "change marker",
    )?
    .ok_or_else(|| {
        Error::InvalidFormat(format!(
            "text:{} requires text:change-id",
            String::from_utf8_lossy(marker.as_ref())
        ))
    })?;
    if !change_types.contains_key(&id) {
        return Err(Error::InvalidFormat(format!(
            "change marker references unknown ID '{id}'"
        )));
    }
    match marker.as_ref() {
        b"change-start" => {
            if ranges.pending.len() >= MAX_SEMANTIC_ITEMS {
                return Err(Error::InvalidFormat(format!(
                    "document exceeds {MAX_SEMANTIC_ITEMS} open change ranges"
                )));
            }
            if ranges
                .pending
                .insert(
                    id.clone(),
                    PendingChangeRange {
                        text: String::new(),
                        seen_paragraph: in_paragraph,
                    },
                )
                .is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "duplicate open change range '{id}'"
                )));
            }
        },
        b"change-end" => {
            let range = ranges.pending.remove(&id).ok_or_else(|| {
                Error::InvalidFormat(format!("text:change-end has no open range for '{id}'"))
            })?;
            ranges.completed_count = ranges.completed_count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("completed change-range count overflow".to_string())
            })?;
            if ranges.completed_count > MAX_SEMANTIC_ITEMS {
                return Err(Error::InvalidFormat(format!(
                    "document exceeds {MAX_SEMANTIC_ITEMS} completed change ranges"
                )));
            }
            ranges.completed.entry(id).or_default().push(range.text);
        },
        b"change" => {},
        _ => unreachable!(),
    }
    Ok(())
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
    source_depth: Option<usize>,
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
        let office_element = is_bound(&namespace, OFFICE_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_semantic_depth(document_depth, "section")?;
                for section in &mut active {
                    if section.source_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "section source declarations must be empty".to_string(),
                        ));
                    }
                    section.depth = checked_semantic_depth(section.depth, "section")?;
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
                        source_depth: None,
                    });
                    next_order += 1;
                } else if let Some(section) = active.last_mut()
                    && section.depth == 2
                    && ((text_element && element.local_name().as_ref() == b"section-source")
                        || (office_element && element.local_name().as_ref() == b"dde-source"))
                {
                    apply_section_source(
                        &reader,
                        element,
                        text_element,
                        office_element,
                        &mut section.section,
                    )?;
                    section.source_depth = Some(section.depth);
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
                } else if let Some(section) = active.last_mut()
                    && section.depth == 1
                    && ((text_element && element.local_name().as_ref() == b"section-source")
                        || (office_element && element.local_name().as_ref() == b"dde-source"))
                {
                    apply_section_source(
                        &reader,
                        element,
                        text_element,
                        office_element,
                        &mut section.section,
                    )?;
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
                if active.iter().any(|section| section.source_depth.is_some()) {
                    return Err(Error::InvalidFormat(
                        "section source declarations must be empty".to_string(),
                    ));
                }
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
                if active.iter().any(|section| section.source_depth.is_some()) {
                    return Err(Error::InvalidFormat(
                        "section source declarations must be empty".to_string(),
                    ));
                }
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
                if active.iter().any(|section| section.source_depth.is_some()) {
                    return Err(Error::InvalidFormat(
                        "section source declarations must be empty".to_string(),
                    ));
                }
                let value = decode_reference(reference, "section")?;
                for section in &mut active {
                    if section.source_depth == Some(section.depth) {
                        section.source_depth = None;
                    }
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
    let xml_id = namespaced_attribute(reader, element, XML_NAMESPACE, b"id", "section")?;
    let protection_key = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"protection-key",
        "section",
    )?;
    let protection_key_digest_algorithm = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"protection-key-digest-algorithm",
        "section",
    )?;
    let display_value =
        namespaced_attribute(reader, element, TEXT_NAMESPACE, b"display", "section")?;
    let condition = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"condition", "section")?;
    let display = match display_value.as_deref() {
        None | Some("true") => SectionDisplay::Visible,
        Some("none") => SectionDisplay::Hidden,
        Some("condition") if condition.is_some() => SectionDisplay::Condition,
        Some("condition") => {
            return Err(Error::InvalidFormat(
                "text:display='condition' requires text:condition".to_string(),
            ));
        },
        Some(value) => {
            return Err(Error::InvalidFormat(format!(
                "unsupported text:display value '{value}'"
            )));
        },
    };
    if condition.is_some() && display != SectionDisplay::Condition {
        return Err(Error::InvalidFormat(
            "text:condition requires text:display='condition'".to_string(),
        ));
    }
    Ok(Section {
        name,
        style,
        protected,
        xml_id,
        protection_key,
        protection_key_digest_algorithm,
        display,
        condition,
        source: None,
        dde_source: None,
        content: String::new(),
    })
}

fn apply_section_source(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_element: bool,
    office_element: bool,
    section: &mut Section,
) -> Result<()> {
    if section.source.is_some() || section.dde_source.is_some() {
        return Err(Error::InvalidFormat(
            "section may have only one source declaration".to_string(),
        ));
    }
    if text_element && element.local_name().as_ref() == b"section-source" {
        let href =
            namespaced_attribute(reader, element, XLINK_NAMESPACE, b"href", "section source")?;
        let link_type =
            namespaced_attribute(reader, element, XLINK_NAMESPACE, b"type", "section source")?;
        if link_type.as_deref().is_some_and(|value| value != "simple") {
            return Err(Error::InvalidFormat(
                "section source xlink:type must be 'simple'".to_string(),
            ));
        }
        let show =
            namespaced_attribute(reader, element, XLINK_NAMESPACE, b"show", "section source")?;
        if show.as_deref().is_some_and(|value| value != "embed") {
            return Err(Error::InvalidFormat(
                "section source xlink:show must be 'embed'".to_string(),
            ));
        }
        if href.is_some() != link_type.is_some() || href.is_none() && show.is_some() {
            return Err(Error::InvalidFormat(
                "section source xlink:href and xlink:type must appear together".to_string(),
            ));
        }
        section.source = Some(SectionSource {
            href,
            section_name: namespaced_attribute(
                reader,
                element,
                TEXT_NAMESPACE,
                b"section-name",
                "section source",
            )?,
            filter_name: namespaced_attribute(
                reader,
                element,
                TEXT_NAMESPACE,
                b"filter-name",
                "section source",
            )?,
        });
    } else if office_element && element.local_name().as_ref() == b"dde-source" {
        let conversion_mode = namespaced_attribute(
            reader,
            element,
            OFFICE_NAMESPACE,
            b"conversion-mode",
            "section DDE source",
        )?;
        if conversion_mode.as_deref().is_some_and(|value| {
            !matches!(
                value,
                "into-default-style-data-style" | "into-english-number" | "keep-text"
            )
        }) {
            return Err(Error::InvalidFormat(
                "unsupported office:conversion-mode on section DDE source".to_string(),
            ));
        }
        section.dde_source = Some(SectionDdeSource {
            name: namespaced_attribute(
                reader,
                element,
                OFFICE_NAMESPACE,
                b"name",
                "section DDE source",
            )?,
            conversion_mode,
            automatic_update: namespaced_attribute(
                reader,
                element,
                OFFICE_NAMESPACE,
                b"automatic-update",
                "section DDE source",
            )?
            .map(|value| parse_boolean(&value, "office:automatic-update"))
            .transpose()?,
        });
    }
    Ok(())
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

        assert_eq!(changes[0].id, "change1");
        assert_eq!(changes[0].change_type, ChangeType::Insertion);
        assert_eq!(changes[0].author.as_deref(), Some("John Doe"));
        assert_eq!(changes[0].date.as_deref(), Some("2024-03-15T10:30:00"));
        assert!(changes[0].content.is_empty());

        assert_eq!(changes[1].id, "change2");
        assert_eq!(changes[1].change_type, ChangeType::Deletion);
        assert_eq!(changes[1].author.as_deref(), Some("Jane Smith"));
        assert_eq!(changes[1].date.as_deref(), Some("2024-03-15T11:00:00"));
        assert!(changes[1].content.is_empty());

        assert_eq!(changes[2].id, "change3");
        assert_eq!(changes[2].change_type, ChangeType::FormatChange);
        assert_eq!(changes[2].author.as_deref(), Some("Bob Wilson"));
        assert_eq!(changes[2].date.as_deref(), Some("2024-03-15T12:00:00"));
        assert!(changes[2].content.is_empty());
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
    fn parses_tracked_change_metadata_deletions_and_referenced_ranges() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:body><o:text><t:tracked-changes><t:changed-region t:id="i1"><t:insertion><o:change-info><d:creator>A &amp; B</d:creator><d:date>2026-07-16T10:00:00</d:date><t:p>review note</t:p></o:change-info></t:insertion></t:changed-region><t:changed-region xml:id="d1"><t:deletion><o:change-info><d:creator>Deleter</d:creator><d:date>2026-07-16</d:date><t:p>not deleted text</t:p></o:change-info><t:p>Gone &amp;<t:s t:c="2"/><t:span><![CDATA[X]]></t:span></t:p><t:p>Second<t:tab/></t:p></t:deletion></t:changed-region><t:changed-region t:id="f1"><t:format-change><o:change-info><d:creator>Stylist</d:creator><d:date>2026-07-15</d:date></o:change-info></t:format-change></t:changed-region></t:tracked-changes><t:p>pre<t:change-start t:change-id="i1"/>In&amp;<o:annotation o:name="note"><t:p>hidden comment</t:p></o:annotation><t:span>sert</t:span><t:s t:c="2"/><![CDATA[!]]><t:change-end t:change-id="i1"/>post<t:change t:change-id="d1"/></t:p><t:p><t:change-start t:change-id="i1"/>Again<t:change-end t:change-id="i1"/> and <t:change-start t:change-id="f1"/>Bold<t:change-end t:change-id="f1"/></t:p></o:text></o:body></o:document-content>"#;
        let changes = OdtParser::parse_track_changes(xml).unwrap();
        assert_eq!(changes.len(), 3);
        assert_eq!(changes[0].id, "i1");
        assert_eq!(changes[0].author.as_deref(), Some("A & B"));
        assert_eq!(changes[0].date.as_deref(), Some("2026-07-16T10:00:00"));
        assert_eq!(changes[0].content, "In&sert  !\nAgain");
        assert_eq!(changes[1].id, "d1");
        assert_eq!(changes[1].change_type, ChangeType::Deletion);
        assert_eq!(changes[1].content, "Gone &  X\nSecond\t");
        assert_eq!(changes[2].id, "f1");
        assert_eq!(changes[2].change_type, ChangeType::FormatChange);
        assert_eq!(changes[2].content, "Bold");
    }

    #[test]
    fn tracked_changes_reject_ambiguous_declarations_and_ranges() {
        let prelude = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:u="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:body><o:text>"#;
        let info = r#"<o:change-info><d:creator>A</d:creator><d:date>D</d:date></o:change-info>"#;
        let suffix = "</o:text></o:body></o:document-content>";

        let missing_id = format!(
            "{prelude}<t:tracked-changes><t:changed-region><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes>{suffix}"
        );
        assert!(OdtParser::parse_track_changes(&missing_id).is_err());

        let duplicate_id = format!(
            "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region><t:changed-region t:id=\"x\"><t:deletion>{info}</t:deletion></t:changed-region></t:tracked-changes>{suffix}"
        );
        assert!(OdtParser::parse_track_changes(&duplicate_id).is_err());

        let multiple_kinds = format!(
            "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion><t:deletion>{info}</t:deletion></t:changed-region></t:tracked-changes>{suffix}"
        );
        assert!(OdtParser::parse_track_changes(&multiple_kinds).is_err());

        let missing_kind = format!(
            "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"/></t:tracked-changes>{suffix}"
        );
        assert!(OdtParser::parse_track_changes(&missing_kind).is_err());

        let unknown_marker = format!(
            "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes><t:p><t:change t:change-id=\"unknown\"/></t:p>{suffix}"
        );
        assert!(OdtParser::parse_track_changes(&unknown_marker).is_err());

        let unmatched_end = format!(
            "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes><t:p><t:change-end t:change-id=\"x\"/></t:p>{suffix}"
        );
        assert!(OdtParser::parse_track_changes(&unmatched_end).is_err());

        let unmatched_start = format!(
            "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes><t:p><t:change-start t:change-id=\"x\"/>open</t:p>{suffix}"
        );
        assert!(OdtParser::parse_track_changes(&unmatched_start).is_err());

        let duplicate_attribute = format!(
            "{prelude}<t:tracked-changes><t:changed-region t:id=\"x\"><t:insertion>{info}</t:insertion></t:changed-region></t:tracked-changes><t:p><t:change t:change-id=\"x\" u:change-id=\"x\"/></t:p>{suffix}"
        );
        assert!(OdtParser::parse_track_changes(&duplicate_attribute).is_err());
        assert!(OdtParser::parse_track_changes("<t:tracked-changes>").is_err());
    }

    #[test]
    fn tracked_changes_enforce_nesting_bound() {
        let mut xml = String::from(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:body><o:text><t:tracked-changes><t:changed-region t:id="d"><t:deletion><o:change-info><d:creator>A</d:creator><d:date>D</d:date></o:change-info><t:p>"#,
        );
        for _ in 0..MAX_SEMANTIC_DEPTH {
            xml.push_str("<t:span>");
        }
        for _ in 0..MAX_SEMANTIC_DEPTH {
            xml.push_str("</t:span>");
        }
        xml.push_str(
            "</t:p></t:deletion></t:changed-region></t:tracked-changes></o:text></o:body></o:document-content>",
        );
        assert!(OdtParser::parse_track_changes(&xml).is_err());
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
        let xml = r#"<x:document-content xmlns:x="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:l="http://www.w3.org/1999/xlink"><x:body><x:text><t:section t:name="Outer &amp; Main" t:style-name="S1" t:protected="1" xml:id="outer" t:protection-key="YWJj" t:protection-key-digest-algorithm="urn:sha256" t:display="condition" t:condition="ooow:visible()"><t:section-source l:type="simple" l:href="https://example.invalid/doc.odt" l:show="embed" t:section-name="Remote" t:filter-name="writer8"/><t:p>One &amp;<t:s t:c="2"/></t:p><t:section t:name="Inner"><t:p>Inner <![CDATA[X]]></t:p></t:section><t:p>Last</t:p></t:section><t:section t:name="Empty"><x:dde-source x:name="Feed" x:conversion-mode="keep-text" x:automatic-update="false"/></t:section></x:text></x:body></x:document-content>"#;
        let sections = OdtParser::parse_sections(xml).unwrap();
        assert_eq!(sections.len(), 3);
        assert_eq!(sections[0].name, "Outer & Main");
        assert_eq!(sections[0].style.as_deref(), Some("S1"));
        assert!(sections[0].protected);
        assert_eq!(sections[0].xml_id.as_deref(), Some("outer"));
        assert_eq!(sections[0].protection_key.as_deref(), Some("YWJj"));
        assert_eq!(
            sections[0].protection_key_digest_algorithm.as_deref(),
            Some("urn:sha256")
        );
        assert_eq!(sections[0].display, SectionDisplay::Condition);
        assert_eq!(sections[0].condition.as_deref(), Some("ooow:visible()"));
        let source = sections[0].source.as_ref().unwrap();
        assert_eq!(
            source.href.as_deref(),
            Some("https://example.invalid/doc.odt")
        );
        assert_eq!(source.section_name.as_deref(), Some("Remote"));
        assert_eq!(source.filter_name.as_deref(), Some("writer8"));
        assert_eq!(sections[0].content, "One &  \nInner X\nLast");
        assert_eq!(sections[1].name, "Inner");
        assert_eq!(sections[1].content, "Inner X");
        assert_eq!(sections[2].name, "Empty");
        assert!(sections[2].content.is_empty());
        let dde = sections[2].dde_source.as_ref().unwrap();
        assert_eq!(dde.name.as_deref(), Some("Feed"));
        assert_eq!(dde.conversion_mode.as_deref(), Some("keep-text"));
        assert_eq!(dde.automatic_update, Some(false));
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
        let missing_condition =
            format!(r#"<t:section xmlns:t="{namespace}" t:name="A" t:display="condition"/>"#);
        assert!(OdtParser::parse_sections(&missing_condition).is_err());
        let stray_condition =
            format!(r#"<t:section xmlns:t="{namespace}" t:name="A" t:condition="x"/>"#);
        assert!(OdtParser::parse_sections(&stray_condition).is_err());
        let duplicate_source = format!(
            r#"<t:section xmlns:t="{namespace}" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" t:name="A"><t:section-source/><o:dde-source/></t:section>"#
        );
        assert!(OdtParser::parse_sections(&duplicate_source).is_err());
        let nonempty_source = format!(
            r#"<t:section xmlns:t="{namespace}" t:name="A"><t:section-source>bad</t:section-source></t:section>"#
        );
        assert!(OdtParser::parse_sections(&nonempty_source).is_err());
        let incomplete_link = format!(
            r#"<t:section xmlns:t="{namespace}" xmlns:l="http://www.w3.org/1999/xlink" t:name="A"><t:section-source l:href="x"/></t:section>"#
        );
        assert!(OdtParser::parse_sections(&incomplete_link).is_err());
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
            xml_id: None,
            protection_key: None,
            protection_key_digest_algorithm: None,
            display: SectionDisplay::Visible,
            condition: None,
            source: None,
            dde_source: None,
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
            xml_id: None,
            protection_key: None,
            protection_key_digest_algorithm: None,
            display: SectionDisplay::Visible,
            condition: None,
            source: None,
            dde_source: None,
            content: "Text".to_string(),
        };
        let cloned = section.clone();
        assert_eq!(section.name, cloned.name);
        assert_eq!(section.protected, cloned.protected);
    }
}
