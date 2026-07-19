//! Semantic access to point and range cross-reference targets.

use super::index::expanded_attributes;
use crate::elements::xml::{
    TEXT_NAMESPACE, append_checked, append_text_control, decode_reference, is_bound,
    namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

mod writing;
pub use writing::{
    ReferenceMarkFragments, insert_reference_mark_xml, remove_reference_mark_xml,
    replace_reference_mark_xml,
};

const MAX_DEPTH: usize = 4_096;
const MAX_MARKS: usize = 1_000_000;

/// A point or range target for `text:reference-ref` fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceMark {
    name: String,
    start: Option<(usize, usize)>,
    end: Option<(usize, usize)>,
    text: String,
    range: bool,
}

impl ReferenceMark {
    /// Create a point reference target.
    pub fn point(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            start: None,
            end: None,
            text: String::new(),
            range: false,
        }
    }

    /// Create a range reference target.
    pub fn range(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            start: None,
            end: None,
            text: String::new(),
            range: true,
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    /// Zero-based paragraph/heading index and character offset.
    pub fn start(&self) -> Option<(usize, usize)> {
        self.start
    }

    /// Zero-based paragraph/heading index and character offset.
    pub fn end(&self) -> Option<(usize, usize)> {
        self.end
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn is_range(&self) -> bool {
        self.range
    }
}

struct PendingReference {
    mark: ReferenceMark,
    order: usize,
    seen_paragraph: bool,
}

pub(crate) fn parse_reference_marks(xml: &str) -> Result<Vec<ReferenceMark>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut paragraph_index = 0usize;
    let mut paragraph: Option<(usize, usize, usize)> = None;
    let mut marker_depth = None;
    let mut pending = HashMap::<String, PendingReference>::new();
    let mut identities = HashSet::<String>::new();
    let mut marks = Vec::new();
    let mut next_order = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid reference-mark XML: {error}"))
            })?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_depth(document_depth)?;
                if marker_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "reference-mark elements must be empty".to_string(),
                    ));
                }
                if let Some(active) = paragraph.as_mut() {
                    let location = Some((active.0, active.1));
                    process_marker(
                        &reader,
                        element,
                        text_element,
                        location,
                        &mut pending,
                        &mut identities,
                        &mut marks,
                        &mut next_order,
                    )?;
                    if text_element {
                        append_control(&reader, element, active, &mut pending)?;
                    }
                    active.2 = checked_depth(active.2)?;
                } else if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                    paragraph = Some((paragraph_index, 0, 1));
                    paragraph_index = paragraph_index.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("reference-mark paragraph count overflow".to_string())
                    })?;
                    for reference in pending.values_mut() {
                        if reference.seen_paragraph {
                            append_checked(&mut reference.mark.text, "\n")?;
                        }
                        reference.seen_paragraph = true;
                    }
                } else {
                    process_marker(
                        &reader,
                        element,
                        text_element,
                        None,
                        &mut pending,
                        &mut identities,
                        &mut marks,
                        &mut next_order,
                    )?;
                }
                if text_element && is_marker(element) {
                    marker_depth = Some(document_depth);
                }
            },
            Event::Empty(ref element) => {
                let location = paragraph.map(|(index, offset, _)| (index, offset));
                process_marker(
                    &reader,
                    element,
                    text_element,
                    location,
                    &mut pending,
                    &mut identities,
                    &mut marks,
                    &mut next_order,
                )?;
                if text_element && let Some(active) = paragraph.as_mut() {
                    append_control(&reader, element, active, &mut pending)?;
                }
            },
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) if marker_depth.is_some() => {
                return Err(Error::InvalidFormat(
                    "reference-mark elements must be empty".to_string(),
                ));
            },
            Event::Text(ref value) if paragraph.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid reference-mark text: {error}"))
                    })?;
                append_text(
                    paragraph.as_mut().expect("checked paragraph"),
                    &mut pending,
                    &value,
                )?;
            },
            Event::CData(ref value) if paragraph.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid reference-mark CDATA: {error}"))
                    })?;
                append_text(
                    paragraph.as_mut().expect("checked paragraph"),
                    &mut pending,
                    &value,
                )?;
            },
            Event::GeneralRef(ref reference) if paragraph.is_some() => {
                append_text(
                    paragraph.as_mut().expect("checked paragraph"),
                    &mut pending,
                    &decode_reference(reference, "reference mark")?,
                )?;
            },
            Event::End(_) => {
                if marker_depth == Some(document_depth) {
                    marker_depth = None;
                }
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("reference-mark XML stack underflow".to_string())
                })?;
                if let Some((_, _, depth)) = paragraph.as_mut() {
                    *depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("reference-mark paragraph stack underflow".to_string())
                    })?;
                    if *depth == 0 {
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
            "incomplete reference-mark XML structure".to_string(),
        ));
    }
    if let Some(name) = pending.keys().next() {
        return Err(Error::InvalidFormat(format!(
            "unclosed reference-mark range '{name}'"
        )));
    }
    marks.sort_by_key(|(order, _)| *order);
    Ok(marks.into_iter().map(|(_, mark)| mark).collect())
}

fn process_marker(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    text_element: bool,
    location: Option<(usize, usize)>,
    pending: &mut HashMap<String, PendingReference>,
    identities: &mut HashSet<String>,
    marks: &mut Vec<(usize, ReferenceMark)>,
    next_order: &mut usize,
) -> Result<()> {
    if !text_element || !is_marker(element) {
        return Ok(());
    }
    validate_marker_attributes(reader, element)?;
    let name = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"name", "reference mark")?
        .ok_or_else(|| Error::InvalidFormat("reference mark requires text:name".to_string()))?;
    validate_reference_name(&name)?;
    match element.local_name().as_ref() {
        b"reference-mark" => {
            ensure_capacity(*next_order)?;
            if !identities.insert(name.clone()) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate reference-mark identity '{name}'"
                )));
            }
            marks.push((
                *next_order,
                ReferenceMark {
                    name,
                    start: location,
                    end: location,
                    text: String::new(),
                    range: false,
                },
            ));
            *next_order += 1;
        },
        b"reference-mark-start" => {
            ensure_capacity(*next_order)?;
            if !identities.insert(name.clone()) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate reference-mark identity '{name}'"
                )));
            }
            if pending
                .insert(
                    name.clone(),
                    PendingReference {
                        mark: ReferenceMark {
                            name: name.clone(),
                            start: location,
                            end: None,
                            text: String::new(),
                            range: true,
                        },
                        order: *next_order,
                        seen_paragraph: paragraph_seen(location),
                    },
                )
                .is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "duplicate open reference-mark range '{name}'"
                )));
            }
            *next_order += 1;
        },
        b"reference-mark-end" => {
            let mut reference = pending.remove(&name).ok_or_else(|| {
                Error::InvalidFormat(format!("reference-mark-end has no open range '{name}'"))
            })?;
            reference.mark.end = location;
            marks.push((reference.order, reference.mark));
        },
        _ => unreachable!(),
    }
    Ok(())
}

fn validate_marker_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    let attributes = expanded_attributes(reader, element, "reference mark")?;
    if attributes.len() != 1
        || attributes[0].namespace_uri.as_deref()
            != Some(std::str::from_utf8(TEXT_NAMESPACE).expect("ODF text namespace is UTF-8"))
        || attributes[0].local_name != "name"
    {
        return Err(Error::InvalidFormat(
            "reference marks allow only text:name".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn validate_reference_name(name: &str) -> Result<()> {
    const MAX_NAME_BYTES: usize = 65_536;
    if name.len() > MAX_NAME_BYTES {
        return Err(Error::InvalidFormat(format!(
            "reference-mark name exceeds {MAX_NAME_BYTES} bytes"
        )));
    }
    if name.chars().any(|character| {
        !matches!(
            character,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}'
                | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(Error::InvalidFormat(
            "reference-mark name contains a forbidden XML character".to_string(),
        ));
    }
    Ok(())
}

fn paragraph_seen(location: Option<(usize, usize)>) -> bool {
    location.is_some()
}

fn is_marker(element: &BytesStart<'_>) -> bool {
    matches!(
        element.local_name().as_ref(),
        b"reference-mark" | b"reference-mark-start" | b"reference-mark-end"
    )
}

fn append_text(
    paragraph: &mut (usize, usize, usize),
    pending: &mut HashMap<String, PendingReference>,
    value: &str,
) -> Result<()> {
    for reference in pending.values_mut() {
        append_checked(&mut reference.mark.text, value)?;
    }
    paragraph.1 = paragraph
        .1
        .checked_add(value.chars().count())
        .ok_or_else(|| Error::InvalidFormat("reference-mark offset overflow".to_string()))?;
    Ok(())
}

fn append_control(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    paragraph: &mut (usize, usize, usize),
    pending: &mut HashMap<String, PendingReference>,
) -> Result<()> {
    let mut value = String::new();
    append_text_control(reader, element, &mut value)?;
    append_text(paragraph, pending, &value)
}

fn ensure_capacity(count: usize) -> Result<()> {
    if count >= MAX_MARKS {
        return Err(Error::InvalidFormat(format!(
            "document exceeds {MAX_MARKS} reference marks"
        )));
    }
    Ok(())
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("reference-mark depth overflow".to_string()))?;
    if depth > MAX_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "reference-mark nesting exceeds {MAX_DEPTH} levels"
        )));
    }
    Ok(depth)
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    #[test]
    fn parses_point_and_range_reference_marks_with_exact_positions_and_text() {
        let xml = format!(
            r#"<x:text xmlns:x="{TEXT}"><x:p>ab<x:reference-mark x:name="point"/>c<x:reference-mark-start x:name="range"/>D&amp;<x:span>E</x:span><x:s x:c="2"/></x:p><x:p>F<![CDATA[!]]><x:reference-mark-end x:name="range"/>z</x:p></x:text>"#
        );
        let marks = parse_reference_marks(&xml).unwrap();
        assert_eq!(marks.len(), 2);
        assert_eq!(marks[0].name(), "point");
        assert!(!marks[0].is_range());
        assert_eq!(marks[0].start(), Some((0, 2)));
        assert_eq!(marks[0].end(), Some((0, 2)));
        assert!(marks[0].text().is_empty());
        assert_eq!(marks[1].name(), "range");
        assert!(marks[1].is_range());
        assert_eq!(marks[1].start(), Some((0, 3)));
        assert_eq!(marks[1].end(), Some((1, 2)));
        assert_eq!(marks[1].text(), "D&E  \nF!");
    }

    #[test]
    fn reference_marks_reject_missing_duplicate_unmatched_and_nonempty_markers() {
        let missing = format!(r#"<x:reference-mark xmlns:x="{TEXT}"/>"#);
        assert!(parse_reference_marks(&missing).is_err());
        let duplicate = format!(
            r#"<x:p xmlns:x="{TEXT}"><x:reference-mark-start x:name="a"/><x:reference-mark-start x:name="a"/></x:p>"#
        );
        assert!(parse_reference_marks(&duplicate).is_err());
        let unmatched = format!(r#"<x:reference-mark-end xmlns:x="{TEXT}" x:name="a"/>"#);
        assert!(parse_reference_marks(&unmatched).is_err());
        let unclosed = format!(r#"<x:reference-mark-start xmlns:x="{TEXT}" x:name="a"/>"#);
        assert!(parse_reference_marks(&unclosed).is_err());
        let nonempty =
            format!(r#"<x:reference-mark xmlns:x="{TEXT}" x:name="a">bad</x:reference-mark>"#);
        assert!(parse_reference_marks(&nonempty).is_err());
        let aliases = format!(
            r#"<x:reference-mark xmlns:x="{TEXT}" xmlns:y="{TEXT}" x:name="a" y:name="b"/>"#
        );
        assert!(parse_reference_marks(&aliases).is_err());
        assert!(parse_reference_marks("<x:reference-mark>").is_err());
    }

    #[test]
    fn reference_marks_enforce_nesting_bound() {
        let mut xml = format!(r#"<x:p xmlns:x="{TEXT}">"#);
        for _ in 0..MAX_DEPTH {
            xml.push_str("<x:span>");
        }
        for _ in 0..MAX_DEPTH {
            xml.push_str("</x:span>");
        }
        xml.push_str("</x:p>");
        assert!(parse_reference_marks(&xml).is_err());
    }
}
