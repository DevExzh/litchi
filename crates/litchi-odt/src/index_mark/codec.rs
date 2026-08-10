//! Namespace-aware parsing for `OpenDocument` index source marks.

use super::model::{TextIndexMark, TextIndexMarkKind};
use super::{MAX_MARK_DEPTH, MAX_MARKS};
use crate::elements::xml::{
    TEXT_NAMESPACE, append_checked, append_text_control, decode_reference, is_bound,
    namespaced_attribute,
};
use crate::index::expanded_attributes;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::collections::HashMap;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct MarkKey {
    kind: TextIndexMarkKind,
    id: String,
}

struct PendingMark {
    mark: TextIndexMark,
    order: usize,
    seen_paragraph: bool,
}

struct ActiveBibliography {
    mark: TextIndexMark,
    order: usize,
    depth: usize,
}

pub(crate) fn parse_text_index_marks(xml: &str) -> Result<Vec<TextIndexMark>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut paragraph_depth: Option<usize> = None;
    let mut pending = HashMap::<MarkKey, PendingMark>::new();
    let mut bibliography: Option<ActiveBibliography> = None;
    let mut marks = Vec::new();
    let mut next_order = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid index-mark XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_depth(document_depth)?;
                if let Some(depth) = paragraph_depth.as_mut() {
                    *depth = checked_depth(*depth)?;
                } else if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                    paragraph_depth = Some(1);
                    start_mark_paragraph(&mut pending)?;
                }
                if let Some(active) = bibliography.as_mut() {
                    active.depth = checked_depth(active.depth)?;
                    return Err(Error::InvalidFormat(
                        "text:bibliography-mark may contain only text".to_string(),
                    ));
                }
                if text_element {
                    process_start(
                        &reader,
                        element,
                        paragraph_depth.is_some(),
                        &mut pending,
                        &mut bibliography,
                        &mut marks,
                        &mut next_order,
                    )?;
                    if paragraph_depth.is_some() {
                        for pending in pending.values_mut() {
                            append_text_control(&reader, element, &mut pending.mark.value)?;
                        }
                    }
                }
            },
            Event::Empty(ref element) if text_element => {
                if bibliography.is_some() {
                    return Err(Error::InvalidFormat(
                        "text:bibliography-mark may contain only text".to_string(),
                    ));
                }
                process_empty(
                    &reader,
                    element,
                    paragraph_depth.is_some(),
                    &mut pending,
                    &mut marks,
                    &mut next_order,
                )?;
                if paragraph_depth.is_some() {
                    for pending in pending.values_mut() {
                        append_text_control(&reader, element, &mut pending.mark.value)?;
                    }
                }
            },
            Event::Text(ref value) if bibliography.is_some() || paragraph_depth.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid index-mark text: {error}"))
                    })?;
                append_visible_text(&mut pending, bibliography.as_mut(), &value)?;
            },
            Event::CData(ref value) if bibliography.is_some() || paragraph_depth.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid index-mark CDATA: {error}"))
                    })?;
                append_visible_text(&mut pending, bibliography.as_mut(), &value)?;
            },
            Event::GeneralRef(ref reference)
                if bibliography.is_some() || paragraph_depth.is_some() =>
            {
                append_visible_text(
                    &mut pending,
                    bibliography.as_mut(),
                    &decode_reference(reference, "index mark")?,
                )?;
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("index-mark XML stack underflow".to_string())
                })?;
                if bibliography
                    .as_ref()
                    .is_some_and(|active| active.depth == 1)
                {
                    let active = bibliography.take().ok_or_else(|| {
                        Error::InvalidFormat("missing completed bibliography mark".to_string())
                    })?;
                    marks.push((active.order, active.mark));
                } else if let Some(active) = bibliography.as_mut() {
                    active.depth = active.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("bibliography-mark stack underflow".to_string())
                    })?;
                }
                if let Some(depth) = paragraph_depth.as_mut() {
                    *depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("index-mark paragraph stack underflow".to_string())
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
    if document_depth != 0 || paragraph_depth.is_some() || bibliography.is_some() {
        return Err(Error::InvalidFormat(
            "incomplete index-mark XML structure".to_string(),
        ));
    }
    if let Some((key, _)) = pending.iter().next() {
        return Err(Error::InvalidFormat(format!(
            "unclosed {:?} index range '{}'",
            key.kind, key.id
        )));
    }
    marks.sort_by_key(|(order, _)| *order);
    Ok(marks.into_iter().map(|(_, mark)| mark).collect())
}

fn process_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    in_paragraph: bool,
    pending: &mut HashMap<MarkKey, PendingMark>,
    bibliography: &mut Option<ActiveBibliography>,
    marks: &mut Vec<(usize, TextIndexMark)>,
    next_order: &mut usize,
) -> Result<()> {
    if let Some(kind) = start_kind(element.local_name().as_ref()) {
        open_range(reader, element, kind, in_paragraph, pending, next_order)
    } else if let Some(kind) = end_kind(element.local_name().as_ref()) {
        close_range(reader, element, kind, pending, marks)
    } else if let Some(kind) = point_kind(element.local_name().as_ref()) {
        if kind == TextIndexMarkKind::Bibliography {
            ensure_mark_capacity(*next_order)?;
            *bibliography = Some(ActiveBibliography {
                mark: point_mark(reader, element, kind, None)?,
                order: *next_order,
                depth: 1,
            });
            *next_order += 1;
            Ok(())
        } else {
            ensure_mark_capacity(*next_order)?;
            marks.push((*next_order, point_mark(reader, element, kind, None)?));
            *next_order += 1;
            Ok(())
        }
    } else {
        Ok(())
    }
}

fn process_empty(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    in_paragraph: bool,
    pending: &mut HashMap<MarkKey, PendingMark>,
    marks: &mut Vec<(usize, TextIndexMark)>,
    next_order: &mut usize,
) -> Result<()> {
    if let Some(kind) = start_kind(element.local_name().as_ref()) {
        open_range(reader, element, kind, in_paragraph, pending, next_order)
    } else if let Some(kind) = end_kind(element.local_name().as_ref()) {
        close_range(reader, element, kind, pending, marks)
    } else if let Some(kind) = point_kind(element.local_name().as_ref()) {
        ensure_mark_capacity(*next_order)?;
        marks.push((*next_order, point_mark(reader, element, kind, Some(""))?));
        *next_order += 1;
        Ok(())
    } else {
        Ok(())
    }
}

fn open_range(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: TextIndexMarkKind,
    in_paragraph: bool,
    pending: &mut HashMap<MarkKey, PendingMark>,
    next_order: &mut usize,
) -> Result<()> {
    ensure_mark_capacity(*next_order)?;
    let id = required_attribute(reader, element, b"id", "index range start")?;
    if id.is_empty() {
        return Err(Error::InvalidFormat(
            "index range ID must not be empty".to_string(),
        ));
    }
    validate_mark_attributes(reader, element, kind, false)?;
    let key = MarkKey {
        kind,
        id: id.clone(),
    };
    let mark = TextIndexMark {
        kind,
        id: Some(id),
        value: String::new(),
        range: true,
        attributes: expanded_attributes(reader, element, "index mark")?,
    };
    if pending
        .insert(
            key.clone(),
            PendingMark {
                mark,
                order: *next_order,
                seen_paragraph: in_paragraph,
            },
        )
        .is_some()
    {
        return Err(Error::InvalidFormat(format!(
            "duplicate open {:?} index range '{}'",
            kind, key.id
        )));
    }
    *next_order += 1;
    Ok(())
}

fn close_range(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: TextIndexMarkKind,
    pending: &mut HashMap<MarkKey, PendingMark>,
    marks: &mut Vec<(usize, TextIndexMark)>,
) -> Result<()> {
    let id = required_attribute(reader, element, b"id", "index range end")?;
    if id.is_empty() {
        return Err(Error::InvalidFormat(
            "index range ID must not be empty".to_string(),
        ));
    }
    let key = MarkKey { kind, id };
    let pending = pending.remove(&key).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "index range end has no open {:?} range '{}'",
            kind, key.id
        ))
    })?;
    marks.push((pending.order, pending.mark));
    Ok(())
}

fn point_mark(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: TextIndexMarkKind,
    bibliography_value: Option<&str>,
) -> Result<TextIndexMark> {
    validate_mark_attributes(reader, element, kind, true)?;
    let value = if kind == TextIndexMarkKind::Bibliography {
        bibliography_value.unwrap_or_default().to_string()
    } else {
        required_attribute(reader, element, b"string-value", "index point mark")?
    };
    Ok(TextIndexMark {
        kind,
        id: None,
        value,
        range: false,
        attributes: expanded_attributes(reader, element, "index mark")?,
    })
}

fn validate_mark_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: TextIndexMarkKind,
    point: bool,
) -> Result<()> {
    if kind == TextIndexMarkKind::User {
        required_attribute(reader, element, b"index-name", "user-index mark")?;
    }
    if kind == TextIndexMarkKind::Bibliography {
        let bibliography_type =
            required_attribute(reader, element, b"bibliography-type", "bibliography mark")?;
        if !is_bibliography_type(&bibliography_type) {
            return Err(Error::InvalidFormat(format!(
                "unsupported text:bibliography-type '{bibliography_type}'"
            )));
        }
    }
    if let Some(level) = namespaced_attribute(
        reader,
        element,
        TEXT_NAMESPACE,
        b"outline-level",
        "index mark",
    )? && level
        .parse::<usize>()
        .ok()
        .as_ref()
        .is_none_or(|level| *level <= 0)
    {
        return Err(Error::InvalidFormat(
            "text:outline-level must be a positive integer".to_string(),
        ));
    }
    if kind == TextIndexMarkKind::Alphabetical
        && let Some(value) = namespaced_attribute(
            reader,
            element,
            TEXT_NAMESPACE,
            b"main-entry",
            "alphabetical-index mark",
        )?
        && !matches!(value.as_str(), "true" | "false" | "1" | "0")
    {
        return Err(Error::InvalidFormat(
            "text:main-entry must be true, false, 1, or 0".to_string(),
        ));
    }
    if point && kind != TextIndexMarkKind::Bibliography {
        required_attribute(reader, element, b"string-value", "index point mark")?;
    }
    Ok(())
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
    context: &str,
) -> Result<String> {
    let value = namespaced_attribute(reader, element, TEXT_NAMESPACE, local_name, context)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{context} requires text:{}",
                String::from_utf8_lossy(local_name)
            ))
        })?;
    Ok(value)
}

pub(crate) fn is_bibliography_type(value: &str) -> bool {
    matches!(
        value,
        "article"
            | "book"
            | "booklet"
            | "conference"
            | "custom1"
            | "custom2"
            | "custom3"
            | "custom4"
            | "custom5"
            | "email"
            | "inbook"
            | "incollection"
            | "inproceedings"
            | "journal"
            | "manual"
            | "mastersthesis"
            | "misc"
            | "phdthesis"
            | "proceedings"
            | "techreport"
            | "unpublished"
            | "www"
    )
}

fn append_visible_text(
    pending: &mut HashMap<MarkKey, PendingMark>,
    bibliography: Option<&mut ActiveBibliography>,
    value: &str,
) -> Result<()> {
    for pending in pending.values_mut() {
        append_checked(&mut pending.mark.value, value)?;
    }
    if let Some(bibliography) = bibliography {
        append_checked(&mut bibliography.mark.value, value)?;
    }
    Ok(())
}

fn start_mark_paragraph(pending: &mut HashMap<MarkKey, PendingMark>) -> Result<()> {
    for pending in pending.values_mut() {
        if pending.seen_paragraph {
            append_checked(&mut pending.mark.value, "\n")?;
        }
        pending.seen_paragraph = true;
    }
    Ok(())
}

pub(crate) fn start_kind(local_name: &[u8]) -> Option<TextIndexMarkKind> {
    match local_name {
        b"toc-mark-start" => Some(TextIndexMarkKind::TableOfContents),
        b"user-index-mark-start" => Some(TextIndexMarkKind::User),
        b"alphabetical-index-mark-start" => Some(TextIndexMarkKind::Alphabetical),
        _ => None,
    }
}

pub(crate) fn end_kind(local_name: &[u8]) -> Option<TextIndexMarkKind> {
    match local_name {
        b"toc-mark-end" => Some(TextIndexMarkKind::TableOfContents),
        b"user-index-mark-end" => Some(TextIndexMarkKind::User),
        b"alphabetical-index-mark-end" => Some(TextIndexMarkKind::Alphabetical),
        _ => None,
    }
}

pub(crate) fn point_kind(local_name: &[u8]) -> Option<TextIndexMarkKind> {
    match local_name {
        b"toc-mark" => Some(TextIndexMarkKind::TableOfContents),
        b"user-index-mark" => Some(TextIndexMarkKind::User),
        b"alphabetical-index-mark" => Some(TextIndexMarkKind::Alphabetical),
        b"bibliography-mark" => Some(TextIndexMarkKind::Bibliography),
        _ => None,
    }
}

fn ensure_mark_capacity(count: usize) -> Result<()> {
    if count >= MAX_MARKS {
        return Err(Error::InvalidFormat(format!(
            "document exceeds {MAX_MARKS} index marks"
        )));
    }
    Ok(())
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("index-mark nesting depth overflow".to_string()))?;
    if depth > MAX_MARK_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "index-mark nesting exceeds {MAX_MARK_DEPTH} levels"
        )));
    }
    Ok(depth)
}
