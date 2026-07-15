//! Semantic parsing for OpenDocument index source marks.

use super::index::{TextIndexAttribute, expanded_attributes};
use crate::elements::xml::{
    TEXT_NAMESPACE, append_checked, append_text_control, decode_reference, is_bound,
    namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::collections::HashMap;

const MAX_MARK_DEPTH: usize = 4_096;
const MAX_MARKS: usize = 1_000_000;

/// An index source-mark family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextIndexMarkKind {
    TableOfContents,
    User,
    Alphabetical,
    Bibliography,
}

/// A point or resolved range mark that contributes an entry to a generated index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexMark {
    kind: TextIndexMarkKind,
    id: Option<String>,
    value: String,
    range: bool,
    attributes: Vec<TextIndexAttribute>,
}

impl TextIndexMark {
    pub fn kind(&self) -> TextIndexMarkKind {
        self.kind
    }

    pub fn id(&self) -> Option<&str> {
        self.id.as_deref()
    }

    /// Point marks return their stored string; range marks return their referenced visible text.
    pub fn value(&self) -> &str {
        &self.value
    }

    pub fn is_range(&self) -> bool {
        self.range
    }

    pub fn attributes(&self) -> &[TextIndexAttribute] {
        &self.attributes
    }

    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name() == local_name
            })
            .map(TextIndexAttribute::value)
    }
}

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
                    let active = bibliography.take().expect("checked bibliography mark");
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
    )? {
        if level
            .parse::<usize>()
            .ok()
            .filter(|level| *level > 0)
            .is_none()
        {
            return Err(Error::InvalidFormat(
                "text:outline-level must be a positive integer".to_string(),
            ));
        }
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

fn is_bibliography_type(value: &str) -> bool {
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

fn start_kind(local_name: &[u8]) -> Option<TextIndexMarkKind> {
    match local_name {
        b"toc-mark-start" => Some(TextIndexMarkKind::TableOfContents),
        b"user-index-mark-start" => Some(TextIndexMarkKind::User),
        b"alphabetical-index-mark-start" => Some(TextIndexMarkKind::Alphabetical),
        _ => None,
    }
}

fn end_kind(local_name: &[u8]) -> Option<TextIndexMarkKind> {
    match local_name {
        b"toc-mark-end" => Some(TextIndexMarkKind::TableOfContents),
        b"user-index-mark-end" => Some(TextIndexMarkKind::User),
        b"alphabetical-index-mark-end" => Some(TextIndexMarkKind::Alphabetical),
        _ => None,
    }
}

fn point_kind(local_name: &[u8]) -> Option<TextIndexMarkKind> {
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

#[cfg(test)]
mod tests {
    use super::*;

    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    #[test]
    fn parses_point_range_and_bibliography_marks_in_document_order() {
        let xml = format!(
            r#"<x:text xmlns:x="{TEXT}" xmlns:u="urn:vendor"><x:p><x:toc-mark-start x:id="same" x:outline-level="2" u:flag="yes"/>Alpha<x:span>&amp;</x:span><x:toc-mark x:string-value="Manual &amp; Entry"/><x:alphabetical-index-mark-start x:id="same" x:key1="K" x:main-entry="1"/>Beta<![CDATA[!]]><x:alphabetical-index-mark-end x:id="same"/></x:p><x:p>Next<x:toc-mark-end x:id="same"/><x:bibliography-mark x:bibliography-type="book" x:author="A &amp; B" x:title="T">[1] &amp; more</x:bibliography-mark><x:user-index-mark x:index-name="Custom" x:string-value="User" x:outline-level="3"/></x:p></x:text>"#
        );
        let marks = parse_text_index_marks(&xml).unwrap();
        assert_eq!(marks.len(), 5);
        assert_eq!(marks[0].kind(), TextIndexMarkKind::TableOfContents);
        assert!(marks[0].is_range());
        assert_eq!(marks[0].id(), Some("same"));
        assert_eq!(marks[0].value(), "Alpha&Beta!\nNext");
        assert_eq!(marks[0].attribute(Some(TEXT), "outline-level"), Some("2"));
        assert_eq!(marks[0].attribute(Some("urn:vendor"), "flag"), Some("yes"));

        assert_eq!(marks[1].kind(), TextIndexMarkKind::TableOfContents);
        assert!(!marks[1].is_range());
        assert_eq!(marks[1].value(), "Manual & Entry");

        assert_eq!(marks[2].kind(), TextIndexMarkKind::Alphabetical);
        assert_eq!(marks[2].id(), Some("same"));
        assert_eq!(marks[2].value(), "Beta!");
        assert_eq!(marks[2].attribute(Some(TEXT), "key1"), Some("K"));

        assert_eq!(marks[3].kind(), TextIndexMarkKind::Bibliography);
        assert_eq!(marks[3].value(), "[1] & more");
        assert_eq!(marks[3].attribute(Some(TEXT), "author"), Some("A & B"));
        assert_eq!(
            marks[3].attribute(Some(TEXT), "bibliography-type"),
            Some("book")
        );

        assert_eq!(marks[4].kind(), TextIndexMarkKind::User);
        assert_eq!(marks[4].value(), "User");
        assert_eq!(marks[4].attribute(Some(TEXT), "index-name"), Some("Custom"));
    }

    #[test]
    fn index_marks_reject_missing_ambiguous_and_unmatched_metadata() {
        let missing = format!(r#"<x:toc-mark xmlns:x="{TEXT}"/>"#);
        assert!(parse_text_index_marks(&missing).is_err());
        let unmatched = format!(r#"<x:toc-mark-end xmlns:x="{TEXT}" x:id="a"/>"#);
        assert!(parse_text_index_marks(&unmatched).is_err());
        let unclosed = format!(r#"<x:toc-mark-start xmlns:x="{TEXT}" x:id="a"/>"#);
        assert!(parse_text_index_marks(&unclosed).is_err());
        let duplicate = format!(
            r#"<x:p xmlns:x="{TEXT}"><x:toc-mark-start x:id="a"/><x:toc-mark-start x:id="a"/></x:p>"#
        );
        assert!(parse_text_index_marks(&duplicate).is_err());
        let aliases = format!(
            r#"<x:toc-mark xmlns:x="{TEXT}" xmlns:y="{TEXT}" x:string-value="A" y:string-value="B"/>"#
        );
        assert!(parse_text_index_marks(&aliases).is_err());
        let invalid_level = format!(
            r#"<x:user-index-mark xmlns:x="{TEXT}" x:index-name="I" x:string-value="V" x:outline-level="0"/>"#
        );
        assert!(parse_text_index_marks(&invalid_level).is_err());
        let invalid_boolean = format!(
            r#"<x:alphabetical-index-mark xmlns:x="{TEXT}" x:string-value="A" x:main-entry="yes"/>"#
        );
        assert!(parse_text_index_marks(&invalid_boolean).is_err());
        let invalid_bibliography_type = format!(
            r#"<x:bibliography-mark xmlns:x="{TEXT}" x:bibliography-type="novel">bad</x:bibliography-mark>"#
        );
        assert!(parse_text_index_marks(&invalid_bibliography_type).is_err());
        let bibliography_child = format!(
            r#"<x:bibliography-mark xmlns:x="{TEXT}" x:bibliography-type="book"><x:span>bad</x:span></x:bibliography-mark>"#
        );
        assert!(parse_text_index_marks(&bibliography_child).is_err());
        assert!(parse_text_index_marks("<x:toc-mark>").is_err());

        let empty_strings =
            format!(r#"<x:user-index-mark xmlns:x="{TEXT}" x:index-name="" x:string-value=""/>"#);
        assert_eq!(
            parse_text_index_marks(&empty_strings).unwrap()[0].value(),
            ""
        );
    }

    #[test]
    fn index_marks_enforce_nesting_bound() {
        let mut xml = format!(r#"<x:p xmlns:x="{TEXT}">"#);
        for _ in 0..MAX_MARK_DEPTH {
            xml.push_str("<x:span>");
        }
        for _ in 0..MAX_MARK_DEPTH {
            xml.push_str("</x:span>");
        }
        xml.push_str("</x:p>");
        assert!(parse_text_index_marks(&xml).is_err());
    }
}
