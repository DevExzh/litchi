//! Namespace-aware scanner and lossless XML span mechanics.

use super::{
    MAX_CONTENT_XML_BYTES, MAX_DYNAMIC_FIELDS, MAX_PARAGRAPHS, MAX_XML_DEPTH, ParagraphSite, Scan,
    Span, TEXT_NAMESPACE,
};
use crate::elements::field::{DynamicTextField, FieldParser};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub(super) fn take_database_field(fields: &mut Vec<Span>, index: usize) -> Result<Span> {
    if index >= fields.len() {
        return Err(Error::InvalidFormat(format!(
            "database field index {index} is out of bounds (field count: {})",
            fields.len()
        )));
    }
    Ok(fields.swap_remove(index))
}

pub(super) fn take_field(fields: &mut Vec<Span>, index: usize) -> Result<Span> {
    if index >= fields.len() {
        return Err(Error::InvalidFormat(format!(
            "dynamic text field index {index} is out of bounds (field count: {})",
            fields.len()
        )));
    }
    Ok(fields.swap_remove(index))
}

pub(super) fn validate_meta_field_mutation(
    output: String,
    field: &DynamicTextField,
) -> Result<String> {
    if matches!(
        field,
        DynamicTextField::MetaField { .. }
            | DynamicTextField::DropDown { .. }
            | DynamicTextField::Script { .. }
    ) {
        FieldParser::parse_dynamic_text_fields(&output)?;
    }
    Ok(output)
}

pub(super) fn replace_span(xml: &str, span: Span, replacement: &str) -> Result<String> {
    let capacity = xml
        .len()
        .checked_sub(span.end - span.start)
        .and_then(|len| len.checked_add(replacement.len()))
        .ok_or_else(|| Error::InvalidFormat("dynamic text XML size overflow".to_string()))?;
    if capacity > MAX_CONTENT_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "mutated content.xml exceeds {MAX_CONTENT_XML_BYTES} bytes"
        )));
    }
    let mut output = String::with_capacity(capacity);
    output.push_str(&xml[..span.start]);
    output.push_str(replacement);
    output.push_str(&xml[span.end..]);
    Ok(output)
}

pub(crate) fn scan(xml: &str, wanted_paragraph: Option<usize>) -> Result<Scan> {
    if xml.len() > MAX_CONTENT_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "content.xml exceeds {MAX_CONTENT_XML_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut paragraph_count = 0usize;
    let mut active_paragraph_depth = None;
    let mut active_fields: Vec<(usize, usize, bool)> = Vec::new();
    let mut active_database_fields: Vec<(usize, usize)> = Vec::new();
    let mut output = Scan::default();

    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_| {
            Error::InvalidFormat("content.xml byte offset exceeds platform limits".to_string())
        })?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid content.xml: {error}")))?;
        let text_namespace = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if value == TEXT_NAMESPACE
        );
        match event {
            Event::Start(ref element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("content.xml nesting depth overflow".to_string())
                })?;
                if depth > MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "content.xml nesting exceeds {MAX_XML_DEPTH} levels"
                    )));
                }
                let local = element.local_name();
                if text_namespace && local.as_ref() == b"p" {
                    if paragraph_count >= MAX_PARAGRAPHS {
                        return Err(Error::InvalidFormat("too many ODF paragraphs".to_string()));
                    }
                    if wanted_paragraph == Some(paragraph_count) {
                        active_paragraph_depth = Some(depth);
                    }
                    paragraph_count += 1;
                }
                if text_namespace && is_dynamic_local(local.as_ref()) {
                    if active_fields
                        .last()
                        .is_some_and(|(_, _, allows_nested)| !allows_nested)
                    {
                        return Err(Error::InvalidFormat(
                            "nested dynamic text fields are not allowed".to_string(),
                        ));
                    }
                    if output.fields.len() >= MAX_DYNAMIC_FIELDS {
                        return Err(Error::InvalidFormat(
                            "too many dynamic text fields".to_string(),
                        ));
                    }
                    active_fields.push((event_start, depth, local.as_ref() == b"meta-field"));
                }
                if text_namespace && is_database_local(local.as_ref()) {
                    if !active_database_fields.is_empty() {
                        return Err(Error::InvalidFormat(
                            "nested database fields are not allowed".to_string(),
                        ));
                    }
                    active_database_fields.push((event_start, depth));
                }
            },
            Event::Empty(ref element) => {
                let local = element.local_name();
                let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                    Error::InvalidFormat(
                        "content.xml byte offset exceeds platform limits".to_string(),
                    )
                })?;
                if text_namespace && local.as_ref() == b"p" {
                    if paragraph_count >= MAX_PARAGRAPHS {
                        return Err(Error::InvalidFormat("too many ODF paragraphs".to_string()));
                    }
                    if wanted_paragraph == Some(paragraph_count) {
                        let qualified_name = std::str::from_utf8(element.name().as_ref())
                            .map_err(|_| {
                                Error::InvalidFormat(
                                    "ODF paragraph name is not valid UTF-8".to_string(),
                                )
                            })?
                            .to_string();
                        output.paragraph = Some(ParagraphSite::Empty {
                            start: event_start,
                            end: event_end,
                            qualified_name,
                        });
                    }
                    paragraph_count += 1;
                }
                if text_namespace && is_dynamic_local(local.as_ref()) {
                    if active_fields
                        .last()
                        .is_some_and(|(_, _, allows_nested)| !allows_nested)
                    {
                        return Err(Error::InvalidFormat(
                            "nested dynamic text fields are not allowed".to_string(),
                        ));
                    }
                    if output.fields.len() >= MAX_DYNAMIC_FIELDS {
                        return Err(Error::InvalidFormat(
                            "too many dynamic text fields".to_string(),
                        ));
                    }
                    output.fields.push(Span {
                        start: event_start,
                        end: event_end,
                    });
                }
                if text_namespace && is_database_local(local.as_ref()) {
                    output.database_fields.push(Span {
                        start: event_start,
                        end: event_end,
                    });
                }
            },
            Event::End(_) => {
                if active_paragraph_depth == Some(depth) {
                    output.paragraph = Some(ParagraphSite::BeforeEnd(event_start));
                    active_paragraph_depth = None;
                }
                if active_fields
                    .last()
                    .is_some_and(|(_, field_depth, _)| *field_depth == depth)
                {
                    let (start, _, _) = active_fields.pop().expect("checked active field");
                    let end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        Error::InvalidFormat(
                            "content.xml byte offset exceeds platform limits".to_string(),
                        )
                    })?;
                    output.fields.push(Span { start, end });
                }
                if active_database_fields
                    .last()
                    .is_some_and(|(_, field_depth)| *field_depth == depth)
                {
                    let (start, _) = active_database_fields
                        .pop()
                        .expect("checked database field");
                    let end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        Error::InvalidFormat(
                            "content.xml byte offset exceeds platform limits".to_string(),
                        )
                    })?;
                    output.database_fields.push(Span { start, end });
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("content.xml nesting depth underflow".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !active_fields.is_empty() || !active_database_fields.is_empty() {
        return Err(Error::InvalidFormat(
            "unterminated dynamic text field".to_string(),
        ));
    }
    output.fields.sort_by_key(|span| span.start);
    output.database_fields.sort_by_key(|span| span.start);
    Ok(output)
}

fn is_database_local(local: &[u8]) -> bool {
    matches!(
        local,
        b"database-display"
            | b"database-next"
            | b"database-row-select"
            | b"database-row-number"
            | b"database-name"
    )
}

fn is_dynamic_local(local: &[u8]) -> bool {
    matches!(
        local,
        b"placeholder"
            | b"conditional-text"
            | b"hidden-text"
            | b"hidden-paragraph"
            | b"sequence"
            | b"sequence-ref"
            | b"variable-set"
            | b"variable-get"
            | b"expression"
            | b"variable-input"
            | b"user-field-get"
            | b"user-field-input"
            | b"text-input"
            | b"drop-down"
            | b"script"
            | b"table-formula"
            | b"measure"
            | b"reference-ref"
            | b"bookmark-ref"
            | b"note-ref"
            | b"page-count"
            | b"paragraph-count"
            | b"word-count"
            | b"character-count"
            | b"table-count"
            | b"image-count"
            | b"object-count"
            | b"page-number"
            | b"date"
            | b"time"
            | b"page-continuation"
            | b"page-variable-set"
            | b"page-variable-get"
            | b"creation-date"
            | b"creation-time"
            | b"print-date"
            | b"print-time"
            | b"editing-cycles"
            | b"editing-duration"
            | b"modification-date"
            | b"modification-time"
            | b"initial-creator"
            | b"description"
            | b"printed-by"
            | b"title"
            | b"subject"
            | b"keywords"
            | b"creator"
            | b"user-defined"
            | b"meta-field"
            | b"dde-connection"
    )
}
