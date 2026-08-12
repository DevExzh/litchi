//! Worksheet-context extraction and byte-preserving publication for auto-filters.

use std::ops::Range as ByteRange;

use litchi_ooxml_common::mce::process_ooxml;
use quick_xml::Writer;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::codec::{
    parse_auto_filter_fragment, write_auto_filter_fragment, write_auto_filter_fragment_in,
};
use super::model::{CORE, Definition, MAX_FRAGMENT_BYTES, STRICT};
use crate::error::{Error, Result, allocation, invalid};

const MAX_WORKSHEET_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_EVENTS: usize = 1_000_000;
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

const SUCCESSORS: &[&[u8]] = &[
    b"mergeCells",
    b"phoneticPr",
    b"conditionalFormatting",
    b"dataValidations",
    b"hyperlinks",
    b"printOptions",
    b"pageMargins",
    b"pageSetup",
    b"headerFooter",
    b"rowBreaks",
    b"colBreaks",
    b"customProperties",
    b"cellWatches",
    b"ignoredErrors",
    b"smartTags",
    b"drawing",
    b"legacyDrawing",
    b"legacyDrawingHF",
    b"picture",
    b"oleObjects",
    b"controls",
    b"webPublishItems",
    b"tableParts",
    b"extLst",
];

/// Parse the worksheet's effective direct auto-filter declaration.
pub fn parse_auto_filter(xml: &[u8]) -> Result<Option<Definition>> {
    let processed = process_ooxml(xml)?;
    let Some(fragment) = capture(processed.as_ref())? else {
        return Ok(None);
    };
    parse_auto_filter_fragment(&fragment).map(Some)
}

/// Replace, insert, or remove one direct worksheet auto-filter declaration.
///
/// Every byte outside the selected `autoFilter` span is retained exactly.
/// A worksheet whose effective declaration is selected through markup
/// compatibility is refused because the active physical branch cannot be
/// rewritten without normalizing unrelated source markup.
pub fn replace_auto_filter(xml: &[u8], value: Option<&Definition>) -> Result<Vec<u8>> {
    if xml.len() > MAX_WORKSHEET_BYTES {
        return Err(invalid("auto-filter worksheet XML is too large"));
    }
    let effective = parse_auto_filter(xml)?;
    if effective.as_ref() == value {
        return Ok(xml.to_vec());
    }
    if let Some(value) = value {
        validate_definition(value)?;
    }
    let layout = scan_layout(xml)?;
    if layout.alternate_content || (effective.is_some() && layout.span.is_none()) {
        return Err(invalid(
            "autoFilter selected through markup compatibility cannot be edited byte-exactly",
        ));
    }
    let replacement = match value {
        Some(value) => write_auto_filter_fragment_in(value, layout.namespace)?,
        None => Vec::new(),
    };
    let span = layout.span.unwrap_or(layout.insertion..layout.insertion);
    let capacity = xml
        .len()
        .checked_sub(span.len())
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or_else(|| invalid("auto-filter worksheet output size overflow"))?;
    if capacity > MAX_WORKSHEET_BYTES {
        return Err(invalid("auto-filter worksheet output is too large"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("auto-filter worksheet output", source))?;
    output.extend_from_slice(&xml[..span.start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&xml[span.end..]);
    if parse_auto_filter(&output)?.as_ref() != value {
        return Err(invalid("auto-filter worksheet write verification failed"));
    }
    Ok(output)
}

pub(crate) fn validate_definition(value: &Definition) -> Result<()> {
    let fragment = write_auto_filter_fragment(value)?;
    if parse_auto_filter_fragment(&fragment)? != *value {
        return Err(invalid("auto-filter semantic write verification failed"));
    }
    Ok(())
}

struct Layout {
    namespace: &'static [u8],
    span: Option<ByteRange<usize>>,
    insertion: usize,
    alternate_content: bool,
}

fn scan_layout(xml: &[u8]) -> Result<Layout> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_namespace = None;
    let mut root_close = None;
    let mut insertion = None;
    let mut open = None;
    let mut span = None;
    let mut alternate_content = false;
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("auto-filter worksheet event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("auto-filter worksheet exceeds event limit"));
        }
        let start = position(&reader)?;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if mce(&namespace) && element.local_name().as_ref() == b"AlternateContent" {
                    alternate_content = true;
                }
                if depth == 0 {
                    if element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("auto-filter publication requires a worksheet root"));
                    }
                    root_namespace = spreadsheet_namespace(&namespace);
                    if root_namespace.is_none() {
                        return Err(invalid(
                            "auto-filter publication requires a SpreadsheetML worksheet root",
                        ));
                    }
                } else if depth == 1 && spreadsheet(&namespace) {
                    let local = element.local_name();
                    if local.as_ref() == b"autoFilter" {
                        if span.is_some() || open.replace(start).is_some() {
                            return Err(invalid("duplicate direct worksheet autoFilter"));
                        }
                    } else if SUCCESSORS.contains(&local.as_ref()) && insertion.is_none() {
                        insertion = Some(start);
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("auto-filter worksheet nesting overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid("auto-filter worksheet nesting is too deep"));
                }
            },
            Event::Empty(element) => {
                if mce(&namespace) && element.local_name().as_ref() == b"AlternateContent" {
                    alternate_content = true;
                }
                if depth == 1 && spreadsheet(&namespace) {
                    let local = element.local_name();
                    if local.as_ref() == b"autoFilter" {
                        if span.replace(start..end).is_some() || open.is_some() {
                            return Err(invalid("duplicate direct worksheet autoFilter"));
                        }
                    } else if SUCCESSORS.contains(&local.as_ref()) && insertion.is_none() {
                        insertion = Some(start);
                    }
                }
            },
            Event::End(element) => {
                if depth == 2
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"autoFilter"
                {
                    let selected = open
                        .take()
                        .ok_or_else(|| invalid("autoFilter close has no direct start"))?;
                    span = Some(selected..end);
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unexpected auto-filter worksheet end element"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "auto-filter publication rejects DTD and processing instructions",
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0 || open.is_some() {
        return Err(invalid("incomplete auto-filter worksheet XML"));
    }
    Ok(Layout {
        namespace: root_namespace.ok_or_else(|| invalid("worksheet XML has no root"))?,
        span,
        insertion: insertion
            .or(root_close)
            .ok_or_else(|| invalid("worksheet has no autoFilter insertion point"))?,
        alternate_content,
    })
}

fn capture(xml: &[u8]) -> Result<Option<Vec<u8>>> {
    let mut reader = NsReader::from_reader(xml);
    let mut capture: Option<(usize, Writer<Vec<u8>>)> = None;
    let mut result = None;
    let mut document_depth = 0usize;
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("auto-filter capture event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("auto-filter capture exceeds event limit"));
        }
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match &event {
            Event::Start(_) => {
                document_depth = document_depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("auto-filter document nesting overflow"))?;
                if document_depth > MAX_DEPTH {
                    return Err(invalid("auto-filter document nesting is too deep"));
                }
            },
            Event::End(_) => {
                document_depth = document_depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("auto-filter document nesting underflow"))?;
            },
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::Eof => {},
        }
        if let Some((depth, writer)) = capture.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            match event {
                Event::Start(_) => {
                    *depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("autoFilter capture nesting overflow"))?;
                    if *depth > MAX_DEPTH {
                        return Err(invalid("autoFilter capture nesting is too deep"));
                    }
                },
                Event::End(_) => {
                    *depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("autoFilter capture nesting underflow"))?;
                },
                Event::Empty(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_)
                | Event::Eof => {},
            }
            if *depth == 0 {
                let (_, writer) = capture.take().unwrap_or_else(|| {
                    crate::error::panic_missing_invariant(
                        "required value was checked before extraction",
                    )
                });
                let value = writer.into_inner();
                if value.len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("autoFilter is too large"));
                }
                if result.replace(value).is_some() {
                    return Err(invalid("duplicate worksheet autoFilter"));
                }
            }
            continue;
        }
        match event {
            Event::Start(e)
                if spreadsheet(&namespace) && e.local_name().as_ref() == b"autoFilter" =>
            {
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Start(e)).map_err(xml_error)?;
                capture = Some((1, writer));
            },
            Event::Empty(e)
                if spreadsheet(&namespace) && e.local_name().as_ref() == b"autoFilter" =>
            {
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Empty(e)).map_err(xml_error)?;
                if result.replace(writer.into_inner()).is_some() {
                    return Err(invalid("duplicate worksheet autoFilter"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if capture.is_some() {
        return Err(invalid("unterminated autoFilter"));
    }
    if document_depth != 0 {
        return Err(invalid("incomplete auto-filter worksheet XML"));
    }
    Ok(result)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source| invalid("auto-filter XML position does not fit usize"))
}

fn spreadsheet_namespace(namespace: &ResolveResult<'_>) -> Option<&'static [u8]> {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == CORE => Some(CORE),
        ResolveResult::Bound(value) if value.as_ref() == STRICT => Some(STRICT),
        _ => None,
    }
}

fn spreadsheet(namespace: &ResolveResult<'_>) -> bool {
    spreadsheet_namespace(namespace).is_some()
}

fn mce(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == MCE)
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
