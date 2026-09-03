#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::ref_option,
    reason = "the public API shape is retained for compatibility"
)]
//! Streaming text extraction for paragraph and run content.

use crate::error::{Error, Result};
use litchi_core::{
    SequentialTextWriter, TextObjectKind, TextOutputError, TextOutputOptions, TextOutputReport,
};
use litchi_ooxml_common::private::BindingTracker;
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::reader::Reader;
use std::io::Write;

use super::super::model::Paragraph;
use super::xml::is_fragment_word_name;

/// Maximum nesting depth accepted when extracting paragraph text.
const MAX_TEXT_SCAN_DEPTH: usize = 128;
/// Maximum number of elements scanned while extracting paragraph text.
const MAX_TEXT_SCAN_NODES: usize = 1_000_000;
/// Maximum raw text/reference bytes decoded in one borrowed chunk.
const MAX_TEXT_DECODE_CHUNK_BYTES: usize = 4096;
const MAX_TEXT_REFERENCE_BYTES: usize = 64 * 1024;

pub(crate) fn extract_word_text(xml_bytes: &[u8]) -> Result<String> {
    let mut result = String::new();
    for_each_word_text_chunk(xml_bytes, |chunk| {
        result
            .try_reserve(chunk.len())
            .map_err(|source| Error::Allocation {
                resource: "Word paragraph text",
                source,
            })?;
        result.push_str(chunk);
        Ok::<(), Error>(())
    })?;
    Ok(result)
}

pub(crate) fn for_each_word_text_chunk<F, E>(
    xml_bytes: &[u8],
    mut append: F,
) -> std::result::Result<(), E>
where
    F: FnMut(&str) -> std::result::Result<(), E>,
    E: From<Error>,
{
    // Plain reader + hand-rolled binding maintenance (change 0229, the
    // litchi-odt 0227 analog): the tracker replicates the push/pop
    // `NsReader` performs inside `read_resolved_event` (the
    // `BindingTracker` byte-exactness contract). Both the old slice-backed
    // `NsReader` and this plain `Reader` borrow their events; the removed work
    // is namespace maintenance, not an event-buffer copy.
    // `NsReader::from_reader` is `Reader::from_reader` with default
    // configuration, so the tokenization and error stream are unchanged.
    let mut reader = Reader::from_reader(xml_bytes);
    let mut tracker = BindingTracker::new();
    let mut pending_pop = false;
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_depth = None;

    loop {
        // The deferred pop of the previous `End`/`Empty` scope runs before
        // the read, exactly where `NsReader::read_event_impl` applies it.
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }
        let event = reader
            .read_event()
            .map_err(|error| E::from(Error::Xml(error.to_string())))?;
        // The push for a `Start`/`Empty` runs before the event is
        // classified, so a namespace error preempts the event exactly where
        // `read_resolved_event` returned `Err`. A push error is a real
        // `NamespaceError`, whose `Display` is what
        // `quick_xml::Error::Namespace` forwards to, so the `Error::Xml`
        // message is byte-identical to the historical failure.
        //
        // `resolve_event` maps `Start`/`Empty`/`End` to
        // `resolve(name, use_default = true)` and every other event to
        // `Unbound`; the `End` name resolves in its own scope because the
        // pop is deferred to the next read. This path consumes the `End`
        // resolution (the `text_depth` closing match below), unlike the
        // litchi-odt text path.
        let namespace = match &event {
            Event::Start(element) => {
                tracker
                    .push(element)
                    .map_err(|error| E::from(Error::Xml(error.to_string())))?;
                tracker.resolve_element(element.name()).0
            },
            Event::Empty(element) => {
                tracker
                    .push(element)
                    .map_err(|error| E::from(Error::Xml(error.to_string())))?;
                // The scope an `Empty` element opens closes immediately:
                // defer its pop to the top of the next iteration.
                pending_pop = true;
                tracker.resolve_element(element.name()).0
            },
            Event::End(element) => {
                pending_pop = true;
                tracker.resolve_element(element.name()).0
            },
            _ => ResolveResult::Unbound,
        };

        if fragment_prefix.is_none()
            && let Event::Start(element) = &event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            let prefix = element.name().prefix().map(|prefix| prefix.into_inner());
            fragment_prefix = Some(match prefix {
                Some(prefix) => Some(checked_text_vec_clone(
                    prefix,
                    "Word text namespace prefix",
                )?),
                None => None,
            });
        }

        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    E::from(Error::InvalidFormat(
                        "Word XML element counter overflow".to_string(),
                    ))
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(E::from(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    ))));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    E::from(Error::InvalidFormat(
                        "Word XML nesting is too deep".to_string(),
                    ))
                })?;
                if depth > MAX_TEXT_SCAN_DEPTH {
                    return Err(E::from(Error::InvalidFormat(format!(
                        "Word XML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                    ))));
                }
                if text_depth.is_none()
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = Some(depth);
                } else if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    let mut encoded = [0_u8; 4];
                    append(character.encode_utf8(&mut encoded))?;
                }
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    E::from(Error::InvalidFormat(
                        "Word XML element counter overflow".to_string(),
                    ))
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(E::from(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    ))));
                }
                if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    let mut encoded = [0_u8; 4];
                    append(character.encode_utf8(&mut encoded))?;
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                append_xml_text_chunks(text.as_ref(), &mut append)?;
            },
            Event::CData(text) if text_depth.is_some() => {
                append_xml_text_chunks_without_entities(text.as_ref(), &mut append)?;
            },
            Event::GeneralRef(reference) => {
                let decoded = decode_word_xml_reference(&reference).map_err(E::from)?;
                if text_depth.is_some() {
                    append_utf8_str_chunks(&decoded, &mut append)?;
                }
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    E::from(Error::InvalidFormat("invalid Word XML nesting".to_string()))
                })?;
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(E::from(Error::InvalidFormat(
                    "unterminated Word XML".to_string(),
                )));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_) => {},
        }
    }
    Ok(())
}

fn checked_text_vec_clone(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

fn is_supported_xml_reference(reference: &[u8]) -> bool {
    if reference == b"lt"
        || reference == b"gt"
        || reference == b"amp"
        || reference == b"apos"
        || reference == b"quot"
    {
        return true;
    }
    let (digits, hexadecimal) = if reference.first() == Some(&b'#') {
        if reference.get(1) == Some(&b'x') || reference.get(1) == Some(&b'X') {
            (reference.get(2..).unwrap_or_default(), true)
        } else {
            (reference.get(1..).unwrap_or_default(), false)
        }
    } else {
        return false;
    };
    !digits.is_empty()
        && digits.iter().all(|digit| {
            if hexadecimal {
                digit.is_ascii_hexdigit()
            } else {
                digit.is_ascii_digit()
            }
        })
}

fn decode_word_xml_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    let reference_bytes = reference.as_ref();
    if reference_bytes.len() > MAX_TEXT_REFERENCE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "Word XML reference exceeds the {MAX_TEXT_REFERENCE_BYTES}-byte limit"
        )));
    }
    if !is_supported_xml_reference(reference_bytes) {
        return Err(Error::InvalidFormat(
            "unsupported Word XML general entity reference".to_string(),
        ));
    }
    decode_xml_reference(reference).map_err(Error::from)
}

fn append_xml_text_chunks<F, E>(raw: &[u8], append: &mut F) -> std::result::Result<(), E>
where
    F: FnMut(&str) -> std::result::Result<(), E>,
    E: From<Error>,
{
    let mut cursor = 0usize;
    let mut plain_start = 0usize;
    while cursor < raw.len() {
        if raw[cursor] != b'&' {
            cursor += 1;
            continue;
        }
        append_xml_text_chunks_without_entities(&raw[plain_start..cursor], append)?;
        let entity_end = raw[cursor + 1..]
            .iter()
            .position(|byte| *byte == b';')
            .map(|relative| cursor + 1 + relative)
            .ok_or_else(|| {
                E::from(Error::InvalidFormat(
                    "unterminated Word XML text reference".to_string(),
                ))
            })?;
        let entity = &raw[cursor..=entity_end];
        if entity.len() > MAX_TEXT_REFERENCE_BYTES {
            return Err(E::from(Error::InvalidFormat(format!(
                "Word XML reference exceeds the {MAX_TEXT_REFERENCE_BYTES}-byte limit"
            ))));
        }
        let entity = std::str::from_utf8(entity).map_err(|_| {
            E::from(Error::InvalidFormat(
                "Word XML text reference is not valid UTF-8".to_string(),
            ))
        })?;
        let decoded = quick_xml::escape::unescape(entity)
            .map_err(|error| E::from(Error::Xml(error.to_string())))?;
        append_utf8_str_chunks(decoded.as_ref(), append)?;
        cursor = entity_end + 1;
        plain_start = cursor;
    }
    append_xml_text_chunks_without_entities(&raw[plain_start..], append)
}

fn append_xml_text_chunks_without_entities<F, E>(
    raw: &[u8],
    append: &mut F,
) -> std::result::Result<(), E>
where
    F: FnMut(&str) -> std::result::Result<(), E>,
    E: From<Error>,
{
    let mut cursor = 0usize;
    let mut plain_start = 0usize;
    while cursor < raw.len() {
        match raw[cursor] {
            b'\r' => {
                append_utf8_chunks(&raw[plain_start..cursor], append)?;
                append("\n")?;
                cursor += if raw.get(cursor + 1) == Some(&b'\n') {
                    2
                } else {
                    1
                };
                plain_start = cursor;
            },
            b'\n' => {
                append_utf8_chunks(&raw[plain_start..cursor], append)?;
                append("\n")?;
                cursor += 1;
                plain_start = cursor;
            },
            _ => cursor += 1,
        }
    }
    append_utf8_chunks(&raw[plain_start..], append)
}

fn append_utf8_chunks<F, E>(value: &[u8], append: &mut F) -> std::result::Result<(), E>
where
    F: FnMut(&str) -> std::result::Result<(), E>,
    E: From<Error>,
{
    let value = std::str::from_utf8(value).map_err(|_| {
        E::from(Error::InvalidFormat(
            "Word XML text is not valid UTF-8".to_string(),
        ))
    })?;
    append_utf8_str_chunks(value, append)
}

fn append_utf8_str_chunks<F, E>(value: &str, append: &mut F) -> std::result::Result<(), E>
where
    F: FnMut(&str) -> std::result::Result<(), E>,
    E: From<Error>,
{
    let mut cursor = 0usize;
    while cursor < value.len() {
        let mut end =
            value.len().min(
                cursor
                    .checked_add(MAX_TEXT_DECODE_CHUNK_BYTES)
                    .ok_or_else(|| {
                        E::from(Error::InvalidFormat(
                            "Word XML text chunk length overflow".to_string(),
                        ))
                    })?,
            );
        while end > cursor && !value.is_char_boundary(end) {
            end -= 1;
        }
        if end == cursor {
            return Err(E::from(Error::InvalidFormat(
                "Word XML text chunk is not valid UTF-8".to_string(),
            )));
        }
        append(&value[cursor..end])?;
        cursor = end;
    }
    Ok(())
}

// The sink path deliberately has its own bounded parser. The established
// `extract_word_text` projection is retained for compatibility and may return
// one caller-owned String for legacy APIs; this path must never accumulate a
// document-wide semantic String or paragraph collection.
pub(crate) const MAX_SEMANTIC_TEXT_RAW_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_SEMANTIC_TEXT_EVENTS: usize = 1_000_000;
const MAX_SEMANTIC_TEXT_DEPTH: usize = 128;
const MAX_SEMANTIC_TEXT_PARAGRAPHS: usize = 1_000_000;
const MAX_SEMANTIC_TEXT_RUNS: usize = 4_000_000;
const MAX_SEMANTIC_TEXT_EVENT_BYTES: usize = 1024 * 1024;
const MAX_SEMANTIC_TEXT_NAME_BYTES: usize = 64 * 1024;
const MAX_SEMANTIC_TEXT_REFERENCE_BYTES: usize = 64 * 1024;
const MAX_SEMANTIC_TEXT_ATTRIBUTE_BYTES: usize = 1024 * 1024;
const MAX_SEMANTIC_TEXT_NAMESPACE_BINDINGS: usize = 4096;
const MAX_SEMANTIC_TEXT_NAMESPACE_BYTES: usize = 4 * 1024 * 1024;
const MAX_SEMANTIC_TEXT_PARAGRAPH_BYTES: usize = 16 * 1024 * 1024;
const MAX_SEMANTIC_TEXT_DOCUMENT_BYTES: usize = 64 * 1024 * 1024;

pub(crate) const fn semantic_text_raw_xml_limit() -> usize {
    MAX_SEMANTIC_TEXT_RAW_XML_BYTES
}

pub(crate) fn write_text_to<W: Write + ?Sized>(
    xml_bytes: &[u8],
    output: &mut W,
    options: TextOutputOptions<'_>,
) -> std::result::Result<TextOutputReport, TextOutputError<Error>> {
    let mut writer = SequentialTextWriter::new(output, options);
    write_text_to_with_writer(xml_bytes, &mut writer)?;
    Ok(writer.finish())
}

pub(crate) fn write_text_to_with_writer<'options, 'output, W: Write + ?Sized>(
    xml_bytes: &[u8],
    writer: &mut SequentialTextWriter<'options, 'output, W>,
) -> std::result::Result<(), TextOutputError<Error>> {
    write_text_to_with_operation_check(xml_bytes, writer, || Ok(()))
}

pub(crate) fn write_text_to_with_operation_check<'options, 'output, W, F>(
    xml_bytes: &[u8],
    writer: &mut SequentialTextWriter<'options, 'output, W>,
    mut operation_check: F,
) -> std::result::Result<(), TextOutputError<Error>>
where
    W: Write + ?Sized,
    F: FnMut() -> Result<()>,
{
    if xml_bytes.len() > MAX_SEMANTIC_TEXT_RAW_XML_BYTES {
        return Err(writer.document_error(Error::InvalidFormat(format!(
            "semantic DOCX raw XML exceeds {MAX_SEMANTIC_TEXT_RAW_XML_BYTES} bytes"
        ))));
    }
    if xml_bytes.len() > MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES {
        return Err(writer.document_error(Error::InvalidFormat(format!(
            "semantic DOCX processed XML exceeds {MAX_SEMANTIC_TEXT_PROCESSED_XML_BYTES} bytes"
        ))));
    }

    operation_check().map_err(|error| writer.document_error(error))?;
    preflight_semantic_xml(xml_bytes, &mut operation_check)
        .map_err(|error| writer.document_error(error))?;
    operation_check().map_err(|error| writer.document_error(error))?;

    let mut reader = NsReader::from_reader(xml_bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut parser = SemanticTextParser::default();
    let mut buffer = Vec::new();
    loop {
        operation_check().map_err(|error| writer.document_error(error))?;
        buffer.clear();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| writer.document_error(Error::Xml(error.to_string())))?;
        operation_check().map_err(|error| writer.document_error(error))?;
        let event = event.into_owned();
        if let Event::Start(element) | Event::Empty(element) = &event {
            validate_semantic_attribute_names(&reader, element)
                .map_err(|error| writer.document_error(error))?;
        }
        let namespace = match &event {
            Event::Start(element) | Event::Empty(element) => {
                reader.resolver().resolve_element(element.name()).0
            },
            Event::End(element) => reader.resolver().resolve_element(element.name()).0,
            _ => ResolveResult::Unbound,
        };
        if parser
            .consume(namespace, event, writer, &mut operation_check)
            .map_err(|error| match error {
                SemanticTextFailure::Document(source) => writer.document_error(source),
                SemanticTextFailure::Output(error) => error,
            })?
        {
            return Ok(());
        }
        operation_check().map_err(|error| writer.document_error(error))?;
    }
}

#[derive(Default)]
struct SemanticTextXmlBudget {
    events: usize,
    depth: usize,
    namespace_bindings: usize,
    namespace_bytes: usize,
}

impl SemanticTextXmlBudget {
    fn observe_event(&mut self, event: &Event<'_>) -> Result<()> {
        self.events = self.events.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("semantic DOCX XML event counter overflow".into())
        })?;
        if self.events > MAX_SEMANTIC_TEXT_EVENTS {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX XML exceeds {MAX_SEMANTIC_TEXT_EVENTS} events"
            )));
        }
        let bytes = semantic_event_bytes(event);
        if bytes > MAX_SEMANTIC_TEXT_EVENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX XML event exceeds {MAX_SEMANTIC_TEXT_EVENT_BYTES} bytes"
            )));
        }
        Ok(())
    }

    fn observe_namespaces(&mut self, element: &quick_xml::events::BytesStart<'_>) -> Result<()> {
        if element.name().as_ref().len() > MAX_SEMANTIC_TEXT_NAME_BYTES {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX XML element name exceeds {MAX_SEMANTIC_TEXT_NAME_BYTES} bytes"
            )));
        }
        let mut attribute_bytes = 0usize;
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
            let key = attribute.key.as_ref();
            attribute_bytes = attribute_bytes
                .checked_add(key.len())
                .and_then(|value| value.checked_add(attribute.value.len()))
                .ok_or_else(|| {
                    Error::InvalidFormat("semantic DOCX attribute length overflow".into())
                })?;
            if attribute_bytes > MAX_SEMANTIC_TEXT_ATTRIBUTE_BYTES {
                return Err(Error::InvalidFormat(format!(
                    "semantic DOCX XML exceeds {MAX_SEMANTIC_TEXT_ATTRIBUTE_BYTES} attribute bytes"
                )));
            }
            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                self.namespace_bindings =
                    self.namespace_bindings.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("semantic DOCX namespace counter overflow".into())
                    })?;
                self.namespace_bytes = self
                    .namespace_bytes
                    .checked_add(key.len())
                    .and_then(|value| value.checked_add(attribute.value.len()))
                    .ok_or_else(|| {
                        Error::InvalidFormat("semantic DOCX namespace byte counter overflow".into())
                    })?;
                if self.namespace_bindings > MAX_SEMANTIC_TEXT_NAMESPACE_BINDINGS {
                    return Err(Error::InvalidFormat(format!(
                        "semantic DOCX XML exceeds {MAX_SEMANTIC_TEXT_NAMESPACE_BINDINGS} namespace bindings"
                    )));
                }
                if self.namespace_bytes > MAX_SEMANTIC_TEXT_NAMESPACE_BYTES {
                    return Err(Error::InvalidFormat(format!(
                        "semantic DOCX XML exceeds {MAX_SEMANTIC_TEXT_NAMESPACE_BYTES} namespace bytes"
                    )));
                }
            }
        }
        Ok(())
    }

    fn start(&mut self) -> Result<()> {
        self.depth = self.depth.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("semantic DOCX XML depth counter overflow".into())
        })?;
        if self.depth > MAX_SEMANTIC_TEXT_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX XML exceeds {MAX_SEMANTIC_TEXT_DEPTH} depth"
            )));
        }
        Ok(())
    }

    fn end(&mut self) -> Result<()> {
        self.depth = self.depth.checked_sub(1).ok_or_else(|| {
            Error::InvalidFormat("semantic DOCX XML has an unexpected closing element".into())
        })?;
        Ok(())
    }
}

fn preflight_semantic_xml<F>(xml_bytes: &[u8], operation_check: &mut F) -> Result<()>
where
    F: FnMut() -> Result<()>,
{
    let mut reader = Reader::from_reader(xml_bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut budget = SemanticTextXmlBudget::default();
    let mut root_seen = false;
    let mut root_closed = false;

    loop {
        operation_check()?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        operation_check()?;
        budget.observe_event(&event)?;
        match event {
            Event::Start(element) => {
                if budget.depth == 0 {
                    if root_seen || root_closed {
                        return Err(Error::InvalidFormat(
                            "semantic DOCX XML has multiple roots".into(),
                        ));
                    }
                    root_seen = true;
                }
                budget.observe_namespaces(&element)?;
                budget.start()?;
            },
            Event::Empty(element) => {
                if budget.depth == 0 {
                    if root_seen || root_closed {
                        return Err(Error::InvalidFormat(
                            "semantic DOCX XML has multiple roots".into(),
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                }
                let empty_depth = budget.depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("semantic DOCX XML depth counter overflow".into())
                })?;
                if empty_depth > MAX_SEMANTIC_TEXT_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "semantic DOCX XML exceeds {MAX_SEMANTIC_TEXT_DEPTH} depth"
                    )));
                }
                budget.observe_namespaces(&element)?;
            },
            Event::End(element) => {
                if element.name().as_ref().len() > MAX_SEMANTIC_TEXT_NAME_BYTES {
                    return Err(Error::InvalidFormat(format!(
                        "semantic DOCX XML end name exceeds {MAX_SEMANTIC_TEXT_NAME_BYTES} bytes"
                    )));
                }
                budget.end()?;
                if budget.depth == 0 {
                    root_closed = true;
                }
            },
            Event::GeneralRef(reference) => {
                if reference.as_ref().len() > MAX_SEMANTIC_TEXT_REFERENCE_BYTES {
                    return Err(Error::InvalidFormat(format!(
                        "semantic DOCX XML reference exceeds {MAX_SEMANTIC_TEXT_REFERENCE_BYTES} bytes"
                    )));
                }
            },
            Event::Eof => {
                if !root_seen {
                    return Err(Error::InvalidFormat(
                        "semantic DOCX XML lacks an element root".into(),
                    ));
                }
                if budget.depth != 0 {
                    return Err(Error::InvalidFormat(
                        "semantic DOCX XML has unbalanced elements".into(),
                    ));
                }
                return Ok(());
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_) => {},
        }
    }
}

fn semantic_event_bytes(event: &Event<'_>) -> usize {
    match event {
        Event::Start(element) => element.as_ref().len(),
        Event::Empty(element) => element.as_ref().len(),
        Event::End(element) => element.as_ref().len(),
        Event::Text(text) => text.as_ref().len(),
        Event::CData(text) => text.as_ref().len(),
        Event::Comment(comment) => comment.as_ref().len(),
        Event::DocType(doctype) => doctype.as_ref().len(),
        Event::PI(pi) => pi.as_ref().len(),
        Event::Decl(decl) => decl.as_ref().len(),
        Event::GeneralRef(reference) => reference.as_ref().len(),
        Event::Eof => 0,
    }
}

fn validate_semantic_attributes(element: &quick_xml::events::BytesStart<'_>) -> Result<()> {
    let mut total = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        total = total
            .checked_add(attribute.key.as_ref().len())
            .and_then(|value| value.checked_add(attribute.value.len()))
            .ok_or_else(|| {
                Error::InvalidFormat("semantic DOCX attribute length overflow".into())
            })?;
        if total > MAX_SEMANTIC_TEXT_ATTRIBUTE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX XML exceeds {MAX_SEMANTIC_TEXT_ATTRIBUTE_BYTES} attribute bytes"
            )));
        }
        let value = attribute
            .normalized_value(XmlVersion::Implicit1_0)
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_characters(&value)?;
    }
    Ok(())
}

fn validate_semantic_attribute_names(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        if let ResolveResult::Unknown(prefix) = reader.resolver().resolve_attribute(attribute.key).0
        {
            return Err(Error::InvalidFormat(format!(
                "unresolved semantic DOCX attribute namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        }
    }
    Ok(())
}

fn validate_semantic_element_namespace(namespace: &ResolveResult<'_>) -> Result<()> {
    if let ResolveResult::Unknown(prefix) = namespace {
        return Err(Error::InvalidFormat(format!(
            "unresolved semantic DOCX element namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )));
    }
    Ok(())
}

fn is_word_element(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == crate::namespace::WORDPROCESSINGML_NAMESPACE
                    || *value == crate::namespace::STRICT_WORDPROCESSINGML_NAMESPACE
        )
}

fn is_word_text_name(namespace: &ResolveResult<'_>, name: QName<'_>) -> bool {
    is_word_element(namespace, name, b"t")
}

fn is_word_control_name(namespace: &ResolveResult<'_>, name: QName<'_>) -> bool {
    is_word_element(namespace, name, b"tab")
        || is_word_element(namespace, name, b"br")
        || is_word_element(namespace, name, b"cr")
        || is_word_element(namespace, name, b"noBreakHyphen")
        || is_word_element(namespace, name, b"softHyphen")
}

fn validate_xml_characters(value: &str) -> Result<()> {
    if value.chars().all(|character| {
        matches!(
            character,
            '\u{9}'
                | '\u{a}'
                | '\u{d}'
                | '\u{20}'..='\u{d7ff}'
                | '\u{e000}'..='\u{fffd}'
                | '\u{10000}'..='\u{10ffff}'
        )
    }) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "semantic DOCX XML contains an invalid XML character".into(),
        ))
    }
}

fn validate_xml_comment(comment: &str) -> Result<()> {
    validate_xml_characters(comment)?;
    if comment.contains("--") || comment.ends_with('-') {
        return Err(Error::InvalidFormat(
            "semantic DOCX XML contains an invalid comment".into(),
        ));
    }
    Ok(())
}

enum SemanticTextFailure {
    Document(Error),
    Output(TextOutputError<Error>),
}

#[derive(Default)]
struct SemanticTextParser {
    budget: SemanticTextXmlBudget,
    root_seen: bool,
    declaration_seen: bool,
    document_event_seen: bool,
    paragraph_depth: Option<usize>,
    text_depth: Option<usize>,
    paragraph: Option<String>,
    paragraphs: usize,
    runs: usize,
    decoded_document_bytes: usize,
}

impl SemanticTextParser {
    fn increment(value: &mut usize, limit: usize, resource: &'static str) -> Result<()> {
        *value = value
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat(format!("{resource} counter overflow")))?;
        if *value > limit {
            return Err(Error::InvalidFormat(format!("{resource} exceeds {limit}")));
        }
        Ok(())
    }

    fn append(&mut self, value: &str) -> Result<()> {
        if value.is_empty() {
            return Ok(());
        }
        let paragraph = self.paragraph.as_mut().ok_or_else(|| {
            Error::InvalidFormat("semantic DOCX text appears outside a paragraph".into())
        })?;
        let paragraph_bytes = paragraph.len().checked_add(value.len()).ok_or_else(|| {
            Error::InvalidFormat("semantic DOCX paragraph length overflow".into())
        })?;
        if paragraph_bytes > MAX_SEMANTIC_TEXT_PARAGRAPH_BYTES {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX paragraph text exceeds {MAX_SEMANTIC_TEXT_PARAGRAPH_BYTES} bytes"
            )));
        }
        let document_bytes = self
            .decoded_document_bytes
            .checked_add(value.len())
            .ok_or_else(|| Error::InvalidFormat("semantic DOCX decoded length overflow".into()))?;
        if document_bytes > MAX_SEMANTIC_TEXT_DOCUMENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX decoded text exceeds {MAX_SEMANTIC_TEXT_DOCUMENT_BYTES} bytes"
            )));
        }
        paragraph
            .try_reserve(value.len())
            .map_err(|source| Error::Allocation {
                resource: "semantic DOCX current paragraph text",
                source,
            })?;
        paragraph.push_str(value);
        self.decoded_document_bytes = document_bytes;
        Ok(())
    }

    fn append_control(&mut self, character: char) -> Result<()> {
        let mut encoded = [0_u8; 4];
        self.append(character.encode_utf8(&mut encoded))
    }

    fn start_element(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &quick_xml::events::BytesStart<'_>,
    ) -> Result<()> {
        validate_semantic_attributes(element)?;
        let name = element.name();
        if name.local_name().as_ref() == b"t" && !is_word_text_name(namespace, name) {
            return Err(Error::InvalidFormat(
                "foreign semantic DOCX text element is not w:t".into(),
            ));
        }
        if is_word_control_name(namespace, name) && self.paragraph.is_none() {
            return Err(Error::InvalidFormat(
                "semantic DOCX control appears outside a paragraph".into(),
            ));
        }
        if is_word_element(namespace, name, b"p") {
            if self.paragraph_depth.is_some() {
                return Err(Error::InvalidFormat(
                    "nested semantic DOCX paragraphs are not supported".into(),
                ));
            }
            Self::increment(
                &mut self.paragraphs,
                MAX_SEMANTIC_TEXT_PARAGRAPHS,
                "semantic DOCX paragraphs",
            )?;
        } else if is_word_element(namespace, name, b"r") {
            Self::increment(&mut self.runs, MAX_SEMANTIC_TEXT_RUNS, "semantic DOCX runs")?;
        }
        if self.text_depth.is_some() {
            return Err(Error::InvalidFormat(
                "nested semantic DOCX elements inside w:t are not permitted".into(),
            ));
        }
        if is_word_control_name(namespace, name) {
            let character = if is_word_element(namespace, name, b"tab") {
                '\t'
            } else if is_word_element(namespace, name, b"noBreakHyphen") {
                '\u{2011}'
            } else if is_word_element(namespace, name, b"softHyphen") {
                '\u{00ad}'
            } else {
                '\n'
            };
            self.append_control(character)?;
        }
        Ok(())
    }

    fn empty_element<W: Write + ?Sized, F: FnMut() -> Result<()>>(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &quick_xml::events::BytesStart<'_>,
        writer: &mut SequentialTextWriter<'_, '_, W>,
        operation_check: &mut F,
    ) -> std::result::Result<(), SemanticTextFailure> {
        validate_semantic_attributes(element).map_err(SemanticTextFailure::Document)?;
        let name = element.name();
        if name.local_name().as_ref() == b"t" && !is_word_text_name(namespace, name) {
            return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                "foreign semantic DOCX text element is not w:t".into(),
            )));
        }
        if self.text_depth.is_some() {
            return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                "nested semantic DOCX elements inside w:t are not permitted".into(),
            )));
        }
        if is_word_element(namespace, name, b"p") {
            if self.paragraph_depth.is_some() {
                return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                    "nested semantic DOCX paragraphs are not supported".into(),
                )));
            }
            Self::increment(
                &mut self.paragraphs,
                MAX_SEMANTIC_TEXT_PARAGRAPHS,
                "semantic DOCX paragraphs",
            )
            .map_err(SemanticTextFailure::Document)?;
            operation_check().map_err(SemanticTextFailure::Document)?;
            writer
                .write_object::<Error>(TextObjectKind::Paragraph, "")
                .map_err(SemanticTextFailure::Output)?;
        } else if is_word_element(namespace, name, b"r") {
            Self::increment(&mut self.runs, MAX_SEMANTIC_TEXT_RUNS, "semantic DOCX runs")
                .map_err(SemanticTextFailure::Document)?;
        } else if is_word_text_name(namespace, name) {
            if self.paragraph_depth.is_none() {
                return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                    "semantic DOCX w:t appears outside a paragraph".into(),
                )));
            }
        } else if is_word_control_name(namespace, name) {
            if self.paragraph.is_none() {
                return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                    "semantic DOCX control appears outside a paragraph".into(),
                )));
            }
            let character = if is_word_element(namespace, name, b"tab") {
                '\t'
            } else if is_word_element(namespace, name, b"noBreakHyphen") {
                '\u{2011}'
            } else if is_word_element(namespace, name, b"softHyphen") {
                '\u{00ad}'
            } else {
                '\n'
            };
            self.append_control(character)
                .map_err(SemanticTextFailure::Document)?;
        }
        Ok(())
    }

    fn end_element<W: Write + ?Sized, F: FnMut() -> Result<()>>(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &quick_xml::events::BytesEnd<'_>,
        writer: &mut SequentialTextWriter<'_, '_, W>,
        operation_check: &mut F,
    ) -> std::result::Result<(), SemanticTextFailure> {
        validate_semantic_element_namespace(namespace).map_err(SemanticTextFailure::Document)?;
        let name = element.name();
        if name.local_name().as_ref() == b"t" && !is_word_text_name(namespace, name) {
            return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                "foreign semantic DOCX text element is not w:t".into(),
            )));
        }
        if is_word_text_name(namespace, name) {
            if self.text_depth != Some(self.budget.depth) {
                return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                    "unbalanced semantic DOCX w:t".into(),
                )));
            }
            self.text_depth = None;
        }
        if is_word_element(namespace, name, b"p") {
            if self.paragraph_depth != Some(self.budget.depth) {
                return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                    "unbalanced semantic DOCX paragraph".into(),
                )));
            }
            let paragraph = self.paragraph.take().ok_or_else(|| {
                SemanticTextFailure::Document(Error::InvalidFormat(
                    "semantic DOCX paragraph state is missing".into(),
                ))
            })?;
            self.paragraph_depth = None;
            operation_check().map_err(SemanticTextFailure::Document)?;
            writer
                .write_object::<Error>(TextObjectKind::Paragraph, &paragraph)
                .map_err(SemanticTextFailure::Output)?;
        }
        Ok(())
    }

    fn text_event(&mut self, text: &quick_xml::events::BytesText<'_>) -> Result<()> {
        let decoded = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let decoded =
            quick_xml::escape::unescape(&decoded).map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_characters(&decoded)?;
        if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX decoded text event exceeds {MAX_SEMANTIC_TEXT_EVENT_BYTES} bytes"
            )));
        }
        self.append(&decoded)
    }

    fn cdata_event(&mut self, text: &quick_xml::events::BytesCData<'_>) -> Result<()> {
        let decoded = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_characters(&decoded)?;
        if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX decoded CDATA event exceeds {MAX_SEMANTIC_TEXT_EVENT_BYTES} bytes"
            )));
        }
        self.append(&decoded)
    }

    fn reference_event(&mut self, reference: &quick_xml::events::BytesRef<'_>) -> Result<()> {
        if reference.as_ref().len() > MAX_SEMANTIC_TEXT_REFERENCE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX XML reference exceeds {MAX_SEMANTIC_TEXT_REFERENCE_BYTES} bytes"
            )));
        }
        let decoded = decode_xml_reference(reference)?;
        validate_xml_characters(&decoded)?;
        if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "semantic DOCX decoded reference exceeds {MAX_SEMANTIC_TEXT_EVENT_BYTES} bytes"
            )));
        }
        self.append(&decoded)
    }

    fn consume<W: Write + ?Sized, F: FnMut() -> Result<()>>(
        &mut self,
        namespace: ResolveResult<'_>,
        event: Event<'_>,
        writer: &mut SequentialTextWriter<'_, '_, W>,
        operation_check: &mut F,
    ) -> std::result::Result<bool, SemanticTextFailure> {
        self.budget
            .observe_event(&event)
            .map_err(SemanticTextFailure::Document)?;
        let declaration_is_first = !self.document_event_seen;
        if !matches!(event, Event::Eof) {
            self.document_event_seen = true;
        }
        match event {
            Event::Start(element) => {
                validate_semantic_element_namespace(&namespace)
                    .map_err(SemanticTextFailure::Document)?;
                if self.budget.depth == 0 {
                    if self.root_seen || !is_word_element(&namespace, element.name(), b"document") {
                        return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                            "semantic DOCX XML has an invalid root".into(),
                        )));
                    }
                    self.root_seen = true;
                }
                self.budget
                    .observe_namespaces(&element)
                    .map_err(SemanticTextFailure::Document)?;
                self.budget.start().map_err(SemanticTextFailure::Document)?;
                self.start_element(&namespace, &element)
                    .map_err(SemanticTextFailure::Document)?;
                let name = element.name();
                if is_word_element(&namespace, name, b"p") {
                    self.paragraph_depth = Some(self.budget.depth);
                    self.paragraph = Some(String::new());
                } else if is_word_text_name(&namespace, name) {
                    if self.paragraph_depth.is_none() {
                        return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                            "semantic DOCX w:t appears outside a paragraph".into(),
                        )));
                    }
                    self.text_depth = Some(self.budget.depth);
                }
            },
            Event::Empty(element) => {
                validate_semantic_element_namespace(&namespace)
                    .map_err(SemanticTextFailure::Document)?;
                if self.budget.depth == 0 {
                    if self.root_seen || !is_word_element(&namespace, element.name(), b"document") {
                        return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                            "semantic DOCX XML has an invalid root".into(),
                        )));
                    }
                    self.root_seen = true;
                }
                self.budget
                    .observe_namespaces(&element)
                    .map_err(SemanticTextFailure::Document)?;
                self.empty_element(&namespace, &element, writer, operation_check)?;
            },
            Event::End(element) => {
                self.end_element(&namespace, &element, writer, operation_check)?;
                self.budget.end().map_err(SemanticTextFailure::Document)?;
            },
            Event::Text(text) if self.text_depth.is_some() => {
                self.text_event(&text)
                    .map_err(SemanticTextFailure::Document)?;
            },
            Event::Text(text) => {
                let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    SemanticTextFailure::Document(Error::Xml(error.to_string()))
                })?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                    SemanticTextFailure::Document(Error::Xml(error.to_string()))
                })?;
                validate_xml_characters(&decoded).map_err(SemanticTextFailure::Document)?;
                if decoded.len() > MAX_SEMANTIC_TEXT_EVENT_BYTES {
                    return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                        "semantic DOCX decoded text event is too large".into(),
                    )));
                }
                if self.budget.depth == 0 && !decoded.as_bytes().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                        "semantic DOCX XML has text outside its root".into(),
                    )));
                }
            },
            Event::CData(text) if self.text_depth.is_some() => {
                self.cdata_event(&text)
                    .map_err(SemanticTextFailure::Document)?;
            },
            Event::CData(_) => {
                return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                    "semantic DOCX CDATA is outside w:t".into(),
                )));
            },
            Event::GeneralRef(reference) if self.text_depth.is_some() => {
                self.reference_event(&reference)
                    .map_err(SemanticTextFailure::Document)?;
            },
            Event::GeneralRef(_) => {
                return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                    "semantic DOCX XML reference is outside w:t".into(),
                )));
            },
            Event::Decl(_) => {
                if self.declaration_seen || !declaration_is_first || self.root_seen {
                    return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                        "XML declarations must be the first DOCX document event".into(),
                    )));
                }
                self.declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                    "DTD declarations are not permitted in semantic DOCX text".into(),
                )));
            },
            Event::PI(_) => {
                return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                    "processing instructions are not permitted in semantic DOCX text".into(),
                )));
            },
            Event::Comment(comment) => {
                let decoded = comment.decode().map_err(|error| {
                    SemanticTextFailure::Document(Error::Xml(error.to_string()))
                })?;
                validate_xml_comment(&decoded).map_err(SemanticTextFailure::Document)?;
            },
            Event::Eof => {
                if !self.root_seen {
                    return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                        "semantic DOCX XML lacks an element root".into(),
                    )));
                }
                if self.budget.depth != 0
                    || self.paragraph_depth.is_some()
                    || self.text_depth.is_some()
                {
                    return Err(SemanticTextFailure::Document(Error::InvalidFormat(
                        "semantic DOCX XML has unbalanced elements".into(),
                    )));
                }
                return Ok(true);
            },
        }
        Ok(false)
    }
}

fn word_special_character(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Option<char> {
    if is_fragment_word_name(namespace, name, b"tab", fragment_prefix) {
        Some('\t')
    } else if is_fragment_word_name(namespace, name, b"br", fragment_prefix)
        || is_fragment_word_name(namespace, name, b"cr", fragment_prefix)
    {
        Some('\n')
    } else if is_fragment_word_name(namespace, name, b"noBreakHyphen", fragment_prefix) {
        Some('\u{2011}')
    } else if is_fragment_word_name(namespace, name, b"softHyphen", fragment_prefix) {
        Some('\u{00ad}')
    } else {
        None
    }
}

impl Paragraph {
    /// Get the text content of this paragraph.
    ///
    /// Concatenates all text from all runs in the paragraph.
    ///
    /// # Performance
    ///
    /// Uses streaming XML parsing with pre-allocated buffer to extract text efficiently.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
    }
}

/// Change-0229 differential oracle: the pre-0229 `NsReader` implementation
/// of [`extract_word_text`], retained test-only so the tracker-driven path
/// is pinned byte-for-byte against it (the litchi-odt 0227 pattern, where
/// the oracle stayed in production because other APIs used it; here nothing
/// else needs it).
#[cfg(test)]
fn extract_word_text_nsreader_oracle(xml_bytes: &[u8]) -> Result<String> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut result = String::with_capacity(xml_bytes.len() / 8);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_depth = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;

        if fragment_prefix.is_none()
            && let Event::Start(element) = &event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }

        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML nesting is too deep".to_string())
                })?;
                if depth > MAX_TEXT_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML nesting exceeds the {MAX_TEXT_SCAN_DEPTH} depth limit"
                    )));
                }
                if text_depth.is_none()
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = Some(depth);
                } else if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_TEXT_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_TEXT_SCAN_NODES} elements"
                    )));
                }
                if let Some(character) =
                    word_special_character(&namespace, element.name(), &fragment_prefix)
                {
                    result.push(character);
                }
            },
            Event::Text(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(&unescaped);
            },
            Event::CData(text) if text_depth.is_some() => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                result.push_str(&decoded);
            },
            Event::GeneralRef(reference) if text_depth.is_some() => {
                result.push_str(&decode_xml_reference(&reference)?);
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_fragment_word_name(&namespace, element.name(), b"t", &fragment_prefix)
                {
                    text_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid Word XML nesting".to_string()))?;
            },
            Event::Eof if depth != 0 || text_depth.is_some() => {
                return Err(Error::InvalidFormat("unterminated Word XML".to_string()));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    result.shrink_to_fit();
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::{Path, PathBuf};

    /// Transitional (loose) and strict WordprocessingML namespace URIs.
    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const W_STRICT: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

    /// Differential parity: the tracker-driven path and the `NsReader`
    /// oracle must agree on the extracted text or fail with a
    /// byte-identical error string.
    fn assert_extract_parity(xml: &[u8]) {
        let tracker = extract_word_text(xml);
        let oracle = extract_word_text_nsreader_oracle(xml);
        match (tracker, oracle) {
            (Ok(tracker), Ok(oracle)) => {
                assert_eq!(tracker, oracle, "tracker/oracle extracted text diverges")
            },
            (Err(tracker), Err(oracle)) => assert_eq!(
                tracker.to_string(),
                oracle.to_string(),
                "tracker/oracle error strings diverge"
            ),
            (tracker, oracle) => {
                panic!("tracker/oracle outcome mismatch: {tracker:?} vs {oracle:?}")
            },
        }
    }

    #[test]
    fn tracker_path_matches_oracle_on_core_extraction() {
        let fixtures: Vec<(String, &str)> = vec![
            // Plain run text.
            (
                format!(r#"<w:p xmlns:w="{W}"><w:r><w:t>Hello</w:t></w:r></w:p>"#),
                "Hello",
            ),
            // All five special characters, as Empty and Start/End elements.
            (
                format!(
                    r#"<w:p xmlns:w="{W}"><w:r><w:tab/><w:br/><w:cr></w:cr><w:noBreakHyphen/><w:softHyphen/></w:r></w:p>"#
                ),
                "\t\n\n\u{2011}\u{00ad}",
            ),
            // Strict-namespace matching: strict `w:t` is extracted.
            (
                format!(r#"<w:p xmlns:w="{W_STRICT}"><w:r><w:t>S</w:t></w:r></w:p>"#),
                "S",
            ),
            // Bare fragment fallback: `w` never declared, the first unbound
            // Start fixes the fragment prefix and `w:t` matches through it.
            (r#"<w:p><w:r><w:t>Hi</w:t></w:r></w:p>"#.to_string(), "Hi"),
            // Unprefixed bare fragment: the first unbound Start has no
            // prefix, so unprefixed `t` matches the fallback.
            (r#"<p><r><t>Hi</t></r></p>"#.to_string(), "Hi"),
            // Prefix shadowing at depth: the rebound scope's `w:t` resolves
            // to the foreign URI and is skipped; the outer binding resumes.
            (
                format!(
                    r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:r xmlns:w="urn:foreign"><w:t>B</w:t></w:r><w:r><w:t>C</w:t></w:r></w:p>"#
                ),
                "AC",
            ),
            // Special characters under the shadowed prefix are skipped too.
            (
                format!(
                    r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t><w:tab/></w:r><w:r xmlns:w="urn:foreign"><w:tab/></w:r><w:r><w:t>C</w:t></w:r></w:p>"#
                ),
                "A\tC",
            ),
        ];
        for (xml, expected) in &fixtures {
            assert_extract_parity(xml.as_bytes());
            assert_eq!(
                extract_word_text(xml.as_bytes()).unwrap(),
                *expected,
                "unexpected extracted text for {xml}"
            );
        }
    }

    #[test]
    fn tracker_path_matches_oracle_on_subtle_namespace_fallbacks() {
        // These cases exercise interactions between the fragment-prefix
        // fallback and binding scopes whose pinned outcome is subtle (an
        // emptied default binding captures the fallback; an emptied prefix
        // binding matches a previously captured prefix); parity with the
        // oracle is the assertion, no hardcoded text.
        let fixtures: Vec<String> = vec![
            // Default-namespace redefinition and unset.
            format!(r#"<p xmlns="{W}"><t>A</t><d xmlns=""><t>B</t></d><t>C</t></p>"#),
            // Emptied prefix binding (`xmlns:w=""`) after the prefix
            // resolved properly: the emptied scope's names resolve
            // `Unknown("w")`, which the fragment fallback may still match.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:r xmlns:w=""><w:t>B</w:t></w:r></w:p>"#
            ),
            // An unbound root with a properly bound block inside: the bound
            // `w:t` wins through the namespace match, not the fallback.
            format!(r#"<p xmlns:w="{W}"><w:r><w:t>in</w:t></w:r><r><t>out</t></r></p>"#),
            // An `Empty` element carrying a declaration: its push and
            // deferred pop bracket the event itself.
            format!(
                r#"<w:p xmlns:w="{W}"><w:marker xmlns:w="urn:foreign"/><w:r><w:t>After</w:t></w:r></w:p>"#
            ),
            // An attribute *value* containing the substring `xmlns` must not
            // disturb the binding scan outcome.
            format!(r#"<w:p xmlns:w="{W}"><w:r w:id="xmlns:shadow"><w:t>A</w:t></w:r></w:p>"#),
            // CDATA and a general entity reference inside and outside `w:t`.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>x<![CDATA[<raw>]]>&amp;y</w:t></w:r><w:instrText><![CDATA[skip]]></w:instrText></w:p>"#
            ),
            // Mixed loose root and strict inner binding.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:s xmlns:w="{W_STRICT}"><w:t>B</w:t></w:s></w:p>"#
            ),
        ];
        for xml in &fixtures {
            assert_extract_parity(xml.as_bytes());
        }
    }

    #[test]
    fn tracker_path_matches_oracle_on_namespace_errors() {
        let xml_ns = "http://www.w3.org/XML/1998/namespace";
        let xmlns_ns = "http://www.w3.org/2000/xmlns/";
        let error_fixtures: Vec<String> = vec![
            // Declaring the `xmlns` prefix itself.
            format!(r#"<w:p xmlns:w="{W}" xmlns:xmlns="urn:example:x"><w:t>x</w:t></w:p>"#),
            // Binding `xml` to a foreign URI.
            format!(r#"<w:p xmlns:w="{W}" xmlns:xml="urn:example:x"><w:t>x</w:t></w:p>"#),
            // Binding another prefix to the reserved xml URI.
            format!(r#"<w:p xmlns:w="{W}" xmlns:q="{xml_ns}"><w:t>x</w:t></w:p>"#),
            // Binding a prefix to the reserved xmlns URI.
            format!(r#"<w:p xmlns:w="{W}" xmlns:q="{xmlns_ns}"><w:t>x</w:t></w:p>"#),
            // The same failure mid-stream after extractable text: the push
            // error preempts the event exactly where the `NsReader` read
            // error did.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:r xmlns:xml="urn:example:x"><w:t>B</w:t></w:r></w:p>"#
            ),
            // A namespace error on an `Empty` element.
            format!(
                r#"<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:tab xmlns:xmlns="urn:x"/></w:p>"#
            ),
        ];
        for xml in &error_fixtures {
            assert_extract_parity(xml.as_bytes());
            assert!(
                extract_word_text(xml.as_bytes()).is_err(),
                "expected an error for {xml}"
            );
        }
        // Rebinding `xml` to its reserved URI is a no-op, not an error.
        let benign = format!(r#"<w:p xmlns:w="{W}" xmlns:xml="{xml_ns}"><w:t>x</w:t></w:p>"#);
        assert_extract_parity(benign.as_bytes());
        assert_eq!(extract_word_text(benign.as_bytes()).unwrap(), "x");

        // Declaration-limit parity (`xmlns:w` accounts for one declaration
        // on the tag): 256 declarations pass, 257 fail identically.
        let declarations = |count: usize| {
            (0..count)
                .map(|index| format!(r#"xmlns:d{index}="urn:example:{index}""#))
                .collect::<Vec<_>>()
                .join(" ")
        };
        let within_limit = format!(
            r#"<w:p xmlns:w="{W}" {}><w:t>x</w:t></w:p>"#,
            declarations(255)
        );
        assert_extract_parity(within_limit.as_bytes());
        assert_eq!(extract_word_text(within_limit.as_bytes()).unwrap(), "x");
        let over_limit = format!(
            r#"<w:p xmlns:w="{W}" {}><w:t>x</w:t></w:p>"#,
            declarations(256)
        );
        assert_extract_parity(over_limit.as_bytes());
        assert!(extract_word_text(over_limit.as_bytes()).is_err());
    }

    #[test]
    fn tracker_path_matches_oracle_on_malformed_and_limited_xml() {
        // Unterminated document: the Eof structural check fires identically.
        let unterminated = format!(r#"<w:p xmlns:w="{W}"><w:r><w:t>unfinished"#);
        assert_extract_parity(unterminated.as_bytes());
        // Tokenizer error mid-tag.
        let broken = format!(r#"<w:p xmlns:w="{W}"><w:t a=">#</w:p>"#);
        assert_extract_parity(broken.as_bytes());
        // Depth limit: 129 nested elements exceed MAX_TEXT_SCAN_DEPTH = 128.
        let deep = format!(
            r#"<w:p xmlns:w="{W}">{}x{}</w:p>"#,
            "<w:a>".repeat(MAX_TEXT_SCAN_DEPTH + 1),
            "</w:a>".repeat(MAX_TEXT_SCAN_DEPTH + 1),
        );
        assert_extract_parity(deep.as_bytes());
        let error = extract_word_text(deep.as_bytes()).unwrap_err();
        assert_eq!(
            error.to_string(),
            extract_word_text_nsreader_oracle(deep.as_bytes())
                .unwrap_err()
                .to_string()
        );
        assert!(error.to_string().contains("depth limit"));
    }

    #[test]
    fn tracker_path_matches_oracle_on_docx_corpus() {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        let mut files = Vec::new();
        collect_docx_corpus(&root, &mut files);
        files.sort();
        assert!(!files.is_empty(), "no .docx corpus fixtures discovered");
        let mut compared = 0usize;
        for path in &files {
            let Some(document_xml) = docx_document_xml(path) else {
                continue;
            };
            assert_extract_parity(&document_xml);
            compared += 1;
        }
        assert!(
            compared > 0,
            "no .docx corpus fixtures yielded word/document.xml"
        );
        eprintln!("corpus parity compared over {compared} document parts");
    }

    fn collect_docx_corpus(directory: &Path, files: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_docx_corpus(&path, files);
            } else if path
                .extension()
                .is_some_and(|extension| extension == "docx")
            {
                files.push(path);
            }
        }
    }

    fn docx_document_xml(path: &Path) -> Option<Vec<u8>> {
        let bytes = std::fs::read(path).ok()?;
        let reader = soapberry_zip::office::ArchiveReader::new(&bytes).ok()?;
        reader.read("word/document.xml").ok()
    }
}
