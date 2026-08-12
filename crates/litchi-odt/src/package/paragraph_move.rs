//! Exact-fragment movement for a deliberately narrow plain-paragraph closure.

#![deny(
    clippy::expect_used,
    clippy::panic,
    clippy::todo,
    clippy::unimplemented,
    clippy::unwrap_used
)]

use crate::{
    constants::ODF_CONTENT,
    protection::Policy,
    transaction::{EnvelopeKind, Snapshot},
};
use litchi_core::{Error, Result};
use litchi_odf_common::core::{AuthoredXmlFragment, XmlSourcePart, XmlSplicePublication};
use litchi_odf_common::package::{
    MAX_CONTENT_REPLACEMENT_BYTES, raw_identical_members, replace_content_xml_spliced,
};
use quick_xml::{
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};
use std::ops::Range;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MCE_NS: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_PARAGRAPHS: usize = 4_096;
const MAX_EVENTS: usize = 1_048_576;
const MAX_DEPTH: usize = 64;

pub(crate) fn move_plain_paragraph(source: &Snapshot, from: usize, to: usize) -> Result<Snapshot> {
    match source.envelope_kind()? {
        EnvelopeKind::Plain => {},
        EnvelopeKind::Signed => return unsupported("signed packages"),
        EnvelopeKind::Encrypted => return unsupported("encrypted packages"),
    }
    let document = source.document()?;
    let content = document.transaction_content_xml();
    if content.len() > MAX_CONTENT_REPLACEMENT_BYTES {
        return invalid("content.xml exceeds the paragraph-move limit");
    }
    if document.protection()? != Policy::default() {
        return unsupported("protected documents");
    }
    if document
        .document_scripts()?
        .is_some_and(|scripts| !scripts.scripts.is_empty() || !scripts.event_listeners.is_empty())
        || !document.script_resources()?.is_empty()
    {
        return unsupported("documents containing scripts");
    }

    let paragraphs = scan_plain_paragraphs(content)?;
    if from >= paragraphs.len() || to >= paragraphs.len() {
        return invalid("plain paragraph move position is out of bounds");
    }
    if from == to {
        return Ok(source.clone());
    }

    let order = paragraph_order(paragraphs.len(), from, to)?;
    let candidate = reorder_exact_fragments(content, &paragraphs, &order)?;
    let source_part = XmlSourcePart::load(document.transaction_package(), ODF_CONTENT)?;
    let proof_source = source_part.clone();
    let mut publication = XmlSplicePublication::new(source_part);
    for (slot, source_index) in order.into_iter().enumerate() {
        if slot == source_index {
            continue;
        }
        let slot_range = paragraphs
            .get(slot)
            .ok_or_else(|| invalid_error("invalid ODT paragraph move slot"))?
            .clone();
        let expected = content
            .as_bytes()
            .get(slot_range.clone())
            .ok_or_else(|| invalid_error("invalid ODT paragraph move proof range"))?;
        let replacement = content
            .as_bytes()
            .get(
                paragraphs
                    .get(source_index)
                    .ok_or_else(|| invalid_error("invalid ODT paragraph move source"))?
                    .clone(),
            )
            .ok_or_else(|| invalid_error("invalid ODT paragraph move fragment range"))?
            .to_vec();
        let proof = proof_source.checked_range(slot_range, expected)?;
        publication.replace(proof, AuthoredXmlFragment::markup(replacement)?)?;
    }
    let target_bytes =
        replace_content_xml_spliced(document.transaction_package(), &candidate, publication)?;
    let target = Snapshot::from_bytes(target_bytes)?;
    let target_document = target.document()?;
    require_raw_untouched_members(
        document.transaction_package(),
        target_document.transaction_package(),
    )?;
    Ok(target)
}

fn scan_plain_paragraphs(xml: &str) -> Result<Vec<Range<usize>>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut body_depth = None;
    let mut text_depth = None;
    let mut paragraph = None::<(usize, usize)>;
    let mut paragraphs = Vec::new();
    let mut saw_root = false;
    let mut saw_body = false;
    let mut saw_text = false;

    loop {
        let event_start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODT content.xml: {error}")))?;
        let mce = is_bound(&namespace, MCE_NS);
        let office = is_bound(&namespace, OFFICE_NS);
        let text = is_bound(&namespace, TEXT_NS);
        let event = event.into_owned();
        if match &event {
            Event::Start(element) | Event::Empty(element) => {
                has_mce_attribute(&mut reader, element)?
            },
            _ => false,
        } {
            return unsupported("markup-compatibility attributes");
        }
        let event_end = position(&reader)?;
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid_error("ODT paragraph move event count overflow"))?;
        if events > MAX_EVENTS {
            return invalid("ODT paragraph move event limit exceeded");
        }
        if mce {
            return unsupported("markup-compatibility content");
        }

        match event {
            Event::Start(element) => {
                let local = element.local_name();
                if depth == 0 {
                    if saw_root || !office || local.as_ref() != b"document-content" {
                        return unsupported("noncanonical content.xml root ownership");
                    }
                    saw_root = true;
                } else if office && local.as_ref() == b"scripts" {
                    return unsupported("documents containing scripts");
                }
                if paragraph.is_some() {
                    return unsupported("paragraphs containing nested markup");
                }
                if office && local.as_ref() == b"body" {
                    if depth != 1 || saw_body || body_depth.is_some() {
                        return unsupported("multiple or nested office:body elements");
                    }
                    saw_body = true;
                    body_depth = Some(depth);
                } else if office && local.as_ref() == b"text" {
                    if body_depth.and_then(|value| value.checked_add(1)) != Some(depth)
                        || saw_text
                        || text_depth.is_some()
                    {
                        return unsupported("noncanonical office:text ownership");
                    }
                    if element.attributes().with_checks(true).next().is_some() {
                        return unsupported("office:text attributes");
                    }
                    saw_text = true;
                    text_depth = Some(depth);
                } else if body_depth.is_some() && text_depth.is_none() {
                    return unsupported("non-text office:body children");
                } else if let Some(owner_depth) = text_depth {
                    if depth != owner_depth.saturating_add(1) || !text || local.as_ref() != b"p" {
                        return unsupported("non-plain office:text children");
                    }
                    if element.attributes().with_checks(true).next().is_some() {
                        return unsupported("plain paragraph attributes");
                    }
                    paragraph = Some((depth, event_start));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ODT paragraph move XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return invalid("ODT paragraph move XML depth limit exceeded");
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if paragraph.is_some() {
                    return unsupported("paragraphs containing nested markup");
                }
                if depth == 0 {
                    return unsupported("empty or noncanonical content.xml root ownership");
                }
                if office && matches!(local.as_ref(), b"body" | b"text") {
                    return unsupported("empty or duplicate body/text ownership");
                } else if office && local.as_ref() == b"scripts" {
                    // The standard empty placeholder carries no executable payload.
                    if depth != 1
                        || body_depth.is_some()
                        || text_depth.is_some()
                        || element.attributes().with_checks(true).next().is_some()
                    {
                        return unsupported("scripts inside the document body");
                    }
                } else if body_depth.is_some() && text_depth.is_none() {
                    return unsupported("non-text office:body children");
                } else if let Some(owner_depth) = text_depth {
                    if depth != owner_depth.saturating_add(1) || !text || local.as_ref() != b"p" {
                        return unsupported("non-plain office:text children");
                    }
                    if element.attributes().with_checks(true).next().is_some() {
                        return unsupported("plain paragraph attributes");
                    }
                    push_paragraph(&mut paragraphs, event_start..event_end)?;
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("ODT paragraph move XML depth underflow"))?;
                let local = element.local_name();
                if let Some((paragraph_depth, start)) = paragraph {
                    if depth == paragraph_depth {
                        if !text || local.as_ref() != b"p" {
                            return invalid("ODT plain paragraph has a mismatched close tag");
                        }
                        push_paragraph(&mut paragraphs, start..event_end)?;
                        paragraph = None;
                    }
                } else if text_depth == Some(depth) {
                    if !office || local.as_ref() != b"text" {
                        return invalid("ODT office:text has a mismatched close tag");
                    }
                    text_depth = None;
                } else if body_depth == Some(depth) {
                    if !office || local.as_ref() != b"body" {
                        return invalid("ODT office:body has a mismatched close tag");
                    }
                    body_depth = None;
                }
            },
            Event::Text(value) => {
                if paragraph.is_none()
                    && text_depth.is_some()
                    && value
                        .as_ref()
                        .iter()
                        .any(|byte| !byte.is_ascii_whitespace())
                {
                    return unsupported("text outside direct paragraphs");
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if paragraph.is_some() => {},
            Event::CData(_) | Event::GeneralRef(_) if text_depth.is_some() => {
                return unsupported("non-whitespace office:text content");
            },
            Event::Decl(_) | Event::Comment(_) | Event::PI(_)
                if paragraph.is_some() || text_depth.is_some() =>
            {
                return unsupported("opaque markup in office:text");
            },
            Event::DocType(_) => return unsupported("documents containing a doctype"),
            Event::Eof => break,
            Event::Decl(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if depth != 0 || paragraph.is_some() || text_depth.is_some() || body_depth.is_some() {
        return invalid("ODT paragraph move XML is truncated");
    }
    if !saw_root || !saw_body || !saw_text {
        return unsupported("missing content.xml/office:body/office:text ownership");
    }
    Ok(paragraphs)
}

fn reorder_exact_fragments(
    xml: &str,
    paragraphs: &[Range<usize>],
    order: &[usize],
) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT exact paragraph move output",
            source,
        })?;
    let mut cursor = 0usize;
    for (slot, source_index) in order.iter().copied().enumerate() {
        let slot_range = paragraphs
            .get(slot)
            .ok_or_else(|| invalid_error("invalid ODT paragraph move slot"))?;
        let source_range = paragraphs
            .get(source_index)
            .ok_or_else(|| invalid_error("invalid ODT paragraph move source"))?;
        output.push_str(
            xml.get(cursor..slot_range.start)
                .ok_or_else(|| invalid_error("invalid ODT paragraph separator range"))?,
        );
        output.push_str(
            xml.get(source_range.clone())
                .ok_or_else(|| invalid_error("invalid ODT paragraph fragment range"))?,
        );
        cursor = slot_range.end;
    }
    output.push_str(
        xml.get(cursor..)
            .ok_or_else(|| invalid_error("invalid ODT paragraph suffix range"))?,
    );
    if output.len() != xml.len() {
        return invalid("ODT paragraph move changed content.xml length");
    }
    Ok(output)
}

fn paragraph_order(length: usize, from: usize, to: usize) -> Result<Vec<usize>> {
    let mut order = Vec::new();
    order
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "ODT plain paragraph move order",
            source,
        })?;
    order.extend(0..length);
    let moved = order.remove(from);
    order.insert(to, moved);
    Ok(order)
}

fn require_raw_untouched_members(
    source: &crate::core::OwnedPackage,
    target: &crate::core::OwnedPackage,
) -> Result<()> {
    let identical = raw_identical_members(source.as_bytes(), target.as_bytes())
        .ok_or_else(|| invalid_error("ODT paragraph move cannot audit raw ZIP members"))?;
    let mut source_paths = source.package()?.files()?;
    let mut target_paths = target.package()?.files()?;
    source_paths.sort();
    target_paths.sort();
    if source_paths != target_paths {
        return unsupported("packages whose member set changes during publication");
    }
    for path in source_paths {
        if path != ODF_CONTENT && !identical.contains(&path) {
            return unsupported("packages that cannot raw-preserve untouched members");
        }
    }
    Ok(())
}

fn push_paragraph(paragraphs: &mut Vec<Range<usize>>, range: Range<usize>) -> Result<()> {
    if paragraphs.len() >= MAX_PARAGRAPHS {
        return invalid("ODT paragraph move paragraph limit exceeded");
    }
    paragraphs
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "ODT plain paragraph move index",
            source,
        })?;
    paragraphs.push(range);
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid_error("ODT paragraph move XML position overflow"))
}

fn is_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected)
}

fn has_mce_attribute(reader: &mut NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODT content.xml attribute: {error}"))
        })?;
        if is_bound(
            &reader.resolver_mut().resolve_attribute(attribute.key).0,
            MCE_NS,
        ) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn unsupported<T>(what: &str) -> Result<T> {
    Err(Error::Unsupported(format!(
        "ODT exact plain paragraph move refuses {what}"
    )))
}
