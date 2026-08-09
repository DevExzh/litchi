#![expect(
    clippy::redundant_locals,
    reason = "the local rebinding narrows the parser state"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_same,
    reason = "the validated binding intentionally replaces its fallible precursor"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
//! Package-facing discovery and section integration for `footnoteColumns`.

use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::codec::Context;
use super::transaction::Snapshot;

const MAX_PART_NODES: usize = 1_000_000;
const MAX_PART_DEPTH: usize = 128;
const MC_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Parse every `w:sectPr` found in a Word main-document part.
///
/// The source part is scanned as authored bytes so an ignorable Word 2012
/// child is not discarded before the focused owner can snapshot it. Namespace
/// context inherited from the document root is retained out-of-band by each
/// snapshot, so detached fragments remain byte-for-byte authored until edited.
///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
pub fn parse_part(part: &dyn Part) -> Result<Vec<Snapshot>> {
    let xml = part.blob();
    let xml = xml;
    let mut snapshots = Vec::new();
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut ignorable = Vec::<Option<String>>::new();
    let mut capture = None;

    loop {
        let event_start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let event_end = position(&reader)?;
        let resolver = reader.resolver().clone();

        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormat("document XML node counter overflow".into()))?;
            if nodes > MAX_PART_NODES {
                return Err(Error::InvalidFormat(format!(
                    "document XML exceeds {MAX_PART_NODES} elements"
                )));
            }
        }

        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("document XML depth overflow".into()))?;
                if depth > MAX_PART_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "document XML nesting exceeds {MAX_PART_DEPTH}"
                    )));
                }
                let effective_ignorable = direct_ignorable(&element, &resolver, decoder)?
                    .or_else(|| ignorable.last().cloned().flatten());
                ignorable.push(effective_ignorable.clone());
                if is_wordprocessing_namespace(&resolver.resolve_element(element.name()).0)
                    && element.local_name().as_ref() == b"sectPr"
                {
                    if capture.is_some() {
                        return Err(Error::InvalidFormat(
                            "document XML contains nested sectPr elements".into(),
                        ));
                    }
                    capture = Some((
                        event_start,
                        depth,
                        Context::from_resolver(&resolver, effective_ignorable)?,
                    ));
                }
            },
            Event::Empty(element) => {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("document XML depth overflow".into()))?;
                if child_depth > MAX_PART_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "document XML nesting exceeds {MAX_PART_DEPTH}"
                    )));
                }
                let effective_ignorable = direct_ignorable(&element, &resolver, decoder)?
                    .or_else(|| ignorable.last().cloned().flatten());
                if is_wordprocessing_namespace(&resolver.resolve_element(element.name()).0)
                    && element.local_name().as_ref() == b"sectPr"
                {
                    let context = Context::from_resolver(&resolver, effective_ignorable)?;
                    let bytes = xml.get(event_start..event_end).ok_or_else(|| {
                        Error::InvalidFormat("section XML range is outside its part".into())
                    })?;
                    snapshots.push(Snapshot::from_xml_with_context(bytes.to_vec(), context)?);
                }
            },
            Event::End(_) => {
                let Some((start, target_depth, context)) = capture.take() else {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("document XML has invalid nesting".into())
                    })?;
                    ignorable.pop();
                    continue;
                };
                if target_depth == depth {
                    let bytes = xml.get(start..event_end).ok_or_else(|| {
                        Error::InvalidFormat("section XML range is outside its part".into())
                    })?;
                    snapshots.push(Snapshot::from_xml_with_context(bytes.to_vec(), context)?);
                } else {
                    capture = Some((start, target_depth, context));
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("document XML has invalid nesting".into())
                })?;
                ignorable.pop();
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

    if depth != 0 || capture.is_some() || !ignorable.is_empty() {
        return Err(Error::InvalidFormat(
            "document XML has incomplete element nesting".into(),
        ));
    }
    Ok(snapshots)
}

fn direct_ignorable(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"Ignorable" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if *value == *MC_NAMESPACE
        ) && !matches!(namespace, ResolveResult::Unknown(value) if value.as_slice() == b"mc")
        {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "document element has duplicate mc:Ignorable attributes".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source_error| Error::InvalidFormat("document XML offset overflow".into()))
}
