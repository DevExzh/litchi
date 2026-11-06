#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
//! Package-facing discovery and section integration for `footnoteColumns`.

use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use litchi_core::xml::ReaderOrigin;
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::codec::Context;
use super::transaction::Snapshot;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

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
    section_ranges(xml)?
        .into_iter()
        .map(|(start, end, context)| {
            Snapshot::from_xml_with_context(section_bytes(xml, start, end)?.to_vec(), context)
        })
        .collect()
}

fn section_ranges(xml: &[u8]) -> Result<Vec<(usize, usize, Context)>> {
    let mut snapshots = Vec::new();
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut ignorable = Vec::<Option<String>>::new();
    let mut capture = None;

    loop {
        let event_start = position(&reader, origin)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let event_end = position(&reader, origin)?;
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
                    snapshots.push((event_start, event_end, context));
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
                    snapshots.push((start, event_end, context));
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
    for attribute in element.checked_attributes() {
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

/// Build detached section views from selected source owners, retaining the
/// inherited namespace and Ignorable context that a visibility pass removes.
pub(crate) fn active_sections(part: &dyn Part) -> Result<Vec<crate::section::Section>> {
    use litchi_ooxml_common::mce;
    let xml = part.blob();
    let ranges = section_ranges(xml)?;
    let offsets = ranges
        .iter()
        .map(|(start, _, _)| {
            u32::try_from(*start)
                .map_err(|_| Error::InvalidFormat("section offset exceeds u32".into()))
        })
        .collect::<Result<Vec<_>>>()?;
    let active = mce::active_offsets(
        xml,
        &offsets,
        &crate::parts::document_part::visible_document_capabilities(),
        &mce::OffsetLimits::default(),
    )?
    .into_iter()
    .collect::<std::collections::HashSet<_>>();
    let mut sections = Vec::new();
    for ((start, end, context), offset) in ranges.into_iter().zip(offsets) {
        if !active.contains(&offset) {
            continue;
        }
        // Validate the source extension before giving the section a standalone
        // namespace envelope. This must not invent a missing Ignorable rule.
        let _ = Snapshot::from_xml_with_context(
            section_bytes(xml, start, end)?.to_vec(),
            context.clone(),
        )?;
        let fragment =
            mce::self_contained_fragment(xml, start, end - start, &mce::Limits::default())?;
        let fragment = with_inherited_ignorable(fragment.as_ref(), context.ignorable())?;
        let visible = mce::process_markup_compatibility(
            &fragment,
            &crate::parts::document_part::visible_document_capabilities(),
            &mce::Limits::default(),
        )?;
        // Preserve the authored fragment when processing only normalizes
        // namespace spelling. When nested compatibility changes visibility,
        // use that selected projection and restore its inherited requirement.
        let projected = visible.report.alternate_content_count != 0
            || visible.report.ignored_elements != 0
            || visible.report.unwrapped_elements != 0
            || visible.report.ignored_attributes != 0;
        let section = if projected {
            with_inherited_ignorable(visible.xml.as_ref(), context.ignorable())?
        } else {
            fragment
        };
        sections.push(crate::section::Section::from_xml_bytes(section)?);
    }
    Ok(sections)
}

fn with_inherited_ignorable(xml: &[u8], ignorable: Option<&str>) -> Result<Vec<u8>> {
    let Some(ignorable) = ignorable else {
        return Ok(xml.to_vec());
    };
    let mut reader = NsReader::from_reader(xml);
    let event = reader
        .read_event()
        .map_err(|error| Error::Xml(error.to_string()))?;
    let (Event::Start(root) | Event::Empty(root)) = event else {
        return Err(Error::InvalidFormat(
            "section fragment has no element root".into(),
        ));
    };
    if direct_ignorable(&root, reader.resolver(), reader.decoder())?.is_some() {
        return Ok(xml.to_vec());
    }
    let mut prefix = "mc".to_owned();
    let mut suffix = 0usize;
    loop {
        let name = format!("{prefix}:Ignorable");
        match reader
            .resolver()
            .resolve_attribute(quick_xml::name::QName(name.as_bytes()))
            .0
        {
            ResolveResult::Bound(Namespace(ns)) if ns != MC_NAMESPACE => {
                suffix += 1;
                prefix = format!("mc{suffix}");
            },
            _ => break,
        }
    }
    let name = format!("{prefix}:Ignorable");
    let bound = matches!(reader.resolver().resolve_attribute(quick_xml::name::QName(name.as_bytes())).0,ResolveResult::Bound(Namespace(ns)) if ns==MC_NAMESPACE);
    let mut addition = format!(
        " {prefix}:Ignorable=\"{}\"",
        quick_xml::escape::escape(ignorable)
    );
    if !bound {
        addition.push_str(&format!(
            " xmlns:{prefix}=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
        ));
    }
    let mut at = usize::try_from(reader.buffer_position())
        .map_err(|_| Error::InvalidFormat("section root offset exceeds usize".into()))?
        - 1;
    if xml.get(at.wrapping_sub(1)) == Some(&b'/') {
        at -= 1;
    }
    let output_len = xml
        .len()
        .checked_add(addition.len())
        .ok_or_else(|| Error::InvalidFormat("section context size overflows".into()))?;
    if output_len > super::validation::MAX_XML_BYTES {
        return Err(Error::InvalidFormat(
            "section context exceeds XML byte limit".into(),
        ));
    }
    let mut result = xml.to_vec();
    result.splice(at..at, addition.bytes());
    Ok(result)
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| Error::InvalidFormat("document XML offset overflow".into()))
}

fn section_bytes(xml: &[u8], start: usize, end: usize) -> Result<&[u8]> {
    xml.get(start..end)
        .ok_or_else(|| Error::InvalidFormat("section source range escapes document XML".into()))
}
