//! Namespace-aware, range-preserving XML edits for the transaction root.

mod empty_slides;

use std::collections::{BTreeSet, HashMap, HashSet};
use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::{DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE};
use quick_xml::Reader;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Slide, invalid};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

pub(super) const MAX_XML_DEPTH: usize = 256;
const MAX_XML_NODES: usize = 1_000_000;

/// The compacted bytes and the conservative witness that the document root
/// survived byte-for-byte. The witness is transient and is not retained in a
/// transaction or snapshot.
#[derive(Debug)]
pub(crate) struct CompactedSlide {
    pub(crate) xml: Vec<u8>,
    pub(crate) root_unchanged: bool,
}

/// Re-emit a changed slide without inter-element formatting.
///
/// A leading UTF-8 byte-order mark, which the reader consumes before its
/// first event, is carried to the output unchanged: it is a byte the producer
/// wrote, not formatting.
pub(crate) fn compact_changed_slide_xml(source: &[u8]) -> Result<CompactedSlide> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let origin = ReaderOrigin::of(source);
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation compact slide XML",
            source,
        })?;
    output.extend_from_slice(&source[..ReaderOrigin::of(source).skipped()]);
    let mut preserve_space = Vec::new();
    let mut pending_whitespace = Vec::new();
    let mut text_run_has_content = false;
    let mut roots = 0usize;
    let mut source_root_start = None;
    let mut source_root_end = None;
    let mut output_root_start = None;
    let mut output_root_end = None;
    let mut outside_non_text_exact = true;
    loop {
        let source_event_start = origin.offset(reader.buffer_position());
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let source_event_end = origin.offset(reader.buffer_position());
        match event {
            Event::Start(element) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                let is_root = preserve_space.is_empty();
                if is_root {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| invalid("opened-presentation slide root count overflow"))?;
                    source_root_start = source_event_start;
                    output_root_start = Some(output.len());
                }
                let inherited = preserve_space.last().copied().unwrap_or(false);
                preserve_space.push(element_preserves_space(
                    &element,
                    reader.decoder(),
                    inherited,
                )?);
                write_compact_start(&mut output, &element, reader.decoder(), false)?;
            },
            Event::Empty(element) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                let is_root = preserve_space.is_empty();
                if is_root {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| invalid("opened-presentation slide root count overflow"))?;
                    source_root_start = source_event_start;
                    output_root_start = Some(output.len());
                }
                let inherited = preserve_space.last().copied().unwrap_or(false);
                let _preserve = element_preserves_space(&element, reader.decoder(), inherited)?;
                write_compact_start(&mut output, &element, reader.decoder(), true)?;
                if is_root {
                    source_root_end = origin.offset(reader.buffer_position());
                    output_root_end = Some(output.len());
                }
            },
            Event::End(element) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                let is_root = preserve_space.len() == 1;
                preserve_space.pop().ok_or_else(|| {
                    invalid("opened-presentation slide has an unexpected end element")
                })?;
                output.extend_from_slice(b"</");
                output.extend_from_slice(element.name().as_ref());
                output.push(b'>');
                if is_root {
                    source_root_end = origin.offset(reader.buffer_position());
                    output_root_end = Some(output.len());
                }
            },
            Event::Text(text) => {
                let bytes = text.as_ref();
                if preserve_space.is_empty() {
                    if bytes.iter().all(u8::is_ascii_whitespace) {
                        continue;
                    }
                    return Err(invalid(
                        "opened-presentation slide has text outside its root",
                    ));
                }
                if bytes.iter().all(u8::is_ascii_whitespace)
                    && !preserve_space.last().copied().unwrap_or(false)
                {
                    if text_run_has_content {
                        output.extend_from_slice(bytes);
                    } else {
                        pending_whitespace.extend_from_slice(bytes);
                    }
                } else {
                    output.extend_from_slice(&pending_whitespace);
                    pending_whitespace.clear();
                    output.extend_from_slice(bytes);
                    text_run_has_content = true;
                }
            },
            Event::CData(data) => {
                if preserve_space.is_empty() {
                    return Err(invalid(
                        "opened-presentation slide has CDATA outside its root",
                    ));
                }
                output.extend_from_slice(&pending_whitespace);
                pending_whitespace.clear();
                output.extend_from_slice(b"<![CDATA[");
                output.extend_from_slice(data.as_ref());
                output.extend_from_slice(b"]]>");
                text_run_has_content = true;
            },
            Event::GeneralRef(reference) => {
                if preserve_space.is_empty() {
                    return Err(invalid(
                        "opened-presentation slide has an entity outside its root",
                    ));
                }
                output.extend_from_slice(&pending_whitespace);
                pending_whitespace.clear();
                output.push(b'&');
                output.extend_from_slice(reference.as_ref());
                output.push(b';');
                text_run_has_content = true;
            },
            Event::Decl(declaration) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                let output_event_start = output.len();
                output.extend_from_slice(b"<?");
                output.extend_from_slice(declaration.as_ref());
                output.extend_from_slice(b"?>");
                if preserve_space.is_empty() {
                    outside_non_text_exact &= exact_event_bytes(
                        source,
                        &output,
                        source_event_start,
                        source_event_end,
                        output_event_start,
                        output.len(),
                    );
                }
            },
            Event::PI(instruction) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                let output_event_start = output.len();
                output.extend_from_slice(b"<?");
                output.extend_from_slice(instruction.as_ref());
                output.extend_from_slice(b"?>");
                if preserve_space.is_empty() {
                    outside_non_text_exact &= exact_event_bytes(
                        source,
                        &output,
                        source_event_start,
                        source_event_end,
                        output_event_start,
                        output.len(),
                    );
                }
            },
            Event::Comment(comment) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                let output_event_start = output.len();
                output.extend_from_slice(b"<!--");
                output.extend_from_slice(comment.as_ref());
                output.extend_from_slice(b"-->");
                if preserve_space.is_empty() {
                    outside_non_text_exact &= exact_event_bytes(
                        source,
                        &output,
                        source_event_start,
                        source_event_end,
                        output_event_start,
                        output.len(),
                    );
                }
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "opened-presentation slide document types are not publishable",
                ));
            },
            Event::Eof => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                break;
            },
        }
    }
    if !preserve_space.is_empty() || roots != 1 {
        return Err(invalid(
            "opened-presentation slide must contain exactly one closed root",
        ));
    }
    let root_unchanged = match (
        source_root_start,
        source_root_end,
        output_root_start,
        output_root_end,
    ) {
        (Some(source_start), Some(source_end), Some(output_start), Some(output_end)) => source
            .get(source_start..source_end)
            .zip(output.get(output_start..output_end))
            .is_some_and(|(source_root, output_root)| source_root == output_root),
        _ => false,
    };
    Ok(CompactedSlide {
        xml: output,
        root_unchanged: root_unchanged && outside_non_text_exact,
    })
}

fn exact_event_bytes(
    source: &[u8],
    output: &[u8],
    source_start: Option<usize>,
    source_end: Option<usize>,
    output_start: usize,
    output_end: usize,
) -> bool {
    let (Some(source_start), Some(source_end)) = (source_start, source_end) else {
        return false;
    };
    source
        .get(source_start..source_end)
        .zip(output.get(output_start..output_end))
        .is_some_and(|(source_event, output_event)| source_event == output_event)
}

fn finish_compact_text_run(pending_whitespace: &mut Vec<u8>, has_content: &mut bool) {
    pending_whitespace.clear();
    *has_content = false;
}

fn element_preserves_space(
    element: &BytesStart<'_>,
    decoder: Decoder,
    inherited: bool,
) -> Result<bool> {
    let mut preserve = inherited;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() != b"xml:space" {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        preserve = match value.as_ref() {
            "preserve" => true,
            "default" => false,
            _ => {
                return Err(invalid(
                    "opened-presentation slide xml:space must be default or preserve",
                ));
            },
        };
    }
    Ok(preserve)
}

fn write_compact_start(
    output: &mut Vec<u8>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    empty: bool,
) -> Result<()> {
    output.push(b'<');
    output.extend_from_slice(element.name().as_ref());
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        output.push(b' ');
        output.extend_from_slice(attribute.key.as_ref());
        output.extend_from_slice(b"=\"");
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        // Literal attribute whitespace is normalized to spaces on reparse.
        // Keep whitespace that originated in character references semantic.
        let escaped = quick_xml::escape::escape(value.as_ref());
        let mut start = 0;
        for (index, byte) in escaped.bytes().enumerate() {
            let reference: &[u8] = match byte {
                b'\t' => b"&#9;",
                b'\n' => b"&#10;",
                b'\r' => b"&#13;",
                _ => continue,
            };
            output.extend_from_slice(&escaped.as_bytes()[start..index]);
            output.extend_from_slice(reference);
            start = index + 1;
        }
        output.extend_from_slice(&escaped.as_bytes()[start..]);
        output.push(b'"');
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
    Ok(())
}

#[derive(Debug)]
struct SlideIdElement {
    id: u32,
    relationship_id: String,
    span: Range<usize>,
    element_name: Option<Vec<u8>>,
    element_namespace: Option<Vec<u8>>,
    relationship_attribute_name: Option<Vec<u8>>,
    relationship_namespace: Option<Vec<u8>>,
}

#[derive(Debug)]
struct OpenElement {
    id: u32,
    relationship_id: String,
    start: usize,
    depth: usize,
}

#[derive(Debug)]
struct TextElement {
    span: Range<usize>,
    empty_name: Option<Vec<u8>>,
    /// Whether the slot is a field's (`a:fld`) cached result.
    field: bool,
}

struct ShapeTextEdit<'a> {
    index: usize,
    shape: Range<usize>,
    text: &'a str,
    changed: bool,
}

pub(crate) fn reorder_slides(xml: &[u8], current: &[Slide], ordered: &[u32]) -> Result<Vec<u8>> {
    let mut current_bindings = Vec::new();
    current_bindings
        .try_reserve_exact(current.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide bindings",
            source,
        })?;
    for slide in current {
        current_bindings.push((slide.id, slide.relationship_id.as_str()));
    }
    let mut ordered_bindings = Vec::new();
    ordered_bindings
        .try_reserve(ordered.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation ordered slide bindings",
            source,
        })?;
    for id in ordered {
        let relationship_id = current
            .iter()
            .find(|slide| slide.id == *id)
            .map_or("", |slide| slide.relationship_id.as_str());
        ordered_bindings.push((*id, relationship_id));
    }
    reorder_slide_bindings(xml, &current_bindings, &ordered_bindings, None)
}

/// Validate and reorder a raw `p:sldIdLst` against exact `(id, r:id)` bindings.
///
/// The entry spans are permuted as opaque byte ranges. Every byte before the
/// first entry, between entry slots, and after the final entry remains at its
/// original position; IDs and relationship IDs are never regenerated.
pub(crate) fn reorder_slide_bindings(
    xml: &[u8],
    current: &[(u32, &str)],
    ordered: &[(u32, &str)],
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<Vec<u8>> {
    if ordered.len() != current.len() {
        return Err(invalid(
            "opened-presentation slide order is not a complete permutation",
        ));
    }
    let elements = slide_id_elements(xml, false, context)?;
    validate_slide_elements(&elements, current, context)?;
    let mut selected = Vec::new();
    selected
        .try_reserve(ordered.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide permutation",
            source,
        })?;
    let mut seen_ids = HashSet::new();
    let mut seen_relationships = HashSet::new();
    seen_ids
        .try_reserve(ordered.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation ordered slide IDs",
            source,
        })?;
    seen_relationships
        .try_reserve(ordered.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation ordered slide relationship IDs",
            source,
        })?;
    let mut indexes = HashMap::new();
    indexes
        .try_reserve(current.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide binding index",
            source,
        })?;
    for (index, binding) in current.iter().enumerate() {
        if index % 256 == 0 {
            check_execution_context(context)?;
        }
        if indexes.insert(*binding, index).is_some() {
            return Err(invalid(
                "opened-presentation slide-order bindings are ambiguous",
            ));
        }
    }
    for (index, (id, relationship_id)) in ordered.iter().enumerate() {
        if index % 256 == 0 {
            check_execution_context(context)?;
        }
        if !seen_ids.insert(*id) || !seen_relationships.insert(*relationship_id) {
            return Err(invalid(
                "opened-presentation slide order repeats an identity",
            ));
        }
        let index = indexes
            .get(&(*id, *relationship_id))
            .copied()
            .ok_or_else(|| {
                invalid("opened-presentation slide order references an unknown binding")
            })?;
        selected.push(index);
    }
    if selected
        .iter()
        .enumerate()
        .all(|(left, right)| left == *right)
    {
        check_execution_context(context)?;
        return clone_xml_bytes(xml, "opened-presentation unchanged XML");
    }
    let first = elements
        .first()
        .ok_or_else(|| invalid("opened-presentation slide list is empty"))?;
    let last = elements
        .last()
        .ok_or_else(|| invalid("opened-presentation slide list is empty"))?;
    check_execution_context(context)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation reordered XML",
            source,
        })?;
    output.extend_from_slice(span_bytes(xml, 0..first.span.start)?);
    for (position, source_index) in selected.into_iter().enumerate() {
        if position % 256 == 0 {
            check_execution_context(context)?;
        }
        let source = elements
            .get(source_index)
            .ok_or_else(|| invalid("opened-presentation slide order lost a binding"))?;
        output.extend_from_slice(span_bytes(xml, source.span.clone())?);
        if let (Some(current), Some(next)) = (elements.get(position), elements.get(position + 1)) {
            output.extend_from_slice(span_bytes(xml, current.span.end..next.span.start)?);
        }
    }
    output.extend_from_slice(span_bytes(xml, last.span.end..xml.len())?);
    Ok(output)
}

/// Validate that a raw presentation XML stream has exactly the supplied
/// ordered slide bindings without producing a rewritten buffer.
pub(crate) fn validate_slide_bindings(
    xml: &[u8],
    current: &[(u32, &str)],
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<()> {
    let elements = slide_id_elements(xml, false, context)?;
    validate_slide_elements(&elements, current, context)
}

fn validate_slide_elements(
    elements: &[SlideIdElement],
    current: &[(u32, &str)],
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<()> {
    if elements.len() != current.len() {
        return Err(invalid(
            "opened-presentation slide-order XML differs from the semantic graph",
        ));
    }
    let mut current_ids = HashSet::new();
    let mut current_relationships = HashSet::new();
    current_ids
        .try_reserve(current.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide IDs",
            source,
        })?;
    current_relationships
        .try_reserve(current.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide relationship IDs",
            source,
        })?;
    for (index, (element, (id, relationship_id))) in elements.iter().zip(current).enumerate() {
        if index % 256 == 0 {
            check_execution_context(context)?;
        }
        if element.id != *id || element.relationship_id != *relationship_id {
            return Err(invalid(
                "opened-presentation slide-order binding changed during staging",
            ));
        }
        if !current_ids.insert(*id) || !current_relationships.insert(*relationship_id) {
            return Err(invalid(
                "opened-presentation slide-order bindings are ambiguous",
            ));
        }
    }
    check_execution_context(context)?;
    Ok(())
}

pub(crate) fn remove_slide(xml: &[u8], current: &[Slide], id: u32) -> Result<Vec<u8>> {
    remove_slide_binding(
        xml,
        current
            .iter()
            .map(|slide| (slide.id, slide.relationship_id.as_str())),
        id,
        None,
    )
}

/// Remove one complete `(slide-id, relationship-id)` binding from a raw
/// presentation root while retaining every byte outside that binding's span.
///
/// The caller supplies the complete semantic binding list, so this helper
/// proves the raw list has the same length, order, and one-to-one identities
/// before allocating the candidate.  Source-backed callers can pass their
/// stable binding iterator directly and retain their execution context.
/// The planner owns physical XML admission, root/markup policy, and the
/// final-slide rule; this lexical helper can produce an empty slide list.
pub(crate) fn remove_slide_binding<'a, I>(
    xml: &[u8],
    current: I,
    selected_id: u32,
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<Vec<u8>>
where
    I: ExactSizeIterator<Item = (u32, &'a str)>,
{
    let current_len = current.len();
    let elements = slide_id_elements(xml, false, context)?;
    if elements.len() != current_len {
        return Err(invalid(
            "opened-presentation slide-order XML differs from the semantic graph",
        ));
    }
    let mut current_ids = HashSet::new();
    let mut current_relationships = HashSet::new();
    current_ids
        .try_reserve(current_len)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide removal IDs",
            source,
        })?;
    current_relationships
        .try_reserve(current_len)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide removal relationship IDs",
            source,
        })?;
    let mut selected_index = None;
    for (index, (element, (id, relationship_id))) in elements.iter().zip(current).enumerate() {
        if index % 256 == 0 {
            check_execution_context(context)?;
        }
        if element.id != id || element.relationship_id != relationship_id {
            return Err(invalid(
                "opened-presentation slide-order binding changed during staging",
            ));
        }
        if !current_ids.insert(id) || !current_relationships.insert(relationship_id) {
            return Err(invalid(
                "opened-presentation slide-order bindings are ambiguous",
            ));
        }
        if id == selected_id {
            selected_index = Some(index);
        }
    }
    let selected_index = selected_index
        .ok_or_else(|| invalid("opened-presentation slide removal identity is missing"))?;
    let target = elements
        .get(selected_index)
        .ok_or_else(|| invalid("opened-presentation slide removal identity disappeared"))?;
    check_execution_context(context)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len().saturating_sub(target.span.len()))
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide removal XML",
            source,
        })?;
    check_execution_context(context)?;
    output.extend_from_slice(span_bytes(xml, 0..target.span.start)?);
    check_execution_context(context)?;
    output.extend_from_slice(span_bytes(xml, target.span.end..xml.len())?);
    check_execution_context(context)?;
    // The removed slide leaves its p14 section too (record 0303).
    if super::sections::present(&output) {
        if let Some(sectioned) = super::sections::remove(&output, selected_id)? {
            check_execution_context(context)?;
            return Ok(sectioned);
        }
    }
    Ok(output)
}

/// The byte spans that [`remove_slide_binding`] deletes for the slide
/// `selected_id`: its `p:sldId` entry and, in a sectioned root, its
/// `p14:sldId` section entry, in document order.
pub(crate) fn slide_binding_removal_spans(
    xml: &[u8],
    selected_id: u32,
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<Vec<Range<usize>>> {
    let entry = slide_id_elements(xml, false, context)?
        .into_iter()
        .find(|element| element.id == selected_id)
        .ok_or_else(|| invalid("opened-presentation slide removal identity is missing"))?;
    let mut spans = vec![entry.span];
    if super::sections::present(xml) {
        spans.extend(super::sections::member_spans(xml, selected_id)?);
    }
    spans.sort_unstable_by_key(|span| span.start);
    Ok(spans)
}

/// Prove that a root delta changes only slide-list membership and order.
/// Non-entry bytes and every retained binding remain exact. Scanning each
/// root once preserves inter-entry trivia without quadratic repeated removal.
///
/// A root with a p14 section list may also change section membership: the
/// proof then runs on both roots with every section entry stripped, and the
/// membership itself must be one that the catalog change keeps
/// ([`super::sections::validate_membership`], record 0303).
pub(crate) fn validate_slide_catalog_delta<'a, 'b, I, J>(
    before: &[u8],
    before_order: I,
    after: &[u8],
    after_order: J,
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<()>
where
    I: ExactSizeIterator<Item = (u32, &'a str)>,
    J: ExactSizeIterator<Item = (u32, &'b str)>,
{
    if !super::sections::present(before) && !super::sections::present(after) {
        return validate_unsectioned_catalog_delta(
            before,
            before_order,
            after,
            after_order,
            context,
        );
    }
    let before_order: Vec<(u32, &'a str)> = before_order.collect();
    let after_order: Vec<(u32, &'b str)> = after_order.collect();
    match (
        super::sections::strip(before)?,
        super::sections::strip(after)?,
    ) {
        (None, None) => validate_unsectioned_catalog_delta(
            before,
            before_order.into_iter(),
            after,
            after_order.into_iter(),
            context,
        ),
        (Some((before_stripped, before_members)), Some((after_stripped, after_members))) => {
            let before_ids: Vec<u32> = before_order.iter().map(|(id, _)| *id).collect();
            let after_ids: Vec<u32> = after_order.iter().map(|(id, _)| *id).collect();
            super::sections::validate_membership(
                &before_members,
                &after_members,
                &before_ids,
                &after_ids,
            )?;
            validate_unsectioned_catalog_delta(
                &before_stripped,
                before_order.into_iter(),
                &after_stripped,
                after_order.into_iter(),
                context,
            )
        },
        _ => Err(invalid(
            "source slide catalog proof adds or removes the section list",
        )),
    }
}

fn validate_unsectioned_catalog_delta<'a, 'b, I, J>(
    before: &[u8],
    before_order: I,
    after: &[u8],
    after_order: J,
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<()>
where
    I: ExactSizeIterator<Item = (u32, &'a str)>,
    J: ExactSizeIterator<Item = (u32, &'b str)>,
{
    check_execution_context(context)?;
    // Account for both parsed entry catalogs and their lookup tables before
    // parsing. The proof is transient and releases this admission on return.
    const MAX_PROOF_BYTES: usize = 128 * 1024 * 1024;
    let amount = before
        .len()
        .checked_add(after.len())
        .and_then(|bytes| bytes.checked_mul(16))
        .and_then(|bytes| bytes.checked_add(4096))
        .ok_or(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::ProofBytes),
            observed: usize::MAX,
            maximum: MAX_PROOF_BYTES,
        })?;
    if amount > MAX_PROOF_BYTES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::ProofBytes),
            observed: amount,
            maximum: MAX_PROOF_BYTES,
        });
    }
    let _reservation = context
        .map(|context| {
            context
                .reserve(litchi_core::Resource::Memory, amount as u64)
                .map_err(|error| {
                    Error::Opc(match error {
                        litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                        other => litchi_opc::OpcError::Execution(other),
                    })
                })
        })
        .transpose()?;
    let left = slide_id_elements(before, false, context)?;
    let right = slide_id_elements(after, false, context)?;
    if left.len() != before_order.len() || right.len() != after_order.len() {
        return Err(invalid(
            "source slide catalog proof differs from root bindings",
        ));
    }
    if left.is_empty() || right.is_empty() {
        return validate_empty_slide_catalog_delta(
            before,
            &left,
            before_order,
            after,
            &right,
            after_order,
            context,
        );
    }
    let allocation = |source| Error::Allocation {
        resource: "source slide catalog proof",
        source,
    };
    let mut retained = HashMap::new();
    retained.try_reserve(left.len()).map_err(allocation)?;
    let mut relationships = HashSet::new();
    relationships.try_reserve(left.len()).map_err(allocation)?;
    for (index, (entry, (id, rid))) in left.iter().zip(before_order).enumerate() {
        if index % 256 == 0 {
            check_execution_context(context)?;
        }
        if entry.id != id
            || entry.relationship_id != rid
            || retained.insert(id, entry).is_some()
            || !relationships.insert(rid)
        {
            return Err(invalid(
                "source slide catalog proof has ambiguous before bindings",
            ));
        }
    }
    let mut after_ids = HashSet::new();
    after_ids.try_reserve(right.len()).map_err(allocation)?;
    let mut after_relationships = HashSet::new();
    after_relationships
        .try_reserve(right.len())
        .map_err(allocation)?;
    for (index, (entry, (id, rid))) in right.iter().zip(after_order).enumerate() {
        if index % 256 == 0 {
            check_execution_context(context)?;
        }
        if entry.id != id
            || entry.relationship_id != rid
            || !after_ids.insert(id)
            || !after_relationships.insert(rid)
        {
            return Err(invalid(
                "source slide catalog proof has ambiguous after bindings",
            ));
        }
        if let Some(previous) = retained.get(&id) {
            if previous.relationship_id != rid
                || span_bytes(before, previous.span.clone())?
                    != span_bytes(after, entry.span.clone())?
            {
                return Err(invalid(
                    "source slide catalog proof changed a retained binding",
                ));
            }
        }
    }
    let gaps = |length: usize, entries: &[SlideIdElement]| -> Result<Vec<Range<usize>>> {
        let mut gaps = Vec::new();
        gaps.try_reserve_exact(entries.len() + 1)
            .map_err(allocation)?;
        let mut cursor = 0;
        for entry in entries {
            if entry.span.start < cursor {
                return Err(invalid("source slide catalog proof spans overlap"));
            }
            gaps.push(cursor..entry.span.start);
            cursor = entry.span.end;
        }
        gaps.push(cursor..length);
        Ok(gaps)
    };
    let left_gaps = gaps(before.len(), &left)?;
    let right_gaps = gaps(after.len(), &right)?;
    let (large_xml, large, small_xml, small) = if left.len() >= right.len() {
        (before, left_gaps, after, right_gaps)
    } else {
        (after, right_gaps, before, left_gaps)
    };
    // Native removal joins whole adjacent gaps; it cannot split one or move
    // trivia across retained slots. Inversion reverses exactly this relation.
    // Consume at least one original gap per target gap. Choosing the shortest
    // match leaves any empty gaps for later slots, including the final slot.
    let mut source_gap = 0;
    for (index, target) in small.iter().enumerate() {
        check_execution_context(context)?;
        let expected = span_bytes(small_xml, target.clone())?;
        let mut offset = 0usize;
        loop {
            let range = large
                .get(source_gap)
                .ok_or_else(|| invalid("source catalog trivia slots differ"))?;
            let bytes = span_bytes(large_xml, range.clone())?;
            let end = offset
                .checked_add(bytes.len())
                .ok_or_else(|| invalid("source catalog trivia size overflows"))?;
            if expected.get(offset..end) != Some(bytes) {
                return Err(invalid(
                    "source patch changes XML outside the slide catalog",
                ));
            }
            source_gap += 1;
            offset = end;
            if offset == expected.len() && (index + 1 < small.len() || source_gap == large.len()) {
                break;
            }
        }
    }
    if source_gap != large.len() {
        return Err(invalid("source catalog trivia slots differ"));
    }
    check_execution_context(context)
}

/// Empty slide lists need a slightly stronger proof than the normal
/// whole-gap comparison.  The empty-list insertion helper expands a
/// self-closing list and emits the first binding with a fresh, namespace-local
/// spelling.  Replaying those exact insertions from the empty source and
/// comparing the complete result proves every root byte outside the entries,
/// including wrappers, attributes, declarations, and trivia.
fn validate_empty_slide_catalog_delta<'a, 'b, I, J>(
    before: &[u8],
    before_elements: &[SlideIdElement],
    before_order: I,
    after: &[u8],
    after_elements: &[SlideIdElement],
    after_order: J,
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<()>
where
    I: ExactSizeIterator<Item = (u32, &'a str)>,
    J: ExactSizeIterator<Item = (u32, &'b str)>,
{
    let before_bindings = collect_catalog_bindings(before_elements, before_order, context)?;
    let after_bindings = collect_catalog_bindings(after_elements, after_order, context)?;
    if before_bindings.is_empty() && after_bindings.is_empty() {
        if before == after {
            return Ok(());
        }
        return Err(invalid(
            "empty source slide catalogs differ outside their slide list",
        ));
    }

    let (empty, populated, populated_bindings) = if before_bindings.is_empty() {
        (before, after, &after_bindings)
    } else if after_bindings.is_empty() {
        (after, before, &before_bindings)
    } else {
        return Err(invalid(
            "empty source slide catalog proof has an unexpected binding shape",
        ));
    };
    let replayed = replay_empty_slide_insertions(empty, populated_bindings, context)?;
    if replayed != populated {
        return Err(invalid(
            "empty source slide catalog changes bytes outside inserted bindings",
        ));
    }
    Ok(())
}

fn collect_catalog_bindings<'a, I>(
    elements: &[SlideIdElement],
    order: I,
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<Vec<(u32, String)>>
where
    I: ExactSizeIterator<Item = (u32, &'a str)>,
{
    let allocation = |source| Error::Allocation {
        resource: "source empty slide catalog bindings",
        source,
    };
    let mut bindings = Vec::new();
    bindings
        .try_reserve_exact(elements.len())
        .map_err(allocation)?;
    let mut ids = HashSet::new();
    let mut relationships = HashSet::new();
    ids.try_reserve(elements.len()).map_err(allocation)?;
    relationships
        .try_reserve(elements.len())
        .map_err(allocation)?;
    for (index, (element, (id, relationship_id))) in elements.iter().zip(order).enumerate() {
        if index % 256 == 0 {
            check_execution_context(context)?;
        }
        if element.id != id
            || element.relationship_id != relationship_id
            || !ids.insert(id)
            || !relationships.insert(relationship_id)
        {
            return Err(invalid("source empty slide catalog has ambiguous bindings"));
        }
        bindings.push((id, relationship_id.to_owned()));
    }
    check_execution_context(context)?;
    Ok(bindings)
}

fn replay_empty_slide_insertions(
    xml: &[u8],
    bindings: &[(u32, String)],
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<Vec<u8>> {
    let mut output = clone_xml_bytes(xml, "source empty slide catalog replay")?;
    let mut current: Vec<(u32, String)> = Vec::new();
    current
        .try_reserve_exact(bindings.len())
        .map_err(|source| Error::Allocation {
            resource: "source empty slide catalog replay bindings",
            source,
        })?;
    for (position, (id, relationship_id)) in bindings.iter().enumerate() {
        if position % 256 == 0 {
            check_execution_context(context)?;
        }
        output = insert_slide_binding(
            &output,
            current.iter().map(|(current_id, current_relationship)| {
                (*current_id, current_relationship.as_str())
            }),
            position,
            *id,
            relationship_id,
        )?;
        current.push((*id, relationship_id.clone()));
    }
    check_execution_context(context)?;
    Ok(output)
}

pub(crate) fn insert_slide(
    xml: &[u8],
    current: &[Slide],
    position: usize,
    id: u32,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    insert_slide_binding(
        xml,
        current
            .iter()
            .map(|slide| (slide.id, slide.relationship_id.as_str())),
        position,
        id,
        relationship_id,
    )
}

/// Insert one self-contained `p:sldId` entry while validating only the lexical
/// presentation binding supplied by the caller.
///
/// Source-backed presentations deliberately avoid materializing opened
/// [`Slide`] values.  This binding-based variant preserves the exact source
/// XML ranges while remaining usable by both presentation facades.
pub(crate) fn insert_slide_binding<'a, I>(
    xml: &[u8],
    current: I,
    position: usize,
    id: u32,
    relationship_id: &str,
) -> Result<Vec<u8>>
where
    I: ExactSizeIterator<Item = (u32, &'a str)>,
{
    // Insertion callers allocate new relationship IDs using the native authoring
    // grammar. Bound and validate them before copying bytes into an attribute.
    if relationship_id.len() > crate::media_parts::MAX_STRING_BYTES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::RelationshipBytes),
            observed: relationship_id.len(),
            maximum: crate::media_parts::MAX_STRING_BYTES,
        });
    }
    crate::media_parts::validate_id(relationship_id)?;
    let current_len = current.len();
    if position > current_len {
        return Err(invalid(
            "opened-presentation slide insertion position is out of bounds",
        ));
    }
    let elements = slide_id_elements(xml, true, None)?;
    if elements.len() != current_len {
        return Err(invalid(
            "opened-presentation slide-order XML differs from the semantic graph",
        ));
    }
    let mut current = current;
    for element in &elements {
        let Some((expected_id, expected_relationship_id)) = current.next() else {
            return Err(invalid(
                "opened-presentation slide-order XML differs from the semantic graph",
            ));
        };
        if element.id != expected_id || element.relationship_id != expected_relationship_id {
            return Err(invalid(
                "opened-presentation slide-order binding changed during staging",
            ));
        }
    }
    if current.next().is_some() {
        return Err(invalid(
            "opened-presentation slide-order XML differs from the semantic graph",
        ));
    }
    if elements.is_empty() {
        return empty_slides::insert(xml, id, relationship_id);
    }
    let exemplar = elements
        .first()
        .ok_or_else(|| invalid("opened-presentation slide list is empty"))?;
    let element_name = exemplar
        .element_name
        .as_deref()
        .ok_or_else(|| invalid("opened-presentation slide entry name is missing"))?;
    let relationship_attribute_name = exemplar
        .relationship_attribute_name
        .as_deref()
        .ok_or_else(|| invalid("opened-presentation relationship attribute name is missing"))?;
    let element_namespace = exemplar
        .element_namespace
        .as_deref()
        .ok_or_else(|| invalid("opened-presentation slide entry namespace is missing"))?;
    let relationship_namespace = exemplar
        .relationship_namespace
        .as_deref()
        .ok_or_else(|| invalid("opened-presentation relationship namespace is missing"))?;
    if !relationship_attribute_name.contains(&b':') {
        return Err(invalid(
            "opened-presentation relationship attribute has no namespace prefix",
        ));
    }
    let insertion = elements
        .get(position)
        .map_or_else(
            || elements.last().map(|element| element.span.end),
            |element| Some(element.span.start),
        )
        .ok_or_else(|| invalid("opened-presentation slide insertion point is missing"))?;
    let mut fragment = Vec::new();
    let capacity = element_name
        .len()
        .checked_add(relationship_attribute_name.len())
        .and_then(|value| value.checked_add(element_namespace.len()))
        .and_then(|value| value.checked_add(relationship_namespace.len()))
        .and_then(|value| value.checked_add(relationship_id.len()))
        .and_then(|value| value.checked_add(96))
        .ok_or_else(|| invalid("opened-presentation slide entry size overflow"))?;
    fragment
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation inserted slide entry",
            source,
        })?;
    fragment.push(b'<');
    fragment.extend_from_slice(element_name);
    append_namespace_binding(&mut fragment, element_name, element_namespace)?;
    append_namespace_binding(
        &mut fragment,
        relationship_attribute_name,
        relationship_namespace,
    )?;
    fragment.extend_from_slice(b" id=\"");
    fragment.extend_from_slice(id.to_string().as_bytes());
    fragment.extend_from_slice(b"\" ");
    fragment.extend_from_slice(relationship_attribute_name);
    fragment.extend_from_slice(b"=\"");
    fragment.extend_from_slice(relationship_id.as_bytes());
    fragment.extend_from_slice(b"\"/>");
    let output_len = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| invalid("opened-presentation inserted XML size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation inserted slide XML",
            source,
        })?;
    output.extend_from_slice(span_bytes(xml, 0..insertion)?);
    output.extend_from_slice(&fragment);
    output.extend_from_slice(span_bytes(xml, insertion..xml.len())?);
    Ok(output)
}

fn append_namespace_binding(output: &mut Vec<u8>, name: &[u8], namespace: &[u8]) -> Result<()> {
    output.extend_from_slice(b" xmlns");
    if let Some(separator) = name.iter().position(|byte| *byte == b':') {
        if separator == 0 {
            return Err(invalid("opened-presentation XML name has an empty prefix"));
        }
        output.push(b':');
        output.extend_from_slice(&name[..separator]);
    }
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(namespace);
    output.push(b'"');
    Ok(())
}

pub(crate) fn rewrite_shape_text_with_limits(
    xml: &[u8],
    shape: Range<usize>,
    text: &str,
    max_text_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<Vec<u8>> {
    let edits = [ShapeTextEdit {
        index: 0,
        shape,
        text,
        changed: true,
    }];
    rewrite_shape_texts(xml, &edits, max_text_bytes, shape_limits.output_bytes())?
        .ok_or_else(|| invalid("opened-presentation shape text rewrite disappeared"))
}

pub(crate) fn shape_text_skeleton_with_limits(
    xml: &[u8],
    shape_limits: crate::shape::Limits,
) -> Result<[u8; 32]> {
    use sha2::{Digest, Sha256};
    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    if scene.is_rewritten() {
        return Err(invalid(
            "source patch text proof cannot select MCE branches",
        ));
    }
    let mut owners = Vec::new();
    owners
        .try_reserve_exact(scene.len())
        .map_err(|source| Error::Allocation {
            resource: "source patch editable text owners",
            source,
        })?;
    for shape in scene.iter() {
        if shape.common().text().is_some() {
            owners.push(shape.common().span()?.range(xml.len())?);
        }
    }
    owners.sort_unstable_by_key(|owner| (owner.start, owner.end));
    if owners.windows(2).any(|pair| pair[0].end > pair[1].start) {
        return Err(invalid("source patch editable text owners overlap"));
    }
    let slots = drawing_text_elements_for_owners(xml, &owners)?;
    let mut hash = Sha256::new();
    hash.update(b"litchi-pptx-source-text-skeleton-v1\0");
    hash.update((scene.len() as u64).to_le_bytes());
    hash.update((owners.len() as u64).to_le_bytes());
    let mut cursor = 0;
    for slot in slots.iter().flatten() {
        if slot.span.start < cursor || slot.span.end > xml.len() {
            return Err(invalid("source patch text slots overlap"));
        }
        hash.update(text_span_bytes(xml, cursor..slot.span.start)?);
        if let Some(name) = &slot.empty_name {
            let raw = text_span_bytes(xml, slot.span.clone())?;
            let open_end = empty_text_open_end(raw)?;
            hash.update(&raw[..open_end]);
            hash.update(b"></");
            hash.update(name);
            hash.update(b">");
        }
        cursor = slot.span.end;
    }
    hash.update(text_span_bytes(xml, cursor..xml.len())?);
    Ok(hash.finalize().into())
}

// Literal CR is normalized by XML readers, including inside xml:space=preserve.
// Character references preserve the caller's logical text. Escape with an exact
// byte admission and retain the borrowed fast path when nothing needs escaping.
pub(super) fn escape_character_data(
    text: &str,
    maximum: usize,
) -> Result<std::borrow::Cow<'_, str>> {
    fn entity(byte: u8) -> Option<&'static str> {
        match byte {
            b'&' => Some("&amp;"),
            b'<' => Some("&lt;"),
            b'>' => Some("&gt;"),
            b'"' => Some("&quot;"),
            b'\'' => Some("&apos;"),
            b'\r' => Some("&#13;"),
            _ => None,
        }
    }
    let mut length = 0usize;
    for byte in text.bytes() {
        length = length
            .checked_add(entity(byte).map_or(1, str::len))
            .ok_or(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
                observed: usize::MAX,
                maximum,
            })?;
        if length > maximum {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
                observed: length,
                maximum,
            });
        }
    }
    if length == text.len() {
        return Ok(std::borrow::Cow::Borrowed(text));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation escaped text",
            source,
        })?;
    let mut start = 0;
    for (index, byte) in text.bytes().enumerate() {
        if let Some(replacement) = entity(byte) {
            output.push_str(&text[start..index]);
            output.push_str(replacement);
            start = index + 1;
        }
    }
    output.push_str(&text[start..]);
    Ok(std::borrow::Cow::Owned(output))
}

pub(crate) fn stage_shape_texts_with_limits(
    xml: &[u8],
    replacements: &[super::ShapeTextReplacement<'_>],
    max_text_bytes: usize,
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    stage_shape_texts_admitted_with_limits(
        xml,
        replacements,
        max_text_bytes,
        max_output_bytes,
        super::MAX_SHAPE_TEXT_REPLACEMENTS,
        max_text_bytes,
        shape_limits,
    )
}

pub(super) fn stage_shape_texts_admitted_with_limits(
    xml: &[u8],
    replacements: &[super::ShapeTextReplacement<'_>],
    max_text_bytes: usize,
    max_output_bytes: usize,
    selection_limit: usize,
    aggregate_text_limit: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    if replacements.is_empty() {
        return Ok((None, 0));
    }
    if replacements.len() > selection_limit {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextReplacements),
            observed: replacements.len(),
            maximum: selection_limit,
        });
    }
    let aggregate_text_bytes = replacements.iter().try_fold(0usize, |total, replacement| {
        total
            .checked_add(replacement.text().len())
            .ok_or_else(|| invalid("opened-presentation aggregate shape text size overflow"))
    })?;
    if aggregate_text_bytes > aggregate_text_limit {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
            observed: aggregate_text_bytes,
            maximum: aggregate_text_limit,
        });
    }

    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    if scene.is_rewritten() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "set_shape_texts",
            "atomic shape text batches do not support markup-compatibility branch selection",
        ));
    }
    if replacements.len() > scene.len() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextReplacements),
            observed: replacements.len(),
            maximum: scene.len(),
        });
    }
    let mut selected = HashSet::new();
    selected
        .try_reserve(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text identities",
            source,
        })?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text edit plan",
            source,
        })?;
    for replacement in replacements {
        let shape = scene.shape(replacement.selector())?;
        let index = shape.common().index();
        if !selected.insert(index) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::InvalidValue,
                "set_shape_text",
                "a shape is selected more than once",
            ));
        }
        let current = shape.common().text().ok_or(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "set_shape_text",
            "the selected shape has no text body",
        ))?;
        edits.push(ShapeTextEdit {
            index,
            shape: shape.common().span()?.range(xml.len())?,
            text: replacement.text(),
            changed: current != replacement.text(),
        });
    }
    edits.sort_unstable_by_key(|edit| (edit.shape.start, edit.shape.end));
    for pair in edits.windows(2) {
        if pair[0].shape.end > pair[1].shape.start {
            return Err(invalid("opened-presentation shape text selections overlap"));
        }
    }
    let changed = edits.iter().filter(|edit| edit.changed).count();
    let staged = rewrite_shape_texts(xml, &edits, max_text_bytes, max_output_bytes)?;
    let Some(candidate) = staged.as_deref() else {
        return Ok((None, 0));
    };
    let published = crate::shape::Scene::read(candidate, shape_limits)?;
    if published.is_rewritten() || published.len() != scene.len() {
        return Err(invalid(
            "opened-presentation shape text batch changed the semantic shape graph",
        ));
    }
    for edit in &edits {
        if published.at(edit.index)?.common().text() != Some(edit.text) {
            return Err(invalid(
                "opened-presentation shape text batch did not round-trip semantically",
            ));
        }
    }
    Ok((staged, changed))
}

fn rewrite_shape_texts(
    xml: &[u8],
    edits: &[ShapeTextEdit<'_>],
    max_text_bytes: usize,
    max_output_bytes: usize,
) -> Result<Option<Vec<u8>>> {
    for edit in edits {
        if edit.text.len() > max_text_bytes {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
                observed: edit.text.len(),
                maximum: max_text_bytes,
            });
        }
        if !edit.text.chars().all(is_xml_char) {
            return Err(invalid(
                "opened-presentation shape text contains an invalid XML character",
            ));
        }
        if edit.shape.start >= edit.shape.end || edit.shape.end > xml.len() {
            return Err(invalid("opened-presentation shape range is invalid"));
        }
    }
    let mut owners = Vec::new();
    owners
        .try_reserve_exact(edits.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text owner ranges",
            source,
        })?;
    owners.extend(edits.iter().map(|edit| edit.shape.clone()));
    let elements = drawing_text_elements_for_owners(xml, &owners)?;
    if elements.iter().any(Vec::is_empty) {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "set_shape_text",
            "the selected shape has no DrawingML text run",
        ));
    }
    if !edits.iter().any(|edit| edit.changed) {
        return Ok(None);
    }
    // A shape text set fills the run slots; writing into a field's cached
    // result, or blanking it, would lose the field's text (record 0304).
    if edits
        .iter()
        .zip(&elements)
        .any(|(edit, spans)| edit.changed && spans.iter().any(|span| span.field))
    {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "set_shape_text",
            "the shape text holds a field (a:fld), which a shape text set would \
                     overwrite; set the text of its runs instead",
        ));
    }
    let mut escaped = Vec::new();
    escaped
        .try_reserve_exact(edits.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation escaped shape texts",
            source,
        })?;
    for edit in edits {
        escaped.push(if edit.changed {
            escape_character_data(edit.text, max_output_bytes)?
        } else {
            std::borrow::Cow::Borrowed("")
        });
    }
    let mut removed = 0usize;
    let mut emitted = 0usize;
    // The emission below copies the bytes between consecutive spans, so the
    // spans it writes must be in bounds, in document order and disjoint. The
    // locator produces them that way; a violation is refused here rather
    // than reaching a slice.
    let mut planned_end = 0usize;
    for ((edit, spans), replacement) in edits.iter().zip(&elements).zip(&escaped) {
        if !edit.changed {
            continue;
        }
        for (position, span) in spans.iter().enumerate() {
            if span.span.start < planned_end
                || span.span.start > span.span.end
                || span.span.end > xml.len()
            {
                return Err(invalid(
                    "opened-presentation shape text spans are out of order",
                ));
            }
            planned_end = span.span.end;
            removed = removed
                .checked_add(span.span.len())
                .ok_or_else(|| invalid("opened-presentation shape text size overflow"))?;
            emitted = emitted
                .checked_add(text_element_output_len(
                    xml,
                    span,
                    position,
                    replacement.len(),
                )?)
                .ok_or_else(|| invalid("opened-presentation shape text size overflow"))?;
        }
    }
    let output_len = xml
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(emitted))
        .ok_or_else(|| invalid("opened-presentation shape text output size overflow"))?;
    if output_len > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: output_len,
            maximum: max_output_bytes,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text XML",
            source,
        })?;
    let mut cursor = 0usize;
    for ((edit, spans), replacement) in edits.iter().zip(&elements).zip(&escaped) {
        if !edit.changed {
            continue;
        }
        for (position, span) in spans.iter().enumerate() {
            output.extend_from_slice(text_span_bytes(xml, cursor..span.span.start)?);
            write_text_element(&mut output, xml, span, position, replacement.as_bytes())?;
            cursor = span.span.end;
        }
    }
    output.extend_from_slice(text_span_bytes(xml, cursor..xml.len())?);
    if output.len() != output_len {
        return Err(invalid(
            "opened-presentation shape text output length changed during emission",
        ));
    }
    Ok(Some(output))
}

fn text_element_output_len(
    xml: &[u8],
    element: &TextElement,
    position: usize,
    escaped_len: usize,
) -> Result<usize> {
    if position != 0 {
        return Ok(element
            .empty_name
            .as_ref()
            .map_or(0, |_| element.span.len()));
    }
    let Some(name) = &element.empty_name else {
        return Ok(escaped_len);
    };
    let open_end = empty_text_open_end(text_span_bytes(xml, element.span.clone())?)?;
    open_end
        .checked_add(escaped_len)
        .and_then(|length| length.checked_add(name.len()))
        .and_then(|length| length.checked_add(4))
        .ok_or_else(|| invalid("opened-presentation empty text expansion size overflow"))
}

fn write_text_element(
    output: &mut Vec<u8>,
    xml: &[u8],
    element: &TextElement,
    position: usize,
    escaped: &[u8],
) -> Result<()> {
    if position != 0 {
        if element.empty_name.is_some() {
            output.extend_from_slice(text_span_bytes(xml, element.span.clone())?);
        }
        return Ok(());
    }
    let Some(name) = &element.empty_name else {
        output.extend_from_slice(escaped);
        return Ok(());
    };
    let raw = text_span_bytes(xml, element.span.clone())?;
    let open_end = empty_text_open_end(raw)?;
    output.extend_from_slice(&raw[..open_end]);
    output.push(b'>');
    output.extend_from_slice(escaped);
    output.extend_from_slice(b"</");
    output.extend_from_slice(name);
    output.push(b'>');
    Ok(())
}

/// The bytes of one planned text span, or a typed refusal for a range that is
/// reversed or out of bounds, so no span can turn into a slicing panic.
fn text_span_bytes(xml: &[u8], range: Range<usize>) -> Result<&[u8]> {
    xml.get(range)
        .ok_or_else(|| invalid("opened-presentation shape text spans are out of order"))
}

/// The bytes of `range` in `xml`, or a typed refusal for a range that is
/// reversed or out of bounds. Every span an edit copies around comes from a
/// reader pass over the same bytes and is ordered by construction; this keeps
/// a violation of that invariant a refusal rather than a slicing panic.
pub(crate) fn span_bytes(xml: &[u8], range: Range<usize>) -> Result<&[u8]> {
    xml.get(range)
        .ok_or_else(|| invalid("opened-presentation XML span is out of order or out of bounds"))
}

fn empty_text_open_end(raw: &[u8]) -> Result<usize> {
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| invalid("opened-presentation empty text tag is malformed"))?;
    let mut open_end = slash;
    while open_end > 0 && raw[open_end - 1].is_ascii_whitespace() {
        open_end -= 1;
    }
    Ok(open_end)
}

/// Locate the byte immediately before the closing presentation shape tree.
///
/// The source authoring path appends native fragments at this exact boundary.
/// Keeping the scan shared with the durable proof means the proof can require
/// that an admitted shape really was appended by that path, rather than merely
/// appearing somewhere among the direct scene roots.
fn shape_tree_insertion(xml: &[u8]) -> Result<usize> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut tree_depth = None;
    let mut trees = 0usize;
    let mut insertion = None;
    loop {
        let start = position(&reader, origin)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let pml_namespace = is_presentation_namespace(&namespace);
        let event = event.into_owned();
        drop(namespace);
        match event {
            Event::Start(element) => {
                bump(&mut nodes)?;
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: depth.checked_add(1).unwrap_or(usize::MAX),
                        maximum: MAX_XML_DEPTH,
                    });
                }
                depth += 1;
                if pml_namespace && element.local_name().as_ref() == b"spTree" {
                    trees = trees.saturating_add(1);
                    if trees != 1 || tree_depth.replace(depth).is_some() {
                        return Err(invalid(
                            "opened-presentation slide has multiple shape trees",
                        ));
                    }
                }
            },
            Event::Empty(element) => {
                bump(&mut nodes)?;
                if pml_namespace && element.local_name().as_ref() == b"spTree" {
                    return Err(invalid(
                        "opened-presentation cannot append to an empty shape-tree element",
                    ));
                }
            },
            Event::End(element) => {
                if tree_depth == Some(depth)
                    && pml_namespace
                    && element.local_name().as_ref() == b"spTree"
                {
                    insertion = Some(start);
                    tree_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation XML depth underflow"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || tree_depth.is_some() {
        return Err(invalid("opened-presentation slide XML is unterminated"));
    }
    if trees != 1 {
        return Err(invalid(
            "opened-presentation slide must have one shape tree",
        ));
    }
    insertion.ok_or_else(|| invalid("opened-presentation slide has no shape tree"))
}

pub(crate) fn append_shape(xml: &[u8], fragment: &[u8]) -> Result<Vec<u8>> {
    if fragment.is_empty() {
        return Err(invalid(
            "opened-presentation shape fragment cannot be empty",
        ));
    }
    let insertion = shape_tree_insertion(xml)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len().saturating_add(fragment.len()))
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation appended shape XML",
            source,
        })?;
    output.extend_from_slice(span_bytes(xml, 0..insertion)?);
    output.extend_from_slice(fragment);
    output.extend_from_slice(span_bytes(xml, insertion..xml.len())?);
    Ok(output)
}

/// Append one native shape after an exact output-size preflight.
pub(crate) fn append_shape_with_limit(
    xml: &[u8],
    fragment: &[u8],
    max_output_bytes: usize,
) -> Result<Vec<u8>> {
    let output_len = xml.len().checked_add(fragment.len()).ok_or(Error::Limit {
        kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
        observed: usize::MAX,
        maximum: max_output_bytes,
    })?;
    if output_len > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: output_len,
            maximum: max_output_bytes,
        });
    }
    let output = append_shape(xml, fragment)?;
    if output.len() != output_len {
        return Err(invalid(
            "native shape append changed the preflighted output size",
        ));
    }
    Ok(output)
}

/// Remove a suffix of native text-box roots when the suffix is exactly what
/// the source authoring path emits.
///
/// Source patches already prove ordinary edits by hashing the XML outside
/// DrawingML text slots.  A newly authored text box necessarily changes that
/// skeleton, so this helper admits the one structural change that the source
/// API owns: one or more canonical `p:sp` roots appended directly before the
/// shape-tree end tag.  Returning `None` is deliberate for every other shape
/// or placement; the caller then applies the ordinary skeleton proof and
/// rejects the pair.
#[cfg(test)]
pub(crate) fn strip_appended_text_boxes(before: &[u8], after: &[u8]) -> Result<Option<Vec<u8>>> {
    strip_appended_text_boxes_with_limits(before, after, crate::shape::Limits::DEFAULT)
}

pub(crate) fn strip_appended_text_boxes_with_limits(
    before: &[u8],
    after: &[u8],
    shape_limits: crate::shape::Limits,
) -> Result<Option<Vec<u8>>> {
    let before_scene = crate::shape::Scene::read(before, shape_limits)?;
    let after_scene = crate::shape::Scene::read(after, shape_limits)?;
    if before_scene.is_rewritten() || after_scene.is_rewritten() {
        return Ok(None);
    }

    let before_roots = before_scene
        .roots()
        .map(|shape| shape.span())
        .collect::<Result<Vec<_>>>()?;
    let after_roots = after_scene
        .roots()
        .map(|shape| shape.span())
        .collect::<Result<Vec<_>>>()?;
    if after_roots.len() <= before_roots.len() {
        return Ok(None);
    }

    let before_insertion = shape_tree_insertion(before)?;
    let after_insertion = shape_tree_insertion(after)?;
    if span_bytes(before, before_insertion..before.len())?
        != span_bytes(after, after_insertion..after.len())?
    {
        return Ok(None);
    }

    let added = &after_roots[before_roots.len()..];
    let added_ranges = added
        .iter()
        .map(|span| span.range(after.len()))
        .collect::<Result<Vec<_>>>()?;
    let Some(first) = added_ranges.first() else {
        return Ok(None);
    };
    let Some(last) = added_ranges.last() else {
        return Ok(None);
    };
    if last.end != after_insertion
        || added_ranges
            .windows(2)
            .any(|pair| pair[0].end != pair[1].start)
    {
        return Ok(None);
    }

    let mut next_id = super::next_shape_id(&before_scene, "text box")?;
    let conformance = crate::media_parts::document_conformance(after)?;
    for (index, shape) in after_scene.roots().skip(before_roots.len()).enumerate() {
        let crate::shape::Shape::Auto(_) = shape else {
            return Ok(None);
        };
        let Some(id) = shape.id() else {
            return Ok(None);
        };
        if id != next_id || shape.placeholder().is_some() {
            return Ok(None);
        }
        let Some(bounds) = shape.bounds() else {
            return Ok(None);
        };
        let Some(text) = shape.text() else {
            return Ok(None);
        };
        let expected_name = format!("Text Box {id}");
        if shape.name() != Some(expected_name.as_str()) {
            return Ok(None);
        }
        let expected = super::text_box_fragment(
            text,
            (bounds.x(), bounds.y(), bounds.width(), bounds.height()),
            id,
            conformance,
            super::Limits::default(),
        )?;
        if shape.xml()? != expected.as_bytes() {
            return Ok(None);
        }
        if index + 1 < added.len() {
            next_id = next_id
                .checked_add(1)
                .ok_or_else(|| invalid("source text-box shape ID overflow"))?;
        }
    }

    let mut stripped = Vec::new();
    let retained = after
        .len()
        .checked_sub(last.end - first.start)
        .ok_or_else(|| invalid("source text-box suffix span underflows"))?;
    stripped
        .try_reserve_exact(retained)
        .map_err(|source| Error::Allocation {
            resource: "source patch stripped text-box XML",
            source,
        })?;
    stripped.extend_from_slice(span_bytes(after, 0..first.start)?);
    stripped.extend_from_slice(span_bytes(after, last.end..after.len())?);
    Ok(Some(stripped))
}

pub(crate) fn remap_shape_fragment(
    source: &[u8],
    root_source_id: u32,
    root_name: &str,
    shape_ids: &HashMap<u32, u32>,
    relationships: &HashMap<String, String>,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut output = Vec::with_capacity(source.len());
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut remap = ShapeRemap {
        root_source_id,
        root_name,
        shape_ids,
        relationships,
        written_shape_ids: HashSet::new(),
    };
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) => {
                if depth == 0 {
                    roots = roots.saturating_add(1);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("opened-presentation shape depth overflow"))?;
                write_remapped_shape_start(
                    &mut output,
                    &element,
                    reader.decoder(),
                    reader.resolver(),
                    false,
                    &mut remap,
                )?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    roots = roots.saturating_add(1);
                }
                write_remapped_shape_start(
                    &mut output,
                    &element,
                    reader.decoder(),
                    reader.resolver(),
                    true,
                    &mut remap,
                )?;
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation shape depth underflow"))?;
                output.extend_from_slice(b"</");
                output.extend_from_slice(element.name().as_ref());
                output.push(b'>');
            },
            Event::Text(text) => output.extend_from_slice(text.as_ref()),
            Event::CData(data) => {
                output.extend_from_slice(b"<![CDATA[");
                output.extend_from_slice(data.as_ref());
                output.extend_from_slice(b"]]");
                output.push(b'>');
            },
            Event::GeneralRef(reference) => {
                output.push(b'&');
                output.extend_from_slice(reference.as_ref());
                output.push(b';');
            },
            Event::Comment(comment) => {
                output.extend_from_slice(b"<!--");
                output.extend_from_slice(comment.as_ref());
                output.extend_from_slice(b"-->");
            },
            Event::PI(instruction) => {
                output.extend_from_slice(b"<?");
                output.extend_from_slice(instruction.as_ref());
                output.extend_from_slice(b"?>");
            },
            Event::Decl(_) | Event::DocType(_) => {
                return Err(invalid(
                    "opened-presentation shape fragment contains document-level markup",
                ));
            },
            Event::Eof => break,
        }
    }
    if depth != 0 || roots != 1 || !remap.written_shape_ids.contains(&root_source_id) {
        return Err(invalid(
            "opened-presentation transferred shape must have one root and a complete identity map",
        ));
    }
    Ok(output)
}

pub(crate) fn connector_connection_ids(source: &[u8]) -> Result<BTreeSet<u32>> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut connections = BTreeSet::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element)
                if matches!(element.local_name().as_ref(), b"stCxn" | b"endCxn") =>
            {
                let mut identity = None;
                for attribute in element.checked_attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    if attribute.key.as_ref() != b"id" {
                        continue;
                    }
                    if identity.is_some() {
                        return Err(invalid(
                            "opened-presentation connector endpoint repeats its identity attribute",
                        ));
                    }
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    identity = Some(value.parse::<u32>().map_err(|_err| {
                        invalid("opened-presentation connector endpoint identity is invalid")
                    })?);
                }
                connections.insert(identity.ok_or_else(|| {
                    invalid("opened-presentation connector endpoint identity is missing")
                })?);
            },
            Event::Decl(_) | Event::DocType(_) => {
                return Err(invalid(
                    "opened-presentation shape fragment contains document-level markup",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(connections)
}

fn shape_removal_plan_with_limits(
    xml: &[u8],
    key: crate::shape::Key<'_>,
    shape_limits: crate::shape::Limits,
) -> Result<(Range<usize>, BTreeSet<u32>, usize)> {
    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    if scene.is_rewritten() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "remove_shape",
            "shape removal does not support markup-compatibility branch selection",
        ));
    }
    shape_removal_plan_in(&scene, xml, key)
}

/// [`shape_removal_plan_with_limits`] on an already read scene whose spans are defined
/// against `xml` (the processed owner of a rewritten scene, record 0303).
pub(super) fn shape_removal_plan_in(
    scene: &crate::shape::Scene<'_>,
    xml: &[u8],
    key: crate::shape::Key<'_>,
) -> Result<(Range<usize>, BTreeSet<u32>, usize)> {
    let selected = scene.shape(key)?;
    let selected_index = selected.common().index();
    let root = scene
        .roots()
        .find(|shape| shape.common().index() == selected_index)
        .ok_or(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "remove_shape",
            "shape removal requires a direct top-level shape",
        ))?;
    let root_span = root.common().span()?.range(xml.len())?;
    let mut identities = BTreeSet::new();
    let mut removed_identities = BTreeSet::new();
    for shape in scene.iter() {
        let span = shape.common().span()?.range(xml.len())?;
        let inside_root = root_span.start <= span.start && span.end <= root_span.end;
        let Some(id) = shape.id() else {
            return Err(invalid(
                "opened-presentation shape removal requires identifiable shapes",
            ));
        };
        if !identities.insert(id) {
            return Err(invalid(
                "opened-presentation shape removal requires unique shape identities",
            ));
        }
        if inside_root {
            removed_identities.insert(id);
        }
    }
    if removed_identities.is_empty() {
        return Err(invalid(
            "opened-presentation removed shape has no non-visual identity",
        ));
    }
    Ok((root_span, removed_identities, scene.len()))
}

pub(crate) fn shape_removal_span_with_limits(
    xml: &[u8],
    key: crate::shape::Key<'_>,
    shape_limits: crate::shape::Limits,
) -> Result<Range<usize>> {
    shape_removal_plan_with_limits(xml, key, shape_limits).map(|(span, _, _)| span)
}

pub(crate) fn stage_shape_removal_with_limits(
    xml: &[u8],
    key: crate::shape::Key<'_>,
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<(Option<Vec<u8>>, usize)> {
    let (span, removed_identities, shape_count) =
        shape_removal_plan_with_limits(xml, key, shape_limits)?;
    let retained = xml
        .len()
        .checked_sub(span.len())
        .ok_or_else(|| invalid("opened-presentation removed shape span underflows"))?;
    if retained > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: retained,
            maximum: max_output_bytes,
        });
    }
    let mut candidate = Vec::new();
    candidate
        .try_reserve_exact(retained)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation removed shape XML",
            source,
        })?;
    candidate.extend_from_slice(span_bytes(xml, 0..span.start)?);
    candidate.extend_from_slice(span_bytes(xml, span.end..xml.len())?);
    validate_shape_removal_dependencies(&candidate, &removed_identities)?;

    let published = crate::shape::Scene::read(&candidate, shape_limits)?;
    if published.is_rewritten()
        || published.len().checked_add(removed_identities.len()) != Some(shape_count)
    {
        return Err(invalid(
            "opened-presentation shape removal changed the semantic shape graph",
        ));
    }
    let original = crate::shape::Scene::read(xml, shape_limits)?;
    let mut retained_shapes = original.iter().filter(|shape| {
        shape
            .id()
            .is_none_or(|id| !removed_identities.contains(&id))
    });
    for shape in published.iter() {
        let Some(previous) = retained_shapes.next() else {
            return Err(invalid(
                "opened-presentation shape removal changed retained shape identities",
            ));
        };
        if previous.id() != shape.id() {
            return Err(invalid(
                "opened-presentation shape removal changed retained shape identities",
            ));
        }
    }
    if retained_shapes.next().is_some() {
        return Err(invalid(
            "opened-presentation shape removal lost a retained shape",
        ));
    }
    Ok((Some(candidate), 1))
}

/// Validate that no connector or animation target in the retained XML points
/// at a removed shape identity. The scan intentionally covers only known OOXML
/// target vocabularies; malformed values in those vocabularies are refused.
pub(crate) fn validate_shape_removal_dependencies(
    candidate: &[u8],
    removed_identities: &BTreeSet<u32>,
) -> Result<()> {
    let references = shape_reference_ids(candidate)?;
    if references
        .iter()
        .any(|identity| removed_identities.contains(identity))
    {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "remove_shape",
            "the removed shape is referenced by a connector or animation target",
        ));
    }
    Ok(())
}

/// Collect connector, timing, show-event and control shape associations from
/// one owner. Unknown identity-bearing markup is an explicit removal boundary.
pub(crate) fn shape_reference_ids(source: &[u8]) -> Result<BTreeSet<u32>> {
    const P14: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    let decoder = reader.decoder();
    let mut references = BTreeSet::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let (Event::Start(element) | Event::Empty(element)) = event else {
            if matches!(event, Event::DocType(_)) {
                return Err(invalid(
                    "opened-presentation shape XML contains document-level markup",
                ));
            }
            if matches!(event, Event::Eof) {
                break;
            }
            continue;
        };
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(value)) => value,
            _ => b"".as_slice(),
        };
        let local = element.local_name();
        let attribute_name = if (namespace == crate::namespace::PRESENTATIONML_NAMESPACE
            || namespace == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE)
            && matches!(
                local.as_ref(),
                b"spTgt" | b"bldP" | b"bldDgm" | b"bldOleChart" | b"bldGraphic" | b"control"
            ) {
            Some(b"spid".as_slice())
        } else if namespace == P14 && local.as_ref() == b"bmkTgt" {
            Some(b"spid".as_slice())
        } else if namespace == P14
            && matches!(
                local.as_ref(),
                b"triggerEvt"
                    | b"playEvt"
                    | b"stopEvt"
                    | b"pauseEvt"
                    | b"resumeEvt"
                    | b"seekEvt"
                    | b"nullEvt"
            )
        {
            Some(b"objId".as_slice())
        } else if (namespace == DRAWINGML_NAMESPACE || namespace == STRICT_DRAWINGML_NAMESPACE)
            && matches!(local.as_ref(), b"stCxn" | b"endCxn")
        {
            Some(b"id".as_slice())
        } else {
            None
        };
        let mut identity = None;
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
            let attribute_local = attribute.key.local_name();
            let legacy_ole = (namespace == crate::namespace::PRESENTATIONML_NAMESPACE
                || namespace == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE)
                && local.as_ref() == b"oleObj"
                && attribute.key.as_ref() == b"spid";
            if Some(attribute.key.as_ref()) != attribute_name {
                if !legacy_ole
                    && matches!(
                        attribute_local.as_ref(),
                        b"spid" | b"objId" | b"shapeId" | b"objectId"
                    )
                {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "remove_shape",
                        "an opaque shape identity association cannot be preserved after removal",
                    ));
                }
                continue;
            }
            if identity.is_some() {
                return Err(invalid(
                    "opened-presentation shape reference repeats its identity attribute",
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            identity = Some(value.trim().parse::<u32>().map_err(|_| {
                if local.as_ref() == b"control" {
                    Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "remove_shape",
                        "legacy control shape associations cannot be resolved for removal",
                    )
                } else {
                    invalid("opened-presentation shape reference identity is invalid")
                }
            })?);
        }
        if let Some(identity) = identity {
            references.insert(identity);
        } else if attribute_name.is_some() {
            return Err(invalid(
                "opened-presentation shape reference identity is missing",
            ));
        }
    }
    Ok(references)
}

struct ShapeRemap<'a> {
    root_source_id: u32,
    root_name: &'a str,
    shape_ids: &'a HashMap<u32, u32>,
    relationships: &'a HashMap<String, String>,
    written_shape_ids: HashSet<u32>,
}

fn write_remapped_shape_start(
    output: &mut Vec<u8>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    empty: bool,
    remap: &mut ShapeRemap<'_>,
) -> Result<()> {
    let is_identity = element.local_name().as_ref() == b"cNvPr";
    let is_connection = matches!(element.local_name().as_ref(), b"stCxn" | b"endCxn");
    let mut attributes = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let decoded = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        attributes.push((attribute.key.as_ref().to_vec(), decoded.into_owned()));
    }
    let source_identity = if is_identity {
        shape_identity_from_attributes(&attributes)?
    } else {
        None
    };
    let mapped_identity = source_identity
        .map(|source_id| {
            remap
                .shape_ids
                .get(&source_id)
                .copied()
                .map(|destination_id| (source_id, destination_id))
                .ok_or_else(|| {
                    invalid(format!(
                        "opened-presentation transferred shape identity {source_id} was not remapped"
                    ))
                })
        })
        .transpose()?
        .map(|(source_id, destination_id)| {
            if !remap.written_shape_ids.insert(source_id) {
                return Err(invalid(format!(
                    "opened-presentation transferred shape repeats identity {source_id}"
                )));
            }
            Ok((source_id, destination_id))
        })
        .transpose()?;
    output.push(b'<');
    output.extend_from_slice(element.name().as_ref());
    for (key, decoded) in attributes {
        output.push(b' ');
        output.extend_from_slice(&key);
        output.extend_from_slice(b"=\"");
        let value = if let Some((_source_id, destination_id)) = mapped_identity
            && key == b"id"
        {
            destination_id.to_string()
        } else if let Some((source_id, destination_id)) = mapped_identity
            && key == b"name"
        {
            if source_id == remap.root_source_id {
                remap.root_name.to_owned()
            } else {
                format!("{decoded} Copy {destination_id}")
            }
        } else if is_connection && key == b"id" {
            let connected_id = decoded.parse::<u32>().map_err(|_err| {
                invalid("opened-presentation connector endpoint identity is invalid")
            })?;
            remap
                .shape_ids
                .get(&connected_id)
                .ok_or_else(|| {
                    Error::shape_transfer_refused(
                        crate::ShapeTransferRefusal::UnresolvedConnectorEndpoint,
                    )
                })?
                .to_string()
        } else {
            let (namespace, _) = resolver.resolve_attribute(quick_xml::name::QName(key.as_slice()));
            let is_relationship_namespace = matches!(
                &namespace,
                ResolveResult::Bound(Namespace(value))
                    if *value == litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE
                        || *value == litchi_ooxml_common::relationships::STRICT_NAMESPACE
            );
            let is_inherited_relationship_alias = matches!(
                &namespace,
                ResolveResult::Unknown(prefix)
                    if prefix.as_slice() == b"r"
                        || remap.relationships.contains_key(decoded.as_str())
            );
            if is_relationship_namespace || is_inherited_relationship_alias {
                remap
                    .relationships
                    .get(decoded.as_str())
                    .cloned()
                    .ok_or_else(|| {
                        invalid(format!(
                            "opened-presentation transferred relationship {decoded} was not remapped"
                        ))
                    })?
            } else {
                decoded
            }
        };
        output.extend_from_slice(quick_xml::escape::escape(&value).as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
    Ok(())
}

fn shape_identity_from_attributes(attributes: &[(Vec<u8>, String)]) -> Result<Option<u32>> {
    let mut identities = attributes
        .iter()
        .filter(|(key, _)| key.as_slice() == b"id")
        .map(|(_, value)| value);
    let Some(value) = identities.next() else {
        return Ok(None);
    };
    if identities.next().is_some() {
        return Err(invalid(
            "opened-presentation non-visual identity has duplicate id attributes",
        ));
    }
    value
        .parse::<u32>()
        .map(Some)
        .map_err(|_err| invalid("opened-presentation non-visual shape identity is invalid"))
}

fn slide_id_elements(
    xml: &[u8],
    capture_names: bool,
    context: Option<&litchi_core::ExecutionContext>,
) -> Result<Vec<SlideIdElement>> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut list_depth = None;
    let mut lists = 0usize;
    let mut open = None;
    let mut elements = Vec::new();
    loop {
        let start = position(&reader, origin)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let pml_namespace = is_presentation_namespace(&namespace);
        let captured_element_namespace = if capture_names
            && elements.is_empty()
            && list_depth == Some(depth)
            && matches!(&event, Event::Empty(element) if element.local_name().as_ref() == b"sldId")
        {
            match &namespace {
                ResolveResult::Bound(Namespace(value)) if pml_namespace => Some(clone_xml_name(
                    value,
                    "opened-presentation slide entry namespace",
                )?),
                _ => None,
            }
        } else {
            None
        };
        let event = event.into_owned();
        drop(namespace);
        let end = position(&reader, origin)?;
        match event {
            Event::Start(element) => {
                bump(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: depth.checked_add(1).unwrap_or(usize::MAX),
                        maximum: MAX_XML_DEPTH,
                    });
                }
                if pml_namespace && element.local_name().as_ref() == b"sldIdLst" {
                    if depth != 1 {
                        return Err(invalid(
                            "opened-presentation slide-ID list must be a direct root child",
                        ));
                    }
                    lists = lists.checked_add(1).ok_or_else(|| {
                        invalid("opened-presentation slide-ID list count overflow")
                    })?;
                    if lists != 1 {
                        return Err(invalid("opened-presentation has multiple slide-ID lists"));
                    }
                    list_depth = Some(depth + 1);
                } else if pml_namespace && element.local_name().as_ref() == b"sldId" {
                    if list_depth != Some(depth) {
                        return Err(invalid(
                            "opened-presentation slide ID is not a direct slide-list child",
                        ));
                    }
                    if capture_names {
                        return Err(invalid(
                            "slide-copy requires empty presentation slide-ID entries",
                        ));
                    }
                    if open.is_some() {
                        return Err(invalid("opened-presentation slide IDs overlap"));
                    }
                    let (id, relationship_id, _relationship_attribute_name, _) =
                        parse_slide_id(&element, &reader, false)?;
                    open = Some(OpenElement {
                        id,
                        relationship_id,
                        start,
                        depth: depth + 1,
                    });
                } else if list_depth == Some(depth) {
                    if !pml_namespace || element.local_name().as_ref() != b"sldId" {
                        return Err(invalid(
                            "opened-presentation slide-ID list has an unsupported child",
                        ));
                    }
                }
                depth += 1;
            },
            Event::Empty(element) => {
                bump(&mut nodes)?;
                if nodes % 256 == 0 {
                    check_execution_context(context)?;
                }
                if pml_namespace && element.local_name().as_ref() == b"sldIdLst" {
                    if depth != 1 || lists != 0 {
                        return Err(invalid(
                            "opened-presentation has an additional or nested slide-ID list",
                        ));
                    }
                    lists = lists.checked_add(1).ok_or_else(|| {
                        invalid("opened-presentation slide-ID list count overflow")
                    })?;
                } else if pml_namespace && element.local_name().as_ref() == b"sldId" {
                    if list_depth != Some(depth) {
                        return Err(invalid(
                            "opened-presentation slide ID is not a direct slide-list child",
                        ));
                    }
                    let capture_element_names = capture_names && elements.is_empty();
                    let (id, relationship_id, relationship_attribute_name, relationship_namespace) =
                        parse_slide_id(&element, &reader, capture_element_names)?;
                    push_slide_id_element(
                        &mut elements,
                        SlideIdElement {
                            id,
                            relationship_id,
                            span: start..end,
                            element_name: if capture_element_names {
                                Some(clone_xml_name(
                                    element.name().as_ref(),
                                    "opened-presentation slide entry name",
                                )?)
                            } else {
                                None
                            },
                            element_namespace: captured_element_namespace,
                            relationship_attribute_name,
                            relationship_namespace,
                        },
                    )?;
                } else if list_depth == Some(depth) {
                    if !pml_namespace || element.local_name().as_ref() != b"sldId" {
                        return Err(invalid(
                            "opened-presentation slide-ID list has an unsupported child",
                        ));
                    }
                }
            },
            Event::End(element) => {
                if pml_namespace && element.local_name().as_ref() == b"sldId" {
                    if open.as_ref().is_none_or(|active| active.depth != depth) {
                        return Err(invalid(
                            "opened-presentation slide ID is nested or not opened",
                        ));
                    }
                    let active = open
                        .take()
                        .ok_or_else(|| invalid("opened-presentation slide ID disappeared"))?;
                    push_slide_id_element(
                        &mut elements,
                        SlideIdElement {
                            id: active.id,
                            relationship_id: active.relationship_id,
                            span: active.start..end,
                            element_name: None,
                            element_namespace: None,
                            relationship_attribute_name: None,
                            relationship_namespace: None,
                        },
                    )?;
                }
                if pml_namespace && element.local_name().as_ref() == b"sldIdLst" {
                    if list_depth != Some(depth) {
                        return Err(invalid(
                            "opened-presentation slide-ID list is nested or not opened",
                        ));
                    }
                    list_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation XML depth underflow"))?;
            },
            Event::Text(value) if list_depth == Some(depth) => {
                if !value
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid("opened-presentation slide-ID list contains text"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if list_depth == Some(depth) => {
                return Err(invalid(
                    "opened-presentation slide-ID list contains unsupported content",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || open.is_some() || list_depth.is_some() {
        return Err(invalid("opened-presentation XML is unterminated"));
    }
    if lists != 1 {
        return Err(invalid("opened-presentation has no slide-ID list"));
    }
    check_execution_context(context)?;
    Ok(elements)
}

fn push_slide_id_element(
    elements: &mut Vec<SlideIdElement>,
    element: SlideIdElement,
) -> Result<()> {
    elements
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide ID elements",
            source,
        })?;
    elements.push(element);
    Ok(())
}

fn parse_slide_id(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    capture_relationship_name: bool,
) -> Result<(u32, String, Option<Vec<u8>>, Option<Vec<u8>>)> {
    let value =
        litchi_ooxml_common::xml::unqualified_attribute_value(element, b"id", reader.decoder())?
            .ok_or_else(|| invalid("opened-presentation slide ID lacks id"))?;
    let id = value
        .parse::<u32>()
        .map_err(|_err| invalid("opened-presentation slide ID is invalid"))?;
    let relationship_id = crate::parts::relationship_attribute(element, reader)?
        .ok_or_else(|| invalid("opened-presentation slide ID lacks r:id"))?;
    let mut relationship_attribute_name = None;
    let mut relationship_namespace = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        let resolved = reader.resolver().resolve_attribute(attribute.key).0;
        if capture_relationship_name
            && matches!(&resolved, ResolveResult::Bound(Namespace(value))
                if *value == b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    || *value == b"http://purl.oclc.org/ooxml/officeDocument/relationships")
        {
            if relationship_attribute_name
                .replace(clone_xml_name(
                    attribute.key.as_ref(),
                    "opened-presentation relationship attribute name",
                )?)
                .is_some()
            {
                return Err(invalid(
                    "opened-presentation slide ID repeats its relationship attribute",
                ));
            }
            let ResolveResult::Bound(Namespace(value)) = resolved else {
                unreachable!("relationship namespace was matched above");
            };
            relationship_namespace = Some(clone_xml_name(
                value,
                "opened-presentation relationship namespace",
            )?);
        }
    }
    if capture_relationship_name && relationship_attribute_name.is_none() {
        return Err(invalid(
            "opened-presentation slide ID relationship name is unresolved",
        ));
    }
    Ok((
        id,
        relationship_id,
        relationship_attribute_name,
        relationship_namespace,
    ))
}

fn clone_xml_name(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.extend_from_slice(value);
    Ok(output)
}

fn clone_xml_bytes(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.extend_from_slice(value);
    Ok(output)
}

fn drawing_text_elements_for_owners(
    xml: &[u8],
    owners: &[Range<usize>],
) -> Result<Vec<Vec<TextElement>>> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    spans
        .try_reserve_exact(owners.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text span groups",
            source,
        })?;
    spans.resize_with(owners.len(), Vec::new);
    let mut active: Option<(usize, usize, usize)> = None;
    // The depth of the open `a:fld`, while one is open.
    let mut field: Option<usize> = None;
    let mut owner_position = 0usize;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    loop {
        let start = position(&reader, origin)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let drawing_namespace = is_drawing(&namespace);
        // The event borrows the input slice, not the reader; only the
        // namespace holds the reader, so ending that borrow needs no copy of
        // the event.
        drop(namespace);
        let end = position(&reader, origin)?;
        match event {
            Event::Start(element) => {
                bump(&mut nodes)?;
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: depth.checked_add(1).unwrap_or(usize::MAX),
                        maximum: MAX_XML_DEPTH,
                    });
                }
                depth += 1;
                while owners
                    .get(owner_position)
                    .is_some_and(|owner| start >= owner.end)
                {
                    owner_position += 1;
                }
                let owner = owners
                    .get(owner_position)
                    .filter(|owner| owner.contains(&start));
                if owner.is_some()
                    && drawing_namespace
                    && field.is_none()
                    && element.local_name().as_ref() == b"fld"
                {
                    field = Some(depth);
                }
                if owner.is_some() && drawing_namespace && element.local_name().as_ref() == b"t" {
                    if active.replace((end, depth, owner_position)).is_some() {
                        return Err(invalid(
                            "opened-presentation DrawingML text elements overlap",
                        ));
                    }
                } else if active.is_some() {
                    return Err(invalid(
                        "opened-presentation DrawingML text contains child markup",
                    ));
                }
            },
            Event::Empty(element) => {
                bump(&mut nodes)?;
                while owners
                    .get(owner_position)
                    .is_some_and(|owner| start >= owner.end)
                {
                    owner_position += 1;
                }
                let owner = owners
                    .get(owner_position)
                    .filter(|owner| owner.contains(&start));
                let text_element =
                    owner.is_some() && drawing_namespace && element.local_name().as_ref() == b"t";
                // Nothing may sit inside an open text element, an empty `a:t`
                // included: recording one would give the rewrite a span that
                // starts inside the enclosing element's span.
                if active.is_some() {
                    return Err(invalid(if text_element {
                        "opened-presentation DrawingML text elements overlap"
                    } else {
                        "opened-presentation DrawingML text contains child markup"
                    }));
                }
                if text_element {
                    let group = spans.get_mut(owner_position).ok_or_else(|| {
                        invalid("opened-presentation shape text span owner disappeared")
                    })?;
                    group.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "opened-presentation shape text spans",
                        source,
                    })?;
                    group.push(TextElement {
                        span: start..end,
                        empty_name: Some(element.name().as_ref().to_vec()),
                        field: field.is_some(),
                    });
                }
            },
            Event::End(element) => {
                if let Some((content_start, active_depth, owner)) = active
                    && active_depth == depth
                    && drawing_namespace
                    && element.local_name().as_ref() == b"t"
                {
                    let group = spans.get_mut(owner).ok_or_else(|| {
                        invalid("opened-presentation shape text span owner disappeared")
                    })?;
                    group.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "opened-presentation shape text spans",
                        source,
                    })?;
                    group.push(TextElement {
                        span: content_start..start,
                        empty_name: None,
                        field: field.is_some(),
                    });
                    active = None;
                }
                if field == Some(depth) {
                    field = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation XML depth underflow"))?;
            },
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => {},
            Event::Comment(_) if active.is_some() => {
                return Err(invalid(
                    "opened-presentation DrawingML text contains a comment",
                ));
            },
            Event::Decl(_) | Event::PI(_) | Event::DocType(_) if active.is_some() => {
                return Err(invalid(
                    "opened-presentation DrawingML text contains forbidden markup",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || active.is_some() {
        return Err(invalid("opened-presentation slide XML is unterminated"));
    }
    Ok(spans)
}

fn is_drawing(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
    )
}

pub(super) fn is_presentation_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == crate::namespace::PRESENTATIONML_NAMESPACE
                || *value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE
    )
}

pub(super) fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("opened-presentation XML position exceeds usize"))
}

fn check_execution_context(context: Option<&litchi_core::ExecutionContext>) -> Result<()> {
    if let Some(context) = context {
        context.check().map_err(|error| {
            Error::Opc(match error {
                litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                error => litchi_opc::OpcError::Execution(error),
            })
        })?;
    }
    Ok(())
}

pub(super) fn bump(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| invalid("opened-presentation XML node count overflow"))?;
    if *nodes > MAX_XML_NODES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
            observed: *nodes,
            maximum: MAX_XML_NODES,
        });
    }
    Ok(())
}

fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x1_0000..=0x10_FFFF)
}

#[cfg(test)]
mod compaction_tests {
    use super::compact_changed_slide_xml;
    use crate::Result;

    fn assert_compacted(source: &[u8], expected: &[u8], root_unchanged: bool) -> Result<()> {
        let compacted = compact_changed_slide_xml(source)?;
        assert_eq!(compacted.xml, expected);
        assert_eq!(compacted.root_unchanged, root_unchanged);
        Ok(())
    }

    #[test]
    fn outer_declaration_whitespace_keeps_the_root_witness() -> Result<()> {
        let source = b"<?xml version=\"1.0\"?>\r\n<p:sld xmlns:p=\"urn:p\"/>\r\n";
        let expected = b"<?xml version=\"1.0\"?><p:sld xmlns:p=\"urn:p\"/>";
        assert_compacted(source, expected, true)
    }

    #[test]
    fn bom_and_empty_root_are_included_in_the_exact_byte_contract() -> Result<()> {
        let source = b"\xEF\xBB\xBF\n<p:sld xmlns:p=\"urn:p\"/>\n";
        let expected = b"\xEF\xBB\xBF<p:sld xmlns:p=\"urn:p\"/>";
        assert_compacted(source, expected, true)
    }

    #[test]
    fn outside_processing_instruction_is_copied_before_root_comparison() -> Result<()> {
        let source = b"<?litchi mode=\"x\"?>\n<p:sld xmlns:p=\"urn:p\"/>";
        let expected = b"<?litchi mode=\"x\"?><p:sld xmlns:p=\"urn:p\"/>";
        assert_compacted(source, expected, true)
    }

    #[test]
    fn outside_comment_is_copied_before_root_comparison() -> Result<()> {
        let source = b"<!--before-->\n<p:sld xmlns:p=\"urn:p\"/>\n<!--after-->";
        let expected = b"<!--before--><p:sld xmlns:p=\"urn:p\"/><!--after-->";
        assert_compacted(source, expected, true)
    }

    #[test]
    fn unknown_markup_inside_the_root_is_preserved_by_the_witness() -> Result<()> {
        let source = b"<p:sld xmlns:p=\"urn:p\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x=\"urn:x\" mc:Ignorable=\"x\"><x:future/></p:sld>";
        assert_compacted(source, source, true)
    }

    #[test]
    fn root_content_whitespace_and_attribute_quote_changes_fall_back() -> Result<()> {
        let whitespace = b"<p:sld xmlns:p=\"urn:p\"> \n <p:child/> </p:sld>";
        assert_compacted(
            whitespace,
            b"<p:sld xmlns:p=\"urn:p\"><p:child/></p:sld>",
            false,
        )?;

        let quote = b"<p:sld xmlns:p='urn:p'/>";
        assert_compacted(quote, b"<p:sld xmlns:p=\"urn:p\"/>", false)
    }

    #[test]
    fn entity_and_invalid_outside_markup_keep_the_existing_refusals() -> Result<()> {
        let entity = b"<p:sld xmlns:p=\"urn:p\">&amp;</p:sld>";
        assert_compacted(entity, entity, true)?;
        assert!(compact_changed_slide_xml(b"<p:sld xmlns:p=\"urn:p\"/>text").is_err());
        assert!(compact_changed_slide_xml(b"<!DOCTYPE p:sld><p:sld/>").is_err());
        Ok(())
    }

    #[test]
    fn invalid_root_shape_does_not_create_a_witness() {
        assert!(compact_changed_slide_xml(b"<p:sld/> <p:sld/>").is_err());
        assert!(compact_changed_slide_xml(b"<p:sld>").is_err());
    }
}

#[cfg(test)]
mod text_escape_tests {
    use super::escape_character_data;
    use crate::{Error, Result};

    #[test]
    fn character_data_preserves_xml_line_endings_with_exact_byte_admission() -> Result<()> {
        let input = "你好\r\n<&>\"'\t";
        let expected = "你好&#13;\n&lt;&amp;&gt;&quot;&apos;\t";
        assert_eq!(escape_character_data(input, expected.len())?, expected);
        assert!(matches!(
            escape_character_data(input, expected.len() - 1),
            Err(Error::Limit { .. })
        ));
        assert!(matches!(
            escape_character_data("🙂\n", 5)?,
            std::borrow::Cow::Borrowed(_)
        ));
        assert!(matches!(
            escape_character_data("🙂\n", 4),
            Err(Error::Limit { .. })
        ));
        Ok(())
    }
}

#[cfg(test)]
mod slide_binding_tests {
    use super::{insert_slide_binding, remove_slide_binding};
    use crate::{Error, Result};

    const XML: &[u8] = br#"<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>
  <p:sldId id="256" r:id="rId1" future='opaque'/><!--between--><p:sldId id='257' r:id='rId2'></p:sldId>
</p:sldIdLst></p:presentation>"#;

    const BINDINGS: [(u32, &str); 2] = [(256, "rId1"), (257, "rId2")];

    #[test]
    fn removal_retains_lexical_spans_and_inter_entry_trivia() -> Result<()> {
        let actual = remove_slide_binding(XML, BINDINGS.into_iter(), 256, None)?;
        let expected = br#"<?xml version="1.0"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>
  <!--between--><p:sldId id='257' r:id='rId2'></p:sldId>
</p:sldIdLst></p:presentation>"#;
        assert_eq!(actual, expected);
        Ok(())
    }

    #[test]
    fn removal_rejects_malformed_and_stale_complete_bindings() {
        let malformed = b"<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"></p:sldIdLst></p:presentation>";
        assert!(remove_slide_binding(malformed, BINDINGS.into_iter(), 256, None).is_err());
        assert!(
            remove_slide_binding(
                XML,
                [(256, "rId-stale"), (257, "rId2")].into_iter(),
                256,
                None,
            )
            .is_err()
        );
        assert!(remove_slide_binding(XML, BINDINGS.into_iter(), 999, None).is_err());
    }

    #[test]
    fn removal_honors_cooperative_cancellation() {
        use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits};
        use std::num::{NonZeroU64, NonZeroUsize};

        let budget = Budget::root(
            "opened-slide-removal-xml-test",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = CancellationSource::pair();
        let limits = ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MAX, 0)
            .expect("valid execution limits");
        let context = ExecutionContext::new(budget, cancellation, limits);
        cancellation_source.cancel();

        assert!(matches!(
            remove_slide_binding(XML, BINDINGS.into_iter(), 256, Some(&context)),
            Err(Error::Opc(litchi_opc::OpcError::Cancelled))
        ));
    }
    #[test]
    fn catalog_delta_proves_formatted_move_removal_and_inverse() -> Result<()> {
        let swapped = [BINDINGS[1], BINDINGS[0]];
        let moved = super::reorder_slide_bindings(XML, &BINDINGS, &swapped, None)?;
        let removed = remove_slide_binding(&moved, swapped.into_iter(), 257, None)?;
        let remaining = [BINDINGS[0]];
        super::validate_slide_catalog_delta(
            XML,
            BINDINGS.into_iter(),
            &removed,
            remaining.into_iter(),
            None,
        )?;
        super::validate_slide_catalog_delta(
            &removed,
            remaining.into_iter(),
            XML,
            BINDINGS.into_iter(),
            None,
        )?;
        Ok(())
    }

    #[test]
    fn catalog_delta_rejects_root_and_retained_binding_mutations() {
        let source = std::str::from_utf8(XML).expect("ASCII fixture");
        for changed in [
            source.replace("opaque", "forged"),
            source.replace("<!--between-->", "<!--forged-->"),
        ] {
            let relocated = source
                .replace("<!--between-->", "")
                .replace("</p:sldIdLst>", "<!--between--></p:sldIdLst>");
            assert!(
                super::validate_slide_catalog_delta(
                    XML,
                    BINDINGS.into_iter(),
                    relocated.as_bytes(),
                    BINDINGS.into_iter(),
                    None
                )
                .is_err()
            );
            assert!(
                super::validate_slide_catalog_delta(
                    XML,
                    BINDINGS.into_iter(),
                    changed.as_bytes(),
                    BINDINGS.into_iter(),
                    None,
                )
                .is_err()
            );
        }
        assert!(
            super::validate_slide_catalog_delta(
                XML,
                BINDINGS.into_iter(),
                XML,
                [BINDINGS[1], BINDINGS[0]].into_iter(),
                None,
            )
            .is_err()
        );
    }

    #[test]
    fn catalog_delta_proves_empty_list_insertion_and_inverse_for_both_dialects() -> Result<()> {
        for namespace in [
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://purl.oclc.org/ooxml/presentationml/main",
        ] {
            for list in ["<p:sldIdLst><!--keep--></p:sldIdLst>", "<p:sldIdLst/>"] {
                let source =
                    format!("<p:presentation xmlns:p='{namespace}'>{list}</p:presentation>");
                let binding = [(256, "rId9")];
                let inserted = insert_slide_binding(
                    source.as_bytes(),
                    std::iter::empty(),
                    0,
                    binding[0].0,
                    binding[0].1,
                )?;
                super::validate_slide_catalog_delta(
                    source.as_bytes(),
                    std::iter::empty(),
                    &inserted,
                    binding.into_iter(),
                    None,
                )?;
                super::validate_slide_catalog_delta(
                    &inserted,
                    binding.into_iter(),
                    source.as_bytes(),
                    std::iter::empty(),
                    None,
                )?;

                let mut forged = inserted.clone();
                let marker = b"<p:presentation";
                let start = forged
                    .windows(marker.len())
                    .position(|window| window == marker)
                    .expect("presentation root");
                forged.splice(
                    start + marker.len()..start + marker.len(),
                    b" forged='x'".iter().copied(),
                );
                assert!(
                    super::validate_slide_catalog_delta(
                        source.as_bytes(),
                        std::iter::empty(),
                        &forged,
                        binding.into_iter(),
                        None,
                    )
                    .is_err()
                );
            }
        }
        Ok(())
    }

    #[test]
    fn catalog_delta_rejects_missing_list_empty_root_forms() {
        let before = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        let after = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256" r:id="rId9"/></p:sldIdLst></p:presentation>"#;
        assert!(
            super::validate_slide_catalog_delta(
                before,
                std::iter::empty(),
                after,
                std::iter::once((256, "rId9")),
                None,
            )
            .is_err()
        );
    }
}

#[cfg(test)]
mod source_text_box_proof_tests {
    use super::{append_shape, strip_appended_text_boxes};
    use crate::media_parts::Conformance;
    use crate::{Result, opened::Limits};

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    const STRICT_DML: &str = "http://purl.oclc.org/ooxml/drawingml/main";

    fn empty_slide(strict: bool) -> Vec<u8> {
        let (pml, dml) = if strict {
            (STRICT_PML, STRICT_DML)
        } else {
            (PML, DML)
        };
        format!(
            r#"<p:sld xmlns:p="{pml}" xmlns:a="{dml}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes()
    }

    fn text_box(strict: bool, text: &str, id: u32) -> Result<String> {
        crate::opened::text_box_fragment(
            text,
            (10, 20, 30, 40),
            id,
            if strict {
                Conformance::Strict
            } else {
                Conformance::Transitional
            },
            Limits::default(),
        )
    }

    fn appended(xml: &[u8], fragment: &str) -> Result<Vec<u8>> {
        append_shape(xml, fragment.as_bytes())
    }

    #[test]
    fn repeated_canonical_suffixes_are_the_only_admitted_new_roots() -> Result<()> {
        for strict in [false, true] {
            let before = appended(&empty_slide(strict), &text_box(strict, "old", 2)?)?;
            let once = appended(&before, &text_box(strict, "one", 3)?)?;
            let after = appended(&once, &text_box(strict, "two", 4)?)?;
            assert_eq!(strip_appended_text_boxes(&before, &after)?, Some(before));
        }
        Ok(())
    }

    #[test]
    fn arbitrary_shapes_pictures_relationships_ids_and_namespaces_are_rejected() -> Result<()> {
        let before = appended(&empty_slide(false), &text_box(false, "old", 2)?)?;
        let canonical = text_box(false, "new", 3)?;

        let mut non_text_box = canonical.clone();
        non_text_box = non_text_box.replace("txBox=\"1\"", "txBox=\"0\"");
        let mismatched_id = text_box(false, "new", 4)?;
        let forged_namespace = canonical.replace(DML, STRICT_DML);
        let picture = format!(
            r#"<p:pic xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}"><p:nvPicPr><p:cNvPr id="3" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill><p:spPr/></p:pic>"#
        );

        for fragment in [non_text_box, mismatched_id, forged_namespace, picture] {
            let candidate = appended(&before, &fragment)?;
            assert!(matches!(
                strip_appended_text_boxes(&before, &candidate),
                Ok(None) | Err(_)
            ));
        }
        Ok(())
    }
}
