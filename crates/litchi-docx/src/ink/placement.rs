//! Source-bound, private selectors used by DOCX Ink authoring.
//!
//! The public Ink model deliberately does not expose XML identity.  This
//! module keeps the source ranges needed by the transaction layer behind the
//! crate boundary and uses the same MCE capability profile as the inventory
//! scanner.  It does not build host wrappers or authored XML.

#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "the bounded source selectors keep their scanner helpers together"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "quick-xml events are refined after each validation step"
)]

use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use litchi_opc::OwnedXmlPart;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::host;
use super::{Limits, xml};
use crate::package::story::StoryDialect;
use crate::{Error, Result};

const TRANSITIONAL_WORD: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_RELATIONSHIPS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const WORD_2010_WORDML: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const WORDPROCESSING_INK: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
const WORDPROCESSING_CANVAS: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
const WORDPROCESSING_GROUP: &[u8] =
    b"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";

const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_NAMESPACE_DECLARATIONS: usize = 256;
const MAX_RELATIONSHIP_ID_BYTES: usize = 1024;
const MAX_RETARGET_TARGETS: usize = 65_536;
const MAX_RETARGET_VALUE_BYTES: usize = 32 * 1024 * 1024;

/// One active paragraph source opening tag and its insertion disposition.
///
/// The range is retained even when `eligible` is false.  This is intentional:
/// an unsupported MCE fallback must not disappear from semantic paragraph
/// ordering merely because an insertion transaction cannot safely edit it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Paragraph {
    pub(crate) span: Range<usize>,
    pub(crate) eligible: bool,
}

/// Return active Word paragraph opening tags in source order.
#[cfg(test)]
pub(crate) fn paragraphs(
    xml: &[u8],
    dialect: StoryDialect,
    limits: Limits,
) -> Result<Vec<Range<usize>>> {
    let records = paragraph_records(xml, dialect, limits)?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(records.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink paragraph source spans",
            source,
        })?;
    result.extend(records.into_iter().map(|record| record.span));
    Ok(result)
}

/// Return active paragraphs together with source-preserving insertion safety.
///
/// This is the transaction-facing form of [`paragraphs`].  In particular, an
/// active paragraph under an external MCE fallback remains in the returned
/// sequence while being marked ineligible for structural insertion.
pub(crate) fn paragraph_records(
    source: &[u8],
    dialect: StoryDialect,
    limits: Limits,
) -> Result<Vec<Paragraph>> {
    let limits = limits.validate()?;
    let candidates = scan_paragraph_candidates(source, dialect, limits)?;
    if candidates.is_empty() {
        return Ok(Vec::new());
    }

    let mut offsets = Vec::new();
    offsets
        .try_reserve_exact(candidates.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink paragraph candidate offsets",
            source,
        })?;
    for candidate in &candidates {
        offsets.push(to_u32(candidate.span.start)?);
    }
    let selected = host::select_active_offsets(source, &offsets, limits)?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(selected.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink active paragraph spans",
            source,
        })?;
    let mut candidate_index = 0usize;
    for offset in selected {
        let offset = usize::try_from(offset).map_err(|error| {
            Error::Invalid(format!("paragraph source offset is invalid: {error}"))
        })?;
        while candidates
            .get(candidate_index)
            .is_some_and(|candidate| candidate.span.start < offset)
        {
            candidate_index = candidate_index.saturating_add(1);
        }
        let candidate = candidates
            .get(candidate_index)
            .filter(|candidate| candidate.span.start == offset)
            .ok_or_else(|| invalid("DOCX Ink active paragraph selection diverged"))?;
        result.push(Paragraph {
            span: candidate.span.clone(),
            eligible: candidate.eligible,
        });
        candidate_index = candidate_index.saturating_add(1);
    }
    Ok(result)
}

/// Check a caller's active paragraph spans without changing their ordinal.
///
/// The input must be exactly the result of [`paragraphs`] for the same source
/// and limits.  Returning one disposition per span prevents callers from
/// silently filtering an unsupported MCE paragraph and shifting later
/// positions.
#[cfg(test)]
pub(crate) fn paragraph_eligibility(
    source: &[u8],
    spans: &[Range<usize>],
    dialect: StoryDialect,
    limits: Limits,
) -> Result<Vec<bool>> {
    let records = paragraph_records(source, dialect, limits)?;
    if records.len() != spans.len()
        || records
            .iter()
            .zip(spans)
            .any(|(record, span)| record.span != *span)
    {
        return Err(invalid(
            "DOCX Ink paragraph source spans are stale or unordered",
        ));
    }
    let mut result = Vec::new();
    result
        .try_reserve_exact(records.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink paragraph eligibility",
            source,
        })?;
    result.extend(records.into_iter().map(|record| record.eligible));
    Ok(result)
}

/// Replace selected content-part relationship values in one source scan.
///
/// The selected ranges identify complete source elements, not arbitrary byte
/// offsets.  Expanded element and attribute names are checked before the
/// existing OPC attribute rewriter is called, so aliases, quotes, unrelated
/// attributes, and source lexical spacing remain untouched.
pub(crate) fn retarget_many(
    source: &OwnedXmlPart,
    targets: &[(Range<usize>, &str)],
    dialect: StoryDialect,
    max_output: usize,
) -> Result<OwnedXmlPart> {
    if targets.is_empty() {
        return Ok(source.clone());
    }
    if targets.len() > MAX_RETARGET_TARGETS {
        return Err(limit(
            "Ink retarget targets",
            targets.len(),
            MAX_RETARGET_TARGETS,
        ));
    }
    if max_output == 0 {
        return Err(limit("Ink retarget output bytes", 0, 1));
    }
    let bytes = source.bytes();
    let mut previous_end = 0usize;
    for (span, relationship_id) in targets {
        if span.start < previous_end || span.start >= span.end || span.end > bytes.len() {
            return Err(invalid("DOCX Ink retarget spans are invalid or unordered"));
        }
        validate_relationship_id(relationship_id)?;
        previous_end = span.end;
    }

    let mut reader = reader(bytes);
    let origin = ReaderOrigin::of(bytes);
    let mut stack_depth = 0usize;
    let mut target_index = 0usize;
    let mut active_target: Option<ActiveRetarget> = None;
    let mut replacement_bytes = 0usize;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(targets.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink retarget attribute edits",
            source,
        })?;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_end = position(origin, reader.buffer_position())?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let (start, _start_end) = opening_span(event_end, &element, false, bytes)?;
                if let Some(active) = active_target.as_ref() {
                    if stack_depth > active.depth {
                        return Err(invalid(
                            "DOCX Ink retarget contentPart contains child content",
                        ));
                    }
                }
                if target_index < targets.len() && targets[target_index].0.start == start {
                    let target = &targets[target_index].0;
                    let attribute = selected_relationship_attribute(
                        &element, &namespace, resolver, dialect, bytes,
                    )?;
                    active_target = Some(ActiveRetarget {
                        index: target_index,
                        depth: stack_depth,
                        end: target.end,
                        value_range: attribute.value_range,
                    });
                    target_index = target_index.saturating_add(1);
                } else if target_index < targets.len() && targets[target_index].0.start < start {
                    return Err(invalid(
                        "DOCX Ink retarget span does not identify an element",
                    ));
                }
                stack_depth = stack_depth
                    .checked_add(1)
                    .ok_or_else(|| limit("Ink retarget XML depth", usize::MAX, usize::MAX))?;
            },
            Event::Empty(element) => {
                let (start, end) = opening_span(event_end, &element, true, bytes)?;
                if active_target.is_some() {
                    return Err(invalid(
                        "DOCX Ink retarget contentPart contains child content",
                    ));
                }
                if target_index < targets.len() && targets[target_index].0.start == start {
                    let target = &targets[target_index].0;
                    if target.end != end {
                        return Err(invalid(
                            "DOCX Ink retarget span does not cover the complete contentPart",
                        ));
                    }
                    let attribute = selected_relationship_attribute(
                        &element, &namespace, resolver, dialect, bytes,
                    )?;
                    let index = target_index;
                    target_index = target_index.saturating_add(1);
                    add_retarget_edit(
                        &mut edits,
                        attribute.value_range,
                        targets[index].1,
                        &mut replacement_bytes,
                    )?;
                } else if target_index < targets.len() && targets[target_index].0.start < start {
                    return Err(invalid(
                        "DOCX Ink retarget span does not identify an element",
                    ));
                }
            },
            Event::End(element) => {
                let end = event_end;
                stack_depth = stack_depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("DOCX Ink retarget XML depth underflow"))?;
                if let Some(active) = active_target.take() {
                    if active.depth != stack_depth || active.end != end {
                        return Err(invalid(
                            "DOCX Ink retarget span does not cover the complete contentPart",
                        ));
                    }
                    add_retarget_edit(
                        &mut edits,
                        active.value_range,
                        targets[active.index].1,
                        &mut replacement_bytes,
                    )?;
                }
                xml::text(element.name().as_ref())?;
            },
            Event::Text(text) => {
                xml::text(text.as_ref())?;
                if active_target.is_some() {
                    return Err(invalid("DOCX Ink retarget contentPart contains child text"));
                }
            },
            Event::CData(text) => {
                xml::text(text.as_ref())?;
                if active_target.is_some() {
                    return Err(invalid(
                        "DOCX Ink retarget contentPart contains child CDATA",
                    ));
                }
            },
            Event::GeneralRef(reference) => {
                if active_target.is_some() {
                    return Err(invalid(
                        "DOCX Ink retarget contentPart contains a child reference",
                    ));
                }
                xml::reference(reference.as_ref())?;
            },
            Event::Comment(comment) => {
                if active_target.is_some() {
                    return Err(invalid(
                        "DOCX Ink retarget contentPart contains a child comment",
                    ));
                }
                xml::text(comment.as_ref())?;
            },
            Event::Decl(declaration) => xml::declaration(&declaration)?,
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DOCX Ink retarget source rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => break,
        }
    }
    if active_target.is_some() || target_index != targets.len() || stack_depth != 0 {
        return Err(invalid("DOCX Ink retarget target inventory is incomplete"));
    }

    let mut output_size = bytes.len();
    let mut changed = false;
    for (range, value) in &edits {
        output_size = output_size
            .checked_sub(range.len())
            .and_then(|size| size.checked_add(value.len()))
            .ok_or_else(|| invalid("DOCX Ink retarget output size overflowed"))?;
        changed |= bytes.get(range.clone()) != Some(value.as_slice());
    }
    if output_size > max_output {
        return Err(limit("Ink retarget output bytes", output_size, max_output));
    }
    if !changed {
        return Ok(source.clone());
    }
    Ok(source.replace_attributes(&edits)?)
}

/// Retarget the one supported VML image relationship in each selected
/// AlternateContent fallback while retaining every other source byte.
///
/// The host ranges are complete, admitted drawing hosts captured by
/// [`super::host::capture`].  This helper repeats the bounded XML and direct
/// Choice/Fallback checks while locating the image attributes, so a stale or
/// malformed range cannot cause an unrelated relationship to be rewritten.
pub(crate) fn retarget_fallbacks(
    source: &OwnedXmlPart,
    targets: &[(Range<usize>, &str)],
    dialect: StoryDialect,
    limits: Limits,
) -> Result<OwnedXmlPart> {
    let limits = limits.validate()?;
    if targets.is_empty() {
        return Ok(source.clone());
    }
    if targets.len() > MAX_RETARGET_TARGETS {
        return Err(limit(
            "Ink fallback retarget targets",
            targets.len(),
            MAX_RETARGET_TARGETS,
        ));
    }
    let bytes = source.bytes();
    if bytes.len() > MAX_XML_BYTES {
        return Err(limit("story XML bytes", bytes.len(), MAX_XML_BYTES));
    }
    if bytes.len() > limits.stories.max_story_bytes {
        return Err(limit(
            "story XML bytes",
            bytes.len(),
            limits.stories.max_story_bytes,
        ));
    }

    let mut previous_end = 0usize;
    for (span, relationship_id) in targets {
        if span.start < previous_end || span.start >= span.end || span.end > bytes.len() {
            return Err(invalid(
                "DOCX Ink fallback retarget spans are invalid or unordered",
            ));
        }
        validate_relationship_id(relationship_id)?;
        previous_end = span.end;
    }

    let mut reader = reader(bytes);
    let origin = ReaderOrigin::of(bytes);
    let mut stack: Vec<FallbackFrame> = Vec::new();
    stack
        .try_reserve(limits.max_xml_depth)
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink fallback XML frames",
            source,
        })?;
    let mut states = Vec::new();
    states
        .try_reserve_exact(targets.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink fallback target states",
            source,
        })?;
    let mut target_index = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut prolog_started = false;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(targets.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink fallback attribute edits",
            source,
        })?;
    let mut replacement_bytes = 0usize;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_end = position(origin, reader.buffer_position())?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                observe_envelope_start(
                    &element,
                    &namespace,
                    resolver,
                    &mut stack,
                    &mut nodes,
                    limits,
                    &mut root_seen,
                    root_closed,
                    &mut prolog_started,
                )?;
                let (start, _start_end) = opening_span(event_end, &element, false, bytes)?;
                let kind = classify_fallback_node(&namespace, &element);
                let parent = stack.last().copied();
                let host_id = if target_index < targets.len()
                    && targets[target_index].0.start == start
                {
                    if kind != FallbackNodeKind::Alternate {
                        return Err(invalid(
                            "DOCX Ink fallback retarget range does not identify AlternateContent",
                        ));
                    }
                    let id = states.len();
                    states.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "DOCX Ink fallback target states",
                        source,
                    })?;
                    states.push(FallbackTargetState {
                        start,
                        end: targets[target_index].0.end,
                        choice_count: 0,
                        fallback_count: 0,
                        other_count: 0,
                        fallback: None,
                        image_elements: 0,
                        image_attribute: None,
                        non_whitespace_text: false,
                    });
                    target_index = target_index.saturating_add(1);
                    Some(id)
                } else if target_index < targets.len() && targets[target_index].0.start < start {
                    return Err(invalid(
                        "DOCX Ink fallback retarget range does not identify an element",
                    ));
                } else {
                    None
                };

                if kind == FallbackNodeKind::Fallback
                    && parent.is_some_and(|frame| frame.kind == FallbackNodeKind::Alternate)
                {
                    let parent = parent.expect("checked parent");
                    if let Some(id) = parent.host_id {
                        let state = states
                            .get_mut(id)
                            .ok_or_else(|| invalid("DOCX Ink fallback target state is missing"))?;
                        if state.fallback.is_some() {
                            return Err(invalid(
                                "DOCX Ink fallback retarget host contains multiple direct Fallback elements",
                            ));
                        }
                        state.fallback_count = state.fallback_count.saturating_add(1);
                        state.fallback =
                            Some((start, opening_span(event_end, &element, false, bytes)?.1));
                    }
                } else if kind == FallbackNodeKind::Choice
                    && parent.is_some_and(|frame| frame.kind == FallbackNodeKind::Alternate)
                {
                    let parent = parent.expect("checked parent");
                    if let Some(id) = parent.host_id {
                        let state = states
                            .get_mut(id)
                            .ok_or_else(|| invalid("DOCX Ink fallback target state is missing"))?;
                        state.choice_count = state.choice_count.saturating_add(1);
                        if state.choice_count > 1 {
                            return Err(invalid(
                                "DOCX Ink fallback retarget host contains multiple direct Choice elements",
                            ));
                        }
                    }
                } else if parent.is_some_and(|frame| frame.kind == FallbackNodeKind::Alternate)
                    && host_id.is_none()
                {
                    if let Some(id) = parent.and_then(|frame| frame.host_id) {
                        let state = states
                            .get_mut(id)
                            .ok_or_else(|| invalid("DOCX Ink fallback target state is missing"))?;
                        state.other_count = state.other_count.saturating_add(1);
                    }
                }

                let fallback_host_id = if kind == FallbackNodeKind::Fallback
                    && parent.is_some_and(|frame| frame.kind == FallbackNodeKind::Alternate)
                {
                    parent.and_then(|frame| frame.host_id)
                } else {
                    parent.and_then(|frame| frame.fallback_host_id)
                };
                if kind == FallbackNodeKind::ImageData {
                    if let Some(id) = fallback_host_id {
                        let state = states
                            .get_mut(id)
                            .ok_or_else(|| invalid("DOCX Ink fallback target state is missing"))?;
                        state.image_elements = state.image_elements.saturating_add(1);
                        if state.image_elements > 1 {
                            return Err(invalid(
                                "DOCX Ink fallback contains multiple VML image elements",
                            ));
                        }
                        state.image_attribute =
                            fallback_image_attribute(&element, resolver, bytes, dialect)?;
                    }
                }
                stack.push(FallbackFrame {
                    kind,
                    start,
                    host_id,
                    fallback_host_id,
                    is_word_run: is_word_run(&namespace, &element),
                });
            },
            Event::Empty(element) => {
                observe_envelope_empty(
                    &element,
                    &namespace,
                    resolver,
                    &mut stack,
                    &mut nodes,
                    limits,
                    &mut root_seen,
                    &mut root_closed,
                    &mut prolog_started,
                )?;
                let (start, _end) = opening_span(event_end, &element, true, bytes)?;
                let kind = classify_fallback_node(&namespace, &element);
                let parent = stack.last().copied();
                if target_index < targets.len() && targets[target_index].0.start == start {
                    return Err(invalid(
                        "DOCX Ink fallback retarget range must cover a complete AlternateContent",
                    ));
                }
                if kind == FallbackNodeKind::Choice
                    && parent.is_some_and(|frame| frame.kind == FallbackNodeKind::Alternate)
                {
                    return Err(invalid("DOCX Ink fallback Choice must be complete"));
                }
                if kind == FallbackNodeKind::Fallback
                    && parent.is_some_and(|frame| frame.kind == FallbackNodeKind::Alternate)
                {
                    return Err(invalid("DOCX Ink fallback Fallback must be complete"));
                }
                if parent.is_some_and(|frame| frame.kind == FallbackNodeKind::Alternate) {
                    if let Some(id) = parent.and_then(|frame| frame.host_id) {
                        let state = states
                            .get_mut(id)
                            .ok_or_else(|| invalid("DOCX Ink fallback target state is missing"))?;
                        match kind {
                            FallbackNodeKind::Choice => {
                                state.choice_count = state.choice_count.saturating_add(1);
                            },
                            FallbackNodeKind::Fallback => {
                                state.fallback_count = state.fallback_count.saturating_add(1);
                            },
                            _ => state.other_count = state.other_count.saturating_add(1),
                        }
                    }
                }
                let fallback_host_id = if kind == FallbackNodeKind::Fallback
                    && parent.is_some_and(|frame| frame.kind == FallbackNodeKind::Alternate)
                {
                    parent.and_then(|frame| frame.host_id)
                } else {
                    parent.and_then(|frame| frame.fallback_host_id)
                };
                if kind == FallbackNodeKind::ImageData {
                    if let Some(id) = fallback_host_id {
                        let state = states
                            .get_mut(id)
                            .ok_or_else(|| invalid("DOCX Ink fallback target state is missing"))?;
                        state.image_elements = state.image_elements.saturating_add(1);
                        if state.image_elements > 1 {
                            return Err(invalid(
                                "DOCX Ink fallback contains multiple VML image elements",
                            ));
                        }
                        state.image_attribute =
                            fallback_image_attribute(&element, resolver, bytes, dialect)?;
                    }
                }
            },
            Event::End(element) => {
                let end = event_end;
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("DOCX Ink fallback source has an unexpected end"))?;
                xml::text(element.name().as_ref())?;
                if let Some(id) = frame.host_id {
                    let state = states
                        .get(id)
                        .ok_or_else(|| invalid("DOCX Ink fallback target state is missing"))?;
                    if state.end != end || state.start != frame.start {
                        return Err(invalid(
                            "DOCX Ink fallback retarget range does not cover complete AlternateContent",
                        ));
                    }
                    let parent_is_word_run = stack.last().is_some_and(|parent| parent.is_word_run);
                    if !parent_is_word_run
                        || state.choice_count != 1
                        || state.fallback_count != 1
                        || state.other_count != 0
                        || state.non_whitespace_text
                        || state.fallback.is_none()
                        || state.image_elements != 1
                        || state.image_attribute.is_none()
                    {
                        return Err(invalid(
                            "DOCX Ink fallback host has no single supported image reference",
                        ));
                    }
                    add_retarget_edit(
                        &mut edits,
                        state
                            .image_attribute
                            .clone()
                            .expect("checked image attribute"),
                        targets[id].1,
                        &mut replacement_bytes,
                    )?;
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                xml::text(text.as_ref())?;
                if stack.is_empty() {
                    if (!root_seen || root_closed)
                        && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                    {
                        return Err(invalid(
                            "DOCX source has non-whitespace text outside its root",
                        ));
                    }
                } else if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    if let Some(frame) = stack.last() {
                        if let Some(id) = frame.host_id {
                            let state = states.get_mut(id).ok_or_else(|| {
                                invalid("DOCX Ink fallback target state is missing")
                            })?;
                            if frame.kind == FallbackNodeKind::Alternate {
                                state.non_whitespace_text = true;
                            }
                        }
                    }
                }
                if !root_seen {
                    prolog_started = true;
                }
            },
            Event::CData(data) => {
                xml::text(data.as_ref())?;
                if stack.is_empty() || !root_seen || root_closed {
                    return Err(invalid("DOCX source has CDATA outside its root"));
                }
            },
            Event::Comment(comment) => {
                xml::text(comment.as_ref())?;
                if !root_seen {
                    prolog_started = true;
                }
            },
            Event::GeneralRef(reference) => {
                xml::reference(reference.as_ref())?;
                if stack.is_empty() || !root_seen || root_closed {
                    return Err(invalid("DOCX source has a reference outside its root"));
                }
            },
            Event::Decl(declaration) => {
                xml::declaration(&declaration)?;
                if declaration_seen || prolog_started || root_seen {
                    return Err(invalid(
                        "DOCX source has an XML declaration outside its prolog",
                    ));
                }
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DOCX source rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || !stack.is_empty() || target_index != targets.len()
                {
                    return Err(invalid("DOCX Ink fallback target inventory is incomplete"));
                }
                break;
            },
        }
    }

    let mut output_size = bytes.len();
    let mut changed = false;
    for (range, value) in &edits {
        output_size = output_size
            .checked_sub(range.len())
            .and_then(|size| size.checked_add(value.len()))
            .ok_or_else(|| invalid("DOCX Ink fallback output size overflowed"))?;
        changed |= bytes.get(range.clone()) != Some(value.as_slice());
    }
    let maximum = limits.stories.max_story_bytes.min(MAX_XML_BYTES);
    if output_size > maximum {
        return Err(limit("story XML bytes", output_size, maximum));
    }
    if !changed {
        return Ok(source.clone());
    }
    Ok(source.replace_attributes(&edits)?)
}

/// Retarget one source content-part relationship value.
#[cfg(test)]
pub(crate) fn retarget(
    source: &OwnedXmlPart,
    anchor_span: Range<usize>,
    dialect: StoryDialect,
    relationship_id: &str,
    max_output: usize,
) -> Result<OwnedXmlPart> {
    retarget_many(
        source,
        &[(anchor_span, relationship_id)],
        dialect,
        max_output,
    )
}

/// Return the opening tag of the one admitted `mc:Fallback` under a complete
/// AlternateContent host.  The wrapper itself is never rebuilt here.
#[cfg(test)]
pub(crate) fn fallback_start(
    source: &[u8],
    host_range: Range<usize>,
    limits: Limits,
) -> Result<Range<usize>> {
    let limits = limits.validate()?;
    if source.len() > MAX_XML_BYTES {
        return Err(limit("story XML bytes", source.len(), MAX_XML_BYTES));
    }
    if host_range.start >= host_range.end || host_range.end > source.len() {
        return Err(invalid("DOCX Ink AlternateContent host range is invalid"));
    }

    let mut reader = reader(source);
    let origin = ReaderOrigin::of(source);
    let mut stack: Vec<HostFrame> = Vec::new();
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut prolog_started = false;
    let mut found = None;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_end = position(origin, reader.buffer_position())?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                observe_envelope_start(
                    &element,
                    &namespace,
                    resolver,
                    &mut stack,
                    &mut nodes,
                    limits,
                    &mut root_seen,
                    root_closed,
                    &mut prolog_started,
                )?;
                let (start, start_end) = opening_span(event_end, &element, false, source)?;
                let kind = classify_host_kind(&namespace, &element);
                let parent = stack.last().copied();
                if kind == HostKind::Fallback
                    && parent.is_some_and(|frame| frame.kind == HostKind::Alternate)
                {
                    let host = parent.expect("checked parent");
                    if host.fallback.is_some() {
                        return Err(invalid(
                            "DOCX Ink AlternateContent contains multiple direct Fallback elements",
                        ));
                    }
                    let host_index = stack.len().saturating_sub(1);
                    stack[host_index].fallback = Some((start, start_end));
                }
                stack.push(HostFrame::new(
                    kind,
                    start,
                    is_word_run(&namespace, &element),
                ));
            },
            Event::Empty(element) => {
                observe_envelope_empty(
                    &element,
                    &namespace,
                    resolver,
                    &mut stack,
                    &mut nodes,
                    limits,
                    &mut root_seen,
                    &mut root_closed,
                    &mut prolog_started,
                )?;
                let (start, _end) = opening_span(event_end, &element, true, source)?;
                let kind = classify_host_kind(&namespace, &element);
                let parent = stack.last().copied();
                if kind == HostKind::Fallback
                    && parent.is_some_and(|frame| frame.kind == HostKind::Alternate)
                {
                    return Err(invalid("DOCX Ink Fallback must be a complete element"));
                }
                if kind == HostKind::Choice
                    && parent.is_some_and(|frame| frame.kind == HostKind::Alternate)
                {
                    return Err(invalid("DOCX Ink Choice must be a complete element"));
                }
                if start == host_range.start {
                    return Err(invalid(
                        "DOCX Ink host range does not cover an AlternateContent element",
                    ));
                }
                if let Some(parent) = stack.last_mut() {
                    if parent.kind == HostKind::Alternate {
                        parent.observe_child(kind);
                    }
                }
            },
            Event::End(element) => {
                let end = event_end;
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("DOCX Ink source has an unexpected end element"))?;
                xml::text(element.name().as_ref())?;
                if frame.kind == HostKind::Alternate && frame.start == host_range.start {
                    if end != host_range.end {
                        return Err(invalid(
                            "DOCX Ink host range does not cover the complete AlternateContent element",
                        ));
                    }
                    if frame.choice_count != 1
                        || frame.fallback_count != 1
                        || frame.other_count != 0
                        || frame.non_whitespace_text
                        || frame.fallback.is_none()
                    {
                        return Err(invalid(
                            "DOCX Ink AlternateContent host is not a complete Choice/Fallback pair",
                        ));
                    }
                    if !stack.last().is_some_and(|parent| parent.is_word_run) {
                        return Err(invalid(
                            "DOCX Ink AlternateContent host is not directly under a Word run",
                        ));
                    }
                    found = frame.fallback.map(|(start, end)| start..end);
                }
                if let Some(parent) = stack.last_mut() {
                    parent.observe_child(frame.kind);
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                xml::text(text.as_ref())?;
                if stack.is_empty() {
                    if !root_seen || root_closed {
                        if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                            return Err(invalid(
                                "DOCX source has non-whitespace text outside its root",
                            ));
                        }
                    }
                } else if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    if let Some(frame) = stack.last_mut() {
                        frame.non_whitespace_text = true;
                    }
                }
                if !root_seen {
                    prolog_started = true;
                }
            },
            Event::CData(data) => {
                xml::text(data.as_ref())?;
                if stack.is_empty() || !root_seen || root_closed {
                    return Err(invalid("DOCX source has CDATA outside its root"));
                }
                if let Some(frame) = stack.last_mut() {
                    frame.non_whitespace_text = true;
                }
            },
            Event::Comment(comment) => {
                xml::text(comment.as_ref())?;
                if !root_seen {
                    prolog_started = true;
                }
            },
            Event::GeneralRef(reference) => {
                xml::reference(reference.as_ref())?;
                if stack.is_empty() || !root_seen || root_closed {
                    return Err(invalid("DOCX source has a reference outside its root"));
                }
                if let Some(frame) = stack.last_mut() {
                    frame.non_whitespace_text = true;
                }
            },
            Event::Decl(declaration) => {
                xml::declaration(&declaration)?;
                if declaration_seen || prolog_started || root_seen {
                    return Err(invalid(
                        "DOCX source has an XML declaration outside its prolog",
                    ));
                }
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DOCX source rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || !stack.is_empty() {
                    return Err(invalid("DOCX source has an unterminated XML root"));
                }
                break;
            },
        }
    }
    found.ok_or_else(|| invalid("DOCX Ink AlternateContent host range was not found"))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum MceKind {
    Alternate,
    Choice,
    Fallback,
    Other,
}

#[derive(Clone, Debug)]
struct ParagraphCandidate {
    span: Range<usize>,
    eligible: bool,
}

fn scan_paragraph_candidates(
    source: &[u8],
    dialect: StoryDialect,
    limits: Limits,
) -> Result<Vec<ParagraphCandidate>> {
    if source.len() > MAX_XML_BYTES {
        return Err(limit("story XML bytes", source.len(), MAX_XML_BYTES));
    }
    let mut reader = reader(source);
    let origin = ReaderOrigin::of(source);
    let mut stack: Vec<ScanFrame> = Vec::new();
    let mut candidates = Vec::new();
    candidates
        .try_reserve(limits.max_xml_nodes.min(1024))
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink paragraph candidates",
            source,
        })?;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut prolog_started = false;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event_position = position(origin, reader.buffer_position())?;
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                observe_envelope_start(
                    &element,
                    &namespace,
                    resolver,
                    &mut stack,
                    &mut nodes,
                    limits,
                    &mut root_seen,
                    root_closed,
                    &mut prolog_started,
                )?;
                let (start, end) = opening_span(event_position, &element, false, source)?;
                let kind = classify_mce(&namespace, &element);
                if is_paragraph(&namespace, &element, dialect) {
                    push_paragraph(
                        &mut candidates,
                        start..end,
                        paragraph_is_eligible(&stack),
                        limits.max_xml_nodes,
                    )?;
                }
                stack.push(ScanFrame {
                    mce: kind,
                    choice_supported: kind != MceKind::Choice
                        || choice_requires_supported(&element, resolver),
                });
            },
            Event::Empty(element) => {
                observe_envelope_empty(
                    &element,
                    &namespace,
                    resolver,
                    &mut stack,
                    &mut nodes,
                    limits,
                    &mut root_seen,
                    &mut root_closed,
                    &mut prolog_started,
                )?;
                let (start, end) = opening_span(event_position, &element, true, source)?;
                if is_paragraph(&namespace, &element, dialect) {
                    push_paragraph(
                        &mut candidates,
                        start..end,
                        paragraph_is_eligible(&stack),
                        limits.max_xml_nodes,
                    )?;
                }
            },
            Event::End(element) => {
                if stack.pop().is_none() {
                    return Err(invalid("DOCX source has an unexpected end element"));
                }
                xml::text(element.name().as_ref())?;
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                xml::text(text.as_ref())?;
                observe_outside_text(
                    &stack,
                    &mut root_seen,
                    root_closed,
                    &mut prolog_started,
                    text.as_ref(),
                )?;
            },
            Event::CData(data) => {
                xml::text(data.as_ref())?;
                if stack.is_empty() || !root_seen || root_closed {
                    return Err(invalid("DOCX source has CDATA outside its root"));
                }
            },
            Event::Comment(comment) => {
                xml::text(comment.as_ref())?;
                if !root_seen {
                    prolog_started = true;
                }
            },
            Event::GeneralRef(reference) => {
                xml::reference(reference.as_ref())?;
                if stack.is_empty() || !root_seen || root_closed {
                    return Err(invalid("DOCX source has a reference outside its root"));
                }
            },
            Event::Decl(declaration) => {
                xml::declaration(&declaration)?;
                if declaration_seen || prolog_started || root_seen {
                    return Err(invalid(
                        "DOCX source has an XML declaration outside its prolog",
                    ));
                }
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DOCX source rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || !stack.is_empty() {
                    return Err(invalid("DOCX source has an unterminated XML root"));
                }
                break;
            },
        }
    }
    Ok(candidates)
}

#[derive(Clone, Copy, Debug)]
struct ScanFrame {
    mce: MceKind,
    choice_supported: bool,
}

fn paragraph_is_eligible(stack: &[ScanFrame]) -> bool {
    !stack.iter().any(|frame| {
        frame.mce == MceKind::Fallback || (frame.mce == MceKind::Choice && !frame.choice_supported)
    })
}

fn push_paragraph(
    candidates: &mut Vec<ParagraphCandidate>,
    span: Range<usize>,
    eligible: bool,
    maximum: usize,
) -> Result<()> {
    if candidates.len() >= maximum {
        return Err(limit(
            "DOCX Ink paragraph candidates",
            candidates.len().saturating_add(1),
            maximum,
        ));
    }
    candidates
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink paragraph candidates",
            source,
        })?;
    candidates.push(ParagraphCandidate { span, eligible });
    Ok(())
}

fn is_paragraph(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    dialect: StoryDialect,
) -> bool {
    element.local_name().as_ref() == b"p"
        && is_namespace(namespace, dialect_word_namespace(dialect))
}

fn classify_mce(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> MceKind {
    if !is_namespace(namespace, MCE_NAMESPACE) {
        return MceKind::Other;
    }
    match element.local_name().as_ref() {
        b"AlternateContent" => MceKind::Alternate,
        b"Choice" => MceKind::Choice,
        b"Fallback" => MceKind::Fallback,
        _ => MceKind::Other,
    }
}

fn choice_requires_supported(element: &BytesStart<'_>, resolver: &NamespaceResolver) -> bool {
    let Some(attribute) = element
        .attributes()
        .filter_map(|attribute| attribute.ok())
        .find(|attribute| {
            attribute.key.local_name().as_ref() == b"Requires"
                && matches!(
                    resolver.resolve_attribute(attribute.key).0,
                    ResolveResult::Unbound
                )
        })
    else {
        return false;
    };
    let Ok(value) =
        attribute.decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
    else {
        return false;
    };
    let mut found = false;
    for token in value.split_whitespace() {
        let qname = QName(token.as_bytes());
        let Some(prefix) = qname.prefix() else {
            return false;
        };
        let ResolveResult::Bound(Namespace(namespace)) =
            resolver.resolve_prefix(Some(prefix), false)
        else {
            return false;
        };
        if namespace != WORD_2010_WORDML
            && namespace != WORDPROCESSING_INK
            && namespace != WORDPROCESSING_CANVAS
            && namespace != WORDPROCESSING_GROUP
        {
            return false;
        }
        found = true;
    }
    found
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FallbackNodeKind {
    Alternate,
    Choice,
    Fallback,
    ImageData,
    Other,
}

#[derive(Clone, Debug)]
struct FallbackTargetState {
    start: usize,
    end: usize,
    choice_count: usize,
    fallback_count: usize,
    other_count: usize,
    fallback: Option<(usize, usize)>,
    image_elements: usize,
    image_attribute: Option<Range<usize>>,
    non_whitespace_text: bool,
}

#[derive(Clone, Copy, Debug)]
struct FallbackFrame {
    kind: FallbackNodeKind,
    start: usize,
    host_id: Option<usize>,
    fallback_host_id: Option<usize>,
    is_word_run: bool,
}

fn classify_fallback_node(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> FallbackNodeKind {
    if is_namespace(namespace, MCE_NAMESPACE) {
        return match element.local_name().as_ref() {
            b"AlternateContent" => FallbackNodeKind::Alternate,
            b"Choice" => FallbackNodeKind::Choice,
            b"Fallback" => FallbackNodeKind::Fallback,
            _ => FallbackNodeKind::Other,
        };
    }
    if is_namespace(namespace, VML_NAMESPACE) && element.local_name().as_ref() == b"imagedata" {
        FallbackNodeKind::ImageData
    } else {
        FallbackNodeKind::Other
    }
}

fn fallback_image_attribute(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    source: &[u8],
    dialect: StoryDialect,
) -> Result<Option<Range<usize>>> {
    let mut selected = None;
    for raw_attribute in element.attributes().with_checks(true) {
        let attribute = raw_attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let supported =
            local.as_ref() == b"id" && is_namespace(&namespace, relationship_namespace(dialect));
        if !supported {
            continue;
        }
        if selected.is_some() {
            return Err(invalid(
                "DOCX Ink fallback image has multiple supported relationship attributes",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.is_empty()
            || value.len() > MAX_RELATIONSHIP_ID_BYTES
            || value
                .chars()
                .any(|character| character.is_control() || character.is_whitespace())
        {
            return Err(invalid(
                "DOCX Ink fallback image has an invalid relationship ID",
            ));
        }
        let start = source_offset(attribute.value.as_ref(), source)?;
        let end = start
            .checked_add(attribute.value.len())
            .ok_or_else(|| invalid("DOCX Ink fallback image attribute range overflowed"))?;
        selected = Some(start..end);
    }
    Ok(selected)
}

#[cfg(test)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum HostKind {
    Alternate,
    Choice,
    Fallback,
    Other,
}

#[cfg(test)]
#[derive(Clone, Copy, Debug)]
struct HostFrame {
    kind: HostKind,
    start: usize,
    choice_count: usize,
    fallback_count: usize,
    other_count: usize,
    fallback: Option<(usize, usize)>,
    non_whitespace_text: bool,
    is_word_run: bool,
}

#[cfg(test)]
impl HostFrame {
    fn new(kind: HostKind, start: usize, is_word_run: bool) -> Self {
        Self {
            kind,
            start,
            choice_count: 0,
            fallback_count: 0,
            other_count: 0,
            fallback: None,
            non_whitespace_text: false,
            is_word_run,
        }
    }

    fn observe_child(&mut self, child: HostKind) {
        if self.kind != HostKind::Alternate {
            return;
        }
        match child {
            HostKind::Choice => self.choice_count = self.choice_count.saturating_add(1),
            HostKind::Fallback => self.fallback_count = self.fallback_count.saturating_add(1),
            HostKind::Other | HostKind::Alternate => {
                self.other_count = self.other_count.saturating_add(1)
            },
        }
    }
}

#[cfg(test)]
fn classify_host_kind(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> HostKind {
    if !is_namespace(namespace, MCE_NAMESPACE) {
        return HostKind::Other;
    }
    match element.local_name().as_ref() {
        b"AlternateContent" => HostKind::Alternate,
        b"Choice" => HostKind::Choice,
        b"Fallback" => HostKind::Fallback,
        _ => HostKind::Other,
    }
}

fn is_word_run(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    element.local_name().as_ref() == b"r"
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == TRANSITIONAL_WORD || *value == STRICT_WORD)
}

#[derive(Clone, Debug)]
struct ActiveRetarget {
    index: usize,
    depth: usize,
    end: usize,
    value_range: Range<usize>,
}

#[derive(Clone, Debug)]
struct SelectedAttribute {
    value_range: Range<usize>,
}

fn selected_relationship_attribute(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    dialect: StoryDialect,
    source: &[u8],
) -> Result<SelectedAttribute> {
    let expected_element = is_namespace(namespace, dialect_word_namespace(dialect))
        || is_namespace(namespace, WORD_2010_WORDML);
    if element.local_name().as_ref() != b"contentPart" || !expected_element {
        return Err(invalid(
            "DOCX Ink retarget target is not a supported expanded contentPart",
        ));
    }
    xml::element(element, resolver)?;
    let relationship_namespace = if is_namespace(namespace, WORD_2010_WORDML) {
        TRANSITIONAL_RELATIONSHIPS
    } else {
        relationship_namespace(dialect)
    };
    let mut selected = None;
    for raw_attribute in element.attributes().with_checks(true) {
        let attribute = raw_attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (attribute_namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() != b"id" || !is_namespace(&attribute_namespace, relationship_namespace) {
            continue;
        }
        if selected.is_some() {
            return Err(invalid(
                "DOCX Ink retarget target has duplicate expanded r:id",
            ));
        }
        let start = source_offset(attribute.value.as_ref(), source)?;
        let end = start
            .checked_add(attribute.value.len())
            .ok_or_else(|| invalid("DOCX Ink retarget attribute range overflowed"))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.is_empty()
            || value.len() > MAX_RELATIONSHIP_ID_BYTES
            || value
                .chars()
                .any(|character| character.is_control() || character.is_whitespace())
        {
            return Err(invalid(
                "DOCX Ink retarget target has an invalid relationship ID",
            ));
        }
        selected = Some(SelectedAttribute {
            value_range: start..end,
        });
    }
    selected.ok_or_else(|| invalid("DOCX Ink retarget target is missing expanded r:id"))
}

fn source_offset(value: &[u8], source: &[u8]) -> Result<usize> {
    let start = (value.as_ptr() as usize)
        .checked_sub(source.as_ptr() as usize)
        .ok_or_else(|| invalid("DOCX Ink retarget attribute is not source-backed"))?;
    let end = start
        .checked_add(value.len())
        .filter(|end| *end <= source.len())
        .ok_or_else(|| invalid("DOCX Ink retarget attribute range is invalid"))?;
    if source.get(start..end) != Some(value) {
        return Err(invalid("DOCX Ink retarget attribute source bytes diverged"));
    }
    Ok(start)
}

fn add_retarget_edit(
    edits: &mut Vec<(Range<usize>, Vec<u8>)>,
    source_range: Range<usize>,
    relationship_id: &str,
    replacement_bytes: &mut usize,
) -> Result<()> {
    let value = escape_attribute_value(relationship_id)?;
    *replacement_bytes = (*replacement_bytes)
        .checked_add(value.len())
        .ok_or_else(|| {
            limit(
                "Ink retarget replacement bytes",
                usize::MAX,
                MAX_RETARGET_VALUE_BYTES,
            )
        })?;
    if *replacement_bytes > MAX_RETARGET_VALUE_BYTES {
        return Err(limit(
            "Ink retarget replacement bytes",
            *replacement_bytes,
            MAX_RETARGET_VALUE_BYTES,
        ));
    }
    edits.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "DOCX Ink retarget attribute edits",
        source,
    })?;
    edits.push((source_range, value));
    Ok(())
}

fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid("DOCX Ink relationship ID is empty or too long"));
    }
    let mut characters = value.chars();
    let Some(first) = characters.next() else {
        return Err(invalid("DOCX Ink relationship ID is empty"));
    };
    if !(first == '_' || first.is_ascii_alphabetic())
        || characters.any(|character| {
            !(character == '_'
                || character == '-'
                || character == '.'
                || character.is_ascii_alphanumeric())
        })
    {
        return Err(invalid("DOCX Ink relationship ID is not an XML ID token"));
    }
    Ok(())
}

fn escape_attribute_value(value: &str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink relationship ID escape",
            source,
        })?;
    for byte in value.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'>' => output.extend_from_slice(b"&gt;"),
            b'\'' => output.extend_from_slice(b"&apos;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            byte => output.push(byte),
        }
    }
    Ok(output)
}

fn observe_envelope_start<T>(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    stack: &mut [T],
    nodes: &mut usize,
    limits: Limits,
    root_seen: &mut bool,
    root_closed: bool,
    prolog_started: &mut bool,
) -> Result<()> {
    xml::element(element, resolver)?;
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid(
            "DOCX source element uses an unknown namespace prefix",
        ));
    }
    observe_node(nodes, limits.max_xml_nodes)?;
    let depth = stack
        .len()
        .checked_add(1)
        .ok_or_else(|| limit("XML depth", usize::MAX, limits.max_xml_depth))?;
    if depth > limits.max_xml_depth {
        return Err(limit("XML depth", depth, limits.max_xml_depth));
    }
    if *root_seen && root_closed {
        return Err(invalid("DOCX source has more than one XML root"));
    }
    if stack.is_empty() {
        if *root_seen {
            return Err(invalid("DOCX source has more than one XML root"));
        }
        *root_seen = true;
    }
    *prolog_started = true;
    Ok(())
}

fn observe_envelope_empty<T>(
    element: &BytesStart<'_>,
    namespace: &ResolveResult<'_>,
    resolver: &NamespaceResolver,
    stack: &mut [T],
    nodes: &mut usize,
    limits: Limits,
    root_seen: &mut bool,
    root_closed: &mut bool,
    prolog_started: &mut bool,
) -> Result<()> {
    xml::element(element, resolver)?;
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid(
            "DOCX source element uses an unknown namespace prefix",
        ));
    }
    observe_node(nodes, limits.max_xml_nodes)?;
    let depth = stack
        .len()
        .checked_add(1)
        .ok_or_else(|| limit("XML depth", usize::MAX, limits.max_xml_depth))?;
    if depth > limits.max_xml_depth {
        return Err(limit("XML depth", depth, limits.max_xml_depth));
    }
    if stack.is_empty() {
        if *root_seen {
            return Err(invalid("DOCX source has more than one XML root"));
        }
        *root_seen = true;
        *root_closed = true;
    } else if *root_closed {
        return Err(invalid("DOCX source has markup after its XML root"));
    }
    *prolog_started = true;
    Ok(())
}

fn observe_outside_text(
    stack: &[ScanFrame],
    root_seen: &mut bool,
    root_closed: bool,
    prolog_started: &mut bool,
    text: &[u8],
) -> Result<()> {
    if stack.is_empty() && (!*root_seen || root_closed) && !text.iter().all(u8::is_ascii_whitespace)
    {
        return Err(invalid(
            "DOCX source has non-whitespace text outside its root",
        ));
    }
    if !*root_seen {
        *prolog_started = true;
    }
    Ok(())
}

fn reader(source: &[u8]) -> NsReader<&[u8]> {
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);
    reader
}

fn observe_node(nodes: &mut usize, maximum: usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("XML nodes", usize::MAX, maximum))?;
    if *nodes > maximum {
        return Err(limit("XML nodes", *nodes, maximum));
    }
    Ok(())
}

fn opening_span(
    event_end: usize,
    element: &BytesStart<'_>,
    empty: bool,
    source: &[u8],
) -> Result<(usize, usize)> {
    let suffix = if empty { 3 } else { 2 };
    let consumed = element
        .as_ref()
        .len()
        .checked_add(suffix)
        .ok_or_else(|| invalid("DOCX source element position overflowed"))?;
    // `NsReader::buffer_position()` is sampled after `read_event()`, so this
    // is the byte immediately following the complete opening tag.
    let start = event_end
        .checked_sub(consumed)
        .ok_or_else(|| invalid("DOCX source element position underflowed"))?;
    if source.get(start).copied() != Some(b'<') {
        return Err(invalid(
            "DOCX source element position is not an opening tag",
        ));
    }
    Ok((start, event_end))
}

/// The byte offset in the reader's input of reader position `position`;
/// `origin` is that input's [`ReaderOrigin`].
fn position(origin: ReaderOrigin, position: u64) -> Result<usize> {
    origin
        .offset(position)
        .ok_or_else(|| Error::Invalid("DOCX source position does not fit usize".into()))
}

fn to_u32(offset: usize) -> Result<u32> {
    u32::try_from(offset)
        .map_err(|error| Error::Invalid(format!("DOCX source position exceeds u32: {error}")))
}

fn dialect_word_namespace(dialect: StoryDialect) -> &'static [u8] {
    match dialect {
        StoryDialect::Transitional => TRANSITIONAL_WORD,
        StoryDialect::Strict => STRICT_WORD,
    }
}

fn relationship_namespace(dialect: StoryDialect) -> &'static [u8] {
    match dialect {
        StoryDialect::Transitional => TRANSITIONAL_RELATIONSHIPS,
        StoryDialect::Strict => STRICT_RELATIONSHIPS,
    }
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::InkLimit {
        resource,
        actual,
        maximum,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};
    use soapberry_zip::office::StreamingArchiveWriter;

    const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    fn story(body: &str) -> String {
        format!(
            r#"<q:document xmlns:q="{WORD}" xmlns:r="{REL}" xmlns:z="{REL}" xmlns:mc="{MC}" xmlns:ext="urn:external" xmlns:wpi="{WORDPROCESSING_INK_STR}"><q:body>{body}</q:body></q:document>"#,
            WORDPROCESSING_INK_STR = std::str::from_utf8(WORDPROCESSING_INK).unwrap(),
        )
    }

    fn limits() -> Limits {
        Limits::default()
    }

    fn source_part(xml: &str) -> OwnedXmlPart {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/part.xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer.write_stored("part.xml", xml.as_bytes()).unwrap();
        let package = OpcPackage::from_bytes(&writer.finish_to_bytes().unwrap()).unwrap();
        let name = PackURI::new("/part.xml").unwrap();
        package.source_xml_part(&name).unwrap()
    }

    #[test]
    fn paragraphs_are_namespace_aware_and_keep_semantic_order() {
        let xml = story(
            r#"<q:p/><f:p xmlns:f="urn:foreign"/><q:tbl><q:tr><q:tc><q:p/></q:tc></q:tr></q:tbl><q:p/>"#,
        );
        let spans = paragraphs(xml.as_bytes(), StoryDialect::Transitional, limits()).unwrap();
        assert_eq!(
            spans
                .iter()
                .map(|span| &xml[span.clone()])
                .collect::<Vec<_>>(),
            vec!["<q:p/>", "<q:p/>", "<q:p/>"]
        );
    }

    #[test]
    fn source_offsets_account_for_prolog_whitespace_and_tag_shapes() {
        let xml = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n  {}",
            story(r#"<q:p/><q:p><q:r/></q:p>"#)
        );
        let spans = paragraphs(xml.as_bytes(), StoryDialect::Transitional, limits()).unwrap();
        assert_eq!(
            spans
                .iter()
                .map(|span| &xml[span.clone()])
                .collect::<Vec<_>>(),
            vec!["<q:p/>", "<q:p>"]
        );
    }

    #[test]
    fn external_mce_fallback_is_retained_but_ineligible() {
        let xml = story(
            r#"<mc:AlternateContent><mc:Choice Requires="ext"><q:p/></mc:Choice><mc:Fallback><q:p/></mc:Fallback></mc:AlternateContent><q:p/>"#,
        );
        let records =
            paragraph_records(xml.as_bytes(), StoryDialect::Transitional, limits()).unwrap();
        assert_eq!(records.len(), 2);
        assert!(!records[0].eligible);
        assert!(records[1].eligible);
        let spans = records
            .iter()
            .map(|record| record.span.clone())
            .collect::<Vec<_>>();
        assert_eq!(
            paragraph_eligibility(xml.as_bytes(), &spans, StoryDialect::Transitional, limits())
                .unwrap(),
            vec![false, true]
        );
    }

    #[test]
    fn paragraph_scanner_rejects_invalid_xml_and_bounds() {
        assert!(paragraphs(
            br#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><q:body><q:p></q:body></q:document>"#,
            StoryDialect::Transitional,
            limits(),
        )
        .is_err());
        let xml = story(r#"<q:p/><q:p/>"#);
        let bounded = Limits {
            max_xml_nodes: 2,
            ..limits()
        };
        assert!(paragraphs(xml.as_bytes(), StoryDialect::Transitional, bounded).is_err());
    }

    #[test]
    fn fallback_start_requires_one_complete_host_fallback() {
        let xml = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n  {}",
            story(
                r#"<q:p><q:r><m:AlternateContent xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006"><m:Choice Requires="wpi"><q:r/></m:Choice><m:Fallback><q:r/></m:Fallback></m:AlternateContent></q:r></q:p>"#,
            )
        );
        let start = xml.find("<m:AlternateContent").unwrap();
        let end = xml[start..].find("</m:AlternateContent>").unwrap()
            + start
            + "</m:AlternateContent>".len();
        let fallback = fallback_start(xml.as_bytes(), start..end, limits()).unwrap();
        assert_eq!(&xml[fallback], "<m:Fallback>");

        let duplicate = xml.replace("</m:Choice>", "</m:Choice><m:Fallback><q:r/></m:Fallback>");
        let start = duplicate.find("<m:AlternateContent").unwrap();
        let end = duplicate[start..].find("</m:AlternateContent>").unwrap()
            + start
            + "</m:AlternateContent>".len();
        assert!(fallback_start(duplicate.as_bytes(), start..end, limits()).is_err());
    }

    #[test]
    fn retarget_preserves_aliases_quotes_and_unrelated_attributes() {
        let xml = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n  {}",
            story(r#"<q:p><q:r><q:contentPart z:id='old' custom="keep"/></q:r></q:p>"#)
        );
        let start = xml.find("<q:contentPart").unwrap();
        let end = xml[start..].find("/>").unwrap() + start + 2;
        let source = source_part(&xml);
        let replacement = retarget(
            &source,
            start..end,
            StoryDialect::Transitional,
            "new",
            xml.len() + 32,
        )
        .unwrap();
        assert_eq!(
            std::str::from_utf8(replacement.bytes()).unwrap(),
            xml.replace("z:id='old'", "z:id='new'")
        );
        assert_eq!(
            retarget(
                &source,
                start..end,
                StoryDialect::Transitional,
                "old",
                xml.len()
            )
            .unwrap()
            .bytes(),
            source.bytes()
        );
    }

    #[test]
    fn retarget_rejects_foreign_namespaces_unordered_targets_and_bounds() {
        let xml = story(
            r#"<q:p><q:r><q:contentPart z:id="one"/><q:contentPart z:id="two"/></q:r></q:p>"#,
        );
        let first = xml.find("<q:contentPart").unwrap();
        let first_end = xml[first..].find("/>").unwrap() + first + 2;
        let second = xml[first_end..].find("<q:contentPart").unwrap() + first_end;
        let second_end = xml[second..].find("/>").unwrap() + second + 2;
        let source = source_part(&xml);
        let retargeted = retarget_many(
            &source,
            &[(first..first_end, "three"), (second..second_end, "four")],
            StoryDialect::Transitional,
            xml.len() + 32,
        )
        .unwrap();
        let bytes = std::str::from_utf8(retargeted.bytes()).unwrap();
        assert!(bytes.contains("z:id=\"three\"") && bytes.contains("z:id=\"four\""));
        assert!(
            retarget_many(
                &source,
                &[(second..second_end, "four"), (first..first_end, "three")],
                StoryDialect::Transitional,
                xml.len() + 32,
            )
            .is_err()
        );
        assert!(
            retarget(
                &source,
                first..first_end,
                StoryDialect::Transitional,
                "new",
                xml.len() - 1
            )
            .is_err()
        );

        let foreign_xml =
            story(r#"<q:p><q:r><x:contentPart xmlns:x="urn:foreign" z:id="bad"/></q:r></q:p>"#);
        let foreign_start = foreign_xml.find("<x:contentPart").unwrap();
        let foreign_end = foreign_xml[foreign_start..].find("/>").unwrap() + foreign_start + 2;
        let foreign = source_part(&foreign_xml);
        assert!(
            retarget(
                &foreign,
                foreign_start..foreign_end,
                StoryDialect::Transitional,
                "new",
                foreign_xml.len() + 32
            )
            .is_err()
        );
    }

    #[test]
    fn retarget_fallbacks_preserves_vml_lexical_content_and_noop_bytes() {
        let xml = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n  {}",
            story(
                r#"<q:p><q:r><m:AlternateContent xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006" data-keep="host"><m:Choice Requires="wpi"><q:r/></m:Choice><m:Fallback><q:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape style="width:10pt;height:10pt" data-keep="shape"><v:imagedata z:id='oldImage' xmlns:o="urn:schemas-microsoft-com:office:office" o:foo="keep" title="keep"/></v:shape></q:pict></m:Fallback></m:AlternateContent></q:r></q:p>"#,
            )
        );
        let start = xml.find("<m:AlternateContent").unwrap();
        let end = xml[start..].find("</m:AlternateContent>").unwrap()
            + start
            + "</m:AlternateContent>".len();
        let source = source_part(&xml);
        let retargeted = retarget_fallbacks(
            &source,
            &[(start..end, "newImage")],
            StoryDialect::Transitional,
            limits(),
        )
        .unwrap();
        let expected = xml.replace("z:id='oldImage'", "z:id='newImage'");
        assert_eq!(retargeted.bytes(), expected.as_bytes());
        for preserved in [
            "style=\"width:10pt;height:10pt\"",
            "data-keep=\"host\"",
            "data-keep=\"shape\"",
            "o:foo=\"keep\"",
            "title=\"keep\"",
        ] {
            assert!(
                std::str::from_utf8(retargeted.bytes())
                    .unwrap()
                    .contains(preserved),
                "retarget changed preserved attribute {preserved:?}"
            );
        }
        let no_op = retarget_fallbacks(
            &source,
            &[(start..end, "oldImage")],
            StoryDialect::Transitional,
            limits(),
        )
        .unwrap();
        assert_eq!(no_op.bytes(), source.bytes());
    }

    #[test]
    fn retarget_fallbacks_rejects_missing_or_duplicate_image_relationships() {
        for body in [
            r#"<q:p><q:r><m:AlternateContent xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006"><m:Choice Requires="wpi"><q:r/></m:Choice><m:Fallback><q:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:imagedata/></q:pict></m:Fallback></m:AlternateContent></q:r></q:p>"#,
            r#"<q:p><q:r><m:AlternateContent xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006"><m:Choice Requires="wpi"><q:r/></m:Choice><m:Fallback><q:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:imagedata z:id="one"/><v:imagedata z:id="two"/></q:pict></m:Fallback></m:AlternateContent></q:r></q:p>"#,
            r#"<q:p><q:r><m:AlternateContent xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006"><m:Choice Requires="wpi"><q:r/></m:Choice><m:Fallback><q:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:imagedata xmlns:o="urn:schemas-microsoft-com:office:office" o:relid="wordImage"/></q:pict></m:Fallback></m:AlternateContent></q:r></q:p>"#,
        ] {
            let xml = story(body);
            let start = xml.find("<m:AlternateContent").unwrap();
            let end = xml[start..].find("</m:AlternateContent>").unwrap()
                + start
                + "</m:AlternateContent>".len();
            let source = source_part(&xml);
            assert!(
                retarget_fallbacks(
                    &source,
                    &[(start..end, "newImage")],
                    StoryDialect::Transitional,
                    limits(),
                )
                .is_err()
            );
        }
    }

    #[test]
    fn retarget_fallbacks_batches_ordered_hosts() {
        let xml = story(
            r#"<q:p><q:r><m:AlternateContent xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006"><m:Choice Requires="wpi"><q:r/></m:Choice><m:Fallback><q:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape style="width:10pt"><v:imagedata z:id="firstImage"/></v:shape></q:pict></m:Fallback></m:AlternateContent><m:AlternateContent xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006"><m:Choice Requires="wpi"><q:r/></m:Choice><m:Fallback><q:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape style="width:20pt"><v:imagedata z:id="secondImage"/></v:shape></q:pict></m:Fallback></m:AlternateContent></q:r></q:p>"#,
        );
        let ranges = xml
            .match_indices("<m:AlternateContent")
            .map(|(start, _)| {
                let end = xml[start..].find("</m:AlternateContent>").unwrap()
                    + start
                    + "</m:AlternateContent>".len();
                start..end
            })
            .collect::<Vec<_>>();
        assert_eq!(ranges.len(), 2);
        let source = source_part(&xml);
        let relationship_ids = ["firstReplacement", "secondReplacement"];
        let targets = ranges
            .iter()
            .zip(relationship_ids)
            .map(|(range, relationship_id)| (range.clone(), relationship_id))
            .collect::<Vec<_>>();
        let retargeted =
            retarget_fallbacks(&source, &targets, StoryDialect::Transitional, limits()).unwrap();
        let output = std::str::from_utf8(retargeted.bytes()).unwrap();
        assert!(output.contains("z:id=\"firstReplacement\""));
        assert!(output.contains("z:id=\"secondReplacement\""));
        assert!(output.contains("style=\"width:10pt\""));
        assert!(output.contains("style=\"width:20pt\""));
    }
}
