use std::ops::Range;

use super::codec::{anchor_id, scan_layout, selected_raw_span_with_limits};
use crate::tag::pml;
use crate::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce::{Capabilities, OffsetLimits, active_offsets};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::Part as OpcPart;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

pub(super) const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
pub(super) const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";

pub(super) fn validate_owner_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml"
    ) {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: "PresentationML shape-tag owner".into(),
            actual: content_type.into(),
        })
    }
}

pub(super) fn validate_staged_anchor(
    owner: &dyn OpcPart,
    xml: &[u8],
    key: crate::shape::Key<'_>,
    expected_id: Option<&str>,
    shape_limits: crate::shape::Limits,
) -> Result<()> {
    validate_owner_content_type(owner.content_type())?;
    let staged = scan_layout(
        xml,
        selected_raw_span_with_limits(xml, key, shape_limits)?,
        shape_limits,
    )?;
    if staged.anchor.as_ref().map(|anchor| anchor.id.as_str()) != expected_id {
        return Err(crate::tag::invalid(
            "staged shape tag anchor did not round-trip",
        ));
    }
    Ok(())
}

pub(super) fn active_pml_offsets(
    xml: &[u8],
    span: &Range<usize>,
    shape_limits: crate::shape::Limits,
) -> Result<Vec<u32>> {
    if xml.len() > shape_limits.input_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
            observed: xml.len(),
            maximum: shape_limits.input_bytes(),
        });
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut offsets = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    loop {
        let start = xml_position(&reader, origin)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let profile = pml(&namespace);
        drop(namespace);
        match event {
            Event::Start(_) => {
                enter_depth(&mut depth, shape_limits)?;
                if start >= span.start && start < span.end && profile.is_some() {
                    bump_nodes(&mut nodes, shape_limits)?;
                    try_push(
                        &mut offsets,
                        offset_u32(start)?,
                        "shape MCE candidate offsets",
                    )?;
                }
            },
            Event::Empty(_) => {
                enter_depth(&mut depth, shape_limits)?;
                leave_depth(&mut depth)?;
                if start >= span.start && start < span.end && profile.is_some() {
                    bump_nodes(&mut nodes, shape_limits)?;
                    try_push(
                        &mut offsets,
                        offset_u32(start)?,
                        "shape MCE candidate offsets",
                    )?;
                }
            },
            Event::End(_) => leave_depth(&mut depth)?,
            Event::Eof => break,
            _ => {},
        }
        if start >= span.end {
            break;
        }
    }
    active_offsets(
        xml,
        &offsets,
        &shape_mce_capabilities(),
        &shape_offset_limits(shape_limits),
    )
    .map_err(Into::into)
}

pub(super) fn preserved_anchor_uses(
    xml: &[u8],
    relationship_id: &str,
    shape_limits: crate::shape::Limits,
) -> Result<usize> {
    if xml.len() > shape_limits.input_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
            observed: xml.len(),
            maximum: shape_limits.input_bytes(),
        });
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut uses = 0usize;
    loop {
        let start = xml_position(&reader, origin)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let profile = pml(&namespace);
        drop(namespace);
        let is_start = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if is_start {
                    enter_depth(&mut depth, shape_limits)?;
                } else {
                    enter_depth(&mut depth, shape_limits)?;
                    leave_depth(&mut depth)?;
                }
                bump_nodes(&mut nodes, shape_limits)?;
                if profile.is_none() || element.local_name().as_ref() != b"tags" {
                    continue;
                }
                let profile = profile
                    .ok_or_else(|| crate::tag::invalid("preserved p:tags profile is missing"))?;
                let (id, _) = anchor_id(
                    &reader,
                    xml,
                    &element,
                    start..xml_position(&reader, origin)?,
                    profile,
                )?;
                if id == relationship_id {
                    uses = uses.checked_add(1).ok_or(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: usize::MAX,
                        maximum: shape_limits.nodes(),
                    })?;
                }
            },
            Event::End(_) => leave_depth(&mut depth)?,
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(uses)
}

pub(super) fn shape_mce_capabilities() -> Capabilities {
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(P14);
    capabilities.understand_namespace(P15);
    capabilities
}

pub(super) fn shape_mce_limits(limits: crate::shape::Limits) -> litchi_ooxml_common::mce::Limits {
    litchi_ooxml_common::mce::Limits {
        max_input_bytes: limits.input_bytes(),
        max_output_bytes: limits.output_bytes(),
        max_depth: limits.depth(),
        ..litchi_ooxml_common::mce::Limits::default()
    }
}

pub(super) fn shape_offset_limits(limits: crate::shape::Limits) -> OffsetLimits {
    OffsetLimits {
        max_source_bytes: limits.input_bytes(),
        max_offsets: limits.nodes(),
        max_marked_bytes: limits.input_bytes(),
        processing: shape_mce_limits(limits),
    }
}

pub(super) fn has_non_namespace_attrs(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if name != b"xmlns" && !name.starts_with(b"xmlns:") {
            return Ok(true);
        }
    }
    Ok(false)
}

pub(super) fn take_active<I>(active: &mut std::iter::Peekable<I>, start: usize) -> Result<bool>
where
    I: Iterator<Item = u32>,
{
    let start = offset_u32(start)?;
    if active.peek().copied() == Some(start) {
        let _ = active.next();
        Ok(true)
    } else if active.peek().is_some_and(|offset| *offset < start) {
        Err(crate::tag::invalid(
            "MCE-active shape offsets are out of source order",
        ))
    } else {
        Ok(false)
    }
}

pub(super) fn try_push<T>(values: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    if values.len() == values.capacity() {
        values
            .try_reserve(1)
            .map_err(|source| crate::tag::allocation(resource, source))?;
    }
    values.push(value);
    Ok(())
}

pub(super) fn bump_nodes(nodes: &mut usize, limits: crate::shape::Limits) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or(Error::Limit {
        kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
        observed: usize::MAX,
        maximum: limits.nodes(),
    })?;
    if *nodes > limits.nodes() {
        Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
            observed: *nodes,
            maximum: limits.nodes(),
        })
    } else {
        Ok(())
    }
}

pub(super) fn enter_depth(depth: &mut usize, limits: crate::shape::Limits) -> Result<()> {
    *depth = depth.checked_add(1).ok_or(Error::Limit {
        kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
        observed: usize::MAX,
        maximum: limits.depth(),
    })?;
    if *depth > limits.depth() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
            observed: *depth,
            maximum: limits.depth(),
        });
    }
    Ok(())
}

pub(super) fn leave_depth(depth: &mut usize) -> Result<()> {
    *depth = depth
        .checked_sub(1)
        .ok_or_else(|| crate::tag::invalid("shape-tag owner XML depth underflow"))?;
    Ok(())
}

pub(super) fn offset_u32(offset: usize) -> Result<u32> {
    u32::try_from(offset)
        .map_err(|_err| crate::tag::invalid("shape-tag XML offset does not fit u32"))
}

/// The byte offset in the reader's input of its current position; `origin`
/// is that input's [`ReaderOrigin`].
pub(super) fn xml_position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| crate::tag::invalid("shape-tag XML offset does not fit usize"))
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
