use std::ops::Range;

use super::codec::{MAX_OWNER_BYTES, MAX_OWNER_NODES, anchor_id, scan_layout, selected_raw_span};
use crate::tag::pml;
use crate::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, OffsetLimits, active_offsets};
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

pub(super) fn validate_staged_anchor<'k>(
    owner: &dyn OpcPart,
    xml: &[u8],
    key: crate::shape::Key<'k>,
    expected_id: Option<&str>,
) -> Result<()> {
    validate_owner_content_type(owner.content_type())?;
    let staged = scan_layout(xml, selected_raw_span(xml, key)?)?;
    if staged.anchor.as_ref().map(|anchor| anchor.id.as_str()) != expected_id {
        return Err(crate::tag::invalid(
            "staged shape tag anchor did not round-trip",
        ));
    }
    Ok(())
}

pub(super) fn active_pml_offsets(xml: &[u8], span: &Range<usize>) -> Result<Vec<u32>> {
    let mut reader = NsReader::from_reader(xml);
    let mut offsets = Vec::new();
    let mut nodes = 0usize;
    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let profile = pml(&namespace);
        drop(namespace);
        match event {
            Event::Start(_) | Event::Empty(_)
                if start >= span.start && start < span.end && profile.is_some() =>
            {
                bump_nodes(&mut nodes)?;
                try_push(
                    &mut offsets,
                    offset_u32(start)?,
                    "shape MCE candidate offsets",
                )?;
            },
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
        &OffsetLimits::default(),
    )
    .map_err(Into::into)
}

pub(super) fn preserved_anchor_uses(xml: &[u8], relationship_id: &str) -> Result<usize> {
    if xml.len() > MAX_OWNER_BYTES {
        return Err(Error::Limit {
            resource: "shape-tag owner XML bytes",
            limit: MAX_OWNER_BYTES,
        });
    }
    let mut reader = NsReader::from_reader(xml);
    let mut nodes = 0usize;
    let mut uses = 0usize;
    loop {
        let start = xml_position(&reader)?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let profile = pml(&namespace);
        drop(namespace);
        match event {
            Event::Start(element) | Event::Empty(element)
                if profile.is_some() && element.local_name().as_ref() == b"tags" =>
            {
                bump_nodes(&mut nodes)?;
                let profile = profile
                    .ok_or_else(|| crate::tag::invalid("preserved p:tags profile is missing"))?;
                let (id, _) = anchor_id(
                    &reader,
                    xml,
                    &element,
                    start..xml_position(&reader)?,
                    profile,
                )?;
                if id == relationship_id {
                    uses = uses.checked_add(1).ok_or(Error::Limit {
                        resource: "preserved shape tag-anchor references",
                        limit: MAX_OWNER_NODES,
                    })?;
                }
            },
            Event::Start(_) | Event::Empty(_) => bump_nodes(&mut nodes)?,
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

pub(super) fn has_non_namespace_attrs(element: &BytesStart<'_>) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
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

pub(super) fn bump_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "shape-tag owner XML nodes",
        limit: MAX_OWNER_NODES,
    })?;
    if *nodes > MAX_OWNER_NODES {
        Err(Error::Limit {
            resource: "shape-tag owner XML nodes",
            limit: MAX_OWNER_NODES,
        })
    } else {
        Ok(())
    }
}

pub(super) fn offset_u32(offset: usize) -> Result<u32> {
    u32::try_from(offset).map_err(|_| crate::tag::invalid("shape-tag XML offset does not fit u32"))
}

pub(super) fn xml_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| crate::tag::invalid("shape-tag XML offset does not fit usize"))
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
