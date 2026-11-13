//! Lossless source ranges for the supported control and OCX attributes.
//!
//! Markup-compatibility processing is intentionally left to the existing
//! control loader.  For publication we locate every raw `p:control` branch
//! carrying the selected relationship ID and edit the corresponding
//! attributes in place.  This updates both a Choice and its Fallback when a
//! producer supplied both, while retaining the original branch envelope and
//! every unknown child byte.

use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::presentation::embedded::{
    MAX_XML_DEPTH, REL, increment_nodes, invalid, is_presentationml_name, limit,
};

use super::super::model::Persistence;
use super::super::{ACTIVEX_NAMESPACE, MAX_SLIDE_XML_BYTES};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const AX: &[u8] = ACTIVEX_NAMESPACE;
const RELATIONSHIPS: &[u8] = REL;
const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

#[derive(Debug, Clone)]
pub(crate) struct SourceAttribute {
    pub(crate) value: String,
    pub(crate) value_span: Range<usize>,
    pub(crate) full_span: Range<usize>,
}

#[derive(Debug, Clone)]
pub(crate) struct SourceControl {
    pub(crate) relationship_id: Option<String>,
    pub(crate) name: Option<SourceAttribute>,
    pub(crate) show_as_icon: Option<SourceAttribute>,
    pub(crate) image_width: Option<SourceAttribute>,
    pub(crate) image_height: Option<SourceAttribute>,
    pub(crate) opening_insert: usize,
}

#[derive(Debug, Clone)]
pub(crate) struct SourceDescriptor {
    pub(crate) class_id: Option<SourceAttribute>,
    pub(crate) license: Option<SourceAttribute>,
    pub(crate) persistence: Option<SourceAttribute>,
    pub(crate) relationship_id: Option<SourceAttribute>,
    pub(crate) ax_prefix: String,
    pub(crate) opening_insert: usize,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct ControlChanges {
    pub(crate) name: Option<Option<String>>,
    pub(crate) show_as_icon: Option<Option<bool>>,
    pub(crate) image_width: Option<Option<u32>>,
    pub(crate) image_height: Option<Option<u32>>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct DescriptorChanges {
    pub(crate) class_id: Option<String>,
    pub(crate) license: Option<Option<String>>,
    pub(crate) persistence: Option<Persistence>,
    pub(crate) remove_binary_relationship: bool,
}

#[derive(Debug, Clone)]
struct Replacement {
    range: Range<usize>,
    value: Vec<u8>,
}

/// Where [`locate_controls`] is relative to the `p:sld` root.
///
/// ```text
/// Prolog --p:sld start tag-----> Open
/// Prolog --empty p:sld tag-----> Closed
/// Open   --p:sld end tag-------> Closed
/// any    --declaration, comment, text, CDATA--> unchanged
/// ```
///
/// The last line holds in every phase: a declaration is accepted anywhere and
/// content outside the root is ignored. Refused: any other top-level element,
/// before or after the root; a processing instruction or DOCTYPE anywhere; and
/// the end of input before `Closed`.
#[derive(Clone, Copy, PartialEq, Eq)]
enum ControlsRootPhase {
    /// No root start tag has been read.
    Prolog,
    /// Inside the root.
    Open,
    /// The root has ended.
    Closed,
}

pub(crate) fn locate_controls(xml: &[u8]) -> Result<Vec<SourceControl>> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml.len(),
            MAX_SLIDE_XML_BYTES,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut c_sld_depth = None;
    let mut controls_depth = None;
    let mut open_control_depth = None;
    let mut controls = Vec::new();
    let mut nodes = 0usize;
    let mut phase = ControlsRootPhase::Prolog;

    loop {
        let start = position(&reader, origin)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        child_depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if child_depth == 1 {
                    if phase != ControlsRootPhase::Prolog
                        || !is_presentationml_name(&namespace, element.name(), b"sld")
                    {
                        return Err(invalid("control XML must have one PresentationML sld root"));
                    }
                    phase = ControlsRootPhase::Open;
                } else if child_depth == 2
                    && is_presentationml_name(&namespace, element.name(), b"cSld")
                {
                    c_sld_depth = Some(child_depth);
                } else if c_sld_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    if controls_depth.is_some() {
                        return Err(invalid("slide contains multiple control containers"));
                    }
                    controls_depth = Some(child_depth);
                } else if controls_depth.is_some()
                    && open_control_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    controls.push(parse_control_source(&element, decoder, &reader, start)?);
                    open_control_depth = Some(child_depth);
                }
                depth = child_depth;
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        child_depth,
                        MAX_XML_DEPTH,
                    ));
                }
                if child_depth == 1 {
                    if phase != ControlsRootPhase::Prolog
                        || !is_presentationml_name(&namespace, element.name(), b"sld")
                    {
                        return Err(invalid("control XML must have one PresentationML sld root"));
                    }
                    phase = ControlsRootPhase::Closed;
                } else if child_depth == 2
                    && is_presentationml_name(&namespace, element.name(), b"cSld")
                {
                    c_sld_depth = Some(child_depth);
                } else if c_sld_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    if controls_depth.is_some() {
                        return Err(invalid("slide contains multiple control containers"));
                    }
                } else if controls_depth.is_some()
                    && open_control_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    controls.push(parse_control_source(&element, decoder, &reader, start)?);
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid control XML nesting"));
                }
                if open_control_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    open_control_depth = None;
                }
                if controls_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    controls_depth = None;
                }
                if c_sld_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"cSld")
                {
                    c_sld_depth = None;
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid("control XML must close with p:sld"));
                    }
                    phase = ControlsRootPhase::Closed;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "control XML rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => match phase {
                ControlsRootPhase::Closed if depth == 0 && open_control_depth.is_none() => {
                    return Ok(controls);
                },
                ControlsRootPhase::Prolog | ControlsRootPhase::Open | ControlsRootPhase::Closed => {
                    return Err(invalid("unterminated PresentationML control slide"));
                },
            },
            // Accepted in every phase and ignored.
            Event::Decl(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

/// Where [`locate_descriptor`] is relative to the `ax:ocx` root.
///
/// ```text
/// Prolog --ax:ocx start tag--------> Open
/// Prolog --empty ax:ocx tag--------> Closed
/// Open   --the root's end tag------> Closed
/// Closed --top-level element-------> Closed (accepted: skipped unread)
/// any    --declaration, comment, text, CDATA--> unchanged
/// ```
///
/// Only the first top-level element is the descriptor. A later one is not
/// checked against `ax:ocx` and nothing in it is read, but it must end before
/// the end of input. The last line holds in every phase: a declaration is
/// accepted anywhere and content outside the root is ignored. Refused: a first
/// top-level element other than `ax:ocx`; a processing instruction or DOCTYPE
/// anywhere; and the end of input before `Closed` or inside a later element.
#[derive(Clone, Copy, PartialEq, Eq)]
enum DescriptorRootPhase {
    /// No root start tag has been read.
    Prolog,
    /// Inside the root.
    Open,
    /// The root has ended.
    Closed,
}

pub(crate) fn locate_descriptor(xml: &[u8]) -> Result<SourceDescriptor> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml.len(),
            MAX_SLIDE_XML_BYTES,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut phase = DescriptorRootPhase::Prolog;
    let mut depth = 0usize;
    let mut result = None;
    let mut nodes = 0usize;
    loop {
        let start = position(&reader, origin)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let empty = matches!(&event, Event::Empty(_));
        let event = event.into_owned();
        match event {
            Event::Start(element) | Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        MAX_XML_DEPTH,
                    )
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        child_depth,
                        MAX_XML_DEPTH,
                    ));
                }
                match phase {
                    DescriptorRootPhase::Prolog => {
                        if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == *AX)
                            || element.name().local_name().as_ref() != b"ocx"
                        {
                            return Err(invalid("control descriptor must have an ax:ocx root"));
                        }
                        phase = DescriptorRootPhase::Open;
                        result = Some(parse_descriptor_source(&element, decoder, &reader, start)?);
                    },
                    // A child of the root: not read.
                    DescriptorRootPhase::Open => {},
                    // A later top-level element, or one nested in it: accepted
                    // and skipped unread.
                    DescriptorRootPhase::Closed => {},
                }
                if empty && depth == 0 {
                    phase = DescriptorRootPhase::Closed;
                }
                depth = child_depth;
                if empty {
                    depth -= 1;
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("invalid control descriptor nesting"));
                }
                depth -= 1;
                if depth == 0 {
                    phase = DescriptorRootPhase::Closed;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "control descriptor rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => match phase {
                DescriptorRootPhase::Closed if depth == 0 => {
                    return result.ok_or_else(|| invalid("control descriptor is empty"));
                },
                DescriptorRootPhase::Prolog
                | DescriptorRootPhase::Open
                | DescriptorRootPhase::Closed => {
                    return Err(invalid("unterminated ActiveX control descriptor"));
                },
            },
            // Accepted in every phase and ignored.
            Event::Decl(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

pub(crate) fn rewrite_control(
    source: &[u8],
    relationship_id: &str,
    changes: &ControlChanges,
) -> Result<Vec<u8>> {
    let controls = locate_controls(source)?;
    let selected: Vec<&SourceControl> = controls
        .iter()
        .filter(|control| control.relationship_id.as_deref() == Some(relationship_id))
        .collect();
    if selected.is_empty() {
        return Err(invalid(format!(
            "control relationship '{relationship_id}' has no source XML anchor"
        )));
    }
    let mut replacements = Vec::new();
    for control in selected {
        let mut insert = Vec::new();
        apply_optional_string(
            &mut replacements,
            &mut insert,
            control.name.as_ref(),
            changes.name.as_ref(),
            b"name",
            control.opening_insert,
        )?;
        apply_optional_bool(
            &mut replacements,
            &mut insert,
            control.show_as_icon.as_ref(),
            changes.show_as_icon.as_ref(),
            b"showAsIcon",
            control.opening_insert,
        )?;
        apply_optional_u32(
            &mut replacements,
            &mut insert,
            control.image_width.as_ref(),
            changes.image_width.as_ref(),
            b"imgW",
            control.opening_insert,
        )?;
        apply_optional_u32(
            &mut replacements,
            &mut insert,
            control.image_height.as_ref(),
            changes.image_height.as_ref(),
            b"imgH",
            control.opening_insert,
        )?;
        if !insert.is_empty() {
            replacements.push(Replacement {
                range: control.opening_insert..control.opening_insert,
                value: insert,
            });
        }
    }
    apply_replacements(source, replacements)
}

pub(crate) fn rewrite_descriptor(source: &[u8], changes: &DescriptorChanges) -> Result<Vec<u8>> {
    let descriptor = locate_descriptor(source)?;
    let mut replacements = Vec::new();
    let mut insert = Vec::new();
    if let Some(value) = changes.class_id.as_deref() {
        let attribute = descriptor
            .class_id
            .as_ref()
            .ok_or_else(|| invalid("ActiveX descriptor classid source attribute is missing"))?;
        replacements.push(Replacement {
            range: attribute.value_span.clone(),
            value: escape(value).into_bytes(),
        });
    }
    apply_descriptor_optional_string(
        &mut replacements,
        &mut insert,
        descriptor.license.as_ref(),
        changes.license.as_ref(),
        &descriptor.ax_prefix,
        "license",
    )?;
    if let Some(persistence) = changes.persistence {
        let value = persistence_token(persistence);
        match (descriptor.persistence.as_ref(), value) {
            (Some(attribute), Some(value)) => replacements.push(Replacement {
                range: attribute.value_span.clone(),
                value: value.as_bytes().to_vec(),
            }),
            (Some(attribute), None) => replacements.push(Replacement {
                range: attribute.full_span.clone(),
                value: Vec::new(),
            }),
            (None, Some(value)) => insert.extend_from_slice(
                format!(" {}:persistence=\"{}\"", descriptor.ax_prefix, value).as_bytes(),
            ),
            (None, None) => {},
        }
    }
    if changes.remove_binary_relationship {
        let attribute = descriptor.relationship_id.as_ref().ok_or_else(|| {
            invalid("ActiveX descriptor binary relationship source attribute is missing")
        })?;
        replacements.push(Replacement {
            range: attribute.full_span.clone(),
            value: Vec::new(),
        });
    }
    if !insert.is_empty() {
        replacements.push(Replacement {
            range: descriptor.opening_insert..descriptor.opening_insert,
            value: insert,
        });
    }
    apply_replacements(source, replacements)
}

fn parse_control_source(
    element: &BytesStart<'_>,
    decoder: Decoder,
    reader: &NsReader<&[u8]>,
    start: usize,
) -> Result<SourceControl> {
    Ok(SourceControl {
        relationship_id: find_attribute(
            element,
            decoder,
            reader,
            start,
            b"id",
            AttributeKind::Relationship,
        )?
        .map(|attribute| attribute.value),
        name: find_attribute(
            element,
            decoder,
            reader,
            start,
            b"name",
            AttributeKind::Unqualified,
        )?,
        show_as_icon: find_attribute(
            element,
            decoder,
            reader,
            start,
            b"showAsIcon",
            AttributeKind::Unqualified,
        )?,
        image_width: find_attribute(
            element,
            decoder,
            reader,
            start,
            b"imgW",
            AttributeKind::Unqualified,
        )?,
        image_height: find_attribute(
            element,
            decoder,
            reader,
            start,
            b"imgH",
            AttributeKind::Unqualified,
        )?,
        opening_insert: opening_insert(element, start)?,
    })
}

fn parse_descriptor_source(
    element: &BytesStart<'_>,
    decoder: Decoder,
    reader: &NsReader<&[u8]>,
    start: usize,
) -> Result<SourceDescriptor> {
    let root_qname = element.name();
    let root_name = root_qname.as_ref();
    let ax_prefix = root_name
        .iter()
        .position(|byte| *byte == b':')
        .map(|offset| String::from_utf8_lossy(&root_name[..offset]).into_owned())
        .ok_or_else(|| invalid("ActiveX descriptor root must use an ax prefix"))?;
    Ok(SourceDescriptor {
        class_id: find_attribute(
            element,
            decoder,
            reader,
            start,
            b"classid",
            AttributeKind::Ax,
        )?,
        license: find_attribute(
            element,
            decoder,
            reader,
            start,
            b"license",
            AttributeKind::Ax,
        )?,
        persistence: find_attribute(
            element,
            decoder,
            reader,
            start,
            b"persistence",
            AttributeKind::Ax,
        )?,
        relationship_id: find_attribute(
            element,
            decoder,
            reader,
            start,
            b"id",
            AttributeKind::Relationship,
        )?,
        ax_prefix,
        opening_insert: opening_insert(element, start)?,
    })
}

#[derive(Clone, Copy)]
enum AttributeKind {
    Unqualified,
    Ax,
    Relationship,
}

fn find_attribute(
    element: &BytesStart<'_>,
    decoder: Decoder,
    reader: &NsReader<&[u8]>,
    start: usize,
    local: &[u8],
    kind: AttributeKind,
) -> Result<Option<SourceAttribute>> {
    let mut found = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != local {
            continue;
        }
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        let matches = match kind {
            AttributeKind::Unqualified => matches!(namespace, ResolveResult::Unbound),
            AttributeKind::Ax => {
                matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == *AX)
            },
            AttributeKind::Relationship => matches!(
                namespace,
                ResolveResult::Bound(Namespace(value))
                    if *value == *RELATIONSHIPS || *value == *STRICT_RELATIONSHIPS
            ),
        };
        if !matches {
            continue;
        }
        if found.is_some() {
            return Err(invalid(format!(
                "duplicate ActiveX attribute '{}'",
                String::from_utf8_lossy(local)
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let span = attribute_span(element.as_ref(), attribute.key.as_ref())?;
        let base = start
            .checked_add(1)
            .and_then(|value| value.checked_add(span.value_start))
            .ok_or_else(|| invalid("ActiveX attribute offset overflow"))?;
        let full_base = start
            .checked_add(1)
            .and_then(|value| value.checked_add(span.name_start))
            .ok_or_else(|| invalid("ActiveX attribute offset overflow"))?;
        found = Some(SourceAttribute {
            value,
            value_span: base..start
                .checked_add(1)
                .and_then(|value| value.checked_add(span.value_end))
                .ok_or_else(|| invalid("ActiveX attribute offset overflow"))?,
            full_span: full_base
                ..start
                    .checked_add(1)
                    .and_then(|value| value.checked_add(span.end))
                    .ok_or_else(|| invalid("ActiveX attribute offset overflow"))?,
        });
    }
    Ok(found)
}

fn apply_optional_string(
    replacements: &mut Vec<Replacement>,
    insert: &mut Vec<u8>,
    attribute: Option<&SourceAttribute>,
    desired: Option<&Option<String>>,
    name: &[u8],
    opening_insert: usize,
) -> Result<()> {
    let Some(desired) = desired else {
        return Ok(());
    };
    let desired = desired.as_deref();
    match (attribute, desired) {
        (Some(attribute), Some(value)) => replacements.push(Replacement {
            range: attribute.value_span.clone(),
            value: escape(value).into_bytes(),
        }),
        (Some(attribute), None) => replacements.push(Replacement {
            range: attribute.full_span.clone(),
            value: Vec::new(),
        }),
        (None, Some(value)) => {
            insert.extend_from_slice(b" ");
            insert.extend_from_slice(name);
            insert.extend_from_slice(b"=\"");
            insert.extend_from_slice(escape(value).as_bytes());
            insert.extend_from_slice(b"\"");
        },
        (None, None) => {},
    }
    let _ = opening_insert;
    Ok(())
}

fn apply_optional_bool(
    replacements: &mut Vec<Replacement>,
    insert: &mut Vec<u8>,
    attribute: Option<&SourceAttribute>,
    desired: Option<&Option<bool>>,
    name: &[u8],
    opening_insert: usize,
) -> Result<()> {
    let Some(desired) = desired else {
        return Ok(());
    };
    let desired = desired.map(|value| if value { "true" } else { "false" });
    apply_optional_string(
        replacements,
        insert,
        attribute,
        Some(&desired.map(str::to_owned)),
        name,
        opening_insert,
    )
}

fn apply_optional_u32(
    replacements: &mut Vec<Replacement>,
    insert: &mut Vec<u8>,
    attribute: Option<&SourceAttribute>,
    desired: Option<&Option<u32>>,
    name: &[u8],
    opening_insert: usize,
) -> Result<()> {
    let Some(desired) = desired else {
        return Ok(());
    };
    let desired = desired.map(|value| value.to_string());
    apply_optional_string(
        replacements,
        insert,
        attribute,
        Some(&desired),
        name,
        opening_insert,
    )
}

fn apply_descriptor_optional_string(
    replacements: &mut Vec<Replacement>,
    insert: &mut Vec<u8>,
    attribute: Option<&SourceAttribute>,
    desired: Option<&Option<String>>,
    prefix: &str,
    name: &str,
) -> Result<()> {
    let Some(desired) = desired else {
        return Ok(());
    };
    match (attribute, desired.as_deref()) {
        (Some(attribute), Some(value)) => replacements.push(Replacement {
            range: attribute.value_span.clone(),
            value: escape(value).into_bytes(),
        }),
        (Some(attribute), None) => replacements.push(Replacement {
            range: attribute.full_span.clone(),
            value: Vec::new(),
        }),
        (None, Some(value)) => {
            insert.extend_from_slice(format!(" {prefix}:{name}=\"{}\"", escape(value)).as_bytes());
        },
        (None, None) => {},
    }
    Ok(())
}

fn persistence_token(value: Persistence) -> Option<&'static str> {
    match value {
        Persistence::PropertyBag => Some("persistPropertyBag"),
        Persistence::Stream => Some("persistStream"),
        Persistence::StreamInit => Some("persistStreamInit"),
        Persistence::Storage => Some("persistStorage"),
        Persistence::Unknown => None,
    }
}

fn apply_replacements(source: &[u8], mut replacements: Vec<Replacement>) -> Result<Vec<u8>> {
    replacements.sort_by_key(|replacement| std::cmp::Reverse(replacement.range.start));
    let mut result = source.to_vec();
    let mut upper = source.len();
    for replacement in replacements {
        if replacement.range.start > replacement.range.end
            || replacement.range.end > source.len()
            || replacement.range.end > upper
        {
            return Err(invalid(
                "ActiveX source patch ranges overlap or escape the source",
            ));
        }
        result.splice(replacement.range.clone(), replacement.value);
        upper = replacement.range.start;
    }
    Ok(result)
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("ActiveX XML source offset overflow"))
}

fn opening_insert(element: &BytesStart<'_>, start: usize) -> Result<usize> {
    let raw = element.as_ref();
    let Some(close) = raw.iter().rposition(|byte| *byte == b'>') else {
        return start
            .checked_add(1)
            .and_then(|value| value.checked_add(raw.len()))
            .ok_or_else(|| invalid("ActiveX XML opening offset overflow"));
    };
    let insert = if close > 0 && raw[close - 1] == b'/' {
        close - 1
    } else {
        close
    };
    start
        .checked_add(1)
        .and_then(|value| value.checked_add(insert))
        .ok_or_else(|| invalid("ActiveX XML opening offset overflow"))
}

struct LocalAttributeSpan {
    name_start: usize,
    value_start: usize,
    value_end: usize,
    end: usize,
}

fn attribute_span(raw: &[u8], key: &[u8]) -> Result<LocalAttributeSpan> {
    let mut index = 0usize;
    while index < raw.len()
        && !raw[index].is_ascii_whitespace()
        && !matches!(raw[index], b'>' | b'/')
    {
        index += 1;
    }
    while index < raw.len() {
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] == b'>' || raw[index] == b'/' {
            break;
        }
        let name_start = index;
        while index < raw.len()
            && !raw[index].is_ascii_whitespace()
            && !matches!(raw[index], b'=' | b'>' | b'/')
        {
            index += 1;
        }
        let name = &raw[name_start..index];
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] != b'=' {
            return Err(invalid("ActiveX attribute has no value"));
        }
        index += 1;
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *raw
            .get(index)
            .ok_or_else(|| invalid("ActiveX attribute value is unterminated"))?;
        if quote != b'"' && quote != b'\'' {
            return Err(invalid("ActiveX attribute value is not quoted"));
        }
        index += 1;
        let value_start = index;
        while index < raw.len() && raw[index] != quote {
            index += 1;
        }
        if index >= raw.len() {
            return Err(invalid("ActiveX attribute value is unterminated"));
        }
        if name == key {
            return Ok(LocalAttributeSpan {
                name_start,
                value_start,
                value_end: index,
                end: index + 1,
            });
        }
        index += 1;
    }
    Err(invalid(format!(
        "ActiveX attribute '{}' has no source span",
        String::from_utf8_lossy(key)
    )))
}

fn escape(value: &str) -> String {
    let mut result = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => result.push_str("&amp;"),
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            '"' => result.push_str("&quot;"),
            '\'' => result.push_str("&apos;"),
            _ => result.push(character),
        }
    }
    result
}

#[cfg(test)]
mod envelope_tests {
    //! What each control source reader accepts before, around and after its root, and the exact error reported otherwise. Each expected
    //! line is the row label followed by the result.

    const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

    /// One reader's root element in the forms the rows are built from.
    struct Root<'a> {
        /// The start tag.
        open: &'a str,
        /// Content that makes the element valid for its reader.
        body: &'a str,
        /// The end tag.
        close: &'a str,
        /// The element as one empty-element tag.
        empty: &'a str,
    }

    fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
        let Root {
            open,
            body,
            close,
            empty,
        } = *root;
        let document = format!("{open}{body}{close}");
        let d = document.as_str();
        vec![
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECLARATION.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", format!("{open}{close}")),
            ("root then whitespace", format!("{d}\n ")),
            ("root then comment", format!("{d}<!--c-->")),
            ("root then PI", format!("{d}<?pi x?>")),
            ("root then text", format!("{d}text")),
            ("root then second root", format!("{d}{d}")),
            ("root then foreign start", format!("{d}<foreign></foreign>")),
            ("root then foreign empty", format!("{d}<foreign/>")),
            ("declaration after root", format!("{d}{DECLARATION}")),
            ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
            (
                "comment before declaration",
                format!("<!--c-->{DECLARATION}{d}"),
            ),
            ("text before root", format!("text{d}")),
            ("unterminated root", format!("{open}{body}")),
            ("stray end tag", format!("{d}</foreign>")),
            ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
            ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
            ("declaration then root", format!("{DECLARATION}{d}")),
            (
                "whitespace before declaration",
                format!(" {DECLARATION}{d}"),
            ),
            ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
            ("comment then root", format!("<!--c-->{d}")),
            ("PI then root", format!("<?pi x?>{d}")),
            ("root then text then second root", format!("{d}text{d}")),
            ("empty-element root then root", format!("{empty}{d}")),
            (
                "declaration after empty-element root",
                format!("{empty}{DECLARATION}"),
            ),
        ]
    }

    fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&[u8]) -> String) {
        let actual = rows(root)
            .into_iter()
            .map(|(label, xml)| format!("{label}: {}", read(xml.as_bytes())))
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
    }

    fn outcome<T, E: std::fmt::Debug>(
        result: Result<T, E>,
        summary: impl Fn(T) -> String,
    ) -> String {
        match result {
            Ok(value) => format!("Ok({})", summary(value)),
            Err(error) => format!("{error:?}"),
        }
    }

    /// The location of every control of one slide.
    #[test]
    fn control_location_envelope() {
        let root = Root {
            open: r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
            body: r"<p:cSld><p:spTree/></p:cSld>",
            close: r"</p:sld>",
            empty: r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"unterminated PresentationML control slide\")",
                "whitespace only: Invalid(\"unterminated PresentationML control slide\")",
                "declaration only: Invalid(\"unterminated PresentationML control slide\")",
                "empty-element root: Ok(0)",
                "start and end root: Ok(0)",
                "root then whitespace: Ok(0)",
                "root then comment: Ok(0)",
                "root then PI: Invalid(\"control XML rejects DTDs and processing instructions\")",
                "root then text: Ok(0)",
                "root then second root: Invalid(\"control XML must have one PresentationML sld root\")",
                "root then foreign start: Invalid(\"control XML must have one PresentationML sld root\")",
                "root then foreign empty: Invalid(\"control XML must have one PresentationML sld root\")",
                "declaration after root: Ok(0)",
                "two declarations: Ok(0)",
                "comment before declaration: Ok(0)",
                "text before root: Ok(0)",
                "unterminated root: Invalid(\"unterminated PresentationML control slide\")",
                "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "DOCTYPE: Invalid(\"control XML rejects DTDs and processing instructions\")",
                "top-level CDATA: Ok(0)",
                "declaration then root: Ok(0)",
                "whitespace before declaration: Ok(0)",
                "PI before declaration: Invalid(\"control XML rejects DTDs and processing instructions\")",
                "comment then root: Ok(0)",
                "PI then root: Invalid(\"control XML rejects DTDs and processing instructions\")",
                "root then text then second root: Invalid(\"control XML must have one PresentationML sld root\")",
                "empty-element root then root: Invalid(\"control XML must have one PresentationML sld root\")",
                "declaration after empty-element root: Ok(0)",
            ],
            |xml| {
                outcome(super::locate_controls(xml), |values| {
                    values.len().to_string()
                })
            },
        );
    }

    /// The ActiveX descriptor reader. Only the first top-level element is the
    /// descriptor; later ones are skipped.
    #[test]
    fn descriptor_location_envelope() {
        let root = Root {
            open: r#"<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ax:classid="first" ax:persistence="persistStream" r:id="rId1">"#,
            body: r"<ax:ocxPr/>",
            close: r"</ax:ocx>",
            empty: r#"<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ax:classid="first" ax:persistence="persistStream" r:id="rId1"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"unterminated ActiveX control descriptor\")",
                "whitespace only: Invalid(\"unterminated ActiveX control descriptor\")",
                "declaration only: Invalid(\"unterminated ActiveX control descriptor\")",
                "empty-element root: Ok(Some(\"first\"))",
                "start and end root: Ok(Some(\"first\"))",
                "root then whitespace: Ok(Some(\"first\"))",
                "root then comment: Ok(Some(\"first\"))",
                "root then PI: Invalid(\"control descriptor rejects DTDs and processing instructions\")",
                "root then text: Ok(Some(\"first\"))",
                "root then second root: Ok(Some(\"first\"))",
                "root then foreign start: Ok(Some(\"first\"))",
                "root then foreign empty: Ok(Some(\"first\"))",
                "declaration after root: Ok(Some(\"first\"))",
                "two declarations: Ok(Some(\"first\"))",
                "comment before declaration: Ok(Some(\"first\"))",
                "text before root: Ok(Some(\"first\"))",
                "unterminated root: Invalid(\"unterminated ActiveX control descriptor\")",
                "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "DOCTYPE: Invalid(\"control descriptor rejects DTDs and processing instructions\")",
                "top-level CDATA: Ok(Some(\"first\"))",
                "declaration then root: Ok(Some(\"first\"))",
                "whitespace before declaration: Ok(Some(\"first\"))",
                "PI before declaration: Invalid(\"control descriptor rejects DTDs and processing instructions\")",
                "comment then root: Ok(Some(\"first\"))",
                "PI then root: Invalid(\"control descriptor rejects DTDs and processing instructions\")",
                "root then text then second root: Ok(Some(\"first\"))",
                "empty-element root then root: Ok(Some(\"first\"))",
                "declaration after empty-element root: Ok(Some(\"first\"))",
            ],
            |xml| {
                outcome(super::locate_descriptor(xml), |value| {
                    format!("{:?}", value.class_id.map(|attribute| attribute.value))
                })
            },
        );
    }
}
