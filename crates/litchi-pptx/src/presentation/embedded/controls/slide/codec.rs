//! Lossless source ranges for the supported control and OCX attributes.
//!
//! Markup-compatibility processing is intentionally left to the existing
//! control loader.  For publication we locate every raw `p:control` branch
//! carrying the selected relationship ID and edit the corresponding
//! attributes in place.  This updates both a Choice and its Fallback when a
//! producer supplied both, while retaining the original branch envelope and
//! every unknown child byte.

use std::ops::Range;

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

pub(crate) fn locate_controls(xml: &[u8]) -> Result<Vec<SourceControl>> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("control slide XML bytes", MAX_SLIDE_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut c_sld_depth = None;
    let mut controls_depth = None;
    let mut open_control_depth = None;
    let mut controls = Vec::new();
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;

    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("control XML depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("control XML depth", MAX_XML_DEPTH));
                }
                if child_depth == 1 {
                    if root_seen || !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid("control XML must have one PresentationML sld root"));
                    }
                    root_seen = true;
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
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("control XML depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("control XML depth", MAX_XML_DEPTH));
                }
                if child_depth == 1 {
                    if root_seen || !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid("control XML must have one PresentationML sld root"));
                    }
                    root_seen = true;
                    root_closed = true;
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
                    root_closed = true;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "control XML rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || depth != 0 || open_control_depth.is_some() {
                    return Err(invalid("unterminated PresentationML control slide"));
                }
                return Ok(controls);
            },
            _ => {},
        }
    }
}

pub(crate) fn locate_descriptor(xml: &[u8]) -> Result<SourceDescriptor> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("control descriptor XML bytes", MAX_SLIDE_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut root_seen = false;
    let mut root_closed = false;
    let mut depth = 0usize;
    let mut result = None;
    let mut nodes = 0usize;
    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let empty = matches!(&event, Event::Empty(_));
        let event = event.into_owned();
        match event {
            Event::Start(element) | Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("control descriptor depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("control descriptor depth", MAX_XML_DEPTH));
                }
                if !root_seen {
                    if !matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == *AX)
                        || element.name().local_name().as_ref() != b"ocx"
                    {
                        return Err(invalid("control descriptor must have an ax:ocx root"));
                    }
                    root_seen = true;
                    result = Some(parse_descriptor_source(&element, decoder, &reader, start)?);
                }
                if empty && depth == 0 {
                    root_closed = true;
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
                    root_closed = true;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "control descriptor rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => {
                if !root_seen || !root_closed || depth != 0 {
                    return Err(invalid("unterminated ActiveX control descriptor"));
                }
                return result.ok_or_else(|| invalid("control descriptor is empty"));
            },
            _ => {},
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
    for attribute in element.attributes().with_checks(true) {
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
    replacements.sort_by(|left, right| right.range.start.cmp(&left.range.start));
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

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid("ActiveX XML source offset overflow"))
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
