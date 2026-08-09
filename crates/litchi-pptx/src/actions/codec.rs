//! Bounded `PresentationML` XML scanning for action settings.

use super::model::Trigger;
use super::{Limits, invalid, limit};
use crate::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, process_markup_compatibility};
use litchi_ooxml_common::relationships::attribute_value;
use litchi_ooxml_common::xml::{is_drawingml_name, unqualified_attribute_value};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

pub(super) const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_ACTION_SETTINGS: usize = 4_096;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;
pub(super) const MAX_ATTRIBUTE_BYTES: usize = 4_096;

const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATIONML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";

pub(super) struct Parsed {
    pub(super) trigger: Trigger,
    pub(super) action: Option<String>,
    pub(super) relationship_id: Option<String>,
    pub(super) tooltip: Option<String>,
    pub(super) target_frame: Option<String>,
}

impl Limits {
    pub(super) fn add_slide_xml(&mut self, bytes: usize) -> Result<()> {
        if bytes > MAX_SLIDE_XML_BYTES {
            return Err(limit("slide XML bytes", MAX_SLIDE_XML_BYTES));
        }
        self.total_slide_xml_bytes = self
            .total_slide_xml_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("total slide XML bytes", MAX_TOTAL_SLIDE_XML_BYTES))?;
        if self.total_slide_xml_bytes > MAX_TOTAL_SLIDE_XML_BYTES {
            return Err(limit("total slide XML bytes", MAX_TOTAL_SLIDE_XML_BYTES));
        }
        Ok(())
    }

    pub(super) fn add_action(&mut self) -> Result<()> {
        self.action_count = self
            .action_count
            .checked_add(1)
            .ok_or_else(|| limit("slide action-setting count", MAX_ACTION_SETTINGS))?;
        if self.action_count > MAX_ACTION_SETTINGS {
            return Err(limit("slide action-setting count", MAX_ACTION_SETTINGS));
        }
        Ok(())
    }
}

pub(super) fn scan(xml_bytes: &[u8], limits: &mut Limits) -> Result<Vec<Parsed>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes", MAX_SLIDE_XML_BYTES));
    }

    let capabilities = Capabilities::ooxml_baseline();
    let mce_limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: MAX_SLIDE_XML_BYTES,
        max_output_bytes: MAX_SLIDE_XML_BYTES,
        max_depth: MAX_XML_DEPTH,
        max_namespace_bindings: 4_096,
        max_directive_tokens: 4_096,
        max_choices_per_alternate: 1_024,
    };
    let xml = process_markup_compatibility(xml_bytes, &capabilities, &mce_limits)?.xml;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut actions = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth", MAX_XML_DEPTH));
                }
                if depth == 1 {
                    validate_slide_root(&namespace, element.name(), saw_root)?;
                    saw_root = true;
                }
                maybe_push(
                    &mut actions,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    limits,
                )?;
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth", MAX_XML_DEPTH));
                }
                if child_depth == 1 {
                    validate_slide_root(&namespace, element.name(), saw_root)?;
                    saw_root = true;
                    closed_root = true;
                }
                maybe_push(
                    &mut actions,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    limits,
                )?;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid slide XML nesting"));
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid(
                            "slide XML must close with a PresentationML sld element",
                        ));
                    }
                    closed_root = true;
                }
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("slide XML must not contain a DTD")),
            Event::Eof => {
                if !saw_root || !closed_root || depth != 0 {
                    return Err(invalid("unterminated or missing PresentationML slide root"));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(actions)
}

fn maybe_push(
    actions: &mut Vec<Parsed>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    limits: &mut Limits,
) -> Result<()> {
    let trigger = if is_drawingml_name(namespace, element.name(), b"hlinkClick") {
        Trigger::Click
    } else if is_drawingml_name(namespace, element.name(), b"hlinkHover") {
        Trigger::Hover
    } else {
        return Ok(());
    };

    limits.add_action()?;
    let relationship_id = bounded_optional(
        attribute_value(element, b"id", decoder, resolver)?,
        "relationship ID",
    )?
    .filter(|value| !value.is_empty());
    let action = bounded_optional(
        unqualified_attribute_value(element, b"action", decoder)?,
        "action string",
    )?;
    let tooltip = bounded_optional(
        unqualified_attribute_value(element, b"tooltip", decoder)?,
        "tooltip",
    )?;
    let target_frame = bounded_optional(
        unqualified_attribute_value(element, b"tgtFrame", decoder)?,
        "target frame",
    )?;
    actions.push(Parsed {
        trigger,
        action,
        relationship_id,
        tooltip,
        target_frame,
    });
    Ok(())
}

fn validate_slide_root(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    root_seen: bool,
) -> Result<()> {
    if root_seen || !is_presentationml_name(namespace, name, b"sld") {
        return Err(invalid(
            "slide XML must have one PresentationML sld root element",
        ));
    }
    Ok(())
}

fn is_presentationml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == PRESENTATIONML_NAMESPACE || *value == STRICT_PRESENTATIONML_NAMESPACE
        },
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"p",
        ResolveResult::Unbound => false,
    }
}

fn bounded_optional(value: Option<String>, what: &'static str) -> Result<Option<String>> {
    if let Some(value) = &value
        && value.len() > MAX_ATTRIBUTE_BYTES
    {
        return Err(limit(what, MAX_ATTRIBUTE_BYTES));
    }
    Ok(value)
}

fn increment_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("slide XML node count", MAX_XML_NODES))?;
    if *nodes > MAX_XML_NODES {
        return Err(limit("slide XML node count", MAX_XML_NODES));
    }
    Ok(())
}
