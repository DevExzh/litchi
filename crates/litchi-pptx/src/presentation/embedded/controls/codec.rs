use super::model::Persistence;
use super::{MAX_CONTROLS, MAX_SLIDE_XML_BYTES};
use crate::Result;
use crate::presentation::embedded::{
    MAX_XML_ATTRIBUTES, MAX_XML_DEPTH, bounded, increment_nodes, invalid, is_presentationml_name,
    limit, relationship_value, validate_root,
};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::NamespaceResolver;
use quick_xml::reader::NsReader;

#[derive(Default)]
pub(crate) struct Parsed {
    pub(crate) shape_id: Option<String>,
    pub(crate) name: Option<String>,
    pub(crate) show_as_icon: Option<bool>,
    pub(crate) image_width: Option<u32>,
    pub(crate) image_height: Option<u32>,
    pub(crate) relationship_id: Option<String>,
}

pub(crate) fn scan(xml_bytes: &[u8], count: &mut usize) -> Result<Vec<Parsed>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("control slide XML bytes", MAX_SLIDE_XML_BYTES));
    }
    let mce = MceLimits {
        max_input_bytes: MAX_SLIDE_XML_BYTES,
        max_output_bytes: MAX_SLIDE_XML_BYTES,
        max_depth: MAX_XML_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let xml =
        process_markup_compatibility(xml_bytes, &MceCapabilities::ooxml_baseline(), &mce)?.xml;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut values = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut c_sld_depth = None;
    let mut controls_depth = None;
    let mut open_control_depth = None;
    let mut saw_container = false;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("control XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("control XML depth", MAX_XML_DEPTH));
                }
                if depth == 1 {
                    validate_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                } else if depth == 2 && is_presentationml_name(&namespace, element.name(), b"cSld")
                {
                    c_sld_depth = Some(depth);
                } else if c_sld_depth == Some(depth - 1)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    if saw_container {
                        return Err(invalid("slide contains multiple control containers"));
                    }
                    saw_container = true;
                    controls_depth = Some(depth);
                } else if controls_depth == Some(depth - 1)
                    && open_control_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    add_control(count)?;
                    values.push(parse_control(&element, decoder, &resolver)?);
                    open_control_depth = Some(depth);
                }
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
                    validate_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                    root_closed = true;
                } else if c_sld_depth == Some(child_depth - 1)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    if saw_container {
                        return Err(invalid("slide contains multiple control containers"));
                    }
                    saw_container = true;
                } else if controls_depth == Some(child_depth - 1)
                    && open_control_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    add_control(count)?;
                    values.push(parse_control(&element, decoder, &resolver)?);
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid control XML nesting"));
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid("control XML must close with p:sld"));
                    }
                    root_closed = true;
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
                return Ok(values);
            },
            _ => {},
        }
    }
}

fn add_control(count: &mut usize) -> Result<()> {
    *count = count
        .checked_add(1)
        .ok_or_else(|| limit("control count", MAX_CONTROLS))?;
    if *count > MAX_CONTROLS {
        return Err(limit("control count", MAX_CONTROLS));
    }
    Ok(())
}

fn parse_control(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Parsed> {
    if element.attributes().with_checks(true).count() > MAX_XML_ATTRIBUTES {
        return Err(limit("control XML attributes", MAX_XML_ATTRIBUTES));
    }
    let optional = |name: &[u8], label: &'static str| -> Result<Option<String>> {
        let value = unqualified_attribute_value(element, name, decoder)?;
        if let Some(value) = &value {
            bounded(value, label)?;
        }
        Ok(value)
    };
    let show_as_icon = optional(b"showAsIcon", "control show-as-icon")?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(invalid("invalid control show-as-icon flag")),
        })
        .transpose()?;
    let number = |name: &[u8], label: &'static str| -> Result<Option<u32>> {
        optional(name, label)?
            .map(|value| {
                value
                    .parse()
                    .map_err(|_| invalid(format!("invalid {label}")))
            })
            .transpose()
    };
    Ok(Parsed {
        shape_id: optional(b"spid", "control shape ID")?,
        name: optional(b"name", "control name")?,
        show_as_icon,
        image_width: number(b"imgW", "control image width")?,
        image_height: number(b"imgH", "control image height")?,
        relationship_id: relationship_value(element, b"id", decoder, resolver)?
            .filter(|value| !value.is_empty()),
    })
}

pub(crate) fn parse_descriptor(
    xml: &[u8],
) -> Result<(String, Option<String>, Persistence, Option<String>)> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("control descriptor XML bytes", MAX_SLIDE_XML_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut root = false;
    let mut result = None;
    loop {
        let decoder = reader.decoder();
        let (_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                increment_nodes(&mut depth)?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("control descriptor depth", MAX_XML_DEPTH));
                }
                if !root {
                    root = true;
                    if element.name().local_name().as_ref() != b"ocx" {
                        return Err(invalid("control descriptor must have an ax:ocx root"));
                    }
                    let class_id = attribute(&element, b"classid", decoder)?
                        .ok_or_else(|| invalid("control descriptor is missing ax:classid"))?;
                    let license = attribute(&element, b"license", decoder)?;
                    let persistence = match attribute(&element, b"persistence", decoder)?.as_deref()
                    {
                        Some("persistPropertyBag") => Persistence::PropertyBag,
                        Some("persistStream") => Persistence::Stream,
                        Some("persistStreamInit") => Persistence::StreamInit,
                        Some("persistStorage") => Persistence::Storage,
                        Some(_) => Persistence::Unknown,
                        None => Persistence::Unknown,
                    };
                    let relationship_id =
                        relationship_value(&element, b"id", decoder, reader.resolver())?;
                    result = Some((class_id, license, persistence, relationship_id));
                }
                depth += 1;
            },
            Event::Empty(element) => {
                increment_nodes(&mut depth)?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("control descriptor depth", MAX_XML_DEPTH));
                }
                if !root {
                    if element.name().local_name().as_ref() != b"ocx" {
                        return Err(invalid("control descriptor must have an ax:ocx root"));
                    }
                    let class_id = attribute(&element, b"classid", decoder)?
                        .ok_or_else(|| invalid("control descriptor is missing ax:classid"))?;
                    let license = attribute(&element, b"license", decoder)?;
                    let persistence = match attribute(&element, b"persistence", decoder)?.as_deref()
                    {
                        Some("persistPropertyBag") => Persistence::PropertyBag,
                        Some("persistStream") => Persistence::Stream,
                        Some("persistStreamInit") => Persistence::StreamInit,
                        Some("persistStorage") => Persistence::Storage,
                        Some(_) | None => Persistence::Unknown,
                    };
                    let relationship_id =
                        relationship_value(&element, b"id", decoder, reader.resolver())?;
                    result = Some((class_id, license, persistence, relationship_id));
                }
                break;
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("invalid control descriptor nesting"));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "control descriptor rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    result.ok_or_else(|| invalid("control descriptor is empty"))
}

fn attribute(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<String>> {
    Ok(unqualified_attribute_value(element, name, decoder)?)
}
