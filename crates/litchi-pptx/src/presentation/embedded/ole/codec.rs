use super::model::Mode;
use crate::Result;
use crate::presentation::embedded::{
    MAX_XML_ATTRIBUTES, MAX_XML_BYTES, MAX_XML_DEPTH, PML, REL, STRICT_PML, STRICT_REL, bounded,
    increment_nodes, invalid, limit,
};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

pub(crate) const OLE_GRAPHIC_DATA_URI: &str =
    "http://schemas.openxmlformats.org/presentationml/2006/ole";
pub(crate) const MAX_OBJECTS: usize = 4_096;

#[derive(Debug, Clone)]
pub(crate) struct Node {
    pub(crate) namespace: Vec<u8>,
    pub(crate) local: String,
    pub(crate) attributes: Vec<Attribute>,
    pub(crate) children: Vec<Node>,
}

#[derive(Debug, Clone)]
pub(crate) struct Attribute {
    pub(crate) namespace: Vec<u8>,
    pub(crate) local: String,
    pub(crate) value: String,
}

pub(crate) fn parse_tree(xml_bytes: &[u8]) -> Result<Node> {
    if xml_bytes.len() > MAX_XML_BYTES {
        return Err(limit("OLE slide XML bytes", MAX_XML_BYTES));
    }
    let mce = Limits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_XML_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let xml = process_markup_compatibility(xml_bytes, &Capabilities::ooxml_baseline(), &mce)?.xml;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
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
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(limit("OLE XML depth", MAX_XML_DEPTH));
                }
                stack.push(make_node(&element, decoder, &resolver, &namespace)?);
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let node = make_node(&element, decoder, &resolver, &namespace)?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::End(element) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected OLE closing element"))?;
                if node.local.as_bytes() != element.local_name().as_ref() {
                    return Err(invalid("mismatched OLE XML closing element"));
                }
                attach(node, &mut stack, &mut root)?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("OLE XML rejects DTDs and processing instructions"));
            },
            Event::Text(_) | Event::CData(_) | Event::Comment(_) => {},
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated OLE XML"));
    }
    let root = root.ok_or_else(|| invalid("OLE slide has no root"))?;
    if root.local != "sld"
        || !(root.namespace.as_slice() == PML || root.namespace.as_slice() == STRICT_PML)
    {
        return Err(invalid("OLE XML must have a PresentationML sld root"));
    }
    Ok(root)
}

fn make_node(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    namespace: &ResolveResult<'_>,
) -> Result<Node> {
    if element.attributes().with_checks(true).count() > MAX_XML_ATTRIBUTES {
        return Err(limit("OLE XML attributes", MAX_XML_ATTRIBUTES));
    }
    let namespace = match namespace {
        ResolveResult::Bound(value) => value.0.to_vec(),
        ResolveResult::Unknown(prefix) => prefix.as_slice().to_vec(),
        ResolveResult::Unbound => Vec::new(),
    };
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| crate::Error::Xml(error.to_string()))?;
        let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
        let attribute_namespace = match attribute_namespace {
            ResolveResult::Bound(value) => value.0.to_vec(),
            ResolveResult::Unknown(prefix) => prefix.as_slice().to_vec(),
            ResolveResult::Unbound => Vec::new(),
        };
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        bounded(&value, "OLE XML attribute")?;
        attributes.push(Attribute {
            namespace: attribute_namespace,
            local: String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned(),
            value,
        });
    }
    Ok(Node {
        namespace,
        local: String::from_utf8_lossy(element.local_name().as_ref()).into_owned(),
        attributes,
        children: Vec::new(),
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("OLE slide has multiple roots"));
    }
    Ok(())
}

pub(crate) fn inventory(root: &Node) -> Result<Vec<Parsed>> {
    let mut result = Vec::new();
    collect_frames(root, &mut result)?;
    if result.len() > MAX_OBJECTS {
        return Err(limit("OLE object count", MAX_OBJECTS));
    }
    Ok(result)
}

#[derive(Debug, Clone)]
pub(crate) struct Parsed {
    pub(crate) shape_id: Option<u32>,
    pub(crate) shape_name: Option<String>,
    pub(crate) legacy_shape_id: Option<String>,
    pub(crate) name: Option<String>,
    pub(crate) program_id: Option<String>,
    pub(crate) show_as_icon: Option<bool>,
    pub(crate) preview_width: Option<u32>,
    pub(crate) preview_height: Option<u32>,
    pub(crate) mode: Mode,
    pub(crate) relationship_id: Option<String>,
    pub(crate) preview_relationship_id: Option<String>,
}

fn collect_frames(node: &Node, result: &mut Vec<Parsed>) -> Result<()> {
    if node.local == "graphicFrame" {
        if let Some(parsed) = parse_frame(node)? {
            result.push(parsed);
        }
    }
    for child in &node.children {
        collect_frames(child, result)?;
    }
    Ok(())
}

fn parse_frame(frame: &Node) -> Result<Option<Parsed>> {
    let Some(graphic) = child(frame, "graphic") else {
        return Ok(None);
    };
    let Some(data) = child(graphic, "graphicData") else {
        return Ok(None);
    };
    if attr(data, "uri", false) != Some(OLE_GRAPHIC_DATA_URI) {
        return Ok(None);
    }
    let Some(object) = child(data, "oleObj") else {
        return Ok(None);
    };
    let non_visual = child(frame, "nvGraphicFramePr").and_then(|node| child(node, "cNvPr"));
    let shape_id = non_visual
        .and_then(|node| attr(node, "id", false))
        .map(|value| value.parse().map_err(|_| invalid("invalid OLE shape ID")))
        .transpose()?;
    let shape_name = non_visual.and_then(|node| attr(node, "name", false).map(str::to_owned));
    let mode = if let Some(embed) = child(object, "embed") {
        if child(object, "link").is_some() {
            return Err(invalid("OLE object contains both embed and link"));
        }
        let relationship_id = attr(object, "id", true)
            .or_else(|| attr(embed, "id", true))
            .map(str::to_owned);
        (Mode::Embedded, relationship_id)
    } else if let Some(link) = child(object, "link") {
        let relationship_id = attr(object, "id", true)
            .or_else(|| attr(link, "id", true))
            .map(str::to_owned);
        (Mode::Linked, relationship_id)
    } else {
        return Err(invalid("OLE object contains neither embed nor link"));
    };
    let preview = child(object, "pic")
        .and_then(|pic| child(pic, "blipFill"))
        .and_then(|fill| child(fill, "blip"))
        .and_then(|blip| attr(blip, "embed", true).map(str::to_owned));
    Ok(Some(Parsed {
        shape_id,
        shape_name,
        legacy_shape_id: attr(object, "spid", false).map(str::to_owned),
        name: attr(object, "name", false).map(str::to_owned),
        program_id: attr(object, "progId", false).map(str::to_owned),
        show_as_icon: attr(object, "showAsIcon", false)
            .map(parse_bool)
            .transpose()?,
        preview_width: attr(object, "imgW", false)
            .map(|value| {
                value
                    .parse()
                    .map_err(|_| invalid("invalid OLE preview width"))
            })
            .transpose()?,
        preview_height: attr(object, "imgH", false)
            .map(|value| {
                value
                    .parse()
                    .map_err(|_| invalid("invalid OLE preview height"))
            })
            .transpose()?,
        mode: mode.0,
        relationship_id: mode.1,
        preview_relationship_id: preview,
    }))
}

fn child<'a>(node: &'a Node, local: &str) -> Option<&'a Node> {
    node.children.iter().find(|child| child.local == local)
}

fn attr<'a>(node: &'a Node, local: &str, relationship: bool) -> Option<&'a str> {
    node.attributes.iter().find_map(|attribute| {
        if attribute.local != local {
            return None;
        }
        if relationship {
            let relation = attribute.namespace.as_slice() == REL
                || attribute.namespace.as_slice() == STRICT_REL
                || attribute.namespace.as_slice() == b"r";
            relation.then_some(attribute.value.as_str())
        } else if attribute.namespace.is_empty() {
            Some(attribute.value.as_str())
        } else {
            None
        }
    })
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("invalid OLE boolean")),
    }
}
