//! Bounded SpreadsheetML OLE markup codec.

use crate::error::Result;
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::*;
use super::{invalid, limit, xml_error};

#[derive(Clone)]
struct Attribute {
    namespace: String,
    name: String,
    value: String,
}
#[derive(Clone)]
pub(super) struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}

pub fn parse_ole_objects(xml: &[u8]) -> Result<Option<OleObjects>> {
    let root = parse_document(xml)?;
    let conformance = conformance(&root)?;
    let mut lists = root
        .children
        .iter()
        .filter(|child| child.namespace == conformance.sml() && child.name == "oleObjects");
    let Some(list) = lists.next() else {
        return Ok(None);
    };
    if lists.next().is_some() {
        return Err(invalid("worksheet has multiple oleObjects collections"));
    }
    no_attributes(list, &[])?;
    whitespace(list)?;
    if list.children.len() > MAX_OBJECTS {
        return Err(limit("object count"));
    }
    let mut objects = Vec::with_capacity(list.children.len());
    for child in &list.children {
        require(child, conformance.sml(), "oleObject")?;
        objects.push(parse_object(child, conformance)?);
    }
    let value = OleObjects { objects };
    validate_value(&value, false)?;
    Ok(Some(value))
}

fn parse_object(node: &Node, conformance: OleObjectConformance) -> Result<OleObject> {
    whitespace(node)?;
    if node.children.len() > 1 {
        return Err(invalid("oleObject has multiple child elements"));
    }
    let program_id = optional(node, "", "progId").map(str::to_owned);
    let data_or_view_aspect = optional(node, "", "dvAspect")
        .map(OleObjectAspect::try_from)
        .transpose()?;
    let link = optional(node, "", "link").map(str::to_owned);
    let update = optional(node, "", "oleUpdate")
        .map(str::parse)
        .transpose()?;
    let auto_load = optional(node, "", "autoLoad")
        .map(|value| parse_bool(value, "autoLoad"))
        .transpose()?;
    let shape_id = required(node, "", "shapeId")?
        .parse()
        .map_err(|_| invalid("invalid oleObject shapeId"))?;
    let relationship_id = required(node, conformance.rel(), "id")?.to_owned();
    no_attributes(
        node,
        &[
            ("", "progId"),
            ("", "dvAspect"),
            ("", "link"),
            ("", "oleUpdate"),
            ("", "autoLoad"),
            ("", "shapeId"),
            (conformance.rel(), "id"),
        ],
    )?;
    let properties = node
        .children
        .first()
        .map(|child| parse_properties(child, conformance))
        .transpose()?;
    Ok(OleObject {
        program_id,
        data_or_view_aspect,
        link,
        update,
        auto_load,
        shape_id,
        relationship_id,
        relationship_kind: OleObjectRelationshipKind::OleObject,
        target: None,
        properties,
    })
}

fn parse_properties(node: &Node, conformance: OleObjectConformance) -> Result<OleObjectProperties> {
    require(node, conformance.sml(), "objectPr")?;
    whitespace(node)?;
    if node.children.len() != 1 {
        return Err(invalid("objectPr must contain exactly one anchor"));
    }
    let preview_relationship_id = required(node, conformance.rel(), "id")?.to_owned();
    let boolean = |name| {
        optional(node, "", name)
            .map(|value| parse_bool(value, name))
            .transpose()
    };
    let value = OleObjectProperties {
        preview_relationship_id,
        preview: None,
        default_size: boolean("defaultSize")?,
        print: boolean("print")?,
        disabled: boolean("disabled")?,
        ui_object: boolean("uiObject")?,
        auto_fill: boolean("autoFill")?,
        auto_line: boolean("autoLine")?,
        auto_pict: boolean("autoPict")?,
        dde: boolean("dde")?,
        macro_name: optional(node, "", "macro").map(str::to_owned),
        alt_text: optional(node, "", "altText").map(str::to_owned),
        anchor: parse_anchor(&node.children[0], conformance)?,
    };
    no_attributes(
        node,
        &[
            (conformance.rel(), "id"),
            ("", "defaultSize"),
            ("", "print"),
            ("", "disabled"),
            ("", "uiObject"),
            ("", "autoFill"),
            ("", "autoLine"),
            ("", "autoPict"),
            ("", "dde"),
            ("", "macro"),
            ("", "altText"),
        ],
    )?;
    Ok(value)
}

fn parse_anchor(node: &Node, conformance: OleObjectConformance) -> Result<OleObjectAnchor> {
    require(node, conformance.sml(), "anchor")?;
    whitespace(node)?;
    no_attributes(node, &[("", "moveWithCells"), ("", "sizeWithCells")])?;
    if node.children.len() != 2 {
        return Err(invalid("object anchor requires from and to markers"));
    }
    require(&node.children[0], conformance.sml(), "from")?;
    require(&node.children[1], conformance.sml(), "to")?;
    Ok(OleObjectAnchor {
        move_with_cells: optional(node, "", "moveWithCells")
            .map(|value| parse_bool(value, "moveWithCells"))
            .transpose()?,
        size_with_cells: optional(node, "", "sizeWithCells")
            .map(|value| parse_bool(value, "sizeWithCells"))
            .transpose()?,
        from: parse_marker(&node.children[0], conformance)?,
        to: parse_marker(&node.children[1], conformance)?,
    })
}

fn parse_marker(node: &Node, conformance: OleObjectConformance) -> Result<OleObjectMarker> {
    whitespace(node)?;
    no_attributes(node, &[])?;
    let expected = ["col", "colOff", "row", "rowOff"];
    if node.children.len() != expected.len() {
        return Err(invalid(
            "object marker requires col, colOff, row, and rowOff",
        ));
    }
    for (child, name) in node.children.iter().zip(expected) {
        require(child, conformance.xdr(), name)?;
        no_attributes(child, &[])?;
        if !child.children.is_empty() {
            return Err(invalid("object marker coordinate must be text-only"));
        }
    }
    Ok(OleObjectMarker {
        column: coordinate(&node.children[0], "column")?,
        column_offset: signed_coordinate(&node.children[1], "column offset")?,
        row: coordinate(&node.children[2], "row")?,
        row_offset: signed_coordinate(&node.children[3], "row offset")?,
    })
}

/// Deterministically serializes a self-contained `oleObjects` fragment.
pub fn write_ole_objects(value: &OleObjects, conformance: OleObjectConformance) -> Result<Vec<u8>> {
    validate_value(value, false)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<x:oleObjects xmlns:x=\"");
    escape(&mut output, conformance.sml());
    output.extend_from_slice(b"\" xmlns:r=\"");
    escape(&mut output, conformance.rel());
    output.extend_from_slice(b"\" xmlns:xdr=\"");
    escape(&mut output, conformance.xdr());
    if value.objects.is_empty() {
        output.extend_from_slice(b"\"/>");
        return Ok(output);
    }
    output.extend_from_slice(b"\">");
    for object in &value.objects {
        output.extend_from_slice(b"<x:oleObject");
        if let Some(value) = &object.program_id {
            attr(&mut output, "progId", value);
        }
        if let Some(value) = object.data_or_view_aspect {
            attr(&mut output, "dvAspect", value.as_str());
        }
        if let Some(value) = &object.link {
            attr(&mut output, "link", value);
        }
        if let Some(value) = object.update {
            attr(&mut output, "oleUpdate", value.as_str());
        }
        if let Some(value) = object.auto_load {
            bool_attr(&mut output, "autoLoad", value);
        }
        attr(&mut output, "shapeId", &object.shape_id.to_string());
        attr(&mut output, "r:id", &object.relationship_id);
        let Some(properties) = &object.properties else {
            output.extend_from_slice(b"/>");
            continue;
        };
        output.extend_from_slice(b"><x:objectPr");
        attr(&mut output, "r:id", &properties.preview_relationship_id);
        for (name, value) in [
            ("defaultSize", properties.default_size),
            ("print", properties.print),
            ("disabled", properties.disabled),
            ("uiObject", properties.ui_object),
            ("autoFill", properties.auto_fill),
            ("autoLine", properties.auto_line),
            ("autoPict", properties.auto_pict),
            ("dde", properties.dde),
        ] {
            if let Some(value) = value {
                bool_attr(&mut output, name, value);
            }
        }
        if let Some(value) = &properties.macro_name {
            attr(&mut output, "macro", value);
        }
        if let Some(value) = &properties.alt_text {
            attr(&mut output, "altText", value);
        }
        output.extend_from_slice(b"><x:anchor");
        if let Some(value) = properties.anchor.move_with_cells {
            bool_attr(&mut output, "moveWithCells", value);
        }
        if let Some(value) = properties.anchor.size_with_cells {
            bool_attr(&mut output, "sizeWithCells", value);
        }
        output.push(b'>');
        write_marker(&mut output, "from", &properties.anchor.from);
        write_marker(&mut output, "to", &properties.anchor.to);
        output.extend_from_slice(b"</x:anchor></x:objectPr></x:oleObject>");
    }
    output.extend_from_slice(b"</x:oleObjects>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized XML bytes"));
    }
    Ok(output)
}

fn write_marker(output: &mut Vec<u8>, name: &str, marker: &OleObjectMarker) {
    output.extend_from_slice(b"<x:");
    output.extend_from_slice(name.as_bytes());
    output.push(b'>');
    for (name, value) in [
        ("col", marker.column.to_string()),
        ("colOff", marker.column_offset.to_string()),
        ("row", marker.row.to_string()),
        ("rowOff", marker.row_offset.to_string()),
    ] {
        output.extend_from_slice(b"<xdr:");
        output.extend_from_slice(name.as_bytes());
        output.push(b'>');
        output.extend_from_slice(value.as_bytes());
        output.extend_from_slice(b"</xdr:");
        output.extend_from_slice(name.as_bytes());
        output.push(b'>');
    }
    output.extend_from_slice(b"</x:");
    output.extend_from_slice(name.as_bytes());
    output.push(b'>');
}

/// Loads OLE anchors and validates all payload and preview relationships for one worksheet.

pub(super) fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("input XML bytes"));
    }
    let mut caps = Capabilities::ooxml_baseline();
    caps.understand_namespace(X14);
    let limits = Limits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &caps, &limits)?.xml;
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let node = make_node(&reader, element, reader.decoder(), &mut strings)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside worksheet root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|value| value.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside worksheet root"));
                }
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected in worksheet OLE markup")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated worksheet XML"));
    }
    root.ok_or_else(|| invalid("missing worksheet root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + name.len())?;
    let mut attributes = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let qname = item.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if attributes
            .iter()
            .any(|attribute: &Attribute| attribute.namespace == namespace && attribute.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}
fn conformance(root: &Node) -> Result<OleObjectConformance> {
    crate_conformance(root)
}
pub(super) fn crate_conformance(root: &Node) -> Result<OleObjectConformance> {
    if root.name != "worksheet" {
        return Err(invalid("expected worksheet root"));
    }
    match root.namespace.as_str() {
        SML => Ok(OleObjectConformance::Transitional),
        STRICT_SML => Ok(OleObjectConformance::Strict),
        _ => Err(invalid("unsupported worksheet namespace")),
    }
}

pub(super) fn insert_collection(
    xml: &[u8],
    fragment: &[u8],
    conformance: OleObjectConformance,
) -> Result<Vec<u8>> {
    let later = [
        b"controls".as_slice(),
        b"webPublishItems",
        b"tableParts",
        b"extLst",
    ];
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut position = None;
    let mut root = false;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("worksheet XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == conformance.sml().as_bytes());
                if depth == 0 {
                    if !core || element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("worksheet root does not match conformance"));
                    }
                    root = true;
                } else if depth == 1 && core && later.contains(&element.local_name().as_ref()) {
                    position.get_or_insert(start);
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
            },
            Event::Empty(element) => {
                let core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == conformance.sml().as_bytes());
                if depth == 1 && core && later.contains(&element.local_name().as_ref()) {
                    position.get_or_insert(start);
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected worksheet closing element"));
                }
                if depth == 1 && element.local_name().as_ref() == b"worksheet" {
                    position.get_or_insert(start);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root || depth != 0 {
        return Err(invalid("invalid worksheet XML"));
    }
    let position = position.ok_or_else(|| invalid("missing worksheet closing element"))?;
    let size = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| limit("updated XML bytes"))?;
    if size > MAX_XML_BYTES {
        return Err(limit("updated XML bytes"));
    }
    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

fn require(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!("expected {name}, got {}", node.name)))
    }
}
fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}
fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name)))
}
fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    if let Some(attribute) = node.attributes.iter().find(|attribute| {
        !allowed.contains(&(attribute.namespace.as_str(), attribute.name.as_str()))
    }) {
        Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )))
    } else {
        Ok(())
    }
}
fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}
fn coordinate(node: &Node, name: &str) -> Result<u32> {
    node.text
        .trim()
        .parse()
        .map_err(|_| invalid(format!("invalid object marker {name}")))
}
fn signed_coordinate(node: &Node, name: &str) -> Result<i64> {
    node.text
        .trim()
        .parse()
        .map_err(|_| invalid(format!("invalid object marker {name}")))
}
fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid boolean '{value}' for {name}"))),
    }
}
fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
fn bool_attr(output: &mut Vec<u8>, name: &str, value: bool) {
    attr(output, name, if value { "1" } else { "0" });
}
fn attr(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'\"');
}
fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
