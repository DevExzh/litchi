//! SpreadsheetML and ActiveX XML codecs.

use super::super::model::*;
use super::super::validation::{
    bounded, nonempty, validate_controls, validate_descriptor, validate_font,
};
use super::super::{
    AX, MAX_CONTROL_NAME_CHARS, MAX_CONTROLS, MAX_DEPTH, MAX_NODES, MAX_OUTPUT_XML, MAX_PROPERTIES,
    MAX_SHAPE_ID, MAX_XML, REL, REL_STRICT, Result, SML, SML_STRICT, X14, XDR, XDR_STRICT, invalid,
    limit, xml_error,
};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::{NsReader, XmlVersion};
use std::collections::HashSet;

impl Controls {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_XML {
            return Err(limit("worksheet XML bytes"));
        }
        let root = mce_dom(xml, true)?;
        if root.local != "worksheet" || !is_sml(&root.ns) {
            return Err(invalid("expected SpreadsheetML worksheet root"));
        }
        let mut containers = root
            .children
            .iter()
            .filter(|n| n.local == "controls" && is_sml(&n.ns));
        let Some(container) = containers.next() else {
            return Ok(Self::default());
        };
        if containers.next().is_some() {
            return Err(invalid("worksheet has multiple controls collections"));
        }
        check_attrs(container, &[])?;
        if container.children.is_empty() {
            return Err(invalid("controls requires at least one control"));
        }
        if container.children.len() > MAX_CONTROLS {
            return Err(limit("worksheet controls"));
        }
        let mut controls = Vec::with_capacity(container.children.len());
        let mut shape_ids = HashSet::new();
        let mut names = HashSet::new();
        for node in &container.children {
            if node.local != "control" || !is_sml(&node.ns) {
                return Err(invalid("unexpected child in controls"));
            }
            check_attrs(
                node,
                &[
                    ("", "shapeId"),
                    (REL, "id"),
                    (REL_STRICT, "id"),
                    ("", "name"),
                ],
            )?;
            let shape_id = req_u32(node, "", "shapeId")?;
            if !(1..=MAX_SHAPE_ID).contains(&shape_id) || !shape_ids.insert(shape_id) {
                return Err(invalid(
                    "control shapeId must be unique and within Office's supported range",
                ));
            }
            let relationship_id =
                relationship_attr(node, "id")?.ok_or_else(|| invalid("control is missing r:id"))?;
            nonempty(&relationship_id, "control relationship ID")?;
            let name = attr(node, "", "name")?;
            if let Some(name) = name.as_ref() {
                bounded(name, "control name")?;
                if name.chars().count() > MAX_CONTROL_NAME_CHARS {
                    return Err(invalid("control name exceeds Office's 32-character limit"));
                }
                if !names.insert(name.clone()) {
                    return Err(invalid("duplicate control name"));
                }
            }
            if node.children.len() > 1 {
                return Err(invalid("control permits at most one controlPr"));
            }
            let properties = node
                .children
                .first()
                .map(parse_control_properties)
                .transpose()?;
            controls.push(Control {
                shape_id,
                relationship_id,
                name,
                properties,
            });
        }
        Ok(Self { controls })
    }

    /// Writes a minimal canonical worksheet containing only the controls collection.
    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate_controls(self)?;
        let sml = if strict { SML_STRICT } else { SML };
        let rel = if strict { REL_STRICT } else { REL };
        let xdr = if strict { XDR_STRICT } else { XDR };
        let mut out = String::with_capacity(512 + self.controls.len() * 256);
        out.push_str(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"",
        );
        out.push_str(sml);
        out.push_str("\" xmlns:r=\"");
        out.push_str(rel);
        out.push_str("\" xmlns:xdr=\"");
        out.push_str(xdr);
        out.push_str("\"><controls>");
        for control in &self.controls {
            write_control(&mut out, control);
        }
        out.push_str("</controls></worksheet>");
        if out.len() > MAX_OUTPUT_XML {
            return Err(limit("canonical worksheet XML bytes"));
        }
        Ok(out.into_bytes())
    }
}

/// Replaces the direct worksheet `controls` child while preserving unrelated bytes.
///
/// An empty value removes the collection. Controls selected only through an MCE
/// `AlternateContent` branch are rejected because rewriting that branch would not
/// be byte-preserving.
pub fn replace_controls_xml(xml: &[u8], controls: &Controls) -> Result<Vec<u8>> {
    let parsed = Controls::parse(xml)?;
    if !controls.controls.is_empty() {
        validate_controls(controls)?;
    }
    let location = controls_span(xml)?;
    if !parsed.controls.is_empty() && location.span.is_none() {
        return Err(invalid(
            "MCE-selected controls cannot be mutated as a direct worksheet child",
        ));
    }
    let fragment = if controls.controls.is_empty() {
        Vec::new()
    } else {
        controls_fragment(controls, location.strict)?
    };
    let (start, end) = location
        .span
        .unwrap_or((location.insertion, location.insertion));
    let size = xml
        .len()
        .checked_sub(end - start)
        .and_then(|n| n.checked_add(fragment.len()))
        .ok_or_else(|| limit("updated worksheet XML bytes"))?;
    if size > MAX_OUTPUT_XML {
        return Err(limit("updated worksheet XML bytes"));
    }
    let mut out = Vec::with_capacity(size);
    out.extend_from_slice(&xml[..start]);
    out.extend_from_slice(&fragment);
    out.extend_from_slice(&xml[end..]);
    Ok(out)
}

impl Descriptor {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_XML {
            return Err(limit("ActiveX descriptor XML bytes"));
        }
        let root = mce_dom(xml, false)?;
        if root.local != "ocx" || root.ns != AX {
            return Err(invalid("expected ActiveX ocx root"));
        }
        check_attrs(
            &root,
            &[
                (AX, "classid"),
                (AX, "license"),
                (AX, "persistence"),
                (REL, "id"),
                (REL_STRICT, "id"),
            ],
        )?;
        let class_id = req_attr(&root, AX, "classid")?;
        let license = attr(&root, AX, "license")?;
        let persistence = parse_persistence(&req_attr(&root, AX, "persistence")?)?;
        let relationship_id = relationship_attr(&root, "id")?;
        bounded(&class_id, "ActiveX class ID")?;
        if let Some(value) = license.as_ref() {
            bounded(value, "ActiveX license")?;
        }
        let mut count = 0usize;
        let properties = parse_properties(&root.children, 0, &mut count)?;
        let value = Self {
            class_id,
            license,
            persistence,
            relationship_id,
            properties,
        };
        validate_descriptor(&value)?;
        Ok(value)
    }

    pub fn to_xml(&self) -> Result<Vec<u8>> {
        validate_descriptor(self)?;
        let mut out = String::with_capacity(512);
        out.push_str(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><ax:ocx xmlns:ax=\"",
        );
        out.push_str(AX);
        out.push_str("\" xmlns:r=\"");
        out.push_str(REL);
        out.push('"');
        qattr(&mut out, "ax:classid", &self.class_id);
        if let Some(v) = self.license.as_deref() {
            qattr(&mut out, "ax:license", v);
        }
        qattr(
            &mut out,
            "ax:persistence",
            persistence_str(self.persistence),
        );
        if let Some(v) = self.relationship_id.as_deref() {
            qattr(&mut out, "r:id", v);
        }
        if self.properties.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            for property in &self.properties {
                write_property(&mut out, property);
            }
            out.push_str("</ax:ocx>");
        }
        if out.len() > MAX_OUTPUT_XML {
            return Err(limit("canonical ActiveX XML bytes"));
        }
        Ok(out.into_bytes())
    }
}

fn controls_fragment(value: &Controls, strict: bool) -> Result<Vec<u8>> {
    validate_controls(value)?;
    let rel = if strict { REL_STRICT } else { REL };
    let xdr = if strict { XDR_STRICT } else { XDR };
    let mut out = String::with_capacity(256 + value.controls.len() * 256);
    out.push_str("<controls xmlns:r=\"");
    out.push_str(rel);
    out.push_str("\" xmlns:xdr=\"");
    out.push_str(xdr);
    out.push_str("\">");
    for control in &value.controls {
        write_control(&mut out, control);
    }
    out.push_str("</controls>");
    if out.len() > MAX_OUTPUT_XML {
        return Err(limit("controls XML bytes"));
    }
    Ok(out.into_bytes())
}

pub(crate) struct ControlsLocation {
    pub(crate) strict: bool,
    pub(crate) span: Option<(usize, usize)>,
    pub(crate) insertion: usize,
}

pub(crate) fn controls_span(xml: &[u8]) -> Result<ControlsLocation> {
    if xml.len() > MAX_XML {
        return Err(limit("worksheet XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut strict = false;
    let mut root = false;
    let mut controls_start = None;
    let mut controls_span = None;
    let mut insertion = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("worksheet XML offset overflow"))?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let namespace = resolved_ns(&resolved)?;
                if depth == 0 {
                    if element.local_name().as_ref() != b"worksheet"
                        || !matches!(namespace.as_str(), SML | SML_STRICT)
                    {
                        return Err(invalid("expected SpreadsheetML worksheet root"));
                    }
                    strict = namespace == SML_STRICT;
                    root = true;
                } else if depth == 1 && namespace == if strict { SML_STRICT } else { SML } {
                    match element.local_name().as_ref() {
                        b"controls" if controls_start.replace(start).is_some() => {
                            return Err(invalid("worksheet has multiple direct controls"));
                        },
                        b"controls" => {},
                        b"webPublishItems" | b"tableParts" | b"extLst" => {
                            insertion.get_or_insert(start);
                        },
                        _ => {},
                    }
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("worksheet XML depth"));
                }
            },
            Event::Empty(element) => {
                let namespace = resolved_ns(&resolved)?;
                if depth == 1 && namespace == if strict { SML_STRICT } else { SML } {
                    if element.local_name().as_ref() == b"controls" {
                        return Err(invalid("empty controls collection is not valid"));
                    }
                    if matches!(
                        element.local_name().as_ref(),
                        b"webPublishItems" | b"tableParts" | b"extLst"
                    ) {
                        insertion.get_or_insert(start);
                    }
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected worksheet closing element"));
                }
                if depth == 2 && element.local_name().as_ref() == b"controls" {
                    let start = controls_start
                        .take()
                        .ok_or_else(|| invalid("mismatched controls closing element"))?;
                    let end = usize::try_from(reader.buffer_position())
                        .map_err(|_| invalid("worksheet XML offset overflow"))?;
                    controls_span = Some((start, end));
                }
                if depth == 1 && element.local_name().as_ref() == b"worksheet" {
                    insertion.get_or_insert(start);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root || depth != 0 || controls_start.is_some() {
        return Err(invalid("invalid worksheet XML"));
    }
    Ok(ControlsLocation {
        strict,
        span: controls_span,
        insertion: insertion.ok_or_else(|| invalid("missing worksheet closing element"))?,
    })
}

pub(crate) fn relationship_ids_in_xml(xml: &[u8]) -> Result<HashSet<String>> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut ids = HashSet::new();
    loop {
        let resolver = reader.resolver().clone();
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(xml_error)?;
                    let (namespace, _) = resolver.resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Bound(Namespace(value)) if matches!(value, b"http://schemas.openxmlformats.org/officeDocument/2006/relationships" | b"http://purl.oclc.org/ooxml/officeDocument/relationships"))
                    {
                        ids.insert(
                            attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Explicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(xml_error)?
                                .into_owned(),
                        );
                    }
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(ids)
}

#[derive(Debug, Clone)]
struct Attribute {
    ns: String,
    local: String,
    value: String,
}
#[derive(Debug, Clone)]
struct Node {
    ns: String,
    local: String,
    attrs: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}

fn mce_dom(xml: &[u8], worksheet: bool) -> Result<Node> {
    let mut caps = Capabilities::ooxml_baseline();
    caps.understand_namespace(X14)
        .understand_namespace(XDR)
        .understand_namespace(XDR_STRICT)
        .understand_namespace(AX);
    let limits = Limits {
        max_input_bytes: MAX_XML,
        max_output_bytes: MAX_OUTPUT_XML,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &caps, &limits)?;
    parse_dom(processed.xml.as_ref(), worksheet)
}

fn parse_dom(xml: &[u8], worksheet: bool) -> Result<Node> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;
    let mut buffer = Vec::new();
    loop {
        let decoder = reader.decoder();
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(e) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                let namespace = resolved_ns(&resolved)?;
                drop(resolved);
                let resolver = reader.resolver().clone();
                stack.push(make_node(&resolver, namespace, &e, decoder)?);
            },
            Event::Empty(e) => {
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                let namespace = resolved_ns(&resolved)?;
                drop(resolved);
                let resolver = reader.resolver().clone();
                let node = make_node(&resolver, namespace, &e, decoder)?;
                append_node(&mut stack, &mut root, node)?;
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML end element"))?;
                append_node(&mut stack, &mut root, node)?;
            },
            Event::Text(e) => {
                let decoded = e.decode().map_err(xml_error)?;
                let value = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_OUTPUT_XML {
                    return Err(limit("XML text bytes"));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.text.push_str(&value);
                } else if !value.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::CData(e) => {
                let value = e.decode().map_err(xml_error)?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_OUTPUT_XML {
                    return Err(limit("XML text bytes"));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.text.push_str(&value);
                } else if !value.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(e) => {
                let name = e.decode().map_err(xml_error)?;
                let value = if let Some(c) = e.resolve_char_ref().map_err(xml_error)? {
                    c.to_string()
                } else {
                    match name.as_ref() {
                        "amp" => "&",
                        "lt" => "<",
                        "gt" => ">",
                        "apos" => "'",
                        "quot" => "\"",
                        _ => return Err(invalid("custom XML entity is rejected")),
                    }
                    .into()
                };
                if let Some(parent) = stack.last_mut() {
                    parent.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    let root = root.ok_or_else(|| invalid("missing XML root"))?;
    if worksheet && root.children.len() > MAX_NODES {
        return Err(limit("worksheet nodes"));
    }
    Ok(root)
}

fn make_node(
    resolver: &NamespaceResolver,
    ns: String,
    e: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Node> {
    let local = std::str::from_utf8(e.local_name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let mut attrs = Vec::new();
    for item in e.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, _) = resolver.resolve_attribute(item.key);
        let ans = resolved_ns(&resolved)?;
        let alocal = std::str::from_utf8(item.key.local_name().as_ref())
            .map_err(xml_error)?
            .to_string();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded(&value, "XML attribute")?;
        if attrs
            .iter()
            .any(|a: &Attribute| a.ns == ans && a.local == alocal)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attrs.push(Attribute {
            ns: ans,
            local: alocal,
            value,
        });
    }
    Ok(Node {
        ns,
        local,
        attrs,
        children: Vec::new(),
        text: String::new(),
    })
}

fn resolved_ns(value: &ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(v)) => std::str::from_utf8(v)
            .map(str::to_string)
            .map_err(xml_error),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}
fn append_node(stack: &mut [Node], root: &mut Option<Node>, node: Node) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn parse_control_properties(node: &Node) -> Result<ControlProperties> {
    if node.local != "controlPr" || !is_sml(&node.ns) {
        return Err(invalid("unexpected control child"));
    }
    check_attrs(
        node,
        &[
            ("", "locked"),
            ("", "defaultSize"),
            ("", "print"),
            ("", "disabled"),
            ("", "recalcAlways"),
            ("", "uiObject"),
            ("", "autoFill"),
            ("", "autoLine"),
            ("", "autoPict"),
            ("", "macro"),
            ("", "altText"),
            (REL, "id"),
            (REL_STRICT, "id"),
        ],
    )?;
    if node.children.len() != 1 {
        return Err(invalid("controlPr requires exactly one anchor"));
    }
    let anchor = parse_anchor(&node.children[0])?;
    let macro_name = attr(node, "", "macro")?;
    let alternate_text = attr(node, "", "altText")?;
    if let Some(v) = macro_name.as_ref() {
        bounded(v, "control macro name")?;
    }
    if let Some(v) = alternate_text.as_ref() {
        bounded(v, "control alternate text")?;
    }
    Ok(ControlProperties {
        anchor,
        locked: opt_bool(node, "locked")?,
        default_size: opt_bool(node, "defaultSize")?,
        print: opt_bool(node, "print")?,
        disabled: opt_bool(node, "disabled")?,
        recalc_always: opt_bool(node, "recalcAlways")?,
        ui_object: opt_bool(node, "uiObject")?,
        auto_fill: opt_bool(node, "autoFill")?,
        auto_line: opt_bool(node, "autoLine")?,
        auto_picture: opt_bool(node, "autoPict")?,
        macro_name,
        alternate_text,
        preview_relationship_id: relationship_attr(node, "id")?,
    })
}
fn parse_anchor(node: &Node) -> Result<ObjectAnchor> {
    if node.local != "anchor" || !is_sml(&node.ns) {
        return Err(invalid("controlPr requires anchor"));
    }
    check_attrs(node, &[("", "moveWithCells"), ("", "sizeWithCells")])?;
    if node.children.len() != 2
        || node.children[0].local != "from"
        || node.children[1].local != "to"
        || !is_sml(&node.children[0].ns)
        || !is_sml(&node.children[1].ns)
    {
        return Err(invalid("anchor requires from then to"));
    }
    Ok(ObjectAnchor {
        from: parse_marker(&node.children[0])?,
        to: parse_marker(&node.children[1])?,
        move_with_cells: opt_bool(node, "moveWithCells")?,
        size_with_cells: opt_bool(node, "sizeWithCells")?,
    })
}
fn parse_marker(node: &Node) -> Result<Marker> {
    check_attrs(node, &[])?;
    let expected = ["col", "colOff", "row", "rowOff"];
    if node.children.len() != expected.len() {
        return Err(invalid("anchor marker requires col, colOff, row, rowOff"));
    }
    for (child, expected) in node.children.iter().zip(expected) {
        if child.local != expected
            || !is_xdr(&child.ns)
            || !child.children.is_empty()
            || !child.attrs.is_empty()
        {
            return Err(invalid("invalid anchor marker grammar"));
        }
    }
    Ok(Marker {
        column: text_i32(&node.children[0])?,
        column_offset: text_i64(&node.children[1])?,
        row: text_i32(&node.children[2])?,
        row_offset: text_i64(&node.children[3])?,
    })
}

fn parse_properties(nodes: &[Node], depth: usize, count: &mut usize) -> Result<Vec<Property>> {
    if depth >= MAX_DEPTH {
        return Err(limit("ActiveX property nesting"));
    }
    let mut result = Vec::with_capacity(nodes.len());
    let mut names = HashSet::new();
    for node in nodes {
        *count = count
            .checked_add(1)
            .ok_or_else(|| limit("ActiveX properties"))?;
        if *count > MAX_PROPERTIES {
            return Err(limit("ActiveX properties"));
        }
        if node.local != "ocxPr" || node.ns != AX {
            return Err(invalid("unexpected ActiveX descriptor child"));
        }
        check_attrs(node, &[(AX, "name"), (AX, "value")])?;
        let name = req_attr(node, AX, "name")?;
        bounded(&name, "ActiveX property name")?;
        if !names.insert(name.clone()) {
            return Err(invalid("duplicate ActiveX property name"));
        }
        let value = attr(node, AX, "value")?;
        if let Some(v) = value.as_ref() {
            bounded(v, "ActiveX property value")?;
        }
        if node.children.len() > 1 {
            return Err(invalid("ActiveX property permits at most one object child"));
        }
        let object = node
            .children
            .first()
            .map(|child| parse_property_object(child, depth + 1, count))
            .transpose()?;
        if value.is_some() && object.is_some() {
            return Err(invalid(
                "ActiveX property value cannot coexist with font or picture",
            ));
        }
        result.push(Property {
            name,
            value,
            object,
        });
    }
    Ok(result)
}
fn parse_property_object(node: &Node, depth: usize, count: &mut usize) -> Result<PropertyObject> {
    if node.ns != AX {
        return Err(invalid("invalid ActiveX property object namespace"));
    }
    match node.local.as_str() {
        "font" => {
            check_attrs(
                node,
                &[(AX, "persistence"), (REL, "id"), (REL_STRICT, "id")],
            )?;
            let persistence = attr(node, AX, "persistence")?
                .map(|v| parse_persistence(&v))
                .transpose()?;
            let font = Font {
                persistence,
                relationship_id: relationship_attr(node, "id")?,
                properties: parse_properties(&node.children, depth, count)?,
            };
            validate_font(&font)?;
            Ok(PropertyObject::Font(font))
        },
        "picture" => {
            check_attrs(node, &[(REL, "id"), (REL_STRICT, "id")])?;
            if !node.children.is_empty() {
                return Err(invalid("ActiveX picture must be empty"));
            }
            Ok(PropertyObject::Picture(Picture {
                relationship_id: relationship_attr(node, "id")?,
            }))
        },
        _ => Err(invalid("ActiveX property child must be font or picture")),
    }
}

fn write_control(out: &mut String, c: &Control) {
    out.push_str("<control");
    qattr(out, "shapeId", &c.shape_id.to_string());
    qattr(out, "r:id", &c.relationship_id);
    if let Some(v) = c.name.as_deref() {
        qattr(out, "name", v);
    }
    let Some(p) = c.properties.as_ref() else {
        out.push_str("/>");
        return;
    };
    out.push_str("><controlPr");
    bool_attr(out, "locked", p.locked);
    bool_attr(out, "defaultSize", p.default_size);
    bool_attr(out, "print", p.print);
    bool_attr(out, "disabled", p.disabled);
    bool_attr(out, "recalcAlways", p.recalc_always);
    bool_attr(out, "uiObject", p.ui_object);
    bool_attr(out, "autoFill", p.auto_fill);
    bool_attr(out, "autoLine", p.auto_line);
    bool_attr(out, "autoPict", p.auto_picture);
    if let Some(v) = p.macro_name.as_deref() {
        qattr(out, "macro", v);
    }
    if let Some(v) = p.alternate_text.as_deref() {
        qattr(out, "altText", v);
    }
    if let Some(v) = p.preview_relationship_id.as_deref() {
        qattr(out, "r:id", v);
    }
    out.push_str("><anchor");
    bool_attr(out, "moveWithCells", p.anchor.move_with_cells);
    bool_attr(out, "sizeWithCells", p.anchor.size_with_cells);
    out.push('>');
    write_marker(out, "from", &p.anchor.from);
    write_marker(out, "to", &p.anchor.to);
    out.push_str("</anchor></controlPr></control>");
}
fn write_marker(out: &mut String, name: &str, m: &Marker) {
    out.push('<');
    out.push_str(name);
    out.push_str("><xdr:col>");
    out.push_str(&m.column.to_string());
    out.push_str("</xdr:col><xdr:colOff>");
    out.push_str(&m.column_offset.to_string());
    out.push_str("</xdr:colOff><xdr:row>");
    out.push_str(&m.row.to_string());
    out.push_str("</xdr:row><xdr:rowOff>");
    out.push_str(&m.row_offset.to_string());
    out.push_str("</xdr:rowOff></");
    out.push_str(name);
    out.push('>');
}
fn write_property(out: &mut String, p: &Property) {
    out.push_str("<ax:ocxPr");
    qattr(out, "ax:name", &p.name);
    if let Some(v) = p.value.as_deref() {
        qattr(out, "ax:value", v);
    }
    match p.object.as_ref() {
        None => out.push_str("/>"),
        Some(o) => {
            out.push('>');
            write_object(out, o);
            out.push_str("</ax:ocxPr>");
        },
    }
}
fn write_object(out: &mut String, value: &PropertyObject) {
    match value {
        PropertyObject::Picture(p) => {
            out.push_str("<ax:picture");
            if let Some(v) = p.relationship_id.as_deref() {
                qattr(out, "r:id", v);
            }
            out.push_str("/>");
        },
        PropertyObject::Font(f) => {
            out.push_str("<ax:font");
            if let Some(v) = f.persistence {
                qattr(out, "ax:persistence", persistence_str(v));
            }
            if let Some(v) = f.relationship_id.as_deref() {
                qattr(out, "r:id", v);
            }
            if f.properties.is_empty() {
                out.push_str("/>");
            } else {
                out.push('>');
                for p in &f.properties {
                    write_property(out, p);
                }
                out.push_str("</ax:font>");
            }
        },
    }
}
fn bool_attr(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(v) = value {
        qattr(out, name, if v { "1" } else { "0" });
    }
}
fn qattr(out: &mut String, name: &str, value: &str) {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    escape(out, value);
    out.push('"');
}
fn escape(out: &mut String, value: &str) {
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            '\r' => out.push_str("&#xD;"),
            '\n' => out.push_str("&#xA;"),
            '\t' => out.push_str("&#x9;"),
            _ => out.push(c),
        }
    }
}

fn check_attrs(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    for a in &node.attrs {
        if !allowed
            .iter()
            .any(|(ns, local)| *ns == a.ns && *local == a.local)
        {
            return Err(invalid(format!(
                "unexpected attribute {{{}}}{} on {}",
                a.ns, a.local, node.local
            )));
        }
    }
    Ok(())
}
fn attr(node: &Node, ns: &str, local: &str) -> Result<Option<String>> {
    let mut values = node.attrs.iter().filter(|a| a.ns == ns && a.local == local);
    let value = values.next().map(|a| a.value.clone());
    if values.next().is_some() {
        return Err(invalid("duplicate attribute"));
    }
    Ok(value)
}
fn req_attr(node: &Node, ns: &str, local: &str) -> Result<String> {
    attr(node, ns, local)?
        .filter(|v| !v.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing required {}", node.local, local)))
}
fn relationship_attr(node: &Node, local: &str) -> Result<Option<String>> {
    let a = attr(node, REL, local)?;
    let b = attr(node, REL_STRICT, local)?;
    if a.is_some() && b.is_some() {
        return Err(invalid("duplicate relationship attribute"));
    }
    Ok(a.or(b))
}
fn req_u32(node: &Node, ns: &str, local: &str) -> Result<u32> {
    req_attr(node, ns, local)?
        .parse()
        .map_err(|_| invalid(format!("invalid unsigned integer {local}")))
}
fn opt_bool(node: &Node, local: &str) -> Result<Option<bool>> {
    attr(node, "", local)?.map(|v| parse_bool(&v)).transpose()
}
fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid("invalid XML boolean")),
    }
}
fn text_i32(node: &Node) -> Result<i32> {
    node.text
        .trim()
        .parse()
        .map_err(|_| invalid("invalid anchor signed integer"))
}
fn text_i64(node: &Node) -> Result<i64> {
    node.text
        .trim()
        .parse()
        .map_err(|_| invalid("invalid anchor coordinate"))
}
fn parse_persistence(value: &str) -> Result<Persistence> {
    match value {
        "persistPropertyBag" => Ok(Persistence::PropertyBag),
        "persistStream" => Ok(Persistence::Stream),
        "persistStreamInit" => Ok(Persistence::StreamInit),
        "persistStorage" => Ok(Persistence::Storage),
        _ => Err(invalid("invalid ActiveX persistence")),
    }
}
fn persistence_str(value: Persistence) -> &'static str {
    match value {
        Persistence::PropertyBag => "persistPropertyBag",
        Persistence::Stream => "persistStream",
        Persistence::StreamInit => "persistStreamInit",
        Persistence::Storage => "persistStorage",
    }
}
fn is_sml(ns: &str) -> bool {
    matches!(ns, SML | SML_STRICT)
}
fn is_xdr(ns: &str) -> bool {
    matches!(ns, XDR | XDR_STRICT)
}
