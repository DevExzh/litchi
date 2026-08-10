//! Bounded `SpreadsheetML` OLE markup codec.

use crate::error::Result;
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::cmp::Reverse;

use super::model::{
    Aspect, MAX_DEPTH, MAX_NODES, MAX_OBJECTS, MAX_STRING_BYTES, MAX_XML_BYTES, OleObject,
    OleObjectAnchor, OleObjectConformance, OleObjectMarker, OleObjectProperties,
    OleObjectRelationshipKind, OleObjectUpdate, OleObjects, SML, STRICT_SML, X14, validate_value,
};
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
        if child.namespace == conformance.sml() && child.name == "oleObject" {
            objects.push(parse_object(child, conformance)?);
        }
    }
    let value = OleObjects { objects };
    validate_value(&value, false)?;
    Ok(Some(value))
}

fn parse_object(node: &Node, conformance: OleObjectConformance) -> Result<OleObject> {
    whitespace(node)?;
    let object_properties = node
        .children
        .iter()
        .filter(|child| child.namespace == conformance.sml() && child.name == "objectPr")
        .collect::<Vec<_>>();
    if object_properties.len() > 1 {
        return Err(invalid("oleObject has multiple child elements"));
    }
    let program_id = optional(node, "", "progId").map(str::to_owned);
    let data_or_view_aspect = optional(node, "", "dvAspect")
        .map(Aspect::try_from)
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
        .map_err(|_source| invalid("invalid oleObject shapeId"))?;
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
    let properties = object_properties
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
    let anchors = node
        .children
        .iter()
        .filter(|child| child.namespace == conformance.sml() && child.name == "anchor")
        .collect::<Vec<_>>();
    if anchors.len() != 1 {
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
        anchor: parse_anchor(anchors[0], conformance)?,
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
    let markers = ["from", "to"].map(|name| {
        node.children
            .iter()
            .filter(|child| child.namespace == conformance.sml() && child.name == name)
            .collect::<Vec<_>>()
    });
    if markers.iter().any(|markers| markers.len() != 1) {
        return Err(invalid("object anchor requires from and to markers"));
    }
    Ok(OleObjectAnchor {
        move_with_cells: optional(node, "", "moveWithCells")
            .map(|value| parse_bool(value, "moveWithCells"))
            .transpose()?,
        size_with_cells: optional(node, "", "sizeWithCells")
            .map(|value| parse_bool(value, "sizeWithCells"))
            .transpose()?,
        from: parse_marker(markers[0][0], conformance)?,
        to: parse_marker(markers[1][0], conformance)?,
    })
}

fn parse_marker(node: &Node, conformance: OleObjectConformance) -> Result<OleObjectMarker> {
    whitespace(node)?;
    no_attributes(node, &[])?;
    let expected = ["col", "colOff", "row", "rowOff"];
    let coordinates = expected
        .iter()
        .map(|name| {
            node.children
                .iter()
                .filter(|child| child.namespace == conformance.xdr() && child.name == *name)
                .collect::<Vec<_>>()
        })
        .collect::<Vec<_>>();
    if coordinates.iter().any(|coordinates| coordinates.len() != 1) {
        return Err(invalid(
            "object marker requires col, colOff, row, and rowOff",
        ));
    }
    for coordinate in &coordinates {
        let child = coordinate[0];
        no_attributes(child, &[])?;
        if !child.children.is_empty() {
            return Err(invalid("object marker coordinate must be text-only"));
        }
    }
    Ok(OleObjectMarker {
        column: coordinate(coordinates[0][0], "column")?,
        column_offset: signed_coordinate(coordinates[1][0], "column offset")?,
        row: coordinate(coordinates[2][0], "row")?,
        row_offset: signed_coordinate(coordinates[3][0], "row offset")?,
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

/// Patch typed OLE metadata into an existing worksheet without rebuilding its
/// `oleObjects` subtree. Only known attributes and marker text are replaced;
/// extension attributes, unknown children, comments, namespace choices, and
/// fallback branches stay in their original byte representation.
pub(super) fn patch_ole_objects_source(
    source: &[u8],
    before: &OleObjects,
    after: &OleObjects,
    conformance: OleObjectConformance,
) -> Result<Vec<u8>> {
    if before == after {
        return Ok(source.to_vec());
    }
    validate_value(after, false)?;
    compatible_source_edit(before, after)?;
    let raw = collect_raw_source(source, conformance)?;
    let mut edits = Vec::new();

    for (before_object, after_object) in before.objects.iter().zip(&after.objects) {
        let key = RawObjectKey {
            shape_id: before_object.shape_id,
            relationship_id: before_object.relationship_id.clone(),
        };
        let object_elements = raw
            .elements
            .iter()
            .filter(|element| {
                element.namespace == conformance.sml()
                    && element.name == "oleObject"
                    && element.object.as_ref() == Some(&key)
            })
            .collect::<Vec<_>>();
        if object_elements.is_empty() {
            return Err(invalid(format!(
                "OLE object shapeId {} is absent from worksheet source",
                key.shape_id
            )));
        }

        patch_optional_attribute(
            source,
            &mut edits,
            object_elements.as_slice(),
            "progId",
            before_object.program_id.as_deref(),
            after_object.program_id.as_deref(),
        )?;
        patch_optional_attribute(
            source,
            &mut edits,
            object_elements.as_slice(),
            "dvAspect",
            before_object.data_or_view_aspect.map(Aspect::as_str),
            after_object.data_or_view_aspect.map(Aspect::as_str),
        )?;
        patch_optional_attribute(
            source,
            &mut edits,
            object_elements.as_slice(),
            "link",
            before_object.link.as_deref(),
            after_object.link.as_deref(),
        )?;
        patch_optional_attribute(
            source,
            &mut edits,
            object_elements.as_slice(),
            "oleUpdate",
            before_object.update.map(OleObjectUpdate::as_str),
            after_object.update.map(OleObjectUpdate::as_str),
        )?;
        patch_optional_bool_attribute(
            source,
            &mut edits,
            object_elements.as_slice(),
            "autoLoad",
            before_object.auto_load,
            after_object.auto_load,
        )?;

        let Some(before_properties) = before_object.properties.as_ref() else {
            continue;
        };
        let after_properties = after_object
            .properties
            .as_ref()
            .ok_or_else(|| invalid("OLE object properties cannot be removed by a source edit"))?;
        let property_elements = raw
            .elements
            .iter()
            .filter(|element| {
                element.namespace == conformance.sml()
                    && element.name == "objectPr"
                    && element.object.as_ref() == Some(&key)
                    && element
                        .parent
                        .is_some_and(|parent| raw.elements[parent].object.as_ref() == Some(&key))
            })
            .collect::<Vec<_>>();
        for property in property_elements {
            patch_optional_bool_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "defaultSize",
                before_properties.default_size,
                after_properties.default_size,
            )?;
            patch_optional_bool_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "print",
                before_properties.print,
                after_properties.print,
            )?;
            patch_optional_bool_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "disabled",
                before_properties.disabled,
                after_properties.disabled,
            )?;
            patch_optional_bool_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "uiObject",
                before_properties.ui_object,
                after_properties.ui_object,
            )?;
            patch_optional_bool_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "autoFill",
                before_properties.auto_fill,
                after_properties.auto_fill,
            )?;
            patch_optional_bool_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "autoLine",
                before_properties.auto_line,
                after_properties.auto_line,
            )?;
            patch_optional_bool_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "autoPict",
                before_properties.auto_pict,
                after_properties.auto_pict,
            )?;
            patch_optional_bool_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "dde",
                before_properties.dde,
                after_properties.dde,
            )?;
            patch_optional_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "macro",
                before_properties.macro_name.as_deref(),
                after_properties.macro_name.as_deref(),
            )?;
            patch_optional_attribute(
                source,
                &mut edits,
                std::slice::from_ref(&property),
                "altText",
                before_properties.alt_text.as_deref(),
                after_properties.alt_text.as_deref(),
            )?;

            let property_id = property_id(&raw, property)?;
            let anchors = raw
                .elements
                .iter()
                .filter(|element| {
                    element.namespace == conformance.sml()
                        && element.name == "anchor"
                        && element.parent == Some(property_id)
                })
                .collect::<Vec<_>>();
            for anchor in anchors {
                patch_optional_bool_attribute(
                    source,
                    &mut edits,
                    std::slice::from_ref(&anchor),
                    "moveWithCells",
                    before_properties.anchor.move_with_cells,
                    after_properties.anchor.move_with_cells,
                )?;
                patch_optional_bool_attribute(
                    source,
                    &mut edits,
                    std::slice::from_ref(&anchor),
                    "sizeWithCells",
                    before_properties.anchor.size_with_cells,
                    after_properties.anchor.size_with_cells,
                )?;
                patch_anchor_marker(
                    source,
                    &mut edits,
                    &raw,
                    anchor,
                    "from",
                    &before_properties.anchor.from,
                    &after_properties.anchor.from,
                    conformance,
                )?;
                patch_anchor_marker(
                    source,
                    &mut edits,
                    &raw,
                    anchor,
                    "to",
                    &before_properties.anchor.to,
                    &after_properties.anchor.to,
                    conformance,
                )?;
            }
        }
    }

    apply_edits(source, edits)
}

fn compatible_source_edit(before: &OleObjects, after: &OleObjects) -> Result<()> {
    if before.objects.len() != after.objects.len() {
        return Err(invalid(
            "OLE source edits cannot add or remove objects; use package storage for topology changes",
        ));
    }
    for (before_object, after_object) in before.objects.iter().zip(&after.objects) {
        if before_object.shape_id != after_object.shape_id
            || before_object.relationship_id != after_object.relationship_id
            || before_object.relationship_kind != after_object.relationship_kind
            || before_object.target != after_object.target
        {
            return Err(invalid(
                "OLE source edits cannot change shape, relationship, or opaque payload identity",
            ));
        }
        match (&before_object.properties, &after_object.properties) {
            (None, None) => {},
            (Some(before), Some(after))
                if before.preview_relationship_id == after.preview_relationship_id
                    && before.preview == after.preview => {},
            (Some(_), Some(_)) => {
                return Err(invalid(
                    "OLE source edits cannot change preview relationship or opaque preview data",
                ));
            },
            _ => {
                return Err(invalid(
                    "OLE source edits cannot add or remove object properties",
                ));
            },
        }
    }
    Ok(())
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RawObjectKey {
    shape_id: u32,
    relationship_id: String,
}

#[derive(Clone, Debug)]
struct RawAttribute {
    namespace: String,
    name: String,
    value_start: usize,
    value_end: usize,
    remove_start: usize,
    end: usize,
}

#[derive(Clone, Debug)]
struct RawElement {
    namespace: String,
    name: String,
    parent: Option<usize>,
    object: Option<RawObjectKey>,
    start_end: usize,
    end_start: Option<usize>,
    empty: bool,
    attributes: Vec<RawAttribute>,
}

#[derive(Clone, Debug)]
struct RawSource {
    elements: Vec<RawElement>,
}

fn collect_raw_source(source: &[u8], conformance: OleObjectConformance) -> Result<RawSource> {
    if source.len() > MAX_XML_BYTES {
        return Err(limit("input XML bytes"));
    }
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<usize> = Vec::new();
    let mut elements: Vec<RawElement> = Vec::new();
    let mut nodes = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_source| invalid("worksheet XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let namespace = resolved(namespace)?;
                let local_name = element.local_name();
                let name = std::str::from_utf8(local_name.as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                let object = if namespace == conformance.sml() && name == "oleObject" {
                    Some(raw_object_key(&reader, element, conformance)?)
                } else {
                    stack.last().and_then(|id| elements[*id].object.clone())
                };
                let parent = stack.last().copied();
                let end = usize::try_from(reader.buffer_position())
                    .map_err(|_source| invalid("worksheet XML offset overflow"))?;
                let attributes = raw_attributes(source, start, end, &reader, element)?;
                let id = elements.len();
                elements.push(RawElement {
                    namespace,
                    name,
                    parent,
                    object,
                    start_end: usize::try_from(reader.buffer_position())
                        .map_err(|_source| invalid("worksheet XML offset overflow"))?,
                    end_start: None,
                    empty,
                    attributes,
                });
                if !empty {
                    stack.push(id);
                }
            },
            Event::End(element) => {
                let id = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected worksheet closing element"))?;
                let local_name = element.local_name();
                let name = std::str::from_utf8(local_name.as_ref()).map_err(xml_error)?;
                if elements[id].name != name {
                    return Err(invalid("worksheet XML element nesting changed during scan"));
                }
                elements[id].end_start = Some(start);
            },
            Event::DocType(_) | Event::PI(_) | Event::CData(_) => {
                return Err(invalid("unsafe XML construct in worksheet OLE source"));
            },
            Event::Eof => break,
            Event::Text(_) | Event::Comment(_) | Event::Decl(_) | Event::GeneralRef(_) => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated worksheet XML"));
    }
    Ok(RawSource { elements })
}

fn raw_object_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: OleObjectConformance,
) -> Result<RawObjectKey> {
    let shape_id = raw_attribute_value(reader, element, "", "shapeId")?
        .ok_or_else(|| invalid("OLE source object is missing shapeId"))?
        .parse()
        .map_err(|_source| invalid("invalid OLE source shapeId"))?;
    let relationship_id = raw_attribute_value(reader, element, conformance.rel(), "id")?
        .ok_or_else(|| invalid("OLE source object is missing relationship ID"))?;
    Ok(RawObjectKey {
        shape_id,
        relationship_id,
    })
}

fn raw_attribute_value(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    name: &str,
) -> Result<Option<String>> {
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        if item.key.as_ref() == b"xmlns" || item.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (attribute_namespace, local) = reader.resolver().resolve_attribute(item.key);
        let attribute_namespace = resolved(attribute_namespace)?;
        let local = std::str::from_utf8(local.as_ref()).map_err(xml_error)?;
        if attribute_namespace == namespace && local == name {
            return Ok(Some(
                item.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(xml_error)?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn raw_attributes(
    source: &[u8],
    start: usize,
    end: usize,
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Vec<RawAttribute>> {
    let mut expanded = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let qname = item.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        expanded.push((
            qname.to_vec(),
            resolved(namespace)?,
            std::str::from_utf8(local.as_ref())
                .map_err(xml_error)?
                .to_owned(),
        ));
    }

    let mut cursor = start
        .checked_add(1)
        .ok_or_else(|| invalid("worksheet XML offset overflow"))?;
    while cursor < end && source[cursor].is_ascii_whitespace() {
        cursor += 1;
    }
    while cursor < end
        && !source[cursor].is_ascii_whitespace()
        && source[cursor] != b'>'
        && source[cursor] != b'/'
    {
        cursor += 1;
    }
    let mut result = Vec::with_capacity(expanded.len());
    let tag_end = end
        .checked_sub(1)
        .ok_or_else(|| invalid("empty worksheet XML element"))?;
    let tag_end = if source.get(tag_end.wrapping_sub(1)) == Some(&b'/') {
        tag_end - 1
    } else {
        tag_end
    };
    while cursor < tag_end {
        while cursor < tag_end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag_end || source[cursor] == b'/' {
            break;
        }
        let attribute_start = cursor;
        while cursor < tag_end
            && !source[cursor].is_ascii_whitespace()
            && source[cursor] != b'='
            && source[cursor] != b'/'
            && source[cursor] != b'>'
        {
            cursor += 1;
        }
        let qname = &source[attribute_start..cursor];
        while cursor < tag_end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag_end || source[cursor] != b'=' {
            return Err(invalid("malformed worksheet XML attribute"));
        }
        cursor += 1;
        while cursor < tag_end && source[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *source
            .get(cursor)
            .ok_or_else(|| invalid("worksheet XML attribute has no value"))?;
        if quote != b'\'' && quote != b'"' {
            return Err(invalid("worksheet XML attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < end && source[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= end {
            return Err(invalid("unterminated worksheet XML attribute"));
        }
        cursor += 1;
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let Some((_, namespace, name)) = expanded.iter().find(|(name, _, _)| name == qname) else {
            return Err(invalid(
                "worksheet XML attribute scan disagrees with parser",
            ));
        };
        let remove_start =
            if attribute_start > start + 1 && source[attribute_start - 1].is_ascii_whitespace() {
                attribute_start - 1
            } else {
                attribute_start
            };
        result.push(RawAttribute {
            namespace: namespace.clone(),
            name: name.clone(),
            value_start,
            value_end,
            remove_start,
            end: cursor,
        });
    }
    if result.len() != expanded.len() {
        return Err(invalid("worksheet XML attribute count changed during scan"));
    }
    Ok(result)
}

fn property_id(raw: &RawSource, property: &RawElement) -> Result<usize> {
    raw.elements
        .iter()
        .position(|candidate| std::ptr::eq(candidate, property))
        .ok_or_else(|| invalid("OLE source element is not owned by its scan"))
}

fn patch_anchor_marker(
    source: &[u8],
    edits: &mut Vec<Edit>,
    raw: &RawSource,
    anchor: &RawElement,
    name: &str,
    before: &OleObjectMarker,
    after: &OleObjectMarker,
    conformance: OleObjectConformance,
) -> Result<()> {
    let anchor_id = property_id(raw, anchor)?;
    let marker = raw
        .elements
        .iter()
        .find(|element| {
            element.namespace == conformance.sml()
                && element.name == name
                && element.parent == Some(anchor_id)
        })
        .ok_or_else(|| invalid(format!("OLE source anchor is missing {name} marker")))?;
    let marker_id = property_id(raw, marker)?;
    for (coordinate_name, before_value, after_value) in [
        ("col", i128::from(before.column), i128::from(after.column)),
        (
            "colOff",
            i128::from(before.column_offset),
            i128::from(after.column_offset),
        ),
        ("row", i128::from(before.row), i128::from(after.row)),
        (
            "rowOff",
            i128::from(before.row_offset),
            i128::from(after.row_offset),
        ),
    ] {
        if before_value == after_value {
            continue;
        }
        let coordinate = raw
            .elements
            .iter()
            .find(|element| {
                element.namespace == conformance.xdr()
                    && element.name == coordinate_name
                    && element.parent == Some(marker_id)
            })
            .ok_or_else(|| {
                invalid(format!(
                    "OLE source anchor is missing {name}/{coordinate_name}"
                ))
            })?;
        let start = coordinate.start_end;
        let end = coordinate
            .end_start
            .ok_or_else(|| invalid("OLE source coordinate is not a text element"))?;
        let (text_start, text_end) = trimmed_text_span(source, start, end)?;
        add_edit(
            edits,
            text_start,
            text_end,
            after_value.to_string().into_bytes(),
        )?;
    }
    Ok(())
}

fn patch_optional_bool_attribute(
    source: &[u8],
    edits: &mut Vec<Edit>,
    elements: &[&RawElement],
    name: &str,
    before: Option<bool>,
    after: Option<bool>,
) -> Result<()> {
    let before = before.map(|value| if value { "1" } else { "0" });
    let after = after.map(|value| if value { "1" } else { "0" });
    patch_optional_attribute(source, edits, elements, name, before, after)
}

fn patch_optional_attribute(
    source: &[u8],
    edits: &mut Vec<Edit>,
    elements: &[&RawElement],
    name: &str,
    before: Option<&str>,
    after: Option<&str>,
) -> Result<()> {
    if before == after {
        return Ok(());
    }
    for element in elements {
        let attribute = element
            .attributes
            .iter()
            .find(|attribute| attribute.namespace.is_empty() && attribute.name == name);
        match (attribute, after) {
            (Some(attribute), Some(value)) => {
                let quote = source
                    .get(attribute.value_start.wrapping_sub(1))
                    .copied()
                    .ok_or_else(|| invalid("OLE source attribute quote is missing"))?;
                let mut replacement = Vec::new();
                escape_for_quote(&mut replacement, value, quote);
                add_edit(
                    edits,
                    attribute.value_start,
                    attribute.value_end,
                    replacement,
                )?;
            },
            (Some(attribute), None) => {
                add_edit(edits, attribute.remove_start, attribute.end, Vec::new())?;
            },
            (None, Some(value)) => {
                let position = element
                    .start_end
                    .checked_sub(1)
                    .and_then(|position| {
                        if element.empty && source.get(position.wrapping_sub(1)) == Some(&b'/') {
                            position.checked_sub(1)
                        } else {
                            Some(position)
                        }
                    })
                    .ok_or_else(|| invalid("OLE source attribute insertion offset overflow"))?;
                let mut replacement = Vec::new();
                replacement.push(b' ');
                replacement.extend_from_slice(name.as_bytes());
                replacement.extend_from_slice(b"=\"");
                escape_for_quote(&mut replacement, value, b'"');
                replacement.push(b'"');
                add_edit(edits, position, position, replacement)?;
            },
            (None, None) => {},
        }
    }
    Ok(())
}

fn trimmed_text_span(source: &[u8], start: usize, end: usize) -> Result<(usize, usize)> {
    if end < start || end > source.len() {
        return Err(invalid("OLE source text span is outside worksheet XML"));
    }
    let mut start = start;
    let mut end = end;
    while start < end && source[start].is_ascii_whitespace() {
        start += 1;
    }
    while end > start && source[end - 1].is_ascii_whitespace() {
        end -= 1;
    }
    if start == end {
        return Err(invalid("OLE source coordinate has no text"));
    }
    Ok((start, end))
}

#[derive(Debug)]
struct Edit {
    start: usize,
    end: usize,
    replacement: Vec<u8>,
}

fn add_edit(edits: &mut Vec<Edit>, start: usize, end: usize, replacement: Vec<u8>) -> Result<()> {
    if start > end {
        return Err(invalid("OLE source edit range is inverted"));
    }
    edits.push(Edit {
        start,
        end,
        replacement,
    });
    Ok(())
}

fn apply_edits(source: &[u8], mut edits: Vec<Edit>) -> Result<Vec<u8>> {
    edits.sort_by_key(|edit| Reverse(edit.start));
    let mut next_start = source.len();
    for edit in &edits {
        if edit.start > edit.end || edit.end > next_start || edit.end > source.len() {
            return Err(invalid("overlapping OLE source edits"));
        }
        next_start = edit.start;
    }
    let mut delta = 0isize;
    for edit in &edits {
        let replacement = isize::try_from(edit.replacement.len())
            .map_err(|_source| limit("updated XML bytes"))?;
        let removed =
            isize::try_from(edit.end - edit.start).map_err(|_source| limit("updated XML bytes"))?;
        delta = delta
            .checked_add(replacement - removed)
            .ok_or_else(|| limit("updated XML bytes"))?;
    }
    let size = if delta >= 0 {
        source
            .len()
            .checked_add(usize::try_from(delta).map_err(|_source| limit("updated XML bytes"))?)
    } else {
        source
            .len()
            .checked_sub(usize::try_from(-delta).map_err(|_source| limit("updated XML bytes"))?)
    }
    .ok_or_else(|| limit("updated XML bytes"))?;
    if size > MAX_XML_BYTES {
        return Err(limit("updated XML bytes"));
    }

    let mut output = Vec::with_capacity(size);
    let mut cursor = 0usize;
    for edit in edits.into_iter().rev() {
        output.extend_from_slice(&source[cursor..edit.start]);
        output.extend_from_slice(&edit.replacement);
        cursor = edit.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn escape_for_quote(output: &mut Vec<u8>, value: &str, quote: u8) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' if quote == b'"' => output.extend_from_slice(b"&quot;"),
            '\'' if quote == b'\'' => output.extend_from_slice(b"&apos;"),
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
    let local_name = element.local_name();
    let name = std::str::from_utf8(local_name.as_ref())
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
            .map_err(|_source| invalid("worksheet XML offset overflow"))?;
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
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
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
fn no_attributes(_node: &Node, _allowed: &[(&str, &str)]) -> Result<()> {
    // Unknown attributes are retained by source-bound transactions. The
    // semantic loader intentionally ignores them so extension producers can
    // coexist with the typed core without forcing a lossy rewrite.
    Ok(())
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
        .map_err(|_source| invalid(format!("invalid object marker {name}")))
}
fn signed_coordinate(node: &Node, name: &str) -> Result<i64> {
    node.text
        .trim()
        .parse()
        .map_err(|_source| invalid(format!("invalid object marker {name}")))
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
