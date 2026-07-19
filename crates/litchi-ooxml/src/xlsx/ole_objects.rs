//! Typed SpreadsheetML OLE-object anchors with inert package payloads.

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 500_000;
const MAX_DEPTH: usize = 256;
const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
const MAX_OBJECTS: usize = 1024;
const MAX_PAYLOAD_BYTES: usize = 64 * 1024 * 1024;
const MAX_TOTAL_PAYLOAD_BYTES: usize = 256 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OleObjectConformance {
    Transitional,
    Strict,
}

impl OleObjectConformance {
    fn sml(self) -> &'static str {
        match self {
            Self::Transitional => SML,
            Self::Strict => STRICT_SML,
        }
    }
    fn rel(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }
    fn xdr(self) -> &'static str {
        match self {
            Self::Transitional => XDR,
            Self::Strict => STRICT_XDR,
        }
    }
    fn ole_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::OLE_OBJECT,
            Self::Strict => rt::STRICT_OLE_OBJECT,
        }
    }
    fn package_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::PACKAGE,
            Self::Strict => rt::STRICT_PACKAGE,
        }
    }
    fn image_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::IMAGE,
            Self::Strict => rt::STRICT_IMAGE,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OleObjectAspect {
    Content,
    Icon,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OleObjectUpdate {
    Always,
    OnCall,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OleObjectRelationshipKind {
    OleObject,
    Package,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleObjectResource {
    pub part_name: String,
    pub content_type: String,
    /// Stored and returned without format sniffing, parsing, or activation.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OleObjectTarget {
    Internal(OleObjectResource),
    /// An inert OPC external target. It is never fetched or activated.
    External(String),
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct OleObjectMarker {
    pub column: u32,
    pub column_offset: i64,
    pub row: u32,
    pub row_offset: i64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleObjectAnchor {
    pub move_with_cells: Option<bool>,
    pub size_with_cells: Option<bool>,
    pub from: OleObjectMarker,
    pub to: OleObjectMarker,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OleObjectProperties {
    pub preview_relationship_id: String,
    pub preview: Option<OleObjectResource>,
    pub default_size: Option<bool>,
    pub print: Option<bool>,
    pub disabled: Option<bool>,
    pub ui_object: Option<bool>,
    pub auto_fill: Option<bool>,
    pub auto_line: Option<bool>,
    pub auto_pict: Option<bool>,
    pub dde: Option<bool>,
    pub macro_name: Option<String>,
    pub alt_text: Option<String>,
    pub anchor: OleObjectAnchor,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetOleObject {
    pub program_id: Option<String>,
    pub data_or_view_aspect: Option<OleObjectAspect>,
    pub link: Option<String>,
    pub update: Option<OleObjectUpdate>,
    pub auto_load: Option<bool>,
    pub shape_id: u32,
    pub relationship_id: String,
    pub relationship_kind: OleObjectRelationshipKind,
    /// Filled by package loading and required by package storage.
    pub target: Option<OleObjectTarget>,
    pub properties: Option<OleObjectProperties>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorksheetOleObjects {
    pub objects: Vec<WorksheetOleObject>,
}

#[derive(Clone)]
struct Attribute {
    namespace: String,
    name: String,
    value: String,
}
#[derive(Clone)]
struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}

/// Parses the optional `oleObjects` collection from a complete worksheet part.
pub fn parse_worksheet_ole_objects(xml: &[u8]) -> Result<Option<WorksheetOleObjects>> {
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
    let value = WorksheetOleObjects { objects };
    validate_value(&value, false)?;
    Ok(Some(value))
}

fn parse_object(node: &Node, conformance: OleObjectConformance) -> Result<WorksheetOleObject> {
    whitespace(node)?;
    if node.children.len() > 1 {
        return Err(invalid("oleObject has multiple child elements"));
    }
    let program_id = optional(node, "", "progId").map(str::to_owned);
    let data_or_view_aspect = optional(node, "", "dvAspect")
        .map(parse_aspect)
        .transpose()?;
    let link = optional(node, "", "link").map(str::to_owned);
    let update = optional(node, "", "oleUpdate")
        .map(parse_update)
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
    Ok(WorksheetOleObject {
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
pub fn write_worksheet_ole_objects(
    value: &WorksheetOleObjects,
    conformance: OleObjectConformance,
) -> Result<Vec<u8>> {
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
            attr(
                &mut output,
                "dvAspect",
                match value {
                    OleObjectAspect::Content => "DVASPECT_CONTENT",
                    OleObjectAspect::Icon => "DVASPECT_ICON",
                },
            );
        }
        if let Some(value) = &object.link {
            attr(&mut output, "link", value);
        }
        if let Some(value) = object.update {
            attr(
                &mut output,
                "oleUpdate",
                match value {
                    OleObjectUpdate::Always => "OLEUPDATE_ALWAYS",
                    OleObjectUpdate::OnCall => "OLEUPDATE_ONCALL",
                },
            );
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
pub fn load_worksheet_ole_objects(
    package: &OpcPackage,
    worksheet_name: &PackURI,
) -> Result<Option<WorksheetOleObjects>> {
    if package
        .rels()
        .iter()
        .any(|relationship| embedded_kind(relationship.reltype()).is_some())
    {
        return Err(invalid(
            "package root cannot source embedded-object relationships",
        ));
    }
    let worksheet = package.get_part(worksheet_name)?;
    require_worksheet(worksheet)?;
    let Some(mut value) = parse_worksheet_ole_objects(worksheet.blob())? else {
        if worksheet
            .rels()
            .iter()
            .any(|relationship| embedded_kind(relationship.reltype()).is_some())
        {
            return Err(invalid(
                "worksheet has embedded-object relationships without oleObjects markup",
            ));
        }
        return Ok(None);
    };
    let mut referenced = HashSet::new();
    let mut targets = HashSet::new();
    let mut total = 0usize;
    for object in &mut value.objects {
        if !referenced.insert(object.relationship_id.clone()) {
            return Err(invalid(format!(
                "duplicate object relationship reference '{}'",
                object.relationship_id
            )));
        }
        let relationship = worksheet
            .rels()
            .get(&object.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "missing object relationship '{}'",
                    object.relationship_id
                ))
            })?;
        let kind = embedded_kind(relationship.reltype()).ok_or_else(|| {
            invalid(format!(
                "relationship '{}' is not an embedded-object relationship",
                object.relationship_id
            ))
        })?;
        object.relationship_kind = kind;
        object.target = Some(if relationship.is_external() {
            if object.link.is_none() {
                return Err(invalid("external OLE relationship requires a link moniker"));
            }
            OleObjectTarget::External(relationship.target_ref().to_owned())
        } else {
            let target = relationship.target_partname()?;
            if !targets.insert(target.to_string()) {
                return Err(invalid(format!("multiple OLE objects target '{target}'")));
            }
            if !target.as_str().starts_with("/xl/embeddings/") {
                return Err(invalid(format!(
                    "embedded object '{target}' is outside /xl/embeddings"
                )));
            }
            let part = package.get_part(&target)?;
            validate_payload(part, kind)?;
            add_payload(&mut total, part.blob().len())?;
            OleObjectTarget::Internal(OleObjectResource {
                part_name: target.to_string(),
                content_type: part.content_type().to_owned(),
                data: part.blob().to_vec(),
            })
        });
        if let Some(properties) = object.properties.as_mut() {
            let relationship = worksheet
                .rels()
                .get(&properties.preview_relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "missing object preview relationship '{}'",
                        properties.preview_relationship_id
                    ))
                })?;
            if !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
                || relationship.is_external()
            {
                return Err(invalid(
                    "object preview relationship must be an internal image",
                ));
            }
            let target = relationship.target_partname()?;
            if !target.as_str().starts_with("/xl/media/") {
                return Err(invalid(format!(
                    "object preview '{target}' is outside /xl/media"
                )));
            }
            let part = package.get_part(&target)?;
            if !is_image_content_type(part.content_type()) || !part.rels().is_empty() {
                return Err(invalid(format!(
                    "object preview '{target}' is not a relationship-free image"
                )));
            }
            add_payload(&mut total, part.blob().len())?;
            properties.preview = Some(OleObjectResource {
                part_name: target.to_string(),
                content_type: part.content_type().to_owned(),
                data: part.blob().to_vec(),
            });
        }
    }
    for relationship in worksheet
        .rels()
        .iter()
        .filter(|relationship| embedded_kind(relationship.reltype()).is_some())
    {
        if !referenced.contains(relationship.r_id()) {
            return Err(invalid(format!(
                "unreferenced embedded-object relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    Ok(Some(value))
}

/// Adds a new OLE collection and its inert relationships to one worksheet.
pub fn store_worksheet_ole_objects(
    package: &mut OpcPackage,
    worksheet_name: &PackURI,
    value: &WorksheetOleObjects,
    conformance: OleObjectConformance,
) -> Result<()> {
    validate_value(value, true)?;
    if load_worksheet_ole_objects(package, worksheet_name)?.is_some() {
        return Err(invalid("worksheet already contains OLE objects"));
    }
    let worksheet = package.get_part(worksheet_name)?;
    let root = parse_document(worksheet.blob())?;
    if crate_conformance(&root)? != conformance {
        return Err(invalid(
            "requested conformance does not match worksheet namespace",
        ));
    }
    let fragment = write_worksheet_ole_objects(value, conformance)?;
    let updated = insert_collection(worksheet.blob(), &fragment, conformance)?;
    let mut relationships: HashMap<String, (String, String, bool)> = HashMap::new();
    let mut parts: HashMap<String, OleObjectResource> = HashMap::new();
    for object in &value.objects {
        let target = object
            .target
            .as_ref()
            .ok_or_else(|| invalid("OLE target is required for package storage"))?;
        let relationship_type = match object.relationship_kind {
            OleObjectRelationshipKind::OleObject => conformance.ole_rel(),
            OleObjectRelationshipKind::Package => conformance.package_rel(),
        };
        match target {
            OleObjectTarget::External(target) => add_relationship_plan(
                &mut relationships,
                &object.relationship_id,
                relationship_type,
                target,
                true,
            )?,
            OleObjectTarget::Internal(resource) => {
                let uri = resource_uri(resource, "/xl/embeddings/")?;
                add_part_plan(package, &mut parts, resource)?;
                add_relationship_plan(
                    &mut relationships,
                    &object.relationship_id,
                    relationship_type,
                    &uri.relative_ref(worksheet_name.base_uri()),
                    false,
                )?;
            },
        }
        if let Some(properties) = &object.properties {
            let preview = properties.preview.as_ref().ok_or_else(|| {
                invalid("object preview resource is required for package storage")
            })?;
            let uri = resource_uri(preview, "/xl/media/")?;
            add_part_plan(package, &mut parts, preview)?;
            add_relationship_plan(
                &mut relationships,
                &properties.preview_relationship_id,
                conformance.image_rel(),
                &uri.relative_ref(worksheet_name.base_uri()),
                false,
            )?;
        }
    }
    for id in relationships.keys() {
        if package.get_part(worksheet_name)?.rels().get(id).is_some() {
            return Err(invalid(format!(
                "worksheet relationship ID '{id}' already exists"
            )));
        }
    }
    package.get_part_mut(worksheet_name)?.set_blob(updated);
    for resource in parts.into_values() {
        let uri = PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?;
        package.add_part(Box::new(BlobPart::new(
            uri,
            resource.content_type,
            resource.data,
        )));
    }
    for (id, (relationship_type, target, external)) in relationships {
        package
            .get_part_mut(worksheet_name)?
            .rels_mut()
            .add_relationship(relationship_type, target, id, external);
    }
    Ok(())
}

fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("input XML bytes"));
    }
    let mut caps = MceCapabilities::ooxml_baseline();
    caps.understand_namespace(X14);
    let limits = MceLimits {
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
fn crate_conformance(root: &Node) -> Result<OleObjectConformance> {
    if root.name != "worksheet" {
        return Err(invalid("expected worksheet root"));
    }
    match root.namespace.as_str() {
        SML => Ok(OleObjectConformance::Transitional),
        STRICT_SML => Ok(OleObjectConformance::Strict),
        _ => Err(invalid("unsupported worksheet namespace")),
    }
}

fn insert_collection(
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
                let core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value.as_ref() == conformance.sml().as_bytes());
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
                let core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value.as_ref() == conformance.sml().as_bytes());
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

fn validate_value(value: &WorksheetOleObjects, require_targets: bool) -> Result<()> {
    if value.objects.len() > MAX_OBJECTS {
        return Err(limit("object count"));
    }
    let mut shapes = HashSet::new();
    let mut ids = HashSet::new();
    let mut total = 0usize;
    for object in &value.objects {
        if !(1..=67_098_623).contains(&object.shape_id) {
            return Err(invalid("OLE shapeId is outside Office's supported range"));
        }
        if !shapes.insert(object.shape_id) {
            return Err(invalid(format!(
                "duplicate OLE shapeId {}",
                object.shape_id
            )));
        }
        validate_id(&object.relationship_id)?;
        if !ids.insert(object.relationship_id.clone()) {
            return Err(invalid(format!(
                "duplicate OLE relationship ID '{}'",
                object.relationship_id
            )));
        }
        if let Some(value) = &object.program_id {
            bounded(value)?;
            if value.len() >= 39
                || value
                    .bytes()
                    .next()
                    .is_some_and(|byte| byte.is_ascii_digit())
                || !value
                    .bytes()
                    .all(|byte| byte.is_ascii_alphanumeric() || byte == b'.')
            {
                return Err(invalid(format!("invalid Office ProgID '{value}'")));
            }
        }
        if let Some(value) = &object.link {
            bounded(value)?;
            if value.len() > 8192 {
                return Err(invalid("OLE link moniker exceeds Office's limit"));
            }
        }
        if require_targets && object.target.is_none() {
            return Err(invalid("OLE target is required for package storage"));
        }
        if let Some(target) = &object.target {
            match target {
                OleObjectTarget::External(value) => {
                    bounded(value)?;
                    if object.link.is_none() {
                        return Err(invalid("external OLE target requires a link moniker"));
                    }
                },
                OleObjectTarget::Internal(resource) => {
                    validate_resource(resource, "/xl/embeddings/")?;
                    if object.relationship_kind == OleObjectRelationshipKind::OleObject
                        && resource.content_type != ct::OFC_OLE_OBJECT
                    {
                        return Err(invalid(
                            "OLE relationship requires the OLE Object content type",
                        ));
                    }
                    add_payload(&mut total, resource.data.len())?;
                },
            }
        }
        if let Some(properties) = &object.properties {
            validate_id(&properties.preview_relationship_id)?;
            if properties.preview_relationship_id == object.relationship_id {
                return Err(invalid("payload and preview relationship IDs must differ"));
            }
            if let Some(value) = &properties.macro_name {
                bounded(value)?;
            }
            if let Some(value) = &properties.alt_text {
                bounded(value)?;
            }
            if require_targets && properties.preview.is_none() {
                return Err(invalid("object preview is required for package storage"));
            }
            if let Some(preview) = &properties.preview {
                validate_resource(preview, "/xl/media/")?;
                if !is_image_content_type(&preview.content_type) {
                    return Err(invalid("object preview has a non-image content type"));
                }
                add_payload(&mut total, preview.data.len())?;
            }
        }
    }
    Ok(())
}

fn validate_payload(part: &dyn Part, kind: OleObjectRelationshipKind) -> Result<()> {
    if kind == OleObjectRelationshipKind::OleObject && part.content_type() != ct::OFC_OLE_OBJECT {
        return Err(invalid(format!(
            "OLE payload '{}' has invalid content type '{}'",
            part.partname(),
            part.content_type()
        )));
    }
    for relationship in part.rels().iter() {
        if !matches!(relationship.reltype(), rt::HYPERLINK | rt::STRICT_HYPERLINK) {
            return Err(invalid(format!(
                "embedded payload '{}' has forbidden outbound relationship",
                part.partname()
            )));
        }
    }
    Ok(())
}
fn embedded_kind(value: &str) -> Option<OleObjectRelationshipKind> {
    match value {
        rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT => Some(OleObjectRelationshipKind::OleObject),
        rt::PACKAGE | rt::STRICT_PACKAGE => Some(OleObjectRelationshipKind::Package),
        _ => None,
    }
}
fn require_worksheet(part: &dyn Part) -> Result<()> {
    if part.content_type() == ct::SML_WORKSHEET {
        Ok(())
    } else {
        Err(invalid(format!(
            "part '{}' is not a worksheet",
            part.partname()
        )))
    }
}
fn resource_uri(resource: &OleObjectResource, prefix: &str) -> Result<PackURI> {
    validate_resource(resource, prefix)?;
    PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)
}
fn validate_resource(resource: &OleObjectResource, prefix: &str) -> Result<()> {
    let uri = PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?;
    if !uri.as_str().starts_with(prefix) {
        return Err(invalid(format!("resource '{uri}' is outside {prefix}")));
    }
    if resource.content_type.is_empty()
        || resource
            .content_type
            .bytes()
            .any(|byte| byte.is_ascii_whitespace())
    {
        return Err(invalid("invalid embedded resource content type"));
    }
    if resource.data.len() > MAX_PAYLOAD_BYTES {
        return Err(limit("individual payload bytes"));
    }
    Ok(())
}
fn add_part_plan(
    package: &OpcPackage,
    parts: &mut HashMap<String, OleObjectResource>,
    resource: &OleObjectResource,
) -> Result<()> {
    if package
        .iter_parts()
        .any(|part| part.partname().as_str() == resource.part_name)
    {
        return Err(invalid(format!(
            "resource part '{}' already exists",
            resource.part_name
        )));
    }
    if let Some(existing) = parts.get(&resource.part_name) {
        if existing != resource {
            return Err(invalid(format!(
                "conflicting resource part '{}'",
                resource.part_name
            )));
        }
    } else {
        parts.insert(resource.part_name.clone(), resource.clone());
    }
    Ok(())
}
fn add_relationship_plan(
    plans: &mut HashMap<String, (String, String, bool)>,
    id: &str,
    kind: &str,
    target: &str,
    external: bool,
) -> Result<()> {
    validate_id(id)?;
    let plan = (kind.to_owned(), target.to_owned(), external);
    if let Some(existing) = plans.get(id) {
        if existing != &plan {
            return Err(invalid(format!("conflicting relationship ID '{id}'")));
        }
    } else {
        plans.insert(id.to_owned(), plan);
    }
    Ok(())
}
fn add_payload(total: &mut usize, size: usize) -> Result<()> {
    if size > MAX_PAYLOAD_BYTES {
        return Err(limit("individual payload bytes"));
    }
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("total payload bytes"))?;
    if *total > MAX_TOTAL_PAYLOAD_BYTES {
        Err(limit("total payload bytes"))
    } else {
        Ok(())
    }
}
fn is_image_content_type(value: &str) -> bool {
    value.starts_with("image/") || matches!(value, "application/x-emf" | "application/x-wmf")
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
fn parse_aspect(value: &str) -> Result<OleObjectAspect> {
    match value {
        "DVASPECT_CONTENT" => Ok(OleObjectAspect::Content),
        "DVASPECT_ICON" => Ok(OleObjectAspect::Icon),
        _ => Err(invalid(format!("invalid OLE data/view aspect '{value}'"))),
    }
}
fn parse_update(value: &str) -> Result<OleObjectUpdate> {
    match value {
        "OLEUPDATE_ALWAYS" => Ok(OleObjectUpdate::Always),
        "OLEUPDATE_ONCALL" => Ok(OleObjectUpdate::OnCall),
        _ => Err(invalid(format!("invalid OLE update mode '{value}'"))),
    }
}
fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID cannot be empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!("invalid relationship ID '{value}'")))
    } else {
        Ok(())
    }
}
fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("string bytes"))
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
        ResolveResult::Bound(Namespace(value)) => Ok(std::str::from_utf8(value.as_ref())
            .map_err(xml_error)?
            .to_owned()),
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
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(name: &str) -> OoxmlError {
    invalid(format!("worksheet OLE {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI: &[u8] =
        include_bytes!("../../../../3rdparty/poi/test-data/spreadsheet/bug64512_embed.xlsx");

    fn marker(row: u32) -> OleObjectMarker {
        OleObjectMarker {
            column: 1,
            column_offset: 0,
            row,
            row_offset: 0,
        }
    }
    fn value() -> WorksheetOleObjects {
        WorksheetOleObjects {
            objects: vec![WorksheetOleObject {
                program_id: Some("Package.2".into()),
                data_or_view_aspect: Some(OleObjectAspect::Icon),
                link: None,
                update: Some(OleObjectUpdate::OnCall),
                auto_load: Some(false),
                shape_id: 1025,
                relationship_id: "rIdOle".into(),
                relationship_kind: OleObjectRelationshipKind::OleObject,
                target: Some(OleObjectTarget::Internal(OleObjectResource {
                    part_name: "/xl/embeddings/oleObject1.bin".into(),
                    content_type: ct::OFC_OLE_OBJECT.into(),
                    data: vec![0xd0, 0xcf, 0x11, 0xe0],
                })),
                properties: Some(OleObjectProperties {
                    preview_relationship_id: "rIdPreview".into(),
                    preview: Some(OleObjectResource {
                        part_name: "/xl/media/image1.emf".into(),
                        content_type: "image/x-emf".into(),
                        data: vec![1, 2, 3],
                    }),
                    default_size: Some(false),
                    print: Some(true),
                    disabled: None,
                    ui_object: None,
                    auto_fill: Some(false),
                    auto_line: Some(false),
                    auto_pict: None,
                    dde: None,
                    macro_name: None,
                    alt_text: Some("Object preview".into()),
                    anchor: OleObjectAnchor {
                        move_with_cells: Some(true),
                        size_with_cells: Some(false),
                        from: marker(1),
                        to: marker(3),
                    },
                }),
            }],
        }
    }
    fn package(conformance: OleObjectConformance) -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            ct::SML_WORKSHEET.into(),
            format!(
                "<x:worksheet xmlns:x=\"{}\"><x:sheetData/><x:tableParts/></x:worksheet>",
                conformance.sml()
            )
            .into_bytes(),
        )));
        (package, uri)
    }

    #[test]
    fn strict_round_trip_covers_complete_typed_properties() {
        let expected = value();
        let fragment =
            write_worksheet_ole_objects(&expected, OleObjectConformance::Strict).unwrap();
        let xml = [
            format!("<x:worksheet xmlns:x=\"{STRICT_SML}\">").as_bytes(),
            fragment.as_slice(),
            b"</x:worksheet>",
        ]
        .concat();
        let parsed = parse_worksheet_ole_objects(&xml).unwrap().unwrap();
        assert_eq!(parsed.objects[0].program_id.as_deref(), Some("Package.2"));
        assert_eq!(
            parsed.objects[0].properties.as_ref().unwrap().anchor.to.row,
            3
        );
        assert!(parsed.objects[0].target.is_none());
    }

    #[test]
    fn loads_real_poi_mce_objects_without_opening_payloads() {
        let package = OpcPackage::from_bytes(POI).unwrap();
        let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        let objects = load_worksheet_ole_objects(&package, &uri).unwrap().unwrap();
        assert_eq!(objects.objects.len(), 2);
        assert_eq!(objects.objects[0].program_id.as_deref(), Some("Package"));
        assert_eq!(objects.objects[1].program_id.as_deref(), Some("Package2"));
        assert!(
            objects
                .objects
                .iter()
                .all(|object| matches!(object.target, Some(OleObjectTarget::Internal(_))))
        );
        assert!(objects.objects.iter().all(|object| {
            object
                .properties
                .as_ref()
                .unwrap()
                .preview
                .as_ref()
                .unwrap()
                .data
                .starts_with(b"\x01\x00\x00\x00")
        }));
    }

    #[test]
    fn strict_package_writer_round_trips_and_inserts_in_schema_order() {
        let (mut package, uri) = package(OleObjectConformance::Strict);
        let expected = value();
        store_worksheet_ole_objects(&mut package, &uri, &expected, OleObjectConformance::Strict)
            .unwrap();
        assert_eq!(
            load_worksheet_ole_objects(&package, &uri).unwrap().unwrap(),
            expected
        );
        let xml = package.get_part(&uri).unwrap().blob();
        assert!(
            memchr::memmem::find(xml, b"<x:oleObjects").unwrap()
                < memchr::memmem::find(xml, b"<x:tableParts").unwrap()
        );
    }

    #[test]
    fn accepts_inert_external_package_target() {
        let (mut package, uri) = package(OleObjectConformance::Transitional);
        let mut expected = value();
        let object = &mut expected.objects[0];
        object.relationship_kind = OleObjectRelationshipKind::Package;
        object.target = Some(OleObjectTarget::External(
            "https://example.invalid/object.xlsx".into(),
        ));
        object.link = Some("'https://example.invalid/object.xlsx'!A1".into());
        object.properties = None;
        store_worksheet_ole_objects(
            &mut package,
            &uri,
            &expected,
            OleObjectConformance::Transitional,
        )
        .unwrap();
        assert_eq!(
            load_worksheet_ole_objects(&package, &uri).unwrap().unwrap(),
            expected
        );
    }

    #[test]
    fn rejects_malformed_markup_caps_and_graphs() {
        for xml in [
            format!(
                r#"<worksheet xmlns="{SML}"><oleObjects><oleObject shapeId="0"/></oleObjects></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><oleObjects><oleObject shapeId="1" r:id="rId1"><objectPr r:id="rId2"><anchor><to/></anchor></objectPr></oleObject></oleObjects></worksheet>"#
            ),
            format!(r#"<!DOCTYPE x><worksheet xmlns="{SML}"/>"#),
        ] {
            assert!(
                parse_worksheet_ole_objects(xml.as_bytes()).is_err(),
                "{xml}"
            );
        }
        assert!(parse_worksheet_ole_objects(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
        let (mut missing, uri) = package(OleObjectConformance::Transitional);
        let fragment =
            write_worksheet_ole_objects(&value(), OleObjectConformance::Transitional).unwrap();
        missing.get_part_mut(&uri).unwrap().set_blob(
            [
                format!("<x:worksheet xmlns:x=\"{SML}\">").as_bytes(),
                fragment.as_slice(),
                b"</x:worksheet>",
            ]
            .concat(),
        );
        assert!(load_worksheet_ole_objects(&missing, &uri).is_err());
        let (mut unreferenced, uri) = package(OleObjectConformance::Transitional);
        unreferenced
            .get_part_mut(&uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::OLE_OBJECT.into(),
                "../embeddings/x.bin".into(),
                "rIdX".into(),
                false,
            );
        assert!(load_worksheet_ole_objects(&unreferenced, &uri).is_err());
    }
}
