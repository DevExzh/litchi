//! MS-XLSX Spreadsheet Data Model package support.
//!
//! The MS-XLDM payload is intentionally opaque. It can contain encrypted
//! connection credentials and a complex virtual file system, so this module
//! only inventories and copies bounded bytes. It never decrypts, decompresses,
//! evaluates expressions, or accesses external resources.

use super::xldm::inspect_xldm;
use crate::error::{OoxmlError, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashMap, HashSet};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const X15: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const CONNECTIONS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
const STRICT_CONNECTIONS_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/connections";
const CONNECTIONS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";

pub const DATA_MODEL_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.model+data";
pub const DATA_MODEL_EXTENSION_URI: &str = "{FCE2AD5D-F65C-4FA6-A056-5C36A1767C68}";
pub const DATA_MODEL_PART_NAME: &str = "/xl/model/item.data";

const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_REWRITE_BYTES: usize = 32 * 1024 * 1024;
const MAX_EXTENSION_BYTES: usize = 4 * 1024 * 1024;
const MAX_PAYLOAD_BYTES: usize = 512 * 1024 * 1024;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_TOTAL_STRING_BYTES: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 200_000;
const MAX_DEPTH: usize = 128;
const MAX_TABLES: usize = 65_536;
const MAX_RELATIONSHIPS: usize = 65_536;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataModelOpaqueXml {
    /// A self-contained `x15:extLst` subtree retained without interpretation.
    pub xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataModelTable {
    pub id: String,
    pub name: String,
    pub connection: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataModelRelationship {
    pub from_table: String,
    pub from_column: String,
    pub to_table: String,
    pub to_column: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataModelDefinition {
    /// Minimum application version. MS-XLSX defines 5 as the default and floor.
    pub min_version_load: u8,
    pub tables: Vec<DataModelTable>,
    pub relationships: Vec<DataModelRelationship>,
    pub extension_list: Option<DataModelOpaqueXml>,
}

impl Default for DataModelDefinition {
    fn default() -> Self {
        Self {
            min_version_load: 5,
            tables: Vec::new(),
            relationships: Vec::new(),
            extension_list: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataModelPayload {
    pub part_name: String,
    /// Opaque MS-XLDM bytes. No inner-file or credential processing is performed.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataModel {
    pub definition: DataModelDefinition,
    pub payload: DataModelPayload,
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

/// Parse an inline `x15:dataModel` descriptor.
pub fn parse_data_model(xml: &[u8]) -> Result<DataModelDefinition> {
    let root = parse_document(xml)?;
    parse_data_model_node(&root)
}

/// Deterministically serialize an inline `x15:dataModel` descriptor.
pub fn write_data_model(value: &DataModelDefinition) -> Result<Vec<u8>> {
    validate_definition(value, false)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<x15:dataModel xmlns:x15=\"");
    escape(&mut output, X15);
    output.push(b'\"');
    if value.min_version_load != 5 {
        attr(
            &mut output,
            "minVersionLoad",
            &value.min_version_load.to_string(),
        );
    }
    if value.tables.is_empty() && value.relationships.is_empty() && value.extension_list.is_none() {
        output.extend_from_slice(b"/>");
        return Ok(output);
    }
    output.push(b'>');
    if !value.tables.is_empty() {
        output.extend_from_slice(b"<x15:modelTables>");
        for table in &value.tables {
            output.extend_from_slice(b"<x15:modelTable");
            attr(&mut output, "id", &table.id);
            attr(&mut output, "name", &table.name);
            attr(&mut output, "connection", &table.connection);
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</x15:modelTables>");
    }
    if !value.relationships.is_empty() {
        output.extend_from_slice(b"<x15:modelRelationships>");
        for relationship in &value.relationships {
            output.extend_from_slice(b"<x15:modelRelationship");
            attr(&mut output, "fromTable", &relationship.from_table);
            attr(&mut output, "fromColumn", &relationship.from_column);
            attr(&mut output, "toTable", &relationship.to_table);
            attr(&mut output, "toColumn", &relationship.to_column);
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</x15:modelRelationships>");
    }
    if let Some(extension) = &value.extension_list {
        output.extend_from_slice(&extension.xml);
    }
    output.extend_from_slice(b"</x15:dataModel>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized descriptor bytes"));
    }
    Ok(output)
}

/// Load the singleton workbook Data Model and retain its MS-XLDM payload inertly.
pub fn load_data_model(package: &OpcPackage, workbook_name: &PackURI) -> Result<Option<DataModel>> {
    let workbook = package.get_part(workbook_name)?;
    let workbook_root = parse_document(workbook.blob())?;
    let (_, definition) = workbook_definition(&workbook_root)?;
    let mut parts = package
        .iter_parts()
        .filter(|part| part.content_type() == DATA_MODEL_CONTENT_TYPE);
    let part = parts.next();
    if parts.next().is_some() {
        return Err(invalid("package contains multiple Data Model parts"));
    }
    let (definition, part) = match (definition, part) {
        (Some(definition), Some(part)) => (definition, part),
        (Some(_), None) => {
            return Err(invalid(
                "workbook dataModel extension has no Data Model part",
            ));
        },
        (None, Some(_)) => {
            return Err(invalid(
                "Data Model part has no workbook dataModel extension",
            ));
        },
        (None, None) => return Ok(None),
    };
    if part.partname().as_str() != DATA_MODEL_PART_NAME {
        return Err(invalid(format!(
            "Data Model part '{}' must be '{DATA_MODEL_PART_NAME}'",
            part.partname()
        )));
    }
    if part.blob().is_empty() {
        return Err(invalid("Data Model payload cannot be empty"));
    }
    if part.blob().len() > MAX_PAYLOAD_BYTES {
        return Err(limit("payload bytes"));
    }
    inspect_xldm(part.blob())?;
    if !part.rels().is_empty() {
        return Err(invalid(
            "Data Model part has forbidden outbound relationships",
        ));
    }
    reject_inbound_relationships(package, part.partname())?;
    validate_connections(package, workbook_name, &definition)?;
    Ok(Some(DataModel {
        definition,
        payload: DataModelPayload {
            part_name: part.partname().to_string(),
            data: part.blob().to_vec(),
        },
    }))
}

/// Store a singleton Data Model after validating the complete mutation plan.
pub fn store_data_model(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &DataModel,
) -> Result<()> {
    validate_definition(&value.definition, false)?;
    validate_payload(&value.payload)?;
    if load_data_model(package, workbook_name)?.is_some() {
        return Err(invalid("workbook already contains a Data Model"));
    }
    if package
        .iter_parts()
        .any(|part| part.partname().as_str() == DATA_MODEL_PART_NAME)
    {
        return Err(invalid(format!(
            "part '{DATA_MODEL_PART_NAME}' already exists"
        )));
    }
    validate_connections(package, workbook_name, &value.definition)?;
    let workbook = package.get_part(workbook_name)?;
    let root = parse_document(workbook.blob())?;
    let (core, existing) = workbook_definition(&root)?;
    if existing.is_some() {
        return Err(invalid("workbook already has a dataModel extension"));
    }
    let descriptor = write_data_model(&value.definition)?;
    let mut fragment = Vec::new();
    fragment.extend_from_slice(b"<x:ext xmlns:x=\"");
    escape(&mut fragment, core);
    fragment.extend_from_slice(b"\" uri=\"");
    escape(&mut fragment, DATA_MODEL_EXTENSION_URI);
    fragment.extend_from_slice(b"\">");
    fragment.extend_from_slice(&descriptor);
    fragment.extend_from_slice(b"</x:ext>");
    let updated = insert_extension(workbook.blob(), core, &fragment)?;
    let uri = PackURI::new(&value.payload.part_name).map_err(OoxmlError::InvalidUri)?;
    package.add_part(Box::new(BlobPart::new(
        uri,
        DATA_MODEL_CONTENT_TYPE.into(),
        value.payload.data.clone(),
    )));
    package.get_part_mut(workbook_name)?.set_blob(updated);
    Ok(())
}

fn parse_data_model_node(root: &Node) -> Result<DataModelDefinition> {
    require(root, X15, "dataModel")?;
    no_attributes(root, &[("", "minVersionLoad")])?;
    whitespace(root)?;
    let min_version_load = optional(root, "", "minVersionLoad")
        .map(|value| {
            value
                .parse::<u8>()
                .map_err(|_| invalid("minVersionLoad must be an unsigned byte"))
        })
        .transpose()?
        .unwrap_or(5);
    let mut tables = Vec::new();
    let mut relationships = Vec::new();
    let mut extension_list = None;
    let mut stage = 0u8;
    for child in &root.children {
        match child.name.as_str() {
            "modelTables" if child.namespace == X15 && stage == 0 => {
                stage = 1;
                tables = parse_tables(child)?;
            },
            "modelRelationships" if child.namespace == X15 && stage <= 1 => {
                stage = 2;
                relationships = parse_relationships(child)?;
            },
            "extLst" if child.namespace == X15 && stage <= 2 => {
                stage = 3;
                let xml = serialize_node(child)?;
                if xml.len() > MAX_EXTENSION_BYTES {
                    return Err(limit("extension bytes"));
                }
                extension_list = Some(DataModelOpaqueXml { xml });
            },
            _ => return Err(invalid("unexpected or out-of-order dataModel child")),
        }
    }
    let value = DataModelDefinition {
        min_version_load,
        tables,
        relationships,
        extension_list,
    };
    validate_definition(&value, true)?;
    Ok(value)
}

fn parse_tables(node: &Node) -> Result<Vec<DataModelTable>> {
    no_attributes(node, &[])?;
    whitespace(node)?;
    if node.children.is_empty() {
        return Err(invalid("modelTables must contain at least one modelTable"));
    }
    if node.children.len() > MAX_TABLES {
        return Err(limit("table count"));
    }
    node.children
        .iter()
        .map(|child| {
            require(child, X15, "modelTable")?;
            no_attributes(child, &[("", "id"), ("", "name"), ("", "connection")])?;
            leaf(child)?;
            Ok(DataModelTable {
                id: required(child, "", "id")?.to_owned(),
                name: required(child, "", "name")?.to_owned(),
                connection: required(child, "", "connection")?.to_owned(),
            })
        })
        .collect()
}

fn parse_relationships(node: &Node) -> Result<Vec<DataModelRelationship>> {
    no_attributes(node, &[])?;
    whitespace(node)?;
    if node.children.is_empty() {
        return Err(invalid(
            "modelRelationships must contain at least one modelRelationship",
        ));
    }
    if node.children.len() > MAX_RELATIONSHIPS {
        return Err(limit("relationship count"));
    }
    node.children
        .iter()
        .map(|child| {
            require(child, X15, "modelRelationship")?;
            no_attributes(
                child,
                &[
                    ("", "fromTable"),
                    ("", "fromColumn"),
                    ("", "toTable"),
                    ("", "toColumn"),
                ],
            )?;
            leaf(child)?;
            Ok(DataModelRelationship {
                from_table: required(child, "", "fromTable")?.to_owned(),
                from_column: required(child, "", "fromColumn")?.to_owned(),
                to_table: required(child, "", "toTable")?.to_owned(),
                to_column: required(child, "", "toColumn")?.to_owned(),
            })
        })
        .collect()
}

fn validate_definition(value: &DataModelDefinition, extension_already_parsed: bool) -> Result<()> {
    if value.min_version_load < 5 {
        return Err(invalid("minVersionLoad must be at least 5"));
    }
    if value.tables.len() > MAX_TABLES {
        return Err(limit("table count"));
    }
    if value.relationships.len() > MAX_RELATIONSHIPS {
        return Err(limit("relationship count"));
    }
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    for table in &value.tables {
        for (field, label) in [
            (&table.id, "table id"),
            (&table.name, "table name"),
            (&table.connection, "connection name"),
        ] {
            bounded_nonempty(field, label)?;
        }
        if !ids.insert(table.id.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate case-insensitive Data Model table id '{}'",
                table.id
            )));
        }
        if !names.insert(table.name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate case-insensitive Data Model table name '{}'",
                table.name
            )));
        }
    }
    let mut relationships = HashSet::new();
    for relationship in &value.relationships {
        for (field, label) in [
            (&relationship.from_table, "fromTable"),
            (&relationship.from_column, "fromColumn"),
            (&relationship.to_table, "toTable"),
            (&relationship.to_column, "toColumn"),
        ] {
            bounded_nonempty(field, label)?;
        }
        if !names.contains(&relationship.from_table.to_lowercase()) {
            return Err(invalid(format!(
                "relationship references unknown fromTable '{}'",
                relationship.from_table
            )));
        }
        if !names.contains(&relationship.to_table.to_lowercase()) {
            return Err(invalid(format!(
                "relationship references unknown toTable '{}'",
                relationship.to_table
            )));
        }
        let key = (
            relationship.from_table.to_lowercase(),
            relationship.from_column.to_lowercase(),
            relationship.to_table.to_lowercase(),
            relationship.to_column.to_lowercase(),
        );
        if !relationships.insert(key) {
            return Err(invalid(
                "duplicate case-insensitive Data Model relationship",
            ));
        }
    }
    if let Some(extension) = &value.extension_list {
        if extension.xml.len() > MAX_EXTENSION_BYTES {
            return Err(limit("extension bytes"));
        }
        if !extension_already_parsed {
            let root = parse_document(&extension.xml)?;
            require(&root, X15, "extLst")?;
        }
    }
    Ok(())
}

fn validate_payload(value: &DataModelPayload) -> Result<()> {
    if value.part_name != DATA_MODEL_PART_NAME {
        return Err(invalid(format!(
            "Data Model part must be '{DATA_MODEL_PART_NAME}'"
        )));
    }
    if value.data.is_empty() {
        return Err(invalid("Data Model payload cannot be empty"));
    }
    if value.data.len() > MAX_PAYLOAD_BYTES {
        return Err(limit("payload bytes"));
    }
    inspect_xldm(&value.data)?;
    PackURI::new(&value.part_name).map_err(OoxmlError::InvalidUri)?;
    Ok(())
}

fn workbook_definition(root: &Node) -> Result<(&str, Option<DataModelDefinition>)> {
    if root.name != "workbook" || !(root.namespace == SML || root.namespace == STRICT_SML) {
        return Err(invalid("expected SpreadsheetML workbook root"));
    }
    let core = root.namespace.as_str();
    let lists: Vec<_> = root
        .children
        .iter()
        .filter(|child| child.namespace == core && child.name == "extLst")
        .collect();
    if lists.len() > 1 {
        return Err(invalid("workbook has multiple direct extLst elements"));
    }
    let mut found = None;
    if let Some(list) = lists.first() {
        for extension in &list.children {
            if extension.namespace == core
                && extension.name == "ext"
                && optional(extension, "", "uri") == Some(DATA_MODEL_EXTENSION_URI)
            {
                if found.is_some() {
                    return Err(invalid("workbook has multiple dataModel extensions"));
                }
                no_attributes(extension, &[("", "uri")])?;
                whitespace(extension)?;
                if extension.children.len() != 1 {
                    return Err(invalid(
                        "dataModel extension must contain exactly one dataModel element",
                    ));
                }
                found = Some(parse_data_model_node(&extension.children[0])?);
            }
        }
    }
    Ok((core, found))
}

fn validate_connections(
    package: &OpcPackage,
    workbook_name: &PackURI,
    definition: &DataModelDefinition,
) -> Result<()> {
    if definition.tables.is_empty() {
        return Ok(());
    }
    let workbook = package.get_part(workbook_name)?;
    let mut relationships = workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            CONNECTIONS_RELATIONSHIP_TYPE | STRICT_CONNECTIONS_RELATIONSHIP_TYPE
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("Data Model tables require a workbook Connections part"))?;
    if relationships.next().is_some() {
        return Err(invalid("workbook has multiple Connections relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("Connections relationship cannot be external"));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(invalid(format!(
            "Connections part '{target}' has content type '{}'",
            part.content_type()
        )));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "Connections part has forbidden outbound relationships",
        ));
    }
    let connections = super::connections::Connections::parse(part.blob())
        .map_err(|error| invalid(format!("invalid Connections part: {error}")))?;
    let names: HashSet<String> = connections
        .connections
        .iter()
        .filter_map(|connection| connection.name.as_ref())
        .map(|name| name.to_lowercase())
        .collect();
    for table in &definition.tables {
        if !names.contains(&table.connection.to_lowercase()) {
            return Err(invalid(format!(
                "Data Model table '{}' references unknown workbook connection '{}'",
                table.name, table.connection
            )));
        }
    }
    Ok(())
}

fn reject_inbound_relationships(package: &OpcPackage, target: &PackURI) -> Result<()> {
    for relationship in package.rels().iter() {
        if !relationship.is_external()
            && relationship.target_partname()?.as_str() == target.as_str()
        {
            return Err(invalid(
                "package relationship targets the relationship-free Data Model part",
            ));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if !relationship.is_external()
                && relationship.target_partname()?.as_str() == target.as_str()
            {
                return Err(invalid(format!(
                    "part '{}' has a relationship to the relationship-free Data Model part",
                    source.partname()
                )));
            }
        }
    }
    Ok(())
}

fn insert_extension(xml: &[u8], core: &str, fragment: &[u8]) -> Result<Vec<u8>> {
    let new_size = xml
        .len()
        .checked_add(fragment.len())
        .and_then(|size| size.checked_add(core.len() + 64))
        .ok_or_else(|| limit("rewrite bytes"))?;
    if new_size > MAX_REWRITE_BYTES {
        return Err(limit("rewrite bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut open_ext = false;
    let mut empty_ext = None;
    let mut root_close = None;
    loop {
        let start =
            usize::try_from(reader.buffer_position()).map_err(|_| limit("rewrite position"))?;
        let event = reader.read_event().map_err(xml_error)?;
        let end =
            usize::try_from(reader.buffer_position()).map_err(|_| limit("rewrite position"))?;
        match event {
            Event::Start(element) => {
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                if depth == 1 && namespace == core && element.local_name().as_ref() == b"extLst" {
                    open_ext = true;
                }
                depth += 1;
            },
            Event::Empty(element) => {
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                if depth == 1 && namespace == core && element.local_name().as_ref() == b"extLst" {
                    empty_ext = Some((start, end, element.name().as_ref().to_vec()));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected workbook closing element"));
                }
                if depth == 2 && open_ext && element.local_name().as_ref() == b"extLst" {
                    let mut output = Vec::with_capacity(new_size);
                    output.extend_from_slice(&xml[..start]);
                    output.extend_from_slice(fragment);
                    output.extend_from_slice(&xml[start..]);
                    return Ok(output);
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                depth -= 1;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if let Some((start, end, qname)) = empty_ext {
        let raw = &xml[start..end];
        let close = memchr::memmem::rfind(raw, b"/>")
            .ok_or_else(|| invalid("invalid empty workbook extLst"))?;
        let mut output = Vec::with_capacity(new_size);
        output.extend_from_slice(&xml[..start + close]);
        output.push(b'>');
        output.extend_from_slice(fragment);
        output.extend_from_slice(b"</");
        output.extend_from_slice(&qname);
        output.push(b'>');
        output.extend_from_slice(&xml[end..]);
        return Ok(output);
    }
    let close = root_close.ok_or_else(|| invalid("missing workbook closing element"))?;
    let mut output = Vec::with_capacity(new_size);
    output.extend_from_slice(&xml[..close]);
    output.extend_from_slice(b"<extLst xmlns=\"");
    escape(&mut output, core);
    output.extend_from_slice(b"\">");
    output.extend_from_slice(fragment);
    output.extend_from_slice(b"</extLst>");
    output.extend_from_slice(&xml[close..]);
    Ok(output)
}

fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("XML bytes"));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
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
                    return Err(invalid("text outside XML root"));
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
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected in Data Model XML")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
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
fn serialize_node(node: &Node) -> Result<Vec<u8>> {
    let mut namespaces = BTreeMap::new();
    collect_namespaces(node, &mut namespaces);
    let mut prefixes = HashMap::new();
    let mut next = 0usize;
    for namespace in namespaces.keys() {
        let prefix = match namespace.as_str() {
            X15 => "x15".into(),
            SML | STRICT_SML => "x".into(),
            _ => {
                let prefix = format!("n{next}");
                next += 1;
                prefix
            },
        };
        prefixes.insert(namespace.clone(), prefix);
    }
    let mut output = Vec::new();
    write_node(node, &prefixes, true, &mut output);
    Ok(output)
}
fn collect_namespaces(node: &Node, output: &mut BTreeMap<String, ()>) {
    if !node.namespace.is_empty() {
        output.insert(node.namespace.clone(), ());
    }
    for attribute in &node.attributes {
        if !attribute.namespace.is_empty() {
            output.insert(attribute.namespace.clone(), ());
        }
    }
    for child in &node.children {
        collect_namespaces(child, output);
    }
}
fn write_node(node: &Node, prefixes: &HashMap<String, String>, root: bool, output: &mut Vec<u8>) {
    output.push(b'<');
    qname(output, &node.namespace, &node.name, prefixes);
    if root {
        let mut values: Vec<_> = prefixes.iter().collect();
        values.sort_by(|a, b| a.1.cmp(b.1));
        for (namespace, prefix) in values {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
            escape(output, namespace);
            output.push(b'\"');
        }
    }
    for attribute in &node.attributes {
        output.push(b' ');
        qname(output, &attribute.namespace, &attribute.name, prefixes);
        output.extend_from_slice(b"=\"");
        escape(output, &attribute.value);
        output.push(b'\"');
    }
    if node.children.is_empty() && node.text.is_empty() {
        output.extend_from_slice(b"/>");
        return;
    }
    output.push(b'>');
    escape_text(output, &node.text);
    for child in &node.children {
        write_node(child, prefixes, false, output);
    }
    output.extend_from_slice(b"</");
    qname(output, &node.namespace, &node.name, prefixes);
    output.push(b'>');
}
fn qname(output: &mut Vec<u8>, namespace: &str, name: &str, prefixes: &HashMap<String, String>) {
    if !namespace.is_empty() {
        output.extend_from_slice(prefixes[namespace].as_bytes());
        output.push(b':');
    }
    output.extend_from_slice(name.as_bytes());
}
fn require(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!("expected {{{namespace}}}{name}")))
    }
}
fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    for attribute in &node.attributes {
        if !allowed
            .iter()
            .any(|(namespace, name)| attribute.namespace == *namespace && attribute.name == *name)
        {
            return Err(invalid(format!(
                "unexpected attribute '{}' on {}",
                attribute.name, node.name
            )));
        }
    }
    Ok(())
}
fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .ok_or_else(|| invalid(format!("missing required {name} attribute")))
}
fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}
fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}
fn leaf(node: &Node) -> Result<()> {
    whitespace(node)?;
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{} cannot have children", node.name)))
    }
}
fn bounded_nonempty(value: &str, label: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{label} cannot be empty")));
    }
    if value.len() > MAX_STRING_BYTES {
        return Err(limit(label));
    }
    Ok(())
}
fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_TOTAL_STRING_BYTES {
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
fn escape_text(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
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
    invalid(format!("Data Model {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::Part;

    fn definition() -> DataModelDefinition {
        DataModelDefinition { min_version_load: 7, tables: vec![DataModelTable { id: "t-sales".into(), name: "Sales".into(), connection: "ModelConnection".into() }, DataModelTable { id: "t-date".into(), name: "Date".into(), connection: "ModelConnection".into() }], relationships: vec![DataModelRelationship { from_table: "Sales".into(), from_column: "DateKey".into(), to_table: "Date".into(), to_column: "DateKey".into() }], extension_list: Some(DataModelOpaqueXml { xml: format!(r#"<x15:extLst xmlns:x15="{X15}"><x15:ext uri="urn:test"><v:opaque xmlns:v="urn:vendor"/></x15:ext></x15:extLst>"#).into_bytes() }) }
    }
    fn package() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        let mut part = BlobPart::new(
            workbook.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            format!(r#"<workbook xmlns="{SML}"><sheets/></workbook>"#).into_bytes(),
        );
        let connections = PackURI::new("/xl/connections.xml").unwrap();
        part.rels_mut().add_relationship(
            CONNECTIONS_RELATIONSHIP_TYPE.into(),
            "connections.xml".into(),
            "rIdConnections".into(),
            false,
        );
        package.add_part(Box::new(part));
        package.add_part(Box::new(BlobPart::new(connections, CONNECTIONS_CONTENT_TYPE.into(), format!(r#"<connections xmlns="{SML}"><connection id="1" name="ModelConnection" refreshedVersion="7"/></connections>"#).into_bytes())));
        (package, workbook)
    }
    fn model() -> DataModel {
        DataModel {
            definition: definition(),
            payload: DataModelPayload {
                part_name: DATA_MODEL_PART_NAME.into(),
                data: super::super::xldm::test_xldm_bytes(),
            },
        }
    }

    #[test]
    fn typed_descriptor_round_trip() {
        let expected = definition();
        let xml = write_data_model(&expected).unwrap();
        let actual = parse_data_model(&xml).unwrap();
        assert_eq!(actual.min_version_load, expected.min_version_load);
        assert_eq!(actual.tables, expected.tables);
        assert_eq!(actual.relationships, expected.relationships);
        assert!(String::from_utf8_lossy(&actual.extension_list.unwrap().xml).contains("opaque"));
    }

    #[test]
    fn package_round_trip_preserves_inert_payload_and_inline_metadata() {
        let (mut package, workbook) = package();
        let expected = model();
        store_data_model(&mut package, &workbook, &expected).unwrap();
        let actual = load_data_model(&package, &workbook).unwrap().unwrap();
        assert_eq!(
            actual.definition.min_version_load,
            expected.definition.min_version_load
        );
        assert_eq!(actual.definition.tables, expected.definition.tables);
        assert_eq!(
            actual.definition.relationships,
            expected.definition.relationships
        );
        assert_eq!(actual.payload, expected.payload);
        assert!(
            String::from_utf8_lossy(&actual.definition.extension_list.unwrap().xml)
                .contains("opaque")
        );
    }

    #[test]
    fn inserts_into_existing_empty_extension_list() {
        let (mut package, workbook) = package();
        package.get_part_mut(&workbook).unwrap().set_blob(
            format!(r#"<workbook xmlns="{SML}"><sheets/><extLst /></workbook>"#).into_bytes(),
        );
        store_data_model(&mut package, &workbook, &model()).unwrap();
        assert!(load_data_model(&package, &workbook).unwrap().is_some());
    }

    #[test]
    fn rejects_hostile_xml_schema_and_bounds() {
        for xml in [
            format!(r#"<!DOCTYPE x><x15:dataModel xmlns:x15="{X15}"/>"#),
            format!(r#"<?bad x?><x15:dataModel xmlns:x15="{X15}"/>"#),
            format!(r#"<x15:dataModel xmlns:x15="{X15}" minVersionLoad="4"/>"#),
            format!(r#"<x15:dataModel xmlns:x15="{X15}"><x15:modelTables/></x15:dataModel>"#),
            format!(
                r#"<x15:dataModel xmlns:x15="{X15}"><x15:modelRelationships><x15:modelRelationship fromTable="Missing" fromColumn="a" toTable="Missing" toColumn="b"/></x15:modelRelationships></x15:dataModel>"#
            ),
        ] {
            assert!(parse_data_model(xml.as_bytes()).is_err());
        }
        assert!(parse_data_model(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
    }

    #[test]
    fn rejects_missing_connection_and_unknown_table_references() {
        let mut value = definition();
        value.tables[0].connection = "Absent".into();
        let (mut package, workbook) = package();
        assert!(
            store_data_model(
                &mut package,
                &workbook,
                &DataModel {
                    definition: value,
                    payload: model().payload
                }
            )
            .is_err()
        );
        let mut value = definition();
        value.relationships[0].to_table = "Absent".into();
        assert!(write_data_model(&value).is_err());
    }

    #[test]
    fn rejects_orphan_duplicate_wrong_path_and_relationship_edges() {
        let (mut pkg, workbook) = package();
        pkg.add_part(Box::new(BlobPart::new(
            PackURI::new(DATA_MODEL_PART_NAME).unwrap(),
            DATA_MODEL_CONTENT_TYPE.into(),
            vec![1],
        )));
        assert!(load_data_model(&pkg, &workbook).is_err());
        let (mut pkg, workbook) = package();
        store_data_model(&mut pkg, &workbook, &model()).unwrap();
        pkg.get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "model/item.data".into(),
                "rIdModel".into(),
                false,
            );
        assert!(load_data_model(&pkg, &workbook).is_err());
        let mut wrong = model();
        wrong.payload.part_name = "/xl/model/other.data".into();
        let (mut pkg, workbook) = package();
        assert!(store_data_model(&mut pkg, &workbook, &wrong).is_err());
    }
}
