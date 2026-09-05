//! Typed, inert inspection of MS-XLDM section 2.5 table metadata.
//!
//! XML is bounded and validated against the section 2.5 object grammar. The
//! original bytes and native payloads remain borrowed. Compression metadata is
//! reported and cross-checked, but no section 2.7 algorithm is executed.

use std::collections::HashSet;
use std::fmt;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

use super::super::generated::{SystemGeneratedData, SystemGeneratedFile, SystemGeneratedKind};
use super::super::native::{
    DictionaryBody, DictionaryType, NativeData, NativeFile, StringPageData,
};
use super::super::{GeneratedNameKind, Storage, classify_generated_path};
use super::model::{
    ColumnPolicy, DictionaryPolicy, HierarchyPolicy, MetadataClass, MetadataCollection,
    MetadataDataObject, MetadataError, MetadataFile, MetadataFileKind, MetadataMember,
    MetadataModel, MetadataObject, MetadataProperty, MetadataResult, RelationshipIndexKind,
    RelationshipPolicy,
};
use super::model::{hybrid_width, no_split_width};

const MAX_METADATA_BYTES: usize = 16 * 1024 * 1024;
const MAX_XML_NODES: usize = 500_000;
const MAX_XML_DEPTH: usize = 128;
const MAX_TEXT_BYTES: usize = 32 * 1024 * 1024;
const MAX_METADATA_FILES: usize = 65_536;
#[derive(Clone, Debug)]
struct XmlNode {
    name: String,
    attributes: Vec<(String, String)>,
    children: Vec<XmlNode>,
    text: String,
}

/// Parse and validate one section 2.5 `.tbl.xml` member.
pub fn parse_file<'a>(storage_path: &'a str, bytes: &'a [u8]) -> MetadataResult<MetadataFile<'a>> {
    let generated = classify_generated_path(storage_path).map_err(|error| {
        MetadataError::new(format!(
            "invalid generated metadata path {storage_path}: {error}"
        ))
    })?;
    let kind = match generated.kind {
        GeneratedNameKind::TableMetadata => MetadataFileKind::Table,
        GeneratedNameKind::ColumnHierarchyMetadata => MetadataFileKind::ColumnHierarchy,
        GeneratedNameKind::UserHierarchyMetadata => MetadataFileKind::UserHierarchy,
        GeneratedNameKind::TableRelationshipMetadata => MetadataFileKind::TableRelationship,
        GeneratedNameKind::DatabaseDefinition
        | GeneratedNameKind::DataSourceViewDefinition
        | GeneratedNameKind::CubeDefinition
        | GeneratedNameKind::DataSourceOrDimensionDefinition
        | GeneratedNameKind::CubeInformation
        | GeneratedNameKind::PartitionInformation
        | GeneratedNameKind::TableInformation
        | GeneratedNameKind::MdxScriptMetadata
        | GeneratedNameKind::MeasureGroupMetadata
        | GeneratedNameKind::PartitionMetadata
        | GeneratedNameKind::ColumnData
        | GeneratedNameKind::TableRelationshipIndex
        | GeneratedNameKind::ColumnPositionToId
        | GeneratedNameKind::ColumnIdToPosition
        | GeneratedNameKind::ColumnHashIndex
        | GeneratedNameKind::ColumnDictionary
        | GeneratedNameKind::UserHierarchyChildCount
        | GeneratedNameKind::UserHierarchyFirstChildPosition
        | GeneratedNameKind::UserHierarchyParentPosition
        | GeneratedNameKind::UserHierarchyMultilevelId => {
            return Err(MetadataError::new(
                "section 2.5 metadata requires a .tbl.xml path",
            ));
        },
    };
    let node = parse_xml(bytes)?;
    let table = project_object(node)?;
    if table.class.as_str() != "XMSimpleTable" {
        return Err(MetadataError::new("a .tbl.xml root MUST be XMSimpleTable"));
    }
    validate_table_file_kind(kind, &table)?;
    Ok(MetadataFile {
        storage_path,
        bytes,
        kind,
        table,
    })
}

/// Discover every logged `.tbl.xml` member and derive its section 2.3 policy.
pub fn inspect<'a>(storage: &'a Storage<'a>) -> MetadataResult<MetadataModel<'a>> {
    let mut files = Vec::new();
    let mut seen = HashSet::new();
    for group in &storage.backup_log.file_groups {
        for logged in &group.files {
            let path = logged.storage_path.as_str();
            if !path.ends_with(".tbl.xml") {
                continue;
            }
            if !seen.insert(path) {
                return Err(MetadataError::new(format!(
                    "duplicate logged metadata member {path}"
                )));
            }
            if files.len() == MAX_METADATA_FILES {
                return Err(MetadataError::new("too many table metadata files"));
            }
            let index = storage
                .files
                .iter()
                .position(|entry| entry.path == path)
                .ok_or_else(|| {
                    MetadataError::new(format!(
                        "logged metadata member {path} is absent from the directory"
                    ))
                })?;
            let bytes = storage.file_payload(index).ok_or_else(|| {
                MetadataError::new(format!("cannot resolve metadata member {path}"))
            })?;
            files.push(parse_file(path, bytes)?);
        }
    }
    let (columns, relationships, hierarchies) = derive_policies(&files)?;
    Ok(MetadataModel {
        files,
        columns,
        relationships,
        hierarchies,
    })
}

/// Revalidate an inspected file and return its exact original XML bytes.
pub fn write_file(file: &MetadataFile<'_>) -> MetadataResult<Vec<u8>> {
    let reparsed = parse_file(file.storage_path, file.bytes)?;
    if reparsed.kind != file.kind || reparsed.table != file.table {
        return Err(MetadataError::new("metadata object graph was mutated"));
    }
    Ok(file.bytes.to_vec())
}

fn parse_xml(bytes: &[u8]) -> MetadataResult<XmlNode> {
    if bytes.len() > MAX_METADATA_BYTES {
        return Err(MetadataError::new("metadata XML is too large"));
    }
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| MetadataError::new(format!("invalid metadata XML: {error}")))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| MetadataError::new("XML node count overflow"))?;
                if nodes > MAX_XML_NODES || stack.len() >= MAX_XML_DEPTH {
                    return Err(MetadataError::new("metadata XML structure limit exceeded"));
                }
                let empty = matches!(event, Event::Empty(_));
                let node = make_xml_node(element, decoder)?;
                if empty {
                    attach_node(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| MetadataError::new("unexpected XML end element"))?;
                attach_node(node, &mut stack, &mut root)?;
            },
            Event::Text(value) => {
                let decoded = value.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                text_bytes = text_bytes
                    .checked_add(decoded.len())
                    .ok_or_else(|| MetadataError::new("XML text size overflow"))?;
                if text_bytes > MAX_TEXT_BYTES {
                    return Err(MetadataError::new("metadata XML text limit exceeded"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(MetadataError::new("text outside metadata root"));
                }
            },
            Event::CData(value) => {
                let decoded = value.decode().map_err(xml_error)?;
                text_bytes = text_bytes
                    .checked_add(decoded.len())
                    .ok_or_else(|| MetadataError::new("XML text size overflow"))?;
                if text_bytes > MAX_TEXT_BYTES {
                    return Err(MetadataError::new("metadata XML text limit exceeded"));
                }
                stack
                    .last_mut()
                    .ok_or_else(|| MetadataError::new("CDATA outside metadata root"))?
                    .text
                    .push_str(&decoded);
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|c| c.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| MetadataError::new("custom XML entities are rejected"))?;
                stack
                    .last_mut()
                    .ok_or_else(|| MetadataError::new("entity outside metadata root"))?
                    .text
                    .push_str(&value);
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(MetadataError::new(
                    "DTD and processing instructions are rejected",
                ));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) => {},
        }
    }
    if !stack.is_empty() {
        return Err(MetadataError::new("unclosed metadata XML element"));
    }
    root.ok_or_else(|| MetadataError::new("missing metadata XML root"))
}

fn make_xml_node(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> MetadataResult<XmlNode> {
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    if name.contains(':') {
        return Err(MetadataError::new(
            "namespaced metadata elements are not allowed",
        ));
    }
    let mut attributes = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(xml_error)?
            .to_owned();
        if key == "xmlns" || key.contains(':') {
            return Err(MetadataError::new("metadata namespaces are not allowed"));
        }
        if attributes.iter().any(|(name, _)| name == &key) {
            return Err(MetadataError::new("duplicate metadata attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        attributes.push((key, value));
    }
    Ok(XmlNode {
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}

fn attach_node(
    node: XmlNode,
    stack: &mut [XmlNode],
    root: &mut Option<XmlNode>,
) -> MetadataResult<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(MetadataError::new("multiple metadata XML roots"));
    }
    Ok(())
}

fn project_object(node: XmlNode) -> MetadataResult<MetadataObject> {
    if node.name != "XMObject" {
        return Err(MetadataError::new("expected XMObject"));
    }
    if !node.text.trim().is_empty() {
        return Err(MetadataError::new("XMObject cannot contain direct text"));
    }
    let mut class = None;
    let mut name = None;
    let mut provider_version = None;
    for (key, value) in node.attributes {
        match key.as_str() {
            "class" => class = Some(parse_class(value)?),
            "name" => name = Some(value),
            "ProviderVersion" => provider_version = Some(parse_i32(&value, "ProviderVersion")?),
            _ => {
                return Err(MetadataError::new(format!(
                    "unknown XMObject attribute {key}"
                )));
            },
        }
    }
    let class = class.ok_or_else(|| MetadataError::new("XMObject requires a class attribute"))?;
    let mut properties = None;
    let mut members = None;
    let mut collections = None;
    let mut data_objects = None;
    for child in node.children {
        match child.name.as_str() {
            "Properties" if properties.is_none() => properties = Some(project_properties(child)?),
            "Members" if members.is_none() => members = Some(project_members(child)?),
            "Collections" if collections.is_none() => {
                collections = Some(project_collections(child)?);
            },
            "DataObjects" if data_objects.is_none() => {
                data_objects = Some(project_data_objects(child)?);
            },
            "Properties" | "Members" | "Collections" | "DataObjects" => {
                return Err(MetadataError::new("duplicate XMObject container"));
            },
            _ => {
                return Err(MetadataError::new(format!(
                    "unknown XMObject child {}",
                    child.name
                )));
            },
        }
    }
    let object = MetadataObject {
        class,
        name,
        provider_version,
        properties: properties.unwrap_or_default(),
        members: members.unwrap_or_default(),
        collections: collections.unwrap_or_default(),
        data_objects: data_objects.unwrap_or_default(),
    };
    validate_object(&object)?;
    Ok(object)
}

fn project_properties(node: XmlNode) -> MetadataResult<Vec<MetadataProperty>> {
    plain_container(&node, "Properties")?;
    let mut result = Vec::new();
    for child in node.children {
        if !child.attributes.is_empty() || !child.children.is_empty() {
            return Err(MetadataError::new("metadata properties MUST be scalar"));
        }
        if result
            .iter()
            .any(|item: &MetadataProperty| item.name == child.name)
        {
            return Err(MetadataError::new(format!(
                "duplicate property {}",
                child.name
            )));
        }
        result.push(MetadataProperty {
            name: child.name,
            value: child.text.trim().to_owned(),
        });
    }
    Ok(result)
}

fn project_members(node: XmlNode) -> MetadataResult<Vec<MetadataMember>> {
    plain_container(&node, "Members")?;
    let mut result = Vec::new();
    for child in node.children {
        plain_container(&child, "Member")?;
        if child.children.len() != 2
            || child.children[0].name != "Name"
            || child.children[1].name != "XMObject"
        {
            return Err(MetadataError::new(
                "Member requires Name followed by XMObject",
            ));
        }
        let name = scalar_text(&child.children[0], "Name")?;
        result.push(MetadataMember {
            name,
            object: Box::new(project_object(child.children[1].clone())?),
        });
    }
    Ok(result)
}

fn project_collections(node: XmlNode) -> MetadataResult<Vec<MetadataCollection>> {
    plain_container(&node, "Collections")?;
    let mut result = Vec::new();
    for child in node.children {
        plain_container(&child, "Collection")?;
        if child
            .children
            .first()
            .is_none_or(|item| item.name != "Name")
        {
            return Err(MetadataError::new("Collection requires Name first"));
        }
        let name = scalar_text(&child.children[0], "Name")?;
        let mut objects = Vec::new();
        for object in child.children.into_iter().skip(1) {
            objects.push(project_object(object)?);
        }
        result.push(MetadataCollection { name, objects });
    }
    Ok(result)
}

fn project_data_objects(node: XmlNode) -> MetadataResult<Vec<MetadataDataObject>> {
    plain_container(&node, "DataObjects")?;
    let mut result = Vec::new();
    for child in node.children {
        plain_container(&child, "DataObject")?;
        if child.children.len() != 1 {
            return Err(MetadataError::new(
                "DataObject requires exactly one XMObject",
            ));
        }
        result.push(MetadataDataObject {
            object: Box::new(project_object(child.children[0].clone())?),
        });
    }
    Ok(result)
}

fn plain_container(node: &XmlNode, expected: &str) -> MetadataResult<()> {
    if node.name != expected || !node.attributes.is_empty() || !node.text.trim().is_empty() {
        return Err(MetadataError::new(format!("invalid {expected} container")));
    }
    Ok(())
}

fn scalar_text(node: &XmlNode, expected: &str) -> MetadataResult<String> {
    if node.name != expected || !node.attributes.is_empty() || !node.children.is_empty() {
        return Err(MetadataError::new(format!("invalid scalar {expected}")));
    }
    Ok(node.text.trim().to_owned())
}

fn parse_class(value: String) -> MetadataResult<MetadataClass> {
    let fixed = [
        "XMSimpleTable",
        "XMRawColumn",
        "XMRelationship",
        "XMRelationshipIndexSparseDIDs",
        "XMRelationshipIndexDenseDIDs",
        "XMRelationshipIndex123DIDs",
        "XMHierarchy",
        "XMUserHierarchy",
        "XMHierarchyDataID2PositionHashIndex",
        "XMColumnSegment",
        "XMPartition",
        "XMMultiPartSegmentMap",
        "XMSegment1Map",
        "XMTableStats",
        "XMColumnStats",
        "XMColumnSegmentStats",
        "XMValueDataDictionary<XM_Long>",
        "XMValueDataDictionary<XM_Real>",
        "XMHashDataDictionary<XM_Real>",
        "XMHashDataDictionary<XM_Long>",
        "XMHashDataDictionary<XM_String>",
        "XMHashDataDictionary<XMVariantPtr>",
        "XM123CompressionInfo",
        "XMRawColumnPartitionDataObject",
        "XMRLECompressionInfo",
        "XMRLEGeneralCompressionInfo",
        "XMColumnSegmentDataObject",
        "XMSegmentEqualMapEx<XMSegmentEqualMap_FastInstantiation>",
        "XMSegmentEqualMapEx<XMSegmentEqualMap_ComplexInstantiation>",
        "XMHybridRLECompressionInfo<class XM123CompressionInfo>",
        "XMHybridRLECompressionInfo<class XMREGeneralCompressionInfo>",
    ];
    let valid = fixed.contains(&value.as_str())
        || no_split_width(&value).is_some()
        || hybrid_width(&value).is_some();
    if !valid {
        return Err(MetadataError::new(format!(
            "unknown XMObject class {value}"
        )));
    }
    Ok(MetadataClass(value))
}

fn validate_object(object: &MetadataObject) -> MetadataResult<()> {
    let class = object.class.as_str();
    let (properties, containers): (&[&str], u8) = match class {
        "XMSimpleTable" => (&["Version", "Settings", "RIViolationCount"], 0b0111),
        "XMTableStats" => (&["SegmentSize", "Usage"], 0b0001),
        "XMRawColumn" => (
            &[
                "Settings",
                "ColumnFlags",
                "Collation",
                "OrderByColumn",
                "Locale",
                "BinaryCharacters",
            ],
            0b1111,
        ),
        "XMRelationship" => (&["PrimaryTable", "PrimaryColumn", "ForeignColumn"], 0b1001),
        "XMRelationshipIndexSparseDIDs" => (&["Flags"], 0b0001),
        "XMRelationshipIndexDenseDIDs" => (&["Records", "TableName", "Flags"], 0b0001),
        "XMRelationshipIndex123DIDs" | "XMHierarchyDataID2PositionHashIndex" => (&[], 0),
        "XMColumnStats" => (
            &[
                "DistinctStates",
                "MinDataID",
                "MaxDataID",
                "OriginalMinSegmentDataID",
                "RLESortOrder",
                "RowCount",
                "HasNulls",
                "RLERuns",
                "OthersRLERuns",
                "Usage",
                "DBType",
                "XMType",
                "CompressionType",
                "CompressionParam",
                "EncodingHint",
                "AggCounter",
                "WhereCounter",
                "OrderByCounter",
            ],
            0b0001,
        ),
        "XMHierarchy" => (
            &[
                "SortOrder",
                "IsProcessed",
                "TypeMaterialization",
                "ColumnPosition2DataID",
                "ColumnDataID2Position",
                "DistinctDataIDs",
                "TableStore",
            ],
            0b0001,
        ),
        "XMUserHierarchy" => (&["IsProcessed", "TableStore", "TableName"], 0b0001),
        "XMColumnSegment" => (&["Records", "Mask"], 0b0011),
        "XMPartition" => (&["IsProcessed", "Partition"], 0b0001),
        "XMMultiPartSegmentMap" => (
            &["FirstPartitionRecordCount", "FirstPartitionSegmentCount"],
            0b0101,
        ),
        "XMSegment1Map" => (&["Records"], 0b0001),
        "XMSegmentEqualMapEx<XMSegmentEqualMap_FastInstantiation>"
        | "XMSegmentEqualMapEx<XMSegmentEqualMap_ComplexInstantiation>" => {
            (&["Segments", "Records", "RecordsPerSegment"], 0b0001)
        },
        "XMValueDataDictionary<XM_Long>" | "XMValueDataDictionary<XM_Real>" => {
            (&["DataVersion", "BaseId", "Magnitude"], 0b0001)
        },
        "XMHashDataDictionary<XM_Real>" => {
            (&["DataVersion", "LastId", "Nullable", "Unique"], 0b0001)
        },
        "XMHashDataDictionary<XM_Long>" => (
            &[
                "DataVersion",
                "LastId",
                "Nullable",
                "Unique",
                "OperatingOn32",
            ],
            0b0001,
        ),
        "XMHashDataDictionary<XM_String>" => (
            &[
                "DataVersion",
                "LastId",
                "Nullable",
                "Unique",
                "DictionaryFlags",
            ],
            0b0001,
        ),
        "XMColumnSegmentStats" => (
            &[
                "DistinctStates",
                "MinDataID",
                "MaxDataID",
                "OriginalMinSegmentDataID",
                "RLESortOrder",
                "RowCount",
                "HasNulls",
                "RLERuns",
                "OthersRLERuns",
            ],
            0b0001,
        ),
        "XMRawColumnPartitionDataObject" => (&["DataVersion", "Partition", "SegmentCount"], 0b0001),
        "XMRLECompressionInfo" => (
            &[
                "BookmarkBits",
                "StorageAllocSize",
                "StorageUsedSize",
                "SegmentNeedsResizing",
            ],
            0b0001,
        ),
        "XM123CompressionInfo" => (&["Min"], 0b0001),
        value if no_split_width(value).is_some() => (&["Min"], 0b0001),
        value if value.starts_with("XMHybridRLECompressionInfo<class ") => (&[], 0b0010),
        _ => (&[], 0),
    };
    let actual = u8::from(!object.properties.is_empty())
        | (u8::from(!object.members.is_empty()) << 1)
        | (u8::from(!object.collections.is_empty()) << 2)
        | (u8::from(!object.data_objects.is_empty()) << 3);
    if actual != containers {
        return Err(MetadataError::new(format!(
            "{class} has invalid or missing containers"
        )));
    }
    require_properties(object, properties)?;
    validate_scalar_constraints(object)?;
    validate_nested_constraints(object)
}

fn require_properties(object: &MetadataObject, required: &[&str]) -> MetadataResult<()> {
    if object.properties.len() != required.len()
        || required.iter().any(|name| object.property(name).is_none())
    {
        return Err(MetadataError::new(format!(
            "{} has an invalid property set",
            object.class.as_str()
        )));
    }
    Ok(())
}

fn validate_scalar_constraints(object: &MetadataObject) -> MetadataResult<()> {
    let class = object.class.as_str();
    for property in &object.properties {
        let name = property.name.as_str();
        let value = property.value.as_str();
        if bool_property(name) {
            parse_bool(value, name)?;
        } else if double_property(name) {
            value
                .parse::<f64>()
                .map_err(|_source| MetadataError::new(format!("invalid double {name}")))?;
        } else if !string_property(name) {
            parse_i64(value, name)?;
        }
    }
    let ranged = |name, min, max| -> MetadataResult<()> {
        if let Some(value) = object.property(name) {
            let value = parse_i64(value, name)?;
            if value < min || value > max {
                return Err(MetadataError::new(format!(
                    "{class}.{name} is outside {min}..={max}"
                )));
            }
        }
        Ok(())
    };
    match class {
        "XMSimpleTable" => ranged("Settings", 0, 4367)?,
        "XMRawColumn" => {
            ranged("Settings", 0, 7994)?;
            ranged("ColumnFlags", 0, 63)?;
            if parse_i64(required_property(object, "ColumnFlags")?, "ColumnFlags")? & 0x8 == 0 {
                return Err(MetadataError::new(
                    "XMRawColumn.ColumnFlags MUST contain 0x8",
                ));
            }
        },
        "XMRelationshipIndexSparseDIDs" | "XMRelationshipIndexDenseDIDs" => ranged("Flags", 0, 1)?,
        "XMTableStats" => ranged("Usage", 0, 2)?,
        "XMColumnStats" => {
            ranged("Usage", 0, 3)?;
            ranged("DBType", 0, 130)?;
            ranged("XMType", 0, 3)?;
            ranged("CompressionType", 0, 2)?;
            ranged("EncodingHint", 0, 2)?;
            if parse_i64(required_property(object, "RLESortOrder")?, "RLESortOrder")? != -1 {
                return Err(MetadataError::new("XMColumnStats.RLESortOrder MUST be -1"));
            }
            let db = parse_i64(required_property(object, "DBType")?, "DBType")?;
            if db > 29 && db != 128 && db != 130 {
                return Err(MetadataError::new("XMColumnStats.DBType is reserved"));
            }
        },
        "XMHierarchy" => {
            let sort = parse_i64(required_property(object, "SortOrder")?, "SortOrder")?;
            if ![0, 2].contains(&sort) {
                return Err(MetadataError::new("XMHierarchy.SortOrder MUST be 0 or 2"));
            }
            let materialization = parse_i64(
                required_property(object, "TypeMaterialization")?,
                "TypeMaterialization",
            )?;
            if ![-1, 0, 1, 2, 3].contains(&materialization) {
                return Err(MetadataError::new("invalid hierarchy materialization"));
            }
            if materialization == -1
                && parse_bool(required_property(object, "IsProcessed")?, "IsProcessed")?
            {
                return Err(MetadataError::new(
                    "processed hierarchy cannot have unspecified materialization",
                ));
            }
            let p2i = parse_i64(
                required_property(object, "ColumnPosition2DataID")?,
                "ColumnPosition2DataID",
            )?;
            let i2p = parse_i64(
                required_property(object, "ColumnDataID2Position")?,
                "ColumnDataID2Position",
            )?;
            if ![-1, 0].contains(&p2i) || ![-1, 1].contains(&i2p) {
                return Err(MetadataError::new("invalid hierarchy index declaration"));
            }
        },
        "XMHashDataDictionary<XM_String>" => {
            ranged("DictionaryFlags", 0, 263)?;
            let flags = parse_i64(
                required_property(object, "DictionaryFlags")?,
                "DictionaryFlags",
            )?;
            if flags & !0x107 != 0 {
                return Err(MetadataError::new("DictionaryFlags contains reserved bits"));
            }
        },
        "XMColumnSegment" => ranged("Mask", 0, 2)?,
        "XMColumnSegmentStats"
            if parse_i64(required_property(object, "RLESortOrder")?, "RLESortOrder")? != -1 =>
        {
            return Err(MetadataError::new(
                "XMColumnSegmentStats.RLESortOrder MUST be -1",
            ));
        },
        "XMRLECompressionInfo"
            if parse_bool(
                required_property(object, "SegmentNeedsResizing")?,
                "SegmentNeedsResizing",
            )? =>
        {
            return Err(MetadataError::new("SegmentNeedsResizing MUST be false"));
        },
        _ => {},
    }
    Ok(())
}

fn validate_nested_constraints(object: &MetadataObject) -> MetadataResult<()> {
    let class = object.class.as_str();
    match class {
        "XMSimpleTable" => {
            require_member_classes(
                object,
                &[
                    (
                        "SegmentMap",
                        &[
                            "XMMultiPartSegmentMap",
                            "XMSegment1Map",
                            "XMSegmentEqualMapEx<XMSegmentEqualMap_FastInstantiation>",
                            "XMSegmentEqualMapEx<XMSegmentEqualMap_ComplexInstantiation>",
                        ],
                    ),
                    ("TableStats", &["XMTableStats"]),
                ],
            )?;
            require_collections(
                object,
                &[
                    ("Partitions", "XMPartition"),
                    ("Columns", "XMRawColumn"),
                    ("Relationships", "XMRelationship"),
                    ("UserHierarchies", "XMUserHierarchy"),
                ],
            )?;
        },
        "XMRawColumn" => {
            require_member_classes(
                object,
                &[
                    ("IntrinsicHierarchy", &["XMHierarchy"]),
                    ("ColumnStats", &["XMColumnStats"]),
                ],
            )?;
            require_collections(object, &[("Segments", "XMColumnSegment")])?;
            if object.data_objects.len() != 2
                || object
                    .data_objects
                    .iter()
                    .filter(|item| item.object.class.as_str() == "XMRawColumnPartitionDataObject")
                    .count()
                    != 1
            {
                return Err(MetadataError::new(
                    "XMRawColumn requires partition and dictionary/index data objects",
                ));
            }
            let allowed = [
                "XMRawColumnPartitionDataObject",
                "XMValueDataDictionary<XM_Long>",
                "XMValueDataDictionary<XM_Real>",
                "XMHashDataDictionary<XM_Real>",
                "XMHashDataDictionary<XM_Long>",
                "XMHashDataDictionary<XM_String>",
                "XMHierarchyDataID2PositionHashIndex",
            ];
            if object
                .data_objects
                .iter()
                .any(|item| !allowed.contains(&item.object.class.as_str()))
            {
                return Err(MetadataError::new("invalid XMRawColumn data object class"));
            }
        },
        "XMRelationship"
            if (object.data_objects.len() != 1
                || !matches!(
                    object.data_objects[0].object.class.as_str(),
                    "XMRelationshipIndexSparseDIDs"
                        | "XMRelationshipIndexDenseDIDs"
                        | "XMRelationshipIndex123DIDs"
                )) =>
        {
            return Err(MetadataError::new(
                "XMRelationship requires one relationship index object",
            ));
        },
        "XMColumnSegment" => require_member_classes(
            object,
            &[
                ("SubSegment", &["XMColumnSegment"]),
                ("CompressionInfo", &["@hybrid"]),
                ("ColumnSegmentStats", &["XMColumnSegmentStats"]),
            ],
        )?,
        "XMMultiPartSegmentMap" => require_collections_multi(
            object,
            &[(
                "Partitions",
                &[
                    "XMSegment1Map",
                    "XMSegmentEqualMapEx<XMSegmentEqualMap_FastInstantiation>",
                    "XMSegmentEqualMapEx<XMSegmentEqualMap_ComplexInstantiation>",
                ],
            )],
        )?,
        value if value.starts_with("XMHybridRLECompressionInfo<class ") => {
            require_member_classes(
                object,
                &[
                    ("RLECompression", &["XMRLECompressionInfo"]),
                    ("SubCompression", &["@matching_subcompression"]),
                ],
            )?;
            let sub = object
                .member("SubCompression")
                .unwrap_or_else(|| crate::error::panic_missing_invariant("validated member"));
            let correct = if let Some(width) = hybrid_width(value) {
                sub.class.no_split_width() == Some(width) && !sub.class.is_hybrid_compression()
            } else if value.contains("XM123CompressionInfo") {
                sub.class.as_str() == "XM123CompressionInfo"
            } else {
                sub.class.as_str() == "XMRLEGeneralCompressionInfo"
            };
            if !correct {
                return Err(MetadataError::new(
                    "hybrid SubCompression does not match its outer class",
                ));
            }
        },
        _ => {},
    }
    Ok(())
}

fn require_member_classes(
    object: &MetadataObject,
    required: &[(&str, &[&str])],
) -> MetadataResult<()> {
    if object.members.len() != required.len() {
        return Err(MetadataError::new(format!(
            "{} has an invalid member count",
            object.class.as_str()
        )));
    }
    for (name, classes) in required {
        let member = object
            .members
            .iter()
            .find(|item| item.name == *name)
            .ok_or_else(|| MetadataError::new(format!("missing member {name}")))?;
        let class = member.object.class.as_str();
        let valid = classes.contains(&class)
            || (classes.contains(&"@hybrid") && member.object.class.is_hybrid_compression())
            || classes.contains(&"@matching_subcompression");
        if !valid {
            return Err(MetadataError::new(format!(
                "member {name} has invalid class {class}"
            )));
        }
    }
    Ok(())
}

fn require_collections(object: &MetadataObject, required: &[(&str, &str)]) -> MetadataResult<()> {
    if object.collections.len() != required.len() {
        return Err(MetadataError::new(format!(
            "{} has an invalid collection count",
            object.class.as_str()
        )));
    }
    for (name, class) in required {
        let collection = object
            .collections
            .iter()
            .find(|item| item.name == *name)
            .ok_or_else(|| MetadataError::new(format!("missing collection {name}")))?;
        if collection
            .objects
            .iter()
            .any(|item| item.class.as_str() != *class)
        {
            return Err(MetadataError::new(format!(
                "collection {name} has an invalid object class"
            )));
        }
    }
    Ok(())
}

fn require_collections_multi(
    object: &MetadataObject,
    required: &[(&str, &[&str])],
) -> MetadataResult<()> {
    if object.collections.len() != required.len() {
        return Err(MetadataError::new("invalid collection count"));
    }
    for (name, classes) in required {
        let collection = object
            .collections
            .iter()
            .find(|item| item.name == *name)
            .ok_or_else(|| MetadataError::new(format!("missing collection {name}")))?;
        if collection
            .objects
            .iter()
            .any(|item| !classes.contains(&item.class.as_str()))
        {
            return Err(MetadataError::new(format!(
                "collection {name} has an invalid class"
            )));
        }
    }
    Ok(())
}

fn validate_table_file_kind(kind: MetadataFileKind, table: &MetadataObject) -> MetadataResult<()> {
    let columns = table
        .collection("Columns")
        .unwrap_or_else(|| crate::error::panic_missing_invariant("validated table"));
    let roles: Vec<i64> = columns
        .iter()
        .map(|column| {
            parse_i64(
                column
                    .property("Settings")
                    .unwrap_or_else(|| crate::error::panic_missing_invariant("validated column")),
                "Settings",
            )
            .map(|value| value & 0x1f)
        })
        .collect::<Result<_, _>>()?;
    match kind {
        MetadataFileKind::ColumnHierarchy if roles.iter().any(|role| ![5, 7].contains(role)) => {
            return Err(MetadataError::new(
                "column hierarchy metadata contains a non-index column",
            ));
        },
        MetadataFileKind::UserHierarchy
            if roles.len() != 4 || ![8, 9, 16, 17].iter().all(|role| roles.contains(role)) =>
        {
            return Err(MetadataError::new(
                "user hierarchy metadata requires all four generated columns",
            ));
        },
        MetadataFileKind::TableRelationship if roles.len() != 1 => {
            return Err(MetadataError::new(
                "relationship metadata requires exactly one generated column",
            ));
        },
        MetadataFileKind::Table
        | MetadataFileKind::ColumnHierarchy
        | MetadataFileKind::UserHierarchy
        | MetadataFileKind::TableRelationship => {},
    }
    Ok(())
}

fn derive_policies(
    files: &[MetadataFile<'_>],
) -> MetadataResult<(
    Vec<ColumnPolicy>,
    Vec<RelationshipPolicy>,
    Vec<HierarchyPolicy>,
)> {
    let mut columns = Vec::new();
    let mut relationships = Vec::new();
    let mut hierarchies = Vec::new();
    for file in files {
        for column in file
            .table
            .collection("Columns")
            .unwrap_or_else(|| crate::error::panic_missing_invariant("validated table"))
        {
            let partition = column
                .data_objects
                .iter()
                .find(|item| item.object.class.as_str() == "XMRawColumnPartitionDataObject")
                .unwrap_or_else(|| crate::error::panic_missing_invariant("validated raw column"))
                .object
                .as_ref();
            let stats = column
                .member("ColumnStats")
                .unwrap_or_else(|| crate::error::panic_missing_invariant("validated raw column"));
            let dictionary_object = column
                .data_objects
                .iter()
                .map(|item| item.object.as_ref())
                .find(|item| item.class.as_str() != "XMRawColumnPartitionDataObject")
                .unwrap_or_else(|| crate::error::panic_missing_invariant("validated raw column"));
            let dictionary = if dictionary_object
                .class
                .as_str()
                .starts_with("XMHashDataDictionary")
            {
                Some(DictionaryPolicy {
                    storage_name: qualify_storage_name(
                        file.storage_path,
                        dictionary_object.name.as_deref().ok_or_else(|| {
                            MetadataError::new("hash dictionary requires its storage name")
                        })?,
                    ),
                    class: dictionary_object.class.clone(),
                    dictionary_flags: dictionary_object
                        .property("DictionaryFlags")
                        .map(|value| {
                            parse_i64(value, "DictionaryFlags").and_then(|value| {
                                u16::try_from(value).map_err(|_source| {
                                    MetadataError::new("DictionaryFlags overflow")
                                })
                            })
                        })
                        .transpose()?,
                    operating_on_32: dictionary_object
                        .property("OperatingOn32")
                        .map(|value| parse_bool(value, "OperatingOn32"))
                        .transpose()?,
                })
            } else {
                None
            };
            let segment_count = usize::try_from(parse_i64(
                required_property(partition, "SegmentCount")?,
                "SegmentCount",
            )?)
            .map_err(|_source| MetadataError::new("negative or overflowing SegmentCount"))?;
            let row_count = u64::try_from(parse_i64(
                required_property(stats, "RowCount")?,
                "RowCount",
            )?)
            .map_err(|_source| MetadataError::new("negative RowCount"))?;
            let compression_type = u8::try_from(parse_i64(
                required_property(stats, "CompressionType")?,
                "CompressionType",
            )?)
            .map_err(|_source| MetadataError::new("CompressionType overflow"))?;
            let compression_param = parse_i64(
                required_property(stats, "CompressionParam")?,
                "CompressionParam",
            )?;
            let segments = column
                .collection("Segments")
                .unwrap_or_else(|| crate::error::panic_missing_invariant("validated raw column"));
            if segments.len() != segment_count {
                return Err(MetadataError::new(
                    "SegmentCount disagrees with the Segments collection",
                ));
            }
            let mut segment_rows = 0u64;
            for segment in segments {
                let records = u64::try_from(parse_i64(
                    required_property(segment, "Records")?,
                    "Records",
                )?)
                .map_err(|_source| MetadataError::new("negative segment Records"))?;
                let stats = segment
                    .member("ColumnSegmentStats")
                    .unwrap_or_else(|| crate::error::panic_missing_invariant("validated segment"));
                let stats_rows = u64::try_from(parse_i64(
                    required_property(stats, "RowCount")?,
                    "RowCount",
                )?)
                .map_err(|_source| MetadataError::new("negative segment RowCount"))?;
                if records != stats_rows {
                    return Err(MetadataError::new(
                        "segment Records disagrees with ColumnSegmentStats.RowCount",
                    ));
                }
                segment_rows = segment_rows
                    .checked_add(records)
                    .ok_or_else(|| MetadataError::new("segment row-count overflow"))?;
                let compression = segment
                    .member("CompressionInfo")
                    .unwrap_or_else(|| crate::error::panic_missing_invariant("validated segment"))
                    .class
                    .as_str();
                if compression_type == 1 {
                    let width = hybrid_width(compression).ok_or_else(|| {
                        MetadataError::new(
                            "NoSplit CompressionType requires no-split segment compression",
                        )
                    })?;
                    if compression_param != i64::from(width) {
                        return Err(MetadataError::new(
                            "CompressionParam disagrees with the no-split bit width",
                        ));
                    }
                }
            }
            if segment_rows != row_count {
                return Err(MetadataError::new(
                    "column RowCount disagrees with its segment rows",
                ));
            }
            columns.push(ColumnPolicy {
                name: column
                    .name
                    .clone()
                    .ok_or_else(|| MetadataError::new("XMRawColumn requires a name"))?,
                data_file: qualify_storage_name(
                    file.storage_path,
                    partition.name.as_deref().ok_or_else(|| {
                        MetadataError::new("XMRawColumnPartitionDataObject requires a storage name")
                    })?,
                ),
                segment_count,
                row_count,
                compression_type,
                settings: u16::try_from(parse_i64(
                    required_property(column, "Settings")?,
                    "Settings",
                )?)
                .map_err(|_source| MetadataError::new("Settings overflow"))?,
                dictionary,
            });
            let hierarchy = column
                .member("IntrinsicHierarchy")
                .unwrap_or_else(|| crate::error::panic_missing_invariant("validated raw column"));
            let processed =
                parse_bool(required_property(hierarchy, "IsProcessed")?, "IsProcessed")?;
            let materialization = parse_i64(
                required_property(hierarchy, "TypeMaterialization")?,
                "TypeMaterialization",
            )?;
            hierarchies.push(HierarchyPolicy {
                table_store: required_property(hierarchy, "TableStore")?.to_owned(),
                processed,
                position_to_id: processed
                    && parse_i64(
                        required_property(hierarchy, "ColumnPosition2DataID")?,
                        "ColumnPosition2DataID",
                    )? == 0,
                id_to_position: processed
                    && parse_i64(
                        required_property(hierarchy, "ColumnDataID2Position")?,
                        "ColumnDataID2Position",
                    )? == 1
                    && materialization == 0,
                id_to_position_hash: processed
                    && parse_i64(
                        required_property(hierarchy, "ColumnDataID2Position")?,
                        "ColumnDataID2Position",
                    )? == 1
                    && materialization == 1,
                level_ids: Vec::new(),
                level_offsets: Vec::new(),
            });
        }
        for relationship in file
            .table
            .collection("Relationships")
            .unwrap_or_else(|| crate::error::panic_missing_invariant("validated table"))
        {
            let index = relationship.data_objects[0].object.class.as_str();
            relationships.push(RelationshipPolicy {
                name: relationship.name.clone(),
                primary_table: required_property(relationship, "PrimaryTable")?.to_owned(),
                primary_column: required_property(relationship, "PrimaryColumn")?.to_owned(),
                foreign_column: required_property(relationship, "ForeignColumn")?.to_owned(),
                index_kind: match index {
                    "XMRelationshipIndexSparseDIDs" => RelationshipIndexKind::Sparse,
                    "XMRelationshipIndexDenseDIDs" => RelationshipIndexKind::Dense,
                    _ => RelationshipIndexKind::OneTwoThree,
                },
            });
        }
        for hierarchy in file
            .table
            .collection("UserHierarchies")
            .unwrap_or_else(|| crate::error::panic_missing_invariant("validated table"))
        {
            let processed =
                parse_bool(required_property(hierarchy, "IsProcessed")?, "IsProcessed")?;
            let (level_ids, level_offsets) =
                parse_user_hierarchy_store(required_property(hierarchy, "TableStore")?)?;
            hierarchies.push(HierarchyPolicy {
                table_store: required_property(hierarchy, "TableName")?.to_owned(),
                processed,
                position_to_id: false,
                id_to_position: false,
                id_to_position_hash: false,
                level_ids,
                level_offsets,
            });
        }
    }
    Ok((columns, relationships, hierarchies))
}

pub(crate) fn validate_columns(
    metadata: &MetadataModel<'_>,
    native: &[NativeFile<'_>],
) -> MetadataResult<()> {
    for column in &metadata.columns {
        let data = unique_native(native, &column.data_file)?;
        let NativeData::Idf(idf) = &data.data else {
            return Err(MetadataError::new(format!(
                "{} is not column IDF data",
                column.data_file
            )));
        };
        if idf.segments.len()
            != column
                .segment_count
                .checked_mul(2)
                .ok_or_else(|| MetadataError::new("SegmentCount overflow"))?
        {
            return Err(MetadataError::new(format!(
                "column {} SegmentCount disagrees with its IDF",
                column.name
            )));
        }
        if let Some(policy) = &column.dictionary {
            let file = unique_native(native, &policy.storage_name)?;
            let NativeData::Dictionary(dictionary) = &file.data else {
                return Err(MetadataError::new(format!(
                    "{} is not a dictionary",
                    policy.storage_name
                )));
            };
            match (
                policy.class.as_str(),
                dictionary.dictionary_type,
                &dictionary.body,
            ) {
                (
                    "XMHashDataDictionary<XM_Long>",
                    DictionaryType::Long,
                    DictionaryBody::Numeric(values),
                ) => {
                    let expected = if policy.operating_on_32 == Some(true) {
                        4
                    } else {
                        8
                    };
                    if values.element_size != expected {
                        return Err(MetadataError::new(
                            "OperatingOn32 disagrees with numeric dictionary width",
                        ));
                    }
                },
                (
                    "XMHashDataDictionary<XM_Real>",
                    DictionaryType::Real,
                    DictionaryBody::Numeric(values),
                ) if values.element_size == 8 => {},
                (
                    "XMHashDataDictionary<XM_String>",
                    DictionaryType::String,
                    DictionaryBody::String(strings),
                ) => {
                    let flags = policy.dictionary_flags.unwrap_or_else(|| {
                        crate::error::panic_missing_invariant("string policy flags")
                    });
                    if dictionary.hash.is_some() != (flags & 0x001 != 0) {
                        return Err(MetadataError::new(
                            "DictionaryFlags lookup bit disagrees with string hash header presence",
                        ));
                    }
                    let compressed = flags & 0x002 != 0;
                    if strings.pages.iter().any(|page| {
                        matches!(page.data, StringPageData::Compressed { .. }) != compressed
                    }) {
                        return Err(MetadataError::new(
                            "DictionaryFlags compression bit disagrees with string page layout",
                        ));
                    }
                },
                _ => {
                    return Err(MetadataError::new(
                        "dictionary metadata class disagrees with native dictionary type",
                    ));
                },
            }
        }
    }
    Ok(())
}

pub(crate) fn validate_hierarchies(
    metadata: &MetadataModel<'_>,
    native: &[NativeFile<'_>],
    generated: &[SystemGeneratedFile<'_>],
) -> MetadataResult<()> {
    for hierarchy in &metadata.hierarchies {
        if hierarchy.table_store.starts_with("U$") {
            let count = generated
                .iter()
                .filter(|file| {
                    file.object_key.contains(&hierarchy.table_store)
                        && matches!(
                            file.kind,
                            SystemGeneratedKind::UserHierarchyChildCount
                                | SystemGeneratedKind::UserHierarchyFirstChildPosition
                                | SystemGeneratedKind::UserHierarchyMultilevelIdentifier
                                | SystemGeneratedKind::UserHierarchyParentPosition
                        )
                })
                .count();
            let expected = if hierarchy.processed { 4 } else { 0 };
            if count != expected {
                return Err(MetadataError::new(format!(
                    "user hierarchy {} generated-file presence disagrees with IsProcessed",
                    hierarchy.table_store
                )));
            }
            continue;
        }
        require_generated_presence(
            generated,
            &hierarchy.table_store,
            SystemGeneratedKind::PositionToIdentifier,
            hierarchy.position_to_id,
        )?;
        require_generated_presence(
            generated,
            &hierarchy.table_store,
            SystemGeneratedKind::IdentifierToPosition,
            hierarchy.id_to_position,
        )?;
        let hash_count = native
            .iter()
            .filter(|file| {
                basename(file.storage_path).contains(&hierarchy.table_store)
                    && matches!(file.data, NativeData::HashIndex(_))
            })
            .count();
        if hash_count != usize::from(hierarchy.id_to_position_hash) {
            return Err(MetadataError::new(format!(
                "hierarchy {} hash-index presence disagrees with TypeMaterialization",
                hierarchy.table_store
            )));
        }
    }
    Ok(())
}

fn require_generated_presence(
    files: &[SystemGeneratedFile<'_>],
    key: &str,
    kind: SystemGeneratedKind,
    expected: bool,
) -> MetadataResult<()> {
    let count = files
        .iter()
        .filter(|file| file.kind == kind && file.object_key.contains(key))
        .count();
    if count != usize::from(expected) {
        return Err(MetadataError::new(format!(
            "hierarchy {key} generated-file presence is inconsistent"
        )));
    }
    Ok(())
}

pub(crate) fn validate_relationships(
    metadata: &MetadataModel<'_>,
    generated: &[SystemGeneratedFile<'_>],
) -> MetadataResult<()> {
    let indexes: Vec<_> = generated
        .iter()
        .filter(|file| file.kind == SystemGeneratedKind::RelationshipIndex)
        .collect();
    if indexes.len() != metadata.relationships.len() {
        return Err(MetadataError::new(
            "relationship definition count disagrees with generated indexes",
        ));
    }
    let mut remaining: Vec<_> = metadata
        .relationships
        .iter()
        .map(|policy| policy.index_kind)
        .collect();
    for file in indexes {
        let actual = match file.data {
            SystemGeneratedData::HashIndex(_) => RelationshipIndexKind::Sparse,
            SystemGeneratedData::Idf(_) => RelationshipIndexKind::Dense,
        };
        let position = remaining
            .iter()
            .position(|kind| {
                *kind == actual
                    || (*kind == RelationshipIndexKind::OneTwoThree
                        && actual == RelationshipIndexKind::Dense)
            })
            .ok_or_else(|| {
                MetadataError::new("relationship index representation disagrees with metadata")
            })?;
        remaining.remove(position);
    }
    Ok(())
}

fn unique_native<'a>(
    files: &'a [NativeFile<'a>],
    name: &str,
) -> MetadataResult<&'a NativeFile<'a>> {
    let mut matches = files
        .iter()
        .filter(|file| file.storage_path == name || basename(file.storage_path) == name);
    let file = matches.next().ok_or_else(|| {
        MetadataError::new(format!("metadata-bound native file {name} is absent"))
    })?;
    if matches.next().is_some() {
        return Err(MetadataError::new(format!(
            "metadata-bound native file {name} is ambiguous"
        )));
    }
    Ok(file)
}

fn basename(path: &str) -> &str {
    path.rsplit('/').next().unwrap_or(path)
}
fn qualify_storage_name(metadata_path: &str, name: &str) -> String {
    if name.contains('/') {
        name.to_owned()
    } else if let Some((parent, _)) = metadata_path.rsplit_once('/') {
        format!("{parent}/{name}")
    } else {
        name.to_owned()
    }
}
fn parse_user_hierarchy_store(value: &str) -> MetadataResult<(Vec<String>, Vec<u64>)> {
    if !value.starts_with('$') || !value.ends_with('$') {
        return Err(MetadataError::new(
            "user hierarchy TableStore MUST start and end with '$'",
        ));
    }
    let parts: Vec<_> = value[1..value.len() - 1].split('$').collect();
    if parts.is_empty() || parts.len() % 2 != 0 {
        return Err(MetadataError::new(
            "user hierarchy TableStore requires ID/offset pairs",
        ));
    }
    let mut ids = Vec::new();
    let mut offsets = Vec::new();
    for pair in parts.as_chunks::<2>().0 {
        if pair[0].is_empty() || ids.iter().any(|id| id == pair[0]) {
            return Err(MetadataError::new(
                "user hierarchy TableStore contains an empty or duplicate level ID",
            ));
        }
        let offset = pair[1]
            .parse::<u64>()
            .map_err(|_source| MetadataError::new("invalid user hierarchy offset"))?;
        if offsets.last().is_some_and(|previous| *previous > offset) {
            return Err(MetadataError::new(
                "user hierarchy offsets MUST be cumulative",
            ));
        }
        ids.push(pair[0].to_owned());
        offsets.push(offset);
    }
    if offsets.first() != Some(&0) {
        return Err(MetadataError::new(
            "the first user hierarchy offset MUST be zero",
        ));
    }
    Ok((ids, offsets))
}
fn required_property<'a>(object: &'a MetadataObject, name: &str) -> MetadataResult<&'a str> {
    object
        .property(name)
        .ok_or_else(|| MetadataError::new(format!("missing property {name}")))
}
fn parse_i64(value: &str, name: &str) -> MetadataResult<i64> {
    value
        .parse()
        .map_err(|_source| MetadataError::new(format!("invalid integer {name}")))
}
fn parse_i32(value: &str, name: &str) -> MetadataResult<i32> {
    value
        .parse()
        .map_err(|_source| MetadataError::new(format!("invalid integer {name}")))
}
fn parse_bool(value: &str, name: &str) -> MetadataResult<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(MetadataError::new(format!("invalid Boolean {name}"))),
    }
}
fn bool_property(name: &str) -> bool {
    matches!(
        name,
        "IsProcessed"
            | "HasNulls"
            | "Nullable"
            | "Unique"
            | "OperatingOn32"
            | "SegmentNeedsResizing"
    )
}
fn double_property(name: &str) -> bool {
    name == "Magnitude"
}
fn string_property(name: &str) -> bool {
    matches!(
        name,
        "Collation"
            | "OrderByColumn"
            | "PrimaryTable"
            | "PrimaryColumn"
            | "ForeignColumn"
            | "TableName"
            | "TableStore"
    )
}
fn xml_error(error: impl fmt::Display) -> MetadataError {
    MetadataError::new(format!("invalid metadata XML: {error}"))
}
