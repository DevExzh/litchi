//! Bounded, inert parsing of MS-XLDM section 2.6 model OLAP XML files.
//!
//! MS-SSAS base object content is retained as a generic tree. XLDM's `Load`
//! wrapper, tabular extension suffixes, information files, filenames, file
//! lists, and identifier relationships are validated without evaluating MDX,
//! accessing data sources, decrypting, or decompressing native data.

use std::collections::HashSet;
use std::error::Error;
use std::fmt;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

use super::metadata::{MetadataFileKind, MetadataModel};
use super::{GeneratedNameKind, Storage, classify_generated_path};

const MAX_OLAP_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_OLAP_NODES: usize = 750_000;
const MAX_OLAP_DEPTH: usize = 192;
const MAX_OLAP_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_OLAP_FILES: usize = 65_536;
const MAX_FILE_LIST_ITEMS: usize = 100_000;

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapError(String);

impl OlapError {
    fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}
impl fmt::Display for OlapError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}
impl Error for OlapError {}
pub type OlapResult<T> = Result<T, OlapError>;

/// Generic retained MS-SSAS base content. Namespace declarations and
/// attributes are inert strings and never drive external behavior.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapElement {
    pub name: String,
    pub attributes: Vec<(String, String)>,
    pub text: String,
    pub children: Vec<OlapElement>,
}

impl OlapElement {
    #[must_use]
    pub fn child(&self, name: &str) -> Option<&OlapElement> {
        self.children.iter().find(|child| child.name == name)
    }
    pub fn children_named<'a>(&'a self, name: &str) -> impl Iterator<Item = &'a OlapElement> + 'a {
        let name = name.to_owned();
        self.children.iter().filter(move |child| child.name == name)
    }
    #[must_use]
    pub fn scalar(&self, name: &str) -> Option<&str> {
        let child = self.child(name)?;
        child.children.is_empty().then_some(child.text.trim())
    }
}

#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum OlapObjectKind {
    Cube,
    Database,
    DataSource,
    DataSourceView,
    Dimension,
    MdxScript,
    MeasureGroup,
    Partition,
}

impl OlapObjectKind {
    fn element_name(self) -> &'static str {
        match self {
            Self::Cube => "Cube",
            Self::Database => "Database",
            Self::DataSource => "DataSource",
            Self::DataSourceView => "DataSourceView",
            Self::Dimension => "Dimension",
            Self::MdxScript => "MdxScript",
            Self::MeasureGroup => "MeasureGroup",
            Self::Partition => "Partition",
        }
    }
    fn suffix(self) -> &'static str {
        match self {
            Self::Cube => ".cub.xml",
            Self::Database => ".db.xml",
            Self::DataSource => ".ds.xml",
            Self::DataSourceView => ".dsv.xml",
            Self::Dimension => ".dim.xml",
            Self::MdxScript => ".scr.xml",
            Self::MeasureGroup => ".det.xml",
            Self::Partition => ".prt.xml",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum OlapFileKind {
    Definition(OlapObjectKind),
    PartitionInformation,
    DimensionInformation,
    CubeInformation,
}

#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct OlapParentReference {
    pub database_id: Option<String>,
    pub cube_id: Option<String>,
}

/// The section 2.6.1.3 suffix plus object-specific file lists.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TabularExtension {
    pub ordinal: i32,
    pub object_version: i32,
    pub persist_location: i32,
    pub data_files: Vec<String>,
    pub permission_files: Vec<String>,
    pub measure_group_files: Vec<String>,
    pub perspective_files: Vec<String>,
    pub assembly_files: Vec<String>,
    pub aggregation_design_files: Vec<String>,
    pub partition_files: Vec<String>,
    pub default_collation_version: Option<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapHierarchy {
    pub id: String,
    pub level_ids: Vec<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapDefinition {
    pub parent: OlapParentReference,
    pub kind: OlapObjectKind,
    pub object_id: String,
    pub object_name: Option<String>,
    pub object: OlapElement,
    pub extension: TabularExtension,
    pub attribute_ids: Vec<String>,
    pub hierarchies: Vec<OlapHierarchy>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct PartitionInformation {
    pub data_version: i32,
    pub rigid_agg_version: i32,
    pub flex_agg_version: i32,
    pub data_index_version: i32,
    pub rigid_index_version: i32,
    pub flex_index_version: i32,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DimensionInformationMap {
    pub offset_header: i64,
    pub offset_data: i64,
    pub record_count: i64,
    pub segment_count: i64,
    pub format_mask: i64,
    pub header_size: i64,
    pub path_count: i64,
    pub data_count: i64,
    pub segment_index_count: i64,
    pub map_data_indices: i64,
    pub min_max_values: i64,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DimensionInformationProperty {
    pub parent_child: bool,
    pub depth: i32,
    pub balanced: bool,
    pub has_holes: bool,
    pub map_dataset: DimensionInformationMap,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DimensionInformation {
    pub data_version: i32,
    pub index_version: i32,
    pub decode_store_version: i32,
    pub level_store_version: i32,
    pub properties: Vec<DimensionInformationProperty>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct CubeInformation;

#[derive(Clone, Debug, Eq, PartialEq)]
#[allow(
    clippy::large_enum_variant,
    reason = "boxing this public schema enum would break callers"
)]
pub enum OlapDocument {
    Definition(OlapDefinition),
    PartitionInformation(PartitionInformation),
    DimensionInformation(DimensionInformation),
    CubeInformation(CubeInformation),
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapFile<'a> {
    pub storage_path: &'a str,
    pub bytes: &'a [u8],
    pub kind: OlapFileKind,
    pub document: OlapDocument,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapModel<'a> {
    pub files: Vec<OlapFile<'a>>,
}

/// Parse one path if it belongs to section 2.6; unrelated generated XML files
/// return `Ok(None)`.
pub fn parse_file<'a>(storage_path: &'a str, bytes: &'a [u8]) -> OlapResult<Option<OlapFile<'a>>> {
    let generated = classify_generated_path(storage_path).map_err(|error| {
        OlapError::new(format!(
            "invalid generated OLAP path {storage_path}: {error}"
        ))
    })?;
    let kind = match generated.kind {
        GeneratedNameKind::DatabaseDefinition => OlapFileKind::Definition(OlapObjectKind::Database),
        GeneratedNameKind::DataSourceViewDefinition => {
            OlapFileKind::Definition(OlapObjectKind::DataSourceView)
        },
        GeneratedNameKind::CubeDefinition => OlapFileKind::Definition(OlapObjectKind::Cube),
        GeneratedNameKind::DataSourceOrDimensionDefinition if storage_path.ends_with(".ds.xml") => {
            OlapFileKind::Definition(OlapObjectKind::DataSource)
        },
        GeneratedNameKind::DataSourceOrDimensionDefinition
            if storage_path.ends_with(".dim.xml") =>
        {
            OlapFileKind::Definition(OlapObjectKind::Dimension)
        },
        GeneratedNameKind::MdxScriptMetadata => OlapFileKind::Definition(OlapObjectKind::MdxScript),
        GeneratedNameKind::MeasureGroupMetadata => {
            OlapFileKind::Definition(OlapObjectKind::MeasureGroup)
        },
        GeneratedNameKind::PartitionMetadata => OlapFileKind::Definition(OlapObjectKind::Partition),
        GeneratedNameKind::PartitionInformation => OlapFileKind::PartitionInformation,
        GeneratedNameKind::TableInformation => OlapFileKind::DimensionInformation,
        GeneratedNameKind::CubeInformation => OlapFileKind::CubeInformation,
        GeneratedNameKind::DataSourceOrDimensionDefinition
        | GeneratedNameKind::TableMetadata
        | GeneratedNameKind::TableRelationshipMetadata
        | GeneratedNameKind::ColumnHierarchyMetadata
        | GeneratedNameKind::UserHierarchyMetadata
        | GeneratedNameKind::ColumnData
        | GeneratedNameKind::TableRelationshipIndex
        | GeneratedNameKind::ColumnPositionToId
        | GeneratedNameKind::ColumnIdToPosition
        | GeneratedNameKind::ColumnHashIndex
        | GeneratedNameKind::ColumnDictionary
        | GeneratedNameKind::UserHierarchyChildCount
        | GeneratedNameKind::UserHierarchyFirstChildPosition
        | GeneratedNameKind::UserHierarchyParentPosition
        | GeneratedNameKind::UserHierarchyMultilevelId => return Ok(None),
    };
    let root = parse_xml(bytes)?;
    let document = match kind {
        OlapFileKind::Definition(expected) => {
            OlapDocument::Definition(parse_definition(root, expected, storage_path)?)
        },
        OlapFileKind::PartitionInformation => {
            OlapDocument::PartitionInformation(parse_partition_information(root)?)
        },
        OlapFileKind::DimensionInformation => {
            OlapDocument::DimensionInformation(parse_dimension_information(root)?)
        },
        OlapFileKind::CubeInformation => {
            parse_empty_root(&root, "Cube")?;
            OlapDocument::CubeInformation(CubeInformation)
        },
    };
    Ok(Some(OlapFile {
        storage_path,
        bytes,
        kind,
        document,
    }))
}

/// Discover section 2.6 members through the validated backup log and
/// cross-check all file-list paths against the virtual directory.
pub fn inspect<'a>(
    storage: &'a Storage<'a>,
    metadata: &MetadataModel<'_>,
) -> OlapResult<OlapModel<'a>> {
    let mut files = Vec::new();
    let mut seen = HashSet::new();
    for group in &storage.backup_log.file_groups {
        for logged in &group.files {
            let path = logged.storage_path.as_str();
            let generated = classify_generated_path(path).map_err(|error| {
                OlapError::new(format!("invalid generated path {path}: {error}"))
            })?;
            if !is_olap_kind(generated.kind) {
                continue;
            }
            if !seen.insert(path) {
                return Err(OlapError::new(format!(
                    "duplicate logged OLAP member {path}"
                )));
            }
            if files.len() == MAX_OLAP_FILES {
                return Err(OlapError::new("too many OLAP files"));
            }
            let index = storage
                .files
                .iter()
                .position(|entry| entry.path == path)
                .ok_or_else(|| {
                    OlapError::new(format!(
                        "logged OLAP member {path} is absent from the directory"
                    ))
                })?;
            let bytes = storage
                .file_payload(index)
                .ok_or_else(|| OlapError::new(format!("cannot resolve OLAP member {path}")))?;
            if let Some(file) = parse_file(path, bytes)? {
                files.push(file);
            }
        }
    }
    let model = OlapModel { files };
    let paths: Vec<_> = storage
        .files
        .iter()
        .map(|entry| entry.path.as_str())
        .collect();
    validate_model(&model, metadata, Some(&paths))?;
    Ok(model)
}

pub fn validate(model: &OlapModel<'_>, metadata: &MetadataModel<'_>) -> OlapResult<()> {
    validate_model(model, metadata, None)
}

pub fn write_file(file: &OlapFile<'_>) -> OlapResult<Vec<u8>> {
    let reparsed = parse_file(file.storage_path, file.bytes)?
        .ok_or_else(|| OlapError::new("file is not section 2.6 OLAP XML"))?;
    if reparsed.kind != file.kind || reparsed.document != file.document {
        return Err(OlapError::new("OLAP metadata model was mutated"));
    }
    Ok(file.bytes.to_vec())
}

fn is_olap_kind(kind: GeneratedNameKind) -> bool {
    matches!(
        kind,
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
    )
}

fn parse_definition(
    root: OlapElement,
    expected: OlapObjectKind,
    path: &str,
) -> OlapResult<OlapDefinition> {
    if root.name != "Load" || !root.text.trim().is_empty() || root.children.len() != 2 {
        return Err(OlapError::new(
            "OLAP definition requires a Load root with two children",
        ));
    }
    let parent_node = unique_child(&root, "ParentObject")?;
    let definition_node = unique_child(&root, "ObjectDefinition")?;
    require_only_children(parent_node, &["DatabaseID", "CubeID"])?;
    let parent = OlapParentReference {
        database_id: optional_scalar(parent_node, "DatabaseID")?,
        cube_id: optional_scalar(parent_node, "CubeID")?,
    };
    if definition_node.children.len() != 1 || !definition_node.text.trim().is_empty() {
        return Err(OlapError::new(
            "ObjectDefinition requires exactly one major object",
        ));
    }
    let object = definition_node.children[0].clone();
    if object.name != expected.element_name() {
        return Err(OlapError::new(format!(
            "{path} contains the wrong major object {}",
            object.name
        )));
    }
    let extension = parse_tabular_extension(&object, expected)?;
    let file_version = generated_version(path, expected.suffix())?;
    if extension.object_version != file_version {
        return Err(OlapError::new(
            "ObjectVersion disagrees with the generated filename",
        ));
    }
    let object_id = required_scalar(&object, "ID")?.to_owned();
    if object_id.is_empty() {
        return Err(OlapError::new("OLAP object ID cannot be empty"));
    }
    let object_name = object.scalar("Name").map(str::to_owned);
    let attribute_ids = if expected == OlapObjectKind::Dimension {
        collect_collection_ids(&object, "Attributes", "Attribute", "ID")?
    } else {
        Vec::new()
    };
    let hierarchies = if expected == OlapObjectKind::Dimension {
        collect_hierarchies(&object)?
    } else {
        Vec::new()
    };
    validate_parent_shape(expected, &parent)?;
    Ok(OlapDefinition {
        parent,
        kind: expected,
        object_id,
        object_name,
        object,
        extension,
        attribute_ids,
        hierarchies,
    })
}

fn parse_tabular_extension(
    object: &OlapElement,
    kind: OlapObjectKind,
) -> OlapResult<TabularExtension> {
    let extras: &[&str] = match kind {
        OlapObjectKind::DataSource => &["PermissionFileList"],
        OlapObjectKind::Cube => &[
            "PermissionFileList",
            "MeasureGroupFileList",
            "PerspectiveFileList",
            "AssemblyFileList",
        ],
        OlapObjectKind::Dimension => &["PermissionFileList"],
        OlapObjectKind::MeasureGroup => &["AggregationDesignFileList", "PartitionFileList"],
        OlapObjectKind::Database
        | OlapObjectKind::DataSourceView
        | OlapObjectKind::MdxScript
        | OlapObjectKind::Partition => &[],
    };
    let suffix_len = 5 + extras.len();
    if object.children.len() < suffix_len {
        return Err(OlapError::new(format!(
            "{} is missing its tabular suffix",
            object.name
        )));
    }
    let suffix = &object.children[object.children.len() - suffix_len..];
    let names = [
        "Ordinal",
        "ObjectVersion",
        "PersistLocation",
        "System",
        "DataFileList",
    ];
    for (node, expected) in suffix.iter().zip(names) {
        if node.name != expected {
            return Err(OlapError::new(format!(
                "tabular suffix requires {expected} in sequence"
            )));
        }
    }
    for (node, expected) in suffix[5..].iter().zip(extras) {
        if node.name != *expected {
            return Err(OlapError::new(format!(
                "tabular suffix requires {expected} in sequence"
            )));
        }
    }
    let system = parse_bool(scalar_node(&suffix[3])?, "System")?;
    if system {
        return Err(OlapError::new("tabular System MUST be false"));
    }
    let mut extension = TabularExtension {
        ordinal: parse_i32(scalar_node(&suffix[0])?, "Ordinal")?,
        object_version: parse_i32(scalar_node(&suffix[1])?, "ObjectVersion")?,
        persist_location: parse_i32(scalar_node(&suffix[2])?, "PersistLocation")?,
        data_files: parse_file_list(scalar_node(&suffix[4])?)?,
        permission_files: Vec::new(),
        measure_group_files: Vec::new(),
        perspective_files: Vec::new(),
        assembly_files: Vec::new(),
        aggregation_design_files: Vec::new(),
        partition_files: Vec::new(),
        default_collation_version: None,
    };
    if extension.ordinal < 0 || extension.persist_location < 0 {
        return Err(OlapError::new(
            "Ordinal and PersistLocation MUST be nonnegative",
        ));
    }
    for node in &suffix[5..] {
        let files = parse_file_list(scalar_node(node)?)?;
        match node.name.as_str() {
            "PermissionFileList" => extension.permission_files = files,
            "MeasureGroupFileList" => extension.measure_group_files = files,
            "PerspectiveFileList" => extension.perspective_files = files,
            "AssemblyFileList" => extension.assembly_files = files,
            "AggregationDesignFileList" => extension.aggregation_design_files = files,
            "PartitionFileList" => extension.partition_files = files,
            _ => unreachable!(),
        }
    }
    if kind == OlapObjectKind::Database {
        let prefix = &object.children[..object.children.len() - suffix_len];
        if let Some(node) = prefix
            .last()
            .filter(|node| node.name == "DefaultCollationVersion")
        {
            let value = scalar_node(node)?.to_owned();
            if !matches!(value.as_str(), "Earliest" | "80" | "90" | "100") {
                return Err(OlapError::new("invalid DefaultCollationVersion"));
            }
            extension.default_collation_version = Some(value);
        }
    }
    Ok(extension)
}

fn parse_partition_information(root: OlapElement) -> OlapResult<PartitionInformation> {
    require_sequence(
        &root,
        "Partition",
        &[
            "DataVersion",
            "RigidAggVersion",
            "FlexAggVersion",
            "DataIndexVersion",
            "RigidIndexVersion",
            "FlexIndexVersion",
        ],
    )?;
    Ok(PartitionInformation {
        data_version: child_i32(&root, "DataVersion")?,
        rigid_agg_version: child_i32(&root, "RigidAggVersion")?,
        flex_agg_version: child_i32(&root, "FlexAggVersion")?,
        data_index_version: child_i32(&root, "DataIndexVersion")?,
        rigid_index_version: child_i32(&root, "RigidIndexVersion")?,
        flex_index_version: child_i32(&root, "FlexIndexVersion")?,
    })
}

fn parse_dimension_information(root: OlapElement) -> OlapResult<DimensionInformation> {
    require_sequence(
        &root,
        "Dimension",
        &[
            "DataVersion",
            "IndexVersion",
            "DecodeStoreVersion",
            "LevelStoreVersion",
            "Properties",
        ],
    )?;
    let properties_node = root
        .child("Properties")
        .unwrap_or_else(|| crate::error::panic_missing_invariant("validated sequence"));
    if !properties_node.text.trim().is_empty()
        || properties_node
            .children
            .iter()
            .any(|child| child.name != "Property")
    {
        return Err(OlapError::new("invalid dimension information Properties"));
    }
    let mut properties = Vec::new();
    for property in &properties_node.children {
        require_sequence(
            property,
            "Property",
            &["ParentChild", "Depth", "Balanced", "HasHoles", "MapDataset"],
        )?;
        let map = property
            .child("MapDataset")
            .unwrap_or_else(|| crate::error::panic_missing_invariant("validated sequence"));
        let names = [
            "m_cbOffsetHeader",
            "m_cbOffsetData",
            "m_cRecord",
            "m_cSegment",
            "m_mskFormat",
            "m_cbHeader",
            "m_cPath",
            "m_cData",
            "m_cSegmentIndex",
            "MapDataIndices",
            "MinMaxValues",
        ];
        require_sequence(map, "MapDataset", &names)?;
        let value = |name| child_i64(map, name);
        properties.push(DimensionInformationProperty {
            parent_child: child_bool(property, "ParentChild")?,
            depth: child_i32(property, "Depth")?,
            balanced: child_bool(property, "Balanced")?,
            has_holes: child_bool(property, "HasHoles")?,
            map_dataset: DimensionInformationMap {
                offset_header: value("m_cbOffsetHeader")?,
                offset_data: value("m_cbOffsetData")?,
                record_count: value("m_cRecord")?,
                segment_count: value("m_cSegment")?,
                format_mask: value("m_mskFormat")?,
                header_size: value("m_cbHeader")?,
                path_count: value("m_cPath")?,
                data_count: value("m_cData")?,
                segment_index_count: value("m_cSegmentIndex")?,
                map_data_indices: value("MapDataIndices")?,
                min_max_values: value("MinMaxValues")?,
            },
        });
    }
    Ok(DimensionInformation {
        data_version: child_i32(&root, "DataVersion")?,
        index_version: child_i32(&root, "IndexVersion")?,
        decode_store_version: child_i32(&root, "DecodeStoreVersion")?,
        level_store_version: child_i32(&root, "LevelStoreVersion")?,
        properties,
    })
}

fn parse_empty_root(root: &OlapElement, name: &str) -> OlapResult<()> {
    if root.name != name || !root.children.is_empty() || !root.text.trim().is_empty() {
        return Err(OlapError::new(format!("{name} information MUST be empty")));
    }
    Ok(())
}

fn validate_model(
    model: &OlapModel<'_>,
    metadata: &MetadataModel<'_>,
    storage_paths: Option<&[&str]>,
) -> OlapResult<()> {
    let definitions: Vec<_> = model
        .files
        .iter()
        .filter_map(|file| match &file.document {
            OlapDocument::Definition(value) => Some((file.storage_path, value)),
            OlapDocument::PartitionInformation(_)
            | OlapDocument::DimensionInformation(_)
            | OlapDocument::CubeInformation(_) => None,
        })
        .collect();
    let mut ids = HashSet::new();
    let mut ordinals = HashSet::new();
    for (_, definition) in &definitions {
        if !ids.insert((definition.kind, definition.object_id.as_str())) {
            return Err(OlapError::new(format!(
                "duplicate {} object ID {}",
                definition.kind.element_name(),
                definition.object_id
            )));
        }
    }
    let count = |kind| {
        definitions
            .iter()
            .filter(|(_, value)| value.kind == kind)
            .count()
    };
    for kind in [
        OlapObjectKind::Database,
        OlapObjectKind::DataSource,
        OlapObjectKind::DataSourceView,
        OlapObjectKind::Cube,
        OlapObjectKind::Dimension,
        OlapObjectKind::MdxScript,
        OlapObjectKind::MeasureGroup,
        OlapObjectKind::Partition,
    ] {
        let actual = count(kind);
        if actual == 0 {
            return Err(OlapError::new(format!(
                "model requires a {} definition",
                kind.element_name()
            )));
        }
        if matches!(
            kind,
            OlapObjectKind::Database
                | OlapObjectKind::DataSource
                | OlapObjectKind::DataSourceView
                | OlapObjectKind::Cube
                | OlapObjectKind::MdxScript
        ) && actual != 1
        {
            return Err(OlapError::new(format!(
                "model requires exactly one {} definition",
                kind.element_name()
            )));
        }
    }
    let database_ids: HashSet<_> = definitions
        .iter()
        .filter(|(_, value)| value.kind == OlapObjectKind::Database)
        .map(|(_, value)| value.object_id.as_str())
        .collect();
    let cube_ids: HashSet<_> = definitions
        .iter()
        .filter(|(_, value)| value.kind == OlapObjectKind::Cube)
        .map(|(_, value)| value.object_id.as_str())
        .collect();
    for (path, definition) in &definitions {
        let ordinal_key = (
            definition.kind,
            definition.parent.database_id.as_deref(),
            definition.parent.cube_id.as_deref(),
            definition.extension.ordinal,
        );
        if !ordinals.insert(ordinal_key) {
            return Err(OlapError::new(format!(
                "duplicate {} Ordinal {} within one parent",
                definition.kind.element_name(),
                definition.extension.ordinal
            )));
        }
        if definition
            .parent
            .database_id
            .as_deref()
            .is_some_and(|id| !database_ids.contains(id))
        {
            return Err(OlapError::new(format!(
                "{path} references an unknown DatabaseID"
            )));
        }
        if definition
            .parent
            .cube_id
            .as_deref()
            .is_some_and(|id| !cube_ids.contains(id))
        {
            return Err(OlapError::new(format!(
                "{path} references an unknown CubeID"
            )));
        }
        if let Some(paths) = storage_paths {
            let folder = persist_folder(path, definition)?;
            for item in all_file_lists(&definition.extension) {
                require_storage_reference(paths, &folder, item)?;
            }
        }
    }
    let table_count = metadata
        .files
        .iter()
        .filter(|file| file.kind == MetadataFileKind::Table)
        .count();
    for kind in [
        OlapObjectKind::Dimension,
        OlapObjectKind::MeasureGroup,
        OlapObjectKind::Partition,
    ] {
        if count(kind) < table_count {
            return Err(OlapError::new(format!(
                "every table requires a {} definition",
                kind.element_name()
            )));
        }
    }
    let dimension_attributes: HashSet<_> = definitions
        .iter()
        .filter(|(_, value)| value.kind == OlapObjectKind::Dimension)
        .flat_map(|(_, value)| value.attribute_ids.iter().map(String::as_str))
        .collect();
    for file in metadata
        .files
        .iter()
        .filter(|file| file.kind == MetadataFileKind::Table)
    {
        for column in file
            .table
            .collection("Columns")
            .unwrap_or_else(|| crate::error::panic_missing_invariant("validated metadata"))
        {
            let id = column
                .name
                .as_deref()
                .ok_or_else(|| OlapError::new("metadata column has no ID"))?;
            if !dimension_attributes.contains(id) {
                return Err(OlapError::new(format!(
                    "column ID {id} has no dimension Attribute"
                )));
            }
        }
    }
    let olap_hierarchies: Vec<_> = definitions
        .iter()
        .flat_map(|(_, value)| value.hierarchies.iter())
        .collect();
    for policy in metadata
        .hierarchies
        .iter()
        .filter(|policy| !policy.level_ids.is_empty())
    {
        if !olap_hierarchies
            .iter()
            .any(|hierarchy| hierarchy.level_ids == policy.level_ids)
        {
            return Err(OlapError::new(format!(
                "user hierarchy {} level IDs do not match a dimension Hierarchy",
                policy.table_store
            )));
        }
        if policy.level_offsets.len() != policy.level_ids.len() {
            return Err(OlapError::new(
                "user hierarchy ID/offset cardinality mismatch",
            ));
        }
    }
    let dimension_values: HashSet<_> = definitions
        .iter()
        .filter(|(_, value)| value.kind == OlapObjectKind::Dimension)
        .flat_map(|(_, value)| descendant_scalars(&value.object))
        .collect();
    for relationship in &metadata.relationships {
        for id in [
            &relationship.primary_table,
            &relationship.primary_column,
            &relationship.foreign_column,
        ] {
            if !dimension_values.contains(id.as_str()) {
                return Err(OlapError::new(format!(
                    "relationship identifier {id} is absent from dimension metadata"
                )));
            }
        }
    }
    Ok(())
}

fn validate_parent_shape(kind: OlapObjectKind, parent: &OlapParentReference) -> OlapResult<()> {
    let valid = match kind {
        OlapObjectKind::Database => parent.database_id.is_none() && parent.cube_id.is_none(),
        OlapObjectKind::MdxScript | OlapObjectKind::MeasureGroup | OlapObjectKind::Partition => {
            parent.database_id.is_some() && parent.cube_id.is_some()
        },
        OlapObjectKind::Cube
        | OlapObjectKind::DataSource
        | OlapObjectKind::DataSourceView
        | OlapObjectKind::Dimension => parent.database_id.is_some() && parent.cube_id.is_none(),
    };
    if !valid {
        return Err(OlapError::new(format!(
            "{} has an invalid ParentObject reference",
            kind.element_name()
        )));
    }
    Ok(())
}

fn all_file_lists(extension: &TabularExtension) -> impl Iterator<Item = &str> {
    [
        &extension.data_files,
        &extension.permission_files,
        &extension.measure_group_files,
        &extension.perspective_files,
        &extension.assembly_files,
        &extension.aggregation_design_files,
        &extension.partition_files,
    ]
    .into_iter()
    .flat_map(|items| items.iter().map(String::as_str))
}

fn persist_folder(path: &str, definition: &OlapDefinition) -> OlapResult<String> {
    let parent = path.rsplit_once('/').map_or("", |value| value.0);
    if matches!(
        definition.kind,
        OlapObjectKind::DataSourceView | OlapObjectKind::MdxScript
    ) {
        return Ok(parent.to_owned());
    }
    let base = path
        .rsplit('/')
        .next()
        .unwrap_or(path)
        .strip_suffix(definition.kind.suffix())
        .ok_or_else(|| OlapError::new("definition suffix changed"))?;
    let id = base
        .rsplit_once('.')
        .ok_or_else(|| OlapError::new("definition version changed"))?
        .0;
    let object_suffix = definition
        .kind
        .suffix()
        .strip_suffix(".xml")
        .unwrap_or_else(|| crate::error::panic_missing_invariant("fixed suffix"));
    let folder = format!(
        "{id}.{}{object_suffix}",
        definition.extension.persist_location
    );
    Ok(if parent.is_empty() {
        folder
    } else {
        format!("{parent}/{folder}")
    })
}

fn require_storage_reference(paths: &[&str], folder: &str, item: &str) -> OlapResult<()> {
    let qualified = if item.contains('/') || folder.is_empty() {
        item.to_owned()
    } else {
        format!("{folder}/{item}")
    };
    if !paths.contains(&qualified.as_str()) {
        return Err(OlapError::new(format!(
            "OLAP file-list member {item} is absent from storage at {qualified}"
        )));
    }
    Ok(())
}

fn collect_collection_ids(
    object: &OlapElement,
    collection: &str,
    item: &str,
    id: &str,
) -> OlapResult<Vec<String>> {
    let Some(collection) = object.child(collection) else {
        return Ok(Vec::new());
    };
    let mut ids = Vec::new();
    for child in &collection.children {
        if child.name != item {
            continue;
        }
        let value = required_scalar(child, id)?.to_owned();
        if ids.contains(&value) {
            return Err(OlapError::new(format!("duplicate {item} ID {value}")));
        }
        ids.push(value);
    }
    Ok(ids)
}

fn collect_hierarchies(object: &OlapElement) -> OlapResult<Vec<OlapHierarchy>> {
    let Some(collection) = object.child("Hierarchies") else {
        return Ok(Vec::new());
    };
    let mut result = Vec::new();
    for hierarchy in collection.children_named("Hierarchy") {
        let id = required_scalar(hierarchy, "ID")?.to_owned();
        let levels = hierarchy
            .child("Levels")
            .map(|levels| {
                levels
                    .children_named("Level")
                    .map(|level| required_scalar(level, "ID").map(str::to_owned))
                    .collect()
            })
            .transpose()?
            .unwrap_or_default();
        result.push(OlapHierarchy {
            id,
            level_ids: levels,
        });
    }
    Ok(result)
}

fn descendant_scalars(element: &OlapElement) -> Vec<&str> {
    let mut values = Vec::new();
    if element.children.is_empty() && !element.text.trim().is_empty() {
        values.push(element.text.trim());
    }
    for child in &element.children {
        values.extend(descendant_scalars(child));
    }
    values
}

fn parse_file_list(value: &str) -> OlapResult<Vec<String>> {
    if value.is_empty() {
        return Ok(Vec::new());
    }
    let mut result = Vec::new();
    for item in value.split(';') {
        if result.len() == MAX_FILE_LIST_ITEMS {
            return Err(OlapError::new("OLAP file list is too large"));
        }
        if item.is_empty()
            || item.starts_with('/')
            || item.contains('\\')
            || item
                .split('/')
                .any(|part| part.is_empty() || part == "." || part == "..")
        {
            return Err(OlapError::new("unsafe or empty OLAP file-list path"));
        }
        if result.iter().any(|existing| existing == item) {
            return Err(OlapError::new("duplicate OLAP file-list item"));
        }
        result.push(item.to_owned());
    }
    Ok(result)
}

fn generated_version(path: &str, suffix: &str) -> OlapResult<i32> {
    let base = path.rsplit('/').next().unwrap_or(path);
    let prefix = base
        .strip_suffix(suffix)
        .ok_or_else(|| OlapError::new("generated definition suffix changed"))?;
    let version = prefix
        .rsplit_once('.')
        .ok_or_else(|| OlapError::new("generated definition has no version"))?
        .1;
    parse_i32(version, "filename version")
}

fn parse_xml(bytes: &[u8]) -> OlapResult<OlapElement> {
    if bytes.len() > MAX_OLAP_XML_BYTES {
        return Err(OlapError::new("OLAP XML is too large"));
    }
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(ref node) | Event::Empty(ref node) => {
                nodes += 1;
                if nodes > MAX_OLAP_NODES || stack.len() >= MAX_OLAP_DEPTH {
                    return Err(OlapError::new("OLAP XML structure limit exceeded"));
                }
                let empty = matches!(event, Event::Empty(_));
                let node = make_node(node, decoder)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| OlapError::new("unexpected OLAP XML end"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let text = text.decode().map_err(xml_error)?;
                let text = quick_xml::escape::unescape(&text).map_err(xml_error)?;
                text_bytes = text_bytes
                    .checked_add(text.len())
                    .ok_or_else(|| OlapError::new("OLAP XML text overflow"))?;
                if text_bytes > MAX_OLAP_TEXT_BYTES {
                    return Err(OlapError::new("OLAP XML text limit exceeded"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&text);
                } else if !text.trim().is_empty() {
                    return Err(OlapError::new("text outside OLAP root"));
                }
            },
            Event::CData(text) => {
                let text = text.decode().map_err(xml_error)?;
                text_bytes = text_bytes
                    .checked_add(text.len())
                    .ok_or_else(|| OlapError::new("OLAP XML text overflow"))?;
                if text_bytes > MAX_OLAP_TEXT_BYTES {
                    return Err(OlapError::new("OLAP XML text limit exceeded"));
                }
                stack
                    .last_mut()
                    .ok_or_else(|| OlapError::new("CDATA outside OLAP root"))?
                    .text
                    .push_str(&text);
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
                    .ok_or_else(|| OlapError::new("custom XML entities are rejected"))?;
                stack
                    .last_mut()
                    .ok_or_else(|| OlapError::new("entity outside OLAP root"))?
                    .text
                    .push_str(&value);
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(OlapError::new(
                    "DTD and processing instructions are rejected",
                ));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(OlapError::new("unclosed OLAP XML element"));
    }
    root.ok_or_else(|| OlapError::new("missing OLAP XML root"))
}

fn make_node(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> OlapResult<OlapElement> {
    let qualified_name = element.name();
    let raw = std::str::from_utf8(qualified_name.as_ref()).map_err(xml_error)?;
    let name = raw.rsplit(':').next().unwrap_or(raw).to_owned();
    let mut attributes = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if attributes.iter().any(|(existing, _)| existing == &key) {
            return Err(OlapError::new("duplicate OLAP XML attribute"));
        }
        attributes.push((key, value));
    }
    Ok(OlapElement {
        name,
        attributes,
        text: String::new(),
        children: Vec::new(),
    })
}
fn attach(
    node: OlapElement,
    stack: &mut [OlapElement],
    root: &mut Option<OlapElement>,
) -> OlapResult<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(OlapError::new("multiple OLAP XML roots"));
    }
    Ok(())
}

fn unique_child<'a>(node: &'a OlapElement, name: &str) -> OlapResult<&'a OlapElement> {
    let mut values = node.children_named(name);
    let value = values
        .next()
        .ok_or_else(|| OlapError::new(format!("missing {name}")))?;
    if values.next().is_some() {
        return Err(OlapError::new(format!("duplicate {name}")));
    }
    Ok(value)
}
fn required_scalar<'a>(node: &'a OlapElement, name: &str) -> OlapResult<&'a str> {
    let child = unique_child(node, name)?;
    scalar_node(child)
}
fn optional_scalar(node: &OlapElement, name: &str) -> OlapResult<Option<String>> {
    let values: Vec<_> = node.children_named(name).collect();
    if values.len() > 1 {
        return Err(OlapError::new(format!("duplicate {name}")));
    }
    values
        .first()
        .map(|node| scalar_node(node).map(str::to_owned))
        .transpose()
}
fn scalar_node(node: &OlapElement) -> OlapResult<&str> {
    if !node.children.is_empty() {
        return Err(OlapError::new(format!("{} MUST be scalar", node.name)));
    }
    Ok(node.text.trim())
}
fn require_only_children(node: &OlapElement, names: &[&str]) -> OlapResult<()> {
    if !node.text.trim().is_empty()
        || node
            .children
            .iter()
            .any(|child| !names.contains(&child.name.as_str()))
    {
        return Err(OlapError::new(format!(
            "{} contains an unknown child",
            node.name
        )));
    }
    Ok(())
}
fn require_sequence(node: &OlapElement, root: &str, names: &[&str]) -> OlapResult<()> {
    if node.name != root
        || !node.text.trim().is_empty()
        || node.children.len() != names.len()
        || node
            .children
            .iter()
            .zip(names)
            .any(|(child, name)| child.name != *name)
    {
        return Err(OlapError::new(format!(
            "invalid {root} information sequence"
        )));
    }
    Ok(())
}
fn parse_i32(value: &str, name: &str) -> OlapResult<i32> {
    value
        .parse()
        .map_err(|_source| OlapError::new(format!("invalid integer {name}")))
}
fn parse_i64(value: &str, name: &str) -> OlapResult<i64> {
    value
        .parse()
        .map_err(|_source| OlapError::new(format!("invalid integer {name}")))
}
fn parse_bool(value: &str, name: &str) -> OlapResult<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(OlapError::new(format!("invalid Boolean {name}"))),
    }
}
fn child_i32(node: &OlapElement, name: &str) -> OlapResult<i32> {
    parse_i32(required_scalar(node, name)?, name)
}
fn child_i64(node: &OlapElement, name: &str) -> OlapResult<i64> {
    parse_i64(required_scalar(node, name)?, name)
}
fn child_bool(node: &OlapElement, name: &str) -> OlapResult<bool> {
    parse_bool(required_scalar(node, name)?, name)
}
fn xml_error(error: impl fmt::Display) -> OlapError {
    OlapError::new(format!("invalid OLAP XML: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn definition(
        kind: OlapObjectKind,
        id: &str,
        version: i32,
        parent: &str,
        base: &str,
        extras: &str,
    ) -> String {
        format!(
            "<Load><ParentObject>{parent}</ParentObject><ObjectDefinition><{name}><ID>{id}</ID>{base}<Ordinal>0</Ordinal><ObjectVersion>{version}</ObjectVersion><PersistLocation>0</PersistLocation><System>false</System><DataFileList></DataFileList>{extras}</{name}></ObjectDefinition></Load>",
            name = kind.element_name()
        )
    }

    #[test]
    fn parses_every_xldm_defined_document_shape_borrowed_and_exact() {
        let xml = definition(
            OlapObjectKind::Dimension,
            "Table",
            0,
            "<DatabaseID>DB</DatabaseID>",
            "<Attributes><Attribute><ID>Column</ID></Attribute></Attributes><Hierarchies><Hierarchy><ID>Geo</ID><Levels><Level><ID>Country</ID></Level><Level><ID>City</ID></Level></Levels></Hierarchy></Hierarchies>",
            "<PermissionFileList></PermissionFileList>",
        );
        let file = parse_file("Model.1.db/Table-T.0.dim.xml", xml.as_bytes())
            .unwrap()
            .unwrap();
        let OlapDocument::Definition(value) = &file.document else {
            panic!()
        };
        assert_eq!(value.attribute_ids, ["Column"]);
        assert_eq!(value.hierarchies[0].level_ids, ["Country", "City"]);
        assert!(std::ptr::eq(file.bytes.as_ptr(), xml.as_ptr()));
        assert_eq!(write_file(&file).unwrap(), xml.as_bytes());

        let partition = b"<Partition><DataVersion>1</DataVersion><RigidAggVersion>0</RigidAggVersion><FlexAggVersion>0</FlexAggVersion><DataIndexVersion>0</DataIndexVersion><RigidIndexVersion>0</RigidIndexVersion><FlexIndexVersion>0</FlexIndexVersion></Partition>";
        assert!(matches!(
            parse_file(
                "Model.1.db/Cube.0.cub/Table.0.det/Table.1.prt/info.2.xml",
                partition
            )
            .unwrap()
            .unwrap()
            .document,
            OlapDocument::PartitionInformation(_)
        ));
        let cube = b"<Cube/>";
        assert!(matches!(
            parse_file("Model.1.db/Cube.0.cub/info.3.xml", cube)
                .unwrap()
                .unwrap()
                .document,
            OlapDocument::CubeInformation(_)
        ));
    }

    #[test]
    fn parses_dimension_information_and_ignored_map_declarations() {
        let map = "<MapDataset><m_cbOffsetHeader>1</m_cbOffsetHeader><m_cbOffsetData>2</m_cbOffsetData><m_cRecord>3</m_cRecord><m_cSegment>4</m_cSegment><m_mskFormat>5</m_mskFormat><m_cbHeader>6</m_cbHeader><m_cPath>7</m_cPath><m_cData>8</m_cData><m_cSegmentIndex>9</m_cSegmentIndex><MapDataIndices>10</MapDataIndices><MinMaxValues>11</MinMaxValues></MapDataset>";
        let xml = format!(
            "<Dimension><DataVersion>1</DataVersion><IndexVersion>2</IndexVersion><DecodeStoreVersion>3</DecodeStoreVersion><LevelStoreVersion>4</LevelStoreVersion><Properties><Property><ParentChild>false</ParentChild><Depth>0</Depth><Balanced>true</Balanced><HasHoles>false</HasHoles>{map}</Property></Properties></Dimension>"
        );
        let file = parse_file("Model.1.db/Table.0.dim/info.5.xml", xml.as_bytes())
            .unwrap()
            .unwrap();
        let OlapDocument::DimensionInformation(value) = file.document else {
            panic!()
        };
        assert_eq!(value.properties[0].map_dataset.min_max_values, 11);
    }

    #[test]
    fn rejects_versions_system_paths_sequences_and_hostile_xml() {
        let valid = definition(OlapObjectKind::Database, "DB", 1, "", "", "");
        for xml in [
            valid.replace("<ObjectVersion>1", "<ObjectVersion>2"),
            valid.replace("<System>false", "<System>true"),
            valid.replace(
                "<Ordinal>0</Ordinal><ObjectVersion>1</ObjectVersion>",
                "<ObjectVersion>1</ObjectVersion><Ordinal>0</Ordinal>",
            ),
            valid.replace(
                "<DataFileList></DataFileList>",
                "<DataFileList>../secret</DataFileList>",
            ),
        ] {
            assert!(parse_file("Model.1.db.xml", xml.as_bytes()).is_err());
        }
        assert!(
            parse_file(
                "Model.1.db.xml",
                b"<!DOCTYPE x [<!ENTITY a 'x'>]><Load>&a;</Load>"
            )
            .is_err()
        );
    }

    #[test]
    fn validates_hierarchy_level_ids_against_section_2_5() {
        let hierarchy = super::super::metadata::HierarchyPolicy {
            table_store: "U$T$Geo".into(),
            processed: true,
            position_to_id: false,
            id_to_position: false,
            id_to_position_hash: false,
            level_ids: vec!["Country".into(), "City".into()],
            level_offsets: vec![0, 2],
        };
        let metadata = MetadataModel {
            files: vec![],
            columns: vec![],
            relationships: vec![],
            hierarchies: vec![hierarchy],
        };
        let definitions = [
            (OlapObjectKind::Database, "DB", 1, "", "", ""),
            (
                OlapObjectKind::DataSource,
                "DS",
                1,
                "<DatabaseID>DB</DatabaseID>",
                "",
                "<PermissionFileList></PermissionFileList>",
            ),
            (
                OlapObjectKind::DataSourceView,
                "DSV",
                1,
                "<DatabaseID>DB</DatabaseID>",
                "",
                "",
            ),
            (
                OlapObjectKind::Cube,
                "Cube",
                1,
                "<DatabaseID>DB</DatabaseID>",
                "",
                "<PermissionFileList></PermissionFileList><MeasureGroupFileList></MeasureGroupFileList><PerspectiveFileList></PerspectiveFileList><AssemblyFileList></AssemblyFileList>",
            ),
            (
                OlapObjectKind::Dimension,
                "T",
                1,
                "<DatabaseID>DB</DatabaseID>",
                "<Hierarchies><Hierarchy><ID>Geo</ID><Levels><Level><ID>Country</ID></Level><Level><ID>City</ID></Level></Levels></Hierarchy></Hierarchies>",
                "<PermissionFileList></PermissionFileList>",
            ),
            (
                OlapObjectKind::MdxScript,
                "Script",
                0,
                "<DatabaseID>DB</DatabaseID><CubeID>Cube</CubeID>",
                "",
                "",
            ),
            (
                OlapObjectKind::MeasureGroup,
                "T",
                1,
                "<DatabaseID>DB</DatabaseID><CubeID>Cube</CubeID>",
                "",
                "<AggregationDesignFileList></AggregationDesignFileList><PartitionFileList></PartitionFileList>",
            ),
            (
                OlapObjectKind::Partition,
                "T",
                1,
                "<DatabaseID>DB</DatabaseID><CubeID>Cube</CubeID>",
                "",
                "",
            ),
        ];
        let mut owned = Vec::new();
        for (kind, id, version, parent, base, extras) in definitions {
            owned.push((kind, definition(kind, id, version, parent, base, extras)));
        }
        let paths = [
            "Model.1.db.xml",
            "Model.1.db/Source.1.ds.xml",
            "Model.1.db/View.1.dsv.xml",
            "Model.1.db/Cube.1.cub.xml",
            "Model.1.db/Table-T.1.dim.xml",
            "Model.1.db/Cube.0.cub/MdxScript.0.scr.xml",
            "Model.1.db/Cube.0.cub/Table-T.1.det.xml",
            "Model.1.db/Cube.0.cub/Table-T.0.det/Table-T.1.prt.xml",
        ];
        let mut files = Vec::new();
        for ((_, xml), path) in owned.iter().zip(paths) {
            files.push(parse_file(path, xml.as_bytes()).unwrap().unwrap());
        }
        let model = OlapModel { files };
        validate(&model, &metadata).unwrap();
        let mut bad = metadata.clone();
        bad.hierarchies[0].level_ids[1] = "State".into();
        assert!(validate(&model, &bad).is_err());
    }
}
