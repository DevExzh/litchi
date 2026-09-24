//! Bounded source-bound proof of the MS-XLDM 2.6 OLAP projection.
//!
//! MS-XLDM delegates the base OLAP object grammar to MS-SSAS while defining
//! the table-model closure in sections 2.6.1.3 and 2.6.5 through 2.6.8.
//! This owner therefore keeps every persisted identifier in its original
//! namespace. A generated TableID, an inner metadata name, an OLAP ID, an
//! OLAP ObjectID, and a workbook table identifier are not interchangeable.
//! A mapping is admitted only when the containing folder and explicit XML
//! reference fields establish it.
//!
//! This module is an inspection/proof boundary. It does not evaluate MDX or
//! DAX, refresh a data source, resolve a path, decompress a data member, or
//! author XML. The returned proof owns only bounded descriptor strings and
//! borrows the already inspected XLDM source; member payloads stay source
//! backed.

use std::collections::{HashMap, HashSet};
use std::error::Error as StdError;
use std::fmt;

use super::metadata::{MetadataFileKind, MetadataModel, MetadataObject};
use super::model::{FileGroupClass, GeneratedNameKind, Storage, StorageProfile};
use super::olap::{
    OlapDefinition, OlapDocument, OlapElement, OlapFileKind, OlapModel, OlapObjectKind,
};

const MAX_DEFAULT_ITEMS: usize = 1_000_000;
const MAX_DEFAULT_STRING_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEFAULT_SOURCE_BYTES: usize = 512 * 1024 * 1024;
const MAX_DEFAULT_WORK: usize = 4_000_000;
// Proof bindings retain the same source descriptors in more than one owner
// (FileGroup, table, relationship, and cube edges). Charge a conservative
// number of retained copies before the first proof-owned index is reserved.
const RETAINED_STRING_COPIES: usize = 4;
const DERIVED_STRING_SLACK_PER_RECORD: usize = 256;

/// Caller quotas for OLAP proof construction.
///
/// These limits cover aggregate work and descriptor retention performed by
/// this proof, so a caller can reject a hostile graph before proof-owned
/// allocation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct OlapProofLimits {
    /// Maximum aggregate count of visited records and XML nodes.
    pub max_items: usize,
    /// Maximum aggregate UTF-8 bytes examined plus conservatively charged
    /// proof-owned descriptor copies.
    pub max_string_bytes: usize,
    /// Maximum already-inspected source size admitted to this proof.
    pub max_source_bytes: usize,
    /// Maximum bounded graph work units, including repeated collection edges.
    pub max_work: usize,
}

impl Default for OlapProofLimits {
    fn default() -> Self {
        Self {
            max_items: MAX_DEFAULT_ITEMS,
            max_string_bytes: MAX_DEFAULT_STRING_BYTES,
            max_source_bytes: MAX_DEFAULT_SOURCE_BYTES,
            max_work: MAX_DEFAULT_WORK,
        }
    }
}

/// A structural or caller-quota failure while proving the OLAP closure.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum OlapProofError {
    /// The source profile has no admitted OLAP projection in this owner.
    UnsupportedProfile,
    /// A caller quota was reached before proof-owned allocation.
    LimitExceeded {
        resource: &'static str,
        actual: usize,
        maximum: usize,
    },
    /// The graph contains an explicit contradiction.
    Invalid { path: String, detail: String },
    /// The graph contains a recognized member whose base grammar or identity
    /// is not sufficient to prove ownership.
    Unproven { path: String, detail: String },
    /// A bounded descriptor vector or map could not reserve its capacity.
    Allocation {
        resource: &'static str,
        detail: String,
    },
}

impl fmt::Display for OlapProofError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnsupportedProfile => {
                formatter.write_str("XLDM OLAP proof requires the version-140 profile")
            },
            Self::LimitExceeded {
                resource,
                actual,
                maximum,
            } => write!(
                formatter,
                "XLDM OLAP proof limit exceeded for {resource}: actual {actual}, maximum {maximum}"
            ),
            Self::Invalid { path, detail } => {
                write!(formatter, "invalid XLDM OLAP graph at {path}: {detail}")
            },
            Self::Unproven { path, detail } => {
                write!(
                    formatter,
                    "unproven XLDM OLAP ownership at {path}: {detail}"
                )
            },
            Self::Allocation { resource, detail } => {
                write!(
                    formatter,
                    "could not reserve XLDM OLAP proof {resource}: {detail}"
                )
            },
        }
    }
}

impl StdError for OlapProofError {}

/// Result type for prove_xldm140_olap.
pub type OlapProofResult<T> = Result<T, OlapProofError>;

/// The field used to establish an explicit XML reference.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum OlapReferenceField {
    DimensionId,
    CubeDimensionId,
    AttributeId,
    RelationshipId,
    RelationshipName,
    Id,
}

/// A retained explicit identifier reference, including its XML field kind.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapReference {
    pub field: OlapReferenceField,
    pub value: String,
}

/// A FileGroup-to-definition binding. ID and ObjectID deliberately stay
/// separate: section 2.1.2.3.1.3.1 defines them as different fields.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapFileGroupBinding {
    pub group_index: usize,
    pub class: FileGroupClass,
    pub id: String,
    pub name: String,
    pub object_id: String,
    pub definition_path: String,
    pub definition_id: String,
    pub definition_object_id: String,
    pub definition_name: Option<String>,
    pub members: Vec<String>,
    pub data_files: Vec<String>,
}

/// A proven table-to-Dimension binding. The generated table_id, metadata
/// name, and OLAP ID/name are exposed independently.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapTableBinding {
    pub table_id: String,
    pub metadata_path: String,
    pub metadata_name: Option<String>,
    pub dimension_path: String,
    pub dimension_id: String,
    pub dimension_object_id: String,
    pub dimension_name: Option<String>,
    pub attribute_ids: Vec<String>,
    pub data_files: Vec<String>,
}

/// A proven relationship binding. The metadata file's containing table is
/// the foreign/many side; primary_table comes from PrimaryTable.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapRelationshipBinding {
    pub metadata_path: String,
    pub relationship_name: Option<String>,
    /// The `RelId` token retained from the generated relationship filename.
    pub generated_relationship_id: String,
    pub containing_table: String,
    pub primary_table: String,
    pub primary_column: String,
    pub foreign_column: String,
    pub dimension_path: String,
    pub dimension_reference: OlapReference,
    pub relationship_index_paths: Vec<String>,
}

/// A cube's explicit dimension and attribute references.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapCubeBinding {
    pub path: String,
    pub id: String,
    pub dimensions: Vec<OlapCubeDimensionBinding>,
    pub measure_group_paths: Vec<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapCubeDimensionBinding {
    pub reference: OlapReference,
    pub attribute_references: Vec<OlapReference>,
    pub unknown: bool,
}

/// A measure-group binding with explicit degenerate/reference dimensions.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapMeasureGroupBinding {
    pub path: String,
    pub id: String,
    pub object_id: String,
    pub table_dimension_id: String,
    pub dimensions: Vec<OlapMeasureGroupDimensionBinding>,
    pub partition_paths: Vec<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapMeasureGroupDimensionBinding {
    pub kind: MeasureGroupDimensionKind,
    pub dimension_reference: OlapReference,
    pub attribute_references: Vec<OlapReference>,
    pub unknown: bool,
}

#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum MeasureGroupDimensionKind {
    Degenerate,
    Reference,
}

/// A partition definition owned by one measure group.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapPartitionBinding {
    pub path: String,
    pub id: String,
    pub object_id: String,
    pub measure_group_path: String,
    pub data_files: Vec<String>,
}

/// A recognized physical or XML member that could not be associated with a
/// proven OLAP owner. It remains visible and makes is_complete false.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct OlapUnknownMember {
    pub path: String,
    pub reason: String,
}

/// The bounded, source-backed OLAP closure proof.
#[derive(Debug)]
pub struct Xldm140OlapProof<'storage, 'source> {
    storage: &'storage Storage<'source>,
    file_groups: Vec<OlapFileGroupBinding>,
    tables: Vec<OlapTableBinding>,
    relationships: Vec<OlapRelationshipBinding>,
    cube: Option<OlapCubeBinding>,
    measure_groups: Vec<OlapMeasureGroupBinding>,
    partitions: Vec<OlapPartitionBinding>,
    unknown_members: Vec<OlapUnknownMember>,
}

impl Xldm140OlapProof<'_, '_> {
    /// Return the exact inspected outer source without copying member bytes.
    #[must_use]
    pub fn storage(&self) -> &Storage<'_> {
        self.storage
    }

    #[must_use]
    pub fn file_groups(&self) -> &[OlapFileGroupBinding] {
        &self.file_groups
    }

    #[must_use]
    pub fn tables(&self) -> &[OlapTableBinding] {
        &self.tables
    }

    #[must_use]
    pub fn relationships(&self) -> &[OlapRelationshipBinding] {
        &self.relationships
    }

    #[must_use]
    pub fn cube(&self) -> Option<&OlapCubeBinding> {
        self.cube.as_ref()
    }

    #[must_use]
    pub fn measure_groups(&self) -> &[OlapMeasureGroupBinding] {
        &self.measure_groups
    }

    #[must_use]
    pub fn partitions(&self) -> &[OlapPartitionBinding] {
        &self.partitions
    }

    #[must_use]
    pub fn unknown_members(&self) -> &[OlapUnknownMember] {
        &self.unknown_members
    }

    /// A proof is complete only when every required table relationship and
    /// every recognized extra member has an admitted owner.
    #[must_use]
    pub fn is_complete(&self) -> bool {
        self.unknown_members.is_empty()
    }
}

/// Prove the version-140 OLAP projection against section 2.2, section 2.5,
/// and the already parsed section 2.6 model.
///
/// `model` and `metadata` are expected to come from the corresponding
/// source-bound inspectors. This owner does not duplicate the complete
/// MS-SSAS base grammar; it proves the section 2.6 table-model edges and
/// refuses any edge that is absent, ambiguous, or outside the inspected
/// storage/FileGroup closure.
///
/// This function never treats a generated filename as sufficient identity. It
/// requires FileGroup ID/ObjectID bindings, explicit OLAP references, exact
/// file-list membership, and the table-local metadata ownership needed by
/// sections 2.6.5 through 2.6.8.
pub fn prove_xldm140_olap<'storage, 'source>(
    storage: &'storage Storage<'source>,
    metadata: &MetadataModel<'_>,
    model: &OlapModel<'_>,
    limits: OlapProofLimits,
) -> OlapProofResult<Xldm140OlapProof<'storage, 'source>> {
    if storage.profile() != StorageProfile::Xldm140 {
        return Err(OlapProofError::UnsupportedProfile);
    }
    if !storage.backup_log.is_olap {
        return Err(unproven(
            "BackupLog",
            "OlapInfo is false for an OLAP proof request",
        ));
    }
    let mut budget = Budget::new(limits);
    budget.source_bytes(storage.bytes().len())?;
    preflight_storage(storage, &mut budget)?;
    preflight_metadata(metadata, &mut budget)?;
    preflight_olap(model, &mut budget)?;
    preflight_retained_strings(storage, metadata, model, &mut budget)?;
    let logged_file_count =
        storage
            .backup_log
            .file_groups
            .iter()
            .try_fold(0usize, |total, group| {
                total
                    .checked_add(group.files.len())
                    .ok_or_else(|| limit("graph work", usize::MAX, budget.limits.max_work))
            })?;
    let graph_nodes = storage
        .files
        .len()
        .checked_add(storage.backup_log.file_groups.len())
        .and_then(|value| value.checked_add(logged_file_count))
        .and_then(|value| value.checked_add(metadata.files.len()))
        .and_then(|value| value.checked_add(model.files.len()))
        .ok_or_else(|| limit("graph work", usize::MAX, budget.limits.max_work))?;
    budget.work_bound(graph_nodes)?;

    let mut physical_paths = HashMap::new();
    reserve_map(
        &mut physical_paths,
        storage.files.len(),
        "physical path index",
    )?;
    for (index, file) in storage.files.iter().enumerate() {
        if physical_paths.insert(file.path.as_str(), index).is_some() {
            return Err(invalid(
                file.path.as_str(),
                "duplicate physical storage path",
            ));
        }
    }

    let mut definitions = Vec::new();
    reserve_vec(&mut definitions, model.files.len(), "OLAP definitions")?;
    let mut definitions_by_key = HashMap::new();
    reserve_map(
        &mut definitions_by_key,
        model.files.len(),
        "OLAP definition index",
    )?;
    for (index, file) in model.files.iter().enumerate() {
        let OlapFileKind::Definition(kind) = file.kind else {
            continue;
        };
        let OlapDocument::Definition(definition) = &file.document else {
            return Err(invalid(
                file.storage_path,
                "definition file has a non-definition document",
            ));
        };
        if definition.kind != kind {
            return Err(invalid(
                file.storage_path,
                "OLAP file kind disagrees with its definition kind",
            ));
        }
        validate_definition_shape(file.storage_path, definition)?;
        let object_id = required_scalar(&definition.object, "ObjectID", file.storage_path)?;
        let object_name = optional_scalar(&definition.object, "Name", file.storage_path)?;
        let entry = DefinitionRef {
            model_index: index,
            kind,
            path: file.storage_path.to_owned(),
            id: definition.object_id.clone(),
            object_id: object_id.to_owned(),
            name: object_name.map(str::to_owned),
        };
        if definitions_by_key
            .insert((kind, definition.object_id.as_str()), definitions.len())
            .is_some()
        {
            return Err(invalid(file.storage_path, "duplicate OLAP kind and ID"));
        }
        definitions.push(entry);
    }

    let tables = collect_table_paths(metadata)?;
    let mut table_bindings = Vec::new();
    reserve_vec(&mut table_bindings, tables.len(), "table bindings")?;
    let mut table_by_id = HashMap::new();
    reserve_map(&mut table_by_id, tables.len(), "table path index")?;

    let mut groups_by_key = HashMap::new();
    reserve_map(
        &mut groups_by_key,
        storage.backup_log.file_groups.len(),
        "FileGroup index",
    )?;
    for (index, group) in storage.backup_log.file_groups.iter().enumerate() {
        if groups_by_key
            .insert((group.class, group.id.as_str()), index)
            .is_some()
        {
            return Err(invalid(&group.id, "duplicate FileGroup class and ID"));
        }
    }

    let database = unique_definition(&definitions, OlapObjectKind::Database)?;
    let database_definition = definition(model, database)?;
    let database_location = required_scalar(
        &database_definition.object,
        "DbStorageLocation",
        database.path.as_str(),
    )?;
    validate_group_persistence(
        storage,
        &definitions,
        model,
        database_location,
        &groups_by_key,
    )?;

    let mut file_group_bindings = Vec::new();
    reserve_vec(
        &mut file_group_bindings,
        storage.backup_log.file_groups.len(),
        "FileGroup bindings",
    )?;
    let mut unknown_members = Vec::new();
    let unknown_capacity = storage
        .backup_log
        .file_groups
        .len()
        .checked_add(model.files.len())
        .and_then(|value| value.checked_add(storage.files.len()))
        .and_then(|value| value.checked_add(metadata.files.len()))
        .ok_or_else(|| limit("unknown OLAP members", usize::MAX, limits.max_items))?;
    reserve_vec(
        &mut unknown_members,
        unknown_capacity,
        "unknown OLAP members",
    )?;
    let mut matched_groups = HashSet::new();
    reserve_set(
        &mut matched_groups,
        storage.backup_log.file_groups.len(),
        "matched FileGroup set",
    )?;

    for definition_ref in &definitions {
        let class = class_for_kind(definition_ref.kind);
        let Some(group_index) = groups_by_key.get(&(class, definition_ref.id.as_str())) else {
            push_unknown(
                &mut unknown_members,
                OlapUnknownMember {
                    path: definition_ref.path.clone(),
                    reason: "OLAP definition has no FileGroup with the same explicit ID".into(),
                },
            )?;
            continue;
        };
        let group = &storage.backup_log.file_groups[*group_index];
        let definition = definition(model, definition_ref)?;
        validate_definition_group(definition_ref, definition, group, database_location)?;
        matched_groups.insert(*group_index);
        let mut members = Vec::new();
        reserve_vec(&mut members, group.files.len(), "FileGroup members")?;
        for file in &group.files {
            members.push(file.storage_path.clone());
        }
        let data_files =
            validate_data_file_list(definition_ref, definition, group, &members, &physical_paths)?;
        file_group_bindings.push(OlapFileGroupBinding {
            group_index: *group_index,
            class: group.class,
            id: group.id.clone(),
            name: group.name.clone(),
            object_id: group.object_id.clone(),
            definition_path: definition_ref.path.clone(),
            definition_id: definition_ref.id.clone(),
            definition_object_id: definition_ref.object_id.clone(),
            definition_name: definition_ref.name.clone(),
            members,
            data_files,
        });
    }
    for (index, group) in storage.backup_log.file_groups.iter().enumerate() {
        if !matched_groups.contains(&index) {
            push_unknown(
                &mut unknown_members,
                OlapUnknownMember {
                    path: group.id.clone(),
                    reason: format!(
                        "FileGroup class {} has no matching OLAP definition",
                        group.class.code()
                    ),
                },
            )?;
        }
    }

    for table in &tables {
        let Some(dimension_ref) = find_dimension_for_table(&definitions, table, model)? else {
            return Err(unproven(
                table.metadata_path.as_str(),
                "table has no Dimension definition bound through its containing .dim folder",
            ));
        };
        if !physical_paths.contains_key(table.metadata_path.as_str()) {
            return Err(unproven(
                table.metadata_path.as_str(),
                "table metadata is absent from the XLDM virtual directory",
            ));
        }
        require_group_member(
            &file_group_bindings,
            dimension_ref.path.as_str(),
            table.metadata_path.as_str(),
        )?;
        let dimension = definition(model, dimension_ref)?;
        let attributes = dimension_attributes(dimension, dimension_ref.path.as_str())?;
        let metadata_file = &metadata.files[table.metadata_index];
        if metadata_file.table.class.as_str() != "XMSimpleTable" {
            return Err(invalid(
                table.metadata_path.as_str(),
                "table metadata root is not an XMSimpleTable",
            ));
        }
        let columns = metadata_columns(metadata_file, table.metadata_path.as_str())?;
        if attributes.len() != columns.len() {
            return Err(invalid(
                dimension_ref.path.as_str(),
                "Dimension Attribute count does not equal table column count",
            ));
        }
        let mut attribute_ids = Vec::new();
        reserve_vec(&mut attribute_ids, attributes.len(), "table attribute IDs")?;
        for attribute in &attributes {
            attribute_ids.push(attribute.value.clone());
        }
        let mut attribute_set = HashSet::new();
        reserve_set(
            &mut attribute_set,
            attribute_ids.len(),
            "dimension attribute set",
        )?;
        for attribute in &attribute_ids {
            if !attribute_set.insert(attribute.as_str()) {
                return Err(invalid(
                    dimension_ref.path.as_str(),
                    "duplicate Dimension Attribute ID",
                ));
            }
        }
        for column in &columns {
            if !attribute_set.contains(column.as_str()) {
                return Err(unproven(
                    table.metadata_path.as_str(),
                    "a metadata column has no explicit matching Dimension Attribute ID",
                ));
            }
        }
        let data_files = file_group_bindings
            .iter()
            .find(|binding| binding.definition_path == dimension_ref.path)
            .map(|binding| binding.data_files.clone())
            .unwrap_or_default();
        let table_binding = OlapTableBinding {
            table_id: table.table_id.clone(),
            metadata_path: table.metadata_path.clone(),
            metadata_name: metadata_file.table.name.clone(),
            dimension_path: dimension_ref.path.clone(),
            dimension_id: dimension_ref.id.clone(),
            dimension_object_id: dimension_ref.object_id.clone(),
            dimension_name: dimension_ref.name.clone(),
            attribute_ids,
            data_files,
        };
        if table_by_id
            .insert(table.table_id.as_str(), table_bindings.len())
            .is_some()
        {
            return Err(invalid(
                table.metadata_path.as_str(),
                "duplicate generated TableID in metadata folders",
            ));
        }
        table_bindings.push(table_binding);
    }

    let relationships = prove_relationships(
        metadata,
        model,
        &table_bindings,
        &definitions,
        &physical_paths,
        &file_group_bindings,
        &mut unknown_members,
    )?;
    let cube = prove_cube(model, &definitions, &table_bindings, &mut unknown_members)?;
    let measure_groups = prove_measure_groups(
        model,
        &definitions,
        &table_bindings,
        &relationships,
        &mut unknown_members,
    )?;
    let partitions = prove_partitions(model, &definitions, &measure_groups, &mut unknown_members)?;
    validate_cube_file_lists(cube.as_ref(), &measure_groups)?;
    validate_measure_group_file_lists(&measure_groups, &partitions)?;
    validate_partition_ownership(&measure_groups, &partitions)?;
    record_unlinked_members(
        storage,
        metadata,
        model,
        &file_group_bindings,
        &mut unknown_members,
    )?;

    Ok(Xldm140OlapProof {
        storage,
        file_groups: file_group_bindings,
        tables: table_bindings,
        relationships,
        cube,
        measure_groups,
        partitions,
        unknown_members,
    })
}

#[derive(Clone, Debug)]
struct DefinitionRef {
    model_index: usize,
    kind: OlapObjectKind,
    path: String,
    id: String,
    object_id: String,
    name: Option<String>,
}

#[derive(Clone, Debug)]
struct TablePath {
    metadata_index: usize,
    table_id: String,
    metadata_path: String,
    folder: String,
}

#[derive(Clone, Debug)]
struct RelationInput {
    metadata_path: String,
    relationship_name: Option<String>,
    generated_relationship_id: String,
    containing_folder: String,
    containing_table: String,
    primary_table: String,
    primary_column: String,
    foreign_column: String,
}

fn class_for_kind(kind: OlapObjectKind) -> FileGroupClass {
    match kind {
        OlapObjectKind::Database => FileGroupClass::Database,
        OlapObjectKind::DataSource => FileGroupClass::DataSource,
        OlapObjectKind::DataSourceView => FileGroupClass::DataSourceView,
        OlapObjectKind::Cube => FileGroupClass::Cube,
        OlapObjectKind::Dimension => FileGroupClass::Dimension,
        OlapObjectKind::MdxScript => FileGroupClass::MdxScript,
        OlapObjectKind::MeasureGroup => FileGroupClass::MeasureGroup,
        OlapObjectKind::Partition => FileGroupClass::Partition,
    }
}

fn definition<'a>(
    model: &'a OlapModel<'_>,
    value: &DefinitionRef,
) -> OlapProofResult<&'a OlapDefinition> {
    let file = model.files.get(value.model_index).ok_or_else(|| {
        invalid(
            value.path.as_str(),
            "definition index is outside the OLAP model",
        )
    })?;
    match &file.document {
        OlapDocument::Definition(definition) => Ok(definition),
        _ => Err(invalid(
            value.path.as_str(),
            "definition index has a non-definition document",
        )),
    }
}

fn unique_definition(
    definitions: &[DefinitionRef],
    kind: OlapObjectKind,
) -> OlapProofResult<&DefinitionRef> {
    let mut found = definitions.iter().filter(|item| item.kind == kind);
    let Some(value) = found.next() else {
        return Err(unproven(
            kind.element_name_for_proof(),
            "required OLAP definition is missing",
        ));
    };
    if found.next().is_some() {
        return Err(invalid(
            kind.element_name_for_proof(),
            "multiple singleton OLAP definitions",
        ));
    }
    Ok(value)
}

fn collect_table_paths(metadata: &MetadataModel<'_>) -> OlapProofResult<Vec<TablePath>> {
    let mut tables = Vec::new();
    reserve_vec(
        &mut tables,
        metadata
            .files
            .iter()
            .filter(|file| file.kind == MetadataFileKind::Table)
            .count(),
        "table paths",
    )?;
    for (metadata_index, file) in metadata.files.iter().enumerate() {
        if file.kind != MetadataFileKind::Table {
            continue;
        }
        let (folder, table_id) = metadata_table_path(file.storage_path)?;
        tables.push(TablePath {
            metadata_index,
            table_id,
            metadata_path: file.storage_path.to_owned(),
            folder,
        });
    }
    Ok(tables)
}

fn metadata_table_path(path: &str) -> OlapProofResult<(String, String)> {
    let (folder, name) = path
        .rsplit_once('/')
        .ok_or_else(|| invalid(path, "table metadata path has no containing folder"))?;
    let folder_name = folder
        .rsplit('/')
        .next()
        .ok_or_else(|| invalid(path, "table metadata path has no dimension folder"))?;
    let Some(folder_prefix) = folder_name.strip_suffix(".dim") else {
        return Err(invalid(path, "table metadata is outside a .dim folder"));
    };
    let Some((table_id, version)) = folder_prefix.rsplit_once('.') else {
        return Err(invalid(path, "dimension folder has no version"));
    };
    if version != "0" || !valid_user_id(table_id) {
        return Err(invalid(
            path,
            "table dimension folder is not the required .0.dim owner",
        ));
    }
    let Some(file_prefix) = name.strip_suffix(".tbl.xml") else {
        return Err(invalid(path, "table metadata has no .tbl.xml suffix"));
    };
    let Some((file_table_id, file_version)) = file_prefix.rsplit_once('.') else {
        return Err(invalid(path, "table metadata has no version"));
    };
    if file_table_id != table_id
        || file_version.is_empty()
        || !file_version.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(invalid(
            path,
            "table metadata filename does not belong to its dimension folder",
        ));
    }
    Ok((folder.to_owned(), table_id.to_owned()))
}

fn find_dimension_for_table<'a>(
    definitions: &'a [DefinitionRef],
    table: &TablePath,
    model: &OlapModel<'_>,
) -> OlapProofResult<Option<&'a DefinitionRef>> {
    let mut result = None;
    for value in definitions
        .iter()
        .filter(|item| item.kind == OlapObjectKind::Dimension)
    {
        let definition = definition(model, value)?;
        let parent = value.path.rsplit_once('/').map_or("", |parts| parts.0);
        let base = value.path.rsplit('/').next().unwrap_or(value.path.as_str());
        let Some(prefix) = base.strip_suffix(".dim.xml") else {
            continue;
        };
        let Some((table_id, _version)) = prefix.rsplit_once('.') else {
            continue;
        };
        let folder_version = definition.extension.persist_location;
        let folder = if parent.is_empty() {
            format!("{table_id}.{folder_version}.dim")
        } else {
            format!("{parent}/{table_id}.{folder_version}.dim")
        };
        if folder == table.folder {
            if result.replace(value).is_some() {
                return Err(invalid(
                    table.metadata_path.as_str(),
                    "multiple Dimension definitions own one table folder",
                ));
            }
        }
    }
    Ok(result)
}

fn metadata_columns(
    file: &super::metadata::MetadataFile<'_>,
    path: &str,
) -> OlapProofResult<Vec<String>> {
    let Some(columns) = file.table.collection("Columns") else {
        return Err(unproven(path, "table has no Columns collection"));
    };
    let mut output = Vec::new();
    reserve_vec(&mut output, columns.len(), "metadata column names")?;
    for column in columns {
        if column.class.as_str() != "XMRawColumn" {
            return Err(invalid(
                path,
                "table Columns contains a non-XMRawColumn object",
            ));
        }
        let name = column
            .name
            .as_deref()
            .ok_or_else(|| unproven(path, "a metadata column has no object name"))?;
        output.push(name.to_owned());
    }
    Ok(output)
}

fn dimension_attributes(
    definition: &OlapDefinition,
    path: &str,
) -> OlapProofResult<Vec<OlapReference>> {
    let Some(collection) = definition.object.child("Attributes") else {
        return Err(unproven(path, "Dimension has no Attributes collection"));
    };
    let mut output = Vec::new();
    reserve_vec(
        &mut output,
        collection.children.len(),
        "Dimension Attributes",
    )?;
    for item in &collection.children {
        if item.name != "Attribute" {
            return Err(invalid(
                path,
                "Dimension Attributes contains a non-Attribute member",
            ));
        }
        require_type(item, "DimensionAttribute", path)?;
        let value = required_scalar(item, "ID", path)?;
        output.push(OlapReference {
            field: OlapReferenceField::Id,
            value: value.to_owned(),
        });
    }
    Ok(output)
}

fn validate_group_persistence(
    storage: &Storage<'_>,
    definitions: &[DefinitionRef],
    model: &OlapModel<'_>,
    database_location: &str,
    groups_by_key: &HashMap<(FileGroupClass, &str), usize>,
) -> OlapProofResult<()> {
    let database_persist_location = model
        .files
        .iter()
        .find_map(|file| match &file.document {
            OlapDocument::Definition(definition) if definition.kind == OlapObjectKind::Database => {
                Some(definition.extension.persist_location)
            },
            _ => None,
        })
        .ok_or_else(|| invalid("section 2.6", "model has no Database definition"))?;
    for group in &storage.backup_log.file_groups {
        if group.persist_location_path != database_location {
            return Err(invalid(
                &group.persist_location_path,
                "FileGroup PersistLocationPath does not match Database DbStorageLocation",
            ));
        }
    }
    for definition_ref in definitions {
        let class = class_for_kind(definition_ref.kind);
        let Some(index) = groups_by_key.get(&(class, definition_ref.id.as_str())) else {
            continue;
        };
        let group = &storage.backup_log.file_groups[*index];
        let definition = definition(model, definition_ref)?;
        if group.object_version != definition.extension.object_version {
            return Err(invalid(
                definition_ref.path.as_str(),
                "FileGroup ObjectVersion does not match OLAP ObjectVersion",
            ));
        }
        // SdfFileGroupType.PersistLocation is the database-folder version
        // (MS-XLDM 2.1.2.3.1.3.1), so every group uses the database object's
        // PersistLocation.  An individual OLAP object's PersistLocation is
        // used for its own generated child folder and may therefore differ
        // (for example, a Dimension uses the .0.dim folder while the
        // database folder is .1.db).
        if group.persist_location != database_persist_location {
            return Err(invalid(
                definition_ref.path.as_str(),
                "FileGroup PersistLocation does not match Database PersistLocation",
            ));
        }
    }
    Ok(())
}

fn validate_definition_group(
    definition_ref: &DefinitionRef,
    definition: &OlapDefinition,
    group: &super::model::FileGroup,
    database_location: &str,
) -> OlapProofResult<()> {
    if group.id != definition.object_id {
        return Err(invalid(
            definition_ref.path.as_str(),
            "FileGroup ID does not match the OLAP ID",
        ));
    }
    let object_id = required_scalar(&definition.object, "ObjectID", definition_ref.path.as_str())?;
    if group.object_id != object_id {
        return Err(invalid(
            definition_ref.path.as_str(),
            "FileGroup ObjectID does not match the OLAP ObjectID",
        ));
    }
    let name = optional_scalar(&definition.object, "Name", definition_ref.path.as_str())?;
    if let Some(name) = name {
        if group.name != name {
            return Err(invalid(
                definition_ref.path.as_str(),
                "FileGroup Name does not match the OLAP Name",
            ));
        }
    }
    if group.persist_location_path != database_location {
        return Err(invalid(
            definition_ref.path.as_str(),
            "FileGroup PersistLocationPath is outside the database storage location",
        ));
    }
    Ok(())
}

fn validate_definition_shape(path: &str, definition: &OlapDefinition) -> OlapProofResult<()> {
    if definition.object.name != definition.kind.element_name_for_proof() {
        return Err(invalid(
            path,
            "OLAP definition root object does not match its file kind",
        ));
    }
    if definition.extension.ordinal < 0
        || definition.extension.object_version < 0
        || definition.extension.persist_location < 0
    {
        return Err(invalid(
            path,
            "OLAP Ordinal, ObjectVersion, and PersistLocation must be nonnegative",
        ));
    }
    let object_id = required_scalar(&definition.object, "ID", path)?;
    if object_id != definition.object_id {
        return Err(invalid(
            path,
            "parsed OLAP ID disagrees with the definition object ID",
        ));
    }
    if let Some(name) = &definition.object_name {
        if optional_scalar(&definition.object, "Name", path)? != Some(name.as_str()) {
            return Err(invalid(
                path,
                "parsed OLAP Name disagrees with the definition object name",
            ));
        }
    }
    let base = path.rsplit('/').next().unwrap_or(path);
    let prefix = base
        .strip_suffix(definition.kind.suffix_for_proof())
        .ok_or_else(|| invalid(path, "OLAP definition filename has the wrong suffix"))?;
    let (user_id, version) = prefix
        .rsplit_once('.')
        .ok_or_else(|| invalid(path, "OLAP definition filename has no version"))?;
    if user_id.is_empty()
        || version.is_empty()
        || !version.bytes().all(|byte| byte.is_ascii_digit())
        || version.parse::<i32>().ok() != Some(definition.extension.object_version)
    {
        return Err(invalid(
            path,
            "OLAP ObjectVersion does not match the generated definition filename",
        ));
    }
    Ok(())
}

fn require_group_member(
    groups: &[OlapFileGroupBinding],
    definition_path: &str,
    member_path: &str,
) -> OlapProofResult<()> {
    let group = groups
        .iter()
        .find(|group| group.definition_path == definition_path)
        .ok_or_else(|| {
            unproven(
                member_path,
                "owning OLAP definition has no proven FileGroup",
            )
        })?;
    if group.members.iter().any(|member| member == member_path) {
        Ok(())
    } else {
        Err(unproven(
            member_path,
            "member is absent from its owning OLAP FileGroup",
        ))
    }
}

fn validate_data_file_list(
    definition_ref: &DefinitionRef,
    definition: &OlapDefinition,
    group: &super::model::FileGroup,
    members: &[String],
    physical_paths: &HashMap<&str, usize>,
) -> OlapProofResult<Vec<String>> {
    let folder = persist_folder(definition_ref.path.as_str(), definition)?;
    let mut member_set = HashSet::new();
    reserve_set(&mut member_set, members.len(), "FileGroup member set")?;
    for member in members {
        member_set.insert(member.as_str());
    }
    for member in members {
        if !physical_paths.contains_key(member.as_str()) {
            return Err(invalid(
                member.as_str(),
                "FileGroup member is absent from the virtual directory",
            ));
        }
    }
    if !member_set.contains(definition_ref.path.as_str()) {
        return Err(invalid(
            definition_ref.path.as_str(),
            "FileGroup does not contain its OLAP definition member",
        ));
    }
    let mut result = Vec::new();
    reserve_vec(
        &mut result,
        definition.extension.data_files.len(),
        "OLAP DataFileList",
    )?;
    for item in &definition.extension.data_files {
        let qualified = qualify_file_list_item(&folder, item, definition_ref.path.as_str())?;
        if !member_set.contains(qualified.as_str()) {
            return Err(invalid(
                definition_ref.path.as_str(),
                "DataFileList member is not in its FileGroup",
            ));
        }
        if !physical_paths.contains_key(qualified.as_str()) {
            return Err(invalid(
                qualified.as_str(),
                "DataFileList member is absent from the virtual directory",
            ));
        }
        result.push(qualified);
    }
    let mut expected_raw = HashSet::new();
    let raw_count = group
        .files
        .iter()
        .filter(|file| is_materialized_data_kind(file.generated.kind))
        .count();
    reserve_set(&mut expected_raw, raw_count, "materialized data set")?;
    for file in &group.files {
        if is_materialized_data_kind(file.generated.kind) {
            validate_materialized_member_owner(
                definition_ref,
                &folder,
                file.storage_path.as_str(),
                file.generated.kind,
            )?;
            if file.generated.normalized_path != file.storage_path {
                return Err(invalid(
                    file.storage_path.as_str(),
                    "generated path disagrees with the logged storage path",
                ));
            }
            expected_raw.insert(file.storage_path.as_str());
        }
    }
    let mut actual = HashSet::new();
    reserve_set(&mut actual, result.len(), "DataFileList set")?;
    for value in &result {
        actual.insert(value.as_str());
    }
    if actual != expected_raw {
        return Err(invalid(
            definition_ref.path.as_str(),
            "DataFileList is not the exact materialized-member set",
        ));
    }
    Ok(result)
}

fn is_materialized_data_kind(kind: GeneratedNameKind) -> bool {
    matches!(
        kind,
        GeneratedNameKind::ColumnData
            | GeneratedNameKind::ColumnDictionary
            | GeneratedNameKind::ColumnHashIndex
            | GeneratedNameKind::ColumnPositionToId
            | GeneratedNameKind::ColumnIdToPosition
            | GeneratedNameKind::ColumnHierarchyMetadata
            | GeneratedNameKind::UserHierarchyMetadata
            | GeneratedNameKind::TableRelationshipIndex
            | GeneratedNameKind::UserHierarchyChildCount
            | GeneratedNameKind::UserHierarchyFirstChildPosition
            | GeneratedNameKind::UserHierarchyParentPosition
            | GeneratedNameKind::UserHierarchyMultilevelId
    )
}

fn validate_materialized_member_owner(
    definition_ref: &DefinitionRef,
    folder: &str,
    path: &str,
    kind: GeneratedNameKind,
) -> OlapProofResult<()> {
    if definition_ref.kind != OlapObjectKind::Dimension {
        return Ok(());
    }
    let table_id = folder
        .rsplit('/')
        .next()
        .and_then(|value| value.strip_suffix(".dim"))
        .and_then(|value| value.rsplit_once('.'))
        .map(|(table_id, _version)| table_id)
        .filter(|value| valid_user_id(value))
        .ok_or_else(|| {
            invalid(
                definition_ref.path.as_str(),
                "Dimension persistence folder has no valid table owner",
            )
        })?;
    let parent = path.rsplit_once('/').map_or("", |parts| parts.0);
    if parent != folder {
        return Err(unproven(
            path,
            "materialized Dimension member is outside its owning table folder",
        ));
    }
    let name = path.rsplit('/').next().unwrap_or(path);
    if !generated_member_has_table_owner(name, kind, table_id) {
        return Err(unproven(
            path,
            format!("materialized Dimension member does not encode its owning table {table_id}"),
        ));
    }
    Ok(())
}

fn generated_member_has_table_owner(name: &str, kind: GeneratedNameKind, table_id: &str) -> bool {
    match kind {
        GeneratedNameKind::ColumnData => {
            let Some(stem) = name.strip_suffix(".idf") else {
                return false;
            };
            let Some((_ordinal, body)) = stem.split_once('.') else {
                return false;
            };
            body.strip_prefix(table_id)
                .is_some_and(|rest| rest.starts_with('.'))
        },
        GeneratedNameKind::ColumnDictionary => {
            let Some(stem) = name.strip_suffix(".dictionary") else {
                return false;
            };
            let Some((_ordinal, body)) = stem.split_once('.') else {
                return false;
            };
            body.strip_prefix(table_id)
                .is_some_and(|rest| rest.starts_with('.'))
        },
        GeneratedNameKind::ColumnHashIndex => {
            let Some(stem) = name.strip_suffix(".hidx") else {
                return false;
            };
            let Some((_ordinal, body)) = stem.split_once('.') else {
                return false;
            };
            body.strip_prefix("H$")
                .and_then(|value| value.strip_prefix(table_id))
                .is_some_and(|rest| rest.starts_with('$'))
        },
        GeneratedNameKind::ColumnPositionToId | GeneratedNameKind::ColumnIdToPosition => {
            let Some(stem) = name.strip_suffix(".idf") else {
                return false;
            };
            let Some((_ordinal, body)) = stem.split_once('.') else {
                return false;
            };
            body.strip_prefix("H$")
                .and_then(|value| value.strip_prefix(table_id))
                .is_some_and(|rest| rest.starts_with('$'))
        },
        GeneratedNameKind::TableRelationshipIndex => {
            let Some(stem) = name.strip_suffix(".idf") else {
                return false;
            };
            let Some((_ordinal, body)) = stem.split_once('.') else {
                return false;
            };
            body.strip_prefix("R$")
                .and_then(|value| value.strip_prefix(table_id))
                .is_some_and(|rest| rest.starts_with('$'))
        },
        GeneratedNameKind::ColumnHierarchyMetadata | GeneratedNameKind::UserHierarchyMetadata => {
            let Some(stem) = name.strip_suffix(".tbl.xml") else {
                return false;
            };
            let Some((body, _version)) = stem.rsplit_once('.') else {
                return false;
            };
            body.strip_prefix("H$")
                .or_else(|| body.strip_prefix("U$"))
                .and_then(|value| value.strip_prefix(table_id))
                .is_some_and(|rest| rest.starts_with('$'))
        },
        GeneratedNameKind::UserHierarchyChildCount
        | GeneratedNameKind::UserHierarchyFirstChildPosition
        | GeneratedNameKind::UserHierarchyParentPosition
        | GeneratedNameKind::UserHierarchyMultilevelId => {
            let Some(stem) = name.strip_suffix(".idf") else {
                return false;
            };
            let Some((_ordinal, body)) = stem.split_once('.') else {
                return false;
            };
            body.strip_prefix("U$")
                .and_then(|value| value.strip_prefix(table_id))
                .is_some_and(|rest| rest.starts_with('$'))
        },
        GeneratedNameKind::TableInformation
        | GeneratedNameKind::TableMetadata
        | GeneratedNameKind::TableRelationshipMetadata
        | GeneratedNameKind::CryptographicKey
        | GeneratedNameKind::DatabaseDefinition
        | GeneratedNameKind::DataSourceViewDefinition
        | GeneratedNameKind::CubeDefinition
        | GeneratedNameKind::DataSourceOrDimensionDefinition
        | GeneratedNameKind::CubeInformation
        | GeneratedNameKind::PartitionInformation
        | GeneratedNameKind::MdxScriptMetadata
        | GeneratedNameKind::MeasureGroupMetadata
        | GeneratedNameKind::PartitionMetadata => false,
    }
}

fn relationship_collection<'a>(
    object: &'a OlapElement,
    path: &str,
    unknown_members: &mut Vec<OlapUnknownMember>,
) -> OlapProofResult<Option<&'a OlapElement>> {
    let mut collections = object
        .children
        .iter()
        .filter(|child| child.name == "Relationships");
    let Some(collection) = collections.next() else {
        return Ok(None);
    };
    if collections.next().is_some() {
        push_unknown(
            unknown_members,
            OlapUnknownMember {
                path: path.to_owned(),
                reason: "object contains duplicate Relationships collections".into(),
            },
        )?;
    }
    if !collection.text.trim().is_empty() {
        return Err(invalid(path, "Relationships collection contains text"));
    }
    for child in &collection.children {
        if child.name != "Relationship" {
            push_unknown(
                unknown_members,
                OlapUnknownMember {
                    path: path.to_owned(),
                    reason: format!(
                        "Relationships collection contains unknown member {}",
                        child.name
                    ),
                },
            )?;
        }
    }
    Ok(Some(collection))
}

fn metadata_relationship_collection<'a>(
    object: &'a MetadataObject,
    path: &str,
) -> OlapProofResult<Option<&'a [MetadataObject]>> {
    let mut collections = object
        .collections
        .iter()
        .filter(|collection| collection.name == "Relationships");
    let Some(collection) = collections.next() else {
        return Ok(None);
    };
    if collections.next().is_some() {
        return Err(invalid(
            path,
            "metadata object contains duplicate Relationships collections",
        ));
    }
    Ok(Some(collection.objects.as_slice()))
}

fn prove_relationships(
    metadata: &MetadataModel<'_>,
    model: &OlapModel<'_>,
    table_bindings: &[OlapTableBinding],
    definitions: &[DefinitionRef],
    physical_paths: &HashMap<&str, usize>,
    file_group_bindings: &[OlapFileGroupBinding],
    unknown_members: &mut Vec<OlapUnknownMember>,
) -> OlapProofResult<Vec<OlapRelationshipBinding>> {
    let mut inputs = Vec::new();
    let relationship_count = metadata
        .files
        .iter()
        .filter(|file| file.kind == MetadataFileKind::TableRelationship)
        .count();
    reserve_vec(&mut inputs, relationship_count, "relationship metadata")?;
    for file in metadata
        .files
        .iter()
        .filter(|file| file.kind == MetadataFileKind::TableRelationship)
    {
        let (containing_folder, containing_table, generated_relationship_id) =
            metadata_table_path_for_relationship(file.storage_path)?;
        if file.table.class.as_str() != "XMSimpleTable" {
            return Err(invalid(
                file.storage_path,
                "relationship metadata root is not an XMSimpleTable",
            ));
        }
        let Some(collection) = metadata_relationship_collection(&file.table, file.storage_path)?
        else {
            return Err(unproven(
                file.storage_path,
                "relationship file has no Relationships collection",
            ));
        };
        if collection.len() != 1 {
            return Err(invalid(
                file.storage_path,
                "relationship file must contain one relationship",
            ));
        }
        let relation = &collection[0];
        if relation.class.as_str() != "XMRelationship" {
            return Err(invalid(
                file.storage_path,
                "relationship collection contains a non-XMRelationship object",
            ));
        }
        let relationship_name = relation.name.clone();
        let primary_table = required_property(relation, "PrimaryTable", file.storage_path)?;
        let primary_column = required_property(relation, "PrimaryColumn", file.storage_path)?;
        let foreign_column = required_property(relation, "ForeignColumn", file.storage_path)?;
        inputs.push(RelationInput {
            metadata_path: file.storage_path.to_owned(),
            relationship_name,
            generated_relationship_id,
            containing_folder,
            containing_table,
            primary_table,
            primary_column,
            foreign_column,
        });
        let _ = containing_folder;
    }

    let mut output = Vec::new();
    reserve_vec(&mut output, inputs.len(), "relationship bindings")?;
    let mut seen_metadata = HashSet::new();
    reserve_set(
        &mut seen_metadata,
        inputs.len(),
        "relationship metadata set",
    )?;
    let mut seen_dimension_relationships = HashSet::new();
    reserve_set(
        &mut seen_dimension_relationships,
        inputs.len(),
        "Dimension relationship set",
    )?;
    for input in &inputs {
        let containing = resolve_table_id(
            table_bindings,
            &input.containing_table,
            &input.metadata_path,
        )?;
        let (table_folder, _) = metadata_table_path(containing.metadata_path.as_str())?;
        if table_folder != input.containing_folder {
            return Err(unproven(
                &input.metadata_path,
                "relationship metadata folder does not match its table metadata owner",
            ));
        }
        // PrimaryTable is the XML metadata table name (MS-XLDM
        // XMRelationshipPropertiesType), not the generated TableID taken
        // from the containing .dim folder. Keep those namespaces separate.
        let primary = resolve_table_metadata_name(
            table_bindings,
            &input.primary_table,
            &input.metadata_path,
        )?;
        ensure_table_column(
            primary,
            &input.primary_column,
            &input.metadata_path,
            "PrimaryColumn",
        )?;
        ensure_table_column(
            containing,
            &input.foreign_column,
            &input.metadata_path,
            "ForeignColumn",
        )?;
        let dimension = definitions
            .iter()
            .find(|value| {
                value.kind == OlapObjectKind::Dimension && value.id == primary.dimension_id
            })
            .ok_or_else(|| {
                unproven(
                    &input.metadata_path,
                    "primary table has no Dimension definition",
                )
            })?;
        let dimension_def = definition(model, dimension)?;
        let relation_collection = relationship_collection(
            &dimension_def.object,
            dimension.path.as_str(),
            unknown_members,
        )?
        .ok_or_else(|| {
            unproven(
                &input.metadata_path,
                "primary Dimension has no Relationships collection",
            )
        })?;
        if !physical_paths.contains_key(input.metadata_path.as_str()) {
            return Err(unproven(
                input.metadata_path.as_str(),
                "relationship metadata is absent from the XLDM virtual directory",
            ));
        }
        require_group_member(
            file_group_bindings,
            containing.dimension_path.as_str(),
            input.metadata_path.as_str(),
        )?;
        let mut candidates = Vec::new();
        for item in relation_collection
            .children
            .iter()
            .filter(|item| item.name == "Relationship")
        {
            if let Some(reference) = relationship_reference(item, &input.metadata_path)? {
                // The Dimension relationship edge is keyed by the RelId in
                // the generated relationship filename. The XMRelationship
                // `name` attribute is retained metadata and must not be
                // accepted as a substitute for that storage identity.
                if reference.value == input.generated_relationship_id {
                    candidates.push(reference);
                }
            }
        }
        if candidates.len() != 1 {
            return Err(unproven(
                input.metadata_path.as_str(),
                "Dimension relationship is missing or ambiguous",
            ));
        }
        let relation_reference = candidates.remove(0);
        if !seen_dimension_relationships
            .insert((dimension.path.clone(), relation_reference.value.clone()))
        {
            return Err(invalid(
                &input.metadata_path,
                "duplicate Dimension relationship binding",
            ));
        }
        let index_paths = relationship_index_paths(
            input,
            containing.dimension_path.as_str(),
            physical_paths,
            file_group_bindings,
        )?;
        output.push(OlapRelationshipBinding {
            metadata_path: input.metadata_path.clone(),
            relationship_name: input.relationship_name.clone(),
            generated_relationship_id: input.generated_relationship_id.clone(),
            containing_table: input.containing_table.clone(),
            primary_table: input.primary_table.clone(),
            primary_column: input.primary_column.clone(),
            foreign_column: input.foreign_column.clone(),
            dimension_path: dimension.path.clone(),
            dimension_reference: relation_reference,
            relationship_index_paths: index_paths,
        });
        seen_metadata.insert(input.metadata_path.as_str());
    }
    for dimension in definitions
        .iter()
        .filter(|value| value.kind == OlapObjectKind::Dimension)
    {
        let definition = definition(model, dimension)?;
        let Some(collection) =
            relationship_collection(&definition.object, dimension.path.as_str(), unknown_members)?
        else {
            continue;
        };
        for item in collection
            .children
            .iter()
            .filter(|item| item.name == "Relationship")
        {
            let Some(reference) = relationship_reference(item, dimension.path.as_str())? else {
                return Err(unproven(
                    dimension.path.as_str(),
                    "Dimension relationship has no explicit identity",
                ));
            };
            if !seen_dimension_relationships
                .iter()
                .any(|(path, value)| path == &dimension.path && value == &reference.value)
            {
                push_unknown(
                    unknown_members,
                    OlapUnknownMember {
                        path: dimension.path.clone(),
                        reason: format!(
                            "Dimension relationship {} has no table relationship metadata owner",
                            reference.value
                        ),
                    },
                )?;
            }
        }
    }
    for file in metadata
        .files
        .iter()
        .filter(|file| file.kind == MetadataFileKind::TableRelationship)
    {
        if !seen_metadata.contains(file.storage_path) {
            push_unknown(
                unknown_members,
                OlapUnknownMember {
                    path: file.storage_path.to_owned(),
                    reason: "relationship metadata was not assigned to a Dimension".into(),
                },
            )?;
        }
    }
    Ok(output)
}

fn metadata_table_path_for_relationship(path: &str) -> OlapProofResult<(String, String, String)> {
    let (folder, name) = path
        .rsplit_once('/')
        .ok_or_else(|| invalid(path, "relationship metadata path has no folder"))?;
    let folder_name = folder
        .rsplit('/')
        .next()
        .ok_or_else(|| invalid(path, "relationship metadata path has no dimension folder"))?;
    let Some(prefix) = folder_name.strip_suffix(".dim") else {
        return Err(invalid(
            path,
            "relationship metadata is outside a .dim folder",
        ));
    };
    let Some((table_id, version)) = prefix.rsplit_once('.') else {
        return Err(invalid(
            path,
            "relationship dimension folder has no version",
        ));
    };
    if version != "0" || table_id.is_empty() {
        return Err(invalid(
            path,
            "relationship metadata owner is not a .0.dim folder",
        ));
    }
    let Some(prefix) = name.strip_suffix(".tbl.xml") else {
        return Err(invalid(
            path,
            "relationship metadata filename is not R$...tbl.xml",
        ));
    };
    let Some((name_prefix, version)) = prefix.rsplit_once('.') else {
        return Err(invalid(
            path,
            "relationship metadata filename has no version",
        ));
    };
    if !version.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid(
            path,
            "relationship metadata filename has an invalid version",
        ));
    }
    let Some(name_prefix) = name_prefix.strip_prefix("R$") else {
        return Err(invalid(
            path,
            "relationship metadata filename is not R$...tbl.xml",
        ));
    };
    let Some((file_table_id, relationship_id)) = name_prefix.split_once('$') else {
        return Err(invalid(path, "relationship metadata filename has no RelId"));
    };
    if file_table_id != table_id || !valid_user_id(relationship_id) {
        return Err(invalid(
            path,
            "relationship metadata filename does not belong to its dimension folder",
        ));
    }
    Ok((
        folder.to_owned(),
        table_id.to_owned(),
        relationship_id.to_owned(),
    ))
}

fn required_property(object: &MetadataObject, name: &str, path: &str) -> OlapProofResult<String> {
    object
        .property(name)
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
        .ok_or_else(|| unproven(path, format!("relationship is missing {name}")))
}

fn resolve_table_metadata_name<'a>(
    tables: &'a [OlapTableBinding],
    value: &str,
    path: &str,
) -> OlapProofResult<&'a OlapTableBinding> {
    let mut found = tables
        .iter()
        .filter(|table| table.metadata_name.as_deref() == Some(value));
    let Some(table) = found.next() else {
        return Err(unproven(
            path,
            "explicit relationship table reference has no matching table",
        ));
    };
    if found.next().is_some() {
        return Err(unproven(
            path,
            "explicit relationship table reference is ambiguous",
        ));
    }
    Ok(table)
}

fn resolve_table_id<'a>(
    tables: &'a [OlapTableBinding],
    value: &str,
    path: &str,
) -> OlapProofResult<&'a OlapTableBinding> {
    let mut found = tables.iter().filter(|table| table.table_id == value);
    let Some(table) = found.next() else {
        return Err(unproven(
            path,
            "relationship metadata containing folder has no matching table",
        ));
    };
    if found.next().is_some() {
        return Err(unproven(
            path,
            "relationship metadata containing folder matches multiple tables",
        ));
    }
    Ok(table)
}

fn ensure_table_column(
    table: &OlapTableBinding,
    value: &str,
    path: &str,
    field: &str,
) -> OlapProofResult<()> {
    if table.attribute_ids.iter().any(|column| column == value) {
        Ok(())
    } else {
        Err(unproven(
            path,
            format!(
                "{field} {value} is not a column of table {}",
                table.table_id
            ),
        ))
    }
}

fn relationship_reference(
    item: &OlapElement,
    path: &str,
) -> OlapProofResult<Option<OlapReference>> {
    let fields = [
        ("RelationshipID", OlapReferenceField::RelationshipId),
        ("ID", OlapReferenceField::Id),
        ("Name", OlapReferenceField::RelationshipName),
    ];
    let mut found = None;
    for (name, field) in fields {
        if let Some(value) = optional_scalar(item, name, path)? {
            if value.is_empty() {
                return Err(invalid(path, "relationship reference is empty"));
            }
            if found.is_some() {
                return Err(unproven(path, "relationship has multiple identity fields"));
            }
            found = Some(OlapReference {
                field,
                value: value.to_owned(),
            });
        }
    }
    Ok(found)
}

fn relationship_index_paths(
    input: &RelationInput,
    containing_dimension_path: &str,
    physical_paths: &HashMap<&str, usize>,
    file_group_bindings: &[OlapFileGroupBinding],
) -> OlapProofResult<Vec<String>> {
    let key = [
        "R$",
        input.containing_table.as_str(),
        "$",
        input.generated_relationship_id.as_str(),
    ]
    .concat();
    let group = file_group_bindings
        .iter()
        .find(|group| group.definition_path == containing_dimension_path)
        .ok_or_else(|| {
            unproven(
                &input.metadata_path,
                "containing table has no proven Dimension FileGroup",
            )
        })?;
    let mut output = Vec::new();
    reserve_vec(&mut output, 1, "relationship index paths")?;
    for member in &group.members {
        let name = member.rsplit('/').next().unwrap_or(member.as_str());
        let parent = member.rsplit_once('/').map_or("", |parts| parts.0);
        if parent != input.containing_folder || !is_relationship_index_path(name) {
            continue;
        }
        let Some(suffix) = name.strip_suffix(".idf") else {
            continue;
        };
        let Some((_version, rest)) = suffix.split_once('.') else {
            continue;
        };
        if rest == format!("{key}.INDEX.0") {
            if !physical_paths.contains_key(member.as_str()) {
                return Err(unproven(
                    &input.metadata_path,
                    "relationship index is absent from the XLDM virtual directory",
                ));
            }
            if !output.is_empty() {
                return Err(unproven(
                    &input.metadata_path,
                    "relationship has multiple physical index members",
                ));
            }
            output.push(member.clone());
        }
    }
    if output.len() != 1 {
        return Err(unproven(
            &input.metadata_path,
            "relationship has zero or multiple physical index members",
        ));
    }
    Ok(output)
}

fn prove_cube(
    model: &OlapModel<'_>,
    definitions: &[DefinitionRef],
    tables: &[OlapTableBinding],
    unknown_members: &mut Vec<OlapUnknownMember>,
) -> OlapProofResult<Option<OlapCubeBinding>> {
    let Some(cube_ref) = definitions
        .iter()
        .find(|value| value.kind == OlapObjectKind::Cube)
    else {
        if tables.is_empty() {
            return Ok(None);
        }
        return Err(unproven("Cube", "every tabular model must have a Cube"));
    };
    let cube = definition(model, cube_ref)?;
    let Some(dimensions) = cube.object.child("Dimensions") else {
        return Err(unproven(
            cube_ref.path.as_str(),
            "Cube has no Dimensions collection",
        ));
    };
    let mut bindings = Vec::new();
    reserve_vec(&mut bindings, dimensions.children.len(), "Cube Dimensions")?;
    let mut known_ids = HashSet::new();
    reserve_set(&mut known_ids, tables.len(), "Cube Dimension index")?;
    for table in tables {
        known_ids.insert(table.dimension_id.as_str());
    }
    let mut seen_dimensions = HashSet::new();
    reserve_set(
        &mut seen_dimensions,
        dimensions.children.len(),
        "Cube Dimension references",
    )?;
    for item in &dimensions.children {
        if item.name != "Dimension" {
            return Err(invalid(
                cube_ref.path.as_str(),
                "Cube Dimensions contains a non-Dimension member",
            ));
        }
        require_type(item, "CubeDimension", cube_ref.path.as_str())?;
        let reference = reference_from(
            item,
            &["DimensionID", "ID"],
            OlapReferenceField::DimensionId,
            cube_ref.path.as_str(),
        )?;
        if !seen_dimensions.insert(reference.value.clone()) {
            return Err(invalid(
                cube_ref.path.as_str(),
                "Cube Dimensions contains a duplicate Dimension reference",
            ));
        }
        let Some(attributes) = item.child("Attributes") else {
            return Err(unproven(
                cube_ref.path.as_str(),
                "Cube Dimension has no Attributes collection",
            ));
        };
        let attribute_references =
            collect_attribute_references(attributes, "CubeAttribute", cube_ref.path.as_str())?;
        let unknown = !known_ids.contains(reference.value.as_str());
        if unknown {
            push_unknown(
                unknown_members,
                OlapUnknownMember {
                    path: cube_ref.path.clone(),
                    reason: format!(
                        "Cube Dimension references unlinked Dimension {}",
                        reference.value
                    ),
                },
            )?;
        }
        if let Some(table) = tables
            .iter()
            .find(|table| table.dimension_id == reference.value)
        {
            ensure_attribute_set(
                &table.attribute_ids,
                &attribute_references,
                cube_ref.path.as_str(),
                "Cube Dimension Attributes",
            )?;
        }
        bindings.push(OlapCubeDimensionBinding {
            reference,
            attribute_references,
            unknown,
        });
    }
    for table in tables {
        if !bindings
            .iter()
            .any(|value| value.reference.value == table.dimension_id)
        {
            return Err(unproven(
                cube_ref.path.as_str(),
                "a table Dimension is absent from Cube Dimensions",
            ));
        }
    }
    let folder = persist_folder(cube_ref.path.as_str(), cube)?;
    let measure_group_paths = qualify_file_list_vec(
        &folder,
        &cube.extension.measure_group_files,
        cube_ref.path.as_str(),
    )?;
    for path in &measure_group_paths {
        if !definitions
            .iter()
            .any(|value| value.kind == OlapObjectKind::MeasureGroup && value.path == *path)
        {
            push_unknown(
                unknown_members,
                OlapUnknownMember {
                    path: path.clone(),
                    reason: "Cube MeasureGroupFileList member has no matching definition".into(),
                },
            )?;
        }
    }
    Ok(Some(OlapCubeBinding {
        path: cube_ref.path.clone(),
        id: cube_ref.id.clone(),
        dimensions: bindings,
        measure_group_paths,
    }))
}

fn prove_measure_groups(
    model: &OlapModel<'_>,
    definitions: &[DefinitionRef],
    tables: &[OlapTableBinding],
    relationships: &[OlapRelationshipBinding],
    unknown_members: &mut Vec<OlapUnknownMember>,
) -> OlapProofResult<Vec<OlapMeasureGroupBinding>> {
    let mut output = Vec::new();
    let measure_group_count = definitions
        .iter()
        .filter(|value| value.kind == OlapObjectKind::MeasureGroup)
        .count();
    reserve_vec(&mut output, measure_group_count, "measure-group bindings")?;
    for group_ref in definitions
        .iter()
        .filter(|value| value.kind == OlapObjectKind::MeasureGroup)
    {
        let group = definition(model, group_ref)?;
        let Some(dimensions) = group.object.child("Dimensions") else {
            return Err(unproven(
                group_ref.path.as_str(),
                "MeasureGroup has no Dimensions collection",
            ));
        };
        let mut dimension_bindings = Vec::new();
        reserve_vec(
            &mut dimension_bindings,
            dimensions.children.len(),
            "measure-group dimensions",
        )?;
        let mut seen_dimensions = HashSet::new();
        reserve_set(
            &mut seen_dimensions,
            dimensions.children.len(),
            "measure-group dimension references",
        )?;
        let mut table_dimension_id = None;
        for item in &dimensions.children {
            if item.name != "Dimension" {
                return Err(invalid(
                    group_ref.path.as_str(),
                    "MeasureGroup Dimensions contains a non-Dimension member",
                ));
            }
            let kind = match element_type(item) {
                Some("DegenerateMeasureGroupDimension") => MeasureGroupDimensionKind::Degenerate,
                Some("ReferenceMeasureGroupDimension") => MeasureGroupDimensionKind::Reference,
                Some(_) => {
                    push_unknown(
                        unknown_members,
                        OlapUnknownMember {
                            path: group_ref.path.clone(),
                            reason: "MeasureGroup contains an unrecognized Dimension type".into(),
                        },
                    )?;
                    continue;
                },
                None => {
                    return Err(unproven(
                        group_ref.path.as_str(),
                        "MeasureGroup Dimension has no xsi:type",
                    ));
                },
            };
            let reference = reference_from(
                item,
                &["CubeDimensionID", "DimensionID", "ID"],
                OlapReferenceField::CubeDimensionId,
                group_ref.path.as_str(),
            )?;
            if !seen_dimensions.insert((kind, reference.value.clone())) {
                return Err(invalid(
                    group_ref.path.as_str(),
                    "MeasureGroup Dimensions contains a duplicate Dimension reference",
                ));
            }
            let attributes = item.child("Attributes").ok_or_else(|| {
                unproven(
                    group_ref.path.as_str(),
                    "MeasureGroup Dimension has no Attributes collection",
                )
            })?;
            let attribute_references = collect_attribute_references(
                attributes,
                "MeasureGroupDimensionAttribute",
                group_ref.path.as_str(),
            )?;
            if kind == MeasureGroupDimensionKind::Degenerate
                && table_dimension_id
                    .replace(reference.value.clone())
                    .is_some()
            {
                return Err(unproven(
                    group_ref.path.as_str(),
                    "MeasureGroup has multiple degenerate dimensions",
                ));
            }
            let unknown = !tables
                .iter()
                .any(|table| table.dimension_id == reference.value);
            if unknown {
                push_unknown(
                    unknown_members,
                    OlapUnknownMember {
                        path: group_ref.path.clone(),
                        reason: format!(
                            "MeasureGroup Dimension references unlinked Dimension {}",
                            reference.value
                        ),
                    },
                )?;
            }
            if let Some(table) = tables
                .iter()
                .find(|table| table.dimension_id == reference.value)
            {
                ensure_attribute_set(
                    &table.attribute_ids,
                    &attribute_references,
                    group_ref.path.as_str(),
                    "MeasureGroup Dimension Attributes",
                )?;
            }
            dimension_bindings.push(OlapMeasureGroupDimensionBinding {
                kind,
                dimension_reference: reference,
                attribute_references,
                unknown,
            });
        }
        let Some(table_dimension_id) = table_dimension_id else {
            return Err(unproven(
                group_ref.path.as_str(),
                "MeasureGroup has no degenerate table Dimension",
            ));
        };
        let folder = persist_folder(group_ref.path.as_str(), group)?;
        let partition_paths = qualify_file_list_vec(
            &folder,
            &group.extension.partition_files,
            group_ref.path.as_str(),
        )?;
        if partition_paths.is_empty() {
            return Err(unproven(
                group_ref.path.as_str(),
                "MeasureGroup has no PartitionFileList members",
            ));
        }
        for path in &partition_paths {
            if !definitions
                .iter()
                .any(|value| value.kind == OlapObjectKind::Partition && value.path == *path)
            {
                push_unknown(
                    unknown_members,
                    OlapUnknownMember {
                        path: path.clone(),
                        reason: "MeasureGroup PartitionFileList member has no matching Partition"
                            .into(),
                    },
                )?;
            }
        }
        if let Some(table) = tables
            .iter()
            .find(|table| table.dimension_id == table_dimension_id)
        {
            for relation in relationships.iter().filter(|relation| {
                relation.primary_table == table.metadata_name.as_deref().unwrap_or("")
            }) {
                let Some(related) = tables
                    .iter()
                    .find(|candidate| candidate.table_id == relation.containing_table)
                else {
                    push_unknown(
                        unknown_members,
                        OlapUnknownMember {
                            path: group_ref.path.clone(),
                            reason: format!(
                                "relationship {} has an unresolved foreign table",
                                relation.metadata_path
                            ),
                        },
                    )?;
                    continue;
                };
                if !dimension_bindings.iter().any(|dimension| {
                    dimension.kind == MeasureGroupDimensionKind::Reference
                        && dimension.dimension_reference.value == related.dimension_id
                }) {
                    push_unknown(
                        unknown_members,
                        OlapUnknownMember {
                            path: group_ref.path.clone(),
                            reason: format!(
                                "MeasureGroup is missing a reference Dimension for {}",
                                related.dimension_id
                            ),
                        },
                    )?;
                }
            }
        }
        output.push(OlapMeasureGroupBinding {
            path: group_ref.path.clone(),
            id: group_ref.id.clone(),
            object_id: group_ref.object_id.clone(),
            table_dimension_id,
            dimensions: dimension_bindings,
            partition_paths,
        });
    }
    for table in tables {
        let mut owners = output
            .iter()
            .filter(|group| group.table_dimension_id == table.dimension_id);
        if owners.next().is_none() {
            return Err(unproven(
                table.dimension_path.as_str(),
                "table has no measure group",
            ));
        }
        if owners.next().is_some() {
            return Err(unproven(
                table.dimension_path.as_str(),
                "table has multiple measure groups",
            ));
        }
    }
    Ok(output)
}

fn prove_partitions(
    model: &OlapModel<'_>,
    definitions: &[DefinitionRef],
    measure_groups: &[OlapMeasureGroupBinding],
    unknown_members: &mut Vec<OlapUnknownMember>,
) -> OlapProofResult<Vec<OlapPartitionBinding>> {
    let mut output = Vec::new();
    let partition_count = definitions
        .iter()
        .filter(|value| value.kind == OlapObjectKind::Partition)
        .count();
    reserve_vec(&mut output, partition_count, "partition bindings")?;
    for partition_ref in definitions
        .iter()
        .filter(|value| value.kind == OlapObjectKind::Partition)
    {
        let partition = definition(model, partition_ref)?;
        let mut owners = measure_groups.iter().filter(|group| {
            group
                .partition_paths
                .iter()
                .any(|path| path == &partition_ref.path)
        });
        let Some(owner) = owners.next() else {
            push_unknown(
                unknown_members,
                OlapUnknownMember {
                    path: partition_ref.path.clone(),
                    reason: "Partition is not a member of any MeasureGroup PartitionFileList"
                        .into(),
                },
            )?;
            continue;
        };
        if owners.next().is_some() {
            return Err(unproven(
                partition_ref.path.as_str(),
                "Partition belongs to multiple measure groups",
            ));
        }
        let folder = persist_folder(partition_ref.path.as_str(), partition)?;
        let data_files = qualify_file_list_vec(
            &folder,
            &partition.extension.data_files,
            partition_ref.path.as_str(),
        )?;
        output.push(OlapPartitionBinding {
            path: partition_ref.path.clone(),
            id: partition_ref.id.clone(),
            object_id: partition_ref.object_id.clone(),
            measure_group_path: owner.path.clone(),
            data_files,
        });
    }
    if output.is_empty() {
        return Err(unproven(
            "Partition",
            "no partition has a proven measure-group owner",
        ));
    }
    Ok(output)
}

fn validate_cube_file_lists(
    cube: Option<&OlapCubeBinding>,
    measure_groups: &[OlapMeasureGroupBinding],
) -> OlapProofResult<()> {
    let Some(cube) = cube else {
        return Ok(());
    };
    let mut expected = HashSet::new();
    reserve_set(
        &mut expected,
        measure_groups.len(),
        "Cube measure-group set",
    )?;
    for group in measure_groups {
        expected.insert(group.path.as_str());
    }
    let mut actual = HashSet::new();
    reserve_set(
        &mut actual,
        cube.measure_group_paths.len(),
        "Cube file-list set",
    )?;
    for path in &cube.measure_group_paths {
        if !actual.insert(path.as_str()) {
            return Err(invalid(
                &cube.path,
                "duplicate Cube MeasureGroupFileList member",
            ));
        }
    }
    if !expected.is_subset(&actual) {
        return Err(unproven(
            &cube.path,
            "Cube MeasureGroupFileList is missing a proven measure-group definition",
        ));
    }
    Ok(())
}

fn validate_measure_group_file_lists(
    groups: &[OlapMeasureGroupBinding],
    partitions: &[OlapPartitionBinding],
) -> OlapProofResult<()> {
    for group in groups {
        let mut expected = HashSet::new();
        let expected_count = partitions
            .iter()
            .filter(|partition| partition.measure_group_path == group.path)
            .count();
        reserve_set(&mut expected, expected_count, "measure-group partition set")?;
        for partition in partitions
            .iter()
            .filter(|partition| partition.measure_group_path == group.path)
        {
            expected.insert(partition.path.as_str());
        }
        let mut actual = HashSet::new();
        reserve_set(
            &mut actual,
            group.partition_paths.len(),
            "partition file-list set",
        )?;
        for path in &group.partition_paths {
            if !actual.insert(path.as_str()) {
                return Err(invalid(&group.path, "duplicate PartitionFileList member"));
            }
        }
        if !expected.is_subset(&actual) {
            return Err(unproven(
                &group.path,
                "PartitionFileList is missing a proven partition definition",
            ));
        }
    }
    Ok(())
}

fn validate_partition_ownership(
    groups: &[OlapMeasureGroupBinding],
    partitions: &[OlapPartitionBinding],
) -> OlapProofResult<()> {
    for partition in partitions {
        if !groups
            .iter()
            .any(|group| group.path == partition.measure_group_path)
        {
            return Err(unproven(
                &partition.path,
                "partition owner disappeared during proof",
            ));
        }
    }
    Ok(())
}

fn record_unlinked_members(
    storage: &Storage<'_>,
    metadata: &MetadataModel<'_>,
    model: &OlapModel<'_>,
    groups: &[OlapFileGroupBinding],
    unknown_members: &mut Vec<OlapUnknownMember>,
) -> OlapProofResult<()> {
    for group in groups {
        for member in &group.members {
            let known = group.data_files.iter().any(|path| path == member)
                || model.files.iter().any(|file| file.storage_path == member)
                || metadata
                    .files
                    .iter()
                    .any(|file| file.storage_path == member);
            if !known
                && !unknown_members
                    .iter()
                    .any(|unknown| unknown.path == *member)
            {
                push_unknown(
                    unknown_members,
                    OlapUnknownMember {
                        path: member.clone(),
                        reason: "physical FileGroup member has no proven OLAP or metadata owner"
                            .into(),
                    },
                )?;
            }
        }
    }
    for file in &storage.files {
        let grouped = groups
            .iter()
            .any(|group| group.members.iter().any(|member| member == &file.path));
        if !grouped
            && file.path.contains('/')
            && !unknown_members
                .iter()
                .any(|unknown| unknown.path == file.path)
        {
            push_unknown(
                unknown_members,
                OlapUnknownMember {
                    path: file.path.clone(),
                    reason: "physical member is absent from every proven FileGroup".into(),
                },
            )?;
        }
    }
    for file in &metadata.files {
        let grouped = groups.iter().any(|group| {
            group
                .members
                .iter()
                .any(|member| member == file.storage_path)
        });
        if !grouped
            && !unknown_members
                .iter()
                .any(|unknown| unknown.path == file.storage_path)
        {
            push_unknown(
                unknown_members,
                OlapUnknownMember {
                    path: file.storage_path.to_owned(),
                    reason: "metadata member is absent from every proven FileGroup".into(),
                },
            )?;
        }
    }
    for file in &model.files {
        let grouped = groups.iter().any(|group| {
            group
                .members
                .iter()
                .any(|member| member == file.storage_path)
        });
        if !grouped
            && !unknown_members
                .iter()
                .any(|unknown| unknown.path == file.storage_path)
        {
            push_unknown(
                unknown_members,
                OlapUnknownMember {
                    path: file.storage_path.to_owned(),
                    reason: "OLAP member is absent from every proven FileGroup".into(),
                },
            )?;
        }
    }
    Ok(())
}

fn collect_attribute_references(
    collection: &OlapElement,
    expected_type: &str,
    path: &str,
) -> OlapProofResult<Vec<OlapReference>> {
    if !collection.text.trim().is_empty() {
        return Err(invalid(path, "Attributes collection contains text"));
    }
    let mut output = Vec::new();
    reserve_vec(
        &mut output,
        collection.children.len(),
        "attribute references",
    )?;
    for item in &collection.children {
        if item.name != "Attribute" {
            return Err(invalid(
                path,
                "Attributes collection contains a non-Attribute member",
            ));
        }
        require_type(item, expected_type, path)?;
        let reference = reference_from(
            item,
            &["AttributeID", "ID"],
            OlapReferenceField::AttributeId,
            path,
        )?;
        output.push(reference);
    }
    Ok(output)
}

fn ensure_attribute_set(
    expected: &[String],
    actual: &[OlapReference],
    path: &str,
    label: &str,
) -> OlapProofResult<()> {
    let mut expected_set = HashSet::new();
    reserve_set(&mut expected_set, expected.len(), "expected attribute set")?;
    for value in expected {
        expected_set.insert(value.as_str());
    }
    let mut actual_set = HashSet::new();
    reserve_set(&mut actual_set, actual.len(), "actual attribute set")?;
    for value in actual {
        actual_set.insert(value.value.as_str());
    }
    if expected_set.len() != expected.len() || actual_set.len() != actual.len() {
        return Err(invalid(path, format!("{label} contains duplicate IDs")));
    }
    if expected_set != actual_set {
        return Err(unproven(
            path,
            format!("{label} does not exactly cover the proven table columns"),
        ));
    }
    Ok(())
}

fn reference_from(
    item: &OlapElement,
    names: &[&str],
    fallback: OlapReferenceField,
    path: &str,
) -> OlapProofResult<OlapReference> {
    let mut found = None;
    for name in names {
        if let Some(value) = optional_scalar(item, name, path)? {
            if value.is_empty() {
                return Err(invalid(path, "explicit OLAP reference is empty"));
            }
            if found.is_some() {
                return Err(unproven(
                    path,
                    "OLAP reference has multiple candidate fields",
                ));
            }
            let field = match *name {
                "DimensionID" => OlapReferenceField::DimensionId,
                "CubeDimensionID" => OlapReferenceField::CubeDimensionId,
                "AttributeID" => OlapReferenceField::AttributeId,
                _ => fallback,
            };
            found = Some(OlapReference {
                field,
                value: value.to_owned(),
            });
        }
    }
    found.ok_or_else(|| unproven(path, "required explicit OLAP reference field is missing"))
}

fn require_type(item: &OlapElement, expected: &str, path: &str) -> OlapProofResult<()> {
    if element_type(item) != Some(expected) {
        return Err(unproven(path, format!("expected xsi:type {expected}")));
    }
    Ok(())
}

fn element_type(item: &OlapElement) -> Option<&str> {
    item.attributes.iter().find_map(|(name, value)| {
        (name == "xsi:type" || name == "type" || name.ends_with(":type")).then_some(value.as_str())
    })
}

fn required_scalar<'a>(
    object: &'a OlapElement,
    name: &str,
    path: &str,
) -> OlapProofResult<&'a str> {
    let mut found = None;
    for child in object.children.iter().filter(|child| child.name == name) {
        if found.is_some() {
            return Err(invalid(path, format!("duplicate {name} element")));
        }
        if !child.children.is_empty() {
            return Err(invalid(path, format!("{name} is not scalar")));
        }
        found = Some(child.text.trim());
    }
    found
        .filter(|value| !value.is_empty())
        .ok_or_else(|| unproven(path, format!("missing {name}")))
}

fn optional_scalar<'a>(
    object: &'a OlapElement,
    name: &str,
    path: &str,
) -> OlapProofResult<Option<&'a str>> {
    let mut found = None;
    for child in object.children.iter().filter(|child| child.name == name) {
        if found.is_some() {
            return Err(invalid(path, format!("duplicate {name} element")));
        }
        if !child.children.is_empty() {
            return Err(invalid(path, format!("{name} is not scalar")));
        }
        found = Some(child.text.trim());
    }
    Ok(found)
}

fn qualify_file_list_vec(
    folder: &str,
    values: &[String],
    path: &str,
) -> OlapProofResult<Vec<String>> {
    let mut result = Vec::new();
    reserve_vec(&mut result, values.len(), "OLAP file list")?;
    for value in values {
        let qualified = qualify_file_list_item(folder, value, path)?;
        result.push(qualified);
        let last = result.len() - 1;
        if result[..last]
            .iter()
            .any(|existing| existing == &result[last])
        {
            return Err(invalid(path, "duplicate OLAP file-list member"));
        }
    }
    Ok(result)
}

fn qualify_file_list_item(folder: &str, item: &str, path: &str) -> OlapProofResult<String> {
    if item.is_empty()
        || item.starts_with('/')
        || item.contains('\\')
        || item
            .split('/')
            .any(|part| part.is_empty() || part == "." || part == "..")
    {
        return Err(invalid(path, "unsafe OLAP file-list path"));
    }
    Ok(if item.contains('/') || folder.is_empty() {
        item.to_owned()
    } else {
        format!("{folder}/{item}")
    })
}

fn persist_folder(path: &str, definition: &OlapDefinition) -> OlapProofResult<String> {
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
        .strip_suffix(definition.kind.suffix_for_proof())
        .ok_or_else(|| invalid(path, "definition suffix changed"))?;
    let id = base
        .rsplit_once('.')
        .ok_or_else(|| invalid(path, "definition has no generated version"))?
        .0;
    let object_suffix = definition.kind.suffix_for_proof().trim_end_matches(".xml");
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

fn decimal_len(value: i32) -> usize {
    let mut magnitude = value.unsigned_abs();
    let mut length = 1;
    while magnitude >= 10 {
        magnitude /= 10;
        length += 1;
    }
    if value.is_negative() {
        length + 1
    } else {
        length
    }
}

/// Section 2.2.3.7.2.2 fixes the physical relationship-index member name to
/// the `.idf` form. Section 2.4.3.1 permits the payload of that member to use
/// either the ordinary `.idf` layout or the sparse hash-index layout; the
/// payload choice does not change the physical name.
#[must_use]
pub fn is_relationship_index_path(path: &str) -> bool {
    let name = path.rsplit('/').next().unwrap_or(path);
    let Some(stem) = name.strip_suffix(".idf") else {
        return false;
    };
    let Some((version, rest)) = stem.split_once('.') else {
        return false;
    };
    if !version.bytes().all(|byte| byte.is_ascii_digit())
        || !rest.starts_with("R$")
        || !rest.ends_with(".INDEX.0")
    {
        return false;
    }
    let Some(identifiers) = rest
        .strip_prefix("R$")
        .and_then(|value| value.strip_suffix(".INDEX.0"))
    else {
        return false;
    };
    let mut count = 0;
    for value in identifiers.split('$') {
        count += 1;
        if !valid_user_id(value) {
            return false;
        }
    }
    count >= 2
}

fn valid_user_id(value: &str) -> bool {
    !value.is_empty()
        && value.bytes().all(|byte| {
            byte.is_ascii_alphanumeric()
                || (b'#'..=b'.').contains(&byte)
                || matches!(
                    byte,
                    b'!' | b'=' | b'@' | b'[' | b']' | b'^' | b'{' | b'}' | b'~'
                )
        })
}

fn preflight_storage(storage: &Storage<'_>, budget: &mut Budget) -> OlapProofResult<()> {
    for file in &storage.files {
        budget.item("physical files")?;
        budget.string(&file.path, "physical paths")?;
    }
    for group in &storage.backup_log.file_groups {
        budget.item("FileGroups")?;
        budget.string(&group.id, "FileGroup IDs")?;
        budget.string(&group.name, "FileGroup names")?;
        budget.string(&group.persist_location_path, "FileGroup paths")?;
        budget.string(&group.storage_location_path, "FileGroup paths")?;
        budget.string(&group.object_id, "FileGroup ObjectIDs")?;
        for file in &group.files {
            budget.item("logged files")?;
            budget.string(&file.source_path, "logged source paths")?;
            budget.string(&file.storage_path, "logged storage paths")?;
            budget.string(&file.generated.normalized_path, "generated paths")?;
        }
    }
    budget.string(&storage.backup_log.server_root, "backup-log text")?;
    budget.string(&storage.backup_log.object_name, "backup-log text")?;
    budget.string(&storage.backup_log.object_id, "backup-log text")?;
    for collation in &storage.backup_log.collations {
        budget.item("collations")?;
        budget.string(collation, "collations")?;
    }
    Ok(())
}

fn preflight_metadata(metadata: &MetadataModel<'_>, budget: &mut Budget) -> OlapProofResult<()> {
    for file in &metadata.files {
        budget.item("metadata files")?;
        budget.string(file.storage_path, "metadata paths")?;
        preflight_metadata_object(&file.table, budget)?;
    }
    for column in &metadata.columns {
        budget.item("metadata policies")?;
        budget.string(&column.name, "metadata policy text")?;
        budget.string(&column.data_file, "metadata policy paths")?;
        if let Some(dictionary) = &column.dictionary {
            budget.item("metadata policies")?;
            budget.string(&dictionary.storage_name, "metadata policy paths")?;
            budget.string(dictionary.class.as_str(), "metadata policy text")?;
        }
    }
    for relationship in &metadata.relationships {
        budget.item("metadata relationships")?;
        if let Some(name) = &relationship.name {
            budget.string(name, "metadata relationship names")?;
        }
        budget.string(&relationship.primary_table, "metadata relationship names")?;
        budget.string(&relationship.primary_column, "metadata relationship names")?;
        budget.string(&relationship.foreign_column, "metadata relationship names")?;
    }
    Ok(())
}

fn preflight_metadata_object(object: &MetadataObject, budget: &mut Budget) -> OlapProofResult<()> {
    budget.item("metadata objects")?;
    budget.string(object.class.as_str(), "metadata object text")?;
    if let Some(name) = &object.name {
        budget.string(name, "metadata object names")?;
    }
    for property in &object.properties {
        budget.item("metadata properties")?;
        budget.string(&property.name, "metadata property text")?;
        budget.string(&property.value, "metadata property text")?;
    }
    for member in &object.members {
        budget.item("metadata members")?;
        budget.string(&member.name, "metadata member names")?;
        preflight_metadata_object(&member.object, budget)?;
    }
    for collection in &object.collections {
        budget.item("metadata collections")?;
        budget.string(&collection.name, "metadata collection names")?;
        for child in &collection.objects {
            preflight_metadata_object(child, budget)?;
        }
    }
    for data in &object.data_objects {
        budget.item("metadata data objects")?;
        preflight_metadata_object(&data.object, budget)?;
    }
    Ok(())
}

fn preflight_olap(model: &OlapModel<'_>, budget: &mut Budget) -> OlapProofResult<()> {
    for file in &model.files {
        budget.item("OLAP files")?;
        budget.string(file.storage_path, "OLAP paths")?;
        if let OlapDocument::Definition(definition) = &file.document {
            budget.string(&definition.object_id, "OLAP IDs")?;
            for value in [&definition.parent.database_id, &definition.parent.cube_id]
                .into_iter()
                .flatten()
            {
                budget.string(value, "OLAP parent IDs")?;
            }
            if let Some(name) = &definition.object_name {
                budget.string(name, "OLAP names")?;
            }
            preflight_olap_element(&definition.object, budget)?;
            for item in all_file_lists_for_budget(&definition.extension) {
                budget.item("OLAP file-list members")?;
                budget.string(item, "OLAP file-list paths")?;
            }
            for item in &definition.attribute_ids {
                budget.item("OLAP attribute IDs")?;
                budget.string(item, "OLAP attribute IDs")?;
            }
            for hierarchy in &definition.hierarchies {
                budget.item("OLAP hierarchies")?;
                budget.string(&hierarchy.id, "OLAP hierarchy IDs")?;
                for level in &hierarchy.level_ids {
                    budget.item("OLAP hierarchy levels")?;
                    budget.string(level, "OLAP hierarchy IDs")?;
                }
            }
        }
    }
    Ok(())
}

fn preflight_olap_element(element: &OlapElement, budget: &mut Budget) -> OlapProofResult<()> {
    budget.item("OLAP XML nodes")?;
    budget.string(&element.name, "OLAP XML names")?;
    budget.string(&element.text, "OLAP XML text")?;
    for (name, value) in &element.attributes {
        budget.item("OLAP XML attributes")?;
        budget.string(name, "OLAP XML attributes")?;
        budget.string(value, "OLAP XML attributes")?;
    }
    for child in &element.children {
        preflight_olap_element(child, budget)?;
    }
    Ok(())
}

fn preflight_retained_strings(
    storage: &Storage<'_>,
    metadata: &MetadataModel<'_>,
    model: &OlapModel<'_>,
    budget: &mut Budget,
) -> OlapProofResult<()> {
    for file in &storage.files {
        charge_retained_string(budget, &file.path, "retained physical paths")?;
    }
    for group in &storage.backup_log.file_groups {
        for value in [
            group.id.as_str(),
            group.name.as_str(),
            group.object_id.as_str(),
        ] {
            charge_retained_string(budget, value, "retained FileGroup text")?;
        }
        for file in &group.files {
            charge_retained_string(budget, &file.storage_path, "retained member paths")?;
        }
    }
    for file in &metadata.files {
        charge_retained_string(budget, file.storage_path, "retained metadata paths")?;
        charge_retained_metadata_object(&file.table, budget)?;
    }
    for column in &metadata.columns {
        charge_retained_string(budget, &column.name, "retained metadata policy text")?;
        charge_retained_string(budget, &column.data_file, "retained metadata policy paths")?;
        if let Some(dictionary) = &column.dictionary {
            charge_retained_string(
                budget,
                &dictionary.storage_name,
                "retained metadata dictionary paths",
            )?;
            charge_retained_string(
                budget,
                dictionary.class.as_str(),
                "retained metadata dictionary text",
            )?;
        }
    }
    for relationship in &metadata.relationships {
        if let Some(name) = &relationship.name {
            charge_retained_string(budget, name, "retained relationship names")?;
        }
        for value in [
            relationship.primary_table.as_str(),
            relationship.primary_column.as_str(),
            relationship.foreign_column.as_str(),
        ] {
            charge_retained_string(budget, value, "retained relationship endpoints")?;
        }
    }
    for file in &model.files {
        charge_retained_string(budget, file.storage_path, "retained OLAP paths")?;
        if let OlapDocument::Definition(definition) = &file.document {
            charge_retained_string(budget, &definition.object_id, "retained OLAP IDs")?;
            if let Some(name) = &definition.object_name {
                charge_retained_string(budget, name, "retained OLAP names")?;
            }
            for value in [&definition.parent.database_id, &definition.parent.cube_id]
                .into_iter()
                .flatten()
            {
                charge_retained_string(budget, value, "retained OLAP parent IDs")?;
            }
            charge_retained_olap_element(&definition.object, budget)?;
            for value in all_file_lists_for_budget(&definition.extension) {
                charge_retained_string(budget, value, "retained OLAP file-list paths")?;
            }
            charge_retained_derived_definition_strings(file.storage_path, definition, budget)?;
            for value in &definition.attribute_ids {
                charge_retained_string(budget, value, "retained OLAP attribute IDs")?;
            }
            for hierarchy in &definition.hierarchies {
                charge_retained_string(budget, &hierarchy.id, "retained OLAP hierarchy IDs")?;
                for value in &hierarchy.level_ids {
                    charge_retained_string(budget, value, "retained OLAP hierarchy levels")?;
                }
            }
        }
    }
    let logged_records = storage
        .backup_log
        .file_groups
        .iter()
        .try_fold(0usize, |total, group| total.checked_add(group.files.len()))
        .ok_or_else(|| {
            limit(
                "retained strings",
                usize::MAX,
                budget.limits.max_string_bytes,
            )
        })?;
    let records = storage
        .files
        .len()
        .checked_add(storage.backup_log.file_groups.len())
        .and_then(|value| value.checked_add(metadata.files.len()))
        .and_then(|value| value.checked_add(model.files.len()))
        .and_then(|value| value.checked_add(logged_records))
        .ok_or_else(|| {
            limit(
                "retained strings",
                usize::MAX,
                budget.limits.max_string_bytes,
            )
        })?;
    let slack = records
        .checked_mul(DERIVED_STRING_SLACK_PER_RECORD)
        .ok_or_else(|| {
            limit(
                "retained strings",
                usize::MAX,
                budget.limits.max_string_bytes,
            )
        })?;
    budget.retained_bytes(slack, "retained generated strings")?;
    Ok(())
}

fn charge_retained_derived_definition_strings(
    path: &str,
    definition: &OlapDefinition,
    budget: &mut Budget,
) -> OlapProofResult<()> {
    let Some((folder_len, full_folder_len)) = persist_folder_lengths(path, definition) else {
        // The semantic pass reports malformed definition paths. Do not turn
        // that structural error into a preflight error merely while sizing a
        // string that the semantic pass will never allocate.
        return Ok(());
    };
    let special = matches!(
        definition.kind,
        OlapObjectKind::DataSourceView | OlapObjectKind::MdxScript
    );
    budget.retained_bytes(folder_len, "retained persistence folders")?;
    if !special && full_folder_len != folder_len {
        budget.retained_bytes(full_folder_len, "retained persistence folders")?;
    }
    for item in all_file_lists_for_budget(&definition.extension) {
        if let Some(length) = qualified_file_list_len(folder_len, item) {
            budget.retained_bytes(length, "retained qualified file-list paths")?;
        }
    }
    Ok(())
}

fn persist_folder_lengths(path: &str, definition: &OlapDefinition) -> Option<(usize, usize)> {
    let parent = path.rsplit_once('/').map_or("", |value| value.0);
    if matches!(
        definition.kind,
        OlapObjectKind::DataSourceView | OlapObjectKind::MdxScript
    ) {
        return Some((parent.len(), parent.len()));
    }
    let base = path
        .rsplit('/')
        .next()
        .unwrap_or(path)
        .strip_suffix(definition.kind.suffix_for_proof())?;
    let id = base.rsplit_once('.')?.0;
    let object_suffix = definition.kind.suffix_for_proof().trim_end_matches(".xml");
    let folder_len = id
        .len()
        .checked_add(1)?
        .checked_add(decimal_len(definition.extension.persist_location))?
        .checked_add(object_suffix.len())?;
    let full_len = if parent.is_empty() {
        folder_len
    } else {
        parent.len().checked_add(1)?.checked_add(folder_len)?
    };
    Some((folder_len, full_len))
}

fn qualified_file_list_len(folder_len: usize, item: &str) -> Option<usize> {
    if item.is_empty()
        || item.starts_with('/')
        || item.contains('\\')
        || item
            .split('/')
            .any(|part| part.is_empty() || part == "." || part == "..")
    {
        return None;
    }
    if item.contains('/') || folder_len == 0 {
        Some(item.len())
    } else {
        folder_len.checked_add(1)?.checked_add(item.len())
    }
}

fn charge_retained_metadata_object(
    object: &MetadataObject,
    budget: &mut Budget,
) -> OlapProofResult<()> {
    charge_retained_string(
        budget,
        object.class.as_str(),
        "retained metadata object text",
    )?;
    if let Some(name) = &object.name {
        charge_retained_string(budget, name, "retained metadata object names")?;
    }
    for property in &object.properties {
        charge_retained_string(budget, &property.name, "retained metadata property text")?;
        charge_retained_string(budget, &property.value, "retained metadata property text")?;
    }
    for member in &object.members {
        charge_retained_string(budget, &member.name, "retained metadata member names")?;
        charge_retained_metadata_object(&member.object, budget)?;
    }
    for collection in &object.collections {
        charge_retained_string(
            budget,
            &collection.name,
            "retained metadata collection names",
        )?;
        for child in &collection.objects {
            charge_retained_metadata_object(child, budget)?;
        }
    }
    for data in &object.data_objects {
        charge_retained_metadata_object(&data.object, budget)?;
    }
    Ok(())
}

fn charge_retained_olap_element(element: &OlapElement, budget: &mut Budget) -> OlapProofResult<()> {
    charge_retained_string(budget, &element.name, "retained OLAP XML names")?;
    charge_retained_string(budget, &element.text, "retained OLAP XML text")?;
    for (name, value) in &element.attributes {
        charge_retained_string(budget, name, "retained OLAP XML attributes")?;
        charge_retained_string(budget, value, "retained OLAP XML attributes")?;
    }
    for child in &element.children {
        charge_retained_olap_element(child, budget)?;
    }
    Ok(())
}

fn charge_retained_string(
    budget: &mut Budget,
    value: &str,
    resource: &'static str,
) -> OlapProofResult<()> {
    for _ in 0..RETAINED_STRING_COPIES {
        budget.retained_bytes(value.len(), resource)?;
    }
    Ok(())
}

fn all_file_lists_for_budget(
    extension: &super::olap::TabularExtension,
) -> impl Iterator<Item = &str> {
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
    .flat_map(|values| values.iter().map(String::as_str))
}

struct Budget {
    limits: OlapProofLimits,
    items: usize,
    strings: usize,
    work: usize,
}

impl Budget {
    fn new(limits: OlapProofLimits) -> Self {
        Self {
            limits,
            items: 0,
            strings: 0,
            work: 0,
        }
    }

    fn item(&mut self, resource: &'static str) -> OlapProofResult<()> {
        self.items = self
            .items
            .checked_add(1)
            .ok_or_else(|| limit(resource, usize::MAX, self.limits.max_items))?;
        if self.items > self.limits.max_items {
            return Err(limit(resource, self.items, self.limits.max_items));
        }
        self.work("item")
    }

    fn string(&mut self, value: &str, resource: &'static str) -> OlapProofResult<()> {
        self.retained_bytes(value.len(), resource)?;
        self.work("string")
    }

    fn retained_bytes(&mut self, amount: usize, resource: &'static str) -> OlapProofResult<()> {
        self.strings = self
            .strings
            .checked_add(amount)
            .ok_or_else(|| limit(resource, usize::MAX, self.limits.max_string_bytes))?;
        if self.strings > self.limits.max_string_bytes {
            return Err(limit(resource, self.strings, self.limits.max_string_bytes));
        }
        Ok(())
    }

    fn source_bytes(&self, actual: usize) -> OlapProofResult<()> {
        if actual > self.limits.max_source_bytes {
            Err(limit("source bytes", actual, self.limits.max_source_bytes))
        } else {
            Ok(())
        }
    }

    fn work_bound(&mut self, graph_nodes: usize) -> OlapProofResult<()> {
        let repeated = graph_nodes
            .checked_mul(graph_nodes)
            .ok_or_else(|| limit("graph work", usize::MAX, self.limits.max_work))?;
        self.work = self
            .work
            .checked_add(repeated)
            .ok_or_else(|| limit("graph work", usize::MAX, self.limits.max_work))?;
        if self.work > self.limits.max_work {
            return Err(limit("graph work", self.work, self.limits.max_work));
        }
        Ok(())
    }

    fn work(&mut self, resource: &'static str) -> OlapProofResult<()> {
        self.work = self
            .work
            .checked_add(1)
            .ok_or_else(|| limit(resource, usize::MAX, self.limits.max_work))?;
        if self.work > self.limits.max_work {
            return Err(limit(resource, self.work, self.limits.max_work));
        }
        Ok(())
    }
}

fn reserve_vec<T>(
    vector: &mut Vec<T>,
    amount: usize,
    resource: &'static str,
) -> OlapProofResult<()> {
    vector
        .try_reserve(amount)
        .map_err(|error| OlapProofError::Allocation {
            resource,
            detail: error.to_string(),
        })
}

fn reserve_map<K: Eq + std::hash::Hash, V>(
    map: &mut HashMap<K, V>,
    amount: usize,
    resource: &'static str,
) -> OlapProofResult<()> {
    map.try_reserve(amount)
        .map_err(|error| OlapProofError::Allocation {
            resource,
            detail: error.to_string(),
        })
}

fn reserve_set<T: Eq + std::hash::Hash>(
    set: &mut HashSet<T>,
    amount: usize,
    resource: &'static str,
) -> OlapProofResult<()> {
    set.try_reserve(amount)
        .map_err(|error| OlapProofError::Allocation {
            resource,
            detail: error.to_string(),
        })
}

fn push_unknown(
    values: &mut Vec<OlapUnknownMember>,
    value: OlapUnknownMember,
) -> OlapProofResult<()> {
    values
        .try_reserve(1)
        .map_err(|error| OlapProofError::Allocation {
            resource: "unknown OLAP members",
            detail: error.to_string(),
        })?;
    values.push(value);
    Ok(())
}

fn invalid(path: impl Into<String>, detail: impl Into<String>) -> OlapProofError {
    OlapProofError::Invalid {
        path: path.into(),
        detail: detail.into(),
    }
}

fn unproven(path: impl Into<String>, detail: impl Into<String>) -> OlapProofError {
    OlapProofError::Unproven {
        path: path.into(),
        detail: detail.into(),
    }
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> OlapProofError {
    OlapProofError::LimitExceeded {
        resource,
        actual,
        maximum,
    }
}

trait ProofSuffix {
    fn suffix_for_proof(self) -> &'static str;
    fn element_name_for_proof(self) -> &'static str;
}

impl ProofSuffix for OlapObjectKind {
    fn suffix_for_proof(self) -> &'static str {
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

    fn element_name_for_proof(self) -> &'static str {
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
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::metadata::{MetadataClass, MetadataCollection, MetadataFile, MetadataObject};
    use crate::model::{
        FileEntry, FileGroup, GeneratedPath, LoggedFile, Offset, Size, test_xldm140_storage,
    };
    use crate::olap::{
        CubeInformation, OlapDocument, OlapFile, OlapFileKind, OlapParentReference,
        TabularExtension,
    };

    fn scalar(name: &str, value: &str) -> OlapElement {
        OlapElement {
            name: name.into(),
            attributes: Vec::new(),
            text: value.into(),
            children: Vec::new(),
        }
    }

    fn element(name: &str, children: Vec<OlapElement>) -> OlapElement {
        OlapElement {
            name: name.into(),
            attributes: Vec::new(),
            text: String::new(),
            children,
        }
    }

    fn typed(name: &str, type_name: &str, children: Vec<OlapElement>) -> OlapElement {
        OlapElement {
            name: name.into(),
            attributes: vec![("xsi:type".into(), type_name.into())],
            text: String::new(),
            children,
        }
    }

    fn object_name(object_id: &str, name: &str, children: Vec<OlapElement>) -> OlapElement {
        let mut result = vec![scalar("Name", name), scalar("ObjectID", object_id)];
        result.extend(children);
        element("Object", result)
    }

    fn extension(
        version: i32,
        persist: i32,
        data_files: &[&str],
        measure_group_files: &[&str],
        partition_files: &[&str],
    ) -> TabularExtension {
        TabularExtension {
            ordinal: 0,
            object_version: version,
            persist_location: persist,
            data_files: data_files.iter().map(|value| (*value).into()).collect(),
            permission_files: Vec::new(),
            measure_group_files: measure_group_files
                .iter()
                .map(|value| (*value).into())
                .collect(),
            perspective_files: Vec::new(),
            assembly_files: Vec::new(),
            aggregation_design_files: Vec::new(),
            partition_files: partition_files
                .iter()
                .map(|value| (*value).into())
                .collect(),
            default_collation_version: None,
        }
    }

    fn definition(
        path: &'static str,
        kind: OlapObjectKind,
        id: &str,
        object_id: &str,
        name: &str,
        children: Vec<OlapElement>,
        extension: TabularExtension,
    ) -> OlapFile<'static> {
        let mut object = object_name(object_id, name, children);
        object.name = kind.element_name_for_proof().into();
        object.children.insert(0, scalar("ID", id));
        OlapFile {
            storage_path: path,
            bytes: &[],
            kind: OlapFileKind::Definition(kind),
            document: OlapDocument::Definition(OlapDefinition {
                parent: OlapParentReference::default(),
                kind,
                object_id: id.into(),
                object_name: Some(name.into()),
                object,
                extension,
                attribute_ids: Vec::new(),
                hierarchies: Vec::new(),
            }),
        }
    }

    fn logged(path: &'static str, kind: GeneratedNameKind) -> LoggedFile {
        LoggedFile {
            source_path: path.into(),
            storage_path: path.into(),
            last_write_timestamp: 0,
            size: 0,
            generated: GeneratedPath {
                normalized_path: path.into(),
                kind,
            },
        }
    }

    fn group(
        class: FileGroupClass,
        id: &str,
        name: &str,
        object_id: &str,
        version: i32,
        persist: i32,
        members: Vec<(&'static str, GeneratedNameKind)>,
    ) -> FileGroup {
        FileGroup {
            class,
            id: id.into(),
            name: name.into(),
            object_version: version,
            persist_location: persist,
            persist_location_path: "Model.1.db".into(),
            storage_location_path: String::new(),
            object_id: object_id.into(),
            files: members
                .into_iter()
                .map(|(path, kind)| logged(path, kind))
                .collect(),
        }
    }

    fn metadata_table() -> MetadataFile<'static> {
        let column = MetadataObject {
            class: MetadataClass("XMRawColumn".into()),
            name: Some("Key".into()),
            provider_version: None,
            properties: Vec::new(),
            members: Vec::new(),
            collections: Vec::new(),
            data_objects: Vec::new(),
        };
        let table = MetadataObject {
            class: MetadataClass("XMSimpleTable".into()),
            name: Some("InnerT1".into()),
            provider_version: None,
            properties: Vec::new(),
            members: Vec::new(),
            collections: vec![MetadataCollection {
                name: "Columns".into(),
                objects: vec![column],
            }],
            data_objects: Vec::new(),
        };
        MetadataFile {
            storage_path: "Model.1.db/T1.0.dim/T1.1.tbl.xml",
            bytes: &[],
            kind: MetadataFileKind::Table,
            table,
        }
    }

    fn complete_fixture() -> (Storage<'static>, MetadataModel<'static>, OlapModel<'static>) {
        let dimension_attributes = element(
            "Attributes",
            vec![typed(
                "Attribute",
                "DimensionAttribute",
                vec![scalar("ID", "Key")],
            )],
        );
        let cube_dimensions = element(
            "Dimensions",
            vec![typed(
                "Dimension",
                "CubeDimension",
                vec![
                    scalar("DimensionID", "DIM1"),
                    element(
                        "Attributes",
                        vec![typed(
                            "Attribute",
                            "CubeAttribute",
                            vec![scalar("AttributeID", "Key")],
                        )],
                    ),
                ],
            )],
        );
        let measure_dimensions = element(
            "Dimensions",
            vec![typed(
                "Dimension",
                "DegenerateMeasureGroupDimension",
                vec![
                    scalar("CubeDimensionID", "DIM1"),
                    element(
                        "Attributes",
                        vec![typed(
                            "Attribute",
                            "MeasureGroupDimensionAttribute",
                            vec![scalar("AttributeID", "Key")],
                        )],
                    ),
                ],
            )],
        );
        let files = vec![
            definition(
                "Model.1.db.xml",
                OlapObjectKind::Database,
                "DB",
                "OBJ-DB",
                "Database",
                vec![scalar("DbStorageLocation", "Model.1.db")],
                extension(1, 1, &[], &[], &[]),
            ),
            definition(
                "Model.1.db/Cube.1.cub.xml",
                OlapObjectKind::Cube,
                "CUBE",
                "OBJ-CUBE",
                "Cube",
                vec![cube_dimensions],
                extension(1, 0, &[], &["T1.1.det.xml"], &[]),
            ),
            definition(
                "Model.1.db/T1.1.dim.xml",
                OlapObjectKind::Dimension,
                "DIM1",
                "OBJ-DIM",
                "Dimension",
                vec![dimension_attributes],
                extension(1, 0, &["1.T1.Key.0.idf"], &[], &[]),
            ),
            definition(
                "Model.1.db/Cube.0.cub/T1.1.det.xml",
                OlapObjectKind::MeasureGroup,
                "MG1",
                "OBJ-MG",
                "MeasureGroup",
                vec![measure_dimensions],
                extension(1, 0, &[], &[], &["T1.1.prt.xml"]),
            ),
            definition(
                "Model.1.db/Cube.0.cub/T1.0.det/T1.1.prt.xml",
                OlapObjectKind::Partition,
                "P1",
                "OBJ-P",
                "Partition",
                Vec::new(),
                extension(1, 0, &[], &[], &[]),
            ),
        ];
        let paths = [
            "Model.1.db.xml",
            "Model.1.db/Cube.1.cub.xml",
            "Model.1.db/T1.1.dim.xml",
            "Model.1.db/T1.0.dim/T1.1.tbl.xml",
            "Model.1.db/T1.0.dim/1.T1.Key.0.idf",
            "Model.1.db/Cube.0.cub/T1.1.det.xml",
            "Model.1.db/Cube.0.cub/T1.0.det/T1.1.prt.xml",
        ];
        let mut storage = test_xldm140_storage(&[0], &paths);
        storage.backup_log.is_olap = true;
        storage.backup_log.file_groups = vec![
            group(
                FileGroupClass::Database,
                "DB",
                "Database",
                "OBJ-DB",
                1,
                1,
                vec![("Model.1.db.xml", GeneratedNameKind::DatabaseDefinition)],
            ),
            group(
                FileGroupClass::Cube,
                "CUBE",
                "Cube",
                "OBJ-CUBE",
                1,
                1,
                vec![(
                    "Model.1.db/Cube.1.cub.xml",
                    GeneratedNameKind::CubeDefinition,
                )],
            ),
            group(
                FileGroupClass::Dimension,
                "DIM1",
                "Dimension",
                "OBJ-DIM",
                1,
                1,
                vec![
                    (
                        "Model.1.db/T1.1.dim.xml",
                        GeneratedNameKind::DataSourceOrDimensionDefinition,
                    ),
                    (
                        "Model.1.db/T1.0.dim/T1.1.tbl.xml",
                        GeneratedNameKind::TableMetadata,
                    ),
                    (
                        "Model.1.db/T1.0.dim/1.T1.Key.0.idf",
                        GeneratedNameKind::ColumnData,
                    ),
                ],
            ),
            group(
                FileGroupClass::MeasureGroup,
                "MG1",
                "MeasureGroup",
                "OBJ-MG",
                1,
                1,
                vec![(
                    "Model.1.db/Cube.0.cub/T1.1.det.xml",
                    GeneratedNameKind::MeasureGroupMetadata,
                )],
            ),
            group(
                FileGroupClass::Partition,
                "P1",
                "Partition",
                "OBJ-P",
                1,
                1,
                vec![(
                    "Model.1.db/Cube.0.cub/T1.0.det/T1.1.prt.xml",
                    GeneratedNameKind::PartitionMetadata,
                )],
            ),
        ];
        let metadata = MetadataModel {
            files: vec![metadata_table()],
            columns: Vec::new(),
            relationships: Vec::new(),
            hierarchies: Vec::new(),
        };
        (storage, metadata, OlapModel { files })
    }

    fn relationship_fixture(
        primary_column: &str,
        foreign_column: &str,
    ) -> (Storage<'static>, MetadataModel<'static>, OlapModel<'static>) {
        let (mut storage, mut metadata, mut model) = complete_fixture();
        let relation_path = "Model.1.db/T1.0.dim/R$T1$Rel.1.tbl.xml";
        let index_path = "Model.1.db/T1.0.dim/1.R$T1$Rel.INDEX.0.idf";
        storage.files.push(FileEntry {
            path: relation_path.into(),
            kind: crate::FileKind::OpaqueBinary,
            offset: Offset(0),
            stored_size: Size(0),
            crc32: 0,
            delete: false,
            created_timestamp: 0,
            access_timestamp: 0,
            last_write_timestamp: 0,
        });
        storage.files.push(FileEntry {
            path: index_path.into(),
            kind: crate::FileKind::OpaqueBinary,
            offset: Offset(0),
            stored_size: Size(0),
            crc32: 0,
            delete: false,
            created_timestamp: 0,
            access_timestamp: 0,
            last_write_timestamp: 0,
        });
        storage.backup_log.file_groups[2].files.push(logged(
            relation_path,
            GeneratedNameKind::TableRelationshipMetadata,
        ));
        storage.backup_log.file_groups[2].files.push(logged(
            index_path,
            GeneratedNameKind::TableRelationshipIndex,
        ));

        metadata.files.push(MetadataFile {
            storage_path: relation_path,
            bytes: &[],
            kind: MetadataFileKind::TableRelationship,
            table: MetadataObject {
                class: MetadataClass("XMSimpleTable".into()),
                name: Some("InnerT1".into()),
                provider_version: None,
                properties: Vec::new(),
                members: Vec::new(),
                collections: vec![MetadataCollection {
                    name: "Relationships".into(),
                    objects: vec![MetadataObject {
                        class: MetadataClass("XMRelationship".into()),
                        name: Some("Rel".into()),
                        provider_version: None,
                        properties: vec![
                            crate::metadata::MetadataProperty {
                                name: "PrimaryTable".into(),
                                value: "InnerT1".into(),
                            },
                            crate::metadata::MetadataProperty {
                                name: "PrimaryColumn".into(),
                                value: primary_column.into(),
                            },
                            crate::metadata::MetadataProperty {
                                name: "ForeignColumn".into(),
                                value: foreign_column.into(),
                            },
                        ],
                        members: Vec::new(),
                        collections: Vec::new(),
                        data_objects: Vec::new(),
                    }],
                }],
                data_objects: Vec::new(),
            },
        });
        let OlapDocument::Definition(dimension) = &mut model.files[2].document else {
            panic!("fixture dimension");
        };
        dimension.extension.data_files.push(index_path.into());
        dimension.object.children.push(element(
            "Relationships",
            vec![element("Relationship", vec![scalar("ID", "Rel")])],
        ));
        let OlapDocument::Definition(measure_group) = &mut model.files[3].document else {
            panic!("fixture measure group");
        };
        let dimensions = measure_group
            .object
            .children
            .iter_mut()
            .find(|node| node.name == "Dimensions")
            .expect("fixture dimensions");
        dimensions.children.push(typed(
            "Dimension",
            "ReferenceMeasureGroupDimension",
            vec![
                scalar("CubeDimensionID", "DIM1"),
                element(
                    "Attributes",
                    vec![typed(
                        "Attribute",
                        "MeasureGroupDimensionAttribute",
                        vec![scalar("AttributeID", "Key")],
                    )],
                ),
            ],
        ));
        (storage, metadata, model)
    }

    #[test]
    fn proves_complete_cube_measure_group_partition_owner_graph() {
        let (storage, metadata, model) = complete_fixture();
        let proof = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .expect("complete fixture should be proven");
        assert!(proof.is_complete());
        assert_eq!(proof.tables()[0].table_id, "T1");
        assert_eq!(proof.tables()[0].metadata_name.as_deref(), Some("InnerT1"));
        assert_eq!(proof.tables()[0].dimension_id, "DIM1");
        assert_eq!(proof.cube().unwrap().dimensions[0].reference.value, "DIM1");
        assert_eq!(proof.measure_groups()[0].table_dimension_id, "DIM1");
        assert_eq!(
            proof.partitions()[0].measure_group_path,
            "Model.1.db/Cube.0.cub/T1.1.det.xml"
        );
        assert_eq!(
            proof.file_groups()[2].data_files,
            ["Model.1.db/T1.0.dim/1.T1.Key.0.idf"]
        );
        assert!(std::ptr::eq(
            proof.storage().bytes().as_ptr(),
            storage.bytes().as_ptr()
        ));
    }

    #[test]
    fn proves_primary_dimension_relationship_and_foreign_side_index_owner() {
        let (storage, metadata, model) = relationship_fixture("Key", "Key");
        let index_path = "Model.1.db/T1.0.dim/1.R$T1$Rel.INDEX.0.idf";

        let proof = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .expect("relationship closure should be proven");
        assert!(proof.is_complete());
        assert_eq!(proof.relationships().len(), 1);
        let relationship = &proof.relationships()[0];
        assert_eq!(relationship.containing_table, "T1");
        assert_eq!(relationship.primary_table, "InnerT1");
        assert_eq!(relationship.relationship_index_paths, [index_path]);
        assert_eq!(relationship.dimension_reference.value, "Rel");
    }

    #[test]
    fn refuses_relationship_column_collision_across_table_identity() {
        let (storage, metadata, model) = relationship_fixture("WrongTableColumn", "Key");
        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(error.to_string().contains("PrimaryColumn"));
    }

    #[test]
    fn primary_table_accepts_metadata_name_but_not_generated_table_id() {
        let (storage, mut metadata, model) = relationship_fixture("Key", "Key");
        metadata.files[1].table.collections[0].objects[0].properties[0].value = "T1".into();
        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(error.to_string().contains("no matching table"));
    }

    #[test]
    fn dimension_relationship_uses_filename_relid_not_metadata_name() {
        let (storage, mut metadata, model) = relationship_fixture("Key", "Key");
        let relation = &mut metadata.files[1].table.collections[0].objects[0];
        relation.name = Some("DisplayOnlyRelationshipName".into());
        let proof = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .expect("the Dimension edge is keyed by the generated RelId");
        assert_eq!(
            proof.relationships()[0].relationship_name.as_deref(),
            Some("DisplayOnlyRelationshipName")
        );

        let (storage, mut metadata, mut model) = relationship_fixture("Key", "Key");
        let relation = &mut metadata.files[1].table.collections[0].objects[0];
        relation.name = Some("DisplayOnlyRelationshipName".into());
        let OlapDocument::Definition(dimension) = &mut model.files[2].document else {
            panic!("fixture dimension");
        };
        let relationships = dimension
            .object
            .children
            .iter_mut()
            .find(|node| node.name == "Relationships")
            .expect("fixture relationships");
        relationships.children[0].children[0].text = "DisplayOnlyRelationshipName".into();
        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(error.to_string().contains("missing or ambiguous"));
    }

    #[test]
    fn refuses_duplicate_dimension_relationship_binding() {
        let (mut storage, mut metadata, model) = relationship_fixture("Key", "Key");
        let duplicate_path = "Model.1.db/T1.0.dim/R$T1$Rel.2.tbl.xml";
        storage.files.push(FileEntry {
            path: duplicate_path.into(),
            kind: crate::FileKind::OpaqueBinary,
            offset: Offset(0),
            stored_size: Size(0),
            crc32: 0,
            delete: false,
            created_timestamp: 0,
            access_timestamp: 0,
            last_write_timestamp: 0,
        });
        storage.backup_log.file_groups[2].files.push(logged(
            duplicate_path,
            GeneratedNameKind::TableRelationshipMetadata,
        ));
        let mut duplicate = metadata.files[1].clone();
        duplicate.storage_path = duplicate_path;
        metadata.files.push(duplicate);

        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("duplicate Dimension relationship binding")
        );
    }

    #[test]
    fn relationship_relid_fallback_keeps_filename_identity() {
        let (storage, mut metadata, model) = relationship_fixture("Key", "Key");
        let relation = &mut metadata.files[1].table.collections[0].objects[0];
        relation.name = None;
        let proof = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .expect("generated RelId supplies the relationship identity");
        assert_eq!(proof.relationships()[0].relationship_name, None);
        assert_eq!(proof.relationships()[0].generated_relationship_id, "Rel");
    }

    #[test]
    fn marks_unknown_or_duplicate_dimension_relationship_collections_incomplete() {
        let (storage, metadata, mut model) = complete_fixture();
        let OlapDocument::Definition(dimension) = &mut model.files[2].document else {
            panic!("fixture dimension");
        };
        dimension.object.children.push(element(
            "Relationships",
            vec![element("Future", Vec::new())],
        ));
        let proof = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .expect("unknown relationship members remain visible");
        assert!(!proof.is_complete());
        assert!(
            proof
                .unknown_members()
                .iter()
                .any(|member| member.reason.contains("unknown member"))
        );

        let (storage, metadata, mut model) = complete_fixture();
        let OlapDocument::Definition(dimension) = &mut model.files[2].document else {
            panic!("fixture dimension");
        };
        dimension.object.children.extend([
            element("Relationships", Vec::new()),
            element("Relationships", Vec::new()),
        ]);
        let proof = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .expect("duplicate relationship collections remain visible");
        assert!(!proof.is_complete());
        assert!(
            proof
                .unknown_members()
                .iter()
                .any(|member| member.reason.contains("duplicate Relationships"))
        );
    }

    #[test]
    fn refuses_materialized_member_with_a_different_table_owner() {
        let (mut storage, metadata, mut model) = complete_fixture();
        let original = "Model.1.db/T1.0.dim/1.T1.Key.0.idf";
        let spoofed = "Model.1.db/T1.0.dim/1.T2.Key.0.idf";
        storage
            .files
            .iter_mut()
            .find(|file| file.path == original)
            .expect("physical column data")
            .path = spoofed.into();
        let logged = storage.backup_log.file_groups[2]
            .files
            .iter_mut()
            .find(|file| file.storage_path == original)
            .expect("logged column data");
        logged.source_path = spoofed.into();
        logged.storage_path = spoofed.into();
        logged.generated.normalized_path = spoofed.into();
        let OlapDocument::Definition(dimension) = &mut model.files[2].document else {
            panic!("fixture dimension");
        };
        dimension.extension.data_files[0] = spoofed.into();

        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("does not encode its owning table T1")
        );

        let (mut storage, metadata, model) = complete_fixture();
        let logged = storage.backup_log.file_groups[2]
            .files
            .iter_mut()
            .find(|file| file.storage_path == original)
            .expect("logged column data");
        logged.generated.normalized_path = spoofed.into();
        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(
            error
                .to_string()
                .contains("generated path disagrees with the logged storage path")
        );
    }

    #[test]
    fn refuses_file_group_object_id_spoof_and_data_file_list_undercoverage() {
        let (mut storage, metadata, model) = complete_fixture();
        storage.backup_log.file_groups[2].object_id = "WRONG".into();
        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(error.to_string().contains("ObjectID"));

        let (storage, metadata, mut model) = complete_fixture();
        let OlapDocument::Definition(dimension) = &mut model.files[2].document else {
            panic!("fixture dimension");
        };
        dimension.extension.data_files.clear();
        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(error.to_string().contains("DataFileList"));
    }

    #[test]
    fn refuses_file_group_persist_location_and_filename_version_drift() {
        let (mut storage, metadata, model) = complete_fixture();
        storage.backup_log.file_groups[0].persist_location = 2;
        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(error.to_string().contains("PersistLocation"));

        let (storage, metadata, mut model) = complete_fixture();
        let OlapDocument::Definition(dimension) = &mut model.files[2].document else {
            panic!("fixture dimension");
        };
        dimension.extension.object_version = 2;
        let error = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .unwrap_err();
        assert!(error.to_string().contains("ObjectVersion"));
    }

    #[test]
    fn leaves_recognized_unlinked_cube_dimension_visible() {
        let (storage, metadata, mut model) = complete_fixture();
        let OlapDocument::Definition(cube) = &mut model.files[1].document else {
            panic!("fixture cube");
        };
        let Some(dimensions) = cube
            .object
            .children
            .iter_mut()
            .find(|node| node.name == "Dimensions")
        else {
            panic!("fixture cube dimensions");
        };
        dimensions.children.push(typed(
            "Dimension",
            "CubeDimension",
            vec![
                scalar("DimensionID", "UNLINKED"),
                element("Attributes", Vec::new()),
            ],
        ));
        let proof = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .expect("unlinked extra members remain readable");
        assert!(!proof.is_complete());
        assert!(
            proof
                .unknown_members()
                .iter()
                .any(|member| member.reason.contains("UNLINKED"))
        );
    }

    #[test]
    fn leaves_unlinked_physical_olap_member_visible() {
        let (mut storage, metadata, model) = complete_fixture();
        let path = "Model.1.db/T1.0.dim/info.1.xml";
        storage.files.push(FileEntry {
            path: path.into(),
            kind: crate::FileKind::OpaqueBinary,
            offset: Offset(0),
            stored_size: Size(0),
            crc32: 0,
            delete: false,
            created_timestamp: 0,
            access_timestamp: 0,
            last_write_timestamp: 0,
        });
        storage.backup_log.file_groups[2]
            .files
            .push(logged(path, GeneratedNameKind::TableInformation));
        let proof = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .expect("unlinked physical members remain readable");
        assert!(!proof.is_complete());
        assert!(proof.unknown_members().iter().any(|member| {
            member.path == path && member.reason.contains("no proven OLAP or metadata owner")
        }));
    }

    #[test]
    fn leaves_unlinked_metadata_and_olap_members_visible() {
        let (storage, mut metadata, mut model) = complete_fixture();
        let mut hierarchy = metadata_table();
        hierarchy.storage_path = "Model.1.db/T1.0.dim/H$T1$Hierarchy.1.tbl.xml";
        hierarchy.kind = MetadataFileKind::UserHierarchy;
        metadata.files.push(hierarchy);
        model.files.push(OlapFile {
            storage_path: "Model.1.db/info.1.cub.xml",
            bytes: &[],
            kind: OlapFileKind::CubeInformation,
            document: OlapDocument::CubeInformation(CubeInformation),
        });

        let proof = prove_xldm140_olap(&storage, &metadata, &model, OlapProofLimits::default())
            .expect("unlinked metadata and OLAP members remain readable");
        assert!(!proof.is_complete());
        assert!(proof.unknown_members().iter().any(|member| {
            member.path == "Model.1.db/T1.0.dim/H$T1$Hierarchy.1.tbl.xml"
                && member.reason.contains("metadata member")
        }));
        assert!(proof.unknown_members().iter().any(|member| {
            member.path == "Model.1.db/info.1.cub.xml" && member.reason.contains("OLAP member")
        }));
    }

    #[test]
    fn preflight_rejects_string_budget_before_proof_allocation() {
        let (storage, metadata, model) = complete_fixture();
        let limits = OlapProofLimits {
            max_string_bytes: 8,
            ..OlapProofLimits::default()
        };
        let error = prove_xldm140_olap(&storage, &metadata, &model, limits).unwrap_err();
        assert!(matches!(
            error,
            OlapProofError::LimitExceeded {
                resource: "physical paths" | "OLAP paths" | "metadata paths" | "OLAP XML names",
                ..
            }
        ));
    }

    #[test]
    fn retained_string_budget_has_an_admitted_boundary_before_indexes() {
        let (storage, metadata, model) = complete_fixture();
        let mut budget = Budget::new(OlapProofLimits::default());
        budget.source_bytes(storage.bytes().len()).unwrap();
        preflight_storage(&storage, &mut budget).unwrap();
        preflight_metadata(&metadata, &mut budget).unwrap();
        preflight_olap(&model, &mut budget).unwrap();
        let source_total = budget.strings;
        preflight_retained_strings(&storage, &metadata, &model, &mut budget).unwrap();
        let admitted_total = budget.strings;
        assert!(admitted_total > source_total);

        let error = prove_xldm140_olap(
            &storage,
            &metadata,
            &model,
            OlapProofLimits {
                max_string_bytes: admitted_total - 1,
                ..OlapProofLimits::default()
            },
        )
        .unwrap_err();
        assert!(matches!(
            error,
            OlapProofError::LimitExceeded { resource, .. } if resource.starts_with("retained")
        ));

        prove_xldm140_olap(
            &storage,
            &metadata,
            &model,
            OlapProofLimits {
                max_string_bytes: admitted_total,
                ..OlapProofLimits::default()
            },
        )
        .expect("the exact preflight boundary should admit the proof");
    }

    #[test]
    fn preflight_rejects_source_and_item_limits_before_owned_indexes() {
        let (storage, metadata, model) = complete_fixture();
        let limits = OlapProofLimits {
            max_items: 0,
            ..OlapProofLimits::default()
        };
        let error = prove_xldm140_olap(&storage, &metadata, &model, limits).unwrap_err();
        assert!(matches!(
            error,
            OlapProofError::LimitExceeded {
                resource: "physical files",
                actual: 1,
                maximum: 0
            }
        ));

        let limits = OlapProofLimits {
            max_source_bytes: storage.bytes().len() - 1,
            ..OlapProofLimits::default()
        };
        let error = prove_xldm140_olap(&storage, &metadata, &model, limits).unwrap_err();
        assert!(matches!(
            error,
            OlapProofError::LimitExceeded {
                resource: "source bytes",
                ..
            }
        ));
    }

    #[test]
    fn graph_work_quota_rejects_repeated_owner_scans() {
        let (storage, metadata, model) = complete_fixture();
        let limits = OlapProofLimits {
            max_work: 300,
            ..OlapProofLimits::default()
        };
        let error = prove_xldm140_olap(&storage, &metadata, &model, limits).unwrap_err();
        assert!(matches!(
            error,
            OlapProofError::LimitExceeded {
                resource: "graph work",
                ..
            }
        ));
    }

    #[test]
    fn relationship_index_helper_requires_idf_member_name_for_sparse_payloads() {
        assert!(is_relationship_index_path(
            "Model.1.db/T1.0.dim/1.R$T1$Rel.INDEX.0.idf"
        ));
        assert!(!is_relationship_index_path(
            "Model.1.db/T1.0.dim/1.R$T1$Rel.INDEX.0.hidx"
        ));
        assert!(!is_relationship_index_path(
            "Model.1.db/T1.0.dim/1.H$T1$Key.0.hidx"
        ));
        assert!(!is_relationship_index_path(
            "Model.1.db/T1.0.dim/1.R$T1$Rel.INDEX.0.bin"
        ));
    }
}
