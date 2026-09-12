//! Source-bound OPC ownership for Custom Data and Custom Data Properties.
//!
//! Custom Data payloads are deliberately inert.  This owner validates only
//! the package graph, the typed `datastoreItem` properties, and bounded byte
//! sizes; it never interprets the binary storage.

use std::collections::HashSet;
use std::sync::Arc;

use litchi_opc::{
    BlobPart, OpcPackage, OwnedContentTypes, OwnedRelationships, PackURI, Part as OpcPart,
    TargetMode,
};

use crate::connections::embedded_data::Bindings;
use crate::error::{Error, Result, invalid};

use super::codec::{
    canonical_extension, parse_properties, rewrite_extension_list, rewrite_id,
    validate_source_properties, write_properties,
};
use super::{
    CustomData, CustomDataView, DATA_CONTENT_TYPE, DATA_RELATIONSHIP_TYPE, PROPERTIES_CONTENT_TYPE,
    PROPERTIES_RELATIONSHIP_TYPE, RemovalDisposition,
};

const MAX_STORAGES: usize = 4_096;
const MAX_PAYLOAD_BYTES: usize = 512 * 1024 * 1024;

/// Bounded Custom Data resource policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_storages: usize,
    max_payload_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_storages: MAX_STORAGES,
            max_payload_bytes: MAX_PAYLOAD_BYTES,
        }
    }
}

impl Limits {
    /// Construct the default Custom Data policy.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_storages: MAX_STORAGES,
            max_payload_bytes: MAX_PAYLOAD_BYTES,
        }
    }

    /// Set the maximum number of storage pairs.
    #[must_use]
    pub const fn with_max_storages(mut self, value: usize) -> Self {
        self.max_storages = value;
        self
    }

    /// Set the maximum bytes in one inert payload.
    #[must_use]
    pub const fn with_max_payload_bytes(mut self, value: usize) -> Self {
        self.max_payload_bytes = value;
        self
    }

    /// Maximum storage count after the protocol cap.
    #[must_use]
    pub const fn max_storages(self) -> usize {
        if self.max_storages < MAX_STORAGES {
            self.max_storages
        } else {
            MAX_STORAGES
        }
    }

    /// Maximum payload bytes after the protocol cap.
    #[must_use]
    pub const fn max_payload_bytes(self) -> usize {
        if self.max_payload_bytes < MAX_PAYLOAD_BYTES {
            self.max_payload_bytes
        } else {
            MAX_PAYLOAD_BYTES
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipState {
    source: PackURI,
    id: String,
    reltype: String,
    target: String,
    mode: TargetMode,
}

/// One typed Custom Data storage with its physical source binding retained.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    value: CustomDataView,
    properties_part_name: PackURI,
    data_part_name: PackURI,
    workbook_relationship_id: String,
    workbook_relationship_target: String,
    data_relationship_id: String,
    data_relationship_target: String,
    source_properties: Arc<Vec<u8>>,
    source_properties_proof: Option<Arc<litchi_opc::OwnedXmlPart>>,
    source_data: Arc<Vec<u8>>,
    /// Exact source relationship member owned by the Properties part.
    /// `None` is reserved for staged, newly allocated entries.
    properties_relationships: Option<OwnedRelationships>,
    /// Exact source relationship member owned by the inert payload part.
    /// This retains an explicitly present empty `.rels` member.
    data_relationships: Option<OwnedRelationships>,
}

impl Part {
    /// Typed metadata and inert bytes for this storage.
    #[must_use]
    pub fn value(&self) -> &CustomDataView {
        &self.value
    }

    /// Custom Data Properties part URI.
    #[must_use]
    pub fn properties_part_name(&self) -> &PackURI {
        &self.properties_part_name
    }

    /// Custom Data payload part URI.
    #[must_use]
    pub fn data_part_name(&self) -> &PackURI {
        &self.data_part_name
    }

    /// Storage UID.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.value.properties.id
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SourceState {
    content_types: OwnedContentTypes,
    workbook_blob: Arc<Vec<u8>>,
    workbook_relationships: OwnedRelationships,
    incoming_relationships: Arc<[RelationshipState]>,
    connections: Bindings,
}

/// Immutable source-bound Custom Data catalog.
#[derive(Debug, Clone)]
pub struct Snapshot {
    entries: Arc<[Part]>,
    source: SourceState,
    limits: Limits,
}

impl Snapshot {
    /// Load all Custom Data storage pairs.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        Self::load_with_limits(package, &Limits::default())
    }

    /// Load all Custom Data storage pairs with explicit limits.
    pub fn load_with_limits(package: &OpcPackage, limits: &Limits) -> Result<Self> {
        let entries = load_entries(package, limits)?;
        let workbook = package.main_document_part()?;
        let incoming_relationships = incoming_relationships(package, &entries)?;
        let connections = Bindings::load(package)?;
        connections.validate_ids(&entries.iter().map(Part::id).collect())?;
        Ok(Self {
            entries: Arc::from(entries.into_boxed_slice()),
            source: SourceState {
                content_types: package.source_content_types()?,
                workbook_blob: workbook.blob_arc(),
                workbook_relationships: package.source_relationships(workbook.partname())?,
                incoming_relationships: Arc::from(incoming_relationships.into_boxed_slice()),
                connections,
            },
            limits: *limits,
        })
    }

    /// Storages in deterministic UID order.
    #[must_use]
    pub fn entries(&self) -> &[Part] {
        &self.entries
    }

    /// Alias emphasizing the custom-data storage collection.
    #[must_use]
    pub fn storages(&self) -> &[Part] {
        self.entries()
    }

    /// Find a storage by its exact decoded UID.
    #[must_use]
    pub fn find(&self, id: &str) -> Option<&Part> {
        self.entries.iter().find(|entry| entry.id() == id)
    }

    /// Resource policy retained by this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Whether no Custom Data storage exists.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Number of Custom Data storages.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Number of recognized connection references to a storage UID.
    #[must_use]
    pub fn connection_references(&self, id: &str) -> usize {
        self.source.connections.count(id)
    }

    fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
            && self.entries.len() == other.entries.len()
            && self
                .entries
                .iter()
                .zip(other.entries.iter())
                .all(|(left, right)| {
                    same_identity(left, right)
                        && left.source_properties == right.source_properties
                        && left.source_data == right.source_data
                        && left.properties_relationships == right.properties_relationships
                        && left.data_relationships == right.data_relationships
                })
    }

    fn same_semantics(&self, entries: &[Part]) -> bool {
        self.entries.len() == entries.len()
            && self.entries.iter().all(|left| {
                entries
                    .iter()
                    .find(|right| same_identity(left, right))
                    .is_some_and(|right| left.value == right.value)
            })
    }

    fn same_published_semantics(&self, entries: &[Part]) -> bool {
        self.entries.len() == entries.len()
            && self.entries.iter().all(|left| {
                entries
                    .iter()
                    .find(|right| same_identity(left, right))
                    .is_some_and(|right| same_value(&left.value, &right.value))
            })
    }
}

fn same_value(left: &CustomDataView, right: &CustomDataView) -> bool {
    left.data == right.data
        && left.properties.id == right.properties.id
        && match (
            canonical_extension(left.properties.extension_list.as_ref()),
            canonical_extension(right.properties.extension_list.as_ref()),
        ) {
            (Ok(left), Ok(right)) => left == right,
            _ => false,
        }
}

/// Clone-staged Custom Data CRUD transaction.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Vec<Part>,
    limits: Limits,
    connections: Bindings,
}

impl<'a> Transaction<'a> {
    /// Start a transaction with default limits.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        Self::with_limits(target, &Limits::default())
    }

    /// Start a transaction with explicit limits.
    pub fn with_limits(target: &'a mut OpcPackage, limits: &Limits) -> Result<Self> {
        let before = Snapshot::load_with_limits(target, limits)?;
        Ok(Self {
            draft: before.entries.to_vec(),
            connections: before.source.connections.clone(),
            target,
            before,
            limits: *limits,
        })
    }

    /// Source snapshot captured at transaction start.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged entries.
    #[must_use]
    pub fn entries(&self) -> &[Part] {
        &self.draft
    }

    /// Replace one storage while retaining its physical graph identity.
    pub fn set(&mut self, index: usize, value: CustomData) -> Result<bool> {
        self.set_shared(index, value.into())
    }

    fn set_shared(&mut self, index: usize, value: CustomDataView) -> Result<bool> {
        validate_value(&value, &self.limits)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| invalid(format!("custom-data index {index} is absent")))?;
        if current.value == value {
            return Ok(false);
        }
        let mut candidate = self.draft.clone();
        candidate[index].value = value;
        validate_entries(self.target, &candidate, &self.limits)?;
        if candidate[index].id().is_empty() && self.connections.count(current.id()) > 0 {
            return Err(Error::Unsupported {
                feature: "renaming a referenced Custom Data storage to the unreferencable empty UID",
            });
        }
        let connections = self
            .connections
            .rebind(current.id(), candidate[index].id())?;
        connections.validate_ids(&candidate.iter().map(Part::id).collect())?;
        self.draft = candidate;
        self.connections = connections;
        Ok(true)
    }

    /// Edit typed properties while retaining inert payload bytes.
    pub fn edit_properties(
        &mut self,
        index: usize,
        edit: impl FnOnce(&mut super::Properties) -> Result<()>,
    ) -> Result<bool> {
        let mut value = self
            .draft
            .get(index)
            .map(|entry| entry.value.clone())
            .ok_or_else(|| invalid(format!("custom-data index {index} is absent")))?;
        let mut properties = value.properties.as_ref().clone();
        edit(&mut properties)?;
        value.properties = Arc::new(properties);
        self.set_shared(index, value)
    }

    /// Replace one inert payload without exposing package relationships.
    pub fn set_data(&mut self, index: usize, data: Vec<u8>) -> Result<bool> {
        let mut value = self
            .draft
            .get(index)
            .map(|entry| entry.value.clone())
            .ok_or_else(|| invalid(format!("custom-data index {index} is absent")))?;
        value.data = Arc::new(data);
        self.set_shared(index, value)
    }

    /// Insert a new storage with an allocated, source-bound graph.
    pub fn insert(&mut self, value: CustomData) -> Result<usize> {
        let value = CustomDataView::from(value);
        validate_value(&value, &self.limits)?;
        if self
            .draft
            .iter()
            .any(|entry| entry.id() == value.properties.id)
        {
            return Err(invalid("Custom Data storage IDs must be unique"));
        }
        let workbook = self.target.main_document_part()?;
        let properties_part_name = allocate_part_name(self.target, &self.draft, "properties")?;
        let data_part_name = allocate_part_name(self.target, &self.draft, "data")?;
        let workbook_relationship_id =
            allocate_relationship_id(workbook, &self.draft, "rIdCustomData", true);
        let data_relationship_id = "rIdCustomData".to_owned();
        let workbook_relationship_target =
            properties_part_name.relative_ref(workbook.partname().base_uri());
        let data_relationship_target = data_part_name.relative_ref(properties_part_name.base_uri());
        let mut candidate = self.draft.clone();
        candidate.push(Part {
            value,
            properties_part_name,
            data_part_name,
            workbook_relationship_id,
            workbook_relationship_target,
            data_relationship_id,
            data_relationship_target,
            source_properties: Arc::new(Vec::new()),
            source_properties_proof: None,
            source_data: Arc::new(Vec::new()),
            properties_relationships: None,
            data_relationships: None,
        });
        validate_entries(self.target, &candidate, &self.limits)?;
        self.draft = candidate;
        Ok(self.draft.len() - 1)
    }

    /// Insert or replace the storage with the same UID.
    pub fn upsert(&mut self, value: CustomData) -> Result<usize> {
        if let Some(index) = self
            .draft
            .iter()
            .position(|entry| entry.id() == value.properties.id)
        {
            self.set(index, value)?;
            Ok(index)
        } else {
            self.insert(value)
        }
    }

    /// Rename a storage and all recognized connection references atomically.
    pub fn rename(&mut self, index: usize, id: impl Into<String>) -> Result<bool> {
        let id = id.into();
        self.edit_properties(index, |properties| {
            properties.id = id;
            Ok(())
        })
    }

    /// Remove one storage, refusing referenced removal without a disposition.
    pub fn remove(&mut self, index: usize) -> Result<Option<Part>> {
        self.remove_with(index, RemovalDisposition::RejectReferenced)
    }

    /// Remove a storage and explicitly detach or retarget its connections.
    pub fn remove_with(
        &mut self,
        index: usize,
        disposition: RemovalDisposition,
    ) -> Result<Option<Part>> {
        if index >= self.draft.len() {
            return Ok(None);
        }
        let mut candidate = self.draft.clone();
        let removed = candidate.remove(index);
        validate_entries(self.target, &candidate, &self.limits)?;
        let count = self.connections.count(removed.id());
        let replacement = match &disposition {
            RemovalDisposition::RejectReferenced if count > 0 => {
                return Err(Error::CustomDataReferenced {
                    id: removed.id().into(),
                    connections: count,
                });
            },
            RemovalDisposition::RejectReferenced | RemovalDisposition::DetachConnections => "",
            RemovalDisposition::RetargetConnections(id) => {
                if id.is_empty() || !candidate.iter().any(|entry| entry.id() == id) {
                    return Err(invalid(
                        "connection retarget UID must name a remaining Custom Data storage",
                    ));
                }
                id.as_str()
            },
        };
        let connections = self.connections.rebind(removed.id(), replacement)?;
        connections.validate_ids(&candidate.iter().map(Part::id).collect())?;
        self.draft = candidate;
        self.connections = connections;
        Ok(Some(removed))
    }

    /// Whether staged typed values or graph identities differ.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.before.same_semantics(&self.draft) || self.connections.is_changed()
    }

    /// Validate the source closure and atomically publish the staged graph.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            return Ok(Commit::new(
                self.before.clone(),
                Patch::new(self.before.clone(), self.before.clone()),
                false,
            ));
        }
        if self.target.is_signed() || self.target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let current = Snapshot::load_with_limits(self.target, &self.limits)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "Custom Data source closure".into(),
            });
        }
        let mut candidate = self.target.clone();
        let content_types = transition_content_types(
            &self.before.source.content_types,
            &self.before.entries,
            &self.draft,
        )?;
        apply_entries(
            &mut candidate,
            self.before.entries(),
            &self.draft,
            &self.limits,
            &content_types,
        )?;
        self.connections.publish(&mut candidate)?;
        let snapshot = Snapshot::load_with_limits(&candidate, &self.limits)?;
        if !snapshot.same_published_semantics(&self.draft)
            || !snapshot.source.connections.same_values(&self.connections)
        {
            return Err(invalid("Custom Data publication changed staged semantics"));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}

/// Exact, source-checked Custom Data catalog replacement.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source state required before application.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by application.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is a source-byte no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return an exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply atomically after source-closure validation.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load_with_limits(target, &self.before.limits)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "Custom Data source closure".into(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if target.is_signed() || target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let mut candidate = target.clone();
        apply_snapshot(&mut candidate, &self.after)?;
        let resulting = Snapshot::load_with_limits(&candidate, &self.after.limits)?;
        if !resulting.same_source(&self.after) {
            return Err(invalid(
                "Custom Data patch publication changed source bytes",
            ));
        }
        *target = candidate;
        Ok(())
    }
}

/// Successful Custom Data transaction publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    /// Whether the catalog changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }
}

fn load_entries(package: &OpcPackage, limits: &Limits) -> Result<Vec<Part>> {
    let workbook = package.main_document_part()?;
    validate_feature_relationships(package, workbook)?;
    let properties_parts = package
        .iter_parts()
        .filter(|part| part.content_type() == PROPERTIES_CONTENT_TYPE)
        .collect::<Vec<_>>();
    if properties_parts.len() > limits.max_storages() {
        return Err(invalid("Custom Data storage count exceeds the size limit"));
    }
    let mut ids = HashSet::new();
    let mut data_targets = Vec::new();
    data_targets
        .try_reserve(properties_parts.len())
        .map_err(|_| invalid("Custom Data payload target allocation failed"))?;
    let mut entries = Vec::with_capacity(properties_parts.len());
    for properties_part in properties_parts {
        let owners = workbook
            .rels()
            .iter()
            .filter(|relationship| {
                relationship.reltype() == PROPERTIES_RELATIONSHIP_TYPE
                    && !relationship.is_external()
                    && relationship
                        .target_partname()
                        .ok()
                        .is_some_and(|target| target.is_equivalent_to(properties_part.partname()))
            })
            .collect::<Vec<_>>();
        if owners.len() != 1 {
            return Err(invalid(format!(
                "Custom Data Properties part '{}' must have exactly one workbook owner",
                properties_part.partname()
            )));
        }
        let owner = owners[0];
        let data_relationships = properties_part
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == DATA_RELATIONSHIP_TYPE)
            .collect::<Vec<_>>();
        if data_relationships.len() != 1 || data_relationships[0].is_external() {
            return Err(invalid(format!(
                "Custom Data Properties part '{}' must have one internal data relationship",
                properties_part.partname()
            )));
        }
        if properties_part
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() != DATA_RELATIONSHIP_TYPE)
        {
            return Err(invalid(
                "Custom Data Properties has an unexpected relationship",
            ));
        }
        let data_relationship = data_relationships[0];
        let data_target = data_relationship.target_partname()?;
        let data_part = package.get_part(&data_target)?;
        // Keep the package's actual Part identity separate from the
        // relationship's lexical target_ref, which is retained below for
        // source-conflict checks.
        let data_part_name = data_part.partname().clone();
        if data_part.content_type() != DATA_CONTENT_TYPE {
            return Err(invalid(format!(
                "Custom Data payload '{}' has content type '{}', expected '{DATA_CONTENT_TYPE}'",
                data_part_name,
                data_part.content_type()
            )));
        }
        if data_part.blob().len() > limits.max_payload_bytes() {
            return Err(invalid("Custom Data payload exceeds the size limit"));
        }
        if !data_part.rels().is_empty() {
            return Err(invalid("Custom Data payload has outbound relationships"));
        }
        let properties = parse_properties(properties_part.blob())?;
        if !ids.insert(properties.id.clone()) {
            return Err(invalid("Custom Data storage IDs must be unique"));
        }
        if data_targets
            .iter()
            .any(|target: &PackURI| target.is_equivalent_to(&data_part_name))
        {
            return Err(invalid(
                "a Custom Data payload cannot be shared by multiple properties parts",
            ));
        }
        data_targets.push(data_part_name.clone());
        entries.push(Part {
            value: CustomDataView {
                properties: Arc::new(properties),
                data: data_part.blob_arc(),
            },
            properties_part_name: properties_part.partname().clone(),
            data_part_name,
            workbook_relationship_id: owner.r_id().to_owned(),
            workbook_relationship_target: owner.target_ref().to_owned(),
            data_relationship_id: data_relationship.r_id().to_owned(),
            data_relationship_target: data_relationship.target_ref().to_owned(),
            source_properties: properties_part.blob_arc(),
            source_properties_proof: Some(Arc::new(
                package.source_xml_part(properties_part.partname())?,
            )),
            source_data: data_part.blob_arc(),
            properties_relationships: Some(
                package.source_relationships(properties_part.partname())?,
            ),
            data_relationships: Some(package.source_relationships(data_part.partname())?),
        });
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == DATA_CONTENT_TYPE)
    {
        if !data_targets
            .iter()
            .any(|target| target.is_equivalent_to(part.partname()))
        {
            return Err(invalid(format!(
                "orphan Custom Data payload '{}'",
                part.partname()
            )));
        }
    }
    entries.sort_unstable_by(|left, right| left.id().cmp(right.id()));
    Ok(entries)
}

fn validate_feature_relationships(package: &OpcPackage, workbook: &dyn OpcPart) -> Result<()> {
    for relationship in package.rels().iter() {
        validate_feature_relationship(package, None, workbook.partname(), relationship)?;
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            validate_feature_relationship(
                package,
                Some(source),
                workbook.partname(),
                relationship,
            )?;
        }
    }
    Ok(())
}

fn validate_feature_relationship(
    package: &OpcPackage,
    source: Option<&dyn OpcPart>,
    workbook_name: &PackURI,
    relationship: &litchi_opc::Relationship,
) -> Result<()> {
    let kind = match relationship.reltype() {
        PROPERTIES_RELATIONSHIP_TYPE => PROPERTIES_RELATIONSHIP_TYPE,
        DATA_RELATIONSHIP_TYPE => DATA_RELATIONSHIP_TYPE,
        _ => return Ok(()),
    };
    if relationship.is_external()
        || relationship.target_query().is_some()
        || relationship.target_fragment().is_some()
    {
        return Err(invalid(format!(
            "Custom Data relationship '{kind}' must target an internal part without a query or fragment"
        )));
    }
    let target = relationship.target_partname()?;
    let target_part = package.get_part(&target)?;
    if kind == PROPERTIES_RELATIONSHIP_TYPE {
        if source.is_none_or(|part| !part.partname().is_equivalent_to(workbook_name)) {
            return Err(invalid(
                "Custom Data Properties relationship must originate from the workbook",
            ));
        }
        if target_part.content_type() != PROPERTIES_CONTENT_TYPE {
            return Err(invalid(
                "Custom Data Properties relationship targets a non-Properties part",
            ));
        }
    } else {
        let Some(source) = source else {
            return Err(invalid(
                "Custom Data relationship must originate from a Properties part",
            ));
        };
        if source.content_type() != PROPERTIES_CONTENT_TYPE {
            return Err(invalid(
                "Custom Data relationship must originate from a Properties part",
            ));
        }
        if target_part.content_type() != DATA_CONTENT_TYPE {
            return Err(invalid(
                "Custom Data relationship targets a non-payload part",
            ));
        }
    }
    Ok(())
}

fn validate_value(value: &CustomDataView, limits: &Limits) -> Result<()> {
    if value.data.len() > limits.max_payload_bytes() {
        return Err(invalid("Custom Data payload exceeds the size limit"));
    }
    validate_source_properties(&value.properties)
}

fn validate_entries(package: &OpcPackage, entries: &[Part], limits: &Limits) -> Result<()> {
    if entries.len() > limits.max_storages() {
        return Err(invalid("Custom Data storage count exceeds the size limit"));
    }
    let mut ids = HashSet::new();
    for entry in entries {
        validate_value(&entry.value, limits)?;
        if !ids.insert(entry.id().to_owned()) {
            return Err(invalid("Custom Data storage IDs must be unique"));
        }
        if entry
            .properties_part_name
            .is_equivalent_to(&entry.data_part_name)
            || package.get_part(&entry.properties_part_name).is_ok()
                && entry.source_properties.is_empty()
            || package.get_part(&entry.data_part_name).is_ok() && entry.source_properties.is_empty()
        {
            return Err(invalid(
                "Custom Data part identity collides with the package",
            ));
        }
    }
    Ok(())
}

fn same_identity(left: &Part, right: &Part) -> bool {
    left.properties_part_name
        .is_equivalent_to(&right.properties_part_name)
        && left.data_part_name.is_equivalent_to(&right.data_part_name)
        && left.workbook_relationship_id == right.workbook_relationship_id
        && left.workbook_relationship_target == right.workbook_relationship_target
        && left.data_relationship_id == right.data_relationship_id
        && left.data_relationship_target == right.data_relationship_target
}

fn incoming_relationships(
    package: &OpcPackage,
    entries: &[Part],
) -> Result<Vec<RelationshipState>> {
    let targets = entries
        .iter()
        .flat_map(|entry| [&entry.properties_part_name, &entry.data_part_name])
        .cloned()
        .collect::<Vec<_>>();
    let mut values = Vec::new();
    let package_source = PackURI::new("/").map_err(invalid)?;
    for relationship in package.rels().iter() {
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname()?;
        if targets
            .iter()
            .any(|candidate| candidate.is_equivalent_to(&target))
        {
            values.push(RelationshipState {
                source: package_source.clone(),
                id: relationship.r_id().to_owned(),
                reltype: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                mode: relationship.target_mode(),
            });
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            if targets
                .iter()
                .any(|candidate| candidate.is_equivalent_to(&target))
            {
                values.push(RelationshipState {
                    source: source.partname().clone(),
                    id: relationship.r_id().to_owned(),
                    reltype: relationship.reltype().to_owned(),
                    target: relationship.target_ref().to_owned(),
                    mode: relationship.target_mode(),
                });
            }
        }
    }
    values.sort_by(|left, right| {
        left.source
            .as_str()
            .cmp(right.source.as_str())
            .then_with(|| left.id.cmp(&right.id))
    });
    Ok(values)
}

fn allocate_part_name(package: &OpcPackage, entries: &[Part], kind: &str) -> Result<PackURI> {
    let mut index = 1u32;
    loop {
        let candidate =
            PackURI::new(format!("/xl/customData/{kind}{index}.xml")).map_err(invalid)?;
        let candidate = if kind == "data" {
            PackURI::new(format!("/xl/customData/data{index}.bin")).map_err(invalid)?
        } else {
            candidate
        };
        if package.get_part(&candidate).is_err()
            && !entries.iter().any(|entry| {
                entry.properties_part_name.is_equivalent_to(&candidate)
                    || entry.data_part_name.is_equivalent_to(&candidate)
            })
        {
            return Ok(candidate);
        }
        index = index
            .checked_add(1)
            .ok_or_else(|| invalid("Custom Data part-name space exhausted"))?;
    }
}

fn allocate_relationship_id(
    workbook: &dyn OpcPart,
    entries: &[Part],
    prefix: &str,
    _include_existing: bool,
) -> String {
    let mut used = workbook
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<HashSet<_>>();
    used.extend(
        entries
            .iter()
            .map(|entry| entry.workbook_relationship_id.clone()),
    );
    let mut candidate = prefix.to_owned();
    let mut suffix = 2u32;
    while !used.insert(candidate.clone()) {
        candidate = format!("{prefix}{suffix}");
        suffix = suffix.saturating_add(1);
    }
    candidate
}

fn apply_entries(
    package: &mut OpcPackage,
    before: &[Part],
    after: &[Part],
    limits: &Limits,
    content_types: &OwnedContentTypes,
) -> Result<()> {
    let before_map = before
        .iter()
        .map(|entry| (identity_key(entry), entry))
        .collect::<std::collections::HashMap<_, _>>();
    let after_map = after
        .iter()
        .map(|entry| (identity_key(entry), entry))
        .collect::<std::collections::HashMap<_, _>>();
    for entry in before {
        if !after_map.contains_key(&identity_key(entry)) {
            remove_entry(package, entry)?;
        }
    }
    for entry in after {
        let key = identity_key(entry);
        if let Some(previous) = before_map.get(&key) {
            let properties_changed = previous.value.properties != entry.value.properties;
            let data_changed = previous.value.data != entry.value.data;
            if properties_changed {
                let proof = if !entry.source_properties.is_empty()
                    && entry.source_properties != previous.source_properties
                {
                    Some(
                        entry
                            .source_properties_proof
                            .as_ref()
                            .ok_or_else(|| invalid("missing properties source provenance"))?
                            .as_ref()
                            .clone(),
                    )
                } else if previous.value.properties.extension_list
                    == entry.value.properties.extension_list
                {
                    Some(rewrite_id(
                        previous
                            .source_properties_proof
                            .as_ref()
                            .ok_or_else(|| invalid("missing properties source provenance"))?,
                        entry.id(),
                    )?)
                } else {
                    let mut proof = rewrite_extension_list(
                        previous
                            .source_properties_proof
                            .as_ref()
                            .ok_or_else(|| invalid("missing properties source provenance"))?,
                        entry.value.properties.extension_list.as_ref(),
                    )?;
                    if previous.id() != entry.id() {
                        proof = rewrite_id(&proof, entry.id())?;
                    }
                    Some(proof)
                };
                if let Some(proof) = proof {
                    package.try_replace_owned_xml_part(&previous.source_properties, proof)?;
                } else {
                    package
                        .get_part_mut(&entry.properties_part_name)?
                        .set_blob(write_properties(&entry.value.properties)?);
                };
            }
            if data_changed {
                let bytes =
                    if !entry.source_data.is_empty() && entry.source_data != previous.source_data {
                        Arc::clone(&entry.source_data)
                    } else {
                        Arc::clone(&entry.value.data)
                    };
                package
                    .get_part_mut(&entry.data_part_name)?
                    .set_blob_shared(bytes);
            }
        } else {
            add_entry(package, entry, limits)?;
        }
    }
    let current_content_types = package.source_content_types()?;
    package.try_replace_content_types(current_content_types.bytes(), content_types)?;
    Ok(())
}

fn transition_content_types(
    before: &OwnedContentTypes,
    before_entries: &[Part],
    after_entries: &[Part],
) -> Result<OwnedContentTypes> {
    let before_map = before_entries
        .iter()
        .map(|entry| (identity_key(entry), entry))
        .collect::<std::collections::HashMap<_, _>>();
    let after_map = after_entries
        .iter()
        .map(|entry| (identity_key(entry), entry))
        .collect::<std::collections::HashMap<_, _>>();
    let mut removed = Vec::new();
    for entry in before_entries {
        if !after_map.contains_key(&identity_key(entry)) {
            removed.push(entry.properties_part_name.clone());
            removed.push(entry.data_part_name.clone());
        }
    }
    let mut additions = Vec::new();
    for entry in after_entries {
        if !before_map.contains_key(&identity_key(entry)) {
            additions.push((&entry.properties_part_name, PROPERTIES_CONTENT_TYPE));
            additions.push((&entry.data_part_name, DATA_CONTENT_TYPE));
        }
    }
    let maximum = litchi_opc::ReadLimits::default().max_content_types_bytes();
    let mut result = before.without_parts(&removed, maximum)?;
    if !additions.is_empty() {
        result = result.with_part_overrides(&additions, maximum)?;
    }
    Ok(result)
}

fn identity_key(entry: &Part) -> (PackURI, PackURI, String, String, String, String) {
    (
        entry.properties_part_name.clone(),
        entry.data_part_name.clone(),
        entry.workbook_relationship_id.clone(),
        entry.workbook_relationship_target.clone(),
        entry.data_relationship_id.clone(),
        entry.data_relationship_target.clone(),
    )
}

fn add_entry(package: &mut OpcPackage, entry: &Part, limits: &Limits) -> Result<()> {
    if package.get_part(&entry.properties_part_name).is_ok()
        || package.get_part(&entry.data_part_name).is_ok()
    {
        return Err(invalid("Custom Data part already exists"));
    }
    let workbook = package.main_document_part()?;
    if workbook
        .rels()
        .get(&entry.workbook_relationship_id)
        .is_some()
    {
        return Err(invalid(
            "Custom Data workbook relationship ID already exists",
        ));
    }
    let workbook_name = workbook.partname().clone();
    // Capture the complete owner member before adding the new edge. The
    // relationship token keeps comments, PIs, prefixes and attribute order
    // in the workbook `.rels` member when the new edge is spliced in.
    let workbook_relationships = package.source_relationships(&workbook_name)?;
    let properties_xml = if entry.source_properties.is_empty() {
        Arc::new(write_properties(&entry.value.properties)?)
    } else {
        Arc::clone(&entry.source_properties)
    };
    if properties_xml.len() > 4 * 1024 * 1024 || entry.value.data.len() > limits.max_payload_bytes()
    {
        return Err(invalid("Custom Data output exceeds the size limit"));
    }
    if let Some(proof) = &entry.source_properties_proof {
        package.try_add_owned_xml_part(proof.as_ref().clone())?;
    } else {
        package.try_add_part(Box::new(BlobPart::new_shared(
            entry.properties_part_name.clone(),
            PROPERTIES_CONTENT_TYPE.into(),
            properties_xml,
        )))?;
    }
    package.try_add_part(Box::new(BlobPart::new_shared(
        entry.data_part_name.clone(),
        DATA_CONTENT_TYPE.into(),
        Arc::clone(&entry.value.data),
    )))?;
    let properties_relationships = package.source_relationships(&entry.properties_part_name)?;
    let properties_replacement = properties_relationships.with_relationship(
        DATA_RELATIONSHIP_TYPE,
        &entry.data_relationship_target,
        &entry.data_relationship_id,
        TargetMode::Internal,
        usize::MAX,
    )?;
    package.try_replace_relationships(&properties_relationships, &properties_replacement)?;
    let workbook_replacement = workbook_relationships.with_relationship(
        PROPERTIES_RELATIONSHIP_TYPE,
        &entry.workbook_relationship_target,
        &entry.workbook_relationship_id,
        TargetMode::Internal,
        usize::MAX,
    )?;
    package.try_replace_relationships(&workbook_relationships, &workbook_replacement)?;
    Ok(())
}

fn remove_entry(package: &mut OpcPackage, entry: &Part) -> Result<()> {
    let workbook = package.main_document_part()?;
    let owner = workbook
        .rels()
        .get(&entry.workbook_relationship_id)
        .ok_or_else(|| invalid("Custom Data workbook relationship is absent"))?;
    if owner.target_ref() != entry.workbook_relationship_target
        || !owner
            .target_partname()?
            .is_equivalent_to(&entry.properties_part_name)
    {
        return Err(invalid("Custom Data workbook relationship changed"));
    }
    let properties = package.get_part(&entry.properties_part_name)?;
    let properties_part_name = properties.partname().clone();
    let data = properties
        .rels()
        .get(&entry.data_relationship_id)
        .ok_or_else(|| invalid("Custom Data payload relationship is absent"))?;
    if data.target_ref() != entry.data_relationship_target
        || !data
            .target_partname()?
            .is_equivalent_to(&entry.data_part_name)
    {
        return Err(invalid("Custom Data payload relationship changed"));
    }
    let workbook_name = workbook.partname().clone();
    let workbook_relationships = package.source_relationships(&workbook_name)?;
    let properties_relationships = package.source_relationships(&properties_part_name)?;
    let workbook_replacement =
        workbook_relationships.without_relationship(&entry.workbook_relationship_id, usize::MAX)?;
    package.try_replace_relationships(&workbook_relationships, &workbook_replacement)?;
    let properties_replacement =
        properties_relationships.without_relationship(&entry.data_relationship_id, usize::MAX)?;
    package.try_replace_relationships(&properties_relationships, &properties_replacement)?;
    if part_is_referenced(package, &entry.properties_part_name)
        || part_is_referenced(package, &entry.data_part_name)
    {
        return Err(invalid(
            "Custom Data part has an unexpected incoming relationship",
        ));
    }
    let data_part_name = package.get_part(&entry.data_part_name)?.partname().clone();
    if !package.remove_part(&properties_part_name) || !package.remove_part(&data_part_name) {
        return Err(invalid("Custom Data part is absent"));
    }
    Ok(())
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .ok()
                .is_some_and(|candidate| candidate.is_equivalent_to(target))
    }) || package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .ok()
                    .is_some_and(|candidate| candidate.is_equivalent_to(target))
        })
    })
}

fn apply_snapshot(package: &mut OpcPackage, snapshot: &Snapshot) -> Result<()> {
    // Custom Data CRUD does not touch workbook XML, but retaining the exact
    // workbook source and relationship graph makes inverse patches robust to
    // lexical relationship changes.
    let workbook_name = package.main_document_part()?.partname().clone();
    package
        .get_part_mut(&workbook_name)?
        .set_blob_shared(Arc::clone(&snapshot.source.workbook_blob));
    let current = load_entries(package, &snapshot.limits)?;
    apply_entries(
        package,
        &current,
        &snapshot.entries,
        &snapshot.limits,
        &snapshot.source.content_types,
    )?;
    restore_relationship_provenance(package, snapshot)?;
    snapshot.source.connections.restore(package)
}

fn restore_relationship_provenance(package: &mut OpcPackage, snapshot: &Snapshot) -> Result<()> {
    let restore = |package: &mut OpcPackage, source: &OwnedRelationships| -> Result<()> {
        let current = package.source_relationships(source.owner())?;
        package.try_replace_relationships(&current, source)?;
        Ok(())
    };
    restore(package, &snapshot.source.workbook_relationships)?;
    for entry in snapshot.entries.iter() {
        let properties = entry
            .properties_relationships
            .as_ref()
            .ok_or_else(|| invalid("missing Properties relationship provenance"))?;
        restore(package, properties)?;
        let data = entry
            .data_relationships
            .as_ref()
            .ok_or_else(|| invalid("missing payload relationship provenance"))?;
        restore(package, data)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::custom_data::ExtensionList;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, PackageWriter};

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    fn fixture() -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        let properties = PackURI::new("/xl/customData/item1.xml").unwrap();
        let data = PackURI::new("/xl/customData/data1.bin").unwrap();
        package.add_part(Box::new(BlobPart::new(
            workbook.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            format!(r#"<workbook xmlns="{SML}"><sheets/></workbook>"#).into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            properties.clone(),
            PROPERTIES_CONTENT_TYPE.into(),
            br#"<x14:datastoreItem xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" id="Storage-1"/>"#.to_vec(),
        )));
        package.add_part(Box::new(BlobPart::new(
            data.clone(),
            DATA_CONTENT_TYPE.into(),
            b"opaque-data".to_vec(),
        )));
        package.rels_mut().add_relationship(
            OFFICE_DOCUMENT.into(),
            "xl/workbook.xml".into(),
            "rIdWorkbook".into(),
            false,
        );
        package
            .get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .add_relationship(
                PROPERTIES_RELATIONSHIP_TYPE.into(),
                "customData/item1.xml".into(),
                "rIdCustomData".into(),
                false,
            );
        package
            .get_part_mut(&properties)
            .unwrap()
            .rels_mut()
            .add_relationship(
                DATA_RELATIONSHIP_TYPE.into(),
                "data1.bin".into(),
                "rIdCustomData".into(),
                false,
            );
        package
    }

    fn with_connections() -> OpcPackage {
        let mut package = fixture();
        let name = PackURI::new("/xl/connections.xml").unwrap();
        let xml = format!(
            r#"<?before?><s:connections xmlns:s="{SML}" xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:q="urn:opaque" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="a q"><?inside?><!--before--><s:connection id="7" type="5" refreshedVersion="3"><s:extLst><s:ext uri="{{D79990A0-CA42-45E3-83F4-45C500A0EAA5}}"><a:connection embeddedDataId = 'Storage&#45;1' q:embeddedDataId='Storage-1' culture='en-US'><!--inside--><q:opaque value='a:Kind'/></a:connection></s:ext></s:extLst></s:connection><s:connection id="8" type="5" refreshedVersion="3"><s:extLst><s:ext uri="{{D79990A0-CA42-45E3-83F4-45C500A0EAA5}}"><a:connection embeddedDataId="Storage-1"/></s:ext></s:extLst></s:connection><!--after--></s:connections><?after?>"#
        );
        package.add_part(Box::new(BlobPart::new(
            name,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml".into(),
            xml.into_bytes(),
        )));
        package
            .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"
                    .into(),
                "connections.xml".into(),
                "rIdConnections".into(),
                false,
            );
        let properties = format!(
            r#"<!--root--><datastoreItem xmlns='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main' xmlns:s='{SML}' xmlns:v='urn:values' id = 'Storage&#45;1'><extLst><s:ext uri='urn:keep'><q:item xmlns:q='urn:q' kind='v:Kind'/><!--inside--></s:ext></extLst><!--end--></datastoreItem>"#
        );
        package
            .get_part_mut(&PackURI::new("/xl/customData/item1.xml").unwrap())
            .unwrap()
            .set_blob(properties.into_bytes());
        as_source(&package)
    }

    fn as_source(package: &OpcPackage) -> OpcPackage {
        as_source_with_relationship_overrides(package, &[])
    }

    fn as_source_with_content_types(package: &OpcPackage, content_types: &[u8]) -> OpcPackage {
        as_source_with_overrides_and_content_types(package, &[], Some(content_types))
    }

    fn signed_fixture() -> OpcPackage {
        let mut package = fixture();
        let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
        package.add_part(Box::new(BlobPart::new(
            origin.clone(),
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN.into(),
            b"<origin/>".to_vec(),
        )));
        package.rels_mut().add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.into(),
            "_xmlsignatures/origin.sigs".into(),
            "rIdSignature".into(),
            false,
        );
        as_source(&package)
    }

    fn remove_signature_infrastructure(package: &mut OpcPackage) {
        let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
        package.rels_mut().remove("rIdSignature").unwrap();
        assert!(package.is_signed());
        assert!(package.remove_part(&origin));
        assert!(!package.is_signed());
        assert!(package.requires_signature_edit_policy());
    }

    // A synthetic producer ZIP, deliberately bypassing the authored compact
    // writer so that tests exercise genuine ingress provenance for spaced XML.
    // Overrides also let the relationship provenance tests publish an
    // explicitly empty `.rels` member, which an in-memory empty collection
    // cannot represent by itself.
    fn as_source_with_relationship_overrides(
        package: &OpcPackage,
        overrides: &[(&str, &[u8])],
    ) -> OpcPackage {
        as_source_with_overrides_and_content_types(package, overrides, None)
    }

    fn as_source_with_overrides_and_content_types(
        package: &OpcPackage,
        overrides: &[(&str, &[u8])],
        content_types: Option<&[u8]>,
    ) -> OpcPackage {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        let mut parts = package.iter_parts().collect::<Vec<_>>();
        parts.sort_by(|a, b| a.partname().as_str().cmp(b.partname().as_str()));
        let mut types = String::from(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        );
        for part in &parts {
            types.push_str(&format!(
                r#"<Override PartName="{}" ContentType="{}"/>"#,
                part.partname(),
                part.content_type()
            ));
            writer
                .write_stored(
                    part.partname().as_str().trim_start_matches('/'),
                    part.blob(),
                )
                .unwrap();
            if let Some((_, relationships)) = overrides
                .iter()
                .find(|(owner, _)| owner.eq_ignore_ascii_case(part.partname().as_str()))
            {
                writer
                    .write_stored(
                        part.partname()
                            .rels_uri()
                            .unwrap()
                            .as_str()
                            .trim_start_matches('/'),
                        relationships,
                    )
                    .unwrap();
            } else if !part.rels().is_empty() {
                let relationships = part.rels().to_xml();
                writer
                    .write_stored(
                        part.partname()
                            .rels_uri()
                            .unwrap()
                            .as_str()
                            .trim_start_matches('/'),
                        relationships.as_bytes(),
                    )
                    .unwrap();
            }
        }
        types.push_str("</Types>");
        writer
            .write_stored(
                "[Content_Types].xml",
                content_types.unwrap_or(types.as_bytes()),
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", package.rels().to_xml().as_bytes())
            .unwrap();
        OpcPackage::from_bytes(&writer.finish_to_bytes().unwrap()).unwrap()
    }

    #[test]
    fn changed_transaction_refuses_low_level_signature_removal_until_unsign() -> Result<()> {
        let mut package = signed_fixture();
        remove_signature_infrastructure(&mut package);

        let before = Snapshot::load(&package).unwrap();
        let noop = Transaction::new(&mut package).unwrap().commit().unwrap();
        assert!(!noop.changed());
        assert!(Snapshot::load(&package).unwrap().same_source(&before));

        let data = PackURI::new("/xl/customData/data1.bin").unwrap();
        let original_data = package.get_part(&data).unwrap().blob().to_vec();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.set_data(0, b"changed after signature removal".to_vec())?;
        assert!(matches!(transaction.commit(), Err(Error::Signed)));
        assert_eq!(package.get_part(&data).unwrap().blob(), original_data);
        assert!(package.requires_signature_edit_policy());

        package.unsign();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.set_data(0, b"changed after explicit unsign".to_vec())?;
        transaction.commit()?;
        assert_eq!(
            package.get_part(&data).unwrap().blob(),
            b"changed after explicit unsign"
        );
        Ok(())
    }

    #[test]
    fn changed_patch_refuses_low_level_signature_removal_until_unsign() {
        let mut patch_source = signed_fixture();
        patch_source.unsign();
        let patch = {
            let mut transaction = Transaction::new(&mut patch_source).unwrap();
            transaction.set_data(0, b"patch payload".to_vec()).unwrap();
            transaction.commit().unwrap().patch().clone()
        };

        let mut target = signed_fixture();
        remove_signature_infrastructure(&mut target);
        let data = PackURI::new("/xl/customData/data1.bin").unwrap();
        let original_data = target.get_part(&data).unwrap().blob().to_vec();
        assert!(matches!(patch.apply(&mut target), Err(Error::Signed)));
        assert_eq!(target.get_part(&data).unwrap().blob(), original_data);
        assert!(target.requires_signature_edit_policy());

        target.unsign();
        patch.apply(&mut target).unwrap();
        assert_eq!(target.get_part(&data).unwrap().blob(), b"patch payload");
    }

    #[test]
    fn removal_reopen_inverse_restores_exact_content_types_member() {
        let mut base = fixture();
        let unrelated = PackURI::new("/xl/unrelated.bin").unwrap();
        base.add_part(Box::new(BlobPart::new(
            unrelated,
            "application/octet-stream".into(),
            b"unrelated".to_vec(),
        )));
        let content_types = br#"<?xml version='1.0'?>
<T:Types xmlns:T='http://schemas.openxmlformats.org/package/2006/content-types'>
 <?before-manifest?>
 <T:Default ContentType='application/vnd.openxmlformats-package.relationships+xml' Extension='rels'/>
 <!-- workbook mapping -->
 <T:Override ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' PartName='/xl/workbook.xml'/>
 <?between-properties?>
 <T:Override ContentType='application/vnd.openxmlformats-officedocument.customDataProperties+xml' PartName='/xl/customData/item1.xml'/>
 <!-- payload mapping -->
 <T:Override PartName='/xl/customData/data1.bin' ContentType='application/binary'/>
 <?between-unrelated?>
 <T:Override ContentType='application/octet-stream' PartName='/xl/unrelated.bin'/>
</T:Types>"#;
        let mut package = as_source_with_content_types(&base, content_types);
        assert_eq!(
            Snapshot::load(&package)
                .unwrap()
                .source
                .content_types
                .bytes(),
            content_types
        );

        let commit = {
            let mut transaction = Transaction::new(&mut package).unwrap();
            transaction.remove(0).unwrap();
            transaction.commit().unwrap()
        };
        let removed_bytes = PackageWriter::to_bytes(&package).unwrap();
        let removed_archive = soapberry_zip::office::ArchiveReader::new(&removed_bytes).unwrap();
        let removed_content_types = removed_archive.read("[Content_Types].xml").unwrap();
        assert!(
            removed_content_types
                .windows(b"/xl/unrelated.bin".len())
                .any(|window| { window == b"/xl/unrelated.bin" })
        );
        assert!(
            removed_content_types
                .windows(b"<?between-unrelated?>".len())
                .any(|window| window == b"<?between-unrelated?>")
        );
        assert!(
            !removed_content_types
                .windows(b"/xl/customData/item1.xml".len())
                .any(|window| window == b"/xl/customData/item1.xml")
        );
        assert!(
            !removed_content_types
                .windows(b"/xl/customData/data1.bin".len())
                .any(|window| window == b"/xl/customData/data1.bin")
        );

        let mut reopened = OpcPackage::from_vec(removed_bytes).unwrap();
        commit.patch().inverse().apply(&mut reopened).unwrap();
        let restored_bytes = PackageWriter::to_bytes(&reopened).unwrap();
        let restored_archive = soapberry_zip::office::ArchiveReader::new(&restored_bytes).unwrap();
        assert_eq!(
            restored_archive.read("[Content_Types].xml").unwrap(),
            content_types
        );
    }

    #[test]
    fn rename_splices_properties_and_all_connection_ids_and_preserves_exact_inverse() {
        let mut package = with_connections();
        let before = Snapshot::load(&package).unwrap();
        assert_eq!(before.connection_references("Storage-1"), 2);
        let connections = PackURI::new("/xl/connections.xml").unwrap();
        let properties = PackURI::new("/xl/customData/item1.xml").unwrap();
        let connection_xml =
            String::from_utf8(package.get_part(&connections).unwrap().blob().to_vec()).unwrap();
        let property_xml =
            String::from_utf8(package.get_part(&properties).unwrap().blob().to_vec()).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.rename(0, "temporary").unwrap();
        let final_id = "O'Brien & _x0041_\u{1}😀";
        transaction.rename(0, final_id).unwrap();
        // Staging identities does not serialize connection parts repeatedly.
        assert_eq!(
            transaction.target.get_part(&connections).unwrap().blob(),
            connection_xml.as_bytes()
        );
        let commit = transaction.commit().unwrap();
        assert_eq!(commit.snapshot().connection_references(final_id), 2);
        assert!(commit.snapshot().find(final_id).is_some());
        let escaped =
            String::from_utf8(crate::source_attributes::escaped_xstring(final_id)).unwrap();
        assert_eq!(
            package.get_part(&connections).unwrap().blob(),
            connection_xml
                .replace(
                    "embeddedDataId = 'Storage&#45;1'",
                    &format!("embeddedDataId = '{escaped}'")
                )
                .replace(
                    "embeddedDataId=\"Storage-1\"",
                    &format!("embeddedDataId=\"{escaped}\"")
                )
                .as_bytes()
        );
        assert_eq!(
            package.get_part(&properties).unwrap().blob(),
            property_xml
                .replace("id = 'Storage&#45;1'", &format!("id = '{escaped}'"))
                .as_bytes()
        );
        commit.patch().inverse().apply(&mut package).unwrap();
        assert!(Snapshot::load(&package).unwrap().same_source(&before));
        commit.patch().apply(&mut package).unwrap();
        assert!(
            Snapshot::load(&package)
                .unwrap()
                .same_source(commit.snapshot())
        );
    }

    #[test]
    fn reference_rename_cancellation_and_failed_mutation_are_exact() {
        let mut package = with_connections();
        let before = Snapshot::load(&package).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.rename(0, "temporary").unwrap();
        assert!(transaction.rename(0, "").is_err());
        assert_eq!(transaction.entries()[0].id(), "temporary");
        transaction
            .edit_properties(0, |properties| {
                properties.id = "Storage-1".into();
                Ok(())
            })
            .unwrap();
        assert!(!transaction.is_changed());
        let commit = transaction.commit().unwrap();
        assert!(!commit.changed());
        assert!(Snapshot::load(&package).unwrap().same_source(&before));
    }

    #[test]
    fn referenced_remove_requires_disposition_and_detach_is_reversible() {
        let mut package = with_connections();
        let before = Snapshot::load(&package).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(matches!(
            transaction.remove(0),
            Err(Error::CustomDataReferenced { connections: 2, .. })
        ));
        assert!(!transaction.is_changed());
        transaction
            .remove_with(0, RemovalDisposition::DetachConnections)
            .unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.snapshot().is_empty());
        let name = PackURI::new("/xl/connections.xml").unwrap();
        let xml = String::from_utf8_lossy(package.get_part(&name).unwrap().blob());
        assert!(xml.contains("embeddedDataId = ''"));
        assert!(xml.contains("embeddedDataId=\"\""));
        assert!(xml.contains("q:embeddedDataId='Storage-1'"));
        commit.patch().inverse().apply(&mut package).unwrap();
        assert!(Snapshot::load(&package).unwrap().same_source(&before));
    }

    #[test]
    fn retarget_removal_follows_the_selected_storage_through_rename() {
        let mut package = with_connections();
        let before = Snapshot::load(&package).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(
            transaction
                .remove_with(0, RemovalDisposition::RetargetConnections("missing".into()))
                .is_err()
        );
        assert!(!transaction.is_changed());
        let second = transaction
            .insert(CustomData::new("Storage-2", vec![2]))
            .unwrap();
        assert_eq!(second, 1);
        transaction
            .remove_with(
                0,
                RemovalDisposition::RetargetConnections("Storage-2".into()),
            )
            .unwrap();
        transaction.rename(0, "Renamed").unwrap();
        let commit = transaction.commit().unwrap();
        assert_eq!(commit.snapshot().connection_references("Renamed"), 2);
        assert_eq!(commit.snapshot().len(), 1);
        commit.patch().inverse().apply(&mut package).unwrap();
        assert!(Snapshot::load(&package).unwrap().same_source(&before));
    }

    #[test]
    fn connection_only_byte_and_graph_changes_conflict_with_custom_data_patch() {
        let source = with_connections();
        let mut package = source.clone();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.set_data(0, vec![9]).unwrap();
        let commit = transaction.commit().unwrap();
        for graph in [false, true] {
            let mut stale = source.clone();
            let name = PackURI::new("/xl/connections.xml").unwrap();
            if graph {
                let workbook = stale
                    .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
                    .unwrap();
                workbook.rels_mut().add_relationship(
                    "urn:passive".into(),
                    "https://invalid.example/".into(),
                    "rIdForeign".into(),
                    true,
                );
            } else {
                let part = stale.get_part_mut(&name).unwrap();
                let xml = String::from_utf8(part.blob().to_vec())
                    .unwrap()
                    .replace("<!--before-->", "<!--different-->");
                part.set_blob(xml.into_bytes());
            }
            let mut stale = as_source(&stale);
            let before = Snapshot::load(&stale).unwrap();
            assert!(matches!(
                commit.patch().apply(&mut stale),
                Err(Error::PatchConflict { .. })
            ));
            assert!(Snapshot::load(&stale).unwrap().same_source(&before));
        }
    }

    #[test]
    fn rejects_dangling_wrong_type_and_unmodeled_connection_identity_edits() {
        let source = with_connections();
        let name = PackURI::new("/xl/connections.xml").unwrap();
        for (from, to) in [
            ("Storage&#45;1", "Missing"),
            ("type=\"5\"", "type=\"102\""),
            ("id=\"8\"", "id=\"7\""),
        ] {
            let mut package = source.clone();
            let xml = String::from_utf8(package.get_part(&name).unwrap().blob().to_vec())
                .unwrap()
                .replace(from, to);
            package
                .get_part_mut(&name)
                .unwrap()
                .set_blob(xml.into_bytes());
            assert!(Snapshot::load(&package).is_err());
        }
        let mut package = source;
        let xml = String::from_utf8(package.get_part(&name).unwrap().blob().to_vec())
            .unwrap()
            .replace("{D79990A0-CA42-45E3-83F4-45C500A0EAA5}", "urn:unmodeled");
        package
            .get_part_mut(&name)
            .unwrap()
            .set_blob(xml.into_bytes());
        let mut package = as_source(&package);
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert!(matches!(
            transaction.rename(0, "New"),
            Err(Error::Unsupported { .. })
        ));
        assert!(!transaction.is_changed());
    }

    #[test]
    fn public_package_rename_and_remove_reopen_with_resolved_references() {
        let mut package = crate::Package::from_opc(with_connections()).unwrap();
        assert!(package.rename_custom_data("Storage-1", "New").unwrap());
        assert!(matches!(
            package.remove_custom_data("New"),
            Err(Error::CustomDataReferenced { .. })
        ));
        let mut bytes = Vec::new();
        package.write_to(&mut bytes).unwrap();
        let mut reopened = crate::Package::from_bytes(bytes).unwrap();
        assert_eq!(
            reopened.custom_data().unwrap().connection_references("New"),
            2
        );
        reopened
            .remove_custom_data_with("New", RemovalDisposition::DetachConnections)
            .unwrap();
        assert!(reopened.custom_data().unwrap().is_empty());
        let output = reopened.to_bytes().unwrap();
        assert!(
            crate::Package::from_bytes(output)
                .unwrap()
                .custom_data()
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn source_bound_custom_data_crud_preserves_payload_and_inverse_bytes() {
        let mut package = fixture();
        let properties = PackURI::new("/xl/customData/item1.xml").unwrap();
        let data = PackURI::new("/xl/customData/data1.bin").unwrap();
        let before_properties = package.get_part(&properties).unwrap().blob().to_vec();
        let before_data = package.get_part(&data).unwrap().blob().to_vec();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction
            .edit_properties(0, |properties| {
                properties.id = "Storage-2".into();
                Ok(())
            })
            .unwrap();
        transaction.set_data(0, b"changed".to_vec()).unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(package.get_part(&data).unwrap().blob(), b"changed");
        commit.patch().inverse().apply(&mut package).unwrap();
        assert_eq!(
            package.get_part(&properties).unwrap().blob(),
            before_properties
        );
        assert_eq!(package.get_part(&data).unwrap().blob(), before_data);
    }

    #[test]
    fn removal_reopen_inverse_restores_exact_relationship_members() {
        let workbook_relationships = br#"<?xml version='1.0'?>
<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'>
 <?keep-workbook?> <!-- before -->
 <r:Relationship Target='customData/item1.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps' Id='rIdCustomData'></r:Relationship>
 <!-- after -->
</r:Relationships>"#;
        let properties_relationships = br#"<?xml version='1.0'?>
<p:Relationships xmlns:p='http://schemas.openxmlformats.org/package/2006/relationships'>
 <!-- properties --> <?keep-properties?>
 <p:Relationship Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData' Target='data1.bin' Id='rIdCustomData'></p:Relationship>
</p:Relationships>"#;
        let data_relationships = br#"<?xml version='1.0'?>
<d:Relationships xmlns:d='http://schemas.openxmlformats.org/package/2006/relationships'>
 <!-- explicit empty payload member --> <?keep-payload?>
</d:Relationships>"#;
        let package = as_source_with_relationship_overrides(
            &fixture(),
            &[
                ("/xl/workbook.xml", workbook_relationships),
                ("/xl/customData/item1.xml", properties_relationships),
                ("/xl/customData/data1.bin", data_relationships),
            ],
        );
        let before = Snapshot::load(&package).unwrap();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let properties_uri = PackURI::new("/xl/customData/item1.xml").unwrap();
        let data_uri = PackURI::new("/xl/customData/data1.bin").unwrap();
        assert!(
            before.entries()[0]
                .data_relationships
                .as_ref()
                .unwrap()
                .member_present()
        );

        let mut package = package;
        let commit = {
            let mut transaction = Transaction::new(&mut package).unwrap();
            transaction.remove(0).unwrap();
            transaction.commit().unwrap()
        };
        let removed = PackageWriter::to_bytes(&package).unwrap();
        let mut reopened = OpcPackage::from_vec(removed).unwrap();
        commit.patch().inverse().apply(&mut reopened).unwrap();

        let restored = PackageWriter::to_bytes(&reopened).unwrap();
        let archive = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
        assert_eq!(
            archive
                .read(workbook_uri.rels_uri().unwrap().membername())
                .unwrap(),
            workbook_relationships
        );
        assert_eq!(
            archive
                .read(properties_uri.rels_uri().unwrap().membername())
                .unwrap(),
            properties_relationships
        );
        assert_eq!(
            archive
                .read(data_uri.rels_uri().unwrap().membername())
                .unwrap(),
            data_relationships
        );
        let restored = OpcPackage::from_vec(restored).unwrap();
        assert!(Snapshot::load(&restored).unwrap().same_source(&before));
    }

    #[test]
    fn detach_reopen_inverse_restores_exact_relationships_and_connection_source() {
        let workbook_relationships = br#"<?xml version='1.0'?>
<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'>
 <?keep-workbook?>
 <r:Relationship Target='connections.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections' Id='rIdConnections'></r:Relationship>
 <!-- custom edge --> <r:Relationship Target='customData/item1.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps' Id='rIdCustomData'></r:Relationship>
</r:Relationships>"#;
        let properties_relationships = br#"<?xml version='1.0'?>
<p:Relationships xmlns:p='http://schemas.openxmlformats.org/package/2006/relationships'>
 <?keep-properties?> <p:Relationship Target='data1.bin' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData' Id='rIdCustomData'></p:Relationship>
</p:Relationships>"#;
        let data_relationships = br#"<?xml version='1.0'?>
<d:Relationships xmlns:d='http://schemas.openxmlformats.org/package/2006/relationships'><?keep-payload?></d:Relationships>"#;
        let source = with_connections();
        let connections_uri = PackURI::new("/xl/connections.xml").unwrap();
        let before_connections = source.get_part(&connections_uri).unwrap().blob().to_vec();
        let package = as_source_with_relationship_overrides(
            &source,
            &[
                ("/xl/workbook.xml", workbook_relationships),
                ("/xl/customData/item1.xml", properties_relationships),
                ("/xl/customData/data1.bin", data_relationships),
            ],
        );
        let before = Snapshot::load(&package).unwrap();
        let mut package = package;
        let commit = {
            let mut transaction = Transaction::new(&mut package).unwrap();
            transaction
                .remove_with(0, RemovalDisposition::DetachConnections)
                .unwrap();
            transaction.commit().unwrap()
        };
        let removed = PackageWriter::to_bytes(&package).unwrap();
        let mut reopened = OpcPackage::from_vec(removed).unwrap();
        commit.patch().inverse().apply(&mut reopened).unwrap();
        let restored = PackageWriter::to_bytes(&reopened).unwrap();
        let archive = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
        assert_eq!(
            archive
                .read(
                    PackURI::new("/xl/workbook.xml")
                        .unwrap()
                        .rels_uri()
                        .unwrap()
                        .membername()
                )
                .unwrap(),
            workbook_relationships
        );
        assert_eq!(
            archive
                .read(
                    PackURI::new("/xl/customData/item1.xml")
                        .unwrap()
                        .rels_uri()
                        .unwrap()
                        .membername()
                )
                .unwrap(),
            properties_relationships
        );
        assert_eq!(
            archive
                .read(
                    PackURI::new("/xl/customData/data1.bin")
                        .unwrap()
                        .rels_uri()
                        .unwrap()
                        .membername()
                )
                .unwrap(),
            data_relationships
        );
        assert_eq!(
            archive.read("xl/connections.xml").unwrap(),
            before_connections
        );
        let restored = OpcPackage::from_vec(restored).unwrap();
        assert!(Snapshot::load(&restored).unwrap().same_source(&before));
    }

    #[test]
    fn insertion_splices_the_existing_workbook_relationship_source() {
        let workbook_relationships = br#"<?xml version='1.0'?>
<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'>
 <?keep-workbook?> <r:Relationship Target='customData/item1.xml' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps' Id='rIdCustomData'></r:Relationship>
</r:Relationships>"#;
        let package = as_source_with_relationship_overrides(
            &fixture(),
            &[("/xl/workbook.xml", workbook_relationships)],
        );
        let mut package = package;
        let commit = {
            let mut transaction = Transaction::new(&mut package).unwrap();
            transaction
                .insert(CustomData::new("Storage-2", b"new-payload".to_vec()))
                .unwrap();
            transaction.commit().unwrap()
        };
        let output = PackageWriter::to_bytes(&package).unwrap();
        let archive = soapberry_zip::office::ArchiveReader::new(&output).unwrap();
        let workbook = archive
            .read(
                PackURI::new("/xl/workbook.xml")
                    .unwrap()
                    .rels_uri()
                    .unwrap()
                    .membername(),
            )
            .unwrap();
        assert!(workbook.starts_with(b"<?xml version='1.0'?>"));
        assert!(
            workbook
                .windows(b"<?keep-workbook?>".len())
                .any(|window| { window == b"<?keep-workbook?>" })
        );
        assert!(
            workbook
                .windows(b"Target='customData/item1.xml'".len())
                .any(|window| { window == b"Target='customData/item1.xml'" })
        );
        assert!(commit.changed());
    }

    #[test]
    fn custom_data_insert_and_remove_are_graph_atomic() {
        let mut package = fixture();
        let mut transaction = Transaction::new(&mut package).unwrap();
        let index = transaction
            .insert(CustomData::new("Storage-2", b"payload".to_vec()))
            .unwrap();
        assert_eq!(index, 1);
        let commit = transaction.commit().unwrap();
        assert_eq!(Snapshot::load(&package).unwrap().len(), 2);
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.remove(0).unwrap();
        transaction.commit().unwrap();
        assert_eq!(Snapshot::load(&package).unwrap().len(), 1);
        assert!(commit.patch().inverse().apply(&mut package).is_err());
    }

    #[test]
    fn metadata_edits_and_snapshots_share_payload_storage() {
        let mut package = fixture();
        let data_name = PackURI::new("/xl/customData/data1.bin").unwrap();
        let data = package.get_part(&data_name).unwrap().blob_arc();
        let snapshot = Snapshot::load(&package).unwrap();
        let cloned = snapshot.clone();
        assert!(Arc::ptr_eq(&snapshot.entries, &cloned.entries));
        assert!(Arc::ptr_eq(&snapshot.entries[0].value.data, &data));
        let view = snapshot.entries[0].value().clone();
        assert!(Arc::ptr_eq(
            &view.properties,
            &snapshot.entries[0].value.properties
        ));
        let mut transaction = Transaction::new(&mut package).unwrap();
        assert_eq!(
            transaction.entries()[0].value().data().as_ptr(),
            data.as_ptr()
        );
        transaction
            .edit_properties(0, |properties| {
                properties.id = "Storage-2".into();
                Ok(())
            })
            .unwrap();
        assert_eq!(
            transaction.entries()[0].value().data().as_ptr(),
            data.as_ptr()
        );
        assert_eq!(snapshot.entries[0].value().properties().id, "Storage-1");
        let commit = transaction.commit().unwrap();
        assert!(Arc::ptr_eq(&commit.snapshot.entries[0].value.data, &data));
        assert!(Arc::ptr_eq(
            &package.get_part(&data_name).unwrap().blob_arc(),
            &data
        ));
        commit.patch().inverse().apply(&mut package).unwrap();
        assert!(Arc::ptr_eq(
            &package.get_part(&data_name).unwrap().blob_arc(),
            &data
        ));
    }

    #[test]
    fn payload_replacement_and_insertion_move_buffers_into_publication() {
        let mut package = fixture();
        let data_name = PackURI::new("/xl/customData/data1.bin").unwrap();
        let before = package.get_part(&data_name).unwrap().blob_arc();
        let replacement = vec![0x5a; 1024 * 1024];
        let replacement_ptr = replacement.as_ptr();
        let inserted = vec![0xa5; 1024 * 1024];
        let inserted_ptr = inserted.as_ptr();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.set_data(0, replacement).unwrap();
        transaction
            .insert(CustomData::new("Storage-2", inserted))
            .unwrap();
        assert_eq!(
            transaction.entries()[0].value().data().as_ptr(),
            replacement_ptr
        );
        assert_eq!(
            transaction.entries()[1].value().data().as_ptr(),
            inserted_ptr
        );
        let commit = transaction.commit().unwrap();
        assert_eq!(
            commit.snapshot.entries[0].value().data().as_ptr(),
            replacement_ptr
        );
        assert_eq!(
            commit.snapshot.entries[1].value().data().as_ptr(),
            inserted_ptr
        );
        let forward = commit.patch().clone();
        forward.inverse().apply(&mut package).unwrap();
        assert!(Arc::ptr_eq(
            &package.get_part(&data_name).unwrap().blob_arc(),
            &before
        ));
        forward.apply(&mut package).unwrap();
        assert_eq!(
            package.get_part(&data_name).unwrap().blob().as_ptr(),
            replacement_ptr
        );
    }

    #[test]
    fn removing_empty_payload_has_an_exact_inverse() {
        let mut package = fixture();
        let data_name = PackURI::new("/xl/customData/data1.bin").unwrap();
        package
            .get_part_mut(&data_name)
            .unwrap()
            .set_blob(Vec::new());
        let properties_name = PackURI::new("/xl/customData/item1.xml").unwrap();
        let properties = br#"<datastoreItem xmlns='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main' id='Storage-1'><!--retained--></datastoreItem>"#.to_vec();
        package
            .get_part_mut(&properties_name)
            .unwrap()
            .set_blob(properties.clone());
        let mut transaction = Transaction::new(&mut package).unwrap();
        // An existing empty binary part is not a newly allocated graph identity.
        transaction.set_data(0, vec![42]).unwrap();
        transaction.set_data(0, Vec::new()).unwrap();
        assert!(!transaction.is_changed());
        transaction.remove(0).unwrap();
        let commit = transaction.commit().unwrap();
        commit.patch().inverse().apply(&mut package).unwrap();
        assert_eq!(
            package.get_part(&properties_name).unwrap().blob(),
            properties
        );
        assert!(package.get_part(&data_name).unwrap().blob().is_empty());
    }

    #[test]
    fn case_aliased_graph_targets_load_and_remove() {
        let mut package = fixture();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        let properties = PackURI::new("/xl/customData/item1.xml").unwrap();
        package
            .get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .retarget("rIdCustomData", "customData/ITEM1.XML".into())
            .unwrap();
        package
            .get_part_mut(&properties)
            .unwrap()
            .rels_mut()
            .retarget("rIdCustomData", "DATA1.BIN".into())
            .unwrap();

        let snapshot = Snapshot::load(&package).unwrap();
        assert_eq!(snapshot.len(), 1);
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.remove(0).unwrap();
        transaction.commit().unwrap();
        assert!(package.get_part(&properties).is_err());
        assert!(
            package
                .get_part(&PackURI::new("/xl/customData/data1.bin").unwrap())
                .is_err()
        );
    }

    #[test]
    fn case_aliased_foreign_inbound_edge_refuses_removal_atomically() {
        let mut package = fixture();
        let foreign = PackURI::new("/xl/foreign.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            foreign.clone(),
            "application/xml".into(),
            b"<foreign/>".to_vec(),
        )));
        package
            .get_part_mut(&foreign)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:foreign-custom-data-edge".into(),
                "customData/ITEM1.XML".into(),
                "rIdForeign".into(),
                false,
            );
        let before = Snapshot::load(&package).unwrap();
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction.remove(0).unwrap();
        assert!(transaction.commit().is_err());
        assert!(Snapshot::load(&package).unwrap().same_source(&before));
        assert!(
            package
                .get_part(&PackURI::new("/xl/customData/item1.xml").unwrap())
                .is_ok()
        );
        assert!(
            package
                .get_part(&foreign)
                .unwrap()
                .rels()
                .get("rIdForeign")
                .is_some()
        );
    }

    #[test]
    fn malformed_type_five_connections_are_rejected_before_custom_data_edits() {
        let mut package = with_connections();
        let connections = PackURI::new("/xl/connections.xml").unwrap();
        let xml = String::from_utf8(package.get_part(&connections).unwrap().blob().to_vec())
            .unwrap()
            .replace(" type=\"5\" refreshedVersion=\"3\"", " type=\"5\"");
        package
            .get_part_mut(&connections)
            .unwrap()
            .set_blob(xml.into_bytes());
        assert!(Snapshot::load(&package).is_err());
    }

    #[test]
    fn custom_data_load_runs_the_complete_connections_graph_validator() {
        let mut root_without_part = fixture();
        root_without_part.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"
                .into(),
            "xl/connections.xml".into(),
            "rIdRootConnections".into(),
            false,
        );
        assert!(Snapshot::load(&root_without_part).is_err());

        let mut case_alias = with_connections();
        case_alias
            .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .rels_mut()
            .retarget("rIdConnections", "CONNECTIONS.XML".into())
            .unwrap();
        assert_eq!(Snapshot::load(&case_alias).unwrap().len(), 1);

        let mut root_edge = with_connections();
        root_edge.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"
                .into(),
            "xl/connections.xml".into(),
            "rIdRootConnections".into(),
            false,
        );
        assert!(Snapshot::load(&root_edge).is_err());

        let mut outbound = with_connections();
        outbound
            .get_part_mut(&PackURI::new("/xl/connections.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:unexpected".into(),
                "foreign.xml".into(),
                "rIdUnexpected".into(),
                false,
            );
        assert!(Snapshot::load(&outbound).is_err());

        let mut orphan = with_connections();
        orphan
            .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .rels_mut()
            .remove("rIdConnections");
        assert!(Snapshot::load(&orphan).is_err());

        let query_table_content_type =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
        let query_table = PackURI::new("/xl/queryTables/queryTable1.xml").unwrap();
        let query_table_xml =
            format!(r#"<queryTable xmlns="{SML}" name="Query1" connectionId="7"/>"#);
        let mut query_orphan = with_connections();
        query_orphan.add_part(Box::new(BlobPart::new(
            query_table.clone(),
            query_table_content_type.into(),
            query_table_xml.as_bytes().to_vec(),
        )));
        assert!(Snapshot::load(&query_orphan).is_err());

        let mut query_multiowner = with_connections();
        query_multiowner.add_part(Box::new(BlobPart::new(
            query_table.clone(),
            query_table_content_type.into(),
            query_table_xml.into_bytes(),
        )));
        for (sheet_name, relationship_id) in [
            ("/xl/worksheets/sheet1.xml", "rIdQuery1"),
            ("/xl/worksheets/sheet2.xml", "rIdQuery2"),
        ] {
            let sheet = PackURI::new(sheet_name).unwrap();
            query_multiowner.add_part(Box::new(BlobPart::new(
                sheet.clone(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml".into(),
                b"<worksheet/>".to_vec(),
            )));
            query_multiowner
                .get_part_mut(&sheet)
                .unwrap()
                .rels_mut()
                .add_relationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
                    .into(),
                "../queryTables/queryTable1.xml".into(),
                relationship_id.into(),
                false,
            );
        }
        assert!(Snapshot::load(&query_multiowner).is_err());
    }

    #[test]
    fn custom_data_relationships_require_strict_feature_owned_part_edges() {
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        let properties = PackURI::new("/xl/customData/item1.xml").unwrap();
        let foreign = PackURI::new("/xl/foreign.bin").unwrap();

        let mut property_query = fixture();
        property_query
            .get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .retarget("rIdCustomData", "customData/item1.xml?view=1#item".into())
            .unwrap();
        assert!(Snapshot::load(&property_query).is_err());

        let mut data_query = fixture();
        data_query
            .get_part_mut(&properties)
            .unwrap()
            .rels_mut()
            .retarget("rIdCustomData", "data1.bin?view=1#payload".into())
            .unwrap();
        assert!(Snapshot::load(&data_query).is_err());

        let mut wrong_property_target = fixture();
        wrong_property_target.add_part(Box::new(BlobPart::new(
            foreign.clone(),
            "application/octet-stream".into(),
            Vec::new(),
        )));
        wrong_property_target
            .get_part_mut(&workbook)
            .unwrap()
            .rels_mut()
            .add_relationship(
                PROPERTIES_RELATIONSHIP_TYPE.into(),
                "foreign.bin".into(),
                "rIdWrongProperties".into(),
                false,
            );
        assert!(Snapshot::load(&wrong_property_target).is_err());

        let mut wrong_data_target = fixture();
        wrong_data_target
            .get_part_mut(&properties)
            .unwrap()
            .rels_mut()
            .add_relationship(
                DATA_RELATIONSHIP_TYPE.into(),
                "../workbook.xml".into(),
                "rIdWrongData".into(),
                false,
            );
        assert!(Snapshot::load(&wrong_data_target).is_err());

        let mut orphan_property_edge = fixture();
        orphan_property_edge.add_part(Box::new(BlobPart::new(
            foreign.clone(),
            "application/octet-stream".into(),
            Vec::new(),
        )));
        orphan_property_edge
            .get_part_mut(&foreign)
            .unwrap()
            .rels_mut()
            .add_relationship(
                PROPERTIES_RELATIONSHIP_TYPE.into(),
                "customData/item1.xml".into(),
                "rIdOrphanProperties".into(),
                false,
            );
        assert!(Snapshot::load(&orphan_property_edge).is_err());

        let mut orphan_data_edge = fixture();
        orphan_data_edge.add_part(Box::new(BlobPart::new(
            foreign.clone(),
            "application/octet-stream".into(),
            Vec::new(),
        )));
        orphan_data_edge
            .get_part_mut(&foreign)
            .unwrap()
            .rels_mut()
            .add_relationship(
                DATA_RELATIONSHIP_TYPE.into(),
                "customData/data1.bin".into(),
                "rIdOrphanData".into(),
                false,
            );
        assert!(Snapshot::load(&orphan_data_edge).is_err());

        let mut external_edge = fixture();
        external_edge.add_part(Box::new(BlobPart::new(
            foreign,
            "application/octet-stream".into(),
            Vec::new(),
        )));
        external_edge
            .get_part_mut(&PackURI::new("/xl/foreign.bin").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                DATA_RELATIONSHIP_TYPE.into(),
                "https://example.invalid/data.bin".into(),
                "rIdExternalData".into(),
                true,
            );
        assert!(Snapshot::load(&external_edge).is_err());
    }

    #[test]
    fn extension_list_edits_preserve_surrounding_source_and_inverse_bytes() {
        let mut package = fixture();
        let properties = PackURI::new("/xl/customData/item1.xml").unwrap();
        let original = format!(
            r#"<?preserve-root?><p:datastoreItem xmlns:p="{x14}" xmlns:s="{sml}" xmlns:v="urn:vendor" id = 'Storage-1'><!--before--><p:extLst><!--old--><s:ext uri='urn:old'><v:opaque/></s:ext></p:extLst><!--after--></p:datastoreItem>"#,
            x14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
            sml = SML,
        );
        package
            .get_part_mut(&properties)
            .unwrap()
            .set_blob(original.clone().into_bytes());
        let mut package = as_source(&package);
        let before = Snapshot::load(&package).unwrap();
        let before_bytes = package.get_part(&properties).unwrap().blob().to_vec();
        let before_extension = before.entries()[0]
            .value()
            .properties()
            .extension_list
            .clone();

        let mut no_op = Transaction::new(&mut package).unwrap();
        no_op
            .edit_properties(0, |value| {
                value.extension_list = before_extension.clone();
                Ok(())
            })
            .unwrap();
        assert!(!no_op.is_changed());
        no_op.commit().unwrap();
        assert_eq!(package.get_part(&properties).unwrap().blob(), before_bytes);

        let replacement = format!(
            r#"<p:extLst xmlns:p="{x14}" xmlns:s="{sml}" xmlns:v="urn:vendor"><s:ext uri='urn:new'><v:opaque/></s:ext></p:extLst>"#,
            x14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
            sml = SML,
        );
        let mut transaction = Transaction::new(&mut package).unwrap();
        transaction
            .edit_properties(0, |value| {
                value.id = "Storage-2".into();
                value.extension_list = Some(ExtensionList {
                    xml: replacement.as_bytes().to_vec(),
                });
                Ok(())
            })
            .unwrap();
        let commit = transaction.commit().unwrap();
        let output = package.get_part(&properties).unwrap().blob().to_vec();
        assert!(output.starts_with(b"<?preserve-root?>"));
        assert!(
            output
                .windows(b"<!--before-->".len())
                .any(|window| window == b"<!--before-->")
        );
        assert!(
            output
                .windows(b"<!--after-->".len())
                .any(|window| window == b"<!--after-->")
        );
        assert!(
            output
                .windows(b"<p:datastoreItem".len())
                .any(|window| window == b"<p:datastoreItem")
        );
        assert!(
            output
                .windows(b"urn:new".len())
                .any(|window| window == b"urn:new")
        );
        assert!(
            output
                .windows(b"Storage-2".len())
                .any(|window| window == b"Storage-2")
        );
        commit.patch().inverse().apply(&mut package).unwrap();
        assert_eq!(package.get_part(&properties).unwrap().blob(), before_bytes);
        commit.patch().apply(&mut package).unwrap();
        assert_eq!(package.get_part(&properties).unwrap().blob(), output);
    }
}
