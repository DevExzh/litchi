//! OPC graph ownership and source-bound transactions for XLSB Custom Data.
//!
//! This module deliberately keeps physical part names, relationship IDs, XML
//! source proofs, and BIFF12 spans private.  The public layer works with
//! [`StorageId`] and semantic storage views; the package facade is the only
//! place that exposes application of a validated patch.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use litchi_core::{ExecutionContext, Reservation, Resource};
use litchi_opc::{
    BlobPart, OpcPackage, OwnedContentTypes, OwnedRelationships, OwnedXmlPart, PackURI, TargetMode,
};

use super::bindings::{self, BindingError, BindingLimits, UidRewrite, WorkBudget};
use super::model::{Storage, StorageId, StorageSelectorInput};
use super::{
    CustomData, CustomDataView, DATA_CONTENT_TYPE, DATA_RELATIONSHIP_TYPE, PROPERTIES_CONTENT_TYPE,
    PROPERTIES_RELATIONSHIP_TYPE, Properties, RemovalDisposition, Result,
};
use crate::package::connections::package::{
    CONNECTIONS_CONTENT_TYPE, CONNECTIONS_RELATIONSHIP_TYPE,
};
use crate::package::error::Error;
use litchi_ooxml_common::custom_data::codec::{
    Limits as XmlLimits, canonical_extension_with_limits, parse_properties_with_limits,
    rewrite_extension_list_with_limits, rewrite_id_with_limits,
    validate_source_properties_with_limits, write_properties_with_limits,
};

const MAX_STORAGES: usize = 4_096;
const MAX_UID_UNITS: usize = 65_535;
const MAX_PROPERTIES_XML_BYTES: usize = 4 * 1024 * 1024;
const MAX_EXTENSION_XML_BYTES: usize = 2 * 1024 * 1024;
const MAX_XML_DEPTH: usize = 128;
const MAX_XML_NODES: usize = 100_000;
const MAX_PAYLOAD_BYTES: usize = 512 * 1024 * 1024;
const MAX_CONNECTIONS_BYTES: usize = 64 * 1024 * 1024;
const MAX_BIFF_RECORDS: usize = 1_048_576;
const MAX_BINDINGS: usize = 65_536;
const MAX_OPAQUE_BLOCKS: usize = 65_536;
const MAX_RELATIONSHIPS: usize = 100_000;
const MAX_RELATIONSHIP_XML_BYTES: usize = 4 * 1024 * 1024;
const MAX_PACKAGE_NODES: usize = 1_000_000;
const MAX_PACKAGE_BYTES: usize = 2 * 1024 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024 * 1024;
const MAX_TEMPORARY_BYTES: usize = 256 * 1024 * 1024;

const fn clamp_limit(value: usize, maximum: usize) -> usize {
    if value < maximum { value } else { maximum }
}

fn invalid(detail: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: "XLSB Custom Data".to_owned(),
        val: detail.into(),
    }
}

fn unsupported(detail: impl Into<String>) -> Error {
    Error::UnsupportedFeature(format!("XLSB Custom Data: {}", detail.into()))
}

fn conflict(part: impl Into<String>) -> Error {
    Error::PatchConflict { part: part.into() }
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "the shared error is non-exhaustive; typed limit and allocation variants are mapped above"
)]
fn map_common_error(error: litchi_ooxml_common::Error) -> Error {
    match error {
        litchi_ooxml_common::Error::Limit {
            resource,
            max,
            actual,
        } => Error::LimitExceeded {
            resource,
            actual,
            maximum: max,
        },
        litchi_ooxml_common::Error::Allocation { resource, source } => {
            Error::Allocation { resource, source }
        },
        other => Error::Common(other),
    }
}

/// Bounded policy for one Custom Data catalog operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_storages: usize,
    max_uid_units: usize,
    max_properties_xml_bytes: usize,
    max_extension_xml_bytes: usize,
    max_xml_depth: usize,
    max_xml_nodes: usize,
    max_payload_bytes: usize,
    max_connections_bytes: usize,
    max_biff_records: usize,
    max_bindings: usize,
    max_opaque_blocks: usize,
    max_relationships: usize,
    max_relationship_xml_bytes: usize,
    max_package_nodes: usize,
    max_package_bytes: usize,
    max_output_bytes: usize,
    max_work_bytes: usize,
    max_temporary_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_storages: MAX_STORAGES,
            max_uid_units: MAX_UID_UNITS,
            max_properties_xml_bytes: MAX_PROPERTIES_XML_BYTES,
            max_extension_xml_bytes: MAX_EXTENSION_XML_BYTES,
            max_xml_depth: MAX_XML_DEPTH,
            max_xml_nodes: MAX_XML_NODES,
            max_payload_bytes: MAX_PAYLOAD_BYTES,
            max_connections_bytes: MAX_CONNECTIONS_BYTES,
            max_biff_records: MAX_BIFF_RECORDS,
            max_bindings: MAX_BINDINGS,
            max_opaque_blocks: MAX_OPAQUE_BLOCKS,
            max_relationships: MAX_RELATIONSHIPS,
            max_relationship_xml_bytes: MAX_RELATIONSHIP_XML_BYTES,
            max_package_nodes: MAX_PACKAGE_NODES,
            max_package_bytes: MAX_PACKAGE_BYTES,
            max_output_bytes: MAX_OUTPUT_BYTES,
            max_work_bytes: MAX_WORK_BYTES,
            max_temporary_bytes: MAX_TEMPORARY_BYTES,
        }
    }
}

impl Limits {
    /// Default bounded policy as a constant for static callers.
    pub const DEFAULT: Self = Self::new();

    /// Construct the default bounded policy.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_storages: MAX_STORAGES,
            max_uid_units: MAX_UID_UNITS,
            max_properties_xml_bytes: MAX_PROPERTIES_XML_BYTES,
            max_extension_xml_bytes: MAX_EXTENSION_XML_BYTES,
            max_xml_depth: MAX_XML_DEPTH,
            max_xml_nodes: MAX_XML_NODES,
            max_payload_bytes: MAX_PAYLOAD_BYTES,
            max_connections_bytes: MAX_CONNECTIONS_BYTES,
            max_biff_records: MAX_BIFF_RECORDS,
            max_bindings: MAX_BINDINGS,
            max_opaque_blocks: MAX_OPAQUE_BLOCKS,
            max_relationships: MAX_RELATIONSHIPS,
            max_relationship_xml_bytes: MAX_RELATIONSHIP_XML_BYTES,
            max_package_nodes: MAX_PACKAGE_NODES,
            max_package_bytes: MAX_PACKAGE_BYTES,
            max_output_bytes: MAX_OUTPUT_BYTES,
            max_work_bytes: MAX_WORK_BYTES,
            max_temporary_bytes: MAX_TEMPORARY_BYTES,
        }
    }

    /// Set the storage-count ceiling.
    #[must_use]
    pub const fn with_max_storages(mut self, value: usize) -> Self {
        self.max_storages = value;
        self
    }

    /// Set the UID UTF-16 code-unit ceiling.
    #[must_use]
    pub const fn with_max_uid_units(mut self, value: usize) -> Self {
        self.max_uid_units = value;
        self
    }

    /// Alias for [`Self::with_max_uid_units`].
    #[must_use]
    pub const fn with_max_uid_characters(self, value: usize) -> Self {
        self.with_max_uid_units(value)
    }

    /// Set the properties XML byte ceiling.
    #[must_use]
    pub const fn with_max_properties_xml_bytes(mut self, value: usize) -> Self {
        self.max_properties_xml_bytes = value;
        self
    }

    /// Set the direct extension XML byte ceiling.
    #[must_use]
    pub const fn with_max_extension_xml_bytes(mut self, value: usize) -> Self {
        self.max_extension_xml_bytes = value;
        self
    }

    /// Set the XML nesting-depth ceiling.
    #[must_use]
    pub const fn with_max_xml_depth(mut self, value: usize) -> Self {
        self.max_xml_depth = value;
        self
    }

    /// Set the XML node/event ceiling.
    #[must_use]
    pub const fn with_max_xml_nodes(mut self, value: usize) -> Self {
        self.max_xml_nodes = value;
        self
    }

    /// Set the inert payload byte ceiling.
    #[must_use]
    pub const fn with_max_payload_bytes(mut self, value: usize) -> Self {
        self.max_payload_bytes = value;
        self
    }

    /// Set the connections-part byte ceiling.
    #[must_use]
    pub const fn with_max_connections_bytes(mut self, value: usize) -> Self {
        self.max_connections_bytes = value;
        self
    }

    /// Set the BIFF12 record-count ceiling.
    #[must_use]
    pub const fn with_max_biff_records(mut self, value: usize) -> Self {
        self.max_biff_records = value;
        self
    }

    /// Alias for [`Self::with_max_biff_records`].
    #[must_use]
    pub const fn with_max_records(self, value: usize) -> Self {
        self.with_max_biff_records(value)
    }

    /// Set the effective-binding ceiling.
    #[must_use]
    pub const fn with_max_bindings(mut self, value: usize) -> Self {
        self.max_bindings = value;
        self
    }

    /// Set the opaque-wrapper provenance ceiling.
    #[must_use]
    pub const fn with_max_opaque_blocks(mut self, value: usize) -> Self {
        self.max_opaque_blocks = value;
        self
    }

    /// Set the relationship-count ceiling.
    #[must_use]
    pub const fn with_max_relationships(mut self, value: usize) -> Self {
        self.max_relationships = value;
        self
    }

    /// Set the relationship XML byte ceiling.
    #[must_use]
    pub const fn with_max_relationship_xml_bytes(mut self, value: usize) -> Self {
        self.max_relationship_xml_bytes = value;
        self
    }

    /// Set the package graph node ceiling.
    #[must_use]
    pub const fn with_max_package_nodes(mut self, value: usize) -> Self {
        self.max_package_nodes = value;
        self
    }

    /// Set the aggregate package byte ceiling.
    #[must_use]
    pub const fn with_max_package_bytes(mut self, value: usize) -> Self {
        self.max_package_bytes = value;
        self
    }

    /// Set the final output byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, value: usize) -> Self {
        self.max_output_bytes = value;
        self
    }

    /// Set the bounded work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, value: usize) -> Self {
        self.max_work_bytes = value;
        self
    }

    /// Set the temporary replacement ceiling.
    #[must_use]
    pub const fn with_max_temporary_bytes(mut self, value: usize) -> Self {
        self.max_temporary_bytes = value;
        self
    }

    /// Maximum number of storage pairs.
    #[must_use]
    pub const fn max_storages(self) -> usize {
        clamp_limit(self.max_storages, MAX_STORAGES)
    }
    /// Maximum UID UTF-16 code units.
    #[must_use]
    pub const fn max_uid_units(self) -> usize {
        clamp_limit(self.max_uid_units, MAX_UID_UNITS)
    }
    /// Maximum properties XML bytes.
    #[must_use]
    pub const fn max_properties_xml_bytes(self) -> usize {
        clamp_limit(self.max_properties_xml_bytes, MAX_PROPERTIES_XML_BYTES)
    }
    /// Maximum direct extension XML bytes.
    #[must_use]
    pub const fn max_extension_xml_bytes(self) -> usize {
        clamp_limit(self.max_extension_xml_bytes, MAX_EXTENSION_XML_BYTES)
    }
    /// Maximum XML depth.
    #[must_use]
    pub const fn max_xml_depth(self) -> usize {
        clamp_limit(self.max_xml_depth, MAX_XML_DEPTH)
    }
    /// Maximum XML node/event count.
    #[must_use]
    pub const fn max_xml_nodes(self) -> usize {
        clamp_limit(self.max_xml_nodes, MAX_XML_NODES)
    }
    /// Maximum inert payload bytes.
    #[must_use]
    pub const fn max_payload_bytes(self) -> usize {
        clamp_limit(self.max_payload_bytes, MAX_PAYLOAD_BYTES)
    }
    /// Maximum connections-part bytes.
    #[must_use]
    pub const fn max_connections_bytes(self) -> usize {
        clamp_limit(self.max_connections_bytes, MAX_CONNECTIONS_BYTES)
    }
    /// Maximum BIFF12 records.
    #[must_use]
    pub const fn max_biff_records(self) -> usize {
        clamp_limit(self.max_biff_records, MAX_BIFF_RECORDS)
    }
    /// Maximum effective bindings.
    #[must_use]
    pub const fn max_bindings(self) -> usize {
        clamp_limit(self.max_bindings, MAX_BINDINGS)
    }
    /// Maximum opaque blocks.
    #[must_use]
    pub const fn max_opaque_blocks(self) -> usize {
        clamp_limit(self.max_opaque_blocks, MAX_OPAQUE_BLOCKS)
    }
    /// Maximum relationship count.
    #[must_use]
    pub const fn max_relationships(self) -> usize {
        clamp_limit(self.max_relationships, MAX_RELATIONSHIPS)
    }
    /// Maximum relationship XML bytes.
    #[must_use]
    pub const fn max_relationship_xml_bytes(self) -> usize {
        clamp_limit(self.max_relationship_xml_bytes, MAX_RELATIONSHIP_XML_BYTES)
    }
    /// Maximum package graph nodes.
    #[must_use]
    pub const fn max_package_nodes(self) -> usize {
        clamp_limit(self.max_package_nodes, MAX_PACKAGE_NODES)
    }
    /// Maximum aggregate package bytes.
    #[must_use]
    pub const fn max_package_bytes(self) -> usize {
        clamp_limit(self.max_package_bytes, MAX_PACKAGE_BYTES)
    }
    /// Maximum final output bytes.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        clamp_limit(self.max_output_bytes, MAX_OUTPUT_BYTES)
    }
    /// Maximum charged work bytes.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        clamp_limit(self.max_work_bytes, MAX_WORK_BYTES)
    }
    /// Maximum temporary replacement bytes.
    #[must_use]
    pub const fn max_temporary_bytes(self) -> usize {
        clamp_limit(self.max_temporary_bytes, MAX_TEMPORARY_BYTES)
    }

    fn binding_limits(self) -> BindingLimits {
        BindingLimits {
            max_source_bytes: self.max_connections_bytes(),
            max_record_payload_bytes: self.max_connections_bytes(),
            max_records: self.max_biff_records(),
            max_bindings: self.max_bindings(),
            max_opaque_blocks: self.max_opaque_blocks(),
            max_wrapper_depth: self.max_xml_depth().max(1),
            max_calculated_member_records: self.max_biff_records(),
            max_culture_units: 84,
            max_client_cube_urn_units: self.max_uid_units(),
            max_output_bytes: self.max_output_bytes(),
            max_work_bytes: self.max_work_bytes(),
            initial_binding_capacity: 16,
            initial_opaque_capacity: 8,
        }
    }

    fn validate(self) -> Result<Self> {
        if self.max_storages() == 0
            || self.max_package_nodes() == 0
            || self.max_package_bytes() == 0
            || self.max_output_bytes() == 0
            || self.max_work_bytes() == 0
        {
            return Err(invalid(
                "Custom Data limits that govern package publication must be non-zero",
            ));
        }
        Ok(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct PartIdentity {
    properties_part_name: PackURI,
    data_part_name: PackURI,
    workbook_relationship_id: String,
    workbook_relationship_target: String,
    data_relationship_id: String,
    data_relationship_target: String,
    source_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MemberProof {
    name: PackURI,
    content_type: String,
    bytes: Arc<Vec<u8>>,
    data_bytes: Arc<Vec<u8>>,
    properties_xml: Option<OwnedXmlPart>,
    relationships: OwnedRelationships,
    data_relationships: OwnedRelationships,
    inbound: Vec<InboundProof>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct InboundProof {
    owner: PackURI,
    relationship_id: String,
    relationship_type: String,
    target: String,
    external: bool,
}

/// Physical identity retained privately by one source-bound storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum StorageOrigin {
    Existing(PartIdentity),
    New(PartIdentity),
}

impl StorageOrigin {
    fn identity(&self) -> &PartIdentity {
        match self {
            Self::Existing(identity) | Self::New(identity) => identity,
        }
    }

    fn source_id(&self) -> Option<&str> {
        match self {
            Self::Existing(identity) => {
                // The source UID is kept in the properties relationship target
                // only indirectly, so this is filled by the package loader via
                // the identity's private side table.  New entries deliberately
                // have no source reference.
                identity.source_id.as_deref()
            },
            Self::New(_) => None,
        }
    }
}

impl PartIdentity {
    /// Compare the OPC graph identity while ignoring the semantic UID that
    /// was decoded from the Properties XML. A rename changes that UID but
    /// keeps both part names and relationship edges physically stable.
    fn same_physical(&self, other: &Self) -> bool {
        self.properties_part_name == other.properties_part_name
            && self.data_part_name == other.data_part_name
            && self.workbook_relationship_id == other.workbook_relationship_id
            && self.workbook_relationship_target == other.workbook_relationship_target
            && self.data_relationship_id == other.data_relationship_id
            && self.data_relationship_target == other.data_relationship_target
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ConnectionState {
    part_name: PackURI,
    content_type: String,
    bytes: Arc<Vec<u8>>,
    relationships: OwnedRelationships,
    inbound: Vec<InboundProof>,
    workbook_relationship_id: String,
    workbook_relationship_target: String,
    references: Vec<String>,
    blocking_opaque: bool,
    record_count: usize,
}

#[derive(Debug, Clone)]
struct SourceState {
    content_types: OwnedContentTypes,
    root_relationships: OwnedRelationships,
    workbook_name: PackURI,
    workbook_bytes: Arc<Vec<u8>>,
    workbook_relationships: OwnedRelationships,
    members: Vec<MemberProof>,
    connection: Option<ConnectionState>,
    signed: bool,
    signature_policy_required: bool,
}

impl PartialEq for SourceState {
    fn eq(&self, other: &Self) -> bool {
        self.content_types == other.content_types
            && self.root_relationships == other.root_relationships
            && self.workbook_name == other.workbook_name
            && self.workbook_bytes == other.workbook_bytes
            && self.workbook_relationships == other.workbook_relationships
            && self.members == other.members
            && self.connection == other.connection
            && self.signed == other.signed
            && self.signature_policy_required == other.signature_policy_required
    }
}

impl Eq for SourceState {}

/// An immutable, source-bound Custom Data catalog.
#[derive(Debug, Clone)]
pub struct Snapshot {
    entries: Arc<[Storage]>,
    source: Arc<SourceState>,
    limits: Limits,
    context: Option<ExecutionContext>,
    /// The source package and decoded views are retained for the lifetime of
    /// the snapshot. Keep their memory charge scoped to that same lifetime;
    /// charging it cumulatively would make repeated source checks exhaust a
    /// caller's budget even though the intermediate snapshots are dropped.
    #[allow(
        dead_code,
        reason = "the Arc token releases the retained snapshot charge on drop"
    )]
    memory_reservation: Option<Arc<Reservation>>,
}

impl Snapshot {
    /// Read a catalog with the default finite policy.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        Self::load_with_limits(package, Limits::default())
    }

    /// Read a catalog with an explicit finite policy.
    pub fn load_with_limits(package: &OpcPackage, limits: Limits) -> Result<Self> {
        Self::load_with_limits_and_context(package, limits, None)
    }

    /// Read a catalog under an owned execution context.
    pub fn load_with_limits_and_context(
        package: &OpcPackage,
        limits: Limits,
        context: impl Into<Option<ExecutionContext>>,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        let context = context.into();
        check_context(context.as_ref())?;
        let (source, mut entries, memory_reservation) =
            capture_source(package, limits, context.as_ref())?;
        entries.sort_unstable_by(|left, right| left.id().cmp(right.id()));
        if entries.len() > limits.max_storages() {
            return Err(Error::LimitExceeded {
                resource: "Custom Data storages",
                actual: entries.len(),
                maximum: limits.max_storages(),
            });
        }
        check_context(context.as_ref())?;
        Ok(Self {
            entries: Arc::from(entries.into_boxed_slice()),
            source: Arc::new(source),
            limits,
            context,
            memory_reservation,
        })
    }

    /// Borrow all semantic storages in deterministic UID order.
    #[must_use]
    pub fn storages(&self) -> &[Storage] {
        &self.entries
    }

    /// Alias for [`Self::storages`].
    #[must_use]
    pub fn entries(&self) -> &[Storage] {
        self.storages()
    }

    /// Find one storage by exact decoded UID.
    #[must_use]
    pub fn find(&self, id: &str) -> Option<&Storage> {
        self.entries.iter().find(|storage| storage.id() == id)
    }

    /// Number of storages in this catalog.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether this catalog is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Retained resource policy.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    pub(crate) fn execution_context(&self) -> Option<ExecutionContext> {
        self.context.clone()
    }

    /// Count effective ExtConn14 references to one UID.
    #[must_use]
    pub fn connection_references(&self, id: &str) -> usize {
        if id.is_empty() {
            return 0;
        }
        self.source.connection.as_ref().map_or(0, |connection| {
            connection
                .references
                .iter()
                .filter(|value| !value.is_empty() && value.as_str() == id)
                .count()
        })
    }

    /// Whether a hidden AC/FRT/unsupported reference candidate is retained.
    #[must_use]
    pub fn has_opaque_reference_candidates(&self) -> bool {
        self.source
            .connection
            .as_ref()
            .is_some_and(|connection| connection.blocking_opaque)
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.limits == other.limits && self.source.as_ref() == other.source.as_ref()
    }

    fn same_semantics(&self, entries: &[Storage]) -> bool {
        semantic_entries_equal(&self.entries, entries)
    }

    fn connection_values(&self) -> &[String] {
        self.source
            .connection
            .as_ref()
            .map_or(&[], |value| value.references.as_slice())
    }
}

/// A detached transaction over an immutable package snapshot.
pub struct Transaction {
    before: Snapshot,
    draft: Vec<Storage>,
    package: OpcPackage,
    limits: Limits,
    context: Option<ExecutionContext>,
    connection_edits: BTreeMap<String, String>,
}

impl std::fmt::Debug for Transaction {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Transaction")
            .field("storages", &self.draft.len())
            .field("changed_connections", &self.connection_edits.len())
            .field("limits", &self.limits)
            .finish_non_exhaustive()
    }
}

impl Transaction {
    /// Start a detached transaction from a package clone.
    pub(crate) fn from_package(package: OpcPackage) -> Result<Self> {
        Self::from_package_with_limits_and_context(package, Limits::default(), None)
    }

    /// Start a detached transaction using explicit limits.
    pub(crate) fn from_package_with_limits(package: OpcPackage, limits: Limits) -> Result<Self> {
        Self::from_package_with_limits_and_context(package, limits, None)
    }

    /// Start a detached transaction under an owned execution context.
    pub(crate) fn from_package_with_limits_and_context(
        package: OpcPackage,
        limits: Limits,
        context: Option<ExecutionContext>,
    ) -> Result<Self> {
        let before = Snapshot::load_with_limits_and_context(&package, limits, context.clone())?;
        Ok(Self {
            draft: clone_storages(&before.entries, "Custom Data transaction draft")?,
            package,
            before,
            limits,
            context,
            connection_edits: BTreeMap::new(),
        })
    }

    /// Source snapshot captured when this transaction was created.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow currently staged semantic storages.
    #[must_use]
    pub fn storages(&self) -> &[Storage] {
        &self.draft
    }

    /// Alias for [`Self::storages`].
    #[must_use]
    pub fn entries(&self) -> &[Storage] {
        self.storages()
    }

    /// Replace one complete storage value while retaining its package graph.
    pub fn set<S: StorageSelectorInput>(&mut self, selector: S, value: CustomData) -> Result<bool> {
        let index = self.resolve(selector)?;
        self.set_at(index, value.into())
    }

    /// Replace only the inert payload of one storage.
    pub fn set_data<S: StorageSelectorInput>(
        &mut self,
        selector: S,
        data: Vec<u8>,
    ) -> Result<bool> {
        let index = self.resolve(selector)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| invalid("Custom Data storage selector is absent"))?;
        let mut value = current.value.clone();
        value.data = Arc::new(data);
        self.set_at(index, value)
    }

    /// Edit typed X14 properties without exposing source XML.
    pub fn edit_properties<S: StorageSelectorInput>(
        &mut self,
        selector: S,
        edit: impl FnOnce(&mut Properties) -> Result<()>,
    ) -> Result<bool> {
        let index = self.resolve(selector)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| invalid("Custom Data storage selector is absent"))?;
        let mut properties = current.value.properties.as_ref().clone();
        edit(&mut properties)?;
        let mut value = current.value.clone();
        value.properties = Arc::new(properties);
        self.set_at(index, value)
    }

    /// Insert a new inert storage and return its owned UID.
    pub fn insert(&mut self, value: CustomData) -> Result<StorageId> {
        check_context(self.context.as_ref())?;
        let value = value.into();
        validate_value(&value, &self.limits)?;
        let id = StorageId::new(value.properties.id.clone())?;
        if self.draft.iter().any(|storage| storage.id() == id.as_str()) {
            return Err(invalid("Custom Data storage IDs must be unique"));
        }
        if self.draft.len() >= self.limits.max_storages() {
            return Err(Error::LimitExceeded {
                resource: "Custom Data storages",
                actual: self.draft.len() + 1,
                maximum: self.limits.max_storages(),
            });
        }
        let identity = allocate_identity(&self.package, &self.draft)?;
        let mut candidate = clone_storages(&self.draft, "Custom Data insertion draft")?;
        candidate
            .try_reserve(1)
            .map_err(|source| allocation("Custom Data insertion draft", source))?;
        candidate.push(Storage {
            value,
            origin: Some(StorageOrigin::New(identity)),
        });
        validate_staged_entries(&candidate, &self.limits)?;
        check_context(self.context.as_ref())?;
        self.draft = candidate;
        Ok(id)
    }

    /// Insert or replace a storage with the same UID.
    pub fn upsert(&mut self, value: CustomData) -> Result<StorageId> {
        let id = StorageId::new(value.properties.id.clone())?;
        if let Some(index) = self
            .draft
            .iter()
            .position(|storage| storage.id() == id.as_str())
        {
            self.set_at(index, value.into())?;
            return Ok(id);
        }
        self.insert(value)
    }

    /// Rename one storage and all effective ExtConn14 references to it.
    pub fn rename<S: StorageSelectorInput>(
        &mut self,
        selector: S,
        id: impl Into<String>,
    ) -> Result<bool> {
        let id = id.into();
        self.edit_properties(selector, |properties| {
            properties.id = id;
            Ok(())
        })
    }

    /// Remove one storage, refusing if effective bindings refer to it.
    pub fn remove<S: StorageSelectorInput>(&mut self, selector: S) -> Result<Option<Storage>> {
        self.remove_with(selector, RemovalDisposition::RejectReferenced)
    }

    /// Remove one storage after an explicit detach or retarget decision.
    pub fn remove_with<S: StorageSelectorInput>(
        &mut self,
        selector: S,
        disposition: RemovalDisposition,
    ) -> Result<Option<Storage>> {
        let Some(index) = self.resolve_optional(selector) else {
            return Ok(None);
        };
        let removed = self.draft[index].clone();
        let source_id = removed.origin.as_ref().and_then(StorageOrigin::source_id);
        let reference_count = source_id.map_or(0, |id| self.before.connection_references(id));
        let target = match disposition {
            RemovalDisposition::RejectReferenced if reference_count != 0 => {
                return Err(Error::CustomDataReferenced {
                    id: removed.id().to_owned(),
                    connections: reference_count,
                });
            },
            RemovalDisposition::RejectReferenced | RemovalDisposition::DetachConnections => {
                String::new()
            },
            RemovalDisposition::RetargetConnections(id) => {
                if id.is_empty()
                    || !self
                        .draft
                        .iter()
                        .enumerate()
                        .any(|(candidate, storage)| candidate != index && storage.id() == id)
                {
                    return Err(invalid(
                        "connection retarget UID must name a remaining non-empty Custom Data storage",
                    ));
                }
                id
            },
        };
        let mut candidate_edits = self.connection_edits.clone();
        if let Some(source_id) = source_id {
            if self.before.source.connection.is_some()
                && (reference_count != 0 || !target.is_empty())
            {
                candidate_edits.insert(source_id.to_owned(), target);
            }
        }
        check_context(self.context.as_ref())?;
        let mut candidate = clone_storages(&self.draft, "Custom Data removal draft")?;
        candidate.remove(index);
        validate_staged_entries(&candidate, &self.limits)?;
        check_context(self.context.as_ref())?;
        self.draft = candidate;
        self.connection_edits = candidate_edits;
        Ok(Some(removed))
    }

    /// Whether the staged semantic graph differs from the source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !staged_entries_equal(&self.before.entries, &self.draft)
            || !self.connection_edits.is_empty()
    }

    /// Alias for [`Self::is_changed`].
    #[must_use]
    pub fn changed(&self) -> bool {
        self.is_changed()
    }

    /// Validate and publish this transaction atomically.
    pub fn commit(self) -> Result<Commit> {
        check_context(self.context.as_ref())?;
        if !self.is_changed() {
            let patch = Patch {
                before: self.before.clone(),
                after: self.before.clone(),
            };
            return Ok(Commit {
                snapshot: self.before.clone(),
                patch,
                changed: false,
            });
        }
        if self.before.source.signed || self.before.source.signature_policy_required {
            return Err(Error::Signed);
        }
        let current = Snapshot::load_with_limits_and_context(
            &self.package,
            self.limits,
            self.context.clone(),
        )?;
        if !current.same_source(&self.before) {
            return Err(conflict("Custom Data source closure"));
        }
        if self.has_identity_blocking_opaque() {
            return Err(unsupported(
                "an opaque AC/FRT/unsupported connection block may hide a Custom Data reference",
            ));
        }
        let mut candidate = self.package.clone();
        apply_draft(
            &mut candidate,
            &self.before,
            &self.draft,
            &self.connection_edits,
            self.limits,
            self.context.as_ref(),
        )?;
        check_context(self.context.as_ref())?;
        let snapshot =
            Snapshot::load_with_limits_and_context(&candidate, self.limits, self.context.clone())?;
        if !snapshot.same_semantics(&self.draft)
            || !connection_values_equal_after(&self.before, &snapshot, &self.connection_edits)
        {
            return Err(invalid("Custom Data publication changed staged semantics"));
        }
        let patch = Patch {
            before: self.before,
            after: snapshot.clone(),
        };
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }

    fn resolve<S: StorageSelectorInput>(&self, selector: S) -> Result<usize> {
        self.resolve_optional(selector)
            .ok_or_else(|| invalid("Custom Data storage selector is absent"))
    }

    fn resolve_optional<S: StorageSelectorInput>(&self, selector: S) -> Option<usize> {
        selector.resolve_storage(&self.draft)
    }

    fn set_at(&mut self, index: usize, value: CustomDataView) -> Result<bool> {
        check_context(self.context.as_ref())?;
        validate_value(&value, &self.limits)?;
        let current = self
            .draft
            .get(index)
            .ok_or_else(|| invalid("Custom Data storage selector is absent"))?
            .clone();
        if current.value == value {
            return Ok(false);
        }
        if self
            .draft
            .iter()
            .enumerate()
            .any(|(candidate, storage)| candidate != index && storage.id() == value.properties.id)
        {
            return Err(invalid("Custom Data storage IDs must be unique"));
        }
        if value.properties.id.is_empty()
            && current
                .origin
                .as_ref()
                .and_then(StorageOrigin::source_id)
                .is_some_and(|id| self.before.connection_references(id) != 0)
        {
            return Err(invalid(
                "a referenced Custom Data storage cannot be renamed to the empty UID",
            ));
        }
        let mut candidate_edits = self.connection_edits.clone();
        if let Some(source_id) = current.origin.as_ref().and_then(StorageOrigin::source_id) {
            let referenced = self
                .before
                .source
                .connection
                .as_ref()
                .is_some_and(|connection| {
                    !source_id.is_empty()
                        && connection
                            .references
                            .iter()
                            .any(|value| !value.is_empty() && value == source_id)
                });
            if source_id != value.properties.id && referenced {
                candidate_edits.insert(source_id.to_owned(), value.properties.id.clone());
            } else if source_id == value.properties.id || !referenced {
                candidate_edits.remove(source_id);
            }
        }
        check_context(self.context.as_ref())?;
        let mut candidate = clone_storages(&self.draft, "Custom Data update draft")?;
        candidate[index].value = value;
        validate_staged_entries(&candidate, &self.limits)?;
        check_context(self.context.as_ref())?;
        self.draft = candidate;
        self.connection_edits = candidate_edits;
        Ok(true)
    }

    fn has_identity_blocking_opaque(&self) -> bool {
        if !self.before.has_opaque_reference_candidates() {
            return false;
        }
        self.connection_edits.iter().any(|(from, to)| from != to)
            || self.before.entries.iter().any(|before| {
                let Some(current) = self
                    .draft
                    .iter()
                    .find(|candidate| same_identity(before, candidate))
                else {
                    return before
                        .origin
                        .as_ref()
                        .is_some_and(|origin| matches!(origin, StorageOrigin::Existing(_)));
                };
                before.id() != current.id()
            })
    }
}

/// A source-checked, reversible package patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Source proof required before applying this patch.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Semantic/source state produced by this patch.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact source no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this patch to an OPC package after source-closure validation.
    pub(crate) fn apply_to_opc(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load_with_limits_and_context(
            target,
            self.before.limits,
            self.before.context.clone(),
        )?;
        if !current.same_source(&self.before) {
            return Err(conflict("Custom Data source closure"));
        }
        if self.is_empty() {
            return Ok(());
        }
        if target.is_signed() || target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        // Reapply only the Custom Data graph delta. The target is allowed to
        // carry edits to unrelated worksheet or opaque parts after the patch
        // was created; replacing the whole retained package would silently
        // discard those edits.
        let connection_edits = connection_edits_between(&self.before, &self.after)?;
        apply_draft(
            target,
            &self.before,
            &self.after.entries,
            &connection_edits,
            self.before.limits,
            self.before.context.as_ref(),
        )?;
        restore_content_types_template_if_scoped(
            target,
            &self.before,
            &self.after,
            self.before.limits,
        )?;
        restore_relationship_templates_if_scoped(
            target,
            &self.before,
            &self.after,
            self.before.limits,
        )?;
        restore_member_templates_if_scoped(target, &self.after, self.before.limits)?;
        let applied = Snapshot::load_with_limits_and_context(
            target,
            self.before.limits,
            self.before.context.clone(),
        )?;
        if !applied.same_source(&self.after)
            || !applied.same_semantics(&self.after.entries)
            || applied.connection_values() != self.after.connection_values()
        {
            return Err(invalid(
                "Custom Data patch postcondition changed the source closure",
            ));
        }
        Ok(())
    }
}

/// Successful publication result.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether any semantic or connection bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting source-bound snapshot.
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

impl Clone for Commit {
    fn clone(&self) -> Self {
        Self {
            snapshot: self.snapshot.clone(),
            patch: self.patch.clone(),
            changed: self.changed,
        }
    }
}

fn check_context(context: Option<&ExecutionContext>) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    context.check().map_err(Error::from)?;
    // A zero-sized charge still checks cancellation and keeps this owner on
    // the same execution-policy seam as the other package services.
    context.consume(Resource::Work, 0).map_err(Error::from)?;
    Ok(())
}

fn charge_context(
    context: Option<&ExecutionContext>,
    resource: Resource,
    amount: usize,
) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    let amount = u64::try_from(amount)
        .map_err(|error| invalid(format!("execution charge exceeds u64: {error}")))?;
    context.consume(resource, amount).map_err(Error::from)?;
    Ok(())
}

/// Reserve memory for allocations retained by a snapshot or for a temporary
/// rewrite. Unlike work and object accounting, this charge is released when
/// the returned token is dropped.
fn reserve_context(
    context: Option<&ExecutionContext>,
    resource: Resource,
    amount: usize,
) -> Result<Option<Arc<Reservation>>> {
    let Some(context) = context else {
        return Ok(None);
    };
    let amount = u64::try_from(amount)
        .map_err(|error| invalid(format!("execution charge exceeds u64: {error}")))?;
    context
        .reserve(resource, amount)
        .map(|reservation| Some(Arc::new(reservation)))
        .map_err(Error::from)
}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "BindingError is non-exhaustive; the owner maps all limit and security refusals explicitly"
)]
fn map_binding_error(error: BindingError) -> Error {
    match error {
        BindingError::SourceLimit { actual, limit }
        | BindingError::RecordPayloadLimit { actual, limit }
        | BindingError::RecordLimit { actual, limit }
        | BindingError::BindingLimit { actual, limit }
        | BindingError::OpaqueLimit { actual, limit }
        | BindingError::CalculatedMemberLimit { actual, limit }
        | BindingError::OutputLimit { actual, limit }
        | BindingError::WorkLimit { actual, limit }
        | BindingError::StringLimit {
            units: actual,
            limit,
            ..
        }
        | BindingError::UidLimit {
            units: actual,
            limit,
        } => Error::LimitExceeded {
            resource: "XLSB Custom Data BIFF12 operation",
            actual,
            maximum: limit,
        },
        BindingError::HiddenReference => {
            unsupported("an opaque AC/FRT/ExtConn15 block may hide a Custom Data reference")
        },
        BindingError::SourceMismatch => conflict("connections.bin source"),
        BindingError::BudgetRejected => Error::Execution(litchi_core::ExecutionError::Cancelled),
        other => invalid(other.to_string()),
    }
}

fn map_binding_error_with_budget(error: BindingError, budget: &mut ContextWorkBudget<'_>) -> Error {
    if matches!(error, BindingError::BudgetRejected) {
        if let Some(error) = budget.rejected.take() {
            return Error::Execution(error);
        }
    }
    map_binding_error(error)
}

fn validate_uid_limits(id: &str, limits: Limits) -> Result<()> {
    let units = id.encode_utf16().count();
    if units > limits.max_uid_units() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data UID UTF-16 units",
            actual: units,
            maximum: limits.max_uid_units(),
        });
    }
    Ok(())
}

fn validate_value(value: &CustomDataView, limits: &Limits) -> Result<()> {
    if value.data.len() > limits.max_payload_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data payload bytes",
            actual: value.data.len(),
            maximum: limits.max_payload_bytes(),
        });
    }
    validate_uid_limits(&value.properties.id, *limits)?;
    if let Some(extension) = value.properties.extension_list.as_ref() {
        if extension.xml.len() > limits.max_extension_xml_bytes() {
            return Err(Error::LimitExceeded {
                resource: "Custom Data extension XML bytes",
                actual: extension.xml.len(),
                maximum: limits.max_extension_xml_bytes(),
            });
        }
    }
    let xml_limits = codec_limits(*limits);
    validate_source_properties_with_limits(&value.properties, &xml_limits)
        .map_err(map_common_error)?;
    Ok(())
}

fn codec_limits(limits: Limits) -> XmlLimits {
    codec_limits_with_output(limits, limits.max_properties_xml_bytes())
}

fn codec_limits_with_output(limits: Limits, output_bytes: usize) -> XmlLimits {
    // The owner validates the required UID in UTF-16 units immediately after
    // decoding. Keep the shared XML string ceiling independent of that owner
    // policy so a small UID limit reports the typed XLSB UID refusal instead
    // of an earlier byte-based common-codec error.
    let string_bytes = XmlLimits::MAX_STRING_BYTES;
    XmlLimits::standard()
        .with_properties_xml_bytes(limits.max_properties_xml_bytes().min(output_bytes))
        .with_extension_xml_bytes(limits.max_extension_xml_bytes().min(output_bytes))
        .with_string_bytes(string_bytes)
        .with_nodes(limits.max_xml_nodes())
        .with_depth(limits.max_xml_depth())
        .with_namespace_bytes(limits.max_properties_xml_bytes())
        .with_attributes(limits.max_xml_nodes())
}

fn validate_staged_entries(entries: &[Storage], limits: &Limits) -> Result<()> {
    if entries.len() > limits.max_storages() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data storages",
            actual: entries.len(),
            maximum: limits.max_storages(),
        });
    }
    let mut ids = HashSet::new();
    ids.try_reserve(entries.len())
        .map_err(|source| allocation("Custom Data IDs", source))?;
    for entry in entries {
        validate_value(&entry.value, limits)?;
        if !ids.insert(entry.id().to_owned()) {
            return Err(invalid("Custom Data storage IDs must be unique"));
        }
    }
    Ok(())
}

fn clone_storages(entries: &[Storage], resource: &'static str) -> Result<Vec<Storage>> {
    let mut cloned = Vec::new();
    cloned
        .try_reserve(entries.len())
        .map_err(|source| allocation(resource, source))?;
    cloned.extend(entries.iter().cloned());
    Ok(cloned)
}

fn semantic_entries_equal(left: &[Storage], right: &[Storage]) -> bool {
    // `left` is always a freshly loaded Snapshot and is sorted by UID. The
    // transaction draft can retain its original physical order after a
    // rename crosses another UID, so compare through the validated unique UID
    // index rather than by position or an O(storages²) search.
    left.len() == right.len()
        && right.iter().all(|candidate| {
            let Ok(index) = left.binary_search_by(|other| other.id().cmp(candidate.id())) else {
                return false;
            };
            let other = &left[index];
            same_identity(candidate, other)
                && candidate.value.data == other.value.data
                && candidate.value.properties == other.value.properties
        })
}

fn staged_entries_equal(left: &[Storage], right: &[Storage]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left, right)| {
            left.id() == right.id()
                && left.value.data == right.value.data
                && left.value.properties == right.value.properties
                && same_identity(left, right)
        })
}

fn connection_values_equal_after(
    before: &Snapshot,
    after: &Snapshot,
    edits: &BTreeMap<String, String>,
) -> bool {
    let before_values = before.connection_values();
    let after_values = after.connection_values();
    before_values.len() == after_values.len()
        && before_values
            .iter()
            .zip(after_values)
            .all(|(before, after)| {
                edits
                    .get(before)
                    .map_or(before == after, |value| value == after)
            })
}

fn connection_edits_between(
    before: &Snapshot,
    after: &Snapshot,
) -> Result<BTreeMap<String, String>> {
    let mut edits = BTreeMap::new();
    let (Some(before_connection), Some(after_connection)) = (
        before.source.connection.as_ref(),
        after.source.connection.as_ref(),
    ) else {
        if before.source.connection.is_some() != after.source.connection.is_some() {
            return Err(conflict("connections.bin source presence"));
        }
        return Ok(edits);
    };
    if before_connection.references.len() != after_connection.references.len() {
        return Err(conflict("connections.bin binding cardinality"));
    }
    for (from, to) in before_connection
        .references
        .iter()
        .zip(&after_connection.references)
    {
        if from == to {
            continue;
        }
        if let Some(previous) = edits.insert(from.clone(), to.clone())
            && previous != *to
        {
            return Err(conflict("connections.bin UID rewrite is ambiguous"));
        }
    }
    // Ensure every changed value is represented by the map, including a
    // detach to the empty UID and many-to-one retargets.
    for (from, to) in before_connection
        .references
        .iter()
        .zip(&after_connection.references)
    {
        if edits.get(from).map(String::as_str) != Some(to.as_str()) && from != to {
            return Err(conflict("connections.bin UID rewrite is incomplete"));
        }
    }
    Ok(edits)
}

fn package_size(package: &OpcPackage) -> Result<usize> {
    let mut total = package.source_content_types()?.bytes().len();
    total = total
        .checked_add(
            package
                .source_relationships(&PackURI::new("/").map_err(invalid)?)?
                .bytes()
                .len(),
        )
        .ok_or_else(|| invalid("Custom Data package relationship byte count overflow"))?;
    for part in package.iter_parts() {
        total = total
            .checked_add(part.blob().len())
            .ok_or_else(|| invalid("Custom Data package byte count overflow"))?;
        total = total
            .checked_add(package.source_relationships(part.partname())?.bytes().len())
            .ok_or_else(|| invalid("Custom Data relationship byte count overflow"))?;
    }
    Ok(total)
}

fn validate_package_limits(package: &OpcPackage, limits: Limits) -> Result<()> {
    if package.part_count() > limits.max_package_nodes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data package nodes",
            actual: package.part_count(),
            maximum: limits.max_package_nodes(),
        });
    }
    let bytes = package_size(package)?;
    if bytes > limits.max_package_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data package bytes",
            actual: bytes,
            maximum: limits.max_package_bytes(),
        });
    }
    let relationships = package_relationship_count(package)?;
    if relationships > limits.max_relationships() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data relationships",
            actual: relationships,
            maximum: limits.max_relationships(),
        });
    }
    Ok(())
}

fn validate_output_limits(package: &OpcPackage, limits: Limits) -> Result<()> {
    let bytes = package_size(package)?;
    if bytes > limits.max_output_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data final output bytes",
            actual: bytes,
            maximum: limits.max_output_bytes(),
        });
    }
    Ok(())
}

fn package_relationship_count(package: &OpcPackage) -> Result<usize> {
    let mut count = package.rels().iter().count();
    for part in package.iter_parts() {
        count = count
            .checked_add(part.rels().iter().count())
            .ok_or_else(|| invalid("Custom Data relationship count overflow"))?;
    }
    Ok(count)
}

fn capture_source(
    package: &OpcPackage,
    limits: Limits,
    context: Option<&ExecutionContext>,
) -> Result<(SourceState, Vec<Storage>, Option<Arc<Reservation>>)> {
    check_context(context)?;
    validate_package_limits(package, limits)?;
    let memory_reservation = reserve_context(context, Resource::Memory, package_size(package)?)?;
    let workbook = package.main_document_part()?;
    if workbook.content_type() != litchi_opc::constants::content_type::XLSB_BIN {
        return Err(invalid("Custom Data owner requires the XLSB workbook part"));
    }
    let workbook_name = workbook.partname().clone();
    let workbook_bytes = workbook.blob_arc();
    let content_types = package.source_content_types()?;
    if content_types.bytes().len() > limits.max_relationship_xml_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data content-types XML bytes",
            actual: content_types.bytes().len(),
            maximum: limits.max_relationship_xml_bytes(),
        });
    }
    let root_relationships = package.source_relationships(&PackURI::new("/").map_err(invalid)?)?;
    if root_relationships.bytes().len() > limits.max_relationship_xml_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data package relationships XML bytes",
            actual: root_relationships.bytes().len(),
            maximum: limits.max_relationship_xml_bytes(),
        });
    }
    let workbook_relationships = package.source_relationships(&workbook_name)?;
    if workbook_relationships.bytes().len() > limits.max_relationship_xml_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data workbook relationships XML bytes",
            actual: workbook_relationships.bytes().len(),
            maximum: limits.max_relationship_xml_bytes(),
        });
    }

    let mut source = SourceState {
        content_types,
        root_relationships,
        workbook_name,
        workbook_bytes,
        workbook_relationships,
        members: Vec::new(),
        connection: None,
        signed: package.is_signed(),
        signature_policy_required: package.requires_signature_edit_policy(),
    };
    // The graph loader also validates all relationship cardinalities.  It is
    // intentionally called before the connection closure so dangling or
    // duplicate storage owners fail before any BIFF source is inspected.
    let entries = load_entries(package, &source, limits, context)?;
    source.members = capture_member_proofs(package, &entries, limits, context)?;
    source.connection = capture_connection(package, &source.workbook_name, limits, context)?;
    validate_connection_ids(&entries, source.connection.as_ref())?;
    check_context(context)?;
    Ok((source, entries, memory_reservation))
}

fn load_entries(
    package: &OpcPackage,
    _source: &SourceState,
    limits: Limits,
    context: Option<&ExecutionContext>,
) -> Result<Vec<Storage>> {
    check_context(context)?;
    let workbook = package.main_document_part()?;
    let workbook_name = workbook.partname();
    let mut properties_parts = Vec::new();
    properties_parts
        .try_reserve(limits.max_storages().min(package.part_count()))
        .map_err(|source| allocation("Custom Data Properties parts", source))?;
    for part in package.iter_parts() {
        check_context(context)?;
        if part.content_type() == PROPERTIES_CONTENT_TYPE {
            if properties_parts.len() >= limits.max_storages() {
                return Err(Error::LimitExceeded {
                    resource: "Custom Data storages",
                    actual: properties_parts.len() + 1,
                    maximum: limits.max_storages(),
                });
            }
            properties_parts.push(part);
        }
    }

    // Validate every relationship using one profile table before selecting any
    // owner. This catches external, query/fragment, wrong-source, and dangling
    // edges that could otherwise leave an orphaned semantic projection.
    validate_feature_relationships(package, workbook_name)?;

    let mut ids = HashSet::new();
    ids.try_reserve(properties_parts.len())
        .map_err(|source| allocation("Custom Data IDs", source))?;
    let mut data_targets = Vec::new();
    data_targets
        .try_reserve(properties_parts.len())
        .map_err(|source| allocation("Custom Data payload targets", source))?;
    let mut entries = Vec::new();
    entries
        .try_reserve(properties_parts.len())
        .map_err(|source| allocation("Custom Data storage entries", source))?;

    for properties_part in properties_parts {
        check_context(context)?;
        if properties_part.blob().len() > limits.max_properties_xml_bytes() {
            return Err(Error::LimitExceeded {
                resource: "Custom Data Properties XML bytes",
                actual: properties_part.blob().len(),
                maximum: limits.max_properties_xml_bytes(),
            });
        }
        let mut owner = None;
        for relationship in workbook.rels().iter() {
            if relationship.reltype() != PROPERTIES_RELATIONSHIP_TYPE
                || relationship.is_external()
                || relationship.target_query().is_some()
                || relationship.target_fragment().is_some()
                || !relationship
                    .target_partname()
                    .ok()
                    .is_some_and(|target| target.is_equivalent_to(properties_part.partname()))
            {
                continue;
            }
            if owner.is_some() {
                return Err(invalid(format!(
                    "Custom Data Properties part '{}' must have exactly one workbook owner",
                    properties_part.partname()
                )));
            }
            owner = Some(relationship);
        }
        let Some(owner) = owner else {
            return Err(invalid(format!(
                "Custom Data Properties part '{}' must have exactly one workbook owner",
                properties_part.partname()
            )));
        };
        let relationships = properties_part.rels();
        let mut data_relationship = None;
        for relationship in relationships.iter() {
            if relationship.reltype() != DATA_RELATIONSHIP_TYPE {
                return Err(invalid(
                    "Custom Data Properties has an unexpected relationship",
                ));
            }
            if data_relationship.is_some() {
                return Err(invalid(format!(
                    "Custom Data Properties part '{}' must have one internal data relationship",
                    properties_part.partname()
                )));
            }
            data_relationship = Some(relationship);
        }
        let Some(data_relationship) = data_relationship else {
            return Err(invalid(format!(
                "Custom Data Properties part '{}' must have one internal data relationship",
                properties_part.partname()
            )));
        };
        if data_relationship.is_external() {
            return Err(invalid(format!(
                "Custom Data Properties part '{}' must have one internal data relationship",
                properties_part.partname()
            )));
        }
        if data_relationship.target_query().is_some()
            || data_relationship.target_fragment().is_some()
        {
            return Err(invalid(
                "Custom Data payload relationship cannot have a query or fragment",
            ));
        }
        let data_part_name = data_relationship.target_partname()?;
        let data_part = package.get_part(&data_part_name)?;
        if data_part.content_type() != DATA_CONTENT_TYPE {
            return Err(Error::InvalidContentType {
                expected: DATA_CONTENT_TYPE.to_owned(),
                got: data_part.content_type().to_owned(),
            });
        }
        if data_part.blob().len() > limits.max_payload_bytes() {
            return Err(Error::LimitExceeded {
                resource: "Custom Data payload bytes",
                actual: data_part.blob().len(),
                maximum: limits.max_payload_bytes(),
            });
        }
        if !data_part.rels().is_empty() {
            return Err(invalid("Custom Data payload has outbound relationships"));
        }
        let properties =
            parse_properties_with_limits(properties_part.blob(), &codec_limits(limits))
                .map_err(map_common_error)?;
        validate_uid_limits(&properties.id, limits)?;
        if !ids.insert(properties.id.clone()) {
            return Err(invalid("Custom Data storage IDs must be unique"));
        }
        if data_targets
            .iter()
            .any(|target: &PackURI| target.is_equivalent_to(data_part.partname()))
        {
            return Err(invalid(
                "a Custom Data payload cannot be shared by multiple properties parts",
            ));
        }
        data_targets.push(data_part.partname().clone());
        let identity = PartIdentity {
            properties_part_name: properties_part.partname().clone(),
            data_part_name: data_part.partname().clone(),
            workbook_relationship_id: owner.r_id().to_owned(),
            workbook_relationship_target: owner.target_ref().to_owned(),
            data_relationship_id: data_relationship.r_id().to_owned(),
            data_relationship_target: data_relationship.target_ref().to_owned(),
            source_id: Some(properties.id.clone()),
        };
        entries.push(Storage {
            value: CustomDataView {
                properties: Arc::new(properties),
                data: data_part.blob_arc(),
            },
            origin: Some(StorageOrigin::Existing(identity)),
        });
        charge_context(context, Resource::Objects, 1)?;
    }
    for part in package.iter_parts() {
        if part.content_type() == DATA_CONTENT_TYPE
            && !data_targets
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

fn capture_member_proofs(
    package: &OpcPackage,
    entries: &[Storage],
    limits: Limits,
    context: Option<&ExecutionContext>,
) -> Result<Vec<MemberProof>> {
    let mut targets = Vec::new();
    targets
        .try_reserve(
            entries
                .len()
                .checked_mul(2)
                .ok_or_else(|| invalid("Custom Data inbound proof target count overflow"))?,
        )
        .map_err(|source| allocation("Custom Data inbound proof targets", source))?;
    for entry in entries {
        let identity = entry
            .origin
            .as_ref()
            .ok_or_else(|| invalid("Custom Data entry has no source identity"))?
            .identity();
        targets.push(identity.properties_part_name.clone());
        targets.push(identity.data_part_name.clone());
    }
    // Capture every incoming edge once and distribute it to the affected
    // Custom Data members. The former per-storage full graph scan multiplied
    // relationship work and temporary capacity by the number of storages.
    let mut inbound = capture_inbound_proofs(package, &targets, limits, context)?;
    let mut members = Vec::new();
    members
        .try_reserve(entries.len())
        .map_err(|source| allocation("Custom Data member proofs", source))?;
    for entry in entries {
        check_context(context)?;
        let identity = entry
            .origin
            .as_ref()
            .ok_or_else(|| invalid("Custom Data entry has no source identity"))?
            .identity();
        let properties = package.get_part(&identity.properties_part_name)?;
        let data = package.get_part(&identity.data_part_name)?;
        let properties_relationships =
            package.source_relationships(&identity.properties_part_name)?;
        let data_relationships = package.source_relationships(&identity.data_part_name)?;
        if properties_relationships.bytes().len() > limits.max_relationship_xml_bytes()
            || data_relationships.bytes().len() > limits.max_relationship_xml_bytes()
        {
            return Err(Error::LimitExceeded {
                resource: "Custom Data relationship XML bytes",
                actual: properties_relationships
                    .bytes()
                    .len()
                    .max(data_relationships.bytes().len()),
                maximum: limits.max_relationship_xml_bytes(),
            });
        }
        let properties_xml = package.source_xml_part(&identity.properties_part_name).ok();
        let properties_key = inbound_key(&identity.properties_part_name);
        let data_key = inbound_key(&identity.data_part_name);
        let mut member_inbound = inbound.remove(&properties_key).unwrap_or_default();
        if data_key != properties_key {
            if let Some(data_inbound) = inbound.remove(&data_key) {
                member_inbound
                    .try_reserve(data_inbound.len())
                    .map_err(|source| allocation("Custom Data inbound proof merge", source))?;
                member_inbound.extend(data_inbound);
            }
        }
        member_inbound.sort_unstable_by(|left, right| {
            left.owner
                .as_str()
                .cmp(right.owner.as_str())
                .then_with(|| left.relationship_id.cmp(&right.relationship_id))
                .then_with(|| left.target.cmp(&right.target))
        });
        members.push(MemberProof {
            name: properties.partname().clone(),
            content_type: properties.content_type().to_owned(),
            bytes: properties.blob_arc(),
            data_bytes: data.blob_arc(),
            properties_xml,
            relationships: properties_relationships,
            data_relationships,
            inbound: member_inbound,
        });
        // The payload is intentionally inert; retaining its allocation in the
        // semantic view and MemberProof is enough to make stale checks exact.
    }
    Ok(members)
}

fn capture_inbound_proofs(
    package: &OpcPackage,
    targets: &[PackURI],
    limits: Limits,
    context: Option<&ExecutionContext>,
) -> Result<HashMap<String, Vec<InboundProof>>> {
    let mut target_keys = HashSet::new();
    target_keys
        .try_reserve(targets.len())
        .map_err(|source| allocation("Custom Data inbound proof keys", source))?;
    for target in targets {
        target_keys.insert(inbound_key(target));
    }

    let relationship_count = package_relationship_count(package)?;
    charge_context(context, Resource::Work, relationship_count)?;

    // Count first, then reserve exactly the number of matching edges. This
    // keeps one aggregate bound for all proof vectors and avoids reserving a
    // full package graph once per storage.
    let mut counts = HashMap::new();
    counts
        .try_reserve(target_keys.len())
        .map_err(|source| allocation("Custom Data inbound proof counts", source))?;
    for key in &target_keys {
        counts.insert(key.clone(), 0usize);
    }

    let mut count = |relationship: &litchi_opc::Relationship| -> Result<()> {
        let Some(key) = inbound_relationship_key(relationship, &target_keys) else {
            return Ok(());
        };
        let Some(value) = counts.get_mut(&key) else {
            return Err(invalid("Custom Data inbound proof target disappeared"));
        };
        *value = value
            .checked_add(1)
            .ok_or_else(|| invalid("Custom Data inbound relationship count overflow"))?;
        Ok(())
    };
    if let Ok(root) = PackURI::new("/") {
        for relationship in package.rels().iter() {
            check_context(context)?;
            count(relationship)?;
        }
        // Keep the root URI alive for the second pass below; the first pass
        // only counts and does not retain an owner.
        drop(root);
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            check_context(context)?;
            count(relationship)?;
        }
    }

    let mut proofs = HashMap::new();
    proofs
        .try_reserve(counts.len())
        .map_err(|source| allocation("Custom Data inbound relationship proofs", source))?;
    for (key, count) in counts {
        let mut values = Vec::new();
        values
            .try_reserve(count)
            .map_err(|source| allocation("Custom Data inbound relationship proofs", source))?;
        proofs.insert(key, values);
    }

    let mut append = |owner: &PackURI, relationship: &litchi_opc::Relationship| -> Result<()> {
        let Some(key) = inbound_relationship_key(relationship, &target_keys) else {
            return Ok(());
        };
        let Some(values) = proofs.get_mut(&key) else {
            return Err(invalid("Custom Data inbound proof target disappeared"));
        };
        values.push(InboundProof {
            owner: owner.clone(),
            relationship_id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        });
        Ok(())
    };
    if let Ok(root) = PackURI::new("/") {
        for relationship in package.rels().iter() {
            check_context(context)?;
            append(&root, relationship)?;
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            check_context(context)?;
            append(part.partname(), relationship)?;
        }
    }
    for values in proofs.values_mut() {
        values.sort_unstable_by(|left, right| {
            left.owner
                .as_str()
                .cmp(right.owner.as_str())
                .then_with(|| left.relationship_id.cmp(&right.relationship_id))
                .then_with(|| left.target.cmp(&right.target))
        });
    }
    // The package relationship limit is already checked at ingress. Keep the
    // explicit guard here as the aggregate proof quota in case that policy is
    // tightened independently in the future.
    let matched = proofs.values().map(Vec::len).sum::<usize>();
    if matched > limits.max_relationships() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data inbound relationship proofs",
            actual: matched,
            maximum: limits.max_relationships(),
        });
    }
    Ok(proofs)
}

fn inbound_key(target: &PackURI) -> String {
    target.as_str().to_ascii_lowercase()
}

fn inbound_relationship_key(
    relationship: &litchi_opc::Relationship,
    targets: &HashSet<String>,
) -> Option<String> {
    if let Ok(target) = relationship.target_partname() {
        let key = inbound_key(&target);
        if targets.contains(&key) {
            return Some(key);
        }
    }
    let key = relationship.target_ref().to_ascii_lowercase();
    targets.contains(&key).then_some(key)
}

fn validate_feature_relationships(package: &OpcPackage, workbook_name: &PackURI) -> Result<()> {
    for relationship in package.rels().iter() {
        if relationship.reltype() == PROPERTIES_RELATIONSHIP_TYPE
            || relationship.reltype() == DATA_RELATIONSHIP_TYPE
            || relationship.reltype() == CONNECTIONS_RELATIONSHIP_TYPE
        {
            return Err(invalid(
                "Custom Data relationship cannot originate from the package root",
            ));
        }
    }
    let workbook = package.get_part(workbook_name)?;
    for relationship in workbook.rels().iter() {
        if relationship.reltype() != PROPERTIES_RELATIONSHIP_TYPE {
            continue;
        }
        if relationship.is_external()
            || relationship.target_query().is_some()
            || relationship.target_fragment().is_some()
        {
            return Err(invalid(
                "Custom Data Properties relationship must be internal and plain",
            ));
        }
        let target = relationship.target_partname()?;
        let target_part = package.get_part(&target)?;
        if target_part.content_type() != PROPERTIES_CONTENT_TYPE {
            return Err(invalid(
                "Custom Data Properties relationship targets a non-Properties part",
            ));
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if relationship.reltype() == PROPERTIES_RELATIONSHIP_TYPE
                && part.partname() != workbook_name
            {
                return Err(invalid(
                    "Custom Data Properties relationship must originate from the workbook",
                ));
            }
            if relationship.reltype() == CONNECTIONS_RELATIONSHIP_TYPE
                && part.partname() != workbook_name
            {
                return Err(invalid(
                    "connections relationship must originate from the workbook",
                ));
            }
            if relationship.reltype() != DATA_RELATIONSHIP_TYPE {
                continue;
            }
            if relationship.is_external()
                || relationship.target_query().is_some()
                || relationship.target_fragment().is_some()
            {
                return Err(invalid(
                    "Custom Data relationship must be internal and plain",
                ));
            }
            if part.content_type() != PROPERTIES_CONTENT_TYPE {
                return Err(invalid(
                    "Custom Data relationship must originate from a Properties part",
                ));
            }
            let target = relationship.target_partname()?;
            let target_part = package.get_part(&target)?;
            if target_part.content_type() != DATA_CONTENT_TYPE {
                return Err(invalid(
                    "Custom Data relationship targets a non-payload part",
                ));
            }
        }
    }
    Ok(())
}

fn capture_connection(
    package: &OpcPackage,
    workbook_name: &PackURI,
    limits: Limits,
    context: Option<&ExecutionContext>,
) -> Result<Option<ConnectionState>> {
    check_context(context)?;
    let workbook = package.get_part(workbook_name)?;
    let mut relationship = None;
    for candidate in workbook.rels().iter() {
        if candidate.reltype() != CONNECTIONS_RELATIONSHIP_TYPE {
            continue;
        }
        if relationship.is_some() {
            return Err(invalid(
                "workbook declares multiple connections relationships",
            ));
        }
        relationship = Some(candidate);
    }
    let mut connection_part = None;
    for part in package.iter_parts() {
        if part.content_type() == CONNECTIONS_CONTENT_TYPE {
            if connection_part.is_some() {
                return Err(invalid(
                    "package contains an orphan or additional connections part",
                ));
            }
            connection_part = Some(part);
        }
    }
    let Some(relationship) = relationship else {
        if connection_part.is_some() {
            return Err(invalid("package contains an orphan connections part"));
        }
        return Ok(None);
    };
    if relationship.is_external()
        || relationship.target_query().is_some()
        || relationship.target_fragment().is_some()
    {
        return Err(invalid(
            "connections relationship must be internal and plain",
        ));
    }
    let part_name = relationship.target_partname()?;
    let part = package.get_part(&part_name)?;
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(Error::InvalidContentType {
            expected: CONNECTIONS_CONTENT_TYPE.to_owned(),
            got: part.content_type().to_owned(),
        });
    }
    if !connection_part
        .is_some_and(|candidate| candidate.partname().is_equivalent_to(part.partname()))
    {
        return Err(invalid(
            "package contains an orphan or additional connections part",
        ));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "connections part must not have outbound relationships",
        ));
    }
    if part.blob().len() > limits.max_connections_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data connections.bin bytes",
            actual: part.blob().len(),
            maximum: limits.max_connections_bytes(),
        });
    }
    let mut binding_limits = limits.binding_limits();
    // A longer UID can grow the BIFF12 stream. Make the rewrite planner see
    // the strictest output ceiling up front, before it allocates its output
    // buffer; post-rewrite checks alone are too late for a bounded operation.
    binding_limits.max_output_bytes = binding_limits
        .max_output_bytes
        .min(limits.max_connections_bytes())
        .min(limits.max_temporary_bytes());
    let mut budget = ContextWorkBudget {
        context,
        rejected: None,
    };
    let scan =
        bindings::scan_connections_with_budget(part.blob(), binding_limits, &[], &mut budget)
            .map_err(|error| map_binding_error_with_budget(error, &mut budget))?;
    let mut references = Vec::new();
    references
        .try_reserve(scan.bindings().len())
        .map_err(|source| allocation("Custom Data connection references", source))?;
    for binding in scan.bindings() {
        check_context(context)?;
        // Retain empty fields as positional slots so detach patches and their
        // exact inverses can reconstruct the original source binding.
        references.push(binding.client_cube_urn().to_owned());
    }
    let part_relationships = package.source_relationships(part.partname())?;
    let mut inbound = capture_inbound_proofs(
        package,
        std::slice::from_ref(part.partname()),
        limits,
        context,
    )?;
    Ok(Some(ConnectionState {
        part_name: part.partname().clone(),
        content_type: part.content_type().to_owned(),
        bytes: part.blob_arc(),
        relationships: part_relationships,
        inbound: inbound
            .remove(&inbound_key(part.partname()))
            .unwrap_or_default(),
        workbook_relationship_id: relationship.r_id().to_owned(),
        workbook_relationship_target: relationship.target_ref().to_owned(),
        references,
        blocking_opaque: scan.has_blocking_opaque(),
        record_count: scan.record_count(),
    }))
}

struct ContextWorkBudget<'a> {
    context: Option<&'a ExecutionContext>,
    rejected: Option<litchi_core::ExecutionError>,
}

impl WorkBudget for ContextWorkBudget<'_> {
    fn check(&mut self) -> std::result::Result<(), BindingError> {
        if let Some(context) = self.context {
            if let Err(error) = context.check() {
                self.rejected = Some(error.clone());
                return Err(BindingError::BudgetRejected);
            }
        }
        Ok(())
    }

    fn charge(&mut self, units: usize) -> std::result::Result<(), BindingError> {
        self.check()?;
        if let Some(context) = self.context {
            let amount = u64::try_from(units).map_err(|_error| BindingError::LengthOverflow)?;
            if let Err(error) = context.consume(Resource::Work, amount) {
                self.rejected = Some(error.clone());
                return Err(BindingError::BudgetRejected);
            }
        }
        Ok(())
    }
}

fn validate_connection_ids(
    entries: &[Storage],
    connection: Option<&ConnectionState>,
) -> Result<()> {
    let mut ids = HashSet::new();
    ids.try_reserve(entries.len())
        .map_err(|source| allocation("Custom Data connection UID index", source))?;
    ids.extend(entries.iter().map(Storage::id));
    if let Some(connection) = connection {
        for id in &connection.references {
            if id.is_empty() {
                // An empty ST_Xstring is a valid detached ExtConn14 value.
                // It is intentionally outside the Properties UID owner set.
                continue;
            }
            if !ids.contains(&id.as_str()) {
                return Err(invalid(format!(
                    "ExtConn14 Custom Data UID '{id}' has no Properties owner"
                )));
            }
        }
    }
    Ok(())
}

fn allocate_identity(package: &OpcPackage, entries: &[Storage]) -> Result<PartIdentity> {
    let workbook = package.main_document_part()?;
    let mut index = 1usize;
    let mut used_ids = HashSet::new();
    used_ids
        .try_reserve(workbook.rels().iter().count())
        .map_err(|source| allocation("Custom Data workbook relationship IDs", source))?;
    for relationship in workbook.rels().iter() {
        used_ids.insert(relationship.r_id().to_owned());
    }
    loop {
        let properties_part_name =
            PackURI::new(format!("/xl/customData/props{index}.xml")).map_err(invalid)?;
        let data_part_name =
            PackURI::new(format!("/xl/customData/data{index}.bin")).map_err(invalid)?;
        let workbook_relationship_id = if index == 1 {
            "rIdCustomDataProps".to_owned()
        } else {
            format!("rIdCustomDataProps{index}")
        };
        let data_relationship_id = if index == 1 {
            "rIdCustomData".to_owned()
        } else {
            format!("rIdCustomData{index}")
        };
        let collides = package.get_part(&properties_part_name).is_ok()
            || package.get_part(&data_part_name).is_ok()
            || entries.iter().any(|entry| {
                entry.origin.as_ref().is_some_and(|origin| {
                    let identity = origin.identity();
                    identity
                        .properties_part_name
                        .is_equivalent_to(&properties_part_name)
                        || identity.data_part_name.is_equivalent_to(&data_part_name)
                })
            })
            || used_ids.contains(&workbook_relationship_id);
        if !collides {
            return Ok(PartIdentity {
                properties_part_name: properties_part_name.clone(),
                data_part_name: data_part_name.clone(),
                workbook_relationship_id,
                workbook_relationship_target: properties_part_name
                    .relative_ref(workbook.partname().base_uri()),
                data_relationship_id,
                data_relationship_target: data_part_name
                    .relative_ref(properties_part_name.base_uri()),
                source_id: None,
            });
        }
        index = index
            .checked_add(1)
            .ok_or_else(|| invalid("Custom Data part-name space exhausted"))?;
    }
}

fn identity_of(storage: &Storage) -> Option<&PartIdentity> {
    storage.origin.as_ref().map(StorageOrigin::identity)
}

fn same_identity(left: &Storage, right: &Storage) -> bool {
    match (identity_of(left), identity_of(right)) {
        (Some(left), Some(right)) => left.same_physical(right),
        _ => false,
    }
}

fn transition_content_types(
    before: &SourceState,
    before_entries: &[Storage],
    after_entries: &[Storage],
    limits: Limits,
    output_limit: usize,
) -> Result<OwnedContentTypes> {
    let mut removed = Vec::new();
    removed
        .try_reserve(
            before_entries
                .len()
                .checked_mul(2)
                .ok_or_else(|| invalid("Custom Data content-types removal count overflow"))?,
        )
        .map_err(|source| allocation("Custom Data content-types removals", source))?;
    for entry in before_entries {
        if !after_entries
            .iter()
            .any(|candidate| same_identity(entry, candidate))
        {
            let identity =
                identity_of(entry).ok_or_else(|| invalid("Custom Data source identity missing"))?;
            removed.push(identity.properties_part_name.clone());
            removed.push(identity.data_part_name.clone());
        }
    }
    let mut additions = Vec::new();
    additions
        .try_reserve(
            after_entries
                .len()
                .checked_mul(2)
                .ok_or_else(|| invalid("Custom Data content-types addition count overflow"))?,
        )
        .map_err(|source| allocation("Custom Data content-types additions", source))?;
    for entry in after_entries {
        if !before_entries
            .iter()
            .any(|candidate| same_identity(entry, candidate))
        {
            let identity =
                identity_of(entry).ok_or_else(|| invalid("Custom Data staged identity missing"))?;
            additions.push((&identity.properties_part_name, PROPERTIES_CONTENT_TYPE));
            additions.push((&identity.data_part_name, DATA_CONTENT_TYPE));
        }
    }
    // Content-types XML is a relationship-graph replacement. Cap the
    // delegated builder by both its format ceiling and the caller's temporary
    // allocation policy before it reserves an output buffer.
    let maximum = limits
        .max_relationship_xml_bytes()
        .min(limits.max_temporary_bytes())
        .min(limits.max_output_bytes())
        .min(output_limit);
    let mut result = before
        .content_types
        .without_parts(&removed, maximum)
        .map_err(|error| map_content_types_error(error, maximum))?;
    if !additions.is_empty() {
        result = result
            .with_part_overrides(&additions, maximum)
            .map_err(|error| map_content_types_error(error, maximum))?;
    }
    if result.bytes().len() > limits.max_relationship_xml_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data content-types XML bytes",
            actual: result.bytes().len(),
            maximum: limits.max_relationship_xml_bytes(),
        });
    }
    Ok(result)
}

fn map_content_types_error(error: litchi_opc::OpcError, maximum: usize) -> Error {
    let detail = error.to_string();
    if detail.contains("output exceeds")
        || detail.contains("source exceeds")
        || detail.contains("selectors exceed")
    {
        return Error::LimitExceeded {
            resource: "Custom Data content-types XML bytes",
            actual: maximum.saturating_add(1),
            maximum,
        };
    }
    Error::Opc(error)
}

/// Restore the retained content-types lexical placement when a reverse patch
/// re-adds a removed Custom Data pair. `with_part_overrides` appends new
/// overrides, which is safe for ordinary publication but would move source
/// entries when applying an inverse patch. Use the exact after-template only
/// when every non-Custom-Data byte in the current target still matches the
/// template; otherwise leave the target's unrelated manifest edits intact and
/// let the source-closure postcondition refuse the patch.
fn restore_content_types_template_if_scoped(
    package: &mut OpcPackage,
    before: &Snapshot,
    after: &Snapshot,
    limits: Limits,
) -> Result<()> {
    let current = package.source_content_types()?;
    if current == after.source.content_types {
        return Ok(());
    }
    let mut custom_parts = Vec::new();
    custom_parts
        .try_reserve(
            before
                .entries
                .len()
                .checked_add(after.entries.len())
                .and_then(|count| count.checked_mul(2))
                .ok_or_else(|| invalid("Custom Data content-types template part count overflow"))?,
        )
        .map_err(|source| allocation("Custom Data content-types template parts", source))?;
    for entry in before.entries.iter().chain(after.entries.iter()) {
        let Some(identity) = identity_of(entry) else {
            continue;
        };
        custom_parts.push(identity.properties_part_name.clone());
        custom_parts.push(identity.data_part_name.clone());
    }
    let maximum = limits
        .max_relationship_xml_bytes()
        .min(limits.max_temporary_bytes())
        .min(limits.max_output_bytes());
    let current_without = current
        .without_parts(&custom_parts, maximum)
        .map_err(|error| map_content_types_error(error, maximum))?;
    let after_without = after
        .source
        .content_types
        .without_parts(&custom_parts, maximum)
        .map_err(|error| map_content_types_error(error, maximum))?;
    if current_without != after_without {
        return Ok(());
    }
    if after.source.content_types.bytes().len() > limits.max_relationship_xml_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data content-types XML bytes",
            actual: after.source.content_types.bytes().len(),
            maximum: limits.max_relationship_xml_bytes(),
        });
    }
    let output_limit = replacement_output_limit(package, current.bytes().len(), limits)?;
    if after.source.content_types.bytes().len() > output_limit {
        return Err(Error::LimitExceeded {
            resource: "Custom Data final output bytes",
            actual: after.source.content_types.bytes().len(),
            maximum: output_limit,
        });
    }
    package.try_replace_content_types(current.bytes(), &after.source.content_types)?;
    Ok(())
}

/// Restore exact relationship-member placement for a reverse patch when the
/// target's non-Custom-Data edges are unchanged. Relationship builders append
/// newly added edges; source templates let an inverse return those edges to
/// their original lexical positions without touching unrelated relationships.
fn restore_relationship_templates_if_scoped(
    package: &mut OpcPackage,
    before: &Snapshot,
    after: &Snapshot,
    limits: Limits,
) -> Result<()> {
    let maximum = limits
        .max_relationship_xml_bytes()
        .min(limits.max_temporary_bytes())
        .min(limits.max_output_bytes());
    let workbook_name = package.main_document_part()?.partname().clone();
    let current_workbook = package.source_relationships(&workbook_name)?;
    let mut workbook_ids = Vec::new();
    workbook_ids
        .try_reserve(
            before
                .entries
                .len()
                .checked_add(after.entries.len())
                .ok_or_else(|| invalid("Custom Data workbook relationship ID count overflow"))?,
        )
        .map_err(|source| allocation("Custom Data workbook relationship IDs", source))?;
    for entry in before.entries.iter().chain(after.entries.iter()) {
        if let Some(identity) = identity_of(entry) {
            if !workbook_ids
                .iter()
                .any(|candidate: &String| candidate == &identity.workbook_relationship_id)
            {
                workbook_ids.push(identity.workbook_relationship_id.clone());
            }
        }
    }
    let desired_workbook = &after.source.workbook_relationships;
    restore_one_relationship_template(
        package,
        &current_workbook,
        desired_workbook,
        &workbook_ids,
        maximum,
        limits,
    )?;

    for desired_member in &after.source.members {
        let mut data_ids = Vec::new();
        data_ids
            .try_reserve(2)
            .map_err(|source| allocation("Custom Data relationship selectors", source))?;
        for entry in before.entries.iter().chain(after.entries.iter()) {
            let Some(identity) = identity_of(entry) else {
                continue;
            };
            if identity
                .properties_part_name
                .is_equivalent_to(&desired_member.name)
                && !data_ids
                    .iter()
                    .any(|candidate: &String| candidate == &identity.data_relationship_id)
            {
                data_ids.push(identity.data_relationship_id.clone());
            }
        }
        let current_member = package.source_relationships(&desired_member.name)?;
        restore_one_relationship_template(
            package,
            &current_member,
            &desired_member.relationships,
            &data_ids,
            maximum,
            limits,
        )?;
    }
    Ok(())
}

fn restore_one_relationship_template(
    package: &mut OpcPackage,
    current: &OwnedRelationships,
    desired: &OwnedRelationships,
    custom_ids: &[String],
    maximum: usize,
    limits: Limits,
) -> Result<()> {
    if current == desired {
        return Ok(());
    }
    let mut selectors = Vec::new();
    selectors
        .try_reserve(custom_ids.len())
        .map_err(|source| allocation("Custom Data relationship selectors", source))?;
    selectors.extend(custom_ids.iter().map(String::as_str));
    let current_without = current
        .without_relationships(&selectors, maximum)
        .map_err(Error::Opc)?;
    let desired_without = desired
        .without_relationships(&selectors, maximum)
        .map_err(Error::Opc)?;
    if current_without != desired_without {
        return Ok(());
    }
    if desired.bytes().len() > limits.max_relationship_xml_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data relationship XML bytes",
            actual: desired.bytes().len(),
            maximum: limits.max_relationship_xml_bytes(),
        });
    }
    let output_limit = replacement_output_limit(package, current.bytes().len(), limits)?;
    if desired.bytes().len() > output_limit {
        return Err(Error::LimitExceeded {
            resource: "Custom Data final output bytes",
            actual: desired.bytes().len(),
            maximum: output_limit,
        });
    }
    package.try_replace_relationships(current, desired)?;
    Ok(())
}

/// Restore source XML and payload bytes for Custom Data parts reintroduced by
/// an inverse patch. A semantic `CustomData` value does not retain its source
/// lexical XML, so the ordinary insertion path necessarily writes canonical
/// properties; the patch's retained member proof supplies the exact bytes.
fn restore_member_templates_if_scoped(
    package: &mut OpcPackage,
    after: &Snapshot,
    limits: Limits,
) -> Result<()> {
    for member in &after.source.members {
        let properties_match = package
            .get_part(&member.name)?
            .blob()
            .eq(member.bytes.as_slice());
        if !properties_match {
            check_restored_properties_limits(member.bytes.len(), limits)?;
            let Some(replacement) = member.properties_xml.as_ref() else {
                continue;
            };
            check_restored_properties_limits(replacement.bytes().len(), limits)?;
            let current = package.source_xml_part(&member.name)?;
            let output_limit = replacement_output_limit(package, current.bytes().len(), limits)?;
            if replacement.bytes().len() > output_limit {
                return Err(Error::LimitExceeded {
                    resource: "Custom Data final output bytes",
                    actual: replacement.bytes().len(),
                    maximum: output_limit,
                });
            }
            package.try_replace_owned_xml_part(current.bytes(), replacement.clone())?;
        }
        let data_name = after.entries.iter().find_map(|entry| {
            let identity = identity_of(entry)?;
            identity
                .properties_part_name
                .is_equivalent_to(&member.name)
                .then_some(identity.data_part_name.clone())
        });
        if let Some(data_name) = data_name {
            let data = package.get_part(&data_name)?;
            if data.blob() != member.data_bytes.as_slice() {
                let output_limit = replacement_output_limit(package, data.blob().len(), limits)?;
                if member.data_bytes.len() > output_limit {
                    return Err(Error::LimitExceeded {
                        resource: "Custom Data final output bytes",
                        actual: member.data_bytes.len(),
                        maximum: output_limit,
                    });
                }
                package
                    .get_part_mut(&data_name)?
                    .set_blob_shared(Arc::clone(&member.data_bytes));
            }
        }
    }
    Ok(())
}

fn check_restored_properties_limits(bytes: usize, limits: Limits) -> Result<()> {
    if bytes > limits.max_temporary_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data temporary Properties restoration bytes",
            actual: bytes,
            maximum: limits.max_temporary_bytes(),
        });
    }
    if bytes > limits.max_properties_xml_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data Properties XML bytes",
            actual: bytes,
            maximum: limits.max_properties_xml_bytes(),
        });
    }
    Ok(())
}

fn apply_draft(
    package: &mut OpcPackage,
    before: &Snapshot,
    draft: &[Storage],
    connection_edits: &BTreeMap<String, String>,
    limits: Limits,
    context: Option<&ExecutionContext>,
) -> Result<()> {
    check_context(context)?;
    validate_staged_entries(draft, &limits)?;
    preflight_output_budget(package, before, limits)?;

    // Remove old graph pairs first in the detached candidate. The transaction
    // has already planned any BIFF detachment/retargeting, so this intermediate
    // state is never visible to a caller.
    for entry in before.entries.iter() {
        check_context(context)?;
        if !draft
            .iter()
            .any(|candidate| same_identity(entry, candidate))
        {
            remove_entry(package, entry, limits)?;
        }
    }
    for entry in draft {
        check_context(context)?;
        let Some(previous) = before
            .entries
            .iter()
            .find(|candidate| same_identity(entry, candidate))
        else {
            add_entry(package, entry, &limits)?;
            continue;
        };
        apply_existing_entry(package, previous, entry, &before.source, &limits)?;
    }

    apply_connection_edits(package, before, connection_edits, limits, context)?;
    let current_content_types = package.source_content_types()?;
    let content_types_limit =
        replacement_output_limit(package, current_content_types.bytes().len(), limits)?;
    let content_types = transition_content_types(
        &before.source,
        &before.entries,
        draft,
        limits,
        content_types_limit,
    )?;
    package.try_replace_content_types(current_content_types.bytes(), &content_types)?;
    validate_package_limits(package, limits)?;
    validate_output_limits(package, limits)?;
    Ok(())
}

fn replacement_output_limit(
    package: &OpcPackage,
    replaced_bytes: usize,
    limits: Limits,
) -> Result<usize> {
    let current = package_size(package)?;
    let base = current
        .checked_sub(replaced_bytes)
        .ok_or_else(|| invalid("Custom Data replacement source size exceeds package size"))?;
    Ok(limits.max_output_bytes().saturating_sub(base))
}

/// Reject an output ceiling that is already below the bytes guaranteed to
/// survive a Custom Data edit. The mutable Custom Data parts, content types,
/// package relationships, workbook relationships, and connections stream are
/// excluded from this lower bound because their exact lengths can change. All
/// other part payloads and relationship images are untouched by this owner and
/// therefore make an exact, allocation-free minimum for the final package.
fn preflight_output_budget(package: &OpcPackage, before: &Snapshot, limits: Limits) -> Result<()> {
    let maximum = limits.max_output_bytes();
    let mut mutable_names = HashSet::new();
    mutable_names
        .try_reserve(
            before
                .source
                .members
                .len()
                .checked_add(
                    before
                        .entries
                        .len()
                        .checked_mul(2)
                        .ok_or_else(|| invalid("Custom Data output name count overflow"))?,
                )
                .ok_or_else(|| invalid("Custom Data output name count overflow"))?,
        )
        .map_err(|source| allocation("Custom Data output mutable names", source))?;
    for member in &before.source.members {
        mutable_names.insert(inbound_key(&member.name));
    }
    for entry in before.entries.iter() {
        if let Some(identity) = identity_of(entry) {
            mutable_names.insert(inbound_key(&identity.properties_part_name));
            mutable_names.insert(inbound_key(&identity.data_part_name));
        }
    }
    let mut immutable = 0usize;
    let connection_name = before
        .source
        .connection
        .as_ref()
        .map(|connection| &connection.part_name);
    for part in package.iter_parts() {
        let custom_member = mutable_names.contains(&inbound_key(part.partname()));
        let connection_member =
            connection_name.is_some_and(|name| name.is_equivalent_to(part.partname()));
        if custom_member || connection_member {
            continue;
        }
        immutable = immutable
            .checked_add(part.blob().len())
            .ok_or_else(|| invalid("Custom Data output byte count overflow"))?;
        immutable = immutable
            .checked_add(package.source_relationships(part.partname())?.bytes().len())
            .ok_or_else(|| invalid("Custom Data output relationship byte count overflow"))?;
    }
    if immutable > maximum {
        return Err(Error::LimitExceeded {
            resource: "Custom Data final output bytes",
            actual: immutable,
            maximum,
        });
    }
    Ok(())
}

fn source_member<'a>(source: &'a SourceState, entry: &Storage) -> Result<&'a MemberProof> {
    let identity =
        identity_of(entry).ok_or_else(|| invalid("Custom Data source identity missing"))?;
    source
        .members
        .iter()
        .find(|member| member.name.is_equivalent_to(&identity.properties_part_name))
        .ok_or_else(|| invalid("Custom Data source member proof is missing"))
}

fn apply_existing_entry(
    package: &mut OpcPackage,
    previous: &Storage,
    entry: &Storage,
    source: &SourceState,
    limits: &Limits,
) -> Result<()> {
    let identity =
        identity_of(entry).ok_or_else(|| invalid("Custom Data staged identity missing"))?;
    let member = source_member(source, previous)?;
    let properties_changed = previous.value.properties != entry.value.properties;
    let data_changed = previous.value.data != entry.value.data;
    if properties_changed {
        if member.bytes.len() > limits.max_temporary_bytes() {
            return Err(Error::LimitExceeded {
                resource: "Custom Data temporary Properties source bytes",
                actual: member.bytes.len(),
                maximum: limits.max_temporary_bytes(),
            });
        }
        let proof = member.properties_xml.as_ref().ok_or_else(|| {
            unsupported("Custom Data Properties source has no retained XML provenance for editing")
        })?;
        // Common codec writers check their output cap before reserving their
        // replacement buffer. Tighten that cap with the temporary allocation
        // and exact remaining package output budget so a rewrite cannot
        // allocate first and fail the aggregate publication check afterwards.
        let output_limit = replacement_output_limit(package, member.bytes.len(), *limits)?;
        let xml_limits =
            codec_limits_with_output(*limits, limits.max_temporary_bytes().min(output_limit));
        let extension_same = canonical_extension_with_limits(
            previous.value.properties.extension_list.as_ref(),
            &xml_limits,
        )
        .map_err(map_common_error)?
            == canonical_extension_with_limits(
                entry.value.properties.extension_list.as_ref(),
                &xml_limits,
            )
            .map_err(map_common_error)?;
        let mut replacement = if extension_same {
            rewrite_id_with_limits(proof, entry.id(), &xml_limits).map_err(map_common_error)?
        } else {
            rewrite_extension_list_with_limits(
                proof,
                entry.value.properties.extension_list.as_ref(),
                &xml_limits,
            )
            .map_err(map_common_error)?
        };
        if !extension_same && previous.id() != entry.id() {
            replacement = rewrite_id_with_limits(&replacement, entry.id(), &xml_limits)
                .map_err(map_common_error)?;
        }
        if replacement.bytes().len() > limits.max_properties_xml_bytes() {
            return Err(Error::LimitExceeded {
                resource: "Custom Data Properties XML bytes",
                actual: replacement.bytes().len(),
                maximum: limits.max_properties_xml_bytes(),
            });
        }
        if replacement.bytes().len() > limits.max_temporary_bytes() {
            return Err(Error::LimitExceeded {
                resource: "Custom Data temporary Properties replacement bytes",
                actual: replacement.bytes().len(),
                maximum: limits.max_temporary_bytes(),
            });
        }
        package.try_replace_owned_xml_part(&member.bytes, replacement)?;
    }
    if data_changed {
        let current = package.get_part(&identity.data_part_name)?;
        if current.blob() != previous.value.data.as_slice() {
            return Err(conflict("Custom Data payload source"));
        }
        let output_limit = replacement_output_limit(package, previous.value.data.len(), *limits)?;
        if entry.value.data.len() > output_limit {
            return Err(Error::LimitExceeded {
                resource: "Custom Data final output bytes",
                actual: entry.value.data.len(),
                maximum: output_limit,
            });
        }
        package
            .get_part_mut(&identity.data_part_name)?
            .set_blob_shared(Arc::clone(&entry.value.data));
    }
    Ok(())
}

fn add_entry(package: &mut OpcPackage, entry: &Storage, limits: &Limits) -> Result<()> {
    let identity =
        identity_of(entry).ok_or_else(|| invalid("Custom Data staged identity missing"))?;
    if package.get_part(&identity.properties_part_name).is_ok()
        || package.get_part(&identity.data_part_name).is_ok()
    {
        return Err(conflict("Custom Data insertion part identity"));
    }
    validate_value(&entry.value, limits)?;
    if entry.value.data.len() > limits.max_temporary_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data temporary payload bytes",
            actual: entry.value.data.len(),
            maximum: limits.max_temporary_bytes(),
        });
    }
    let output_limit = replacement_output_limit(package, 0, *limits)?;
    let properties_xml = write_properties_with_limits(
        &entry.value.properties,
        &codec_limits_with_output(*limits, limits.max_temporary_bytes().min(output_limit)),
    )
    .map_err(map_common_error)?;
    if properties_xml.len() > limits.max_properties_xml_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data Properties XML bytes",
            actual: properties_xml.len(),
            maximum: limits.max_properties_xml_bytes(),
        });
    }
    if properties_xml.len() > limits.max_temporary_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data temporary Properties bytes",
            actual: properties_xml.len(),
            maximum: limits.max_temporary_bytes(),
        });
    }
    let workbook_name = package.main_document_part()?.partname().clone();
    let workbook_relationships_before = package.source_relationships(&workbook_name)?;
    if workbook_relationships_before
        .bytes()
        .len()
        .checked_add(properties_xml.len())
        .is_none()
    {
        return Err(invalid("Custom Data insertion output size overflow"));
    }
    package.try_add_owned_xml_part_bytes(
        identity.properties_part_name.clone(),
        PROPERTIES_CONTENT_TYPE.to_owned(),
        Arc::new(properties_xml),
    )?;
    let data_output_limit = replacement_output_limit(package, 0, *limits)?;
    if entry.value.data.len() > data_output_limit {
        return Err(Error::LimitExceeded {
            resource: "Custom Data final output bytes",
            actual: entry.value.data.len(),
            maximum: data_output_limit,
        });
    }
    package.try_add_part(Box::new(BlobPart::new(
        identity.data_part_name.clone(),
        DATA_CONTENT_TYPE.to_owned(),
        entry.value.data.as_ref().clone(),
    )))?;

    let properties_relationships = package.source_relationships(&identity.properties_part_name)?;
    let properties_relationship_output_limit =
        replacement_output_limit(package, properties_relationships.bytes().len(), *limits)?
            .min(limits.max_relationship_xml_bytes())
            .min(limits.max_temporary_bytes());
    let properties_replacement = properties_relationships.with_relationship(
        DATA_RELATIONSHIP_TYPE,
        &identity.data_relationship_target,
        &identity.data_relationship_id,
        TargetMode::Internal,
        properties_relationship_output_limit,
    )?;
    package.try_replace_relationships(&properties_relationships, &properties_replacement)?;
    let workbook_relationships = package.source_relationships(&workbook_name)?;
    let workbook_relationship_output_limit =
        replacement_output_limit(package, workbook_relationships.bytes().len(), *limits)?
            .min(limits.max_relationship_xml_bytes())
            .min(limits.max_temporary_bytes());
    let workbook_replacement = workbook_relationships.with_relationship(
        PROPERTIES_RELATIONSHIP_TYPE,
        &identity.workbook_relationship_target,
        &identity.workbook_relationship_id,
        TargetMode::Internal,
        workbook_relationship_output_limit,
    )?;
    package.try_replace_relationships(&workbook_relationships, &workbook_replacement)?;
    Ok(())
}

fn remove_entry(package: &mut OpcPackage, entry: &Storage, limits: Limits) -> Result<()> {
    let identity =
        identity_of(entry).ok_or_else(|| invalid("Custom Data source identity missing"))?;
    let workbook = package.main_document_part()?;
    let owner = workbook
        .rels()
        .get(&identity.workbook_relationship_id)
        .ok_or_else(|| conflict("Custom Data workbook relationship"))?;
    if owner.reltype() != PROPERTIES_RELATIONSHIP_TYPE
        || owner.is_external()
        || owner.target_ref() != identity.workbook_relationship_target
        || !owner
            .target_partname()?
            .is_equivalent_to(&identity.properties_part_name)
    {
        return Err(conflict("Custom Data workbook relationship"));
    }
    let properties = package.get_part(&identity.properties_part_name)?;
    let payload = properties
        .rels()
        .get(&identity.data_relationship_id)
        .ok_or_else(|| conflict("Custom Data payload relationship"))?;
    if payload.reltype() != DATA_RELATIONSHIP_TYPE
        || payload.is_external()
        || payload.target_ref() != identity.data_relationship_target
        || !payload
            .target_partname()?
            .is_equivalent_to(&identity.data_part_name)
    {
        return Err(conflict("Custom Data payload relationship"));
    }
    if inbound_count(package, &identity.properties_part_name) != 1
        || inbound_count(package, &identity.data_part_name) != 1
    {
        return Err(invalid(
            "Custom Data part has an unexpected incoming relationship",
        ));
    }
    let workbook_relationships = package.source_relationships(workbook.partname())?;
    let workbook_relationship_output_limit =
        replacement_output_limit(package, workbook_relationships.bytes().len(), limits)?
            .min(limits.max_relationship_xml_bytes())
            .min(limits.max_temporary_bytes());
    let workbook_replacement = workbook_relationships.without_relationship(
        &identity.workbook_relationship_id,
        workbook_relationship_output_limit,
    )?;
    package.try_replace_relationships(&workbook_relationships, &workbook_replacement)?;
    let properties_relationships = package.source_relationships(&identity.properties_part_name)?;
    let properties_relationship_output_limit =
        replacement_output_limit(package, properties_relationships.bytes().len(), limits)?
            .min(limits.max_relationship_xml_bytes())
            .min(limits.max_temporary_bytes());
    let properties_replacement = properties_relationships.without_relationship(
        &identity.data_relationship_id,
        properties_relationship_output_limit,
    )?;
    package.try_replace_relationships(&properties_relationships, &properties_replacement)?;
    if !package.remove_part(&identity.properties_part_name)
        || !package.remove_part(&identity.data_part_name)
    {
        return Err(conflict("Custom Data part removal"));
    }
    Ok(())
}

fn inbound_count(package: &OpcPackage, target: &PackURI) -> usize {
    let package_count = package
        .rels()
        .iter()
        .filter(|relationship| {
            relationship
                .target_partname()
                .ok()
                .is_some_and(|candidate| candidate.is_equivalent_to(target))
        })
        .count();
    package_count
        + package
            .iter_parts()
            .map(|part| {
                part.rels()
                    .iter()
                    .filter(|relationship| {
                        relationship
                            .target_partname()
                            .ok()
                            .is_some_and(|candidate| candidate.is_equivalent_to(target))
                    })
                    .count()
            })
            .sum::<usize>()
}

fn apply_connection_edits(
    package: &mut OpcPackage,
    before: &Snapshot,
    edits: &BTreeMap<String, String>,
    limits: Limits,
    context: Option<&ExecutionContext>,
) -> Result<()> {
    if edits.is_empty() {
        return Ok(());
    }
    let state = before
        .source
        .connection
        .as_ref()
        .ok_or_else(|| conflict("connections.bin source"))?;
    let part = package.get_part(&state.part_name)?;
    if part.content_type() != state.content_type || part.blob() != state.bytes.as_slice() {
        return Err(conflict("connections.bin source"));
    }
    let source = part.blob_arc();
    if source.len() > limits.max_temporary_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data temporary connections source bytes",
            actual: source.len(),
            maximum: limits.max_temporary_bytes(),
        });
    }
    // The source bytes are retained by the snapshot; the rewrite output and
    // scanner scratch are temporary and must release their reservation when
    // this operation returns.
    let _memory_reservation = reserve_context(context, Resource::Memory, source.len())?;
    let mut budget = ContextWorkBudget {
        context,
        rejected: None,
    };
    let mut binding_limits = limits.binding_limits();
    // The binding rewrite allocates before the final package-size check.  Its
    // output ceiling must therefore include the exact bytes left after
    // replacing the current connections stream, in addition to the binding,
    // connections, and temporary ceilings.
    let aggregate_output_limit = replacement_output_limit(package, source.len(), limits)?;
    binding_limits.max_output_bytes = binding_limits
        .max_output_bytes
        .min(limits.max_connections_bytes())
        .min(limits.max_temporary_bytes())
        .min(aggregate_output_limit);
    let scan =
        bindings::scan_connections_with_budget(source.as_slice(), binding_limits, &[], &mut budget)
            .map_err(|error| map_binding_error_with_budget(error, &mut budget))?;
    let mut rewrites = Vec::new();
    rewrites
        .try_reserve(edits.len())
        .map_err(|source| allocation("Custom Data connection rewrites", source))?;
    for (from, to) in edits {
        rewrites.push(UidRewrite {
            from: from.as_str(),
            to: to.as_str(),
        });
    }
    // The binding rewriter performs its exact final-length preflight and
    // fallible reservation after planning replacements. Do not pre-reserve
    // the entire source stream here: a shrinking rewrite must not consume the
    // caller's output/temporary budget for bytes it will discard.
    let mut output = Vec::new();
    let outcome = bindings::rewrite_client_cube_urns_with_budget(
        source.as_slice(),
        &scan,
        &rewrites,
        &mut output,
        binding_limits,
        &mut budget,
    )
    .map_err(map_binding_error)?;
    if !outcome.changed {
        return Ok(());
    }
    if output.len() > limits.max_connections_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data connections.bin bytes",
            actual: output.len(),
            maximum: limits.max_connections_bytes(),
        });
    }
    if output.len() > limits.max_temporary_bytes() {
        return Err(Error::LimitExceeded {
            resource: "Custom Data temporary connections replacement bytes",
            actual: output.len(),
            maximum: limits.max_temporary_bytes(),
        });
    }
    package
        .get_part_mut(&state.part_name)?
        .set_blob_shared(Arc::new(output));
    Ok(())
}
