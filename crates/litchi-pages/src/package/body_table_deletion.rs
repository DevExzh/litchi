//! Exact-source body-table deletion for focused Pages packages.
//!
//! Deletion is coordinated here because a body table owns records in several
//! native components: the body storage anchor, drawable/model objects,
//! formula dependency state, package metadata registries, and sometimes
//! component registrations.  The child modules only prepare their bounded
//! part of one candidate.  This module owns selector resolution, the shared
//! transaction ledger, one physical reassembly, candidate reopening, and the
//! exact reversible patch.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the deletion transaction keeps public evidence beside its private plan"
)]

mod archive;
mod formula;
mod graph;
mod metadata;
mod text;

use std::collections::HashMap;
use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use thiserror::Error as ThisError;

use super::{Package, PackageError, body_table_catalog, table_lock};
use crate::selector::BodyTableSelector;

const PREVIEW_ENTRY_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// Finite resources charged by one complete body-table deletion transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source package bytes.
    InputBytes,
    /// Complete candidate package bytes.
    OutputBytes,
    /// Physical ZIP entries.
    Entries,
    /// Bytes in one physical ZIP entry.
    EntryBytes,
    /// Aggregate physical ZIP entry bytes.
    TotalEntryBytes,
    /// ZIP names and structural metadata bytes.
    PackageBytes,
    /// Bytes in one decoded native IWA payload.
    PayloadBytes,
    /// Aggregate decoded native IWA payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected by the transaction.
    PayloadObjects,
    /// Native payload messages inspected by the transaction.
    PayloadMessages,
    /// Native framing and metadata items inspected by the transaction.
    PayloadItems,
    /// Native object/data references inspected by the transaction.
    PayloadReferences,
    /// Strict input wire bytes inspected by a focused codec.
    WireBytes,
    /// Strict output wire bytes produced by a focused codec.
    WireOutputBytes,
    /// Strict wire fields inspected by a focused codec.
    WireFields,
    /// Strict wire nesting depth.
    WireNesting,
    /// Aggregate bounded codec work.
    WireWork,
    /// Aggregate work across all deletion phases, including reopen and
    /// locality verification.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::TransactionWork => "transaction work",
        })
    }
}

/// Failure from resolving, preparing, or publishing one body-table deletion.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The selected rooted graph or native payload is malformed.
    #[error("the selected Pages body-table deletion source is invalid")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Pages body-table deletion {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category that exceeded its ceiling.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Pages body-table deletion")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested semantic state.
    #[error("the deleted Pages body table failed semantic verification")]
    Verification,
    /// The supplied patch was created from a different exact source artifact.
    #[error("the Pages body-table deletion patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// No rooted body table matched the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// More than one rooted body table matched an exact name.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The rooted graph or selector was not unique.
    #[error("the Pages body-table deletion selector is ambiguous")]
    AmbiguousSelector,
    /// The source cannot publish a preservation-safe changed artifact.
    #[error("the Pages package source does not support exact body-table deletion")]
    UnsupportedSource,
    /// A surviving object owns a selected formula dependency, so deleting the
    /// table would change another table's calculation semantics.
    #[error("a surviving Pages formula dependency prevents body-table deletion")]
    UnsupportedDependency,
    /// The edit already stages its one removal.
    #[error("the Pages body-table deletion edit already stages a removal")]
    OperationAlreadyStaged,
    /// The edit was committed without a staged removal.
    #[error("the Pages body-table deletion edit stages no removal")]
    NoStagedOperation,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::AmbiguousTableName { .. } | Self::AmbiguousSelector { .. } => {
                RefusalKind::InvalidValue
            },
            Self::UnsupportedSource { .. } | Self::UnsupportedDependency { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::OperationAlreadyStaged { .. } | Self::NoStagedOperation { .. } => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One object location retained by the deletion-wide object index.
///
/// The index is built once by `graph` and borrowed by formula, metadata, and
/// archive preparation.  Keeping the component name with the location avoids
/// a repeated package-wide name lookup while retaining the source order used
/// by locality verification.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct DeletionObjectLocation {
    pub(super) identifier: NonZeroU64,
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) component_name: Box<str>,
}

/// One source-wide object index shared by every deletion phase.
#[derive(Debug, Clone, Default)]
pub(super) struct DeletionObjectIndex {
    /// Map an identifier to an index in `ordered`, so component names and
    /// locations have one owned allocation per source object.
    pub(super) by_identifier: HashMap<NonZeroU64, usize>,
    pub(super) ordered: Vec<DeletionObjectLocation>,
}

impl DeletionObjectIndex {
    pub(super) fn get(&self, identifier: NonZeroU64) -> Option<&DeletionObjectLocation> {
        self.by_identifier
            .get(&identifier)
            .and_then(|index| self.ordered.get(*index))
    }

    pub(super) fn iter(&self) -> impl Iterator<Item = &DeletionObjectLocation> {
        self.ordered.iter()
    }
}

/// The selector result captured before any deletion phase allocates a
/// candidate.  The public snapshot is the semantic identity used by the
/// commit and inverse patch; `target` is private ownership evidence.
#[derive(Clone)]
pub(super) struct DeletionRequest {
    pub(super) target: table_lock::BodyTableTarget,
    pub(super) removed_table: body_table_catalog::Snapshot,
}

impl fmt::Debug for DeletionRequest {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DeletionRequest")
            .field("removed_table", &self.removed_table)
            .finish_non_exhaustive()
    }
}

/// One payload replacement plus the archive-header references it is allowed
/// to prune while preserving all unrelated raw fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct MessageEdit {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
    pub(super) object_identifier: NonZeroU64,
    pub(super) message_type: u32,
    pub(super) data: Vec<u8>,
    pub(super) remove_object_references: Vec<NonZeroU64>,
    pub(super) remove_data_references: Vec<NonZeroU64>,
}

/// One object selected for physical archive removal.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) struct ObjectRemoval {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) identifier: NonZeroU64,
}

/// One selected component-registration removal in package metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct ComponentRemoval {
    pub(super) component_index: usize,
    pub(super) identifier: u64,
    pub(super) locator: Box<str>,
    pub(super) name: Box<str>,
}

/// One exact component UUID-registry removal request.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) struct UuidRemoval {
    pub(super) component_identifier: u64,
    pub(super) object_identifier: NonZeroU64,
    pub(super) expected_uuid: litchi_iwa_protos::package_metadata_codec::UuidBits,
}

/// One exact component-to-component object external-reference removal.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) struct ExternalReferenceRemoval {
    pub(super) source_component_identifier: u64,
    pub(super) target_component_identifier: u64,
    pub(super) object_identifier: NonZeroU64,
    pub(super) expected_is_weak: Option<bool>,
}

/// One exact `ComponentDataReference` owner removal.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) struct DataReferenceOwnerRemoval {
    pub(super) component_identifier: u64,
    pub(super) data_identifier: NonZeroU64,
    pub(super) object_identifier: NonZeroU64,
    pub(super) expected_count: u32,
}

/// Additive output of graph, formula, and metadata preparation.
///
/// Each phase appends only records it proved from the same source snapshot.
/// The archive phase rejects duplicate or conflicting records before mutating
/// an in-memory archive, so the final physical operation remains atomic.
#[derive(Debug, Clone, Default)]
pub(super) struct RemovalPlan {
    pub(super) message_edits: Vec<MessageEdit>,
    pub(super) object_removals: Vec<ObjectRemoval>,
    pub(super) component_removals: Vec<ComponentRemoval>,
    pub(super) uuid_removals: Vec<UuidRemoval>,
    pub(super) external_reference_removals: Vec<ExternalReferenceRemoval>,
    pub(super) data_reference_owner_removals: Vec<DataReferenceOwnerRemoval>,
    pub(super) formula_contexts: Vec<NonZeroU64>,
}

/// Graph phase output.  The index and source table snapshots are retained so
/// later phases and candidate verification never redo the rooted discovery.
#[derive(Debug, Clone)]
pub(super) struct GraphPlan {
    pub(super) request: DeletionRequest,
    pub(super) index: DeletionObjectIndex,
    pub(super) source_tables: Vec<body_table_catalog::Snapshot>,
    pub(super) removals: RemovalPlan,
}

/// Formula phase output, additive to graph ownership removals.
#[derive(Debug, Clone)]
pub(super) struct FormulaPlan {
    pub(super) removals: RemovalPlan,
}

/// Metadata phase output, additive to graph and formula removals.
#[derive(Debug, Clone)]
pub(super) struct MetadataPlan {
    pub(super) removals: RemovalPlan,
}

/// An edited decompressed IWA component prepared by the archive phase.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct ComponentEdit {
    pub(super) component_index: usize,
    pub(super) name: Box<str>,
    pub(super) data: Vec<u8>,
}

/// Final physical worklist.  `body_table_deletion` performs the one ZIP
/// reassembly from these owned edits and names; child modules never publish a
/// partial candidate.
#[derive(Debug, Clone, Default)]
pub(super) struct ArchivePlan {
    pub(super) edits: Vec<ComponentEdit>,
    pub(super) deleted_entries: Vec<Box<str>>,
    /// Object identifiers that the merged graph/formula plan authorizes for
    /// physical removal.  The coordinator uses this final set for a
    /// candidate-wide dangling-reference proof after all rewrites exist.
    pub(super) removed_object_ids: Vec<NonZeroU64>,
    pub(super) removed_objects: usize,
    pub(super) removed_components: usize,
    pub(super) touched_components: usize,
}

/// Exact-source reversible body-table deletion patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    proof: table_lock::BodyTableTarget,
    removed_table: body_table_catalog::Snapshot,
    source_tables: Arc<[body_table_catalog::Snapshot]>,
    target_tables: Arc<[body_table_catalog::Snapshot]>,
    source_preview_count: usize,
    target_preview_count: usize,
    touched_components: usize,
    removed_objects: usize,
    removed_components: usize,
    removes_table: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("removed_table", &self.removed_table)
            .field("source_table_count", &self.source_tables.len())
            .field("target_table_count", &self.target_tables.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the semantic table snapshot removed by this patch.
    #[must_use]
    pub const fn removed_table(&self) -> &body_table_catalog::Snapshot {
        &self.removed_table
    }

    /// Return the source artifact fingerprint used for diagnostics.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target artifact fingerprint used for diagnostics.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return the exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            proof: self.proof.clone(),
            removed_table: self.removed_table.clone(),
            source_tables: Arc::clone(&self.target_tables),
            target_tables: Arc::clone(&self.source_tables),
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
            touched_components: self.touched_components,
            removed_objects: self.removed_objects,
            removed_components: self.removed_components,
            removes_table: !self.removes_table,
        }
    }

    /// Deletion patches always carry a changed exact artifact.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        false
    }
}

/// Compact evidence from one published body-table deletion.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    removed_objects: usize,
    removed_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    fn published(plan: &ArchivePlan, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: plan.touched_components,
            removed_objects: plan.removed_objects,
            removed_components: plan.removed_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Exact package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of physically removed native objects.
    #[must_use]
    pub const fn removed_objects(self) -> usize {
        self.removed_objects
    }

    /// Number of physically removed native components.
    #[must_use]
    pub const fn removed_components(self) -> usize {
        self.removed_components
    }

    /// Number of removed root preview entries.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one body-table deletion transaction.
#[must_use = "a body-table deletion commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Borrow the semantic snapshot associated with this deletion patch.
    ///
    /// For an inverse patch this is the table being restored; the snapshot
    /// remains the original deletion's semantic identity in both directions.
    #[must_use]
    pub fn removed_table(&self) -> &body_table_catalog::Snapshot {
        self.patch.removed_table()
    }
}

/// One immutable edit that stages exactly one rooted body-table removal.
///
/// Start it with [`Package::edit_body_table_deletion`]; the source package is
/// never changed, and dropping the edit discards the staged removal.
pub struct Edit<'a> {
    source: &'a Package,
    selector: Option<BodyTableSelector<'a>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("staged", &self.selector.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    /// Stage the removal of the selected body table.
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] after one removal.
    pub fn remove(mut self, selector: impl Into<BodyTableSelector<'a>>) -> Result<Self, Error> {
        if self.selector.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        self.selector = Some(selector.into());
        Ok(self)
    }

    /// Resolve, publish and reopen the staged removal as one exact-source
    /// transaction.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit,
    /// [`Refusal::TableNotFound`], [`Refusal::AmbiguousTableName`] and
    /// [`Refusal::AmbiguousSelector`] for the selector,
    /// [`Refusal::UnsupportedSource`] for a source that is not exact, and typed
    /// source, dependency, limit and verification errors. Failure never
    /// publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let selector = self
            .selector
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        let source = &self.source.state.source;
        let mut budget = table_lock::WireBudget::new(source.limits()).map_err(map_lock_error)?;
        budget
            .charge_source_catalog(source)
            .map_err(map_lock_error)?;
        let targets = table_lock::native_body_table_targets_with_budget(self.source, &mut budget)
            .map_err(map_lock_error)?;
        let target = select_target(targets, selector, &mut budget)?;
        if !source.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        let removed_table =
            body_table_catalog::Snapshot::from_target(target.clone()).map_err(map_catalog_error)?;
        let request = DeletionRequest {
            target,
            removed_table,
        };
        publish_deletion(self.source, request, &mut budget)
    }
}

impl Package {
    /// Start an immutable single-removal body-table deletion edit.
    ///
    /// ```no_run
    /// use litchi_pages::Package;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("report.pages")?;
    /// let commit = package.edit_body_table_deletion().remove(0usize)?.commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_body_table_deletion(&commit.patch().inverse())?;
    /// assert_eq!(restored.package().body_tables()?.len(), package.body_tables()?.len());
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub const fn edit_body_table_deletion(&self) -> Edit<'_> {
        Edit {
            source: self,
            selector: None,
        }
    }

    /// Apply a deletion patch only to its exact immutable source package.
    pub fn apply_body_table_deletion(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = &self.state.source;
        let mut budget = table_lock::WireBudget::new(source.limits()).map_err(map_lock_error)?;
        let source_bytes = source.source_bytes();
        budget
            .charge_input_source(source_bytes)
            .and_then(|_| budget.charge_payload_work(source_bytes.len()))
            .map_err(map_lock_error)?;
        let source_owner = source.shared_source();
        if !patch.artifacts.authorizes_source(&source_owner) {
            return Err(Error::PatchConflict);
        }
        if !source.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        if patch.removes_table {
            let current = self
                .resolve_body_table_with_budget(
                    BodyTableSelector::index(patch.proof.table_position),
                    &mut budget,
                )
                .map_err(map_lock_error)?;
            if current != patch.proof
                || body_table_catalog::Snapshot::from_target(current).map_err(map_catalog_error)?
                    != patch.removed_table
            {
                return Err(Error::PatchConflict);
            }
        } else {
            let current = catalog_with_budget(self, &mut budget)?;
            if !catalog_matches_snapshots(&current, &patch.source_tables) {
                return Err(Error::PatchConflict);
            }
        }
        let target = patch.artifacts.target();
        let candidate = reopen_candidate(self, target, &mut budget)?;
        verify_candidate(&candidate, patch, &mut budget)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components: patch.touched_components,
                removed_objects: patch.removed_objects,
                removed_components: patch.removed_components,
                deleted_previews: if patch.removes_table {
                    patch
                        .source_preview_count
                        .saturating_sub(patch.target_preview_count)
                } else {
                    patch
                        .target_preview_count
                        .saturating_sub(patch.source_preview_count)
                },
                full_reparse_performed: true,
            },
        })
    }
}

fn publish_deletion(
    source: &Package,
    request: DeletionRequest,
    budget: &mut table_lock::WireBudget,
) -> Result<Commit, Error> {
    let graph = graph::prepare(source, request, budget)?;
    let mut formula = formula::prepare(source, &graph, budget)?;
    graph::complete_removals(source, &graph, &mut formula, budget)?;
    let metadata = metadata::prepare(source, &graph, &formula, budget)?;
    let mut archive = archive::prepare(source, &graph, &formula, &metadata, budget)?;
    let source_preview_count = preview_count(source, budget)?;
    append_preview_deletions(source, &mut archive.deleted_entries, budget)?;

    let mut edits = Vec::new();
    edits
        .try_reserve_exact(archive.edits.len())
        .map_err(|_| Error::Allocation {
            amount: archive.edits.len(),
        })?;
    for edit in &archive.edits {
        edits.push(EntryEdit::new(edit.name.as_ref(), &edit.data));
    }
    let mut deleted_names = Vec::new();
    deleted_names
        .try_reserve_exact(archive.deleted_entries.len())
        .map_err(|_| Error::Allocation {
            amount: archive.deleted_entries.len(),
        })?;
    for entry in &archive.deleted_entries {
        deleted_names.push(entry.as_ref());
    }
    let prepared = source
        .state
        .source
        .package()
        .prepare_reassembly_with_deletions(&edits, &deleted_names, source.state.source.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .charge_output_bytes(requirements.output_bytes())
        .and_then(|_| budget.charge_payload_work(requirements.output_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.retained_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.scratch_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.allocations()))
        .map_err(map_lock_error)?;
    let target_bytes = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let target: Arc<[u8]> = target_bytes.into();
    let source_table_count = graph.source_tables.len();
    let source_tables: Arc<[body_table_catalog::Snapshot]> = graph.source_tables.into();
    let candidate = reopen_candidate(source, Arc::clone(&target), budget)?;
    verify_removed_object_references(&candidate, &archive.removed_object_ids, budget)?;
    let target_tables = catalog_with_budget(&candidate, budget)?;
    let target_table_count = target_tables.len();
    if target_table_count.checked_add(1) != Some(source_table_count)
        || !catalog_matches_deleted_source(
            &target_tables,
            &source_tables,
            graph.request.removed_table.index(),
        )
    {
        return Err(Error::Verification);
    }
    let target_preview_count = preview_count(&candidate, budget)?;
    if target_preview_count != 0 {
        return Err(Error::Verification);
    }
    verify_locality(source, &candidate, &archive, budget)?;
    let target_tables: Arc<[body_table_catalog::Snapshot]> = target_tables.into_values().into();
    let source_bytes = source.state.source.shared_source();
    let artifacts = ExactArtifacts::new(source_bytes, target);
    let patch = Patch {
        artifacts,
        proof: graph.request.target,
        removed_table: graph.request.removed_table,
        source_tables,
        target_tables,
        source_preview_count,
        target_preview_count,
        touched_components: archive.touched_components,
        removed_objects: archive.removed_objects,
        removed_components: archive.removed_components,
        removes_table: true,
    };
    Ok(Commit {
        package: candidate,
        patch,
        diagnostics: Diagnostics::published(&archive, source_preview_count),
    })
}

fn reopen_candidate(
    source: &Package,
    bytes: Arc<[u8]>,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, Error> {
    budget
        .charge_payload_work(bytes.len())
        .map_err(map_lock_error)?;
    let catalog = source
        .state
        .source
        .reopen(bytes)
        .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .and_then(|_| table_lock::charge_reopen_work(&catalog, budget))
        .map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn verify_candidate(
    candidate: &Package,
    patch: &Patch,
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
    let target_tables = catalog_with_budget(candidate, budget)?;
    if !catalog_matches_snapshots(&target_tables, &patch.target_tables)
        || (patch.removes_table
            && !catalog_matches_deleted_source(
                &target_tables,
                &patch.source_tables,
                patch.removed_table.index(),
            ))
        || (!patch.removes_table
            && !snapshots_match_deleted_source(
                &patch.source_tables,
                &patch.target_tables,
                patch.removed_table.index(),
            ))
    {
        return Err(Error::Verification);
    }
    if preview_count(candidate, budget)? != patch.target_preview_count {
        return Err(Error::Verification);
    }
    Ok(())
}

/// Prove that the final candidate contains neither a removed object nor a
/// surviving aggregate/field object reference to one.  Data-reference IDs
/// are deliberately traversed only for accounting: they occupy a separate
/// namespace and must not be compared with object IDs.
fn verify_removed_object_references(
    candidate: &Package,
    removed_ids: &[NonZeroU64],
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
    if removed_ids.is_empty() {
        return Ok(());
    }

    let mut object_count = 0usize;
    let mut message_count = 0usize;
    let mut field_count = 0usize;
    let mut reference_count = 0usize;
    let mut object_reference_count = 0usize;
    for component in candidate.state.source.components().iter() {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        object_count = object_count
            .checked_add(component.archive().objects.len())
            .ok_or(Error::InvalidSource)?;
        for object in &component.archive().objects {
            budget
                .charge_payload_work(
                    object
                        .messages
                        .len()
                        .checked_add(1)
                        .ok_or(Error::InvalidSource)?,
                )
                .map_err(map_lock_error)?;
            message_count = message_count
                .checked_add(object.messages.len())
                .ok_or(Error::InvalidSource)?;
            for info in &object.archive_info.message_infos {
                budget.charge_payload_work(1).map_err(map_lock_error)?;
                field_count = field_count
                    .checked_add(info.field_infos.len())
                    .ok_or(Error::InvalidSource)?;
                object_reference_count = object_reference_count
                    .checked_add(info.object_references.len())
                    .ok_or(Error::InvalidSource)?;
                reference_count = reference_count
                    .checked_add(info.object_references.len())
                    .and_then(|value| value.checked_add(info.data_references.len()))
                    .ok_or(Error::InvalidSource)?;
                budget
                    .charge_payload_work(
                        info.object_references
                            .len()
                            .checked_add(info.data_references.len())
                            .ok_or(Error::InvalidSource)?,
                    )
                    .map_err(map_lock_error)?;
                for field in &info.field_infos {
                    budget.charge_payload_work(1).map_err(map_lock_error)?;
                    object_reference_count = object_reference_count
                        .checked_add(field.object_references.len())
                        .ok_or(Error::InvalidSource)?;
                    reference_count = reference_count
                        .checked_add(field.object_references.len())
                        .and_then(|value| value.checked_add(field.data_references.len()))
                        .ok_or(Error::InvalidSource)?;
                    budget
                        .charge_payload_work(
                            field
                                .object_references
                                .len()
                                .checked_add(field.data_references.len())
                                .ok_or(Error::InvalidSource)?,
                        )
                        .map_err(map_lock_error)?;
                }
            }
        }
    }
    budget
        .charge_payload_objects(object_count)
        .and_then(|_| budget.charge_payload_messages(message_count))
        .and_then(|_| budget.charge_payload_items(field_count))
        .and_then(|_| budget.charge_payload_references(reference_count))
        .map_err(map_lock_error)?;

    let mut binary_search_steps = 0usize;
    let mut remaining = removed_ids.len();
    while remaining > 0 {
        binary_search_steps = binary_search_steps
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        remaining >>= 1;
    }
    let searches = object_count
        .checked_add(object_reference_count)
        .ok_or(Error::InvalidSource)?;
    budget
        .charge_payload_work(
            searches
                .checked_mul(
                    binary_search_steps
                        .checked_add(1)
                        .ok_or(Error::InvalidSource)?,
                )
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(map_lock_error)?;

    for component in candidate.state.source.components().iter() {
        for object in &component.archive().objects {
            if object
                .archive_info
                .identifier
                .and_then(NonZeroU64::new)
                .is_some_and(|identifier| removed_ids.binary_search(&identifier).is_ok())
            {
                return Err(Error::Verification);
            }
            for info in &object.archive_info.message_infos {
                if info
                    .object_references
                    .iter()
                    .filter_map(|identifier| NonZeroU64::new(*identifier))
                    .any(|identifier| removed_ids.binary_search(&identifier).is_ok())
                {
                    return Err(Error::Verification);
                }
                for field in &info.field_infos {
                    if field
                        .object_references
                        .iter()
                        .filter_map(|identifier| NonZeroU64::new(*identifier))
                        .any(|identifier| removed_ids.binary_search(&identifier).is_ok())
                    {
                        return Err(Error::Verification);
                    }
                }
            }
        }
    }
    Ok(())
}

fn catalog_with_budget(
    package: &Package,
    budget: &mut table_lock::WireBudget,
) -> Result<body_table_catalog::Catalog, Error> {
    budget
        .charge_source_catalog(&package.state.source)
        .map_err(map_lock_error)?;
    let targets =
        table_lock::body_table_catalog_with_budget(package, budget).map_err(map_lock_error)?;
    body_table_catalog::Catalog::from_targets(targets, budget).map_err(map_catalog_error)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    plan: &ArchivePlan,
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
    let before = source.state.source.package();
    let after = candidate.state.source.package();
    if before.len().saturating_sub(plan.deleted_entries.len()) != after.len() {
        return Err(Error::Verification);
    }
    let locality_comparisons = before
        .len()
        .checked_mul(
            plan.deleted_entries
                .len()
                .checked_add(plan.edits.len())
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|value| value.checked_add(after.len()))
        .ok_or(Error::InvalidSource)?;
    budget
        .charge_payload_work(locality_comparisons)
        .map_err(map_lock_error)?;
    for (before_entry, after_entry) in before
        .iter()
        .filter(|entry| {
            !plan
                .deleted_entries
                .iter()
                .any(|name| name.as_ref() == entry.name())
        })
        .zip(after.iter())
    {
        let edited = plan
            .edits
            .iter()
            .any(|edit| edit.name.as_ref() == before_entry.name());
        if before_entry.name() != after_entry.name() {
            return Err(Error::Verification);
        }
        if !edited && before_entry.data() != after_entry.data() {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn preview_count(package: &Package, budget: &mut table_lock::WireBudget) -> Result<usize, Error> {
    budget
        .charge_payload_work(
            package
                .state
                .source
                .package()
                .len()
                .checked_mul(PREVIEW_ENTRY_NAMES.len())
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(map_lock_error)?;
    Ok(PREVIEW_ENTRY_NAMES
        .iter()
        .filter(|name| {
            package
                .state
                .source
                .package()
                .iter()
                .any(|entry| entry.name() == **name)
        })
        .count())
}

fn select_target(
    targets: Vec<table_lock::BodyTableTarget>,
    selector: BodyTableSelector<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<table_lock::BodyTableTarget, Error> {
    match selector {
        BodyTableSelector::Position(position) => {
            budget
                .charge_payload_work(position.get().saturating_add(1).min(targets.len()))
                .map_err(map_lock_error)?;
            targets
                .into_iter()
                .nth(position.get())
                .ok_or(Error::Refused(Refusal::TableNotFound))
        },
        BodyTableSelector::Name(name) => {
            let mut selected = None;
            for target in targets {
                let comparison_work = target
                    .table_name
                    .len()
                    .checked_add(name.len())
                    .ok_or(Error::InvalidSource)?;
                budget
                    .charge_payload_work(comparison_work)
                    .map_err(map_lock_error)?;
                if target.table_name.as_ref() != name {
                    continue;
                }
                if selected.is_some() {
                    return Err(Error::Refused(Refusal::AmbiguousTableName));
                }
                selected = Some(target);
            }
            selected.ok_or(Error::Refused(Refusal::TableNotFound))
        },
    }
}

fn catalog_matches_snapshots(
    catalog: &body_table_catalog::Catalog,
    expected: &[body_table_catalog::Snapshot],
) -> bool {
    catalog.len() == expected.len()
        && catalog
            .iter()
            .zip(expected)
            .all(|(actual, expected)| same_table_snapshot(actual, expected))
}

fn catalog_matches_deleted_source(
    after: &body_table_catalog::Catalog,
    before: &[body_table_catalog::Snapshot],
    removed_index: usize,
) -> bool {
    snapshots_match_deleted_source(after.as_slice(), before, removed_index)
}

fn snapshots_match_deleted_source(
    after: &[body_table_catalog::Snapshot],
    before: &[body_table_catalog::Snapshot],
    removed_index: usize,
) -> bool {
    if before.len() != after.len().saturating_add(1) || removed_index >= before.len() {
        return false;
    }
    before
        .iter()
        .enumerate()
        .filter(|(index, _)| *index != removed_index)
        .map(|(_, snapshot)| snapshot)
        .zip(after)
        .all(|(before, after)| same_table_snapshot(before, after))
}

fn same_table_snapshot(
    left: &body_table_catalog::Snapshot,
    right: &body_table_catalog::Snapshot,
) -> bool {
    left.name() == right.name() && left.rows() == right.rows() && left.columns() == right.columns()
}

fn append_preview_deletions(
    source: &Package,
    deleted_entries: &mut Vec<Box<str>>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
    let package_len = source.state.source.package().len();
    let deleted_len = deleted_entries.len();
    let comparisons = package_len
        .checked_add(deleted_len)
        .and_then(|value| value.checked_mul(PREVIEW_ENTRY_NAMES.len()))
        .ok_or(Error::InvalidSource)?;
    budget
        .charge_payload_items(PREVIEW_ENTRY_NAMES.len())
        .and_then(|_| budget.charge_payload_work(comparisons))
        .map_err(map_lock_error)?;
    deleted_entries
        .try_reserve(PREVIEW_ENTRY_NAMES.len())
        .map_err(|_| Error::Allocation {
            amount: PREVIEW_ENTRY_NAMES.len(),
        })?;
    for name in PREVIEW_ENTRY_NAMES {
        if source
            .state
            .source
            .package()
            .iter()
            .any(|entry| entry.name() == name)
            && !deleted_entries.iter().any(|entry| entry.as_ref() == name)
        {
            deleted_entries.push(name.into());
        }
    }
    Ok(())
}

fn map_catalog_error(error: body_table_catalog::Error) -> Error {
    match error {
        body_table_catalog::Error::AmbiguousTableName => {
            Error::Refused(Refusal::AmbiguousTableName)
        },
        body_table_catalog::Error::InvalidSource => Error::InvalidSource,
        body_table_catalog::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_catalog_limit(kind),
            observed,
            maximum,
        },
        body_table_catalog::Error::Allocation { amount } => Error::Allocation { amount },
        body_table_catalog::Error::InvalidName(_) => Error::InvalidSource,
    }
}

fn map_lock_error(error: table_lock::Error) -> Error {
    match error {
        table_lock::Error::Refused(crate::table::lock::Refusal::TableNotFound) => {
            Error::Refused(Refusal::TableNotFound)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::AmbiguousTableName) => {
            Error::Refused(Refusal::AmbiguousTableName)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::AmbiguousSelector) => {
            Error::Refused(Refusal::AmbiguousSelector)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        table_lock::Error::InvalidSource => Error::InvalidSource,
        table_lock::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::Error::Allocation { amount } => Error::Allocation { amount },
        table_lock::Error::Verification => Error::Verification,
        table_lock::Error::PatchConflict => Error::PatchConflict,
    }
}

const fn map_lock_limit(kind: table_lock::LimitKind) -> LimitKind {
    use table_lock::LimitKind as Lock;
    match kind {
        Lock::InputBytes => LimitKind::InputBytes,
        Lock::OutputBytes => LimitKind::OutputBytes,
        Lock::Entries => LimitKind::Entries,
        Lock::EntryBytes => LimitKind::EntryBytes,
        Lock::TotalEntryBytes => LimitKind::TotalEntryBytes,
        Lock::PackageBytes => LimitKind::PackageBytes,
        Lock::PayloadBytes => LimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => LimitKind::PayloadObjects,
        Lock::PayloadMessages => LimitKind::PayloadMessages,
        Lock::PayloadItems => LimitKind::PayloadItems,
        Lock::PayloadReferences => LimitKind::PayloadReferences,
        Lock::WireBytes => LimitKind::WireBytes,
        Lock::WireFields => LimitKind::WireFields,
        Lock::WireNesting => LimitKind::WireNesting,
        Lock::WireWork => LimitKind::WireWork,
    }
}

const fn map_catalog_limit(kind: body_table_catalog::LimitKind) -> LimitKind {
    use body_table_catalog::LimitKind as Catalog;
    match kind {
        Catalog::InputBytes => LimitKind::InputBytes,
        Catalog::OutputBytes => LimitKind::OutputBytes,
        Catalog::Entries => LimitKind::Entries,
        Catalog::EntryBytes => LimitKind::EntryBytes,
        Catalog::TotalEntryBytes => LimitKind::TotalEntryBytes,
        Catalog::PackageBytes => LimitKind::PackageBytes,
        Catalog::PayloadBytes => LimitKind::PayloadBytes,
        Catalog::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        Catalog::PayloadObjects => LimitKind::PayloadObjects,
        Catalog::PayloadMessages => LimitKind::PayloadMessages,
        Catalog::PayloadItems => LimitKind::PayloadItems,
        Catalog::PayloadReferences => LimitKind::PayloadReferences,
        Catalog::WireBytes => LimitKind::WireBytes,
        Catalog::WireFields => LimitKind::WireFields,
        Catalog::WireNesting => LimitKind::WireNesting,
        Catalog::WireWork => LimitKind::WireWork,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalPayloadBytes,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::PackageBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadObjects,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::PayloadLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        _ => Error::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
