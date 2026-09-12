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
use thiserror::Error;

use super::{Package, PackageError, body_table_catalog, table_lock};
use crate::selector::BodyTableSelector;

const PREVIEW_ENTRY_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// Finite resources charged by one complete body-table deletion transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableDeletionLimitKind {
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

impl fmt::Display for BodyTableDeletionLimitKind {
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
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableDeletionError {
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
    /// The selected rooted graph or native payload is malformed.
    #[error("the selected Pages body-table deletion source is invalid")]
    InvalidSource,
    /// A surviving object owns a selected formula dependency, so deleting the
    /// table would change another table's calculation semantics.
    #[error("a surviving Pages formula dependency prevents body-table deletion")]
    UnsupportedDependency,
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Pages body-table deletion {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: BodyTableDeletionLimitKind,
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
    pub(super) removed_table: body_table_catalog::BodyTableSnapshot,
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
    pub(super) source_tables: Vec<body_table_catalog::BodyTableSnapshot>,
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
pub struct BodyTableDeletionPatch {
    artifacts: ExactArtifacts,
    proof: table_lock::BodyTableTarget,
    removed_table: body_table_catalog::BodyTableSnapshot,
    source_tables: Arc<[body_table_catalog::BodyTableSnapshot]>,
    target_tables: Arc<[body_table_catalog::BodyTableSnapshot]>,
    source_preview_count: usize,
    target_preview_count: usize,
    touched_components: usize,
    removed_objects: usize,
    removed_components: usize,
    removes_table: bool,
}

impl fmt::Debug for BodyTableDeletionPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableDeletionPatch")
            .field("removed_table", &self.removed_table)
            .field("source_table_count", &self.source_tables.len())
            .field("target_table_count", &self.target_tables.len())
            .finish_non_exhaustive()
    }
}

impl BodyTableDeletionPatch {
    /// Return the semantic table snapshot removed by this patch.
    #[must_use]
    pub const fn removed_table(&self) -> &body_table_catalog::BodyTableSnapshot {
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
pub struct BodyTableDeletionDiagnostics {
    changed: bool,
    touched_components: usize,
    removed_objects: usize,
    removed_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl BodyTableDeletionDiagnostics {
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
pub struct BodyTableDeletionCommit {
    package: Package,
    patch: BodyTableDeletionPatch,
    diagnostics: BodyTableDeletionDiagnostics,
}

impl BodyTableDeletionCommit {
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

    /// Borrow the exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyTableDeletionPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableDeletionDiagnostics {
        &self.diagnostics
    }

    /// Borrow the semantic snapshot associated with this deletion patch.
    ///
    /// For an inverse patch this is the table being restored; the snapshot
    /// remains the original deletion's semantic identity in both directions.
    #[must_use]
    pub fn removed_table(&self) -> &body_table_catalog::BodyTableSnapshot {
        self.patch.removed_table()
    }
}

impl Package {
    /// Remove one rooted body table as one exact-source transaction.
    pub fn remove_body_table<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<BodyTableDeletionCommit, BodyTableDeletionError> {
        let source = &self.state.source;
        let mut budget = table_lock::WireBudget::new(source.limits()).map_err(map_lock_error)?;
        budget
            .charge_source_catalog(source)
            .map_err(map_lock_error)?;
        let targets = table_lock::native_body_table_targets_with_budget(self, &mut budget)
            .map_err(map_lock_error)?;
        let selector = selector.into();
        let target = select_target(targets, selector, &mut budget)?;
        if !source.source_is_exact() {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        let removed_table = body_table_catalog::BodyTableSnapshot::from_target(target.clone())
            .map_err(map_catalog_error)?;
        let request = DeletionRequest {
            target,
            removed_table,
        };
        publish_deletion(self, request, &mut budget)
    }

    /// Apply a deletion patch only to its exact immutable source package.
    pub fn apply_body_table_deletion(
        &self,
        patch: &BodyTableDeletionPatch,
    ) -> Result<BodyTableDeletionCommit, BodyTableDeletionError> {
        let source = &self.state.source;
        let mut budget = table_lock::WireBudget::new(source.limits()).map_err(map_lock_error)?;
        let source_bytes = source.source_bytes();
        budget
            .charge_input_source(source_bytes)
            .and_then(|_| budget.charge_payload_work(source_bytes.len()))
            .map_err(map_lock_error)?;
        let source_owner = source.shared_source();
        if !patch.artifacts.authorizes_source(&source_owner) {
            return Err(BodyTableDeletionError::PatchConflict);
        }
        if !source.source_is_exact() {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        if patch.removes_table {
            let current = self
                .resolve_body_table_with_budget(
                    BodyTableSelector::index(patch.proof.table_position),
                    &mut budget,
                )
                .map_err(map_lock_error)?;
            if current != patch.proof
                || body_table_catalog::BodyTableSnapshot::from_target(current)
                    .map_err(map_catalog_error)?
                    != patch.removed_table
            {
                return Err(BodyTableDeletionError::PatchConflict);
            }
        } else {
            let current = catalog_with_budget(self, &mut budget)?;
            if !catalog_matches_snapshots(&current, &patch.source_tables) {
                return Err(BodyTableDeletionError::PatchConflict);
            }
        }
        let target = patch.artifacts.target();
        let candidate = reopen_candidate(self, target, &mut budget)?;
        verify_candidate(&candidate, patch, &mut budget)?;
        Ok(BodyTableDeletionCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableDeletionDiagnostics {
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
) -> Result<BodyTableDeletionCommit, BodyTableDeletionError> {
    let graph = graph::prepare(source, request, budget)?;
    let mut formula = formula::prepare(source, &graph, budget)?;
    graph::complete_removals(source, &graph, &mut formula, budget)?;
    let metadata = metadata::prepare(source, &graph, &formula, budget)?;
    let mut archive = archive::prepare(source, &graph, &formula, &metadata, budget)?;
    let source_preview_count = preview_count(source, budget)?;
    append_preview_deletions(source, &mut archive.deleted_entries, budget)?;

    let mut edits = Vec::new();
    edits.try_reserve_exact(archive.edits.len()).map_err(|_| {
        BodyTableDeletionError::Allocation {
            amount: archive.edits.len(),
        }
    })?;
    for edit in &archive.edits {
        edits.push(EntryEdit::new(edit.name.as_ref(), &edit.data));
    }
    let mut deleted_names = Vec::new();
    deleted_names
        .try_reserve_exact(archive.deleted_entries.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
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
    let source_tables: Arc<[body_table_catalog::BodyTableSnapshot]> = graph.source_tables.into();
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
        return Err(BodyTableDeletionError::Verification);
    }
    let target_preview_count = preview_count(&candidate, budget)?;
    if target_preview_count != 0 {
        return Err(BodyTableDeletionError::Verification);
    }
    verify_locality(source, &candidate, &archive, budget)?;
    let target_tables: Arc<[body_table_catalog::BodyTableSnapshot]> =
        target_tables.into_values().into();
    let source_bytes = source.state.source.shared_source();
    let artifacts = ExactArtifacts::new(source_bytes, target);
    let patch = BodyTableDeletionPatch {
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
    Ok(BodyTableDeletionCommit {
        package: candidate,
        patch,
        diagnostics: BodyTableDeletionDiagnostics::published(&archive, source_preview_count),
    })
}

fn reopen_candidate(
    source: &Package,
    bytes: Arc<[u8]>,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableDeletionError> {
    budget
        .charge_payload_work(bytes.len())
        .map_err(map_lock_error)?;
    let catalog = litchi_iwa_archive::SourceCatalog::from_shared_bytes_with_limits(
        bytes,
        source.state.source.limits(),
    )
    .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .and_then(|_| table_lock::charge_reopen_work(&catalog, budget))
        .map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn verify_candidate(
    candidate: &Package,
    patch: &BodyTableDeletionPatch,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
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
        return Err(BodyTableDeletionError::Verification);
    }
    if preview_count(candidate, budget)? != patch.target_preview_count {
        return Err(BodyTableDeletionError::Verification);
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
) -> Result<(), BodyTableDeletionError> {
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
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        for object in &component.archive().objects {
            budget
                .charge_payload_work(
                    object
                        .messages
                        .len()
                        .checked_add(1)
                        .ok_or(BodyTableDeletionError::InvalidSource)?,
                )
                .map_err(map_lock_error)?;
            message_count = message_count
                .checked_add(object.messages.len())
                .ok_or(BodyTableDeletionError::InvalidSource)?;
            for info in &object.archive_info.message_infos {
                budget.charge_payload_work(1).map_err(map_lock_error)?;
                field_count = field_count
                    .checked_add(info.field_infos.len())
                    .ok_or(BodyTableDeletionError::InvalidSource)?;
                object_reference_count = object_reference_count
                    .checked_add(info.object_references.len())
                    .ok_or(BodyTableDeletionError::InvalidSource)?;
                reference_count = reference_count
                    .checked_add(info.object_references.len())
                    .and_then(|value| value.checked_add(info.data_references.len()))
                    .ok_or(BodyTableDeletionError::InvalidSource)?;
                budget
                    .charge_payload_work(
                        info.object_references
                            .len()
                            .checked_add(info.data_references.len())
                            .ok_or(BodyTableDeletionError::InvalidSource)?,
                    )
                    .map_err(map_lock_error)?;
                for field in &info.field_infos {
                    budget.charge_payload_work(1).map_err(map_lock_error)?;
                    object_reference_count = object_reference_count
                        .checked_add(field.object_references.len())
                        .ok_or(BodyTableDeletionError::InvalidSource)?;
                    reference_count = reference_count
                        .checked_add(field.object_references.len())
                        .and_then(|value| value.checked_add(field.data_references.len()))
                        .ok_or(BodyTableDeletionError::InvalidSource)?;
                    budget
                        .charge_payload_work(
                            field
                                .object_references
                                .len()
                                .checked_add(field.data_references.len())
                                .ok_or(BodyTableDeletionError::InvalidSource)?,
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
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        remaining >>= 1;
    }
    let searches = object_count
        .checked_add(object_reference_count)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(
            searches
                .checked_mul(
                    binary_search_steps
                        .checked_add(1)
                        .ok_or(BodyTableDeletionError::InvalidSource)?,
                )
                .ok_or(BodyTableDeletionError::InvalidSource)?,
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
                return Err(BodyTableDeletionError::Verification);
            }
            for info in &object.archive_info.message_infos {
                if info
                    .object_references
                    .iter()
                    .filter_map(|identifier| NonZeroU64::new(*identifier))
                    .any(|identifier| removed_ids.binary_search(&identifier).is_ok())
                {
                    return Err(BodyTableDeletionError::Verification);
                }
                for field in &info.field_infos {
                    if field
                        .object_references
                        .iter()
                        .filter_map(|identifier| NonZeroU64::new(*identifier))
                        .any(|identifier| removed_ids.binary_search(&identifier).is_ok())
                    {
                        return Err(BodyTableDeletionError::Verification);
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
) -> Result<body_table_catalog::BodyTableCatalog, BodyTableDeletionError> {
    budget
        .charge_source_catalog(&package.state.source)
        .map_err(map_lock_error)?;
    let targets =
        table_lock::body_table_catalog_with_budget(package, budget).map_err(map_lock_error)?;
    body_table_catalog::BodyTableCatalog::from_targets(targets, budget).map_err(map_catalog_error)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    plan: &ArchivePlan,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let before = source.state.source.package();
    let after = candidate.state.source.package();
    if before.len().saturating_sub(plan.deleted_entries.len()) != after.len() {
        return Err(BodyTableDeletionError::Verification);
    }
    let locality_comparisons = before
        .len()
        .checked_mul(
            plan.deleted_entries
                .len()
                .checked_add(plan.edits.len())
                .ok_or(BodyTableDeletionError::InvalidSource)?,
        )
        .and_then(|value| value.checked_add(after.len()))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
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
            return Err(BodyTableDeletionError::Verification);
        }
        if !edited && before_entry.data() != after_entry.data() {
            return Err(BodyTableDeletionError::Verification);
        }
    }
    Ok(())
}

fn preview_count(
    package: &Package,
    budget: &mut table_lock::WireBudget,
) -> Result<usize, BodyTableDeletionError> {
    budget
        .charge_payload_work(
            package
                .state
                .source
                .package()
                .len()
                .checked_mul(PREVIEW_ENTRY_NAMES.len())
                .ok_or(BodyTableDeletionError::InvalidSource)?,
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
) -> Result<table_lock::BodyTableTarget, BodyTableDeletionError> {
    match selector {
        BodyTableSelector::Position(position) => {
            budget
                .charge_payload_work(position.get().saturating_add(1).min(targets.len()))
                .map_err(map_lock_error)?;
            targets
                .into_iter()
                .nth(position.get())
                .ok_or(BodyTableDeletionError::TableNotFound)
        },
        BodyTableSelector::Name(name) => {
            let mut selected = None;
            for target in targets {
                let comparison_work = target
                    .table_name
                    .len()
                    .checked_add(name.len())
                    .ok_or(BodyTableDeletionError::InvalidSource)?;
                budget
                    .charge_payload_work(comparison_work)
                    .map_err(map_lock_error)?;
                if target.table_name.as_ref() != name {
                    continue;
                }
                if selected.is_some() {
                    return Err(BodyTableDeletionError::AmbiguousTableName);
                }
                selected = Some(target);
            }
            selected.ok_or(BodyTableDeletionError::TableNotFound)
        },
    }
}

fn catalog_matches_snapshots(
    catalog: &body_table_catalog::BodyTableCatalog,
    expected: &[body_table_catalog::BodyTableSnapshot],
) -> bool {
    catalog.len() == expected.len()
        && catalog
            .iter()
            .zip(expected)
            .all(|(actual, expected)| same_table_snapshot(actual, expected))
}

fn catalog_matches_deleted_source(
    after: &body_table_catalog::BodyTableCatalog,
    before: &[body_table_catalog::BodyTableSnapshot],
    removed_index: usize,
) -> bool {
    snapshots_match_deleted_source(after.as_slice(), before, removed_index)
}

fn snapshots_match_deleted_source(
    after: &[body_table_catalog::BodyTableSnapshot],
    before: &[body_table_catalog::BodyTableSnapshot],
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
    left: &body_table_catalog::BodyTableSnapshot,
    right: &body_table_catalog::BodyTableSnapshot,
) -> bool {
    left.name() == right.name() && left.rows() == right.rows() && left.columns() == right.columns()
}

fn append_preview_deletions(
    source: &Package,
    deleted_entries: &mut Vec<Box<str>>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let package_len = source.state.source.package().len();
    let deleted_len = deleted_entries.len();
    let comparisons = package_len
        .checked_add(deleted_len)
        .and_then(|value| value.checked_mul(PREVIEW_ENTRY_NAMES.len()))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_items(PREVIEW_ENTRY_NAMES.len())
        .and_then(|_| budget.charge_payload_work(comparisons))
        .map_err(map_lock_error)?;
    deleted_entries
        .try_reserve(PREVIEW_ENTRY_NAMES.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
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

fn map_catalog_error(error: body_table_catalog::BodyTableCatalogError) -> BodyTableDeletionError {
    match error {
        body_table_catalog::BodyTableCatalogError::AmbiguousTableName => {
            BodyTableDeletionError::AmbiguousTableName
        },
        body_table_catalog::BodyTableCatalogError::InvalidSource => {
            BodyTableDeletionError::InvalidSource
        },
        body_table_catalog::BodyTableCatalogError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableDeletionError::LimitExceeded {
            kind: map_catalog_limit(kind),
            observed,
            maximum,
        },
        body_table_catalog::BodyTableCatalogError::Allocation { amount } => {
            BodyTableDeletionError::Allocation { amount }
        },
        body_table_catalog::BodyTableCatalogError::InvalidName(_) => {
            BodyTableDeletionError::InvalidSource
        },
    }
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableDeletionError {
    match error {
        table_lock::BodyTableLockError::TableNotFound => BodyTableDeletionError::TableNotFound,
        table_lock::BodyTableLockError::AmbiguousTableName => {
            BodyTableDeletionError::AmbiguousTableName
        },
        table_lock::BodyTableLockError::AmbiguousSelector => {
            BodyTableDeletionError::AmbiguousSelector
        },
        table_lock::BodyTableLockError::UnsupportedSource => {
            BodyTableDeletionError::UnsupportedSource
        },
        table_lock::BodyTableLockError::InvalidSource => BodyTableDeletionError::InvalidSource,
        table_lock::BodyTableLockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableDeletionError::LimitExceeded {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::BodyTableLockError::Allocation { amount } => {
            BodyTableDeletionError::Allocation { amount }
        },
        table_lock::BodyTableLockError::Verification => BodyTableDeletionError::Verification,
        table_lock::BodyTableLockError::PatchConflict => BodyTableDeletionError::PatchConflict,
    }
}

const fn map_lock_limit(kind: table_lock::BodyTableLockLimitKind) -> BodyTableDeletionLimitKind {
    use table_lock::BodyTableLockLimitKind as Lock;
    match kind {
        Lock::InputBytes => BodyTableDeletionLimitKind::InputBytes,
        Lock::OutputBytes => BodyTableDeletionLimitKind::OutputBytes,
        Lock::Entries => BodyTableDeletionLimitKind::Entries,
        Lock::EntryBytes => BodyTableDeletionLimitKind::EntryBytes,
        Lock::TotalEntryBytes => BodyTableDeletionLimitKind::TotalEntryBytes,
        Lock::PackageBytes => BodyTableDeletionLimitKind::PackageBytes,
        Lock::PayloadBytes => BodyTableDeletionLimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => BodyTableDeletionLimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => BodyTableDeletionLimitKind::PayloadObjects,
        Lock::PayloadMessages => BodyTableDeletionLimitKind::PayloadMessages,
        Lock::PayloadItems => BodyTableDeletionLimitKind::PayloadItems,
        Lock::PayloadReferences => BodyTableDeletionLimitKind::PayloadReferences,
        Lock::WireBytes => BodyTableDeletionLimitKind::WireBytes,
        Lock::WireFields => BodyTableDeletionLimitKind::WireFields,
        Lock::WireNesting => BodyTableDeletionLimitKind::WireNesting,
        Lock::WireWork => BodyTableDeletionLimitKind::WireWork,
    }
}

const fn map_catalog_limit(
    kind: body_table_catalog::BodyTableCatalogLimitKind,
) -> BodyTableDeletionLimitKind {
    use body_table_catalog::BodyTableCatalogLimitKind as Catalog;
    match kind {
        Catalog::InputBytes => BodyTableDeletionLimitKind::InputBytes,
        Catalog::OutputBytes => BodyTableDeletionLimitKind::OutputBytes,
        Catalog::Entries => BodyTableDeletionLimitKind::Entries,
        Catalog::EntryBytes => BodyTableDeletionLimitKind::EntryBytes,
        Catalog::TotalEntryBytes => BodyTableDeletionLimitKind::TotalEntryBytes,
        Catalog::PackageBytes => BodyTableDeletionLimitKind::PackageBytes,
        Catalog::PayloadBytes => BodyTableDeletionLimitKind::PayloadBytes,
        Catalog::TotalPayloadBytes => BodyTableDeletionLimitKind::TotalPayloadBytes,
        Catalog::PayloadObjects => BodyTableDeletionLimitKind::PayloadObjects,
        Catalog::PayloadMessages => BodyTableDeletionLimitKind::PayloadMessages,
        Catalog::PayloadItems => BodyTableDeletionLimitKind::PayloadItems,
        Catalog::PayloadReferences => BodyTableDeletionLimitKind::PayloadReferences,
        Catalog::WireBytes => BodyTableDeletionLimitKind::WireBytes,
        Catalog::WireFields => BodyTableDeletionLimitKind::WireFields,
        Catalog::WireNesting => BodyTableDeletionLimitKind::WireNesting,
        Catalog::WireWork => BodyTableDeletionLimitKind::WireWork,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyTableDeletionError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableDeletionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => BodyTableDeletionLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    BodyTableDeletionLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => BodyTableDeletionLimitKind::Entries,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => {
                    BodyTableDeletionLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    BodyTableDeletionLimitKind::TotalEntryBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    BodyTableDeletionLimitKind::PayloadBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    BodyTableDeletionLimitKind::TotalPayloadBytes
                },
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    BodyTableDeletionLimitKind::PackageBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyTableDeletionError::Allocation { amount }
        },
        _ => BodyTableDeletionError::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> BodyTableDeletionError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => BodyTableDeletionError::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => BodyTableDeletionError::LimitExceeded {
            kind: BodyTableDeletionLimitKind::PayloadObjects,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::PayloadLimit { observed, limit } => BodyTableDeletionError::LimitExceeded {
            kind: BodyTableDeletionLimitKind::PayloadBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        _ => BodyTableDeletionError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
