//! Selector-first, source-preserving Pages body-table appearance transactions.
//!
//! This module owns the rooted Pages body table selection and publication
//! boundary.  The table-style wire projection is deliberately kept behind the
//! hidden protos codec; native object identifiers and generated protobuf
//! values never cross the public `table::appearance` module.

use std::collections::HashSet;
use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::archive::{
    ArchiveReferenceKind, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceVisitor, FieldObjectReferenceTransition, ObjectReferenceTransition,
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::package_metadata_codec::{
    AdditionSaveTokenBatch, Batch as MetadataBatch, ComponentSelector, ExternalReferenceAddition,
    ObjectUuidAddition, PackageMetadataVisitor, RewriteError as MetadataRewriteError,
    RewriteOptions as MetadataRewriteOptions, SaveTokenBatch, UuidBits,
    inspect_package_metadata_with_visitor, prepare_package_metadata_additions_and_save_tokens,
};
use litchi_iwa_protos::table_appearance_codec as codec;
use thiserror::Error as ThisError;

use super::{Package, table_lock};
use crate::{selector::BodyTableSelector, table::appearance::Appearance};

const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const MAX_INHERITANCE_DEPTH: usize = 64;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const METADATA_ENTRY_NAME: &str = "Index/Metadata.iwa";
const ROOT_PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// A content-free location associated with a body-table appearance operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableAppearancePath {
    /// The complete Pages package.
    Package,
    /// One rooted table at checked zero-based positions.
    Table { table: usize },
}

/// A finite resource governed by a body-table appearance operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableAppearanceLimitKind {
    /// Complete source package bytes.
    InputBytes,
    /// Complete candidate package bytes.
    OutputBytes,
    /// Physical package members.
    Entries,
    /// Bytes in one physical member.
    EntryBytes,
    /// Aggregate physical member bytes.
    TotalEntryBytes,
    /// Physical package names and container metadata.
    PackageBytes,
    /// Bytes in one decoded native payload.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native framing or metadata items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Strict wire input bytes.
    WireBytes,
    /// Strict wire output bytes.
    WireOutputBytes,
    /// Strict wire fields.
    WireFields,
    /// Strict wire nesting.
    WireNesting,
    /// Strict wire work.
    WireWork,
    /// Aggregate focused transaction work.
    TransactionWork,
}

impl fmt::Display for BodyTableAppearanceLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A content-redacted body-table appearance failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum BodyTableAppearanceError {
    /// No body table matched the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// More than one body table matched an exact name selector.
    #[error("more than one Pages body table has the requested name")]
    AmbiguousTableName,
    /// The selected rooted table route is ambiguous.
    #[error("the selected Pages body-table route is ambiguous")]
    AmbiguousSelector,
    /// A changed operation targeted a locked table.
    #[error("the selected Pages body table is locked at {path:?}")]
    TableLocked { path: BodyTableAppearancePath },
    /// A native style dependency is not safe to rewrite in this owner.
    #[error("the selected Pages body table appearance has an unsupported dependency at {path:?}")]
    UnsupportedDependency { path: BodyTableAppearancePath },
    /// The source is not an exact supported native profile.
    #[error("this Pages source does not support exact body-table appearance editing")]
    UnsupportedSource,
    /// Rooted ownership or wire framing is invalid.
    #[error("the Pages body body-table appearance source is invalid at {path:?}")]
    InvalidSource { path: BodyTableAppearancePath },
    /// A finite resource ceiling was exceeded.
    #[error(
        "Pages body body-table appearance {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: BodyTableAppearanceLimitKind,
        observed: u64,
        maximum: u64,
        path: BodyTableAppearancePath,
    },
    /// A bounded allocation failed before publication.
    #[error(
        "could not allocate {amount} units for the Pages body body-table appearance transaction"
    )]
    Allocation {
        amount: usize,
        path: BodyTableAppearancePath,
    },
    /// Candidate reopening or locality verification failed.
    #[error("the edited Pages body table appearance failed semantic verification")]
    Verification,
    /// A patch was applied to a package other than its exact source.
    #[error("the body-table appearance patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Clone, PartialEq, Eq)]
struct BodyTableAppearanceTarget {
    native: table_lock::BodyTableTarget,
    appearance: Appearance,
    style_identifier: Option<u64>,
}

/// Immutable appearance settings staged against one package snapshot.
pub struct BodyTableAppearanceEdit<'a> {
    source: &'a Package,
    target: BodyTableAppearanceTarget,
    before: Appearance,
    appearance: Appearance,
    budget: TransactionBudget,
}

impl fmt::Debug for BodyTableAppearanceEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableAppearanceEdit")
            .field("path", &self.path())
            .field("appearance", &self.appearance)
            .finish_non_exhaustive()
    }
}

impl BodyTableAppearanceEdit<'_> {
    /// Return the selected table path.
    #[must_use]
    pub const fn path(&self) -> BodyTableAppearancePath {
        BodyTableAppearancePath::Table {
            table: self.target.native.table_position,
        }
    }

    /// Return the staged appearance.
    #[must_use]
    pub const fn appearance(&self) -> Appearance {
        self.appearance
    }

    /// Replace the staged appearance without touching package bytes.
    #[must_use]
    pub fn set(mut self, appearance: Appearance) -> Self {
        self.appearance = appearance;
        self
    }

    /// Validate and atomically publish the staged appearance.
    ///
    /// Exact semantic no-ops reuse the immutable source allocation. Changed
    /// edits revalidate the selected rooted target, enforce ownership and
    /// lock safety, then rewrite and fully reopen the candidate.
    pub fn commit(self) -> Result<BodyTableAppearanceCommit, BodyTableAppearanceError> {
        commit_edit(self)
    }
}

/// A reversible process-local exact-source appearance patch.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyTableAppearancePatch {
    artifacts: OwnedExactArtifacts,
    target: BodyTableAppearanceTarget,
    before: Appearance,
    after: Appearance,
    source_previews: usize,
    target_previews: usize,
    touched_components: usize,
}

impl fmt::Debug for BodyTableAppearancePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableAppearancePatch")
            .field("path", &self.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyTableAppearancePatch {
    /// Return the selected table path.
    #[must_use]
    pub const fn path(&self) -> BodyTableAppearancePath {
        BodyTableAppearancePath::Table {
            table: self.target.native.table_position,
        }
    }

    /// Return exact source appearance settings.
    #[must_use]
    pub const fn before(&self) -> Appearance {
        self.before
    }

    /// Return exact target appearance settings.
    #[must_use]
    pub const fn after(&self) -> Appearance {
        self.after
    }

    /// Return the source diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch is an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            target: self.target.clone(),
            before: self.after,
            after: self.before,
            source_previews: self.target_previews,
            target_previews: self.source_previews,
            touched_components: self.touched_components,
        }
    }
}

/// Content-free publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyTableAppearanceDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl BodyTableAppearanceDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(deleted_previews: usize, touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Whether exact package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of canonical root previews deleted in this direction.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether a complete candidate package was reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// One fully validated immutable publication.
#[must_use = "a body-table appearance commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyTableAppearanceCommit {
    package: Package,
    patch: BodyTableAppearancePatch,
    diagnostics: BodyTableAppearanceDiagnostics,
}

impl BodyTableAppearanceCommit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyTableAppearancePatch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableAppearanceDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted table's effective appearance.
    pub fn body_table_appearance<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Appearance, BodyTableAppearanceError> {
        let mut budget = TransactionBudget::new(self)?;
        let result = resolve_target_with_budget(self, selector, &mut budget, true);
        Ok(result?.appearance)
    }

    /// Start a selector-first immutable body-table appearance edit.
    pub fn edit_body_table_appearance<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<BodyTableAppearanceEdit<'_>, BodyTableAppearanceError> {
        let mut budget = TransactionBudget::new(self)?;
        // Stage only after the same strict graph/metadata admission used by
        // a read. Changed commits repeat this proof because an edit is an
        // immutable snapshot boundary, not an authority token; exact no-ops
        // retain this staged immutable proof and skip changed-only work.
        let target = resolve_target_with_budget(self, selector, &mut budget, true)?;
        Ok(BodyTableAppearanceEdit {
            source: self,
            before: target.appearance,
            appearance: target.appearance,
            target,
            budget,
        })
    }

    /// Apply a reversible exact-source appearance patch.
    pub fn apply_body_table_appearance(
        &self,
        patch: &BodyTableAppearancePatch,
    ) -> Result<BodyTableAppearanceCommit, BodyTableAppearanceError> {
        let mut budget = TransactionBudget::new(self)?;
        let source_catalog = physical_source(self)?;
        let source_owner = SharedBytes::from_shared_slice(source_catalog.shared_source());
        if !patch.artifacts.authorizes_owner(&source_owner) {
            return Err(BodyTableAppearanceError::PatchConflict);
        }
        let selected =
            resolve_at_with_budget(self, patch.target.native.table_position, &mut budget, true)?;
        if selected.appearance != patch.before || selected.native != patch.target.native {
            return Err(BodyTableAppearanceError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(BodyTableAppearanceCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyTableAppearanceDiagnostics::unchanged(),
            });
        }
        if !source_catalog.source_is_exact() {
            return Err(BodyTableAppearanceError::PatchConflict);
        }
        let target_owner = patch.artifacts.target_owner();
        budget.charge_output_bytes(target_owner.len(), BodyTableAppearancePath::Package)?;
        budget.charge_transaction_work(
            target_owner.len().saturating_mul(2),
            BodyTableAppearancePath::Package,
        )?;
        budget.preflight_candidate_reopen_bytes(
            target_owner.as_ref(),
            BodyTableAppearancePath::Package,
        )?;
        let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
            Arc::<[u8]>::from(target_owner.as_slice()),
            self.state.source.limits(),
        )
        .map_err(|error| map_archive_error(error, BodyTableAppearancePath::Package))?;
        budget.preflight_candidate_catalog(&candidate_source, BodyTableAppearancePath::Package)?;
        let candidate = Package::from_source_catalog(candidate_source)
            .map_err(|_| BodyTableAppearanceError::Verification)?;
        let after = resolve_at_with_budget(
            &candidate,
            patch.target.native.table_position,
            &mut budget,
            true,
        )?;
        if after.appearance != patch.after || after.native != patch.target.native {
            return Err(BodyTableAppearanceError::Verification);
        }
        verify_candidate_locality(
            self,
            &candidate,
            patch.target.clone(),
            selected.style_identifier,
            after.style_identifier,
            patch.source_previews,
            patch.target_previews,
            BodyTableAppearancePath::Package,
            &mut budget,
        )?;
        Ok(BodyTableAppearanceCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableAppearanceDiagnostics::published(
                patch.source_previews.saturating_sub(patch.target_previews),
                patch.touched_components,
            ),
        })
    }
}

fn commit_edit(
    edit: BodyTableAppearanceEdit<'_>,
) -> Result<BodyTableAppearanceCommit, BodyTableAppearanceError> {
    // Capture the content-free path before moving the operation ledger out of
    // the edit.  The ledger is deliberately not Clone/Copy: every subsequent
    // check must debit the same transaction rather than a fresh budget.
    let edit_path = edit.path();
    let catalog = physical_source(edit.source)?;
    let source_owner = SharedBytes::from_shared_slice(catalog.shared_source());
    if edit.before == edit.appearance {
        return Ok(BodyTableAppearanceCommit {
            package: edit.source.snapshot(),
            patch: BodyTableAppearancePatch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                target: edit.target.clone(),
                before: edit.before,
                after: edit.appearance,
                source_previews: 0,
                target_previews: 0,
                touched_components: 0,
            },
            diagnostics: BodyTableAppearanceDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(BodyTableAppearanceError::UnsupportedSource);
    }
    let mut budget = edit.budget;
    let revalidated = resolve_at_with_budget(
        edit.source,
        edit.target.native.table_position,
        &mut budget,
        true,
    )?;
    if revalidated.appearance != edit.before
        || revalidated.native != edit.target.native
        || revalidated.style_identifier != edit.target.style_identifier
    {
        return Err(BodyTableAppearanceError::PatchConflict);
    }
    if revalidated.native.explicit_locked == Some(true) {
        return Err(BodyTableAppearanceError::TableLocked { path: edit_path });
    }
    budget.charge_transaction_work(edit.source.source_bytes().len(), edit_path)?;
    let target = edit.target.clone();
    let old_style = target
        .style_identifier
        .ok_or(BodyTableAppearanceError::UnsupportedDependency { path: edit_path })?;
    let metadata_selector_set = preflight_metadata_for_style(
        edit.source,
        target.native.component_index,
        old_style,
        edit_path,
        &mut budget,
    )?;
    let (new_style, uuid, native_edits) = rewrite_native_appearance(
        edit.source,
        target.clone(),
        old_style,
        edit.appearance,
        edit_path,
        &mut budget,
    )?;
    let metadata_edit = rewrite_metadata_for_style(
        edit.source,
        target.native.component_index,
        new_style,
        uuid,
        native_edits.style_component,
        metadata_selector_set,
        edit_path,
        &mut budget,
    )?;
    let previews = root_preview_deletions(catalog, edit_path, &mut budget)?;
    let mut edits = native_edits.edits;
    budget.charge_allocations(1, edit_path)?;
    edits
        .try_reserve(1)
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: edits.len().saturating_add(1),
            path: edit_path,
        })?;
    edits.push(metadata_edit);
    let touched_components = edits
        .iter()
        .enumerate()
        .filter(|(index, edit)| {
            edits[..*index]
                .iter()
                .all(|previous| previous.name != edit.name)
        })
        .count();
    let mut entry_edits = Vec::new();
    budget.charge_allocations(edits.len().saturating_add(1), edit_path)?;
    entry_edits.try_reserve_exact(edits.len()).map_err(|_| {
        BodyTableAppearanceError::Allocation {
            amount: edits.len(),
            path: edit_path,
        }
    })?;
    for edit in &edits {
        entry_edits.push(EntryEdit::new(edit.name.as_str(), edit.data.as_slice()));
    }
    let physical_limits = catalog.limits();
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&entry_edits, &previews, physical_limits)
        .map_err(|error| map_archive_error(error, edit_path))?;
    let requirements = prepared.execution_requirements();
    budget.preflight_reassembly(requirements, edit_path)?;
    let target_bytes = prepared
        .execute(requirements.exact_limits())
        .map_err(|error| map_archive_error(error, edit_path))?;
    budget.preflight_candidate_reopen_bytes(target_bytes.as_slice(), edit_path)?;
    let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
        target_bytes.into(),
        edit.source.state.source.limits(),
    )
    .map_err(|error| map_archive_error(error, edit_path))?;
    budget.preflight_candidate_catalog(&candidate_source, edit_path)?;
    let package = Package::from_source_catalog(candidate_source)
        .map_err(|_| BodyTableAppearanceError::Verification)?;
    let after = resolve_at_with_budget(&package, target.native.table_position, &mut budget, true)?;
    if after.appearance != edit.appearance || after.native != target.native {
        return Err(BodyTableAppearanceError::Verification);
    }
    verify_candidate_locality(
        edit.source,
        &package,
        target.clone(),
        Some(old_style),
        after.style_identifier,
        previews.len(),
        0,
        edit_path,
        &mut budget,
    )?;
    let target_catalog = physical_source(&package)?;
    let target_owner = SharedBytes::from_shared_slice(target_catalog.shared_source());
    let patch = BodyTableAppearancePatch {
        artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
        target,
        before: edit.before,
        after: edit.appearance,
        source_previews: previews.len(),
        target_previews: 0,
        touched_components,
    };
    Ok(BodyTableAppearanceCommit {
        package,
        patch,
        diagnostics: BodyTableAppearanceDiagnostics::published(previews.len(), touched_components),
    })
}

#[derive(Debug)]
struct NativeEdit {
    name: String,
    data: Vec<u8>,
}

#[derive(Debug)]
struct NativeOutput {
    edits: Vec<NativeEdit>,
    style_component: usize,
}

struct TransactionBudget {
    /// The Pages lock budget is the physical/package ledger.  Keeping it in
    /// this operation budget makes source/catalog and candidate/reopen scans
    /// share the same counters as the focused appearance work below.
    wire: table_lock::WireBudget,
    maximum_wire_bytes: usize,
    maximum_output_bytes: usize,
    maximum_fields: usize,
    maximum_work: usize,
    maximum_styles: usize,
    maximum_components: usize,
    maximum_references: usize,
    maximum_additions: usize,
    maximum_allocations: usize,
    maximum_transaction_work: usize,
    maximum_nesting: usize,
    remaining_fields: usize,
    remaining_work: usize,
    remaining_output_bytes: usize,
    remaining_styles: usize,
    remaining_components: usize,
    remaining_references: usize,
    remaining_additions: usize,
    remaining_allocations: usize,
    remaining_transaction_work: usize,
}

impl TransactionBudget {
    fn new(source: &Package) -> Result<Self, BodyTableAppearanceError> {
        let physical = source.state.source.limits();
        let mut wire = table_lock::WireBudget::new(physical)
            .map_err(|error| map_lock_error(error, BodyTableAppearancePath::Package))?;
        wire.charge_source_catalog(&source.state.source)
            .map_err(|error| map_lock_error(error, BodyTableAppearancePath::Package))?;
        let archive = physical.archive_limits();
        let maximum_wire_bytes = physical.max_iwa_stream_bytes().max(1);
        let maximum_output_bytes = usize::try_from(physical.max_total_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let maximum_fields = maximum_wire_bytes.saturating_mul(8).max(1);
        let maximum_work = maximum_wire_bytes.saturating_mul(32).max(1);
        let maximum_styles = maximum_wire_bytes.max(1);
        let maximum_components = maximum_wire_bytes.max(1);
        let maximum_references = archive.max_metadata_items().saturating_mul(8).max(1);
        let maximum_additions = 64;
        let maximum_transaction_work = usize::try_from(physical.max_total_bytes())
            .unwrap_or(usize::MAX)
            // Selection, strict metadata authority, COW preparation, and the
            // candidate locality pass each traverse bounded source facts.
            // Reserve the complete finite multipass envelope up front.
            .saturating_mul(128)
            .max(1);
        // ArchiveInfo header nesting and protobuf message nesting are distinct
        // axes. The appearance codec consumes the shared protobuf wire ceiling;
        // the physical archive parser continues to enforce its own header cap.
        let maximum_nesting = WireLimits::MAX_NESTING.max(1);
        // Strict metadata ownership is inspected alongside the native style
        // graph.  Bound its visitor/codec scratch by the source component
        // cardinality instead of using a fixed allowance that rejects valid
        // multi-component packages before publication.
        let source_objects = source
            .state
            .source
            .components()
            .iter()
            .fold(0usize, |count, component| {
                count.saturating_add(component.archive().objects.len())
            });
        let source_messages = source
            .state
            .source
            .components()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
            .fold(0usize, |count, object| {
                count.saturating_add(object.messages.len())
            });
        let maximum_allocations = source
            .state
            .source
            .components()
            .len()
            .saturating_mul(64)
            .saturating_add(source_objects.saturating_mul(16))
            .saturating_add(source_messages.saturating_mul(8))
            .saturating_add(256)
            .max(256);
        Ok(Self {
            wire,
            maximum_wire_bytes,
            maximum_output_bytes,
            maximum_fields,
            maximum_work,
            maximum_styles,
            maximum_components,
            maximum_references,
            maximum_additions,
            maximum_allocations,
            maximum_transaction_work,
            maximum_nesting,
            remaining_fields: maximum_fields,
            remaining_work: maximum_work,
            remaining_output_bytes: maximum_output_bytes,
            remaining_styles: maximum_styles,
            remaining_components: maximum_components,
            remaining_references: maximum_references,
            remaining_additions: maximum_additions,
            remaining_allocations: maximum_allocations,
            remaining_transaction_work: maximum_transaction_work,
        })
    }

    fn codec_options(&self, source: &[u8]) -> codec::DecodeOptions {
        // Input, output, retained, and scratch are charged together after a
        // successful stage. Give each pre-allocation axis at most one quarter
        // of the remaining aggregate envelope so their combined requirements
        // cannot overrun the operation ledger before the report is merged.
        let stage_bytes = self.remaining_transaction_work / 4;
        codec::DecodeOptions::new(
            self.maximum_wire_bytes
                .min(source.len().max(1))
                .min(stage_bytes),
            stage_bytes,
            self.remaining_fields,
            self.remaining_work,
            u32::try_from(self.maximum_nesting).unwrap_or(u32::MAX),
            self.remaining_styles,
        )
        .with_max_allocations(self.remaining_allocations)
        .with_max_retained_bytes(stage_bytes)
        .with_max_scratch_bytes(stage_bytes)
    }

    fn charge_fields(
        &mut self,
        amount: usize,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        charge_budget(
            &mut self.remaining_fields,
            self.maximum_fields,
            amount,
            BodyTableAppearanceLimitKind::WireFields,
            path,
        )
    }

    fn charge_work(
        &mut self,
        amount: usize,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        charge_budget(
            &mut self.remaining_work,
            self.maximum_work,
            amount,
            BodyTableAppearanceLimitKind::WireWork,
            path,
        )
    }

    fn charge_styles(
        &mut self,
        amount: usize,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        charge_budget(
            &mut self.remaining_styles,
            self.maximum_styles,
            amount,
            BodyTableAppearanceLimitKind::PayloadItems,
            path,
        )
    }

    fn charge_components(
        &mut self,
        amount: usize,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        charge_budget(
            &mut self.remaining_components,
            self.maximum_components,
            amount,
            BodyTableAppearanceLimitKind::PayloadItems,
            path,
        )
    }

    fn charge_references(
        &mut self,
        amount: usize,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        charge_budget(
            &mut self.remaining_references,
            self.maximum_references,
            amount,
            BodyTableAppearanceLimitKind::PayloadReferences,
            path,
        )
    }

    fn charge_additions(
        &mut self,
        amount: usize,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        charge_budget(
            &mut self.remaining_additions,
            self.maximum_additions,
            amount,
            BodyTableAppearanceLimitKind::PayloadItems,
            path,
        )
    }

    fn charge_allocations(
        &mut self,
        amount: usize,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        charge_budget(
            &mut self.remaining_allocations,
            self.maximum_allocations,
            amount,
            BodyTableAppearanceLimitKind::TransactionWork,
            path,
        )
    }

    fn charge_transaction_work(
        &mut self,
        amount: usize,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        charge_budget(
            &mut self.remaining_transaction_work,
            self.maximum_transaction_work,
            amount,
            BodyTableAppearanceLimitKind::TransactionWork,
            path,
        )
    }

    fn charge_output_bytes(
        &mut self,
        amount: usize,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        charge_budget(
            &mut self.remaining_output_bytes,
            self.maximum_output_bytes,
            amount,
            BodyTableAppearanceLimitKind::OutputBytes,
            path,
        )
    }

    fn charge_nesting(
        &mut self,
        amount: u32,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        let observed = usize::try_from(amount).unwrap_or(usize::MAX);
        if observed > self.maximum_nesting {
            return Err(BodyTableAppearanceError::LimitExceeded {
                kind: BodyTableAppearanceLimitKind::WireNesting,
                observed: u64::try_from(observed).unwrap_or(u64::MAX),
                maximum: u64::try_from(self.maximum_nesting).unwrap_or(u64::MAX),
                path,
            });
        }
        Ok(())
    }

    /// Reserve the bytes and logical work needed before a candidate source is
    /// handed to the ZIP/catalog parser.  The parser has its own physical
    /// limits, but it cannot see this operation's aggregate ledger.
    fn preflight_candidate_reopen_bytes(
        &mut self,
        bytes: &[u8],
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        self.charge_transaction_work(bytes.len(), path)?;
        self.wire
            .charge_output_bytes(bytes.len())
            .map_err(|error| map_lock_error(error, path))?;
        self.wire
            .charge_payload_work(bytes.len())
            .map_err(|error| map_lock_error(error, path))
    }

    /// Charge the catalog walk performed by `Package::from_source_catalog`
    /// before allowing its semantic snapshot to allocate.
    fn preflight_candidate_catalog(
        &mut self,
        catalog: &SourceCatalog,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        let entry_count = catalog.package().iter().count();
        let component_count = catalog.components().len();
        self.charge_transaction_work(catalog.shared_source().len(), path)?;
        self.charge_allocations(
            entry_count
                .saturating_add(component_count)
                .saturating_add(1),
            path,
        )?;
        self.wire
            .charge_source_catalog(catalog)
            .map_err(|error| map_lock_error(error, path))?;
        table_lock::charge_reopen_work(catalog, &mut self.wire)
            .map_err(|error| map_lock_error(error, path))
    }

    fn preflight_reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        self.charge_output_bytes(requirements.output_bytes(), path)?;
        self.charge_allocations(requirements.allocations(), path)?;
        self.charge_transaction_work(requirements.offset_count(), path)?;
        self.charge_transaction_work(
            requirements
                .output_bytes()
                .saturating_add(requirements.scratch_bytes())
                .saturating_add(requirements.retained_bytes()),
            path,
        )
    }

    fn consume_report(
        &mut self,
        report: codec::DecodeReport,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        self.charge_fields(report.fields(), path)?;
        self.charge_work(report.work_bytes(), path)?;
        self.charge_nesting(report.max_depth(), path)?;
        self.charge_allocations(report.allocations(), path)?;
        self.charge_transaction_work(
            report
                .input_bytes()
                .saturating_add(report.output_bytes())
                .saturating_add(report.retained_bytes())
                .saturating_add(report.scratch_bytes()),
            path,
        )
    }

    fn preflight_requirements(
        &mut self,
        requirements: codec::RewriteExecutionRequirements,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        self.charge_fields(requirements.fields(), path)?;
        self.charge_work(requirements.work_bytes(), path)?;
        self.charge_nesting(requirements.max_depth(), path)?;
        self.charge_allocations(requirements.allocations(), path)?;
        self.charge_transaction_work(
            requirements
                .input_bytes()
                .saturating_add(requirements.output_bytes())
                .saturating_add(requirements.retained_bytes())
                .saturating_add(requirements.scratch_bytes()),
            path,
        )
    }

    fn preflight_metadata_requirements(
        &mut self,
        requirements: litchi_iwa_protos::package_metadata_codec::RewriteExecutionRequirements,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        self.charge_fields(requirements.fields(), path)?;
        self.charge_work(requirements.work_bytes(), path)?;
        self.charge_components(requirements.components(), path)?;
        self.charge_references(requirements.references(), path)?;
        self.charge_allocations(requirements.allocations(), path)?;
        self.charge_transaction_work(
            requirements
                .output_bytes()
                .saturating_add(requirements.retained_bytes())
                .saturating_add(requirements.scratch_bytes()),
            path,
        )
    }

    fn metadata_options(
        &self,
        source: &Package,
        bytes: usize,
        additions: usize,
    ) -> MetadataRewriteOptions {
        let base = metadata_options(source, bytes, additions);
        MetadataRewriteOptions::new(
            base.max_input_bytes().min(self.remaining_transaction_work),
            base.max_output_bytes().min(self.remaining_transaction_work),
            base.max_fields().min(self.remaining_fields),
            base.max_work_bytes().min(self.remaining_work),
            base.recursion_limit(),
            base.max_components().min(self.remaining_components),
            base.max_references().min(self.remaining_references),
            base.max_additions().min(self.remaining_additions),
        )
    }

    fn consume_metadata_report(
        &mut self,
        report: litchi_iwa_protos::package_metadata_codec::RewriteReport,
        path: BodyTableAppearancePath,
    ) -> Result<(), BodyTableAppearanceError> {
        self.charge_fields(report.fields(), path)?;
        self.charge_work(report.work_bytes(), path)?;
        self.charge_components(report.components_scanned(), path)?;
        self.charge_references(report.references_scanned(), path)?;
        self.charge_additions(report.additions(), path)?;
        self.charge_nesting(report.max_depth(), path)?;
        self.charge_transaction_work(
            report
                .input_bytes()
                .saturating_add(report.output_bytes())
                .saturating_add(report.retained_bytes())
                .saturating_add(report.scratch_bytes()),
            path,
        )
    }
}

fn charge_budget(
    remaining: &mut usize,
    maximum: usize,
    amount: usize,
    kind: BodyTableAppearanceLimitKind,
    path: BodyTableAppearancePath,
) -> Result<(), BodyTableAppearanceError> {
    if amount > *remaining {
        let observed = maximum.saturating_sub(*remaining).saturating_add(amount);
        return Err(BodyTableAppearanceError::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path,
        });
    }
    *remaining -= amount;
    Ok(())
}

fn rewrite_native_appearance(
    source: &Package,
    target: BodyTableAppearanceTarget,
    old_style: u64,
    appearance: Appearance,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(u64, UuidBits, NativeOutput), BodyTableAppearanceError> {
    let style_location = resolved_location(source, old_style, path, budget)?;
    let style_payload = style_message_data(source, old_style, path, budget)?;
    let (style, style_report) =
        codec::decode_table_style_with_report(style_payload, budget.codec_options(style_payload))
            .map_err(|error| map_codec_error(error, path))?;
    budget.consume_report(style_report, path)?;
    let stylesheet_id = style
        .stylesheet_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(BodyTableAppearanceError::UnsupportedDependency { path })?;
    let stylesheet_location = resolved_location(source, stylesheet_id, path, budget)?;
    let stylesheet_payload = stylesheet_message_data(source, stylesheet_id, path, budget)?;
    if target.native.model_component_index != style_location.0
        || style_location.0 != stylesheet_location.0
    {
        return Err(BodyTableAppearanceError::UnsupportedDependency { path });
    }
    let new_style = next_style_identifier(source, path, budget)?;
    let uuid = style_uuid(new_style);
    if uuid.lower() == 0 && uuid.upper() == 0 {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let model_payload = model_payload(source, &target.native, budget)
        .map_err(|_| BodyTableAppearanceError::InvalidSource { path })?;
    let physical = physical_source(source)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|error| map_archive_error(error, path))?;
    budget.charge_allocations(3, path)?;
    let mut component_indices = Vec::new();
    component_indices
        .try_reserve_exact(3)
        .map_err(|_| BodyTableAppearanceError::Allocation { amount: 3, path })?;
    component_indices.push(target.native.component_index);
    component_indices.push(style_location.0);
    if !component_indices.contains(&stylesheet_location.0) {
        component_indices.push(stylesheet_location.0);
    }
    component_indices.sort_unstable();
    component_indices.dedup();
    let snappy_limits = physical
        .limits()
        .snappy_limits()
        .map_err(|error| map_archive_error(error, path))?;
    let maximum_archive_bytes = archive_limits.max_archive_bytes();
    let maximum_compressed_bytes = SnappyStream::maximum_compressed_len(maximum_archive_bytes)
        .map_err(|error| map_core_error(error, path))?;
    budget.charge_allocations(
        component_indices.len().saturating_mul(4).saturating_add(8),
        path,
    )?;
    budget.charge_transaction_work(
        component_indices
            .len()
            .saturating_mul(maximum_archive_bytes.saturating_add(maximum_compressed_bytes)),
        path,
    )?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(component_indices.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: component_indices.len(),
            path,
        })?;
    let variation = codec::canonical_table_style_variation(
        codec::TableStyleVariationWrite {
            parent_identifier: old_style,
            stylesheet_identifier: stylesheet_id,
            overrides: appearance_overrides(appearance),
        },
        budget.codec_options(style_payload),
    )
    .map_err(|error| map_codec_error(error, path))?;
    budget.consume_report(variation.report(), path)?;
    let model_plan = codec::prepare_table_model_style_rewrite(
        model_payload,
        old_style,
        new_style,
        budget.codec_options(model_payload),
    )
    .map_err(|error| map_codec_error(error, path))?;
    budget.preflight_requirements(model_plan.execution_requirements(), path)?;
    let (model_rewritten, _) = model_plan
        .execute(model_plan.execution_requirements().exact_limits())
        .map_err(|error| map_codec_error(error, path))?;
    let stylesheet_plan = codec::prepare_stylesheet_append(
        stylesheet_payload,
        codec::StylesheetStyleAppend {
            style_identifier: new_style,
            parent_identifier: Some(old_style),
        },
        budget.codec_options(stylesheet_payload),
    )
    .map_err(|error| map_codec_error(error, path))?;
    budget.preflight_requirements(stylesheet_plan.execution_requirements(), path)?;
    let (stylesheet_rewritten, _) = stylesheet_plan
        .execute(stylesheet_plan.execution_requirements().exact_limits())
        .map_err(|error| map_codec_error(error, path))?;
    for component_index in component_indices {
        let component = source
            .state
            .source
            .components()
            .get_index(component_index)
            .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
        budget.charge_allocations(1, path)?;
        budget.charge_transaction_work(component.name().len(), path)?;
        let name = component.name().to_owned();
        let entry = physical
            .package()
            .iter()
            .find(|entry| entry.name() == name)
            .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
        if entry.is_opaque() {
            return Err(BodyTableAppearanceError::UnsupportedSource);
        }
        budget.charge_transaction_work(entry.data().len(), path)?;
        let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
            .map_err(|error| map_core_error(error, path))?;
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(|error| map_core_error(error, path))?;
        archive
            .validate_canonical_object_framing(stream.as_bytes())
            .map_err(|error| map_core_error(error, path))?;
        if component_index == target.native.component_index {
            replace_model_in_archive(
                &mut archive,
                target.native.model_identifier.get(),
                old_style,
                new_style,
                target.native.model_message_type,
                model_rewritten.as_slice(),
                archive_limits,
                path,
                budget,
            )?;
        }
        if component_index == style_location.0 {
            append_style_object(
                &mut archive,
                new_style,
                old_style,
                stylesheet_id,
                variation.bytes(),
                archive_limits,
                path,
                budget,
            )?;
        }
        if component_index == stylesheet_location.0 {
            replace_stylesheet_in_archive(
                &mut archive,
                stylesheet_id,
                new_style,
                stylesheet_rewritten.as_slice(),
                archive_limits,
                path,
                budget,
            )?;
        }
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(|error| map_core_error(error, path))?;
        let compressed =
            SnappyStream::compress(&bytes).map_err(|error| map_core_error(error, path))?;
        edits.push(NativeEdit {
            name,
            data: compressed,
        });
    }
    Ok((
        new_style,
        uuid,
        NativeOutput {
            edits,
            style_component: style_location.0,
        },
    ))
}

fn appearance_overrides(appearance: Appearance) -> codec::AppearanceOverrides {
    codec::AppearanceOverrides {
        row_banding: Some(matches!(
            appearance.row_banding,
            crate::table::appearance::Banding::Enabled
        )),
        row_sizing: Some(matches!(
            appearance.row_sizing,
            crate::table::appearance::RowSizing::FitCellContents
        )),
        body_horizontal: Some(matches!(
            appearance.gridlines.body_horizontal,
            crate::table::appearance::GridlineVisibility::Visible
        )),
        body_vertical: Some(matches!(
            appearance.gridlines.body_vertical,
            crate::table::appearance::GridlineVisibility::Visible
        )),
        header_columns_horizontal: Some(matches!(
            appearance.gridlines.header_columns_horizontal,
            crate::table::appearance::GridlineVisibility::Visible
        )),
        header_rows_vertical: Some(matches!(
            appearance.gridlines.header_rows_vertical,
            crate::table::appearance::GridlineVisibility::Visible
        )),
        footer_rows_vertical: Some(matches!(
            appearance.gridlines.footer_rows_vertical,
            crate::table::appearance::GridlineVisibility::Visible
        )),
    }
}

fn next_style_identifier(
    source: &Package,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<u64, BodyTableAppearanceError> {
    let mut maximum = 0u64;
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            maximum = maximum.max(
                object
                    .archive_info
                    .identifier
                    .ok_or(BodyTableAppearanceError::InvalidSource { path })?,
            );
            budget.charge_transaction_work(1, path)?;
        }
    }
    let metadata = metadata_payload(source, path, budget)?;
    let mut visitor = IdentifierCensus::default();
    let inspection = inspect_package_metadata_with_visitor(
        metadata,
        budget.metadata_options(source, metadata.len(), 0),
        &mut visitor,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(inspection.report(), path)?;
    if visitor.unknown {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let identifier = maximum
        .max(inspection.last_object_identifier())
        .max(visitor.maximum)
        .checked_add(1)
        .filter(|identifier| *identifier != 0)
        .ok_or(BodyTableAppearanceError::LimitExceeded {
            kind: BodyTableAppearanceLimitKind::PayloadObjects,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
            path,
        })?;
    let mut collision = UuidCollisionVisitor::new(style_uuid(identifier));
    let collision_inspection = inspect_package_metadata_with_visitor(
        metadata,
        budget.metadata_options(source, metadata.len(), 0),
        &mut collision,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(collision_inspection.report(), path)?;
    if collision.unknown || collision.found {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    Ok(identifier)
}

fn style_uuid(identifier: u64) -> UuidBits {
    UuidBits::new(
        identifier ^ 0x9e37_79b9_7f4a_7c15,
        identifier.rotate_left(29) ^ 0xd1b5_4a32_d192_ed03,
    )
}

fn resolved_location(
    source: &Package,
    identifier: u64,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(usize, usize), BodyTableAppearanceError> {
    let mut location = None;
    for (component_index, component) in source.state.source.components().iter().enumerate() {
        budget.charge_components(1, path)?;
        budget.charge_transaction_work(component.archive().objects.len(), path)?;
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            if object.archive_info.identifier != Some(identifier) {
                continue;
            }
            if location.replace((component_index, object_index)).is_some() {
                return Err(BodyTableAppearanceError::InvalidSource { path });
            }
        }
    }
    location.ok_or(BodyTableAppearanceError::InvalidSource { path })
}

fn stylesheet_message_data<'source>(
    source: &'source Package,
    identifier: u64,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<&'source [u8], BodyTableAppearanceError> {
    appearance_message_data(source, identifier, STYLESHEET_MESSAGE_TYPE, path, budget)
}

fn appearance_message_data<'source>(
    source: &'source Package,
    identifier: u64,
    expected_type: u32,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<&'source [u8], BodyTableAppearanceError> {
    let (component_index, object_index) = resolved_location(source, identifier, path, budget)?;
    let object = source
        .state
        .source
        .components()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(object_index))
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    if object.archive_info.identifier != Some(identifier) {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    validate_appearance_message_roles(object, expected_type, path, budget)?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == expected_type);
    let message = messages
        .next()
        .ok_or(BodyTableAppearanceError::UnsupportedDependency { path })?;
    if messages.next().is_some() {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    Ok(message.data.as_slice())
}

fn map_codec_error(
    error: codec::DecodeError,
    path: BodyTableAppearancePath,
) -> BodyTableAppearanceError {
    if let Some(amount) = error.allocation_requested() {
        return BodyTableAppearanceError::Allocation { amount, path };
    }
    let Some(limit) = error.resource_limit() else {
        return BodyTableAppearanceError::InvalidSource { path };
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::InputBytes { observed, maximum } => {
            (BodyTableAppearanceLimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::OutputBytes { observed, maximum } => (
            BodyTableAppearanceLimitKind::WireOutputBytes,
            observed,
            maximum,
        ),
        codec::DecodeLimit::Fields { observed, maximum } => {
            (BodyTableAppearanceLimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::WorkBytes { observed, maximum } => {
            (BodyTableAppearanceLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => {
            return BodyTableAppearanceError::LimitExceeded {
                kind: BodyTableAppearanceLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
                path,
            };
        },
        codec::DecodeLimit::Styles { observed, maximum } => (
            BodyTableAppearanceLimitKind::PayloadItems,
            observed,
            maximum,
        ),
        codec::DecodeLimit::Allocations { observed, maximum } => (
            BodyTableAppearanceLimitKind::TransactionWork,
            observed,
            maximum,
        ),
        codec::DecodeLimit::RetainedBytes { observed, maximum }
        | codec::DecodeLimit::ScratchBytes { observed, maximum } => (
            BodyTableAppearanceLimitKind::TransactionWork,
            observed,
            maximum,
        ),
        _ => return BodyTableAppearanceError::InvalidSource { path },
    };
    BodyTableAppearanceError::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        path,
    }
}

fn snapshot_appearance(snapshot: codec::AppearanceSnapshot) -> Appearance {
    Appearance {
        row_banding: if snapshot.row_banding {
            crate::table::appearance::Banding::Enabled
        } else {
            crate::table::appearance::Banding::Disabled
        },
        row_sizing: if snapshot.row_sizing {
            crate::table::appearance::RowSizing::FitCellContents
        } else {
            crate::table::appearance::RowSizing::Fixed
        },
        gridlines: crate::table::appearance::Gridlines {
            body_horizontal: bool_gridline(snapshot.body_horizontal),
            body_vertical: bool_gridline(snapshot.body_vertical),
            header_columns_horizontal: bool_gridline(snapshot.header_columns_horizontal),
            header_rows_vertical: bool_gridline(snapshot.header_rows_vertical),
            footer_rows_vertical: bool_gridline(snapshot.footer_rows_vertical),
        },
    }
}

const fn bool_gridline(value: bool) -> crate::table::appearance::GridlineVisibility {
    if value {
        crate::table::appearance::GridlineVisibility::Visible
    } else {
        crate::table::appearance::GridlineVisibility::Hidden
    }
}

struct SelectorVisitor<'source> {
    locators: &'source [&'source str],
    identifiers: Vec<Option<u64>>,
    effective_locators: Vec<Option<String>>,
    duplicate: bool,
    unknown: bool,
}

struct ParentReferenceVisitor<'source> {
    source_identifier: u64,
    source_locator: &'source str,
    target_identifier: u64,
    object_identifier: u64,
    exact: usize,
    related: usize,
    versioned: bool,
    wrong_weakness: bool,
    unexpected: bool,
    unknown: bool,
}

struct RegistryTarget {
    target_identifier: u64,
    expected_component_identifier: u64,
    allow_external_reference: bool,
    current_uuid: usize,
    versioned_uuid: bool,
    cross_component_uuid: bool,
    external_reference: bool,
    data_owner: bool,
    ambiguous: bool,
    root_data_map: bool,
}

struct RegistrySetVisitor<'targets> {
    targets: &'targets mut [RegistryTarget],
    unknown: bool,
}

struct GlobalReferenceAuthority<'source> {
    physical_identifiers: &'source HashSet<u64>,
    style_identifier: u64,
    invalid: bool,
}

struct MetadataSelectorSet {
    identifiers: Vec<u64>,
    effective_locators: Vec<String>,
    last_object_identifier: u64,
}

#[derive(Debug, Clone, Copy)]
struct MetadataRoute {
    component_index: usize,
    object_index: usize,
    message_index: usize,
}

fn normalized_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

fn metadata_route(
    source: &Package,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<Option<MetadataRoute>, BodyTableAppearanceError> {
    let mut route = None;
    for (component_index, component) in source.state.source.components().iter().enumerate() {
        budget.charge_components(1, path)?;
        budget.charge_transaction_work(component.archive().objects.len(), path)?;
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            budget.charge_transaction_work(object.messages.len(), path)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != METADATA_MESSAGE_TYPE {
                    continue;
                }
                if component.name() != METADATA_ENTRY_NAME || route.is_some() {
                    return Ok(None);
                }
                route = Some(MetadataRoute {
                    component_index,
                    object_index,
                    message_index,
                });
            }
        }
    }
    Ok(route)
}

impl MetadataSelectorSet {
    fn selectors(
        &self,
        path: BodyTableAppearancePath,
    ) -> Result<Vec<ComponentSelector<'_>>, BodyTableAppearanceError> {
        let mut selectors = Vec::new();
        selectors
            .try_reserve_exact(self.identifiers.len())
            .map_err(|_| BodyTableAppearanceError::Allocation {
                amount: self.identifiers.len(),
                path,
            })?;
        selectors.extend(
            self.identifiers
                .iter()
                .zip(self.effective_locators.iter())
                .map(|(identifier, locator)| ComponentSelector::new(*identifier, locator.as_str())),
        );
        Ok(selectors)
    }
}

#[derive(Default)]
struct IdentifierCensus {
    maximum: u64,
    unknown: bool,
}

impl IdentifierCensus {
    fn record(&mut self, identifier: u64) {
        self.maximum = self.maximum.max(identifier);
    }
}

impl PackageMetadataVisitor for IdentifierCensus {
    fn visit_unknown_field(&mut self) -> Result<(), MetadataRewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(component.identifier());
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: litchi_iwa_protos::package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(binding.object_identifier());
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(reference.target_component_identifier());
        if let Some(identifier) = reference.object_identifier() {
            self.record(identifier);
        }
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::DataReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(reference.data_identifier());
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: litchi_iwa_protos::package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(owner.data_identifier());
        self.record(owner.object_identifier());
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataRewriteError> {
        self.record(identifier);
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), MetadataRewriteError> {
        self.record(object_identifier);
        Ok(())
    }
}

struct UuidCollisionVisitor {
    target: UuidBits,
    found: bool,
    unknown: bool,
}

impl UuidCollisionVisitor {
    const fn new(target: UuidBits) -> Self {
        Self {
            target,
            found: false,
            unknown: false,
        }
    }
}

impl PackageMetadataVisitor for UuidCollisionVisitor {
    fn visit_unknown_field(&mut self) -> Result<(), MetadataRewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: litchi_iwa_protos::package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.found |= binding.uuid() == self.target;
        Ok(())
    }
}

impl ArchiveReferenceVisitor for GlobalReferenceAuthority<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        match occurrence.kind {
            ArchiveReferenceKind::Object => {
                if !self
                    .physical_identifiers
                    .contains(&occurrence.referenced_identifier)
                {
                    self.invalid = true;
                }
            },
            ArchiveReferenceKind::Data => {
                if occurrence.referenced_identifier == self.style_identifier {
                    self.invalid = true;
                }
            },
        }
        Ok(())
    }
}

impl<'source> SelectorVisitor<'source> {
    fn new(
        locators: &'source [&'source str],
        path: BodyTableAppearancePath,
    ) -> Result<Self, BodyTableAppearanceError> {
        let mut identifiers = Vec::new();
        identifiers.try_reserve_exact(locators.len()).map_err(|_| {
            BodyTableAppearanceError::Allocation {
                amount: locators.len(),
                path,
            }
        })?;
        identifiers.resize(locators.len(), None);
        let mut effective_locators = Vec::new();
        effective_locators
            .try_reserve_exact(locators.len())
            .map_err(|_| BodyTableAppearanceError::Allocation {
                amount: locators.len(),
                path,
            })?;
        effective_locators.resize_with(locators.len(), || None);
        Ok(Self {
            locators,
            identifiers,
            effective_locators,
            duplicate: false,
            unknown: false,
        })
    }
}

impl PackageMetadataVisitor for SelectorVisitor<'_> {
    fn visit_unknown_field(&mut self) -> Result<(), MetadataRewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        if !component.is_current() {
            return Ok(());
        }
        for (index, locator) in self.locators.iter().enumerate() {
            if component.preferred_locator() != *locator
                && component.effective_locator() != *locator
            {
                continue;
            }
            if self.identifiers[index]
                .replace(component.identifier())
                .is_some()
            {
                self.duplicate = true;
            }
            self.effective_locators[index] = Some(component.effective_locator().to_owned());
        }
        Ok(())
    }
}

impl<'source> ParentReferenceVisitor<'source> {
    fn new(
        source: ComponentSelector<'source>,
        target: ComponentSelector<'source>,
        object: u64,
    ) -> Self {
        Self {
            source_identifier: source.identifier(),
            source_locator: source.locator(),
            target_identifier: target.identifier(),
            object_identifier: object,
            exact: 0,
            related: 0,
            versioned: false,
            wrong_weakness: false,
            unexpected: false,
            unknown: false,
        }
    }
}

impl PackageMetadataVisitor for ParentReferenceVisitor<'_> {
    fn visit_unknown_field(&mut self) -> Result<(), MetadataRewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        let source = reference.source();
        let related = source.identifier() == self.source_identifier
            && reference.target_component_identifier() == self.target_identifier
            && reference.object_identifier() == Some(self.object_identifier);
        if !related {
            if reference.object_identifier() == Some(self.object_identifier) {
                self.unexpected = true;
            }
            return Ok(());
        }
        self.related = self.related.saturating_add(1);
        if reference.is_versioned() || !source.is_current() {
            self.versioned = true;
            return Ok(());
        }
        if source.effective_locator() != self.source_locator {
            self.wrong_weakness = true;
            return Ok(());
        }
        if reference.is_weak().is_some() {
            self.wrong_weakness = true;
            return Ok(());
        }
        self.exact = self.exact.saturating_add(1);
        Ok(())
    }
}

impl RegistryTarget {
    fn new(
        component: ComponentSelector<'_>,
        target_identifier: u64,
        allow_external_reference: bool,
    ) -> Self {
        Self {
            target_identifier,
            expected_component_identifier: component.identifier(),
            allow_external_reference,
            current_uuid: 0,
            versioned_uuid: false,
            cross_component_uuid: false,
            external_reference: false,
            data_owner: false,
            ambiguous: false,
            root_data_map: false,
        }
    }

    fn valid(&self) -> bool {
        self.current_uuid == 1
            && !self.versioned_uuid
            && !self.cross_component_uuid
            && !self.data_owner
            && !self.ambiguous
            && !self.root_data_map
            && (self.allow_external_reference || !self.external_reference)
    }
}

impl PackageMetadataVisitor for RegistrySetVisitor<'_> {
    fn visit_unknown_field(&mut self) -> Result<(), MetadataRewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: litchi_iwa_protos::package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        for target in self.targets.iter_mut() {
            if binding.object_identifier() != target.target_identifier {
                continue;
            }
            if binding.component().is_current() {
                target.current_uuid = target.current_uuid.saturating_add(1);
                if binding.component().identifier() != target.expected_component_identifier {
                    target.cross_component_uuid = true;
                }
            } else {
                target.versioned_uuid = true;
            }
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        let Some(identifier) = reference.object_identifier() else {
            return Ok(());
        };
        for target in self.targets.iter_mut() {
            if identifier == target.target_identifier {
                target.external_reference = true;
            }
        }
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::DataReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        for target in self.targets.iter_mut() {
            if reference.data_identifier() == target.target_identifier {
                target.data_owner = true;
            }
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: litchi_iwa_protos::package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        for target in self.targets.iter_mut() {
            if owner.object_identifier() == target.target_identifier
                || owner.data_identifier() == target.target_identifier
            {
                target.data_owner = true;
            }
        }
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataRewriteError> {
        for target in self.targets.iter_mut() {
            if identifier == target.target_identifier {
                target.ambiguous = true;
            }
        }
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), MetadataRewriteError> {
        for target in self.targets.iter_mut() {
            if object_identifier == target.target_identifier {
                target.root_data_map = true;
            }
        }
        Ok(())
    }
}

fn metadata_payload<'source>(
    source: &'source Package,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<&'source [u8], BodyTableAppearanceError> {
    let route = metadata_route(source, path, budget)?
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    source
        .state
        .source
        .components()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.messages.get(route.message_index))
        .filter(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(BodyTableAppearanceError::InvalidSource { path })
}

fn metadata_options(source: &Package, bytes: usize, additions: usize) -> MetadataRewriteOptions {
    let physical_max = source.state.source.limits().max_iwa_stream_bytes();
    MetadataRewriteOptions::new(
        bytes.max(1).min(physical_max),
        bytes.saturating_add(4_096).max(1).min(physical_max),
        bytes.saturating_mul(64).clamp(1, WireLimits::MAX_FIELDS),
        bytes
            .saturating_mul(256)
            .saturating_add(additions.saturating_mul(bytes))
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        64,
        bytes
            .max(source.state.source.components().len())
            .saturating_mul(4)
            .max(1),
        source
            .state
            .source
            .limits()
            .archive_limits()
            .max_metadata_items()
            .max(1),
        additions.max(1),
    )
}

fn metadata_selectors(
    source: &Package,
    component_indices: &[usize],
    payload: &[u8],
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<MetadataSelectorSet, BodyTableAppearanceError> {
    budget.charge_allocations(
        component_indices.len().saturating_mul(3).saturating_add(4),
        path,
    )?;
    let mut locators = Vec::new();
    locators
        .try_reserve_exact(component_indices.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: component_indices.len(),
            path,
        })?;
    for &index in component_indices {
        let component = source
            .state
            .source
            .components()
            .get_index(index)
            .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
        locators.push(normalized_locator(component.name()));
    }
    let locator_refs = locators;
    let mut visitor = SelectorVisitor::new(&locator_refs, path)?;
    let inspection = inspect_package_metadata_with_visitor(
        payload,
        budget.metadata_options(source, payload.len(), component_indices.len()),
        &mut visitor,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(inspection.report(), path)?;
    if visitor.unknown || visitor.duplicate || visitor.identifiers.iter().any(Option::is_none) {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(visitor.identifiers.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: visitor.identifiers.len(),
            path,
        })?;
    let mut effective_locators = Vec::new();
    effective_locators
        .try_reserve_exact(visitor.identifiers.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: visitor.identifiers.len(),
            path,
        })?;
    for (identifier, locator) in visitor
        .identifiers
        .into_iter()
        .zip(visitor.effective_locators)
    {
        identifiers.push(identifier.ok_or(BodyTableAppearanceError::InvalidSource { path })?);
        effective_locators.push(locator.ok_or(BodyTableAppearanceError::InvalidSource { path })?);
    }
    Ok(MetadataSelectorSet {
        identifiers,
        effective_locators,
        last_object_identifier: inspection.last_object_identifier(),
    })
}

fn validate_parent_external_reference(
    source: &Package,
    payload: &[u8],
    model_selector: ComponentSelector<'_>,
    style_selector: ComponentSelector<'_>,
    old_style: u64,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let mut visitor = ParentReferenceVisitor::new(model_selector, style_selector, old_style);
    let inspection = inspect_package_metadata_with_visitor(
        payload,
        budget.metadata_options(source, payload.len(), 0),
        &mut visitor,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(inspection.report(), path)?;
    if visitor.unknown
        || visitor.exact != 1
        || visitor.related != 1
        || visitor.versioned
        || visitor.wrong_weakness
        || visitor.unexpected
    {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    Ok(())
}

fn validate_style_registry_entries(
    source: &Package,
    payload: &[u8],
    targets: &mut [RegistryTarget],
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let mut visitor = RegistrySetVisitor {
        targets,
        unknown: false,
    };
    let inspection = inspect_package_metadata_with_visitor(
        payload,
        budget.metadata_options(source, payload.len(), 0),
        &mut visitor,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(inspection.report(), path)?;
    if visitor.unknown || visitor.targets.iter().any(|target| !target.valid()) {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    Ok(())
}

fn validate_metadata_graph_registries(
    source: &Package,
    payload: &[u8],
    model_component: usize,
    model_identifier: u64,
    routed_identifiers: &[u64],
    parent_styles: &[codec::TableStyleNode<'_>],
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let mut component_indices = Vec::new();
    let target_count = routed_identifiers
        .len()
        .saturating_add(parent_styles.len())
        .saturating_add(1);
    budget.charge_allocations(target_count, path)?;
    component_indices
        .try_reserve_exact(target_count)
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: target_count,
            path,
        })?;
    component_indices.push(model_component);
    for identifier in routed_identifiers
        .iter()
        .copied()
        .chain(parent_styles.iter().map(|node| (*node).identifier()))
    {
        let component = resolved_location(source, identifier, path, budget)?.0;
        if !component_indices.contains(&component) {
            component_indices.push(component);
        }
    }
    let selector_set = metadata_selectors(source, &component_indices, payload, path, budget)?;
    budget.charge_allocations(selector_set.identifiers.len(), path)?;
    let selectors = selector_set.selectors(path)?;
    let mut targets = Vec::new();
    budget.charge_allocations(target_count, path)?;
    targets
        .try_reserve_exact(target_count)
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: target_count,
            path,
        })?;
    let model_selector_index = component_indices
        .iter()
        .position(|component| *component == model_component)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    targets.push(RegistryTarget::new(
        *selectors
            .get(model_selector_index)
            .ok_or(BodyTableAppearanceError::InvalidSource { path })?,
        model_identifier,
        true,
    ));
    for identifier in routed_identifiers
        .iter()
        .copied()
        .chain(parent_styles.iter().map(|node| (*node).identifier()))
    {
        let component = resolved_location(source, identifier, path, budget)?.0;
        let selector_index = component_indices
            .iter()
            .position(|candidate| *candidate == component)
            .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
        targets.push(RegistryTarget::new(
            *selectors
                .get(selector_index)
                .ok_or(BodyTableAppearanceError::InvalidSource { path })?,
            identifier,
            true,
        ));
    }
    validate_style_registry_entries(source, payload, &mut targets, path, budget)
}

fn map_metadata_error(
    error: MetadataRewriteError,
    path: BodyTableAppearancePath,
) -> BodyTableAppearanceError {
    if let Some(amount) = error.allocation_request() {
        return BodyTableAppearanceError::Allocation { amount, path };
    }
    let Some(limit) = error.resource_limit() else {
        return BodyTableAppearanceError::InvalidSource { path };
    };
    let (kind, observed, maximum) = match limit {
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::InputBytes {
            observed,
            maximum,
        } => (BodyTableAppearanceLimitKind::WireBytes, observed, maximum),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::OutputBytes {
            observed,
            maximum,
        } => (
            BodyTableAppearanceLimitKind::WireOutputBytes,
            observed,
            maximum,
        ),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
            (BodyTableAppearanceLimitKind::WireFields, observed, maximum)
        },
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
            (BodyTableAppearanceLimitKind::WireWork, observed, maximum)
        },
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
            return BodyTableAppearanceError::LimitExceeded {
                kind: BodyTableAppearanceLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
                path,
            };
        },
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
            observed,
            maximum,
        } => (
            BodyTableAppearanceLimitKind::PayloadItems,
            observed,
            maximum,
        ),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::References {
            observed,
            maximum,
        } => (
            BodyTableAppearanceLimitKind::PayloadReferences,
            observed,
            maximum,
        ),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Additions {
            observed,
            maximum,
        } => (
            BodyTableAppearanceLimitKind::PayloadItems,
            observed,
            maximum,
        ),
        _ => return BodyTableAppearanceError::InvalidSource { path },
    };
    BodyTableAppearanceError::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        path,
    }
}

fn preflight_metadata_for_style(
    source: &Package,
    model_component: usize,
    old_style: u64,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<MetadataSelectorSet, BodyTableAppearanceError> {
    let style_component = resolved_location(source, old_style, path, budget)?.0;
    let payload = metadata_payload(source, path, budget)?;
    let one_component = [model_component];
    let two_components = [model_component, style_component];
    let component_indices: &[usize] = if model_component == style_component {
        &one_component
    } else {
        &two_components
    };
    let selector_set = metadata_selectors(source, component_indices, payload, path, budget)?;
    budget.charge_allocations(selector_set.identifiers.len(), path)?;
    let selectors = selector_set.selectors(path)?;
    let model_selector = selectors
        .first()
        .copied()
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    let style_selector = if model_component == style_component {
        model_selector
    } else {
        selectors
            .get(1)
            .copied()
            .ok_or(BodyTableAppearanceError::InvalidSource { path })?
    };
    if model_component != style_component {
        validate_parent_external_reference(
            source,
            payload,
            model_selector,
            style_selector,
            old_style,
            path,
            budget,
        )?;
    }
    let style_payload = style_message_data(source, old_style, path, budget)?;
    let (style, style_report) =
        codec::decode_table_style_with_report(style_payload, budget.codec_options(style_payload))
            .map_err(|error| map_codec_error(error, path))?;
    budget.consume_report(style_report, path)?;
    let stylesheet_identifier = style
        .stylesheet_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(BodyTableAppearanceError::UnsupportedDependency { path })?;
    let stylesheet_component = resolved_location(source, stylesheet_identifier, path, budget)?.0;
    if stylesheet_component != style_component {
        return Err(BodyTableAppearanceError::UnsupportedDependency { path });
    }
    let stylesheet_selector = style_selector;
    let mut registry_targets = [
        RegistryTarget::new(
            style_selector,
            old_style,
            model_component != style_component,
        ),
        RegistryTarget::new(stylesheet_selector, stylesheet_identifier, false),
    ];
    // Native Numbers commonly uses the StylesheetArchive object itself as
    // the current component root. ComponentInfo then names that root through
    // `identifier` and does not repeat it in the component UUID map.
    if stylesheet_identifier != stylesheet_selector.identifier() {
        validate_style_registry_entries(source, payload, &mut registry_targets, path, budget)?;
    } else {
        validate_style_registry_entries(source, payload, &mut registry_targets[..1], path, budget)?;
    }
    Ok(selector_set)
}

fn rewrite_metadata_for_style(
    source: &Package,
    model_component: usize,
    new_style: u64,
    uuid: UuidBits,
    style_component: usize,
    selector_set: MetadataSelectorSet,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<NativeEdit, BodyTableAppearanceError> {
    let payload = metadata_payload(source, path, budget)?;
    budget.charge_allocations(selector_set.identifiers.len(), path)?;
    let selectors = selector_set.selectors(path)?;
    let expected_last = selector_set.last_object_identifier;
    let model_selector = selectors
        .first()
        .copied()
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    let style_selector = if model_component == style_component {
        model_selector
    } else {
        selectors
            .get(1)
            .copied()
            .ok_or(BodyTableAppearanceError::InvalidSource { path })?
    };
    budget.charge_allocations(4, path)?;
    let uuid_additions = vec![ObjectUuidAddition::new(style_selector, new_style, uuid)];
    let mut external_additions = Vec::new();
    if model_component != style_component {
        external_additions.push(ExternalReferenceAddition::new(
            model_selector,
            style_selector,
            new_style,
            None,
        ));
    }
    let mut save_selectors = Vec::new();
    save_selectors.push(model_selector);
    if model_component != style_component {
        save_selectors.push(style_selector);
    }
    let metadata_options = budget.metadata_options(
        source,
        payload.len(),
        uuid_additions
            .len()
            .saturating_add(external_additions.len()),
    );
    let batch = AdditionSaveTokenBatch::new(
        MetadataBatch::new(
            expected_last,
            new_style,
            &uuid_additions,
            &external_additions,
        ),
        SaveTokenBatch::new(save_selectors.as_slice()),
    );
    let prepared =
        prepare_package_metadata_additions_and_save_tokens(payload, batch, metadata_options)
            .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(prepared.prepare_report(), path)?;
    let requirements = prepared.execution_requirements();
    budget.preflight_metadata_requirements(requirements, path)?;
    let rewritten = prepared
        .execute(requirements.exact_limits())
        .map_err(|error| map_metadata_error(error, path))?
        .into_bytes();
    rewrite_metadata_entry(source, rewritten, path, budget)
}

fn rewrite_metadata_entry(
    source: &Package,
    payload: Vec<u8>,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<NativeEdit, BodyTableAppearanceError> {
    let route = metadata_route(source, path, budget)?
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    let physical = physical_source(source)?;
    let entry = physical
        .package()
        .iter()
        .find(|entry| entry.name() == METADATA_ENTRY_NAME)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    if entry.is_opaque() {
        return Err(BodyTableAppearanceError::UnsupportedSource);
    }
    budget.charge_transaction_work(entry.data().len(), path)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|error| map_archive_error(error, path))?;
    let maximum_archive_bytes = archive_limits.max_archive_bytes();
    let maximum_compressed_bytes = SnappyStream::maximum_compressed_len(maximum_archive_bytes)
        .map_err(|error| map_core_error(error, path))?;
    budget.charge_allocations(4, path)?;
    budget.charge_transaction_work(
        maximum_archive_bytes.saturating_add(maximum_compressed_bytes),
        path,
    )?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical
            .limits()
            .snappy_limits()
            .map_err(|error| map_archive_error(error, path))?,
    )
    .map_err(|error| map_core_error(error, path))?;
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(|error| map_core_error(error, path))?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|error| map_core_error(error, path))?;
    let object = archive
        .objects
        .get_mut(route.object_index)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    let message = object
        .messages
        .get(route.message_index)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    if message.type_ != METADATA_MESSAGE_TYPE {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    object
        .replace_message_preserving_header_with_limits(
            route.message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
            archive_limits,
        )
        .map_err(|error| map_core_error(error, path))?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|error| map_core_error(error, path))?;
    let compressed = SnappyStream::compress(&bytes).map_err(|error| map_core_error(error, path))?;
    Ok(NativeEdit {
        name: METADATA_ENTRY_NAME.to_owned(),
        data: compressed,
    })
}

/// Reject an object that mixes two known appearance roles.  The native wire
/// format permits unrelated message types on one object, but treating a
/// TableInfo/TableModel/style/preset/network/stylesheet alias as an
/// unrelated message would let a redirected role survive a COW transition.
/// Unknown message types remain opaque and are preserved.
fn validate_appearance_message_roles(
    object: &ArchiveObject,
    expected_type: u32,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    budget.charge_transaction_work(object.messages.len(), path)?;
    let mut expected = 0usize;
    for message in &object.messages {
        if !matches!(
            message.type_,
            STYLESHEET_MESSAGE_TYPE
                | 6_000
                | 6_001
                | TABLE_STYLE_MESSAGE_TYPE
                | TABLE_STYLE_PRESET_MESSAGE_TYPE
                | TABLE_STYLE_NETWORK_MESSAGE_TYPE
        ) {
            continue;
        }
        if message.type_ != expected_type {
            return Err(BodyTableAppearanceError::InvalidSource { path });
        }
        expected = expected.saturating_add(1);
    }
    if expected != 1 {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    Ok(())
}

fn copy_bytes(
    source: &[u8],
    path: BodyTableAppearancePath,
) -> Result<Vec<u8>, BodyTableAppearanceError> {
    let mut copied = Vec::new();
    copied
        .try_reserve_exact(source.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: source.len(),
            path,
        })?;
    copied.extend_from_slice(source);
    Ok(copied)
}

fn copy_identifiers(
    source: &[u64],
    additional: usize,
    path: BodyTableAppearancePath,
) -> Result<Vec<u64>, BodyTableAppearanceError> {
    let capacity = source
        .len()
        .checked_add(additional)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    let mut copied = Vec::new();
    copied
        .try_reserve_exact(capacity)
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: capacity,
            path,
        })?;
    copied.extend_from_slice(source);
    Ok(copied)
}

fn replace_model_in_archive(
    archive: &mut Archive,
    model_identifier: u64,
    old_style: u64,
    new_style: u64,
    model_type: u32,
    replacement: &[u8],
    limits: litchi_iwa_core::Limits,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let object = archive
        .object_mut(model_identifier)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    validate_appearance_message_roles(object, model_type, path, budget)?;
    let mut indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == model_type);
    let message_index = indexes
        .next()
        .map(|(index, _)| index)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    if indexes.next().is_some() {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    let nested_references = info.field_infos.iter().try_fold(0usize, |count, field| {
        count
            .checked_add(field.object_references.len())
            .and_then(|value| value.checked_add(field.data_references.len()))
    });
    let nested_references =
        nested_references.ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    budget.charge_fields(info.field_infos.len(), path)?;
    budget.charge_references(
        info.object_references
            .len()
            .saturating_add(info.data_references.len())
            .saturating_add(nested_references),
        path,
    )?;
    budget.charge_transaction_work(
        info.field_infos
            .iter()
            .map(|field| field.path.as_slice().len())
            .sum::<usize>()
            .saturating_add(info.field_infos.len()),
        path,
    )?;
    budget.charge_allocations(info.object_references.len().saturating_add(4), path)?;
    budget.charge_transaction_work(replacement.len(), path)?;
    let before = copy_identifiers(&info.object_references, 0, path)?;
    if before
        .iter()
        .filter(|identifier| **identifier == old_style)
        .count()
        != 1
        || before.contains(&new_style)
    {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    // ArchiveInfo transitions retain surviving aggregate references in their
    // source order and publish newly introduced identifiers as a final suffix.
    // The selected field-local edge still changes directly from old to new.
    let mut after = Vec::new();
    after
        .try_reserve_exact(before.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: before.len(),
            path,
        })?;
    after.extend(
        before
            .iter()
            .copied()
            .filter(|identifier| *identifier != old_style),
    );
    after.push(new_style);
    let replacement = copy_bytes(replacement, path)?;
    replace_message_with_reference_transition(
        object,
        message_index,
        RawMessage {
            type_: object.messages[message_index].type_,
            data: replacement,
        },
        before.as_slice(),
        after.as_slice(),
        &[3],
        Some((
            std::slice::from_ref(&old_style),
            std::slice::from_ref(&new_style),
        )),
        false,
        limits,
        path,
        budget,
    )
}

fn append_style_object(
    archive: &mut Archive,
    identifier: u64,
    parent: u64,
    stylesheet: u64,
    payload: &[u8],
    limits: litchi_iwa_core::Limits,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    if archive
        .objects
        .iter()
        .any(|object| object.archive_info.identifier == Some(identifier))
    {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    budget.charge_allocations(payload.len().saturating_add(2), path)?;
    archive
        .objects
        .try_reserve(1)
        .map_err(|_| BodyTableAppearanceError::Allocation { amount: 1, path })?;
    let payload = copy_bytes(payload, path)?;
    let mut messages = Vec::new();
    messages
        .try_reserve_exact(1)
        .map_err(|_| BodyTableAppearanceError::Allocation { amount: 1, path })?;
    messages.push(RawMessage {
        type_: TABLE_STYLE_MESSAGE_TYPE,
        data: payload,
    });
    let mut object = ArchiveObject::new_with_limits(identifier, messages, limits)
        .map_err(|error| map_core_error(error, path))?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(2)
        .map_err(|_| BodyTableAppearanceError::Allocation { amount: 2, path })?;
    references.extend([parent, stylesheet]);
    object.archive_info.message_infos[0].object_references = references;
    archive.objects.push(object);
    Ok(())
}

fn replace_stylesheet_in_archive(
    archive: &mut Archive,
    stylesheet_identifier: u64,
    new_style: u64,
    replacement: &[u8],
    limits: litchi_iwa_core::Limits,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let object = archive
        .object_mut(stylesheet_identifier)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    validate_appearance_message_roles(object, STYLESHEET_MESSAGE_TYPE, path, budget)?;
    let mut indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == STYLESHEET_MESSAGE_TYPE);
    let message_index = indexes
        .next()
        .map(|(index, _)| index)
        .ok_or(BodyTableAppearanceError::UnsupportedDependency { path })?;
    if indexes.next().is_some() {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    budget.charge_allocations(info.object_references.len().saturating_add(4), path)?;
    budget.charge_transaction_work(replacement.len(), path)?;
    let before = copy_identifiers(&info.object_references, 0, path)?;
    if before.contains(&new_style) {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let mut after = copy_identifiers(&before, 1, path)?;
    after.push(new_style);
    let replacement = copy_bytes(replacement, path)?;
    replace_message_with_reference_transition(
        object,
        message_index,
        RawMessage {
            type_: STYLESHEET_MESSAGE_TYPE,
            data: replacement,
        },
        before.as_slice(),
        after.as_slice(),
        &[1],
        Some((&before, &after)),
        false,
        limits,
        path,
        budget,
    )
}

fn replace_message_with_reference_transition(
    object: &mut ArchiveObject,
    message_index: usize,
    message: RawMessage,
    before: &[u64],
    after: &[u64],
    field_path: &[u32],
    field_references: Option<(&[u64], &[u64])>,
    required_field_path: bool,
    limits: litchi_iwa_core::Limits,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    budget.charge_fields(info.field_infos.len(), path)?;
    budget.charge_references(before.len().saturating_add(after.len()), path)?;
    budget.charge_allocations(
        info.field_infos
            .len()
            .saturating_mul(4)
            .saturating_add(before.len())
            .saturating_add(after.len()),
        path,
    )?;
    let nested_references = info
        .field_infos
        .iter()
        .map(|field| field.object_references.len())
        .sum::<usize>();
    budget.charge_references(nested_references, path)?;
    budget.charge_allocations(nested_references.saturating_add(2), path)?;
    let fields = &object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?
        .field_infos;
    let mut field_indices = Vec::new();
    field_indices.try_reserve_exact(fields.len()).map_err(|_| {
        BodyTableAppearanceError::Allocation {
            amount: fields.len(),
            path,
        }
    })?;
    for (index, field) in fields.iter().enumerate() {
        if field.path.as_slice() == field_path {
            field_indices.push(index);
        }
    }
    if field_indices.len() > 1 || (required_field_path && field_indices.len() != 1) {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let mut before_set = HashSet::new();
    before_set
        .try_reserve(before.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: before.len(),
            path,
        })?;
    let mut after_set = HashSet::new();
    after_set
        .try_reserve(after.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: after.len(),
            path,
        })?;
    before_set.extend(before.iter().copied());
    after_set.extend(after.iter().copied());
    if before_set.len() != before.len() || after_set.len() != after.len() {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let mut field_before = Vec::new();
    let mut field_after = Vec::new();
    field_before
        .try_reserve_exact(field_indices.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: field_indices.len(),
            path,
        })?;
    field_after
        .try_reserve_exact(field_indices.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: field_indices.len(),
            path,
        })?;
    for (index, field) in object.archive_info.message_infos[message_index]
        .field_infos
        .iter()
        .enumerate()
    {
        if field.path.as_slice() != field_path
            && field
                .object_references
                .iter()
                .any(|identifier| before_set.contains(identifier) != after_set.contains(identifier))
        {
            return Err(BodyTableAppearanceError::InvalidSource { path });
        }
        if field_indices.binary_search(&index).is_err() {
            continue;
        }
        if field
            .r#type
            .is_some_and(|kind| kind != FieldType::ObjectReference)
        {
            return Err(BodyTableAppearanceError::InvalidSource { path });
        }
        let expected_before = field_references.map(|(before, _)| before).unwrap_or(before);
        let expected_after = field_references.map(|(_, after)| after).unwrap_or(after);
        if field.object_references != expected_before {
            return Err(BodyTableAppearanceError::InvalidSource { path });
        }
        field_before.push(copy_identifiers(&field.object_references, 0, path)?);
        field_after.push(copy_identifiers(expected_after, 0, path)?);
    }
    let mut transitions = Vec::new();
    transitions
        .try_reserve_exact(field_indices.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: field_indices.len(),
            path,
        })?;
    for index in 0..field_indices.len() {
        transitions.push(FieldObjectReferenceTransition {
            field_info_index: field_indices[index],
            expected_path: field_path,
            before: field_before[index].as_slice(),
            after: field_after[index].as_slice(),
        });
    }
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            message_index,
            message,
            ObjectReferenceTransition {
                aggregate_before: before,
                aggregate_after: after,
                fields: transitions.as_slice(),
            },
            limits,
        )
        .map_err(|error| map_core_error(error, path))?;
    Ok(())
}

fn resolve_target_with_budget<'table>(
    source: &Package,
    selector: impl Into<BodyTableSelector<'table>>,
    budget: &mut TransactionBudget,
    require_metadata: bool,
) -> Result<BodyTableAppearanceTarget, BodyTableAppearanceError> {
    // The physical catalog and the semantic resolver must debit the same
    // operation ledger.  `TransactionBudget::new` charges the first source;
    // this call is pointer-deduplicated for that source and also charges a
    // reopened candidate before its graph is traversed.
    budget
        .wire
        .charge_source_catalog(&source.state.source)
        .map_err(|error| map_lock_error(error, BodyTableAppearancePath::Package))?;
    budget.charge_transaction_work(
        source.source_bytes().len(),
        BodyTableAppearancePath::Package,
    )?;
    let native = source
        .resolve_body_table_with_budget(selector.into(), &mut budget.wire)
        .map_err(|error| map_lock_error(error, BodyTableAppearancePath::Package))?;
    let path = BodyTableAppearancePath::Table {
        table: native.table_position,
    };
    table_lock::validate_body_table_target(source, &native, &mut budget.wire)
        .map_err(|error| map_lock_error(error, path))?;
    ensure_unique_physical_identifier(source, native.model_identifier.get(), path, budget)?;
    let payload = model_payload(source, &native, budget)?;
    let (style_identifier, appearance) = resolve_codec_appearance(
        source,
        payload,
        native.component_index,
        native.model_identifier.get(),
        native.model_message_type,
        path,
        budget,
        require_metadata,
    )?;
    Ok(BodyTableAppearanceTarget {
        native,
        appearance,
        style_identifier,
    })
}

fn resolve_at_with_budget(
    source: &Package,
    table: usize,
    budget: &mut TransactionBudget,
    require_metadata: bool,
) -> Result<BodyTableAppearanceTarget, BodyTableAppearanceError> {
    resolve_target_with_budget(
        source,
        BodyTableSelector::index(table),
        budget,
        require_metadata,
    )
}

fn model_payload<'a>(
    source: &'a Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut TransactionBudget,
) -> Result<&'a [u8], BodyTableAppearanceError> {
    let path = BodyTableAppearancePath::Table {
        table: target.table_position,
    };
    let object = source
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .and_then(|component| component.archive().objects.get(target.model_object_index))
        .filter(|object| object.archive_info.identifier == Some(target.model_identifier.get()))
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    validate_appearance_message_roles(object, target.model_message_type, path, budget)?;
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    Ok(message.data.as_slice())
}

fn resolve_codec_appearance(
    source: &Package,
    model_payload: &[u8],
    model_component: usize,
    model_identifier: u64,
    model_message_type: u32,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
    require_metadata: bool,
) -> Result<(Option<u64>, Appearance), BodyTableAppearanceError> {
    // The model projection itself is part of the one aggregate wire budget.
    // DecodeReport is consumed even when the semantic projection is empty.
    let (model, model_report) =
        codec::decode_table_model_with_report(model_payload, budget.codec_options(model_payload))
            .map_err(|error| map_codec_error(error, path))?;
    budget.consume_report(model_report, path)?;
    let direct_style_identifier = model.style_identifier();
    validate_model_archive_metadata(
        source,
        model_identifier,
        model_message_type,
        direct_style_identifier,
        model.style_preset_identifier(),
        path,
        budget,
    )?;
    // A nonzero direct edge is authoritative. Read-only discovery may fall
    // back through preset -> network -> style, while the transaction target
    // remains `None` so changed writes on an inherited/preset route fail
    // closed before allocation.
    let mut routed_identifiers = [0u64; 2];
    let mut routed_identifier_count = 0usize;
    let style_identifier = if direct_style_identifier != 0 {
        direct_style_identifier
    } else if let Some(preset_identifier) = model.style_preset_identifier() {
        routed_identifiers[0] = preset_identifier;
        ensure_unique_physical_identifier(source, preset_identifier, path, budget)?;
        let preset_payload = appearance_message_data(
            source,
            preset_identifier,
            TABLE_STYLE_PRESET_MESSAGE_TYPE,
            path,
            budget,
        )?;
        let (preset, preset_report) = codec::decode_table_style_preset_with_report(
            preset_payload,
            budget.codec_options(preset_payload),
        )
        .map_err(|error| map_codec_error(error, path))?;
        budget.consume_report(preset_report, path)?;
        let network_identifier = preset
            .style_network_identifier()
            .filter(|identifier| *identifier != 0)
            .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
        routed_identifiers[1] = network_identifier;
        routed_identifier_count = 2;
        ensure_unique_physical_identifier(source, network_identifier, path, budget)?;
        let network_payload = appearance_message_data(
            source,
            network_identifier,
            TABLE_STYLE_NETWORK_MESSAGE_TYPE,
            path,
            budget,
        )?;
        let (network, network_report) = codec::decode_table_style_network_with_report(
            network_payload,
            budget.codec_options(network_payload),
        )
        .map_err(|error| map_codec_error(error, path))?;
        budget.consume_report(network_report, path)?;
        network.table_style_identifier()
    } else {
        if require_metadata {
            let metadata = metadata_payload(source, path, budget)?;
            validate_metadata_graph_registries(
                source,
                metadata,
                model_component,
                model_identifier,
                &[],
                &[],
                path,
                budget,
            )?;
        }
        return Ok((None, Appearance::default()));
    };
    let mut nodes = Vec::new();
    budget.charge_allocations(8, path)?;
    let mut current = Some(style_identifier);
    let mut stylesheet = None;
    for _ in 0..MAX_INHERITANCE_DEPTH {
        let Some(identifier) = current else { break };
        ensure_unique_physical_identifier(source, identifier, path, budget)?;
        budget.charge_styles(1, path)?;
        if nodes
            .iter()
            .any(|node: &codec::TableStyleNode<'_>| node.identifier() == identifier)
        {
            return Err(BodyTableAppearanceError::InvalidSource { path });
        }
        let style_payload = style_message_data(source, identifier, path, budget)?;
        let (style, style_report) = codec::decode_table_style_with_report(
            style_payload,
            budget.codec_options(style_payload),
        )
        .map_err(|error| map_codec_error(error, path))?;
        budget.consume_report(style_report, path)?;
        let style_sheet = style
            .stylesheet_identifier()
            .filter(|identifier| *identifier != 0)
            .ok_or(BodyTableAppearanceError::UnsupportedDependency { path })?;
        if style_sheet == 0 {
            return Err(BodyTableAppearanceError::UnsupportedDependency { path });
        }
        if let Some(previous) = stylesheet {
            if previous != style_sheet {
                return Err(BodyTableAppearanceError::UnsupportedDependency { path });
            }
        } else {
            stylesheet = Some(style_sheet);
        }
        current = style.parent_identifier();
        nodes
            .try_reserve(1)
            .map_err(|_| BodyTableAppearanceError::Allocation {
                amount: nodes.len().saturating_add(1),
                path,
            })?;
        nodes.push(codec::TableStyleNode::new(identifier, style));
    }
    if current.is_some() {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let stylesheet_identifier =
        stylesheet.ok_or(BodyTableAppearanceError::UnsupportedDependency { path })?;
    validate_global_style_inbound(
        source,
        style_identifier,
        model_identifier,
        model_message_type,
        path,
        budget,
    )?;
    let style_location = resolved_location(source, style_identifier, path, budget)?;
    let stylesheet_location = resolved_location(source, stylesheet_identifier, path, budget)?;
    ensure_unique_physical_identifier(source, stylesheet_identifier, path, budget)?;
    let _stylesheet_payload = stylesheet_message_data(source, stylesheet_identifier, path, budget)?;
    if style_location.0 != stylesheet_location.0 {
        return Err(BodyTableAppearanceError::UnsupportedDependency { path });
    }
    if require_metadata {
        let metadata = metadata_payload(source, path, budget)?;
        validate_metadata_graph_registries(
            source,
            metadata,
            model_component,
            model_identifier,
            &routed_identifiers[..routed_identifier_count],
            &nodes[1..],
            path,
            budget,
        )?;
        let _ =
            preflight_metadata_for_style(source, model_component, style_identifier, path, budget)?;
    }
    budget.charge_fields(nodes.len(), path)?;
    budget.charge_work(nodes.len().saturating_mul(model_payload.len()), path)?;
    let effective = codec::resolve_table_style_appearance(
        &nodes,
        style_identifier,
        budget.codec_options(model_payload),
    )
    .map_err(|error| map_codec_error(error, path))?;
    Ok((
        (direct_style_identifier != 0).then_some(direct_style_identifier),
        snapshot_appearance(effective),
    ))
}

fn validate_model_archive_metadata(
    source: &Package,
    model_identifier: u64,
    model_message_type: u32,
    style_identifier: u64,
    preset_identifier: Option<u64>,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let (component_index, object_index) =
        resolved_location(source, model_identifier, path, budget)?;
    let object = source
        .state
        .source
        .components()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(object_index))
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    validate_appearance_message_roles(object, model_message_type, path, budget)?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == model_message_type)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    let nested_references = info.field_infos.iter().try_fold(0usize, |count, field| {
        count
            .checked_add(field.object_references.len())
            .and_then(|value| value.checked_add(field.data_references.len()))
    });
    let nested_references =
        nested_references.ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    budget.charge_fields(info.field_infos.len(), path)?;
    budget.charge_references(
        info.object_references
            .len()
            .saturating_add(info.data_references.len())
            .saturating_add(nested_references),
        path,
    )?;
    budget.charge_transaction_work(
        info.field_infos
            .iter()
            .map(|field| field.path.as_slice().len())
            .sum::<usize>()
            .saturating_add(info.field_infos.len()),
        path,
    )?;
    let mut expected = [0u64; 2];
    let mut expected_len = 0usize;
    if style_identifier != 0 {
        expected[expected_len] = style_identifier;
        expected_len += 1;
    }
    if let Some(preset_identifier) = preset_identifier.filter(|identifier| *identifier != 0) {
        expected[expected_len] = preset_identifier;
        expected_len += 1;
    }
    if info.object_references.as_slice() != &expected[..expected_len] {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    if !info.data_references.is_empty() {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    let mut style_fields = 0usize;
    let mut preset_fields = 0usize;
    for field in &info.field_infos {
        if !field.data_references.is_empty() {
            return Err(BodyTableAppearanceError::InvalidSource { path });
        }
        let is_style_path = field.path.as_slice() == [3];
        let is_preset_path = field.path.as_slice() == [48];
        if is_style_path {
            style_fields = style_fields.saturating_add(1);
        }
        if is_preset_path {
            preset_fields = preset_fields.saturating_add(1);
        }
        if (is_style_path || is_preset_path)
            && field
                .r#type
                .is_some_and(|kind| kind != FieldType::ObjectReference)
        {
            return Err(BodyTableAppearanceError::InvalidSource { path });
        }
        if field.object_references.is_empty() && !is_style_path && !is_preset_path {
            continue;
        }
        let valid_style = style_identifier != 0
            && is_style_path
            && field.object_references.as_slice() == [style_identifier];
        let valid_preset = preset_identifier.is_some_and(|preset_identifier| {
            is_preset_path && field.object_references.as_slice() == [preset_identifier]
        });
        if !valid_style && !valid_preset {
            return Err(BodyTableAppearanceError::InvalidSource { path });
        }
    }
    if style_fields > 1 || preset_fields > 1 {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    Ok(())
}

fn style_message_data<'source>(
    source: &'source Package,
    identifier: u64,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<&'source [u8], BodyTableAppearanceError> {
    appearance_message_data(source, identifier, TABLE_STYLE_MESSAGE_TYPE, path, budget)
}

fn validate_global_style_inbound(
    source: &Package,
    style_identifier: u64,
    selected_model_identifier: u64,
    selected_model_type: u32,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let physical = physical_source(source)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|error| map_archive_error(error, path))?;
    let object_count = source
        .state
        .source
        .components()
        .iter()
        .try_fold(0usize, |count, component| {
            count.checked_add(component.archive().objects.len())
        })
        .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
    budget.charge_allocations(object_count, path)?;
    budget.charge_transaction_work(object_count, path)?;
    let mut physical_identifiers = HashSet::new();
    physical_identifiers
        .try_reserve(object_count)
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: object_count,
            path,
        })?;
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
            if !physical_identifiers.insert(identifier) {
                return Err(BodyTableAppearanceError::InvalidSource { path });
            }
        }
    }
    let mut authority = GlobalReferenceAuthority {
        physical_identifiers: &physical_identifiers,
        style_identifier,
        invalid: false,
    };
    for component in source.state.source.components().iter() {
        budget.charge_components(1, path)?;
        for object in &component.archive().objects {
            let object_identifier = object
                .archive_info
                .identifier
                .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
            if object.messages.len() != object.archive_info.message_infos.len() {
                return Err(BodyTableAppearanceError::InvalidSource { path });
            }
            let mut object_fields = 0usize;
            let mut object_references = 0usize;
            let mut object_work = 1usize;
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                let nested_object_references = info
                    .field_infos
                    .iter()
                    .try_fold(0usize, |count, field| {
                        count.checked_add(field.object_references.len())
                    })
                    .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
                let nested_data_references = info
                    .field_infos
                    .iter()
                    .try_fold(0usize, |count, field| {
                        count.checked_add(field.data_references.len())
                    })
                    .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
                let field_path_items = info
                    .field_infos
                    .iter()
                    .try_fold(0usize, |count, field| {
                        count.checked_add(field.path.path.len())
                    })
                    .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
                object_fields = object_fields
                    .checked_add(info.field_infos.len())
                    .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
                object_references = object_references
                    .checked_add(info.object_references.len())
                    .and_then(|count| count.checked_add(info.data_references.len()))
                    .and_then(|count| count.checked_add(nested_object_references))
                    .and_then(|count| count.checked_add(nested_data_references))
                    .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
                object_work = object_work
                    .checked_add(message.data.len())
                    .and_then(|work| work.checked_add(info.object_references.len()))
                    .and_then(|work| work.checked_add(info.data_references.len()))
                    .and_then(|work| work.checked_add(info.field_infos.len()))
                    .and_then(|work| work.checked_add(field_path_items))
                    .and_then(|work| work.checked_add(nested_object_references))
                    .and_then(|work| work.checked_add(nested_data_references))
                    .ok_or(BodyTableAppearanceError::InvalidSource { path })?;
            }
            budget.charge_fields(object_fields, path)?;
            budget.charge_references(object_references, path)?;
            budget.charge_transaction_work(object_work, path)?;
            object
                .inspect_references_with_policy_and_limits(
                    &mut authority,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(|error| map_core_error(error, path))?;
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                let aggregate_occurrences = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == style_identifier)
                    .count();
                let field_occurrences = info
                    .field_infos
                    .iter()
                    .flat_map(|field| &field.object_references)
                    .filter(|identifier| **identifier == style_identifier)
                    .count();
                if aggregate_occurrences == 0 && field_occurrences == 0 {
                    continue;
                }
                if aggregate_occurrences != 1 || field_occurrences > 1 {
                    return Err(BodyTableAppearanceError::InvalidSource { path });
                }
                let selected_model = object_identifier == selected_model_identifier
                    && message.type_ == selected_model_type;
                let known_style_role = matches!(
                    message.type_,
                    6_001
                        | TABLE_STYLE_MESSAGE_TYPE
                        | TABLE_STYLE_NETWORK_MESSAGE_TYPE
                        | STYLESHEET_MESSAGE_TYPE
                );
                if !selected_model && !known_style_role {
                    return Err(BodyTableAppearanceError::UnsupportedDependency { path });
                }
            }
        }
    }
    if authority.invalid {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    Ok(())
}

fn ensure_unique_physical_identifier(
    source: &Package,
    identifier: u64,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let mut matches = 0usize;
    for component in source.state.source.components().iter() {
        budget.charge_components(1, path)?;
        budget.charge_transaction_work(component.archive().objects.len(), path)?;
        matches = matches.saturating_add(
            component
                .archive()
                .objects
                .iter()
                .filter(|object| object.archive_info.identifier == Some(identifier))
                .count(),
        );
    }
    if matches != 1 {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    Ok(())
}

fn physical_source(source: &Package) -> Result<&SourceCatalog, BodyTableAppearanceError> {
    Ok(&source.state.source)
}

fn preview_names(
    catalog: &SourceCatalog,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<Vec<&'static str>, BodyTableAppearanceError> {
    budget.charge_allocations(ROOT_PREVIEW_NAMES.len(), path)?;
    budget.charge_transaction_work(
        catalog
            .package()
            .len()
            .saturating_mul(ROOT_PREVIEW_NAMES.len()),
        path,
    )?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(ROOT_PREVIEW_NAMES.len())
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: ROOT_PREVIEW_NAMES.len(),
            path,
        })?;
    for name in ROOT_PREVIEW_NAMES {
        if catalog.package().iter().any(|entry| entry.name() == name) {
            names.push(name);
        }
    }
    Ok(names)
}

fn root_preview_deletions(
    catalog: &SourceCatalog,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<Vec<&'static str>, BodyTableAppearanceError> {
    let names = preview_names(catalog, path, budget)?;
    if names.len() != ROOT_PREVIEW_NAMES.len() {
        return Err(BodyTableAppearanceError::InvalidSource { path });
    }
    Ok(names)
}

fn verify_candidate_locality(
    source: &Package,
    candidate: &Package,
    target: BodyTableAppearanceTarget,
    old_style: Option<u64>,
    new_style: Option<u64>,
    source_previews: usize,
    target_previews: usize,
    path: BodyTableAppearancePath,
    budget: &mut TransactionBudget,
) -> Result<(), BodyTableAppearanceError> {
    let source_catalog = physical_source(source)?;
    let candidate_catalog = physical_source(candidate)?;
    budget.charge_transaction_work(
        source
            .source_bytes()
            .len()
            .saturating_add(candidate.source_bytes().len()),
        path,
    )?;
    budget.charge_allocations(
        ROOT_PREVIEW_NAMES.len().saturating_mul(2).saturating_add(8),
        path,
    )?;
    let source_preview_names = preview_names(source_catalog, path, budget)
        .map_err(|_| BodyTableAppearanceError::Verification)?;
    let candidate_preview_names = preview_names(candidate_catalog, path, budget)
        .map_err(|_| BodyTableAppearanceError::Verification)?;
    if source_preview_names.len() != source_previews
        || candidate_preview_names.len() != target_previews
    {
        return Err(BodyTableAppearanceError::Verification);
    }

    let mut component_indices = Vec::new();
    component_indices
        .try_reserve_exact(3)
        .map_err(|_| BodyTableAppearanceError::Allocation { amount: 3, path })?;
    component_indices.push(target.native.component_index);
    if let Some(identifier) = old_style {
        component_indices.push(resolved_location(source, identifier, path, budget)?.0);
    }
    if let Some(identifier) = new_style {
        component_indices.push(resolved_location(candidate, identifier, path, budget)?.0);
    }
    let mut allowed_names = Vec::new();
    allowed_names
        .try_reserve_exact(component_indices.len().saturating_add(1))
        .map_err(|_| BodyTableAppearanceError::Allocation {
            amount: component_indices.len().saturating_add(1),
            path,
        })?;
    for component_index in component_indices {
        let Some(component) = source
            .state
            .source
            .components()
            .get_index(component_index)
            .or_else(|| {
                candidate
                    .state
                    .source
                    .components()
                    .get_index(component_index)
            })
        else {
            return Err(BodyTableAppearanceError::Verification);
        };
        if !allowed_names
            .iter()
            .any(|name: &String| name == component.name())
        {
            allowed_names.push(component.name().to_owned());
        }
    }
    if !allowed_names.iter().any(|name| name == METADATA_ENTRY_NAME) {
        allowed_names.push(METADATA_ENTRY_NAME.to_owned());
    }

    let mut before_entries = source_catalog
        .package()
        .iter()
        .filter(|entry| !source_preview_names.contains(&entry.name()));
    let mut after_entries = candidate_catalog
        .package()
        .iter()
        .filter(|entry| !candidate_preview_names.contains(&entry.name()));
    budget.charge_transaction_work(
        source_catalog
            .package()
            .len()
            .saturating_add(candidate_catalog.package().len())
            .saturating_mul(allowed_names.len().saturating_add(1)),
        path,
    )?;
    loop {
        match (before_entries.next(), after_entries.next()) {
            (Some(before), Some(after)) if before.name() == after.name() => {
                let changed = before.raw_name() != after.raw_name()
                    || before.data() != after.data()
                    || before.metadata() != after.metadata()
                    || before.raw_record().local_record() != after.raw_record().local_record()
                    || before.raw_record().compressed_data()
                        != after.raw_record().compressed_data();
                let expected = allowed_names.iter().any(|name| name == before.name());
                if changed != expected {
                    return Err(BodyTableAppearanceError::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(BodyTableAppearanceError::Verification),
        }
    }
    Ok(())
}

fn map_lock_error(
    error: table_lock::BodyTableLockError,
    path: BodyTableAppearancePath,
) -> BodyTableAppearanceError {
    use table_lock::BodyTableLockError as LockError;
    use table_lock::BodyTableLockLimitKind as LockLimit;

    match error {
        LockError::TableNotFound => BodyTableAppearanceError::TableNotFound,
        LockError::AmbiguousTableName => BodyTableAppearanceError::AmbiguousTableName,
        LockError::AmbiguousSelector => BodyTableAppearanceError::AmbiguousSelector,
        LockError::UnsupportedSource => BodyTableAppearanceError::UnsupportedSource,
        LockError::InvalidSource => BodyTableAppearanceError::InvalidSource { path },
        LockError::PatchConflict => BodyTableAppearanceError::PatchConflict,
        LockError::Verification => BodyTableAppearanceError::Verification,
        LockError::Allocation { amount } => BodyTableAppearanceError::Allocation { amount, path },
        LockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableAppearanceError::LimitExceeded {
            kind: match kind {
                LockLimit::InputBytes => BodyTableAppearanceLimitKind::InputBytes,
                LockLimit::OutputBytes => BodyTableAppearanceLimitKind::OutputBytes,
                LockLimit::Entries => BodyTableAppearanceLimitKind::Entries,
                LockLimit::EntryBytes => BodyTableAppearanceLimitKind::EntryBytes,
                LockLimit::TotalEntryBytes => BodyTableAppearanceLimitKind::TotalEntryBytes,
                LockLimit::PackageBytes => BodyTableAppearanceLimitKind::PackageBytes,
                LockLimit::PayloadBytes => BodyTableAppearanceLimitKind::PayloadBytes,
                LockLimit::TotalPayloadBytes => BodyTableAppearanceLimitKind::TotalPayloadBytes,
                LockLimit::PayloadObjects => BodyTableAppearanceLimitKind::PayloadObjects,
                LockLimit::PayloadMessages => BodyTableAppearanceLimitKind::PayloadMessages,
                LockLimit::PayloadItems => BodyTableAppearanceLimitKind::PayloadItems,
                LockLimit::PayloadReferences => BodyTableAppearanceLimitKind::PayloadReferences,
                LockLimit::WireBytes => BodyTableAppearanceLimitKind::WireBytes,
                LockLimit::WireFields => BodyTableAppearanceLimitKind::WireFields,
                LockLimit::WireNesting => BodyTableAppearanceLimitKind::WireNesting,
                LockLimit::WireWork => BodyTableAppearanceLimitKind::WireWork,
            },
            observed,
            maximum,
            path,
        },
    }
}

fn map_archive_error(
    error: litchi_iwa_archive::Error,
    path: BodyTableAppearancePath,
) -> BodyTableAppearanceError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => {
            let kind = match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    BodyTableAppearanceLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    BodyTableAppearanceLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => BodyTableAppearanceLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    BodyTableAppearanceLimitKind::PackageBytes
                },
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => {
                    BodyTableAppearanceLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    BodyTableAppearanceLimitKind::TotalEntryBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    BodyTableAppearanceLimitKind::PayloadBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    BodyTableAppearanceLimitKind::TotalPayloadBytes
                },
            };
            BodyTableAppearanceError::LimitExceeded {
                kind,
                observed,
                maximum,
                path,
            }
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyTableAppearanceError::Allocation { amount, path }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error, path),
        _ => BodyTableAppearanceError::InvalidSource { path },
    }
}

fn map_core_error(
    error: litchi_iwa_core::Error,
    path: BodyTableAppearancePath,
) -> BodyTableAppearanceError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => {
            let kind = match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    BodyTableAppearanceLimitKind::PayloadBytes
                },
                litchi_iwa_core::LimitKind::Objects => BodyTableAppearanceLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    BodyTableAppearanceLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    BodyTableAppearanceLimitKind::PayloadItems
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    BodyTableAppearanceLimitKind::WireNesting
                },
            };
            BodyTableAppearanceError::LimitExceeded {
                kind,
                observed: u64::try_from(observed).unwrap_or(u64::MAX),
                maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
                path,
            }
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyTableAppearanceError::Allocation {
                amount: requested,
                path,
            }
        },
        _ => BodyTableAppearanceError::InvalidSource { path },
    }
}
