//! Exact-source transactions for native Keynote table merge regions.
//!
//! The transaction owns only checked [`Region`] values.  Slide/table
//! selection, package authority, archive locality, and resource accounting
//! remain in the private table core; the shared Numbers wire adapter owns the
//! merge-owner rewrite.  No generated table model or native object identity
//! crosses this boundary.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::too_many_arguments,
    reason = "The package boundary deliberately redacts native failure detail."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{
    Error as ArchiveError, LimitKind as ArchiveLimitKind, SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{WireLimits, table::merge::Region};
use litchi_iwa_core::{RawMessage, SnappyStream};
use litchi_numbers_wire::table_merges::{self, ReadLimits};

use super::super::{Package, PayloadLimitKind, ReadError, SemanticLimitKind, slide_table_core};
use super::{
    MERGE_READER_ALLOCATIONS_PER_REGION, MERGE_READER_SCRATCH_PER_REGION, REGION_BYTES,
    SlideTableMergesError, SlideTableMergesLimitKind, SlideTableMergesPath,
};
use crate::{SlideSelector, slide::table::TableSelector};

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;

/// One immutable exact-source merge edit.
///
/// A single edit may stage several merge/unmerge calls.  The staged value is
/// always a checked region set, so a later call observes the result of the
/// earlier call without exposing a native formula index or object ID.
pub struct SlideTableMergesEdit<'a> {
    source: &'a Package,
    selection: MergeSelection,
    after: Vec<Region>,
}

impl fmt::Debug for SlideTableMergesEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableMergesEdit")
            .field("path", &self.selection.path())
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableMergesEdit<'_> {
    /// Return the selected slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.target.slide_position
    }

    /// Return the selected table position.
    #[must_use]
    pub const fn table_position(&self) -> Position {
        self.selection.target.table_position
    }

    /// Return the semantic location of this edit.
    #[must_use]
    pub const fn path(&self) -> SlideTableMergesPath {
        self.selection.path()
    }

    /// Return the source merge regions in source order.
    #[must_use]
    pub fn before(&self) -> &[Region] {
        &self.selection.before
    }

    /// Return the currently staged merge regions in semantic order.
    #[must_use]
    pub fn after(&self) -> &[Region] {
        &self.after
    }

    /// Stage one checked merge region.
    ///
    /// A region that already intersects a source or earlier staged region is
    /// rejected.  This keeps duplicate and partial-overlap requests aligned
    /// with native iWork merge semantics; an unmerge of a missing region is
    /// the explicit idempotent no-op operation.
    ///
    /// This changes merge geometry while preserving the underlying cell values.
    pub fn merge(&mut self, region: Region) -> Result<&mut Self, SlideTableMergesError> {
        self.selection
            .budget
            .work(1)
            .map_err(super::map_core_error)?;
        validate_region(&self.selection.target, region)?;
        if self.after.len() >= WireLimits::MAX_FIELDS {
            return Err(SlideTableMergesError::LimitExceeded {
                kind: SlideTableMergesLimitKind::PayloadItems,
                observed: self.after.len() as u64 + 1,
                maximum: WireLimits::MAX_FIELDS as u64,
            });
        }
        self.selection
            .budget
            .work(self.after.len())
            .map_err(super::map_core_error)?;
        if self.after.iter().any(|existing| existing.overlaps(region)) {
            return Err(SlideTableMergesError::InvalidRegion);
        }
        self.selection
            .budget
            .owned_value(REGION_BYTES)
            .map_err(super::map_core_error)?;
        self.after
            .try_reserve(1)
            .map_err(|_| SlideTableMergesError::Allocation {
                amount: REGION_BYTES,
            })?;
        self.after.push(region);
        Ok(self)
    }

    /// Stage removal of one exact merge region.
    ///
    /// Removing a region that is not present leaves the edit unchanged.  The
    /// resulting commit is byte-exact and reports `changed() == false` when
    /// no other operation was staged.
    pub fn unmerge(&mut self, region: Region) -> Result<bool, SlideTableMergesError> {
        self.selection
            .budget
            .work(1)
            .map_err(super::map_core_error)?;
        self.selection
            .budget
            .work(self.after.len())
            .map_err(super::map_core_error)?;
        if let Some(index) = self.after.iter().position(|existing| *existing == region) {
            // A region already present in the selected source (or staged by a
            // prior checked merge) is necessarily in bounds.  Avoid an
            // upfront bounds check so an absent request remains the native
            // idempotent no-op, including outside the table.
            self.after.remove(index);
            return Ok(true);
        }
        Ok(false)
    }

    /// Publish the checked edit after candidate reopen and locality proofs.
    pub fn commit(self) -> Result<SlideTableMergesCommit, SlideTableMergesError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible merge patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTableMergesPatch {
    artifacts: ExactArtifacts,
    selection: MergeSelection,
    before: Vec<Region>,
    after: Vec<Region>,
    touched_components: usize,
    deleted_previews: usize,
    source_previews_absent: bool,
    target_previews_absent: bool,
}

impl fmt::Debug for SlideTableMergesPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableMergesPatch")
            .field("path", &self.selection.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableMergesPatch {
    /// Return the semantic location guarded by this patch.
    #[must_use]
    pub const fn path(&self) -> SlideTableMergesPath {
        self.selection.path()
    }

    /// Return the exact source regions.
    #[must_use]
    pub fn before(&self) -> &[Region] {
        &self.before
    }

    /// Return the exact target regions.
    #[must_use]
    pub fn after(&self) -> &[Region] {
        &self.after
    }

    /// Return the source artifact fingerprint guarded by this patch.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target artifact fingerprint carried by this patch.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether applying this patch leaves the source bytes unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Reverse the exact source and target artifacts and semantic states.
    #[must_use]
    pub fn inverse(&self) -> Self {
        let mut selection = self.selection.clone();
        selection.before = self.after.clone();
        Self {
            artifacts: self.artifacts.inverse(),
            selection,
            before: self.after.clone(),
            after: self.before.clone(),
            touched_components: self.touched_components,
            deleted_previews: self.deleted_previews,
            source_previews_absent: self.target_previews_absent,
            target_previews_absent: self.source_previews_absent,
        }
    }
}

/// Compact diagnostics for one published merge transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTableMergesDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideTableMergesDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Return whether the candidate changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of invalidated root previews.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the candidate was reparsed before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one merge transaction.
#[must_use = "a Keynote merge commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTableMergesCommit {
    package: Package,
    patch: SlideTableMergesPatch,
    diagnostics: SlideTableMergesDiagnostics,
}

impl SlideTableMergesCommit {
    /// Borrow the verified package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its verified package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideTableMergesPatch {
        &self.patch
    }

    /// Borrow transaction diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableMergesDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq, Eq)]
struct MergeSelection {
    target: slide_table_core::Target,
    before: Vec<Region>,
    budget: slide_table_core::Budget,
}

impl fmt::Debug for MergeSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("MergeSelection")
            .field("path", &self.path())
            .field("before", &self.before)
            .field("locked", &self.target.locked)
            .finish_non_exhaustive()
    }
}

impl MergeSelection {
    const fn path(&self) -> SlideTableMergesPath {
        SlideTableMergesPath::Table {
            slide: self.target.slide_position,
            table: self.target.table_position,
        }
    }
}

impl Package {
    /// Begin an immutable exact edit of one selected slide-table merge set.
    pub fn edit_slide_table_merges<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<SlideTableMergesEdit<'_>, SlideTableMergesError> {
        let mut budget = slide_table_core::Budget::new(self).map_err(super::map_core_error)?;
        let mut selection = select_merges(self, slide.into(), table.into(), &mut budget)?;
        let after = clone_regions(&selection.before, &mut budget)?;
        selection.budget = budget;
        Ok(SlideTableMergesEdit {
            source: self,
            selection,
            after,
        })
    }

    /// Apply an exact-source checked reversible slide-table merge patch.
    pub fn apply_slide_table_merges(
        &self,
        patch: &SlideTableMergesPatch,
    ) -> Result<SlideTableMergesCommit, SlideTableMergesError> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideTableMergesError::PatchConflict);
        }
        let mut budget = slide_table_core::Budget::new(self).map_err(super::map_core_error)?;
        if previews_absent(self, &mut budget)? != patch.source_previews_absent {
            return Err(SlideTableMergesError::PatchConflict);
        }
        let current = select_merges(
            self,
            SlideSelector::position(patch.selection.target.slide_position),
            TableSelector::position(patch.selection.target.table_position),
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(SlideTableMergesError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(SlideTableMergesCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTableMergesDiagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch, budget)
    }
}

fn select_merges(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    budget: &mut slide_table_core::Budget,
) -> Result<MergeSelection, SlideTableMergesError> {
    let target = slide_table_core::select_table(package, slide, table, budget)
        .map_err(super::map_core_error)?;
    let before = package.decode_selected_merges(target.clone(), budget)?;
    Ok(MergeSelection {
        target,
        before,
        budget: *budget,
    })
}

fn clone_regions(
    regions: &[Region],
    budget: &mut slide_table_core::Budget,
) -> Result<Vec<Region>, SlideTableMergesError> {
    let bytes = regions
        .len()
        .checked_mul(REGION_BYTES)
        .ok_or(slide_table_core::Error::InvalidSource)
        .map_err(super::map_core_error)?;
    budget.owned_value(bytes).map_err(super::map_core_error)?;
    let mut clone = Vec::new();
    clone
        .try_reserve_exact(regions.len())
        .map_err(|_| SlideTableMergesError::Allocation { amount: bytes })?;
    clone.extend_from_slice(regions);
    Ok(clone)
}

fn validate_region(
    target: &slide_table_core::Target,
    region: Region,
) -> Result<(), SlideTableMergesError> {
    if region.end_row() >= target.rows || region.end_column() >= target.columns {
        return Err(SlideTableMergesError::InvalidRegion);
    }
    Ok(())
}

fn normalize_regions(
    target: &slide_table_core::Target,
    before: &[Region],
    desired: &[Region],
    budget: &mut slide_table_core::Budget,
) -> Result<Vec<Region>, SlideTableMergesError> {
    if desired.len() > WireLimits::MAX_FIELDS {
        return Err(SlideTableMergesError::LimitExceeded {
            kind: SlideTableMergesLimitKind::PayloadItems,
            observed: desired.len() as u64,
            maximum: WireLimits::MAX_FIELDS as u64,
        });
    }
    let mut normalized: Vec<Region> = Vec::new();
    let bytes = desired
        .len()
        .checked_mul(REGION_BYTES)
        .ok_or(slide_table_core::Error::InvalidSource)
        .map_err(super::map_core_error)?;
    budget.owned_value(bytes).map_err(super::map_core_error)?;
    normalized
        .try_reserve_exact(desired.len())
        .map_err(|_| SlideTableMergesError::Allocation { amount: bytes })?;
    for region in before {
        budget.work(desired.len()).map_err(super::map_core_error)?;
        if desired.contains(region) {
            validate_region(target, *region)?;
            if normalized.iter().any(|existing| existing.overlaps(*region)) {
                return Err(SlideTableMergesError::InvalidSource);
            }
            normalized.push(*region);
        }
    }
    for region in desired {
        budget.work(before.len()).map_err(super::map_core_error)?;
        validate_region(target, *region)?;
        if before.contains(region) {
            continue;
        }
        if normalized.iter().any(|existing| existing.overlaps(*region)) {
            return Err(SlideTableMergesError::InvalidRegion);
        }
        normalized.push(*region);
    }
    Ok(normalized)
}

fn commit_edit(
    source: &Package,
    selection: &MergeSelection,
    desired: Vec<Region>,
) -> Result<SlideTableMergesCommit, SlideTableMergesError> {
    let mut budget = selection.budget;
    let desired = normalize_regions(&selection.target, &selection.before, &desired, &mut budget)?;
    if selection.before == desired {
        let source_previews_absent = previews_absent(source, &mut budget)?;
        let bytes = shared_source_artifact(source, &mut budget)?;
        return Ok(SlideTableMergesCommit {
            package: source.snapshot(),
            patch: SlideTableMergesPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: selection.before.clone(),
                after: desired,
                touched_components: 0,
                deleted_previews: 0,
                source_previews_absent,
                target_previews_absent: source_previews_absent,
            },
            diagnostics: SlideTableMergesDiagnostics::unchanged(),
        });
    }
    let source_previews_absent = previews_absent(source, &mut budget)?;
    let (candidate, deleted_previews) = rewrite_merges(source, selection, &desired, &mut budget)?;
    if !previews_absent(&candidate, &mut budget)? {
        return Err(SlideTableMergesError::Verification);
    }
    let mut reopen_budget = budget;
    let selected = select_merges(
        &candidate,
        SlideSelector::position(selection.target.slide_position),
        TableSelector::position(selection.target.table_position),
        &mut reopen_budget,
    )?;
    if !same_selection(&selected, selection) || selected.before != desired {
        return Err(SlideTableMergesError::Verification);
    }
    slide_table_core::verify_locality(
        source,
        &candidate,
        &selection.target,
        true,
        &mut reopen_budget,
    )
    .map_err(super::map_core_error)?;
    let mut artifact_budget = reopen_budget;
    let target = shared_source_artifact(&candidate, &mut artifact_budget)?;
    let source_artifact = shared_source_artifact(source, &mut artifact_budget)?;
    Ok(SlideTableMergesCommit {
        package: candidate,
        patch: SlideTableMergesPatch {
            artifacts: ExactArtifacts::new(source_artifact, target),
            selection: selection.clone(),
            before: selection.before.clone(),
            after: desired,
            touched_components: 1,
            deleted_previews,
            source_previews_absent,
            target_previews_absent: true,
        },
        diagnostics: SlideTableMergesDiagnostics::published(deleted_previews),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &SlideTableMergesPatch,
    mut budget: slide_table_core::Budget,
) -> Result<SlideTableMergesCommit, SlideTableMergesError> {
    let target_source = patch.artifacts.target();
    let candidate = parse_candidate(Arc::clone(&target_source), source, &mut budget)?;
    if previews_absent(&candidate, &mut budget)? != patch.target_previews_absent {
        return Err(SlideTableMergesError::Verification);
    }
    let selected = select_merges(
        &candidate,
        SlideSelector::position(patch.selection.target.slide_position),
        TableSelector::position(patch.selection.target.table_position),
        &mut budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(SlideTableMergesError::Verification);
    }
    slide_table_core::verify_locality(
        source,
        &candidate,
        &patch.selection.target,
        patch.target_previews_absent,
        &mut budget,
    )
    .map_err(super::map_core_error)?;
    Ok(SlideTableMergesCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideTableMergesDiagnostics::published(patch.deleted_previews),
    })
}

fn rewrite_merges(
    source: &Package,
    selection: &MergeSelection,
    desired: &[Region],
    budget: &mut slide_table_core::Budget,
) -> Result<(Package, usize), SlideTableMergesError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.target.model.component.as_ref())
        .ok_or(SlideTableMergesError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableMergesError::UnsupportedSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    // Reuse the core's bounded component reader for this second temporary
    // archive.  It preflights Snappy and archive allocations, then charges the
    // actual object/message/field/reference inventory and validates canonical
    // framing before the mutable archive is inspected below.
    let mut archive = slide_table_core::component_archive(
        source,
        selection.target.model.component.as_ref(),
        budget,
    )
    .map_err(super::map_core_error)?;
    let original_source = archive
        .object(selection.target.model.identifier)
        .ok_or(SlideTableMergesError::InvalidSource)?
        .messages
        .get(selection.target.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(SlideTableMergesError::InvalidSource)?
        .data
        .as_slice();
    budget
        .allocations(original_source.len())
        .map_err(super::map_core_error)?;
    budget
        .retained(original_source.len())
        .map_err(super::map_core_error)?;
    let mut original = Vec::new();
    original
        .try_reserve_exact(original_source.len())
        .map_err(|_| SlideTableMergesError::Allocation {
            amount: original_source.len(),
        })?;
    original.extend_from_slice(original_source);

    let wire = budget.residual(source).map_err(super::map_core_error)?;
    let max_regions = desired
        .len()
        .max(selection.before.len())
        .clamp(1, WireLimits::MAX_FIELDS);
    let limits = ReadLimits {
        wire,
        max_regions,
        max_overlap_checks: budget.remaining_work().map_err(super::map_core_error)?,
    };
    let write = table_merges::rewrite_table_merges(
        &original,
        desired,
        litchi_core::id::generate_guid_bytes(),
        limits,
    )
    .map_err(|error| {
        super::charge_attempted(budget, &error);
        super::map_merge_error(error)
    })?;
    budget
        .input(write.report.input_bytes())
        .map_err(super::map_core_error)?;
    budget
        .fields(write.report.fields())
        .map_err(super::map_core_error)?;
    budget
        .work(write.report.work())
        .map_err(super::map_core_error)?;
    if !write.changed {
        return Err(SlideTableMergesError::Verification);
    }
    let rewritten = write.data;
    budget
        .output(rewritten.len())
        .map_err(super::map_core_error)?;
    budget
        .allocations(rewritten.len())
        .map_err(super::map_core_error)?;
    budget
        .retained(rewritten.len())
        .map_err(super::map_core_error)?;
    budget
        .scratch(rewritten.len())
        .map_err(super::map_core_error)?;

    let verify_limits = ReadLimits {
        wire: budget.residual(source).map_err(super::map_core_error)?,
        max_regions,
        max_overlap_checks: budget.remaining_work().map_err(super::map_core_error)?,
    };
    let verified = table_merges::read_table_merges(&rewritten, verify_limits).map_err(|error| {
        super::charge_attempted(budget, &error);
        super::map_merge_error(error)
    })?;
    charge_read(budget, &verified)?;
    if verified.regions != desired {
        return Err(SlideTableMergesError::Verification);
    }

    archive
        .object_mut(selection.target.model.identifier)
        .ok_or(SlideTableMergesError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.target.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_native_error)?;
    let encoded_bound = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_native_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_bound).map_err(map_native_error)?;
    budget
        .output(
            encoded_bound
                .checked_add(compressed_bound)
                .ok_or(slide_table_core::Error::InvalidSource)
                .map_err(super::map_core_error)?,
        )
        .map_err(super::map_core_error)?;
    budget
        .allocations(encoded_bound)
        .map_err(super::map_core_error)?;
    budget
        .retained(encoded_bound)
        .map_err(super::map_core_error)?;
    budget
        .allocations(compressed_bound)
        .map_err(super::map_core_error)?;
    budget
        .retained(compressed_bound)
        .map_err(super::map_core_error)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_native_error)?;
    if bytes.len() != encoded_bound {
        return Err(SlideTableMergesError::Verification);
    }
    let compressed = SnappyStream::compress(&bytes).map_err(map_native_error)?;
    if compressed.len() > compressed_bound {
        return Err(SlideTableMergesError::Verification);
    }
    slide_table_core::charge_preview_scan(source, budget).map_err(super::map_core_error)?;
    let previews = super::super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| SlideTableMergesError::InvalidSource)?;
    let edits = [EntryEdit::new(
        selection.target.model.component.as_ref(),
        &compressed,
    )];
    charge_reassembly_prepare(catalog, compressed.len(), previews.names(), budget)?;
    let prepared = catalog
        .prepare_reassembly_with_deletions(&edits, previews.names(), physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .reassembly(requirements)
        .map_err(super::map_core_error)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate = parse_candidate(output.into(), source, budget)?;
    Ok((candidate, previews.len()))
}

fn charge_read(
    budget: &mut slide_table_core::Budget,
    read: &table_merges::MergeRead,
) -> Result<(), SlideTableMergesError> {
    budget
        .input(read.report.input_bytes())
        .map_err(super::map_core_error)?;
    budget
        .fields(read.report.fields())
        .map_err(super::map_core_error)?;
    budget
        .work(read.report.work())
        .map_err(super::map_core_error)?;
    let region_count = read.regions.len();
    let allocations = region_count
        .checked_mul(MERGE_READER_ALLOCATIONS_PER_REGION)
        .ok_or(slide_table_core::Error::InvalidSource)
        .map_err(super::map_core_error)?;
    budget
        .allocations(allocations)
        .map_err(super::map_core_error)?;
    let scratch = region_count
        .checked_mul(MERGE_READER_SCRATCH_PER_REGION)
        .ok_or(slide_table_core::Error::InvalidSource)
        .map_err(super::map_core_error)?;
    budget.scratch(scratch).map_err(super::map_core_error)?;
    let retained = region_count
        .checked_mul(REGION_BYTES)
        .ok_or(slide_table_core::Error::InvalidSource)
        .map_err(super::map_core_error)?;
    budget.retained(retained).map_err(super::map_core_error)
}

fn shared_source_artifact(
    package: &Package,
    budget: &mut slide_table_core::Budget,
) -> Result<Arc<[u8]>, SlideTableMergesError> {
    let source = physical_catalog(package)?;
    budget
        .artifact(source.package().source_bytes().len(), 0)
        .map_err(super::map_core_error)?;
    Ok(source.shared_source())
}

fn parse_candidate(
    source: Arc<[u8]>,
    original: &Package,
    budget: &mut slide_table_core::Budget,
) -> Result<Package, SlideTableMergesError> {
    let source_bytes = source.len();
    budget
        .preflight_candidate(original, source_bytes)
        .map_err(super::map_core_error)?;
    budget
        .preflight_semantic_scan(original)
        .map_err(super::map_core_error)?;
    let candidate = Package::from_source_with_options(source, original.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    Ok(candidate)
}

fn charge_reassembly_prepare(
    catalog: &SourceCatalog,
    edited_bytes: usize,
    deleted_names: &[&str],
    budget: &mut slide_table_core::Budget,
) -> Result<(), SlideTableMergesError> {
    budget
        .preflight_reassembly(catalog, edited_bytes, deleted_names.len())
        .map_err(super::map_core_error)
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, SlideTableMergesError> {
    slide_table_core::physical_catalog(package).map_err(super::map_core_error)
}

fn previews_absent(
    package: &Package,
    budget: &mut slide_table_core::Budget,
) -> Result<bool, SlideTableMergesError> {
    slide_table_core::charge_preview_scan(package, budget).map_err(super::map_core_error)?;
    super::super::rendering_invalidation::root_previews_absent(physical_catalog(package)?.package())
        .map_err(|_| SlideTableMergesError::Verification)
}

fn same_selection(left: &MergeSelection, right: &MergeSelection) -> bool {
    slide_table_core::same_target(&left.target, &right.target)
}

fn map_read_error(error: ReadError) -> SlideTableMergesError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableMergesError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideTableMergesLimitKind::References,
                _ => SlideTableMergesLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableMergesError::LimitExceeded {
            kind: match kind {
                PayloadLimitKind::Bytes => SlideTableMergesLimitKind::InputBytes,
                PayloadLimitKind::Fields => SlideTableMergesLimitKind::WireFields,
                PayloadLimitKind::Nesting => SlideTableMergesLimitKind::WireNesting,
                PayloadLimitKind::Work => SlideTableMergesLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableMergesError::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        _ => SlideTableMergesError::InvalidSource,
    }
}

fn map_archive_error(error: ArchiveError) -> SlideTableMergesError {
    match error {
        ArchiveError::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableMergesError::LimitExceeded {
            kind: match kind {
                ArchiveLimitKind::InputBytes => SlideTableMergesLimitKind::InputBytes,
                ArchiveLimitKind::OutputBytes => SlideTableMergesLimitKind::OutputBytes,
                ArchiveLimitKind::Entries => SlideTableMergesLimitKind::Entries,
                ArchiveLimitKind::EntryBytes | ArchiveLimitKind::CompressedEntryBytes => {
                    SlideTableMergesLimitKind::EntryBytes
                },
                ArchiveLimitKind::TotalBytes => SlideTableMergesLimitKind::TotalBytes,
                _ => SlideTableMergesLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        ArchiveError::Allocation { amount, .. } => SlideTableMergesError::Allocation { amount },
        ArchiveError::Iwa(error) => map_native_error(error),
        _ => SlideTableMergesError::InvalidSource,
    }
}

fn map_native_error(error: litchi_iwa_core::Error) -> SlideTableMergesError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableMergesError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => SlideTableMergesLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableMergesLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => SlideTableMergesLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => SlideTableMergesLimitKind::EntryBytes,
                _ => SlideTableMergesLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableMergesError::Allocation { amount: requested }
        },
        _ => SlideTableMergesError::InvalidSource,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Package, SlideSelector, TableSelector};

    const NATIVE_SOURCE: &[u8] =
        include_bytes!("../../../../../test-data/iwork/keynote/slide-table-merges-native.key");

    fn source_bytes(package: &Package) -> Vec<u8> {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes).expect("write Keynote source");
        bytes
    }

    #[test]
    fn merge_and_unmerge_are_source_checked_and_reversible() {
        let package = Package::from_bytes(NATIVE_SOURCE).expect("native Keynote source");
        let baseline = source_bytes(&package);
        let added = Region::new(0, 0, 2, 2).expect("checked merge region");
        let mut edit = package
            .edit_slide_table_merges(SlideSelector::index(0), TableSelector::index(0))
            .expect("select table");
        edit.merge(added).expect("stage merge");
        let commit = edit.commit().expect("commit merge");
        assert!(commit.diagnostics().changed());
        assert_eq!(
            commit
                .package()
                .slide_table_merges(0, 0)
                .expect("read merged regions"),
            [Region::new(3, 1, 2, 2).unwrap(), added]
        );

        let inverse_patch = commit.patch().inverse();
        let changed = commit.into_package();
        let restored = changed
            .apply_slide_table_merges(&inverse_patch)
            .expect("apply inverse")
            .into_package();
        assert_eq!(source_bytes(&restored), baseline);
    }

    #[test]
    fn missing_unmerge_is_byte_exact_noop() {
        let package = Package::from_bytes(NATIVE_SOURCE).expect("native Keynote source");
        use std::sync::atomic::Ordering;
        assert_eq!(
            package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            0
        );
        let baseline = source_bytes(&package);
        let missing = Region::new(0, 0, 2, 2).expect("checked merge region");
        let outside = Region::new(100, 100, 1, 2).expect("checked outside region");
        let mut edit = package.edit_slide_table_merges(0, 0).expect("select table");
        assert!(!edit.unmerge(missing).expect("stage missing unmerge"));
        assert!(
            !edit
                .unmerge(outside)
                .expect("stage outside missing unmerge")
        );
        let commit = edit.commit().expect("commit no-op");
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(source_bytes(commit.package()), baseline);
        let applied = package
            .apply_slide_table_merges(commit.patch())
            .expect("apply no-op");
        assert_eq!(source_bytes(applied.package()), baseline);
        assert_eq!(
            package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            0
        );
    }

    #[test]
    fn duplicate_or_out_of_bounds_merge_is_rejected_without_public_ids() {
        let package = Package::from_bytes(NATIVE_SOURCE).expect("native Keynote source");
        let existing = Region::new(3, 1, 2, 2).expect("checked existing region");
        let mut duplicate_edit = package.edit_slide_table_merges(0, 0).expect("select table");
        assert!(matches!(
            duplicate_edit.merge(existing),
            Err(SlideTableMergesError::InvalidRegion)
        ));

        let outside = Region::new(4, 2, 1, 2).expect("checked region");
        let mut outside_edit = package.edit_slide_table_merges(0, 0).expect("select table");
        assert!(matches!(
            outside_edit.merge(outside),
            Err(SlideTableMergesError::InvalidRegion)
        ));
    }

    #[test]
    fn second_archive_budget_failure_is_atomic() {
        let package = Package::from_bytes(NATIVE_SOURCE).expect("native Keynote source");
        let baseline = source_bytes(&package);
        let mut admission = slide_table_core::Budget::new(&package).expect("merge budget");
        let selection = select_merges(
            &package,
            SlideSelector::index(0),
            TableSelector::index(0),
            &mut admission,
        )
        .expect("select table");

        // Leave enough allocation allowance for both archive preflights, but
        // none for the parsed archive inventory. The second archive must fail
        // before any candidate bytes can be published.
        let mut budget = selection.budget;
        let remaining = budget
            .remaining_allocations()
            .expect("remaining allocation allowance");
        assert!(remaining > 2);
        budget
            .allocations(remaining - 2)
            .expect("reserve the two archive preflight allocations");

        let desired = [Region::new(0, 0, 2, 2).expect("checked merge region")];
        let error = rewrite_merges(&package, &selection, &desired, &mut budget)
            .expect_err("archive inventory must exceed the constrained budget");
        assert!(matches!(
            error,
            SlideTableMergesError::LimitExceeded {
                kind: SlideTableMergesLimitKind::Allocations,
                ..
            }
        ));
        assert_eq!(source_bytes(&package), baseline);
    }
}
