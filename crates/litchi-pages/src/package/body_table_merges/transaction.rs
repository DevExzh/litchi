//! Exact-source selector-first transactions for Pages body-table merges.
//!
//! The transaction owns only checked [`Region`] values and reuses the parent
//! module's rooted ownership proof, read boundary, and finite operation budget.
//! Changed publication rewrites one selected IWA component and fully reopens
//! the candidate before returning it to the caller.

use std::{fmt, sync::Arc};

use litchi_core::id::generate_guid_bytes;
use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_core::RawMessage;
use litchi_numbers_wire::table_merges::{self, ReadLimits};

use super::{
    BodyTableMergesError, BodyTableMergesLimitKind, Package, map_common_error, map_lock_error,
    map_merge_error, map_package_error, model_message, page_layout, read_regions, table_lock,
};
use crate::selector::BodyTableSelector;
use crate::table::merge::Region;

/// A mutable selector-first edit for one Pages body table's merged-cell
/// geometry.
///
/// The edit owns only archive-free [`Region`] values. Native object
/// identifiers, merge-owner identities, and formula records remain private
/// proof data and are resolved again before every changed commit.
pub struct BodyTableMergesEdit<'a> {
    source: &'a Package,
    target: table_lock::BodyTableTarget,
    before: Vec<Region>,
    regions: Vec<Region>,
    budget: table_lock::WireBudget,
}

impl fmt::Debug for BodyTableMergesEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableMergesEdit")
            .field("before", &self.before)
            .field("regions", &self.regions)
            .finish_non_exhaustive()
    }
}

impl BodyTableMergesEdit<'_> {
    /// Return the regions staged for publication.
    #[must_use]
    pub fn regions(&self) -> &[Region] {
        &self.regions
    }

    /// Stage one new merged-cell rectangle.
    ///
    /// A rectangle that overlaps an existing staged merge is rejected,
    /// matching the native table invariant. Use [`Self::unmerge`]
    /// first when replacing an existing rectangle.
    ///
    /// This changes merge geometry while preserving the underlying cell values.
    pub fn merge(&mut self, region: Region) -> Result<&mut Self, BodyTableMergesError> {
        validate_region(&self.target, region)?;
        self.budget
            .charge_payload_items(1)
            .and_then(|_| {
                self.budget
                    .charge_payload_work(self.regions.len().saturating_add(1))
            })
            .map_err(map_lock_error)?;
        if self
            .regions
            .iter()
            .any(|existing| existing.overlaps(region))
        {
            return Err(BodyTableMergesError::OverlappingRegion);
        }
        self.regions
            .try_reserve(1)
            .map_err(|_| BodyTableMergesError::Allocation { amount: 1 })?;
        self.regions.push(region);
        Ok(self)
    }

    /// Remove one staged merged-cell rectangle.
    ///
    /// Removing a rectangle that is absent is a successful semantic no-op,
    /// which makes this operation safe to use for idempotent cleanup.
    pub fn unmerge(&mut self, region: Region) -> Result<bool, BodyTableMergesError> {
        let work = self
            .regions
            .len()
            .checked_mul(2)
            .ok_or(BodyTableMergesError::InvalidSource)?;
        self.budget
            .charge_payload_work(work)
            .map_err(map_lock_error)?;
        let Some(index) = self.regions.iter().position(|existing| *existing == region) else {
            return Ok(false);
        };
        self.regions.remove(index);
        Ok(true)
    }

    /// Validate and atomically publish the staged merge topology.
    pub fn commit(self) -> Result<BodyTableMergesCommit, BodyTableMergesError> {
        commit_edit(self)
    }
}

/// Reversible exact-source patch for one body-table merge topology.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyTableMergesPatch {
    artifacts: OwnedExactArtifacts,
    proof: table_lock::BodyTableTarget,
    before: Vec<Region>,
    after: Vec<Region>,
}

impl fmt::Debug for BodyTableMergesPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableMergesPatch")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyTableMergesPatch {
    /// Return the source merge topology.
    #[must_use]
    pub fn before(&self) -> &[Region] {
        &self.before
    }

    /// Return the target merge topology.
    #[must_use]
    pub fn after(&self) -> &[Region] {
        &self.after
    }

    /// Return the source artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether semantic topology and exact package bytes are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            proof: self.proof.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Compact evidence from one committed body-table merge transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyTableMergesDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl BodyTableMergesDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from source bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one body-table merge transaction.
#[must_use = "a Pages body-table merge commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyTableMergesCommit {
    package: Package,
    patch: BodyTableMergesPatch,
    diagnostics: BodyTableMergesDiagnostics,
}

impl BodyTableMergesCommit {
    /// Borrow the fully reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyTableMergesPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableMergesDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Start a selector-first immutable body-table merge edit.
    pub fn edit_body_table_merges<'table, S>(
        &self,
        selector: S,
    ) -> Result<BodyTableMergesEdit<'_>, BodyTableMergesError>
    where
        S: Into<BodyTableSelector<'table>>,
    {
        let mut budget = transaction_budget(self)?;
        let target = self
            .resolve_body_table_with_budget(selector.into(), &mut budget)
            .map_err(map_lock_error)?;
        let before = read_regions(self, &target, &mut budget)?;
        Ok(BodyTableMergesEdit {
            source: self,
            target,
            regions: before.clone(),
            before,
            budget,
        })
    }

    /// Apply a reversible exact-source body-table merge patch.
    pub fn apply_body_table_merges(
        &self,
        patch: &BodyTableMergesPatch,
    ) -> Result<BodyTableMergesCommit, BodyTableMergesError> {
        let mut budget = transaction_budget(self)?;
        let source = self.state.source.shared_source();
        if !patch
            .artifacts
            .authorizes_owner(&SharedBytes::from_shared_slice(source.clone()))
        {
            return Err(BodyTableMergesError::PatchConflict);
        }
        let target = self
            .resolve_body_table_with_budget(
                BodyTableSelector::index(patch.proof.table_position),
                &mut budget,
            )
            .map_err(map_lock_error)?;
        if target.model_identifier != patch.proof.model_identifier {
            return Err(BodyTableMergesError::PatchConflict);
        }
        let current = read_regions(self, &target, &mut budget)?;
        if current != patch.before {
            return Err(BodyTableMergesError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(BodyTableMergesCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyTableMergesDiagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact() {
            return Err(BodyTableMergesError::PatchConflict);
        }
        let target_owner = patch.artifacts.target_owner();
        let candidate = reopen_candidate(self, target_owner.as_slice(), &mut budget)?;
        let verified_target = candidate
            .resolve_body_table_with_budget(
                BodyTableSelector::index(patch.proof.table_position),
                &mut budget,
            )
            .map_err(map_lock_error)?;
        if verified_target.model_identifier != patch.proof.model_identifier {
            return Err(BodyTableMergesError::Verification);
        }
        if read_regions(&candidate, &verified_target, &mut budget)? != patch.after {
            return Err(BodyTableMergesError::Verification);
        }
        verify_locality(self, &candidate, &patch.proof, &mut budget)?;
        Ok(BodyTableMergesCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableMergesDiagnostics::published(),
        })
    }
}

fn transaction_budget(package: &Package) -> Result<table_lock::WireBudget, BodyTableMergesError> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    budget
        .charge_source_catalog(&package.state.source)
        .map_err(map_lock_error)?;
    Ok(budget)
}

fn validate_region(
    target: &table_lock::BodyTableTarget,
    region: Region,
) -> Result<(), BodyTableMergesError> {
    if region.end_row() >= target.table_rows || region.end_column() >= target.table_columns {
        return Err(BodyTableMergesError::InvalidRegion);
    }
    Ok(())
}

fn commit_edit(
    edit: BodyTableMergesEdit<'_>,
) -> Result<BodyTableMergesCommit, BodyTableMergesError> {
    let source = edit.source;
    let mut budget = edit.budget;
    let after = normalize_regions(&edit.before, &edit.regions, &mut budget)?;
    let source_owner = SharedBytes::from_shared_slice(source.state.source.shared_source());
    if edit.before == after {
        return Ok(BodyTableMergesCommit {
            package: source.snapshot(),
            patch: BodyTableMergesPatch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                proof: edit.target,
                before: edit.before,
                after,
            },
            diagnostics: BodyTableMergesDiagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(BodyTableMergesError::UnsupportedSource);
    }
    table_lock::validate_body_table_target(source, &edit.target, &mut budget)
        .map_err(map_lock_error)?;
    let package = rewrite_model(source, &edit.target, &edit.before, &after, &mut budget)?;
    let candidate_target = package
        .resolve_body_table_with_budget(
            BodyTableSelector::index(edit.target.table_position),
            &mut budget,
        )
        .map_err(map_lock_error)?;
    if candidate_target.model_identifier != edit.target.model_identifier
        || read_regions(&package, &candidate_target, &mut budget)? != after
    {
        return Err(BodyTableMergesError::Verification);
    }
    verify_locality(source, &package, &edit.target, &mut budget)?;
    let target_owner = SharedBytes::from_shared_slice(package.state.source.shared_source());
    let artifacts = OwnedExactArtifacts::new(source_owner, target_owner);
    Ok(BodyTableMergesCommit {
        package,
        patch: BodyTableMergesPatch {
            artifacts,
            proof: edit.target,
            before: edit.before,
            after,
        },
        diagnostics: BodyTableMergesDiagnostics::published(),
    })
}

fn normalize_regions(
    before: &[Region],
    staged: &[Region],
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<Region>, BodyTableMergesError> {
    let comparisons = before
        .len()
        .checked_mul(staged.len())
        .and_then(|count| count.checked_mul(2))
        .and_then(|count| count.checked_add(staged.len()))
        .ok_or(BodyTableMergesError::InvalidSource)?;
    budget
        .charge_payload_items(staged.len())
        .and_then(|_| budget.charge_payload_work(comparisons))
        .map_err(map_lock_error)?;
    let mut normalized = Vec::new();
    normalized
        .try_reserve_exact(staged.len())
        .map_err(|_| BodyTableMergesError::Allocation {
            amount: staged.len(),
        })?;
    normalized.extend(
        before
            .iter()
            .copied()
            .filter(|region| staged.contains(region)),
    );
    normalized.extend(
        staged
            .iter()
            .copied()
            .filter(|region| !before.contains(region)),
    );
    Ok(normalized)
}

#[derive(Clone, Copy)]
struct MergeRewriteBounds {
    compressed_bound: usize,
    package_output_bound: usize,
    rewritten_message_bound: usize,
    archive_bound: usize,
}

fn rewrite_model(
    source: &Package,
    target: &table_lock::BodyTableTarget,
    before: &[Region],
    after: &[Region],
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableMergesError> {
    let catalog = &source.state.source;
    let component = catalog
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    let component_name = component.name();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyTableMergesError::UnsupportedSource);
    }
    let original = model_message(source, target)?.data.as_slice();
    let bounds = preflight_merge_rewrite(source, target, before, after, original.len(), budget)?;
    let (mut archive, archive_limits) =
        page_layout::editable_archive(source, component_name).map_err(map_page_layout_error)?;
    let object = archive
        .objects
        .get_mut(target.model_object_index)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableMergesError::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, target.model_message_index)
        .map_err(map_page_layout_error)?;
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    let current = message.data.clone();
    if current != original {
        return Err(BodyTableMergesError::InvalidSource);
    }
    let limits = writer_limits(budget, bounds, before.len().max(after.len()))?;
    let write = table_merges::rewrite_table_merges(&current, after, generate_guid_bytes(), limits)
        .map_err(map_merge_error)?;
    if !write.changed || write.data.as_slice() == current.as_slice() {
        return Err(BodyTableMergesError::Verification);
    }
    budget
        .charge_payload_work(write.report.input_bytes())
        .and_then(|_| budget.charge_codec_report(write.report.fields(), write.report.work(), 0, 0))
        .and_then(|_| budget.charge_payload_work(write.data.len()))
        .map_err(map_lock_error)?;
    if write.data.len() > bounds.rewritten_message_bound {
        return Err(BodyTableMergesError::Verification);
    }
    object
        .replace_message_preserving_header_with_limits(
            target.model_message_index,
            RawMessage {
                type_: target.model_message_type,
                data: write.data,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed =
        page_layout::compress_archive(archive, archive_limits).map_err(map_page_layout_error)?;
    if compressed.len() > bounds.compressed_bound {
        return Err(BodyTableMergesError::Verification);
    }
    let edits = [EntryEdit::new(component_name, compressed.as_slice())];
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], catalog.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    if requirements.output_bytes() > bounds.package_output_bound
        || requirements.retained_bytes() > bounds.package_output_bound
    {
        return Err(BodyTableMergesError::Verification);
    }
    budget
        .charge_payload_work(requirements.output_bytes())
        .and_then(|_| budget.charge_payload_work(requirements.retained_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.scratch_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.allocations()))
        .map_err(map_lock_error)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    if output.len() != requirements.output_bytes() {
        return Err(BodyTableMergesError::Verification);
    }
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), catalog.limits())
            .map_err(map_archive_error)?;
    table_lock::charge_reopen_work(&candidate_source, budget).map_err(map_lock_error)?;
    Package::from_source_catalog(candidate_source).map_err(map_package_error)
}

fn preflight_merge_rewrite(
    source: &Package,
    target: &table_lock::BodyTableTarget,
    before: &[Region],
    after: &[Region],
    original_message_len: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<MergeRewriteBounds, BodyTableMergesError> {
    // Canonical native merge formulas are small. Reserve a deliberately
    // conservative per-new-region bound before the shared writer can create
    // its output vector; retained formula records are copied unchanged.
    const MAX_NEW_FORMULA_BYTES: usize = 1_024;
    const MAX_VARINT_BYTES: usize = 10;
    let added = after
        .iter()
        .filter(|region| !before.contains(region))
        .count();
    let rewritten_message_bound = original_message_len
        .checked_add(
            added
                .checked_mul(MAX_NEW_FORMULA_BYTES)
                .ok_or(BodyTableMergesError::InvalidSource)?,
        )
        .and_then(|value| value.checked_add(MAX_VARINT_BYTES.saturating_mul(4)))
        .ok_or(BodyTableMergesError::InvalidSource)?;
    let component = source
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    let archive_source_length = parsed_archive_source_length(component.archive())?;
    let archive_bound = archive_source_length
        .checked_sub(original_message_len)
        .and_then(|value| value.checked_add(rewritten_message_bound))
        .and_then(|value| value.checked_add(MAX_VARINT_BYTES.saturating_mul(4)))
        .ok_or(BodyTableMergesError::InvalidSource)?;
    let compressed_bound = table_lock::snappy_compressed_bound(archive_bound)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    let entry = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or(BodyTableMergesError::InvalidSource)?;
    let replacement_compressed_bound = match entry.metadata().central().compression_method() {
        0 => compressed_bound,
        8 => table_lock::deflate_compressed_bound(compressed_bound)
            .ok_or(BodyTableMergesError::InvalidSource)?,
        _ => return Err(BodyTableMergesError::UnsupportedSource),
    };
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| {
            BodyTableMergesError::LimitExceeded {
                kind: BodyTableMergesLimitKind::EntryBytes,
                observed: u64::MAX,
                maximum: source.state.source.limits().max_entry_bytes(),
            }
        })?;
    let package_output_bound = source
        .state
        .source
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_add(replacement_compressed_bound))
        .ok_or(BodyTableMergesError::InvalidSource)?;
    budget
        .charge_output_bytes(rewritten_message_bound)
        .and_then(|_| budget.charge_output_bytes(archive_bound))
        .and_then(|_| budget.charge_output_bytes(compressed_bound))
        .and_then(|_| budget.charge_output_bytes(replacement_compressed_bound))
        .and_then(|_| budget.charge_output_bytes(package_output_bound))
        .and_then(|_| budget.charge_payload_bytes(archive_bound))
        .and_then(|_| budget.charge_total_payload_bytes(archive_bound))
        .and_then(|_| budget.charge_payload_work(archive_bound))
        .and_then(|_| budget.charge_payload_work(compressed_bound))
        .and_then(|_| budget.charge_payload_work(replacement_compressed_bound))
        .and_then(|_| budget.charge_payload_work(package_output_bound))
        .and_then(|_| {
            budget.precharge_candidate_reopen(
                &source.state.source,
                package_output_bound,
                target.model_component_index,
                compressed_bound,
                archive_bound,
                target.model_object_index,
                target.model_message_index,
                rewritten_message_bound,
            )
        })
        .map_err(map_lock_error)?;
    Ok(MergeRewriteBounds {
        compressed_bound,
        package_output_bound,
        rewritten_message_bound,
        archive_bound,
    })
}

fn writer_limits(
    budget: &table_lock::WireBudget,
    bounds: MergeRewriteBounds,
    max_regions: usize,
) -> Result<ReadLimits, BodyTableMergesError> {
    let fields = budget.remaining_wire_fields();
    if fields == 0 {
        return Err(BodyTableMergesError::LimitExceeded {
            kind: BodyTableMergesLimitKind::WireFields,
            observed: 1,
            maximum: 0,
        });
    }
    let work = budget.remaining_wire_work();
    if work == 0 {
        return Err(BodyTableMergesError::LimitExceeded {
            kind: BodyTableMergesLimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let base = budget.wire_limits();
    if bounds.rewritten_message_bound > base.max_output_bytes() {
        return Err(BodyTableMergesError::LimitExceeded {
            kind: BodyTableMergesLimitKind::WireOutputBytes,
            observed: bounds.rewritten_message_bound as u64,
            maximum: base.max_output_bytes() as u64,
        });
    }
    let wire = base
        .with_input_bytes(bounds.archive_bound.min(base.max_input_bytes()).max(1))
        .and_then(|limits| limits.with_output_bytes(bounds.rewritten_message_bound.max(1)))
        .and_then(|limits| limits.with_fields(fields))
        .and_then(|limits| limits.with_rewrite_work(work))
        .map_err(map_common_error)?;
    Ok(ReadLimits {
        wire,
        max_regions: max_regions.max(1),
        max_overlap_checks: work,
    })
}

fn parsed_archive_source_length(
    archive: &litchi_iwa_core::Archive,
) -> Result<usize, BodyTableMergesError> {
    let Some(last) = archive.objects.last() else {
        return Ok(0);
    };
    let length = last
        .data_offset
        .checked_add(last.data_length)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    usize::try_from(length).map_err(|_| BodyTableMergesError::InvalidSource)
}

fn reopen_candidate(
    source: &Package,
    bytes: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableMergesError> {
    budget
        .charge_output_bytes(bytes.len())
        .and_then(|_| budget.charge_payload_work(bytes.len()))
        .map_err(map_lock_error)?;
    let catalog = SourceCatalog::from_shared_bytes_with_limits(
        Arc::<[u8]>::from(bytes),
        source.state.source.limits(),
    )
    .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .and_then(|_| table_lock::charge_reopen_work(&catalog, budget))
        .map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableMergesError> {
    let left = &source.state.source;
    let right = &candidate.state.source;
    if left.components().len() != right.components().len()
        || left.package().len() != right.package().len()
    {
        return Err(BodyTableMergesError::Verification);
    }
    let selected_name = left
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableMergesError::Verification)?
        .name();
    for (before, after) in left.package().iter().zip(right.package().iter()) {
        budget
            .charge_payload_work(
                before
                    .data()
                    .len()
                    .saturating_add(after.data().len())
                    .saturating_add(before.raw_name().len())
                    .saturating_add(after.raw_name().len()),
            )
            .map_err(map_lock_error)?;
        let selected = before.name() == selected_name;
        if before.name() != after.name()
            || before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
            || (!selected
                && (before.raw_record().local_record() != after.raw_record().local_record()
                    || !central_record_preserved_except_offset(
                        before.raw_record().central_directory_record(),
                        after.raw_record().central_directory_record(),
                    )
                    || before.data() != after.data()))
        {
            return Err(BodyTableMergesError::Verification);
        }
    }
    for (component_index, (before, after)) in left
        .components()
        .iter()
        .zip(right.components().iter())
        .enumerate()
    {
        if before.name() != after.name()
            || before.archive().objects.len() != after.archive().objects.len()
        {
            return Err(BodyTableMergesError::Verification);
        }
        for (object_index, (before_object, after_object)) in before
            .archive()
            .objects
            .iter()
            .zip(after.archive().objects.iter())
            .enumerate()
        {
            if component_index != target.model_component_index
                || object_index != target.model_object_index
            {
                budget
                    .charge_payload_work(
                        before_object
                            .data_length
                            .try_into()
                            .unwrap_or(usize::MAX)
                            .saturating_add(
                                after_object.data_length.try_into().unwrap_or(usize::MAX),
                            ),
                    )
                    .map_err(map_lock_error)?;
                if !before_object.same_content_ignoring_offsets(after_object) {
                    return Err(BodyTableMergesError::Verification);
                }
                continue;
            }
            if before_object.archive_info.identifier != after_object.archive_info.identifier
                || before_object.archive_info.should_merge != after_object.archive_info.should_merge
                || before_object.messages.len() != after_object.messages.len()
                || before_object.archive_info.message_infos.len()
                    != after_object.archive_info.message_infos.len()
            {
                return Err(BodyTableMergesError::Verification);
            }
            for (message_index, (before_message, after_message)) in before_object
                .messages
                .iter()
                .zip(after_object.messages.iter())
                .enumerate()
            {
                let before_info = before_object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(BodyTableMergesError::Verification)?;
                let after_info = after_object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(BodyTableMergesError::Verification)?;
                budget
                    .charge_payload_work(
                        before_message
                            .data
                            .len()
                            .saturating_add(after_message.data.len()),
                    )
                    .map_err(map_lock_error)?;
                if message_index == target.model_message_index {
                    if before_message.type_ != after_message.type_
                        || !message_info_preserved_except_length(before_info, after_info)
                        || before_message.data == after_message.data
                    {
                        return Err(BodyTableMergesError::Verification);
                    }
                } else if before_message != after_message || before_info != after_info {
                    return Err(BodyTableMergesError::Verification);
                }
            }
        }
    }
    Ok(())
}

fn message_info_preserved_except_length(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn central_record_preserved_except_offset(before: &[u8], after: &[u8]) -> bool {
    const OFFSET: std::ops::Range<usize> = 42..46;
    before.len() == after.len()
        && before.len() >= OFFSET.end
        && before[..OFFSET.start] == after[..OFFSET.start]
        && before[OFFSET.end..] == after[OFFSET.end..]
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyTableMergesError {
    map_lock_error(table_lock::map_archive_error(error))
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyTableMergesError {
    match error {
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyTableMergesError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableMergesError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => {
                    BodyTableMergesLimitKind::TotalPayloadBytes
                },
                litchi_iwa_core::LimitKind::Objects => BodyTableMergesLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    BodyTableMergesLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes => {
                    BodyTableMergesLimitKind::PayloadBytes
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems => {
                    BodyTableMergesLimitKind::PayloadItems
                },
                litchi_iwa_core::LimitKind::HeaderNesting => BodyTableMergesLimitKind::WireNesting,
                _ => BodyTableMergesLimitKind::PayloadBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => BodyTableMergesError::InvalidSource,
    }
}

fn map_page_layout_error(error: page_layout::PageLayoutError) -> BodyTableMergesError {
    match error {
        page_layout::PageLayoutError::UnsupportedSource => BodyTableMergesError::UnsupportedSource,
        page_layout::PageLayoutError::InvalidSource
        | page_layout::PageLayoutError::InvalidLayout(_) => BodyTableMergesError::InvalidSource,
        page_layout::PageLayoutError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableMergesError::LimitExceeded {
            kind: match kind {
                page_layout::PageLayoutLimitKind::InputBytes => {
                    BodyTableMergesLimitKind::InputBytes
                },
                page_layout::PageLayoutLimitKind::OutputBytes => {
                    BodyTableMergesLimitKind::OutputBytes
                },
                page_layout::PageLayoutLimitKind::Entries => BodyTableMergesLimitKind::Entries,
                page_layout::PageLayoutLimitKind::EntryBytes => {
                    BodyTableMergesLimitKind::EntryBytes
                },
                page_layout::PageLayoutLimitKind::TotalEntryBytes => {
                    BodyTableMergesLimitKind::TotalEntryBytes
                },
                page_layout::PageLayoutLimitKind::PackageBytes => {
                    BodyTableMergesLimitKind::PackageBytes
                },
                page_layout::PageLayoutLimitKind::PayloadBytes => {
                    BodyTableMergesLimitKind::PayloadBytes
                },
                page_layout::PageLayoutLimitKind::TotalPayloadBytes => {
                    BodyTableMergesLimitKind::TotalPayloadBytes
                },
                page_layout::PageLayoutLimitKind::PayloadObjects => {
                    BodyTableMergesLimitKind::PayloadObjects
                },
                page_layout::PageLayoutLimitKind::PayloadMessages => {
                    BodyTableMergesLimitKind::PayloadMessages
                },
                page_layout::PageLayoutLimitKind::PayloadItems => {
                    BodyTableMergesLimitKind::PayloadItems
                },
                page_layout::PageLayoutLimitKind::WireBytes => BodyTableMergesLimitKind::WireBytes,
                page_layout::PageLayoutLimitKind::WireFields => {
                    BodyTableMergesLimitKind::WireFields
                },
                page_layout::PageLayoutLimitKind::WireNesting => {
                    BodyTableMergesLimitKind::WireNesting
                },
                page_layout::PageLayoutLimitKind::WireWork => BodyTableMergesLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        page_layout::PageLayoutError::Allocation { amount } => {
            BodyTableMergesError::Allocation { amount }
        },
        page_layout::PageLayoutError::Verification => BodyTableMergesError::Verification,
        page_layout::PageLayoutError::PatchConflict => BodyTableMergesError::PatchConflict,
    }
}
