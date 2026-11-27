//! Exact-source selector-first transactions for Pages body-table merges.
//!
//! The transaction owns only checked [`Region`] values and reuses the parent
//! module's rooted ownership proof, read boundary, and finite operation budget.
//! Changed publication rewrites one selected IWA component and fully reopens
//! the candidate before returning it to the caller.

use std::{fmt, sync::Arc};

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_core::RawMessage;
use litchi_numbers_wire::table_merges::{self, ReadLimits};

use super::{
    Error, LimitKind, Package, map_common_error, map_lock_error, map_merge_error,
    map_package_error, model_message, page_layout, read_regions, table_lock,
};
use crate::selector::BodyTableSelector;
use crate::table::merge::Region;

/// A mutable selector-first edit for one Pages body table's merged-cell
/// geometry.
///
/// The edit owns only archive-free [`Region`] values. Native object
/// identifiers, merge-owner identities, and formula records remain private
/// proof data and are resolved again before every changed commit.
pub struct Edit<'a> {
    source: &'a Package,
    selector: table_lock::OwnedBodyTableSelector,
    target: Option<table_lock::BodyTableTarget>,
    before: Option<Vec<Region>>,
    regions: Vec<Region>,
    budget: Option<table_lock::WireBudget>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("before", &self.before)
            .field("regions", &self.regions)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn resolve(&mut self) -> Result<(), Error> {
        if self.target.is_none() {
            let mut budget = match self.budget.take() {
                Some(budget) => budget,
                None => transaction_budget(self.source)?,
            };
            let target = match self
                .selector
                .resolve_with_budget(self.source, &mut budget)
                .map_err(map_lock_error)
            {
                Ok(target) => target,
                Err(error) => {
                    self.budget = Some(budget);
                    return Err(error);
                },
            };
            let before = match read_regions(self.source, &target, &mut budget) {
                Ok(before) => before,
                Err(error) => {
                    self.budget = Some(budget);
                    return Err(error);
                },
            };
            if self.before.is_none() {
                self.before = Some(before.clone());
                self.regions = before;
            }
            self.target = Some(target);
            self.budget = Some(budget);
        }
        Ok(())
    }

    /// Return the regions staged for publication.
    pub fn regions(&mut self) -> Result<&[Region], Error> {
        self.resolve()?;
        Ok(&self.regions)
    }

    /// Stage one new merged-cell rectangle.
    ///
    /// A rectangle that overlaps an existing staged merge is rejected,
    /// matching the native table invariant. Use [`Self::unmerge`]
    /// first when replacing an existing rectangle.
    ///
    /// This changes merge geometry while preserving the underlying cell values.
    pub fn merge(&mut self, region: Region) -> Result<&mut Self, Error> {
        self.resolve()?;
        validate_region(self.target.as_ref().ok_or(Error::InvalidSource)?, region)?;
        let budget = self.budget.as_mut().ok_or(Error::InvalidSource)?;
        budget.charge_payload_items(1).map_err(map_lock_error)?;
        budget
            .charge_payload_work(self.regions.len().saturating_add(1))
            .map_err(map_lock_error)?;
        if self
            .regions
            .iter()
            .any(|existing| existing.overlaps(region))
        {
            return Err(Error::Refused(
                crate::table::merge::transaction::Refusal::OverlappingRegion,
            ));
        }
        self.regions
            .try_reserve(1)
            .map_err(|_| Error::Allocation { amount: 1 })?;
        self.regions.push(region);
        Ok(self)
    }

    /// Remove one staged merged-cell rectangle.
    ///
    /// Removing a rectangle that is absent is a successful semantic no-op,
    /// which makes this operation safe to use for idempotent cleanup.
    pub fn unmerge(&mut self, region: Region) -> Result<bool, Error> {
        self.resolve()?;
        let work = self
            .regions
            .len()
            .checked_mul(2)
            .ok_or(Error::InvalidSource)?;
        self.budget
            .as_mut()
            .ok_or(Error::InvalidSource)?
            .charge_payload_work(work)
            .map_err(map_lock_error)?;
        let Some(index) = self.regions.iter().position(|existing| *existing == region) else {
            return Ok(false);
        };
        self.regions.remove(index);
        Ok(true)
    }

    /// Validate and atomically publish the staged merge topology.
    pub fn commit(mut self) -> Result<Commit, Error> {
        self.resolve()?;
        commit_edit(self)
    }
}

/// Reversible exact-source patch for one body-table merge topology.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    proof: table_lock::BodyTableTarget,
    before: Vec<Region>,
    after: Vec<Region>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
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
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
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
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
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

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Start a selector-first immutable body-table merge edit.
    pub fn edit_body_table_merges<'table, S>(&self, selector: S) -> Result<Edit<'_>, Error>
    where
        S: Into<BodyTableSelector<'table>>,
    {
        Ok(Edit {
            source: self,
            selector: table_lock::OwnedBodyTableSelector::try_from(self, selector.into())
                .map_err(map_lock_error)?,
            target: None,
            regions: Vec::new(),
            before: None,
            budget: None,
        })
    }

    /// Apply a reversible exact-source body-table merge patch.
    pub fn apply_body_table_merges(&self, patch: &Patch) -> Result<Commit, Error> {
        let mut budget = transaction_budget(self)?;
        let source = self.state.source.shared_source();
        if !patch
            .artifacts
            .authorizes_owner(&SharedBytes::from_shared_slice(source.clone()))
        {
            return Err(Error::PatchConflict);
        }
        let target = self
            .resolve_body_table_with_budget(
                BodyTableSelector::index(patch.proof.table_position),
                &mut budget,
            )
            .map_err(map_lock_error)?;
        if target.model_identifier != patch.proof.model_identifier {
            return Err(Error::PatchConflict);
        }
        let current = read_regions(self, &target, &mut budget)?;
        if current != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact() {
            return Err(Error::PatchConflict);
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
            return Err(Error::Verification);
        }
        if read_regions(&candidate, &verified_target, &mut budget)? != patch.after {
            return Err(Error::Verification);
        }
        verify_locality(self, &candidate, &patch.proof, &mut budget)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }
}

fn transaction_budget(package: &Package) -> Result<table_lock::WireBudget, Error> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    budget
        .charge_source_catalog(&package.state.source)
        .map_err(map_lock_error)?;
    Ok(budget)
}

fn validate_region(target: &table_lock::BodyTableTarget, region: Region) -> Result<(), Error> {
    if region.end_row() >= target.table_rows || region.end_column() >= target.table_columns {
        return Err(Error::Refused(
            crate::table::merge::transaction::Refusal::InvalidRegion,
        ));
    }
    Ok(())
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let target = edit.target.ok_or(Error::InvalidSource)?;
    let before = edit.before.ok_or(Error::InvalidSource)?;
    let mut budget = edit.budget.ok_or(Error::InvalidSource)?;
    let after = normalize_regions(&before, &edit.regions, &mut budget)?;
    let source_owner = SharedBytes::from_shared_slice(source.state.source.shared_source());
    if before == after {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                proof: target,
                before,
                after,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(Error::Refused(
            crate::table::merge::transaction::Refusal::UnsupportedSource,
        ));
    }
    table_lock::validate_body_table_target(source, &target, &mut budget).map_err(map_lock_error)?;
    let package = rewrite_model(source, &target, &before, &after, &mut budget)?;
    let candidate_target = package
        .resolve_body_table_with_budget(
            BodyTableSelector::index(target.table_position),
            &mut budget,
        )
        .map_err(map_lock_error)?;
    if candidate_target.model_identifier != target.model_identifier
        || read_regions(&package, &candidate_target, &mut budget)? != after
    {
        return Err(Error::Verification);
    }
    verify_locality(source, &package, &target, &mut budget)?;
    let target_owner = SharedBytes::from_shared_slice(package.state.source.shared_source());
    let artifacts = OwnedExactArtifacts::new(source_owner, target_owner);
    Ok(Commit {
        package,
        patch: Patch {
            artifacts,
            proof: target,
            before,
            after,
        },
        diagnostics: Diagnostics::published(),
    })
}

fn normalize_regions(
    before: &[Region],
    staged: &[Region],
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<Region>, Error> {
    let comparisons = before
        .len()
        .checked_mul(staged.len())
        .and_then(|count| count.checked_mul(2))
        .and_then(|count| count.checked_add(staged.len()))
        .ok_or(Error::InvalidSource)?;
    budget
        .charge_payload_items(staged.len())
        .and_then(|_| budget.charge_payload_work(comparisons))
        .map_err(map_lock_error)?;
    let mut normalized = Vec::new();
    normalized
        .try_reserve_exact(staged.len())
        .map_err(|_| Error::Allocation {
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
) -> Result<Package, Error> {
    let catalog = &source.state.source;
    let component = catalog
        .components()
        .get_index(target.model_component_index)
        .ok_or(Error::InvalidSource)?;
    let component_name = component.name();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::Refused(
            crate::table::merge::transaction::Refusal::UnsupportedSource,
        ));
    }
    let original = model_message(source, target)?.data.as_slice();
    let bounds = preflight_merge_rewrite(source, target, before, after, original.len(), budget)?;
    let (mut archive, archive_limits) =
        page_layout::editable_archive(source, component_name).map_err(map_page_layout_error)?;
    let object = archive
        .objects
        .get_mut(target.model_object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, target.model_message_index)
        .map_err(map_page_layout_error)?;
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(Error::InvalidSource)?;
    let current = message.data.clone();
    if current != original {
        return Err(Error::InvalidSource);
    }
    let limits = writer_limits(budget, bounds, before.len().max(after.len()))?;
    let write = table_merges::rewrite_table_merges_derived(&current, after, limits)
        .map_err(map_merge_error)?;
    if !write.changed || write.data.as_slice() == current.as_slice() {
        return Err(Error::Verification);
    }
    budget
        .charge_payload_work(write.report.input_bytes())
        .and_then(|_| budget.charge_codec_report(write.report.fields(), write.report.work(), 0, 0))
        .and_then(|_| budget.charge_payload_work(write.data.len()))
        .map_err(map_lock_error)?;
    if write.data.len() > bounds.rewritten_message_bound {
        return Err(Error::Verification);
    }
    object
        .replace_message_preserving_header(
            target.model_message_index,
            RawMessage {
                type_: target.model_message_type,
                data: write.data,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed = page_layout::compress_archive(source, archive, archive_limits)
        .map_err(map_page_layout_error)?;
    if compressed.len() > bounds.compressed_bound {
        return Err(Error::Verification);
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
        return Err(Error::Verification);
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
        return Err(Error::Verification);
    }
    let candidate_source = catalog.reopen(output.into()).map_err(map_archive_error)?;
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
) -> Result<MergeRewriteBounds, Error> {
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
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|value| value.checked_add(MAX_VARINT_BYTES.saturating_mul(4)))
        .ok_or(Error::InvalidSource)?;
    let component = source
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(Error::InvalidSource)?;
    let archive_source_length = parsed_archive_source_length(component.archive())?;
    let archive_bound = archive_source_length
        .checked_sub(original_message_len)
        .and_then(|value| value.checked_add(rewritten_message_bound))
        .and_then(|value| value.checked_add(MAX_VARINT_BYTES.saturating_mul(4)))
        .ok_or(Error::InvalidSource)?;
    let compressed_bound =
        table_lock::snappy_compressed_bound(archive_bound).ok_or(Error::InvalidSource)?;
    let entry = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or(Error::InvalidSource)?;
    let replacement_compressed_bound = match entry.metadata().central().compression_method() {
        0 => compressed_bound,
        8 => table_lock::deflate_compressed_bound(compressed_bound).ok_or(Error::InvalidSource)?,
        _ => {
            return Err(Error::Refused(
                crate::table::merge::transaction::Refusal::UnsupportedSource,
            ));
        },
    };
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| Error::Limit {
            kind: LimitKind::EntryBytes,
            observed: u64::MAX,
            maximum: source.state.source.limits().max_entry_bytes(),
        })?;
    let package_output_bound = source
        .state
        .source
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_add(replacement_compressed_bound))
        .ok_or(Error::InvalidSource)?;
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
) -> Result<ReadLimits, Error> {
    let fields = budget.remaining_wire_fields();
    if fields == 0 {
        return Err(Error::Limit {
            kind: LimitKind::WireFields,
            observed: 1,
            maximum: 0,
        });
    }
    let work = budget.remaining_wire_work();
    if work == 0 {
        return Err(Error::Limit {
            kind: LimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let base = budget.wire_limits();
    if bounds.rewritten_message_bound > base.max_output_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::WireOutputBytes,
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

fn parsed_archive_source_length(archive: &litchi_iwa_core::Archive) -> Result<usize, Error> {
    let Some(last) = archive.objects.last() else {
        return Ok(0);
    };
    let length = last
        .data_offset
        .checked_add(last.data_length)
        .ok_or(Error::InvalidSource)?;
    usize::try_from(length).map_err(|_| Error::InvalidSource)
}

fn reopen_candidate(
    source: &Package,
    bytes: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<Package, Error> {
    budget
        .charge_output_bytes(bytes.len())
        .and_then(|_| budget.charge_payload_work(bytes.len()))
        .map_err(map_lock_error)?;
    let catalog = source
        .state
        .source
        .reopen(Arc::<[u8]>::from(bytes))
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
) -> Result<(), Error> {
    let left = &source.state.source;
    let right = &candidate.state.source;
    if left.components().len() != right.components().len()
        || left.package().len() != right.package().len()
    {
        return Err(Error::Verification);
    }
    let selected_name = left
        .components()
        .get_index(target.model_component_index)
        .ok_or(Error::Verification)?
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
            return Err(Error::Verification);
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
            return Err(Error::Verification);
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
                    return Err(Error::Verification);
                }
                continue;
            }
            if before_object.archive_info.identifier != after_object.archive_info.identifier
                || before_object.archive_info.should_merge != after_object.archive_info.should_merge
                || before_object.messages.len() != after_object.messages.len()
                || before_object.archive_info.message_infos.len()
                    != after_object.archive_info.message_infos.len()
            {
                return Err(Error::Verification);
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
                    .ok_or(Error::Verification)?;
                let after_info = after_object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(Error::Verification)?;
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
                        return Err(Error::Verification);
                    }
                } else if before_message != after_message || before_info != after_info {
                    return Err(Error::Verification);
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

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    map_lock_error(table_lock::map_archive_error(error))
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => LimitKind::TotalPayloadBytes,
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes => LimitKind::PayloadBytes,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                _ => LimitKind::PayloadBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => Error::InvalidSource,
    }
}

fn map_page_layout_error(error: page_layout::Error) -> Error {
    match error {
        page_layout::Error::Refused(
            crate::page_layout::transaction::Refusal::UnsupportedSource,
        ) => Error::Refused(crate::table::merge::transaction::Refusal::UnsupportedSource),
        page_layout::Error::InvalidSource | page_layout::Error::InvalidLayout(_) => {
            Error::InvalidSource
        },
        page_layout::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                page_layout::LimitKind::InputBytes => LimitKind::InputBytes,
                page_layout::LimitKind::OutputBytes => LimitKind::OutputBytes,
                page_layout::LimitKind::Entries => LimitKind::Entries,
                page_layout::LimitKind::EntryBytes => LimitKind::EntryBytes,
                page_layout::LimitKind::TotalEntryBytes => LimitKind::TotalEntryBytes,
                page_layout::LimitKind::PackageBytes => LimitKind::PackageBytes,
                page_layout::LimitKind::PayloadBytes => LimitKind::PayloadBytes,
                page_layout::LimitKind::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
                page_layout::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
                page_layout::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
                page_layout::LimitKind::PayloadItems => LimitKind::PayloadItems,
                page_layout::LimitKind::WireBytes => LimitKind::WireBytes,
                page_layout::LimitKind::WireFields => LimitKind::WireFields,
                page_layout::LimitKind::WireNesting => LimitKind::WireNesting,
                page_layout::LimitKind::WireWork => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
        page_layout::Error::Allocation { amount } => Error::Allocation { amount },
        page_layout::Error::Verification => Error::Verification,
        page_layout::Error::PatchConflict => Error::PatchConflict,
    }
}

#[cfg(test)]
mod budget_tests {
    use super::{Error, LimitKind, Package};
    use crate::selector::BodyTableSelector;
    use crate::table::merge::transaction::Refusal;

    fn fixture() -> Result<Package, Box<dyn std::error::Error>> {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/pages/body-table-visible.pages");
        Ok(Package::open(path)?)
    }

    #[test]
    fn deferred_resolution_preserves_wire_budget_after_failure()
    -> Result<(), Box<dyn std::error::Error>> {
        let package = fixture()?;
        let mut edit = package.edit_body_table_merges(BodyTableSelector::index(usize::MAX))?;

        assert!(matches!(
            edit.regions(),
            Err(Error::Refused(Refusal::TableNotFound))
        ));
        let budget = edit
            .budget
            .as_mut()
            .expect("a failed lazy resolution retains its wire budget");
        budget.exhaust_payload_work_for_test();
        assert!(matches!(
            edit.regions(),
            Err(Error::Limit {
                kind: LimitKind::WireWork,
                ..
            })
        ));
        Ok(())
    }
}
