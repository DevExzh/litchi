//! Exact-source selector-first transactions for Numbers table merges.
//!
//! The transaction owns only archive-free [`Region`] values and a bounded
//! desired state.  Native identifiers, archive positions, and merge-owner
//! UUIDs stay inside the package and wire boundaries.  Changed publication
//! rewrites one selected IWA component and reopens the complete candidate
//! before returning it to the caller.

use std::{fmt, sync::Arc};

use litchi_core::id::generate_guid_bytes;
use litchi_iwa_archive::package::{EntryEdit, SharedBytes};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{Archive, Limits, RawMessage, SnappyStream};
use litchi_numbers_wire::table_merges as merge_wire;

use super::{
    Budget, Package, TableMergesError, TableMergesLimitKind, decode_model_merges,
    select_semantic_positions,
};
use crate::table::merge::Region;
use crate::{SheetSelector, TableSelector};

/// A mutable, selector-first merge state staged against one immutable package.
///
/// The state starts with the source order.  Removing an existing region and
/// adding it again keeps the final source order; newly added regions retain
/// their staging order after all surviving source regions.
pub struct TableMergesEdit<'a> {
    source: &'a Package,
    sheet_position: usize,
    table_position: usize,
    before: Vec<Region>,
    regions: Vec<Region>,
    max_regions: usize,
    budget: Budget,
    rows: u32,
    columns: u32,
}

impl fmt::Debug for TableMergesEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("TableMergesEdit")
            .field("sheet_position", &self.sheet_position)
            .field("table_position", &self.table_position)
            .field("regions", &self.regions)
            .finish_non_exhaustive()
    }
}

impl TableMergesEdit<'_> {
    /// Return the currently staged merged regions.
    #[must_use]
    pub fn regions(&self) -> &[Region] {
        &self.regions
    }

    /// Stage one new merged-cell rectangle.
    ///
    /// A rectangle that intersects any staged rectangle is rejected.  This
    /// includes an equal rectangle, matching Numbers' merge operation.
    ///
    /// This changes merge geometry while preserving the underlying cell values.
    pub fn merge(&mut self, region: Region) -> Result<&mut Self, TableMergesError> {
        if region.end_row() >= self.rows || region.end_column() >= self.columns {
            return Err(TableMergesError::InvalidRegion);
        }
        self.budget
            .charge_work(self.regions.len().saturating_add(1))?;
        if self
            .regions
            .iter()
            .copied()
            .any(|existing| existing.overlaps(region))
        {
            return Err(TableMergesError::OverlappingRegion);
        }
        let observed = self
            .regions
            .len()
            .checked_add(1)
            .ok_or(TableMergesError::InvalidSource)?;
        if observed > self.max_regions {
            return Err(TableMergesError::LimitExceeded {
                kind: TableMergesLimitKind::Regions,
                observed: usize_as_u64(observed),
                maximum: usize_as_u64(self.max_regions),
            });
        }
        self.regions
            .try_reserve(1)
            .map_err(|_| TableMergesError::Allocation { amount: 1 })?;
        self.regions.push(region);
        Ok(self)
    }

    /// Stage removal of one exact merged-cell rectangle.
    ///
    /// Missing rectangles are deliberately successful no-ops, matching the
    /// native Numbers operation's `false` result without exposing that host
    /// mutation API at this focused boundary.
    pub fn unmerge(&mut self, region: Region) -> Result<bool, TableMergesError> {
        self.budget
            .charge_work(self.regions.len().saturating_add(1))?;
        if let Some(index) = self
            .regions
            .iter()
            .position(|candidate| *candidate == region)
        {
            self.regions.remove(index);
            return Ok(true);
        }
        Ok(false)
    }

    /// Commit the staged state as a fully reopened exact-source package.
    pub fn commit(self) -> Result<TableMergesCommit, TableMergesError> {
        let source_catalog = physical_source(self.source)?;
        let source_bytes = source_catalog.__source_owner();
        let source_fingerprint = fingerprint(&source_bytes);
        let mut edit = self;
        let after = normalize_regions(&edit.before, &edit.regions, &mut edit.budget)?;

        if after == edit.before {
            return Ok(TableMergesCommit {
                package: edit.source.snapshot(),
                patch: TableMergesPatch {
                    source: source_bytes.clone(),
                    target: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    sheet_position: edit.sheet_position,
                    table_position: edit.table_position,
                    before: edit.before,
                    after,
                },
                diagnostics: TableMergesDiagnostics::unchanged(),
            });
        }

        if !source_catalog.source_is_exact() {
            return Err(TableMergesError::UnsupportedSource);
        }

        let package = rewrite(
            edit.source,
            edit.sheet_position,
            edit.table_position,
            edit.before.len(),
            &after,
            &mut edit.budget,
        )?;
        let target = physical_source(&package)?.__source_owner();
        let target_fingerprint = fingerprint(&target);
        let patch = TableMergesPatch {
            source: source_bytes,
            target,
            source_fingerprint,
            target_fingerprint,
            sheet_position: edit.sheet_position,
            table_position: edit.table_position,
            before: edit.before,
            after,
        };
        let diagnostics = TableMergesDiagnostics::published(&patch.before, &patch.after);
        Ok(TableMergesCommit {
            package,
            patch,
            diagnostics,
        })
    }
}

/// An exact-source reversible table-merges patch.
#[derive(Clone, PartialEq, Eq)]
pub struct TableMergesPatch {
    source: SharedBytes,
    target: SharedBytes,
    source_fingerprint: u64,
    target_fingerprint: u64,
    sheet_position: usize,
    table_position: usize,
    before: Vec<Region>,
    after: Vec<Region>,
}

impl fmt::Debug for TableMergesPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("TableMergesPatch")
            .field("sheet_position", &self.sheet_position)
            .field("table_position", &self.table_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl TableMergesPatch {
    /// Return the source merge state required before applying this patch.
    #[must_use]
    pub fn before(&self) -> &[Region] {
        &self.before
    }

    /// Return the merge state produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[Region] {
        &self.after
    }

    /// Return whether this patch preserves both semantic state and bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source.as_ref() == self.target.as_ref()
            && self.source_fingerprint == self.target_fingerprint
    }

    /// Return an inverse patch that restores the exact source artifact.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            sheet_position: self.sheet_position,
            table_position: self.table_position,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Compact evidence describing one committed merge transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableMergesDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
    regions_added: usize,
    regions_removed: usize,
}

impl TableMergesDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
            regions_added: 0,
            regions_removed: 0,
        }
    }

    fn published(before: &[Region], after: &[Region]) -> Self {
        let regions_added = after
            .iter()
            .filter(|region| !before.contains(region))
            .count();
        let regions_removed = before
            .iter()
            .filter(|region| !after.contains(region))
            .count();
        Self {
            changed: true,
            touched_components: 1,
            full_reparse_performed: true,
            regions_added,
            regions_removed,
        }
    }

    /// Return whether the committed package differs semantically from source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }

    /// Return the number of new merged regions published.
    #[must_use]
    pub const fn regions_added(self) -> usize {
        self.regions_added
    }

    /// Return the number of removed merged regions published.
    #[must_use]
    pub const fn regions_removed(self) -> usize {
        self.regions_removed
    }
}

/// The fully reopened immutable result of one merge transaction.
#[must_use = "a Numbers table-merges commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct TableMergesCommit {
    package: Package,
    patch: TableMergesPatch,
    diagnostics: TableMergesDiagnostics,
}

impl TableMergesCommit {
    /// Borrow the fully reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &TableMergesPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &TableMergesDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Start a selector-first transaction over one rooted table's merge state.
    pub fn edit_table_merges<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<TableMergesEdit<'_>, TableMergesError> {
        let mut budget = Budget::new(self)?;
        let (sheet_position, table_position) =
            select_semantic_positions(self, sheet.into(), table.into(), &mut budget)?;
        let target = super::super::table_headers::resolve::resolve_target(
            self,
            sheet_position,
            table_position,
        )
        .map_err(map_header_error)?;
        let model = super::super::table_headers::rewrite::selected_payload(self, target)
            .map_err(map_header_error)?;
        let before = decode_model_merges(model, &mut budget)?;
        let regions = clone_regions(&before)?;
        Ok(TableMergesEdit {
            source: self,
            sheet_position,
            table_position,
            before,
            regions,
            max_regions: budget.max_regions,
            budget,
            rows: target.rows,
            columns: target.columns,
        })
    }

    /// Apply an exact-source-checked reversible merge patch.
    pub fn apply_table_merges(
        &self,
        patch: &TableMergesPatch,
    ) -> Result<TableMergesCommit, TableMergesError> {
        let source_catalog = physical_source(self)?;
        if fingerprint(source_catalog.source_bytes()) != patch.source_fingerprint
            || source_catalog.source_bytes() != patch.source.as_ref()
        {
            return Err(TableMergesError::PatchConflict);
        }
        let mut budget = Budget::new(self)?;
        let current = read_regions_at(
            self,
            patch.sheet_position,
            patch.table_position,
            &mut budget,
        )?;
        if current != patch.before {
            return Err(TableMergesError::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source.as_ref() != patch.target.as_ref()
                || patch.source_fingerprint != patch.target_fingerprint
            {
                return Err(TableMergesError::PatchConflict);
            }
            return Ok(TableMergesCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: TableMergesDiagnostics::unchanged(),
            });
        }
        if !source_catalog.source_is_exact()
            || fingerprint(&patch.target) != patch.target_fingerprint
        {
            return Err(TableMergesError::PatchConflict);
        }
        budget.charge_work(patch.target.len().saturating_mul(2))?;
        let candidate =
            Package::from_source_owner_with_options(patch.target.clone(), self.state.options)
                .map_err(super::map_package_error)?;
        let actual = read_regions_at(
            &candidate,
            patch.sheet_position,
            patch.table_position,
            &mut budget,
        )?;
        if actual != patch.after {
            return Err(TableMergesError::Verification);
        }
        Ok(TableMergesCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: TableMergesDiagnostics::published(&patch.before, &patch.after),
        })
    }
}

fn rewrite(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    source_regions: usize,
    after: &[Region],
    budget: &mut Budget,
) -> Result<Package, TableMergesError> {
    let target = super::super::table_headers::resolve::resolve_target(
        source,
        sheet_position,
        table_position,
    )
    .map_err(map_header_error)?;
    let source_catalog = physical_source(source)?;
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(TableMergesError::InvalidSource)?;
    let component_name = component.name().to_owned();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(TableMergesError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(TableMergesError::UnsupportedSource);
    }
    let physical_limits = source_catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(super::map_archive_error)?;
    // The source catalog already owns a parsed archive for this component.
    // Charge its complete inventory and exact encoded length before opening
    // the compressed entry again.  This keeps decompression and archive
    // parsing behind the same aggregate transaction ledger.
    let cached_archive = component.archive();
    let cached_object_count = cached_archive.objects.len();
    let cached_message_count = archive_message_count(cached_archive)?;
    let cached_encoded_len = cached_archive
        .encoded_len_with_limits(archive_limits)
        .map_err(super::map_core_error)?;
    budget.charge_objects(cached_object_count)?;
    budget.charge_messages(cached_message_count)?;
    budget.charge_work(entry.data().len())?;
    budget.charge_work(cached_encoded_len)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits
            .snappy_limits()
            .map_err(super::map_archive_error)?,
    )
    .map_err(super::map_core_error)?;
    if stream.as_bytes().len() != cached_encoded_len {
        return Err(TableMergesError::InvalidSource);
    }
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(super::map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(super::map_core_error)?;
    if archive.objects.len() != cached_object_count
        || archive_message_count(&archive)? != cached_message_count
    {
        return Err(TableMergesError::InvalidSource);
    }
    drop(stream);

    let source_payload = archive
        .objects
        .get(target.object_index)
        .and_then(|object| object.messages.get(target.message_index))
        .filter(|message| message.type_ == target.message_type)
        .map(|message| message.data.as_slice())
        .ok_or(TableMergesError::InvalidSource)?;
    if archive
        .objects
        .get(target.object_index)
        .is_none_or(|object| object.archive_info.identifier != Some(target.model_identifier))
    {
        return Err(TableMergesError::InvalidSource);
    }
    let limits = write_limits(
        source_payload.len(),
        after.len(),
        source_regions,
        archive_limits.max_archive_bytes(),
        budget,
    )?;
    let written =
        merge_wire::rewrite_table_merges(source_payload, after, generate_guid_bytes(), limits)
            .map_err(|error| {
                super::charge_attempted(budget, &error);
                super::map_merge_error(error)
            })?;
    budget.charge_wire_report(
        written.report.input_bytes(),
        written.report.fields(),
        written.report.work(),
    )?;
    if !written.changed || written.data == source_payload {
        return Err(TableMergesError::Verification);
    }
    budget.charge_retained(written.data.len())?;
    let retained = retain_bytes(&written.data)?;
    let object = archive
        .objects
        .get_mut(target.object_index)
        .ok_or(TableMergesError::InvalidSource)?;
    super::super::table_headers::resolve::validate_message_metadata(object, target.message_index)
        .map_err(map_header_error)?;
    object
        .replace_message_preserving_header_with_limits(
            target.message_index,
            RawMessage {
                type_: target.message_type,
                data: written.data,
            },
            archive_limits,
        )
        .map_err(super::map_core_error)?;
    let encoded_bound = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(super::map_core_error)?;
    budget.charge_work(encoded_bound)?;
    budget.charge_scratch(encoded_bound)?;
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(super::map_core_error)?;
    if rewritten.len() > encoded_bound {
        return Err(TableMergesError::Verification);
    }
    let snappy_limits = physical_limits
        .snappy_limits()
        .map_err(super::map_archive_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_bound).map_err(super::map_core_error)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(TableMergesError::LimitExceeded {
            kind: TableMergesLimitKind::EntryBytes,
            observed: usize_as_u64(compressed_bound),
            maximum: usize_as_u64(snappy_limits.max_compressed_stream()),
        });
    }
    budget.charge_work(compressed_bound)?;
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten).map_err(super::map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(TableMergesError::Verification);
    }
    drop(rewritten);
    let edits = [EntryEdit::new(&component_name, &compressed)];
    let source_bytes_len = source.source_bytes().len();
    budget.charge_work(source_bytes_len)?;
    let prepared = source_catalog
        .package()
        .prepare_reassembly(&edits, physical_limits)
        .map_err(super::map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.charge_allocations(requirements.allocations())?;
    budget.charge_scratch(requirements.scratch_bytes())?;
    budget.charge_work(
        requirements
            .output_bytes()
            .saturating_mul(2)
            .saturating_add(requirements.scratch_bytes()),
    )?;
    let (candidate_objects, candidate_messages, candidate_stream_bytes) = candidate_inventory(
        &source.state.components,
        &component_name,
        cached_object_count,
        cached_message_count,
        encoded_bound,
        archive_limits,
    )?;
    budget.charge_objects(candidate_objects)?;
    budget.charge_messages(candidate_messages)?;
    budget.charge_work(
        candidate_stream_bytes
            .saturating_add(candidate_objects)
            .saturating_add(source.state.components.catalog().len()),
    )?;
    // The prepared plan reports the exact candidate ZIP length, so account
    // for its final source-owner allocation before execution materializes it.
    budget.charge_work(requirements.output_bytes())?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(super::map_archive_error)?;
    drop(compressed);
    let candidate = Package::from_owned_bytes_with_options(output, source.state.options)
        .map_err(super::map_package_error)?;
    let source_previews =
        super::super::table_headers::rewrite::root_preview_deletions(source_catalog)
            .map_err(map_header_error)?;
    budget.charge_work(source.source_bytes().len())?;
    super::super::table_headers::rewrite::verify_exact_locality(
        source,
        &candidate,
        target,
        &source_previews,
        source_previews.len(),
        &retained,
    )
    .map_err(map_header_error)?;
    let actual = read_regions_at(&candidate, sheet_position, table_position, budget)?;
    if actual != after {
        return Err(TableMergesError::Verification);
    }
    Ok(candidate)
}

fn normalize_regions(
    before: &[Region],
    desired: &[Region],
    budget: &mut Budget,
) -> Result<Vec<Region>, TableMergesError> {
    let mut normalized = Vec::new();
    normalized
        .try_reserve_exact(desired.len())
        .map_err(|_| TableMergesError::Allocation {
            amount: desired.len(),
        })?;
    for region in before {
        budget.charge_work(desired.len())?;
        if desired.contains(region) {
            normalized.push(*region);
        }
    }
    for region in desired {
        budget.charge_work(before.len().saturating_add(1))?;
        if !before.contains(region) {
            normalized.push(*region);
        }
    }
    Ok(normalized)
}

fn clone_regions(source: &[Region]) -> Result<Vec<Region>, TableMergesError> {
    let mut clone = Vec::new();
    clone
        .try_reserve_exact(source.len())
        .map_err(|_| TableMergesError::Allocation {
            amount: source.len(),
        })?;
    clone.extend_from_slice(source);
    Ok(clone)
}

fn archive_message_count(archive: &Archive) -> Result<usize, TableMergesError> {
    archive.objects.iter().try_fold(0usize, |count, object| {
        count
            .checked_add(object.messages.len())
            .ok_or(TableMergesError::InvalidSource)
    })
}

fn candidate_inventory(
    components: &super::super::Components,
    modified_component: &str,
    modified_objects: usize,
    modified_messages: usize,
    modified_stream_bytes: usize,
    archive_limits: Limits,
) -> Result<(usize, usize, usize), TableMergesError> {
    let mut objects = modified_objects;
    let mut messages = modified_messages;
    let mut stream_bytes = modified_stream_bytes;
    for (name, archive) in components.iter_archives() {
        if name == modified_component {
            continue;
        }
        objects = objects
            .checked_add(archive.objects.len())
            .ok_or(TableMergesError::InvalidSource)?;
        messages = messages
            .checked_add(archive_message_count(archive)?)
            .ok_or(TableMergesError::InvalidSource)?;
        stream_bytes = stream_bytes
            .checked_add(
                archive
                    .encoded_len_with_limits(archive_limits)
                    .map_err(super::map_core_error)?,
            )
            .ok_or(TableMergesError::InvalidSource)?;
    }
    Ok((objects, messages, stream_bytes))
}

fn read_regions_at(
    package: &Package,
    sheet_position: usize,
    table_position: usize,
    budget: &mut Budget,
) -> Result<Vec<Region>, TableMergesError> {
    let (sheet_position, table_position) = select_semantic_positions(
        package,
        SheetSelector::index(sheet_position),
        TableSelector::index(table_position),
        budget,
    )?;
    let model = super::resolve_model_payload(
        &package.state.components,
        &package.state.index,
        sheet_position,
        table_position,
        budget,
    )?;
    decode_model_merges(model, budget)
}

fn retain_bytes(source: &[u8]) -> Result<Arc<[u8]>, TableMergesError> {
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(source.len())
        .map_err(|_| TableMergesError::Allocation {
            amount: source.len(),
        })?;
    retained.extend_from_slice(source);
    Ok(retained.into())
}

fn write_limits(
    source_bytes: usize,
    regions: usize,
    source_regions: usize,
    maximum_output: usize,
    budget: &Budget,
) -> Result<merge_wire::ReadLimits, TableMergesError> {
    let residual = budget.residual_wire()?;
    let output = maximum_output
        .max(source_bytes.saturating_add(1))
        .min(WireLimits::MAX_OUTPUT_BYTES);
    let wire = residual
        .with_output_bytes(output)
        .map_err(super::map_common_error)?;
    Ok(merge_wire::ReadLimits {
        wire,
        max_regions: regions.max(source_regions).clamp(1, WireLimits::MAX_FIELDS),
        max_overlap_checks: residual.max_rewrite_work(),
    })
}

fn physical_source(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, TableMergesError> {
    super::super::table_headers::rewrite::physical_source(package).map_err(map_header_error)
}

fn map_header_error(error: super::super::table_headers::Error) -> TableMergesError {
    use super::super::table_headers::{Error, LimitKind};
    match error {
        Error::SheetNotFound => TableMergesError::SheetNotFound,
        Error::TableNotFound => TableMergesError::TableNotFound,
        Error::UnsupportedSource => TableMergesError::UnsupportedSource,
        Error::LimitExceeded {
            kind,
            observed,
            maximum,
            ..
        } => TableMergesError::LimitExceeded {
            kind: match kind {
                LimitKind::InputBytes => TableMergesLimitKind::InputBytes,
                LimitKind::OutputBytes => TableMergesLimitKind::OutputBytes,
                LimitKind::Entries => TableMergesLimitKind::Entries,
                LimitKind::EntryBytes => TableMergesLimitKind::EntryBytes,
                LimitKind::TotalEntryBytes => TableMergesLimitKind::TotalEntryBytes,
                LimitKind::PackageBytes => TableMergesLimitKind::PackageBytes,
                LimitKind::PayloadBytes => TableMergesLimitKind::PayloadBytes,
                LimitKind::TotalPayloadBytes => TableMergesLimitKind::TotalPayloadBytes,
                LimitKind::PayloadObjects => TableMergesLimitKind::PayloadObjects,
                LimitKind::PayloadMessages => TableMergesLimitKind::PayloadMessages,
                LimitKind::PayloadItems => TableMergesLimitKind::PayloadItems,
                LimitKind::PayloadReferences => TableMergesLimitKind::PayloadReferences,
                LimitKind::WireBytes => TableMergesLimitKind::WireBytes,
                LimitKind::WireOutputBytes => TableMergesLimitKind::WireOutputBytes,
                LimitKind::WireFields => TableMergesLimitKind::WireFields,
                LimitKind::WireNesting => TableMergesLimitKind::WireNesting,
                LimitKind::WireWork => TableMergesLimitKind::WireWork,
                LimitKind::TransactionWork => TableMergesLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        Error::Allocation { amount, .. } => TableMergesError::Allocation { amount },
        Error::Verification => TableMergesError::Verification,
        Error::PatchConflict => TableMergesError::PatchConflict,
        Error::InvalidSource { .. }
        | Error::InvalidSettings { .. }
        | Error::TableLocked { .. }
        | Error::UnsupportedDependency { .. } => TableMergesError::InvalidSource,
    }
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

fn usize_as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
