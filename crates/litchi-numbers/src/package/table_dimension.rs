//! Exact-source selector-first row and column size transactions.

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::numbers_table_cell_storage_codec as codec;

use super::{Package, table_headers};
use crate::table::dimension::transaction::{
    Commit, Diagnostics, Edit, Evidence, LimitKind, Patch, Path, TransactionError,
};
use crate::{
    selector::{SheetSelector, TableSelector},
    table::{
        dimension::{Dimension, Points, Size},
        lock::State as LockState,
    },
};

const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;

#[cfg(test)]
mod phase_observer {
    use std::sync::atomic::{AtomicUsize, Ordering};

    pub(super) static COMPONENT_ENCODE: AtomicUsize = AtomicUsize::new(0);
    pub(super) static REASSEMBLY: AtomicUsize = AtomicUsize::new(0);
    pub(super) static OUTPUT: AtomicUsize = AtomicUsize::new(0);
    pub(super) static REOPEN: AtomicUsize = AtomicUsize::new(0);
    pub(super) static LOCALITY: AtomicUsize = AtomicUsize::new(0);
    pub(super) static FINGERPRINT: AtomicUsize = AtomicUsize::new(0);

    pub(super) fn hit(counter: &AtomicUsize) {
        counter.fetch_add(1, Ordering::Relaxed);
    }

    pub(super) fn reset() {
        for counter in [
            &COMPONENT_ENCODE,
            &REASSEMBLY,
            &OUTPUT,
            &REOPEN,
            &LOCALITY,
            &FINGERPRINT,
        ] {
            counter.store(0, Ordering::Relaxed);
        }
    }

    pub(super) fn snapshot() -> [usize; 6] {
        [
            COMPONENT_ENCODE.load(Ordering::Relaxed),
            REASSEMBLY.load(Ordering::Relaxed),
            OUTPUT.load(Ordering::Relaxed),
            REOPEN.load(Ordering::Relaxed),
            LOCALITY.load(Ordering::Relaxed),
            FINGERPRINT.load(Ordering::Relaxed),
        ]
    }
}

#[cfg(test)]
macro_rules! phase {
    ($counter:ident) => {
        phase_observer::hit(&phase_observer::$counter)
    };
}

#[cfg(not(test))]
macro_rules! phase {
    ($counter:ident) => {};
}

#[derive(Clone, Copy)]
struct Selected {
    target: table_headers::Target,
    dimension: Dimension,
    evidence: Evidence,
    size: Size,
    budget: ProjectionBudget,
}

#[derive(Clone, Copy)]
struct ProjectionBudget {
    fields: usize,
    work: usize,
    references: usize,
    text: usize,
    maximum_fields: usize,
    maximum_work: usize,
    maximum_references: usize,
    maximum_text: usize,
    path: Path,
}

impl ProjectionBudget {
    fn new(source: &Package, path: Path) -> Self {
        Self {
            fields: 0,
            work: 0,
            references: 0,
            text: 0,
            maximum_fields: usize::try_from(source.state.options.archive().max_total_bytes())
                .unwrap_or(usize::MAX),
            maximum_work: usize::try_from(source.state.options.archive().max_total_bytes())
                .unwrap_or(usize::MAX),
            maximum_references: source.state.options.semantic().max_references(),
            maximum_text: source.state.options.semantic().max_output_text_bytes(),
            path,
        }
    }
    fn options(&self, bytes: usize) -> codec::DecodeOptions {
        codec::DecodeOptions::new(
            bytes.max(1),
            self.maximum_fields.saturating_sub(self.fields),
            self.maximum_work.saturating_sub(self.work),
            16,
            self.maximum_references.saturating_sub(self.references),
            self.maximum_text.saturating_sub(self.text),
        )
    }
    fn charge(&mut self, report: codec::DecodeReport) -> Result<(), TransactionError> {
        self.fields = charge(
            self.fields,
            report.fields(),
            self.maximum_fields,
            LimitKind::WireFields,
            self.path,
        )?;
        self.work = charge(
            self.work,
            report.work_bytes(),
            self.maximum_work,
            LimitKind::WireWork,
            self.path,
        )?;
        self.references = charge(
            self.references,
            report.references(),
            self.maximum_references,
            LimitKind::PayloadReferences,
            self.path,
        )?;
        self.text = charge(
            self.text,
            report.text_bytes(),
            self.maximum_text,
            LimitKind::PayloadBytes,
            self.path,
        )?;
        Ok(())
    }

    fn charge_upper_bound(
        &mut self,
        bound: codec::DecodeResourceUpperBound,
    ) -> Result<(), TransactionError> {
        self.fields = charge(
            self.fields,
            bound.fields(),
            self.maximum_fields,
            LimitKind::WireFields,
            self.path,
        )?;
        self.work = charge(
            self.work,
            bound.work_bytes(),
            self.maximum_work,
            LimitKind::WireWork,
            self.path,
        )?;
        self.references = charge(
            self.references,
            bound.references(),
            self.maximum_references,
            LimitKind::PayloadReferences,
            self.path,
        )?;
        self.text = charge(
            self.text,
            bound.text_bytes(),
            self.maximum_text,
            LimitKind::PayloadBytes,
            self.path,
        )?;
        Ok(())
    }

    fn charge_source_axes(&mut self, report: codec::DecodeReport) -> Result<(), TransactionError> {
        self.fields = charge(
            self.fields,
            report.fields(),
            self.maximum_fields,
            LimitKind::WireFields,
            self.path,
        )?;
        self.references = charge(
            self.references,
            report.references(),
            self.maximum_references,
            LimitKind::PayloadReferences,
            self.path,
        )?;
        self.text = charge(
            self.text,
            report.text_bytes(),
            self.maximum_text,
            LimitKind::PayloadBytes,
            self.path,
        )?;
        Ok(())
    }

    fn map_error(self, error: codec::DecodeError) -> TransactionError {
        map_codec_error_with_offsets(
            error,
            self.path,
            self.fields,
            self.work,
            self.references,
            self.text,
        )
    }
}

fn charge(
    current: usize,
    amount: usize,
    maximum: usize,
    kind: LimitKind,
    path: Path,
) -> Result<usize, TransactionError> {
    let observed = current.checked_add(amount).unwrap_or(usize::MAX);
    if observed > maximum {
        return Err(TransactionError::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path,
        });
    }
    Ok(observed)
}

impl Package {
    /// Read one rooted row or column's explicit/default size.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, allocation, or resource error.
    pub fn table_dimension_size<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        dimension: Dimension,
    ) -> Result<Size, TransactionError> {
        Ok(resolve(self, sheet, table, dimension)?.size)
    }

    /// Start one selector-first immutable row or column size edit.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, allocation, or resource error.
    pub fn edit_table_dimension_size<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        dimension: Dimension,
    ) -> Result<Edit<'_>, TransactionError> {
        let selected = resolve(self, sheet, table, dimension)?;
        Ok(Edit {
            source: self,
            sheet_position: selected.target.sheet_position,
            table_position: selected.target.table_position,
            dimension,
            before: selected.size,
            size: selected.size,
            evidence: selected.evidence,
        })
    }

    /// Apply a reversible patch to its exact retained source artifact.
    ///
    /// # Errors
    ///
    /// Returns a conflict unless this package is the patch's exact source.
    pub fn apply_table_dimension_size(&self, patch: &Patch) -> Result<Commit, TransactionError> {
        let target_bytes = patch.artifacts.target_owner();
        preflight_transaction_work(self, Some(target_bytes.as_ref()))?;
        let catalog = physical_source(self)?;
        let source = catalog.__source_owner();
        phase!(FINGERPRINT);
        if !patch.artifacts.authorizes_owner(&source) {
            return Err(TransactionError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: unchanged(),
            });
        }
        let selected = resolve_at(
            self,
            patch.sheet_position,
            patch.table_position,
            patch.dimension,
        )?;
        if selected.evidence != patch.evidence || selected.size != patch.before {
            return Err(TransactionError::PatchConflict);
        }
        let candidate = Package::from_source_owner_with_options(target_bytes, self.state.options)
            .map_err(map_read_error)?;
        let after = resolve_at(
            &candidate,
            patch.sheet_position,
            patch.table_position,
            patch.dimension,
        )?;
        if after.size != patch.after || after.evidence != patch.evidence {
            return Err(TransactionError::Verification);
        }
        let target_payload = selected_payload(&candidate, patch.evidence)?;
        verify_locality(
            self,
            &candidate,
            patch.evidence,
            patch.source_previews,
            patch.target_previews,
            target_payload,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: published(patch.source_previews.saturating_sub(patch.target_previews)),
        })
    }
}

pub(crate) fn commit(edit: Edit<'_>) -> Result<Commit, TransactionError> {
    let catalog = physical_source(edit.source)?;
    if edit.before == edit.size {
        preflight_transaction_work(edit.source, None)?;
        let source = catalog.__source_owner();
        phase!(FINGERPRINT);
        return Ok(Commit {
            package: edit.source.snapshot(),
            patch: Patch {
                artifacts: litchi_iwa_archive::package::OwnedExactArtifacts::new(
                    source.clone(),
                    source,
                ),
                sheet_position: edit.sheet_position,
                table_position: edit.table_position,
                dimension: edit.dimension,
                before: edit.before,
                after: edit.size,
                evidence: edit.evidence,
                touched_components: 0,
                source_previews: 0,
                target_previews: 0,
            },
            diagnostics: unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(TransactionError::UnsupportedSource);
    }
    let selected = resolve_at(
        edit.source,
        edit.sheet_position,
        edit.table_position,
        edit.dimension,
    )?;
    if selected.evidence != edit.evidence || selected.size != edit.before {
        return Err(TransactionError::InvalidSource { path: edit.path() });
    }
    if selected.target.locked == LockState::Locked {
        return Err(TransactionError::TableLocked { path: edit.path() });
    }
    let source = catalog.__source_owner();
    let previews = root_previews(catalog)?;
    let package = rewrite(edit.source, selected, edit.size, &previews)?;
    let target_payload = selected_payload(&package, edit.evidence)?;
    verify_locality(
        edit.source,
        &package,
        edit.evidence,
        previews.len(),
        0,
        target_payload,
    )?;
    let target = physical_source(&package)?.__source_owner();
    phase!(FINGERPRINT);
    Ok(Commit {
        package,
        patch: Patch {
            artifacts: litchi_iwa_archive::package::OwnedExactArtifacts::new(source, target),
            sheet_position: edit.sheet_position,
            table_position: edit.table_position,
            dimension: edit.dimension,
            before: edit.before,
            after: edit.size,
            evidence: edit.evidence,
            touched_components: 1,
            source_previews: previews.len(),
            target_previews: 0,
        },
        diagnostics: published(previews.len()),
    })
}

fn resolve<'sheet, 'table>(
    source: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    table: impl Into<TableSelector<'table>>,
    dimension: Dimension,
) -> Result<Selected, TransactionError> {
    let selected_sheet = source
        .state
        .document
        .sheet(sheet)
        .map_err(|_| TransactionError::InvalidSource {
            path: Path::Package,
        })?
        .ok_or(TransactionError::SheetNotFound)?;
    let table_position = match table.into() {
        TableSelector::Index(index) => selected_sheet.tables().nth(index).map(|_| index),
        TableSelector::Name(name) => {
            let mut matches = selected_sheet
                .tables()
                .enumerate()
                .filter(|(_, candidate)| candidate.name() == name);
            let first = matches.next().map(|(index, _)| index);
            if matches.next().is_some() {
                return Err(TransactionError::AmbiguousSelector);
            }
            first
        },
    }
    .ok_or(TransactionError::TableNotFound)?;
    resolve_at(source, selected_sheet.index(), table_position, dimension)
}

fn resolve_at(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    dimension: Dimension,
) -> Result<Selected, TransactionError> {
    let path = Path::Dimension {
        sheet: sheet_position,
        table: table_position,
        dimension,
    };
    let mut budget = ProjectionBudget::new(source, path);
    let transaction_work = preflight_transaction_work(source, None)?;
    budget.work = charge(
        budget.work,
        transaction_work,
        budget.maximum_work,
        LimitKind::TransactionWork,
        path,
    )?;
    let target = table_headers::resolve::resolve_target(source, sheet_position, table_position)
        .map_err(map_header_error)?;
    let limit = match dimension {
        Dimension::Row(_) => target.rows,
        Dimension::Column(_) => target.columns,
    };
    let index = u32::try_from(dimension.index()).map_err(|_| TransactionError::OutOfBounds {
        path,
        length: limit,
    })?;
    if index >= limit {
        return Err(TransactionError::OutOfBounds {
            path,
            length: limit,
        });
    }
    let ownership =
        table_headers::ownership::validate_selected_ownership_with_report(source, target)
            .map_err(map_header_error)?;
    let model =
        table_headers::rewrite::selected_payload(source, target).map_err(map_header_error)?;
    budget.work = charge(
        budget.work,
        ownership.work.saturating_add(ownership.transaction_work),
        budget.maximum_work,
        LimitKind::TransactionWork,
        path,
    )?;
    budget.references = charge(
        budget.references,
        ownership.references,
        budget.maximum_references,
        LimitKind::PayloadReferences,
        path,
    )?;
    let (model_snapshot, report) =
        codec::decode_table_model_with_report(model, budget.options(model.len()))
            .map_err(|error| budget.map_error(error))?;
    budget.charge(report)?;
    if model_snapshot.number_of_rows() != target.rows
        || model_snapshot.number_of_columns() != target.columns
    {
        return Err(TransactionError::InvalidSource { path });
    }
    let store = model_snapshot.base_data_store();
    let (data_store, report) =
        codec::decode_data_store_with_report(store, budget.options(store.len()))
            .map_err(|error| budget.map_error(error))?;
    budget.charge(report)?;
    let model_object = object_at(source, target.component_index, target.object_index, path)?;
    let column_id = checked_reference(data_store.column_headers(), path)?;
    let selected_slot = match dimension {
        Dimension::Row(_) => {
            usize::try_from(index / 65_536).map_err(|_| TransactionError::InvalidSource { path })?
        },
        Dimension::Column(_) => usize::MAX,
    };
    let mut collector = BucketCollector {
        selected: selected_slot,
        seen: 0,
        bucket: None,
        references: Vec::new(),
        allocation_failed: false,
    };
    let (_, report) = codec::decode_header_storage_with_visitor(
        data_store.row_headers(),
        budget.options(data_store.row_headers().len()),
        &mut collector,
    )
    .map_err(|error| budget.map_error(error))?;
    budget.charge(report)?;
    if collector.allocation_failed {
        return Err(TransactionError::Allocation {
            amount: collector.seen,
            path,
        });
    }
    let expected = usize::try_from(target.rows.div_ceil(65_536))
        .map_err(|_| TransactionError::InvalidSource { path })?;
    if collector.references.len() != expected {
        return Err(TransactionError::InvalidSource { path });
    }
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(collector.references.len())
        .map_err(|_| TransactionError::Allocation {
            amount: collector.references.len(),
            path,
        })?;
    for reference in &collector.references {
        let identifier = checked_reference(*reference, path)?;
        if identifier == column_id {
            return Err(TransactionError::InvalidSource { path });
        }
        table_headers::resolve::require_declared_reference(
            model_object,
            target.message_index,
            identifier,
            &[4, 1, 2],
        )
        .map_err(map_header_error)?;
        identifiers.push(identifier);
    }
    identifiers.sort_unstable();
    if identifiers.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(TransactionError::InvalidSource { path });
    }
    let info = model_object
        .archive_info
        .message_infos
        .get(target.message_index)
        .ok_or(TransactionError::InvalidSource { path })?;
    let mut declared = Vec::new();
    for field in &info.field_infos {
        if field.path.as_slice() == [4, 1, 2] {
            declared
                .try_reserve(field.object_references.len())
                .map_err(|_| TransactionError::Allocation {
                    amount: field.object_references.len(),
                    path,
                })?;
            declared.extend_from_slice(&field.object_references);
        } else if field
            .object_references
            .iter()
            .any(|id| identifiers.binary_search(id).is_ok())
        {
            return Err(TransactionError::InvalidSource { path });
        }
    }
    if !declared.is_empty() {
        declared.sort_unstable();
        if declared != identifiers {
            return Err(TransactionError::InvalidSource { path });
        }
    }
    let (bucket_id, declared_path): (u64, &[u32]) = match dimension {
        Dimension::Column(_) => (column_id, &[4, 2]),
        Dimension::Row(_) => (
            checked_reference(
                collector
                    .bucket
                    .ok_or(TransactionError::InvalidSource { path })?,
                path,
            )?,
            &[4, 1, 2],
        ),
    };
    table_headers::resolve::require_declared_reference(
        model_object,
        target.message_index,
        bucket_id,
        declared_path,
    )
    .map_err(map_header_error)?;
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, bucket_id)
        .map_err(map_read_error)?
        .ok_or(TransactionError::InvalidSource { path })?;
    let object = object_at(
        source,
        resolved.component_index,
        resolved.object_index,
        path,
    )?;
    if object.archive_info.identifier != Some(bucket_id) {
        return Err(TransactionError::InvalidSource { path });
    }
    let mut messages = resolved
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == HEADER_BUCKET_MESSAGE_TYPE);
    let (message_index, message) = messages
        .next()
        .ok_or(TransactionError::InvalidSource { path })?;
    if messages.next().is_some() {
        return Err(TransactionError::InvalidSource { path });
    }
    table_headers::resolve::validate_message_metadata(object, message_index)
        .map_err(map_header_error)?;
    if object.archive_info.message_infos.len() != object.messages.len() {
        return Err(TransactionError::InvalidSource { path });
    }
    let (minimum, maximum) = match dimension {
        Dimension::Row(_) => {
            let slot = index / 65_536;
            let minimum = slot.saturating_mul(65_536);
            (minimum, minimum.saturating_add(65_536).min(limit))
        },
        Dimension::Column(_) => (0, limit),
    };
    let mut reader = HeaderReader {
        index,
        found: None,
        duplicate: false,
        seen: Vec::with_capacity(0),
        allocation_failed: false,
        minimum,
        maximum,
    };
    let (_, report) = codec::decode_header_storage_bucket_with_visitor(
        &message.data,
        budget.options(message.data.len()),
        &mut reader,
    )
    .map_err(|error| budget.map_error(error))?;
    budget.charge(report)?;
    if reader.allocation_failed {
        return Err(TransactionError::Allocation {
            amount: reader.seen.len().saturating_add(1),
            path,
        });
    }
    reader.seen.sort_unstable();
    if reader.duplicate || reader.seen.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(TransactionError::InvalidSource { path });
    }
    let size = reader
        .found
        .map(size_from_bits)
        .transpose()?
        .unwrap_or(Size::Default);
    Ok(Selected {
        target,
        dimension,
        evidence: Evidence {
            model_component: target.component_index,
            model_object: target.object_index,
            model_message: target.message_index,
            model_identifier: target.model_identifier,
            bucket_component: resolved.component_index,
            bucket_object: resolved.object_index,
            bucket_message: message_index,
            bucket_identifier: bucket_id,
        },
        size,
        budget,
    })
}

struct BucketCollector {
    selected: usize,
    seen: usize,
    bucket: Option<codec::ReferenceSnapshot>,
    references: Vec<codec::ReferenceSnapshot>,
    allocation_failed: bool,
}
impl codec::StorageVisitor for BucketCollector {
    fn visit_header_bucket(
        &mut self,
        record: codec::ReferenceRecord<'_>,
    ) -> Result<(), codec::DecodeError> {
        if self.seen == self.selected {
            self.bucket = Some(record.reference());
        }
        if self.references.try_reserve(1).is_err() {
            self.allocation_failed = true;
            return Ok(());
        }
        self.references.push(record.reference());
        self.seen = self.seen.saturating_add(1);
        Ok(())
    }
}

struct HeaderReader {
    index: u32,
    found: Option<u32>,
    duplicate: bool,
    seen: Vec<u32>,
    minimum: u32,
    maximum: u32,
    allocation_failed: bool,
}
impl codec::StorageVisitor for HeaderReader {
    fn visit_header(&mut self, header: codec::HeaderSnapshot) -> Result<(), codec::DecodeError> {
        let size = f32::from_bits(header.size_bits());
        if header.index() < self.minimum
            || header.index() >= self.maximum
            || !size.is_finite()
            || size < 0.0
            || (size == 0.0 && header.size_bits() != 0)
        {
            self.duplicate = true;
        }
        if self.seen.try_reserve(1).is_err() {
            self.allocation_failed = true;
            return Ok(());
        }
        self.seen.push(header.index());
        if header.index() == self.index {
            if self.found.replace(header.size_bits()).is_some() {
                self.duplicate = true;
            }
        }
        Ok(())
    }
}

fn size_from_bits(bits: u32) -> Result<Size, TransactionError> {
    let value = f32::from_bits(bits);
    if value == 0.0 {
        if bits == 0 {
            return Ok(Size::Default);
        }
        return Err(TransactionError::InvalidSource {
            path: Path::Package,
        });
    }
    Points::new(value)
        .map(Size::Points)
        .map_err(|_| TransactionError::InvalidSource {
            path: Path::Package,
        })
}

fn rewrite(
    source: &Package,
    selected: Selected,
    after: Size,
    previews: &[&str],
) -> Result<Package, TransactionError> {
    let catalog = physical_source(source)?;
    let path = Path::Dimension {
        sheet: selected.target.sheet_position,
        table: selected.target.table_position,
        dimension: selected.dimension,
    };
    let payload = selected_payload(source, selected.evidence)?;
    let index = u32::try_from(selected.dimension.index())
        .map_err(|_| TransactionError::InvalidSource { path })?;
    let limit = match selected.dimension {
        Dimension::Row(_) => selected.target.rows,
        Dimension::Column(_) => selected.target.columns,
    };
    let edit = match after {
        Size::Default => codec::HeaderSizeEdit::remove(index),
        Size::Points(points) => codec::HeaderSizeEdit::set(index, points.value().to_bits()),
    };
    let mut budget = selected.budget;
    let planning_options = budget.options(payload.len().saturating_add(64));
    let planning_work = header_plan_work_upper_bound(payload.len())?;
    budget.work = charge(
        budget.work,
        planning_work,
        budget.maximum_work,
        LimitKind::TransactionWork,
        path,
    )?;
    let plan = codec::plan_header_storage_bucket_sizes(payload, limit, &[edit], planning_options)
        .map_err(|error| budget.map_error(error))?;
    let requirements = plan.requirements();
    let reported_planning_work = requirements
        .source()
        .work_bytes()
        .checked_add(requirements.rewrite_work_bytes())
        .unwrap_or(usize::MAX);
    if reported_planning_work > planning_work {
        return Err(TransactionError::InvalidSource { path });
    }
    budget.charge_source_axes(requirements.source())?;
    let execution_options = budget.options(requirements.output_bytes().max(1));
    budget.charge_upper_bound(requirements.result_upper_bound())?;
    let component = source
        .state
        .components
        .catalog()
        .get_index(selected.evidence.bucket_component)
        .ok_or(TransactionError::InvalidSource { path })?;
    let component_name = component.name();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(TransactionError::InvalidSource { path })?;
    if entry.is_opaque() {
        return Err(TransactionError::UnsupportedSource);
    }
    let physical_limits = catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let publication_work = publication_work_upper_bound(
        source.source_bytes().len(),
        stream.as_bytes().len(),
        requirements.output_bytes(),
        preflight_transaction_work(source, None)?,
    )?;
    budget.work = charge(
        budget.work,
        publication_work,
        budget.maximum_work,
        LimitKind::TransactionWork,
        path,
    )?;
    let (rewritten_payload, _report) =
        codec::execute_header_storage_bucket_size_plan(plan, execution_options)
            .map_err(|error| budget.map_error(error))?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    let object = archive
        .objects
        .get_mut(selected.evidence.bucket_object)
        .ok_or(TransactionError::InvalidSource { path })?;
    if object.archive_info.identifier != Some(selected.evidence.bucket_identifier) {
        return Err(TransactionError::InvalidSource { path });
    }
    table_headers::resolve::validate_message_metadata(object, selected.evidence.bucket_message)
        .map_err(map_header_error)?;
    if object
        .messages
        .get(selected.evidence.bucket_message)
        .is_none_or(|message| message.type_ != HEADER_BUCKET_MESSAGE_TYPE)
    {
        return Err(TransactionError::InvalidSource { path });
    }
    object
        .replace_message_preserving_header_with_limits(
            selected.evidence.bucket_message,
            RawMessage {
                type_: HEADER_BUCKET_MESSAGE_TYPE,
                data: rewritten_payload,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    phase!(COMPONENT_ENCODE);
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    phase!(REASSEMBLY);
    phase!(OUTPUT);
    let output = catalog
        .package()
        .reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            previews,
            physical_limits,
        )
        .map_err(map_archive_error)?;
    phase!(REOPEN);
    let candidate = Package::from_owned_bytes_with_options(output, source.state.options)
        .map_err(map_read_error)?;
    let reselected = resolve_at(
        &candidate,
        selected.target.sheet_position,
        selected.target.table_position,
        selected.dimension,
    )?;
    if reselected.size != after || reselected.evidence != selected.evidence {
        return Err(TransactionError::Verification);
    }
    Ok(candidate)
}

fn selected_payload(source: &Package, evidence: Evidence) -> Result<&[u8], TransactionError> {
    source
        .state
        .components
        .catalog()
        .get_index(evidence.bucket_component)
        .and_then(|component| component.archive().objects.get(evidence.bucket_object))
        .filter(|object| object.archive_info.identifier == Some(evidence.bucket_identifier))
        .and_then(|object| object.messages.get(evidence.bucket_message))
        .filter(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(TransactionError::InvalidSource {
            path: Path::Package,
        })
}

fn checked_reference(
    reference: codec::ReferenceSnapshot,
    path: Path,
) -> Result<u64, TransactionError> {
    if reference.identifier() == 0 || reference.deprecated_is_external() == Some(true) {
        return Err(TransactionError::InvalidSource { path });
    }
    Ok(reference.identifier())
}

fn object_at(
    source: &Package,
    component: usize,
    object: usize,
    path: Path,
) -> Result<&litchi_iwa_core::ArchiveObject, TransactionError> {
    source
        .state
        .components
        .catalog()
        .get_index(component)
        .and_then(|entry| entry.archive().objects.get(object))
        .ok_or(TransactionError::InvalidSource { path })
}

fn physical_source(
    source: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, TransactionError> {
    source
        .state
        .components
        .physical()
        .ok_or(TransactionError::UnsupportedSource)
}

fn root_previews(
    source: &litchi_iwa_archive::SourceCatalog,
) -> Result<Vec<&'static str>, TransactionError> {
    table_headers::rewrite::root_preview_deletions(source).map_err(map_header_error)
}

fn preflight_transaction_work(
    source: &Package,
    target: Option<&[u8]>,
) -> Result<usize, TransactionError> {
    table_headers::rewrite::preflight_transaction_work(source, target).map_err(map_header_error)
}

fn publication_work_upper_bound(
    source_bytes: usize,
    component_bytes: usize,
    replacement_payload_bytes: usize,
    reopen_projection_work: usize,
) -> Result<usize, TransactionError> {
    const COMPONENT_SLACK: usize = 4 * 1024;
    const PACKAGE_SLACK: usize = 8 * 1024;
    let archive_output = checked_add_work(
        checked_add_work(component_bytes, replacement_payload_bytes)?,
        COMPONENT_SLACK,
    )?;
    let compressed_output = maximum_snappy_output(archive_output)?;
    let package_output = checked_add_work(
        checked_add_work(source_bytes, compressed_output)?,
        PACKAGE_SLACK,
    )?;
    // Component encode + compression, package reassembly/output, candidate
    // reopen/projection, byte-exact locality, and both artifact fingerprints.
    let component = checked_add_work(archive_output.saturating_mul(2), compressed_output)?;
    let package = checked_add_work(source_bytes, package_output)?;
    let reopen = checked_add_work(package_output.saturating_mul(2), reopen_projection_work)?;
    let locality = checked_add_work(source_bytes, package_output)?;
    let fingerprints = checked_add_work(source_bytes, package_output)?;
    checked_add_work(
        checked_add_work(checked_add_work(component, package)?, reopen)?,
        checked_add_work(locality, fingerprints)?,
    )
}

fn header_plan_work_upper_bound(source_bytes: usize) -> Result<usize, TransactionError> {
    let log = if source_bytes <= 1 {
        0
    } else {
        usize::try_from(usize::BITS - (source_bytes - 1).leading_zeros()).unwrap_or(usize::MAX)
    };
    // Every record and edit consumes at least one source byte. This bounds
    // the strict/Buffa scans, fallible staging, duplicate-order sort, and all
    // binary searches before the codec exposes its exact requirements.
    let per_byte = checked_add_work(40, log.saturating_mul(4))?;
    checked_add_work(source_bytes.saturating_mul(per_byte), 256)
}

fn maximum_snappy_output(input_bytes: usize) -> Result<usize, TransactionError> {
    const RAW_OVERHEAD: usize = 32;
    const FRAME_HEADER: usize = 4;
    let mut total = 0usize;
    let mut remaining = input_bytes;
    while remaining != 0 {
        let chunk = remaining.min(SnappyStream::WRITE_CHUNK_SIZE);
        let compressed = checked_add_work(checked_add_work(RAW_OVERHEAD, chunk)?, chunk / 6)?;
        total = checked_add_work(total, checked_add_work(FRAME_HEADER, compressed)?)?;
        remaining = remaining
            .checked_sub(chunk)
            .ok_or(TransactionError::InvalidSource {
                path: Path::Package,
            })?;
    }
    Ok(total)
}

fn checked_add_work(left: usize, right: usize) -> Result<usize, TransactionError> {
    left.checked_add(right)
        .ok_or(TransactionError::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
            path: Path::Package,
        })
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    evidence: Evidence,
    source_previews: usize,
    target_previews: usize,
    payload: &[u8],
) -> Result<(), TransactionError> {
    phase!(LOCALITY);
    let before = physical_source(source)?;
    let after = physical_source(candidate)?;
    if root_previews(before)?.len() != source_previews
        || root_previews(after)?.len() != target_previews
    {
        return Err(TransactionError::Verification);
    }
    let before_preview_names = root_previews(before)?;
    let after_preview_names = root_previews(after)?;
    let mut before_entries = before
        .package()
        .iter()
        .filter(|entry| !before_preview_names.contains(&entry.name()));
    let mut after_entries = after
        .package()
        .iter()
        .filter(|entry| !after_preview_names.contains(&entry.name()));
    loop {
        match (before_entries.next(), after_entries.next()) {
            (Some(left), Some(right)) if left.name() == right.name() => {
                let selected = source
                    .state
                    .components
                    .catalog()
                    .get_index(evidence.bucket_component)
                    .is_some_and(|component| component.name() == left.name());
                let preserved = if selected {
                    table_headers::rewrite::selected_package_member_preserved(left, right)
                } else {
                    table_headers::rewrite::package_member_preserved(left, right)
                };
                if !preserved {
                    return Err(TransactionError::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(TransactionError::Verification),
        }
    }
    if source.state.components.catalog().len() != candidate.state.components.catalog().len() {
        return Err(TransactionError::Verification);
    }
    for (component_index, (left, right)) in source
        .state
        .components
        .catalog()
        .iter()
        .zip(candidate.state.components.catalog().iter())
        .enumerate()
    {
        if left.name() != right.name()
            || left.archive().objects.len() != right.archive().objects.len()
        {
            return Err(TransactionError::Verification);
        }
        for (object_index, (left_object, right_object)) in left
            .archive()
            .objects
            .iter()
            .zip(&right.archive().objects)
            .enumerate()
        {
            if component_index != evidence.bucket_component
                || object_index != evidence.bucket_object
            {
                if left_object != right_object {
                    return Err(TransactionError::Verification);
                }
                continue;
            }
            if left_object.archive_info.identifier != right_object.archive_info.identifier
                || left_object.messages.len() != right_object.messages.len()
            {
                return Err(TransactionError::Verification);
            }
            for (message_index, (left_message, right_message)) in left_object
                .messages
                .iter()
                .zip(&right_object.messages)
                .enumerate()
            {
                if message_index == evidence.bucket_message {
                    if right_message.type_ != HEADER_BUCKET_MESSAGE_TYPE
                        || right_message.data != payload
                    {
                        return Err(TransactionError::Verification);
                    }
                } else if left_message != right_message {
                    return Err(TransactionError::Verification);
                }
            }
        }
    }
    Ok(())
}

fn unchanged() -> Diagnostics {
    Diagnostics {
        changed: false,
        touched_components: 0,
        deleted_previews: 0,
        full_reparse_performed: false,
    }
}
fn published(deleted_previews: usize) -> Diagnostics {
    Diagnostics {
        changed: true,
        touched_components: 1,
        deleted_previews,
        full_reparse_performed: true,
    }
}

fn map_header_error(error: table_headers::Error) -> TransactionError {
    match error {
        table_headers::Error::SheetNotFound => TransactionError::SheetNotFound,
        table_headers::Error::TableNotFound => TransactionError::TableNotFound,
        table_headers::Error::UnsupportedSource => TransactionError::UnsupportedSource,
        table_headers::Error::TableLocked { .. } => TransactionError::TableLocked {
            path: Path::Package,
        },
        table_headers::Error::LimitExceeded {
            observed, maximum, ..
        } => TransactionError::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed,
            maximum,
            path: Path::Package,
        },
        table_headers::Error::Allocation { amount, .. } => TransactionError::Allocation {
            amount,
            path: Path::Package,
        },
        _ => TransactionError::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_read_error(error: super::Error) -> TransactionError {
    match error {
        super::Error::Archive(error) => map_archive_error(error),
        super::Error::Common(error) => map_common_error(error),
        super::Error::InputTooLarge { observed, maximum } => TransactionError::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed,
            maximum,
            path: Path::Package,
        },
        super::Error::SemanticLimit {
            observed, maximum, ..
        } => TransactionError::LimitExceeded {
            kind: LimitKind::PayloadItems,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        _ => TransactionError::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> TransactionError {
    match error {
        litchi_iwa_archive::Error::Allocation { amount, .. } => TransactionError::Allocation {
            amount,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Limit {
            observed, maximum, ..
        } => TransactionError::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed,
            maximum,
            path: Path::Package,
        },
        _ => TransactionError::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> TransactionError {
    match error {
        litchi_iwa_core::Error::Allocation { requested, .. } => TransactionError::Allocation {
            amount: requested,
            path: Path::Package,
        },
        litchi_iwa_core::Error::Limit {
            observed, maximum, ..
        } => TransactionError::LimitExceeded {
            kind: LimitKind::PayloadBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        _ => TransactionError::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_common_error(error: litchi_iwa_common::Error) -> TransactionError {
    match error {
        litchi_iwa_common::Error::Allocation { amount, .. } => TransactionError::Allocation {
            amount,
            path: Path::Package,
        },
        _ => TransactionError::InvalidSource {
            path: Path::Package,
        },
    }
}

fn map_codec_error_with_offsets(
    error: codec::DecodeError,
    path: Path,
    consumed_fields: usize,
    consumed_work: usize,
    consumed_references: usize,
    consumed_text: usize,
) -> TransactionError {
    let Some(limit) = error.resource_limit() else {
        return TransactionError::InvalidSource { path };
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::Bytes { observed, maximum } => {
            (LimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::References { observed, maximum } => (
            LimitKind::PayloadReferences,
            consumed_references.saturating_add(observed),
            consumed_references.saturating_add(maximum),
        ),
        codec::DecodeLimit::Text { observed, maximum } => (
            LimitKind::PayloadBytes,
            consumed_text.saturating_add(observed),
            consumed_text.saturating_add(maximum),
        ),
        codec::DecodeLimit::Fields { observed, maximum } => (
            LimitKind::WireFields,
            consumed_fields.saturating_add(observed),
            consumed_fields.saturating_add(maximum),
        ),
        codec::DecodeLimit::Work { observed, maximum } => (
            LimitKind::WireWork,
            consumed_work.saturating_add(observed),
            consumed_work.saturating_add(maximum),
        ),
        codec::DecodeLimit::Nesting { observed, maximum } => {
            return TransactionError::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
                path,
            };
        },
        codec::DecodeLimit::Allocation { requested } => {
            return TransactionError::Allocation {
                amount: requested,
                path,
            };
        },
        _ => return TransactionError::InvalidSource { path },
    };
    TransactionError::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        path,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{PackageLimits, PackageReadOptions, PackageSemanticLimits};

    fn fixture() -> std::path::PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers")
    }

    fn with_work_limit(
        source: &[u8],
        maximum: usize,
    ) -> Result<Package, Box<dyn std::error::Error>> {
        let limits = PackageLimits::new(
            PackageLimits::MAX_INPUT_BYTES,
            PackageLimits::MAX_ENTRIES,
            PackageLimits::MAX_ENTRY_BYTES,
            u64::try_from(maximum)?,
            PackageLimits::MAX_IWA_STREAM_BYTES,
        )?;
        Ok(Package::from_bytes_with_options(
            source,
            PackageReadOptions::new(limits, PackageSemanticLimits::default()),
        )?)
    }

    #[test]
    fn publication_max_minus_one_preempts_every_output_phase()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = std::fs::read(fixture())?;
        let mut lower = source.len();
        let mut upper = usize::try_from(PackageLimits::MAX_TOTAL_BYTES)?;
        while lower < upper {
            let middle = lower + (upper - lower) / 2;
            let succeeds = with_work_limit(&source, middle).is_ok_and(|package| {
                package
                    .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))
                    .and_then(|edit| edit.set(Size::points(124.0).unwrap()).commit())
                    .is_ok()
            });
            if succeeds {
                upper = middle;
            } else {
                lower = middle.saturating_add(1);
            }
        }
        let exact = lower;
        assert!(exact > source.len());
        let _exact_commit = with_work_limit(&source, exact)?
            .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))?
            .set(Size::points(124.0)?)
            .commit()?;

        let restricted = with_work_limit(&source, exact - 1)?;
        let edit = restricted
            .edit_table_dimension_size(0usize, 0usize, Dimension::Column(2))?
            .set(Size::points(124.0)?);
        phase_observer::reset();
        let error = edit
            .commit()
            .expect_err("max-minus-one publication work must refuse before output phases");
        assert!(matches!(
            error,
            TransactionError::LimitExceeded {
                kind: LimitKind::TransactionWork,
                observed,
                maximum,
                path: Path::Dimension {
                    sheet: 0,
                    table: 0,
                    dimension: Dimension::Column(2),
                },
            } if observed > maximum && maximum == u64::try_from(exact - 1)?
        ));
        assert_eq!(phase_observer::snapshot(), [0; 6]);
        Ok(())
    }
}
