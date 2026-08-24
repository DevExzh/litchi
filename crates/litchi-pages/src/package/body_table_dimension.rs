//! Exact-source row and column size transactions for rooted Pages tables.
//!
//! The table graph is resolved by [`super::table_lock`].  This module owns
//! only the archive-free dimension value and the narrow header-bucket rewrite;
//! object identifiers and wire records remain private proof data.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused Pages transaction keeps its proof beside the rewrite"
)]

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::{LimitKind as ArchiveLimitKind, SourceCatalog, package::EntryEdit};
use litchi_iwa_core::LimitKind as CoreLimitKind;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::table_dimension_codec as codec;
use thiserror::Error;

use super::{Package, PackageError, page_layout, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::dimension::{Dimension, Points, Size};

const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const HEADER_BUCKET_ROWS: usize = 65_536;
const ROOT_PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const MAX_VARINT_BYTES: usize = 10;

/// Finite resources charged by one body-table dimension transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableDimensionLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete candidate package output bytes.
    OutputBytes,
    /// ZIP entries.
    Entries,
    /// One ZIP entry's bytes.
    EntryBytes,
    /// Aggregate ZIP entry bytes.
    TotalEntryBytes,
    /// Decoded IWA payload bytes.
    PayloadBytes,
    /// Aggregate decoded IWA payload bytes.
    TotalPayloadBytes,
    /// Native payload objects.
    PayloadObjects,
    /// Native payload messages.
    PayloadMessages,
    /// Native payload metadata items.
    PayloadItems,
    /// Native references.
    PayloadReferences,
    /// Codec wire bytes.
    WireBytes,
    /// Codec output bytes.
    WireOutputBytes,
    /// Codec fields.
    WireFields,
    /// Codec nesting.
    WireNesting,
    /// Codec work.
    WireWork,
    /// Aggregate transaction work.
    TransactionWork,
}

impl fmt::Display for BodyTableDimensionLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
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

/// Failure from a body-table dimension read or exact-source transaction.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableDimensionError {
    /// No rooted table matched the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// More than one table matched an exact name.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The rooted selector or graph is ambiguous.
    #[error("the Pages body-table dimension selector is ambiguous")]
    AmbiguousSelector,
    /// The source is not an exact editable package artifact.
    #[error("the Pages package source does not support exact body-table dimension editing")]
    UnsupportedSource,
    /// The selected graph or wire payload is malformed.
    #[error("the selected Pages body-table dimension source is invalid")]
    InvalidSource,
    /// A selected table is protected from editing.
    #[error("the selected Pages body table is locked")]
    TableLocked,
    /// A finite transaction ceiling was exceeded.
    #[error(
        "Pages body-table dimensions {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: BodyTableDimensionLimitKind,
        /// Observed amount.
        observed: u64,
        /// Maximum configured amount.
        maximum: u64,
    },
    /// A fallible bounded allocation failed.
    #[error("could not allocate {amount} units for Pages body-table dimensions")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Candidate reopening did not reproduce the requested semantic state.
    #[error("the edited Pages body-table dimension failed semantic verification")]
    Verification,
    /// A patch was created from a different exact package artifact.
    #[error("the Pages body-table dimension patch does not match the exact source package")]
    PatchConflict,
}

/// Mutable selector-first body-table dimension edit.
pub struct BodyTableDimensionEdit<'a> {
    source: &'a Package,
    target: table_lock::BodyTableTarget,
    dimension: Dimension,
    before: Size,
    size: Size,
}

impl fmt::Debug for BodyTableDimensionEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableDimensionEdit")
            .field("dimension", &self.dimension)
            .field("before", &self.before)
            .field("size", &self.size)
            .finish_non_exhaustive()
    }
}

impl BodyTableDimensionEdit<'_> {
    /// Return the staged size.
    #[must_use]
    pub const fn size(&self) -> Size {
        self.size
    }

    /// Replace the staged size.
    #[must_use]
    pub fn set(mut self, size: Size) -> Self {
        self.size = size;
        self
    }

    /// Stage an explicit point size.
    #[must_use]
    pub fn set_points(self, points: Points) -> Self {
        self.set(Size::Points(points))
    }

    /// Stage the native default size (removing a simple override).
    #[must_use]
    pub fn reset(self) -> Self {
        self.set(Size::Default)
    }

    /// Validate and publish this edit atomically.
    pub fn commit(self) -> Result<BodyTableDimensionCommit, BodyTableDimensionError> {
        commit_edit(self)
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
struct DimensionEvidence {
    model_component: usize,
    model_object: usize,
    model_message: usize,
    model_identifier: NonZeroU64,
    bucket_component: usize,
    bucket_object: usize,
    bucket_message: usize,
    bucket_identifier: NonZeroU64,
}

/// Exact-source reversible dimension patch.
#[derive(Clone, PartialEq)]
pub struct BodyTableDimensionPatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    proof: table_lock::BodyTableTarget,
    evidence: DimensionEvidence,
    dimension: Dimension,
    before: Size,
    after: Size,
    source_preview_count: usize,
    target_preview_count: usize,
}

impl fmt::Debug for BodyTableDimensionPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableDimensionPatch")
            .field("dimension", &self.dimension)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyTableDimensionPatch {
    /// Return the source semantic size.
    #[must_use]
    pub const fn before(&self) -> Size {
        self.before
    }

    /// Return the target semantic size.
    #[must_use]
    pub const fn after(&self) -> Size {
        self.after
    }

    /// Return the source fingerprint used for conflict detection.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the target fingerprint used for conflict detection.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Whether this patch leaves both semantic state and exact bytes unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && (Arc::ptr_eq(&self.source, &self.target) || self.source == self.target)
    }

    /// Return the exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            proof: self.proof.clone(),
            evidence: self.evidence,
            dimension: self.dimension,
            before: self.after,
            after: self.before,
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
        }
    }
}

/// Compact diagnostics from one dimension publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyTableDimensionDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl BodyTableDimensionDiagnostics {
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

    /// Number of deleted root previews.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of a dimension transaction.
#[must_use = "a body-table dimension commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyTableDimensionCommit {
    package: Package,
    patch: BodyTableDimensionPatch,
    diagnostics: BodyTableDimensionDiagnostics,
}

impl BodyTableDimensionCommit {
    /// Borrow the validated package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyTableDimensionPatch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableDimensionDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted body's explicit or default row/column size.
    pub fn body_table_dimension_size<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        dimension: Dimension,
    ) -> Result<Size, BodyTableDimensionError> {
        let target = resolve_target(self, selector.into())?;
        let mut budget = transaction_budget(self)?;
        read_dimension(self, &target, dimension, &mut budget)
    }

    /// Start a selector-first row/column size edit.
    pub fn edit_body_table_dimension_size<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        dimension: Dimension,
    ) -> Result<BodyTableDimensionEdit<'_>, BodyTableDimensionError> {
        let target = resolve_target(self, selector.into())?;
        let mut budget = transaction_budget(self)?;
        let before = read_dimension(self, &target, dimension, &mut budget)?;
        Ok(BodyTableDimensionEdit {
            source: self,
            target,
            dimension,
            before,
            size: before,
        })
    }

    /// Apply a reversible patch to its exact source package.
    pub fn apply_body_table_dimension_size(
        &self,
        patch: &BodyTableDimensionPatch,
    ) -> Result<BodyTableDimensionCommit, BodyTableDimensionError> {
        let source = self.state.source.source_bytes();
        if page_layout::fingerprint(source) != patch.source_fingerprint
            || source != patch.source.as_ref()
        {
            return Err(BodyTableDimensionError::PatchConflict);
        }
        let mut budget = transaction_budget(self)?;
        if read_dimension(self, &patch.proof, patch.dimension, &mut budget)? != patch.before {
            return Err(BodyTableDimensionError::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source_preview_count != patch.target_preview_count {
                return Err(BodyTableDimensionError::PatchConflict);
            }
            return Ok(BodyTableDimensionCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyTableDimensionDiagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact()
            || page_layout::fingerprint(&patch.target) != patch.target_fingerprint
        {
            return Err(BodyTableDimensionError::PatchConflict);
        }
        let candidate = reopen_target(self, Arc::clone(&patch.target), &mut budget)?;
        if read_dimension(&candidate, &patch.proof, patch.dimension, &mut budget)? != patch.after {
            return Err(BodyTableDimensionError::Verification);
        }
        verify_locality(
            self,
            &candidate,
            patch.evidence,
            patch.source_preview_count,
            patch.target_preview_count,
            &mut budget,
        )?;
        Ok(BodyTableDimensionCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableDimensionDiagnostics::published(
                patch
                    .source_preview_count
                    .saturating_sub(patch.target_preview_count),
            ),
        })
    }
}

fn commit_edit(
    edit: BodyTableDimensionEdit<'_>,
) -> Result<BodyTableDimensionCommit, BodyTableDimensionError> {
    let source = edit.source;
    let source_bytes: Arc<[u8]> = source.state.source.shared_source();
    let source_fingerprint = page_layout::fingerprint(source_bytes.as_ref());
    let source_preview_count = preview_count(source);
    let mut budget = transaction_budget(source)?;
    if edit.before == edit.size {
        let evidence = read_evidence(source, &edit.target, edit.dimension, &mut budget)?;
        return Ok(BodyTableDimensionCommit {
            package: source.snapshot(),
            patch: BodyTableDimensionPatch {
                source: Arc::clone(&source_bytes),
                target: source_bytes,
                source_fingerprint,
                target_fingerprint: source_fingerprint,
                proof: edit.target,
                evidence,
                dimension: edit.dimension,
                before: edit.before,
                after: edit.size,
                source_preview_count,
                target_preview_count: source_preview_count,
            },
            diagnostics: BodyTableDimensionDiagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(BodyTableDimensionError::UnsupportedSource);
    }
    let selected = select_dimension(source, &edit.target, edit.dimension, &mut budget)?;
    if selected.size != edit.before {
        return Err(BodyTableDimensionError::InvalidSource);
    }
    if edit.target.explicit_locked == Some(true) {
        return Err(BodyTableDimensionError::TableLocked);
    }
    let previews = preview_names(source);
    let package = rewrite_dimension(source, selected, edit.size, &previews, &mut budget)?;
    let candidate_selected = select_dimension(&package, &edit.target, edit.dimension, &mut budget)?;
    if candidate_selected.size != edit.size {
        return Err(BodyTableDimensionError::Verification);
    }
    verify_locality(
        source,
        &package,
        selected.evidence,
        source_preview_count,
        preview_count(&package),
        &mut budget,
    )?;
    let target = package.state.source.shared_source();
    let target_fingerprint = page_layout::fingerprint(target.as_ref());
    let target_preview_count = preview_count(&package);
    Ok(BodyTableDimensionCommit {
        package,
        patch: BodyTableDimensionPatch {
            source: source_bytes,
            target,
            source_fingerprint,
            target_fingerprint,
            proof: edit.target,
            evidence: selected.evidence,
            dimension: edit.dimension,
            before: edit.before,
            after: edit.size,
            source_preview_count,
            target_preview_count,
        },
        diagnostics: BodyTableDimensionDiagnostics::published(
            source_preview_count.saturating_sub(target_preview_count),
        ),
    })
}

fn resolve_target(
    package: &Package,
    selector: BodyTableSelector<'_>,
) -> Result<table_lock::BodyTableTarget, BodyTableDimensionError> {
    package.resolve_body_table(selector).map_err(map_lock_error)
}

fn transaction_budget(
    package: &Package,
) -> Result<table_lock::WireBudget, BodyTableDimensionError> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    budget
        .charge_source_catalog(&package.state.source)
        .map_err(map_lock_error)?;
    Ok(budget)
}

#[derive(Clone, Copy)]
struct Selected {
    evidence: DimensionEvidence,
    dimension: Dimension,
    size: Size,
    limit: u32,
}

fn select_dimension(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    dimension: Dimension,
    budget: &mut table_lock::WireBudget,
) -> Result<Selected, BodyTableDimensionError> {
    table_lock::validate_body_table_target(package, target, budget).map_err(map_lock_error)?;
    let message = model_message(package, target)?;
    let (model, report) = decode_model(&message.data, budget)?;
    budget
        .charge_codec_report(
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
        )
        .map_err(map_lock_error)?;
    let limit = match dimension {
        Dimension::Row(_) => model.number_of_rows(),
        Dimension::Column(_) => model.number_of_columns(),
    };
    let index =
        u32::try_from(dimension.index()).map_err(|_| BodyTableDimensionError::InvalidSource)?;
    if index >= limit {
        return Err(BodyTableDimensionError::InvalidSource);
    }
    let mut collector = BucketCollector::default();
    let options = codec_options(budget, model.base_data_store().len());
    let (store, store_report) =
        codec::decode_data_store_with_visitor(model.base_data_store(), options, &mut collector)
            .map_err(|error| map_codec_error(error, budget))?;
    budget
        .charge_codec_report(
            store_report.fields(),
            store_report.work_bytes(),
            store_report.max_depth(),
            store_report.references(),
        )
        .map_err(map_lock_error)?;
    let column_reference = store.column_headers();
    let column_id = column_reference.identifier();
    if column_id == 0 || column_reference.deprecated_is_external() == Some(true) {
        return Err(BodyTableDimensionError::InvalidSource);
    }
    let expected_rows =
        usize::try_from(u64::from(model.number_of_rows()).div_ceil(HEADER_BUCKET_ROWS as u64))
            .map_err(|_| BodyTableDimensionError::InvalidSource)?;
    if collector.references.len() != expected_rows {
        return Err(BodyTableDimensionError::InvalidSource);
    }
    let mut row_ids = Vec::new();
    row_ids
        .try_reserve_exact(collector.references.len())
        .map_err(|_| BodyTableDimensionError::Allocation {
            amount: collector.references.len(),
        })?;
    for reference in &collector.references {
        if reference.identifier() == 0 || reference.deprecated_is_external() == Some(true) {
            return Err(BodyTableDimensionError::InvalidSource);
        }
        row_ids.push(reference.identifier());
    }
    row_ids.sort_unstable();
    if row_ids.windows(2).any(|pair| pair[0] == pair[1])
        || row_ids.binary_search(&column_id).is_ok()
    {
        return Err(BodyTableDimensionError::InvalidSource);
    }
    let model_object = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .and_then(|component| component.archive().objects.get(target.model_object_index))
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    validate_declared_references(
        model_object,
        target.model_message_index,
        &row_ids,
        &[4, 1, 2],
    )?;
    validate_declared_references(
        model_object,
        target.model_message_index,
        &[column_id],
        &[4, 2],
    )?;
    let (bucket_id, bucket_slot) = match dimension {
        Dimension::Column(_) => (column_id, None),
        Dimension::Row(row) => {
            let slot = row / HEADER_BUCKET_ROWS;
            (
                *row_ids
                    .get(slot)
                    .ok_or(BodyTableDimensionError::InvalidSource)?,
                Some(slot),
            )
        },
    };
    for identifier in &row_ids {
        let _ = locate_bucket(package, *identifier)?;
    }
    let bucket = locate_bucket(package, bucket_id)?;
    let minimum = bucket_slot.unwrap_or(0).saturating_mul(HEADER_BUCKET_ROWS) as u32;
    let maximum = bucket_slot
        .map(|_| minimum.saturating_add(HEADER_BUCKET_ROWS as u32).min(limit))
        .unwrap_or(limit);
    let payload = bucket_message(package, bucket)?;
    let mut reader = HeaderReader {
        index,
        minimum,
        maximum,
        found: None,
        seen: Vec::new(),
        duplicate: false,
    };
    let (_bucket, header_report) = codec::decode_header_storage_bucket_with_visitor(
        payload,
        codec_options(budget, payload.len()),
        &mut reader,
    )
    .map_err(|error| map_codec_error(error, budget))?;
    budget
        .charge_codec_report(
            header_report.fields(),
            header_report.work_bytes(),
            header_report.max_depth(),
            header_report.references(),
        )
        .map_err(map_lock_error)?;
    reader.seen.sort_unstable();
    if reader.duplicate || reader.seen.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(BodyTableDimensionError::InvalidSource);
    }
    let size = reader
        .found
        .map(size_from_bits)
        .transpose()?
        .unwrap_or(Size::Default);
    Ok(Selected {
        evidence: DimensionEvidence {
            model_component: target.model_component_index,
            model_object: target.model_object_index,
            model_message: target.model_message_index,
            model_identifier: target.model_identifier,
            bucket_component: bucket.0,
            bucket_object: bucket.1,
            bucket_message: bucket.2,
            bucket_identifier: NonZeroU64::new(bucket_id)
                .ok_or(BodyTableDimensionError::InvalidSource)?,
        },
        dimension,
        size,
        limit,
    })
}

fn read_dimension(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    dimension: Dimension,
    budget: &mut table_lock::WireBudget,
) -> Result<Size, BodyTableDimensionError> {
    Ok(select_dimension(package, target, dimension, budget)?.size)
}

fn read_evidence(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    dimension: Dimension,
    budget: &mut table_lock::WireBudget,
) -> Result<DimensionEvidence, BodyTableDimensionError> {
    Ok(select_dimension(package, target, dimension, budget)?.evidence)
}

#[derive(Default)]
struct BucketCollector {
    references: Vec<codec::ReferenceSnapshot>,
}

impl codec::StorageVisitor for BucketCollector {
    fn visit_header_bucket(
        &mut self,
        record: codec::ReferenceRecord<'_>,
    ) -> Result<(), codec::DecodeError> {
        self.references
            .try_reserve(1)
            .map_err(|_| codec::DecodeError::allocation(self.references.len().saturating_add(1)))?;
        self.references.push(record.reference());
        Ok(())
    }
}

struct HeaderReader {
    index: u32,
    minimum: u32,
    maximum: u32,
    found: Option<u32>,
    seen: Vec<u32>,
    duplicate: bool,
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
        self.seen
            .try_reserve(1)
            .map_err(|_| codec::DecodeError::allocation(self.seen.len().saturating_add(1)))?;
        self.seen.push(header.index());
        if header.index() == self.index && self.found.replace(header.size_bits()).is_some() {
            self.duplicate = true;
        }
        Ok(())
    }
}

fn size_from_bits(bits: u32) -> Result<Size, BodyTableDimensionError> {
    if bits == 0 {
        return Ok(Size::Default);
    }
    Points::new(f32::from_bits(bits))
        .map(Size::Points)
        .map_err(|_| BodyTableDimensionError::InvalidSource)
}

fn decode_model<'a>(
    source: &'a [u8],
    budget: &table_lock::WireBudget,
) -> Result<(codec::TableModelSnapshot<'a>, codec::DecodeReport), BodyTableDimensionError> {
    codec::decode_table_model_with_report(source, codec_options(budget, source.len()))
        .map_err(|error| map_codec_error(error, budget))
}

fn codec_options(budget: &table_lock::WireBudget, bytes: usize) -> codec::DecodeOptions {
    let limits = budget.wire_limits();
    codec::DecodeOptions::new(
        bytes.max(1).min(limits.max_input_bytes()),
        budget.remaining_wire_fields(),
        budget.remaining_wire_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        budget.maximum_payload_references(),
        limits.max_input_bytes(),
    )
}

fn model_message<'a>(
    package: &'a Package,
    target: &table_lock::BodyTableTarget,
) -> Result<&'a RawMessage, BodyTableDimensionError> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableDimensionError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    Ok(message)
}

fn locate_bucket(
    package: &Package,
    identifier: u64,
) -> Result<(usize, usize, usize), BodyTableDimensionError> {
    let mut found = None;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            if object.archive_info.identifier != Some(identifier) {
                continue;
            }
            if found.is_some() {
                return Err(BodyTableDimensionError::InvalidSource);
            }
            let mut message_index = None;
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ != HEADER_BUCKET_MESSAGE_TYPE {
                    continue;
                }
                if message_index.replace(index).is_some() {
                    return Err(BodyTableDimensionError::InvalidSource);
                }
            }
            let message_index = message_index.ok_or(BodyTableDimensionError::InvalidSource)?;
            validate_message_metadata(object, message_index)?;
            found = Some((component_index, object_index, message_index));
        }
    }
    found.ok_or(BodyTableDimensionError::InvalidSource)
}

fn bucket_message(
    package: &Package,
    location: (usize, usize, usize),
) -> Result<&[u8], BodyTableDimensionError> {
    package
        .state
        .source
        .components()
        .get_index(location.0)
        .and_then(|component| component.archive().objects.get(location.1))
        .and_then(|object| object.messages.get(location.2))
        .filter(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(BodyTableDimensionError::InvalidSource)
}

fn validate_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), BodyTableDimensionError> {
    let message = object
        .messages
        .get(message_index)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    if message.type_ != info.type_
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(BodyTableDimensionError::InvalidSource);
    }
    Ok(())
}

fn validate_declared_references(
    object: &ArchiveObject,
    message_index: usize,
    identifiers: &[u64],
    path: &[u32],
) -> Result<(), BodyTableDimensionError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    let mut has_declared_path = false;
    for field in &info.field_infos {
        if field.path.as_slice() == path {
            has_declared_path = true;
            for value in &field.object_references {
                if !identifiers.contains(value) {
                    return Err(BodyTableDimensionError::InvalidSource);
                }
            }
        } else if field
            .object_references
            .iter()
            .any(|value| identifiers.contains(value))
        {
            return Err(BodyTableDimensionError::InvalidSource);
        }
    }
    for identifier in identifiers {
        if info
            .object_references
            .iter()
            .filter(|value| **value == *identifier)
            .count()
            != 1
        {
            return Err(BodyTableDimensionError::InvalidSource);
        }
        if has_declared_path
            && info
                .field_infos
                .iter()
                .filter(|field| field.path.as_slice() == path)
                .flat_map(|field| field.object_references.iter())
                .filter(|value| **value == *identifier)
                .count()
                != 1
        {
            return Err(BodyTableDimensionError::InvalidSource);
        }
    }
    Ok(())
}

fn rewrite_dimension(
    source: &Package,
    selected: Selected,
    after: Size,
    previews: &[&'static str],
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableDimensionError> {
    let catalog = &source.state.source;
    let component = catalog
        .components()
        .get_index(selected.evidence.bucket_component)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    let component_name = component.name();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyTableDimensionError::UnsupportedSource);
    }
    budget
        .charge_payload_bytes(entry.data().len())
        .map_err(map_lock_error)?;
    budget
        .charge_total_payload_bytes(entry.data().len())
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(entry.data().len())
        .map_err(map_lock_error)?;
    let payload = bucket_message(
        source,
        (
            selected.evidence.bucket_component,
            selected.evidence.bucket_object,
            selected.evidence.bucket_message,
        ),
    )?;
    let bounds = preflight_dimension_rewrite(source, selected, entry, payload.len(), budget)?;
    let dimension_index = match selected.dimension {
        Dimension::Row(row) | Dimension::Column(row) => {
            u32::try_from(row).map_err(|_| BodyTableDimensionError::InvalidSource)?
        },
    };
    let edit = match after {
        Size::Default => codec::HeaderSizeEdit::remove(dimension_index),
        Size::Points(points) => {
            codec::HeaderSizeEdit::set(dimension_index, points.value().to_bits())
        },
    };
    // The plan retains staged records and index scratch before exposing its
    // exact requirements. Charge a source-sized scan before it can allocate.
    budget
        .charge_payload_work(payload.len())
        .map_err(map_lock_error)?;
    let plan_options = codec_options(budget, payload.len().saturating_add(64));
    let plan =
        codec::plan_header_storage_bucket_sizes(payload, selected.limit, &[edit], plan_options)
            .map_err(|error| map_codec_error(error, budget))?;
    let requirements = plan.requirements();
    charge_report(budget, requirements.source())?;
    charge_bound(budget, requirements.result_upper_bound())?;
    budget
        .charge_payload_work(requirements.rewrite_work_bytes())
        .map_err(map_lock_error)?;
    let execution_options = codec_options(budget, requirements.output_bytes().max(1));
    let (rewritten_payload, report) =
        codec::execute_header_storage_bucket_size_plan(plan, execution_options)
            .map_err(|error| map_codec_error(error, budget))?;
    let source_requirements = requirements.source();
    let result_bound = requirements.result_upper_bound();
    let report_source = report.source();
    let report_result = report.result();
    if report.output_bytes() != requirements.output_bytes()
        || report.rewrite_work_bytes() > requirements.rewrite_work_bytes()
        || report_source.fields() > source_requirements.fields()
        || report_source.work_bytes() > source_requirements.work_bytes()
        || report_source.references() > source_requirements.references()
        || report_result.fields() > result_bound.fields()
        || report_result.work_bytes() > result_bound.work_bytes()
        || report_result.references() > result_bound.references()
    {
        return Err(BodyTableDimensionError::Verification);
    }
    let (mut archive, archive_limits) =
        page_layout::editable_archive(source, component_name).map_err(map_page_layout_error)?;
    let object = archive
        .objects
        .get_mut(selected.evidence.bucket_object)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    if object.archive_info.identifier != Some(selected.evidence.bucket_identifier.get()) {
        return Err(BodyTableDimensionError::InvalidSource);
    }
    validate_message_metadata(object, selected.evidence.bucket_message)?;
    object
        .replace_message_preserving_header_with_limits(
            selected.evidence.bucket_message,
            RawMessage {
                type_: HEADER_BUCKET_MESSAGE_TYPE,
                data: rewritten_payload.clone(),
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed =
        page_layout::compress_archive(archive, archive_limits).map_err(map_page_layout_error)?;
    if compressed.len() > bounds.compressed_bound {
        return Err(BodyTableDimensionError::Verification);
    }
    let edits = [EntryEdit::new(component_name, &compressed)];
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, previews, catalog.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    if requirements.output_bytes() > bounds.package_output_bound
        || requirements.retained_bytes() > bounds.package_output_bound
    {
        return Err(BodyTableDimensionError::Verification);
    }
    // The broad output bound was charged before any codec/archive mutation.
    // Reassembly's exact scratch and allocation counts are additional work
    // owned by this transaction; execute only after those requirements pass.
    budget
        .charge_payload_work(requirements.scratch_bytes())
        .and_then(|_| budget.charge_payload_work(requirements.allocations()))
        .map_err(map_lock_error)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    if output.len() != requirements.output_bytes() {
        return Err(BodyTableDimensionError::Verification);
    }
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), catalog.limits())
            .map_err(map_archive_error)?;
    Package::from_source_catalog(candidate_source).map_err(map_package_error)
}

#[derive(Clone, Copy)]
struct DimensionRewriteBounds {
    compressed_bound: usize,
    package_output_bound: usize,
}

fn preflight_dimension_rewrite(
    source: &Package,
    selected: Selected,
    entry: &litchi_iwa_archive::package::Entry,
    original_message_length: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<DimensionRewriteBounds, BodyTableDimensionError> {
    let catalog = &source.state.source;
    let component = catalog
        .components()
        .get_index(selected.evidence.bucket_component)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    let stream_length = parsed_archive_source_length(component.archive())?;
    let message_count = component
        .archive()
        .objects
        .iter()
        .try_fold(0usize, |count, object| {
            count.checked_add(object.messages.len())
        })
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    let rewritten_message_bound = original_message_length
        .checked_add(MAX_VARINT_BYTES.saturating_mul(4))
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    let archive_bound = stream_length
        .checked_sub(original_message_length)
        .and_then(|value| value.checked_add(rewritten_message_bound))
        .and_then(|value| {
            value.checked_add(
                message_count
                    .checked_mul(MAX_VARINT_BYTES.saturating_mul(3))?
                    .checked_add(MAX_VARINT_BYTES.saturating_mul(2))?,
            )
        })
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    let compressed_bound = table_lock::snappy_compressed_bound(archive_bound)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    let replacement_compressed_bound = match entry.metadata().central().compression_method() {
        0 => compressed_bound,
        8 => table_lock::deflate_compressed_bound(compressed_bound)
            .ok_or(BodyTableDimensionError::InvalidSource)?,
        _ => return Err(BodyTableDimensionError::UnsupportedSource),
    };
    let old_compressed_size =
        usize::try_from(entry.metadata().compressed_size()).map_err(|_| {
            BodyTableDimensionError::LimitExceeded {
                kind: BodyTableDimensionLimitKind::EntryBytes,
                observed: u64::MAX,
                maximum: catalog.limits().max_entry_bytes(),
            }
        })?;
    let package_output_bound = catalog
        .source_bytes()
        .len()
        .checked_sub(old_compressed_size)
        .and_then(|value| value.checked_add(replacement_compressed_bound))
        .ok_or(BodyTableDimensionError::InvalidSource)?;
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
        .map_err(map_lock_error)?;
    budget
        .precharge_candidate_reopen(
            catalog,
            package_output_bound,
            selected.evidence.bucket_component,
            compressed_bound,
            archive_bound,
            selected.evidence.bucket_object,
            selected.evidence.bucket_message,
            rewritten_message_bound,
        )
        .map_err(map_lock_error)?;
    Ok(DimensionRewriteBounds {
        compressed_bound,
        package_output_bound,
    })
}

fn parsed_archive_source_length(
    archive: &litchi_iwa_core::Archive,
) -> Result<usize, BodyTableDimensionError> {
    let Some(last) = archive.objects.last() else {
        return Ok(0);
    };
    let length = last
        .data_offset
        .checked_add(last.data_length)
        .ok_or(BodyTableDimensionError::InvalidSource)?;
    usize::try_from(length).map_err(|_| BodyTableDimensionError::InvalidSource)
}

fn charge_report(
    budget: &mut table_lock::WireBudget,
    report: codec::DecodeReport,
) -> Result<(), BodyTableDimensionError> {
    budget
        .charge_codec_report(
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
        )
        .map_err(map_lock_error)
}

fn charge_bound(
    budget: &mut table_lock::WireBudget,
    bound: codec::DecodeResourceUpperBound,
) -> Result<(), BodyTableDimensionError> {
    budget
        .charge_codec_report(
            bound.fields(),
            bound.work_bytes(),
            bound.max_depth(),
            bound.references(),
        )
        .map_err(map_lock_error)
}

fn preview_names(package: &Package) -> Vec<&'static str> {
    ROOT_PREVIEW_NAMES
        .iter()
        .copied()
        .filter(|name| {
            package
                .state
                .source
                .package()
                .iter()
                .any(|entry| entry.name() == *name)
        })
        .collect()
}

fn preview_count(package: &Package) -> usize {
    preview_names(package).len()
}

fn reopen_target(
    source: &Package,
    target: Arc<[u8]>,
    budget: &mut table_lock::WireBudget,
) -> Result<Package, BodyTableDimensionError> {
    budget
        .charge_input_source(target.as_ref())
        .map_err(map_lock_error)?;
    let catalog =
        SourceCatalog::from_shared_bytes_with_limits(target, source.state.source.limits())
            .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&catalog)
        .map_err(map_lock_error)?;
    table_lock::charge_reopen_work(&catalog, budget).map_err(map_lock_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    evidence: DimensionEvidence,
    source_previews: usize,
    target_previews: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDimensionError> {
    if preview_count(source) != source_previews || preview_count(candidate) != target_previews {
        return Err(BodyTableDimensionError::Verification);
    }
    let before = &source.state.source;
    let after_source = &candidate.state.source;
    let mut before_entries = before
        .package()
        .iter()
        .filter(|entry| !ROOT_PREVIEW_NAMES.contains(&entry.name()));
    let mut after_entries = after_source
        .package()
        .iter()
        .filter(|entry| !ROOT_PREVIEW_NAMES.contains(&entry.name()));
    loop {
        match (before_entries.next(), after_entries.next()) {
            (Some(left), Some(right)) if left.name() == right.name() => {
                if left.name()
                    != before
                        .components()
                        .get_index(evidence.bucket_component)
                        .ok_or(BodyTableDimensionError::Verification)?
                        .name()
                    && {
                        budget
                            .charge_payload_work(
                                left.data().len().saturating_add(right.data().len()),
                            )
                            .map_err(map_lock_error)?;
                        left.data() != right.data()
                    }
                {
                    return Err(BodyTableDimensionError::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(BodyTableDimensionError::Verification),
        }
    }
    if before.components().len() != after_source.components().len() {
        return Err(BodyTableDimensionError::Verification);
    }
    for (component_index, (left, right)) in before
        .components()
        .iter()
        .zip(after_source.components().iter())
        .enumerate()
    {
        if left.name() != right.name()
            || left.archive().objects.len() != right.archive().objects.len()
        {
            return Err(BodyTableDimensionError::Verification);
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
                budget
                    .charge_payload_work(
                        left_object
                            .data_length
                            .try_into()
                            .unwrap_or(usize::MAX)
                            .saturating_add(
                                right_object.data_length.try_into().unwrap_or(usize::MAX),
                            ),
                    )
                    .map_err(map_lock_error)?;
                if !left_object.same_content_ignoring_offsets(right_object) {
                    return Err(BodyTableDimensionError::Verification);
                }
                continue;
            }
            if left_object.archive_info.identifier != right_object.archive_info.identifier
                || left_object.archive_info.should_merge != right_object.archive_info.should_merge
                || left_object.messages.len() != right_object.messages.len()
                || left_object.archive_info.message_infos.len()
                    != right_object.archive_info.message_infos.len()
            {
                return Err(BodyTableDimensionError::Verification);
            }
            for (message_index, (left_message, right_message)) in left_object
                .messages
                .iter()
                .zip(&right_object.messages)
                .enumerate()
            {
                let left_info = left_object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(BodyTableDimensionError::Verification)?;
                let right_info = right_object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(BodyTableDimensionError::Verification)?;
                if message_index == evidence.bucket_message {
                    if right_message.type_ != HEADER_BUCKET_MESSAGE_TYPE
                        || !message_info_preserved_except_length(left_info, right_info)
                    {
                        return Err(BodyTableDimensionError::Verification);
                    }
                    budget
                        .charge_payload_work(
                            left_message
                                .data
                                .len()
                                .saturating_add(right_message.data.len()),
                        )
                        .map_err(map_lock_error)?;
                    if right_message.data == left_message.data {
                        return Err(BodyTableDimensionError::Verification);
                    }
                } else {
                    budget
                        .charge_payload_work(
                            left_message
                                .data
                                .len()
                                .saturating_add(right_message.data.len()),
                        )
                        .map_err(map_lock_error)?;
                    if left_message != right_message || left_info != right_info {
                        return Err(BodyTableDimensionError::Verification);
                    }
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

fn map_codec_error(
    error: codec::DecodeError,
    _budget: &table_lock::WireBudget,
) -> BodyTableDimensionError {
    let Some(limit) = error.resource_limit() else {
        return BodyTableDimensionError::InvalidSource;
    };
    match limit {
        codec::DecodeLimit::Bytes { observed, maximum } => {
            limit_error(BodyTableDimensionLimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::References { observed, maximum } => limit_error(
            BodyTableDimensionLimitKind::PayloadReferences,
            observed,
            maximum,
        ),
        codec::DecodeLimit::Text { observed, maximum } => {
            limit_error(BodyTableDimensionLimitKind::PayloadBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            limit_error(BodyTableDimensionLimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::Work { observed, maximum } => {
            limit_error(BodyTableDimensionLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => limit_error(
            BodyTableDimensionLimitKind::WireNesting,
            observed as usize,
            maximum as usize,
        ),
        codec::DecodeLimit::Allocation { requested } => {
            BodyTableDimensionError::Allocation { amount: requested }
        },
        _ => BodyTableDimensionError::InvalidSource,
    }
}

fn limit_error(
    kind: BodyTableDimensionLimitKind,
    observed: usize,
    maximum: usize,
) -> BodyTableDimensionError {
    BodyTableDimensionError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableDimensionError {
    match error {
        table_lock::BodyTableLockError::TableNotFound => BodyTableDimensionError::TableNotFound,
        table_lock::BodyTableLockError::AmbiguousTableName => {
            BodyTableDimensionError::AmbiguousTableName
        },
        table_lock::BodyTableLockError::AmbiguousSelector => {
            BodyTableDimensionError::AmbiguousSelector
        },
        table_lock::BodyTableLockError::UnsupportedSource => {
            BodyTableDimensionError::UnsupportedSource
        },
        table_lock::BodyTableLockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableDimensionError::LimitExceeded {
            kind: match kind {
                table_lock::BodyTableLockLimitKind::InputBytes => {
                    BodyTableDimensionLimitKind::InputBytes
                },
                table_lock::BodyTableLockLimitKind::OutputBytes => {
                    BodyTableDimensionLimitKind::OutputBytes
                },
                table_lock::BodyTableLockLimitKind::Entries => BodyTableDimensionLimitKind::Entries,
                table_lock::BodyTableLockLimitKind::EntryBytes => {
                    BodyTableDimensionLimitKind::EntryBytes
                },
                table_lock::BodyTableLockLimitKind::TotalEntryBytes => {
                    BodyTableDimensionLimitKind::TotalEntryBytes
                },
                table_lock::BodyTableLockLimitKind::PayloadBytes => {
                    BodyTableDimensionLimitKind::PayloadBytes
                },
                table_lock::BodyTableLockLimitKind::TotalPayloadBytes => {
                    BodyTableDimensionLimitKind::TotalPayloadBytes
                },
                table_lock::BodyTableLockLimitKind::PayloadObjects => {
                    BodyTableDimensionLimitKind::PayloadObjects
                },
                table_lock::BodyTableLockLimitKind::PayloadMessages => {
                    BodyTableDimensionLimitKind::PayloadMessages
                },
                table_lock::BodyTableLockLimitKind::PayloadItems => {
                    BodyTableDimensionLimitKind::PayloadItems
                },
                table_lock::BodyTableLockLimitKind::PayloadReferences => {
                    BodyTableDimensionLimitKind::PayloadReferences
                },
                table_lock::BodyTableLockLimitKind::WireBytes => {
                    BodyTableDimensionLimitKind::WireBytes
                },
                table_lock::BodyTableLockLimitKind::WireFields => {
                    BodyTableDimensionLimitKind::WireFields
                },
                table_lock::BodyTableLockLimitKind::WireNesting => {
                    BodyTableDimensionLimitKind::WireNesting
                },
                table_lock::BodyTableLockLimitKind::WireWork => {
                    BodyTableDimensionLimitKind::WireWork
                },
                table_lock::BodyTableLockLimitKind::PackageBytes => {
                    BodyTableDimensionLimitKind::PayloadItems
                },
            },
            observed,
            maximum,
        },
        table_lock::BodyTableLockError::Allocation { amount } => {
            BodyTableDimensionError::Allocation { amount }
        },
        table_lock::BodyTableLockError::InvalidSource => BodyTableDimensionError::InvalidSource,
        table_lock::BodyTableLockError::Verification => BodyTableDimensionError::Verification,
        table_lock::BodyTableLockError::PatchConflict => BodyTableDimensionError::PatchConflict,
    }
}

fn map_page_layout_error(error: page_layout::PageLayoutError) -> BodyTableDimensionError {
    match error {
        page_layout::PageLayoutError::Allocation { amount } => {
            BodyTableDimensionError::Allocation { amount }
        },
        page_layout::PageLayoutError::LimitExceeded {
            observed, maximum, ..
        } => BodyTableDimensionError::LimitExceeded {
            kind: BodyTableDimensionLimitKind::PayloadBytes,
            observed,
            maximum,
        },
        page_layout::PageLayoutError::UnsupportedSource => {
            BodyTableDimensionError::UnsupportedSource
        },
        _ => BodyTableDimensionError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyTableDimensionError {
    match error {
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyTableDimensionError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableDimensionError::LimitExceeded {
            kind: match kind {
                ArchiveLimitKind::InputBytes => BodyTableDimensionLimitKind::InputBytes,
                ArchiveLimitKind::OutputBytes => BodyTableDimensionLimitKind::OutputBytes,
                ArchiveLimitKind::Entries => BodyTableDimensionLimitKind::Entries,
                ArchiveLimitKind::MemberNameBytes | ArchiveLimitKind::MetadataBytes => {
                    BodyTableDimensionLimitKind::PayloadItems
                },
                ArchiveLimitKind::CompressedEntryBytes | ArchiveLimitKind::EntryBytes => {
                    BodyTableDimensionLimitKind::EntryBytes
                },
                ArchiveLimitKind::TotalBytes => BodyTableDimensionLimitKind::TotalEntryBytes,
                ArchiveLimitKind::IwaStreamBytes => BodyTableDimensionLimitKind::PayloadBytes,
                ArchiveLimitKind::IwaTotalBytes => BodyTableDimensionLimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
        },
        _ => BodyTableDimensionError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyTableDimensionError {
    match error {
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyTableDimensionError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableDimensionError::LimitExceeded {
            kind: match kind {
                CoreLimitKind::ArchiveBytes
                | CoreLimitKind::ObjectBytes
                | CoreLimitKind::MessageBytes
                | CoreLimitKind::SnappyChunkBytes
                | CoreLimitKind::SnappyStreamBytes => BodyTableDimensionLimitKind::PayloadBytes,
                CoreLimitKind::Objects => BodyTableDimensionLimitKind::PayloadObjects,
                CoreLimitKind::Messages | CoreLimitKind::MessagesPerObject => {
                    BodyTableDimensionLimitKind::PayloadMessages
                },
                CoreLimitKind::HeaderBytes
                | CoreLimitKind::HeaderMemoryBytes
                | CoreLimitKind::SnappyCompressedChunkBytes
                | CoreLimitKind::SnappyCompressedStreamBytes => {
                    BodyTableDimensionLimitKind::WireBytes
                },
                CoreLimitKind::HeaderFields => BodyTableDimensionLimitKind::WireFields,
                CoreLimitKind::HeaderNesting => BodyTableDimensionLimitKind::WireNesting,
                CoreLimitKind::MetadataItems | CoreLimitKind::SnappyFrames => {
                    BodyTableDimensionLimitKind::PayloadItems
                },
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => BodyTableDimensionError::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> BodyTableDimensionError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => BodyTableDimensionError::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => BodyTableDimensionError::LimitExceeded {
            kind: BodyTableDimensionLimitKind::PayloadObjects,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::PayloadLimit { observed, limit } => BodyTableDimensionError::LimitExceeded {
            kind: BodyTableDimensionLimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: limit as u64,
        },
        _ => BodyTableDimensionError::InvalidSource,
    }
}
