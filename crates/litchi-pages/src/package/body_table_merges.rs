//! Bounded, selector-first Pages body-table merge readback.
//!
//! The Pages package owns rooted attachment and table-model selection. The
//! selected model payload is then handed to the shared borrowed merge codec;
//! no native identifier or generated protobuf value crosses this module's
//! public boundary.

use litchi_iwa_archive::ComponentCatalog;
use litchi_iwa_core::RawMessage;
use litchi_numbers_wire::table_merges::{self, ReadLimits};
use std::fmt;
use thiserror::Error;

use super::{Package, page_layout, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::merge::Region;

mod reader;
mod transaction;

pub use reader::MergeReader;
pub use transaction::{
    BodyTableMergesCommit, BodyTableMergesDiagnostics, BodyTableMergesEdit, BodyTableMergesPatch,
};

/// A finite resource governed by one body-table merge read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableMergesLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete candidate package output bytes (unused by this read).
    OutputBytes,
    /// ZIP entries inspected while resolving the rooted table.
    Entries,
    /// Bytes retained by one ZIP entry.
    EntryBytes,
    /// Aggregate ZIP entry bytes.
    TotalEntryBytes,
    /// ZIP metadata bytes.
    PackageBytes,
    /// Bytes in one decoded native payload.
    PayloadBytes,
    /// Aggregate decoded native payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected.
    PayloadObjects,
    /// Native payload messages inspected.
    PayloadMessages,
    /// Native payload metadata items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Strict merge-wire input bytes.
    WireBytes,
    /// Strict merge-wire output bytes (unused by this read).
    WireOutputBytes,
    /// Strict merge-wire fields.
    WireFields,
    /// Strict merge-wire nesting.
    WireNesting,
    /// Aggregate merge-wire work.
    WireWork,
}

impl fmt::Display for BodyTableMergesLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "ZIP entries",
            Self::EntryBytes => "ZIP entry bytes",
            Self::TotalEntryBytes => "total ZIP entry bytes",
            Self::PackageBytes => "package metadata bytes",
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
        })
    }
}

/// Failure while reading one Pages body table's merged-cell geometry.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableMergesError {
    /// No rooted body table matched the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// More than one rooted body table matched an exact name.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The rooted selector or ownership graph is ambiguous.
    #[error("the Pages body-table merge selector is ambiguous")]
    AmbiguousSelector,
    /// The source cannot be inspected through the exact native profile.
    #[error("the Pages package source does not support body-table merge reads")]
    UnsupportedSource,
    /// The selected rooted graph or merge formula is malformed.
    #[error("the selected Pages body-table merge source is invalid")]
    InvalidSource,
    /// A requested region is outside the selected table dimensions.
    #[error("the requested Pages body-table merge region is outside the table bounds")]
    InvalidRegion,
    /// A requested region overlaps an already staged merged-cell rectangle.
    #[error("the requested Pages body-table merge region overlaps an existing merge")]
    OverlappingRegion,
    /// A finite read resource ceiling was exceeded.
    #[error(
        "Pages body-table merges {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: BodyTableMergesLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded semantic result allocation failed.
    #[error("could not allocate {amount} units for Pages body-table merges")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Complete candidate reopening did not reproduce the requested regions.
    #[error("the edited Pages body-table merges failed semantic verification")]
    Verification,
    /// The supplied patch was not created from this exact package artifact.
    #[error("the Pages body-table merge patch does not match the exact source package")]
    PatchConflict,
}

fn select_reader_target(
    targets: Vec<table_lock::BodyTableTarget>,
    selector: BodyTableSelector<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<table_lock::BodyTableTarget, table_lock::BodyTableLockError> {
    match selector {
        BodyTableSelector::Position(position) => {
            budget.charge_payload_work(position.get().saturating_add(1).min(targets.len()))?;
            targets
                .into_iter()
                .nth(position.get())
                .ok_or(table_lock::BodyTableLockError::TableNotFound)
        },
        BodyTableSelector::Name(name) => {
            let mut first = None;
            for target in targets {
                budget.charge_payload_work(target.table_name.len().saturating_add(name.len()))?;
                if target.table_name.as_ref() != name {
                    continue;
                }
                if first.replace(target).is_some() {
                    return Err(table_lock::BodyTableLockError::AmbiguousTableName);
                }
            }
            first.ok_or(table_lock::BodyTableLockError::TableNotFound)
        },
    }
}

impl Package {
    /// Read the validated merged-cell rectangles of one rooted body table.
    ///
    /// The selector is resolved through the same ownership proof used by the
    /// other Pages table readers. Root selection, model-wire validation, and
    /// merge-formula decoding share one finite operation budget; the catalog
    /// is not read and then scanned again. An absent merge owner is a valid
    /// empty result.
    ///
    /// # Errors
    ///
    /// Returns [`BodyTableMergesError::InvalidSource`] when a merge formula is
    /// malformed, references another table, overlaps another region, or
    /// extends outside the selected table dimensions.
    pub fn body_table_merges<'table, S>(
        &self,
        selector: S,
    ) -> Result<Vec<Region>, BodyTableMergesError>
    where
        S: Into<BodyTableSelector<'table>>,
    {
        let mut budget =
            table_lock::WireBudget::new(self.state.source.limits()).map_err(map_lock_error)?;
        let target = self
            .resolve_body_table_with_budget(selector.into(), &mut budget)
            .map_err(map_lock_error)?;
        read_regions(self, &target, &mut budget)
    }
}

fn read_regions(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<Region>, BodyTableMergesError> {
    table_lock::validate_body_table_target(package, target, budget).map_err(map_lock_error)?;
    let message = model_message(package, target)?;
    let limits = read_limits(budget)?;
    let read = table_merges::read_table_merges(&message.data, limits).map_err(map_merge_error)?;

    // The shared reader reports the selected message and formula passes
    // separately. Charge each pass once in the enclosing Pages budget; the
    // returned regions were bounded by max_regions before its result vector
    // was reserved.
    budget
        .charge_payload_work(read.report.input_bytes())
        .map_err(map_lock_error)?;
    budget
        .charge_codec_report(read.report.fields(), read.report.work(), 0, 0)
        .map_err(map_lock_error)?;

    for region in &read.regions {
        if region.end_row() >= target.table_rows || region.end_column() >= target.table_columns {
            return Err(BodyTableMergesError::InvalidSource);
        }
    }
    Ok(read.regions)
}

fn model_message<'source>(
    package: &'source Package,
    target: &table_lock::BodyTableTarget,
) -> Result<&'source RawMessage, BodyTableMergesError> {
    model_message_from_components(package.state.source.components(), target)
}

fn model_message_from_components<'source>(
    components: &'source ComponentCatalog,
    target: &table_lock::BodyTableTarget,
) -> Result<&'source RawMessage, BodyTableMergesError> {
    let component = components
        .get_index(target.model_component_index)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableMergesError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(BodyTableMergesError::InvalidSource)?;
    Ok(message)
}

fn read_limits(budget: &table_lock::WireBudget) -> Result<ReadLimits, BodyTableMergesError> {
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
    // A merge read walks the model, owner, formula store, pair, and formula
    // payloads. Their input report is cumulative, so the selected model's
    // length is only one part of the remaining budget. Use the remaining
    // aggregate work as a conservative byte ceiling, bounded by the common
    // wire profile; the enclosing budget charges the report after success.
    let input = work.min(base.max_input_bytes());
    let wire = base
        .with_input_bytes(input)
        .and_then(|limits| limits.with_fields(fields))
        .and_then(|limits| limits.with_rewrite_work(work))
        .map_err(map_common_error)?;
    Ok(ReadLimits {
        wire,
        max_regions: fields,
        max_overlap_checks: work,
    })
}

fn map_merge_error(error: table_merges::MergeReadError) -> BodyTableMergesError {
    map_common_error(error.error().clone())
}

fn map_package_error(error: super::PackageError) -> BodyTableMergesError {
    map_lock_error(table_lock::map_package_error(error))
}

fn map_common_error(error: litchi_iwa_common::Error) -> BodyTableMergesError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => BodyTableMergesError::LimitExceeded {
            kind: map_common_limit(kind),
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            BodyTableMergesError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => BodyTableMergesError::InvalidSource,
    }
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableMergesError {
    match error {
        table_lock::BodyTableLockError::TableNotFound => BodyTableMergesError::TableNotFound,
        table_lock::BodyTableLockError::AmbiguousTableName => {
            BodyTableMergesError::AmbiguousTableName
        },
        table_lock::BodyTableLockError::AmbiguousSelector => {
            BodyTableMergesError::AmbiguousSelector
        },
        table_lock::BodyTableLockError::UnsupportedSource => {
            BodyTableMergesError::UnsupportedSource
        },
        table_lock::BodyTableLockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableMergesError::LimitExceeded {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::BodyTableLockError::Allocation { amount } => {
            BodyTableMergesError::Allocation { amount }
        },
        table_lock::BodyTableLockError::InvalidSource
        | table_lock::BodyTableLockError::Verification
        | table_lock::BodyTableLockError::PatchConflict => BodyTableMergesError::InvalidSource,
    }
}

const fn map_common_limit(kind: litchi_iwa_common::LimitKind) -> BodyTableMergesLimitKind {
    use litchi_iwa_common::LimitKind as Common;
    match kind {
        Common::InputBytes => BodyTableMergesLimitKind::WireBytes,
        Common::Fields => BodyTableMergesLimitKind::WireFields,
        Common::OutputBytes => BodyTableMergesLimitKind::WireOutputBytes,
        Common::Nesting => BodyTableMergesLimitKind::WireNesting,
        Common::RewriteWork => BodyTableMergesLimitKind::WireWork,
        Common::TableRows
        | Common::TableColumns
        | Common::TableCells
        | Common::MaterializedCells => BodyTableMergesLimitKind::PayloadItems,
    }
}

const fn map_lock_limit(kind: table_lock::BodyTableLockLimitKind) -> BodyTableMergesLimitKind {
    use table_lock::BodyTableLockLimitKind as Lock;
    match kind {
        Lock::InputBytes => BodyTableMergesLimitKind::InputBytes,
        Lock::OutputBytes => BodyTableMergesLimitKind::OutputBytes,
        Lock::Entries => BodyTableMergesLimitKind::Entries,
        Lock::EntryBytes => BodyTableMergesLimitKind::EntryBytes,
        Lock::TotalEntryBytes => BodyTableMergesLimitKind::TotalEntryBytes,
        Lock::PackageBytes => BodyTableMergesLimitKind::PackageBytes,
        Lock::PayloadBytes => BodyTableMergesLimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => BodyTableMergesLimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => BodyTableMergesLimitKind::PayloadObjects,
        Lock::PayloadMessages => BodyTableMergesLimitKind::PayloadMessages,
        Lock::PayloadItems => BodyTableMergesLimitKind::PayloadItems,
        Lock::PayloadReferences => BodyTableMergesLimitKind::PayloadReferences,
        Lock::WireBytes => BodyTableMergesLimitKind::WireBytes,
        Lock::WireFields => BodyTableMergesLimitKind::WireFields,
        Lock::WireNesting => BodyTableMergesLimitKind::WireNesting,
        Lock::WireWork => BodyTableMergesLimitKind::WireWork,
    }
}
