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
use thiserror::Error as ThisError;

use super::{Package, page_layout, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::merge::Region;

mod reader;
mod transaction;

pub use reader::MergeReader;
pub use transaction::{Commit, Diagnostics, Edit, Patch};

/// A finite resource governed by one body-table merge read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
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
    /// Resolved formula-owner entries inspected by a merge formula.
    OwnerEntries,
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

impl fmt::Display for LimitKind {
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
            Self::OwnerEntries => "formula owner entries",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// Failure while reading one Pages body table's merged-cell geometry.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The selected rooted graph or merge formula is malformed.
    #[error("the selected Pages body-table merge source is invalid")]
    InvalidSource,
    /// A finite read resource ceiling was exceeded.
    #[error(
        "Pages body-table merges {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
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
    /// The rooted selector or ownership graph is ambiguous.
    #[error("the Pages body-table merge selector is ambiguous")]
    AmbiguousSelector,
    /// The source cannot be inspected through the exact native profile.
    #[error("the Pages package source does not support body-table merge reads")]
    UnsupportedSource,
    /// A requested region is outside the selected table dimensions.
    #[error("the requested Pages body-table merge region is outside the table bounds")]
    InvalidRegion,
    /// A requested region overlaps an already staged merged-cell rectangle.
    #[error("the requested Pages body-table merge region overlaps an existing merge")]
    OverlappingRegion,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::AmbiguousTableName { .. }
            | Self::AmbiguousSelector { .. }
            | Self::InvalidRegion { .. }
            | Self::OverlappingRegion { .. } => RefusalKind::InvalidValue,
            Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

fn select_reader_target(
    targets: Vec<table_lock::BodyTableTarget>,
    selector: BodyTableSelector<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<table_lock::BodyTableTarget, table_lock::Error> {
    match selector {
        BodyTableSelector::Position(position) => {
            budget.charge_payload_work(position.get().saturating_add(1).min(targets.len()))?;
            targets
                .into_iter()
                .nth(position.get())
                .ok_or(table_lock::Error::Refused(
                    crate::table::lock::Refusal::TableNotFound,
                ))
        },
        BodyTableSelector::Name(name) => {
            let mut first = None;
            for target in targets {
                budget.charge_payload_work(target.table_name.len().saturating_add(name.len()))?;
                if target.table_name.as_ref() != name {
                    continue;
                }
                if first.replace(target).is_some() {
                    return Err(table_lock::Error::Refused(
                        crate::table::lock::Refusal::AmbiguousTableName,
                    ));
                }
            }
            first.ok_or(table_lock::Error::Refused(
                crate::table::lock::Refusal::TableNotFound,
            ))
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
    /// Returns [`Error::InvalidSource`] when a merge formula is
    /// malformed, references another table, overlaps another region, or
    /// extends outside the selected table dimensions.
    pub fn body_table_merges<'table, S>(&self, selector: S) -> Result<Vec<Region>, Error>
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
) -> Result<Vec<Region>, Error> {
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
            return Err(Error::InvalidSource);
        }
    }
    Ok(read.regions)
}

fn model_message<'source>(
    package: &'source Package,
    target: &table_lock::BodyTableTarget,
) -> Result<&'source RawMessage, Error> {
    model_message_from_components(package.state.source.components(), target)
}

fn model_message_from_components<'source>(
    components: &'source ComponentCatalog,
    target: &table_lock::BodyTableTarget,
) -> Result<&'source RawMessage, Error> {
    let component = components
        .get_index(target.model_component_index)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(Error::InvalidSource)?;
    object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(Error::InvalidSource)?;
    Ok(message)
}

fn read_limits(budget: &table_lock::WireBudget) -> Result<ReadLimits, Error> {
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

fn map_merge_error(error: table_merges::MergeReadError) -> Error {
    if let Some(owner_limit) = error.owner_entries_limit() {
        return Error::Limit {
            kind: LimitKind::OwnerEntries,
            observed: owner_limit.observed() as u64,
            maximum: owner_limit.maximum() as u64,
        };
    }
    map_common_error(error.error().clone())
}

fn map_package_error(error: super::PackageError) -> Error {
    map_lock_error(table_lock::map_package_error(error))
}

fn map_common_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: map_common_limit(kind),
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
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
        table_lock::Error::InvalidSource
        | table_lock::Error::Verification
        | table_lock::Error::PatchConflict => Error::InvalidSource,
    }
}

const fn map_common_limit(kind: litchi_iwa_common::LimitKind) -> LimitKind {
    use litchi_iwa_common::LimitKind as Common;
    match kind {
        Common::InputBytes => LimitKind::WireBytes,
        Common::Fields => LimitKind::WireFields,
        Common::OutputBytes => LimitKind::WireOutputBytes,
        Common::Nesting => LimitKind::WireNesting,
        Common::RewriteWork => LimitKind::WireWork,
        Common::TableRows
        | Common::TableColumns
        | Common::TableCells
        | Common::MaterializedCells => LimitKind::PayloadItems,
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
