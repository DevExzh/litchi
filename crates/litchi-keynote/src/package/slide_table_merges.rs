//! Selector-first readback of native merged-cell regions in Keynote tables.
//!
//! Table selection and package authority remain in [`super::slide_table_core`]
//! while the selected model's merge-owner wire path is decoded by the shared
//! Numbers wire adapter.  The operation keeps one resource ledger from root
//! selection through the borrowed merge scan; it never mutates or reassembles
//! the source package.

mod reader;
mod transaction;

pub use reader::MergeReader;
pub use transaction::{Commit, Diagnostics, Edit, Patch};

use std::fmt;

use litchi_core::Position;
use litchi_iwa_common::{LimitKind as WireLimitKind, WireLimits, table::merge::Region};
use litchi_numbers_wire::table_merges::{
    self, FORMULA_INDEX_BYTES, MergeReadError, PAIR_REFERENCE_BYTES, REGION_BYTES,
};
use thiserror::Error as ThisError;

use super::Package;
use super::slide_table_core as core;
use crate::SlideSelector;
use crate::slide::table::TableSelector;

const MERGE_READER_ALLOCATIONS_PER_REGION: usize = 3;
const MERGE_READER_SCRATCH_PER_REGION: usize = PAIR_REFERENCE_BYTES + FORMULA_INDEX_BYTES;

/// Finite resources charged by a focused Keynote merged-cell read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Source package bytes inspected while resolving the selected table.
    InputBytes,
    /// Candidate package bytes produced by a lower-level read or transaction.
    OutputBytes,
    /// Package entries visited by the operation.
    Entries,
    /// Bytes retained for one package entry.
    EntryBytes,
    /// Total package-entry bytes retained by the operation.
    TotalBytes,
    /// Native archive objects visited by the root proof.
    PayloadObjects,
    /// Native archive messages visited by the root proof.
    PayloadMessages,
    /// Native merge-pair items staged by the selected decoder.
    PayloadItems,
    /// Native references inspected by the root proof.
    References,
    /// Resolved formula-owner entries inspected by a merge formula.
    OwnerEntries,
    /// Wire fields inspected by the selected merge path.
    WireFields,
    /// Maximum nested wire depth inspected by the selected merge path.
    WireNesting,
    /// Aggregate wire work performed by the operation.
    WireWork,
    /// Fallible allocations admitted by the operation.
    Allocations,
    /// Bytes retained by decoded semantic values.
    Retained,
    /// Temporary bytes retained while decoding.
    Scratch,
    /// Components inspected by the root proof.
    Components,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::References => "references",
            Self::OwnerEntries => "formula owner entries",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
        })
    }
}

/// Content-free semantic location for a focused merged-cell read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete package.
    Package,
    /// One checked slide/table selection.
    Table { slide: Position, table: Position },
}

impl fmt::Display for Path {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Table { slide, table } => {
                write!(
                    formatter,
                    "slide {} table {} merged cells",
                    slide.get(),
                    table.get()
                )
            },
        }
    }
}

/// Failure from a focused Keynote merged-cell read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    #[error("the selected Keynote slide-table merge source is invalid")]
    InvalidSource,
    #[error(
        "Keynote slide-table merges {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote slide-table merge operation")]
    Allocation { amount: usize },
    #[error("the edited Keynote slide-table merges failed semantic verification")]
    Verification,
    #[error("the Keynote slide-table merge patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    #[error("this Keynote source does not support focused slide-table merge reads")]
    UnsupportedSource,
    #[error("the requested Keynote slide-table merge graph has an unsupported dependency")]
    UnsupportedDependency,
    #[error("the requested Keynote slide-table merge topology is unsupported")]
    UnsupportedTopology,
    #[error("the Keynote slide-table merge selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    #[error("the requested Keynote slide-table merge region is invalid")]
    InvalidRegion,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. }
            | Self::UnsupportedDependency { .. }
            | Self::UnsupportedTopology { .. } => RefusalKind::UnsupportedContent,
            Self::AmbiguousSelector { .. }
            | Self::EmptySlideName { .. }
            | Self::InvalidRegion { .. } => RefusalKind::InvalidValue,
            Self::SlideNameNotFound { .. }
            | Self::SlidePositionNotFound { .. }
            | Self::TablePositionNotFound { .. } => RefusalKind::NotFound,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl Package {
    /// Read all validated merged-cell regions for one selected slide table.
    ///
    /// The selectors are resolved through the slide's owned drawable and
    /// z-order graph.  The selected table model and its merge-owner formula
    /// storage are then decoded with the same cumulative budget used for root
    /// selection.  This read is source-preserving and does not expose native
    /// object identifiers.
    pub fn slide_table_merges<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Vec<Region>, Error> {
        let mut budget = core::Budget::new(self).map_err(map_core_error)?;
        let target = core::select_table(self, slide.into(), table.into(), &mut budget)
            .map_err(map_core_error)?;
        self.decode_selected_merges(target, &mut budget)
    }

    /// Read one table from a checked shared component catalog.
    ///
    /// This ingress is intentionally private to the metadata-only reader.  It
    /// keeps the physical selector above source-preserving, while reusing the
    /// same target proof and lazy merge codec after component inventory has
    /// been charged by the dedicated read-only core path.
    pub(super) fn slide_table_merges_from_components<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Vec<Region>, Error> {
        let mut budget = core::Budget::new(self).map_err(map_core_error)?;
        let target =
            core::select_table_from_components(self, slide.into(), table.into(), &mut budget)
                .map_err(map_core_error)?;
        self.decode_selected_merges(target, &mut budget)
    }

    fn decode_selected_merges(
        &self,
        target: core::Target,
        budget: &mut core::Budget,
    ) -> Result<Vec<Region>, Error> {
        let source = core::model_payload(self, &target).map_err(map_core_error)?;
        let wire = budget.residual(self).map_err(map_core_error)?;
        // The shared decoder reserves one pair-slice, one index vector, and
        // one region vector for the complete selected store before decoding
        // formulas. Keep its exact reservation shape inside this package's
        // ledger so a large pair count cannot bypass allocation or scratch
        // ceilings merely because the final semantic vector is small.
        let max_regions = merge_region_capacity(
            budget.remaining_allocations().map_err(map_core_error)?,
            budget.remaining_scratch().map_err(map_core_error)?,
            budget.remaining_retained().map_err(map_core_error)?,
        );
        let limits = table_merges::ReadLimits {
            wire,
            max_regions,
            max_overlap_checks: budget.remaining_work().map_err(map_core_error)?,
        };
        let read = table_merges::read_table_merges(source, limits).map_err(|error| {
            charge_attempted(budget, &error);
            map_merge_error(error)
        })?;

        budget
            .input(read.report.input_bytes())
            .map_err(map_core_error)?;
        budget
            .fields(read.report.fields())
            .map_err(map_core_error)?;
        budget.work(read.report.work()).map_err(map_core_error)?;
        let region_count = read.regions.len();
        let allocations = region_count
            .checked_mul(MERGE_READER_ALLOCATIONS_PER_REGION)
            .ok_or(core::Error::InvalidSource)
            .map_err(map_core_error)?;
        budget.allocations(allocations).map_err(map_core_error)?;
        let scratch = region_count
            .checked_mul(MERGE_READER_SCRATCH_PER_REGION)
            .ok_or(core::Error::InvalidSource)
            .map_err(map_core_error)?;
        budget.scratch(scratch).map_err(map_core_error)?;
        let retained = read
            .regions
            .len()
            .checked_mul(REGION_BYTES)
            .ok_or(core::Error::InvalidSource)
            .map_err(map_core_error)?;
        budget.retained(retained).map_err(map_core_error)?;
        Ok(read.regions)
    }
}

fn merge_region_capacity(allocations: usize, scratch: usize, retained: usize) -> usize {
    allocations
        .checked_div(MERGE_READER_ALLOCATIONS_PER_REGION)
        .unwrap_or(0)
        .min(
            scratch
                .checked_div(MERGE_READER_SCRATCH_PER_REGION)
                .unwrap_or(0),
        )
        .min(retained.checked_div(REGION_BYTES).unwrap_or(0))
        .min(WireLimits::MAX_FIELDS)
}

fn charge_attempted(budget: &mut core::Budget, error: &MergeReadError) {
    let attempted = error.attempted();
    let _ = budget.input(attempted.input_bytes);
    let _ = budget.fields(attempted.fields);
    let _ = budget.work(attempted.work);
}

fn map_merge_error(error: MergeReadError) -> Error {
    if let Some(owner_limit) = error.owner_entries_limit() {
        return Error::Limit {
            kind: LimitKind::OwnerEntries,
            observed: owner_limit.observed() as u64,
            maximum: owner_limit.maximum() as u64,
        };
    }
    match error.error() {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: map_wire_limit(*kind),
            observed: *observed as u64,
            maximum: *limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            Error::Allocation { amount: *amount }
        },
        _ => Error::InvalidSource,
    }
}

fn map_wire_limit(kind: WireLimitKind) -> LimitKind {
    match kind {
        WireLimitKind::InputBytes => LimitKind::InputBytes,
        WireLimitKind::Fields => LimitKind::WireFields,
        WireLimitKind::OutputBytes => LimitKind::OutputBytes,
        WireLimitKind::Nesting => LimitKind::WireNesting,
        WireLimitKind::RewriteWork => LimitKind::WireWork,
        WireLimitKind::TableRows
        | WireLimitKind::TableColumns
        | WireLimitKind::TableCells
        | WireLimitKind::MaterializedCells => LimitKind::PayloadItems,
    }
}

fn map_core_error(error: core::Error) -> Error {
    match error {
        core::Error::UnsupportedSource => Error::Refused(Refusal::UnsupportedSource),
        core::Error::UnsupportedDependency => Error::Refused(Refusal::UnsupportedDependency),
        core::Error::UnsupportedTopology => Error::Refused(Refusal::UnsupportedTopology),
        core::Error::AmbiguousSelector => Error::Refused(Refusal::AmbiguousSelector),
        core::Error::EmptySlideName => Error::Refused(Refusal::EmptySlideName),
        core::Error::SlideNameNotFound => Error::Refused(Refusal::SlideNameNotFound),
        core::Error::SlidePositionNotFound(position) => {
            Error::Refused(Refusal::SlidePositionNotFound { position })
        },
        core::Error::TablePositionNotFound(position) => {
            Error::Refused(Refusal::TablePositionNotFound { position })
        },
        core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_core_limit(kind),
            observed,
            maximum,
        },
        core::Error::Allocation(amount) => Error::Allocation { amount },
        core::Error::InvalidSource
        | core::Error::Read
        | core::Error::Wire
        | core::Error::Codec
        | core::Error::Archive
        | core::Error::Verification => Error::InvalidSource,
    }
}

fn map_core_limit(kind: core::LimitKind) -> LimitKind {
    match kind {
        core::LimitKind::InputBytes => LimitKind::InputBytes,
        core::LimitKind::OutputBytes => LimitKind::OutputBytes,
        core::LimitKind::Entries => LimitKind::Entries,
        core::LimitKind::EntryBytes => LimitKind::EntryBytes,
        core::LimitKind::TotalBytes => LimitKind::TotalBytes,
        core::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
        core::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
        core::LimitKind::References => LimitKind::References,
        core::LimitKind::WireFields => LimitKind::WireFields,
        core::LimitKind::WireNesting => LimitKind::WireNesting,
        core::LimitKind::WireWork => LimitKind::WireWork,
        core::LimitKind::Allocations => LimitKind::Allocations,
        core::LimitKind::Retained => LimitKind::Retained,
        core::LimitKind::Scratch => LimitKind::Scratch,
        core::LimitKind::Components => LimitKind::Components,
    }
}

#[cfg(test)]
mod tests {
    use super::{MERGE_READER_SCRATCH_PER_REGION, REGION_BYTES, merge_region_capacity};

    #[test]
    fn zero_scratch_budget_cannot_admit_a_merge_region() {
        assert_eq!(
            merge_region_capacity(usize::MAX, MERGE_READER_SCRATCH_PER_REGION - 1, usize::MAX),
            0
        );
        assert_eq!(merge_region_capacity(usize::MAX, 0, REGION_BYTES), 0);
        assert_eq!(
            merge_region_capacity(3, MERGE_READER_SCRATCH_PER_REGION, REGION_BYTES),
            1
        );
    }
}
