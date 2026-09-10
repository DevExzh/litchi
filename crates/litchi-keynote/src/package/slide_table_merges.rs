//! Selector-first readback of native merged-cell regions in Keynote tables.
//!
//! Table selection and package authority remain in [`super::slide_table_core`]
//! while the selected model's merge-owner wire path is decoded by the shared
//! Numbers wire adapter.  The operation keeps one resource ledger from root
//! selection through the borrowed merge scan; it never mutates or reassembles
//! the source package.

use std::fmt;

use litchi_core::Position;
use litchi_iwa_common::{LimitKind as WireLimitKind, WireLimits, table::merge::Region};
use litchi_numbers_wire::table_merges::{
    self, FORMULA_INDEX_BYTES, MergeReadError, PAIR_REFERENCE_BYTES, REGION_BYTES,
};
use thiserror::Error;

use super::Package;
use super::slide_table_core as core;
use crate::SlideSelector;
use crate::slide::table::TableSelector;

const MERGE_READER_ALLOCATIONS_PER_REGION: usize = 3;
const MERGE_READER_SCRATCH_PER_REGION: usize = PAIR_REFERENCE_BYTES + FORMULA_INDEX_BYTES;

/// Finite resources charged by a focused Keynote merged-cell read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableMergesLimitKind {
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

impl fmt::Display for SlideTableMergesLimitKind {
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
pub enum SlideTableMergesPath {
    /// The complete package.
    Package,
    /// One checked slide/table selection.
    Table { slide: Position, table: Position },
}

impl fmt::Display for SlideTableMergesPath {
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

/// Failure from a focused Keynote merged-cell read.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTableMergesError {
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
    #[error("the selected Keynote slide-table merge source is invalid")]
    InvalidSource,
    #[error(
        "Keynote slide-table merges {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideTableMergesLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote slide-table merge read")]
    Allocation { amount: usize },
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
    ) -> Result<Vec<Region>, SlideTableMergesError> {
        let mut budget = core::Budget::new(self).map_err(map_core_error)?;
        let target = core::select_table(self, slide.into(), table.into(), &mut budget)
            .map_err(map_core_error)?;
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
            charge_attempted(&mut budget, &error);
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

fn map_merge_error(error: MergeReadError) -> SlideTableMergesError {
    match error.error() {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SlideTableMergesError::LimitExceeded {
            kind: map_wire_limit(*kind),
            observed: *observed as u64,
            maximum: *limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideTableMergesError::Allocation { amount: *amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => SlideTableMergesError::InvalidSource,
    }
}

fn map_wire_limit(kind: WireLimitKind) -> SlideTableMergesLimitKind {
    match kind {
        WireLimitKind::InputBytes => SlideTableMergesLimitKind::InputBytes,
        WireLimitKind::Fields => SlideTableMergesLimitKind::WireFields,
        WireLimitKind::OutputBytes => SlideTableMergesLimitKind::OutputBytes,
        WireLimitKind::Nesting => SlideTableMergesLimitKind::WireNesting,
        WireLimitKind::RewriteWork => SlideTableMergesLimitKind::WireWork,
        WireLimitKind::TableRows
        | WireLimitKind::TableColumns
        | WireLimitKind::TableCells
        | WireLimitKind::MaterializedCells => SlideTableMergesLimitKind::PayloadItems,
    }
}

fn map_core_error(error: core::Error) -> SlideTableMergesError {
    match error {
        core::Error::UnsupportedSource => SlideTableMergesError::UnsupportedSource,
        core::Error::UnsupportedDependency => SlideTableMergesError::UnsupportedDependency,
        core::Error::UnsupportedTopology => SlideTableMergesError::UnsupportedTopology,
        core::Error::AmbiguousSelector => SlideTableMergesError::AmbiguousSelector,
        core::Error::EmptySlideName => SlideTableMergesError::EmptySlideName,
        core::Error::SlideNameNotFound => SlideTableMergesError::SlideNameNotFound,
        core::Error::SlidePositionNotFound(position) => {
            SlideTableMergesError::SlidePositionNotFound { position }
        },
        core::Error::TablePositionNotFound(position) => {
            SlideTableMergesError::TablePositionNotFound { position }
        },
        core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableMergesError::LimitExceeded {
            kind: map_core_limit(kind),
            observed,
            maximum,
        },
        core::Error::Allocation(amount) => SlideTableMergesError::Allocation { amount },
        core::Error::InvalidSource
        | core::Error::Read
        | core::Error::Wire
        | core::Error::Codec
        | core::Error::Archive
        | core::Error::Verification => SlideTableMergesError::InvalidSource,
    }
}

fn map_core_limit(kind: core::LimitKind) -> SlideTableMergesLimitKind {
    match kind {
        core::LimitKind::InputBytes => SlideTableMergesLimitKind::InputBytes,
        core::LimitKind::OutputBytes => SlideTableMergesLimitKind::OutputBytes,
        core::LimitKind::Entries => SlideTableMergesLimitKind::Entries,
        core::LimitKind::EntryBytes => SlideTableMergesLimitKind::EntryBytes,
        core::LimitKind::TotalBytes => SlideTableMergesLimitKind::TotalBytes,
        core::LimitKind::PayloadObjects => SlideTableMergesLimitKind::PayloadObjects,
        core::LimitKind::PayloadMessages => SlideTableMergesLimitKind::PayloadMessages,
        core::LimitKind::References => SlideTableMergesLimitKind::References,
        core::LimitKind::WireFields => SlideTableMergesLimitKind::WireFields,
        core::LimitKind::WireNesting => SlideTableMergesLimitKind::WireNesting,
        core::LimitKind::WireWork => SlideTableMergesLimitKind::WireWork,
        core::LimitKind::Allocations => SlideTableMergesLimitKind::Allocations,
        core::LimitKind::Retained => SlideTableMergesLimitKind::Retained,
        core::LimitKind::Scratch => SlideTableMergesLimitKind::Scratch,
        core::LimitKind::Components => SlideTableMergesLimitKind::Components,
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
