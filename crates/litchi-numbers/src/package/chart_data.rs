//! Selector-first, archive-free Numbers chart data reads.
//!
//! The chart graph is rooted by the existing chart-arrangement walker.  Only
//! the selected drawable payload enters the strict borrowed chart-data codec;
//! labels and values are copied into the common semantic model after the
//! codec and package budgets have admitted the complete operation.

#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "The package boundary converts bounded native arithmetic and redacts graph failures."
)]

use std::fmt;
use std::mem::size_of;

use litchi_core::Position;
use litchi_iwa_common::chart::data::ChartData;
use litchi_iwa_protos::chart_data_codec as codec;
use thiserror::Error as ThisError;

use super::Package;
use super::chart_arrangement::{
    self, ChartBudget, select_chart_with_budget, selected_chart_payload,
};
use crate::sheet::chart::arrangement::Error as ChartArrangementError;
use crate::{ChartSelector, SheetSelector};

/// Resource category reported by a focused Numbers chart-data read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Source bytes inspected by the rooted graph or codec.
    WireBytes,
    /// Encoded fields visited by strict preflight or Buffa.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate strict projection work.
    WireWork,
    /// Numeric cells visited by the grid projection.
    Cells,
    /// Row and column labels retained by the projection.
    Labels,
    /// UTF-8 text retained by the semantic model.
    TextBytes,
    /// Logical temporary allocations.
    Allocations,
    /// Bytes retained by one read operation.
    RetainedBytes,
    /// Native references inspected by the selector.
    References,
    /// Compressed IWA member bytes inspected by the rooted owner.
    CompressedBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Cells => "cells",
            Self::Labels => "labels",
            Self::TextBytes => "text bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::References => "references",
            Self::CompressedBytes => "compressed bytes",
        })
    }
}

/// Failure from a semantic Numbers chart-data read.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted graph or selected payload is malformed.
    #[error("the Numbers chart-data source is invalid")]
    InvalidSource,
    /// A finite operation budget was exceeded.
    #[error("Numbers chart data {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource that exceeded its ceiling.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for Numbers chart data")]
    Allocation { amount: usize },
    /// The patch was created from another exact package artifact.
    #[error("the Numbers chart-data patch does not match the exact source package")]
    PatchConflict,
    /// Candidate reopening or semantic readback did not reproduce the target.
    #[error("the edited Numbers chart data failed semantic verification")]
    Verification,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The package does not retain a supported physical source.
    #[error("this Numbers source does not support chart data reads")]
    UnsupportedSource,
    /// The selected graph crosses an unsupported native dependency boundary.
    #[error("the requested Numbers chart-data graph is unsupported")]
    UnsupportedDependency,
    /// A semantic selector matched more than one native owner.
    #[error("the Numbers chart-data selector is ambiguous")]
    AmbiguousSelector,
    /// An empty sheet name was supplied.
    #[error("the Numbers sheet selector name cannot be empty")]
    EmptySheetName,
    /// No sheet matched the requested name.
    #[error("the Numbers workbook has no sheet matching the requested name")]
    SheetNameNotFound,
    /// No sheet matched the requested position.
    #[error("the Numbers workbook has no sheet at position {position:?}")]
    SheetPositionNotFound { position: Position },
    /// No chart matched the requested position.
    #[error("the selected Numbers sheet has no chart at position {position:?}")]
    ChartPositionNotFound { position: Position },
    /// Chart-data writes must preserve the source axes and rectangular shape.
    #[error("Numbers chart-data edits must preserve the existing labels and dimensions")]
    ShapeMismatch,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. }
            | Self::UnsupportedDependency { .. }
            | Self::ShapeMismatch { .. } => RefusalKind::UnsupportedContent,
            Self::AmbiguousSelector { .. } | Self::EmptySheetName { .. } => {
                RefusalKind::InvalidValue
            },
            Self::SheetNameNotFound { .. }
            | Self::SheetPositionNotFound { .. }
            | Self::ChartPositionNotFound { .. } => RefusalKind::NotFound,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl Package {
    /// Read one modern sheet chart's rectangular inline numeric data through
    /// semantic selectors.
    ///
    /// Numeric cells are returned as finite `f64` values and empty or
    /// date/duration-only cells as `None`. Legacy chart payloads and grids
    /// whose dimensions or numeric values do not satisfy this contract are
    /// rejected without exposing native identifiers or wire objects.
    pub fn sheet_chart_data<'sheet>(
        &self,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        chart_selector: impl Into<ChartSelector>,
    ) -> Result<ChartData, Error> {
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement_error)?;
        let selection = select_chart_with_budget(
            self,
            sheet_selector.into(),
            chart_selector.into(),
            false,
            &mut budget,
        )
        .map_err(map_arrangement_error)?;
        read_selected_chart_data(self, &selection, &mut budget)
    }
}

/// Read the selected chart through an already-admitted chart transaction
/// budget.  Focused chart-data writers use this after selecting and checking
/// the same drawable so the selector walk and codec projection are charged
/// exactly once per operation.
pub(super) fn read_selected_chart_data(
    package: &Package,
    selection: &chart_arrangement::ChartSelection,
    budget: &mut ChartBudget,
) -> Result<ChartData, Error> {
    let payload = selected_chart_payload(package, selection).map_err(map_arrangement_error)?;
    let options = budget.data_codec_options().map_err(map_arrangement_error)?;
    let (snapshot, report) = match codec::decode_modern_with_report(payload, &options) {
        Ok(decoded) => decoded,
        Err(error) => {
            let report = error.report();
            if let Err(budget_error) = budget.data_codec_report(report) {
                return Err(map_arrangement_error(budget_error));
            }
            if let Some(limit) = map_codec_limit(&error) {
                return Err(limit);
            }
            return Err(Error::InvalidSource);
        },
    };
    budget
        .data_codec_report(report)
        .map_err(map_arrangement_error)?;

    // Label and row iterators intentionally remain lazy in the codec, so
    // account for their bounded source walks and semantic copy loop before
    // attempting any fallible allocation or ownership transfer.
    let materialization_work = materialization_work(snapshot, report)?;
    budget
        .data_materialization_work(materialization_work)
        .map_err(map_arrangement_error)?;

    // The strict codec has already checked dimensions and finite numeric
    // values.  These helpers still preflight every owned allocation so a
    // fallible copy cannot publish a partially materialized model.
    let row_names = own_labels(snapshot.row_labels(), budget)?;
    let column_names = own_labels(snapshot.column_labels(), budget)?;
    let values = own_values(snapshot, budget)?;

    ChartData::new(row_names, column_names, values).map_err(|_| Error::InvalidSource)
}

fn materialization_work(
    snapshot: codec::ChartDataSnapshot<'_>,
    report: codec::DecodeReport,
) -> Result<usize, Error> {
    let source_walks = snapshot
        .grid_source()
        .len()
        .checked_mul(10)
        .ok_or(Error::InvalidSource)?;
    let copy_work = report
        .text_bytes()
        .checked_add(report.cell_count())
        .and_then(|amount| amount.checked_add(snapshot.row_count()))
        .ok_or(Error::InvalidSource)?;
    source_walks
        .checked_add(copy_work)
        .ok_or(Error::InvalidSource)
}

fn own_labels(
    labels: codec::LabelList<'_>,
    budget: &mut ChartBudget,
) -> Result<Vec<String>, Error> {
    let count = labels.len();
    let text_bytes = labels
        .iter()
        .try_fold(0usize, |total, value| total.checked_add(value.len()))
        .ok_or(Error::InvalidSource)?;
    let retained = count
        .checked_mul(size_of::<String>())
        .and_then(|amount| amount.checked_add(text_bytes))
        .ok_or(Error::InvalidSource)?;
    let string_allocations = labels.iter().filter(|value| !value.is_empty()).count();
    let allocations = usize::from(count != 0)
        .checked_add(string_allocations)
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;

    let mut output = Vec::new();
    output
        .try_reserve_exact(count)
        .map_err(|_| Error::Allocation { amount: count })?;
    for value in labels.iter() {
        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|_| Error::Allocation {
                amount: value.len(),
            })?;
        owned.push_str(value);
        output.push(owned);
    }
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(output)
}

fn own_values(
    snapshot: codec::ChartDataSnapshot<'_>,
    budget: &mut ChartBudget,
) -> Result<Vec<Vec<Option<f64>>>, Error> {
    let row_count = snapshot.row_count();
    let column_count = snapshot.column_count();
    let cell_count = row_count
        .checked_mul(column_count)
        .ok_or(Error::InvalidSource)?;
    let row_slots = row_count
        .checked_mul(size_of::<Vec<Option<f64>>>())
        .ok_or(Error::InvalidSource)?;
    let cell_slots = cell_count
        .checked_mul(size_of::<Option<f64>>())
        .ok_or(Error::InvalidSource)?;
    let retained = row_slots
        .checked_add(cell_slots)
        .ok_or(Error::InvalidSource)?;
    let allocations = usize::from(row_count != 0)
        .checked_add(row_count)
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;

    let mut values = Vec::new();
    values
        .try_reserve_exact(row_count)
        .map_err(|_| Error::Allocation { amount: row_count })?;
    for row in snapshot.rows().iter() {
        if row.len() != column_count {
            return Err(Error::InvalidSource);
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(column_count)
            .map_err(|_| Error::Allocation {
                amount: column_count,
            })?;
        for value in row.values() {
            owned.push(value);
        }
        if owned.len() != column_count {
            return Err(Error::InvalidSource);
        }
        values.push(owned);
    }
    if values.len() != row_count {
        return Err(Error::InvalidSource);
    }
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(values)
}

pub(super) fn map_arrangement_error(error: ChartArrangementError) -> Error {
    match error {
        ChartArrangementError::Refused(chart_arrangement::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        ChartArrangementError::Refused(chart_arrangement::Refusal::UnsupportedDependency) => {
            Error::Refused(Refusal::UnsupportedDependency)
        },
        ChartArrangementError::Refused(chart_arrangement::Refusal::AmbiguousSelector) => {
            Error::Refused(Refusal::AmbiguousSelector)
        },
        ChartArrangementError::Refused(chart_arrangement::Refusal::EmptySheetName) => {
            Error::Refused(Refusal::EmptySheetName)
        },
        ChartArrangementError::Refused(chart_arrangement::Refusal::SheetNameNotFound) => {
            Error::Refused(Refusal::SheetNameNotFound)
        },
        ChartArrangementError::Refused(chart_arrangement::Refusal::SheetPositionNotFound {
            position,
        }) => Error::Refused(Refusal::SheetPositionNotFound { position }),
        ChartArrangementError::Refused(chart_arrangement::Refusal::ChartPositionNotFound {
            position,
        }) => Error::Refused(Refusal::ChartPositionNotFound { position }),
        ChartArrangementError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_arrangement_limit_kind(kind),
            observed,
            maximum,
        },
        ChartArrangementError::Allocation { amount } => Error::Allocation { amount },
        ChartArrangementError::InvalidSource => Error::InvalidSource,
        ChartArrangementError::Verification => Error::Verification,
        ChartArrangementError::PatchConflict => Error::PatchConflict,
    }
}

const fn map_arrangement_limit_kind(kind: chart_arrangement::LimitKind) -> LimitKind {
    match kind {
        chart_arrangement::LimitKind::WireBytes | chart_arrangement::LimitKind::OutputBytes => {
            LimitKind::WireBytes
        },
        chart_arrangement::LimitKind::WireFields => LimitKind::WireFields,
        chart_arrangement::LimitKind::WireNesting => LimitKind::WireNesting,
        chart_arrangement::LimitKind::WireWork | chart_arrangement::LimitKind::ScratchBytes => {
            LimitKind::WireWork
        },
        chart_arrangement::LimitKind::Allocations => LimitKind::Allocations,
        chart_arrangement::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
        chart_arrangement::LimitKind::References | chart_arrangement::LimitKind::Identifiers => {
            LimitKind::References
        },
        chart_arrangement::LimitKind::CompressedBytes => LimitKind::CompressedBytes,
    }
}

pub(super) fn map_codec_limit(error: &codec::DecodeError) -> Option<Error> {
    let (kind, observed, maximum) = match error.resource_limit()? {
        codec::DecodeLimit::Bytes { observed, maximum } => {
            (LimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (LimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::Work { observed, maximum } => (LimitKind::WireWork, observed, maximum),
        codec::DecodeLimit::Nesting { observed, maximum } => {
            (LimitKind::WireNesting, observed as usize, maximum as usize)
        },
        codec::DecodeLimit::Cells { observed, maximum } => (LimitKind::Cells, observed, maximum),
        codec::DecodeLimit::Labels { observed, maximum } => (LimitKind::Labels, observed, maximum),
        codec::DecodeLimit::Text { observed, maximum } => (LimitKind::TextBytes, observed, maximum),
        _ => return None,
    };
    Some(Error::Limit {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    })
}
