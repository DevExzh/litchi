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
use thiserror::Error;

use super::Package;
use super::chart_arrangement::{
    self, ChartArrangementError, ChartBudget, select_chart_with_budget, selected_chart_payload,
};
use crate::{ChartSelector, SheetSelector};

/// Resource category reported by a focused Numbers chart-data read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartDataLimitKind {
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
}

impl fmt::Display for ChartDataLimitKind {
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
        })
    }
}

/// Failure from a semantic Numbers chart-data read.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartDataError {
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
    /// The rooted graph or selected payload is malformed.
    #[error("the Numbers chart-data source is invalid")]
    InvalidSource,
    /// A finite operation budget was exceeded.
    #[error("Numbers chart data {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource that exceeded its ceiling.
        kind: ChartDataLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for Numbers chart data")]
    Allocation { amount: usize },
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
    ) -> Result<ChartData, ChartDataError> {
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement_error)?;
        let selection = select_chart_with_budget(
            self,
            sheet_selector.into(),
            chart_selector.into(),
            false,
            &mut budget,
        )
        .map_err(map_arrangement_error)?;
        let payload = selected_chart_payload(self, &selection).map_err(map_arrangement_error)?;
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
                return Err(ChartDataError::InvalidSource);
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
        let row_names = own_labels(snapshot.row_labels(), &mut budget)?;
        let column_names = own_labels(snapshot.column_labels(), &mut budget)?;
        let values = own_values(snapshot, &mut budget)?;

        ChartData::new(row_names, column_names, values).map_err(|_| ChartDataError::InvalidSource)
    }
}

fn materialization_work(
    snapshot: codec::ChartDataSnapshot<'_>,
    report: codec::DecodeReport,
) -> Result<usize, ChartDataError> {
    let source_walks = snapshot
        .grid_source()
        .len()
        .checked_mul(10)
        .ok_or(ChartDataError::InvalidSource)?;
    let copy_work = report
        .text_bytes()
        .checked_add(report.cell_count())
        .and_then(|amount| amount.checked_add(snapshot.row_count()))
        .ok_or(ChartDataError::InvalidSource)?;
    source_walks
        .checked_add(copy_work)
        .ok_or(ChartDataError::InvalidSource)
}

fn own_labels(
    labels: codec::LabelList<'_>,
    budget: &mut ChartBudget,
) -> Result<Vec<String>, ChartDataError> {
    let count = labels.len();
    let text_bytes = labels
        .iter()
        .try_fold(0usize, |total, value| total.checked_add(value.len()))
        .ok_or(ChartDataError::InvalidSource)?;
    let retained = count
        .checked_mul(size_of::<String>())
        .and_then(|amount| amount.checked_add(text_bytes))
        .ok_or(ChartDataError::InvalidSource)?;
    let string_allocations = labels.iter().filter(|value| !value.is_empty()).count();
    let allocations = usize::from(count != 0)
        .checked_add(string_allocations)
        .ok_or(ChartDataError::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;

    let mut output = Vec::new();
    output
        .try_reserve_exact(count)
        .map_err(|_| ChartDataError::Allocation { amount: count })?;
    for value in labels.iter() {
        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|_| ChartDataError::Allocation {
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
) -> Result<Vec<Vec<Option<f64>>>, ChartDataError> {
    let row_count = snapshot.row_count();
    let column_count = snapshot.column_count();
    let cell_count = row_count
        .checked_mul(column_count)
        .ok_or(ChartDataError::InvalidSource)?;
    let row_slots = row_count
        .checked_mul(size_of::<Vec<Option<f64>>>())
        .ok_or(ChartDataError::InvalidSource)?;
    let cell_slots = cell_count
        .checked_mul(size_of::<Option<f64>>())
        .ok_or(ChartDataError::InvalidSource)?;
    let retained = row_slots
        .checked_add(cell_slots)
        .ok_or(ChartDataError::InvalidSource)?;
    let allocations = usize::from(row_count != 0)
        .checked_add(row_count)
        .ok_or(ChartDataError::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;

    let mut values = Vec::new();
    values
        .try_reserve_exact(row_count)
        .map_err(|_| ChartDataError::Allocation { amount: row_count })?;
    for row in snapshot.rows().iter() {
        if row.len() != column_count {
            return Err(ChartDataError::InvalidSource);
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(column_count)
            .map_err(|_| ChartDataError::Allocation {
                amount: column_count,
            })?;
        for value in row.values() {
            owned.push(value);
        }
        if owned.len() != column_count {
            return Err(ChartDataError::InvalidSource);
        }
        values.push(owned);
    }
    if values.len() != row_count {
        return Err(ChartDataError::InvalidSource);
    }
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(values)
}

fn map_arrangement_error(error: ChartArrangementError) -> ChartDataError {
    match error {
        ChartArrangementError::UnsupportedSource => ChartDataError::UnsupportedSource,
        ChartArrangementError::UnsupportedDependency => ChartDataError::UnsupportedDependency,
        ChartArrangementError::AmbiguousSelector => ChartDataError::AmbiguousSelector,
        ChartArrangementError::EmptySheetName => ChartDataError::EmptySheetName,
        ChartArrangementError::SheetNameNotFound => ChartDataError::SheetNameNotFound,
        ChartArrangementError::SheetPositionNotFound { position } => {
            ChartDataError::SheetPositionNotFound { position }
        },
        ChartArrangementError::ChartPositionNotFound { position } => {
            ChartDataError::ChartPositionNotFound { position }
        },
        ChartArrangementError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ChartDataError::LimitExceeded {
            kind: map_arrangement_limit_kind(kind),
            observed,
            maximum,
        },
        ChartArrangementError::Allocation { amount } => ChartDataError::Allocation { amount },
        ChartArrangementError::InvalidSource
        | ChartArrangementError::Verification
        | ChartArrangementError::PatchConflict => ChartDataError::InvalidSource,
    }
}

const fn map_arrangement_limit_kind(
    kind: chart_arrangement::ChartArrangementLimitKind,
) -> ChartDataLimitKind {
    match kind {
        chart_arrangement::ChartArrangementLimitKind::WireBytes
        | chart_arrangement::ChartArrangementLimitKind::OutputBytes => {
            ChartDataLimitKind::WireBytes
        },
        chart_arrangement::ChartArrangementLimitKind::WireFields => ChartDataLimitKind::WireFields,
        chart_arrangement::ChartArrangementLimitKind::WireNesting => {
            ChartDataLimitKind::WireNesting
        },
        chart_arrangement::ChartArrangementLimitKind::WireWork
        | chart_arrangement::ChartArrangementLimitKind::ScratchBytes => {
            ChartDataLimitKind::WireWork
        },
        chart_arrangement::ChartArrangementLimitKind::Allocations => {
            ChartDataLimitKind::Allocations
        },
        chart_arrangement::ChartArrangementLimitKind::RetainedBytes => {
            ChartDataLimitKind::RetainedBytes
        },
        chart_arrangement::ChartArrangementLimitKind::References => ChartDataLimitKind::References,
    }
}

fn map_codec_limit(error: &codec::DecodeError) -> Option<ChartDataError> {
    let (kind, observed, maximum) = match error.resource_limit()? {
        codec::DecodeLimit::Bytes { observed, maximum } => {
            (ChartDataLimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (ChartDataLimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::Work { observed, maximum } => {
            (ChartDataLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => (
            ChartDataLimitKind::WireNesting,
            observed as usize,
            maximum as usize,
        ),
        codec::DecodeLimit::Cells { observed, maximum } => {
            (ChartDataLimitKind::Cells, observed, maximum)
        },
        codec::DecodeLimit::Labels { observed, maximum } => {
            (ChartDataLimitKind::Labels, observed, maximum)
        },
        codec::DecodeLimit::Text { observed, maximum } => {
            (ChartDataLimitKind::TextBytes, observed, maximum)
        },
        _ => return None,
    };
    Some(ChartDataError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    })
}
