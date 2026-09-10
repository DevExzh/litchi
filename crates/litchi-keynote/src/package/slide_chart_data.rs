//! Selector-first, archive-free Keynote chart data reads.
//!
//! The chart graph is resolved by the existing chart-title authority. This
//! adapter only owns the final semantic copy of the selected rectangular grid;
//! native identifiers, archive objects, and generated protobuf values remain
//! private to the package and protocol crates.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "The package boundary maps native failures to content-redacted semantic errors."
)]

use std::{fmt, mem::size_of};

use litchi_core::Position;
use litchi_iwa_common::chart::data::{ChartData, DataError};
use litchi_iwa_protos::chart_data_codec::{self, DecodeOptions, LabelList};
use thiserror::Error;

use super::Package;
use super::chart_axis_support::{self, AxisSupportBudget, AxisSupportError};
use super::slide_chart_title::{
    self, ChartGraphScanBudget, ChartSelection, ChartTitleError, ChartTitleLimitKind,
    select_chart_with_budget,
};
use crate::{ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = chart_data_codec::MODERN_CHART_DRAWABLE_MESSAGE_TYPE;
const MAX_DATA_CELL_COUNT: usize = 1_000_000;
const MAX_DATA_LABEL_COUNT: usize = 1_000_000;
const MAX_DATA_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// A finite resource governed while one chart-data view is prepared.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideChartDataLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// Native payload bytes inspected by the selector.
    PayloadBytes,
    /// Native references inspected by the selector.
    PayloadReferences,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate graph and codec work.
    WireWork,
    /// Decoder allocation units.
    WireAllocations,
    /// Decoder-retained bytes.
    WireRetainedBytes,
    /// Numeric grid cells visited.
    Cells,
    /// Semantic slide count.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic text-storage objects.
    TextStorages,
    /// Semantic text fragments.
    TextFragments,
    /// Aggregate semantic text bytes.
    TextBytes,
    /// ZIP/IWA entries and objects.
    Entries,
    /// Bytes in one ZIP/IWA entry or message.
    EntryBytes,
    /// Aggregate retained package bytes.
    TotalBytes,
    /// Allocations needed by the data view.
    Allocations,
    /// Bytes retained by the data view.
    RetainedBytes,
    /// Number of borrowed row/column labels.
    LabelCount,
}

impl fmt::Display for SlideChartDataLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::WireBytes => "wire bytes",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadReferences => "payload references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::WireAllocations => "wire allocations",
            Self::WireRetainedBytes => "wire retained bytes",
            Self::Cells => "cells",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::LabelCount => "label count",
        })
    }
}

/// A content-redacted failure raised by a chart-data read.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideChartDataError {
    /// The source was not retained as an exact physical package.
    #[error("this Keynote source does not support physical chart data reads")]
    UnsupportedSource,
    /// An exact-name selector was ambiguous.
    #[error("the Keynote chart data selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound { position: Position },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// The selected chart graph or data payload was malformed.
    #[error("the Keynote chart data source is malformed or unsupported")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote chart data {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: SlideChartDataLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded semantic allocation failed before the result was published.
    #[error("could not allocate {amount} units for Keynote chart data")]
    Allocation { amount: usize },
}

impl Package {
    /// Read one selected chart's modern inline rectangular grid.
    ///
    /// Numeric cells are returned as finite `f64` values and missing,
    /// date-only, or duration-only native cells become `None`. Legacy chart
    /// payloads and graphs that do not prove the modern inline grid are
    /// rejected as [`SlideChartDataError::InvalidSource`].
    pub fn slide_chart_data<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartData, SlideChartDataError> {
        let mut budget = ChartGraphScanBudget::new(self).map_err(map_chart_title_error)?;
        budget
            .charge_selection_scans(self, true)
            .map_err(map_axis_support_error)?;
        let selection = select_chart_with_budget(
            self,
            slide_selector.into(),
            chart_selector.into(),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        read_selected_data(self, selection, &mut budget)
    }
}

fn read_selected_data(
    package: &Package,
    selection: ChartSelection,
    budget: &mut ChartGraphScanBudget,
) -> Result<ChartData, SlideChartDataError> {
    let (component_name, chart_object) = package
        .object_with_component(selection.chart_identifier)
        .ok_or(SlideChartDataError::InvalidSource)?;
    if component_name != selection.slide_component_name {
        return Err(SlideChartDataError::InvalidSource);
    }
    let (message_index, message) =
        chart_axis_support::unique_message(chart_object, CHART_MESSAGE_TYPE, budget)
            .map_err(map_axis_support_error)?;
    chart_axis_support::validate_selected_message_metadata(chart_object, message_index)
        .map_err(map_axis_support_error)?;

    let (limits, residual_work) = budget.chart_metadata_residual_limits();
    if residual_work == 0 {
        return Err(SlideChartDataError::LimitExceeded {
            kind: SlideChartDataLimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let maximum_depth = u32::try_from(limits.max_nesting()).map_err(|_error| {
        SlideChartDataError::LimitExceeded {
            kind: SlideChartDataLimitKind::WireNesting,
            observed: usize_to_u64(limits.max_nesting()),
            maximum: u64::from(u32::MAX),
        }
    })?;
    let options = DecodeOptions::new(
        message.data.len().max(1).min(limits.max_input_bytes()),
        limits.max_fields().min(residual_work).max(1),
        residual_work,
        maximum_depth,
        limits
            .max_fields()
            .min(residual_work)
            .clamp(1, MAX_DATA_CELL_COUNT),
        limits
            .max_fields()
            .min(residual_work)
            .clamp(1, MAX_DATA_LABEL_COUNT),
        limits
            .max_input_bytes()
            .min(residual_work)
            .clamp(1, MAX_DATA_TEXT_BYTES),
    );
    let (snapshot, report) =
        chart_data_codec::decode_modern_with_report(message.data.as_slice(), &options).map_err(
            |error| {
                let charge = charge_data_report(budget, error.report());
                match charge {
                    Ok(()) => map_chart_data_decode_error(error),
                    Err(budget_error) => budget_error,
                }
            },
        )?;
    charge_data_report(budget, report)?;
    charge_snapshot_projection(snapshot, budget)?;

    let row_names = copy_labels(snapshot.row_labels(), budget)?;
    let column_names = copy_labels(snapshot.column_labels(), budget)?;
    let values = copy_values(snapshot, budget)?;
    ChartData::new(row_names, column_names, values).map_err(map_data_error)
}

fn charge_data_report(
    budget: &mut ChartGraphScanBudget,
    report: chart_data_codec::DecodeReport,
) -> Result<(), SlideChartDataError> {
    let amount = report
        .source_bytes()
        .checked_add(report.fields())
        .and_then(|value| value.checked_add(report.work_bytes()))
        .and_then(|value| value.checked_add(report.max_depth() as usize))
        .and_then(|value| value.checked_add(report.cell_count()))
        .and_then(|value| value.checked_add(report.label_count()))
        .and_then(|value| value.checked_add(report.text_bytes()))
        .and_then(|value| value.checked_add(report.allocations()))
        .and_then(|value| value.checked_add(report.retained_bytes()))
        .ok_or(SlideChartDataError::InvalidSource)?;
    budget.charge_work(amount).map_err(map_axis_support_error)
}

/// Reserve the repeated borrowed-view walks performed while materializing the
/// common model.  The codec report accounts for its own strict preflight and
/// scalar projection; this envelope covers the two label passes, row framing,
/// and value-field walks owned by this package adapter.
fn charge_snapshot_projection(
    snapshot: chart_data_codec::ChartDataSnapshot<'_>,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), SlideChartDataError> {
    let cells = snapshot
        .row_count()
        .checked_mul(snapshot.column_count())
        .ok_or(SlideChartDataError::InvalidSource)?;
    let amount = snapshot
        .grid_source()
        .len()
        .max(1)
        .checked_mul(10)
        .and_then(|value| value.checked_add(cells))
        .ok_or(SlideChartDataError::InvalidSource)?;
    preflight_data_budget(budget, SlideChartDataLimitKind::WireWork, amount)?;
    budget.charge_work(amount).map_err(map_axis_support_error)
}

fn copy_labels<'source>(
    labels: LabelList<'source>,
    budget: &mut ChartGraphScanBudget,
) -> Result<Vec<String>, SlideChartDataError> {
    let count = labels.len();
    let mut text_bytes = 0usize;
    let mut string_allocations = 0usize;
    let mut seen = 0usize;
    for label in labels.iter() {
        seen = seen
            .checked_add(1)
            .ok_or(SlideChartDataError::InvalidSource)?;
        text_bytes = text_bytes
            .checked_add(label.len())
            .ok_or(SlideChartDataError::InvalidSource)?;
        string_allocations = string_allocations
            .checked_add(usize::from(!label.is_empty()))
            .ok_or(SlideChartDataError::InvalidSource)?;
    }
    if seen != count {
        return Err(SlideChartDataError::InvalidSource);
    }
    let allocations = usize::from(count != 0)
        .checked_add(string_allocations)
        .ok_or(SlideChartDataError::InvalidSource)?;
    let retained = count
        .checked_mul(size_of::<String>())
        .and_then(|value| value.checked_add(text_bytes))
        .ok_or(SlideChartDataError::InvalidSource)?;
    let charge = count
        .checked_add(text_bytes)
        .and_then(|value| value.checked_add(allocations))
        .and_then(|value| value.checked_add(retained))
        .ok_or(SlideChartDataError::InvalidSource)?;
    preflight_data_budget(budget, SlideChartDataLimitKind::Allocations, allocations)?;
    preflight_data_budget(budget, SlideChartDataLimitKind::RetainedBytes, retained)?;
    preflight_data_budget(budget, SlideChartDataLimitKind::WireWork, charge)?;

    let mut copied = Vec::new();
    copied
        .try_reserve_exact(count)
        .map_err(|_error| SlideChartDataError::Allocation { amount: count })?;
    for label in labels.iter() {
        let mut owned = String::new();
        owned
            .try_reserve_exact(label.len())
            .map_err(|_error| SlideChartDataError::Allocation {
                amount: label.len(),
            })?;
        owned.push_str(label);
        copied.push(owned);
    }
    if copied.len() != count {
        return Err(SlideChartDataError::InvalidSource);
    }
    budget.charge_work(charge).map_err(map_axis_support_error)?;
    Ok(copied)
}

fn copy_values(
    snapshot: chart_data_codec::ChartDataSnapshot<'_>,
    budget: &mut ChartGraphScanBudget,
) -> Result<Vec<Vec<Option<f64>>>, SlideChartDataError> {
    let rows = snapshot.row_count();
    let columns = snapshot.column_count();
    let cells = rows
        .checked_mul(columns)
        .ok_or(SlideChartDataError::InvalidSource)?;
    let value_work = rows
        .checked_add(cells)
        .and_then(|value| value.checked_add(rows.checked_mul(size_of::<Vec<Option<f64>>>())?))
        .and_then(|value| value.checked_add(cells.checked_mul(size_of::<Option<f64>>())?))
        .ok_or(SlideChartDataError::InvalidSource)?;
    let allocations = usize::from(rows != 0)
        .checked_add(rows)
        .ok_or(SlideChartDataError::InvalidSource)?;
    let retained = rows
        .checked_mul(size_of::<Vec<Option<f64>>>())
        .and_then(|value| value.checked_add(cells.checked_mul(size_of::<Option<f64>>())?))
        .ok_or(SlideChartDataError::InvalidSource)?;
    let charge = value_work
        .checked_add(allocations)
        .ok_or(SlideChartDataError::InvalidSource)?;
    preflight_data_budget(budget, SlideChartDataLimitKind::Allocations, allocations)?;
    preflight_data_budget(budget, SlideChartDataLimitKind::RetainedBytes, retained)?;
    preflight_data_budget(budget, SlideChartDataLimitKind::WireWork, charge)?;

    let mut copied = Vec::new();
    copied
        .try_reserve_exact(rows)
        .map_err(|_error| SlideChartDataError::Allocation { amount: rows })?;
    for row in snapshot.rows().iter() {
        if row.len() != columns {
            return Err(SlideChartDataError::InvalidSource);
        }
        let mut values = Vec::new();
        values
            .try_reserve_exact(columns)
            .map_err(|_error| SlideChartDataError::Allocation { amount: columns })?;
        for value in row.values() {
            values.push(value);
        }
        if values.len() != columns {
            return Err(SlideChartDataError::InvalidSource);
        }
        copied.push(values);
    }
    if copied.len() != rows {
        return Err(SlideChartDataError::InvalidSource);
    }
    budget.charge_work(charge).map_err(map_axis_support_error)?;
    Ok(copied)
}

fn preflight_data_budget(
    budget: &ChartGraphScanBudget,
    kind: SlideChartDataLimitKind,
    amount: usize,
) -> Result<(), SlideChartDataError> {
    // ChartGraphScanBudget deliberately exposes one aggregate operation
    // ledger shared by selector and metadata readers.  Allocation and
    // retained-byte checks use that same finite remainder, but run before
    // try_reserve so a rejected semantic copy cannot partially materialize.
    let (_, remaining) = budget.chart_metadata_residual_limits();
    if amount > remaining {
        return Err(SlideChartDataError::LimitExceeded {
            kind,
            observed: usize_to_u64(amount),
            maximum: usize_to_u64(remaining),
        });
    }
    Ok(())
}

fn map_data_error(_error: DataError) -> SlideChartDataError {
    SlideChartDataError::InvalidSource
}

fn map_axis_support_error(error: AxisSupportError) -> SlideChartDataError {
    map_chart_title_error(slide_chart_title::map_axis_support_error(error))
}

fn map_chart_title_error(error: ChartTitleError) -> SlideChartDataError {
    match error {
        ChartTitleError::UnsupportedSource => SlideChartDataError::UnsupportedSource,
        ChartTitleError::AmbiguousSelector => SlideChartDataError::AmbiguousSelector,
        ChartTitleError::EmptySlideName => SlideChartDataError::EmptySlideName,
        ChartTitleError::SlideNameNotFound => SlideChartDataError::SlideNameNotFound,
        ChartTitleError::SlidePositionNotFound { position } => {
            SlideChartDataError::SlidePositionNotFound { position }
        },
        ChartTitleError::ChartNameNotFound => SlideChartDataError::ChartNameNotFound,
        ChartTitleError::ChartPositionNotFound { position } => {
            SlideChartDataError::ChartPositionNotFound { position }
        },
        ChartTitleError::EmptyChartName => SlideChartDataError::EmptyChartName,
        ChartTitleError::InvalidSource
        | ChartTitleError::Verification
        | ChartTitleError::PatchConflict => SlideChartDataError::InvalidSource,
        ChartTitleError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => SlideChartDataError::LimitExceeded {
            kind: map_title_limit_kind(kind),
            observed,
            maximum,
        },
        ChartTitleError::Allocation { amount } => SlideChartDataError::Allocation { amount },
    }
}

fn map_title_limit_kind(kind: ChartTitleLimitKind) -> SlideChartDataLimitKind {
    match kind {
        ChartTitleLimitKind::InputBytes => SlideChartDataLimitKind::InputBytes,
        ChartTitleLimitKind::OutputBytes => SlideChartDataLimitKind::WireBytes,
        ChartTitleLimitKind::WireBytes => SlideChartDataLimitKind::WireBytes,
        ChartTitleLimitKind::Entries => SlideChartDataLimitKind::Entries,
        ChartTitleLimitKind::EntryBytes => SlideChartDataLimitKind::EntryBytes,
        ChartTitleLimitKind::TotalBytes => SlideChartDataLimitKind::TotalBytes,
        ChartTitleLimitKind::Slides => SlideChartDataLimitKind::Slides,
        ChartTitleLimitKind::References => SlideChartDataLimitKind::References,
        ChartTitleLimitKind::TextStorages => SlideChartDataLimitKind::TextStorages,
        ChartTitleLimitKind::TextFragments => SlideChartDataLimitKind::TextFragments,
        ChartTitleLimitKind::TextBytes | ChartTitleLimitKind::TitleBytes => {
            SlideChartDataLimitKind::TextBytes
        },
        ChartTitleLimitKind::WireFields => SlideChartDataLimitKind::WireFields,
        ChartTitleLimitKind::WireNesting => SlideChartDataLimitKind::WireNesting,
        ChartTitleLimitKind::WireWork => SlideChartDataLimitKind::WireWork,
    }
}

fn map_chart_data_decode_error(error: chart_data_codec::DecodeError) -> SlideChartDataError {
    if let Some((observed, maximum)) = error.input_limit_values() {
        return SlideChartDataError::LimitExceeded {
            kind: SlideChartDataLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideChartDataError::LimitExceeded {
            kind: SlideChartDataLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideChartDataError::LimitExceeded {
            kind: SlideChartDataLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.cell_limit_values() {
        return SlideChartDataError::LimitExceeded {
            kind: SlideChartDataLimitKind::Cells,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.label_limit_values() {
        return SlideChartDataError::LimitExceeded {
            kind: SlideChartDataLimitKind::LabelCount,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.text_limit_values() {
        return SlideChartDataError::LimitExceeded {
            kind: SlideChartDataLimitKind::TextBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.depth_limit_values() {
        return SlideChartDataError::LimitExceeded {
            kind: SlideChartDataLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        };
    }
    SlideChartDataError::InvalidSource
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn semantic_copy_budget_refuses_retained_bytes_before_allocation() {
        let source = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/keynote/chart-caption-native.key"),
        )
        .expect("native chart fixture");
        let package = Package::from_bytes(&source).expect("native package");
        let mut budget = ChartGraphScanBudget::new(&package).expect("chart budget");
        let (_, remaining) = budget.chart_metadata_residual_limits();
        assert!(remaining > 1);
        budget
            .charge_work(remaining - 1)
            .expect("reserve all but one work unit");

        assert_eq!(
            preflight_data_budget(&budget, SlideChartDataLimitKind::RetainedBytes, 2),
            Err(SlideChartDataError::LimitExceeded {
                kind: SlideChartDataLimitKind::RetainedBytes,
                observed: 2,
                maximum: 1,
            })
        );
    }

    #[test]
    fn semantic_projection_budget_refuses_a_low_work_ceiling() {
        let source = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/keynote/chart-caption-native.key"),
        )
        .expect("native chart fixture");
        let package = Package::from_bytes(&source).expect("native package");
        let grid = one_cell_grid();
        let (snapshot, _report) =
            chart_data_codec::decode_grid_with_report(&grid, &DecodeOptions::for_source(&grid))
                .expect("one-cell chart grid");
        let mut budget = ChartGraphScanBudget::new(&package).expect("chart budget");
        let (_, remaining) = budget.chart_metadata_residual_limits();
        let projection = snapshot
            .grid_source()
            .len()
            .checked_mul(10)
            .and_then(|value| value.checked_add(1))
            .expect("projection work");
        assert!(remaining > projection);
        budget
            .charge_work(remaining - projection + 1)
            .expect("leave less than the projection envelope");

        assert!(matches!(
            charge_snapshot_projection(snapshot, &mut budget),
            Err(SlideChartDataError::LimitExceeded {
                kind: SlideChartDataLimitKind::WireWork,
                observed,
                maximum,
            }) if observed == projection as u64 && maximum == (projection - 1) as u64
        ));
    }

    fn one_cell_grid() -> Vec<u8> {
        fn varint(mut value: u32, output: &mut Vec<u8>) {
            while value >= 0x80 {
                output.push((value as u8) | 0x80);
                value >>= 7;
            }
            output.push(value as u8);
        }

        fn text_field(number: u32, value: &[u8], output: &mut Vec<u8>) {
            varint((number << 3) | 2, output);
            varint(value.len() as u32, output);
            output.extend_from_slice(value);
        }

        let mut value = Vec::new();
        varint(9, &mut value);
        value.extend_from_slice(&1.0_f64.to_le_bytes());

        let mut row = Vec::new();
        text_field(1, &value, &mut row);

        let mut grid = Vec::new();
        text_field(1, b"Row", &mut grid);
        text_field(2, b"Value", &mut grid);
        text_field(3, &row, &mut grid);
        grid
    }
}
