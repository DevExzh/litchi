//! Selector-first Pages body-chart data reads.
//!
//! The Arrange reader proves the rooted body-chart ownership path before this
//! module asks the shared chart-data codec to inspect the selected drawable.
//! Native identifiers and protobuf objects therefore stay private to the
//! package boundary.  The returned value is the common archive-free
//! [`litchi_iwa_common::chart::data::ChartData`] grid.

use std::fmt;
use std::mem::size_of;

use litchi_core::Position;
use litchi_iwa_common::chart::data::{ChartData, DataError};
use litchi_iwa_protos::chart_data_codec::{self, DecodeLimit, DecodeOptions};
use thiserror::Error;

use super::Package;
use super::body_chart_arrangement::{
    self, ArrangementBudget, BodyChartArrangementError, BodyChartArrangementLimitKind,
};
use crate::selector::BodyChartSelector;

const MAX_DATA_CELLS: usize = 1_000_000;
const MAX_DATA_LABEL_COUNT: usize = 1_000_000;
const MAX_DATA_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// Finite resources governed by one Pages body-chart data read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyChartDataLimitKind {
    /// Complete package input bytes inspected by rooted selection.
    InputBytes,
    /// Native payload bytes inspected.
    PayloadBytes,
    /// Native references inspected.
    PayloadReferences,
    /// Strict chart-data wire bytes.
    WireBytes,
    /// Strict chart-data wire fields.
    WireFields,
    /// Strict chart-data wire nesting.
    WireNesting,
    /// Strict chart-data wire work.
    WireWork,
    /// Borrowed decoder allocations.
    WireAllocations,
    /// Borrowed decoder retained bytes.
    WireRetainedBytes,
    /// Numeric cell count.
    CellCount,
    /// Number of row and column labels.
    LabelCount,
    /// UTF-8 label bytes.
    TextBytes,
}

impl fmt::Display for BodyChartDataLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::WireAllocations => "wire allocations",
            Self::WireRetainedBytes => "wire retained bytes",
            Self::CellCount => "cell count",
            Self::LabelCount => "label count",
            Self::TextBytes => "text bytes",
        })
    }
}

/// Failure while reading one rooted Pages body chart's semantic data grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum BodyChartDataError {
    /// No rooted body chart matched the source-order selector.
    #[error("the Pages body has no chart at position {position:?}")]
    ChartNotFound { position: Position },
    /// The source does not retain a supported exact native artifact.
    #[error("this Pages source does not support body-chart data reads")]
    UnsupportedSource,
    /// The rooted graph or selected chart-data payload was malformed.
    #[error("the selected Pages body-chart data source is invalid")]
    InvalidSource,
    /// A finite read resource ceiling was exceeded.
    #[error("Pages body-chart data {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: BodyChartDataLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded semantic allocation failed.
    #[error("could not allocate {amount} units for Pages body-chart data")]
    Allocation { amount: usize },
}

impl Package {
    /// Read one rooted body chart's modern inline numeric data grid.
    ///
    /// Root selection, strict chart-data decoding, and semantic materializing
    /// allocations consume one aggregate bounded budget.  The decoder keeps
    /// native labels and values borrowed until this method returns the
    /// archive-free common model. Date- and duration-only native cells are
    /// represented as missing values; legacy pre-Unity grids are refused.
    pub fn body_chart_data(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<ChartData, BodyChartDataError> {
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = body_chart_arrangement::resolve_target(self, selector.into(), &mut budget)
            .map_err(map_arrangement_error)?;
        let source = body_chart_arrangement::selected_chart_payload(self, &target)
            .map_err(map_arrangement_error)?;

        let options = data_options(&budget, source)?;
        let (snapshot, report) = match chart_data_codec::decode_modern_with_report(source, &options)
        {
            Ok(decoded) => decoded,
            Err(error) => {
                let report = error.report();
                charge_data_report(&mut budget, report)?;
                return Err(map_data_codec_error(error));
            },
        };
        charge_data_report(&mut budget, report)?;
        charge_materialization_work(&mut budget, snapshot, report)?;

        let row_names = own_labels(snapshot.row_labels(), &mut budget)?;
        let column_names = own_labels(snapshot.column_labels(), &mut budget)?;
        let values = own_values(snapshot.rows(), snapshot.column_count(), &mut budget)?;

        ChartData::new(row_names, column_names, values).map_err(map_data_error)
    }
}

fn data_options(
    budget: &ArrangementBudget,
    source: &[u8],
) -> Result<DecodeOptions, BodyChartDataError> {
    let limits = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    let source_bytes = source.len().max(1).min(limits.max_input_bytes());
    let fields = limits.max_fields().max(1);
    let work = limits.max_rewrite_work().max(1);
    let max_depth =
        u32::try_from(limits.max_nesting()).map_err(|_| BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireNesting,
            observed: limits.max_nesting() as u64,
            maximum: u32::MAX as u64,
        })?;
    let max_cells = limits.max_fields().clamp(1, MAX_DATA_CELLS);
    let max_labels = limits.max_fields().clamp(1, MAX_DATA_LABEL_COUNT);
    let max_text = limits.max_input_bytes().clamp(1, MAX_DATA_TEXT_BYTES);

    Ok(DecodeOptions::new(
        source_bytes,
        fields,
        work,
        max_depth,
        max_cells,
        max_labels,
        max_text,
    ))
}

fn charge_data_report(
    budget: &mut ArrangementBudget,
    report: chart_data_codec::DecodeReport,
) -> Result<(), BodyChartDataError> {
    budget
        .input(report.source_bytes())
        .map_err(map_arrangement_error)?;
    budget
        .fields(report.fields())
        .map_err(map_arrangement_error)?;
    budget
        .work(report.work_bytes())
        .map_err(map_arrangement_error)?;
    budget
        .allocations(report.allocations())
        .map_err(map_arrangement_error)?;
    budget
        .retained(report.retained_bytes())
        .map_err(map_arrangement_error)?;

    let limits = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    if usize::try_from(report.max_depth()).unwrap_or(usize::MAX) > limits.max_nesting() {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireNesting,
            observed: u64::from(report.max_depth()),
            maximum: limits.max_nesting() as u64,
        });
    }
    if report.cell_count() > MAX_DATA_CELLS {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::CellCount,
            observed: report.cell_count() as u64,
            maximum: MAX_DATA_CELLS as u64,
        });
    }
    if report.label_count() > MAX_DATA_LABEL_COUNT {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::LabelCount,
            observed: report.label_count() as u64,
            maximum: MAX_DATA_LABEL_COUNT as u64,
        });
    }
    if report.text_bytes() > MAX_DATA_TEXT_BYTES {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::TextBytes,
            observed: report.text_bytes() as u64,
            maximum: MAX_DATA_TEXT_BYTES as u64,
        });
    }
    Ok(())
}

fn charge_materialization_work(
    budget: &mut ArrangementBudget,
    snapshot: chart_data_codec::ChartDataSnapshot<'_>,
    report: chart_data_codec::DecodeReport,
) -> Result<(), BodyChartDataError> {
    // The borrowed iterators intentionally re-scan source spans after strict
    // decode: labels are walked once to preflight text/allocation shape and
    // once to copy, while rows and values are walked once to materialize the
    // common model. Reserve a source-sized envelope before any such walk so
    // the aggregate budget also bounds this lazy replay work.
    let work = snapshot
        .grid_source()
        .len()
        .checked_mul(10)
        .and_then(|amount| amount.checked_add(report.text_bytes()))
        .and_then(|amount| amount.checked_add(report.cell_count()))
        .ok_or(BodyChartDataError::InvalidSource)?;
    budget.work(work).map_err(map_arrangement_error)
}

fn own_labels(
    labels: chart_data_codec::LabelList<'_>,
    budget: &mut ArrangementBudget,
) -> Result<Vec<String>, BodyChartDataError> {
    let count = labels.len();
    if count > MAX_DATA_LABEL_COUNT {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::LabelCount,
            observed: count as u64,
            maximum: MAX_DATA_LABEL_COUNT as u64,
        });
    }
    let (text_bytes, string_allocations, seen) = labels.iter().try_fold(
        (0usize, 0usize, 0usize),
        |(text_bytes, string_allocations, seen), label| {
            let text_bytes = text_bytes
                .checked_add(label.len())
                .ok_or(BodyChartDataError::InvalidSource)?;
            let string_allocations = string_allocations
                .checked_add(usize::from(!label.is_empty()))
                .ok_or(BodyChartDataError::InvalidSource)?;
            let seen = seen
                .checked_add(1)
                .ok_or(BodyChartDataError::InvalidSource)?;
            Ok((text_bytes, string_allocations, seen))
        },
    )?;
    if seen != count {
        return Err(BodyChartDataError::InvalidSource);
    }
    if text_bytes > MAX_DATA_TEXT_BYTES {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::TextBytes,
            observed: text_bytes as u64,
            maximum: MAX_DATA_TEXT_BYTES as u64,
        });
    }
    let allocations = usize::from(count != 0)
        .checked_add(string_allocations)
        .ok_or(BodyChartDataError::InvalidSource)?;
    let retained = count
        .checked_mul(size_of::<String>())
        .and_then(|amount| amount.checked_add(text_bytes))
        .ok_or(BodyChartDataError::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;

    let mut owned = Vec::new();
    owned
        .try_reserve_exact(count)
        .map_err(|_| BodyChartDataError::Allocation { amount: count })?;
    for label in labels.iter() {
        let mut value = String::new();
        value
            .try_reserve_exact(label.len())
            .map_err(|_| BodyChartDataError::Allocation {
                amount: label.len(),
            })?;
        value.push_str(label);
        owned.push(value);
    }
    if owned.len() != count {
        return Err(BodyChartDataError::InvalidSource);
    }
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(owned)
}

fn own_values(
    rows: chart_data_codec::GridRows<'_>,
    columns: usize,
    budget: &mut ArrangementBudget,
) -> Result<Vec<Vec<Option<f64>>>, BodyChartDataError> {
    let row_count = rows.len();
    let cells = row_count
        .checked_mul(columns)
        .ok_or(BodyChartDataError::InvalidSource)?;
    if cells > MAX_DATA_CELLS {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::CellCount,
            observed: cells as u64,
            maximum: MAX_DATA_CELLS as u64,
        });
    }
    let allocations = usize::from(row_count != 0)
        .checked_add(row_count)
        .ok_or(BodyChartDataError::InvalidSource)?;
    let retained = row_count
        .checked_mul(size_of::<Vec<Option<f64>>>())
        .and_then(|amount| amount.checked_add(cells.checked_mul(size_of::<Option<f64>>())?))
        .ok_or(BodyChartDataError::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;

    let mut owned = Vec::new();
    owned
        .try_reserve_exact(row_count)
        .map_err(|_| BodyChartDataError::Allocation { amount: row_count })?;
    for row in rows.iter() {
        if row.len() != columns {
            return Err(BodyChartDataError::InvalidSource);
        }
        let mut values = Vec::new();
        values
            .try_reserve_exact(columns)
            .map_err(|_| BodyChartDataError::Allocation { amount: columns })?;
        for value in row.values() {
            values.push(value);
        }
        if values.len() != columns {
            return Err(BodyChartDataError::InvalidSource);
        }
        owned.push(values);
    }
    if owned.len() != row_count {
        return Err(BodyChartDataError::InvalidSource);
    }
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(owned)
}

fn map_arrangement_error(error: BodyChartArrangementError) -> BodyChartDataError {
    match error {
        BodyChartArrangementError::ChartNotFound { position }
        | BodyChartArrangementError::ChartPositionNotFound { position } => {
            BodyChartDataError::ChartNotFound { position }
        },
        BodyChartArrangementError::UnsupportedSource => BodyChartDataError::UnsupportedSource,
        BodyChartArrangementError::InvalidSource
        | BodyChartArrangementError::Verification
        | BodyChartArrangementError::PatchConflict => BodyChartDataError::InvalidSource,
        BodyChartArrangementError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyChartDataError::LimitExceeded {
            kind: map_arrangement_limit(kind),
            observed,
            maximum,
        },
        BodyChartArrangementError::Allocation { amount } => {
            BodyChartDataError::Allocation { amount }
        },
    }
}

fn map_arrangement_limit(kind: BodyChartArrangementLimitKind) -> BodyChartDataLimitKind {
    match kind {
        BodyChartArrangementLimitKind::InputBytes
        | BodyChartArrangementLimitKind::PayloadBytes
        | BodyChartArrangementLimitKind::TotalPayloadBytes => BodyChartDataLimitKind::InputBytes,
        BodyChartArrangementLimitKind::PayloadReferences => {
            BodyChartDataLimitKind::PayloadReferences
        },
        BodyChartArrangementLimitKind::WireBytes
        | BodyChartArrangementLimitKind::WireOutputBytes => BodyChartDataLimitKind::WireBytes,
        BodyChartArrangementLimitKind::WireFields
        | BodyChartArrangementLimitKind::PayloadItems
        | BodyChartArrangementLimitKind::PayloadMessages
        | BodyChartArrangementLimitKind::PayloadObjects => BodyChartDataLimitKind::WireFields,
        BodyChartArrangementLimitKind::WireNesting => BodyChartDataLimitKind::WireNesting,
        BodyChartArrangementLimitKind::WireWork => BodyChartDataLimitKind::WireWork,
        BodyChartArrangementLimitKind::WireAllocations => BodyChartDataLimitKind::WireAllocations,
        BodyChartArrangementLimitKind::WireRetainedBytes
        | BodyChartArrangementLimitKind::WireScratchBytes => {
            BodyChartDataLimitKind::WireRetainedBytes
        },
        BodyChartArrangementLimitKind::OutputBytes
        | BodyChartArrangementLimitKind::Entries
        | BodyChartArrangementLimitKind::EntryBytes
        | BodyChartArrangementLimitKind::TotalEntryBytes
        | BodyChartArrangementLimitKind::PackageBytes => BodyChartDataLimitKind::InputBytes,
    }
}

fn map_data_codec_error(error: chart_data_codec::DecodeError) -> BodyChartDataError {
    let Some(limit) = error.resource_limit() else {
        return BodyChartDataError::InvalidSource;
    };
    match limit {
        DecodeLimit::Bytes { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Fields { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Work { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Nesting { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        DecodeLimit::Cells { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::CellCount,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Labels { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::LabelCount,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Text { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::TextBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => BodyChartDataError::InvalidSource,
    }
}

fn map_data_error(_error: DataError) -> BodyChartDataError {
    BodyChartDataError::InvalidSource
}

#[cfg(test)]
mod tests {
    use super::*;

    const NATIVE: &[u8] =
        include_bytes!("../../../../test-data/iwork/pages/chart-data-native.pages");

    #[test]
    fn selected_payload_cell_limit_is_reported_before_materializing_values() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let mut budget = ArrangementBudget::new(&package).expect("arrangement budget");
        let target = body_chart_arrangement::resolve_target(
            &package,
            BodyChartSelector::index(0),
            &mut budget,
        )
        .expect("native chart target");
        let source = body_chart_arrangement::selected_chart_payload(&package, &target)
            .expect("selected chart payload");
        let options = data_options(&budget, source)
            .expect("data options")
            .with_max_cells(1);
        let error = chart_data_codec::decode_modern_with_report(source, &options)
            .expect_err("the two-by-four grid exceeds one cell");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Cells {
                observed: 2,
                maximum: 1
            })
        ));
    }

    #[test]
    fn postdecode_replay_work_is_refused_by_the_aggregate_budget() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let mut selection_budget = ArrangementBudget::new(&package).expect("arrangement budget");
        let target = body_chart_arrangement::resolve_target(
            &package,
            BodyChartSelector::index(0),
            &mut selection_budget,
        )
        .expect("native chart target");
        let source = body_chart_arrangement::selected_chart_payload(&package, &target)
            .expect("selected chart payload");
        let options = data_options(&selection_budget, source).expect("data options");
        let (snapshot, report) =
            chart_data_codec::decode_modern_with_report(source, &options).expect("chart data");

        let mut exhausted = ArrangementBudget::new(&package).expect("fresh arrangement budget");
        let BodyChartArrangementError::LimitExceeded { maximum, .. } =
            exhausted.work(usize::MAX).expect_err("finite work ceiling")
        else {
            panic!("expected a work limit");
        };
        exhausted
            .work(usize::try_from(maximum).expect("addressable work limit"))
            .expect("consume the work budget");
        let result = charge_materialization_work(&mut exhausted, snapshot, report);
        assert!(matches!(
            result,
            Err(BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireWork,
                ..
            })
        ));
    }

    #[test]
    fn value_copy_refuses_exhausted_retention_budget() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let mut selection_budget = ArrangementBudget::new(&package).expect("arrangement budget");
        let target = body_chart_arrangement::resolve_target(
            &package,
            BodyChartSelector::index(0),
            &mut selection_budget,
        )
        .expect("native chart target");
        let source = body_chart_arrangement::selected_chart_payload(&package, &target)
            .expect("selected chart payload");
        let options = data_options(&selection_budget, source).expect("data options");
        let snapshot = chart_data_codec::decode_modern(source, &options).expect("chart data");

        let mut budget = ArrangementBudget::new(&package).expect("fresh arrangement budget");
        budget
            .retained(
                usize::try_from(package.state.source.limits().max_input_bytes())
                    .expect("addressable retention limit"),
            )
            .expect("consume the retention budget");
        let result = own_values(snapshot.rows(), snapshot.column_count(), &mut budget);
        assert!(matches!(
            result,
            Err(BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireRetainedBytes,
                ..
            })
        ));
    }
}
