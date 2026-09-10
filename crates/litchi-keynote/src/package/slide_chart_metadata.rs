//! Selector-first, archive-free Keynote chart metadata reads.
//!
//! The chart graph is resolved by the existing title owner.  This adapter
//! only projects the selected native chart payload into the common chart
//! metadata value; native object identifiers and protobuf values stay inside
//! this module.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "The package boundary maps native failures to content-redacted semantic errors."
)]

use std::fmt;

use litchi_core::Position;
use litchi_iwa_common::chart::{kind::Kind, metadata::ChartMetadata};
use litchi_iwa_protos::chart_metadata_codec::{self, DecodeOptions, LabelList};
use thiserror::Error;

use super::Package;
use super::chart_axis_support::{self, AxisSupportBudget, AxisSupportError};
use super::slide_chart_title::{
    self, ChartGraphScanBudget, ChartSelection, ChartTitleError, ChartTitleLimitKind,
    select_chart_with_budget,
};
use crate::{ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = chart_metadata_codec::MODERN_CHART_DRAWABLE_MESSAGE_TYPE;
const MAX_METADATA_LABEL_COUNT: usize = 1_000_000;
const MAX_METADATA_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// A finite resource governed while one chart metadata view is prepared.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideChartMetadataLimitKind {
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
    /// Allocations needed by the metadata view.
    Allocations,
    /// Bytes retained by the metadata view.
    RetainedBytes,
    /// Number of borrowed row/column labels.
    LabelCount,
}

impl fmt::Display for SlideChartMetadataLimitKind {
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

/// A content-redacted failure raised by a chart metadata read.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideChartMetadataError {
    /// The source was not retained as an exact physical package.
    #[error("this Keynote source does not support physical chart metadata reads")]
    UnsupportedSource,
    /// An exact-name selector was ambiguous.
    #[error("the Keynote chart metadata selector is ambiguous")]
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
    /// The selected chart graph or metadata payload was malformed.
    #[error("the Keynote chart metadata source is malformed or unsupported")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote chart metadata {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: SlideChartMetadataLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded semantic allocation failed before the result was published.
    #[error("could not allocate {amount} units for Keynote chart metadata")]
    Allocation { amount: usize },
}

impl Package {
    /// Read the archive-free metadata of one selected Keynote chart.
    pub fn slide_chart_metadata<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartMetadata, SlideChartMetadataError> {
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
        read_selected_metadata(self, selection, &mut budget)
    }
}

fn read_selected_metadata(
    package: &Package,
    selection: ChartSelection,
    budget: &mut ChartGraphScanBudget,
) -> Result<ChartMetadata, SlideChartMetadataError> {
    let (component_name, chart_object) = package
        .object_with_component(selection.chart_identifier)
        .ok_or(SlideChartMetadataError::InvalidSource)?;
    if component_name != selection.slide_component_name {
        return Err(SlideChartMetadataError::InvalidSource);
    }
    let (message_index, message) =
        chart_axis_support::unique_message(chart_object, CHART_MESSAGE_TYPE, budget)
            .map_err(map_axis_support_error)?;
    chart_axis_support::validate_selected_message_metadata(chart_object, message_index)
        .map_err(map_axis_support_error)?;

    let (limits, residual_work) = budget.chart_metadata_residual_limits();
    if residual_work == 0 {
        return Err(SlideChartMetadataError::LimitExceeded {
            kind: SlideChartMetadataLimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let maximum_depth = u32::try_from(limits.max_nesting()).map_err(|_error| {
        SlideChartMetadataError::LimitExceeded {
            kind: SlideChartMetadataLimitKind::WireNesting,
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
            .clamp(1, MAX_METADATA_LABEL_COUNT),
        limits
            .max_input_bytes()
            .min(residual_work)
            .clamp(1, MAX_METADATA_TEXT_BYTES),
    );
    let (snapshot, report) =
        chart_metadata_codec::decode_modern_with_report(message.data.as_slice(), &options)
            .map_err(|error| {
                let report = error.report();
                let charge = budget.charge_chart_metadata_report(report);
                match charge {
                    Ok(()) => map_chart_metadata_decode_error(error),
                    Err(budget_error) => map_chart_title_error(budget_error),
                }
            })?;
    budget
        .charge_chart_metadata_report(report)
        .map_err(map_chart_title_error)?;
    if snapshot.non_style_ref().map(|reference| reference.get())
        != Some(selection.non_style_identifier)
    {
        return Err(SlideChartMetadataError::InvalidSource);
    }

    let row_names = copy_labels(snapshot.row_labels(), budget)?;
    let column_names = copy_labels(snapshot.column_labels(), budget)?;
    Ok(ChartMetadata::from_owned(
        Kind::from_native(snapshot.chart_type()),
        selection.title,
        row_names,
        column_names,
        snapshot.series_count(),
        snapshot.contains_default_data().unwrap_or(false),
    ))
}

fn copy_labels<'source>(
    labels: LabelList<'source>,
    budget: &mut ChartGraphScanBudget,
) -> Result<Vec<String>, SlideChartMetadataError> {
    let count = labels.len();
    let mut text_bytes = 0usize;
    let mut string_allocations = 0usize;
    let mut seen = 0usize;
    for label in labels.iter() {
        seen = seen
            .checked_add(1)
            .ok_or(SlideChartMetadataError::InvalidSource)?;
        text_bytes = text_bytes
            .checked_add(label.len())
            .ok_or(SlideChartMetadataError::InvalidSource)?;
        string_allocations = string_allocations
            .checked_add(usize::from(!label.is_empty()))
            .ok_or(SlideChartMetadataError::InvalidSource)?;
    }
    if seen != count {
        return Err(SlideChartMetadataError::InvalidSource);
    }
    let charge = count
        .checked_add(text_bytes)
        .and_then(|value| value.checked_add(string_allocations))
        .ok_or(SlideChartMetadataError::InvalidSource)?;
    budget.charge_work(charge).map_err(map_axis_support_error)?;

    let mut copied = Vec::new();
    copied
        .try_reserve_exact(count)
        .map_err(|_error| SlideChartMetadataError::Allocation { amount: count })?;
    for label in labels.iter() {
        let mut owned = String::new();
        owned.try_reserve_exact(label.len()).map_err(|_error| {
            SlideChartMetadataError::Allocation {
                amount: label.len(),
            }
        })?;
        owned.push_str(label);
        copied.push(owned);
    }
    if copied.len() != count {
        return Err(SlideChartMetadataError::InvalidSource);
    }
    Ok(copied)
}

fn map_axis_support_error(error: AxisSupportError) -> SlideChartMetadataError {
    map_chart_title_error(slide_chart_title::map_axis_support_error(error))
}

fn map_chart_title_error(error: ChartTitleError) -> SlideChartMetadataError {
    match error {
        ChartTitleError::UnsupportedSource => SlideChartMetadataError::UnsupportedSource,
        ChartTitleError::AmbiguousSelector => SlideChartMetadataError::AmbiguousSelector,
        ChartTitleError::EmptySlideName => SlideChartMetadataError::EmptySlideName,
        ChartTitleError::SlideNameNotFound => SlideChartMetadataError::SlideNameNotFound,
        ChartTitleError::SlidePositionNotFound { position } => {
            SlideChartMetadataError::SlidePositionNotFound { position }
        },
        ChartTitleError::ChartNameNotFound => SlideChartMetadataError::ChartNameNotFound,
        ChartTitleError::ChartPositionNotFound { position } => {
            SlideChartMetadataError::ChartPositionNotFound { position }
        },
        ChartTitleError::EmptyChartName => SlideChartMetadataError::EmptyChartName,
        ChartTitleError::InvalidSource
        | ChartTitleError::Verification
        | ChartTitleError::PatchConflict => SlideChartMetadataError::InvalidSource,
        ChartTitleError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => SlideChartMetadataError::LimitExceeded {
            kind: map_title_limit_kind(kind),
            observed,
            maximum,
        },
        ChartTitleError::Allocation { amount } => SlideChartMetadataError::Allocation { amount },
    }
}

fn map_title_limit_kind(kind: ChartTitleLimitKind) -> SlideChartMetadataLimitKind {
    match kind {
        ChartTitleLimitKind::InputBytes => SlideChartMetadataLimitKind::InputBytes,
        ChartTitleLimitKind::OutputBytes => SlideChartMetadataLimitKind::WireBytes,
        ChartTitleLimitKind::WireBytes => SlideChartMetadataLimitKind::WireBytes,
        ChartTitleLimitKind::Entries => SlideChartMetadataLimitKind::Entries,
        ChartTitleLimitKind::EntryBytes => SlideChartMetadataLimitKind::EntryBytes,
        ChartTitleLimitKind::TotalBytes => SlideChartMetadataLimitKind::TotalBytes,
        ChartTitleLimitKind::Slides => SlideChartMetadataLimitKind::Slides,
        ChartTitleLimitKind::References => SlideChartMetadataLimitKind::References,
        ChartTitleLimitKind::TextStorages => SlideChartMetadataLimitKind::TextStorages,
        ChartTitleLimitKind::TextFragments => SlideChartMetadataLimitKind::TextFragments,
        ChartTitleLimitKind::TextBytes | ChartTitleLimitKind::TitleBytes => {
            SlideChartMetadataLimitKind::TextBytes
        },
        ChartTitleLimitKind::WireFields => SlideChartMetadataLimitKind::WireFields,
        ChartTitleLimitKind::WireNesting => SlideChartMetadataLimitKind::WireNesting,
        ChartTitleLimitKind::WireWork => SlideChartMetadataLimitKind::WireWork,
    }
}

fn map_chart_metadata_decode_error(
    error: chart_metadata_codec::DecodeError,
) -> SlideChartMetadataError {
    if let Some(amount) = error.allocation_amount() {
        return SlideChartMetadataError::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.input_limit_values() {
        return SlideChartMetadataError::LimitExceeded {
            kind: SlideChartMetadataLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideChartMetadataError::LimitExceeded {
            kind: SlideChartMetadataLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideChartMetadataError::LimitExceeded {
            kind: SlideChartMetadataLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.label_limit_values() {
        return SlideChartMetadataError::LimitExceeded {
            kind: SlideChartMetadataLimitKind::LabelCount,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.text_limit_values() {
        return SlideChartMetadataError::LimitExceeded {
            kind: SlideChartMetadataLimitKind::TextBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.depth_limit_values() {
        return SlideChartMetadataError::LimitExceeded {
            kind: SlideChartMetadataLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        };
    }
    SlideChartMetadataError::InvalidSource
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
