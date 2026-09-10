//! Selector-first, borrowed Numbers chart metadata reads.
//!
//! The sheet/chart graph is rooted by the existing chart-arrangement walker.
//! Only the selected drawable payload is handed to the strict metadata codec;
//! labels and the optional title are copied into the archive-free common
//! model after every source and allocation limit has been charged.

#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "The package boundary converts bounded native arithmetic and redacts graph failures."
)]

use std::fmt;
use std::mem::size_of;
use std::num::NonZeroU64;

use litchi_core::Position;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::metadata::ChartMetadata;
use litchi_iwa_protos::chart_metadata_codec as codec;
use litchi_iwa_protos::keynote_chart_title_codec as title_codec;
use thiserror::Error;

use super::Package;
use super::chart_arrangement::{
    self, ChartArrangementError, ChartBudget, ChartSelection, object_from_resolved,
    parse_wire_view_with_budget, select_chart_with_budget, unique_typed_message,
    validate_message_metadata,
};
use crate::{ChartSelector, SheetSelector};

const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const GENERATED_CHART_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;

/// Resource category reported by a focused Numbers chart-metadata read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartMetadataLimitKind {
    /// Source bytes inspected by the rooted graph or codec.
    WireBytes,
    /// Encoded fields visited by strict preflight or Buffa.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate strict projection work.
    WireWork,
    /// Borrowed labels retained by the projection.
    Labels,
    /// UTF-8 text retained by the owned common model.
    TextBytes,
    /// Logical temporary allocations.
    Allocations,
    /// Bytes retained by one read operation.
    RetainedBytes,
    /// Native references inspected by the selector.
    References,
}

impl fmt::Display for ChartMetadataLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Labels => "labels",
            Self::TextBytes => "text bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::References => "references",
        })
    }
}

/// Failure from a semantic Numbers chart-metadata read.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartMetadataError {
    /// The package does not retain a supported physical source.
    #[error("this Numbers source does not support chart metadata reads")]
    UnsupportedSource,
    /// The selected graph crosses an unsupported native dependency boundary.
    #[error("the requested Numbers chart-metadata graph is unsupported")]
    UnsupportedDependency,
    /// A semantic selector matched more than one native owner.
    #[error("the Numbers chart-metadata selector is ambiguous")]
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
    #[error("the Numbers chart-metadata source is invalid")]
    InvalidSource,
    /// A referenced title object belongs to another chart graph.
    #[error("the Numbers chart-metadata source has a foreign title reference")]
    ForeignReference,
    /// A finite operation budget was exceeded.
    #[error("Numbers chart metadata {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource that exceeded its ceiling.
        kind: ChartMetadataLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for Numbers chart metadata")]
    Allocation { amount: usize },
}

impl Package {
    /// Read one sheet chart's archive-free metadata through semantic
    /// selectors.
    pub fn sheet_chart_metadata<'sheet>(
        &self,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        chart_selector: impl Into<ChartSelector>,
    ) -> Result<ChartMetadata, ChartMetadataError> {
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement_error)?;
        let selection = select_chart_with_budget(
            self,
            sheet_selector.into(),
            chart_selector.into(),
            false,
            &mut budget,
        )
        .map_err(map_arrangement_error)?;
        let payload = chart_arrangement::selected_chart_payload(self, &selection)
            .map_err(map_arrangement_error)?;
        let options = budget
            .metadata_codec_options()
            .map_err(map_arrangement_error)?;
        let (snapshot, report) = match codec::decode_modern_with_report(payload, &options) {
            Ok(decoded) => decoded,
            Err(error) => {
                let report = error.report();
                if let Err(budget_error) = budget.metadata_codec_report(report) {
                    return Err(map_arrangement_error(budget_error));
                }
                if let Some(limit) = map_metadata_codec_limit(&error) {
                    return Err(limit);
                }
                return Err(ChartMetadataError::InvalidSource);
            },
        };
        budget
            .metadata_codec_report(report)
            .map_err(map_arrangement_error)?;

        let title = read_title(self, &selection, snapshot.non_style_ref(), &mut budget)?;
        let row_names = own_labels(snapshot.row_labels(), &mut budget)?;
        let column_names = own_labels(snapshot.column_labels(), &mut budget)?;
        let title = own_title(title, &mut budget)?;

        Ok(ChartMetadata::from_owned(
            Kind::from_native(snapshot.chart_type()),
            title,
            row_names,
            column_names,
            snapshot.series_count(),
            snapshot.contains_default_data().unwrap_or(false),
        ))
    }
}

fn read_title<'source>(
    package: &'source Package,
    selection: &ChartSelection,
    identifier: Option<NonZeroU64>,
    budget: &mut ChartBudget,
) -> Result<Option<&'source str>, ChartMetadataError> {
    let Some(identifier) = identifier else {
        return Ok(None);
    };
    let identifier = identifier.get();
    let resolved = package
        .state
        .index
        .resolve_ref_id(&package.state.components, identifier)
        .map_err(|_| ChartMetadataError::InvalidSource)?
        .ok_or(ChartMetadataError::InvalidSource)?;
    let chart_component = package
        .state
        .components
        .catalog()
        .get_index(selection.component_index)
        .ok_or(ChartMetadataError::InvalidSource)?;
    let chart_object = chart_component
        .archive()
        .objects
        .get(selection.object_index)
        .ok_or(ChartMetadataError::InvalidSource)?;
    let chart_info = chart_object
        .archive_info
        .message_infos
        .get(selection.message_index)
        .ok_or(ChartMetadataError::InvalidSource)?;
    if !chart_non_style_reference_is_owned(chart_info, identifier) {
        return Err(ChartMetadataError::ForeignReference);
    }
    let object = object_from_resolved(package, resolved).map_err(map_arrangement_error)?;
    if object.archive_info.identifier != Some(identifier) {
        return Err(ChartMetadataError::InvalidSource);
    }
    validate_message_metadata(object).map_err(map_arrangement_error)?;
    let Some((_message_index, message)) =
        unique_typed_message(object, CHART_NON_STYLE_MESSAGE_TYPE)
            .map_err(map_arrangement_error)?
    else {
        return Err(ChartMetadataError::InvalidSource);
    };
    let view = parse_wire_view_with_budget(message.data.as_slice(), 1, budget)
        .map_err(map_arrangement_error)?;
    let mut extension = None;
    for field in view.fields() {
        if field.number() != GENERATED_CHART_NON_STYLE_EXTENSION_FIELD {
            continue;
        }
        if extension.is_some() || field.wire_type() != 2 {
            return Err(ChartMetadataError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| ChartMetadataError::InvalidSource)?;
        extension = Some(field.payload());
    }
    let Some(extension) = extension else {
        return Ok(None);
    };
    let options = budget
        .title_codec_options(extension)
        .map_err(map_arrangement_error)?;
    let (snapshot, report) = match title_codec::decode_chart_title_with_report(extension, options) {
        Ok(decoded) => decoded,
        Err(error) => {
            budget
                .title_codec_failure(extension)
                .map_err(map_arrangement_error)?;
            if let Some(limit) = map_title_codec_limit(&error) {
                return Err(limit);
            }
            if let Some(amount) = error.allocation_amount() {
                return Err(ChartMetadataError::Allocation { amount });
            }
            return Err(ChartMetadataError::InvalidSource);
        },
    };
    budget
        .title_codec_report(report)
        .map_err(map_arrangement_error)?;
    Ok(snapshot.visible_title())
}

/// Prove that the selected chart message owns its non-style reference.
///
/// `MessageInfo::object_references` is the aggregate ownership index.  Field
/// metadata is optional in older archives, but when present it must identify
/// the generated chart extension (`10000`) and its `chart_non_style` field
/// (`10`).  A reference appearing only in a data list, an unrelated path, or
/// more than once is rejected before the object index is entered.
fn chart_non_style_reference_is_owned(
    info: &litchi_iwa_core::MessageInfo,
    identifier: u64,
) -> bool {
    let aggregate_count = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if aggregate_count != 1 || info.data_references.contains(&identifier) {
        return false;
    }
    let mut field_count = 0usize;
    for field in &info.field_infos {
        let object_count = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        let data_count = field
            .data_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if object_count == 0 && data_count == 0 {
            continue;
        }
        if data_count != 0 || object_count != 1 || field.path.as_slice() != [10_000, 10] {
            return false;
        }
        field_count = field_count.saturating_add(object_count);
    }
    field_count <= 1
}

fn own_labels(
    labels: codec::LabelList<'_>,
    budget: &mut ChartBudget,
) -> Result<Vec<String>, ChartMetadataError> {
    let count = labels.len();
    let text_bytes = labels
        .iter()
        .try_fold(0usize, |total, value| total.checked_add(value.len()))
        .ok_or(ChartMetadataError::InvalidSource)?;
    let retained = count
        .checked_mul(size_of::<String>())
        .and_then(|amount| amount.checked_add(text_bytes))
        .ok_or(ChartMetadataError::InvalidSource)?;
    let string_allocations = labels.iter().filter(|value| !value.is_empty()).count();
    let allocations = usize::from(count != 0)
        .checked_add(string_allocations)
        .ok_or(ChartMetadataError::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(count)
        .map_err(|_| ChartMetadataError::Allocation { amount: count })?;
    for value in labels.iter() {
        let mut owned = String::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|_| ChartMetadataError::Allocation {
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

fn own_title(
    title: Option<&str>,
    budget: &mut ChartBudget,
) -> Result<Option<String>, ChartMetadataError> {
    let Some(title) = title else {
        return Ok(None);
    };
    let retained = size_of::<String>()
        .checked_add(title.len())
        .ok_or(ChartMetadataError::InvalidSource)?;
    let allocations = usize::from(!title.is_empty());
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;
    let mut value = String::new();
    value
        .try_reserve_exact(title.len())
        .map_err(|_| ChartMetadataError::Allocation {
            amount: title.len(),
        })?;
    value.push_str(title);
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(Some(value))
}

fn map_arrangement_error(error: ChartArrangementError) -> ChartMetadataError {
    match error {
        ChartArrangementError::UnsupportedSource => ChartMetadataError::UnsupportedSource,
        ChartArrangementError::UnsupportedDependency => ChartMetadataError::UnsupportedDependency,
        ChartArrangementError::AmbiguousSelector => ChartMetadataError::AmbiguousSelector,
        ChartArrangementError::EmptySheetName => ChartMetadataError::EmptySheetName,
        ChartArrangementError::SheetNameNotFound => ChartMetadataError::SheetNameNotFound,
        ChartArrangementError::SheetPositionNotFound { position } => {
            ChartMetadataError::SheetPositionNotFound { position }
        },
        ChartArrangementError::ChartPositionNotFound { position } => {
            ChartMetadataError::ChartPositionNotFound { position }
        },
        ChartArrangementError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ChartMetadataError::LimitExceeded {
            kind: map_limit_kind(kind),
            observed,
            maximum,
        },
        ChartArrangementError::Allocation { amount } => ChartMetadataError::Allocation { amount },
        ChartArrangementError::InvalidSource
        | ChartArrangementError::Verification
        | ChartArrangementError::PatchConflict => ChartMetadataError::InvalidSource,
    }
}

fn map_metadata_codec_limit(error: &codec::DecodeError) -> Option<ChartMetadataError> {
    let (kind, observed, maximum) = match error.resource_limit()? {
        codec::DecodeLimit::Bytes { observed, maximum } => {
            (ChartMetadataLimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (ChartMetadataLimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::Work { observed, maximum } => {
            (ChartMetadataLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => (
            ChartMetadataLimitKind::WireNesting,
            observed as usize,
            maximum as usize,
        ),
        codec::DecodeLimit::Labels { observed, maximum } => {
            (ChartMetadataLimitKind::Labels, observed, maximum)
        },
        codec::DecodeLimit::Text { observed, maximum } => {
            (ChartMetadataLimitKind::TextBytes, observed, maximum)
        },
        _ => return None,
    };
    Some(ChartMetadataError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    })
}

fn map_title_codec_limit(error: &title_codec::DecodeError) -> Option<ChartMetadataError> {
    let (kind, observed, maximum) = match error.resource_limit()? {
        title_codec::DecodeLimit::Bytes { observed, maximum } => {
            (ChartMetadataLimitKind::WireBytes, observed, maximum)
        },
        title_codec::DecodeLimit::Fields { observed, maximum } => {
            (ChartMetadataLimitKind::WireFields, observed, maximum)
        },
        title_codec::DecodeLimit::Work { observed, maximum } => {
            (ChartMetadataLimitKind::WireWork, observed, maximum)
        },
        title_codec::DecodeLimit::Output { observed, maximum }
        | title_codec::DecodeLimit::Title { observed, maximum } => {
            (ChartMetadataLimitKind::TextBytes, observed, maximum)
        },
        title_codec::DecodeLimit::Nesting { observed, maximum } => (
            ChartMetadataLimitKind::WireNesting,
            observed as usize,
            maximum as usize,
        ),
        _ => return None,
    };
    Some(ChartMetadataError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    })
}

const fn map_limit_kind(
    kind: chart_arrangement::ChartArrangementLimitKind,
) -> ChartMetadataLimitKind {
    match kind {
        chart_arrangement::ChartArrangementLimitKind::WireBytes
        | chart_arrangement::ChartArrangementLimitKind::OutputBytes => {
            ChartMetadataLimitKind::WireBytes
        },
        chart_arrangement::ChartArrangementLimitKind::WireFields => {
            ChartMetadataLimitKind::WireFields
        },
        chart_arrangement::ChartArrangementLimitKind::WireNesting => {
            ChartMetadataLimitKind::WireNesting
        },
        chart_arrangement::ChartArrangementLimitKind::WireWork
        | chart_arrangement::ChartArrangementLimitKind::ScratchBytes => {
            ChartMetadataLimitKind::WireWork
        },
        chart_arrangement::ChartArrangementLimitKind::Allocations => {
            ChartMetadataLimitKind::Allocations
        },
        chart_arrangement::ChartArrangementLimitKind::RetainedBytes => {
            ChartMetadataLimitKind::RetainedBytes
        },
        chart_arrangement::ChartArrangementLimitKind::References => {
            ChartMetadataLimitKind::References
        },
    }
}
