//! Selector-first Pages body-chart metadata reads.
//!
//! The rooted chart graph is resolved by the Arrange reader.  This module
//! only decodes the selected modern chart payload and transfers its borrowed
//! labels and optional visible title into the archive-free common model.
//! Native identifiers remain private to the package boundary.

use std::fmt;
use std::mem::size_of;
use std::num::NonZeroU64;

use litchi_core::Position;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::metadata::ChartMetadata;
use litchi_iwa_common::wire::WireFieldView;
use litchi_iwa_protos::chart_metadata_codec::{self, DecodeOptions};
use litchi_iwa_protos::keynote_chart_title_codec::{
    self as chart_title_codec, DecodeLimit as ChartTitleDecodeLimit,
};
use thiserror::Error;

use super::body_chart_arrangement::{
    self, ArrangementBudget, BodyChartArrangementError, BodyChartArrangementLimitKind,
};
use super::{Package, body_chart_arrangement::ChartTarget};
use crate::selector::BodyChartSelector;

const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const GENERATED_CHART_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;
const MAX_METADATA_LABEL_COUNT: usize = 1_000_000;
const MAX_METADATA_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// Finite resources governed by one Pages body-chart metadata read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyChartMetadataLimitKind {
    /// Complete package input bytes inspected by rooted selection.
    InputBytes,
    /// Native payload bytes inspected.
    PayloadBytes,
    /// Native references inspected.
    PayloadReferences,
    /// Strict metadata wire bytes.
    WireBytes,
    /// Strict metadata wire fields.
    WireFields,
    /// Strict metadata wire nesting.
    WireNesting,
    /// Strict metadata wire work.
    WireWork,
    /// Borrowed decoder allocations.
    WireAllocations,
    /// Borrowed decoder retained bytes.
    WireRetainedBytes,
    /// Number of row and column labels retained.
    LabelCount,
    /// UTF-8 metadata text retained.
    TextBytes,
}

impl fmt::Display for BodyChartMetadataLimitKind {
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
            Self::LabelCount => "label count",
            Self::TextBytes => "text bytes",
        })
    }
}

/// Failure while reading one rooted Pages body chart's semantic metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum BodyChartMetadataError {
    /// No rooted body chart matched the source-order selector.
    #[error("the Pages body has no chart at position {position:?}")]
    ChartNotFound { position: Position },
    /// The source does not retain a supported exact native artifact.
    #[error("this Pages source does not support body-chart metadata reads")]
    UnsupportedSource,
    /// The rooted graph or selected metadata payload was malformed.
    #[error("the selected Pages body-chart metadata source is invalid")]
    InvalidSource,
    /// The selected chart points at a nonstyle object outside its component
    /// or without the declared generated-extension ownership path.
    #[error("the selected Pages body-chart metadata has a foreign title reference")]
    ForeignReference,
    /// A finite read resource ceiling was exceeded.
    #[error(
        "Pages body-chart metadata {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: BodyChartMetadataLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded semantic allocation failed.
    #[error("could not allocate {amount} units for Pages body-chart metadata")]
    Allocation { amount: usize },
}

impl Package {
    /// Read one rooted body chart's format-neutral metadata.
    ///
    /// Root selection and modern chart decoding consume one aggregate bounded
    /// budget.  The selected payload remains borrowed through strict decode;
    /// only the returned semantic strings are materialized.
    pub fn body_chart_metadata(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<ChartMetadata, BodyChartMetadataError> {
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = body_chart_arrangement::resolve_target(self, selector.into(), &mut budget)
            .map_err(map_arrangement_error)?;
        let source = body_chart_arrangement::selected_chart_payload(self, &target)
            .map_err(map_arrangement_error)?;

        let options = metadata_options(&budget, source)?;
        let (snapshot, report) =
            match chart_metadata_codec::decode_modern_with_report(source, &options) {
                Ok(decoded) => decoded,
                Err(error) => {
                    let failed = error.report();
                    charge_metadata_report(&mut budget, failed, true)?;
                    return Err(map_metadata_codec_error(error));
                },
            };
        charge_metadata_report(&mut budget, report, false)?;

        let title = snapshot
            .non_style_ref()
            .map(|reference| read_visible_title(self, &target, reference, &mut budget))
            .transpose()?
            .flatten();
        let row_names = own_labels(snapshot.row_labels(), &mut budget)?;
        let column_names = own_labels(snapshot.column_labels(), &mut budget)?;

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

fn metadata_options(
    budget: &ArrangementBudget,
    source: &[u8],
) -> Result<DecodeOptions, BodyChartMetadataError> {
    let limits = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    let source_bytes = source.len().max(1).min(limits.max_input_bytes());
    let fields = limits.max_fields().max(1);
    let work = limits.max_rewrite_work().max(1);
    let max_depth =
        u32::try_from(limits.max_nesting()).map_err(|_| BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::WireNesting,
            observed: limits.max_nesting() as u64,
            maximum: u32::MAX as u64,
        })?;
    let max_labels = limits.max_fields().clamp(1, MAX_METADATA_LABEL_COUNT);
    let max_text = limits.max_input_bytes().clamp(1, MAX_METADATA_TEXT_BYTES);

    Ok(DecodeOptions::new(
        source_bytes,
        fields,
        work,
        max_depth,
        max_labels,
        max_text,
    ))
}

fn charge_metadata_report(
    budget: &mut ArrangementBudget,
    report: chart_metadata_codec::DecodeReport,
    failed: bool,
) -> Result<(), BodyChartMetadataError> {
    budget
        .input(report.source_bytes())
        .map_err(map_arrangement_error)?;
    budget
        .fields(report.fields())
        .map_err(map_arrangement_error)?;
    let work = report
        .work_bytes()
        .checked_add(if failed {
            report.failure_work_bytes()
        } else {
            0
        })
        .ok_or(BodyChartMetadataError::InvalidSource)?;
    budget.work(work).map_err(map_arrangement_error)?;
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
        return Err(BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::WireNesting,
            observed: u64::from(report.max_depth()),
            maximum: limits.max_nesting() as u64,
        });
    }
    if report.label_count() > MAX_METADATA_LABEL_COUNT {
        return Err(BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::LabelCount,
            observed: report.label_count() as u64,
            maximum: MAX_METADATA_LABEL_COUNT as u64,
        });
    }
    if report.text_bytes() > MAX_METADATA_TEXT_BYTES {
        return Err(BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::TextBytes,
            observed: report.text_bytes() as u64,
            maximum: MAX_METADATA_TEXT_BYTES as u64,
        });
    }
    Ok(())
}

fn own_labels(
    labels: chart_metadata_codec::LabelList<'_>,
    budget: &mut ArrangementBudget,
) -> Result<Vec<String>, BodyChartMetadataError> {
    let count = labels.len();
    if count > MAX_METADATA_LABEL_COUNT {
        return Err(BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::LabelCount,
            observed: count as u64,
            maximum: MAX_METADATA_LABEL_COUNT as u64,
        });
    }
    let text_bytes = labels.iter().try_fold(0usize, |total, label| {
        total
            .checked_add(label.len())
            .ok_or(BodyChartMetadataError::InvalidSource)
    })?;
    let string_allocations = labels.iter().filter(|label| !label.is_empty()).count();
    let allocations = usize::from(count != 0)
        .checked_add(string_allocations)
        .ok_or(BodyChartMetadataError::InvalidSource)?;
    let retained = count
        .checked_mul(size_of::<String>())
        .and_then(|amount| amount.checked_add(text_bytes))
        .ok_or(BodyChartMetadataError::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(count)
        .map_err(|_| BodyChartMetadataError::Allocation { amount: count })?;
    for label in labels.iter() {
        let mut value = String::new();
        value
            .try_reserve_exact(label.len())
            .map_err(|_| BodyChartMetadataError::Allocation {
                amount: label.len(),
            })?;
        value.push_str(label);
        owned.push(value);
    }
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(owned)
}

fn read_visible_title(
    package: &Package,
    target: &ChartTarget,
    reference: NonZeroU64,
    budget: &mut ArrangementBudget,
) -> Result<Option<String>, BodyChartMetadataError> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.component_index)
        .ok_or(BodyChartMetadataError::InvalidSource)?;
    let chart = component
        .archive()
        .objects
        .get(target.drawable_object_index)
        .ok_or(BodyChartMetadataError::InvalidSource)?;
    if chart.archive_info.identifier != Some(target.drawable_identifier.get()) {
        return Err(BodyChartMetadataError::InvalidSource);
    }
    if !body_chart_arrangement::object_metadata_is_owned(
        chart,
        target.drawable_message_index,
        reference,
        &[10_000, 10],
        true,
    ) {
        return Err(BodyChartMetadataError::ForeignReference);
    }
    let located = body_chart_arrangement::locate_unique_object(package, reference, budget)
        .map_err(map_arrangement_error)?;
    let Some((message_index, message)) = body_chart_arrangement::unique_optional_message(
        located.object,
        CHART_NON_STYLE_MESSAGE_TYPE,
    )
    .map_err(map_arrangement_error)?
    else {
        return Err(BodyChartMetadataError::InvalidSource);
    };
    body_chart_arrangement::validate_message_metadata(located.object, message_index)
        .map_err(map_arrangement_error)?;

    let view = budget
        .parse(message.data.as_slice(), 1)
        .map_err(map_arrangement_error)?;
    let mut extension = None;
    let mut count = 0usize;
    for field in view
        .fields()
        .filter(|field| field.number() == GENERATED_CHART_NON_STYLE_EXTENSION_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(BodyChartMetadataError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| BodyChartMetadataError::InvalidSource)?;
        count = count
            .checked_add(1)
            .ok_or(BodyChartMetadataError::InvalidSource)?;
        extension = Some(field);
    }
    let Some(field) = extension else {
        return Ok(None);
    };
    if count != 1 {
        return Err(BodyChartMetadataError::InvalidSource);
    }
    decode_title(field, budget)
}

fn decode_title(
    field: WireFieldView<'_>,
    budget: &mut ArrangementBudget,
) -> Result<Option<String>, BodyChartMetadataError> {
    let source = field.payload();
    let limits = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::WireNesting,
            observed: limits.max_nesting() as u64,
            maximum: u32::MAX as u64,
        })?;
    let options = chart_title_codec::DecodeOptions::new(
        source.len().max(1).min(limits.max_input_bytes()),
        limits.max_fields().max(1),
        limits.max_rewrite_work().max(1),
        recursion,
    )
    .with_max_output_bytes(limits.max_input_bytes().max(1))
    .with_max_title_bytes(limits.max_input_bytes().max(1));
    // The title codec predates the public failure-report API used by the
    // chart metadata codec. Reserve a bounded source-sized envelope before
    // entering it so malformed titles still consume aggregate read budget.
    // Build options from the pre-reservation residual profile; rebuilding
    // them after reservation would unnecessarily constrain an otherwise valid
    // title twice.
    let worst_fields = source
        .len()
        .checked_mul(4)
        .ok_or(BodyChartMetadataError::InvalidSource)?
        .max(1);
    let worst_work = source
        .len()
        .checked_mul(8)
        .ok_or(BodyChartMetadataError::InvalidSource)?
        .max(1);
    budget.input(source.len()).map_err(map_arrangement_error)?;
    budget.fields(worst_fields).map_err(map_arrangement_error)?;
    budget.work(worst_work).map_err(map_arrangement_error)?;
    let snapshot =
        chart_title_codec::decode_chart_title(source, options).map_err(map_title_codec_error)?;
    let title = snapshot.visible_title();
    let Some(title) = title else {
        return Ok(None);
    };
    let allocations = usize::from(!title.is_empty());
    let retained = size_of::<String>()
        .checked_add(title.len())
        .ok_or(BodyChartMetadataError::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(title.len())
        .map_err(|_| BodyChartMetadataError::Allocation {
            amount: title.len(),
        })?;
    owned.push_str(title);
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(Some(owned))
}

fn map_arrangement_error(error: BodyChartArrangementError) -> BodyChartMetadataError {
    match error {
        BodyChartArrangementError::ChartNotFound { position }
        | BodyChartArrangementError::ChartPositionNotFound { position } => {
            BodyChartMetadataError::ChartNotFound { position }
        },
        BodyChartArrangementError::UnsupportedSource => BodyChartMetadataError::UnsupportedSource,
        BodyChartArrangementError::InvalidSource
        | BodyChartArrangementError::Verification
        | BodyChartArrangementError::PatchConflict => BodyChartMetadataError::InvalidSource,
        BodyChartArrangementError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyChartMetadataError::LimitExceeded {
            kind: map_arrangement_limit(kind),
            observed,
            maximum,
        },
        BodyChartArrangementError::Allocation { amount } => {
            BodyChartMetadataError::Allocation { amount }
        },
    }
}

fn map_arrangement_limit(kind: BodyChartArrangementLimitKind) -> BodyChartMetadataLimitKind {
    match kind {
        BodyChartArrangementLimitKind::InputBytes
        | BodyChartArrangementLimitKind::PayloadBytes
        | BodyChartArrangementLimitKind::TotalPayloadBytes => {
            BodyChartMetadataLimitKind::InputBytes
        },
        BodyChartArrangementLimitKind::PayloadReferences => {
            BodyChartMetadataLimitKind::PayloadReferences
        },
        BodyChartArrangementLimitKind::WireBytes
        | BodyChartArrangementLimitKind::WireOutputBytes => BodyChartMetadataLimitKind::WireBytes,
        BodyChartArrangementLimitKind::WireFields
        | BodyChartArrangementLimitKind::PayloadItems
        | BodyChartArrangementLimitKind::PayloadMessages
        | BodyChartArrangementLimitKind::PayloadObjects => BodyChartMetadataLimitKind::WireFields,
        BodyChartArrangementLimitKind::WireNesting => BodyChartMetadataLimitKind::WireNesting,
        BodyChartArrangementLimitKind::WireWork => BodyChartMetadataLimitKind::WireWork,
        BodyChartArrangementLimitKind::WireAllocations => {
            BodyChartMetadataLimitKind::WireAllocations
        },
        BodyChartArrangementLimitKind::WireRetainedBytes
        | BodyChartArrangementLimitKind::WireScratchBytes => {
            BodyChartMetadataLimitKind::WireRetainedBytes
        },
        BodyChartArrangementLimitKind::OutputBytes
        | BodyChartArrangementLimitKind::Entries
        | BodyChartArrangementLimitKind::EntryBytes
        | BodyChartArrangementLimitKind::TotalEntryBytes
        | BodyChartArrangementLimitKind::PackageBytes => BodyChartMetadataLimitKind::InputBytes,
    }
}

fn map_metadata_codec_error(error: chart_metadata_codec::DecodeError) -> BodyChartMetadataError {
    if let Some(amount) = error.allocation_amount() {
        return BodyChartMetadataError::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.input_limit_values() {
        return BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.label_limit_values() {
        return BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::LabelCount,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.text_limit_values() {
        return BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::TextBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.depth_limit_values() {
        return BodyChartMetadataError::LimitExceeded {
            kind: BodyChartMetadataLimitKind::WireNesting,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    BodyChartMetadataError::InvalidSource
}

fn map_title_codec_error(error: chart_title_codec::DecodeError) -> BodyChartMetadataError {
    if let Some(amount) = error.allocation_amount() {
        return BodyChartMetadataError::Allocation { amount };
    }
    if let Some(limit) = error.resource_limit() {
        return match limit {
            ChartTitleDecodeLimit::Bytes { observed, maximum } => {
                BodyChartMetadataError::LimitExceeded {
                    kind: BodyChartMetadataLimitKind::WireBytes,
                    observed: observed as u64,
                    maximum: maximum as u64,
                }
            },
            ChartTitleDecodeLimit::Fields { observed, maximum } => {
                BodyChartMetadataError::LimitExceeded {
                    kind: BodyChartMetadataLimitKind::WireFields,
                    observed: observed as u64,
                    maximum: maximum as u64,
                }
            },
            ChartTitleDecodeLimit::Work { observed, maximum } => {
                BodyChartMetadataError::LimitExceeded {
                    kind: BodyChartMetadataLimitKind::WireWork,
                    observed: observed as u64,
                    maximum: maximum as u64,
                }
            },
            ChartTitleDecodeLimit::Nesting { observed, maximum } => {
                BodyChartMetadataError::LimitExceeded {
                    kind: BodyChartMetadataLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            ChartTitleDecodeLimit::Output { observed, maximum }
            | ChartTitleDecodeLimit::Title { observed, maximum } => {
                BodyChartMetadataError::LimitExceeded {
                    kind: BodyChartMetadataLimitKind::TextBytes,
                    observed: observed as u64,
                    maximum: maximum as u64,
                }
            },
            _ => BodyChartMetadataError::InvalidSource,
        };
    }
    BodyChartMetadataError::InvalidSource
}

#[cfg(test)]
mod tests {
    use super::*;

    const NATIVE: &[u8] =
        include_bytes!("../../../../test-data/iwork/pages/chart-arrangement-native.pages");

    #[test]
    fn title_reference_outside_selected_metadata_is_rejected_as_foreign() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let mut budget = ArrangementBudget::new(&package).expect("arrangement budget");
        let target = body_chart_arrangement::resolve_target(
            &package,
            BodyChartSelector::index(0),
            &mut budget,
        )
        .expect("native chart target");
        let result = read_visible_title(
            &package,
            &target,
            NonZeroU64::new(u64::MAX).expect("nonzero reference"),
            &mut budget,
        );
        assert!(matches!(
            result,
            Err(BodyChartMetadataError::ForeignReference)
        ));
    }

    #[test]
    fn metadata_options_refuse_exhausted_aggregate_work_before_decode() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let mut budget = ArrangementBudget::new(&package).expect("arrangement budget");
        let BodyChartArrangementError::LimitExceeded { maximum, .. } =
            budget.work(usize::MAX).expect_err("finite work ceiling")
        else {
            panic!("expected a work limit");
        };
        budget
            .work(usize::try_from(maximum).expect("addressable work limit"))
            .expect("consume the remaining work budget");
        let result = metadata_options(&budget, b"chart");
        assert!(matches!(
            result,
            Err(BodyChartMetadataError::LimitExceeded {
                kind: BodyChartMetadataLimitKind::WireWork,
                ..
            })
        ));
    }

    #[test]
    fn title_decode_refuses_exhausted_retention_budget() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let mut budget = ArrangementBudget::new(&package).expect("arrangement budget");
        budget
            .retained(
                usize::try_from(package.state.source.limits().max_input_bytes())
                    .expect("addressable retention limit"),
            )
            .expect("consume the retention budget");
        let mut title = vec![0xa8, 0x01, 0x01, 0xba, 0x01, 13];
        title.extend_from_slice(b"bounded title");
        let mut source = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut source,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &title,
        )
        .expect("title extension framing");
        let view = litchi_iwa_common::wire::WireView::parse(&source).expect("title field");
        let result = decode_title(view.fields().next().expect("title extension"), &mut budget);
        assert!(matches!(
            result,
            Err(BodyChartMetadataError::LimitExceeded {
                kind: BodyChartMetadataLimitKind::WireRetainedBytes,
                ..
            })
        ));
    }
}
