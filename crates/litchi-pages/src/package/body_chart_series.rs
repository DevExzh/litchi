//! Selector-first Pages body-chart series-symbol transactions.
//!
//! This adapter owns only Pages selection, bounded admission and exact-source
//! publication.  The series override graph (including symbol visibility,
//! shape and point-size inheritance) remains in `litchi_iwa_style`.

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{
    Entry, EntryEdit, EntryInsertion, OwnedExactArtifacts, SharedBytes,
};
use litchi_iwa_archive::unstable::{
    CandidateIngressCharge, CandidateIngressError, CandidateIngressPlan, CandidateIngressVisitor,
};
use litchi_iwa_common::chart::error_bar::Series as ErrorBarSeries;
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::number_format::{LabelAffixes, NumberFormat};
use litchi_iwa_common::chart::pie::{LabelDistance, LabelVisibility, WedgeExplosion};
use litchi_iwa_common::chart::series::{
    ConnectionLine, ConnectionLineFamily, Stroke, Symbol, SymbolFamily, SymbolFill, Trendline,
};
use litchi_iwa_common::chart::series_labels::Visibility;
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_common::wire::{WireFieldView, WireView};
use litchi_iwa_core::{
    Archive, ArchiveEncodingPreparationChunk, ArchiveEncodingPreparationError, ArchiveObject,
    RawMessage, SnappyCompressionPreparationChunk, SnappyCompressionPreparationError, SnappyStream,
};
use litchi_iwa_protos::chart_metadata_codec::{self, DecodeOptions};
use litchi_iwa_protos::chart_style_codec;
use litchi_iwa_style::graph::{self as style_graph, Components, Workspace};
use litchi_iwa_style::series as series_engine;
use litchi_iwa_style::series::property as series_property;
use litchi_iwa_style::series::style as style_engine;
use thiserror::Error as ThisError;

use super::body_chart_arrangement::{self as arrangement, ArrangementBudget, ChartTarget};
#[cfg(test)]
use super::body_drawable_graph as package_graph;
use super::{Package, PackageError};
use crate::selector::BodyChartSelector;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_STYLE_MESSAGE_TYPES_FIRST: u32 = 5_022;
const CHART_STYLE_MESSAGE_TYPES_LAST: u32 = 5_029;
/// `TSWP.ParagraphStyleArchive`.
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const OBJECT_CONTAINER_MESSAGE_TYPE: u32 = 11_008;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const STYLE_EXTENSION_FIELD: u32 = 10_000;

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub(super) struct MutationPlan {
    components: Vec<Arc<str>>,
    objects: Vec<(Arc<str>, Vec<u64>)>,
    allowed_message_types: Vec<u32>,
    allowed_fields: Vec<u32>,
    chart_fields: Vec<u32>,
    planned_new_members: Vec<Arc<str>>,
    planned_removed_members: Vec<Arc<str>>,
}

impl MutationPlan {
    /// The plan of the inverse publication: members the forward edit created
    /// leave, and members it removed come back.
    fn inverse(&self) -> Self {
        let mut inverse = self.clone();
        std::mem::swap(
            &mut inverse.planned_new_members,
            &mut inverse.planned_removed_members,
        );
        inverse
    }
}

/// A finite resource charged by one Pages body-chart series-symbol operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// ZIP entries inspected or retained.
    Entries,
    /// One ZIP entry's bytes.
    EntryBytes,
    /// Aggregate ZIP entry bytes.
    TotalEntryBytes,
    /// ZIP metadata bytes.
    PackageBytes,
    /// Native payload bytes.
    PayloadBytes,
    /// Aggregate native payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadMessages,
    /// Native references.
    PayloadReferences,
    /// Native metadata items such as registry entries.
    PayloadItems,
    /// Native object identifiers allocated or released.
    PayloadObjects,
    /// Wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Wire and graph work.
    WireWork,
    /// Bounded allocations.
    Allocations,
    /// Retained source/candidate bytes.
    RetainedBytes,
    /// Candidate scratch bytes.
    ScratchBytes,
    /// Physical archive resource.
    Archive(litchi_iwa_core::LimitKind),
    /// Package ingress/reassembly resource.
    Package(litchi_iwa_archive::LimitKind),
    /// Shared wire resource.
    Wire(litchi_iwa_common::LimitKind),
    /// Chart-style codec resource.
    Codec(chart_style_codec::LimitKind),
    /// Package-metadata registry resource.
    Registry(litchi_iwa_protos::package_metadata_codec::RewriteLimit),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InputBytes => formatter.write_str("input bytes"),
            Self::OutputBytes => formatter.write_str("output bytes"),
            Self::Entries => formatter.write_str("ZIP entries"),
            Self::EntryBytes => formatter.write_str("ZIP entry bytes"),
            Self::TotalEntryBytes => formatter.write_str("total ZIP entry bytes"),
            Self::PackageBytes => formatter.write_str("package bytes"),
            Self::PayloadBytes => formatter.write_str("payload bytes"),
            Self::TotalPayloadBytes => formatter.write_str("total payload bytes"),
            Self::PayloadMessages => formatter.write_str("payload messages"),
            Self::PayloadReferences => formatter.write_str("payload references"),
            Self::PayloadItems => formatter.write_str("payload items"),
            Self::PayloadObjects => formatter.write_str("payload objects"),
            Self::WireFields => formatter.write_str("wire fields"),
            Self::WireNesting => formatter.write_str("wire nesting"),
            Self::WireWork => formatter.write_str("wire work"),
            Self::Allocations => formatter.write_str("allocations"),
            Self::RetainedBytes => formatter.write_str("retained bytes"),
            Self::ScratchBytes => formatter.write_str("scratch bytes"),
            Self::Archive(kind) => write!(formatter, "archive {kind}"),
            Self::Package(kind) => write!(formatter, "package {kind:?}"),
            Self::Wire(kind) => write!(formatter, "wire {kind}"),
            Self::Codec(kind) => write!(formatter, "codec {kind:?}"),
            Self::Registry(kind) => write!(formatter, "registry {kind:?}"),
        }
    }
}

/// A well-formed symbol request rejected without host fallback.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No chart matched the checked body position.
    #[error("the Pages body has no chart at position {position:?}")]
    ChartNotFound { position: Position },
    /// The source is not an exact package artifact.
    #[error("this Pages source does not support exact body-chart series edits")]
    UnsupportedSource,
    /// The property does not apply to the chart's kind.
    #[error("the property does not apply to the selected chart kind")]
    UnsupportedKind,
    /// The value vector does not match the chart's visible series count.
    #[error("the series vector has {given} values; expected {expected}")]
    SeriesCount { expected: usize, given: usize },
    /// The series is past the chart's series.
    #[error("the selected chart has no series {series}")]
    SeriesNotFound {
        /// The requested zero-based series.
        series: usize,
    },
    /// The property has no reset (only the series-style properties do).
    #[error("the chart property cannot be reset")]
    NotResettable,
    /// A style object is shared by more than one chart.
    #[error("the selected chart style is shared by another chart")]
    SharedStyle,
    /// An image fill needs the embedded-data graph this family does not
    /// write.
    #[error("image fills are not written by this family")]
    ImageFill,
    /// The property does not exist on the requested axis, or the chart has
    /// no primary axis of that direction.
    #[error("the chart axis has no such property")]
    UnsupportedAxis,
    /// The chart has no style object for the property.
    #[error("the chart has no style object for the property")]
    MissingStyle,
    /// No operation was staged.
    #[error("no body-chart series-symbol change is staged")]
    NoStagedOperation,
    /// The patch belongs to another semantic property family.
    #[error("the patch does not belong to body-chart series symbols")]
    WrongProperty,
    /// The engine produced a publication shape this route does not own.
    #[error("the series-symbol graph produced unsupported publication changes")]
    UnsupportedPublication,
}

impl Refusal {
    /// Classify this refusal for the unified-operations boundary.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::ChartNotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedSource
            | Self::UnsupportedKind
            | Self::SeriesCount { .. }
            | Self::WrongProperty
            | Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::SeriesNotFound { .. } => RefusalKind::NotFound,
            Self::NotResettable => RefusalKind::UnsupportedOperation,
            Self::UnsupportedAxis | Self::MissingStyle => RefusalKind::UnsupportedTarget,
            Self::SharedStyle | Self::ImageFill | Self::UnsupportedPublication => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

/// Failure from a Pages body-chart series-symbol read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A typed terminal refusal.
    #[error(transparent)]
    Refused(Refusal),
    /// The selected chart graph or metadata is malformed.
    #[error("the selected Pages body-chart series-symbol source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Pages body-chart series symbols {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} Pages body-chart series-symbol units")]
    Allocation { amount: usize },
    /// A finite finalization arithmetic operation overflowed.
    #[error("Pages body-chart series-symbol finalization overflow: {reason}")]
    Overflow { reason: &'static str },
    /// Candidate semantic readback did not reproduce the requested vector.
    #[error("the edited Pages body-chart series symbols failed semantic verification")]
    Verification,
    /// The patch was created from another exact source artifact.
    #[error("the Pages body-chart series-symbol patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(value: Refusal) -> Self {
        Self::Refused(value)
    }
}

/// One per-series chart property: a value for every visible series, in
/// series order.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// The line joining a series' points (line, area, radar and scatter
    /// charts).
    ConnectionLines(Vec<ConnectionLine>),
    /// How far each wedge is pulled out of a pie or donut.
    WedgeExplosions(Vec<WedgeExplosion>),
    /// How far each wedge's label sits from the pie centre.
    LabelDistances(Vec<LabelDistance>),
    /// One optional symbol value per visible chart series; `None` hides it.
    Symbols(Vec<Option<Symbol>>),
    /// The trendline of each series.
    Trendlines(Vec<Trendline>),
    /// The error bars of each series.
    ErrorBars(Vec<ErrorBarSeries>),
    /// Which labels each pie wedge shows.
    PieLabels(Vec<LabelVisibility>),
    /// The value-label visibility of each series.
    ValueLabels(Vec<Visibility>),
    /// The value-label prefix and suffix of each series.
    ValueLabelAffixes(Vec<LabelAffixes>),
    /// The value-label number format of each series.
    ValueLabelNumberFormats(Vec<NumberFormat>),
    /// The stroke of each series (`None`: no stroke).
    Strokes(Vec<Option<Stroke>>),
    /// The fill of each series.
    Fills(Vec<Fill>),
    /// The fill of each series' data symbols.
    SymbolFills(Vec<SymbolFill>),
    /// The outline of each series' data symbols (`None`: no outline).
    SymbolOutlines(Vec<Option<Stroke>>),
}

/// The property a [`Property`] value sets, for reads and resets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyKind {
    /// [`Property::ConnectionLines`].
    ConnectionLines,
    /// [`Property::WedgeExplosions`].
    WedgeExplosions,
    /// [`Property::LabelDistances`].
    LabelDistances,
    /// [`Property::Symbols`].
    Symbols,
    /// [`Property::Trendlines`].
    Trendlines,
    /// [`Property::ErrorBars`].
    ErrorBars,
    /// [`Property::PieLabels`].
    PieLabels,
    /// [`Property::ValueLabels`].
    ValueLabels,
    /// [`Property::ValueLabelAffixes`].
    ValueLabelAffixes,
    /// [`Property::ValueLabelNumberFormats`].
    ValueLabelNumberFormats,
    /// [`Property::Strokes`].
    Strokes,
    /// [`Property::Fills`].
    Fills,
    /// [`Property::SymbolFills`].
    SymbolFills,
    /// [`Property::SymbolOutlines`].
    SymbolOutlines,
}

impl PropertyKind {
    /// The series-style property this kind names, if it is one (the
    /// properties [`Edit::reset`] applies to).
    const fn style(self) -> Option<style_engine::Kind> {
        match self {
            Self::Strokes => Some(style_engine::Kind::Strokes),
            Self::Fills => Some(style_engine::Kind::Fills),
            Self::SymbolFills => Some(style_engine::Kind::SymbolFills),
            Self::SymbolOutlines => Some(style_engine::Kind::SymbolOutlines),
            _ => None,
        }
    }
}

impl Property {
    /// Return the property kind.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        match self {
            Self::ConnectionLines(_) => PropertyKind::ConnectionLines,
            Self::WedgeExplosions(_) => PropertyKind::WedgeExplosions,
            Self::LabelDistances(_) => PropertyKind::LabelDistances,
            Self::Symbols(_) => PropertyKind::Symbols,
            Self::Trendlines(_) => PropertyKind::Trendlines,
            Self::ErrorBars(_) => PropertyKind::ErrorBars,
            Self::PieLabels(_) => PropertyKind::PieLabels,
            Self::ValueLabels(_) => PropertyKind::ValueLabels,
            Self::ValueLabelAffixes(_) => PropertyKind::ValueLabelAffixes,
            Self::ValueLabelNumberFormats(_) => PropertyKind::ValueLabelNumberFormats,
            Self::Strokes(_) => PropertyKind::Strokes,
            Self::Fills(_) => PropertyKind::Fills,
            Self::SymbolFills(_) => PropertyKind::SymbolFills,
            Self::SymbolOutlines(_) => PropertyKind::SymbolOutlines,
        }
    }

    /// Return the number of values supplied.
    #[must_use]
    pub fn len(&self) -> usize {
        match self {
            Self::ConnectionLines(values) => values.len(),
            Self::WedgeExplosions(values) => values.len(),
            Self::LabelDistances(values) => values.len(),
            Self::Symbols(values) => values.len(),
            Self::Trendlines(values) => values.len(),
            Self::ErrorBars(values) => values.len(),
            Self::PieLabels(values) => values.len(),
            Self::ValueLabels(values) => values.len(),
            Self::ValueLabelAffixes(values) => values.len(),
            Self::ValueLabelNumberFormats(values) => values.len(),
            Self::Strokes(values) | Self::SymbolOutlines(values) => values.len(),
            Self::Fills(values) => values.len(),
            Self::SymbolFills(values) => values.len(),
        }
    }

    /// Whether the property contains no series values.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// The per-series style values of a series-style property.
    fn style_values(&self) -> Option<Vec<Option<style_engine::Value>>> {
        Some(match self {
            Self::Strokes(values) | Self::SymbolOutlines(values) => values
                .iter()
                .map(|value| Some(style_engine::Value::Stroke(*value)))
                .collect(),
            Self::Fills(values) => values
                .iter()
                .map(|value| Some(style_engine::Value::Fill(value.clone())))
                .collect(),
            Self::SymbolFills(values) => values
                .iter()
                .map(|value| Some(style_engine::Value::SymbolFill(value.clone())))
                .collect(),
            _ => return None,
        })
    }
}

/// The selected body chart with its kind and visible series count.
#[derive(Clone, PartialEq)]
pub(super) struct Target {
    pub(super) chart: ChartTarget,
    pub(super) kind: ChartKind,
    pub(super) series: usize,
}

struct PackageComponents<'a>(&'a Package);

impl Components for PackageComponents<'_> {
    fn count(&self) -> usize {
        self.0.state.source.components().len()
    }

    fn component(&self, index: usize) -> Option<(&str, &Archive)> {
        self.0
            .state
            .source
            .components()
            .get_index(index)
            .map(|component| (component.name(), component.archive()))
    }
}

fn map_arrangement_limit(kind: arrangement::LimitKind) -> LimitKind {
    match kind {
        arrangement::LimitKind::InputBytes => LimitKind::InputBytes,
        arrangement::LimitKind::PayloadBytes => LimitKind::PayloadBytes,
        arrangement::LimitKind::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        arrangement::LimitKind::PayloadReferences => LimitKind::PayloadReferences,
        arrangement::LimitKind::PayloadItems => LimitKind::PayloadItems,
        arrangement::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
        arrangement::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
        arrangement::LimitKind::WireBytes => {
            LimitKind::Wire(litchi_iwa_common::LimitKind::InputBytes)
        },
        arrangement::LimitKind::WireOutputBytes => {
            LimitKind::Wire(litchi_iwa_common::LimitKind::OutputBytes)
        },
        arrangement::LimitKind::WireFields => LimitKind::WireFields,
        arrangement::LimitKind::WireNesting => LimitKind::WireNesting,
        arrangement::LimitKind::WireWork => LimitKind::WireWork,
        arrangement::LimitKind::WireAllocations => LimitKind::Allocations,
        arrangement::LimitKind::WireRetainedBytes => LimitKind::RetainedBytes,
        arrangement::LimitKind::WireScratchBytes => LimitKind::ScratchBytes,
        arrangement::LimitKind::OutputBytes => LimitKind::OutputBytes,
        arrangement::LimitKind::Entries => LimitKind::Entries,
        arrangement::LimitKind::EntryBytes => LimitKind::EntryBytes,
        arrangement::LimitKind::TotalEntryBytes => LimitKind::TotalEntryBytes,
        arrangement::LimitKind::PackageBytes => LimitKind::PackageBytes,
    }
}

fn map_arrangement_error(error: arrangement::Error) -> Error {
    match error {
        arrangement::Error::Refused(arrangement::Refusal::ChartNotFound { position })
        | arrangement::Error::Refused(arrangement::Refusal::ChartPositionNotFound { position }) => {
            Refusal::ChartNotFound { position }.into()
        },
        arrangement::Error::Refused(arrangement::Refusal::UnsupportedSource) => {
            Refusal::UnsupportedSource.into()
        },
        arrangement::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_arrangement_limit(kind),
            observed,
            maximum,
        },
        arrangement::Error::Allocation { amount } => Error::Allocation { amount },
        arrangement::Error::InvalidSource => Error::InvalidSource,
        arrangement::Error::Verification => Error::Verification,
        arrangement::Error::PatchConflict => Error::PatchConflict,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: LimitKind::Package(kind),
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Archive(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_registry_error(error: litchi_iwa_protos::package_metadata_codec::RewriteError) -> Error {
    if let Some(limit) = error.resource_limit() {
        let (observed, maximum) = match limit {
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::InputBytes {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::OutputBytes {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Fields {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Work { observed, maximum }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::References {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Additions {
                observed,
                maximum,
            } => (observed as u64, maximum as u64),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Nesting {
                observed,
                maximum,
            } => (u64::from(observed), u64::from(maximum)),
            _ => return Error::InvalidSource,
        };
        return Error::Limit {
            kind: LimitKind::Registry(limit),
            observed,
            maximum,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    Error::InvalidSource
}

fn map_graph_error(error: style_graph::Error) -> Error {
    match error {
        style_graph::Error::Archive(error) => map_core_error(error),
        style_graph::Error::Wire(error) => match error {
            litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit,
            } => Error::Limit {
                kind: LimitKind::Wire(kind),
                observed: observed as u64,
                maximum: limit as u64,
            },
            litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
            _ => Error::InvalidSource,
        },
        style_graph::Error::Codec(error) => match error {
            chart_style_codec::Error::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: LimitKind::Codec(kind),
                observed: observed as u64,
                maximum: maximum as u64,
            },
            chart_style_codec::Error::Allocation { bytes } => Error::Allocation { amount: bytes },
            _ => Error::InvalidSource,
        },
        style_graph::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                style_graph::GraphLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                style_graph::GraphLimitKind::AllocationEvents => LimitKind::Allocations,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        style_graph::Error::Registry(error) => map_registry_error(error),
        style_graph::Error::IdentifierOverflow => Refusal::UnsupportedPublication.into(),
        style_graph::Error::Invalid("a series style is shared by several charts") => {
            Refusal::SharedStyle.into()
        },
        _ => Error::InvalidSource,
    }
}

fn map_symbol_error(error: series_engine::SymbolError) -> Error {
    match error {
        series_engine::SymbolError::Graph(error) => map_graph_error(error),
        series_engine::SymbolError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                series_engine::LimitKind::InputBytes => LimitKind::InputBytes,
                series_engine::LimitKind::Fields => LimitKind::WireFields,
                series_engine::LimitKind::OutputBytes => LimitKind::OutputBytes,
                series_engine::LimitKind::Nesting => LimitKind::WireNesting,
                series_engine::LimitKind::Work => LimitKind::WireWork,
                series_engine::LimitKind::Allocations => LimitKind::Allocations,
                series_engine::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                series_engine::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                series_engine::LimitKind::References => LimitKind::PayloadReferences,
                series_engine::LimitKind::RegistryEntries => LimitKind::Registry(
                    litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
                        observed,
                        maximum,
                    },
                ),
                series_engine::LimitKind::Identifiers => LimitKind::PayloadObjects,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        series_engine::SymbolError::Allocation { amount } => Error::Allocation { amount },
        series_engine::SymbolError::Invalid(_) => Error::InvalidSource,
        series_engine::SymbolError::IdentifierOverflow => Refusal::UnsupportedPublication.into(),
        series_engine::SymbolError::Refused(refusal) => match refusal {
            series_engine::Refusal::UnsupportedKind => Refusal::UnsupportedKind.into(),
            series_engine::Refusal::SharedStyle => Refusal::SharedStyle.into(),
            series_engine::Refusal::ImageFill => Refusal::ImageFill.into(),
            series_engine::Refusal::UnsupportedAxis => Refusal::UnsupportedAxis.into(),
            series_engine::Refusal::MissingStyle => Refusal::MissingStyle.into(),
            _ => Refusal::UnsupportedPublication.into(),
        },
        _ => Error::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                crate::LimitKind::Source(kind) => match kind {
                    crate::SourceLimitKind::InputBytes => LimitKind::InputBytes,
                    crate::SourceLimitKind::Members => LimitKind::Entries,
                    crate::SourceLimitKind::EntryBytes => LimitKind::EntryBytes,
                    crate::SourceLimitKind::AggregateBytes => LimitKind::TotalEntryBytes,
                    crate::SourceLimitKind::ComponentBytes => LimitKind::PayloadBytes,
                },
                crate::LimitKind::Semantic(kind) => match kind {
                    crate::SemanticLimitKind::Sections => LimitKind::PayloadItems,
                    crate::SemanticLimitKind::TextBytes => LimitKind::PayloadBytes,
                },
            },
            observed,
            maximum,
        },
        PackageError::PayloadLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadObjects,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::SectionNamesTooLarge { observed, limit }
        | PackageError::Semantic(crate::Error::TextTooLarge { observed, limit }) => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::Semantic(
            crate::Error::TooManySections { actual, limit }
            | crate::Error::TooManyBodyStorages { actual, limit },
        ) => Error::Limit {
            kind: LimitKind::PayloadItems,
            observed: actual as u64,
            maximum: limit as u64,
        },
        PackageError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn reopen_series_candidate(source: &Package, target: Arc<[u8]>) -> Result<Package, Error> {
    let catalog = source
        .state
        .source
        .reopen(target)
        .map_err(map_archive_error)?;
    Package::from_source_catalog_with_semantic_limits(catalog, source.state.semantic_limits)
        .map_err(map_package_error)
}

fn archive_limits(package: &Package) -> Result<litchi_iwa_core::Limits, Error> {
    package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)
}

fn positive_limit(value: usize, kind: LimitKind) -> Result<usize, Error> {
    if value == 0 {
        return Err(Error::Limit {
            kind,
            observed: 1,
            maximum: 0,
        });
    }
    Ok(value)
}

fn engine_limits(
    package: &Package,
    budget: &ArrangementBudget,
) -> Result<series_engine::Limits, Error> {
    let physical = package.state.source.limits();
    let package_bytes = usize::try_from(physical.max_input_bytes())
        .map_err(|_| Error::InvalidSource)?
        .max(1);
    let total_bytes = usize::try_from(physical.max_total_bytes())
        .map_err(|_| Error::InvalidSource)?
        .max(1);
    let wire = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    let codec_max = chart_style_codec::Limits::DEFAULT.max_bytes();
    let input = budget
        .residual_input_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes)
        .min(litchi_iwa_common::WireLimits::MAX_INPUT_BYTES)
        .min(codec_max);
    let input = positive_limit(input, LimitKind::InputBytes)?;
    let output = budget
        .residual_output_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes)
        .min(codec_max);
    let output = positive_limit(output, LimitKind::OutputBytes)?;
    let fields = wire
        .max_fields()
        .min(chart_style_codec::Limits::DEFAULT.max_fields())
        .min(litchi_iwa_common::WireLimits::MAX_FIELDS);
    let fields = positive_limit(fields, LimitKind::WireFields)?;
    let work = budget
        .residual_work_bytes()
        .map_err(map_arrangement_error)?
        .min(total_bytes.saturating_mul(32).max(1))
        .min(litchi_iwa_common::WireLimits::MAX_REWRITE_WORK);
    let work = positive_limit(work, LimitKind::WireWork)?;
    let allocations = budget
        .residual_allocations()
        .map_err(map_arrangement_error)?;
    let allocations = positive_limit(allocations, LimitKind::Allocations)?;
    let retained = budget
        .residual_retained_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes);
    let retained = positive_limit(retained, LimitKind::RetainedBytes)?;
    let scratch = budget
        .residual_scratch_bytes()
        .map_err(map_arrangement_error)?
        .min(package_bytes);
    let scratch = positive_limit(scratch, LimitKind::ScratchBytes)?;
    let references = budget
        .residual_references()
        .map_err(map_arrangement_error)?;
    let references = positive_limit(references, LimitKind::PayloadReferences)?;
    let registry_entries = budget
        .residual_registry_entries()
        .map_err(map_arrangement_error)?;
    let registry_entries = positive_limit(registry_entries, LimitKind::PayloadItems)?;
    let identifiers = budget
        .residual_identifiers()
        .map_err(map_arrangement_error)?;
    let identifiers = positive_limit(identifiers, LimitKind::PayloadObjects)?;
    Ok(series_engine::Limits::new(
        input,
        fields,
        output,
        u32::try_from(wire.max_nesting()).map_err(|_| Error::InvalidSource)?,
        work,
        allocations,
        retained,
        scratch,
        references,
        registry_entries,
        identifiers,
    ))
}

fn charge_engine_report(
    budget: &mut ArrangementBudget,
    report: &series_engine::Report,
) -> Result<(), Error> {
    let max_nesting = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?
        .max_nesting();
    budget
        .input(report.input_bytes())
        .and_then(|_| budget.output(report.output_bytes()))
        .and_then(|_| budget.fields(report.fields()))
        .and_then(|_| budget.work(report.work_bytes()))
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.scratch(report.scratch_bytes()))
        .and_then(|_| budget.references(report.references()))
        .and_then(|_| budget.registry_entries(report.registry_entries()))
        .and_then(|_| budget.identifiers(report.identifiers()))
        .map_err(map_arrangement_error)?;
    if report.max_depth() as usize > max_nesting {
        return Err(Error::Limit {
            kind: LimitKind::WireNesting,
            observed: u64::from(report.max_depth()),
            maximum: max_nesting as u64,
        });
    }
    Ok(())
}

fn map_finish_error(error: style_graph::FinishError) -> Error {
    match error {
        style_graph::FinishError::Graph(error) => map_graph_error(error),
        style_graph::FinishError::Overflow(reason) => Error::Overflow { reason },
        style_graph::FinishError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                style_graph::FinishLimitKind::Work => LimitKind::WireWork,
                style_graph::FinishLimitKind::Allocations => LimitKind::Allocations,
                style_graph::FinishLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                style_graph::FinishLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        style_graph::FinishError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn preflight_finish(
    budget: &ArrangementBudget,
    requirements: style_graph::FinishRequirements,
) -> Result<(), Error> {
    budget
        .preflight_work(requirements.work())
        .and_then(|_| budget.preflight_allocations(requirements.allocations()))
        .and_then(|_| budget.preflight_retained(requirements.retained_bytes()))
        .and_then(|_| budget.preflight_scratch(requirements.scratch_bytes()))
        .map_err(map_arrangement_error)
}

fn finish_limits(budget: &ArrangementBudget) -> Result<style_graph::FinishLimits, Error> {
    Ok(style_graph::FinishLimits::new(
        budget
            .residual_work_bytes()
            .map_err(map_arrangement_error)?,
        budget
            .residual_allocations()
            .map_err(map_arrangement_error)?,
        budget
            .residual_retained_bytes()
            .map_err(map_arrangement_error)?,
        budget
            .residual_scratch_bytes()
            .map_err(map_arrangement_error)?,
    ))
}

fn charge_finish_report(
    budget: &mut ArrangementBudget,
    report: style_graph::FinishReport,
) -> Result<(), Error> {
    budget
        .work(report.work())
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.scratch(report.scratch_bytes()))
        .map_err(map_arrangement_error)
}

fn target_with_budget(
    package: &Package,
    selector: BodyChartSelector,
    budget: &mut ArrangementBudget,
) -> Result<ChartTarget, Error> {
    arrangement::resolve_target(package, selector, budget).map_err(map_arrangement_error)
}

fn admit_workspace(package: &Package, budget: &mut ArrangementBudget) -> Result<(), Error> {
    let source_bytes = package.state.source.shared_source().len();
    if !budget.workspace_source_admitted {
        budget.input(source_bytes).map_err(map_arrangement_error)?;
        budget
            .retained(source_bytes)
            .map_err(map_arrangement_error)?;
        budget.workspace_source_admitted = true;
    }
    Ok(())
}

fn map_construction_error(error: style_graph::ConstructionError) -> Error {
    match error {
        style_graph::ConstructionError::Graph(error) => map_graph_error(error),
        style_graph::ConstructionError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                style_graph::ConstructionLimitKind::Work => LimitKind::WireWork,
                style_graph::ConstructionLimitKind::Allocations => LimitKind::Allocations,
                style_graph::ConstructionLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                style_graph::ConstructionLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        style_graph::ConstructionError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn construction_limits(
    budget: &ArrangementBudget,
) -> Result<style_graph::ConstructionLimits, Error> {
    Ok(style_graph::ConstructionLimits::new(
        budget
            .residual_work_bytes()
            .map_err(map_arrangement_error)?,
        budget
            .residual_allocations()
            .map_err(map_arrangement_error)?,
        budget
            .residual_retained_bytes()
            .map_err(map_arrangement_error)?,
        budget
            .residual_scratch_bytes()
            .map_err(map_arrangement_error)?,
    ))
}

fn charge_workspace_construction(
    budget: &mut ArrangementBudget,
    requirements: style_graph::ConstructionRequirements,
) -> Result<(), Error> {
    budget
        .preflight_work(requirements.total_work())
        .and_then(|_| budget.preflight_allocations(requirements.index_allocations()))
        .and_then(|_| budget.preflight_retained(requirements.index_retained_bytes()))
        .and_then(|_| budget.preflight_scratch(requirements.index_scratch_bytes()))
        .map_err(map_arrangement_error)?;
    budget
        .work(requirements.total_work())
        .and_then(|_| budget.allocations(requirements.index_allocations()))
        .and_then(|_| budget.retained(requirements.index_retained_bytes()))
        .and_then(|_| budget.scratch(requirements.index_scratch_bytes()))
        .map_err(map_arrangement_error)
}

fn metadata_options(budget: &ArrangementBudget, source: &[u8]) -> Result<DecodeOptions, Error> {
    let limits = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    let input = positive_limit(
        source.len().min(limits.max_input_bytes()),
        LimitKind::InputBytes,
    )?;
    let fields = positive_limit(limits.max_fields(), LimitKind::WireFields)?;
    let work = positive_limit(limits.max_rewrite_work(), LimitKind::WireWork)?;
    let labels = positive_limit(limits.max_fields(), LimitKind::PayloadItems)?;
    let text = positive_limit(limits.max_input_bytes(), LimitKind::PayloadBytes)?;
    Ok(DecodeOptions::new(
        input,
        fields,
        work,
        u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?,
        labels.min(1_000_000),
        text.min(64 * 1024 * 1024),
    ))
}

fn charge_metadata_report(
    budget: &mut ArrangementBudget,
    report: chart_metadata_codec::DecodeReport,
) -> Result<(), Error> {
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
        .map_err(map_arrangement_error)
}

fn map_metadata_error(error: chart_metadata_codec::DecodeError) -> Error {
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.input_limit_values() {
        return Error::Limit {
            kind: LimitKind::InputBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.label_limit_values() {
        return Error::Limit {
            kind: LimitKind::PayloadItems,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.text_limit_values() {
        return Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.depth_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireNesting,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

fn series_count(kind: ChartKind, components: usize) -> Result<usize, Error> {
    series_engine::visible_series(kind, components).ok_or_else(|| match kind {
        ChartKind::Scatter2d | ChartKind::Bubble2d => Refusal::UnsupportedKind.into(),
        _ => Error::InvalidSource,
    })
}

fn target_with_metadata(
    package: &Package,
    selector: BodyChartSelector,
    budget: &mut ArrangementBudget,
) -> Result<Target, Error> {
    let chart = target_with_budget(package, selector, budget)?;
    let source =
        arrangement::selected_chart_payload(package, &chart).map_err(map_arrangement_error)?;
    let options = metadata_options(budget, source)?;
    let (snapshot, report) = match chart_metadata_codec::decode_modern_with_report(source, &options)
    {
        Ok(decoded) => decoded,
        Err(error) => {
            charge_metadata_report(budget, error.report())?;
            return Err(map_metadata_error(error));
        },
    };
    charge_metadata_report(budget, report)?;
    let kind = ChartKind::from_native(snapshot.chart_type());
    let series = series_count(kind, snapshot.series_count())?;
    Ok(Target {
        chart,
        kind,
        series,
    })
}

fn preflight_publication(
    source: &Package,
    edits: &[(String, Archive)],
    insertions: &[style_graph::Insertion],
    deletions: &[String],
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let vector_caps = [
        (edits.len(), size_of::<(String, Vec<u8>)>()),
        (
            insertions.len(),
            size_of::<(String, Vec<u8>, Option<String>)>(),
        ),
        (edits.len(), size_of::<EntryEdit<'static>>()),
        (insertions.len(), size_of::<EntryInsertion<'static>>()),
        (source.state.source.package().len(), size_of::<&str>()),
        (deletions.len(), size_of::<&str>()),
    ];
    let collection_units = vector_caps.iter().try_fold(0usize, |total, (count, _)| {
        total
            .checked_add(usize::from(*count != 0))
            .ok_or(Error::InvalidSource)
    })?;
    let collection_bytes = vector_caps
        .iter()
        .try_fold(0usize, |total, (count, size)| {
            count
                .checked_mul(*size)
                .and_then(|amount| total.checked_add(amount))
                .ok_or(Error::InvalidSource)
        })?;
    budget
        .preflight_work(collection_bytes)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_allocations(collection_units)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(collection_bytes)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_scratch(collection_bytes)
        .map_err(map_arrangement_error)?;
    budget
        .work(collection_bytes)
        .map_err(map_arrangement_error)?;
    budget
        .allocations(collection_units)
        .map_err(map_arrangement_error)?;
    budget
        .retained(collection_bytes)
        .map_err(map_arrangement_error)?;
    budget
        .scratch(collection_bytes)
        .map_err(map_arrangement_error)?;
    Ok(())
}

fn map_archive_encoding_preparation_error(error: ArchiveEncodingPreparationError<Error>) -> Error {
    match error {
        ArchiveEncodingPreparationError::Admission(error) => error,
        ArchiveEncodingPreparationError::ArithmeticOverflow { resource } => {
            Error::Overflow { reason: resource }
        },
        ArchiveEncodingPreparationError::Core(error) => map_core_error(error),
        _ => Error::InvalidSource,
    }
}

fn map_snappy_preparation_error(error: SnappyCompressionPreparationError<Error>) -> Error {
    match error {
        SnappyCompressionPreparationError::Admission(error) => error,
        SnappyCompressionPreparationError::ArithmeticOverflow { resource } => {
            Error::Overflow { reason: resource }
        },
        SnappyCompressionPreparationError::Core(error) => map_core_error(error),
        _ => Error::InvalidSource,
    }
}

fn prepare_archive_encoding<'archive>(
    archive: &'archive Archive,
    limits: litchi_iwa_core::Limits,
    budget: &mut ArrangementBudget,
) -> Result<litchi_iwa_core::ArchiveEncodingPlan<'archive>, Error> {
    let mut visitor = |chunk: ArchiveEncodingPreparationChunk| {
        budget
            .work(chunk.traversal_bytes())
            .map_err(map_arrangement_error)
    };
    archive
        .prepare_encoding(limits, &mut visitor)
        .map_err(map_archive_encoding_preparation_error)
}

fn charge_archive_encoding_requirements(
    requirements: litchi_iwa_core::ArchiveEncodingRequirements,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    budget
        .work(requirements.work_bytes())
        .and_then(|_| budget.work(requirements.identifier_sort_work()))
        .and_then(|_| budget.allocations(requirements.allocation_events()))
        .and_then(|_| budget.retained(requirements.retained_bytes()))
        .and_then(|_| budget.scratch(requirements.scratch_bytes()))
        .map_err(map_arrangement_error)
}

fn prepare_snappy_compression(
    input_len: usize,
    limits: litchi_iwa_core::SnappyLimits,
    budget: &mut ArrangementBudget,
) -> Result<litchi_iwa_core::SnappyCompressionPlan, Error> {
    let mut visitor = |chunk: SnappyCompressionPreparationChunk| {
        budget
            .work(chunk.traversal_bytes())
            .map_err(map_arrangement_error)
    };
    SnappyStream::prepare_compression(input_len, limits, &mut visitor)
        .map_err(map_snappy_preparation_error)
}

fn charge_snappy_compression_requirements(
    requirements: litchi_iwa_core::SnappyCompressionRequirements,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let scratch_bytes = requirements
        .temporary_frame_capacity_bytes()
        .checked_add(requirements.encoder_workspace_bytes())
        .ok_or(Error::Overflow {
            reason: "Snappy compression scratch overflow",
        })?;
    budget
        .work(requirements.work_bytes())
        .and_then(|_| budget.allocations(requirements.allocation_events()))
        .and_then(|_| budget.retained(requirements.output_capacity_bytes()))
        .and_then(|_| budget.scratch(scratch_bytes))
        .map_err(map_arrangement_error)
}

fn encode_archive_for_publication(
    archive: &Archive,
    archive_limits: litchi_iwa_core::Limits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    budget: &mut ArrangementBudget,
) -> Result<Vec<u8>, Error> {
    let prepared = prepare_archive_encoding(archive, archive_limits, budget)?;
    let requirements = prepared.requirements();
    charge_archive_encoding_requirements(requirements, budget)?;
    let encoded = prepared.execute().map_err(map_core_error)?;
    let prepared = prepare_snappy_compression(encoded.len(), snappy_limits, budget)?;
    let requirements = prepared.requirements();
    charge_snappy_compression_requirements(requirements, budget)?;
    prepared.execute(&encoded).map_err(map_core_error)
}

struct IngressBudget<'a>(&'a mut ArrangementBudget);

impl CandidateIngressVisitor for IngressBudget<'_> {
    type Error = Error;

    fn admit(&mut self, charge: CandidateIngressCharge) -> Result<(), Self::Error> {
        self.0
            .retained(charge.retained_bytes())
            .and_then(|_| self.0.work(charge.traversal_bytes()))
            .and_then(|_| self.0.allocations(charge.allocations()))
            .map_err(map_arrangement_error)
    }
}

fn map_candidate_ingress_error(error: CandidateIngressError<Error>) -> Error {
    match error {
        CandidateIngressError::Physical(error) => map_archive_error(error),
        CandidateIngressError::Admission(error) => error,
        CandidateIngressError::CloneArithmetic { .. } => Error::InvalidSource,
        _ => Error::InvalidSource,
    }
}

fn charge_candidate_plan(
    plan: &CandidateIngressPlan<'_>,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let requirements = plan.requirements();
    budget
        .input(requirements.input_bytes())
        .and_then(|_| budget.work(requirements.traversal_bytes()))
        .and_then(|_| budget.allocations(requirements.allocations()))
        .and_then(|_| budget.retained(requirements.retained_bytes()))
        .map_err(map_arrangement_error)?;
    Ok(())
}

fn publish_changes(
    source: &Package,
    changes: style_graph::Changes,
    budget: &mut ArrangementBudget,
) -> Result<(Package, usize), Error> {
    let touched_components = changes
        .edits
        .len()
        .checked_add(changes.insertions.len())
        .and_then(|value| value.checked_add(changes.deletions.len()))
        .ok_or(Error::InvalidSource)?;
    preflight_publication(
        source,
        &changes.edits,
        &changes.insertions,
        &changes.deletions,
        budget,
    )?;
    let archive_limits = archive_limits(source)?;
    let snappy_limits = source
        .state
        .source
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let mut edited = Vec::new();
    edited
        .try_reserve_exact(changes.edits.len())
        .map_err(|_| Error::Allocation {
            amount: changes.edits.len(),
        })?;
    for (name, archive) in changes.edits {
        let compressed =
            encode_archive_for_publication(&archive, archive_limits, snappy_limits, budget)?;
        edited.push((name, compressed));
    }

    let catalog = &source.state.source;
    let mut inserted = Vec::new();
    inserted
        .try_reserve_exact(changes.insertions.len())
        .map_err(|_| Error::Allocation {
            amount: changes.insertions.len(),
        })?;
    for insertion in changes.insertions {
        let compressed = encode_archive_for_publication(
            &insertion.archive,
            archive_limits,
            snappy_limits,
            budget,
        )?;
        inserted.push((insertion.name, compressed, insertion.before));
    }

    let mut edits = Vec::new();
    edits
        .try_reserve_exact(edited.len())
        .map_err(|_| Error::Allocation {
            amount: edited.len(),
        })?;
    edits.extend(
        edited
            .iter()
            .map(|(name, data)| EntryEdit::new(name.as_str(), data.as_slice())),
    );

    let mut insertions = Vec::new();
    insertions
        .try_reserve_exact(inserted.len())
        .map_err(|_| Error::Allocation {
            amount: inserted.len(),
        })?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(catalog.package().len())
        .map_err(|_| Error::Allocation {
            amount: catalog.package().len(),
        })?;
    names.extend(catalog.package().iter().map(|entry| entry.name()));
    for (name, data, before) in &inserted {
        let insertion = match before.as_deref() {
            Some(before) => {
                let position = names
                    .iter()
                    .position(|candidate| *candidate == before)
                    .ok_or(Error::InvalidSource)?;
                EntryInsertion::at(position, name.as_str(), data.as_slice())
            },
            None => EntryInsertion::new(name.as_str(), data.as_slice()),
        };
        insertions.push(insertion);
    }
    let mut deletions = Vec::new();
    deletions
        .try_reserve_exact(changes.deletions.len())
        .map_err(|_| Error::Allocation {
            amount: changes.deletions.len(),
        })?;
    deletions.extend(changes.deletions.iter().map(String::as_str));

    let mut ingress = IngressBudget(budget);
    let prepared = catalog
        .__prepare_reassembly_with_changes_admitted(
            &insertions,
            &edits,
            &deletions,
            catalog.limits(),
            &mut ingress,
        )
        .map_err(map_candidate_ingress_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .output(requirements.output_bytes())
        .and_then(|_| budget.work(requirements.output_bytes()))
        .and_then(|_| budget.allocations(requirements.allocations()))
        .and_then(|_| budget.retained(requirements.retained_bytes()))
        .and_then(|_| budget.scratch(requirements.scratch_bytes()))
        .map_err(map_arrangement_error)?;
    // The prepared execution report is the exact ZIP-size authority. The
    // earlier projection only admits the archive staging phase; it must not
    // reject a valid insertion after preparation has already been admitted.
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate_plan =
        litchi_iwa_archive::SourceCatalog::__candidate_ingress(&output, catalog.limits())
            .map_err(map_archive_error)?;
    charge_candidate_plan(&candidate_plan, budget)?;
    budget
        .candidate_reopen(output.len())
        .map_err(map_arrangement_error)?;
    let target = Arc::<[u8]>::from(output);
    let candidate = reopen_series_candidate(source, target)?;
    Ok((candidate, touched_components))
}

fn package_object(package: &Package, identifier: u64) -> Option<(&str, &ArchiveObject)> {
    package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| {
            component
                .archive()
                .object(identifier)
                .map(|object| (component.name(), object))
        })
}

fn structural_scope_object(package: &Package, identifier: u64) -> bool {
    let Some((component, object)) = package_object(package, identifier) else {
        return false;
    };
    match component {
        "Index/Document.iwa" if identifier == 1 => object
            .messages
            .iter()
            .any(|message| message.type_ == ROOT_MESSAGE_TYPE),
        "Index/Metadata.iwa" => object
            .messages
            .iter()
            .any(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE),
        _ => false,
    }
}

fn add_planned_object(
    objects: &mut Vec<(Arc<str>, Vec<u64>)>,
    component: &str,
    identifier: u64,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    if let Some((_, identifiers)) = objects
        .iter_mut()
        .find(|(name, _)| name.as_ref() == component)
    {
        if !identifiers.contains(&identifier) {
            budget
                .preflight_allocations(1)
                .and_then(|_| budget.preflight_retained(size_of::<u64>()))
                .and_then(|_| budget.preflight_work(1))
                .map_err(map_arrangement_error)?;
            budget
                .allocations(1)
                .and_then(|_| budget.retained(size_of::<u64>()))
                .and_then(|_| budget.work(1))
                .map_err(map_arrangement_error)?;
            identifiers
                .try_reserve_exact(1)
                .map_err(|_| Error::Allocation { amount: 1 })?;
            identifiers.push(identifier);
        }
    } else {
        let retained = component
            .len()
            .checked_add(size_of::<(Arc<str>, Vec<u64>)>())
            .and_then(|value| value.checked_add(size_of::<u64>()))
            .ok_or(Error::InvalidSource)?;
        budget
            .preflight_allocations(3)
            .and_then(|_| budget.preflight_retained(retained))
            .and_then(|_| budget.preflight_work(1))
            .map_err(map_arrangement_error)?;
        budget
            .allocations(3)
            .and_then(|_| budget.retained(retained))
            .and_then(|_| budget.work(1))
            .map_err(map_arrangement_error)?;
        objects
            .try_reserve_exact(1)
            .map_err(|_| Error::Allocation { amount: 1 })?;
        let mut identifiers = Vec::new();
        identifiers
            .try_reserve_exact(1)
            .map_err(|_| Error::Allocation { amount: 1 })?;
        identifiers.push(identifier);
        objects.push((Arc::from(component), identifiers));
    }
    Ok(())
}

fn add_planned_member(
    members: &mut Vec<Arc<str>>,
    name: &str,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    if members.iter().any(|member| member.as_ref() == name) {
        return Ok(());
    }
    let retained = name
        .len()
        .checked_add(size_of::<Arc<str>>())
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_allocations(2)
        .and_then(|_| budget.preflight_retained(retained))
        .and_then(|_| budget.preflight_work(name.len()))
        .map_err(map_arrangement_error)?;
    budget
        .allocations(2)
        .and_then(|_| budget.retained(retained))
        .and_then(|_| budget.work(name.len()))
        .map_err(map_arrangement_error)?;
    members
        .try_reserve_exact(1)
        .map_err(|_| Error::Allocation { amount: 1 })?;
    members.push(Arc::from(name));
    Ok(())
}

fn mutation_plan(
    source: &Package,
    components: &[Arc<str>],
    changes: &style_graph::Changes,
    scope: &series_engine::MutationScope,
    budget: &mut ArrangementBudget,
) -> Result<MutationPlan, Error> {
    let member_allowed = |name: &str| {
        scope.allowed_members().iter().any(|member| member == name)
            || scope
                .planned_new_members()
                .iter()
                .any(|member| member == name)
            || scope
                .planned_removed_members()
                .iter()
                .any(|member| member == name)
    };
    let existing_allowed = |identifier: u64| {
        scope.allowed_objects().contains(&identifier)
            || scope.planned_removed_objects().contains(&identifier)
            || structural_scope_object(source, identifier)
    };
    let new_allowed = |identifier: u64| scope.planned_new_objects().contains(&identifier);
    let mut objects = Vec::new();
    let source_object_count = scope
        .allowed_objects()
        .len()
        .checked_add(scope.planned_removed_objects().len())
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_allocations(1)
        .and_then(|_| {
            budget.preflight_retained(source_object_count.saturating_mul(size_of::<u64>()))
        })
        .map_err(map_arrangement_error)?;
    budget
        .allocations(1)
        .and_then(|_| budget.retained(source_object_count.saturating_mul(size_of::<u64>())))
        .map_err(map_arrangement_error)?;
    let mut source_objects = Vec::new();
    source_objects
        .try_reserve_exact(source_object_count)
        .map_err(|_| Error::Allocation {
            amount: source_object_count,
        })?;
    source_objects.extend_from_slice(scope.allowed_objects());
    source_objects.extend_from_slice(scope.planned_removed_objects());
    for identifier in source_objects {
        let Some((component, _)) = package_object(source, identifier) else {
            return Err(Error::Verification);
        };
        add_planned_object(&mut objects, component, identifier, budget)?;
    }
    for (name, archive) in &changes.edits {
        if !member_allowed(name) {
            return Err(Error::Verification);
        }
        for object in &archive.objects {
            let identifier = object.archive_info.identifier.ok_or(Error::Verification)?;
            if let Some((_, source_object)) = package_object(source, identifier)
                && same_object_metadata(source_object, object)
            {
                continue;
            }
            if package_object(source, identifier).is_some() {
                if !existing_allowed(identifier) {
                    return Err(Error::Verification);
                }
                if structural_scope_object(source, identifier) {
                    add_planned_object(&mut objects, name, identifier, budget)?;
                }
            } else if !new_allowed(identifier)
                || !object
                    .messages
                    .iter()
                    .all(|message| scope.allowed_message_types().contains(&message.type_))
            {
                return Err(Error::Verification);
            } else {
                add_planned_object(&mut objects, name, identifier, budget)?;
            }
        }
    }
    for insertion in &changes.insertions {
        if !member_allowed(&insertion.name) {
            return Err(Error::Verification);
        }
        for object in &insertion.archive.objects {
            let identifier = object.archive_info.identifier.ok_or(Error::Verification)?;
            if !new_allowed(identifier)
                || !object
                    .messages
                    .iter()
                    .all(|message| scope.allowed_message_types().contains(&message.type_))
            {
                return Err(Error::Verification);
            }
            add_planned_object(&mut objects, &insertion.name, identifier, budget)?;
        }
    }
    for name in components.iter().map(Arc::as_ref) {
        if !member_allowed(name) {
            return Err(Error::Verification);
        }
    }
    let mut scoped_components = Vec::new();
    for name in components {
        add_planned_member(&mut scoped_components, name, budget)?;
    }
    let mut planned_new_members = Vec::new();
    for name in scope.planned_new_members() {
        add_planned_member(&mut planned_new_members, name, budget)?;
    }
    let mut planned_removed_members = Vec::new();
    for name in scope.planned_removed_members() {
        add_planned_member(&mut planned_removed_members, name, budget)?;
    }
    budget
        .preflight_allocations(2)
        .and_then(|_| {
            budget.preflight_retained(
                scope
                    .allowed_message_types()
                    .len()
                    .saturating_mul(size_of::<u32>()),
            )
        })
        .and_then(|_| {
            budget.preflight_retained(
                scope
                    .allowed_fields()
                    .len()
                    .saturating_mul(size_of::<u32>()),
            )
        })
        .map_err(map_arrangement_error)?;
    budget.allocations(2).map_err(map_arrangement_error)?;
    let mut allowed_message_types = Vec::new();
    allowed_message_types
        .try_reserve_exact(scope.allowed_message_types().len())
        .map_err(|_| Error::Allocation {
            amount: scope.allowed_message_types().len(),
        })?;
    allowed_message_types.extend_from_slice(scope.allowed_message_types());
    let mut allowed_fields = Vec::new();
    allowed_fields
        .try_reserve_exact(scope.allowed_fields().len())
        .map_err(|_| Error::Allocation {
            amount: scope.allowed_fields().len(),
        })?;
    allowed_fields.extend_from_slice(scope.allowed_fields());
    budget
        .preflight_allocations(1)
        .and_then(|_| {
            budget.preflight_retained(
                scope
                    .allowed_chart_fields()
                    .len()
                    .saturating_mul(size_of::<u32>()),
            )
        })
        .map_err(map_arrangement_error)?;
    budget.allocations(1).map_err(map_arrangement_error)?;
    let mut chart_fields = Vec::new();
    chart_fields
        .try_reserve_exact(scope.allowed_chart_fields().len())
        .map_err(|_| Error::Allocation {
            amount: scope.allowed_chart_fields().len(),
        })?;
    chart_fields.extend_from_slice(scope.allowed_chart_fields());
    Ok(MutationPlan {
        components: scoped_components,
        objects,
        allowed_message_types,
        allowed_fields,
        chart_fields,
        planned_new_members,
        planned_removed_members,
    })
}

fn planned_contains(plan: &MutationPlan, component: &str, identifier: u64) -> bool {
    plan.objects
        .iter()
        .any(|(name, ids)| name.as_ref() == component && ids.contains(&identifier))
}

impl MutationPlan {
    fn allowed_object(&self, identifier: u64) -> bool {
        self.objects
            .iter()
            .any(|(_, identifiers)| identifiers.contains(&identifier))
    }

    fn planned_new_member(&self, name: &str) -> bool {
        self.planned_new_members
            .iter()
            .any(|member| member.as_ref() == name)
    }

    fn planned_removed_member(&self, name: &str) -> bool {
        self.planned_removed_members
            .iter()
            .any(|member| member.as_ref() == name)
    }
}

fn unique_wire_field<'a>(
    view: &'a WireView<'a>,
    number: u32,
) -> Result<Option<WireFieldView<'a>>, Error> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if found.replace(field).is_some() {
            return Err(Error::Verification);
        }
    }
    Ok(found)
}

fn same_wire_fields_except(
    source: &WireView<'_>,
    candidate: &WireView<'_>,
    excluded: &[u32],
    allowed_fields: &[u32],
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let mut source_fields = source.fields().filter(|field| {
        !(excluded.contains(&field.number()) && allowed_fields.contains(&field.number()))
    });
    let mut candidate_fields = candidate.fields().filter(|field| {
        !(excluded.contains(&field.number()) && allowed_fields.contains(&field.number()))
    });
    loop {
        match (source_fields.next(), candidate_fields.next()) {
            (None, None) => return Ok(()),
            (Some(left), Some(right)) => {
                budget
                    .work(left.raw().len().saturating_add(right.raw().len()))
                    .map_err(map_arrangement_error)?;
                if left.raw() != right.raw() {
                    return Err(Error::Verification);
                }
            },
            _ => {
                return Err(Error::Verification);
            },
        }
    }
}

fn verify_message_unknowns(
    message_type: u32,
    source: &[u8],
    candidate: &[u8],
    plan: &MutationPlan,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let source_view = budget.parse(source, 1).map_err(map_arrangement_error)?;
    let candidate_view = budget.parse(candidate, 1).map_err(map_arrangement_error)?;
    let outer_excluded: &[u32] = match message_type {
        CHART_MESSAGE_TYPE | CHART_STYLE_MESSAGE_TYPES_FIRST..=CHART_STYLE_MESSAGE_TYPES_LAST => {
            &[STYLE_EXTENSION_FIELD]
        },
        STYLESHEET_MESSAGE_TYPE => &[1, 5],
        // A legend typography variation patched in place: its override
        // count and character properties.
        PARAGRAPH_STYLE_MESSAGE_TYPE => &[10, 11],
        OBJECT_CONTAINER_MESSAGE_TYPE => &[2],
        PACKAGE_METADATA_MESSAGE_TYPE => &[1, 3, 8, 10, 12],
        ROOT_MESSAGE_TYPE => &[7, 9],
        _ => return Err(Error::Verification),
    };
    same_wire_fields_except(
        &source_view,
        &candidate_view,
        outer_excluded,
        &plan.allowed_fields,
        budget,
    )?;
    if message_type == PACKAGE_METADATA_MESSAGE_TYPE {
        // Component registrations match by identifier: a planned new member
        // registers one more component, a planned removed member one less;
        // every other registration keeps its bytes outside its object-UUID
        // and external-reference lists.
        let mut keyed = |view: &WireView<'_>| -> Result<Vec<(Vec<u8>, Vec<u8>)>, Error> {
            let mut entries = Vec::new();
            for field in view.fields().filter(|field| field.number() == 3) {
                let entry = budget
                    .parse(field.payload(), 2)
                    .map_err(map_arrangement_error)?;
                let identifier = unique_wire_field(&entry, 1)?
                    .ok_or(Error::Verification)?
                    .payload()
                    .to_vec();
                budget
                    .work(field.payload().len())
                    .map_err(map_arrangement_error)?;
                entries
                    .try_reserve(1)
                    .map_err(|_| Error::Allocation { amount: 1 })?;
                entries.push((identifier, field.payload().to_vec()));
            }
            Ok(entries)
        };
        let source_entries = keyed(&source_view)?;
        let candidate_entries = keyed(&candidate_view)?;
        let mut matched = 0_usize;
        let mut removed = 0_usize;
        for (identifier, source_entry) in &source_entries {
            let Some((_, candidate_entry)) = candidate_entries
                .iter()
                .find(|(candidate, _)| candidate == identifier)
            else {
                removed = removed.saturating_add(1);
                continue;
            };
            matched = matched.saturating_add(1);
            let source_view = budget
                .parse(source_entry, 2)
                .map_err(map_arrangement_error)?;
            let candidate_view = budget
                .parse(candidate_entry, 2)
                .map_err(map_arrangement_error)?;
            same_wire_fields_except(
                &source_view,
                &candidate_view,
                &[6, 11],
                &plan.allowed_fields,
                budget,
            )?;
        }
        let added = candidate_entries.len().saturating_sub(matched);
        if removed > plan.planned_removed_members.len() || added > plan.planned_new_members.len() {
            return Err(Error::Verification);
        }
        return Ok(());
    }
    if !outer_excluded.contains(&STYLE_EXTENSION_FIELD) {
        return Ok(());
    }
    match (
        unique_wire_field(&source_view, STYLE_EXTENSION_FIELD)?,
        unique_wire_field(&candidate_view, STYLE_EXTENSION_FIELD)?,
    ) {
        (None, None) => Ok(()),
        (Some(source), Some(candidate)) => {
            let source_view = budget
                .parse(source.payload(), 2)
                .map_err(map_arrangement_error)?;
            let candidate_view = budget
                .parse(candidate.payload(), 2)
                .map_err(map_arrangement_error)?;
            let excluded: &[u32] = match message_type {
                CHART_MESSAGE_TYPE => &plan.chart_fields,
                _ => &plan.allowed_fields,
            };
            same_wire_fields_except(
                &source_view,
                &candidate_view,
                excluded,
                &plan.allowed_fields,
                budget,
            )
        },
        _ => Err(Error::Verification),
    }
}

fn verify_object_messages(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    plan: &MutationPlan,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    if source.messages.len() != candidate.messages.len() {
        return Err(Error::Verification);
    }
    for (before, after) in source.messages.iter().zip(&candidate.messages) {
        if before.type_ != after.type_ {
            return Err(Error::Verification);
        }
        if before.data == after.data {
            continue;
        }
        if !plan.allowed_message_types.contains(&before.type_) {
            return Err(Error::Verification);
        }
        verify_message_unknowns(before.type_, &before.data, &after.data, plan, budget)?;
    }
    Ok(())
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    plan: &MutationPlan,
    budget: &mut ArrangementBudget,
) -> Result<(), Error> {
    let changed = |name: &str| plan.components.iter().any(|item| item.as_ref() == name);
    let source_entries = source.state.source.package();
    let candidate_entries = candidate.state.source.package();
    for before in source_entries.iter() {
        let after = candidate_entries
            .iter()
            .find(|item| item.name() == before.name());
        let Some(after) = after else {
            if !plan.planned_removed_member(before.name()) {
                return Err(Error::Verification);
            }
            continue;
        };
        budget
            .work(entry_work(before)?.saturating_add(entry_work(after)?))
            .map_err(map_arrangement_error)?;
        if before.name() != after.name()
            || before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
            || before.metadata().local() != after.metadata().local()
            || before.metadata().central() != after.metadata().central()
        {
            return Err(Error::Verification);
        }
        if changed(before.name()) {
            if !zip_local_header_preserved(
                before.raw_record().local_record(),
                after.raw_record().local_record(),
            ) || !zip_record_equal_outside(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
                &[(16, 28), (42, 46)],
            ) {
                return Err(Error::Verification);
            }
        } else if before.data() != after.data()
            || before.raw_record().local_record() != after.raw_record().local_record()
            || before.raw_record().compressed_data() != after.raw_record().compressed_data()
            || !zip_record_equal_outside(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
                &[(42, 46)],
            )
        {
            return Err(Error::Verification);
        }
    }
    for after in candidate_entries.iter() {
        if source_entries
            .iter()
            .all(|before| before.name() != after.name())
            && !plan.planned_new_member(after.name())
        {
            return Err(Error::Verification);
        }
    }

    for component in source.state.source.components().iter() {
        let name = component.name();
        let Some(after_component) = candidate.state.source.components().get(name) else {
            if plan.planned_removed_member(name) {
                continue;
            }
            return Err(Error::Verification);
        };
        if !changed(name) {
            if component.archive().objects.len() != after_component.archive().objects.len()
                || component
                    .archive()
                    .objects
                    .iter()
                    .zip(&after_component.archive().objects)
                    .any(|(before, after)| !same_object_metadata(before, after))
            {
                return Err(Error::Verification);
            }
            continue;
        }
        for before in &component.archive().objects {
            let identifier = before.archive_info.identifier.ok_or(Error::Verification)?;
            let Some(after) = after_component.archive().object(identifier) else {
                if !planned_contains(plan, name, identifier) {
                    return Err(Error::Verification);
                }
                continue;
            };
            budget
                .work(object_work(before)?.saturating_add(object_work(after)?))
                .map_err(map_arrangement_error)?;
            let planned = planned_contains(plan, name, identifier);
            let unchanged = same_object_metadata(before, after);
            if !planned && !unchanged {
                return Err(Error::Verification);
            }
            if planned && !same_archive_info_except_references(before, after, true, plan) {
                return Err(Error::Verification);
            }
            if planned && !before.same_source_headers(after) {
                return Err(Error::Verification);
            }
            if !unchanged {
                verify_object_messages(before, after, plan, budget)?;
            }
        }
        for after in &after_component.archive().objects {
            let identifier = after.archive_info.identifier.ok_or(Error::Verification)?;
            if component.archive().object(identifier).is_none()
                && !planned_contains(plan, name, identifier)
            {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

fn entry_work(entry: &Entry) -> Result<usize, Error> {
    entry
        .name()
        .len()
        .checked_add(entry.raw_name().len())
        .and_then(|amount| amount.checked_add(entry.data().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().local_record().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().central_directory_record().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().compressed_data().len()))
        .ok_or(Error::InvalidSource)
}

fn object_work(object: &ArchiveObject) -> Result<usize, Error> {
    let mut work = size_of::<ArchiveObject>()
        .checked_add(object.messages.len())
        .and_then(|amount| amount.checked_add(object.archive_info.message_infos.len()))
        .ok_or(Error::InvalidSource)?;
    for message in &object.messages {
        work = work
            .checked_add(size_of::<RawMessage>())
            .and_then(|amount| amount.checked_add(message.data.len()))
            .ok_or(Error::InvalidSource)?;
    }
    Ok(work)
}

fn same_object_metadata(before: &ArchiveObject, after: &ArchiveObject) -> bool {
    before.same_content_ignoring_offsets(after)
}

fn same_archive_info_except_references(
    before: &ArchiveObject,
    after: &ArchiveObject,
    allow_message_lengths: bool,
    plan: &MutationPlan,
) -> bool {
    if before.archive_info.identifier != after.archive_info.identifier
        || before.archive_info.should_merge != after.archive_info.should_merge
        || before.archive_info.message_infos.len() != after.archive_info.message_infos.len()
    {
        return false;
    }
    before
        .archive_info
        .message_infos
        .iter()
        .zip(&after.archive_info.message_infos)
        .all(|(left, right)| {
            left.type_ == right.type_
                && left.versions == right.versions
                && (allow_message_lengths || left.length == right.length)
                && left.base_message_index == right.base_message_index
                && left.diff_merge_version == right.diff_merge_version
                && left.diff_field_path == right.diff_field_path
                && left.fields_to_remove == right.fields_to_remove
                && left.diff_read_version == right.diff_read_version
                && reference_transition_allowed(
                    &left.object_references,
                    &right.object_references,
                    plan,
                )
                && left.data_references == right.data_references
                && left.field_infos.len() == right.field_infos.len()
                && left
                    .field_infos
                    .iter()
                    .zip(&right.field_infos)
                    .all(|(left, right)| {
                        left.path == right.path
                            && left.r#type == right.r#type
                            && left.unknown_field_rule == right.unknown_field_rule
                            && left.known_field_rule == right.known_field_rule
                            && left.known_field_version == right.known_field_version
                            && left.known_field_feature_identifier
                                == right.known_field_feature_identifier
                            && reference_transition_allowed(
                                &left.object_references,
                                &right.object_references,
                                plan,
                            )
                            && left.data_references == right.data_references
                    })
        })
}

fn reference_transition_allowed(before: &[u64], after: &[u64], plan: &MutationPlan) -> bool {
    let count =
        |values: &[u64], identifier| values.iter().filter(|value| **value == identifier).count();
    before.iter().chain(after).all(|identifier| {
        let left = count(before, *identifier);
        let right = count(after, *identifier);
        left == right || plan.allowed_object(*identifier)
    })
}

fn zip_record_equal_outside(before: &[u8], after: &[u8], excluded: &[(usize, usize)]) -> bool {
    if before.len() != after.len() {
        return false;
    }
    let mut cursor = 0;
    for &(start, end) in excluded {
        if start < cursor
            || end < start
            || end > before.len()
            || before[cursor..start] != after[cursor..start]
        {
            return false;
        }
        cursor = end;
    }
    before[cursor..] == after[cursor..]
}

fn zip_local_header_preserved(before: &[u8], after: &[u8]) -> bool {
    let Some(before_header) = zip_local_header_len(before) else {
        return false;
    };
    let Some(after_header) = zip_local_header_len(after) else {
        return false;
    };
    before_header == after_header
        && zip_record_equal_outside(
            &before[..before_header],
            &after[..after_header],
            &[(14, 26)],
        )
}

fn zip_local_header_len(record: &[u8]) -> Option<usize> {
    if record.len() < 30 || record.get(..4) != Some(b"PK\x03\x04") {
        return None;
    }
    let name = usize::from(u16::from_le_bytes([record[26], record[27]]));
    let extra = usize::from(u16::from_le_bytes([record[28], record[29]]));
    30usize
        .checked_add(name)?
        .checked_add(extra)
        .filter(|length| *length <= record.len())
}

// ---------------------------------------------------------------------------
// The shared core of the Pages chart-property families (series, style,
// axis): one transaction budget, the bounded workspace, exact-source
// publication under the engine's mutation scope and candidate reopen.
// ---------------------------------------------------------------------------

/// The transaction budget of one chart-property operation.
pub(super) struct Ledger {
    budget: ArrangementBudget,
}

/// A published candidate of the shared core.
pub(super) struct Published {
    pub(super) package: Package,
    pub(super) touched: usize,
    pub(super) plan: MutationPlan,
    pub(super) components: Arc<[Arc<str>]>,
}

impl Ledger {
    /// Admit one operation on `package`.
    pub(super) fn new(package: &Package) -> Result<Self, Error> {
        Ok(Self {
            budget: ArrangementBudget::new(package).map_err(map_arrangement_error)?,
        })
    }

    /// Resolve the selected chart, its kind and its series count.
    pub(super) fn target(
        &mut self,
        package: &Package,
        selector: BodyChartSelector,
    ) -> Result<Target, Error> {
        target_with_metadata(package, selector, &mut self.budget)
    }

    /// Run one engine read over a bounded workspace of `package`.
    pub(super) fn read<T>(
        &mut self,
        package: &Package,
        read: impl FnOnce(
            &Workspace<'_>,
            series_engine::Limits,
        ) -> series_engine::SymbolResult<(T, series_engine::Report)>,
    ) -> Result<T, Error> {
        let budget = &mut self.budget;
        admit_workspace(package, budget)?;
        let components = PackageComponents(package);
        let limits = archive_limits(package)?;
        let census_limits = construction_limits(budget)?;
        let construction = Workspace::construction_requirements(&components, census_limits)
            .map_err(map_construction_error)?;
        charge_workspace_construction(budget, construction)?;
        let workspace = Workspace::new_bounded(&components, limits, construction, census_limits)
            .map_err(map_construction_error)?;
        let engine_limits = engine_limits(package, budget)?;
        let (value, report) = read(&workspace, engine_limits).map_err(map_symbol_error)?;
        charge_engine_report(budget, &report)?;
        Ok(value)
    }

    /// Run one engine write over a bounded workspace of `source` and publish
    /// its changes under the engine's mutation scope; `None` when nothing
    /// changes.
    pub(super) fn write(
        &mut self,
        source: &Package,
        write: impl FnOnce(
            &mut Workspace<'_>,
            series_engine::Limits,
        ) -> series_engine::SymbolResult<(bool, series_engine::Report)>,
    ) -> Result<Option<Published>, Error> {
        if !source.state.source.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        let budget = &mut self.budget;
        admit_workspace(source, budget)?;
        let components = PackageComponents(source);
        let limits = archive_limits(source)?;
        let census_limits = construction_limits(budget)?;
        let construction = Workspace::construction_requirements(&components, census_limits)
            .map_err(map_construction_error)?;
        charge_workspace_construction(budget, construction)?;
        let mut workspace =
            Workspace::new_bounded(&components, limits, construction, census_limits)
                .map_err(map_construction_error)?;
        let engine_limits = engine_limits(source, budget)?;
        let (changed, report) = write(&mut workspace, engine_limits).map_err(map_symbol_error)?;
        if !changed {
            charge_engine_report(budget, &report)?;
            return Ok(None);
        }
        let scope = report.scope().cloned().ok_or(Error::Verification)?;
        charge_engine_report(budget, &report)?;
        let finish_requirements = workspace.finish_requirements().map_err(map_finish_error)?;
        preflight_finish(budget, finish_requirements)?;
        let finish_limits = finish_limits(budget)?;
        let (changes, finish_report) = workspace
            .finish_bounded(finish_limits)
            .map_err(map_finish_error)?;
        charge_finish_report(budget, finish_report)?;
        if changes.is_empty() {
            return Ok(None);
        }
        let name_count = changes
            .edits
            .len()
            .checked_add(changes.insertions.len())
            .and_then(|value| value.checked_add(changes.deletions.len()))
            .ok_or(Error::InvalidSource)?;
        let name_bytes = changes
            .edits
            .iter()
            .map(|(name, _)| name.len())
            .chain(
                changes
                    .insertions
                    .iter()
                    .map(|insertion| insertion.name.len()),
            )
            .chain(changes.deletions.iter().map(String::len))
            .try_fold(0usize, |total, bytes| total.checked_add(bytes))
            .ok_or(Error::InvalidSource)?;
        budget
            .preflight_allocations(name_count)
            .and_then(|_| budget.preflight_retained(name_bytes))
            .and_then(|_| budget.preflight_work(name_bytes))
            .map_err(map_arrangement_error)?;
        budget
            .allocations(name_count)
            .and_then(|_| budget.retained(name_bytes))
            .and_then(|_| budget.work(name_bytes))
            .map_err(map_arrangement_error)?;
        let mut names = Vec::new();
        names
            .try_reserve_exact(name_count)
            .map_err(|_| Error::Allocation { amount: name_count })?;
        names.extend(
            changes
                .edits
                .iter()
                .map(|(name, _)| Arc::<str>::from(name.as_str()))
                .chain(
                    changes
                        .insertions
                        .iter()
                        .map(|insertion| Arc::<str>::from(insertion.name.as_str())),
                )
                .chain(
                    changes
                        .deletions
                        .iter()
                        .map(|name| Arc::<str>::from(name.as_str())),
                ),
        );
        let plan = mutation_plan(source, &names, &changes, &scope, budget)?;
        let (candidate, touched) = publish_changes(source, changes, budget)?;
        verify_locality(source, &candidate, &plan, budget)?;
        Ok(Some(Published {
            package: candidate,
            touched,
            plan,
            components: names.into(),
        }))
    }

    /// Reopen the exact target of a patch over `package` and verify its
    /// locality against `plan`.
    pub(super) fn reopen(
        &mut self,
        package: &Package,
        artifacts: &OwnedExactArtifacts,
        plan: &MutationPlan,
    ) -> Result<Package, Error> {
        let budget = &mut self.budget;
        admit_workspace(package, budget)?;
        let source = package.state.source.shared_source();
        let source_owner = SharedBytes::from_shared_slice(Arc::clone(&source));
        budget
            .work(
                source
                    .len()
                    .checked_add(artifacts.source_owner().as_ref().len())
                    .ok_or(Error::InvalidSource)?,
            )
            .map_err(map_arrangement_error)?;
        if !artifacts.authorizes_owner(&source_owner) {
            return Err(Error::PatchConflict);
        }
        if !package.state.source.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        let target_owner = artifacts.target_owner();
        budget
            .preflight_output(target_owner.as_ref().len())
            .map_err(map_arrangement_error)?;
        budget
            .preflight_scratch(target_owner.as_ref().len())
            .map_err(map_arrangement_error)?;
        let candidate_plan = litchi_iwa_archive::SourceCatalog::__candidate_ingress(
            target_owner.as_ref(),
            package.state.source.limits(),
        )
        .map_err(map_archive_error)?;
        charge_candidate_plan(&candidate_plan, budget)?;
        budget
            .output(target_owner.as_ref().len())
            .and_then(|_| budget.scratch(target_owner.as_ref().len()))
            .map_err(map_arrangement_error)?;
        budget
            .candidate_reopen(target_owner.as_ref().len())
            .map_err(map_arrangement_error)?;
        let candidate_bytes = Arc::<[u8]>::from(target_owner.as_ref());
        let candidate = reopen_series_candidate(package, candidate_bytes)?;
        verify_locality(package, &candidate, plan, budget)?;
        Ok(candidate)
    }
}

/// The exact artifacts of a commit from `source` to `target`.
pub(super) fn artifacts(source: &Package, target: &Package) -> OwnedExactArtifacts {
    OwnedExactArtifacts::new(
        SharedBytes::from_shared_slice(source.state.source.shared_source()),
        SharedBytes::from_shared_slice(target.state.source.shared_source()),
    )
}

/// The exact no-op artifacts of `source`.
pub(super) fn noop_artifacts(source: &Package) -> OwnedExactArtifacts {
    let owner = SharedBytes::from_shared_slice(source.state.source.shared_source());
    OwnedExactArtifacts::new(owner.clone(), owner)
}

/// Whether `artifacts` were produced from `package`'s exact bytes.
pub(super) fn authorizes(package: &Package, artifacts: &OwnedExactArtifacts) -> bool {
    artifacts.authorizes_owner(&SharedBytes::from_shared_slice(
        package.state.source.shared_source(),
    ))
}

/// Inverse a publication plan.
pub(super) fn inverse_plan(plan: &MutationPlan) -> MutationPlan {
    plan.inverse()
}

impl Diagnostics {
    pub(super) const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
        }
    }

    pub(super) const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
        }
    }
}

// ---------------------------------------------------------------------------
// The per-series family.
// ---------------------------------------------------------------------------

/// One staged per-series operation.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    Set(Property),
    Reset(PropertyKind, usize),
}

/// A staged immutable body-chart per-series edit.
pub struct Edit<'a> {
    source: &'a Package,
    selector: BodyChartSelector,
    change: Option<Staged>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("selector", &self.selector)
            .field("staged", &self.change.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage `property` (one value per visible series), replacing any
    /// staged operation.
    ///
    /// # Errors
    ///
    /// [`Refusal::SeriesCount`] for a value naming no series.
    pub fn set(mut self, property: Property) -> Result<Self, Error> {
        if property.is_empty() {
            return Err(Refusal::SeriesCount {
                expected: 1,
                given: 0,
            }
            .into());
        }
        self.change = Some(Staged::Set(property));
        Ok(self)
    }

    /// Stage removing series `series`' own value of a series-style property
    /// (strokes, fills, symbol fills, symbol outlines), so it inherits its
    /// theme style's again (the host's `reset_body_chart_series_*`).
    ///
    /// # Errors
    ///
    /// [`Refusal::NotResettable`] for the other properties.
    pub fn reset(mut self, kind: PropertyKind, series: usize) -> Result<Self, Error> {
        if kind.style().is_none() {
            return Err(Refusal::NotResettable.into());
        }
        self.change = Some(Staged::Reset(kind, series));
        Ok(self)
    }

    /// Publish the staged exact-source transaction and reopen its candidate.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::UnsupportedKind`],
    /// [`Refusal::SeriesCount`], [`Refusal::SeriesNotFound`],
    /// [`Refusal::SharedStyle`], [`Refusal::ImageFill`], source, limit and
    /// verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        let staged = self
            .change
            .ok_or_else(|| Error::from(Refusal::NoStagedOperation))?;
        let source = self.source;
        let mut ledger = Ledger::new(source)?;
        let target = ledger.target(source, self.selector)?;
        let kind = match &staged {
            Staged::Set(property) => property.kind(),
            Staged::Reset(kind, _) => *kind,
        };
        let before = read_property(&mut ledger, source, &target, kind)?;
        let published = match &staged {
            Staged::Set(property) => {
                if property.len() != target.series {
                    return Err(Refusal::SeriesCount {
                        expected: target.series,
                        given: property.len(),
                    }
                    .into());
                }
                if before == *property {
                    None
                } else {
                    ledger.write(source, |workspace, limits| {
                        write_property(workspace, &target, property, limits)
                    })?
                }
            },
            Staged::Reset(kind, series) => {
                if *series >= target.series {
                    return Err(Refusal::SeriesNotFound { series: *series }.into());
                }
                let style_kind = kind.style().ok_or(Refusal::NotResettable)?;
                let mut values = before.style_values().ok_or(Error::InvalidSource)?;
                values[*series] = None;
                ledger.write(source, |workspace, limits| {
                    style_engine::write(
                        workspace,
                        target.chart.drawable_identifier.get(),
                        target.kind,
                        style_kind,
                        &values,
                        limits,
                    )
                })?
            },
        };
        let Some(published) = published else {
            return Ok(Commit {
                package: source.clone(),
                patch: Patch {
                    artifacts: noop_artifacts(source),
                    position: target.chart.position,
                    target,
                    before: before.clone(),
                    after: before,
                    changed_components: Arc::from([]),
                    plan: MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        };
        let candidate_target = ledger.target(
            &published.package,
            BodyChartSelector::position(target.chart.position),
        )?;
        if candidate_target.kind != target.kind || candidate_target.series != target.series {
            return Err(Error::Verification);
        }
        let after = read_property(&mut ledger, &published.package, &candidate_target, kind)?;
        if let Staged::Set(property) = &staged
            && after != *property
        {
            return Err(Error::Verification);
        }
        Ok(Commit {
            patch: Patch {
                artifacts: artifacts(source, &published.package),
                position: target.chart.position,
                target,
                before,
                after,
                changed_components: published.components,
                plan: published.plan,
            },
            package: published.package,
            diagnostics: Diagnostics::published(published.touched),
        })
    }
}

fn unsupported() -> series_engine::SymbolError {
    series_engine::SymbolError::Refused(series_engine::Refusal::UnsupportedKind)
}

fn connection_lines(
    kind: ChartKind,
) -> series_engine::SymbolResult<series_property::ConnectionLines> {
    ConnectionLineFamily::for_chart_kind(kind)
        .map(series_property::ConnectionLines)
        .ok_or_else(unsupported)
}

fn symbol_family(kind: ChartKind) -> series_engine::SymbolResult<SymbolFamily> {
    SymbolFamily::for_chart_kind(kind).ok_or_else(unsupported)
}

fn value_labels(kind: ChartKind) -> series_engine::SymbolResult<series_property::ValueLabels> {
    series_property::ValueLabels::for_chart_kind(kind).ok_or_else(unsupported)
}

fn value_label_affixes(
    kind: ChartKind,
) -> series_engine::SymbolResult<series_property::ValueLabelAffixes> {
    series_property::ValueLabelAffixes::for_chart_kind(kind).ok_or_else(unsupported)
}

fn value_label_formats(
    kind: ChartKind,
) -> series_engine::SymbolResult<series_property::ValueLabelNumberFormats> {
    series_property::ValueLabelNumberFormats::for_chart_kind(kind).ok_or_else(unsupported)
}

/// Read per-series property `kind` of the target chart.
fn read_property(
    ledger: &mut Ledger,
    package: &Package,
    target: &Target,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let chart = target.chart.drawable_identifier.get();
    let count = target.series;
    let chart_kind = target.kind;
    ledger.read(package, |workspace, limits| {
        if let Some(style_kind) = kind.style() {
            let (values, report) =
                style_engine::read(workspace, chart, chart_kind, count, style_kind, limits)?;
            return Ok((from_style_values(kind, values)?, report));
        }
        macro_rules! generic {
            ($variant:ident, $property:expr) => {{
                let (values, report) =
                    series_engine::read(workspace, chart, count, &$property, limits)?;
                (Property::$variant(values), report)
            }};
        }
        Ok(match kind {
            PropertyKind::Symbols => {
                let (values, report) = series_engine::symbols(
                    workspace,
                    chart,
                    symbol_family(chart_kind)?,
                    count,
                    limits,
                )?;
                (Property::Symbols(values), report)
            },
            PropertyKind::ConnectionLines => {
                generic!(ConnectionLines, connection_lines(chart_kind)?)
            },
            PropertyKind::WedgeExplosions => {
                generic!(WedgeExplosions, series_property::WedgeExplosions)
            },
            PropertyKind::LabelDistances => {
                generic!(LabelDistances, series_property::LabelDistances)
            },
            PropertyKind::Trendlines => generic!(Trendlines, series_property::Trendlines),
            PropertyKind::ErrorBars => generic!(ErrorBars, series_property::ErrorBars),
            PropertyKind::PieLabels => generic!(PieLabels, series_property::PieLabels),
            PropertyKind::ValueLabels => generic!(ValueLabels, value_labels(chart_kind)?),
            PropertyKind::ValueLabelAffixes => {
                generic!(ValueLabelAffixes, value_label_affixes(chart_kind)?)
            },
            PropertyKind::ValueLabelNumberFormats => {
                generic!(ValueLabelNumberFormats, value_label_formats(chart_kind)?)
            },
            PropertyKind::Strokes
            | PropertyKind::Fills
            | PropertyKind::SymbolFills
            | PropertyKind::SymbolOutlines => {
                return Err(series_engine::SymbolError::Invalid("a series-style kind"));
            },
        })
    })
}

/// Write per-series `property` to the target chart.
fn write_property(
    workspace: &mut Workspace<'_>,
    target: &Target,
    property: &Property,
    limits: series_engine::Limits,
) -> series_engine::SymbolResult<(bool, series_engine::Report)> {
    let chart = target.chart.drawable_identifier.get();
    let chart_kind = target.kind;
    if let (Some(style_kind), Some(values)) = (property.kind().style(), property.style_values()) {
        return style_engine::write(workspace, chart, chart_kind, style_kind, &values, limits);
    }
    match property {
        Property::Symbols(values) => {
            series_engine::set_symbols(workspace, chart, symbol_family(chart_kind)?, values, limits)
        },
        Property::ConnectionLines(values) => series_engine::write(
            workspace,
            chart,
            &connection_lines(chart_kind)?,
            values,
            limits,
        ),
        Property::WedgeExplosions(values) => series_engine::write(
            workspace,
            chart,
            &series_property::WedgeExplosions,
            values,
            limits,
        ),
        Property::LabelDistances(values) => series_engine::write(
            workspace,
            chart,
            &series_property::LabelDistances,
            values,
            limits,
        ),
        Property::Trendlines(values) => series_engine::write(
            workspace,
            chart,
            &series_property::Trendlines,
            values,
            limits,
        ),
        Property::ErrorBars(values) => series_engine::write(
            workspace,
            chart,
            &series_property::ErrorBars,
            values,
            limits,
        ),
        Property::PieLabels(values) => series_engine::write(
            workspace,
            chart,
            &series_property::PieLabels,
            values,
            limits,
        ),
        Property::ValueLabels(values) => {
            series_engine::write(workspace, chart, &value_labels(chart_kind)?, values, limits)
        },
        Property::ValueLabelAffixes(values) => series_engine::write(
            workspace,
            chart,
            &value_label_affixes(chart_kind)?,
            values,
            limits,
        ),
        Property::ValueLabelNumberFormats(values) => series_engine::write(
            workspace,
            chart,
            &value_label_formats(chart_kind)?,
            values,
            limits,
        ),
        Property::Strokes(_)
        | Property::Fills(_)
        | Property::SymbolFills(_)
        | Property::SymbolOutlines(_) => {
            Err(series_engine::SymbolError::Invalid("a series-style value"))
        },
    }
}

fn from_style_values(
    kind: PropertyKind,
    values: Vec<style_engine::Value>,
) -> series_engine::SymbolResult<Property> {
    let invalid =
        || series_engine::SymbolError::Invalid("a series-style value of another property");
    Ok(match kind {
        PropertyKind::Strokes | PropertyKind::SymbolOutlines => {
            let strokes = values
                .into_iter()
                .map(|value| match value {
                    style_engine::Value::Stroke(stroke) => Ok(stroke),
                    _ => Err(invalid()),
                })
                .collect::<series_engine::SymbolResult<Vec<_>>>()?;
            if kind == PropertyKind::Strokes {
                Property::Strokes(strokes)
            } else {
                Property::SymbolOutlines(strokes)
            }
        },
        PropertyKind::Fills => Property::Fills(
            values
                .into_iter()
                .map(|value| match value {
                    style_engine::Value::Fill(fill) => Ok(fill),
                    _ => Err(invalid()),
                })
                .collect::<series_engine::SymbolResult<_>>()?,
        ),
        PropertyKind::SymbolFills => Property::SymbolFills(
            values
                .into_iter()
                .map(|value| match value {
                    style_engine::Value::SymbolFill(fill) => Ok(fill),
                    _ => Err(invalid()),
                })
                .collect::<series_engine::SymbolResult<_>>()?,
        ),
        _ => return Err(invalid()),
    })
}

/// Exact-source reversible body-chart per-series patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    position: Position,
    target: Target,
    before: Property,
    after: Property,
    changed_components: Arc<[Arc<str>]>,
    plan: MutationPlan,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .field("series", &self.after.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether this patch is an exact byte and semantic no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop() && self.before == self.after
    }

    /// Return the exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            position: self.position,
            target: self.target.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
            changed_components: self.changed_components.clone(),
            plan: self.plan.inverse(),
        }
    }
}

/// Compact publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
}

impl Diagnostics {
    /// Whether package bytes or semantic values changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten existing components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
}

/// Fully reopened immutable result of one body-chart per-series edit.
#[must_use = "a chart-series commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

fn apply_patch(package: &Package, patch: &Patch) -> Result<Commit, Error> {
    let mut ledger = Ledger::new(package)?;
    if !authorizes(package, &patch.artifacts) {
        return Err(Error::PatchConflict);
    }
    let current = ledger.target(package, BodyChartSelector::position(patch.position))?;
    let before = read_property(&mut ledger, package, &current, patch.before.kind())?;
    if !current.chart.same_identity(&patch.target.chart)
        || current.kind != patch.target.kind
        || current.series != patch.target.series
        || before != patch.before
    {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Commit {
            package: package.clone(),
            patch: patch.clone(),
            diagnostics: Diagnostics::unchanged(),
        });
    }
    let candidate = ledger.reopen(package, &patch.artifacts, &patch.plan)?;
    let verified = ledger.target(&candidate, BodyChartSelector::position(patch.position))?;
    let after = read_property(&mut ledger, &candidate, &verified, patch.after.kind())?;
    if !verified.chart.same_identity(&patch.target.chart)
        || verified.kind != patch.target.kind
        || verified.series != patch.target.series
        || after != patch.after
    {
        return Err(Error::Verification);
    }
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics::published(patch.changed_components.len()),
    })
}

impl Package {
    /// Read one per-series property of the selected body chart, one value
    /// per visible series in series order.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedKind`], source and limit
    /// errors.
    pub fn body_chart_series(
        &self,
        selector: impl Into<BodyChartSelector>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let mut ledger = Ledger::new(self)?;
        let target = ledger.target(self, selector.into())?;
        read_property(&mut ledger, self, &target, kind)
    }

    /// Begin a selector-first per-series chart transaction.
    ///
    /// ```no_run
    /// use litchi_pages::Package;
    /// use litchi_pages::chart::series::PropertyKind;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.pages")?;
    /// let commit = package
    ///     .edit_body_chart_series(0)?
    ///     .reset(PropertyKind::Fills, 0)?
    ///     .commit()?;
    /// let restored = commit.package().apply_body_chart_series(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector errors.
    pub fn edit_body_chart_series(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        let mut ledger = Ledger::new(self)?;
        ledger.target(self, selector)?;
        Ok(Edit {
            source: self,
            selector,
            change: None,
        })
    }

    /// Apply a patch produced from this exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; locality and verification errors.
    pub fn apply_body_chart_series(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const NATIVE: &[u8] =
        include_bytes!("../../../../test-data/iwork/pages/chart-data-native.pages");

    fn append_unknown_marker(data: &mut Vec<u8>, marker: u64) {
        litchi_iwa_common::wire::append_varint_field(data, 12_345, marker)
            .expect("bounded unknown marker");
    }

    fn empty_plan() -> MutationPlan {
        MutationPlan::default()
    }

    #[test]
    fn production_locality_rejects_unplanned_object_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let component = package
            .state
            .source
            .components()
            .iter()
            .find(|component| component.archive().objects.len() >= 2)
            .expect("fixture component with two objects");
        let component_name = Arc::<str>::from(component.name());
        let ids = component
            .archive()
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .take(2)
            .collect::<Vec<_>>();
        let allowed = ids[0];
        let unplanned = ids[1];
        let mut archive = component.archive().clone();
        let message = archive
            .object_mut(unplanned)
            .expect("unplanned object")
            .messages
            .first_mut()
            .expect("unplanned message");
        append_unknown_marker(&mut message.data, 0xfeed);
        let candidate =
            package_graph::publish_archives(&package, vec![(Arc::clone(&component_name), archive)])
                .expect("candidate publication");
        let plan = MutationPlan {
            components: vec![Arc::clone(&component_name)],
            objects: vec![(component_name, vec![allowed])],
            ..MutationPlan::default()
        };
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(&package, &candidate, &plan, &mut budget,),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_changed_chart_unknown_field() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let (component_name, chart_id) = package
            .state
            .source
            .components()
            .iter()
            .find_map(|component| {
                component
                    .archive()
                    .objects
                    .iter()
                    .find(|object| {
                        object
                            .messages
                            .iter()
                            .any(|message| message.type_ == CHART_MESSAGE_TYPE)
                    })
                    .and_then(|object| object.archive_info.identifier)
                    .map(|identifier| (Arc::<str>::from(component.name()), identifier))
            })
            .expect("fixture chart object");
        let mut archive = package
            .state
            .source
            .components()
            .get(component_name.as_ref())
            .expect("chart component")
            .archive()
            .clone();
        let message = archive
            .object_mut(chart_id)
            .expect("chart object")
            .messages
            .iter_mut()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .expect("chart message");
        append_unknown_marker(&mut message.data, 0xbeef);
        let candidate =
            package_graph::publish_archives(&package, vec![(Arc::clone(&component_name), archive)])
                .expect("candidate publication");
        let plan = MutationPlan {
            components: vec![Arc::clone(&component_name)],
            objects: vec![(component_name, vec![chart_id])],
            ..MutationPlan::default()
        };
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(&package, &candidate, &plan, &mut budget,),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_extra_zip_member() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let candidate = package_graph::publish_changes(
            &package,
            Vec::new(),
            &[("Data/unplanned-series.bin", b"unused".as_slice())],
            &[],
            &[],
        )
        .expect("extra-member candidate publication");
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(&package, &candidate, &empty_plan(), &mut budget,),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_unplanned_member_removal() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let removable = package
            .state
            .source
            .package()
            .iter()
            .find(|entry| entry.name() != "Index/Document.iwa" && !entry.is_opaque())
            .expect("removable decoded member");
        let removed =
            package_graph::publish_changes(&package, Vec::new(), &[], &[], &[removable.name()])
                .expect("unplanned IWA removal");
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(&package, &removed, &empty_plan(), &mut budget,),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_zip_metadata_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let entry = package
            .state
            .source
            .package()
            .iter()
            .find(|entry| !entry.is_opaque())
            .expect("decoded package member");
        let central = entry.raw_record().central_directory_record();
        let mut bytes = package.state.source.shared_source().to_vec();
        let offset = bytes
            .windows(central.len())
            .position(|window| window == central)
            .expect("central record in source bytes");
        bytes[offset + 38] ^= 1;
        let candidate = Package::from_bytes(&bytes).expect("metadata-tampered candidate");
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(&package, &candidate, &empty_plan(), &mut budget,),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_archive_metadata_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let component = package
            .state
            .source
            .components()
            .iter()
            .find(|component| !component.archive().objects.is_empty())
            .expect("fixture component");
        let component_name = Arc::<str>::from(component.name());
        let object_id = component.archive().objects[0]
            .archive_info
            .identifier
            .expect("object identifier");
        let mut archive = component.archive().clone();
        let object = archive.object_mut(object_id).expect("metadata object");
        object.archive_info.should_merge = Some(!object.archive_info.should_merge.unwrap_or(false));
        let candidate =
            package_graph::publish_archives(&package, vec![(Arc::clone(&component_name), archive)])
                .expect("metadata candidate publication");
        let plan = MutationPlan {
            components: vec![Arc::clone(&component_name)],
            objects: vec![(component_name, vec![object_id])],
            ..MutationPlan::default()
        };
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(&package, &candidate, &plan, &mut budget,),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_locality_rejects_unplanned_registry_object_mutation() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let component = package
            .state
            .source
            .components()
            .iter()
            .find(|component| component.name() == "Index/Metadata.iwa")
            .expect("package metadata component");
        let component_name = Arc::<str>::from(component.name());
        let object_id = component.archive().objects[0]
            .archive_info
            .identifier
            .expect("registry object identifier");
        let mut archive = component.archive().clone();
        let message = archive
            .object_mut(object_id)
            .expect("registry object")
            .messages
            .first_mut()
            .expect("registry message");
        append_unknown_marker(&mut message.data, 0xabad1dea);
        let candidate =
            package_graph::publish_archives(&package, vec![(Arc::clone(&component_name), archive)])
                .expect("registry candidate publication");
        let mut budget = ArrangementBudget::new(&package).expect("locality budget");
        assert!(matches!(
            verify_locality(&package, &candidate, &empty_plan(), &mut budget,),
            Err(Error::Verification)
        ));
    }

    #[test]
    fn production_limit_mapping_preserves_scoped_axes() {
        assert_eq!(
            map_arrangement_limit(arrangement::LimitKind::PayloadMessages),
            LimitKind::PayloadMessages
        );
        assert_eq!(
            map_arrangement_limit(arrangement::LimitKind::PayloadObjects),
            LimitKind::PayloadObjects
        );
        assert_eq!(
            map_arrangement_limit(arrangement::LimitKind::Entries),
            LimitKind::Entries
        );
        assert_eq!(
            map_arrangement_limit(arrangement::LimitKind::WireOutputBytes),
            LimitKind::Wire(litchi_iwa_common::LimitKind::OutputBytes)
        );
        assert!(matches!(
            map_symbol_error(series_engine::SymbolError::IdentifierOverflow),
            Error::Refused(Refusal::UnsupportedPublication)
        ));
        assert!(matches!(
            map_package_error(PackageError::Limit {
                kind: crate::LimitKind::Semantic(crate::SemanticLimitKind::TextBytes),
                observed: 2,
                maximum: 1,
            }),
            Error::Limit {
                kind: LimitKind::PayloadBytes,
                observed: 2,
                maximum: 1,
            }
        ));
        assert!(matches!(
            map_archive_error(litchi_iwa_archive::Error::Iwa(
                litchi_iwa_core::Error::Limit {
                    kind: litchi_iwa_core::LimitKind::ObjectBytes,
                    observed: 8,
                    maximum: 4,
                },
            )),
            Error::Limit {
                kind: LimitKind::Archive(litchi_iwa_core::LimitKind::ObjectBytes),
                observed: 8,
                maximum: 4,
            }
        ));
        assert!(matches!(
            map_package_error(PackageError::SectionNamesTooLarge {
                observed: 12,
                limit: 10,
            }),
            Error::Limit {
                kind: LimitKind::PayloadBytes,
                observed: 12,
                maximum: 10,
            }
        ));
        assert!(matches!(
            map_package_error(PackageError::Semantic(crate::Error::TooManySections {
                actual: 7,
                limit: 6,
            })),
            Error::Limit {
                kind: LimitKind::PayloadItems,
                observed: 7,
                maximum: 6,
            }
        ));
        assert!(matches!(
            map_package_error(PackageError::Semantic(crate::Error::TooManyBodyStorages {
                actual: 9,
                limit: 8,
            })),
            Error::Limit {
                kind: LimitKind::PayloadItems,
                observed: 9,
                maximum: 8,
            }
        ));
        assert!(matches!(
            map_package_error(PackageError::Semantic(crate::Error::TextTooLarge {
                observed: 11,
                limit: 10,
            })),
            Error::Limit {
                kind: LimitKind::PayloadBytes,
                observed: 11,
                maximum: 10,
            }
        ));
    }
}
