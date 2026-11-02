//! Selector-first Keynote slide-chart per-series transactions (D7-A, record
//! 0414) and the shared core of the Keynote chart-property families.
//!
//! The shared W1-E engines (`litchi_iwa_style::{series, chart_style}`) own
//! the graph work: series overrides and their allocation, the effective
//! series styles, the chart- and axis-level style objects. This leaf adapter
//! supplies only Keynote chart selection, semantic values, the transaction
//! ledger, exact-source publication under the engine's source-derived
//! mutation scope, candidate reopen/readback, and typed refusal mapping. The
//! `chart::style` and `chart::axis` families
//! (`super::slide_chart_style`) publish through the same core.

#![allow(
    clippy::cast_sign_loss,
    clippy::needless_pass_by_value,
    reason = "The focused adapter keeps the semantic boundary beside its staged transaction."
)]

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, EntryInsertion, ExactArtifacts};
use litchi_iwa_core::{Archive, Limits as ArchiveLimits};
use litchi_iwa_protos::chart_metadata_codec::{self, DecodeOptions};
use litchi_iwa_protos::chart_style_codec;
use litchi_iwa_style::graph::{self, Components, Workspace};
use litchi_iwa_style::series as engine;
use thiserror::Error as ThisError;

use super::chart_axis_support::{self, AxisSupportError, AxisSupportSelectorError};
use super::slide_chart_font::{self, FieldLedger};
use super::slide_table_core::{Budget as TransactionBudget, Error as TransactionError};
use super::{Package, PayloadLimitKind, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{ChartSelector, SlideSelector};
use litchi_core::Position;
use litchi_iwa_common::chart::error_bar::Series as ErrorBarSeries;
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::number_format::{LabelAffixes, NumberFormat};
use litchi_iwa_common::chart::pie::{
    LabelDistance, LabelVisibility, LeaderLineVisibility, WedgeExplosion,
};
use litchi_iwa_common::chart::series::{
    ConnectionLine, ConnectionLineFamily, Stroke, Symbol, SymbolFamily, SymbolFill, Trendline,
};
use litchi_iwa_common::chart::series_labels::Visibility;
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_style::series::property as series_property;
use litchi_iwa_style::series::style as style_engine;
pub use litchi_iwa_style::series::style::{AutoFit, ValueLabelLocation};

/// One per-series chart property: a value for every series, in series
/// order.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// The line joining a series' points (line, area, radar and scatter
    /// charts).
    ConnectionLines(Vec<ConnectionLine>),
    /// How far each wedge is pulled out of a pie or donut.
    WedgeExplosions(Vec<WedgeExplosion>),
    /// Whether each pie wedge's label draws a leader line.
    LeaderLines(Vec<LeaderLineVisibility>),
    /// How far each wedge's label sits from the pie centre.
    LabelDistances(Vec<LabelDistance>),
    /// A symbol per visible series; `None` hides that series' marker.
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
    /// Where each bar or column series draws its value labels.
    ValueLabelLocations(Vec<ValueLabelLocation>),
    /// Whether each series' error bars are spaced automatically.
    ErrorBarAutoFits(Vec<AutoFit>),
    /// Whether each series' value labels move to fit.
    ValueLabelAutoFits(Vec<AutoFit>),
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
    /// [`Property::LeaderLines`].
    LeaderLines,
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
    /// [`Property::ValueLabelLocations`].
    ValueLabelLocations,
    /// [`Property::ErrorBarAutoFits`].
    ErrorBarAutoFits,
    /// [`Property::ValueLabelAutoFits`].
    ValueLabelAutoFits,
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
            Self::ValueLabelLocations => Some(style_engine::Kind::ValueLabelLocations),
            Self::ErrorBarAutoFits => Some(style_engine::Kind::ErrorBarAutoFits),
            Self::ValueLabelAutoFits => Some(style_engine::Kind::ValueLabelAutoFits),
            Self::SymbolFills => Some(style_engine::Kind::SymbolFills),
            Self::SymbolOutlines => Some(style_engine::Kind::SymbolOutlines),
            _ => None,
        }
    }
}

impl Property {
    /// The property this value sets.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        match self {
            Self::ConnectionLines(_) => PropertyKind::ConnectionLines,
            Self::WedgeExplosions(_) => PropertyKind::WedgeExplosions,
            Self::LeaderLines(_) => PropertyKind::LeaderLines,
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
            Self::ValueLabelLocations(_) => PropertyKind::ValueLabelLocations,
            Self::ErrorBarAutoFits(_) => PropertyKind::ErrorBarAutoFits,
            Self::ValueLabelAutoFits(_) => PropertyKind::ValueLabelAutoFits,
            Self::SymbolFills(_) => PropertyKind::SymbolFills,
            Self::SymbolOutlines(_) => PropertyKind::SymbolOutlines,
        }
    }

    /// Return the number of series named by this property.
    #[must_use]
    pub fn len(&self) -> usize {
        match self {
            Self::ConnectionLines(values) => values.len(),
            Self::WedgeExplosions(values) => values.len(),
            Self::LeaderLines(values) => values.len(),
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
            Self::ValueLabelLocations(values) => values.len(),
            Self::ErrorBarAutoFits(values) => values.len(),
            Self::ValueLabelAutoFits(values) => values.len(),
            Self::SymbolFills(values) => values.len(),
        }
    }

    /// Whether this property names no series.
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
            Self::ValueLabelLocations(values) => values
                .iter()
                .map(|value| Some(style_engine::Value::ValueLabelLocation(*value)))
                .collect(),
            Self::ErrorBarAutoFits(values) => values
                .iter()
                .map(|value| Some(style_engine::Value::AutoFit(*value)))
                .collect(),
            Self::ValueLabelAutoFits(values) => values
                .iter()
                .map(|value| Some(style_engine::Value::AutoFit(*value)))
                .collect(),
            Self::SymbolFills(values) => values
                .iter()
                .map(|value| Some(style_engine::Value::SymbolFill(value.clone())))
                .collect(),
            _ => return None,
        })
    }
}

/// Finite resource dimensions exposed by the focused route.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Wire input bytes.
    InputBytes,
    /// Wire output bytes.
    OutputBytes,
    /// Bytes in one protobuf payload or archive member.
    WireBytes,
    /// Package entries, objects and messages.
    Entries,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Bytes in one raw ZIP member name.
    MemberNameBytes,
    /// Aggregate raw ZIP header variable metadata bytes.
    MetadataBytes,
    /// Compressed bytes in one ZIP member.
    CompressedEntryBytes,
    /// Decompressed bytes in one IWA component.
    IwaStreamBytes,
    /// Aggregate decompressed bytes across IWA components.
    IwaTotalBytes,
    /// Aggregate IWA objects inspected by detection.
    IwaObjects,
    /// Protobuf fields inspected by detection.
    IwaFields,
    /// Protobuf nesting inspected by detection.
    IwaNesting,
    /// Aggregate detection wire work.
    IwaWork,
    /// Semantic slides visited by the shared chart selector.
    Slides,
    /// Graph references.
    References,
    /// Font-name or other semantic text bytes retained by the selector.
    TextBytes,
    /// Native chart title bytes retained by the selector.
    TitleBytes,
    /// Series-style registry entries rewritten.
    RegistryEntries,
    /// Package-wide object identifiers allocated or released.
    Identifiers,
    /// Snappy frames inspected by the IWA boundary.
    SnappyFrames,
    /// Parsed wire fields.
    WireFields,
    /// Maximum protobuf nesting depth.
    WireNesting,
    /// Aggregate graph work.
    WireWork,
    /// Graph allocation events.
    Allocations,
    /// Retained graph bytes.
    RetainedBytes,
    /// Temporary graph scratch bytes.
    ScratchBytes,
    /// Caller bytes retained by deferred selectors.
    SelectorBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::MemberNameBytes => "member name bytes",
            Self::MetadataBytes => "metadata bytes",
            Self::CompressedEntryBytes => "compressed entry bytes",
            Self::IwaStreamBytes => "IWA stream bytes",
            Self::IwaTotalBytes => "total IWA bytes",
            Self::IwaObjects => "IWA objects",
            Self::IwaFields => "IWA fields",
            Self::IwaNesting => "IWA nesting",
            Self::IwaWork => "IWA work",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextBytes => "text bytes",
            Self::TitleBytes => "title bytes",
            Self::RegistryEntries => "registry entries",
            Self::Identifiers => "identifiers",
            Self::SnappyFrames => "Snappy frames",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::SelectorBytes => "selector bytes",
        })
    }
}

/// A well-formed request this route refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The source is not an exact physical package.
    #[error("this Keynote source does not support chart-series edits")]
    UnsupportedSource,
    /// The selector did not resolve one slide/chart.
    #[error("the selected Keynote chart was not found")]
    ChartNotFound,
    /// An exact-name selector was ambiguous.
    #[error("the Keynote chart-series selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked semantic chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound { position: Position },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// The property does not apply to the chart's kind (data symbols of a
    /// bar chart, a pie start angle of a column chart).
    #[error("the property does not apply to the selected chart kind")]
    UnsupportedKind,
    /// The value does not name the chart's visible series count.
    #[error("chart series value count does not match the chart")]
    SeriesCount,
    /// The series is past the chart's series.
    #[error("the selected chart has no series {series}")]
    SeriesNotFound {
        /// The requested zero-based series.
        series: usize,
    },
    /// The property has no reset (only the series-style properties do).
    #[error("the chart property cannot be reset")]
    NotResettable,
    /// The style the write would change is shared with another chart.
    #[error("the chart style is shared with another chart")]
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
    /// Commit was requested without a staged value.
    #[error("the Keynote chart-series transaction has no staged operation")]
    NoStagedOperation,
    /// The graph is valid but outside the focused ownership profile.
    #[error("the selected chart-series graph is outside the focused route")]
    UnsupportedDependency,
}

impl Refusal {
    /// Map this semantic refusal to the unified refusal taxonomy.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource | Self::UnsupportedDependency => {
                RefusalKind::UnsupportedContent
            },
            Self::ChartNotFound
            | Self::SlideNameNotFound
            | Self::SlidePositionNotFound { .. }
            | Self::ChartNameNotFound
            | Self::ChartPositionNotFound { .. } => RefusalKind::NotFound,
            Self::AmbiguousSelector | Self::EmptySlideName | Self::EmptyChartName => {
                RefusalKind::InvalidValue
            },
            Self::UnsupportedKind | Self::SeriesCount => RefusalKind::InvalidValue,
            Self::SeriesNotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedAxis | Self::MissingStyle => RefusalKind::UnsupportedTarget,
            Self::SharedStyle | Self::ImageFill => RefusalKind::UnsupportedContent,
            Self::NotResettable | Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
        }
    }
}

/// A content-redacted chart-series failure.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted graph is malformed.
    #[error("the Keynote chart-series source is invalid")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote chart-series {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource dimension.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} Keynote chart-series units")]
    Allocation { amount: usize },
    /// A finite finalization calculation overflowed before allocation.
    #[error("Keynote chart-series finalization arithmetic overflow")]
    Overflow,
    /// The reopened candidate failed semantic verification.
    #[error("the Keynote chart-series candidate failed verification")]
    Verification,
    /// The patch was created from another exact source.
    #[error("the Keynote chart-series patch does not match the exact source")]
    PatchConflict,
}

/// Compact evidence describing one chart-series commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    pub(super) const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    pub(super) const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }

    /// Whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Whether the candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Selection {
    pub(super) slide: Position,
    pub(super) chart: Position,
    pub(super) identifier: u64,
}

fn select(
    package: &Package,
    slide: SlideSelector<'_>,
    chart: ChartSelector<'_>,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
) -> Result<Selection, Error> {
    let selected = slide_chart_font::select_chart_for_series(package, slide, chart, budget, fields)
        .map_err(map_axis_support_error)?;
    Ok(Selection {
        slide: selected.slide_position,
        chart: selected.chart_position,
        identifier: selected.chart_identifier,
    })
}

fn map_axis_support_error(error: AxisSupportError) -> Error {
    match error {
        AxisSupportError::Selector(selector) => match selector {
            AxisSupportSelectorError::UnsupportedSource => Refusal::UnsupportedSource.into(),
            AxisSupportSelectorError::AmbiguousSelector => Refusal::AmbiguousSelector.into(),
            AxisSupportSelectorError::EmptySlideName => Refusal::EmptySlideName.into(),
            AxisSupportSelectorError::SlideNameNotFound => Refusal::SlideNameNotFound.into(),
            AxisSupportSelectorError::SlidePositionNotFound { position } => {
                Refusal::SlidePositionNotFound { position }.into()
            },
            AxisSupportSelectorError::ChartNameNotFound => Refusal::ChartNameNotFound.into(),
            AxisSupportSelectorError::ChartPositionNotFound { position } => {
                Refusal::ChartPositionNotFound { position }.into()
            },
            AxisSupportSelectorError::EmptyChartName => Refusal::EmptyChartName.into(),
        },
        AxisSupportError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                chart_axis_support::AxisSupportLimitKind::InputBytes => LimitKind::InputBytes,
                chart_axis_support::AxisSupportLimitKind::OutputBytes => LimitKind::OutputBytes,
                chart_axis_support::AxisSupportLimitKind::WireBytes => LimitKind::WireBytes,
                chart_axis_support::AxisSupportLimitKind::Entries => LimitKind::Entries,
                chart_axis_support::AxisSupportLimitKind::EntryBytes => LimitKind::EntryBytes,
                chart_axis_support::AxisSupportLimitKind::TotalBytes => LimitKind::TotalBytes,
                chart_axis_support::AxisSupportLimitKind::Slides => LimitKind::Slides,
                chart_axis_support::AxisSupportLimitKind::TextStorages
                | chart_axis_support::AxisSupportLimitKind::TextFragments => LimitKind::Entries,
                chart_axis_support::AxisSupportLimitKind::References => LimitKind::References,
                chart_axis_support::AxisSupportLimitKind::TextBytes => LimitKind::TextBytes,
                chart_axis_support::AxisSupportLimitKind::WireFields => LimitKind::WireFields,
                chart_axis_support::AxisSupportLimitKind::WireNesting => LimitKind::WireNesting,
                chart_axis_support::AxisSupportLimitKind::WireWork => LimitKind::WireWork,
                chart_axis_support::AxisSupportLimitKind::TitleBytes => LimitKind::TitleBytes,
            },
            observed,
            maximum,
        },
        AxisSupportError::Allocation { amount } => Error::Allocation { amount },
        AxisSupportError::InvalidSource => Error::InvalidSource,
    }
}

fn kind_and_count(
    package: &Package,
    selection: Selection,
    budget: &mut TransactionBudget,
) -> Result<(ChartKind, usize), Error> {
    let payload = package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| {
            component
                .archive()
                .object(selection.identifier)
                .and_then(|object| {
                    object
                        .messages
                        .iter()
                        .find(|message| message.type_ == 5_021)
                        .map(|message| message.data.as_slice())
                })
        })
        .ok_or(Error::InvalidSource)?;
    let limits = budget.residual(package).map_err(map_transaction_error)?;
    budget
        .preflight_allocations(1)
        .and_then(|_| budget.preflight_retained(payload.len()))
        .and_then(|_| budget.preflight_scratch(payload.len()))
        .map_err(map_transaction_error)?;
    let text_maximum = package.semantic_limits().max_text_bytes();
    if payload.len() > text_maximum {
        return Err(Error::Limit {
            kind: LimitKind::WireBytes,
            observed: payload.len() as u64,
            maximum: text_maximum as u64,
        });
    }
    let maximum_depth = u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?;
    let options = DecodeOptions::new(
        payload.len().max(1).min(limits.max_input_bytes()),
        limits.max_fields().max(1),
        limits.max_rewrite_work().max(1),
        maximum_depth,
        limits.max_fields().clamp(1, 1_000_000),
        limits
            .max_input_bytes()
            .min(text_maximum)
            .clamp(1, 64 * 1024 * 1024),
    );
    let (snapshot, report) =
        match chart_metadata_codec::decode_modern_with_report(payload, &options) {
            Ok(decoded) => decoded,
            Err(error) => {
                charge_metadata_report(budget, error.report())?;
                return Err(map_chart_metadata_decode_error(error));
            },
        };
    charge_metadata_report(budget, report)?;
    let kind = ChartKind::from_native(snapshot.chart_type());
    let series =
        engine::visible_series(kind, snapshot.series_count()).ok_or(Error::InvalidSource)?;
    Ok((kind, series))
}

fn charge_metadata_report(
    budget: &mut TransactionBudget,
    report: chart_metadata_codec::DecodeReport,
) -> Result<(), Error> {
    budget
        .input(report.source_bytes())
        .and_then(|_| budget.fields(report.fields()))
        .and_then(|_| budget.work(report.work_bytes()))
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.nesting(report.max_depth() as usize))
        .map_err(map_transaction_error)
}

fn map_chart_metadata_decode_error(error: chart_metadata_codec::DecodeError) -> Error {
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.input_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireBytes,
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
    if let Some((observed, maximum)) = error.depth_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireNesting,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

struct LeafComponents<'a> {
    components: Vec<(&'a str, &'a Archive)>,
}

impl<'a> LeafComponents<'a> {
    fn view_requirements(package: &Package) -> Result<(usize, usize), Error> {
        let count = package.state.source.components().len();
        let bytes = count
            .checked_mul(size_of::<(&str, &Archive)>())
            .and_then(|value| value.checked_add(size_of::<Self>()))
            .ok_or(Error::InvalidSource)?;
        Ok((usize::from(count != 0).saturating_add(1), bytes))
    }

    fn new(package: &'a Package) -> Result<Self, Error> {
        let count = package.state.source.components().len();
        let mut components = Vec::new();
        components
            .try_reserve_exact(count)
            .map_err(|_| Error::Allocation { amount: count })?;
        for index in 0..count {
            let component = package
                .state
                .source
                .components()
                .get_index(index)
                .ok_or(Error::InvalidSource)?;
            components.push((component.name(), component.archive()));
        }
        Ok(Self { components })
    }
}

impl Components for LeafComponents<'_> {
    fn count(&self) -> usize {
        self.components.len()
    }

    fn component(&self, index: usize) -> Option<(&str, &Archive)> {
        self.components.get(index).copied()
    }
}

fn graph_limits(package: &Package) -> Result<ArchiveLimits, Error> {
    package
        .structure_archive_limits()
        .map_err(|_| Error::InvalidSource)
}

fn transaction_budget(package: &Package) -> Result<TransactionBudget, Error> {
    let mut budget = TransactionBudget::new(package).map_err(map_transaction_error)?;
    budget
        .preflight_semantic_scan(package)
        .map_err(map_transaction_error)?;
    budget
        .inventory_components(package)
        .map_err(map_transaction_error)?;
    Ok(budget)
}

fn transaction_fields(package: &Package) -> Result<FieldLedger, Error> {
    slide_chart_font::field_ledger_for_series(package).map_err(map_axis_support_error)
}

struct WorkspaceAdmission<'a> {
    components: Box<LeafComponents<'a>>,
    requirements: graph::ConstructionRequirements,
    limits: graph::ConstructionLimits,
}

fn map_construction(error: graph::ConstructionError) -> Error {
    match error {
        graph::ConstructionError::Graph(error) => map_graph(error),
        graph::ConstructionError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                graph::ConstructionLimitKind::Work => LimitKind::WireWork,
                graph::ConstructionLimitKind::Allocations => LimitKind::Allocations,
                graph::ConstructionLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                graph::ConstructionLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => return Error::InvalidSource,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        graph::ConstructionError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn admit_workspace<'a>(
    package: &'a Package,
    budget: &mut TransactionBudget,
) -> Result<WorkspaceAdmission<'a>, Error> {
    let components = package.state.source.components();
    let component_count = components.len();
    let object_count = package.state.total_objects;
    let (view_allocations, view_retained) = LeafComponents::view_requirements(package)?;
    budget
        .allocations(view_allocations)
        .and_then(|_| budget.retained(view_retained))
        .map_err(map_transaction_error)?;
    let components_view = Box::new(LeafComponents::new(package)?);
    let construction_limits = budget
        .construction_limits()
        .map_err(map_transaction_error)?;
    let requirements = Workspace::construction_requirements(&*components_view, construction_limits)
        .map_err(map_construction)?;
    budget
        .components(component_count)
        .and_then(|_| budget.payload_objects(object_count))
        .and_then(|_| budget.charge_construction(requirements))
        .map_err(map_transaction_error)
        .map(|()| WorkspaceAdmission {
            components: components_view,
            requirements,
            limits: construction_limits,
        })
}

fn map_transaction_error(error: TransactionError) -> Error {
    match error {
        TransactionError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                super::slide_table_core::LimitKind::InputBytes => LimitKind::InputBytes,
                super::slide_table_core::LimitKind::OutputBytes => LimitKind::OutputBytes,
                super::slide_table_core::LimitKind::WireFields => LimitKind::WireFields,
                super::slide_table_core::LimitKind::WireWork => LimitKind::WireWork,
                super::slide_table_core::LimitKind::Allocations => LimitKind::Allocations,
                super::slide_table_core::LimitKind::Retained => LimitKind::RetainedBytes,
                super::slide_table_core::LimitKind::Scratch => LimitKind::ScratchBytes,
                super::slide_table_core::LimitKind::WireNesting => LimitKind::WireNesting,
                super::slide_table_core::LimitKind::References => LimitKind::References,
                super::slide_table_core::LimitKind::Entries
                | super::slide_table_core::LimitKind::Components
                | super::slide_table_core::LimitKind::PayloadObjects
                | super::slide_table_core::LimitKind::PayloadMessages => LimitKind::Entries,
                super::slide_table_core::LimitKind::EntryBytes => LimitKind::EntryBytes,
                super::slide_table_core::LimitKind::TotalBytes => LimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        TransactionError::Allocation(amount) => Error::Allocation { amount },
        TransactionError::UnsupportedSource => Refusal::UnsupportedSource.into(),
        TransactionError::UnsupportedDependency => Refusal::UnsupportedDependency.into(),
        TransactionError::InvalidSource
        | TransactionError::UnsupportedTopology
        | TransactionError::AmbiguousSelector
        | TransactionError::EmptySlideName
        | TransactionError::SlideNameNotFound
        | TransactionError::SlidePositionNotFound(_)
        | TransactionError::TablePositionNotFound(_)
        | TransactionError::Read
        | TransactionError::Wire
        | TransactionError::Codec
        | TransactionError::Archive
        | TransactionError::Verification => Error::InvalidSource,
    }
}

fn map_selector_error(error: super::edit::SelectorError) -> Error {
    match error {
        super::edit::SelectorError::Limit { observed, maximum } => Error::Limit {
            kind: LimitKind::SelectorBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        super::edit::SelectorError::Allocation { amount } => Error::Allocation { amount },
    }
}

fn map_graph(error: graph::Error) -> Error {
    match error {
        graph::Error::Limit {
            kind: graph::GraphLimitKind::ScratchBytes,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::ScratchBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        graph::Error::Limit {
            kind: graph::GraphLimitKind::AllocationEvents,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::Allocations,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        graph::Error::Archive(error) => match error {
            litchi_iwa_core::Error::Allocation { requested, .. } => {
                Error::Allocation { amount: requested }
            },
            litchi_iwa_core::Error::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: map_core_limit(kind),
                observed: observed as u64,
                maximum: maximum as u64,
            },
            _ => Error::InvalidSource,
        },
        graph::Error::Registry(error) => map_registry_error(error),
        graph::Error::Wire(error) => map_wire_error(error),
        graph::Error::Codec(error) => map_codec_error(error),
        graph::Error::MissingObject(_)
        | graph::Error::AmbiguousObject(_)
        | graph::Error::IdentifierOverflow => Error::InvalidSource,
        graph::Error::Invalid("a series style is shared by several charts") => {
            Refusal::UnsupportedDependency.into()
        },
        graph::Error::Invalid(_) => Error::InvalidSource,
        _ => Error::InvalidSource,
    }
}

fn map_finish_error(error: graph::FinishError) -> Error {
    match error {
        graph::FinishError::Graph(error) => map_graph(error),
        graph::FinishError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                graph::FinishLimitKind::Work => LimitKind::WireWork,
                graph::FinishLimitKind::Allocations => LimitKind::Allocations,
                graph::FinishLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                graph::FinishLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => return Error::InvalidSource,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        graph::FinishError::Allocation { amount } => Error::Allocation { amount },
        graph::FinishError::Overflow(_) => Error::Overflow,
        _ => Error::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::Entries,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_codec_error(error: chart_style_codec::Error) -> Error {
    match error {
        chart_style_codec::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                chart_style_codec::LimitKind::Bytes => LimitKind::WireBytes,
                chart_style_codec::LimitKind::Fields => LimitKind::WireFields,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        chart_style_codec::Error::Allocation { bytes } => Error::Allocation { amount: bytes },
        _ => Error::InvalidSource,
    }
}

fn map_registry_error(error: litchi_iwa_protos::package_metadata_codec::RewriteError) -> Error {
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::InputBytes {
                observed,
                maximum,
            } => (LimitKind::InputBytes, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::OutputBytes {
                observed,
                maximum,
            } => (LimitKind::OutputBytes, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Fields {
                observed,
                maximum,
            } => (LimitKind::WireFields, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (LimitKind::WireWork, observed, maximum)
            },
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Nesting {
                observed,
                maximum,
            } => (LimitKind::WireNesting, observed as usize, maximum as usize),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
                observed,
                maximum,
            }
            | litchi_iwa_protos::package_metadata_codec::RewriteLimit::Additions {
                observed,
                maximum,
            } => (LimitKind::Entries, observed, maximum),
            litchi_iwa_protos::package_metadata_codec::RewriteLimit::References {
                observed,
                maximum,
            } => (LimitKind::References, observed, maximum),
            _ => return Error::InvalidSource,
        };
        return Error::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

fn map_core_limit(kind: litchi_iwa_core::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_core::LimitKind::ArchiveBytes => LimitKind::TotalBytes,
        litchi_iwa_core::LimitKind::Objects
        | litchi_iwa_core::LimitKind::Messages
        | litchi_iwa_core::LimitKind::MessagesPerObject
        | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
        litchi_iwa_core::LimitKind::ObjectBytes
        | litchi_iwa_core::LimitKind::MessageBytes
        | litchi_iwa_core::LimitKind::HeaderBytes
        | litchi_iwa_core::LimitKind::HeaderMemoryBytes
        | litchi_iwa_core::LimitKind::SnappyChunkBytes
        | litchi_iwa_core::LimitKind::SnappyStreamBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => LimitKind::WireBytes,
        litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
        litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
        litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::SnappyFrames,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes => LimitKind::MemberNameBytes,
                litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::MetadataBytes,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    LimitKind::CompressedEntryBytes
                },
                litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::IwaStreamBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::IwaTotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
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
            kind: map_core_limit(kind),
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_detection_error(error: litchi_iwa_detect::Error) -> Error {
    match error {
        litchi_iwa_detect::Error::IwaCore(error) => map_core_error(error),
        litchi_iwa_detect::Error::IwaCommon(error) => map_wire_error(error),
        litchi_iwa_detect::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_detect::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_detect::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_detect::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_detect::LimitKind::MemberNameBytes => LimitKind::MemberNameBytes,
                litchi_iwa_detect::LimitKind::MetadataBytes => LimitKind::MetadataBytes,
                litchi_iwa_detect::LimitKind::CompressedEntryBytes => {
                    LimitKind::CompressedEntryBytes
                },
                litchi_iwa_detect::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_detect::LimitKind::TotalBytes => LimitKind::TotalBytes,
                litchi_iwa_detect::LimitKind::IwaStreamBytes => LimitKind::IwaStreamBytes,
                litchi_iwa_detect::LimitKind::IwaTotalBytes => LimitKind::IwaTotalBytes,
                litchi_iwa_detect::LimitKind::IwaObjects => LimitKind::IwaObjects,
                litchi_iwa_detect::LimitKind::IwaFields => LimitKind::IwaFields,
                litchi_iwa_detect::LimitKind::IwaNesting => LimitKind::IwaNesting,
                litchi_iwa_detect::LimitKind::IwaWork => LimitKind::IwaWork,
                _ => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_detect::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                SemanticLimitKind::Objects => LimitKind::Entries,
                SemanticLimitKind::Slides => LimitKind::Slides,
                SemanticLimitKind::References => LimitKind::References,
                SemanticLimitKind::TextStorages | SemanticLimitKind::TextFragments => {
                    LimitKind::Entries
                },
                SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                PayloadLimitKind::Bytes => LimitKind::WireBytes,
                PayloadLimitKind::Fields => LimitKind::WireFields,
                PayloadLimitKind::Nesting => LimitKind::WireNesting,
                PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        ReadError::Detection(error) => map_detection_error(error),
        ReadError::Io(_)
        | ReadError::UnexpectedFormat { .. }
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_) => Error::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

const fn is_series_rewritable_type(message_type: u32) -> bool {
    matches!(message_type, 401 | 2_022 | 5_021..=5_029 | 11_006 | 11_008)
}

fn map_font_error(error: slide_chart_font::Error) -> Error {
    match error {
        slide_chart_font::Error::Refused(slide_chart_font::Refusal::UnsupportedArchiveInfo) => {
            Refusal::UnsupportedDependency.into()
        },
        slide_chart_font::Error::Refused(_) => Refusal::UnsupportedSource.into(),
        slide_chart_font::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                slide_chart_font::LimitKind::InputBytes => LimitKind::InputBytes,
                slide_chart_font::LimitKind::OutputBytes => LimitKind::OutputBytes,
                slide_chart_font::LimitKind::WireBytes => LimitKind::WireBytes,
                slide_chart_font::LimitKind::Entries => LimitKind::Entries,
                slide_chart_font::LimitKind::EntryBytes => LimitKind::EntryBytes,
                slide_chart_font::LimitKind::TotalBytes => LimitKind::TotalBytes,
                slide_chart_font::LimitKind::Slides => LimitKind::Slides,
                slide_chart_font::LimitKind::Identifiers => LimitKind::Identifiers,
                slide_chart_font::LimitKind::References => LimitKind::References,
                slide_chart_font::LimitKind::WireFields => LimitKind::WireFields,
                slide_chart_font::LimitKind::WireNesting => LimitKind::WireNesting,
                slide_chart_font::LimitKind::WireWork => LimitKind::WireWork,
                slide_chart_font::LimitKind::Allocations => LimitKind::Allocations,
                slide_chart_font::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                slide_chart_font::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                slide_chart_font::LimitKind::SelectorBytes => LimitKind::SelectorBytes,
                slide_chart_font::LimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed,
            maximum,
        },
        slide_chart_font::Error::Allocation { amount } => Error::Allocation { amount },
        slide_chart_font::Error::InvalidSource => Error::InvalidSource,
        slide_chart_font::Error::Verification => Error::Verification,
        slide_chart_font::Error::PatchConflict => Error::PatchConflict,
    }
}

fn map_series_error(error: engine::SymbolError) -> Error {
    match error {
        engine::SymbolError::Graph(error) => map_graph(error),
        engine::SymbolError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                engine::LimitKind::InputBytes => LimitKind::InputBytes,
                engine::LimitKind::Fields => LimitKind::WireFields,
                engine::LimitKind::OutputBytes => LimitKind::OutputBytes,
                engine::LimitKind::Nesting => LimitKind::WireNesting,
                engine::LimitKind::Work => LimitKind::WireWork,
                engine::LimitKind::Allocations => LimitKind::Allocations,
                engine::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                engine::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                engine::LimitKind::References => LimitKind::References,
                engine::LimitKind::RegistryEntries => LimitKind::RegistryEntries,
                engine::LimitKind::Identifiers => LimitKind::Identifiers,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        engine::SymbolError::Allocation { amount } => Error::Allocation { amount },
        engine::SymbolError::Invalid("a series style is shared by several charts") => {
            Refusal::UnsupportedDependency.into()
        },
        engine::SymbolError::Invalid(_) => Error::InvalidSource,
        engine::SymbolError::IdentifierOverflow => Refusal::UnsupportedDependency.into(),
        engine::SymbolError::Refused(refusal) => match refusal {
            engine::Refusal::UnsupportedKind => Refusal::UnsupportedKind.into(),
            engine::Refusal::SharedStyle => Refusal::SharedStyle.into(),
            engine::Refusal::ImageFill => Refusal::ImageFill.into(),
            engine::Refusal::UnsupportedAxis => Refusal::UnsupportedAxis.into(),
            engine::Refusal::MissingStyle => Refusal::MissingStyle.into(),
            _ => Refusal::UnsupportedDependency.into(),
        },
        _ => Error::InvalidSource,
    }
}

fn series_limits(
    package: &Package,
    budget: &TransactionBudget,
    fields: &FieldLedger,
) -> Result<engine::Limits, Error> {
    const CHART_CODEC_MAX_BYTES: usize = 256 * 1024 * 1024;
    const CHART_CODEC_MAX_FIELDS: usize = litchi_iwa_common::WireLimits::MAX_FIELDS;
    let wire = budget.residual(package).map_err(map_transaction_error)?;
    let max_fields = fields
        .remaining()
        .map_err(map_font_error)?
        .min(wire.max_fields())
        .clamp(1, CHART_CODEC_MAX_FIELDS);
    let max_input_bytes = wire.max_input_bytes().clamp(1, CHART_CODEC_MAX_BYTES);
    let max_output_bytes = wire.max_output_bytes().clamp(1, CHART_CODEC_MAX_BYTES);
    let references = budget
        .remaining_references()
        .map_err(map_transaction_error)?;
    Ok(engine::Limits::new(
        max_input_bytes,
        max_fields,
        max_output_bytes,
        u32::try_from(wire.max_nesting()).map_err(|_| Error::InvalidSource)?,
        wire.max_rewrite_work(),
        budget
            .remaining_allocations()
            .map_err(map_transaction_error)?,
        budget.remaining_retained().map_err(map_transaction_error)?,
        budget.remaining_scratch().map_err(map_transaction_error)?,
        references,
        references,
        references,
    ))
}

fn charge_series_report(
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
    report: engine::Report,
) -> Result<(), Error> {
    let references = report
        .references()
        .checked_add(report.registry_entries())
        .and_then(|value| value.checked_add(report.identifiers()))
        .ok_or(Error::InvalidSource)?;
    budget
        .input(report.input_bytes())
        .and_then(|_| budget.output(report.output_bytes()))
        .and_then(|_| budget.work(report.work_bytes()))
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.scratch(report.scratch_bytes()))
        .and_then(|_| budget.references(references))
        .and_then(|_| budget.nesting(report.max_depth() as usize))
        .map_err(map_transaction_error)?;
    fields.charge(report.fields()).map_err(map_font_error)
}

fn finish_workspace(
    workspace: Workspace<'_>,
    budget: &mut TransactionBudget,
) -> Result<graph::Changes, Error> {
    let requirements = workspace.finish_requirements().map_err(map_finish_error)?;
    let limits = graph::FinishLimits::new(
        if requirements.work() == 0 {
            0
        } else {
            budget.remaining_work().map_err(map_transaction_error)?
        },
        if requirements.allocations() == 0 {
            0
        } else {
            budget
                .remaining_allocations()
                .map_err(map_transaction_error)?
        },
        if requirements.retained_bytes() == 0 {
            0
        } else {
            budget.remaining_retained().map_err(map_transaction_error)?
        },
        if requirements.scratch_bytes() == 0 {
            0
        } else {
            budget.remaining_scratch().map_err(map_transaction_error)?
        },
    );
    for (kind, observed, maximum) in [
        (LimitKind::WireWork, requirements.work(), limits.max_work()),
        (
            LimitKind::Allocations,
            requirements.allocations(),
            limits.max_allocations(),
        ),
        (
            LimitKind::RetainedBytes,
            requirements.retained_bytes(),
            limits.max_retained_bytes(),
        ),
        (
            LimitKind::ScratchBytes,
            requirements.scratch_bytes(),
            limits.max_scratch_bytes(),
        ),
    ] {
        if observed > maximum {
            return Err(Error::Limit {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
    }
    let (changes, report) = workspace.finish_bounded(limits).map_err(map_finish_error)?;
    // Component/index admission is charged by `admit_workspace`; this report
    // covers only the final vectors, copied names, and their traversal.
    budget
        .output(report.output_bytes())
        .and_then(|_| budget.work(report.work()))
        .and_then(|_| budget.allocations(report.allocations()))
        .and_then(|_| budget.retained(report.retained_bytes()))
        .and_then(|_| budget.scratch(report.scratch_bytes()))
        .map_err(map_transaction_error)?;
    Ok(changes)
}

fn map_archive_preparation_error(
    error: litchi_iwa_core::ArchiveEncodingPreparationError<Error>,
) -> Error {
    match error {
        litchi_iwa_core::ArchiveEncodingPreparationError::Admission(error) => error,
        litchi_iwa_core::ArchiveEncodingPreparationError::ArithmeticOverflow { .. } => {
            Error::Overflow
        },
        litchi_iwa_core::ArchiveEncodingPreparationError::Core(error) => {
            map_graph(graph::Error::Archive(error))
        },
        _ => Error::InvalidSource,
    }
}

fn map_snappy_preparation_error(
    error: litchi_iwa_core::SnappyCompressionPreparationError<Error>,
) -> Error {
    match error {
        litchi_iwa_core::SnappyCompressionPreparationError::Admission(error) => error,
        litchi_iwa_core::SnappyCompressionPreparationError::ArithmeticOverflow { .. } => {
            Error::Overflow
        },
        litchi_iwa_core::SnappyCompressionPreparationError::Core(error) => {
            map_graph(graph::Error::Archive(error))
        },
        _ => Error::InvalidSource,
    }
}

fn charge_archive_preparation_chunk(
    budget: &mut TransactionBudget,
    chunk: litchi_iwa_core::ArchiveEncodingPreparationChunk,
) -> Result<(), Error> {
    budget
        .work(chunk.traversal_bytes())
        .map_err(map_transaction_error)
}

fn charge_snappy_preparation_chunk(
    budget: &mut TransactionBudget,
    chunk: litchi_iwa_core::SnappyCompressionPreparationChunk,
) -> Result<(), Error> {
    budget
        .work(chunk.traversal_bytes())
        .map_err(map_transaction_error)
}

struct PreparedPublication<'a> {
    name: &'a str,
    archive: litchi_iwa_core::ArchiveEncodingPlan<'a>,
    snappy: litchi_iwa_core::SnappyCompressionPlan,
    position: Option<usize>,
}

fn prepare_publication<'a>(
    archive: &'a Archive,
    name: &'a str,
    position: Option<usize>,
    archive_limits: ArchiveLimits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    budget: &mut TransactionBudget,
) -> Result<PreparedPublication<'a>, Error> {
    let mut archive_visitor = |chunk| charge_archive_preparation_chunk(budget, chunk);
    let archive_plan = archive
        .prepare_encoding(archive_limits, &mut archive_visitor)
        .map_err(map_archive_preparation_error)?;
    let archive_requirements = archive_plan.requirements();
    budget
        .work(archive_requirements.work_bytes())
        .and_then(|_| budget.work(archive_requirements.identifier_sort_work()))
        .and_then(|_| budget.allocations(archive_requirements.allocation_events()))
        .and_then(|_| budget.retained(archive_requirements.retained_bytes()))
        .and_then(|_| budget.scratch(archive_requirements.scratch_bytes()))
        .map_err(map_transaction_error)?;

    let mut snappy_visitor = |chunk| charge_snappy_preparation_chunk(budget, chunk);
    let snappy_plan = litchi_iwa_core::SnappyStream::prepare_compression(
        archive_requirements.encoded_bytes(),
        snappy_limits,
        &mut snappy_visitor,
    )
    .map_err(map_snappy_preparation_error)?;
    let snappy_requirements = snappy_plan.requirements();
    let snappy_scratch = snappy_requirements
        .temporary_frame_capacity_bytes()
        .checked_add(snappy_requirements.encoder_workspace_bytes())
        .ok_or(Error::Overflow)?;
    budget
        .work(snappy_requirements.work_bytes())
        .and_then(|_| budget.allocations(snappy_requirements.allocation_events()))
        .and_then(|_| budget.retained(snappy_requirements.output_capacity_bytes()))
        .and_then(|_| budget.scratch(snappy_scratch))
        .map_err(map_transaction_error)?;
    Ok(PreparedPublication {
        name,
        archive: archive_plan,
        snappy: snappy_plan,
        position,
    })
}

fn execute_publication(prepared: PreparedPublication<'_>) -> Result<Vec<u8>, Error> {
    let encoded = prepared
        .archive
        .execute()
        .map_err(|error| map_graph(graph::Error::Archive(error)))?;
    prepared
        .snappy
        .execute(&encoded)
        .map_err(|error| map_graph(graph::Error::Archive(error)))
}

fn publish_changes(
    source: &Package,
    changes: graph::Changes,
    selected_chart: u64,
    scope: &engine::MutationScope,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
    chart_extension_fields: &[u32],
) -> Result<(Package, usize, slide_chart_font::MutationPlan), Error> {
    let touched_components = changes
        .edits
        .len()
        .checked_add(changes.insertions.len())
        .and_then(|count| count.checked_add(changes.deletions.len()))
        .ok_or(Error::InvalidSource)?;
    let extra_deletions = match &source.state.source {
        PhysicalSource::Package(catalog) => {
            super::rendering_invalidation::root_preview_deletions(catalog.package())
                .map_err(|_| Error::InvalidSource)?
        },
        PhysicalSource::Semantic(_) => return Err(Refusal::UnsupportedSource.into()),
    };
    let plan = slide_chart_font::mutation_plan_for_series_scope(
        source,
        selected_chart,
        &changes,
        scope,
        extra_deletions.names(),
        budget,
    )
    .map_err(map_font_error)?;
    let candidate = publish_series_with_prepared_policy(
        source,
        changes,
        budget,
        fields,
        &plan,
        extra_deletions.names(),
        is_series_rewritable_type,
        chart_extension_fields,
    )?;
    Ok((candidate, touched_components, plan))
}

fn publish_series_with_prepared_policy(
    package: &Package,
    changes: graph::Changes,
    budget: &mut TransactionBudget,
    fields: &mut FieldLedger,
    plan: &slide_chart_font::MutationPlan,
    extra_deletions: &[&str],
    is_rewritable_type: fn(u32) -> bool,
    chart_extension_fields: &[u32],
) -> Result<Package, Error> {
    let catalog = match &package.state.source {
        PhysicalSource::Package(catalog) => catalog,
        PhysicalSource::Semantic(_) => return Err(Refusal::UnsupportedSource.into()),
    };
    let archive_limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = catalog
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let edit_capacity = changes.edits.len();
    let insertion_capacity = changes.insertions.len();
    let deletion_capacity = changes
        .deletions
        .len()
        .checked_add(extra_deletions.len())
        .ok_or(Error::Overflow)?;
    let collection_scratch = edit_capacity
        .checked_mul(size_of::<(&str, Vec<u8>)>())
        .and_then(|value| value.checked_add(edit_capacity.checked_mul(size_of::<EntryEdit<'_>>())?))
        .and_then(|value| {
            value.checked_add(catalog.package().len().checked_mul(size_of::<&str>())?)
        })
        .and_then(|value| {
            value.checked_add(insertion_capacity.checked_mul(size_of::<(
                &str,
                Vec<u8>,
                Option<usize>,
            )>())?)
        })
        .and_then(|value| {
            value.checked_add(insertion_capacity.checked_mul(size_of::<EntryInsertion<'_>>())?)
        })
        .and_then(|value| value.checked_add(deletion_capacity.checked_mul(size_of::<&str>())?))
        .and_then(|value| {
            value.checked_add(edit_capacity.checked_mul(size_of::<PreparedPublication<'_>>())?)
        })
        .and_then(|value| {
            value.checked_add(insertion_capacity.checked_mul(size_of::<PreparedPublication<'_>>())?)
        })
        .ok_or(Error::Overflow)?;
    let collection_allocations = usize::from(edit_capacity != 0) * 3
        + usize::from(insertion_capacity != 0) * 3
        + usize::from(deletion_capacity != 0)
        + usize::from(!catalog.package().is_empty());
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
        .chain(extra_deletions.iter().map(|name| name.len()))
        .try_fold(0usize, |total, bytes| total.checked_add(bytes))
        .ok_or(Error::Overflow)?;
    budget
        .allocations(collection_allocations)
        .and_then(|_| budget.scratch(collection_scratch))
        .and_then(|_| budget.work(name_bytes))
        .map_err(map_transaction_error)?;

    let mut edited = Vec::new();
    edited
        .try_reserve_exact(edit_capacity)
        .map_err(|_| Error::Allocation {
            amount: edit_capacity,
        })?;
    let mut inserted = Vec::new();
    inserted
        .try_reserve_exact(insertion_capacity)
        .map_err(|_| Error::Allocation {
            amount: insertion_capacity,
        })?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(edit_capacity)
        .map_err(|_| Error::Allocation {
            amount: edit_capacity,
        })?;
    let mut insertions = Vec::new();
    insertions
        .try_reserve_exact(insertion_capacity)
        .map_err(|_| Error::Allocation {
            amount: insertion_capacity,
        })?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(catalog.package().len())
        .map_err(|_| Error::Allocation {
            amount: catalog.package().len(),
        })?;
    names.extend(catalog.package().iter().map(|entry| entry.name()));
    let mut deletions = Vec::new();
    deletions
        .try_reserve_exact(deletion_capacity)
        .map_err(|_| Error::Allocation {
            amount: deletion_capacity,
        })?;
    deletions.extend(changes.deletions.iter().map(String::as_str));
    for name in extra_deletions {
        if !deletions.contains(name) {
            deletions.push(name);
        }
    }

    let mut prepared_edits = Vec::new();
    prepared_edits
        .try_reserve_exact(edit_capacity)
        .map_err(|_| Error::Allocation {
            amount: edit_capacity,
        })?;
    for (name, archive) in &changes.edits {
        prepared_edits.push(prepare_publication(
            archive,
            name.as_str(),
            None,
            archive_limits,
            snappy_limits,
            budget,
        )?);
    }
    let mut prepared_insertions = Vec::new();
    prepared_insertions
        .try_reserve_exact(insertion_capacity)
        .map_err(|_| Error::Allocation {
            amount: insertion_capacity,
        })?;
    for insertion in &changes.insertions {
        let position = match insertion.before.as_deref() {
            Some(before) => Some(
                names
                    .iter()
                    .position(|name| *name == before)
                    .ok_or(Error::InvalidSource)?,
            ),
            None => None,
        };
        prepared_insertions.push(prepare_publication(
            &insertion.archive,
            insertion.name.as_str(),
            position,
            archive_limits,
            snappy_limits,
            budget,
        )?);
    }

    for prepared in prepared_edits {
        let name = prepared.name;
        edited.push((name, execute_publication(prepared)?));
    }
    edits.extend(
        edited
            .iter()
            .map(|(name, data)| EntryEdit::new(name, data.as_slice())),
    );
    for prepared in prepared_insertions {
        let name = prepared.name;
        let position = prepared.position;
        inserted.push((name, execute_publication(prepared)?, position));
    }
    insertions.extend(
        inserted
            .iter()
            .map(|(name, data, position)| match position {
                Some(position) => EntryInsertion::at(*position, name, data.as_slice()),
                None => EntryInsertion::new(name, data.as_slice()),
            }),
    );

    let output = slide_chart_font::reassemble_prepared_payloads(
        package,
        budget,
        &insertions,
        &edits,
        &deletions,
    )
    .map_err(map_font_error)?;
    let output_len = output.len();
    slide_chart_font::preflight_locality_verification(package, output_len, plan, budget)
        .map_err(map_font_error)?;
    slide_chart_font::preflight_candidate(package, output_len, changes.insertions.len(), budget)
        .map_err(map_font_error)?;
    let candidate = package.structure_reopen(output).map_err(map_read_error)?;
    slide_chart_font::verify_locality_members(
        package,
        &candidate,
        plan,
        budget,
        fields,
        is_rewritable_type,
        chart_extension_fields,
    )
    .map_err(map_font_error)?;
    Ok(candidate)
}

// ---------------------------------------------------------------------------
// The shared core of the Keynote chart-property families (series, style,
// axis): one transaction ledger, the bounded workspace, exact-source
// publication under the engine's mutation scope and candidate reopen.
// ---------------------------------------------------------------------------

/// The selected chart with its kind and visible series count.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Target {
    pub(super) selection: Selection,
    pub(super) kind: ChartKind,
    pub(super) series: usize,
}

/// The transaction ledger of one chart-property operation.
pub(super) struct Ledger {
    budget: TransactionBudget,
    fields: FieldLedger,
}

impl Ledger {
    /// Admit one operation on `package`.
    pub(super) fn new(package: &Package) -> Result<Self, Error> {
        Ok(Self {
            budget: transaction_budget(package)?,
            fields: transaction_fields(package)?,
        })
    }

    /// Resolve the selected chart, its kind and its series count.
    pub(super) fn target(
        &mut self,
        package: &Package,
        slide: SlideSelector<'_>,
        chart: ChartSelector<'_>,
    ) -> Result<Target, Error> {
        let selection = select(package, slide, chart, &mut self.budget, &mut self.fields)?;
        let (kind, series) = kind_and_count(package, selection, &mut self.budget)?;
        Ok(Target {
            selection,
            kind,
            series,
        })
    }

    /// Run one engine read over a bounded workspace of `package`.
    pub(super) fn read<T>(
        &mut self,
        package: &Package,
        read: impl FnOnce(&Workspace<'_>, engine::Limits) -> engine::SymbolResult<(T, engine::Report)>,
    ) -> Result<T, Error> {
        let WorkspaceAdmission {
            components,
            requirements,
            limits,
        } = admit_workspace(package, &mut self.budget)?;
        let workspace =
            Workspace::new_bounded(&*components, graph_limits(package)?, requirements, limits)
                .map_err(map_construction)?;
        let limits = series_limits(package, &self.budget, &self.fields)?;
        let (value, report) = read(&workspace, limits).map_err(map_series_error)?;
        charge_series_report(&mut self.budget, &mut self.fields, report)?;
        Ok(value)
    }

    /// Run one engine write over a bounded workspace of `source` and publish
    /// its changes under the engine's mutation scope; `None` when nothing
    /// changes. `chart_extension_fields` are the chart-drawable extension
    /// fields the write may change.
    pub(super) fn write(
        &mut self,
        source: &Package,
        target: &Target,
        chart_extension_fields: &[u32],
        write: impl FnOnce(
            &mut Workspace<'_>,
            engine::Limits,
        ) -> engine::SymbolResult<(bool, engine::Report)>,
    ) -> Result<Option<Published>, Error> {
        if !source.structure_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        self.budget
            .preflight_semantic_scan(source)
            .map_err(map_transaction_error)?;
        let WorkspaceAdmission {
            components,
            requirements,
            limits,
        } = admit_workspace(source, &mut self.budget)?;
        let mut workspace =
            Workspace::new_bounded(&*components, graph_limits(source)?, requirements, limits)
                .map_err(map_construction)?;
        let limits = series_limits(source, &self.budget, &self.fields)?;
        let (changed, report) = write(&mut workspace, limits).map_err(map_series_error)?;
        if !changed {
            charge_series_report(&mut self.budget, &mut self.fields, report)?;
            return Ok(None);
        }
        let scope = report.scope().cloned().ok_or(Error::Verification)?;
        charge_series_report(&mut self.budget, &mut self.fields, report)?;
        let changes = finish_workspace(workspace, &mut self.budget)?;
        let (package, touched, plan) = publish_changes(
            source,
            changes,
            target.selection.identifier,
            &scope,
            &mut self.budget,
            &mut self.fields,
            chart_extension_fields,
        )?;
        Ok(Some(Published {
            package,
            touched,
            plan,
        }))
    }

    /// Reopen the exact target of a patch over `package` and verify its
    /// locality against the patch's mutation plan.
    pub(super) fn reopen(
        &mut self,
        package: &Package,
        artifacts: &ExactArtifacts,
        plan: &slide_chart_font::MutationPlan,
        chart_extension_fields: &[u32],
    ) -> Result<Package, Error> {
        let target = artifacts.target();
        slide_chart_font::preflight_candidate(package, target.len(), 0, &mut self.budget)
            .map_err(map_font_error)?;
        let candidate = package
            .structure_reopen_shared(target)
            .map_err(map_read_error)?;
        slide_chart_font::verify_locality_members(
            package,
            &candidate,
            plan,
            &mut self.budget,
            &mut self.fields,
            is_series_rewritable_type,
            chart_extension_fields,
        )
        .map_err(map_font_error)?;
        Ok(candidate)
    }
}

/// A published candidate of the shared core.
pub(super) struct Published {
    pub(super) package: Package,
    pub(super) touched: usize,
    pub(super) plan: slide_chart_font::MutationPlan,
}

/// The exact artifacts of a commit from `source` to `target`.
pub(super) fn artifacts(source: &Package, target: &Package) -> Result<ExactArtifacts, Error> {
    let source = physical_source(source)?;
    let target = physical_source(target)?;
    Ok(ExactArtifacts::new(source, target))
}

/// The exact no-op artifacts of `source`.
pub(super) fn noop_artifacts(source: &Package) -> Result<ExactArtifacts, Error> {
    let source = physical_source(source)?;
    Ok(ExactArtifacts::new(Arc::clone(&source), source))
}

/// Whether `artifacts` were produced from `package`'s exact bytes.
pub(super) fn authorizes(package: &Package, artifacts: &ExactArtifacts) -> Result<bool, Error> {
    Ok(artifacts.authorizes_source(&physical_source(package)?))
}

/// The chart-drawable extension fields a series write may change (the
/// sparse series override and private-style arrays).
const SERIES_CHART_FIELDS: &[u32] = &[17, 19];

// ---------------------------------------------------------------------------
// The per-series family.
// ---------------------------------------------------------------------------

/// A compact exact-source patch for one per-series chart property.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    slide: Position,
    chart: Position,
    before: Property,
    after: Property,
    plan: slide_chart_font::MutationPlan,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("slide", &self.slide)
            .field("chart", &self.chart)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch is an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide: self.slide,
            chart: self.chart,
            before: self.after.clone(),
            after: self.before.clone(),
            plan: slide_chart_font::inverse_mutation_plan(&self.plan),
        }
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide
    }

    /// Return the selected semantic chart position.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart
    }

    /// Borrow the source property.
    #[must_use]
    pub const fn before(&self) -> &Property {
        &self.before
    }

    /// Borrow the target property.
    #[must_use]
    pub const fn after(&self) -> &Property {
        &self.after
    }

    /// Return the exact-source fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the exact-target fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
}

/// One verified chart-series commit.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the candidate package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Borrow the exact patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Consume the commit and return its package and exact patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Consume this commit.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
}

/// One staged per-series operation.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    Set(Property),
    Reset(PropertyKind, usize),
}

/// One immutable selector-first chart-series edit.
pub struct Edit<'a> {
    source: &'a Package,
    slide: super::edit::OwnedSlideSelector,
    chart: super::edit::OwnedChartSelector,
    staged: Option<Staged>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("staged", &self.staged.is_some())
            .finish_non_exhaustive()
    }
}

/// Own the slide and chart selectors of an edit within the source's text
/// ceiling.
pub(super) fn owned_selectors<'s, 'c>(
    source: &Package,
    slide: impl Into<SlideSelector<'s>>,
    chart: impl Into<ChartSelector<'c>>,
) -> Result<
    (
        super::edit::OwnedSlideSelector,
        super::edit::OwnedChartSelector,
    ),
    Error,
> {
    let mut retained = 0;
    let maximum = source.semantic_limits().max_text_bytes();
    let slide =
        super::edit::OwnedSlideSelector::from_selector(slide.into(), maximum, &mut retained)
            .map_err(map_selector_error)?;
    let chart =
        super::edit::OwnedChartSelector::from_selector(chart.into(), maximum, &mut retained)
            .map_err(map_selector_error)?;
    Ok((slide, chart))
}

impl<'a> Edit<'a> {
    fn new<'s, 'c>(
        source: &'a Package,
        slide: impl Into<SlideSelector<'s>>,
        chart: impl Into<ChartSelector<'c>>,
    ) -> Result<Self, Error> {
        let (slide, chart) = owned_selectors(source, slide, chart)?;
        Ok(Self {
            source,
            slide,
            chart,
            staged: None,
        })
    }

    /// Stage `property` (one value per visible series) as the chart's
    /// value, replacing any staged operation.
    ///
    /// # Errors
    ///
    /// [`Refusal::SeriesCount`] for a value naming no series.
    pub fn set(&mut self, property: Property) -> Result<&mut Self, Error> {
        if property.is_empty() {
            return Err(Refusal::SeriesCount.into());
        }
        self.staged = Some(Staged::Set(property));
        Ok(self)
    }

    /// Stage removing series `series`' own value of a series-style property
    /// (strokes, fills, symbol fills, symbol outlines), so it inherits its
    /// theme style's again (the host's `reset_slide_chart_series_*`).
    ///
    /// # Errors
    ///
    /// [`Refusal::NotResettable`] for the other properties.
    pub fn reset(&mut self, kind: PropertyKind, series: usize) -> Result<&mut Self, Error> {
        if kind.style().is_none() {
            return Err(Refusal::NotResettable.into());
        }
        self.staged = Some(Staged::Reset(kind, series));
        Ok(self)
    }

    /// Validate and atomically publish the staged operation. A value equal
    /// to the chart's is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::UnsupportedKind`],
    /// [`Refusal::SeriesCount`], [`Refusal::SeriesNotFound`],
    /// [`Refusal::SharedStyle`], [`Refusal::ImageFill`], source, limit and
    /// verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        let staged = self.staged.ok_or(Refusal::NoStagedOperation)?;
        let mut ledger = Ledger::new(self.source)?;
        let target = ledger.target(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
        )?;
        let kind = match &staged {
            Staged::Set(property) => property.kind(),
            Staged::Reset(kind, _) => *kind,
        };
        let before = read_property(&mut ledger, self.source, &target, kind)?;
        let published = match &staged {
            Staged::Set(property) => {
                if property.len() != target.series {
                    return Err(Refusal::SeriesCount.into());
                }
                if before == *property {
                    None
                } else {
                    ledger.write(
                        self.source,
                        &target,
                        SERIES_CHART_FIELDS,
                        |workspace, limits| write_property(workspace, &target, property, limits),
                    )?
                }
            },
            Staged::Reset(kind, series) => {
                if *series >= target.series {
                    return Err(Refusal::SeriesNotFound { series: *series }.into());
                }
                let style_kind = kind.style().ok_or(Refusal::NotResettable)?;
                let mut values = before.style_values().ok_or(Error::InvalidSource)?;
                values[*series] = None;
                ledger.write(
                    self.source,
                    &target,
                    SERIES_CHART_FIELDS,
                    |workspace, limits| {
                        style_engine::write(
                            workspace,
                            target.selection.identifier,
                            target.kind,
                            style_kind,
                            &values,
                            limits,
                        )
                    },
                )?
            },
        };
        let Some(published) = published else {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: noop_artifacts(self.source)?,
                    slide: target.selection.slide,
                    chart: target.selection.chart,
                    before: before.clone(),
                    after: before,
                    plan: slide_chart_font::MutationPlan::default(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        };
        let candidate_target = ledger.target(
            &published.package,
            self.slide.as_selector(),
            self.chart.as_selector(),
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
                artifacts: artifacts(self.source, &published.package)?,
                slide: target.selection.slide,
                chart: target.selection.chart,
                before,
                after,
                plan: published.plan,
            },
            package: published.package,
            diagnostics: Diagnostics::published(published.touched),
        })
    }
}

fn connection_lines(kind: ChartKind) -> engine::SymbolResult<series_property::ConnectionLines> {
    ConnectionLineFamily::for_chart_kind(kind)
        .map(series_property::ConnectionLines)
        .ok_or(engine::SymbolError::Refused(
            engine::Refusal::UnsupportedKind,
        ))
}

fn symbol_family(kind: ChartKind) -> engine::SymbolResult<SymbolFamily> {
    SymbolFamily::for_chart_kind(kind).ok_or(engine::SymbolError::Refused(
        engine::Refusal::UnsupportedKind,
    ))
}

fn value_labels(kind: ChartKind) -> engine::SymbolResult<series_property::ValueLabels> {
    series_property::ValueLabels::for_chart_kind(kind).ok_or(engine::SymbolError::Refused(
        engine::Refusal::UnsupportedKind,
    ))
}

fn value_label_affixes(
    kind: ChartKind,
) -> engine::SymbolResult<series_property::ValueLabelAffixes> {
    series_property::ValueLabelAffixes::for_chart_kind(kind).ok_or(engine::SymbolError::Refused(
        engine::Refusal::UnsupportedKind,
    ))
}

fn value_label_formats(
    kind: ChartKind,
) -> engine::SymbolResult<series_property::ValueLabelNumberFormats> {
    series_property::ValueLabelNumberFormats::for_chart_kind(kind).ok_or(
        engine::SymbolError::Refused(engine::Refusal::UnsupportedKind),
    )
}

/// Read per-series property `kind` of the target chart.
fn read_property(
    ledger: &mut Ledger,
    package: &Package,
    target: &Target,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let chart = target.selection.identifier;
    let count = target.series;
    let chart_kind = target.kind;
    ledger.read(package, |workspace, limits| {
        if let Some(style_kind) = kind.style() {
            let (values, report) =
                style_engine::read(workspace, chart, chart_kind, count, style_kind, limits)?;
            return Ok((from_style_values(kind, values)?, report));
        }
        Ok(match kind {
            PropertyKind::Symbols => {
                let (values, report) =
                    engine::symbols(workspace, chart, symbol_family(chart_kind)?, count, limits)?;
                (Property::Symbols(values), report)
            },
            PropertyKind::ConnectionLines => {
                let (values, report) = engine::read(
                    workspace,
                    chart,
                    count,
                    &connection_lines(chart_kind)?,
                    limits,
                )?;
                (Property::ConnectionLines(values), report)
            },
            PropertyKind::WedgeExplosions => {
                let (values, report) = engine::read(
                    workspace,
                    chart,
                    count,
                    &series_property::WedgeExplosions,
                    limits,
                )?;
                (Property::WedgeExplosions(values), report)
            },
            PropertyKind::LeaderLines => {
                let (values, report) = engine::read(
                    workspace,
                    chart,
                    count,
                    &series_property::LeaderLines,
                    limits,
                )?;
                (Property::LeaderLines(values), report)
            },
            PropertyKind::LabelDistances => {
                let (values, report) = engine::read(
                    workspace,
                    chart,
                    count,
                    &series_property::LabelDistances,
                    limits,
                )?;
                (Property::LabelDistances(values), report)
            },
            PropertyKind::Trendlines => {
                let (values, report) = engine::read(
                    workspace,
                    chart,
                    count,
                    &series_property::Trendlines,
                    limits,
                )?;
                (Property::Trendlines(values), report)
            },
            PropertyKind::ErrorBars => {
                let (values, report) =
                    engine::read(workspace, chart, count, &series_property::ErrorBars, limits)?;
                (Property::ErrorBars(values), report)
            },
            PropertyKind::PieLabels => {
                let (values, report) =
                    engine::read(workspace, chart, count, &series_property::PieLabels, limits)?;
                (Property::PieLabels(values), report)
            },
            PropertyKind::ValueLabels => {
                let (values, report) =
                    engine::read(workspace, chart, count, &value_labels(chart_kind)?, limits)?;
                (Property::ValueLabels(values), report)
            },
            PropertyKind::ValueLabelAffixes => {
                let (values, report) = engine::read(
                    workspace,
                    chart,
                    count,
                    &value_label_affixes(chart_kind)?,
                    limits,
                )?;
                (Property::ValueLabelAffixes(values), report)
            },
            PropertyKind::ValueLabelNumberFormats => {
                let (values, report) = engine::read(
                    workspace,
                    chart,
                    count,
                    &value_label_formats(chart_kind)?,
                    limits,
                )?;
                (Property::ValueLabelNumberFormats(values), report)
            },
            PropertyKind::Strokes
            | PropertyKind::Fills
            | PropertyKind::ValueLabelLocations
            | PropertyKind::ErrorBarAutoFits
            | PropertyKind::ValueLabelAutoFits
            | PropertyKind::SymbolFills
            | PropertyKind::SymbolOutlines => {
                return Err(engine::SymbolError::Invalid("a series-style kind"));
            },
        })
    })
}

/// Write per-series `property` to the target chart.
fn write_property(
    workspace: &mut Workspace<'_>,
    target: &Target,
    property: &Property,
    limits: engine::Limits,
) -> engine::SymbolResult<(bool, engine::Report)> {
    let chart = target.selection.identifier;
    let chart_kind = target.kind;
    if let (Some(style_kind), Some(values)) = (property.kind().style(), property.style_values()) {
        return style_engine::write(workspace, chart, chart_kind, style_kind, &values, limits);
    }
    match property {
        Property::Symbols(values) => {
            engine::set_symbols(workspace, chart, symbol_family(chart_kind)?, values, limits)
        },
        Property::ConnectionLines(values) => engine::write(
            workspace,
            chart,
            &connection_lines(chart_kind)?,
            values,
            limits,
        ),
        Property::WedgeExplosions(values) => engine::write(
            workspace,
            chart,
            &series_property::WedgeExplosions,
            values,
            limits,
        ),
        Property::LeaderLines(values) => engine::write(
            workspace,
            chart,
            &series_property::LeaderLines,
            values,
            limits,
        ),
        Property::LabelDistances(values) => engine::write(
            workspace,
            chart,
            &series_property::LabelDistances,
            values,
            limits,
        ),
        Property::Trendlines(values) => engine::write(
            workspace,
            chart,
            &series_property::Trendlines,
            values,
            limits,
        ),
        Property::ErrorBars(values) => engine::write(
            workspace,
            chart,
            &series_property::ErrorBars,
            values,
            limits,
        ),
        Property::PieLabels(values) => engine::write(
            workspace,
            chart,
            &series_property::PieLabels,
            values,
            limits,
        ),
        Property::ValueLabels(values) => {
            engine::write(workspace, chart, &value_labels(chart_kind)?, values, limits)
        },
        Property::ValueLabelAffixes(values) => engine::write(
            workspace,
            chart,
            &value_label_affixes(chart_kind)?,
            values,
            limits,
        ),
        Property::ValueLabelNumberFormats(values) => engine::write(
            workspace,
            chart,
            &value_label_formats(chart_kind)?,
            values,
            limits,
        ),
        Property::Strokes(_)
        | Property::Fills(_)
        | Property::ValueLabelLocations(_)
        | Property::ErrorBarAutoFits(_)
        | Property::ValueLabelAutoFits(_)
        | Property::SymbolFills(_)
        | Property::SymbolOutlines(_) => Err(engine::SymbolError::Invalid("a series-style value")),
    }
}

fn from_style_values(
    kind: PropertyKind,
    values: Vec<style_engine::Value>,
) -> engine::SymbolResult<Property> {
    let invalid = || engine::SymbolError::Invalid("a series-style value of another property");
    Ok(match kind {
        PropertyKind::Strokes | PropertyKind::SymbolOutlines => {
            let strokes = values
                .into_iter()
                .map(|value| match value {
                    style_engine::Value::Stroke(stroke) => Ok(stroke),
                    _ => Err(invalid()),
                })
                .collect::<engine::SymbolResult<Vec<_>>>()?;
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
                .collect::<engine::SymbolResult<_>>()?,
        ),
        PropertyKind::ValueLabelLocations => Property::ValueLabelLocations(
            values
                .into_iter()
                .map(|value| match value {
                    style_engine::Value::ValueLabelLocation(fill) => Ok(fill),
                    _ => Err(invalid()),
                })
                .collect::<engine::SymbolResult<_>>()?,
        ),
        PropertyKind::ErrorBarAutoFits => Property::ErrorBarAutoFits(
            values
                .into_iter()
                .map(|value| match value {
                    style_engine::Value::AutoFit(fill) => Ok(fill),
                    _ => Err(invalid()),
                })
                .collect::<engine::SymbolResult<_>>()?,
        ),
        PropertyKind::ValueLabelAutoFits => Property::ValueLabelAutoFits(
            values
                .into_iter()
                .map(|value| match value {
                    style_engine::Value::AutoFit(fill) => Ok(fill),
                    _ => Err(invalid()),
                })
                .collect::<engine::SymbolResult<_>>()?,
        ),
        PropertyKind::SymbolFills => Property::SymbolFills(
            values
                .into_iter()
                .map(|value| match value {
                    style_engine::Value::SymbolFill(fill) => Ok(fill),
                    _ => Err(invalid()),
                })
                .collect::<engine::SymbolResult<_>>()?,
        ),
        _ => return Err(invalid()),
    })
}

fn physical_source(package: &Package) -> Result<Arc<[u8]>, Error> {
    match &package.state.source {
        PhysicalSource::Package(catalog) => Ok(catalog.shared_source()),
        PhysicalSource::Semantic(_) => Err(Refusal::UnsupportedSource.into()),
    }
}

impl Package {
    /// Read one per-series property of one slide chart, one value per
    /// visible series in series order.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedKind`], source and limit
    /// errors.
    pub fn slide_chart_series<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let mut ledger = Ledger::new(self)?;
        let target = ledger.target(self, slide.into(), chart.into())?;
        read_property(&mut ledger, self, &target, kind)
    }

    /// Start an exact-source per-series chart edit.
    ///
    /// ```no_run
    /// use litchi_keynote::Package;
    /// use litchi_keynote::chart::series::PropertyKind;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.key")?;
    /// let mut edit = package.edit_slide_chart_series(0, 0)?;
    /// edit.reset(PropertyKind::Fills, 0)?;
    /// let commit = edit.commit()?;
    /// let restored = commit.package().apply_slide_chart_series(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector ceiling errors.
    pub fn edit_slide_chart_series<'slide, 'chart>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        chart: impl Into<ChartSelector<'chart>>,
    ) -> Result<Edit<'_>, Error> {
        Edit::new(self, slide, chart)
    }

    /// Apply a patch only to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source; locality and verification errors.
    pub fn apply_slide_chart_series(&self, patch: &Patch) -> Result<Commit, Error> {
        let mut ledger = Ledger::new(self)?;
        if !authorizes(self, &patch.artifacts)? {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let candidate = ledger.reopen(self, &patch.artifacts, &patch.plan, SERIES_CHART_FIELDS)?;
        let target = ledger.target(
            &candidate,
            SlideSelector::position(patch.slide),
            ChartSelector::position(patch.chart),
        )?;
        let after = read_property(&mut ledger, &candidate, &target, patch.after.kind())?;
        if after != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(0),
        })
    }
}

#[cfg(test)]
mod limit_mapping_tests {
    use super::*;

    fn assert_limit(error: Error, kind: LimitKind, observed: u64, maximum: u64) {
        assert_eq!(
            error,
            Error::Limit {
                kind,
                observed,
                maximum,
            }
        );
    }

    #[test]
    fn preserves_archive_and_snappy_units_and_values() {
        const OBSERVED: usize = 17;
        const MAXIMUM: usize = 11;

        assert_limit(
            map_graph(graph::Error::Archive(litchi_iwa_core::Error::Limit {
                kind: litchi_iwa_core::LimitKind::ArchiveBytes,
                observed: OBSERVED,
                maximum: MAXIMUM,
            })),
            LimitKind::TotalBytes,
            OBSERVED as u64,
            MAXIMUM as u64,
        );
        assert_limit(
            map_graph(graph::Error::Archive(litchi_iwa_core::Error::Limit {
                kind: litchi_iwa_core::LimitKind::SnappyFrames,
                observed: OBSERVED,
                maximum: MAXIMUM,
            })),
            LimitKind::SnappyFrames,
            OBSERVED as u64,
            MAXIMUM as u64,
        );
    }

    #[test]
    fn preserves_nested_archive_and_detection_units_and_values() {
        const OBSERVED: usize = 31;
        const MAXIMUM: usize = 23;

        assert_limit(
            map_archive_error(litchi_iwa_archive::Error::Iwa(
                litchi_iwa_core::Error::Limit {
                    kind: litchi_iwa_core::LimitKind::SnappyFrames,
                    observed: OBSERVED,
                    maximum: MAXIMUM,
                },
            )),
            LimitKind::SnappyFrames,
            OBSERVED as u64,
            MAXIMUM as u64,
        );
        assert_eq!(
            map_archive_error(litchi_iwa_archive::Error::Iwa(
                litchi_iwa_core::Error::Allocation {
                    resource: "nested IWA",
                    requested: OBSERVED,
                },
            )),
            Error::Allocation { amount: OBSERVED },
        );
        assert_limit(
            map_read_error(ReadError::Detection(
                litchi_iwa_detect::Error::LimitExceeded {
                    kind: litchi_iwa_detect::LimitKind::IwaFields,
                    observed: OBSERVED as u64,
                    maximum: MAXIMUM as u64,
                },
            )),
            LimitKind::IwaFields,
            OBSERVED as u64,
            MAXIMUM as u64,
        );
        assert_limit(
            map_read_error(ReadError::Detection(litchi_iwa_detect::Error::IwaCommon(
                litchi_iwa_common::Error::LimitExceeded {
                    kind: litchi_iwa_common::LimitKind::Fields,
                    observed: OBSERVED,
                    limit: MAXIMUM,
                },
            ))),
            LimitKind::WireFields,
            OBSERVED as u64,
            MAXIMUM as u64,
        );
        assert_eq!(
            map_read_error(ReadError::Detection(litchi_iwa_detect::Error::Allocation {
                amount: OBSERVED
            },)),
            Error::Allocation { amount: OBSERVED },
        );
    }

    #[test]
    fn preserves_shared_font_and_transaction_byte_units_and_values() {
        const OBSERVED: u64 = 23;
        const MAXIMUM: u64 = 19;

        for (kind, expected) in [
            (
                chart_axis_support::AxisSupportLimitKind::EntryBytes,
                LimitKind::EntryBytes,
            ),
            (
                chart_axis_support::AxisSupportLimitKind::TotalBytes,
                LimitKind::TotalBytes,
            ),
            (
                chart_axis_support::AxisSupportLimitKind::Slides,
                LimitKind::Slides,
            ),
            (
                chart_axis_support::AxisSupportLimitKind::TextBytes,
                LimitKind::TextBytes,
            ),
            (
                chart_axis_support::AxisSupportLimitKind::TitleBytes,
                LimitKind::TitleBytes,
            ),
        ] {
            assert_limit(
                map_axis_support_error(AxisSupportError::LimitExceeded {
                    kind,
                    observed: OBSERVED,
                    maximum: MAXIMUM,
                }),
                expected,
                OBSERVED,
                MAXIMUM,
            );
        }

        for kind in [
            slide_chart_font::LimitKind::EntryBytes,
            slide_chart_font::LimitKind::TotalBytes,
            slide_chart_font::LimitKind::Identifiers,
            slide_chart_font::LimitKind::TextBytes,
        ] {
            let expected = match kind {
                slide_chart_font::LimitKind::EntryBytes => LimitKind::EntryBytes,
                slide_chart_font::LimitKind::TotalBytes => LimitKind::TotalBytes,
                slide_chart_font::LimitKind::Identifiers => LimitKind::Identifiers,
                slide_chart_font::LimitKind::TextBytes => LimitKind::TextBytes,
                _ => unreachable!(),
            };
            assert_limit(
                map_font_error(slide_chart_font::Error::Limit {
                    kind,
                    observed: OBSERVED,
                    maximum: MAXIMUM,
                }),
                expected,
                OBSERVED,
                MAXIMUM,
            );
        }

        for kind in [
            super::super::slide_table_core::LimitKind::EntryBytes,
            super::super::slide_table_core::LimitKind::TotalBytes,
        ] {
            let expected = match kind {
                super::super::slide_table_core::LimitKind::EntryBytes => LimitKind::EntryBytes,
                super::super::slide_table_core::LimitKind::TotalBytes => LimitKind::TotalBytes,
                _ => unreachable!(),
            };
            assert_limit(
                map_transaction_error(TransactionError::Limit {
                    kind,
                    observed: OBSERVED,
                    maximum: MAXIMUM,
                }),
                expected,
                OBSERVED,
                MAXIMUM,
            );
        }
    }

    #[test]
    fn preserves_series_registry_and_identifier_units_and_values() {
        const OBSERVED: usize = 29;
        const MAXIMUM: usize = 23;

        assert_limit(
            map_series_error(engine::SymbolError::Limit {
                kind: engine::LimitKind::RegistryEntries,
                observed: OBSERVED,
                maximum: MAXIMUM,
            }),
            LimitKind::RegistryEntries,
            OBSERVED as u64,
            MAXIMUM as u64,
        );
        assert_limit(
            map_series_error(engine::SymbolError::Limit {
                kind: engine::LimitKind::Identifiers,
                observed: OBSERVED,
                maximum: MAXIMUM,
            }),
            LimitKind::Identifiers,
            OBSERVED as u64,
            MAXIMUM as u64,
        );
    }
}
