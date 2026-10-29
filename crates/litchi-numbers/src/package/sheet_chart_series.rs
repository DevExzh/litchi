//! Implementation of the public [`crate::sheet::chart::series`]
//! transaction: per-series chart properties of one sheet chart on W1-E's
//! style-variation graph engine (`litchi_iwa_style::series`, record 0375a),
//! record 0378.
//!
//! Each property holds one value per series, in series order (the chart's
//! row or column names, by its series direction). A write runs the engine
//! over a workspace of the package's components: overrides are patched,
//! copied, allocated or removed, the chart's sparse array, the stylesheet,
//! the object containers and the package metadata registry follow, and the
//! changes are republished once. The host editor's `set_sheet_chart_*`
//! all-series and one-series routes collapse into [`Edit::set`].

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;
use std::mem::size_of;

use litchi_iwa_archive::package::{EntryEdit, EntryInsertion};
use litchi_iwa_common::chart::error_bar::Series as ErrorBarSeries;
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::pie::{LabelDistance, LabelVisibility, WedgeExplosion};
use litchi_iwa_common::chart::series::{
    ConnectionLine, ConnectionLineFamily, Stroke, Symbol, SymbolFamily, SymbolFill, Trendline,
};
use litchi_iwa_common::chart::series_labels::Visibility;
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::chart_metadata_codec as metadata_codec;
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_style::series::{
    self as engine,
    property::{
        ConnectionLines, ErrorBars, LabelDistances, PieLabels, Trendlines, ValueLabels,
        WedgeExplosions,
    },
};
use thiserror::Error as ThisError;

use super::Package;
use super::chart_arrangement::{
    self as arrangement, ChartBudget, ChartSelection, parse_wire_view_with_budget,
    select_chart_with_budget,
};
use super::sheet_chart_series_style as style;
use super::sheet_drawable_graph as graph;
use super::structure::{self, Published, references};
use super::style_graph::{self, LeafComponents};
use crate::selector::{ChartSelector, SheetSelector};

structure::structural_family!("chart-series", Engine(litchi_iwa_style::series::LimitKind));

/// `TSCH.ChartDrawableArchive` extension holding the `TSCH.ChartArchive`.
const CHART_ARCHIVE_FIELD: u32 = 10_000;
/// `TSCH.ChartArchive.series_direction`.
const SERIES_DIRECTION_FIELD: u32 = 5;
const SERIES_BY_ROW: u64 = 1;
const SERIES_BY_COLUMN: u64 = 2;
const GRAPH_INDEX_SLOT_BYTES: usize = 64;
const GRAPH_AMBIGUOUS_SLOT_BYTES: usize = 32;
const GRAPH_EXTRA_ALLOCATIONS: usize = 4;
const ZIP_MEMBER_OVERHEAD_BYTES: usize = 128;

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
    /// How far each wedge's label sits from the pie centre.
    LabelDistances(Vec<LabelDistance>),
    /// The data symbol of each series (`None` hides it).
    Symbols(Vec<Option<Symbol>>),
    /// The trendline of each series.
    Trendlines(Vec<Trendline>),
    /// The error bars of each series.
    ErrorBars(Vec<ErrorBarSeries>),
    /// Which labels each pie wedge shows.
    PieLabels(Vec<LabelVisibility>),
    /// The value-label visibility of each series.
    ValueLabels(Vec<Visibility>),
    /// The stroke of each series (`None`: no stroke).
    Strokes(Vec<Option<Stroke>>),
    /// The fill of each series.
    Fills(Vec<Fill>),
    /// The fill of each series' data symbols.
    SymbolFills(Vec<SymbolFill>),
    /// The outline of each series' data symbols (`None`: no outline).
    SymbolOutlines(Vec<Option<Stroke>>),
}

/// The property a [`Property`] value sets, for reads.
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
    const fn style(self) -> Option<style::Kind> {
        match self {
            Self::Strokes => Some(style::Kind::Strokes),
            Self::Fills => Some(style::Kind::Fills),
            Self::SymbolFills => Some(style::Kind::SymbolFills),
            Self::SymbolOutlines => Some(style::Kind::SymbolOutlines),
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
            Self::LabelDistances(_) => PropertyKind::LabelDistances,
            Self::Symbols(_) => PropertyKind::Symbols,
            Self::Trendlines(_) => PropertyKind::Trendlines,
            Self::ErrorBars(_) => PropertyKind::ErrorBars,
            Self::PieLabels(_) => PropertyKind::PieLabels,
            Self::ValueLabels(_) => PropertyKind::ValueLabels,
            Self::Strokes(_) => PropertyKind::Strokes,
            Self::Fills(_) => PropertyKind::Fills,
            Self::SymbolFills(_) => PropertyKind::SymbolFills,
            Self::SymbolOutlines(_) => PropertyKind::SymbolOutlines,
        }
    }

    /// The per-series style values of a series-style property.
    fn style_values(&self) -> Option<Vec<style::Value>> {
        Some(match self {
            Self::Strokes(values) | Self::SymbolOutlines(values) => {
                values.iter().copied().map(style::Value::Stroke).collect()
            },
            Self::Fills(values) => values.iter().cloned().map(style::Value::Fill).collect(),
            Self::SymbolFills(values) => values
                .iter()
                .cloned()
                .map(style::Value::SymbolFill)
                .collect(),
            _ => return None,
        })
    }

    /// The number of series the value names.
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
            Self::Strokes(values) | Self::SymbolOutlines(values) => values.len(),
            Self::Fills(values) => values.len(),
            Self::SymbolFills(values) => values.len(),
        }
    }

    /// Whether the value names no series.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

/// A typed, content-redacted chart-series failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted chart graph or an override payload is malformed.
    #[error("rooted chart-series source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} chart-series items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("chart-series candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("chart-series patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The chart selector did not resolve one chart of the sheet.
    #[error("chart selector did not resolve")]
    ChartNotFound,
    /// The property does not apply to the chart's kind (connection lines of
    /// a pie chart, data symbols of a bar chart).
    #[error("the property does not apply to this chart kind")]
    UnsupportedKind,
    /// The value names another number of series than the chart has.
    #[error("the chart has {expected} series, the value names {given}")]
    SeriesCount {
        /// The chart's series count.
        expected: usize,
        /// The number of values given.
        given: usize,
    },
    /// The series is past the chart's series.
    #[error("the chart has no series {series}")]
    SeriesNotFound {
        /// The requested zero-based series.
        series: usize,
    },
    /// The property has no reset (only the series-style properties do).
    #[error("the property cannot be reset")]
    NotResettable,
    /// The series style is shared with another chart, which the write would
    /// change too.
    #[error("the series style is shared with another chart")]
    SharedStyle,
    /// An image fill needs the embedded-data graph this family does not
    /// write.
    #[error("image fills are not written by this family")]
    ImageFill,
    /// Commit was requested before staging a change.
    #[error("no chart-series change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph (a shared series
    /// style the write would have to change).
    #[error("a surviving object depends on the chart-series graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support chart-series edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::ChartNotFound | Self::SeriesNotFound { .. } => {
                RefusalKind::NotFound
            },
            Self::UnsupportedKind | Self::NotResettable => RefusalKind::UnsupportedTarget,
            Self::SharedStyle | Self::ImageFill => RefusalKind::UnsupportedContent,
            Self::SeriesCount { .. } => RefusalKind::InvalidValue,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::UnsupportedDependency | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        structure::map_wire(&error).into()
    }
}

impl From<litchi_iwa_style::graph::Error> for Error {
    fn from(error: litchi_iwa_style::graph::Error) -> Self {
        match style_graph::fault(&error) {
            style_graph::StyleGraphFault::Shared(fault) => fault.into(),
            style_graph::StyleGraphFault::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: match kind {
                    litchi_iwa_style::graph::GraphLimitKind::ScratchBytes => {
                        LimitKind::ScratchBytes
                    },
                    litchi_iwa_style::graph::GraphLimitKind::AllocationEvents => {
                        LimitKind::AllocationEvents
                    },
                },
                observed,
                maximum,
            },
        }
    }
}

impl From<litchi_iwa_style::series::SymbolError> for Error {
    fn from(error: litchi_iwa_style::series::SymbolError) -> Self {
        match error {
            litchi_iwa_style::series::SymbolError::Graph(error) => error.into(),
            litchi_iwa_style::series::SymbolError::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: LimitKind::Engine(kind),
                observed: observed as u64,
                maximum: maximum as u64,
            },
            litchi_iwa_style::series::SymbolError::Allocation { amount } => {
                Error::Allocation { amount }
            },
            litchi_iwa_style::series::SymbolError::Invalid(_)
            | litchi_iwa_style::series::SymbolError::IdentifierOverflow => Error::InvalidSource,
            _ => Error::InvalidSource,
        }
    }
}

/// Chart-series changes staged against one chart of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    chart: ChartSelector,
    changes: Vec<Staged>,
    budget: ChartBudget,
}

/// One staged change.
#[derive(Debug, Clone, PartialEq)]
enum Staged {
    Set(Property),
    Reset(PropertyKind, usize),
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("chart", &self.chart)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage `property` (one value per series) as the chart's value.
    /// Changes apply in order.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set(&mut self, property: Property) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(Staged::Set(property));
        Ok(self)
    }

    /// Stage removing series `series`' own value of a series-style property
    /// (strokes, fills, symbol fills, symbol outlines), so it inherits its
    /// theme style's again (the host's `reset_sheet_chart_series_*`).
    ///
    /// # Errors
    ///
    /// [`Refusal::NotResettable`] for the other properties,
    /// [`Error::Allocation`].
    pub fn reset(&mut self, kind: PropertyKind, series: usize) -> Result<&mut Self, Error> {
        if kind.style().is_none() {
            return Err(Refusal::NotResettable.into());
        }
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(Staged::Reset(kind, series));
        Ok(self)
    }

    /// Validate and atomically publish the staged changes. A change that
    /// leaves a property as it is is skipped; an edit whose changes are all
    /// skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::UnsupportedKind`],
    /// [`Refusal::SeriesCount`], source and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut budget = self.budget;
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for change in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let published = match change {
                Staged::Set(property) => {
                    apply_one(base, self.sheet, self.chart, property, &mut budget)?
                },
                Staged::Reset(kind, series) => {
                    reset_one(base, self.sheet, self.chart, *kind, *series, &mut budget)?
                },
            };
            if let Some(published) = published {
                if let Staged::Set(property) = change
                    && read(
                        &published.package,
                        self.sheet,
                        self.chart,
                        property.kind(),
                        &mut budget,
                    )? != *property
                {
                    return Err(Error::Verification);
                }
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(self.source);
        };
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        )
    }
}

/// The chart drawable's identifier, kind and series count.
struct Target {
    chart: u64,
    kind: ChartKind,
    series: usize,
}

fn target(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    budget: &mut ChartBudget,
) -> Result<Target, Error> {
    let selection = select_chart_with_budget(
        package,
        SheetSelector::position(sheet.into()),
        chart,
        false,
        budget,
    )
    .map_err(map_arrangement)?;
    target_from_selection(package, &selection, budget)
}

fn target_from_selection(
    package: &Package,
    selection: &ChartSelection,
    budget: &mut ChartBudget,
) -> Result<Target, Error> {
    let payload =
        arrangement::selected_chart_payload(package, selection).map_err(map_arrangement)?;
    let component = arrangement::package_component(package, selection.component_index)
        .map_err(map_arrangement)?;
    let object = component
        .archive()
        .objects
        .get(selection.object_index)
        .ok_or(Error::InvalidSource)?;
    let message = object
        .messages
        .get(selection.message_index)
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_reference_scan(&message.data)
        .map_err(map_arrangement)?;
    let archives = references::records(&message.data, &[CHART_ARCHIVE_FIELD])?;
    budget
        .charge_reference_scan(&message.data, archives.len())
        .map_err(map_arrangement)?;
    let [archive] = archives.as_slice() else {
        return Err(Error::InvalidSource);
    };
    let archive_view = parse_wire_view_with_budget(archive, 1, budget).map_err(map_arrangement)?;
    let direction = graph::unique_field(&archive_view, SERIES_DIRECTION_FIELD, 0)?
        .map(|field| {
            let (value, consumed) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                .map_err(|_| Error::InvalidSource)?;
            if consumed != field.payload().len()
                || litchi_iwa_common::varint::encoded_len(value) != consumed
            {
                return Err(Error::InvalidSource);
            }
            Ok(value)
        })
        .transpose()?;
    let options = budget.metadata_codec_options().map_err(map_arrangement)?;
    let (metadata, report) = match metadata_codec::decode_modern_with_report(payload, &options) {
        Ok(decoded) => decoded,
        Err(error) => {
            budget
                .metadata_codec_report(error.report())
                .map_err(map_arrangement)?;
            return Err(map_metadata_error(&error));
        },
    };
    budget
        .metadata_codec_report(report)
        .map_err(map_arrangement)?;
    let components = match direction {
        Some(SERIES_BY_ROW) => metadata.row_labels().len(),
        Some(SERIES_BY_COLUMN) => metadata.column_labels().len(),
        _ => return Err(Error::InvalidSource),
    };
    let kind = ChartKind::from_native(metadata.chart_type());
    let series = series_count(kind, components)?;
    Ok(Target {
        chart: selection.chart_identifier,
        kind,
        series,
    })
}

/// The visible series of `components` data rows or columns (the host's
/// `chart_series_count`): one per component, except a scatter (`X + Y`) or
/// bubble (`X + Y + size`) chart and the multi-data point charts, which show
/// one.
fn series_count(kind: ChartKind, components: usize) -> Result<usize, Error> {
    let count = match kind {
        ChartKind::Scatter2d if components == 2 => 1,
        ChartKind::Bubble2d if components == 3 => 1,
        ChartKind::Scatter2d | ChartKind::Bubble2d => return Err(Error::InvalidSource),
        ChartKind::MultiDataScatter2d | ChartKind::MultiDataBubble2d => 1,
        _ => components,
    };
    if count == 0 {
        return Err(Error::InvalidSource);
    }
    Ok(count)
}

fn map_arrangement(error: arrangement::Error) -> Error {
    match error {
        arrangement::Error::Refused(refusal) => match refusal {
            arrangement::Refusal::UnsupportedSource => Refusal::UnsupportedSource.into(),
            arrangement::Refusal::UnsupportedDependency => Refusal::UnsupportedDependency.into(),
            arrangement::Refusal::AmbiguousSelector => Error::InvalidSource,
            arrangement::Refusal::EmptySheetName
            | arrangement::Refusal::SheetNameNotFound
            | arrangement::Refusal::SheetPositionNotFound { .. } => Refusal::SheetNotFound.into(),
            arrangement::Refusal::ChartPositionNotFound { .. } => Refusal::ChartNotFound.into(),
        },
        arrangement::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                arrangement::LimitKind::WireBytes | arrangement::LimitKind::CompressedBytes => {
                    LimitKind::PayloadBytes
                },
                arrangement::LimitKind::WireFields => LimitKind::Engine(engine::LimitKind::Fields),
                arrangement::LimitKind::WireNesting => {
                    LimitKind::Engine(engine::LimitKind::Nesting)
                },
                arrangement::LimitKind::WireWork => LimitKind::Engine(engine::LimitKind::Work),
                arrangement::LimitKind::OutputBytes => {
                    LimitKind::Engine(engine::LimitKind::OutputBytes)
                },
                arrangement::LimitKind::Allocations => LimitKind::AllocationEvents,
                arrangement::LimitKind::RetainedBytes => {
                    LimitKind::Engine(engine::LimitKind::RetainedBytes)
                },
                arrangement::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                arrangement::LimitKind::References => {
                    LimitKind::Engine(engine::LimitKind::References)
                },
                arrangement::LimitKind::Identifiers => LimitKind::Identifiers,
            },
            observed,
            maximum,
        },
        arrangement::Error::Allocation { amount } => Error::Allocation { amount },
        arrangement::Error::InvalidSource => Error::InvalidSource,
        arrangement::Error::Verification => Error::Verification,
        arrangement::Error::PatchConflict => Error::PatchConflict,
    }
}

fn map_metadata_error(error: &metadata_codec::DecodeError) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        metadata_codec::DecodeLimit::Bytes { observed, maximum } => {
            (LimitKind::PayloadBytes, observed, maximum)
        },
        metadata_codec::DecodeLimit::Fields { observed, maximum } => (
            LimitKind::Engine(engine::LimitKind::Fields),
            observed,
            maximum,
        ),
        metadata_codec::DecodeLimit::Work { observed, maximum } => (
            LimitKind::Engine(engine::LimitKind::Work),
            observed,
            maximum,
        ),
        metadata_codec::DecodeLimit::Nesting { observed, maximum } => (
            LimitKind::Engine(engine::LimitKind::Nesting),
            observed as usize,
            maximum as usize,
        ),
        metadata_codec::DecodeLimit::Labels { observed, maximum }
        | metadata_codec::DecodeLimit::Text { observed, maximum } => (
            LimitKind::Engine(engine::LimitKind::RetainedBytes),
            observed,
            maximum,
        ),
        _ => return Error::InvalidSource,
    };
    Error::Limit {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

#[derive(Debug, Clone, Copy)]
struct GraphWorkspaceCost {
    work: usize,
    allocations: usize,
    scratch: usize,
}

fn graph_workspace_cost(package: &Package) -> Result<GraphWorkspaceCost, Error> {
    let catalog = package.state.components.catalog();
    let object_count = catalog.iter().try_fold(0usize, |count, component| {
        count
            .checked_add(component.archive().objects.len())
            .ok_or(Error::InvalidSource)
    })?;
    let component_count = catalog.len();
    let work = component_count
        .checked_add(object_count)
        .ok_or(Error::InvalidSource)?;
    let allocations = component_count
        .checked_add(object_count.checked_mul(2).ok_or(Error::InvalidSource)?)
        .and_then(|value| value.checked_add(GRAPH_EXTRA_ALLOCATIONS))
        .ok_or(Error::InvalidSource)?;
    let scratch = component_count
        .checked_mul(size_of::<(usize, usize)>())
        .and_then(|value| {
            value.checked_add(
                object_count.checked_mul(GRAPH_INDEX_SLOT_BYTES + GRAPH_AMBIGUOUS_SLOT_BYTES)?,
            )
        })
        .ok_or(Error::InvalidSource)?;
    Ok(GraphWorkspaceCost {
        work,
        allocations,
        scratch,
    })
}

fn prepare_components<'a>(
    package: &'a Package,
    budget: &mut ChartBudget,
) -> Result<(LeafComponents<'a>, litchi_iwa_core::Limits), Error> {
    admit_graph_workspace(package, budget)?;
    let components = LeafComponents::new(package).map_err(Error::from)?;
    let limits = structure::exact_source(package)?
        .limits()
        .effective_archive_limits()
        .map_err(|error| structure::map_archive(&error))?;
    Ok((components, limits))
}

fn admit_graph_workspace(package: &Package, budget: &mut ChartBudget) -> Result<(), Error> {
    let cost = graph_workspace_cost(package)?;
    budget.preflight_work(cost.work).map_err(map_arrangement)?;
    budget
        .preflight_allocations(cost.allocations)
        .map_err(map_arrangement)?;
    budget
        .preflight_scratch(cost.scratch)
        .map_err(map_arrangement)?;
    budget.work(cost.work).map_err(map_arrangement)?;
    budget
        .allocations(cost.allocations)
        .map_err(map_arrangement)?;
    budget.scratch(cost.scratch).map_err(map_arrangement)?;
    Ok(())
}

fn publish_changes_exact(
    package: &Package,
    changes: &litchi_iwa_style::graph::Changes,
    budget: &mut ChartBudget,
) -> Result<Published, Error> {
    let catalog = structure::exact_source(package)?;
    let archive_limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(|error| structure::map_archive(&error))?;
    let snappy_limits = catalog
        .limits()
        .snappy_limits()
        .map_err(|error| structure::map_archive(&error))?;
    let mut output_bound = package.source_bytes().len();
    let mut retained_bound = 0usize;
    let mut scratch_bound = 0usize;
    let mut work_bound = 0usize;
    let mut member_count = 0usize;
    let mut bounds = |name: &str, archive: &Archive| -> Result<(), Error> {
        let encoded = archive
            .encoded_len(archive_limits)
            .map_err(|error| structure::map_core(&error))?;
        let compressed = SnappyStream::maximum_compressed_len(encoded, snappy_limits)
            .map_err(|error| structure::map_core(&error))?;
        let framing = name
            .len()
            .checked_mul(2)
            .and_then(|value| value.checked_add(ZIP_MEMBER_OVERHEAD_BYTES))
            .ok_or(Error::InvalidSource)?;
        output_bound = output_bound
            .checked_add(compressed)
            .and_then(|value| value.checked_add(framing))
            .ok_or(Error::InvalidSource)?;
        retained_bound = retained_bound
            .checked_add(encoded)
            .and_then(|value| value.checked_add(compressed))
            .ok_or(Error::InvalidSource)?;
        scratch_bound = scratch_bound
            .checked_add(encoded)
            .and_then(|value| value.checked_add(compressed))
            .ok_or(Error::InvalidSource)?;
        work_bound = work_bound
            .checked_add(encoded)
            .and_then(|value| value.checked_add(compressed))
            .ok_or(Error::InvalidSource)?;
        member_count = member_count.checked_add(1).ok_or(Error::InvalidSource)?;
        Ok(())
    };
    for (name, archive) in &changes.edits {
        bounds(name, archive)?;
    }
    for insertion in &changes.insertions {
        bounds(&insertion.name, &insertion.archive)?;
    }
    let vector_allocations = member_count
        .checked_add(changes.deletions.len())
        .and_then(|value| value.checked_add(3))
        .ok_or(Error::InvalidSource)?;
    budget
        .preflight_output(output_bound)
        .map_err(map_arrangement)?;
    budget
        .preflight_allocations(vector_allocations)
        .map_err(map_arrangement)?;
    budget
        .preflight_scratch(scratch_bound)
        .map_err(map_arrangement)?;
    budget
        .preflight_retained(retained_bound)
        .map_err(map_arrangement)?;
    budget.preflight_work(work_bound).map_err(map_arrangement)?;
    budget.work(work_bound).map_err(map_arrangement)?;
    budget
        .allocations(vector_allocations)
        .map_err(map_arrangement)?;
    budget.retained(retained_bound).map_err(map_arrangement)?;
    budget.scratch(scratch_bound).map_err(map_arrangement)?;

    let mut edited = Vec::new();
    edited
        .try_reserve_exact(changes.edits.len())
        .map_err(|_| Error::Allocation {
            amount: changes.edits.len(),
        })?;
    for (name, archive) in &changes.edits {
        let bytes = archive
            .to_bytes(archive_limits)
            .map_err(|error| structure::map_core(&error))?;
        let compressed = SnappyStream::compress(&bytes, snappy_limits)
            .map_err(|error| structure::map_core(&error))?;
        edited.push((name.as_str(), compressed));
    }
    let names: Vec<&str> = catalog.package().iter().map(|entry| entry.name()).collect();
    let mut inserted = Vec::new();
    inserted
        .try_reserve_exact(changes.insertions.len())
        .map_err(|_| Error::Allocation {
            amount: changes.insertions.len(),
        })?;
    for insertion in &changes.insertions {
        let position = match &insertion.before {
            Some(before) => Some(
                names
                    .iter()
                    .position(|name| name == before)
                    .ok_or(Error::InvalidSource)?,
            ),
            None => None,
        };
        let bytes = insertion
            .archive
            .to_bytes(archive_limits)
            .map_err(|error| structure::map_core(&error))?;
        let compressed = SnappyStream::compress(&bytes, snappy_limits)
            .map_err(|error| structure::map_core(&error))?;
        inserted.push((insertion.name.as_str(), compressed, position));
    }
    let edits: Vec<EntryEdit<'_>> = edited
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect();
    let insertions: Vec<EntryInsertion<'_>> = inserted
        .iter()
        .map(|(name, data, position)| match position {
            Some(position) => EntryInsertion::at(*position, name, data),
            None => EntryInsertion::new(name, data),
        })
        .collect();
    let deletions: Vec<&str> = changes.deletions.iter().map(String::as_str).collect();
    let prepared = catalog
        .prepare_reassembly_with_changes(
            &insertions,
            &edits,
            &deletions,
            package.state.options.archive(),
        )
        .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    budget.reassembly(requirements).map_err(map_arrangement)?;
    budget
        .candidate_reopen(requirements.output_bytes())
        .map_err(map_arrangement)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(|_| Error::InvalidSource)?;
    let candidate = Package::from_owned_bytes_with_options(output, package.state.options)
        .map_err(|_| Error::Verification)?;
    Ok(Published {
        package: candidate,
        touched_components: edits.len() + insertions.len() + deletions.len(),
        deleted_previews: 0,
    })
}

fn lines(kind: ChartKind) -> Result<ConnectionLines, Error> {
    Ok(ConnectionLines(
        ConnectionLineFamily::for_chart_kind(kind).ok_or(Refusal::UnsupportedKind)?,
    ))
}

fn symbols(kind: ChartKind) -> Result<SymbolFamily, Error> {
    Ok(SymbolFamily::for_chart_kind(kind).ok_or(Refusal::UnsupportedKind)?)
}

fn value_labels(kind: ChartKind) -> Result<ValueLabels, Error> {
    Ok(ValueLabels::for_chart_kind(kind).ok_or(Refusal::UnsupportedKind)?)
}

fn read(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    kind: PropertyKind,
    budget: &mut ChartBudget,
) -> Result<Property, Error> {
    let target = target(package, sheet, chart, budget)?;
    if let Some(style_kind) = kind.style() {
        admit_graph_workspace(package, budget)?;
        let values = style::read(
            package,
            target.chart,
            target.kind,
            target.series,
            style_kind,
        )
        .map_err(from_style)?;
        return from_style_values(kind, values);
    }
    let (components, limits) = prepare_components(package, budget)?;
    let workspace = Workspace::new(&components, limits).map_err(Error::from)?;
    let (chart, count) = (target.chart, target.series);
    Ok(match kind {
        PropertyKind::ConnectionLines => Property::ConnectionLines(engine::values(
            &workspace,
            chart,
            count,
            &lines(target.kind)?,
        )?),
        PropertyKind::WedgeExplosions => {
            Property::WedgeExplosions(engine::values(&workspace, chart, count, &WedgeExplosions)?)
        },
        PropertyKind::LabelDistances => {
            Property::LabelDistances(engine::values(&workspace, chart, count, &LabelDistances)?)
        },
        PropertyKind::Symbols => {
            let limits = budget.series_limits();
            let (values, report) =
                engine::symbols(&workspace, chart, symbols(target.kind)?, count, limits)
                    .map_err(Error::from)?;
            budget
                .charge_series_report(report)
                .map_err(map_arrangement)?;
            Property::Symbols(values)
        },
        PropertyKind::Trendlines => {
            Property::Trendlines(engine::values(&workspace, chart, count, &Trendlines)?)
        },
        PropertyKind::ErrorBars => {
            Property::ErrorBars(engine::values(&workspace, chart, count, &ErrorBars)?)
        },
        PropertyKind::PieLabels => {
            Property::PieLabels(engine::values(&workspace, chart, count, &PieLabels)?)
        },
        PropertyKind::ValueLabels => Property::ValueLabels(engine::values(
            &workspace,
            chart,
            count,
            &value_labels(target.kind)?,
        )?),
        PropertyKind::Strokes
        | PropertyKind::Fills
        | PropertyKind::SymbolFills
        | PropertyKind::SymbolOutlines => return Err(Error::InvalidSource),
    })
}

fn apply_one(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    property: &Property,
    budget: &mut ChartBudget,
) -> Result<Option<Published>, Error> {
    let target = target(source, sheet, chart, budget)?;
    if property.len() != target.series {
        return Err(Refusal::SeriesCount {
            expected: target.series,
            given: property.len(),
        }
        .into());
    }
    if let (Some(style_kind), Some(values)) = (property.kind().style(), property.style_values()) {
        admit_graph_workspace(source, budget)?;
        let current = from_style_values(
            property.kind(),
            style::read(source, target.chart, target.kind, target.series, style_kind)
                .map_err(from_style)?,
        )?;
        if current == *property {
            return Ok(None);
        }
        admit_style_candidate(budget)?;
        let values: Vec<Option<style::Value>> = values.into_iter().map(Some).collect();
        let published = style::write(source, target.chart, target.kind, style_kind, &values)
            .map_err(from_style)?;
        return publish_changes(published, sheet, chart, property.kind(), budget);
    }
    let (components, limits) = prepare_components(source, budget)?;
    let mut workspace = Workspace::new(&components, limits).map_err(Error::from)?;
    let chart_id = target.chart;
    let mut symbol_scope = None;
    let changed = match property {
        Property::ConnectionLines(values) => {
            engine::set_values(&mut workspace, chart_id, &lines(target.kind)?, values)?
        },
        Property::WedgeExplosions(values) => {
            engine::set_values(&mut workspace, chart_id, &WedgeExplosions, values)?
        },
        Property::LabelDistances(values) => {
            engine::set_values(&mut workspace, chart_id, &LabelDistances, values)?
        },
        Property::Symbols(values) => {
            let limits = budget.series_limits();
            let (changed, report) = engine::set_symbols(
                &mut workspace,
                chart_id,
                symbols(target.kind)?,
                values,
                limits,
            )
            .map_err(Error::from)?;
            symbol_scope = report.scope().cloned();
            budget
                .charge_series_report(report)
                .map_err(map_arrangement)?;
            changed
        },
        Property::Trendlines(values) => {
            engine::set_values(&mut workspace, chart_id, &Trendlines, values)?
        },
        Property::ErrorBars(values) => {
            engine::set_values(&mut workspace, chart_id, &ErrorBars, values)?
        },
        Property::PieLabels(values) => {
            engine::set_values(&mut workspace, chart_id, &PieLabels, values)?
        },
        Property::ValueLabels(values) => engine::set_values(
            &mut workspace,
            chart_id,
            &value_labels(target.kind)?,
            values,
        )?,
        Property::Strokes(_)
        | Property::Fills(_)
        | Property::SymbolFills(_)
        | Property::SymbolOutlines(_) => return Err(Error::InvalidSource),
    };
    let changes = workspace.finish();
    if !changed || changes.is_empty() {
        return Ok(None);
    }
    if let Some(scope) = symbol_scope.as_ref() {
        validate_symbol_scope(scope, target.chart, source, &changes)?;
    }
    let published = publish_changes_exact(source, &changes, budget)?;
    publish_changes(Some(published), sheet, chart, property.kind(), budget)
}

fn validate_symbol_scope(
    scope: &engine::MutationScope,
    chart: u64,
    source: &Package,
    changes: &litchi_iwa_style::graph::Changes,
) -> Result<(), Error> {
    if scope.selected_chart() != chart {
        return Err(Error::Verification);
    }
    let edited_outside_scope = changes.edits.iter().any(|(name, archive)| {
        if !scope
            .allowed_members()
            .iter()
            .any(|allowed| allowed == name)
        {
            return true;
        }
        let Some(original) = source
            .state
            .components
            .catalog()
            .iter()
            .find(|component| component.name() == name)
            .map(|component| component.archive())
        else {
            return true;
        };
        archive.objects.iter().any(|object| {
            let Some(identifier) = object.archive_info.identifier else {
                return true;
            };
            let unchanged = original
                .objects
                .iter()
                .find(|candidate| candidate.archive_info.identifier == Some(identifier))
                .is_some_and(|candidate| candidate == object);
            if unchanged {
                return false;
            }
            scope.allowed_objects().contains(&identifier)
                && object
                    .messages
                    .iter()
                    .all(|message| scope.allowed_message_types().contains(&message.type_))
        })
    });
    let inserted_outside_scope = changes.insertions.iter().any(|insertion| {
        !scope
            .planned_new_members()
            .iter()
            .any(|allowed| allowed == &insertion.name)
            || insertion.archive.objects.iter().any(|object| {
                object
                    .archive_info
                    .identifier
                    .is_none_or(|identifier| !scope.planned_new_objects().contains(&identifier))
            })
    });
    let deleted_outside_scope = changes.deletions.iter().any(|name| {
        !scope
            .planned_removed_members()
            .iter()
            .any(|removed| removed == name)
    });
    if edited_outside_scope || inserted_outside_scope || deleted_outside_scope {
        return Err(Error::Verification);
    }
    Ok(())
}

fn publish_changes(
    published: Option<Published>,
    sheet: usize,
    chart: ChartSelector,
    kind: PropertyKind,
    budget: &mut ChartBudget,
) -> Result<Option<Published>, Error> {
    let Some(published) = published else {
        return Ok(None);
    };
    let mut read_budget = *budget;
    let _ = read(&published.package, sheet, chart, kind, &mut read_budget)?;
    *budget = read_budget;
    Ok(Some(published))
}

fn admit_style_candidate(budget: &mut ChartBudget) -> Result<(), Error> {
    let bound = budget.series_limits().max_output_bytes();
    budget.preflight_output(bound).map_err(map_arrangement)?;
    budget.preflight_retained(bound).map_err(map_arrangement)?;
    budget.preflight_scratch(bound).map_err(map_arrangement)?;
    budget.preflight_allocations(1).map_err(map_arrangement)?;
    budget.output(bound).map_err(map_arrangement)?;
    budget.retained(bound).map_err(map_arrangement)?;
    budget.scratch(bound).map_err(map_arrangement)?;
    budget.allocations(1).map_err(map_arrangement)?;
    budget.candidate_reopen(bound).map_err(map_arrangement)?;
    Ok(())
}

/// Reset series `series`' own value of a series-style property.
fn reset_one(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    kind: PropertyKind,
    series: usize,
    budget: &mut ChartBudget,
) -> Result<Option<Published>, Error> {
    let style_kind = kind.style().ok_or(Refusal::NotResettable)?;
    let target = target(source, sheet, chart, budget)?;
    if series >= target.series {
        return Err(Refusal::SeriesNotFound { series }.into());
    }
    let current = style::read(source, target.chart, target.kind, target.series, style_kind)
        .map_err(from_style)?;
    let values: Vec<Option<style::Value>> = current
        .into_iter()
        .enumerate()
        .map(|(index, value)| (index != series).then_some(value))
        .collect();
    admit_style_candidate(budget)?;
    let published =
        style::write(source, target.chart, target.kind, style_kind, &values).map_err(from_style)?;
    publish_changes(published, sheet, chart, kind, budget)
}

fn from_style(failure: style::Failure) -> Error {
    match failure {
        style::Failure::Refused(style::Refused::Kind) => Refusal::UnsupportedKind.into(),
        style::Failure::Refused(style::Refused::Shared) => Refusal::SharedStyle.into(),
        style::Failure::Refused(style::Refused::ImageFill) => Refusal::ImageFill.into(),
        style::Failure::Fault(fault) => fault.into(),
    }
}

fn from_style_values(kind: PropertyKind, values: Vec<style::Value>) -> Result<Property, Error> {
    let strokes = |values: Vec<style::Value>| -> Result<Vec<Option<Stroke>>, Error> {
        values
            .into_iter()
            .map(|value| match value {
                style::Value::Stroke(stroke) => Ok(stroke),
                _ => Err(Error::InvalidSource),
            })
            .collect()
    };
    Ok(match kind {
        PropertyKind::Strokes => Property::Strokes(strokes(values)?),
        PropertyKind::SymbolOutlines => Property::SymbolOutlines(strokes(values)?),
        PropertyKind::Fills => Property::Fills(
            values
                .into_iter()
                .map(|value| match value {
                    style::Value::Fill(fill) => Ok(fill),
                    _ => Err(Error::InvalidSource),
                })
                .collect::<Result<_, _>>()?,
        ),
        PropertyKind::SymbolFills => Property::SymbolFills(
            values
                .into_iter()
                .map(|value| match value {
                    style::Value::SymbolFill(fill) => Ok(fill),
                    _ => Err(Error::InvalidSource),
                })
                .collect::<Result<_, _>>()?,
        ),
        _ => return Err(Error::InvalidSource),
    })
}

fn resolve_sheet(
    package: &Package,
    selector: SheetSelector<'_>,
    budget: &mut ChartBudget,
) -> Result<usize, Error> {
    arrangement::resolve_sheet_position(package, selector, budget)
        .map(|position| position.get())
        .map_err(map_arrangement)
}

impl Package {
    /// The per-series values of one chart property of one sheet chart, in
    /// series order.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedKind`] and source errors.
    pub fn sheet_chart_series<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement)?;
        let sheet = resolve_sheet(self, sheet.into(), &mut budget)?;
        read(self, sheet, chart.into(), kind, &mut budget)
    }

    /// Start an immutable per-series property edit of one sheet chart.
    ///
    /// # Errors
    ///
    /// Selector errors and source errors.
    pub fn edit_sheet_chart_series<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = ChartBudget::for_package(self).map_err(map_arrangement)?;
        let sheet = resolve_sheet(self, sheet.into(), &mut budget)?;
        let chart = chart.into();
        target(self, sheet, chart, &mut budget)?;
        Ok(Edit {
            source: self,
            sheet,
            chart,
            changes: Vec::new(),
            budget,
        })
    }

    /// Apply a chart-series patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_chart_series(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
