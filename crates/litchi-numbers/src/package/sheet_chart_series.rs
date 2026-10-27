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

use litchi_iwa_common::chart::error_bar::Series as ErrorBarSeries;
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::pie::{LabelDistance, LabelVisibility, WedgeExplosion};
use litchi_iwa_common::chart::series::{
    ConnectionLine, ConnectionLineFamily, Symbol, SymbolFamily, Trendline,
};
use litchi_iwa_common::chart::series_labels::Visibility;
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
use super::sheet_drawable_graph::{self as graph, CHART_MESSAGE_TYPE};
use super::structure::{self, Published, references};
use super::style_graph::{self, LeafComponents};
use crate::selector::{ChartSelector, SheetSelector};
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("chart-series");

/// `TSCH.ChartDrawableArchive` extension holding the `TSCH.ChartArchive`.
const CHART_ARCHIVE_FIELD: u32 = 10_000;
/// `TSCH.ChartArchive.series_direction`.
const SERIES_DIRECTION_FIELD: u32 = 5;
const SERIES_BY_ROW: u64 = 1;
const SERIES_BY_COLUMN: u64 = 2;

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
        }
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
            Self::SheetNotFound | Self::ChartNotFound => RefusalKind::NotFound,
            Self::UnsupportedKind => RefusalKind::UnsupportedTarget,
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
        style_graph::fault(&error).into()
    }
}

/// Chart-series changes staged against one chart of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    chart: ChartSelector,
    changes: Vec<Property>,
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
        self.changes.push(property);
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
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for property in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) = apply_one(base, self.sheet, self.chart, property)? {
                if read(&published.package, self.sheet, self.chart, property.kind())? != *property {
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

fn target(package: &Package, sheet: usize, chart: ChartSelector) -> Result<Target, Error> {
    let drawables = graph::sheet_drawables(package, sheet)?;
    let entry = drawables
        .select(DrawableSelector::from(chart))
        .ok_or(Refusal::ChartNotFound)?;
    let (index, message_type) = entry.message.ok_or(Error::InvalidSource)?;
    if message_type != CHART_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
    }
    let metadata = package
        .sheet_chart_metadata(sheet, chart)
        .map_err(|_error| Error::InvalidSource)?;
    let object = graph::object(package, entry.identifier)?;
    let archives = references::records(&object.messages[index].data, &[CHART_ARCHIVE_FIELD])?;
    let [archive] = archives.as_slice() else {
        return Err(Error::InvalidSource);
    };
    let view =
        litchi_iwa_common::wire::WireView::parse(archive, litchi_iwa_common::WireLimits::DEFAULT)?;
    let direction = graph::unique_field(&view, SERIES_DIRECTION_FIELD, 0)?
        .map(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()))
        .transpose()
        .map_err(|_error| Error::InvalidSource)?
        .map(|(value, _)| value);
    let components = match direction {
        Some(SERIES_BY_ROW) => metadata.row_names().len(),
        Some(SERIES_BY_COLUMN) => metadata.column_names().len(),
        _ => return Err(Error::InvalidSource),
    };
    let series = series_count(metadata.kind(), components)?;
    Ok(Target {
        chart: entry.identifier,
        kind: metadata.kind(),
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
) -> Result<Property, Error> {
    let target = target(package, sheet, chart)?;
    let components = LeafComponents::new(package)?;
    let limits = structure::exact_source(package)?
        .limits()
        .effective_archive_limits()
        .map_err(|error| structure::map_archive(&error))?;
    let workspace = Workspace::new(&components, limits)?;
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
        PropertyKind::Symbols => Property::Symbols(engine::symbols(
            &workspace,
            chart,
            symbols(target.kind)?,
            count,
        )?),
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
    })
}

fn apply_one(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    property: &Property,
) -> Result<Option<Published>, Error> {
    let target = target(source, sheet, chart)?;
    if property.len() != target.series {
        return Err(Refusal::SeriesCount {
            expected: target.series,
            given: property.len(),
        }
        .into());
    }
    let components = LeafComponents::new(source)?;
    let limits = structure::exact_source(source)?
        .limits()
        .effective_archive_limits()
        .map_err(|error| structure::map_archive(&error))?;
    let mut workspace = Workspace::new(&components, limits)?;
    let chart = target.chart;
    let changed = match property {
        Property::ConnectionLines(values) => {
            engine::set_values(&mut workspace, chart, &lines(target.kind)?, values)?
        },
        Property::WedgeExplosions(values) => {
            engine::set_values(&mut workspace, chart, &WedgeExplosions, values)?
        },
        Property::LabelDistances(values) => {
            engine::set_values(&mut workspace, chart, &LabelDistances, values)?
        },
        Property::Symbols(values) => {
            engine::set_symbols(&mut workspace, chart, symbols(target.kind)?, values)?
        },
        Property::Trendlines(values) => {
            engine::set_values(&mut workspace, chart, &Trendlines, values)?
        },
        Property::ErrorBars(values) => {
            engine::set_values(&mut workspace, chart, &ErrorBars, values)?
        },
        Property::PieLabels(values) => {
            engine::set_values(&mut workspace, chart, &PieLabels, values)?
        },
        Property::ValueLabels(values) => {
            engine::set_values(&mut workspace, chart, &value_labels(target.kind)?, values)?
        },
    };
    let changes = workspace.finish();
    if !changed || changes.is_empty() {
        return Ok(None);
    }
    Ok(Some(style_graph::publish(source, changes)?))
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
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
        let sheet = resolve_sheet(self, sheet.into())?;
        read(self, sheet, chart.into(), kind)
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
        let sheet = resolve_sheet(self, sheet.into())?;
        let chart = chart.into();
        target(self, sheet, chart)?;
        Ok(Edit {
            source: self,
            sheet,
            chart,
            changes: Vec::new(),
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
