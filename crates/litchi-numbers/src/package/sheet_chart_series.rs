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

use litchi_iwa_archive::package::{EntryEdit, EntryInsertion, OwnedExactArtifacts};
use litchi_iwa_common::chart::error_bar::Series as ErrorBarSeries;
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::pie::{LabelDistance, LabelVisibility, WedgeExplosion};
use litchi_iwa_common::chart::series::{
    ConnectionLine, ConnectionLineFamily, Stroke, Symbol, SymbolFamily, SymbolFill, Trendline,
};
use litchi_iwa_common::chart::series_labels::Visibility;
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
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

/// Resource category reported by the Numbers chart-series transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// One decoded IWA component payload.
    PayloadBytes,
    /// Native object identifiers.
    Identifiers,
    /// Temporary graph scratch bytes.
    ScratchBytes,
    /// Fallible graph allocation events.
    AllocationEvents,
    /// A bounded shared series-engine axis.
    Engine(litchi_iwa_style::series::LimitKind),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::PackageBytes => formatter.write_str("package bytes"),
            Self::PayloadBytes => formatter.write_str("payload bytes"),
            Self::Identifiers => formatter.write_str("object identifiers"),
            Self::ScratchBytes => formatter.write_str("scratch bytes"),
            Self::AllocationEvents => formatter.write_str("allocation events"),
            Self::Engine(kind) => kind.fmt(formatter),
        }
    }
}

/// A typed chart-series transaction failure.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    #[error(transparent)]
    Refused(Refusal),
    #[error("the rooted Numbers chart-series source is invalid")]
    InvalidSource,
    #[error("Numbers chart-series {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} Numbers chart-series units")]
    Allocation { amount: usize },
    #[error("Numbers chart-series bounded finalization arithmetic overflowed")]
    Overflow,
    #[error("the edited Numbers chart series failed semantic verification")]
    Verification,
    #[error("the Numbers chart-series patch does not match the exact source package")]
    PatchConflict,
}

impl From<structure::Fault> for Error {
    fn from(fault: structure::Fault) -> Self {
        match fault {
            structure::Fault::Unsupported => Refusal::UnsupportedSource.into(),
            structure::Fault::Invalid => Self::InvalidSource,
            structure::Fault::Dependency => Refusal::UnsupportedDependency.into(),
            structure::Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    structure::FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    structure::FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    structure::FaultLimit::Identifiers => LimitKind::Identifiers,
                },
                observed,
                maximum,
            },
            structure::Fault::Allocation { amount } => Self::Allocation { amount },
            structure::Fault::Verification => Self::Verification,
        }
    }
}

const PROPERTY_KINDS: usize = 12;

const fn property_index(kind: PropertyKind) -> usize {
    match kind {
        PropertyKind::ConnectionLines => 0,
        PropertyKind::WedgeExplosions => 1,
        PropertyKind::LabelDistances => 2,
        PropertyKind::Symbols => 3,
        PropertyKind::Trendlines => 4,
        PropertyKind::ErrorBars => 5,
        PropertyKind::PieLabels => 6,
        PropertyKind::ValueLabels => 7,
        PropertyKind::Strokes => 8,
        PropertyKind::Fills => 9,
        PropertyKind::SymbolFills => 10,
        PropertyKind::SymbolOutlines => 11,
    }
}

/// An exact-source reversible chart-series patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    sheet: usize,
    chart: ChartSelector,
    chart_identifier: u64,
    before: [Option<Property>; PROPERTY_KINDS],
    after: [Option<Property>; PROPERTY_KINDS],
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("sheet", &self.sheet)
            .field("chart", &self.chart)
            .finish_non_exhaustive()
    }
}

impl Patch {
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop() && self.before == self.after
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet: self.sheet,
            chart: self.chart,
            chart_identifier: self.chart_identifier,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Compact chart-series publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
}

impl Diagnostics {
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
}

/// A verified immutable chart-series publication.
#[must_use = "a chart-series commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    fn unchanged(
        source: &Package,
        sheet: usize,
        chart: ChartSelector,
        chart_identifier: u64,
        before: [Option<Property>; PROPERTY_KINDS],
        after: [Option<Property>; PROPERTY_KINDS],
    ) -> Result<Self, Error> {
        let owner = structure::source_owner(source)?;
        Ok(Self {
            package: source.clone(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
                sheet,
                chart,
                chart_identifier,
                before,
                after,
            },
            diagnostics: Diagnostics {
                changed: false,
                touched_components: 0,
            },
        })
    }

    fn published(
        source: &Package,
        published: Published,
        sheet: usize,
        chart: ChartSelector,
        chart_identifier: u64,
        before: [Option<Property>; PROPERTY_KINDS],
        after: [Option<Property>; PROPERTY_KINDS],
    ) -> Result<Self, Error> {
        let source_owner = structure::source_owner(source)?;
        let target_owner = structure::source_owner(&published.package)?;
        Ok(Self {
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
                sheet,
                chart,
                chart_identifier,
                before,
                after,
            },
            diagnostics: Diagnostics {
                changed: true,
                touched_components: published.touched_components,
            },
            package: published.package,
        })
    }
}

/// `TSCH.ChartDrawableArchive` extension holding the `TSCH.ChartArchive`.
const CHART_ARCHIVE_FIELD: u32 = 10_000;
/// `TSCH.ChartArchive.series_direction`.
const SERIES_DIRECTION_FIELD: u32 = 5;
const SERIES_BY_ROW: u64 = 1;
const SERIES_BY_COLUMN: u64 = 2;
const MAX_STAGED_OPERATIONS: usize = 4;
const SYMBOL_GENERATED_FIELDS: std::ops::RangeInclusive<u32> = 10_000..=10_003;

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

impl From<litchi_iwa_style::graph::FinishError> for Error {
    fn from(error: litchi_iwa_style::graph::FinishError) -> Self {
        match error {
            litchi_iwa_style::graph::FinishError::Graph(error) => error.into(),
            litchi_iwa_style::graph::FinishError::Overflow(_) => Error::Overflow,
            litchi_iwa_style::graph::FinishError::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: match kind {
                    litchi_iwa_style::graph::FinishLimitKind::Work => {
                        LimitKind::Engine(engine::LimitKind::Work)
                    },
                    litchi_iwa_style::graph::FinishLimitKind::Allocations => {
                        LimitKind::AllocationEvents
                    },
                    litchi_iwa_style::graph::FinishLimitKind::RetainedBytes => {
                        LimitKind::Engine(engine::LimitKind::RetainedBytes)
                    },
                    litchi_iwa_style::graph::FinishLimitKind::ScratchBytes => {
                        LimitKind::ScratchBytes
                    },
                    _ => return Error::InvalidSource,
                },
                observed: observed as u64,
                maximum: maximum as u64,
            },
            litchi_iwa_style::graph::FinishError::Allocation { amount } => {
                Error::Allocation { amount }
            },
            _ => Error::InvalidSource,
        }
    }
}

impl From<litchi_iwa_style::graph::ConstructionError> for Error {
    fn from(error: litchi_iwa_style::graph::ConstructionError) -> Self {
        match error {
            litchi_iwa_style::graph::ConstructionError::Graph(error) => error.into(),
            litchi_iwa_style::graph::ConstructionError::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: match kind {
                    litchi_iwa_style::graph::ConstructionLimitKind::Work => {
                        LimitKind::Engine(engine::LimitKind::Work)
                    },
                    litchi_iwa_style::graph::ConstructionLimitKind::Allocations => {
                        LimitKind::AllocationEvents
                    },
                    litchi_iwa_style::graph::ConstructionLimitKind::RetainedBytes => {
                        LimitKind::Engine(engine::LimitKind::RetainedBytes)
                    },
                    litchi_iwa_style::graph::ConstructionLimitKind::ScratchBytes => {
                        LimitKind::ScratchBytes
                    },
                    _ => return Error::InvalidSource,
                },
                observed: observed as u64,
                maximum: maximum as u64,
            },
            litchi_iwa_style::graph::ConstructionError::Allocation { amount } => {
                Error::Allocation { amount }
            },
            _ => Error::InvalidSource,
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
    chart_identifier: u64,
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
        if self.changes.len() >= MAX_STAGED_OPERATIONS {
            return Err(Error::Limit {
                kind: LimitKind::Engine(engine::LimitKind::Allocations),
                observed: self.changes.len().saturating_add(1) as u64,
                maximum: MAX_STAGED_OPERATIONS as u64,
            });
        }
        let residual = self.budget.series_limits().map_err(map_arrangement)?;
        if property.len() > residual.max_references() {
            return Err(Error::Limit {
                kind: LimitKind::Engine(engine::LimitKind::References),
                observed: property.len() as u64,
                maximum: residual.max_references() as u64,
            });
        }
        self.budget
            .preflight_allocations(1)
            .map_err(map_arrangement)?;
        structure::reserve(&mut self.changes, 1)?;
        self.budget.allocations(1).map_err(map_arrangement)?;
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
        if self.changes.len() >= MAX_STAGED_OPERATIONS {
            return Err(Error::Limit {
                kind: LimitKind::Engine(engine::LimitKind::Allocations),
                observed: self.changes.len().saturating_add(1) as u64,
                maximum: MAX_STAGED_OPERATIONS as u64,
            });
        }
        self.budget
            .preflight_allocations(1)
            .map_err(map_arrangement)?;
        structure::reserve(&mut self.changes, 1)?;
        self.budget.allocations(1).map_err(map_arrangement)?;
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
        let mut before: [Option<Property>; PROPERTY_KINDS] = std::array::from_fn(|_| None);
        let mut after: [Option<Property>; PROPERTY_KINDS] = std::array::from_fn(|_| None);
        for change in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let kind = match change {
                Staged::Set(property) => property.kind(),
                Staged::Reset(kind, _) => *kind,
            };
            let index = property_index(kind);
            if before[index].is_none() {
                before[index] = Some(read(base, self.sheet, self.chart, kind, &mut budget)?);
            }
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
                after[index] = Some(read(
                    current.as_ref().expect("published package is retained"),
                    self.sheet,
                    self.chart,
                    kind,
                    &mut budget,
                )?);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(
                self.source,
                self.sheet,
                self.chart,
                self.chart_identifier,
                before,
                after,
            );
        };
        Commit::published(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
            self.sheet,
            self.chart,
            self.chart_identifier,
            before,
            after,
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

struct PreparedComponents<'a> {
    components: Vec<LeafComponents<'a>>,
    limits: litchi_iwa_core::Limits,
    requirements: litchi_iwa_style::graph::ConstructionRequirements,
    construction_limits: litchi_iwa_style::graph::ConstructionLimits,
}

impl PreparedComponents<'_> {
    fn workspace(&self) -> Result<Workspace<'_>, Error> {
        Workspace::new_bounded(
            &self.components[0],
            self.limits,
            self.requirements,
            self.construction_limits,
        )
        .map_err(Error::from)
    }
}

fn prepare_components<'a>(
    package: &'a Package,
    budget: &mut ChartBudget,
) -> Result<PreparedComponents<'a>, Error> {
    let (view_allocations, view_retained) =
        LeafComponents::view_requirements(package).map_err(Error::from)?;
    budget
        .preflight_allocations(view_allocations)
        .map_err(map_arrangement)?;
    budget
        .preflight_retained(view_retained)
        .map_err(map_arrangement)?;
    budget
        .allocations(view_allocations)
        .map_err(map_arrangement)?;
    budget.retained(view_retained).map_err(map_arrangement)?;
    let mut components = Vec::new();
    components
        .try_reserve_exact(1)
        .map_err(|_| Error::Allocation { amount: 1 })?;
    components.push(LeafComponents::new(package).map_err(Error::from)?);
    let construction_limits = budget.construction_limits().map_err(map_arrangement)?;
    let requirements = Workspace::construction_requirements(&components[0], construction_limits)
        .map_err(Error::from)?;
    budget
        .charge_construction(requirements)
        .map_err(map_arrangement)?;
    let limits = structure::exact_source(package)?
        .limits()
        .effective_archive_limits()
        .map_err(|error| structure::map_archive(&error))?;
    Ok(PreparedComponents {
        components,
        limits,
        requirements,
        construction_limits,
    })
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
    let mut retained_bound = 0usize;
    let mut scratch_bound = 0usize;
    let mut work_bound = 0usize;
    let mut member_count = 0usize;
    let mut bounds = |_name: &str, archive: &Archive| -> Result<(), Error> {
        let encoded = archive
            .encoded_len(archive_limits)
            .map_err(|error| structure::map_core(&error))?;
        let compressed = SnappyStream::maximum_compressed_len(encoded, snappy_limits)
            .map_err(|error| structure::map_core(&error))?;
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
        .checked_mul(2)
        .and_then(|value| value.checked_add(changes.edits.len()))
        .and_then(|value| value.checked_add(changes.insertions.len()))
        .and_then(|value| value.checked_add(changes.deletions.len()))
        .and_then(|value| value.checked_add(3))
        .ok_or(Error::InvalidSource)?;
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
        .map_err(|error| structure::map_archive(&error))?;
    let requirements = prepared.execution_requirements();
    budget.reassembly(requirements).map_err(map_arrangement)?;
    budget
        .candidate_reopen(requirements.output_bytes())
        .map_err(map_arrangement)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(|error| structure::map_archive(&error))?;
    let candidate = Package::from_owned_bytes_with_options(output, package.state.options)
        .map_err(|_| Error::Verification)?;
    verify_member_locality(package, &candidate, changes, budget)?;
    Ok(Published {
        package: candidate,
        touched_components: edits.len() + insertions.len() + deletions.len(),
        deleted_previews: 0,
    })
}

fn verify_member_locality(
    source: &Package,
    candidate: &Package,
    changes: &litchi_iwa_style::graph::Changes,
    budget: &mut ChartBudget,
) -> Result<(), Error> {
    let source_catalog = structure::exact_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = structure::exact_source(candidate).map_err(|_| Error::Verification)?;
    let mut work = 0usize;
    for entry in source_catalog
        .package()
        .iter()
        .chain(candidate_catalog.package().iter())
    {
        work = work
            .checked_add(arrangement::zip_entry_work(entry, true).map_err(map_arrangement)?)
            .ok_or(Error::InvalidSource)?;
    }
    budget.preflight_work(work).map_err(map_arrangement)?;
    budget.work(work).map_err(map_arrangement)?;

    let source_changed = |name: &str| changes.edits.iter().any(|(edited, _)| edited == name);
    let source_deleted = |name: &str| changes.deletions.iter().any(|deleted| deleted == name);
    let source_inserted = |name: &str| {
        changes
            .insertions
            .iter()
            .any(|insertion| insertion.name == name)
    };
    let source_expected_count = source_catalog
        .package()
        .iter()
        .filter(|entry| !source_deleted(entry.name()))
        .count()
        .checked_add(changes.insertions.len())
        .ok_or(Error::InvalidSource)?;
    if candidate_catalog.package().iter().count() != source_expected_count {
        return Err(Error::Verification);
    }
    for entry in source_catalog.package().iter() {
        let Some(target) = candidate_catalog
            .package()
            .iter()
            .find(|candidate| candidate.name() == entry.name())
        else {
            if source_deleted(entry.name()) {
                continue;
            }
            return Err(Error::Verification);
        };
        if source_deleted(entry.name()) {
            return Err(Error::Verification);
        }
        if source_changed(entry.name()) {
            if !arrangement::same_entry_fixed_metadata(entry, target)
                || !arrangement::same_changed_entry_records(entry, target)
            {
                return Err(Error::Verification);
            }
        } else if !arrangement::same_unselected_entry_records(entry, target) {
            return Err(Error::Verification);
        }
    }
    for entry in candidate_catalog.package().iter() {
        if source_catalog
            .package()
            .iter()
            .any(|source| source.name() == entry.name())
        {
            continue;
        }
        if !source_inserted(entry.name()) {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn apply_series_patch(package: &Package, patch: &Patch) -> Result<Commit, Error> {
    let owner = structure::source_owner(package)?;
    if !patch.artifacts.authorizes_owner(&owner) {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Commit {
            package: package.clone(),
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: false,
                touched_components: 0,
            },
        });
    }
    let mut budget = ChartBudget::for_package(package).map_err(map_arrangement)?;
    let target_bytes = patch.artifacts.target_owner().as_ref().len();
    budget
        .preflight_output(target_bytes)
        .map_err(map_arrangement)?;
    budget
        .preflight_retained(target_bytes)
        .map_err(map_arrangement)?;
    budget.preflight_allocations(1).map_err(map_arrangement)?;
    budget
        .preflight_scratch(target_bytes)
        .map_err(map_arrangement)?;
    budget
        .candidate_reopen(target_bytes)
        .map_err(map_arrangement)?;
    let candidate = Package::from_source_owner_with_options(
        patch.artifacts.target_owner(),
        package.state.options,
    )
    .map_err(|_| Error::Verification)?;
    let selected = target(&candidate, patch.sheet, patch.chart, &mut budget)?;
    if selected.chart != patch.chart_identifier {
        return Err(Error::PatchConflict);
    }
    for (index, expected) in patch.after.iter().enumerate() {
        let Some(expected) = expected else {
            continue;
        };
        let kind = match index {
            0 => PropertyKind::ConnectionLines,
            1 => PropertyKind::WedgeExplosions,
            2 => PropertyKind::LabelDistances,
            3 => PropertyKind::Symbols,
            4 => PropertyKind::Trendlines,
            5 => PropertyKind::ErrorBars,
            6 => PropertyKind::PieLabels,
            7 => PropertyKind::ValueLabels,
            8 => PropertyKind::Strokes,
            9 => PropertyKind::Fills,
            10 => PropertyKind::SymbolFills,
            11 => PropertyKind::SymbolOutlines,
            _ => return Err(Error::Verification),
        };
        if read(&candidate, patch.sheet, patch.chart, kind, &mut budget)? != *expected {
            return Err(Error::Verification);
        }
    }
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics {
            changed: true,
            touched_components: 0,
        },
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
        let _prepared = prepare_components(package, budget)?;
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
    let prepared = prepare_components(package, budget)?;
    let workspace = prepared.workspace()?;
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
            let limits = budget.series_limits().map_err(map_arrangement)?;
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
        let _prepared = prepare_components(source, budget)?;
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
    let prepared = prepare_components(source, budget)?;
    let mut workspace = prepared.workspace()?;
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
            let limits = budget.series_limits().map_err(map_arrangement)?;
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
    let finish_limits = budget.finish_limits().map_err(map_arrangement)?;
    workspace.finish_requirements().map_err(Error::from)?;
    let (changes, finish_report) = workspace
        .finish_bounded(finish_limits)
        .map_err(Error::from)?;
    budget
        .charge_finish_report(finish_report)
        .map_err(map_arrangement)?;
    if !changed || changes.is_empty() {
        return Ok(None);
    }
    if let Some(scope) = symbol_scope.as_ref() {
        validate_symbol_scope(scope, target.chart, source, &changes, budget)?;
    }
    let published = publish_changes_exact(source, &changes, budget)?;
    publish_changes(Some(published), sheet, chart, property.kind(), budget)
}

fn validate_symbol_scope(
    scope: &engine::MutationScope,
    chart: u64,
    source: &Package,
    changes: &litchi_iwa_style::graph::Changes,
    budget: &mut ChartBudget,
) -> Result<(), Error> {
    if scope.selected_chart() != chart {
        return Err(Error::Verification);
    }
    for (name, archive) in &changes.edits {
        if !scope
            .allowed_members()
            .iter()
            .any(|allowed| allowed == name)
        {
            return Err(Error::Verification);
        }
        let Some(original) = source
            .state
            .components
            .catalog()
            .iter()
            .find(|component| component.name() == name)
            .map(|component| component.archive())
        else {
            return Err(Error::Verification);
        };
        for object in &archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                return Err(Error::Verification);
            };
            let original_object = original
                .objects
                .iter()
                .find(|candidate| candidate.archive_info.identifier == Some(identifier));
            if original_object
                .as_ref()
                .is_some_and(|candidate| *candidate == object)
            {
                continue;
            }
            let object_allowed = scope.allowed_objects().contains(&identifier)
                || scope.planned_new_objects().contains(&identifier)
                || object.messages.iter().all(|message| {
                    message.type_ == 11_006
                        && scope.allowed_members().iter().any(|member| {
                            member == name
                                && (member.ends_with("Metadata.iwa")
                                    || member.ends_with("DocumentMetadata.iwa"))
                        })
                });
            let messages_allowed = object
                .messages
                .iter()
                .all(|message| scope.allowed_message_types().contains(&message.type_));
            if !object_allowed || !messages_allowed {
                return Err(Error::Verification);
            }
            if let Some(original_object) = original_object.as_ref()
                && (original_object.messages.len() != object.messages.len()
                    || object.messages.iter().any(|message| {
                        !original_object
                            .messages
                            .iter()
                            .any(|source| source.type_ == message.type_)
                    }))
            {
                return Err(Error::Verification);
            }
            if let Some(original_object) = original_object.as_ref() {
                validate_symbol_object_metadata(original_object, object, scope, budget)?;
            }
            if original_object.is_none() && !validate_new_symbol_object_metadata(object, scope) {
                return Err(Error::Verification);
            }
            for message in &object.messages {
                let original_message = original_object.as_ref().and_then(|candidate| {
                    candidate
                        .messages
                        .iter()
                        .find(|source| source.type_ == message.type_)
                });
                validate_symbol_message(original_message, message, scope.allowed_fields(), budget)?;
            }
        }
        if original.objects.iter().any(|original_object| {
            let Some(identifier) = original_object.archive_info.identifier else {
                return true;
            };
            archive
                .objects
                .iter()
                .all(|candidate| candidate.archive_info.identifier != Some(identifier))
                && !scope.planned_removed_objects().contains(&identifier)
        }) {
            return Err(Error::Verification);
        }
    }
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
    if inserted_outside_scope || deleted_outside_scope {
        return Err(Error::Verification);
    }
    Ok(())
}

fn scope_allows_candidate_identifier(scope: &engine::MutationScope, identifier: u64) -> bool {
    scope.allowed_objects().contains(&identifier)
        || scope.planned_new_objects().contains(&identifier)
}

fn scoped_reference_transition(
    source: &[u64],
    candidate: &[u64],
    scope: &engine::MutationScope,
) -> bool {
    if !candidate.iter().all(|identifier| {
        source.contains(identifier) || scope_allows_candidate_identifier(scope, *identifier)
    }) || !source.iter().all(|identifier| {
        candidate.contains(identifier) || scope.planned_removed_objects().contains(identifier)
    }) {
        return false;
    }
    let source_kept = source
        .iter()
        .filter(|identifier| candidate.contains(identifier));
    let candidate_existing = candidate
        .iter()
        .filter(|identifier| source.contains(identifier));
    source_kept.eq(candidate_existing)
}

#[derive(Clone, Copy)]
enum HeaderLevel {
    ArchiveInfo,
    MessageInfo,
    FieldInfo,
    FieldPath,
}

fn header_known(level: HeaderLevel, number: u32) -> bool {
    match level {
        HeaderLevel::ArchiveInfo => matches!(number, 1..=3),
        HeaderLevel::MessageInfo => matches!(number, 1..=6 | 8..=11),
        HeaderLevel::FieldInfo => matches!(number, 1..=8),
        HeaderLevel::FieldPath => number == 1,
    }
}

fn header_nested(level: HeaderLevel, number: u32) -> Option<HeaderLevel> {
    match (level, number) {
        (HeaderLevel::ArchiveInfo, 2) => Some(HeaderLevel::MessageInfo),
        (HeaderLevel::MessageInfo, 4) => Some(HeaderLevel::FieldInfo),
        (HeaderLevel::MessageInfo, 9 | 10) => Some(HeaderLevel::FieldPath),
        (HeaderLevel::FieldInfo, 1) => Some(HeaderLevel::FieldPath),
        _ => None,
    }
}

fn compare_header_unknowns(
    source_bytes: &[u8],
    candidate_bytes: &[u8],
    level: HeaderLevel,
    depth: usize,
    budget: &mut ChartBudget,
) -> Result<(), Error> {
    let source_view =
        parse_wire_view_with_budget(source_bytes, depth, budget).map_err(map_arrangement)?;
    let candidate_view =
        parse_wire_view_with_budget(candidate_bytes, depth, budget).map_err(map_arrangement)?;
    let mut source_unknown = source_view
        .fields()
        .filter(|field| !header_known(level, field.number()));
    let mut candidate_unknown = candidate_view
        .fields()
        .filter(|field| !header_known(level, field.number()));
    loop {
        match (source_unknown.next(), candidate_unknown.next()) {
            (Some(source), Some(candidate)) if source.raw() == candidate.raw() => {},
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }
    for number in 1..=11 {
        if !header_known(level, number) {
            continue;
        }
        let mut source_fields = source_view
            .fields()
            .filter(|field| field.number() == number);
        let mut candidate_fields = candidate_view
            .fields()
            .filter(|field| field.number() == number);
        loop {
            match (source_fields.next(), candidate_fields.next()) {
                (Some(source), Some(candidate)) => {
                    if let Some(nested) = header_nested(level, number) {
                        let source_payload = source.payload();
                        let candidate_payload = candidate.payload();
                        compare_header_unknowns(
                            source_payload,
                            candidate_payload,
                            nested,
                            depth.saturating_add(1),
                            budget,
                        )?;
                    }
                },
                (None, None) => break,
                _ => return Err(Error::Verification),
            }
        }
    }
    Ok(())
}

fn validate_symbol_object_metadata(
    source: &litchi_iwa_core::ArchiveObject,
    candidate: &litchi_iwa_core::ArchiveObject,
    scope: &engine::MutationScope,
    budget: &mut ChartBudget,
) -> Result<(), Error> {
    if source.archive_info.should_merge != candidate.archive_info.should_merge
        || source.source_headers().is_some() != candidate.source_headers().is_some()
    {
        return Err(Error::Verification);
    }
    if let (Some((source_header, _)), Some((candidate_header, _))) =
        (source.source_headers(), candidate.source_headers())
    {
        compare_header_unknowns(
            source_header,
            candidate_header,
            HeaderLevel::ArchiveInfo,
            1,
            budget,
        )?;
    }
    if source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len() {
        return Err(Error::Verification);
    }
    for (index, (source_info, candidate_info)) in source
        .archive_info
        .message_infos
        .iter()
        .zip(&candidate.archive_info.message_infos)
        .enumerate()
    {
        let Some(source_message) = source.messages.get(index) else {
            return Err(Error::Verification);
        };
        let Some(candidate_message) = candidate.messages.get(index) else {
            return Err(Error::Verification);
        };
        let source_length = u32::try_from(source_message.data.len()).ok();
        let candidate_length = u32::try_from(candidate_message.data.len()).ok();
        if source_info.type_ != candidate_info.type_
            || source_info.length != source_length.unwrap_or(u32::MAX)
            || candidate_info.length != candidate_length.unwrap_or(u32::MAX)
            || source_info.versions != candidate_info.versions
            || source_info.base_message_index != candidate_info.base_message_index
            || source_info.diff_merge_version != candidate_info.diff_merge_version
            || source_info.diff_field_path != candidate_info.diff_field_path
            || source_info.fields_to_remove != candidate_info.fields_to_remove
            || source_info.diff_read_version != candidate_info.diff_read_version
            || !scoped_reference_transition(
                &source_info.object_references,
                &candidate_info.object_references,
                scope,
            )
            || !scoped_reference_transition(
                &source_info.data_references,
                &candidate_info.data_references,
                scope,
            )
            || source_info.field_infos.len() != candidate_info.field_infos.len()
        {
            return Err(Error::Verification);
        }
        for (source_field, candidate_field) in source_info
            .field_infos
            .iter()
            .zip(&candidate_info.field_infos)
        {
            if source_field.path != candidate_field.path
                || source_field.r#type != candidate_field.r#type
                || source_field.unknown_field_rule != candidate_field.unknown_field_rule
                || source_field.known_field_rule != candidate_field.known_field_rule
                || source_field.known_field_version != candidate_field.known_field_version
                || source_field.known_field_feature_identifier
                    != candidate_field.known_field_feature_identifier
                || !scoped_reference_transition(
                    &source_field.object_references,
                    &candidate_field.object_references,
                    scope,
                )
                || !scoped_reference_transition(
                    &source_field.data_references,
                    &candidate_field.data_references,
                    scope,
                )
            {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

fn validate_new_symbol_object_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    scope: &engine::MutationScope,
) -> bool {
    if object.archive_info.should_merge.is_some()
        || object.archive_info.message_infos.len() != object.messages.len()
    {
        return false;
    }
    object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
        .all(|(message, info)| {
            u32::try_from(message.data.len()).ok() == Some(info.length)
                && message.type_ == info.type_
                && !info.versions.is_empty()
                && info.base_message_index.is_none()
                && info.diff_merge_version.is_empty()
                && info.diff_field_path.is_none()
                && info.fields_to_remove.is_empty()
                && info.diff_read_version.is_empty()
                && info.field_infos.is_empty()
                && info
                    .object_references
                    .iter()
                    .all(|identifier| scope_allows_candidate_identifier(scope, *identifier))
                && info
                    .data_references
                    .iter()
                    .all(|identifier| scope_allows_candidate_identifier(scope, *identifier))
                && info.field_infos.iter().all(|field| {
                    field
                        .object_references
                        .iter()
                        .all(|identifier| scope_allows_candidate_identifier(scope, *identifier))
                        && field
                            .data_references
                            .iter()
                            .all(|identifier| scope_allows_candidate_identifier(scope, *identifier))
                })
        })
}

fn validate_symbol_message(
    original: Option<&RawMessage>,
    candidate: &RawMessage,
    allowed_fields: &[u32],
    budget: &mut ChartBudget,
) -> Result<(), Error> {
    let candidate_view =
        parse_wire_view_with_budget(&candidate.data, 1, budget).map_err(map_arrangement)?;
    let original_view = original
        .map(|message| parse_wire_view_with_budget(&message.data, 1, budget))
        .transpose()
        .map_err(map_arrangement)?;
    let original_fields = original_view.as_ref().map_or(0, |view| view.len());
    let pair_work = candidate_view
        .len()
        .checked_mul(original_fields)
        .ok_or(Error::Verification)?;
    budget.preflight_work(pair_work).map_err(map_arrangement)?;
    budget.work(pair_work).map_err(map_arrangement)?;
    let mut candidate_unknown = candidate_view.fields().filter(|field| {
        !allowed_fields.contains(&field.number())
            && !SYMBOL_GENERATED_FIELDS.contains(&field.number())
    });
    let mut original_unknown = original_view.as_ref().map(|view| {
        view.fields().filter(|field| {
            !allowed_fields.contains(&field.number())
                && !SYMBOL_GENERATED_FIELDS.contains(&field.number())
        })
    });
    loop {
        match (
            candidate_unknown.next(),
            original_unknown.as_mut().and_then(Iterator::next),
        ) {
            (Some(candidate), Some(original)) if candidate.raw() == original.raw() => {},
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
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
    let bound = budget
        .series_limits()
        .map_err(map_arrangement)?
        .max_output_bytes();
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
        let chart_identifier = target(self, sheet, chart, &mut budget)?.chart;
        Ok(Edit {
            source: self,
            sheet,
            chart,
            chart_identifier,
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
        apply_series_patch(self, patch)
    }
}
