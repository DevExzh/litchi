//! Exact-source chart lifecycle of a Numbers sheet chart (record 0416; the
//! host's `set_sheet_chart_kind` and `set_sheet_chart_direction`).
//!
//! The chart is addressed by its sheet and a [`ChartSelector`]. An [`Edit`]
//! stages kind and series-direction changes; [`Edit::commit`] rewrites the
//! two fields of the chart drawable's chart archive in place through the
//! shared `litchi_iwa_style::chart_lifecycle` values (every other byte of
//! the payload and package is kept), publishes once and reads the chart
//! back. The chart's frame is edited through
//! [`crate::sheet::drawable::properties`] and its data through
//! [`crate::sheet::chart::data`].

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::chart::Direction;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_core::RawMessage;
use litchi_iwa_style::chart_lifecycle as engine;
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, CHART_MESSAGE_TYPE};
use super::structure::{self, Fault, Published};
use crate::selector::{ChartSelector, SheetSelector};
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("chart-lifecycle");

/// A typed, content-redacted chart-lifecycle failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted chart graph or the chart payload is malformed.
    #[error("rooted chart-lifecycle source is invalid")]
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
    #[error("could not allocate {amount} chart-lifecycle items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("chart-lifecycle candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("chart-lifecycle patch conflicts with this package")]
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
    /// The kind is not a supported concrete chart kind.
    #[error("the chart kind cannot be assigned")]
    UnsupportedKind,
    /// The series direction is not a supported direction.
    #[error("the series direction cannot be assigned")]
    UnsupportedDirection,
    /// Commit was requested before staging a change.
    #[error("no chart-lifecycle change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph.
    #[error("a surviving object depends on the chart")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support chart-lifecycle edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::ChartNotFound => RefusalKind::NotFound,
            Self::UnsupportedKind | Self::UnsupportedDirection => RefusalKind::InvalidValue,
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

impl From<engine::Error> for Error {
    fn from(_error: engine::Error) -> Self {
        Self::InvalidSource
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Change {
    Kind(Kind),
    Direction(Direction),
}

/// Chart-lifecycle changes staged against one chart of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    chart: ChartSelector,
    changes: Vec<Change>,
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
    /// Stage a new native chart kind, keeping the chart's data and styles
    /// (the host's `set_sheet_chart_kind`).
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedKind`] for an undefined or unsupported kind;
    /// [`Error::Allocation`].
    pub fn set_kind(&mut self, kind: Kind) -> Result<&mut Self, Error> {
        if !engine::assignable_kind(kind) {
            return Err(Refusal::UnsupportedKind.into());
        }
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(Change::Kind(kind));
        Ok(self)
    }

    /// Stage whether the chart's series are its rows or its columns (the
    /// host's `set_sheet_chart_direction`).
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedDirection`]; [`Error::Allocation`].
    pub fn set_direction(&mut self, direction: Direction) -> Result<&mut Self, Error> {
        if direction.is_unsupported() {
            return Err(Refusal::UnsupportedDirection.into());
        }
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(Change::Direction(direction));
        Ok(self)
    }

    /// Validate and atomically publish the staged changes. A change that
    /// leaves the chart as it is is skipped; an edit whose changes are all
    /// skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; source, resource and verification
    /// errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let (identifier, index) = chart_drawable(self.source, self.sheet, self.chart)?;
        let object = graph::object(self.source, identifier)?;
        let before = object.messages[index].data.clone();
        let mut payload = before.clone();
        let mut expected = engine::values(&before, WireLimits::DEFAULT)?;
        for change in &self.changes {
            payload = match *change {
                Change::Kind(kind) => {
                    expected.0 = Some(kind);
                    engine::with_kind(&payload, kind, WireLimits::DEFAULT)?
                },
                Change::Direction(direction) => {
                    expected.1 = Some(direction);
                    engine::with_direction(&payload, direction, WireLimits::DEFAULT)?
                },
            };
        }
        if engine::values(&before, WireLimits::DEFAULT)? == expected {
            return Commit::unchanged(self.source);
        }
        let mut data = Some(payload);
        let mut replace = |object: &mut litchi_iwa_core::ArchiveObject,
                           limits: litchi_iwa_core::Limits| {
            object
                .replace_message_preserving_header(
                    index,
                    RawMessage {
                        type_: CHART_MESSAGE_TYPE,
                        data: data.take().ok_or(Fault::Invalid)?,
                    },
                    limits,
                )
                .map(drop)
                .map_err(|error| structure::map_core(&error))
        };
        let published: Published =
            graph::publish_in_place(self.source, &mut [(identifier, &mut replace)])?;
        if read(&published.package, self.sheet, self.chart)? != expected {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

/// The chart drawable object and its chart message.
fn chart_drawable(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
) -> Result<(u64, usize), Error> {
    let drawables = graph::sheet_drawables(package, sheet)?;
    let entry = drawables
        .select(DrawableSelector::from(chart))
        .ok_or(Refusal::ChartNotFound)?;
    let (index, message_type) = entry.message.ok_or(Error::InvalidSource)?;
    if message_type != CHART_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
    }
    Ok((entry.identifier, index))
}

fn read(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
) -> Result<(Option<Kind>, Option<Direction>), Error> {
    let (identifier, index) = chart_drawable(package, sheet, chart)?;
    let object = graph::object(package, identifier)?;
    Ok(engine::values(
        &object.messages[index].data,
        WireLimits::DEFAULT,
    )?)
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// Whether one sheet chart's series are its rows or its columns.
    ///
    /// # Errors
    ///
    /// Selector and source errors; a chart without a stored direction is
    /// [`Error::InvalidSource`].
    pub fn sheet_chart_direction<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Direction, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        read(self, sheet, chart.into())?
            .1
            .ok_or(Error::InvalidSource)
    }

    /// Start an immutable chart-lifecycle edit of one sheet chart.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::chart::lifecycle::{Direction, Kind};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_sheet_chart_lifecycle(0, 0)?;
    /// edit.set_kind(Kind::Bar2d)?.set_direction(Direction::Columns)?;
    /// let commit = edit.commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_sheet_chart_lifecycle(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn edit_sheet_chart_lifecycle<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let chart = chart.into();
        chart_drawable(self, sheet, chart)?;
        Ok(Edit {
            source: self,
            sheet,
            chart,
            changes: Vec::new(),
        })
    }

    /// Apply a chart-lifecycle patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_chart_lifecycle(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
