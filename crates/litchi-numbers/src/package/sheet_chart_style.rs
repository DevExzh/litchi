//! Implementation of the public [`crate::sheet::chart::style`] transaction:
//! chart-level inspector properties of one sheet chart on the shared chart
//! style engine (`litchi_iwa_protos::chart_style_codec`, record 0375),
//! record 0378.
//!
//! The chart's style objects are found through the codec's reference
//! projection of the chart drawable; each change maps a common value
//! (`litchi_iwa_common::{chart, chart3d}`) to the codec's library write rule
//! and rewrites the one style payload in place. This is the behaviour of
//! the host editor's `set_sheet_chart_{pie_start_angle, donut_inner_radius,
//! radar_start_angle, 3d_series_gap, 3d_rotation, 3d_depth, rounded_corners}`
//! routes, collapsed into one property verb.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::chart::legend;
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::chart_style_codec::{self as codec, Archive, Limits};
use litchi_iwa_style::chart_style::chart as engine;
use litchi_iwa_style::series::{Refusal as EngineRefusal, SymbolError};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, CHART_MESSAGE_TYPE};
use super::structure::{self, Fault, Published};
use crate::selector::{ChartSelector, SheetSelector};
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("chart-style");

pub use litchi_iwa_style::chart_style::chart::{Property, PropertyKind, RadarGridShape};

/// A typed, content-redacted chart-style failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted chart graph or a style payload is malformed.
    #[error("rooted chart-style source is invalid")]
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
    #[error("could not allocate {amount} chart-style items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("chart-style candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("chart-style patch conflicts with this package")]
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
    /// The property does not apply to the chart's kind (a pie start angle
    /// of a column chart, a 3D depth of a 2D chart).
    #[error("the property does not apply to this chart kind")]
    UnsupportedKind,
    /// The chart has no style object of the property's archive.
    #[error("the chart has no style object for the property")]
    MissingStyle,
    /// Commit was requested before staging a change.
    #[error("no chart-style change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph.
    #[error("a surviving object depends on the chart style")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support chart-style edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::ChartNotFound => RefusalKind::NotFound,
            Self::UnsupportedKind | Self::MissingStyle => RefusalKind::UnsupportedTarget,
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

impl From<codec::Error> for Error {
    fn from(error: codec::Error) -> Self {
        match error {
            codec::Error::Allocation { bytes, .. } => Self::Allocation { amount: bytes },
            _ => Self::InvalidSource,
        }
    }
}

impl From<SymbolError> for Error {
    fn from(error: SymbolError) -> Self {
        match error {
            SymbolError::Refused(EngineRefusal::UnsupportedKind) => Refusal::UnsupportedKind.into(),
            SymbolError::Graph(litchi_iwa_style::graph::Error::Codec(error)) => error.into(),
            _ => Self::InvalidSource,
        }
    }
}

/// One staged change.
#[derive(Debug, Clone, Copy, PartialEq)]
enum Staged {
    Set(Property),
    Reset(PropertyKind),
}

/// Chart-style changes staged against one chart of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    chart: ChartSelector,
    changes: Vec<Staged>,
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
    /// Stage `property` as the chart's value. Changes apply in order.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set(&mut self, property: Property) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(Staged::Set(property));
        Ok(self)
    }

    /// Stage resetting `kind` to its native default.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn reset(&mut self, kind: PropertyKind) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(Staged::Reset(kind));
        Ok(self)
    }

    /// Validate and atomically publish the staged changes. A change that
    /// leaves a property as it is is skipped; an edit whose changes are all
    /// skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::UnsupportedKind`],
    /// [`Refusal::MissingStyle`], source and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for change in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let property = match *change {
                Staged::Set(property) => property,
                Staged::Reset(kind) => kind.native_default(),
            };
            if let Some(published) = apply_one(base, self.sheet, self.chart, property)? {
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

/// The chart's drawable, kind and style object of `archive`.
struct Target {
    kind: ChartKind,
    style: u64,
    index: usize,
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

fn read_legend(package: &Package, sheet: usize, chart: ChartSelector) -> Result<Property, Error> {
    let (identifier, index) = chart_drawable(package, sheet, chart)?;
    let object = graph::object(package, identifier)?;
    Ok(Property::LegendFrame(engine::legend_frame(
        &object.messages[index].data,
    )?))
}

fn apply_legend(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    frame: legend::Frame,
) -> Result<Option<Published>, Error> {
    if read_legend(source, sheet, chart)? == Property::LegendFrame(frame) {
        return Ok(None);
    }
    let (identifier, index) = chart_drawable(source, sheet, chart)?;
    let object = graph::object(source, identifier)?;
    let mut data = Some(engine::rewrite_legend_frame(
        &object.messages[index].data,
        frame,
    )?);
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
    let published = graph::publish_in_place(source, &mut [(identifier, &mut replace)])?;
    if read_legend(&published.package, sheet, chart)? != Property::LegendFrame(frame) {
        return Err(Error::Verification);
    }
    Ok(Some(published))
}

fn target(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    archive: Archive,
) -> Result<Target, Error> {
    let drawables = graph::sheet_drawables(package, sheet)?;
    let entry = drawables
        .select(DrawableSelector::from(chart))
        .ok_or(Refusal::ChartNotFound)?;
    let (index, message_type) = entry.message.ok_or(Error::InvalidSource)?;
    if message_type != CHART_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
    }
    let kind = package
        .sheet_chart_metadata(sheet, chart)
        .map_err(|_error| Error::InvalidSource)?
        .kind();
    let object = graph::object(package, entry.identifier)?;
    let references = codec::references(&object.messages[index].data, Limits::DEFAULT)?;
    let style = references
        .chart_level(archive)
        .ok_or(Refusal::MissingStyle)?;
    let style_object = graph::object(package, style)?;
    let style_index = structure::unique_message(style_object, &[archive.message_type()])?;
    Ok(Target {
        kind,
        style,
        index: style_index,
    })
}

fn read(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let Some(archive) = kind.archive() else {
        return read_legend(package, sheet, chart);
    };
    let target = target(package, sheet, chart, archive)?;
    if !kind.applies_to(target.kind) {
        return Err(Refusal::UnsupportedKind.into());
    }
    let object = graph::object(package, target.style)?;
    let properties = codec::decode(
        &object.messages[target.index].data,
        archive,
        Limits::DEFAULT,
    )?;
    Ok(engine::value(&properties, kind, target.kind)?)
}

fn apply_one(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    value: Property,
) -> Result<Option<Published>, Error> {
    if let Property::LegendFrame(frame) = value {
        return apply_legend(source, sheet, chart, frame);
    }
    let kind = value.kind();
    let archive = kind.archive().ok_or(Error::InvalidSource)?;
    let target = target(source, sheet, chart, archive)?;
    if !kind.applies_to(target.kind) {
        return Err(Refusal::UnsupportedKind.into());
    }
    let object = graph::object(source, target.style)?;
    let payload = &object.messages[target.index].data;
    let properties = codec::decode(payload, archive, Limits::DEFAULT)?;
    if engine::value(&properties, kind, target.kind)? == value {
        return Ok(None);
    }
    let rewritten = engine::rewrite(payload, value, target.kind)?;
    if rewritten == *payload {
        return Ok(None);
    }
    let mut data = Some(rewritten);
    let message_type = archive.message_type();
    let index = target.index;
    let mut replace = |object: &mut litchi_iwa_core::ArchiveObject,
                       limits: litchi_iwa_core::Limits| {
        object
            .replace_message_preserving_header(
                index,
                RawMessage {
                    type_: message_type,
                    data: data.take().ok_or(Fault::Invalid)?,
                },
                limits,
            )
            .map(drop)
            .map_err(|error| structure::map_core(&error))
    };
    let published = graph::publish_in_place(source, &mut [(target.style, &mut replace)])?;
    if read(&published.package, sheet, chart, kind)? != value {
        return Err(Error::Verification);
    }
    Ok(Some(published))
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// The stored (or native default) value of one chart-level style
    /// property of one sheet chart.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedKind`],
    /// [`Refusal::MissingStyle`] and source errors.
    pub fn sheet_chart_style<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        read(self, sheet, chart.into(), kind)
    }

    /// Start an immutable chart-level style edit of one sheet chart.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::chart::style::PropertyKind;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_sheet_chart_style(0, 0)?;
    /// edit.reset(PropertyKind::RoundedCorners)?;
    /// let commit = edit.commit()?;
    /// let restored = commit.package().apply_sheet_chart_style(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// Selector errors and source errors.
    pub fn edit_sheet_chart_style<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let chart = chart.into();
        graph::sheet_drawables(self, sheet)?
            .select(DrawableSelector::from(chart))
            .ok_or(Refusal::ChartNotFound)?;
        Ok(Edit {
            source: self,
            sheet,
            chart,
            changes: Vec::new(),
        })
    }

    /// Apply a chart-style patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_chart_style(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
