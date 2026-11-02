//! Implementation of the public [`crate::sheet::chart::axis`] transaction:
//! the axis-style switches of one sheet chart's primary value or category
//! axis (line, gridlines, minor tick marks, the value axis's minimum label,
//! the major tick-mark location) on the shared chart style codec
//! (`litchi_iwa_protos::chart_style_codec`, record 0375), record 0378.
//!
//! Each property is one field of the axis style's generated extension; a
//! write rewrites the primary axis style in place when only this chart uses
//! it (the host's `AxisStyleSlot::ensure_exclusive`). The host editor's
//! `set_sheet_chart_axis_*_visible`, `set_sheet_chart_axis_tick_mark_location`
//! and `set_sheet_chart_value_axis_minimum_label_visible` routes collapse
//! into [`Edit::set`].

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::chart_style_codec::{self as codec, Archive, Limits};
use litchi_iwa_style::chart_style::axis as engine;
use litchi_iwa_style::series::{Refusal as EngineRefusal, SymbolError};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, CHART_MESSAGE_TYPE};
use super::structure::{self, Fault, Published};
use crate::selector::{ChartSelector, SheetSelector};
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("chart-axis");

pub use litchi_iwa_style::chart_style::axis::{Property, PropertyKind};

/// A typed, content-redacted chart-axis failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted chart graph or an axis style is malformed.
    #[error("rooted chart-axis source is invalid")]
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
    #[error("could not allocate {amount} chart-axis items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("chart-axis candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("chart-axis patch conflicts with this package")]
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
    /// The chart has no such axis (a pie chart), or the property does not
    /// apply to it (the minimum label of a category axis).
    #[error("the chart has no such axis property")]
    UnsupportedAxis,
    /// Another chart uses the axis style.
    #[error("the axis style is shared with another chart")]
    SharedStyle,
    /// Commit was requested before staging a change.
    #[error("no chart-axis change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph.
    #[error("a surviving object depends on the chart axis style")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support chart-axis edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::ChartNotFound => RefusalKind::NotFound,
            Self::UnsupportedAxis => RefusalKind::UnsupportedTarget,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::SharedStyle | Self::UnsupportedDependency | Self::UnsupportedSource => {
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

impl From<SymbolError> for Error {
    fn from(error: SymbolError) -> Self {
        match error {
            SymbolError::Refused(EngineRefusal::UnsupportedAxis) => Refusal::UnsupportedAxis.into(),
            SymbolError::Graph(litchi_iwa_style::graph::Error::Codec(error)) => error.into(),
            _ => Self::InvalidSource,
        }
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

/// Axis-style changes staged against one chart of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    chart: ChartSelector,
    changes: Vec<(Axis, Property)>,
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
    /// Stage `property` on the chart's primary `axis`. Changes apply in
    /// order.
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedAxis`] for the minimum label of a category
    /// axis, [`Error::Allocation`].
    pub fn set(&mut self, axis: Axis, property: Property) -> Result<&mut Self, Error> {
        if property.kind().field(axis).is_none() {
            return Err(Refusal::UnsupportedAxis.into());
        }
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push((axis, property));
        Ok(self)
    }

    /// Validate and atomically publish the staged changes. A change that
    /// leaves a property as it is is skipped; an edit whose changes are all
    /// skipped is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::UnsupportedAxis`],
    /// [`Refusal::SharedStyle`], source and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for (axis, property) in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) = apply_one(base, self.sheet, self.chart, *axis, *property)? {
                if read(
                    &published.package,
                    self.sheet,
                    self.chart,
                    *axis,
                    property.kind(),
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

/// The primary axis style of the chart: (object, message index).
fn axis_style(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    axis: Axis,
    archive: Archive,
) -> Result<(u64, usize), Error> {
    let drawables = graph::sheet_drawables(package, sheet)?;
    let entry = drawables
        .select(DrawableSelector::from(chart))
        .ok_or(Refusal::ChartNotFound)?;
    let (index, message_type) = entry.message.ok_or(Error::InvalidSource)?;
    if message_type != CHART_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
    }
    let object = graph::object(package, entry.identifier)?;
    let references = codec::references(&object.messages[index].data, Limits::DEFAULT)?;
    let styles = axis_objects(&references, axis, archive);
    let style = *styles.first().ok_or(Refusal::UnsupportedAxis)?;
    let style_object = graph::object(package, style)?;
    let style_index = structure::unique_message(style_object, &[archive.message_type()])?;
    Ok((style, style_index))
}

/// The chart's style or non-style objects of `axis`.
fn axis_objects(references: &codec::StyleReferences, axis: Axis, archive: Archive) -> &[u64] {
    match (axis, archive) {
        (Axis::Value, Archive::AxisNonStyle) => references.value_axis_non_styles(),
        (_, Archive::AxisNonStyle) => references.category_axis_non_styles(),
        (Axis::Value, _) => references.value_axis_styles(),
        _ => references.category_axis_styles(),
    }
}

/// Whether only one chart axis references `style` (the host's
/// `ensure_exclusive`).
fn exclusive(package: &Package, style: u64, archive: Archive) -> Result<bool, Error> {
    let mut owners = 0_usize;
    for object in package.state.components.iter_objects() {
        for message in &object.messages {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            let references = codec::references(&message.data, Limits::DEFAULT)?;
            owners += axis_objects(&references, Axis::Value, archive)
                .iter()
                .chain(axis_objects(&references, Axis::Category, archive))
                .filter(|identifier| **identifier == style)
                .count();
        }
    }
    Ok(owners == 1)
}

fn read(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    axis: Axis,
    kind: PropertyKind,
) -> Result<Property, Error> {
    if kind.field(axis).is_none() {
        return Err(Refusal::UnsupportedAxis.into());
    }
    let archive = kind.archive();
    let (style, index) = axis_style(package, sheet, chart, axis, archive)?;
    let object = graph::object(package, style)?;
    let properties = codec::decode(&object.messages[index].data, archive, Limits::DEFAULT)?;
    Ok(engine::value(&properties, axis, kind)?)
}

fn apply_one(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    axis: Axis,
    value: Property,
) -> Result<Option<Published>, Error> {
    let kind = value.kind();
    if read(source, sheet, chart, axis, kind)? == value {
        return Ok(None);
    }
    let archive = kind.archive();
    let (style, index) = axis_style(source, sheet, chart, axis, archive)?;
    if !exclusive(source, style, archive)? {
        return Err(Refusal::SharedStyle.into());
    }
    let object = graph::object(source, style)?;
    let rewritten = engine::rewrite(&object.messages[index].data, axis, value)?;
    if rewritten == object.messages[index].data {
        return Ok(None);
    }
    let mut data = Some(rewritten);
    let message_type = archive.message_type();
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
    Ok(Some(graph::publish_in_place(
        source,
        &mut [(style, &mut replace)],
    )?))
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// The stored (or native default) value of one axis-style property of a
    /// sheet chart's primary `axis`.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::UnsupportedAxis`] and source errors.
    pub fn sheet_chart_axis<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
        axis: Axis,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        read(self, sheet, chart.into(), axis, kind)
    }

    /// Start an immutable axis-style edit of one sheet chart.
    ///
    /// # Errors
    ///
    /// Selector errors and source errors.
    pub fn edit_sheet_chart_axis<'sheet>(
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

    /// Apply a chart-axis patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_chart_axis(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
