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

use litchi_iwa_common::chart::axis::style::Visibility;
use litchi_iwa_common::chart::axis::{Axis, TickMarkLocation};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::chart_style_codec::{self as codec, Archive, Limits, Write};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, CHART_MESSAGE_TYPE};
use super::structure::{self, Fault, Published};
use crate::selector::{ChartSelector, SheetSelector};
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("chart-axis");

/// One axis-style property with its value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Property {
    /// Whether the axis line shows.
    Line(Visibility),
    /// Whether major gridlines show.
    MajorGridlines(Visibility),
    /// Whether minor gridlines show.
    MinorGridlines(Visibility),
    /// Whether minor tick marks show.
    MinorTickMarks(Visibility),
    /// Where major tick marks are drawn.
    TickMarkLocation(TickMarkLocation),
    /// Whether the value axis shows its minimum label (value axis only).
    MinimumLabel(Visibility),
}

/// The property a [`Property`] value sets, for reads.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyKind {
    /// [`Property::Line`].
    Line,
    /// [`Property::MajorGridlines`].
    MajorGridlines,
    /// [`Property::MinorGridlines`].
    MinorGridlines,
    /// [`Property::MinorTickMarks`].
    MinorTickMarks,
    /// [`Property::TickMarkLocation`].
    TickMarkLocation,
    /// [`Property::MinimumLabel`].
    MinimumLabel,
}

impl Property {
    /// The property this value sets.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        match self {
            Self::Line(_) => PropertyKind::Line,
            Self::MajorGridlines(_) => PropertyKind::MajorGridlines,
            Self::MinorGridlines(_) => PropertyKind::MinorGridlines,
            Self::MinorTickMarks(_) => PropertyKind::MinorTickMarks,
            Self::TickMarkLocation(_) => PropertyKind::TickMarkLocation,
            Self::MinimumLabel(_) => PropertyKind::MinimumLabel,
        }
    }
}

impl PropertyKind {
    /// The extension field of the property on `axis` (the host's
    /// `axis_style` constants), or `None` when the axis has no such field.
    const fn field(self, axis: Axis) -> Option<u32> {
        Some(match (self, axis) {
            (Self::Line, Axis::Category) => 24,
            (Self::Line, Axis::Value) => 25,
            (Self::MajorGridlines, Axis::Category) => 27,
            (Self::MajorGridlines, Axis::Value) => 28,
            (Self::MinorGridlines, Axis::Category) => 32,
            (Self::MinorGridlines, Axis::Value) => 33,
            (Self::MinorTickMarks, Axis::Category) => 34,
            (Self::MinorTickMarks, Axis::Value) => 35,
            (Self::TickMarkLocation, Axis::Category) => 36,
            (Self::TickMarkLocation, Axis::Value) => 37,
            (Self::MinimumLabel, Axis::Value) => 31,
            _ => return None,
        })
    }

    /// The value of a switch the style does not store.
    const fn default_visible(self) -> bool {
        matches!(self, Self::MinimumLabel | Self::MinorTickMarks)
    }
}

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
    let styles = match axis {
        Axis::Value => references.value_axis_styles(),
        _ => references.category_axis_styles(),
    };
    let style = *styles.first().ok_or(Refusal::UnsupportedAxis)?;
    let style_object = graph::object(package, style)?;
    let style_index =
        structure::unique_message(style_object, &[Archive::AxisStyle.message_type()])?;
    Ok((style, style_index))
}

/// Whether only one chart axis references `style` (the host's
/// `ensure_exclusive`).
fn exclusive(package: &Package, style: u64) -> Result<bool, Error> {
    let mut owners = 0_usize;
    for object in package.state.components.iter_objects() {
        for message in &object.messages {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            let references = codec::references(&message.data, Limits::DEFAULT)?;
            owners += references
                .value_axis_styles()
                .iter()
                .chain(references.category_axis_styles())
                .filter(|identifier| **identifier == style)
                .count();
        }
    }
    Ok(owners == 1)
}

const fn visibility(visible: bool) -> Visibility {
    if visible {
        Visibility::Visible
    } else {
        Visibility::Hidden
    }
}

fn read(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    axis: Axis,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let field = kind.field(axis).ok_or(Refusal::UnsupportedAxis)?;
    let (style, index) = axis_style(package, sheet, chart, axis)?;
    let object = graph::object(package, style)?;
    let properties = codec::decode(
        &object.messages[index].data,
        Archive::AxisStyle,
        Limits::DEFAULT,
    )?;
    if kind == PropertyKind::TickMarkLocation {
        return Ok(Property::TickMarkLocation(
            properties
                .int32(field)?
                .map(TickMarkLocation::from_native)
                .unwrap_or_default(),
        ));
    }
    let visible = visibility(properties.bool(field)?.unwrap_or(kind.default_visible()));
    Ok(match kind {
        PropertyKind::Line => Property::Line(visible),
        PropertyKind::MajorGridlines => Property::MajorGridlines(visible),
        PropertyKind::MinorGridlines => Property::MinorGridlines(visible),
        PropertyKind::MinorTickMarks => Property::MinorTickMarks(visible),
        _ => Property::MinimumLabel(visible),
    })
}

fn apply_one(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    axis: Axis,
    value: Property,
) -> Result<Option<Published>, Error> {
    let kind = value.kind();
    let field = kind.field(axis).ok_or(Refusal::UnsupportedAxis)?;
    if read(source, sheet, chart, axis, kind)? == value {
        return Ok(None);
    }
    let (style, index) = axis_style(source, sheet, chart, axis)?;
    if !exclusive(source, style)? {
        return Err(Refusal::SharedStyle.into());
    }
    let write = match value {
        Property::TickMarkLocation(location) => Write::Int32 {
            field,
            value: Some(location.native_value()),
        },
        Property::Line(visible)
        | Property::MajorGridlines(visible)
        | Property::MinorGridlines(visible)
        | Property::MinorTickMarks(visible)
        | Property::MinimumLabel(visible) => Write::Bool {
            field,
            value: Some(visible.is_visible()),
        },
    };
    let object = graph::object(source, style)?;
    let rewritten = codec::rewrite(
        &object.messages[index].data,
        Archive::AxisStyle,
        &[write],
        codec::EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    if !rewritten.changed() {
        return Ok(None);
    }
    let mut data = Some(rewritten.into_bytes().into_owned());
    let message_type = Archive::AxisStyle.message_type();
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
