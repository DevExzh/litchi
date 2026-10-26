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
use litchi_iwa_common::chart::{corner, donut, pie, radar};
use litchi_iwa_common::chart3d::{self, Scale};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::chart_style_codec::{
    self as codec, Archive, Limits, Properties, field, property,
};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, CHART_MESSAGE_TYPE};
use super::structure::{self, Fault, Published};
use crate::selector::{ChartSelector, SheetSelector};
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("chart-style");

/// One chart-level style property with its value.
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// Wedges rotation of a pie or donut chart.
    PieStartAngle(pie::StartAngle),
    /// Segment inner radius of a donut chart.
    DonutInnerRadius(donut::InnerRadius),
    /// Rotation of a radar chart.
    RadarStartAngle(radar::StartAngle),
    /// Gap between series of a 3D line or area chart.
    SeriesGap3d(chart3d::SeriesGap),
    /// Scene rotation of a 3D chart.
    Rotation3d(chart3d::Rotation),
    /// Depth of a 3D chart.
    Depth3d(chart3d::Depth),
    /// Rounded corners of bars and columns.
    RoundedCorners(corner::Rounding),
}

/// The property a [`Property`] value sets, for reads and resets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyKind {
    /// [`Property::PieStartAngle`].
    PieStartAngle,
    /// [`Property::DonutInnerRadius`].
    DonutInnerRadius,
    /// [`Property::RadarStartAngle`].
    RadarStartAngle,
    /// [`Property::SeriesGap3d`].
    SeriesGap3d,
    /// [`Property::Rotation3d`].
    Rotation3d,
    /// [`Property::Depth3d`].
    Depth3d,
    /// [`Property::RoundedCorners`].
    RoundedCorners,
}

impl Property {
    /// The property this value sets.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        match self {
            Self::PieStartAngle(_) => PropertyKind::PieStartAngle,
            Self::DonutInnerRadius(_) => PropertyKind::DonutInnerRadius,
            Self::RadarStartAngle(_) => PropertyKind::RadarStartAngle,
            Self::SeriesGap3d(_) => PropertyKind::SeriesGap3d,
            Self::Rotation3d(_) => PropertyKind::Rotation3d,
            Self::Depth3d(_) => PropertyKind::Depth3d,
            Self::RoundedCorners(_) => PropertyKind::RoundedCorners,
        }
    }
}

impl PropertyKind {
    /// The native default of this property (what a reset writes).
    #[must_use]
    pub const fn native_default(self) -> Property {
        match self {
            Self::PieStartAngle => Property::PieStartAngle(pie::StartAngle::ZERO),
            Self::DonutInnerRadius => Property::DonutInnerRadius(donut::InnerRadius::DEFAULT),
            Self::RadarStartAngle => Property::RadarStartAngle(radar::StartAngle::ZERO),
            Self::SeriesGap3d => Property::SeriesGap3d(chart3d::SeriesGap::NATIVE_DEFAULT),
            Self::Rotation3d => Property::Rotation3d(chart3d::Rotation::DEFAULT),
            Self::Depth3d => Property::Depth3d(chart3d::Depth::DEFAULT),
            Self::RoundedCorners => Property::RoundedCorners(corner::Rounding::NONE),
        }
    }

    const fn archive(self) -> Archive {
        match self {
            Self::RoundedCorners => Archive::ChartStyle,
            _ => Archive::ChartNonStyle,
        }
    }

    const fn applies_to(self, kind: ChartKind) -> bool {
        match self {
            Self::PieStartAngle => kind.supports_pie_start_angle(),
            Self::DonutInnerRadius => kind.supports_donut_inner_radius(),
            Self::RadarStartAngle => kind.supports_radar_start_angle(),
            Self::SeriesGap3d => kind.supports_3d_series_gap(),
            Self::Rotation3d => kind.supports_3d_scene(),
            Self::Depth3d => kind.supports_3d_depth(),
            Self::RoundedCorners => true,
        }
    }
}

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

/// A stored value outside its semantic range.
fn invalid<E>(_error: E) -> Error {
    Error::InvalidSource
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

const fn scale3d(scale: Scale) -> property::Scale3d {
    match scale {
        Scale::Area => property::Scale3d::Area,
        Scale::Bar => property::Scale3d::Bar,
        Scale::Column => property::Scale3d::Column,
        Scale::Line => property::Scale3d::Line,
        _ => property::Scale3d::Pie,
    }
}

/// The 3D depth field of a chart kind.
const fn depth_field(scale: Scale) -> u32 {
    scale3d(scale).field()
}

/// The stored value of `kind`.
fn read_value(
    properties: &Properties<'_>,
    kind: PropertyKind,
    chart: ChartKind,
) -> Result<Property, Error> {
    use field::{chart_non_style as non_style, chart_style as style};
    Ok(match kind {
        PropertyKind::PieStartAngle => Property::PieStartAngle(
            properties
                .float(non_style::PIE_START_ANGLE)?
                .map(pie::StartAngle::from_degrees)
                .transpose()
                .map_err(invalid)?
                .unwrap_or_default(),
        ),
        PropertyKind::DonutInnerRadius => Property::DonutInnerRadius(
            properties
                .float(non_style::DONUT_INNER_RADIUS)?
                .map(donut::InnerRadius::from_native)
                .transpose()
                .map_err(invalid)?
                .unwrap_or_default(),
        ),
        PropertyKind::RadarStartAngle => Property::RadarStartAngle(
            properties
                .float(non_style::RADAR_START_ANGLE)?
                .map(radar::StartAngle::from_degrees)
                .transpose()
                .map_err(invalid)?
                .unwrap_or_default(),
        ),
        PropertyKind::SeriesGap3d => Property::SeriesGap3d(
            properties
                .float(non_style::SERIES_GAP_3D)?
                .map(chart3d::SeriesGap::from_native_ratio)
                .transpose()
                .map_err(invalid)?
                .unwrap_or_default(),
        ),
        PropertyKind::Rotation3d => Property::Rotation3d(
            properties
                .vector(non_style::ROTATION_3D)?
                .map(|v| chart3d::Rotation::from_native([v.x, v.y, v.z, v.w]))
                .transpose()
                .map_err(invalid)?
                .unwrap_or_default(),
        ),
        PropertyKind::Depth3d => {
            let scale = Scale::for_kind(chart).ok_or(Refusal::UnsupportedKind)?;
            Property::Depth3d(
                properties
                    .vector(depth_field(scale))?
                    .map(|v| chart3d::Depth::from_native(scale, [v.x, v.y, v.z, v.w]))
                    .transpose()
                    .map_err(invalid)?
                    .unwrap_or_default(),
            )
        },
        PropertyKind::RoundedCorners => Property::RoundedCorners(corner::Rounding::new(
            properties
                .float(style::ROUNDED_CORNER_RADIUS)?
                .map(corner::Radius::from_native)
                .transpose()
                .map_err(invalid)?
                .unwrap_or(corner::Radius::ZERO),
            properties
                .bool(style::ROUNDED_CORNER_OUTER_END_ONLY)?
                .unwrap_or(false),
        )),
    })
}

/// The codec's library write rule for `value`.
fn write_for(value: Property, chart: ChartKind) -> Result<property::PropertyWrite<'static>, Error> {
    Ok(match value {
        Property::PieStartAngle(angle) => property::pie_start_angle(angle.degrees()),
        Property::DonutInnerRadius(radius) => {
            property::donut_inner_radius(radius.native_fraction())
        },
        Property::RadarStartAngle(angle) => property::radar_start_angle(angle.degrees()),
        Property::SeriesGap3d(gap) => property::series_gap_3d(gap.native_ratio()),
        Property::Rotation3d(rotation) => {
            property::rotation_3d(rotation.x_degrees(), rotation.y_degrees())
        },
        Property::Depth3d(depth) => {
            let scale = Scale::for_kind(chart).ok_or(Refusal::UnsupportedKind)?;
            property::depth_3d(
                scale3d(scale),
                depth.native_ratio(scale),
                (depth != chart3d::Depth::DEFAULT).then(|| depth.native_vector(scale)),
            )
        },
        Property::RoundedCorners(rounding) => property::rounded_corners(
            rounding.radius().native_fraction(),
            rounding.outside_corners_only(),
        ),
    })
}

fn read(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let archive = kind.archive();
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
    read_value(&properties, kind, target.kind)
}

fn apply_one(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    value: Property,
) -> Result<Option<Published>, Error> {
    let kind = value.kind();
    let archive = kind.archive();
    let target = target(source, sheet, chart, archive)?;
    if !kind.applies_to(target.kind) {
        return Err(Refusal::UnsupportedKind.into());
    }
    let object = graph::object(source, target.style)?;
    let payload = &object.messages[target.index].data;
    let properties = codec::decode(payload, archive, Limits::DEFAULT)?;
    if read_value(&properties, kind, target.kind)? == value {
        return Ok(None);
    }
    let write = write_for(value, target.kind)?;
    let rewritten = codec::rewrite(
        payload,
        archive,
        write.writes(),
        write.empty(),
        Limits::DEFAULT,
    )?;
    if !rewritten.changed() {
        return Ok(None);
    }
    let mut data = Some(rewritten.into_bytes().into_owned());
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
