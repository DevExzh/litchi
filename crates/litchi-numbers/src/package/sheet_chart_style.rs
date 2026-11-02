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
use litchi_iwa_common::shape::shadow::Shadow as StoredShadow;
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::chart_style_codec::{self as codec, Archive, Limits};
use litchi_iwa_style::chart_style::chart as engine;
use litchi_iwa_style::series::style as series_style;
use litchi_iwa_style::series::{Refusal as EngineRefusal, SymbolError};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, CHART_MESSAGE_TYPE};
use super::structure::{self, Fault, Published};
use crate::selector::{ChartSelector, SheetSelector};
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("chart-style");

pub use litchi_iwa_common::shape::shadow::Drop;
pub use litchi_iwa_style::chart_style::chart::{
    Lighting3d, Property, PropertyKind, RadarGridShape, RadarSeriesStyle, Shadow,
};

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
    /// Another chart uses the style object the write would change (the
    /// host's `ensure_exclusive` refusal; editing it would restyle that
    /// chart too).
    #[error("the chart style is shared with another chart")]
    SharedStyle,
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
#[derive(Debug, Clone, PartialEq)]
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
            let property = match change {
                Staged::Set(property) => property.clone(),
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

/// A chart-series failure as this family's error. The structural limits
/// keep their kind; the series engine's wire, archive and semantic limits
/// are reported as the payload or package bytes they bound.
fn from_series(error: super::sheet_chart_series::Error) -> Error {
    use super::sheet_chart_series::{Error as Series, LimitKind as SeriesLimit, Refusal as Why};
    match error {
        Series::Refused(Why::SheetNotFound) => Refusal::SheetNotFound.into(),
        Series::Refused(Why::ChartNotFound) => Refusal::ChartNotFound.into(),
        Series::Refused(Why::UnsupportedKind) => Refusal::UnsupportedKind.into(),
        Series::Refused(Why::SharedStyle) => Refusal::SharedStyle.into(),
        Series::Refused(Why::UnsupportedDependency) => Refusal::UnsupportedDependency.into(),
        Series::Refused(Why::UnsupportedSource) => Refusal::UnsupportedSource.into(),
        Series::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                SeriesLimit::PackageBytes | SeriesLimit::Archive(_) | SeriesLimit::Detection(_) => {
                    LimitKind::PackageBytes
                },
                SeriesLimit::Identifiers => LimitKind::Identifiers,
                SeriesLimit::ScratchBytes => LimitKind::ScratchBytes,
                SeriesLimit::AllocationEvents => LimitKind::AllocationEvents,
                _ => LimitKind::PayloadBytes,
            },
            observed,
            maximum,
        },
        Series::Allocation { amount } => Error::Allocation { amount },
        Series::Verification => Error::Verification,
        Series::PatchConflict => Error::PatchConflict,
        _ => Error::InvalidSource,
    }
}

/// The kind and the effective series styles of the visible series of one
/// sheet chart (the host's `effective_chart_series_style_slots`).
fn radar_styles(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
) -> Result<(ChartKind, Vec<u64>), Error> {
    let (identifier, kind, count) =
        super::sheet_chart_series::chart_series(package, sheet, chart).map_err(from_series)?;
    let object = graph::object(package, identifier)?;
    let index = structure::unique_message(object, &[CHART_MESSAGE_TYPE])?;
    let references = codec::references(&object.messages[index].data, Limits::DEFAULT)?;
    let mut styles = Vec::new();
    structure::reserve(&mut styles, count)?;
    for series in 0..count {
        styles.push(
            references
                .series_style(series)
                .ok_or(Error::InvalidSource)?,
        );
    }
    Ok((kind, styles))
}

/// The series style payload of `style`.
fn series_payload(package: &Package, style: u64) -> Result<&[u8], SymbolError> {
    let object = graph::object(package, style)
        .map_err(|_fault| SymbolError::Invalid("a series style is missing"))?;
    let index = structure::unique_message(object, &[Archive::SeriesStyle.message_type()])
        .map_err(|_fault| SymbolError::Invalid("a series style has no single payload"))?;
    Ok(&object.messages[index].data)
}

/// The uniform radar series style of one sheet chart (the host's
/// `sheet_chart_radar_series_style`).
fn read_radar(package: &Package, sheet: usize, chart: ChartSelector) -> Result<Property, Error> {
    let (kind, styles) = radar_styles(package, sheet, chart)?;
    if !PropertyKind::RadarSeriesStyle.applies_to(kind) {
        return Err(Refusal::UnsupportedKind.into());
    }
    Ok(Property::RadarSeriesStyle(series_style::radar::effective(
        &styles,
        |style| series_payload(package, style),
    )?))
}

/// The host's `set_sheet_chart_radar_series_style`: every visible series'
/// effective style whose value differs is rewritten in place, each
/// exclusive to the chart.
fn apply_radar(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    style: RadarSeriesStyle,
) -> Result<Option<Published>, Error> {
    let (kind, styles) = radar_styles(source, sheet, chart)?;
    if !PropertyKind::RadarSeriesStyle.applies_to(kind) {
        return Err(Refusal::UnsupportedKind.into());
    }
    if series_style::radar::effective(&styles, |identifier| series_payload(source, identifier))
        .is_ok_and(|current| current == style)
    {
        return Ok(None);
    }
    let rewrites = series_style::radar::rewrites(&styles, style, |identifier| {
        series_payload(source, identifier)
    })?;
    for (identifier, _) in &rewrites {
        if !exclusive_series(source, *identifier)? {
            return Err(Refusal::SharedStyle.into());
        }
    }
    if rewrites.is_empty() {
        return Ok(None);
    }
    let message_type = Archive::SeriesStyle.message_type();
    let mut closures = Vec::new();
    structure::reserve(&mut closures, rewrites.len())?;
    for (object, data) in rewrites {
        let index = structure::unique_message(graph::object(source, object)?, &[message_type])?;
        let mut data = Some(data);
        let replace = move |object: &mut litchi_iwa_core::ArchiveObject,
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
        closures.push((object, replace));
    }
    let mut edits: Vec<(u64, graph::ObjectEdit<'_>)> = Vec::new();
    structure::reserve(&mut edits, closures.len())?;
    for (object, replace) in &mut closures {
        edits.push((*object, replace));
    }
    let published = graph::publish_in_place(source, &mut edits)?;
    if read_radar(&published.package, sheet, chart)? != Property::RadarSeriesStyle(style) {
        return Err(Error::Verification);
    }
    Ok(Some(published))
}

fn read(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
    kind: PropertyKind,
) -> Result<Property, Error> {
    if kind == PropertyKind::RadarSeriesStyle {
        return read_radar(package, sheet, chart);
    }
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
    if kind == PropertyKind::Shadow {
        let mut uniform = None;
        for (style, index) in series_styles(package, sheet, chart)? {
            let shadow =
                series_style::series_shadow(&graph::object(package, style)?.messages[index].data)?;
            if uniform.is_some_and(|known| known != shadow) {
                return Err(Error::InvalidSource);
            }
            uniform = Some(shadow);
        }
        let series = uniform.ok_or(Error::InvalidSource)?;
        return Ok(Property::Shadow(engine::shadow(&properties, series)?));
    }
    Ok(engine::value(&properties, kind, target.kind)?)
}

/// The chart's series style objects (theme styles with private overrides,
/// all of them) and the index of each one's series style message.
fn series_styles(
    package: &Package,
    sheet: usize,
    chart: ChartSelector,
) -> Result<Vec<(u64, usize)>, Error> {
    let (identifier, index) = chart_drawable(package, sheet, chart)?;
    let object = graph::object(package, identifier)?;
    let references = codec::references(&object.messages[index].data, Limits::DEFAULT)?;
    let mut styles = Vec::new();
    for style in litchi_iwa_style::series::chart_series_styles(&references)? {
        let index = structure::unique_message(
            graph::object(package, style)?,
            &[Archive::SeriesStyle.message_type()],
        )?;
        styles.push((style, index));
    }
    Ok(styles)
}

/// Whether only one chart references series style `style` (the host's
/// `ChartSeriesStyleSlot::ensure_exclusive`).
fn exclusive_series(package: &Package, style: u64) -> Result<bool, Error> {
    let mut owners = 0_usize;
    for object in package.state.components.iter_objects() {
        for message in &object.messages {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            let references = codec::references(&message.data, Limits::DEFAULT)?;
            let private = references
                .series_private_styles()
                .is_some_and(|sparse| sparse.entries().iter().any(|(_, entry)| *entry == style));
            if private || references.series_theme_styles().contains(&style) {
                owners += 1;
            }
        }
    }
    Ok(owners == 1)
}

/// The host's `set_chart_shadow`: the chart style and every series style
/// must be exclusive; each series style stores the drop shadow (or the
/// disabled shadow) and the chart style the grouping switch, published in
/// place together.
fn apply_shadow(
    source: &Package,
    sheet: usize,
    chart: ChartSelector,
    target: &Target,
    value: &Property,
) -> Result<Option<Published>, Error> {
    let Property::Shadow(shadow) = *value else {
        return Err(Error::InvalidSource);
    };
    let styles = series_styles(source, sheet, chart)?;
    for &(style, _) in &styles {
        if !exclusive_series(source, style)? {
            return Err(Refusal::SharedStyle.into());
        }
    }
    let stored = shadow
        .drop_shadow()
        .map_or(StoredShadow::Disabled, StoredShadow::Drop);
    let mut rewrites: Vec<(u64, usize, u32, Vec<u8>)> = Vec::new();
    for (style, index) in styles {
        let payload = &graph::object(source, style)?.messages[index].data;
        let data = series_style::rewrite_series_shadow(payload, stored)?;
        if data != *payload {
            rewrites.push((style, index, Archive::SeriesStyle.message_type(), data));
        }
    }
    let payload = &graph::object(source, target.style)?.messages[target.index].data;
    let data = engine::rewrite(payload, value, target.kind)?;
    if data != *payload {
        rewrites.push((
            target.style,
            target.index,
            Archive::ChartStyle.message_type(),
            data,
        ));
    }
    if rewrites.is_empty() {
        return Ok(None);
    }
    let mut closures: Vec<_> = rewrites
        .into_iter()
        .map(|(object, index, message_type, data)| {
            let mut data = Some(data);
            let replace = move |object: &mut litchi_iwa_core::ArchiveObject,
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
            (object, replace)
        })
        .collect();
    let mut edits: Vec<(u64, graph::ObjectEdit<'_>)> = Vec::new();
    for (object, replace) in &mut closures {
        edits.push((*object, replace));
    }
    Ok(Some(graph::publish_in_place(source, &mut edits)?))
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
    if let Property::RadarSeriesStyle(style) = value {
        return apply_radar(source, sheet, chart, style);
    }
    let kind = value.kind();
    let archive = kind.archive().ok_or(Error::InvalidSource)?;
    let target = target(source, sheet, chart, archive)?;
    if !kind.applies_to(target.kind) {
        return Err(Refusal::UnsupportedKind.into());
    }
    let object = graph::object(source, target.style)?;
    let payload = &object.messages[target.index].data;
    let current = if kind == PropertyKind::Shadow {
        read(source, sheet, chart, kind)?
    } else {
        engine::value(
            &codec::decode(payload, archive, Limits::DEFAULT)?,
            kind,
            target.kind,
        )?
    };
    if current == value {
        return Ok(None);
    }
    if !exclusive(source, target.style, archive)? {
        return Err(Refusal::SharedStyle.into());
    }
    if kind == PropertyKind::Shadow {
        let published = apply_shadow(source, sheet, chart, &target, &value)?;
        if let Some(published) = &published
            && read(&published.package, sheet, chart, kind)? != value
        {
            return Err(Error::Verification);
        }
        return Ok(published);
    }
    let rewritten = engine::rewrite(payload, &value, target.kind)?;
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

/// Whether only one chart references `style` as its `archive` style (the
/// host's `ChartStyleSlot`/`ChartNonStyleSlot::ensure_exclusive`).
fn exclusive(package: &Package, style: u64, archive: Archive) -> Result<bool, Error> {
    let mut owners = 0_usize;
    for object in package.state.components.iter_objects() {
        for message in &object.messages {
            if message.type_ == CHART_MESSAGE_TYPE
                && codec::references(&message.data, Limits::DEFAULT)?.chart_level(archive)
                    == Some(style)
            {
                owners += 1;
            }
        }
    }
    Ok(owners == 1)
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
