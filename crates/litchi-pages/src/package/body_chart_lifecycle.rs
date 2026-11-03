//! Exact-source chart lifecycle of a Pages body chart (record 0416; the
//! host's `set_body_chart_kind` and `set_body_chart_direction`).
//!
//! A chart is addressed by a [`BodyChartSelector`]. An [`Edit`] stages kind
//! and series-direction changes; [`Edit::commit`] rewrites the two fields of
//! the chart drawable's chart archive in place through the shared
//! `litchi_iwa_style::chart_lifecycle` values (every other byte is kept),
//! publishes once and reads the chart back. [`Package::body_chart_geometry`]
//! reads the chart's frame.

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::chart::Direction;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_style::chart_lifecycle as engine;
use thiserror::Error as ThisError;

use super::Package;
use super::body_chart_arrangement::{self as arrangement, ArrangementBudget, ChartTarget};
use super::body_drawable_graph::{self as graph, MessageRewrite};
use crate::selector::BodyChartSelector;

const CHART_MESSAGE_TYPE: u32 = 5_021;
/// `TSCH.ChartDrawableArchive.super` (`TSD.DrawableArchive`).
const CHART_DRAWABLE: u32 = 1;

/// A finite resource governed by a chart-lifecycle transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes or fields of the bounded chart traversal.
    Traversal,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("chart traversal")
    }
}

/// A typed, content-redacted chart-lifecycle failure.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted body or chart payload is malformed.
    #[error("the Pages chart-lifecycle source is invalid")]
    InvalidSource,
    /// A finite limit was exceeded.
    #[error("Pages chart-lifecycle {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exceeded resource.
        kind: LimitKind,
        /// The observed amount.
        observed: u64,
        /// The configured maximum.
        maximum: u64,
    },
    /// The published candidate did not read back as staged.
    #[error("the Pages chart-lifecycle candidate failed verification")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("the Pages chart-lifecycle patch does not match this package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by
/// [`Refusal::kind`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No body chart matches the selector.
    #[error("the body has no chart at {position:?}")]
    ChartNotFound {
        /// The requested chart position.
        position: Position,
    },
    /// The kind is not a supported concrete chart kind.
    #[error("the chart kind cannot be assigned")]
    UnsupportedKind,
    /// The series direction is not a supported direction.
    #[error("the series direction cannot be assigned")]
    UnsupportedDirection,
    /// Commit was requested before staging a change.
    #[error("no chart-lifecycle change is staged")]
    NoStagedOperation,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support chart-lifecycle edits")]
    UnsupportedSource,
    /// A frame needs a position and finite parts with non-negative sizes.
    #[error("the chart frame is invalid")]
    InvalidGeometry,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::ChartNotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedKind | Self::UnsupportedDirection | Self::InvalidGeometry => {
                RefusalKind::InvalidValue
            },
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::UnsupportedSource => RefusalKind::UnsupportedContent,
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

impl From<arrangement::Error> for Error {
    fn from(error: arrangement::Error) -> Self {
        match error {
            arrangement::Error::Refused(crate::chart::arrangement::Refusal::ChartNotFound {
                position,
            }) => Refusal::ChartNotFound { position }.into(),
            arrangement::Error::Limit {
                observed, maximum, ..
            } => Self::Limit {
                kind: LimitKind::Traversal,
                observed,
                maximum,
            },
            _ => Self::InvalidSource,
        }
    }
}

impl From<graph::GraphError> for Error {
    fn from(_error: graph::GraphError) -> Self {
        Self::InvalidSource
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Change {
    Kind(Kind),
    Direction(Direction),
    Geometry(crate::drawable::Geometry),
}

/// The kind and series direction a chart stores.
type Values = (Option<Kind>, Option<Direction>);

/// Chart-lifecycle changes staged against one body chart.
pub struct Edit<'a> {
    source: &'a Package,
    selector: BodyChartSelector,
    changes: Vec<Change>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("selector", &self.selector)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage a new native chart kind, keeping the chart's data and styles
    /// (the host's `set_body_chart_kind`).
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedKind`] for an undefined or unsupported kind.
    pub fn set_kind(&mut self, kind: Kind) -> Result<&mut Self, Error> {
        if !engine::assignable_kind(kind) {
            return Err(Refusal::UnsupportedKind.into());
        }
        self.changes.push(Change::Kind(kind));
        Ok(self)
    }

    /// Stage whether the chart's series are its rows or its columns (the
    /// host's `set_body_chart_direction`).
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedDirection`].
    pub fn set_direction(&mut self, direction: Direction) -> Result<&mut Self, Error> {
        if direction.is_unsupported() {
            return Err(Refusal::UnsupportedDirection.into());
        }
        self.changes.push(Change::Direction(direction));
        Ok(self)
    }

    /// Stage a new frame (the host's `set_body_chart_geometry`). Pages
    /// keeps a chart's body attachment offsets in step with its position,
    /// so the commit moves them too: the horizontal offset from the body's
    /// left margin, the vertical offset from the page.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidGeometry`] for a frame without a position or with
    /// non-finite parts or negative sizes.
    pub fn set_geometry(
        &mut self,
        geometry: crate::drawable::Geometry,
    ) -> Result<&mut Self, Error> {
        if !geometry.is_valid() || geometry.position.is_none() {
            return Err(Refusal::InvalidGeometry.into());
        }
        self.changes.push(Change::Geometry(geometry));
        Ok(self)
    }

    /// Validate and atomically publish the staged changes. Changes that
    /// leave the chart as it is make an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`]; source, limit and verification
    /// errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let (target, before_payload) = select(self.source, self.selector)?;
        let before = engine::values(&before_payload, WireLimits::DEFAULT)?;
        let mut after = before;
        let before_frame = geometry(&before_payload)?;
        let mut frame = before_frame;
        let mut payload = before_payload;
        for change in &self.changes {
            payload = match *change {
                Change::Geometry(new) => {
                    frame = Some(new);
                    with_geometry(&payload, new)?
                },
                Change::Kind(kind) => {
                    after.0 = Some(kind);
                    engine::with_kind(&payload, kind, WireLimits::DEFAULT)?
                },
                Change::Direction(direction) => {
                    after.1 = Some(direction);
                    engine::with_direction(&payload, direction, WireLimits::DEFAULT)?
                },
            };
        }
        let source_bytes = self.source.state.source.shared_source();
        if after == before && frame == before_frame {
            return Ok(Commit {
                package: self.source.clone(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    selector: self.selector,
                    before,
                    after,
                },
                diagnostics: Diagnostics { changed: false },
            });
        }
        if !self.source.state.source.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        let mut rewrites = vec![MessageRewrite {
            component: Arc::clone(&target.component_name),
            object: target.drawable_identifier,
            message_index: target.drawable_message_index,
            message_type: CHART_MESSAGE_TYPE,
            data: payload,
        }];
        let moved =
            frame.and_then(|frame| frame.position) != before_frame.and_then(|frame| frame.position);
        if moved {
            let position = frame
                .and_then(|frame| frame.position)
                .ok_or(Error::InvalidSource)?;
            rewrites.push(attachment_rewrite(self.source, &target, position)?);
        }
        let candidate = graph::publish(self.source, rewrites)?;
        if values(&candidate, self.selector)? != after
            || geometry(&select(&candidate, self.selector)?.1)? != frame
        {
            return Err(Error::Verification);
        }
        let target_bytes = candidate.state.source.shared_source();
        Ok(Commit {
            package: candidate,
            patch: Patch {
                artifacts: ExactArtifacts::new(source_bytes, target_bytes),
                selector: self.selector,
                before,
                after,
            },
            diagnostics: Diagnostics { changed: true },
        })
    }
}

/// `payload` with the chart drawable's frame set to `frame` (only the
/// parts that differ are rewritten).
pub(super) fn with_geometry(
    payload: &[u8],
    frame: crate::drawable::Geometry,
) -> Result<Vec<u8>, Error> {
    use litchi_iwa_protos::shape_properties_codec as codec;
    let current = geometry(payload)?;
    let mut write = codec::GeometryWrite::preserve();
    if current.and_then(|current| current.position) != frame.position {
        write = write.with_position(
            frame
                .position
                .map(|point| codec::Point::new(point.x, point.y)),
        );
    }
    if current.and_then(|current| current.size) != frame.size {
        write = write.with_size(
            frame
                .size
                .map(|size| codec::Size::new(size.width, size.height)),
        );
    }
    if current.and_then(|current| current.flags) != frame.flags {
        write = write.with_flags(frame.flags.map(crate::drawable::GeometryFlags::bits));
    }
    if current.and_then(|current| current.angle).map(f32::to_bits) != frame.angle.map(f32::to_bits)
    {
        write = write.with_angle(frame.angle);
    }
    litchi_iwa_common::wire::transform_length_delimited_field(payload, CHART_DRAWABLE, |drawable| {
        codec::rewrite_drawable_archive(
            drawable,
            write,
            codec::DrawablePropertiesWrite::preserve(),
            codec::DecodeOptions::for_source(drawable),
        )
        .map(codec::RewriteOutput::into_bytes)
        .map_err(|_error| litchi_iwa_common::Error::InvalidFormat("chart frame".to_owned()))
    })
    .map_err(|_error| Error::InvalidSource)
}

const ATTACHMENT_H_OFFSET: u32 = 3;
const ATTACHMENT_V_OFFSET: u32 = 5;

/// The chart's body attachment with its offsets moved to `position` (the
/// host's `set_chart_attachment_position`).
fn attachment_rewrite(
    package: &Package,
    target: &ChartTarget,
    position: crate::drawable::Point,
) -> Result<MessageRewrite, Error> {
    use litchi_iwa_common::wire::patch_fixed32_field;
    let component = package
        .state
        .source
        .components()
        .get_index(target.attachment_component_index)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.attachment_object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.attachment_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    let message = object
        .messages
        .get(target.attachment_message_index)
        .ok_or(Error::InvalidSource)?;
    let left_margin = {
        let mut budget = super::body_drawable_graph::Budget::new(package)?;
        let index = super::body_drawable_graph::ObjectIndex::new(package, &mut budget)?;
        graph::resolve_body(&index, &mut budget)?
            .root
            .left_margin
            .unwrap_or_default()
    };
    let horizontal = position.x - left_margin;
    if !horizontal.is_finite() {
        return Err(Refusal::InvalidGeometry.into());
    }
    let data = patch_fixed32_field(
        &message.data,
        ATTACHMENT_H_OFFSET,
        true,
        Some(horizontal.to_bits()),
    )
    .map_err(|_error| Error::InvalidSource)?;
    let data = patch_fixed32_field(&data, ATTACHMENT_V_OFFSET, true, Some(position.y.to_bits()))
        .map_err(|_error| Error::InvalidSource)?;
    Ok(MessageRewrite {
        component: Arc::from(component.name()),
        object: target.attachment_identifier,
        message_index: target.attachment_message_index,
        message_type: message.type_,
        data,
    })
}

/// The selected chart and an owned copy of its chart payload.
fn select(package: &Package, selector: BodyChartSelector) -> Result<(ChartTarget, Vec<u8>), Error> {
    let mut budget = ArrangementBudget::new(package)?;
    let target = arrangement::resolve_target(package, selector, &mut budget)?;
    let payload = arrangement::selected_chart_payload(package, &target)?.to_vec();
    Ok((target, payload))
}

fn values(package: &Package, selector: BodyChartSelector) -> Result<Values, Error> {
    let (_, payload) = select(package, selector)?;
    Ok(engine::values(&payload, WireLimits::DEFAULT)?)
}

/// A reversible, exact-source chart-lifecycle patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    selector: BodyChartSelector,
    before: Values,
    after: Values,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("selector", &self.selector)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether applying the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The patch that restores the source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selector: self.selector,
            before: self.after,
            after: self.before,
        }
    }
}

/// What a chart-lifecycle commit changed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether the commit changed the package.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A published chart-lifecycle edit.
#[must_use = "a chart-lifecycle commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The published package.
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// The published package, by value.
    pub fn into_package(self) -> Package {
        self.package
    }

    /// The published package and the patch that made it.
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// The patch that made the package.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What changed.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// The frame a chart payload's `TSD.DrawableArchive` stores.
pub(super) fn geometry(payload: &[u8]) -> Result<Option<crate::drawable::Geometry>, Error> {
    use crate::drawable::{Geometry, GeometryFlags, Point, Size};
    use litchi_iwa_protos::shape_properties_codec as codec;
    let view = litchi_iwa_common::wire::WireView::parse(payload, WireLimits::DEFAULT)
        .map_err(|_error| Error::InvalidSource)?;
    let mut drawable = None;
    for field in view.fields() {
        if field.number() == CHART_DRAWABLE {
            if field.wire_type() != 2 || drawable.replace(field.payload()).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    let drawable = drawable.ok_or(Error::InvalidSource)?;
    let decoded =
        codec::decode_drawable_archive(drawable, codec::DecodeOptions::for_source(drawable))
            .map_err(|_error| Error::InvalidSource)?;
    Ok(decoded.geometry().map(|geometry| Geometry {
        position: geometry.position().map(|point| Point {
            x: point.x(),
            y: point.y(),
        }),
        size: geometry.size().map(|size| Size {
            width: size.width(),
            height: size.height(),
        }),
        flags: geometry.flags().map(GeometryFlags::from_bits),
        angle: geometry.angle(),
    }))
}

impl Package {
    /// The stored frame of one body chart (`None` without a frame).
    ///
    /// # Errors
    ///
    /// Selector and source errors.
    pub fn body_chart_geometry(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Option<crate::drawable::Geometry>, Error> {
        let (_, payload) = select(self, selector.into())?;
        geometry(&payload)
    }

    /// Whether one body chart's series are its rows or its columns.
    ///
    /// # Errors
    ///
    /// Selector and source errors; a chart without a stored direction is
    /// [`Error::InvalidSource`].
    pub fn body_chart_direction(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Direction, Error> {
        values(self, selector.into())?.1.ok_or(Error::InvalidSource)
    }

    /// Begin staging chart-lifecycle changes on one body chart.
    ///
    /// ```no_run
    /// use litchi_pages::Package;
    /// use litchi_pages::chart::lifecycle::{Direction, Kind};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.pages")?;
    /// let mut edit = package.edit_body_chart_lifecycle(0)?;
    /// edit.set_kind(Kind::Bar2d)?.set_direction(Direction::Columns)?;
    /// let commit = edit.commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_body_chart_lifecycle(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::ChartNotFound`] and source errors.
    pub fn edit_body_chart_lifecycle(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        select(self, selector)?;
        Ok(Edit {
            source: self,
            selector,
            changes: Vec::new(),
        })
    }

    /// Replay an exact-source chart-lifecycle patch.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when the patch was made against another
    /// artifact.
    pub fn apply_body_chart_lifecycle(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
            || values(self, patch.selector)? != patch.before
        {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics { changed: false },
            });
        }
        let candidate = graph::reopen(self, patch.artifacts.target())?;
        if values(&candidate, patch.selector)? != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics { changed: true },
        })
    }
}
