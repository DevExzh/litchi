//! Selector-first, source-preserving frame, property and playback edits of
//! body-anchored images, movies, audio clips and shapes.
//!
//! The drawable is found by kind and text-flow position in the rooted body
//! storage; the shared [`super::body_drawable_graph`] proof establishes its
//! ownership. Each staged change is a field-level rewrite of the selected
//! drawable message (and, when a movie or audio clip moves, of its body
//! attachment's offsets): unknown fields, field order and every other object
//! stay byte-identical.

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::media::playback::MediaPlaybackSettings;
use litchi_iwa_common::shape::geometry::{FlipAxis, Point, Size};
use litchi_iwa_common::wire::{patch_fixed32_field, transform_length_delimited_field};
use litchi_iwa_protos::shape_properties_codec as codec;
use thiserror::Error as ThisError;

use litchi_iwa_common::shape::line::Segment;
use litchi_iwa_protos::shape_style_codec as line_codec;

use super::Package;
use super::body_drawable_graph::{
    self as graph, Budget, GraphError, GraphLimit, MessageRewrite, ObjectIndex, kind_name,
    message_type,
};
use super::movie_playback::{__decode_movie_playback_payload, __rewrite_movie_playback_payload};
use crate::drawable::{Geometry, GeometryFlags, Properties};
use crate::selector::DrawableSelector;

const SHAPE_INFO_SHAPE_FIELD: u32 = 1;
const MEDIA_DRAWABLE_FIELD: u32 = 1;
const IMAGE_ORIGINAL_SIZE_FIELD: u32 = 4;
const IMAGE_NATURAL_SIZE_FIELD: u32 = 9;
const MOVIE_ORIGINAL_SIZE_FIELD: u32 = 20;
const MOVIE_NATURAL_SIZE_FIELD: u32 = 21;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;
const ATTACHMENT_H_OFFSET_TYPE_FIELD: u32 = 2;
const ATTACHMENT_H_OFFSET_FIELD: u32 = 3;
const ATTACHMENT_V_OFFSET_TYPE_FIELD: u32 = 4;
const ATTACHMENT_V_OFFSET_FIELD: u32 = 5;
const H_OFFSET_BODY_MARGIN: u64 = 0;
const V_OFFSET_PAGE: u64 = 1;

/// A finite resource charged by one body-drawable property operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected by the graph proof.
    InputBytes,
    /// Bytes in the rewritten payload or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and reassembly work.
    WireWork,
    /// Semantic section or body-storage cardinality.
    PayloadItems,
    /// Semantic UTF-8 text retained by the reopened document.
    TextBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::PayloadItems => "payload items",
            Self::TextBytes => "text bytes",
        })
    }
}

/// Failure from a body-drawable property read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted drawable graph or the selected payload is malformed.
    #[error("the selected Pages body-drawable source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Pages body-drawable properties {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for Pages body-drawable properties")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested state.
    #[error("the edited Pages body-drawable properties failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body-drawable properties patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No body drawable of the selected kind is at the selected position.
    #[error("the Pages body has no {kind} at position {position:?}")]
    NotFound {
        /// The selected kind: `image`, `movie` or `audio`.
        kind: &'static str,
        /// The selected position.
        position: Position,
    },
    /// The edit does not apply to the selected kind (for example a size or
    /// a flip of an audio clip, or playback of an image).
    #[error("{operation} does not apply to a Pages body {kind}")]
    UnsupportedOperation {
        /// The refused verb.
        operation: &'static str,
        /// The selected kind.
        kind: &'static str,
    },
    /// The value cannot be stored (non-finite or negative frame parts, an
    /// invalid playback value, a movie frame without a position).
    #[error("invalid Pages body-drawable value: {reason}")]
    InvalidValue {
        /// Why the value was refused.
        reason: &'static str,
    },
    /// The drawable lacks the stored state the verb needs (no frame, no
    /// layout flags, no original size).
    #[error("the Pages body {kind} has no {missing}")]
    MissingState {
        /// The selected kind.
        kind: &'static str,
        /// The missing state.
        missing: &'static str,
    },
    /// The movie or audio attachment uses offset bases this route does not
    /// rewrite.
    #[error("the Pages body media attachment uses unsupported offset bases")]
    UnsupportedAnchor,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact body-drawable edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedOperation { .. } => RefusalKind::UnsupportedOperation,
            Self::InvalidValue { .. } => RefusalKind::InvalidValue,
            Self::MissingState { .. } => RefusalKind::UnsupportedTarget,
            Self::UnsupportedAnchor | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<GraphError> for Error {
    fn from(error: GraphError) -> Self {
        match error {
            GraphError::InvalidSource => Self::InvalidSource,
            GraphError::Allocation { amount } => Self::Allocation { amount },
            GraphError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    GraphLimit::InputBytes => LimitKind::InputBytes,
                    GraphLimit::OutputBytes => LimitKind::OutputBytes,
                    GraphLimit::WireFields => LimitKind::WireFields,
                    GraphLimit::WireNesting => LimitKind::WireNesting,
                    GraphLimit::WireWork => LimitKind::WireWork,
                    GraphLimit::SemanticItems => LimitKind::PayloadItems,
                    GraphLimit::SemanticTextBytes => LimitKind::TextBytes,
                },
                observed,
                maximum,
            },
        }
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        graph::map_wire_error(error).into()
    }
}

/// The stored frame, properties and playback of one body drawable.
#[derive(Debug, Clone, PartialEq)]
pub struct Snapshot {
    selector: DrawableSelector,
    geometry: Option<Geometry>,
    properties: Properties,
    playback: Option<MediaPlaybackSettings>,
    original_size: Option<Size>,
    natural_size: Option<Size>,
    line_segment: Option<Segment>,
    line_reversed: bool,
}

impl Snapshot {
    /// The segment a native straight line draws, from its start to its end
    /// (`None` for every other drawable).
    #[must_use]
    pub const fn line_segment(&self) -> Option<Segment> {
        self.line_segment
    }

    /// The selector this snapshot was read through.
    #[must_use]
    pub const fn selector(&self) -> DrawableSelector {
        self.selector
    }

    /// The stored frame; `None` when the drawable has no frame.
    #[must_use]
    pub const fn geometry(&self) -> Option<Geometry> {
        self.geometry
    }

    /// The shared drawable properties.
    #[must_use]
    pub const fn properties(&self) -> &Properties {
        &self.properties
    }

    /// Trim, poster, repeat and volume settings of a movie or audio clip;
    /// `None` for an image.
    #[must_use]
    pub const fn playback(&self) -> Option<MediaPlaybackSettings> {
        self.playback
    }

    /// The stored original media size (the Original Size target).
    #[must_use]
    pub const fn original_size(&self) -> Option<Size> {
        self.original_size
    }

    /// The stored untransformed media size.
    #[must_use]
    pub const fn natural_size(&self) -> Option<Size> {
        self.natural_size
    }
}

/// Frame, property and playback changes staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    before: Snapshot,
    after: Snapshot,
    move_attachment: bool,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// The state read when the edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// The state staged for publication.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    fn kind_name(&self) -> &'static str {
        kind_name(self.target.selector)
    }

    fn stored_geometry(&self) -> Result<Geometry, Error> {
        self.after.geometry.ok_or_else(|| {
            Refusal::MissingState {
                kind: self.kind_name(),
                missing: "frame",
            }
            .into()
        })
    }

    /// Replace the frame of an image, movie or shape.
    ///
    /// A movie frame must have a position: Pages keeps a movie's body
    /// attachment offsets in step with it, and this edit moves them too.
    ///
    /// # Errors
    ///
    /// Refuses audio (use [`Self::set_position`]), a drawable without a
    /// frame, and non-finite or negative parts.
    pub fn set_geometry(&mut self, geometry: Geometry) -> Result<&mut Self, Error> {
        if matches!(self.target.selector, DrawableSelector::Audio(_)) {
            return Err(Refusal::UnsupportedOperation {
                operation: "set_geometry",
                kind: "audio",
            }
            .into());
        }
        self.stored_geometry()?;
        if !geometry.is_valid() {
            return Err(Refusal::InvalidValue {
                reason: "frame parts must be finite and sizes non-negative",
            }
            .into());
        }
        if matches!(self.target.selector, DrawableSelector::Movie(_)) {
            if geometry.position.is_none() {
                return Err(Refusal::InvalidValue {
                    reason: "a movie frame needs a position",
                }
                .into());
            }
            self.move_attachment = true;
        }
        self.after.geometry = Some(geometry);
        Ok(self)
    }

    /// Move the drawable; a movie or audio clip's attachment offsets move
    /// with it.
    ///
    /// # Errors
    ///
    /// Refuses a drawable without a frame and a non-finite position.
    pub fn set_position(&mut self, position: Point) -> Result<&mut Self, Error> {
        let geometry = self.stored_geometry()?;
        if !position.x.is_finite() || !position.y.is_finite() {
            return Err(Refusal::InvalidValue {
                reason: "a position must be finite",
            }
            .into());
        }
        if matches!(
            self.target.selector,
            DrawableSelector::Movie(_) | DrawableSelector::Audio(_)
        ) {
            self.move_attachment = true;
        }
        self.after.geometry = Some(Geometry {
            position: Some(position),
            ..geometry
        });
        Ok(self)
    }

    /// Apply one native Arrange Flip command to an image, movie or shape.
    ///
    /// # Errors
    ///
    /// Refuses audio and a frame without layout flags.
    pub fn flip(&mut self, axis: FlipAxis) -> Result<&mut Self, Error> {
        if matches!(self.target.selector, DrawableSelector::Audio(_)) {
            return Err(Refusal::UnsupportedOperation {
                operation: "flip",
                kind: "audio",
            }
            .into());
        }
        let geometry = self.stored_geometry()?;
        let flipped = geometry.flipped(axis).ok_or(Refusal::MissingState {
            kind: self.kind_name(),
            missing: "layout flags",
        })?;
        self.after.geometry = Some(flipped);
        Ok(self)
    }

    /// Restore the displayed size of an image or movie from its stored
    /// original size (the Original Size command); position, flags and
    /// rotation are kept.
    ///
    /// # Errors
    ///
    /// Refuses audio, shapes and a drawable without a positive original size.
    pub fn restore_original_size(&mut self) -> Result<&mut Self, Error> {
        if matches!(
            self.target.selector,
            DrawableSelector::Audio(_) | DrawableSelector::Shape(_) | DrawableSelector::TextBox(_)
        ) {
            return Err(Refusal::UnsupportedOperation {
                operation: "restore_original_size",
                kind: self.kind_name(),
            }
            .into());
        }
        let geometry = self.stored_geometry()?;
        let original = self
            .after
            .original_size
            .filter(|size| {
                size.width.is_finite()
                    && size.height.is_finite()
                    && size.width > 0.0
                    && size.height > 0.0
            })
            .ok_or(Refusal::MissingState {
                kind: self.kind_name(),
                missing: "original size",
            })?;
        self.after.geometry = Some(Geometry {
            size: Some(original),
            ..geometry
        });
        Ok(self)
    }

    /// Move a native straight line onto `segment`, keeping its stored
    /// direction (the frame, flags, rotation and the path's length follow).
    ///
    /// # Errors
    ///
    /// Refuses a drawable that is not a native straight line.
    pub fn set_line_segment(&mut self, segment: Segment) -> Result<&mut Self, Error> {
        if self.after.line_segment.is_none() {
            return Err(Refusal::UnsupportedOperation {
                operation: "set_line_segment",
                kind: self.kind_name(),
            }
            .into());
        }
        let directed = if self.after.line_reversed {
            segment.reversed()
        } else {
            segment
        };
        let (position, size, angle) = directed.geometry();
        self.after.geometry = Some(Geometry {
            position: Some(position),
            size: Some(size),
            flags: Some(GeometryFlags::from_bits(line_codec::LINE_FLAGS)),
            angle: Some(angle),
        });
        self.after.line_segment = Some(segment);
        Ok(self)
    }

    /// Replace the shared drawable properties.
    pub fn set_properties(&mut self, properties: Properties) -> &mut Self {
        self.after.properties = properties;
        self
    }

    /// Replace the playback settings of a movie or audio clip.
    ///
    /// # Errors
    ///
    /// Refuses an image or shape and an invalid playback value.
    pub fn set_playback(&mut self, settings: MediaPlaybackSettings) -> Result<&mut Self, Error> {
        if matches!(
            self.target.selector,
            DrawableSelector::Image(_) | DrawableSelector::Shape(_) | DrawableSelector::TextBox(_)
        ) {
            return Err(Refusal::UnsupportedOperation {
                operation: "set_playback",
                kind: self.kind_name(),
            }
            .into());
        }
        let settings = settings.canonicalize().map_err(|_| Refusal::InvalidValue {
            reason: "invalid playback settings",
        })?;
        self.after.playback = Some(settings);
        Ok(self)
    }

    /// Validate and publish the staged edit.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// Exact-source reversible body-drawable properties patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    selector: DrawableSelector,
    before: Snapshot,
    after: Snapshot,
    touched_components: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("selector", &self.selector)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The selected drawable.
    #[must_use]
    pub const fn selector(&self) -> DrawableSelector {
        self.selector
    }

    /// The state required before this patch applies.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// The state this patch produces.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// The number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(&self) -> usize {
        self.touched_components
    }

    /// Whether both the semantic state and the exact artifact are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selector: self.selector,
            before: self.after.clone(),
            after: self.before.clone(),
            touched_components: self.touched_components,
        }
    }
}

/// Content-free diagnostics of one body-drawable properties publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
}

impl Diagnostics {
    /// Whether package bytes and semantic state changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// The number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
}

/// The validated immutable result of one body-drawable properties transaction.
#[must_use = "a body-drawable properties commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Target {
    selector: DrawableSelector,
    attachment: NonZeroU64,
    attachment_message: usize,
    attachment_component: Arc<str>,
    drawable: NonZeroU64,
    drawable_message: usize,
    drawable_component: Arc<str>,
    left_margin_bits: Option<u32>,
}

impl Package {
    /// Read the frame, properties and playback of one body image, movie or
    /// audio clip.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when no drawable matches or the graph is invalid.
    pub fn body_drawable_properties(
        &self,
        selector: impl Into<DrawableSelector>,
    ) -> Result<Snapshot, Error> {
        Ok(resolve(self, selector.into())?.1)
    }

    /// Begin a selector-first frame, property and playback edit of one body
    /// image, movie or audio clip.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when no drawable matches or the graph is invalid.
    pub fn edit_body_drawable_properties(
        &self,
        selector: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let (target, before) = resolve(self, selector.into())?;
        Ok(Edit {
            source: self,
            target,
            after: before.clone(),
            before,
            move_attachment: false,
        })
    }

    /// Replay an exact-source body-drawable properties patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact or the drawable no longer matches its `before` state.
    pub fn apply_body_drawable_properties(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = self.state.source.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let (_, current) = resolve(self, patch.selector)?;
        if current != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics {
                    changed: false,
                    touched_components: 0,
                },
            });
        }
        let candidate = graph::reopen(self, patch.artifacts.target())?;
        let (_, reopened) = resolve(&candidate, patch.selector)?;
        if reopened != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components: patch.touched_components,
            },
        })
    }
}

fn resolve(package: &Package, selector: DrawableSelector) -> Result<(Target, Snapshot), Error> {
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let selected =
        graph::select_drawable(&index, &body, selector, &mut budget)?.ok_or(Refusal::NotFound {
            kind: kind_name(selector),
            position: selector.position(),
        })?;
    let snapshot = read_snapshot(
        selector,
        selected.payload,
        selected.drawable_payload,
        &budget,
    )?;
    let target = Target {
        selector,
        attachment: selected.anchor.attachment,
        attachment_message: selected.anchor.attachment_message,
        attachment_component: index.component_name(selected.attachment.component_index)?,
        drawable: selected.anchor.drawable,
        drawable_message: selected.message_index,
        drawable_component: index.component_name(selected.drawable.component_index)?,
        left_margin_bits: body.root.left_margin.map(f32::to_bits),
    };
    Ok((target, snapshot))
}

fn validate_attachment_offsets(payload: &[u8], budget: &mut Budget) -> Result<(), Error> {
    let view = budget.parse(payload, 2)?;
    let h_type = graph::unique_field(&view, ATTACHMENT_H_OFFSET_TYPE_FIELD, 0)?
        .map(|field| graph::canonical_varint(field.payload()))
        .transpose()?;
    let v_type = graph::unique_field(&view, ATTACHMENT_V_OFFSET_TYPE_FIELD, 0)?
        .map(|field| graph::canonical_varint(field.payload()))
        .transpose()?;
    if h_type != Some(H_OFFSET_BODY_MARGIN) || v_type != Some(V_OFFSET_PAGE) {
        return Err(Refusal::UnsupportedAnchor.into());
    }
    for number in [ATTACHMENT_H_OFFSET_FIELD, ATTACHMENT_V_OFFSET_FIELD] {
        let field = graph::unique_field(&view, number, 5)?.ok_or(Error::InvalidSource)?;
        let value = f32::from_le_bytes(
            field
                .payload()
                .try_into()
                .map_err(|_| Error::InvalidSource)?,
        );
        if !value.is_finite() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn codec_options(source: &[u8], budget: &Budget) -> codec::DecodeOptions {
    let limits = budget.wire_limits();
    let bytes = source.len().max(1);
    codec::DecodeOptions::for_source(source)
        .with_max_input_bytes(bytes.min(limits.max_input_bytes()))
        .with_max_output_bytes(bytes.saturating_add(4_096).min(limits.max_output_bytes()))
}

fn map_codec_error(error: &codec::DecodeError) -> Error {
    match error.limit_kind() {
        Some(_) => Error::Limit {
            kind: LimitKind::WireWork,
            observed: 0,
            maximum: 0,
        },
        None => Error::InvalidSource,
    }
}

fn read_snapshot(
    selector: DrawableSelector,
    payload: &[u8],
    drawable_payload: &[u8],
    budget: &Budget,
) -> Result<Snapshot, Error> {
    let drawable =
        codec::decode_drawable_archive(drawable_payload, codec_options(drawable_payload, budget))
            .map_err(|error| map_codec_error(&error))?;
    let geometry = drawable.geometry().map(|geometry| Geometry {
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
    });
    if geometry.is_some_and(|geometry| !geometry.is_valid()) {
        return Err(Error::InvalidSource);
    }
    let stored = drawable.properties();
    let properties = Properties {
        hyperlink: stored.hyperlink_url().map(str::to_owned),
        locked: stored.locked(),
        aspect_ratio_locked: stored.aspect_ratio_locked(),
        accessibility_description: stored.accessibility_description().map(str::to_owned),
    };
    let sizes = match selector {
        DrawableSelector::Image(_) => Some((IMAGE_ORIGINAL_SIZE_FIELD, IMAGE_NATURAL_SIZE_FIELD)),
        DrawableSelector::Movie(_) | DrawableSelector::Audio(_) => {
            Some((MOVIE_ORIGINAL_SIZE_FIELD, MOVIE_NATURAL_SIZE_FIELD))
        },
        DrawableSelector::Shape(_) | DrawableSelector::TextBox(_) => None,
    };
    let (original_size, natural_size) = match sizes {
        Some((original_field, natural_field)) => {
            let view = litchi_iwa_common::wire::WireView::parse(payload, budget.wire_limits())?;
            (
                stored_size(&view, original_field, budget)?,
                stored_size(&view, natural_field, budget)?,
            )
        },
        None => (None, None),
    };
    let playback = match selector {
        DrawableSelector::Image(_) | DrawableSelector::Shape(_) | DrawableSelector::TextBox(_) => {
            None
        },
        DrawableSelector::Movie(_) | DrawableSelector::Audio(_) => Some(
            __decode_movie_playback_payload(payload, budget.wire_limits())
                .map_err(|_| Error::InvalidSource)?,
        ),
    };
    let (line_segment, line_reversed) = match selector {
        DrawableSelector::Shape(_) => line_of(payload)?,
        _ => (None, false),
    };
    Ok(Snapshot {
        selector,
        geometry,
        properties,
        playback,
        original_size,
        natural_size,
        line_segment,
        line_reversed,
    })
}

/// The segment of a native straight line and whether its stored path runs
/// backwards (the host's `body_line_segment`).
fn line_of(shape_info: &[u8]) -> Result<(Option<Segment>, bool), Error> {
    let Some(stored) = line_codec::line(shape_info, line_codec::Limits::DEFAULT)
        .map_err(|_| Error::InvalidSource)?
    else {
        return Ok((None, false));
    };
    let [x, y] = stored.geometry.position.ok_or(Error::InvalidSource)?;
    let [width, height] = stored.geometry.size.ok_or(Error::InvalidSource)?;
    let segment = Segment::from_geometry(
        Point { x, y },
        Size { width, height },
        stored.geometry.angle.ok_or(Error::InvalidSource)?,
        stored.natural_width,
    )
    .map_err(|_| Error::InvalidSource)?;
    Ok(if stored.horizontal_flip {
        (Some(segment.reversed()), true)
    } else {
        (Some(segment), false)
    })
}

fn stored_size(
    view: &litchi_iwa_common::wire::WireView<'_>,
    number: u32,
    budget: &Budget,
) -> Result<Option<Size>, Error> {
    let Some(field) = graph::unique_field(view, number, 2)? else {
        return Ok(None);
    };
    let size = litchi_iwa_common::wire::WireView::parse(field.payload(), budget.wire_limits())?;
    let mut parts = [0.0_f32; 2];
    for (slot, part) in [SIZE_WIDTH_FIELD, SIZE_HEIGHT_FIELD]
        .into_iter()
        .zip(&mut parts)
    {
        let field = graph::unique_field(&size, slot, 5)?.ok_or(Error::InvalidSource)?;
        *part = f32::from_le_bytes(
            field
                .payload()
                .try_into()
                .map_err(|_| Error::InvalidSource)?,
        );
    }
    Ok(Some(Size {
        width: parts[0],
        height: parts[1],
    }))
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let source_bytes = source.state.source.shared_source();
    let geometry_changed = edit.after.geometry != edit.before.geometry;
    let properties_changed = edit.after.properties != edit.before.properties;
    let playback_changed = edit.after.playback != edit.before.playback;
    let line_changed = edit.after.line_segment != edit.before.line_segment;
    let position_changed = edit.move_attachment
        && edit.after.geometry.and_then(|geometry| geometry.position)
            != edit.before.geometry.and_then(|geometry| geometry.position);
    if !geometry_changed && !properties_changed && !playback_changed && !position_changed {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                selector: edit.target.selector,
                before: edit.before.clone(),
                after: edit.before,
                touched_components: 0,
            },
            diagnostics: Diagnostics {
                changed: false,
                touched_components: 0,
            },
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let (current, before) = resolve(source, edit.target.selector)?;
    if current != edit.target || before != edit.before {
        return Err(Error::PatchConflict);
    }
    let budget = Budget::new(source)?;
    let drawable_object = find_object(source, &current.drawable_component, current.drawable)?;
    let mut payload = drawable_object
        .messages
        .get(current.drawable_message)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    let mut before = edit.before.clone();
    if line_changed {
        let directed = edit.after.geometry.ok_or(Error::InvalidSource)?;
        let segment = edit.after.line_segment.ok_or(Error::InvalidSource)?;
        let geometry = line_codec::Geometry {
            position: directed.position.map(|point| [point.x, point.y]),
            size: directed.size.map(|size| [size.width, size.height]),
            flags: directed.flags.map(GeometryFlags::bits),
            angle: directed.angle,
        };
        payload = line_codec::rewrite_line(
            &payload,
            geometry,
            segment.length(),
            line_codec::Limits::DEFAULT,
        )
        .map_err(|_| Error::InvalidSource)?
        .into_bytes()
        .into_owned();
        // The frame is written; only property changes remain.
        before.geometry = edit.after.geometry;
    }
    if (geometry_changed && !line_changed) || properties_changed {
        let rewrite = |drawable: &[u8]| rewrite_drawable(drawable, &before, &edit.after, &budget);
        payload = if matches!(
            current.selector,
            DrawableSelector::Shape(_) | DrawableSelector::TextBox(_)
        ) {
            transform_length_delimited_field(&payload, MEDIA_DRAWABLE_FIELD, |shape| {
                transform_length_delimited_field(shape, SHAPE_INFO_SHAPE_FIELD, rewrite)
            })?
        } else {
            transform_length_delimited_field(&payload, MEDIA_DRAWABLE_FIELD, rewrite)?
        };
    }
    if playback_changed {
        let settings = edit.after.playback.ok_or(Error::InvalidSource)?;
        payload = __rewrite_movie_playback_payload(&payload, settings, budget.wire_limits())
            .map_err(|_| Error::InvalidSource)?;
    }
    let message_type = message_type(current.selector);
    let mut rewrites = vec![MessageRewrite {
        component: Arc::clone(&current.drawable_component),
        object: current.drawable,
        message_index: current.drawable_message,
        message_type,
        data: payload,
    }];
    if position_changed {
        let position = edit
            .after
            .geometry
            .and_then(|geometry| geometry.position)
            .ok_or(Error::InvalidSource)?;
        let left_margin = current.left_margin_bits.map_or(0.0, f32::from_bits);
        let attachment_object =
            find_object(source, &current.attachment_component, current.attachment)?;
        let attachment = attachment_object
            .messages
            .get(current.attachment_message)
            .ok_or(Error::InvalidSource)?
            .data
            .as_slice();
        let mut offsets_budget = Budget::new(source)?;
        validate_attachment_offsets(attachment, &mut offsets_budget)?;
        let horizontal = position.x - left_margin;
        if !horizontal.is_finite() {
            return Err(Refusal::InvalidValue {
                reason: "the position overflows the attachment offset",
            }
            .into());
        }
        let attachment = patch_fixed32_field(
            attachment,
            ATTACHMENT_H_OFFSET_FIELD,
            true,
            Some(horizontal.to_bits()),
        )?;
        let attachment = patch_fixed32_field(
            &attachment,
            ATTACHMENT_V_OFFSET_FIELD,
            true,
            Some(position.y.to_bits()),
        )?;
        rewrites.push(MessageRewrite {
            component: Arc::clone(&current.attachment_component),
            object: current.attachment,
            message_index: current.attachment_message,
            message_type: graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
            data: attachment,
        });
    }
    let touched_components =
        if rewrites.len() > 1 && current.attachment_component != current.drawable_component {
            2
        } else {
            1
        };
    let candidate = graph::publish(source, rewrites)?;
    let (reopened_target, reopened) = resolve(&candidate, current.selector)?;
    // A line's endpoints are derived back from the stored frame and path
    // width, so they hold up to `f32` rounding; everything else is exact.
    let mut expected = edit.after.clone();
    if line_changed
        && let (Some(wanted), Some(read)) = (expected.line_segment, reopened.line_segment)
        && segments_match(wanted, read)
    {
        expected.line_segment = Some(read);
    }
    if reopened_target.drawable != current.drawable || reopened != expected {
        return Err(Error::Verification);
    }
    let edit_after = expected;
    let target_bytes = candidate.state.source.shared_source();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: ExactArtifacts::new(source_bytes, target_bytes),
            selector: current.selector,
            before: edit.before,
            after: edit_after,
            touched_components,
        },
        diagnostics: Diagnostics {
            changed: true,
            touched_components,
        },
    })
}

/// Whether two line segments agree up to the `f32` rounding of a frame
/// round trip.
fn segments_match(left: Segment, right: Segment) -> bool {
    let close = |a: f32, b: f32| (a - b).abs() <= 1e-3 * a.abs().max(b.abs()).max(1.0);
    close(left.start().x, right.start().x)
        && close(left.start().y, right.start().y)
        && close(left.end().x, right.end().x)
        && close(left.end().y, right.end().y)
}

fn find_object<'a>(
    package: &'a Package,
    component: &str,
    identifier: NonZeroU64,
) -> Result<&'a litchi_iwa_core::ArchiveObject, Error> {
    package
        .state
        .source
        .components()
        .get(component)
        .and_then(|component| component.archive().object(identifier.get()))
        .ok_or(Error::InvalidSource)
}

fn rewrite_drawable(
    source: &[u8],
    before: &Snapshot,
    after: &Snapshot,
    budget: &Budget,
) -> Result<Vec<u8>, Error> {
    let mut geometry = codec::GeometryWrite::preserve();
    if let (Some(old), Some(new)) = (before.geometry, after.geometry) {
        if old.position != new.position {
            geometry = geometry.with_position(
                new.position
                    .map(|point| codec::Point::new(point.x, point.y)),
            );
        }
        if old.size != new.size {
            geometry = geometry.with_size(
                new.size
                    .map(|size| codec::Size::new(size.width, size.height)),
            );
        }
        if old.flags != new.flags {
            geometry = geometry.with_flags(new.flags.map(GeometryFlags::bits));
        }
        if old.angle.map(f32::to_bits) != new.angle.map(f32::to_bits) {
            geometry = geometry.with_angle(new.angle);
        }
    }
    let mut properties = codec::DrawablePropertiesWrite::preserve();
    let (old, new) = (&before.properties, &after.properties);
    if old.hyperlink != new.hyperlink {
        properties = properties.with_hyperlink_url(new.hyperlink.as_deref());
    }
    if old.locked != new.locked {
        properties = properties.with_locked(new.locked);
    }
    if old.aspect_ratio_locked != new.aspect_ratio_locked {
        properties = properties.with_aspect_ratio_locked(new.aspect_ratio_locked);
    }
    if old.accessibility_description != new.accessibility_description {
        properties =
            properties.with_accessibility_description(new.accessibility_description.as_deref());
    }
    codec::rewrite_drawable_archive(source, geometry, properties, codec_options(source, budget))
        .map(codec::RewriteOutput::into_bytes)
        .map_err(|error| map_codec_error(&error))
}
