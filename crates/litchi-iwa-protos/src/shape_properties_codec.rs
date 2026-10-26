//! Bounded, source-preserving geometry and drawable-property projection.
//!
//! The focused Pages, Numbers, and Keynote crates all edit the same small
//! native edge: a `TSD.DrawableArchive` geometry plus four user-facing
//! properties.  This module owns the wire policy for that edge.  It accepts
//! only the direct drawable payload; each format package is responsible for
//! selecting its shape object and splicing the returned bytes into the
//! surrounding `ShapeInfoArchive` or media archive.
//!
//! The semantic values are intentionally neutral.  Buffa lazy views are used
//! as a bounded ingress cross-check, while the caller-owned source bytes stay
//! authoritative for every rewrite.  Unknown fields, field order, and raw
//! nested payloads therefore survive a geometry or property edit.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict scanner is intentionally ordered before the private Buffa checks."
)]

use std::{fmt, mem::size_of, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_media_creation_generated::LitchiIwaProjection as geometry_projection;
use crate::buffa_keynote_media_properties_generated::LitchiIwaKeynoteMoviePropertiesProjection as properties_projection;

const GEOMETRY_FIELD: u32 = 1;
const HYPERLINK_FIELD: u32 = 4;
const LOCKED_FIELD: u32 = 5;
const ASPECT_RATIO_LOCKED_FIELD: u32 = 7;
const ACCESSIBILITY_DESCRIPTION_FIELD: u32 = 8;

const POSITION_FIELD: u32 = 1;
const SIZE_FIELD: u32 = 2;
const FLAGS_FIELD: u32 = 3;
const ANGLE_FIELD: u32 = 4;

const POINT_X_FIELD: u32 = 1;
const POINT_Y_FIELD: u32 = 2;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;
const MAX_RECURSION: u32 = 64;
const BUFFA_LOGICAL_ALLOCATIONS_PER_TRAVERSAL: usize = 1;
const OUTPUT_ALLOCATIONS: usize = 1;
const MAX_MESSAGE_BYTES: usize = 512 * 1024 * 1024;

/// Finite limits for one direct `TSD.DrawableArchive` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl DecodeOptions {
    /// Construct an explicit source/field/work/nesting policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes: max_input_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_allocations: usize::MAX,
            max_retained_bytes: usize::MAX,
            max_scratch_bytes: usize::MAX,
        }
    }

    /// Construct a bounded policy from one source slice.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(16).max(1),
            bytes.saturating_mul(64).max(1),
            16,
        )
        .with_max_output_bytes(bytes.saturating_mul(2).max(1))
        .with_max_allocations(bytes.saturating_mul(16).max(16))
        .with_max_retained_bytes(bytes.saturating_mul(32).max(1))
        .with_max_scratch_bytes(
            bytes
                .saturating_mul(size_of::<ParsedField>())
                .saturating_mul(16)
                .max(1),
        )
    }

    /// Replace the source/input ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, maximum: usize) -> Self {
        self.max_input_bytes = maximum;
        self
    }

    /// Alias using the message-oriented name used by sibling codecs.
    #[must_use]
    pub const fn with_max_message_bytes(self, maximum: usize) -> Self {
        self.with_max_input_bytes(maximum)
    }

    /// Replace the candidate-output ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace field and work ceilings.
    #[must_use]
    pub const fn with_resource_limits(mut self, fields: usize, work_bytes: usize) -> Self {
        self.max_fields = fields;
        self.max_work_bytes = work_bytes;
        self
    }

    /// Replace the nesting ceiling.
    #[must_use]
    pub const fn with_recursion_limit(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self
    }

    /// Replace the logical allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Replace the retained-byte ceiling.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = maximum;
        self
    }

    /// Replace the parser scratch ceiling.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, maximum: usize) -> Self {
        self.max_scratch_bytes = maximum;
        self
    }

    /// Return the input ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Alias using the message-oriented name used by sibling codecs.
    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.max_input_bytes()
    }

    /// Return the output ceiling.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Return the field ceiling.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Return the work ceiling.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_input_bytes)
            // Unknown fields remain source-owned. Buffa skips them without
            // retaining a second unknown-field buffer under this setting.
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Resource axis rejected by a strict scan or prepared execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Input source bytes.
    InputBytes { observed: usize, maximum: usize },
    /// Candidate output bytes.
    OutputBytes { observed: usize, maximum: usize },
    /// Visited protobuf fields.
    Fields { observed: usize, maximum: usize },
    /// Wire work bytes.
    Work { observed: usize, maximum: usize },
    /// Nested message depth.
    Nesting { observed: u32, maximum: u32 },
    /// Logical staging allocations.
    Allocations { observed: usize, maximum: usize },
    /// Source/candidate bytes retained together.
    Retained { observed: usize, maximum: usize },
    /// Parser scratch bytes.
    Scratch { observed: usize, maximum: usize },
}

/// Strict shape-property codec failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Plain(&'static str),
    Limit(DecodeLimit),
    Allocation(&'static str),
    InvalidUtf8(&'static str),
}

impl DecodeError {
    const fn plain(message: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Plain(message),
        }
    }

    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limit(limit),
        }
    }

    const fn allocation(resource: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation(resource),
        }
    }

    const fn invalid_utf8(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidUtf8(field),
        }
    }

    /// Return the resource failure, when this is a bounded refusal.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Limit(limit) => Some(limit),
            _ => None,
        }
    }

    /// Alias used by focused package adapters.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<DecodeLimit> {
        self.resource_limit()
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.kind {
            DecodeErrorKind::Plain(message) => formatter.write_str(message),
            DecodeErrorKind::Allocation(resource) => {
                write!(
                    formatter,
                    "shape properties allocation failed for {resource}"
                )
            },
            DecodeErrorKind::InvalidUtf8(field) => write!(formatter, "invalid UTF-8 in {field}"),
            DecodeErrorKind::Limit(DecodeLimit::InputBytes { observed, maximum }) => {
                write!(
                    formatter,
                    "shape properties input bytes {observed} exceed {maximum}"
                )
            },
            DecodeErrorKind::Limit(DecodeLimit::OutputBytes { observed, maximum }) => {
                write!(
                    formatter,
                    "shape properties output bytes {observed} exceed {maximum}"
                )
            },
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => {
                write!(
                    formatter,
                    "shape properties fields {observed} exceed {maximum}"
                )
            },
            DecodeErrorKind::Limit(DecodeLimit::Work { observed, maximum }) => {
                write!(
                    formatter,
                    "shape properties work {observed} exceeds {maximum}"
                )
            },
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => {
                write!(
                    formatter,
                    "shape properties nesting {observed} exceeds {maximum}"
                )
            },
            DecodeErrorKind::Limit(DecodeLimit::Allocations { observed, maximum }) => {
                write!(
                    formatter,
                    "shape properties allocations {observed} exceed {maximum}"
                )
            },
            DecodeErrorKind::Limit(DecodeLimit::Retained { observed, maximum }) => {
                write!(
                    formatter,
                    "shape properties retained bytes {observed} exceed {maximum}"
                )
            },
            DecodeErrorKind::Limit(DecodeLimit::Scratch { observed, maximum }) => {
                write!(
                    formatter,
                    "shape properties scratch bytes {observed} exceed {maximum}"
                )
            },
        }
    }
}

impl std::error::Error for DecodeError {}

/// Neutral finite drawable point.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Point {
    x: f32,
    y: f32,
}

impl Point {
    /// Construct a point. Finiteness is checked when a write is prepared.
    #[must_use]
    pub const fn new(x: f32, y: f32) -> Self {
        Self { x, y }
    }

    /// Horizontal coordinate.
    #[must_use]
    pub const fn x(self) -> f32 {
        self.x
    }

    /// Vertical coordinate.
    #[must_use]
    pub const fn y(self) -> f32 {
        self.y
    }
}

/// Neutral drawable size in document points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Size {
    width: f32,
    height: f32,
}

impl Size {
    /// Construct a size. Finiteness and non-negativity are checked when a
    /// write is prepared.
    #[must_use]
    pub const fn new(width: f32, height: f32) -> Self {
        Self { width, height }
    }

    /// Width.
    #[must_use]
    pub const fn width(self) -> f32 {
        self.width
    }

    /// Height.
    #[must_use]
    pub const fn height(self) -> f32 {
        self.height
    }
}

/// Optional geometry carried by one drawable.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GeometrySnapshot {
    position: Option<Point>,
    size: Option<Size>,
    flags: Option<u32>,
    angle: Option<f32>,
}

impl GeometrySnapshot {
    /// Position, preserving field presence.
    #[must_use]
    pub const fn position(self) -> Option<Point> {
        self.position
    }

    /// Size, preserving field presence.
    #[must_use]
    pub const fn size(self) -> Option<Size> {
        self.size
    }

    /// Native reflection/transform flags.
    #[must_use]
    pub const fn flags(self) -> Option<u32> {
        self.flags
    }

    /// Native rotation in degrees.
    #[must_use]
    pub const fn angle(self) -> Option<f32> {
        self.angle
    }
}

/// Borrowed shared properties of one drawable.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DrawablePropertiesSnapshot<'source> {
    hyperlink_url: Option<&'source str>,
    locked: Option<bool>,
    aspect_ratio_locked: Option<bool>,
    accessibility_description: Option<&'source str>,
}

impl<'source> DrawablePropertiesSnapshot<'source> {
    /// Explicit hyperlink URL, if present.
    #[must_use]
    pub const fn hyperlink_url(self) -> Option<&'source str> {
        self.hyperlink_url
    }

    /// Explicit lock state, if present.
    #[must_use]
    pub const fn locked(self) -> Option<bool> {
        self.locked
    }

    /// Explicit aspect-ratio lock state, if present.
    #[must_use]
    pub const fn aspect_ratio_locked(self) -> Option<bool> {
        self.aspect_ratio_locked
    }

    /// Explicit accessibility description, if present.
    #[must_use]
    pub const fn accessibility_description(self) -> Option<&'source str> {
        self.accessibility_description
    }
}

/// Semantic projection of one direct drawable archive.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct DrawableSnapshot<'source> {
    geometry: Option<GeometrySnapshot>,
    properties: DrawablePropertiesSnapshot<'source>,
}

impl<'source> DrawableSnapshot<'source> {
    /// Return the optional geometry envelope.
    #[must_use]
    pub const fn geometry(self) -> Option<GeometrySnapshot> {
        self.geometry
    }

    /// Return the shared drawable properties.
    #[must_use]
    pub const fn properties(self) -> DrawablePropertiesSnapshot<'source> {
        self.properties
    }
}

/// One source-preserving optional-field update.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldUpdate<T> {
    /// Retain the complete source field span.
    Preserve,
    /// Replace or append the canonical field.
    Set(T),
    /// Remove the canonical field if present.
    Clear,
}

impl<T> FieldUpdate<T> {
    /// Preserve the source field.
    #[must_use]
    pub const fn preserve() -> Self {
        Self::Preserve
    }

    /// Set a value.
    #[must_use]
    pub const fn set(value: T) -> Self {
        Self::Set(value)
    }

    /// Clear a field.
    #[must_use]
    pub const fn clear() -> Self {
        Self::Clear
    }
}

/// Requested geometry changes. The default preserves every source field.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GeometryWrite {
    position: FieldUpdate<Point>,
    size: FieldUpdate<Size>,
    flags: FieldUpdate<u32>,
    angle: FieldUpdate<f32>,
    remove_envelope: bool,
}

impl Default for GeometryWrite {
    fn default() -> Self {
        Self::preserve()
    }
}

impl GeometryWrite {
    /// Preserve every geometry field.
    #[must_use]
    pub const fn preserve() -> Self {
        Self {
            position: FieldUpdate::Preserve,
            size: FieldUpdate::Preserve,
            flags: FieldUpdate::Preserve,
            angle: FieldUpdate::Preserve,
            remove_envelope: false,
        }
    }

    /// Remove the outer geometry envelope without adding any canonical field.
    ///
    /// This is useful when a semantic edit must restore a source that had no
    /// `GeometryArchive` field. Unknown fields in the envelope are removed as
    /// part of the envelope; callers should use this only when the target
    /// source state proves that the envelope was absent.
    #[must_use]
    pub const fn remove_envelope(mut self) -> Self {
        self.remove_envelope = true;
        self
    }

    /// Construct an exact replacement. `None` clears each selected field.
    #[must_use]
    pub fn new(
        position: Option<Point>,
        size: Option<Size>,
        flags: Option<u32>,
        angle: Option<f32>,
    ) -> Self {
        Self::preserve()
            .with_position(position)
            .with_size(size)
            .with_flags(flags)
            .with_angle(angle)
    }

    /// Construct a replacement from an optional existing geometry envelope.
    #[must_use]
    pub fn from_snapshot(snapshot: Option<GeometrySnapshot>) -> Self {
        match snapshot {
            Some(snapshot) => Self::new(
                snapshot.position,
                snapshot.size,
                snapshot.flags,
                snapshot.angle,
            ),
            None => Self::new(None, None, None, None),
        }
    }

    /// Set or clear position.
    #[must_use]
    pub fn with_position(mut self, value: Option<Point>) -> Self {
        self.position = option_update(value);
        self
    }

    /// Set or clear size.
    #[must_use]
    pub fn with_size(mut self, value: Option<Size>) -> Self {
        self.size = option_update(value);
        self
    }

    /// Set or clear flags.
    #[must_use]
    pub fn with_flags(mut self, value: Option<u32>) -> Self {
        self.flags = option_update(value);
        self
    }

    /// Set or clear angle.
    #[must_use]
    pub fn with_angle(mut self, value: Option<f32>) -> Self {
        self.angle = option_update(value);
        self
    }

    /// Set an explicit position.
    #[must_use]
    pub const fn set_position(mut self, value: Point) -> Self {
        self.position = FieldUpdate::Set(value);
        self
    }

    /// Clear position.
    #[must_use]
    pub const fn clear_position(mut self) -> Self {
        self.position = FieldUpdate::Clear;
        self
    }

    /// Set an explicit size.
    #[must_use]
    pub const fn set_size(mut self, value: Size) -> Self {
        self.size = FieldUpdate::Set(value);
        self
    }

    /// Clear size.
    #[must_use]
    pub const fn clear_size(mut self) -> Self {
        self.size = FieldUpdate::Clear;
        self
    }

    /// Set explicit flags.
    #[must_use]
    pub const fn set_flags(mut self, value: u32) -> Self {
        self.flags = FieldUpdate::Set(value);
        self
    }

    /// Clear flags.
    #[must_use]
    pub const fn clear_flags(mut self) -> Self {
        self.flags = FieldUpdate::Clear;
        self
    }

    /// Set explicit rotation.
    #[must_use]
    pub const fn set_angle(mut self, value: f32) -> Self {
        self.angle = FieldUpdate::Set(value);
        self
    }

    /// Clear rotation.
    #[must_use]
    pub const fn clear_angle(mut self) -> Self {
        self.angle = FieldUpdate::Clear;
        self
    }
}

/// Requested shared drawable-property changes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DrawablePropertiesWrite<'value> {
    hyperlink_url: FieldUpdate<&'value str>,
    locked: FieldUpdate<bool>,
    aspect_ratio_locked: FieldUpdate<bool>,
    accessibility_description: FieldUpdate<&'value str>,
}

impl<'value> DrawablePropertiesWrite<'value> {
    /// Preserve every source property.
    #[must_use]
    pub const fn preserve() -> Self {
        Self {
            hyperlink_url: FieldUpdate::Preserve,
            locked: FieldUpdate::Preserve,
            aspect_ratio_locked: FieldUpdate::Preserve,
            accessibility_description: FieldUpdate::Preserve,
        }
    }

    /// Construct an exact replacement. `None` clears a property.
    #[must_use]
    pub fn new(
        hyperlink_url: Option<&'value str>,
        locked: Option<bool>,
        aspect_ratio_locked: Option<bool>,
        accessibility_description: Option<&'value str>,
    ) -> Self {
        Self::preserve()
            .with_hyperlink_url(hyperlink_url)
            .with_locked(locked)
            .with_aspect_ratio_locked(aspect_ratio_locked)
            .with_accessibility_description(accessibility_description)
    }

    /// Construct an exact replacement from a borrowed snapshot.
    #[must_use]
    pub fn from_snapshot(snapshot: DrawablePropertiesSnapshot<'value>) -> Self {
        Self::new(
            snapshot.hyperlink_url,
            snapshot.locked,
            snapshot.aspect_ratio_locked,
            snapshot.accessibility_description,
        )
    }

    /// Set or clear the hyperlink.
    #[must_use]
    pub fn with_hyperlink_url(mut self, value: Option<&'value str>) -> Self {
        self.hyperlink_url = option_update(value);
        self
    }

    /// Alias using the shorter native property name.
    #[must_use]
    pub fn with_hyperlink(self, value: Option<&'value str>) -> Self {
        self.with_hyperlink_url(value)
    }

    /// Set or clear lock state.
    #[must_use]
    pub fn with_locked(mut self, value: Option<bool>) -> Self {
        self.locked = option_update(value);
        self
    }

    /// Set or clear aspect-ratio lock state.
    #[must_use]
    pub fn with_aspect_ratio_locked(mut self, value: Option<bool>) -> Self {
        self.aspect_ratio_locked = option_update(value);
        self
    }

    /// Set or clear accessibility text.
    #[must_use]
    pub fn with_accessibility_description(mut self, value: Option<&'value str>) -> Self {
        self.accessibility_description = option_update(value);
        self
    }

    /// Set an explicit hyperlink.
    #[must_use]
    pub const fn set_hyperlink_url(mut self, value: &'value str) -> Self {
        self.hyperlink_url = FieldUpdate::Set(value);
        self
    }

    /// Alias using the shorter native property name.
    #[must_use]
    pub const fn set_hyperlink(self, value: &'value str) -> Self {
        self.set_hyperlink_url(value)
    }

    /// Clear the hyperlink.
    #[must_use]
    pub const fn clear_hyperlink_url(mut self) -> Self {
        self.hyperlink_url = FieldUpdate::Clear;
        self
    }

    /// Set an explicit lock state.
    #[must_use]
    pub const fn set_locked(mut self, value: bool) -> Self {
        self.locked = FieldUpdate::Set(value);
        self
    }

    /// Clear lock state.
    #[must_use]
    pub const fn clear_locked(mut self) -> Self {
        self.locked = FieldUpdate::Clear;
        self
    }

    /// Set an explicit aspect-ratio lock state.
    #[must_use]
    pub const fn set_aspect_ratio_locked(mut self, value: bool) -> Self {
        self.aspect_ratio_locked = FieldUpdate::Set(value);
        self
    }

    /// Clear aspect-ratio lock state.
    #[must_use]
    pub const fn clear_aspect_ratio_locked(mut self) -> Self {
        self.aspect_ratio_locked = FieldUpdate::Clear;
        self
    }

    /// Set an explicit accessibility description.
    #[must_use]
    pub const fn set_accessibility_description(mut self, value: &'value str) -> Self {
        self.accessibility_description = FieldUpdate::Set(value);
        self
    }

    /// Clear accessibility text.
    #[must_use]
    pub const fn clear_accessibility_description(mut self) -> Self {
        self.accessibility_description = FieldUpdate::Clear;
        self
    }
}

fn option_update<T>(value: Option<T>) -> FieldUpdate<T> {
    match value {
        Some(value) => FieldUpdate::Set(value),
        None => FieldUpdate::Clear,
    }
}

/// Exact successful scan accounting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    /// Source bytes visited.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Alias for [`Self::input_bytes`].
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.input_bytes
    }

    /// Visited fields.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Charged wire work.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nested message depth.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Logical allocations admitted by the scan.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source bytes retained by the report's operation.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Parser scratch bytes reserved by the scan.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Exact resource requirements for a prepared rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    /// Candidate output bytes.
    pub output_bytes: usize,
    /// Aggregate visited fields.
    pub fields: usize,
    /// Aggregate wire work.
    pub work_bytes: usize,
    /// Maximum nested depth.
    pub max_depth: u32,
    /// Logical allocations.
    pub allocations: usize,
    /// Source/candidate retained bytes.
    pub retained_bytes: usize,
    /// Scratch bytes.
    pub scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    /// Use exact requirements as execution ceilings.
    #[must_use]
    pub const fn exact(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits {
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }

    /// Alias for [`Self::exact`].
    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }
}

/// Caller-supplied ceilings replayed before output allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    /// Candidate output ceiling.
    pub output_bytes: usize,
    /// Field ceiling.
    pub fields: usize,
    /// Work ceiling.
    pub work_bytes: usize,
    /// Maximum depth.
    pub max_depth: u32,
    /// Allocation ceiling.
    pub allocations: usize,
    /// Retained-byte ceiling.
    pub retained_bytes: usize,
    /// Scratch-byte ceiling.
    pub scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Use exact prepared requirements as execution ceilings.
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        requirements.exact()
    }

    /// Replace the output ceiling.
    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }

    /// Replace the field ceiling.
    #[must_use]
    pub const fn with_fields(mut self, value: usize) -> Self {
        self.fields = value;
        self
    }

    /// Replace the work ceiling.
    #[must_use]
    pub const fn with_work_bytes(mut self, value: usize) -> Self {
        self.work_bytes = value;
        self
    }

    /// Replace the maximum-depth ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }

    /// Replace the allocation ceiling.
    #[must_use]
    pub const fn with_allocations(mut self, value: usize) -> Self {
        self.allocations = value;
        self
    }

    /// Replace the retained-byte ceiling.
    #[must_use]
    pub const fn with_retained_bytes(mut self, value: usize) -> Self {
        self.retained_bytes = value;
        self
    }

    /// Replace the scratch-byte ceiling.
    #[must_use]
    pub const fn with_scratch_bytes(mut self, value: usize) -> Self {
        self.scratch_bytes = value;
        self
    }
}

/// Exact successful rewrite accounting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    input_bytes: usize,
    requirements: RewriteExecutionRequirements,
    changed: bool,
}

impl RewriteReport {
    /// Source bytes retained for the rewrite.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Alias for [`Self::input_bytes`].
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.input_bytes
    }

    /// Return candidate output bytes.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.requirements.output_bytes
    }

    /// Return the aggregate field count.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.requirements.fields
    }

    /// Return aggregate work bytes.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.requirements.work_bytes
    }

    /// Return the maximum nested message depth.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.requirements.max_depth
    }

    /// Return logical allocations admitted by the rewrite.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.requirements.allocations
    }

    /// Return source/candidate retained bytes.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.requirements.retained_bytes
    }

    /// Return parser/output scratch bytes.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.requirements.scratch_bytes
    }

    /// Return whether the semantic rewrite changed any selected field.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the full exact requirements.
    #[must_use]
    pub const fn requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }
}

/// Candidate bytes plus exact rewrite accounting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RewriteOutput {
    bytes: Vec<u8>,
    report: RewriteReport,
}

impl RewriteOutput {
    /// Borrow candidate bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Alias for [`Self::bytes`].
    #[must_use]
    pub fn output(&self) -> &[u8] {
        self.bytes()
    }

    /// Alias for [`Self::bytes`].
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.bytes()
    }

    /// Consume the output and return candidate bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// Alias for [`Self::into_bytes`].
    #[must_use]
    pub fn into_output(self) -> Vec<u8> {
        self.into_bytes()
    }

    /// Return rewrite accounting.
    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }

    /// Return whether the candidate differs semantically from the source.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.report.changed
    }
}

/// Prepared rewrite retaining only borrowed source spans and typed requests.
#[derive(Debug)]
pub struct PreparedShapePropertiesRewrite<'source, 'value> {
    source: &'source [u8],
    geometry: GeometryWrite,
    properties: DrawablePropertiesWrite<'value>,
    scan: Scan<'source>,
    requirements: RewriteExecutionRequirements,
    changed: bool,
}

impl<'source, 'value> PreparedShapePropertiesRewrite<'source, 'value> {
    /// Return strict source accounting captured during preparation.
    #[must_use]
    pub fn prepare_report(&self) -> DecodeReport {
        self.scan.report(self.source.len())
    }

    /// Return exact resources needed by execution.
    #[must_use]
    pub const fn requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Alias matching sibling prepared-codec APIs.
    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements()
    }

    /// Execute after replaying all requirements against one aggregate ledger.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_execution_limits(self.requirements, limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| DecodeError::allocation("rewrite output"))?;
        build_output(
            self.source,
            &self.scan,
            self.geometry,
            self.properties,
            &mut output,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::plain(
                "shape properties rewrite output measurement disagreed with execution",
            ));
        }
        Ok(RewriteOutput {
            bytes: output,
            report: RewriteReport {
                input_bytes: self.source.len(),
                requirements: self.requirements,
                changed: self.changed,
            },
        })
    }
}

/// Decode one direct drawable archive.
pub fn decode_drawable_archive(
    source: &[u8],
    options: DecodeOptions,
) -> Result<DrawableSnapshot<'_>, DecodeError> {
    Ok(decode_drawable_archive_with_report(source, options)?.0)
}

/// Decode a drawable and return exact bounded accounting.
pub fn decode_drawable_archive_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(DrawableSnapshot<'_>, DecodeReport), DecodeError> {
    let scan = scan_drawable(source, options)?;
    Ok((scan.snapshot, scan.report(source.len())))
}

/// Prepare a source-preserving direct drawable rewrite.
pub fn prepare_drawable_archive_rewrite<'source, 'value>(
    source: &'source [u8],
    geometry: GeometryWrite,
    properties: DrawablePropertiesWrite<'value>,
    options: DecodeOptions,
) -> Result<PreparedShapePropertiesRewrite<'source, 'value>, DecodeError> {
    validate_input(source, options)?;
    validate_geometry_write(geometry)?;
    validate_properties_write(properties)?;
    let scan = scan_drawable(source, options)?;
    let output_bytes = measure_output(source, &scan, geometry, properties, options)?;
    let changed = semantic_change(&scan, geometry, properties);
    let requirements = requirements_for(source, &scan, output_bytes, options)?;
    Ok(PreparedShapePropertiesRewrite {
        source,
        geometry,
        properties,
        scan,
        requirements,
        changed,
    })
}

/// Rewrite one direct drawable archive with exact source/output admission.
pub fn rewrite_drawable_archive(
    source: &[u8],
    geometry: GeometryWrite,
    properties: DrawablePropertiesWrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_drawable_archive_rewrite(source, geometry, properties, options)?;
    let limits = prepared.requirements().exact();
    prepared.execute(limits)
}

#[derive(Debug, Clone, Copy)]
struct ParsedField {
    number: u32,
    wire: u8,
    key_canonical: bool,
    start: usize,
    value_start: usize,
    value_end: usize,
    end: usize,
    varint: Option<u64>,
    varint_canonical: Option<bool>,
    length_canonical: Option<bool>,
}

#[derive(Debug)]
struct NestedFixed {
    fields: Vec<ParsedField>,
    first: f32,
    second: f32,
}

#[derive(Debug)]
struct GeometryScan {
    fields: Vec<ParsedField>,
    position: Option<NestedFixed>,
    size: Option<NestedFixed>,
    snapshot: GeometrySnapshot,
}

#[derive(Debug)]
struct Scan<'source> {
    fields: Vec<ParsedField>,
    geometry_field: Option<ParsedField>,
    geometry: Option<GeometryScan>,
    snapshot: DrawableSnapshot<'source>,
    accounting: Accounting,
}

#[derive(Debug, Clone, Copy)]
struct Accounting {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    scratch_bytes: usize,
}

#[derive(Debug, Clone, Copy)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    scratch_bytes: usize,
}

impl Budget {
    const fn new() -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            allocations: 0,
            scratch_bytes: 0,
        }
    }

    fn message(&mut self, source_bytes: usize, options: DecodeOptions) -> Result<(), DecodeError> {
        self.work(source_bytes, options)?;
        Ok(())
    }

    fn field(&mut self, span: usize, options: DecodeOptions) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(1).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
        if self.fields > options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: self.fields,
                maximum: options.max_fields,
            }));
        }
        self.work(span, options)
    }

    fn work(&mut self, amount: usize, options: DecodeOptions) -> Result<(), DecodeError> {
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
        if self.work_bytes > options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn reserve_fields(
        &mut self,
        capacity: usize,
        options: DecodeOptions,
    ) -> Result<Vec<ParsedField>, DecodeError> {
        let scratch = capacity
            .checked_mul(size_of::<ParsedField>())
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Scratch {
                    observed: usize::MAX,
                    maximum: options.max_scratch_bytes,
                })
            })?;
        self.scratch_bytes = self.scratch_bytes.checked_add(scratch).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
        if self.scratch_bytes > options.max_scratch_bytes {
            return Err(DecodeError::limit(DecodeLimit::Scratch {
                observed: self.scratch_bytes,
                maximum: options.max_scratch_bytes,
            }));
        }
        self.allocations = self.allocations.checked_add(1).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
        if self.allocations > options.max_allocations {
            return Err(DecodeError::limit(DecodeLimit::Allocations {
                observed: self.allocations,
                maximum: options.max_allocations,
            }));
        }
        let mut fields = Vec::new();
        fields
            .try_reserve_exact(capacity)
            .map_err(|_| DecodeError::allocation("parsed field list"))?;
        Ok(fields)
    }

    fn buffa_traversal(
        &mut self,
        fields: usize,
        work_bytes: usize,
        scratch_fields: usize,
        options: DecodeOptions,
    ) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(fields).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
        if self.fields > options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: self.fields,
                maximum: options.max_fields,
            }));
        }
        self.work_bytes = self.work_bytes.checked_add(work_bytes).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
        if self.work_bytes > options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: options.max_work_bytes,
            }));
        }
        let scratch = scratch_fields
            .checked_mul(size_of::<ParsedField>())
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Scratch {
                    observed: usize::MAX,
                    maximum: options.max_scratch_bytes,
                })
            })?;
        self.scratch_bytes = self.scratch_bytes.checked_add(scratch).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
        if self.scratch_bytes > options.max_scratch_bytes {
            return Err(DecodeError::limit(DecodeLimit::Scratch {
                observed: self.scratch_bytes,
                maximum: options.max_scratch_bytes,
            }));
        }
        self.allocations = self
            .allocations
            .checked_add(BUFFA_LOGICAL_ALLOCATIONS_PER_TRAVERSAL)
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Allocations {
                    observed: usize::MAX,
                    maximum: options.max_allocations,
                })
            })?;
        if self.allocations > options.max_allocations {
            return Err(DecodeError::limit(DecodeLimit::Allocations {
                observed: self.allocations,
                maximum: options.max_allocations,
            }));
        }
        Ok(())
    }

    fn depth(&mut self, depth: u32, options: DecodeOptions) -> Result<(), DecodeError> {
        self.max_depth = self.max_depth.max(depth);
        if depth > options.recursion_limit {
            return Err(DecodeError::limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: options.recursion_limit,
            }));
        }
        Ok(())
    }

    const fn accounting(self) -> Accounting {
        Accounting {
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            scratch_bytes: self.scratch_bytes,
        }
    }
}

impl<'source> Scan<'source> {
    fn report(&self, source_bytes: usize) -> DecodeReport {
        DecodeReport {
            input_bytes: source_bytes,
            fields: self.accounting.fields,
            work_bytes: self.accounting.work_bytes,
            max_depth: self.accounting.max_depth,
            allocations: self.accounting.allocations,
            retained_bytes: source_bytes,
            scratch_bytes: self.accounting.scratch_bytes,
        }
    }
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let maximum = options.max_input_bytes.min(MAX_MESSAGE_BYTES);
    if source.len() > maximum || options.max_input_bytes > MAX_MESSAGE_BYTES {
        return Err(DecodeError::limit(DecodeLimit::InputBytes {
            observed: source.len().max(options.max_input_bytes),
            maximum,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}

fn scan_drawable<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<Scan<'source>, DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new();
    let fields = parse_message(source, options, &mut budget, 1)?;
    let geometry_field = optional_known(&fields, GEOMETRY_FIELD, "DrawableArchive.geometry")?;
    let geometry = if let Some(field) = geometry_field {
        require_wire(field, 2, "DrawableArchive.geometry")?;
        let payload = payload(source, field)?;
        Some(parse_geometry(payload, options, &mut budget, 2)?)
    } else {
        None
    };

    let hyperlink_field =
        optional_known(&fields, HYPERLINK_FIELD, "DrawableArchive.hyperlink_url")?;
    let hyperlink_url = hyperlink_field
        .map(|field| {
            require_wire(field, 2, "DrawableArchive.hyperlink_url")?;
            let bytes = payload(source, field)?;
            str::from_utf8(bytes).map_err(|_| DecodeError::invalid_utf8("hyperlink_url"))
        })
        .transpose()?;
    let locked_field = optional_known(&fields, LOCKED_FIELD, "DrawableArchive.locked")?;
    let locked = locked_field
        .map(|field| {
            require_wire(field, 0, "DrawableArchive.locked")?;
            canonical_bool(field)
        })
        .transpose()?;
    let aspect_field = optional_known(
        &fields,
        ASPECT_RATIO_LOCKED_FIELD,
        "DrawableArchive.aspect_ratio_locked",
    )?;
    let aspect_ratio_locked = aspect_field
        .map(|field| {
            require_wire(field, 0, "DrawableArchive.aspect_ratio_locked")?;
            canonical_bool(field)
        })
        .transpose()?;
    let accessibility_field = optional_known(
        &fields,
        ACCESSIBILITY_DESCRIPTION_FIELD,
        "DrawableArchive.accessibility_description",
    )?;
    let accessibility_description = accessibility_field
        .map(|field| {
            require_wire(field, 2, "DrawableArchive.accessibility_description")?;
            let bytes = payload(source, field)?;
            str::from_utf8(bytes)
                .map_err(|_| DecodeError::invalid_utf8("accessibility_description"))
        })
        .transpose()?;

    force_buffa(
        source,
        options,
        &fields,
        geometry.as_ref(),
        hyperlink_url,
        locked,
        aspect_ratio_locked,
        accessibility_description,
        &mut budget,
    )?;
    let properties = DrawablePropertiesSnapshot {
        hyperlink_url,
        locked,
        aspect_ratio_locked,
        accessibility_description,
    };
    let snapshot = DrawableSnapshot {
        geometry: geometry.as_ref().map(|value| value.snapshot),
        properties,
    };
    Ok(Scan {
        fields,
        geometry_field,
        geometry,
        snapshot,
        accounting: budget.accounting(),
    })
}

fn parse_geometry(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<GeometryScan, DecodeError> {
    let fields = parse_message(source, options, budget, depth)?;
    let position_field = optional_known(&fields, POSITION_FIELD, "Geometry.position")?;
    let size_field = optional_known(&fields, SIZE_FIELD, "Geometry.size")?;
    let flags_field = optional_known(&fields, FLAGS_FIELD, "Geometry.flags")?;
    let angle_field = optional_known(&fields, ANGLE_FIELD, "Geometry.angle")?;

    let position = position_field
        .map(|field| {
            require_wire(field, 2, "Geometry.position")?;
            parse_fixed_pair(
                payload(source, field)?,
                options,
                budget,
                depth + 1,
                "Point",
                true,
            )
        })
        .transpose()?;
    let size = size_field
        .map(|field| {
            require_wire(field, 2, "Geometry.size")?;
            parse_fixed_pair(
                payload(source, field)?,
                options,
                budget,
                depth + 1,
                "Size",
                false,
            )
        })
        .transpose()?;
    let flags = flags_field
        .map(|field| {
            require_wire(field, 0, "Geometry.flags")?;
            canonical_u32(field, "Geometry.flags")
        })
        .transpose()?;
    let angle = angle_field
        .map(|field| {
            require_wire(field, 5, "Geometry.angle")?;
            let value = fixed32(source, field, "Geometry.angle")?;
            if !value.is_finite() {
                return Err(DecodeError::plain("Geometry.angle must be finite"));
            }
            Ok(value)
        })
        .transpose()?;
    let snapshot = GeometrySnapshot {
        position: position
            .as_ref()
            .map(|value| Point::new(value.first, value.second)),
        size: size
            .as_ref()
            .map(|value| Size::new(value.first, value.second)),
        flags,
        angle,
    };
    Ok(GeometryScan {
        fields,
        position,
        size,
        snapshot,
    })
}

fn parse_fixed_pair(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
    name: &'static str,
    point: bool,
) -> Result<NestedFixed, DecodeError> {
    let fields = parse_message(source, options, budget, depth)?;
    let first_number = if point {
        POINT_X_FIELD
    } else {
        SIZE_WIDTH_FIELD
    };
    let second_number = if point {
        POINT_Y_FIELD
    } else {
        SIZE_HEIGHT_FIELD
    };
    let first = unique_known(
        &fields,
        first_number,
        if point { "Point.x" } else { "Size.width" },
    )?;
    let second = unique_known(
        &fields,
        second_number,
        if point { "Point.y" } else { "Size.height" },
    )?;
    require_wire(first, 5, name)?;
    require_wire(second, 5, name)?;
    let first = fixed32(source, first, name)?;
    let second = fixed32(source, second, name)?;
    if !first.is_finite() || !second.is_finite() {
        return Err(DecodeError::plain(if point {
            "Point coordinates must be finite"
        } else {
            "Size dimensions must be finite"
        }));
    }
    if !point && (first < 0.0 || second < 0.0) {
        return Err(DecodeError::plain("Size dimensions must be non-negative"));
    }
    Ok(NestedFixed {
        fields,
        first,
        second,
    })
}

fn parsed_fields_bytes(fields: &[ParsedField]) -> usize {
    fields.iter().fold(0usize, |total, field| {
        total.saturating_add(field.end.saturating_sub(field.start))
    })
}

fn force_buffa(
    source: &[u8],
    options: DecodeOptions,
    fields: &[ParsedField],
    geometry: Option<&GeometryScan>,
    hyperlink_url: Option<&str>,
    locked: Option<bool>,
    aspect_ratio_locked: Option<bool>,
    accessibility_description: Option<&str>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let top_fields = fields.len();
    budget.buffa_traversal(top_fields, source.len(), top_fields, options)?;
    let creation: geometry_projection::DrawableArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::plain("Buffa drawable geometry projection rejected source"))?;
    let geometry_fields = geometry.map_or(0, |value| value.fields.len());
    let geometry_work = geometry.map_or(0, |value| parsed_fields_bytes(&value.fields));
    budget.buffa_traversal(geometry_fields, geometry_work, geometry_fields, options)?;
    let projected_geometry = creation
        .geometry
        .get()
        .map_err(|_| DecodeError::plain("Buffa drawable geometry projection rejected geometry"))?;
    match (geometry, projected_geometry) {
        (None, None) => {},
        (Some(expected), Some(projected)) => {
            let position_fields = expected
                .position
                .as_ref()
                .map_or(0, |value| value.fields.len());
            let position_work = expected
                .position
                .as_ref()
                .map_or(0, |value| parsed_fields_bytes(&value.fields));
            budget.buffa_traversal(position_fields, position_work, position_fields, options)?;
            let projected_position = projected
                .position
                .get()
                .map_err(|_| DecodeError::plain("Buffa point projection rejected source"))?
                .map(|point| Point::new(point.x, point.y));
            let size_fields = expected.size.as_ref().map_or(0, |value| value.fields.len());
            let size_work = expected
                .size
                .as_ref()
                .map_or(0, |value| parsed_fields_bytes(&value.fields));
            budget.buffa_traversal(size_fields, size_work, size_fields, options)?;
            let projected_size = projected
                .size
                .get()
                .map_err(|_| DecodeError::plain("Buffa size projection rejected source"))?
                .map(|size| Size::new(size.width, size.height));
            let candidate = GeometrySnapshot {
                position: projected_position,
                size: projected_size,
                flags: projected.flags,
                angle: projected.angle,
            };
            if candidate != expected.snapshot {
                return Err(DecodeError::plain(
                    "Buffa drawable geometry projection disagrees with strict scan",
                ));
            }
        },
        _ => {
            return Err(DecodeError::plain(
                "Buffa drawable geometry projection disagrees with strict presence",
            ));
        },
    }

    budget.buffa_traversal(top_fields, source.len(), top_fields, options)?;
    let properties: properties_projection::DrawableArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::plain("Buffa drawable property projection rejected source"))?;
    if properties.hyperlink != hyperlink_url
        || properties.locked != locked
        || properties.aspect_ratio_locked != aspect_ratio_locked
        || properties.accessibility_description != accessibility_description
    {
        return Err(DecodeError::plain(
            "Buffa drawable property projection disagrees with strict scan",
        ));
    }
    Ok(())
}

fn unique_known(
    fields: &[ParsedField],
    number: u32,
    name: &'static str,
) -> Result<ParsedField, DecodeError> {
    let mut found = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number == number)
    {
        if found.is_some() {
            return Err(DecodeError::plain(name));
        }
        found = Some(field);
    }
    found.ok_or_else(|| DecodeError::plain(name))
}

fn optional_known(
    fields: &[ParsedField],
    number: u32,
    name: &'static str,
) -> Result<Option<ParsedField>, DecodeError> {
    let mut found = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number == number)
    {
        if found.is_some() {
            return Err(DecodeError::plain(name));
        }
        found = Some(field);
    }
    Ok(found)
}

fn require_wire(field: ParsedField, expected: u8, name: &'static str) -> Result<(), DecodeError> {
    if !field.key_canonical {
        return Err(DecodeError::plain(
            "protobuf known field key is not canonical",
        ));
    }
    if field.wire != expected {
        return Err(DecodeError::plain(name));
    }
    Ok(())
}

fn payload(source: &[u8], field: ParsedField) -> Result<&[u8], DecodeError> {
    if field.length_canonical != Some(true) {
        return Err(DecodeError::plain("protobuf known length is not canonical"));
    }
    source
        .get(field.value_start..field.value_end)
        .ok_or_else(|| DecodeError::plain("protobuf field payload is outside source"))
}

fn fixed32(source: &[u8], field: ParsedField, name: &'static str) -> Result<f32, DecodeError> {
    let bytes = source
        .get(field.value_start..field.value_end)
        .ok_or_else(|| DecodeError::plain(name))?;
    let bytes: [u8; 4] = bytes.try_into().map_err(|_| DecodeError::plain(name))?;
    Ok(f32::from_le_bytes(bytes))
}

fn canonical_u32(field: ParsedField, name: &'static str) -> Result<u32, DecodeError> {
    if field.varint_canonical != Some(true) {
        return Err(DecodeError::plain(name));
    }
    let value = field.varint.ok_or_else(|| DecodeError::plain(name))?;
    if value > u64::from(u32::MAX) {
        return Err(DecodeError::plain(name));
    }
    Ok(value as u32)
}

fn canonical_bool(field: ParsedField) -> Result<bool, DecodeError> {
    if field.varint_canonical != Some(true) {
        return Err(DecodeError::plain("protobuf boolean is not canonical"));
    }
    match field.varint {
        Some(0) => Ok(false),
        Some(1) => Ok(true),
        _ => Err(DecodeError::plain("protobuf boolean is not canonical")),
    }
}

fn parse_message(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<Vec<ParsedField>, DecodeError> {
    budget.depth(depth, options)?;
    budget.message(source.len(), options)?;
    let capacity = source.len().min(options.max_fields);
    let mut fields = budget.reserve_fields(capacity, options)?;
    let mut cursor = 0usize;
    while cursor < source.len() {
        let start = cursor;
        let (key, key_end, key_canonical) = read_varint(source, cursor)?;
        cursor = key_end;
        let number_value = key >> 3;
        let wire = (key & 7) as u8;
        if number_value == 0 || number_value > u64::from(0x1fff_ffff_u32) {
            return Err(DecodeError::plain("protobuf field number is invalid"));
        }
        let number = number_value as u32;
        if wire == 4 {
            return Err(DecodeError::plain("unexpected protobuf end-group"));
        }
        let (value_start, value_end, varint, varint_canonical, length_canonical) = match wire {
            0 => {
                let value_start = cursor;
                let (value, value_end, canonical) = read_varint(source, cursor)?;
                (value_start, value_end, Some(value), Some(canonical), None)
            },
            1 => {
                let value_end = cursor
                    .checked_add(8)
                    .ok_or_else(|| DecodeError::plain("protobuf fixed64 field length overflow"))?;
                if value_end > source.len() {
                    return Err(DecodeError::plain("truncated protobuf fixed64 field"));
                }
                (cursor, value_end, None, None, None)
            },
            2 => {
                let (length, payload_start, canonical) = read_length(source, cursor)?;
                let length = usize::try_from(length)
                    .map_err(|_| DecodeError::plain("protobuf length does not fit usize"))?;
                let value_end = payload_start.checked_add(length).ok_or_else(|| {
                    DecodeError::plain("protobuf length-delimited field overflow")
                })?;
                if value_end > source.len() {
                    return Err(DecodeError::plain(
                        "truncated protobuf length-delimited field",
                    ));
                }
                (payload_start, value_end, None, None, Some(canonical))
            },
            3 => {
                let mut group_cursor = cursor;
                skip_group(
                    source,
                    &mut group_cursor,
                    options,
                    budget,
                    depth + 1,
                    number,
                )?;
                (cursor, group_cursor, None, None, None)
            },
            5 => {
                let value_end = cursor
                    .checked_add(4)
                    .ok_or_else(|| DecodeError::plain("protobuf fixed32 field length overflow"))?;
                if value_end > source.len() {
                    return Err(DecodeError::plain("truncated protobuf fixed32 field"));
                }
                (cursor, value_end, None, None, None)
            },
            _ => return Err(DecodeError::plain("protobuf wire type is invalid")),
        };
        cursor = value_end;
        budget.field(cursor.saturating_sub(start), options)?;
        fields.push(ParsedField {
            number,
            wire,
            key_canonical,
            start,
            value_start,
            value_end,
            end: cursor,
            varint,
            varint_canonical,
            length_canonical,
        });
    }
    Ok(fields)
}

fn skip_group(
    source: &[u8],
    cursor: &mut usize,
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
    opener: u32,
) -> Result<(), DecodeError> {
    budget.depth(depth, options)?;
    loop {
        if *cursor >= source.len() {
            return Err(DecodeError::plain("unterminated protobuf group"));
        }
        let start = *cursor;
        let (key, key_end, _) = read_varint(source, *cursor)?;
        *cursor = key_end;
        let number_value = key >> 3;
        let wire = (key & 7) as u8;
        if number_value == 0 || number_value > u64::from(0x1fff_ffff_u32) {
            return Err(DecodeError::plain("protobuf group field number is invalid"));
        }
        let number = number_value as u32;
        match wire {
            0 => {
                let (_, end, _) = read_varint(source, *cursor)?;
                *cursor = end;
            },
            1 => {
                *cursor = cursor
                    .checked_add(8)
                    .ok_or_else(|| DecodeError::plain("protobuf group fixed64 length overflow"))?;
            },
            2 => {
                let (length, payload_start, _) = read_length(source, *cursor)?;
                let length = usize::try_from(length)
                    .map_err(|_| DecodeError::plain("protobuf group length does not fit usize"))?;
                *cursor = payload_start
                    .checked_add(length)
                    .ok_or_else(|| DecodeError::plain("protobuf group length overflow"))?;
            },
            3 => {
                skip_group(source, cursor, options, budget, depth + 1, number)?;
            },
            4 => {
                if number != opener {
                    return Err(DecodeError::plain("mismatched protobuf group"));
                }
                budget.field(cursor.saturating_sub(start), options)?;
                return Ok(());
            },
            5 => {
                *cursor = cursor
                    .checked_add(4)
                    .ok_or_else(|| DecodeError::plain("protobuf group fixed32 length overflow"))?;
            },
            _ => return Err(DecodeError::plain("protobuf group wire type is invalid")),
        }
        if *cursor > source.len() {
            return Err(DecodeError::plain("truncated protobuf group"));
        }
        budget.field(cursor.saturating_sub(start), options)?;
    }
}

fn read_varint(source: &[u8], start: usize) -> Result<(u64, usize, bool), DecodeError> {
    let mut value = 0u64;
    let mut shift = 0u32;
    let mut cursor = start;
    while cursor < source.len() && shift < 64 {
        let byte = source[cursor];
        cursor += 1;
        if shift == 63 && byte > 1 {
            return Err(DecodeError::plain("protobuf varint overflows u64"));
        }
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Ok((value, cursor, cursor - start == varint_len(value)));
        }
        shift += 7;
    }
    Err(DecodeError::plain("truncated or overlong protobuf varint"))
}

fn read_length(source: &[u8], start: usize) -> Result<(u64, usize, bool), DecodeError> {
    read_varint(source, start)
}

fn requirements_for(
    source: &[u8],
    scan: &Scan<'_>,
    output_bytes: usize,
    options: DecodeOptions,
) -> Result<RewriteExecutionRequirements, DecodeError> {
    let fields = scan.accounting.fields;
    let work_bytes = scan
        .accounting
        .work_bytes
        .checked_add(output_bytes)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    let allocations = scan
        .accounting
        .allocations
        .checked_add(OUTPUT_ALLOCATIONS)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
    let retained_bytes = source.len().checked_add(output_bytes).ok_or_else(|| {
        DecodeError::limit(DecodeLimit::Retained {
            observed: usize::MAX,
            maximum: options.max_retained_bytes,
        })
    })?;
    let scratch_bytes = scan
        .accounting
        .scratch_bytes
        .checked_add(output_bytes)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes,
        max_depth: scan.accounting.max_depth,
        allocations,
        retained_bytes,
        scratch_bytes,
    };
    check_requirements(requirements, options)?;
    Ok(requirements)
}

fn check_requirements(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    for (observed, maximum, limit) in [
        (
            requirements.output_bytes,
            options.max_output_bytes,
            DecodeLimit::OutputBytes {
                observed: requirements.output_bytes,
                maximum: options.max_output_bytes,
            },
        ),
        (
            requirements.fields,
            options.max_fields,
            DecodeLimit::Fields {
                observed: requirements.fields,
                maximum: options.max_fields,
            },
        ),
        (
            requirements.work_bytes,
            options.max_work_bytes,
            DecodeLimit::Work {
                observed: requirements.work_bytes,
                maximum: options.max_work_bytes,
            },
        ),
        (
            requirements.allocations,
            options.max_allocations,
            DecodeLimit::Allocations {
                observed: requirements.allocations,
                maximum: options.max_allocations,
            },
        ),
        (
            requirements.retained_bytes,
            options.max_retained_bytes,
            DecodeLimit::Retained {
                observed: requirements.retained_bytes,
                maximum: options.max_retained_bytes,
            },
        ),
        (
            requirements.scratch_bytes,
            options.max_scratch_bytes,
            DecodeLimit::Scratch {
                observed: requirements.scratch_bytes,
                maximum: options.max_scratch_bytes,
            },
        ),
    ] {
        if observed > maximum {
            return Err(DecodeError::limit(limit));
        }
    }
    if requirements.max_depth > options.recursion_limit {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    Ok(())
}

fn check_execution_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    for (observed, maximum, limit) in [
        (
            requirements.output_bytes,
            limits.output_bytes,
            DecodeLimit::OutputBytes {
                observed: requirements.output_bytes,
                maximum: limits.output_bytes,
            },
        ),
        (
            requirements.fields,
            limits.fields,
            DecodeLimit::Fields {
                observed: requirements.fields,
                maximum: limits.fields,
            },
        ),
        (
            requirements.work_bytes,
            limits.work_bytes,
            DecodeLimit::Work {
                observed: requirements.work_bytes,
                maximum: limits.work_bytes,
            },
        ),
        (
            requirements.allocations,
            limits.allocations,
            DecodeLimit::Allocations {
                observed: requirements.allocations,
                maximum: limits.allocations,
            },
        ),
        (
            requirements.retained_bytes,
            limits.retained_bytes,
            DecodeLimit::Retained {
                observed: requirements.retained_bytes,
                maximum: limits.retained_bytes,
            },
        ),
        (
            requirements.scratch_bytes,
            limits.scratch_bytes,
            DecodeLimit::Scratch {
                observed: requirements.scratch_bytes,
                maximum: limits.scratch_bytes,
            },
        ),
    ] {
        if observed > maximum {
            return Err(DecodeError::limit(limit));
        }
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    Ok(())
}

fn validate_geometry_write(write: GeometryWrite) -> Result<(), DecodeError> {
    if write.remove_envelope && geometry_adds_field(write) {
        return Err(DecodeError::plain(
            "geometry envelope removal cannot add a field",
        ));
    }
    if let FieldUpdate::Set(point) = write.position
        && (!point.x.is_finite() || !point.y.is_finite())
    {
        return Err(DecodeError::plain("drawable position must be finite"));
    }
    if let FieldUpdate::Set(size) = write.size
        && (!size.width.is_finite()
            || !size.height.is_finite()
            || size.width < 0.0
            || size.height < 0.0)
    {
        return Err(DecodeError::plain(
            "drawable size must be finite and non-negative",
        ));
    }
    if let FieldUpdate::Set(angle) = write.angle
        && !angle.is_finite()
    {
        return Err(DecodeError::plain("drawable angle must be finite"));
    }
    Ok(())
}

fn validate_properties_write(write: DrawablePropertiesWrite<'_>) -> Result<(), DecodeError> {
    if let FieldUpdate::Set(value) = write.hyperlink_url {
        validate_property_text(value, "hyperlink_url")?;
    }
    if let FieldUpdate::Set(value) = write.accessibility_description {
        validate_property_text(value, "accessibility_description")?;
    }
    Ok(())
}

fn validate_property_text(value: &str, field: &'static str) -> Result<(), DecodeError> {
    if value.len() > MAX_MESSAGE_BYTES {
        return Err(DecodeError::limit(DecodeLimit::OutputBytes {
            observed: value.len(),
            maximum: MAX_MESSAGE_BYTES,
        }));
    }
    let _ = field;
    Ok(())
}

fn semantic_change(
    scan: &Scan<'_>,
    geometry: GeometryWrite,
    properties: DrawablePropertiesWrite<'_>,
) -> bool {
    let current_geometry = scan.snapshot.geometry;
    let current_properties = scan.snapshot.properties;
    (geometry.remove_envelope && scan.geometry.is_some())
        || field_changes(
            current_geometry.and_then(|value| value.position),
            geometry.position,
        )
        || field_changes(current_geometry.and_then(|value| value.size), geometry.size)
        || field_changes(
            current_geometry.and_then(|value| value.flags),
            geometry.flags,
        )
        || field_changes(
            current_geometry.and_then(|value| value.angle),
            geometry.angle,
        )
        || field_changes(current_properties.hyperlink_url, properties.hyperlink_url)
        || field_changes(current_properties.locked, properties.locked)
        || field_changes(
            current_properties.aspect_ratio_locked,
            properties.aspect_ratio_locked,
        )
        || field_changes(
            current_properties.accessibility_description,
            properties.accessibility_description,
        )
}

fn field_changes<T: PartialEq + Copy>(current: Option<T>, update: FieldUpdate<T>) -> bool {
    match update {
        FieldUpdate::Preserve => false,
        FieldUpdate::Set(value) => current != Some(value),
        FieldUpdate::Clear => current.is_some(),
    }
}

fn measure_output(
    _source: &[u8],
    scan: &Scan<'_>,
    geometry: GeometryWrite,
    properties: DrawablePropertiesWrite<'_>,
    options: DecodeOptions,
) -> Result<usize, DecodeError> {
    let mut size = 0usize;
    for field in &scan.fields {
        let contribution = if field.number == GEOMETRY_FIELD {
            if geometry.remove_envelope {
                0
            } else if let Some(geometry_scan) = scan.geometry.as_ref() {
                let body = measure_geometry(geometry_scan, geometry)?;
                length_delimited_field_len(GEOMETRY_FIELD, body)
            } else {
                field.end - field.start
            }
        } else if field.number == HYPERLINK_FIELD {
            measure_string_field(field, properties.hyperlink_url)?
        } else if field.number == LOCKED_FIELD {
            measure_bool_field(field, properties.locked)?
        } else if field.number == ASPECT_RATIO_LOCKED_FIELD {
            measure_bool_field(field, properties.aspect_ratio_locked)?
        } else if field.number == ACCESSIBILITY_DESCRIPTION_FIELD {
            measure_string_field(field, properties.accessibility_description)?
        } else {
            field.end - field.start
        };
        size = size.checked_add(contribution).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::OutputBytes {
                observed: usize::MAX,
                maximum: options.max_output_bytes,
            })
        })?;
    }
    if scan.geometry_field.is_none() && geometry_adds_field(geometry) {
        let body = measure_new_geometry(geometry)?;
        size = size
            .checked_add(length_delimited_field_len(GEOMETRY_FIELD, body))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::OutputBytes {
                    observed: usize::MAX,
                    maximum: options.max_output_bytes,
                })
            })?;
    }
    if !has_field(&scan.fields, HYPERLINK_FIELD)
        && matches!(properties.hyperlink_url, FieldUpdate::Set(_))
    {
        size = size
            .checked_add(string_field_len(HYPERLINK_FIELD, properties.hyperlink_url))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::OutputBytes {
                    observed: usize::MAX,
                    maximum: options.max_output_bytes,
                })
            })?;
    }
    if !has_field(&scan.fields, LOCKED_FIELD) && matches!(properties.locked, FieldUpdate::Set(_)) {
        size = size
            .checked_add(varint_field_len_bool_update(
                LOCKED_FIELD,
                properties.locked,
            ))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::OutputBytes {
                    observed: usize::MAX,
                    maximum: options.max_output_bytes,
                })
            })?;
    }
    if !has_field(&scan.fields, ASPECT_RATIO_LOCKED_FIELD)
        && matches!(properties.aspect_ratio_locked, FieldUpdate::Set(_))
    {
        size = size
            .checked_add(varint_field_len_bool_update(
                ASPECT_RATIO_LOCKED_FIELD,
                properties.aspect_ratio_locked,
            ))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::OutputBytes {
                    observed: usize::MAX,
                    maximum: options.max_output_bytes,
                })
            })?;
    }
    if !has_field(&scan.fields, ACCESSIBILITY_DESCRIPTION_FIELD)
        && matches!(properties.accessibility_description, FieldUpdate::Set(_))
    {
        size = size
            .checked_add(string_field_len(
                ACCESSIBILITY_DESCRIPTION_FIELD,
                properties.accessibility_description,
            ))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::OutputBytes {
                    observed: usize::MAX,
                    maximum: options.max_output_bytes,
                })
            })?;
    }
    Ok(size)
}

fn measure_new_geometry(write: GeometryWrite) -> Result<usize, DecodeError> {
    let mut size = 0usize;
    if matches!(write.position, FieldUpdate::Set(_)) {
        size = size
            .checked_add(length_delimited_field_len(POSITION_FIELD, fixed_pair_len()))
            .ok_or_else(|| DecodeError::plain("geometry size overflow"))?;
    }
    if matches!(write.size, FieldUpdate::Set(_)) {
        size = size
            .checked_add(length_delimited_field_len(SIZE_FIELD, fixed_pair_len()))
            .ok_or_else(|| DecodeError::plain("geometry size overflow"))?;
    }
    if matches!(write.flags, FieldUpdate::Set(_)) {
        size = size
            .checked_add(varint_field_len_u32(FLAGS_FIELD, write.flags))
            .ok_or_else(|| DecodeError::plain("geometry size overflow"))?;
    }
    if matches!(write.angle, FieldUpdate::Set(_)) {
        size = size
            .checked_add(fixed32_field_len(ANGLE_FIELD))
            .ok_or_else(|| DecodeError::plain("geometry size overflow"))?;
    }
    Ok(size)
}

fn geometry_adds_field(write: GeometryWrite) -> bool {
    matches!(write.position, FieldUpdate::Set(_))
        || matches!(write.size, FieldUpdate::Set(_))
        || matches!(write.flags, FieldUpdate::Set(_))
        || matches!(write.angle, FieldUpdate::Set(_))
}

fn measure_string_field(
    field: &ParsedField,
    update: FieldUpdate<&str>,
) -> Result<usize, DecodeError> {
    match update {
        FieldUpdate::Preserve => Ok(field.end - field.start),
        FieldUpdate::Clear => Ok(0),
        FieldUpdate::Set(value) => Ok(length_delimited_field_len(field.number, value.len())),
    }
}

fn measure_bool_field(
    field: &ParsedField,
    update: FieldUpdate<bool>,
) -> Result<usize, DecodeError> {
    match update {
        FieldUpdate::Preserve => Ok(field.end - field.start),
        FieldUpdate::Clear => Ok(0),
        FieldUpdate::Set(value) => Ok(varint_field_len_bool(field.number, value)),
    }
}

fn has_field(fields: &[ParsedField], number: u32) -> bool {
    fields.iter().any(|field| field.number == number)
}

fn fixed_pair_len() -> usize {
    fixed32_field_len(1) + fixed32_field_len(2)
}

fn fixed32_field_len(number: u32) -> usize {
    varint_len(u64::from(number << 3 | 5)) + 4
}

fn length_delimited_field_len(number: u32, payload: usize) -> usize {
    varint_len(u64::from(number << 3 | 2)) + varint_len(payload as u64) + payload
}

fn string_field_len(number: u32, update: FieldUpdate<&str>) -> usize {
    match update {
        FieldUpdate::Set(value) => length_delimited_field_len(number, value.len()),
        _ => 0,
    }
}

fn varint_field_len_u32(number: u32, update: FieldUpdate<u32>) -> usize {
    match update {
        FieldUpdate::Set(value) => varint_len(u64::from(number << 3)) + varint_len(value.into()),
        _ => 0,
    }
}

fn varint_field_len_bool(number: u32, _value: bool) -> usize {
    varint_len(u64::from(number << 3)) + 1
}

fn varint_field_len_bool_update(number: u32, update: FieldUpdate<bool>) -> usize {
    match update {
        FieldUpdate::Set(value) => varint_field_len_bool(number, value),
        _ => 0,
    }
}

fn build_output(
    source: &[u8],
    scan: &Scan<'_>,
    geometry: GeometryWrite,
    properties: DrawablePropertiesWrite<'_>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut saw_geometry = false;
    let mut saw_hyperlink = false;
    let mut saw_locked = false;
    let mut saw_aspect = false;
    let mut saw_accessibility = false;
    for field in &scan.fields {
        match field.number {
            GEOMETRY_FIELD => {
                if geometry.remove_envelope {
                    continue;
                }
                saw_geometry = true;
                let geometry_scan = scan
                    .geometry
                    .as_ref()
                    .ok_or_else(|| DecodeError::plain("geometry scan is missing"))?;
                let geometry_source = payload(source, *field)?;
                let body_len = measure_geometry(geometry_scan, geometry)?;
                append_length_delimited_prefix(output, GEOMETRY_FIELD, body_len);
                build_geometry(geometry_source, geometry_scan, geometry, output)?;
            },
            HYPERLINK_FIELD => {
                saw_hyperlink = true;
                append_string_update(
                    output,
                    field.number,
                    properties.hyperlink_url,
                    source,
                    *field,
                )?;
            },
            LOCKED_FIELD => {
                saw_locked = true;
                if matches!(properties.locked, FieldUpdate::Preserve) {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    append_bool_update(output, field.number, properties.locked)?;
                }
            },
            ASPECT_RATIO_LOCKED_FIELD => {
                saw_aspect = true;
                if matches!(properties.aspect_ratio_locked, FieldUpdate::Preserve) {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    append_bool_update(output, field.number, properties.aspect_ratio_locked)?;
                }
            },
            ACCESSIBILITY_DESCRIPTION_FIELD => {
                saw_accessibility = true;
                append_string_update(
                    output,
                    field.number,
                    properties.accessibility_description,
                    source,
                    *field,
                )?;
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
    }
    if !saw_geometry && geometry_adds_field(geometry) {
        let body_len = measure_new_geometry(geometry)?;
        append_length_delimited_prefix(output, GEOMETRY_FIELD, body_len);
        build_new_geometry(geometry, output);
    }
    if !saw_hyperlink {
        append_string_absent(output, HYPERLINK_FIELD, properties.hyperlink_url);
    }
    if !saw_locked {
        append_bool_absent(output, LOCKED_FIELD, properties.locked);
    }
    if !saw_aspect {
        append_bool_absent(
            output,
            ASPECT_RATIO_LOCKED_FIELD,
            properties.aspect_ratio_locked,
        );
    }
    if !saw_accessibility {
        append_string_absent(
            output,
            ACCESSIBILITY_DESCRIPTION_FIELD,
            properties.accessibility_description,
        );
    }
    Ok(())
}

fn measure_geometry(geometry: &GeometryScan, write: GeometryWrite) -> Result<usize, DecodeError> {
    let mut size = 0usize;
    for field in &geometry.fields {
        let contribution = match field.number {
            POSITION_FIELD => {
                measure_nested_pair(field, geometry.position.as_ref(), write.position)?
            },
            SIZE_FIELD => measure_nested_pair(field, geometry.size.as_ref(), write.size)?,
            FLAGS_FIELD => measure_existing_varint(field, write.flags)?,
            ANGLE_FIELD => measure_existing_fixed32(field, write.angle)?,
            _ => field.end - field.start,
        };
        size = size
            .checked_add(contribution)
            .ok_or_else(|| DecodeError::plain("geometry size overflow"))?;
    }
    if !has_field(&geometry.fields, POSITION_FIELD) && matches!(write.position, FieldUpdate::Set(_))
    {
        size = size
            .checked_add(length_delimited_field_len(POSITION_FIELD, fixed_pair_len()))
            .ok_or_else(|| DecodeError::plain("geometry size overflow"))?;
    }
    if !has_field(&geometry.fields, SIZE_FIELD) && matches!(write.size, FieldUpdate::Set(_)) {
        size = size
            .checked_add(length_delimited_field_len(SIZE_FIELD, fixed_pair_len()))
            .ok_or_else(|| DecodeError::plain("geometry size overflow"))?;
    }
    if !has_field(&geometry.fields, FLAGS_FIELD) && matches!(write.flags, FieldUpdate::Set(_)) {
        size = size
            .checked_add(varint_field_len_u32(FLAGS_FIELD, write.flags))
            .ok_or_else(|| DecodeError::plain("geometry size overflow"))?;
    }
    if !has_field(&geometry.fields, ANGLE_FIELD) && matches!(write.angle, FieldUpdate::Set(_)) {
        size = size
            .checked_add(fixed32_field_len(ANGLE_FIELD))
            .ok_or_else(|| DecodeError::plain("geometry size overflow"))?;
    }
    Ok(size)
}

fn measure_nested_pair<T: Copy>(
    field: &ParsedField,
    current: Option<&NestedFixed>,
    update: FieldUpdate<T>,
) -> Result<usize, DecodeError> {
    match update {
        FieldUpdate::Preserve => Ok(field.end - field.start),
        FieldUpdate::Clear => Ok(0),
        FieldUpdate::Set(_) => {
            if let Some(current) = current {
                let mut size = 0usize;
                for nested in &current.fields {
                    size = size
                        .checked_add(if nested.number == 1 || nested.number == 2 {
                            fixed32_field_len(nested.number)
                        } else {
                            nested.end - nested.start
                        })
                        .ok_or_else(|| DecodeError::plain("nested pair size overflow"))?;
                }
                if !has_field(&current.fields, 1) {
                    size = size
                        .checked_add(fixed32_field_len(1))
                        .ok_or_else(|| DecodeError::plain("nested pair size overflow"))?;
                }
                if !has_field(&current.fields, 2) {
                    size = size
                        .checked_add(fixed32_field_len(2))
                        .ok_or_else(|| DecodeError::plain("nested pair size overflow"))?;
                }
                Ok(length_delimited_field_len(field.number, size))
            } else {
                Ok(length_delimited_field_len(field.number, fixed_pair_len()))
            }
        },
    }
}

fn measure_existing_varint(
    field: &ParsedField,
    update: FieldUpdate<u32>,
) -> Result<usize, DecodeError> {
    match update {
        FieldUpdate::Preserve => Ok(field.end - field.start),
        FieldUpdate::Clear => Ok(0),
        FieldUpdate::Set(_) => Ok(varint_field_len_u32(field.number, update)),
    }
}

fn measure_existing_fixed32(
    field: &ParsedField,
    update: FieldUpdate<f32>,
) -> Result<usize, DecodeError> {
    match update {
        FieldUpdate::Preserve => Ok(field.end - field.start),
        FieldUpdate::Clear => Ok(0),
        FieldUpdate::Set(_) => Ok(fixed32_field_len(field.number)),
    }
}

fn build_geometry(
    source: &[u8],
    geometry: &GeometryScan,
    write: GeometryWrite,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut saw_position = false;
    let mut saw_size = false;
    let mut saw_flags = false;
    let mut saw_angle = false;
    for field in &geometry.fields {
        match field.number {
            POSITION_FIELD => {
                saw_position = true;
                append_pair_update(
                    output,
                    source,
                    *field,
                    geometry.position.as_ref(),
                    write.position,
                )?;
            },
            SIZE_FIELD => {
                saw_size = true;
                append_pair_update(output, source, *field, geometry.size.as_ref(), write.size)?;
            },
            FLAGS_FIELD => {
                saw_flags = true;
                if matches!(write.flags, FieldUpdate::Preserve) {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    append_u32_update(output, field.number, write.flags)?;
                }
            },
            ANGLE_FIELD => {
                saw_angle = true;
                if matches!(write.angle, FieldUpdate::Preserve) {
                    output.extend_from_slice(&source[field.start..field.end]);
                } else {
                    append_f32_update(output, field.number, write.angle)?;
                }
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
    }
    if !saw_position {
        append_pair_absent(output, POSITION_FIELD, write.position, true);
    }
    if !saw_size {
        append_pair_absent(output, SIZE_FIELD, write.size, false);
    }
    if !saw_flags {
        append_u32_absent(output, FLAGS_FIELD, write.flags);
    }
    if !saw_angle {
        append_f32_absent(output, ANGLE_FIELD, write.angle);
    }
    Ok(())
}

fn build_new_geometry(write: GeometryWrite, output: &mut Vec<u8>) {
    append_pair_absent(output, POSITION_FIELD, write.position, true);
    append_pair_absent(output, SIZE_FIELD, write.size, false);
    append_u32_absent(output, FLAGS_FIELD, write.flags);
    append_f32_absent(output, ANGLE_FIELD, write.angle);
}

fn append_pair_update<T: Copy + PairValues>(
    output: &mut Vec<u8>,
    source: &[u8],
    field: ParsedField,
    current: Option<&NestedFixed>,
    update: FieldUpdate<T>,
) -> Result<(), DecodeError> {
    match update {
        FieldUpdate::Preserve => output.extend_from_slice(&source[field.start..field.end]),
        FieldUpdate::Clear => {},
        FieldUpdate::Set(value) => {
            let _ = value;
            let body_len = current
                .map(|nested| {
                    nested
                        .fields
                        .iter()
                        .map(|nested_field| {
                            if nested_field.number == 1 || nested_field.number == 2 {
                                fixed32_field_len(nested_field.number)
                            } else {
                                nested_field.end - nested_field.start
                            }
                        })
                        .try_fold(0usize, |sum, size| sum.checked_add(size))
                        .and_then(|size| {
                            let missing = usize::from(!has_field(&nested.fields, 1))
                                + usize::from(!has_field(&nested.fields, 2));
                            size.checked_add(missing * 5)
                        })
                        .unwrap_or(fixed_pair_len())
                })
                .unwrap_or(fixed_pair_len());
            append_length_delimited_prefix(output, field.number, body_len);
            if let Some(nested) = current {
                let nested_source = payload(source, field)?;
                build_pair(nested_source, nested, value, output);
            } else {
                build_new_pair(value, output);
            }
        },
    }
    Ok(())
}

fn build_pair<T: Copy + PairValues>(
    source: &[u8],
    current: &NestedFixed,
    value: T,
    output: &mut Vec<u8>,
) {
    let (first, second) = pair_values(value);
    let mut saw_first = false;
    let mut saw_second = false;
    for field in &current.fields {
        match field.number {
            1 => {
                saw_first = true;
                append_fixed32_field(output, 1, first);
            },
            2 => {
                saw_second = true;
                append_fixed32_field(output, 2, second);
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
    }
    if !saw_first {
        append_fixed32_field(output, 1, first);
    }
    if !saw_second {
        append_fixed32_field(output, 2, second);
    }
}

fn build_new_pair<T: Copy + PairValues>(value: T, output: &mut Vec<u8>) {
    let (first, second) = pair_values(value);
    append_fixed32_field(output, 1, first);
    append_fixed32_field(output, 2, second);
}

trait PairValues {
    fn pair(self) -> (f32, f32);
}

impl PairValues for Point {
    fn pair(self) -> (f32, f32) {
        (self.x, self.y)
    }
}

impl PairValues for Size {
    fn pair(self) -> (f32, f32) {
        (self.width, self.height)
    }
}

fn pair_values<T: Copy + PairValues>(value: T) -> (u32, u32) {
    let (first, second) = value.pair();
    (first.to_bits(), second.to_bits())
}

fn append_length_delimited_prefix(output: &mut Vec<u8>, number: u32, payload_len: usize) {
    push_varint(output, u64::from(number << 3 | 2));
    push_varint(output, payload_len as u64);
}

fn append_fixed32_field(output: &mut Vec<u8>, number: u32, value: u32) {
    push_varint(output, u64::from(number << 3 | 5));
    output.extend_from_slice(&value.to_le_bytes());
}

fn append_u32_update(
    output: &mut Vec<u8>,
    number: u32,
    update: FieldUpdate<u32>,
) -> Result<(), DecodeError> {
    match update {
        FieldUpdate::Preserve => Err(DecodeError::plain(
            "preserved scalar field requires its source span",
        )),
        FieldUpdate::Set(value) => {
            append_u32_field(output, number, value);
            Ok(())
        },
        FieldUpdate::Clear => Ok(()),
    }
}

fn append_u32_absent(output: &mut Vec<u8>, number: u32, update: FieldUpdate<u32>) {
    if let FieldUpdate::Set(value) = update {
        append_u32_field(output, number, value);
    }
}

fn append_u32_field(output: &mut Vec<u8>, number: u32, value: u32) {
    push_varint(output, u64::from(number << 3));
    push_varint(output, u64::from(value));
}

fn append_f32_update(
    output: &mut Vec<u8>,
    number: u32,
    update: FieldUpdate<f32>,
) -> Result<(), DecodeError> {
    match update {
        FieldUpdate::Preserve => Err(DecodeError::plain(
            "preserved scalar field requires its source span",
        )),
        FieldUpdate::Set(value) => {
            append_fixed32_field(output, number, value.to_bits());
            Ok(())
        },
        FieldUpdate::Clear => Ok(()),
    }
}

fn append_f32_absent(output: &mut Vec<u8>, number: u32, update: FieldUpdate<f32>) {
    if let FieldUpdate::Set(value) = update {
        append_fixed32_field(output, number, value.to_bits());
    }
}

fn append_pair_absent<T: Copy + PairValues>(
    output: &mut Vec<u8>,
    number: u32,
    update: FieldUpdate<T>,
    _point: bool,
) {
    if let FieldUpdate::Set(value) = update {
        let (first, second) = pair_values(value);
        append_length_delimited_prefix(output, number, fixed_pair_len());
        append_fixed32_field(output, 1, first);
        append_fixed32_field(output, 2, second);
    }
}

fn append_string_update(
    output: &mut Vec<u8>,
    number: u32,
    update: FieldUpdate<&str>,
    source: &[u8],
    field: ParsedField,
) -> Result<(), DecodeError> {
    match update {
        FieldUpdate::Preserve => output.extend_from_slice(&source[field.start..field.end]),
        FieldUpdate::Set(value) => {
            append_length_delimited_prefix(output, number, value.len());
            output.extend_from_slice(value.as_bytes());
        },
        FieldUpdate::Clear => {},
    }
    Ok(())
}

fn append_string_absent(output: &mut Vec<u8>, number: u32, update: FieldUpdate<&str>) {
    if let FieldUpdate::Set(value) = update {
        append_length_delimited_prefix(output, number, value.len());
        output.extend_from_slice(value.as_bytes());
    }
}

fn append_bool_update(
    output: &mut Vec<u8>,
    number: u32,
    update: FieldUpdate<bool>,
) -> Result<(), DecodeError> {
    match update {
        FieldUpdate::Preserve => Err(DecodeError::plain(
            "preserved scalar field requires its source span",
        )),
        FieldUpdate::Set(value) => {
            push_varint(output, u64::from(number << 3));
            push_varint(output, u64::from(value));
            Ok(())
        },
        FieldUpdate::Clear => Ok(()),
    }
}

fn append_bool_absent(output: &mut Vec<u8>, number: u32, update: FieldUpdate<bool>) {
    if let FieldUpdate::Set(value) = update {
        push_varint(output, u64::from(number << 3));
        push_varint(output, u64::from(value));
    }
}

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
            .with_max_output_bytes(4096)
            .with_resource_limits(1024, 16 * 1024)
            .with_max_allocations(1024)
            .with_max_retained_bytes(source.len().saturating_add(8192))
            .with_max_scratch_bytes(64 * 1024)
    }

    fn push_noncanonical_varint(output: &mut Vec<u8>, value: u64) {
        let mut canonical = Vec::new();
        push_varint(&mut canonical, value);
        let last = canonical.pop().expect("varint has a terminal byte");
        output.extend_from_slice(&canonical);
        output.push(last | 0x80);
        output.push(0);
    }

    fn drawable_fixture() -> Vec<u8> {
        let mut point = Vec::new();
        append_fixed32_field(&mut point, 99, 0x1234_5678);
        append_fixed32_field(&mut point, POINT_X_FIELD, 12.5_f32.to_bits());
        append_fixed32_field(&mut point, POINT_Y_FIELD, 24.25_f32.to_bits());

        let mut size = Vec::new();
        append_u32_field(&mut size, 98, 7);
        append_fixed32_field(&mut size, SIZE_WIDTH_FIELD, 96.0_f32.to_bits());
        append_fixed32_field(&mut size, SIZE_HEIGHT_FIELD, 48.0_f32.to_bits());

        let mut geometry = Vec::new();
        append_string_absent(&mut geometry, 20, FieldUpdate::Set("geometry-unknown"));
        append_length_delimited_prefix(&mut geometry, POSITION_FIELD, point.len());
        geometry.extend_from_slice(&point);
        append_length_delimited_prefix(&mut geometry, SIZE_FIELD, size.len());
        geometry.extend_from_slice(&size);
        append_u32_field(&mut geometry, FLAGS_FIELD, 3);
        append_fixed32_field(&mut geometry, ANGLE_FIELD, 15.0_f32.to_bits());

        let mut source = Vec::new();
        append_u32_field(&mut source, 99, 7);
        append_length_delimited_prefix(&mut source, 2, 2);
        source.extend_from_slice(&[0x08, 0x01]);
        append_length_delimited_prefix(&mut source, 3, 2);
        source.extend_from_slice(&[0x08, 0x02]);
        append_length_delimited_prefix(&mut source, GEOMETRY_FIELD, geometry.len());
        source.extend_from_slice(&geometry);
        append_string_absent(
            &mut source,
            HYPERLINK_FIELD,
            FieldUpdate::Set("https://example.test/original"),
        );
        append_bool_absent(&mut source, LOCKED_FIELD, FieldUpdate::Set(false));
        append_u32_field(&mut source, 100, 11);
        append_bool_absent(
            &mut source,
            ASPECT_RATIO_LOCKED_FIELD,
            FieldUpdate::Set(true),
        );
        append_string_absent(
            &mut source,
            ACCESSIBILITY_DESCRIPTION_FIELD,
            FieldUpdate::Set("Original description"),
        );
        source
    }

    #[test]
    fn decode_reports_semantic_geometry_and_properties() {
        let source = drawable_fixture();
        let (snapshot, report) = decode_drawable_archive_with_report(&source, options(&source))
            .expect("fixture should decode");
        let geometry = snapshot.geometry().expect("geometry should be present");
        assert_eq!(geometry.position(), Some(Point::new(12.5, 24.25)));
        assert_eq!(geometry.size(), Some(Size::new(96.0, 48.0)));
        assert_eq!(geometry.flags(), Some(3));
        assert_eq!(geometry.angle(), Some(15.0));
        let properties = snapshot.properties();
        assert_eq!(
            properties.hyperlink_url(),
            Some("https://example.test/original")
        );
        assert_eq!(properties.locked(), Some(false));
        assert_eq!(properties.aspect_ratio_locked(), Some(true));
        assert_eq!(
            properties.accessibility_description(),
            Some("Original description")
        );
        assert_eq!(report.input_bytes(), source.len());
        assert!(report.fields() >= 12);
        assert!(report.allocations() >= BUFFA_LOGICAL_ALLOCATIONS_PER_TRAVERSAL * 2);
        assert!(report.scratch_bytes() > 0);
        assert!(report.work_bytes() > source.len());
    }

    #[test]
    fn rewrite_splices_known_values_and_retains_unknown_nested_fields() {
        let source = drawable_fixture();
        let output = rewrite_drawable_archive(
            &source,
            GeometryWrite::preserve()
                .set_position(Point::new(100.0, 200.0))
                .set_angle(270.0),
            DrawablePropertiesWrite::preserve().set_accessibility_description("Edited 文本"),
            options(&source),
        )
        .expect("fixture rewrite should succeed");

        assert!(output.changed());
        assert!(
            output
                .bytes()
                .windows(3)
                .any(|window| window == [0x98, 0x06, 7])
        );
        assert!(
            output
                .bytes()
                .windows(4)
                .any(|window| window == [0xA2, 0x01, 0x10, b'g'])
        );
        assert!(
            output
                .bytes()
                .windows(4)
                .any(|window| window == [0x12, 0x02, 0x08, 1])
        );
        assert!(
            output
                .bytes()
                .windows(4)
                .any(|window| window == [0x1A, 0x02, 0x08, 2])
        );
        assert!(
            output
                .bytes()
                .windows(3)
                .any(|window| window == [0x9D, 0x06, 0x78])
        );

        let (snapshot, _) =
            decode_drawable_archive_with_report(output.bytes(), options(output.bytes()))
                .expect("rewritten fixture should decode");
        let geometry = snapshot.geometry().expect("geometry should remain present");
        assert_eq!(geometry.position(), Some(Point::new(100.0, 200.0)));
        assert_eq!(geometry.size(), Some(Size::new(96.0, 48.0)));
        assert_eq!(geometry.flags(), Some(3));
        assert_eq!(geometry.angle(), Some(270.0));
        assert_eq!(snapshot.properties().locked(), Some(false));
        assert_eq!(
            snapshot.properties().accessibility_description(),
            Some("Edited 文本")
        );
    }

    #[test]
    fn preserve_is_exact_and_clear_keeps_unknown_geometry_fields() {
        let source = drawable_fixture();
        let preserved = rewrite_drawable_archive(
            &source,
            GeometryWrite::preserve(),
            DrawablePropertiesWrite::preserve(),
            options(&source),
        )
        .expect("preserve rewrite should succeed");
        assert!(!preserved.changed());
        assert_eq!(preserved.bytes(), source.as_slice());

        let cleared = rewrite_drawable_archive(
            &source,
            GeometryWrite::new(None, None, None, None),
            DrawablePropertiesWrite::new(None, None, None, None),
            options(&source),
        )
        .expect("clear rewrite should succeed");
        assert!(cleared.changed());
        assert!(
            cleared
                .bytes()
                .windows(3)
                .any(|window| window == [0x98, 0x06, 7])
        );
        assert!(
            cleared
                .bytes()
                .windows(3)
                .any(|window| window == [0xA0, 0x06, 11])
        );
        assert!(
            cleared
                .bytes()
                .windows(4)
                .any(|window| window == [0xA2, 0x01, 0x10, b'g'])
        );
        assert!(
            cleared
                .bytes()
                .windows(4)
                .any(|window| window == [0x12, 0x02, 0x08, 1])
        );
        assert!(
            cleared
                .bytes()
                .windows(4)
                .any(|window| window == [0x1A, 0x02, 0x08, 2])
        );
        let snapshot = decode_drawable_archive(cleared.bytes(), options(cleared.bytes()))
            .expect("cleared fixture should decode");
        let geometry = snapshot.geometry().expect("geometry envelope is retained");
        assert_eq!(geometry.position(), None);
        assert_eq!(geometry.size(), None);
        assert_eq!(geometry.flags(), None);
        assert_eq!(geometry.angle(), None);
        assert_eq!(snapshot.properties().hyperlink_url(), None);
    }

    #[test]
    fn exact_replacement_can_create_and_then_remove_all_known_fields() {
        let source = Vec::new();
        let created = rewrite_drawable_archive(
            &source,
            GeometryWrite::new(
                Some(Point::new(1.0, 2.0)),
                Some(Size::new(3.0, 4.0)),
                Some(3),
                Some(45.0),
            ),
            DrawablePropertiesWrite::new(
                Some("https://example.test/new"),
                Some(true),
                Some(false),
                Some("Created description"),
            ),
            options(&source),
        )
        .expect("known fields should be creatable");
        let snapshot = decode_drawable_archive(created.bytes(), options(created.bytes()))
            .expect("created fields should decode");
        let geometry = snapshot.geometry().expect("created geometry should exist");
        assert_eq!(geometry.position(), Some(Point::new(1.0, 2.0)));
        assert_eq!(geometry.size(), Some(Size::new(3.0, 4.0)));
        assert_eq!(geometry.flags(), Some(3));
        assert_eq!(geometry.angle(), Some(45.0));
        assert_eq!(snapshot.properties().locked(), Some(true));

        let cleared = rewrite_drawable_archive(
            created.bytes(),
            GeometryWrite::new(None, None, None, None),
            DrawablePropertiesWrite::new(None, None, None, None),
            options(created.bytes()),
        )
        .expect("created fields should be removable");
        assert_eq!(cleared.bytes(), &[0x0A, 0x00]);

        let envelope_removed = rewrite_drawable_archive(
            created.bytes(),
            GeometryWrite::new(None, None, None, None).remove_envelope(),
            DrawablePropertiesWrite::new(None, None, None, None),
            options(created.bytes()),
        )
        .expect("the explicitly absent geometry envelope should be restorable");
        assert!(envelope_removed.changed());
        assert!(envelope_removed.bytes().is_empty());
    }

    #[test]
    fn exact_execution_limits_are_checked_before_replay() {
        let source = drawable_fixture();
        let prepared = prepare_drawable_archive_rewrite(
            &source,
            GeometryWrite::preserve().set_flags(9),
            DrawablePropertiesWrite::preserve(),
            options(&source),
        )
        .expect("rewrite should be admitted");
        let requirements = prepared.requirements();
        let mut limits = requirements.exact();
        limits.output_bytes = limits.output_bytes.saturating_sub(1);
        let error = prepared
            .execute(limits)
            .expect_err("output ceiling should reject before replay");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::OutputBytes { .. })
        ));
    }

    #[test]
    fn buffa_traversal_limits_are_charged_before_projection() {
        let source = drawable_fixture();
        let constrained = options(&source).with_max_allocations(8);
        let error = decode_drawable_archive(&source, constrained)
            .expect_err("the final Buffa traversal should be precharged");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Allocations { .. })
        ));
    }

    #[test]
    fn malformed_known_fields_are_rejected() {
        let mut source = Vec::new();
        append_u32_field(&mut source, LOCKED_FIELD, 2);
        let error = decode_drawable_archive(&source, options(&source))
            .expect_err("noncanonical bool should be rejected");
        assert!(error.resource_limit().is_none());

        let mut duplicate = Vec::new();
        append_bool_absent(&mut duplicate, LOCKED_FIELD, FieldUpdate::Set(false));
        append_bool_absent(&mut duplicate, LOCKED_FIELD, FieldUpdate::Set(true));
        assert!(decode_drawable_archive(&duplicate, options(&duplicate)).is_err());
    }

    #[test]
    fn opaque_unknown_noncanonical_framing_is_preserved() {
        let mut source = Vec::new();
        push_noncanonical_varint(&mut source, u64::from(99_u32 << 3));
        push_noncanonical_varint(&mut source, 1);
        push_noncanonical_varint(&mut source, u64::from(100_u32 << 3 | 2));
        push_noncanonical_varint(&mut source, 1);
        source.push(b'x');

        decode_drawable_archive(&source, options(&source))
            .expect("opaque noncanonical fields should be readable");
        let output = rewrite_drawable_archive(
            &source,
            GeometryWrite::preserve(),
            DrawablePropertiesWrite::preserve().set_locked(true),
            options(&source),
        )
        .expect("opaque noncanonical fields should survive a rewrite");
        assert!(
            output
                .bytes()
                .windows(5)
                .any(|window| { window == [0x98, 0x86, 0x00, 0x81, 0x00] })
        );
        assert!(
            output
                .bytes()
                .windows(6)
                .any(|window| { window == [0xA2, 0x86, 0x00, 0x81, 0x00, b'x'] })
        );
        assert!(
            output
                .bytes()
                .windows(2)
                .any(|window| window == [0x28, 0x01])
        );
    }

    #[test]
    fn known_noncanonical_keys_lengths_and_values_are_rejected() {
        let mut key = Vec::new();
        push_noncanonical_varint(&mut key, u64::from(LOCKED_FIELD << 3));
        key.push(0);
        assert!(decode_drawable_archive(&key, options(&key)).is_err());

        let mut length = vec![0x42];
        push_noncanonical_varint(&mut length, 1);
        length.push(b'x');
        assert!(decode_drawable_archive(&length, options(&length)).is_err());

        let mut value = vec![0x28];
        push_noncanonical_varint(&mut value, 1);
        assert!(decode_drawable_archive(&value, options(&value)).is_err());
    }

    #[test]
    fn large_keys_are_rejected_before_narrowing_and_groups_match() {
        let mut large_key = vec![0xff; 9];
        large_key.push(0x01);
        assert!(decode_drawable_archive(&large_key, options(&large_key)).is_err());

        let mut large_group_child = Vec::new();
        push_varint(&mut large_group_child, u64::from(99_u32 << 3 | 3));
        large_group_child.extend_from_slice(&large_key);
        assert!(decode_drawable_archive(&large_group_child, options(&large_group_child)).is_err());

        let mut mismatched_group = Vec::new();
        push_varint(&mut mismatched_group, u64::from(99_u32 << 3 | 3));
        push_varint(&mut mismatched_group, u64::from(100_u32 << 3 | 4));
        assert!(decode_drawable_archive(&mismatched_group, options(&mismatched_group)).is_err());
    }
}
