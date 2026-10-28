//! Selector-first, source-preserving frame, property and playback edits of
//! one sheet drawable: text boxes, shapes, images, movies, audio clips and
//! charts (record 0378).
//!
//! The drawable is found through the shared [`super::sheet_drawable_graph`]
//! proof. Each staged change is a field-level rewrite of the selected
//! drawable message: unknown fields, field order and every other object stay
//! byte-identical, the root previews are kept (as the legacy host editor
//! does for these in-place edits) and the package metadata is untouched.
//!
//! This is the behaviour of the host editor's `set_sheet_*_geometry`,
//! `flip_sheet_*`, `restore_sheet_*_original_size`, `set_sheet_*_properties`,
//! `set_sheet_audio_position` and `set_sheet_*_playback_settings` routes,
//! collapsed into one family.

use std::fmt;
use std::time::Duration;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::media::playback::{
    MediaLoopMode, MediaPlaybackSettings, MediaVolume, TimeField,
};
use litchi_iwa_common::wire::{WireView, transform_length_delimited_field};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::movie_playback_codec as playback_codec;
use litchi_iwa_protos::shape_properties_codec as codec;
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{
    self as graph, IMAGE_MESSAGE_TYPE, MOVIE_MESSAGE_TYPE, SHAPE_INFO_MESSAGE_TYPE,
};
use super::structure::{self, Artifacts, Fault, FaultLimit};
use crate::selector::SheetSelector;
use crate::sheet::drawable::{
    DrawableSelector, FlipAxis, Geometry, GeometryFlags, Kind, Point, Properties, Size,
};

const IMAGE_ORIGINAL_SIZE_FIELD: u32 = 4;
const IMAGE_NATURAL_SIZE_FIELD: u32 = 9;
const MOVIE_ORIGINAL_SIZE_FIELD: u32 = 20;
const MOVIE_NATURAL_SIZE_FIELD: u32 = 21;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;
const WRAPPED_DRAWABLE_FIELD: u32 = 1;

/// A finite resource charged by one sheet-drawable property transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Native object identifiers.
    Identifiers,
    /// Temporary graph scratch bytes.
    ScratchBytes,
    /// Fallible graph allocation events.
    AllocationEvents,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::Identifiers => "object identifiers",
            Self::ScratchBytes => "scratch bytes",
            Self::AllocationEvents => "allocation events",
        })
    }
}

/// A typed, content-redacted sheet-drawable property failure.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted drawable graph or the selected payload is malformed.
    #[error("the selected Numbers sheet-drawable source is invalid")]
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
    #[error("could not allocate {amount} sheet-drawable items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested state.
    #[error("the edited Numbers sheet-drawable properties failed verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Numbers sheet-drawable properties patch does not match this package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No drawable matches the selector.
    #[error("the sheet has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The edit does not apply to the selected kind (a size of an audio
    /// clip, a flip of a text box, playback of an image, any edit of a
    /// table through this family).
    #[error("{operation} does not apply to a Numbers sheet {kind}")]
    UnsupportedOperation {
        /// The refused verb.
        operation: &'static str,
        /// The selected kind.
        kind: &'static str,
    },
    /// The value cannot be stored (non-finite or negative frame parts, an
    /// invalid playback value).
    #[error("invalid Numbers sheet-drawable value: {reason}")]
    InvalidValue {
        /// Why the value was refused.
        reason: &'static str,
    },
    /// The drawable lacks the stored state the verb needs (no frame, no
    /// layout flags, no original size).
    #[error("the Numbers sheet {kind} has no {missing}")]
    MissingState {
        /// The selected kind.
        kind: &'static str,
        /// The missing state.
        missing: &'static str,
    },
    /// The source does not keep an exact artifact for a changed publication,
    /// or the drawable is stored in more than one component.
    #[error("this Numbers source does not support exact sheet-drawable edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedOperation { .. } => RefusalKind::UnsupportedOperation,
            Self::InvalidValue { .. } => RefusalKind::InvalidValue,
            Self::MissingState { .. } => RefusalKind::UnsupportedTarget,
            Self::UnsupportedSource => RefusalKind::UnsupportedContent,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Unsupported | Fault::Dependency => Self::Refused(Refusal::UnsupportedSource),
            Fault::Invalid => Self::InvalidSource,
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                },
                observed,
                maximum,
            },
            Fault::GraphLimit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    structure::GraphLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                    structure::GraphLimitKind::AllocationEvents => LimitKind::AllocationEvents,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Verification => Self::Verification,
        }
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        structure::map_wire(&error).into()
    }
}

/// The stored frame, properties and playback of one sheet drawable.
#[derive(Debug, Clone, PartialEq)]
pub struct Snapshot {
    kind: Kind,
    geometry: Option<Geometry>,
    properties: Properties,
    playback: Option<MediaPlaybackSettings>,
    original_size: Option<Size>,
    natural_size: Option<Size>,
}

impl Snapshot {
    /// What the drawable is.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
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
    /// `None` for every other kind.
    #[must_use]
    pub const fn playback(&self) -> Option<MediaPlaybackSettings> {
        self.playback
    }

    /// The stored original media size (the Original Size target) of an
    /// image or movie.
    #[must_use]
    pub const fn original_size(&self) -> Option<Size> {
        self.original_size
    }

    /// The stored untransformed media size of an image or movie.
    #[must_use]
    pub const fn natural_size(&self) -> Option<Size> {
        self.natural_size
    }
}

/// The private target of one edit: the sheet, the drawable and its message.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Target {
    sheet: usize,
    selector: DrawableSelector,
    drawable: u64,
    message_index: usize,
    message_type: u32,
}

/// Frame, property and playback changes staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    before: Snapshot,
    after: Snapshot,
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

    const fn kind_name(&self) -> &'static str {
        self.before.kind.name()
    }

    fn refuse(&self, operation: &'static str) -> Error {
        Refusal::UnsupportedOperation {
            operation,
            kind: self.kind_name(),
        }
        .into()
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

    /// Replace the frame of a text box, shape, image, movie or chart.
    ///
    /// # Errors
    ///
    /// Refuses audio (use [`Self::set_position`]), a drawable without a
    /// frame, and non-finite or negative parts.
    pub fn set_geometry(&mut self, geometry: Geometry) -> Result<&mut Self, Error> {
        if self.before.kind == Kind::Audio {
            return Err(self.refuse("set_geometry"));
        }
        self.stored_geometry()?;
        if !geometry.is_valid() {
            return Err(Refusal::InvalidValue {
                reason: "frame parts must be finite and sizes non-negative",
            }
            .into());
        }
        self.after.geometry = Some(geometry);
        Ok(self)
    }

    /// Move the drawable, keeping its size, flags and rotation.
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
        self.after.geometry = Some(Geometry {
            position: Some(position),
            ..geometry
        });
        Ok(self)
    }

    /// Apply one native Arrange Flip command to a shape, image or movie.
    ///
    /// # Errors
    ///
    /// Refuses the other kinds and a frame without layout flags.
    pub fn flip(&mut self, axis: FlipAxis) -> Result<&mut Self, Error> {
        if !matches!(self.before.kind, Kind::Shape | Kind::Image | Kind::Movie) {
            return Err(self.refuse("flip"));
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
    /// Refuses the other kinds and a drawable without a positive original
    /// size.
    pub fn restore_original_size(&mut self) -> Result<&mut Self, Error> {
        if !matches!(self.before.kind, Kind::Image | Kind::Movie) {
            return Err(self.refuse("restore_original_size"));
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

    /// Replace the shared drawable properties.
    pub fn set_properties(&mut self, properties: Properties) -> &mut Self {
        self.after.properties = properties;
        self
    }

    /// Replace the playback settings of a movie or audio clip.
    ///
    /// # Errors
    ///
    /// Refuses the other kinds and an invalid playback value.
    pub fn set_playback(&mut self, settings: MediaPlaybackSettings) -> Result<&mut Self, Error> {
        if !matches!(self.before.kind, Kind::Movie | Kind::Audio) {
            return Err(self.refuse("set_playback"));
        }
        let settings = settings
            .canonicalize()
            .map_err(|_error| Refusal::InvalidValue {
                reason: "invalid playback settings",
            })?;
        self.after.playback = Some(settings);
        Ok(self)
    }

    /// Validate and publish the staged edit.
    ///
    /// # Errors
    ///
    /// Source, resource and verification errors; failure never publishes a
    /// partial package. An unchanged edit is an exact no-op.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// Exact-source reversible sheet-drawable properties patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: Artifacts,
    sheet: usize,
    selector: DrawableSelector,
    before: Snapshot,
    after: Snapshot,
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

    /// Whether both the semantic state and the exact artifact are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet: self.sheet,
            selector: self.selector,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Content-free diagnostics of one sheet-drawable properties publication.
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

/// The validated immutable result of one sheet-drawable properties
/// transaction.
#[must_use = "a sheet-drawable properties commit contains the validated package snapshot"]
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

impl Package {
    /// Read the frame, properties and playback of one sheet drawable.
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`], [`Refusal::DrawableNotFound`], and
    /// [`Error::InvalidSource`] for a malformed graph.
    pub fn sheet_drawable_properties<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Snapshot, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        Ok(resolve(self, sheet, drawable.into())?.1)
    }

    /// Begin a selector-first frame, property and playback edit of one
    /// sheet drawable.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::drawable::{DrawableSelector, Point};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_sheet_drawable_properties(0, DrawableSelector::text_box(0))?;
    /// edit.set_position(Point { x: 72.0, y: 144.0 })?;
    /// let commit = edit.commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`], [`Refusal::DrawableNotFound`],
    /// [`Refusal::UnsupportedOperation`] for a table, and
    /// [`Error::InvalidSource`] for a malformed graph.
    pub fn edit_sheet_drawable_properties<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let (target, before) = resolve(self, sheet, drawable.into())?;
        Ok(Edit {
            source: self,
            target,
            after: before.clone(),
            before,
        })
    }

    /// Replay an exact-source sheet-drawable properties patch.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when the patch was made against another
    /// artifact or the drawable no longer matches its `before` state.
    pub fn apply_sheet_drawable_properties(&self, patch: &Patch) -> Result<Commit, Error> {
        let (_, current) = resolve(self, patch.sheet, patch.selector)?;
        if current != patch.before {
            return Err(Error::PatchConflict);
        }
        let package = patch.artifacts.apply(self)?.ok_or(Error::PatchConflict)?;
        if patch.is_noop() {
            return Ok(Commit {
                package,
                patch: patch.clone(),
                diagnostics: Diagnostics {
                    changed: false,
                    touched_components: 0,
                },
            });
        }
        let (_, reopened) = resolve(&package, patch.sheet, patch.selector)?;
        if reopened != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components: 1,
            },
        })
    }
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

fn resolve(
    package: &Package,
    sheet: usize,
    selector: DrawableSelector,
) -> Result<(Target, Snapshot), Error> {
    let drawables = graph::sheet_drawables(package, sheet)?;
    let entry = drawables
        .select(selector)
        .ok_or(Refusal::DrawableNotFound { selector })?;
    let Some((message_index, message_type)) = entry.message else {
        return Err(Refusal::UnsupportedOperation {
            operation: "edit_sheet_drawable_properties",
            kind: entry.kind.name(),
        }
        .into());
    };
    if entry.kind == Kind::Other {
        return Err(Refusal::UnsupportedOperation {
            operation: "edit_sheet_drawable_properties",
            kind: entry.kind.name(),
        }
        .into());
    }
    let object = graph::object(package, entry.identifier)?;
    let payload = object
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    let drawable = graph::drawable_payload(message_type, payload)?;
    if graph::drawable_parent(drawable)? != Some(drawables.sheet.identifier) {
        return Err(Error::InvalidSource);
    }
    let (geometry, properties) = graph::frame_and_properties(drawable)?;
    let sizes = match message_type {
        IMAGE_MESSAGE_TYPE => Some((IMAGE_ORIGINAL_SIZE_FIELD, IMAGE_NATURAL_SIZE_FIELD)),
        MOVIE_MESSAGE_TYPE => Some((MOVIE_ORIGINAL_SIZE_FIELD, MOVIE_NATURAL_SIZE_FIELD)),
        _ => None,
    };
    let (original_size, natural_size) = match sizes {
        Some((original, natural)) => (
            stored_size(payload, original)?,
            stored_size(payload, natural)?,
        ),
        None => (None, None),
    };
    let playback = if message_type == MOVIE_MESSAGE_TYPE {
        Some(decode_playback(payload)?)
    } else {
        None
    };
    Ok((
        Target {
            sheet,
            selector,
            drawable: entry.identifier,
            message_index,
            message_type,
        },
        Snapshot {
            kind: entry.kind,
            geometry,
            properties,
            playback,
            original_size,
            natural_size,
        },
    ))
}

fn stored_size(payload: &[u8], number: u32) -> Result<Option<Size>, Error> {
    let view = WireView::parse(payload, WireLimits::DEFAULT)?;
    let Some(field) = graph::unique_field(&view, number, 2)? else {
        return Ok(None);
    };
    let size = WireView::parse(field.payload(), WireLimits::DEFAULT)?;
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
                .map_err(|_error| Error::InvalidSource)?,
        );
    }
    Ok(Some(Size {
        width: parts[0],
        height: parts[1],
    }))
}

fn playback_options(source: &[u8]) -> playback_codec::DecodeOptions {
    let limits = WireLimits::DEFAULT;
    let base = playback_codec::DecodeOptions::for_source(source);
    let recursion_limit = u32::try_from(limits.max_nesting())
        .unwrap_or(u32::MAX)
        .min(8);
    base.with_max_input_bytes(base.max_input_bytes().min(limits.max_input_bytes()))
        .with_max_output_bytes(limits.max_output_bytes())
        .with_resource_limits(
            source
                .len()
                .max(1)
                .saturating_mul(4)
                .min(limits.max_fields()),
            source
                .len()
                .max(1)
                .saturating_mul(8)
                .min(limits.max_rewrite_work()),
        )
        .with_recursion_limit(recursion_limit)
}

fn seconds(value: f32, field: TimeField) -> Result<Duration, Error> {
    let _ = field;
    if !value.is_finite() || value < 0.0 {
        return Err(Error::InvalidSource);
    }
    Duration::try_from_secs_f32(value).map_err(|_error| Error::InvalidSource)
}

fn as_seconds(value: Duration) -> Result<f32, Error> {
    let seconds = value.as_secs_f64();
    if !seconds.is_finite() || seconds > f64::from(f32::MAX) {
        return Err(Refusal::InvalidValue {
            reason: "a playback time is out of range",
        }
        .into());
    }
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the native field is a float; the range was checked"
    )]
    Ok(seconds as f32)
}

fn decode_playback(payload: &[u8]) -> Result<MediaPlaybackSettings, Error> {
    let snapshot = playback_codec::decode_movie_playback(payload, playback_options(payload))
        .map_err(|_error| Error::InvalidSource)?;
    let start_time = snapshot
        .start_time
        .map(|value| seconds(value, TimeField::Start))
        .transpose()?;
    let end_time = seconds(snapshot.end_time, TimeField::End)?;
    let poster_time = snapshot
        .poster_time
        .map(|value| seconds(value, TimeField::Poster))
        .transpose()?;
    let loop_mode = snapshot.loop_mode.map(MediaLoopMode::from_raw);
    let volume = snapshot
        .volume
        .map(MediaVolume::new)
        .transpose()
        .map_err(|_error| Error::InvalidSource)?;
    MediaPlaybackSettings {
        start_time,
        end_time,
        poster_time,
        loop_mode,
        volume,
    }
    .canonicalize()
    .map_err(|_error| Error::InvalidSource)
}

fn rewrite_playback(payload: &[u8], settings: MediaPlaybackSettings) -> Result<Vec<u8>, Error> {
    let write = playback_codec::MoviePlaybackWrite::from_values(
        settings.start_time.map(as_seconds).transpose()?,
        as_seconds(settings.end_time)?,
        settings.poster_time.map(as_seconds).transpose()?,
        settings.loop_mode.map(MediaLoopMode::as_raw),
        settings.volume.map(MediaVolume::as_f32),
    );
    playback_codec::rewrite_movie_playback(payload, write, playback_options(payload))
        .map_err(|_error| Error::InvalidSource)
}

fn rewrite_drawable(source: &[u8], before: &Snapshot, after: &Snapshot) -> Result<Vec<u8>, Error> {
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
    codec::rewrite_drawable_archive(source, geometry, properties, graph::codec_options(source))
        .map(codec::RewriteOutput::into_bytes)
        .map_err(|_error| Error::InvalidSource)
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let target = edit.target;
    let geometry_changed = edit.after.geometry != edit.before.geometry;
    let properties_changed = edit.after.properties != edit.before.properties;
    let playback_changed = edit.after.playback != edit.before.playback;
    if !geometry_changed && !properties_changed && !playback_changed {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: Artifacts::unchanged(source)?,
                sheet: target.sheet,
                selector: target.selector,
                before: edit.before.clone(),
                after: edit.before,
            },
            diagnostics: Diagnostics {
                changed: false,
                touched_components: 0,
            },
        });
    }
    let object = graph::object(source, target.drawable)?;
    let mut payload = object
        .messages
        .get(target.message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    if geometry_changed || properties_changed {
        let rewrite = |drawable: &[u8]| rewrite_drawable(drawable, &edit.before, &edit.after);
        payload = if target.message_type == SHAPE_INFO_MESSAGE_TYPE {
            transform_length_delimited_field(&payload, WRAPPED_DRAWABLE_FIELD, |shape| {
                transform_length_delimited_field(shape, WRAPPED_DRAWABLE_FIELD, rewrite)
            })?
        } else {
            transform_length_delimited_field(&payload, WRAPPED_DRAWABLE_FIELD, rewrite)?
        };
    }
    if playback_changed {
        let settings = edit.after.playback.ok_or(Error::InvalidSource)?;
        payload = rewrite_playback(&payload, settings)?;
    }
    let mut payload = Some(payload);
    let mut replace = |object: &mut litchi_iwa_core::ArchiveObject,
                       limits: litchi_iwa_core::Limits| {
        object
            .replace_message_preserving_header(
                target.message_index,
                RawMessage {
                    type_: target.message_type,
                    data: payload.take().ok_or(Fault::Invalid)?,
                },
                limits,
            )
            .map(drop)
            .map_err(|error| structure::map_core(&error))
    };
    let published = graph::publish_in_place(source, &mut [(target.drawable, &mut replace)])?;
    let (reopened_target, reopened) = resolve(&published.package, target.sheet, target.selector)?;
    if reopened_target.drawable != target.drawable || reopened != edit.after {
        return Err(Error::Verification);
    }
    Ok(Commit {
        patch: Patch {
            artifacts: Artifacts::new(source, &published.package)?,
            sheet: target.sheet,
            selector: target.selector,
            before: edit.before,
            after: edit.after,
        },
        diagnostics: Diagnostics {
            changed: true,
            touched_components: published.touched_components,
        },
        package: published.package,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn playback_seconds_round_trip() {
        let value = as_seconds(Duration::from_millis(1_250)).unwrap();
        assert_eq!(
            seconds(value, TimeField::End).unwrap(),
            Duration::from_millis(1_250)
        );
        assert!(seconds(f32::NAN, TimeField::Start).is_err());
    }

    #[test]
    fn refusal_kinds_are_classified() {
        use litchi_core::edit::RefusalKind;
        assert_eq!(Refusal::SheetNotFound.kind(), RefusalKind::NotFound);
        assert_eq!(
            Refusal::UnsupportedOperation {
                operation: "flip",
                kind: "text box"
            }
            .kind(),
            RefusalKind::UnsupportedOperation
        );
    }
}
