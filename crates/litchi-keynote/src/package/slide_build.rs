//! Exact-source, selector-first creation, update and removal of slide
//! builds (record 0372; a port of the migration host's `add_slide_build`,
//! `set_slide_build`, `remove_slide_build` and `slide_builds`).
//!
//! A build is addressed by its slide and its position in the slide's build
//! list (`KN.SlideArchive.builds`, the order Keynote's Build Order lists
//! them); it animates one drawable addressed by a [`DrawableSelector`]. An
//! [`Edit`] stages operations against one immutable package; selectors
//! resolve against that base snapshot. [`Edit::commit`] applies them in one
//! structural workspace, publishes one candidate, reopens it and reads every
//! touched build back.

mod codec;

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use thiserror::Error as ThisError;

use self::codec::Native;
use super::structure::StructureError;
use super::structure::wire::{self, Fields, LENGTH_DELIMITED};
use super::structure::workspace::Workspace;
use super::{Package, PhysicalSource, ReadError, SLIDE_MESSAGE_TYPE, SLIDE_NODE_MESSAGE_TYPE};
use crate::build::{Acceleration, Effect, Start, TimingCurve, UnknownText};
use crate::{DrawableSelector, Seconds, SlideSelector, SlideSelectorError};

const BUILD_MESSAGE_TYPE: u32 = 8;
const BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;
const SLIDE_BUILDS: u32 = 2;
const SLIDE_OWNED_DRAWABLES: u32 = 7;
const SLIDE_BUILD_CHUNKS: u32 = 43;
const NODE_BUILD_EVENT_COUNT: u32 = 15;
const NODE_HAS_EXPLICIT_BUILDS: u32 = 20;
const NODE_BUILD_EVENT_CACHE_VERSION: u32 = 26;
const NODE_EXPLICIT_BUILDS_CACHE_VERSION: u32 = 27;

/// Whether a build brings its drawable in, takes it out, or acts on it.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Stage {
    /// Build In.
    In,
    /// Build Out.
    Out,
    /// Action (rotate, scale, opacity, move, emphasis).
    Action,
    /// A stage of another native kind, kept as read.
    Unknown(UnknownText),
}

/// One build's animation: stage, effect, timing and start.
///
/// Constructors give Keynote's native defaults for each effect family (the
/// host builder's `appear_in`, `dissolve_in`, …, `keyboard_out`). Native
/// parameters a build carries beside these values (delivery, text delivery
/// options, custom effect parameters) are kept as read when settings are
/// written back.
///
/// ```
/// use litchi_keynote::Seconds;
/// use litchi_keynote::build::Start;
/// use litchi_keynote::slide::build::{Settings, Stage};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let settings = Settings::appear(Stage::In)
///     .with_start(Start::AfterPrevious)?
///     .with_duration(Seconds::new(0.75)?)?;
/// assert_eq!(settings.stage(), &Stage::In);
/// # Ok(())
/// # }
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Settings {
    stage: Stage,
    effect: Effect,
    duration: Seconds,
    delay: Seconds,
    start: Start,
    timing_curve: Option<TimingCurve>,
    native: Native,
}

fn seconds(value: f64) -> Seconds {
    Seconds::new(value).unwrap_or(Seconds::ZERO)
}

impl Settings {
    fn with_defaults(stage: Stage, effect: Effect, duration: f64) -> Self {
        let native = Native::defaults(&stage, &effect);
        Self {
            stage,
            effect,
            duration: seconds(duration),
            delay: Seconds::ZERO,
            start: Start::OnClick,
            timing_curve: None,
            native,
        }
    }

    /// Appear (in or out), one second.
    #[must_use]
    pub fn appear(stage: Stage) -> Self {
        Self::with_defaults(stage, Effect::Appear, 1.0)
    }

    /// A text-object effect (Dissolve by character 1 s, Shimmer 1.5 s,
    /// Skid 1.25 s, Swoosh 1 s, Trace 2 s), in or out.
    #[must_use]
    pub fn object(stage: Stage, effect: crate::build::ObjectEffect) -> Self {
        use crate::build::ObjectEffect;
        let duration = match effect {
            ObjectEffect::Shimmer => 1.5,
            ObjectEffect::Skid(_) => 1.25,
            ObjectEffect::Trace(_) => 2.0,
            _ => 1.0,
        };
        Self::with_defaults(stage, Effect::Object(effect), duration)
    }

    /// A Rotate, Scale, Opacity or Move action, one second.
    #[must_use]
    pub fn action(action: crate::build::Action) -> Self {
        Self::with_defaults(Stage::Action, Effect::Action(action), 1.0)
    }

    /// An emphasis action (Blink, Bounce, Flip, Jiggle, Pulse one second;
    /// Pop half a second).
    #[must_use]
    pub fn emphasis(emphasis: crate::build::Emphasis) -> Self {
        let duration = if matches!(emphasis, crate::build::Emphasis::Pop(_)) {
            0.5
        } else {
            1.0
        };
        Self::with_defaults(Stage::Action, Effect::Emphasis(emphasis), duration)
    }

    /// Keyboard, in or out, three seconds.
    #[must_use]
    pub fn keyboard(stage: Stage, keyboard: crate::build::Keyboard) -> Self {
        Self::with_defaults(stage, Effect::Keyboard(keyboard), 3.0)
    }

    /// The stage.
    #[must_use]
    pub const fn stage(&self) -> &Stage {
        &self.stage
    }

    /// The effect.
    #[must_use]
    pub const fn effect(&self) -> &Effect {
        &self.effect
    }

    /// The duration.
    #[must_use]
    pub const fn duration(&self) -> Seconds {
        self.duration
    }

    /// The delay before an automatic start.
    #[must_use]
    pub const fn delay(&self) -> Seconds {
        self.delay
    }

    /// The start.
    #[must_use]
    pub const fn start(&self) -> Start {
        self.start
    }

    /// The custom timing curve of a typed action with custom acceleration.
    #[must_use]
    pub const fn timing_curve(&self) -> Option<&TimingCurve> {
        self.timing_curve.as_ref()
    }

    /// This build with another duration.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidTiming`] for a zero duration.
    pub fn with_duration(mut self, duration: Seconds) -> Result<Self, Error> {
        if duration == Seconds::ZERO {
            return Err(Error::Refused(Refusal::InvalidTiming));
        }
        self.duration = duration;
        Ok(self)
    }

    /// This build with another delay (only automatic After Transition and
    /// After Previous starts wait).
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidTiming`] for a delay on an On Click or
    /// With Previous start.
    pub fn with_delay(mut self, delay: Seconds) -> Result<Self, Error> {
        if matches!(self.start, Start::OnClick | Start::WithPrevious) && delay != Seconds::ZERO {
            return Err(Error::Refused(Refusal::InvalidTiming));
        }
        self.delay = delay;
        Ok(self)
    }

    /// This build with another start.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidTiming`] for an On Click or With Previous
    /// start while the build has a delay.
    pub fn with_start(mut self, start: Start) -> Result<Self, Error> {
        if matches!(start, Start::OnClick | Start::WithPrevious) && self.delay != Seconds::ZERO {
            return Err(Error::Refused(Refusal::InvalidTiming));
        }
        self.start = start;
        Ok(self)
    }

    /// This build with another effect of the same stage (native defaults of
    /// the build are kept).
    #[must_use]
    pub fn with_effect(mut self, effect: Effect) -> Self {
        self.effect = effect;
        self
    }

    /// This typed action with a custom timing curve: the action's
    /// acceleration becomes [`Acceleration::Custom`].
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`Refusal::NotAnAction`] for an
    /// effect other than Rotate, Scale, Opacity or Move.
    pub fn with_timing_curve(mut self, curve: TimingCurve) -> Result<Self, Error> {
        use crate::build::{Action, Motion, Opacity, Rotation, Scale};
        let invalid = |_| Error::Refused(Refusal::NotAnAction);
        let custom = Acceleration::Custom;
        self.effect = match self.effect {
            Effect::Action(Action::Rotate(rotation)) => Effect::Action(Action::Rotate(
                Rotation::new(rotation.total_degrees(), rotation.direction(), custom)
                    .map_err(invalid)?,
            )),
            Effect::Action(Action::Scale(scale)) => Effect::Action(Action::Scale(
                Scale::new(scale.factor(), custom).map_err(invalid)?,
            )),
            Effect::Action(Action::Opacity(opacity)) => Effect::Action(Action::Opacity(
                Opacity::new(opacity.percent(), custom).map_err(invalid)?,
            )),
            Effect::Action(Action::Move(motion)) => Effect::Action(Action::Move(Motion::new(
                motion.path().clone(),
                motion.align_to_path(),
                custom,
            ))),
            _ => return Err(Error::Refused(Refusal::NotAnAction)),
        };
        self.timing_curve = Some(curve);
        Ok(self)
    }
}

/// One build a slide holds, as read from the package.
#[derive(Debug, Clone, PartialEq)]
pub struct Build {
    drawable: DrawableSelector,
    settings: Settings,
    chunks: usize,
}

impl Build {
    /// The drawable the build animates.
    #[must_use]
    pub const fn drawable(&self) -> DrawableSelector {
        self.drawable
    }

    /// The build's settings.
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }

    /// How many timing chunks (build events) the build owns.
    #[must_use]
    pub const fn chunks(&self) -> usize {
        self.chunks
    }
}

/// A finite resource governed by a slide-build transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Native objects in the package.
    Objects,
    /// The native object-identifier space.
    Identifiers,
    /// Builds on one slide.
    Builds,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::Objects => "objects",
            Self::Identifiers => "object identifiers",
            Self::Builds => "builds",
        })
    }
}

/// A slide-build failure that never exposes native identifiers.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The package cannot be edited safely.
    #[error("the Keynote source cannot be edited safely")]
    InvalidSource,
    /// A resource limit was exceeded.
    #[error("Keynote slide-build {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exceeded resource.
        kind: LimitKind,
        /// The observed amount.
        observed: u64,
        /// The configured maximum.
        maximum: u64,
    },
    /// The published candidate did not read back as staged.
    #[error("the Keynote slide-build candidate failed verification")]
    Verification,
    /// The patch does not match the exact source package.
    #[error("the Keynote slide-build patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request the slide-build family refuses.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The package was not opened from an exact physical source.
    #[error("this Keynote source does not support structural edits")]
    UnsupportedSource,
    /// No slide has the selected navigator name.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// No slide is at the selected position.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// The requested position.
        position: Position,
    },
    /// Several slides share the selected name, or the name is empty.
    #[error("the Keynote slide selector is ambiguous or empty")]
    AmbiguousSelector,
    /// No drawable is at the selected position.
    #[error("the Keynote slide has no drawable at position {position:?}")]
    DrawablePositionNotFound {
        /// The requested position.
        position: Position,
    },
    /// No build is at the selected position.
    #[error("the Keynote slide has no build at position {position:?}")]
    BuildPositionNotFound {
        /// The requested position.
        position: Position,
    },
    /// The build was removed earlier in this edit.
    #[error("the Keynote build was removed earlier in this edit")]
    BuildRemoved,
    /// The start is not valid at the build's place in the build order
    /// (After Transition only first; With/After Previous never first).
    #[error("the Keynote build start is not valid at this position in the build order")]
    InvalidStart,
    /// An action build must carry typed action parameters to be created or
    /// changed; an unmodelled action keeps its native parameters only when
    /// unchanged.
    #[error("the Keynote action needs typed parameters")]
    OpaqueAction,
    /// A zero duration, or a delay on an On Click or With Previous start.
    #[error("the Keynote build timing is not valid")]
    InvalidTiming,
    /// A custom acceleration needs a timing curve, which only typed
    /// actions carry.
    #[error("a custom Keynote acceleration needs a timing curve on a typed action")]
    NotAnAction,
    /// The build or its timing chunks use a native shape this family does
    /// not edit (a transient merge payload, chunks outside the slide).
    #[error("the Keynote build uses an unsupported native shape")]
    UnsupportedShape,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SlideNameNotFound
            | Self::SlidePositionNotFound { .. }
            | Self::DrawablePositionNotFound { .. }
            | Self::BuildPositionNotFound { .. }
            | Self::BuildRemoved => RefusalKind::NotFound,
            Self::AmbiguousSelector
            | Self::InvalidStart
            | Self::OpaqueAction
            | Self::InvalidTiming
            | Self::NotAnAction => RefusalKind::InvalidValue,
            Self::UnsupportedSource | Self::UnsupportedShape => RefusalKind::UnsupportedContent,
        }
    }
}

#[derive(Clone)]
enum Operation {
    Add {
        slide: u64,
        drawable: u64,
        settings: Settings,
    },
    Set {
        slide: u64,
        build: u64,
        settings: Settings,
    },
    Remove {
        slide: u64,
        build: u64,
    },
}

/// One bounded slide-build edit staged against an immutable package.
///
/// ```no_run
/// use litchi_keynote::slide::build::{Settings, Stage};
/// use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let package = Package::open("deck.key")?;
/// let mut edit = package.edit_slide_builds();
/// edit.add(SlideSelector::index(0), DrawableSelector::index(0), Settings::appear(Stage::In))?;
/// let commit = edit.commit()?;
/// assert_eq!(commit.package().slide_build_list(SlideSelector::index(0))?.len(), 1);
/// # Ok(())
/// # }
/// ```
pub struct Edit<'a> {
    source: &'a Package,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Edit")
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

struct SlideFields {
    slide: u64,
    owned: Vec<u64>,
    builds: Vec<u64>,
    chunks: Vec<u64>,
}

fn read_slide(package: &Package, selector: SlideSelector<'_>) -> Result<SlideFields, Error> {
    let show = package.show().map_err(map_read_error)?;
    let position = show
        .select_slide(selector)
        .map_err(|error| match error {
            SlideSelectorError::EmptySlideName | SlideSelectorError::DuplicateSlideName { .. } => {
                Error::Refused(Refusal::AmbiguousSelector)
            },
        })?
        .ok_or(match selector {
            SlideSelector::Name(_) => Error::Refused(Refusal::SlideNameNotFound),
            SlideSelector::Position(position) => {
                Error::Refused(Refusal::SlidePositionNotFound { position })
            },
        })?
        .position();
    let record = package
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource)?;
    let wire = package.wire_limits().map_err(|_| Error::InvalidSource)?;
    let object = package
        .object(record.slide_identifier)
        .ok_or(Error::InvalidSource)?;
    let payload = object
        .messages
        .iter()
        .find(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let fields = Fields::parse(&payload.data, wire).map_err(lift)?;
    let references = |number| -> Result<Vec<u64>, Error> {
        fields
            .all(number)
            .map(|field| {
                wire::reference_identifier(field.message().map_err(lift)?, wire).map_err(lift)
            })
            .collect()
    };
    Ok(SlideFields {
        slide: record.slide_identifier,
        owned: references(SLIDE_OWNED_DRAWABLES)?,
        builds: references(SLIDE_BUILDS)?,
        chunks: references(SLIDE_BUILD_CHUNKS)?,
    })
}

fn unique_payload(package: &Package, identifier: u64, message_type: u32) -> Result<&[u8], Error> {
    let object = package.object(identifier).ok_or(Error::InvalidSource)?;
    let mut found = object
        .messages
        .iter()
        .filter(|message| message.type_ == message_type);
    let first = found.next().ok_or(Error::InvalidSource)?;
    if found.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok(&first.data)
}

/// The slide's builds in build order, with each one's chunks in event
/// order.
fn read_builds(
    package: &Package,
    fields: &SlideFields,
) -> Result<Vec<(u64, Build, Vec<u64>)>, Error> {
    let wire = package.wire_limits().map_err(|_| Error::InvalidSource)?;
    let mut chunks: Vec<(u64, u64, codec::Chunk)> = Vec::with_capacity(fields.chunks.len());
    for chunk in &fields.chunks {
        let (build, value) = codec::read_chunk(
            unique_payload(package, *chunk, BUILD_CHUNK_MESSAGE_TYPE)?,
            wire,
        )
        .map_err(lift)?;
        let build = build.ok_or(Error::InvalidSource)?;
        if !fields.builds.contains(&build) {
            return Err(Error::InvalidSource);
        }
        chunks.push((*chunk, build, value));
    }
    let first_chunk = fields.chunks.first().copied();
    let mut builds = Vec::with_capacity(fields.builds.len());
    for build in &fields.builds {
        let payload = unique_payload(package, *build, BUILD_MESSAGE_TYPE)?;
        let drawable = codec::build_drawable(payload, wire)
            .map_err(lift)?
            .ok_or(Error::InvalidSource)?;
        let position = fields
            .owned
            .iter()
            .position(|owned| *owned == drawable)
            .ok_or(Error::InvalidSource)?;
        let own = chunks
            .iter()
            .filter(|(_, owner, _)| owner == build)
            .collect::<Vec<_>>();
        let starts = own
            .first()
            .is_some_and(|(id, _, _)| Some(*id) == first_chunk);
        let settings = codec::read_settings(payload, own.first().map(|(_, _, c)| *c), starts, wire)
            .map_err(lift)?;
        builds.push((
            *build,
            Build {
                drawable: DrawableSelector::index(position),
                settings,
                chunks: own.len(),
            },
            own.iter().map(|(id, _, _)| *id).collect(),
        ));
    }
    Ok(builds)
}

fn check_start(start: Start, event_index: usize) -> Result<(), Error> {
    match (start, event_index) {
        (Start::AfterTransition, 0) | (Start::OnClick, _) => Ok(()),
        (Start::WithPrevious | Start::AfterPrevious, index) if index > 0 => Ok(()),
        _ => Err(Error::Refused(Refusal::InvalidStart)),
    }
}

fn check_settings(settings: &Settings) -> Result<(), Error> {
    if matches!(settings.stage, Stage::Action) && !codec::is_typed_action(&settings.effect) {
        return Err(Error::Refused(Refusal::OpaqueAction));
    }
    if codec::action_acceleration(&settings.effect) == Some(Acceleration::Custom)
        && settings.timing_curve.is_none()
    {
        return Err(Error::Refused(Refusal::NotAnAction));
    }
    Ok(())
}

impl<'a> Edit<'a> {
    const fn new(source: &'a Package) -> Self {
        Self {
            source,
            operations: Vec::new(),
        }
    }

    /// Stage a new build of one drawable, appended to the slide's build
    /// order (host `add_slide_build`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::OpaqueAction`] for an
    /// action without typed parameters, [`Refusal::NotAnAction`] for a
    /// custom acceleration without a timing curve, and
    /// [`Refusal::InvalidStart`] for a start that is not valid as the next
    /// build event.
    pub fn add<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        drawable: impl Into<DrawableSelector>,
        settings: Settings,
    ) -> Result<&mut Self, Error> {
        check_settings(&settings)?;
        let fields = read_slide(self.source, slide.into())?;
        let drawable = drawable.into();
        let target = *fields.owned.get(drawable.as_index()).ok_or(Error::Refused(
            Refusal::DrawablePositionNotFound {
                position: drawable.as_position(),
            },
        ))?;
        let staged = self.events_after(fields.slide, fields.chunks.len());
        check_start(settings.start, staged)?;
        self.operations.push(Operation::Add {
            slide: fields.slide,
            drawable: target,
            settings,
        });
        Ok(self)
    }

    /// The slide's build-event count after the operations staged so far.
    fn events_after(&self, slide: u64, base: usize) -> usize {
        self.operations
            .iter()
            .fold(base, |count, operation| match operation {
                Operation::Add { slide: s, .. } if *s == slide => count + 1,
                _ => count,
            })
    }

    fn select_build(
        &self,
        slide: SlideSelector<'_>,
        build: Position,
    ) -> Result<(SlideFields, u64, Build, Vec<u64>), Error> {
        let fields = read_slide(self.source, slide)?;
        let builds = read_builds(self.source, &fields)?;
        let (identifier, value, chunks) = builds.into_iter().nth(build.get()).ok_or(
            Error::Refused(Refusal::BuildPositionNotFound { position: build }),
        )?;
        if self.operations.iter().any(|operation| {
            matches!(operation, Operation::Remove { build: removed, .. } if *removed == identifier)
        }) {
            return Err(Error::Refused(Refusal::BuildRemoved));
        }
        Ok((fields, identifier, value, chunks))
    }

    /// Stage new settings for one build and every timing chunk it owns
    /// (host `set_slide_build`). Native parameters the settings do not model
    /// keep their bytes; setting the current settings stages nothing.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors, [`Refusal::OpaqueAction`] when an
    /// action without typed parameters would change,
    /// [`Refusal::NotAnAction`] and [`Refusal::InvalidStart`].
    pub fn set<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        build: Position,
        settings: Settings,
    ) -> Result<&mut Self, Error> {
        let (fields, identifier, current, chunks) = self.select_build(slide.into(), build)?;
        if current.settings == settings {
            return Ok(self);
        }
        let opaque = |settings: &Settings| {
            matches!(settings.stage, Stage::Action) && !codec::is_typed_action(&settings.effect)
        };
        if (opaque(&current.settings) || opaque(&settings))
            && (current.settings.stage != settings.stage
                || current.settings.effect != settings.effect)
        {
            return Err(Error::Refused(Refusal::OpaqueAction));
        }
        if codec::action_acceleration(&settings.effect) == Some(Acceleration::Custom)
            && settings.timing_curve.is_none()
        {
            return Err(Error::Refused(Refusal::NotAnAction));
        }
        let first = chunks
            .first()
            .ok_or(Error::Refused(Refusal::UnsupportedShape))?;
        let event = fields
            .chunks
            .iter()
            .position(|chunk| chunk == first)
            .ok_or(Error::InvalidSource)?;
        check_start(settings.start, event)?;
        self.operations.push(Operation::Set {
            slide: fields.slide,
            build: identifier,
            settings,
        });
        Ok(self)
    }

    /// Stage the removal of one build and its timing chunks (host
    /// `remove_slide_build`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::InvalidStart`] when a
    /// remaining build's start would become invalid.
    pub fn remove<'s>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        build: Position,
    ) -> Result<&mut Self, Error> {
        let slide = slide.into();
        let (fields, identifier, _, _) = self.select_build(slide, build)?;
        let builds = read_builds(self.source, &fields)?;
        let mut event = 0;
        for (other, value, chunks) in &builds {
            if *other == identifier || chunks.is_empty() {
                continue;
            }
            check_start(value.settings.start, event)?;
            event += chunks.len();
        }
        self.operations.push(Operation::Remove {
            slide: fields.slide,
            build: identifier,
        });
        Ok(self)
    }

    /// Apply every staged operation in one workspace, publish one candidate,
    /// reopen it and read the touched builds back.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Verification`] when the candidate does not read back
    /// as staged, [`Refusal::UnsupportedShape`] for unsupported native
    /// shapes, and typed source and limit errors.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_operations(self.source, &self.operations)
    }
}

fn lift(error: StructureError) -> Error {
    match error {
        StructureError::IdentifierOverflow => Error::Limit {
            kind: LimitKind::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        },
        _ => Error::InvalidSource,
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::Objects,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => Error::InvalidSource,
    }
}

fn physical_source(package: &Package) -> Result<Arc<[u8]>, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) if source.source_is_exact() => Ok(source.shared_source()),
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => {
            Err(Error::Refused(Refusal::UnsupportedSource))
        },
    }
}

/// Rewrite the slide's build and chunk references, keeping the header's
/// aggregate in step (host `patch_slide_build_references`).
fn patch_slide_references(
    workspace: &mut Workspace<'_>,
    slide: u64,
    remove: &[u64],
    add: Option<(u64, u64)>,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let payload = workspace
        .payload(slide, SLIDE_MESSAGE_TYPE)
        .map_err(lift)?
        .to_vec();
    let mut fields = Fields::parse(&payload, wire).map_err(lift)?;
    for number in [SLIDE_BUILDS, SLIDE_BUILD_CHUNKS] {
        let mut values = Vec::new();
        for field in fields.all(number) {
            let raw = field.message().map_err(lift)?.to_vec();
            if !remove.contains(&wire::reference_identifier(&raw, wire).map_err(lift)?) {
                values.push(raw);
            }
        }
        if let Some((build, chunk)) = add {
            values.push(wire::reference(if number == SLIDE_BUILDS {
                build
            } else {
                chunk
            }));
        }
        fields.replace_all(number, LENGTH_DELIMITED, &values);
    }
    let limits = workspace.archive_limits();
    let object = workspace.object_mut(slide).map_err(lift)?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    object
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data: fields.encode(),
            },
            limits,
        )
        .map_err(|error| lift(error.into()))?;
    let info = &mut object.archive_info.message_infos[index];
    info.object_references.retain(|id| !remove.contains(id));
    for field in &mut info.field_infos {
        field.object_references.retain(|id| !remove.contains(id));
    }
    if let Some((build, chunk)) = add {
        for id in [build, chunk] {
            if !info.object_references.contains(&id) {
                info.object_references.push(id);
            }
        }
    }
    Ok(())
}

/// The slide node's build-event cache (host `patch_slide_build_cache`).
fn patch_node_cache(workspace: &mut Workspace<'_>, node: u64, events: usize) -> Result<(), Error> {
    let wire = workspace.wire;
    let count = u64::try_from(events).map_err(|_| Error::InvalidSource)?;
    let payload = workspace
        .payload(node, SLIDE_NODE_MESSAGE_TYPE)
        .map_err(lift)?
        .to_vec();
    let mut fields = Fields::parse(&payload, wire).map_err(lift)?;
    if count == 0 {
        fields.remove(NODE_BUILD_EVENT_COUNT);
    } else {
        fields.set_varint(NODE_BUILD_EVENT_COUNT, count);
    }
    fields.set_varint(
        NODE_BUILD_EVENT_CACHE_VERSION,
        if count == 0 { u64::from(u32::MAX) } else { 2 },
    );
    fields.set_varint(NODE_HAS_EXPLICIT_BUILDS, u64::from(count != 0));
    fields.set_varint(NODE_EXPLICIT_BUILDS_CACHE_VERSION, 2);
    replace(workspace, node, SLIDE_NODE_MESSAGE_TYPE, fields.encode())
}

fn replace(
    workspace: &mut Workspace<'_>,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
) -> Result<(), Error> {
    let limits = workspace.archive_limits();
    let object = workspace.object_mut(identifier).map_err(lift)?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or(Error::InvalidSource)?;
    object
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(|error| lift(error.into()))?;
    Ok(())
}

fn slide_node(source: &Package, slide: u64) -> Result<u64, Error> {
    let count = source.show().map_err(map_read_error)?.slides().len();
    for position in 0..count {
        let record = source
            .slide_record_at(position)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource)?;
        if record.slide_identifier == slide {
            return Ok(record.node_identifier);
        }
    }
    Err(Error::InvalidSource)
}

fn commit_operations(source: &Package, operations: &[Operation]) -> Result<Commit, Error> {
    let source_bytes = physical_source(source)?;
    if operations.is_empty() {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
            },
            diagnostics: Diagnostics::default(),
        });
    }
    let mut workspace = Workspace::new(source).map_err(lift)?;
    let wire = workspace.wire;
    let mut diagnostics = Diagnostics {
        operations: operations.len(),
        ..Diagnostics::default()
    };
    let mut touched = Vec::new();
    for operation in operations {
        match operation {
            Operation::Add {
                slide,
                drawable,
                settings,
            } => {
                let member = workspace.component_of(*slide).ok_or(Error::InvalidSource)?;
                let build = workspace.allocate().map_err(lift)?;
                let chunk = workspace.allocate().map_err(lift)?;
                let uuid = workspace.clone_uuid(None, chunk).map_err(lift)?;
                let registered = workspace.clone_uuid(None, build).map_err(lift)?;
                #[allow(
                    clippy::cast_possible_truncation,
                    reason = "the seed is the low 32 bits"
                )]
                let seed = uuid.lower as u32;
                let limits = workspace.archive_limits();
                for (identifier, message_type, data) in [
                    (
                        build,
                        BUILD_MESSAGE_TYPE,
                        codec::encode_build(*drawable, settings, seed),
                    ),
                    (
                        chunk,
                        BUILD_CHUNK_MESSAGE_TYPE,
                        codec::encode_chunk(build, (uuid.lower, uuid.upper), settings),
                    ),
                ] {
                    let object = ArchiveObject::new(
                        identifier,
                        vec![RawMessage {
                            type_: message_type,
                            data,
                        }],
                        limits,
                    )
                    .map_err(|error| lift(error.into()))?;
                    workspace.insert(&member, object).map_err(lift)?;
                }
                patch_slide_references(&mut workspace, *slide, &[], Some((build, chunk)))?;
                let component = workspace
                    .registry
                    .component_for_member(&member)
                    .ok_or(Error::InvalidSource)?
                    .identifier;
                workspace
                    .registry
                    .component_mut(component)
                    .ok_or(Error::InvalidSource)?
                    .add_object_uuid(build, registered);
                diagnostics.created_builds += 1;
                touched.push(*slide);
            },
            Operation::Set {
                slide,
                build,
                settings,
            } => {
                let object = workspace.require(*build).map_err(lift)?;
                if object.messages.iter().any(|message| message.type_ == 0) {
                    return Err(Error::Refused(Refusal::UnsupportedShape));
                }
                let payload = workspace
                    .payload(*build, BUILD_MESSAGE_TYPE)
                    .map_err(lift)?
                    .to_vec();
                let patched = codec::patch_build(&payload, settings, wire).map_err(lift)?;
                replace(&mut workspace, *build, BUILD_MESSAGE_TYPE, patched)?;
                let slide_payload = workspace
                    .payload(*slide, SLIDE_MESSAGE_TYPE)
                    .map_err(lift)?
                    .to_vec();
                let fields = Fields::parse(&slide_payload, wire).map_err(lift)?;
                let mut chunks = Vec::new();
                for field in fields.all(SLIDE_BUILD_CHUNKS) {
                    chunks.push(
                        wire::reference_identifier(field.message().map_err(lift)?, wire)
                            .map_err(lift)?,
                    );
                }
                for chunk in chunks {
                    let payload = workspace
                        .payload(chunk, BUILD_CHUNK_MESSAGE_TYPE)
                        .map_err(lift)?
                        .to_vec();
                    if codec::read_chunk(&payload, wire).map_err(lift)?.0 == Some(*build) {
                        let patched = codec::patch_chunk(&payload, settings, wire).map_err(lift)?;
                        replace(&mut workspace, chunk, BUILD_CHUNK_MESSAGE_TYPE, patched)?;
                    }
                }
                diagnostics.updated_builds += 1;
                touched.push(*slide);
            },
            Operation::Remove { slide, build } => {
                let slide_payload = workspace
                    .payload(*slide, SLIDE_MESSAGE_TYPE)
                    .map_err(lift)?
                    .to_vec();
                let fields = Fields::parse(&slide_payload, wire).map_err(lift)?;
                let mut chunks = Vec::new();
                for field in fields.all(SLIDE_BUILD_CHUNKS) {
                    let chunk = wire::reference_identifier(field.message().map_err(lift)?, wire)
                        .map_err(lift)?;
                    let payload = workspace
                        .payload(chunk, BUILD_CHUNK_MESSAGE_TYPE)
                        .map_err(lift)?;
                    if codec::read_chunk(payload, wire).map_err(lift)?.0 == Some(*build) {
                        chunks.push(chunk);
                    }
                }
                let mut removed = chunks.clone();
                removed.push(*build);
                patch_slide_references(&mut workspace, *slide, &removed, None)?;
                let member = workspace.component_of(*slide).ok_or(Error::InvalidSource)?;
                for object in &removed {
                    workspace.remove_object(*object).map_err(lift)?;
                }
                let component = workspace
                    .registry
                    .component_for_member(&member)
                    .ok_or(Error::InvalidSource)?
                    .identifier;
                if workspace
                    .registry
                    .component(component)
                    .is_some_and(|component| component.uuid_of(*build).is_some())
                {
                    workspace
                        .registry
                        .component_mut(component)
                        .ok_or(Error::InvalidSource)?
                        .remove_object_uuids(&[*build]);
                }
                diagnostics.removed_builds += 1;
                touched.push(*slide);
            },
        }
    }
    touched.sort_unstable();
    touched.dedup();
    for slide in &touched {
        let payload = workspace
            .payload(*slide, SLIDE_MESSAGE_TYPE)
            .map_err(lift)?
            .to_vec();
        let events = Fields::parse(&payload, wire)
            .map_err(lift)?
            .all(SLIDE_BUILD_CHUNKS)
            .count();
        if operations
            .iter()
            .any(|operation| !matches!(operation, Operation::Set { .. }))
        {
            let node = slide_node(source, *slide)?;
            patch_node_cache(&mut workspace, node, events)?;
        }
    }
    if diagnostics.removed_builds > 0 {
        workspace.release_removed_identifier_suffix();
    }
    let bytes = workspace.publish().map_err(lift)?;
    let candidate = source.structure_reopen(bytes).map_err(map_read_error)?;
    // Read every touched build back.
    let read = candidate_builds(&candidate)?;
    for operation in operations {
        let (build, settings) = match operation {
            Operation::Set {
                build, settings, ..
            } => (Some(*build), settings),
            Operation::Add { settings, .. } => (None, settings),
            Operation::Remove { build, .. } => {
                if read.iter().any(|(identifier, _)| identifier == build) {
                    return Err(Error::Verification);
                }
                continue;
            },
        };
        let found = read.iter().any(|(identifier, value)| {
            build.is_none_or(|build| *identifier == build)
                && value.settings.stage == settings.stage
                && value.settings.effect == settings.effect
                && value.settings.duration == settings.duration
                && value.settings.delay == settings.delay
                && value.settings.start == settings.start
        });
        if !found {
            return Err(Error::Verification);
        }
    }
    let target = physical_source(&candidate)?;
    let patch = Patch {
        artifacts: ExactArtifacts::new(source_bytes, target),
    };
    diagnostics.changed = !patch.is_noop();
    Ok(Commit {
        package: candidate,
        patch,
        diagnostics,
    })
}

fn candidate_builds(package: &Package) -> Result<Vec<(u64, Build)>, Error> {
    let count = package.show().map_err(map_read_error)?.slides().len();
    let mut all = Vec::new();
    for position in 0..count {
        let fields = read_slide(package, SlideSelector::index(position))?;
        all.extend(
            read_builds(package, &fields)?
                .into_iter()
                .map(|(identifier, build, _)| (identifier, build)),
        );
    }
    Ok(all)
}

/// Exact-source-checked reversible slide-build patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Return whether source and target are byte-identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Compact evidence for one verified slide-build commit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    operations: usize,
    created_builds: usize,
    updated_builds: usize,
    removed_builds: usize,
}

impl Diagnostics {
    /// Return whether the package changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return how many operations were applied.
    #[must_use]
    pub const fn operations(self) -> usize {
        self.operations
    }

    /// Return how many builds were created.
    #[must_use]
    pub const fn created_builds(self) -> usize {
        self.created_builds
    }

    /// Return how many builds were updated.
    #[must_use]
    pub const fn updated_builds(self) -> usize {
        self.updated_builds
    }

    /// Return how many builds were removed.
    #[must_use]
    pub const fn removed_builds(self) -> usize {
        self.removed_builds
    }
}

/// The verified result of one slide-build edit.
#[must_use]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the committed package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Take the committed package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Split the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

impl Package {
    /// Read a slide's builds in build order (host `slide_builds`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Error::InvalidSource`] for a
    /// malformed build graph.
    pub fn slide_build_list<'s>(
        &self,
        slide: impl Into<SlideSelector<'s>>,
    ) -> Result<Box<[Build]>, Error> {
        let fields = read_slide(self, slide.into())?;
        Ok(read_builds(self, &fields)?
            .into_iter()
            .map(|(_, build, _)| build)
            .collect())
    }

    /// Start one slide-build edit (add, set and remove builds).
    #[must_use]
    pub const fn edit_slide_builds(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Apply a slide-build patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the
    /// patch's exact source.
    pub fn apply_slide_builds(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let candidate = self
            .structure_reopen(patch.artifacts.target().to_vec())
            .map_err(map_read_error)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                ..Diagnostics::default()
            },
        })
    }
}
