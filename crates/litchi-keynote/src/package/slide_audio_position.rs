//! Exact-source, selector-first position transactions for slide-owned audio.
//!
//! Audio controls use the same native movie envelope as file-backed movies,
//! but their displayed size is commonly zero (and may be absent).  This owner
//! therefore rewrites only `geometry.position`; it never constructs a
//! positive-size [`crate::slide::media::geometry::MovieGeometry`] or touches size, flags, angle, playback,
//! assets, comments, or other native fields.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The semantic boundary redacts lower-layer failure details."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_protos::keynote_movie_geometry_codec;
use thiserror::Error;

use super::slide_movie_geometry::{
    GeometryBudget, GeometrySelection, codec_options, map_geometry_codec_error, physical_catalog,
    previews_absent, rewrite_geometry_message, select_audio_with_budget, verify_locality,
};
use super::{Package, ReadError, SemanticLimitKind};
use crate::slide::media::Point;
use crate::{MovieSelector, SlideSelector};

/// Resource categories reported by an audio-position transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideAudioPositionLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
    /// Bytes in the selected audio movie payload.
    PositionBytes,
    /// Logical transaction allocations.
    Allocations,
    /// Retained transaction bytes.
    Retained,
    /// Scratch transaction bytes.
    Scratch,
    /// Rewritten physical components.
    Components,
}

impl fmt::Display for SlideAudioPositionLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::PositionBytes => "audio position bytes",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
        })
    }
}

/// Content-redacted failure raised by an audio-position transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideAudioPositionError {
    /// The source was not retained as an exact physical package.
    #[error("this Keynote source does not support physical audio-position edits")]
    UnsupportedSource,
    /// The selected audio graph is not safely owned by one component.
    #[error("the requested Keynote audio-position graph is unsupported")]
    UnsupportedDependency,
    /// A selector matched more than one semantic object.
    #[error("the Keynote audio-position selector is ambiguous")]
    AmbiguousSelector,
    /// The slide name selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide matched an exact name selector.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// No slide existed at a checked position.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// No audio control existed at a checked source-order position.
    #[error("the selected Keynote slide has no audio at position {position:?}")]
    MoviePositionNotFound { position: Position },
    /// The selected media graph or payload was malformed.
    #[error("the Keynote audio-position source is invalid")]
    InvalidSource,
    /// A supplied or decoded coordinate was not finite.
    #[error("Keynote audio position must contain finite coordinates")]
    InvalidPosition,
    /// A finite operation resource ceiling was exceeded.
    #[error("Keynote audio position {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: SlideAudioPositionLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote audio-position transaction")]
    Allocation { amount: usize },
    /// Candidate reopening did not reproduce the staged position and locality.
    #[error("the edited Keynote audio position failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote audio-position patch does not match the exact source package")]
    PatchConflict,
}

impl From<super::slide_movie_geometry::SlideMovieGeometryError> for SlideAudioPositionError {
    fn from(error: super::slide_movie_geometry::SlideMovieGeometryError) -> Self {
        use super::slide_movie_geometry::{SlideMovieGeometryError, SlideMovieGeometryLimitKind};
        match error {
            SlideMovieGeometryError::UnsupportedSource => Self::UnsupportedSource,
            SlideMovieGeometryError::UnsupportedDependency => Self::UnsupportedDependency,
            // Position-only audio edits preserve the native lock bit.  The
            // shared geometry selector can still report this for the
            // file-movie contract, but it is not an audio-position outcome.
            SlideMovieGeometryError::Locked => Self::UnsupportedDependency,
            SlideMovieGeometryError::AmbiguousSelector => Self::AmbiguousSelector,
            SlideMovieGeometryError::EmptySlideName => Self::EmptySlideName,
            SlideMovieGeometryError::SlideNameNotFound => Self::SlideNameNotFound,
            SlideMovieGeometryError::SlidePositionNotFound { position } => {
                Self::SlidePositionNotFound { position }
            },
            SlideMovieGeometryError::MoviePositionNotFound { position } => {
                Self::MoviePositionNotFound { position }
            },
            SlideMovieGeometryError::InvalidSource => Self::InvalidSource,
            SlideMovieGeometryError::LimitExceeded {
                kind,
                observed,
                maximum,
            } => Self::LimitExceeded {
                kind: match kind {
                    SlideMovieGeometryLimitKind::InputBytes => {
                        SlideAudioPositionLimitKind::InputBytes
                    },
                    SlideMovieGeometryLimitKind::OutputBytes => {
                        SlideAudioPositionLimitKind::OutputBytes
                    },
                    SlideMovieGeometryLimitKind::WireBytes => {
                        SlideAudioPositionLimitKind::WireBytes
                    },
                    SlideMovieGeometryLimitKind::Entries => SlideAudioPositionLimitKind::Entries,
                    SlideMovieGeometryLimitKind::EntryBytes => {
                        SlideAudioPositionLimitKind::EntryBytes
                    },
                    SlideMovieGeometryLimitKind::TotalBytes => {
                        SlideAudioPositionLimitKind::TotalBytes
                    },
                    SlideMovieGeometryLimitKind::Slides => SlideAudioPositionLimitKind::Slides,
                    SlideMovieGeometryLimitKind::References => {
                        SlideAudioPositionLimitKind::References
                    },
                    SlideMovieGeometryLimitKind::WireFields => {
                        SlideAudioPositionLimitKind::WireFields
                    },
                    SlideMovieGeometryLimitKind::WireNesting => {
                        SlideAudioPositionLimitKind::WireNesting
                    },
                    SlideMovieGeometryLimitKind::WireWork => SlideAudioPositionLimitKind::WireWork,
                    SlideMovieGeometryLimitKind::GeometryBytes => {
                        SlideAudioPositionLimitKind::PositionBytes
                    },
                    SlideMovieGeometryLimitKind::Allocations => {
                        SlideAudioPositionLimitKind::Allocations
                    },
                    SlideMovieGeometryLimitKind::Retained => SlideAudioPositionLimitKind::Retained,
                    SlideMovieGeometryLimitKind::Scratch => SlideAudioPositionLimitKind::Scratch,
                    SlideMovieGeometryLimitKind::Components => {
                        SlideAudioPositionLimitKind::Components
                    },
                },
                observed,
                maximum,
            },
            SlideMovieGeometryError::Allocation { amount } => Self::Allocation { amount },
            SlideMovieGeometryError::Verification => Self::Verification,
            SlideMovieGeometryError::PatchConflict => Self::PatchConflict,
        }
    }
}

/// One mutable audio position staged against an immutable package snapshot.
pub struct SlideAudioPositionEdit<'a> {
    source: &'a Package,
    budget: GeometryBudget,
    selection: GeometrySelection,
    before: Point,
    after: Point,
}

impl fmt::Debug for SlideAudioPositionEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideAudioPositionEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("audio_position", &self.selection.movie_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish()
    }
}

impl<'a> SlideAudioPositionEdit<'a> {
    fn new<'slide>(
        source: &'a Package,
        slide: impl Into<SlideSelector<'slide>>,
        audio: impl Into<MovieSelector>,
    ) -> Result<Self, SlideAudioPositionError> {
        let mut budget = GeometryBudget::new(source)?;
        let source_bytes = physical_catalog(source)?.source_bytes().len();
        budget.source(source_bytes)?;
        let selection = select_audio_with_budget(source, slide.into(), audio.into(), &mut budget)?;
        let before = selection
            .before_position
            .ok_or(SlideAudioPositionError::InvalidSource)?;
        Ok(Self {
            source,
            budget,
            selection,
            before,
            after: before,
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected audio position within the slide's source order.
    #[must_use]
    pub const fn audio_position(&self) -> Position {
        self.selection.movie_position
    }

    /// Return the finite position observed when this edit began.
    #[must_use]
    pub const fn before(&self) -> Point {
        self.before
    }

    /// Return the finite position staged for publication.
    #[must_use]
    pub const fn after(&self) -> Point {
        self.after
    }

    /// Stage a finite audio position.
    pub fn set(mut self, position: Point) -> Result<Self, SlideAudioPositionError> {
        validate_position(position)?;
        self.after = position;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<SlideAudioPositionCommit, SlideAudioPositionError> {
        commit_edit(
            self.source,
            &self.selection,
            self.before,
            self.after,
            self.budget,
        )
    }
}

/// An exact-source checked reversible audio-position patch.
#[derive(Clone, PartialEq)]
pub struct SlideAudioPositionPatch {
    artifacts: ExactArtifacts,
    selection: GeometrySelection,
    before: Point,
    after: Point,
    deleted_previews: usize,
    restored_previews: usize,
    source_previews_absent: bool,
    target_previews_absent: bool,
}

impl fmt::Debug for SlideAudioPositionPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideAudioPositionPatch")
            .field("slide_position", &self.selection.slide_position)
            .field("audio_position", &self.selection.movie_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideAudioPositionPatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected audio position within the slide's source order.
    #[must_use]
    pub const fn audio_position(&self) -> Position {
        self.selection.movie_position
    }

    /// Return the position required from the source package.
    #[must_use]
    pub const fn before(&self) -> Point {
        self.before
    }

    /// Return the position produced by this patch.
    #[must_use]
    pub const fn after(&self) -> Point {
        self.after
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch preserves exact source bytes and state.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            before: self.after,
            after: self.before,
            deleted_previews: self.restored_previews,
            restored_previews: self.deleted_previews,
            source_previews_absent: self.target_previews_absent,
            target_previews_absent: self.source_previews_absent,
        }
    }
}

/// Compact evidence describing one audio-position publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideAudioPositionDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideAudioPositionDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return how many stale root rendering previews were deleted.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one audio-position transaction.
#[must_use = "an audio-position commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideAudioPositionCommit {
    package: Package,
    patch: SlideAudioPositionPatch,
    diagnostics: SlideAudioPositionDiagnostics,
}

impl SlideAudioPositionCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideAudioPositionPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideAudioPositionDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the finite position of one slide-owned audio control.
    ///
    /// `audio` is a source-order movie selector: every movie archive sibling
    /// on the selected slide occupies a position, including file-backed movie
    /// siblings.  Selecting a file-backed sibling therefore reports a
    /// wrong-kind graph instead of renumbering the audio controls.
    pub fn slide_audio_position<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        audio: impl Into<MovieSelector>,
    ) -> Result<Point, SlideAudioPositionError> {
        let mut budget = GeometryBudget::new(self)?;
        let source_bytes = physical_catalog(self)?.source_bytes().len();
        budget.source(source_bytes)?;
        select_audio_with_budget(self, slide.into(), audio.into(), &mut budget)?
            .before_position
            .ok_or(SlideAudioPositionError::InvalidSource)
    }

    /// Begin an exact immutable edit of one selected audio position.
    ///
    /// The movie selector uses the same source order as
    /// [`Package::slide_audio_position`], counting all movie archive
    /// siblings on the slide.
    pub fn edit_slide_audio_position<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        audio: impl Into<MovieSelector>,
    ) -> Result<SlideAudioPositionEdit<'_>, SlideAudioPositionError> {
        SlideAudioPositionEdit::new(self, slide, audio)
    }

    /// Apply an exact-source checked audio-position patch.
    pub fn apply_slide_audio_position(
        &self,
        patch: &SlideAudioPositionPatch,
    ) -> Result<SlideAudioPositionCommit, SlideAudioPositionError> {
        let mut budget = GeometryBudget::new(self)?;
        let catalog = physical_catalog(self)?;
        let source_len = catalog.source_bytes().len();
        budget.source(source_len)?;
        budget.work(source_len)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(SlideAudioPositionError::PatchConflict);
        }
        if previews_absent(self)? != patch.source_previews_absent {
            return Err(SlideAudioPositionError::PatchConflict);
        }
        let current = select_audio_with_budget(
            self,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection)
            || current.before_position != Some(patch.before)
        {
            return Err(SlideAudioPositionError::PatchConflict);
        }
        if patch.is_noop() {
            budget.validate_package(self)?;
            return Ok(SlideAudioPositionCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideAudioPositionDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(SlideAudioPositionError::UnsupportedSource);
        }
        reopen_target_patch(self, patch, &mut budget)
    }
}

fn validate_position(position: Point) -> Result<(), SlideAudioPositionError> {
    if position.x.is_finite() && position.y.is_finite() {
        Ok(())
    } else {
        Err(SlideAudioPositionError::InvalidPosition)
    }
}

fn same_selection(left: &GeometrySelection, right: &GeometrySelection) -> bool {
    left.slide_position == right.slide_position
        && left.movie_position == right.movie_position
        && left.slide_identifier == right.slide_identifier
        && left.node_identifier == right.node_identifier
        && left.movie_identifier == right.movie_identifier
        && left.message_index == right.message_index
        && left.slide_component_name == right.slide_component_name
        && left.locked == right.locked
}

fn commit_edit(
    source: &Package,
    selection: &GeometrySelection,
    before: Point,
    after: Point,
    mut budget: GeometryBudget,
) -> Result<SlideAudioPositionCommit, SlideAudioPositionError> {
    validate_position(after)?;
    if before == after {
        let catalog = physical_catalog(source)?;
        budget.validate_package(source)?;
        budget.work(source.source_bytes().len())?;
        let bytes = catalog.shared_source();
        let previews_absent = previews_absent(source)?;
        return Ok(SlideAudioPositionCommit {
            package: source.snapshot(),
            patch: SlideAudioPositionPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before,
                after,
                deleted_previews: 0,
                restored_previews: 0,
                source_previews_absent: previews_absent,
                target_previews_absent: previews_absent,
            },
            diagnostics: SlideAudioPositionDiagnostics::unchanged(),
        });
    }
    let catalog = physical_catalog(source)?;
    if !catalog.source_is_exact() {
        return Err(SlideAudioPositionError::UnsupportedSource);
    }
    let current = select_audio_with_budget(
        source,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
        &mut budget,
    )?;
    if !same_selection(&current, selection) || current.before_position != Some(before) {
        return Err(SlideAudioPositionError::PatchConflict);
    }
    let (candidate, deleted_previews) = rewrite_geometry_message(
        source,
        selection.slide_component_name.as_ref(),
        selection.movie_identifier,
        selection.message_index,
        |original, budget| {
            let options = codec_options(source, original, budget)?;
            let write =
                keynote_movie_geometry_codec::MoviePositionWrite::from_values(after.x, after.y);
            let prepared = keynote_movie_geometry_codec::prepare_movie_position_rewrite(
                original, write, options,
            )
            .map_err(map_geometry_codec_error)?;
            budget.codec_report(prepared.prepare_report())?;
            let requirements = prepared.execution_requirements();
            budget.codec_requirements(requirements)?;
            prepared
                .execute(requirements.exact_limits())
                .map_err(map_geometry_codec_error)
                .map(|output| output.into_output())
        },
        &mut budget,
    )?;
    budget.validate_package(&candidate)?;
    if !previews_absent(&candidate)? {
        return Err(SlideAudioPositionError::Verification);
    }
    let selected = select_audio_with_budget(
        &candidate,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
        &mut budget,
    )?;
    if !same_selection(&selected, selection) || selected.before_position != Some(after) {
        return Err(SlideAudioPositionError::Verification);
    }
    verify_locality(
        source,
        &candidate,
        selection.slide_component_name.as_ref(),
        selection.movie_identifier,
        selection.message_index,
        true,
        &mut budget,
    )?;
    let source_bytes = catalog.shared_source();
    let target = physical_catalog(&candidate)?.shared_source();
    Ok(SlideAudioPositionCommit {
        package: candidate,
        patch: SlideAudioPositionPatch {
            artifacts: ExactArtifacts::new(source_bytes, target),
            selection: selection.clone(),
            before,
            after,
            deleted_previews,
            restored_previews: 0,
            source_previews_absent: previews_absent(source)?,
            target_previews_absent: true,
        },
        diagnostics: SlideAudioPositionDiagnostics::published(deleted_previews),
    })
}

fn reopen_target_patch(
    source: &Package,
    patch: &SlideAudioPositionPatch,
    budget: &mut GeometryBudget,
) -> Result<SlideAudioPositionCommit, SlideAudioPositionError> {
    budget.candidate_reopen(patch.artifacts.target().len())?;
    budget.allocations(1)?;
    budget.retained(patch.artifacts.target().len())?;
    budget.scratch(patch.artifacts.target().len())?;
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    budget.validate_package(&candidate)?;
    if previews_absent(&candidate)? != patch.target_previews_absent {
        return Err(SlideAudioPositionError::Verification);
    }
    let selected = select_audio_with_budget(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        MovieSelector::position(patch.selection.movie_position),
        budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before_position != Some(patch.after)
    {
        return Err(SlideAudioPositionError::Verification);
    }
    verify_locality(
        source,
        &candidate,
        patch.selection.slide_component_name.as_ref(),
        patch.selection.movie_identifier,
        patch.selection.message_index,
        patch.target_previews_absent,
        budget,
    )?;
    Ok(SlideAudioPositionCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideAudioPositionDiagnostics::published(patch.deleted_previews),
    })
}

fn map_read_error(error: ReadError) -> SlideAudioPositionError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideAudioPositionError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Slides => SlideAudioPositionLimitKind::Slides,
                SemanticLimitKind::References => SlideAudioPositionLimitKind::References,
                _ => SlideAudioPositionLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideAudioPositionError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideAudioPositionLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideAudioPositionLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideAudioPositionLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideAudioPositionLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideAudioPositionError::Allocation { amount },
        _ => SlideAudioPositionError::InvalidSource,
    }
}
