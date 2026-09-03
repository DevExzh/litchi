//! Exact-source, selector-first Keynote movie/audio-playback transactions.
//!
//! This owner deliberately handles the scalar playback edge only. It resolves
//! a rooted file-backed movie or independently positioned audio control,
//! delegates the MovieArchive wire projection to the neutral strict codec, and
//! rewrites the selected media component in a private candidate before
//! reopening it through the normal Keynote ingress. Movie/audio graph
//! creation/removal, media replacement, geometry, and metadata allocation are
//! outside this module.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The semantic boundary deliberately redacts lower-layer failures."
)]

use std::fmt;
use std::sync::Arc;
use std::time::Duration;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len, wire::WireView};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::movie_playback_codec;
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::slide::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};
use crate::{MovieKind, MovieSelector, SlideSelector};

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const MOVIE_SUPER_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;

#[derive(Debug, Clone, Copy)]
struct PlaybackBudget {
    max_input: usize,
    max_output: usize,
    max_fields: usize,
    max_work: usize,
    max_nesting: usize,
    max_references: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    max_components: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    nesting: usize,
    references: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
    components: usize,
}

impl PlaybackBudget {
    fn for_package(package: &Package) -> Result<Self, SlideMoviePlaybackError> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let physical = package.state.options.archive();
        let source = usize::try_from(physical.max_input_bytes())
            .map_err(|_| SlideMoviePlaybackError::InvalidSource)?;
        let aggregate = source
            .checked_mul(4)
            .ok_or(SlideMoviePlaybackError::InvalidSource)?;
        let semantic = package.semantic_limits();
        Ok(Self {
            max_input: aggregate,
            max_output: aggregate,
            max_fields: wire.max_fields(),
            max_work: wire.max_rewrite_work(),
            max_nesting: wire.max_nesting(),
            max_references: semantic.max_references(),
            max_allocations: semantic
                .max_objects()
                .checked_add(semantic.max_references())
                .ok_or(SlideMoviePlaybackError::InvalidSource)?,
            max_retained: aggregate,
            max_scratch: aggregate,
            max_components: semantic.max_objects(),
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
            components: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: SlideMoviePlaybackLimitKind,
    ) -> Result<(), SlideMoviePlaybackError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SlideMoviePlaybackError::InvalidSource)?;
        if observed > maximum {
            return Err(SlideMoviePlaybackError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn source(&mut self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            SlideMoviePlaybackLimitKind::InputBytes,
        )
    }

    fn wire_scan(&mut self, payload: &[u8], nesting: usize) -> Result<(), SlideMoviePlaybackError> {
        self.source(payload.len())?;
        self.fields(payload.len())?;
        self.work(payload.len())?;
        self.nesting(nesting)
    }

    fn output(&mut self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            SlideMoviePlaybackLimitKind::OutputBytes,
        )
    }

    fn preflight_output(&self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        let observed = self
            .output
            .checked_add(amount)
            .ok_or(SlideMoviePlaybackError::InvalidSource)?;
        if observed > self.max_output {
            return Err(SlideMoviePlaybackError::LimitExceeded {
                kind: SlideMoviePlaybackLimitKind::OutputBytes,
                observed: observed as u64,
                maximum: self.max_output as u64,
            });
        }
        Ok(())
    }

    fn preflight_work(&self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        let observed = self
            .work
            .checked_add(amount)
            .ok_or(SlideMoviePlaybackError::InvalidSource)?;
        if observed > self.max_work {
            return Err(SlideMoviePlaybackError::LimitExceeded {
                kind: SlideMoviePlaybackLimitKind::WireWork,
                observed: observed as u64,
                maximum: self.max_work as u64,
            });
        }
        Ok(())
    }

    fn fields(&mut self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        Self::add(
            &mut self.fields,
            amount,
            self.max_fields,
            SlideMoviePlaybackLimitKind::WireFields,
        )
    }

    fn work(&mut self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            SlideMoviePlaybackLimitKind::WireWork,
        )
    }

    fn nesting(&mut self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        if amount > self.max_nesting {
            return Err(SlideMoviePlaybackError::LimitExceeded {
                kind: SlideMoviePlaybackLimitKind::WireNesting,
                observed: amount as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(amount);
        Ok(())
    }

    fn references(&mut self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SlideMoviePlaybackLimitKind::References,
        )
    }

    fn allocations(&mut self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SlideMoviePlaybackLimitKind::Allocations,
        )
    }

    fn retained(&mut self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SlideMoviePlaybackLimitKind::Retained,
        )
    }

    fn scratch(&mut self, amount: usize) -> Result<(), SlideMoviePlaybackError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            SlideMoviePlaybackLimitKind::Scratch,
        )
    }

    fn component(&mut self) -> Result<(), SlideMoviePlaybackError> {
        Self::add(
            &mut self.components,
            1,
            self.max_components,
            SlideMoviePlaybackLimitKind::Components,
        )
    }

    fn residual_wire_limits(
        &self,
        base: litchi_iwa_common::WireLimits,
    ) -> Result<litchi_iwa_common::WireLimits, SlideMoviePlaybackError> {
        let input = self.max_input.saturating_sub(self.input).max(1);
        let fields = self.max_fields.saturating_sub(self.fields).max(1);
        let work = self.max_work.saturating_sub(self.work).max(1);
        base.with_input_bytes(base.max_input_bytes().min(input))
            .and_then(|value| value.with_fields(base.max_fields().min(fields)))
            .and_then(|value| value.with_rewrite_work(base.max_rewrite_work().min(work)))
            .and_then(|value| value.with_nesting(base.max_nesting().min(self.max_nesting)))
            .map_err(map_wire_error)
    }

    fn codec_report(
        &mut self,
        report: movie_playback_codec::DecodeReport,
    ) -> Result<(), SlideMoviePlaybackError> {
        self.source(report.input_bytes())?;
        Self::add(
            &mut self.fields,
            report.fields(),
            self.max_fields,
            SlideMoviePlaybackLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            report.work_bytes(),
            self.max_work,
            SlideMoviePlaybackLimitKind::WireWork,
        )?;
        self.scratch(report.scratch_bytes())?;
        self.retained(report.input_bytes())?;
        self.nesting(report.max_depth() as usize)?;
        self.allocations(report.allocations())?;
        Ok(())
    }

    fn codec_requirements(
        &mut self,
        requirements: movie_playback_codec::RewriteExecutionRequirements,
    ) -> Result<(), SlideMoviePlaybackError> {
        self.output(requirements.output_bytes)?;
        Self::add(
            &mut self.fields,
            requirements.fields,
            self.max_fields,
            SlideMoviePlaybackLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            requirements.work_bytes,
            self.max_work,
            SlideMoviePlaybackLimitKind::WireWork,
        )?;
        self.allocations(requirements.allocations)?;
        self.retained(requirements.retained_bytes)?;
        self.scratch(requirements.scratch_bytes)?;
        self.nesting(requirements.max_depth as usize)?;
        Ok(())
    }

    fn physical(&mut self, decompressed: usize) -> Result<(), SlideMoviePlaybackError> {
        self.source(decompressed)?;
        self.work(decompressed)?;
        self.output(
            decompressed
                .checked_mul(2)
                .ok_or(SlideMoviePlaybackError::InvalidSource)?,
        )
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SlideMoviePlaybackError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        Ok(())
    }

    fn candidate_reopen(&mut self, bytes: usize) -> Result<(), SlideMoviePlaybackError> {
        self.source(bytes)?;
        self.work(bytes)?;
        Ok(())
    }
}

/// Resource categories reported by a movie-playback transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideMoviePlaybackLimitKind {
    InputBytes,
    OutputBytes,
    WireBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    Slides,
    References,
    WireFields,
    WireNesting,
    WireWork,
    PlaybackBytes,
    Allocations,
    Retained,
    Scratch,
    Components,
}

impl fmt::Display for SlideMoviePlaybackLimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
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
            Self::PlaybackBytes => "movie playback bytes",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
        })
    }
}

/// Content-redacted failure raised by a movie-playback transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideMoviePlaybackError {
    #[error("this Keynote source does not support physical movie-playback edits")]
    UnsupportedSource,
    #[error("the requested Keynote movie-playback graph operation is unsupported")]
    UnsupportedDependency,
    #[error("the Keynote movie-playback selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no movie at position {position:?}")]
    MoviePositionNotFound { position: Position },
    #[error("the Keynote movie-playback source cannot be edited safely")]
    InvalidSource,
    #[error("Keynote movie-playback {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        kind: SlideMoviePlaybackLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote movie-playback transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote movie playback failed semantic verification")]
    Verification,
    #[error("the Keynote movie-playback patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable playback value staged against an immutable package snapshot.
pub struct SlideMoviePlaybackEdit<'a> {
    source: &'a Package,
    selection: PlaybackSelection,
    after: MediaPlaybackSettings,
}

impl fmt::Debug for SlideMoviePlaybackEdit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideMoviePlaybackEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field("has_before", &self.selection.before.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> SlideMoviePlaybackEdit<'a> {
    fn new<'slide>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
    ) -> Result<Self, SlideMoviePlaybackError> {
        let selection = select_movie(source, slide_selector.into(), movie_selector.into(), true)?;
        let after = selection
            .before
            .ok_or(SlideMoviePlaybackError::UnsupportedDependency)?;
        Ok(Self {
            source,
            selection,
            after,
        })
    }

    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }

    #[must_use]
    pub const fn before(&self) -> Option<MediaPlaybackSettings> {
        self.selection.before
    }

    #[must_use]
    pub const fn after(&self) -> MediaPlaybackSettings {
        self.after
    }

    pub fn set(mut self, settings: MediaPlaybackSettings) -> Result<Self, SlideMoviePlaybackError> {
        settings
            .validate()
            .map_err(|_| SlideMoviePlaybackError::InvalidSource)?;
        self.after = settings;
        Ok(self)
    }

    pub fn commit(self) -> Result<SlideMoviePlaybackCommit, SlideMoviePlaybackError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible playback patch.
#[derive(Clone, PartialEq)]
pub struct SlideMoviePlaybackPatch {
    artifacts: ExactArtifacts,
    selection: PlaybackSelection,
    before: MediaPlaybackSettings,
    after: MediaPlaybackSettings,
    touched_components: usize,
    deleted_previews: usize,
}

impl fmt::Debug for SlideMoviePlaybackPatch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideMoviePlaybackPatch")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field(
                "changed_poster",
                &(self.before.poster_time != self.after.poster_time),
            )
            .finish_non_exhaustive()
    }
}

impl SlideMoviePlaybackPatch {
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }

    #[must_use]
    pub const fn before(&self) -> MediaPlaybackSettings {
        self.before
    }

    #[must_use]
    pub const fn after(&self) -> MediaPlaybackSettings {
        self.after
    }

    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            before: self.after,
            after: self.before,
            touched_components: self.touched_components,
            deleted_previews: self.deleted_previews,
        }
    }
}

/// Compact playback publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideMoviePlaybackDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideMoviePlaybackDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one playback transaction.
#[must_use = "a Keynote movie-playback commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideMoviePlaybackCommit {
    package: Package,
    patch: SlideMoviePlaybackPatch,
    diagnostics: SlideMoviePlaybackDiagnostics,
}

impl SlideMoviePlaybackCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub const fn patch(&self) -> &SlideMoviePlaybackPatch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &SlideMoviePlaybackDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq)]
struct PlaybackSelection {
    slide_position: Position,
    movie_position: Position,
    slide_identifier: u64,
    node_identifier: u64,
    movie_identifier: u64,
    message_index: usize,
    slide_component_name: Arc<str>,
    before: Option<MediaPlaybackSettings>,
}

impl fmt::Debug for PlaybackSelection {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("PlaybackSelection")
            .field("slide_position", &self.slide_position)
            .field("movie_position", &self.movie_position)
            .field("has_before", &self.before.is_some())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Read playback settings for one existing file-backed movie or audio
    /// control.
    pub fn slide_movie_playback_settings<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
    ) -> Result<Option<MediaPlaybackSettings>, SlideMoviePlaybackError> {
        let mut budget = PlaybackBudget::for_package(self)?;
        Ok(select_movie_with_budget(
            self,
            slide_selector.into(),
            movie_selector.into(),
            false,
            &mut budget,
        )?
        .before)
    }

    /// Begin an exact immutable edit of an existing file-backed movie or audio
    /// control's playback.
    pub fn edit_slide_movie_playback_settings<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
    ) -> Result<SlideMoviePlaybackEdit<'_>, SlideMoviePlaybackError> {
        SlideMoviePlaybackEdit::new(self, slide_selector, movie_selector)
    }

    /// Apply an exact-source checked playback patch.
    pub fn apply_slide_movie_playback_settings(
        &self,
        patch: &SlideMoviePlaybackPatch,
    ) -> Result<SlideMoviePlaybackCommit, SlideMoviePlaybackError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(SlideMoviePlaybackError::PatchConflict);
        }
        let mut budget = PlaybackBudget::for_package(self)?;
        let current = select_movie_with_budget(
            self,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            true,
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != Some(patch.before) {
            return Err(SlideMoviePlaybackError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideMoviePlaybackCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideMoviePlaybackDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(SlideMoviePlaybackError::PatchConflict);
        }
        reopen_target_patch(self, patch, &mut budget)
    }
}

fn commit_edit(
    source: &Package,
    selection: &PlaybackSelection,
    after: MediaPlaybackSettings,
) -> Result<SlideMoviePlaybackCommit, SlideMoviePlaybackError> {
    let catalog = physical_catalog(source)?;
    let source_bytes = catalog.shared_source();
    let mut budget = PlaybackBudget::for_package(source)?;
    budget.source(source.source_bytes().len())?;
    let current = select_movie_with_budget(
        source,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
        true,
        &mut budget,
    )?;
    let before = selection
        .before
        .ok_or(SlideMoviePlaybackError::UnsupportedDependency)?;
    if !same_selection(&current, selection) || current.before != Some(before) {
        return Err(SlideMoviePlaybackError::InvalidSource);
    }
    if before == after {
        return Ok(SlideMoviePlaybackCommit {
            package: source.snapshot(),
            patch: SlideMoviePlaybackPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                selection: selection.clone(),
                before,
                after,
                touched_components: 0,
                deleted_previews: 0,
            },
            diagnostics: SlideMoviePlaybackDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(SlideMoviePlaybackError::UnsupportedSource);
    }
    let (candidate, deleted_previews) = rewrite_movie(source, selection, after, &mut budget)?;
    candidate.validate().map_err(map_read_error)?;
    let target = physical_catalog(&candidate)?.shared_source();
    let target_selection = select_movie(
        &candidate,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
        true,
    )?;
    if !same_selection(&target_selection, selection) || target_selection.before != Some(after) {
        return Err(SlideMoviePlaybackError::Verification);
    }
    verify_locality(source, &candidate, selection, &mut budget)?;
    Ok(SlideMoviePlaybackCommit {
        package: candidate,
        patch: SlideMoviePlaybackPatch {
            artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
            selection: selection.clone(),
            before,
            after,
            touched_components: 1,
            deleted_previews,
        },
        diagnostics: SlideMoviePlaybackDiagnostics::published(1, deleted_previews),
    })
}

fn reopen_target_patch(
    source: &Package,
    patch: &SlideMoviePlaybackPatch,
    budget: &mut PlaybackBudget,
) -> Result<SlideMoviePlaybackCommit, SlideMoviePlaybackError> {
    budget.candidate_reopen(patch.artifacts.target().len())?;
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    let selected = select_movie_with_budget(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        MovieSelector::position(patch.selection.movie_position),
        true,
        budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != Some(patch.after) {
        return Err(SlideMoviePlaybackError::Verification);
    }
    verify_locality(source, &candidate, &patch.selection, budget)?;
    Ok(SlideMoviePlaybackCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideMoviePlaybackDiagnostics::published(
            patch.touched_components,
            patch.deleted_previews,
        ),
    })
}

fn select_movie(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    mutation_guards: bool,
) -> Result<PlaybackSelection, SlideMoviePlaybackError> {
    let mut budget = PlaybackBudget::for_package(package)?;
    select_movie_with_budget(
        package,
        slide_selector,
        movie_selector,
        mutation_guards,
        &mut budget,
    )
}

fn select_movie_with_budget(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    _mutation_guards: bool,
    budget: &mut PlaybackBudget,
) -> Result<PlaybackSelection, SlideMoviePlaybackError> {
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideMoviePlaybackError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (slide_component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideMoviePlaybackError::InvalidSource)?;
    if slide.messages.len() != slide.archive_info.message_infos.len() {
        return Err(SlideMoviePlaybackError::InvalidSource);
    }
    let slide_message = unique_message(slide, SLIDE_MESSAGE_TYPE)?;
    let base_limits = package.wire_limits().map_err(map_wire_error)?;
    let limits = budget.residual_wire_limits(base_limits)?;
    let movies = repeated_references(slide_message.1, SLIDE_OWNED_DRAWABLES_FIELD, limits)?;
    budget.wire_scan(slide_message.1, 1)?;
    budget.references(movies.len())?;
    let mut movie_entries = Vec::new();
    movie_entries.try_reserve_exact(movies.len()).map_err(|_| {
        SlideMoviePlaybackError::Allocation {
            amount: movies.len(),
        }
    })?;
    for identifier in movies {
        let (component, movie) = package
            .object_with_component(identifier)
            .ok_or(SlideMoviePlaybackError::InvalidSource)?;
        // A slide may own shapes, placeholders, or other drawable records
        // beside movies. Only a type-3007 payload participates in the movie
        // source order; unrelated drawable records are ignored here.
        if movie
            .messages
            .iter()
            .all(|message| message.type_ != MOVIE_MESSAGE_TYPE)
        {
            continue;
        }
        if component != slide_component_name
            || movie.messages.len() != movie.archive_info.message_infos.len()
        {
            return Err(SlideMoviePlaybackError::InvalidSource);
        }
        let (message_index, payload) = unique_message(movie, MOVIE_MESSAGE_TYPE)?;
        let path = super::SemanticPath::SlideDrawable {
            slide: slide_position.get(),
            index: movie_entries.len(),
        };
        let semantic_limits =
            budget.residual_wire_limits(package.semantic_wire_limits().map_err(map_read_error)?)?;
        let preflight =
            super::preflight_movie(payload, semantic_limits, path).map_err(map_read_error)?;
        let (info, _) = super::decode_movie_info(
            payload,
            budget.residual_wire_limits(package.semantic_wire_limits().map_err(map_read_error)?)?,
            path,
        )
        .map_err(map_read_error)?;
        budget.wire_scan(payload, 1)?;
        budget.references(preflight.data_references)?;
        let parent = movie_parent(payload, budget.residual_wire_limits(base_limits)?)?;
        budget.wire_scan(payload, 2)?;
        if parent != record.slide_identifier {
            return Err(SlideMoviePlaybackError::InvalidSource);
        }
        movie_entries.push((identifier, message_index, info.kind(), info.playback()));
    }
    let movie_position = movie_selector.as_position();
    let (movie_identifier, message_index, movie_kind, _) = *movie_entries
        .get(movie_position.get())
        .ok_or(SlideMoviePlaybackError::MoviePositionNotFound {
            position: movie_position,
        })?;
    if !matches!(movie_kind, MovieKind::File | MovieKind::Audio) {
        return Err(SlideMoviePlaybackError::InvalidSource);
    }
    let movie = package
        .object(movie_identifier)
        .ok_or(SlideMoviePlaybackError::InvalidSource)?;
    let payload = movie
        .messages
        .get(message_index)
        .ok_or(SlideMoviePlaybackError::InvalidSource)?
        .data
        .as_slice();
    if movie_data_field_count(payload, base_limits)? != 1 {
        return Err(SlideMoviePlaybackError::InvalidSource);
    }
    ensure_unique_movie_identity(package, slide_component_name, movie_identifier)?;
    ensure_unique_movie_owner(
        package,
        record.slide_identifier,
        movie_identifier,
        base_limits,
        budget,
    )?;
    let options = codec_options(package, payload, budget)?;
    let (snapshot, report) =
        movie_playback_codec::decode_movie_playback_with_report(payload, options)
            .map_err(map_codec_error)?;
    budget.codec_report(report)?;
    let before = settings_from_snapshot(snapshot)?;
    Ok(PlaybackSelection {
        slide_position,
        movie_position,
        slide_identifier: record.slide_identifier,
        node_identifier: record.node_identifier,
        movie_identifier,
        message_index,
        slide_component_name: Arc::from(slide_component_name),
        before: Some(before),
    })
}

fn ensure_unique_movie_identity(
    package: &Package,
    selected_component: &str,
    movie_identifier: u64,
) -> Result<(), SlideMoviePlaybackError> {
    let mut matches = 0usize;
    let mut selected_matches = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object.archive_info.identifier == Some(movie_identifier) {
                matches = matches
                    .checked_add(1)
                    .ok_or(SlideMoviePlaybackError::InvalidSource)?;
                if component.name() == selected_component {
                    selected_matches = selected_matches
                        .checked_add(1)
                        .ok_or(SlideMoviePlaybackError::InvalidSource)?;
                }
            }
        }
    }
    if matches != 1 || selected_matches != 1 {
        // Cross-component movie graphs and physical duplicate identifiers are
        // intentionally unsupported by this scalar owner.
        return Err(SlideMoviePlaybackError::UnsupportedDependency);
    }
    Ok(())
}

fn rewrite_movie(
    source: &Package,
    selection: &PlaybackSelection,
    after: MediaPlaybackSettings,
    budget: &mut PlaybackBudget,
) -> Result<(Package, usize), SlideMoviePlaybackError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.slide_component_name.as_ref())
        .ok_or(SlideMoviePlaybackError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideMoviePlaybackError::InvalidSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.source(entry.data().len())?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    budget.physical(stream.as_bytes().len())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    let object = archive
        .object(selection.movie_identifier)
        .ok_or(SlideMoviePlaybackError::InvalidSource)?;
    let original = object
        .messages
        .get(selection.message_index)
        .ok_or(SlideMoviePlaybackError::InvalidSource)?
        .data
        .as_slice();
    let before = selection
        .before
        .ok_or(SlideMoviePlaybackError::UnsupportedDependency)?;
    if before == after {
        return Ok((source.snapshot(), 0));
    }

    // Size the archive and its worst-case Snappy member before the codec is
    // allowed to execute. The prepared codec still performs its own exact
    // requirements replay; this outer bound covers archive serialization,
    // compression, and the package member replacement without claiming one
    // physical allocation.
    let encoded_bound = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_bound).map_err(map_core_error)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(SlideMoviePlaybackError::LimitExceeded {
            kind: SlideMoviePlaybackLimitKind::EntryBytes,
            observed: compressed_bound as u64,
            maximum: snappy_limits.max_compressed_stream() as u64,
        });
    }
    let package_bound = source
        .source_bytes()
        .len()
        .checked_sub(entry.data().len())
        .and_then(|value| value.checked_add(compressed_bound))
        .ok_or(SlideMoviePlaybackError::InvalidSource)?;
    budget.preflight_output(
        encoded_bound
            .checked_add(compressed_bound)
            .ok_or(SlideMoviePlaybackError::InvalidSource)?,
    )?;
    budget.preflight_output(package_bound)?;
    budget.preflight_work(
        encoded_bound
            .checked_add(compressed_bound)
            .ok_or(SlideMoviePlaybackError::InvalidSource)?,
    )?;
    let options = codec_options(source, original, budget)?;
    let write = movie_playback_write(after)?;
    let prepared = movie_playback_codec::prepare_movie_playback_rewrite(original, write, options)
        .map_err(map_codec_error)?;
    budget.codec_report(prepared.prepare_report())?;
    let requirements = prepared.execution_requirements();
    budget.codec_requirements(requirements)?;
    let rewritten = prepared
        .execute(requirements.exact_limits())
        .map_err(map_codec_error)?;
    archive
        .object_mut(selection.movie_identifier)
        .ok_or(SlideMoviePlaybackError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.message_index,
            RawMessage {
                type_: MOVIE_MESSAGE_TYPE,
                data: rewritten.into_bytes(),
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.physical(bytes.len())?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    let edit = EntryEdit::new(
        selection.slide_component_name.as_ref(),
        compressed.as_slice(),
    );
    // Playback scalars do not change slide geometry or text rendering.  Keep
    // the existing preview members byte-for-byte; only the selected IWA
    // component is part of this owner’s locality contract.
    let edits = [edit];
    let prepared_reassembly = catalog
        .prepare_reassembly_with_deletions(&edits, &[], physical_limits)
        .map_err(map_archive_error)?;
    let reassembly_requirements = prepared_reassembly.execution_requirements();
    budget.reassembly(reassembly_requirements)?;
    budget.candidate_reopen(reassembly_requirements.output_bytes())?;
    let output = prepared_reassembly
        .execute(reassembly_requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((candidate, 0))
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideMoviePlaybackError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideMoviePlaybackError::EmptySlideName);
            }
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(map_slide_selector_error)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideMoviePlaybackError::SlideNameNotFound)
        },
    }
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &PlaybackSelection,
    budget: &mut PlaybackBudget,
) -> Result<(), SlideMoviePlaybackError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    if source_catalog.package().len() != candidate_catalog.package().len() {
        return Err(SlideMoviePlaybackError::Verification);
    }
    for source_entry in source_catalog.package().iter() {
        let candidate_entry = candidate_catalog
            .package()
            .iter()
            .find(|entry| entry.name() == source_entry.name())
            .ok_or(SlideMoviePlaybackError::Verification)?;
        if source_entry.name() != selection.slide_component_name.as_ref()
            && source_entry.data() != candidate_entry.data()
        {
            return Err(SlideMoviePlaybackError::Verification);
        }
    }
    for candidate_entry in candidate_catalog.package().iter() {
        if source_catalog
            .package()
            .iter()
            .all(|entry| entry.name() != candidate_entry.name())
        {
            return Err(SlideMoviePlaybackError::Verification);
        }
    }
    budget.component()?;
    let source_archive = component_archive(source, selection.slide_component_name.as_ref())?;
    let candidate_archive = component_archive(candidate, selection.slide_component_name.as_ref())?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(SlideMoviePlaybackError::Verification);
    }
    for source_object in &source_archive.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(SlideMoviePlaybackError::Verification)?;
        let candidate_object = candidate_archive
            .object(identifier)
            .ok_or(SlideMoviePlaybackError::Verification)?;
        if identifier == selection.movie_identifier {
            let source_message = source_object
                .messages
                .get(selection.message_index)
                .ok_or(SlideMoviePlaybackError::Verification)?;
            let candidate_message = candidate_object
                .messages
                .get(selection.message_index)
                .ok_or(SlideMoviePlaybackError::Verification)?;
            if source_message.type_ != candidate_message.type_ {
                return Err(SlideMoviePlaybackError::Verification);
            }
            let mut expected = source_object.clone();
            expected
                .replace_message_preserving_header_with_limits(
                    selection.message_index,
                    candidate_message.clone(),
                    source
                        .state
                        .options
                        .archive()
                        .effective_archive_limits()
                        .map_err(map_archive_error)?,
                )
                .map_err(map_core_error)?;
            // The physical provenance lengths describe the selected
            // candidate framing, not the immutable source framing. All raw
            // header/message metadata remains compared by
            // `same_content_ignoring_offsets` below.
            expected.header_length = candidate_object.header_length;
            expected.data_length = candidate_object.data_length;
            if !expected.same_content_ignoring_offsets(candidate_object) {
                return Err(SlideMoviePlaybackError::Verification);
            }
        } else if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(SlideMoviePlaybackError::Verification);
        }
        budget.work(source_object.messages.len())?;
    }
    Ok(())
}

fn component_archive(package: &Package, name: &str) -> Result<Archive, SlideMoviePlaybackError> {
    let catalog = physical_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(SlideMoviePlaybackError::InvalidSource)?;
    let snappy = package
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream =
        SnappyStream::decompress_with_limits(entry.data(), snappy).map_err(map_core_error)?;
    Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)
}

fn unique_message(
    object: &litchi_iwa_core::ArchiveObject,
    message_type: u32,
) -> Result<(usize, &[u8]), SlideMoviePlaybackError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(SlideMoviePlaybackError::InvalidSource)?;
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(SlideMoviePlaybackError::InvalidSource);
        }
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(SlideMoviePlaybackError::InvalidSource);
        }
    }
    selected.ok_or(SlideMoviePlaybackError::InvalidSource)
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: litchi_iwa_common::WireLimits,
) -> Result<Vec<u64>, SlideMoviePlaybackError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(fields.len())
        .map_err(|_| SlideMoviePlaybackError::Allocation {
            amount: fields.len(),
        })?;
    for field in fields
        .fields()
        .filter(|field| field.number() == field_number)
    {
        if field.wire_type() != 2 {
            return Err(SlideMoviePlaybackError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        values.push(strict_reference_payload(
            field.payload(),
            limits,
            "Keynote movie playback reference",
        )?);
    }
    Ok(values)
}

fn movie_parent(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<u64, SlideMoviePlaybackError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let super_field = fields
        .fields()
        .filter(|field| field.number() == MOVIE_SUPER_FIELD)
        .collect::<Vec<_>>();
    if super_field.len() != 1 || super_field[0].wire_type() != 2 {
        return Err(SlideMoviePlaybackError::InvalidSource);
    }
    let nested =
        WireView::parse_with_limits(super_field[0].payload(), limits).map_err(map_wire_error)?;
    let parent = nested
        .fields()
        .filter(|field| field.number() == DRAWABLE_PARENT_FIELD)
        .collect::<Vec<_>>();
    if parent.len() != 1 || parent[0].wire_type() != 2 {
        return Err(SlideMoviePlaybackError::InvalidSource);
    }
    strict_reference_payload(parent[0].payload(), limits, "Keynote movie parent")
}

fn strict_reference_payload(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    _context: &'static str,
) -> Result<u64, SlideMoviePlaybackError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    for field in fields.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideMoviePlaybackError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideMoviePlaybackError::InvalidSource)?;
                if width != encoded_len(value) || value == 0 {
                    return Err(SlideMoviePlaybackError::InvalidSource);
                }
                identifier = Some(value);
            },
            // The legacy Reference type and deprecated external marker are
            // not authoritative for this strict scalar edge. Reject their
            // presence, including explicit defaults, instead of silently
            // accepting an external/typed route.
            2 | 3 => return Err(SlideMoviePlaybackError::InvalidSource),
            _ => {},
        }
    }
    identifier.ok_or(SlideMoviePlaybackError::InvalidSource)
}

fn movie_data_field_count(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<usize, SlideMoviePlaybackError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut count = 0usize;
    for field in fields.fields().filter(|field| field.number() == 14) {
        if field.wire_type() != 2 {
            return Err(SlideMoviePlaybackError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        super::validate_movie_data_reference(field.payload(), limits).map_err(map_wire_error)?;
        count = count
            .checked_add(1)
            .ok_or(SlideMoviePlaybackError::InvalidSource)?;
    }
    Ok(count)
}

fn ensure_unique_movie_owner(
    package: &Package,
    slide_identifier: u64,
    movie_identifier: u64,
    limits: litchi_iwa_common::WireLimits,
    budget: &mut PlaybackBudget,
) -> Result<(), SlideMoviePlaybackError> {
    let mut occurrences = 0usize;
    let mut selected_occurrences = 0usize;
    for component in package.state.source.components().iter() {
        budget.component()?;
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != SLIDE_MESSAGE_TYPE {
                    continue;
                }
                let owner_identifier = object.archive_info.identifier;
                let ids = repeated_references(&message.data, SLIDE_OWNED_DRAWABLES_FIELD, limits)?;
                budget.wire_scan(&message.data, 1)?;
                budget.references(ids.len())?;
                for id in ids {
                    if id == movie_identifier {
                        occurrences = occurrences
                            .checked_add(1)
                            .ok_or(SlideMoviePlaybackError::InvalidSource)?;
                        if owner_identifier == Some(slide_identifier) {
                            selected_occurrences = selected_occurrences
                                .checked_add(1)
                                .ok_or(SlideMoviePlaybackError::InvalidSource)?;
                        }
                    }
                }
            }
        }
    }
    if occurrences != 1 || selected_occurrences != 1 {
        return Err(SlideMoviePlaybackError::InvalidSource);
    }
    Ok(())
}

fn same_selection(left: &PlaybackSelection, right: &PlaybackSelection) -> bool {
    left.slide_position == right.slide_position
        && left.movie_position == right.movie_position
        && left.slide_identifier == right.slide_identifier
        && left.node_identifier == right.node_identifier
        && left.movie_identifier == right.movie_identifier
        && left.message_index == right.message_index
        && left.slide_component_name == right.slide_component_name
}

fn codec_options(
    package: &Package,
    payload: &[u8],
    budget: &PlaybackBudget,
) -> Result<movie_playback_codec::DecodeOptions, SlideMoviePlaybackError> {
    let limits =
        budget.residual_wire_limits(package.semantic_wire_limits().map_err(map_read_error)?)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideMoviePlaybackError::InvalidSource)?;
    let output = budget
        .max_output
        .saturating_sub(budget.output)
        .max(payload.len().max(1));
    Ok(movie_playback_codec::DecodeOptions::new(
        limits.max_input_bytes().min(payload.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    )
    .with_max_output_bytes(output))
}

fn movie_playback_write(
    settings: MediaPlaybackSettings,
) -> Result<movie_playback_codec::MoviePlaybackWrite, SlideMoviePlaybackError> {
    let canonical = settings
        .canonicalize()
        .map_err(|_| SlideMoviePlaybackError::InvalidSource)?;
    Ok(movie_playback_codec::MoviePlaybackWrite::from_values(
        canonical.start_time.map(duration_seconds).transpose()?,
        duration_seconds(canonical.end_time)?,
        canonical.poster_time.map(duration_seconds).transpose()?,
        canonical.loop_mode.map(MediaLoopMode::as_raw),
        canonical.volume.map(MediaVolume::as_f32),
    ))
}

fn duration_seconds(value: Duration) -> Result<f32, SlideMoviePlaybackError> {
    let seconds = value.as_secs_f64();
    if !seconds.is_finite() || seconds > f64::from(f32::MAX) {
        return Err(SlideMoviePlaybackError::InvalidSource);
    }
    Ok(seconds as f32)
}

fn settings_from_snapshot(
    snapshot: movie_playback_codec::MoviePlaybackSnapshot,
) -> Result<MediaPlaybackSettings, SlideMoviePlaybackError> {
    let duration = |value: f32| {
        if !value.is_finite() || value < 0.0 {
            return Err(SlideMoviePlaybackError::InvalidSource);
        }
        Duration::try_from_secs_f32(value).map_err(|_| SlideMoviePlaybackError::InvalidSource)
    };
    MediaPlaybackSettings::new(duration(snapshot.end_time)?)
        .with_start_time(snapshot.start_time.map(duration).transpose()?)
        .with_poster_time(snapshot.poster_time.map(duration).transpose()?)
        .with_loop_mode(snapshot.loop_mode.map(MediaLoopMode::from_raw))
        .with_volume(
            snapshot
                .volume
                .map(MediaVolume::new)
                .transpose()
                .map_err(|_| SlideMoviePlaybackError::InvalidSource)?,
        )
        .canonicalize()
        .map_err(|_| SlideMoviePlaybackError::InvalidSource)
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideMoviePlaybackError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideMoviePlaybackError::UnsupportedSource),
    }
}

fn map_slide_selector_error(error: crate::SlideSelectorError) -> SlideMoviePlaybackError {
    match error {
        crate::SlideSelectorError::DuplicateSlideName { .. } => {
            SlideMoviePlaybackError::AmbiguousSelector
        },
        crate::SlideSelectorError::EmptySlideName => SlideMoviePlaybackError::EmptySlideName,
    }
}

fn map_read_error(error: ReadError) -> SlideMoviePlaybackError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideMoviePlaybackError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Slides => SlideMoviePlaybackLimitKind::Slides,
                SemanticLimitKind::References => SlideMoviePlaybackLimitKind::References,
                SemanticLimitKind::Objects => SlideMoviePlaybackLimitKind::Entries,
                _ => SlideMoviePlaybackLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideMoviePlaybackError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideMoviePlaybackLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideMoviePlaybackLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideMoviePlaybackLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideMoviePlaybackLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideMoviePlaybackError::Allocation { amount },
        _ => SlideMoviePlaybackError::InvalidSource,
    }
}

fn map_codec_error(error: movie_playback_codec::DecodeError) -> SlideMoviePlaybackError {
    if let Some(limit) = error.limit_kind() {
        let (observed, maximum) = error.limit_values().unwrap_or((0, 0));
        let kind = match limit {
            movie_playback_codec::DecodeLimit::InputBytes => SlideMoviePlaybackLimitKind::WireBytes,
            movie_playback_codec::DecodeLimit::OutputBytes => {
                SlideMoviePlaybackLimitKind::OutputBytes
            },
            movie_playback_codec::DecodeLimit::Fields => SlideMoviePlaybackLimitKind::WireFields,
            movie_playback_codec::DecodeLimit::Work => SlideMoviePlaybackLimitKind::WireWork,
            movie_playback_codec::DecodeLimit::Nesting => SlideMoviePlaybackLimitKind::WireNesting,
            movie_playback_codec::DecodeLimit::Allocations
            | movie_playback_codec::DecodeLimit::Scratch => SlideMoviePlaybackLimitKind::WireWork,
            movie_playback_codec::DecodeLimit::Retained => SlideMoviePlaybackLimitKind::OutputBytes,
            _ => SlideMoviePlaybackLimitKind::WireWork,
        };
        return SlideMoviePlaybackError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    SlideMoviePlaybackError::InvalidSource
}

fn map_wire_error(_error: litchi_iwa_common::Error) -> SlideMoviePlaybackError {
    SlideMoviePlaybackError::InvalidSource
}
fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideMoviePlaybackError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideMoviePlaybackError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    SlideMoviePlaybackLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideMoviePlaybackLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SlideMoviePlaybackLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideMoviePlaybackLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    SlideMoviePlaybackLimitKind::TotalBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    SlideMoviePlaybackLimitKind::PlaybackBytes
                },
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    SlideMoviePlaybackLimitKind::WireBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideMoviePlaybackError::Allocation { amount }
        },
        _ => SlideMoviePlaybackError::InvalidSource,
    }
}
fn map_core_error(error: litchi_iwa_core::Error) -> SlideMoviePlaybackError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideMoviePlaybackError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes => {
                    SlideMoviePlaybackLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideMoviePlaybackLimitKind::Entries
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems => {
                    SlideMoviePlaybackLimitKind::WireFields
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideMoviePlaybackLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes => {
                    SlideMoviePlaybackLimitKind::PlaybackBytes
                },
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    SlideMoviePlaybackLimitKind::EntryBytes
                },
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideMoviePlaybackError::Allocation { amount: requested }
        },
        _ => SlideMoviePlaybackError::InvalidSource,
    }
}
