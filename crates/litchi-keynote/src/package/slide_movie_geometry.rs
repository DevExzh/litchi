//! Exact-source, selector-first geometry transactions for file-backed movies.
//!
//! Geometry is a rendering-affecting edge: successful edits invalidate the
//! root previews.  This module deliberately owns only an existing movie's
//! position and displayed size.  Movie flags, angle, media, captions,
//! playback, builds, metadata, and object allocation remain opaque.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The semantic boundary redacts lower-layer failure details."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len, wire::WireView};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_movie_geometry_codec;
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::slide::media::geometry::MovieGeometry;
use crate::slide::media::{Point, Size};
use crate::{MovieKind, MovieSelector, SlideSelector};

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const MOVIE_SUPER_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;

/// Resource categories reported by a movie-geometry transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideMovieGeometryLimitKind {
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
    GeometryBytes,
    Allocations,
    Retained,
    Scratch,
    Components,
}

impl fmt::Display for SlideMovieGeometryLimitKind {
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
            Self::GeometryBytes => "movie geometry bytes",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
        })
    }
}

/// Content-redacted failure raised by a movie-geometry transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideMovieGeometryError {
    #[error("this Keynote source does not support physical movie-geometry edits")]
    UnsupportedSource,
    #[error("the requested Keynote movie-geometry graph is unsupported")]
    UnsupportedDependency,
    #[error("the Keynote movie-geometry selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no movie at position {position:?}")]
    MoviePositionNotFound { position: Position },
    #[error("the Keynote movie-geometry source is invalid")]
    InvalidSource,
    #[error("Keynote movie geometry {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        kind: SlideMovieGeometryLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote movie-geometry transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote movie geometry failed semantic verification")]
    Verification,
    #[error("the Keynote movie-geometry patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy)]
struct GeometryBudget {
    max_input: usize,
    max_output: usize,
    max_fields: usize,
    max_work: usize,
    max_nesting: usize,
    max_references: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    nesting: usize,
    references: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
}

impl GeometryBudget {
    fn new(package: &Package) -> Result<Self, SlideMovieGeometryError> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let source: usize = package
            .state
            .options
            .archive()
            .max_input_bytes()
            .try_into()
            .map_err(|_| SlideMovieGeometryError::InvalidSource)?;
        let aggregate = source
            .checked_mul(4)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        Ok(Self {
            max_input: aggregate,
            max_output: aggregate,
            max_fields: wire.max_fields(),
            max_work: wire.max_rewrite_work(),
            max_nesting: wire.max_nesting(),
            max_references: package.semantic_limits().max_references(),
            max_allocations: aggregate,
            max_retained: aggregate,
            max_scratch: aggregate,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: SlideMovieGeometryLimitKind,
    ) -> Result<(), SlideMovieGeometryError> {
        let value = current
            .checked_add(amount)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if value > maximum {
            return Err(SlideMovieGeometryError::LimitExceeded {
                kind,
                observed: value as u64,
                maximum: maximum as u64,
            });
        }
        *current = value;
        Ok(())
    }

    fn source(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.input,
            bytes,
            self.max_input,
            SlideMovieGeometryLimitKind::InputBytes,
        )
    }

    fn output(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.output,
            bytes,
            self.max_output,
            SlideMovieGeometryLimitKind::OutputBytes,
        )
    }

    fn references(&mut self, amount: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SlideMovieGeometryLimitKind::References,
        )
    }

    fn work(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.work,
            bytes,
            self.max_work,
            SlideMovieGeometryLimitKind::WireWork,
        )
    }

    fn allocations(&mut self, amount: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SlideMovieGeometryLimitKind::Allocations,
        )
    }

    fn retained(&mut self, amount: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SlideMovieGeometryLimitKind::Retained,
        )
    }

    fn scratch(&mut self, amount: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            SlideMovieGeometryLimitKind::Scratch,
        )
    }

    fn physical(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        self.source(bytes)?;
        self.work(bytes)?;
        self.output(
            bytes
                .checked_mul(2)
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SlideMovieGeometryError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.work(requirements.output_bytes())
    }

    fn candidate_reopen(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        self.source(bytes)?;
        self.work(bytes)
    }

    fn residual(
        &self,
        package: &Package,
    ) -> Result<litchi_iwa_common::WireLimits, SlideMovieGeometryError> {
        let base = package.wire_limits().map_err(map_wire_error)?;
        base.with_input_bytes(
            base.max_input_bytes()
                .min(self.max_input.saturating_sub(self.input).max(1)),
        )
        .and_then(|v| {
            v.with_fields(
                base.max_fields()
                    .min(self.max_fields.saturating_sub(self.fields).max(1)),
            )
        })
        .and_then(|v| {
            v.with_rewrite_work(
                base.max_rewrite_work()
                    .min(self.max_work.saturating_sub(self.work).max(1)),
            )
        })
        .and_then(|v| v.with_nesting(base.max_nesting().min(self.max_nesting)))
        .map_err(map_wire_error)
    }

    fn codec_report(
        &mut self,
        report: keynote_movie_geometry_codec::DecodeReport,
    ) -> Result<(), SlideMovieGeometryError> {
        self.source(report.input_bytes())?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())?;
        Self::add(
            &mut self.fields,
            report.fields(),
            self.max_fields,
            SlideMovieGeometryLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            report.work_bytes(),
            self.max_work,
            SlideMovieGeometryLimitKind::WireWork,
        )?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn codec_requirements(
        &mut self,
        requirements: keynote_movie_geometry_codec::RewriteExecutionRequirements,
    ) -> Result<(), SlideMovieGeometryError> {
        self.output(requirements.output_bytes)?;
        Self::add(
            &mut self.fields,
            requirements.fields,
            self.max_fields,
            SlideMovieGeometryLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            requirements.work_bytes,
            self.max_work,
            SlideMovieGeometryLimitKind::WireWork,
        )?;
        self.allocations(requirements.allocations)?;
        self.retained(requirements.retained_bytes)?;
        self.scratch(requirements.scratch_bytes)?;
        if usize::try_from(requirements.max_depth).unwrap_or(usize::MAX) > self.max_nesting {
            return Err(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::WireNesting,
                observed: requirements.max_depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }
}

/// One mutable geometry value staged against an immutable package snapshot.
pub struct SlideMovieGeometryEdit<'a> {
    source: &'a Package,
    selection: GeometrySelection,
    after: MovieGeometry,
}

impl fmt::Debug for SlideMovieGeometryEdit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideMovieGeometryEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .finish_non_exhaustive()
    }
}

impl<'a> SlideMovieGeometryEdit<'a> {
    fn new<'slide>(
        source: &'a Package,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<Self, SlideMovieGeometryError> {
        let selection = select_movie(source, slide.into(), movie.into())?;
        let before = selection
            .before
            .ok_or(SlideMovieGeometryError::UnsupportedDependency)?;
        Ok(Self {
            source,
            selection,
            after: before,
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
    pub const fn before(&self) -> Option<MovieGeometry> {
        self.selection.before
    }
    #[must_use]
    pub const fn after(&self) -> MovieGeometry {
        self.after
    }

    pub fn set(mut self, geometry: MovieGeometry) -> Result<Self, SlideMovieGeometryError> {
        self.after = geometry;
        Ok(self)
    }

    pub fn commit(self) -> Result<SlideMovieGeometryCommit, SlideMovieGeometryError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible movie-geometry patch.
#[derive(Clone, PartialEq)]
pub struct SlideMovieGeometryPatch {
    artifacts: ExactArtifacts,
    selection: GeometrySelection,
    before: MovieGeometry,
    after: MovieGeometry,
    touched_components: usize,
    deleted_previews: usize,
    source_previews_absent: bool,
    target_previews_absent: bool,
}

impl fmt::Debug for SlideMovieGeometryPatch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideMovieGeometryPatch")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .finish_non_exhaustive()
    }
}

impl SlideMovieGeometryPatch {
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }
    #[must_use]
    pub const fn before(&self) -> MovieGeometry {
        self.before
    }
    #[must_use]
    pub const fn after(&self) -> MovieGeometry {
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
            source_previews_absent: self.target_previews_absent,
            target_previews_absent: self.source_previews_absent,
        }
    }
}

/// Compact movie-geometry publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideMovieGeometryDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideMovieGeometryDiagnostics {
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

/// Fully verified result of one movie-geometry transaction.
#[must_use = "a Keynote movie-geometry commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideMovieGeometryCommit {
    package: Package,
    patch: SlideMovieGeometryPatch,
    diagnostics: SlideMovieGeometryDiagnostics,
}

impl SlideMovieGeometryCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    #[must_use]
    pub const fn patch(&self) -> &SlideMovieGeometryPatch {
        &self.patch
    }
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideMovieGeometryDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq)]
struct GeometrySelection {
    slide_position: Position,
    movie_position: Position,
    slide_identifier: u64,
    node_identifier: u64,
    movie_identifier: u64,
    message_index: usize,
    slide_component_name: Arc<str>,
    before: Option<MovieGeometry>,
}

impl fmt::Debug for GeometrySelection {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("GeometrySelection")
            .field("slide_position", &self.slide_position)
            .field("movie_position", &self.movie_position)
            .field("has_before", &self.before.is_some())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Read the validated position and displayed size of one file-backed movie.
    pub fn slide_movie_geometry<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<Option<MovieGeometry>, SlideMovieGeometryError> {
        Ok(select_movie(self, slide.into(), movie.into())?.before)
    }

    /// Begin an exact immutable edit of an existing file-backed movie geometry.
    pub fn edit_slide_movie_geometry<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<SlideMovieGeometryEdit<'_>, SlideMovieGeometryError> {
        SlideMovieGeometryEdit::new(self, slide, movie)
    }

    /// Apply an exact-source checked movie-geometry patch.
    pub fn apply_slide_movie_geometry(
        &self,
        patch: &SlideMovieGeometryPatch,
    ) -> Result<SlideMovieGeometryCommit, SlideMovieGeometryError> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideMovieGeometryError::PatchConflict);
        }
        if previews_absent(self)? != patch.source_previews_absent {
            return Err(SlideMovieGeometryError::PatchConflict);
        }
        let current = select_movie(
            self,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
        )?;
        if !same_selection(&current, &patch.selection) || current.before != Some(patch.before) {
            return Err(SlideMovieGeometryError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideMovieGeometryCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideMovieGeometryDiagnostics::unchanged(),
            });
        }
        reopen_target_patch(self, patch)
    }
}

fn commit_edit(
    source: &Package,
    selection: &GeometrySelection,
    after: MovieGeometry,
) -> Result<SlideMovieGeometryCommit, SlideMovieGeometryError> {
    let before = selection
        .before
        .ok_or(SlideMovieGeometryError::UnsupportedDependency)?;
    if before == after {
        let bytes: Arc<[u8]> = Arc::from(source.source_bytes());
        return Ok(SlideMovieGeometryCommit {
            package: source.snapshot(),
            patch: SlideMovieGeometryPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before,
                after,
                touched_components: 0,
                deleted_previews: 0,
                source_previews_absent: previews_absent(source)?,
                target_previews_absent: previews_absent(source)?,
            },
            diagnostics: SlideMovieGeometryDiagnostics::unchanged(),
        });
    }
    let mut budget = GeometryBudget::new(source)?;
    let (candidate, deleted_previews) = rewrite_movie(source, selection, after, &mut budget)?;
    candidate.validate().map_err(map_read_error)?;
    if !previews_absent(&candidate)? {
        return Err(SlideMovieGeometryError::Verification);
    }
    let selected = select_movie(
        &candidate,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
    )?;
    if !same_selection(&selected, selection) || selected.before != Some(after) {
        return Err(SlideMovieGeometryError::Verification);
    }
    verify_locality(source, &candidate, selection, true, &mut budget)?;
    let target = physical_catalog(&candidate)?.shared_source();
    Ok(SlideMovieGeometryCommit {
        package: candidate,
        patch: SlideMovieGeometryPatch {
            artifacts: ExactArtifacts::new(Arc::from(source.source_bytes()), target),
            selection: selection.clone(),
            before,
            after,
            touched_components: 1,
            deleted_previews,
            source_previews_absent: previews_absent(source)?,
            target_previews_absent: true,
        },
        diagnostics: SlideMovieGeometryDiagnostics::published(1, deleted_previews),
    })
}

fn reopen_target_patch(
    source: &Package,
    patch: &SlideMovieGeometryPatch,
) -> Result<SlideMovieGeometryCommit, SlideMovieGeometryError> {
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    if previews_absent(&candidate)? != patch.target_previews_absent {
        return Err(SlideMovieGeometryError::Verification);
    }
    let selected = select_movie(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        MovieSelector::position(patch.selection.movie_position),
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != Some(patch.after) {
        return Err(SlideMovieGeometryError::Verification);
    }
    let mut budget = GeometryBudget::new(source)?;
    budget.candidate_reopen(patch.artifacts.target().len())?;
    verify_locality(
        source,
        &candidate,
        &patch.selection,
        patch.target_previews_absent,
        &mut budget,
    )?;
    Ok(SlideMovieGeometryCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideMovieGeometryDiagnostics::published(
            patch.touched_components,
            patch.deleted_previews,
        ),
    })
}

fn select_movie(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
) -> Result<GeometrySelection, SlideMovieGeometryError> {
    let catalog = physical_catalog(package)?;
    if !catalog.package().iter().any(|entry| {
        let name = entry.name().to_ascii_lowercase();
        name.starts_with("data/")
            && !entry.is_opaque()
            && [".mov", ".mp4", ".m4v", ".mpeg", ".mpg"]
                .iter()
                .any(|extension| name.ends_with(extension))
    }) {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    let mut budget = GeometryBudget::new(package)?;
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideMovieGeometryError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let (slide_message_index, slide_payload) = unique_message(slide, SLIDE_MESSAGE_TYPE)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let ids = repeated_references(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits)?;
    validate_archive_info_references(slide, slide_message_index, &ids, true)?;
    budget.references(ids.len())?;
    let mut movies = Vec::new();
    for identifier in ids {
        let (movie_component, object) = package
            .object_with_component(identifier)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if object
            .messages
            .iter()
            .all(|m| m.type_ != MOVIE_MESSAGE_TYPE)
        {
            continue;
        }
        if movie_component != component_name {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
        let (message_index, payload) = unique_message(object, MOVIE_MESSAGE_TYPE)?;
        let preflight = super::preflight_movie(
            payload,
            limits,
            super::SemanticPath::SlideDrawable {
                slide: slide_position.get(),
                index: movies.len(),
            },
        )
        .map_err(map_read_error)?;
        let (info, _) = super::decode_movie_info(
            payload,
            limits,
            super::SemanticPath::SlideDrawable {
                slide: slide_position.get(),
                index: movies.len(),
            },
        )
        .map_err(map_read_error)?;
        if info.kind() != MovieKind::File {
            continue;
        }
        if preflight.movie_data_fields != 1 || preflight.data_references == 0 {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
        if movie_parent(payload, limits)? != record.slide_identifier {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        let options = codec_options(package, payload, &budget)?;
        let (snapshot, report) =
            keynote_movie_geometry_codec::decode_movie_geometry_with_report(payload, options)
                .map_err(map_geometry_codec_error)?;
        budget.codec_report(report)?;
        let geometry = MovieGeometry::new(
            Point {
                x: snapshot.x(),
                y: snapshot.y(),
            },
            Size {
                width: snapshot.width(),
                height: snapshot.height(),
            },
        )
        .map(Some)
        .map_err(|_| SlideMovieGeometryError::InvalidSource)?;
        movies.push((identifier, message_index, geometry));
    }
    let movie_position = movie_selector.as_position();
    let (movie_identifier, message_index, before) = *movies.get(movie_position.get()).ok_or(
        SlideMovieGeometryError::MoviePositionNotFound {
            position: movie_position,
        },
    )?;
    ensure_unique_movie_identity(package, component_name, movie_identifier)?;
    ensure_unique_movie_owner(package, record.slide_identifier, movie_identifier, limits)?;
    let (_, movie_object) = package
        .object_with_component(movie_identifier)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let (_, movie_payload) = unique_message(movie_object, MOVIE_MESSAGE_TYPE)?;
    let movie_refs = movie_archive_references(movie_payload, limits)?;
    validate_archive_info_references(movie_object, message_index, &movie_refs, false)?;
    for identifier in &movie_refs {
        package
            .object_with_component(*identifier)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        let owners = package
            .state
            .source
            .components()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
            .filter(|object| object.archive_info.identifier == Some(*identifier))
            .count();
        if owners != 1 {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
    }
    if before.is_none() {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    Ok(GeometrySelection {
        slide_position,
        movie_position,
        slide_identifier: record.slide_identifier,
        node_identifier: record.node_identifier,
        movie_identifier,
        message_index,
        slide_component_name: Arc::from(component_name),
        before,
    })
}

fn rewrite_movie(
    source: &Package,
    selection: &GeometrySelection,
    after: MovieGeometry,
    budget: &mut GeometryBudget,
) -> Result<(Package, usize), SlideMovieGeometryError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|e| e.name() == selection.slide_component_name.as_ref())
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.physical(entry.data().len())?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    budget.physical(stream.as_bytes().len())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    let object = archive
        .object(selection.movie_identifier)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let original = object
        .messages
        .get(selection.message_index)
        .ok_or(SlideMovieGeometryError::InvalidSource)?
        .data
        .as_slice();
    let encoded_bound = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?
        .checked_add(original.len())
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_bound).map_err(map_core_error)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(SlideMovieGeometryError::LimitExceeded {
            kind: SlideMovieGeometryLimitKind::EntryBytes,
            observed: compressed_bound as u64,
            maximum: snappy_limits.max_compressed_stream() as u64,
        });
    }
    let package_bound = source
        .source_bytes()
        .len()
        .checked_sub(entry.data().len())
        .and_then(|value| value.checked_add(compressed_bound))
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    budget.output(
        encoded_bound
            .checked_add(compressed_bound)
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    budget.output(package_bound)?;
    budget.work(
        encoded_bound
            .checked_add(compressed_bound)
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    let options = codec_options(source, original, budget)?;
    let write = keynote_movie_geometry_codec::MovieGeometryWrite::from_values(
        after.position().x,
        after.position().y,
        after.size().width,
        after.size().height,
    );
    let prepared =
        keynote_movie_geometry_codec::prepare_movie_geometry_rewrite(original, write, options)
            .map_err(map_geometry_codec_error)?;
    budget.codec_report(prepared.prepare_report())?;
    let requirements = prepared.execution_requirements();
    budget.codec_requirements(requirements)?;
    let rewritten = prepared
        .execute(requirements.exact_limits())
        .map_err(map_geometry_codec_error)?
        .into_output();
    archive
        .object_mut(selection.movie_identifier)
        .ok_or(SlideMovieGeometryError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.message_index,
            RawMessage {
                type_: MOVIE_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.physical(bytes.len())?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| SlideMovieGeometryError::InvalidSource)?;
    let edit = EntryEdit::new(
        selection.slide_component_name.as_ref(),
        compressed.as_slice(),
    );
    let edits = [edit];
    let prepared = catalog
        .prepare_reassembly_with_deletions(&edits, previews.names(), physical_limits)
        .map_err(map_archive_error)?;
    let req = prepared.execution_requirements();
    budget.reassembly(req)?;
    budget.candidate_reopen(req.output_bytes())?;
    let output = prepared
        .execute(req.exact_limits())
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((candidate, previews.len()))
}

fn codec_options(
    package: &Package,
    payload: &[u8],
    budget: &GeometryBudget,
) -> Result<keynote_movie_geometry_codec::DecodeOptions, SlideMovieGeometryError> {
    let limits = budget.residual(package)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideMovieGeometryError::InvalidSource)?;
    let output = budget
        .max_output
        .saturating_sub(budget.output)
        .max(payload.len().max(1));
    Ok(keynote_movie_geometry_codec::DecodeOptions::new(
        limits.max_input_bytes().min(payload.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    )
    .with_max_output_bytes(output)
    .with_max_allocations(budget.max_input.max(1))
    .with_max_retained_bytes(budget.max_output.max(1))
    .with_max_scratch_bytes(budget.max_output.max(1)))
}

fn map_geometry_codec_error(
    error: keynote_movie_geometry_codec::DecodeError,
) -> SlideMovieGeometryError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            keynote_movie_geometry_codec::DecodeLimit::Bytes { observed, maximum } => {
                (SlideMovieGeometryLimitKind::WireBytes, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Fields { observed, maximum } => {
                (SlideMovieGeometryLimitKind::WireFields, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Work { observed, maximum } => {
                (SlideMovieGeometryLimitKind::WireWork, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Nesting { observed, maximum } => (
                SlideMovieGeometryLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            keynote_movie_geometry_codec::DecodeLimit::Allocations { observed, maximum } => {
                (SlideMovieGeometryLimitKind::Allocations, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Retained { observed, maximum } => {
                (SlideMovieGeometryLimitKind::Retained, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Scratch { observed, maximum } => {
                (SlideMovieGeometryLimitKind::Scratch, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Output { observed, maximum } => {
                (SlideMovieGeometryLimitKind::OutputBytes, observed, maximum)
            },
            _ => return SlideMovieGeometryError::InvalidSource,
        };
        return SlideMovieGeometryError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    SlideMovieGeometryError::InvalidSource
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &GeometrySelection,
    target_previews_absent: bool,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let previews = super::rendering_invalidation::root_preview_deletions(source_catalog.package())
        .map_err(|_| SlideMovieGeometryError::Verification)?;
    let candidate_previews =
        super::rendering_invalidation::root_preview_deletions(candidate_catalog.package())
            .map_err(|_| SlideMovieGeometryError::Verification)?;
    for entry in source_catalog.package().iter() {
        budget.work(entry.data().len())?;
        let candidate_entry = candidate_catalog
            .package()
            .iter()
            .find(|candidate| candidate.name() == entry.name());
        if target_previews_absent && previews.names().contains(&entry.name()) {
            if candidate_entry.is_some() {
                return Err(SlideMovieGeometryError::Verification);
            }
            continue;
        }
        let other = candidate_entry.ok_or(SlideMovieGeometryError::Verification)?;
        budget.work(other.data().len())?;
        if entry.name() != selection.slide_component_name.as_ref() && entry.data() != other.data() {
            return Err(SlideMovieGeometryError::Verification);
        }
    }
    for entry in candidate_catalog.package().iter() {
        if source_catalog
            .package()
            .iter()
            .all(|source| source.name() != entry.name())
        {
            if target_previews_absent || !candidate_previews.names().contains(&entry.name()) {
                return Err(SlideMovieGeometryError::Verification);
            }
        }
    }

    let source_archive = component_archive(source, selection.slide_component_name.as_ref())?;
    let candidate_archive = component_archive(candidate, selection.slide_component_name.as_ref())?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(SlideMovieGeometryError::Verification);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    for source_object in &source_archive.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(SlideMovieGeometryError::Verification)?;
        let candidate_object = candidate_archive
            .object(identifier)
            .ok_or(SlideMovieGeometryError::Verification)?;
        if identifier == selection.movie_identifier {
            let source_message = source_object
                .messages
                .get(selection.message_index)
                .ok_or(SlideMovieGeometryError::Verification)?;
            let candidate_message = candidate_object
                .messages
                .get(selection.message_index)
                .ok_or(SlideMovieGeometryError::Verification)?;
            if source_message.type_ != candidate_message.type_ {
                return Err(SlideMovieGeometryError::Verification);
            }
            let mut expected = source_object.clone();
            expected
                .replace_message_preserving_header_with_limits(
                    selection.message_index,
                    candidate_message.clone(),
                    archive_limits,
                )
                .map_err(map_core_error)?;
            expected.header_length = candidate_object.header_length;
            expected.data_length = candidate_object.data_length;
            if !expected.same_content_ignoring_offsets(candidate_object) {
                return Err(SlideMovieGeometryError::Verification);
            }
        } else if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(SlideMovieGeometryError::Verification);
        }
        budget.work(source_object.messages.len())?;
    }
    Ok(())
}

fn component_archive(package: &Package, name: &str) -> Result<Archive, SlideMovieGeometryError> {
    let catalog = physical_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
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

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideMovieGeometryError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideMovieGeometryError::EmptySlideName);
            }
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(map_slide_selector_error)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideMovieGeometryError::SlideNameNotFound)
        },
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
}

fn unique_message(
    object: &litchi_iwa_core::ArchiveObject,
    message_type: u32,
) -> Result<(usize, &[u8]), SlideMovieGeometryError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }
    selected.ok_or(SlideMovieGeometryError::InvalidSource)
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: litchi_iwa_common::WireLimits,
) -> Result<Vec<u64>, SlideMovieGeometryError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut result = Vec::new();
    for field in fields
        .fields()
        .filter(|field| field.number() == field_number)
    {
        if field.wire_type() != 2 {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        result.push(strict_reference_payload(field.payload(), limits)?);
    }
    Ok(result)
}

fn movie_parent(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<u64, SlideMovieGeometryError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let super_fields = fields
        .fields()
        .filter(|f| f.number() == MOVIE_SUPER_FIELD)
        .collect::<Vec<_>>();
    if super_fields.len() != 1 || super_fields[0].wire_type() != 2 {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    let nested =
        WireView::parse_with_limits(super_fields[0].payload(), limits).map_err(map_wire_error)?;
    let parent = nested
        .fields()
        .filter(|f| f.number() == DRAWABLE_PARENT_FIELD)
        .collect::<Vec<_>>();
    if parent.len() != 1 || parent[0].wire_type() != 2 {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    strict_reference_payload(parent[0].payload(), limits)
}

fn strict_reference_payload(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<u64, SlideMovieGeometryError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    for field in fields.fields() {
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideMovieGeometryError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideMovieGeometryError::InvalidSource)?;
                if width != encoded_len(value) || value == 0 {
                    return Err(SlideMovieGeometryError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(SlideMovieGeometryError::InvalidSource),
            _ => {},
        }
    }
    identifier.ok_or(SlideMovieGeometryError::InvalidSource)
}

fn validate_archive_info_references(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    expected: &[u64],
    allow_unselected: bool,
) -> Result<(), SlideMovieGeometryError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    if info
        .object_references
        .iter()
        .enumerate()
        .any(|(index, identifier)| info.object_references[..index].contains(identifier))
        || expected.iter().any(|identifier| {
            info.object_references
                .iter()
                .filter(|candidate| *candidate == identifier)
                .count()
                != 1
        })
        || info.field_infos.iter().any(|field| {
            field
                .object_references
                .iter()
                .enumerate()
                .any(|(index, identifier)| {
                    field.object_references[..index].contains(identifier)
                        || !info.object_references.contains(identifier)
                        || (expected.contains(identifier)
                            && (!allow_unselected
                                || field.path.as_slice() != [SLIDE_OWNED_DRAWABLES_FIELD]))
                })
                || field
                    .data_references
                    .iter()
                    .any(|identifier| expected.contains(identifier))
        })
    {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    Ok(())
}

fn movie_archive_references(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<Vec<u64>, SlideMovieGeometryError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let super_fields = fields
        .fields()
        .filter(|field| field.number() == MOVIE_SUPER_FIELD)
        .collect::<Vec<_>>();
    if super_fields.len() != 1 || super_fields[0].wire_type() != 2 {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    let drawable =
        WireView::parse_with_limits(super_fields[0].payload(), limits).map_err(map_wire_error)?;
    let mut references = Vec::new();
    let mut title_seen = false;
    let mut caption_seen = false;
    for field in drawable
        .fields()
        .filter(|field| matches!(field.number(), 10 | 11))
    {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        let seen = if field.number() == 10 {
            &mut title_seen
        } else {
            &mut caption_seen
        };
        if std::mem::replace(seen, true) {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        references.push(strict_reference_payload(field.payload(), limits)?);
    }
    let mut style_seen = false;
    for field in fields.fields().filter(|field| field.number() == 19) {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        if std::mem::replace(&mut style_seen, true) {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        references.push(strict_reference_payload(field.payload(), limits)?);
    }
    if references.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    Ok(references)
}

fn ensure_unique_movie_identity(
    package: &Package,
    component: &str,
    identifier: u64,
) -> Result<(), SlideMovieGeometryError> {
    let mut total = 0usize;
    let mut selected = 0usize;
    for current in package.state.source.components().iter() {
        for object in &current.archive().objects {
            if object.archive_info.identifier == Some(identifier) {
                total += 1;
                if current.name() == component {
                    selected += 1;
                }
            }
        }
    }
    if total != 1 || selected != 1 {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    Ok(())
}

fn ensure_unique_movie_owner(
    package: &Package,
    slide_identifier: u64,
    movie_identifier: u64,
    limits: litchi_iwa_common::WireLimits,
) -> Result<(), SlideMovieGeometryError> {
    let mut total = 0usize;
    let mut selected = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != SLIDE_MESSAGE_TYPE {
                    continue;
                }
                for id in repeated_references(&message.data, SLIDE_OWNED_DRAWABLES_FIELD, limits)? {
                    if id == movie_identifier {
                        total += 1;
                        if object.archive_info.identifier == Some(slide_identifier) {
                            selected += 1;
                        }
                    }
                }
            }
        }
    }
    if total != 1 || selected != 1 {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    Ok(())
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideMovieGeometryError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideMovieGeometryError::UnsupportedSource),
    }
}

fn previews_absent(package: &Package) -> Result<bool, SlideMovieGeometryError> {
    let catalog = physical_catalog(package)?;
    super::rendering_invalidation::root_previews_absent(catalog.package())
        .map_err(|_| SlideMovieGeometryError::Verification)
}

fn map_slide_selector_error(error: crate::SlideSelectorError) -> SlideMovieGeometryError {
    match error {
        crate::SlideSelectorError::DuplicateSlideName { .. } => {
            SlideMovieGeometryError::AmbiguousSelector
        },
        crate::SlideSelectorError::EmptySlideName => SlideMovieGeometryError::EmptySlideName,
    }
}
fn map_read_error(error: ReadError) -> SlideMovieGeometryError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideMovieGeometryError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Slides => SlideMovieGeometryLimitKind::Slides,
                SemanticLimitKind::References => SlideMovieGeometryLimitKind::References,
                _ => SlideMovieGeometryLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideMovieGeometryError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideMovieGeometryLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideMovieGeometryLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideMovieGeometryLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideMovieGeometryLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideMovieGeometryError::Allocation { amount },
        _ => SlideMovieGeometryError::InvalidSource,
    }
}
fn map_wire_error(_error: litchi_iwa_common::Error) -> SlideMovieGeometryError {
    SlideMovieGeometryError::InvalidSource
}
fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideMovieGeometryError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideMovieGeometryError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    SlideMovieGeometryLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideMovieGeometryLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SlideMovieGeometryLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideMovieGeometryLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    SlideMovieGeometryLimitKind::TotalBytes
                },
                _ => SlideMovieGeometryLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideMovieGeometryError::Allocation { amount }
        },
        _ => SlideMovieGeometryError::InvalidSource,
    }
}
fn map_core_error(error: litchi_iwa_core::Error) -> SlideMovieGeometryError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideMovieGeometryError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideMovieGeometryLimitKind::Entries
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems => {
                    SlideMovieGeometryLimitKind::WireFields
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideMovieGeometryLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    SlideMovieGeometryLimitKind::EntryBytes
                },
                _ => SlideMovieGeometryLimitKind::WireBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideMovieGeometryError::Allocation { amount: requested }
        },
        _ => SlideMovieGeometryError::InvalidSource,
    }
}
