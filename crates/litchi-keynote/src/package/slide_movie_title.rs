//! Exact-source, selector-first Keynote movie-title transactions.
//!
//! Movie titles use the same native CaptionInfo/storage graph as movie
//! captions, with the `MovieArchive.title` edge and
//! `CaptionOrTitleKind::Title` profile. The graph, Metadata, preview, and
//! resource-accounting machinery is shared with the movie-caption owner.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The semantic boundary deliberately redacts lower-layer failures."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use thiserror::Error as ThisError;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{MovieSelector, SlideSelector};

const MAX_TITLE_BYTES: usize = 64 * 1024 * 1024;

/// A finite resource governed while a movie-title transaction is prepared or
/// published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    OutputBytes,
    WireBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    Slides,
    References,
    TextStorages,
    TextFragments,
    TextBytes,
    WireFields,
    WireNesting,
    WireWork,
    TitleBytes,
}

impl fmt::Display for LimitKind {
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
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::TitleBytes => "movie title bytes",
        })
    }
}

/// A content-redacted failure raised by a movie-title transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    #[error("the Keynote movie-title source cannot be edited safely")]
    InvalidSource,
    #[error("Keynote movie-title {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote movie-title transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote movie title failed semantic verification")]
    Verification,
    #[error("the Keynote movie-title patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    #[error("this Keynote source does not support physical movie-title edits")]
    UnsupportedSource,
    #[error("the requested Keynote movie-title graph operation is unsupported")]
    UnsupportedDependency,
    #[error("the Keynote movie-title selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no movie at position {position:?}")]
    MoviePositionNotFound { position: Position },
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. } | Self::UnsupportedDependency { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::AmbiguousSelector { .. } | Self::EmptySlideName { .. } => {
                RefusalKind::InvalidValue
            },
            Self::SlideNameNotFound { .. }
            | Self::SlidePositionNotFound { .. }
            | Self::MoviePositionNotFound { .. } => RefusalKind::NotFound,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One mutable movie-title value staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    selection: super::slide_movie_caption::MovieCaptionSelection,
    after: Option<String>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn new<'slide>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
    ) -> Result<Self, Error> {
        let selection = select_title(source, slide_selector.into(), movie_selector.into(), true)?;
        let after = selection.text.as_deref().map(copy_title).transpose()?;
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
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    pub fn set(mut self, title: impl AsRef<str>) -> Result<Self, Error> {
        self.after = Some(copy_title(title.as_ref())?);
        Ok(self)
    }

    pub fn clear(mut self) -> Result<Self, Error> {
        self.after = None;
        Ok(self)
    }

    pub fn commit(self) -> Result<Commit, Error> {
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let current = select_title(
            self.source,
            SlideSelector::position(self.selection.slide_position),
            MovieSelector::position(self.selection.movie_position),
            true,
        )?;
        if !current.same_identity(&self.selection) || current.text != self.selection.text {
            return Err(Error::InvalidSource);
        }
        if self.selection.text == self.after {
            self.source.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    selection: self.selection,
                    target_selection: current,
                    after: self.after,
                    touched_components: 0,
                    deleted_previews: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        self.source.validate().map_err(map_read_error)?;
        if self
            .selection
            .text
            .as_deref()
            .is_some_and(super::slide_movie_caption::contains_dependent_marker)
            || self
                .after
                .as_deref()
                .is_some_and(super::slide_movie_caption::contains_dependent_marker)
        {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        let mut budget = super::slide_chart_caption::CaptionBudget::for_package(self.source)
            .map_err(map_chart_error)?;
        budget
            .charge_catalog_scan(self.source)
            .map_err(map_chart_error)?;
        budget
            .charge_selection_scan(self.source, 1)
            .map_err(map_chart_error)?;
        let (package, touched_components, deleted_previews) = rewrite_title(
            self.source,
            &self.selection,
            self.after.as_deref(),
            &mut budget,
        )?;
        budget
            .charge_selection_scan(&package, 1)
            .map_err(map_chart_error)?;
        let candidate = select_title(
            &package,
            SlideSelector::position(self.selection.slide_position),
            MovieSelector::position(self.selection.movie_position),
            true,
        )?;
        if !candidate.same_movie_identity(&self.selection) || candidate.text != self.after {
            return Err(Error::Verification);
        }
        budget
            .charge_validation_scan(&package, 1)
            .map_err(map_chart_error)?;
        verify_title_transition(
            self.source,
            &package,
            &self.selection,
            &candidate,
            true,
            &mut budget,
        )?;
        let target = physical_catalog(&package)?.shared_source();
        budget
            .charge_exact_artifacts(source_bytes.len(), target.len())
            .map_err(map_chart_error)?;
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(source_bytes, target),
                selection: self.selection,
                target_selection: candidate,
                after: self.after,
                touched_components,
                deleted_previews,
                target_requires_invalidated_previews: true,
            },
            diagnostics: Diagnostics::published(touched_components, deleted_previews),
        })
    }
}

/// An exact-source-checked reversible movie-title patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    selection: super::slide_movie_caption::MovieCaptionSelection,
    target_selection: super::slide_movie_caption::MovieCaptionSelection,
    after: Option<String>,
    touched_components: usize,
    deleted_previews: usize,
    target_requires_invalidated_previews: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl Patch {
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }

    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
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
        self.selection == self.target_selection && self.artifacts.is_byte_noop()
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.target_selection.clone(),
            target_selection: self.selection.clone(),
            after: self.selection.text.clone(),
            touched_components: self.touched_components,
            deleted_previews: 0,
            target_requires_invalidated_previews: self.touched_components != 0
                && !self.target_requires_invalidated_previews,
        }
    }
}

/// Compact evidence describing one movie-title commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
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

/// The fully verified result of one immutable movie-title transaction.
#[must_use = "a Keynote movie-title commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the title of one selected file-backed movie.
    pub fn slide_movie_title<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
    ) -> Result<Option<String>, Error> {
        Ok(select_title(self, slide_selector.into(), movie_selector.into(), true)?.text)
    }

    /// Start an exact immutable edit of one selected movie title.
    pub fn edit_slide_movie_title<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
    ) -> Result<Edit<'_>, Error> {
        Edit::new(self, slide_selector, movie_selector)
    }

    /// Apply an exact-source-checked movie-title patch.
    pub fn apply_slide_movie_title(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let current = select_title(
            self,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            true,
        )?;
        if !current.same_identity(&patch.selection) || current.text != patch.selection.text {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        let mut budget = super::slide_chart_caption::CaptionBudget::for_package(self)
            .map_err(map_chart_error)?;
        budget.charge_catalog_scan(self).map_err(map_chart_error)?;
        budget
            .charge_selection_scan(self, 1)
            .map_err(map_chart_error)?;
        budget
            .charge_exact_artifacts(source.len(), patch.artifacts.target().len())
            .map_err(map_chart_error)?;
        budget
            .charge_candidate_reopen(patch.artifacts.target().len())
            .map_err(map_chart_error)?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        budget
            .charge_selection_scan(&candidate, 1)
            .map_err(map_chart_error)?;
        let selected = select_title(
            &candidate,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            true,
        )?;
        if !selected.same_identity(&patch.target_selection) || selected.text != patch.after {
            return Err(Error::Verification);
        }
        budget
            .charge_validation_scan(&candidate, 1)
            .map_err(map_chart_error)?;
        verify_title_transition(
            self,
            &candidate,
            &patch.selection,
            &selected,
            patch.target_requires_invalidated_previews,
            &mut budget,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(patch.touched_components, patch.deleted_previews),
        })
    }
}

fn select_title(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    mutation_guards: bool,
) -> Result<super::slide_movie_caption::MovieCaptionSelection, Error> {
    super::slide_movie_caption::select_movie_text(
        package,
        slide_selector,
        movie_selector,
        mutation_guards,
        super::slide_movie_caption::MovieTextKind::Title,
    )
    .map_err(map_movie_caption_error)
}

fn rewrite_title(
    source: &Package,
    selection: &super::slide_movie_caption::MovieCaptionSelection,
    after: Option<&str>,
    budget: &mut super::slide_chart_caption::CaptionBudget,
) -> Result<(Package, usize, usize), Error> {
    let result = if selection.storage_identifier.is_some() && after.is_some() {
        let end = selection
            .text
            .as_deref()
            .ok_or(Error::InvalidSource)?
            .encode_utf16()
            .count();
        super::slide_chart_caption::rewrite_existing_caption_text_with_metadata_budget(
            source,
            selection.storage_identifier.ok_or(Error::InvalidSource)?,
            selection.slide_node_identifier,
            &selection.slide_component_name,
            end,
            after.ok_or(Error::InvalidSource)?,
            budget,
        )
    } else {
        super::slide_chart_caption::rewrite_caption_graph_operation_with_budget(
            source,
            &selection.slide_component_name,
            selection.movie_identifier,
            selection.reference_identifier,
            selection.storage_identifier,
            after,
            super::slide_chart_caption::CaptionEdgeKind::MovieTitle,
            budget,
        )
    };
    result.map_err(map_chart_error)
}

fn verify_title_transition(
    source: &Package,
    candidate: &Package,
    before: &super::slide_movie_caption::MovieCaptionSelection,
    target: &super::slide_movie_caption::MovieCaptionSelection,
    require_invalidated_previews: bool,
    budget: &mut super::slide_chart_caption::CaptionBudget,
) -> Result<(), Error> {
    if before.storage_identifier.is_some() && target.storage_identifier.is_some() {
        super::slide_chart_caption::verify_existing_text_metadata_candidate(
            source,
            candidate,
            before.storage_identifier.ok_or(Error::InvalidSource)?,
            before.slide_node_identifier,
            require_invalidated_previews,
            budget,
        )
        .map_err(map_chart_error)
    } else {
        super::slide_chart_caption::verify_caption_graph_transition(
            source,
            candidate,
            &before.slide_component_name,
            before.reference_identifier,
            before.caption_info_identifier,
            before.storage_identifier,
            before.placement_identifier,
            before.style_identifier,
            target.reference_identifier,
            target.caption_info_identifier,
            target.storage_identifier,
            target.placement_identifier,
            target.style_identifier,
        )
        .map_err(map_chart_error)?;
        if require_invalidated_previews
            && !super::rendering_invalidation::root_previews_absent(
                candidate.state.source.package(),
            )
            .map_err(|_| Error::Verification)?
        {
            return Err(Error::Verification);
        }
        Ok(())
    }
}

fn copy_title(value: &str) -> Result<String, Error> {
    if value.len() > MAX_TITLE_BYTES {
        return Err(Error::Limit {
            kind: LimitKind::TitleBytes,
            observed: usize_to_u64(value.len()),
            maximum: usize_to_u64(MAX_TITLE_BYTES),
        });
    }
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|_| Error::Allocation {
            amount: value.len(),
        })?;
    copy.push_str(value);
    Ok(copy)
}

fn physical_catalog(package: &Package) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::Refused(Refusal::UnsupportedSource)),
    }
}

fn map_movie_caption_error(error: super::slide_movie_caption::Error) -> Error {
    match error {
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::UnsupportedSource,
        ) => Error::Refused(Refusal::UnsupportedSource),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::UnsupportedDependency,
        ) => Error::Refused(Refusal::UnsupportedDependency),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::AmbiguousSelector,
        ) => Error::Refused(Refusal::AmbiguousSelector),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::EmptySlideName,
        ) => Error::Refused(Refusal::EmptySlideName),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::SlideNameNotFound,
        ) => Error::Refused(Refusal::SlideNameNotFound),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::SlidePositionNotFound { position },
        ) => Error::Refused(Refusal::SlidePositionNotFound { position }),
        super::slide_movie_caption::Error::Refused(
            crate::slide::movie::caption::Refusal::MoviePositionNotFound { position },
        ) => Error::Refused(Refusal::MoviePositionNotFound { position }),
        super::slide_movie_caption::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_caption_limit(kind),
            observed,
            maximum,
        },
        super::slide_movie_caption::Error::Allocation { amount } => Error::Allocation { amount },
        super::slide_movie_caption::Error::Verification => Error::Verification,
        super::slide_movie_caption::Error::PatchConflict => Error::PatchConflict,
        _ => Error::InvalidSource,
    }
}

fn map_chart_error(error: super::slide_chart_caption::Error) -> Error {
    match error {
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::UnsupportedSource,
        ) => Error::Refused(Refusal::UnsupportedSource),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::UnsupportedDependency,
        ) => Error::Refused(Refusal::UnsupportedDependency),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::AmbiguousSelector,
        ) => Error::Refused(Refusal::AmbiguousSelector),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::EmptySlideName,
        ) => Error::Refused(Refusal::EmptySlideName),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::SlideNameNotFound,
        ) => Error::Refused(Refusal::SlideNameNotFound),
        super::slide_chart_caption::Error::Refused(
            crate::chart::caption::Refusal::SlidePositionNotFound { position },
        ) => Error::Refused(Refusal::SlidePositionNotFound { position }),
        super::slide_chart_caption::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_chart_limit(kind),
            observed,
            maximum,
        },
        super::slide_chart_caption::Error::Allocation { amount } => Error::Allocation { amount },
        super::slide_chart_caption::Error::Verification => Error::Verification,
        super::slide_chart_caption::Error::PatchConflict => Error::PatchConflict,
        _ => Error::InvalidSource,
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                SemanticLimitKind::Objects => LimitKind::Entries,
                SemanticLimitKind::Slides => LimitKind::Slides,
                SemanticLimitKind::References => LimitKind::References,
                SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_caption_limit(kind: super::slide_movie_caption::LimitKind) -> LimitKind {
    use super::slide_movie_caption::LimitKind as K;
    match kind {
        K::InputBytes => LimitKind::InputBytes,
        K::OutputBytes => LimitKind::OutputBytes,
        K::WireBytes => LimitKind::WireBytes,
        K::Entries => LimitKind::Entries,
        K::EntryBytes => LimitKind::EntryBytes,
        K::TotalBytes => LimitKind::TotalBytes,
        K::Slides => LimitKind::Slides,
        K::References => LimitKind::References,
        K::TextStorages => LimitKind::TextStorages,
        K::TextFragments => LimitKind::TextFragments,
        K::TextBytes => LimitKind::TextBytes,
        K::WireFields => LimitKind::WireFields,
        K::WireNesting => LimitKind::WireNesting,
        K::WireWork => LimitKind::WireWork,
        K::CaptionBytes => LimitKind::TitleBytes,
    }
}

fn map_chart_limit(kind: super::slide_chart_caption::LimitKind) -> LimitKind {
    use super::slide_chart_caption::LimitKind as K;
    match kind {
        K::InputBytes => LimitKind::InputBytes,
        K::OutputBytes => LimitKind::OutputBytes,
        K::WireBytes => LimitKind::WireBytes,
        K::Entries => LimitKind::Entries,
        K::EntryBytes => LimitKind::EntryBytes,
        K::TotalBytes => LimitKind::TotalBytes,
        K::Slides => LimitKind::Slides,
        K::References => LimitKind::References,
        K::TextStorages => LimitKind::TextStorages,
        K::TextFragments => LimitKind::TextFragments,
        K::TextBytes => LimitKind::TextBytes,
        K::WireFields => LimitKind::WireFields,
        K::WireNesting => LimitKind::WireNesting,
        K::WireWork => LimitKind::WireWork,
        K::CaptionBytes => LimitKind::TitleBytes,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
