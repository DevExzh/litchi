//! Exact-source, selector-first Keynote movie-caption transactions.
//!
//! Movie captions use the same native CaptionInfo/text-storage graph as chart
//! captions, but are reached through a movie drawable. Existing storage text,
//! canonical stand-in creation, and fresh-stand-in removal all share the chart
//! owner's bounded physical lifecycle seam.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction redacts native graph failures at the semantic boundary."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_protos::{keynote_movie_caption_codec, pages_movie_caption_codec};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{MovieSelector, SlideSelector};

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const BUILD_MESSAGE_TYPE: u32 = 8;
const SLIDE_BUILDS_FIELD: u32 = 2;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const MOVIE_SUPER_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;

const MAX_CAPTION_BYTES: usize = 64 * 1024 * 1024;

/// The two semantic text edges exposed by a Keynote movie drawable.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum MovieTextKind {
    Caption,
    Title,
}

impl MovieTextKind {
    const fn expected_child_info_kind(self) -> i32 {
        match self {
            Self::Caption => 1,
            Self::Title => 2,
        }
    }

    const fn edge_field_path(self) -> [u32; 2] {
        match self {
            Self::Caption => [11, 1],
            Self::Title => [10, 1],
        }
    }

    fn accepts_edge_path(self, path: &[u32]) -> bool {
        let field = self.edge_field_path()[0];
        matches!(path, [number, 1] if *number == field)
            || matches!(path, [1, number, 1] if *number == field)
            || matches!(path, [1, 1, number, 1] if *number == field)
    }
}

/// A finite resource governed while a movie-caption transaction is prepared
/// or published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideMovieCaptionLimitKind {
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
    /// Semantic text-storage objects.
    TextStorages,
    /// Semantic text fragments.
    TextFragments,
    /// Aggregate semantic text bytes.
    TextBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf scan and rewrite work.
    WireWork,
    /// UTF-8 bytes in one caption value.
    CaptionBytes,
}

impl fmt::Display for SlideMovieCaptionLimitKind {
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
            Self::CaptionBytes => "caption bytes",
        })
    }
}

/// A content-redacted failure raised by a movie-caption transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideMovieCaptionError {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical movie-caption edits")]
    UnsupportedSource,
    /// The requested operation needs a graph shape outside this owner.
    #[error("the requested Keynote movie-caption graph operation is unsupported")]
    UnsupportedDependency,
    /// An exact-name slide selector was ambiguous.
    #[error("the Keynote movie-caption selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// A checked semantic movie position does not exist.
    #[error("the selected Keynote slide has no movie at position {position:?}")]
    MoviePositionNotFound { position: Position },
    /// The source movie, caption graph, or text storage is malformed or unsafe.
    #[error("the Keynote movie-caption source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote movie-caption {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: SlideMovieCaptionLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote movie-caption transaction")]
    Allocation { amount: usize },
    /// Full candidate reopening did not reproduce the requested caption.
    #[error("the edited Keynote movie caption failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote movie-caption patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable movie-caption value staged against an immutable package.
pub struct SlideMovieCaptionEdit<'a> {
    source: &'a Package,
    selection: MovieCaptionSelection,
    after: Option<String>,
}

impl fmt::Debug for SlideMovieCaptionEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideMovieCaptionEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> SlideMovieCaptionEdit<'a> {
    fn new<'slide, 'movie>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
    ) -> Result<Self, SlideMovieCaptionError> {
        let selection = select_caption(source, slide_selector.into(), movie_selector.into(), true)?;
        let after = selection.text.as_deref().map(copy_caption).transpose()?;
        Ok(Self {
            source,
            selection,
            after,
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected semantic movie position within its slide.
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }

    /// Borrow the caption text observed when this edit began.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    /// Borrow the caption text staged for publication.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Stage whole-text replacement or creation of a native caption.
    pub fn set(mut self, caption: impl AsRef<str>) -> Result<Self, SlideMovieCaptionError> {
        self.after = Some(copy_caption(caption.as_ref())?);
        Ok(self)
    }

    /// Stage caption removal.
    pub fn clear(mut self) -> Result<Self, SlideMovieCaptionError> {
        self.after = None;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<SlideMovieCaptionCommit, SlideMovieCaptionError> {
        let source_catalog = physical_catalog(self.source)?;
        let source_bytes = source_catalog.shared_source();
        let current = select_caption(
            self.source,
            SlideSelector::position(self.selection.slide_position),
            MovieSelector::position(self.selection.movie_position),
            true,
        )?;
        if !current.same_identity(&self.selection) || current.text != self.selection.text {
            return Err(SlideMovieCaptionError::InvalidSource);
        }
        if self.selection.text == self.after {
            self.source.validate().map_err(map_read_error)?;
            return Ok(SlideMovieCaptionCommit {
                package: self.source.snapshot(),
                patch: SlideMovieCaptionPatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    selection: self.selection.clone(),
                    target_selection: self.selection,
                    after: self.after,
                    touched_components: 0,
                    deleted_previews: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: SlideMovieCaptionDiagnostics::unchanged(),
            });
        }
        if !source_catalog.source_is_exact() {
            return Err(SlideMovieCaptionError::UnsupportedSource);
        }
        self.source.validate().map_err(map_read_error)?;
        if self
            .selection
            .text
            .as_deref()
            .is_some_and(contains_dependent_marker)
            || self.after.as_deref().is_some_and(contains_dependent_marker)
        {
            return Err(SlideMovieCaptionError::UnsupportedDependency);
        }
        let mut budget = super::slide_chart_caption::CaptionBudget::for_package(self.source)
            .map_err(map_chart_caption_error)?;
        budget
            .charge_catalog_scan(self.source)
            .map_err(map_chart_caption_error)?;
        budget
            .charge_selection_scan(self.source, 1)
            .map_err(map_chart_caption_error)?;
        let (package, touched_components, deleted_previews) =
            match (self.selection.storage_identifier, self.after.as_deref()) {
                (Some(storage_identifier), Some(desired)) => {
                    let end = self
                        .selection
                        .text
                        .as_deref()
                        .ok_or(SlideMovieCaptionError::InvalidSource)?
                        .encode_utf16()
                        .count();
                    super::slide_chart_caption::rewrite_existing_caption_text_with_metadata_budget(
                        self.source,
                        storage_identifier,
                        self.selection.slide_node_identifier,
                        &self.selection.slide_component_name,
                        end,
                        desired,
                        &mut budget,
                    )
                    .map_err(map_chart_caption_error)?
                },
                (None, Some(desired)) => {
                    super::slide_chart_caption::rewrite_caption_graph_operation_with_budget(
                        self.source,
                        &self.selection.slide_component_name,
                        self.selection.movie_identifier,
                        self.selection.reference_identifier,
                        self.selection.storage_identifier,
                        Some(desired),
                        super::slide_chart_caption::CaptionEdgeKind::Movie,
                        &mut budget,
                    )
                    .map_err(map_chart_caption_error)?
                },
                (Some(_), None) => {
                    super::slide_chart_caption::rewrite_caption_graph_operation_with_budget(
                        self.source,
                        &self.selection.slide_component_name,
                        self.selection.movie_identifier,
                        self.selection.reference_identifier,
                        self.selection.storage_identifier,
                        None,
                        super::slide_chart_caption::CaptionEdgeKind::Movie,
                        &mut budget,
                    )
                    .map_err(map_chart_caption_error)?
                },
                (None, None) => return Err(SlideMovieCaptionError::InvalidSource),
            };
        budget
            .charge_selection_scan(&package, 1)
            .map_err(map_chart_caption_error)?;
        let candidate = select_caption(
            &package,
            SlideSelector::position(self.selection.slide_position),
            MovieSelector::position(self.selection.movie_position),
            true,
        )?;
        if !candidate.same_movie_identity(&self.selection) || candidate.text != self.after {
            return Err(SlideMovieCaptionError::Verification);
        }
        budget
            .charge_validation_scan(&package, 1)
            .map_err(map_chart_caption_error)?;
        if self.selection.storage_identifier.is_some() && candidate.storage_identifier.is_some() {
            super::slide_chart_caption::verify_existing_text_metadata_candidate(
                self.source,
                &package,
                self.selection
                    .storage_identifier
                    .ok_or(SlideMovieCaptionError::InvalidSource)?,
                self.selection.slide_node_identifier,
                true,
                &mut budget,
            )
            .map_err(map_chart_caption_error)?;
        } else {
            let expected_objects = if self.selection.storage_identifier.is_none() {
                self.source.state.total_objects.saturating_add(4)
            } else {
                self.source.state.total_objects.saturating_add(1)
            };
            if package.state.total_objects != expected_objects {
                return Err(SlideMovieCaptionError::Verification);
            }
            super::slide_chart_caption::verify_caption_graph_transition(
                self.source,
                &package,
                &self.selection.slide_component_name,
                self.selection.reference_identifier,
                self.selection.caption_info_identifier,
                self.selection.storage_identifier,
                self.selection.placement_identifier,
                self.selection.style_identifier,
                candidate.reference_identifier,
                candidate.caption_info_identifier,
                candidate.storage_identifier,
                candidate.placement_identifier,
                candidate.style_identifier,
            )
            .map_err(map_chart_caption_error)?;
            if !super::rendering_invalidation::root_previews_absent(package.state.source.package())
                .map_err(map_rendering_error)?
            {
                return Err(SlideMovieCaptionError::Verification);
            }
        }
        let target = physical_catalog(&package)?.shared_source();
        budget
            .charge_exact_artifacts(source_bytes.len(), target.len())
            .map_err(map_chart_caption_error)?;
        Ok(SlideMovieCaptionCommit {
            package,
            patch: SlideMovieCaptionPatch {
                artifacts: ExactArtifacts::new(source_bytes, target),
                selection: self.selection,
                target_selection: candidate,
                after: self.after,
                touched_components,
                deleted_previews,
                target_requires_invalidated_previews: true,
            },
            diagnostics: SlideMovieCaptionDiagnostics::published(
                touched_components,
                deleted_previews,
            ),
        })
    }
}

/// An exact-source-checked reversible semantic movie-caption patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideMovieCaptionPatch {
    artifacts: ExactArtifacts,
    selection: MovieCaptionSelection,
    target_selection: MovieCaptionSelection,
    after: Option<String>,
    touched_components: usize,
    deleted_previews: usize,
    target_requires_invalidated_previews: bool,
}

impl fmt::Debug for SlideMovieCaptionPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideMovieCaptionPatch")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl SlideMovieCaptionPatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected semantic movie position within its slide.
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }

    /// Borrow the caption required from the source package.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    /// Borrow the caption produced by the target package.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
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

    /// Return whether this patch preserves semantic caption state and bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.selection == self.target_selection && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from target back to source.
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

/// Compact evidence describing one movie-caption commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SlideMovieCaptionDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideMovieCaptionDiagnostics {
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

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable movie-caption transaction.
#[must_use = "a Keynote movie-caption commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideMovieCaptionCommit {
    package: Package,
    patch: SlideMovieCaptionPatch,
    diagnostics: SlideMovieCaptionDiagnostics,
}

impl SlideMovieCaptionCommit {
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
    pub const fn patch(&self) -> &SlideMovieCaptionPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideMovieCaptionDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the caption text of one selected file-backed movie.
    pub fn slide_movie_caption<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
    ) -> Result<Option<String>, SlideMovieCaptionError> {
        Ok(select_caption(self, slide_selector.into(), movie_selector.into(), true)?.text)
    }

    /// Start an exact immutable edit of one selected movie caption.
    pub fn edit_slide_movie_caption<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
    ) -> Result<SlideMovieCaptionEdit<'_>, SlideMovieCaptionError> {
        SlideMovieCaptionEdit::new(self, slide_selector, movie_selector)
    }

    /// Apply an exact-source-checked movie-caption patch.
    pub fn apply_slide_movie_caption(
        &self,
        patch: &SlideMovieCaptionPatch,
    ) -> Result<SlideMovieCaptionCommit, SlideMovieCaptionError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(SlideMovieCaptionError::PatchConflict);
        }
        let current = select_caption(
            self,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            true,
        )?;
        if !current.same_identity(&patch.selection) || current.text != patch.selection.text {
            return Err(SlideMovieCaptionError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideMovieCaptionCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideMovieCaptionDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(SlideMovieCaptionError::PatchConflict);
        }
        let mut budget = super::slide_chart_caption::CaptionBudget::for_package(self)
            .map_err(map_chart_caption_error)?;
        budget
            .charge_catalog_scan(self)
            .map_err(map_chart_caption_error)?;
        budget
            .charge_selection_scan(self, 1)
            .map_err(map_chart_caption_error)?;
        budget
            .charge_exact_artifacts(source.len(), patch.artifacts.target().len())
            .map_err(map_chart_caption_error)?;
        budget
            .charge_candidate_reopen(patch.artifacts.target().len())
            .map_err(map_chart_caption_error)?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        budget
            .charge_selection_scan(&candidate, 1)
            .map_err(map_chart_caption_error)?;
        let selected = select_caption(
            &candidate,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            true,
        )?;
        if !selected.same_identity(&patch.target_selection) || selected.text != patch.after {
            return Err(SlideMovieCaptionError::Verification);
        }
        budget
            .charge_validation_scan(&candidate, 1)
            .map_err(map_chart_caption_error)?;
        if patch.selection.storage_identifier.is_some()
            && patch.target_selection.storage_identifier.is_some()
        {
            super::slide_chart_caption::verify_existing_text_metadata_candidate(
                self,
                &candidate,
                patch
                    .selection
                    .storage_identifier
                    .ok_or(SlideMovieCaptionError::InvalidSource)?,
                patch.selection.slide_node_identifier,
                patch.target_requires_invalidated_previews,
                &mut budget,
            )
            .map_err(map_chart_caption_error)?;
        } else {
            let source_count = self.state.total_objects;
            let candidate_count = candidate.state.total_objects;
            let count_matches = match (
                patch.selection.storage_identifier,
                patch.target_selection.storage_identifier,
            ) {
                (None, Some(_)) => {
                    candidate_count == source_count.saturating_add(4)
                        || source_count == candidate_count.saturating_add(1)
                },
                (Some(_), None) => {
                    candidate_count == source_count.saturating_add(1)
                        || source_count == candidate_count.saturating_add(4)
                },
                _ => false,
            };
            if !count_matches {
                return Err(SlideMovieCaptionError::Verification);
            }
            super::slide_chart_caption::verify_caption_graph_transition(
                self,
                &candidate,
                &patch.selection.slide_component_name,
                patch.selection.reference_identifier,
                patch.selection.caption_info_identifier,
                patch.selection.storage_identifier,
                patch.selection.placement_identifier,
                patch.selection.style_identifier,
                patch.target_selection.reference_identifier,
                patch.target_selection.caption_info_identifier,
                patch.target_selection.storage_identifier,
                patch.target_selection.placement_identifier,
                patch.target_selection.style_identifier,
            )
            .map_err(map_chart_caption_error)?;
            if patch.target_requires_invalidated_previews
                && !super::rendering_invalidation::root_previews_absent(
                    candidate.state.source.package(),
                )
                .map_err(map_rendering_error)?
            {
                return Err(SlideMovieCaptionError::Verification);
            }
        }
        Ok(SlideMovieCaptionCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SlideMovieCaptionDiagnostics::published(
                patch.touched_components,
                patch.deleted_previews,
            ),
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct MovieCaptionSelection {
    pub(super) slide_position: Position,
    pub(super) movie_position: Position,
    pub(super) slide_identifier: u64,
    pub(super) slide_node_identifier: u64,
    pub(super) movie_identifier: u64,
    pub(super) slide_component_name: String,
    pub(super) reference_identifier: Option<u64>,
    pub(super) caption_info_identifier: Option<u64>,
    pub(super) storage_identifier: Option<u64>,
    pub(super) placement_identifier: Option<u64>,
    pub(super) style_identifier: Option<u64>,
    pub(super) text: Option<String>,
}

impl MovieCaptionSelection {
    pub(super) fn same_movie_identity(&self, other: &Self) -> bool {
        self.slide_position == other.slide_position
            && self.movie_position == other.movie_position
            && self.slide_identifier == other.slide_identifier
            && self.slide_node_identifier == other.slide_node_identifier
            && self.movie_identifier == other.movie_identifier
            && self.slide_component_name == other.slide_component_name
    }

    pub(super) fn same_identity(&self, other: &Self) -> bool {
        self.same_movie_identity(other)
            && self.reference_identifier == other.reference_identifier
            && self.caption_info_identifier == other.caption_info_identifier
            && self.storage_identifier == other.storage_identifier
            && self.placement_identifier == other.placement_identifier
            && self.style_identifier == other.style_identifier
    }
}

fn select_caption(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    mutation_guards: bool,
) -> Result<MovieCaptionSelection, SlideMovieCaptionError> {
    select_movie_text(
        package,
        slide_selector,
        movie_selector,
        mutation_guards,
        MovieTextKind::Caption,
    )
}

/// Resolve either the movie caption or title edge while retaining one strict
/// ownership census and one semantic selection representation.
pub(super) fn select_movie_text(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    mutation_guards: bool,
    kind: MovieTextKind,
) -> Result<MovieCaptionSelection, SlideMovieCaptionError> {
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideMovieCaptionError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    let (_node_component_name, _node) = package
        .object_with_component(record.node_identifier)
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    if slide.messages.len() != 1 {
        return Err(SlideMovieCaptionError::InvalidSource);
    }
    let payload = super::unique_payload(&slide.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
        .map_err(map_read_error)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let movie_identifiers = repeated_references(payload, SLIDE_OWNED_DRAWABLES_FIELD, limits)?;
    let mut movies = Vec::new();
    movies.try_reserve(movie_identifiers.len()).map_err(|_| {
        SlideMovieCaptionError::Allocation {
            amount: movie_identifiers.len(),
        }
    })?;
    for identifier in movie_identifiers {
        let (movie_component, movie) = package
            .object_with_component(identifier)
            .ok_or(SlideMovieCaptionError::InvalidSource)?;
        let movie_messages = movie
            .messages
            .iter()
            .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            .count();
        match movie_messages {
            0 => {},
            1 if movie_component == component_name && movie.messages.len() == 1 => {
                movies.push(identifier);
            },
            _ => return Err(SlideMovieCaptionError::InvalidSource),
        }
    }
    let movie_position = movie_selector.as_position();
    let movie_identifier =
        *movies
            .get(movie_position.get())
            .ok_or(SlideMovieCaptionError::MoviePositionNotFound {
                position: movie_position,
            })?;
    let movie = package
        .object(movie_identifier)
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    validate_selected_message_metadata(movie, 0)?;
    let movie_payload = &movie.messages[0].data;
    let (movie_info, _movie_references) = super::decode_movie_info(
        movie_payload,
        package.semantic_wire_limits().map_err(map_read_error)?,
        super::SemanticPath::SlideDrawable {
            slide: slide_position.get(),
            index: movie_position.get(),
        },
    )
    .map_err(map_read_error)?;
    if movie_info.kind() != crate::MovieKind::File {
        return Err(SlideMovieCaptionError::InvalidSource);
    }
    let movie_options = movie_decode_options(package, movie_payload)?;
    let movie_snapshot =
        keynote_movie_caption_codec::decode_movie_caption(movie_payload, movie_options)
            .map_err(map_movie_codec_error)?;
    if movie_parent_identifier(movie_payload, limits)? != record.slide_identifier {
        return Err(SlideMovieCaptionError::InvalidSource);
    }
    let reference_identifier = match kind {
        MovieTextKind::Caption => movie_snapshot.caption_identifier(),
        MovieTextKind::Title => movie_snapshot.title_identifier(),
    };
    let opposite_identifier = match kind {
        MovieTextKind::Caption => movie_snapshot.title_identifier(),
        MovieTextKind::Title => movie_snapshot.caption_identifier(),
    };
    if reference_identifier.is_some() && reference_identifier == opposite_identifier {
        return Err(SlideMovieCaptionError::InvalidSource);
    }
    if mutation_guards {
        prove_exclusive_movie_drawable(
            package,
            record.slide_identifier,
            component_name,
            movie_identifier,
        )?;
    }
    let empty = |reference_identifier| MovieCaptionSelection {
        slide_position,
        movie_position,
        slide_identifier: record.slide_identifier,
        slide_node_identifier: record.node_identifier,
        movie_identifier,
        slide_component_name: component_name.to_owned(),
        reference_identifier,
        caption_info_identifier: None,
        storage_identifier: None,
        placement_identifier: None,
        style_identifier: None,
        text: None,
    };
    let Some(reference_identifier) = reference_identifier else {
        return Ok(empty(None));
    };
    let (info_component, info_object) = package
        .object_with_component(reference_identifier)
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    if info_component != component_name || info_object.messages.len() != 1 {
        return Err(SlideMovieCaptionError::InvalidSource);
    }
    if info_object.messages[0].type_ == STANDIN_MESSAGE_TYPE {
        validate_selected_message_metadata(info_object, 0)?;
        if !info_object.messages[0].data.is_empty() {
            return Err(SlideMovieCaptionError::InvalidSource);
        }
        if mutation_guards {
            prove_exclusive_caption_standin(package, movie_identifier, reference_identifier, kind)?;
        }
        return Ok(empty(Some(reference_identifier)));
    }
    if info_object.messages[0].type_ != CAPTION_INFO_MESSAGE_TYPE {
        return Err(SlideMovieCaptionError::InvalidSource);
    }
    validate_selected_message_metadata(info_object, 0)?;
    let info_payload = &info_object.messages[0].data;
    let snapshot = pages_movie_caption_codec::decode_caption_info(
        info_payload,
        caption_info_decode_options(package, info_payload)?,
    )
    .map_err(map_caption_info_codec_error)?;
    let storage_identifier = snapshot
        .owned_storage_identifier()
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    if storage_identifier == 0
        || snapshot.deprecated_storage_identifier() != Some(storage_identifier)
        || snapshot.parent_identifier() != movie_identifier
        || snapshot.is_text_box() != Some(true)
        || snapshot.child_info_kind() != Some(kind.expected_child_info_kind())
    {
        return Err(SlideMovieCaptionError::InvalidSource);
    }
    let placement_identifier = snapshot
        .placement_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    let style_identifier = snapshot
        .style_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    let identifiers = [
        movie_identifier,
        reference_identifier,
        storage_identifier,
        placement_identifier,
        style_identifier,
    ];
    for (index, identifier) in identifiers.iter().enumerate() {
        if identifiers[..index].contains(identifier) {
            return Err(SlideMovieCaptionError::InvalidSource);
        }
    }
    require_private_object(
        package,
        storage_identifier,
        STORAGE_MESSAGE_TYPE,
        Some(component_name),
    )?;
    require_private_object(
        package,
        placement_identifier,
        CAPTION_PLACEMENT_MESSAGE_TYPE,
        Some(component_name),
    )?;
    require_private_object(
        package,
        style_identifier,
        SHAPE_STYLE_MESSAGE_TYPE,
        Some(component_name),
    )?;
    if mutation_guards {
        prove_exclusive_caption_storage(
            package,
            movie_identifier,
            reference_identifier,
            storage_identifier,
            placement_identifier,
            style_identifier,
            kind,
        )?;
    }
    let text = super::slide_text::read_owned_storage_text(package, storage_identifier)
        .map_err(map_slide_text_error)?;
    Ok(MovieCaptionSelection {
        slide_position,
        movie_position,
        slide_identifier: record.slide_identifier,
        slide_node_identifier: record.node_identifier,
        movie_identifier,
        slide_component_name: component_name.to_owned(),
        reference_identifier: Some(reference_identifier),
        caption_info_identifier: Some(reference_identifier),
        storage_identifier: Some(storage_identifier),
        placement_identifier: Some(placement_identifier),
        style_identifier: Some(style_identifier),
        text: Some(text),
    })
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideMovieCaptionError> {
    match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(map_read_error)?
            .map(|_record| position)
            .ok_or(SlideMovieCaptionError::SlidePositionNotFound { position }),
        SlideSelector::Name(name) => {
            let selector = SlideSelector::try_name(name).map_err(map_slide_selector_error)?;
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(map_slide_selector_error)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideMovieCaptionError::SlideNameNotFound)
        },
    }
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: litchi_iwa_common::WireLimits,
) -> Result<Vec<u64>, SlideMovieCaptionError> {
    let fields = litchi_iwa_common::wire::WireView::parse_with_limits(payload, limits)
        .map_err(map_wire_error)?;
    let mut references = Vec::new();
    references
        .try_reserve(fields.len())
        .map_err(|_| SlideMovieCaptionError::Allocation {
            amount: fields.len(),
        })?;
    for field in fields.fields() {
        if field.number() != field_number {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(SlideMovieCaptionError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        references.push(
            super::validate_reference_payload(field.payload(), limits, "Keynote slide movie")
                .map_err(map_wire_error)?,
        );
    }
    Ok(references)
}

fn unique_length_delimited_payload(
    payload: &[u8],
    field_number: u32,
    limits: litchi_iwa_common::WireLimits,
) -> Result<Option<&[u8]>, SlideMovieCaptionError> {
    let fields = litchi_iwa_common::wire::WireView::parse_with_limits(payload, limits)
        .map_err(map_wire_error)?;
    let mut selected = None;
    for field in fields
        .fields()
        .filter(|field| field.number() == field_number)
    {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(SlideMovieCaptionError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        selected = Some(field.payload());
    }
    Ok(selected)
}

fn movie_parent_identifier(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<u64, SlideMovieCaptionError> {
    let drawable = unique_length_delimited_payload(payload, MOVIE_SUPER_FIELD, limits)?
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    let parent = unique_length_delimited_payload(drawable, DRAWABLE_PARENT_FIELD, limits)?
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    super::validate_reference_payload(parent, limits, "Keynote movie parent")
        .map_err(map_wire_error)
}

fn prove_exclusive_movie_drawable(
    package: &Package,
    slide_identifier: u64,
    slide_component_name: &str,
    movie_identifier: u64,
) -> Result<(), SlideMovieCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let selected_build_identifiers = selected_movie_build_identifiers(
        package,
        slide_identifier,
        slide_component_name,
        movie_identifier,
        limits,
    )?;
    let mut payload_occurrences = 0usize;
    let mut aggregate_occurrences = 0usize;
    let mut field_occurrences = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(SlideMovieCaptionError::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
                let canonical_slide =
                    owner_identifier == slide_identifier && message.type_ == SLIDE_MESSAGE_TYPE;
                let selected_build = component.name() == slide_component_name
                    && message.type_ == BUILD_MESSAGE_TYPE
                    && selected_build_identifiers.contains(&owner_identifier);

                if message.type_ == SLIDE_MESSAGE_TYPE {
                    let local =
                        repeated_references(&message.data, SLIDE_OWNED_DRAWABLES_FIELD, limits)?
                            .into_iter()
                            .filter(|identifier| *identifier == movie_identifier)
                            .count();
                    if local != 0 && (!canonical_slide || local != 1) {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    payload_occurrences = payload_occurrences
                        .checked_add(local)
                        .ok_or(SlideMovieCaptionError::InvalidSource)?;
                }

                let aggregate_count = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == movie_identifier)
                    .count();
                let aggregate_data_count = info
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == movie_identifier)
                    .count();
                if aggregate_data_count != 0 {
                    return Err(SlideMovieCaptionError::UnsupportedDependency);
                }
                if aggregate_count != 0 {
                    if aggregate_count != 1 || (!canonical_slide && !selected_build) {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    aggregate_occurrences = aggregate_occurrences
                        .checked_add(aggregate_count)
                        .ok_or(SlideMovieCaptionError::InvalidSource)?;
                }

                for field in &info.field_infos {
                    let field_count = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == movie_identifier)
                        .count();
                    let field_data_count = field
                        .data_references
                        .iter()
                        .filter(|identifier| **identifier == movie_identifier)
                        .count();
                    if field_data_count != 0 {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    if field_count != 0 {
                        if !canonical_slide
                            || field_count != 1
                            || !matches!(
                                field.path.path.as_slice(),
                                [7, 1] | [1, 7, 1] | [1, 1, 7, 1]
                            )
                        {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                        field_occurrences = field_occurrences
                            .checked_add(field_count)
                            .ok_or(SlideMovieCaptionError::InvalidSource)?;
                    }
                }
            }
        }
    }
    let expected_aggregate_occurrences = 1usize
        .checked_add(selected_build_identifiers.len())
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    if payload_occurrences == 1
        && aggregate_occurrences == expected_aggregate_occurrences
        && field_occurrences <= 1
    {
        Ok(())
    } else {
        Err(SlideMovieCaptionError::InvalidSource)
    }
}

/// Resolve the only additional inbound edge authored by a fresh movie
/// creation: a same-component movie-start build registered by the selected
/// slide.  Header membership alone is insufficient because an unregistered
/// build (or a build whose payload names another drawable) must remain a
/// rejected cross-graph dependency.
fn selected_movie_build_identifiers(
    package: &Package,
    slide_identifier: u64,
    slide_component_name: &str,
    movie_identifier: u64,
    limits: litchi_iwa_common::WireLimits,
) -> Result<Vec<u64>, SlideMovieCaptionError> {
    let slide = package
        .object(slide_identifier)
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    let slide_payload =
        super::unique_payload(&slide.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
            .map_err(map_read_error)?;
    let build_identifiers = repeated_references(slide_payload, SLIDE_BUILDS_FIELD, limits)?;
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(build_identifiers.len())
        .map_err(|_| SlideMovieCaptionError::Allocation {
            amount: build_identifiers.len(),
        })?;
    for (index, identifier) in build_identifiers.iter().copied().enumerate() {
        if build_identifiers[..index].contains(&identifier) {
            return Err(SlideMovieCaptionError::UnsupportedDependency);
        }
        let (component, object) = package
            .object_with_component(identifier)
            .ok_or(SlideMovieCaptionError::InvalidSource)?;
        if component != slide_component_name {
            continue;
        }
        let mut build_message = None;
        for (message_index, message) in object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == BUILD_MESSAGE_TYPE)
        {
            if build_message.replace((message_index, message)).is_some() {
                return Err(SlideMovieCaptionError::UnsupportedDependency);
            }
        }
        let Some((message_index, message)) = build_message else {
            continue;
        };
        let info = object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(SlideMovieCaptionError::InvalidSource)?;
        if info.type_ != BUILD_MESSAGE_TYPE
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(SlideMovieCaptionError::InvalidSource);
        }
        let payload = message.data.as_slice();
        if build_target_identifier(payload, limits)? != movie_identifier {
            continue;
        }
        validate_movie_start_build(payload, limits)?;
        match info.object_references.as_slice() {
            [] => {},
            [reference] if *reference == movie_identifier => selected.push(identifier),
            _ => return Err(SlideMovieCaptionError::UnsupportedDependency),
        }
    }
    Ok(selected)
}

/// Require the canonical movie-start target/effect projection from one build
/// payload.  The generated native build model is deliberately not used at
/// this ingress boundary; this keeps the guard on the bounded wire view.
fn build_target_identifier(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<u64, SlideMovieCaptionError> {
    let drawable = unique_length_delimited_payload(payload, 1, limits)?
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    super::validate_reference_payload(drawable, limits, "Keynote movie build")
        .map_err(map_wire_error)
}

fn validate_movie_start_build(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<(), SlideMovieCaptionError> {
    let attributes = unique_length_delimited_payload(payload, 4, limits)?
        .ok_or(SlideMovieCaptionError::UnsupportedDependency)?;
    let animation = unique_length_delimited_payload(attributes, 18, limits)?
        .ok_or(SlideMovieCaptionError::UnsupportedDependency)?;
    let effect = unique_length_delimited_payload(animation, 2, limits)?
        .ok_or(SlideMovieCaptionError::UnsupportedDependency)?;
    if effect != b"apple:movie-start" {
        return Err(SlideMovieCaptionError::UnsupportedDependency);
    }
    Ok(())
}

fn require_private_object(
    package: &Package,
    identifier: u64,
    message_type: u32,
    expected_component: Option<&str>,
) -> Result<(), SlideMovieCaptionError> {
    let (component, object) = package
        .object_with_component(identifier)
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    if expected_component.is_some_and(|expected| expected != component)
        || object.messages.len() != 1
        || object.messages[0].type_ != message_type
    {
        return Err(SlideMovieCaptionError::InvalidSource);
    }
    validate_selected_message_metadata(object, 0)
}

fn validate_selected_message_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
) -> Result<(), SlideMovieCaptionError> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideMovieCaptionError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || message.base_message_index.is_some()
        || !message.diff_merge_version.is_empty()
        || message.diff_field_path.is_some()
        || !message.fields_to_remove.is_empty()
        || !message.diff_read_version.is_empty()
    {
        return Err(SlideMovieCaptionError::InvalidSource);
    }
    Ok(())
}

/// Prove that all physical aliases of the selected CaptionInfo/storage graph
/// belong to the selected movie. The aggregate and FieldInfo census is kept
/// deliberately global: an alias in an unrelated component is unsafe even
/// when the selected payload itself looks canonical.
fn prove_exclusive_caption_storage(
    package: &Package,
    movie_identifier: u64,
    caption_info_identifier: u64,
    storage_identifier: u64,
    placement_identifier: u64,
    style_identifier: u64,
    kind: MovieTextKind,
) -> Result<(), SlideMovieCaptionError> {
    let mut movie_edges = 0usize;
    let mut info_aggregate_edges = 0usize;
    let mut info_field_edges = 0usize;
    let mut storage_payload_owner = 0usize;
    let mut storage_aggregate_owner = 0usize;
    let mut storage_field_owner = 0usize;
    let mut style_payload_owner = 0usize;
    let mut style_aggregate_owner = 0usize;
    let mut style_field_owner = 0usize;
    let mut placement_payload_owner = 0usize;
    let mut placement_aggregate_owner = 0usize;
    let mut placement_field_owner = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(SlideMovieCaptionError::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
                if message.type_ == MOVIE_MESSAGE_TYPE {
                    let snapshot = keynote_movie_caption_codec::decode_movie_caption(
                        &message.data,
                        movie_decode_options(package, &message.data)?,
                    )
                    .map_err(map_movie_codec_error)?;
                    let opposite_identifier = match kind {
                        MovieTextKind::Caption => snapshot.title_identifier(),
                        MovieTextKind::Title => snapshot.caption_identifier(),
                    };
                    if opposite_identifier == Some(caption_info_identifier) {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    let edge_identifier = match kind {
                        MovieTextKind::Caption => snapshot.caption_identifier(),
                        MovieTextKind::Title => snapshot.title_identifier(),
                    };
                    if edge_identifier == Some(caption_info_identifier) {
                        movie_edges = movie_edges
                            .checked_add(1)
                            .ok_or(SlideMovieCaptionError::InvalidSource)?;
                        if owner_identifier != movie_identifier {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                    }
                }
                let aggregate_info = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == caption_info_identifier)
                    .count();
                let data_info = info
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == caption_info_identifier)
                    .count();
                if data_info != 0 {
                    return Err(SlideMovieCaptionError::UnsupportedDependency);
                }
                let mut local_info_fields = 0usize;
                for field in &info.field_infos {
                    let field_count = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == caption_info_identifier)
                        .count();
                    let field_data = field
                        .data_references
                        .iter()
                        .filter(|identifier| **identifier == caption_info_identifier)
                        .count();
                    if field_data != 0 {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    if field_count != 0 {
                        if owner_identifier != movie_identifier
                            || message.type_ != MOVIE_MESSAGE_TYPE
                            || field_count != 1
                            || !kind.accepts_edge_path(field.path.path.as_slice())
                        {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                        local_info_fields = local_info_fields
                            .checked_add(field_count)
                            .ok_or(SlideMovieCaptionError::InvalidSource)?;
                    }
                }
                if aggregate_info != 0 {
                    if owner_identifier != movie_identifier
                        || message.type_ != MOVIE_MESSAGE_TYPE
                        || aggregate_info != 1
                    {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                }
                info_aggregate_edges = info_aggregate_edges
                    .checked_add(aggregate_info)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
                info_field_edges = info_field_edges
                    .checked_add(local_info_fields)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;

                if message.type_ == CAPTION_INFO_MESSAGE_TYPE {
                    let snapshot = pages_movie_caption_codec::decode_caption_info(
                        &message.data,
                        caption_info_decode_options(package, &message.data)?,
                    )
                    .map_err(map_caption_info_codec_error)?;
                    let owns_storage = snapshot.owned_storage_identifier()
                        == Some(storage_identifier)
                        || snapshot.deprecated_storage_identifier() == Some(storage_identifier);
                    if owns_storage {
                        storage_payload_owner = storage_payload_owner
                            .checked_add(1)
                            .ok_or(SlideMovieCaptionError::InvalidSource)?;
                        if owner_identifier != caption_info_identifier {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                    }
                    if snapshot.style_identifier() == Some(style_identifier) {
                        style_payload_owner = style_payload_owner
                            .checked_add(1)
                            .ok_or(SlideMovieCaptionError::InvalidSource)?;
                        if owner_identifier != caption_info_identifier {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                    }
                    if snapshot.placement_identifier() == Some(placement_identifier) {
                        placement_payload_owner = placement_payload_owner
                            .checked_add(1)
                            .ok_or(SlideMovieCaptionError::InvalidSource)?;
                        if owner_identifier != caption_info_identifier {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                    }
                }
                let aggregate_style = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == style_identifier)
                    .count();
                let aggregate_placement = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == placement_identifier)
                    .count();
                if info.data_references.iter().any(|identifier| {
                    *identifier == style_identifier || *identifier == placement_identifier
                }) {
                    return Err(SlideMovieCaptionError::UnsupportedDependency);
                }
                if aggregate_style != 0 {
                    let allowed_owner = (owner_identifier == caption_info_identifier
                        && message.type_ == CAPTION_INFO_MESSAGE_TYPE)
                        || (owner_identifier == movie_identifier
                            && message.type_ == MOVIE_MESSAGE_TYPE);
                    if !allowed_owner || aggregate_style != 1 {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                }
                if aggregate_placement != 0 {
                    if owner_identifier != caption_info_identifier
                        || message.type_ != CAPTION_INFO_MESSAGE_TYPE
                        || aggregate_placement != 1
                    {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                }
                let mut local_style_fields = 0usize;
                let mut local_placement_fields = 0usize;
                for field in &info.field_infos {
                    let style_count = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == style_identifier)
                        .count();
                    let placement_count = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == placement_identifier)
                        .count();
                    if field.data_references.iter().any(|identifier| {
                        *identifier == style_identifier || *identifier == placement_identifier
                    }) {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    if style_count != 0 {
                        if owner_identifier != caption_info_identifier
                            || message.type_ != CAPTION_INFO_MESSAGE_TYPE
                            || style_count != 1
                            || !matches!(field.path.path.as_slice(), [1, 1, 2] | [1, 2])
                        {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                        local_style_fields = local_style_fields
                            .checked_add(style_count)
                            .ok_or(SlideMovieCaptionError::InvalidSource)?;
                    }
                    if placement_count != 0 {
                        if owner_identifier != caption_info_identifier
                            || message.type_ != CAPTION_INFO_MESSAGE_TYPE
                            || placement_count != 1
                            || !matches!(field.path.path.as_slice(), [2])
                        {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                        local_placement_fields = local_placement_fields
                            .checked_add(placement_count)
                            .ok_or(SlideMovieCaptionError::InvalidSource)?;
                    }
                }
                style_aggregate_owner = style_aggregate_owner
                    .checked_add(aggregate_style)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
                style_field_owner = style_field_owner
                    .checked_add(local_style_fields)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
                placement_aggregate_owner = placement_aggregate_owner
                    .checked_add(aggregate_placement)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
                placement_field_owner = placement_field_owner
                    .checked_add(local_placement_fields)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
                let aggregate_storage = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == storage_identifier)
                    .count();
                let data_storage = info
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == storage_identifier)
                    .count();
                if data_storage != 0 {
                    return Err(SlideMovieCaptionError::UnsupportedDependency);
                }
                let mut local_storage_fields = 0usize;
                for field in &info.field_infos {
                    let field_count = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == storage_identifier)
                        .count();
                    let field_data = field
                        .data_references
                        .iter()
                        .filter(|identifier| **identifier == storage_identifier)
                        .count();
                    if field_data != 0 {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    if field_count != 0 {
                        if owner_identifier != caption_info_identifier
                            || message.type_ != CAPTION_INFO_MESSAGE_TYPE
                            || field_count > 2
                            || !matches!(field.path.path.as_slice(), [1, 2] | [1, 4])
                        {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                        local_storage_fields = local_storage_fields
                            .checked_add(field_count)
                            .ok_or(SlideMovieCaptionError::InvalidSource)?;
                    }
                }
                if aggregate_storage != 0 {
                    if owner_identifier != caption_info_identifier
                        || message.type_ != CAPTION_INFO_MESSAGE_TYPE
                        || aggregate_storage != 1
                    {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                }
                storage_aggregate_owner = storage_aggregate_owner
                    .checked_add(aggregate_storage)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
                storage_field_owner = storage_field_owner
                    .checked_add(local_storage_fields)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
            }
        }
    }
    if movie_edges == 1
        && info_aggregate_edges == 1
        && info_field_edges <= 1
        && storage_payload_owner == 1
        && storage_aggregate_owner == 1
        && storage_field_owner <= 2
        && style_payload_owner == 1
        && style_aggregate_owner == 2
        && style_field_owner <= 1
        && placement_payload_owner == 1
        && placement_aggregate_owner == 1
        && placement_field_owner <= 1
    {
        Ok(())
    } else {
        Err(SlideMovieCaptionError::InvalidSource)
    }
}

fn prove_exclusive_caption_standin(
    package: &Package,
    movie_identifier: u64,
    standin_identifier: u64,
    kind: MovieTextKind,
) -> Result<(), SlideMovieCaptionError> {
    let mut payload_edges = 0usize;
    let mut aggregate_edges = 0usize;
    let mut field_edges = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(SlideMovieCaptionError::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == MOVIE_MESSAGE_TYPE {
                    let snapshot = keynote_movie_caption_codec::decode_movie_caption(
                        &message.data,
                        movie_decode_options(package, &message.data)?,
                    )
                    .map_err(map_movie_codec_error)?;
                    let opposite_identifier = match kind {
                        MovieTextKind::Caption => snapshot.title_identifier(),
                        MovieTextKind::Title => snapshot.caption_identifier(),
                    };
                    if opposite_identifier == Some(standin_identifier) {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    let edge_identifier = match kind {
                        MovieTextKind::Caption => snapshot.caption_identifier(),
                        MovieTextKind::Title => snapshot.title_identifier(),
                    };
                    if edge_identifier == Some(standin_identifier) {
                        payload_edges += 1;
                        if payload_edges > 1 || owner_identifier != movie_identifier {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                    }
                }
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(SlideMovieCaptionError::InvalidSource)?;
                if info.data_references.contains(&standin_identifier) {
                    return Err(SlideMovieCaptionError::UnsupportedDependency);
                }
                let aggregate = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == standin_identifier)
                    .count();
                if aggregate != 0 {
                    if owner_identifier != movie_identifier
                        || message.type_ != MOVIE_MESSAGE_TYPE
                        || aggregate != 1
                    {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    aggregate_edges += aggregate;
                }
                for field in &info.field_infos {
                    let field_count = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == standin_identifier)
                        .count();
                    if field.data_references.contains(&standin_identifier) {
                        return Err(SlideMovieCaptionError::UnsupportedDependency);
                    }
                    if field_count != 0 {
                        if owner_identifier != movie_identifier
                            || message.type_ != MOVIE_MESSAGE_TYPE
                            || field_count != 1
                            || !kind.accepts_edge_path(field.path.path.as_slice())
                        {
                            return Err(SlideMovieCaptionError::UnsupportedDependency);
                        }
                        field_edges += field_count;
                    }
                }
            }
        }
    }
    if payload_edges == 1 && aggregate_edges == 1 && field_edges <= 1 {
        Ok(())
    } else {
        Err(SlideMovieCaptionError::InvalidSource)
    }
}

fn movie_decode_options(
    package: &Package,
    payload: &[u8],
) -> Result<keynote_movie_caption_codec::DecodeOptions, SlideMovieCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideMovieCaptionError::InvalidSource)?;
    Ok(keynote_movie_caption_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    )
    .with_max_output_bytes(limits.max_output_bytes()))
}

fn caption_info_decode_options(
    package: &Package,
    payload: &[u8],
) -> Result<pages_movie_caption_codec::DecodeOptions, SlideMovieCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideMovieCaptionError::InvalidSource)?;
    Ok(pages_movie_caption_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideMovieCaptionError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideMovieCaptionError::UnsupportedSource),
    }
}

fn copy_caption(value: &str) -> Result<String, SlideMovieCaptionError> {
    if value.len() > MAX_CAPTION_BYTES {
        return Err(SlideMovieCaptionError::LimitExceeded {
            kind: SlideMovieCaptionLimitKind::CaptionBytes,
            observed: usize_to_u64(value.len()),
            maximum: usize_to_u64(MAX_CAPTION_BYTES),
        });
    }
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|_| SlideMovieCaptionError::Allocation {
            amount: value.len(),
        })?;
    copy.push_str(value);
    Ok(copy)
}

pub(super) fn contains_dependent_marker(text: &str) -> bool {
    text.contains('\u{000e}') || text.contains('\u{fffc}')
}

fn map_slide_selector_error(error: crate::SlideSelectorError) -> SlideMovieCaptionError {
    match error {
        crate::SlideSelectorError::EmptySlideName => SlideMovieCaptionError::EmptySlideName,
        crate::SlideSelectorError::DuplicateSlideName { .. } => {
            SlideMovieCaptionError::AmbiguousSelector
        },
    }
}

fn map_read_error(error: ReadError) -> SlideMovieCaptionError {
    match error {
        ReadError::Archive(error) => map_archive_error(error),
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideMovieCaptionError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => SlideMovieCaptionLimitKind::Entries,
                SemanticLimitKind::Slides => SlideMovieCaptionLimitKind::Slides,
                SemanticLimitKind::References => SlideMovieCaptionLimitKind::References,
                SemanticLimitKind::TextStorages => SlideMovieCaptionLimitKind::TextStorages,
                SemanticLimitKind::TextFragments => SlideMovieCaptionLimitKind::TextFragments,
                SemanticLimitKind::TextBytes => SlideMovieCaptionLimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideMovieCaptionError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideMovieCaptionLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideMovieCaptionLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideMovieCaptionLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideMovieCaptionLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => SlideMovieCaptionError::Allocation { amount },
        _ => SlideMovieCaptionError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideMovieCaptionError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideMovieCaptionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideMovieCaptionLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideMovieCaptionLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SlideMovieCaptionLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    SlideMovieCaptionLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    SlideMovieCaptionLimitKind::TotalBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    SlideMovieCaptionLimitKind::WireBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideMovieCaptionError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideMovieCaptionError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideMovieCaptionError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideMovieCaptionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    SlideMovieCaptionLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => SlideMovieCaptionLimitKind::Entries,
                litchi_iwa_core::LimitKind::HeaderFields => SlideMovieCaptionLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideMovieCaptionLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::SnappyFrames => SlideMovieCaptionLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideMovieCaptionError::Allocation { amount: requested }
        },
        _ => SlideMovieCaptionError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SlideMovieCaptionError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SlideMovieCaptionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => SlideMovieCaptionLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => {
                    SlideMovieCaptionLimitKind::OutputBytes
                },
                litchi_iwa_common::LimitKind::Fields => SlideMovieCaptionLimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => SlideMovieCaptionLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => SlideMovieCaptionLimitKind::WireWork,
                _ => SlideMovieCaptionLimitKind::WireBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        _ => SlideMovieCaptionError::InvalidSource,
    }
}

fn map_movie_codec_error(
    error: keynote_movie_caption_codec::DecodeError,
) -> SlideMovieCaptionError {
    if let Some((observed, maximum)) = error.output_limit_values() {
        return SlideMovieCaptionError::LimitExceeded {
            kind: SlideMovieCaptionLimitKind::OutputBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return SlideMovieCaptionError::Allocation { amount };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            keynote_movie_caption_codec::WireResourceLimit::Bytes { observed, maximum } => {
                SlideMovieCaptionError::LimitExceeded {
                    kind: SlideMovieCaptionLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            keynote_movie_caption_codec::WireResourceLimit::Nesting { observed, maximum } => {
                SlideMovieCaptionError::LimitExceeded {
                    kind: SlideMovieCaptionLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => SlideMovieCaptionError::InvalidSource,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideMovieCaptionError::LimitExceeded {
            kind: SlideMovieCaptionLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideMovieCaptionError::LimitExceeded {
            kind: SlideMovieCaptionLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    SlideMovieCaptionError::InvalidSource
}

fn map_caption_info_codec_error(
    error: pages_movie_caption_codec::DecodeError,
) -> SlideMovieCaptionError {
    if let Some((observed, maximum)) = error.message_byte_limit_values() {
        return SlideMovieCaptionError::LimitExceeded {
            kind: SlideMovieCaptionLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideMovieCaptionError::LimitExceeded {
            kind: SlideMovieCaptionLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideMovieCaptionError::LimitExceeded {
            kind: SlideMovieCaptionLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    SlideMovieCaptionError::InvalidSource
}

fn map_slide_text_error(error: super::slide_text::SlideTextError) -> SlideMovieCaptionError {
    match error {
        super::slide_text::SlideTextError::UnsupportedSource => {
            SlideMovieCaptionError::UnsupportedSource
        },
        super::slide_text::SlideTextError::DependentContent
        | super::slide_text::SlideTextError::ObjectMarkerReplacement => {
            SlideMovieCaptionError::UnsupportedDependency
        },
        super::slide_text::SlideTextError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => SlideMovieCaptionError::LimitExceeded {
            kind: map_slide_text_limit_kind(kind),
            observed,
            maximum,
        },
        super::slide_text::SlideTextError::Allocation { amount } => {
            SlideMovieCaptionError::Allocation { amount }
        },
        super::slide_text::SlideTextError::Verification => SlideMovieCaptionError::Verification,
        _ => SlideMovieCaptionError::InvalidSource,
    }
}

fn map_chart_caption_error(
    error: super::slide_chart_caption::ChartCaptionError,
) -> SlideMovieCaptionError {
    use super::slide_chart_caption::ChartCaptionError;
    match error {
        ChartCaptionError::UnsupportedSource => SlideMovieCaptionError::UnsupportedSource,
        ChartCaptionError::UnsupportedDependency => SlideMovieCaptionError::UnsupportedDependency,
        ChartCaptionError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => SlideMovieCaptionError::LimitExceeded {
            kind: map_chart_caption_limit_kind(kind),
            observed,
            maximum,
        },
        ChartCaptionError::Allocation { amount } => SlideMovieCaptionError::Allocation { amount },
        ChartCaptionError::Verification => SlideMovieCaptionError::Verification,
        _ => SlideMovieCaptionError::InvalidSource,
    }
}

fn map_rendering_error(
    _error: super::rendering_invalidation::RenderingInvalidationError,
) -> SlideMovieCaptionError {
    SlideMovieCaptionError::InvalidSource
}

fn map_chart_caption_limit_kind(
    kind: super::slide_chart_caption::ChartCaptionLimitKind,
) -> SlideMovieCaptionLimitKind {
    use super::slide_chart_caption::ChartCaptionLimitKind;
    match kind {
        ChartCaptionLimitKind::InputBytes => SlideMovieCaptionLimitKind::InputBytes,
        ChartCaptionLimitKind::OutputBytes => SlideMovieCaptionLimitKind::OutputBytes,
        ChartCaptionLimitKind::WireBytes => SlideMovieCaptionLimitKind::WireBytes,
        ChartCaptionLimitKind::Entries => SlideMovieCaptionLimitKind::Entries,
        ChartCaptionLimitKind::EntryBytes => SlideMovieCaptionLimitKind::EntryBytes,
        ChartCaptionLimitKind::TotalBytes => SlideMovieCaptionLimitKind::TotalBytes,
        ChartCaptionLimitKind::Slides => SlideMovieCaptionLimitKind::Slides,
        ChartCaptionLimitKind::References => SlideMovieCaptionLimitKind::References,
        ChartCaptionLimitKind::TextStorages => SlideMovieCaptionLimitKind::TextStorages,
        ChartCaptionLimitKind::TextFragments => SlideMovieCaptionLimitKind::TextFragments,
        ChartCaptionLimitKind::TextBytes => SlideMovieCaptionLimitKind::TextBytes,
        ChartCaptionLimitKind::WireFields => SlideMovieCaptionLimitKind::WireFields,
        ChartCaptionLimitKind::WireNesting => SlideMovieCaptionLimitKind::WireNesting,
        ChartCaptionLimitKind::WireWork => SlideMovieCaptionLimitKind::WireWork,
        ChartCaptionLimitKind::CaptionBytes => SlideMovieCaptionLimitKind::CaptionBytes,
    }
}

fn map_slide_text_limit_kind(
    kind: super::slide_text::SlideTextLimitKind,
) -> SlideMovieCaptionLimitKind {
    use super::slide_text::SlideTextLimitKind;
    match kind {
        SlideTextLimitKind::InputBytes => SlideMovieCaptionLimitKind::InputBytes,
        SlideTextLimitKind::OutputBytes => SlideMovieCaptionLimitKind::OutputBytes,
        SlideTextLimitKind::Entries => SlideMovieCaptionLimitKind::Entries,
        SlideTextLimitKind::EntryBytes => SlideMovieCaptionLimitKind::EntryBytes,
        SlideTextLimitKind::TotalBytes => SlideMovieCaptionLimitKind::TotalBytes,
        SlideTextLimitKind::Slides => SlideMovieCaptionLimitKind::Slides,
        SlideTextLimitKind::References => SlideMovieCaptionLimitKind::References,
        SlideTextLimitKind::TextStorages => SlideMovieCaptionLimitKind::TextStorages,
        SlideTextLimitKind::TextFragments => SlideMovieCaptionLimitKind::TextFragments,
        SlideTextLimitKind::TextBytes => SlideMovieCaptionLimitKind::TextBytes,
        SlideTextLimitKind::TextUnits => SlideMovieCaptionLimitKind::WireWork,
        SlideTextLimitKind::WireBytes => SlideMovieCaptionLimitKind::WireBytes,
        SlideTextLimitKind::WireFields => SlideMovieCaptionLimitKind::WireFields,
        SlideTextLimitKind::WireNesting => SlideMovieCaptionLimitKind::WireNesting,
        SlideTextLimitKind::WireWork => SlideMovieCaptionLimitKind::WireWork,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
