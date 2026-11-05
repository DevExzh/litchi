//! Exact-source, selector-first properties transactions for slide media.
//!
//! This owner deliberately treats a media drawable as a `MovieArchive` and
//! counts every such archive in slide source order.  File movies and
//! independently positioned audio therefore share one typed selector space;
//! selecting an unsupported movie kind for mutation is reported as a graph
//! error instead of silently changing the meaning of a selector.  Only the
//! four user-facing drawable-property fields are projected.  The read path
//! accepts every classified movie kind, including sparse placeholder and
//! live-video archives; mutation remains limited to file movies and audio
//! controls.  The Buffa codec and archive reassembly seam preserve all other
//! movie fields and their original wire bytes.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The semantic boundary redacts lower-layer failure details."
)]

use std::collections::HashSet;
use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::{
    decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireDescent, WireFieldView, WireView, preflight_wire_tree},
};
use litchi_iwa_core::{
    ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceScope, ArchiveReferenceVisitor, FieldType, MessageInfo,
};
use litchi_iwa_protos::keynote_media_properties_codec;
use thiserror::Error as ThisError;

use super::slide_media_lifecycle::graph_caption_witness::CaptionWitnessBudget;
use super::slide_movie_geometry::{
    GeometryBudget, physical_catalog, previews_absent, rewrite_geometry_message, verify_locality,
};
use super::{Package, ReadError, SemanticLimitKind};
use crate::slide::media::{MediaProperties, MovieKind};
use crate::{MovieSelector, SlideSelector};

mod data;

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const MOVIE_SUPER_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const MOVIE_AUDIO_ONLY_FIELD: u32 = 9;
const MOVIE_FLAGS_FIELD: u32 = 13;
const MOVIE_LIVE_VIDEO_FIELD: u32 = 30;
const MOVIE_COMMENT_FIELD: u32 = 6;
const MOVIE_TITLE_FIELD: u32 = 10;
const MOVIE_CAPTION_FIELD: u32 = 11;
const MOVIE_STYLE_FIELD: u32 = 19;
const MOVIE_COMMENT_MESSAGE_TYPE: u32 = 3_056;
const MOVIE_DATA_FIELD: u32 = 14;
const MOVIE_POSTER_DATA_FIELD: u32 = 15;
const CONTAINER_MESSAGE_TYPE: u32 = 3_003;
const GROUP_MESSAGE_TYPE: u32 = 3_008;
const MOVIE_STANDIN_MESSAGE_TYPE: u32 = 3_097;
const MOVIE_CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const MOVIE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const MOVIE_AUDIO_STYLE_MESSAGE_TYPE: u32 = 3_016;
const SLIDE_NAME_FIELD: u32 = 10;
const MAX_PROPERTY_BYTES: usize = 64 * 1024 * 1024;

impl CaptionWitnessBudget for GeometryBudget {
    type Error = Error;

    fn invalid_source() -> Self::Error {
        Error::InvalidSource
    }

    fn charge_codec_pass(&mut self, payload: &[u8]) -> Result<(), Self::Error> {
        self.fields(payload.len().max(1)).map_err(Error::from)?;
        self.work(
            payload
                .len()
                .checked_mul(32)
                .ok_or(Error::InvalidSource)?
                .max(1),
        )
        .map_err(Error::from)?;
        self.allocations(1).map_err(Error::from)?;
        self.scratch(payload.len().max(1)).map_err(Error::from)
    }

    fn charge_header_metadata(&mut self, info: &MessageInfo) -> Result<(), Self::Error> {
        let fields = info
            .object_references
            .len()
            .checked_add(info.data_references.len())
            .and_then(|count| count.checked_add(info.field_infos.len()))
            .ok_or_else(Self::invalid_source)?;
        let references = info
            .object_references
            .len()
            .checked_add(info.data_references.len())
            .and_then(|count| {
                info.field_infos.iter().try_fold(count, |count, field| {
                    count
                        .checked_add(field.object_references.len())
                        .and_then(|count| count.checked_add(field.data_references.len()))
                })
            })
            .ok_or_else(Self::invalid_source)?;
        self.fields(fields).map_err(Error::from)?;
        self.references(references).map_err(Error::from)?;
        self.work(fields.checked_add(references).ok_or(Error::InvalidSource)?)
            .map_err(Error::from)
    }
}

/// Resource categories reported by a media-properties transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
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
    /// Bytes in the selected movie-property payload.
    PropertyBytes,
    /// Logical transaction allocations.
    Allocations,
    /// Retained transaction bytes.
    Retained,
    /// Scratch transaction bytes.
    Scratch,
    /// Rewritten physical components.
    Components,
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
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::PropertyBytes => "media property bytes",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
        })
    }
}

/// Content-redacted failure raised by a media-properties transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The selected media graph or payload was malformed.
    #[error("the Keynote media-properties source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Keynote media properties {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote media-properties transaction")]
    Allocation { amount: usize },
    /// Candidate reopening did not reproduce the staged properties and locality.
    #[error("the edited Keynote media properties failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote media-properties patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The source was not retained as an exact physical package.
    #[error("this Keynote source does not support physical media-property edits")]
    UnsupportedSource,
    /// The selected media graph is not safely owned by one component.
    #[error("the requested Keynote media-property graph is unsupported")]
    UnsupportedDependency,
    /// A selector matched more than one semantic object.
    #[error("the Keynote media-property selector is ambiguous")]
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
    /// No movie archive existed at a checked source-order position.
    #[error("the selected Keynote slide has no media at position {position:?}")]
    MoviePositionNotFound { position: Position },
    /// The selected media is not supported by the mutation path.
    #[error("the selected Keynote media kind is unsupported for properties")]
    WrongMediaKind,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::WrongMediaKind { .. } => RefusalKind::UnsupportedTarget,
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

/// One mutable media-properties value staged against an immutable package
/// snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    budget: GeometryBudget,
    selection: MediaPropertiesSelection,
    before: Arc<MediaProperties>,
    after: Arc<MediaProperties>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field("before", self.before.as_ref())
            .field("after", self.after.as_ref())
            .finish()
    }
}

impl<'a> Edit<'a> {
    fn new<'slide>(
        source: &'a Package,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<Self, Error> {
        let mut budget = GeometryBudget::new(source)?;
        budget.source(physical_catalog(source)?.source_bytes().len())?;
        let selection = select_media_properties(source, slide.into(), movie.into(), &mut budget)?;
        let before = Arc::clone(&selection.before);
        Ok(Self {
            source,
            budget,
            selection,
            before: Arc::clone(&before),
            after: before,
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected source-order media position.
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }

    /// Return the properties observed when this edit began.
    #[must_use]
    pub fn before(&self) -> &MediaProperties {
        self.before.as_ref()
    }

    /// Return the properties currently staged for publication.
    #[must_use]
    pub fn after(&self) -> &MediaProperties {
        self.after.as_ref()
    }

    /// Stage a complete replacement of the four modeled properties.
    pub fn set(mut self, properties: MediaProperties) -> Result<Self, Error> {
        validate_properties(&properties)?;
        self.budget.allocations(1)?;
        let adopted_bytes = properties
            .hyperlink_url()
            .map_or(0, str::len)
            .checked_add(properties.accessibility_description().map_or(0, str::len))
            .and_then(|bytes| bytes.checked_add(size_of::<MediaProperties>()))
            .ok_or(Error::InvalidSource)?;
        self.budget.retained(adopted_bytes)?;
        self.after = Arc::new(properties);
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(
            self.source,
            &self.selection,
            self.before,
            self.after,
            self.budget,
        )
    }
}

/// An exact-source checked reversible media-properties patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    selection: MediaPropertiesSelection,
    before: Arc<MediaProperties>,
    after: Arc<MediaProperties>,
    deleted_previews: usize,
    restored_previews: usize,
    source_previews_absent: bool,
    target_previews_absent: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the properties required from the source package.
    #[must_use]
    pub fn before(&self) -> &MediaProperties {
        self.before.as_ref()
    }

    /// Return the properties produced by this patch.
    #[must_use]
    pub fn after(&self) -> &MediaProperties {
        self.after.as_ref()
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected source-order media position.
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }

    /// Return the source package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch is an exact source no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from the target back to the source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            deleted_previews: self.restored_previews,
            restored_previews: self.deleted_previews,
            source_previews_absent: self.target_previews_absent,
            target_previews_absent: self.source_previews_absent,
        }
    }
}

/// Compact evidence describing one media-properties publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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

    /// Return how many stale root previews were deleted.
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

/// Fully verified result of one media-properties transaction.
#[must_use = "a media-properties commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
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

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq)]
struct MediaPropertiesSelection {
    slide_position: Position,
    movie_position: Position,
    slide_identifier: u64,
    node_identifier: u64,
    movie_identifier: u64,
    message_index: usize,
    slide_component_name: Arc<str>,
    kind: MovieKind,
    before: Arc<MediaProperties>,
}

impl fmt::Debug for MediaPropertiesSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("MediaPropertiesSelection")
            .field("slide_position", &self.slide_position)
            .field("movie_position", &self.movie_position)
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Read the four modeled properties of one source-order slide media item.
    ///
    /// The selector counts every native `MovieArchive` sibling in source
    /// order.  This keeps file movies and audio controls in one stable index
    /// and avoids an audio-only renumbering that would make a caller edit the
    /// wrong object after a producer changes sibling order.
    ///
    /// Reading validates ownership, references, and the selected property
    /// fields without loading media assets. Properties therefore remain
    /// readable when a sparse source omits media content. Editing additionally
    /// requires a file movie or audio control with valid, materialized media
    /// assets before publication.
    pub fn slide_media_properties<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<MediaProperties, Error> {
        let mut budget = GeometryBudget::new(self)?;
        budget.source(physical_catalog(self)?.source_bytes().len())?;
        let selection =
            select_media_property_fields(self, slide.into(), movie.into(), &mut budget)?;
        clone_properties(selection.before.as_ref(), &mut budget)
    }

    /// Begin an exact immutable edit of one source-order slide media item.
    ///
    /// Start from the current value to preserve the other properties:
    ///
    /// ```no_run
    /// use litchi_keynote::{MovieSelector, Package, SlideSelector};
    /// # fn update(package: &Package) -> Result<(), Box<dyn std::error::Error>> {
    /// let slide = SlideSelector::index(0);
    /// let movie = MovieSelector::index(0);
    /// let properties = package.slide_media_properties(slide, movie)?
    ///     .with_accessibility_description(Some("Opening narration".to_owned()));
    /// let commit = package.edit_slide_media_properties(slide, movie)?
    ///     .set(properties)?
    ///     .commit()?;
    /// let restored = commit.package()
    ///     .apply_slide_media_properties(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    pub fn edit_slide_media_properties<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<Edit<'_>, Error> {
        Edit::new(self, slide, movie)
    }

    /// Apply an exact-source checked media-properties patch.
    pub fn apply_slide_media_properties(&self, patch: &Patch) -> Result<Commit, Error> {
        let mut budget = GeometryBudget::new(self)?;
        let catalog = physical_catalog(self)?;
        let source_len = catalog.source_bytes().len();
        budget.source(source_len)?;
        budget.work(source_len)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        if previews_absent(self)? != patch.source_previews_absent {
            return Err(Error::PatchConflict);
        }
        let current = select_media_properties(
            self,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            budget.validate_package(self)?;
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        reopen_target_patch(self, patch, &mut budget)
    }
}

fn commit_edit(
    source: &Package,
    selection: &MediaPropertiesSelection,
    before: Arc<MediaProperties>,
    after: Arc<MediaProperties>,
    mut budget: GeometryBudget,
) -> Result<Commit, Error> {
    validate_properties(after.as_ref())?;
    if before == after {
        let catalog = physical_catalog(source)?;
        budget.validate_package(source)?;
        budget.work(source.source_bytes().len())?;
        let bytes = catalog.shared_source();
        let previews_absent = previews_absent(source)?;
        return Ok(Commit {
            package: source.snapshot(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: Arc::clone(&before),
                after: Arc::clone(&after),
                deleted_previews: 0,
                restored_previews: 0,
                source_previews_absent: previews_absent,
                target_previews_absent: previews_absent,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    let catalog = physical_catalog(source)?;
    if !catalog.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let current = select_media_properties(
        source,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
        &mut budget,
    )?;
    if !same_selection(&current, selection) || current.before != before {
        return Err(Error::PatchConflict);
    }
    let (candidate, deleted_previews) = rewrite_geometry_message(
        source,
        selection.slide_component_name.as_ref(),
        selection.movie_identifier,
        selection.message_index,
        |original, budget| {
            let options =
                property_codec_options(source, original, budget).map_err(to_geometry_error)?;
            let write = keynote_media_properties_codec::MoviePropertiesWrite::new(
                after.hyperlink_url(),
                after.locked(),
                after.aspect_ratio_locked(),
                after.accessibility_description(),
            );
            let prepared = keynote_media_properties_codec::prepare_movie_properties_rewrite(
                original, write, options,
            )
            .map_err(map_properties_codec_error)
            .map_err(to_geometry_error)?;
            budget_property_report(budget, prepared.prepare_report()).map_err(to_geometry_error)?;
            let requirements = prepared.execution_requirements();
            budget_property_requirements(budget, requirements).map_err(to_geometry_error)?;
            prepared
                .execute(requirements.exact_limits())
                .map_err(map_properties_codec_error)
                .map_err(to_geometry_error)
                .map(|output| output.into_output())
        },
        &mut budget,
    )?;
    budget.validate_package(&candidate)?;
    if !previews_absent(&candidate)? {
        return Err(Error::Verification);
    }
    let selected = select_media_properties(
        &candidate,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
        &mut budget,
    )?;
    if !same_selection(&selected, selection) || selected.before.as_ref() != after.as_ref() {
        return Err(Error::Verification);
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
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: ExactArtifacts::new(source_bytes, target),
            selection: selection.clone(),
            before,
            after,
            deleted_previews,
            restored_previews: 0,
            source_previews_absent: previews_absent(source)?,
            target_previews_absent: true,
        },
        diagnostics: Diagnostics::published(deleted_previews),
    })
}

fn reopen_target_patch(
    source: &Package,
    patch: &Patch,
    budget: &mut GeometryBudget,
) -> Result<Commit, Error> {
    budget.candidate_reopen(patch.artifacts.target().len())?;
    budget.allocations(1)?;
    budget.retained(patch.artifacts.target().len())?;
    budget.scratch(patch.artifacts.target().len())?;
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    budget.validate_package(&candidate)?;
    if previews_absent(&candidate)? != patch.target_previews_absent {
        return Err(Error::Verification);
    }
    let selected = select_media_properties(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        MovieSelector::position(patch.selection.movie_position),
        budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(Error::Verification);
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
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics::published(patch.deleted_previews),
    })
}

fn select_media_properties(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    budget: &mut GeometryBudget,
) -> Result<MediaPropertiesSelection, Error> {
    let selection = select_media_property_fields(package, slide_selector, movie_selector, budget)?;
    // Placeholder and live-video archives expose the same four semantic
    // drawable fields, but their media closure is owned by the slide layout
    // or camera subsystem. Keep their read path available while refusing to
    // route those records through a file/audio asset rewrite transaction.
    if !matches!(selection.kind, MovieKind::File | MovieKind::Audio) {
        return Err(Error::Refused(Refusal::WrongMediaKind));
    }
    let assets = data::validate_selected_media_assets(
        package,
        selection.slide_position,
        selection.movie_position,
        budget,
    )?;
    if assets.content.is_empty() || assets.poster.is_some_and(|poster| poster.is_empty()) {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(selection)
}

fn select_media_property_fields(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    budget: &mut GeometryBudget,
) -> Result<MediaPropertiesSelection, Error> {
    let slide_position = resolve_slide_position(package, slide_selector, budget)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(Error::Refused(Refusal::SlidePositionNotFound {
            position: slide_position,
        }))?;
    let (component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(Error::InvalidSource)?;
    let (slide_message_index, slide_payload) = unique_message(slide, SLIDE_MESSAGE_TYPE, budget)?;
    let limits = package.semantic_wire_limits().map_err(map_read_error)?;
    let drawable_ids =
        repeated_references(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits, budget)?;
    validate_slide_header(
        slide,
        record.slide_identifier,
        slide_message_index,
        &drawable_ids,
        budget,
    )?;

    let requested_movie_position = movie_selector.as_position();
    let mut movie_position = 0usize;
    let mut selected = None;
    for identifier in drawable_ids.iter().copied() {
        let (movie_component, object) = package
            .object_with_component(identifier)
            .ok_or(Error::InvalidSource)?;
        let Some((message_index, payload)) = optional_message(object, MOVIE_MESSAGE_TYPE, budget)?
        else {
            continue;
        };
        let current_position = Position::new(movie_position);
        movie_position = movie_position.checked_add(1).ok_or(Error::InvalidSource)?;
        if current_position != requested_movie_position {
            continue;
        }
        if movie_component != component_name {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        if object.archive_info.identifier != Some(identifier) {
            return Err(Error::InvalidSource);
        }
        let kind = classify_movie_kind(payload, limits, budget)?;
        if movie_parent(payload, limits, budget)? != record.slide_identifier {
            return Err(Error::InvalidSource);
        }
        super::slide_media_replacement::validate_selected_message_metadata(object, message_index)
            .map_err(|_| Error::InvalidSource)?;
        validate_selected_movie_references(
            package,
            component_name,
            object,
            message_index,
            payload,
            limits,
            budget,
        )?;
        let options = property_codec_options(package, payload, budget)?;
        let (snapshot, report) =
            keynote_media_properties_codec::decode_movie_properties_with_report(payload, options)
                .map_err(map_properties_codec_error)?;
        budget_property_report(budget, report)?;
        let before = shared_properties(owned_properties(snapshot, budget)?, budget)?;
        selected = Some((identifier, message_index, kind, before));
    }
    let (movie_identifier, message_index, kind, before) = selected.ok_or_else(|| {
        if requested_movie_position.get() >= movie_position {
            Error::Refused(Refusal::MoviePositionNotFound {
                position: requested_movie_position,
            })
        } else {
            Error::InvalidSource
        }
    })?;
    budget.allocations(1)?;
    budget.retained(component_name.len())?;
    validate_single_drawable_ownership(
        package,
        record.slide_identifier,
        movie_identifier,
        component_name,
        limits,
        budget,
    )?;
    Ok(MediaPropertiesSelection {
        slide_position,
        movie_position: requested_movie_position,
        slide_identifier: record.slide_identifier,
        node_identifier: record.node_identifier,
        movie_identifier,
        message_index,
        slide_component_name: Arc::from(component_name),
        kind,
        before,
    })
}

fn same_selection(left: &MediaPropertiesSelection, right: &MediaPropertiesSelection) -> bool {
    left.slide_position == right.slide_position
        && left.movie_position == right.movie_position
        && left.slide_identifier == right.slide_identifier
        && left.node_identifier == right.node_identifier
        && left.movie_identifier == right.movie_identifier
        && left.message_index == right.message_index
        && left.slide_component_name == right.slide_component_name
        && left.kind == right.kind
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
    budget: &mut GeometryBudget,
) -> Result<Position, Error> {
    match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(map_read_error)?
            .map(|_| position)
            .ok_or(Error::Refused(Refusal::SlidePositionNotFound { position })),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(Error::Refused(Refusal::EmptySlideName));
            }
            let maximum = package.semantic_limits().max_slides();
            let mut match_position = None;
            for index in 0..maximum {
                let position = Position::new(index);
                let Some(record) = package.slide_record_at(index).map_err(map_read_error)? else {
                    break;
                };
                let (_, slide) = package
                    .object_with_component(record.slide_identifier)
                    .ok_or(Error::InvalidSource)?;
                let (_, payload) = unique_message(slide, SLIDE_MESSAGE_TYPE, budget)?;
                let limits = package.semantic_wire_limits().map_err(map_read_error)?;
                if slide_name(payload, limits, budget)?.is_some_and(|value| value == name) {
                    if match_position.replace(position).is_some() {
                        return Err(Error::Refused(Refusal::AmbiguousSelector));
                    }
                }
            }
            match_position.ok_or(Error::Refused(Refusal::SlideNameNotFound))
        },
    }
}

fn slide_name<'a>(
    payload: &'a [u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<Option<&'a str>, Error> {
    let view = parse_wire(payload, limits, budget)?;
    let mut name = None;
    for field in view
        .fields()
        .filter(|field| field.number() == SLIDE_NAME_FIELD)
    {
        if name.is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        name = Some(std::str::from_utf8(field.payload()).map_err(|_| Error::InvalidSource)?);
    }
    Ok(name)
}

fn parse_wire<'a>(
    payload: &'a [u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<WireView<'a>, Error> {
    // Skip visits inspect this payload once without descending into opaque
    // length-delimited fields. Reserve that work before the preflight.
    budget.work(payload.len())?;
    let report = preflight_wire_tree(payload, limits, |_visit| Ok(WireDescent::Skip))
        .map_err(|_| Error::InvalidSource)?;
    budget.fields(report.fields())?;
    budget.nesting(report.max_depth())?;
    let field_storage = report
        .fields()
        .checked_mul(size_of::<WireFieldView<'static>>())
        .and_then(|bytes| bytes.checked_mul(2))
        .ok_or(Error::InvalidSource)?;
    // WireView may grow its span vector while parsing; one reserve per field
    // is a conservative ceiling for all allocation attempts.
    budget.allocations(report.fields())?;
    budget.scratch(field_storage)?;
    budget.fields(report.fields())?;
    budget.work(payload.len())?;
    let view = WireView::parse(payload, limits).map_err(|_| Error::InvalidSource)?;
    Ok(view)
}

fn unique_message<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    budget: &mut GeometryBudget,
) -> Result<(usize, &'a [u8]), Error> {
    optional_message(object, message_type, budget)?.ok_or(Error::InvalidSource)
}

fn optional_message<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    budget: &mut GeometryBudget,
) -> Result<Option<(usize, &'a [u8])>, Error> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(Error::InvalidSource)?;
        if info.type_ != message.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(Error::InvalidSource);
        }
        budget.work(
            message
                .data
                .len()
                .checked_add(1)
                .ok_or(Error::InvalidSource)?,
        )?;
        if message.type_ != message_type {
            continue;
        }
        if found.replace((index, message.data.as_slice())).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(found)
}

fn validate_slide_header(
    slide: &ArchiveObject,
    slide_identifier: u64,
    message_index: usize,
    drawable_ids: &[u64],
    budget: &mut GeometryBudget,
) -> Result<(), Error> {
    if slide.archive_info.identifier != Some(slide_identifier) {
        return Err(Error::InvalidSource);
    }
    let Some(info) = slide.archive_info.message_infos.get(message_index) else {
        return Err(Error::InvalidSource);
    };
    if info.type_ != SLIDE_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
    }
    let mut seen = HashSet::new();
    budget.allocations(usize::from(!info.object_references.is_empty()))?;
    let seen_storage = info
        .object_references
        .len()
        .checked_mul(size_of::<u64>())
        .and_then(|bytes| bytes.checked_mul(2))
        .ok_or(Error::InvalidSource)?;
    budget.retained(seen_storage)?;
    seen.try_reserve(info.object_references.len())
        .map_err(|_| Error::Allocation {
            amount: info.object_references.len(),
        })?;
    for identifier in &info.object_references {
        if !seen.insert(*identifier) {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
    }
    for identifier in drawable_ids {
        if info
            .object_references
            .iter()
            .filter(|reference| *reference == identifier)
            .count()
            != 1
        {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
    }
    for field in &info.field_infos {
        if field.path.as_slice() == [SLIDE_OWNED_DRAWABLES_FIELD]
            && (field.object_references.as_slice() != drawable_ids
                || !field.data_references.is_empty())
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<Vec<u64>, Error> {
    let view = parse_wire(payload, limits, budget)?;
    let count = view
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    budget.allocations(usize::from(count != 0))?;
    let mut result = Vec::new();
    budget.retained(
        count
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?,
    )?;
    result
        .try_reserve_exact(count)
        .map_err(|_| Error::Allocation { amount: count })?;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        result.push(reference_identifier(field.payload(), limits, budget)?);
    }
    budget.references(count)?;
    Ok(result)
}

fn reference_identifier(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<u64, Error> {
    let view = parse_wire(payload, limits, budget)?;
    let mut identifier = None;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        if field.number() != 1 {
            return Err(Error::InvalidSource);
        }
        if identifier.is_some() || field.wire_type() != 0 {
            return Err(Error::InvalidSource);
        }
        let (value, width) =
            decode_varint_from_bytes(field.payload()).map_err(|_| Error::InvalidSource)?;
        if value == 0 || width != encoded_len(value) {
            return Err(Error::InvalidSource);
        }
        identifier = Some(value);
    }
    identifier.ok_or(Error::InvalidSource)
}

fn classify_movie_kind(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<MovieKind, Error> {
    let view = parse_wire(payload, limits, budget)?;
    let mut audio = None;
    let mut flags = None;
    let mut live = None;
    let mut super_fields = 0usize;
    for field in view.fields() {
        match field.number() {
            MOVIE_AUDIO_ONLY_FIELD => {
                audio = Some(unique_bool(field, audio, "audio")?);
            },
            MOVIE_FLAGS_FIELD => {
                flags = Some(unique_u32(field, flags, "flags")?);
            },
            MOVIE_LIVE_VIDEO_FIELD => {
                live = Some(unique_bool(field, live, "live")?);
            },
            MOVIE_SUPER_FIELD => {
                super_fields = super_fields.checked_add(1).ok_or(Error::InvalidSource)?;
                if super_fields > 1 {
                    return Err(Error::InvalidSource);
                }
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                field
                    .validate_canonical_framing()
                    .map_err(|_| Error::InvalidSource)?;
            },
            _ => {},
        }
    }
    if super_fields != 1 {
        return Err(Error::InvalidSource);
    }
    Ok(if live == Some(true) {
        MovieKind::LiveVideo
    } else if audio == Some(true) {
        MovieKind::Audio
    } else if flags.is_some_and(|value| value & 1 != 0) {
        MovieKind::Placeholder
    } else {
        MovieKind::File
    })
}

fn unique_bool(
    field: WireFieldView<'_>,
    current: Option<bool>,
    _context: &str,
) -> Result<bool, Error> {
    if current.is_some() || field.wire_type() != 0 {
        return Err(Error::InvalidSource);
    }
    field
        .validate_canonical_framing()
        .map_err(|_| Error::InvalidSource)?;
    let (value, width) =
        decode_varint_from_bytes(field.payload()).map_err(|_| Error::InvalidSource)?;
    if width != encoded_len(value) || value > 1 {
        return Err(Error::InvalidSource);
    }
    Ok(value != 0)
}

fn unique_u32(
    field: WireFieldView<'_>,
    current: Option<u32>,
    _context: &str,
) -> Result<u32, Error> {
    if current.is_some() || field.wire_type() != 0 {
        return Err(Error::InvalidSource);
    }
    field
        .validate_canonical_framing()
        .map_err(|_| Error::InvalidSource)?;
    let (value, width) =
        decode_varint_from_bytes(field.payload()).map_err(|_| Error::InvalidSource)?;
    if value > u64::from(u32::MAX) || width != encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    Ok(value as u32)
}

fn movie_parent(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<u64, Error> {
    let root = parse_wire(payload, limits, budget)?;
    let mut parent = None;
    for field in root
        .fields()
        .filter(|field| field.number() == MOVIE_SUPER_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let drawable = parse_wire(field.payload(), limits, budget)?;
        for child in drawable
            .fields()
            .filter(|child| child.number() == DRAWABLE_PARENT_FIELD)
        {
            if parent.is_some() || child.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
            child
                .validate_canonical_framing()
                .map_err(|_| Error::InvalidSource)?;
            parent = Some(reference_identifier(child.payload(), limits, budget)?);
        }
    }
    parent.ok_or(Error::InvalidSource)
}

fn validate_selected_movie_references(
    package: &Package,
    component_name: &str,
    movie_object: &ArchiveObject,
    message_index: usize,
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<(), Error> {
    let root = parse_wire(payload, limits, budget)?;
    let super_field = root
        .fields()
        .find(|field| field.number() == MOVIE_SUPER_FIELD)
        .ok_or(Error::InvalidSource)?;
    let drawable = parse_wire(super_field.payload(), limits, budget)?;
    budget.allocations(1)?;
    budget.retained(
        4usize
            .checked_mul(size_of::<(u64, u32)>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(4)
        .map_err(|_| Error::Allocation { amount: 4 })?;
    let mut style_seen = false;
    let mut comment_seen = false;
    let mut title_seen = false;
    let mut caption_seen = false;

    for field in root
        .fields()
        .filter(|field| field.number() == MOVIE_STYLE_FIELD)
    {
        if std::mem::replace(&mut style_seen, true) {
            return Err(Error::InvalidSource);
        }
        append_movie_reference(field, MOVIE_STYLE_FIELD, &mut references, limits, budget)?;
    }
    for field in drawable.fields().filter(|field| {
        matches!(
            field.number(),
            MOVIE_COMMENT_FIELD | MOVIE_TITLE_FIELD | MOVIE_CAPTION_FIELD
        )
    }) {
        let seen = if field.number() == MOVIE_TITLE_FIELD {
            &mut title_seen
        } else if field.number() == MOVIE_CAPTION_FIELD {
            &mut caption_seen
        } else {
            &mut comment_seen
        };
        if std::mem::replace(seen, true) {
            return Err(Error::InvalidSource);
        }
        append_movie_reference(field, field.number(), &mut references, limits, budget)?;
    }

    let data_references = movie_data_references(root, limits, budget)?;
    let info = movie_object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let extra_object_reference_count = info
        .object_references
        .len()
        .checked_sub(references.len())
        .ok_or(Error::InvalidSource)?;
    let style_witnesses = if extra_object_reference_count == 0 {
        [None, None]
    } else {
        super::slide_media_lifecycle::graph_caption_witness::prove_movie_caption_style_witness(
            package,
            component_name,
            movie_object,
            limits,
            budget,
        )?
    };
    let mut transitive_style_ids = [0_u64; 2];
    let mut transitive_style_count = 0usize;
    for identifier in style_witnesses.into_iter().flatten() {
        if references.iter().any(|(known, _)| *known == identifier) {
            continue;
        }
        if transitive_style_count == transitive_style_ids.len()
            || transitive_style_ids[..transitive_style_count].contains(&identifier)
        {
            return Err(Error::InvalidSource);
        }
        transitive_style_ids[transitive_style_count] = identifier;
        transitive_style_count += 1;
    }
    budget.references(transitive_style_count)?;
    let expected_object_reference_count = references
        .len()
        .checked_add(transitive_style_count)
        .ok_or(Error::InvalidSource)?;
    if info.object_references.len() != expected_object_reference_count
        || references.iter().any(|(identifier, _)| {
            info.object_references
                .iter()
                .filter(|known| *known == identifier)
                .count()
                != 1
        })
        || (0..transitive_style_count).any(|index| {
            let identifier = transitive_style_ids[index];
            info.object_references
                .iter()
                .filter(|known| **known == identifier)
                .count()
                != 1
        })
        || info.object_references.iter().any(|identifier| {
            !references.iter().any(|(known, _)| *known == *identifier)
                && !transitive_style_ids[..transitive_style_count].contains(identifier)
        })
    {
        return Err(Error::InvalidSource);
    }
    validate_selected_movie_header_metadata(info, &references, &data_references)?;

    for (identifier, field_number) in references {
        let (reference_component, referenced) = package
            .object_with_component(identifier)
            .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
        if reference_component != component_name && field_number != MOVIE_STYLE_FIELD {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        let allowed = match field_number {
            MOVIE_TITLE_FIELD | MOVIE_CAPTION_FIELD => {
                &[MOVIE_STANDIN_MESSAGE_TYPE, MOVIE_CAPTION_INFO_MESSAGE_TYPE][..]
            },
            MOVIE_COMMENT_FIELD => &[MOVIE_COMMENT_MESSAGE_TYPE][..],
            MOVIE_STYLE_FIELD => &[MOVIE_STYLE_MESSAGE_TYPE, MOVIE_AUDIO_STYLE_MESSAGE_TYPE][..],
            _ => return Err(Error::InvalidSource),
        };
        let Some(_) = reference_message_type(referenced, allowed, budget)? else {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        };
    }
    for &identifier in &transitive_style_ids[..transitive_style_count] {
        let (reference_component, referenced) = package
            .object_with_component(identifier)
            .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
        if reference_component != component_name {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        let Some(_) = reference_message_type(referenced, &[MOVIE_STYLE_MESSAGE_TYPE], budget)?
        else {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        };
    }
    Ok(())
}

fn append_movie_reference(
    field: WireFieldView<'_>,
    expected_number: u32,
    references: &mut Vec<(u64, u32)>,
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<(), Error> {
    if field.number() != expected_number || field.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    field
        .validate_canonical_framing()
        .map_err(|_| Error::InvalidSource)?;
    let identifier = reference_identifier(field.payload(), limits, budget)?;
    if references.iter().any(|(known, _)| *known == identifier) {
        return Err(Error::InvalidSource);
    }
    budget.references(1)?;
    references.push((identifier, expected_number));
    Ok(())
}

fn movie_data_references(
    root: WireView<'_>,
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<Vec<(u32, u64)>, Error> {
    let count = root
        .fields()
        .filter(|field| matches!(field.number(), MOVIE_DATA_FIELD | MOVIE_POSTER_DATA_FIELD))
        .count();
    budget.allocations(usize::from(count != 0))?;
    budget.retained(
        count
            .checked_mul(size_of::<(u32, u64)>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(count)
        .map_err(|_| Error::Allocation { amount: count })?;
    for field in root
        .fields()
        .filter(|field| matches!(field.number(), MOVIE_DATA_FIELD | MOVIE_POSTER_DATA_FIELD))
    {
        if field.wire_type() != 2
            || references
                .iter()
                .any(|(number, _)| *number == field.number())
        {
            return Err(Error::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        let identifier = reference_identifier(field.payload(), limits, budget)?;
        budget.references(1)?;
        references.push((field.number(), identifier));
    }
    Ok(references)
}

fn validate_selected_movie_header_metadata(
    info: &MessageInfo,
    references: &[(u64, u32)],
    data_references: &[(u32, u64)],
) -> Result<(), Error> {
    if info.data_references.len() != data_references.len()
        || data_references.iter().any(|(_, identifier)| {
            info.data_references
                .iter()
                .filter(|known| **known == *identifier)
                .count()
                != 1
        })
    {
        return Err(Error::InvalidSource);
    }

    for field in &info.field_infos {
        if !field.object_references.is_empty() {
            let expected = match field.path.as_slice() {
                [MOVIE_SUPER_FIELD, MOVIE_COMMENT_FIELD] => Some(MOVIE_COMMENT_FIELD),
                [MOVIE_SUPER_FIELD, MOVIE_TITLE_FIELD] => Some(MOVIE_TITLE_FIELD),
                [MOVIE_SUPER_FIELD, MOVIE_CAPTION_FIELD] => Some(MOVIE_CAPTION_FIELD),
                [MOVIE_STYLE_FIELD] => Some(MOVIE_STYLE_FIELD),
                _ => None,
            };
            let Some(expected_number) = expected else {
                return Err(Error::InvalidSource);
            };
            let Some((identifier, _)) = references
                .iter()
                .find(|(_, number)| *number == expected_number)
            else {
                return Err(Error::InvalidSource);
            };
            if field.object_references.as_slice() != [*identifier] {
                return Err(Error::InvalidSource);
            }
        }
        if !field.data_references.is_empty() {
            let Some(expected_number) = (match field.path.as_slice() {
                [MOVIE_DATA_FIELD] => Some(MOVIE_DATA_FIELD),
                [MOVIE_POSTER_DATA_FIELD] => Some(MOVIE_POSTER_DATA_FIELD),
                _ => None,
            }) else {
                return Err(Error::InvalidSource);
            };
            if field.r#type != Some(FieldType::DataReference) || field.data_references.len() != 1 {
                return Err(Error::InvalidSource);
            }
            let Some((_, identifier)) = data_references
                .iter()
                .find(|(number, _)| *number == expected_number)
            else {
                return Err(Error::InvalidSource);
            };
            if field.data_references.as_slice() != [*identifier] {
                return Err(Error::InvalidSource);
            }
        }
    }
    Ok(())
}

fn reference_message_type(
    object: &ArchiveObject,
    allowed: &[u32],
    budget: &mut GeometryBudget,
) -> Result<Option<u32>, Error> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(Error::InvalidSource)?;
        if info.type_ != message.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(Error::InvalidSource);
        }
        budget.work(
            message
                .data
                .len()
                .checked_add(1)
                .ok_or(Error::InvalidSource)?,
        )?;
        if !allowed.contains(&message.type_) {
            continue;
        }
        if found.replace(message.type_).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(found)
}

fn validate_single_drawable_ownership(
    package: &Package,
    slide_identifier: u64,
    movie_identifier: u64,
    expected_component: &str,
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<(), Error> {
    let mut object_matches = 0usize;
    let mut object_component_matches = false;
    let mut total_references = 0usize;
    let mut selected_references = 0usize;
    let (_, slide_object) = package
        .object_with_component(slide_identifier)
        .ok_or(Error::InvalidSource)?;
    let slide_message_index = slide_object
        .archive_info
        .message_infos
        .iter()
        .enumerate()
        .find_map(|(index, info)| (info.type_ == SLIDE_MESSAGE_TYPE).then_some(index))
        .ok_or(Error::InvalidSource)?;
    let expected_metadata_references = slide_object
        .archive_info
        .message_infos
        .get(slide_message_index)
        .ok_or(Error::InvalidSource)?
        .object_references
        .iter()
        .filter(|identifier| **identifier == movie_identifier)
        .count()
        .checked_add(
            slide_object
                .archive_info
                .message_infos
                .get(slide_message_index)
                .ok_or(Error::InvalidSource)?
                .field_infos
                .iter()
                .flat_map(|field| field.object_references.iter())
                .filter(|identifier| **identifier == movie_identifier)
                .count(),
        )
        .ok_or(Error::InvalidSource)?;
    let mut metadata_census = MovieInboundMetadataCensus {
        package,
        movie_identifier,
        slide_identifier,
        slide_message_index,
        selected_references: 0,
        invalid: false,
    };
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;

    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let inventory_work = object
                .messages
                .len()
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            budget.work(inventory_work)?;
            let object_identifier = object.archive_info.identifier;
            if object.messages.len() != object.archive_info.message_infos.len() {
                return Err(Error::InvalidSource);
            }
            let mut metadata_fields = 0usize;
            let mut metadata_references = 0usize;
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                if message.type_ != info.type_
                    || usize::try_from(info.length).ok() != Some(message.data.len())
                {
                    return Err(Error::InvalidSource);
                }
                metadata_fields = metadata_fields
                    .checked_add(info.field_infos.len())
                    .ok_or(Error::InvalidSource)?;
                metadata_references = metadata_references
                    .checked_add(info.object_references.len())
                    .and_then(|value| value.checked_add(info.data_references.len()))
                    .ok_or(Error::InvalidSource)?;
                for field in &info.field_infos {
                    metadata_references = metadata_references
                        .checked_add(field.object_references.len())
                        .and_then(|value| value.checked_add(field.data_references.len()))
                        .ok_or(Error::InvalidSource)?;
                }
            }
            budget.fields(metadata_fields)?;
            budget.references(metadata_references)?;
            budget.work(
                metadata_fields
                    .checked_add(metadata_references)
                    .ok_or(Error::InvalidSource)?,
            )?;
            object
                .inspect_references_with_policy(
                    &mut metadata_census,
                    ArchiveReferencePolicy::KnownReferences,
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
            if object_identifier == Some(movie_identifier) {
                object_matches = object_matches.checked_add(1).ok_or(Error::InvalidSource)?;
                object_component_matches = component.name() == expected_component;
            }
            if !object
                .messages
                .iter()
                .any(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            {
                continue;
            }
            if object.messages.len() != object.archive_info.message_infos.len() {
                return Err(Error::InvalidSource);
            }
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != SLIDE_MESSAGE_TYPE {
                    continue;
                }
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(Error::InvalidSource)?;
                if info.type_ != message.type_
                    || usize::try_from(info.length).ok() != Some(message.data.len())
                {
                    return Err(Error::InvalidSource);
                }
                budget.work(1)?;
                let view = parse_wire(&message.data, limits, budget)?;
                for field in view
                    .fields()
                    .filter(|field| field.number() == SLIDE_OWNED_DRAWABLES_FIELD)
                {
                    if field.wire_type() != 2 {
                        return Err(Error::InvalidSource);
                    }
                    field
                        .validate_canonical_framing()
                        .map_err(|_| Error::InvalidSource)?;
                    let identifier = reference_identifier(field.payload(), limits, budget)?;
                    budget.references(1)?;
                    if identifier != movie_identifier {
                        continue;
                    }
                    total_references = total_references
                        .checked_add(1)
                        .ok_or(Error::InvalidSource)?;
                    if object_identifier == Some(slide_identifier) {
                        selected_references = selected_references
                            .checked_add(1)
                            .ok_or(Error::InvalidSource)?;
                    }
                    if total_references > 1 {
                        return Err(Error::Refused(Refusal::UnsupportedDependency));
                    }
                }
            }
        }
    }

    if object_matches != 1
        || !object_component_matches
        || total_references != 1
        || selected_references != 1
        || metadata_census.invalid
        || metadata_census.selected_references != expected_metadata_references
    {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(())
}

struct MovieInboundMetadataCensus<'a> {
    package: &'a Package,
    movie_identifier: u64,
    slide_identifier: u64,
    slide_message_index: usize,
    selected_references: usize,
    invalid: bool,
}

impl ArchiveReferenceVisitor for MovieInboundMetadataCensus<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.kind != ArchiveReferenceKind::Object
            || occurrence.referenced_identifier != self.movie_identifier
        {
            return Ok(());
        }
        let Some((_, owner)) = self
            .package
            .object_with_component(occurrence.object_identifier)
        else {
            self.invalid = true;
            return Ok(());
        };
        let Some(info) = owner
            .archive_info
            .message_infos
            .get(occurrence.message_index)
        else {
            self.invalid = true;
            return Ok(());
        };
        let ownership_edge = match occurrence.scope {
            ArchiveReferenceScope::Message => {
                matches!(
                    info.type_,
                    SLIDE_MESSAGE_TYPE | CONTAINER_MESSAGE_TYPE | GROUP_MESSAGE_TYPE
                )
            },
            ArchiveReferenceScope::Field { field_index } => {
                let Some(field) = info.field_infos.get(field_index) else {
                    self.invalid = true;
                    return Ok(());
                };
                match info.type_ {
                    SLIDE_MESSAGE_TYPE => field.path.as_slice() == [SLIDE_OWNED_DRAWABLES_FIELD],
                    CONTAINER_MESSAGE_TYPE => field.path.as_slice() == [3],
                    GROUP_MESSAGE_TYPE => field.path.as_slice() == [2],
                    _ => false,
                }
            },
        };
        if !ownership_edge {
            return Ok(());
        }
        if occurrence.object_identifier != self.slide_identifier
            || occurrence.message_index != self.slide_message_index
        {
            self.invalid = true;
            return Ok(());
        }
        let Some(count) = self.selected_references.checked_add(1) else {
            self.invalid = true;
            return Ok(());
        };
        self.selected_references = count;
        Ok(())
    }
}

fn validate_properties(properties: &MediaProperties) -> Result<(), Error> {
    let bytes = properties
        .hyperlink_url()
        .map_or(0, str::len)
        .checked_add(properties.accessibility_description().map_or(0, str::len))
        .ok_or(Error::InvalidSource)?;
    if bytes > MAX_PROPERTY_BYTES {
        return Err(Error::Limit {
            kind: LimitKind::PropertyBytes,
            observed: bytes as u64,
            maximum: MAX_PROPERTY_BYTES as u64,
        });
    }
    Ok(())
}

fn owned_properties(
    snapshot: keynote_media_properties_codec::MoviePropertiesSnapshot<'_>,
    budget: &mut GeometryBudget,
) -> Result<MediaProperties, Error> {
    Ok(MediaProperties::from_parts(
        owned_string(snapshot.hyperlink_url(), budget)?,
        snapshot.locked(),
        snapshot.aspect_ratio_locked(),
        owned_string(snapshot.accessibility_description(), budget)?,
    ))
}

fn clone_properties(
    properties: &MediaProperties,
    budget: &mut GeometryBudget,
) -> Result<MediaProperties, Error> {
    Ok(MediaProperties::from_parts(
        owned_string(properties.hyperlink_url(), budget)?,
        properties.locked(),
        properties.aspect_ratio_locked(),
        owned_string(properties.accessibility_description(), budget)?,
    ))
}

fn shared_properties(
    properties: MediaProperties,
    budget: &mut GeometryBudget,
) -> Result<Arc<MediaProperties>, Error> {
    budget.allocations(1)?;
    budget.retained(size_of::<MediaProperties>())?;
    Ok(Arc::new(properties))
}

fn owned_string(value: Option<&str>, budget: &mut GeometryBudget) -> Result<Option<String>, Error> {
    let Some(value) = value else {
        return Ok(None);
    };
    budget.allocations(1)?;
    budget.retained(value.len())?;
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_| Error::Allocation {
            amount: value.len(),
        })?;
    output.push_str(value);
    Ok(Some(output))
}

fn property_codec_options(
    package: &Package,
    payload: &[u8],
    budget: &GeometryBudget,
) -> Result<keynote_media_properties_codec::DecodeOptions, Error> {
    let limits = budget.residual(package)?;
    let recursion = u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?;
    let output = budget.remaining_output()?.min(limits.max_output_bytes());
    let allocations = budget.remaining_allocations()?;
    let retained = budget.remaining_retained()?;
    let scratch = budget.remaining_scratch()?;
    Ok(keynote_media_properties_codec::DecodeOptions::new(
        limits.max_input_bytes().min(payload.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    )
    .with_max_output_bytes(output)
    .with_max_allocations(allocations)
    .with_max_retained_bytes(retained)
    .with_max_scratch_bytes(scratch))
}

fn budget_property_report(
    budget: &mut GeometryBudget,
    report: keynote_media_properties_codec::DecodeReport,
) -> Result<(), Error> {
    budget.source(report.input_bytes())?;
    budget.allocations(report.allocations())?;
    budget.retained(report.retained_bytes())?;
    budget.scratch(report.scratch_bytes())?;
    budget.fields(report.fields())?;
    budget.work(report.work_bytes())?;
    budget.nesting(report.max_depth() as usize)?;
    Ok(())
}

fn budget_property_requirements(
    budget: &mut GeometryBudget,
    requirements: keynote_media_properties_codec::RewriteExecutionRequirements,
) -> Result<(), Error> {
    budget.output(requirements.output_bytes)?;
    budget.fields(requirements.fields)?;
    budget.work(requirements.work_bytes)?;
    budget.allocations(requirements.allocations)?;
    budget.retained(requirements.retained_bytes)?;
    budget.scratch(requirements.scratch_bytes)?;
    budget.nesting(requirements.max_depth as usize)?;
    Ok(())
}

fn map_properties_codec_error(error: keynote_media_properties_codec::DecodeError) -> Error {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            keynote_media_properties_codec::DecodeLimit::Bytes { observed, maximum } => {
                (LimitKind::WireBytes, observed, maximum)
            },
            keynote_media_properties_codec::DecodeLimit::Fields { observed, maximum } => {
                (LimitKind::WireFields, observed, maximum)
            },
            keynote_media_properties_codec::DecodeLimit::Work { observed, maximum } => {
                (LimitKind::WireWork, observed, maximum)
            },
            keynote_media_properties_codec::DecodeLimit::Nesting { observed, maximum } => {
                (LimitKind::WireNesting, observed as usize, maximum as usize)
            },
            keynote_media_properties_codec::DecodeLimit::Allocations { observed, maximum } => {
                (LimitKind::Allocations, observed, maximum)
            },
            keynote_media_properties_codec::DecodeLimit::Retained { observed, maximum } => {
                (LimitKind::Retained, observed, maximum)
            },
            keynote_media_properties_codec::DecodeLimit::Scratch { observed, maximum } => {
                (LimitKind::Scratch, observed, maximum)
            },
            keynote_media_properties_codec::DecodeLimit::Output { observed, maximum } => {
                (LimitKind::OutputBytes, observed, maximum)
            },
            _ => return Error::InvalidSource,
        };
        return Error::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

fn to_geometry_error(error: Error) -> super::slide_movie_geometry::Error {
    use super::slide_movie_geometry::{
        Error as SlideMovieGeometryError, LimitKind as SlideMovieGeometryLimitKind,
    };
    match error {
        Error::Refused(Refusal::UnsupportedSource) => SlideMovieGeometryError::Refused(
            crate::slide::movie::geometry::Refusal::UnsupportedSource,
        ),
        Error::Refused(Refusal::UnsupportedDependency)
        | Error::Refused(Refusal::WrongMediaKind) => SlideMovieGeometryError::Refused(
            crate::slide::movie::geometry::Refusal::UnsupportedDependency,
        ),
        Error::Refused(Refusal::AmbiguousSelector) => SlideMovieGeometryError::Refused(
            crate::slide::movie::geometry::Refusal::AmbiguousSelector,
        ),
        Error::Refused(Refusal::EmptySlideName) => {
            SlideMovieGeometryError::Refused(crate::slide::movie::geometry::Refusal::EmptySlideName)
        },
        Error::Refused(Refusal::SlideNameNotFound) => SlideMovieGeometryError::Refused(
            crate::slide::movie::geometry::Refusal::SlideNameNotFound,
        ),
        Error::Refused(Refusal::SlidePositionNotFound { position }) => {
            SlideMovieGeometryError::Refused(
                crate::slide::movie::geometry::Refusal::SlidePositionNotFound { position },
            )
        },
        Error::Refused(Refusal::MoviePositionNotFound { position }) => {
            SlideMovieGeometryError::Refused(
                crate::slide::movie::geometry::Refusal::MoviePositionNotFound { position },
            )
        },
        Error::InvalidSource => SlideMovieGeometryError::InvalidSource,
        Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideMovieGeometryError::Limit {
            kind: match kind {
                LimitKind::InputBytes => SlideMovieGeometryLimitKind::InputBytes,
                LimitKind::OutputBytes => SlideMovieGeometryLimitKind::OutputBytes,
                LimitKind::WireBytes => SlideMovieGeometryLimitKind::WireBytes,
                LimitKind::Entries => SlideMovieGeometryLimitKind::Entries,
                LimitKind::EntryBytes => SlideMovieGeometryLimitKind::EntryBytes,
                LimitKind::TotalBytes => SlideMovieGeometryLimitKind::TotalBytes,
                LimitKind::Slides => SlideMovieGeometryLimitKind::Slides,
                LimitKind::References => SlideMovieGeometryLimitKind::References,
                LimitKind::WireFields => SlideMovieGeometryLimitKind::WireFields,
                LimitKind::WireNesting => SlideMovieGeometryLimitKind::WireNesting,
                LimitKind::WireWork => SlideMovieGeometryLimitKind::WireWork,
                LimitKind::PropertyBytes => SlideMovieGeometryLimitKind::GeometryBytes,
                LimitKind::Allocations => SlideMovieGeometryLimitKind::Allocations,
                LimitKind::Retained => SlideMovieGeometryLimitKind::Retained,
                LimitKind::Scratch => SlideMovieGeometryLimitKind::Scratch,
                LimitKind::Components => SlideMovieGeometryLimitKind::Components,
            },
            observed,
            maximum,
        },
        Error::Allocation { amount } => SlideMovieGeometryError::Allocation { amount },
        Error::Verification => SlideMovieGeometryError::Verification,
        Error::PatchConflict => SlideMovieGeometryError::PatchConflict,
    }
}

impl From<super::slide_movie_geometry::Error> for Error {
    fn from(error: super::slide_movie_geometry::Error) -> Self {
        use super::slide_movie_geometry::{
            Error as SlideMovieGeometryError, LimitKind as SlideMovieGeometryLimitKind,
        };
        match error {
            SlideMovieGeometryError::Refused(
                crate::slide::movie::geometry::Refusal::UnsupportedSource,
            ) => Self::Refused(Refusal::UnsupportedSource),
            SlideMovieGeometryError::Refused(
                crate::slide::movie::geometry::Refusal::UnsupportedDependency,
            ) => Self::Refused(Refusal::UnsupportedDependency),
            SlideMovieGeometryError::Refused(
                crate::slide::movie::geometry::Refusal::AmbiguousSelector,
            ) => Self::Refused(Refusal::AmbiguousSelector),
            SlideMovieGeometryError::Refused(
                crate::slide::movie::geometry::Refusal::EmptySlideName,
            ) => Self::Refused(Refusal::EmptySlideName),
            SlideMovieGeometryError::Refused(
                crate::slide::movie::geometry::Refusal::SlideNameNotFound,
            ) => Self::Refused(Refusal::SlideNameNotFound),
            SlideMovieGeometryError::Refused(
                crate::slide::movie::geometry::Refusal::SlidePositionNotFound { position },
            ) => Self::Refused(Refusal::SlidePositionNotFound { position }),
            SlideMovieGeometryError::Refused(
                crate::slide::movie::geometry::Refusal::MoviePositionNotFound { position },
            ) => Self::Refused(Refusal::MoviePositionNotFound { position }),
            SlideMovieGeometryError::InvalidSource => Self::InvalidSource,
            SlideMovieGeometryError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    SlideMovieGeometryLimitKind::InputBytes => LimitKind::InputBytes,
                    SlideMovieGeometryLimitKind::OutputBytes => LimitKind::OutputBytes,
                    SlideMovieGeometryLimitKind::WireBytes => LimitKind::WireBytes,
                    SlideMovieGeometryLimitKind::EntryBytes => LimitKind::EntryBytes,
                    SlideMovieGeometryLimitKind::Entries => LimitKind::Entries,
                    SlideMovieGeometryLimitKind::TotalBytes => LimitKind::TotalBytes,
                    SlideMovieGeometryLimitKind::Slides => LimitKind::Slides,
                    SlideMovieGeometryLimitKind::References => LimitKind::References,
                    SlideMovieGeometryLimitKind::WireFields => LimitKind::WireFields,
                    SlideMovieGeometryLimitKind::WireNesting => LimitKind::WireNesting,
                    SlideMovieGeometryLimitKind::WireWork => LimitKind::WireWork,
                    SlideMovieGeometryLimitKind::GeometryBytes => LimitKind::PropertyBytes,
                    SlideMovieGeometryLimitKind::Allocations => LimitKind::Allocations,
                    SlideMovieGeometryLimitKind::Retained => LimitKind::Retained,
                    SlideMovieGeometryLimitKind::Scratch => LimitKind::Scratch,
                    SlideMovieGeometryLimitKind::Components => LimitKind::Components,
                },
                observed,
                maximum,
            },
            SlideMovieGeometryError::Allocation { amount } => Self::Allocation { amount },
            SlideMovieGeometryError::Verification => Self::Verification,
            SlideMovieGeometryError::PatchConflict => Self::PatchConflict,
            SlideMovieGeometryError::Refused(crate::slide::movie::geometry::Refusal::Locked) => {
                Self::Refused(Refusal::UnsupportedDependency)
            },
        }
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::limit_from_semantic(kind, observed, maximum),
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::limit_from_payload(kind, observed, maximum),
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

impl Error {
    fn limit_from_semantic(kind: SemanticLimitKind, observed: usize, maximum: usize) -> Self {
        Self::Limit {
            kind: match kind {
                SemanticLimitKind::Slides => LimitKind::Slides,
                SemanticLimitKind::References => LimitKind::References,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        }
    }

    fn limit_from_payload(kind: super::PayloadLimitKind, observed: usize, maximum: usize) -> Self {
        Self::Limit {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        }
    }
}
