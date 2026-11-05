//! Selector-first duplication and removal of a Keynote slide movie or audio
//! drawable.
//!
//! The owner in this module is intentionally the only place where the
//! private native graph is turned into a transaction.  Public callers select
//! a slide and a source-order media position; object identifiers, component
//! names, UUIDs, archive headers, and data records stay inside this module or
//! one of its bounded child adapters.  Every changed package is built from
//! the exact source ZIP and reopened before the commit is returned.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::too_many_lines,
    clippy::too_many_arguments,
    reason = "The lifecycle transaction keeps source admission, graph closure, and publication together."
)]

use std::{fmt, mem::size_of, sync::Arc};

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{Archive, Limits as ArchiveLimits, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    keynote_media_lifecycle_codec as lifecycle_codec, package_metadata_codec as identity_codec,
    package_metadata_media_codec as media_codec,
};
use sha1::{Digest, Sha1};
use thiserror::Error as ThisError;

use super::snappy_failure::SnappyFailure;
use super::{Package, PhysicalSource, SLIDE_MESSAGE_TYPE, SemanticPath};
use crate::{MovieKind, MovieSelector, SlideSelector};

pub(in crate::package) mod budget;
mod clone_payload;
mod comment_clone;
pub(in crate::package) mod comment_graph;
mod comment_removal;
mod graph;
pub(super) mod graph_caption_witness;
mod identifier_watermark;
mod metadata;
mod node_cache;

pub(in crate::package) use budget::LifecycleBudget;
use graph::MediaGraphSelection;

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const BUILD_MESSAGE_TYPE: u32 = 8;
const BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const DATA_METADATA_MAP_MESSAGE_TYPE: u32 = 11_015;
const MOVIE_DATA_FIELD: u32 = 14;
const POSTER_IMAGE_DATA_FIELD: u32 = 15;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_BUILDS_FIELD: u32 = 2;
const SLIDE_BUILD_CHUNKS_FIELD: u32 = 43;
const DATA_METADATA_MAP_FIELD: u32 = 10;
const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const DATA_PREFIX: &str = "Data/";
const SHA1_BYTES: usize = 20;

/// Finite resource categories returned by one lifecycle operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    Slides,
    References,
    MediaBytes,
    WireFields,
    WireNesting,
    WireWork,
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::MediaBytes => "media bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
        };
        formatter.write_str(name)
    }
}

/// Errors raised while selecting, rewriting, or verifying a slide-media
/// lifecycle transaction.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    #[error("the Keynote package source is malformed or inconsistent")]
    InvalidSource,
    #[error("lifecycle resource limit {kind} exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("allocation of {amount} bytes was refused")]
    Allocation { amount: usize },
    #[error("candidate verification failed")]
    Verification,
    #[error("the lifecycle patch does not apply to this exact source")]
    PatchConflict,
    #[error("the package could not be read")]
    Read,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    #[error("the source is not an exact physical Keynote package")]
    UnsupportedSource,
    #[error("Keynote slide name must not be empty")]
    EmptySlideName,
    #[error("Keynote slide name was not found")]
    SlideNameNotFound,
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    #[error("Keynote slide position was not found: {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("Keynote media position was not found: {position:?}")]
    MoviePositionNotFound { position: Position },
    #[error(
        "selected Keynote media kind does not match the requested operation (expected {expected:?}, actual {actual:?})"
    )]
    KindMismatch {
        expected: MovieKind,
        actual: MovieKind,
    },
    #[error("Keynote media lifecycle edit cannot safely own this selected comment graph")]
    UnsupportedComment,
    #[error("an audio drawable has no poster image")]
    AudioPoster,
    /// The edit already stages its one duplication or removal.
    #[error("the media lifecycle edit already stages an operation")]
    OperationAlreadyStaged,
    /// The edit was committed without a staged operation.
    #[error("the media lifecycle edit stages no operation")]
    NoStagedOperation,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. }
            | Self::UnsupportedComment { .. }
            | Self::AudioPoster { .. } => RefusalKind::UnsupportedContent,
            Self::EmptySlideName { .. } | Self::AmbiguousSelector { .. } => {
                RefusalKind::InvalidValue
            },
            Self::SlideNameNotFound { .. }
            | Self::SlidePositionNotFound { .. }
            | Self::MoviePositionNotFound { .. } => RefusalKind::NotFound,
            Self::KindMismatch { .. } => RefusalKind::UnsupportedTarget,
            Self::OperationAlreadyStaged { .. } | Self::NoStagedOperation { .. } => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// The 128-bit identity used by native build chunks, kept private to the
/// package adapter.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(in crate::package) struct SuperUuid {
    pub(in crate::package) lower: u64,
    pub(in crate::package) upper: u64,
}

impl From<identity_codec::UuidBits> for SuperUuid {
    fn from(value: identity_codec::UuidBits) -> Self {
        Self {
            lower: value.lower(),
            upper: value.upper(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LifecycleAction {
    Duplicate,
    Remove,
}

/// A source-order identity retained by a patch without exposing native IDs.
#[derive(Clone, PartialEq, Eq)]
struct SelectionFingerprint {
    slide_position: Position,
    movie_position: Position,
    kind: MovieKind,
    source_media_count: usize,
    target_media_count: usize,
}

impl fmt::Debug for SelectionFingerprint {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SelectionFingerprint")
            .field("slide_position", &self.slide_position)
            .field("movie_position", &self.movie_position)
            .field("kind", &self.kind)
            .field("source_media_count", &self.source_media_count)
            .field("target_media_count", &self.target_media_count)
            .finish()
    }
}

/// Exact source and target artifacts produced by a movie/audio lifecycle edit.
#[must_use]
#[derive(Clone)]
pub struct Patch {
    artifacts: ExactArtifacts,
    action: LifecycleAction,
    selection: SelectionFingerprint,
    source_selection: Arc<MediaGraphSelection>,
    created_objects: usize,
    removed_objects: usize,
    removed_data: usize,
    touched_members: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("action", &self.action)
            .field("selection", &self.selection)
            .field("created_objects", &self.created_objects)
            .field("removed_objects", &self.removed_objects)
            .field("removed_data", &self.removed_data)
            .field("touched_members", &self.touched_members)
            .field("artifacts", &self.artifacts)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the compact source artifact fingerprint used for diagnostics.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the compact target artifact fingerprint used for diagnostics.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return the source-order slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the source-order media position.
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }

    /// Return the selected media's semantic kind.
    #[must_use]
    pub const fn kind(&self) -> MovieKind {
        self.selection.kind
    }

    /// Return the number of media drawables before the edit.
    #[must_use]
    pub const fn source_media_count(&self) -> usize {
        self.selection.source_media_count
    }

    /// Return the number of media drawables after the edit.
    #[must_use]
    pub const fn target_media_count(&self) -> usize {
        self.selection.target_media_count
    }

    /// Return how many IWA objects were appended by duplication.
    #[must_use]
    pub const fn created_objects(&self) -> usize {
        self.created_objects
    }

    /// Return how many IWA objects were removed.
    #[must_use]
    pub const fn removed_objects(&self) -> usize {
        self.removed_objects
    }

    /// Return how many DataInfo records and physical data members were removed.
    #[must_use]
    pub const fn removed_data(&self) -> usize {
        self.removed_data
    }

    /// Return how many ZIP members were rewritten or deleted.
    #[must_use]
    pub const fn touched_members(&self) -> usize {
        self.touched_members
    }

    /// Return whether source and target are byte-identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse transaction.
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            action: match self.action {
                LifecycleAction::Duplicate => LifecycleAction::Remove,
                LifecycleAction::Remove => LifecycleAction::Duplicate,
            },
            selection: SelectionFingerprint {
                source_media_count: self.selection.target_media_count,
                target_media_count: self.selection.source_media_count,
                ..self.selection.clone()
            },
            source_selection: Arc::clone(&self.source_selection),
            created_objects: self.removed_objects,
            removed_objects: self.created_objects,
            removed_data: self.removed_data,
            touched_members: self.touched_members,
        }
    }
}

/// Resource and topology evidence for a committed lifecycle edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_members: usize,
    created_objects: usize,
    removed_objects: usize,
    removed_data: usize,
    source_media_count: usize,
    target_media_count: usize,
}

impl Diagnostics {
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    #[must_use]
    pub const fn touched_members(self) -> usize {
        self.touched_members
    }
    #[must_use]
    pub const fn created_objects(self) -> usize {
        self.created_objects
    }
    #[must_use]
    pub const fn removed_objects(self) -> usize {
        self.removed_objects
    }
    #[must_use]
    pub const fn removed_data(self) -> usize {
        self.removed_data
    }
    #[must_use]
    pub const fn source_media_count(self) -> usize {
        self.source_media_count
    }
    #[must_use]
    pub const fn target_media_count(self) -> usize {
        self.target_media_count
    }
}

/// The verified result of one immutable lifecycle transaction.
#[must_use]
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
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// One staged movie or audio duplication or removal.
#[derive(Debug, Clone, Copy)]
struct Staged<'a> {
    slide: SlideSelector<'a>,
    media: MovieSelector,
    action: LifecycleAction,
    expected_kind: Option<MovieKind>,
}

/// One immutable edit that stages exactly one duplication or removal of a
/// source-order movie or audio drawable on a slide.
///
/// Start it with [`Package::edit_slide_media_lifecycle`]; the source package
/// is never changed, and dropping the edit discards the staged operation.
#[derive(Debug)]
pub struct Edit<'a> {
    source: &'a Package,
    staged: Option<Staged<'a>>,
}

impl<'a> Edit<'a> {
    fn stage(
        mut self,
        slide: SlideSelector<'a>,
        media: MovieSelector,
        action: LifecycleAction,
        expected_kind: Option<MovieKind>,
    ) -> Result<Self, Error> {
        if self.staged.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        self.staged = Some(Staged {
            slide,
            media,
            action,
            expected_kind,
        });
        Ok(self)
    }

    /// Stage the duplication of one source-order movie or audio drawable on
    /// a slide.
    ///
    /// Direct comments and their replies receive independent storage while
    /// preserving their text, storage UUIDs, and shared authors.
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn duplicate(
        self,
        slide: impl Into<SlideSelector<'a>>,
        media: impl Into<MovieSelector>,
    ) -> Result<Self, Error> {
        self.stage(slide.into(), media.into(), LifecycleAction::Duplicate, None)
    }

    /// Stage the removal of one source-order movie or audio drawable from a
    /// slide.
    ///
    /// Shared comment storage survives until its last drawable owner is
    /// removed. Shared authors remain available after their component
    /// dependency is released.
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn remove(
        self,
        slide: impl Into<SlideSelector<'a>>,
        media: impl Into<MovieSelector>,
    ) -> Result<Self, Error> {
        self.stage(slide.into(), media.into(), LifecycleAction::Remove, None)
    }

    /// Stage the duplication of a file-backed movie; the commit refuses any
    /// other media kind with [`Refusal::KindMismatch`]. The selector remains
    /// source-order and therefore stays stable when native object identifiers
    /// change. Direct comments follow the cloning policy of
    /// [`Self::duplicate`].
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn duplicate_movie(
        self,
        slide: impl Into<SlideSelector<'a>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<Self, Error> {
        self.stage(
            slide.into(),
            movie.into(),
            LifecycleAction::Duplicate,
            Some(MovieKind::File),
        )
    }

    /// Stage the duplication of an audio drawable, including supported
    /// direct comment threads; the commit refuses any other media kind with
    /// [`Refusal::KindMismatch`].
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn duplicate_audio(
        self,
        slide: impl Into<SlideSelector<'a>>,
        audio: impl Into<MovieSelector>,
    ) -> Result<Self, Error> {
        self.stage(
            slide.into(),
            audio.into(),
            LifecycleAction::Duplicate,
            Some(MovieKind::Audio),
        )
    }

    /// Stage the removal of a file-backed movie; the commit refuses any other
    /// media kind with [`Refusal::KindMismatch`].
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn remove_movie(
        self,
        slide: impl Into<SlideSelector<'a>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<Self, Error> {
        self.stage(
            slide.into(),
            movie.into(),
            LifecycleAction::Remove,
            Some(MovieKind::File),
        )
    }

    /// Stage the removal of an audio drawable; the commit refuses any other
    /// media kind with [`Refusal::KindMismatch`].
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] after one operation.
    pub fn remove_audio(
        self,
        slide: impl Into<SlideSelector<'a>>,
        audio: impl Into<MovieSelector>,
    ) -> Result<Self, Error> {
        self.stage(
            slide.into(),
            audio.into(),
            LifecycleAction::Remove,
            Some(MovieKind::Audio),
        )
    }

    /// Select, publish and reopen the staged operation as one exact-source
    /// transaction.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit; selector, kind,
    /// comment, source, resource and verification errors. Failure never
    /// publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let staged = self
            .staged
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        run_lifecycle(
            self.source,
            staged.slide,
            staged.media,
            staged.action,
            staged.expected_kind,
        )
    }
}

impl Package {
    /// Start an immutable single-operation movie or audio lifecycle edit.
    ///
    /// ```no_run
    /// use litchi_keynote::{MovieSelector, Package, SlideSelector};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("talk.key")?;
    /// let commit = package
    ///     .edit_slide_media_lifecycle()
    ///     .duplicate(SlideSelector::index(0), MovieSelector::index(0))?
    ///     .commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_slide_media_lifecycle(&commit.patch().inverse())?;
    /// assert!(restored.diagnostics().changed());
    /// # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub const fn edit_slide_media_lifecycle(&self) -> Edit<'_> {
        Edit {
            source: self,
            staged: None,
        }
    }

    /// Apply a patch only when this package is the exact retained source.
    ///
    /// Applying a patch reopens and validates its target artifact.  No native
    /// identifier from the patch is trusted as a source selector.
    pub fn apply_slide_media_lifecycle(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let mut budget = LifecycleBudget::for_package(self)?;
        budget.charge_output(patch.artifacts.target().len())?;
        budget.charge_allocations(patch.artifacts.target().len())?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(|_| Error::PatchConflict)?;
        candidate.validate().map_err(|_| Error::PatchConflict)?;
        verify_candidate_selection(&candidate, patch, &mut budget)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: diagnostics_for_patch(patch),
        })
    }
}

fn run_lifecycle(
    source: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    action: LifecycleAction,
    expected_kind: Option<MovieKind>,
) -> Result<Commit, Error> {
    let catalog = physical_catalog(source)?;
    let mut budget = LifecycleBudget::for_package(source)?;
    budget.charge_entries(catalog.package().len())?;
    let wire_limits = source.semantic_wire_limits().map_err(|_| Error::Read)?;
    let selection = graph::select_media(
        source,
        slide_selector,
        movie_selector,
        wire_limits,
        &mut budget,
    )?;
    if let Some(expected) = expected_kind
        && selection.kind != expected
    {
        return Err(Error::Refused(Refusal::KindMismatch {
            expected,
            actual: selection.kind,
        }));
    }
    budget.charge_entries(1)?;
    budget.charge_references(
        selection
            .private_object_ids
            .len()
            .checked_add(selection.build_ids.len())
            .and_then(|value| value.checked_add(selection.chunk_ids.len()))
            .ok_or(Error::InvalidSource)?,
    )?;
    let source_media_count = count_slide_media(source, &selection, wire_limits, &mut budget)?;
    let target_media_count = match action {
        LifecycleAction::Duplicate => source_media_count
            .checked_add(1)
            .ok_or(Error::InvalidSource)?,
        LifecycleAction::Remove => source_media_count
            .checked_sub(1)
            .ok_or(Error::InvalidSource)?,
    };
    let (candidate, patch) = rewrite_lifecycle(
        source,
        catalog,
        selection,
        action,
        source_media_count,
        target_media_count,
        wire_limits,
        &mut budget,
    )?;
    let diagnostics = diagnostics_for_patch(&patch);
    Ok(Commit {
        package: candidate,
        patch,
        diagnostics,
    })
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(catalog) if catalog.source_is_exact() => Ok(catalog),
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => {
            Err(Error::Refused(Refusal::UnsupportedSource))
        },
    }
}

fn diagnostics_for_patch(patch: &Patch) -> Diagnostics {
    Diagnostics {
        changed: !patch.is_noop(),
        touched_members: patch.touched_members,
        created_objects: patch.created_objects,
        removed_objects: patch.removed_objects,
        removed_data: patch.removed_data,
        source_media_count: patch.selection.source_media_count,
        target_media_count: patch.selection.target_media_count,
    }
}

fn count_slide_media(
    package: &Package,
    selection: &MediaGraphSelection,
    limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<usize, Error> {
    let object = package
        .object(selection.slide_identifier)
        .ok_or(Error::InvalidSource)?;
    let payload = unique_payload(&object.messages, SLIDE_MESSAGE_TYPE)?;
    budget.charge_wire_work(payload.len().max(1))?;
    budget.charge_allocations(payload.len())?;
    let view = WireView::parse(payload, limits).map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_fields(view.len())?;
    let mut count = 0usize;
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
        let identifier = reference_identifier(field.payload(), limits)?;
        budget.charge_references(1)?;
        let Some(drawable) = package.object(identifier) else {
            return Err(Error::InvalidSource);
        };
        let Some(message) = drawable
            .messages
            .iter()
            .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        else {
            continue;
        };
        let (_info, _) = decode_movie_info(
            &message.data,
            limits,
            SemanticPath::SlideDrawable {
                slide: selection.slide_position.get(),
                index: count,
            },
        )
        .map_err(|_| Error::InvalidSource)?;
        budget.charge_wire_work(message.data.len().max(1))?;
        // `MovieSelector` addresses the complete source-ordered MovieArchive
        // projection, including native placeholders and live-video drawables.
        // The lifecycle graph still rejects those kinds as edit targets, but
        // omitting them here would shift every later selector and would make
        // candidate verification disagree with the original selection.
        count = count.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    Ok(count)
}

fn unique_payload(messages: &[RawMessage], message_type: u32) -> Result<&[u8], Error> {
    let mut payload = None;
    for message in messages
        .iter()
        .filter(|message| message.type_ == message_type)
    {
        if payload.replace(message.data.as_slice()).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    payload.ok_or(Error::InvalidSource)
}

/// Descendant graph helpers use this narrow package-level adapter rather than
/// importing the parent package's private reader directly.  Keeping the
/// result typed as the package reader error also prevents raw IDs or reader
/// internals from crossing the lifecycle boundary.
fn decode_movie_info(
    payload: &[u8],
    limits: WireLimits,
    path: SemanticPath,
) -> super::ReadResult<(super::MovieInfo, usize)> {
    super::decode_movie_info(payload, limits, path)
}

fn reference_identifier(payload: &[u8], limits: WireLimits) -> Result<u64, Error> {
    let view = WireView::parse(payload, limits).map_err(|_| Error::InvalidSource)?;
    let mut identifier = None;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        if field.number() != 1 {
            continue;
        }
        if field.wire_type() != 0 || identifier.is_some() {
            return Err(Error::InvalidSource);
        }
        let (value, width) = litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
            .map_err(|_| Error::InvalidSource)?;
        if width != field.payload().len()
            || litchi_iwa_common::varint::encoded_len(value) != width
            || value == 0
        {
            return Err(Error::InvalidSource);
        }
        identifier = Some(value);
    }
    identifier.ok_or(Error::InvalidSource)
}

fn load_metadata(
    package: &Package,
    catalog: &SourceCatalog,
    archive_limits: ArchiveLimits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    budget: &mut LifecycleBudget,
) -> Result<MetadataArchive, Error> {
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == METADATA_COMPONENT)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    // The source catalog already owns the exact parsed Metadata component.
    // Its encoded-length preflight is allocation-free and gives us a tight
    // decompressed bound before Snappy can grow its output buffer.
    let source_component = package
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == METADATA_COMPONENT)
        .ok_or(Error::InvalidSource)?;
    let expected_stream_len = source_component
        .archive()
        .encoded_len(archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    if expected_stream_len == 0 || expected_stream_len > snappy_limits.max_decompressed_stream() {
        return Err(Error::InvalidSource);
    }
    let chunk_limit = expected_stream_len.min(SnappyStream::MAX_UNCOMPRESSED_CHUNK);
    let tight_snappy_limits = litchi_iwa_core::SnappyLimits::new(chunk_limit, expected_stream_len)
        .and_then(|limits| {
            limits.with_input_limits(
                snappy_limits.max_compressed_chunk(),
                snappy_limits.max_compressed_stream(),
                snappy_limits.max_frames(),
            )
        })
        .map_err(|_| Error::InvalidSource)?;
    let source_object_count = source_component.archive().objects.len();
    let source_message_count =
        source_component
            .archive()
            .objects
            .iter()
            .try_fold(0usize, |total, object| {
                total
                    .checked_add(object.messages.len())
                    .ok_or(Error::InvalidSource)
            })?;
    let parse_scratch = source_object_count
        .checked_mul(size_of::<litchi_iwa_core::ArchiveObject>())
        .and_then(|bytes| {
            source_message_count
                .checked_mul(size_of::<RawMessage>())
                .and_then(|message_bytes| bytes.checked_add(message_bytes))
        })
        // Snappy's decoded buffer remains live while Archive parsing retains
        // copied message payloads, so reserve both source-width copies in the
        // shared ledger before either allocation begins.
        .and_then(|bytes| {
            expected_stream_len
                .checked_mul(2)
                .and_then(|decoded_bytes| bytes.checked_add(decoded_bytes))
        })
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(parse_scratch)?;
    let stream = SnappyStream::decompress(entry.data(), tight_snappy_limits)
        .map_err(|_| Error::InvalidSource)?;
    if stream.as_bytes().len() != expected_stream_len {
        return Err(Error::InvalidSource);
    }
    let archive =
        Archive::parse(stream.as_bytes(), archive_limits).map_err(|_| Error::InvalidSource)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes(), archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    budget.charge_entries(archive.objects.len())?;
    for object in &archive.objects {
        budget.charge_wire_fields(object.messages.len())?;
        let work = object.messages.iter().try_fold(0usize, |total, message| {
            total
                .checked_add(message.data.len().max(1))
                .ok_or(Error::InvalidSource)
        })?;
        budget.charge_wire_work(work)?;
    }
    let mut selected = None;
    for (object_index, object) in archive.objects.iter().enumerate() {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                continue;
            }
            if selected
                .replace((object_index, message_index, object.archive_info.identifier))
                .is_some()
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    let (object_index, message_index, object_identifier) = selected.ok_or(Error::InvalidSource)?;
    let payload_source = archive
        .objects
        .get(object_index)
        .and_then(|object| object.messages.get(message_index))
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(payload_source.len())?;
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(payload_source.len())
        .map_err(|_| Error::Allocation {
            amount: payload_source.len(),
        })?;
    payload.extend_from_slice(payload_source);
    let _ = package;
    Ok(MetadataArchive {
        archive,
        object_identifier: object_identifier.ok_or(Error::InvalidSource)?,
        message_index,
        payload,
    })
}

fn metadata_options(
    package: &Package,
    payload_length: usize,
    budget: &mut LifecycleBudget,
) -> Result<(identity_codec::RewriteOptions, media_codec::DecodeOptions), Error> {
    let semantic = package.semantic_limits();
    let wire = package
        .semantic_wire_limits()
        .map_err(|_| Error::InvalidSource)?;
    let doubled = payload_length
        .checked_mul(2)
        .and_then(|value| value.checked_add(4096))
        .ok_or(Error::InvalidSource)?;
    let fields = payload_length
        .checked_mul(16)
        .and_then(|value| value.checked_add(64))
        .ok_or(Error::InvalidSource)?;
    let work = payload_length
        .checked_mul(128)
        .and_then(|value| value.checked_add(4096))
        .ok_or(Error::InvalidSource)?;
    let depth = u32::try_from(wire.max_nesting()).map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_work(work)?;
    let identity = identity_codec::RewriteOptions::new(
        payload_length.max(1),
        doubled,
        fields,
        work,
        depth,
        semantic.max_objects(),
        semantic.max_references(),
        semantic.max_objects(),
    );
    let media = media_codec::DecodeOptions::new(
        payload_length.max(1),
        fields,
        work,
        semantic.max_objects(),
        semantic.max_objects(),
        semantic.max_references(),
        SHA1_BYTES,
        4096,
        depth,
    )
    .with_max_output_bytes(doubled);
    Ok((identity, media))
}

fn metadata_snapshot<'source>(
    payload: &'source [u8],
    identity_options: &identity_codec::RewriteOptions,
    media_options: &media_codec::DecodeOptions,
    map: Option<media_codec::DataMetadataMapSource<'source>>,
    budget: &mut LifecycleBudget,
) -> Result<metadata::MetadataSnapshot<'source>, Error> {
    let mut callback_error = None;
    let mut charge = |amount: usize| match budget.charge_allocations(amount) {
        Ok(()) => Ok(()),
        Err(error) => {
            callback_error = Some(error);
            Err(metadata::MetadataError::Allocation { amount })
        },
    };
    let snapshot = metadata::MetadataSnapshot::inspect(
        payload,
        *identity_options,
        *media_options,
        map,
        &mut charge,
    )
    .map_err(|error| {
        callback_error
            .take()
            .unwrap_or_else(|| map_metadata_error(error))
    })?;
    let resources = snapshot.resources();
    charge_identity_report_from_codec(resources.identity_report(), budget)?;
    charge_media_decode_report(resources.media_report(), budget)?;
    budget.charge_allocations(resources.scratch_bytes())?;
    Ok(snapshot)
}

fn map_witness<'source>(
    package: &'source Package,
    payload: &'source [u8],
    options: media_codec::DecodeOptions,
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<Option<media_codec::DataMetadataMapSource<'source>>, Error> {
    let view = WireView::parse(payload, wire_limits).map_err(|_| Error::InvalidSource)?;
    let mut map_identifier = None;
    for field in view
        .fields()
        .filter(|field| field.number() == DATA_METADATA_MAP_FIELD)
    {
        if field.wire_type() != 2 || map_identifier.is_some() {
            return Err(Error::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        map_identifier = Some(reference_identifier(field.payload(), wire_limits)?);
    }
    let Some(identifier) = map_identifier else {
        return Ok(None);
    };
    let Some((_component, object)) = package.object_with_component(identifier) else {
        return Err(Error::InvalidSource);
    };
    let mut map_payload = None;
    for message in object
        .messages
        .iter()
        .filter(|message| message.type_ == DATA_METADATA_MAP_MESSAGE_TYPE)
    {
        if map_payload.replace(message.data.as_slice()).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let map_payload = map_payload.ok_or(Error::InvalidSource)?;
    budget.charge_wire_work(map_payload.len())?;
    let source = media_codec::DataMetadataMapSource::from_source(identifier, map_payload, options)
        .map_err(|_| Error::InvalidSource)?;
    budget.charge_wire_fields(source.fields())?;
    budget.charge_wire_work(source.work_bytes())?;
    budget.charge_nesting(source.max_depth() as usize)?;
    budget.charge_references(source.entries())?;
    budget.charge_allocations(source.scratch_bytes())?;
    for _ in 1..source.allocations() {
        budget.charge_allocations(0)?;
    }
    Ok(Some(source))
}

fn map_metadata_error(error: metadata::MetadataError) -> Error {
    match error {
        metadata::MetadataError::Allocation { amount } => Error::Allocation { amount },
        metadata::MetadataError::Identity(error) => {
            if let Some(amount) = error.allocation_request() {
                Error::Allocation { amount }
            } else if let Some(limit) = error.resource_limit() {
                map_identity_limit(limit)
            } else {
                Error::InvalidSource
            }
        },
        metadata::MetadataError::MediaDecode(error) => error
            .resource_limit()
            .map_or(Error::InvalidSource, map_media_limit),
        metadata::MetadataError::MediaRewrite(error) => {
            if let Some(amount) = error.allocation_request() {
                Error::Allocation { amount }
            } else {
                error
                    .resource_limit()
                    .map_or(Error::InvalidSource, map_media_limit)
            }
        },
        metadata::MetadataError::Missing
        | metadata::MetadataError::Ambiguous
        | metadata::MetadataError::Invalid => Error::InvalidSource,
    }
}

fn map_identity_limit(limit: identity_codec::RewriteLimit) -> Error {
    use LimitKind as K;
    use identity_codec::RewriteLimit as L;
    let (kind, observed, maximum) = match limit {
        L::InputBytes { observed, maximum } => (K::InputBytes, observed, maximum),
        L::OutputBytes { observed, maximum } => (K::OutputBytes, observed, maximum),
        L::Fields { observed, maximum } => (K::WireFields, observed, maximum),
        L::Work { observed, maximum } => (K::WireWork, observed, maximum),
        L::Components { observed, maximum } => (K::Entries, observed, maximum),
        L::References { observed, maximum } | L::Additions { observed, maximum } => {
            (K::References, observed, maximum)
        },
        L::Nesting { observed, maximum } => {
            return Error::Limit {
                kind: K::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        _ => return Error::InvalidSource,
    };
    Error::Limit {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn map_media_limit(limit: media_codec::DecodeLimit) -> Error {
    use LimitKind as K;
    use media_codec::DecodeLimit as L;
    let (kind, observed, maximum) = match limit {
        L::Bytes { observed, maximum } => (K::InputBytes, observed, maximum),
        L::OutputBytes { observed, maximum } => (K::OutputBytes, observed, maximum),
        L::Fields { observed, maximum } => (K::WireFields, observed, maximum),
        L::Work { observed, maximum } => (K::WireWork, observed, maximum),
        L::Components { observed, maximum } => (K::Entries, observed, maximum),
        L::DataRecords { observed, maximum } | L::Owners { observed, maximum } => {
            (K::References, observed, maximum)
        },
        L::DigestBytes { observed, maximum } | L::NameBytes { observed, maximum } => {
            (K::EntryBytes, observed, maximum)
        },
        L::Nesting { observed, maximum } => {
            return Error::Limit {
                kind: K::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        _ => return Error::InvalidSource,
    };
    Error::Limit {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn charge_identity_snapshot_report(
    report: metadata::IdentityRewriteReport,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.scratch_bytes())?;
    budget.charge_references(report.references_scanned())?;
    Ok(())
}

fn charge_identity_report_from_codec(
    report: identity_codec::RewriteReport,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.scratch_bytes())?;
    budget.charge_references(report.references_scanned())?;
    Ok(())
}

fn charge_media_decode_report(
    report: media_codec::DecodeReport,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_references(report.data_references().saturating_add(report.owners()))?;
    Ok(())
}

fn charge_media_report(
    report: media_codec::RewriteReport,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_references(
        report
            .owners()
            .saturating_add(report.owner_additions())
            .saturating_add(report.owner_removals()),
    )?;
    budget.charge_allocations(report.scratch_bytes())?;
    Ok(())
}

fn selected_object_ids(
    selection: &MediaGraphSelection,
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, Error> {
    let capacity = 1usize
        .checked_add(selection.private_object_ids.len())
        .and_then(|value| value.checked_add(selection.build_ids.len()))
        .and_then(|value| value.checked_add(selection.chunk_ids.len()))
        .ok_or(Error::InvalidSource)?;
    let mut ids = Vec::new();
    let bytes = capacity
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    ids.try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation { amount: capacity })?;
    ids.push(selection.movie_identifier);
    ids.extend(selection.private_object_ids.iter().copied());
    ids.extend(selection.build_ids.iter().copied());
    ids.extend(selection.chunk_ids.iter().copied());
    ids.sort_unstable();
    ids.dedup();
    Ok(ids)
}

fn allocate_clone_identities(
    package: &Package,
    component_archive: &Archive,
    snapshot: &metadata::MetadataSnapshot<'_>,
    component: metadata::ComponentIdentity<'_>,
    source_ids: &[u64],
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(Vec<(u64, u64)>, Vec<u64>, Vec<SuperUuid>), Error> {
    let mut global_max = snapshot
        .last_identifier()
        .max(snapshot.metadata_identifier_maximum());
    // Object IDs are package-global, so a selected component's local maximum
    // is insufficient.  The census remains linear in the already indexed
    // source components and is charged before any result vectors grow.
    let mut object_count = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            object_count = object_count.checked_add(1).ok_or(Error::InvalidSource)?;
            if let Some(identifier) = object.archive_info.identifier {
                global_max = global_max.max(identifier);
            }
        }
    }
    budget.charge_entries(object_count)?;
    let bytes = source_ids
        .len()
        .checked_mul(
            size_of::<(u64, u64)>()
                .checked_add(size_of::<u64>())
                .and_then(|bytes| bytes.checked_add(size_of::<SuperUuid>()))
                .and_then(|bytes| {
                    bytes.checked_add(size_of::<(u64, SuperUuid, Option<SuperUuid>)>())
                })
                .ok_or(Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    let mut remap = Vec::new();
    let mut new_ids = Vec::new();
    let mut new_uuids = Vec::new();
    remap
        .try_reserve_exact(source_ids.len())
        .map_err(|_| Error::Allocation {
            amount: source_ids.len().saturating_mul(size_of::<(u64, u64)>()),
        })?;
    new_ids
        .try_reserve_exact(source_ids.len())
        .map_err(|_| Error::Allocation {
            amount: source_ids.len().saturating_mul(size_of::<u64>()),
        })?;
    new_uuids
        .try_reserve_exact(source_ids.len())
        .map_err(|_| Error::Allocation {
            amount: source_ids.len().saturating_mul(size_of::<SuperUuid>()),
        })?;
    let mut generated_build_uuids: Vec<(u64, SuperUuid, Option<SuperUuid>)> = Vec::new();
    generated_build_uuids
        .try_reserve_exact(source_ids.len())
        .map_err(|_| Error::Allocation {
            amount: source_ids
                .len()
                .saturating_mul(size_of::<(u64, SuperUuid, Option<SuperUuid>)>()),
        })?;
    for &source_id in source_ids {
        global_max = global_max.checked_add(1).ok_or(Error::InvalidSource)?;
        remap.push((source_id, global_max));
        new_ids.push(global_max);
        let uuid = match snapshot.object_uuid(component, source_id) {
            Ok(uuid) => SuperUuid::from(uuid),
            Err(metadata::MetadataError::Missing) => SuperUuid {
                lower: source_id,
                upper: source_id.rotate_left(17),
            },
            Err(error) => return Err(map_metadata_error(error)),
        };
        let new_uuid = derive_clone_uuid(uuid, global_max);
        budget.charge_wire_work(new_uuids.len())?;
        if snapshot.has_uuid(identity_codec::UuidBits::new(
            new_uuid.lower,
            new_uuid.upper,
        )) || new_uuids.contains(&new_uuid)
        {
            return Err(Error::InvalidSource);
        }
        if selection_contains_build_id(source_id, source_ids, component_archive) {
            if generated_build_uuids
                .iter()
                .any(|(_, existing, _)| *existing == new_uuid)
            {
                return Err(Error::InvalidSource);
            }
            generated_build_uuids.push((global_max, new_uuid, None));
        }
        new_uuids.push(new_uuid);
    }

    // BuildChunk UUIDs identify a build within the playback graph. Both
    // payload locations, and all chunks for that build, must agree. The
    // migration builder uses a separate ObjectUuidMap identity; native files
    // commonly use the same UUID for both domains. Derive clone identities
    // independently for both domains, preserving their source relationship.
    for &chunk_id in source_ids {
        let Some(chunk_index) = source_ids.iter().position(|id| *id == chunk_id) else {
            return Err(Error::InvalidSource);
        };
        let Some(chunk_object) = component_archive.object(chunk_id) else {
            continue;
        };
        let Some(message) = chunk_object
            .messages
            .iter()
            .find(|message| message.type_ == BUILD_CHUNK_MESSAGE_TYPE)
        else {
            continue;
        };
        let options = lifecycle_options(wire_limits, message.data.len())?;
        let (chunk, report) =
            lifecycle_codec::decode_build_chunk_with_report(&message.data, options)
                .map_err(|_| Error::InvalidSource)?;
        charge_lifecycle_decode_report(report, budget)?;
        match snapshot.object_uuid(component, chunk_id) {
            Err(metadata::MetadataError::Missing) => {},
            Ok(_) => return Err(Error::InvalidSource),
            Err(error) => return Err(map_metadata_error(error)),
        }
        let old_build = chunk.build().identifier();
        let _registry_witness = snapshot
            .object_uuid(component, old_build)
            .map_err(map_metadata_error)?;
        let nested_uuid = chunk.chunk_identifier().map(|uuid| {
            let uuid = uuid.uuid();
            SuperUuid {
                lower: uuid.lower(),
                upper: uuid.upper(),
            }
        });
        let direct_uuid = chunk.build_id().map(|uuid| {
            let uuid = uuid.uuid();
            SuperUuid {
                lower: uuid.lower(),
                upper: uuid.upper(),
            }
        });
        if nested_uuid.is_none() || direct_uuid != nested_uuid {
            return Err(Error::InvalidSource);
        }
        let new_build = remap_lookup(&remap, old_build)?;
        let new_uuid = derive_clone_uuid(nested_uuid.ok_or(Error::InvalidSource)?, new_build);
        if snapshot.has_uuid(identity_codec::UuidBits::new(
            new_uuid.lower,
            new_uuid.upper,
        )) {
            return Err(Error::InvalidSource);
        }
        budget.charge_wire_work(generated_build_uuids.len())?;
        for (identifier, uuid, source_chunk_uuid) in &mut generated_build_uuids {
            if *identifier == new_build {
                if source_chunk_uuid.is_some_and(|source| Some(source) != nested_uuid) {
                    return Err(Error::InvalidSource);
                }
                *source_chunk_uuid = nested_uuid;
            } else if *uuid == new_uuid
                || source_chunk_uuid
                    .is_some_and(|source| derive_clone_uuid(source, *identifier) == new_uuid)
            {
                return Err(Error::InvalidSource);
            }
        }
        new_uuids[chunk_index] = new_uuid;
    }
    Ok((remap, new_ids, new_uuids))
}

fn derive_clone_uuid(source: SuperUuid, target_identifier: u64) -> SuperUuid {
    let mut digest = Sha1::new();
    digest.update(b"litchi.keynote.media.clone.v1");
    digest.update(source.lower.to_le_bytes());
    digest.update(source.upper.to_le_bytes());
    digest.update(target_identifier.to_le_bytes());
    let digest = digest.finalize();
    let mut lower = [0; 8];
    let mut upper = [0; 8];
    lower.copy_from_slice(&digest[..8]);
    upper.copy_from_slice(&digest[8..16]);
    let mut uuid = SuperUuid {
        lower: u64::from_le_bytes(lower),
        upper: u64::from_le_bytes(upper),
    };
    if uuid.lower == 0 && uuid.upper == 0 {
        uuid.upper = 1;
    }
    uuid
}

fn selection_contains_build_id(
    identifier: u64,
    source_ids: &[u64],
    component_archive: &Archive,
) -> bool {
    component_archive.object(identifier).is_some_and(|object| {
        object
            .messages
            .iter()
            .any(|message| message.type_ == BUILD_MESSAGE_TYPE)
    }) && source_ids.binary_search(&identifier).is_ok()
}

fn charge_lifecycle_decode_report(
    report: lifecycle_codec::DecodeReport,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.scratch_bytes())?;
    budget.charge_allocations(report.retained_bytes())?;
    budget.charge_references(report.fields())
}

fn append_clones(
    package: &Package,
    edited: &mut Archive,
    source_archive: &Archive,
    selection: &MediaGraphSelection,
    source_ids: &[u64],
    remap: &[(u64, u64)],
    new_ids: &[u64],
    new_uuids: &[SuperUuid],
    archive_limits: ArchiveLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    let movie = source_archive
        .object(selection.movie_identifier)
        .ok_or(Error::InvalidSource)?;
    let style_witnesses = graph_caption_witness::prove_movie_caption_style_witness(
        package,
        selection.component_name.as_ref(),
        movie,
        package
            .semantic_wire_limits()
            .map_err(|_| Error::InvalidSource)?,
        budget,
    )?;

    let mut transitive_styles = [0; 2];
    let mut style_count = 0;
    for identifier in style_witnesses.into_iter().flatten() {
        transitive_styles[style_count] = identifier;
        style_count += 1;
    }
    let mut clones = Vec::new();
    budget.charge_allocations(
        source_ids
            .len()
            .saturating_mul(size_of::<litchi_iwa_core::ArchiveObject>()),
    )?;
    clones
        .try_reserve_exact(source_ids.len())
        .map_err(|_| Error::Allocation {
            amount: source_ids
                .len()
                .saturating_mul(size_of::<litchi_iwa_core::ArchiveObject>()),
        })?;
    for (index, &source_id) in source_ids.iter().enumerate() {
        let source_object = source_archive
            .object(source_id)
            .ok_or(Error::InvalidSource)?;
        let new_id = *new_ids.get(index).ok_or(Error::InvalidSource)?;
        let uuid = selection
            .chunk_ids
            .contains(&source_id)
            .then(|| new_uuids[index]);
        let clone = graph::clone_object(
            source_object,
            new_id,
            remap,
            archive_limits,
            source_id == selection.movie_identifier,
            uuid,
            &transitive_styles[..style_count],
            budget,
        )?;
        clones.push(clone);
    }
    edited
        .append_objects(clones, archive_limits)
        .map_err(|_| Error::InvalidSource)
}

fn append_owner_additions<'a>(
    selection: &MediaGraphSelection,
    remap: &[(u64, u64)],
    component: identity_codec::ComponentSelector<'a>,
    output: &mut Vec<media_codec::DataReferenceOwnerAddition<'a>>,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    let mut counts: Vec<(u64, u64, u32)> = Vec::new();
    let count_bytes = selection
        .data_references
        .len()
        .checked_mul(size_of::<(u64, u64, u32)>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(count_bytes)?;
    counts
        .try_reserve_exact(selection.data_references.len())
        .map_err(|_| Error::Allocation {
            amount: selection
                .data_references
                .len()
                .saturating_mul(size_of::<(u64, u64, u32)>()),
        })?;
    for &(data, object) in &selection.data_references {
        let new_object = remap_lookup(remap, object)?;
        counts.push((data, new_object, 1));
    }
    budget.charge_wire_work(
        counts
            .len()
            .checked_mul(size_of::<(u64, u64, u32)>() * usize::BITS as usize)
            .ok_or(Error::InvalidSource)?,
    )?;
    counts.sort_unstable_by_key(|entry| (entry.0, entry.1));
    let output_bytes = counts
        .len()
        .checked_mul(size_of::<media_codec::DataReferenceOwnerAddition<'a>>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(output_bytes)?;
    output
        .try_reserve_exact(counts.len())
        .map_err(|_| Error::Allocation {
            amount: output_bytes,
        })?;
    let mut index = 0usize;
    while index < counts.len() {
        let (data, object, mut count) = counts[index];
        index += 1;
        while let Some(entry) = counts.get(index) {
            if entry.0 != data || entry.1 != object {
                break;
            }
            count = count.checked_add(entry.2).ok_or(Error::InvalidSource)?;
            index += 1;
        }
        output.push(media_codec::DataReferenceOwnerAddition::new(
            media_codec::ComponentSelector::new(component.identifier(), component.locator()),
            data,
            object,
            count,
        ));
    }
    Ok(())
}

fn append_owner_removals<'a>(
    selection: &MediaGraphSelection,
    component: metadata::ComponentIdentity<'_>,
    snapshot: &metadata::MetadataSnapshot<'_>,
    component_selector: identity_codec::ComponentSelector<'a>,
    output: &mut Vec<media_codec::DataReferenceOwnerRemoval<'a>>,
    data_removals: &mut Vec<media_codec::DataInfoRemoval>,
    removed_data_ids: &mut Vec<u64>,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    let mut counts: Vec<(u64, u64, u32)> = Vec::new();
    let count_bytes = selection
        .data_references
        .len()
        .checked_mul(size_of::<(u64, u64, u32)>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(count_bytes)?;
    counts
        .try_reserve_exact(selection.data_references.len())
        .map_err(|_| Error::Allocation {
            amount: selection
                .data_references
                .len()
                .saturating_mul(size_of::<(u64, u64, u32)>()),
        })?;
    for &(data, object) in &selection.data_references {
        counts.push((data, object, 1));
    }
    budget.charge_wire_work(
        counts
            .len()
            .checked_mul(size_of::<(u64, u64, u32)>() * usize::BITS as usize)
            .ok_or(Error::InvalidSource)?,
    )?;
    counts.sort_unstable_by_key(|entry| (entry.0, entry.1));
    let output_bytes = counts
        .len()
        .checked_mul(size_of::<media_codec::DataReferenceOwnerRemoval<'a>>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(output_bytes)?;
    output
        .try_reserve_exact(counts.len())
        .map_err(|_| Error::Allocation {
            amount: output_bytes,
        })?;
    let mut data_counts: Vec<(u64, usize)> = Vec::new();
    let data_count_bytes = counts
        .len()
        .checked_mul(size_of::<(u64, usize)>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(data_count_bytes)?;
    data_counts
        .try_reserve_exact(counts.len())
        .map_err(|_| Error::Allocation {
            amount: data_count_bytes,
        })?;
    let mut index = 0usize;
    while index < counts.len() {
        let (data, object, mut count) = counts[index];
        index += 1;
        while let Some(entry) = counts.get(index) {
            if entry.0 != data || entry.1 != object {
                break;
            }
            count = count.checked_add(entry.2).ok_or(Error::InvalidSource)?;
            index += 1;
        }
        let owner = snapshot
            .owner(component, data, object)
            .map_err(map_metadata_error)?;
        if owner.count() != count {
            return Err(Error::InvalidSource);
        }
        output.push(media_codec::DataReferenceOwnerRemoval::new(
            media_codec::ComponentSelector::new(
                component_selector.identifier(),
                component_selector.locator(),
            ),
            data,
            object,
            count,
        ));
        if data_counts.last().is_some_and(|entry| entry.0 == data) {
            let (_, selected) = data_counts.last_mut().ok_or(Error::InvalidSource)?;
            *selected = selected.checked_add(1).ok_or(Error::InvalidSource)?;
        } else {
            data_counts.push((data, 1));
        }
    }
    budget.charge_wire_work(
        snapshot
            .owners()
            .len()
            .checked_mul(data_counts.len().max(1))
            .ok_or(Error::InvalidSource)?,
    )?;
    for (data, selected_records) in data_counts {
        let total = snapshot
            .current_owner_count(data)
            .map_err(map_metadata_error)?;
        if selected_records > total {
            return Err(Error::InvalidSource);
        }
        if selected_records == total {
            data_removals.push(media_codec::DataInfoRemoval::new(data));
            removed_data_ids.push(data);
        }
    }
    Ok(())
}

fn validate_final_data_removals(
    package: &Package,
    catalog: &SourceCatalog,
    snapshot: &metadata::MetadataSnapshot<'_>,
    removed_objects: &[u64],
    data_removals: &[media_codec::DataInfoRemoval],
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    for removal in data_removals {
        let data = snapshot
            .data_info(removal.identifier())
            .map_err(map_metadata_error)?;
        let path_length = DATA_PREFIX
            .len()
            .checked_add(data.file_name().len())
            .ok_or(Error::InvalidSource)?;
        budget.charge_allocations(path_length)?;
        let mut path = String::new();
        path.try_reserve_exact(path_length)
            .map_err(|_| Error::Allocation {
                amount: path_length,
            })?;
        path.push_str(DATA_PREFIX);
        path.push_str(data.file_name());
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == path)
            .ok_or(Error::InvalidSource)?;
        budget.charge_media_bytes(entry.data().len())?;
        budget.charge_wire_work(entry.data().len())?;
        if entry.is_opaque()
            || data.materialized_length()
                != Some(u64::try_from(entry.data().len()).map_err(|_| Error::InvalidSource)?)
            || Sha1::digest(entry.data()).as_slice() != data.digest()
        {
            return Err(Error::InvalidSource);
        }
        for other in snapshot.data_records() {
            if other.identifier() != data.identifier() && other.file_name() == data.file_name() {
                return Err(Error::InvalidSource);
            }
        }
        for component in package.state.source.components().iter() {
            budget.charge_entries(component.archive().objects.len())?;
            for object in &component.archive().objects {
                let identifier = object.archive_info.identifier;
                if identifier.is_some_and(|value| removed_objects.binary_search(&value).is_ok()) {
                    continue;
                }
                let mut data_references = 0usize;
                let mut field_count = 0usize;
                for info in &object.archive_info.message_infos {
                    data_references = data_references
                        .checked_add(info.data_references.len())
                        .ok_or(Error::InvalidSource)?;
                    field_count = field_count
                        .checked_add(info.field_infos.len())
                        .ok_or(Error::InvalidSource)?;
                    for field in &info.field_infos {
                        data_references = data_references
                            .checked_add(field.data_references.len())
                            .ok_or(Error::InvalidSource)?;
                    }
                }
                budget.charge_references(data_references)?;
                budget.charge_wire_fields(
                    object
                        .archive_info
                        .message_infos
                        .len()
                        .checked_add(field_count)
                        .ok_or(Error::InvalidSource)?,
                )?;
                budget.charge_wire_work(data_references.saturating_add(field_count))?;
                for info in &object.archive_info.message_infos {
                    if info.data_references.contains(&data.identifier())
                        || info
                            .field_infos
                            .iter()
                            .any(|field| field.data_references.contains(&data.identifier()))
                    {
                        return Err(Error::InvalidSource);
                    }
                }
            }
        }
    }
    Ok(())
}

/// Charge and verify the selected media bytes before either a duplicate or a
/// removal is published.  A lifecycle transaction must account for the
/// physical read/hash even when the data record remains shared after removal.
fn validate_selected_media_bytes(
    catalog: &SourceCatalog,
    snapshot: &metadata::MetadataSnapshot<'_>,
    selection: &MediaGraphSelection,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(2)
        .map_err(|_| Error::Allocation {
            amount: 2usize.saturating_mul(size_of::<u64>()),
        })?;
    budget.charge_allocations(2usize.saturating_mul(size_of::<u64>()))?;
    if let Some(identifier) = selection.content_identifier {
        identifiers.push(identifier);
    }
    if let Some(identifier) = selection.poster_identifier {
        identifiers.push(identifier);
    }
    identifiers.sort_unstable();
    identifiers.dedup();
    for identifier in identifiers {
        let data = snapshot.data_info(identifier).map_err(map_metadata_error)?;
        let path_length = DATA_PREFIX
            .len()
            .checked_add(data.file_name().len())
            .ok_or(Error::InvalidSource)?;
        budget.charge_allocations(path_length)?;
        let mut path = String::new();
        path.try_reserve_exact(path_length)
            .map_err(|_| Error::Allocation {
                amount: path_length,
            })?;
        path.push_str(DATA_PREFIX);
        path.push_str(data.file_name());
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == path)
            .ok_or(Error::InvalidSource)?;
        if entry.is_opaque()
            || data.materialized_length()
                != Some(u64::try_from(entry.data().len()).map_err(|_| Error::InvalidSource)?)
            || Sha1::digest(entry.data()).as_slice() != data.digest()
        {
            return Err(Error::InvalidSource);
        }
        budget.charge_media_bytes(entry.data().len())?;
        budget.charge_wire_work(entry.data().len())?;
    }
    Ok(())
}

fn rewrite_slide_archive_object(
    archive: &mut Archive,
    identifier: u64,
    edit: lifecycle_codec::SlideLifecycleEdit<'_>,
    archive_limits: ArchiveLimits,
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    let object = archive.object_mut(identifier).ok_or(Error::InvalidSource)?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let source = &object.messages[index].data;
    let options = lifecycle_options(wire_limits, source.len())?;
    let (payload, report) =
        lifecycle_codec::rewrite_slide_lifecycle_with_report(source, edit, options)
            .map_err(|_| Error::InvalidSource)?;
    charge_lifecycle_report(report, budget)?;
    graph::replace_slide_message_with_lifecycle_refs(object, index, payload, archive_limits, budget)
}

fn lifecycle_options(
    wire_limits: WireLimits,
    source_length: usize,
) -> Result<lifecycle_codec::DecodeOptions, Error> {
    let output = source_length
        .checked_mul(2)
        .and_then(|value| value.checked_add(4096))
        .ok_or(Error::InvalidSource)?;
    let work = source_length
        .checked_mul(256)
        .and_then(|value| value.checked_add(4096))
        .ok_or(Error::InvalidSource)?;
    let depth = u32::try_from(wire_limits.max_nesting()).map_err(|_| Error::InvalidSource)?;
    Ok(lifecycle_codec::DecodeOptions::new(
        source_length.max(1),
        output,
        wire_limits.max_fields(),
        work,
        wire_limits.max_fields(),
        depth,
    )
    .with_max_message_bytes(source_length.max(1))
    .with_max_output_bytes(output)
    .with_max_fields(wire_limits.max_fields())
    .with_max_work_bytes(work)
    .with_max_references(wire_limits.max_fields())
    .with_max_depth(depth))
}

fn charge_lifecycle_report(
    report: lifecycle_codec::RewriteReport,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_wire_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.retained_bytes())?;
    budget.charge_references(report.fields())
}

fn serialize_component(
    archive: &Archive,
    archive_limits: ArchiveLimits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    budget: &mut LifecycleBudget,
) -> Result<Vec<u8>, Error> {
    let expected_length = archive
        .encoded_len(archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    budget.charge_allocations(expected_length)?;
    let bytes = archive
        .to_bytes(archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    if bytes.len() != expected_length {
        return Err(Error::InvalidSource);
    }
    let maximum_compressed = SnappyStream::maximum_compressed_len(bytes.len(), snappy_limits)
        .map_err(map_snappy_error)?;
    // `compress` reserves its maximum stream bound and uses one compressed
    // frame scratch buffer at a time.  Charging that bound before entering
    // the encoder keeps the operation ledger valid if allocation fails.
    // `compress` retains its maximum output reservation while each
    // `compress_vec` frame is alive.  The frame upper bounds can accumulate
    // across the operation, so charge a second maximum bound for that
    // transient frame storage before entering the encoder.
    let compression_allocation = maximum_compressed
        .checked_mul(2)
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(compression_allocation)?;
    let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_snappy_error)?;
    if compressed.len() > snappy_limits.max_compressed_stream() {
        return Err(Error::Limit {
            kind: LimitKind::EntryBytes,
            observed: compressed.len() as u64,
            maximum: snappy_limits.max_compressed_stream() as u64,
        });
    }
    budget.charge_output(compressed.len())?;
    Ok(compressed)
}

fn remap_lookup(remap: &[(u64, u64)], source: u64) -> Result<u64, Error> {
    remap
        .binary_search_by_key(&source, |entry| entry.0)
        .ok()
        .and_then(|index| remap.get(index).map(|entry| entry.1))
        .ok_or(Error::InvalidSource)
}

fn new_ids_for_movie(
    selection: &MediaGraphSelection,
    source_ids: &[u64],
    new_ids: &[u64],
) -> Result<u64, Error> {
    let index = source_ids
        .binary_search(&selection.movie_identifier)
        .map_err(|_| Error::InvalidSource)?;
    new_ids.get(index).copied().ok_or(Error::InvalidSource)
}

fn new_ids_for_builds(
    selection: &MediaGraphSelection,
    source_ids: &[u64],
    new_ids: &[u64],
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, Error> {
    let mut output = Vec::new();
    let bytes = selection
        .build_ids
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    output
        .try_reserve_exact(selection.build_ids.len())
        .map_err(|_| Error::Allocation {
            amount: selection.build_ids.len().saturating_mul(size_of::<u64>()),
        })?;
    for &id in &selection.build_ids {
        output.push(
            new_ids[source_ids
                .binary_search(&id)
                .map_err(|_| Error::InvalidSource)?],
        );
    }
    Ok(output)
}

fn new_ids_for_chunks(
    selection: &MediaGraphSelection,
    source_ids: &[u64],
    new_ids: &[u64],
    budget: &mut LifecycleBudget,
) -> Result<Vec<u64>, Error> {
    let mut output = Vec::new();
    let bytes = selection
        .chunk_ids
        .len()
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(bytes)?;
    output
        .try_reserve_exact(selection.chunk_ids.len())
        .map_err(|_| Error::Allocation {
            amount: selection.chunk_ids.len().saturating_mul(size_of::<u64>()),
        })?;
    for &id in &selection.chunk_ids {
        output.push(
            new_ids[source_ids
                .binary_search(&id)
                .map_err(|_| Error::InvalidSource)?],
        );
    }
    Ok(output)
}

fn verify_candidate_delta(
    source: &Package,
    candidate: &Package,
    selection: &MediaGraphSelection,
    source_ids: &[u64],
    new_ids: &[u64],
    action: LifecycleAction,
    source_media_count: usize,
    target_media_count: usize,
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    let actual_count = count_slide_media(candidate, selection, wire_limits, budget)?;
    if actual_count != target_media_count {
        return Err(Error::Verification);
    }
    match action {
        LifecycleAction::Duplicate => {
            if actual_count != source_media_count.saturating_add(1) {
                return Err(Error::Verification);
            }
            for &identifier in new_ids {
                if candidate.object(identifier).is_none() {
                    return Err(Error::Verification);
                }
            }
            if let Some(source_plan) = selection.comment_graph.as_ref() {
                let source_position = source_ids
                    .binary_search(&source_plan.root_storage_identifier)
                    .map_err(|_| Error::Verification)?;
                let new_root = *new_ids.get(source_position).ok_or(Error::Verification)?;
                let cloned_plan = comment_graph::plan_comment_graph(
                    candidate,
                    selection.component_name.as_ref(),
                    new_root,
                    wire_limits,
                    budget,
                )?;
                if cloned_plan.root_storage_uuid() != source_plan.root_storage_uuid()
                    || cloned_plan.author_ids != source_plan.author_ids
                    || cloned_plan.storage_identities.len() != source_plan.storage_identities.len()
                {
                    return Err(Error::Verification);
                }
                for (source_identity, cloned_identity) in source_plan
                    .storage_identities
                    .iter()
                    .zip(&cloned_plan.storage_identities)
                {
                    budget.charge_wire_work(source_ids.len())?;
                    let position = source_ids
                        .binary_search(&source_identity.identifier)
                        .map_err(|_| Error::Verification)?;
                    if new_ids.get(position) != Some(&cloned_identity.identifier)
                        || source_identity.uuid != cloned_identity.uuid
                    {
                        return Err(Error::Verification);
                    }
                }
            }
        },
        LifecycleAction::Remove => {
            for &identifier in source_ids {
                if candidate.object(identifier).is_some() {
                    return Err(Error::Verification);
                }
            }
            graph::validate_removal_closure(candidate, source_ids, budget)?;
        },
    }
    budget.charge_references(new_ids.len().saturating_add(source_ids.len()))?;
    let _ = source;
    Ok(())
}

fn verify_candidate_selection(
    candidate: &Package,
    patch: &Patch,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    let limits = candidate
        .semantic_wire_limits()
        .map_err(|_| Error::Verification)?;
    let count = count_slide_media(candidate, &patch.source_selection, limits, budget)?;
    if count != patch.selection.target_media_count {
        return Err(Error::Verification);
    }
    budget.charge_references(count)?;
    Ok(())
}

fn verify_zip_locality(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    changed_members: &str,
    changed_node_member: Option<&str>,
    deleted_data_paths: &[String],
    deleted_preview_names: &[&str],
) -> Result<(), Error> {
    let mut candidates = candidate.package().iter();
    for entry in source.package().iter() {
        if deleted_preview_names.contains(&entry.name())
            || deleted_data_paths.iter().any(|name| name == entry.name())
        {
            continue;
        }
        let candidate_entry = candidates.next().ok_or(Error::Verification)?;
        if entry.name() != candidate_entry.name() {
            return Err(Error::Verification);
        }
        if entry.name() == changed_members
            || entry.name() == METADATA_COMPONENT
            || Some(entry.name()) == changed_node_member
        {
            if !super::rendering_invalidation::selected_package_member_preserved(
                entry,
                candidate_entry,
            ) {
                return Err(Error::Verification);
            }
            continue;
        }
        if entry.raw_name() != candidate_entry.raw_name()
            || entry.is_opaque() != candidate_entry.is_opaque()
            || entry.data() != candidate_entry.data()
            || entry.metadata() != candidate_entry.metadata()
            || entry.raw_record().local_record() != candidate_entry.raw_record().local_record()
            || entry.raw_record().compressed_data()
                != candidate_entry.raw_record().compressed_data()
            || !super::rendering_invalidation::central_record_preserved(
                entry.raw_record().central_directory_record(),
                candidate_entry.raw_record().central_directory_record(),
            )
        {
            return Err(Error::Verification);
        }
    }
    if candidates.next().is_some() {
        return Err(Error::Verification);
    }
    Ok(())
}

struct MetadataArchive {
    archive: Archive,
    object_identifier: u64,
    message_index: usize,
    payload: Vec<u8>,
}

fn rewrite_lifecycle(
    source: &Package,
    catalog: &SourceCatalog,
    selection: MediaGraphSelection,
    action: LifecycleAction,
    source_media_count: usize,
    target_media_count: usize,
    wire_limits: WireLimits,
    budget: &mut LifecycleBudget,
) -> Result<(Package, Patch), Error> {
    let archive_limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let snappy_limits = catalog
        .limits()
        .snappy_limits()
        .map_err(|_| Error::InvalidSource)?;
    let metadata = load_metadata(source, catalog, archive_limits, snappy_limits, budget)?;
    let metadata_options = metadata_options(source, metadata.payload.len(), budget)?;
    let map = map_witness(
        source,
        &metadata.payload,
        media_codec::DecodeOptions::for_source(&metadata.payload),
        wire_limits,
        budget,
    )?;
    let locator = selection
        .component_name
        .strip_prefix("Index/")
        .and_then(|value| value.strip_suffix(".iwa"))
        .ok_or(Error::InvalidSource)?;
    let snapshot = metadata_snapshot(
        &metadata.payload,
        &metadata_options.0,
        &metadata_options.1,
        map,
        budget,
    )?;
    validate_selected_media_bytes(catalog, &snapshot, &selection, budget)?;
    let component = snapshot
        .current_component(locator)
        .map_err(map_metadata_error)?;
    if let Some(plan) = selection.comment_graph.as_ref() {
        for dependency in &plan.author_dependencies {
            budget.charge_references(1)?;
            if dependency.component_name == selection.component_name {
                continue;
            }
            let target_locator = dependency
                .component_name
                .strip_prefix("Index/")
                .and_then(|value| value.strip_suffix(".iwa"))
                .ok_or(Error::InvalidSource)?;
            budget.charge_wire_work(snapshot.external_dependency_lookup_work())?;
            snapshot
                .require_current_external_dependency(
                    component,
                    target_locator,
                    dependency.author_identifier,
                )
                .map_err(map_metadata_error)?;
        }
    }
    let selected_source_ids = selected_object_ids(&selection, budget)?;
    let removal_plan = if action == LifecycleAction::Remove {
        selection
            .comment_graph
            .as_ref()
            .map(|plan| {
                comment_removal::plan_comment_removal(
                    source,
                    selection.component_name.as_ref(),
                    &selected_source_ids,
                    plan,
                    archive_limits,
                    budget,
                )
            })
            .transpose()?
    } else {
        None
    };
    let source_ids = removal_plan
        .as_ref()
        .map_or(selected_source_ids.as_slice(), |plan| {
            plan.removed_object_ids.as_slice()
        });
    let released_watermark = if action == LifecycleAction::Remove {
        identifier_watermark::plan_release(source, source_ids, snapshot.last_identifier(), budget)?
    } else {
        None
    };
    let component_archive = source
        .state
        .source
        .components()
        .iter()
        .find(|candidate| candidate.name() == selection.component_name.as_ref())
        .ok_or(Error::InvalidSource)?
        .archive();
    let component_clone_bytes = component_archive
        .encoded_len(archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    budget.charge_allocations(component_clone_bytes)?;
    let mut edited_component = component_archive.clone();

    let (remap, new_ids, new_uuids) = match action {
        LifecycleAction::Duplicate => allocate_clone_identities(
            source,
            &edited_component,
            &snapshot,
            component,
            source_ids,
            wire_limits,
            budget,
        )?,
        LifecycleAction::Remove => (Vec::new(), Vec::new(), Vec::new()),
    };
    let mut identity_additions = Vec::new();
    let mut identity_removals = Vec::new();
    let mut external_removals = Vec::new();
    if let Some(plan) = removal_plan.as_ref() {
        let bytes = plan
            .unused_external_author_ids
            .len()
            .checked_mul(size_of::<identity_codec::ExternalReferenceRemoval<'_>>())
            .ok_or(Error::InvalidSource)?;
        if bytes != 0 {
            budget.charge_allocations(bytes)?;
            external_removals
                .try_reserve_exact(plan.unused_external_author_ids.len())
                .map_err(|_| Error::Allocation { amount: bytes })?;
        }
        let comment_graph = selection
            .comment_graph
            .as_ref()
            .ok_or(Error::InvalidSource)?;
        for &author_identifier in &plan.unused_external_author_ids {
            budget.charge_wire_work(comment_graph.author_dependencies.len().max(1))?;
            let dependency = comment_graph
                .author_dependencies
                .iter()
                .find(|dependency| dependency.author_identifier == author_identifier)
                .ok_or(Error::InvalidSource)?;
            let target_locator = dependency
                .component_name
                .strip_prefix("Index/")
                .and_then(|value| value.strip_suffix(".iwa"))
                .ok_or(Error::InvalidSource)?;
            budget.charge_wire_work(snapshot.external_dependency_lookup_work())?;
            external_removals.push(
                snapshot
                    .prepare_current_external_dependency_removal(
                        component,
                        target_locator,
                        author_identifier,
                    )
                    .map_err(map_metadata_error)?,
            );
        }
    }
    let mut owner_additions = Vec::new();
    let mut owner_removals = Vec::new();
    let mut data_removals = Vec::new();
    let mut removed_data_ids = Vec::new();
    let identity_capacity_bytes = source_ids
        .len()
        .checked_mul(size_of::<identity_codec::ObjectUuidAddition<'_>>())
        .and_then(|bytes| {
            source_ids
                .len()
                .checked_mul(size_of::<identity_codec::ObjectUuidRemoval<'_>>())
                .and_then(|removal_bytes| bytes.checked_add(removal_bytes))
        })
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(identity_capacity_bytes)?;
    identity_additions
        .try_reserve_exact(source_ids.len())
        .map_err(|_| Error::Allocation {
            amount: identity_capacity_bytes,
        })?;
    identity_removals
        .try_reserve_exact(source_ids.len())
        .map_err(|_| Error::Allocation {
            amount: identity_capacity_bytes,
        })?;
    let owner_capacity = selection.data_references.len();
    let data_capacity_bytes = owner_capacity
        .checked_mul(size_of::<media_codec::DataInfoRemoval>())
        .and_then(|bytes| bytes.checked_add(owner_capacity.saturating_mul(size_of::<u64>())))
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(data_capacity_bytes)?;
    data_removals
        .try_reserve_exact(owner_capacity)
        .map_err(|_| Error::Allocation {
            amount: owner_capacity.saturating_mul(size_of::<media_codec::DataInfoRemoval>()),
        })?;
    removed_data_ids
        .try_reserve_exact(owner_capacity)
        .map_err(|_| Error::Allocation {
            amount: owner_capacity.saturating_mul(size_of::<u64>()),
        })?;
    let component_selector = component.selector();

    match action {
        LifecycleAction::Duplicate => {
            append_clones(
                source,
                &mut edited_component,
                component_archive,
                &selection,
                source_ids,
                &remap,
                &new_ids,
                &new_uuids,
                archive_limits,
                budget,
            )?;

            for ((old, new), uuid) in remap.iter().zip(new_uuids.iter()) {
                match snapshot.object_uuid(component, *old) {
                    Ok(_) => identity_additions.push(identity_codec::ObjectUuidAddition::new(
                        component_selector,
                        *new,
                        identity_codec::UuidBits::new(uuid.lower, uuid.upper),
                    )),
                    Err(metadata::MetadataError::Missing) => {},
                    Err(error) => return Err(map_metadata_error(error)),
                }
            }
            append_owner_additions(
                &selection,
                &remap,
                component_selector,
                &mut owner_additions,
                budget,
            )?;
        },
        LifecycleAction::Remove => {
            for identifier in source_ids {
                match snapshot.object_uuid(component, *identifier) {
                    Ok(uuid) => identity_removals.push(identity_codec::ObjectUuidRemoval::new(
                        component_selector,
                        *identifier,
                        uuid,
                    )),
                    Err(metadata::MetadataError::Missing) => {},
                    Err(error) => return Err(map_metadata_error(error)),
                }
            }
            append_owner_removals(
                &selection,
                component,
                &snapshot,
                component_selector,
                &mut owner_removals,
                &mut data_removals,
                &mut removed_data_ids,
                budget,
            )?;
            validate_final_data_removals(
                source,
                catalog,
                &snapshot,
                source_ids,
                &data_removals,
                budget,
            )?;
        },
    }

    let mut appended_build_ids = Vec::new();
    let mut appended_chunk_ids = Vec::new();
    let mut slide_remove_ids = Vec::new();
    if matches!(action, LifecycleAction::Duplicate) {
        appended_build_ids = new_ids_for_builds(&selection, source_ids, &new_ids, budget)?;
        appended_chunk_ids = new_ids_for_chunks(&selection, source_ids, &new_ids, budget)?;
    } else {
        let remove_capacity = 1usize
            .checked_add(selection.build_ids.len())
            .and_then(|value| value.checked_add(selection.chunk_ids.len()))
            .ok_or(Error::InvalidSource)?;
        let remove_bytes = remove_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?;
        budget.charge_allocations(remove_bytes)?;
        slide_remove_ids
            .try_reserve_exact(remove_capacity)
            .map_err(|_| Error::Allocation {
                amount: remove_bytes,
            })?;
        slide_remove_ids.push(selection.movie_identifier);
        slide_remove_ids.extend(selection.build_ids.iter().copied());
        slide_remove_ids.extend(selection.chunk_ids.iter().copied());
    }
    let movie_append = if matches!(action, LifecycleAction::Duplicate) {
        Some(new_ids_for_movie(&selection, source_ids, &new_ids)?)
    } else {
        None
    };
    let slide_id_edit = if let Some(movie_append) = movie_append.as_ref() {
        lifecycle_codec::SlideLifecycleEdit::empty()
            .with_owned_drawables(std::slice::from_ref(movie_append))
            .with_drawables_z_order(std::slice::from_ref(movie_append))
            .with_builds(&appended_build_ids)
            .with_build_chunks(&appended_chunk_ids)
    } else {
        lifecycle_codec::SlideLifecycleEdit::empty().with_removed_identifiers(&slide_remove_ids)
    };
    rewrite_slide_archive_object(
        &mut edited_component,
        selection.slide_identifier,
        slide_id_edit,
        archive_limits,
        wire_limits,
        budget,
    )?;
    if matches!(action, LifecycleAction::Remove) {
        for identifier in source_ids.iter().copied() {
            edited_component
                .remove_object_checked(identifier, archive_limits)
                .map_err(|_| Error::InvalidSource)?;
        }
    }

    let node_edit = node_cache::prepare_slide_node_build_cache(
        source,
        &selection,
        selection.component_name.as_ref(),
        &mut edited_component,
        archive_limits,
        wire_limits,
        budget,
    )?;

    let identity_batch = match action {
        LifecycleAction::Duplicate => {
            let new_last = *new_ids.last().ok_or(Error::InvalidSource)?;
            metadata::IdentityBatch::additions(
                snapshot.last_identifier(),
                new_last,
                &identity_additions,
            )
        },
        LifecycleAction::Remove => {
            let batch =
                metadata::IdentityBatch::removals(snapshot.last_identifier(), &identity_removals)
                    .with_external_reference_removals(&external_removals);
            match released_watermark {
                Some(new_last_identifier) => batch.with_new_last_identifier(new_last_identifier),
                None => batch,
            }
        },
    };
    let mut media_batch =
        media_codec::MediaRewriteBatch::new(&[], &data_removals, &owner_additions, &owner_removals);
    if let Some(map_source) = map {
        media_batch = media_batch.with_data_metadata_map_source(map_source);
    }
    let mut metadata_budget_error = None;
    let mut charge_metadata = |amount: usize| match budget.charge_allocations(amount) {
        Ok(()) => Ok(()),
        Err(error) => {
            metadata_budget_error = Some(error);
            Err(metadata::MetadataError::Allocation { amount })
        },
    };
    let rewritten_metadata = metadata::rewrite_metadata(
        &snapshot,
        identity_batch,
        media_batch,
        metadata_options.0,
        metadata_options.1,
        &mut charge_metadata,
    )
    .map_err(|error| {
        metadata_budget_error
            .take()
            .unwrap_or_else(|| map_metadata_error(error))
    })?;
    let (metadata_payload, removed_paths, identity_report, media_report) =
        rewritten_metadata.into_parts();
    if let Some(report) = identity_report {
        charge_identity_snapshot_report(report, budget)?;
    }
    if let Some(report) = media_report {
        charge_media_report(report, budget)?;
    }
    let mut metadata_archive = metadata.archive;
    metadata_archive
        .object_mut(metadata.object_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            metadata.message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data: metadata_payload,
            },
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;

    let component_bytes =
        serialize_component(&edited_component, archive_limits, snappy_limits, budget)?;
    let metadata_bytes =
        serialize_component(&metadata_archive, archive_limits, snappy_limits, budget)?;
    let node_bytes = node_edit
        .as_ref()
        .map(|edit| serialize_component(&edit.archive, archive_limits, snappy_limits, budget))
        .transpose()?;
    let component_edit = EntryEdit::new(selection.component_name.as_ref(), &component_bytes);
    let metadata_edit = EntryEdit::new(METADATA_COMPONENT, &metadata_bytes);
    let preview_plan = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| Error::InvalidSource)?;
    let removed_path_capacity = removed_paths.len();
    let removed_path_bytes = removed_paths.iter().try_fold(0usize, |total, path| {
        total
            .checked_add(DATA_PREFIX.len())
            .and_then(|total| total.checked_add(path.len()))
            .ok_or(Error::InvalidSource)
    })?;
    budget.charge_allocations(
        removed_path_capacity
            .checked_mul(size_of::<String>())
            .and_then(|bytes| bytes.checked_add(removed_path_bytes))
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut deleted_names = Vec::new();
    let mut removed_data_paths = Vec::new();
    removed_data_paths
        .try_reserve_exact(removed_path_capacity)
        .map_err(|_| Error::Allocation {
            amount: removed_path_capacity
                .saturating_mul(size_of::<String>())
                .saturating_add(removed_path_bytes),
        })?;
    for path in removed_paths {
        let mut full = String::new();
        full.try_reserve_exact(DATA_PREFIX.len().saturating_add(path.len()))
            .map_err(|_| Error::Allocation {
                amount: DATA_PREFIX.len().saturating_add(path.len()),
            })?;
        full.push_str(DATA_PREFIX);
        full.push_str(path);
        removed_data_paths.push(full);
    }
    let deleted_capacity = removed_data_paths
        .len()
        .checked_add(preview_plan.len())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(
        deleted_capacity
            .checked_mul(size_of::<&str>())
            .ok_or(Error::InvalidSource)?,
    )?;
    deleted_names
        .try_reserve_exact(deleted_capacity)
        .map_err(|_| Error::Allocation {
            amount: deleted_capacity.saturating_mul(size_of::<&str>()),
        })?;
    deleted_names.extend(removed_data_paths.iter().map(String::as_str));
    deleted_names.extend(preview_plan.names());
    let node_entry_edit = node_edit
        .as_ref()
        .zip(node_bytes.as_ref())
        .map(|(edit, bytes)| EntryEdit::new(edit.name.as_ref(), bytes));
    let mut entry_edits = [component_edit, metadata_edit, metadata_edit];
    let edit_count = if let Some(edit) = node_entry_edit {
        entry_edits[2] = edit;
        3
    } else {
        2
    };
    let prepared = catalog
        .package()
        .prepare_reassembly_with_changes(
            &[],
            &entry_edits[..edit_count],
            &deleted_names,
            catalog.limits(),
        )
        .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    budget.charge_output(requirements.output_bytes())?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(|_| Error::InvalidSource)?;
    let target: Arc<[u8]> = output.into();
    let candidate = Package::from_source_with_options(Arc::clone(&target), source.state.options)
        .map_err(|_| Error::Read)?;
    candidate.validate().map_err(|_| Error::Verification)?;

    let candidate_catalog = physical_catalog(&candidate)?;
    node_cache::validate_candidate_package_node_cache(&candidate, &selection, wire_limits, budget)?;
    verify_candidate_delta(
        source,
        &candidate,
        &selection,
        source_ids,
        &new_ids,
        action,
        source_media_count,
        target_media_count,
        wire_limits,
        budget,
    )?;
    if let Some(plan) = removal_plan.as_ref() {
        verify_comment_removal_retention(source, &candidate, &selection, plan, budget)?;
    }
    if let Some(new_last_identifier) = released_watermark {
        identifier_watermark::verify_release(&candidate, new_last_identifier, budget)?;
    }
    verify_zip_locality(
        catalog,
        candidate_catalog,
        &selection.component_name,
        node_edit.as_ref().map(|edit| edit.name.as_ref()),
        &removed_data_paths,
        preview_plan.names(),
    )?;
    let patch = Patch {
        artifacts: ExactArtifacts::new(catalog.shared_source(), target),
        action,
        selection: SelectionFingerprint {
            slide_position: selection.slide_position,
            movie_position: selection.movie_position,
            kind: selection.kind,
            source_media_count,
            target_media_count,
        },
        source_selection: Arc::new(selection),
        created_objects: new_ids.len(),
        removed_objects: if matches!(action, LifecycleAction::Remove) {
            source_ids.len()
        } else {
            0
        },
        removed_data: removed_data_ids.len(),
        touched_members: edit_count
            .checked_add(removed_data_paths.len())
            .and_then(|value| value.checked_add(preview_plan.len()))
            .ok_or(Error::InvalidSource)?,
    };
    Ok((candidate, patch))
}

/// Recheck all retained comment nodes and shared authors against the exact source.
fn verify_comment_removal_retention(
    source: &Package,
    candidate: &Package,
    selection: &MediaGraphSelection,
    plan: &comment_removal::CommentRemovalPlan,
    budget: &mut LifecycleBudget,
) -> Result<(), Error> {
    let graph = selection
        .comment_graph
        .as_ref()
        .ok_or(Error::Verification)?;
    for &identifier in plan
        .retained_comment_storage_ids
        .iter()
        .chain(&graph.author_ids)
    {
        budget.charge_references(1)?;
        let (before_component, before) = source
            .object_with_component(identifier)
            .ok_or(Error::Verification)?;
        let (after_component, after) = candidate
            .object_with_component(identifier)
            .ok_or(Error::Verification)?;
        let bytes = before.messages.iter().try_fold(
            usize::try_from(before.header_length).map_err(|_| Error::Verification)?,
            |bytes, message| {
                bytes
                    .checked_add(message.data.len())
                    .ok_or(Error::Verification)
            },
        )?;
        budget.charge_wire_work(bytes.max(1))?;
        if before_component != after_component
            || before.archive_info != after.archive_info
            || before.messages != after.messages
        {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

/// Preserve a retained Snappy encoder refusal as a typed owner failure
/// (see [`SnappyFailure`]).
fn map_snappy_error(error: litchi_iwa_core::Error) -> Error {
    match SnappyFailure::from(error) {
        SnappyFailure::CompressedBytes { observed, maximum } => Error::Limit {
            kind: LimitKind::EntryBytes,
            observed,
            maximum,
        },
        SnappyFailure::DecodedBytes { observed, maximum } => Error::Limit {
            kind: LimitKind::EntryBytes,
            observed,
            maximum,
        },
        SnappyFailure::Work { observed, maximum } => Error::Limit {
            kind: LimitKind::WireWork,
            observed,
            maximum,
        },
        SnappyFailure::Allocation { requested } => Error::Allocation { amount: requested },
        SnappyFailure::Invalid => Error::InvalidSource,
    }
}
