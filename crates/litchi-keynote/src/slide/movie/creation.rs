//! Semantic evidence and exact-source patches for fresh slide movies.
//!
//! Native object identities, archive records, and data paths remain private
//! to the package owner. A patch authorizes one exact source and retains its
//! exact inverse, including previews removed during creation.

use std::fmt;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use thiserror::Error as ThisError;

use super::Options;
use crate::{Package, SlideSelector};

/// Finite resources consumed by one movie-creation transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Source package bytes.
    InputBytes,
    /// Candidate package bytes.
    OutputBytes,
    /// Package entries.
    Entries,
    /// Native objects inspected or created privately by the transaction.
    Objects,
    /// Bytes in one entry.
    EntryBytes,
    /// Aggregate retained bytes.
    TotalBytes,
    /// Slides.
    Slides,
    /// Object references.
    References,
    /// Aggregate movie and poster bytes.
    MediaBytes,
    /// Wire fields.
    WireFields,
    /// Wire nesting.
    WireNesting,
    /// Wire work.
    WireWork,
    /// Allocations.
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "source package bytes",
            Self::OutputBytes => "candidate package bytes",
            Self::Entries => "package entries",
            Self::Objects => "archive objects",
            Self::EntryBytes => "bytes in one entry",
            Self::TotalBytes => "aggregate retained bytes",
            Self::Slides => "slides",
            Self::References => "object references",
            Self::MediaBytes => "movie and poster bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
        })
    }
}

/// Content-redacted failures from fresh slide-movie creation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The source graph or metadata cannot safely admit creation.
    #[error("the Keynote source graph cannot safely admit slide movie creation")]
    InvalidSource,
    /// A configured finite resource ceiling was exceeded.
    #[error("slide movie creation {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    /// An allocation failed before publication.
    #[error("could not allocate {amount} units for slide movie creation")]
    Allocation { amount: usize },
    /// The reopened candidate did not reproduce the planned graph and movie.
    #[error("the created slide movie failed verification")]
    Verification,
    /// The patch is not authorized for this exact source snapshot.
    #[error("the slide movie creation patch does not match this exact source")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The package has no exact physical source suitable for this transaction.
    #[error("this Keynote source does not support slide movie creation")]
    UnsupportedSource,
    /// The slide name is empty.
    #[error("Keynote slide name must not be empty")]
    EmptySlideName,
    /// No slide has the requested name.
    #[error("the requested Keynote slide name was not found")]
    SlideNameNotFound,
    /// More than one slide matches the selector.
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    /// The requested slide position is absent.
    #[error("Keynote slide position {position:?} was not found")]
    SlidePositionNotFound { position: Position },
    /// The movie filename is not a bounded safe video filename.
    #[error("slide movie requires one safe recognized movie filename")]
    InvalidMovieFilename,
    /// The poster filename is not a bounded safe image filename.
    #[error("slide movie requires one safe recognized poster filename")]
    InvalidPosterFilename,
    /// The payload is empty or has no recognized video signature.
    #[error("slide movie requires a nonempty recognized movie payload")]
    UnsupportedMovie,
    /// The poster payload is empty or has no recognized image signature.
    #[error("slide movie requires a nonempty recognized poster payload")]
    UnsupportedPoster,
    /// The edit already stages its one creation.
    #[error("the slide movie creation edit already stages a movie")]
    OperationAlreadyStaged,
    /// The edit was committed without a staged creation.
    #[error("the slide movie creation edit stages no movie")]
    NoStagedOperation,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. }
            | Self::UnsupportedMovie { .. }
            | Self::UnsupportedPoster { .. } => RefusalKind::UnsupportedContent,
            Self::EmptySlideName { .. }
            | Self::AmbiguousSelector { .. }
            | Self::InvalidMovieFilename { .. }
            | Self::InvalidPosterFilename { .. } => RefusalKind::InvalidValue,
            Self::SlideNameNotFound { .. } | Self::SlidePositionNotFound { .. } => {
                RefusalKind::NotFound
            },
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

/// Exact source and target artifacts for creating one slide-owned movie.
#[must_use]
#[derive(Clone)]
pub struct Patch {
    pub(crate) artifacts: ExactArtifacts,
    pub(crate) slide_position: Position,
    pub(crate) movie_position: Position,
    pub(crate) source_media_count: usize,
    pub(crate) target_media_count: usize,
    pub(crate) options: Options,
    pub(crate) movie_data_digest: [u8; 20],
    pub(crate) movie_data_len: usize,
    pub(crate) poster_data_digest: [u8; 20],
    pub(crate) poster_data_len: usize,
    pub(crate) target_contains_created_movie: bool,
    pub(crate) created_objects: usize,
    pub(crate) removed_objects: usize,
    pub(crate) created_data: usize,
    pub(crate) removed_data: usize,
    pub(crate) touched_members: usize,
    pub(crate) deleted_previews: usize,
    pub(crate) restored_previews: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("slide_position", &self.slide_position)
            .field("movie_position", &self.movie_position)
            .field("source_media_count", &self.source_media_count)
            .field("target_media_count", &self.target_media_count)
            .field("options", &self.options)
            .field("movie_bytes", &self.movie_data_len)
            .field("poster_bytes", &self.poster_data_len)
            .field("created_objects", &self.created_objects)
            .field("removed_objects", &self.removed_objects)
            .field("created_data", &self.created_data)
            .field("removed_data", &self.removed_data)
            .field("touched_members", &self.touched_members)
            .field("deleted_previews", &self.deleted_previews)
            .field("restored_previews", &self.restored_previews)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the diagnostic fingerprint of the exact source artifact.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the diagnostic fingerprint of the exact target artifact.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return the validated placement, dimensions, and canonical duration of
    /// the new movie.
    #[must_use]
    pub const fn options(&self) -> Options {
        self.options
    }

    /// Return whether source and target are byte-identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse, including any removed or restored previews.
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            source_media_count: self.target_media_count,
            target_media_count: self.source_media_count,
            target_contains_created_movie: !self.target_contains_created_movie,
            created_objects: self.removed_objects,
            removed_objects: self.created_objects,
            created_data: self.removed_data,
            removed_data: self.created_data,
            deleted_previews: self.restored_previews,
            restored_previews: self.deleted_previews,
            ..self.clone()
        }
    }

    /// Return the selected source-order slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the new movie position in the slide media sequence.
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.movie_position
    }

    /// Return the media count in the source.
    #[must_use]
    pub const fn source_media_count(&self) -> usize {
        self.source_media_count
    }

    /// Return the media count in the target.
    #[must_use]
    pub const fn target_media_count(&self) -> usize {
        self.target_media_count
    }

    /// Return the SHA-1 digest of the source movie payload.
    #[must_use]
    pub const fn movie_data_digest(&self) -> [u8; 20] {
        self.movie_data_digest
    }

    /// Return the source movie payload length in bytes.
    #[must_use]
    pub const fn movie_data_len(&self) -> usize {
        self.movie_data_len
    }

    /// Return the SHA-1 digest of the source poster payload.
    #[must_use]
    pub const fn poster_data_digest(&self) -> [u8; 20] {
        self.poster_data_digest
    }

    /// Return the source poster payload length in bytes.
    #[must_use]
    pub const fn poster_data_len(&self) -> usize {
        self.poster_data_len
    }

    /// Return the number of created archive objects.
    #[must_use]
    pub const fn created_objects(&self) -> usize {
        self.created_objects
    }

    /// Return the number of removed archive objects.
    #[must_use]
    pub const fn removed_objects(&self) -> usize {
        self.removed_objects
    }

    /// Return the number of inserted movie and poster data records.
    #[must_use]
    pub const fn created_data(&self) -> usize {
        self.created_data
    }

    /// Return the number of removed movie and poster data records.
    #[must_use]
    pub const fn removed_data(&self) -> usize {
        self.removed_data
    }

    /// Return the number of changed package members.
    #[must_use]
    pub const fn touched_members(&self) -> usize {
        self.touched_members
    }

    /// Return the number of deleted previews.
    #[must_use]
    pub const fn deleted_previews(&self) -> usize {
        self.deleted_previews
    }

    /// Return the number of restored previews.
    #[must_use]
    pub const fn restored_previews(&self) -> usize {
        self.restored_previews
    }
}

/// Resource and topology evidence for a verified creation or its inverse.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    pub(crate) changed: bool,
    pub(crate) source_media_count: usize,
    pub(crate) target_media_count: usize,
    pub(crate) created_objects: usize,
    pub(crate) removed_objects: usize,
    pub(crate) created_data: usize,
    pub(crate) removed_data: usize,
    pub(crate) touched_members: usize,
    pub(crate) deleted_previews: usize,
    pub(crate) restored_previews: usize,
}

impl Diagnostics {
    pub(crate) fn for_patch(patch: &Patch) -> Self {
        Self {
            changed: !patch.is_noop(),
            source_media_count: patch.source_media_count,
            target_media_count: patch.target_media_count,
            created_objects: patch.created_objects,
            removed_objects: patch.removed_objects,
            created_data: patch.created_data,
            removed_data: patch.removed_data,
            touched_members: patch.touched_members,
            deleted_previews: patch.deleted_previews,
            restored_previews: patch.restored_previews,
        }
    }

    /// Return whether the transaction changed the package bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the media count in the source.
    #[must_use]
    pub const fn source_media_count(self) -> usize {
        self.source_media_count
    }

    /// Return the media count in the target.
    #[must_use]
    pub const fn target_media_count(self) -> usize {
        self.target_media_count
    }

    /// Return the number of created archive objects.
    #[must_use]
    pub const fn created_objects(self) -> usize {
        self.created_objects
    }

    /// Return the number of removed archive objects.
    #[must_use]
    pub const fn removed_objects(self) -> usize {
        self.removed_objects
    }

    /// Return the number of inserted movie and poster data records.
    #[must_use]
    pub const fn created_data(self) -> usize {
        self.created_data
    }

    /// Return the number of removed movie and poster data records.
    #[must_use]
    pub const fn removed_data(self) -> usize {
        self.removed_data
    }

    /// Return the number of changed package members.
    #[must_use]
    pub const fn touched_members(self) -> usize {
        self.touched_members
    }

    /// Return the number of deleted previews.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return the number of restored previews.
    #[must_use]
    pub const fn restored_previews(self) -> usize {
        self.restored_previews
    }
}

/// A reopened package and its exact-source reversible creation patch.
#[must_use]
#[derive(Debug)]
pub struct Commit {
    pub(crate) package: Package,
    pub(crate) patch: Patch,
    pub(crate) diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the verified package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its verified package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source patch.
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the creation diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// One staged slide movie creation.
#[derive(Debug, Clone, Copy)]
struct Staged<'a> {
    slide: SlideSelector<'a>,
    preferred_movie_filename: &'a str,
    movie_data: &'a [u8],
    preferred_poster_filename: &'a str,
    poster_data: &'a [u8],
    options: Options,
}

/// One immutable edit that stages exactly one independently positioned,
/// file-backed movie on a slide.
///
/// Start it with [`Package::edit_slide_movie_creation`]; the source package is
/// never changed, and dropping the edit discards the staged movie. Both movie
/// and poster bytes stay borrowed until the commit has assembled and reopened
/// the candidate.
pub struct Edit<'a> {
    source: &'a Package,
    staged: Option<Staged<'a>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide", &self.staged.as_ref().map(|staged| staged.slide))
            .field(
                "movie_data_len",
                &self.staged.as_ref().map(|staged| staged.movie_data.len()),
            )
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    pub(crate) const fn new(source: &'a Package) -> Self {
        Self {
            source,
            staged: None,
        }
    }

    /// Stage one movie on the selected slide: `movie_data` (named after
    /// `preferred_movie_filename`) with the poster frame `poster_data` (named
    /// after `preferred_poster_filename`), placed by `options`.
    ///
    /// # Errors
    ///
    /// [`Refusal::OperationAlreadyStaged`] after one creation.
    pub fn add(
        mut self,
        slide: impl Into<SlideSelector<'a>>,
        preferred_movie_filename: &'a str,
        movie_data: &'a [u8],
        preferred_poster_filename: &'a str,
        poster_data: &'a [u8],
        options: Options,
    ) -> Result<Self, Error> {
        if self.staged.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        self.staged = Some(Staged {
            slide: slide.into(),
            preferred_movie_filename,
            movie_data,
            preferred_poster_filename,
            poster_data,
            options,
        });
        Ok(self)
    }

    /// Validate, publish and reopen the staged movie as one exact-source
    /// transaction; the returned patch authorizes replay only for that exact
    /// source.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit; selector, input,
    /// source, resource and verification errors. Failure never publishes a
    /// partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let staged = self
            .staged
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        self.source.create_slide_movie(
            staged.slide,
            staged.preferred_movie_filename,
            staged.movie_data,
            staged.preferred_poster_filename,
            staged.poster_data,
            staged.options,
        )
    }
}
