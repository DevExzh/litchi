//! Cheaply shareable archive-free Keynote document snapshots.

use std::fmt;
use std::path::Path;
use std::sync::Arc;

use thiserror::Error as ThisError;

use crate::package::{ReadError as PackageReadError, SemanticLimits as PackageSemanticLimits};
use crate::show::Show;

/// Stable source-capture resource category for archive-free Keynote ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DocumentSourceLimitKind {
    /// Encoded bytes supplied by the caller.
    InputBytes,
    /// Independently addressed items discovered in the source.
    Entries,
    /// Decoded bytes contributed by one source item.
    EntryBytes,
    /// Aggregate decoded bytes contributed by all source items.
    AggregateBytes,
    /// Bytes contributed by one document payload component.
    ComponentBytes,
}

impl fmt::Display for DocumentSourceLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
            Self::ComponentBytes => "component bytes",
        })
    }
}

/// Invalid caller-selected source-capture limits.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[error(
    "Keynote document source {kind} limit must be non-zero and no greater than {maximum}, got {value}"
)]
#[non_exhaustive]
pub struct DocumentSourceLimitsError {
    /// Resource category whose requested limit is invalid.
    pub kind: DocumentSourceLimitKind,
    /// Requested resource ceiling.
    pub value: u64,
    /// Format-wide hard ceiling for this resource.
    pub maximum: u64,
}

/// Checked physical resource ceilings for archive-free Keynote source capture.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentSourceLimits {
    input_bytes: u64,
    entries: usize,
    entry_bytes: u64,
    aggregate_bytes: u64,
    component_bytes: usize,
}

impl DocumentSourceLimits {
    /// Build a checked source-capture resource profile.
    ///
    /// # Errors
    ///
    /// Returns an error if any ceiling is zero or exceeds its hard maximum.
    pub fn new(
        max_input_bytes: u64,
        max_entries: usize,
        max_entry_bytes: u64,
        max_aggregate_bytes: u64,
        max_component_bytes: usize,
    ) -> Result<Self, DocumentSourceLimitsError> {
        check_source_limit(
            DocumentSourceLimitKind::InputBytes,
            max_input_bytes,
            litchi_iwa_detect::Limits::HARD_MAX_INPUT_BYTES,
        )?;
        check_source_limit(
            DocumentSourceLimitKind::Entries,
            usize_u64(max_entries),
            usize_u64(litchi_iwa_detect::Limits::HARD_MAX_FILES),
        )?;
        check_source_limit(
            DocumentSourceLimitKind::EntryBytes,
            max_entry_bytes,
            litchi_iwa_detect::Limits::HARD_MAX_ENTRY_SIZE,
        )?;
        check_source_limit(
            DocumentSourceLimitKind::AggregateBytes,
            max_aggregate_bytes,
            litchi_iwa_detect::Limits::HARD_MAX_TOTAL_SIZE,
        )?;
        check_source_limit(
            DocumentSourceLimitKind::ComponentBytes,
            usize_u64(max_component_bytes),
            usize_u64(litchi_iwa_detect::Limits::HARD_MAX_IWA_STREAM_SIZE),
        )?;
        Ok(Self {
            input_bytes: max_input_bytes,
            entries: max_entries,
            entry_bytes: max_entry_bytes,
            aggregate_bytes: max_aggregate_bytes,
            component_bytes: max_component_bytes,
        })
    }

    /// Maximum encoded bytes accepted from one source.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.input_bytes
    }

    /// Maximum independently addressed source items accepted.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.entries
    }

    /// Maximum decoded bytes accepted from one source item.
    #[must_use]
    pub const fn max_entry_bytes(self) -> u64 {
        self.entry_bytes
    }

    /// Maximum aggregate decoded bytes accepted from all source items.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> u64 {
        self.aggregate_bytes
    }

    /// Maximum bytes accepted from one document payload component.
    #[must_use]
    pub const fn max_component_bytes(self) -> usize {
        self.component_bytes
    }

    fn detector_limits(self) -> litchi_iwa_detect::Result<litchi_iwa_detect::Limits> {
        litchi_iwa_detect::Limits::new(
            self.input_bytes,
            self.entries,
            self.entry_bytes,
            self.aggregate_bytes,
            self.component_bytes,
        )
    }
}

impl Default for DocumentSourceLimits {
    fn default() -> Self {
        let limits = litchi_iwa_detect::Limits::default();
        Self {
            input_bytes: limits.max_input_bytes(),
            entries: limits.max_files(),
            entry_bytes: limits.max_entry_size(),
            aggregate_bytes: limits.max_total_size(),
            component_bytes: limits.max_iwa_stream_size(),
        }
    }
}

/// Stable semantic resource category for archive-free Keynote projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DocumentSemanticLimitKind {
    /// Native IWA objects indexed for package-wide lookup.
    Objects,
    /// Semantic slides decoded from the native show tree.
    Slides,
    /// Semantic graph-reference occurrences traversed from the show root.
    References,
    /// Native text-storage objects decoded by the semantic reader.
    TextStorages,
    /// Rich-text fragment ranges retained by the semantic reader.
    TextFragments,
    /// Aggregate bytes retained from text storage and semantic identifiers.
    TextBytes,
}

impl fmt::Display for DocumentSemanticLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Objects => "objects",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
        })
    }
}

/// Invalid caller-selected semantic resource ceilings.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[error(
    "Keynote document semantic {kind} limit must be non-zero and no greater than {maximum}, got {value}"
)]
#[non_exhaustive]
pub struct DocumentSemanticLimitsError {
    /// Resource category whose requested limit is invalid.
    pub kind: DocumentSemanticLimitKind,
    /// Requested resource ceiling.
    pub value: usize,
    /// Format-wide hard ceiling for this resource.
    pub maximum: usize,
}

/// Checked semantic resource ceilings for one archive-free Keynote projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentSemanticLimits {
    objects: usize,
    slides: usize,
    references: usize,
    text_storages: usize,
    text_fragments: usize,
    text_bytes: usize,
}

impl DocumentSemanticLimits {
    /// Hard ceiling for native IWA objects indexed by one Keynote package.
    pub const MAX_OBJECTS: usize = crate::MAX_OBJECTS;
    /// Hard ceiling for semantic slides decoded from one Keynote package.
    pub const MAX_SLIDES: usize = crate::MAX_SLIDES;
    /// Hard ceiling for semantic graph-reference occurrences.
    pub const MAX_REFERENCES: usize = crate::MAX_REFERENCES;
    /// Hard ceiling for decoded text-storage objects in one Keynote package.
    pub const MAX_TEXT_STORAGES: usize = crate::MAX_TEXT_STORAGES;
    /// Hard ceiling for retained rich-text fragment ranges.
    pub const MAX_TEXT_FRAGMENTS: usize = crate::MAX_TEXT_FRAGMENTS;
    /// Hard ceiling for aggregate decoded text bytes in one Keynote package.
    pub const MAX_TEXT_BYTES: usize = crate::MAX_TEXT_BYTES;

    /// Build a checked semantic resource profile.
    ///
    /// # Errors
    ///
    /// Returns an error when any requested ceiling is zero or exceeds its
    /// format-wide hard ceiling.
    pub const fn new(
        max_objects: usize,
        max_slides: usize,
        max_references: usize,
        max_text_storages: usize,
        max_text_fragments: usize,
        max_text_bytes: usize,
    ) -> Result<Self, DocumentSemanticLimitsError> {
        if max_objects == 0 || max_objects > Self::MAX_OBJECTS {
            return Err(DocumentSemanticLimitsError {
                kind: DocumentSemanticLimitKind::Objects,
                value: max_objects,
                maximum: Self::MAX_OBJECTS,
            });
        }
        if max_slides == 0 || max_slides > Self::MAX_SLIDES {
            return Err(DocumentSemanticLimitsError {
                kind: DocumentSemanticLimitKind::Slides,
                value: max_slides,
                maximum: Self::MAX_SLIDES,
            });
        }
        if max_references == 0 || max_references > Self::MAX_REFERENCES {
            return Err(DocumentSemanticLimitsError {
                kind: DocumentSemanticLimitKind::References,
                value: max_references,
                maximum: Self::MAX_REFERENCES,
            });
        }
        if max_text_storages == 0 || max_text_storages > Self::MAX_TEXT_STORAGES {
            return Err(DocumentSemanticLimitsError {
                kind: DocumentSemanticLimitKind::TextStorages,
                value: max_text_storages,
                maximum: Self::MAX_TEXT_STORAGES,
            });
        }
        if max_text_fragments == 0 || max_text_fragments > Self::MAX_TEXT_FRAGMENTS {
            return Err(DocumentSemanticLimitsError {
                kind: DocumentSemanticLimitKind::TextFragments,
                value: max_text_fragments,
                maximum: Self::MAX_TEXT_FRAGMENTS,
            });
        }
        if max_text_bytes == 0 || max_text_bytes > Self::MAX_TEXT_BYTES {
            return Err(DocumentSemanticLimitsError {
                kind: DocumentSemanticLimitKind::TextBytes,
                value: max_text_bytes,
                maximum: Self::MAX_TEXT_BYTES,
            });
        }
        Ok(Self {
            objects: max_objects,
            slides: max_slides,
            references: max_references,
            text_storages: max_text_storages,
            text_fragments: max_text_fragments,
            text_bytes: max_text_bytes,
        })
    }

    /// Maximum number of native IWA objects indexed for package-wide lookup.
    #[must_use]
    pub const fn max_objects(self) -> usize {
        self.objects
    }

    /// Maximum number of semantic slides decoded from the native show tree.
    #[must_use]
    pub const fn max_slides(self) -> usize {
        self.slides
    }

    /// Maximum semantic graph-reference occurrences traversed.
    #[must_use]
    pub const fn max_references(self) -> usize {
        self.references
    }

    /// Maximum number of native text-storage objects decoded.
    #[must_use]
    pub const fn max_text_storages(self) -> usize {
        self.text_storages
    }

    /// Maximum rich-text fragment ranges retained.
    #[must_use]
    pub const fn max_text_fragments(self) -> usize {
        self.text_fragments
    }

    /// Maximum aggregate byte length of retained semantic text and identifiers.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.text_bytes
    }

    fn package_limits(self) -> PackageSemanticLimits {
        PackageSemanticLimits::new(
            self.objects,
            self.slides,
            self.references,
            self.text_storages,
            self.text_fragments,
            self.text_bytes,
        )
        .expect("validated document semantic limits must convert to package limits")
    }
}

impl Default for DocumentSemanticLimits {
    fn default() -> Self {
        Self {
            objects: Self::MAX_OBJECTS,
            slides: Self::MAX_SLIDES,
            references: Self::MAX_REFERENCES,
            text_storages: Self::MAX_TEXT_STORAGES,
            text_fragments: Self::MAX_TEXT_FRAGMENTS,
            text_bytes: Self::MAX_TEXT_BYTES,
        }
    }
}

/// Stable filesystem failure category for archive-free Keynote ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DocumentIoKind {
    /// The source does not exist.
    NotFound,
    /// Access to the source was denied.
    PermissionDenied,
    /// The operation conflicted with an existing resource.
    AlreadyExists,
    /// The caller supplied invalid input.
    InvalidInput,
    /// The source contained invalid data.
    InvalidData,
    /// The operation timed out.
    TimedOut,
    /// The operation was interrupted.
    Interrupted,
    /// The source ended unexpectedly.
    UnexpectedEof,
    /// Another content-free I/O category.
    Other,
}

impl fmt::Display for DocumentIoKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::NotFound => "not found",
            Self::PermissionDenied => "permission denied",
            Self::AlreadyExists => "already exists",
            Self::InvalidInput => "invalid input",
            Self::InvalidData => "invalid data",
            Self::TimedOut => "timed out",
            Self::Interrupted => "interrupted",
            Self::UnexpectedEof => "unexpected end of input",
            Self::Other => "other I/O failure",
        })
    }
}

/// Stable resource category reported by an archive-free Keynote read.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DocumentReadLimitKind {
    /// Complete captured input bytes.
    InputBytes,
    /// Independently addressed source items.
    Entries,
    /// Physical metadata bytes inspected during capture.
    MetadataBytes,
    /// Bytes contributed by one source entry or semantic payload.
    EntryBytes,
    /// Aggregate source or payload bytes.
    AggregateBytes,
    /// Decoded bytes contributed by one IWA component.
    ComponentBytes,
    /// Native IWA objects indexed for projection.
    Objects,
    /// Semantic slides decoded from the show tree.
    Slides,
    /// Semantic graph references traversed from the show root.
    References,
    /// Native text-storage objects decoded.
    TextStorages,
    /// Rich-text fragments retained.
    TextFragments,
    /// Aggregate semantic text bytes retained.
    TextBytes,
    /// Encoded or rewritten semantic payload bytes.
    PayloadBytes,
    /// Parsed semantic payload fields.
    PayloadFields,
    /// Nested semantic payload traversal depth.
    PayloadNesting,
    /// Aggregate semantic payload traversal work.
    PayloadWork,
    /// A future content-free resource category not known by this release.
    Other,
}

impl fmt::Display for DocumentReadLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Entries => "entries",
            Self::MetadataBytes => "metadata bytes",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
            Self::ComponentBytes => "component bytes",
            Self::Objects => "objects",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadFields => "payload fields",
            Self::PayloadNesting => "payload nesting",
            Self::PayloadWork => "payload work",
            Self::Other => "other resource",
        })
    }
}

/// Errors raised while publishing an archive-free Keynote document.
///
/// Display and error chains contain only closed categories and numeric bounds;
/// lower-layer archive, wire, protobuf, path, and authored-content details are
/// intentionally discarded at this boundary.
#[derive(Debug, ThisError, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum DocumentReadError {
    /// A filesystem operation failed.
    #[error("Keynote document I/O failure: {kind}")]
    Io { kind: DocumentIoKind },
    /// The recognized source belongs to a different application.
    #[error("iWork source is not a Keynote document")]
    NotKeynote,
    /// A checked resource ceiling was exceeded.
    #[error("Keynote document {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: DocumentReadLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// The source could not be captured safely.
    #[error("invalid Keynote document source")]
    InvalidSource,
    /// The captured source is not a valid Keynote document.
    #[error("invalid Keynote document format")]
    InvalidFormat,
    /// A bounded allocation failed before semantic state was published.
    #[error("Keynote document allocation failed for {amount} units")]
    Allocation {
        /// Requested elements or bytes.
        amount: usize,
    },
}

/// Deterministic measurements retained for a source-backed semantic document.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentStats {
    /// Number of semantic slides resolved from the Keynote show tree.
    pub slide_count: usize,
}

#[derive(Debug)]
struct State {
    show: Show,
    metadata: Option<litchi_core::Metadata>,
    stats: Option<DocumentStats>,
}

/// Physical and semantic resource profiles for archive-free document ingress.
///
/// The source profile belongs to format detection rather than exact package
/// preservation, so it applies equally to complete ZIP artifacts and frozen
/// app-authored package directories. The canonical properties diagnostic has
/// the independent hard ceiling [`crate::MAX_DOCUMENT_PROPERTIES_BYTES`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentReadOptions {
    source: DocumentSourceLimits,
    semantic: DocumentSemanticLimits,
}

impl DocumentReadOptions {
    /// Combine checked source-capture and semantic resource profiles.
    #[must_use]
    pub const fn new(source: DocumentSourceLimits, semantic: DocumentSemanticLimits) -> Self {
        Self { source, semantic }
    }

    /// Return the bounded source-capture profile.
    #[must_use]
    pub const fn source(self) -> DocumentSourceLimits {
        self.source
    }

    /// Return the semantic graph-decoding profile.
    #[must_use]
    pub const fn semantic(self) -> DocumentSemanticLimits {
        self.semantic
    }
}

/// An immutable, cheaply clonable semantic Keynote document snapshot.
#[derive(Debug, Clone)]
pub struct Document {
    state: Arc<State>,
}

impl Document {
    /// Open a complete Keynote package or an app-authored package directory.
    ///
    /// This constructor eagerly freezes the source and completes the bounded
    /// semantic graph projection before publishing the archive-free snapshot.
    /// A directory contributes only its IWA index and canonical
    /// `Metadata/Properties.plist`; media, previews, exact package bytes,
    /// writing, and editing are intentionally not represented by `Document`.
    /// The properties diagnostic never exceeds
    /// [`crate::MAX_DOCUMENT_PROPERTIES_BYTES`], even when the broader source
    /// entry ceiling is larger.
    /// Use [`crate::Package`] for exact complete-package preservation.
    ///
    /// # Errors
    ///
    /// Returns an error when the source is missing, unsafe, ambiguous,
    /// encrypted, belongs to another iWork application, is malformed, changes
    /// during capture, or exceeds a checked resource ceiling.
    pub fn open(path: impl AsRef<Path>) -> Result<Self, DocumentReadError> {
        Self::open_with_options(path, DocumentReadOptions::default())
    }

    /// Open an archive-free Keynote document under explicit source and
    /// semantic limits.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::open`].
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: DocumentReadOptions,
    ) -> Result<Self, DocumentReadError> {
        let source_limits = options
            .source()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source = litchi_iwa_detect::PreparedSource::__from_path_with_keynote_properties(
            path,
            source_limits,
        )
        .map_err(map_detection_error)?
        .ok_or(DocumentReadError::InvalidFormat)?;
        if source.format() != litchi_iwa_detect::Format::Keynote {
            return Err(DocumentReadError::NotKeynote);
        }
        crate::package::semantic_document_from_prepared_source(
            source,
            options.semantic().package_limits(),
        )
        .map_err(map_package_read_error)
    }

    /// Decode complete Keynote package bytes into an archive-free snapshot.
    ///
    /// The input is borrowed only for bounded capture. The returned document
    /// retains semantic show values and source diagnostics, but no package
    /// bytes, archive objects, protobuf messages, or native identifiers.
    /// Use [`crate::Package::from_bytes`] when exact package preservation or
    /// editing is required.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are missing a recognized Keynote root,
    /// violate a physical or semantic ceiling, or contain malformed semantic
    /// content.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, DocumentReadError> {
        Self::from_bytes_with_options(bytes, DocumentReadOptions::default())
    }

    /// Decode complete Keynote package bytes under explicit resource limits.
    ///
    /// The source profile is translated once at this package boundary; the
    /// semantic profile is passed through unchanged. The package handle is
    /// dropped before this method returns, so the result remains archive-free.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::from_bytes`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: DocumentReadOptions,
    ) -> Result<Self, DocumentReadError> {
        let source = options.source();
        let package_limits = crate::Limits::new(
            source.max_input_bytes(),
            source.max_entries(),
            source.max_entry_bytes(),
            source.max_aggregate_bytes(),
            source.max_component_bytes(),
        )
        .map_err(map_archive_error)?;
        let package = crate::Package::from_bytes_with_options(
            bytes,
            crate::ReadOptions::new(package_limits, options.semantic().package_limits()),
        )
        .map_err(map_package_read_error)?;
        let show = package.show().map_err(map_package_read_error)?.clone();
        let metadata = package
            .metadata()
            .map_err(map_package_read_error)?
            .ok_or(DocumentReadError::InvalidFormat)?;
        let package_stats = package.stats().map_err(map_package_read_error)?;
        let stats = DocumentStats {
            slide_count: package_stats.slide_count,
        };
        Ok(Self::from_source(show, metadata, stats))
    }

    /// Decode already-shared immutable Keynote package bytes into an
    /// archive-free snapshot without copying the input allocation.
    ///
    /// The shared source is consumed by bounded capture and semantic
    /// projection; the returned document retains no package bytes, archive
    /// objects, protobuf messages, or native identifiers. Use
    /// [`crate::Package`] when exact package preservation or editing is
    /// required.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are missing a recognized Keynote root,
    /// violate a physical or semantic ceiling, or contain malformed semantic
    /// content.
    pub fn from_shared_bytes(bytes: Arc<[u8]>) -> Result<Self, DocumentReadError> {
        Self::from_shared_bytes_with_options(bytes, DocumentReadOptions::default())
    }

    /// Decode already-shared immutable Keynote package bytes under explicit
    /// source and semantic limits.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::from_shared_bytes`].
    pub fn from_shared_bytes_with_options(
        bytes: Arc<[u8]>,
        options: DocumentReadOptions,
    ) -> Result<Self, DocumentReadError> {
        let source_limits = options
            .source()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source =
            litchi_iwa_detect::PreparedSource::from_shared_bytes_with_limits(bytes, source_limits)
                .map_err(map_detection_error)?
                .ok_or(DocumentReadError::InvalidFormat)?;
        if source.format() != litchi_iwa_detect::Format::Keynote {
            return Err(DocumentReadError::NotKeynote);
        }
        crate::package::semantic_document_from_prepared_source(
            source,
            options.semantic().package_limits(),
        )
        .map_err(map_package_read_error)
    }

    /// Create a snapshot from an already decoded semantic show.
    #[must_use]
    pub fn from_show(show: Show) -> Self {
        Self {
            state: Arc::new(State {
                show,
                metadata: None,
                stats: None,
            }),
        }
    }

    pub(crate) fn from_source(
        show: Show,
        metadata: litchi_core::Metadata,
        stats: DocumentStats,
    ) -> Self {
        Self {
            state: Arc::new(State {
                show,
                metadata: Some(metadata),
                stats: Some(stats),
            }),
        }
    }

    /// Capture another cheap handle to the same snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow the immutable semantic show.
    #[must_use]
    pub fn show(&self) -> &Show {
        &self.state.show
    }

    /// Borrow the slides without copying the snapshot.
    #[must_use]
    pub fn slides(&self) -> &[crate::Slide] {
        self.state.show.slides()
    }

    /// Extract rooted text in semantic presentation order.
    ///
    /// # Errors
    ///
    /// Returns an allocation error if the exact output buffer cannot be
    /// reserved.
    pub fn text(&self) -> Result<String, DocumentReadError> {
        crate::package::semantic_text(&self.state.show).map_err(map_package_read_error)
    }

    /// Borrow source-derived metadata.
    ///
    /// The value combines semantic Show fields with the canonical properties
    /// diagnostic when that sidecar exists. `Some` denotes a validated source
    /// origin; it does not prove the optional sidecar was present.
    ///
    /// Values built with [`Self::from_show`] have no source diagnostics and
    /// return `None`.
    #[must_use]
    pub fn metadata(&self) -> Option<&litchi_core::Metadata> {
        self.state.metadata.as_ref()
    }

    /// Return source diagnostics retained during bounded ingress.
    ///
    /// Values built with [`Self::from_show`] have no physical source and
    /// return `None`.
    #[must_use]
    pub fn stats(&self) -> Option<DocumentStats> {
        self.state.stats
    }

    /// Validate semantic invariants of the detached snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed semantic error if the retained show settings are no
    /// longer canonical.
    pub fn validate(&self) -> crate::Result<()> {
        self.state.show.settings().validate()
    }
}

fn check_source_limit(
    kind: DocumentSourceLimitKind,
    value: u64,
    maximum: u64,
) -> Result<(), DocumentSourceLimitsError> {
    if value == 0 || value > maximum {
        return Err(DocumentSourceLimitsError {
            kind,
            value,
            maximum,
        });
    }
    Ok(())
}

fn usize_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn map_package_read_error(error: PackageReadError) -> DocumentReadError {
    match error {
        PackageReadError::Io(error) => DocumentReadError::Io {
            kind: io_kind(error.kind()),
        },
        PackageReadError::Archive(error) => map_archive_error(error),
        PackageReadError::Detection(error) => map_detection_error(error),
        PackageReadError::NotKeynote => DocumentReadError::NotKeynote,
        PackageReadError::InvalidFormat(_) | PackageReadError::Decode(_) => {
            DocumentReadError::InvalidFormat
        },
        PackageReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => DocumentReadError::Limit {
            kind: semantic_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        PackageReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => DocumentReadError::Limit {
            kind: payload_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        PackageReadError::Allocation { amount, .. } => DocumentReadError::Allocation { amount },
        PackageReadError::TextStorage { .. } | PackageReadError::Metadata(_) => {
            DocumentReadError::InvalidFormat
        },
    }
}

fn map_detection_error(error: litchi_iwa_detect::Error) -> DocumentReadError {
    match error {
        litchi_iwa_detect::Error::Io(error) => DocumentReadError::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_detect::Error::IwaCore(error) => map_iwa_core_error(error),
        litchi_iwa_detect::Error::IwaCommon(error) => map_iwa_common_error(error),
        litchi_iwa_detect::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => DocumentReadError::Limit {
            kind: detection_limit_kind(kind),
            observed,
            maximum,
        },
        litchi_iwa_detect::Error::Allocation { amount } => DocumentReadError::Allocation { amount },
        litchi_iwa_detect::Error::SourceChanged => DocumentReadError::InvalidSource,
        litchi_iwa_detect::Error::InvalidFormat(_)
        | litchi_iwa_detect::Error::Archive(_)
        | litchi_iwa_detect::Error::InvalidLimits
        | litchi_iwa_detect::Error::Encrypted => DocumentReadError::InvalidFormat,
        _ => DocumentReadError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> DocumentReadError {
    match error {
        litchi_iwa_archive::Error::Io(error) => DocumentReadError::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_archive::Error::Iwa(error) => map_iwa_core_error(error),
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => DocumentReadError::Limit {
            kind: archive_limit_kind(kind),
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            DocumentReadError::Allocation { amount }
        },
        litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. } => DocumentReadError::InvalidSource,
        litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => DocumentReadError::InvalidFormat,
    }
}

fn map_iwa_core_error(error: litchi_iwa_core::Error) -> DocumentReadError {
    match error {
        litchi_iwa_core::Error::Io(error) => DocumentReadError::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => DocumentReadError::Limit {
            kind: iwa_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            DocumentReadError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Snappy { .. } => DocumentReadError::InvalidFormat,
    }
}

fn map_iwa_common_error(error: litchi_iwa_common::Error) -> DocumentReadError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => DocumentReadError::Limit {
            kind: wire_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            DocumentReadError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => DocumentReadError::InvalidFormat,
    }
}

fn io_kind(kind: std::io::ErrorKind) -> DocumentIoKind {
    match kind {
        std::io::ErrorKind::NotFound => DocumentIoKind::NotFound,
        std::io::ErrorKind::PermissionDenied => DocumentIoKind::PermissionDenied,
        std::io::ErrorKind::AlreadyExists => DocumentIoKind::AlreadyExists,
        std::io::ErrorKind::InvalidInput => DocumentIoKind::InvalidInput,
        std::io::ErrorKind::InvalidData => DocumentIoKind::InvalidData,
        std::io::ErrorKind::TimedOut => DocumentIoKind::TimedOut,
        std::io::ErrorKind::Interrupted => DocumentIoKind::Interrupted,
        std::io::ErrorKind::UnexpectedEof => DocumentIoKind::UnexpectedEof,
        _ => DocumentIoKind::Other,
    }
}

fn semantic_limit_kind(kind: crate::package::SemanticLimitKind) -> DocumentReadLimitKind {
    match kind {
        crate::package::SemanticLimitKind::Objects => DocumentReadLimitKind::Objects,
        crate::package::SemanticLimitKind::Slides => DocumentReadLimitKind::Slides,
        crate::package::SemanticLimitKind::References => DocumentReadLimitKind::References,
        crate::package::SemanticLimitKind::TextStorages => DocumentReadLimitKind::TextStorages,
        crate::package::SemanticLimitKind::TextFragments => DocumentReadLimitKind::TextFragments,
        crate::package::SemanticLimitKind::TextBytes => DocumentReadLimitKind::TextBytes,
    }
}

fn payload_limit_kind(kind: crate::package::PayloadLimitKind) -> DocumentReadLimitKind {
    match kind {
        crate::package::PayloadLimitKind::Bytes => DocumentReadLimitKind::PayloadBytes,
        crate::package::PayloadLimitKind::Fields => DocumentReadLimitKind::PayloadFields,
        crate::package::PayloadLimitKind::Nesting => DocumentReadLimitKind::PayloadNesting,
        crate::package::PayloadLimitKind::Work => DocumentReadLimitKind::PayloadWork,
    }
}

fn detection_limit_kind(kind: litchi_iwa_detect::LimitKind) -> DocumentReadLimitKind {
    match kind {
        litchi_iwa_detect::LimitKind::InputBytes => DocumentReadLimitKind::InputBytes,
        litchi_iwa_detect::LimitKind::Entries => DocumentReadLimitKind::Entries,
        litchi_iwa_detect::LimitKind::MetadataBytes => DocumentReadLimitKind::MetadataBytes,
        litchi_iwa_detect::LimitKind::EntryBytes
        | litchi_iwa_detect::LimitKind::MemberNameBytes
        | litchi_iwa_detect::LimitKind::CompressedEntryBytes => DocumentReadLimitKind::EntryBytes,
        litchi_iwa_detect::LimitKind::TotalBytes
        | litchi_iwa_detect::LimitKind::OutputBytes
        | litchi_iwa_detect::LimitKind::IwaTotalBytes => DocumentReadLimitKind::AggregateBytes,
        litchi_iwa_detect::LimitKind::IwaStreamBytes => DocumentReadLimitKind::ComponentBytes,
        litchi_iwa_detect::LimitKind::IwaObjects => DocumentReadLimitKind::Objects,
        litchi_iwa_detect::LimitKind::IwaFields => DocumentReadLimitKind::PayloadFields,
        litchi_iwa_detect::LimitKind::IwaNesting => DocumentReadLimitKind::PayloadNesting,
        litchi_iwa_detect::LimitKind::IwaWork => DocumentReadLimitKind::PayloadWork,
        _ => DocumentReadLimitKind::Other,
    }
}

fn archive_limit_kind(kind: litchi_iwa_archive::LimitKind) -> DocumentReadLimitKind {
    match kind {
        litchi_iwa_archive::LimitKind::InputBytes => DocumentReadLimitKind::InputBytes,
        litchi_iwa_archive::LimitKind::Entries => DocumentReadLimitKind::Entries,
        litchi_iwa_archive::LimitKind::MetadataBytes => DocumentReadLimitKind::MetadataBytes,
        litchi_iwa_archive::LimitKind::EntryBytes
        | litchi_iwa_archive::LimitKind::MemberNameBytes
        | litchi_iwa_archive::LimitKind::CompressedEntryBytes => DocumentReadLimitKind::EntryBytes,
        litchi_iwa_archive::LimitKind::TotalBytes
        | litchi_iwa_archive::LimitKind::OutputBytes
        | litchi_iwa_archive::LimitKind::IwaTotalBytes => DocumentReadLimitKind::AggregateBytes,
        litchi_iwa_archive::LimitKind::IwaStreamBytes => DocumentReadLimitKind::ComponentBytes,
    }
}

fn iwa_limit_kind(kind: litchi_iwa_core::LimitKind) -> DocumentReadLimitKind {
    match kind {
        litchi_iwa_core::LimitKind::ArchiveBytes => DocumentReadLimitKind::AggregateBytes,
        litchi_iwa_core::LimitKind::Objects => DocumentReadLimitKind::Objects,
        litchi_iwa_core::LimitKind::Messages
        | litchi_iwa_core::LimitKind::MessagesPerObject
        | litchi_iwa_core::LimitKind::HeaderFields
        | litchi_iwa_core::LimitKind::MetadataItems => DocumentReadLimitKind::PayloadFields,
        litchi_iwa_core::LimitKind::ObjectBytes
        | litchi_iwa_core::LimitKind::MessageBytes
        | litchi_iwa_core::LimitKind::HeaderBytes
        | litchi_iwa_core::LimitKind::HeaderMemoryBytes => DocumentReadLimitKind::PayloadBytes,
        litchi_iwa_core::LimitKind::HeaderNesting => DocumentReadLimitKind::PayloadNesting,
        litchi_iwa_core::LimitKind::SnappyChunkBytes
        | litchi_iwa_core::LimitKind::SnappyStreamBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
        | litchi_iwa_core::LimitKind::SnappyFrames => DocumentReadLimitKind::ComponentBytes,
    }
}

fn wire_limit_kind(kind: litchi_iwa_common::LimitKind) -> DocumentReadLimitKind {
    match kind {
        litchi_iwa_common::LimitKind::InputBytes | litchi_iwa_common::LimitKind::OutputBytes => {
            DocumentReadLimitKind::PayloadBytes
        },
        litchi_iwa_common::LimitKind::Fields => DocumentReadLimitKind::PayloadFields,
        litchi_iwa_common::LimitKind::Nesting => DocumentReadLimitKind::PayloadNesting,
        litchi_iwa_common::LimitKind::RewriteWork => DocumentReadLimitKind::PayloadWork,
        litchi_iwa_common::LimitKind::TableRows
        | litchi_iwa_common::LimitKind::TableColumns
        | litchi_iwa_common::LimitKind::TableCells
        | litchi_iwa_common::LimitKind::MaterializedCells => DocumentReadLimitKind::PayloadFields,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    fn native_fixture_path() -> std::path::PathBuf {
        std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/basic.key")
    }

    #[test]
    fn snapshots_are_send_sync_and_shareable() {
        assert_send_sync::<Document>();
        let document = Document::from_show(Show::builder().build());
        let snapshot = document.snapshot();
        assert_eq!(document.show(), snapshot.show());
    }

    #[test]
    fn shared_byte_ingress_matches_borrowed_ingress() {
        let bytes = std::fs::read(native_fixture_path()).expect("native Keynote fixture");
        let borrowed = Document::from_bytes(&bytes).expect("borrowed semantic ingress");
        let shared_source: Arc<[u8]> = bytes.clone().into();
        let shared = Document::from_shared_bytes(shared_source).expect("shared semantic ingress");

        assert_eq!(shared.show(), borrowed.show());
        assert_eq!(
            shared.text().expect("shared text"),
            borrowed.text().expect("borrowed text")
        );
        assert_eq!(
            shared
                .metadata()
                .map(|metadata| metadata.application.as_deref()),
            borrowed
                .metadata()
                .map(|metadata| metadata.application.as_deref())
        );
        assert_eq!(shared.stats(), borrowed.stats());
        drop(bytes);
        assert_eq!(
            shared.text().expect("shared text after source drop"),
            borrowed.text().expect("borrowed text after source drop")
        );
    }

    #[test]
    fn shared_byte_ingress_rejects_unrecognized_input() {
        let error = Document::from_shared_bytes(Arc::from([0_u8, 1, 2].as_slice()))
            .expect_err("non-package bytes must not publish a document");
        assert!(matches!(error, DocumentReadError::InvalidFormat));
    }
}
