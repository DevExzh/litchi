//! The cheaply shareable, archive-free read-only Keynote projection.
//!
//! [`Presentation`] is the deck-family semantic projection: an immutable
//! snapshot of the decoded show without package bytes, archive objects or
//! native identifiers. It is the only reader of app-authored package
//! directories; [`crate::Package`] owns exact complete-package preservation
//! and editing and projects a `Presentation` through
//! [`crate::Package::presentation`].

use std::fmt;
use std::path::Path;
use std::sync::Arc;

use thiserror::Error as ThisError;

use crate::limits::OpenOptions;
use crate::package::ReadError as PackageReadError;
use crate::show::Show;

/// Stable filesystem failure category for archive-free Keynote ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum IoKind {
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

impl fmt::Display for IoKind {
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
pub enum LimitKind {
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

impl fmt::Display for LimitKind {
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
pub enum Error {
    /// A filesystem operation failed.
    #[error("Keynote document I/O failure: {kind}")]
    Io { kind: IoKind },
    /// The recognized source belongs to a different application.
    #[error("iWork source is a {detected:?} document, not a Keynote presentation")]
    UnexpectedFormat {
        /// The application the source belongs to.
        detected: litchi_core::FileFormat,
    },
    /// A checked resource ceiling was exceeded.
    #[error("Keynote document {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
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
pub struct Stats {
    /// Number of semantic slides resolved from the Keynote show tree.
    pub slide_count: usize,
}

#[derive(Debug)]
struct State {
    show: Show,
    metadata: Option<litchi_core::Metadata>,
    stats: Option<Stats>,
}

/// An immutable, cheaply clonable semantic Keynote presentation snapshot.
#[derive(Debug, Clone)]
pub struct Presentation {
    state: Arc<State>,
}

impl Presentation {
    /// Open a complete Keynote package or an app-authored package directory.
    ///
    /// This constructor eagerly freezes the source and completes the bounded
    /// semantic graph projection before publishing the archive-free snapshot.
    /// A directory contributes only its IWA index and canonical
    /// `Metadata/Properties.plist`; media, previews, exact package bytes,
    /// writing, and editing are intentionally not represented by `Presentation`.
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
    pub fn open(path: impl AsRef<Path>) -> Result<Self, Error> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open an archive-free Keynote presentation under explicit source and
    /// semantic limits.
    ///
    /// The source limits govern format detection; their input dimension is
    /// capped at the detector's own hard ceiling.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::open`].
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions) -> Result<Self, Error> {
        let source_limits = options
            .source_limits_value()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source = litchi_iwa_detect::PreparedSource::__from_path_with_keynote_properties(
            path,
            source_limits,
        )
        .map_err(map_detection_error)?
        .ok_or(Error::InvalidFormat)?;
        if source.format() != litchi_iwa_detect::Format::Keynote {
            return Err(Error::UnexpectedFormat {
                detected: crate::package::file_format(source.format()),
            });
        }
        crate::package::presentation_from_prepared_source(source, options.semantic_limits_value())
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
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, Error> {
        Self::from_bytes_with_options(bytes, OpenOptions::default())
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
    pub fn from_bytes_with_options(bytes: &[u8], options: OpenOptions) -> Result<Self, Error> {
        let package = crate::Package::from_bytes_with_options(bytes, options)
            .map_err(map_package_read_error)?;
        let show = package.show().map_err(map_package_read_error)?.clone();
        let metadata = package
            .metadata()
            .map_err(map_package_read_error)?
            .ok_or(Error::InvalidFormat)?;
        let package_stats = package.stats().map_err(map_package_read_error)?;
        let stats = Stats {
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
    pub fn from_shared_bytes(bytes: Arc<[u8]>) -> Result<Self, Error> {
        Self::from_shared_bytes_with_options(bytes, OpenOptions::default())
    }

    /// Decode already-shared immutable Keynote package bytes under explicit
    /// source and semantic limits.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::from_shared_bytes`].
    pub fn from_shared_bytes_with_options(
        bytes: Arc<[u8]>,
        options: OpenOptions,
    ) -> Result<Self, Error> {
        let source_limits = options
            .source_limits_value()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source = litchi_iwa_detect::PreparedSource::from_shared_bytes(bytes, source_limits)
            .map_err(map_detection_error)?
            .ok_or(Error::InvalidFormat)?;
        if source.format() != litchi_iwa_detect::Format::Keynote {
            return Err(Error::UnexpectedFormat {
                detected: crate::package::file_format(source.format()),
            });
        }
        crate::package::presentation_from_prepared_source(source, options.semantic_limits_value())
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

    pub(crate) fn from_source(show: Show, metadata: litchi_core::Metadata, stats: Stats) -> Self {
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
    pub fn text(&self) -> Result<String, Error> {
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
    pub fn stats(&self) -> Option<Stats> {
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

fn usize_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn map_package_read_error(error: PackageReadError) -> Error {
    match error {
        PackageReadError::Io(error) => Error::Io {
            kind: io_kind(error.kind()),
        },
        PackageReadError::Archive(error) => map_archive_error(error),
        PackageReadError::Detection(error) => map_detection_error(error),
        PackageReadError::UnexpectedFormat { detected } => Error::UnexpectedFormat { detected },
        PackageReadError::InvalidFormat(_) | PackageReadError::Decode(_) => Error::InvalidFormat,
        PackageReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: semantic_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        PackageReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: payload_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        PackageReadError::Allocation { amount, .. } => Error::Allocation { amount },
        PackageReadError::TextStorage { .. } | PackageReadError::Metadata(_) => {
            Error::InvalidFormat
        },
    }
}

fn map_detection_error(error: litchi_iwa_detect::Error) -> Error {
    match error {
        litchi_iwa_detect::Error::Io(error) => Error::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_detect::Error::IwaCore(error) => map_iwa_core_error(error),
        litchi_iwa_detect::Error::IwaCommon(error) => map_iwa_common_error(error),
        litchi_iwa_detect::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: detection_limit_kind(kind),
            observed,
            maximum,
        },
        litchi_iwa_detect::Error::Allocation { amount } => Error::Allocation { amount },
        litchi_iwa_detect::Error::SourceChanged => Error::InvalidSource,
        litchi_iwa_detect::Error::InvalidFormat(_)
        | litchi_iwa_detect::Error::Archive(_)
        | litchi_iwa_detect::Error::InvalidLimits
        | litchi_iwa_detect::Error::Encrypted => Error::InvalidFormat,
        _ => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Io(error) => Error::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_archive::Error::Iwa(error) => map_iwa_core_error(error),
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: archive_limit_kind(kind),
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. } => Error::InvalidSource,
        _ => Error::InvalidFormat,
    }
}

fn map_iwa_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Io(error) => Error::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: iwa_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidFormat,
    }
}

fn map_iwa_common_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: wire_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidFormat,
    }
}

fn io_kind(kind: std::io::ErrorKind) -> IoKind {
    match kind {
        std::io::ErrorKind::NotFound => IoKind::NotFound,
        std::io::ErrorKind::PermissionDenied => IoKind::PermissionDenied,
        std::io::ErrorKind::AlreadyExists => IoKind::AlreadyExists,
        std::io::ErrorKind::InvalidInput => IoKind::InvalidInput,
        std::io::ErrorKind::InvalidData => IoKind::InvalidData,
        std::io::ErrorKind::TimedOut => IoKind::TimedOut,
        std::io::ErrorKind::Interrupted => IoKind::Interrupted,
        std::io::ErrorKind::UnexpectedEof => IoKind::UnexpectedEof,
        _ => IoKind::Other,
    }
}

fn semantic_limit_kind(kind: crate::limits::SemanticLimitKind) -> LimitKind {
    match kind {
        crate::limits::SemanticLimitKind::Objects => LimitKind::Objects,
        crate::limits::SemanticLimitKind::Slides => LimitKind::Slides,
        crate::limits::SemanticLimitKind::References => LimitKind::References,
        crate::limits::SemanticLimitKind::TextStorages => LimitKind::TextStorages,
        crate::limits::SemanticLimitKind::TextFragments => LimitKind::TextFragments,
        crate::limits::SemanticLimitKind::TextBytes => LimitKind::TextBytes,
    }
}

fn payload_limit_kind(kind: crate::package::PayloadLimitKind) -> LimitKind {
    match kind {
        crate::package::PayloadLimitKind::Bytes => LimitKind::PayloadBytes,
        crate::package::PayloadLimitKind::Fields => LimitKind::PayloadFields,
        crate::package::PayloadLimitKind::Nesting => LimitKind::PayloadNesting,
        crate::package::PayloadLimitKind::Work => LimitKind::PayloadWork,
    }
}

fn detection_limit_kind(kind: litchi_iwa_detect::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_detect::LimitKind::InputBytes => LimitKind::InputBytes,
        litchi_iwa_detect::LimitKind::Entries => LimitKind::Entries,
        litchi_iwa_detect::LimitKind::MetadataBytes => LimitKind::MetadataBytes,
        litchi_iwa_detect::LimitKind::EntryBytes
        | litchi_iwa_detect::LimitKind::MemberNameBytes
        | litchi_iwa_detect::LimitKind::CompressedEntryBytes => LimitKind::EntryBytes,
        litchi_iwa_detect::LimitKind::TotalBytes
        | litchi_iwa_detect::LimitKind::OutputBytes
        | litchi_iwa_detect::LimitKind::IwaTotalBytes => LimitKind::AggregateBytes,
        litchi_iwa_detect::LimitKind::IwaStreamBytes => LimitKind::ComponentBytes,
        litchi_iwa_detect::LimitKind::IwaObjects => LimitKind::Objects,
        litchi_iwa_detect::LimitKind::IwaFields => LimitKind::PayloadFields,
        litchi_iwa_detect::LimitKind::IwaNesting => LimitKind::PayloadNesting,
        litchi_iwa_detect::LimitKind::IwaWork => LimitKind::PayloadWork,
        _ => LimitKind::Other,
    }
}

fn archive_limit_kind(kind: litchi_iwa_archive::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
        litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
        litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::MetadataBytes,
        litchi_iwa_archive::LimitKind::EntryBytes
        | litchi_iwa_archive::LimitKind::MemberNameBytes
        | litchi_iwa_archive::LimitKind::CompressedEntryBytes => LimitKind::EntryBytes,
        litchi_iwa_archive::LimitKind::TotalBytes
        | litchi_iwa_archive::LimitKind::OutputBytes
        | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::AggregateBytes,
        litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::ComponentBytes,
    }
}

fn iwa_limit_kind(kind: litchi_iwa_core::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_core::LimitKind::ArchiveBytes => LimitKind::AggregateBytes,
        litchi_iwa_core::LimitKind::Objects => LimitKind::Objects,
        litchi_iwa_core::LimitKind::Messages
        | litchi_iwa_core::LimitKind::MessagesPerObject
        | litchi_iwa_core::LimitKind::HeaderFields
        | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::PayloadFields,
        litchi_iwa_core::LimitKind::ObjectBytes
        | litchi_iwa_core::LimitKind::MessageBytes
        | litchi_iwa_core::LimitKind::HeaderBytes
        | litchi_iwa_core::LimitKind::HeaderMemoryBytes => LimitKind::PayloadBytes,
        litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::PayloadNesting,
        litchi_iwa_core::LimitKind::SnappyChunkBytes
        | litchi_iwa_core::LimitKind::SnappyStreamBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
        | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::ComponentBytes,
    }
}

fn wire_limit_kind(kind: litchi_iwa_common::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_common::LimitKind::InputBytes | litchi_iwa_common::LimitKind::OutputBytes => {
            LimitKind::PayloadBytes
        },
        litchi_iwa_common::LimitKind::Fields => LimitKind::PayloadFields,
        litchi_iwa_common::LimitKind::Nesting => LimitKind::PayloadNesting,
        litchi_iwa_common::LimitKind::RewriteWork => LimitKind::PayloadWork,
        litchi_iwa_common::LimitKind::TableRows
        | litchi_iwa_common::LimitKind::TableColumns
        | litchi_iwa_common::LimitKind::TableCells
        | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::PayloadFields,
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
        assert_send_sync::<Presentation>();
        let document = Presentation::from_show(Show::builder().build());
        let snapshot = document.snapshot();
        assert_eq!(document.show(), snapshot.show());
    }

    #[test]
    fn shared_byte_ingress_matches_borrowed_ingress() {
        let bytes = std::fs::read(native_fixture_path()).expect("native Keynote fixture");
        let borrowed = Presentation::from_bytes(&bytes).expect("borrowed semantic ingress");
        let shared_source: Arc<[u8]> = bytes.clone().into();
        let shared =
            Presentation::from_shared_bytes(shared_source).expect("shared semantic ingress");

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
        let error = Presentation::from_shared_bytes(Arc::from([0_u8, 1, 2].as_slice()))
            .expect_err("non-package bytes must not publish a document");
        assert!(matches!(error, Error::InvalidFormat));
    }
}
