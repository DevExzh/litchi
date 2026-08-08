//! Bounded Apple iWork format detection.
//!
//! This leaf owns packaged and legacy bundle detection, including bounded ZIP
//! ingress, nested `Index.zip` inspection, root `Document.iwa` validation, and
//! filesystem safety checks. It does not depend on a format facade.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Detection classifiers stay beside the ingress routines they support."
)]

use litchi_iwa_archive::{
    self, ComponentCatalog, DetectionRoot, DirectoryMarkers, FrozenDirectoryBundle,
    LogicalSourceCatalog, SourceCatalog,
};
use litchi_iwa_common::wire::{WireField, parse_wire_fields};
use litchi_iwa_core::{Archive, SnappyLimits, SnappyStream};
use litchi_iwa_package::FrozenEntryStore;
use std::{
    fmt,
    fs::{self, File, Metadata, OpenOptions},
    io::{Read, Seek, SeekFrom},
    path::Path,
    sync::Arc,
    time::SystemTime,
};

const MAX_INPUT_BYTES: u64 = 1024 * 1024 * 1024;

/// Errors returned by a bounded iWork detection attempt.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// A filesystem or stream operation failed.
    Io(std::io::Error),
    /// The neutral IWA substrate rejected a stream or archive.
    IwaCore(litchi_iwa_core::Error),
    /// Shared wire validation rejected a protobuf payload.
    IwaCommon(litchi_iwa_common::Error),
    /// The input is not a valid or unambiguous iWork format.
    InvalidFormat(String),
    /// ZIP ingress or bundle structure was rejected.
    Archive(String),
    /// A physical package resource ceiling was exceeded.
    LimitExceeded {
        /// Content-free physical resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A caller supplied an invalid physical resource profile.
    InvalidLimits,
    /// A bounded physical allocation failed before a source was published.
    Allocation {
        /// Elements or bytes requested by the failed allocation.
        amount: usize,
    },
    /// The package uses an encrypted iWork container marker.
    Encrypted,
    /// A positional or directory source changed while it was being captured.
    SourceChanged,
}

/// Physical resource category reported by [`Error::LimitExceeded`].
///
/// This detector-owned vocabulary prevents callers from depending on the ZIP,
/// Snappy, or neutral IWA implementations used below the detection boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete packaged input bytes.
    InputBytes,
    /// Complete bytes produced by package reassembly.
    OutputBytes,
    /// Number of packaged members.
    Entries,
    /// Bytes in one packaged member name.
    MemberNameBytes,
    /// Aggregate packaged-header metadata bytes.
    MetadataBytes,
    /// Compressed bytes in one packaged member.
    CompressedEntryBytes,
    /// Expanded bytes in one packaged member.
    EntryBytes,
    /// Aggregate expanded package bytes.
    TotalBytes,
    /// Decoded bytes in one IWA component.
    IwaStreamBytes,
    /// Aggregate decoded bytes across IWA components.
    IwaTotalBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "package entries",
            Self::MemberNameBytes => "member name bytes",
            Self::MetadataBytes => "package metadata bytes",
            Self::CompressedEntryBytes => "compressed entry bytes",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "expanded package bytes",
            Self::IwaStreamBytes => "decoded IWA component bytes",
            Self::IwaTotalBytes => "aggregate decoded IWA bytes",
        })
    }
}

/// Result type for iWork detection.
pub type Result<T> = std::result::Result<T, Error>;

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "I/O error: {error}"),
            Self::IwaCore(error) => error.fmt(formatter),
            Self::IwaCommon(error) => error.fmt(formatter),
            Self::InvalidFormat(message) => write!(formatter, "Invalid IWA format: {message}"),
            Self::Archive(message) => write!(formatter, "Archive parsing error: {message}"),
            Self::LimitExceeded {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "iWork detection {kind} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::InvalidLimits => formatter.write_str("invalid iWork detection limits"),
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "iWork detection allocation failed for {amount} units"
                )
            },
            Self::Encrypted => {
                formatter.write_str("password-protected iWork documents are not supported")
            },
            Self::SourceChanged => {
                formatter.write_str("iWork source changed while it was being captured")
            },
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::IwaCore(error) => Some(error),
            Self::IwaCommon(error) => Some(error),
            Self::InvalidFormat(_)
            | Self::Archive(_)
            | Self::LimitExceeded { .. }
            | Self::InvalidLimits
            | Self::Allocation { .. }
            | Self::Encrypted
            | Self::SourceChanged => None,
        }
    }
}

impl From<std::io::Error> for Error {
    fn from(error: std::io::Error) -> Self {
        Self::Io(error)
    }
}

impl From<litchi_iwa_core::Error> for Error {
    fn from(error: litchi_iwa_core::Error) -> Self {
        match error {
            litchi_iwa_core::Error::InvalidLimits { .. } => Self::InvalidLimits,
            litchi_iwa_core::Error::Limit {
                observed, maximum, ..
            } => Self::LimitExceeded {
                kind: LimitKind::IwaStreamBytes,
                observed: usize_u64(observed),
                maximum: usize_u64(maximum),
            },
            litchi_iwa_core::Error::Io(io_error) => Self::Io(io_error),
            litchi_iwa_core::Error::Allocation { requested, .. } => {
                Self::Allocation { amount: requested }
            },
            other @ (litchi_iwa_core::Error::InvalidArchive { .. }
            | litchi_iwa_core::Error::HeaderCodec { .. }
            | litchi_iwa_core::Error::Snappy { .. }) => Self::IwaCore(other),
        }
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self::IwaCommon(error)
    }
}
/// The application family of a detected iWork document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Format {
    /// Apple Pages.
    Pages,
    /// Apple Keynote.
    Keynote,
    /// Apple Numbers.
    Numbers,
}

/// Opaque, single-use iWork source prepared for one concrete format owner.
///
/// Construction snapshots and validates the physical package, decodes its IWA
/// components, and classifies the application from that same immutable state.
/// Passing this value to a focused Pages, Keynote, or Numbers adapter therefore
/// avoids a second ZIP traversal or Snappy/IWA decode. The physical catalog is
/// deliberately not exposed by the ordinary detection API.
pub struct PreparedSource {
    backing: PreparedBacking,
    format: Format,
    limits: Limits,
}

enum PreparedBacking {
    Package(SourceCatalog),
    Semantic {
        components: Arc<ComponentCatalog>,
        limits: litchi_iwa_archive::Limits,
    },
}

impl fmt::Debug for PreparedSource {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedSource")
            .field("format", &self.format)
            .field("limits", &self.limits)
            .finish_non_exhaustive()
    }
}

impl PreparedSource {
    /// Prepare borrowed package bytes under the default detection profile.
    ///
    /// The bytes are copied once into immutable shared ownership. Use
    /// [`Self::from_shared_bytes`] when the caller already owns an `Arc`.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is malformed, ambiguous, encrypted,
    /// or exceeds a physical resource ceiling. Non-ZIP and unrecognized iWork
    /// inputs return `Ok(None)`.
    pub fn from_bytes(value: &[u8]) -> Result<Option<Self>> {
        Self::from_bytes_with_limits(value, Limits::default())
    }

    /// Prepare borrowed package bytes under explicit detection limits.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_bytes`].
    pub fn from_bytes_with_limits(value: &[u8], limits: Limits) -> Result<Option<Self>> {
        if !check_prepared_candidate(value, limits)? {
            return Ok(None);
        }
        let catalog = SourceCatalog::from_bytes_with_limits(value, archive_limits(limits)?)
            .map_err(map_archive_error)?;
        Self::from_catalog(catalog, limits)
    }

    /// Prepare already-owned immutable package bytes without copying them.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is malformed, ambiguous, encrypted,
    /// or exceeds a physical resource ceiling. Non-ZIP and unrecognized iWork
    /// inputs return `Ok(None)`.
    pub fn from_shared_bytes(value: Arc<[u8]>) -> Result<Option<Self>> {
        Self::from_shared_bytes_with_limits(value, Limits::default())
    }

    /// Prepare an immutable package under explicit physical limits.
    ///
    /// The same checked profile is retained by the underlying catalog so a
    /// selected format owner cannot accidentally reinterpret already-validated
    /// bytes under a different physical policy.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is malformed, ambiguous, encrypted,
    /// or exceeds a physical resource ceiling. Non-ZIP and unrecognized iWork
    /// inputs return `Ok(None)`.
    pub fn from_shared_bytes_with_limits(value: Arc<[u8]>, limits: Limits) -> Result<Option<Self>> {
        if !check_prepared_candidate(&value, limits)? {
            return Ok(None);
        }
        let catalog = SourceCatalog::from_shared_bytes_with_limits(value, archive_limits(limits)?)
            .map_err(map_archive_error)?;
        Self::from_catalog(catalog, limits)
    }

    /// Prepare an immutable, already-normalized logical entry snapshot.
    ///
    /// This explicitly unstable integration route performs no synthetic ZIP
    /// serialization and never claims exact-package provenance. The frozen
    /// entry store remains alive through complete admission and application
    /// classification, then is released before the selected format decoder
    /// receives the component-only semantic snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when any entry is unsafe, ambiguous, encrypted,
    /// malformed, or exceeds a physical resource ceiling. An entry set with
    /// no recognized canonical application root returns `Ok(None)`.
    #[doc(hidden)]
    pub fn from_frozen_entries(value: FrozenEntryStore) -> Result<Option<Self>> {
        Self::from_frozen_entries_with_limits(value, Limits::default())
    }

    /// Prepare immutable logical entries under explicit physical limits.
    ///
    /// The input-byte ceiling is not reinterpreted as an expanded logical
    /// payload ceiling; per-entry and aggregate expanded ceilings govern the
    /// already-decoded member payloads.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_frozen_entries`].
    #[doc(hidden)]
    pub fn from_frozen_entries_with_limits(
        value: FrozenEntryStore,
        limits: Limits,
    ) -> Result<Option<Self>> {
        let logical =
            LogicalSourceCatalog::from_frozen_entries_with_limits(value, archive_limits(limits)?)
                .map_err(map_archive_error)?;
        let Some(format) = component_catalog(logical.components())? else {
            return Ok(None);
        };
        let archive_limits = logical.limits();
        let components = Arc::new(logical.into_components());
        Ok(Some(Self {
            backing: PreparedBacking::Semantic {
                components,
                limits: archive_limits,
            },
            format,
            limits,
        }))
    }

    /// Prepare a packaged file or app-authored directory bundle.
    ///
    /// Regular files are opened once, read through a bounded stable file
    /// handle, and classified from the retained package catalog. Directories
    /// are frozen through the archive-owned index adapter and classified from
    /// those same retained components. Symbolic links and special nodes are
    /// rejected. No semantic adapter reopens the path.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is missing, unsafe, malformed, changes
    /// during capture, or exceeds a physical resource ceiling. An accessible
    /// regular file or directory that is not recognized as iWork returns
    /// `Ok(None)`.
    pub fn from_path(value: impl AsRef<Path>) -> Result<Option<Self>> {
        Self::from_path_with_limits(value, Limits::default())
    }

    /// Prepare a packaged file or app-authored directory bundle under
    /// explicit physical limits.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_path`].
    pub fn from_path_with_limits(value: impl AsRef<Path>, limits: Limits) -> Result<Option<Self>> {
        let path = value.as_ref();
        match kind(path)? {
            Kind::File => {
                let source = read_stable_package_file(path, limits)?;
                Self::from_shared_bytes_with_limits(source, limits)
            },
            Kind::Dir => {
                let directory =
                    FrozenDirectoryBundle::open_with_limits(path, archive_limits(limits)?)
                        .map_err(map_archive_error)?;
                Self::from_directory(directory, limits)
            },
            Kind::Missing => Err(Error::Io(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                "iWork source path does not exist",
            ))),
        }
    }

    fn from_catalog(catalog: SourceCatalog, limits: Limits) -> Result<Option<Self>> {
        let Some(format) = component_catalog(catalog.components())? else {
            return Ok(None);
        };
        Ok(Some(Self {
            backing: PreparedBacking::Package(catalog),
            format,
            limits,
        }))
    }

    fn from_directory(directory: FrozenDirectoryBundle, limits: Limits) -> Result<Option<Self>> {
        let Some(format) = component_catalog(directory.components())? else {
            if marker_outcome(directory.markers()) != Outcome::None {
                return Err(Error::InvalidFormat(
                    "iWork directory marker has no canonical application root".to_owned(),
                ));
            }
            return Ok(None);
        };
        match marker_outcome(directory.markers()) {
            Outcome::None => {},
            Outcome::Found(marker) if marker == format => {},
            Outcome::Found(_) => {
                return Err(Error::InvalidFormat(
                    "iWork directory marker conflicts with the canonical application root"
                        .to_owned(),
                ));
            },
            Outcome::Conflict => {
                return Err(Error::InvalidFormat(
                    "iWork directory contains conflicting application markers".to_owned(),
                ));
            },
        }
        Ok(Some(Self {
            backing: PreparedBacking::Semantic {
                limits: directory.limits(),
                components: directory.into_components(),
            },
            format,
            limits,
        }))
    }

    /// Return the application selected from the retained component snapshot.
    #[must_use]
    pub const fn format(&self) -> Format {
        self.format
    }

    /// Return the checked detection profile used to prepare this source.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Consume a packaged source into its physical catalog for a focused
    /// preserve-mode adapter.
    ///
    /// This is an explicitly unstable cross-crate handoff. Application code
    /// should pass `PreparedSource` to a format owner instead of depending on
    /// archive vocabulary directly.
    #[doc(hidden)]
    #[must_use]
    pub fn __into_source_catalog(self) -> Option<SourceCatalog> {
        match self.backing {
            PreparedBacking::Package(catalog) => Some(catalog),
            PreparedBacking::Semantic { .. } => None,
        }
    }

    /// Consume any prepared source into its semantic component snapshot and
    /// the physical limits that authorized it.
    ///
    /// This unstable handoff intentionally erases exact-package provenance.
    /// It is appropriate only for archive-free semantic projections.
    #[doc(hidden)]
    #[must_use]
    pub fn __into_components(self) -> (Arc<ComponentCatalog>, litchi_iwa_archive::Limits) {
        match self.backing {
            PreparedBacking::Package(catalog) => {
                let limits = catalog.limits();
                (Arc::new(catalog.into_components()), limits)
            },
            PreparedBacking::Semantic { components, limits } => (components, limits),
        }
    }
}

fn marker_outcome(markers: DirectoryMarkers) -> Outcome {
    classify(markers.pages(), markers.keynote(), markers.numbers())
}

fn read_stable_package_file(path: &Path, limits: Limits) -> Result<Arc<[u8]>> {
    let _checked = archive_limits(limits)?;
    let mut options = OpenOptions::new();
    options.read(true);
    #[cfg(unix)]
    {
        use std::os::unix::fs::OpenOptionsExt;
        options.custom_flags(libc::O_CLOEXEC | libc::O_NOFOLLOW | libc::O_NONBLOCK);
    }
    let mut file = options.open(path)?;
    let before = FileSnapshot::from_metadata(&file.metadata()?);
    if !before.is_regular_file() {
        return Err(Error::InvalidFormat(
            "iWork source path is not a regular file".to_owned(),
        ));
    }
    if before.len > limits.max_input_bytes {
        return Err(Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed: before.len,
            maximum: limits.max_input_bytes,
        });
    }

    let capacity =
        usize::try_from(before.len).map_err(|_error| Error::Allocation { amount: usize::MAX })?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(capacity)
        .map_err(|_error| Error::Allocation { amount: capacity })?;
    file.by_ref()
        .take(limits.max_input_bytes.saturating_add(1))
        .read_to_end(&mut bytes)?;
    let observed = u64::try_from(bytes.len()).unwrap_or(u64::MAX);
    if observed > limits.max_input_bytes {
        return Err(Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed,
            maximum: limits.max_input_bytes,
        });
    }

    let after = FileSnapshot::from_metadata(&file.metadata()?);
    let path_after = match fs::symlink_metadata(path) {
        Ok(metadata) => FileSnapshot::from_metadata(&metadata),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            return Err(Error::SourceChanged);
        },
        Err(error) => return Err(Error::Io(error)),
    };
    if before != after || after != path_after || observed != before.len {
        return Err(Error::SourceChanged);
    }
    Ok(bytes.into())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct FileSnapshot {
    len: u64,
    modified: Option<SystemTime>,
    is_file: bool,
    #[cfg(unix)]
    device: u64,
    #[cfg(unix)]
    inode: u64,
    #[cfg(unix)]
    mode: u32,
    #[cfg(unix)]
    modified_seconds: i64,
    #[cfg(unix)]
    modified_nanoseconds: i64,
    #[cfg(unix)]
    changed_seconds: i64,
    #[cfg(unix)]
    changed_nanoseconds: i64,
}

impl FileSnapshot {
    fn from_metadata(metadata: &Metadata) -> Self {
        #[cfg(unix)]
        use std::os::unix::fs::MetadataExt;

        Self {
            len: metadata.len(),
            modified: metadata.modified().ok(),
            is_file: metadata.is_file(),
            #[cfg(unix)]
            device: metadata.dev(),
            #[cfg(unix)]
            inode: metadata.ino(),
            #[cfg(unix)]
            mode: metadata.mode(),
            #[cfg(unix)]
            modified_seconds: metadata.mtime(),
            #[cfg(unix)]
            modified_nanoseconds: metadata.mtime_nsec(),
            #[cfg(unix)]
            changed_seconds: metadata.ctime(),
            #[cfg(unix)]
            changed_nanoseconds: metadata.ctime_nsec(),
        }
    }

    const fn is_regular_file(self) -> bool {
        self.is_file
    }
}

fn check_prepared_candidate(value: &[u8], limits: Limits) -> Result<bool> {
    let input_size = u64::try_from(value.len()).unwrap_or(u64::MAX);
    if input_size > limits.max_input_bytes {
        return Err(Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed: input_size,
            maximum: limits.max_input_bytes,
        });
    }
    Ok(is_zip_signature(value))
}

/// Resource ceilings for one iWork detection attempt.
///
/// The defaults are conservative enough for untrusted input while allowing
/// ordinary media-heavy documents. Callers may tighten any ceiling, but the
/// checked constructor never permits a limit above the format-wide hard
/// ceiling. Detection remains fail-closed when a limit is exceeded.
#[allow(
    clippy::struct_field_names,
    reason = "The max_* vocabulary makes each detection ceiling self-documenting."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: u64,
    max_files: usize,
    max_entry_size: u64,
    max_total_size: u64,
    max_iwa_stream_size: usize,
}

impl Limits {
    /// Maximum input size accepted by the default safety profile.
    pub const HARD_MAX_INPUT_BYTES: u64 = MAX_INPUT_BYTES;
    /// Maximum ZIP member count accepted by the default safety profile.
    pub const HARD_MAX_FILES: usize = litchi_iwa_archive::Limits::MAX_ENTRIES;
    /// Maximum uncompressed ZIP member size accepted by the default profile.
    pub const HARD_MAX_ENTRY_SIZE: u64 = litchi_iwa_archive::Limits::MAX_ENTRY_BYTES;
    /// Maximum aggregate uncompressed ZIP size accepted by the default profile.
    pub const HARD_MAX_TOTAL_SIZE: u64 = litchi_iwa_archive::Limits::MAX_TOTAL_BYTES;
    /// Maximum decompressed size of one IWA component.
    pub const HARD_MAX_IWA_STREAM_SIZE: usize = SnappyStream::MAX_DECOMPRESSED_STREAM;

    /// Construct checked detection ceilings.
    ///
    /// # Errors
    ///
    /// Returns an error when a ceiling is zero or exceeds its format-wide
    /// hard maximum.
    pub fn new(
        max_input_bytes: u64,
        max_files: usize,
        max_entry_size: u64,
        max_total_size: u64,
        max_iwa_stream_size: usize,
    ) -> Result<Self> {
        if max_input_bytes == 0
            || max_files == 0
            || max_entry_size == 0
            || max_total_size == 0
            || max_iwa_stream_size == 0
        {
            return Err(Error::InvalidLimits);
        }
        if max_input_bytes > Self::HARD_MAX_INPUT_BYTES {
            return Err(Error::InvalidLimits);
        }
        if max_files > Self::HARD_MAX_FILES {
            return Err(Error::InvalidLimits);
        }
        if max_entry_size > Self::HARD_MAX_ENTRY_SIZE {
            return Err(Error::InvalidLimits);
        }
        if max_total_size > Self::HARD_MAX_TOTAL_SIZE {
            return Err(Error::InvalidLimits);
        }
        if max_iwa_stream_size > Self::HARD_MAX_IWA_STREAM_SIZE {
            return Err(Error::InvalidLimits);
        }

        Ok(Self {
            max_input_bytes,
            max_files,
            max_entry_size,
            max_total_size,
            max_iwa_stream_size,
        })
    }

    /// Maximum complete input size accepted by this profile.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum number of ZIP members indexed by one probe.
    #[must_use]
    pub const fn max_files(self) -> usize {
        self.max_files
    }

    /// Maximum declared uncompressed size of one ZIP member.
    #[must_use]
    pub const fn max_entry_size(self) -> u64 {
        self.max_entry_size
    }

    /// Maximum aggregate declared uncompressed ZIP size.
    #[must_use]
    pub const fn max_total_size(self) -> u64 {
        self.max_total_size
    }

    /// Maximum decompressed size of one IWA component.
    #[must_use]
    pub const fn max_iwa_stream_size(self) -> usize {
        self.max_iwa_stream_size
    }

    fn snappy_limits(self) -> Result<SnappyLimits> {
        Ok(SnappyLimits::new(
            self.max_iwa_stream_size
                .min(SnappyStream::MAX_UNCOMPRESSED_CHUNK),
            self.max_iwa_stream_size,
        )?)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: MAX_INPUT_BYTES,
            max_files: litchi_iwa_archive::Limits::MAX_ENTRIES,
            max_entry_size: litchi_iwa_archive::Limits::MAX_ENTRY_BYTES,
            max_total_size: litchi_iwa_archive::Limits::MAX_TOTAL_BYTES,
            max_iwa_stream_size: SnappyStream::MAX_DECOMPRESSED_STREAM,
        }
    }
}

/// Detect an iWork application from complete packaged bytes.
///
/// ZIP and Snappy metadata are validated under explicit file-count and size
/// limits. A package with conflicting application-root evidence is reported as
/// a typed format error; an unrelated or unrecognized byte slice returns
/// `Ok(None)`.
///
/// # Errors
///
/// Returns a typed error when the package is malformed, ambiguous, encrypted,
/// or exceeds a configured resource ceiling.
pub fn bytes(value: &[u8]) -> Result<Option<Format>> {
    bytes_with_limits(value, Limits::default())
}

/// Detect an iWork application using caller-selected resource ceilings.
///
/// # Errors
///
/// Returns a typed error when the package is malformed, ambiguous, encrypted,
/// or exceeds the supplied resource ceilings.
pub fn bytes_with_limits(value: &[u8], limits: Limits) -> Result<Option<Format>> {
    let input_size = u64::try_from(value.len()).unwrap_or(u64::MAX);
    if input_size > limits.max_input_bytes {
        return Err(Error::LimitExceeded {
            kind: LimitKind::InputBytes,
            observed: input_size,
            maximum: limits.max_input_bytes,
        });
    }
    if !is_zip_signature(value) {
        return Ok(None);
    }
    let root = litchi_iwa_archive::inspect_detection_root(value, archive_limits(limits)?)
        .map_err(map_archive_error)?;
    classify_root(&root)
}

/// Detect an iWork application from an already-parsed component catalog.
///
/// This route is intended for source-owning format coordinators. It performs
/// no ZIP, Snappy, or IWA work and classifies the same immutable components
/// that the caller will subsequently project. The catalog's ingress limits
/// therefore remain authoritative instead of being replaced by detector
/// defaults during a second parse.
///
/// # Errors
///
/// Returns a typed error when canonical root evidence is ambiguous,
/// unrecognized, or conflicts with Keynote component markers. A catalog with
/// no canonical application root returns `Ok(None)`.
pub fn component_catalog(catalog: &ComponentCatalog) -> Result<Option<Format>> {
    if catalog.is_empty() {
        return Ok(None);
    }

    let mut marks = Marks::default();
    let mut root_archive = None;
    for component in catalog.iter() {
        let Some(name) = index_name(component.name()) else {
            continue;
        };
        marks.see_index(name);
        if name == "Document.iwa" && root_archive.replace(component.archive()).is_some() {
            return Err(Error::InvalidFormat(
                "iWork package contains multiple Document.iwa components".to_owned(),
            ));
        }
    }

    if !marks.iwa {
        return Ok(None);
    }
    let Some(resolved_root) = root_archive else {
        return Ok(None);
    };
    let Some(format) = root_format_archive(resolved_root)? else {
        return Err(Error::InvalidFormat(
            "Document.iwa has no recognized iWork application root".to_owned(),
        ));
    };
    if marks.accepts(format) {
        Ok(Some(format))
    } else {
        Err(Error::InvalidFormat(
            "iWork component markers conflict with the Document.iwa application root".to_owned(),
        ))
    }
}

fn archive_limits(limits: Limits) -> Result<litchi_iwa_archive::Limits> {
    litchi_iwa_archive::Limits::new(
        limits.max_input_bytes,
        limits.max_files,
        limits.max_entry_size,
        limits.max_total_size,
        limits.max_iwa_stream_size,
    )
    .map_err(map_archive_error)
}

fn map_archive_error(archive_error: litchi_iwa_archive::Error) -> Error {
    match archive_error {
        litchi_iwa_archive::Error::Io(io_error) => Error::Io(io_error),
        litchi_iwa_archive::Error::Iwa(iwa_error) => Error::from(iwa_error),
        litchi_iwa_archive::Error::Encrypted => Error::Encrypted,
        litchi_iwa_archive::Error::InvalidLimits(_) => Error::InvalidLimits,
        litchi_iwa_archive::Error::Zip { message }
        | litchi_iwa_archive::Error::InvalidBundle(message) => {
            Error::Archive(format!("iWork archive ingress: {message}"))
        },
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: map_archive_limit(kind),
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. } => Error::SourceChanged,
        litchi_iwa_archive::Error::Reassembly(message) => {
            Error::Archive(format!("iWork archive reassembly: {message}"))
        },
    }
}

const fn map_archive_limit(kind: litchi_iwa_archive::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
        litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
        litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
        litchi_iwa_archive::LimitKind::MemberNameBytes => LimitKind::MemberNameBytes,
        litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::MetadataBytes,
        litchi_iwa_archive::LimitKind::CompressedEntryBytes => LimitKind::CompressedEntryBytes,
        litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
        litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalBytes,
        litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::IwaStreamBytes,
        litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::IwaTotalBytes,
    }
}

fn usize_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn classify_root(root: &DetectionRoot) -> Result<Option<Format>> {
    if !root.has_iwa_components() {
        return Ok(None);
    }
    let marks = Marks {
        iwa: true,
        keynote: root.has_keynote_components(),
    };
    let Some(document) = root.document() else {
        return Ok(None);
    };
    let Some(format) = root_format_archive(document)? else {
        return Err(Error::InvalidFormat(
            "Document.iwa has no recognized iWork application root".to_owned(),
        ));
    };
    if marks.accepts(format) {
        Ok(Some(format))
    } else {
        Err(Error::InvalidFormat(
            "iWork component markers conflict with the Document.iwa application root".to_owned(),
        ))
    }
}

fn is_zip_signature(value: &[u8]) -> bool {
    value.starts_with(b"PK\x03\x04")
        || value.starts_with(b"PK\x05\x06")
        || value.starts_with(b"PK\x07\x08")
}

/// Detect the owning iWork application from the root `DocumentArchive` payload.
///
/// Message type identifiers overlap between Pages, Numbers, and Keynote, so they
/// cannot reliably identify an application. The root protobuf schemas have
/// stable, application-specific required message shapes: Pages uses its shared
/// document at field 15, Numbers uses references at fields 4/5/6 plus its shared
/// document at field 8, and Keynote uses a reference at field 2 plus its shared
/// document at field 3. Malformed or multiply matching payloads fail closed.
#[must_use]
pub fn detect_application_from_document(payload: &[u8]) -> Option<Format> {
    let fields = wire_fields(payload)?;
    let pages = unique_field(payload, &fields, 15, 2).is_some_and(valid_shared_document);
    let numbers = [4, 5, 6]
        .into_iter()
        .all(|field| unique_field(payload, &fields, field, 2).is_some_and(valid_reference))
        && unique_field(payload, &fields, 8, 2).is_some_and(valid_shared_document);
    let keynote = unique_field(payload, &fields, 2, 2).is_some_and(valid_reference)
        && unique_field(payload, &fields, 3, 2).is_some_and(valid_shared_document);

    match (pages, numbers, keynote) {
        (true, false, false) => Some(Format::Pages),
        (false, true, false) => Some(Format::Numbers),
        (false, false, true) => Some(Format::Keynote),
        _ => None,
    }
}

fn wire_fields(payload: &[u8]) -> Option<Vec<WireField>> {
    parse_wire_fields(payload).ok()
}

fn unique_field<'a>(
    payload: &'a [u8],
    fields: &[WireField],
    number: u32,
    wire_type: u8,
) -> Option<&'a [u8]> {
    let mut matches = fields.iter().filter(|field| field.number() == number);
    let field = matches.next()?;
    if matches.next().is_some() || field.wire_type() != wire_type {
        return None;
    }
    field.checked_payload(payload).ok()
}

fn valid_reference(payload: &[u8]) -> bool {
    wire_fields(payload)
        .and_then(|fields| unique_field(payload, &fields, 1, 0))
        .is_some()
}

fn valid_shared_document(payload: &[u8]) -> bool {
    wire_fields(payload)
        .and_then(|fields| unique_field(payload, &fields, 1, 2))
        .and_then(wire_fields)
        .is_some()
}

fn root_format(data: &[u8], limits: Limits) -> Result<Option<Format>> {
    let stream = SnappyStream::decompress_with_limits(data, limits.snappy_limits()?)?;
    let archive = Archive::parse(stream.as_bytes())?;
    root_format_archive(&archive)
}

fn root_format_archive(archive: &Archive) -> Result<Option<Format>> {
    let mut detected = None;

    for application in archive
        .objects
        .iter()
        .filter(|object| object.archive_info.identifier == Some(1))
        .flat_map(|object| &object.messages)
        .filter_map(|message| detect_application_from_document(&message.data))
    {
        let format = application;
        if detected.is_some() {
            return Err(Error::InvalidFormat(
                "Document.iwa contains multiple application roots".to_owned(),
            ));
        }
        detected = Some(format);
    }

    Ok(detected)
}

/// Detect an iWork application from a seekable stream.
///
/// Detection starts at byte zero and restores the caller's original cursor on
/// every path. Streams larger than the selected input ceiling are rejected
/// without being read.
///
/// # Errors
///
/// Returns a typed error when the stream cannot be inspected, restored, or
/// contains malformed or over-limit iWork input.
pub fn reader<R: Read + Seek>(value: &mut R) -> Result<Option<Format>> {
    reader_with_limits(value, Limits::default())
}

/// Detect an iWork application from a seekable stream under explicit limits.
///
/// # Errors
///
/// Returns a typed error when the stream cannot be inspected, restored, or
/// contains malformed or over-limit iWork input.
pub fn reader_with_limits<R: Read + Seek>(value: &mut R, limits: Limits) -> Result<Option<Format>> {
    let original = value.stream_position()?;
    let detected = (|| {
        let input_length = value.seek(SeekFrom::End(0))?;
        if input_length > limits.max_input_bytes {
            return Err(Error::InvalidFormat(format!(
                "iWork detection input is {input_length} bytes, exceeding the {} byte limit",
                limits.max_input_bytes
            )));
        }

        let input_size = usize::try_from(input_length).map_err(|error| {
            Error::InvalidFormat(format!("iWork input length does not fit usize: {error}"))
        })?;
        let mut data = Vec::new();
        data.try_reserve_exact(input_size).map_err(|error| {
            Error::InvalidFormat(format!(
                "unable to reserve iWork detection input buffer: {error}"
            ))
        })?;
        data.resize(input_size, 0);

        value.seek(SeekFrom::Start(0))?;
        value.read_exact(&mut data)?;
        let mut extra = [0];
        if value.read(&mut extra)? != 0 {
            return Err(Error::InvalidFormat(
                "iWork detection source changed while it was being read".to_owned(),
            ));
        }
        bytes_with_limits(&data, limits)
    })();
    value.seek(SeekFrom::Start(original))?;
    detected
}

/// Detect a packaged iWork file or a legacy directory bundle.
///
/// Symbolic links, conflicting markers, malformed Index.zip archives, and
/// directory traversal errors are typed errors.
///
/// # Errors
///
/// Returns a typed error when the path is inaccessible, unsafe, malformed, or
/// exceeds a configured resource ceiling.
pub fn path(value: impl AsRef<Path>) -> Result<Option<Format>> {
    path_with_limits(value, Limits::default())
}

/// Detect a packaged file or legacy directory bundle under explicit limits.
///
/// # Errors
///
/// Returns a typed error when the path is inaccessible, unsafe, malformed, or
/// exceeds the supplied resource ceilings.
pub fn path_with_limits(value: impl AsRef<Path>, limits: Limits) -> Result<Option<Format>> {
    let path = value.as_ref();
    match kind(path)? {
        Kind::File => reader_with_limits(&mut File::open(path)?, limits),
        Kind::Dir => directory(path, limits),
        Kind::Missing => Ok(None),
    }
}

fn directory(root: &Path, limits: Limits) -> Result<Option<Format>> {
    let mut evidence = classify(
        marker(root, "index.xml")?,
        marker(root, "index.apxl")?,
        marker(root, "index.numbers")?,
    );
    if evidence == Outcome::Conflict {
        return Err(Error::InvalidFormat(
            "iWork bundle contains conflicting legacy application markers".to_owned(),
        ));
    }

    let index_zip = root.join("Index.zip");
    evidence = evidence.merge(match kind(&index_zip)? {
        Kind::File => match reader_with_limits(&mut File::open(&index_zip)?, limits)? {
            Some(format) => Outcome::Found(format),
            None => Outcome::Conflict,
        },
        Kind::Dir => Outcome::Conflict,
        Kind::Missing => Outcome::None,
    });
    if evidence == Outcome::Conflict {
        return Err(Error::InvalidFormat(
            "iWork bundle contains an invalid or conflicting Index.zip".to_owned(),
        ));
    }

    let index = root.join("Index");
    evidence = evidence.merge(match kind(&index)? {
        Kind::Dir => directory_outcome(&index, limits)?,
        Kind::File => Outcome::Conflict,
        Kind::Missing => Outcome::None,
    });

    match evidence {
        Outcome::Found(format) => Ok(Some(format)),
        Outcome::None => Ok(None),
        Outcome::Conflict => Err(Error::InvalidFormat(
            "iWork bundle contains conflicting application evidence".to_owned(),
        )),
    }
}

fn directory_outcome(index: &Path, limits: Limits) -> Result<Outcome> {
    let mut marks = Marks::default();
    let mut document = None;
    let mut entry_count = 0usize;
    let mut total_size = 0u64;
    for entry_result in fs::read_dir(index)? {
        let entry = entry_result?;
        entry_count = entry_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("iWork index entry count overflow".to_owned()))?;
        if entry_count > limits.max_files {
            return Err(Error::InvalidFormat(format!(
                "iWork index directory contains more than the {} entry limit",
                limits.max_files
            )));
        }
        let entry_kind = entry.file_type()?;
        if entry_kind.is_symlink() {
            return Err(Error::InvalidFormat(format!(
                "iWork bundle index contains symbolic link {}",
                entry.path().display()
            )));
        }
        if entry_kind.is_file() {
            let entry_size = entry.metadata()?.len();
            total_size = total_size.checked_add(entry_size).ok_or_else(|| {
                Error::InvalidFormat("iWork index byte count overflow".to_owned())
            })?;
            if total_size > limits.max_total_size {
                return Err(Error::InvalidFormat(format!(
                    "iWork index directory contains {total_size} bytes, exceeding the {} byte limit",
                    limits.max_total_size
                )));
            }
            let entry_name_os = entry.file_name();
            let entry_name = entry_name_os.to_str().ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork bundle index contains a non-UTF-8 entry: {}",
                    entry.path().display()
                ))
            })?;
            marks.see_index(entry_name);
            if entry_name == "Document.iwa" {
                document = Some(entry.path());
            }
        }
    }
    if !marks.iwa {
        return Ok(Outcome::None);
    }
    let Some(document_path) = document else {
        return Err(Error::InvalidFormat(
            "iWork bundle index contains IWA components but no Document.iwa".to_owned(),
        ));
    };
    let document_size = fs::metadata(&document_path)?.len();
    if document_size > limits.max_input_bytes || document_size > limits.max_entry_size {
        let limit = limits.max_input_bytes.min(limits.max_entry_size);
        return Err(Error::InvalidFormat(format!(
            "iWork Document.iwa is {document_size} bytes, exceeding the {limit} byte limit"
        )));
    }
    let Some(format) = root_format(&fs::read(&document_path)?, limits)? else {
        return Err(Error::InvalidFormat(
            "Document.iwa has no recognized iWork application root".to_owned(),
        ));
    };
    Ok(if marks.accepts(format) {
        Outcome::Found(format)
    } else {
        Outcome::Conflict
    })
}

#[derive(Debug, Default, Clone, Copy)]
struct Marks {
    iwa: bool,
    keynote: bool,
}

impl Marks {
    #[allow(
        clippy::case_sensitive_file_extension_comparisons,
        reason = "IWA member names are case-sensitive protocol names."
    )]
    fn see_index(&mut self, name: &str) {
        if !name.ends_with(".iwa") {
            return;
        }
        self.iwa = true;
        self.keynote |= is_component(name, "MasterSlide")
            || is_component(name, "Slide")
            || is_component(name, "TemplateSlide");
    }

    fn accepts(self, format: Format) -> bool {
        !self.keynote || format == Format::Keynote
    }
}

#[allow(
    clippy::case_sensitive_file_extension_comparisons,
    reason = "IWA member names are case-sensitive protocol names."
)]
fn is_component(component_name: &str, stem: &str) -> bool {
    let Some(name_without_suffix) = component_name.strip_suffix(".iwa") else {
        return false;
    };
    let Some(suffix) = name_without_suffix.strip_prefix(stem) else {
        return false;
    };
    suffix.is_empty()
        || suffix.strip_prefix('-').is_some_and(|version| {
            !version.is_empty()
                && version
                    .split('-')
                    .all(|part| !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit()))
        })
}

fn index_name(name: &str) -> Option<&str> {
    name.strip_prefix("Index/")
        .or_else(|| (!name.contains('/')).then_some(name))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Outcome {
    None,
    Found(Format),
    Conflict,
}

impl Outcome {
    fn merge(self, other: Self) -> Self {
        match (self, other) {
            (Self::Conflict, _) | (_, Self::Conflict) => Self::Conflict,
            (Self::None, outcome) | (outcome, Self::None) => outcome,
            (Self::Found(left), Self::Found(right)) if left == right => Self::Found(left),
            (Self::Found(_), Self::Found(_)) => Self::Conflict,
        }
    }
}

fn classify(pages: bool, keynote: bool, numbers: bool) -> Outcome {
    match usize::from(pages) + usize::from(keynote) + usize::from(numbers) {
        0 => Outcome::None,
        1 if pages => Outcome::Found(Format::Pages),
        1 if keynote => Outcome::Found(Format::Keynote),
        1 => Outcome::Found(Format::Numbers),
        _ => Outcome::Conflict,
    }
}

fn marker(root: &Path, name: &str) -> Result<bool> {
    match kind(&root.join(name))? {
        Kind::File => Ok(true),
        Kind::Missing => Ok(false),
        Kind::Dir => Err(Error::InvalidFormat(format!(
            "iWork marker {} is a directory",
            root.join(name).display()
        ))),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Missing,
    File,
    Dir,
}

fn kind(value: &Path) -> Result<Kind> {
    match fs::symlink_metadata(value) {
        Ok(metadata) => {
            let kind = metadata.file_type();
            if kind.is_symlink() {
                Err(Error::InvalidFormat(format!(
                    "iWork detection refuses symbolic link {}",
                    value.display()
                )))
            } else if kind.is_file() {
                Ok(Kind::File)
            } else if kind.is_dir() {
                Ok(Kind::Dir)
            } else {
                Err(Error::InvalidFormat(format!(
                    "iWork detection refuses unsupported filesystem node {}",
                    value.display()
                )))
            }
        },
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(Kind::Missing),
        Err(error) => Err(Error::Io(error)),
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "detector tests use fixed fallible package fixtures and compare independent failures"
)]
mod tests {
    use super::*;
    use litchi_iwa_core::{ArchiveObject, RawMessage};
    use litchi_iwa_package::{Entry, EntryStore};
    use litchi_iwa_protos::{kn, tn, tp, tsa, tsk, tsp};
    use prost::Message;
    use std::io::Cursor;
    use std::sync::atomic::{AtomicU64, Ordering};

    fn shared_document() -> tsa::DocumentArchive {
        tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        }
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn document_payload(format: Format) -> Vec<u8> {
        match format {
            Format::Pages => tp::DocumentArchive {
                super_: shared_document(),
                ..Default::default()
            }
            .encode_to_vec(),
            Format::Numbers => tn::DocumentArchive {
                super_: shared_document(),
                stylesheet: reference(1),
                sidebar_order: reference(2),
                theme: reference(3),
                ..Default::default()
            }
            .encode_to_vec(),
            Format::Keynote => kn::DocumentArchive {
                super_: shared_document(),
                show: reference(1),
                ..Default::default()
            }
            .encode_to_vec(),
        }
    }

    fn package(names: &[(&str, &[u8])]) -> Vec<u8> {
        litchi_iwa_archive::package::to_bytes(
            names.iter().copied(),
            litchi_iwa_archive::Limits::default(),
        )
        .unwrap()
    }

    fn document(format: Format) -> Vec<u8> {
        let shared_document = || tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        };
        let reference = |identifier| tsp::Reference {
            identifier,
            ..Default::default()
        };
        let (message_type, payload) = match format {
            Format::Pages => (
                10_000,
                tp::DocumentArchive {
                    super_: shared_document(),
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
            Format::Keynote => (
                1,
                kn::DocumentArchive {
                    super_: shared_document(),
                    show: reference(1),
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
            Format::Numbers => (
                6,
                tn::DocumentArchive {
                    super_: shared_document(),
                    stylesheet: reference(1),
                    sidebar_order: reference(2),
                    theme: reference(3),
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
        };
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: message_type,
                        data: payload,
                    }],
                )
                .unwrap(),
            ],
        };
        SnappyStream::compress(&archive.to_bytes().unwrap()).unwrap()
    }

    fn document_package(format: Format, extra_names: &[&str]) -> Vec<u8> {
        let root = document(format);
        let mut files = vec![("Index/Document.iwa", root.as_slice())];
        files.extend(extra_names.iter().map(|name| (*name, b"iwa".as_slice())));
        package(&files)
    }

    fn frozen_package(bytes: &[u8]) -> FrozenEntryStore {
        let catalog = litchi_iwa_archive::package::Catalog::from_bytes(bytes)
            .expect("parse logical test package");
        let entries = catalog
            .into_iter()
            .map(|entry| {
                let (name, data) = entry.into_parts();
                Entry::new(name, data)
            })
            .collect();
        EntryStore::try_from_entries(entries)
            .expect("unique logical test entries")
            .freeze()
    }

    fn frozen_entries(entries: &[(&str, &[u8])]) -> FrozenEntryStore {
        EntryStore::try_from_entries(
            entries
                .iter()
                .map(|(name, data)| Entry::new((*name).to_owned(), (*data).to_vec()))
                .collect(),
        )
        .expect("unique logical test entries")
        .freeze()
    }

    #[test]
    fn test_document_payload_detection() {
        assert_eq!(
            detect_application_from_document(&document_payload(Format::Pages)),
            Some(Format::Pages)
        );
        assert_eq!(
            detect_application_from_document(&document_payload(Format::Numbers)),
            Some(Format::Numbers)
        );
        assert_eq!(
            detect_application_from_document(&document_payload(Format::Keynote)),
            Some(Format::Keynote)
        );

        let pages_with_references = tp::DocumentArchive {
            super_: shared_document(),
            stylesheet: Some(reference(1)),
            floating_drawables: Some(reference(2)),
            ..Default::default()
        }
        .encode_to_vec();
        assert_eq!(
            detect_application_from_document(&pages_with_references),
            Some(Format::Pages)
        );

        let mut conflicting = document_payload(Format::Pages);
        conflicting.extend(document_payload(Format::Numbers));
        assert_eq!(detect_application_from_document(&conflicting), None);

        let mut conflicting = document_payload(Format::Pages);
        conflicting.extend(document_payload(Format::Keynote));
        assert_eq!(detect_application_from_document(&conflicting), None);

        assert_eq!(detect_application_from_document(&[0x78, 0x00]), None);
        assert_eq!(detect_application_from_document(&[0x7a, 0x00]), None);
        assert_eq!(detect_application_from_document(&[0x80]), None);
    }

    #[test]
    fn detects_root_application_with_shared_table_components() {
        assert_eq!(
            bytes(&document_package(Format::Pages, &[])).unwrap(),
            Some(Format::Pages)
        );
        assert_eq!(
            bytes(&document_package(
                Format::Keynote,
                &[
                    "Index/MasterSlide-12.iwa",
                    "Index/Slide-1.iwa",
                    "Index/TemplateSlide-31.iwa",
                    "Index/CalculationEngine-81.iwa"
                ]
            ))
            .unwrap(),
            Some(Format::Keynote)
        );
        assert_eq!(
            bytes(&document_package(
                Format::Numbers,
                &["Index/CalculationEngine-174.iwa"]
            ))
            .unwrap(),
            Some(Format::Numbers)
        );
        assert_eq!(
            bytes(&document_package(
                Format::Pages,
                &["Index/CalculationEngine.iwa"]
            ))
            .unwrap(),
            Some(Format::Pages)
        );
        assert!(bytes(&document_package(Format::Numbers, &["Index/Slide-1.iwa"])).is_err());
        assert!(
            bytes(&document_package(
                Format::Pages,
                &["Index/MasterSlide-12.iwa"]
            ))
            .is_err()
        );
        assert!(bytes(&package(&[("Index/Document.iwa", b"not iwa")])).is_err());
        assert_eq!(
            bytes(&package(&[("Index/Unknown.iwa", b"iwa")])).unwrap(),
            None
        );
        assert_eq!(
            bytes(&package(&[("Data/image.png", b"iwa")])).unwrap(),
            None
        );

        let duplicate_root = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: document_payload(Format::Pages),
                    }],
                )
                .unwrap(),
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: document_payload(Format::Pages),
                    }],
                )
                .unwrap(),
            ],
        };
        assert!(root_format_archive(&duplicate_root).is_err());
    }

    #[test]
    fn detects_from_the_same_parsed_component_snapshot() {
        for expected in [Format::Pages, Format::Numbers, Format::Keynote] {
            let bytes = document_package(expected, &[]);
            let source = SourceCatalog::from_bytes(&bytes).unwrap();

            assert_eq!(
                component_catalog(source.components()).unwrap(),
                Some(expected)
            );
            assert_eq!(
                component_catalog(source.components()).unwrap(),
                super::bytes(&bytes).unwrap()
            );
        }

        let media_package = package(&[("Data/image.png", b"not an IWA component")]);
        let source = SourceCatalog::from_bytes(&media_package).unwrap();
        assert_eq!(component_catalog(source.components()).unwrap(), None);

        let pages_root = document(Format::Pages);
        let legacy_index = package(&[("Document.iwa", &pages_root)]);
        let legacy = package(&[("legacy.pages/Index.zip", &legacy_index)]);
        let legacy_source = SourceCatalog::from_bytes(&legacy).unwrap();
        assert_eq!(
            component_catalog(legacy_source.components()).unwrap(),
            Some(Format::Pages)
        );

        let duplicate = package(&[
            ("Index/Document.iwa", &pages_root),
            ("Document.iwa", &pages_root),
        ]);
        let duplicate_source = SourceCatalog::from_bytes(&duplicate).unwrap();
        assert!(component_catalog(duplicate_source.components()).is_err());
    }

    #[test]
    fn prepared_source_retains_one_shared_snapshot_and_profile() {
        let bytes: Arc<[u8]> = document_package(Format::Pages, &[]).into();
        let limits = Limits::new(
            u64::try_from(bytes.len()).unwrap(),
            32,
            1024 * 1024,
            2 * 1024 * 1024,
            4 * 1024 * 1024,
        )
        .unwrap();

        let prepared = PreparedSource::from_shared_bytes_with_limits(Arc::clone(&bytes), limits)
            .unwrap()
            .unwrap();
        assert_eq!(prepared.format(), Format::Pages);
        assert_eq!(prepared.limits(), limits);

        let catalog = prepared
            .__into_source_catalog()
            .expect("byte-prepared sources retain a package catalog");
        assert!(Arc::ptr_eq(&bytes, &catalog.shared_source()));
        assert_eq!(catalog.limits().max_input_bytes(), limits.max_input_bytes());
        assert_eq!(catalog.limits().max_entries(), limits.max_files());
        assert_eq!(
            catalog.limits().max_iwa_stream_bytes(),
            limits.max_iwa_stream_size()
        );
    }

    #[test]
    fn frozen_logical_preparation_matches_packaged_classification() {
        for expected in [Format::Pages, Format::Numbers, Format::Keynote] {
            let bytes = document_package(expected, &[]);
            let prepared = PreparedSource::from_frozen_entries(frozen_package(&bytes))
                .unwrap()
                .expect("recognized logical source");
            assert_eq!(prepared.format(), expected);
            assert!(prepared.__into_source_catalog().is_none());

            let component_prepared = PreparedSource::from_frozen_entries(frozen_package(&bytes))
                .unwrap()
                .expect("recognized logical source");
            let (logical, retained_limits) = component_prepared.__into_components();
            let packaged = SourceCatalog::from_bytes(&bytes).unwrap();
            assert_eq!(retained_limits, archive_limits(Limits::default()).unwrap());
            assert_eq!(
                logical
                    .iter()
                    .map(litchi_iwa_archive::Component::name)
                    .collect::<Vec<_>>(),
                packaged
                    .components()
                    .iter()
                    .map(litchi_iwa_archive::Component::name)
                    .collect::<Vec<_>>()
            );
        }
    }

    #[test]
    fn frozen_logical_preparation_skips_operation_storage() {
        let root = document(Format::Pages);
        let prepared = PreparedSource::from_frozen_entries(frozen_entries(&[
            ("Index/Document.iwa", &root),
            ("Index/OperationStorage.iwa", b"bvxn operation log"),
        ]))
        .unwrap()
        .expect("recognized Pages source");
        let (components, _limits) = prepared.__into_components();
        assert_eq!(components.len(), 1);
        assert_eq!(
            components
                .iter()
                .next()
                .map(litchi_iwa_archive::Component::name),
            Some("Index/Document.iwa")
        );
    }

    #[test]
    fn frozen_logical_limits_keep_typed_categories() {
        let cases = [
            (
                frozen_entries(&[("a", b""), ("b", b"")]),
                Limits::new(16, 1, 1, 1, 1024).unwrap(),
                LimitKind::Entries,
            ),
            (
                frozen_entries(&[("ab", b"")]),
                Limits::new(1, 1, 1, 1, 1024).unwrap(),
                LimitKind::MemberNameBytes,
            ),
            (
                frozen_entries(&[("a", b""), ("b", b"")]),
                Limits::new(1, 2, 1, 1, 1024).unwrap(),
                LimitKind::MetadataBytes,
            ),
            (
                frozen_entries(&[("a", b"12")]),
                Limits::new(16, 1, 1, 2, 1024).unwrap(),
                LimitKind::EntryBytes,
            ),
            (
                frozen_entries(&[("a", b"1"), ("b", b"2")]),
                Limits::new(16, 2, 1, 1, 1024).unwrap(),
                LimitKind::TotalBytes,
            ),
        ];
        for (entries, limits, expected) in cases {
            let error = PreparedSource::from_frozen_entries_with_limits(entries, limits)
                .expect_err("logical limit must be retained");
            assert!(matches!(
                error,
                Error::LimitExceeded { kind, .. } if kind == expected
            ));
        }
    }

    #[test]
    fn frozen_logical_preparation_rejects_encryption_and_unexpanded_indexes() {
        assert!(matches!(
            PreparedSource::from_frozen_entries(frozen_entries(&[("Metadata/.iwpv2", b"marker")])),
            Err(Error::Encrypted)
        ));
        assert!(matches!(
            PreparedSource::from_frozen_entries(frozen_entries(&[("legacy/Index.zip", b"nested")])),
            Err(Error::Archive(_))
        ));
    }

    #[test]
    fn prepared_source_leaves_unrecognized_inputs_unclaimed() {
        assert!(
            PreparedSource::from_bytes(b"not a ZIP package")
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn detects_nested_legacy_indexes_and_rejects_ambiguous_or_encrypted_packages() {
        let root = document(Format::Pages);
        let index = package(&[("Document.iwa", &root)]);
        let outer = package(&[("legacy.pages/Index.zip", &index)]);
        assert_eq!(bytes(&outer).unwrap(), Some(Format::Pages));

        let mixed = package(&[
            ("legacy.pages/Index.zip", &index),
            ("Index/CalculationEngine.iwa", b"iwa"),
        ]);
        assert!(bytes(&mixed).is_err());

        let ambiguous = package(&[("a/Index.zip", &index), ("b/Index.zip", &index)]);
        assert!(bytes(&ambiguous).is_err());

        let root = document(Format::Pages);
        let encrypted = package(&[
            ("Index/Document.iwa", &root),
            ("Metadata/.iwpv2", b"encryption metadata"),
        ]);
        assert!(matches!(bytes(&encrypted), Err(Error::Encrypted)));
    }

    #[test]
    fn archive_failures_keep_content_free_typed_categories() {
        let kinds = [
            (
                litchi_iwa_archive::LimitKind::InputBytes,
                LimitKind::InputBytes,
            ),
            (
                litchi_iwa_archive::LimitKind::OutputBytes,
                LimitKind::OutputBytes,
            ),
            (litchi_iwa_archive::LimitKind::Entries, LimitKind::Entries),
            (
                litchi_iwa_archive::LimitKind::MemberNameBytes,
                LimitKind::MemberNameBytes,
            ),
            (
                litchi_iwa_archive::LimitKind::MetadataBytes,
                LimitKind::MetadataBytes,
            ),
            (
                litchi_iwa_archive::LimitKind::CompressedEntryBytes,
                LimitKind::CompressedEntryBytes,
            ),
            (
                litchi_iwa_archive::LimitKind::EntryBytes,
                LimitKind::EntryBytes,
            ),
            (
                litchi_iwa_archive::LimitKind::TotalBytes,
                LimitKind::TotalBytes,
            ),
            (
                litchi_iwa_archive::LimitKind::IwaStreamBytes,
                LimitKind::IwaStreamBytes,
            ),
            (
                litchi_iwa_archive::LimitKind::IwaTotalBytes,
                LimitKind::IwaTotalBytes,
            ),
        ];
        for (archive_kind, expected) in kinds {
            let error = map_archive_error(litchi_iwa_archive::Error::Limit {
                kind: archive_kind,
                observed: 9,
                maximum: 8,
            });
            assert!(matches!(
                error,
                Error::LimitExceeded {
                    kind,
                    observed: 9,
                    maximum: 8,
                } if kind == expected
            ));
        }

        let errors = [
            map_archive_error(litchi_iwa_archive::Error::Encrypted),
            map_archive_error(litchi_iwa_archive::Error::InvalidLimits(
                "private implementation detail".to_owned(),
            )),
            map_archive_error(litchi_iwa_archive::Error::Allocation {
                resource: "private implementation detail",
                amount: 17,
            }),
        ];
        assert!(matches!(errors[0], Error::Encrypted));
        assert!(matches!(errors[1], Error::InvalidLimits));
        assert!(matches!(errors[2], Error::Allocation { amount: 17 }));
        for error in &errors {
            assert!(std::error::Error::source(error).is_none());
            assert!(!error.to_string().contains("private implementation detail"));
        }
    }

    #[test]
    fn checked_limits_preserve_defaults_and_bound_each_layer() {
        let valid = document_package(Format::Pages, &[]);
        let defaults = Limits::default();
        assert_eq!(
            bytes_with_limits(&valid, defaults).unwrap(),
            Some(Format::Pages)
        );
        assert_eq!(defaults.max_input_bytes(), Limits::HARD_MAX_INPUT_BYTES);
        assert_eq!(defaults.max_files(), Limits::HARD_MAX_FILES);
        assert_eq!(defaults.max_entry_size(), Limits::HARD_MAX_ENTRY_SIZE);
        assert_eq!(defaults.max_total_size(), Limits::HARD_MAX_TOTAL_SIZE);
        assert_eq!(
            defaults.max_iwa_stream_size(),
            Limits::HARD_MAX_IWA_STREAM_SIZE
        );

        let input_bound = Limits::new(1, 1, 1, 1, 1).unwrap();
        assert!(bytes_with_limits(&valid, input_bound).is_err());

        let stream_bound = Limits::new(
            Limits::HARD_MAX_INPUT_BYTES,
            Limits::HARD_MAX_FILES,
            Limits::HARD_MAX_ENTRY_SIZE,
            Limits::HARD_MAX_TOTAL_SIZE,
            1,
        )
        .unwrap();
        assert!(bytes_with_limits(&valid, stream_bound).is_err());
    }

    #[test]
    fn checked_limits_reject_zero_and_hard_ceiling_escapes() {
        assert!(Limits::new(0, 1, 1, 1, 1).is_err());
        assert!(Limits::new(1, 0, 1, 1, 1).is_err());
        assert!(Limits::new(1, 1, 0, 1, 1).is_err());
        assert!(Limits::new(1, 1, 1, 0, 1).is_err());
        assert!(Limits::new(1, 1, 1, 1, 0).is_err());
        assert!(Limits::new(Limits::HARD_MAX_INPUT_BYTES + 1, 1, 1, 1, 1).is_err());
        assert!(Limits::new(1, Limits::HARD_MAX_FILES + 1, 1, 1, 1).is_err());
        assert!(Limits::new(1, 1, Limits::HARD_MAX_ENTRY_SIZE + 1, 1, 1).is_err());
        assert!(Limits::new(1, 1, 1, Limits::HARD_MAX_TOTAL_SIZE + 1, 1).is_err());
        assert!(Limits::new(1, 1, 1, 1, Limits::HARD_MAX_IWA_STREAM_SIZE + 1).is_err());
    }

    #[test]
    fn reader_restores_nonzero_cursor_on_success_and_rejection() {
        let mut valid = Cursor::new(document_package(Format::Pages, &[]));
        valid.set_position(9);
        assert_eq!(reader(&mut valid).unwrap(), Some(Format::Pages));
        assert_eq!(valid.position(), 9);

        let mut invalid = Cursor::new(b"not an iWork package".to_vec());
        invalid.set_position(4);
        assert_eq!(reader(&mut invalid).unwrap(), None);
        assert_eq!(invalid.position(), 4);
    }

    #[test]
    fn path_supports_files_legacy_bundles_and_index_zip() -> std::io::Result<()> {
        let temp = Temp::new()?;

        let packaged = temp.0.join("document.pages");
        fs::write(&packaged, document_package(Format::Pages, &[]))?;
        assert_eq!(path(&packaged).unwrap(), Some(Format::Pages));

        let legacy = temp.0.join("legacy.key");
        fs::create_dir(&legacy)?;
        fs::write(legacy.join("index.apxl"), [])?;
        assert_eq!(path(&legacy).unwrap(), Some(Format::Keynote));

        let bundle = temp.0.join("sheet.numbers");
        fs::create_dir(&bundle)?;
        fs::write(
            bundle.join("Index.zip"),
            document_package(Format::Numbers, &["Index/CalculationEngine-174.iwa"]),
        )?;
        assert_eq!(path(&bundle).unwrap(), Some(Format::Numbers));

        let agreeing = temp.0.join("agreeing.pages");
        fs::create_dir(&agreeing)?;
        fs::write(agreeing.join("index.xml"), [])?;
        fs::write(
            agreeing.join("Index.zip"),
            document_package(Format::Pages, &[]),
        )?;
        assert_eq!(path(&agreeing).unwrap(), Some(Format::Pages));

        let unpacked = temp.0.join("unpacked.key");
        fs::create_dir_all(unpacked.join("Index"))?;
        fs::write(
            unpacked.join("Index/Document.iwa"),
            document(Format::Keynote),
        )?;
        fs::write(unpacked.join("Index/Slide-1.iwa"), [])?;
        assert_eq!(path(&unpacked).unwrap(), Some(Format::Keynote));

        let tight = Limits::new(1, 1, 1, 1, 1).unwrap();
        assert!(path_with_limits(&packaged, tight).is_err());
        assert!(path_with_limits(&unpacked, tight).is_err());
        Ok(())
    }

    #[test]
    fn unpacked_index_entry_count_is_bounded_before_document_read() -> std::io::Result<()> {
        let temp = Temp::new()?;
        let unpacked = temp.0.join("bounded.pages");
        fs::create_dir_all(unpacked.join("Index"))?;
        fs::write(unpacked.join("Index/Document.iwa"), document(Format::Pages))?;
        fs::write(unpacked.join("Index/Extra.iwa"), [])?;

        let limits = Limits::new(
            Limits::HARD_MAX_INPUT_BYTES,
            1,
            Limits::HARD_MAX_ENTRY_SIZE,
            Limits::HARD_MAX_TOTAL_SIZE,
            Limits::HARD_MAX_IWA_STREAM_SIZE,
        )
        .unwrap();
        let error = path_with_limits(&unpacked, limits).unwrap_err();
        assert!(error.to_string().contains("more than the 1 entry limit"));
        Ok(())
    }

    #[test]
    fn unpacked_index_total_size_is_bounded() -> std::io::Result<()> {
        let temp = Temp::new()?;
        let unpacked = temp.0.join("total-size.pages");
        fs::create_dir_all(unpacked.join("Index"))?;
        fs::write(unpacked.join("Index/Document.iwa"), document(Format::Pages))?;
        fs::write(unpacked.join("Index/Extra.iwa"), b"extra bytes")?;
        let total = fs::metadata(unpacked.join("Index/Document.iwa"))?.len()
            + fs::metadata(unpacked.join("Index/Extra.iwa"))?.len();
        let limits = Limits::new(
            Limits::HARD_MAX_INPUT_BYTES,
            Limits::HARD_MAX_FILES,
            Limits::HARD_MAX_ENTRY_SIZE,
            total - 1,
            Limits::HARD_MAX_IWA_STREAM_SIZE,
        )
        .unwrap();
        let error = path_with_limits(&unpacked, limits).unwrap_err();
        assert!(error.to_string().contains("iWork index directory contains"));
        assert!(error.to_string().contains("byte limit"));
        Ok(())
    }

    #[test]
    fn path_rejects_generic_and_conflicting_index_directories() -> std::io::Result<()> {
        let temp = Temp::new()?;

        let generic = temp.0.join("generic");
        fs::create_dir_all(generic.join("Index"))?;
        fs::write(generic.join("Index/Unknown.iwa"), [])?;
        assert!(path(&generic).is_err());

        let media = temp.0.join("media-only");
        fs::create_dir_all(media.join("Data"))?;
        fs::create_dir(media.join("Assets"))?;
        fs::write(media.join("theme-preview.jpg"), [])?;
        assert_eq!(path(&media).unwrap(), None);

        let conflict = temp.0.join("conflict");
        fs::create_dir_all(conflict.join("Index"))?;
        fs::write(conflict.join("Index/Document.iwa"), document(Format::Pages))?;
        fs::write(conflict.join("Index/Slide.iwa"), [])?;
        fs::write(conflict.join("Index/CalculationEngine.iwa"), [])?;
        assert!(path(&conflict).is_err());

        let legacy_conflict = temp.0.join("legacy-conflict");
        fs::create_dir(&legacy_conflict)?;
        fs::write(legacy_conflict.join("index.xml"), [])?;
        fs::write(
            legacy_conflict.join("Index.zip"),
            document_package(Format::Numbers, &[]),
        )?;
        assert!(path(&legacy_conflict).is_err());

        let representation_conflict = temp.0.join("representation-conflict");
        fs::create_dir_all(representation_conflict.join("Index"))?;
        fs::write(
            representation_conflict.join("Index.zip"),
            document_package(Format::Pages, &[]),
        )?;
        fs::write(
            representation_conflict.join("Index/Document.iwa"),
            document(Format::Keynote),
        )?;
        fs::write(representation_conflict.join("Index/Slide-1.iwa"), [])?;
        assert!(path(&representation_conflict).is_err());
        Ok(())
    }

    struct Temp(std::path::PathBuf);

    impl Temp {
        fn new() -> std::io::Result<Self> {
            static NEXT: AtomicU64 = AtomicU64::new(0);
            let path = std::env::temp_dir().join(format!(
                "litchi-iwa-detect-{}-{}",
                std::process::id(),
                NEXT.fetch_add(1, Ordering::Relaxed)
            ));
            fs::create_dir(&path)?;
            Ok(Self(path))
        }
    }

    impl Drop for Temp {
        fn drop(&mut self) {
            drop(fs::remove_dir_all(&self.0));
        }
    }
}
