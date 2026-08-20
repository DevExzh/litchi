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
    self, ComponentCatalog, DetectionRoot, DirectoryMarkers, DirectoryMetadataSidecars,
    FrozenDirectoryBundle, LogicalEntryLimits, LogicalSourceCatalog, SemanticMetadataSidecars,
    SemanticProfile, SemanticProjection, SourceCatalog,
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
/// Maximum canonical properties diagnostic retained for a semantic owner.
pub const MAX_PROPERTIES_BYTES: usize = litchi_iwa_archive::MAX_DIRECTORY_PROPERTIES_BYTES as usize;

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
    /// Aggregate IWA objects across semantic components.
    IwaObjects,
    /// Protobuf-style fields parsed while classifying an IWA root.
    IwaFields,
    /// Nested protobuf traversal while classifying an IWA root.
    IwaNesting,
    /// Aggregate bounded wire work while classifying an IWA root.
    IwaWork,
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
            Self::IwaObjects => "aggregate IWA objects",
            Self::IwaFields => "IWA root fields",
            Self::IwaNesting => "IWA root nesting",
            Self::IwaWork => "IWA root work",
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
                kind,
                observed,
                maximum,
            } => Self::LimitExceeded {
                kind: if kind == litchi_iwa_core::LimitKind::Objects {
                    LimitKind::IwaObjects
                } else {
                    LimitKind::IwaStreamBytes
                },
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
        match error {
            litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit,
            } => Self::LimitExceeded {
                kind: match kind {
                    litchi_iwa_common::LimitKind::InputBytes => LimitKind::IwaStreamBytes,
                    litchi_iwa_common::LimitKind::Fields => LimitKind::IwaFields,
                    litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                    litchi_iwa_common::LimitKind::Nesting => LimitKind::IwaNesting,
                    litchi_iwa_common::LimitKind::RewriteWork => LimitKind::IwaWork,
                    litchi_iwa_common::LimitKind::TableRows
                    | litchi_iwa_common::LimitKind::TableColumns
                    | litchi_iwa_common::LimitKind::TableCells
                    | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::IwaFields,
                },
                observed: usize_u64(observed),
                maximum: usize_u64(limit),
            },
            litchi_iwa_common::Error::Allocation { amount, .. } => Self::Allocation { amount },
            litchi_iwa_common::Error::InvalidLimit { .. } => Self::InvalidLimits,
            invalid @ litchi_iwa_common::Error::InvalidFormat(_) => Self::IwaCommon(invalid),
        }
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

/// Canonical metadata diagnostics retained for a format-owned semantic reader.
///
/// This explicitly unstable DTO exposes only the three fixed authorities used
/// by format-owned Pages and Numbers readers. It deliberately carries neither
/// physical package entries nor caller-selected paths.
#[doc(hidden)]
#[derive(Clone, Default)]
pub struct PreparedMetadataSidecars {
    properties_plist: Option<Arc<[u8]>>,
    build_version_history_plist: Option<Arc<[u8]>>,
    document_identifier: Option<Arc<[u8]>>,
}

impl fmt::Debug for PreparedMetadataSidecars {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedMetadataSidecars")
            .field(
                "properties_plist_bytes",
                &self.properties_plist.as_deref().map(<[u8]>::len),
            )
            .field(
                "build_version_history_plist_bytes",
                &self.build_version_history_plist.as_deref().map(<[u8]>::len),
            )
            .field(
                "document_identifier_bytes",
                &self.document_identifier.as_deref().map(<[u8]>::len),
            )
            .finish()
    }
}

impl PreparedMetadataSidecars {
    /// Borrow canonical `Metadata/Properties.plist`, when present.
    #[must_use]
    pub fn properties_plist(&self) -> Option<&[u8]> {
        self.properties_plist.as_deref()
    }

    /// Borrow canonical `Metadata/BuildVersionHistory.plist`, when present.
    #[must_use]
    pub fn build_version_history_plist(&self) -> Option<&[u8]> {
        self.build_version_history_plist.as_deref()
    }

    /// Borrow canonical `Metadata/DocumentIdentifier`, when present.
    #[must_use]
    pub fn document_identifier(&self) -> Option<&[u8]> {
        self.document_identifier.as_deref()
    }
}

impl From<DirectoryMetadataSidecars> for PreparedMetadataSidecars {
    fn from(sidecars: DirectoryMetadataSidecars) -> Self {
        let (properties_plist, build_version_history_plist, document_identifier) =
            sidecars.__into_parts();
        Self {
            properties_plist,
            build_version_history_plist,
            document_identifier,
        }
    }
}

impl From<SemanticMetadataSidecars> for PreparedMetadataSidecars {
    fn from(sidecars: SemanticMetadataSidecars) -> Self {
        let (properties_plist, build_version_history_plist, document_identifier) =
            sidecars.into_parts();
        Self {
            properties_plist: properties_plist.map(Arc::from),
            build_version_history_plist: build_version_history_plist.map(Arc::from),
            document_identifier: document_identifier.map(Arc::from),
        }
    }
}

/// Archive-free semantic state consumed by one format owner.
///
/// This DTO preserves the checked archive profile and immutable component
/// snapshot without exposing exact-package authority.
#[doc(hidden)]
pub struct PreparedSemanticSource {
    components: Arc<ComponentCatalog>,
    limits: litchi_iwa_archive::Limits,
    sidecars: PreparedMetadataSidecars,
}

impl fmt::Debug for PreparedSemanticSource {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedSemanticSource")
            .field("components", &self.components.len())
            .field("limits", &self.limits)
            .field("sidecars", &self.sidecars)
            .finish()
    }
}

impl PreparedSemanticSource {
    /// Borrow the immutable semantic component catalog.
    #[must_use]
    pub fn components(&self) -> &ComponentCatalog {
        &self.components
    }

    /// Return the physical profile that authorized the component snapshot.
    #[must_use]
    pub const fn archive_limits(&self) -> litchi_iwa_archive::Limits {
        self.limits
    }

    /// Borrow the fixed canonical metadata diagnostics.
    #[must_use]
    pub const fn sidecars(&self) -> &PreparedMetadataSidecars {
        &self.sidecars
    }

    /// Consume this DTO into the format-owned semantic parts.
    #[must_use]
    pub fn __into_parts(
        self,
    ) -> (
        Arc<ComponentCatalog>,
        litchi_iwa_archive::Limits,
        PreparedMetadataSidecars,
    ) {
        (self.components, self.limits, self.sidecars)
    }
}

enum PreparedBacking {
    Package(SourceCatalog),
    FormatOnly {
        limits: litchi_iwa_archive::Limits,
    },
    Semantic {
        components: Arc<ComponentCatalog>,
        limits: litchi_iwa_archive::Limits,
        sidecars: PreparedMetadataSidecars,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PathProfile {
    IndexOnly,
    Properties,
    SemanticMetadata(Format),
    SemanticMetadataAny,
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

    /// Prepare borrowed package bytes with Pages' fixed metadata authorities
    /// checked from physical ZIP headers before package payload decode.
    ///
    /// Ordinary detection and other format owners remain on the generic ZIP
    /// profile. This unstable opt-in caps each exact normalized Pages metadata
    /// member at 64 KiB and rejects unsupported compression.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_bytes_with_limits`], plus a
    /// refusal when a selected Pages metadata authority violates its physical
    /// profile.
    #[doc(hidden)]
    pub fn __from_bytes_with_pages_metadata(value: &[u8], limits: Limits) -> Result<Option<Self>> {
        Self::from_bytes_with_catalog_metadata(value, limits, Some(Format::Pages))
    }

    /// Prepare borrowed package bytes for an archive-free semantic reader.
    ///
    /// The detector first classifies the canonical application root and then
    /// retains only the IWA components and the three bounded semantic metadata
    /// authorities. Unrelated `Data/`, `Preview/`, and unknown members remain
    /// in the exact source allocation but are not materialized as package
    /// entries. This route is intended for read-only metadata, list, and
    /// one-query workflows; preserve-mode owners continue to use
    /// [`Self::from_bytes_with_limits`].
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_bytes_with_limits`], plus a
    /// refusal when a selected semantic metadata authority violates its
    /// physical profile.
    #[doc(hidden)]
    pub fn __from_bytes_with_semantic_metadata(
        value: &[u8],
        limits: Limits,
    ) -> Result<Option<Self>> {
        Self::from_bytes_with_catalog_metadata(value, limits, None)
    }

    /// Prepare borrowed package bytes for Numbers' archive-free semantic
    /// reader under the fixed canonical metadata profile.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_bytes_with_limits`], plus a
    /// refusal when a selected canonical metadata authority violates its
    /// physical profile.
    #[doc(hidden)]
    pub fn __from_bytes_with_numbers_semantics(
        value: &[u8],
        limits: Limits,
    ) -> Result<Option<Self>> {
        Self::from_bytes_with_semantic_projection(value, limits, Format::Numbers)
    }

    fn from_bytes_with_catalog_metadata(
        value: &[u8],
        limits: Limits,
        expected: Option<Format>,
    ) -> Result<Option<Self>> {
        if !check_prepared_candidate(value, limits)? {
            return Ok(None);
        }
        let archive_limits = archive_limits(limits)?;
        let root = litchi_iwa_archive::inspect_semantic_detection_root(value, archive_limits)
            .map_err(map_archive_error)?;
        let Some(format) = classify_root(&root)? else {
            return Ok(None);
        };
        if expected.is_some_and(|expected| format != expected) {
            return Ok(Some(Self::format_only(format, limits, archive_limits)));
        }
        let catalog = SourceCatalog::__from_bytes_with_logical_entry_limits(
            value,
            archive_limits,
            LogicalEntryLimits::SEMANTIC_METADATA,
        )
        .map_err(map_archive_error)?;
        Self::from_catalog(catalog, limits)
    }

    fn from_bytes_with_semantic_projection(
        value: &[u8],
        limits: Limits,
        expected: Format,
    ) -> Result<Option<Self>> {
        if !check_prepared_candidate(value, limits)? {
            return Ok(None);
        }
        let archive_limits = archive_limits(limits)?;
        let root = litchi_iwa_archive::inspect_semantic_detection_root(value, archive_limits)
            .map_err(map_archive_error)?;
        let Some(format) = classify_root(&root)? else {
            return Ok(None);
        };
        if format != expected {
            return Ok(Some(Self::format_only(format, limits, archive_limits)));
        }
        let projection = SemanticProjection::from_bytes_with_limits(
            value,
            archive_limits,
            SemanticProfile::Metadata,
        )
        .map_err(map_archive_error)?;
        let (components, sidecars, projected_limits) = projection.into_parts();
        Ok(Some(Self {
            backing: PreparedBacking::Semantic {
                components: Arc::new(components),
                limits: projected_limits,
                sidecars: sidecars.into(),
            },
            format,
            limits,
        }))
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

    /// Prepare shared package bytes with Pages' fixed metadata authorities
    /// checked from physical ZIP headers before package payload decode.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_shared_bytes_with_limits`],
    /// plus a refusal when a selected Pages metadata authority violates its
    /// physical profile.
    #[doc(hidden)]
    pub fn __from_shared_bytes_with_pages_metadata(
        value: Arc<[u8]>,
        limits: Limits,
    ) -> Result<Option<Self>> {
        Self::from_shared_bytes_with_catalog_metadata(value, limits, Some(Format::Pages))
    }

    /// Prepare shared package bytes for an archive-free semantic reader.
    ///
    /// The exact shared source allocation remains authoritative while the
    /// selected format consumes only IWA components and the three bounded
    /// semantic metadata authorities. Unrelated media and opaque members are
    /// therefore preserved for a later exact write but are not materialized
    /// during this read-only preparation.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_shared_bytes_with_limits`],
    /// plus a refusal when a selected semantic metadata authority violates
    /// its physical profile.
    #[doc(hidden)]
    pub fn __from_shared_bytes_with_semantic_metadata(
        value: Arc<[u8]>,
        limits: Limits,
    ) -> Result<Option<Self>> {
        Self::from_shared_bytes_with_catalog_metadata(value, limits, None)
    }

    /// Prepare shared package bytes for Numbers' archive-free semantic reader
    /// under the fixed canonical metadata profile.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_shared_bytes_with_limits`],
    /// plus a refusal when a selected canonical metadata authority violates
    /// its physical profile.
    #[doc(hidden)]
    pub fn __from_shared_bytes_with_numbers_semantics(
        value: Arc<[u8]>,
        limits: Limits,
    ) -> Result<Option<Self>> {
        Self::from_bytes_with_semantic_projection(&value, limits, Format::Numbers)
    }

    fn from_shared_bytes_with_catalog_metadata(
        value: Arc<[u8]>,
        limits: Limits,
        expected: Option<Format>,
    ) -> Result<Option<Self>> {
        if !check_prepared_candidate(&value, limits)? {
            return Ok(None);
        }
        let archive_limits = archive_limits(limits)?;
        let root = litchi_iwa_archive::inspect_semantic_detection_root(&value, archive_limits)
            .map_err(map_archive_error)?;
        let Some(format) = classify_root(&root)? else {
            return Ok(None);
        };
        if expected.is_some_and(|expected| format != expected) {
            return Ok(Some(Self::format_only(format, limits, archive_limits)));
        }
        let catalog = SourceCatalog::__from_shared_bytes_with_logical_entry_limits(
            value,
            archive_limits,
            LogicalEntryLimits::SEMANTIC_METADATA,
        )
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
        logical
            .components()
            .__validate_semantic_object_limit()
            .map_err(map_archive_error)?;
        let archive_limits = logical.limits();
        let components = Arc::new(logical.into_components());
        Ok(Some(Self {
            backing: PreparedBacking::Semantic {
                components,
                limits: archive_limits,
                sidecars: PreparedMetadataSidecars::default(),
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
        Self::from_path_with_profile(value.as_ref(), limits, PathProfile::IndexOnly)
    }

    /// Prepare a path while retaining the canonical properties diagnostic for
    /// archive-free format-owned semantic readers.
    ///
    /// Ordinary format detection remains index-only. This unstable opt-in is
    /// used only when the selected semantic owner exposes metadata.
    #[doc(hidden)]
    pub fn __from_path_with_properties(
        value: impl AsRef<Path>,
        limits: Limits,
    ) -> Result<Option<Self>> {
        Self::from_path_with_profile(value.as_ref(), limits, PathProfile::Properties)
    }

    /// Prepare a packaged file or app-authored directory for an archive-free
    /// semantic reader.
    ///
    /// Regular-file ingress retains the exact source allocation while
    /// materializing only canonical IWA components and the three bounded
    /// semantic metadata authorities. Directory ingress uses the same frozen
    /// semantic metadata profile. Unrelated media, previews, and opaque
    /// members remain outside the semantic snapshot. On Windows, the
    /// packaged-file adapter retains its existing generic capture capability;
    /// byte and shared-byte ingress remain selective on every platform.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_path_with_limits`], plus a
    /// refusal when a selected semantic metadata authority is unsafe.
    #[doc(hidden)]
    pub fn __from_path_with_semantic_metadata(
        value: impl AsRef<Path>,
        limits: Limits,
    ) -> Result<Option<Self>> {
        Self::from_path_with_profile(value.as_ref(), limits, PathProfile::SemanticMetadataAny)
    }

    /// Prepare a path while retaining exactly Pages' three canonical metadata
    /// diagnostics for its archive-free semantic reader.
    ///
    /// Ordinary detection remains index-only, and the Keynote properties
    /// profile remains properties-only. This opt-in freezes
    /// `Metadata/Properties.plist`, `Metadata/BuildVersionHistory.plist`, and
    /// `Metadata/DocumentIdentifier` from one pinned directory authority.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_path_with_limits`], plus a
    /// refusal when a selected canonical metadata authority is unsafe.
    #[doc(hidden)]
    pub fn __from_path_with_pages_metadata(
        value: impl AsRef<Path>,
        limits: Limits,
    ) -> Result<Option<Self>> {
        Self::from_path_with_profile(
            value.as_ref(),
            limits,
            PathProfile::SemanticMetadata(Format::Pages),
        )
    }

    /// Prepare a path for Numbers' archive-free semantic reader while
    /// retaining exactly the three fixed canonical metadata diagnostics.
    ///
    /// Directory sidecars are inspected only after the frozen application
    /// root and markers agree that the source is Numbers. Packaged files use
    /// the semantic ZIP profile; Windows path ingress fails closed while the
    /// byte and shared-byte entry points remain available.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_path_with_limits`], plus a
    /// refusal when a selected canonical metadata authority is unsafe.
    #[doc(hidden)]
    pub fn __from_path_with_numbers_semantics(
        value: impl AsRef<Path>,
        limits: Limits,
    ) -> Result<Option<Self>> {
        Self::from_path_with_profile(
            value.as_ref(),
            limits,
            PathProfile::SemanticMetadata(Format::Numbers),
        )
    }

    fn from_path_with_profile(
        path: &Path,
        limits: Limits,
        profile: PathProfile,
    ) -> Result<Option<Self>> {
        match kind(path)? {
            Kind::File => {
                // Focused semantic path ingress promises a stable, no-follow
                // physical snapshot. The current Windows adapter cannot pin
                // that identity across reparse-point resolution, so keep the
                // byte/shared-byte APIs available there and fail closed for
                // this path-owned profile.
                #[cfg(windows)]
                if let PathProfile::SemanticMetadata(format) = profile {
                    return Err(Error::InvalidFormat(format!(
                        "{format:?} package path ingress is unsupported on Windows"
                    )));
                }
                let source = read_stable_package_file(path, limits)?;
                match profile {
                    PathProfile::SemanticMetadata(Format::Numbers) => {
                        Self::from_bytes_with_semantic_projection(&source, limits, Format::Numbers)
                    },
                    PathProfile::SemanticMetadata(expected) => {
                        Self::from_shared_bytes_with_catalog_metadata(
                            source.into(),
                            limits,
                            Some(expected),
                        )
                    },
                    PathProfile::SemanticMetadataAny => {
                        // Preserve the existing packaged-file capability on
                        // Windows, whose path adapter cannot promise the
                        // semantic profile's no-reparse capture invariant.
                        #[cfg(windows)]
                        return Self::from_shared_bytes_with_limits(source.into(), limits);
                        #[cfg(not(windows))]
                        Self::from_shared_bytes_with_catalog_metadata(source.into(), limits, None)
                    },
                    PathProfile::IndexOnly | PathProfile::Properties => {
                        Self::from_shared_bytes_with_limits(source.into(), limits)
                    },
                }
            },
            Kind::Dir => {
                let archive = archive_limits(limits)?;
                let directory = match profile {
                    PathProfile::IndexOnly => {
                        FrozenDirectoryBundle::open_with_limits(path, archive)
                    },
                    PathProfile::Properties => {
                        FrozenDirectoryBundle::open_with_properties(path, archive)
                    },
                    PathProfile::SemanticMetadata(expected) => {
                        FrozenDirectoryBundle::open_with_semantic_metadata_when(
                            path,
                            archive,
                            |components, markers| {
                                matches!(component_catalog(components), Ok(Some(format)) if format == expected)
                                    && (marker_outcome(markers) == Outcome::None
                                        || matches!(
                                            marker_outcome(markers),
                                            Outcome::Found(format) if format == expected
                                        ))
                            },
                        )
                    },
                    PathProfile::SemanticMetadataAny => {
                        FrozenDirectoryBundle::open_with_semantic_metadata_when(
                            path,
                            archive,
                            |components, markers| {
                                component_catalog(components).is_ok_and(|format| format.is_some())
                                    && (marker_outcome(markers) == Outcome::None
                                        || matches!(marker_outcome(markers), Outcome::Found(_)))
                            },
                        )
                    },
                }
                .map_err(map_archive_error)?;
                Self::from_directory(
                    directory,
                    limits,
                    matches!(
                        profile,
                        PathProfile::SemanticMetadata(_) | PathProfile::SemanticMetadataAny
                    ),
                )
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

    fn format_only(
        format: Format,
        limits: Limits,
        archive_limits: litchi_iwa_archive::Limits,
    ) -> Self {
        Self {
            backing: PreparedBacking::FormatOnly {
                limits: archive_limits,
            },
            format,
            limits,
        }
    }

    fn from_directory(
        directory: FrozenDirectoryBundle,
        limits: Limits,
        semantic_profile: bool,
    ) -> Result<Option<Self>> {
        let Some(format) = component_catalog(directory.components())? else {
            if marker_outcome(directory.markers()) != Outcome::None {
                return Err(Error::InvalidFormat(
                    "iWork directory marker has no canonical application root".to_owned(),
                ));
            }
            return Ok(None);
        };
        if semantic_profile {
            directory
                .components()
                .__validate_semantic_object_limit()
                .map_err(map_archive_error)?;
        }
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
        let archive_limits = directory.limits();
        let (components, sidecars) = directory.into_semantic_metadata_parts();
        Ok(Some(Self {
            backing: PreparedBacking::Semantic {
                limits: archive_limits,
                components,
                sidecars: sidecars.into(),
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
            PreparedBacking::FormatOnly { .. } | PreparedBacking::Semantic { .. } => None,
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
            PreparedBacking::FormatOnly { limits } => {
                (Arc::new(ComponentCatalog::__empty()), limits)
            },
            PreparedBacking::Semantic {
                components, limits, ..
            } => (components, limits),
        }
    }

    /// Consume a prepared source into semantic components and the optional
    /// frozen canonical `Metadata/Properties.plist` diagnostic.
    ///
    /// This unstable handoff never yields exact package provenance. The
    /// returned sidecar is the only non-IWA directory member captured by the
    /// semantic directory adapter.
    ///
    /// # Errors
    ///
    /// Returns an allocation error if the canonical packaged sidecar cannot
    /// be copied into immutable ownership before the package catalog is
    /// released.
    #[doc(hidden)]
    pub fn __into_semantic_parts(
        self,
    ) -> Result<(
        Arc<ComponentCatalog>,
        litchi_iwa_archive::Limits,
        Option<Arc<[u8]>>,
    )> {
        match self.backing {
            PreparedBacking::Package(catalog) => {
                let limits = catalog.limits();
                let properties = catalog
                    .package()
                    .iter()
                    .find(|entry| entry.name() == "Metadata/Properties.plist")
                    .map(|entry| {
                        if entry.is_opaque() {
                            return Err(Error::InvalidFormat(
                                "canonical Keynote properties use unsupported compression"
                                    .to_owned(),
                            ));
                        }
                        if entry.data().len() > MAX_PROPERTIES_BYTES {
                            return Err(Error::LimitExceeded {
                                kind: LimitKind::EntryBytes,
                                observed: u64::try_from(entry.data().len()).unwrap_or(u64::MAX),
                                maximum: MAX_PROPERTIES_BYTES as u64,
                            });
                        }
                        copy_semantic_sidecar(entry.data())
                    })
                    .transpose()?;
                Ok((Arc::new(catalog.into_components()), limits, properties))
            },
            PreparedBacking::FormatOnly { limits } => {
                Ok((Arc::new(ComponentCatalog::__empty()), limits, None))
            },
            PreparedBacking::Semantic {
                components,
                limits,
                sidecars,
            } => Ok((components, limits, sidecars.properties_plist)),
        }
    }

    /// Consume a Pages-prepared source into one typed archive-free semantic
    /// handoff.
    ///
    /// For a packaged source, this copies only Pages' three exact canonical
    /// metadata authorities before dropping the physical catalog. Selected
    /// entries using an unsupported compression method are rejected rather
    /// than treated as decoded metadata.
    ///
    /// # Errors
    ///
    /// Returns an allocation error if a selected packaged diagnostic cannot be
    /// copied into immutable ownership, or a format error for an opaque
    /// selected authority.
    #[doc(hidden)]
    pub fn __into_pages_semantic_source(self) -> Result<PreparedSemanticSource> {
        self.into_semantic_source(Format::Pages)
    }

    /// Consume a Numbers-prepared source into one typed archive-free semantic
    /// handoff.
    ///
    /// # Errors
    ///
    /// Returns a format error unless this source was classified as Numbers,
    /// or an allocation/physical error while finalizing the fixed sidecars.
    #[doc(hidden)]
    pub fn __into_numbers_semantic_source(self) -> Result<PreparedSemanticSource> {
        self.into_semantic_source(Format::Numbers)
    }

    fn into_semantic_source(self, expected: Format) -> Result<PreparedSemanticSource> {
        if self.format != expected {
            return Err(Error::InvalidFormat(format!(
                "prepared iWork source is not a {expected:?} document"
            )));
        }
        match self.backing {
            PreparedBacking::Package(catalog) => {
                let limits = catalog.limits();
                let sidecars = copy_semantic_metadata_sidecars(catalog.package(), expected)?;
                Ok(PreparedSemanticSource {
                    components: Arc::new(catalog.into_components()),
                    limits,
                    sidecars,
                })
            },
            PreparedBacking::FormatOnly { .. } => Err(Error::InvalidFormat(format!(
                "prepared iWork source is not a {expected:?} document"
            ))),
            PreparedBacking::Semantic {
                components,
                limits,
                sidecars,
            } => Ok(PreparedSemanticSource {
                components,
                limits,
                sidecars,
            }),
        }
    }
}

const SEMANTIC_METADATA_AUTHORITIES: [(&str, SemanticSidecar); 3] = [
    ("Metadata/Properties.plist", SemanticSidecar::Properties),
    (
        "Metadata/BuildVersionHistory.plist",
        SemanticSidecar::BuildVersionHistory,
    ),
    (
        "Metadata/DocumentIdentifier",
        SemanticSidecar::DocumentIdentifier,
    ),
];

#[derive(Clone, Copy)]
enum SemanticSidecar {
    Properties,
    BuildVersionHistory,
    DocumentIdentifier,
}

fn copy_semantic_metadata_sidecars(
    package: &litchi_iwa_archive::package::Catalog,
    format: Format,
) -> Result<PreparedMetadataSidecars> {
    let mut sidecars = PreparedMetadataSidecars::default();
    let selected = package
        .__semantic_metadata_sidecars()
        .map_err(map_archive_error)?;
    for (authority, sidecar, entry) in [
        (
            SEMANTIC_METADATA_AUTHORITIES[0].0,
            SemanticSidecar::Properties,
            selected.properties_plist(),
        ),
        (
            SEMANTIC_METADATA_AUTHORITIES[1].0,
            SemanticSidecar::BuildVersionHistory,
            selected.build_version_history_plist(),
        ),
        (
            SEMANTIC_METADATA_AUTHORITIES[2].0,
            SemanticSidecar::DocumentIdentifier,
            selected.document_identifier(),
        ),
    ] {
        let Some(entry) = entry else {
            continue;
        };
        if entry.is_opaque() {
            return Err(Error::InvalidFormat(format!(
                "canonical {format:?} metadata authority {authority} uses unsupported compression"
            )));
        }
        if entry.data().len() > MAX_PROPERTIES_BYTES {
            return Err(Error::LimitExceeded {
                kind: LimitKind::EntryBytes,
                observed: u64::try_from(entry.data().len()).unwrap_or(u64::MAX),
                maximum: MAX_PROPERTIES_BYTES as u64,
            });
        }
        let data = Some(copy_semantic_sidecar(entry.data())?);
        match sidecar {
            SemanticSidecar::Properties => sidecars.properties_plist = data,
            SemanticSidecar::BuildVersionHistory => sidecars.build_version_history_plist = data,
            SemanticSidecar::DocumentIdentifier => sidecars.document_identifier = data,
        }
    }
    Ok(sidecars)
}

fn copy_semantic_sidecar(source: &[u8]) -> Result<Arc<[u8]>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(source.len())
        .map_err(|_error| Error::Allocation {
            amount: source.len(),
        })?;
    copy.extend_from_slice(source);
    Ok(Arc::from(copy.into_boxed_slice()))
}

fn marker_outcome(markers: DirectoryMarkers) -> Outcome {
    classify(markers.pages(), markers.keynote(), markers.numbers())
}

fn read_stable_package_file(path: &Path, limits: Limits) -> Result<Vec<u8>> {
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
    Ok(bytes)
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
    classify_application_payload(payload).ok().flatten()
}

/// Strictly classify a root `DocumentArchive` payload for detector consumers.
///
/// Unlike [`detect_application_from_document`], this preserves malformed-wire
/// and resource-limit failures as typed detector errors.
#[doc(hidden)]
pub fn __classify_application_payload(payload: &[u8]) -> Result<Option<Format>> {
    classify_application_payload(payload)
}

fn classify_application_payload(payload: &[u8]) -> Result<Option<Format>> {
    let fields = parse_canonical_wire_fields(payload)?;
    let pages = match unique_field(payload, &fields, 15, 2)? {
        Some(shared) => valid_shared_document(shared)?,
        None => false,
    };
    let mut numbers = true;
    for field in [4, 5, 6] {
        numbers &= match unique_field(payload, &fields, field, 2)? {
            Some(reference) => valid_reference(reference)?,
            None => false,
        };
    }
    numbers &= match unique_field(payload, &fields, 8, 2)? {
        Some(shared) => valid_shared_document(shared)?,
        None => false,
    };
    let keynote = match unique_field(payload, &fields, 2, 2)? {
        Some(reference) => valid_reference(reference)?,
        None => false,
    } && match unique_field(payload, &fields, 3, 2)? {
        Some(shared) => valid_shared_document(shared)?,
        None => false,
    };

    let resolved = match (pages, numbers, keynote) {
        (true, false, false) => Some(Format::Pages),
        (false, true, false) => Some(Format::Numbers),
        (false, false, true) => Some(Format::Keynote),
        _ => None,
    };
    if resolved.is_none() && !pages && !numbers && !keynote {
        validate_malformed_application_authority(payload, &fields)?;
    }
    Ok(resolved)
}

fn unique_field<'a>(
    payload: &'a [u8],
    fields: &[WireField],
    number: u32,
    wire_type: u8,
) -> Result<Option<&'a [u8]>> {
    let mut matches = fields.iter().filter(|field| field.number() == number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() || field.wire_type() != wire_type {
        return Ok(None);
    }
    Ok(Some(field.checked_payload(payload)?))
}

fn valid_reference(payload: &[u8]) -> Result<bool> {
    let fields = parse_canonical_wire_fields(payload)?;
    let Some(identifier) = unique_field(payload, &fields, 1, 0)? else {
        return Ok(false);
    };
    Ok(is_canonical_reference_identifier(identifier))
}

fn is_canonical_reference_identifier(payload: &[u8]) -> bool {
    let Ok((identifier, width)) = litchi_iwa_common::decode_varint_from_bytes(payload) else {
        return false;
    };
    width == payload.len() && width == litchi_iwa_common::varint::encoded_len(identifier)
}

fn valid_shared_document(payload: &[u8]) -> Result<bool> {
    let fields = parse_canonical_wire_fields(payload)?;
    let Some(document) = unique_field(payload, &fields, 1, 2)? else {
        return Ok(false);
    };
    let _fields = parse_canonical_wire_fields(document)?;
    Ok(true)
}

fn parse_canonical_wire_fields(payload: &[u8]) -> Result<Vec<WireField>> {
    let fields = parse_wire_fields(payload)?;
    for field in &fields {
        field.validate_canonical_framing(payload)?;
    }
    Ok(fields)
}

fn validate_malformed_application_authority(payload: &[u8], fields: &[WireField]) -> Result<()> {
    if fields.iter().any(|field| field.number() == 15) {
        let shared = strict_unique_field(payload, fields, 15, 2)?;
        validate_strict_shared_document(shared)?;
        return Ok(());
    }
    if [4, 5, 6, 8]
        .iter()
        .all(|number| fields.iter().any(|field| field.number() == *number))
    {
        for number in [4, 5, 6] {
            validate_strict_reference(strict_unique_field(payload, fields, number, 2)?)?;
        }
        validate_strict_shared_document(strict_unique_field(payload, fields, 8, 2)?)?;
        return Ok(());
    }
    if [2, 3]
        .iter()
        .all(|number| fields.iter().any(|field| field.number() == *number))
    {
        validate_strict_reference(strict_unique_field(payload, fields, 2, 2)?)?;
        validate_strict_shared_document(strict_unique_field(payload, fields, 3, 2)?)?;
    }
    Ok(())
}

fn strict_unique_field<'a>(
    payload: &'a [u8],
    fields: &[WireField],
    number: u32,
    wire_type: u8,
) -> Result<&'a [u8]> {
    let mut matches = fields.iter().filter(|field| field.number() == number);
    let Some(field) = matches.next() else {
        return Err(invalid_application_authority(number, "is missing"));
    };
    if matches.next().is_some() {
        return Err(invalid_application_authority(number, "is duplicated"));
    }
    if field.wire_type() != wire_type {
        return Err(invalid_application_authority(
            number,
            "has the wrong wire type",
        ));
    }
    Ok(field.checked_payload(payload)?)
}

fn validate_strict_reference(payload: &[u8]) -> Result<()> {
    let fields = parse_canonical_wire_fields(payload)?;
    let identifier = strict_unique_field(payload, &fields, 1, 0)?;
    if !is_canonical_reference_identifier(identifier) {
        return Err(invalid_application_authority(
            1,
            "has a noncanonical reference identifier",
        ));
    }
    Ok(())
}

fn validate_strict_shared_document(payload: &[u8]) -> Result<()> {
    let fields = parse_canonical_wire_fields(payload)?;
    let document = strict_unique_field(payload, &fields, 1, 2)?;
    let _document_fields = parse_canonical_wire_fields(document)?;
    Ok(())
}

fn invalid_application_authority(number: u32, reason: &str) -> Error {
    litchi_iwa_common::Error::InvalidFormat(format!(
        "protobuf field {number} {reason} in an application authority"
    ))
    .into()
}

fn root_format(data: &[u8], limits: Limits) -> Result<Option<Format>> {
    let stream = SnappyStream::decompress_with_limits(data, limits.snappy_limits()?)?;
    let archive = Archive::parse(stream.as_bytes())?;
    root_format_archive(&archive)
}

fn root_format_archive(archive: &Archive) -> Result<Option<Format>> {
    let mut detected = None;

    for message in archive
        .objects
        .iter()
        .filter(|object| object.archive_info.identifier == Some(1))
        .flat_map(|object| &object.messages)
    {
        let Some(format) = classify_application_payload(&message.data)? else {
            continue;
        };
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

    fn append_length_delimited_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        litchi_iwa_common::encode_varint_into(output, (u64::from(number) << 3) | 2);
        litchi_iwa_common::encode_varint_into(output, u64::try_from(value.len()).unwrap());
        output.extend_from_slice(value);
    }

    fn numbers_payload_with_first_reference(first_reference: &[u8]) -> Vec<u8> {
        let second = reference(2).encode_to_vec();
        let third = reference(3).encode_to_vec();
        let shared = shared_document().encode_to_vec();
        let mut payload = Vec::new();
        for (number, value) in [
            (4, first_reference),
            (5, second.as_slice()),
            (6, third.as_slice()),
            (8, shared.as_slice()),
        ] {
            append_length_delimited_field(&mut payload, number, value);
        }
        payload
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

    fn document_package_with_members(format: Format, members: &[(&str, &[u8])]) -> Vec<u8> {
        let root = document(format);
        let mut files = vec![("Index/Document.iwa", root.as_slice())];
        files.extend_from_slice(members);
        package(&files)
    }

    fn legacy_pages_package_with_outer_members(members: &[(&str, &[u8])]) -> Vec<u8> {
        let root = document(Format::Pages);
        let index = package(&[("Document.iwa", root.as_slice())]);
        let physical_names = members
            .iter()
            .map(|(logical_name, _data)| format!("legacy.pages/{logical_name}"))
            .collect::<Vec<_>>();
        let mut entries = vec![("legacy.pages/Index.zip", index.as_slice())];
        entries.extend(
            physical_names
                .iter()
                .zip(members)
                .map(|(physical_name, (_logical_name, data))| (physical_name.as_str(), *data)),
        );
        package(&entries)
    }

    fn patch_zip_member_compression_to_opaque(bytes: &mut [u8], name: &str) {
        const LOCAL_HEADER_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x03, 0x04];
        const CENTRAL_HEADER_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x01, 0x02];
        const UNSUPPORTED_METHOD: [u8; 2] = 99_u16.to_le_bytes();

        let name = name.as_bytes();
        let mut cursor = 0;
        let mut changed = 0;
        while let Some(relative) = bytes[cursor..]
            .windows(name.len())
            .position(|candidate| candidate == name)
        {
            let position = cursor + relative;
            if position >= 30 && bytes[position - 30..position - 26] == LOCAL_HEADER_SIGNATURE {
                bytes[position - 22..position - 20].copy_from_slice(&UNSUPPORTED_METHOD);
                changed += 1;
            } else if position >= 46
                && bytes[position - 46..position - 42] == CENTRAL_HEADER_SIGNATURE
            {
                bytes[position - 36..position - 34].copy_from_slice(&UNSUPPORTED_METHOD);
                changed += 1;
            }
            cursor = position.saturating_add(name.len());
        }
        assert_eq!(
            changed, 2,
            "{name:?} must have one local and one central ZIP record"
        );
    }

    fn patch_zip_member_local_name(bytes: &mut [u8], name: &str, replacement: &str) {
        assert_eq!(name.len(), replacement.len());
        const LOCAL_HEADER_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x03, 0x04];
        let position = bytes
            .windows(name.len())
            .position(|candidate| candidate == name.as_bytes())
            .expect("test ZIP contains requested local member");
        assert!(position >= 30);
        assert_eq!(bytes[position - 30..position - 26], LOCAL_HEADER_SIGNATURE);
        bytes[position..position + name.len()].copy_from_slice(replacement.as_bytes());
    }

    fn patch_zip_member_raw_names(bytes: &mut [u8], name: &str, replacement: &str) {
        assert_eq!(name.len(), replacement.len());
        let mut cursor = 0;
        let mut changed = 0;
        while let Some(relative) = bytes[cursor..]
            .windows(name.len())
            .position(|candidate| candidate == name.as_bytes())
        {
            let position = cursor + relative;
            bytes[position..position + name.len()].copy_from_slice(replacement.as_bytes());
            changed += 1;
            cursor = position + name.len();
        }
        assert_eq!(changed, 2, "test member must have local and central names");
    }

    fn patch_zip_member_local_compression(bytes: &mut [u8], name: &str, method: u16) {
        const LOCAL_HEADER_SIGNATURE: [u8; 4] = [0x50, 0x4b, 0x03, 0x04];
        let position = bytes
            .windows(name.len())
            .position(|candidate| candidate == name.as_bytes())
            .expect("test ZIP contains requested local member");
        assert!(position >= 30);
        assert_eq!(bytes[position - 30..position - 26], LOCAL_HEADER_SIGNATURE);
        bytes[position - 22..position - 20].copy_from_slice(&method.to_le_bytes());
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
    fn strict_document_payload_classifier_preserves_typed_outcomes() {
        let pages = document_payload(Format::Pages);
        assert_eq!(
            __classify_application_payload(&pages).unwrap(),
            Some(Format::Pages)
        );

        let mut duplicate = pages.clone();
        duplicate.extend_from_slice(&pages);
        assert!(matches!(
            __classify_application_payload(&duplicate),
            Err(Error::IwaCommon(litchi_iwa_common::Error::InvalidFormat(_)))
        ));

        assert!(matches!(
            __classify_application_payload(&[0x78, 0x00]),
            Err(Error::IwaCommon(litchi_iwa_common::Error::InvalidFormat(_)))
        ));
        assert!(matches!(
            __classify_application_payload(&[0x80]),
            Err(Error::IwaCommon(litchi_iwa_common::Error::InvalidFormat(_)))
        ));
        assert!(matches!(
            __classify_application_payload(&[0xfa, 0x00, 0x00]),
            Err(Error::IwaCommon(litchi_iwa_common::Error::InvalidFormat(_)))
        ));
    }

    #[test]
    fn strict_document_payload_classifier_reports_wire_field_limit() {
        let payload =
            [0x08, 0x00].repeat(litchi_iwa_common::WireLimits::default().max_fields() + 1);
        assert!(matches!(
            __classify_application_payload(&payload),
            Err(Error::LimitExceeded {
                kind: LimitKind::IwaFields,
                observed,
                maximum,
            }) if observed == maximum + 1
        ));
    }

    #[test]
    fn strict_document_payload_classifier_validates_nested_reference_identifiers() {
        for invalid in [
            &[0x08, 0x81, 0x00][..],
            &[0x08, 0x01, 0x08, 0x02][..],
            &[0x0a, 0x01, 0x01][..],
        ] {
            assert!(matches!(
                __classify_application_payload(&numbers_payload_with_first_reference(invalid)),
                Err(Error::IwaCommon(litchi_iwa_common::Error::InvalidFormat(_)))
            ));
        }

        assert_eq!(
            __classify_application_payload(&numbers_payload_with_first_reference(&[0x08, 0x00]))
                .unwrap(),
            Some(Format::Numbers),
            "zero is a canonical TSP.Reference encoding used by generated roots"
        );
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
    fn semantic_prepared_source_skips_keynote_media_and_retains_exact_source() {
        let mut bytes = document_package_with_members(
            Format::Keynote,
            &[
                ("Metadata/Properties.plist", b"properties"),
                ("Data/image.png", b"opaque media"),
            ],
        );
        patch_zip_member_compression_to_opaque(&mut bytes, "Data/image.png");
        let source: Arc<[u8]> = bytes.into();

        let prepared = PreparedSource::__from_shared_bytes_with_semantic_metadata(
            Arc::clone(&source),
            Limits::default(),
        )
        .expect("Keynote semantic source should admit opaque unrelated media")
        .expect("synthetic Keynote source should classify");
        assert_eq!(prepared.format(), Format::Keynote);

        let catalog = prepared
            .__into_source_catalog()
            .expect("semantic ZIP preparation retains exact source ownership");
        assert!(Arc::ptr_eq(&source, &catalog.shared_source()));
        assert_eq!(
            catalog
                .package()
                .iter()
                .map(litchi_iwa_archive::package::Entry::name)
                .collect::<Vec<_>>(),
            ["Index/Document.iwa", "Metadata/Properties.plist"]
        );
        let mut exact = Vec::new();
        catalog
            .package()
            .write_to(&mut exact)
            .expect("selective preparation must preserve exact source bytes");
        assert_eq!(exact, source.as_ref());

        let full = SourceCatalog::from_shared_bytes(catalog.shared_source())
            .expect("an explicit full package request should materialize media");
        assert!(
            full.package()
                .iter()
                .any(|entry| entry.name() == "Data/image.png")
        );
        assert!(
            full.package()
                .iter()
                .find(|entry| entry.name() == "Data/image.png")
                .is_some_and(litchi_iwa_archive::package::Entry::is_opaque)
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
    fn pages_semantic_handoff_copies_only_the_three_canonical_sidecars() {
        let properties = b"canonical-pages-properties";
        let history = b"canonical-pages-history";
        let identifier = b"canonical-pages-identifier";
        let bytes: Arc<[u8]> = document_package_with_members(
            Format::Pages,
            &[
                ("Metadata/Properties.plist", properties),
                ("Metadata/BuildVersionHistory.plist", history),
                ("Metadata/DocumentIdentifier", identifier),
                ("Metadata/Properties.plist.bak", b"properties-decoy"),
                ("Metadata/BuildVersionHistory.plist~", b"history-decoy"),
                ("Metadata/DocumentIdentifier.txt", b"identifier-decoy"),
                ("Decoy/Properties.plist", b"basename-decoy"),
            ],
        )
        .into();

        let prepared = PreparedSource::from_shared_bytes(Arc::clone(&bytes))
            .unwrap()
            .expect("fixture has a Pages root");
        let semantic = prepared
            .__into_pages_semantic_source()
            .expect("canonical Pages sidecars are copied");
        assert_eq!(
            Arc::strong_count(&bytes),
            1,
            "the archive-free Pages consumer must not retain exact-package bytes"
        );
        let sidecars = semantic.sidecars();
        assert_eq!(sidecars.properties_plist(), Some(properties.as_slice()));
        assert_eq!(
            sidecars.build_version_history_plist(),
            Some(history.as_slice())
        );
        assert_eq!(sidecars.document_identifier(), Some(identifier.as_slice()));
    }

    #[test]
    fn pages_zip_profile_enforces_modern_canonical_boundaries_before_handoff() {
        for authority in SEMANTIC_METADATA_AUTHORITIES.map(|(authority, _sidecar)| authority) {
            let exact = vec![b'x'; MAX_PROPERTIES_BYTES];
            let accepted = document_package_with_members(Format::Pages, &[(authority, &exact)]);
            let prepared =
                PreparedSource::__from_bytes_with_pages_metadata(&accepted, Limits::default())
                    .unwrap()
                    .expect("fixture has a Pages root");
            let sidecars = prepared
                .__into_pages_semantic_source()
                .expect("exactly 64 KiB physical authority is admitted")
                .__into_parts()
                .2;
            let captured = match authority {
                "Metadata/Properties.plist" => sidecars.properties_plist(),
                "Metadata/BuildVersionHistory.plist" => sidecars.build_version_history_plist(),
                "Metadata/DocumentIdentifier" => sidecars.document_identifier(),
                _ => unreachable!("test iterates fixed Pages metadata authorities"),
            };
            assert_eq!(captured, Some(exact.as_slice()));

            let one_over = vec![b'x'; MAX_PROPERTIES_BYTES + 1];
            let rejected = document_package_with_members(Format::Pages, &[(authority, &one_over)]);
            assert!(matches!(
                PreparedSource::__from_bytes_with_pages_metadata(
                    &rejected,
                    Limits::default(),
                ),
                Err(Error::LimitExceeded {
                    kind: LimitKind::EntryBytes,
                    observed,
                    maximum,
                }) if observed == one_over.len() as u64
                    && maximum == MAX_PROPERTIES_BYTES as u64
            ));
        }
    }

    #[test]
    fn pages_zip_profile_normalizes_legacy_outer_canonical_boundaries() {
        for authority in SEMANTIC_METADATA_AUTHORITIES.map(|(authority, _sidecar)| authority) {
            let exact = vec![b'x'; MAX_PROPERTIES_BYTES];
            let accepted = legacy_pages_package_with_outer_members(&[(authority, &exact)]);
            PreparedSource::__from_bytes_with_pages_metadata(&accepted, Limits::default())
                .unwrap()
                .expect("legacy fixture has a Pages root")
                .__into_pages_semantic_source()
                .expect("exact legacy outer authority is admitted");

            let one_over = vec![b'x'; MAX_PROPERTIES_BYTES + 1];
            let rejected = legacy_pages_package_with_outer_members(&[(authority, &one_over)]);
            assert!(matches!(
                PreparedSource::__from_bytes_with_pages_metadata(
                    &rejected,
                    Limits::default(),
                ),
                Err(Error::LimitExceeded {
                    kind: LimitKind::EntryBytes,
                    observed,
                    maximum,
                }) if observed == one_over.len() as u64
                    && maximum == MAX_PROPERTIES_BYTES as u64
            ));
        }
    }

    #[test]
    fn pages_zip_profile_ignores_near_names_and_unprefixed_legacy_decoys() {
        let oversized = vec![b'x'; MAX_PROPERTIES_BYTES + 1];
        let near_names = document_package_with_members(
            Format::Pages,
            &[
                ("Metadata/Properties.plist.bak", &oversized),
                ("Metadata/BuildVersionHistory.plist~", &oversized),
                ("Metadata/DocumentIdentifier.txt", &oversized),
                ("Decoy/Properties.plist", &oversized),
            ],
        );
        let sidecars =
            PreparedSource::__from_bytes_with_pages_metadata(&near_names, Limits::default())
                .unwrap()
                .expect("near names do not affect Pages classification")
                .__into_pages_semantic_source()
                .unwrap()
                .__into_parts()
                .2;
        assert!(sidecars.properties_plist().is_none());
        assert!(sidecars.build_version_history_plist().is_none());
        assert!(sidecars.document_identifier().is_none());

        let root = document(Format::Pages);
        let index = package(&[("Document.iwa", root.as_slice())]);
        let legacy = package(&[
            ("legacy.pages/Index.zip", index.as_slice()),
            (
                "legacy.pages/Metadata/DocumentIdentifier",
                b"first".as_slice(),
            ),
            ("Metadata/DocumentIdentifier", &oversized),
        ]);
        let sidecars = PreparedSource::__from_bytes_with_pages_metadata(&legacy, Limits::default())
            .expect("unprefixed outer metadata is outside the legacy authority")
            .expect("legacy fixture has a Pages root")
            .__into_pages_semantic_source()
            .expect("only the explicitly prefixed authority is selected")
            .__into_parts()
            .2;
        assert_eq!(sidecars.document_identifier(), Some(b"first".as_slice()));
    }

    #[test]
    fn pages_zip_profile_requires_exact_raw_authority_names() {
        let authority = "Metadata/Properties.plist";
        let near_name = r"Metadata\Properties.plist";
        let mut bytes =
            document_package_with_members(Format::Pages, &[(authority, b"properties-decoy")]);
        // The test writer normalizes paths, so construct hostile raw headers
        // after serialization instead of trusting its input spelling.
        patch_zip_member_raw_names(&mut bytes, authority, near_name);
        let sidecars = PreparedSource::__from_bytes_with_pages_metadata(&bytes, Limits::default())
            .expect("raw near-name is outside the Pages metadata profile")
            .expect("fixture has a Pages root")
            .__into_pages_semantic_source()
            .expect("raw near-name is ignored at handoff")
            .__into_parts()
            .2;
        assert_eq!(sidecars.properties_plist(), None);
    }

    #[test]
    fn pages_zip_profile_refuses_selected_local_central_mismatches() {
        let authority = "Metadata/Properties.plist";
        let replacement = "Metadata/Properties.plisX";
        let mut name_mismatch =
            document_package_with_members(Format::Pages, &[(authority, b"properties")]);
        patch_zip_member_local_name(&mut name_mismatch, authority, replacement);
        assert!(matches!(
            PreparedSource::__from_bytes_with_pages_metadata(
                &name_mismatch,
                Limits::default(),
            ),
            Err(Error::Archive(message))
                if message.contains("mismatched local and central names")
        ));

        let mut method_mismatch =
            document_package_with_members(Format::Pages, &[(authority, b"properties")]);
        patch_zip_member_local_compression(&mut method_mismatch, authority, 99);
        assert!(matches!(
            PreparedSource::__from_bytes_with_pages_metadata(
                &method_mismatch,
                Limits::default(),
            ),
            Err(Error::Archive(message))
                if message.contains("mismatched local and central compression methods")
        ));
    }

    #[test]
    fn pages_zip_profile_refuses_opaque_canonical_authorities_during_preparation() {
        for authority in SEMANTIC_METADATA_AUTHORITIES.map(|(authority, _sidecar)| authority) {
            let mut bytes = document_package_with_members(
                Format::Pages,
                &[(authority, b"canonical-pages-sidecar")],
            );
            patch_zip_member_compression_to_opaque(&mut bytes, authority);

            assert!(matches!(
                PreparedSource::__from_bytes_with_pages_metadata(&bytes, Limits::default()),
                Err(Error::Archive(message))
                    if message.contains(authority)
                        && message.contains("unsupported ZIP compression")
            ));
        }
    }

    #[test]
    fn pages_zip_profile_applies_metadata_policy_only_after_pages_classification() {
        let oversized = vec![b'x'; MAX_PROPERTIES_BYTES + 1];
        let mut keynote = document_package_with_members(
            Format::Keynote,
            &[("Metadata/BuildVersionHistory.plist", &oversized)],
        );
        patch_zip_member_compression_to_opaque(&mut keynote, "Metadata/BuildVersionHistory.plist");

        let prepared =
            PreparedSource::__from_bytes_with_pages_metadata(&keynote, Limits::default())
                .expect("non-Pages metadata remains outside the Pages ZIP profile")
                .expect("fixture has a Keynote root");
        assert_eq!(prepared.format(), Format::Keynote);
    }

    #[test]
    fn pages_shared_and_regular_file_profiles_use_the_same_zip_preflight() -> std::io::Result<()> {
        let one_over = vec![b'x'; MAX_PROPERTIES_BYTES + 1];
        let bytes: Arc<[u8]> = document_package_with_members(
            Format::Pages,
            &[("Metadata/Properties.plist", &one_over)],
        )
        .into();
        assert!(matches!(
            PreparedSource::__from_shared_bytes_with_pages_metadata(
                Arc::clone(&bytes),
                Limits::default(),
            ),
            Err(Error::LimitExceeded {
                kind: LimitKind::EntryBytes,
                ..
            })
        ));

        let temp = Temp::new()?;
        let path = temp.0.join("oversized.pages");
        fs::write(&path, &bytes)?;
        assert!(matches!(
            PreparedSource::__from_path_with_pages_metadata(&path, Limits::default()),
            Err(Error::LimitExceeded {
                kind: LimitKind::EntryBytes,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn keynote_semantic_handoff_remains_properties_only() {
        let properties = b"canonical-keynote-properties";
        let history = b"opaque-history-must-not-affect-keynote";
        let identifier = b"opaque-identifier-must-not-affect-keynote";
        let mut bytes = document_package_with_members(
            Format::Keynote,
            &[
                ("Metadata/Properties.plist", properties),
                ("Metadata/BuildVersionHistory.plist", history),
                ("Metadata/DocumentIdentifier", identifier),
            ],
        );
        patch_zip_member_compression_to_opaque(&mut bytes, "Metadata/BuildVersionHistory.plist");
        patch_zip_member_compression_to_opaque(&mut bytes, "Metadata/DocumentIdentifier");

        let prepared = PreparedSource::from_bytes(&bytes)
            .unwrap()
            .expect("fixture has a Keynote root");
        let (_components, _limits, captured) = prepared
            .__into_semantic_parts()
            .expect("Keynote must ignore non-properties Pages sidecars");
        assert_eq!(captured.as_deref(), Some(properties.as_slice()));
    }

    #[test]
    fn pages_semantic_handoff_rejects_opaque_canonical_sidecars() {
        for authority in SEMANTIC_METADATA_AUTHORITIES.map(|(authority, _sidecar)| authority) {
            let mut bytes = document_package_with_members(
                Format::Pages,
                &[(authority, b"canonical-pages-sidecar")],
            );
            patch_zip_member_compression_to_opaque(&mut bytes, authority);

            let prepared = PreparedSource::from_bytes(&bytes)
                .unwrap()
                .expect("fixture has a Pages root");
            let error = prepared
                .__into_pages_semantic_source()
                .expect_err("opaque canonical Pages metadata must fail closed");
            assert!(matches!(
                error,
                Error::InvalidFormat(message)
                    if message.contains(authority) && message.contains("unsupported compression")
            ));
        }
    }

    #[test]
    fn pages_semantic_handoff_enforces_the_exact_per_sidecar_ceiling() {
        for authority in SEMANTIC_METADATA_AUTHORITIES.map(|(authority, _sidecar)| authority) {
            let exact = vec![b'x'; MAX_PROPERTIES_BYTES];
            let accepted = document_package_with_members(Format::Pages, &[(authority, &exact)]);
            let prepared = PreparedSource::from_bytes(&accepted)
                .unwrap()
                .expect("fixture has a Pages root");
            let sidecars = prepared
                .__into_pages_semantic_source()
                .expect("exactly 64 KiB sidecar must be accepted")
                .__into_parts()
                .2;
            let captured = match authority {
                "Metadata/Properties.plist" => sidecars.properties_plist(),
                "Metadata/BuildVersionHistory.plist" => sidecars.build_version_history_plist(),
                "Metadata/DocumentIdentifier" => sidecars.document_identifier(),
                _ => unreachable!("test iterates fixed Pages metadata authorities"),
            };
            assert_eq!(captured, Some(exact.as_slice()));

            let oversized = vec![b'x'; MAX_PROPERTIES_BYTES + 1];
            let rejected = document_package_with_members(Format::Pages, &[(authority, &oversized)]);
            let prepared = PreparedSource::from_bytes(&rejected)
                .unwrap()
                .expect("fixture has a Pages root");
            assert!(matches!(
                prepared.__into_pages_semantic_source(),
                Err(Error::LimitExceeded {
                    kind: LimitKind::EntryBytes,
                    observed,
                    maximum,
                }) if observed == oversized.len() as u64
                    && maximum == MAX_PROPERTIES_BYTES as u64
            ));
        }
    }

    #[test]
    fn directory_pages_metadata_profile_is_opt_in_and_exact() -> std::io::Result<()> {
        let temp = Temp::new()?;
        let bundle = temp.0.join("profile.pages");
        let properties = b"directory-properties";
        let history = b"directory-history";
        let identifier = b"directory-identifier";
        fs::create_dir(&bundle)?;
        fs::write(
            bundle.join("Index.zip"),
            document_package(Format::Pages, &[]),
        )?;
        fs::create_dir(bundle.join("Metadata"))?;
        fs::write(bundle.join("Metadata/Properties.plist"), properties)?;
        fs::write(bundle.join("Metadata/BuildVersionHistory.plist"), history)?;
        fs::write(bundle.join("Metadata/DocumentIdentifier"), identifier)?;
        fs::write(bundle.join("Metadata/Properties.plist.bak"), b"decoy")?;

        let generic = PreparedSource::from_path(&bundle)
            .unwrap()
            .expect("fixture has a Pages root")
            .__into_pages_semantic_source()
            .unwrap();
        assert_eq!(generic.sidecars().properties_plist(), None);
        assert_eq!(generic.sidecars().build_version_history_plist(), None);
        assert_eq!(generic.sidecars().document_identifier(), None);

        let properties_only =
            PreparedSource::__from_path_with_properties(&bundle, Limits::default())
                .unwrap()
                .expect("fixture has a Pages root");
        let (_components, _limits, captured) = properties_only.__into_semantic_parts().unwrap();
        assert_eq!(captured.as_deref(), Some(properties.as_slice()));

        let pages = PreparedSource::__from_path_with_pages_metadata(&bundle, Limits::default())
            .unwrap()
            .expect("fixture has a Pages root")
            .__into_pages_semantic_source()
            .unwrap();
        assert_eq!(
            pages.sidecars().properties_plist(),
            Some(properties.as_slice())
        );
        assert_eq!(
            pages.sidecars().build_version_history_plist(),
            Some(history.as_slice())
        );
        assert_eq!(
            pages.sidecars().document_identifier(),
            Some(identifier.as_slice())
        );
        Ok(())
    }

    #[test]
    fn directory_pages_profile_applies_sidecar_policy_only_after_pages_classification() -> Result<()>
    {
        let temp = Temp::new()?;
        let keynote = temp.0.join("profile-isolation.key");
        fs::create_dir(&keynote)?;
        fs::write(
            keynote.join("Index.zip"),
            document_package(Format::Keynote, &[]),
        )?;
        fs::create_dir_all(keynote.join("Metadata/BuildVersionHistory.plist"))?;

        let prepared = PreparedSource::__from_path_with_pages_metadata(&keynote, Limits::default())
            .expect("non-Pages sidecars must be outside the Pages profile")
            .expect("fixture has a Keynote root");
        assert_eq!(prepared.format(), Format::Keynote);

        let pages = temp.0.join("profile-strict.pages");
        fs::create_dir(&pages)?;
        fs::write(
            pages.join("Index.zip"),
            document_package(Format::Pages, &[]),
        )?;
        fs::create_dir_all(pages.join("Metadata/BuildVersionHistory.plist"))?;
        assert!(matches!(
            PreparedSource::__from_path_with_pages_metadata(&pages, Limits::default()),
            Err(Error::Archive(message))
                if message.contains("BuildVersionHistory.plist")
                    && message.contains("not a regular file")
        ));
        Ok(())
    }

    #[cfg(windows)]
    #[test]
    fn pages_metadata_path_profile_fails_closed_without_pinned_windows_identity()
    -> std::io::Result<()> {
        let temp = Temp::new()?;
        let package = temp.0.join("document.pages");
        fs::write(&package, document_package(Format::Pages, &[]))?;
        assert!(matches!(
            PreparedSource::__from_path_with_pages_metadata(&package, Limits::default()),
            Err(Error::InvalidFormat(message))
                if message == "Pages package path ingress is unsupported on Windows"
        ));

        let bytes = document_package(Format::Pages, &[]);
        assert!(
            PreparedSource::__from_bytes_with_pages_metadata(&bytes, Limits::default())
                .unwrap()
                .is_some()
        );
        Ok(())
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

    #[test]
    fn numbers_semantic_ingress_hands_off_only_exact_sidecars_and_releases_shared_input() {
        let properties = b"numbers-properties";
        let history = b"numbers-history";
        let identifier = b"numbers-identifier";
        let bytes: Arc<[u8]> = document_package_with_members(
            Format::Numbers,
            &[
                ("Metadata/Properties.plist", properties),
                ("Metadata/BuildVersionHistory.plist", history),
                ("Metadata/DocumentIdentifier", identifier),
                ("Metadata/Properties.plist.bak", b"properties-decoy"),
                ("Metadata/DocumentIdentifier.txt", b"identifier-decoy"),
            ],
        )
        .into();

        let prepared = PreparedSource::__from_shared_bytes_with_numbers_semantics(
            Arc::clone(&bytes),
            Limits::default(),
        )
        .unwrap()
        .expect("fixture has a Numbers root");
        assert_eq!(
            Arc::strong_count(&bytes),
            1,
            "Numbers preparation must not retain the shared package snapshot"
        );
        let shared = prepared
            .__into_numbers_semantic_source()
            .expect("Numbers sidecars hand off");
        assert_eq!(
            Arc::strong_count(&bytes),
            1,
            "the semantic Numbers handoff must release its package snapshot"
        );
        assert_eq!(
            shared.sidecars().properties_plist(),
            Some(properties.as_slice())
        );
        assert_eq!(
            shared.sidecars().build_version_history_plist(),
            Some(history.as_slice())
        );
        assert_eq!(
            shared.sidecars().document_identifier(),
            Some(identifier.as_slice())
        );

        let borrowed =
            PreparedSource::__from_bytes_with_numbers_semantics(&bytes, Limits::default())
                .unwrap()
                .expect("borrowed Numbers input is accepted")
                .__into_numbers_semantic_source()
                .expect("borrowed Numbers source hands off");
        assert_eq!(
            borrowed.sidecars().properties_plist(),
            Some(properties.as_slice())
        );
    }

    #[test]
    fn numbers_zip_profile_enforces_exact_raw_sidecar_boundaries() {
        for authority in SEMANTIC_METADATA_AUTHORITIES.map(|(authority, _)| authority) {
            let exact = vec![b'x'; MAX_PROPERTIES_BYTES];
            let accepted = document_package_with_members(Format::Numbers, &[(authority, &exact)]);
            PreparedSource::__from_bytes_with_numbers_semantics(&accepted, Limits::default())
                .unwrap()
                .expect("fixture has a Numbers root")
                .__into_numbers_semantic_source()
                .expect("64 KiB canonical Numbers sidecar is accepted");

            let oversized = vec![b'x'; MAX_PROPERTIES_BYTES + 1];
            let rejected =
                document_package_with_members(Format::Numbers, &[(authority, &oversized)]);
            assert!(matches!(
                PreparedSource::__from_bytes_with_numbers_semantics(&rejected, Limits::default()),
                Err(Error::LimitExceeded {
                    kind: LimitKind::EntryBytes,
                    observed,
                    maximum,
                }) if observed == oversized.len() as u64 && maximum == MAX_PROPERTIES_BYTES as u64
            ));
        }

        let authority = "Metadata/Properties.plist";
        let mut near_name =
            document_package_with_members(Format::Numbers, &[(authority, b"numbers-properties")]);
        patch_zip_member_raw_names(&mut near_name, authority, r"Metadata\Properties.plist");
        assert!(matches!(
            PreparedSource::__from_bytes_with_numbers_semantics(&near_name, Limits::default()),
            Err(Error::Archive(message))
                if message.contains("non-canonical or one-sided ZIP names")
        ));
    }

    #[test]
    fn numbers_zip_profile_refuses_selected_opaque_authorities() {
        let authority = "Metadata/BuildVersionHistory.plist";
        let mut bytes = document_package_with_members(
            Format::Numbers,
            &[(authority, b"canonical-numbers-sidecar")],
        );
        patch_zip_member_compression_to_opaque(&mut bytes, authority);
        assert!(matches!(
            PreparedSource::__from_bytes_with_numbers_semantics(&bytes, Limits::default()),
            Err(Error::Archive(message))
                if message.contains(authority) && message.contains("unsupported ZIP compression")
        ));
    }

    #[test]
    fn numbers_zip_profile_leaves_foreign_metadata_uninspected() {
        let oversized = vec![b'x'; MAX_PROPERTIES_BYTES + 1];
        for format in [Format::Pages, Format::Keynote] {
            let mut bytes = document_package_with_members(
                format,
                &[("Metadata/BuildVersionHistory.plist", &oversized)],
            );
            patch_zip_member_compression_to_opaque(
                &mut bytes,
                "Metadata/BuildVersionHistory.plist",
            );
            let prepared =
                PreparedSource::__from_bytes_with_numbers_semantics(&bytes, Limits::default())
                    .expect("foreign metadata must remain outside Numbers profile")
                    .expect("fixture has a foreign iWork root");
            assert_eq!(prepared.format(), format);
        }
    }

    #[test]
    fn directory_numbers_semantic_profile_is_conditional_and_exact() -> Result<()> {
        let temp = Temp::new()?;
        let numbers = temp.0.join("profile.numbers");
        fs::create_dir(&numbers)?;
        fs::write(
            numbers.join("Index.zip"),
            document_package(Format::Numbers, &[]),
        )?;
        fs::create_dir(numbers.join("Metadata"))?;
        fs::write(numbers.join("Metadata/Properties.plist"), b"properties")?;
        fs::write(
            numbers.join("Metadata/BuildVersionHistory.plist"),
            b"history",
        )?;
        fs::write(numbers.join("Metadata/DocumentIdentifier"), b"identifier")?;
        fs::write(numbers.join("Metadata/Properties.plist.bak"), b"decoy")?;
        let semantic =
            PreparedSource::__from_path_with_numbers_semantics(&numbers, Limits::default())
                .unwrap()
                .expect("Numbers bundle is classified")
                .__into_numbers_semantic_source()
                .unwrap();
        assert_eq!(
            semantic.sidecars().properties_plist(),
            Some(&b"properties"[..])
        );
        assert_eq!(
            semantic.sidecars().build_version_history_plist(),
            Some(&b"history"[..])
        );
        assert_eq!(
            semantic.sidecars().document_identifier(),
            Some(&b"identifier"[..])
        );

        for (format, extension) in [(Format::Pages, "pages"), (Format::Keynote, "key")] {
            let foreign = temp.0.join(format!("foreign.{extension}"));
            fs::create_dir(&foreign)?;
            fs::write(foreign.join("Index.zip"), document_package(format, &[]))?;
            fs::create_dir_all(foreign.join("Metadata/DocumentIdentifier"))?;
            let prepared =
                PreparedSource::__from_path_with_numbers_semantics(&foreign, Limits::default())
                    .expect("foreign directory sidecars must not be captured")
                    .expect("foreign document remains detectable");
            assert_eq!(prepared.format(), format);
        }
        Ok(())
    }

    #[cfg(windows)]
    #[test]
    fn numbers_semantic_path_profile_fails_closed_without_pinned_windows_identity()
    -> std::io::Result<()> {
        let temp = Temp::new()?;
        let package = temp.0.join("document.numbers");
        fs::write(&package, document_package(Format::Numbers, &[]))?;
        assert!(matches!(
            PreparedSource::__from_path_with_numbers_semantics(&package, Limits::default()),
            Err(Error::InvalidFormat(message))
                if message == "Numbers package path ingress is unsupported on Windows"
        ));
        let bytes = document_package(Format::Numbers, &[]);
        assert!(
            PreparedSource::__from_bytes_with_numbers_semantics(&bytes, Limits::default())
                .unwrap()
                .is_some()
        );
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
