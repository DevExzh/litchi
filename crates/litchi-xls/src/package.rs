//! Canonical immutable owner for legacy XLS packages.
//!
//! `Package` is the source-preserving owner used by the unified operations
//! layer.  An eager package retains the parsed raw reader and its original
//! cursor bytes.  A source-backed package retains the validated positional CFB
//! view and the lazy semantic reader.  Neither path reserializes an opened
//! package through [`crate::Builder`] merely to publish an unchanged artifact.
//!
//! The raw, generic reader lives under [`crate::raw`] and is intentionally
//! available only through the explicit `reader_ref` escape on the canonical
//! [`Workbook`] projection.  This keeps format-specific read capability while
//! leaving the package and projection shapes unambiguous for facade callers.

use crate::{
    CompatibilityProfile, Error as ReaderError, Leniency, OpenOptions, SourceBackedError,
    SourceBackedLimits, SourceBackedWorkbook, SourceLimits, ToleranceReport,
};
use litchi_cfb::{OleError, OleFile, SharedOleFile, SharedOleFileLimits};
use litchi_core::sheet::{Cell as CellTrait, CellValue, Worksheet as WorksheetTrait};
use litchi_core::{Durability, Metadata, ReadAt};
use litchi_ole_common::property_set::{PropertySetReader, SharedPropertySetReader};
use std::fmt;
use std::io::{self, Cursor, Read, Write};
use std::path::Path;
use std::sync::{Arc, OnceLock};

const SIGNATURE_ENTRIES: [&str; 4] = [
    "_xmlsignatures",
    "_signatures",
    "DigitalSignature",
    "\u{0005}DigitalSignature",
];

fn has_signature(ole: &OleFile<Cursor<Vec<u8>>>) -> bool {
    SIGNATURE_ENTRIES.iter().any(|name| ole.exists(&[*name]))
}

fn has_shared_signature(ole: &SharedOleFile) -> bool {
    SIGNATURE_ENTRIES.iter().any(|name| ole.exists(&[*name]))
}

/// The parsed eager raw reader retained by the canonical owner.
///
/// The `Reader` name is supplied by the raw-reader migration.  Keeping this
/// alias local makes the owner independent of the public root re-exports.
type RawReader = crate::raw::Reader<Cursor<Vec<u8>>>;

#[path = "package/cells.rs"]
mod cells;
/// The canonical primary transaction for an immutable XLS package.
#[path = "package/edit.rs"]
pub mod edit;
#[path = "package/metadata.rs"]
pub mod metadata;
pub use cells::History;
pub use edit::{Commit, Diagnostics, Edit, EditOptions, LimitKind, Patch};

/// Errors returned by the canonical XLS package and projection APIs.
#[non_exhaustive]
pub enum Error {
    /// A legacy XLS semantic reader rejected the package.
    Reader(ReaderError),
    /// A source-backed semantic read failed.
    Source(SourceBackedError),
    /// A previously attempted source-backed materialization failed.
    CachedSource(Arc<SourceBackedError>),
    /// The retained CFB source or its atomic publisher failed.
    Cfb(OleError),
    /// A shared OLE object or property-set operation failed with typed
    /// refusal/conflict information.
    Object(litchi_ole_common::Error),
    /// Sequential output failed.
    Io(io::Error),
    /// A checked allocation could not be reserved.
    Allocation {
        /// Logical output being allocated.
        resource: &'static str,
        /// Original allocator failure.
        source: std::collections::TryReserveError,
    },
    /// Ordinary publication cannot preserve the active password-to-open
    /// encryption state.  Use [`Package::to_encrypted_bytes`] only when the
    /// caller explicitly wants the original encrypted artifact byte-for-byte.
    EncryptionPolicy,
    /// A requested canonical projection is unavailable for this owner state.
    InvalidData(&'static str),
}

/// Result type for canonical XLS package operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Admit edited bytes under `source`'s retained read policy.
pub(crate) fn republish(source: &Package, bytes: Vec<u8>) -> Result<Package> {
    source.republish_bytes(bytes)
}

impl From<ReaderError> for Error {
    fn from(error: ReaderError) -> Self {
        Self::Reader(error)
    }
}

impl From<SourceBackedError> for Error {
    fn from(error: SourceBackedError) -> Self {
        Self::Source(error)
    }
}

impl From<OleError> for Error {
    fn from(error: OleError) -> Self {
        Self::Cfb(error)
    }
}

impl From<litchi_ole_common::Error> for Error {
    fn from(error: litchi_ole_common::Error) -> Self {
        Self::Object(error)
    }
}

impl From<io::Error> for Error {
    fn from(error: io::Error) -> Self {
        Self::Io(error)
    }
}

impl fmt::Debug for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let kind = match self {
            Self::Reader(_) => "Reader",
            Self::Source(_) => "Source",
            Self::CachedSource(_) => "CachedSource",
            Self::Cfb(_) => "Cfb",
            Self::Object(_) => "Object",
            Self::Io(_) => "Io",
            Self::Allocation { .. } => "Allocation",
            Self::EncryptionPolicy => "EncryptionPolicy",
            Self::InvalidData(_) => "InvalidData",
        };
        formatter.debug_tuple(kind).finish()
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Reader(error) => fmt::Display::fmt(error, formatter),
            Self::Source(error) => fmt::Display::fmt(error, formatter),
            Self::CachedSource(error) => fmt::Display::fmt(error.as_ref(), formatter),
            Self::Cfb(error) => fmt::Display::fmt(error, formatter),
            Self::Object(error) => fmt::Display::fmt(error, formatter),
            Self::Io(error) => fmt::Display::fmt(error, formatter),
            Self::Allocation { resource, source } => {
                write!(formatter, "could not reserve {resource}: {source}")
            },
            Self::EncryptionPolicy => formatter.write_str(
                "ordinary XLS publication is refused for a password-to-open encrypted source",
            ),
            Self::InvalidData(message) => formatter.write_str(message),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Reader(error) => Some(error),
            Self::Source(error) => Some(error),
            Self::CachedSource(error) => Some(error.as_ref()),
            Self::Cfb(error) => Some(error),
            Self::Object(error) => Some(error),
            Self::Io(error) => Some(error),
            Self::Allocation { source, .. } => Some(source),
            Self::EncryptionPolicy | Self::InvalidData(_) => None,
        }
    }
}

/// The immutable state held by one package.
#[derive(Clone)]
enum PackageState {
    /// A fully parsed reader retaining the exact source cursor bytes.
    Eager {
        reader: Arc<RawReader>,
        encryption: Option<crate::EncryptionProfile>,
    },
    /// A validated positional source and its lazy semantic projection.
    Source {
        cfb: Arc<SharedOleFile>,
        workbook: SourceBackedWorkbook,
        /// Source-backed admission currently accepts clear Workbook streams
        /// only, so no password-to-open profile is retained on this path.
        encryption: Option<crate::EncryptionProfile>,
        /// Materialization is deliberately demand-driven.  A source package
        /// can serve source-backed reads without ever constructing this
        /// eager reader.
        materialized: OnceLock<std::result::Result<Arc<RawReader>, Arc<SourceBackedError>>>,
    },
}

/// Policies retained by the owner that admitted its source.
///
/// The eager reader uses the physical CFB limits and compatibility knobs from
/// [`OpenOptions`].  A source-backed handoff retains the source owner's full
/// scan/materialization profile; converting it to the eager profile would
/// silently widen or narrow later edit admission.
#[derive(Debug, Clone, Copy)]
enum RetainedPolicy {
    Eager {
        source_limits: SourceLimits,
        leniency: Leniency,
        compatibility_profile: CompatibilityProfile,
    },
    Source {
        limits: SourceBackedLimits,
        leniency: Leniency,
    },
}

impl fmt::Debug for PackageState {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Eager { encryption, .. } => formatter
                .debug_struct("Eager")
                .field("encryption", encryption)
                .finish_non_exhaustive(),
            Self::Source {
                cfb,
                workbook,
                encryption,
                ..
            } => formatter
                .debug_struct("Source")
                .field("file_size", &cfb.file_size())
                .field("workbook", workbook)
                .field("encryption", encryption)
                .finish_non_exhaustive(),
        }
    }
}

/// An immutable, cheaply cloned legacy XLS package.
pub struct Package {
    inner: Arc<PackageInner>,
}

#[derive(Clone)]
struct PackageInner {
    state: PackageState,
    policy: RetainedPolicy,
    signed: bool,
    metadata: OnceLock<Metadata>,
    /// The complete tab catalog, captured on first catalog read and shared
    /// by package projections and adapters without copying its names.
    catalog: OnceLock<Arc<[crate::SheetMetadata]>>,
    edit_snapshot: Option<crate::cell_values::Snapshot>,
    /// The Workbook stream, read once for worksheet-feature reads.
    workbook_stream: OnceLock<Arc<[u8]>>,
}

impl fmt::Debug for PackageInner {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PackageInner")
            .field("state", &self.state)
            .field("policy", &self.policy)
            .finish_non_exhaustive()
    }
}

impl Clone for Package {
    fn clone(&self) -> Self {
        Self {
            inner: Arc::clone(&self.inner),
        }
    }
}

impl fmt::Debug for Package {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Package")
            .field("source_backed", &self.is_source_backed())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Create a minimal BIFF8 workbook containing one empty `Sheet1` worksheet.
    ///
    /// The authored image is admitted under the same strict default policies as
    /// an opened package. Publication and subsequent edits use this owner.
    pub fn new() -> Result<Self> {
        crate::Builder::new().build()
    }

    /// Open an XLS package from a copied byte slice under default policies.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(bytes, OpenOptions::default())
    }

    /// Open an XLS package from a copied byte slice under explicit policies.
    pub fn from_bytes_with_options(bytes: &[u8], options: OpenOptions<'_>) -> Result<Self> {
        if bytes.len() > options.source_limits_value().max_input_bytes() {
            return Err(Error::Reader(ReaderError::Limit {
                kind: LimitKind::InputBytes,
                observed: u64::try_from(bytes.len()).unwrap_or(u64::MAX),
                maximum: u64::try_from(options.source_limits_value().max_input_bytes())
                    .unwrap_or(u64::MAX),
            }));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "XLS package input",
                source,
            })?;
        owned.extend_from_slice(bytes);
        Self::from_owned_bytes(owned, options)
    }

    /// Read an XLS package from a sequential source.
    pub fn from_reader<R: Read>(mut reader: R) -> Result<Self> {
        Self::from_reader_with_options(&mut reader, OpenOptions::default())
    }

    /// Read a sequential source under explicit policies.
    pub fn from_reader_with_options<R: Read>(
        reader: &mut R,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        let bytes = read_bounded(reader, options.source_limits_value())?;
        Self::from_owned_bytes(bytes, options)
    }

    /// Open a filesystem path under default policies.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open a password-to-open encrypted XLS path.
    ///
    /// The returned package retains the source encryption profile. Ordinary
    /// publication remains protected by the encryption policy; callers must
    /// choose an explicit encrypted-source route when writing it.
    #[cfg(any(unix, windows))]
    pub fn open_with_password(path: impl AsRef<Path>, password: &str) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::new().password(password))
    }

    /// Open a filesystem path under explicit policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions<'_>) -> Result<Self> {
        let mut file = std::fs::File::open(path).map_err(Error::Io)?;
        Self::from_reader_with_options(&mut file, options)
    }

    /// Open an already parsed in-memory CFB file, preserving the password
    /// handoff when the `FILEPASS` record needs a password.
    pub fn from_ole_file_or_password_required(
        mut ole_file: OleFile<Cursor<Vec<u8>>>,
        options: OpenOptions<'_>,
    ) -> Result<std::result::Result<Self, OleFile<Cursor<Vec<u8>>>>> {
        crate::raw::preflight_ole_file(&ole_file, options.source_limits_value())?;
        let signed = has_signature(&ole_file);
        let metadata = eager_metadata(&mut ole_file);
        let reader = match RawReader::from_ole_file_or_password_required(ole_file, options) {
            Ok(Ok(reader)) => reader,
            Ok(Err(ole_file)) => return Ok(Err(ole_file)),
            Err(error) => return Err(Error::Reader(error)),
        };
        Ok(Ok(Self::from_reader_parts(
            reader, options, metadata, signed,
        )?))
    }

    /// Open an already parsed in-memory CFB file.
    pub fn from_ole_file(
        mut ole_file: OleFile<Cursor<Vec<u8>>>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        crate::raw::preflight_ole_file(&ole_file, options.source_limits_value())?;
        let signed = has_signature(&ole_file);
        let metadata = eager_metadata(&mut ole_file);
        let reader = RawReader::from_ole_file_with_options(ole_file, options)?;
        Self::from_reader_parts(reader, options, metadata, signed)
    }

    /// Adopt a source-backed semantic reader and the exact CFB view it came
    /// from. A deferred cell publication is materialized at this explicit
    /// facade handoff so the package's retained CFB and semantic owner always
    /// describe the same bytes; the semantic owner also retains any original
    /// dynamic-source fence needed by package output and compatibility routes.
    /// Ordinary source-backed commits remain lazy.
    pub fn from_source(workbook: SourceBackedWorkbook) -> Result<Self> {
        let workbook = workbook.materialized()?;
        let cfb = workbook.shared_cfb();
        let signed = has_shared_signature(&cfb);
        let source_limits = workbook.retained_limits();
        let leniency = workbook.retained_leniency();
        Ok(Self {
            inner: Arc::new(PackageInner {
                state: PackageState::Source {
                    cfb,
                    workbook,
                    encryption: None,
                    materialized: OnceLock::new(),
                },
                policy: RetainedPolicy::Source {
                    limits: source_limits,
                    leniency,
                },
                signed,
                metadata: OnceLock::new(),
                catalog: OnceLock::new(),
                edit_snapshot: None,
                workbook_stream: OnceLock::new(),
            }),
        })
    }

    pub(crate) fn from_owned_bytes(bytes: Vec<u8>, options: OpenOptions<'_>) -> Result<Self> {
        let cfb_limits =
            litchi_cfb::OleFileLimits::new(options.source_limits_value().max_input_bytes() as u64)?;
        let mut ole_file = OleFile::open(Cursor::new(bytes), cfb_limits)?;
        crate::raw::preflight_ole_file(&ole_file, options.source_limits_value())?;
        let signed = has_signature(&ole_file);
        let metadata = eager_metadata(&mut ole_file);
        let reader = RawReader::from_ole_file_with_options(ole_file, options)?;
        Self::from_reader_parts(reader, options, metadata, signed)
    }

    fn from_reader_parts(
        reader: RawReader,
        options: OpenOptions<'_>,
        metadata: Metadata,
        signed: bool,
    ) -> Result<Self> {
        let encryption = reader.encryption();
        let metadata_cache = OnceLock::new();
        let _ = metadata_cache.set(metadata);
        Ok(Self {
            inner: Arc::new(PackageInner {
                state: PackageState::Eager {
                    reader: Arc::new(reader),
                    encryption,
                },
                policy: RetainedPolicy::Eager {
                    source_limits: options.source_limits_value(),
                    leniency: options.leniency_value(),
                    compatibility_profile: options.compatibility_profile_value(),
                },
                signed,
                metadata: metadata_cache,
                catalog: OnceLock::new(),
                edit_snapshot: None,
                workbook_stream: OnceLock::new(),
            }),
        })
    }

    /// Physical source limits on an eager owner. Positional owners retain
    /// their own scan and materialization profile instead.
    #[must_use]
    pub fn source_limits(&self) -> Option<SourceLimits> {
        match self.inner.policy {
            RetainedPolicy::Eager { source_limits, .. } => Some(source_limits),
            RetainedPolicy::Source { .. } => None,
        }
    }

    /// Raw variable-payload codecs use the package's retained physical
    /// profile, including when a source-backed package is materialized only
    /// for a compatibility projection such as worksheet features.
    pub(crate) fn raw_record_limits(&self) -> Result<SourceLimits> {
        match self.inner.policy {
            RetainedPolicy::Eager { source_limits, .. } => Ok(source_limits),
            RetainedPolicy::Source { limits, .. } => {
                SourceLimits::from_physical_profile(limits.max_input_bytes).map_err(Error::Reader)
            },
        }
    }

    /// Positional read policy, when this package adopts a lazy reader.
    #[must_use]
    pub fn source_backed_limits(&self) -> Option<SourceBackedLimits> {
        match self.inner.policy {
            RetainedPolicy::Source { limits, .. } => Some(limits),
            RetainedPolicy::Eager { .. } => None,
        }
    }

    /// Retained policy for tolerated reader defects.
    #[must_use]
    pub fn leniency(&self) -> Leniency {
        match self.inner.policy {
            RetainedPolicy::Eager { leniency, .. } | RetainedPolicy::Source { leniency, .. } => {
                leniency
            },
        }
    }

    /// Explicit eager-reader compatibility policy, if this is an eager owner.
    #[must_use]
    pub fn compatibility_profile(&self) -> Option<CompatibilityProfile> {
        match self.inner.policy {
            RetainedPolicy::Eager {
                compatibility_profile,
                ..
            } => Some(compatibility_profile),
            RetainedPolicy::Source { .. } => None,
        }
    }

    /// Return the immutable read projection.
    #[must_use]
    pub fn workbook(&self) -> Workbook {
        Workbook {
            package: Arc::clone(&self.inner),
        }
    }

    /// Share the complete workbook-tab catalog; see
    /// [`Workbook::sheet_catalog`].
    ///
    /// # Errors
    ///
    /// Returns the positional owner's source or metadata read error while
    /// reading the catalog. A source-backed package checks its source on
    /// every direct access, including accesses after the catalog is shared.
    pub fn sheet_catalog(&self) -> Result<Arc<[crate::SheetMetadata]>> {
        sheet_catalog(&self.inner)
    }

    /// Return package metadata.
    ///
    /// Eager metadata is captured before BIFF parsing, matching the existing
    /// reader's property-set repair ordering. Source-backed metadata is read
    /// on first use and retained only after both source-version fences pass;
    /// absent or malformed property sets are an empty metadata projection.
    pub fn metadata(&self) -> Result<Metadata> {
        match &self.inner.state {
            PackageState::Eager { .. } => {
                Ok(self.inner.metadata.get().cloned().unwrap_or_default())
            },
            PackageState::Source { cfb, .. } => {
                ensure_source_fence(&self.inner)?;
                let result = (|| -> Result<Metadata> {
                    let expected = cfb
                        .source_version()
                        .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                    let metadata = if let Some(metadata) = self.inner.metadata.get() {
                        metadata.clone()
                    } else {
                        match cfb.get_metadata() {
                            Ok(metadata) => Metadata::from(metadata),
                            Err(error) if metadata_error_is_soft(&error) => Metadata::default(),
                            Err(error) => {
                                return Err(Error::Source(SourceBackedError::from(error)));
                            },
                        }
                    };
                    let observed = cfb
                        .source_version()
                        .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                    if observed != expected {
                        return Err(Error::Source(SourceBackedError::SourceChanged {
                            expected,
                            observed,
                        }));
                    }
                    Ok(self.inner.metadata.get_or_init(|| metadata).clone())
                })();
                finish_source_fence(&self.inner, result)
            },
        }
    }

    /// Whether this package uses the positional source-backed owner.
    #[must_use]
    pub fn is_source_backed(&self) -> bool {
        matches!(self.inner.state, PackageState::Source { .. })
    }

    /// The retained shared CFB source, if this is a source-backed package.
    #[must_use]
    pub fn source(&self) -> Option<&SharedOleFile> {
        match &self.inner.state {
            PackageState::Source { cfb, .. } => Some(cfb.as_ref()),
            PackageState::Eager { .. } => None,
        }
    }

    /// Clone the retained shared CFB source, if this is source-backed.
    #[must_use]
    pub fn shared_source(&self) -> Option<Arc<SharedOleFile>> {
        match &self.inner.state {
            PackageState::Source { cfb, .. } => Some(Arc::clone(cfb)),
            PackageState::Eager { .. } => None,
        }
    }

    /// The source-backed semantic owner, if present.
    #[must_use]
    pub fn source_workbook(&self) -> Option<&SourceBackedWorkbook> {
        match &self.inner.state {
            PackageState::Source { workbook, .. } => Some(workbook),
            PackageState::Eager { .. } => None,
        }
    }

    /// Check the retained dynamic source without materializing its bytes.
    ///
    /// Scoped edit families use this version-only fence when a policy error
    /// must not hide a concurrent source change.
    pub(crate) fn ensure_source_fence(&self) -> Result<()> {
        ensure_source_fence(&self.inner)
    }

    /// Borrow the retained eager reader, materializing a source-backed reader
    /// only on this explicit compatibility escape.
    pub fn reader_ref(&self) -> Result<&crate::raw::Reader<Cursor<Vec<u8>>>> {
        reader_ref(&self.inner)
    }

    /// Admit edited bytes under this package's retained read policy.
    pub(crate) fn republish_bytes(&self, bytes: Vec<u8>) -> Result<Self> {
        ensure_source_fence(&self.inner)?;
        let result = (|| -> Result<Self> {
            match self.inner.policy {
                RetainedPolicy::Eager {
                    source_limits,
                    leniency,
                    compatibility_profile,
                } => Self::from_owned_bytes(
                    bytes,
                    OpenOptions::default()
                        .source_limits(source_limits)
                        .leniency(leniency)
                        .compatibility_profile(compatibility_profile),
                ),
                RetainedPolicy::Source { limits, leniency } => {
                    let cfb_limits = self
                        .source()
                        .map(SharedOleFile::limits)
                        .ok_or(Error::InvalidData("source-backed XLS package lost its CFB"))?;
                    let cfb = Arc::new(SharedOleFile::open(
                        Arc::new(litchi_core::OwnedSource::new(bytes)),
                        cfb_limits,
                    )?);
                    let workbook = SourceBackedWorkbook::from_shared_ole_file_with_options(
                        cfb,
                        crate::SourceBackedOptions::default()
                            .limits(limits)
                            .leniency(leniency),
                    )?;
                    Self::from_source(workbook)
                },
            }
        })();
        finish_source_fence(&self.inner, result)
    }

    /// The Workbook stream (`Workbook`, or `Book` from older producers), read
    /// once per package and shared.
    pub(crate) fn workbook_stream(&self) -> Result<Arc<[u8]>> {
        if let Some(stream) = self.inner.workbook_stream.get() {
            ensure_source_fence(&self.inner)?;
            return Ok(Arc::clone(stream));
        }
        ensure_source_fence(&self.inner)?;
        let result = (|| -> Result<Vec<u8>> {
            match &self.inner.state {
                PackageState::Eager { reader, .. } => {
                    let data = reader.get_ref().get_ref();
                    let mut ole = OleFile::open(
                        Cursor::new(data.as_slice()),
                        litchi_cfb::OleFileLimits::default(),
                    )?;
                    let name = if ole.exists(&["Workbook"]) {
                        "Workbook"
                    } else {
                        "Book"
                    };
                    Ok(ole.open_stream(&[name])?)
                },
                PackageState::Source { cfb, .. } => {
                    let expected = cfb
                        .source_version()
                        .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                    let name = if cfb.exists(&["Workbook"]) {
                        "Workbook"
                    } else {
                        "Book"
                    };
                    let bytes = cfb.open_stream(&[name])?;
                    let observed = cfb
                        .source_version()
                        .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                    if observed != expected {
                        return Err(Error::Source(SourceBackedError::SourceChanged {
                            expected,
                            observed,
                        }));
                    }
                    Ok(bytes)
                },
            }
        })();
        let bytes = finish_source_fence(&self.inner, result)?;
        let stream: Arc<[u8]> = Arc::from(bytes);
        Ok(Arc::clone(
            self.inner.workbook_stream.get_or_init(|| stream),
        ))
    }

    /// Return the exact source bytes, including password-to-open encryption.
    ///
    /// This method is deliberately named as a source escape.  Ordinary
    /// `to_bytes`, `write_to`, and `save` refuse encrypted input.
    pub fn source_bytes(&self) -> Result<Vec<u8>> {
        exact_bytes(&self.inner)
    }

    /// Return bytes suitable for ordinary publication.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        ensure_publication_allowed(&self.inner)?;
        exact_bytes(&self.inner)
    }

    /// Return the original encrypted artifact exactly, without attempting to
    /// decrypt or reserialize it.
    pub fn to_encrypted_bytes(&self) -> Result<Vec<u8>> {
        exact_bytes(&self.inner)
    }

    /// Explicit plain-output route.  A password-to-open source is refused
    /// until a supported decrypt/re-encrypt owner is available.
    pub fn to_plain_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes()
    }

    /// Whether the retained input uses password-to-open encryption.
    /// Ordinary publication refuses these inputs to avoid emitting plaintext.
    #[must_use]
    pub fn is_encrypted(&self) -> bool {
        self.encryption().is_some()
    }

    /// Whether the retained input carries a compound-file signature stream.
    #[must_use]
    pub(crate) fn is_signed(&self) -> bool {
        self.inner.signed
    }

    /// Return the password-to-open profile retained from source admission.
    ///
    /// A CryptoAPI RC4 profile includes its exact key length. Source-backed
    /// XLS admission is currently clear-only and therefore returns `None`.
    #[must_use]
    pub fn encryption(&self) -> Option<crate::EncryptionProfile> {
        match &self.inner.state {
            PackageState::Eager { encryption, .. } | PackageState::Source { encryption, .. } => {
                *encryption
            },
        }
    }

    /// Stream the exact package bytes to a sequential sink.
    pub fn write_to<W: Write>(&self, writer: &mut W) -> Result<()> {
        ensure_publication_allowed(&self.inner)?;
        stream_exact(&self.inner, writer)
    }

    /// Atomically replace a destination with the exact package bytes.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.save_with_durability(path, Durability::Full)
    }

    /// Atomically replace a destination with explicit durability.
    pub fn save_with_durability<P: AsRef<Path>>(
        &self,
        path: P,
        durability: Durability,
    ) -> Result<()> {
        ensure_publication_allowed(&self.inner)?;
        match &self.inner.state {
            PackageState::Eager { reader, .. } => {
                let bytes = reader.get_ref().get_ref();
                litchi_cfb::atomic::save(
                    path,
                    durability,
                    litchi_cfb::atomic::Overwrite::Replace,
                    |writer| {
                        writer.write_all(bytes).map_err(Error::Io)?;
                        Ok(())
                    },
                    || Ok(()),
                )
            },
            PackageState::Source { cfb, .. } => {
                ensure_source_fence(&self.inner)?;
                let source = cfb.source_arc();
                let expected = cfb.captured_source_version();
                let result = litchi_cfb::atomic::save(
                    path,
                    durability,
                    litchi_cfb::atomic::Overwrite::Replace,
                    |writer| {
                        let before = source.version()?;
                        if before != expected {
                            return Err(Error::Source(SourceBackedError::SourceChanged {
                                expected,
                                observed: before,
                            }));
                        }
                        stream_source(cfb, &*source, cfb.file_size(), writer)
                    },
                    || {
                        let after = source.version()?;
                        if after != expected {
                            return Err(Error::Source(SourceBackedError::SourceChanged {
                                expected,
                                observed: after,
                            }));
                        }
                        Ok(())
                    },
                );
                finish_source_fence(&self.inner, result)
            },
        }
    }

    /// Begin the editable cell-values adapter on demand.
    ///
    /// No snapshot is constructed while a package is opened. Primary edits
    /// retain their validated snapshot for subsequent operations; this explicit
    /// low-level adapter also retains and intersects that selected profile.
    pub(crate) fn edit_snapshot_with_options(
        &self,
        options: EditOptions,
    ) -> Result<crate::cell_values::Snapshot> {
        let (limits, cfb_limits) =
            self.edit_profile(options.object_limits_value(), options.cfb_limits_value())?;
        ensure_source_fence(&self.inner)?;
        let result = (|| -> Result<crate::cell_values::Snapshot> {
            if let Some(cfb) = self.source() {
                cfb.source_version().map_err(SourceBackedError::from)?;
            }
            if let Some(snapshot) = &self.inner.edit_snapshot {
                let result = snapshot.constrain_limits(limits, cfb_limits);
                if let Some(cfb) = self.source() {
                    cfb.source_version().map_err(SourceBackedError::from)?;
                }
                return result.map_err(Error::Reader);
            }
            match &self.inner.state {
                PackageState::Eager { reader, .. } => {
                    if self.inner.signed {
                        let probe = litchi_ole_common::object::Probe::open_with_options(
                            reader.get_ref().get_ref().clone(),
                            litchi_ole_common::object::OpenOptions::new()
                                .limits(limits)
                                .cfb_limits(cfb_limits),
                        )
                        .map_err(Error::Object)?;
                        crate::cell_values::Snapshot::from_probe(probe).map_err(Error::Reader)
                    } else {
                        crate::cell_values::Snapshot::from_slice_with_limits_and_cfb(
                            reader.get_ref().get_ref(),
                            limits,
                            cfb_limits,
                        )
                        .map_err(Error::Reader)
                    }
                },
                PackageState::Source { cfb, .. } => {
                    if self.inner.signed {
                        crate::cell_values::Snapshot::from_shared_probe(cfb, limits, cfb_limits)
                            .map_err(Error::Reader)
                    } else {
                        crate::cell_values::Snapshot::from_shared_ole_file(cfb, limits, cfb_limits)
                            .map_err(Error::Reader)
                    }
                },
            }
        })();
        finish_source_fence(&self.inner, result)
    }

    pub(crate) fn edit_profile(
        &self,
        limits: litchi_ole_common::object::Limits,
        cfb_limits: litchi_cfb::OleFileLimits,
    ) -> Result<(litchi_ole_common::object::Limits, litchi_cfb::OleFileLimits)> {
        let limits = intersect_object_limits(limits, self.inner.policy);
        let retained_input = retained_input_limit(&self.inner);
        let source_cfb_limits = match &self.inner.state {
            PackageState::Source { cfb, .. } => Some(cfb.limits()),
            PackageState::Eager { .. } => None,
        };
        let cfb_limits = if source_cfb_limits.is_none() {
            let admitted = self.inner.edit_snapshot.as_ref().map_or_else(
                litchi_cfb::OleFileLimits::default,
                crate::cell_values::Snapshot::cfb_limits,
            );
            cfb_limits
                .with_max_directory_bytes(
                    cfb_limits
                        .max_directory_bytes()
                        .min(admitted.max_directory_bytes()),
                )?
                .with_max_allocation_table_bytes(
                    cfb_limits
                        .max_allocation_table_bytes()
                        .min(admitted.max_allocation_table_bytes()),
                )?
        } else {
            cfb_limits
        };
        let cfb_limits = intersect_cfb_limits(cfb_limits, retained_input, source_cfb_limits)?;
        if let Some(snapshot) = &self.inner.edit_snapshot {
            return snapshot
                .effective_limits(limits, cfb_limits)
                .map_err(Error::Reader);
        }
        Ok((limits, cfb_limits))
    }

    /// The cell-values read view of this package (its worksheets, cells,
    /// styles and Workbook stream) under its retained edit profile. Edits go
    /// through [`Package::edit`].
    ///
    /// # Errors
    ///
    /// Returns an error when the workbook cannot be read as cell values.
    pub fn cell_values(&self) -> Result<crate::cell_values::Snapshot> {
        self.edit_snapshot_with_options(EditOptions::default())
    }
}

/// An owned read-only workbook projection.
pub struct Workbook {
    package: Arc<PackageInner>,
}

impl Clone for Workbook {
    fn clone(&self) -> Self {
        Self {
            package: Arc::clone(&self.package),
        }
    }
}

impl fmt::Debug for Workbook {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Workbook")
            .field(
                "source_backed",
                &matches!(self.package.state, PackageState::Source { .. }),
            )
            .finish_non_exhaustive()
    }
}

impl Workbook {
    /// Borrow the eager reader, materializing source-backed input only here.
    pub fn reader_ref(&self) -> Result<&crate::raw::Reader<Cursor<Vec<u8>>>> {
        reader_ref(&self.package)
    }

    /// Borrow the source-backed projection when this package has one.
    #[must_use]
    pub fn source_backed_ref(&self) -> Option<&SourceBackedWorkbook> {
        match &self.package.state {
            PackageState::Source { workbook, .. } => Some(workbook),
            PackageState::Eager { .. } => None,
        }
    }

    /// Return the number of workbook tabs in original tab order.
    ///
    /// This includes chart and macro tabs as well as ordinary worksheets.
    pub fn sheet_count(&self) -> Result<usize> {
        match &self.package.state {
            PackageState::Source { workbook, .. } => workbook.sheet_count().map_err(Error::Source),
            PackageState::Eager { reader, .. } => Ok(reader.sheets().len()),
        }
    }

    /// Select one workbook tab by its semantic name or checked zero-based
    /// position and return its owned catalog metadata.
    ///
    /// The position is in the complete workbook-tab catalog. In particular,
    /// it is not the ordinal accepted by `xls_worksheet` when chart or macro
    /// tabs precede an ordinary worksheet.
    pub fn sheet<'a>(
        &self,
        selector: impl Into<litchi_core::selector::Selector<'a, std::convert::Infallible>>,
    ) -> Result<Option<crate::SheetMetadata>> {
        let selector = selector.into();
        match &self.package.state {
            PackageState::Source { workbook, .. } => {
                workbook.sheet(selector).map_err(Error::Source)
            },
            PackageState::Eager { reader, .. } => {
                let sheet = match selector {
                    litchi_core::selector::Selector::Name(name) => {
                        reader.sheet_by_name(name.as_ref())
                    },
                    litchi_core::selector::Selector::Position(position) => {
                        reader.sheet(position.get())
                    },
                    litchi_core::selector::Selector::Id(id) => match id {},
                    _ => return Err(Error::InvalidData("unsupported sheet selector")),
                };
                Ok(sheet.cloned())
            },
        }
    }

    /// Share the complete workbook-tab catalog in original tab order.
    ///
    /// The package captures the catalog once, under its source fence when
    /// needed, and shares the immutable result without copying its names.
    /// Callers that must observe a changed source retain their own freshness
    /// fence around this projection.
    ///
    /// ```
    /// let package = litchi_xls::Package::new()?;
    /// let workbook = package.workbook();
    /// assert_eq!(workbook.sheet_catalog()?.len(), 1);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the positional owner's source or metadata read error while
    /// reading the catalog. A source-backed package checks its source on
    /// every direct access, including accesses after the catalog is shared.
    pub fn sheet_catalog(&self) -> Result<Arc<[crate::SheetMetadata]>> {
        sheet_catalog(&self.package)
    }

    /// Return complete workbook-tab metadata in original tab order.
    ///
    /// This includes chart and macro tabs. Each descriptor distinguishes the
    /// tab position from its optional ordinary worksheet ordinal.
    pub fn sheets(&self) -> Result<Vec<crate::SheetMetadata>> {
        match &self.package.state {
            PackageState::Source { workbook, .. } => {
                workbook.sheet_metadata().map_err(Error::Source)
            },
            PackageState::Eager { reader, .. } => Ok(reader.sheets().to_vec()),
        }
    }

    /// Return ordinary worksheet names.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        match &self.package.state {
            PackageState::Source { workbook, .. } => {
                workbook.worksheet_names().map_err(Error::Source)
            },
            PackageState::Eager { reader, .. } => {
                Ok(litchi_core::sheet::WorkbookTrait::worksheet_names(reader.as_ref()).to_vec())
            },
        }
    }

    /// Return the ordinary worksheet count.
    pub fn worksheet_count(&self) -> Result<usize> {
        match &self.package.state {
            PackageState::Source { workbook, .. } => {
                workbook.worksheet_count().map_err(Error::Source)
            },
            PackageState::Eager { reader, .. } => Ok(
                litchi_core::sheet::WorkbookTrait::worksheet_count(reader.as_ref()),
            ),
        }
    }

    /// Return the source-backed tolerance report or the eager reader report.
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        match &self.package.state {
            PackageState::Source { cfb, workbook, .. } => {
                ensure_source_fence(&self.package)?;
                let result = (|| -> Result<ToleranceReport> {
                    let expected = cfb
                        .source_version()
                        .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                    let mut report = workbook.tolerance_report();
                    report.append_property_sets(|path| cfb.property_set_stream(path));
                    let observed = cfb
                        .source_version()
                        .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                    if observed != expected {
                        return Err(Error::Source(SourceBackedError::SourceChanged {
                            expected,
                            observed,
                        }));
                    }
                    Ok(report)
                })();
                finish_source_fence(&self.package, result)
            },
            PackageState::Eager { reader, .. } => Ok(reader.tolerance_report().clone()),
        }
    }

    /// Extract ordinary worksheet text using the owner route.
    pub fn text(&self) -> Result<String> {
        match &self.package.state {
            PackageState::Source { workbook, .. } => workbook.text().map_err(Error::Source),
            PackageState::Eager { reader, .. } => eager_text(reader.as_ref()),
        }
    }

    /// Resolve one worksheet by a canonical name or zero-based position.
    pub fn worksheet<'a>(
        &self,
        selector: impl Into<litchi_core::selector::Selector<'a, std::convert::Infallible>>,
    ) -> Result<Option<Worksheet>> {
        let selector = selector.into();
        match &self.package.state {
            PackageState::Source { workbook, .. } => match selector {
                litchi_core::selector::Selector::Name(name) => workbook
                    .worksheet_by_name(name.as_ref())
                    .map(|worksheet| worksheet.map(Worksheet::source))
                    .map_err(Error::Source),
                litchi_core::selector::Selector::Position(position) => workbook
                    .worksheet_by_index(position.get())
                    .map(|worksheet| worksheet.map(Worksheet::source))
                    .map_err(Error::Source),
                litchi_core::selector::Selector::Id(id) => match id {},
                _ => return Err(Error::InvalidData("unsupported worksheet selector")),
            },
            PackageState::Eager { reader, .. } => {
                let index = match selector {
                    litchi_core::selector::Selector::Name(name) => {
                        let names =
                            litchi_core::sheet::WorkbookTrait::worksheet_names(reader.as_ref());
                        names
                            .iter()
                            .position(|candidate| candidate.eq_ignore_ascii_case(name.as_ref()))
                    },
                    litchi_core::selector::Selector::Position(position) => Some(position.get()),
                    litchi_core::selector::Selector::Id(id) => match id {},
                    _ => return Err(Error::InvalidData("unsupported worksheet selector")),
                };
                Ok(index
                    .filter(|index| {
                        *index < litchi_core::sheet::WorkbookTrait::worksheet_count(reader.as_ref())
                    })
                    .map(|index| Worksheet::eager(Arc::clone(reader), index)))
            },
        }
    }

    /// Enumerate ordinary worksheets as owned handles.
    pub fn worksheets(&self) -> Result<Vec<Worksheet>> {
        let count = self.worksheet_count()?;
        let mut worksheets = Vec::new();
        worksheets
            .try_reserve_exact(count)
            .map_err(|source| Error::Allocation {
                resource: "XLS workbook worksheet projection",
                source,
            })?;
        for index in 0..count {
            if let Some(worksheet) = self.worksheet(index)? {
                worksheets.push(worksheet);
            }
        }
        Ok(worksheets)
    }
}

/// A read-only owned worksheet handle from either lazy or eager XLS storage.
pub struct Worksheet {
    inner: WorksheetInner,
}

enum WorksheetInner {
    Source(crate::SourceBackedWorksheet),
    Eager {
        reader: Arc<RawReader>,
        index: usize,
    },
}

impl fmt::Debug for Worksheet {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Worksheet").finish_non_exhaustive()
    }
}

impl Clone for Worksheet {
    fn clone(&self) -> Self {
        match &self.inner {
            WorksheetInner::Source(worksheet) => Self::source(worksheet.clone()),
            WorksheetInner::Eager { reader, index } => Self::eager(Arc::clone(reader), *index),
        }
    }
}

impl Worksheet {
    fn source(worksheet: crate::SourceBackedWorksheet) -> Self {
        Self {
            inner: WorksheetInner::Source(worksheet),
        }
    }

    fn eager(reader: Arc<RawReader>, index: usize) -> Self {
        Self {
            inner: WorksheetInner::Eager { reader, index },
        }
    }

    /// Zero-based worksheet index.
    pub fn index(&self) -> Result<usize> {
        match &self.inner {
            WorksheetInner::Source(worksheet) => worksheet.index().map_err(Error::Source),
            WorksheetInner::Eager { index, .. } => Ok(*index),
        }
    }

    /// Worksheet name.
    pub fn name(&self) -> Result<String> {
        match &self.inner {
            WorksheetInner::Source(worksheet) => worksheet.name().map_err(Error::Source),
            WorksheetInner::Eager { reader, index } => {
                let worksheet = reader.xls_worksheet(*index)?;
                Ok(worksheet.name().to_string())
            },
        }
    }

    /// Read one cell value at zero-based row and column coordinates.
    pub fn cell_value(
        &self,
        reference: crate::cell_values::Reference,
    ) -> Result<Option<CellValue>> {
        let row = u32::from(reference.row());
        let column = u32::from(reference.column());
        match &self.inner {
            WorksheetInner::Source(worksheet) => {
                worksheet.cell_value(row, column).map_err(Error::Source)
            },
            WorksheetInner::Eager { reader, index } => {
                let worksheet = reader.xls_worksheet(*index)?;
                Ok(worksheet
                    .get_cell(row, column)
                    .map(|cell| cell.value().clone()))
            },
        }
    }

    /// Return the underlying source-backed worksheet for advanced lazy reads.
    #[must_use]
    pub fn source_backed_ref(&self) -> Option<&crate::SourceBackedWorksheet> {
        match &self.inner {
            WorksheetInner::Source(worksheet) => Some(worksheet),
            WorksheetInner::Eager { .. } => None,
        }
    }
}

fn source_fence_workbook(inner: &PackageInner) -> Option<&SourceBackedWorkbook> {
    // A package created from a deferred workbook owns the composed CFB and
    // semantic target, but the retained workbook is the only owner that also
    // knows the original dynamic source. Route package-level bytes and eager
    // compatibility escapes through that owner when its fence is present.
    match &inner.state {
        PackageState::Source { workbook, .. } if workbook.needs_source_fence() => Some(workbook),
        PackageState::Eager { .. } | PackageState::Source { .. } => None,
    }
}

fn ensure_source_fence(inner: &PackageInner) -> Result<()> {
    if let Some(workbook) = source_fence_workbook(inner) {
        return workbook.ensure_current().map_err(Error::Source);
    }
    if let PackageState::Source { cfb, .. } = &inner.state {
        let expected = cfb.captured_source_version();
        let observed = cfb.source_version().map_err(SourceBackedError::from)?;
        if observed != expected {
            return Err(Error::Source(SourceBackedError::SourceChanged {
                expected,
                observed,
            }));
        }
    }
    Ok(())
}

fn finish_source_fence<T>(inner: &PackageInner, result: Result<T>) -> Result<T> {
    ensure_source_fence(inner)?;
    result
}

fn reader_ref(inner: &Arc<PackageInner>) -> Result<&RawReader> {
    match &inner.state {
        PackageState::Eager { reader, .. } => Ok(reader.as_ref()),
        PackageState::Source {
            workbook,
            materialized,
            cfb,
            ..
        } => {
            ensure_source_fence(inner)?;
            let result = (|| -> Result<&RawReader> {
                cfb.source_version().map_err(SourceBackedError::from)?;
                let result = materialized
                    .get_or_init(|| workbook.materialize_eager().map(Arc::new).map_err(Arc::new));
                cfb.source_version().map_err(SourceBackedError::from)?;
                match result {
                    Ok(reader) => Ok(reader.as_ref()),
                    Err(error) => Err(Error::CachedSource(Arc::clone(error))),
                }
            })();
            finish_source_fence(inner, result)
        },
    }
}

fn exact_bytes(inner: &Arc<PackageInner>) -> Result<Vec<u8>> {
    ensure_source_fence(inner)?;
    let result = (|| -> Result<Vec<u8>> {
        match &inner.state {
            PackageState::Eager { reader, .. } => Ok(reader.get_ref().get_ref().clone()),
            PackageState::Source { cfb, .. } => {
                let length = usize::try_from(cfb.file_size()).map_err(|_| {
                    Error::InvalidData("XLS source size does not fit the current address space")
                })?;
                let mut bytes = Vec::new();
                bytes
                    .try_reserve_exact(length)
                    .map_err(|source| Error::Allocation {
                        resource: "XLS package output",
                        source,
                    })?;
                bytes.resize(length, 0);
                let source = cfb.source_arc();
                let expected = cfb
                    .source_version()
                    .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                if let Err(error) = source.read_exact_at(0, &mut bytes) {
                    return Err(source_read_error(cfb, expected, error));
                }
                let observed = cfb
                    .source_version()
                    .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                if observed != expected {
                    return Err(Error::Source(SourceBackedError::SourceChanged {
                        expected,
                        observed,
                    }));
                }
                Ok(bytes)
            },
        }
    })();
    finish_source_fence(inner, result)
}

fn stream_exact<W: Write>(inner: &Arc<PackageInner>, writer: &mut W) -> Result<()> {
    ensure_source_fence(inner)?;
    let result = (|| -> Result<()> {
        match &inner.state {
            PackageState::Eager { reader, .. } => writer
                .write_all(reader.get_ref().get_ref())
                .map_err(Error::Io),
            PackageState::Source { cfb, .. } => {
                let source = cfb.source_arc();
                let expected = cfb
                    .source_version()
                    .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                stream_source(cfb, &*source, cfb.file_size(), writer)?;
                let observed = cfb
                    .source_version()
                    .map_err(|error| Error::Source(SourceBackedError::from(error)))?;
                if observed != expected {
                    return Err(Error::Source(SourceBackedError::SourceChanged {
                        expected,
                        observed,
                    }));
                }
                Ok(())
            },
        }
    })();
    finish_source_fence(inner, result)
}

fn stream_source<W: Write + ?Sized>(
    cfb: &SharedOleFile,
    source: &dyn ReadAt,
    length: u64,
    writer: &mut W,
) -> Result<()> {
    let mut offset = 0_u64;
    let mut buffer = [0_u8; 64 * 1024];
    while offset < length {
        let count = usize::try_from((length - offset).min(buffer.len() as u64))
            .map_err(|_| Error::InvalidData("XLS source range does not fit usize"))?;
        if let Err(error) = source.read_exact_at(offset, &mut buffer[..count]) {
            let expected = cfb.captured_source_version();
            return Err(source_read_error(cfb, expected, error));
        }
        writer.write_all(&buffer[..count]).map_err(Error::Io)?;
        offset = offset
            .checked_add(count as u64)
            .ok_or(Error::InvalidData("XLS source offset overflow"))?;
    }
    Ok(())
}

fn source_read_error(
    cfb: &SharedOleFile,
    expected: litchi_core::SourceVersion,
    error: io::Error,
) -> Error {
    match cfb.source_version() {
        Ok(observed) if observed != expected => {
            Error::Source(SourceBackedError::SourceChanged { expected, observed })
        },
        Ok(_) => Error::Io(error),
        Err(version_error) => Error::Source(SourceBackedError::from(version_error)),
    }
}

fn ensure_publication_allowed(inner: &Arc<PackageInner>) -> Result<()> {
    let encrypted = match &inner.state {
        PackageState::Eager { encryption, .. } | PackageState::Source { encryption, .. } => {
            encryption.is_some()
        },
    };
    if encrypted {
        Err(Error::EncryptionPolicy)
    } else {
        Ok(())
    }
}

fn retained_input_limit(inner: &Arc<PackageInner>) -> u64 {
    match inner.policy {
        RetainedPolicy::Eager { source_limits, .. } => {
            u64::try_from(source_limits.max_input_bytes()).unwrap_or(u64::MAX)
        },
        RetainedPolicy::Source { limits, .. } => limits.max_input_bytes,
    }
}

fn intersect_object_limits(
    mut limits: litchi_ole_common::object::Limits,
    policy: RetainedPolicy,
) -> litchi_ole_common::object::Limits {
    if let RetainedPolicy::Eager { source_limits, .. } = policy {
        limits.max_stream_size = limits
            .max_stream_size
            .min(u64::try_from(source_limits.max_entry_bytes()).unwrap_or(u64::MAX));
        limits.max_total_size = limits
            .max_total_size
            .min(u64::try_from(source_limits.max_aggregate_bytes()).unwrap_or(u64::MAX));
    }
    limits
}

fn intersect_cfb_limits(
    requested: litchi_cfb::OleFileLimits,
    retained_input: u64,
    source: Option<SharedOleFileLimits>,
) -> Result<litchi_cfb::OleFileLimits> {
    let input = requested
        .max_input_bytes()
        .min(retained_input)
        .min(source.map_or(u64::MAX, SharedOleFileLimits::max_input_bytes));
    let directory = requested
        .max_directory_bytes()
        .min(source.map_or(u64::MAX, SharedOleFileLimits::max_directory_bytes));
    let allocation = requested
        .max_allocation_table_bytes()
        .min(source.map_or(u64::MAX, SharedOleFileLimits::max_allocation_table_bytes));
    litchi_cfb::OleFileLimits::new(input)
        .map_err(Error::Cfb)?
        .with_max_directory_bytes(directory)
        .map_err(Error::Cfb)?
        .with_max_allocation_table_bytes(allocation)
        .map_err(Error::Cfb)
}

fn eager_metadata(ole_file: &mut OleFile<Cursor<Vec<u8>>>) -> Metadata {
    ole_file
        .get_metadata()
        .map(Metadata::from)
        .unwrap_or_default()
}

fn metadata_error_is_soft(error: &litchi_ole_common::Error) -> bool {
    matches!(
        error,
        litchi_ole_common::Error::Cfb(
            OleError::InvalidFormat(_)
                | OleError::InvalidData(_)
                | OleError::NotOleFile
                | OleError::CorruptedFile(_)
                | OleError::StreamNotFound,
        )
    )
}

fn eager_text(reader: &RawReader) -> Result<String> {
    let mut output = String::new();
    for index in 0..litchi_core::sheet::WorkbookTrait::worksheet_count(reader) {
        let worksheet = reader.xls_worksheet(index)?;
        for row in 0..worksheet.row_count() {
            for column in 0..worksheet.column_count() {
                if column > 0 {
                    output.push('\t');
                }
                if let Some(cell) = worksheet.get_cell(
                    u32::try_from(row).unwrap_or(u32::MAX),
                    u32::try_from(column).unwrap_or(u32::MAX),
                ) {
                    append_cell_text(&mut output, cell.value());
                }
            }
            output.push('\n');
        }
    }
    Ok(output)
}

fn append_cell_text(output: &mut String, cell: &CellValue) {
    match cell {
        CellValue::Empty => {},
        CellValue::Bool(value) => output.push_str(if *value { "TRUE" } else { "FALSE" }),
        CellValue::Int(value) => output.push_str(&value.to_string()),
        CellValue::Float(value) => output.push_str(&value.to_string()),
        CellValue::String(value) | CellValue::Error(value) => output.push_str(value),
        CellValue::DateTime(value) => output.push_str(&value.to_string()),
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => match cached_value.as_deref() {
            Some(CellValue::String(value)) => output.push_str(value),
            Some(CellValue::Int(value)) => output.push_str(&value.to_string()),
            Some(CellValue::Float(value)) => output.push_str(&value.to_string()),
            Some(CellValue::Bool(value)) => output.push_str(if *value { "TRUE" } else { "FALSE" }),
            _ => {
                output.push('=');
                output.push_str(formula);
            },
        },
    }
}

fn read_bounded<R: Read>(reader: &mut R, limits: SourceLimits) -> Result<Vec<u8>> {
    let maximum = limits.max_input_bytes();
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(maximum.min(64 * 1024))
        .map_err(|source| Error::Allocation {
            resource: "XLS package input",
            source,
        })?;
    bytes.clear();
    let mut buffer = [0_u8; 64 * 1024];
    loop {
        let request = maximum
            .saturating_sub(bytes.len())
            .saturating_add(1)
            .min(buffer.len());
        let count = reader.read(&mut buffer[..request])?;
        if count == 0 {
            break;
        }
        let next = bytes
            .len()
            .checked_add(count)
            .ok_or(Error::InvalidData("XLS package input length overflow"))?;
        if next > maximum {
            return Err(Error::Reader(ReaderError::Limit {
                kind: LimitKind::InputBytes,
                observed: next as u64,
                maximum: maximum as u64,
            }));
        }
        bytes
            .try_reserve(count)
            .map_err(|source| Error::Allocation {
                resource: "XLS package input",
                source,
            })?;
        bytes.extend_from_slice(&buffer[..count]);
    }
    Ok(bytes)
}

fn sheet_catalog(inner: &PackageInner) -> Result<Arc<[crate::SheetMetadata]>> {
    match &inner.state {
        PackageState::Source { workbook, .. } => {
            ensure_source_fence(inner)?;
            if let Some(catalog) = inner.catalog.get() {
                return Ok(Arc::clone(catalog));
            }
            let captured = workbook.sheet_metadata().map_err(Error::Source)?;
            let captured: Arc<[crate::SheetMetadata]> = Arc::from(captured.into_boxed_slice());
            Ok(Arc::clone(inner.catalog.get_or_init(|| captured)))
        },
        PackageState::Eager { reader, .. } => {
            if let Some(catalog) = inner.catalog.get() {
                return Ok(Arc::clone(catalog));
            }
            let captured = reader.shared_sheets();
            Ok(Arc::clone(inner.catalog.get_or_init(|| captured)))
        },
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, Package};
    use crate::{Builder, SourceBackedError, SourceBackedWorkbook};
    use litchi_cfb::OleError;
    use litchi_core::{ReadAt, SourceVersion};
    use std::io;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

    struct AdmissionSource {
        bytes: Arc<[u8]>,
        observations: AtomicUsize,
        revision: AtomicU64,
        bump_after: AtomicUsize,
    }

    impl AdmissionSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: Arc::from(bytes),
                observations: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
                bump_after: AtomicUsize::new(0),
            }
        }

        fn bump_after_next_observation(&self) {
            let next = self.observations.load(Ordering::SeqCst).saturating_add(1);
            self.bump_after.store(next, Ordering::SeqCst);
        }

        fn package(self: &Arc<Self>) -> Package {
            let reader: Arc<dyn ReadAt> = self.clone();
            let workbook = SourceBackedWorkbook::from_read_at(reader).unwrap();
            Package::from_source(workbook).unwrap()
        }
    }

    impl ReadAt for AdmissionSource {
        fn len(&self) -> io::Result<u64> {
            u64::try_from(self.bytes.len())
                .map_err(|_| io::Error::other("source length does not fit u64"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let start = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
            if start >= self.bytes.len() || output.is_empty() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            let ordinal = self.observations.fetch_add(1, Ordering::SeqCst) + 1;
            let observed =
                SourceVersion::new(0x584c_535f_4144_4d49, self.revision.load(Ordering::SeqCst));
            if self.bump_after.load(Ordering::SeqCst) == ordinal {
                self.bump_after.store(0, Ordering::SeqCst);
                self.revision.fetch_add(1, Ordering::SeqCst);
            }
            Ok(observed)
        }
    }

    fn source_bytes() -> Vec<u8> {
        let mut builder = Builder::new();
        let sheet = builder.add_worksheet("Sheet1").unwrap();
        builder.write_number(sheet, 0, 0, 1.0).unwrap();
        builder.build().unwrap().to_bytes().unwrap()
    }

    #[test]
    fn republish_keeps_candidate_admission_error_when_source_is_stable() {
        let source = Arc::new(AdmissionSource::new(source_bytes()));
        let package = source.package();

        let error = package.republish_bytes(vec![0, 1, 2]).unwrap_err();
        assert!(matches!(
            error,
            Error::Cfb(OleError::NotOleFile | OleError::InvalidFormat(_))
        ));
    }

    #[test]
    fn republish_promotes_source_change_over_candidate_admission_error() {
        let source = Arc::new(AdmissionSource::new(source_bytes()));
        let package = source.package();
        // The hook mutates the original immediately after the leading source
        // observation returns, while invalid candidate admission is pending.
        source.bump_after_next_observation();

        let error = package.republish_bytes(vec![0, 1, 2]).unwrap_err();
        assert!(matches!(
            error,
            Error::Source(SourceBackedError::SourceChanged { .. })
        ));
    }
}
