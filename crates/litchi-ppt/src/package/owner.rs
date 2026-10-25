//! Canonical immutable owner for legacy PowerPoint packages.
//!
//! [`Package`] is the source-preserving owner used by the unified operations
//! API. It retains the validated CFB view and the parsed presentation
//! projection together. Cloning a package therefore clones one `Arc`; it does
//! not reopen the container or parse the presentation again.
//!
//! The eager constructors move the already-parsed CFB into
//! [`litchi_cfb::SharedOleFile::from_owned_ole_file`]. The positional
//! constructors adopt [`super::SourceBackedPackage`]'s existing CFB index.
//! Both routes preserve the original CFB bytes for exact publication.

use super::{Error, OpenOptions, Result, SourceBackedPackage, SourceLimits};
use crate::presentation::Presentation;
use litchi_cfb::{OleError, OleFile, SharedOleFile, SharedOleFileLimits};
#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{Durability, Metadata, OwnedSource, ReadAt, SourceVersion};
use litchi_ole_common::property_set::SharedPropertySetReader;
use std::fmt;
use std::io::{self, Cursor, Read, Write};
use std::path::Path;
use std::sync::Arc;

/// The immutable state retained by one canonical package.
struct PackageInner {
    shared: Arc<SharedOleFile>,
    presentation: Presentation,
    source_limits: SourceLimits,
    leniency: crate::leniency::Leniency,
    /// The digest facts of the exact package bytes (block digests, block
    /// root, signature presence), computed once on demand or seeded by the
    /// scoped publisher that produced them (records 0357a, 0357b).
    exact: crate::scoped::ExactCache,
}

/// An immutable, cheaply cloned legacy PowerPoint package.
///
/// The package retains the exact source CFB and a parsed [`Presentation`]
/// projection. It is safe to send and share between threads when the native
/// presentation projection is enabled by the crate's normal feature set.
/// Ordinary publication preserves the original bytes and refuses encrypted
/// sources; [`Self::to_encrypted_bytes`] is the explicit encrypted pass-through
/// route.
///
/// # Example
///
/// ```rust,no_run
/// use litchi_ppt::Package;
///
/// let package = Package::open("presentation.ppt")?;
/// let presentation = package.presentation()?;
/// println!("{} slides", presentation.slide_count());
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package {
    inner: Arc<PackageInner>,
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
            .field("source_limits", &self.inner.source_limits)
            .field("leniency", &self.inner.leniency)
            .field("encrypted", &self.is_encrypted())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Build the smallest valid plaintext legacy PowerPoint package.
    ///
    /// The empty artifact contains zero slides and is owned by the same
    /// immutable package type as a package opened from disk.
    pub fn new() -> Result<Self> {
        crate::Builder::new().build()
    }

    /// Open a package from a copied byte slice under the default policies.
    ///
    /// The copy is fallibly allocated and is retained by the sealed CFB owner;
    /// opening does not create a second byte buffer.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(bytes, OpenOptions::default())
    }

    /// Open a package from a copied byte slice under explicit policies.
    pub fn from_bytes_with_options(bytes: &[u8], options: OpenOptions<'_>) -> Result<Self> {
        let maximum = options.source_limits_value().max_input_bytes();
        if bytes.len() > maximum {
            return Err(Error::Limit {
                kind: super::SourceLimitKind::InputBytes,
                observed: bytes.len(),
                maximum,
            });
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|_| Error::AllocationFailed("PPT package input"))?;
        owned.extend_from_slice(bytes);
        Self::from_owned_bytes(owned, options)
    }

    /// Read a package from a sequential source under the default policies.
    ///
    /// The source is read in bounded chunks and never grows beyond the
    /// selected complete-input limit.
    pub fn from_reader<R: Read>(mut reader: R) -> Result<Self> {
        Self::from_reader_with_options(&mut reader, OpenOptions::default())
    }

    /// Read a package from a sequential source under explicit policies.
    ///
    /// At most one byte beyond the configured input limit is read, so an
    /// oversized sequential source is rejected before an unbounded allocation
    /// can occur.
    pub fn from_reader_with_options<R: Read>(
        reader: &mut R,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        let maximum = options.source_limits_value().max_input_bytes();
        let mut bytes = Vec::new();
        let mut buffer = [0_u8; 64 * 1024];
        loop {
            let remaining = maximum.saturating_add(1).saturating_sub(bytes.len());
            if remaining == 0 {
                break;
            }
            let request = remaining.min(buffer.len());
            let read = match reader.read(&mut buffer[..request]) {
                Ok(read) if read <= request => read,
                Ok(_) => return Err(Error::Io(io::Error::from(io::ErrorKind::InvalidData))),
                Err(error) if error.kind() == io::ErrorKind::Interrupted => continue,
                Err(error) => return Err(Error::Io(error)),
            };
            let observed = bytes.len().saturating_add(read);
            if observed > maximum {
                return Err(Error::Limit {
                    kind: super::SourceLimitKind::InputBytes,
                    observed,
                    maximum,
                });
            }
            if read == 0 {
                break;
            }
            bytes
                .try_reserve(read)
                .map_err(|_| Error::AllocationFailed("PPT package input"))?;
            bytes.extend_from_slice(&buffer[..read]);
        }
        Self::from_owned_bytes(bytes, options)
    }

    /// Open a filesystem-backed package under the default policies.
    #[cfg(any(unix, windows))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open a filesystem-backed package with a password-to-open credential.
    #[cfg(any(unix, windows))]
    pub fn open_with_password(path: impl AsRef<Path>, password: &str) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::new().password(password))
    }

    /// Open a filesystem-backed package under explicit policies.
    #[cfg(any(unix, windows))]
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions<'_>) -> Result<Self> {
        let source = FileSource::open(path).map_err(Error::Io)?;
        Self::from_read_at_with_options(Arc::new(source), options)
    }

    /// Adopt a positional source under the default policies.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_options(source, OpenOptions::default())
    }

    /// Adopt a positional source under explicit policies.
    ///
    /// The source-backed CFB index is validated once by
    /// [`SourceBackedPackage`]. The presentation projection reuses that index
    /// and does not reopen the container.
    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        let source_package = SourceBackedPackage::from_read_at_with_options(source, options)?;
        Self::from_source(source_package, options)
    }

    /// Adopt an already-parsed in-memory CFB file under explicit policies.
    ///
    /// The parsed CFB index moves into the retained shared owner without a
    /// second container parse or a byte copy.
    pub fn from_ole_file(ole: OleFile<Cursor<Vec<u8>>>, options: OpenOptions<'_>) -> Result<Self> {
        let mut reader = crate::raw::Reader::from_ole_file_with_options(ole, options)?;
        let presentation = reader.presentation_with_options(options)?;
        let shared = seal_ole_file(reader.into_ole_file())?;
        Ok(Self::from_parts(shared, presentation, options))
    }

    /// Adopt an already-parsed CFB file while preserving a password handoff.
    ///
    /// When the native reader reports [`Error::PasswordRequired`], the
    /// unchanged parsed CFB is returned to the caller. This lets a facade ask
    /// for a password without reparsing or discarding the container.
    pub fn from_ole_file_or_password_required(
        ole: OleFile<Cursor<Vec<u8>>>,
        options: OpenOptions<'_>,
    ) -> Result<std::result::Result<Self, OleFile<Cursor<Vec<u8>>>>> {
        let mut reader = crate::raw::Reader::from_ole_file_with_options(ole, options)?;
        let presentation = match reader.presentation_with_options(options) {
            Ok(presentation) => presentation,
            Err(Error::PasswordRequired) => return Ok(Err(reader.into_ole_file())),
            #[cfg(not(feature = "encryption"))]
            Err(Error::UnsupportedEncryption(super::EncryptionKind::CryptoApi)) => {
                return Ok(Err(reader.into_ole_file()));
            },
            Err(error) => return Err(error),
        };
        let shared = seal_ole_file(reader.into_ole_file())?;
        Ok(Ok(Self::from_parts(shared, presentation, options)))
    }

    /// Adopt a validated positional package under explicit policies.
    ///
    /// The source package's CFB index and the resulting presentation are
    /// retained directly. `options` supplies only the requested open and
    /// leniency policy; its password is never stored in the owner.
    pub fn from_source(source: SourceBackedPackage, options: OpenOptions<'_>) -> Result<Self> {
        let shared = source.shared_ole_file_arc();
        let presentation = source.presentation_with_options(options)?;
        let limits = presentation.source_limits();
        let leniency = options.leniency_value();
        shared.source_version()?;
        Ok(Self {
            inner: Arc::new(PackageInner {
                shared,
                presentation,
                source_limits: limits,
                leniency,
                exact: crate::scoped::ExactCache::default(),
            }),
        })
    }

    pub(crate) fn from_owned_bytes_for_builder(
        bytes: Vec<u8>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        Self::from_owned_bytes(bytes, options)
    }

    /// Open a package from owned bytes under explicit policies, without the
    /// copy [`Self::from_bytes_with_options`] makes; the input limit applies.
    pub(crate) fn from_owned_bytes_with_options(
        bytes: Vec<u8>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        let maximum = options.source_limits_value().max_input_bytes();
        if bytes.len() > maximum {
            return Err(Error::Limit {
                kind: super::SourceLimitKind::InputBytes,
                observed: bytes.len(),
                maximum,
            });
        }
        Self::from_owned_bytes(bytes, options)
    }

    fn from_owned_bytes(bytes: Vec<u8>, options: OpenOptions<'_>) -> Result<Self> {
        let mut reader = crate::raw::Reader::from_reader_with_options(Cursor::new(bytes), options)?;
        let presentation = reader.presentation_with_options(options)?;
        let shared = seal_ole_file(reader.into_ole_file())?;
        Ok(Self::from_parts(shared, presentation, options))
    }

    fn from_parts(
        shared: Arc<SharedOleFile>,
        presentation: Presentation,
        options: OpenOptions<'_>,
    ) -> Self {
        let source_limits = presentation.source_limits();
        Self {
            inner: Arc::new(PackageInner {
                shared,
                presentation,
                source_limits,
                leniency: options.leniency_value(),
                exact: crate::scoped::ExactCache::default(),
            }),
        }
    }

    /// The digest facts of the package's exact bytes, for the scoped
    /// families' patches (record 0357b).
    pub(crate) fn exact(&self) -> &crate::scoped::ExactCache {
        &self.inner.exact
    }

    /// Check the already-admitted, root-reachable CFB directory for a
    /// signature component.
    ///
    /// Scoped publication only needs the component's presence. Reusing the
    /// retained directory avoids reparsing the exact bytes for every replay;
    /// the source fences keep the answer tied to the same positional source
    /// version as the package bytes. The traversal follows the same reachable
    /// storage paths as the eager reader, rather than treating an orphaned
    /// physical directory slot as a signature. When the event budget covers
    /// the admitted index, the same reachable set is scanned directly (with
    /// the root entry excluded); tighter budgets use the bounded path walk.
    pub(crate) fn contains_signature_component(&self) -> Result<bool> {
        let expected = self.inner.shared.source_version()?;
        let directory_entry_count = self.inner.shared.directory_entries().count();
        let maximum = self
            .inner
            .source_limits
            .max_events()
            .min(directory_entry_count);
        if self.inner.source_limits.max_events() >= directory_entry_count {
            let found = self
                .inner
                .shared
                .directory_entries()
                .filter(|entry| entry.entry_type != litchi_cfb::consts::STGTY_ROOT)
                .any(|entry| crate::package::codec::is_signature_component(&entry.name));
            ensure_source_version(&self.inner.shared, expected)?;
            return Ok(found);
        }
        let mut pending = vec![Vec::<String>::new()];
        let mut discovered = 0usize;
        let mut found = false;
        while let Some(path) = pending.pop() {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            let remaining = maximum.saturating_sub(discovered);
            self.inner
                .shared
                .visit_directory_entry_refs(&refs, remaining, |entry| {
                    discovered = discovered.saturating_add(1);
                    if crate::package::codec::is_signature_component(&entry.name) {
                        found = true;
                    }
                    if entry.entry_type == litchi_cfb::consts::STGTY_STORAGE {
                        let mut child = path.clone();
                        child.push(entry.name.clone());
                        pending.push(child);
                    }
                    Ok::<(), Error>(())
                })?;
            if found {
                break;
            }
        }
        ensure_source_version(&self.inner.shared, expected)?;
        Ok(found)
    }

    /// Seed the presentation's one-slide shape memo with shapes proven equal
    /// to a parse of slide `persist_id` of this package (record 0264).
    pub(crate) fn seed_slide_shapes(
        &self,
        persist_id: u32,
        shapes: Arc<crate::slide::PatchedShapes>,
    ) {
        self.inner
            .presentation
            .inner
            .shape_memo
            .seed(persist_id, shapes);
    }

    /// Return a cloned read projection after checking positional freshness.
    pub fn presentation(&self) -> Result<Presentation> {
        let expected = self.inner.shared.source_version()?;
        let presentation = self.inner.presentation.clone();
        ensure_source_version(&self.inner.shared, expected)?;
        Ok(presentation)
    }

    /// Return the standard OLE metadata projected into the core API.
    pub fn metadata(&self) -> Result<Metadata> {
        let expected = self.inner.shared.source_version()?;
        let metadata =
            <SharedOleFile as SharedPropertySetReader>::get_metadata(self.inner.shared.as_ref())
                .map(Metadata::from)?;
        ensure_source_version(&self.inner.shared, expected)?;
        Ok(metadata)
    }

    /// Return the finite source policy retained by this package.
    #[must_use]
    pub fn source_limits(&self) -> SourceLimits {
        self.inner.source_limits
    }

    /// Return the leniency policy retained by this package.
    #[must_use]
    pub fn leniency(&self) -> crate::leniency::Leniency {
        self.inner.leniency
    }

    /// Return the VBA inspection profile retained by this package.
    #[cfg(feature = "vba-inspection")]
    #[must_use]
    pub(crate) fn vba_project_limits(&self) -> crate::VbaProjectLimits {
        self.inner.presentation.vba_project_limits()
    }

    /// Check and return the current source identity.
    pub fn source_version(&self) -> Result<SourceVersion> {
        Ok(self.inner.shared.source_version()?)
    }

    /// Borrow the retained shared CFB view for read-only editor handoff.
    #[must_use]
    pub fn shared_ole_file(&self) -> &SharedOleFile {
        self.inner.shared.as_ref()
    }

    /// Clone the retained shared CFB view for read-only editor handoff.
    #[must_use]
    pub fn shared_ole_file_arc(&self) -> Arc<SharedOleFile> {
        Arc::clone(&self.inner.shared)
    }

    /// Whether the retained presentation uses password-to-open encryption.
    #[must_use]
    pub fn is_encrypted(&self) -> bool {
        self.inner.presentation.is_encrypted()
    }

    /// Return the exact native encryption profile when encryption support is
    /// enabled.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub fn encryption(&self) -> Option<crate::EncryptionProfile> {
        self.inner.presentation.encryption()
    }

    /// Return the exact original bytes, including encrypted sources.
    pub fn to_encrypted_bytes(&self) -> Result<Vec<u8>> {
        exact_bytes(&self.inner.shared)
    }

    /// Return bytes suitable for ordinary publication.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        ensure_publication_allowed(self)?;
        self.to_encrypted_bytes()
    }

    /// Write the exact source bytes to a sequential sink.
    pub fn write_to<W: Write>(&self, writer: &mut W) -> Result<()> {
        ensure_publication_allowed(self)?;
        stream_exact(&self.inner.shared, writer)
    }

    /// Atomically replace a destination with the exact source bytes.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.save_with_durability(path, Durability::Full)
    }

    /// Atomically replace a destination with explicit durability.
    pub fn save_with_durability<P: AsRef<Path>>(
        &self,
        path: P,
        durability: Durability,
    ) -> Result<()> {
        ensure_publication_allowed(self)?;
        let shared = Arc::clone(&self.inner.shared);
        let verify_shared = Arc::clone(&shared);
        let expected = shared.source_version()?;
        litchi_cfb::atomic::save(
            path,
            durability,
            litchi_cfb::atomic::Overwrite::Replace,
            move |writer| stream_exact_with_expected(&shared, expected, writer),
            move || ensure_source_version(&verify_shared, expected),
        )
    }
}

fn seal_ole_file(ole: OleFile<Cursor<Vec<u8>>>) -> Result<Arc<SharedOleFile>> {
    // `OwnedSource` mints the same process-local immutable version policy used
    // by the existing facade's sealed CFB handoff. The parsed directory/FAT
    // index moves into SharedOleFile; no second parse or byte copy occurs.
    let version = OwnedSource::new(Vec::new()).version()?;
    let shared = SharedOleFile::from_owned_ole_file(ole, version, SharedOleFileLimits::default())?;
    Ok(Arc::new(shared))
}

fn ensure_publication_allowed(package: &Package) -> Result<()> {
    if package.is_encrypted() {
        return Err(Error::EncryptionPolicy);
    }
    Ok(())
}

fn ensure_source_version(shared: &SharedOleFile, expected: SourceVersion) -> Result<()> {
    let observed = shared.source_version()?;
    if observed != expected {
        return Err(Error::Ole(OleError::SourceChanged { expected, observed }));
    }
    Ok(())
}

fn exact_bytes(shared: &SharedOleFile) -> Result<Vec<u8>> {
    let expected = shared.source_version()?;
    let length = shared.file_size();
    let length = usize::try_from(length).map_err(|_| Error::Limit {
        kind: super::SourceLimitKind::InputBytes,
        observed: usize::MAX,
        maximum: SourceLimits::MAX_INPUT_BYTES,
    })?;
    let source = shared.source_arc();
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(length)
        .map_err(|_| Error::AllocationFailed("PPT package output"))?;
    bytes.resize(length, 0);
    let mut offset = 0_u64;
    while offset < u64::try_from(length).unwrap_or(u64::MAX) {
        let remaining = u64::try_from(length).unwrap_or(u64::MAX) - offset;
        let count = usize::try_from(remaining.min(64 * 1024)).unwrap_or(64 * 1024);
        let offset_usize = usize::try_from(offset).map_err(|_| Error::Limit {
            kind: super::SourceLimitKind::InputBytes,
            observed: usize::MAX,
            maximum: SourceLimits::MAX_INPUT_BYTES,
        })?;
        source
            .read_exact_at(offset, &mut bytes[offset_usize..offset_usize + count])
            .map_err(|error| source_read_error(shared, expected, error))?;
        offset = offset
            .checked_add(u64::try_from(count).unwrap_or(u64::MAX))
            .ok_or_else(|| Error::Io(io::Error::other("PPT source offset overflow")))?;
    }
    ensure_source_version(shared, expected)?;
    Ok(bytes)
}

fn stream_exact<W: Write>(shared: &SharedOleFile, writer: &mut W) -> Result<()> {
    let expected = shared.source_version()?;
    stream_exact_with_expected(shared, expected, writer)
}

fn stream_exact_with_expected<W: Write>(
    shared: &SharedOleFile,
    expected: SourceVersion,
    writer: &mut W,
) -> Result<()> {
    ensure_source_version(shared, expected)?;
    let source = shared.source_arc();
    let length = shared.file_size();
    let mut offset = 0_u64;
    let mut buffer = [0_u8; 64 * 1024];
    while offset < length {
        let count = usize::try_from((length - offset).min(buffer.len() as u64)).map_err(|_| {
            Error::Limit {
                kind: super::SourceLimitKind::InputBytes,
                observed: usize::MAX,
                maximum: SourceLimits::MAX_INPUT_BYTES,
            }
        })?;
        source
            .read_exact_at(offset, &mut buffer[..count])
            .map_err(|error| source_read_error(shared, expected, error))?;
        writer.write_all(&buffer[..count])?;
        offset = offset
            .checked_add(u64::try_from(count).unwrap_or(u64::MAX))
            .ok_or_else(|| Error::Io(io::Error::other("PPT source offset overflow")))?;
    }
    ensure_source_version(shared, expected)
}

fn source_read_error(shared: &SharedOleFile, expected: SourceVersion, error: io::Error) -> Error {
    match shared.source_version() {
        Ok(observed) if observed != expected => {
            Error::Ole(OleError::SourceChanged { expected, observed })
        },
        Ok(_) => Error::Io(error),
        Err(version_error) => Error::Ole(version_error),
    }
}
