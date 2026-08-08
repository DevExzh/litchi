//! Immutable ingress for legacy directory-backed iWork bundles.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "filesystem acquisition helpers stay beside the state they validate"
)]

use std::{
    ffi::OsString,
    fmt,
    fs::{self, Metadata, OpenOptions},
    io::Read,
    path::{Path, PathBuf},
    sync::Arc,
    time::SystemTime,
};

use crate::{ComponentCatalog, Error, LimitKind, Limits, Result};

const READ_CHUNK_BYTES: usize = 16 * 1024;

/// Physical representation captured from a legacy directory bundle.
///
/// Neither variant is an exact ZIP representation of the complete document.
/// In particular, [`Self::IndexZip`] identifies only the directory's index
/// subartifact and must never be passed to package preserve-mode output as if
/// it were a complete `.pages`, `.numbers`, or `.key` file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DirectoryProvenance {
    /// The bundle contained one direct-IWA `Index.zip` file.
    IndexZip,
    /// The bundle contained an unpacked `Index/` directory.
    LooseIndex,
}

/// Legacy application-marker evidence captured beside a directory index.
///
/// The archive layer records presence without assigning application meaning;
/// the focused detector remains responsible for reconciling these markers
/// with the canonical root `Document.iwa` payload.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct DirectoryMarkers {
    pages: bool,
    keynote: bool,
    numbers: bool,
}

impl DirectoryMarkers {
    /// Whether the bundle contained the legacy Pages `index.xml` marker.
    #[must_use]
    pub const fn pages(self) -> bool {
        self.pages
    }

    /// Whether the bundle contained the legacy Keynote `index.apxl` marker.
    #[must_use]
    pub const fn keynote(self) -> bool {
        self.keynote
    }

    /// Whether the bundle contained the legacy Numbers `index.numbers`
    /// marker.
    #[must_use]
    pub const fn numbers(self) -> bool {
        self.numbers
    }
}

/// One immutable logical member captured from an unpacked `Index/` directory.
pub struct FrozenDirectoryEntry {
    name: Box<str>,
    data: Box<[u8]>,
}

impl FrozenDirectoryEntry {
    /// Return the normalized `Index/<member>` name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the captured member bytes.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }
}

impl fmt::Debug for FrozenDirectoryEntry {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FrozenDirectoryEntry")
            .field("name", &self.name)
            .field("bytes", &self.data.len())
            .finish()
    }
}

/// A bounded immutable snapshot of one legacy directory bundle's IWA index.
///
/// Opening copies every in-scope source byte before publishing the value. The
/// snapshot therefore performs no ambient filesystem I/O during later
/// component traversal and is cheaply cloneable through shared state.
/// `Metadata/` and `Data/` are deliberately outside this index adapter; a
/// caller must not infer that they were preserved merely because this value
/// exists.
#[derive(Clone)]
pub struct FrozenDirectoryBundle {
    state: Arc<State>,
}

struct State {
    provenance: DirectoryProvenance,
    markers: DirectoryMarkers,
    limits: Limits,
    components: Arc<ComponentCatalog>,
    storage: Storage,
}

enum Storage {
    IndexZip(Box<[u8]>),
    LooseIndex(Box<[FrozenDirectoryEntry]>),
}

impl fmt::Debug for FrozenDirectoryBundle {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("FrozenDirectoryBundle")
            .field("provenance", &self.state.provenance)
            .field("markers", &self.state.markers)
            .field("limits", &self.state.limits)
            .field("components", &self.state.components.len())
            .field("logical_entries", &self.loose_entries().len())
            .field("index_zip_bytes", &self.index_zip_bytes().map(<[u8]>::len))
            .finish()
    }
}

impl FrozenDirectoryBundle {
    /// Snapshot a legacy directory bundle under the default physical limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the root or index representation is missing,
    /// ambiguous, unsafe, malformed, changes during capture, or exceeds a
    /// physical resource ceiling.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_limits(path, Limits::default())
    }

    /// Snapshot a legacy directory bundle under explicit physical limits.
    ///
    /// Exactly one of `Index.zip` and `Index/` must exist. Symbolic links and
    /// special filesystem nodes are rejected at every traversed boundary.
    /// Loose entries are normalized and sorted before their IWA streams are
    /// decoded, so filesystem enumeration order cannot affect semantics.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::open`].
    pub fn open_with_limits(path: impl AsRef<Path>, limits: Limits) -> Result<Self> {
        let checked_limits = limits.validate()?;
        let root = path.as_ref();
        let root_before = require_directory(root, "directory bundle root")?;
        let markers_before = snapshot_markers(root)?;
        let index_zip_path = root.join("Index.zip");
        let index_path = root.join("Index");
        let index_zip = inspect_node(&index_zip_path, "directory bundle Index.zip")?;
        let index = inspect_node(&index_path, "directory bundle Index")?;

        let snapshot = match (index_zip, index) {
            (Node::File(version), Node::Missing) => Self::from_index_zip(
                &index_zip_path,
                version,
                markers_before.markers,
                checked_limits,
            )?,
            (Node::Missing, Node::Directory(version)) => Self::from_loose_index(
                &index_path,
                version,
                markers_before.markers,
                checked_limits,
            )?,
            (Node::File(_), Node::Directory(_)) => {
                return Err(Error::InvalidBundle(
                    "directory bundle contains both Index.zip and Index/ representations"
                        .to_owned(),
                ));
            },
            (Node::Missing, Node::Missing) => {
                return Err(Error::InvalidBundle(
                    "directory bundle contains neither Index.zip nor Index/".to_owned(),
                ));
            },
            _ => {
                return Err(Error::InvalidBundle(
                    "directory bundle index representation has an invalid node type".to_owned(),
                ));
            },
        };

        let markers_after = snapshot_markers(root)?;
        if markers_after != markers_before {
            return Err(changed("verifying directory bundle application markers"));
        }
        ensure_path_version(root, &root_before, "verifying directory bundle root")?;
        match snapshot.state.provenance {
            DirectoryProvenance::IndexZip => {
                ensure_file_without_peer(
                    &index_zip_path,
                    &index_path,
                    "verifying directory bundle Index.zip representation",
                )?;
            },
            DirectoryProvenance::LooseIndex => {
                ensure_directory_without_peer(
                    &index_path,
                    &index_zip_path,
                    "verifying directory bundle Index/ representation",
                )?;
            },
        }
        Ok(snapshot)
    }

    fn from_index_zip(
        path: &Path,
        expected: FileVersion,
        markers: DirectoryMarkers,
        limits: Limits,
    ) -> Result<Self> {
        let bytes = read_stable_file(path, &expected, limits, FileRole::IndexZip)?;
        let components = ComponentCatalog::from_directory_index_zip(&bytes, limits)?;
        if components.is_empty() {
            return Err(Error::InvalidBundle(
                "directory bundle Index.zip contains no decodable IWA components".to_owned(),
            ));
        }
        Ok(Self {
            state: Arc::new(State {
                provenance: DirectoryProvenance::IndexZip,
                markers,
                limits,
                components: Arc::new(components),
                storage: Storage::IndexZip(bytes),
            }),
        })
    }

    fn from_loose_index(
        path: &Path,
        expected: FileVersion,
        markers: DirectoryMarkers,
        limits: Limits,
    ) -> Result<Self> {
        let manifest = scan_manifest(path, limits)?;
        if manifest.is_empty() {
            return Err(Error::InvalidBundle(
                "directory bundle Index/ contains no entries".to_owned(),
            ));
        }
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(manifest.len())
            .map_err(|_error| Error::Allocation {
                resource: "directory bundle entries",
                amount: manifest.len(),
            })?;
        for item in &manifest {
            let data = read_stable_file(&item.path, &item.version, limits, FileRole::LooseEntry)?;
            entries.push(FrozenDirectoryEntry {
                name: item.logical_name.clone(),
                data,
            });
        }

        let observed = scan_manifest(path, limits)?;
        if manifest != observed {
            return Err(changed("verifying directory bundle Index/ manifest"));
        }
        ensure_path_version(
            path,
            &expected,
            "verifying directory bundle Index/ directory",
        )?;

        let components = ComponentCatalog::from_logical_entries(
            entries.iter().map(|entry| (entry.name(), entry.data())),
            limits,
        )?;
        if components.is_empty() {
            return Err(Error::InvalidBundle(
                "directory bundle Index/ contains no decodable IWA components".to_owned(),
            ));
        }
        Ok(Self {
            state: Arc::new(State {
                provenance: DirectoryProvenance::LooseIndex,
                markers,
                limits,
                components: Arc::new(components),
                storage: Storage::LooseIndex(entries.into_boxed_slice()),
            }),
        })
    }

    /// Return the retained physical ingress profile.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.state.limits
    }

    /// Return which directory-index representation was captured.
    #[must_use]
    pub fn provenance(&self) -> DirectoryProvenance {
        self.state.provenance
    }

    /// Return the frozen legacy marker evidence captured beside the index.
    #[must_use]
    pub fn markers(&self) -> DirectoryMarkers {
        self.state.markers
    }

    /// Borrow parsed components in deterministic normalized-name order.
    #[must_use]
    pub fn components(&self) -> &ComponentCatalog {
        &self.state.components
    }

    /// Consume this snapshot into a shared immutable component catalog.
    ///
    /// This is the semantic-only handoff for directory-backed documents. It
    /// deliberately cannot yield a [`crate::SourceCatalog`], because neither
    /// directory representation is an exact ZIP source for the whole iWork
    /// artifact.
    #[must_use]
    pub fn into_components(self) -> Arc<ComponentCatalog> {
        match Arc::try_unwrap(self.state) {
            Ok(state) => state.components,
            Err(state) => Arc::clone(&state.components),
        }
    }

    /// Borrow the exact `Index.zip` subartifact, if that representation was
    /// captured.
    ///
    /// These bytes are not the exact representation of the complete directory
    /// bundle and are therefore intentionally separate from
    /// `SourceCatalog::source_bytes`.
    #[must_use]
    pub fn index_zip_bytes(&self) -> Option<&[u8]> {
        match &self.state.storage {
            Storage::IndexZip(bytes) => Some(bytes),
            Storage::LooseIndex(_) => None,
        }
    }

    /// Borrow normalized loose index entries, or an empty slice for an
    /// `Index.zip`-backed directory.
    #[must_use]
    pub fn loose_entries(&self) -> &[FrozenDirectoryEntry] {
        match &self.state.storage {
            Storage::IndexZip(_) => &[],
            Storage::LooseIndex(entries) => entries,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MarkerSnapshot {
    markers: DirectoryMarkers,
    pages: Option<FileVersion>,
    keynote: Option<FileVersion>,
    numbers: Option<FileVersion>,
}

fn snapshot_markers(root: &Path) -> Result<MarkerSnapshot> {
    let pages = inspect_marker(&root.join("index.xml"), "legacy Pages marker")?;
    let keynote = inspect_marker(&root.join("index.apxl"), "legacy Keynote marker")?;
    let numbers = inspect_marker(&root.join("index.numbers"), "legacy Numbers marker")?;
    Ok(MarkerSnapshot {
        markers: DirectoryMarkers {
            pages: pages.is_some(),
            keynote: keynote.is_some(),
            numbers: numbers.is_some(),
        },
        pages,
        keynote,
        numbers,
    })
}

fn inspect_marker(path: &Path, label: &str) -> Result<Option<FileVersion>> {
    match inspect_node(path, label)? {
        Node::Missing => Ok(None),
        Node::File(version) => Ok(Some(version)),
        Node::Directory(_) => Err(Error::InvalidBundle(format!(
            "{label} is not a regular file"
        ))),
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ManifestEntry {
    logical_name: Box<str>,
    path: PathBuf,
    version: FileVersion,
}

fn scan_manifest(path: &Path, limits: Limits) -> Result<Vec<ManifestEntry>> {
    let mut entries = Vec::new();
    let mut total_bytes = 0u64;
    let mut metadata_bytes = 0u64;
    for entry_result in fs::read_dir(path)? {
        let directory_entry = entry_result?;
        let file_name = utf8_name(directory_entry.file_name())?;
        if file_name == "Index.zip" {
            return Err(Error::InvalidBundle(
                "directory bundle loose Index/ contains a nested Index.zip".to_owned(),
            ));
        }
        let logical_name = format!("Index/{file_name}");
        let name_bytes = u64::try_from(logical_name.len()).map_err(|_error| {
            Error::InvalidBundle("directory entry name length does not fit u64".to_owned())
        })?;
        check_member_name(name_bytes, limits)?;
        metadata_bytes = metadata_bytes.checked_add(name_bytes).ok_or_else(|| {
            Error::InvalidBundle("directory entry metadata byte count overflowed".to_owned())
        })?;
        check_metadata_bytes(metadata_bytes, limits)?;

        let entry_path = directory_entry.path();
        let node = inspect_node(&entry_path, "directory bundle Index/ member")?;
        let Node::File(version) = node else {
            return Err(Error::InvalidBundle(format!(
                "directory bundle Index/ member {logical_name} is not a regular file"
            )));
        };
        check_loose_entry_size(version.len, limits)?;
        total_bytes = total_bytes.checked_add(version.len).ok_or_else(|| {
            Error::InvalidBundle("directory bundle Index/ byte count overflowed".to_owned())
        })?;
        check_total_bytes(total_bytes, limits)?;
        entries.try_reserve(1).map_err(|_error| Error::Allocation {
            resource: "directory bundle manifest",
            amount: 1,
        })?;
        entries.push(ManifestEntry {
            logical_name: logical_name.into_boxed_str(),
            path: entry_path,
            version,
        });
        check_entry_count(entries.len(), limits)?;
    }
    entries.sort_unstable_by(|left, right| left.logical_name.cmp(&right.logical_name));
    Ok(entries)
}

fn utf8_name(name: OsString) -> Result<String> {
    name.into_string().map_err(|_name| {
        Error::InvalidBundle("directory bundle Index/ contains a non-UTF-8 member name".to_owned())
    })
}

#[derive(Debug, Clone, Copy)]
enum FileRole {
    IndexZip,
    LooseEntry,
}

fn read_stable_file(
    path: &Path,
    expected: &FileVersion,
    limits: Limits,
    role: FileRole,
) -> Result<Box<[u8]>> {
    check_file_size(expected.len, limits, role)?;
    let before_path = require_file(path, "directory bundle member")?;
    if &before_path != expected {
        return Err(changed("opening directory bundle member"));
    }

    let mut open = OpenOptions::new();
    open.read(true);
    #[cfg(unix)]
    {
        use std::os::unix::fs::OpenOptionsExt;

        open.custom_flags(libc::O_CLOEXEC | libc::O_NOFOLLOW | libc::O_NONBLOCK);
    }
    let mut file = open.open(path)?;
    let opened_metadata = file.metadata()?;
    let opened = FileVersion::from_metadata(&opened_metadata);
    if opened != *expected || !opened_metadata.is_file() {
        return Err(changed("opening directory bundle member"));
    }
    ensure_path_version(path, expected, "opening directory bundle member")?;

    let mut bytes = Vec::new();
    let initial = usize::try_from(expected.len)
        .unwrap_or(READ_CHUNK_BYTES)
        .min(READ_CHUNK_BYTES);
    bytes
        .try_reserve_exact(initial)
        .map_err(|_error| Error::Allocation {
            resource: "directory bundle member bytes",
            amount: initial,
        })?;
    let mut buffer = [0u8; READ_CHUNK_BYTES];
    loop {
        let read = match file.read(&mut buffer) {
            Ok(read) => read,
            Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
            Err(error) => return Err(error.into()),
        };
        if read == 0 {
            break;
        }
        let observed = bytes.len().checked_add(read).ok_or_else(|| {
            Error::InvalidBundle("directory bundle member length overflowed usize".to_owned())
        })?;
        let observed_u64 = u64::try_from(observed).map_err(|_error| {
            Error::InvalidBundle("directory bundle member length does not fit u64".to_owned())
        })?;
        check_file_size(observed_u64, limits, role)?;
        bytes
            .try_reserve(read)
            .map_err(|_error| Error::Allocation {
                resource: "directory bundle member bytes",
                amount: read,
            })?;
        bytes.extend_from_slice(&buffer[..read]);
    }

    let after = FileVersion::from_metadata(&file.metadata()?);
    if after != opened {
        return Err(changed("reading directory bundle member"));
    }
    ensure_path_version(path, expected, "verifying directory bundle member")?;
    if u64::try_from(bytes.len()).ok() != Some(expected.len) {
        return Err(changed("reading directory bundle member"));
    }
    Ok(bytes.into_boxed_slice())
}

fn check_file_size(observed: u64, limits: Limits, role: FileRole) -> Result<()> {
    if observed > limits.max_input_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed,
            maximum: limits.max_input_bytes(),
        });
    }
    if matches!(role, FileRole::LooseEntry) && observed > limits.max_entry_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::EntryBytes,
            observed,
            maximum: limits.max_entry_bytes(),
        });
    }
    Ok(())
}

fn check_loose_entry_size(observed: u64, limits: Limits) -> Result<()> {
    check_file_size(observed, limits, FileRole::LooseEntry)
}

fn check_entry_count(observed: usize, limits: Limits) -> Result<()> {
    if observed > limits.max_entries() {
        return Err(Error::Limit {
            kind: LimitKind::Entries,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(limits.max_entries()).unwrap_or(u64::MAX),
        });
    }
    Ok(())
}

fn check_member_name(observed: u64, limits: Limits) -> Result<()> {
    let maximum = limits.max_input_bytes().min(Limits::MAX_MEMBER_NAME_BYTES);
    if observed > maximum {
        return Err(Error::Limit {
            kind: LimitKind::MemberNameBytes,
            observed,
            maximum,
        });
    }
    Ok(())
}

fn check_metadata_bytes(observed: u64, limits: Limits) -> Result<()> {
    let maximum = limits.max_input_bytes().min(Limits::MAX_METADATA_BYTES);
    if observed > maximum {
        return Err(Error::Limit {
            kind: LimitKind::MetadataBytes,
            observed,
            maximum,
        });
    }
    Ok(())
}

fn check_total_bytes(observed: u64, limits: Limits) -> Result<()> {
    if observed > limits.max_total_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::TotalBytes,
            observed,
            maximum: limits.max_total_bytes(),
        });
    }
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Node {
    Missing,
    File(FileVersion),
    Directory(FileVersion),
}

fn inspect_node(path: &Path, label: &str) -> Result<Node> {
    match fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => Err(Error::InvalidBundle(format!(
            "{label} must not be a symbolic link"
        ))),
        Ok(metadata) if metadata.is_file() => Ok(Node::File(FileVersion::from_metadata(&metadata))),
        Ok(metadata) if metadata.is_dir() => {
            Ok(Node::Directory(FileVersion::from_metadata(&metadata)))
        },
        Ok(_) => Err(Error::InvalidBundle(format!(
            "{label} is not a regular file or directory"
        ))),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(Node::Missing),
        Err(error) => Err(error.into()),
    }
}

fn require_file(path: &Path, label: &str) -> Result<FileVersion> {
    match inspect_node(path, label)? {
        Node::File(version) => Ok(version),
        Node::Missing => Err(Error::InvalidBundle(format!("{label} is missing"))),
        Node::Directory(_) => Err(Error::InvalidBundle(format!(
            "{label} is not a regular file"
        ))),
    }
}

fn require_directory(path: &Path, label: &str) -> Result<FileVersion> {
    match inspect_node(path, label)? {
        Node::Directory(version) => Ok(version),
        Node::Missing => Err(Error::InvalidBundle(format!("{label} is missing"))),
        Node::File(_) => Err(Error::InvalidBundle(format!("{label} is not a directory"))),
    }
}

fn ensure_path_version(path: &Path, expected: &FileVersion, context: &'static str) -> Result<()> {
    let observed = fs::symlink_metadata(path).map_err(|error| {
        if error.kind() == std::io::ErrorKind::NotFound {
            changed(context)
        } else {
            Error::Io(error)
        }
    })?;
    if observed.file_type().is_symlink() || FileVersion::from_metadata(&observed) != *expected {
        return Err(changed(context));
    }
    Ok(())
}

fn ensure_file_without_peer(file: &Path, peer: &Path, context: &'static str) -> Result<()> {
    if !matches!(
        inspect_node(file, "directory bundle Index.zip")?,
        Node::File(_)
    ) || !matches!(inspect_node(peer, "directory bundle Index")?, Node::Missing)
    {
        return Err(changed(context));
    }
    Ok(())
}

fn ensure_directory_without_peer(
    directory: &Path,
    peer: &Path,
    context: &'static str,
) -> Result<()> {
    if !matches!(
        inspect_node(directory, "directory bundle Index")?,
        Node::Directory(_)
    ) || !matches!(
        inspect_node(peer, "directory bundle Index.zip")?,
        Node::Missing
    ) {
        return Err(changed(context));
    }
    Ok(())
}

const fn changed(context: &'static str) -> Error {
    Error::DirectoryChanged { context }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct FileVersion {
    len: u64,
    modified: Option<SystemTime>,
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

impl FileVersion {
    fn from_metadata(metadata: &Metadata) -> Self {
        #[cfg(unix)]
        use std::os::unix::fs::MetadataExt;

        Self {
            len: metadata.len(),
            modified: metadata.modified().ok(),
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
}

#[cfg(test)]
mod tests {
    use std::{
        fs,
        sync::atomic::{AtomicU64, Ordering},
    };

    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

    fn iwa(identifier: u64, message_type: u32) -> Result<Vec<u8>> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: message_type,
                    data: vec![1, 2, 3],
                }],
            )?],
        };
        Ok(SnappyStream::compress(&archive.to_bytes()?)?)
    }

    fn index_zip(entries: &[(&str, Vec<u8>)]) -> Result<Vec<u8>> {
        let mut writer = StreamingArchiveWriter::new();
        for (name, bytes) in entries {
            writer.write_stored(name, bytes)?;
        }
        Ok(writer.finish_to_bytes()?)
    }

    #[test]
    fn freezes_direct_index_zip_without_claiming_whole_package_provenance() -> Result<()> {
        let temp = Temp::new()?;
        let bytes = index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?;
        fs::write(temp.path().join("Index.zip"), &bytes)?;

        let frozen = FrozenDirectoryBundle::open(temp.path())?;

        assert_eq!(frozen.provenance(), DirectoryProvenance::IndexZip);
        assert_eq!(frozen.index_zip_bytes(), Some(bytes.as_slice()));
        assert!(frozen.loose_entries().is_empty());
        assert_eq!(frozen.components().len(), 1);
        assert_eq!(
            frozen
                .components()
                .get("Index/Document.iwa")
                .map(crate::Component::name),
            Some("Index/Document.iwa")
        );
        let state_lifetime = Arc::downgrade(&frozen.state);
        let components = frozen.into_components();
        assert!(state_lifetime.upgrade().is_none());
        assert_eq!(components.len(), 1);
        Ok(())
    }

    #[test]
    fn freezes_loose_index_in_deterministic_name_order() -> Result<()> {
        let temp = Temp::new()?;
        let index = temp.path().join("Index");
        fs::create_dir(&index)?;
        fs::write(index.join("Zulu.iwa"), iwa(2, 6_001)?)?;
        fs::write(index.join("Alpha.iwa"), iwa(1, 6_000)?)?;
        fs::write(index.join("opaque.bin"), b"opaque")?;

        let frozen = FrozenDirectoryBundle::open(temp.path())?;

        assert_eq!(frozen.provenance(), DirectoryProvenance::LooseIndex);
        assert!(frozen.index_zip_bytes().is_none());
        assert_eq!(
            frozen
                .loose_entries()
                .iter()
                .map(FrozenDirectoryEntry::name)
                .collect::<Vec<_>>(),
            ["Index/Alpha.iwa", "Index/Zulu.iwa", "Index/opaque.bin"]
        );
        assert_eq!(
            frozen
                .components()
                .iter()
                .map(crate::Component::name)
                .collect::<Vec<_>>(),
            ["Index/Alpha.iwa", "Index/Zulu.iwa"]
        );
        Ok(())
    }

    #[test]
    fn freezes_legacy_marker_evidence_without_classifying_it() -> Result<()> {
        let temp = Temp::new()?;
        fs::create_dir(temp.path().join("Index"))?;
        fs::write(temp.path().join("Index/Document.iwa"), iwa(1, 10_000)?)?;
        fs::write(temp.path().join("index.xml"), [])?;
        fs::write(temp.path().join("index.apxl"), [])?;

        let frozen = FrozenDirectoryBundle::open(temp.path())?;

        assert!(frozen.markers().pages());
        assert!(frozen.markers().keynote());
        assert!(!frozen.markers().numbers());
        Ok(())
    }

    #[test]
    fn rejects_ambiguous_and_nested_index_representations() -> Result<()> {
        let temp = Temp::new()?;
        fs::create_dir(temp.path().join("Index"))?;
        fs::write(temp.path().join("Index/Document.iwa"), iwa(1, 10_000)?)?;
        fs::write(
            temp.path().join("Index.zip"),
            index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open(temp.path()),
            Err(Error::InvalidBundle(message)) if message.contains("both Index.zip and Index/")
        ));

        let nested = Temp::new()?;
        let inner = index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?;
        fs::write(
            nested.path().join("Index.zip"),
            index_zip(&[("legacy.pages/Index.zip", inner)])?,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open(nested.path()),
            Err(Error::InvalidBundle(message)) if message.contains("contains nested index")
        ));
        Ok(())
    }

    #[test]
    fn loose_index_charges_entry_and_aggregate_limits_before_decode() -> Result<()> {
        let temp = Temp::new()?;
        let index = temp.path().join("Index");
        fs::create_dir(&index)?;
        let first = iwa(1, 6_000)?;
        let second = iwa(2, 6_001)?;
        fs::write(index.join("Alpha.iwa"), &first)?;
        fs::write(index.join("Bravo.iwa"), &second)?;
        let total = u64::try_from(first.len() + second.len()).map_err(|_error| {
            Error::InvalidBundle("test aggregate length does not fit u64".to_owned())
        })?;
        let maximum_entry = u64::try_from(first.len().max(second.len())).map_err(|_error| {
            Error::InvalidBundle("test member length does not fit u64".to_owned())
        })?;

        let count_limited = Limits::new(
            Limits::MAX_INPUT_BYTES,
            1,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_limits(temp.path(), count_limited),
            Err(Error::Limit {
                kind: LimitKind::Entries,
                observed: 2,
                maximum: 1
            })
        ));

        let total_limited = Limits::new(
            Limits::MAX_INPUT_BYTES,
            2,
            maximum_entry,
            total - 1,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_limits(temp.path(), total_limited),
            Err(Error::Limit { kind: LimitKind::TotalBytes, observed, maximum })
                if observed == total && maximum == total - 1
        ));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn rejects_root_and_member_symbolic_links() -> Result<()> {
        use std::os::unix::fs::symlink;

        let target = Temp::new()?;
        fs::create_dir(target.path().join("Index"))?;
        fs::write(target.path().join("Index/Document.iwa"), iwa(1, 10_000)?)?;
        let parent = Temp::new()?;
        let root_link = parent.path().join("linked.pages");
        symlink(target.path(), &root_link)?;
        assert!(matches!(
            FrozenDirectoryBundle::open(&root_link),
            Err(Error::InvalidBundle(message)) if message.contains("symbolic link")
        ));

        let member = Temp::new()?;
        fs::create_dir(member.path().join("Index"))?;
        let outside = member.path().join("outside.iwa");
        fs::write(&outside, iwa(1, 10_000)?)?;
        symlink(&outside, member.path().join("Index/Document.iwa"))?;
        assert!(matches!(
            FrozenDirectoryBundle::open(member.path()),
            Err(Error::InvalidBundle(message)) if message.contains("symbolic link")
        ));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn rejects_special_nodes_in_loose_index() -> Result<()> {
        use std::os::unix::net::UnixListener;

        let temp = Temp::new()?;
        fs::create_dir(temp.path().join("Index"))?;
        fs::write(temp.path().join("Index/Document.iwa"), iwa(1, 10_000)?)?;
        let _socket = UnixListener::bind(temp.path().join("Index/socket.iwa"))?;

        assert!(matches!(
            FrozenDirectoryBundle::open(temp.path()),
            Err(Error::InvalidBundle(message)) if message.contains("not a regular file or directory")
        ));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn detects_replaced_file_against_captured_version() -> Result<()> {
        let temp = Temp::new()?;
        let path = temp.path().join("Document.iwa");
        fs::write(&path, b"first")?;
        let version = require_file(&path, "test member")?;
        let replacement = temp.path().join("replacement.iwa");
        fs::write(&replacement, b"other")?;
        fs::rename(&replacement, &path)?;

        assert!(matches!(
            read_stable_file(&path, &version, Limits::default(), FileRole::LooseEntry),
            Err(Error::DirectoryChanged { .. })
        ));
        Ok(())
    }

    struct Temp(PathBuf);

    impl Temp {
        fn new() -> std::io::Result<Self> {
            static NEXT: AtomicU64 = AtomicU64::new(0);
            let path = std::env::temp_dir().join(format!(
                "litchi-iwa-archive-directory-{}-{}",
                std::process::id(),
                NEXT.fetch_add(1, Ordering::Relaxed)
            ));
            fs::create_dir(&path)?;
            Ok(Self(path))
        }

        fn path(&self) -> &Path {
            &self.0
        }
    }

    impl Drop for Temp {
        fn drop(&mut self) {
            if let Err(_error) = fs::remove_dir_all(&self.0) {}
        }
    }
}
