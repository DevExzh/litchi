//! Immutable ingress for legacy directory-backed iWork bundles.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "filesystem acquisition helpers stay beside the state they validate"
)]

use std::{fmt, fs, io::Read, path::Path, sync::Arc};

#[cfg(all(not(unix), not(windows)))]
use std::{
    ffi::OsString,
    fs::{Metadata, OpenOptions},
    path::PathBuf,
    time::SystemTime,
};

#[cfg(unix)]
use std::ffi::{CStr, CString};

#[cfg(unix)]
use rustix::{
    fd::OwnedFd,
    fs::{self as unix_fs, AtFlags, Dir, FileType, Mode, OFlags, Stat},
};

use crate::catalog::DirectoryIndexReport;
use crate::{ComponentCatalog, Error, LimitKind, Limits, Result};

const READ_CHUNK_BYTES: usize = 16 * 1024;
const PROPERTIES_LOGICAL_NAME: &str = "Metadata/Properties.plist";
const BUILD_VERSION_HISTORY_LOGICAL_NAME: &str = "Metadata/BuildVersionHistory.plist";
const DOCUMENT_IDENTIFIER_LOGICAL_NAME: &str = "Metadata/DocumentIdentifier";
/// Maximum size of each canonical metadata diagnostic retained during a
/// bounded semantic directory capture.
pub const MAX_DIRECTORY_PROPERTIES_BYTES: u64 = 64 * 1024;

/// Canonical Pages metadata sidecars frozen beside a directory bundle's IWA
/// index.
///
/// The values are optional because Pages packages may omit any of these
/// diagnostics.  This value never represents arbitrary `Metadata/` members.
#[derive(Debug, Clone, Default)]
pub struct DirectoryMetadataSidecars {
    properties: Option<Arc<[u8]>>,
    build_version_history: Option<Arc<[u8]>>,
    document_identifier: Option<Arc<[u8]>>,
}

impl DirectoryMetadataSidecars {
    /// Borrow canonical `Metadata/Properties.plist`, if it was present.
    #[must_use]
    pub fn properties_plist(&self) -> Option<&[u8]> {
        self.properties.as_deref()
    }

    /// Borrow canonical `Metadata/BuildVersionHistory.plist`, if present.
    #[must_use]
    pub fn build_version_history_plist(&self) -> Option<&[u8]> {
        self.build_version_history.as_deref()
    }

    /// Borrow canonical `Metadata/DocumentIdentifier`, if present.
    #[must_use]
    pub fn document_identifier(&self) -> Option<&[u8]> {
        self.document_identifier.as_deref()
    }

    /// Consume this fixed DTO into its canonical immutable byte owners.
    #[doc(hidden)]
    #[must_use]
    pub fn __into_parts(self) -> (Option<Arc<[u8]>>, Option<Arc<[u8]>>, Option<Arc<[u8]>>) {
        (
            self.properties,
            self.build_version_history,
            self.document_identifier,
        )
    }
}

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
/// component traversal and is cheaply cloneable through shared state. The
/// ordinary constructors remain index-only. A format-owned semantic reader
/// may opt into capturing canonical `Metadata/Properties.plist` with the same
/// pinned root; every other `Metadata/` member and all `Data/` members remain
/// outside this adapter.
#[derive(Clone)]
pub struct FrozenDirectoryBundle {
    state: Arc<State>,
}

struct State {
    provenance: DirectoryProvenance,
    markers: DirectoryMarkers,
    limits: Limits,
    components: Arc<ComponentCatalog>,
    sidecars: DirectoryMetadataSidecars,
    index_report: DirectoryIndexReport,
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
            .field(
                "properties_bytes",
                &self.state.sidecars.properties.as_deref().map(<[u8]>::len),
            )
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
    /// On Unix, pre-existing symbolic links in ancestor directories are
    /// resolved by the initial root open; the resolved root itself is opened
    /// with `O_NOFOLLOW` and pinned, and every later lookup is relative to that
    /// descriptor. Replacing an ancestor, root pathname, or `Index` pathname
    /// therefore cannot redirect an in-progress capture.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::open`].
    pub fn open_with_limits(path: impl AsRef<Path>, limits: Limits) -> Result<Self> {
        Self::open_with_profile(path.as_ref(), limits, CaptureProfile::None)
    }

    /// Snapshot a directory index plus its canonical properties diagnostic.
    ///
    /// This semantic-only profile is intentionally opt-in so ordinary
    /// directory detection and Pages/Numbers projection retain their existing
    /// index-only resource contract.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::open_with_limits`] and charges the
    /// optional diagnostic as one logical entry before allocating its bytes.
    #[doc(hidden)]
    pub fn open_with_properties(path: impl AsRef<Path>, limits: Limits) -> Result<Self> {
        Self::open_with_profile(path.as_ref(), limits, CaptureProfile::Properties)
    }

    /// Snapshot a directory index plus Pages' three canonical metadata
    /// diagnostics.
    ///
    /// The profile captures only `Metadata/Properties.plist`,
    /// `Metadata/BuildVersionHistory.plist`, and `Metadata/DocumentIdentifier`
    /// from one pinned `Metadata` directory. All selected files are charged
    /// before their payload bytes are allocated.
    #[doc(hidden)]
    pub fn open_with_pages_metadata(path: impl AsRef<Path>, limits: Limits) -> Result<Self> {
        Self::open_with_profile(path.as_ref(), limits, CaptureProfile::PagesMetadata)
    }

    /// Snapshot Pages metadata only when the already-frozen index and markers
    /// satisfy a format-owned predicate.
    ///
    /// The predicate runs after index acquisition from the pinned root but
    /// before any selected metadata authority is inspected or read. This
    /// keeps Pages-only sidecar policy from affecting Keynote, Numbers, or
    /// unrecognized directory bundles without reopening an ambient path.
    #[doc(hidden)]
    pub fn open_with_pages_metadata_when<F>(
        path: impl AsRef<Path>,
        limits: Limits,
        mut predicate: F,
    ) -> Result<Self>
    where
        F: FnMut(&ComponentCatalog, DirectoryMarkers) -> bool,
    {
        Self::open_with_profile_when(
            path.as_ref(),
            limits,
            CaptureProfile::PagesMetadata,
            Some(&mut predicate),
        )
    }

    fn open_with_profile(root: &Path, limits: Limits, profile: CaptureProfile) -> Result<Self> {
        Self::open_with_profile_when(root, limits, profile, None)
    }

    fn open_with_profile_when(
        root: &Path,
        limits: Limits,
        profile: CaptureProfile,
        predicate: Option<&mut dyn FnMut(&ComponentCatalog, DirectoryMarkers) -> bool>,
    ) -> Result<Self> {
        let checked_limits = limits.validate()?;

        #[cfg(unix)]
        {
            Self::open_unix(root, checked_limits, profile, predicate)
        }

        // Directory ingress relies on a pinned root descriptor on Unix.  The
        // legacy path-and-metadata snapshot is not safe enough on Windows,
        // where it cannot provide that invariant, so reject rather than
        // publishing an ambiguous capture.
        #[cfg(windows)]
        {
            let _ = (root, checked_limits, profile, predicate);
            Err(Error::InvalidBundle(
                "directory-backed bundle ingress is unsupported on Windows".to_owned(),
            ))
        }

        #[cfg(all(not(unix), not(windows)))]
        {
            let root_before = require_directory(root, "directory bundle root")?;
            reject_path_encryption_markers(root)?;
            let markers_before = snapshot_markers(root)?;
            let delayed = predicate.is_some();
            let mut predicate = predicate;
            let mut sidecars = if profile.captures_any() && !delayed {
                inspect_sidecars(root, profile)?
            } else {
                MetadataCapture::default()
            };
            let index_limits = limits_without_sidecars(checked_limits, &sidecars)?;
            let index_zip_path = root.join("Index.zip");
            let index_path = root.join("Index");
            let index_zip = inspect_node(&index_zip_path, "directory bundle Index.zip")?;
            let index = inspect_node(&index_path, "directory bundle Index")?;

            let snapshot = match (index_zip, index) {
                (Node::File(version), Node::Missing) => Self::from_index_zip(
                    &index_zip_path,
                    version,
                    markers_before.markers,
                    index_limits,
                    checked_limits,
                )
                .map_err(|error| remap_reserved_limit(error, checked_limits, &sidecars))?,
                (Node::Missing, Node::Directory(version)) => Self::from_loose_index(
                    &index_path,
                    version,
                    markers_before.markers,
                    index_limits,
                    checked_limits,
                )
                .map_err(|error| remap_reserved_limit(error, checked_limits, &sidecars))?,
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
            let capture_sidecars = match predicate.as_mut() {
                Some(predicate) => predicate(snapshot.components(), markers_before.markers),
                None => profile.captures_any(),
            };
            if delayed && capture_sidecars {
                sidecars = inspect_sidecars(root, profile)?;
            }
            validate_sidecars_budget(snapshot.state.index_report, checked_limits, &sidecars)?;
            if capture_sidecars {
                read_sidecars(root, &mut sidecars, checked_limits)?;
                verify_sidecars(root, &sidecars)?;
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
            reject_path_encryption_markers(root)?;
            snapshot.with_sidecars(sidecars.into_sidecars())
        }
    }

    #[cfg(unix)]
    fn open_unix(
        root: &Path,
        limits: Limits,
        profile: CaptureProfile,
        predicate: Option<&mut dyn FnMut(&ComponentCatalog, DirectoryMarkers) -> bool>,
    ) -> Result<Self> {
        let root_fd = unix_fs::open(
            root,
            OFlags::RDONLY
                | OFlags::DIRECTORY
                | OFlags::NOFOLLOW
                | OFlags::NONBLOCK
                | OFlags::CLOEXEC,
            Mode::empty(),
        )
        .map_err(|error| {
            let root_is_symlink =
                matches!(error, rustix::io::Errno::LOOP | rustix::io::Errno::NOTDIR)
                    && fs::symlink_metadata(root)
                        .is_ok_and(|metadata| metadata.file_type().is_symlink());
            if root_is_symlink {
                Error::InvalidBundle("directory bundle root must not be a symbolic link".to_owned())
            } else {
                unix_error(error)
            }
        })?;
        Self::open_unix_root(&root_fd, limits, profile, predicate)
    }

    #[cfg(unix)]
    fn open_unix_root(
        root_fd: &OwnedFd,
        limits: Limits,
        profile: CaptureProfile,
        mut predicate: Option<&mut dyn FnMut(&ComponentCatalog, DirectoryMarkers) -> bool>,
    ) -> Result<Self> {
        let root_before = unix_file_version(&unix_fs::fstat(root_fd).map_err(unix_error)?)?;
        unix_require_kind(root_fd, FileType::Directory, "directory bundle root")?;
        unix_reject_encryption_markers(root_fd)?;
        let markers_before = unix_snapshot_markers(root_fd)?;
        let delayed = predicate.is_some();
        let mut sidecars = if profile.captures_any() && !delayed {
            unix_inspect_sidecars(root_fd, profile)?
        } else {
            MetadataCapture::default()
        };
        let index_limits = limits_without_sidecars(limits, &sidecars)?;
        let index_zip = unix_inspect_node(root_fd, c"Index.zip", "directory bundle Index.zip")?;
        let index = unix_inspect_node(root_fd, c"Index", "directory bundle Index")?;

        let snapshot = match (index_zip, index) {
            (Node::File(version), Node::Missing) => {
                let bytes = unix_read_stable_file_at(
                    root_fd,
                    c"Index.zip",
                    &version,
                    index_limits,
                    FileRole::IndexZip,
                )
                .map_err(|error| remap_reserved_limit(error, limits, &sidecars))?;
                let snapshot = Self::from_captured_index_zip(
                    bytes,
                    markers_before.markers,
                    index_limits,
                    limits,
                )
                .map_err(|error| remap_reserved_limit(error, limits, &sidecars))?;
                unix_ensure_node_version(
                    root_fd,
                    c"Index.zip",
                    &version,
                    "verifying directory bundle Index.zip after component parsing",
                )?;
                if !matches!(
                    unix_inspect_node(root_fd, c"Index", "directory bundle Index")?,
                    Node::Missing
                ) {
                    return Err(changed(
                        "verifying directory bundle Index.zip representation",
                    ));
                }
                snapshot
            },
            (Node::Missing, Node::Directory(version)) => {
                let index_fd = unix_open_directory_at(root_fd, c"Index", &version)?;
                let manifest = unix_scan_manifest(&index_fd, index_limits)
                    .map_err(|error| remap_reserved_limit(error, limits, &sidecars))?;
                if manifest.is_empty() {
                    return Err(Error::InvalidBundle(
                        "directory bundle Index/ contains no entries".to_owned(),
                    ));
                }
                let entries = unix_read_manifest(&index_fd, &manifest, index_limits)
                    .map_err(|error| remap_reserved_limit(error, limits, &sidecars))?;
                let snapshot = Self::from_captured_loose(entries, markers_before.markers, limits)
                    .map_err(|error| remap_reserved_limit(error, limits, &sidecars))?;

                let observed = unix_scan_manifest(&index_fd, index_limits)
                    .map_err(|error| remap_reserved_limit(error, limits, &sidecars))?;
                if observed != manifest {
                    return Err(changed(
                        "verifying directory bundle Index/ manifest after component parsing",
                    ));
                }
                let index_after =
                    unix_file_version(&unix_fs::fstat(&index_fd).map_err(unix_error)?)?;
                if index_after != version {
                    return Err(changed(
                        "verifying directory bundle Index/ after component parsing",
                    ));
                }
                unix_ensure_node_version(
                    root_fd,
                    c"Index",
                    &version,
                    "verifying pinned directory bundle Index/ representation",
                )?;
                if !matches!(
                    unix_inspect_node(root_fd, c"Index.zip", "directory bundle Index.zip")?,
                    Node::Missing
                ) {
                    return Err(changed("verifying directory bundle Index/ representation"));
                }
                snapshot
            },
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

        unix_reject_encryption_markers(root_fd)?;
        if unix_snapshot_markers(root_fd)? != markers_before {
            return Err(changed("verifying directory bundle application markers"));
        }
        let capture_sidecars = match predicate.as_mut() {
            Some(predicate) => predicate(snapshot.components(), markers_before.markers),
            None => profile.captures_any(),
        };
        if delayed && capture_sidecars {
            sidecars = unix_inspect_sidecars(root_fd, profile)?;
        }
        validate_sidecars_budget(snapshot.state.index_report, limits, &sidecars)?;
        if capture_sidecars {
            unix_read_sidecars(root_fd, &mut sidecars, limits)?;
            unix_verify_sidecars(root_fd, &sidecars)?;
        }
        let root_after = unix_file_version(&unix_fs::fstat(root_fd).map_err(unix_error)?)?;
        if root_after != root_before {
            return Err(changed("verifying pinned directory bundle root"));
        }
        snapshot.with_sidecars(sidecars.into_sidecars())
    }

    #[cfg(all(not(unix), not(windows)))]
    fn from_index_zip(
        path: &Path,
        expected: FileVersion,
        markers: DirectoryMarkers,
        index_limits: Limits,
        retained_limits: Limits,
    ) -> Result<Self> {
        let bytes = read_stable_file(path, &expected, index_limits, FileRole::IndexZip)?;
        let snapshot =
            Self::from_captured_index_zip(bytes, markers, index_limits, retained_limits)?;
        ensure_path_version(
            path,
            &expected,
            "verifying directory bundle Index.zip after component parsing",
        )?;
        Ok(snapshot)
    }

    fn from_captured_index_zip(
        bytes: Box<[u8]>,
        markers: DirectoryMarkers,
        index_limits: Limits,
        retained_limits: Limits,
    ) -> Result<Self> {
        let (components, index_report) = ComponentCatalog::from_directory_index_zip_with_report(
            &bytes,
            index_limits,
            retained_limits,
        )?;
        if components.is_empty() {
            return Err(Error::InvalidBundle(
                "directory bundle Index.zip contains no decodable IWA components".to_owned(),
            ));
        }
        Ok(Self {
            state: Arc::new(State {
                provenance: DirectoryProvenance::IndexZip,
                markers,
                limits: retained_limits,
                components: Arc::new(components),
                sidecars: DirectoryMetadataSidecars::default(),
                index_report,
                storage: Storage::IndexZip(bytes),
            }),
        })
    }

    #[cfg(all(not(unix), not(windows)))]
    fn from_loose_index(
        path: &Path,
        expected: FileVersion,
        markers: DirectoryMarkers,
        index_limits: Limits,
        retained_limits: Limits,
    ) -> Result<Self> {
        let manifest = scan_manifest(path, index_limits)?;
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
            let data = read_stable_file(
                &item.path,
                &item.version,
                index_limits,
                FileRole::LooseEntry,
            )?;
            entries.push(FrozenDirectoryEntry {
                name: item.logical_name.clone(),
                data,
            });
        }

        let snapshot = Self::from_captured_loose(entries, markers, retained_limits)?;
        let observed = scan_manifest(path, index_limits)?;
        if manifest != observed {
            return Err(changed(
                "verifying directory bundle Index/ manifest after component parsing",
            ));
        }
        ensure_path_version(
            path,
            &expected,
            "verifying directory bundle Index/ after component parsing",
        )?;
        Ok(snapshot)
    }

    fn from_captured_loose(
        entries: Vec<FrozenDirectoryEntry>,
        markers: DirectoryMarkers,
        retained_limits: Limits,
    ) -> Result<Self> {
        let components = ComponentCatalog::from_logical_entries(
            entries.iter().map(|entry| (entry.name(), entry.data())),
            retained_limits,
        )?;
        if components.is_empty() {
            return Err(Error::InvalidBundle(
                "directory bundle Index/ contains no decodable IWA components".to_owned(),
            ));
        }
        let index_report = entries.iter().try_fold(
            DirectoryIndexReport {
                input_bytes: 0,
                entries: 0,
                metadata_bytes: 0,
                expanded_bytes: 0,
            },
            |mut report, entry| {
                report.entries = report.entries.checked_add(1).ok_or_else(|| {
                    Error::InvalidBundle("directory entry count overflowed usize".to_owned())
                })?;
                report.metadata_bytes = report
                    .metadata_bytes
                    .checked_add(u64::try_from(entry.name.len()).map_err(|_error| {
                        Error::InvalidBundle(
                            "directory entry name length does not fit u64".to_owned(),
                        )
                    })?)
                    .ok_or_else(|| {
                        Error::InvalidBundle("directory metadata length overflowed u64".to_owned())
                    })?;
                report.expanded_bytes = report
                    .expanded_bytes
                    .checked_add(u64::try_from(entry.data.len()).map_err(|_error| {
                        Error::InvalidBundle("directory entry length does not fit u64".to_owned())
                    })?)
                    .ok_or_else(|| {
                        Error::InvalidBundle("directory byte length overflowed u64".to_owned())
                    })?;
                report.input_bytes = report.expanded_bytes;
                Ok::<DirectoryIndexReport, Error>(report)
            },
        )?;
        Ok(Self {
            state: Arc::new(State {
                provenance: DirectoryProvenance::LooseIndex,
                markers,
                limits: retained_limits,
                components: Arc::new(components),
                sidecars: DirectoryMetadataSidecars::default(),
                index_report,
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

    /// Borrow the frozen canonical `Metadata/Properties.plist` diagnostic.
    ///
    /// The bytes are retained only for archive-free semantic metadata. Their
    /// presence does not turn this directory snapshot into an exact package
    /// artifact and does not imply that any other sidecar was captured.
    #[must_use]
    pub fn properties_plist(&self) -> Option<&[u8]> {
        self.state.sidecars.properties_plist()
    }

    /// Consume the snapshot into semantic components and the optional frozen
    /// canonical properties diagnostic.
    #[doc(hidden)]
    #[must_use]
    pub fn into_semantic_parts(self) -> (Arc<ComponentCatalog>, Option<Arc<[u8]>>) {
        match Arc::try_unwrap(self.state) {
            Ok(state) => (state.components, state.sidecars.properties),
            Err(state) => (
                Arc::clone(&state.components),
                state.sidecars.properties.as_ref().map(Arc::clone),
            ),
        }
    }

    /// Consume this snapshot into semantic components and Pages' canonical
    /// metadata sidecars. This semantic-only handoff never yields exact
    /// package provenance.
    #[doc(hidden)]
    #[must_use]
    pub fn into_pages_semantic_parts(self) -> (Arc<ComponentCatalog>, DirectoryMetadataSidecars) {
        match Arc::try_unwrap(self.state) {
            Ok(state) => (state.components, state.sidecars),
            Err(state) => (Arc::clone(&state.components), state.sidecars.clone()),
        }
    }

    /// Consume this snapshot into a shared immutable component catalog.
    ///
    /// This is the semantic-only handoff for directory-backed documents. It
    /// deliberately cannot yield a [`crate::SourceCatalog`], because neither
    /// directory representation is an exact ZIP source for the whole iWork
    /// artifact.
    #[must_use]
    pub fn into_components(self) -> Arc<ComponentCatalog> {
        self.into_semantic_parts().0
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

    fn with_sidecars(mut self, sidecars: DirectoryMetadataSidecars) -> Result<Self> {
        Arc::get_mut(&mut self.state)
            .expect("new directory snapshot has exclusive state")
            .sidecars = sidecars;
        Ok(self)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CaptureProfile {
    None,
    Properties,
    PagesMetadata,
}

impl CaptureProfile {
    const fn captures_any(self) -> bool {
        !matches!(self, Self::None)
    }

    const fn captures(self, sidecar: Sidecar) -> bool {
        match self {
            Self::None => false,
            Self::Properties => matches!(sidecar, Sidecar::Properties),
            Self::PagesMetadata => true,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Sidecar {
    Properties,
    BuildVersionHistory,
    DocumentIdentifier,
}

impl Sidecar {
    const ALL: [Self; 3] = [
        Self::Properties,
        Self::BuildVersionHistory,
        Self::DocumentIdentifier,
    ];

    const fn logical_name(self) -> &'static str {
        match self {
            Self::Properties => PROPERTIES_LOGICAL_NAME,
            Self::BuildVersionHistory => BUILD_VERSION_HISTORY_LOGICAL_NAME,
            Self::DocumentIdentifier => DOCUMENT_IDENTIFIER_LOGICAL_NAME,
        }
    }

    #[cfg(all(not(unix), not(windows)))]
    const fn basename(self) -> &'static str {
        match self {
            Self::Properties => "Properties.plist",
            Self::BuildVersionHistory => "BuildVersionHistory.plist",
            Self::DocumentIdentifier => "DocumentIdentifier",
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Properties => "directory bundle Metadata/Properties.plist",
            Self::BuildVersionHistory => "directory bundle Metadata/BuildVersionHistory.plist",
            Self::DocumentIdentifier => "directory bundle Metadata/DocumentIdentifier",
        }
    }

    const fn verification_context(self) -> &'static str {
        match self {
            Self::Properties => "verifying directory bundle Metadata/Properties.plist",
            Self::BuildVersionHistory => {
                "verifying directory bundle Metadata/BuildVersionHistory.plist"
            },
            Self::DocumentIdentifier => "verifying directory bundle Metadata/DocumentIdentifier",
        }
    }

    #[cfg(unix)]
    const fn c_name(self) -> &'static CStr {
        match self {
            Self::Properties => c"Properties.plist",
            Self::BuildVersionHistory => c"BuildVersionHistory.plist",
            Self::DocumentIdentifier => c"DocumentIdentifier",
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct SidecarCapture {
    selected: bool,
    file: Option<FileVersion>,
    data: Option<Box<[u8]>>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct MetadataCapture {
    directory: Option<FileVersion>,
    properties: SidecarCapture,
    build_version_history: SidecarCapture,
    document_identifier: SidecarCapture,
}

impl MetadataCapture {
    fn sidecar(&self, sidecar: Sidecar) -> &SidecarCapture {
        match sidecar {
            Sidecar::Properties => &self.properties,
            Sidecar::BuildVersionHistory => &self.build_version_history,
            Sidecar::DocumentIdentifier => &self.document_identifier,
        }
    }

    fn sidecar_mut(&mut self, sidecar: Sidecar) -> &mut SidecarCapture {
        match sidecar {
            Sidecar::Properties => &mut self.properties,
            Sidecar::BuildVersionHistory => &mut self.build_version_history,
            Sidecar::DocumentIdentifier => &mut self.document_identifier,
        }
    }

    fn into_sidecars(self) -> DirectoryMetadataSidecars {
        DirectoryMetadataSidecars {
            properties: self.properties.data.map(Arc::from),
            build_version_history: self.build_version_history.data.map(Arc::from),
            document_identifier: self.document_identifier.data.map(Arc::from),
        }
    }
}

fn sidecar_totals(capture: &MetadataCapture) -> Result<(u64, usize)> {
    Sidecar::ALL
        .iter()
        .try_fold((0_u64, 0_usize), |(bytes, entries), sidecar| {
            if !capture.sidecar(*sidecar).selected {
                return Ok((bytes, entries));
            }
            let Some(file) = capture.sidecar(*sidecar).file else {
                return Ok((bytes, entries));
            };
            if file.len > MAX_DIRECTORY_PROPERTIES_BYTES {
                return Err(Error::Limit {
                    kind: LimitKind::EntryBytes,
                    observed: file.len,
                    maximum: MAX_DIRECTORY_PROPERTIES_BYTES,
                });
            }
            Ok((
                bytes.checked_add(file.len).ok_or_else(|| {
                    Error::InvalidBundle("directory sidecar length overflowed u64".to_owned())
                })?,
                entries.checked_add(1).ok_or_else(|| {
                    Error::InvalidBundle("directory sidecar count overflowed usize".to_owned())
                })?,
            ))
        })
}

fn limits_without_sidecars(limits: Limits, capture: &MetadataCapture) -> Result<Limits> {
    let (sidecar_bytes, sidecar_entries) = sidecar_totals(capture)?;
    if sidecar_entries == 0 {
        return Ok(limits);
    }
    for sidecar in Sidecar::ALL {
        let Some(file) = capture.sidecar(sidecar).file else {
            continue;
        };
        check_file_size(file.len, limits, FileRole::Sidecar)?;
    }
    if sidecar_bytes >= limits.max_input_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: sidecar_bytes
                .saturating_add(u64::from(sidecar_bytes == limits.max_input_bytes())),
            maximum: limits.max_input_bytes(),
        });
    }
    if sidecar_entries >= limits.max_entries() {
        return Err(Error::Limit {
            kind: LimitKind::Entries,
            observed: u64::try_from(sidecar_entries.saturating_add(1)).unwrap_or(u64::MAX),
            maximum: u64::try_from(limits.max_entries()).unwrap_or(u64::MAX),
        });
    }
    if sidecar_bytes >= limits.max_total_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::TotalBytes,
            observed: sidecar_bytes
                .saturating_add(u64::from(sidecar_bytes == limits.max_total_bytes())),
            maximum: limits.max_total_bytes(),
        });
    }
    let mut name_bytes = 0_u64;
    for sidecar in Sidecar::ALL {
        if capture.sidecar(sidecar).selected && capture.sidecar(sidecar).file.is_some() {
            let name = u64::try_from(sidecar.logical_name().len()).unwrap_or(u64::MAX);
            check_member_name(name, limits)?;
            name_bytes = name_bytes.checked_add(name).ok_or_else(|| {
                Error::InvalidBundle("directory sidecar metadata length overflowed u64".to_owned())
            })?;
        }
    }
    check_metadata_bytes(name_bytes, limits)?;
    let remaining_input = limits
        .max_input_bytes()
        .checked_sub(sidecar_bytes)
        .ok_or_else(|| Error::Limit {
            kind: LimitKind::InputBytes,
            observed: sidecar_bytes,
            maximum: limits.max_input_bytes(),
        })?;
    let remaining_metadata =
        limits
            .max_metadata_bytes()
            .checked_sub(name_bytes)
            .ok_or(Error::Limit {
                kind: LimitKind::MetadataBytes,
                observed: name_bytes,
                maximum: limits.max_metadata_bytes(),
            })?;
    Limits::new(
        remaining_input,
        limits
            .max_entries()
            .checked_sub(sidecar_entries)
            .ok_or_else(|| Error::Limit {
                kind: LimitKind::Entries,
                observed: u64::try_from(sidecar_entries).unwrap_or(u64::MAX),
                maximum: u64::try_from(limits.max_entries()).unwrap_or(u64::MAX),
            })?,
        limits.max_entry_bytes(),
        limits
            .max_total_bytes()
            .checked_sub(sidecar_bytes)
            .ok_or_else(|| Error::Limit {
                kind: LimitKind::TotalBytes,
                observed: sidecar_bytes,
                maximum: limits.max_total_bytes(),
            })?,
        limits.max_iwa_stream_bytes(),
    )?
    .with_archive_limits(limits.archive_limits())?
    .with_derived_metadata_bytes(remaining_metadata.min(remaining_input))
}

fn remap_reserved_limit(error: Error, limits: Limits, capture: &MetadataCapture) -> Error {
    let Ok((sidecar_bytes, sidecar_entries)) = sidecar_totals(capture) else {
        return error;
    };
    match error {
        Error::Limit {
            kind: LimitKind::InputBytes,
            observed,
            ..
        } => Error::Limit {
            kind: LimitKind::InputBytes,
            observed: observed.saturating_add(sidecar_bytes),
            maximum: limits.max_input_bytes(),
        },
        Error::Limit {
            kind: LimitKind::Entries,
            observed,
            ..
        } => Error::Limit {
            kind: LimitKind::Entries,
            observed: observed.saturating_add(u64::try_from(sidecar_entries).unwrap_or(u64::MAX)),
            maximum: limits.max_entries() as u64,
        },
        Error::Limit {
            kind: LimitKind::TotalBytes,
            observed,
            ..
        } => Error::Limit {
            kind: LimitKind::TotalBytes,
            observed: observed.saturating_add(sidecar_bytes),
            maximum: limits.max_total_bytes(),
        },
        Error::Limit {
            kind: LimitKind::MetadataBytes,
            observed,
            ..
        } => {
            let reserved = Sidecar::ALL.iter().fold(0_u64, |total, sidecar| {
                if capture.sidecar(*sidecar).selected && capture.sidecar(*sidecar).file.is_some() {
                    total.saturating_add(
                        u64::try_from(sidecar.logical_name().len()).unwrap_or(u64::MAX),
                    )
                } else {
                    total
                }
            });
            Error::Limit {
                kind: LimitKind::MetadataBytes,
                observed: observed.saturating_add(reserved),
                maximum: limits.max_metadata_bytes(),
            }
        },
        other => other,
    }
}

fn validate_sidecars_budget(
    index: DirectoryIndexReport,
    limits: Limits,
    capture: &MetadataCapture,
) -> Result<()> {
    let (sidecar_bytes, sidecar_entries) = sidecar_totals(capture)?;
    if sidecar_entries == 0 {
        return Ok(());
    }
    let input_bytes = index
        .input_bytes
        .checked_add(sidecar_bytes)
        .ok_or_else(|| Error::InvalidBundle("directory input length overflowed u64".to_owned()))?;
    check_aggregate_input_bytes(input_bytes, limits)?;
    for sidecar in Sidecar::ALL {
        if !capture.sidecar(sidecar).selected {
            continue;
        }
        if let Some(file) = capture.sidecar(sidecar).file {
            if file.len > limits.max_entry_bytes() {
                return Err(Error::Limit {
                    kind: LimitKind::EntryBytes,
                    observed: file.len,
                    maximum: limits.max_entry_bytes(),
                });
            }
        }
    }
    let entries = index
        .entries
        .checked_add(sidecar_entries)
        .ok_or_else(|| Error::InvalidBundle("directory entry count overflowed usize".to_owned()))?;
    check_entry_count(entries, limits)?;
    let name_bytes = Sidecar::ALL.iter().try_fold(0_u64, |total, sidecar| {
        if !capture.sidecar(*sidecar).selected || capture.sidecar(*sidecar).file.is_none() {
            return Ok(total);
        }
        let bytes = u64::try_from(sidecar.logical_name().len()).unwrap_or(u64::MAX);
        check_member_name(bytes, limits)?;
        total.checked_add(bytes).ok_or_else(|| {
            Error::InvalidBundle("directory sidecar metadata length overflowed u64".to_owned())
        })
    })?;
    let metadata_bytes = index
        .metadata_bytes
        .checked_add(name_bytes)
        .ok_or_else(|| {
            Error::InvalidBundle("directory metadata length overflowed u64".to_owned())
        })?;
    check_metadata_bytes(metadata_bytes, limits)?;
    let expanded = index
        .expanded_bytes
        .checked_add(sidecar_bytes)
        .ok_or_else(|| {
            Error::InvalidBundle("directory expanded length overflowed u64".to_owned())
        })?;
    check_total_bytes(expanded, limits)?;
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MarkerSnapshot {
    markers: DirectoryMarkers,
    pages: Option<FileVersion>,
    keynote: Option<FileVersion>,
    numbers: Option<FileVersion>,
}

#[cfg(unix)]
fn unix_snapshot_markers(root_fd: &OwnedFd) -> Result<MarkerSnapshot> {
    let pages = unix_inspect_marker(root_fd, c"index.xml", "legacy Pages marker")?;
    let keynote = unix_inspect_marker(root_fd, c"index.apxl", "legacy Keynote marker")?;
    let numbers = unix_inspect_marker(root_fd, c"index.numbers", "legacy Numbers marker")?;
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

#[cfg(unix)]
fn unix_inspect_marker(root_fd: &OwnedFd, name: &CStr, label: &str) -> Result<Option<FileVersion>> {
    match unix_inspect_node(root_fd, name, label)? {
        Node::Missing => Ok(None),
        Node::File(version) => Ok(Some(version)),
        Node::Directory(_) => Err(Error::InvalidBundle(format!(
            "{label} is not a regular file"
        ))),
    }
}

#[cfg(unix)]
fn unix_reject_encryption_markers(root_fd: &OwnedFd) -> Result<()> {
    for name in [c".iwpv2", c".iwph"] {
        if unix_node_exists(root_fd, name)? {
            return Err(Error::Encrypted);
        }
    }

    match unix_inspect_node(root_fd, c"Metadata", "directory bundle Metadata")? {
        Node::Missing => Ok(()),
        Node::Directory(version) => {
            let metadata_fd = unix_open_directory_at(root_fd, c"Metadata", &version)?;
            for name in [c".iwpv2", c".iwph"] {
                if unix_node_exists(&metadata_fd, name)? {
                    return Err(Error::Encrypted);
                }
            }
            Ok(())
        },
        Node::File(_) => Err(Error::InvalidBundle(
            "directory bundle Metadata is not a directory".to_owned(),
        )),
    }
}

#[cfg(unix)]
fn unix_inspect_sidecars(root_fd: &OwnedFd, profile: CaptureProfile) -> Result<MetadataCapture> {
    match unix_inspect_node(root_fd, c"Metadata", "directory bundle Metadata")? {
        Node::Missing => Ok(MetadataCapture::default()),
        Node::File(_) => Err(Error::InvalidBundle(
            "directory bundle Metadata is not a directory".to_owned(),
        )),
        Node::Directory(directory) => {
            let metadata_fd = unix_open_directory_at(root_fd, c"Metadata", &directory)?;
            let mut capture = MetadataCapture {
                directory: Some(directory),
                ..MetadataCapture::default()
            };
            for sidecar in Sidecar::ALL {
                if !profile.captures(sidecar) {
                    continue;
                }
                capture.sidecar_mut(sidecar).selected = true;
                match unix_inspect_node(&metadata_fd, sidecar.c_name(), sidecar.label())? {
                    Node::Missing => {},
                    Node::Directory(_) => {
                        return Err(Error::InvalidBundle(format!(
                            "{} is not a regular file",
                            sidecar.label()
                        )));
                    },
                    Node::File(file) => capture.sidecar_mut(sidecar).file = Some(file),
                }
            }
            Ok(capture)
        },
    }
}

#[cfg(unix)]
fn unix_read_sidecars(
    root_fd: &OwnedFd,
    capture: &mut MetadataCapture,
    limits: Limits,
) -> Result<()> {
    let Some(directory) = capture.directory else {
        return Ok(());
    };
    let metadata_fd = unix_open_directory_at(root_fd, c"Metadata", &directory)?;
    for sidecar in Sidecar::ALL {
        if !capture.sidecar(sidecar).selected {
            continue;
        }
        let Some(file) = capture.sidecar(sidecar).file else {
            continue;
        };
        let data = unix_read_stable_file_at(
            &metadata_fd,
            sidecar.c_name(),
            &file,
            limits,
            FileRole::Sidecar,
        )?;
        capture.sidecar_mut(sidecar).data = Some(data);
    }
    let observed = unix_file_version(&unix_fs::fstat(&metadata_fd).map_err(unix_error)?)?;
    if observed != directory {
        return Err(changed("reading directory bundle Metadata"));
    }
    Ok(())
}

#[cfg(unix)]
fn unix_verify_sidecars(root_fd: &OwnedFd, capture: &MetadataCapture) -> Result<()> {
    match capture.directory {
        None => {
            if !matches!(
                unix_inspect_node(root_fd, c"Metadata", "directory bundle Metadata")?,
                Node::Missing
            ) {
                return Err(changed("verifying directory bundle Metadata"));
            }
        },
        Some(directory) => {
            let metadata_fd = unix_open_directory_at(root_fd, c"Metadata", &directory)?;
            for sidecar in Sidecar::ALL {
                if !capture.sidecar(sidecar).selected {
                    continue;
                }
                match capture.sidecar(sidecar).file {
                    None => {
                        if !matches!(
                            unix_inspect_node(&metadata_fd, sidecar.c_name(), sidecar.label())?,
                            Node::Missing
                        ) {
                            return Err(changed(sidecar.verification_context()));
                        }
                    },
                    Some(file) => unix_ensure_node_version(
                        &metadata_fd,
                        sidecar.c_name(),
                        &file,
                        sidecar.verification_context(),
                    )?,
                }
            }
            let observed = unix_file_version(&unix_fs::fstat(&metadata_fd).map_err(unix_error)?)?;
            if observed != directory {
                return Err(changed("verifying directory bundle Metadata"));
            }
        },
    }
    Ok(())
}

#[cfg(unix)]
fn unix_node_exists(parent: &OwnedFd, name: &CStr) -> Result<bool> {
    match unix_fs::statat(parent, name, AtFlags::SYMLINK_NOFOLLOW) {
        Ok(_stat) => Ok(true),
        Err(error) if error == rustix::io::Errno::NOENT => Ok(false),
        Err(error) => Err(unix_error(error)),
    }
}

#[cfg(unix)]
fn unix_inspect_node(parent: &OwnedFd, name: &CStr, label: &str) -> Result<Node> {
    let stat = match unix_fs::statat(parent, name, AtFlags::SYMLINK_NOFOLLOW) {
        Ok(stat) => stat,
        Err(error) if error == rustix::io::Errno::NOENT => return Ok(Node::Missing),
        Err(error) => return Err(unix_error(error)),
    };
    let version = unix_file_version(&stat)?;
    match FileType::from_raw_mode(stat.st_mode) {
        FileType::RegularFile => Ok(Node::File(version)),
        FileType::Directory => Ok(Node::Directory(version)),
        FileType::Symlink => Err(Error::InvalidBundle(format!(
            "{label} must not be a symbolic link"
        ))),
        FileType::Fifo
        | FileType::Socket
        | FileType::CharacterDevice
        | FileType::BlockDevice
        | FileType::Unknown => Err(Error::InvalidBundle(format!(
            "{label} is not a regular file or directory"
        ))),
    }
}

#[cfg(unix)]
fn unix_require_kind(fd: &OwnedFd, expected: FileType, label: &str) -> Result<()> {
    let stat = unix_fs::fstat(fd).map_err(unix_error)?;
    if FileType::from_raw_mode(stat.st_mode) != expected {
        return Err(Error::InvalidBundle(format!(
            "{label} has an invalid filesystem node type"
        )));
    }
    Ok(())
}

#[cfg(unix)]
fn unix_open_directory_at(
    parent: &OwnedFd,
    name: &CStr,
    expected: &FileVersion,
) -> Result<OwnedFd> {
    let fd = unix_fs::openat(
        parent,
        name,
        OFlags::RDONLY | OFlags::DIRECTORY | OFlags::NOFOLLOW | OFlags::NONBLOCK | OFlags::CLOEXEC,
        Mode::empty(),
    )
    .map_err(unix_error)?;
    unix_require_kind(&fd, FileType::Directory, "directory bundle directory")?;
    let observed = unix_file_version(&unix_fs::fstat(&fd).map_err(unix_error)?)?;
    if &observed != expected {
        return Err(changed("opening pinned directory bundle directory"));
    }
    Ok(fd)
}

#[cfg(unix)]
fn unix_ensure_node_version(
    parent: &OwnedFd,
    name: &CStr,
    expected: &FileVersion,
    context: &'static str,
) -> Result<()> {
    let observed = match unix_inspect_node(parent, name, "directory bundle source node")? {
        Node::File(version) | Node::Directory(version) => version,
        Node::Missing => return Err(changed(context)),
    };
    if &observed != expected {
        return Err(changed(context));
    }
    Ok(())
}

#[cfg(unix)]
#[derive(Debug, Clone, PartialEq, Eq)]
struct UnixManifestEntry {
    logical_name: Box<str>,
    physical_name: CString,
    version: FileVersion,
}

#[cfg(unix)]
fn unix_scan_manifest(index_fd: &OwnedFd, limits: Limits) -> Result<Vec<UnixManifestEntry>> {
    let mut directory = Dir::read_from(index_fd).map_err(unix_error)?;
    let mut entries = Vec::new();
    let mut total_bytes = 0u64;
    let mut metadata_bytes = 0u64;
    for entry_result in &mut directory {
        let directory_entry = entry_result.map_err(unix_error)?;
        let physical_name = directory_entry.file_name();
        if matches!(physical_name.to_bytes(), b"." | b"..") {
            continue;
        }
        let file_name = std::str::from_utf8(physical_name.to_bytes()).map_err(|_error| {
            Error::InvalidBundle(
                "directory bundle Index/ contains a non-UTF-8 member name".to_owned(),
            )
        })?;
        validate_loose_basename(file_name)?;
        if file_name == "Index.zip" {
            return Err(Error::InvalidBundle(
                "directory bundle loose Index/ contains a nested Index.zip".to_owned(),
            ));
        }
        if is_encryption_marker(file_name) {
            return Err(Error::Encrypted);
        }

        let name_length = "Index/".len().checked_add(file_name.len()).ok_or_else(|| {
            Error::InvalidBundle("directory entry name length overflowed usize".to_owned())
        })?;
        let name_bytes = u64::try_from(name_length).map_err(|_error| {
            Error::InvalidBundle("directory entry name length does not fit u64".to_owned())
        })?;
        check_member_name(name_bytes, limits)?;
        metadata_bytes = metadata_bytes.checked_add(name_bytes).ok_or_else(|| {
            Error::InvalidBundle("directory entry metadata byte count overflowed".to_owned())
        })?;
        check_metadata_bytes(metadata_bytes, limits)?;

        let node = unix_inspect_node(index_fd, physical_name, "directory bundle Index/ member")?;
        let Node::File(version) = node else {
            return Err(Error::InvalidBundle(format!(
                "directory bundle Index/ member {file_name} is not a regular file"
            )));
        };
        check_loose_entry_size(version.len, limits)?;
        total_bytes = total_bytes.checked_add(version.len).ok_or_else(|| {
            Error::InvalidBundle("directory bundle Index/ byte count overflowed".to_owned())
        })?;
        check_aggregate_input_bytes(total_bytes, limits)?;
        check_total_bytes(total_bytes, limits)?;
        let next_entries = entries.len().checked_add(1).ok_or_else(|| {
            Error::InvalidBundle("directory bundle entry count overflowed usize".to_owned())
        })?;
        check_entry_count(next_entries, limits)?;
        let logical_name = format!("Index/{file_name}");
        entries.try_reserve(1).map_err(|_error| Error::Allocation {
            resource: "directory bundle manifest",
            amount: 1,
        })?;
        entries.push(UnixManifestEntry {
            logical_name: logical_name.into_boxed_str(),
            physical_name: physical_name.to_owned(),
            version,
        });
    }
    entries.sort_unstable_by(|left, right| left.logical_name.cmp(&right.logical_name));
    Ok(entries)
}

#[cfg(unix)]
fn unix_read_manifest(
    index_fd: &OwnedFd,
    manifest: &[UnixManifestEntry],
    limits: Limits,
) -> Result<Vec<FrozenDirectoryEntry>> {
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(manifest.len())
        .map_err(|_error| Error::Allocation {
            resource: "directory bundle entries",
            amount: manifest.len(),
        })?;
    for item in manifest {
        let data = unix_read_stable_file_at(
            index_fd,
            &item.physical_name,
            &item.version,
            limits,
            FileRole::LooseEntry,
        )?;
        entries.push(FrozenDirectoryEntry {
            name: item.logical_name.clone(),
            data,
        });
    }
    Ok(entries)
}

#[cfg(unix)]
fn unix_read_stable_file_at(
    parent: &OwnedFd,
    name: &CStr,
    expected: &FileVersion,
    limits: Limits,
    role: FileRole,
) -> Result<Box<[u8]>> {
    check_file_size(expected.len, limits, role)?;
    let fd = unix_fs::openat(
        parent,
        name,
        OFlags::RDONLY | OFlags::NOFOLLOW | OFlags::NONBLOCK | OFlags::CLOEXEC,
        Mode::empty(),
    )
    .map_err(unix_error)?;
    unix_require_kind(&fd, FileType::RegularFile, "directory bundle member")?;
    let opened = unix_file_version(&unix_fs::fstat(&fd).map_err(unix_error)?)?;
    if &opened != expected {
        return Err(changed("opening pinned directory bundle member"));
    }

    let mut file = fs::File::from(fd);
    let bytes = read_bounded_contents(&mut file, expected.len, limits, role)?;
    let after = unix_file_version(&unix_fs::fstat(&file).map_err(unix_error)?)?;
    if after != opened {
        return Err(changed("reading pinned directory bundle member"));
    }
    unix_ensure_node_version(
        parent,
        name,
        expected,
        "verifying pinned directory bundle member",
    )?;
    if u64::try_from(bytes.len()).ok() != Some(expected.len) {
        return Err(changed("reading pinned directory bundle member"));
    }
    Ok(bytes)
}

#[cfg(unix)]
fn unix_file_version(stat: &Stat) -> Result<FileVersion> {
    Ok(FileVersion {
        len: unix_checked_u64(stat.st_size, "length")?,
        device: unix_checked_u64(stat.st_dev, "device identity")?,
        inode: unix_checked_u64(stat.st_ino, "inode identity")?,
        mode: unix_checked_u32(stat.st_mode, "mode")?,
        modified_seconds: unix_checked_i64(stat.st_mtime, "modification time")?,
        modified_nanoseconds: unix_checked_i64(stat.st_mtime_nsec, "modification nanoseconds")?,
        changed_seconds: unix_checked_i64(stat.st_ctime, "change time")?,
        changed_nanoseconds: unix_checked_i64(stat.st_ctime_nsec, "change nanoseconds")?,
    })
}

#[cfg(unix)]
fn unix_checked_u64<T>(value: T, field: &'static str) -> Result<u64>
where
    u64: TryFrom<T>,
{
    u64::try_from(value).map_err(|_error| {
        Error::InvalidBundle(format!("directory source {field} does not fit u64"))
    })
}

#[cfg(unix)]
fn unix_checked_u32<T>(value: T, field: &'static str) -> Result<u32>
where
    u32: TryFrom<T>,
{
    u32::try_from(value).map_err(|_error| {
        Error::InvalidBundle(format!("directory source {field} does not fit u32"))
    })
}

#[cfg(unix)]
fn unix_checked_i64<T>(value: T, field: &'static str) -> Result<i64>
where
    i64: TryFrom<T>,
{
    i64::try_from(value).map_err(|_error| {
        Error::InvalidBundle(format!("directory source {field} does not fit i64"))
    })
}

#[cfg(unix)]
fn unix_error(error: rustix::io::Errno) -> Error {
    Error::Io(std::io::Error::from_raw_os_error(error.raw_os_error()))
}

#[cfg(all(not(unix), not(windows)))]
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

#[cfg(all(not(unix), not(windows)))]
fn inspect_marker(path: &Path, label: &str) -> Result<Option<FileVersion>> {
    match inspect_node(path, label)? {
        Node::Missing => Ok(None),
        Node::File(version) => Ok(Some(version)),
        Node::Directory(_) => Err(Error::InvalidBundle(format!(
            "{label} is not a regular file"
        ))),
    }
}

#[cfg(all(not(unix), not(windows)))]
#[derive(Debug, Clone, PartialEq, Eq)]
struct ManifestEntry {
    logical_name: Box<str>,
    path: PathBuf,
    version: FileVersion,
}

#[cfg(all(not(unix), not(windows)))]
fn scan_manifest(path: &Path, limits: Limits) -> Result<Vec<ManifestEntry>> {
    let mut entries = Vec::new();
    let mut total_bytes = 0u64;
    let mut metadata_bytes = 0u64;
    for entry_result in fs::read_dir(path)? {
        let directory_entry = entry_result?;
        let file_name = utf8_name(directory_entry.file_name())?;
        validate_loose_basename(&file_name)?;
        if file_name == "Index.zip" {
            return Err(Error::InvalidBundle(
                "directory bundle loose Index/ contains a nested Index.zip".to_owned(),
            ));
        }
        if is_encryption_marker(&file_name) {
            return Err(Error::Encrypted);
        }
        let name_length = "Index/".len().checked_add(file_name.len()).ok_or_else(|| {
            Error::InvalidBundle("directory entry name length overflowed usize".to_owned())
        })?;
        let name_bytes = u64::try_from(name_length).map_err(|_error| {
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
                "directory bundle Index/ member {file_name} is not a regular file"
            )));
        };
        check_loose_entry_size(version.len, limits)?;
        total_bytes = total_bytes.checked_add(version.len).ok_or_else(|| {
            Error::InvalidBundle("directory bundle Index/ byte count overflowed".to_owned())
        })?;
        check_aggregate_input_bytes(total_bytes, limits)?;
        check_total_bytes(total_bytes, limits)?;
        let next_entries = entries.len().checked_add(1).ok_or_else(|| {
            Error::InvalidBundle("directory bundle entry count overflowed usize".to_owned())
        })?;
        check_entry_count(next_entries, limits)?;
        let logical_name = format!("Index/{file_name}");
        entries.try_reserve(1).map_err(|_error| Error::Allocation {
            resource: "directory bundle manifest",
            amount: 1,
        })?;
        entries.push(ManifestEntry {
            logical_name: logical_name.into_boxed_str(),
            path: entry_path,
            version,
        });
    }
    entries.sort_unstable_by(|left, right| left.logical_name.cmp(&right.logical_name));
    Ok(entries)
}

#[cfg(all(not(unix), not(windows)))]
fn utf8_name(name: OsString) -> Result<String> {
    name.into_string().map_err(|_name| {
        Error::InvalidBundle("directory bundle Index/ contains a non-UTF-8 member name".to_owned())
    })
}

#[derive(Debug, Clone, Copy)]
enum FileRole {
    IndexZip,
    LooseEntry,
    Sidecar,
}

#[cfg(all(not(unix), not(windows)))]
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
    let mut file = open.open(path)?;
    let opened_metadata = file.metadata()?;
    let opened = FileVersion::from_metadata(&opened_metadata);
    if opened != *expected || !opened_metadata.is_file() {
        return Err(changed("opening directory bundle member"));
    }
    ensure_path_version(path, expected, "opening directory bundle member")?;

    let bytes = read_bounded_contents(&mut file, expected.len, limits, role)?;

    let after = FileVersion::from_metadata(&file.metadata()?);
    if after != opened {
        return Err(changed("reading directory bundle member"));
    }
    ensure_path_version(path, expected, "verifying directory bundle member")?;
    if u64::try_from(bytes.len()).ok() != Some(expected.len) {
        return Err(changed("reading directory bundle member"));
    }
    Ok(bytes)
}

fn read_bounded_contents(
    reader: &mut impl Read,
    expected_len: u64,
    limits: Limits,
    role: FileRole,
) -> Result<Box<[u8]>> {
    let mut bytes = Vec::new();
    let initial = usize::try_from(expected_len)
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
        let read = match reader.read(&mut buffer) {
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
    if !matches!(role, FileRole::IndexZip) && observed > limits.max_entry_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::EntryBytes,
            observed,
            maximum: limits.max_entry_bytes(),
        });
    }
    if matches!(role, FileRole::Sidecar) && observed > MAX_DIRECTORY_PROPERTIES_BYTES {
        return Err(Error::Limit {
            kind: LimitKind::EntryBytes,
            observed,
            maximum: MAX_DIRECTORY_PROPERTIES_BYTES,
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
    let maximum = limits.max_metadata_bytes();
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

fn check_aggregate_input_bytes(observed: u64, limits: Limits) -> Result<()> {
    if observed > limits.max_input_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed,
            maximum: limits.max_input_bytes(),
        });
    }
    Ok(())
}

#[cfg(all(not(unix), not(windows)))]
fn reject_path_encryption_markers(root: &Path) -> Result<()> {
    for name in [".iwpv2", ".iwph"] {
        if !matches!(
            inspect_node(&root.join(name), "directory encryption marker")?,
            Node::Missing
        ) {
            return Err(Error::Encrypted);
        }
    }
    let metadata = root.join("Metadata");
    match inspect_node(&metadata, "directory bundle Metadata")? {
        Node::Missing => Ok(()),
        Node::Directory(_) => {
            for name in [".iwpv2", ".iwph"] {
                if !matches!(
                    inspect_node(&metadata.join(name), "directory metadata encryption marker")?,
                    Node::Missing
                ) {
                    return Err(Error::Encrypted);
                }
            }
            Ok(())
        },
        Node::File(_) => Err(Error::InvalidBundle(
            "directory bundle Metadata is not a directory".to_owned(),
        )),
    }
}

#[cfg(all(not(unix), not(windows)))]
fn inspect_sidecars(root: &Path, profile: CaptureProfile) -> Result<MetadataCapture> {
    let metadata_path = root.join("Metadata");
    match inspect_node(&metadata_path, "directory bundle Metadata")? {
        Node::Missing => Ok(MetadataCapture::default()),
        Node::File(_) => Err(Error::InvalidBundle(
            "directory bundle Metadata is not a directory".to_owned(),
        )),
        Node::Directory(directory) => {
            let mut capture = MetadataCapture {
                directory: Some(directory),
                ..MetadataCapture::default()
            };
            for sidecar in Sidecar::ALL {
                if !profile.captures(sidecar) {
                    continue;
                }
                capture.sidecar_mut(sidecar).selected = true;
                match inspect_node(&metadata_path.join(sidecar.basename()), sidecar.label())? {
                    Node::Missing => {},
                    Node::Directory(_) => {
                        return Err(Error::InvalidBundle(format!(
                            "{} is not a regular file",
                            sidecar.label()
                        )));
                    },
                    Node::File(file) => capture.sidecar_mut(sidecar).file = Some(file),
                }
            }
            Ok(capture)
        },
    }
}

#[cfg(all(not(unix), not(windows)))]
fn read_sidecars(root: &Path, capture: &mut MetadataCapture, limits: Limits) -> Result<()> {
    for sidecar in Sidecar::ALL {
        if !capture.sidecar(sidecar).selected {
            continue;
        }
        let Some(file) = capture.sidecar(sidecar).file.as_ref() else {
            continue;
        };
        let data = read_stable_file(
            &root.join("Metadata").join(sidecar.basename()),
            file,
            limits,
            FileRole::Sidecar,
        )?;
        capture.sidecar_mut(sidecar).data = Some(data);
    }
    Ok(())
}

#[cfg(all(not(unix), not(windows)))]
fn verify_sidecars(root: &Path, capture: &MetadataCapture) -> Result<()> {
    let metadata_path = root.join("Metadata");
    match &capture.directory {
        None => {
            if !matches!(
                inspect_node(&metadata_path, "directory bundle Metadata")?,
                Node::Missing
            ) {
                return Err(changed("verifying directory bundle Metadata"));
            }
        },
        Some(directory) => {
            ensure_path_version(
                &metadata_path,
                directory,
                "verifying directory bundle Metadata",
            )?;
            for sidecar in Sidecar::ALL {
                if !capture.sidecar(sidecar).selected {
                    continue;
                }
                let sidecar_path = metadata_path.join(sidecar.basename());
                match &capture.sidecar(sidecar).file {
                    None => {
                        if !matches!(inspect_node(&sidecar_path, sidecar.label())?, Node::Missing) {
                            return Err(changed(sidecar.verification_context()));
                        }
                    },
                    Some(file) => {
                        ensure_path_version(&sidecar_path, file, sidecar.verification_context())?
                    },
                }
            }
        },
    }
    Ok(())
}

fn is_encryption_marker(name: &str) -> bool {
    matches!(name.rsplit('/').next(), Some(".iwpv2" | ".iwph"))
}

fn validate_loose_basename(name: &str) -> Result<()> {
    if name.is_empty()
        || matches!(name, "." | "..")
        || name.contains(['/', '\\', ':'])
        || name.chars().any(char::is_control)
    {
        return Err(Error::InvalidBundle(format!(
            "directory bundle Index/ member name is not an exact portable basename: {name:?}"
        )));
    }
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Node {
    Missing,
    File(FileVersion),
    Directory(FileVersion),
}

#[cfg(all(not(unix), not(windows)))]
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

#[cfg(all(not(unix), not(windows)))]
fn require_file(path: &Path, label: &str) -> Result<FileVersion> {
    match inspect_node(path, label)? {
        Node::File(version) => Ok(version),
        Node::Missing => Err(Error::InvalidBundle(format!("{label} is missing"))),
        Node::Directory(_) => Err(Error::InvalidBundle(format!(
            "{label} is not a regular file"
        ))),
    }
}

#[cfg(all(not(unix), not(windows)))]
fn require_directory(path: &Path, label: &str) -> Result<FileVersion> {
    match inspect_node(path, label)? {
        Node::Directory(version) => Ok(version),
        Node::Missing => Err(Error::InvalidBundle(format!("{label} is missing"))),
        Node::File(_) => Err(Error::InvalidBundle(format!("{label} is not a directory"))),
    }
}

#[cfg(all(not(unix), not(windows)))]
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

#[cfg(all(not(unix), not(windows)))]
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

#[cfg(all(not(unix), not(windows)))]
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
    #[cfg(all(not(unix), not(windows)))]
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
    #[cfg(all(not(unix), not(windows)))]
    fn from_metadata(metadata: &Metadata) -> Self {
        Self {
            len: metadata.len(),
            modified: metadata.modified().ok(),
        }
    }
}

#[cfg(test)]
mod tests {
    use std::{
        fs,
        path::PathBuf,
        sync::atomic::{AtomicU64, Ordering},
    };

    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

    fn test_file_version(len: u64) -> FileVersion {
        FileVersion {
            len,
            #[cfg(all(not(unix), not(windows)))]
            modified: None,
            #[cfg(unix)]
            device: 0,
            #[cfg(unix)]
            inode: 0,
            #[cfg(unix)]
            mode: 0,
            #[cfg(unix)]
            modified_seconds: 0,
            #[cfg(unix)]
            modified_nanoseconds: 0,
            #[cfg(unix)]
            changed_seconds: 0,
            #[cfg(unix)]
            changed_nanoseconds: 0,
        }
    }

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

    #[cfg(windows)]
    #[test]
    fn windows_directory_ingress_fails_closed() -> Result<()> {
        let temp = Temp::new()?;
        assert!(matches!(
            FrozenDirectoryBundle::open(temp.path()),
            Err(Error::InvalidBundle(message)) if message.contains("unsupported on Windows")
        ));
        Ok(())
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
    fn freezes_only_the_canonical_properties_diagnostic_under_aggregate_limits() -> Result<()> {
        let temp = Temp::new()?;
        let index = index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?;
        let properties = b"canonical-properties";
        fs::write(temp.path().join("Index.zip"), &index)?;
        fs::create_dir(temp.path().join("Metadata"))?;
        fs::write(temp.path().join("Metadata/Properties.plist"), properties)?;
        fs::write(
            temp.path().join("Metadata/BuildVersionHistory.plist"),
            b"ignored",
        )?;
        fs::create_dir(temp.path().join("Data"))?;
        fs::write(temp.path().join("Data/media.bin"), b"ignored")?;

        let exact_input = u64::try_from(index.len() + properties.len()).map_err(|_error| {
            Error::InvalidBundle("test aggregate length does not fit u64".to_owned())
        })?;
        let limits = Limits::new(
            exact_input,
            Limits::MAX_ENTRIES,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        let index_only = FrozenDirectoryBundle::open_with_limits(temp.path(), limits)?;
        assert_eq!(index_only.properties_plist(), None);
        let frozen = FrozenDirectoryBundle::open_with_properties(temp.path(), limits)?;
        assert_eq!(frozen.properties_plist(), Some(properties.as_slice()));
        let (components, captured) = frozen.into_semantic_parts();
        assert_eq!(components.len(), 1);
        assert_eq!(captured.as_deref(), Some(properties.as_slice()));

        let too_small = Limits::new(
            exact_input - 1,
            Limits::MAX_ENTRIES,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_properties(temp.path(), too_small),
            Err(Error::Limit {
                kind: LimitKind::InputBytes,
                observed,
                maximum,
            }) if observed == exact_input && maximum == exact_input - 1
        ));

        let property_bytes = u64::try_from(properties.len()).map_err(|_error| {
            Error::InvalidBundle("test properties length does not fit u64".to_owned())
        })?;
        let no_index_input = Limits::new(
            property_bytes,
            Limits::MAX_ENTRIES,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_properties(temp.path(), no_index_input),
            Err(Error::Limit {
                kind: LimitKind::InputBytes,
                observed,
                maximum,
            }) if observed == property_bytes + 1 && maximum == property_bytes
        ));

        let no_index_entry = Limits::new(
            Limits::MAX_INPUT_BYTES,
            1,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_properties(temp.path(), no_index_entry),
            Err(Error::Limit {
                kind: LimitKind::Entries,
                observed: 2,
                maximum: 1,
            })
        ));

        let no_index_total = Limits::new(
            Limits::MAX_INPUT_BYTES,
            Limits::MAX_ENTRIES,
            Limits::MAX_ENTRY_BYTES,
            property_bytes,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_properties(temp.path(), no_index_total),
            Err(Error::Limit {
                kind: LimitKind::TotalBytes,
                observed,
                maximum,
            }) if observed == property_bytes + 1 && maximum == property_bytes
        ));
        Ok(())
    }

    #[test]
    fn pages_metadata_profile_freezes_only_all_three_canonical_sidecars() -> Result<()> {
        let temp = Temp::new()?;
        let index = index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?;
        let properties = b"canonical-properties";
        let history = b"canonical-history";
        let identifier = b"canonical-identifier";
        fs::write(temp.path().join("Index.zip"), &index)?;
        fs::create_dir(temp.path().join("Metadata"))?;
        fs::write(temp.path().join("Metadata/Properties.plist"), properties)?;
        fs::write(
            temp.path().join("Metadata/BuildVersionHistory.plist"),
            history,
        )?;
        fs::write(temp.path().join("Metadata/DocumentIdentifier"), identifier)?;
        fs::create_dir(temp.path().join("Decoy"))?;
        fs::write(temp.path().join("Decoy/Properties.plist"), b"decoy")?;
        fs::write(temp.path().join("Metadata/Other.plist"), b"ignored")?;

        let frozen =
            FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), Limits::default())?;
        let (components, sidecars) = frozen.into_pages_semantic_parts();
        assert_eq!(components.len(), 1);
        assert_eq!(sidecars.properties_plist(), Some(properties.as_slice()));
        assert_eq!(
            sidecars.build_version_history_plist(),
            Some(history.as_slice())
        );
        assert_eq!(sidecars.document_identifier(), Some(identifier.as_slice()));
        Ok(())
    }

    #[test]
    fn pages_metadata_profile_charges_all_three_sidecars_before_index_capture() -> Result<()> {
        let temp = Temp::new()?;
        let index = index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?;
        let sidecars = [
            b"properties".as_slice(),
            b"history".as_slice(),
            b"identifier".as_slice(),
        ];
        fs::write(temp.path().join("Index.zip"), &index)?;
        fs::create_dir(temp.path().join("Metadata"))?;
        fs::write(temp.path().join("Metadata/Properties.plist"), sidecars[0])?;
        fs::write(
            temp.path().join("Metadata/BuildVersionHistory.plist"),
            sidecars[1],
        )?;
        fs::write(temp.path().join("Metadata/DocumentIdentifier"), sidecars[2])?;
        let exact = u64::try_from(
            index.len() + sidecars.iter().map(|sidecar| sidecar.len()).sum::<usize>(),
        )
        .map_err(|_error| Error::InvalidBundle("test input does not fit u64".to_owned()))?;
        let limits = Limits::new(
            exact,
            Limits::MAX_ENTRIES,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), limits).is_ok());
        let too_small = Limits::new(
            exact - 1,
            Limits::MAX_ENTRIES,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), too_small),
            Err(Error::Limit { kind: LimitKind::InputBytes, observed, maximum })
                if observed == exact && maximum == exact - 1
        ));
        Ok(())
    }

    #[test]
    fn pages_metadata_profile_enforces_each_hard_and_caller_entry_ceiling() -> Result<()> {
        let authorities = [
            "Properties.plist",
            "BuildVersionHistory.plist",
            "DocumentIdentifier",
        ];
        for authority in authorities {
            let temp = Temp::new()?;
            fs::write(
                temp.path().join("Index.zip"),
                index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?,
            )?;
            fs::create_dir(temp.path().join("Metadata"))?;
            let exact = vec![
                b'x';
                usize::try_from(MAX_DIRECTORY_PROPERTIES_BYTES).map_err(|_error| {
                    Error::InvalidBundle("test hard cap does not fit usize".to_owned())
                },)?
            ];
            fs::write(temp.path().join("Metadata").join(authority), &exact)?;
            assert!(
                FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), Limits::default())
                    .is_ok(),
                "exact hard cap must accept {authority}"
            );

            fs::write(
                temp.path().join("Metadata").join(authority),
                vec![b'x'; exact.len() + 1],
            )?;
            assert!(matches!(
                FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), Limits::default()),
                Err(Error::Limit {
                    kind: LimitKind::EntryBytes,
                    observed,
                    maximum,
                }) if observed == MAX_DIRECTORY_PROPERTIES_BYTES + 1
                    && maximum == MAX_DIRECTORY_PROPERTIES_BYTES
            ));

            fs::write(temp.path().join("Metadata").join(authority), b"1234")?;
            let caller_limited = Limits::new(
                Limits::MAX_INPUT_BYTES,
                Limits::MAX_ENTRIES,
                3,
                Limits::MAX_TOTAL_BYTES,
                Limits::MAX_IWA_STREAM_BYTES,
            )?;
            assert!(matches!(
                FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), caller_limited),
                Err(Error::Limit {
                    kind: LimitKind::EntryBytes,
                    observed: 4,
                    maximum: 3,
                })
            ));
        }
        Ok(())
    }

    #[test]
    fn sidecar_incremental_reads_enforce_the_hard_cap_before_growth_allocation() -> Result<()> {
        let oversized = vec![
            b'x';
            usize::try_from(MAX_DIRECTORY_PROPERTIES_BYTES + 1).map_err(
                |_error| { Error::InvalidBundle("test hard cap does not fit usize".to_owned()) }
            )?
        ];
        let mut reader = std::io::Cursor::new(oversized);
        assert!(matches!(
            read_bounded_contents(
                &mut reader,
                1,
                Limits::default(),
                FileRole::Sidecar
            ),
            Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed,
                maximum,
            }) if observed > MAX_DIRECTORY_PROPERTIES_BYTES
                && maximum == MAX_DIRECTORY_PROPERTIES_BYTES
        ));
        Ok(())
    }

    #[test]
    fn pages_metadata_profile_charges_exact_entry_and_expanded_totals() -> Result<()> {
        let temp = Temp::new()?;
        let index = index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?;
        let sidecars = [
            b"properties".as_slice(),
            b"history".as_slice(),
            b"identifier".as_slice(),
        ];
        fs::write(temp.path().join("Index.zip"), &index)?;
        fs::create_dir(temp.path().join("Metadata"))?;
        fs::write(temp.path().join("Metadata/Properties.plist"), sidecars[0])?;
        fs::write(
            temp.path().join("Metadata/BuildVersionHistory.plist"),
            sidecars[1],
        )?;
        fs::write(temp.path().join("Metadata/DocumentIdentifier"), sidecars[2])?;

        let source = FrozenDirectoryBundle::open(temp.path())?;
        let index_entries = source.state.index_report.entries;
        let expanded = source
            .state
            .index_report
            .expanded_bytes
            .checked_add(
                u64::try_from(sidecars.iter().map(|sidecar| sidecar.len()).sum::<usize>())
                    .map_err(|_error| {
                        Error::InvalidBundle("test sidecar total does not fit u64".to_owned())
                    })?,
            )
            .ok_or_else(|| Error::InvalidBundle("test expanded total overflowed".to_owned()))?;

        let exact_entries = Limits::new(
            Limits::MAX_INPUT_BYTES,
            index_entries + 3,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(
            FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), exact_entries).is_ok()
        );
        let one_under_entries = Limits::new(
            Limits::MAX_INPUT_BYTES,
            index_entries + 2,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), one_under_entries),
            Err(Error::Limit {
                kind: LimitKind::Entries,
                observed,
                maximum,
            }) if observed == u64::try_from(index_entries + 3).unwrap_or(u64::MAX)
                && maximum == u64::try_from(index_entries + 2).unwrap_or(u64::MAX)
        ));

        let exact_total = Limits::new(
            Limits::MAX_INPUT_BYTES,
            Limits::MAX_ENTRIES,
            Limits::MAX_ENTRY_BYTES,
            expanded,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), exact_total).is_ok());
        let one_under_total = Limits::new(
            Limits::MAX_INPUT_BYTES,
            Limits::MAX_ENTRIES,
            Limits::MAX_ENTRY_BYTES,
            expanded - 1,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), one_under_total),
            Err(Error::Limit {
                kind: LimitKind::TotalBytes,
                observed,
                maximum,
            }) if observed == expanded && maximum == expanded - 1
        ));
        Ok(())
    }

    #[test]
    fn pages_metadata_profile_reserves_exact_name_metadata_before_index_parse() -> Result<()> {
        let sidecar_names = [
            PROPERTIES_LOGICAL_NAME,
            BUILD_VERSION_HISTORY_LOGICAL_NAME,
            DOCUMENT_IDENTIFIER_LOGICAL_NAME,
        ];
        let sidecar_metadata = sidecar_names.iter().try_fold(0_u64, |total, name| {
            total
                .checked_add(u64::try_from(name.len()).map_err(|_error| {
                    Error::InvalidBundle("test name length does not fit u64".to_owned())
                })?)
                .ok_or_else(|| Error::InvalidBundle("test metadata total overflowed".to_owned()))
        })?;
        let index_metadata = u64::try_from("Index/Document.iwa".len())
            .map_err(|_error| Error::InvalidBundle("test name does not fit u64".to_owned()))?;
        let outer = Limits::default();
        let capture = MetadataCapture {
            properties: SidecarCapture {
                selected: true,
                file: Some(test_file_version(1)),
                data: None,
            },
            build_version_history: SidecarCapture {
                selected: true,
                file: Some(test_file_version(1)),
                data: None,
            },
            document_identifier: SidecarCapture {
                selected: true,
                file: Some(test_file_version(1)),
                data: None,
            },
            ..MetadataCapture::default()
        };
        let reserved = limits_without_sidecars(outer, &capture)?;
        assert_eq!(
            reserved.max_metadata_bytes(),
            outer.max_metadata_bytes() - sidecar_metadata
        );

        let exact = reserved.with_derived_metadata_bytes(index_metadata)?;
        assert_eq!(exact.max_metadata_bytes(), index_metadata);
        let one_under = exact.with_derived_metadata_bytes(index_metadata - 1)?;
        let error = remap_reserved_limit(
            Error::Limit {
                kind: LimitKind::MetadataBytes,
                observed: index_metadata,
                maximum: one_under.max_metadata_bytes(),
            },
            outer,
            &capture,
        );
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::MetadataBytes,
                observed,
                maximum,
            } if observed == sidecar_metadata + index_metadata
                && maximum == outer.max_metadata_bytes()
        ));

        let temp = Temp::new()?;
        fs::write(
            temp.path().join("Index.zip"),
            index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?,
        )?;
        fs::create_dir(temp.path().join("Metadata"))?;
        fs::write(temp.path().join("Metadata/Properties.plist"), b"p")?;
        fs::write(temp.path().join("Metadata/BuildVersionHistory.plist"), b"b")?;
        fs::write(temp.path().join("Metadata/DocumentIdentifier"), b"i")?;
        let index = FrozenDirectoryBundle::open(temp.path())?;
        let exact_metadata = index
            .state
            .index_report
            .metadata_bytes
            .checked_add(sidecar_metadata)
            .ok_or_else(|| Error::InvalidBundle("test metadata total overflowed".to_owned()))?;
        let exact_limits = Limits::default().with_derived_metadata_bytes(exact_metadata)?;
        assert!(FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), exact_limits).is_ok());
        let one_under_limits = Limits::default().with_derived_metadata_bytes(exact_metadata - 1)?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), one_under_limits),
            Err(Error::Limit {
                kind: LimitKind::MetadataBytes,
                observed,
                maximum,
            }) if observed == exact_metadata && maximum == exact_metadata - 1
        ));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn rejects_symbolic_properties_diagnostic() -> Result<()> {
        use std::os::unix::fs::symlink;

        let temp = Temp::new()?;
        fs::write(
            temp.path().join("Index.zip"),
            index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?,
        )?;
        fs::create_dir(temp.path().join("Metadata"))?;
        fs::write(temp.path().join("outside.plist"), b"outside")?;
        symlink(
            temp.path().join("outside.plist"),
            temp.path().join("Metadata/Properties.plist"),
        )?;

        assert!(matches!(
            FrozenDirectoryBundle::open_with_properties(temp.path(), Limits::default()),
            Err(Error::InvalidBundle(message))
                if message.contains("Metadata/Properties.plist")
        ));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn pages_metadata_profile_rejects_selected_symlink_and_directory() -> Result<()> {
        use std::os::unix::fs::symlink;
        use std::os::unix::net::UnixListener;

        let temp = Temp::new()?;
        fs::write(
            temp.path().join("Index.zip"),
            index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?,
        )?;
        fs::create_dir(temp.path().join("Metadata"))?;
        fs::write(temp.path().join("outside"), b"outside")?;
        symlink(
            temp.path().join("outside"),
            temp.path().join("Metadata/DocumentIdentifier"),
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), Limits::default()),
            Err(Error::InvalidBundle(message)) if message.contains("Metadata/DocumentIdentifier")
        ));
        fs::remove_file(temp.path().join("Metadata/DocumentIdentifier"))?;
        fs::create_dir(temp.path().join("Metadata/BuildVersionHistory.plist"))?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), Limits::default()),
            Err(Error::InvalidBundle(message)) if message.contains("Metadata/BuildVersionHistory.plist")
        ));
        fs::remove_dir(temp.path().join("Metadata/BuildVersionHistory.plist"))?;
        let socket_path = temp.path().join("s");
        let _socket = UnixListener::bind(&socket_path)?;
        fs::rename(socket_path, temp.path().join("Metadata/Properties.plist"))?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_pages_metadata(temp.path(), Limits::default()),
            Err(Error::InvalidBundle(message)) if message.contains("Metadata/Properties.plist")
        ));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn index_only_profile_ignores_all_pages_sidecar_authorities() -> Result<()> {
        use std::os::unix::fs::symlink;

        let temp = Temp::new()?;
        let index = index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?;
        fs::write(temp.path().join("Index.zip"), &index)?;
        fs::create_dir(temp.path().join("Metadata"))?;
        let outside = temp.path().join("outside.plist");
        fs::write(&outside, b"outside")?;
        symlink(&outside, temp.path().join("Metadata/Properties.plist"))?;
        fs::create_dir(temp.path().join("Metadata/BuildVersionHistory.plist"))?;
        fs::write(
            temp.path().join("Metadata/DocumentIdentifier"),
            vec![b'x'; usize::try_from(MAX_DIRECTORY_PROPERTIES_BYTES + 1).unwrap()],
        )?;

        let limits = Limits::new(
            u64::try_from(index.len()).unwrap(),
            1,
            Limits::MAX_ENTRY_BYTES,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        let frozen = FrozenDirectoryBundle::open_with_limits(temp.path(), limits)?;
        assert!(frozen.properties_plist().is_none());
        let (_components, sidecars) = frozen.into_pages_semantic_parts();
        assert!(sidecars.properties_plist().is_none());
        assert!(sidecars.build_version_history_plist().is_none());
        assert!(sidecars.document_identifier().is_none());
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn pinned_pages_sidecars_refuse_replacement_deletion_and_type_changes() -> Result<()> {
        fn root_fd(path: &Path) -> Result<OwnedFd> {
            unix_fs::open(
                path,
                OFlags::RDONLY | OFlags::DIRECTORY | OFlags::NOFOLLOW | OFlags::CLOEXEC,
                Mode::empty(),
            )
            .map_err(unix_error)
        }

        let replacement = Temp::new()?;
        fs::create_dir(replacement.path().join("Metadata"))?;
        fs::write(
            replacement.path().join("Metadata/Properties.plist"),
            b"first",
        )?;
        let replacement_fd = root_fd(replacement.path())?;
        let mut capture = unix_inspect_sidecars(&replacement_fd, CaptureProfile::PagesMetadata)?;
        fs::write(replacement.path().join("replacement.plist"), b"other")?;
        fs::rename(
            replacement.path().join("replacement.plist"),
            replacement.path().join("Metadata/Properties.plist"),
        )?;
        assert!(matches!(
            unix_read_sidecars(&replacement_fd, &mut capture, Limits::default()),
            Err(Error::DirectoryChanged { .. })
        ));

        let deletion = Temp::new()?;
        fs::create_dir(deletion.path().join("Metadata"))?;
        fs::write(
            deletion.path().join("Metadata/DocumentIdentifier"),
            b"identifier",
        )?;
        let deletion_fd = root_fd(deletion.path())?;
        let mut capture = unix_inspect_sidecars(&deletion_fd, CaptureProfile::PagesMetadata)?;
        fs::remove_file(deletion.path().join("Metadata/DocumentIdentifier"))?;
        assert!(unix_read_sidecars(&deletion_fd, &mut capture, Limits::default()).is_err());

        let changed_type = Temp::new()?;
        fs::create_dir(changed_type.path().join("Metadata"))?;
        fs::write(
            changed_type
                .path()
                .join("Metadata/BuildVersionHistory.plist"),
            b"history",
        )?;
        let changed_type_fd = root_fd(changed_type.path())?;
        let mut capture = unix_inspect_sidecars(&changed_type_fd, CaptureProfile::PagesMetadata)?;
        fs::remove_file(
            changed_type
                .path()
                .join("Metadata/BuildVersionHistory.plist"),
        )?;
        fs::create_dir(
            changed_type
                .path()
                .join("Metadata/BuildVersionHistory.plist"),
        )?;
        assert!(unix_read_sidecars(&changed_type_fd, &mut capture, Limits::default()).is_err());
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
    fn rejects_directory_encryption_markers_in_every_supported_location() -> Result<()> {
        for location in [".iwpv2", "Metadata/.iwph", "Index/.iwpv2"] {
            let temp = Temp::new()?;
            fs::create_dir(temp.path().join("Index"))?;
            fs::write(temp.path().join("Index/Document.iwa"), iwa(1, 10_000)?)?;
            if location.starts_with("Metadata/") {
                fs::create_dir(temp.path().join("Metadata"))?;
            }
            fs::write(temp.path().join(location), b"encrypted")?;

            assert!(matches!(
                FrozenDirectoryBundle::open(temp.path()),
                Err(Error::Encrypted)
            ));
        }
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn rejects_nonportable_loose_member_basenames() -> Result<()> {
        for name in ["colon:name.iwa", "line\nbreak.iwa", r"back\slash.iwa"] {
            let temp = Temp::new()?;
            fs::create_dir(temp.path().join("Index"))?;
            fs::write(temp.path().join("Index").join(name), b"unsafe")?;

            assert!(matches!(
                FrozenDirectoryBundle::open(temp.path()),
                Err(Error::InvalidBundle(message))
                    if message.contains("not an exact portable basename")
            ));
        }
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

        let aggregate_input_limited = Limits::new(
            total - 1,
            2,
            maximum_entry,
            Limits::MAX_TOTAL_BYTES,
            Limits::MAX_IWA_STREAM_BYTES,
        )?;
        assert!(matches!(
            FrozenDirectoryBundle::open_with_limits(temp.path(), aggregate_input_limited),
            Err(Error::Limit {
                kind: LimitKind::InputBytes,
                observed,
                maximum
            }) if observed == total && maximum == total - 1
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
        let root_fd = unix_fs::open(
            temp.path(),
            OFlags::RDONLY | OFlags::DIRECTORY | OFlags::NOFOLLOW | OFlags::CLOEXEC,
            Mode::empty(),
        )
        .map_err(unix_error)?;
        let Node::File(version) = unix_inspect_node(&root_fd, c"Document.iwa", "test member")?
        else {
            return Err(Error::InvalidBundle(
                "test member was not a regular file".to_owned(),
            ));
        };
        let replacement = temp.path().join("replacement.iwa");
        fs::write(&replacement, b"other")?;
        fs::rename(&replacement, &path)?;

        assert!(matches!(
            unix_read_stable_file_at(
                &root_fd,
                c"Document.iwa",
                &version,
                Limits::default(),
                FileRole::LooseEntry
            ),
            Err(Error::DirectoryChanged { .. })
        ));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn exact_index_revalidation_runs_after_component_parsing() -> Result<()> {
        let temp = Temp::new()?;
        let source = index_zip(&[("Index/Document.iwa", iwa(1, 10_000)?)])?;
        fs::write(temp.path().join("Index.zip"), &source)?;
        let root_fd = unix_fs::open(
            temp.path(),
            OFlags::RDONLY | OFlags::DIRECTORY | OFlags::NOFOLLOW | OFlags::CLOEXEC,
            Mode::empty(),
        )
        .map_err(unix_error)?;
        let Node::File(version) = unix_inspect_node(&root_fd, c"Index.zip", "test Index.zip")?
        else {
            return Err(Error::InvalidBundle(
                "test Index.zip was not a regular file".to_owned(),
            ));
        };
        let captured = unix_read_stable_file_at(
            &root_fd,
            c"Index.zip",
            &version,
            Limits::default(),
            FileRole::IndexZip,
        )?;
        let snapshot = FrozenDirectoryBundle::from_captured_index_zip(
            captured,
            DirectoryMarkers::default(),
            Limits::default(),
            Limits::default(),
        )?;
        assert_eq!(snapshot.components().len(), 1);

        let replacement = temp.path().join("replacement.zip");
        fs::write(&replacement, &source)?;
        fs::rename(replacement, temp.path().join("Index.zip"))?;
        assert!(matches!(
            unix_ensure_node_version(
                &root_fd,
                c"Index.zip",
                &version,
                "test final Index.zip verification"
            ),
            Err(Error::DirectoryChanged { .. })
        ));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn opened_root_is_pinned_when_its_path_is_replaced() -> Result<()> {
        let parent = Temp::new()?;
        let source = parent.path().join("source.pages");
        fs::create_dir(&source)?;
        fs::create_dir(source.join("Index"))?;
        fs::write(source.join("Index/Document.iwa"), iwa(1, 10_000)?)?;
        let root_fd = unix_fs::open(
            &source,
            OFlags::RDONLY | OFlags::DIRECTORY | OFlags::NOFOLLOW | OFlags::CLOEXEC,
            Mode::empty(),
        )
        .map_err(unix_error)?;

        fs::rename(&source, parent.path().join("original.pages"))?;
        fs::create_dir(&source)?;
        fs::create_dir(source.join("Index"))?;
        fs::write(source.join("Index/not-a-document.iwa"), b"malicious")?;

        let frozen = FrozenDirectoryBundle::open_unix_root(
            &root_fd,
            Limits::default(),
            CaptureProfile::None,
            None,
        )?;
        assert_eq!(frozen.components().len(), 1);
        assert!(frozen.components().get("Index/Document.iwa").is_some());
        assert!(
            frozen
                .components()
                .get("Index/not-a-document.iwa")
                .is_none()
        );
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn replaced_index_path_cannot_redirect_or_publish_pinned_capture() -> Result<()> {
        let temp = Temp::new()?;
        let index = temp.path().join("Index");
        fs::create_dir(&index)?;
        fs::write(index.join("Document.iwa"), iwa(1, 10_000)?)?;
        let root_fd = unix_fs::open(
            temp.path(),
            OFlags::RDONLY | OFlags::DIRECTORY | OFlags::NOFOLLOW | OFlags::CLOEXEC,
            Mode::empty(),
        )
        .map_err(unix_error)?;
        let Node::Directory(version) = unix_inspect_node(&root_fd, c"Index", "test Index")? else {
            return Err(Error::InvalidBundle(
                "test Index was not a directory".to_owned(),
            ));
        };
        let index_fd = unix_open_directory_at(&root_fd, c"Index", &version)?;
        let manifest = unix_scan_manifest(&index_fd, Limits::default())?;
        let entries = unix_read_manifest(&index_fd, &manifest, Limits::default())?;
        let frozen = FrozenDirectoryBundle::from_captured_loose(
            entries,
            DirectoryMarkers::default(),
            Limits::default(),
        )?;
        assert!(frozen.components().get("Index/Document.iwa").is_some());

        fs::rename(&index, temp.path().join("OriginalIndex"))?;
        fs::create_dir(&index)?;
        fs::write(index.join("Document.iwa"), iwa(2, 6_001)?)?;
        assert_eq!(unix_scan_manifest(&index_fd, Limits::default())?, manifest);
        assert!(matches!(
            unix_ensure_node_version(
                &root_fd,
                c"Index",
                &version,
                "test final pinned Index verification"
            ),
            Err(Error::DirectoryChanged { .. })
        ));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn preexisting_ancestor_symlink_is_resolved_before_root_is_pinned() -> Result<()> {
        use std::os::unix::fs::symlink;

        let parent = Temp::new()?;
        let actual = parent.path().join("actual");
        fs::create_dir(&actual)?;
        let source = actual.join("document.pages");
        fs::create_dir(&source)?;
        fs::create_dir(source.join("Index"))?;
        fs::write(source.join("Index/Document.iwa"), iwa(1, 10_000)?)?;
        let ancestor = parent.path().join("ancestor-link");
        symlink(&actual, &ancestor)?;

        let frozen = FrozenDirectoryBundle::open(ancestor.join("document.pages"))?;
        assert_eq!(frozen.components().len(), 1);
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
