//! Mutable iWork ZIP package with entry-order preservation.

use std::collections::HashSet;
use std::fs::{self, File};
use std::io::Write;
use std::path::{Component, Path};
use std::sync::Arc;

use soapberry_zip::office::{ArchiveLimits, ArchiveReader, StreamingArchiveWriter};
use tempfile::NamedTempFile;

use crate::archive::Archive;
use crate::snappy::{SnappyLimits, SnappyStream};
use crate::zip_utils::{is_encrypted_iwork_archive, nested_index_zip_name};
use crate::{Error, Result};

#[path = "package_state.rs"]
mod package_state;
use package_state::PackageState;

/// A mutable single-file Pages, Numbers, or Keynote package.
///
/// All ZIP members are retained as raw uncompressed bytes. IWA entries can be
/// parsed, updated transactionally, and written back while media, previews, and
/// metadata remain byte-for-byte unchanged.
///
/// Cloning a package is cheap: the entry table is shared until a mutation is
/// attempted. This keeps editor staging paths memory-efficient when a
/// transaction is rejected and discarded. The first mutation of a shared
/// package performs the required copy-on-write detachment.
#[derive(Debug, Clone, Default)]
pub struct IWorkPackage {
    state: Arc<PackageState>,
    limits: PackageLimits,
}

/// An immutable, cheaply shareable package snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    state: Arc<PackageState>,
    limits: PackageLimits,
}

/// Resource ceilings applied while ingesting one iWork package.
///
/// The limits cover every ZIP archive opened during ingress, including a
/// legacy nested `Index.zip`. They bound central-directory metadata before
/// package members are materialized and can only be tightened below the
/// format-wide hard ceilings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PackageLimits {
    max_input_bytes: u64,
    max_entries: usize,
    max_entry_bytes: u64,
    max_total_bytes: u64,
    max_iwa_stream_bytes: usize,
}

impl PackageLimits {
    /// Hard ceiling for bytes read from one package path or byte slice.
    pub const MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;
    /// Hard ceiling for non-directory package members in one ZIP archive.
    pub const MAX_ENTRIES: usize = 100_000;
    /// Hard ceiling for one declared uncompressed package member.
    pub const MAX_ENTRY_BYTES: u64 = 512 * 1024 * 1024;
    /// Hard ceiling for the declared uncompressed size of one ZIP archive.
    pub const MAX_TOTAL_BYTES: u64 = 2 * 1024 * 1024 * 1024;
    /// Hard ceiling for one decompressed IWA component.
    pub const MAX_IWA_STREAM_BYTES: usize = SnappyStream::MAX_DECOMPRESSED_STREAM;

    /// Build checked package-ingress ceilings.
    pub fn new(max_entries: usize, max_entry_bytes: u64, max_total_bytes: u64) -> Result<Self> {
        Self::new_with_limits(
            Self::MAX_INPUT_BYTES,
            max_entries,
            max_entry_bytes,
            max_total_bytes,
            Self::MAX_IWA_STREAM_BYTES,
        )
    }

    /// Build checked package-ingress ceilings, including filesystem and IWA
    /// decompression budgets.
    pub fn new_with_limits(
        max_input_bytes: u64,
        max_entries: usize,
        max_entry_bytes: u64,
        max_total_bytes: u64,
        max_iwa_stream_bytes: usize,
    ) -> Result<Self> {
        let limits = Self {
            max_input_bytes,
            max_entries,
            max_entry_bytes,
            max_total_bytes,
            max_iwa_stream_bytes,
        };
        limits.validate()
    }

    fn validate(self) -> Result<Self> {
        if self.max_input_bytes == 0
            || self.max_entries == 0
            || self.max_entry_bytes == 0
            || self.max_total_bytes == 0
            || self.max_iwa_stream_bytes == 0
        {
            return Err(Error::InvalidFormat(
                "iWork package limits must be non-zero".to_owned(),
            ));
        }
        if self.max_input_bytes > Self::MAX_INPUT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "iWork package input limit exceeds the {} byte hard ceiling",
                Self::MAX_INPUT_BYTES
            )));
        }
        if self.max_entries > Self::MAX_ENTRIES {
            return Err(Error::InvalidFormat(format!(
                "iWork package entry limit exceeds the {} entry hard ceiling",
                Self::MAX_ENTRIES
            )));
        }
        if self.max_entry_bytes > Self::MAX_ENTRY_BYTES {
            return Err(Error::InvalidFormat(format!(
                "iWork package entry limit exceeds the {} byte hard ceiling",
                Self::MAX_ENTRY_BYTES
            )));
        }
        if self.max_total_bytes > Self::MAX_TOTAL_BYTES {
            return Err(Error::InvalidFormat(format!(
                "iWork package total limit exceeds the {} byte hard ceiling",
                Self::MAX_TOTAL_BYTES
            )));
        }
        if self.max_iwa_stream_bytes > Self::MAX_IWA_STREAM_BYTES {
            return Err(Error::InvalidFormat(format!(
                "iWork package IWA stream limit exceeds the {} byte hard ceiling",
                Self::MAX_IWA_STREAM_BYTES
            )));
        }
        Ok(self)
    }

    /// Tighten the maximum bytes read from one package path or byte slice.
    pub fn with_input_bytes(mut self, max_input_bytes: u64) -> Result<Self> {
        self.max_input_bytes = max_input_bytes;
        self.validate()
    }

    /// Tighten the maximum decompressed size of one IWA component.
    pub fn with_iwa_stream_bytes(mut self, max_iwa_stream_bytes: usize) -> Result<Self> {
        self.max_iwa_stream_bytes = max_iwa_stream_bytes;
        self.validate()
    }

    /// Maximum bytes read from one package path or byte slice.
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum number of non-directory members accepted in one archive.
    pub const fn max_entries(self) -> usize {
        self.max_entries
    }

    /// Maximum declared uncompressed size of one member.
    pub const fn max_entry_bytes(self) -> u64 {
        self.max_entry_bytes
    }

    /// Maximum declared uncompressed size of one archive.
    pub const fn max_total_bytes(self) -> u64 {
        self.max_total_bytes
    }

    /// Maximum decompressed size of one IWA component.
    pub const fn max_iwa_stream_bytes(self) -> usize {
        self.max_iwa_stream_bytes
    }

    fn archive_limits(self) -> ArchiveLimits {
        ArchiveLimits {
            max_files: self.max_entries,
            max_entry_size: self.max_entry_bytes,
            max_total_size: self.max_total_bytes,
        }
    }

    fn snappy_limits(self) -> Result<SnappyLimits> {
        SnappyLimits::new(
            self.max_iwa_stream_bytes
                .min(SnappyStream::MAX_UNCOMPRESSED_CHUNK),
            self.max_iwa_stream_bytes,
        )
        .map_err(|error| Error::Snappy(error.to_string()))
    }

    fn check_input_size(self, size: u64, label: &str) -> Result<()> {
        if size > self.max_input_bytes {
            return Err(Error::InvalidFormat(format!(
                "{label} is {size} bytes, exceeding the {} byte limit",
                self.max_input_bytes
            )));
        }
        Ok(())
    }

    fn check_iwa_stream_size(self, size: usize) -> Result<()> {
        if size > self.max_iwa_stream_bytes {
            return Err(Error::Snappy(format!(
                "IWA stream is {size} bytes, exceeding the {} byte limit",
                self.max_iwa_stream_bytes
            )));
        }
        Ok(())
    }
}

impl Default for PackageLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: Self::MAX_INPUT_BYTES,
            max_entries: Self::MAX_ENTRIES,
            max_entry_bytes: Self::MAX_ENTRY_BYTES,
            max_total_bytes: Self::MAX_TOTAL_BYTES,
            max_iwa_stream_bytes: Self::MAX_IWA_STREAM_BYTES,
        }
    }
}

impl IWorkPackage {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, PackageLimits::default())
    }

    /// Open a package from a path under caller-selected ingress ceilings.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: PackageLimits) -> Result<Self> {
        let path = path.as_ref();
        let size = std::fs::metadata(path)?.len();
        limits.check_input_size(size, "iWork package input")?;
        Self::from_bytes_with_limits(&std::fs::read(path)?, limits)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, PackageLimits::default())
    }

    /// Parse a package under caller-selected ingress ceilings.
    ///
    /// The same ceilings are applied to the outer ZIP and, when present, the
    /// embedded legacy `Index.zip` archive.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: PackageLimits) -> Result<Self> {
        let input_size = u64::try_from(bytes.len()).map_err(|_| {
            Error::InvalidFormat("iWork package input length does not fit u64".to_owned())
        })?;
        limits.check_input_size(input_size, "iWork package input")?;
        let archive = ArchiveReader::new_with_limits(bytes, limits.archive_limits())
            .map_err(|error| Error::Bundle(format!("Failed to open iWork ZIP: {error}")))?;
        if is_encrypted_iwork_archive(&archive) {
            return Err(Error::InvalidFormat(
                "password-protected iWork documents are not supported".to_owned(),
            ));
        }
        if !archive.file_names().any(|name| name.ends_with(".iwa"))
            && let Some(index_name) = nested_index_zip_name(&archive)?
        {
            return Self::from_legacy_bundle(&archive, &index_name, limits);
        }
        Self::from_flat_archive(&archive, limits)
    }

    fn from_flat_archive(archive: &ArchiveReader<'_>, limits: PackageLimits) -> Result<Self> {
        let mut entries = Vec::new();
        let mut seen = HashSet::new();
        for name in archive.file_names() {
            validate_entry_name(name)?;
            if !seen.insert(name.to_owned()) {
                return Err(Error::Bundle(format!(
                    "Duplicate package entry is ambiguous: {name}"
                )));
            }
            let data = archive.read(name).map_err(|error| {
                Error::Bundle(format!("Failed to read package entry {name}: {error}"))
            })?;
            entries.push((name.to_owned(), data));
        }
        Ok(Self {
            state: Arc::new(PackageState::from_entries(entries)),
            limits,
        })
    }

    /// Expand the pre-iWork '13 nested bundle representation into the modern,
    /// flat package representation used by the rest of the mutable API. The
    /// IWA members come first and all non-directory assets are retained with
    /// the legacy bundle prefix removed.
    fn from_legacy_bundle(
        archive: &ArchiveReader<'_>,
        index_name: &str,
        limits: PackageLimits,
    ) -> Result<Self> {
        let prefix = index_name.strip_suffix("Index.zip").ok_or_else(|| {
            Error::InvalidFormat(format!("invalid legacy package index name: {index_name}"))
        })?;
        let index_data = archive.read(index_name).map_err(|error| {
            Error::Bundle(format!(
                "Failed to read legacy package index {index_name}: {error}"
            ))
        })?;
        let index_size = u64::try_from(index_data.len()).map_err(|_| {
            Error::InvalidFormat("legacy package index length does not fit u64".to_owned())
        })?;
        limits.check_input_size(index_size, "legacy iWork Index.zip")?;
        let index = ArchiveReader::new_with_limits(&index_data, limits.archive_limits()).map_err(
            |error| {
                Error::Bundle(format!(
                    "Failed to open legacy package index {index_name}: {error}"
                ))
            },
        )?;

        let mut entries = Vec::new();
        let mut seen = HashSet::new();
        for name in index.file_names().filter(|name| !name.ends_with('/')) {
            validate_entry_name(name)?;
            if !name.ends_with(".iwa") {
                return Err(Error::InvalidFormat(format!(
                    "legacy package index contains a non-IWA member: {name}"
                )));
            }
            insert_unique_archive_entry(&index, name, &mut entries, &mut seen)?;
        }
        if entries.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "legacy package index {index_name} contains no IWA components"
            )));
        }

        for outer_name in archive
            .file_names()
            .filter(|name| *name != index_name && !name.ends_with('/'))
        {
            let name = outer_name.strip_prefix(prefix).unwrap_or(outer_name);
            validate_entry_name(name)?;
            if !seen.insert(name.to_owned()) {
                return Err(Error::InvalidFormat(format!(
                    "legacy package entries normalize to the same name: {name}"
                )));
            }
            let data = archive.read(outer_name).map_err(|error| {
                Error::Bundle(format!(
                    "Failed to read legacy package entry {outer_name}: {error}"
                ))
            })?;
            entries.push((name.to_owned(), data));
        }
        Ok(Self {
            state: Arc::new(PackageState::from_entries(entries)),
            limits,
        })
    }

    pub fn len(&self) -> usize {
        self.state.entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.state.entries.is_empty()
    }

    /// Return the resource profile retained for lazy archive reads and edits.
    pub const fn limits(&self) -> PackageLimits {
        self.limits
    }

    /// Capture an immutable package snapshot without copying package entries.
    pub fn snapshot(&self) -> Snapshot {
        Snapshot {
            state: Arc::clone(&self.state),
            limits: self.limits,
        }
    }

    pub fn entry_names(&self) -> impl Iterator<Item = &str> {
        self.state.entries.iter().map(|(name, _)| name.as_str())
    }

    /// Enumerate package members that contain IWA object archives.
    ///
    /// Some legacy packages contain an `OperationStorage.iwa` member whose
    /// `bvxn` payload is a separate operation-log format. It is intentionally
    /// retained as a raw entry but excluded from object-archive scans.
    pub fn iwa_entry_names(&self) -> impl Iterator<Item = &str> {
        self.state
            .entries
            .iter()
            .filter(|(name, data)| {
                name.ends_with(".iwa") && !is_legacy_operation_storage(name, data)
            })
            .map(|(name, _)| name.as_str())
    }

    /// Locate the package's calculation-engine component without allocating.
    ///
    /// Pages and Numbers may add numeric suffixes when they save a package,
    /// for example `Index/CalculationEngine-174.iwa`. Multiple matching
    /// components are rejected because choosing one would make formula edits
    /// ambiguous.
    pub fn calculation_engine_entry_name(&self) -> Result<Option<&str>> {
        let mut entries = self
            .iwa_entry_names()
            .filter(|name| is_calculation_engine_entry_name(name));
        let Some(entry) = entries.next() else {
            return Ok(None);
        };
        if entries.next().is_some() {
            return Err(Error::InvalidFormat(
                "iWork package contains multiple CalculationEngine components".to_owned(),
            ));
        }
        Ok(Some(entry))
    }

    pub fn contains_entry(&self, name: &str) -> bool {
        self.entry_position(normalize_entry_name(name)).is_some()
    }

    pub fn entry(&self, name: &str) -> Option<&[u8]> {
        let position = self.entry_position(normalize_entry_name(name))?;
        Some(self.state.entries[position].1.as_slice())
    }

    /// Borrow a raw package member for low-level mutation.
    ///
    /// The returned vector is intentionally an escape hatch for format-specific
    /// editors. Package entry, aggregate, and ZIP safety budgets are rechecked
    /// before serialization, so an oversized direct mutation cannot be
    /// published accidentally.
    pub fn entry_mut(&mut self, name: &str) -> Option<&mut Vec<u8>> {
        let position = self.entry_position(normalize_entry_name(name))?;
        Some(&mut Arc::make_mut(&mut self.state).entries[position].1)
    }

    /// Create or replace a package member.
    pub fn insert_entry(
        &mut self,
        name: impl Into<String>,
        data: Vec<u8>,
    ) -> Result<Option<Vec<u8>>> {
        let supplied_name = name.into();
        let name = normalize_entry_name(&supplied_name).to_string();
        validate_entry_name(&name)?;
        let position = self.entry_position(&name);
        self.validate_entry_update(position, &data)?;
        if let Some(position) = position {
            return Ok(Some(std::mem::replace(
                &mut Arc::make_mut(&mut self.state).entries[position].1,
                data,
            )));
        }
        self.insert_new_entry(name, data);
        Ok(None)
    }

    /// Delete a package member.
    pub fn remove_entry(&mut self, name: &str) -> Option<Vec<u8>> {
        let position = self.entry_position(normalize_entry_name(name))?;
        let state = Arc::make_mut(&mut self.state);
        let removed = state.entries.remove(position).1;
        state.rebuild_positions();
        Some(removed)
    }

    /// Parse a compressed `.iwa` package member.
    pub fn archive(&self, name: &str) -> Result<Archive> {
        let normalized = normalize_entry_name(name);
        if !normalized.ends_with(".iwa") {
            return Err(Error::Bundle(format!(
                "Package entry {normalized} is not an IWA component"
            )));
        }
        let compressed = self
            .entry(normalized)
            .ok_or_else(|| Error::Bundle(format!("IWA package entry not found: {normalized}")))?;
        if is_legacy_operation_storage(normalized, compressed) {
            return Err(Error::InvalidFormat(format!(
                "package entry {normalized} is a legacy operation log, not an IWA object archive"
            )));
        }
        let stream = SnappyStream::decompress_with_limits(
            &mut std::io::Cursor::new(compressed),
            self.limits.snappy_limits()?,
        )?;
        Archive::parse(stream.data())
    }

    /// Serialize and replace a parsed `.iwa` package member.
    pub fn replace_archive(&mut self, name: &str, archive: &Archive) -> Result<Option<Vec<u8>>> {
        let normalized = normalize_entry_name(name).to_string();
        validate_entry_name(&normalized)?;
        if !normalized.ends_with(".iwa") {
            return Err(Error::Bundle(format!(
                "Package entry {normalized} is not an IWA component"
            )));
        }
        let data = archive.to_bytes()?;
        self.limits.check_iwa_stream_size(data.len())?;
        let compressed = SnappyStream::compress(&data)?;
        self.insert_entry(normalized, compressed)
    }

    /// Serialize and insert a new IWA component before an existing package member.
    pub(crate) fn insert_archive_before(
        &mut self,
        name: &str,
        archive: &Archive,
        before: &str,
    ) -> Result<()> {
        let normalized = normalize_entry_name(name).to_string();
        let before = normalize_entry_name(before);
        validate_entry_name(&normalized)?;
        if !normalized.ends_with(".iwa") {
            return Err(Error::Bundle(format!(
                "Package entry {normalized} is not an IWA component"
            )));
        }
        if self.contains_entry(&normalized) {
            return Err(Error::InvalidFormat(format!(
                "IWA package entry already exists: {normalized}"
            )));
        }
        let position = self
            .entry_position(before)
            .ok_or_else(|| Error::Bundle(format!("IWA insertion anchor not found: {before}")))?;
        let data = archive.to_bytes()?;
        self.limits.check_iwa_stream_size(data.len())?;
        let compressed = SnappyStream::compress(&data)?;
        self.validate_entry_update(None, &compressed)?;
        let state = Arc::make_mut(&mut self.state);
        state.entries.insert(position, (normalized, compressed));
        state.rebuild_positions();
        Ok(())
    }

    /// Parse, mutate, validate, and replace an IWA component as one operation.
    /// If the callback or serialization fails, the original package is unchanged.
    pub fn update_archive<F>(&mut self, name: &str, update: F) -> Result<()>
    where
        F: FnOnce(&mut Archive) -> Result<()>,
    {
        let mut archive = self.archive(name)?;
        update(&mut archive)?;
        archive.validate()?;
        self.replace_archive(name, &archive)?;
        Ok(())
    }

    /// Encode the package as a ZIP using stored members and the original order.
    ///
    /// Pages and Numbers use a leading `Index/Document.iwa` for package type
    /// discovery, so newly-created document indexes are inserted first.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.validate_state()?;
        let mut writer = StreamingArchiveWriter::new();
        self.write_entries(&mut writer)?;
        writer
            .finish_to_bytes()
            .map_err(|error| Error::Bundle(format!("Failed to finish iWork ZIP: {error}")))
    }

    /// Stream the package as a ZIP to a caller-owned sequential sink.
    ///
    /// The package is validated before the first byte is written. ZIP central
    /// directory records are finalized by the sink writer, so this method does
    /// not allocate a second buffer containing the complete package.
    pub fn write_to<W: Write>(&self, sink: W) -> Result<()> {
        self.validate_state()?;
        let mut writer = StreamingArchiveWriter::with_writer(sink);
        self.write_entries(&mut writer)?;
        writer
            .finish()
            .map(|_| ())
            .map_err(|error| Error::Bundle(format!("Failed to finish iWork ZIP: {error}")))
    }

    /// Atomically save the package to a regular file in the destination
    /// directory without buffering the complete ZIP in memory.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let path = path.as_ref();
        let parent = path
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));

        let existing = match fs::symlink_metadata(path) {
            Ok(metadata) if metadata.file_type().is_symlink() => {
                return Err(Error::Bundle(format!(
                    "iWork package destination must not be a symbolic link: {}",
                    path.display()
                )));
            },
            Ok(metadata) if metadata.is_file() => Some(metadata),
            Ok(_) => {
                return Err(Error::Bundle(format!(
                    "iWork package destination is not a regular file: {}",
                    path.display()
                )));
            },
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => None,
            Err(error) => return Err(error.into()),
        };

        let mut temporary = NamedTempFile::new_in(parent)?;
        self.write_to(temporary.as_file_mut())?;

        // Apply the existing mode after writing: a read-only destination must
        // not make the temporary file unwritable while the ZIP is finalized.
        if let Some(metadata) = existing {
            fs::set_permissions(temporary.path(), metadata.permissions())?;
        }
        temporary.as_file_mut().sync_all()?;
        temporary
            .persist(path)
            .map_err(|error| Error::Io(error.error))?;

        // Make the rename durable on filesystems that support directory sync.
        if let Ok(directory) = File::open(parent) {
            directory.sync_all()?;
        }
        Ok(())
    }

    fn write_entries<W: Write>(&self, writer: &mut StreamingArchiveWriter<W>) -> Result<()> {
        for (name, data) in &self.state.entries {
            writer.write_stored(name, data).map_err(|error| {
                Error::Bundle(format!("Failed to write package entry {name}: {error}"))
            })?;
        }
        Ok(())
    }

    fn entry_position(&self, name: &str) -> Option<usize> {
        self.state.position(name)
    }

    fn validate_entry_update(&self, position: Option<usize>, data: &[u8]) -> Result<()> {
        let current_total = self.validate_state()?;
        self.validate_entry_data(data)?;
        if position.is_none() && self.state.entries.len() >= self.limits.max_entries {
            return Err(Error::Bundle(format!(
                "iWork package entry count exceeds the {} entry limit",
                self.limits.max_entries
            )));
        }

        let previous_size = position
            .map(|index| self.state.entries[index].1.len())
            .unwrap_or(0);
        let previous_size = u64::try_from(previous_size)
            .map_err(|_| Error::Bundle("package member length does not fit u64".to_owned()))?;
        let data_size = u64::try_from(data.len())
            .map_err(|_| Error::Bundle("package member length does not fit u64".to_owned()))?;
        let projected_total = current_total
            .checked_sub(previous_size)
            .and_then(|total| total.checked_add(data_size))
            .ok_or_else(|| Error::Bundle("iWork package total size overflow".to_owned()))?;
        if projected_total > self.limits.max_total_bytes {
            return Err(Error::Bundle(format!(
                "iWork package total size exceeds the {} byte limit",
                self.limits.max_total_bytes
            )));
        }
        Ok(())
    }

    fn validate_state(&self) -> Result<u64> {
        if self.state.entries.len() > self.limits.max_entries {
            return Err(Error::Bundle(format!(
                "iWork package entry count exceeds the {} entry limit",
                self.limits.max_entries
            )));
        }
        for (name, data) in &self.state.entries {
            validate_entry_name(name)?;
            self.validate_entry_data(data)?;
        }
        let total = self
            .state
            .entries
            .iter()
            .try_fold(0_u64, |total, (_, data)| {
                let size = u64::try_from(data.len()).map_err(|_| {
                    Error::Bundle("package member length does not fit u64".to_owned())
                })?;
                total
                    .checked_add(size)
                    .ok_or_else(|| Error::Bundle("iWork package total size overflow".to_owned()))
            })?;
        if total > self.limits.max_total_bytes {
            return Err(Error::Bundle(format!(
                "iWork package total size exceeds the {} byte limit",
                self.limits.max_total_bytes
            )));
        }
        Ok(total)
    }

    fn validate_entry_data(&self, data: &[u8]) -> Result<()> {
        let size = u64::try_from(data.len())
            .map_err(|_| Error::Bundle("package member length does not fit u64".to_owned()))?;
        if size > self.limits.max_entry_bytes {
            return Err(Error::Bundle(format!(
                "iWork package member is {size} bytes, exceeding the {} byte limit",
                self.limits.max_entry_bytes
            )));
        }
        Ok(())
    }

    fn insert_new_entry(&mut self, name: String, data: Vec<u8>) {
        let state = Arc::make_mut(&mut self.state);
        if name == "Index/Document.iwa" {
            state.entries.insert(0, (name, data));
        } else {
            state.entries.push((name, data));
        }
        state.rebuild_positions();
    }
}

impl Snapshot {
    /// Return the number of retained package members.
    pub fn len(&self) -> usize {
        self.state.entries.len()
    }

    /// Report whether the package contains no members.
    pub fn is_empty(&self) -> bool {
        self.state.entries.is_empty()
    }

    /// Return the resource profile retained by this immutable snapshot.
    pub const fn limits(&self) -> PackageLimits {
        self.limits
    }

    /// Enumerate package members in preserved source order.
    pub fn entry_names(&self) -> impl Iterator<Item = &str> {
        self.state.entries.iter().map(|(name, _)| name.as_str())
    }

    /// Borrow one package member without copying it.
    pub fn entry(&self, name: &str) -> Option<&[u8]> {
        let name = normalize_entry_name(name);
        let position = self.state.position(name)?;
        Some(self.state.entries[position].1.as_slice())
    }

    /// Start a mutable copy-on-write edit from this snapshot.
    pub fn edit(&self) -> IWorkPackage {
        IWorkPackage {
            state: Arc::clone(&self.state),
            limits: self.limits,
        }
    }

    /// Encode the unchanged snapshot as an iWork ZIP package.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.edit().to_bytes()
    }

    /// Stream the unchanged snapshot as a ZIP to a caller-owned sink.
    pub fn write_to<W: Write>(&self, sink: W) -> Result<()> {
        self.edit().write_to(sink)
    }

    /// Atomically save the unchanged snapshot to a regular file.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        self.edit().save(path)
    }

    /// Report whether a normalized package member exists without creating an
    /// editable package view.
    pub fn contains_entry(&self, name: &str) -> bool {
        self.state.position(normalize_entry_name(name)).is_some()
    }

    /// Enumerate object-archive members in preserved source order.
    pub fn iwa_entry_names(&self) -> impl Iterator<Item = &str> {
        self.state
            .entries
            .iter()
            .filter(|(name, data)| {
                name.ends_with(".iwa") && !is_legacy_operation_storage(name, data)
            })
            .map(|(name, _)| name.as_str())
    }

    /// Locate the package's calculation-engine component without first
    /// materializing an editable package.
    pub fn calculation_engine_entry_name(&self) -> Result<Option<&str>> {
        let mut entries = self
            .iwa_entry_names()
            .filter(|name| is_calculation_engine_entry_name(name));
        let Some(entry) = entries.next() else {
            return Ok(None);
        };
        if entries.next().is_some() {
            return Err(Error::InvalidFormat(
                "iWork package contains multiple CalculationEngine components".to_owned(),
            ));
        }
        Ok(Some(entry))
    }

    /// Parse a compressed `.iwa` member from this immutable snapshot.
    pub fn archive(&self, name: &str) -> Result<Archive> {
        self.edit().archive(name)
    }
}

fn insert_unique_archive_entry(
    archive: &ArchiveReader<'_>,
    name: &str,
    entries: &mut Vec<(String, Vec<u8>)>,
    seen: &mut HashSet<String>,
) -> Result<()> {
    if !seen.insert(name.to_owned()) {
        return Err(Error::Bundle(format!(
            "Duplicate package entry is ambiguous: {name}"
        )));
    }
    let data = archive
        .read(name)
        .map_err(|error| Error::Bundle(format!("Failed to read package entry {name}: {error}")))?;
    entries.push((name.to_owned(), data));
    Ok(())
}

fn normalize_entry_name(name: &str) -> &str {
    name.strip_prefix('/').unwrap_or(name)
}

pub(crate) fn is_calculation_engine_entry_name(name: &str) -> bool {
    const BASE_NAME: &str = "CalculationEngine.iwa";
    const VERSIONED_PREFIX: &str = "CalculationEngine-";

    name.rsplit('/').next().is_some_and(|file_name| {
        file_name == BASE_NAME
            || file_name
                .strip_prefix(VERSIONED_PREFIX)
                .and_then(|suffix| suffix.strip_suffix(".iwa"))
                .is_some_and(|version| {
                    !version.is_empty()
                        && version.split('-').all(|part| {
                            !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit())
                        })
                })
    })
}

fn is_legacy_operation_storage(name: &str, data: &[u8]) -> bool {
    name.rsplit('/').next() == Some("OperationStorage.iwa") && data.starts_with(b"bvxn")
}

fn validate_entry_name(name: &str) -> Result<()> {
    if name.is_empty() || name.contains('\0') || name.contains('\\') {
        return Err(Error::Bundle(format!(
            "Invalid package entry name: {name:?}"
        )));
    }
    let path = Path::new(name);
    if path.is_absolute()
        || path
            .components()
            .any(|component| !matches!(component, Component::Normal(_)))
    {
        return Err(Error::Bundle(format!("Unsafe package entry name: {name}")));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{ArchiveObject, RawMessage};

    fn archive() -> Archive {
        Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 99,
                        data: vec![1, 2, 3],
                    }],
                )
                .unwrap(),
            ],
        }
    }

    fn legacy_package() -> Vec<u8> {
        let compressed = SnappyStream::compress(&archive().to_bytes().unwrap()).unwrap();
        let mut index = StreamingArchiveWriter::new();
        index
            .write_stored("Index/Document.iwa", &compressed)
            .unwrap();
        let index = index.finish_to_bytes().unwrap();

        let mut outer = StreamingArchiveWriter::new();
        outer
            .write_stored("mac.numbers/preview.jpg", b"preview")
            .unwrap();
        outer.write_stored("mac.numbers/Index.zip", &index).unwrap();
        outer
            .write_stored("mac.numbers/Metadata/Properties.plist", b"plist")
            .unwrap();
        outer.finish_to_bytes().unwrap()
    }

    fn empty_package_with_limits(limits: PackageLimits) -> IWorkPackage {
        let bytes = StreamingArchiveWriter::new().finish_to_bytes().unwrap();
        IWorkPackage::from_bytes_with_limits(&bytes, limits).unwrap()
    }

    #[test]
    fn package_limits_are_checked_and_exposed() {
        let limits = PackageLimits::new(7, 11, 23).unwrap();
        assert_eq!(limits.max_input_bytes(), PackageLimits::MAX_INPUT_BYTES);
        assert_eq!(limits.max_entries(), 7);
        assert_eq!(limits.max_entry_bytes(), 11);
        assert_eq!(limits.max_total_bytes(), 23);
        assert_eq!(
            limits.max_iwa_stream_bytes(),
            PackageLimits::MAX_IWA_STREAM_BYTES
        );

        assert!(PackageLimits::new(0, 1, 1).is_err());
        assert!(PackageLimits::new(1, 0, 1).is_err());
        assert!(PackageLimits::new(1, 1, 0).is_err());
        assert!(
            PackageLimits::new_with_limits(
                0,
                PackageLimits::MAX_ENTRIES,
                PackageLimits::MAX_ENTRY_BYTES,
                PackageLimits::MAX_TOTAL_BYTES,
                PackageLimits::MAX_IWA_STREAM_BYTES,
            )
            .is_err()
        );
        assert!(
            PackageLimits::new_with_limits(PackageLimits::MAX_INPUT_BYTES + 1, 1, 1, 1, 1,)
                .is_err()
        );
        assert!(
            PackageLimits::new_with_limits(
                PackageLimits::MAX_INPUT_BYTES,
                1,
                1,
                1,
                PackageLimits::MAX_IWA_STREAM_BYTES + 1,
            )
            .is_err()
        );
        assert!(PackageLimits::new(PackageLimits::MAX_ENTRIES + 1, 1, 1).is_err());
        assert!(PackageLimits::new(1, PackageLimits::MAX_ENTRY_BYTES + 1, 1).is_err());
        assert!(PackageLimits::new(1, 1, PackageLimits::MAX_TOTAL_BYTES + 1).is_err());

        let tightened = limits
            .with_input_bytes(31)
            .unwrap()
            .with_iwa_stream_bytes(47)
            .unwrap();
        assert_eq!(tightened.max_input_bytes(), 31);
        assert_eq!(tightened.max_iwa_stream_bytes(), 47);
        assert!(limits.with_input_bytes(0).is_err());
        assert!(limits.with_iwa_stream_bytes(0).is_err());
    }

    #[test]
    fn package_limits_bound_file_input_and_lazy_iwa_decompression() {
        let decompressed = archive().to_bytes().unwrap();
        assert!(decompressed.len() > 8);
        let compressed = SnappyStream::compress(&decompressed).unwrap();
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("Index/Document.iwa", &compressed)
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let input_limits = PackageLimits::new(
            10,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
        )
        .unwrap()
        .with_input_bytes(u64::try_from(bytes.len() - 1).unwrap())
        .unwrap();
        let error = IWorkPackage::from_bytes_with_limits(&bytes, input_limits).unwrap_err();
        assert!(error.to_string().contains("iWork package input"));

        let stream_limits = PackageLimits::new(
            10,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
        )
        .unwrap()
        .with_iwa_stream_bytes(8)
        .unwrap();
        let mut package = IWorkPackage::from_bytes_with_limits(&bytes, stream_limits).unwrap();
        assert_eq!(package.limits(), stream_limits);
        assert_eq!(package.snapshot().limits(), stream_limits);
        let error = package.archive("Index/Document.iwa").unwrap_err();
        assert!(error.to_string().contains("Snappy block expands"));
        let error = package
            .replace_archive("Index/Other.iwa", &archive())
            .unwrap_err();
        assert!(error.to_string().contains("IWA stream is"));

        let file = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(file.path(), &bytes).unwrap();
        let error = IWorkPackage::open_with_limits(file.path(), input_limits).unwrap_err();
        assert!(error.to_string().contains("iWork package input"));
    }

    #[test]
    fn package_mutations_respect_entry_and_aggregate_budgets() {
        let entry_limits = PackageLimits::new(1, 4, 4).unwrap();
        let mut package = empty_package_with_limits(entry_limits);
        let error = package
            .insert_entry("Data/too-large", vec![1, 2, 3, 4, 5])
            .unwrap_err();
        assert!(error.to_string().contains("member"));
        assert!(package.is_empty());

        package
            .insert_entry("Data/asset", vec![1, 2, 3, 4])
            .unwrap();
        let error = package.insert_entry("Data/second", vec![5]).unwrap_err();
        assert!(error.to_string().contains("entry count"));
        assert_eq!(package.len(), 1);

        let total_limits = PackageLimits::new(10, 4, 4).unwrap();
        let mut package = empty_package_with_limits(total_limits);
        package
            .insert_entry("Data/asset", vec![1, 2, 3, 4])
            .unwrap();
        let error = package.insert_entry("Data/second", vec![5]).unwrap_err();
        assert!(error.to_string().contains("total size"));

        package.entry_mut("Data/asset").unwrap().push(5);
        let error = package.to_bytes().unwrap_err();
        assert!(error.to_string().contains("member"));
    }

    #[test]
    fn package_limits_reject_flat_archive_metadata_before_materialization() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("Data/asset", b"asset").unwrap();
        writer.write_stored("Metadata/other", b"other").unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let one_entry = PackageLimits::new(
            1,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
        )
        .unwrap();
        assert!(IWorkPackage::from_bytes_with_limits(&bytes, one_entry).is_err());

        let small_entry = PackageLimits::new(10, 4, PackageLimits::MAX_TOTAL_BYTES).unwrap();
        assert!(IWorkPackage::from_bytes_with_limits(&bytes, small_entry).is_err());
    }

    #[test]
    fn package_limits_apply_to_legacy_nested_index_zip() {
        let compressed = SnappyStream::compress(&archive().to_bytes().unwrap()).unwrap();
        let mut index = StreamingArchiveWriter::new();
        index
            .write_stored("Index/Document.iwa", &compressed)
            .unwrap();
        index
            .write_stored("Index/Metadata.iwa", &compressed)
            .unwrap();
        let index = index.finish_to_bytes().unwrap();

        let mut outer = StreamingArchiveWriter::new();
        outer.write_stored("mac.pages/Index.zip", &index).unwrap();
        let bytes = outer.finish_to_bytes().unwrap();

        let one_entry = PackageLimits::new(
            1,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
        )
        .unwrap();
        let error = IWorkPackage::from_bytes_with_limits(&bytes, one_entry).unwrap_err();
        assert!(error.to_string().contains("legacy package index"));
    }

    #[test]
    fn package_entry_and_archive_crud_round_trip() {
        let mut package = IWorkPackage::new();
        package
            .insert_entry("Metadata/Properties.plist", b"plist".to_vec())
            .unwrap();
        package
            .replace_archive("Index/Document.iwa", &archive())
            .unwrap();

        package
            .update_archive("Index/Document.iwa", |archive| {
                archive.object_mut(1).unwrap().replace_message(
                    0,
                    RawMessage {
                        type_: 100,
                        data: vec![4, 5],
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let bytes = package.to_bytes().unwrap();
        let mut reparsed = IWorkPackage::from_bytes(&bytes).unwrap();
        let document = reparsed.archive("Index/Document.iwa").unwrap();
        assert_eq!(document.object(1).unwrap().messages[0].type_, 100);
        assert_eq!(document.object(1).unwrap().messages[0].data, [4, 5]);
        assert_eq!(
            reparsed.remove_entry("Metadata/Properties.plist"),
            Some(b"plist".to_vec())
        );
        assert_eq!(reparsed.entry_names().next(), Some("Index/Document.iwa"));
    }

    #[test]
    fn snapshots_share_storage_until_the_first_mutation() {
        let mut package = IWorkPackage::new();
        package
            .insert_entry("Data/original", b"original".to_vec())
            .unwrap();

        let snapshot = package.snapshot();
        let mut edit = snapshot.edit();
        assert_eq!(snapshot.limits(), package.limits());
        assert_eq!(edit.limits(), package.limits());
        assert_eq!(
            snapshot.entry("Data/original"),
            Some(b"original".as_slice())
        );
        assert_eq!(edit.entry("Data/original"), Some(b"original".as_slice()));

        edit.insert_entry("Data/changed", b"changed".to_vec())
            .unwrap();
        assert_eq!(snapshot.len(), 1);
        assert_eq!(snapshot.entry("Data/changed"), None);
        assert_eq!(edit.len(), 2);
        assert_eq!(edit.entry("Data/changed"), Some(b"changed".as_slice()));
    }

    #[test]
    fn cloning_a_package_does_not_copy_entries_before_mutation() {
        let mut package = IWorkPackage::new();
        package.insert_entry("Data/asset", vec![7; 1024]).unwrap();
        let clone = package.clone();

        assert!(Arc::ptr_eq(&package.state, &clone.state));
        assert_eq!(package.entry("Data/asset"), clone.entry("Data/asset"));

        package.entry_mut("Data/asset").unwrap()[0] = 9;
        assert_eq!(package.entry("Data/asset").unwrap()[0], 9);
        assert_eq!(clone.entry("Data/asset").unwrap()[0], 7);
        assert!(!Arc::ptr_eq(&package.state, &clone.state));
    }

    #[test]
    fn snapshots_expose_indexed_read_operations_without_editing() {
        let mut package = IWorkPackage::new();
        package
            .insert_entry("Index/CalculationEngine-174.iwa", vec![1])
            .unwrap();
        package
            .insert_entry(
                "Index/Document.iwa",
                SnappyStream::compress(&archive().to_bytes().unwrap()).unwrap(),
            )
            .unwrap();

        let snapshot = package.snapshot();
        assert!(snapshot.contains_entry("/Index/Document.iwa"));
        assert_eq!(
            snapshot.iwa_entry_names().collect::<Vec<_>>(),
            ["Index/Document.iwa", "Index/CalculationEngine-174.iwa"]
        );
        assert_eq!(
            snapshot.calculation_engine_entry_name().unwrap(),
            Some("Index/CalculationEngine-174.iwa")
        );
        assert_eq!(
            snapshot
                .archive("Index/Document.iwa")
                .unwrap()
                .objects
                .len(),
            1
        );
    }

    #[test]
    fn structural_mutations_rebuild_the_shared_name_index() {
        let mut package = IWorkPackage::new();
        let expected_archive = SnappyStream::compress(&archive().to_bytes().unwrap()).unwrap();
        package.insert_entry("Data/a", vec![1]).unwrap();
        package.insert_entry("Data/b", vec![2]).unwrap();
        package
            .insert_archive_before("Index/Document.iwa", &archive(), "Data/b")
            .unwrap();

        assert_eq!(
            package.entry("Index/Document.iwa"),
            Some(expected_archive.as_slice())
        );
        assert_eq!(package.entry("Data/a"), Some(&[1][..]));
        assert_eq!(package.entry("Data/b"), Some(&[2][..]));
        assert_eq!(
            package.remove_entry("Index/Document.iwa"),
            Some(expected_archive)
        );
        assert!(!package.contains_entry("Index/Document.iwa"));
        assert_eq!(package.entry("Data/a"), Some(&[1][..]));
        assert_eq!(package.entry("Data/b"), Some(&[2][..]));
    }

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn snapshots_are_send_and_sync() {
        assert_send_sync::<Snapshot>();
    }

    #[test]
    fn preserves_member_order() {
        let mut package = IWorkPackage::new();
        package.insert_entry("Data/a", vec![1]).unwrap();
        package.insert_entry("Data/b", vec![2]).unwrap();
        package
            .replace_archive("Index/Document.iwa", &archive())
            .unwrap();

        let bytes = package.to_bytes().unwrap();
        let reparsed = IWorkPackage::from_bytes(&bytes).unwrap();
        assert_eq!(
            reparsed.entry_names().collect::<Vec<_>>(),
            ["Index/Document.iwa", "Data/a", "Data/b"]
        );
    }

    #[test]
    fn write_to_matches_memory_encoding() {
        let mut package = IWorkPackage::new();
        package.insert_entry("Data/a", vec![1, 2, 3]).unwrap();
        package
            .insert_entry("Metadata/Properties.plist", b"plist".to_vec())
            .unwrap();

        let expected = package.to_bytes().unwrap();
        let mut streamed = Vec::new();
        package.write_to(&mut streamed).unwrap();
        let mut snapshot_streamed = Vec::new();
        package.snapshot().write_to(&mut snapshot_streamed).unwrap();

        assert_eq!(streamed, expected);
        assert_eq!(snapshot_streamed, expected);
    }

    #[test]
    fn save_rejects_non_file_destinations() -> std::io::Result<()> {
        let directory = tempfile::tempdir()?;
        let package = IWorkPackage::new();

        let error = package.save(directory.path()).unwrap_err();
        assert!(error.to_string().contains("not a regular file"));
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn save_rejects_symbolic_link_destinations_without_replacing_target() -> std::io::Result<()> {
        use std::os::unix::fs::symlink;

        let directory = tempfile::tempdir()?;
        let target = directory.path().join("target.pages");
        let link = directory.path().join("document.pages");
        std::fs::write(&target, b"sentinel")?;
        symlink(&target, &link)?;

        let package = IWorkPackage::new();
        let error = package.save(&link).unwrap_err();

        assert!(error.to_string().contains("symbolic link"));
        assert_eq!(std::fs::read(&target)?, b"sentinel");
        assert!(std::fs::symlink_metadata(&link)?.file_type().is_symlink());
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn save_preserves_existing_regular_file_permissions() -> std::io::Result<()> {
        use std::os::unix::fs::PermissionsExt;

        let directory = tempfile::tempdir()?;
        let destination = directory.path().join("document.pages");
        std::fs::write(&destination, b"old")?;
        let mut permissions = std::fs::metadata(&destination)?.permissions();
        permissions.set_mode(0o640);
        std::fs::set_permissions(&destination, permissions)?;

        let mut package = IWorkPackage::new();
        package.insert_entry("Data/a", b"new".to_vec()).unwrap();
        package.save(&destination).unwrap();

        let mode = std::fs::metadata(&destination)?.permissions().mode() & 0o777;
        assert_eq!(mode, 0o640);
        assert_eq!(
            IWorkPackage::open(&destination).unwrap().entry("Data/a"),
            Some(&b"new"[..])
        );
        Ok(())
    }

    #[test]
    fn rejects_unsafe_entry_names() {
        let mut package = IWorkPackage::new();
        assert!(package.insert_entry("../escape", Vec::new()).is_err());
        assert!(package.insert_entry("/absolute", Vec::new()).is_ok());
        assert!(package.insert_entry("bad\\name", Vec::new()).is_err());
    }

    #[test]
    fn expands_legacy_nested_bundle_for_crud_without_losing_assets() {
        let mut package = IWorkPackage::from_bytes(&legacy_package()).unwrap();
        assert_eq!(
            package.entry_names().collect::<Vec<_>>(),
            [
                "Index/Document.iwa",
                "preview.jpg",
                "Metadata/Properties.plist"
            ]
        );
        assert_eq!(package.entry("preview.jpg"), Some(b"preview".as_slice()));
        assert_eq!(
            package.entry("Metadata/Properties.plist"),
            Some(b"plist".as_slice())
        );

        package
            .update_archive("Index/Document.iwa", |archive| {
                archive.object_mut(1).unwrap().messages[0].type_ = 100;
                Ok(())
            })
            .unwrap();
        let reparsed = IWorkPackage::from_bytes(&package.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reparsed.archive("Index/Document.iwa").unwrap().objects[0].messages[0].type_,
            100
        );
        assert_eq!(reparsed.entry("preview.jpg"), Some(b"preview".as_slice()));
        assert!(!reparsed.contains_entry("Index.zip"));
    }

    #[test]
    fn excludes_legacy_operation_log_from_iwa_archive_scans() {
        let mut package = IWorkPackage::new();
        package
            .insert_entry("Index/OperationStorage.iwa", b"bvxn log".to_vec())
            .unwrap();
        package
            .replace_archive("Index/Document.iwa", &archive())
            .unwrap();

        assert_eq!(
            package.iwa_entry_names().collect::<Vec<_>>(),
            ["Index/Document.iwa"]
        );
        let error = package.archive("Index/OperationStorage.iwa").unwrap_err();
        assert!(error.to_string().contains("legacy operation log"));
    }

    #[test]
    fn discovers_canonical_and_app_versioned_calculation_engines_strictly() {
        for entry in [
            "Index/CalculationEngine.iwa",
            "Index/CalculationEngine-174.iwa",
            "Index/CalculationEngine-10-2.iwa",
        ] {
            let mut package = IWorkPackage::new();
            package.insert_entry(entry, vec![1]).unwrap();
            assert_eq!(
                package.calculation_engine_entry_name().unwrap(),
                Some(entry)
            );
        }

        for entry in [
            "Index/CalculationEngine-.iwa",
            "Index/CalculationEngine-copy.iwa",
            "Index/CalculationEngine-1-.iwa",
            "Index/CalculationEngine-1.txt",
        ] {
            let mut package = IWorkPackage::new();
            package.insert_entry(entry, vec![1]).unwrap();
            assert_eq!(package.calculation_engine_entry_name().unwrap(), None);
        }
    }

    #[test]
    fn rejects_ambiguous_calculation_engines() {
        let mut package = IWorkPackage::new();
        package
            .insert_entry("Index/CalculationEngine.iwa", vec![1])
            .unwrap();
        package
            .insert_entry("Index/CalculationEngine-1.iwa", vec![2])
            .unwrap();

        assert!(package.calculation_engine_entry_name().is_err());
    }
}
