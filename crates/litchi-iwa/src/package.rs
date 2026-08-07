//! Mutable iWork ZIP package with entry-order preservation.

use std::fs::{self, File};
use std::io::Write;
use std::path::{Component, Path};
use std::sync::Arc;

use tempfile::NamedTempFile;

use crate::archive::{Archive, ArchiveLimits as IwaArchiveLimits};
use crate::snappy::{SnappyLimits, SnappyStream};
use crate::{Error, Result};
use litchi_core::ReadAt;
use litchi_iwa_package::{Entry, Patch};

#[path = "package_state.rs"]
mod package_state;
use package_state::{PackageState, entry_store_error, package_patch_error};

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
    mutation_revision: u64,
    source: Option<Arc<[u8]>>,
}

/// An immutable, cheaply shareable package snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    state: Arc<PackageState>,
    limits: PackageLimits,
    source: Option<Arc<[u8]>>,
}

/// The published result of an atomic package-level snapshot edit.
///
/// A commit owns the callback's result together with the immutable snapshot
/// produced after the candidate edit passed package-state validation. The
/// source snapshot used to start the edit is never mutated. The commit also
/// retains a source-checked, reversible package patch whose entry payloads
/// continue to share the snapshots' copy-on-write allocations.
#[must_use = "a commit contains the published snapshot"]
#[derive(Debug)]
pub struct Commit<T> {
    value: T,
    snapshot: Snapshot,
    patch: Patch,
}

impl<T> Commit<T> {
    /// Borrow the callback's result.
    pub fn value(&self) -> &T {
        &self.value
    }

    /// Borrow the immutable snapshot published by the commit.
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible package patch produced by this commit.
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit and return its result and published snapshot.
    pub fn into_parts(self) -> (T, Snapshot) {
        (self.value, self.snapshot)
    }

    /// Consume the commit and return its result, snapshot, and patch.
    pub fn into_parts_with_patch(self) -> (T, Snapshot, Patch) {
        (self.value, self.snapshot, self.patch)
    }
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
    iwa_archive_limits: IwaArchiveLimits,
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
            iwa_archive_limits: IwaArchiveLimits::default(),
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

    /// Set the checked resource budget applied to every parsed or serialized
    /// IWA archive in this package.
    ///
    /// The effective archive byte ceiling is the lower of this budget and
    /// [`Self::max_iwa_stream_bytes`], so a caller cannot make decompression
    /// allocate beyond the package's stream budget.
    pub fn with_archive_limits(mut self, limits: IwaArchiveLimits) -> Result<Self> {
        self.iwa_archive_limits = limits;
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

    /// Resource budget retained for one IWA archive parse or serialization.
    pub const fn archive_limits(self) -> IwaArchiveLimits {
        self.iwa_archive_limits
    }

    pub(crate) fn effective_archive_limits(self) -> Result<IwaArchiveLimits> {
        Ok(self.iwa_archive_limits.with_archive_bytes(
            self.max_iwa_stream_bytes
                .min(self.iwa_archive_limits.max_archive_bytes()),
        )?)
    }

    fn snappy_limits(self) -> Result<SnappyLimits> {
        let max_stream_bytes = self.effective_archive_limits()?.max_archive_bytes();
        SnappyLimits::new(
            max_stream_bytes.min(SnappyStream::MAX_UNCOMPRESSED_CHUNK),
            max_stream_bytes,
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
        let limit = self.effective_archive_limits()?.max_archive_bytes();
        if size > limit {
            return Err(Error::Snappy(format!(
                "IWA stream is {size} bytes, exceeding the {limit} byte limit"
            )));
        }
        Ok(())
    }

    fn physical_archive_limits(self) -> Result<litchi_iwa_archive::Limits> {
        let limits = litchi_iwa_archive::Limits::new(
            self.max_input_bytes,
            self.max_entries,
            self.max_entry_bytes,
            self.max_total_bytes,
            self.max_iwa_stream_bytes,
        )?;
        Ok(limits.with_archive_limits(self.effective_archive_limits()?)?)
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
            iwa_archive_limits: IwaArchiveLimits::default(),
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
    ///
    /// Flat packages retain the owned source bytes so an unchanged package
    /// can be written back byte-for-byte, including ZIP metadata and comments.
    /// Legacy nested packages still normalize through the archive adapter.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: PackageLimits) -> Result<Self> {
        let path = path.as_ref();
        let size = std::fs::metadata(path)?.len();
        limits.check_input_size(size, "iWork package input")?;
        let source: Arc<[u8]> = std::fs::read(path)?.into();
        Self::from_shared_bytes_with_limits(source, limits)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, PackageLimits::default())
    }

    /// Parse a package from an immutable, already-owned byte source.
    ///
    /// The source allocation is retained by the physical archive provenance,
    /// so callers that already own an `Arc<[u8]>` avoid an ingress copy while
    /// preserving exact ZIP bytes for a no-op write.
    pub fn from_shared_bytes(source: Arc<[u8]>) -> Result<Self> {
        Self::from_shared_bytes_with_limits(source, PackageLimits::default())
    }

    /// Parse a package from an immutable positional source under the default
    /// package limits.
    ///
    /// The source is read without a shared cursor and its version is checked
    /// before the parsed package is published. Semantic IWA archives remain
    /// lazy and use the package's bounded archive cache after ingress.
    pub fn from_read_at(source: &dyn ReadAt) -> Result<Self> {
        Self::from_read_at_with_limits(source, PackageLimits::default())
    }

    /// Parse a package from an immutable positional source under caller-
    /// selected package limits.
    ///
    /// The physical archive owner checks the source length before allocating,
    /// snapshots through [`ReadAt`], and rejects a changed source revision.
    pub fn from_read_at_with_limits(source: &dyn ReadAt, limits: PackageLimits) -> Result<Self> {
        let catalog = litchi_iwa_archive::package::Catalog::from_read_at_with_limits(
            source,
            limits.physical_archive_limits()?,
        )?;
        let source = catalog.source_is_exact().then(|| catalog.shared_source());
        Self::from_catalog(catalog, limits, source)
    }

    /// Parse a package under caller-selected ingress ceilings.
    ///
    /// The same ceilings are applied to the outer ZIP and, when present, the
    /// embedded legacy `Index.zip` archive. The owned source baseline is
    /// retained for flat packages, so an unchanged package can be written
    /// back byte-for-byte without rebuilding the ZIP envelope.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: PackageLimits) -> Result<Self> {
        Self::from_shared_bytes_with_limits(bytes.to_vec().into(), limits)
    }

    /// Parse an already-owned package source under caller-selected ingress
    /// ceilings without copying the source bytes.
    pub fn from_shared_bytes_with_limits(source: Arc<[u8]>, limits: PackageLimits) -> Result<Self> {
        let input_size = u64::try_from(source.len()).map_err(|_| {
            Error::InvalidFormat("iWork package input length does not fit u64".to_owned())
        })?;
        limits.check_input_size(input_size, "iWork package input")?;
        let catalog = litchi_iwa_archive::package::Catalog::from_shared_bytes_with_limits(
            Arc::clone(&source),
            limits.physical_archive_limits()?,
        )?;
        let source = catalog.source_is_exact().then_some(source);
        Self::from_catalog(catalog, limits, source)
    }

    fn from_catalog(
        catalog: litchi_iwa_archive::package::Catalog,
        limits: PackageLimits,
        source: Option<Arc<[u8]>>,
    ) -> Result<Self> {
        let mut entries = Vec::new();
        entries.try_reserve(catalog.len()).map_err(|_error| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "package entries",
                amount: catalog.len(),
            })
        })?;
        for entry in catalog {
            let (name, data) = entry.into_parts();
            validate_entry_name(&name)?;
            entries.push(Entry::new(name, data));
        }
        let archive_limits = limits.effective_archive_limits()?;
        let package = Self {
            state: Arc::new(PackageState::from_entries(entries, archive_limits)?),
            limits,
            mutation_revision: 0,
            source,
        };
        package.validate_state()?;
        Ok(package)
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

    /// Return the monotonic revision of this mutable package view.
    ///
    /// Format-specific editors use this compact token to reject an archive
    /// parsed before a caller performed another mutation. Cloning a package
    /// preserves the revision, while the first subsequent mutation advances
    /// only the mutated clone's revision.
    pub(crate) const fn mutation_revision(&self) -> u64 {
        self.mutation_revision
    }

    /// Capture an immutable package snapshot without copying package entries.
    pub fn snapshot(&self) -> Snapshot {
        Snapshot {
            state: Arc::clone(&self.state),
            limits: self.limits,
            source: self.source.clone(),
        }
    }

    pub fn entry_names(&self) -> impl Iterator<Item = &str> {
        self.state.entries.iter().map(Entry::name)
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
            .filter(|entry| {
                entry.name().ends_with(".iwa")
                    && !is_legacy_operation_storage(entry.name(), entry.data())
            })
            .map(Entry::name)
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
        Some(self.state.entries.get_at(position)?.data())
    }

    /// Borrow a raw package member for package invariant tests.
    ///
    /// This test-only helper ensures final validation still rejects a direct
    /// byte mutation without exposing an unrestricted mutable buffer in the
    /// production API.
    #[cfg(test)]
    fn entry_mut(&mut self, name: &str) -> Option<&mut Vec<u8>> {
        let name = normalize_entry_name(name);
        let position = self.entry_position(name)?;
        self.mark_mutated();
        let state = Arc::make_mut(&mut self.state);
        state.invalidate_archive(name);
        Some(state.entries.get_at_mut(position)?.data_mut())
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
            let state = Arc::make_mut(&mut self.state);
            let Some(previous) = state.entries.replace_data(position, data) else {
                return Err(Error::Bundle(
                    "package entry index became invalid during replacement".to_owned(),
                ));
            };
            state.invalidate_archive(&name);
            self.mark_mutated();
            return Ok(Some(previous));
        }
        self.insert_new_entry(name, data)?;
        self.mark_mutated();
        Ok(None)
    }

    /// Delete a package member.
    pub fn remove_entry(&mut self, name: &str) -> Option<Vec<u8>> {
        let normalized = normalize_entry_name(name);
        let position = self.entry_position(normalized)?;
        let state = Arc::make_mut(&mut self.state);
        let removed = state.entries.remove_at(position)?.into_parts().1;
        state.invalidate_archive(normalized);
        self.mark_mutated();
        Some(removed)
    }

    /// Parse a compressed `.iwa` package member.
    ///
    /// The immutable parsed representation is retained in the package-state
    /// cache and cloned only at this owned-return boundary. Repeated reads of
    /// one snapshot therefore share decompression and protobuf parsing.
    pub fn archive(&self, name: &str) -> Result<Archive> {
        let normalized = normalize_entry_name(name);
        Ok(self.parsed_archive(normalized)?.as_ref().clone())
    }

    /// Borrow a parsed `.iwa` package member through a bounded read cache.
    ///
    /// The callback never observes package-owned mutable state. Parsed
    /// archives are retained per copy-on-write package state under the
    /// decompressed-byte budget carried by the archive limits.
    pub(crate) fn with_parsed_archive<T, F>(&self, name: &str, read: F) -> Result<T>
    where
        F: FnOnce(&Archive) -> Result<T>,
    {
        let archive = self.parsed_archive(name)?;
        read(&archive)
    }

    pub(crate) fn parsed_archive(&self, name: &str) -> Result<Arc<Archive>> {
        let normalized = normalize_entry_name(name);
        self.state
            .get_or_parse_archive(normalized, |limits| self.parse_archive(normalized, limits))
    }

    fn parse_archive(
        &self,
        normalized: &str,
        archive_limits: IwaArchiveLimits,
    ) -> Result<(Archive, usize)> {
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
        let stream =
            SnappyStream::decompress_with_limits(compressed, self.limits.snappy_limits()?)?;
        let weight = stream.as_bytes().len().max(1);
        let archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)?;
        Ok((archive, weight))
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
        let data = archive.to_bytes_with_limits(self.limits.effective_archive_limits()?)?;
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
        let data = archive.to_bytes_with_limits(self.limits.effective_archive_limits()?)?;
        self.limits.check_iwa_stream_size(data.len())?;
        let compressed = SnappyStream::compress(&data)?;
        self.validate_entry_update(None, &compressed)?;
        let state = Arc::make_mut(&mut self.state);
        state
            .entries
            .try_insert_at(position, Entry::new(normalized, compressed))
            .map_err(entry_store_error)?;
        self.mark_mutated();
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
        archive.validate_with_limits(self.limits.effective_archive_limits()?)?;
        self.replace_archive(name, &archive)?;
        Ok(())
    }

    /// Validate the staged package state without encoding it.
    ///
    /// This is the explicit validation boundary used by package-invariant
    /// tests. It performs the same
    /// package-member and aggregate budget checks used by serialization, but
    /// does not allocate a ZIP buffer or publish any output.
    pub fn validate(&self) -> Result<()> {
        self.validate_state().map(|_| ())
    }

    /// Encode the package as a ZIP using stored members and the original order.
    ///
    /// Pages and Numbers use a leading `Index/Document.iwa` for package type
    /// discovery, so newly-created document indexes are inserted first.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.validate_state()?;
        if let Some(source) = self.source.as_deref() {
            let mut bytes = Vec::new();
            bytes.try_reserve_exact(source.len()).map_err(|_error| {
                Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                    resource: "iWork package source bytes",
                    amount: source.len(),
                })
            })?;
            bytes.extend_from_slice(source);
            return Ok(bytes);
        }
        litchi_iwa_archive::package::to_bytes(
            self.state
                .entries
                .as_slice()
                .iter()
                .map(|entry| (entry.name(), entry.data())),
            self.limits.physical_archive_limits()?,
        )
        .map_err(|error| Error::Bundle(format!("iWork package egress: {error}")))
    }

    /// Stream the package as a ZIP to a caller-owned sequential sink.
    ///
    /// The package is validated before the first byte is written. ZIP central
    /// directory records are finalized by the sink writer, so this method does
    /// not allocate a second buffer containing the complete package.
    pub fn write_to<W: Write>(&self, sink: W) -> Result<()> {
        self.validate_state()?;
        if let Some(source) = self.source.as_deref() {
            let mut sink = sink;
            sink.write_all(source)?;
            return Ok(());
        }
        litchi_iwa_archive::package::write_to(
            self.state
                .entries
                .as_slice()
                .iter()
                .map(|entry| (entry.name(), entry.data())),
            sink,
            self.limits.physical_archive_limits()?,
        )
        .map_err(|error| Error::Bundle(format!("iWork package egress: {error}")))
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
            .and_then(|index| self.state.entries.get_at(index))
            .map_or(0, |entry| entry.data().len());
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
        for entry in self.state.entries.iter() {
            validate_entry_name(entry.name())?;
            self.validate_entry_data(entry.data())?;
        }
        let total = self.state.entries.iter().try_fold(0_u64, |total, entry| {
            let size = u64::try_from(entry.data().len())
                .map_err(|_| Error::Bundle("package member length does not fit u64".to_owned()))?;
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

    fn insert_new_entry(&mut self, name: String, data: Vec<u8>) -> Result<()> {
        let state = Arc::make_mut(&mut self.state);
        let position = usize::from(name != "Index/Document.iwa") * state.entries.len();
        state
            .entries
            .try_insert_at(position, Entry::new(name, data))
            .map_err(entry_store_error)?;
        Ok(())
    }

    fn mark_mutated(&mut self) {
        self.mutation_revision = self.mutation_revision.wrapping_add(1);
        self.source = None;
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
        self.state.entries.iter().map(Entry::name)
    }

    /// Borrow one package member without copying it.
    pub fn entry(&self, name: &str) -> Option<&[u8]> {
        let name = normalize_entry_name(name);
        let position = self.state.position(name)?;
        Some(self.state.entries.get_at(position)?.data())
    }

    /// Validate this immutable snapshot without creating serialized output.
    ///
    /// The check borrows the shared package state and applies the same
    /// resource and entry-name invariants as [`IWorkPackage::validate`].
    pub fn validate(&self) -> Result<()> {
        self.edit().validate()
    }

    /// Apply a source-checked package patch without mutating this snapshot.
    pub fn apply(&self, patch: &Patch) -> Result<Self> {
        self.validate()?;
        let entries = patch
            .apply(&self.state.entries)
            .map_err(package_patch_error)?;
        let state = PackageState::from_store(entries, self.limits.effective_archive_limits()?);
        let target = Self {
            state: Arc::new(state),
            limits: self.limits,
            source: None,
        };
        target.validate()?;
        Ok(target)
    }

    /// Start a mutable copy-on-write edit from this snapshot.
    pub fn edit(&self) -> IWorkPackage {
        IWorkPackage {
            state: Arc::clone(&self.state),
            limits: self.limits,
            mutation_revision: 0,
            source: self.source.clone(),
        }
    }

    /// Stage, validate, and publish one atomic copy-on-write edit.
    ///
    /// The callback receives a private mutable package view. If it returns an
    /// error, or if the resulting package violates its configured limits or
    /// member-name invariants, no new snapshot is published and the source
    /// snapshot remains unchanged.
    pub fn edit_with<T, F>(&self, edit: F) -> Result<Commit<T>>
    where
        F: FnOnce(&mut IWorkPackage) -> Result<T>,
    {
        let mut candidate = self.edit();
        let value = edit(&mut candidate)?;
        candidate.validate_state()?;
        let snapshot = candidate.snapshot();
        Ok(Commit {
            value,
            patch: Patch::between(&self.state.entries, &snapshot.state.entries),
            snapshot,
        })
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
            .filter(|entry| {
                entry.name().ends_with(".iwa")
                    && !is_legacy_operation_storage(entry.name(), entry.data())
            })
            .map(Entry::name)
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
    use std::sync::Arc;
    use std::sync::Barrier;
    use std::sync::atomic::{AtomicUsize, Ordering};
    use std::thread;

    use litchi_iwa_package::EntryChangeKind;

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

    fn archive_with_two_objects() -> Archive {
        let mut archive = archive();
        archive.objects.push(
            ArchiveObject::new(
                2,
                vec![RawMessage {
                    type_: 100,
                    data: vec![4, 5, 6],
                }],
            )
            .unwrap(),
        );
        archive
    }

    fn zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
        litchi_iwa_archive::package::to_bytes(
            entries.iter().copied(),
            litchi_iwa_archive::Limits::default(),
        )
        .unwrap()
    }

    fn physical_entries(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
        litchi_iwa_archive::package::Catalog::from_bytes(bytes)
            .unwrap()
            .into_iter()
            .map(litchi_iwa_archive::package::Entry::into_parts)
            .collect()
    }

    fn legacy_package() -> Vec<u8> {
        let compressed = SnappyStream::compress(&archive().to_bytes().unwrap()).unwrap();
        let index = zip(&[("Index/Document.iwa", &compressed)]);
        zip(&[
            ("mac.numbers/preview.jpg", b"preview"),
            ("mac.numbers/Index.zip", &index),
            ("mac.numbers/Metadata/Properties.plist", b"plist"),
        ])
    }

    fn empty_package_with_limits(limits: PackageLimits) -> IWorkPackage {
        let bytes = zip(&[]);
        IWorkPackage::from_bytes_with_limits(&bytes, limits).unwrap()
    }

    #[test]
    fn shared_bytes_preserve_the_source_for_noop_writes() {
        let mut bytes = zip(&[("preview.jpg", b"preview")]);
        let end_of_central_directory = bytes.len() - 22;
        let comment = b"preserve-mode-comment";
        bytes[end_of_central_directory + 20..end_of_central_directory + 22]
            .copy_from_slice(&(comment.len() as u16).to_le_bytes());
        bytes.extend_from_slice(comment);
        let source: Arc<[u8]> = bytes.into();
        let package = IWorkPackage::from_shared_bytes(source.clone()).unwrap();

        assert_eq!(package.to_bytes().unwrap(), source.as_ref());
        let mut written = Vec::new();
        package.write_to(&mut written).unwrap();
        assert_eq!(written, source.as_ref());
        assert_eq!(package.entry_names().collect::<Vec<_>>(), ["preview.jpg"]);
    }

    #[test]
    fn borrowed_bytes_preserve_the_source_for_noop_writes() {
        let mut bytes = zip(&[("preview.jpg", b"preview")]);
        let end_of_central_directory = bytes.len() - 22;
        let comment = b"borrowed-preserve-mode-comment";
        bytes[end_of_central_directory + 20..end_of_central_directory + 22]
            .copy_from_slice(&(comment.len() as u16).to_le_bytes());
        bytes.extend_from_slice(comment);

        let package = IWorkPackage::from_bytes(&bytes).unwrap();

        assert_eq!(package.to_bytes().unwrap(), bytes);
        let mut written = Vec::new();
        package.write_to(&mut written).unwrap();
        assert_eq!(written, bytes);
    }

    #[test]
    fn path_open_preserves_flat_source_for_noop_writes() -> crate::Result<()> {
        let mut bytes = zip(&[("preview.jpg", b"preview")]);
        let end_of_central_directory = bytes.len() - 22;
        let comment = b"path-preserve-mode-comment";
        bytes[end_of_central_directory + 20..end_of_central_directory + 22]
            .copy_from_slice(&(comment.len() as u16).to_le_bytes());
        bytes.extend_from_slice(comment);

        let file = tempfile::NamedTempFile::new()?;
        std::fs::write(file.path(), &bytes)?;
        let package = IWorkPackage::open(file.path())?;

        assert_eq!(package.to_bytes()?, bytes);
        let mut written = Vec::new();
        package.write_to(&mut written)?;
        assert_eq!(written, bytes);
        Ok(())
    }

    #[test]
    fn reads_a_package_from_a_positional_source() {
        let bytes = zip(&[("preview.jpg", b"preview")]);
        let source = litchi_core::OwnedSource::new(bytes.clone());
        let package = IWorkPackage::from_read_at(&source).unwrap();

        assert_eq!(package.entry("preview.jpg"), Some(&b"preview"[..]));
        assert_eq!(package.to_bytes().unwrap(), bytes);
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
        let bytes = zip(&[("Index/Document.iwa", &compressed)]);

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
        assert!(matches!(
            error,
            Error::IwaCore(core)
                if matches!(
                    core.as_ref(),
                    litchi_iwa_core::Error::Limit {
                        kind: litchi_iwa_core::LimitKind::SnappyChunkBytes,
                        ..
                    }
                )
        ));
        let error = package
            .replace_archive("Index/Other.iwa", &archive())
            .unwrap_err();
        assert!(error.to_string().contains("IWA header byte"));

        let file = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(file.path(), &bytes).unwrap();
        let error = IWorkPackage::open_with_limits(file.path(), input_limits).unwrap_err();
        assert!(error.to_string().contains("iWork package input"));
    }

    #[test]
    fn custom_archive_limits_reach_lazy_parse_before_payload_allocation() -> crate::Result<()> {
        let source = archive_with_two_objects();
        let decompressed = source.to_bytes()?;
        let compressed = SnappyStream::compress(&decompressed)?;
        let bytes = zip(&[("Index/Document.iwa", &compressed)]);

        let archive_limits = IwaArchiveLimits::default().with_objects(1)?;
        let package_limits = PackageLimits::default().with_archive_limits(archive_limits)?;
        let package = IWorkPackage::from_bytes_with_limits(&bytes, package_limits)?;
        let error = package.archive("Index/Document.iwa").unwrap_err();
        assert!(error.to_string().contains("IWA object limit exceeded"));

        let byte_limits = IwaArchiveLimits::default().with_archive_bytes(1)?;
        let package_limits = PackageLimits::default().with_archive_limits(byte_limits)?;
        let package = IWorkPackage::from_bytes_with_limits(&bytes, package_limits)?;
        let error = package.archive("Index/Document.iwa").unwrap_err();
        assert!(matches!(
            error,
            Error::IwaCore(core)
                if matches!(
                    core.as_ref(),
                    litchi_iwa_core::Error::Limit {
                        kind: litchi_iwa_core::LimitKind::SnappyChunkBytes,
                        ..
                    }
                )
        ));
        Ok(())
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
    fn explicit_validation_checks_escape_hatches_without_mutating_snapshots() {
        let limits = PackageLimits::new(1, 4, 4).unwrap();
        let mut package = empty_package_with_limits(limits);
        package
            .insert_entry("Data/asset", vec![1, 2, 3, 4])
            .unwrap();
        let snapshot = package.snapshot();

        assert!(package.validate().is_ok());
        assert!(snapshot.validate().is_ok());

        package.entry_mut("Data/asset").unwrap().push(5);
        assert!(package.validate().is_err());
        assert!(snapshot.validate().is_ok());
        assert_eq!(snapshot.entry("Data/asset"), Some([1, 2, 3, 4].as_slice()));
    }

    #[test]
    fn package_limits_reject_flat_archive_metadata_before_materialization() {
        let bytes = zip(&[("Data/asset", b"asset"), ("Metadata/other", b"other")]);

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
        let index = zip(&[
            ("Index/Document.iwa", &compressed),
            ("Index/Metadata.iwa", &compressed),
        ]);
        let bytes = zip(&[("mac.pages/Index.zip", &index)]);

        let one_entry = PackageLimits::new(
            1,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
        )
        .unwrap();
        let error = IWorkPackage::from_bytes_with_limits(&bytes, one_entry).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("ZIP entries limit exceeded: observed 2, maximum 1"),
            "{error}"
        );
    }

    #[test]
    fn package_limits_apply_to_flattened_legacy_entry_count() {
        let compressed = SnappyStream::compress(&archive().to_bytes().unwrap()).unwrap();
        let index = zip(&[
            ("Index/Document.iwa", &compressed),
            ("Index/Metadata.iwa", &compressed),
        ]);
        let bytes = zip(&[
            ("mac.pages/Index.zip", &index),
            ("mac.pages/Data/asset", b"asset"),
        ]);

        let limits = PackageLimits::new(
            2,
            PackageLimits::MAX_ENTRY_BYTES,
            PackageLimits::MAX_TOTAL_BYTES,
        )
        .unwrap();
        let error = IWorkPackage::from_bytes_with_limits(&bytes, limits).unwrap_err();
        assert!(error.to_string().contains("entry count"));
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
    fn parsed_archive_cache_reuses_reads_and_invalidates_copy_on_write_mutations() {
        let mut package = IWorkPackage::new();
        let original = archive();
        package
            .replace_archive("Index/Document.iwa", &original)
            .unwrap();

        let first = package.parsed_archive("Index/Document.iwa").unwrap();
        let second = package.parsed_archive("Index/Document.iwa").unwrap();
        assert!(Arc::ptr_eq(&first, &second));
        assert!(first.object(1).is_some());

        let mut edited = package.clone();
        let mut replacement = archive();
        replacement.objects[0].archive_info.identifier = Some(2);
        edited
            .replace_archive("Index/Document.iwa", &replacement)
            .unwrap();

        let edited_archive = edited.parsed_archive("Index/Document.iwa").unwrap();
        assert!(!Arc::ptr_eq(&first, &edited_archive));
        assert!(edited_archive.object(2).is_some());
        assert!(
            package
                .parsed_archive("Index/Document.iwa")
                .is_ok_and(|archive| {
                    Arc::ptr_eq(&first, &archive) && archive.object(1).is_some()
                })
        );
    }

    #[test]
    fn parsed_archive_single_flight_shares_one_arc_across_threads() {
        const CALLERS: usize = 8;
        let state = Arc::new(PackageState::default());
        let ready = Arc::new(Barrier::new(CALLERS + 1));
        let first_parse_started = Arc::new(Barrier::new(2));
        let parse_count = Arc::new(AtomicUsize::new(0));
        let handles = (0..CALLERS)
            .map(|_| {
                let state = Arc::clone(&state);
                let ready = Arc::clone(&ready);
                let first_parse_started = Arc::clone(&first_parse_started);
                let parse_count = Arc::clone(&parse_count);
                thread::spawn(move || {
                    ready.wait();
                    state.get_or_parse_archive("Index/Document.iwa", |_| {
                        if parse_count.fetch_add(1, Ordering::SeqCst) == 0 {
                            first_parse_started.wait();
                        }
                        Ok((archive(), 1))
                    })
                })
            })
            .collect::<Vec<_>>();

        ready.wait();
        first_parse_started.wait();
        let mut results = handles
            .into_iter()
            .map(|handle| handle.join().unwrap().unwrap())
            .collect::<Vec<_>>();
        let first = results.pop().unwrap();
        assert!(results.iter().all(|result| Arc::ptr_eq(&first, result)));
        assert_eq!(parse_count.load(Ordering::SeqCst), 1);
    }

    #[test]
    fn failed_archive_flight_wakes_waiters_and_allows_a_retry() {
        const CALLERS: usize = 6;
        let state = Arc::new(PackageState::default());
        let ready = Arc::new(Barrier::new(CALLERS + 1));
        let first_parse_started = Arc::new(Barrier::new(2));
        let parse_count = Arc::new(AtomicUsize::new(0));
        let handles = (0..CALLERS)
            .map(|_| {
                let state = Arc::clone(&state);
                let ready = Arc::clone(&ready);
                let first_parse_started = Arc::clone(&first_parse_started);
                let parse_count = Arc::clone(&parse_count);
                thread::spawn(move || {
                    ready.wait();
                    state.get_or_parse_archive("Index/Document.iwa", |_| {
                        if parse_count.fetch_add(1, Ordering::SeqCst) == 0 {
                            first_parse_started.wait();
                        }
                        Err(Error::InvalidFormat("synthetic parse failure".to_owned()))
                    })
                })
            })
            .collect::<Vec<_>>();

        ready.wait();
        first_parse_started.wait();
        for handle in handles {
            let error = handle.join().unwrap().unwrap_err();
            assert!(error.to_string().contains("synthetic parse failure"));
        }

        let retry_count = AtomicUsize::new(0);
        let retry = state
            .get_or_parse_archive("Index/Document.iwa", |_| {
                retry_count.fetch_add(1, Ordering::SeqCst);
                Ok((archive(), 1))
            })
            .unwrap();
        assert!(retry.object(1).is_some());
        assert_eq!(retry_count.load(Ordering::SeqCst), 1);
        assert!(parse_count.load(Ordering::SeqCst) >= 1);
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
    fn snapshot_edit_with_publishes_a_validated_commit() -> crate::Result<()> {
        let mut package = IWorkPackage::new();
        package.insert_entry("Data/original", b"original".to_vec())?;
        let source = package.snapshot();

        let commit = source.edit_with(|edit| {
            edit.insert_entry("Data/changed", b"changed".to_vec())?;
            Ok(edit.len())
        })?;

        assert_eq!(commit.value(), &2);
        assert_eq!(commit.patch().version(), Patch::VERSION);
        assert_eq!(commit.patch().len(), 1);
        let change = &commit.patch().changes()[0];
        assert_eq!(change.name(), "Data/changed");
        assert_eq!(change.kind(), EntryChangeKind::Added);
        assert_eq!(change.after_position(), Some(1));
        assert_eq!(change.after_len(), Some(7));
        assert_eq!(source.len(), 1);
        assert_eq!(source.entry("Data/changed"), None);
        assert_eq!(commit.snapshot().len(), 2);
        assert_eq!(
            commit.snapshot().entry("Data/changed"),
            Some(b"changed".as_slice())
        );

        let replayed = source.apply(commit.patch())?;
        assert_eq!(replayed.entry("Data/changed"), Some(b"changed".as_slice()));
        let reverted = replayed.apply(&commit.patch().inverse())?;
        assert_eq!(reverted.entry("Data/changed"), None);
        assert_eq!(
            reverted.entry("Data/original"),
            Some(b"original".as_slice())
        );
        Ok(())
    }

    #[test]
    fn package_patch_rejects_a_stale_source() -> crate::Result<()> {
        let mut package = IWorkPackage::new();
        package.insert_entry("Data/original", b"original".to_vec())?;
        let source = package.snapshot();
        let commit = source.edit_with(|edit| {
            edit.insert_entry("Data/changed", b"changed".to_vec())?;
            Ok(())
        })?;

        let mut stale = IWorkPackage::new();
        stale.insert_entry("Data/original", b"different".to_vec())?;
        let error = stale.snapshot().apply(commit.patch()).unwrap_err();
        assert!(error.to_string().contains("patch source does not match"));
        Ok(())
    }

    #[test]
    fn snapshot_edit_with_discards_callback_failures() -> crate::Result<()> {
        let mut package = IWorkPackage::new();
        package.insert_entry("Data/original", b"original".to_vec())?;
        let source = package.snapshot();

        let error = source.edit_with(|edit| -> crate::Result<()> {
            edit.insert_entry("Data/changed", b"changed".to_vec())?;
            Err(Error::InvalidFormat("reject staged edit".to_owned()))
        });

        assert!(error.is_err());
        assert_eq!(source.len(), 1);
        assert_eq!(source.entry("Data/changed"), None);
        Ok(())
    }

    #[test]
    fn snapshot_edit_with_discards_candidates_that_fail_final_validation() {
        let limits = PackageLimits::new(1, 4, 4).unwrap();
        let mut package = empty_package_with_limits(limits);
        package
            .insert_entry("Data/asset", vec![1, 2, 3, 4])
            .unwrap();
        let source = package.snapshot();

        let error = source.edit_with(|edit| {
            edit.entry_mut("Data/asset").unwrap().push(5);
            Ok(())
        });

        assert!(error.unwrap_err().to_string().contains("member"));
        assert_eq!(source.entry("Data/asset"), Some([1, 2, 3, 4].as_slice()));
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
    fn preserves_opaque_entry_bytes_and_order_across_egress() {
        let input = zip(&[
            ("Metadata/first", b"\0metadata\xff"),
            ("Index/Unknown.iwa", b"opaque iwa bytes"),
            ("Data/last", b"\xffasset\0"),
        ]);
        let package = IWorkPackage::from_bytes(&input).unwrap();
        let expected = physical_entries(&input);

        assert_eq!(physical_entries(&package.to_bytes().unwrap()), expected);
        let mut streamed = Vec::new();
        package.write_to(&mut streamed).unwrap();
        assert_eq!(physical_entries(&streamed), expected);
    }

    #[test]
    fn egress_validates_before_emitting_zip_bytes() {
        let limits = PackageLimits::new(1, 4, 4).unwrap();
        let mut package = empty_package_with_limits(limits);
        package
            .insert_entry("Data/asset", vec![1, 2, 3, 4])
            .unwrap();
        package.entry_mut("Data/asset").unwrap().push(5);

        let mut sink = Vec::new();
        assert!(package.write_to(&mut sink).is_err());
        assert!(sink.is_empty());
    }

    #[test]
    fn update_archive_failure_preserves_bytes_order_and_revision() {
        let mut package = IWorkPackage::new();
        package
            .insert_entry("Data/asset", b"asset".to_vec())
            .unwrap();
        package
            .replace_archive("Index/Document.iwa", &archive())
            .unwrap();
        let before_entries = package.entry_names().map(str::to_owned).collect::<Vec<_>>();
        let before_bytes = package.entry("Index/Document.iwa").unwrap().to_vec();
        let before_revision = package.mutation_revision();

        let error = package
            .update_archive("Index/Document.iwa", |_archive| {
                Err(Error::InvalidFormat("reject archive edit".to_owned()))
            })
            .unwrap_err();

        assert!(error.to_string().contains("reject archive edit"));
        assert_eq!(
            package.entry_names().collect::<Vec<_>>(),
            before_entries
                .iter()
                .map(String::as_str)
                .collect::<Vec<_>>()
        );
        assert_eq!(
            package.entry("Index/Document.iwa"),
            Some(before_bytes.as_slice())
        );
        assert_eq!(package.mutation_revision(), before_revision);
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
