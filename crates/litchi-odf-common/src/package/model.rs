//! Neutral ODF package and manifest models.

use soapberry_zip::office::{
    ArchiveLimits as ZipArchiveLimits, ArchiveReader, ArchiveReaderNames, IndexedArchive,
    IndexedArchiveNames,
};
use std::collections::HashMap;
use std::sync::Arc;

#[cfg(test)]
use std::cell::Cell;

#[cfg(test)]
thread_local! {
    static INDEX_BUILDS: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
pub(crate) fn note_index_build() {
    INDEX_BUILDS.with(|count| count.set(count.get().saturating_add(1)));
}

#[cfg(test)]
pub(crate) fn reset_index_build_count() {
    INDEX_BUILDS.with(|count| count.set(0));
}

#[cfg(test)]
pub(crate) fn index_build_count() -> usize {
    INDEX_BUILDS.with(Cell::get)
}

/// A borrowed, lazily decoded ODF ZIP archive.
pub struct Archive<'data> {
    pub(super) reader: ArchiveReaderKind<'data>,
}

/// Declared metadata for one ODF archive member.
///
/// The values are copied from the ZIP central directory.  Keeping this
/// format-neutral view here prevents the raw ZIP metadata type from crossing
/// the ODF package boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ArchiveMetadata {
    pub(crate) compressed_size: u64,
    pub(crate) uncompressed_size: u64,
    pub(crate) directory: bool,
}

/// Checked ZIP catalog limits used by the ODF package owner.
///
/// This is deliberately a format-owned policy type.  The underlying
/// `soapberry-zip` limits remain an implementation detail of the archive
/// indexer and never appear in ODF package signatures.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ArchiveLimits {
    max_files: usize,
    max_member_name_bytes: u64,
    max_metadata_bytes: u64,
    max_compressed_size: u64,
    max_entry_size: u64,
    max_total_size: u64,
}

impl ArchiveLimits {
    /// A profile with all ZIP ceilings disabled.  Package owners still apply
    /// their own finite hard ceilings before this profile reaches the indexer.
    pub const UNBOUNDED: Self = Self {
        max_files: usize::MAX,
        max_member_name_bytes: u64::MAX,
        max_metadata_bytes: u64::MAX,
        max_compressed_size: u64::MAX,
        max_entry_size: u64::MAX,
        max_total_size: u64::MAX,
    };

    /// Construct an explicit ZIP catalog profile.
    #[must_use]
    pub const fn new(
        max_files: usize,
        max_member_name_bytes: u64,
        max_metadata_bytes: u64,
        max_compressed_size: u64,
        max_entry_size: u64,
        max_total_size: u64,
    ) -> Self {
        Self {
            max_files,
            max_member_name_bytes,
            max_metadata_bytes,
            max_compressed_size,
            max_entry_size,
            max_total_size,
        }
    }

    /// Return the maximum number of non-directory members.
    #[must_use]
    pub const fn max_files(self) -> usize {
        self.max_files
    }

    /// Return the maximum raw member-name bytes.
    #[must_use]
    pub const fn max_member_name_bytes(self) -> u64 {
        self.max_member_name_bytes
    }

    /// Return the maximum aggregate central-directory metadata bytes.
    #[must_use]
    pub const fn max_metadata_bytes(self) -> u64 {
        self.max_metadata_bytes
    }

    /// Return the maximum declared compressed bytes for one member.
    #[must_use]
    pub const fn max_compressed_size(self) -> u64 {
        self.max_compressed_size
    }

    /// Return the maximum declared uncompressed bytes for one member.
    #[must_use]
    pub const fn max_entry_size(self) -> u64 {
        self.max_entry_size
    }

    /// Return the maximum aggregate declared uncompressed bytes.
    #[must_use]
    pub const fn max_total_size(self) -> u64 {
        self.max_total_size
    }

    /// Return a copy with a different member-count ceiling.
    #[must_use]
    pub const fn with_max_files(mut self, maximum: usize) -> Self {
        self.max_files = maximum;
        self
    }

    /// Return a copy with a different raw member-name ceiling.
    #[must_use]
    pub const fn with_max_member_name_bytes(mut self, maximum: u64) -> Self {
        self.max_member_name_bytes = maximum;
        self
    }

    /// Return a copy with a different central-directory metadata ceiling.
    #[must_use]
    pub const fn with_max_metadata_bytes(mut self, maximum: u64) -> Self {
        self.max_metadata_bytes = maximum;
        self
    }

    /// Return a copy with a different per-member compressed-size ceiling.
    #[must_use]
    pub const fn with_max_compressed_size(mut self, maximum: u64) -> Self {
        self.max_compressed_size = maximum;
        self
    }

    /// Return a copy with a different per-member uncompressed-size ceiling.
    #[must_use]
    pub const fn with_max_entry_size(mut self, maximum: u64) -> Self {
        self.max_entry_size = maximum;
        self
    }

    /// Return a copy with a different aggregate uncompressed-size ceiling.
    #[must_use]
    pub const fn with_max_total_size(mut self, maximum: u64) -> Self {
        self.max_total_size = maximum;
        self
    }

    pub(crate) const fn into_zip_limits(self) -> ZipArchiveLimits {
        ZipArchiveLimits {
            max_files: self.max_files,
            max_member_name_bytes: self.max_member_name_bytes,
            max_metadata_bytes: self.max_metadata_bytes,
            max_compressed_size: self.max_compressed_size,
            max_entry_size: self.max_entry_size,
            max_total_size: self.max_total_size,
        }
    }
}

impl Default for ArchiveLimits {
    fn default() -> Self {
        Self::from_zip_limits(ZipArchiveLimits::default())
    }
}

impl ArchiveLimits {
    const fn from_zip_limits(limits: ZipArchiveLimits) -> Self {
        Self {
            max_files: limits.max_files,
            max_member_name_bytes: limits.max_member_name_bytes,
            max_metadata_bytes: limits.max_metadata_bytes,
            max_compressed_size: limits.max_compressed_size,
            max_entry_size: limits.max_entry_size,
            max_total_size: limits.max_total_size,
        }
    }
}

impl ArchiveMetadata {
    /// Return the declared compressed member size.
    #[must_use]
    pub const fn compressed_size(self) -> u64 {
        self.compressed_size
    }

    /// Return the declared uncompressed member size.
    #[must_use]
    pub const fn uncompressed_size(self) -> u64 {
        self.uncompressed_size
    }

    /// Return whether the central-directory member is a directory.
    #[must_use]
    pub const fn is_directory(self) -> bool {
        self.directory
    }
}

/// The archive index retained by an owned ODF package.
pub(crate) type PreparedArchive = Arc<IndexedArchive<Arc<Vec<u8>>>>;

/// A borrowed archive reader or an already-indexed owned archive.
pub(super) enum ArchiveReaderKind<'data> {
    Borrowed(ArchiveReader<'data>),
    Prepared(PreparedArchive),
}

pub enum ArchiveNames<'data> {
    Borrowed(ArchiveReaderNames<'data>),
    Prepared(IndexedArchiveNames<'data, Arc<Vec<u8>>>),
}

impl<'data> Iterator for ArchiveNames<'data> {
    type Item = &'data str;

    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Self::Borrowed(names) => names.next(),
            Self::Prepared(names) => names.next(),
        }
    }
}

impl ExactSizeIterator for ArchiveNames<'_> {}

/// The family-neutral portion of `META-INF/manifest.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Manifest {
    /// The media type declared for the root entry.
    pub mimetype: String,
    /// Entries keyed by their normalized manifest path.
    pub entries: HashMap<String, Entry>,
}

/// One family-neutral ODF manifest file entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    /// The media type declared by `manifest:media-type`.
    pub media_type: String,
    /// The optional plaintext size declared by `manifest:size`.
    pub size: Option<u64>,
}

impl Manifest {
    /// Return the media type for a manifest path, if declared.
    #[must_use]
    pub fn get_media_type(&self, path: &str) -> Option<&str> {
        self.entries
            .get(path)
            .map(|entry| entry.media_type.as_str())
    }

    /// Check whether a manifest path is declared.
    #[must_use]
    pub fn has_path(&self, path: &str) -> bool {
        self.entries.contains_key(path)
    }

    /// Iterate over declared manifest paths.
    pub fn paths(&self) -> impl Iterator<Item = &str> {
        self.entries.keys().map(String::as_str)
    }

    /// Return a manifest entry by path.
    #[must_use]
    pub fn get_entry(&self, path: &str) -> Option<&Entry> {
        self.entries.get(path)
    }
}
