//! High-level ZIP archive API optimized for Office document formats.
//!
//! This module provides a simplified interface for reading and writing ZIP archives,
//! specifically optimized for OOXML, ODF, and iWork file formats that use Deflate
//! compression exclusively.
//!
//! # Reading Archives
//!
//! ```rust,no_run
//! use soapberry_zip::office::ArchiveReader;
//!
//! let data = std::fs::read("document.docx")?;
//! let archive = ArchiveReader::new(&data)?;
//!
//! // Read a specific file
//! let content = archive.read("word/document.xml")?;
//!
//! // Iterate over all files
//! for name in archive.file_names() {
//!     println!("{}", name);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```
//!
//! # Writing Archives
//!
//! ```rust,no_run
//! use soapberry_zip::office::StreamingArchiveWriter;
//!
//! let mut writer = StreamingArchiveWriter::new();
//! writer.write_stored("mimetype", b"application/vnd.oasis.opendocument.text")?;
//! writer.write_deflated("content.xml", b"<office:document-content>...</office:document-content>")?;
//! let bytes = writer.finish_to_bytes()?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use crate::{
    CompressionMethod, Error, ErrorKind, ZipArchive, ZipArchiveWriter, ZipSliceArchive,
    ZipVerification,
};
use flate2::Compression;
use flate2::read::DeflateDecoder;
use flate2::write::DeflateEncoder;
use std::collections::HashMap;
use std::io::{Read, Write};

pub use crate::LimitResource;

/// Resource limits applied while indexing an Office ZIP package.
///
/// The defaults accommodate large embedded media while rejecting implausible
/// archive metadata before any decompression allocation occurs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ArchiveLimits {
    /// Maximum number of non-directory entries.
    pub max_files: usize,
    /// Maximum bytes in one raw member name.
    pub max_member_name_bytes: u64,
    /// Maximum aggregate central-directory metadata bytes.
    ///
    /// This includes raw member names, extra fields, and file comments for all
    /// entries, including directories. It is checked before name normalization
    /// or ownership allocation.
    pub max_metadata_bytes: u64,
    /// Maximum declared compressed bytes for one non-directory entry.
    pub max_compressed_size: u64,
    /// Maximum declared uncompressed size of one entry.
    pub max_entry_size: u64,
    /// Maximum sum of all declared uncompressed entry sizes.
    pub max_total_size: u64,
}

impl ArchiveLimits {
    /// Disable resource ceilings while retaining integer and allocation checks.
    pub const UNBOUNDED: Self = Self {
        max_files: usize::MAX,
        max_member_name_bytes: u64::MAX,
        max_metadata_bytes: u64::MAX,
        max_compressed_size: u64::MAX,
        max_entry_size: u64::MAX,
        max_total_size: u64::MAX,
    };
}

impl Default for ArchiveLimits {
    fn default() -> Self {
        Self {
            max_files: 100_000,
            max_member_name_bytes: 4 * 1024,
            max_metadata_bytes: 64 * 1024 * 1024,
            max_compressed_size: 512 * 1024 * 1024,
            max_entry_size: 512 * 1024 * 1024,
            max_total_size: 2 * 1024 * 1024 * 1024,
        }
    }
}

/// High-performance ZIP archive reader for Office document formats.
///
/// Provides a simple API for reading ZIP archives with automatic decompression.
/// Optimized for OOXML (.docx, .xlsx, .pptx), ODF (.odt, .ods, .odp), and
/// iWork (.pages, .numbers, .key) formats.
///
/// # Performance
///
/// - Zero-copy parsing of archive structure
/// - Lazy decompression - only decompress files when accessed
/// - Pre-indexed file lookup for O(1) access by name
pub struct ArchiveReader<'data> {
    archive: ZipSliceArchive<&'data [u8]>,
    /// Pre-built index for fast file lookup by name
    index: HashMap<String, EntryInfo>,
    /// Directory declarations, retained for metadata lookup without changing
    /// the file-only behavior of the main index.
    directories: HashMap<String, Metadata>,
    /// Physical member order, retained for order-sensitive package formats.
    order: Vec<String>,
}

/// Information about an archive entry for fast lookup
#[derive(Debug, Clone)]
struct EntryInfo {
    wayfinder: crate::ZipArchiveEntryWayfinder,
    compression_method: CompressionMethod,
    uncompressed_size: u64,
}

/// Declared ZIP member metadata available without accessing member payloads.
///
/// The values originate in the central directory and are not independently
/// verified until a file is read. This compact copyable view supports safe
/// structural inspection without decompression or cache population.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Metadata {
    compressed_size: u64,
    uncompressed_size: u64,
    directory: bool,
}

impl Metadata {
    /// Returns the declared compressed member size.
    #[inline]
    pub const fn compressed_size(&self) -> u64 {
        self.compressed_size
    }

    /// Returns the declared uncompressed member size.
    #[inline]
    pub const fn uncompressed_size(&self) -> u64 {
        self.uncompressed_size
    }

    /// Returns whether the central-directory member is a directory.
    #[inline]
    pub const fn is_directory(&self) -> bool {
        self.directory
    }
}

impl<'data> ArchiveReader<'data> {
    /// Create a new archive reader from a byte slice.
    ///
    /// This parses the ZIP central directory and builds an index for fast
    /// file lookup. The actual file contents are not decompressed until
    /// accessed via `read()`.
    pub fn new(data: &'data [u8]) -> Result<Self, Error> {
        Self::new_with_limits(data, ArchiveLimits::default())
    }

    /// Create a reader with explicit resource limits.
    pub fn new_with_limits(data: &'data [u8], limits: ArchiveLimits) -> Result<Self, Error> {
        let archive = ZipArchive::from_slice(data)?;

        // Build index for fast lookup
        let mut index = HashMap::new();
        let mut directories = HashMap::new();
        let mut total_metadata_bytes = 0u64;
        let mut total_uncompressed_size = 0u64;
        let mut ordered_names = Vec::new();
        for entry_result in archive.entries() {
            let entry = entry_result?;
            let path = entry.file_path();

            let member_name_bytes = path.as_ref().len() as u64;
            if member_name_bytes > limits.max_member_name_bytes {
                return Err(limit_error(
                    LimitResource::MemberNameBytes,
                    member_name_bytes,
                    limits.max_member_name_bytes,
                ));
            }

            let metadata_bytes = entry.metadata_size_hint();
            total_metadata_bytes = total_metadata_bytes
                .checked_add(metadata_bytes)
                .ok_or_else(|| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: "archive central-directory metadata total overflows u64".to_string(),
                    })
                })?;
            if total_metadata_bytes > limits.max_metadata_bytes {
                return Err(limit_error(
                    LimitResource::MetadataBytes,
                    total_metadata_bytes,
                    limits.max_metadata_bytes,
                ));
            }

            let directory = entry.is_dir();

            if !directory && index.len() >= limits.max_files {
                let actual = (index.len() as u64).checked_add(1).ok_or_else(|| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: "archive file count overflows u64".to_string(),
                    })
                })?;
                return Err(limit_error(
                    LimitResource::FileCount,
                    actual,
                    limits.max_files as u64,
                ));
            }

            let compressed_size = entry.compressed_size_hint();
            if !directory && compressed_size > limits.max_compressed_size {
                return Err(limit_error(
                    LimitResource::CompressedSize,
                    compressed_size,
                    limits.max_compressed_size,
                ));
            }

            let uncompressed_size = entry.uncompressed_size_hint();
            if !directory && uncompressed_size > limits.max_entry_size {
                return Err(limit_error(
                    LimitResource::EntrySize,
                    uncompressed_size,
                    limits.max_entry_size,
                ));
            }
            if !directory {
                total_uncompressed_size = total_uncompressed_size
                    .checked_add(uncompressed_size)
                    .ok_or_else(|| {
                        Error::from(ErrorKind::InvalidInput {
                            msg: "archive uncompressed size total overflows u64".to_string(),
                        })
                    })?;
                if total_uncompressed_size > limits.max_total_size {
                    return Err(limit_error(
                        LimitResource::TotalSize,
                        total_uncompressed_size,
                        limits.max_total_size,
                    ));
                }
            }

            let name = match path.try_normalize() {
                Ok(normalized) => normalized.as_ref().to_string(),
                Err(_) => {
                    // Fallback to raw path as lossy UTF-8
                    String::from_utf8_lossy(path.as_ref()).to_string()
                },
            };

            // Directories are never exposed or decompressed by this API. They
            // consume name and metadata budgets above, but not file or payload
            // budgets. Retaining their compact declarations makes structural
            // inspection possible without changing file lookup behavior.
            if directory {
                directories.entry(name).or_insert(Metadata {
                    compressed_size,
                    uncompressed_size,
                    directory: true,
                });
                continue;
            }

            let local_header_offset = entry.local_header_offset();
            if index
                .insert(
                    name.clone(),
                    EntryInfo {
                        wayfinder: entry.wayfinder(),
                        compression_method: entry.compression_method(),
                        uncompressed_size,
                    },
                )
                .is_some()
            {
                return Err(ErrorKind::InvalidInput {
                    msg: "archive contains duplicate normalized file names".to_string(),
                }
                .into());
            }
            ordered_names.push((local_header_offset, name));
        }

        ordered_names.sort_by_key(|(offset, _)| *offset);
        let order = ordered_names.into_iter().map(|(_, name)| name).collect();

        Ok(Self {
            archive,
            index,
            directories,
            order,
        })
    }

    /// Get the number of files in the archive (excluding directories).
    #[inline]
    pub fn len(&self) -> usize {
        self.index.len()
    }

    /// Check if the archive is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.index.is_empty()
    }

    /// Check if a file exists in the archive.
    #[inline]
    pub fn contains(&self, name: &str) -> bool {
        // Try exact match first
        if self.index.contains_key(name) {
            return true;
        }
        // Try without leading slash
        let normalized = name.strip_prefix('/').unwrap_or(name);
        self.index.contains_key(normalized)
    }

    /// Return declared metadata for a normalized member name.
    ///
    /// This performs only hash-map lookup over the central-directory index. It
    /// never reads, decompresses, verifies, or allocates member payload data.
    pub fn metadata(&self, name: &str) -> Result<Metadata, Error> {
        let normalized = name.strip_prefix('/').unwrap_or(name);
        if let Some(info) = self.index.get(normalized) {
            return Ok(Metadata {
                compressed_size: info.wayfinder.compressed_size_hint(),
                uncompressed_size: info.uncompressed_size,
                directory: false,
            });
        }
        self.directories
            .get(normalized)
            .copied()
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(normalized.to_string())))
    }

    /// Get an iterator over all file names in the archive.
    pub fn file_names(&self) -> impl Iterator<Item = &str> {
        self.order.iter().map(String::as_str)
    }

    /// Whether an archive entry uses the ZIP Store method.
    ///
    /// ODF encryption is applied to an already-deflated byte stream, so the
    /// enclosing ZIP entry must not perform another compression transform.
    pub fn is_stored(&self, name: &str) -> Result<bool, Error> {
        let normalized = name.strip_prefix('/').unwrap_or(name);
        let info = self
            .index
            .get(normalized)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(normalized.to_string())))?;
        Ok(info.compression_method == CompressionMethod::Store)
    }

    /// Read and decompress a file from the archive.
    ///
    /// Returns the decompressed contents of the file. Supports both stored
    /// (uncompressed) and deflated entries.
    pub fn read(&self, name: &str) -> Result<Vec<u8>, Error> {
        // Normalize name - remove leading slash if present
        let normalized = name.strip_prefix('/').unwrap_or(name);

        let info = self
            .index
            .get(normalized)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(normalized.to_string())))?;

        let entry = self.archive.get_entry(info.wayfinder)?;
        let data = entry.data();

        match info.compression_method {
            CompressionMethod::Store => {
                // Stored (uncompressed) - verify and return directly
                let verifier = entry.claim_verifier();
                verifier.valid(ZipVerification {
                    crc: crate::crc32(data),
                    uncompressed_size: data.len() as u64,
                })?;
                Ok(data.to_vec())
            },
            CompressionMethod::Deflate => {
                let size = usize::try_from(info.uncompressed_size).map_err(|_| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: format!(
                            "archive entry size {} does not fit this platform",
                            info.uncompressed_size
                        ),
                    })
                })?;
                let mut decompressed = Vec::new();
                decompressed.try_reserve_exact(size).map_err(|error| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: format!("could not allocate {size} bytes for archive entry: {error}"),
                    })
                })?;

                let decoder = entry.verifying_reader(DeflateDecoder::new(data));
                decoder
                    .take(info.uncompressed_size.saturating_add(1))
                    .read_to_end(&mut decompressed)?;
                if decompressed.len() != size {
                    return Err(ErrorKind::InvalidSize {
                        expected: info.uncompressed_size,
                        actual: decompressed.len() as u64,
                    }
                    .into());
                }
                Ok(decompressed)
            },
            other => Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                other.as_id().as_u16(),
            ))),
        }
    }

    /// Read a file as a UTF-8 string.
    ///
    /// Convenience method that reads and decodes the file as UTF-8.
    pub fn read_string(&self, name: &str) -> Result<String, Error> {
        let bytes = self.read(name)?;
        String::from_utf8(bytes).map_err(|e| {
            Error::from(ErrorKind::Io(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                e,
            )))
        })
    }

    /// Read and decompress multiple files in parallel.
    ///
    /// This uses rayon for parallel decompression, providing significant speedup
    /// when reading many compressed files (typical for OOXML/ODF documents).
    ///
    /// Returns a vector of (name, result) pairs in the same order as input.
    /// Each result is either the decompressed bytes or an error.
    ///
    /// # Example
    /// ```rust,no_run
    /// use soapberry_zip::office::ArchiveReader;
    ///
    /// let data = std::fs::read("document.docx")?;
    /// let archive = ArchiveReader::new(&data)?;
    ///
    /// let files = vec!["word/document.xml", "word/styles.xml"];
    /// let results = archive.read_many_parallel(&files);
    ///
    /// for (name, result) in results {
    ///     match result {
    ///         Ok(bytes) => println!("{}: {} bytes", name, bytes.len()),
    ///         Err(e) => eprintln!("{}: error: {}", name, e),
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn read_many_parallel<'a, S: AsRef<str> + Sync>(
        &self,
        names: &'a [S],
    ) -> Vec<(&'a S, Result<Vec<u8>, Error>)> {
        use rayon::prelude::*;

        names
            .par_iter()
            .map(|name| (name, self.read(name.as_ref())))
            .collect()
    }

    /// Read all files from the archive in parallel.
    ///
    /// Results retain physical source order. Every member has an explicit
    /// result, including corrupt or otherwise unreadable members.
    ///
    /// This is optimal when you need to access most/all files in the archive.
    pub fn read_all_parallel(&self) -> Vec<(String, Result<Vec<u8>, Error>)> {
        use rayon::prelude::*;

        self.order
            .par_iter()
            .map(|name| (name.clone(), self.read(name)))
            .collect()
    }
}

#[inline]
fn limit_error(resource: LimitResource, actual: u64, maximum: u64) -> Error {
    ErrorKind::LimitExceeded {
        resource,
        actual,
        maximum,
    }
    .into()
}

impl std::fmt::Debug for ArchiveReader<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("ArchiveReader")
            .field("file_count", &self.index.len())
            .finish()
    }
}

/// High-performance streaming ZIP archive writer for Office document formats.
///
/// This is the recommended writer for creating complete ZIP archives.
pub struct StreamingArchiveWriter<W: Write> {
    archive: ZipArchiveWriter<W>,
}

impl StreamingArchiveWriter<std::io::Cursor<Vec<u8>>> {
    /// Create a new streaming archive writer that writes to memory.
    pub fn new() -> Self {
        Self {
            archive: ZipArchiveWriter::new(std::io::Cursor::new(Vec::new())),
        }
    }

    /// Finish writing and return the ZIP archive bytes.
    pub fn finish_to_bytes(self) -> Result<Vec<u8>, Error> {
        let cursor = self.archive.finish()?;
        Ok(cursor.into_inner())
    }
}

impl<W: Write> StreamingArchiveWriter<W> {
    /// Create a new streaming archive writer with a custom writer.
    pub fn with_writer(writer: W) -> Self {
        Self {
            archive: ZipArchiveWriter::new(writer),
        }
    }

    /// Write a file without compression (stored).
    pub fn write_stored(&mut self, name: &str, data: &[u8]) -> Result<(), Error> {
        self.archive.write_stored_file(name, data)
    }

    /// Write a file with Deflate compression.
    pub fn write_deflated(&mut self, name: &str, data: &[u8]) -> Result<(), Error> {
        let (mut entry, config) = self
            .archive
            .new_file(name)
            .compression_method(CompressionMethod::Deflate)
            .start()?;

        let encoder = DeflateEncoder::new(&mut entry, Compression::default());
        let mut writer = config.wrap(encoder);
        writer.write_all(data)?;
        let (encoder, desc) = writer.finish()?;
        encoder.finish()?;
        entry.finish(desc)?;
        Ok(())
    }

    /// Finish writing the archive.
    pub fn finish(self) -> Result<W, Error> {
        self.archive.finish()
    }
}

impl Default for StreamingArchiveWriter<std::io::Cursor<Vec<u8>>> {
    fn default() -> Self {
        Self::new()
    }
}

// Ensure ArchiveReader is Send + Sync for parallel iteration
// This is a compile-time assertion
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<ArchiveReader<'static>>();
};

/// Lazy ZIP archive reader with on-demand decompression and caching.
///
/// Unlike `ArchiveReader::read_all_parallel()` which decompresses everything upfront,
/// this reader decompresses files on-demand as they are accessed. This is optimal for:
/// - Large archives where only a subset of files are needed
/// - Pipelining decompression with parsing (process files as they become available)
/// - Reducing memory pressure by not holding all decompressed data at once
///
/// The reader uses interior mutability for thread-safe caching of decompressed data.
///
/// # Example
/// ```rust,no_run
/// use soapberry_zip::office::LazyArchiveReader;
///
/// let data = std::fs::read("document.docx")?;
/// let archive = LazyArchiveReader::new(&data)?;
///
/// // Files are decompressed on first access and cached
/// let content = archive.read("word/document.xml")?;
///
/// // Subsequent reads return cached data (no re-decompression)
/// let content2 = archive.read("word/document.xml")?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct LazyArchiveReader<'data> {
    /// The underlying archive reader (for decompression)
    inner: ArchiveReader<'data>,
    /// Thread-safe cache of decompressed files
    cache: std::sync::RwLock<HashMap<String, std::sync::Arc<Vec<u8>>>>,
}

impl<'data> LazyArchiveReader<'data> {
    /// Create a new lazy archive reader from a byte slice.
    pub fn new(data: &'data [u8]) -> Result<Self, Error> {
        Self::new_with_limits(data, ArchiveLimits::default())
    }

    /// Create a lazy reader with explicit resource limits.
    pub fn new_with_limits(data: &'data [u8], limits: ArchiveLimits) -> Result<Self, Error> {
        let inner = ArchiveReader::new_with_limits(data, limits)?;
        Ok(Self {
            inner,
            cache: std::sync::RwLock::new(HashMap::new()),
        })
    }

    /// Return declared metadata without reading, decompressing, or caching a member.
    #[inline]
    pub fn metadata(&self, name: &str) -> Result<Metadata, Error> {
        self.inner.metadata(name)
    }

    /// Get the number of files in the archive.
    #[inline]
    pub fn len(&self) -> usize {
        self.inner.len()
    }

    /// Check if the archive is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Check if a file exists in the archive.
    #[inline]
    pub fn contains(&self, name: &str) -> bool {
        self.inner.contains(name)
    }

    /// Get an iterator over all file names in the archive.
    pub fn file_names(&self) -> impl Iterator<Item = &str> {
        self.inner.file_names()
    }

    /// Read and decompress a file, using cache if available.
    ///
    /// Returns a cloned Vec for API compatibility. For zero-copy access,
    /// use `read_shared()` which returns an Arc.
    pub fn read(&self, name: &str) -> Result<Vec<u8>, Error> {
        self.read_shared(name).map(|arc| (*arc).clone())
    }

    /// Read and decompress a file, returning a shared reference.
    ///
    /// This is more efficient than `read()` when the same file is accessed
    /// multiple times, as it avoids cloning the decompressed data.
    pub fn read_shared(&self, name: &str) -> Result<std::sync::Arc<Vec<u8>>, Error> {
        let normalized = name.strip_prefix('/').unwrap_or(name);

        // Fast path: check if already cached (read lock)
        {
            let cache = self.cache.read().unwrap();
            if let Some(data) = cache.get(normalized) {
                return Ok(std::sync::Arc::clone(data));
            }
        }

        // Slow path: decompress and cache (write lock)
        let data = self.inner.read(normalized)?;
        let arc = std::sync::Arc::new(data);

        {
            let mut cache = self.cache.write().unwrap();
            // Double-check in case another thread cached it while we were decompressing
            if let Some(existing) = cache.get(normalized) {
                return Ok(std::sync::Arc::clone(existing));
            }
            cache.insert(normalized.to_string(), std::sync::Arc::clone(&arc));
        }

        Ok(arc)
    }

    /// Read multiple files in parallel without caching.
    ///
    /// This is the fastest method for bulk decompression when you need to read
    /// many files at once and don't need caching. Avoids all cloning overhead.
    ///
    /// Results retain the caller-provided order and preserve every member
    /// error. Successful results are not added to the lazy cache.
    #[inline]
    pub fn read_many_parallel<'a>(
        &self,
        names: &'a [&'a str],
    ) -> Vec<(&'a str, Result<Vec<u8>, Error>)> {
        self.read_many_parallel_results(names)
    }

    /// Read multiple files in parallel while preserving individual errors.
    ///
    /// This is intended for parsers that must distinguish a missing or corrupt
    /// required part from a successfully read package.
    pub fn read_many_parallel_results<'a>(
        &self,
        names: &'a [&'a str],
    ) -> Vec<(&'a str, Result<Vec<u8>, Error>)> {
        use rayon::prelude::*;

        names
            .par_iter()
            .map(|name| (*name, self.inner.read(name)))
            .collect()
    }

    /// Read multiple files in parallel with caching.
    ///
    /// This efficiently decompresses multiple files in parallel while still
    /// benefiting from caching. Files already in cache are returned immediately.
    /// Use this when you expect to read the same files multiple times.
    ///
    /// Results retain the caller-provided order and preserve every member
    /// error. Only successful decompressions are cached.
    pub fn read_many_parallel_cached<'a>(
        &self,
        names: &'a [&'a str],
    ) -> Vec<(&'a str, Result<Vec<u8>, Error>)> {
        use rayon::prelude::*;

        names
            .par_iter()
            .map(|name| (*name, self.read(name)))
            .collect()
    }

    /// Read all files in parallel, caching results.
    ///
    /// Results retain physical source order and preserve every member error.
    /// Successful decompressions are cached; failures are never cached.
    pub fn read_all_parallel(&self) -> Vec<(String, Result<Vec<u8>, Error>)> {
        use rayon::prelude::*;

        let names: Vec<&str> = self.inner.file_names().collect();
        names
            .into_par_iter()
            .map(|name| (name.to_string(), self.read(name)))
            .collect()
    }

    /// Get the number of cached files.
    pub fn cache_size(&self) -> usize {
        self.cache.read().unwrap().len()
    }

    /// Clear the decompression cache to free memory.
    pub fn clear_cache(&self) {
        self.cache.write().unwrap().clear();
    }

    /// Take ownership of cached data, consuming the cache.
    ///
    /// Returns all cached files and clears the cache. This is useful when
    /// you want to take ownership of the decompressed data without cloning.
    pub fn take_cache(&self) -> HashMap<String, Vec<u8>> {
        let mut cache = self.cache.write().unwrap();
        let mut result = HashMap::with_capacity(cache.len());
        for (name, arc) in cache.drain() {
            // Try to unwrap the Arc; if there are other references, clone instead
            match std::sync::Arc::try_unwrap(arc) {
                Ok(data) => {
                    result.insert(name, data);
                },
                Err(arc) => {
                    result.insert(name, (*arc).clone());
                },
            }
        }
        result
    }
}

impl std::fmt::Debug for LazyArchiveReader<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("LazyArchiveReader")
            .field("file_count", &self.inner.len())
            .field("cache_size", &self.cache_size())
            .finish()
    }
}

// Ensure LazyArchiveReader is Send + Sync
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<LazyArchiveReader<'static>>();
};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_round_trip_stored() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("test.txt", b"Hello, World!").unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert!(reader.contains("test.txt"));
        assert_eq!(reader.read("test.txt").unwrap(), b"Hello, World!");
    }

    #[test]
    fn test_round_trip_deflated() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_deflated("content.xml", b"<root>Hello</root>")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert!(reader.contains("content.xml"));
        assert_eq!(reader.read("content.xml").unwrap(), b"<root>Hello</root>");
    }

    #[test]
    fn test_multiple_files() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("mimetype", b"application/test")
            .unwrap();
        writer.write_deflated("content.xml", b"<content/>").unwrap();
        writer.write_deflated("styles.xml", b"<styles/>").unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(reader.len(), 3);
        assert_eq!(reader.read("mimetype").unwrap(), b"application/test");
        assert_eq!(reader.read("content.xml").unwrap(), b"<content/>");
        assert_eq!(reader.read("styles.xml").unwrap(), b"<styles/>");
    }

    #[test]
    fn rejects_archives_exceeding_configured_resource_limits() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("first.xml", b"1234").unwrap();
        writer.write_deflated("second.xml", b"5678").unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let file_error = ArchiveReader::new_with_limits(
            &bytes,
            ArchiveLimits {
                max_files: 1,
                ..ArchiveLimits::UNBOUNDED
            },
        )
        .unwrap_err();
        assert_limit(file_error, LimitResource::FileCount, 2, 1);

        let entry_error = ArchiveReader::new_with_limits(
            &bytes,
            ArchiveLimits {
                max_entry_size: 3,
                ..ArchiveLimits::UNBOUNDED
            },
        )
        .unwrap_err();
        assert_limit(entry_error, LimitResource::EntrySize, 4, 3);

        let total_error = ArchiveReader::new_with_limits(
            &bytes,
            ArchiveLimits {
                max_total_size: 7,
                ..ArchiveLimits::UNBOUNDED
            },
        )
        .unwrap_err();
        assert_limit(total_error, LimitResource::TotalSize, 8, 7);

        assert!(ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).is_ok());
        assert!(LazyArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).is_ok());
    }

    #[test]
    fn enforces_member_name_limit_at_the_declared_boundary() {
        let bytes = fixture(&[FixtureEntry::stored(b"name", b"data")]);

        let mut exact = ArchiveLimits::UNBOUNDED;
        exact.max_member_name_bytes = 4;
        assert!(ArchiveReader::new_with_limits(&bytes, exact).is_ok());

        let mut over = exact;
        over.max_member_name_bytes = 3;
        assert_limit(
            ArchiveReader::new_with_limits(&bytes, over).unwrap_err(),
            LimitResource::MemberNameBytes,
            4,
            3,
        );
    }

    #[test]
    fn enforces_aggregate_central_directory_metadata_limit() {
        let bytes = fixture(&[FixtureEntry {
            name: b"a",
            extra: b"xyz",
            comment: b"q",
            compressed_size: 0,
            uncompressed_size: 0,
            data: b"",
        }]);

        let mut exact = ArchiveLimits::UNBOUNDED;
        exact.max_metadata_bytes = 5;
        assert!(ArchiveReader::new_with_limits(&bytes, exact).is_ok());

        let mut over = exact;
        over.max_metadata_bytes = 4;
        assert_limit(
            ArchiveReader::new_with_limits(&bytes, over).unwrap_err(),
            LimitResource::MetadataBytes,
            5,
            4,
        );
    }

    #[test]
    fn enforces_compressed_member_limit_before_data_access() {
        let bytes = fixture(&[FixtureEntry {
            name: b"a",
            extra: b"",
            comment: b"",
            compressed_size: 3,
            uncompressed_size: 0,
            data: b"",
        }]);

        let mut exact = ArchiveLimits::UNBOUNDED;
        exact.max_compressed_size = 3;
        assert!(ArchiveReader::new_with_limits(&bytes, exact).is_ok());

        let mut over = exact;
        over.max_compressed_size = 2;
        assert_limit(
            ArchiveReader::new_with_limits(&bytes, over).unwrap_err(),
            LimitResource::CompressedSize,
            3,
            2,
        );
    }

    #[test]
    fn accepts_exact_and_rejects_over_uncompressed_entry_limits() {
        let bytes = fixture(&[FixtureEntry::stored(b"a", b"abc")]);

        let mut exact = ArchiveLimits::UNBOUNDED;
        exact.max_entry_size = 3;
        assert!(ArchiveReader::new_with_limits(&bytes, exact).is_ok());

        let mut over = exact;
        over.max_entry_size = 2;
        assert_limit(
            ArchiveReader::new_with_limits(&bytes, over).unwrap_err(),
            LimitResource::EntrySize,
            3,
            2,
        );
    }

    #[test]
    fn accepts_exact_and_rejects_over_aggregate_uncompressed_limits() {
        let bytes = fixture(&[
            FixtureEntry::stored(b"a", b"abc"),
            FixtureEntry::stored(b"b", b"wxyz"),
        ]);

        let mut exact = ArchiveLimits::UNBOUNDED;
        exact.max_total_size = 7;
        assert!(ArchiveReader::new_with_limits(&bytes, exact).is_ok());

        let mut over = exact;
        over.max_total_size = 6;
        assert_limit(
            ArchiveReader::new_with_limits(&bytes, over).unwrap_err(),
            LimitResource::TotalSize,
            7,
            6,
        );
    }

    #[test]
    fn accepts_exact_and_rejects_over_file_count_limits() {
        let bytes = fixture(&[
            FixtureEntry::stored(b"a", b""),
            FixtureEntry::stored(b"b", b""),
        ]);

        let mut exact = ArchiveLimits::UNBOUNDED;
        exact.max_files = 2;
        assert!(ArchiveReader::new_with_limits(&bytes, exact).is_ok());

        let mut over = exact;
        over.max_files = 1;
        assert_limit(
            ArchiveReader::new_with_limits(&bytes, over).unwrap_err(),
            LimitResource::FileCount,
            2,
            1,
        );
    }

    #[test]
    fn directories_consume_metadata_but_not_payload_or_file_budgets() {
        let bytes = fixture(&[FixtureEntry {
            name: b"folder/",
            extra: b"",
            comment: b"",
            compressed_size: u32::MAX,
            uncompressed_size: u32::MAX,
            data: b"",
        }]);
        let limits = ArchiveLimits {
            max_files: 0,
            max_member_name_bytes: 7,
            max_metadata_bytes: 7,
            max_compressed_size: 0,
            max_entry_size: 0,
            max_total_size: 0,
        };

        assert!(ArchiveReader::new_with_limits(&bytes, limits).is_ok());
    }

    #[test]
    fn rejects_aggregate_uncompressed_size_overflow() {
        let zip64 = zip64_sizes(u64::MAX, 0);
        let bytes = fixture(&[
            FixtureEntry {
                name: b"a",
                extra: &zip64,
                comment: b"",
                compressed_size: u32::MAX,
                uncompressed_size: u32::MAX,
                data: b"",
            },
            FixtureEntry {
                name: b"b",
                extra: &zip64,
                comment: b"",
                compressed_size: u32::MAX,
                uncompressed_size: u32::MAX,
                data: b"",
            },
        ]);

        let error = ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).unwrap_err();
        assert!(
            matches!(error.kind(), ErrorKind::InvalidInput { msg } if msg.contains("overflows"))
        );
    }

    #[test]
    fn rejects_malformed_central_directory_variable_declarations() {
        let mut bytes = fixture(&[FixtureEntry::stored(b"a", b"")]);
        let end = bytes.len() - 22;
        let central_directory_offset =
            u32::from_le_bytes(bytes[end + 16..end + 20].try_into().unwrap()) as usize;
        bytes[central_directory_offset + 28..central_directory_offset + 30]
            .copy_from_slice(&2u16.to_le_bytes());

        assert!(ArchiveReader::new(&bytes).is_err());
    }

    #[test]
    fn metadata_lookup_uses_only_the_central_directory_index() {
        let bytes = fixture(&[
            FixtureEntry {
                name: b"body.xml",
                extra: b"",
                comment: b"",
                compressed_size: 3,
                uncompressed_size: 5,
                data: b"",
            },
            FixtureEntry {
                name: b"assets/",
                extra: b"",
                comment: b"",
                compressed_size: u32::MAX,
                uncompressed_size: u32::MAX,
                data: b"",
            },
        ]);

        let reader = ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).unwrap();
        let file = reader.metadata("/body.xml").unwrap();
        assert_eq!(file.compressed_size(), 3);
        assert_eq!(file.uncompressed_size(), 5);
        assert!(!file.is_directory());

        let directory = reader.metadata("assets/").unwrap();
        assert_eq!(directory.compressed_size(), u64::from(u32::MAX));
        assert_eq!(directory.uncompressed_size(), u64::from(u32::MAX));
        assert!(directory.is_directory());
        assert!(
            matches!(reader.metadata("missing"), Err(error) if matches!(error.kind(), ErrorKind::FileNotFound(_)))
        );

        let lazy = LazyArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).unwrap();
        assert_eq!(lazy.cache_size(), 0);
        assert_eq!(lazy.metadata("body.xml").unwrap(), file);
        assert_eq!(lazy.cache_size(), 0);
    }

    #[test]
    fn eager_bulk_reads_preserve_source_order_and_all_member_errors() {
        let mut bytes = bulk_fixture();
        corrupt_payload(&mut bytes, b"bad");
        let reader = ArchiveReader::new(&bytes).unwrap();

        let requested = ["last", "bad", "missing", "first"];
        let results = reader.read_many_parallel(&requested);
        assert_eq!(results.len(), requested.len());
        assert_eq!(*results[0].0, "last");
        assert_eq!(results[0].1.as_ref().unwrap(), b"last");
        assert_eq!(*results[1].0, "bad");
        assert!(
            matches!(results[1].1, Err(ref error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. }))
        );
        assert_eq!(*results[2].0, "missing");
        assert!(
            matches!(results[2].1, Err(ref error) if matches!(error.kind(), ErrorKind::FileNotFound(_)))
        );
        assert_eq!(*results[3].0, "first");
        assert_eq!(results[3].1.as_ref().unwrap(), b"first");

        let all = reader.read_all_parallel();
        assert_eq!(
            all.iter()
                .map(|(name, _)| name.as_str())
                .collect::<Vec<_>>(),
            ["first", "bad", "last"]
        );
        assert!(
            matches!(all[1].1, Err(ref error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. }))
        );
    }

    #[test]
    fn lazy_bulk_reads_propagate_errors_and_cache_only_successes() {
        let mut bytes = bulk_fixture();
        corrupt_payload(&mut bytes, b"bad");
        let reader = LazyArchiveReader::new(&bytes).unwrap();
        let requested = ["last", "bad", "missing", "first"];

        let results = reader.read_many_parallel_cached(&requested);
        assert_eq!(results.len(), requested.len());
        assert_eq!(results[0].1.as_ref().unwrap(), b"last");
        assert!(
            matches!(results[1].1, Err(ref error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. }))
        );
        assert!(
            matches!(results[2].1, Err(ref error) if matches!(error.kind(), ErrorKind::FileNotFound(_)))
        );
        assert_eq!(results[3].1.as_ref().unwrap(), b"first");
        assert_eq!(reader.cache_size(), 2);

        let all = reader.read_all_parallel();
        assert_eq!(
            all.iter()
                .map(|(name, _)| name.as_str())
                .collect::<Vec<_>>(),
            ["first", "bad", "last"]
        );
        assert!(
            matches!(all[1].1, Err(ref error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. }))
        );
        assert_eq!(reader.cache_size(), 2);
    }

    fn assert_limit(error: Error, resource: LimitResource, actual: u64, maximum: u64) {
        match error.kind() {
            ErrorKind::LimitExceeded {
                resource: found_resource,
                actual: found_actual,
                maximum: found_maximum,
            } => {
                assert_eq!(
                    (*found_resource, *found_actual, *found_maximum),
                    (resource, actual, maximum)
                );
            },
            other => panic!("expected limit error, got {other:?}"),
        }
    }

    #[derive(Clone, Copy)]
    struct FixtureEntry<'a> {
        name: &'a [u8],
        extra: &'a [u8],
        comment: &'a [u8],
        compressed_size: u32,
        uncompressed_size: u32,
        data: &'a [u8],
    }

    impl<'a> FixtureEntry<'a> {
        fn stored(name: &'a [u8], data: &'a [u8]) -> Self {
            let size = u32::try_from(data.len()).unwrap();
            Self {
                name,
                extra: b"",
                comment: b"",
                compressed_size: size,
                uncompressed_size: size,
                data,
            }
        }
    }

    fn fixture(entries: &[FixtureEntry<'_>]) -> Vec<u8> {
        let mut archive = Vec::new();
        let mut central_directory = Vec::new();

        for entry in entries {
            let local_header_offset = u32::try_from(archive.len()).unwrap();
            push_u32(&mut archive, 0x0403_4b50);
            push_u16(&mut archive, 20);
            push_u16(&mut archive, 0);
            push_u16(&mut archive, 0);
            push_u16(&mut archive, 0);
            push_u16(&mut archive, 0);
            push_u32(&mut archive, 0);
            push_u32(&mut archive, entry.compressed_size);
            push_u32(&mut archive, entry.uncompressed_size);
            push_u16(&mut archive, u16::try_from(entry.name.len()).unwrap());
            push_u16(&mut archive, 0);
            archive.extend_from_slice(entry.name);
            archive.extend_from_slice(entry.data);

            push_u32(&mut central_directory, 0x0201_4b50);
            push_u16(&mut central_directory, 20);
            push_u16(&mut central_directory, 20);
            push_u16(&mut central_directory, 0);
            push_u16(&mut central_directory, 0);
            push_u16(&mut central_directory, 0);
            push_u16(&mut central_directory, 0);
            push_u32(&mut central_directory, 0);
            push_u32(&mut central_directory, entry.compressed_size);
            push_u32(&mut central_directory, entry.uncompressed_size);
            push_u16(
                &mut central_directory,
                u16::try_from(entry.name.len()).unwrap(),
            );
            push_u16(
                &mut central_directory,
                u16::try_from(entry.extra.len()).unwrap(),
            );
            push_u16(
                &mut central_directory,
                u16::try_from(entry.comment.len()).unwrap(),
            );
            push_u16(&mut central_directory, 0);
            push_u16(&mut central_directory, 0);
            push_u32(&mut central_directory, 0);
            push_u32(&mut central_directory, local_header_offset);
            central_directory.extend_from_slice(entry.name);
            central_directory.extend_from_slice(entry.extra);
            central_directory.extend_from_slice(entry.comment);
        }

        let central_directory_offset = u32::try_from(archive.len()).unwrap();
        let central_directory_size = u32::try_from(central_directory.len()).unwrap();
        archive.extend_from_slice(&central_directory);
        push_u32(&mut archive, 0x0605_4b50);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        let count = u16::try_from(entries.len()).unwrap();
        push_u16(&mut archive, count);
        push_u16(&mut archive, count);
        push_u32(&mut archive, central_directory_size);
        push_u32(&mut archive, central_directory_offset);
        push_u16(&mut archive, 0);
        archive
    }

    fn zip64_sizes(uncompressed_size: u64, compressed_size: u64) -> Vec<u8> {
        let mut extra = Vec::new();
        push_u16(&mut extra, 1);
        push_u16(&mut extra, 16);
        extra.extend_from_slice(&uncompressed_size.to_le_bytes());
        extra.extend_from_slice(&compressed_size.to_le_bytes());
        extra
    }

    fn push_u16(output: &mut Vec<u8>, value: u16) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn push_u32(output: &mut Vec<u8>, value: u32) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn bulk_fixture() -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("first", b"first").unwrap();
        writer.write_stored("bad", b"bad").unwrap();
        writer.write_stored("last", b"last").unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn corrupt_payload(archive: &mut [u8], payload: &[u8]) {
        let offsets: Vec<usize> = archive
            .windows(payload.len())
            .enumerate()
            .filter_map(|(offset, candidate)| (candidate == payload).then_some(offset))
            .collect();
        archive[offsets[1]] ^= 0x80;
    }
}
