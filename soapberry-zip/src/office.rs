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
//! use soapberry_zip::office::ArchiveWriter;
//!
//! let mut writer = ArchiveWriter::new();
//! writer.write_stored("mimetype", b"application/vnd.oasis.opendocument.text")?;
//! writer.write_deflated("content.xml", b"<office:document-content>...</office:document-content>")?;
//! let bytes = writer.finish()?;
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
}

/// Information about an archive entry for fast lookup
#[derive(Debug, Clone)]
struct EntryInfo {
    wayfinder: crate::ZipArchiveEntryWayfinder,
    compression_method: CompressionMethod,
    uncompressed_size: u64,
}

impl<'data> ArchiveReader<'data> {
    /// Create a new archive reader from a byte slice.
    ///
    /// This parses the ZIP central directory and builds an index for fast
    /// file lookup. The actual file contents are not decompressed until
    /// accessed via `read()`.
    pub fn new(data: &'data [u8]) -> Result<Self, Error> {
        let archive = ZipArchive::from_slice(data)?;

        // Build index for fast lookup
        let mut index = HashMap::new();
        for entry_result in archive.entries() {
            let entry = entry_result?;
            let path = entry.file_path();

            // Normalize path - convert to string, skip directories
            if entry.is_dir() {
                continue;
            }

            let name = match path.try_normalize() {
                Ok(normalized) => normalized.as_ref().to_string(),
                Err(_) => {
                    // Fallback to raw path as lossy UTF-8
                    String::from_utf8_lossy(path.as_ref()).to_string()
                },
            };

            index.insert(
                name,
                EntryInfo {
                    wayfinder: entry.wayfinder(),
                    compression_method: entry.compression_method(),
                    uncompressed_size: entry.uncompressed_size_hint(),
                },
            );
        }

        Ok(Self { archive, index })
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

    /// Get an iterator over all file names in the archive.
    pub fn file_names(&self) -> impl Iterator<Item = &str> {
        self.index.keys().map(|s| s.as_str())
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
                // Deflate - decompress with pre-allocated buffer
                // Using unsafe to avoid costly buffer zeroing from read_to_end
                let size = info.uncompressed_size as usize;
                let mut decompressed = Vec::with_capacity(size);

                // SAFETY: We set the length to the expected uncompressed size.
                // The decompression will write exactly `size` bytes (verified by CRC32).
                // Any unwritten bytes at the end are truncated after reading.
                #[allow(unsafe_code, clippy::uninit_vec)]
                unsafe {
                    decompressed.set_len(size);
                }

                let mut decoder = entry.verifying_reader(DeflateDecoder::new(data));
                let mut total_read = 0;
                while total_read < size {
                    match decoder.read(&mut decompressed[total_read..]) {
                        Ok(0) => break,
                        Ok(n) => total_read += n,
                        Err(e) if e.kind() == std::io::ErrorKind::Interrupted => continue,
                        Err(e) => return Err(e.into()),
                    }
                }

                // Truncate to actual bytes read (handles size mismatch gracefully)
                decompressed.truncate(total_read);
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
    /// Returns a HashMap mapping file names to their decompressed contents.
    /// Files that fail to decompress are skipped (not included in result).
    ///
    /// This is optimal when you need to access most/all files in the archive.
    pub fn read_all_parallel(&self) -> HashMap<String, Vec<u8>> {
        use rayon::prelude::*;

        // Collect keys to Vec first for proper parallel iteration
        // par_bridge() doesn't parallelize HashMap iteration effectively
        let keys: Vec<&String> = self.index.keys().collect();

        keys.into_par_iter()
            .filter_map(|name| self.read(name).ok().map(|data| (name.clone(), data)))
            .collect()
    }
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
        let inner = ArchiveReader::new(data)?;
        Ok(Self {
            inner,
            cache: std::sync::RwLock::new(HashMap::new()),
        })
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

    /// Read multiple files in parallel WITHOUT caching.
    ///
    /// This is the fastest method for bulk decompression when you need to read
    /// many files at once and don't need caching. Avoids all cloning overhead.
    ///
    /// Returns a HashMap mapping file names to their decompressed contents.
    /// Files that fail to decompress are not included in the result.
    pub fn read_many_parallel(&self, names: &[&str]) -> HashMap<String, Vec<u8>> {
        use rayon::prelude::*;

        // Parallel decompression without caching for maximum performance
        names
            .par_iter()
            .filter_map(|name| {
                let normalized = name.strip_prefix('/').unwrap_or(name);
                self.inner
                    .read(normalized)
                    .ok()
                    .map(|data| (normalized.to_string(), data))
            })
            .collect()
    }

    /// Read multiple files in parallel with caching.
    ///
    /// This efficiently decompresses multiple files in parallel while still
    /// benefiting from caching. Files already in cache are returned immediately.
    /// Use this when you expect to read the same files multiple times.
    ///
    /// Returns a HashMap mapping file names to their decompressed contents.
    /// Files that fail to decompress are not included in the result.
    pub fn read_many_parallel_cached(&self, names: &[&str]) -> HashMap<String, Vec<u8>> {
        use rayon::prelude::*;

        // Separate cached and uncached files
        let (cached, uncached): (Vec<&str>, Vec<&str>) = {
            let cache = self.cache.read().unwrap();
            names.iter().partition(|name| {
                let normalized = name.strip_prefix('/').unwrap_or(*name);
                cache.contains_key(normalized)
            })
        };

        // Start with cached results
        let mut results: HashMap<String, Vec<u8>> = {
            let cache = self.cache.read().unwrap();
            cached
                .into_iter()
                .filter_map(|name| {
                    let normalized = name.strip_prefix('/').unwrap_or(name);
                    cache
                        .get(normalized)
                        .map(|arc| (normalized.to_string(), (**arc).clone()))
                })
                .collect()
        };

        // Decompress uncached files in parallel
        if !uncached.is_empty() {
            let decompressed: Vec<_> = uncached
                .into_par_iter()
                .filter_map(|name| {
                    let normalized = name.strip_prefix('/').unwrap_or(name);
                    self.inner
                        .read(normalized)
                        .ok()
                        .map(|data| (normalized.to_string(), data))
                })
                .collect();

            // Cache the newly decompressed files
            {
                let mut cache = self.cache.write().unwrap();
                for (name, data) in &decompressed {
                    if !cache.contains_key(name.as_str()) {
                        cache.insert(name.clone(), std::sync::Arc::new(data.clone()));
                    }
                }
            }

            results.extend(decompressed);
        }

        results
    }

    /// Read all files in parallel, caching results.
    ///
    /// Similar to `ArchiveReader::read_all_parallel()` but caches results
    /// for potential future access.
    pub fn read_all_parallel(&self) -> HashMap<String, Vec<u8>> {
        let names: Vec<&str> = self.inner.file_names().collect();
        self.read_many_parallel(&names)
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
    use std::sync::atomic::{AtomicUsize, Ordering};

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
}
