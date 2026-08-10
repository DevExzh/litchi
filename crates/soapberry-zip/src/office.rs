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

use crate::path::{RawPath, ZipFilePath};
use crate::{
    CompressionMethod, Error, ErrorKind, RECOMMENDED_BUFFER_SIZE, ReaderAt, ZipArchive,
    ZipArchiveWriter, ZipLocator, ZipSliceArchive, ZipVerification,
};
use flate2::Compression;
use flate2::read::DeflateDecoder;
use flate2::write::DeflateEncoder;
use rayon::prelude::*;
use std::collections::HashMap;
use std::io::{Read, Write};
use std::num::{NonZeroU64, NonZeroUsize};

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

/// CPU-affinity policy for the local workers owned by a [`ParallelReadSession`].
///
/// The archive substrate currently supports only inheriting operating-system
/// placement. A caller must select this policy explicitly when constructing
/// [`ParallelReadLimits`]; no global scheduler or affinity policy is inferred.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ParallelAffinity {
    /// Do not change operating-system worker affinity.
    Inherit,
}

/// Validated finite limits for a local [`ParallelReadSession`].
///
/// The task and byte caps bound one submitted batch. A batch below
/// `min_parallel_bytes` is read serially even when the session owns more than
/// one worker, avoiding small-task scheduling overhead.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParallelReadLimits {
    workers: NonZeroUsize,
    max_in_flight_tasks: NonZeroUsize,
    max_in_flight_bytes: NonZeroU64,
    min_parallel_bytes: u64,
    affinity: ParallelAffinity,
}

impl ParallelReadLimits {
    /// Creates limits with [`ParallelAffinity::Inherit`].
    ///
    /// # Errors
    ///
    /// Returns an error when the worker count exceeds the task cap or when the
    /// parallel-work threshold exceeds the finite byte cap.
    pub fn new(
        workers: NonZeroUsize,
        max_in_flight_tasks: NonZeroUsize,
        max_in_flight_bytes: NonZeroU64,
        min_parallel_bytes: u64,
    ) -> Result<Self, Error> {
        Self::with_affinity(
            workers,
            max_in_flight_tasks,
            max_in_flight_bytes,
            min_parallel_bytes,
            ParallelAffinity::Inherit,
        )
    }

    /// Creates limits with an explicit affinity policy.
    ///
    /// # Errors
    ///
    /// Returns an error when the worker count exceeds the task cap or when the
    /// parallel-work threshold exceeds the finite byte cap.
    pub fn with_affinity(
        workers: NonZeroUsize,
        max_in_flight_tasks: NonZeroUsize,
        max_in_flight_bytes: NonZeroU64,
        min_parallel_bytes: u64,
        affinity: ParallelAffinity,
    ) -> Result<Self, Error> {
        if workers > max_in_flight_tasks {
            return Err(ErrorKind::InvalidParallelReadLimits {
                reason: "workers must not exceed max_in_flight_tasks",
            }
            .into());
        }
        if min_parallel_bytes > max_in_flight_bytes.get() {
            return Err(ErrorKind::InvalidParallelReadLimits {
                reason: "min_parallel_bytes must not exceed max_in_flight_bytes",
            }
            .into());
        }
        Ok(Self {
            workers,
            max_in_flight_tasks,
            max_in_flight_bytes,
            min_parallel_bytes,
            affinity,
        })
    }

    /// Maximum workers the local session may create.
    #[must_use]
    pub const fn workers(self) -> NonZeroUsize {
        self.workers
    }

    /// Maximum tasks in one submitted batch.
    #[must_use]
    pub const fn max_in_flight_tasks(self) -> NonZeroUsize {
        self.max_in_flight_tasks
    }

    /// Maximum declared uncompressed bytes in one submitted batch.
    #[must_use]
    pub const fn max_in_flight_bytes(self) -> NonZeroU64 {
        self.max_in_flight_bytes
    }

    /// Smallest batch size eligible for parallel execution.
    #[must_use]
    pub const fn min_parallel_bytes(self) -> u64 {
        self.min_parallel_bytes
    }

    /// Explicit worker-affinity policy.
    #[must_use]
    pub const fn affinity(self) -> ParallelAffinity {
        self.affinity
    }
}

/// Cooperative cancellation probe used by an explicit parallel read.
///
/// The probe is checked before scheduling, between batches, before each member
/// read, and after each member read. A currently-running decompressor is not
/// forcefully interrupted; cancellation is therefore member-granular.
pub trait CancellationProbe: Send + Sync {
    /// Returns whether the operation should stop at its next interruption point.
    fn is_cancelled(&self) -> bool;
}

impl<F> CancellationProbe for F
where
    F: Fn() -> bool + Send + Sync,
{
    fn is_cancelled(&self) -> bool {
        self()
    }
}

/// Reusable local scheduler for explicit archive bulk reads.
///
/// The session owns a Rayon pool created with the requested worker count. It
/// never initializes or installs Rayon’s process-global pool. A one-worker
/// session uses the same bounded batching policy but executes serially.
pub struct ParallelReadSession {
    limits: ParallelReadLimits,
    pool: Option<rayon::ThreadPool>,
}

impl ParallelReadSession {
    /// Creates a reusable local scheduler with validated finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the local Rayon worker pool cannot be created.
    pub fn new(limits: ParallelReadLimits) -> Result<Self, Error> {
        let workers = limits.workers().get();
        let pool = if workers == 1 {
            None
        } else {
            Some(
                rayon::ThreadPoolBuilder::new()
                    .num_threads(workers)
                    .build()
                    .map_err(|error| ErrorKind::ParallelReadWorkerPool {
                        workers,
                        message: error.to_string(),
                    })?,
            )
        };
        Ok(Self { limits, pool })
    }

    /// Validated policy used by this session.
    #[must_use]
    pub const fn limits(&self) -> ParallelReadLimits {
        self.limits
    }

    /// Explicit worker count requested for this local session.
    #[must_use]
    pub const fn worker_count(&self) -> NonZeroUsize {
        self.limits.workers()
    }

    fn read_many<'name, MetadataFor, ReadMember>(
        &self,
        names: &'name [&'name str],
        cancellation: &dyn CancellationProbe,
        metadata_for: MetadataFor,
        read_member: ReadMember,
    ) -> Result<Vec<(&'name str, Result<Vec<u8>, Error>)>, Error>
    where
        MetadataFor: Fn(&str) -> Result<Metadata, Error> + Sync,
        ReadMember: Fn(&str) -> Result<Vec<u8>, Error> + Sync,
    {
        self.check_cancelled(cancellation)?;
        let mut results = Vec::new();
        results.try_reserve(names.len()).map_err(|error| {
            Error::from(ErrorKind::InvalidInput {
                msg: format!("could not reserve parallel read results: {error}"),
            })
        })?;

        let mut batch = Vec::new();
        let mut batch_bytes = 0_u64;
        for name in names {
            self.check_cancelled(cancellation)?;
            let metadata = match metadata_for(name) {
                Ok(metadata) => metadata,
                Err(error) => {
                    self.flush_batch(
                        &mut results,
                        &mut batch,
                        &mut batch_bytes,
                        cancellation,
                        &read_member,
                    )?;
                    results.push((*name, Err(error)));
                    continue;
                },
            };
            let bytes = metadata.uncompressed_size();
            if bytes > self.limits.max_in_flight_bytes().get() {
                return Err(ErrorKind::ParallelReadInFlightBytesExceeded {
                    actual: bytes,
                    maximum: self.limits.max_in_flight_bytes().get(),
                }
                .into());
            }
            let exceeds_task_cap = batch.len() == self.limits.max_in_flight_tasks().get();
            let next_bytes = batch_bytes.checked_add(bytes).ok_or_else(|| {
                Error::from(ErrorKind::ParallelReadInFlightBytesExceeded {
                    actual: u64::MAX,
                    maximum: self.limits.max_in_flight_bytes().get(),
                })
            })?;
            if !batch.is_empty()
                && (next_bytes > self.limits.max_in_flight_bytes().get() || exceeds_task_cap)
            {
                self.flush_batch(
                    &mut results,
                    &mut batch,
                    &mut batch_bytes,
                    cancellation,
                    &read_member,
                )?;
            }
            batch_bytes = batch_bytes.checked_add(bytes).ok_or_else(|| {
                Error::from(ErrorKind::ParallelReadInFlightBytesExceeded {
                    actual: u64::MAX,
                    maximum: self.limits.max_in_flight_bytes().get(),
                })
            })?;
            batch.push(*name);
        }
        self.flush_batch(
            &mut results,
            &mut batch,
            &mut batch_bytes,
            cancellation,
            &read_member,
        )?;
        Ok(results)
    }

    fn flush_batch<'name, ReadMember>(
        &self,
        results: &mut Vec<(&'name str, Result<Vec<u8>, Error>)>,
        batch: &mut Vec<&'name str>,
        batch_bytes: &mut u64,
        cancellation: &dyn CancellationProbe,
        read_member: &ReadMember,
    ) -> Result<(), Error>
    where
        ReadMember: Fn(&str) -> Result<Vec<u8>, Error> + Sync,
    {
        if batch.is_empty() {
            return Ok(());
        }
        self.check_cancelled(cancellation)?;
        let parallel = batch.len() > 1 && *batch_bytes >= self.limits.min_parallel_bytes();
        let batch = std::mem::take(batch);
        *batch_bytes = 0;
        let results_for_batch: Vec<(&'name str, Result<Vec<u8>, Error>)> = match self.pool.as_ref()
        {
            Some(pool) if parallel => pool.install(|| {
                batch
                    .par_iter()
                    .map(|name| (*name, self.read_member(name, cancellation, read_member)))
                    .collect()
            }),
            Some(_) | None => batch
                .into_iter()
                .map(|name| (name, self.read_member(name, cancellation, read_member)))
                .collect(),
        };
        if cancellation.is_cancelled()
            || results_for_batch.iter().any(|(_, result)| {
                matches!(result, Err(error) if matches!(error.kind(), ErrorKind::Cancelled))
            })
        {
            return Err(cancelled_error());
        }
        results.extend(results_for_batch);
        Ok(())
    }

    fn read_member<ReadMember>(
        &self,
        name: &str,
        cancellation: &dyn CancellationProbe,
        read_member: &ReadMember,
    ) -> Result<Vec<u8>, Error>
    where
        ReadMember: Fn(&str) -> Result<Vec<u8>, Error> + Sync,
    {
        self.check_cancelled(cancellation)?;
        let result = read_member(name);
        self.check_cancelled(cancellation)?;
        result
    }

    fn check_cancelled(&self, cancellation: &dyn CancellationProbe) -> Result<(), Error> {
        if cancellation.is_cancelled() {
            Err(cancelled_error())
        } else {
            Ok(())
        }
    }
}

impl std::fmt::Debug for ParallelReadSession {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ParallelReadSession")
            .field("limits", &self.limits)
            .field("uses_local_pool", &self.pool.is_some())
            .finish()
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

/// Opaque identifier for one non-directory member in an [`IndexedArchive`].
///
/// An ID is stable for the lifetime of the archive that produced it. Its
/// representation is intentionally private so callers cannot manufacture an
/// unchecked physical entry reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct EntryId(usize);

/// One validated, positionally-readable ZIP archive index.
///
/// Unlike [`ArchiveReader`], this type is not restricted to a contiguous byte
/// slice. It owns an already-located [`ZipArchive`] and scans its central
/// directory exactly once under [`ArchiveLimits`]. Member contents remain
/// unread until [`Self::read`] or [`Self::read_entry`] is called.
///
/// The type deliberately has no payload cache and never uses implicit global
/// scheduling. Callers that opt into bounded local parallelism use
/// [`Self::read_many_with_session`] or [`Self::read_all_with_session`].
pub struct IndexedArchive<R> {
    archive: ZipArchive<R>,
    index: HashMap<String, EntryId>,
    entries: Vec<IndexedEntry>,
    directories: HashMap<String, Metadata>,
    order: Vec<EntryId>,
}

#[derive(Debug, Clone)]
struct IndexedEntry {
    name: String,
    info: EntryInfo,
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

    /// Reads multiple members through an explicit local [`ParallelReadSession`].
    ///
    /// Results retain caller input order. Cancellation returns one outer error
    /// and discards every successful member result from the interrupted call.
    pub fn read_many_with_session<'name>(
        &self,
        session: &ParallelReadSession,
        names: &'name [&'name str],
        cancellation: &dyn CancellationProbe,
    ) -> Result<Vec<(&'name str, Result<Vec<u8>, Error>)>, Error> {
        session.read_many(
            names,
            cancellation,
            |name| self.metadata(name),
            |name| self.read(name),
        )
    }

    /// Reads every member through an explicit local [`ParallelReadSession`].
    ///
    /// Results retain physical source order. Cancellation discards all results
    /// from the interrupted call.
    pub fn read_all_with_session(
        &self,
        session: &ParallelReadSession,
        cancellation: &dyn CancellationProbe,
    ) -> Result<Vec<(String, Result<Vec<u8>, Error>)>, Error> {
        let names = self.file_names().collect::<Vec<_>>();
        self.read_many_with_session(session, &names, cancellation)
            .map(|results| {
                results
                    .into_iter()
                    .map(|(name, result)| (name.to_string(), result))
                    .collect()
            })
    }

    /// Reads multiple members serially.
    ///
    /// This compatibility method no longer uses Rayon’s global pool. Create a
    /// [`ParallelReadSession`] and call [`Self::read_many_with_session`] to
    /// request bounded local parallelism.
    #[deprecated(
        since = "0.0.1",
        note = "this compatibility method is serial; use ParallelReadSession with read_many_with_session"
    )]
    pub fn read_many_parallel<'a, S: AsRef<str> + Sync>(
        &self,
        names: &'a [S],
    ) -> Vec<(&'a S, Result<Vec<u8>, Error>)> {
        names
            .iter()
            .map(|name| (name, self.read(name.as_ref())))
            .collect()
    }

    /// Reads all members serially.
    ///
    /// This compatibility method no longer uses Rayon’s global pool. Create a
    /// [`ParallelReadSession`] and call [`Self::read_all_with_session`] to
    /// request bounded local parallelism.
    #[deprecated(
        since = "0.0.1",
        note = "this compatibility method is serial; use ParallelReadSession with read_all_with_session"
    )]
    pub fn read_all_parallel(&self) -> Vec<(String, Result<Vec<u8>, Error>)> {
        self.order
            .iter()
            .map(|name| (name.clone(), self.read(name)))
            .collect()
    }
}

impl<R> IndexedArchive<R>
where
    R: ReaderAt,
{
    /// Locate and index a positional ZIP source with default resource limits.
    ///
    /// `end_offset` is the exclusive source length used by the ZIP locator.
    /// Call [`Self::from_zip_archive_with_limits`] when the caller has already
    /// located the archive and wants to avoid another EOCD search.
    pub fn from_reader(reader: R, end_offset: u64) -> Result<Self, Error> {
        Self::from_reader_with_limits(reader, end_offset, ArchiveLimits::default())
    }

    /// Locate and index a positional ZIP source with explicit resource limits.
    ///
    /// The central directory is located and scanned once. Payload bytes are not
    /// read or decompressed during construction.
    pub fn from_reader_with_limits(
        reader: R,
        end_offset: u64,
        limits: ArchiveLimits,
    ) -> Result<Self, Error> {
        let mut buffer = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipLocator::new()
            .locate_in_reader(reader, &mut buffer, end_offset)
            .map_err(|(_reader, error)| error)?;
        Self::from_zip_archive_with_limits(archive, limits)
    }

    /// Build an index from an already located ZIP archive using default limits.
    ///
    /// This is the preferred constructor for callers that retain one validated
    /// positional ZIP locator result as their physical-package state.
    pub fn from_zip_archive(archive: ZipArchive<R>) -> Result<Self, Error> {
        Self::from_zip_archive_with_limits(archive, ArchiveLimits::default())
    }

    /// Build an index from an already located ZIP archive with explicit limits.
    ///
    /// Every central-directory entry is validated and classified exactly once.
    /// Directories are retained only for metadata lookup; non-directory entries
    /// receive stable opaque [`EntryId`] values.
    pub fn from_zip_archive_with_limits(
        archive: ZipArchive<R>,
        limits: ArchiveLimits,
    ) -> Result<Self, Error> {
        let mut index = HashMap::new();
        let mut entries = Vec::new();
        let mut directories = HashMap::new();
        let mut ordered_entries = Vec::new();
        let mut total_metadata_bytes = 0_u64;
        let mut total_uncompressed_size = 0_u64;
        let mut buffer = vec![0_u8; RECOMMENDED_BUFFER_SIZE];

        {
            let mut central_entries = archive.entries(&mut buffer);
            while let Some(entry) = central_entries.next_entry()? {
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
                            msg: "archive central-directory metadata total overflows u64"
                                .to_string(),
                        })
                    })?;
                if total_metadata_bytes > limits.max_metadata_bytes {
                    return Err(limit_error(
                        LimitResource::MetadataBytes,
                        total_metadata_bytes,
                        limits.max_metadata_bytes,
                    ));
                }

                let compressed_size = entry.compressed_size_hint();
                let uncompressed_size = entry.uncompressed_size_hint();
                let name = normalized_member_name(path);

                if entry.is_dir() {
                    directories.entry(name).or_insert(Metadata {
                        compressed_size,
                        uncompressed_size,
                        directory: true,
                    });
                    continue;
                }

                if entries.len() >= limits.max_files {
                    let actual = (entries.len() as u64).checked_add(1).ok_or_else(|| {
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
                if compressed_size > limits.max_compressed_size {
                    return Err(limit_error(
                        LimitResource::CompressedSize,
                        compressed_size,
                        limits.max_compressed_size,
                    ));
                }
                if uncompressed_size > limits.max_entry_size {
                    return Err(limit_error(
                        LimitResource::EntrySize,
                        uncompressed_size,
                        limits.max_entry_size,
                    ));
                }
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

                let entry_id = EntryId(entries.len());
                if index.insert(name.clone(), entry_id).is_some() {
                    return Err(ErrorKind::InvalidInput {
                        msg: "archive contains duplicate normalized file names".to_string(),
                    }
                    .into());
                }
                ordered_entries.push((entry.local_header_offset(), entry_id));
                entries.push(IndexedEntry {
                    name,
                    info: EntryInfo {
                        wayfinder: entry.wayfinder(),
                        compression_method: entry.compression_method(),
                        uncompressed_size,
                    },
                });
            }
        }

        ordered_entries.sort_unstable_by_key(|(offset, _)| *offset);
        let order = ordered_entries.into_iter().map(|(_, id)| id).collect();

        Ok(Self {
            archive,
            index,
            entries,
            directories,
            order,
        })
    }

    /// Number of indexed non-directory members.
    #[inline]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether this archive has no indexed non-directory members.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Return whether a normalized member name exists.
    #[inline]
    pub fn contains(&self, name: &str) -> bool {
        self.entry_id(name).is_some()
    }

    /// Resolve a member name to its stable opaque entry ID.
    #[inline]
    pub fn entry_id(&self, name: &str) -> Option<EntryId> {
        let normalized = name.strip_prefix('/').unwrap_or(name);
        self.index.get(normalized).copied()
    }

    /// Return declared metadata for a member without payload access.
    pub fn metadata(&self, name: &str) -> Result<Metadata, Error> {
        match self.entry_id(name) {
            Some(id) => self.metadata_for(id),
            None => {
                let normalized = name.strip_prefix('/').unwrap_or(name);
                self.directories
                    .get(normalized)
                    .copied()
                    .ok_or_else(|| Error::from(ErrorKind::FileNotFound(normalized.to_string())))
            },
        }
    }

    /// Return declared metadata for one indexed file entry.
    pub fn metadata_for(&self, entry_id: EntryId) -> Result<Metadata, Error> {
        let entry = self.indexed_entry(entry_id)?;
        Ok(Metadata {
            compressed_size: entry.info.wayfinder.compressed_size_hint(),
            uncompressed_size: entry.info.uncompressed_size,
            directory: false,
        })
    }

    /// Iterate normalized non-directory names in physical local-header order.
    pub fn file_names(&self) -> impl Iterator<Item = &str> {
        self.order.iter().map(|id| self.entries[id.0].name.as_str())
    }

    /// Whether an indexed file uses ZIP Store compression.
    pub fn is_stored(&self, name: &str) -> Result<bool, Error> {
        let entry_id = self.entry_id(name).ok_or_else(|| {
            let normalized = name.strip_prefix('/').unwrap_or(name);
            Error::from(ErrorKind::FileNotFound(normalized.to_string()))
        })?;
        Ok(self.indexed_entry(entry_id)?.info.compression_method == CompressionMethod::Store)
    }

    /// Read and verify one member by name.
    pub fn read(&self, name: &str) -> Result<Vec<u8>, Error> {
        let entry_id = self.entry_id(name).ok_or_else(|| {
            let normalized = name.strip_prefix('/').unwrap_or(name);
            Error::from(ErrorKind::FileNotFound(normalized.to_string()))
        })?;
        self.read_entry(entry_id)
    }

    /// Read and verify one member by its stable opaque entry ID.
    ///
    /// ZIP local-header, data-descriptor, decompressed-size, and CRC checks are
    /// intentionally deferred until this method is called.
    pub fn read_entry(&self, entry_id: EntryId) -> Result<Vec<u8>, Error> {
        let indexed = self.indexed_entry(entry_id)?;
        let entry = self.archive.get_entry(indexed.info.wayfinder)?;
        let size = usize::try_from(indexed.info.uncompressed_size).map_err(|_| {
            Error::from(ErrorKind::InvalidInput {
                msg: format!(
                    "archive entry size {} does not fit this platform",
                    indexed.info.uncompressed_size
                ),
            })
        })?;
        let mut output = Vec::new();
        output.try_reserve_exact(size).map_err(|error| {
            Error::from(ErrorKind::InvalidInput {
                msg: format!("could not allocate {size} bytes for archive entry: {error}"),
            })
        })?;

        match indexed.info.compression_method {
            CompressionMethod::Store => {
                let reader = entry.verifying_reader(entry.reader());
                reader
                    .take(indexed.info.uncompressed_size.saturating_add(1))
                    .read_to_end(&mut output)?;
            },
            CompressionMethod::Deflate => {
                let decoder = DeflateDecoder::new(entry.reader());
                let reader = entry.verifying_reader(decoder);
                reader
                    .take(indexed.info.uncompressed_size.saturating_add(1))
                    .read_to_end(&mut output)?;
            },
            other => {
                return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                    other.as_id().as_u16(),
                )));
            },
        }

        if output.len() != size {
            return Err(ErrorKind::InvalidSize {
                expected: indexed.info.uncompressed_size,
                actual: output.len() as u64,
            }
            .into());
        }
        Ok(output)
    }

    /// Reads multiple members through an explicit local [`ParallelReadSession`].
    ///
    /// This method is available only for positional sources that are safe to
    /// access concurrently. Results retain caller input order, and
    /// cancellation discards every result from the interrupted call.
    pub fn read_many_with_session<'name>(
        &self,
        session: &ParallelReadSession,
        names: &'name [&'name str],
        cancellation: &dyn CancellationProbe,
    ) -> Result<Vec<(&'name str, Result<Vec<u8>, Error>)>, Error>
    where
        R: Send + Sync,
    {
        session.read_many(
            names,
            cancellation,
            |name| self.metadata(name),
            |name| self.read(name),
        )
    }

    /// Reads every indexed member through an explicit local [`ParallelReadSession`].
    ///
    /// Results retain physical source order, and cancellation discards every
    /// result from the interrupted call.
    pub fn read_all_with_session(
        &self,
        session: &ParallelReadSession,
        cancellation: &dyn CancellationProbe,
    ) -> Result<Vec<(String, Result<Vec<u8>, Error>)>, Error>
    where
        R: Send + Sync,
    {
        let names = self.file_names().collect::<Vec<_>>();
        self.read_many_with_session(session, &names, cancellation)
            .map(|results| {
                results
                    .into_iter()
                    .map(|(name, result)| (name.to_string(), result))
                    .collect()
            })
    }

    /// Consume this index and return the located positional archive.
    #[must_use]
    pub fn into_zip_archive(self) -> ZipArchive<R> {
        self.archive
    }

    fn indexed_entry(&self, entry_id: EntryId) -> Result<&IndexedEntry, Error> {
        self.entries.get(entry_id.0).ok_or_else(|| {
            Error::from(ErrorKind::FileNotFound(format!(
                "unknown indexed ZIP entry {}",
                entry_id.0
            )))
        })
    }
}

impl<R> std::fmt::Debug for IndexedArchive<R> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("IndexedArchive")
            .field("file_count", &self.entries.len())
            .finish()
    }
}

fn normalized_member_name(path: ZipFilePath<RawPath<'_>>) -> String {
    match path.try_normalize() {
        Ok(normalized) => normalized.as_ref().to_string(),
        Err(_) => String::from_utf8_lossy(path.as_ref()).to_string(),
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

#[inline]
fn cancelled_error() -> Error {
    ErrorKind::Cancelled.into()
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

// Ensure ArchiveReader can be borrowed by a local parallel-read session.
// This is a compile-time assertion
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<ArchiveReader<'static>>();
};

/// Lazy ZIP archive reader with on-demand decompression and caching.
///
/// Unlike an explicit bulk-read session, this reader decompresses files on demand.
/// This is optimal for:
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

    /// Reads multiple members through an explicit session without populating the cache.
    ///
    /// Results retain caller input order. Cancellation returns an outer error
    /// and does not publish successful values into this reader's cache.
    pub fn read_many_with_session<'name>(
        &self,
        session: &ParallelReadSession,
        names: &'name [&'name str],
        cancellation: &dyn CancellationProbe,
    ) -> Result<Vec<(&'name str, Result<Vec<u8>, Error>)>, Error> {
        session.read_many(
            names,
            cancellation,
            |name| self.inner.metadata(name),
            |name| self.inner.read(name),
        )
    }

    /// Reads every member through an explicit session without populating the cache.
    ///
    /// Results retain physical source order. Cancellation discards every
    /// result from the interrupted call.
    pub fn read_all_with_session(
        &self,
        session: &ParallelReadSession,
        cancellation: &dyn CancellationProbe,
    ) -> Result<Vec<(String, Result<Vec<u8>, Error>)>, Error> {
        let names = self.inner.file_names().collect::<Vec<_>>();
        self.read_many_with_session(session, &names, cancellation)
            .map(|results| {
                results
                    .into_iter()
                    .map(|(name, result)| (name.to_string(), result))
                    .collect()
            })
    }

    /// Reads multiple files serially without caching.
    ///
    /// This compatibility method no longer uses Rayon’s global pool. Create a
    /// [`ParallelReadSession`] and call [`Self::read_many_with_session`] to
    /// request bounded local parallelism.
    #[deprecated(
        since = "0.0.1",
        note = "this compatibility method is serial; use ParallelReadSession with read_many_with_session"
    )]
    pub fn read_many_parallel<'a>(
        &self,
        names: &'a [&'a str],
    ) -> Vec<(&'a str, Result<Vec<u8>, Error>)> {
        names
            .iter()
            .map(|name| (*name, self.inner.read(name)))
            .collect()
    }

    /// Reads multiple files serially while preserving individual errors.
    ///
    /// This compatibility method no longer uses Rayon’s global pool. Create a
    /// [`ParallelReadSession`] and call [`Self::read_many_with_session`] to
    /// request bounded local parallelism.
    #[deprecated(
        since = "0.0.1",
        note = "this compatibility method is serial; use ParallelReadSession with read_many_with_session"
    )]
    pub fn read_many_parallel_results<'a>(
        &self,
        names: &'a [&'a str],
    ) -> Vec<(&'a str, Result<Vec<u8>, Error>)> {
        names
            .iter()
            .map(|name| (*name, self.inner.read(name)))
            .collect()
    }

    /// Reads multiple files serially with caching.
    ///
    /// This compatibility method no longer uses Rayon’s global pool. Explicit
    /// session reads intentionally bypass the cache so cancellation cannot
    /// publish a partial cache population.
    #[deprecated(
        since = "0.0.1",
        note = "this compatibility method is serial; explicit session reads bypass the cache"
    )]
    pub fn read_many_parallel_cached<'a>(
        &self,
        names: &'a [&'a str],
    ) -> Vec<(&'a str, Result<Vec<u8>, Error>)> {
        names.iter().map(|name| (*name, self.read(name))).collect()
    }

    /// Reads all files serially, caching results.
    ///
    /// This compatibility method no longer uses Rayon’s global pool. Explicit
    /// session reads intentionally bypass the cache so cancellation cannot
    /// publish a partial cache population.
    #[deprecated(
        since = "0.0.1",
        note = "this compatibility method is serial; explicit session reads bypass the cache"
    )]
    pub fn read_all_parallel(&self) -> Vec<(String, Result<Vec<u8>, Error>)> {
        let names: Vec<&str> = self.inner.file_names().collect();
        names
            .into_iter()
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
    fn indexed_archive_reads_stored_and_deflated_members_by_stable_id() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("mimetype", b"application/test")
            .unwrap();
        writer
            .write_deflated("content.xml", b"<content>Hello</content>")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        assert!(local_member_has_data_descriptor(&bytes, b"content.xml"));

        let archive = indexed_archive(bytes);
        assert_eq!(archive.len(), 2);
        assert!(archive.contains("/mimetype"));
        assert!(archive.is_stored("mimetype").unwrap());
        assert!(!archive.is_stored("content.xml").unwrap());
        assert_eq!(
            archive.file_names().collect::<Vec<_>>(),
            ["mimetype", "content.xml"]
        );

        let id = archive
            .entry_id("content.xml")
            .expect("indexed content entry");
        let metadata = archive.metadata_for(id).unwrap();
        assert_eq!(metadata.uncompressed_size(), 24);
        assert_eq!(archive.read_entry(id).unwrap(), b"<content>Hello</content>");
        assert_eq!(archive.read("mimetype").unwrap(), b"application/test");
    }

    #[test]
    fn indexed_archive_defers_crc_validation_until_payload_read() {
        let mut bytes = bulk_fixture();
        corrupt_payload(&mut bytes, b"bad");
        let archive = indexed_archive(bytes);

        assert_eq!(archive.read("first").unwrap(), b"first");
        assert!(archive.read("bad").is_err());
    }

    #[test]
    fn indexed_archive_rejects_duplicate_normalized_names_and_limits() {
        let duplicate = fixture(&[
            FixtureEntry::stored(b"dir/../same.xml", b"one"),
            FixtureEntry::stored(b"same.xml", b"two"),
        ]);
        assert!(matches!(
            indexed_archive_result(duplicate, ArchiveLimits::UNBOUNDED),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { msg } if msg.contains("duplicate normalized"))
        ));

        let bytes = fixture(&[
            FixtureEntry::stored(b"first.xml", b"1234"),
            FixtureEntry::stored(b"second.xml", b"5678"),
        ]);
        let limits = ArchiveLimits {
            max_files: 1,
            ..ArchiveLimits::UNBOUNDED
        };
        assert_limit(
            indexed_archive_result(bytes, limits).unwrap_err(),
            LimitResource::FileCount,
            2,
            1,
        );
    }

    #[test]
    fn indexed_archive_applies_zip64_sizes_from_one_located_archive() {
        let zip64 = zip64_sizes(5, 0);
        let bytes = fixture(&[FixtureEntry {
            name: b"zip64.bin",
            extra: &zip64,
            comment: b"",
            compressed_size: u32::MAX,
            uncompressed_size: u32::MAX,
            data: b"",
        }]);
        let source_length = bytes.len() as u64;
        let mut buffer = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
        let located = ZipLocator::new()
            .locate_in_reader(std::io::Cursor::new(bytes), &mut buffer, source_length)
            .map_err(|(_reader, error)| error)
            .unwrap();

        let limits = ArchiveLimits {
            max_entry_size: 4,
            ..ArchiveLimits::UNBOUNDED
        };
        assert_limit(
            IndexedArchive::from_zip_archive_with_limits(located, limits).unwrap_err(),
            LimitResource::EntrySize,
            5,
            4,
        );
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
        let session = test_parallel_session(4);
        let never_cancel = || false;

        let requested = ["last", "bad", "missing", "first"];
        let results = reader
            .read_many_with_session(&session, &requested, &never_cancel)
            .unwrap();
        assert_eq!(results.len(), requested.len());
        assert_eq!(results[0].0, "last");
        assert_eq!(results[0].1.as_ref().unwrap(), b"last");
        assert_eq!(results[1].0, "bad");
        assert!(
            matches!(results[1].1, Err(ref error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. }))
        );
        assert_eq!(results[2].0, "missing");
        assert!(
            matches!(results[2].1, Err(ref error) if matches!(error.kind(), ErrorKind::FileNotFound(_)))
        );
        assert_eq!(results[3].0, "first");
        assert_eq!(results[3].1.as_ref().unwrap(), b"first");

        let all = reader
            .read_all_with_session(&session, &never_cancel)
            .unwrap();
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
    #[allow(deprecated)]
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

    #[test]
    fn local_sessions_with_one_two_and_four_workers_preserve_results_and_order() {
        let mut bytes = bulk_fixture();
        corrupt_payload(&mut bytes, b"bad");
        let reader = ArchiveReader::new(&bytes).unwrap();
        let requested = ["last", "bad", "missing", "first"];
        let never_cancel = || false;

        for workers in [1, 2, 4] {
            let session = test_parallel_session(workers);
            let results = reader
                .read_many_with_session(&session, &requested, &never_cancel)
                .unwrap();
            assert_eq!(
                results.iter().map(|(name, _)| *name).collect::<Vec<_>>(),
                requested
            );
            assert_eq!(results[0].1.as_ref().unwrap(), b"last");
            assert!(
                matches!(results[1].1, Err(ref error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. }))
            );
            assert!(
                matches!(results[2].1, Err(ref error) if matches!(error.kind(), ErrorKind::FileNotFound(_)))
            );
            assert_eq!(results[3].1.as_ref().unwrap(), b"first");
        }
    }

    #[test]
    fn parallel_session_limits_are_finite_and_validated() {
        assert!(matches!(
            ParallelReadLimits::new(
                std::num::NonZeroUsize::new(2).unwrap(),
                std::num::NonZeroUsize::new(1).unwrap(),
                std::num::NonZeroU64::new(1024).unwrap(),
                0,
            ),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidParallelReadLimits { .. })
        ));
        assert!(matches!(
            ParallelReadLimits::new(
                std::num::NonZeroUsize::new(1).unwrap(),
                std::num::NonZeroUsize::new(1).unwrap(),
                std::num::NonZeroU64::new(7).unwrap(),
                8,
            ),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidParallelReadLimits { .. })
        ));

        let limits = test_parallel_limits(4);
        assert_eq!(limits.workers().get(), 4);
        assert_eq!(limits.max_in_flight_tasks().get(), 8);
        assert_eq!(limits.max_in_flight_bytes().get(), 4096);
        assert_eq!(limits.min_parallel_bytes(), 0);
        assert_eq!(limits.affinity(), ParallelAffinity::Inherit);
    }

    #[test]
    fn pre_cancelled_session_reads_no_member_and_does_not_populate_lazy_cache() {
        let bytes = bulk_fixture();
        let reader = LazyArchiveReader::new(&bytes).unwrap();
        let session = test_parallel_session(2);
        let cancelled = || true;

        let error = reader
            .read_many_with_session(&session, &["missing"], &cancelled)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::Cancelled));
        assert_eq!(reader.cache_size(), 0);
    }

    #[test]
    fn cancellation_discards_batch_results_and_does_not_publish_lazy_cache_entries() {
        let bytes = bulk_fixture();
        let reader = LazyArchiveReader::new(&bytes).unwrap();
        let session = test_parallel_session(1);
        let cancellation = CancelAfter::new(4);

        let error = reader
            .read_many_with_session(&session, &["first"], &cancellation)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::Cancelled));
        assert_eq!(reader.cache_size(), 0);
    }

    #[test]
    fn local_session_uses_its_explicit_worker_count() {
        let session = test_parallel_session(4);
        assert_eq!(session.worker_count().get(), 4);
        assert_eq!(
            session
                .pool
                .as_ref()
                .map(rayon::ThreadPool::current_num_threads),
            Some(4)
        );
    }

    #[test]
    fn indexed_archive_can_use_an_explicit_local_session() {
        let archive = indexed_archive(bulk_fixture());
        let session = test_parallel_session(2);
        let never_cancel = || false;
        let results = archive
            .read_many_with_session(&session, &["last", "first"], &never_cancel)
            .unwrap();

        assert_eq!(results[0].0, "last");
        assert_eq!(results[0].1.as_ref().unwrap(), b"last");
        assert_eq!(results[1].0, "first");
        assert_eq!(results[1].1.as_ref().unwrap(), b"first");
    }

    fn indexed_archive(bytes: Vec<u8>) -> IndexedArchive<std::io::Cursor<Vec<u8>>> {
        indexed_archive_result(bytes, ArchiveLimits::default()).expect("valid indexed archive")
    }

    fn test_parallel_limits(workers: usize) -> ParallelReadLimits {
        ParallelReadLimits::new(
            std::num::NonZeroUsize::new(workers).expect("test worker count is nonzero"),
            std::num::NonZeroUsize::new(8).expect("test task count is nonzero"),
            std::num::NonZeroU64::new(4096).expect("test byte count is nonzero"),
            0,
        )
        .expect("test limits are valid")
    }

    fn test_parallel_session(workers: usize) -> ParallelReadSession {
        ParallelReadSession::new(test_parallel_limits(workers))
            .expect("local test worker pool is available")
    }

    struct CancelAfter {
        checks: std::sync::atomic::AtomicUsize,
        cancel_on_check: usize,
    }

    impl CancelAfter {
        fn new(cancel_on_check: usize) -> Self {
            Self {
                checks: std::sync::atomic::AtomicUsize::new(0),
                cancel_on_check,
            }
        }
    }

    impl CancellationProbe for CancelAfter {
        fn is_cancelled(&self) -> bool {
            self.checks
                .fetch_add(1, std::sync::atomic::Ordering::AcqRel)
                >= self.cancel_on_check
        }
    }

    fn indexed_archive_result(
        bytes: Vec<u8>,
        limits: ArchiveLimits,
    ) -> Result<IndexedArchive<std::io::Cursor<Vec<u8>>>, Error> {
        let length = bytes.len() as u64;
        IndexedArchive::from_reader_with_limits(std::io::Cursor::new(bytes), length, limits)
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

    fn local_member_has_data_descriptor(archive: &[u8], wanted_name: &[u8]) -> bool {
        const LOCAL_HEADER: [u8; 4] = 0x0403_4b50_u32.to_le_bytes();
        archive.windows(4).enumerate().any(|(offset, signature)| {
            if signature != LOCAL_HEADER || offset.saturating_add(30) > archive.len() {
                return false;
            }
            let name_len =
                u16::from_le_bytes([archive[offset + 26], archive[offset + 27]]) as usize;
            let name_end = offset.saturating_add(30).saturating_add(name_len);
            name_end <= archive.len()
                && &archive[offset + 30..name_end] == wanted_name
                && u16::from_le_bytes([archive[offset + 6], archive[offset + 7]]) & 0x08 != 0
        })
    }
}
