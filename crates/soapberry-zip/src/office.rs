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
    CompressionMethod, Error, ErrorKind, PreservationIndex, RECOMMENDED_BUFFER_SIZE, ReaderAt,
    ZipArchive, ZipArchiveWriter, ZipLocator, ZipSliceArchive, ZipVerification,
};
use flate2::Compression;
use flate2::read::DeflateDecoder;
use flate2::write::DeflateEncoder;
use rayon::prelude::*;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

pub use crate::LimitResource;

/// Validation policy for an indexed Office archive.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ArchiveValidationPolicy {
    /// Preserve the compatibility path normalization behavior.
    #[default]
    Normalized,
    /// Reject unsafe/raw path spellings and cross-check the offset-zero
    /// `mimetype` local header against its central-directory record.
    StrictPackage,
}

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

/// Zero-allocation iterator over an [`ArchiveReader`] file-name order.
pub struct ArchiveReaderNames<'a> {
    names: std::slice::Iter<'a, String>,
}

impl<'a> Iterator for ArchiveReaderNames<'a> {
    type Item = &'a str;

    fn next(&mut self) -> Option<Self::Item> {
        self.names.next().map(String::as_str)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.names.size_hint()
    }
}

impl ExactSizeIterator for ArchiveReaderNames<'_> {}

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
    has_encrypted_entries: bool,
}

/// Zero-allocation iterator over an [`IndexedArchive`] file-name order.
pub struct IndexedArchiveNames<'a, R> {
    archive: &'a IndexedArchive<R>,
    order: std::slice::Iter<'a, EntryId>,
}

impl<'a, R> Iterator for IndexedArchiveNames<'a, R> {
    type Item = &'a str;

    fn next(&mut self) -> Option<Self::Item> {
        self.order
            .next()
            .map(|id| self.archive.entries[id.0].name.as_str())
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.order.size_hint()
    }
}

impl<R> ExactSizeIterator for IndexedArchiveNames<'_, R> {}

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

            let (name, lossy_name) = normalized_member_name(path);
            let name = canonical_member_name(name);

            // Directories are never exposed or decompressed by this API. They
            // consume name and metadata budgets above, but not file or payload
            // budgets. Retaining their compact declarations makes structural
            // inspection possible without changing file lookup behavior.
            if directory {
                if directories.contains_key(&name) {
                    return Err(duplicate_member_error(
                        &name,
                        lossy_name,
                        "duplicate normalized directory names",
                    ));
                }
                if index.contains_key(&name) {
                    return Err(file_directory_collision_error(&name, lossy_name));
                }
                directories.insert(
                    name,
                    Metadata {
                        compressed_size,
                        uncompressed_size,
                        directory: true,
                    },
                );
                continue;
            }

            if directories.contains_key(&name) {
                return Err(file_directory_collision_error(&name, lossy_name));
            }

            let local_header_offset = entry.local_header_offset();
            if index.contains_key(&name) {
                return Err(duplicate_member_error(
                    &name,
                    lossy_name,
                    "duplicate normalized file names",
                ));
            }
            index.insert(
                name.clone(),
                EntryInfo {
                    wayfinder: entry.wayfinder(),
                    compression_method: entry.compression_method(),
                    uncompressed_size,
                },
            );
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
        let lookup = lookup_member_name(name);
        !lookup.explicit_directory && self.index.contains_key(&lookup.name)
    }

    /// Return declared metadata for a normalized member name.
    ///
    /// This performs only hash-map lookup over the central-directory index. It
    /// never reads, decompresses, verifies, or allocates member payload data.
    pub fn metadata(&self, name: &str) -> Result<Metadata, Error> {
        let lookup = lookup_member_name(name);
        if !lookup.explicit_directory {
            if let Some(info) = self.index.get(&lookup.name) {
                return Ok(Metadata {
                    compressed_size: info.wayfinder.compressed_size_hint(),
                    uncompressed_size: info.uncompressed_size,
                    directory: false,
                });
            }
        }
        self.directories
            .get(&lookup.name)
            .copied()
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))
    }

    /// Get an iterator over all file names in the archive.
    pub fn file_names(&self) -> ArchiveReaderNames<'_> {
        ArchiveReaderNames {
            names: self.order.iter(),
        }
    }

    /// Whether an archive entry uses the ZIP Store method.
    ///
    /// ODF encryption is applied to an already-deflated byte stream, so the
    /// enclosing ZIP entry must not perform another compression transform.
    pub fn is_stored(&self, name: &str) -> Result<bool, Error> {
        let lookup = lookup_member_name(name);
        let info = self
            .index
            .get(&lookup.name)
            .filter(|_| !lookup.explicit_directory)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))?;
        Ok(info.compression_method == CompressionMethod::Store)
    }

    /// Read and decompress a file from the archive.
    ///
    /// Returns the decompressed contents of the file. Supports both stored
    /// (uncompressed) and deflated entries.
    pub fn read(&self, name: &str) -> Result<Vec<u8>, Error> {
        let lookup = lookup_member_name(name);

        let info = self
            .index
            .get(&lookup.name)
            .filter(|_| !lookup.explicit_directory)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))?;

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
    /// Return the exclusive byte offset of the located ZIP archive.
    ///
    /// This is retained from the initial EOCD location and therefore requires
    /// no source read or central-directory rescan. Callers can compare it with
    /// the positional source length before promising raw-member preservation.
    #[must_use]
    pub fn archive_end_offset(&self) -> u64 {
        self.archive.end_offset()
    }

    /// Build a raw-member preservation index from this already located ZIP.
    ///
    /// This borrows the positional archive held by this index and does not run
    /// another EOCD search. `scratch` is used to scan the existing central
    /// directory and must be large enough for its records.
    pub fn preservation_index<'archive>(
        &'archive self,
        scratch: &mut [u8],
    ) -> Result<PreservationIndex<'archive, R>, Error> {
        PreservationIndex::new(&self.archive, scratch)
    }

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
        Self::from_reader_with_limits_and_policy(
            reader,
            end_offset,
            limits,
            ArchiveValidationPolicy::Normalized,
        )
    }

    /// Locate and index a positional ZIP source with explicit limits and
    /// validation policy.
    pub fn from_reader_with_limits_and_policy(
        reader: R,
        end_offset: u64,
        limits: ArchiveLimits,
        policy: ArchiveValidationPolicy,
    ) -> Result<Self, Error> {
        let mut buffer = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipLocator::new()
            .locate_in_reader(reader, &mut buffer, end_offset)
            .map_err(|(_reader, error)| error)?;
        Self::from_zip_archive_with_limits_and_policy(archive, limits, policy)
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
        Self::from_zip_archive_with_limits_and_policy(
            archive,
            limits,
            ArchiveValidationPolicy::Normalized,
        )
    }

    /// Build an index from an already located ZIP archive with explicit limits
    /// and validation policy.
    pub fn from_zip_archive_with_limits_and_policy(
        archive: ZipArchive<R>,
        limits: ArchiveLimits,
        policy: ArchiveValidationPolicy,
    ) -> Result<Self, Error> {
        let mut index = HashMap::new();
        let mut entries = Vec::new();
        let mut directories = HashMap::new();
        let mut ordered_entries = Vec::new();
        let mut total_metadata_bytes = 0_u64;
        let mut total_uncompressed_size = 0_u64;
        let mut strict_mimetype = None;
        let mut has_encrypted_entries = false;
        let mut buffer = vec![0_u8; RECOMMENDED_BUFFER_SIZE];

        {
            let mut central_entries = archive.entries(&mut buffer);
            while let Some(entry) = central_entries.next_entry()? {
                has_encrypted_entries |= entry.flags() & 1 != 0;
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
                let (name, lossy_name) = match policy {
                    ArchiveValidationPolicy::Normalized => normalized_member_name(path),
                    ArchiveValidationPolicy::StrictPackage => (strict_member_name(path)?, false),
                };
                let name = canonical_member_name(name);

                if entry.is_dir() {
                    if directories.contains_key(&name) {
                        return Err(duplicate_member_error(
                            &name,
                            lossy_name,
                            "duplicate normalized directory names",
                        ));
                    }
                    if index.contains_key(&name) {
                        return Err(file_directory_collision_error(&name, lossy_name));
                    }
                    directories.insert(
                        name,
                        Metadata {
                            compressed_size,
                            uncompressed_size,
                            directory: true,
                        },
                    );
                    continue;
                }

                if directories.contains_key(&name) {
                    return Err(file_directory_collision_error(&name, lossy_name));
                }

                if matches!(policy, ArchiveValidationPolicy::StrictPackage)
                    && path.as_ref() == b"mimetype"
                {
                    strict_mimetype = Some((
                        entry.wayfinder(),
                        entry.flags(),
                        entry.compression_method().as_id().as_u16(),
                        entry.crc32(),
                        compressed_size,
                        uncompressed_size,
                    ));
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
                if index.contains_key(&name) {
                    return Err(duplicate_member_error(
                        &name,
                        lossy_name,
                        "duplicate normalized file names",
                    ));
                }
                index.insert(name.clone(), entry_id);
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

        if matches!(policy, ArchiveValidationPolicy::StrictPackage) {
            validate_strict_mimetype(&archive, strict_mimetype)?;
        }

        ordered_entries.sort_unstable_by_key(|(offset, _)| *offset);
        let order = ordered_entries.into_iter().map(|(_, id)| id).collect();

        Ok(Self {
            archive,
            index,
            entries,
            directories,
            order,
            has_encrypted_entries,
        })
    }

    /// Number of indexed non-directory members.
    #[inline]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Number of central-directory entries, including directory records.
    #[inline]
    pub fn preservation_entry_count(&self) -> usize {
        usize::try_from(self.archive.entries_hint()).unwrap_or(usize::MAX)
    }

    /// Exact source bytes occupied by the central directory, EOCD, and
    /// archive comment retained by raw-preservation planning.
    #[inline]
    pub fn preservation_metadata_bytes(&self) -> u64 {
        self.archive
            .end_offset()
            .saturating_sub(self.archive.directory_offset())
    }

    /// Whether any central-directory entry declares traditional ZIP
    /// encryption through general-purpose bit zero.
    #[inline]
    pub fn has_encrypted_entries(&self) -> bool {
        self.has_encrypted_entries
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
        let lookup = lookup_member_name(name);
        if lookup.explicit_directory {
            return None;
        }
        self.index.get(&lookup.name).copied()
    }

    /// Return declared metadata for a member without payload access.
    pub fn metadata(&self, name: &str) -> Result<Metadata, Error> {
        let lookup = lookup_member_name(name);
        match self.entry_id(name) {
            Some(id) => self.metadata_for(id),
            None => self
                .directories
                .get(&lookup.name)
                .copied()
                .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name))),
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
    pub fn file_names(&self) -> IndexedArchiveNames<'_, R> {
        IndexedArchiveNames {
            archive: self,
            order: self.order.iter(),
        }
    }

    /// Whether an indexed file uses ZIP Store compression.
    pub fn is_stored(&self, name: &str) -> Result<bool, Error> {
        let entry_id = self
            .entry_id(name)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup_member_name(name).name)))?;
        Ok(self.indexed_entry(entry_id)?.info.compression_method == CompressionMethod::Store)
    }

    /// Read and verify one member by name.
    pub fn read(&self, name: &str) -> Result<Vec<u8>, Error> {
        let entry_id = self
            .entry_id(name)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup_member_name(name).name)))?;
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
            Error::from(ErrorKind::Allocation {
                resource: "indexed archive entry output",
                source: error,
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

#[derive(Debug)]
struct LookupMemberName {
    name: String,
    explicit_directory: bool,
}

fn normalized_member_name(path: ZipFilePath<RawPath<'_>>) -> (String, bool) {
    match path.try_normalize() {
        Ok(normalized) => (normalized.as_ref().to_string(), false),
        Err(_) => {
            // Keep the existing lossy-UTF-8 compatibility behavior, but apply
            // the same path normalization as valid UTF-8 names so a name
            // returned by `file_names()` is always a usable lookup key.
            let lossy = String::from_utf8_lossy(path.as_ref());
            let normalized = ZipFilePath::from_str(&lossy);
            (normalized.as_str().to_string(), true)
        },
    }
}

fn canonical_member_name(name: String) -> String {
    name.trim_end_matches('/').to_string()
}

fn lookup_member_name(name: &str) -> LookupMemberName {
    let explicit_directory = name
        .as_bytes()
        .last()
        .is_some_and(|byte| matches!(byte, b'/' | b'\\'));
    let normalized = ZipFilePath::from_str(name);
    LookupMemberName {
        name: canonical_member_name(normalized.as_str().to_string()),
        explicit_directory,
    }
}

fn duplicate_member_error(name: &str, lossy_name: bool, kind: &str) -> Error {
    let suffix = if lossy_name {
        " (lossy UTF-8 name collision)"
    } else {
        ""
    };
    ErrorKind::InvalidInput {
        msg: format!("archive contains {kind}: {name}{suffix}"),
    }
    .into()
}

fn file_directory_collision_error(name: &str, lossy_name: bool) -> Error {
    let suffix = if lossy_name {
        " (lossy UTF-8 name collision)"
    } else {
        ""
    };
    ErrorKind::InvalidInput {
        msg: format!(
            "archive contains file/directory name collision after normalization: {name}{suffix}"
        ),
    }
    .into()
}

fn strict_member_name(path: ZipFilePath<RawPath<'_>>) -> Result<String, Error> {
    let raw = path.as_ref();
    if raw.is_empty() || raw.iter().any(|byte| *byte < 0x20) {
        return Err(ErrorKind::InvalidInput {
            msg: "strict package archive contains an empty or control-character member name"
                .to_string(),
        }
        .into());
    }
    let normalized = path.try_normalize().map_err(|_| ErrorKind::InvalidInput {
        msg: "strict package archive contains a non-UTF-8 member name".to_string(),
    })?;
    let normalized_name = normalized.as_str();
    if normalized_name.is_empty() {
        return Err(ErrorKind::InvalidInput {
            msg: "strict package archive contains a member that normalizes to an empty name"
                .to_string(),
        }
        .into());
    }
    let canonical = raw == normalized_name.as_bytes()
        || (path.is_dir()
            && raw.len() == normalized_name.len() + 1
            && &raw[..normalized_name.len()] == normalized_name.as_bytes()
            && raw[normalized_name.len()] == b'/');
    if !canonical {
        return Err(ErrorKind::InvalidInput {
            msg: "strict package archive contains an unsafe or non-canonical member name"
                .to_string(),
        }
        .into());
    }
    Ok(normalized_name.to_string())
}

fn validate_strict_mimetype<R: ReaderAt>(
    archive: &ZipArchive<R>,
    central: Option<(crate::ZipArchiveEntryWayfinder, u16, u16, u32, u64, u64)>,
) -> Result<(), Error> {
    let Some((wayfinder, flags, method, crc, compressed_size, uncompressed_size)) = central else {
        return Err(ErrorKind::InvalidInput {
            msg: "strict package archive has no central mimetype member".to_string(),
        }
        .into());
    };
    if wayfinder.local_header_offset() != 0 {
        return Err(ErrorKind::InvalidInput {
            msg: "central mimetype member does not point to offset-zero local header".to_string(),
        }
        .into());
    }

    let entry = archive.get_entry(wayfinder)?;
    let local = entry.local_header_fixed()?;
    let mut variable =
        vec![0_u8; usize::from(local.file_name_len) + usize::from(local.extra_field_len)];
    let local_header = entry.local_header(&mut variable)?;
    if local_header.file_path().as_bytes() != b"mimetype"
        || local.flags != flags
        || local.compression_method.as_u16() != method
        || local.crc32 != crc
        || u64::from(local.compressed_size) != compressed_size
        || u64::from(local.uncompressed_size) != uncompressed_size
    {
        return Err(ErrorKind::InvalidInput {
            msg: "central mimetype metadata does not match offset-zero local header".to_string(),
        }
        .into());
    }
    Ok(())
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

/// Bounded limits for the ZIP transport streaming writer.
///
/// The writer deliberately supports only ZIP32 output in this first transport
/// substrate. The output ceiling, per-entry ceiling, aggregate uncompressed
/// ceiling, member count, and metadata ceiling are finite by default and are
/// checked before an entry starts or before an input chunk is accepted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingArchiveLimits {
    /// Maximum number of file members accepted by a writer.
    pub max_entries: usize,
    /// Maximum UTF-8 member-name bytes for one file member.
    pub max_member_name_bytes: u64,
    /// Maximum aggregate variable central-directory metadata bytes.
    ///
    /// The streaming convenience methods currently add no extra fields, so
    /// this is the aggregate member-name budget.  ZIP fixed-size headers are
    /// not included, matching [`ArchiveLimits::max_metadata_bytes`].
    pub max_metadata_bytes: u64,
    /// Maximum compressed bytes accepted for one streamed member.
    ///
    /// This is the compressed payload ceiling, excluding the local header,
    /// data descriptor, and central-directory record. It is checked before a
    /// compressed write can exceed the limit and again before the entry is
    /// finalized.
    pub max_compressed_size: u64,
    /// Maximum uncompressed bytes accepted for one streamed member.
    pub max_entry_size: u64,
    /// Maximum aggregate uncompressed bytes accepted across members.
    pub max_total_size: u64,
    /// Maximum complete ZIP bytes accepted by the output sink.
    pub max_output_bytes: u64,
}

impl StreamingArchiveLimits {
    /// Creates explicit streaming metadata limits while retaining the default
    /// finite payload and output ceilings.
    #[must_use]
    pub const fn new(
        max_entries: usize,
        max_member_name_bytes: u64,
        max_metadata_bytes: u64,
    ) -> Self {
        Self {
            max_entries,
            max_member_name_bytes,
            max_metadata_bytes,
            max_compressed_size: DEFAULT_STREAM_MAX_COMPRESSED_SIZE,
            max_entry_size: DEFAULT_STREAM_MAX_ENTRY_SIZE,
            max_total_size: DEFAULT_STREAM_MAX_TOTAL_SIZE,
            max_output_bytes: DEFAULT_STREAM_MAX_OUTPUT_BYTES,
        }
    }

    /// Replaces the finite payload and output ceilings.
    #[must_use]
    pub const fn with_byte_limits(
        mut self,
        max_entry_size: u64,
        max_total_size: u64,
        max_output_bytes: u64,
    ) -> Self {
        self.max_entry_size = max_entry_size;
        self.max_total_size = max_total_size;
        self.max_output_bytes = max_output_bytes;
        if self.max_compressed_size > max_output_bytes {
            self.max_compressed_size = max_output_bytes;
        }
        self
    }

    /// Replaces the finite compressed-payload ceiling for one member.
    #[must_use]
    pub const fn with_compressed_size_limit(mut self, max_compressed_size: u64) -> Self {
        self.max_compressed_size = max_compressed_size;
        self
    }
}

impl Default for StreamingArchiveLimits {
    fn default() -> Self {
        let limits = ArchiveLimits::default();
        Self {
            max_entries: DEFAULT_STREAM_MAX_ENTRIES,
            max_member_name_bytes: limits.max_member_name_bytes,
            max_metadata_bytes: limits.max_metadata_bytes,
            max_compressed_size: DEFAULT_STREAM_MAX_COMPRESSED_SIZE,
            max_entry_size: DEFAULT_STREAM_MAX_ENTRY_SIZE,
            max_total_size: DEFAULT_STREAM_MAX_TOTAL_SIZE,
            max_output_bytes: DEFAULT_STREAM_MAX_OUTPUT_BYTES,
        }
    }
}

const ZIP32_MAX_VALUE: u64 = u32::MAX as u64;
const ZIP32_MAX_ENTRIES: usize = u16::MAX as usize - 1;
const ZIP32_MAX_MEMBER_NAME_BYTES: u64 = u16::MAX as u64;
const MIN_STREAM_OUTPUT_BYTES: u64 = 22;
const DEFAULT_STREAM_MAX_ENTRIES: usize = 65_534;
const DEFAULT_STREAM_MAX_COMPRESSED_SIZE: u64 = 512 * 1024 * 1024;
const DEFAULT_STREAM_MAX_ENTRY_SIZE: u64 = 512 * 1024 * 1024;
const DEFAULT_STREAM_MAX_TOTAL_SIZE: u64 = 2 * 1024 * 1024 * 1024;
const DEFAULT_STREAM_MAX_OUTPUT_BYTES: u64 = 512 * 1024 * 1024;
const STREAM_COPY_BUFFER_SIZE: usize = 16 * 1024;

/// Content-free progress exposed after a non-atomic streaming failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingArchiveProgress {
    /// Bytes reported as accepted by the output sink.
    output_bytes: u64,
    /// Whether a failed entry permanently poisoned the writer.
    poisoned: bool,
}

impl StreamingArchiveProgress {
    /// Bytes reported as accepted by the output sink.
    #[must_use]
    pub const fn output_bytes(self) -> u64 {
        self.output_bytes
    }

    /// Whether the writer rejects all subsequent entry and finish operations.
    #[must_use]
    pub const fn is_poisoned(self) -> bool {
        self.poisoned
    }
}

/// A byte resource bounded by the sequential streaming transport.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StreamingLimitResource {
    /// Compressed payload bytes for one member.
    CompressedBytes,
    /// Bytes accepted by the output sink for the complete ZIP stream.
    OutputBytes,
}

impl std::fmt::Display for StreamingLimitResource {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::CompressedBytes => "compressed member bytes",
            Self::OutputBytes => "output bytes",
        })
    }
}

/// Typed attribution for a streaming byte ceiling reached after publication
/// has started.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct StreamingLimitExceeded {
    resource: StreamingLimitResource,
    actual: u64,
    maximum: u64,
}

impl StreamingLimitExceeded {
    /// The bounded resource that exceeded its ceiling.
    #[must_use]
    pub const fn resource(self) -> StreamingLimitResource {
        self.resource
    }

    /// The attempted or observed byte count.
    #[must_use]
    pub const fn actual(self) -> u64 {
        self.actual
    }

    /// The configured byte ceiling.
    #[must_use]
    pub const fn maximum(self) -> u64 {
        self.maximum
    }
}

impl std::fmt::Display for StreamingLimitExceeded {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "streaming ZIP {} limit exceeded: attempted {}, maximum {}",
            self.resource, self.actual, self.maximum
        )
    }
}

impl std::error::Error for StreamingLimitExceeded {}

#[derive(Debug)]
struct StreamingLimitMarker {
    limit: StreamingLimitExceeded,
}

impl std::fmt::Display for StreamingLimitMarker {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.limit.fmt(formatter)
    }
}

impl std::error::Error for StreamingLimitMarker {}

#[derive(Debug)]
struct StreamingPayloadLimitMarker {
    resource: LimitResource,
    actual: u64,
    maximum: u64,
}

impl std::fmt::Display for StreamingPayloadLimitMarker {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "streaming ZIP {} limit exceeded: attempted {}, maximum {}",
            self.resource, self.actual, self.maximum
        )
    }
}

impl std::error::Error for StreamingPayloadLimitMarker {}

fn streaming_payload_limit_io_error(
    resource: LimitResource,
    actual: u64,
    maximum: u64,
) -> std::io::Error {
    std::io::Error::new(
        std::io::ErrorKind::Other,
        StreamingPayloadLimitMarker {
            resource,
            actual,
            maximum,
        },
    )
}

fn streaming_payload_limit_from_io_error(
    io_error: &std::io::Error,
) -> Option<(LimitResource, u64, u64)> {
    io_error
        .get_ref()
        .and_then(|source| source.downcast_ref::<StreamingPayloadLimitMarker>())
        .map(|marker| (marker.resource, marker.actual, marker.maximum))
}

fn streaming_limit_io_error(
    resource: StreamingLimitResource,
    actual: u64,
    maximum: u64,
) -> std::io::Error {
    std::io::Error::new(
        std::io::ErrorKind::Other,
        StreamingLimitMarker {
            limit: StreamingLimitExceeded {
                resource,
                actual,
                maximum,
            },
        },
    )
}

fn streaming_limit_from_error(error: &Error) -> Option<StreamingLimitExceeded> {
    let io_error = match error.kind() {
        ErrorKind::IO(error) | ErrorKind::Io(error) => error,
        _ => return None,
    };
    streaming_limit_from_io_error(io_error)
}

fn streaming_limit_from_io_error(io_error: &std::io::Error) -> Option<StreamingLimitExceeded> {
    if let Some((actual, maximum)) = crate::writer::owned_entry_limit_from_io_error(io_error) {
        return Some(StreamingLimitExceeded {
            resource: StreamingLimitResource::CompressedBytes,
            actual,
            maximum,
        });
    }
    io_error
        .get_ref()
        .and_then(|source| source.downcast_ref::<StreamingLimitMarker>())
        .map(|marker| marker.limit)
}

/// A streaming publication failure with content-free output progress.
#[derive(Debug)]
pub struct StreamingArchiveFailure {
    error: Error,
    progress: StreamingArchiveProgress,
    limit: Option<StreamingLimitExceeded>,
}

impl StreamingArchiveFailure {
    /// The underlying ZIP or sink error.
    #[must_use]
    pub fn error(&self) -> &Error {
        &self.error
    }

    /// Bytes accepted before the failure and poison state.
    #[must_use]
    pub const fn progress(&self) -> StreamingArchiveProgress {
        self.progress
    }

    /// Returns typed attribution when a streaming byte ceiling caused this
    /// incomplete publication.
    #[must_use]
    pub const fn limit(&self) -> Option<StreamingLimitExceeded> {
        self.limit
    }

    /// Consume the typed failure and return its underlying ZIP error.
    #[must_use]
    pub fn into_error(self) -> Error {
        self.error
    }
}

impl std::fmt::Display for StreamingArchiveFailure {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.error.fmt(formatter)
    }
}

impl std::error::Error for StreamingArchiveFailure {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self.error.kind() {
            ErrorKind::IO(error) | ErrorKind::Io(error) => Some(error),
            _ => Some(&self.error),
        }
    }
}

#[derive(Debug)]
struct BoundedOutput<W> {
    writer: W,
    accepted: u64,
    maximum: u64,
    counter: Arc<AtomicU64>,
}

impl<W> BoundedOutput<W> {
    fn new(writer: W, maximum: u64, counter: Arc<AtomicU64>) -> Self {
        Self {
            writer,
            accepted: 0,
            maximum,
            counter,
        }
    }

    fn into_inner(self) -> W {
        self.writer
    }
}

impl<W: Write> Write for BoundedOutput<W> {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        if buffer.is_empty() {
            return Ok(0);
        }
        let requested = u64::try_from(buffer.len()).unwrap_or(u64::MAX);
        let remaining = self.maximum.saturating_sub(self.accepted);
        if requested > remaining {
            return Err(streaming_limit_io_error(
                StreamingLimitResource::OutputBytes,
                self.accepted.saturating_add(requested),
                self.maximum,
            ));
        }
        let written = self.writer.write(buffer)?;
        self.accepted = self.accepted.saturating_add(written as u64);
        self.counter.store(self.accepted, Ordering::Release);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.writer.flush()
    }
}

/// A ZIP entry sink that bounds compressed payload bytes before forwarding
/// them to the archive writer.
struct LimitedEntryWriter<'entry, 'archive, W> {
    inner: &'entry mut crate::ZipEntryWriter<'archive, BoundedOutput<W>>,
    maximum: u64,
}

impl<'entry, 'archive, W> LimitedEntryWriter<'entry, 'archive, W> {
    fn new(
        inner: &'entry mut crate::ZipEntryWriter<'archive, BoundedOutput<W>>,
        maximum: u64,
    ) -> Self {
        Self { inner, maximum }
    }

    fn compressed_bytes(&self) -> u64 {
        self.inner.compressed_bytes()
    }

    fn ensure_within_limit(&self) -> Result<(), Error> {
        let compressed = self.compressed_bytes();
        if compressed > self.maximum {
            Err(limit_error(
                LimitResource::CompressedSize,
                compressed,
                self.maximum,
            ))
        } else {
            Ok(())
        }
    }
}

impl<W: Write> Write for LimitedEntryWriter<'_, '_, W> {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let requested = u64::try_from(buffer.len()).unwrap_or(u64::MAX);
        let compressed = self.compressed_bytes();
        let remaining = self.maximum.saturating_sub(compressed);
        if requested > remaining {
            return Err(streaming_limit_io_error(
                StreamingLimitResource::CompressedBytes,
                compressed.saturating_add(requested),
                self.maximum,
            ));
        }
        self.inner.write(buffer)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

/// Bounded ZIP transport writer for sequential Office package members.
///
/// This substrate does not construct semantic XLSX, DOCX, PPTX, or ODF
/// models. Format crates remain responsible for validating and serializing
/// those models before handing a member reader to this writer.
pub struct StreamingArchiveWriter<W: Write> {
    archive: ZipArchiveWriter<BoundedOutput<W>>,
    limits: StreamingArchiveLimits,
    entries: usize,
    metadata_bytes: u64,
    total_uncompressed_bytes: u64,
    output_bytes: u64,
    poisoned: bool,
    names: HashSet<String>,
    last_limit: Option<StreamingLimitExceeded>,
    output_counter: Arc<AtomicU64>,
}

/// A consuming, bounded ZIP entry writer.
///
/// The entry owns the archive writer while it is active, so callers can pass
/// this value through a streaming pipeline without holding a mutable borrow
/// into a parent archive. [`Write`] accepts uncompressed bytes for either
/// Store or Deflate output. [`Self::finish`] consumes the entry and recovers a
/// [`StreamingArchiveWriter`] for the next member.
pub struct StreamingArchiveEntry<W: Write> {
    entry: Option<crate::ZipOwnedEntryWriter<BoundedOutput<W>>>,
    limits: StreamingArchiveLimits,
    entries: usize,
    metadata_bytes: u64,
    total_uncompressed_bytes: u64,
    names: HashSet<String>,
    normalized_name: String,
    name_bytes: u64,
    uncompressed_bytes: u64,
    output_bytes: u64,
    poisoned: bool,
    last_limit: Option<StreamingLimitExceeded>,
    failure: Option<Error>,
    output_counter: Arc<AtomicU64>,
}

impl<W: Write> StreamingArchiveEntry<W> {
    /// Number of uncompressed payload bytes accepted by this entry.
    #[must_use]
    pub const fn uncompressed_bytes(&self) -> u64 {
        self.uncompressed_bytes
    }

    /// Number of compressed payload bytes accepted by this entry.
    #[must_use]
    pub fn compressed_bytes(&self) -> u64 {
        self.entry
            .as_ref()
            .map(crate::ZipOwnedEntryWriter::compressed_bytes)
            .unwrap_or(0)
    }

    /// Content-free progress for the active entry.
    #[must_use]
    pub fn progress(&self) -> StreamingArchiveProgress {
        StreamingArchiveProgress {
            output_bytes: self.output_counter.load(Ordering::Acquire),
            poisoned: self.poisoned,
        }
    }

    /// Whether a payload or sink failure permanently invalidated this entry.
    #[must_use]
    pub const fn is_poisoned(&self) -> bool {
        self.poisoned
    }

    fn refresh_output_bytes(&mut self) {
        self.output_bytes = self.output_counter.load(Ordering::Acquire);
    }

    fn capture_io_failure(&mut self, error: std::io::Error) -> std::io::Error {
        let limit = streaming_limit_from_io_error(&error);
        self.last_limit = limit;
        self.refresh_output_bytes();

        let returned = if let Some((resource, actual, maximum)) =
            streaming_payload_limit_from_io_error(&error)
        {
            self.failure = Some(limit_error(resource, actual, maximum));
            error
        } else if let Some(StreamingLimitExceeded {
            resource: StreamingLimitResource::CompressedBytes,
            actual,
            maximum,
        }) = limit
        {
            self.failure = Some(limit_error(LimitResource::CompressedSize, actual, maximum));
            error
        } else {
            let kind = error.kind();
            let message = error.to_string();
            // `std::io::Error` is not cloneable.  Keep the original object in
            // the retained typed failure and return a lightweight immediate
            // notification to the `Write` caller; `finish_with_progress` is
            // the publication result that carries the complete source chain.
            let returned = std::io::Error::new(kind, message);
            // Move the original error into the typed failure.  Rebuilding the
            // error from its display text here would erase nested/custom
            // sources before `finish_with_progress` can report them.
            self.failure = Some(Error::from(error));
            returned
        };
        self.poisoned = true;
        returned
    }

    fn failure(self) -> StreamingArchiveFailure {
        StreamingArchiveFailure {
            error: self.failure.unwrap_or_else(|| {
                ErrorKind::InvalidInput {
                    msg: "streaming ZIP entry was poisoned".to_string(),
                }
                .into()
            }),
            progress: StreamingArchiveProgress {
                output_bytes: self.output_counter.load(Ordering::Acquire),
                poisoned: self.poisoned,
            },
            limit: self.last_limit,
        }
    }

    /// Finishes the entry and recovers the bounded archive writer.
    pub fn finish(self) -> Result<StreamingArchiveWriter<W>, StreamingArchiveFailure> {
        self.finish_with_progress()
            .map(|(writer, _progress)| writer)
    }

    /// Finishes the entry while preserving output progress on failure.
    pub fn finish_with_progress(
        mut self,
    ) -> Result<(StreamingArchiveWriter<W>, StreamingArchiveProgress), StreamingArchiveFailure>
    {
        if self.poisoned {
            return Err(self.failure());
        }

        let entry = match self.entry.take() {
            Some(entry) => entry,
            None => {
                self.poisoned = true;
                self.failure = Some(
                    ErrorKind::InvalidInput {
                        msg: "streaming ZIP entry writer was already finished".to_string(),
                    }
                    .into(),
                );
                return Err(self.failure());
            },
        };
        let archive = match entry.finish() {
            Ok(archive) => archive,
            Err(error) => {
                self.poisoned = true;
                let limit = streaming_limit_from_error(&error).or(self.last_limit);
                let error = match limit {
                    Some(StreamingLimitExceeded {
                        resource: StreamingLimitResource::CompressedBytes,
                        actual,
                        maximum,
                    }) => limit_error(LimitResource::CompressedSize, actual, maximum),
                    _ => error,
                };
                self.failure = Some(error);
                self.last_limit = limit;
                return Err(self.failure());
            },
        };

        let mut writer = StreamingArchiveWriter {
            archive,
            limits: self.limits,
            entries: self.entries.saturating_add(1),
            metadata_bytes: self.metadata_bytes.saturating_add(self.name_bytes),
            total_uncompressed_bytes: self
                .total_uncompressed_bytes
                .saturating_add(self.uncompressed_bytes),
            output_bytes: self.output_bytes,
            poisoned: false,
            names: self.names,
            last_limit: None,
            output_counter: self.output_counter,
        };
        let inserted = writer.names.insert(self.normalized_name);
        debug_assert!(inserted);
        writer.refresh_output_bytes();
        let progress = writer.progress();
        Ok((writer, progress))
    }
}

impl<W: Write> Write for StreamingArchiveEntry<W> {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        if self.poisoned {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "streaming ZIP entry writer is poisoned",
            ));
        }
        if buffer.is_empty() {
            return Ok(0);
        }

        let requested = u64::try_from(buffer.len()).unwrap_or(u64::MAX);
        let next_entry = self.uncompressed_bytes.saturating_add(requested);
        if next_entry > self.limits.max_entry_size {
            let error = streaming_payload_limit_io_error(
                LimitResource::EntrySize,
                next_entry,
                self.limits.max_entry_size,
            );
            return Err(self.capture_io_failure(error));
        }
        let next_total = self.total_uncompressed_bytes.saturating_add(next_entry);
        if next_total > self.limits.max_total_size {
            let error = streaming_payload_limit_io_error(
                LimitResource::TotalSize,
                next_total,
                self.limits.max_total_size,
            );
            return Err(self.capture_io_failure(error));
        }

        let result = self
            .entry
            .as_mut()
            .ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::Other, "streaming ZIP entry finished")
            })
            .and_then(|entry| entry.write(buffer));
        match result {
            Ok(0) => {
                let error = std::io::Error::new(
                    std::io::ErrorKind::WriteZero,
                    "streaming ZIP entry sink accepted no bytes",
                );
                Err(self.capture_io_failure(error))
            },
            Ok(written) => {
                self.uncompressed_bytes = self.uncompressed_bytes.saturating_add(written as u64);
                self.refresh_output_bytes();
                Ok(written)
            },
            Err(error) => Err(self.capture_io_failure(error)),
        }
    }

    fn flush(&mut self) -> std::io::Result<()> {
        if self.poisoned {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "streaming ZIP entry writer is poisoned",
            ));
        }
        let result = self
            .entry
            .as_mut()
            .ok_or_else(|| {
                std::io::Error::new(std::io::ErrorKind::Other, "streaming ZIP entry finished")
            })
            .and_then(Write::flush);
        match result {
            Ok(()) => Ok(()),
            Err(error) => Err(self.capture_io_failure(error)),
        }
    }
}

impl StreamingArchiveWriter<std::io::Cursor<Vec<u8>>> {
    /// Create a new streaming archive writer that writes to memory.
    pub fn new() -> Self {
        Self::with_limits(StreamingArchiveLimits::default())
    }

    /// Create a new in-memory writer with explicit finite metadata limits.
    pub fn with_limits(limits: StreamingArchiveLimits) -> Self {
        let output_counter = Arc::new(AtomicU64::new(0));
        Self {
            archive: ZipArchiveWriter::new(BoundedOutput::new(
                std::io::Cursor::new(Vec::new()),
                limits.max_output_bytes,
                Arc::clone(&output_counter),
            )),
            limits,
            entries: 0,
            metadata_bytes: 0,
            total_uncompressed_bytes: 0,
            output_bytes: 0,
            poisoned: false,
            names: HashSet::new(),
            last_limit: None,
            output_counter,
        }
    }

    /// Finish writing and return the ZIP archive bytes.
    pub fn finish_to_bytes(self) -> Result<Vec<u8>, Error> {
        Ok(self.finish()?.into_inner())
    }
}

impl<W: Write> StreamingArchiveWriter<W> {
    /// Create a new streaming archive writer with a custom writer.
    pub fn with_writer(writer: W) -> Self {
        Self::with_writer_and_limits(writer, StreamingArchiveLimits::default())
    }

    /// Create a new streaming archive writer with a custom writer and
    /// explicit metadata limits.
    pub fn with_writer_and_limits(writer: W, limits: StreamingArchiveLimits) -> Self {
        let output_counter = Arc::new(AtomicU64::new(0));
        Self {
            archive: ZipArchiveWriter::new(BoundedOutput::new(
                writer,
                limits.max_output_bytes,
                Arc::clone(&output_counter),
            )),
            limits,
            entries: 0,
            metadata_bytes: 0,
            total_uncompressed_bytes: 0,
            output_bytes: 0,
            poisoned: false,
            names: HashSet::new(),
            last_limit: None,
            output_counter,
        }
    }

    /// Return the metadata policy used by this writer.
    #[must_use]
    pub const fn limits(&self) -> StreamingArchiveLimits {
        self.limits
    }

    /// Return the number of successfully finalized file members.
    #[must_use]
    pub const fn entry_count(&self) -> usize {
        self.entries
    }

    /// Return the aggregate variable central-directory metadata bytes retained
    /// for finalized members.
    #[must_use]
    pub const fn metadata_bytes(&self) -> u64 {
        self.metadata_bytes
    }

    /// Return aggregate uncompressed bytes accepted for finalized members.
    #[must_use]
    pub const fn total_uncompressed_bytes(&self) -> u64 {
        self.total_uncompressed_bytes
    }

    /// Return content-free progress for this non-atomic writer.
    #[must_use]
    pub const fn progress(&self) -> StreamingArchiveProgress {
        StreamingArchiveProgress {
            output_bytes: self.output_bytes,
            poisoned: self.poisoned,
        }
    }

    /// Return bytes reported as accepted by the output sink.
    #[must_use]
    pub const fn output_bytes(&self) -> u64 {
        self.output_bytes
    }

    /// Whether a failed entry permanently poisons this writer.
    #[must_use]
    pub const fn is_poisoned(&self) -> bool {
        self.poisoned
    }

    /// Returns typed attribution for the byte ceiling that poisoned this
    /// writer, when the failure came from a streaming byte limit.
    #[must_use]
    pub const fn last_limit(&self) -> Option<StreamingLimitExceeded> {
        self.last_limit
    }

    fn refresh_output_bytes(&mut self) {
        self.output_bytes = self.output_counter.load(Ordering::Acquire);
    }

    fn invalid_limits(&self) -> Option<Error> {
        let reason = if self.limits.max_entries > ZIP32_MAX_ENTRIES {
            "max_entries must be below the ZIP32 entry-count ceiling"
        } else if self.limits.max_member_name_bytes > ZIP32_MAX_MEMBER_NAME_BYTES {
            "max_member_name_bytes exceeds the ZIP32 member-name field"
        } else if self.limits.max_metadata_bytes > self.limits.max_output_bytes {
            "max_metadata_bytes exceeds max_output_bytes"
        } else if self.limits.max_compressed_size >= ZIP32_MAX_VALUE {
            "max_compressed_size must be below the ZIP32 size ceiling"
        } else if self.limits.max_compressed_size > self.limits.max_output_bytes {
            "max_compressed_size exceeds max_output_bytes"
        } else if self.limits.max_entry_size >= ZIP32_MAX_VALUE {
            "max_entry_size must be below the ZIP32 size ceiling"
        } else if self.limits.max_total_size >= ZIP32_MAX_VALUE {
            "max_total_size must be below the ZIP32 size ceiling"
        } else if self.limits.max_output_bytes >= ZIP32_MAX_VALUE {
            "max_output_bytes must be below the ZIP32 size ceiling"
        } else if self.limits.max_output_bytes < MIN_STREAM_OUTPUT_BYTES {
            "max_output_bytes is too small for an empty ZIP archive"
        } else {
            return None;
        };
        Some(ErrorKind::InvalidInput { msg: reason.into() }.into())
    }

    fn ensure_usable(&self) -> Result<(), Error> {
        if self.poisoned {
            return Err(ErrorKind::InvalidInput {
                msg: format!(
                    "streaming archive writer is poisoned after {} accepted output bytes",
                    self.output_bytes
                ),
            }
            .into());
        }
        if let Some(error) = self.invalid_limits() {
            return Err(error);
        }
        Ok(())
    }

    fn poison(&mut self, error: Error) -> Error {
        let limit = streaming_limit_from_error(&error);
        self.last_limit = limit;
        self.refresh_output_bytes();
        self.poisoned = true;
        match limit {
            Some(StreamingLimitExceeded {
                resource: StreamingLimitResource::CompressedBytes,
                actual,
                maximum,
            }) => limit_error(LimitResource::CompressedSize, actual, maximum),
            Some(StreamingLimitExceeded {
                resource: StreamingLimitResource::OutputBytes,
                ..
            })
            | None => error,
        }
    }

    fn validate_streaming_entry(&self, name: &str) -> Result<(String, u64), Error> {
        self.ensure_usable()?;
        let raw_name = name.trim_end_matches('/');
        let raw_name_bytes = u64::try_from(raw_name.len()).unwrap_or(u64::MAX);
        let maximum_name_bytes = self
            .limits
            .max_member_name_bytes
            .min(ZIP32_MAX_MEMBER_NAME_BYTES);
        if raw_name_bytes > maximum_name_bytes {
            return Err(limit_error(
                LimitResource::MemberNameBytes,
                raw_name_bytes,
                maximum_name_bytes,
            ));
        }
        let path = ZipFilePath::from_str(raw_name);
        let normalized_name = canonical_member_name(path.as_str().to_string());
        let name_bytes = u64::try_from(normalized_name.len()).unwrap_or(u64::MAX);
        if name_bytes > ZIP32_MAX_MEMBER_NAME_BYTES
            || name_bytes > self.limits.max_member_name_bytes
        {
            return Err(limit_error(
                LimitResource::MemberNameBytes,
                name_bytes,
                maximum_name_bytes,
            ));
        }

        if self.names.contains(&normalized_name) {
            return Err(ErrorKind::InvalidInput {
                msg: format!("duplicate normalized member name: {normalized_name}"),
            }
            .into());
        }

        let next_metadata = self.metadata_bytes.saturating_add(name_bytes);
        if next_metadata > self.limits.max_metadata_bytes {
            return Err(limit_error(
                LimitResource::MetadataBytes,
                next_metadata,
                self.limits.max_metadata_bytes,
            ));
        }

        let next_entries = self.entries.saturating_add(1);
        let max_entries = u64::try_from(self.limits.max_entries).unwrap_or(u64::MAX);
        let actual_entries = u64::try_from(next_entries).unwrap_or(u64::MAX);
        if next_entries > self.limits.max_entries {
            return Err(limit_error(
                LimitResource::FileCount,
                actual_entries,
                max_entries,
            ));
        }

        Ok((normalized_name, name_bytes))
    }

    fn validate_known_entry(
        &self,
        name: &str,
        uncompressed_bytes: u64,
    ) -> Result<(String, u64), Error> {
        let (normalized_name, name_bytes) = self.validate_streaming_entry(name)?;
        if uncompressed_bytes > self.limits.max_entry_size {
            return Err(limit_error(
                LimitResource::EntrySize,
                uncompressed_bytes,
                self.limits.max_entry_size,
            ));
        }
        let total = self
            .total_uncompressed_bytes
            .saturating_add(uncompressed_bytes);
        if total > self.limits.max_total_size {
            return Err(limit_error(
                LimitResource::TotalSize,
                total,
                self.limits.max_total_size,
            ));
        }
        Ok((normalized_name, name_bytes))
    }

    fn record_streaming_entry(&mut self, normalized_name: String, name_bytes: u64) {
        let inserted = self.names.insert(normalized_name);
        debug_assert!(inserted);
        self.entries += 1;
        self.metadata_bytes += name_bytes;
    }

    fn reserve_streaming_entry(&mut self) -> Result<(), Error> {
        self.names.try_reserve(1).map_err(|error| {
            ErrorKind::InvalidInput {
                msg: format!("could not reserve streaming ZIP member-name index: {error}"),
            }
            .into()
        })
    }

    fn copy_stream<R: Read, O: Write>(
        reader: &mut R,
        output: &mut O,
        max_entry_size: u64,
        max_total_size: u64,
        committed_total: u64,
    ) -> Result<u64, Error> {
        let mut buffer = [0u8; STREAM_COPY_BUFFER_SIZE];
        let mut accepted_uncompressed = 0_u64;
        loop {
            let read = match reader.read(&mut buffer) {
                Ok(read) => read,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) => return Err(error.into()),
            };
            if read == 0 {
                return Ok(accepted_uncompressed);
            }
            if read > buffer.len() {
                return Err(ErrorKind::InvalidInput {
                    msg: "stream source returned more bytes than requested".to_string(),
                }
                .into());
            }
            let read = u64::try_from(read).unwrap_or(u64::MAX);
            let next_entry = accepted_uncompressed.saturating_add(read);
            if next_entry > max_entry_size {
                return Err(limit_error(
                    LimitResource::EntrySize,
                    next_entry,
                    max_entry_size,
                ));
            }
            let next_total = committed_total
                .saturating_add(accepted_uncompressed)
                .saturating_add(read);
            if next_total > max_total_size {
                return Err(limit_error(
                    LimitResource::TotalSize,
                    next_total,
                    max_total_size,
                ));
            }
            output.write_all(&buffer[..usize::try_from(read).unwrap_or(usize::MAX)])?;
            accepted_uncompressed = next_entry;
        }
    }

    fn write_reader<R: Read>(
        &mut self,
        name: &str,
        mut reader: R,
        compression_method: CompressionMethod,
    ) -> Result<(), Error> {
        let (normalized_name, name_bytes) = self.validate_streaming_entry(name)?;
        self.reserve_streaming_entry()?;
        let started = self
            .archive
            .new_file(&normalized_name)
            .compression_method(compression_method)
            .start();
        let (mut entry, config) = match started {
            Ok(started) => started,
            Err(error) => return Err(self.poison(error)),
        };

        let max_entry_size = self.limits.max_entry_size;
        let max_total_size = self.limits.max_total_size;
        let max_compressed_size = self.limits.max_compressed_size;
        let committed_total = self.total_uncompressed_bytes;

        let result = match compression_method {
            CompressionMethod::Store => (|| {
                let (accepted, descriptor) = {
                    let mut limited_entry =
                        LimitedEntryWriter::new(&mut entry, max_compressed_size);
                    let mut data_writer = config.wrap(&mut limited_entry);
                    let accepted = Self::copy_stream(
                        &mut reader,
                        &mut data_writer,
                        max_entry_size,
                        max_total_size,
                        committed_total,
                    )?;
                    let (_, descriptor) = data_writer.finish()?;
                    limited_entry.ensure_within_limit()?;
                    (accepted, descriptor)
                };
                entry.finish(descriptor)?;
                Ok(accepted)
            })(),
            CompressionMethod::Deflate => (|| {
                let (accepted, descriptor) = {
                    let mut limited_entry =
                        LimitedEntryWriter::new(&mut entry, max_compressed_size);
                    let encoder = DeflateEncoder::new(&mut limited_entry, Compression::default());
                    let mut data_writer = config.wrap(encoder);
                    let accepted = Self::copy_stream(
                        &mut reader,
                        &mut data_writer,
                        max_entry_size,
                        max_total_size,
                        committed_total,
                    )?;
                    let (encoder, descriptor) = data_writer.finish()?;
                    encoder.finish()?;
                    limited_entry.ensure_within_limit()?;
                    (accepted, descriptor)
                };
                entry.finish(descriptor)?;
                Ok(accepted)
            })(),
            other => Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                other.as_id().as_u16(),
            ))),
        };

        let accepted_uncompressed = match result {
            Ok(accepted) => accepted,
            Err(error) => return Err(self.poison(error)),
        };
        self.total_uncompressed_bytes = self
            .total_uncompressed_bytes
            .saturating_add(accepted_uncompressed);
        self.record_streaming_entry(normalized_name, name_bytes);
        self.refresh_output_bytes();
        Ok(())
    }

    /// Starts a consuming, bounded entry writer.
    ///
    /// The archive writer is moved into the returned entry. After all
    /// uncompressed payload bytes have been written, call
    /// [`StreamingArchiveEntry::finish`] to recover the archive writer and
    /// continue with another member. Store and Deflate are the only supported
    /// Office transport methods.
    pub fn start_entry(
        mut self,
        name: &str,
        compression_method: CompressionMethod,
    ) -> Result<StreamingArchiveEntry<W>, StreamingArchiveFailure> {
        let (normalized_name, name_bytes) = match self.validate_streaming_entry(name) {
            Ok(value) => value,
            Err(error) => {
                return Err(StreamingArchiveFailure {
                    limit: streaming_limit_from_error(&error),
                    error,
                    progress: self.progress(),
                });
            },
        };
        if !matches!(
            compression_method,
            CompressionMethod::Store | CompressionMethod::Deflate
        ) {
            return Err(StreamingArchiveFailure {
                error: ErrorKind::UnsupportedCompressionMethod(compression_method.as_id().as_u16())
                    .into(),
                progress: self.progress(),
                limit: None,
            });
        }
        if let Err(error) = self.reserve_streaming_entry() {
            return Err(StreamingArchiveFailure {
                error,
                progress: self.progress(),
                limit: None,
            });
        }

        let StreamingArchiveWriter {
            archive,
            limits,
            entries,
            metadata_bytes,
            total_uncompressed_bytes,
            output_bytes,
            poisoned: _,
            names,
            last_limit,
            output_counter,
        } = self;
        let entry = match archive
            .start_file_owned(&normalized_name, compression_method)
            .map(|entry| entry.with_compressed_limit(limits.max_compressed_size))
        {
            Ok(entry) => entry,
            Err(error) => {
                let limit = streaming_limit_from_error(&error).or(last_limit);
                let error = match limit {
                    Some(StreamingLimitExceeded {
                        resource: StreamingLimitResource::CompressedBytes,
                        actual,
                        maximum,
                    }) => limit_error(LimitResource::CompressedSize, actual, maximum),
                    _ => error,
                };
                return Err(StreamingArchiveFailure {
                    error,
                    progress: StreamingArchiveProgress {
                        output_bytes: output_counter.load(Ordering::Acquire),
                        poisoned: true,
                    },
                    limit,
                });
            },
        };

        Ok(StreamingArchiveEntry {
            entry: Some(entry),
            limits,
            entries,
            metadata_bytes,
            total_uncompressed_bytes,
            names,
            normalized_name,
            name_bytes,
            uncompressed_bytes: 0,
            output_bytes,
            poisoned: false,
            last_limit: None,
            failure: None,
            output_counter,
        })
    }

    /// Write a file without compression (stored).
    pub fn write_stored(&mut self, name: &str, data: &[u8]) -> Result<(), Error> {
        let data_bytes = u64::try_from(data.len()).unwrap_or(u64::MAX);
        let (normalized_name, name_bytes) = self.validate_known_entry(name, data_bytes)?;
        self.reserve_streaming_entry()?;
        if data_bytes > self.limits.max_compressed_size {
            return Err(limit_error(
                LimitResource::CompressedSize,
                data_bytes,
                self.limits.max_compressed_size,
            ));
        }
        match self.archive.write_stored_file(&normalized_name, data) {
            Ok(()) => {
                self.total_uncompressed_bytes =
                    self.total_uncompressed_bytes.saturating_add(data_bytes);
                self.record_streaming_entry(normalized_name, name_bytes);
                self.refresh_output_bytes();
                Ok(())
            },
            Err(error) => Err(self.poison(error)),
        }
    }

    /// Write a file with Deflate compression.
    pub fn write_deflated(&mut self, name: &str, data: &[u8]) -> Result<(), Error> {
        let data_bytes = u64::try_from(data.len()).unwrap_or(u64::MAX);
        let (normalized_name, name_bytes) = self.validate_known_entry(name, data_bytes)?;
        self.reserve_streaming_entry()?;
        let started = self
            .archive
            .new_file(&normalized_name)
            .compression_method(CompressionMethod::Deflate)
            .start();
        let result = match started {
            Ok((mut entry, config)) => (|| {
                let descriptor = {
                    let mut limited_entry =
                        LimitedEntryWriter::new(&mut entry, self.limits.max_compressed_size);
                    let encoder = DeflateEncoder::new(&mut limited_entry, Compression::default());
                    let mut writer = config.wrap(encoder);
                    writer.write_all(data)?;
                    let (encoder, desc) = writer.finish()?;
                    encoder.finish()?;
                    limited_entry.ensure_within_limit()?;
                    desc
                };
                entry.finish(descriptor)?;
                Ok(())
            })(),
            Err(error) => Err(error),
        };
        if let Err(error) = result {
            return Err(self.poison(error));
        }
        self.total_uncompressed_bytes = self.total_uncompressed_bytes.saturating_add(data_bytes);
        self.record_streaming_entry(normalized_name, name_bytes);
        self.refresh_output_bytes();
        Ok(())
    }

    /// Consume a reader value into a stored ZIP member.
    ///
    /// The source is read incrementally and is not retained after this method
    /// returns.  The output uses a data descriptor, so `W` only needs to
    /// implement [`Write`], not [`std::io::Seek`].
    pub fn write_stored_stream<R: Read>(&mut self, name: &str, reader: R) -> Result<(), Error> {
        self.write_reader(name, reader, CompressionMethod::Store)
    }

    /// Consume a reader value into a Deflate-compressed ZIP member.
    ///
    /// The source is read incrementally and is not retained after this method
    /// returns.  Compression and CRC state are bounded to the encoder's
    /// working buffers plus the central-directory metadata retained by the
    /// archive.
    pub fn write_deflated_stream<R: Read>(&mut self, name: &str, reader: R) -> Result<(), Error> {
        self.write_reader(name, reader, CompressionMethod::Deflate)
    }

    /// Consume a reader value with one of the supported Office ZIP methods.
    ///
    /// [`CompressionMethod::Store`] and [`CompressionMethod::Deflate`] are
    /// supported.  Other methods are rejected before any archive bytes are
    /// written.
    pub fn write_stream<R: Read>(
        &mut self,
        name: &str,
        reader: R,
        compression_method: CompressionMethod,
    ) -> Result<(), Error> {
        if !matches!(
            compression_method,
            CompressionMethod::Store | CompressionMethod::Deflate
        ) {
            return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                compression_method.as_id().as_u16(),
            )));
        }
        self.write_reader(name, reader, compression_method)
    }

    /// Alias for [`Self::write_stored_stream`] using reader-oriented naming.
    pub fn write_stored_reader<R: Read>(&mut self, name: &str, reader: R) -> Result<(), Error> {
        self.write_stored_stream(name, reader)
    }

    /// Alias for [`Self::write_deflated_stream`] using reader-oriented naming.
    pub fn write_deflated_reader<R: Read>(&mut self, name: &str, reader: R) -> Result<(), Error> {
        self.write_deflated_stream(name, reader)
    }

    /// Finish writing the archive.
    pub fn finish(self) -> Result<W, Error> {
        self.finish_with_progress()
            .map(|(writer, _progress)| writer)
            .map_err(StreamingArchiveFailure::into_error)
    }

    /// Finish writing the archive while preserving typed progress on failure.
    ///
    /// This is useful for caller-owned non-atomic sinks: if central-directory
    /// or final-flush output fails, the returned error still reports the
    /// content-free number of bytes accepted by the sink.
    pub fn finish_with_progress(
        mut self,
    ) -> Result<(W, StreamingArchiveProgress), StreamingArchiveFailure> {
        if let Err(error) = self.ensure_usable() {
            return Err(StreamingArchiveFailure {
                error,
                progress: self.progress(),
                limit: self.last_limit,
            });
        }
        self.refresh_output_bytes();
        match self.archive.finish() {
            Ok(output) => {
                let progress = StreamingArchiveProgress {
                    output_bytes: self.output_counter.load(Ordering::Acquire),
                    poisoned: false,
                };
                Ok((output.into_inner(), progress))
            },
            Err(error) => Err(StreamingArchiveFailure {
                limit: streaming_limit_from_error(&error).or(self.last_limit),
                error,
                progress: StreamingArchiveProgress {
                    output_bytes: self.output_counter.load(Ordering::Acquire),
                    poisoned: true,
                },
            }),
        }
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
        let lookup = lookup_member_name(name);
        if lookup.explicit_directory {
            return Err(ErrorKind::FileNotFound(lookup.name).into());
        }
        let normalized = lookup.name;

        // Fast path: check if already cached (read lock)
        {
            let cache = self.cache.read().unwrap();
            if let Some(data) = cache.get(&normalized) {
                return Ok(std::sync::Arc::clone(data));
            }
        }

        // Slow path: decompress and cache (write lock)
        let data = self.inner.read(name)?;
        let arc = std::sync::Arc::new(data);

        {
            let mut cache = self.cache.write().unwrap();
            // Double-check in case another thread cached it while we were decompressing
            if let Some(existing) = cache.get(&normalized) {
                return Ok(std::sync::Arc::clone(existing));
            }
            cache.insert(normalized, std::sync::Arc::clone(&arc));
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
    use std::io::{self, Cursor};

    #[derive(Debug)]
    struct ShortWriter {
        bytes: Vec<u8>,
        max_write: usize,
    }

    impl ShortWriter {
        fn new(max_write: usize) -> Self {
            assert!(max_write > 0);
            Self {
                bytes: Vec::new(),
                max_write,
            }
        }
    }

    impl Write for ShortWriter {
        fn write(&mut self, data: &[u8]) -> io::Result<usize> {
            let written = data.len().min(self.max_write);
            self.bytes.extend_from_slice(&data[..written]);
            Ok(written)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[derive(Debug)]
    struct FailingWriter {
        bytes: Vec<u8>,
        fail_after: usize,
    }

    impl FailingWriter {
        fn new(fail_after: usize) -> Self {
            Self {
                bytes: Vec::new(),
                fail_after,
            }
        }
    }

    impl Write for FailingWriter {
        fn write(&mut self, data: &[u8]) -> io::Result<usize> {
            if self.bytes.len() >= self.fail_after {
                return Err(io::Error::new(io::ErrorKind::BrokenPipe, "sink failed"));
            }
            let available = self.fail_after - self.bytes.len();
            let written = data.len().min(available);
            self.bytes.extend_from_slice(&data[..written]);
            Ok(written)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[derive(Debug)]
    struct NestedSinkSource;

    impl std::fmt::Display for NestedSinkSource {
        fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            formatter.write_str("nested sink failure")
        }
    }

    impl std::error::Error for NestedSinkSource {}

    #[derive(Debug)]
    struct NestedFailingWriter {
        bytes: Vec<u8>,
        fail_after: usize,
    }

    impl NestedFailingWriter {
        fn new(fail_after: usize) -> Self {
            Self {
                bytes: Vec::new(),
                fail_after,
            }
        }
    }

    impl Write for NestedFailingWriter {
        fn write(&mut self, data: &[u8]) -> io::Result<usize> {
            if self.bytes.len() >= self.fail_after {
                return Err(io::Error::new(io::ErrorKind::BrokenPipe, NestedSinkSource));
            }
            let available = self.fail_after - self.bytes.len();
            let written = data.len().min(available);
            self.bytes.extend_from_slice(&data[..written]);
            Ok(written)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[derive(Debug)]
    struct ZeroWriter;

    impl Write for ZeroWriter {
        fn write(&mut self, _data: &[u8]) -> io::Result<usize> {
            Ok(0)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[derive(Debug)]
    enum ReadStep {
        Interrupted,
        Bytes(Vec<u8>),
        Error,
    }

    #[derive(Debug)]
    struct ScriptedReader {
        steps: std::collections::VecDeque<ReadStep>,
    }

    impl ScriptedReader {
        fn new(steps: impl IntoIterator<Item = ReadStep>) -> Self {
            Self {
                steps: steps.into_iter().collect(),
            }
        }
    }

    impl Read for ScriptedReader {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            match self
                .steps
                .pop_front()
                .unwrap_or(ReadStep::Bytes(Vec::new()))
            {
                ReadStep::Interrupted => {
                    Err(io::Error::new(io::ErrorKind::Interrupted, "try again"))
                },
                ReadStep::Bytes(bytes) => {
                    let count = bytes.len().min(output.len());
                    output[..count].copy_from_slice(&bytes[..count]);
                    if count < bytes.len() {
                        self.steps
                            .push_front(ReadStep::Bytes(bytes[count..].to_vec()));
                    }
                    Ok(count)
                },
                ReadStep::Error => Err(io::Error::new(io::ErrorKind::InvalidData, "source failed")),
            }
        }
    }

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
    fn consuming_entry_writer_recovers_bounded_archive_for_both_methods() {
        let mut writer = StreamingArchiveWriter::new();
        let mut entry = writer
            .start_entry("first.bin", CompressionMethod::Store)
            .unwrap();
        entry.write_all(b"first payload").unwrap();
        writer = entry.finish().unwrap();

        let mut entry = writer
            .start_entry("second.bin", CompressionMethod::Deflate)
            .unwrap();
        entry.write_all(b"second payload").unwrap();
        writer = entry.finish().unwrap();

        let bytes = writer.finish_to_bytes().unwrap();
        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(reader.read("first.bin").unwrap(), b"first payload");
        assert_eq!(reader.read("second.bin").unwrap(), b"second payload");
    }

    #[test]
    fn consuming_entry_writer_preserves_limits_and_poison_progress() {
        let limits = StreamingArchiveLimits::new(4, 16, 4096).with_byte_limits(3, 8, 4096);
        let writer = StreamingArchiveWriter::with_limits(limits);
        let mut entry = writer
            .start_entry("bounded.bin", CompressionMethod::Store)
            .unwrap();
        let error = entry.write_all(b"over").unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::Other);
        assert!(entry.is_poisoned());
        assert!(entry.progress().is_poisoned());
        let failure = match entry.finish() {
            Ok(_) => panic!("poisoned entry unexpectedly finished"),
            Err(failure) => failure,
        };
        assert!(matches!(
            failure.error().kind(),
            ErrorKind::LimitExceeded {
                resource: LimitResource::EntrySize,
                actual: 4,
                maximum: 3,
            }
        ));
        assert!(failure.progress().is_poisoned());
    }

    #[test]
    fn consuming_entry_writer_rejects_duplicate_before_header_and_handles_short_sink() {
        let mut writer = StreamingArchiveWriter::new();
        let mut entry = writer
            .start_entry("dir/../same.bin", CompressionMethod::Deflate)
            .unwrap();
        entry.write_all(b"first").unwrap();
        writer = entry.finish().unwrap();

        let duplicate = match writer.start_entry("same.bin", CompressionMethod::Store) {
            Ok(_) => panic!("duplicate normalized member unexpectedly started"),
            Err(failure) => failure,
        };
        assert!(matches!(
            duplicate.error().kind(),
            ErrorKind::InvalidInput { .. }
        ));
        assert!(!duplicate.progress().is_poisoned());

        let mut sink = ShortWriter::new(2);
        let mut writer = StreamingArchiveWriter::with_writer(&mut sink);
        let mut entry = writer
            .start_entry("short.bin", CompressionMethod::Deflate)
            .unwrap();
        entry.write_all(b"short sink payload").unwrap();
        writer = entry.finish().unwrap();
        writer.finish().unwrap();
        let reader = ArchiveReader::new(&sink.bytes).unwrap();
        assert_eq!(reader.read("short.bin").unwrap(), b"short sink payload");
    }

    #[test]
    fn consuming_entry_writer_preflights_raw_names_and_drops_incomplete_publication() {
        let limits = StreamingArchiveLimits::new(4, 8, 64);
        let mut sink = ShortWriter::new(3);
        let failure = match StreamingArchiveWriter::with_writer_and_limits(&mut sink, limits)
            .start_entry(&"x".repeat(9), CompressionMethod::Store)
        {
            Ok(_) => panic!("oversized raw name unexpectedly started"),
            Err(failure) => failure,
        };
        assert!(matches!(
            failure.error().kind(),
            ErrorKind::LimitExceeded {
                resource: LimitResource::MemberNameBytes,
                actual: 9,
                maximum: 8,
            }
        ));
        assert_eq!(failure.progress().output_bytes(), 0);
        assert!(sink.bytes.is_empty());

        let mut sink = ShortWriter::new(3);
        {
            let writer = StreamingArchiveWriter::with_writer(&mut sink);
            let mut entry = writer
                .start_entry("unfinished.bin", CompressionMethod::Deflate)
                .unwrap();
            entry.write_all(b"partial payload").unwrap();
        }
        assert!(ArchiveReader::new(&sink.bytes).is_err());
    }

    #[test]
    fn consuming_entry_writer_enforces_store_compressed_and_aggregate_boundaries() {
        let base = StreamingArchiveLimits::new(4, 32, 128).with_byte_limits(16, 16, 4096);
        let mut writer = StreamingArchiveWriter::with_limits(base.with_compressed_size_limit(3));
        let mut entry = writer
            .start_entry("exact.bin", CompressionMethod::Store)
            .unwrap();
        entry.write_all(b"abc").unwrap();
        writer = entry.finish().unwrap();
        assert_eq!(writer.total_uncompressed_bytes(), 3);

        let mut entry = writer
            .start_entry("over.bin", CompressionMethod::Store)
            .unwrap();
        let error = entry.write_all(b"abcd").unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::Other);
        let failure = match entry.finish() {
            Ok(_) => panic!("compressed limit unexpectedly accepted payload"),
            Err(failure) => failure,
        };
        assert!(matches!(
            failure.error().kind(),
            ErrorKind::LimitExceeded {
                resource: LimitResource::CompressedSize,
                actual: 4,
                maximum: 3,
            }
        ));

        let limits = StreamingArchiveLimits::new(4, 32, 128).with_byte_limits(16, 3, 4096);
        let mut writer = StreamingArchiveWriter::with_limits(limits);
        let mut entry = writer
            .start_entry("first.bin", CompressionMethod::Store)
            .unwrap();
        entry.write_all(b"ab").unwrap();
        writer = entry.finish().unwrap();
        let mut entry = writer
            .start_entry("second.bin", CompressionMethod::Store)
            .unwrap();
        let error = entry.write_all(b"cd").unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::Other);
        let failure = match entry.finish() {
            Ok(_) => panic!("aggregate limit unexpectedly accepted payload"),
            Err(failure) => failure,
        };
        assert!(matches!(
            failure.error().kind(),
            ErrorKind::LimitExceeded {
                resource: LimitResource::TotalSize,
                actual: 4,
                maximum: 3,
            }
        ));
    }

    #[test]
    fn consuming_entry_writer_reports_sink_failure_during_descriptor_finish() {
        let mut sink = FailingWriter::new(52);
        let mut entry = StreamingArchiveWriter::with_writer(&mut sink)
            .start_entry("descriptor.bin", CompressionMethod::Store)
            .unwrap();
        entry.write_all(b"payload").unwrap();
        let failure = match entry.finish() {
            Ok(_) => panic!("descriptor sink failure unexpectedly succeeded"),
            Err(failure) => failure,
        };
        assert!(matches!(
            failure.error().kind(),
            ErrorKind::IO(_) | ErrorKind::Io(_)
        ));
        assert!(failure.progress().is_poisoned());
        assert!(failure.progress().output_bytes() <= 52);
    }

    #[test]
    fn consuming_entry_writer_retains_nested_sink_source() {
        // The stored local header is 30 bytes plus the member name.  Permit
        // that prefix, then fail on the first payload write so the entry
        // capture path—not descriptor finalization—owns the original error.
        let name = "nested.bin";
        let mut sink = NestedFailingWriter::new(30 + name.len());
        let mut entry = StreamingArchiveWriter::with_writer(&mut sink)
            .start_entry(name, CompressionMethod::Store)
            .unwrap();
        let immediate = entry.write(b"payload").unwrap_err();
        assert_eq!(immediate.kind(), io::ErrorKind::BrokenPipe);

        let failure = match entry.finish() {
            Ok(_) => panic!("nested sink failure unexpectedly finished"),
            Err(failure) => failure,
        };
        let error = match failure.error().kind() {
            ErrorKind::IO(error) | ErrorKind::Io(error) => error,
            other => panic!("expected retained I/O error, got {other:?}"),
        };
        assert!(
            error
                .get_ref()
                .and_then(|source| source.downcast_ref::<NestedSinkSource>())
                .is_some(),
            "nested sink source was lost"
        );
        let source = std::error::Error::source(&failure).expect("failure source");
        assert!(source.downcast_ref::<io::Error>().is_some());
    }

    #[test]
    fn consuming_entry_writer_enforces_deflate_and_output_exact_boundaries() {
        let mut probe = StreamingArchiveWriter::new();
        let mut entry = probe.start_entry("x", CompressionMethod::Deflate).unwrap();
        entry.write_all(b"payload").unwrap();
        probe = entry.finish().unwrap();
        let expected_output = probe.finish_to_bytes().unwrap();

        let limits = StreamingArchiveLimits::new(4, 32, 1).with_byte_limits(
            64,
            64,
            expected_output.len() as u64,
        );
        let mut exact = StreamingArchiveWriter::with_limits(limits);
        let mut entry = exact.start_entry("x", CompressionMethod::Deflate).unwrap();
        entry.write_all(b"payload").unwrap();
        exact = entry.finish().unwrap();
        assert_eq!(exact.finish_to_bytes().unwrap(), expected_output);

        let limits = limits.with_byte_limits(64, 64, expected_output.len() as u64 - 1);
        let over = StreamingArchiveWriter::with_limits(limits);
        let mut entry = over.start_entry("x", CompressionMethod::Deflate).unwrap();
        entry.write_all(b"payload").unwrap();
        let over = match entry.finish() {
            Ok(writer) => writer,
            Err(_) => panic!("entry output limit failed before archive finalization"),
        };
        let failure = match over.finish_with_progress() {
            Ok(_) => panic!("output limit unexpectedly accepted one over"),
            Err(failure) => failure,
        };
        assert_eq!(
            failure.limit().map(StreamingLimitExceeded::resource),
            Some(StreamingLimitResource::OutputBytes)
        );

        let limited = StreamingArchiveWriter::with_limits(
            StreamingArchiveLimits::new(4, 32, 128)
                .with_byte_limits(64, 64, 4096)
                .with_compressed_size_limit(0),
        );
        let entry = limited
            .start_entry("empty", CompressionMethod::Deflate)
            .unwrap();
        let failure = match entry.finish() {
            Ok(_) => panic!("compressed limit unexpectedly accepted an empty deflate stream"),
            Err(failure) => failure,
        };
        let compressed_actual = failure
            .limit()
            .expect("compressed limit attribution")
            .actual();
        assert!(compressed_actual > 0);

        let mut exact = StreamingArchiveWriter::with_limits(
            StreamingArchiveLimits::new(4, 32, 128)
                .with_byte_limits(64, 64, 4096)
                .with_compressed_size_limit(compressed_actual.saturating_add(2)),
        );
        let entry = exact
            .start_entry("empty", CompressionMethod::Deflate)
            .unwrap();
        exact = entry.finish().unwrap();
        exact.finish_to_bytes().unwrap();
    }

    #[test]
    fn owned_stream_entries_handle_empty_and_large_members() {
        let large = (0..(1024 * 1024))
            .map(|index| (index as u8).wrapping_mul(31))
            .collect::<Vec<_>>();
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored_stream("empty.bin", Cursor::new(Vec::<u8>::new()))
            .unwrap();
        writer
            .write_deflated_stream("large.bin", Cursor::new(large.clone()))
            .unwrap();

        assert_eq!(writer.entry_count(), 2);
        assert_eq!(writer.metadata_bytes(), "empty.binlarge.bin".len() as u64);
        let bytes = writer.finish_to_bytes().unwrap();

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(reader.read("empty.bin").unwrap(), Vec::<u8>::new());
        assert_eq!(reader.read("large.bin").unwrap(), large);
        assert!(local_member_has_data_descriptor(&bytes, b"empty.bin"));
        assert!(local_member_has_data_descriptor(&bytes, b"large.bin"));
    }

    #[test]
    fn owned_stream_entries_preserve_order_crc_and_descriptor_metadata() {
        let first = b"first stream payload".to_vec();
        let second = b"second stream payload with deflate".to_vec();
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored_reader("first.bin", Cursor::new(first.clone()))
            .unwrap();
        writer
            .write_deflated_reader("second.bin", Cursor::new(second.clone()))
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let archive = ZipArchive::from_slice(&bytes).unwrap();
        let mut entries = archive.entries();
        let first_entry = entries.next_entry().unwrap().unwrap();
        assert_eq!(first_entry.file_path().as_ref(), b"first.bin");
        assert_eq!(first_entry.crc32(), crate::crc32(&first));
        assert!(first_entry.has_data_descriptor());
        let second_entry = entries.next_entry().unwrap().unwrap();
        assert_eq!(second_entry.file_path().as_ref(), b"second.bin");
        assert_eq!(second_entry.crc32(), crate::crc32(&second));
        assert!(second_entry.has_data_descriptor());
        assert!(entries.next_entry().unwrap().is_none());

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(
            reader.file_names().collect::<Vec<_>>(),
            ["first.bin", "second.bin"]
        );
        assert_eq!(reader.read("first.bin").unwrap(), first);
        assert_eq!(reader.read("second.bin").unwrap(), second);
    }

    #[test]
    fn owned_stream_output_accepts_short_non_seek_writes() {
        let mut sink = ShortWriter::new(3);
        {
            let mut writer = StreamingArchiveWriter::with_writer(&mut sink);
            writer
                .write_deflated_stream("short.xml", Cursor::new(b"short sink".to_vec()))
                .unwrap();
            writer.finish().unwrap();
        }

        let reader = ArchiveReader::new(&sink.bytes).unwrap();
        assert_eq!(reader.read("short.xml").unwrap(), b"short sink");
    }

    #[test]
    fn owned_stream_output_reports_failing_sink_after_partial_output() {
        let mut sink = FailingWriter::new(64);
        let error = {
            let mut writer = StreamingArchiveWriter::with_writer(&mut sink);
            writer
                .write_deflated_stream("failing.xml", Cursor::new(vec![b'x'; 4096]))
                .unwrap_err()
        };

        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert!(sink.bytes.len() <= sink.fail_after);
    }

    #[test]
    fn owned_stream_output_is_deterministic_and_rejects_unsupported_methods_early() {
        fn build_archive() -> Vec<u8> {
            let mut writer = StreamingArchiveWriter::new();
            writer
                .write_stored_stream("a", Cursor::new(b"stored".to_vec()))
                .unwrap();
            writer
                .write_stream(
                    "b",
                    Cursor::new(b"deflated".to_vec()),
                    CompressionMethod::Deflate,
                )
                .unwrap();
            writer.finish_to_bytes().unwrap()
        }

        let first = build_archive();
        let second = build_archive();
        assert_eq!(first, second);

        let mut writer = StreamingArchiveWriter::new();
        let error = writer
            .write_stream(
                "unsupported",
                Cursor::new(b"payload".to_vec()),
                CompressionMethod::Bzip2,
            )
            .unwrap_err();
        assert!(matches!(
            error.kind(),
            ErrorKind::UnsupportedCompressionMethod(12)
        ));
        assert_eq!(writer.entry_count(), 0);
        writer
            .write_stored_stream("after-error", Cursor::new(b"ok".to_vec()))
            .unwrap();
        assert!(ArchiveReader::new(&writer.finish_to_bytes().unwrap()).is_ok());
    }

    #[test]
    fn owned_stream_output_enforces_finite_metadata_limits() {
        let mut writer = StreamingArchiveWriter::with_limits(StreamingArchiveLimits::new(1, 8, 16));
        writer
            .write_stored_stream("first", Cursor::new(b"one".to_vec()))
            .unwrap();
        assert_eq!(writer.metadata_bytes(), 5);

        let error = writer
            .write_stored_stream("second", Cursor::new(b"two".to_vec()))
            .unwrap_err();
        assert_limit(error, LimitResource::FileCount, 2, 1);

        let mut name_limited =
            StreamingArchiveWriter::with_limits(StreamingArchiveLimits::new(4, 3, 16));
        let error = name_limited
            .write_stored_stream("four", Cursor::new(b"payload".to_vec()))
            .unwrap_err();
        assert_limit(error, LimitResource::MemberNameBytes, 4, 3);

        let mut metadata_limited =
            StreamingArchiveWriter::with_limits(StreamingArchiveLimits::new(4, 16, 5));
        metadata_limited
            .write_stored_stream("one", Cursor::new(b"one".to_vec()))
            .unwrap();
        let error = metadata_limited
            .write_stored_stream("three", Cursor::new(b"three".to_vec()))
            .unwrap_err();
        assert_limit(error, LimitResource::MetadataBytes, 8, 5);
    }

    #[test]
    fn owned_stream_output_enforces_exact_and_one_over_byte_limits() {
        let mut exact = StreamingArchiveWriter::new();
        exact
            .write_stored_stream("exact", Cursor::new(b"abc".to_vec()))
            .unwrap();
        assert_eq!(exact.total_uncompressed_bytes, 3);

        let mut entry_limited = StreamingArchiveWriter::with_limits(
            StreamingArchiveLimits::new(4, 16, 16).with_byte_limits(3, 8, 4096),
        );
        entry_limited
            .write_stored_stream("exact", Cursor::new(b"abc".to_vec()))
            .unwrap();
        let error = entry_limited
            .write_stored_stream("over", Cursor::new(b"over".to_vec()))
            .unwrap_err();
        assert_limit(error, LimitResource::EntrySize, 4, 3);

        let mut aggregate_limited = StreamingArchiveWriter::with_limits(
            StreamingArchiveLimits::new(4, 16, 16).with_byte_limits(8, 3, 4096),
        );
        aggregate_limited
            .write_stored_stream("first", Cursor::new(b"abc".to_vec()))
            .unwrap();
        let error = aggregate_limited
            .write_stored_stream("second", Cursor::new(b"d".to_vec()))
            .unwrap_err();
        assert_limit(error, LimitResource::TotalSize, 4, 3);

        let mut output_probe = StreamingArchiveWriter::new();
        output_probe.write_stored("one", b"payload").unwrap();
        let output = output_probe.finish_to_bytes().unwrap();
        let expected_output_bytes = output.len() as u64;

        let output_limits =
            StreamingArchiveLimits::new(4, 16, 16).with_byte_limits(8, 8, expected_output_bytes);
        let mut output_exact = StreamingArchiveWriter::with_limits(output_limits);
        output_exact.write_stored("one", b"payload").unwrap();
        assert_eq!(output_exact.finish_to_bytes().unwrap(), output);

        let output_limits = output_limits.with_byte_limits(8, 8, expected_output_bytes - 1);
        let mut output_over = StreamingArchiveWriter::with_limits(output_limits);
        output_over.write_stored("one", b"payload").unwrap();
        let failure = output_over.finish_with_progress().unwrap_err();
        assert!(matches!(
            failure.error().kind(),
            ErrorKind::IO(_) | ErrorKind::Io(_)
        ));
        let limit = failure.limit().expect("typed output limit");
        assert_eq!(limit.resource(), StreamingLimitResource::OutputBytes);
        assert_eq!(limit.maximum(), expected_output_bytes - 1);
        assert!(failure.progress().is_poisoned());
        assert!(failure.progress().output_bytes() < expected_output_bytes);
    }

    #[test]
    fn owned_stream_output_enforces_compressed_member_limits() {
        let limits = StreamingArchiveLimits::new(4, 16, 64)
            .with_byte_limits(16, 32, 4096)
            .with_compressed_size_limit(3);

        let mut exact = StreamingArchiveWriter::with_limits(limits);
        exact
            .write_stored_stream("exact", Cursor::new(b"abc".to_vec()))
            .unwrap();
        let bytes = exact.finish_to_bytes().unwrap();
        assert_eq!(
            ArchiveReader::new(&bytes).unwrap().read("exact").unwrap(),
            b"abc"
        );

        let mut one_over = StreamingArchiveWriter::with_limits(limits);
        let error = one_over
            .write_stored_stream("over", Cursor::new(b"abcd".to_vec()))
            .unwrap_err();
        assert_limit(error, LimitResource::CompressedSize, 4, 3);
        let limit = one_over.last_limit().expect("typed compressed limit");
        assert_eq!(limit.resource(), StreamingLimitResource::CompressedBytes);
        assert_eq!(limit.actual(), 4);
        assert_eq!(limit.maximum(), 3);
        assert!(one_over.is_poisoned());

        let mut deflated = StreamingArchiveWriter::with_limits(
            StreamingArchiveLimits::new(4, 16, 64)
                .with_byte_limits(16, 32, 4096)
                .with_compressed_size_limit(1),
        );
        let error = deflated
            .write_deflated_stream("deflated", Cursor::new(b"payload".to_vec()))
            .unwrap_err();
        match error.kind() {
            ErrorKind::LimitExceeded {
                resource: LimitResource::CompressedSize,
                actual,
                maximum: 1,
            } => assert!(*actual > 1),
            other => panic!("expected compressed limit error, got {other:?}"),
        }
        assert!(deflated.is_poisoned());
    }

    #[test]
    fn owned_stream_deflate_limit_applies_to_compressed_bytes_not_input_bytes() {
        let payload = vec![b'a'; 4096];
        let mut probe = StreamingArchiveWriter::new();
        probe.write_deflated("probe", &payload).unwrap();
        let bytes = probe.finish_to_bytes().unwrap();
        let archive = ZipArchive::from_slice(&bytes).unwrap();
        let entry = archive.entries().next_entry().unwrap().unwrap();
        let compressed = entry.compressed_size_hint();
        assert!(compressed < payload.len() as u64);

        let limits = StreamingArchiveLimits::new(4, 16, 64)
            .with_byte_limits(8192, 8192, 4096)
            .with_compressed_size_limit(compressed);
        let mut writer = StreamingArchiveWriter::with_limits(limits);
        writer.write_deflated("compressed", &payload).unwrap();
        let output = writer.finish_to_bytes().unwrap();
        assert_eq!(
            ArchiveReader::new(&output)
                .unwrap()
                .read("compressed")
                .unwrap(),
            payload
        );
    }

    #[test]
    fn owned_stream_output_rejects_duplicate_normalized_names_before_header() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("dir/../same", b"one").unwrap();
        let output_before_duplicate = writer.output_bytes();
        let error = writer
            .write_stored_stream("same/", Cursor::new(b"two".to_vec()))
            .unwrap_err();
        assert!(
            matches!(error.kind(), ErrorKind::InvalidInput { msg } if msg.contains("duplicate normalized member name"))
        );
        assert_eq!(writer.output_bytes(), output_before_duplicate);
        assert_eq!(writer.entry_count(), 1);
        assert!(!writer.is_poisoned());
        let bytes = writer.finish_to_bytes().unwrap();
        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(reader.file_names().collect::<Vec<_>>(), ["same"]);
        assert_eq!(reader.read("same").unwrap(), b"one");
    }

    #[test]
    fn owned_stream_output_handles_interrupted_and_source_errors_with_poisoning() {
        let mut writer = StreamingArchiveWriter::new();
        let reader =
            ScriptedReader::new([ReadStep::Interrupted, ReadStep::Bytes(b"accepted".to_vec())]);
        writer.write_stored_stream("ok", reader).unwrap();
        assert!(!writer.is_poisoned());

        let reader = ScriptedReader::new([ReadStep::Bytes(b"partial".to_vec()), ReadStep::Error]);
        let error = writer.write_stored_stream("bad", reader).unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert!(writer.is_poisoned());
        let progress = writer.progress();
        assert!(progress.is_poisoned());
        assert!(progress.output_bytes() > 0);

        let error = writer
            .write_stored_stream("after", Cursor::new(b"rejected".to_vec()))
            .unwrap_err();
        assert!(
            matches!(error.kind(), ErrorKind::InvalidInput { msg } if msg.contains("poisoned"))
        );
    }

    #[test]
    fn owned_stream_output_reports_write_zero_and_rejects_post_failure_finish() {
        let mut writer = StreamingArchiveWriter::with_writer_and_limits(
            ZeroWriter,
            StreamingArchiveLimits::default(),
        );
        let error = writer
            .write_stored_stream("zero", Cursor::new(b"payload".to_vec()))
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert!(writer.is_poisoned());
        assert_eq!(writer.output_bytes(), 0);

        let error = writer.finish().unwrap_err();
        assert!(
            matches!(error.kind(), ErrorKind::InvalidInput { msg } if msg.contains("poisoned"))
        );
    }

    #[test]
    fn streaming_limits_reject_zip32_overrides_before_output() {
        let invalid = StreamingArchiveLimits::new(usize::MAX, u64::MAX, u64::MAX).with_byte_limits(
            u64::MAX,
            u64::MAX,
            u64::MAX,
        );
        let mut writer = StreamingArchiveWriter::with_limits(invalid);
        let error = writer
            .write_stored_stream("x", Cursor::new(b"x".to_vec()))
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
        assert_eq!(writer.output_bytes(), 0);
        assert!(!writer.is_poisoned());
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
    fn archive_readers_reject_duplicate_directories_and_file_directory_collisions() {
        for entries in [
            vec![
                FixtureEntry::stored(b"folder/", b""),
                FixtureEntry::stored(b"./folder/", b""),
            ],
            vec![
                FixtureEntry::stored(b"folder", b"file"),
                FixtureEntry::stored(b"folder/", b""),
            ],
        ] {
            let bytes = fixture(&entries);
            assert!(matches!(
                ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED),
                Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { msg } if msg.contains("directory") || msg.contains("file/directory"))
            ));
            assert!(matches!(
                indexed_archive_result(bytes, ArchiveLimits::UNBOUNDED),
                Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { msg } if msg.contains("directory") || msg.contains("file/directory"))
            ));
        }
    }

    #[test]
    fn archive_readers_reject_lossy_utf8_name_collisions() {
        let bytes = fixture(&[
            FixtureEntry::stored(b"collision\xff.bin", b"one"),
            FixtureEntry::stored(b"collision\xfe.bin", b"two"),
        ]);

        for result in [
            ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).map(|_| ()),
            indexed_archive_result(bytes.clone(), ArchiveLimits::UNBOUNDED).map(|_| ()),
        ] {
            assert!(matches!(
                result,
                Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { msg } if msg.contains("duplicate normalized file names"))
            ));
        }
    }

    #[test]
    fn archive_readers_reject_lossy_names_colliding_with_valid_replacement_names() {
        // `from_utf8_lossy` maps each malformed sequence to U+FFFD.  A valid
        // ZIP name may contain that same scalar value, so comparing only the
        // raw bytes would otherwise leave two members under one public lookup
        // key.  Exercise both central-directory orders because the index must
        // reject whichever spelling it encounters second.
        let valid_replacement = b"collision\xef\xbf\xbd.bin";
        for entries in [
            [
                FixtureEntry::stored(b"collision\xff.bin", b"invalid"),
                FixtureEntry::stored(valid_replacement, b"valid"),
            ],
            [
                FixtureEntry::stored(valid_replacement, b"valid"),
                FixtureEntry::stored(b"collision\xff.bin", b"invalid"),
            ],
        ] {
            let bytes = fixture(&entries);
            for result in [
                ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).map(|_| ()),
                indexed_archive_result(bytes.clone(), ArchiveLimits::UNBOUNDED).map(|_| ()),
            ] {
                assert!(matches!(
                    result,
                    Err(error)
                        if matches!(error.kind(), ErrorKind::InvalidInput { msg }
                            if msg.contains("duplicate normalized file names"))
                ));
            }
        }
    }

    #[test]
    fn archive_lookups_apply_the_same_normalization_as_ingress() {
        let bytes = fixture(&[FixtureEntry::stored(b"dir/../body.xml", b"body")]);
        let reader = ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).unwrap();
        assert!(reader.contains("/./dir/../body.xml"));
        assert_eq!(reader.read("/./dir/../body.xml").unwrap(), b"body");

        let indexed = indexed_archive_result(bytes, ArchiveLimits::UNBOUNDED).unwrap();
        assert!(indexed.contains("/./dir/../body.xml"));
        assert_eq!(indexed.read("/./dir/../body.xml").unwrap(), b"body");
    }

    #[test]
    fn archive_readers_normalize_lossy_utf8_names_for_lookup() {
        let bytes = fixture(&[FixtureEntry::stored(b"dir/../lossy\xff.bin", b"body")]);
        let query = "/./dir/../lossy\u{FFFD}.bin";

        let reader = ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).unwrap();
        assert_eq!(reader.file_names().collect::<Vec<_>>(), ["lossy�.bin"]);
        assert!(reader.contains(query));
        assert_eq!(reader.read(query).unwrap(), b"body");

        let indexed = indexed_archive_result(bytes.clone(), ArchiveLimits::UNBOUNDED).unwrap();
        assert_eq!(indexed.file_names().collect::<Vec<_>>(), ["lossy�.bin"]);
        assert!(indexed.contains(query));
        assert_eq!(indexed.read(query).unwrap(), b"body");

        let lazy = LazyArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).unwrap();
        assert!(lazy.contains(query));
        assert_eq!(lazy.read_shared(query).unwrap().as_slice(), b"body");
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

    #[test]
    fn indexed_archive_borrows_a_preservation_index_with_exact_raw_names() {
        let bytes = fixture(&[
            FixtureEntry::stored(b"before.bin", b"before"),
            FixtureEntry::stored(b"folder/", b""),
            FixtureEntry::stored("caf\u{e9}.txt".as_bytes(), b"utf8"),
            FixtureEntry::stored(b"\xffraw.bin", b"opaque"),
        ]);
        let indexed = indexed_archive(bytes);
        let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
        let preservation = indexed.preservation_index(&mut scratch).unwrap();
        let names: Vec<_> = preservation
            .entries()
            .iter()
            .map(|entry| entry.raw_name_bytes().to_vec())
            .collect();

        assert_eq!(
            names,
            vec![
                b"before.bin".to_vec(),
                b"folder/".to_vec(),
                "caf\u{e9}.txt".as_bytes().to_vec(),
                b"\xffraw.bin".to_vec(),
            ]
        );
        assert_eq!(preservation.entries().len(), 4);
    }

    #[test]
    fn indexed_archive_retains_the_located_archive_end_without_rescanning() {
        let mut bytes = fixture(&[FixtureEntry::stored(b"payload.bin", b"payload")]);
        let archive_end = bytes.len() as u64;
        bytes.extend_from_slice(b"opaque trailing bytes");
        let source_len = bytes.len() as u64;
        let indexed = IndexedArchive::from_reader(std::io::Cursor::new(bytes), source_len)
            .expect("ZIP before the opaque suffix remains locatable");

        assert_eq!(indexed.archive_end_offset(), archive_end);
        assert_ne!(indexed.archive_end_offset(), source_len);
    }

    #[test]
    fn preservation_ids_follow_central_records_not_office_entry_ordinals() {
        let mut bytes = fixture(&[
            FixtureEntry::stored(b"first.bin", b"first"),
            FixtureEntry::stored(b"folder/", b""),
            FixtureEntry::stored(b"second.bin", b"second"),
        ]);
        let archive = ZipArchive::from_slice(&bytes).unwrap();
        let central = archive.directory_offset() as usize;
        let eocd = archive.eocd_offset() as usize;
        let first_len = central_record_len(&bytes, central);
        let second_len = central_record_len(&bytes, central + first_len);
        let first = bytes[central..central + first_len].to_vec();
        let second = bytes[central + first_len..central + first_len + second_len].to_vec();
        let third = bytes[central + first_len + second_len..eocd].to_vec();
        bytes[central..eocd].copy_from_slice(&[third, second, first].concat());

        let indexed = indexed_archive(bytes);
        assert_eq!(
            indexed.file_names().collect::<Vec<_>>(),
            vec!["first.bin", "second.bin"]
        );
        let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
        let preservation = indexed.preservation_index(&mut scratch).unwrap();
        let names: Vec<_> = preservation
            .entries()
            .iter()
            .map(|entry| entry.raw_name_bytes().to_vec())
            .collect();

        assert_eq!(
            names,
            vec![
                b"second.bin".to_vec(),
                b"folder/".to_vec(),
                b"first.bin".to_vec(),
            ]
        );
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

    fn central_record_len(bytes: &[u8], offset: usize) -> usize {
        const FIXED_SIZE: usize = 46;
        let name_len = u16::from_le_bytes([bytes[offset + 28], bytes[offset + 29]]) as usize;
        let extra_len = u16::from_le_bytes([bytes[offset + 30], bytes[offset + 31]]) as usize;
        let comment_len = u16::from_le_bytes([bytes[offset + 32], bytes[offset + 33]]) as usize;
        FIXED_SIZE + name_len + extra_len + comment_len
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
