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

use crate::accounting::{
    AccountingReadKind, AccountingWriteKind, CountingReader, accounting_overflow, usize_to_u64,
};
use crate::crc::crc32_chunk;
use crate::path::{RawPath, ZipFilePath};
use crate::reader_at::validate_read_count;
use crate::writer::ReusedDeflateEncoder;
use crate::{
    CompressionMethod, Error, ErrorKind, PreservationIndex, RECOMMENDED_BUFFER_SIZE, ReaderAt,
    ZipArchive, ZipArchiveWriter, ZipLocator, ZipOperationAccounting, ZipReader, ZipSliceArchive,
    ZipVerification,
};
use flate2::read::DeflateDecoder;
use flate2::{Decompress, FlushDecompress, Status};
use rayon::prelude::*;
use std::collections::{HashMap, HashSet};
use std::io::{self, BufRead, Cursor, Read, Write};
use std::mem::size_of;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Arc, Condvar, Mutex};

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
    /// This includes the fixed 46-byte central record, raw member names, extra
    /// fields, and file comments for every entry, including directories. It is
    /// checked before name normalization or ownership allocation.
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
    /// Every central entry's wayfinder, sorted by local-header offset. This
    /// is used only to prove non-overlap before publishing a borrowed slice.
    layout: Vec<BorrowedLayoutEntry>,
    /// A successful strict local-layout proof. The slice source is immutable,
    /// so the proof remains valid for the reader lifetime.
    strict_layout_cache: StrictLayoutCache,
}

/// Reusable decoder state for one sequential operation over a slice-backed
/// ZIP archive.
///
/// The session is intentionally mutable: one Deflate decoder is reset between
/// members, while Store members bypass it. It is not a cache and does not
/// change archive verification policy. Create a fresh session for each
/// independent operation.
pub struct ArchiveReadSession<'archive, 'data> {
    archive: &'archive ArchiveReader<'data>,
    decoder: Option<DeflateDecoder<CountingReader<&'data [u8]>>>,
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
    flags: u16,
    compression_method: CompressionMethod,
    uncompressed_size: u64,
    central_name: Vec<u8>,
}

#[derive(Debug, Clone)]
struct BorrowedLayoutEntry {
    wayfinder: crate::ZipArchiveEntryWayfinder,
    central_name: Vec<u8>,
}

#[derive(Debug, Clone)]
enum IndexedLayoutName {
    Entry(EntryId),
    Directory(Vec<u8>),
}

#[derive(Debug, Clone)]
struct IndexedLayoutEntry {
    wayfinder: crate::ZipArchiveEntryWayfinder,
    name: IndexedLayoutName,
}

/// The bytes of one fixed ZIP local file header.
const STRICT_FIXED_LOCAL_HEADER_BYTES: u64 = 30;

/// The largest offset one central record's declared local span can reach,
/// using resident central metadata and no positional read.
///
/// See [`crate::MAX_LOCAL_SPAN_RESIDUAL`] for why both halves of the local
/// variable region are counted.
fn zero_io_max_span_end(entry: &crate::ZipArchiveEntryWayfinder) -> u64 {
    zero_io_min_span_end(entry).saturating_add(crate::MAX_LOCAL_SPAN_RESIDUAL)
}

/// The smallest offset one central record's declared local span can reach,
/// using resident central metadata and no positional read.
///
/// A local variable region and a data descriptor can only push a span end
/// further out, so a record whose payload alone already runs past an offset
/// overlaps it whatever its local header says.
fn zero_io_min_span_end(entry: &crate::ZipArchiveEntryWayfinder) -> u64 {
    entry
        .local_header_offset()
        .saturating_add(STRICT_FIXED_LOCAL_HEADER_BYTES)
        .saturating_add(entry.compressed_size_hint())
}

fn strict_overlap_error() -> Error {
    Error::from(ErrorKind::InvalidInput {
        msg: "strict streaming refuses overlapping ZIP local spans".to_string(),
    })
}

fn strict_layout_index_error() -> Error {
    Error::from(ErrorKind::InvalidInput {
        msg: "strict layout proof target index is invalid".to_string(),
    })
}

/// What one reader has already established about its own physical layout.
///
/// Successes only. A proof that fails leaves the memo exactly as it was, so a
/// read that failed for cancellation, a budget error or a transient source
/// error retries in full, and no error is ever cached.
///
/// The memo never decides a verdict. Every entry it holds is a fact about the
/// archive's bytes, so a target's accept-or-refuse outcome is the same whether
/// the memo is empty or full; the memo only removes repeated work.
#[derive(Debug, Default)]
struct StrictLayoutMemo {
    /// Fully validated target layouts, keyed by local-header offset.
    targets: HashMap<u64, crate::StrictEntryLayout>,
    /// Where each already-probed record's declared local span ends, keyed by
    /// local-header offset.
    bounds: HashMap<u64, crate::LocalSpanBound>,
    /// Prefix maximum of every record's zero-I/O span upper bound over the
    /// offset-sorted layout. A pure function of the central directory, built
    /// once, and used only to stop the predecessor scan early. Its presence
    /// also records that the distinct-offset check has already passed.
    prefix_max: Vec<u64>,
}

impl StrictLayoutMemo {
    fn is_empty(&self) -> bool {
        self.targets.is_empty() && self.bounds.is_empty() && self.prefix_max.is_empty()
    }

    fn merge(&mut self, proven: StrictLayoutProven) -> Result<(), Error> {
        if let Some(prefix_max) = proven.prefix_max {
            self.prefix_max = prefix_max;
        }
        let layout = proven.layout;
        self.targets.try_reserve(1).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "strict layout target memo",
                source,
            })
        })?;
        self.targets.insert(layout.local_header_offset, layout);
        self.bounds
            .try_reserve(proven.learned.len().saturating_add(1))
            .map_err(|source| {
                Error::from(ErrorKind::Allocation {
                    resource: "strict layout span memo",
                    source,
                })
            })?;
        for (offset, bound) in proven.learned {
            self.bounds.insert(offset, bound);
        }
        self.bounds.insert(
            layout.local_header_offset,
            crate::LocalSpanBound::Exact(layout.span_end),
        );
        Ok(())
    }
}

/// The facts one target-scoped proof established, ready to merge.
#[derive(Debug)]
struct StrictLayoutProven {
    layout: crate::StrictEntryLayout,
    learned: Vec<(u64, crate::LocalSpanBound)>,
    prefix_max: Option<Vec<u64>>,
}

/// Prove that no other central record's declared local span intersects the
/// target's.
///
/// `layout_len` and the three accessors describe one reader's physical layout,
/// sorted by local-header offset and including directory records. The proof
/// establishes, in this order:
///
/// 1. **Distinct local-header offsets**, over every record. A record's
///    local-header offset is copied verbatim from its central record, so this
///    is a pure central-directory property and costs no read.
/// 2. **The target's own layout**, in full and unchanged: method, flags, name,
///    local and central sizes, CRC, data span, data-descriptor resolution and
///    the per-entry refusal of a span that reaches into the central directory.
/// 3. **Successors**, at no cost. A record that starts after the target cannot
///    reach backwards, so the only question is whether the target's own span
///    runs into it, and the answer is already resident metadata.
/// 4. **Predecessors**, pruned by a zero-I/O central-directory bracket, and
///    otherwise settled by one 30-byte read of that record's fixed local
///    header.
///
/// The verdict is a function of the archive bytes and the target alone. It
/// does not depend on what this reader read earlier, which is what an accepted
/// ADR 0005 requires of anything a cache participates in.
fn prove_target_scoped_strict_layout<Wayfinder, Validate, Bound, Resolve>(
    layout_len: usize,
    wayfinder_at: Wayfinder,
    validate_at: Validate,
    bound_at: Bound,
    resolve_at: Resolve,
    target: crate::ZipArchiveEntryWayfinder,
    memo: &StrictLayoutMemo,
) -> Result<StrictLayoutProven, Error>
where
    Wayfinder: Fn(usize) -> Option<crate::ZipArchiveEntryWayfinder>,
    Validate: Fn(usize) -> Result<crate::StrictEntryLayout, Error>,
    Bound: Fn(usize) -> Result<crate::LocalSpanBound, Error>,
    Resolve: Fn(usize, crate::LocalSpanBound) -> Result<u64, Error>,
{
    let target_offset = target.local_header_offset();

    // (1) Distinct local-header offsets, and the prefix maximum that bounds
    // the predecessor scan.  Both are pure central-directory facts, so one
    // pass establishes them for every later read of this reader.
    let built_prefix_max = if memo.prefix_max.len() == layout_len {
        None
    } else {
        let mut prefix_max = Vec::new();
        prefix_max.try_reserve_exact(layout_len).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "strict layout span bracket",
                source,
            })
        })?;
        let mut running = 0_u64;
        let mut previous_offset = None;
        for position in 0..layout_len {
            let entry = wayfinder_at(position).ok_or_else(strict_layout_index_error)?;
            let offset = entry.local_header_offset();
            if previous_offset == Some(offset) {
                return Err(Error::from(ErrorKind::InvalidInput {
                    msg: "strict streaming refuses duplicate ZIP local spans".to_string(),
                }));
            }
            previous_offset = Some(offset);
            running = running.max(zero_io_max_span_end(&entry));
            prefix_max.push(running);
        }
        Some(prefix_max)
    };
    let prefix_max: &[u64] = built_prefix_max
        .as_deref()
        .unwrap_or(memo.prefix_max.as_slice());

    // Offsets are sorted and now known distinct, so the target has exactly one
    // position.
    let target_position = {
        let mut low = 0_usize;
        let mut high = layout_len;
        let mut found = None;
        while low < high {
            let middle = low + (high - low) / 2;
            let offset = wayfinder_at(middle)
                .ok_or_else(strict_layout_index_error)?
                .local_header_offset();
            match offset.cmp(&target_offset) {
                std::cmp::Ordering::Less => low = middle + 1,
                std::cmp::Ordering::Greater => high = middle,
                std::cmp::Ordering::Equal => {
                    found = Some(middle);
                    break;
                },
            }
        }
        found.ok_or_else(|| {
            Error::from(ErrorKind::InvalidInput {
                msg: "strict layout proof has no target local header".to_string(),
            })
        })?
    };

    // (2) The target's own layout, unchanged.
    let layout = validate_at(target_position)?;
    if layout.local_header_offset != target_offset {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "strict layout proof target offset does not match central metadata".to_string(),
        }));
    }

    // (3) Successors, at zero I/O.  The layout is sorted and the offsets are
    // distinct, so the immediately following record has the smallest offset of
    // any later record; if it is not already past the target's exact span end,
    // the two spans intersect.
    if let Some(next) = wayfinder_at(target_position.saturating_add(1)) {
        if next.local_header_offset() < layout.span_end {
            return Err(strict_overlap_error());
        }
    }

    // (4) Predecessors.
    let mut learned: Vec<(u64, crate::LocalSpanBound)> = Vec::new();
    for position in (0..target_position).rev() {
        if prefix_max
            .get(position)
            .is_some_and(|reach| *reach <= target_offset)
        {
            // No record at or before this position can reach the target.
            break;
        }
        let entry = wayfinder_at(position).ok_or_else(strict_layout_index_error)?;
        let offset = entry.local_header_offset();
        if zero_io_max_span_end(&entry) <= target_offset {
            // Proven disjoint from resident metadata alone.
            continue;
        }
        if zero_io_min_span_end(&entry) > target_offset {
            // Refuted from resident metadata alone: this record's payload
            // already runs past the target's local header.
            return Err(strict_overlap_error());
        }
        let bound = match memo.bounds.get(&offset).copied() {
            Some(bound) => bound,
            None => {
                let bound = bound_at(position)?;
                learned.try_reserve(1).map_err(|source| {
                    Error::from(ErrorKind::Allocation {
                        resource: "strict layout neighbour spans",
                        source,
                    })
                })?;
                learned.push((offset, bound));
                bound
            },
        };
        if bound.min_span_end() > target_offset {
            return Err(strict_overlap_error());
        }
        if bound.max_span_end() > target_offset {
            // The record declares a data descriptor whose encoded width is the
            // only thing left deciding the verdict.  Settle it exactly: a
            // bound that simply reserved the widest descriptor would refuse
            // every gapless descriptor-bearing archive, where a record's
            // payload ends exactly one descriptor before the next record.
            let span_end = resolve_at(position, bound)?;
            learned.try_reserve(1).map_err(|source| {
                Error::from(ErrorKind::Allocation {
                    resource: "strict layout neighbour spans",
                    source,
                })
            })?;
            learned.push((offset, crate::LocalSpanBound::Exact(span_end)));
            if span_end > target_offset {
                return Err(strict_overlap_error());
            }
        }
    }

    Ok(StrictLayoutProven {
        layout,
        learned,
        prefix_max: built_prefix_max,
    })
}

#[derive(Debug)]
enum StrictLayoutCacheState {
    Empty,
    Building { owner: std::thread::ThreadId },
    Ready(StrictLayoutMemo),
}

struct StrictLayoutCache {
    state: Mutex<StrictLayoutCacheState>,
    wake: Condvar,
    #[cfg(test)]
    build_count: std::sync::atomic::AtomicUsize,
}

impl StrictLayoutCache {
    fn new() -> Self {
        Self {
            state: Mutex::new(StrictLayoutCacheState::Empty),
            wake: Condvar::new(),
            #[cfg(test)]
            build_count: std::sync::atomic::AtomicUsize::new(0),
        }
    }

    #[cfg(test)]
    fn is_ready(&self) -> bool {
        matches!(
            &*lock_strict_layout_state(self),
            StrictLayoutCacheState::Ready(_)
        )
    }

    #[cfg(test)]
    fn build_count(&self) -> usize {
        self.build_count.load(Ordering::Acquire)
    }
}

fn lock_strict_layout_state(
    cache: &StrictLayoutCache,
) -> std::sync::MutexGuard<'_, StrictLayoutCacheState> {
    match cache.state.lock() {
        Ok(guard) => guard,
        Err(poisoned) => poisoned.into_inner(),
    }
}

struct StrictLayoutBuildGuard<'a> {
    cache: &'a StrictLayoutCache,
    owner: std::thread::ThreadId,
    active: bool,
}

impl<'a> StrictLayoutBuildGuard<'a> {
    fn new(cache: &'a StrictLayoutCache, owner: std::thread::ThreadId) -> Self {
        Self {
            cache,
            owner,
            active: true,
        }
    }
}

impl Drop for StrictLayoutBuildGuard<'_> {
    fn drop(&mut self) {
        if !self.active {
            return;
        }
        let mut state = lock_strict_layout_state(self.cache);
        if matches!(
            &*state,
            StrictLayoutCacheState::Building { owner } if *owner == self.owner
        ) {
            *state = StrictLayoutCacheState::Empty;
            self.cache.wake.notify_all();
        }
    }
}

/// Prove one target's strict layout, reusing whatever this reader already
/// established.
///
/// The memo is taken out of the cache for the duration of one proof, so the
/// existing single-flight contract is unchanged: a second thread waits, and a
/// source that re-enters its own archive on the proving thread still reports a
/// re-entrancy error instead of deadlocking. Only successes are merged back.
fn strict_layout_for_cached<Prove>(
    cache: &StrictLayoutCache,
    target: crate::ZipArchiveEntryWayfinder,
    prove: Prove,
) -> Result<crate::StrictEntryLayout, Error>
where
    Prove: FnOnce(&StrictLayoutMemo) -> Result<StrictLayoutProven, Error>,
{
    let owner = std::thread::current().id();
    let target_offset = target.local_header_offset();
    let mut memo = loop {
        let mut state = lock_strict_layout_state(cache);
        match &mut *state {
            StrictLayoutCacheState::Ready(memo) => {
                if let Some(layout) = memo.targets.get(&target_offset).copied() {
                    return Ok(layout);
                }
                let taken = std::mem::take(memo);
                *state = StrictLayoutCacheState::Building { owner };
                #[cfg(test)]
                cache.build_count.fetch_add(1, Ordering::AcqRel);
                drop(state);
                break taken;
            },
            StrictLayoutCacheState::Empty => {
                *state = StrictLayoutCacheState::Building { owner };
                #[cfg(test)]
                cache.build_count.fetch_add(1, Ordering::AcqRel);
                drop(state);
                break StrictLayoutMemo::default();
            },
            StrictLayoutCacheState::Building {
                owner: building_owner,
            } if *building_owner == owner => {
                return Err(Error::from(ErrorKind::InvalidInput {
                    msg: "strict layout proof build re-entered on its owning thread".to_string(),
                }));
            },
            StrictLayoutCacheState::Building { .. } => {
                state = match cache.wake.wait(state) {
                    Ok(guard) => guard,
                    Err(poisoned) => poisoned.into_inner(),
                };
                drop(state);
            },
        }
    };

    let mut build_guard = StrictLayoutBuildGuard::new(cache, owner);
    let outcome = match prove(&memo) {
        Ok(proven) => {
            let layout = proven.layout;
            // A merge failure is an allocation failure and keeps its identity;
            // every fact already merged stays true either way.
            let merged = memo.merge(proven);
            publish_strict_layout_memo(cache, memo);
            merged.map(|()| layout)
        },
        Err(error) => {
            publish_strict_layout_memo(cache, memo);
            Err(error)
        },
    };
    build_guard.active = false;
    outcome
}

/// Return the memo to the cache and wake anyone waiting on it.
///
/// A memo that learned nothing is published as `Empty`, so a reader whose
/// first read failed is indistinguishable from one that has never been read.
fn publish_strict_layout_memo(cache: &StrictLayoutCache, memo: StrictLayoutMemo) {
    let mut state = lock_strict_layout_state(cache);
    *state = if memo.is_empty() {
        StrictLayoutCacheState::Empty
    } else {
        StrictLayoutCacheState::Ready(memo)
    };
    cache.wake.notify_all();
}

/// Opaque identifier for one non-directory member in an [`IndexedArchive`].
///
/// An ID is stable for the lifetime of the archive that produced it. Its
/// representation is intentionally private so callers cannot manufacture an
/// unchecked physical entry reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct EntryId(usize);

/// A source-issued, fully verified compressed member payload.
///
/// The fields are intentionally private.  Callers can only obtain this value
/// from [`IndexedArchive::read_entry_precompressed_with_progress`] or
/// [`IndexedArchive::read_entry_precompressed_and_decoded_with_progress`], which
/// validate the source layout, capture the exact bounded compressed range,
/// decode that immutable capture, compare or return its logical bytes, and
/// record the actual decoded CRC.  This prevents an
/// unchecked `(compressed, crc, size)` tuple from crossing into a writer.
#[derive(Debug, Clone)]
pub struct VerifiedPrecompressedEntry {
    method: CompressionMethod,
    compressed: Arc<Vec<u8>>,
    compressed_size: u64,
    uncompressed_size: u64,
    crc32: u32,
}

impl VerifiedPrecompressedEntry {
    /// Returns the verified source compression method.
    #[must_use]
    pub const fn compression_method(&self) -> CompressionMethod {
        self.method
    }

    /// Returns the exact captured compressed payload size.
    #[must_use]
    pub const fn compressed_size(&self) -> u64 {
        self.compressed_size
    }

    /// Returns the verified decoded payload size.
    #[must_use]
    pub const fn uncompressed_size(&self) -> u64 {
        self.uncompressed_size
    }

    /// Returns the actual CRC computed from the captured decoded payload.
    #[must_use]
    pub const fn crc32(&self) -> u32 {
        self.crc32
    }

    pub(crate) fn compressed_payload(&self) -> &[u8] {
        self.compressed.as_slice()
    }
}

/// Progress reported while issuing a verified precompressed entry.
///
/// The callback receives byte counts rather than unverified payload bytes. It
/// is called after each bounded compressed capture chunk and each bounded
/// decoded verification chunk.
/// Counts are cumulative within each phase. A decoded count may repeat when
/// a bounded decoder step consumes input without producing output.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PrecompressedProgress {
    /// Bytes captured from the bounded source compressed range.
    Compressed { bytes: u64 },
    /// Bytes decoded and checked against the declared bound (and expected bytes,
    /// when supplied). Final size and CRC validation precede token publication.
    Decoded { bytes: u64 },
}

/// Typed failure while issuing a verified precompressed entry.
///
/// Unlike [`VerifiedEntryReaderError`], a progress callback failure aborts
/// immediately. The operation does not drain an unbounded or malformed member
/// after cancellation merely to preserve an unrelated secondary error.
#[derive(Debug)]
#[non_exhaustive]
pub enum VerifiedPrecompressedError<E> {
    /// The source entry failed ZIP layout, size, checksum, or content checks.
    Archive(Error),
    /// The positional source returned an I/O failure while capturing or
    /// validating the compressed payload.
    Transport(io::Error),
    /// The caller's progress or execution callback requested cancellation.
    Callback(E),
}

impl<E> VerifiedPrecompressedError<E> {
    /// Returns the archive failure, when present.
    #[must_use]
    pub fn archive(&self) -> Option<&Error> {
        match self {
            Self::Archive(error) => Some(error),
            Self::Transport(_) | Self::Callback(_) => None,
        }
    }

    /// Returns the source transport failure, when present.
    #[must_use]
    pub fn transport(&self) -> Option<&io::Error> {
        match self {
            Self::Transport(error) => Some(error),
            Self::Archive(_) | Self::Callback(_) => None,
        }
    }

    /// Returns the callback failure, when present.
    #[must_use]
    pub fn callback(&self) -> Option<&E> {
        match self {
            Self::Callback(error) => Some(error),
            Self::Archive(_) | Self::Transport(_) => None,
        }
    }
}

impl<E: std::fmt::Display> std::fmt::Display for VerifiedPrecompressedError<E> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Archive(error) => write!(formatter, "verified ZIP archive failed: {error}"),
            Self::Transport(error) => write!(formatter, "verified ZIP transport failed: {error}"),
            Self::Callback(error) => {
                write!(formatter, "verified ZIP progress callback failed: {error}")
            },
        }
    }
}

impl<E: std::error::Error + 'static> std::error::Error for VerifiedPrecompressedError<E> {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Archive(error) => Some(error),
            Self::Transport(error) => Some(error),
            Self::Callback(error) => Some(error),
        }
    }
}

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
/// Strict sink reads cache their successful physical-layout proof; positional
/// `ReaderAt` sources must remain byte-stable throughout construction and the
/// archive's lifetime. `ReaderAt` callbacks must not re-enter this archive.
pub struct IndexedArchive<R> {
    archive: ZipArchive<R>,
    layout: Vec<IndexedLayoutEntry>,
    index: HashMap<String, EntryId>,
    entries: Vec<IndexedEntry>,
    directories: HashMap<String, Metadata>,
    order: Vec<EntryId>,
    has_encrypted_entries: bool,
    /// Whether any central-directory record declares a data descriptor.
    ///
    /// This is captured while the index is built so opaque package owners can
    /// make a metadata-only catalog decision without iterating the central
    /// directory again.
    has_data_descriptor_entries: bool,
    /// Whether the located archive or any central-directory record uses ZIP64
    /// metadata.
    ///
    /// The archive-level bit is available directly from the retained ZIP
    /// framing, separately from per-entry ZIP64 fields.
    has_zip64_metadata: bool,
    /// Whether every central record's declared local span ends before the
    /// located central directory.  This includes directory records and is
    /// intentionally computed from the same central pass as the index.
    all_local_spans_bounded: bool,
    /// A successful strict local-layout proof. The `ReaderAt` source must
    /// remain byte-stable during construction and for this archive's lifetime.
    strict_layout_cache: StrictLayoutCache,
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

/// Reusable decoder state for one sequential operation over an indexed ZIP
/// archive.
///
/// The session borrows its archive and is intentionally mutable: a single
/// Deflate decoder is reset between members, while Store members bypass it.
/// It is not a cache and does not change archive ownership or verification
/// policy. Create a fresh session for each independent operation.
pub struct IndexedReadSession<'a, R>
where
    R: ReaderAt,
{
    archive: &'a IndexedArchive<R>,
    decoder: Option<DeflateDecoder<CountingReader<ZipReader<&'a R>>>>,
}

impl<'a, R> IndexedReadSession<'a, R>
where
    R: ReaderAt,
{
    /// Create a sequential read session borrowing an indexed archive.
    #[must_use]
    pub fn new(archive: &'a IndexedArchive<R>) -> Self {
        Self {
            archive,
            decoder: None,
        }
    }

    /// Read and verify one member by name, reusing this session's decoder.
    ///
    /// Name admission is the archive's own: the same normalization, the same
    /// explicit-directory rejection, and the same
    /// [`FileNotFound`](ErrorKind::FileNotFound) identity as
    /// [`IndexedArchive::read`].
    pub fn read(&mut self, name: &str) -> Result<Vec<u8>, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_with_accounting(name, &mut accounting)
    }

    /// Read and verify one member by name while recording actual payload work.
    pub fn read_with_accounting(
        &mut self,
        name: &str,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<Vec<u8>, Error> {
        let entry_id = self.archive.entry_id_for_name(name)?;
        self.read_entry_with_accounting(entry_id, accounting)
    }

    /// Read and verify one member by its stable opaque entry ID.
    ///
    /// Deflate decoder state is reused only within this session. Store
    /// members continue to use their independent positional reader and never
    /// initialize the decoder.
    pub fn read_entry(&mut self, entry_id: EntryId) -> Result<Vec<u8>, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_entry_with_accounting(entry_id, &mut accounting)
    }

    /// Read and verify one indexed member while recording actual payload work.
    pub fn read_entry_with_accounting(
        &mut self,
        entry_id: EntryId,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<Vec<u8>, Error> {
        let indexed = self.archive.indexed_entry(entry_id)?;
        let entry = self.archive.archive.get_entry(indexed.info.wayfinder)?;
        let size = usize::try_from(indexed.info.uncompressed_size).map_err(|_| {
            Error::from(ErrorKind::InvalidInput {
                msg: format!(
                    "archive entry size {} does not fit this platform",
                    indexed.info.uncompressed_size
                ),
            })
        })?;
        let read_limit = indexed
            .info
            .uncompressed_size
            .checked_add(1)
            .ok_or_else(|| {
                Error::from(ErrorKind::InvalidInput {
                    msg: format!(
                        "archive entry size {} cannot be bounded with an overrun sentinel",
                        indexed.info.uncompressed_size
                    ),
                })
            })?;
        let read_capacity = usize::try_from(read_limit).map_err(|_| {
            Error::from(ErrorKind::InvalidInput {
                msg: format!(
                    "archive entry size {} plus an overrun sentinel does not fit this platform",
                    indexed.info.uncompressed_size
                ),
            })
        })?;
        let mut output = Vec::new();
        output.try_reserve_exact(read_capacity).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "indexed archive entry output",
                source,
            })
        })?;

        match indexed.info.compression_method {
            CompressionMethod::Store => {
                let mut source = CountingReader::new(entry.reader());
                let result = entry
                    .verifying_reader(&mut source)
                    .take(read_limit)
                    .read_to_end(&mut output)
                    .map_err(Error::from);
                let accounting_result = accounting
                    .add_stored_payload_bytes_read(source.count())
                    .and_then(|()| {
                        accounting.add_stored_payload_bytes_accepted(usize_to_u64(
                            output.len(),
                            "stored payload bytes accepted",
                        )?)
                    });
                if let Err(error) = result {
                    drop(accounting_result);
                    return Err(error);
                }
                accounting_result?;
            },
            CompressionMethod::Deflate => {
                let (result, compressed_read, produced) = {
                    let decoder = match self.decoder.as_mut() {
                        Some(decoder) => {
                            let _previous = decoder.reset(CountingReader::new(entry.reader()));
                            decoder
                        },
                        None => self
                            .decoder
                            .insert(DeflateDecoder::new(CountingReader::new(entry.reader()))),
                    };
                    let (result, produced_count) = {
                        let mut produced_reader = CountingReader::new(&mut *decoder);
                        let result = entry
                            .verifying_reader(&mut produced_reader)
                            .take(read_limit)
                            .read_to_end(&mut output)
                            .map_err(Error::from);
                        let produced_count = produced_reader.count();
                        (result, produced_count)
                    };
                    let compressed_read = decoder.get_ref().count();
                    (result, compressed_read, produced_count)
                };
                let accounting_result = accounting
                    .add_compressed_deflate_payload_bytes_read(compressed_read)
                    .and_then(|()| accounting.add_deflate_bytes_produced(produced))
                    .and_then(|()| {
                        accounting.add_deflate_bytes_accepted(usize_to_u64(
                            output.len(),
                            "decompressed Deflate bytes accepted",
                        )?)
                    });
                if let Err(error) = result {
                    drop(accounting_result);
                    return Err(error);
                }
                accounting_result?;
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
                actual: usize_to_u64(output.len(), "decompressed ZIP bytes")?,
            }
            .into());
        }
        Ok(output)
    }
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

/// Fixed scratch capacity used by [`IndexedArchive::with_verified_entry_reader`].
///
/// The callback reader never allocates a payload-sized buffer.  A callback may
/// retain only the bytes returned for the current `fill_buf` call, and the
/// lifetime of those bytes is scoped to the callback invocation.
pub const VERIFIED_ENTRY_READER_BUFFER_SIZE: usize = 16 * 1024;

// flate2 1.1.10's read decoder allocates a 32 KiB compressed-input
// BufReader.  The locked zlib-rs 0.6.7 inflate allocator lays out one
// 32 KiB window plus its private state and 64-byte alignment padding.  A
// 64 KiB state envelope is intentionally conservative across supported target
// layouts; it is a bound, not the measured size of one particular machine.
const FLATE2_DEFLATE_INPUT_BUFFER_SIZE: usize = 32 * 1024;
const ZLIB_RS_INFLATE_STATE_UPPER_BOUND_BYTES: usize = 64 * 1024;

// Precompressed transfer captures are source reads rather than callback
// decoder buffers. Keep their requests bounded to one 64 KiB range so a
// short-read ReaderAt does not turn a large member into thousands of 16 KiB
// positional requests.
const PRECOMPRESSED_CAPTURE_BUFFER_SIZE: usize = 64 * 1024;

/// Failure from a callback-scoped, verified indexed-entry read.
///
/// Archive and transport failures are primary: when the callback has already
/// returned an error, that error is retained in the `callback_error` field of those
/// variants.  A callback failure is returned directly only after the complete
/// entry has been drained and verified successfully.
#[derive(Debug)]
#[non_exhaustive]
pub enum VerifiedEntryReaderError<E> {
    /// The entry failed ZIP layout, size, checksum, or other archive checks.
    Archive {
        /// Primary archive failure.
        error: Error,
        /// Callback failure observed before archive finalization failed.
        callback_error: Option<E>,
    },
    /// The positional source or decompressor returned an I/O failure.
    Transport {
        /// Primary transport failure.
        error: io::Error,
        /// Callback failure observed before transport finalization failed.
        callback_error: Option<E>,
    },
    /// The callback failed after archive verification completed successfully.
    Callback(E),
}

impl<E> VerifiedEntryReaderError<E> {
    /// Returns the primary archive failure, when this is an archive failure.
    #[must_use]
    pub fn archive(&self) -> Option<&Error> {
        match self {
            Self::Archive { error, .. } => Some(error),
            Self::Transport { .. } | Self::Callback(_) => None,
        }
    }

    /// Returns the primary transport failure, when this is a transport failure.
    #[must_use]
    pub fn transport(&self) -> Option<&io::Error> {
        match self {
            Self::Transport { error, .. } => Some(error),
            Self::Archive { .. } | Self::Callback(_) => None,
        }
    }

    /// Returns the callback failure, including a secondary failure retained by
    /// an archive or transport primary error.
    #[must_use]
    pub fn callback(&self) -> Option<&E> {
        match self {
            Self::Archive { callback_error, .. } | Self::Transport { callback_error, .. } => {
                callback_error.as_ref()
            },
            Self::Callback(error) => Some(error),
        }
    }

    /// Alias for [`Self::archive`].
    #[must_use]
    pub fn archive_error(&self) -> Option<&Error> {
        self.archive()
    }

    /// Alias for [`Self::transport`].
    #[must_use]
    pub fn transport_error(&self) -> Option<&io::Error> {
        self.transport()
    }

    /// Alias for [`Self::callback`].
    #[must_use]
    pub fn callback_error(&self) -> Option<&E> {
        self.callback()
    }

    /// Extracts the callback error, if one was retained.
    pub fn into_callback(self) -> Option<E> {
        match self {
            Self::Archive { callback_error, .. } | Self::Transport { callback_error, .. } => {
                callback_error
            },
            Self::Callback(error) => Some(error),
        }
    }

    /// Alias for [`Self::into_callback`].
    pub fn into_callback_error(self) -> Option<E> {
        self.into_callback()
    }
}

impl<E: std::fmt::Display> std::fmt::Display for VerifiedEntryReaderError<E> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Archive {
                error,
                callback_error,
            } => match callback_error {
                Some(callback) => write!(
                    formatter,
                    "verified ZIP archive failed: {error} (callback also failed: {callback})"
                ),
                None => write!(formatter, "verified ZIP archive failed: {error}"),
            },
            Self::Transport {
                error,
                callback_error,
            } => match callback_error {
                Some(callback) => write!(
                    formatter,
                    "verified ZIP transport failed: {error} (callback also failed: {callback})"
                ),
                None => write!(formatter, "verified ZIP transport failed: {error}"),
            },
            Self::Callback(error) => write!(formatter, "verified ZIP callback failed: {error}"),
        }
    }
}

impl<E: std::error::Error + 'static> std::error::Error for VerifiedEntryReaderError<E> {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Archive { error, .. } => Some(error),
            Self::Transport { error, .. } => Some(error),
            Self::Callback(error) => Some(error),
        }
    }
}

#[derive(Debug)]
enum VerifiedReaderFailure {
    Archive(Error),
    Transport(io::Error),
    Accounting(Error),
}

impl VerifiedReaderFailure {
    fn as_io_error(&self) -> io::Error {
        match self {
            Self::Archive(error) | Self::Accounting(error) => {
                io::Error::new(io::ErrorKind::InvalidData, error.to_string())
            },
            Self::Transport(error) => io::Error::new(error.kind(), error.to_string()),
        }
    }

    fn into_public<E>(self, callback: Option<E>) -> VerifiedEntryReaderError<E> {
        match self {
            Self::Archive(error) => VerifiedEntryReaderError::Archive {
                error,
                callback_error: callback,
            },
            Self::Transport(error) => VerifiedEntryReaderError::Transport {
                error,
                callback_error: callback,
            },
            Self::Accounting(error) => VerifiedEntryReaderError::Archive {
                error,
                callback_error: callback,
            },
        }
    }
}

/// A fixed-buffer reader that defers archive failures until the caller's
/// callback has returned.  The callback can therefore inspect a valid prefix,
/// while the owner still drains and verifies the complete member afterward.
struct VerifiedEntryBufReader<'accounting, D> {
    reader: D,
    expected: ZipVerification,
    accounting: Option<&'accounting mut ZipOperationAccounting>,
    accounting_kind: AccountingReadKind,
    buffer: [u8; VERIFIED_ENTRY_READER_BUFFER_SIZE],
    start: usize,
    end: usize,
    produced: u64,
    crc: u32,
    eof: bool,
    problem: Option<VerifiedReaderFailure>,
}

impl<'accounting, D> VerifiedEntryBufReader<'accounting, D> {
    fn new(
        reader: D,
        expected: ZipVerification,
        accounting: Option<&'accounting mut ZipOperationAccounting>,
        accounting_kind: AccountingReadKind,
    ) -> Self {
        Self {
            reader,
            expected,
            accounting,
            accounting_kind,
            buffer: [0; VERIFIED_ENTRY_READER_BUFFER_SIZE],
            start: 0,
            end: 0,
            produced: 0,
            crc: 0,
            eof: false,
            problem: None,
        }
    }

    fn set_problem(&mut self, problem: VerifiedReaderFailure) {
        if self.problem.is_none() {
            self.problem = Some(problem);
        }
    }

    fn problem_io(&self) -> io::Error {
        self.problem.as_ref().map_or_else(
            || {
                io::Error::new(
                    io::ErrorKind::Other,
                    "verified ZIP reader failed without a diagnostic",
                )
            },
            VerifiedReaderFailure::as_io_error,
        )
    }

    fn record_produced(&mut self, read: usize) -> Result<(), ()> {
        if !matches!(self.accounting_kind, AccountingReadKind::Deflate) {
            return Ok(());
        }
        let bytes = match usize_to_u64(read, "decompressed Deflate bytes produced") {
            Ok(bytes) => bytes,
            Err(error) => {
                self.set_problem(VerifiedReaderFailure::Accounting(error));
                return Err(());
            },
        };
        let result = match self.accounting.as_mut() {
            Some(accounting) => accounting.add_deflate_bytes_produced(bytes),
            None => Ok(()),
        };
        if let Err(error) = result {
            self.set_problem(VerifiedReaderFailure::Accounting(error));
            return Err(());
        }
        Ok(())
    }

    fn record_accepted(&mut self, accepted: usize) -> Result<(), ()> {
        if accepted == 0 {
            return Ok(());
        }
        let bytes = match usize_to_u64(accepted, "decompressed ZIP bytes accepted") {
            Ok(bytes) => bytes,
            Err(error) => {
                self.set_problem(VerifiedReaderFailure::Accounting(error));
                return Err(());
            },
        };
        let result = match self.accounting.as_mut() {
            Some(accounting) => match self.accounting_kind {
                AccountingReadKind::Stored => accounting.add_stored_payload_bytes_accepted(bytes),
                AccountingReadKind::Deflate => accounting.add_deflate_bytes_accepted(bytes),
            },
            None => Ok(()),
        };
        if let Err(error) = result {
            self.set_problem(VerifiedReaderFailure::Accounting(error));
            return Err(());
        }
        Ok(())
    }

    fn fill_buf_internal(&mut self) -> io::Result<&[u8]>
    where
        D: Read,
    {
        if self.start < self.end {
            return Ok(&self.buffer[self.start..self.end]);
        }
        self.start = 0;
        self.end = 0;
        if self.problem.is_some() {
            return Err(self.problem_io());
        }
        if self.eof {
            return Ok(&[]);
        }

        let remaining = self.expected.size().saturating_sub(self.produced);
        let request = usize::try_from(remaining.saturating_add(1))
            .unwrap_or(VERIFIED_ENTRY_READER_BUFFER_SIZE)
            .min(VERIFIED_ENTRY_READER_BUFFER_SIZE);
        let read = match read_with_interrupt_budget(&mut self.reader, &mut self.buffer[..request]) {
            Ok(read) => read,
            Err(error) => {
                self.set_problem(VerifiedReaderFailure::Transport(error));
                return Err(self.problem_io());
            },
        };
        if read == 0 {
            self.eof = true;
            if self.produced != self.expected.size() {
                self.set_problem(VerifiedReaderFailure::Archive(
                    ErrorKind::InvalidSize {
                        expected: self.expected.size(),
                        actual: self.produced,
                    }
                    .into(),
                ));
                return Err(self.problem_io());
            }
            return Ok(&[]);
        }

        let read_u64 = match u64::try_from(read) {
            Ok(read_u64) => read_u64,
            Err(_) => {
                self.set_problem(VerifiedReaderFailure::Transport(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "verified ZIP reader output count overflows u64",
                )));
                return Err(self.problem_io());
            },
        };
        let produced = match self.produced.checked_add(read_u64) {
            Some(produced) => produced,
            None => {
                self.set_problem(VerifiedReaderFailure::Transport(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "verified ZIP reader output count overflows u64",
                )));
                return Err(self.problem_io());
            },
        };
        self.produced = produced;
        self.crc = crc32_chunk(&self.buffer[..read], self.crc);
        if self.record_produced(read).is_err() {
            return Err(self.problem_io());
        }
        if produced > self.expected.size() {
            self.set_problem(VerifiedReaderFailure::Archive(
                ErrorKind::InvalidSize {
                    expected: self.expected.size(),
                    actual: produced,
                }
                .into(),
            ));
        }
        let visible = usize::try_from(remaining).unwrap_or(read).min(read);
        self.end = visible;
        if visible == 0 && self.problem.is_some() {
            return Err(self.problem_io());
        }
        Ok(&self.buffer[..visible])
    }

    fn finish(&mut self) -> Result<(), VerifiedReaderFailure>
    where
        D: Read,
    {
        loop {
            if let Some(problem) = self.problem.take() {
                return Err(problem);
            }
            let empty = match self.fill_buf_internal() {
                Ok(buffer) => buffer.is_empty(),
                Err(_) => {
                    return Err(self.problem.take().unwrap_or_else(|| {
                        VerifiedReaderFailure::Transport(io::Error::new(
                            io::ErrorKind::Other,
                            "verified ZIP reader failed without a diagnostic",
                        ))
                    }));
                },
            };
            if empty {
                break;
            }
            self.start = self.end;
        }
        self.expected
            .valid_strict(ZipVerification {
                crc: self.crc,
                uncompressed_size: self.produced,
            })
            .map_err(VerifiedReaderFailure::Archive)
    }

    fn abort(&mut self) -> Result<(), VerifiedReaderFailure> {
        self.problem.take().map_or(Ok(()), Err)
    }
}

impl<D: Read> Read for VerifiedEntryBufReader<'_, D> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let amount = {
            let available = self.fill_buf_internal()?;
            let amount = available.len().min(output.len());
            output[..amount].copy_from_slice(&available[..amount]);
            amount
        };
        self.consume(amount);
        Ok(amount)
    }
}

impl<D: Read> BufRead for VerifiedEntryBufReader<'_, D> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        self.fill_buf_internal()
    }

    fn consume(&mut self, amount: usize) {
        let available = self.end.saturating_sub(self.start);
        if amount > available {
            self.set_problem(VerifiedReaderFailure::Archive(
                ErrorKind::InvalidInput {
                    msg: format!(
                        "verified ZIP reader consumed {amount} bytes with only {available} buffered"
                    ),
                }
                .into(),
            ));
            return;
        }
        if self.record_accepted(amount).is_err() {
            return;
        }
        self.start += amount;
    }
}

fn read_with_interrupt_budget<R: Read>(reader: &mut R, buffer: &mut [u8]) -> io::Result<usize> {
    const MAX_INTERRUPTED_RETRIES: usize = 8;
    let mut retries = 0;
    loop {
        match reader.read(buffer) {
            Ok(read) => return validate_read_count(read, buffer.len()),
            Err(error)
                if error.kind() == io::ErrorKind::Interrupted
                    && retries < MAX_INTERRUPTED_RETRIES =>
            {
                retries += 1;
            },
            Err(error) => return Err(error),
        }
    }
}

fn precompressed_archive_error<E>(error: Error) -> VerifiedPrecompressedError<E> {
    VerifiedPrecompressedError::Archive(error)
}

fn precompressed_decode_error<E>(error: io::Error) -> VerifiedPrecompressedError<E> {
    VerifiedPrecompressedError::Archive(
        ErrorKind::InvalidInput {
            msg: format!("captured compressed payload could not be decoded: {error}"),
        }
        .into(),
    )
}

enum CapturedDecodedOutput<'a> {
    Compare(&'a [u8]),
    Collect { bytes: Vec<u8>, length: usize },
}

impl CapturedDecodedOutput<'_> {
    fn len(&self) -> usize {
        match self {
            Self::Compare(bytes) => bytes.len(),
            Self::Collect { length, .. } => *length,
        }
    }

    fn reserve<E>(&mut self) -> Result<(), VerifiedPrecompressedError<E>> {
        if let Self::Collect { bytes, length } = self {
            bytes.try_reserve_exact(*length).map_err(|source| {
                precompressed_archive_error(
                    ErrorKind::Allocation {
                        resource: "verified precompressed decoded payload",
                        source,
                    }
                    .into(),
                )
            })?;
        }
        Ok(())
    }

    fn accept<E>(
        &mut self,
        chunk: &[u8],
        start: usize,
        end: usize,
    ) -> Result<(), VerifiedPrecompressedError<E>> {
        if end > self.len() {
            return Err(precompressed_archive_error(
                ErrorKind::InvalidSize {
                    expected: u64::try_from(self.len()).unwrap_or(u64::MAX),
                    actual: u64::try_from(end).unwrap_or(u64::MAX),
                }
                .into(),
            ));
        }
        match self {
            Self::Compare(expected) => {
                if expected.get(start..end) != Some(chunk) {
                    return Err(precompressed_archive_error(ErrorKind::InvalidInput {
                        msg: "captured compressed payload decodes differently from expected logical bytes".to_string(),
                    }.into()));
                }
            },
            Self::Collect { bytes, .. } => {
                // The full declared length was reserved after layout validation.
                // Checked bounds prevent allocation or oversized output here.
                bytes.extend_from_slice(chunk);
            },
        }
        Ok(())
    }
}

fn verify_captured_precompressed_payload<E, F>(
    method: CompressionMethod,
    compressed: &[u8],
    expected_decoded: &mut CapturedDecodedOutput<'_>,
    expected: ZipVerification,
    progress: &mut F,
) -> Result<u32, VerifiedPrecompressedError<E>>
where
    F: FnMut(PrecompressedProgress) -> Result<(), E>,
{
    match method {
        CompressionMethod::Store => {
            let mut reader = Cursor::new(compressed);
            let crc =
                verify_captured_decoded_reader(&mut reader, expected_decoded, expected, progress)?;
            if reader.position() != u64::try_from(compressed.len()).unwrap_or(u64::MAX) {
                return Err(precompressed_archive_error(
                    ErrorKind::InvalidSize {
                        expected: u64::try_from(compressed.len())
                            .expect("compressed length fits u64"),
                        actual: reader.position(),
                    }
                    .into(),
                ));
            }
            Ok(crc)
        },
        CompressionMethod::Deflate => {
            verify_captured_deflate_payload(compressed, expected_decoded, expected, progress)
        },
        other => Err(precompressed_archive_error(
            ErrorKind::UnsupportedCompressionMethod(other.as_id().as_u16()).into(),
        )),
    }
}

fn verify_captured_deflate_payload<E, F>(
    compressed: &[u8],
    expected_decoded: &mut CapturedDecodedOutput<'_>,
    expected: ZipVerification,
    progress: &mut F,
) -> Result<u32, VerifiedPrecompressedError<E>>
where
    F: FnMut(PrecompressedProgress) -> Result<(), E>,
{
    const DEFLATE_INPUT_CHUNK_SIZE: usize = 64 * 1024;
    let mut decoder = Decompress::new(false);
    let mut input_offset = 0usize;
    let mut compared = 0usize;
    let mut actual_crc = 0_u32;
    let mut output = [0_u8; VERIFIED_ENTRY_READER_BUFFER_SIZE];

    loop {
        let input_end = input_offset
            .checked_add(DEFLATE_INPUT_CHUNK_SIZE)
            .unwrap_or(compressed.len())
            .min(compressed.len());
        let input_is_final = input_end == compressed.len();
        let before_in = decoder.total_in();
        let before_out = decoder.total_out();
        let status = decoder
            .decompress(
                &compressed[input_offset..input_end],
                &mut output,
                if input_is_final {
                    FlushDecompress::Finish
                } else {
                    FlushDecompress::None
                },
            )
            .map_err(|error| {
                precompressed_decode_error(io::Error::new(
                    io::ErrorKind::InvalidData,
                    error.to_string(),
                ))
            })?;
        let consumed = decoder
            .total_in()
            .checked_sub(before_in)
            .and_then(|count| usize::try_from(count).ok())
            .ok_or_else(|| {
                precompressed_archive_error(
                    ErrorKind::InvalidInput {
                        msg: "captured Deflate input count overflows usize".to_string(),
                    }
                    .into(),
                )
            })?;
        let produced = decoder
            .total_out()
            .checked_sub(before_out)
            .and_then(|count| usize::try_from(count).ok())
            .ok_or_else(|| {
                precompressed_archive_error(
                    ErrorKind::InvalidInput {
                        msg: "captured Deflate output count overflows usize".to_string(),
                    }
                    .into(),
                )
            })?;
        let remaining = compressed.len().checked_sub(input_offset).ok_or_else(|| {
            precompressed_archive_error(
                ErrorKind::InvalidInput {
                    msg: "captured Deflate input position exceeds payload length".to_string(),
                }
                .into(),
            )
        })?;
        if consumed > remaining || produced > output.len() {
            return Err(precompressed_archive_error(
                ErrorKind::InvalidInput {
                    msg: "captured Deflate decoder reported an invalid progress count".to_string(),
                }
                .into(),
            ));
        }
        input_offset = input_offset.checked_add(consumed).ok_or_else(|| {
            precompressed_archive_error(
                ErrorKind::InvalidInput {
                    msg: "captured Deflate input position overflows usize".to_string(),
                }
                .into(),
            )
        })?;
        if produced != 0 {
            compare_captured_decoded_chunk(
                &output[..produced],
                expected_decoded,
                &mut compared,
                &mut actual_crc,
                progress,
            )?;
        } else if let Err(error) = progress(PrecompressedProgress::Decoded {
            bytes: u64::try_from(compared).expect("decoded comparison count fits in u64"),
        }) {
            return Err(VerifiedPrecompressedError::Callback(error));
        }

        if status == Status::StreamEnd {
            break;
        }
        if consumed == 0 && produced == 0 {
            return Err(precompressed_archive_error(
                ErrorKind::InvalidInput {
                    msg: "captured Deflate stream ended before its final block".to_string(),
                }
                .into(),
            ));
        }
    }

    let expected_compressed = usize_to_u64(compressed.len(), "captured Deflate payload length")
        .map_err(precompressed_archive_error)?;
    let consumed = decoder.total_in();
    if consumed != expected_compressed || input_offset != compressed.len() {
        return Err(precompressed_archive_error(
            ErrorKind::InvalidSize {
                expected: expected_compressed,
                actual: consumed,
            }
            .into(),
        ));
    }
    finish_captured_decoded_verification(expected_decoded, expected, compared, actual_crc)
}

fn compare_captured_decoded_chunk<E, F>(
    chunk: &[u8],
    expected_decoded: &mut CapturedDecodedOutput<'_>,
    compared: &mut usize,
    actual_crc: &mut u32,
    progress: &mut F,
) -> Result<(), VerifiedPrecompressedError<E>>
where
    F: FnMut(PrecompressedProgress) -> Result<(), E>,
{
    let end = compared.checked_add(chunk.len()).ok_or_else(|| {
        precompressed_archive_error(
            ErrorKind::InvalidInput {
                msg: "captured decoded payload length overflows usize".to_string(),
            }
            .into(),
        )
    })?;
    expected_decoded.accept(chunk, *compared, end)?;
    *compared = end;
    *actual_crc = crc32_chunk(chunk, *actual_crc);
    if let Err(error) = progress(PrecompressedProgress::Decoded {
        bytes: u64::try_from(*compared).expect("decoded comparison count fits in u64"),
    }) {
        return Err(VerifiedPrecompressedError::Callback(error));
    }
    Ok(())
}

fn finish_captured_decoded_verification<E>(
    expected_decoded: &mut CapturedDecodedOutput<'_>,
    expected: ZipVerification,
    compared: usize,
    actual_crc: u32,
) -> Result<u32, VerifiedPrecompressedError<E>> {
    if compared != expected_decoded.len() {
        return Err(precompressed_archive_error(
            ErrorKind::InvalidSize {
                expected: u64::try_from(expected_decoded.len()).expect("decoded length fits u64"),
                actual: u64::try_from(compared).expect("decoded length fits u64"),
            }
            .into(),
        ));
    }
    let actual = ZipVerification {
        crc: actual_crc,
        uncompressed_size: u64::try_from(compared).expect("decoded length fits u64"),
    };
    expected
        .valid(actual)
        .map_err(precompressed_archive_error)?;
    Ok(actual_crc)
}

fn verify_captured_decoded_reader<E, F, R: Read>(
    reader: &mut R,
    expected_decoded: &mut CapturedDecodedOutput<'_>,
    expected: ZipVerification,
    progress: &mut F,
) -> Result<u32, VerifiedPrecompressedError<E>>
where
    F: FnMut(PrecompressedProgress) -> Result<(), E>,
{
    let mut buffer = [0_u8; VERIFIED_ENTRY_READER_BUFFER_SIZE];
    let mut compared = 0usize;
    let mut actual_crc = 0_u32;

    loop {
        let read =
            read_with_interrupt_budget(reader, &mut buffer).map_err(precompressed_decode_error)?;
        if read == 0 {
            break;
        }
        compare_captured_decoded_chunk(
            &buffer[..read],
            expected_decoded,
            &mut compared,
            &mut actual_crc,
            progress,
        )?;
    }
    finish_captured_decoded_verification(expected_decoded, expected, compared, actual_crc)
}

fn complete_verified_callback<T, E>(
    callback: Result<T, E>,
    verification: Result<(), VerifiedReaderFailure>,
) -> Result<T, VerifiedEntryReaderError<E>> {
    match (verification, callback) {
        (Ok(()), Ok(value)) => Ok(value),
        (Ok(()), Err(error)) => Err(VerifiedEntryReaderError::Callback(error)),
        (Err(failure), Ok(_)) => Err(failure.into_public(None)),
        (Err(VerifiedReaderFailure::Accounting(_)), Err(error)) => {
            Err(VerifiedEntryReaderError::Callback(error))
        },
        (Err(failure), Err(error)) => Err(failure.into_public(Some(error))),
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

        // Admit a bounded number of physical central records before retaining
        // any raw name. The fixed record is part of the high-level metadata
        // budget, and the central-directory byte span is an additional hard
        // bound even when callers disable the configured budget.
        let declared_entry_count = archive.entries_hint();
        let declared_central_size = archive.central_directory_size();
        let physical_bound = physical_entry_bound(
            declared_central_size,
            declared_entry_count,
            limits.max_metadata_bytes,
        )?;
        let declared_entry_count = usize::try_from(declared_entry_count).map_err(|_| {
            Error::from(ErrorKind::InvalidInput {
                msg: "ZIP central-directory entry count does not fit this platform".to_string(),
            })
        })?;
        let file_bound = physical_bound.min(limits.max_files);
        let mut index = HashMap::new();
        index.try_reserve(file_bound).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "archive reader file index",
                source,
            })
        })?;
        let mut directories = HashMap::new();
        directories.try_reserve(physical_bound).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "archive reader directory index",
                source,
            })
        })?;
        let mut total_metadata_bytes = 0u64;
        let mut total_uncompressed_size = 0u64;
        let mut ordered_names = Vec::new();
        ordered_names.try_reserve(file_bound).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "archive reader physical order",
                source,
            })
        })?;
        let mut layout = Vec::new();
        layout.try_reserve(physical_bound).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "archive reader physical layout",
                source,
            })
        })?;
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

            let metadata_bytes = entry.metadata_size_hint().checked_add(46).ok_or_else(|| {
                Error::from(ErrorKind::InvalidInput {
                    msg: "archive central-directory metadata size overflows u64".to_string(),
                })
            })?;
            let next_metadata_bytes = total_metadata_bytes
                .checked_add(metadata_bytes)
                .ok_or_else(|| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: "archive central-directory metadata total overflows u64".to_string(),
                    })
                })?;
            if next_metadata_bytes > limits.max_metadata_bytes {
                return Err(limit_error(
                    LimitResource::MetadataBytes,
                    next_metadata_bytes,
                    limits.max_metadata_bytes,
                ));
            }
            total_metadata_bytes = next_metadata_bytes;
            if layout.len() >= physical_bound {
                if layout.len() >= declared_entry_count {
                    return Err(Error::from(ErrorKind::InvalidInput {
                        msg: "central directory contains more records than EOCD declares"
                            .to_string(),
                    }));
                }
                return Err(limit_error(
                    LimitResource::MetadataBytes,
                    total_metadata_bytes,
                    limits.max_metadata_bytes,
                ));
            }
            layout.try_reserve(1).map_err(|source| {
                Error::from(ErrorKind::Allocation {
                    resource: "archive reader physical layout",
                    source,
                })
            })?;

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

            let (name, lossy_name) = normalized_member_name(path, "archive reader member key")?;
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
                directories.try_reserve(1).map_err(|source| {
                    Error::from(ErrorKind::Allocation {
                        resource: "archive reader directory index",
                        source,
                    })
                })?;
                let central_name =
                    clone_name_fallibly(path.as_ref(), "archive reader directory name")?;
                layout.push(BorrowedLayoutEntry {
                    wayfinder: entry.wayfinder(),
                    central_name,
                });
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

            let local_header_offset = entry.local_header_offset();
            if directories.contains_key(&name) {
                return Err(file_directory_collision_error(&name, lossy_name));
            }

            if index.contains_key(&name) {
                return Err(duplicate_member_error(
                    &name,
                    lossy_name,
                    "duplicate normalized file names",
                ));
            }
            index.try_reserve(1).map_err(|source| {
                Error::from(ErrorKind::Allocation {
                    resource: "archive reader file index",
                    source,
                })
            })?;
            ordered_names.try_reserve(1).map_err(|source| {
                Error::from(ErrorKind::Allocation {
                    resource: "archive reader physical order",
                    source,
                })
            })?;
            let central_name = clone_name_fallibly(path.as_ref(), "archive reader member name")?;
            let layout_name =
                clone_name_fallibly(&central_name, "archive reader physical member name")?;
            layout.push(BorrowedLayoutEntry {
                wayfinder: entry.wayfinder(),
                central_name: layout_name,
            });
            let index_name = clone_str_fallibly(&name, "archive reader file index key")?;
            index.insert(
                index_name,
                EntryInfo {
                    wayfinder: entry.wayfinder(),
                    flags: entry.flags(),
                    compression_method: entry.compression_method(),
                    uncompressed_size,
                    central_name,
                },
            );
            ordered_names.push((local_header_offset, name));
        }

        ordered_names.sort_by_key(|(offset, _)| *offset);
        let order = collect_order_fallibly(ordered_names, "archive reader file order")?;
        layout.sort_unstable_by_key(|entry| entry.wayfinder.local_header_offset());

        Ok(Self {
            archive,
            index,
            directories,
            order,
            layout,
            strict_layout_cache: StrictLayoutCache::new(),
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
        let Ok(lookup) = lookup_member_name(name) else {
            return false;
        };
        !lookup.explicit_directory && self.index.contains_key(&lookup.name)
    }

    /// Return declared metadata for a normalized member name.
    ///
    /// This performs only hash-map lookup over the central-directory index. It
    /// never reads, decompresses, verifies, or allocates member payload data.
    pub fn metadata(&self, name: &str) -> Result<Metadata, Error> {
        let lookup = lookup_member_name(name)?;
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
        let lookup = lookup_member_name(name)?;
        let info = self
            .index
            .get(&lookup.name)
            .filter(|_| !lookup.explicit_directory)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))?;
        Ok(info.compression_method == CompressionMethod::Store)
    }

    /// Start a sequential read operation that can reuse Deflate decoder state.
    #[must_use]
    pub fn read_session(&self) -> ArchiveReadSession<'_, 'data> {
        ArchiveReadSession {
            archive: self,
            decoder: None,
        }
    }

    /// Read and decompress a file from the archive.
    ///
    /// Returns the decompressed contents of the file. Supports both stored
    /// (uncompressed) and deflated entries.
    pub fn read(&self, name: &str) -> Result<Vec<u8>, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_with_accounting(name, &mut accounting)
    }

    /// Read and decompress a file while recording actual payload work.
    pub fn read_with_accounting(
        &self,
        name: &str,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<Vec<u8>, Error> {
        let lookup = lookup_member_name(name)?;

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
                let verification = verifier.valid(ZipVerification {
                    crc: crate::crc32(data),
                    uncompressed_size: usize_to_u64(data.len(), "stored payload length")?,
                });
                let accounting_result = accounting.add_stored_payload_bytes_read(usize_to_u64(
                    data.len(),
                    "stored payload bytes read",
                )?);
                if let Err(error) = verification {
                    drop(accounting_result);
                    return Err(error);
                }
                accounting_result?;
                let mut output = Vec::new();
                output.try_reserve_exact(data.len()).map_err(|source| {
                    Error::from(ErrorKind::Allocation {
                        resource: "archive entry output",
                        source,
                    })
                })?;
                output.extend_from_slice(data);
                accounting.add_stored_payload_bytes_accepted(usize_to_u64(
                    output.len(),
                    "stored payload bytes accepted",
                )?)?;
                Ok(output)
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
                let read_limit = info.uncompressed_size.checked_add(1).ok_or_else(|| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: format!(
                            "archive entry size {} cannot be bounded with an overrun sentinel",
                            info.uncompressed_size
                        ),
                    })
                })?;
                let read_capacity = usize::try_from(read_limit).map_err(|_| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: format!(
                            "archive entry size {} plus an overrun sentinel does not fit this platform",
                            info.uncompressed_size
                        ),
                    })
                })?;
                let mut decompressed = Vec::new();
                decompressed
                    .try_reserve_exact(read_capacity)
                    .map_err(|source| {
                        Error::from(ErrorKind::Allocation {
                            resource: "archive entry output",
                            source,
                        })
                    })?;

                let (result, compressed_read, produced) = {
                    let mut compressed = CountingReader::new(data);
                    let mut decoder = DeflateDecoder::new(&mut compressed);
                    let (result, produced_count) = {
                        let mut produced_reader = CountingReader::new(&mut decoder);
                        let verifier = entry.verifying_reader(&mut produced_reader);
                        let result = verifier
                            .take(read_limit)
                            .read_to_end(&mut decompressed)
                            .map_err(Error::from);
                        let produced_count = produced_reader.count();
                        (result, produced_count)
                    };
                    drop(decoder);
                    (result, compressed.count(), produced_count)
                };
                let accepted = usize_to_u64(decompressed.len(), "decompressed ZIP bytes accepted")?;
                let accounting_result = accounting
                    .add_compressed_deflate_payload_bytes_read(compressed_read)
                    .and_then(|()| accounting.add_deflate_bytes_produced(produced))
                    .and_then(|()| accounting.add_deflate_bytes_accepted(accepted));
                if let Err(error) = result {
                    drop(accounting_result);
                    return Err(error);
                }
                accounting_result?;
                if decompressed.len() != size {
                    return Err(ErrorKind::InvalidSize {
                        expected: info.uncompressed_size,
                        actual: usize_to_u64(decompressed.len(), "decompressed ZIP bytes")?,
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

    /// Borrow and verify a stored member without materializing a second copy.
    ///
    /// `Some` is returned only for unencrypted ZIP Store members. Unencrypted
    /// Deflated and otherwise unsupported members return `None`, allowing a
    /// caller to fall back to [`Self::read`]. Targets declaring encryption in
    /// general-purpose flag bits 0 or 6 return a typed error before this
    /// ineligible-compression fallback. Before the borrowed slice is
    /// published, the local header, data descriptor (when present), declared
    /// size, and CRC are validated. The returned bytes borrow the source
    /// archive for the lifetime of this reader and are never inserted into a
    /// decompression cache.
    ///
    /// Archive limits are admission limits: the member's declared metadata
    /// and size were checked by [`Self::new_with_limits`] before this method
    /// can be called. Since this method does not allocate payload storage, it
    /// does not create a separate materialization budget charge.
    ///
    /// For compatibility with the tolerant global ZIP verifier, a declared
    /// CRC-32 of zero normally means "not verified". Nonempty members with
    /// that declaration return `None` because borrowed access cannot publish
    /// an unverifiable source slice; callers can use [`Self::read`] for the
    /// owned fallback. Empty members with a zero CRC remain eligible after
    /// their structural checks.
    /// ZIP64 EOCD archives conservatively return `None`; callers can use
    /// [`Self::read`] or [`Self::read_to`] as the owned fallback for those
    /// archives.
    pub fn read_stored_borrowed(&self, name: &str) -> Result<Option<&'data [u8]>, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_stored_borrowed_with_accounting(name, &mut accounting)
    }

    /// Borrow and verify a stored member without charging payload work.
    ///
    /// Borrowed Store access intentionally has zero accounting work because it
    /// publishes the source slice directly and does not materialize or stream
    /// a payload.
    pub fn read_stored_borrowed_with_accounting(
        &self,
        name: &str,
        _accounting: &mut ZipOperationAccounting,
    ) -> Result<Option<&'data [u8]>, Error> {
        let lookup = lookup_member_name(name)?;

        let info = self
            .index
            .get(&lookup.name)
            .filter(|_| !lookup.explicit_directory)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))?;
        if info.flags & ((1 << 0) | (1 << 6)) != 0 {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "borrowed access refuses encrypted entries".to_string(),
            }));
        }
        if info.compression_method != CompressionMethod::Store {
            return Ok(None);
        }
        if self.archive.is_zip64() {
            return Ok(None);
        }
        if self
            .layout
            .iter()
            .any(|entry| !entry.wayfinder.borrowed_provenance_supported())
        {
            return Ok(None);
        }

        self.validate_borrowed_spans()?;
        let entry = self
            .archive
            .get_stored_entry_borrowed(info.wayfinder, &info.central_name)?;
        let data = entry.data();
        let verification = entry.claim_verifier();
        if !data.is_empty() && verification.crc() == 0 {
            return Ok(None);
        }
        let actual_crc = crate::crc32(data);
        verification.valid_strict(ZipVerification {
            crc: actual_crc,
            uncompressed_size: usize_to_u64(data.len(), "stored ZIP bytes")?,
        })?;
        Ok(Some(data))
    }

    fn validate_borrowed_spans(&self) -> Result<(), Error> {
        self.archive.validate_borrowed_layout()?;
        let mut previous_end = None;
        for entry in &self.layout {
            let (start, end) = self
                .archive
                .borrowed_entry_span(entry.wayfinder, &entry.central_name)?;
            if previous_end.is_some_and(|previous_end| start < previous_end) {
                return Err(Error::from(ErrorKind::InvalidInput {
                    msg: "borrowed access cannot prove non-overlapping ZIP spans".to_string(),
                }));
            }
            previous_end = Some(end);
        }
        if previous_end.is_some_and(|previous_end| previous_end > self.archive.directory_offset()) {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "borrowed access span extends into the central directory".to_string(),
            }));
        }
        Ok(())
    }

    /// Prove that the bytes one member declares are not claimed by any other
    /// record of this archive.
    ///
    /// The slice-backed reader pays no positional read, so the neighbour probe
    /// here is a 30-byte parse rather than a read; the acceptance contract is
    /// deliberately identical to the source-backed reader's.
    fn build_strict_layout_proof(
        &self,
        target: crate::ZipArchiveEntryWayfinder,
        memo: &StrictLayoutMemo,
    ) -> Result<StrictLayoutProven, Error> {
        self.archive.validate_borrowed_layout()?;
        prove_target_scoped_strict_layout(
            self.layout.len(),
            |position| self.layout.get(position).map(|entry| entry.wayfinder),
            |position| {
                let entry = self
                    .layout
                    .get(position)
                    .ok_or_else(strict_layout_index_error)?;
                self.archive
                    .validate_strict_entry_layout(entry.wayfinder, &entry.central_name)
            },
            |position| {
                let entry = self
                    .layout
                    .get(position)
                    .ok_or_else(strict_layout_index_error)?;
                self.archive.local_span_bound(entry.wayfinder)
            },
            |position, bound| {
                let entry = self
                    .layout
                    .get(position)
                    .ok_or_else(strict_layout_index_error)?;
                self.archive.resolve_span_end(entry.wayfinder, bound)
            },
            target,
            memo,
        )
    }

    fn strict_layout_for(
        &self,
        target: crate::ZipArchiveEntryWayfinder,
    ) -> Result<crate::StrictEntryLayout, Error> {
        strict_layout_for_cached(&self.strict_layout_cache, target, |memo| {
            self.build_strict_layout_proof(target, memo)
        })
    }

    /// Decompress and verify one member directly into a caller-owned sink.
    ///
    /// The sink receives at most the declared uncompressed member size. A
    /// successful return means the declared size and CRC have both been
    /// checked. The operation uses a fixed-size scratch buffer and does not
    /// retain a complete decompressed member.
    ///
    /// This method is not atomic: a sink may contain a valid prefix when the
    /// operation returns an I/O, checksum, or size error. The returned count
    /// is the number of bytes accepted by the sink on success. Archive entry
    /// limits are checked while constructing this reader.
    pub fn read_to<W: Write>(&self, name: &str, sink: &mut W) -> Result<u64, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_to_with_accounting(name, sink, &mut accounting)
    }

    /// Decompress and verify one member into a sink while recording actual
    /// source traversal and Deflate destination acceptance.
    pub fn read_to_with_accounting<W: Write>(
        &self,
        name: &str,
        sink: &mut W,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<u64, Error> {
        let lookup = lookup_member_name(name)?;
        let info = self
            .index
            .get(&lookup.name)
            .filter(|_| !lookup.explicit_directory)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))?;

        match info.compression_method {
            CompressionMethod::Store | CompressionMethod::Deflate => {},
            other => {
                return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                    other.as_id().as_u16(),
                )));
            },
        }
        self.archive.validate_strict_stream_target(info.wayfinder)?;
        if info.compression_method == CompressionMethod::Store
            && info.wayfinder.compressed_size_hint() != info.uncompressed_size
        {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: info.uncompressed_size,
                actual: info.wayfinder.compressed_size_hint(),
            }));
        }
        let target_layout = self.strict_layout_for(info.wayfinder)?;
        let verifier = target_layout.verifier;
        let payload = self.archive.strict_payload(target_layout)?;
        match info.compression_method {
            CompressionMethod::Store => {
                let mut source = CountingReader::new(payload);
                let result = stream_verified_with_accounting(
                    &mut source,
                    verifier,
                    sink,
                    accounting,
                    AccountingReadKind::Stored,
                );
                let result = result.and_then(|bytes| {
                    let consumed = source.count();
                    let expected = info.wayfinder.compressed_size_hint();
                    if consumed != expected {
                        return Err(Error::from(ErrorKind::InvalidSize {
                            expected,
                            actual: consumed,
                        }));
                    }
                    Ok(bytes)
                });
                let accounting_result = accounting.add_stored_payload_bytes_read(source.count());
                match result {
                    Err(error) => {
                        // Keep the observed source count even when the stream failed.  The
                        // stream error remains primary if recording the count overflows.
                        let _ = accounting_result;
                        Err(error)
                    },
                    Ok(bytes) => {
                        accounting_result?;
                        Ok(bytes)
                    },
                }
            },
            CompressionMethod::Deflate => {
                let (result, compressed_consumed, compressed_read) = {
                    let mut source = CountingReader::new(payload);
                    let mut decoder = DeflateDecoder::new(&mut source);
                    let result = stream_verified_with_accounting(
                        &mut decoder,
                        verifier,
                        sink,
                        accounting,
                        AccountingReadKind::Deflate,
                    );
                    let compressed_consumed = decoder.total_in();
                    drop(decoder);
                    let compressed_read = source.count();
                    (result, compressed_consumed, compressed_read)
                };
                let result = result.and_then(|bytes| {
                    let expected = info.wayfinder.compressed_size_hint();
                    if compressed_consumed != expected {
                        return Err(Error::from(ErrorKind::InvalidSize {
                            expected,
                            actual: compressed_consumed,
                        }));
                    }
                    Ok(bytes)
                });
                let accounting_result =
                    accounting.add_compressed_deflate_payload_bytes_read(compressed_read);
                match result {
                    Err(error) => {
                        // Keep the observed source count even when the stream failed.  The
                        // stream error remains primary if recording the count overflows.
                        let _ = accounting_result;
                        Err(error)
                    },
                    Ok(bytes) => {
                        accounting_result?;
                        Ok(bytes)
                    },
                }
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

impl<'archive, 'data> ArchiveReadSession<'archive, 'data> {
    /// Read and verify one member by name.
    ///
    /// Deflate decoder state is reused only within this session. Store members
    /// continue to use their independent source slice and never initialize
    /// the decoder.
    pub fn read(&mut self, name: &str) -> Result<Vec<u8>, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_with_accounting(name, &mut accounting)
    }

    /// Read and verify one member while recording actual payload work.
    pub fn read_with_accounting(
        &mut self,
        name: &str,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<Vec<u8>, Error> {
        let lookup = lookup_member_name(name)?;

        let info = self
            .archive
            .index
            .get(&lookup.name)
            .filter(|_| !lookup.explicit_directory)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))?;

        let entry = self.archive.archive.get_entry_borrowed(info.wayfinder)?;
        let data = entry.data();

        match info.compression_method {
            CompressionMethod::Store => {
                // Stored (uncompressed) - verify and return directly.
                let verifier = entry.claim_verifier();
                let verification = verifier.valid(ZipVerification {
                    crc: crate::crc32(data),
                    uncompressed_size: usize_to_u64(data.len(), "stored payload length")?,
                });
                let accounting_result = accounting.add_stored_payload_bytes_read(usize_to_u64(
                    data.len(),
                    "stored payload bytes read",
                )?);
                if let Err(error) = verification {
                    drop(accounting_result);
                    return Err(error);
                }
                accounting_result?;
                let mut output = Vec::new();
                output.try_reserve_exact(data.len()).map_err(|source| {
                    Error::from(ErrorKind::Allocation {
                        resource: "archive entry output",
                        source,
                    })
                })?;
                output.extend_from_slice(data);
                accounting.add_stored_payload_bytes_accepted(usize_to_u64(
                    output.len(),
                    "stored payload bytes accepted",
                )?)?;
                Ok(output)
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
                let read_limit = info.uncompressed_size.checked_add(1).ok_or_else(|| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: format!(
                            "archive entry size {} cannot be bounded with an overrun sentinel",
                            info.uncompressed_size
                        ),
                    })
                })?;
                let read_capacity = usize::try_from(read_limit).map_err(|_| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: format!(
                            "archive entry size {} plus an overrun sentinel does not fit this platform",
                            info.uncompressed_size
                        ),
                    })
                })?;
                let mut decompressed = Vec::new();
                decompressed
                    .try_reserve_exact(read_capacity)
                    .map_err(|source| {
                        Error::from(ErrorKind::Allocation {
                            resource: "archive entry output",
                            source,
                        })
                    })?;

                let (result, compressed_read, produced) = {
                    let decoder = match self.decoder.as_mut() {
                        Some(decoder) => {
                            let _previous = decoder.reset(CountingReader::new(data));
                            decoder
                        },
                        None => self
                            .decoder
                            .insert(DeflateDecoder::new(CountingReader::new(data))),
                    };
                    let (result, produced_count) = {
                        let mut produced_reader = CountingReader::new(&mut *decoder);
                        let verifier = entry.verifying_reader(&mut produced_reader);
                        let result = verifier
                            .take(read_limit)
                            .read_to_end(&mut decompressed)
                            .map_err(Error::from);
                        let produced_count = produced_reader.count();
                        (result, produced_count)
                    };
                    let compressed_read = decoder.get_ref().count();
                    (result, compressed_read, produced_count)
                };
                let accepted = usize_to_u64(decompressed.len(), "decompressed ZIP bytes accepted")?;
                let accounting_result = accounting
                    .add_compressed_deflate_payload_bytes_read(compressed_read)
                    .and_then(|()| accounting.add_deflate_bytes_produced(produced))
                    .and_then(|()| accounting.add_deflate_bytes_accepted(accepted));
                if let Err(error) = result {
                    // Keep the observed source count even when the stream failed. The
                    // stream error remains primary if recording the count overflows.
                    let _ = accounting_result;
                    return Err(error);
                }
                accounting_result?;
                if decompressed.len() != size {
                    return Err(ErrorKind::InvalidSize {
                        expected: info.uncompressed_size,
                        actual: usize_to_u64(decompressed.len(), "decompressed ZIP bytes")?,
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
    /// another EOCD search. `scratch` is used as the fast path while scanning
    /// the existing central directory; valid records larger than that
    /// recommendation use a bounded fallible spill buffer.
    pub fn preservation_index<'archive>(
        &'archive self,
        scratch: &mut [u8],
    ) -> Result<PreservationIndex<'archive, R>, Error> {
        self.preservation_index_with_limits(scratch, ArchiveLimits::default())
    }

    /// Build a raw-member preservation index under explicit archive limits.
    ///
    /// Metadata and member-name limits are applied to every retained source
    /// central record, including directory records, because preservation must
    /// own their exact raw metadata even though ordinary file indexes exclude
    /// directories. The file-count limit retains its non-directory meaning.
    pub fn preservation_index_with_limits<'archive>(
        &'archive self,
        scratch: &mut [u8],
        limits: ArchiveLimits,
    ) -> Result<PreservationIndex<'archive, R>, Error> {
        PreservationIndex::new_with_limits(&self.archive, scratch, limits)
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
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(RECOMMENDED_BUFFER_SIZE)
            .map_err(|source| {
                Error::from(ErrorKind::Allocation {
                    resource: "indexed archive locator scratch",
                    source,
                })
            })?;
        buffer.resize(RECOMMENDED_BUFFER_SIZE, 0);
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
        let declared_entry_count = archive.entries_hint();
        let declared_central_size = archive.central_directory_size();
        let physical_bound = physical_entry_bound(
            declared_central_size,
            declared_entry_count,
            limits.max_metadata_bytes,
        )?;
        let declared_entry_count = usize::try_from(declared_entry_count).map_err(|_| {
            Error::from(ErrorKind::InvalidInput {
                msg: "ZIP central-directory entry count does not fit this platform".to_string(),
            })
        })?;
        let file_bound = physical_bound.min(limits.max_files);
        let mut index = HashMap::new();
        index.try_reserve(file_bound).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "indexed archive file index",
                source,
            })
        })?;
        let mut entries = Vec::new();
        entries.try_reserve(file_bound).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "indexed archive entries",
                source,
            })
        })?;
        let mut directories = HashMap::new();
        directories.try_reserve(physical_bound).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "indexed archive directory index",
                source,
            })
        })?;
        let mut layout = Vec::new();
        layout.try_reserve(physical_bound).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "indexed archive physical layout",
                source,
            })
        })?;
        let mut ordered_entries = Vec::new();
        ordered_entries.try_reserve(file_bound).map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "indexed archive physical order",
                source,
            })
        })?;
        let mut total_metadata_bytes = 0_u64;
        let mut total_uncompressed_size = 0_u64;
        let mut strict_mimetype = None;
        let mut has_encrypted_entries = false;
        let mut has_data_descriptor_entries = false;
        let mut has_zip64_metadata = archive.is_zip64();
        let mut all_local_spans_bounded = true;
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(RECOMMENDED_BUFFER_SIZE)
            .map_err(|source| {
                Error::from(ErrorKind::Allocation {
                    resource: "indexed archive central-directory scratch",
                    source,
                })
            })?;
        buffer.resize(RECOMMENDED_BUFFER_SIZE, 0);

        {
            let mut central_entries = archive.entries_with_metadata_limit(&mut buffer, u64::MAX);
            while let Some(entry) = central_entries.next_entry()? {
                has_encrypted_entries |= entry.flags() & 1 != 0;
                has_data_descriptor_entries |= entry.has_data_descriptor();
                has_zip64_metadata |= entry.is_zip64();
                let local_span_end = entry
                    .local_header_offset()
                    .checked_add(30)
                    .and_then(|offset| offset.checked_add(entry.metadata_size_hint()))
                    .and_then(|offset| offset.checked_add(entry.compressed_size_hint()));
                all_local_spans_bounded &=
                    local_span_end.is_some_and(|end| end <= archive.directory_offset());
                let path = entry.file_path();
                let member_name_bytes = path.as_ref().len() as u64;
                if member_name_bytes > limits.max_member_name_bytes {
                    return Err(limit_error(
                        LimitResource::MemberNameBytes,
                        member_name_bytes,
                        limits.max_member_name_bytes,
                    ));
                }
                if layout.len() >= physical_bound {
                    if layout.len() >= declared_entry_count {
                        return Err(Error::from(ErrorKind::InvalidInput {
                            msg: "central directory contains more records than EOCD declares"
                                .to_string(),
                        }));
                    }
                    let actual = total_metadata_bytes.saturating_add(CENTRAL_FIXED_RECORD_BYTES);
                    return Err(limit_error(
                        LimitResource::MetadataBytes,
                        actual,
                        limits.max_metadata_bytes,
                    ));
                }
                let metadata_bytes = entry
                    .metadata_size_hint()
                    .checked_add(CENTRAL_FIXED_RECORD_BYTES)
                    .ok_or_else(|| {
                        Error::from(ErrorKind::InvalidInput {
                            msg: "archive central-directory metadata size overflows u64"
                                .to_string(),
                        })
                    })?;
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
                layout.try_reserve(1).map_err(|source| {
                    Error::from(ErrorKind::Allocation {
                        resource: "indexed archive physical layout",
                        source,
                    })
                })?;

                let compressed_size = entry.compressed_size_hint();
                let uncompressed_size = entry.uncompressed_size_hint();
                let (name, lossy_name) = match policy {
                    ArchiveValidationPolicy::Normalized => {
                        normalized_member_name(path, "indexed archive member key")?
                    },
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
                    directories.try_reserve(1).map_err(|source| {
                        Error::from(ErrorKind::Allocation {
                            resource: "indexed archive directory index",
                            source,
                        })
                    })?;
                    let central_name =
                        clone_name_fallibly(path.as_ref(), "indexed archive directory name")?;
                    layout.push(IndexedLayoutEntry {
                        wayfinder: entry.wayfinder(),
                        name: IndexedLayoutName::Directory(central_name),
                    });
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
                let local_header_offset = entry.local_header_offset();
                if index.contains_key(&name) {
                    return Err(duplicate_member_error(
                        &name,
                        lossy_name,
                        "duplicate normalized file names",
                    ));
                }
                index.try_reserve(1).map_err(|source| {
                    Error::from(ErrorKind::Allocation {
                        resource: "indexed archive file index",
                        source,
                    })
                })?;
                entries.try_reserve(1).map_err(|source| {
                    Error::from(ErrorKind::Allocation {
                        resource: "indexed archive entries",
                        source,
                    })
                })?;
                ordered_entries.try_reserve(1).map_err(|source| {
                    Error::from(ErrorKind::Allocation {
                        resource: "indexed archive physical order",
                        source,
                    })
                })?;
                let central_name =
                    clone_name_fallibly(path.as_ref(), "indexed archive member name")?;
                let index_name = clone_str_fallibly(&name, "indexed archive file index key")?;
                layout.push(IndexedLayoutEntry {
                    wayfinder: entry.wayfinder(),
                    name: IndexedLayoutName::Entry(entry_id),
                });
                index.insert(index_name, entry_id);
                ordered_entries.push((local_header_offset, entry_id));
                entries.push(IndexedEntry {
                    name,
                    info: EntryInfo {
                        wayfinder: entry.wayfinder(),
                        flags: entry.flags(),
                        compression_method: entry.compression_method(),
                        uncompressed_size,
                        central_name,
                    },
                });
            }
        }

        if matches!(policy, ArchiveValidationPolicy::StrictPackage) {
            validate_strict_mimetype(&archive, strict_mimetype)?;
        }

        ordered_entries.sort_unstable_by_key(|(offset, _)| *offset);
        layout.sort_unstable_by_key(|entry| entry.wayfinder.local_header_offset());
        let order = collect_order_fallibly(ordered_entries, "indexed archive file order")?;

        Ok(Self {
            archive,
            layout,
            index,
            entries,
            directories,
            order,
            has_encrypted_entries,
            has_data_descriptor_entries,
            has_zip64_metadata,
            all_local_spans_bounded,
            strict_layout_cache: StrictLayoutCache::new(),
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

    /// Conservative additional ownership required by one raw-preservation
    /// replacement replay. The value is a scalar so callers do not acquire
    /// ZIP implementation types. `None` indicates checked arithmetic overflow.
    /// The caller's metadata scratch buffer is separate and is not included.
    #[must_use]
    pub fn preservation_memory_upper_bound(&self) -> Option<u64> {
        crate::preserve::preservation_memory_upper_bound(
            self.preservation_entry_count(),
            self.preservation_metadata_bytes(),
        )
    }

    /// Memory upper bound for one callback-scoped verified reader.
    ///
    /// Store readers own the fixed verification state, callback buffer, and
    /// one strict positional source counter/reader. Deflated readers
    /// additionally own flate2's fixed compressed-input window and decoder
    /// value plus the locked zlib-rs inflate allocation envelope. The
    /// zero-sized type parameters below preserve the actual pointer/layout
    /// widths without exposing the source type. This method reports a scalar
    /// budget and keeps backend types private.
    pub fn verified_entry_reader_memory_upper_bound(
        &self,
        entry_id: EntryId,
    ) -> Result<u64, Error> {
        let entry = self.indexed_entry(entry_id)?;
        let fixed_state = size_of::<VerifiedEntryBufReader<'static, &'static mut ()>>()
            .checked_add(size_of::<CountingReader<ZipReader<&'static ()>>>())
            .ok_or_else(|| {
                Error::from(ErrorKind::InvalidInput {
                    msg: "verified reader fixed state size overflows usize".to_string(),
                })
            })?;
        let mut amount = u64::try_from(fixed_state).map_err(|_| {
            Error::from(ErrorKind::InvalidInput {
                msg: "verified reader fixed state size exceeds u64".to_string(),
            })
        })?;
        if entry.info.compression_method == CompressionMethod::Deflate {
            amount = amount
                .checked_add(
                    u64::try_from(FLATE2_DEFLATE_INPUT_BUFFER_SIZE).map_err(|_| {
                        Error::from(ErrorKind::InvalidInput {
                            msg: "Deflate input buffer size exceeds u64".to_string(),
                        })
                    })?,
                )
                .and_then(|value| {
                    value.checked_add(
                        u64::try_from(size_of::<DeflateDecoder<&'static mut ()>>()).ok()?,
                    )
                })
                .and_then(|value| {
                    value.checked_add(u64::try_from(ZLIB_RS_INFLATE_STATE_UPPER_BOUND_BYTES).ok()?)
                })
                .ok_or_else(|| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: "verified Deflate reader memory bound overflows u64".to_string(),
                    })
                })?;
        }
        Ok(amount)
    }

    /// Whether any central-directory entry declares traditional ZIP
    /// encryption through general-purpose bit zero.
    #[inline]
    pub fn has_encrypted_entries(&self) -> bool {
        self.has_encrypted_entries
    }

    /// Whether any indexed central record declares a data descriptor.
    ///
    /// This accessor is hidden implementation plumbing for format owners
    /// that retain an [`IndexedArchive`].  It returns the fact captured by the
    /// archive's existing central-directory pass and never exposes a raw ZIP
    /// record or wayfinder.
    #[doc(hidden)]
    #[must_use]
    pub fn has_data_descriptor_entries(&self) -> bool {
        self.has_data_descriptor_entries
    }

    /// Whether the located archive uses ZIP64 end-of-central-directory
    /// framing. This excludes projected ZIP64 per-entry fields in an
    /// otherwise ZIP32 archive.
    #[doc(hidden)]
    #[must_use]
    pub fn archive_is_zip64(&self) -> bool {
        self.archive.is_zip64()
    }

    /// Whether the located archive or any indexed central record uses ZIP64
    /// metadata.
    ///
    /// This accessor is hidden implementation plumbing for format owners
    /// that need a conservative catalog decision without a second central
    /// directory pass.
    #[doc(hidden)]
    #[must_use]
    pub fn has_zip64_metadata(&self) -> bool {
        self.has_zip64_metadata
    }

    /// Whether every central record's declared local span is bounded by the
    /// located central directory.
    ///
    /// Directory records are included.  The result is captured during index
    /// construction and is therefore a zero-pass metadata fact for callers.
    #[doc(hidden)]
    #[must_use]
    pub fn all_local_spans_bounded(&self) -> bool {
        self.all_local_spans_bounded
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
        let Ok(lookup) = lookup_member_name(name) else {
            return None;
        };
        if lookup.explicit_directory {
            return None;
        }
        self.index.get(&lookup.name).copied()
    }

    /// Return declared metadata for a member without payload access.
    pub fn metadata(&self, name: &str) -> Result<Metadata, Error> {
        let lookup = lookup_member_name(name)?;
        if !lookup.explicit_directory {
            if let Some(entry_id) = self.index.get(&lookup.name).copied() {
                return self.metadata_for(entry_id);
            }
        }
        self.directories
            .get(&lookup.name)
            .copied()
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))
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
        let lookup = lookup_member_name(name)?;
        let entry_id = self
            .index
            .get(&lookup.name)
            .filter(|_| !lookup.explicit_directory)
            .copied()
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))?;
        Ok(self.indexed_entry(entry_id)?.info.compression_method == CompressionMethod::Store)
    }

    /// The central name one layout position validates against.
    fn strict_layout_central_name(&self, position: usize) -> Result<&[u8], Error> {
        let layout_entry = self
            .layout
            .get(position)
            .ok_or_else(strict_layout_index_error)?;
        Ok(match &layout_entry.name {
            IndexedLayoutName::Entry(entry_id) => {
                &self
                    .entries
                    .get(entry_id.0)
                    .ok_or_else(|| {
                        Error::from(ErrorKind::InvalidInput {
                            msg: "indexed strict layout references an unknown entry".to_string(),
                        })
                    })?
                    .info
                    .central_name
            },
            IndexedLayoutName::Directory(name) => name,
        })
    }

    /// Where the member after `position` begins, or the central directory for
    /// the last one.
    ///
    /// This is a read bound for the prover's speculative local-header read and
    /// nothing else.  The target's span is still checked against its
    /// successor's offset below, so a window that stops short of a variable
    /// region only costs one more read; it never changes a verdict.
    fn strict_layout_read_bound(&self, position: usize) -> u64 {
        position
            .checked_add(1)
            .and_then(|next| self.layout.get(next))
            .map_or_else(
                || self.archive.directory_offset(),
                |next| next.wayfinder.local_header_offset(),
            )
    }

    /// Prove that the bytes one member declares are not claimed by any other
    /// record of this archive.
    fn build_strict_layout_proof(
        &self,
        target: crate::ZipArchiveEntryWayfinder,
        memo: &StrictLayoutMemo,
    ) -> Result<StrictLayoutProven, Error> {
        prove_target_scoped_strict_layout(
            self.layout.len(),
            |position| self.layout.get(position).map(|entry| entry.wayfinder),
            |position| {
                let layout_entry = self
                    .layout
                    .get(position)
                    .ok_or_else(strict_layout_index_error)?;
                let central_name = self.strict_layout_central_name(position)?;
                self.archive.validate_strict_entry_layout(
                    layout_entry.wayfinder,
                    central_name,
                    self.strict_layout_read_bound(position),
                )
            },
            |position| {
                let layout_entry = self
                    .layout
                    .get(position)
                    .ok_or_else(strict_layout_index_error)?;
                self.archive.local_span_bound(layout_entry.wayfinder)
            },
            |position, bound| {
                let layout_entry = self
                    .layout
                    .get(position)
                    .ok_or_else(strict_layout_index_error)?;
                self.archive.resolve_span_end(layout_entry.wayfinder, bound)
            },
            target,
            memo,
        )
    }

    fn strict_layout_for(
        &self,
        target: crate::ZipArchiveEntryWayfinder,
    ) -> Result<crate::StrictEntryLayout, Error> {
        strict_layout_for_cached(&self.strict_layout_cache, target, |memo| {
            self.build_strict_layout_proof(target, memo)
        })
    }

    /// Start a sequential read operation that can reuse Deflate decoder state.
    #[must_use]
    pub fn read_session(&self) -> IndexedReadSession<'_, R> {
        IndexedReadSession::new(self)
    }

    /// Read and verify one member by name.
    pub fn read(&self, name: &str) -> Result<Vec<u8>, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_with_accounting(name, &mut accounting)
    }

    /// Read and verify one member while recording actual payload work.
    pub fn read_with_accounting(
        &self,
        name: &str,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<Vec<u8>, Error> {
        let entry_id = self.entry_id_for_name(name)?;
        self.read_entry_with_accounting(entry_id, accounting)
    }

    /// Resolve one member name to its stable opaque entry ID.
    ///
    /// This is the single admission step shared by the one-shot reads and by
    /// [`IndexedReadSession`], so name normalization, the explicit-directory
    /// rejection, and the `FileNotFound` identity cannot drift between them.
    fn entry_id_for_name(&self, name: &str) -> Result<EntryId, Error> {
        let lookup = lookup_member_name(name)?;
        self.index
            .get(&lookup.name)
            .filter(|_| !lookup.explicit_directory)
            .copied()
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))
    }

    /// Run a callback against one verified indexed member without retaining
    /// the complete decoded payload.
    ///
    /// The callback receives a fixed-buffer [`BufRead`] view. It may return a
    /// valid prefix early; the reader then drains and verifies the remainder
    /// before returning success. A callback panic unwinds normally and does
    /// not trigger a drain. Callback errors are returned only after successful
    /// archive finalization.
    pub fn with_verified_entry_reader<T, E, F>(
        &self,
        entry_id: EntryId,
        callback: F,
    ) -> Result<T, VerifiedEntryReaderError<E>>
    where
        F: for<'reader> FnOnce(&'reader mut dyn BufRead) -> Result<T, E>,
    {
        let mut accounting = ZipOperationAccounting::default();
        self.with_verified_entry_reader_with_accounting_mode(
            entry_id,
            callback,
            &mut accounting,
            true,
        )
    }

    /// Run a callback against one indexed member without retaining the
    /// complete decoded payload, aborting as soon as the callback returns an
    /// error.
    ///
    /// A successful callback is still drained and fully verified. When the
    /// callback returns an error, the reader is dropped immediately after any
    /// archive or transport failure already observed by the callback is
    /// retained. The remaining member is therefore not verified on that
    /// error path. Callback panics unwind without a drain.
    pub fn with_verified_entry_reader_abortable<T, E, F>(
        &self,
        entry_id: EntryId,
        callback: F,
    ) -> Result<T, VerifiedEntryReaderError<E>>
    where
        F: for<'reader> FnOnce(&'reader mut dyn BufRead) -> Result<T, E>,
    {
        let mut accounting = ZipOperationAccounting::default();
        self.with_verified_entry_reader_with_accounting_mode(
            entry_id,
            callback,
            &mut accounting,
            false,
        )
    }

    /// Run a callback against one verified indexed member while recording
    /// decoded bytes produced, callback-accepted bytes, and physical source
    /// traversal in `accounting`.
    ///
    /// The callback receives a fixed-buffer [`BufRead`] view. Any bytes read
    /// after the callback stops are drained for verification and are counted
    /// as produced/source bytes but not as callback-accepted bytes.
    pub fn with_verified_entry_reader_with_accounting<T, E, F>(
        &self,
        entry_id: EntryId,
        callback: F,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<T, VerifiedEntryReaderError<E>>
    where
        F: for<'reader> FnOnce(&'reader mut dyn BufRead) -> Result<T, E>,
    {
        self.with_verified_entry_reader_with_accounting_mode(entry_id, callback, accounting, true)
    }

    fn with_verified_entry_reader_with_accounting_mode<T, E, F>(
        &self,
        entry_id: EntryId,
        callback: F,
        accounting: &mut ZipOperationAccounting,
        drain_on_callback_error: bool,
    ) -> Result<T, VerifiedEntryReaderError<E>>
    where
        F: for<'reader> FnOnce(&'reader mut dyn BufRead) -> Result<T, E>,
    {
        let indexed = match self.indexed_entry(entry_id) {
            Ok(indexed) => indexed,
            Err(source) => {
                return Err(VerifiedEntryReaderError::Archive {
                    error: source,
                    callback_error: None,
                });
            },
        };
        let wayfinder = indexed.info.wayfinder;
        let compression_method = indexed.info.compression_method;
        match compression_method {
            CompressionMethod::Store | CompressionMethod::Deflate => {},
            other => {
                return Err(VerifiedEntryReaderError::Archive {
                    error: ErrorKind::UnsupportedCompressionMethod(other.as_id().as_u16()).into(),
                    callback_error: None,
                });
            },
        }
        if let Err(source) = self.archive.validate_strict_stream_target(wayfinder) {
            return Err(VerifiedEntryReaderError::Archive {
                error: source,
                callback_error: None,
            });
        }
        if compression_method == CompressionMethod::Store
            && wayfinder.compressed_size_hint() != wayfinder.uncompressed_size_hint()
        {
            return Err(VerifiedEntryReaderError::Archive {
                error: ErrorKind::InvalidSize {
                    expected: wayfinder.uncompressed_size_hint(),
                    actual: wayfinder.compressed_size_hint(),
                }
                .into(),
                callback_error: None,
            });
        }
        let target_layout = match self.strict_layout_for(wayfinder) {
            Ok(layout) => layout,
            Err(source) => {
                return Err(VerifiedEntryReaderError::Archive {
                    error: source,
                    callback_error: None,
                });
            },
        };
        let payload = self.archive.strict_payload_reader(wayfinder, target_layout);
        let mut callback = Some(callback);

        match compression_method {
            CompressionMethod::Store => {
                let callback = match callback.take() {
                    Some(callback) => callback,
                    None => {
                        return Err(VerifiedEntryReaderError::Archive {
                            error: ErrorKind::InvalidInput {
                                msg: "verified ZIP callback was consumed before dispatch"
                                    .to_string(),
                            }
                            .into(),
                            callback_error: None,
                        });
                    },
                };
                let mut source = CountingReader::new(payload);
                let (callback_result, finalization) = {
                    let mut reader = VerifiedEntryBufReader::new(
                        &mut source,
                        target_layout.verifier,
                        Some(accounting),
                        AccountingReadKind::Stored,
                    );
                    let callback_result = callback(&mut reader);
                    let finalization = if drain_on_callback_error || callback_result.is_ok() {
                        reader.finish()
                    } else {
                        reader.abort()
                    };
                    (callback_result, finalization)
                };
                let compressed_consumed = source.count();
                let accounting_result = accounting
                    .add_stored_payload_bytes_read(compressed_consumed)
                    .map_err(VerifiedReaderFailure::Accounting);
                let verification = match finalization {
                    Err(failure) => Err(failure),
                    Ok(()) if drain_on_callback_error || callback_result.is_ok() => {
                        if compressed_consumed != wayfinder.compressed_size_hint() {
                            Err(VerifiedReaderFailure::Archive(
                                ErrorKind::InvalidSize {
                                    expected: wayfinder.compressed_size_hint(),
                                    actual: compressed_consumed,
                                }
                                .into(),
                            ))
                        } else {
                            accounting_result
                        }
                    },
                    Ok(()) => accounting_result,
                };
                complete_verified_callback(callback_result, verification)
            },
            CompressionMethod::Deflate => {
                let callback = match callback.take() {
                    Some(callback) => callback,
                    None => {
                        return Err(VerifiedEntryReaderError::Archive {
                            error: ErrorKind::InvalidInput {
                                msg: "verified ZIP callback was consumed before dispatch"
                                    .to_string(),
                            }
                            .into(),
                            callback_error: None,
                        });
                    },
                };
                let mut source = CountingReader::new(payload);
                let (callback_result, finalization, compressed_consumed, compressed_read) = {
                    let mut decoder = DeflateDecoder::new(&mut source);
                    let (callback_result, finalization) = {
                        let mut reader = VerifiedEntryBufReader::new(
                            &mut decoder,
                            target_layout.verifier,
                            Some(accounting),
                            AccountingReadKind::Deflate,
                        );
                        let callback_result = callback(&mut reader);
                        let finalization = if drain_on_callback_error || callback_result.is_ok() {
                            reader.finish()
                        } else {
                            reader.abort()
                        };
                        (callback_result, finalization)
                    };
                    let compressed_consumed = decoder.total_in();
                    drop(decoder);
                    (
                        callback_result,
                        finalization,
                        compressed_consumed,
                        source.count(),
                    )
                };
                let accounting_result = accounting
                    .add_compressed_deflate_payload_bytes_read(compressed_read)
                    .map_err(VerifiedReaderFailure::Accounting);
                let verification = match finalization {
                    Err(failure) => Err(failure),
                    Ok(()) if drain_on_callback_error || callback_result.is_ok() => {
                        if compressed_consumed != wayfinder.compressed_size_hint() {
                            Err(VerifiedReaderFailure::Archive(
                                ErrorKind::InvalidSize {
                                    expected: wayfinder.compressed_size_hint(),
                                    actual: compressed_consumed,
                                }
                                .into(),
                            ))
                        } else {
                            accounting_result
                        }
                    },
                    Ok(()) => accounting_result,
                };
                complete_verified_callback(callback_result, verification)
            },
            _ => Err(VerifiedEntryReaderError::Archive {
                error: ErrorKind::InvalidInput {
                    msg: "unsupported verified ZIP compression dispatch".to_string(),
                }
                .into(),
                callback_error: None,
            }),
        }
    }

    /// Read and verify one member by its stable opaque entry ID.
    ///
    /// ZIP local-header, data-descriptor, decompressed-size, and CRC checks are
    /// intentionally deferred until this method is called.
    pub fn read_entry(&self, entry_id: EntryId) -> Result<Vec<u8>, Error> {
        let mut session = self.read_session();
        session.read_entry(entry_id)
    }

    /// Read and verify one indexed member while recording actual payload work.
    pub fn read_entry_with_accounting(
        &self,
        entry_id: EntryId,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<Vec<u8>, Error> {
        let mut session = self.read_session();
        session.read_entry_with_accounting(entry_id, accounting)
    }

    /// Capture and verify the exact compressed payload for one indexed member.
    ///
    /// `expected_decoded` must be the logical bytes that the caller has
    /// already validated.  The method first proves the source layout and
    /// captures exactly the bounded compressed range, invoking `progress` for
    /// each fixed-size capture chunk.  It then decodes the immutable capture,
    /// compares every decoded chunk with `expected_decoded`, computes the
    /// actual CRC, and checks the authoritative central or data-descriptor
    /// metadata.  The returned token is the only way this compressed payload
    /// can reach the preservation writer.
    ///
    /// A progress failure aborts immediately as
    /// [`VerifiedPrecompressedError::Callback`].  This operation deliberately
    /// does not drain a huge or malformed member after cancellation merely to
    /// preserve the ordinary decoded callback reader's secondary-error
    /// behavior.
    pub fn read_entry_precompressed_with_progress<E, F>(
        &self,
        entry_id: EntryId,
        expected_decoded: &[u8],
        progress: F,
    ) -> Result<VerifiedPrecompressedEntry, VerifiedPrecompressedError<E>>
    where
        F: FnMut(PrecompressedProgress) -> Result<(), E>,
    {
        self.capture_precompressed(
            entry_id,
            &mut CapturedDecodedOutput::Compare(expected_decoded),
            progress,
        )
    }

    /// Read logical bytes and issue a verified compressed token in one source capture.
    ///
    /// Unlike [`Self::read_entry_precompressed_with_progress`], this method does
    /// not require a prior decoded read. Store/Deflate bytes are captured once,
    /// decoded once, and checked against the authoritative size and CRC before
    /// either result is returned. The caller must still validate the decoded
    /// bytes semantically before authorizing a document-level transfer.
    ///
    /// Retains the complete compressed and decoded payloads simultaneously, plus
    /// bounded capture/decoder scratch. Archive limits apply; callers with an
    /// external memory budget must reserve both declared sizes before calling.
    /// The decoded callback reports bounded progress, not semantic acceptance.
    /// Callback failure aborts immediately and returns neither result.
    pub fn read_entry_precompressed_and_decoded_with_progress<E, F>(
        &self,
        entry_id: EntryId,
        progress: F,
    ) -> Result<(VerifiedPrecompressedEntry, Vec<u8>), VerifiedPrecompressedError<E>>
    where
        F: FnMut(PrecompressedProgress) -> Result<(), E>,
    {
        let declared = self
            .indexed_entry(entry_id)
            .map_err(precompressed_archive_error)?
            .info
            .uncompressed_size;
        let length = usize::try_from(declared).map_err(|_| {
            precompressed_archive_error(
                ErrorKind::InvalidInput {
                    msg: format!("decoded ZIP payload size {declared} does not fit this platform"),
                }
                .into(),
            )
        })?;
        let mut output = CapturedDecodedOutput::Collect {
            bytes: Vec::new(),
            length,
        };
        let token = self.capture_precompressed(entry_id, &mut output, progress)?;
        let CapturedDecodedOutput::Collect { bytes, .. } = output else {
            unreachable!("decoded capture uses a collecting output")
        };
        Ok((token, bytes))
    }

    fn capture_precompressed<E, F>(
        &self,
        entry_id: EntryId,
        expected_decoded: &mut CapturedDecodedOutput<'_>,
        mut progress: F,
    ) -> Result<VerifiedPrecompressedEntry, VerifiedPrecompressedError<E>>
    where
        F: FnMut(PrecompressedProgress) -> Result<(), E>,
    {
        let indexed = self
            .indexed_entry(entry_id)
            .map_err(precompressed_archive_error)?;
        let method = indexed.info.compression_method;
        match method {
            CompressionMethod::Store | CompressionMethod::Deflate => {},
            other => {
                return Err(precompressed_archive_error(
                    ErrorKind::UnsupportedCompressionMethod(other.as_id().as_u16()).into(),
                ));
            },
        }

        let expected_decoded_size = usize_to_u64(
            expected_decoded.len(),
            "verified precompressed decoded payload length",
        )
        .map_err(precompressed_archive_error)?;
        if expected_decoded_size != indexed.info.uncompressed_size {
            return Err(precompressed_archive_error(
                ErrorKind::InvalidSize {
                    expected: indexed.info.uncompressed_size,
                    actual: expected_decoded_size,
                }
                .into(),
            ));
        }

        let wayfinder = indexed.info.wayfinder;
        self.archive
            .validate_strict_stream_target(wayfinder)
            .map_err(precompressed_archive_error)?;
        if method == CompressionMethod::Store
            && wayfinder.compressed_size_hint() != indexed.info.uncompressed_size
        {
            return Err(precompressed_archive_error(
                ErrorKind::InvalidSize {
                    expected: indexed.info.uncompressed_size,
                    actual: wayfinder.compressed_size_hint(),
                }
                .into(),
            ));
        }

        let target_layout = self
            .strict_layout_for(wayfinder)
            .map_err(precompressed_archive_error)?;
        let compressed_size = wayfinder.compressed_size_hint();
        let compressed_capacity = usize::try_from(compressed_size).map_err(|_| {
            precompressed_archive_error(
                ErrorKind::InvalidInput {
                    msg: format!(
                        "compressed ZIP payload size {compressed_size} does not fit this platform"
                    ),
                }
                .into(),
            )
        })?;
        expected_decoded.reserve()?;
        let mut compressed = Vec::new();
        compressed
            .try_reserve_exact(compressed_capacity)
            .map_err(|source| {
                precompressed_archive_error(
                    ErrorKind::Allocation {
                        resource: "verified precompressed payload",
                        source,
                    }
                    .into(),
                )
            })?;

        let mut captured = 0_u64;
        let mut chunk = [0_u8; PRECOMPRESSED_CAPTURE_BUFFER_SIZE];
        let mut source = self.archive.strict_payload_reader(wayfinder, target_layout);
        while captured < compressed_size {
            let remaining = compressed_size
                .checked_sub(captured)
                .expect("compressed capture position is bounded");
            let request = usize::try_from(
                remaining.min(u64::try_from(chunk.len()).expect("capture chunk fits in u64")),
            )
            .expect("bounded capture chunk fits in usize");
            let read = match source.read(&mut chunk[..request]) {
                Ok(read) => read,
                Err(error) => {
                    return Err(VerifiedPrecompressedError::Transport(error));
                },
            };
            if read == 0 {
                return Err(precompressed_archive_error(
                    ErrorKind::InvalidSize {
                        expected: compressed_size,
                        actual: captured,
                    }
                    .into(),
                ));
            }
            let read_u64 = u64::try_from(read).expect("read length fits in u64");
            captured = match captured.checked_add(read_u64) {
                Some(captured) => captured,
                None => {
                    return Err(precompressed_archive_error(
                        ErrorKind::InvalidInput {
                            msg: "compressed capture byte count overflows u64".to_string(),
                        }
                        .into(),
                    ));
                },
            };
            compressed.extend_from_slice(&chunk[..read]);
            if let Err(error) = progress(PrecompressedProgress::Compressed { bytes: captured }) {
                return Err(VerifiedPrecompressedError::Callback(error));
            }
        }

        let expected_verifier = match source.claim_verifier() {
            Ok(verifier) => verifier,
            Err(error) => {
                return Err(precompressed_archive_error(error));
            },
        };
        let actual_crc = match verify_captured_precompressed_payload(
            method,
            compressed.as_slice(),
            expected_decoded,
            expected_verifier,
            &mut progress,
        ) {
            Ok(crc) => crc,
            Err(error) => {
                return Err(error);
            },
        };

        let compressed_size = match usize_to_u64(
            compressed.len(),
            "verified precompressed compressed payload length",
        ) {
            Ok(size) => size,
            Err(error) => {
                return Err(precompressed_archive_error(error));
            },
        };
        let token = VerifiedPrecompressedEntry {
            method,
            compressed: Arc::new(compressed),
            compressed_size,
            uncompressed_size: expected_decoded_size,
            crc32: actual_crc,
        };
        Ok(token)
    }

    /// Decompress and verify one indexed member directly into a caller-owned
    /// sink without retaining the complete decompressed member.
    ///
    /// The sink may contain a valid prefix when an I/O, checksum, or size
    /// error is returned. A successful return reports the number of bytes
    /// accepted by the sink. Entry limits were checked while constructing the
    /// index.
    pub fn read_entry_to<W: Write>(&self, entry_id: EntryId, sink: &mut W) -> Result<u64, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_entry_to_with_accounting(entry_id, sink, &mut accounting)
    }

    /// Decompress and verify one indexed member into a sink while recording
    /// actual source traversal and Deflate destination acceptance.
    pub fn read_entry_to_with_accounting<W: Write>(
        &self,
        entry_id: EntryId,
        sink: &mut W,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<u64, Error> {
        let indexed = self.indexed_entry(entry_id)?;
        match indexed.info.compression_method {
            CompressionMethod::Store | CompressionMethod::Deflate => {},
            other => {
                return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                    other.as_id().as_u16(),
                )));
            },
        }
        self.archive
            .validate_strict_stream_target(indexed.info.wayfinder)?;
        if indexed.info.compression_method == CompressionMethod::Store
            && indexed.info.wayfinder.compressed_size_hint() != indexed.info.uncompressed_size
        {
            return Err(Error::from(ErrorKind::InvalidSize {
                expected: indexed.info.uncompressed_size,
                actual: indexed.info.wayfinder.compressed_size_hint(),
            }));
        }
        let target_layout = self.strict_layout_for(indexed.info.wayfinder)?;
        let verifier = target_layout.verifier;
        let payload = self
            .archive
            .strict_payload_reader(indexed.info.wayfinder, target_layout);
        match indexed.info.compression_method {
            CompressionMethod::Store => {
                let mut source = CountingReader::new(payload);
                let result = stream_verified_with_accounting(
                    &mut source,
                    verifier,
                    sink,
                    accounting,
                    AccountingReadKind::Stored,
                );
                let result = result.and_then(|bytes| {
                    let consumed = source.count();
                    let expected = indexed.info.wayfinder.compressed_size_hint();
                    if consumed != expected {
                        return Err(Error::from(ErrorKind::InvalidSize {
                            expected,
                            actual: consumed,
                        }));
                    }
                    Ok(bytes)
                });
                let accounting_result = accounting.add_stored_payload_bytes_read(source.count());
                match result {
                    Err(error) => {
                        drop(accounting_result);
                        Err(error)
                    },
                    Ok(bytes) => {
                        accounting_result?;
                        Ok(bytes)
                    },
                }
            },
            CompressionMethod::Deflate => {
                let (result, compressed_consumed, compressed_read) = {
                    let mut source = CountingReader::new(payload);
                    let mut decoder = DeflateDecoder::new(&mut source);
                    let result = stream_verified_with_accounting(
                        &mut decoder,
                        verifier,
                        sink,
                        accounting,
                        AccountingReadKind::Deflate,
                    );
                    let compressed_consumed = decoder.total_in();
                    drop(decoder);
                    let compressed_read = source.count();
                    (result, compressed_consumed, compressed_read)
                };
                let result = result.and_then(|bytes| {
                    let expected = indexed.info.wayfinder.compressed_size_hint();
                    if compressed_consumed != expected {
                        return Err(Error::from(ErrorKind::InvalidSize {
                            expected,
                            actual: compressed_consumed,
                        }));
                    }
                    Ok(bytes)
                });
                let accounting_result =
                    accounting.add_compressed_deflate_payload_bytes_read(compressed_read);
                match result {
                    Err(error) => {
                        drop(accounting_result);
                        Err(error)
                    },
                    Ok(bytes) => {
                        accounting_result?;
                        Ok(bytes)
                    },
                }
            },
            other => Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                other.as_id().as_u16(),
            ))),
        }
    }

    /// Decompress and verify one member by normalized name directly into a
    /// caller-owned sink.
    ///
    /// This is the positional-source counterpart to [`Self::read_to`]. The
    /// sink may contain a valid prefix when an error is returned.
    pub fn read_to<W: Write>(&self, name: &str, sink: &mut W) -> Result<u64, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_to_with_accounting(name, sink, &mut accounting)
    }

    /// Decompress and verify one indexed member by name into a sink while
    /// recording actual source traversal and Deflate destination acceptance.
    pub fn read_to_with_accounting<W: Write>(
        &self,
        name: &str,
        sink: &mut W,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<u64, Error> {
        let lookup = lookup_member_name(name)?;
        let entry_id = self
            .index
            .get(&lookup.name)
            .filter(|_| !lookup.explicit_directory)
            .copied()
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(lookup.name)))?;
        self.read_entry_to_with_accounting(entry_id, sink, accounting)
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

fn normalized_member_name(
    path: ZipFilePath<RawPath<'_>>,
    resource: &'static str,
) -> Result<(String, bool), Error> {
    match std::str::from_utf8(path.as_ref()) {
        Ok(valid) => Ok((normalize_str_fallibly(valid, resource)?, false)),
        Err(_) => {
            // Keep the existing lossy-UTF-8 compatibility behavior, but apply
            // the same path normalization as valid UTF-8 names so a name
            // returned by `file_names()` is always a usable lookup key.
            let lossy = lossy_string_fallibly(path.as_ref(), resource)?;
            Ok((normalize_str_fallibly(&lossy, resource)?, true))
        },
    }
}

fn canonical_member_name(mut name: String) -> String {
    let canonical_len = name.trim_end_matches('/').len();
    name.truncate(canonical_len);
    name
}

fn normalize_str_fallibly(name: &str, resource: &'static str) -> Result<String, Error> {
    let name = name.rfind(':').map_or(name, |offset| &name[offset + 1..]);
    let mut normalized = String::new();
    normalized
        .try_reserve_exact(name.len())
        .map_err(|source| Error::from(ErrorKind::Allocation { resource, source }))?;
    for component in name.split(['/', '\\']) {
        if component.is_empty() || component == "." {
            continue;
        }
        if component == ".." {
            let last_separator = normalized.rfind('/');
            normalized.truncate(last_separator.unwrap_or(0));
            continue;
        }
        if !normalized.is_empty() {
            normalized
                .try_reserve(1)
                .map_err(|source| Error::from(ErrorKind::Allocation { resource, source }))?;
            normalized.push('/');
        }
        normalized
            .try_reserve(component.len())
            .map_err(|source| Error::from(ErrorKind::Allocation { resource, source }))?;
        normalized.push_str(component);
    }
    Ok(normalized)
}

fn lookup_member_name(name: &str) -> Result<LookupMemberName, Error> {
    let explicit_directory = name
        .as_bytes()
        .last()
        .is_some_and(|byte| matches!(byte, b'/' | b'\\'));
    Ok(LookupMemberName {
        name: canonical_member_name(normalize_str_fallibly(name, "archive lookup name")?),
        explicit_directory,
    })
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
    let raw_name = std::str::from_utf8(raw).map_err(|_| ErrorKind::InvalidInput {
        msg: "strict package archive contains a non-UTF-8 member name".to_string(),
    })?;
    let normalized_name = normalize_str_fallibly(raw_name, "strict package member name")?;
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
    Ok(normalized_name)
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
    let variable_length = usize::from(local.file_name_len)
        .checked_add(usize::from(local.extra_field_len))
        .ok_or_else(|| Error::from(ErrorKind::Eof))?;
    let mut variable = Vec::new();
    variable
        .try_reserve_exact(variable_length)
        .map_err(|source| {
            Error::from(ErrorKind::Allocation {
                resource: "strict ZIP mimetype local metadata",
                source,
            })
        })?;
    variable.resize(variable_length, 0);
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

fn clone_name_fallibly(name: &[u8], resource: &'static str) -> Result<Vec<u8>, Error> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(name.len())
        .map_err(|source| Error::from(ErrorKind::Allocation { resource, source }))?;
    owned.extend_from_slice(name);
    Ok(owned)
}

fn clone_str_fallibly(value: &str, resource: &'static str) -> Result<String, Error> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::from(ErrorKind::Allocation { resource, source }))?;
    owned.push_str(value);
    Ok(owned)
}

fn lossy_string_fallibly(bytes: &[u8], resource: &'static str) -> Result<String, Error> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|source| Error::from(ErrorKind::Allocation { resource, source }))?;
    let mut remaining = bytes;
    while !remaining.is_empty() {
        match std::str::from_utf8(remaining) {
            Ok(valid) => {
                owned.push_str(valid);
                break;
            },
            Err(error) => {
                let valid_up_to = error.valid_up_to();
                if valid_up_to != 0 {
                    let valid = remaining.get(..valid_up_to).ok_or_else(|| {
                        Error::from(ErrorKind::InvalidInput {
                            msg: "invalid UTF-8 replacement range".to_string(),
                        })
                    })?;
                    let valid = std::str::from_utf8(valid).map_err(|_| {
                        Error::from(ErrorKind::InvalidInput {
                            msg: "invalid UTF-8 replacement range".to_string(),
                        })
                    })?;
                    owned.push_str(valid);
                }
                let Some(invalid_len) = error.error_len() else {
                    break;
                };
                let skip = valid_up_to.checked_add(invalid_len).ok_or_else(|| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: "invalid UTF-8 replacement range overflows usize".to_string(),
                    })
                })?;
                owned
                    .try_reserve(3)
                    .map_err(|source| Error::from(ErrorKind::Allocation { resource, source }))?;
                owned.push('\u{fffd}');
                remaining = remaining.get(skip..).ok_or_else(|| {
                    Error::from(ErrorKind::InvalidInput {
                        msg: "invalid UTF-8 replacement range".to_string(),
                    })
                })?;
            },
        }
    }
    Ok(owned)
}

fn collect_order_fallibly<T>(
    ordered: Vec<(u64, T)>,
    resource: &'static str,
) -> Result<Vec<T>, Error> {
    let mut order = Vec::new();
    order
        .try_reserve_exact(ordered.len())
        .map_err(|source| Error::from(ErrorKind::Allocation { resource, source }))?;
    for (_, item) in ordered {
        order.push(item);
    }
    Ok(order)
}

const CENTRAL_FIXED_RECORD_BYTES: u64 = 46;
fn physical_entry_bound(
    central_directory_size: u64,
    entry_count: u64,
    max_metadata_bytes: u64,
) -> Result<usize, Error> {
    if entry_count > central_directory_size / CENTRAL_FIXED_RECORD_BYTES {
        return Err(Error::from(ErrorKind::InvalidInput {
            msg: "EOCD entry count exceeds declared central-directory capacity".to_string(),
        }));
    }
    let bound = entry_count.min(max_metadata_bytes / CENTRAL_FIXED_RECORD_BYTES);
    usize::try_from(bound).map_err(|_| {
        Error::from(ErrorKind::InvalidInput {
            msg: "ZIP physical entry bound does not fit this platform".to_string(),
        })
    })
}

/// The streaming writer prepares one reusable Deflate state before opening a
/// Deflate member. Reaching a Deflate branch without it is a writer bug, not a
/// caller error, so it is refused rather than papered over with a fresh state.
fn missing_deflate_state() -> Error {
    Error::from(ErrorKind::InvalidInput {
        msg: "streaming Deflate state was not prepared".to_string(),
    })
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

/// Copy a verified member to a sink using bounded scratch space while
/// recording destination acceptance. The one-byte probe after the declared
/// size detects an overlong logical stream without publishing bytes beyond
/// the central-directory claim. CRC verification is performed only after the
/// complete declared payload has been accepted by the sink. Deflate calls also
/// record decoder output; Store calls only record logical payload acceptance.
fn stream_verified_with_accounting<D, W>(
    mut reader: D,
    verifier: ZipVerification,
    sink: &mut W,
    accounting: &mut ZipOperationAccounting,
    accounting_kind: AccountingReadKind,
) -> Result<u64, Error>
where
    D: Read,
    W: Write,
{
    let expected_size = verifier.size();
    let mut copied = 0_u64;
    let mut crc = 0_u32;
    let mut buffer = [0u8; STREAM_COPY_BUFFER_SIZE];

    while copied < expected_size {
        let remaining = expected_size - copied;
        let request = usize::try_from(remaining)
            .unwrap_or(STREAM_COPY_BUFFER_SIZE)
            .min(buffer.len());
        let read = loop {
            match reader.read(&mut buffer[..request]) {
                Ok(read) => break validate_read_count(read, request)?,
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(error) => return Err(error.into()),
            }
        };
        if read == 0 {
            return Err(ErrorKind::InvalidSize {
                expected: expected_size,
                actual: copied,
            }
            .into());
        }
        if matches!(accounting_kind, AccountingReadKind::Deflate) {
            accounting.add_deflate_bytes_produced(usize_to_u64(
                read,
                "decompressed Deflate bytes produced",
            )?)?;
        }

        let mut accepted = 0;
        while accepted < read {
            let written = loop {
                match sink.write(&buffer[accepted..read]) {
                    Ok(written) => break written,
                    Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
                    Err(error) => return Err(error.into()),
                }
            };
            if written == 0 {
                return Err(Error::from(std::io::Error::new(
                    std::io::ErrorKind::WriteZero,
                    "failed to write decompressed ZIP output",
                )));
            }
            if written > read - accepted {
                return Err(ErrorKind::InvalidInput {
                    msg: "ZIP output sink returned more bytes than requested".to_string(),
                }
                .into());
            }
            let written_u64 = usize_to_u64(written, "decompressed ZIP bytes accepted")?;
            match accounting_kind {
                AccountingReadKind::Stored => {
                    accounting.add_stored_payload_bytes_accepted(written_u64)?;
                },
                AccountingReadKind::Deflate => {
                    accounting.add_deflate_bytes_accepted(written_u64)?;
                },
            }
            accepted = accepted
                .checked_add(written)
                .ok_or_else(|| accounting_overflow("decompressed ZIP bytes accepted"))?;
        }
        crc = crc32_chunk(&buffer[..read], crc);
        copied = copied
            .checked_add(usize_to_u64(read, "ZIP logical payload bytes copied")?)
            .ok_or_else(|| accounting_overflow("ZIP logical payload bytes copied"))?;
    }

    let mut probe = [0_u8; 1];
    let extra = loop {
        match reader.read(&mut probe) {
            Ok(extra) => break validate_read_count(extra, probe.len())?,
            Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
            Err(error) => return Err(error.into()),
        }
    };
    if extra != 0 {
        if matches!(accounting_kind, AccountingReadKind::Deflate) {
            if let Ok(extra) = usize_to_u64(extra, "decompressed Deflate bytes produced") {
                let _ = accounting.add_deflate_bytes_produced(extra);
            }
        }
        let actual = copied
            .checked_add(usize_to_u64(extra, "ZIP logical payload bytes probed")?)
            .ok_or_else(|| accounting_overflow("ZIP logical payload bytes probed"))?;
        return Err(ErrorKind::InvalidSize {
            expected: expected_size,
            actual,
        }
        .into());
    }

    verifier.valid_strict(ZipVerification {
        crc,
        uncompressed_size: copied,
    })?;
    Ok(copied)
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
/// The writer uses ZIP32 framing while all configured payload ceilings remain
/// below ZIP64 thresholds, and promotes unknown-size entries to ZIP64 before
/// their local header when an admitted entry or compressed-size ceiling can
/// reach a ZIP32 sentinel. The output ceiling, per-entry ceiling, aggregate
/// uncompressed ceiling, member count, and metadata ceiling are finite by
/// default and are checked before an entry starts or before an input chunk is
/// accepted.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingArchiveLimits {
    /// Maximum number of file members accepted by a writer.
    pub max_entries: usize,
    /// Maximum UTF-8 member-name bytes for one file member.
    pub max_member_name_bytes: u64,
    /// Maximum aggregate variable central-directory metadata bytes.
    ///
    /// This includes normalized member names and generated ZIP64 extra-field
    /// bytes. ZIP fixed-size headers are not included, matching
    /// [`ArchiveLimits::max_metadata_bytes`].
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
const ZIP32_MAX_MEMBER_NAME_BYTES: u64 = u16::MAX as u64;
const ZIP64_EXTRA_FIELD_HEADER_BYTES: u64 = 4;
const ZIP64_EXTRA_FIELD_VALUE_BYTES: u64 = 8;
const MIN_STREAM_OUTPUT_BYTES: u64 = 22;
const DEFAULT_STREAM_MAX_ENTRIES: usize = 65_534;
const DEFAULT_STREAM_MAX_COMPRESSED_SIZE: u64 = 512 * 1024 * 1024;
const DEFAULT_STREAM_MAX_ENTRY_SIZE: u64 = 512 * 1024 * 1024;
const DEFAULT_STREAM_MAX_TOTAL_SIZE: u64 = 2 * 1024 * 1024 * 1024;
const DEFAULT_STREAM_MAX_OUTPUT_BYTES: u64 = 512 * 1024 * 1024;
const STREAM_COPY_BUFFER_SIZE: usize = 16 * 1024;

/// Admission state retained until a generated member has been finalized.
///
/// `metadata_bytes` includes the normalized member name and any ZIP64 extra
/// fields that the low-level writer will generate for the central record.
#[derive(Debug)]
struct StreamingEntryAdmission {
    normalized_name: String,
    metadata_bytes: u64,
    next_entries: usize,
    next_metadata_bytes: u64,
}

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
        let requested = usize_to_u64(buffer.len(), "streaming ZIP output bytes requested")
            .map_err(|error| std::io::Error::new(std::io::ErrorKind::Other, error))?;
        let attempted = self.accepted.checked_add(requested).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "streaming ZIP output byte count overflow",
            )
        })?;
        if attempted > self.maximum {
            return Err(streaming_limit_io_error(
                StreamingLimitResource::OutputBytes,
                attempted,
                self.maximum,
            ));
        }
        let written = self.writer.write(buffer)?;
        if written > buffer.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "ZIP output sink returned more bytes than requested",
            ));
        }
        let written_u64 = usize_to_u64(written, "streaming ZIP output bytes accepted")
            .map_err(|error| std::io::Error::new(std::io::ErrorKind::Other, error))?;
        self.accepted = self.accepted.checked_add(written_u64).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "streaming ZIP output byte count overflow",
            )
        })?;
        self.counter.store(self.accepted, Ordering::Release);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.writer.flush()
    }
}

/// A ZIP entry sink that bounds compressed payload bytes before forwarding
/// them to the archive writer.
struct LimitedEntryWriter<'entry, 'archive, 'accounting, W> {
    inner: &'entry mut crate::ZipEntryWriter<'archive, BoundedOutput<W>>,
    maximum: u64,
    accounting: &'accounting mut ZipOperationAccounting,
    accounting_kind: AccountingWriteKind,
}

impl<'entry, 'archive, 'accounting, W> LimitedEntryWriter<'entry, 'archive, 'accounting, W> {
    fn new(
        inner: &'entry mut crate::ZipEntryWriter<'archive, BoundedOutput<W>>,
        maximum: u64,
        accounting: &'accounting mut ZipOperationAccounting,
        accounting_kind: AccountingWriteKind,
    ) -> Self {
        Self {
            inner,
            maximum,
            accounting,
            accounting_kind,
        }
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

impl<W: Write> Write for LimitedEntryWriter<'_, '_, '_, W> {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let requested = usize_to_u64(buffer.len(), "generated ZIP payload bytes requested")
            .map_err(|error| std::io::Error::new(std::io::ErrorKind::Other, error))?;
        let compressed = self.compressed_bytes();
        let attempted = compressed.checked_add(requested).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "generated ZIP payload byte count overflow",
            )
        })?;
        if attempted > self.maximum {
            return Err(streaming_limit_io_error(
                StreamingLimitResource::CompressedBytes,
                attempted,
                self.maximum,
            ));
        }
        let written = self.inner.write(buffer)?;
        if written != 0 {
            self.accounting_kind
                .add(
                    self.accounting,
                    usize_to_u64(written, "generated ZIP payload bytes accepted")
                        .map_err(|error| std::io::Error::new(std::io::ErrorKind::Other, error))?,
                )
                .map_err(|error| std::io::Error::new(std::io::ErrorKind::Other, error))?;
        }
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

/// Bounded in-memory scratch for one deflated member.
///
/// The scratch enforces the same compressed-size budget as
/// [`LimitedEntryWriter`], but before any archive byte is emitted so a limit
/// breach cannot poison the archive mid-entry.
struct CompressedScratch<'a> {
    inner: &'a mut Vec<u8>,
    maximum: u64,
}

/// Name admission policy used by the Office streaming writer.
///
/// The ordinary policy retains normalized names so compatibility callers keep
/// the existing duplicate-name checks.  A generated policy owns a bounded
/// cursor instead; its descriptor proves the next normalized name and is
/// advanced only after the low-level ZIP entry has been published.
enum StreamingNamePolicy {
    Ordinary(HashSet<String>),
    Generated(Box<crate::generated_names::GeneratedNamePlan>),
}

impl StreamingNamePolicy {
    fn check_next(&self, normalized_name: &str) -> Result<(), Error> {
        match self {
            Self::Ordinary(names) => {
                if names.contains(normalized_name) {
                    return Err(ErrorKind::InvalidInput {
                        msg: format!("duplicate normalized member name: {normalized_name}"),
                    }
                    .into());
                }
                Ok(())
            },
            Self::Generated(plan) => plan.check_next(normalized_name),
        }
    }

    fn reserve(&mut self) -> Result<(), Error> {
        match self {
            Self::Ordinary(names) => names.try_reserve(1).map_err(|error| {
                ErrorKind::InvalidInput {
                    msg: format!("could not reserve streaming ZIP member-name index: {error}"),
                }
                .into()
            }),
            Self::Generated(_) => Ok(()),
        }
    }

    fn record(&mut self, normalized_name: String) -> Result<(), Error> {
        match self {
            Self::Ordinary(names) => {
                let inserted = names.insert(normalized_name);
                debug_assert!(inserted);
                Ok(())
            },
            Self::Generated(plan) => {
                drop(normalized_name);
                plan.advance()
            },
        }
    }

    fn is_complete(&self) -> bool {
        match self {
            Self::Ordinary(_) => true,
            Self::Generated(plan) => plan.is_complete(),
        }
    }
}

impl<'a> CompressedScratch<'a> {
    fn new(inner: &'a mut Vec<u8>, maximum: u64) -> Self {
        Self { inner, maximum }
    }
}

impl Write for CompressedScratch<'_> {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let written = usize_to_u64(self.inner.len(), "compressed scratch bytes")
            .map_err(|error| std::io::Error::new(std::io::ErrorKind::Other, error))?;
        let requested = usize_to_u64(buffer.len(), "compressed scratch bytes requested")
            .map_err(|error| std::io::Error::new(std::io::ErrorKind::Other, error))?;
        let attempted = written.checked_add(requested).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "compressed scratch byte count overflow",
            )
        })?;
        if attempted > self.maximum {
            return Err(streaming_limit_io_error(
                StreamingLimitResource::CompressedBytes,
                attempted,
                self.maximum,
            ));
        }
        self.inner.write(buffer)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
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
    name_policy: StreamingNamePolicy,
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
    metadata_charge: u64,
    total_uncompressed_bytes: u64,
    name_policy: StreamingNamePolicy,
    normalized_name: String,
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

        let next_entries = match self.entries.checked_add(1) {
            Some(entries) => entries,
            None => {
                self.poisoned = true;
                self.failure = Some(accounting_overflow("streaming ZIP entry count"));
                return Err(self.failure());
            },
        };
        let next_metadata_bytes = match self.metadata_bytes.checked_add(self.metadata_charge) {
            Some(metadata_bytes) => metadata_bytes,
            None => {
                self.poisoned = true;
                self.failure = Some(accounting_overflow(
                    "streaming ZIP aggregate metadata bytes",
                ));
                return Err(self.failure());
            },
        };
        let next_total_uncompressed_bytes = match self
            .total_uncompressed_bytes
            .checked_add(self.uncompressed_bytes)
        {
            Some(total) => total,
            None => {
                self.poisoned = true;
                self.failure = Some(accounting_overflow(
                    "streaming ZIP aggregate uncompressed bytes",
                ));
                return Err(self.failure());
            },
        };

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
            entries: next_entries,
            metadata_bytes: next_metadata_bytes,
            total_uncompressed_bytes: next_total_uncompressed_bytes,
            output_bytes: self.output_bytes,
            poisoned: false,
            name_policy: self.name_policy,
            last_limit: None,
            output_counter: self.output_counter,
        };
        if let Err(error) = writer.name_policy.record(self.normalized_name) {
            let error = writer.poison(error);
            return Err(StreamingArchiveFailure {
                error,
                progress: writer.progress(),
                limit: writer.last_limit,
            });
        }
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

        let requested = match usize_to_u64(buffer.len(), "streaming ZIP entry bytes requested") {
            Ok(requested) => requested,
            Err(error) => {
                return Err(self.capture_io_failure(std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    error.to_string(),
                )));
            },
        };
        let next_entry = match self.uncompressed_bytes.checked_add(requested) {
            Some(next_entry) => next_entry,
            None => {
                return Err(self.capture_io_failure(std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    accounting_overflow("streaming ZIP entry uncompressed bytes").to_string(),
                )));
            },
        };
        if next_entry > self.limits.max_entry_size {
            let error = streaming_payload_limit_io_error(
                LimitResource::EntrySize,
                next_entry,
                self.limits.max_entry_size,
            );
            return Err(self.capture_io_failure(error));
        }
        let next_total = match self.total_uncompressed_bytes.checked_add(next_entry) {
            Some(next_total) => next_total,
            None => {
                return Err(self.capture_io_failure(std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    accounting_overflow("streaming ZIP aggregate uncompressed bytes").to_string(),
                )));
            },
        };
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
                let written_u64 = match usize_to_u64(written, "streaming ZIP entry bytes accepted")
                {
                    Ok(written) => written,
                    Err(error) => {
                        return Err(self.capture_io_failure(std::io::Error::new(
                            std::io::ErrorKind::InvalidInput,
                            error.to_string(),
                        )));
                    },
                };
                self.uncompressed_bytes = match self.uncompressed_bytes.checked_add(written_u64) {
                    Some(uncompressed_bytes) => uncompressed_bytes,
                    None => {
                        return Err(self.capture_io_failure(std::io::Error::new(
                            std::io::ErrorKind::InvalidInput,
                            accounting_overflow("streaming ZIP entry uncompressed bytes")
                                .to_string(),
                        )));
                    },
                };
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
            name_policy: StreamingNamePolicy::Ordinary(HashSet::new()),
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
            name_policy: StreamingNamePolicy::Ordinary(HashSet::new()),
            last_limit: None,
            output_counter,
        }
    }

    /// Create a sequential writer with explicit central-directory scratch.
    ///
    /// Finalized central records are appended to the caller-supplied store
    /// and replayed through the configured buffer at archive finalization.
    /// The output sink remains sequential and need not implement `Seek`.
    /// Scratch is never opened implicitly; the caller chooses its storage,
    /// confidentiality, and cleanup policy. A memory-backed store retains its
    /// contents in memory.
    ///
    /// Working storage includes the replay buffer and one active record/name
    /// bounded by ZIP field limits. This does not bound the complete writer:
    /// normalized names are still retained for exact duplicate checks
    /// under `limits`. Spool failures after output begins are reported by the
    /// existing progress-bearing entry and archive finalization methods.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid limits or failure to initialize scratch.
    /// Construction does not write to the output sink.
    pub fn with_writer_and_limits_and_spool<S>(
        writer: W,
        limits: StreamingArchiveLimits,
        spool: S,
        spool_limits: crate::DirectorySpoolLimits,
    ) -> Result<Self, Error>
    where
        S: Read + Write + std::io::Seek + Send + Sync + 'static,
    {
        let output_counter = Arc::new(AtomicU64::new(0));
        let archive = ZipArchiveWriter::builder().build_with_spool(
            BoundedOutput::new(writer, limits.max_output_bytes, Arc::clone(&output_counter)),
            spool,
            spool_limits,
        )?;
        let result = Self {
            archive,
            limits,
            entries: 0,
            metadata_bytes: 0,
            total_uncompressed_bytes: 0,
            output_bytes: 0,
            poisoned: false,
            name_policy: StreamingNamePolicy::Ordinary(HashSet::new()),
            last_limit: None,
            output_counter,
        };
        result.ensure_usable()?;
        Ok(result)
    }

    /// Create a sequential writer whose member names must follow a checked
    /// generated-name plan.
    ///
    /// The plan is consumed by the writer and carried through each owned
    /// entry. Each caller-supplied name must exactly match the plan's next
    /// canonical raw name before any trimming or ZIP path normalization. The
    /// plan advances only after that member's central record has been
    /// published. Unlike the ordinary writer, this mode does not retain a
    /// growing normalized-name set.
    pub fn with_writer_and_limits_and_spool_and_name_plan<S>(
        writer: W,
        limits: StreamingArchiveLimits,
        spool: S,
        spool_limits: crate::DirectorySpoolLimits,
        plan: crate::generated_names::GeneratedNamePlan,
    ) -> Result<Self, Error>
    where
        S: Read + Write + std::io::Seek + Send + Sync + 'static,
    {
        if let Some(error) = Self::invalid_limits_for(limits) {
            return Err(error);
        }
        let maximum_entries = usize_to_u64(limits.max_entries, "streaming ZIP entry limit")?;
        if plan.entry_count() > maximum_entries {
            return Err(limit_error(
                LimitResource::FileCount,
                plan.entry_count(),
                maximum_entries,
            ));
        }
        let plan_name_bytes = usize_to_u64(
            plan.max_name_bytes(),
            "generated streaming ZIP member name bytes",
        )?;
        let maximum_name_bytes = limits
            .max_member_name_bytes
            .min(ZIP32_MAX_MEMBER_NAME_BYTES);
        if plan_name_bytes > maximum_name_bytes {
            return Err(limit_error(
                LimitResource::MemberNameBytes,
                plan_name_bytes,
                maximum_name_bytes,
            ));
        }
        let output_counter = Arc::new(AtomicU64::new(0));
        let archive = ZipArchiveWriter::builder().build_with_spool(
            BoundedOutput::new(writer, limits.max_output_bytes, Arc::clone(&output_counter)),
            spool,
            spool_limits,
        )?;
        let result = Self {
            archive,
            limits,
            entries: 0,
            metadata_bytes: 0,
            total_uncompressed_bytes: 0,
            output_bytes: 0,
            poisoned: false,
            name_policy: StreamingNamePolicy::Generated(Box::new(plan)),
            last_limit: None,
            output_counter,
        };
        result.ensure_usable()?;
        Ok(result)
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
    /// for finalized members, including generated ZIP64 extra fields.
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

    fn invalid_limits_for(limits: StreamingArchiveLimits) -> Option<Error> {
        let reason = if limits.max_member_name_bytes > ZIP32_MAX_MEMBER_NAME_BYTES {
            "max_member_name_bytes exceeds the ZIP32 member-name field"
        } else if limits.max_metadata_bytes > limits.max_output_bytes {
            "max_metadata_bytes exceeds max_output_bytes"
        } else if limits.max_compressed_size > limits.max_output_bytes {
            "max_compressed_size exceeds max_output_bytes"
        } else if limits.max_output_bytes < MIN_STREAM_OUTPUT_BYTES {
            "max_output_bytes is too small for an empty ZIP archive"
        } else {
            return None;
        };
        Some(ErrorKind::InvalidInput { msg: reason.into() }.into())
    }

    fn invalid_limits(&self) -> Option<Error> {
        Self::invalid_limits_for(self.limits)
    }

    /// Whether an unknown-size entry must establish ZIP64 framing before its
    /// local header is emitted.
    ///
    /// The byte ceilings are the admission contract for streamed payloads.
    /// When either payload ceiling can reach the ZIP32 sentinel, a later
    /// chunk may require a 64-bit descriptor even though its final size is not
    /// known at entry start. Count, aggregate, and output boundaries are
    /// handled by the archive writer's ZIP64 tail promotion and do not change
    /// the descriptor width for an entry whose own size remains ZIP32-sized.
    #[inline]
    fn streaming_entry_uses_zip64(&self) -> bool {
        self.limits.max_compressed_size >= ZIP32_MAX_VALUE
            || self.limits.max_entry_size >= ZIP32_MAX_VALUE
    }

    /// Number of bytes the low-level writer will add to one central record's
    /// ZIP64 extra field for the supplied size/offset knowledge.
    #[inline]
    fn generated_central_zip64_extra_bytes(
        force_size_fields: bool,
        compressed_size: Option<u64>,
        uncompressed_size: Option<u64>,
        local_header_offset: u64,
    ) -> u64 {
        let has_uncompressed_size =
            force_size_fields || uncompressed_size.is_some_and(|size| size >= ZIP32_MAX_VALUE);
        let has_compressed_size =
            force_size_fields || compressed_size.is_some_and(|size| size >= ZIP32_MAX_VALUE);
        let has_offset = local_header_offset >= ZIP32_MAX_VALUE;

        let field_count =
            has_uncompressed_size as u64 + has_compressed_size as u64 + has_offset as u64;
        if field_count == 0 {
            0
        } else {
            ZIP64_EXTRA_FIELD_HEADER_BYTES + field_count * ZIP64_EXTRA_FIELD_VALUE_BYTES
        }
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

    fn validate_entry_name(&self, name: &str) -> Result<(String, u64, usize), Error> {
        self.ensure_usable()?;
        let generated_name = match &self.name_policy {
            StreamingNamePolicy::Ordinary(_) => false,
            StreamingNamePolicy::Generated(plan) => {
                // Generated plans describe the exact public name sequence.
                // Check the caller's raw spelling before the compatibility
                // path normalization below can erase a mismatch.
                plan.check_next(name)?;
                true
            },
        };
        let raw_name = name.trim_end_matches('/');
        let raw_name_bytes = usize_to_u64(raw_name.len(), "streaming ZIP member name bytes")?;
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
        let name_bytes = usize_to_u64(normalized_name.len(), "normalized ZIP member name bytes")?;
        if name_bytes > ZIP32_MAX_MEMBER_NAME_BYTES
            || name_bytes > self.limits.max_member_name_bytes
        {
            return Err(limit_error(
                LimitResource::MemberNameBytes,
                name_bytes,
                maximum_name_bytes,
            ));
        }

        if !generated_name {
            self.name_policy.check_next(&normalized_name)?;
        }

        let next_entries = self
            .entries
            .checked_add(1)
            .ok_or_else(|| accounting_overflow("streaming ZIP entry count"))?;
        let max_entries = usize_to_u64(self.limits.max_entries, "streaming ZIP entry limit")?;
        let actual_entries = usize_to_u64(next_entries, "streaming ZIP entry count")?;
        if next_entries > self.limits.max_entries {
            return Err(limit_error(
                LimitResource::FileCount,
                actual_entries,
                max_entries,
            ));
        }

        Ok((normalized_name, name_bytes, next_entries))
    }

    fn complete_entry_admission(
        &self,
        normalized_name: String,
        name_bytes: u64,
        next_entries: usize,
        generated_central_extra_bytes: u64,
    ) -> Result<StreamingEntryAdmission, Error> {
        let metadata_bytes = name_bytes
            .checked_add(generated_central_extra_bytes)
            .ok_or_else(|| accounting_overflow("streaming ZIP member metadata bytes"))?;
        let next_metadata = self
            .metadata_bytes
            .checked_add(metadata_bytes)
            .ok_or_else(|| accounting_overflow("streaming ZIP aggregate metadata bytes"))?;
        if next_metadata > self.limits.max_metadata_bytes {
            return Err(limit_error(
                LimitResource::MetadataBytes,
                next_metadata,
                self.limits.max_metadata_bytes,
            ));
        }

        Ok(StreamingEntryAdmission {
            normalized_name,
            metadata_bytes,
            next_entries,
            next_metadata_bytes: next_metadata,
        })
    }

    fn validate_streaming_entry(&self, name: &str) -> Result<StreamingEntryAdmission, Error> {
        let (normalized_name, name_bytes, next_entries) = self.validate_entry_name(name)?;
        let generated_central_extra_bytes = Self::generated_central_zip64_extra_bytes(
            self.streaming_entry_uses_zip64(),
            None,
            None,
            self.archive.stream_offset(),
        );
        self.complete_entry_admission(
            normalized_name,
            name_bytes,
            next_entries,
            generated_central_extra_bytes,
        )
    }

    fn validate_known_payload(&self, uncompressed_bytes: u64) -> Result<(), Error> {
        if uncompressed_bytes > self.limits.max_entry_size {
            return Err(limit_error(
                LimitResource::EntrySize,
                uncompressed_bytes,
                self.limits.max_entry_size,
            ));
        }
        let total = self
            .total_uncompressed_bytes
            .checked_add(uncompressed_bytes)
            .ok_or_else(|| accounting_overflow("streaming ZIP aggregate uncompressed bytes"))?;
        if total > self.limits.max_total_size {
            return Err(limit_error(
                LimitResource::TotalSize,
                total,
                self.limits.max_total_size,
            ));
        }
        Ok(())
    }

    fn validate_known_entry(
        &self,
        name: &str,
        uncompressed_bytes: u64,
        compressed_bytes: u64,
    ) -> Result<StreamingEntryAdmission, Error> {
        let (normalized_name, name_bytes, next_entries) = self.validate_entry_name(name)?;
        self.validate_known_payload(uncompressed_bytes)?;
        let generated_central_extra_bytes = Self::generated_central_zip64_extra_bytes(
            false,
            Some(compressed_bytes),
            Some(uncompressed_bytes),
            self.archive.stream_offset(),
        );
        self.complete_entry_admission(
            normalized_name,
            name_bytes,
            next_entries,
            generated_central_extra_bytes,
        )
    }

    fn record_streaming_entry(&mut self, admission: StreamingEntryAdmission) -> Result<(), Error> {
        self.name_policy.record(admission.normalized_name)?;
        self.entries = admission.next_entries;
        self.metadata_bytes = admission.next_metadata_bytes;
        Ok(())
    }

    fn reserve_streaming_entry(&mut self) -> Result<(), Error> {
        self.name_policy.reserve()
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
            let read_bytes = usize_to_u64(read, "stream source bytes read")?;
            let next_entry = accepted_uncompressed
                .checked_add(read_bytes)
                .ok_or_else(|| accounting_overflow("streaming ZIP entry uncompressed bytes"))?;
            if next_entry > max_entry_size {
                return Err(limit_error(
                    LimitResource::EntrySize,
                    next_entry,
                    max_entry_size,
                ));
            }
            let next_total = committed_total
                .checked_add(accepted_uncompressed)
                .and_then(|total| total.checked_add(read_bytes))
                .ok_or_else(|| accounting_overflow("streaming ZIP aggregate uncompressed bytes"))?;
            if next_total > max_total_size {
                return Err(limit_error(
                    LimitResource::TotalSize,
                    next_total,
                    max_total_size,
                ));
            }
            output.write_all(&buffer[..read])?;
            accepted_uncompressed = next_entry;
        }
    }

    fn write_reader<R: Read>(
        &mut self,
        name: &str,
        reader: R,
        compression_method: CompressionMethod,
    ) -> Result<(), Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.write_reader_with_accounting(name, reader, compression_method, &mut accounting)
    }

    fn write_reader_with_accounting<R: Read>(
        &mut self,
        name: &str,
        mut reader: R,
        compression_method: CompressionMethod,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        let admission = self.validate_streaming_entry(name)?;
        self.reserve_streaming_entry()?;
        let zip64 = self.streaming_entry_uses_zip64();
        // One compressor serves every Deflate member of this archive: taken
        // before the entry borrows the archive, returned only after the
        // member's final Deflate output succeeded.
        let mut deflate_state = (compression_method == CompressionMethod::Deflate)
            .then(|| self.archive.take_reusable_deflate());
        let started = self
            .archive
            .new_file(&admission.normalized_name)
            .compression_method(compression_method)
            .zip64(zip64)
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
                    let mut limited_entry = LimitedEntryWriter::new(
                        &mut entry,
                        max_compressed_size,
                        accounting,
                        AccountingWriteKind::Stored,
                    );
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
                    let mut limited_entry = LimitedEntryWriter::new(
                        &mut entry,
                        max_compressed_size,
                        accounting,
                        AccountingWriteKind::GeneratedDeflate,
                    );
                    let state = match deflate_state.as_deref_mut() {
                        Some(state) => state,
                        None => return Err(missing_deflate_state()),
                    };
                    let encoder = ReusedDeflateEncoder::new(state, &mut limited_entry);
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
        if let Some(state) = deflate_state {
            self.archive.restore_reusable_deflate(state);
        }
        self.total_uncompressed_bytes = match self
            .total_uncompressed_bytes
            .checked_add(accepted_uncompressed)
        {
            Some(total) => total,
            None => {
                return Err(self.poison(accounting_overflow(
                    "streaming ZIP aggregate uncompressed bytes",
                )));
            },
        };
        if let Err(error) = self.record_streaming_entry(admission) {
            return Err(self.poison(error));
        }
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
        let admission = match self.validate_streaming_entry(name) {
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

        let zip64 = self.streaming_entry_uses_zip64();

        let StreamingArchiveWriter {
            archive,
            limits,
            entries,
            metadata_bytes,
            total_uncompressed_bytes,
            output_bytes,
            poisoned: _,
            name_policy,
            last_limit,
            output_counter,
        } = self;
        let entry = match (if zip64 {
            archive.start_file_owned_zip64(&admission.normalized_name, compression_method)
        } else {
            archive.start_file_owned(&admission.normalized_name, compression_method)
        })
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
            metadata_charge: admission.metadata_bytes,
            total_uncompressed_bytes,
            name_policy,
            normalized_name: admission.normalized_name,
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
        let mut accounting = ZipOperationAccounting::default();
        self.write_stored_with_accounting(name, data, &mut accounting)
    }

    /// Write a stored file while recording payload bytes accepted by the
    /// archive sink. ZIP framing bytes are excluded from the counter.
    pub fn write_stored_with_accounting(
        &mut self,
        name: &str,
        data: &[u8],
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        let data_bytes = usize_to_u64(data.len(), "stored payload length")?;
        let admission = self.validate_known_entry(name, data_bytes, data_bytes)?;
        self.reserve_streaming_entry()?;
        if data_bytes > self.limits.max_compressed_size {
            return Err(limit_error(
                LimitResource::CompressedSize,
                data_bytes,
                self.limits.max_compressed_size,
            ));
        }
        match self.archive.write_stored_file_with_accounting(
            &admission.normalized_name,
            data,
            accounting,
        ) {
            Ok(()) => {
                self.total_uncompressed_bytes =
                    match self.total_uncompressed_bytes.checked_add(data_bytes) {
                        Some(total) => total,
                        None => {
                            return Err(self.poison(accounting_overflow(
                                "streaming ZIP aggregate uncompressed bytes",
                            )));
                        },
                    };
                if let Err(error) = self.record_streaming_entry(admission) {
                    return Err(self.poison(error));
                }
                self.refresh_output_bytes();
                Ok(())
            },
            Err(error) => Err(self.poison(error)),
        }
    }

    /// Write a file with Deflate compression.
    ///
    /// The member is written through the streaming entry API, so the local
    /// header carries a data descriptor. Callers that already hold the full
    /// payload and need a canonical local header with upfront CRC-32 and
    /// sizes can use [`Self::write_deflated_sized`] instead.
    pub fn write_deflated(&mut self, name: &str, data: &[u8]) -> Result<(), Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.write_deflated_with_accounting(name, data, &mut accounting)
    }

    /// Write a Deflate-compressed file while recording generated payload bytes
    /// accepted by the archive sink. ZIP framing bytes are excluded.
    pub fn write_deflated_with_accounting(
        &mut self,
        name: &str,
        data: &[u8],
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        let data_bytes = usize_to_u64(data.len(), "Deflate payload length")?;
        let admission = self.validate_streaming_entry(name)?;
        self.validate_known_payload(data_bytes)?;
        self.reserve_streaming_entry()?;
        let zip64 = self.streaming_entry_uses_zip64();
        // One compressor serves every Deflate member of this archive: taken
        // before the entry borrows the archive, returned only after the
        // member's final Deflate output succeeded.
        let mut deflate_state = self.archive.take_reusable_deflate();
        let started = self
            .archive
            .new_file(&admission.normalized_name)
            .compression_method(CompressionMethod::Deflate)
            .zip64(zip64)
            .start();
        let result = match started {
            Ok((mut entry, config)) => (|| {
                let descriptor = {
                    let mut limited_entry = LimitedEntryWriter::new(
                        &mut entry,
                        self.limits.max_compressed_size,
                        accounting,
                        AccountingWriteKind::GeneratedDeflate,
                    );
                    let encoder = ReusedDeflateEncoder::new(&mut deflate_state, &mut limited_entry);
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
        self.archive.restore_reusable_deflate(deflate_state);
        self.total_uncompressed_bytes = match self.total_uncompressed_bytes.checked_add(data_bytes)
        {
            Some(total) => total,
            None => {
                return Err(self.poison(accounting_overflow(
                    "streaming ZIP aggregate uncompressed bytes",
                )));
            },
        };
        if let Err(error) = self.record_streaming_entry(admission) {
            return Err(self.poison(error));
        }
        self.refresh_output_bytes();
        Ok(())
    }

    /// Write an in-memory payload with Deflate compression and upfront sizes.
    ///
    /// The payload is compressed into a bounded scratch buffer before any
    /// archive byte is emitted, so the local header declares the final CRC-32
    /// and both sizes and no data descriptor is needed. A compressed-size
    /// limit breach is therefore reported before the archive changes, leaving
    /// the writer usable instead of poisoned mid-entry. Prefer this over
    /// [`Self::write_deflated`] when the produced archive should be probeable
    /// from its central directory alone.
    pub fn write_deflated_sized(&mut self, name: &str, data: &[u8]) -> Result<(), Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.write_deflated_sized_with_accounting(name, data, &mut accounting)
    }

    /// Write a sized Deflate member while recording generated compressed
    /// payload bytes accepted by the archive sink. ZIP framing bytes are
    /// excluded.
    pub fn write_deflated_sized_with_accounting(
        &mut self,
        name: &str,
        data: &[u8],
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        let data_bytes = usize_to_u64(data.len(), "Deflate payload length")?;
        let (normalized_name, name_bytes, next_entries) = self.validate_entry_name(name)?;
        self.validate_known_payload(data_bytes)?;
        self.reserve_streaming_entry()?;
        let mut compressed = Vec::new();
        compressed
            .try_reserve(
                data.len()
                    .min(usize::try_from(self.limits.max_compressed_size).unwrap_or(usize::MAX)),
            )
            .map_err(|source| {
                Error::from(ErrorKind::Allocation {
                    resource: "deflated member scratch buffer",
                    source,
                })
            })?;
        // One compressor serves every Deflate member of this archive. The
        // state returns to the archive only after this member's final Deflate
        // output succeeded; a refused member drops its unfinished stream and
        // the next member starts from a fresh one, so the writer stays usable.
        let mut deflate_state = self.archive.take_reusable_deflate();
        let compression = (|| {
            let mut scratch =
                CompressedScratch::new(&mut compressed, self.limits.max_compressed_size);
            let mut encoder = ReusedDeflateEncoder::new(&mut deflate_state, &mut scratch);
            encoder.write_all(data)?;
            encoder.finish().map(|_scratch| ())
        })();
        if let Err(error) = compression {
            let error = Error::from(error);
            return Err(match streaming_limit_from_error(&error) {
                Some(StreamingLimitExceeded {
                    resource: StreamingLimitResource::CompressedBytes,
                    actual,
                    maximum,
                }) => limit_error(LimitResource::CompressedSize, actual, maximum),
                _ => error,
            });
        }
        self.archive.restore_reusable_deflate(deflate_state);
        let compressed_bytes = usize_to_u64(compressed.len(), "compressed Deflate payload length")?;
        let metadata_extra = Self::generated_central_zip64_extra_bytes(
            false,
            Some(compressed_bytes),
            Some(data_bytes),
            self.archive.stream_offset(),
        );
        let admission = self.complete_entry_admission(
            normalized_name,
            name_bytes,
            next_entries,
            metadata_extra,
        )?;
        let crc32 = crate::crc32(data);
        match self.archive.write_generated_deflate_file_with_accounting(
            &admission.normalized_name,
            crc32,
            data_bytes,
            &compressed,
            accounting,
        ) {
            Ok(()) => {
                self.total_uncompressed_bytes =
                    match self.total_uncompressed_bytes.checked_add(data_bytes) {
                        Some(total) => total,
                        None => {
                            return Err(self.poison(accounting_overflow(
                                "streaming ZIP aggregate uncompressed bytes",
                            )));
                        },
                    };
                if let Err(error) = self.record_streaming_entry(admission) {
                    return Err(self.poison(error));
                }
                self.refresh_output_bytes();
                Ok(())
            },
            Err(error) => Err(self.poison(error)),
        }
    }

    /// Consume a reader value into a stored ZIP member.
    ///
    /// The source is read incrementally and is not retained after this method
    /// returns.  The output uses a data descriptor, so `W` only needs to
    /// implement [`Write`], not [`std::io::Seek`].
    pub fn write_stored_stream<R: Read>(&mut self, name: &str, reader: R) -> Result<(), Error> {
        self.write_reader(name, reader, CompressionMethod::Store)
    }

    /// Consume a reader into a stored ZIP member while recording stored
    /// payload bytes accepted by the archive sink.
    pub fn write_stored_stream_with_accounting<R: Read>(
        &mut self,
        name: &str,
        reader: R,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        self.write_reader_with_accounting(name, reader, CompressionMethod::Store, accounting)
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

    /// Consume a reader into a Deflate-compressed ZIP member while recording
    /// generated payload bytes accepted by the archive sink.
    pub fn write_deflated_stream_with_accounting<R: Read>(
        &mut self,
        name: &str,
        reader: R,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        self.write_reader_with_accounting(name, reader, CompressionMethod::Deflate, accounting)
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

    /// Consume a reader with one of the supported Office ZIP methods while
    /// recording payload bytes accepted by the archive sink.
    pub fn write_stream_with_accounting<R: Read>(
        &mut self,
        name: &str,
        reader: R,
        compression_method: CompressionMethod,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        if !matches!(
            compression_method,
            CompressionMethod::Store | CompressionMethod::Deflate
        ) {
            return Err(Error::from(ErrorKind::UnsupportedCompressionMethod(
                compression_method.as_id().as_u16(),
            )));
        }
        self.write_reader_with_accounting(name, reader, compression_method, accounting)
    }

    /// Alias for [`Self::write_stored_stream`] using reader-oriented naming.
    pub fn write_stored_reader<R: Read>(&mut self, name: &str, reader: R) -> Result<(), Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.write_stored_reader_with_accounting(name, reader, &mut accounting)
    }

    /// Alias for [`Self::write_stored_stream_with_accounting`].
    pub fn write_stored_reader_with_accounting<R: Read>(
        &mut self,
        name: &str,
        reader: R,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        self.write_stored_stream_with_accounting(name, reader, accounting)
    }

    /// Alias for [`Self::write_deflated_stream`] using reader-oriented naming.
    pub fn write_deflated_reader<R: Read>(&mut self, name: &str, reader: R) -> Result<(), Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.write_deflated_reader_with_accounting(name, reader, &mut accounting)
    }

    /// Alias for [`Self::write_deflated_stream_with_accounting`].
    pub fn write_deflated_reader_with_accounting<R: Read>(
        &mut self,
        name: &str,
        reader: R,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<(), Error> {
        self.write_deflated_stream_with_accounting(name, reader, accounting)
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
        if !self.name_policy.is_complete() {
            return Err(StreamingArchiveFailure {
                error: ErrorKind::InvalidInput {
                    msg: "generated ZIP member-name plan is not exhausted".to_string(),
                }
                .into(),
                progress: self.progress(),
                limit: None,
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
///
/// Completed payloads are retained in a bounded weighted LRU cache. Concurrent
/// cold reads of one member share one decompression flight; a failed flight is
/// removed after waking its waiters so a later call can retry.
///
/// [`LazyArchiveCacheLimits`] controls the retained cache, while the archive's
/// [`ArchiveLimits`] continue to govern declared and materialized ZIP sizes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LazyArchiveCacheLimits {
    max_bytes: usize,
    max_entries: usize,
    max_active_flights: usize,
    max_flight_key_bytes: usize,
}

impl LazyArchiveCacheLimits {
    /// Default maximum retained payload bytes.
    pub const DEFAULT_MAX_BYTES: usize = 8 * 1024 * 1024;
    /// Default maximum number of retained payloads.
    pub const DEFAULT_MAX_ENTRIES: usize = 128;
    /// Default maximum number of active same-member decompression flights.
    pub const DEFAULT_MAX_ACTIVE_FLIGHTS: usize = 64;
    /// Default aggregate bytes occupied by active-flight keys.
    pub const DEFAULT_MAX_FLIGHT_KEY_BYTES: usize = 256 * 1024;

    /// Construct finite cache limits.
    ///
    /// # Errors
    ///
    /// Returns [`ErrorKind::InvalidInput`] when either limit is zero.
    pub fn new(max_bytes: usize, max_entries: usize) -> Result<Self, Error> {
        if max_bytes == 0 {
            return Err(ErrorKind::InvalidInput {
                msg: "lazy archive cache byte limit must be non-zero".to_string(),
            }
            .into());
        }
        if max_entries == 0 {
            return Err(ErrorKind::InvalidInput {
                msg: "lazy archive cache entry limit must be non-zero".to_string(),
            }
            .into());
        }
        Ok(Self {
            max_bytes,
            max_entries,
            max_active_flights: Self::DEFAULT_MAX_ACTIVE_FLIGHTS,
            max_flight_key_bytes: Self::DEFAULT_MAX_FLIGHT_KEY_BYTES,
        })
    }

    /// Construct finite cache and active-flight limits.
    ///
    /// The active-flight count bounds retained flight objects, while the
    /// aggregate key-byte limit bounds the names retained by those flights.
    /// A request that cannot become a flight falls back to a direct read and
    /// therefore preserves the reader's ordinary typed result.
    pub fn new_with_active_flight_limits(
        max_bytes: usize,
        max_entries: usize,
        max_active_flights: usize,
        max_flight_key_bytes: usize,
    ) -> Result<Self, Error> {
        Self::new(max_bytes, max_entries)?
            .with_active_flight_limits(max_active_flights, max_flight_key_bytes)
    }

    /// Alias for [`Self::new_with_active_flight_limits`].
    pub fn new_with_flight_limits(
        max_bytes: usize,
        max_entries: usize,
        max_active_flights: usize,
        max_flight_key_bytes: usize,
    ) -> Result<Self, Error> {
        Self::new_with_active_flight_limits(
            max_bytes,
            max_entries,
            max_active_flights,
            max_flight_key_bytes,
        )
    }

    /// Add explicit active-flight object and key-byte limits to this policy.
    pub fn with_active_flight_limits(
        mut self,
        max_active_flights: usize,
        max_flight_key_bytes: usize,
    ) -> Result<Self, Error> {
        if max_active_flights == 0 {
            return Err(ErrorKind::InvalidInput {
                msg: "lazy archive active-flight limit must be non-zero".to_string(),
            }
            .into());
        }
        if max_flight_key_bytes == 0 {
            return Err(ErrorKind::InvalidInput {
                msg: "lazy archive active-flight key-byte limit must be non-zero".to_string(),
            }
            .into());
        }
        self.max_active_flights = max_active_flights;
        self.max_flight_key_bytes = max_flight_key_bytes;
        Ok(self)
    }

    /// Alias for [`Self::with_active_flight_limits`].
    pub fn with_flight_limits(
        self,
        max_active_flights: usize,
        max_flight_key_bytes: usize,
    ) -> Result<Self, Error> {
        self.with_active_flight_limits(max_active_flights, max_flight_key_bytes)
    }

    /// Maximum total decompressed bytes retained by the cache.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    /// Maximum number of completed payloads retained by the cache.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.max_entries
    }

    /// Maximum number of concurrently retained same-member flights.
    #[must_use]
    pub const fn max_active_flights(self) -> usize {
        self.max_active_flights
    }

    /// Maximum aggregate bytes occupied by active-flight keys.
    #[must_use]
    pub const fn max_flight_key_bytes(self) -> usize {
        self.max_flight_key_bytes
    }

    /// Alias for [`Self::max_flight_key_bytes`].
    #[must_use]
    pub const fn max_active_key_bytes(self) -> usize {
        self.max_flight_key_bytes
    }
}

impl Default for LazyArchiveCacheLimits {
    fn default() -> Self {
        Self {
            max_bytes: Self::DEFAULT_MAX_BYTES,
            max_entries: Self::DEFAULT_MAX_ENTRIES,
            max_active_flights: Self::DEFAULT_MAX_ACTIVE_FLIGHTS,
            max_flight_key_bytes: Self::DEFAULT_MAX_FLIGHT_KEY_BYTES,
        }
    }
}

#[derive(Debug)]
struct LazyCacheEntry {
    data: Arc<Vec<u8>>,
    weight: usize,
    last_used: u64,
}

#[derive(Debug, Default)]
struct LazyCacheState {
    entries: HashMap<String, LazyCacheEntry>,
    flights: HashMap<String, Arc<LazyFlight>>,
    active_flights: usize,
    active_key_bytes: usize,
    total_bytes: usize,
    next_recency: u64,
    generation: u64,
}

impl LazyCacheState {
    fn touch(&mut self, name: &str) {
        let recency = self.next_recency();
        if let Some(entry) = self.entries.get_mut(name) {
            entry.last_used = recency;
        }
    }

    fn next_recency(&mut self) -> u64 {
        if self.next_recency == u64::MAX {
            let mut entries = self.entries.values_mut().collect::<Vec<_>>();
            entries.sort_unstable_by_key(|entry| entry.last_used);
            for (index, entry) in entries.into_iter().enumerate() {
                entry.last_used = u64::try_from(index).unwrap_or(u64::MAX);
            }
            self.next_recency = u64::try_from(self.entries.len()).unwrap_or(u64::MAX);
        }
        let recency = self.next_recency;
        self.next_recency = self.next_recency.saturating_add(1);
        recency
    }

    fn evict_oldest(&mut self) -> bool {
        let Some(oldest_name) = self
            .entries
            .iter()
            // The cache owns one reference. An additional reference belongs
            // to a caller and pins the payload until that caller releases it.
            .filter(|(_, entry)| Arc::strong_count(&entry.data) == 1)
            .min_by_key(|(_, entry)| entry.last_used)
            .map(|(name, _)| name.clone())
        else {
            return false;
        };
        if let Some(removed) = self.entries.remove(&oldest_name) {
            self.total_bytes = self.total_bytes.saturating_sub(removed.weight);
            true
        } else {
            false
        }
    }

    fn insert(&mut self, name: String, data: Arc<Vec<u8>>, limits: LazyArchiveCacheLimits) {
        let weight = data.len();
        if weight > limits.max_bytes() {
            return;
        }
        while self.entries.len() >= limits.max_entries()
            || self.total_bytes.saturating_add(weight) > limits.max_bytes()
        {
            if !self.evict_oldest() {
                return;
            }
        }
        if self.entries.try_reserve(1).is_err() {
            return;
        }
        let last_used = self.next_recency();
        if let Some(previous) = self.entries.insert(
            name,
            LazyCacheEntry {
                data,
                weight,
                last_used,
            },
        ) {
            self.total_bytes = self.total_bytes.saturating_sub(previous.weight);
        }
        self.total_bytes = self.total_bytes.saturating_add(weight);
    }

    fn clear_entries(&mut self) {
        self.generation = self.generation.wrapping_add(1);
        self.entries.clear();
        // Detach in-progress work from this generation. Existing owners and
        // waiters retain their Arc<LazyFlight>, while subsequent readers must
        // install a new-generation flight rather than join old work.
        self.flights.clear();
        self.total_bytes = 0;
    }
}

#[derive(Debug, Default)]
struct LazyFlightState {
    complete: bool,
    data: Option<Arc<Vec<u8>>>,
}

#[derive(Debug)]
struct LazyFlight {
    generation: u64,
    key_bytes: usize,
    active: AtomicBool,
    state: std::sync::Mutex<LazyFlightState>,
    completed: std::sync::Condvar,
}

impl LazyFlight {
    fn new(generation: u64, key_bytes: usize) -> Self {
        Self {
            generation,
            key_bytes,
            active: AtomicBool::new(true),
            state: std::sync::Mutex::new(LazyFlightState::default()),
            completed: std::sync::Condvar::new(),
        }
    }

    fn complete_success(&self, data: Arc<Vec<u8>>) {
        let mut state = lock_lazy_cache(&self.state);
        state.data = Some(data);
        state.complete = true;
        self.completed.notify_all();
    }

    fn complete_failure(&self) {
        let mut state = lock_lazy_cache(&self.state);
        state.complete = true;
        self.completed.notify_all();
    }

    fn wait(&self) -> Option<Arc<Vec<u8>>> {
        let mut state = lock_lazy_cache(&self.state);
        while !state.complete {
            state = self
                .completed
                .wait(state)
                .unwrap_or_else(std::sync::PoisonError::into_inner);
        }
        state.data.as_ref().map(Arc::clone)
    }
}

fn lock_lazy_cache<T>(mutex: &std::sync::Mutex<T>) -> std::sync::MutexGuard<'_, T> {
    mutex
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner)
}

enum LazyCacheLookup {
    Hit(Arc<Vec<u8>>),
    Wait(Arc<LazyFlight>),
    Load(Arc<LazyFlight>),
    Bypass,
}

fn finish_lazy_flight(cache: &mut LazyCacheState, name: &str, flight: &Arc<LazyFlight>) {
    if flight.active.swap(false, Ordering::AcqRel) {
        cache.active_flights = cache.active_flights.saturating_sub(1);
        cache.active_key_bytes = cache.active_key_bytes.saturating_sub(flight.key_bytes);
    }
    if cache
        .flights
        .get(name)
        .is_some_and(|registered| Arc::ptr_eq(registered, flight))
    {
        cache.flights.remove(name);
    }
}

fn lazy_flight_key_limit_error(actual: usize, maximum: usize) -> Error {
    ErrorKind::InvalidInput {
        msg: format!(
            "lazy archive active-flight key is {actual} bytes; maximum is {maximum} bytes"
        ),
    }
    .into()
}

fn try_clone_lazy_key(name: &str) -> Option<String> {
    let mut clone = String::new();
    clone.try_reserve_exact(name.len()).ok()?;
    clone.push_str(name);
    Some(clone)
}

pub struct LazyArchiveReader<'data> {
    /// The underlying archive reader (for decompression)
    inner: ArchiveReader<'data>,
    /// Thread-safe cache and same-member decompression flights.
    cache: std::sync::Mutex<LazyCacheState>,
    cache_limits: LazyArchiveCacheLimits,
    #[cfg(test)]
    cold_loads: std::sync::atomic::AtomicU64,
}

impl<'data> LazyArchiveReader<'data> {
    /// Create a new lazy archive reader from a byte slice.
    pub fn new(data: &'data [u8]) -> Result<Self, Error> {
        Self::new_with_limits(data, ArchiveLimits::default())
    }

    /// Create a lazy reader with explicit resource limits.
    pub fn new_with_limits(data: &'data [u8], limits: ArchiveLimits) -> Result<Self, Error> {
        Self::new_with_limits_and_cache_limits(data, limits, LazyArchiveCacheLimits::default())
    }

    /// Create a lazy reader with explicit cache limits.
    pub fn new_with_cache_limits(
        data: &'data [u8],
        cache_limits: LazyArchiveCacheLimits,
    ) -> Result<Self, Error> {
        Self::new_with_limits_and_cache_limits(data, ArchiveLimits::default(), cache_limits)
    }

    /// Create a lazy reader with explicit archive and cache limits.
    pub fn new_with_limits_and_cache_limits(
        data: &'data [u8],
        limits: ArchiveLimits,
        cache_limits: LazyArchiveCacheLimits,
    ) -> Result<Self, Error> {
        let inner = ArchiveReader::new_with_limits(data, limits)?;
        Ok(Self {
            inner,
            cache: std::sync::Mutex::new(LazyCacheState::default()),
            cache_limits,
            #[cfg(test)]
            cold_loads: std::sync::atomic::AtomicU64::new(0),
        })
    }

    /// Return the finite retention policy used by this reader's cache.
    #[must_use]
    pub const fn cache_limits(&self) -> LazyArchiveCacheLimits {
        self.cache_limits
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
        let mut accounting = ZipOperationAccounting::default();
        self.read_with_accounting(name, &mut accounting)
    }

    /// Read and decompress a file while recording actual payload work.
    pub fn read_with_accounting(
        &self,
        name: &str,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<Vec<u8>, Error> {
        self.read_shared_with_accounting(name, accounting)
            .map(|arc| (*arc).clone())
    }

    /// Borrow and verify a stored member without populating the lazy cache.
    ///
    /// This is the zero-copy structural-ingress fast path for callers that
    /// consume a stored XML member immediately. Unencrypted Deflated members
    /// return `None` and should use [`Self::read`] instead. Targets declaring
    /// encryption in general-purpose flag bits 0 or 6 return a typed error.
    /// Nonempty Store members with a declared CRC-32 of zero also return
    /// `None` because borrowed access cannot verify them.
    /// The returned slice remains tied to the source bytes borrowed by this
    /// reader.
    #[inline]
    pub fn read_stored_borrowed(&self, name: &str) -> Result<Option<&'data [u8]>, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_stored_borrowed_with_accounting(name, &mut accounting)
    }

    /// Borrow and verify a stored member without charging payload work.
    #[inline]
    pub fn read_stored_borrowed_with_accounting(
        &self,
        name: &str,
        _accounting: &mut ZipOperationAccounting,
    ) -> Result<Option<&'data [u8]>, Error> {
        self.inner.read_stored_borrowed(name)
    }

    /// Read and decompress a file, returning a shared reference.
    ///
    /// This is more efficient than `read()` when the same file is accessed
    /// multiple times, as it avoids cloning the decompressed data.
    pub fn read_shared(&self, name: &str) -> Result<std::sync::Arc<Vec<u8>>, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_shared_with_accounting(name, &mut accounting)
    }

    /// Read and decompress a file while recording actual payload work.
    ///
    /// Cache hits and waiters that receive a completed flight do not charge any
    /// payload bytes. Only the cold-flight loader charges the supplied value.
    pub fn read_shared_with_accounting(
        &self,
        name: &str,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<std::sync::Arc<Vec<u8>>, Error> {
        let mut read_member =
            |member_name: &str, member_accounting: &mut ZipOperationAccounting| {
                self.inner
                    .read_with_accounting(member_name, member_accounting)
            };
        self.read_shared_with_reader(name, accounting, &mut read_member)
    }

    fn read_shared_with_reader(
        &self,
        name: &str,
        accounting: &mut ZipOperationAccounting,
        read_member: &mut dyn FnMut(&str, &mut ZipOperationAccounting) -> Result<Vec<u8>, Error>,
    ) -> Result<std::sync::Arc<Vec<u8>>, Error> {
        if name.len() > self.cache_limits.max_flight_key_bytes() {
            return Err(lazy_flight_key_limit_error(
                name.len(),
                self.cache_limits.max_flight_key_bytes(),
            ));
        }
        let lookup = lookup_member_name(name)?;
        if lookup.explicit_directory {
            return Err(ErrorKind::FileNotFound(lookup.name).into());
        }
        let normalized = lookup.name;
        if normalized.len() > self.cache_limits.max_flight_key_bytes() {
            return Err(lazy_flight_key_limit_error(
                normalized.len(),
                self.cache_limits.max_flight_key_bytes(),
            ));
        }

        loop {
            let cache_lookup = {
                let mut cache = lock_lazy_cache(&self.cache);
                if let Some(data) = cache
                    .entries
                    .get(&normalized)
                    .map(|entry| Arc::clone(&entry.data))
                {
                    cache.touch(&normalized);
                    LazyCacheLookup::Hit(data)
                } else if let Some(flight) = cache.flights.get(&normalized) {
                    LazyCacheLookup::Wait(Arc::clone(flight))
                } else if cache.active_flights >= self.cache_limits.max_active_flights()
                    || cache
                        .active_key_bytes
                        .checked_add(normalized.len())
                        .is_none()
                    || cache.active_key_bytes + normalized.len()
                        > self.cache_limits.max_flight_key_bytes()
                {
                    // The active-flight policy is a coordination budget, not
                    // a read correctness limit. Once it is full, decompress
                    // directly without retaining another key or flight.
                    LazyCacheLookup::Bypass
                } else if cache.flights.try_reserve(1).is_err() {
                    // Cache bookkeeping is best effort. A failed internal
                    // reservation must not change decompression/error behavior.
                    LazyCacheLookup::Bypass
                } else {
                    let mut flight_name = String::new();
                    if flight_name.try_reserve_exact(normalized.len()).is_err() {
                        LazyCacheLookup::Bypass
                    } else {
                        flight_name.push_str(&normalized);
                        let flight = Arc::new(LazyFlight::new(cache.generation, normalized.len()));
                        cache.flights.insert(flight_name, Arc::clone(&flight));
                        cache.active_flights += 1;
                        cache.active_key_bytes += normalized.len();
                        #[cfg(test)]
                        self.cold_loads
                            .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
                        LazyCacheLookup::Load(flight)
                    }
                }
            };

            match cache_lookup {
                LazyCacheLookup::Hit(data) => return Ok(data),
                LazyCacheLookup::Bypass => {
                    return read_member(&normalized, accounting).map(Arc::new);
                },
                LazyCacheLookup::Wait(flight) => {
                    // A failed flight carries no Error: the original owner
                    // returns its original error, while waiters retry and
                    // therefore preserve each operation's error/source chain.
                    if let Some(data) = flight.wait() {
                        return Ok(data);
                    }
                },
                LazyCacheLookup::Load(flight) => {
                    let result = read_member(&normalized, accounting).map(Arc::new);
                    match result {
                        Ok(data) => {
                            let mut cache = lock_lazy_cache(&self.cache);
                            if cache.generation == flight.generation {
                                if let Some(cache_name) = try_clone_lazy_key(&normalized) {
                                    cache.insert(cache_name, Arc::clone(&data), self.cache_limits);
                                }
                            }
                            // Publish before removing the flight. This keeps
                            // oversized and generation-fenced successes shared
                            // with callers that arrive during completion.
                            flight.complete_success(Arc::clone(&data));
                            finish_lazy_flight(&mut cache, &normalized, &flight);
                            return Ok(data);
                        },
                        Err(error) => {
                            let mut cache = lock_lazy_cache(&self.cache);
                            // Wake waiters before allowing a retrying loader to
                            // install a replacement flight.
                            flight.complete_failure();
                            finish_lazy_flight(&mut cache, &normalized, &flight);
                            return Err(error);
                        },
                    }
                },
            }
        }
    }

    /// Read multiple members serially, sharing the lazy cache and one
    /// resettable Deflate decoder for cold members.
    ///
    /// Results retain caller input order and preserve the per-member error
    /// behavior of [`Self::read_shared`]. Cache hits and same-member waiters
    /// use the existing cache state machine; Store members bypass the decoder.
    /// This API is intended for an owning serial load that has already
    /// performed its structural and declared-size preflight.
    pub fn read_many_serial_shared<'name>(
        &self,
        names: &'name [&'name str],
    ) -> Vec<(&'name str, Result<std::sync::Arc<Vec<u8>>, Error>)> {
        let mut session = self.inner.read_session();
        names
            .iter()
            .map(|name| {
                let mut accounting = ZipOperationAccounting::default();
                let mut read_member =
                    |member_name: &str, member_accounting: &mut ZipOperationAccounting| {
                        session.read_with_accounting(member_name, member_accounting)
                    };
                (
                    *name,
                    self.read_shared_with_reader(name, &mut accounting, &mut read_member),
                )
            })
            .collect()
    }

    /// Decompress and verify one member directly into a caller-owned sink.
    ///
    /// This operation intentionally bypasses the lazy decompression cache, so
    /// consumers can process a large member incrementally without retaining a
    /// second complete payload. The sink may contain a valid prefix when an
    /// I/O, checksum, or size error is returned.
    #[inline]
    pub fn read_to<W: Write>(&self, name: &str, sink: &mut W) -> Result<u64, Error> {
        let mut accounting = ZipOperationAccounting::default();
        self.read_to_with_accounting(name, sink, &mut accounting)
    }

    /// Decompress and verify one member into a sink while recording actual
    /// source traversal and Deflate destination acceptance.
    #[inline]
    pub fn read_to_with_accounting<W: Write>(
        &self,
        name: &str,
        sink: &mut W,
        accounting: &mut ZipOperationAccounting,
    ) -> Result<u64, Error> {
        self.inner.read_to_with_accounting(name, sink, accounting)
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
        lock_lazy_cache(&self.cache).entries.len()
    }

    /// Get the total decompressed bytes retained by the cache.
    pub fn cache_bytes(&self) -> usize {
        lock_lazy_cache(&self.cache).total_bytes
    }

    /// Get the number of active same-member decompression flights.
    pub fn active_flight_count(&self) -> usize {
        lock_lazy_cache(&self.cache).active_flights
    }

    /// Get the aggregate bytes occupied by active-flight keys.
    pub fn active_flight_key_bytes(&self) -> usize {
        lock_lazy_cache(&self.cache).active_key_bytes
    }

    /// Clear the decompression cache to free memory.
    pub fn clear_cache(&self) {
        lock_lazy_cache(&self.cache).clear_entries();
    }

    /// Take ownership of cached data, consuming the cache.
    ///
    /// Returns all cached files and clears the cache. This is useful when
    /// you want to take ownership of the decompressed data without cloning.
    pub fn take_cache(&self) -> HashMap<String, Vec<u8>> {
        let mut cache = lock_lazy_cache(&self.cache);
        cache.generation = cache.generation.wrapping_add(1);
        // See `clear_entries`: detached flights still wake their existing
        // waiters, but a post-take reader must not join their old generation.
        cache.flights.clear();
        let mut result = HashMap::with_capacity(cache.entries.len());
        for (name, entry) in cache.entries.drain() {
            // Try to unwrap the Arc; if there are other references, clone instead
            match std::sync::Arc::try_unwrap(entry.data) {
                Ok(data) => {
                    result.insert(name, data);
                },
                Err(arc) => {
                    result.insert(name, (*arc).clone());
                },
            }
        }
        cache.total_bytes = 0;
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
    use crate::Crc32Option;
    use flate2::Compression;
    use flate2::write::DeflateEncoder;
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
    struct FailOnRangeReaderAt {
        bytes: Vec<u8>,
        fail_start: u64,
        fail_end: u64,
    }

    impl ReaderAt for FailOnRangeReaderAt {
        fn read_at(&self, buffer: &mut [u8], offset: u64) -> io::Result<usize> {
            let request_end = offset.saturating_add(buffer.len() as u64);
            if offset < self.fail_end && request_end > self.fail_start {
                return Err(io::Error::new(
                    io::ErrorKind::UnexpectedEof,
                    "injected indexed source read failure",
                ));
            }
            self.bytes.as_slice().read_at(buffer, offset)
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

    #[derive(Debug)]
    struct OverreportingReader;

    impl Read for OverreportingReader {
        fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
            Ok(buffer.len().saturating_add(1))
        }
    }

    #[derive(Debug)]
    struct InterruptingWriter {
        bytes: Vec<u8>,
        interrupted: bool,
    }

    impl Write for InterruptingWriter {
        fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
            if !self.interrupted {
                self.interrupted = true;
                return Err(io::Error::new(io::ErrorKind::Interrupted, "try again"));
            }
            self.bytes.extend_from_slice(buffer);
            Ok(buffer.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[derive(Debug)]
    struct OverreportingWriter;

    impl Write for OverreportingWriter {
        fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
            Ok(buffer.len().saturating_add(1))
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn bounded_output_rejects_sink_overreport_before_progress_accounting() {
        let counter = std::sync::Arc::new(std::sync::atomic::AtomicU64::new(0));
        let mut output = BoundedOutput::new(OverreportingWriter, 64, counter.clone());
        let error = output.write(b"payload").unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::InvalidData);
        assert_eq!(output.accepted, 0);
        assert_eq!(counter.load(Ordering::Acquire), 0);
    }

    #[test]
    fn accounted_stream_retries_interrupted_source_and_sink() {
        let payload = b"retry payload";
        let mut source =
            ScriptedReader::new([ReadStep::Interrupted, ReadStep::Bytes(payload.to_vec())]);
        let mut sink = InterruptingWriter {
            bytes: Vec::new(),
            interrupted: false,
        };
        let mut accounting = ZipOperationAccounting::default();
        let copied = stream_verified_with_accounting(
            &mut source,
            ZipVerification {
                crc: crate::crc32(payload),
                uncompressed_size: payload.len() as u64,
            },
            &mut sink,
            &mut accounting,
            AccountingReadKind::Deflate,
        )
        .unwrap();
        assert_eq!(copied, payload.len() as u64);
        assert_eq!(sink.bytes, payload);
        assert_eq!(accounting.deflate_bytes_produced(), payload.len() as u64);
        assert_eq!(accounting.deflate_bytes_accepted(), payload.len() as u64);
    }

    #[test]
    fn accounted_stream_rejects_source_overreport_without_progress() {
        let mut source = OverreportingReader;
        let mut sink = Vec::new();
        let mut accounting = ZipOperationAccounting::default();
        let error = stream_verified_with_accounting(
            &mut source,
            ZipVerification {
                crc: 1,
                uncompressed_size: 1,
            },
            &mut sink,
            &mut accounting,
            AccountingReadKind::Stored,
        )
        .unwrap_err();
        match error.kind() {
            ErrorKind::IO(error) | ErrorKind::Io(error) => {
                assert_eq!(error.kind(), io::ErrorKind::InvalidData);
            },
            other => panic!("expected invalid read-count error, got {other:?}"),
        }
        assert!(sink.is_empty());
        assert_eq!(accounting.stored_payload_bytes_accepted(), 0);
        assert_eq!(accounting.stored_payload_bytes_read(), 0);
    }

    #[test]
    fn deflate_overrun_keeps_size_error_before_accounting_overflow() {
        let mut source = ScriptedReader::new([ReadStep::Bytes(vec![b'x'])]);
        let mut sink = Vec::new();
        let mut accounting = ZipOperationAccounting::default();
        accounting.add_deflate_bytes_produced(u64::MAX).unwrap();

        let error = stream_verified_with_accounting(
            &mut source,
            ZipVerification {
                crc: 0,
                uncompressed_size: 0,
            },
            &mut sink,
            &mut accounting,
            AccountingReadKind::Deflate,
        )
        .unwrap_err();
        assert!(matches!(
            error.kind(),
            ErrorKind::InvalidSize {
                expected: 0,
                actual: 1
            }
        ));
        assert_eq!(accounting.deflate_bytes_produced(), u64::MAX);
        assert!(sink.is_empty());
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
    fn read_to_streams_stored_and_deflated_members_with_short_writes() {
        let payload = (0..(STREAM_COPY_BUFFER_SIZE * 2 + 37))
            .map(|index| (index as u8).wrapping_mul(31))
            .collect::<Vec<_>>();
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", &payload).unwrap();
        writer.write_deflated("deflated.bin", &payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let reader = ArchiveReader::new(&bytes).unwrap();
        let mut stored = ShortWriter::new(7);
        assert_eq!(
            reader.read_to("stored.bin", &mut stored).unwrap(),
            payload.len() as u64
        );
        assert_eq!(stored.bytes, payload);

        let mut deflated = ShortWriter::new(11);
        assert_eq!(
            reader.read_to("deflated.bin", &mut deflated).unwrap(),
            payload.len() as u64
        );
        assert_eq!(deflated.bytes, payload);
    }

    #[test]
    fn accounting_distinguishes_store_deflate_and_lazy_cache_work() {
        let stored_payload = b"stored payload";
        let deflated_payload = b"deflated payload with enough bytes to exercise accounting";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", stored_payload).unwrap();
        writer
            .write_deflated("deflated.bin", deflated_payload)
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let reader = ArchiveReader::new(&bytes).unwrap();
        let mut stored_accounting = ZipOperationAccounting::default();
        assert_eq!(
            reader
                .read_with_accounting("stored.bin", &mut stored_accounting)
                .unwrap(),
            stored_payload
        );
        assert_eq!(
            stored_accounting.stored_payload_bytes_read(),
            stored_payload.len() as u64
        );
        assert_eq!(stored_accounting.compressed_deflate_payload_bytes_read(), 0);
        assert_eq!(stored_accounting.deflate_bytes_produced(), 0);

        let mut deflated_accounting = ZipOperationAccounting::default();
        assert_eq!(
            reader
                .read_with_accounting("deflated.bin", &mut deflated_accounting)
                .unwrap(),
            deflated_payload
        );
        assert!(deflated_accounting.compressed_deflate_payload_bytes_read() > 0);
        assert_eq!(
            deflated_accounting.deflate_bytes_produced(),
            deflated_payload.len() as u64
        );
        assert_eq!(
            deflated_accounting.deflate_bytes_accepted(),
            deflated_payload.len() as u64
        );

        let mut sink = ShortWriter::new(3);
        let mut streamed_accounting = ZipOperationAccounting::default();
        assert_eq!(
            reader
                .read_to_with_accounting("deflated.bin", &mut sink, &mut streamed_accounting,)
                .unwrap(),
            deflated_payload.len() as u64
        );
        assert_eq!(sink.bytes, deflated_payload);
        assert_eq!(
            streamed_accounting.deflate_bytes_accepted(),
            deflated_payload.len() as u64
        );

        let mut stored_sink = FailingWriter::new(5);
        let mut stored_stream_accounting = ZipOperationAccounting::default();
        let error = reader
            .read_to_with_accounting(
                "stored.bin",
                &mut stored_sink,
                &mut stored_stream_accounting,
            )
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert_eq!(
            stored_stream_accounting.stored_payload_bytes_read(),
            stored_payload.len() as u64
        );
        assert_eq!(stored_stream_accounting.stored_payload_bytes_accepted(), 5);

        let lazy = LazyArchiveReader::new(&bytes).unwrap();
        assert_eq!(lazy.cache_size(), 0);
        let mut cold = ZipOperationAccounting::default();
        lazy.read_with_accounting("deflated.bin", &mut cold)
            .unwrap();
        assert!(cold.compressed_deflate_payload_bytes_read() > 0);
        assert_eq!(lazy.cache_size(), 1);
        assert_eq!(lazy.cache_bytes(), deflated_payload.len());
        let mut hit = ZipOperationAccounting::default();
        lazy.read_with_accounting("deflated.bin", &mut hit).unwrap();
        assert_eq!(hit, ZipOperationAccounting::default());
        lazy.clear_cache();
        assert_eq!(lazy.cache_size(), 0);
        let mut after_clear = ZipOperationAccounting::default();
        lazy.read_with_accounting("deflated.bin", &mut after_clear)
            .unwrap();
        assert!(after_clear.compressed_deflate_payload_bytes_read() > 0);

        let bypass_limits = LazyArchiveCacheLimits::new(1, 1).unwrap();
        let bypass = LazyArchiveReader::new_with_cache_limits(&bytes, bypass_limits).unwrap();
        let mut bypass_accounting = ZipOperationAccounting::default();
        bypass
            .read_with_accounting("deflated.bin", &mut bypass_accounting)
            .unwrap();
        let bypass_compressed_read = bypass_accounting.compressed_deflate_payload_bytes_read();
        assert!(bypass_compressed_read > 0);
        assert_eq!(bypass.cache_size(), 0);
        assert_eq!(bypass.cache_bytes(), 0);
        assert_eq!(
            bypass.cold_loads.load(std::sync::atomic::Ordering::SeqCst),
            1
        );
        let mut bypass_again_accounting = ZipOperationAccounting::default();
        bypass
            .read_with_accounting("deflated.bin", &mut bypass_again_accounting)
            .unwrap();
        assert_eq!(
            bypass_again_accounting.compressed_deflate_payload_bytes_read(),
            bypass_compressed_read
        );
        assert_eq!(bypass.cache_size(), 0);
        assert_eq!(bypass.cache_bytes(), 0);
        assert_eq!(
            bypass.cold_loads.load(std::sync::atomic::Ordering::SeqCst),
            2
        );

        let indexed = indexed_archive(bytes.clone());
        let indexed_id = indexed.entry_id("deflated.bin").unwrap();
        let mut indexed_accounting = ZipOperationAccounting::default();
        assert_eq!(
            indexed
                .read_entry_with_accounting(indexed_id, &mut indexed_accounting)
                .unwrap(),
            deflated_payload
        );
        assert!(indexed_accounting.compressed_deflate_payload_bytes_read() > 0);
        let mut indexed_sink = ShortWriter::new(4);
        let mut indexed_stream_accounting = ZipOperationAccounting::default();
        indexed
            .read_entry_to_with_accounting(
                indexed_id,
                &mut indexed_sink,
                &mut indexed_stream_accounting,
            )
            .unwrap();
        assert_eq!(indexed_sink.bytes, deflated_payload);
        assert_eq!(
            indexed_stream_accounting.deflate_bytes_accepted(),
            deflated_payload.len() as u64
        );

        let mut indexed_deflated_failure = FailingWriter::new(3);
        let mut indexed_deflated_failure_accounting = ZipOperationAccounting::default();
        let error = indexed
            .read_entry_to_with_accounting(
                indexed_id,
                &mut indexed_deflated_failure,
                &mut indexed_deflated_failure_accounting,
            )
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert!(indexed_deflated_failure_accounting.compressed_deflate_payload_bytes_read() > 0);
        assert_eq!(
            indexed_deflated_failure_accounting.deflate_bytes_accepted(),
            indexed_deflated_failure.bytes.len() as u64
        );

        let indexed_stored_id = indexed.entry_id("stored.bin").unwrap();
        let mut indexed_stored_accounting = ZipOperationAccounting::default();
        assert_eq!(
            indexed
                .read_entry_with_accounting(indexed_stored_id, &mut indexed_stored_accounting)
                .unwrap(),
            stored_payload
        );
        assert_eq!(
            indexed_stored_accounting.stored_payload_bytes_read(),
            stored_payload.len() as u64
        );
        assert_eq!(
            indexed_stored_accounting.stored_payload_bytes_accepted(),
            stored_payload.len() as u64
        );
        assert_eq!(
            indexed_stored_accounting.compressed_deflate_payload_bytes_read(),
            0
        );
        let mut indexed_stored_sink = ShortWriter::new(2);
        let mut indexed_stored_stream_accounting = ZipOperationAccounting::default();
        assert_eq!(
            indexed
                .read_entry_to_with_accounting(
                    indexed_stored_id,
                    &mut indexed_stored_sink,
                    &mut indexed_stored_stream_accounting,
                )
                .unwrap(),
            stored_payload.len() as u64
        );
        assert_eq!(indexed_stored_sink.bytes, stored_payload);
        assert_eq!(
            indexed_stored_stream_accounting.stored_payload_bytes_read(),
            stored_payload.len() as u64
        );
        assert_eq!(
            indexed_stored_stream_accounting.stored_payload_bytes_accepted(),
            stored_payload.len() as u64
        );
        assert_eq!(
            indexed_stored_stream_accounting.compressed_deflate_payload_bytes_read(),
            0
        );

        let mut indexed_stored_failure = FailingWriter::new(2);
        let mut indexed_stored_failure_accounting = ZipOperationAccounting::default();
        let error = indexed
            .read_entry_to_with_accounting(
                indexed_stored_id,
                &mut indexed_stored_failure,
                &mut indexed_stored_failure_accounting,
            )
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert!(indexed_stored_failure_accounting.stored_payload_bytes_read() > 0);
        assert_eq!(
            indexed_stored_failure_accounting.stored_payload_bytes_accepted(),
            indexed_stored_failure.bytes.len() as u64
        );

        let mut borrowed = ZipOperationAccounting::default();
        assert_eq!(
            reader
                .read_stored_borrowed_with_accounting("stored.bin", &mut borrowed)
                .unwrap(),
            Some(&stored_payload[..])
        );
        assert_eq!(borrowed, ZipOperationAccounting::default());

        let mut writing = StreamingArchiveWriter::new();
        let mut writing_accounting = ZipOperationAccounting::default();
        writing
            .write_stored_with_accounting("stored.bin", stored_payload, &mut writing_accounting)
            .unwrap();
        writing
            .write_deflated_with_accounting(
                "deflated.bin",
                deflated_payload,
                &mut writing_accounting,
            )
            .unwrap();
        writing
            .write_deflated_sized_with_accounting(
                "sized.bin",
                deflated_payload,
                &mut writing_accounting,
            )
            .unwrap();
        writing.finish_to_bytes().unwrap();
        assert_eq!(
            writing_accounting.stored_payload_bytes_emitted(),
            stored_payload.len() as u64
        );
        assert!(writing_accounting.generated_deflate_payload_bytes_emitted() > 0);
        assert_eq!(writing_accounting.precompressed_payload_bytes_emitted(), 0);
    }

    #[test]
    fn accounting_preserves_partial_deflate_sink_progress() {
        let payload = vec![b'x'; STREAM_COPY_BUFFER_SIZE + 17];
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("payload.bin", &payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let reader = ArchiveReader::new(&bytes).unwrap();
        let mut sink = FailingWriter::new(19);
        let mut accounting = ZipOperationAccounting::default();
        let error = reader
            .read_to_with_accounting("payload.bin", &mut sink, &mut accounting)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert!(accounting.compressed_deflate_payload_bytes_read() > 0);
        assert_eq!(accounting.deflate_bytes_accepted(), sink.bytes.len() as u64);
        assert!(accounting.deflate_bytes_produced() >= accounting.deflate_bytes_accepted());
    }

    #[test]
    fn indexed_and_lazy_read_to_stream_without_cache_population() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_deflated("content.xml", b"<content>Hello</content>")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let indexed = indexed_archive(bytes.clone());
        let id = indexed.entry_id("content.xml").unwrap();
        let mut indexed_output = Vec::new();
        assert_eq!(
            indexed.read_entry_to(id, &mut indexed_output).unwrap(),
            b"<content>Hello</content>".len() as u64
        );
        assert_eq!(indexed_output, b"<content>Hello</content>");

        let lazy = LazyArchiveReader::new(&bytes).unwrap();
        let mut lazy_output = Vec::new();
        assert_eq!(
            lazy.read_to("content.xml", &mut lazy_output).unwrap(),
            b"<content>Hello</content>".len() as u64
        );
        assert_eq!(lazy_output, b"<content>Hello</content>");
        assert_eq!(lazy.cache_size(), 0);
    }

    #[test]
    fn read_to_retains_typed_integrity_errors_and_sink_failures() {
        let mut bytes = bulk_fixture();
        corrupt_payload(&mut bytes, b"bad");
        let reader = ArchiveReader::new(&bytes).unwrap();
        let mut output = Vec::new();
        let error = reader.read_to("bad", &mut output).unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidChecksum { .. }));
        assert_eq!(output, vec![b'b' ^ 0x80, b'a', b'd']);

        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_deflated("payload.bin", &[b'x'; STREAM_COPY_BUFFER_SIZE + 1])
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let reader = ArchiveReader::new(&bytes).unwrap();
        let mut failing = FailingWriter::new(5);
        let error = reader.read_to("payload.bin", &mut failing).unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert_eq!(failing.bytes.len(), 5);
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
    fn streaming_limits_reject_oversized_member_name_before_output() {
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
    fn streaming_zip64_admission_selects_framing_for_borrowed_and_owned_entries() {
        let zip32_boundary = u64::from(u32::MAX);
        let limits = StreamingArchiveLimits::new(u16::MAX as usize, 64, 4096)
            .with_byte_limits(zip32_boundary, zip32_boundary, zip32_boundary)
            .with_compressed_size_limit(zip32_boundary);
        let mut writer = StreamingArchiveWriter::with_limits(limits);

        writer
            .write_deflated("borrowed-deflated", b"borrowed payload")
            .unwrap();
        writer
            .write_stored_stream("reader-stored", Cursor::new(b"reader payload".to_vec()))
            .unwrap();

        let mut entry = writer
            .start_entry("owned-stored", CompressionMethod::Store)
            .unwrap();
        entry.write_all(b"owned payload").unwrap();
        writer = entry.finish().unwrap();

        let bytes = writer.finish_to_bytes().unwrap();
        assert!(ZipArchive::from_slice(&bytes).unwrap().is_zip64());
        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(
            reader.read("borrowed-deflated").unwrap(),
            b"borrowed payload"
        );
        assert_eq!(reader.read("reader-stored").unwrap(), b"reader payload");
        assert_eq!(reader.read("owned-stored").unwrap(), b"owned payload");

        for name in [
            b"borrowed-deflated" as &[u8],
            b"reader-stored",
            b"owned-stored",
        ] {
            let local = local_header_offset_for_name(&bytes, name);
            assert_eq!(
                u16::from_le_bytes([bytes[local + 4], bytes[local + 5]]),
                45,
                "ZIP64 local header version for {}",
                String::from_utf8_lossy(name)
            );
            assert!(local_member_has_data_descriptor(&bytes, name));
        }
    }

    #[test]
    fn streaming_zip64_metadata_budget_includes_generated_central_extra() {
        let zip32_boundary = u64::from(u32::MAX);
        let borrowed_name = "borrowed.bin";
        let owned_name = "owned.bin";
        let borrowed_metadata = borrowed_name.len() as u64 + 20;
        let expected_metadata = borrowed_metadata + owned_name.len() as u64 + 20;
        let limits = StreamingArchiveLimits::new(4, 64, expected_metadata)
            .with_byte_limits(zip32_boundary, zip32_boundary, zip32_boundary)
            .with_compressed_size_limit(zip32_boundary);
        let mut writer = StreamingArchiveWriter::with_limits(limits);
        writer.write_deflated(borrowed_name, b"payload").unwrap();
        assert_eq!(writer.metadata_bytes(), borrowed_metadata);

        let mut entry = writer
            .start_entry(owned_name, CompressionMethod::Store)
            .unwrap();
        entry.write_all(b"owned").unwrap();
        writer = entry.finish().unwrap();
        assert_eq!(writer.metadata_bytes(), expected_metadata);

        let name_only_budget = StreamingArchiveLimits::new(4, 64, borrowed_name.len() as u64)
            .with_byte_limits(zip32_boundary, zip32_boundary, zip32_boundary)
            .with_compressed_size_limit(zip32_boundary);
        let mut name_only = StreamingArchiveWriter::with_limits(name_only_budget);
        let error = name_only
            .write_deflated(borrowed_name, b"payload")
            .unwrap_err();
        assert_limit(
            error,
            LimitResource::MetadataBytes,
            borrowed_metadata,
            borrowed_name.len() as u64,
        );
        assert_eq!(name_only.output_bytes(), 0);
    }

    #[test]
    fn generated_central_zip64_extra_size_matches_size_and_offset_fields() {
        let below = u64::from(u32::MAX) - 1;
        let at = u64::from(u32::MAX);
        assert_eq!(
            StreamingArchiveWriter::<Cursor<Vec<u8>>>::generated_central_zip64_extra_bytes(
                false,
                Some(below),
                Some(below),
                below,
            ),
            0
        );
        assert_eq!(
            StreamingArchiveWriter::<Cursor<Vec<u8>>>::generated_central_zip64_extra_bytes(
                false,
                Some(at),
                Some(at),
                below,
            ),
            20
        );
        assert_eq!(
            StreamingArchiveWriter::<Cursor<Vec<u8>>>::generated_central_zip64_extra_bytes(
                false,
                Some(below),
                Some(below),
                at,
            ),
            12
        );
        assert_eq!(
            StreamingArchiveWriter::<Cursor<Vec<u8>>>::generated_central_zip64_extra_bytes(
                false,
                Some(at),
                Some(at),
                at,
            ),
            28
        );
        assert_eq!(
            StreamingArchiveWriter::<Cursor<Vec<u8>>>::generated_central_zip64_extra_bytes(
                true, None, None, below,
            ),
            20
        );
        assert_eq!(
            StreamingArchiveWriter::<Cursor<Vec<u8>>>::generated_central_zip64_extra_bytes(
                true, None, None, at,
            ),
            28
        );
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
    fn archive_read_session_reuses_deflate_state_with_accounting_parity() {
        let stored_first = b"stored first descriptor-free payload";
        let deflated_first = b"deflated first payload with a data descriptor";
        let stored_second = b"stored second payload with a data descriptor";
        let deflated_second = b"deflated second payload with a data descriptor";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored-first", stored_first).unwrap();
        writer
            .write_deflated("deflated-first", deflated_first)
            .unwrap();
        writer
            .write_stored_stream("stored-second", Cursor::new(stored_second.to_vec()))
            .unwrap();
        writer
            .write_deflated("deflated-second", deflated_second)
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        assert!(local_member_has_data_descriptor(&bytes, b"deflated-first"));
        assert!(local_member_has_data_descriptor(&bytes, b"stored-second"));
        assert!(local_member_has_data_descriptor(&bytes, b"deflated-second"));

        let reader = ArchiveReader::new(&bytes).unwrap();
        let members = [
            ("stored-first", stored_first.as_slice()),
            ("deflated-first", deflated_first.as_slice()),
            ("stored-second", stored_second.as_slice()),
            ("deflated-second", deflated_second.as_slice()),
        ];
        let mut expected = Vec::new();
        for (name, payload) in members {
            let mut accounting = ZipOperationAccounting::default();
            let decoded = reader.read_with_accounting(name, &mut accounting).unwrap();
            assert_eq!(decoded, payload);
            expected.push((decoded, accounting));
        }

        let mut session = reader.read_session();
        for ((name, payload), (expected_payload, expected_accounting)) in
            members.into_iter().zip(expected)
        {
            let mut accounting = ZipOperationAccounting::default();
            let decoded = session.read_with_accounting(name, &mut accounting).unwrap();
            assert_eq!(decoded, payload);
            assert_eq!(decoded, expected_payload);
            assert_eq!(accounting, expected_accounting);
        }
    }

    #[test]
    fn archive_read_session_resets_after_crc_and_corrupt_deflate_failures() {
        const VALID: &[u8] = b"valid payload after a failed Deflate read";

        let mut crc_writer = StreamingArchiveWriter::new();
        crc_writer
            .write_deflated_sized("bad-crc", b"payload with a bad CRC")
            .unwrap();
        crc_writer.write_deflated_sized("valid", VALID).unwrap();
        let mut crc_bytes = crc_writer.finish_to_bytes().unwrap();
        let crc_central = central_header_offset_for_name(&crc_bytes, b"bad-crc");
        let crc = u32::from_le_bytes(
            crc_bytes[crc_central + 16..crc_central + 20]
                .try_into()
                .unwrap(),
        );
        crc_bytes[crc_central + 16..crc_central + 20].copy_from_slice(&(crc ^ 1).to_le_bytes());
        let crc_reader = ArchiveReader::new(&crc_bytes).unwrap();
        let mut crc_session = crc_reader.read_session();
        let error = crc_session.read("bad-crc").unwrap_err();
        assert_materialized_checksum_error(error);
        assert_eq!(crc_session.read("valid").unwrap(), VALID);

        let mut corrupt_writer = StreamingArchiveWriter::new();
        corrupt_writer
            .write_deflated_sized("corrupt", b"payload with corrupt Deflate bytes")
            .unwrap();
        corrupt_writer.write_deflated_sized("valid", VALID).unwrap();
        let mut corrupt_bytes = corrupt_writer.finish_to_bytes().unwrap();
        let corrupt_local = local_header_offset_for_name(&corrupt_bytes, b"corrupt");
        let name_len = usize::from(u16::from_le_bytes(
            corrupt_bytes[corrupt_local + 26..corrupt_local + 28]
                .try_into()
                .unwrap(),
        ));
        let extra_len = usize::from(u16::from_le_bytes(
            corrupt_bytes[corrupt_local + 28..corrupt_local + 30]
                .try_into()
                .unwrap(),
        ));
        let compressed_start = corrupt_local + 30 + name_len + extra_len;
        corrupt_bytes[compressed_start] ^= 0xff;
        let corrupt_reader = ArchiveReader::new(&corrupt_bytes).unwrap();
        let mut corrupt_session = corrupt_reader.read_session();
        assert!(corrupt_session.read("corrupt").is_err());
        assert_eq!(corrupt_session.read("valid").unwrap(), VALID);

        let mut truncated_writer = StreamingArchiveWriter::new();
        truncated_writer
            .write_deflated_sized("truncated", b"payload with a truncated Deflate stream")
            .unwrap();
        truncated_writer
            .write_deflated_sized("valid", VALID)
            .unwrap();
        let mut truncated_bytes = truncated_writer.finish_to_bytes().unwrap();
        let truncated_local = local_header_offset_for_name(&truncated_bytes, b"truncated");
        let truncated_central = central_header_offset_for_name(&truncated_bytes, b"truncated");
        let compressed_size = u32::from_le_bytes(
            truncated_bytes[truncated_local + 18..truncated_local + 22]
                .try_into()
                .unwrap(),
        );
        assert!(compressed_size > 1);
        truncated_bytes[truncated_local + 18..truncated_local + 22]
            .copy_from_slice(&(compressed_size - 1).to_le_bytes());
        truncated_bytes[truncated_central + 20..truncated_central + 24]
            .copy_from_slice(&(compressed_size - 1).to_le_bytes());
        let truncated_reader = ArchiveReader::new(&truncated_bytes).unwrap();
        let mut truncated_session = truncated_reader.read_session();
        assert!(truncated_session.read("truncated").is_err());
        assert_eq!(truncated_session.read("valid").unwrap(), VALID);
    }

    #[test]
    fn archive_read_session_recovers_after_store_and_unsupported_failures() {
        const VALID: &[u8] = b"valid Deflate payload after a failed member";

        let mut store_writer = StreamingArchiveWriter::new();
        store_writer
            .write_stored("bad-store", b"stored payload")
            .unwrap();
        store_writer.write_deflated_sized("valid", VALID).unwrap();
        let mut store_bytes = store_writer.finish_to_bytes().unwrap();
        let store_local = local_header_offset_for_name(&store_bytes, b"bad-store");
        let store_name_len = usize::from(u16::from_le_bytes(
            store_bytes[store_local + 26..store_local + 28]
                .try_into()
                .unwrap(),
        ));
        let store_extra_len = usize::from(u16::from_le_bytes(
            store_bytes[store_local + 28..store_local + 30]
                .try_into()
                .unwrap(),
        ));
        store_bytes[store_local + 30 + store_name_len + store_extra_len] ^= 0xff;
        let store_reader = ArchiveReader::new(&store_bytes).unwrap();
        let mut store_session = store_reader.read_session();
        let store_error = store_session.read("bad-store").unwrap_err();
        assert_materialized_checksum_error(store_error);
        assert_eq!(store_session.read("valid").unwrap(), VALID);

        let mut unsupported_writer = StreamingArchiveWriter::new();
        unsupported_writer
            .write_deflated_sized("unsupported", b"unsupported method payload")
            .unwrap();
        unsupported_writer
            .write_deflated_sized("valid", VALID)
            .unwrap();
        let mut unsupported_bytes = unsupported_writer.finish_to_bytes().unwrap();
        let unsupported_local = local_header_offset_for_name(&unsupported_bytes, b"unsupported");
        let unsupported_central =
            central_header_offset_for_name(&unsupported_bytes, b"unsupported");
        unsupported_bytes[unsupported_local + 8..unsupported_local + 10]
            .copy_from_slice(&12_u16.to_le_bytes());
        unsupported_bytes[unsupported_central + 10..unsupported_central + 12]
            .copy_from_slice(&12_u16.to_le_bytes());
        let unsupported_reader = ArchiveReader::new(&unsupported_bytes).unwrap();
        let mut unsupported_session = unsupported_reader.read_session();
        let unsupported_error = unsupported_session.read("unsupported").unwrap_err();
        assert!(matches!(
            unsupported_error.kind(),
            ErrorKind::UnsupportedCompressionMethod(12)
        ));
        assert_eq!(unsupported_session.read("valid").unwrap(), VALID);
    }

    #[test]
    fn archive_read_session_resets_after_declared_size_overrun_and_underrun() {
        const ACTUAL: &[u8] = b"payload whose declared size is intentionally wrong";
        const VALID: &[u8] = b"valid payload after a size failure";

        for declared_size in [3_u32, 128_u32] {
            let mut writer = StreamingArchiveWriter::new();
            writer.write_deflated_sized("bad-size", ACTUAL).unwrap();
            writer.write_deflated_sized("valid", VALID).unwrap();
            let mut bytes = writer.finish_to_bytes().unwrap();
            rewrite_uncompressed_size(&mut bytes, b"bad-size", declared_size);
            let reader = ArchiveReader::new(&bytes).unwrap();
            let mut session = reader.read_session();
            let error = session.read("bad-size").unwrap_err();
            assert_materialized_size_error(error);
            assert_eq!(session.read("valid").unwrap(), VALID);
        }
    }

    #[test]
    fn sized_deflated_members_declare_upfront_sizes_without_a_data_descriptor() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("mimetype", b"application/test")
            .unwrap();
        writer
            .write_deflated_sized("content.xml", b"<content>Hello</content>")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        assert!(!local_member_has_data_descriptor(&bytes, b"content.xml"));

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(reader.len(), 2);
        assert_eq!(reader.read("mimetype").unwrap(), b"application/test");
        assert_eq!(
            reader.read("content.xml").unwrap(),
            b"<content>Hello</content>"
        );
    }

    #[test]
    fn sized_deflated_limit_breach_leaves_the_writer_usable() {
        let limits = StreamingArchiveLimits::default().with_compressed_size_limit(1024);
        let mut writer = StreamingArchiveWriter::with_limits(limits);
        writer
            .write_stored("mimetype", b"application/test")
            .unwrap();
        // xorshift64* output is effectively incompressible, so the 4 KiB
        // payload must exceed the 1 KiB compressed-size budget.
        let mut state = 0x1234_5678_9abc_def0u64;
        let payload: Vec<u8> = (0..4096)
            .map(|_| {
                state ^= state << 13;
                state ^= state >> 7;
                state ^= state << 17;
                (state >> 32) as u8
            })
            .collect();
        let error = writer
            .write_deflated_sized("content.xml", &payload)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::LimitExceeded { .. }));
        assert!(!writer.is_poisoned());
        // The archive never observed the rejected member, so a follow-up
        // in-budget write still succeeds.
        writer.write_deflated_sized("ok.xml", b"<ok/>").unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(reader.read("ok.xml").unwrap(), b"<ok/>");
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
    fn indexed_read_session_reuses_deflate_state_with_accounting_parity() {
        let stored_first = b"stored first descriptor-free payload";
        let deflated_first = b"deflated first payload with a data descriptor";
        let stored_second = b"stored second payload with a data descriptor";
        let deflated_second = b"deflated second payload with a data descriptor";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored-first", stored_first).unwrap();
        writer
            .write_deflated("deflated-first", deflated_first)
            .unwrap();
        writer
            .write_stored_stream("stored-second", Cursor::new(stored_second.to_vec()))
            .unwrap();
        writer
            .write_deflated("deflated-second", deflated_second)
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        assert!(local_member_has_data_descriptor(&bytes, b"deflated-first"));
        assert!(local_member_has_data_descriptor(&bytes, b"stored-second"));
        assert!(local_member_has_data_descriptor(&bytes, b"deflated-second"));

        let archive = indexed_archive(bytes);
        let members = [
            ("stored-first", stored_first.as_slice()),
            ("deflated-first", deflated_first.as_slice()),
            ("stored-second", stored_second.as_slice()),
            ("deflated-second", deflated_second.as_slice()),
        ];
        let mut expected = Vec::new();
        for (name, payload) in members {
            let entry_id = archive.entry_id(name).unwrap();
            let mut accounting = ZipOperationAccounting::default();
            let decoded = archive
                .read_entry_with_accounting(entry_id, &mut accounting)
                .unwrap();
            assert_eq!(decoded, payload);
            expected.push((decoded, accounting));
        }

        let mut session = archive.read_session();
        for ((name, payload), (expected_payload, expected_accounting)) in
            members.into_iter().zip(expected)
        {
            let entry_id = archive.entry_id(name).unwrap();
            let mut accounting = ZipOperationAccounting::default();
            let decoded = session
                .read_entry_with_accounting(entry_id, &mut accounting)
                .unwrap();
            assert_eq!(decoded, payload);
            assert_eq!(decoded, expected_payload);
            assert_eq!(accounting, expected_accounting);
        }
    }

    #[test]
    fn indexed_read_session_reads_by_name_exactly_as_the_one_shot_read() {
        const VALID: &[u8] = b"valid payload read after a refused member";
        let stored = b"stored payload that bypasses the decoder";
        let deflated = b"deflated payload that resets the decoder";

        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", stored).unwrap();
        writer
            .write_deflated_sized("bad-crc.bin", b"payload with a bad CRC")
            .unwrap();
        writer
            .write_deflated_sized("deflated.bin", deflated)
            .unwrap();
        writer.write_deflated_sized("valid.bin", VALID).unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let central = central_header_offset_for_name(&bytes, b"bad-crc.bin");
        let crc = u32::from_le_bytes(bytes[central + 16..central + 20].try_into().unwrap());
        bytes[central + 16..central + 20].copy_from_slice(&(crc ^ 1).to_le_bytes());
        let archive = indexed_archive(bytes);

        let names: Vec<String> = archive.file_names().map(str::to_string).collect();
        assert_eq!(
            names,
            ["stored.bin", "bad-crc.bin", "deflated.bin", "valid.bin"]
        );

        // One session reads every member in turn. Payload bytes, payload
        // accounting and the refusal identity are the one-shot read's.
        let mut session = archive.read_session();
        for name in &names {
            let mut one_shot_accounting = ZipOperationAccounting::default();
            let one_shot = archive.read_with_accounting(name, &mut one_shot_accounting);
            let mut session_accounting = ZipOperationAccounting::default();
            let through_session = session.read_with_accounting(name, &mut session_accounting);
            match (one_shot, through_session) {
                (Ok(expected), Ok(actual)) => {
                    assert_eq!(actual, expected, "member {name}");
                    assert_eq!(session_accounting, one_shot_accounting, "member {name}");
                },
                (Err(expected), Err(actual)) => {
                    assert_eq!(actual.to_string(), expected.to_string(), "member {name}");
                },
                (expected, actual) => {
                    panic!("member {name} diverged: {expected:?} then {actual:?}")
                },
            }
        }

        // The refused member left no state behind: every later read through
        // the same session still decodes, in any order.
        assert_eq!(session.read("deflated.bin").unwrap(), deflated);
        assert_eq!(session.read("valid.bin").unwrap(), VALID);
        assert_eq!(session.read("stored.bin").unwrap(), stored);
        assert_materialized_checksum_error(session.read("bad-crc.bin").unwrap_err());
        assert_eq!(session.read("valid.bin").unwrap(), VALID);

        // Name admission is the archive's, including the explicit-directory
        // rejection and the `FileNotFound` identity.
        for absent in ["absent.bin", "deflated.bin/", "stored.bin\\"] {
            let expected = archive.read(absent).unwrap_err();
            assert!(
                matches!(expected.kind(), ErrorKind::FileNotFound(_)),
                "{absent}"
            );
            assert_eq!(
                session.read(absent).unwrap_err().to_string(),
                expected.to_string(),
                "{absent}"
            );
        }
    }

    #[test]
    fn indexed_read_session_resets_after_crc_and_corrupt_deflate_failures() {
        const VALID: &[u8] = b"valid payload after a failed Deflate read";

        let mut crc_writer = StreamingArchiveWriter::new();
        crc_writer
            .write_deflated_sized("bad-crc", b"payload with a bad CRC")
            .unwrap();
        crc_writer.write_deflated_sized("valid", VALID).unwrap();
        let mut crc_bytes = crc_writer.finish_to_bytes().unwrap();
        let crc_central = central_header_offset_for_name(&crc_bytes, b"bad-crc");
        let crc = u32::from_le_bytes(
            crc_bytes[crc_central + 16..crc_central + 20]
                .try_into()
                .unwrap(),
        );
        crc_bytes[crc_central + 16..crc_central + 20].copy_from_slice(&(crc ^ 1).to_le_bytes());
        let crc_archive = indexed_archive(crc_bytes);
        let bad_crc = crc_archive.entry_id("bad-crc").unwrap();
        let valid = crc_archive.entry_id("valid").unwrap();
        let mut session = crc_archive.read_session();
        let error = session.read_entry(bad_crc).unwrap_err();
        assert_materialized_checksum_error(error);
        assert_eq!(session.read_entry(valid).unwrap(), VALID);

        let mut corrupt_writer = StreamingArchiveWriter::new();
        corrupt_writer
            .write_deflated_sized("corrupt", b"payload with corrupt Deflate bytes")
            .unwrap();
        corrupt_writer.write_deflated_sized("valid", VALID).unwrap();
        let mut corrupt_bytes = corrupt_writer.finish_to_bytes().unwrap();
        let corrupt_local = local_header_offset_for_name(&corrupt_bytes, b"corrupt");
        let name_len = usize::from(u16::from_le_bytes(
            corrupt_bytes[corrupt_local + 26..corrupt_local + 28]
                .try_into()
                .unwrap(),
        ));
        let extra_len = usize::from(u16::from_le_bytes(
            corrupt_bytes[corrupt_local + 28..corrupt_local + 30]
                .try_into()
                .unwrap(),
        ));
        let compressed_start = corrupt_local + 30 + name_len + extra_len;
        corrupt_bytes[compressed_start] ^= 0xff;
        let corrupt_archive = indexed_archive(corrupt_bytes);
        let corrupt = corrupt_archive.entry_id("corrupt").unwrap();
        let valid = corrupt_archive.entry_id("valid").unwrap();
        let mut session = corrupt_archive.read_session();
        assert!(session.read_entry(corrupt).is_err());
        assert_eq!(session.read_entry(valid).unwrap(), VALID);

        let mut truncated_writer = StreamingArchiveWriter::new();
        truncated_writer
            .write_deflated_sized("truncated", b"payload with a truncated Deflate stream")
            .unwrap();
        truncated_writer
            .write_deflated_sized("valid", VALID)
            .unwrap();
        let mut truncated_bytes = truncated_writer.finish_to_bytes().unwrap();
        let truncated_local = local_header_offset_for_name(&truncated_bytes, b"truncated");
        let truncated_central = central_header_offset_for_name(&truncated_bytes, b"truncated");
        let compressed_size = u32::from_le_bytes(
            truncated_bytes[truncated_local + 18..truncated_local + 22]
                .try_into()
                .unwrap(),
        );
        assert!(compressed_size > 1);
        truncated_bytes[truncated_local + 18..truncated_local + 22]
            .copy_from_slice(&(compressed_size - 1).to_le_bytes());
        truncated_bytes[truncated_central + 20..truncated_central + 24]
            .copy_from_slice(&(compressed_size - 1).to_le_bytes());
        let truncated_archive = indexed_archive(truncated_bytes);
        let truncated = truncated_archive.entry_id("truncated").unwrap();
        let valid = truncated_archive.entry_id("valid").unwrap();
        let mut session = truncated_archive.read_session();
        assert!(session.read_entry(truncated).is_err());
        assert_eq!(session.read_entry(valid).unwrap(), VALID);
    }

    #[test]
    fn indexed_read_session_resets_after_declared_size_overrun_and_underrun() {
        const ACTUAL: &[u8] = b"payload whose declared size is intentionally wrong";
        const VALID: &[u8] = b"valid payload after a size failure";

        for declared_size in [3_u32, 128_u32] {
            let mut writer = StreamingArchiveWriter::new();
            writer.write_deflated_sized("bad-size", ACTUAL).unwrap();
            writer.write_deflated_sized("valid", VALID).unwrap();
            let mut bytes = writer.finish_to_bytes().unwrap();
            rewrite_uncompressed_size(&mut bytes, b"bad-size", declared_size);
            let archive = indexed_archive(bytes);
            let bad_size = archive.entry_id("bad-size").unwrap();
            let valid = archive.entry_id("valid").unwrap();
            let mut session = archive.read_session();
            let error = session.read_entry(bad_size).unwrap_err();
            assert_materialized_size_error(error);
            assert_eq!(session.read_entry(valid).unwrap(), VALID);
        }
    }

    #[test]
    fn indexed_read_to_reports_typed_crc_failure_for_data_descriptor() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored_stream("bad.bin", Cursor::new(b"bad".to_vec()))
            .unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        assert!(local_member_has_data_descriptor(&bytes, b"bad.bin"));
        corrupt_payload(&mut bytes, b"bad");

        let archive = indexed_archive(bytes);
        let entry_id = archive.entry_id("bad.bin").unwrap();
        let mut output = Vec::new();
        let error = archive.read_entry_to(entry_id, &mut output).unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidChecksum { .. }));
        assert_eq!(output, vec![b'b' ^ 0x80, b'a', b'd']);
    }

    #[test]
    fn read_to_rejects_truncated_and_overlong_store_and_deflate_members() {
        for deflated in [false, true] {
            for (payload, declared_size) in
                [(b"abc".as_slice(), 4_u32), (b"abcde".as_slice(), 4_u32)]
            {
                let mut writer = StreamingArchiveWriter::new();
                if deflated {
                    writer.write_deflated("member.bin", payload).unwrap();
                } else {
                    writer.write_stored("member.bin", payload).unwrap();
                }
                let mut bytes = writer.finish_to_bytes().unwrap();
                rewrite_uncompressed_size(&mut bytes, b"member.bin", declared_size);

                let reader = ArchiveReader::new(&bytes).unwrap();
                let mut output = Vec::new();
                let error = reader.read_to("member.bin", &mut output).unwrap_err();
                assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));

                let archive = indexed_archive(bytes);
                let mut output = Vec::new();
                let error = archive.read_to("member.bin", &mut output).unwrap_err();
                assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));
            }
        }
    }

    #[test]
    fn materialized_reads_reject_truncated_and_overlong_store_and_deflate_members() {
        for deflated in [false, true] {
            for (payload, declared_size) in
                [(b"abc".as_slice(), 4_u32), (b"abcde".as_slice(), 4_u32)]
            {
                let mut writer = StreamingArchiveWriter::new();
                if deflated {
                    writer.write_deflated("member.bin", payload).unwrap();
                } else {
                    writer.write_stored("member.bin", payload).unwrap();
                }
                let mut bytes = writer.finish_to_bytes().unwrap();
                rewrite_uncompressed_size(&mut bytes, b"member.bin", declared_size);

                {
                    let reader =
                        ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).unwrap();
                    let error = reader.read("member.bin").unwrap_err();
                    assert_materialized_size_error(error);
                }

                let archive = indexed_archive_result(bytes, ArchiveLimits::UNBOUNDED).unwrap();
                let entry_id = archive.entry_id("member.bin").unwrap();
                let error = archive.read_entry(entry_id).unwrap_err();
                assert_materialized_size_error(error);
            }
        }
    }

    #[test]
    fn indexed_read_to_supports_zip64_member_metadata_and_short_writes() {
        let bytes = include_bytes!("../assets/zip64.zip").to_vec();
        let archive = indexed_archive_result(bytes, ArchiveLimits::UNBOUNDED).unwrap();
        let mut output = ShortWriter::new(3);
        let count = archive.read_to("README", &mut output).unwrap();
        assert_eq!(count, 36);
        assert_eq!(output.bytes, b"This small file is in ZIP64 format.\n");
    }

    #[test]
    fn borrowed_store_defers_zip64_eocd_to_owned_fallback() {
        let bytes = include_bytes!("../assets/zip64.zip");
        let archive = ZipArchive::from_slice(bytes).unwrap();
        assert!(archive.is_zip64());

        let reader = ArchiveReader::new(bytes).unwrap();
        assert_eq!(reader.read_stored_borrowed("README").unwrap(), None);
        assert_eq!(
            reader.read("README").unwrap(),
            b"This small file is in ZIP64 format.\n"
        );
    }

    #[test]
    fn prefixed_zip64_uses_locator_resolved_central_size() {
        let source = include_bytes!("../assets/zip64.zip");
        let mut bytes = vec![0xa5; 7];
        bytes.extend_from_slice(source);

        let archive = ZipArchive::from_slice(bytes.as_slice()).unwrap();
        assert!(archive.is_zip64());
        let central_end = archive
            .directory_offset()
            .checked_add(archive.central_directory_size())
            .unwrap();
        assert!(central_end <= archive.eocd_offset());

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(
            reader.read("README").unwrap(),
            b"This small file is in ZIP64 format.\n"
        );
    }

    #[test]
    fn indexed_read_to_reports_source_and_zero_progress_sink_errors() {
        let payload = vec![b'x'; STREAM_COPY_BUFFER_SIZE + 1];
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("payload.bin", &payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let zip = ZipArchive::from_slice(&bytes).unwrap();
        let wayfinder = zip.entries().next().unwrap().unwrap().wayfinder();
        let entry = zip.get_entry(wayfinder).unwrap();
        let (fail_start, fail_end) = entry.compressed_data_range();
        let archive = IndexedArchive::from_reader_with_limits(
            FailOnRangeReaderAt {
                bytes: bytes.clone(),
                fail_start,
                fail_end,
            },
            bytes.len() as u64,
            ArchiveLimits::default(),
        )
        .unwrap();
        let mut output = Vec::new();
        let error = archive.read_to("payload.bin", &mut output).unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
        assert!(output.is_empty());

        let reader = ArchiveReader::new(&bytes).unwrap();
        let mut zero = ZeroWriter;
        let error = reader.read_to("payload.bin", &mut zero).unwrap_err();
        match error.kind() {
            ErrorKind::IO(error) | ErrorKind::Io(error) => {
                assert_eq!(error.kind(), io::ErrorKind::WriteZero);
            },
            other => panic!("expected zero-progress sink failure, got {other:?}"),
        }
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
        exact.max_metadata_bytes = 5 + CENTRAL_FIXED_RECORD_BYTES;
        assert!(ArchiveReader::new_with_limits(&bytes, exact).is_ok());
        assert!(indexed_archive_result(bytes.clone(), exact).is_ok());

        let mut over = exact;
        over.max_metadata_bytes -= 1;
        assert_limit(
            ArchiveReader::new_with_limits(&bytes, over).unwrap_err(),
            LimitResource::MetadataBytes,
            5 + CENTRAL_FIXED_RECORD_BYTES,
            4 + CENTRAL_FIXED_RECORD_BYTES,
        );
        assert_limit(
            indexed_archive_result(bytes, over).unwrap_err(),
            LimitResource::MetadataBytes,
            5 + CENTRAL_FIXED_RECORD_BYTES,
            4 + CENTRAL_FIXED_RECORD_BYTES,
        );
    }

    #[test]
    fn positional_index_handles_central_records_larger_than_recommended_scratch() {
        let (bytes, metadata_bytes) = oversized_metadata_fixture();

        // The central record is valid but its variable fields exceed the
        // public 64 KiB recommendation. The positional iterator must spill
        // only this record and continue to expose its borrowed metadata.
        let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipArchive::from_seekable(Cursor::new(bytes.clone()), &mut scratch)
            .expect("locate archive with the recommended scratch size");
        let mut entries = archive.entries(&mut scratch);
        let entry = entries
            .next_entry()
            .expect("read oversized central record")
            .expect("record exists");
        assert_eq!(entry.file_path().as_ref().len(), 4 * 1024);
        assert_eq!(entry.metadata_size_hint(), metadata_bytes);
        assert!(entries.next_entry().expect("finish directory").is_none());

        let indexed = indexed_archive_result(bytes.clone(), ArchiveLimits::default())
            .expect("high-level positional index accepts valid metadata");
        assert_eq!(indexed.len(), 1);

        let mut constrained = ArchiveLimits::UNBOUNDED;
        constrained.max_metadata_bytes = metadata_bytes - 1;
        assert_limit(
            indexed_archive_result(bytes, constrained).unwrap_err(),
            LimitResource::MetadataBytes,
            metadata_bytes + CENTRAL_FIXED_RECORD_BYTES,
            metadata_bytes - 1,
        );
    }

    #[test]
    fn positional_iterator_continues_from_oversized_record_to_ordinary_record() {
        let (bytes, metadata_bytes) = oversized_and_ordinary_fixture(true);
        let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipArchive::from_seekable(Cursor::new(bytes), &mut scratch)
            .expect("locate archive with the recommended scratch size");
        let mut entries = archive.entries(&mut scratch);

        {
            let entry = entries
                .next_entry()
                .expect("read oversized central record")
                .expect("oversized record exists");
            assert_eq!(entry.file_path().as_ref().len(), 4 * 1024);
            assert_eq!(entry.metadata_size_hint(), metadata_bytes);
        }

        let ordinary = entries
            .next_entry()
            .expect("continue after oversized record")
            .expect("ordinary record exists");
        assert_eq!(ordinary.file_path().as_ref(), b"ordinary");
        assert!(entries.next_entry().expect("finish directory").is_none());
    }

    #[test]
    fn positional_iterator_continues_from_ordinary_record_to_oversized_record() {
        let (bytes, metadata_bytes) = oversized_and_ordinary_fixture(false);
        let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipArchive::from_seekable(Cursor::new(bytes), &mut scratch)
            .expect("locate archive with the recommended scratch size");
        let mut entries = archive.entries(&mut scratch);

        let ordinary = entries
            .next_entry()
            .expect("read ordinary central record")
            .expect("ordinary record exists");
        assert_eq!(ordinary.file_path().as_ref(), b"ordinary");

        let oversized = entries
            .next_entry()
            .expect("continue to oversized record")
            .expect("oversized record exists");
        assert_eq!(oversized.file_path().as_ref().len(), 4 * 1024);
        assert_eq!(oversized.metadata_size_hint(), metadata_bytes);
        assert!(entries.next_entry().expect("finish directory").is_none());
    }

    #[test]
    fn prefixed_positional_archive_applies_base_offset_to_oversized_records() {
        let (archive_bytes, metadata_bytes) = oversized_metadata_fixture();
        let prefix = vec![0xa5; 7];
        let prefix_len = prefix.len() as u64;
        let mut bytes = prefix.clone();
        bytes.extend_from_slice(&archive_bytes);

        let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipArchive::from_seekable(Cursor::new(bytes), &mut scratch)
            .expect("locate prefixed archive");
        let mut entries = archive.entries(&mut scratch);
        let entry = entries
            .next_entry()
            .expect("read prefixed oversized record")
            .expect("record exists");
        assert_eq!(entry.metadata_size_hint(), metadata_bytes);
        assert_eq!(entry.local_header_offset(), prefix_len);
    }

    #[test]
    fn truncated_oversized_variable_section_returns_typed_eof() {
        let (bytes, metadata_bytes) = oversized_metadata_fixture();
        let located = ZipArchive::from_slice(&bytes).expect("valid source fixture");
        let central_offset = usize::try_from(located.directory_offset()).unwrap();
        let eocd_offset = usize::try_from(located.eocd_offset()).unwrap();
        let central_size = eocd_offset - central_offset;
        let missing = 1024;
        assert!(metadata_bytes > missing as u64);

        let mut truncated = bytes[..eocd_offset - missing].to_vec();
        truncated.extend_from_slice(&bytes[eocd_offset..]);
        let truncated_eocd_offset = eocd_offset - missing;
        let truncated_central_size = u32::try_from(central_size - missing).unwrap();
        truncated[truncated_eocd_offset + 12..truncated_eocd_offset + 16]
            .copy_from_slice(&truncated_central_size.to_le_bytes());

        let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
        let archive = ZipArchive::from_seekable(Cursor::new(truncated), &mut scratch)
            .expect("locate truncated archive for structural iteration");
        let mut entries = archive.entries(&mut scratch);
        let error = entries
            .next_entry()
            .expect_err("truncated oversized metadata must fail");
        assert!(matches!(error.kind(), ErrorKind::Eof));
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
        let directory_extra = zip64_sizes(0, 0);
        let bytes = fixture(&[FixtureEntry {
            name: b"folder/",
            extra: &directory_extra,
            comment: b"",
            compressed_size: u32::MAX,
            uncompressed_size: u32::MAX,
            data: b"",
        }]);
        let limits = ArchiveLimits {
            max_files: 0,
            max_member_name_bytes: 7,
            max_metadata_bytes: 7 + CENTRAL_FIXED_RECORD_BYTES + directory_extra.len() as u64,
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
        let directory_extra = zip64_sizes(u64::from(u32::MAX), u64::from(u32::MAX));
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
                extra: &directory_extra,
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
    fn lazy_shared_concurrent_cold_reads_share_one_flight() {
        const CALLERS: usize = 8;
        let payload = vec![b'x'; 1024 * 1024];
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("payload", &payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let reader = Arc::new(LazyArchiveReader::new(&bytes).unwrap());
        let ready = Arc::new(std::sync::Barrier::new(CALLERS + 1));
        let values = std::thread::scope(|scope| {
            let handles = (0..CALLERS)
                .map(|_| {
                    let reader = Arc::clone(&reader);
                    let ready = Arc::clone(&ready);
                    scope.spawn(move || {
                        ready.wait();
                        reader.read_shared("payload")
                    })
                })
                .collect::<Vec<_>>();

            ready.wait();
            handles
                .into_iter()
                .map(|handle| handle.join().expect("lazy reader worker should not panic"))
                .collect::<Result<Vec<_>, _>>()
        })
        .expect("concurrent cold reads should succeed");
        assert!(
            values
                .windows(2)
                .all(|pair| Arc::ptr_eq(&pair[0], &pair[1]))
        );
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), 1);
        assert_eq!(reader.cache_size(), 1);
        assert_eq!(reader.cache_bytes(), payload.len());
    }

    #[test]
    fn lazy_shared_failed_flight_wakes_waiters_and_allows_retry() {
        const CALLERS: usize = 6;
        let mut bytes = bulk_fixture();
        corrupt_payload(&mut bytes, b"bad");
        let reader = Arc::new(LazyArchiveReader::new(&bytes).unwrap());
        let ready = Arc::new(std::sync::Barrier::new(CALLERS + 1));
        std::thread::scope(|scope| {
            let handles = (0..CALLERS)
                .map(|_| {
                    let reader = Arc::clone(&reader);
                    let ready = Arc::clone(&ready);
                    scope.spawn(move || {
                        ready.wait();
                        reader.read_shared("bad")
                    })
                })
                .collect::<Vec<_>>();

            ready.wait();
            for handle in handles {
                let result = handle.join().expect("lazy reader worker should not panic");
                assert!(matches!(
                    result,
                    Err(error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. })
                ));
            }
        });
        assert_eq!(reader.cache_size(), 0);
        let first_failure_count = reader.cold_loads.load(Ordering::SeqCst);

        let retry = reader.read_shared("bad");
        assert!(matches!(
            retry,
            Err(error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. })
        ));
        assert_eq!(
            reader.cold_loads.load(Ordering::SeqCst),
            first_failure_count + 1
        );
        assert_eq!(reader.cache_size(), 0);
    }

    fn register_lazy_test_flight(reader: &LazyArchiveReader<'_>, name: &str) -> Arc<LazyFlight> {
        let mut cache = lock_lazy_cache(&reader.cache);
        let flight = Arc::new(LazyFlight::new(cache.generation, name.len()));
        assert!(
            cache
                .flights
                .insert(name.to_string(), Arc::clone(&flight))
                .is_none()
        );
        cache.active_flights += 1;
        cache.active_key_bytes += name.len();
        flight
    }

    #[test]
    fn lazy_active_flight_limit_bypasses_and_recovers() {
        let bytes = fixture(&[FixtureEntry::stored(b"one", b"one")]);
        let cache_limits =
            LazyArchiveCacheLimits::new_with_active_flight_limits(64, 2, 1, 5).unwrap();
        let reader = LazyArchiveReader::new_with_cache_limits(&bytes, cache_limits).unwrap();
        let held = register_lazy_test_flight(&reader, "hold");
        assert_eq!(reader.active_flight_count(), 1);
        assert_eq!(reader.active_flight_key_bytes(), 4);

        // Both the active-flight count and aggregate key-byte budget are full,
        // so this distinct member is read directly without another flight.
        assert_eq!(reader.read_shared("one").unwrap().as_slice(), b"one");
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), 0);
        assert_eq!(reader.cache_size(), 0);

        held.complete_failure();
        {
            let mut cache = lock_lazy_cache(&reader.cache);
            finish_lazy_flight(&mut cache, "hold", &held);
        }
        assert_eq!(reader.active_flight_count(), 0);
        assert_eq!(reader.active_flight_key_bytes(), 0);

        assert_eq!(reader.read_shared("one").unwrap().as_slice(), b"one");
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), 1);
        assert_eq!(reader.cache_size(), 1);
    }

    #[test]
    fn lazy_flight_key_limit_is_typed_and_does_not_retain_state() {
        let bytes = fixture(&[FixtureEntry::stored(b"one", b"one")]);
        let cache_limits =
            LazyArchiveCacheLimits::new_with_active_flight_limits(64, 2, 2, 4).unwrap();
        let reader = LazyArchiveReader::new_with_cache_limits(&bytes, cache_limits).unwrap();

        let result = reader.read_shared("fives");
        assert!(matches!(
            result,
            Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { .. })
        ));
        assert_eq!(reader.active_flight_count(), 0);
        assert_eq!(reader.active_flight_key_bytes(), 0);
        assert_eq!(reader.cache_size(), 0);
    }

    #[test]
    fn lazy_missing_name_flight_wakes_waiter_and_releases_budget() {
        let bytes = fixture(&[FixtureEntry::stored(b"one", b"one")]);
        let cache_limits =
            LazyArchiveCacheLimits::new_with_active_flight_limits(64, 2, 1, 64).unwrap();
        let reader =
            Arc::new(LazyArchiveReader::new_with_cache_limits(&bytes, cache_limits).unwrap());
        let held = register_lazy_test_flight(&reader, "missing");

        let waited = std::thread::scope(|scope| {
            let waiter_reader = Arc::clone(&reader);
            let waiter = scope.spawn(move || waiter_reader.read_shared("missing"));
            // The cache map and this test own two references already. The
            // third reference proves the waiter has captured the flight.
            while Arc::strong_count(&held) < 3 {
                std::thread::yield_now();
            }

            held.complete_failure();
            {
                let mut cache = lock_lazy_cache(&reader.cache);
                finish_lazy_flight(&mut cache, "missing", &held);
            }
            waiter.join().expect("missing-name waiter should not panic")
        });
        assert!(matches!(
            waited,
            Err(error) if matches!(error.kind(), ErrorKind::FileNotFound(_))
        ));
        assert_eq!(reader.active_flight_count(), 0);
        assert_eq!(reader.active_flight_key_bytes(), 0);
        assert_eq!(reader.cache_size(), 0);
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), 1);

        let retry = reader.read_shared("missing");
        assert!(matches!(
            retry,
            Err(error) if matches!(error.kind(), ErrorKind::FileNotFound(_))
        ));
        assert_eq!(reader.active_flight_count(), 0);
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), 2);
    }

    #[test]
    fn lazy_cache_limits_bound_weight_and_entries_with_exact_lru() {
        let bytes = fixture(&[
            FixtureEntry::stored(b"first", b"one"),
            FixtureEntry::stored(b"second", b"two"),
            FixtureEntry::stored(b"third", b"tre"),
        ]);
        let cache_limits = LazyArchiveCacheLimits::new(6, 2).unwrap();
        let reader = LazyArchiveReader::new_with_cache_limits(&bytes, cache_limits).unwrap();

        assert_eq!(reader.cache_limits(), cache_limits);
        assert_eq!(reader.read_shared("first").unwrap().as_slice(), b"one");
        assert_eq!(reader.read_shared("second").unwrap().as_slice(), b"two");
        assert_eq!(reader.cache_size(), 2);
        assert_eq!(reader.cache_bytes(), 6);

        // Touch `first`, then insert `third`: exact LRU must evict `second`.
        assert_eq!(reader.read_shared("first").unwrap().as_slice(), b"one");
        assert_eq!(reader.read_shared("third").unwrap().as_slice(), b"tre");
        assert_eq!(reader.cache_size(), 2);
        assert_eq!(reader.cache_bytes(), 6);

        let cold_loads = reader.cold_loads.load(Ordering::SeqCst);
        assert_eq!(reader.read_shared("second").unwrap().as_slice(), b"two");
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), cold_loads + 1);
        assert_eq!(reader.cache_size(), 2);
        assert_eq!(reader.cache_bytes(), 6);
    }

    #[test]
    fn lazy_read_returns_fresh_vec_while_shared_reads_preserve_arc() {
        let bytes = fixture(&[FixtureEntry::stored(b"payload", b"body")]);
        let reader = LazyArchiveReader::new(&bytes).unwrap();

        let mut first = reader.read("payload").unwrap();
        first[0] = b'X';
        assert_eq!(reader.read("payload").unwrap(), b"body");

        let shared_first = reader.read_shared("payload").unwrap();
        let shared_second = reader.read_shared("payload").unwrap();
        assert!(Arc::ptr_eq(&shared_first, &shared_second));
    }

    #[test]
    fn lazy_serial_shared_batch_reuses_session_and_preserves_cache_identity() {
        let stored = b"stored payload";
        let first = b"first Deflate payload";
        let second = b"second Deflate payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored", stored).unwrap();
        writer.write_deflated("first", first).unwrap();
        writer.write_deflated("second", second).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let reader = LazyArchiveReader::new(&bytes).unwrap();
        let names = ["second", "stored", "first", "second", "first"];
        let expected = [
            second.as_slice(),
            stored.as_slice(),
            first.as_slice(),
            second.as_slice(),
            first.as_slice(),
        ];

        let first_batch = reader.read_many_serial_shared(&names);
        assert_eq!(first_batch.len(), names.len());
        for (result, expected) in first_batch.iter().zip(expected) {
            assert_eq!(result.1.as_ref().unwrap().as_slice(), expected);
        }
        assert_eq!(reader.cache_size(), 3);
        let cold_loads = reader.cold_loads.load(Ordering::SeqCst);
        assert_eq!(cold_loads, 3);

        let second_batch = reader.read_many_serial_shared(&names);
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), cold_loads);
        for ((name, first_result), (_, second_result)) in first_batch.into_iter().zip(second_batch)
        {
            let first_arc = first_result.unwrap();
            let second_arc = second_result.unwrap();
            assert!(Arc::ptr_eq(&first_arc, &second_arc));
            let cached = reader.read_shared(name).unwrap();
            assert!(Arc::ptr_eq(&second_arc, &cached));
        }
    }

    #[test]
    fn lazy_serial_shared_batch_recovers_after_a_failed_cold_member() {
        const VALID: &[u8] = b"valid payload after a failed serial batch member";
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_deflated_sized("bad", b"payload with corrupt Deflate bytes")
            .unwrap();
        writer.write_deflated_sized("valid", VALID).unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let local = local_header_offset_for_name(&bytes, b"bad");
        let name_len = usize::from(u16::from_le_bytes(
            bytes[local + 26..local + 28].try_into().unwrap(),
        ));
        let extra_len = usize::from(u16::from_le_bytes(
            bytes[local + 28..local + 30].try_into().unwrap(),
        ));
        bytes[local + 30 + name_len + extra_len] ^= 0xff;

        let reader = LazyArchiveReader::new(&bytes).unwrap();
        let names = ["bad", "valid"];
        let first = reader.read_many_serial_shared(&names);
        assert!(first[0].1.is_err());
        assert_eq!(first[1].1.as_ref().unwrap().as_slice(), VALID);
        assert_eq!(reader.cache_size(), 1);
        let cold_loads = reader.cold_loads.load(Ordering::SeqCst);
        assert_eq!(cold_loads, 2);

        let retry = reader.read_many_serial_shared(&["bad"]);
        assert!(retry[0].1.is_err());
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), cold_loads + 1);
        assert_eq!(reader.read_shared("valid").unwrap().as_slice(), VALID);
        assert_eq!(reader.active_flight_count(), 0);
    }

    #[test]
    fn lazy_serial_shared_batch_bypasses_full_flight_budget_and_recovers() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_deflated("payload", b"serial bypass payload")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let cache_limits =
            LazyArchiveCacheLimits::new_with_active_flight_limits(1024, 4, 1, 64).unwrap();
        let reader = LazyArchiveReader::new_with_cache_limits(&bytes, cache_limits).unwrap();
        let held = register_lazy_test_flight(&reader, "held");

        let bypassed = reader.read_many_serial_shared(&["payload"]);
        assert_eq!(bypassed.len(), 1);
        assert_eq!(
            bypassed[0].1.as_ref().unwrap().as_slice(),
            b"serial bypass payload"
        );
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), 0);
        assert_eq!(reader.cache_size(), 0);
        assert_eq!(reader.active_flight_count(), 1);

        held.complete_failure();
        {
            let mut cache = lock_lazy_cache(&reader.cache);
            finish_lazy_flight(&mut cache, "held", &held);
        }
        let loaded = reader.read_many_serial_shared(&["payload"]);
        assert_eq!(
            loaded[0].1.as_ref().unwrap().as_slice(),
            b"serial bypass payload"
        );
        assert_eq!(reader.cold_loads.load(Ordering::SeqCst), 1);
        assert_eq!(reader.cache_size(), 1);
        assert_eq!(reader.active_flight_count(), 0);
    }

    #[test]
    fn lazy_cache_limits_reject_zero_capacity() {
        assert!(matches!(
            LazyArchiveCacheLimits::new(0, 1),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { .. })
        ));
        assert!(matches!(
            LazyArchiveCacheLimits::new(1, 0),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { .. })
        ));
        assert!(matches!(
            LazyArchiveCacheLimits::new_with_active_flight_limits(1, 1, 0, 1),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { .. })
        ));
        assert!(matches!(
            LazyArchiveCacheLimits::new_with_active_flight_limits(1, 1, 1, 0),
            Err(error) if matches!(error.kind(), ErrorKind::InvalidInput { .. })
        ));
    }

    #[test]
    fn lazy_cache_skips_oversized_and_externally_pinned_payloads() {
        let bytes = fixture(&[
            FixtureEntry::stored(b"first", b"one"),
            FixtureEntry::stored(b"second", b"two"),
            FixtureEntry::stored(b"third", b"tre"),
            FixtureEntry::stored(b"large", b"larger!!"),
        ]);
        let cache_limits = LazyArchiveCacheLimits::new(6, 2).unwrap();
        let reader = LazyArchiveReader::new_with_cache_limits(&bytes, cache_limits).unwrap();

        let first = reader.read_shared("first").unwrap();
        let second = reader.read_shared("second").unwrap();
        let _ = reader.read_shared("third").unwrap();
        assert_eq!(reader.cache_size(), 2);
        assert_eq!(reader.cache_bytes(), 6);
        assert_eq!(first.as_slice(), b"one");
        assert_eq!(second.as_slice(), b"two");

        drop(first);
        let third = reader.read_shared("third").unwrap();
        assert_eq!(third.as_slice(), b"tre");
        assert_eq!(reader.cache_size(), 2);
        assert_eq!(reader.cache_bytes(), 6);

        let large = reader.read_shared("large").unwrap();
        assert_eq!(large.as_slice(), b"larger!!");
        assert_eq!(reader.cache_size(), 2);
        assert_eq!(reader.cache_bytes(), 6);
        drop(large);
    }

    #[test]
    fn lazy_cache_uses_canonical_member_names_and_clear_fences_flights() {
        let bytes = fixture(&[FixtureEntry::stored(b"dir/../body.xml", b"body")]);
        let reader =
            Arc::new(LazyArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED).unwrap());
        let first = reader.read_shared("/./dir/../body.xml").unwrap();
        let second = reader.read_shared("body.xml").unwrap();
        assert!(Arc::ptr_eq(&first, &second));
        assert_eq!(reader.cache_size(), 1);
        let initial_cold_loads = reader.cold_loads.load(Ordering::SeqCst);

        drop(first);
        drop(second);
        reader.clear_cache();

        let old = register_lazy_test_flight(&reader, "body.xml");
        let (fresh, waited, active_after_clear, flights_empty, cold_loads) =
            std::thread::scope(|scope| {
                let waiter_reader = Arc::clone(&reader);
                let waiter = scope.spawn(move || waiter_reader.read_shared("body.xml"));
                while Arc::strong_count(&old) < 3 {
                    std::thread::yield_now();
                }

                // Clearing detaches the old flight. A post-clear read must create
                // a new generation flight rather than join this one.
                reader.clear_cache();
                let active_after_clear = reader.active_flight_count();
                let flights_empty = lock_lazy_cache(&reader.cache).flights.is_empty();
                let fresh = reader.read_shared("body.xml");
                let cold_loads = reader.cold_loads.load(Ordering::SeqCst);

                let stale = Arc::new(b"stale".to_vec());
                {
                    let mut cache = lock_lazy_cache(&reader.cache);
                    if cache.generation == old.generation {
                        cache.insert(
                            "body.xml".to_string(),
                            Arc::clone(&stale),
                            reader.cache_limits,
                        );
                    }
                    old.complete_success(stale);
                    finish_lazy_flight(&mut cache, "body.xml", &old);
                }
                let waited = waiter
                    .join()
                    .expect("old-generation waiter should not panic");
                (fresh, waited, active_after_clear, flights_empty, cold_loads)
            });
        assert_eq!(active_after_clear, 1);
        assert!(flights_empty);
        assert_eq!(fresh.unwrap().as_slice(), b"body");
        assert_eq!(cold_loads, initial_cold_loads + 1);
        assert_eq!(waited.unwrap().as_slice(), b"stale");
        assert_eq!(reader.read_shared("body.xml").unwrap().as_slice(), b"body");
        assert_eq!(reader.cache_size(), 1);
        assert_eq!(reader.active_flight_count(), 0);
        assert_eq!(reader.active_flight_key_bytes(), 0);
    }

    #[test]
    fn lazy_cache_take_fences_old_flights_and_wakes_waiters() {
        let bytes = fixture(&[FixtureEntry::stored(b"body.xml", b"body")]);
        let cache_limits =
            LazyArchiveCacheLimits::new_with_active_flight_limits(64, 2, 2, 64).unwrap();
        let reader =
            Arc::new(LazyArchiveReader::new_with_cache_limits(&bytes, cache_limits).unwrap());
        let old = register_lazy_test_flight(&reader, "body.xml");

        let (taken, fresh, waited, active_after_take, flights_empty, cold_loads) =
            std::thread::scope(|scope| {
                let waiter_reader = Arc::clone(&reader);
                let waiter = scope.spawn(move || waiter_reader.read_shared("body.xml"));
                while Arc::strong_count(&old) < 3 {
                    std::thread::yield_now();
                }

                let taken = reader.take_cache();
                let active_after_take = reader.active_flight_count();
                let flights_empty = lock_lazy_cache(&reader.cache).flights.is_empty();
                let fresh = reader.read_shared("body.xml");
                let cold_loads = reader.cold_loads.load(Ordering::SeqCst);

                let stale = Arc::new(b"stale".to_vec());
                {
                    let mut cache = lock_lazy_cache(&reader.cache);
                    if cache.generation == old.generation {
                        cache.insert(
                            "body.xml".to_string(),
                            Arc::clone(&stale),
                            reader.cache_limits,
                        );
                    }
                    old.complete_success(stale);
                    finish_lazy_flight(&mut cache, "body.xml", &old);
                }
                let waited = waiter
                    .join()
                    .expect("old-generation waiter should not panic");
                (
                    taken,
                    fresh,
                    waited,
                    active_after_take,
                    flights_empty,
                    cold_loads,
                )
            });
        assert!(taken.is_empty());
        assert_eq!(active_after_take, 1);
        assert!(flights_empty);
        assert_eq!(fresh.unwrap().as_slice(), b"body");
        assert_eq!(cold_loads, 1);
        assert_eq!(waited.unwrap().as_slice(), b"stale");
        assert_eq!(reader.read_shared("body.xml").unwrap().as_slice(), b"body");
        assert_eq!(reader.cache_size(), 1);
        assert_eq!(reader.active_flight_count(), 0);
        assert_eq!(reader.active_flight_key_bytes(), 0);
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
        let indexed = indexed_archive(bytes.clone());
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
        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(
            reader.file_names().collect::<Vec<_>>(),
            vec!["first.bin", "second.bin"]
        );

        let archive = ZipArchive::from_slice(&bytes).unwrap();
        let central = archive.directory_offset() as usize;
        let eocd = archive.eocd_offset() as usize;
        let first_len = central_record_len(&bytes, central);
        let second_len = central_record_len(&bytes, central + first_len);
        let first = bytes[central..central + first_len].to_vec();
        let second = bytes[central + first_len..central + first_len + second_len].to_vec();
        let third = bytes[central + first_len + second_len..eocd].to_vec();
        bytes[central..eocd].copy_from_slice(&[third, second, first].concat());

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(
            reader.file_names().collect::<Vec<_>>(),
            vec!["first.bin", "second.bin"]
        );
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

    #[test]
    fn equal_local_header_offsets_keep_existing_reader_ordering() {
        let mut bytes = fixture(&[
            FixtureEntry::stored(b"first.bin", b"first"),
            FixtureEntry::stored(b"second.bin", b"second"),
        ]);
        let archive = ZipArchive::from_slice(&bytes).unwrap();
        let central = archive.directory_offset() as usize;
        let first_len = central_record_len(&bytes, central);
        let second = central + first_len;
        bytes[second + 42..second + 46].copy_from_slice(&0u32.to_le_bytes());

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(
            reader.file_names().collect::<Vec<_>>(),
            vec!["first.bin", "second.bin"]
        );
        let indexed = indexed_archive(bytes);
        assert_eq!(
            indexed.file_names().collect::<Vec<_>>(),
            vec!["first.bin", "second.bin"]
        );
    }

    #[test]
    fn borrowed_store_validates_signed_and_unsigned_descriptors_and_preserves_identity() {
        let payload = b"descriptor payload";
        for signature in [false, true] {
            let bytes = stored_descriptor_fixture(payload, signature);
            let reader = ArchiveReader::new(&bytes).unwrap();
            let borrowed = reader.read_stored_borrowed("stored.bin").unwrap().unwrap();

            let archive = ZipArchive::from_slice(bytes.as_slice()).unwrap();
            let record = archive.entries().next().unwrap().unwrap();
            let physical = archive
                .get_entry_borrowed(record.wayfinder())
                .unwrap()
                .data();
            assert_eq!(borrowed, payload);
            assert_eq!(borrowed.as_ptr(), physical.as_ptr());
            assert_eq!(borrowed.len(), physical.len());
        }
    }

    #[test]
    fn borrowed_store_validates_descriptors_after_a_nonzero_prelude() {
        let payload = b"leading descriptor payload";
        for signature in [false, true] {
            let bytes = stored_descriptor_fixture_with_prefix(payload, signature, 17);
            let reader = ArchiveReader::new(&bytes).unwrap();
            let borrowed = reader.read_stored_borrowed("stored.bin").unwrap().unwrap();
            let archive = ZipArchive::from_slice(bytes.as_slice()).unwrap();
            let record = archive.entries().next().unwrap().unwrap();
            let physical = archive
                .get_entry_borrowed(record.wayfinder())
                .unwrap()
                .data();
            assert_eq!(borrowed, payload);
            assert_eq!(borrowed.as_ptr(), physical.as_ptr());
            assert_eq!(borrowed.len(), physical.len());
        }
    }

    #[test]
    fn strict_read_to_rejects_target_metadata_before_sink() {
        let payload = b"strict sink payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", payload).unwrap();
        let base = writer.finish_to_bytes().unwrap();
        let local = local_header_offset_for_name(&base, b"stored.bin");
        let central = central_header_offset_for_name(&base, b"stored.bin");

        for mutation in 0..9 {
            let mut bytes = base.clone();
            match mutation {
                0 => bytes[local + 30] ^= 1,
                1 => bytes[local + 8..local + 10].copy_from_slice(&8u16.to_le_bytes()),
                2 => bytes[local + 6..local + 8].copy_from_slice(&8u16.to_le_bytes()),
                3 => bytes[local + 14..local + 18].copy_from_slice(&u32::MAX.to_le_bytes()),
                4 => bytes[local + 18..local + 22].copy_from_slice(&1u32.to_le_bytes()),
                5 => bytes[local + 22..local + 26].copy_from_slice(&1u32.to_le_bytes()),
                6 => bytes[central + 16..central + 20].copy_from_slice(&u32::MAX.to_le_bytes()),
                7 => bytes[central + 20..central + 24].copy_from_slice(&1u32.to_le_bytes()),
                8 => bytes[central + 24..central + 28].copy_from_slice(&1u32.to_le_bytes()),
                _ => unreachable!(),
            }
            let reader = ArchiveReader::new(&bytes).unwrap();
            let mut sink = vec![0xA5];
            let _error = reader.read_to("stored.bin", &mut sink).unwrap_err();
            assert_eq!(sink, vec![0xA5], "mutation {mutation}");
        }
    }

    #[test]
    fn strict_read_to_rejects_target_encryption_but_ignores_unrelated_encryption() {
        for deflated in [false, true] {
            let payload = b"strict encrypted target";
            let mut writer = StreamingArchiveWriter::new();
            writer.write_stored("neighbor.bin", b"neighbor").unwrap();
            if deflated {
                writer.write_deflated("target.bin", payload).unwrap();
            } else {
                writer.write_stored("target.bin", payload).unwrap();
            }
            let base = writer.finish_to_bytes().unwrap();
            let neighbor_local = local_header_offset_for_name(&base, b"neighbor.bin");
            let neighbor_central = central_header_offset_for_name(&base, b"neighbor.bin");
            let mut unrelated = base.clone();
            unrelated[neighbor_local + 6..neighbor_local + 8].copy_from_slice(&1u16.to_le_bytes());
            unrelated[neighbor_central + 8..neighbor_central + 10]
                .copy_from_slice(&1u16.to_le_bytes());
            let reader = ArchiveReader::new(&unrelated).unwrap();
            let mut output = Vec::new();
            assert_eq!(
                reader.read_to("target.bin", &mut output).unwrap(),
                payload.len() as u64
            );
            assert_eq!(output, payload);

            let target_local = local_header_offset_for_name(&base, b"target.bin");
            let target_central = central_header_offset_for_name(&base, b"target.bin");
            for encrypted_flag in [1u16, 1 << 6] {
                let mut encrypted = base.clone();
                encrypted[target_local + 6..target_local + 8]
                    .copy_from_slice(&encrypted_flag.to_le_bytes());
                encrypted[target_central + 8..target_central + 10]
                    .copy_from_slice(&encrypted_flag.to_le_bytes());
                let reader = ArchiveReader::new(&encrypted).unwrap();
                let mut sink = vec![0xA5];
                let _error = reader.read_to("target.bin", &mut sink).unwrap_err();
                assert_eq!(sink, vec![0xA5]);
            }
        }
    }

    #[test]
    fn strict_read_to_accepts_prefix_and_rejects_span_intrusion_before_sink() {
        let payload = b"strict prefix payload";
        for signature in [false, true] {
            let bytes = stored_descriptor_fixture_with_prefix(payload, signature, 17);
            let reader = ArchiveReader::new(&bytes).unwrap();
            let mut output = Vec::new();
            assert_eq!(
                reader.read_to("stored.bin", &mut output).unwrap(),
                payload.len() as u64
            );
            assert_eq!(output, payload);
        }

        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("first.bin", b"first payload").unwrap();
        writer.write_stored("target.bin", b"target").unwrap();
        let base = writer.finish_to_bytes().unwrap();
        let first_local = local_header_offset_for_name(&base, b"first.bin");
        let first_central = central_header_offset_for_name(&base, b"first.bin");
        let second_local = local_header_offset_for_name(&base, b"target.bin");
        let name_len = usize::from(u16::from_le_bytes([
            base[first_local + 26],
            base[first_local + 27],
        ]));
        let extra_len = usize::from(u16::from_le_bytes([
            base[first_local + 28],
            base[first_local + 29],
        ]));
        let payload_offset = first_local + 30 + name_len + extra_len;
        let declared = u32::try_from(second_local - payload_offset + 1).unwrap();
        let mut overlap = base.clone();
        overlap[first_local + 18..first_local + 22].copy_from_slice(&declared.to_le_bytes());
        overlap[first_central + 20..first_central + 24].copy_from_slice(&declared.to_le_bytes());
        let reader = ArchiveReader::new(&overlap).unwrap();
        let mut sink = vec![0xA5];
        let _error = reader.read_to("target.bin", &mut sink).unwrap_err();
        assert_eq!(sink, vec![0xA5]);

        let target_local = local_header_offset_for_name(&base, b"target.bin");
        let target_central = central_header_offset_for_name(&base, b"target.bin");
        let target_name_len = usize::from(u16::from_le_bytes([
            base[target_local + 26],
            base[target_local + 27],
        ]));
        let target_extra_len = usize::from(u16::from_le_bytes([
            base[target_local + 28],
            base[target_local + 29],
        ]));
        let target_payload = target_local + 30 + target_name_len + target_extra_len;
        let central_start = central_header_offset_for_name(&base, b"first.bin");
        let declared = u32::try_from(central_start - target_payload + 1).unwrap();
        let mut intrusion = base;
        intrusion[target_local + 18..target_local + 22].copy_from_slice(&declared.to_le_bytes());
        intrusion[target_central + 20..target_central + 24]
            .copy_from_slice(&declared.to_le_bytes());
        intrusion[target_local + 22..target_local + 26].copy_from_slice(&declared.to_le_bytes());
        intrusion[target_central + 24..target_central + 28]
            .copy_from_slice(&declared.to_le_bytes());
        let reader = ArchiveReader::new(&intrusion).unwrap();
        let mut sink = vec![0xA5];
        let _error = reader.read_to("target.bin", &mut sink).unwrap_err();
        assert_eq!(sink, vec![0xA5]);
    }

    #[test]
    fn strict_read_to_rejects_deflate_trailing_bytes_after_sink_prefix() {
        let payload = b"strict deflate payload";
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_deflated_sized("deflated.bin", payload)
            .unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let local = local_header_offset_for_name(&bytes, b"deflated.bin");
        let central = central_header_offset_for_name(&bytes, b"deflated.bin");
        assert!(!local_member_has_data_descriptor(&bytes, b"deflated.bin"));
        let compressed = u32::from_le_bytes(bytes[local + 18..local + 22].try_into().unwrap());
        let junk = [0xA5, 0x5A, 0xC3];
        let eocd = bytes.len() - 22;
        bytes.splice(central..central, junk);
        let shifted_central = central + junk.len();
        let declared = compressed
            .checked_add(u32::try_from(junk.len()).unwrap())
            .unwrap();
        bytes[local + 18..local + 22].copy_from_slice(&declared.to_le_bytes());
        bytes[shifted_central + 20..shifted_central + 24].copy_from_slice(&declared.to_le_bytes());
        let shifted_eocd = eocd + junk.len();
        bytes[shifted_eocd + 16..shifted_eocd + 20]
            .copy_from_slice(&u32::try_from(shifted_central).unwrap().to_le_bytes());

        {
            let reader = ArchiveReader::new(&bytes).unwrap();
            let mut sink = b"prefix".to_vec();
            let error = reader.read_to("deflated.bin", &mut sink).unwrap_err();
            assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));
            assert!(sink.starts_with(b"prefix"));
            assert!(sink.len() >= b"prefix".len() + payload.len());
        }

        let indexed = indexed_archive(bytes);
        let entry_id = indexed.entry_id("deflated.bin").unwrap();
        let mut accounting = ZipOperationAccounting::default();
        let mut indexed_sink = b"prefix".to_vec();
        let error = indexed
            .read_entry_to_with_accounting(entry_id, &mut indexed_sink, &mut accounting)
            .unwrap_err();
        match error.kind() {
            ErrorKind::InvalidSize { expected, actual } => {
                assert_eq!(*expected, u64::from(declared));
                assert_eq!(*actual, u64::from(compressed));
            },
            other => panic!("expected indexed total_in size error, got {other:?}"),
        }
        assert!(indexed_sink.starts_with(b"prefix"));
        assert!(indexed_sink.len() >= b"prefix".len() + payload.len());
        assert_eq!(accounting.deflate_bytes_produced(), payload.len() as u64);
        assert_eq!(accounting.deflate_bytes_accepted(), payload.len() as u64);
    }

    #[test]
    fn indexed_read_entry_to_with_accounting_uses_strict_preflight() {
        let payload = b"indexed strict payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", payload).unwrap();
        writer.write_deflated("deflated.bin", payload).unwrap();
        let indexed = indexed_archive(writer.finish_to_bytes().unwrap());
        for name in ["stored.bin", "deflated.bin"] {
            let entry_id = indexed.entry_id(name).unwrap();
            let mut accounting = ZipOperationAccounting::default();
            let mut sink = Vec::new();
            assert_eq!(
                indexed
                    .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                    .unwrap(),
                payload.len() as u64
            );
            assert_eq!(sink, payload);
        }

        for signature in [false, true] {
            let indexed = indexed_archive(stored_descriptor_fixture(payload, signature));
            let entry_id = indexed.entry_id("stored.bin").unwrap();
            let mut accounting = ZipOperationAccounting::default();
            let mut sink = Vec::new();
            assert_eq!(
                indexed
                    .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                    .unwrap(),
                payload.len() as u64
            );
            assert_eq!(sink, payload);
        }
        for signature in [false, true] {
            for (field_offset, checksum) in [(0usize, true), (4, false), (8, false)] {
                let mut bytes = stored_descriptor_fixture(payload, signature);
                let descriptor = descriptor_start(payload, signature);
                let value = if checksum { 0 } else { u32::MAX };
                let field_start = descriptor + field_offset;
                bytes[field_start..field_start + 4].copy_from_slice(&value.to_le_bytes());
                let indexed = indexed_archive(bytes);
                let entry_id = indexed.entry_id("stored.bin").unwrap();
                let mut accounting = ZipOperationAccounting::default();
                let mut sink = vec![0xA5];
                let error = indexed
                    .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                    .unwrap_err();
                if checksum {
                    assert!(matches!(error.kind(), ErrorKind::InvalidChecksum { .. }));
                } else {
                    assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));
                }
                assert_eq!(sink, vec![0xA5]);
                assert_eq!(accounting, ZipOperationAccounting::default());
            }
        }
        let indexed = indexed_archive(stored_descriptor_fixture_with_prefix(payload, true, 13));
        let entry_id = indexed.entry_id("stored.bin").unwrap();
        let mut accounting = ZipOperationAccounting::default();
        let mut sink = Vec::new();
        assert_eq!(
            indexed
                .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                .unwrap(),
            payload.len() as u64
        );
        assert_eq!(sink, payload);

        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("neighbor.bin", b"neighbor").unwrap();
        writer.write_deflated("target.bin", payload).unwrap();
        let base = writer.finish_to_bytes().unwrap();
        let neighbor_local = local_header_offset_for_name(&base, b"neighbor.bin");
        let neighbor_central = central_header_offset_for_name(&base, b"neighbor.bin");
        let target_id = {
            let mut unrelated = base.clone();
            unrelated[neighbor_local + 6..neighbor_local + 8].copy_from_slice(&1u16.to_le_bytes());
            unrelated[neighbor_central + 8..neighbor_central + 10]
                .copy_from_slice(&1u16.to_le_bytes());
            indexed_archive(unrelated).entry_id("target.bin").unwrap()
        };
        let mut unrelated = base.clone();
        unrelated[neighbor_local + 6..neighbor_local + 8].copy_from_slice(&1u16.to_le_bytes());
        unrelated[neighbor_central + 8..neighbor_central + 10].copy_from_slice(&1u16.to_le_bytes());
        let indexed = indexed_archive(unrelated);
        let mut accounting = ZipOperationAccounting::default();
        let mut sink = Vec::new();
        assert_eq!(
            indexed
                .read_entry_to_with_accounting(target_id, &mut sink, &mut accounting)
                .unwrap(),
            payload.len() as u64
        );
        assert_eq!(sink, payload);

        let target_local = local_header_offset_for_name(&base, b"target.bin");
        let target_central = central_header_offset_for_name(&base, b"target.bin");
        for encrypted_flag in [1u16, 1 << 6] {
            let mut encrypted = base.clone();
            encrypted[target_local + 6..target_local + 8]
                .copy_from_slice(&encrypted_flag.to_le_bytes());
            encrypted[target_central + 8..target_central + 10]
                .copy_from_slice(&encrypted_flag.to_le_bytes());
            let indexed = indexed_archive(encrypted);
            let entry_id = indexed.entry_id("target.bin").unwrap();
            let mut accounting = ZipOperationAccounting::default();
            let mut sink = vec![0xA5];
            let error = indexed
                .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                .unwrap_err();
            assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
            assert_eq!(sink, vec![0xA5]);
            assert_eq!(accounting, ZipOperationAccounting::default());
        }

        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("zero-crc.bin", payload).unwrap();
        let mut zero_crc = writer.finish_to_bytes().unwrap();
        let zero_local = local_header_offset_for_name(&zero_crc, b"zero-crc.bin");
        let zero_central = central_header_offset_for_name(&zero_crc, b"zero-crc.bin");
        let zero_name_len = usize::from(u16::from_le_bytes([
            zero_crc[zero_local + 26],
            zero_crc[zero_local + 27],
        ]));
        let zero_extra_len = usize::from(u16::from_le_bytes([
            zero_crc[zero_local + 28],
            zero_crc[zero_local + 29],
        ]));
        let zero_payload = zero_local + 30 + zero_name_len + zero_extra_len;
        zero_crc[zero_local + 14..zero_local + 18].copy_from_slice(&0u32.to_le_bytes());
        zero_crc[zero_central + 16..zero_central + 20].copy_from_slice(&0u32.to_le_bytes());
        zero_crc[zero_payload] ^= 0x80;
        let indexed = indexed_archive(zero_crc);
        let entry_id = indexed.entry_id("zero-crc.bin").unwrap();
        let mut accounting = ZipOperationAccounting::default();
        let mut sink = Vec::new();
        let error = indexed
            .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidChecksum { .. }));
        assert_eq!(sink.len(), payload.len());
        assert!(accounting.stored_payload_bytes_read() > 0);
        assert_eq!(indexed.read("zero-crc.bin").unwrap().len(), payload.len());

        let indexed = indexed_archive(include_bytes!("../assets/zip64.zip").to_vec());
        let entry_id = indexed.entry_id("README").unwrap();
        let mut accounting = ZipOperationAccounting::default();
        let mut sink = Vec::new();
        assert!(
            indexed
                .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                .is_ok()
        );
        assert_eq!(sink, b"This small file is in ZIP64 format.\n");

        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("unresolved.bin", payload).unwrap();
        let mut unresolved = writer.finish_to_bytes().unwrap();
        let unresolved_central = central_header_offset_for_name(&unresolved, b"unresolved.bin");
        unresolved[unresolved_central + 42..unresolved_central + 46]
            .copy_from_slice(&u32::MAX.to_le_bytes());
        let error = indexed_archive_result(unresolved, ArchiveLimits::UNBOUNDED).unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("first.bin", b"first payload").unwrap();
        writer.write_stored("target.bin", b"target").unwrap();
        let base = writer.finish_to_bytes().unwrap();
        let first_local = local_header_offset_for_name(&base, b"first.bin");
        let first_central = central_header_offset_for_name(&base, b"first.bin");
        let target_local = local_header_offset_for_name(&base, b"target.bin");
        let first_name_len = usize::from(u16::from_le_bytes([
            base[first_local + 26],
            base[first_local + 27],
        ]));
        let first_extra_len = usize::from(u16::from_le_bytes([
            base[first_local + 28],
            base[first_local + 29],
        ]));
        let first_payload = first_local + 30 + first_name_len + first_extra_len;
        let declared = u32::try_from(target_local - first_payload + 1).unwrap();
        let mut overlap = base.clone();
        overlap[first_local + 18..first_local + 22].copy_from_slice(&declared.to_le_bytes());
        overlap[first_central + 20..first_central + 24].copy_from_slice(&declared.to_le_bytes());
        let indexed = indexed_archive(overlap);
        let entry_id = indexed.entry_id("target.bin").unwrap();
        let mut accounting = ZipOperationAccounting::default();
        let mut sink = vec![0xA5];
        let error = indexed
            .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
        assert_eq!(sink, vec![0xA5]);
        assert_eq!(accounting, ZipOperationAccounting::default());

        let target_central = central_header_offset_for_name(&base, b"target.bin");
        let target_name_len = usize::from(u16::from_le_bytes([
            base[target_local + 26],
            base[target_local + 27],
        ]));
        let target_extra_len = usize::from(u16::from_le_bytes([
            base[target_local + 28],
            base[target_local + 29],
        ]));
        let target_payload = target_local + 30 + target_name_len + target_extra_len;
        let declared = u32::try_from(target_central - target_payload + 1).unwrap();
        let mut intrusion = base.clone();
        intrusion[target_local + 18..target_local + 22].copy_from_slice(&declared.to_le_bytes());
        intrusion[target_local + 22..target_local + 26].copy_from_slice(&declared.to_le_bytes());
        intrusion[target_central + 20..target_central + 24]
            .copy_from_slice(&declared.to_le_bytes());
        intrusion[target_central + 24..target_central + 28]
            .copy_from_slice(&declared.to_le_bytes());
        let indexed = indexed_archive(intrusion);
        let entry_id = indexed.entry_id("target.bin").unwrap();
        let mut accounting = ZipOperationAccounting::default();
        let mut sink = vec![0xA5];
        let error = indexed
            .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::Eof));
        assert_eq!(sink, vec![0xA5]);
        assert_eq!(accounting, ZipOperationAccounting::default());
    }

    #[test]
    fn indexed_strict_preflight_rejects_local_central_mismatches() {
        let payload = b"indexed metadata mismatch payload";
        for deflated in [false, true] {
            let mut writer = StreamingArchiveWriter::new();
            if deflated {
                writer.write_deflated_sized("member.bin", payload).unwrap();
            } else {
                writer.write_stored("member.bin", payload).unwrap();
            }
            let base = writer.finish_to_bytes().unwrap();
            assert!(!local_member_has_data_descriptor(&base, b"member.bin"));
            let local = local_header_offset_for_name(&base, b"member.bin");
            let central = central_header_offset_for_name(&base, b"member.bin");
            let crc = crate::crc32(payload);
            let local_compressed =
                u32::from_le_bytes(base[local + 18..local + 22].try_into().unwrap());
            let central_compressed =
                u32::from_le_bytes(base[central + 20..central + 24].try_into().unwrap());
            let local_uncompressed =
                u32::from_le_bytes(base[local + 22..local + 26].try_into().unwrap());
            let central_uncompressed =
                u32::from_le_bytes(base[central + 24..central + 28].try_into().unwrap());

            for mutation in 0..9 {
                let mut bytes = base.clone();
                match mutation {
                    0 => bytes[local + 30] ^= 1,
                    1 => bytes[local + 8..local + 10]
                        .copy_from_slice(&(if deflated { 0u16 } else { 8u16 }).to_le_bytes()),
                    2 => bytes[local + 6..local + 8].copy_from_slice(&8u16.to_le_bytes()),
                    3 => bytes[local + 14..local + 18]
                        .copy_from_slice(&crc.wrapping_add(1).to_le_bytes()),
                    4 => bytes[central + 16..central + 20]
                        .copy_from_slice(&crc.wrapping_add(1).to_le_bytes()),
                    5 => bytes[local + 18..local + 22]
                        .copy_from_slice(&local_compressed.wrapping_add(1).to_le_bytes()),
                    6 => bytes[central + 20..central + 24]
                        .copy_from_slice(&central_compressed.wrapping_add(1).to_le_bytes()),
                    7 => bytes[local + 22..local + 26]
                        .copy_from_slice(&local_uncompressed.wrapping_add(1).to_le_bytes()),
                    8 => bytes[central + 24..central + 28]
                        .copy_from_slice(&central_uncompressed.wrapping_add(1).to_le_bytes()),
                    _ => unreachable!(),
                }

                let indexed = indexed_archive(bytes);
                let entry_id = indexed.entry_id("member.bin").unwrap();
                let mut sink = vec![0xA5];
                let mut accounting = ZipOperationAccounting::default();
                let error = indexed
                    .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                    .unwrap_err();
                if mutation <= 2 {
                    assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
                } else if mutation <= 4 {
                    assert!(matches!(error.kind(), ErrorKind::InvalidChecksum { .. }));
                } else {
                    assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));
                }
                assert_eq!(sink, vec![0xA5]);
                assert_eq!(accounting, ZipOperationAccounting::default());
            }
        }
    }

    #[test]
    fn indexed_strict_stream_validates_tiny_zip64_descriptors_and_resolution() {
        let payload = b"z";
        let size = u64::try_from(payload.len()).unwrap();
        for signature in [false, true] {
            let bytes = zip64_descriptor_fixture(payload, signature);
            let indexed = indexed_archive_result(bytes.clone(), ArchiveLimits::UNBOUNDED).unwrap();
            let entry_id = indexed.entry_id("zip64.bin").unwrap();
            let mut sink = Vec::new();
            let mut accounting = ZipOperationAccounting::default();
            assert_eq!(
                indexed
                    .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                    .unwrap(),
                size
            );
            assert_eq!(sink, payload);

            let descriptor = zip64_descriptor_start(&bytes, payload.len());
            let crc_offset = if signature { 4 } else { 0 };
            let compressed_offset = if signature { 8 } else { 4 };
            let uncompressed_offset = if signature { 16 } else { 12 };
            for (offset, checksum) in [
                (crc_offset, true),
                (compressed_offset, false),
                (uncompressed_offset, false),
            ] {
                let mut corrupt = bytes.clone();
                let field_start = descriptor + offset;
                if checksum {
                    corrupt[field_start..field_start + 4].copy_from_slice(&0u32.to_le_bytes());
                } else {
                    corrupt[field_start..field_start + 8].copy_from_slice(&u64::MAX.to_le_bytes());
                }
                let indexed = indexed_archive_result(corrupt, ArchiveLimits::UNBOUNDED).unwrap();
                let entry_id = indexed.entry_id("zip64.bin").unwrap();
                let mut sink = vec![0xA5];
                let mut accounting = ZipOperationAccounting::default();
                let error = indexed
                    .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                    .unwrap_err();
                if checksum {
                    assert!(matches!(error.kind(), ErrorKind::InvalidChecksum { .. }));
                } else {
                    assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));
                }
                assert_eq!(sink, vec![0xA5]);
                assert_eq!(accounting, ZipOperationAccounting::default());
            }
        }

        let central_missing_extra =
            zip64_descriptor_fixture_with_central_sizes(payload, false, &[], u32::MAX, u32::MAX);
        let unresolved_compressed = zip64_descriptor_fixture_with_central_sizes(
            payload,
            false,
            &[],
            u32::MAX,
            u32::try_from(payload.len()).unwrap(),
        );
        let unresolved_uncompressed = zip64_descriptor_fixture_with_central_sizes(
            payload,
            false,
            &[],
            u32::try_from(payload.len()).unwrap(),
            u32::MAX,
        );
        for bytes in [
            central_missing_extra,
            unresolved_compressed,
            unresolved_uncompressed,
        ] {
            let error = indexed_strict_error(bytes);
            assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
        }

        let mut truncated_local_extra = zip64_descriptor_fixture(payload, false);
        let local = local_header_offset_for_name(&truncated_local_extra, b"zip64.bin");
        truncated_local_extra[local + 28..local + 30].copy_from_slice(&8u16.to_le_bytes());
        let error = indexed_strict_error(truncated_local_extra);
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let mut unresolved_offset = zip64_descriptor_fixture(payload, false);
        let central = central_header_offset_for_name(&unresolved_offset, b"zip64.bin");
        unresolved_offset[central + 42..central + 46].copy_from_slice(&u32::MAX.to_le_bytes());
        let error = indexed_strict_error(unresolved_offset);
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let mut unresolved_disk = zip64_descriptor_fixture(payload, false);
        let central = central_header_offset_for_name(&unresolved_disk, b"zip64.bin");
        unresolved_disk[central + 34..central + 36].copy_from_slice(&u16::MAX.to_le_bytes());
        let error = indexed_strict_error(unresolved_disk);
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    }

    #[test]
    fn indexed_archive_bounds_directory_metadata_before_layout_publication() {
        let empty = indexed_archive_result(
            fixture(&[]),
            ArchiveLimits {
                max_metadata_bytes: 0,
                ..ArchiveLimits::UNBOUNDED
            },
        )
        .unwrap();
        assert_eq!(empty.len(), 0);

        let one = fixture(&[FixtureEntry {
            name: b"folder/",
            extra: b"xy",
            comment: b"z",
            compressed_size: 0,
            uncompressed_size: 0,
            data: b"",
        }]);
        let one_metadata = CENTRAL_FIXED_RECORD_BYTES + 7 + 2 + 1;
        let one_limits = ArchiveLimits {
            max_files: 0,
            max_member_name_bytes: 7,
            max_metadata_bytes: one_metadata,
            max_compressed_size: 0,
            max_entry_size: 0,
            max_total_size: 0,
        };
        let indexed = indexed_archive_result(one.clone(), one_limits).unwrap();
        assert_eq!(indexed.len(), 0);
        assert!(indexed.metadata("folder/").unwrap().is_directory());

        let mut one_over = one_limits;
        one_over.max_metadata_bytes -= 1;
        assert_limit(
            indexed_archive_result(one, one_over).unwrap_err(),
            LimitResource::MetadataBytes,
            one_metadata,
            one_metadata - 1,
        );

        let multiple = fixture(&[
            FixtureEntry {
                name: b"a/",
                extra: b"123",
                comment: b"q",
                compressed_size: 0,
                uncompressed_size: 0,
                data: b"",
            },
            FixtureEntry {
                name: b"bb/",
                extra: b"xy",
                comment: b"z",
                compressed_size: 0,
                uncompressed_size: 0,
                data: b"",
            },
        ]);
        let multiple_metadata = 2 * (CENTRAL_FIXED_RECORD_BYTES + 6);
        let multiple_limits = ArchiveLimits {
            max_files: 0,
            max_member_name_bytes: 3,
            max_metadata_bytes: multiple_metadata,
            max_compressed_size: 0,
            max_entry_size: 0,
            max_total_size: 0,
        };
        let indexed = indexed_archive_result(multiple.clone(), multiple_limits).unwrap();
        assert_eq!(indexed.len(), 0);
        assert!(indexed.metadata("a/").unwrap().is_directory());
        assert!(indexed.metadata("bb/").unwrap().is_directory());

        let mut multiple_over = multiple_limits;
        multiple_over.max_metadata_bytes -= 1;
        assert_limit(
            indexed_archive_result(multiple, multiple_over).unwrap_err(),
            LimitResource::MetadataBytes,
            multiple_metadata,
            multiple_metadata - 1,
        );
    }

    #[test]
    fn indexed_strict_zero_crc_handles_store_deflate_and_empty_members() {
        let payload = b"indexed nonempty zero CRC payload";
        let actual_crc = crate::crc32(payload);
        let indexed = indexed_archive(skipped_crc_fixture(payload));

        for name in ["stored.bin", "deflated.bin"] {
            let entry_id = indexed.entry_id(name).unwrap();
            let mut sink = Vec::new();
            let mut accounting = ZipOperationAccounting::default();
            let error = indexed
                .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                .unwrap_err();
            assert_zero_crc_checksum(error, actual_crc);
            assert_eq!(sink, payload);
            assert_eq!(indexed.read(name).unwrap(), payload);
        }

        for name in ["empty-stored.bin", "empty-deflated.bin"] {
            let entry_id = indexed.entry_id(name).unwrap();
            let mut sink = Vec::new();
            let mut accounting = ZipOperationAccounting::default();
            assert_eq!(
                indexed
                    .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                    .unwrap(),
                0
            );
            assert!(sink.is_empty());
            assert!(indexed.read(name).unwrap().is_empty());
        }
    }

    #[test]
    fn indexed_strict_layout_accepts_prefix_and_gaps_and_rejects_store_deflate_spans() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", b"stored").unwrap();
        writer
            .write_deflated_sized("deflated.bin", b"deflated")
            .unwrap();
        let base = writer.finish_to_bytes().unwrap();
        let central = central_header_offset_for_name(&base, b"stored.bin");
        let eocd = base.len() - 22;

        let prefix_len = 7usize;
        let mut prefixed = vec![0xA5; prefix_len];
        prefixed.extend_from_slice(&base);
        for name in [b"stored.bin".as_slice(), b"deflated.bin".as_slice()] {
            let central_offset = central_header_offset_for_name(&prefixed, name);
            let old_local = u32::from_le_bytes(
                prefixed[central_offset + 42..central_offset + 46]
                    .try_into()
                    .unwrap(),
            );
            let shifted_local = old_local
                .checked_add(u32::try_from(prefix_len).unwrap())
                .unwrap();
            prefixed[central_offset + 42..central_offset + 46]
                .copy_from_slice(&shifted_local.to_le_bytes());
        }
        let prefixed_eocd = eocd + prefix_len;
        prefixed[prefixed_eocd + 16..prefixed_eocd + 20]
            .copy_from_slice(&u32::try_from(central + prefix_len).unwrap().to_le_bytes());
        let indexed = indexed_archive(prefixed);
        for (name, expected) in [
            ("stored.bin", b"stored".as_slice()),
            ("deflated.bin", b"deflated".as_slice()),
        ] {
            let entry_id = indexed.entry_id(name).unwrap();
            let mut sink = Vec::new();
            assert_eq!(
                indexed.read_entry_to(entry_id, &mut sink).unwrap(),
                expected.len() as u64
            );
            assert_eq!(sink, expected);
        }

        let second_local = local_header_offset_for_name(&base, b"deflated.bin");
        let gap_len = 11usize;
        let mut gapped = base.clone();
        gapped.splice(second_local..second_local, vec![0x5A; gap_len]);
        let shifted_central = central + gap_len;
        let shifted_second_central = central_header_offset_for_name(&gapped, b"deflated.bin");
        let shifted_second_local = second_local + gap_len;
        gapped[shifted_second_central + 42..shifted_second_central + 46]
            .copy_from_slice(&u32::try_from(shifted_second_local).unwrap().to_le_bytes());
        let shifted_eocd = eocd + gap_len;
        gapped[shifted_eocd + 16..shifted_eocd + 20]
            .copy_from_slice(&u32::try_from(shifted_central).unwrap().to_le_bytes());
        let indexed = indexed_archive(gapped);
        for (name, expected) in [
            ("stored.bin", b"stored".as_slice()),
            ("deflated.bin", b"deflated".as_slice()),
        ] {
            let entry_id = indexed.entry_id(name).unwrap();
            let mut sink = Vec::new();
            assert_eq!(
                indexed.read_entry_to(entry_id, &mut sink).unwrap(),
                expected.len() as u64
            );
            assert_eq!(sink, expected);
        }

        for first_deflated in [false, true] {
            for target_deflated in [false, true] {
                let mut writer = StreamingArchiveWriter::new();
                if first_deflated {
                    writer
                        .write_deflated_sized("first.bin", b"first payload")
                        .unwrap();
                } else {
                    writer.write_stored("first.bin", b"first payload").unwrap();
                }
                if target_deflated {
                    writer
                        .write_deflated_sized("target.bin", b"target payload")
                        .unwrap();
                } else {
                    writer
                        .write_stored("target.bin", b"target payload")
                        .unwrap();
                }
                let base = writer.finish_to_bytes().unwrap();
                let first_local = local_header_offset_for_name(&base, b"first.bin");
                let first_central = central_header_offset_for_name(&base, b"first.bin");
                let target_local = local_header_offset_for_name(&base, b"target.bin");
                let name_len = usize::from(u16::from_le_bytes(
                    base[first_local + 26..first_local + 28].try_into().unwrap(),
                ));
                let extra_len = usize::from(u16::from_le_bytes(
                    base[first_local + 28..first_local + 30].try_into().unwrap(),
                ));
                let first_payload = first_local + 30 + name_len + extra_len;
                let declared = u32::try_from(target_local - first_payload + 1).unwrap();
                let mut overlap = base.clone();
                overlap[first_local + 18..first_local + 22]
                    .copy_from_slice(&declared.to_le_bytes());
                overlap[first_central + 20..first_central + 24]
                    .copy_from_slice(&declared.to_le_bytes());
                let indexed = indexed_archive(overlap);
                let entry_id = indexed.entry_id("target.bin").unwrap();
                let mut sink = vec![0xA5];
                let mut accounting = ZipOperationAccounting::default();
                let error = indexed
                    .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                    .unwrap_err();
                assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
                assert_eq!(sink, vec![0xA5]);
                assert_eq!(accounting, ZipOperationAccounting::default());
            }
        }

        for target_deflated in [false, true] {
            let mut writer = StreamingArchiveWriter::new();
            writer.write_stored("first.bin", b"first payload").unwrap();
            if target_deflated {
                writer
                    .write_deflated_sized("target.bin", b"target payload")
                    .unwrap();
            } else {
                writer
                    .write_stored("target.bin", b"target payload")
                    .unwrap();
            }
            let base = writer.finish_to_bytes().unwrap();
            let target_local = local_header_offset_for_name(&base, b"target.bin");
            let target_central = central_header_offset_for_name(&base, b"target.bin");
            let target_name_len = usize::from(u16::from_le_bytes([
                base[target_local + 26],
                base[target_local + 27],
            ]));
            let target_extra_len = usize::from(u16::from_le_bytes([
                base[target_local + 28],
                base[target_local + 29],
            ]));
            let target_payload = target_local + 30 + target_name_len + target_extra_len;
            let central_start = central_header_offset_for_name(&base, b"first.bin");
            let declared = u32::try_from(central_start - target_payload + 1).unwrap();
            let mut intrusion = base.clone();
            intrusion[target_local + 18..target_local + 22]
                .copy_from_slice(&declared.to_le_bytes());
            intrusion[target_central + 20..target_central + 24]
                .copy_from_slice(&declared.to_le_bytes());
            if !target_deflated {
                intrusion[target_local + 22..target_local + 26]
                    .copy_from_slice(&declared.to_le_bytes());
                intrusion[target_central + 24..target_central + 28]
                    .copy_from_slice(&declared.to_le_bytes());
            }
            let indexed = indexed_archive(intrusion);
            let entry_id = indexed.entry_id("target.bin").unwrap();
            let mut sink = vec![0xA5];
            let mut accounting = ZipOperationAccounting::default();
            let error = indexed
                .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                .unwrap_err();
            assert!(matches!(error.kind(), ErrorKind::Eof));
            assert_eq!(sink, vec![0xA5]);
            assert_eq!(accounting, ZipOperationAccounting::default());
        }
    }

    #[test]
    fn borrowed_store_rejects_corrupt_descriptor_crc_and_sizes() {
        let payload = b"descriptor payload";
        for signature in [false, true] {
            for (field_offset, checksum) in [(0usize, true), (4, false), (8, false)] {
                let mut bytes = stored_descriptor_fixture(payload, signature);
                let descriptor = descriptor_start(payload, signature);
                let value = if checksum { 0 } else { u32::MAX };
                let field_start = descriptor + field_offset;
                bytes[field_start..field_start + 4].copy_from_slice(&value.to_le_bytes());

                let reader = ArchiveReader::new(&bytes).unwrap();
                let error = reader.read_stored_borrowed("stored.bin").unwrap_err();
                if checksum {
                    assert!(
                        matches!(error.kind(), ErrorKind::InvalidChecksum { .. }),
                        "signature={signature}, field_offset={field_offset}, error_kind={:?}",
                        error.kind()
                    );
                } else {
                    assert!(
                        matches!(error.kind(), ErrorKind::InvalidSize { .. }),
                        "signature={signature}, field_offset={field_offset}, error_kind={:?}",
                        error.kind()
                    );
                }
            }
        }
    }

    #[test]
    fn borrowed_store_rejects_local_central_metadata_mismatches() {
        let payload = b"fixed stored payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", payload).unwrap();
        let base = writer.finish_to_bytes().unwrap();
        assert!(!local_member_has_data_descriptor(&base, b"stored.bin"));
        let local = local_header_offset_for_name(&base, b"stored.bin");

        let mut method = base.clone();
        method[local + 8..local + 10].copy_from_slice(&8u16.to_le_bytes());
        let error = ArchiveReader::new(&method)
            .unwrap()
            .read_stored_borrowed("stored.bin")
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let mut flags = base.clone();
        flags[local + 6..local + 8].copy_from_slice(&8u16.to_le_bytes());
        let error = ArchiveReader::new(&flags)
            .unwrap()
            .read_stored_borrowed("stored.bin")
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let central = central_header_offset_for_name(&base, b"stored.bin");
        for encrypted_flag in [1u16, 1 << 6] {
            let mut encrypted = base.clone();
            encrypted[local + 6..local + 8].copy_from_slice(&encrypted_flag.to_le_bytes());
            encrypted[central + 8..central + 10].copy_from_slice(&encrypted_flag.to_le_bytes());
            let error = ArchiveReader::new(&encrypted)
                .unwrap()
                .read_stored_borrowed("stored.bin")
                .unwrap_err();
            assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
        }

        let mut name = base.clone();
        name[local + 30] ^= 0x20;
        let error = ArchiveReader::new(&name)
            .unwrap()
            .read_stored_borrowed("stored.bin")
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let mut crc = base.clone();
        crc[local + 14..local + 18].copy_from_slice(&0u32.to_le_bytes());
        let error = ArchiveReader::new(&crc)
            .unwrap()
            .read_stored_borrowed("stored.bin")
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidChecksum { .. }));

        let mut compressed_size = base.clone();
        compressed_size[local + 18..local + 22].copy_from_slice(&1u32.to_le_bytes());
        let error = ArchiveReader::new(&compressed_size)
            .unwrap()
            .read_stored_borrowed("stored.bin")
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));

        let mut uncompressed_size = base.clone();
        uncompressed_size[local + 22..local + 26].copy_from_slice(&1u32.to_le_bytes());
        let error = ArchiveReader::new(&uncompressed_size)
            .unwrap()
            .read_stored_borrowed("stored.bin")
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));
    }

    #[test]
    fn unrelated_encrypted_store_does_not_block_an_unencrypted_target() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("target.bin", b"target").unwrap();
        writer.write_stored("encrypted.bin", b"encrypted").unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let encrypted_local = local_header_offset_for_name(&bytes, b"encrypted.bin");
        let encrypted_central = central_header_offset_for_name(&bytes, b"encrypted.bin");
        bytes[encrypted_local + 6..encrypted_local + 8].copy_from_slice(&1u16.to_le_bytes());
        bytes[encrypted_central + 8..encrypted_central + 10].copy_from_slice(&1u16.to_le_bytes());

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(
            reader.read_stored_borrowed("target.bin").unwrap(),
            Some(&b"target"[..])
        );
    }

    #[test]
    fn deflate_borrowing_returns_none_and_owned_read_remains_available() {
        let payload = b"deflated fallback payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("deflated.bin", payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let reader = ArchiveReader::new(&bytes).unwrap();

        assert_eq!(reader.read_stored_borrowed("deflated.bin").unwrap(), None);
        assert_eq!(reader.read("deflated.bin").unwrap(), payload);
    }

    #[test]
    fn encrypted_deflate_borrowing_is_rejected_before_ineligible_fallback() {
        let payload = b"encrypted deflated payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("deflated.bin", payload).unwrap();
        let base = writer.finish_to_bytes().unwrap();
        let local = local_header_offset_for_name(&base, b"deflated.bin");
        let central = central_header_offset_for_name(&base, b"deflated.bin");

        for encrypted_flag in [1u16, 1 << 6] {
            let mut encrypted = base.clone();
            encrypted[local + 6..local + 8].copy_from_slice(&encrypted_flag.to_le_bytes());
            encrypted[central + 8..central + 10].copy_from_slice(&encrypted_flag.to_le_bytes());
            let reader = ArchiveReader::new(&encrypted).unwrap();
            let error = reader.read_stored_borrowed("deflated.bin").unwrap_err();
            assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

            let mut accounting = ZipOperationAccounting::default();
            let error = reader
                .read_stored_borrowed_with_accounting("deflated.bin", &mut accounting)
                .unwrap_err();
            assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
        }
    }

    #[test]
    fn borrowed_store_refuses_an_overlapping_non_target_span() {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("underclaimed.bin", b"underclaim payload")
            .unwrap();
        writer.write_stored("target.bin", b"target").unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let first_local = local_header_offset_for_name(&bytes, b"underclaimed.bin");
        let first_central = central_header_offset_for_name(&bytes, b"underclaimed.bin");
        let second_local = local_header_offset_for_name(&bytes, b"target.bin");
        let name_len = usize::from(u16::from_le_bytes([
            bytes[first_local + 26],
            bytes[first_local + 27],
        ]));
        let extra_len = usize::from(u16::from_le_bytes([
            bytes[first_local + 28],
            bytes[first_local + 29],
        ]));
        let payload_offset = first_local + 30 + name_len + extra_len;
        let declared_size = u32::try_from(second_local - payload_offset + 1).unwrap();
        bytes[first_local + 18..first_local + 22].copy_from_slice(&declared_size.to_le_bytes());
        bytes[first_central + 20..first_central + 24].copy_from_slice(&declared_size.to_le_bytes());

        let reader = ArchiveReader::new(&bytes).unwrap();
        let error = reader.read_stored_borrowed("target.bin").unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    }

    #[test]
    fn borrowed_store_returns_none_for_nonempty_zero_crc_even_when_payload_is_corrupt() {
        let payload = b"zero CRC payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", payload).unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let local = local_header_offset_for_name(&bytes, b"stored.bin");
        let central = central_header_offset_for_name(&bytes, b"stored.bin");
        let name_len = usize::from(u16::from_le_bytes([bytes[local + 26], bytes[local + 27]]));
        let extra_len = usize::from(u16::from_le_bytes([bytes[local + 28], bytes[local + 29]]));
        let payload_offset = local + 30 + name_len + extra_len;
        bytes[local + 14..local + 18].copy_from_slice(&0u32.to_le_bytes());
        bytes[central + 16..central + 20].copy_from_slice(&0u32.to_le_bytes());
        bytes[payload_offset] ^= 0x80;

        let reader = ArchiveReader::new(&bytes).unwrap();
        assert_eq!(reader.read_stored_borrowed("stored.bin").unwrap(), None);
        let mut accounting = ZipOperationAccounting::default();
        assert_eq!(
            reader
                .read_stored_borrowed_with_accounting("stored.bin", &mut accounting)
                .unwrap(),
            None
        );
    }

    #[test]
    fn strict_verified_paths_reject_nonempty_zero_crc_and_accept_empty() {
        let payload = b"nonempty zero CRC payload";
        assert_ne!(crate::crc32(payload), 0);
        let bytes = skipped_crc_fixture(payload);
        let actual_crc = crate::crc32(payload);

        let reader = ArchiveReader::new(&bytes).unwrap();
        let mut stored_output = Vec::new();
        let error = reader
            .read_to("stored.bin", &mut stored_output)
            .unwrap_err();
        assert_zero_crc_checksum(error, actual_crc);
        assert_eq!(stored_output, payload);

        let mut deflated_output = Vec::new();
        let error = reader
            .read_to("deflated.bin", &mut deflated_output)
            .unwrap_err();
        assert_zero_crc_checksum(error, actual_crc);
        assert_eq!(deflated_output, payload);

        let mut accounted_output = Vec::new();
        let mut accounting = ZipOperationAccounting::default();
        let error = reader
            .read_to_with_accounting("stored.bin", &mut accounted_output, &mut accounting)
            .unwrap_err();
        assert_zero_crc_checksum(error, actual_crc);
        assert_eq!(accounted_output, payload);

        assert_eq!(reader.read("stored.bin").unwrap(), payload);
        let mut owned_accounting = ZipOperationAccounting::default();
        assert_eq!(
            reader
                .read_with_accounting("deflated.bin", &mut owned_accounting)
                .unwrap(),
            payload
        );

        let indexed = indexed_archive(bytes.clone());
        assert_eq!(indexed.read("deflated.bin").unwrap(), payload);
        let indexed_id = indexed.entry_id("deflated.bin").unwrap();
        let mut indexed_output = Vec::new();
        let error = indexed
            .read_entry_to(indexed_id, &mut indexed_output)
            .unwrap_err();
        assert_zero_crc_checksum(error, actual_crc);
        assert_eq!(indexed_output, payload);

        let lazy = LazyArchiveReader::new(&bytes).unwrap();
        let mut lazy_output = Vec::new();
        let mut lazy_accounting = ZipOperationAccounting::default();
        let error = lazy
            .read_to_with_accounting("deflated.bin", &mut lazy_output, &mut lazy_accounting)
            .unwrap_err();
        assert_zero_crc_checksum(error, actual_crc);
        assert_eq!(lazy_output, payload);
        assert_eq!(lazy.cache_size(), 0);

        for name in ["empty-stored.bin", "empty-deflated.bin"] {
            let mut output = Vec::new();
            assert_eq!(reader.read_to(name, &mut output).unwrap(), 0);
            assert!(output.is_empty());
        }
    }

    #[test]
    fn borrowed_store_rejects_payload_corruption_with_a_nonzero_crc() {
        let payload = b"nonzero CRC payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", payload).unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let local = local_header_offset_for_name(&bytes, b"stored.bin");
        let name_len = usize::from(u16::from_le_bytes([bytes[local + 26], bytes[local + 27]]));
        let extra_len = usize::from(u16::from_le_bytes([bytes[local + 28], bytes[local + 29]]));
        bytes[local + 30 + name_len + extra_len] ^= 0x80;

        let reader = ArchiveReader::new(&bytes).unwrap();
        let error = reader.read_stored_borrowed("stored.bin").unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidChecksum { .. }));
    }

    #[test]
    fn borrowed_store_refuses_overlapping_local_spans() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("first.bin", b"first").unwrap();
        writer.write_stored("second.bin", b"second").unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let first_local = local_header_offset_for_name(&bytes, b"first.bin");
        let second_central = central_header_offset_for_name(&bytes, b"second.bin");
        bytes[second_central + 42..second_central + 46]
            .copy_from_slice(&(u32::try_from(first_local).unwrap()).to_le_bytes());

        let reader = ArchiveReader::new(&bytes).unwrap();
        let error = reader.read_stored_borrowed("first.bin").unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    }

    #[test]
    fn borrowed_store_accepts_valid_local_zip64_size_extra() {
        let name = b"stored.bin";
        let payload = b"zip64 local payload";
        let size = u64::try_from(payload.len()).unwrap();
        let mut extra = Vec::new();
        push_u16(&mut extra, 1);
        push_u16(&mut extra, 16);
        extra.extend_from_slice(&size.to_le_bytes());
        extra.extend_from_slice(&size.to_le_bytes());

        let mut archive = Vec::new();
        push_u32(&mut archive, 0x0403_4b50);
        push_u16(&mut archive, 45);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u32(&mut archive, crate::crc32(payload));
        push_u32(&mut archive, u32::MAX);
        push_u32(&mut archive, u32::MAX);
        push_u16(&mut archive, u16::try_from(name.len()).unwrap());
        push_u16(&mut archive, u16::try_from(extra.len()).unwrap());
        archive.extend_from_slice(name);
        archive.extend_from_slice(&extra);
        archive.extend_from_slice(payload);

        let central_directory_offset = u32::try_from(archive.len()).unwrap();
        let mut central = Vec::new();
        push_u32(&mut central, 0x0201_4b50);
        push_u16(&mut central, 45);
        push_u16(&mut central, 45);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, crate::crc32(payload));
        push_u32(&mut central, u32::MAX);
        push_u32(&mut central, u32::MAX);
        push_u16(&mut central, u16::try_from(name.len()).unwrap());
        push_u16(&mut central, u16::try_from(extra.len()).unwrap());
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, 0);
        push_u32(&mut central, 0);
        central.extend_from_slice(name);
        central.extend_from_slice(&extra);
        let central_size = u32::try_from(central.len()).unwrap();
        archive.extend_from_slice(&central);
        push_u32(&mut archive, 0x0605_4b50);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 1);
        push_u16(&mut archive, 1);
        push_u32(&mut archive, central_size);
        push_u32(&mut archive, central_directory_offset);
        push_u16(&mut archive, 0);

        let reader = ArchiveReader::new(&archive).unwrap();
        assert_eq!(
            reader.read_stored_borrowed("stored.bin").unwrap(),
            Some(&payload[..])
        );

        let physical_archive = ZipArchive::from_slice(archive.as_slice()).unwrap();
        let physical = physical_archive
            .get_entry_borrowed(
                physical_archive
                    .entries()
                    .next()
                    .unwrap()
                    .unwrap()
                    .wayfinder(),
            )
            .unwrap()
            .data();
        let borrowed = reader.read_stored_borrowed("stored.bin").unwrap().unwrap();
        assert_eq!(borrowed.as_ptr(), physical.as_ptr());
        assert_eq!(borrowed.len(), physical.len());

        let local = local_header_offset_for_name(&archive, name);
        let local_extra_start = local + 30 + name.len();
        let mut truncated = archive.clone();
        truncated[local + 28..local + 30].copy_from_slice(&12u16.to_le_bytes());
        let error = ArchiveReader::new(&truncated)
            .unwrap()
            .read_stored_borrowed("stored.bin")
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let mut mismatched = archive.clone();
        mismatched[local_extra_start + 4..local_extra_start + 12]
            .copy_from_slice(&(size + 1).to_le_bytes());
        let error = ArchiveReader::new(&mismatched)
            .unwrap()
            .read_stored_borrowed("stored.bin")
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidSize { .. }));

        let duplicate_field = archive[local_extra_start..local_extra_start + extra.len()].to_vec();
        let payload_offset = local_extra_start + extra.len();
        let mut duplicated = archive.clone();
        duplicated.splice(payload_offset..payload_offset, duplicate_field);
        duplicated[local + 28..local + 30].copy_from_slice(&40u16.to_le_bytes());
        let eocd = duplicated.len() - 22;
        let central_offset = u32::from_le_bytes([
            duplicated[eocd + 16],
            duplicated[eocd + 17],
            duplicated[eocd + 18],
            duplicated[eocd + 19],
        ])
        .checked_add(u32::try_from(extra.len()).unwrap())
        .unwrap();
        duplicated[eocd + 16..eocd + 20].copy_from_slice(&central_offset.to_le_bytes());
        let error = ArchiveReader::new(&duplicated)
            .unwrap()
            .read_stored_borrowed("stored.bin")
            .unwrap_err();
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
    }

    fn stored_descriptor_fixture(payload: &[u8], signature: bool) -> Vec<u8> {
        stored_descriptor_fixture_with_prefix(payload, signature, 0)
    }

    fn stored_descriptor_fixture_with_prefix(
        payload: &[u8],
        signature: bool,
        prefix_len: usize,
    ) -> Vec<u8> {
        let name = b"stored.bin";
        let size = u32::try_from(payload.len()).unwrap();
        let crc = crate::crc32(payload);
        let local_header_offset = u32::try_from(prefix_len).unwrap();
        let mut archive = vec![0xA5; prefix_len];

        push_u32(&mut archive, 0x0403_4b50);
        push_u16(&mut archive, 20);
        push_u16(&mut archive, 0x08);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u32(&mut archive, 0);
        push_u32(&mut archive, 0);
        push_u32(&mut archive, 0);
        push_u16(&mut archive, u16::try_from(name.len()).unwrap());
        push_u16(&mut archive, 0);
        archive.extend_from_slice(name);
        archive.extend_from_slice(payload);
        if signature {
            push_u32(&mut archive, 0x0807_4b50);
        }
        push_u32(&mut archive, crc);
        push_u32(&mut archive, size);
        push_u32(&mut archive, size);

        let central_directory_offset = u32::try_from(archive.len()).unwrap();
        let mut central_directory = Vec::new();
        push_u32(&mut central_directory, 0x0201_4b50);
        push_u16(&mut central_directory, 20);
        push_u16(&mut central_directory, 20);
        push_u16(&mut central_directory, 0x08);
        push_u16(&mut central_directory, 0);
        push_u16(&mut central_directory, 0);
        push_u16(&mut central_directory, 0);
        push_u32(&mut central_directory, crc);
        push_u32(&mut central_directory, size);
        push_u32(&mut central_directory, size);
        push_u16(&mut central_directory, u16::try_from(name.len()).unwrap());
        push_u16(&mut central_directory, 0);
        push_u16(&mut central_directory, 0);
        push_u16(&mut central_directory, 0);
        push_u16(&mut central_directory, 0);
        push_u32(&mut central_directory, 0);
        push_u32(&mut central_directory, local_header_offset);
        central_directory.extend_from_slice(name);
        let central_directory_size = u32::try_from(central_directory.len()).unwrap();
        archive.extend_from_slice(&central_directory);

        push_u32(&mut archive, 0x0605_4b50);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 1);
        push_u16(&mut archive, 1);
        push_u32(&mut archive, central_directory_size);
        push_u32(&mut archive, central_directory_offset);
        push_u16(&mut archive, 0);
        archive
    }

    fn descriptor_start(payload: &[u8], signature: bool) -> usize {
        30 + b"stored.bin".len() + payload.len() + if signature { 4 } else { 0 }
    }

    fn local_header_offset_for_name(archive: &[u8], wanted_name: &[u8]) -> usize {
        const LOCAL_HEADER: [u8; 4] = 0x0403_4b50_u32.to_le_bytes();
        archive
            .windows(4)
            .enumerate()
            .find_map(|(offset, signature)| {
                if signature != LOCAL_HEADER || offset.saturating_add(30) > archive.len() {
                    return None;
                }
                let name_len = usize::from(u16::from_le_bytes([
                    archive[offset + 26],
                    archive[offset + 27],
                ]));
                let name_end = offset.saturating_add(30).saturating_add(name_len);
                (name_end <= archive.len() && &archive[offset + 30..name_end] == wanted_name)
                    .then_some(offset)
            })
            .expect("local member header")
    }

    fn central_header_offset_for_name(archive: &[u8], wanted_name: &[u8]) -> usize {
        const CENTRAL_HEADER: [u8; 4] = 0x0201_4b50_u32.to_le_bytes();
        archive
            .windows(4)
            .enumerate()
            .find_map(|(offset, signature)| {
                if signature != CENTRAL_HEADER || offset.saturating_add(46) > archive.len() {
                    return None;
                }
                let name_len = usize::from(u16::from_le_bytes([
                    archive[offset + 28],
                    archive[offset + 29],
                ]));
                let name_end = offset.saturating_add(46).saturating_add(name_len);
                (name_end <= archive.len() && &archive[offset + 46..name_end] == wanted_name)
                    .then_some(offset)
            })
            .expect("central member header")
    }

    #[test]
    fn indexed_strict_streams_reuse_the_bounded_layout_path() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", b"abc").unwrap();
        writer.write_deflated("deflated.bin", b"xyz").unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let indexed = indexed_archive(bytes);
        let stored = indexed.entry_id("stored.bin").unwrap();
        let deflated = indexed.entry_id("deflated.bin").unwrap();

        let mut stored_sink = Vec::new();
        let mut stored_accounting = ZipOperationAccounting::default();
        assert_eq!(
            indexed
                .read_entry_to_with_accounting(stored, &mut stored_sink, &mut stored_accounting,)
                .unwrap(),
            3
        );
        assert_eq!(stored_sink, b"abc");
        assert_eq!(stored_accounting.stored_payload_bytes_read(), 3);
        assert_eq!(stored_accounting.stored_payload_bytes_accepted(), 3);
        assert_eq!(stored_accounting.compressed_deflate_payload_bytes_read(), 0);
        assert_eq!(stored_accounting.deflate_bytes_produced(), 0);
        assert_eq!(stored_accounting.deflate_bytes_accepted(), 0);
        assert!(indexed.strict_layout_cache.is_ready());

        let mut repeated_stored_sink = Vec::new();
        let mut repeated_stored_accounting = ZipOperationAccounting::default();
        assert_eq!(
            indexed
                .read_entry_to_with_accounting(
                    stored,
                    &mut repeated_stored_sink,
                    &mut repeated_stored_accounting,
                )
                .unwrap(),
            3
        );
        assert_eq!(repeated_stored_sink, b"abc");
        assert_eq!(repeated_stored_accounting.stored_payload_bytes_read(), 3);
        assert_eq!(
            repeated_stored_accounting.stored_payload_bytes_accepted(),
            3
        );

        let mut deflated_sink = Vec::new();
        let mut deflated_accounting = ZipOperationAccounting::default();
        assert_eq!(
            indexed
                .read_entry_to_with_accounting(
                    deflated,
                    &mut deflated_sink,
                    &mut deflated_accounting,
                )
                .unwrap(),
            3
        );
        assert_eq!(deflated_sink, b"xyz");
        let deflated_metadata = indexed.metadata_for(deflated).unwrap();
        assert_eq!(
            deflated_accounting.compressed_deflate_payload_bytes_read(),
            deflated_metadata.compressed_size()
        );
        assert_eq!(deflated_accounting.deflate_bytes_produced(), 3);
        assert_eq!(deflated_accounting.deflate_bytes_accepted(), 3);
        assert_eq!(deflated_accounting.stored_payload_bytes_read(), 0);
        assert_eq!(deflated_accounting.stored_payload_bytes_accepted(), 0);
        assert!(indexed.strict_layout_cache.is_ready());
    }

    #[test]
    fn indexed_strict_first_read_uses_one_single_flight_builder() {
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", b"abc").unwrap();
        let indexed = Arc::new(indexed_archive(writer.finish_to_bytes().unwrap()));
        let entry_id = indexed.entry_id("stored.bin").unwrap();
        let barrier = Arc::new(std::sync::Barrier::new(4));
        let mut joins = Vec::new();
        for _ in 0..4 {
            let indexed = Arc::clone(&indexed);
            let barrier = Arc::clone(&barrier);
            joins.push(std::thread::spawn(move || {
                barrier.wait();
                let mut sink = Vec::new();
                indexed.read_entry_to(entry_id, &mut sink).unwrap();
                sink
            }));
        }
        for join in joins {
            assert_eq!(join.join().unwrap(), b"abc");
        }
        assert_eq!(indexed.strict_layout_cache.build_count(), 1);
    }

    // ---------------------------------------------------------------------
    // Change 0580: target-scoped strict layout proof.
    //
    // The fixtures below are the ones no existing test builds: a record whose
    // *local* variable region is inflated so its declared span swallows other
    // records.  Local name length and local extra length are the two span
    // inputs the central directory never carries, so they are the only way to
    // build an overlap that a central-directory-only analysis cannot see.
    // ---------------------------------------------------------------------

    /// One member of a layout-scope fixture.
    struct ScopedMember {
        name: Vec<u8>,
        payload: Vec<u8>,
        /// The `file_name_length` the *local* header declares. A value larger
        /// than `name.len()` extends the declared local span without changing
        /// any byte the central directory carries.
        local_name_len: u16,
        /// The `extra_field_length` the *local* header declares.
        local_extra_len: u16,
        /// The `compressed_size` the *local* header declares, when it is to
        /// differ from the central record's. This is change 0583's input: the
        /// two records describe one physical payload region, and only the
        /// local one is what a streaming reader follows.
        local_compressed_size: Option<u32>,
        /// The `uncompressed_size` the *local* header declares, when it is to
        /// differ from the central record's.
        local_uncompressed_size: Option<u32>,
        /// The general-purpose bit flags the *local* header declares, when they
        /// are to differ from the central record's. Bit 3 is the one that
        /// matters here: it decides whether a reader following local headers
        /// believes this record's sizes or looks for a data descriptor.
        local_flags: Option<u16>,
        /// The CRC the *local* header declares, when it is to differ from what
        /// the descriptor setting would write.
        local_crc: Option<u32>,
        /// Filler bytes written between the payload and the data descriptor, so
        /// the descriptor does not sit where the central payload length places
        /// it.
        descriptor_gap: usize,
        /// Where this member's local record is written. `None` appends it
        /// after everything written so far.
        at: Option<usize>,
        /// `Some(signed)` writes this member with general-purpose bit 3 set,
        /// zeroed local sizes and CRC, and a trailing data descriptor.
        descriptor: Option<bool>,
    }

    impl ScopedMember {
        fn new(name: &[u8], payload: &[u8]) -> Self {
            Self {
                name: name.to_vec(),
                payload: payload.to_vec(),
                local_name_len: u16::try_from(name.len()).unwrap(),
                local_extra_len: 0,
                local_compressed_size: None,
                local_uncompressed_size: None,
                local_flags: None,
                local_crc: None,
                descriptor_gap: 0,
                at: None,
                descriptor: None,
            }
        }

        fn local_compressed_size(mut self, size: u32) -> Self {
            self.local_compressed_size = Some(size);
            self
        }

        fn local_uncompressed_size(mut self, size: u32) -> Self {
            self.local_uncompressed_size = Some(size);
            self
        }

        fn local_flags(mut self, flags: u16) -> Self {
            self.local_flags = Some(flags);
            self
        }

        fn local_crc(mut self, crc: u32) -> Self {
            self.local_crc = Some(crc);
            self
        }

        fn descriptor_gap(mut self, gap: usize) -> Self {
            self.descriptor_gap = gap;
            self
        }

        fn descriptor(mut self, signed: bool) -> Self {
            self.descriptor = Some(signed);
            self
        }

        fn flags(&self) -> u16 {
            if self.descriptor.is_some() { 0x08 } else { 0 }
        }

        fn local_name_len(mut self, len: u16) -> Self {
            self.local_name_len = len;
            self
        }

        fn local_extra_len(mut self, len: u16) -> Self {
            self.local_extra_len = len;
            self
        }

        fn at(mut self, offset: usize) -> Self {
            self.at = Some(offset);
            self
        }

        fn local_record(&self) -> Vec<u8> {
            let size = u32::try_from(self.payload.len()).unwrap();
            let mut record = Vec::new();
            push_u32(&mut record, 0x0403_4b50);
            push_u16(&mut record, 20);
            push_u16(
                &mut record,
                self.local_flags.unwrap_or_else(|| self.flags()),
            );
            push_u16(&mut record, 0);
            push_u16(&mut record, 0);
            push_u16(&mut record, 0);
            let declared = if self.descriptor.is_some() { 0 } else { size };
            push_u32(
                &mut record,
                self.local_crc.unwrap_or(if self.descriptor.is_some() {
                    0
                } else {
                    crate::crc32(&self.payload)
                }),
            );
            push_u32(&mut record, self.local_compressed_size.unwrap_or(declared));
            push_u32(
                &mut record,
                self.local_uncompressed_size.unwrap_or(declared),
            );
            push_u16(&mut record, self.local_name_len);
            push_u16(&mut record, self.local_extra_len);
            let variable = usize::from(self.local_name_len) + usize::from(self.local_extra_len);
            let mut region = vec![0u8; variable];
            let copied = self.name.len().min(variable);
            region[..copied].copy_from_slice(&self.name[..copied]);
            record.extend_from_slice(&region);
            record.extend_from_slice(&self.payload);
            if let Some(signed) = self.descriptor {
                record.extend_from_slice(&vec![0u8; self.descriptor_gap]);
                if signed {
                    push_u32(&mut record, 0x0807_4b50);
                }
                push_u32(&mut record, crate::crc32(&self.payload));
                push_u32(&mut record, size);
                push_u32(&mut record, size);
            }
            record
        }
    }

    /// Assemble an archive whose members may be written at explicit offsets,
    /// so one member's declared local span can physically contain another's
    /// complete local record.
    fn scoped_fixture(members: &[ScopedMember]) -> Vec<u8> {
        let mut archive: Vec<u8> = Vec::new();
        let mut offsets = Vec::new();
        for member in members {
            let record = member.local_record();
            let offset = member.at.unwrap_or(archive.len());
            let end = offset + record.len();
            if archive.len() < end {
                archive.resize(end, 0);
            }
            archive[offset..end].copy_from_slice(&record);
            offsets.push(offset);
        }

        let central_directory_offset = u32::try_from(archive.len()).unwrap();
        let mut central_directory = Vec::new();
        for (member, offset) in members.iter().zip(&offsets) {
            let size = u32::try_from(member.payload.len()).unwrap();
            push_u32(&mut central_directory, 0x0201_4b50);
            push_u16(&mut central_directory, 20);
            push_u16(&mut central_directory, 20);
            push_u16(&mut central_directory, member.flags());
            push_u16(&mut central_directory, 0);
            push_u16(&mut central_directory, 0);
            push_u16(&mut central_directory, 0);
            push_u32(&mut central_directory, crate::crc32(&member.payload));
            push_u32(&mut central_directory, size);
            push_u32(&mut central_directory, size);
            push_u16(
                &mut central_directory,
                u16::try_from(member.name.len()).unwrap(),
            );
            push_u16(&mut central_directory, 0);
            push_u16(&mut central_directory, 0);
            push_u16(&mut central_directory, 0);
            push_u16(&mut central_directory, 0);
            push_u32(&mut central_directory, 0);
            push_u32(&mut central_directory, u32::try_from(*offset).unwrap());
            central_directory.extend_from_slice(&member.name);
        }
        let central_directory_size = u32::try_from(central_directory.len()).unwrap();
        archive.extend_from_slice(&central_directory);
        push_u32(&mut archive, 0x0605_4b50);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        let count = u16::try_from(members.len()).unwrap();
        push_u16(&mut archive, count);
        push_u16(&mut archive, count);
        push_u32(&mut archive, central_directory_size);
        push_u32(&mut archive, central_directory_offset);
        push_u16(&mut archive, 0);
        archive
    }

    /// Change 0575's adversarial witness, byte for byte.
    ///
    /// `docs/performance/results/change-0575/overlap_witness.py` builds the
    /// same 4,405-byte archive. `A.bin`'s local extra field is 4,096 bytes, so
    /// A's declared local span `[0, 4163)` physically contains `B.bin`'s entire
    /// local record at offset 2,048. A's *central* record declares an extra
    /// length of 0, so no central-directory-only analysis can see the overlap.
    /// `C.bin` begins exactly where A's span ends and overlaps nothing.
    fn overlap_witness_archive() -> Vec<u8> {
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"A.bin", &[b'A'; 32]).local_extra_len(4096),
            ScopedMember::new(b"B.bin", &[b'B'; 32]).at(2048),
            ScopedMember::new(b"C.bin", &[b'C'; 32]).at(4163),
        ]);
        // The same unknown extra-field header the Python witness writes, so
        // the two archives are byte-identical.  `0xFACE` is not a defined
        // extra-field id, so every conforming reader skips it.
        let mut bytes = bytes;
        bytes[35..37].copy_from_slice(&0xFACE_u16.to_le_bytes());
        bytes[37..39].copy_from_slice(&4092_u16.to_le_bytes());
        assert_eq!(bytes.len(), 4405, "witness archive size is pinned to 0575");
        bytes
    }

    fn scoped_indexed_read(bytes: &[u8], name: &str) -> Result<Vec<u8>, Error> {
        let length = bytes.len() as u64;
        let indexed = IndexedArchive::from_reader_with_limits(
            std::io::Cursor::new(bytes.to_vec()),
            length,
            ArchiveLimits::UNBOUNDED,
        )?;
        let entry_id = indexed
            .entry_id(name)
            .ok_or_else(|| Error::from(ErrorKind::FileNotFound(name.to_string())))?;
        let mut sink = Vec::new();
        indexed.read_entry_to(entry_id, &mut sink)?;
        Ok(sink)
    }

    fn scoped_borrowed_read(bytes: &[u8], name: &str) -> Result<Vec<u8>, Error> {
        let reader = ArchiveReader::new(bytes)?;
        let mut sink = Vec::new();
        reader.read_to(name, &mut sink)?;
        Ok(sink)
    }

    fn scoped_invalid_input_message(error: &Error) -> String {
        match error.kind() {
            ErrorKind::InvalidInput { msg } => msg.clone(),
            other => panic!("expected InvalidInput, found {other:?}"),
        }
    }

    #[derive(Debug)]
    struct ScopedCountingReaderAt {
        bytes: Vec<u8>,
        reads: Mutex<Vec<(u64, usize)>>,
    }

    impl ScopedCountingReaderAt {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                reads: Mutex::new(Vec::new()),
            }
        }

        fn clear(&self) {
            self.reads.lock().unwrap().clear();
        }

        fn calls(&self) -> usize {
            self.reads.lock().unwrap().len()
        }
    }

    impl ReaderAt for ScopedCountingReaderAt {
        fn read_at(&self, buf: &mut [u8], offset: u64) -> io::Result<usize> {
            let start = usize::try_from(offset).unwrap_or(self.bytes.len());
            let count = if start >= self.bytes.len() {
                0
            } else {
                let count = buf.len().min(self.bytes.len() - start);
                buf[..count].copy_from_slice(&self.bytes[start..start + count]);
                count
            };
            self.reads.lock().unwrap().push((offset, count));
            Ok(count)
        }
    }

    /// Positional reads the strict-layout proof alone issues for one target.
    fn scoped_strict_layout_reads(bytes: &[u8], name: &str) -> usize {
        let length = bytes.len() as u64;
        let indexed = IndexedArchive::from_reader_with_limits(
            ScopedCountingReaderAt::new(bytes.to_vec()),
            length,
            ArchiveLimits::UNBOUNDED,
        )
        .expect("fixture indexes");
        let entry_id = indexed.entry_id(name).expect("member is present");
        let wayfinder = indexed
            .indexed_entry(entry_id)
            .expect("entry id resolves")
            .info
            .wayfinder;
        indexed.archive.get_ref().clear();
        indexed
            .strict_layout_for(wayfinder)
            .expect("target proves its own layout");
        indexed.archive.get_ref().calls()
    }

    #[test]
    fn target_scoped_layout_refuses_both_overlapping_members_and_admits_the_third() {
        let bytes = overlap_witness_archive();

        // A and B overlap each other: A's declared local span contains B's
        // entire local record.  Reading either one is refused from its own
        // side of the overlap, exactly as the archive-wide proof refused them.
        for name in ["A.bin", "B.bin"] {
            for error in [
                scoped_indexed_read(&bytes, name).unwrap_err(),
                scoped_borrowed_read(&bytes, name).unwrap_err(),
            ] {
                assert_eq!(
                    scoped_invalid_input_message(&error),
                    "strict streaming refuses overlapping ZIP local spans",
                    "{name} participates in the overlap and must stay refused",
                );
            }
        }

        // C overlaps nothing.  This is the whole semantic delta of change
        // 0580: the archive-wide proof refused this read because two members
        // C does not touch overlap each other.
        assert_eq!(
            scoped_indexed_read(&bytes, "C.bin").unwrap(),
            vec![b'C'; 32]
        );
        assert_eq!(
            scoped_borrowed_read(&bytes, "C.bin").unwrap(),
            vec![b'C'; 32]
        );
    }

    #[test]
    fn a_distant_predecessor_that_reaches_the_target_is_still_refused() {
        // `first.bin` declares a 16 KiB local extra field, so its span runs to
        // offset 16,425 and contains three later members' complete local
        // records.  The central directory declares no local extra field for
        // it, so its zero-I/O bracket is `[0, 43)` at minimum: only
        // `first.bin`'s own 30-byte local header reveals the reach.
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"first.bin", b"first payload").local_extra_len(16 * 1024),
            ScopedMember::new(b"second.bin", b"second").at(4096),
            ScopedMember::new(b"third.bin", b"third").at(8192),
            ScopedMember::new(b"fourth.bin", b"fourth").at(12288),
        ]);

        for name in ["second.bin", "third.bin", "fourth.bin"] {
            let error = scoped_indexed_read(&bytes, name).unwrap_err();
            assert_eq!(
                scoped_invalid_input_message(&error),
                "strict streaming refuses overlapping ZIP local spans",
                "{name} lies inside first.bin's declared local span",
            );
            let error = scoped_borrowed_read(&bytes, name).unwrap_err();
            assert_eq!(
                scoped_invalid_input_message(&error),
                "strict streaming refuses overlapping ZIP local spans",
            );
        }
    }

    #[test]
    fn a_predecessor_reaching_past_the_central_name_length_bracket_is_still_refused() {
        // The residual window has to count *both* halves of the local variable
        // region.  Here `first.bin` declares a 100-byte local name and a
        // 65,535-byte local extra field, so its declared span ends at 65,697.
        // A bracket that assumed the local name length equals the 9-byte
        // central name length would stop at 65,630 + 24 and prune this record,
        // admitting a read of bytes it claims.
        let target_offset = 65_640;
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"first.bin", &[b'f'; 32])
                .local_name_len(100)
                .local_extra_len(u16::MAX),
            ScopedMember::new(b"target.bin", b"target").at(target_offset),
        ]);
        // The reach is real: first.bin's declared span end is past the target.
        assert!(30 + 100 + usize::from(u16::MAX) + 32 > target_offset);
        // And it is past where a central-name-length bracket would stop.
        assert!(target_offset > 30 + "first.bin".len() + 32 + 65_535 + 24);

        let error = scoped_indexed_read(&bytes, "target.bin").unwrap_err();
        assert_eq!(
            scoped_invalid_input_message(&error),
            "strict streaming refuses overlapping ZIP local spans",
        );
        let error = scoped_borrowed_read(&bytes, "target.bin").unwrap_err();
        assert_eq!(
            scoped_invalid_input_message(&error),
            "strict streaming refuses overlapping ZIP local spans",
        );
    }

    // ---------------------------------------------------------------------
    // Change 0583: a neighbour's payload length is the larger of its local
    // and central declarations.
    //
    // Change 0580 bounds a neighbour's span with the local variable-region
    // length — which it reads — and the *central* payload length, which it
    // does not cross-check against the local header's own `compressed_size`.
    // Change 0582's finding 2 showed that a local header declaring a longer
    // payload than its central record hides bytes a streaming reader assigns
    // to it, and lets another member be read out of them.
    // ---------------------------------------------------------------------

    #[test]
    fn a_predecessor_whose_local_compressed_size_reaches_the_target_is_refused() {
        // Change 0582's `crafted/neighbour-local-csize-smuggles.zip`, rebuilt.
        // `pred.bin` carries a 16-byte payload in its central record and
        // declares a 100,000-byte payload in its local header.  A reader that
        // trusts local headers places its payload at [38, 100038), which
        // contains `target.bin`'s entire local record; a reader that trusts
        // the central directory places it at [38, 54), which does not.
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"pred.bin", &[b'p'; 16]).local_compressed_size(100_000),
            ScopedMember::new(b"target.bin", &[b't'; 16]).at(200),
        ]);
        // Central metadata alone cannot see the reach: pred's central payload
        // ends at 54, far short of the target.
        assert!(30 + "pred.bin".len() + 16 < 200);
        // Only the two bytes at offset 18..22 of pred's local header do.
        assert_eq!(30 + "pred.bin".len() + 100_000, 100_038);

        for error in [
            scoped_indexed_read(&bytes, "target.bin").unwrap_err(),
            scoped_borrowed_read(&bytes, "target.bin").unwrap_err(),
        ] {
            assert_eq!(
                scoped_invalid_input_message(&error),
                "strict streaming refuses overlapping ZIP local spans",
                "target.bin lies inside the payload pred.bin's local header claims",
            );
        }
    }

    #[test]
    fn a_descriptor_bearing_predecessors_payload_end_is_not_moved() {
        // A declared data descriptor's payload end is not only a refusal
        // threshold: it is the offset the descriptor is *read* at.  Moving it
        // does not make the bound larger, it makes it different, and a
        // descriptor parsed somewhere else can match where the true one did
        // not — so a maximum applied to this branch can turn a refusal into an
        // acceptance.  The maximum is therefore confined to the exact branch.
        //
        // `pred.bin`'s central record declares a descriptor; its *local*
        // header does not, and declares a payload four bytes longer than the
        // central one.  Four filler bytes sit between the payload and the real
        // 12-byte descriptor, so the descriptor is exactly where the inflated
        // local length would place it and exactly not where the central length
        // does.  `target.bin` begins two bytes past the real span end.
        let payload = [b'p'; 16];
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"pred.bin", &payload)
                .descriptor(false)
                .descriptor_gap(4)
                .local_flags(0)
                .local_crc(crate::crc32(&payload))
                .local_compressed_size(20)
                .local_uncompressed_size(20),
            ScopedMember::new(b"target.bin", &[b't'; 16]).at(72),
        ]);
        // 30 + 8 + 16 = 54 is where the central length puts the payload end;
        // 54..58 is filler and the real unsigned descriptor starts at 58,
        // which is 30 + 8 + 20, where the inflated local length points.
        assert_eq!(30 + "pred.bin".len() + payload.len(), 54);
        assert_eq!(&bytes[54..58], &[0u8; 4]);
        assert_eq!(&bytes[58..62], &crate::crc32(&payload).to_le_bytes());
        assert_eq!(&bytes[62..66], &16_u32.to_le_bytes());
        assert_eq!(&bytes[66..70], &16_u32.to_le_bytes());

        // Resolving at 54 finds filler, not a descriptor, so the neighbour
        // cannot be bounded and the read is refused.  Resolving at 58 would
        // find a valid descriptor ending at 70, two bytes clear of the target,
        // and the read would succeed.
        assert!(scoped_indexed_read(&bytes, "target.bin").is_err());
        assert!(scoped_borrowed_read(&bytes, "target.bin").is_err());
    }

    #[test]
    fn a_descriptor_bearing_predecessors_local_size_is_not_a_span_length() {
        // The same inflated local size with general-purpose bit 3 set in both
        // records.  Bit 3 means the true sizes are in the trailing data
        // descriptor and these fields are placeholders, so a reader following
        // local headers scans for the descriptor instead of trusting this
        // field: inflating it moves no boundary.
        //
        // It is also what keeps a verdict independent of read order.  Full
        // validation skips the local-versus-central size comparison for a
        // descriptor-bearing record, so `pred.bin` validates as a target and
        // its exact span end is memoised as its neighbour bound.  A maximum
        // applied here would disagree with that memo, and this archive would
        // accept or refuse `target.bin` depending on which member was read
        // first.
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"pred.bin", &[b'p'; 16])
                .descriptor(false)
                .local_compressed_size(100_000),
            ScopedMember::new(b"target.bin", &[b't'; 16]).at(200),
        ]);

        // A fresh reader that has read nothing else.
        assert_eq!(
            scoped_indexed_read(&bytes, "target.bin").unwrap(),
            vec![b't'; 16]
        );
        assert_eq!(
            scoped_borrowed_read(&bytes, "target.bin").unwrap(),
            vec![b't'; 16]
        );

        // And a reader that has already proven `pred.bin`'s own layout, which
        // is the order that populates the memo.
        let length = bytes.len() as u64;
        let indexed = IndexedArchive::from_reader_with_limits(
            std::io::Cursor::new(bytes.clone()),
            length,
            ArchiveLimits::UNBOUNDED,
        )
        .unwrap();
        let reader = ArchiveReader::new(&bytes).unwrap();
        for name in ["pred.bin", "target.bin"] {
            let entry_id = indexed.entry_id(name).unwrap();
            let mut sink = Vec::new();
            indexed.read_entry_to(entry_id, &mut sink).unwrap();
            assert_eq!(sink.len(), 16, "{name} reads its own 16 bytes");
            let mut sink = Vec::new();
            reader.read_to(name, &mut sink).unwrap();
            assert_eq!(sink.len(), 16);
        }
    }

    #[test]
    fn a_predecessor_whose_local_compressed_size_understates_keeps_the_central_bound() {
        // The maximum runs in both directions, and this pins the direction the
        // central record wins.  `pred.bin` carries a 100-byte central payload
        // and declares 16 locally.  The target sits at 134, which is:
        //
        //   * past `30 + central = 130`, so the zero-I/O bracket does not
        //     refuse it and the local header really is read;
        //   * past `30 + 8 + 16 = 54`, the span end the *local* claim gives,
        //     so a bound that took the local value would accept;
        //   * inside `30 + 8 + 100 = 138`, the span end the central claim
        //     gives, so the maximum refuses.
        let central_payload = [b'p'; 100];
        let local_claim = 16_usize;
        let target_offset = 134_usize;
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"pred.bin", &central_payload)
                .local_compressed_size(u32::try_from(local_claim).unwrap()),
            ScopedMember::new(b"target.bin", &[b't'; 16]).at(target_offset),
        ]);
        let fixed = 30 + "pred.bin".len();
        assert!(
            30 + central_payload.len() <= target_offset,
            "the zero-I/O bracket must not settle this",
        );
        assert!(
            fixed + local_claim < target_offset,
            "the local claim clears the target",
        );
        assert!(
            fixed + central_payload.len() > target_offset,
            "the central claim does not",
        );

        for error in [
            scoped_indexed_read(&bytes, "target.bin").unwrap_err(),
            scoped_borrowed_read(&bytes, "target.bin").unwrap_err(),
        ] {
            assert_eq!(
                scoped_invalid_input_message(&error),
                "strict streaming refuses overlapping ZIP local spans",
            );
        }
    }

    #[test]
    fn a_local_zip64_size_sentinel_is_not_a_neighbour_span_length() {
        // `u32::MAX` in a local `compressed_size` is the ZIP64 sentinel, not a
        // length: the real value lives in a ZIP64 extra field inside the
        // variable region, which the neighbour probe deliberately does not
        // read.  A maximum that took the sentinel literally would reserve
        // 4 GiB for every ZIP64 member and refuse valid archives, so the
        // sentinel falls back to the central length — exactly the bound
        // change 0580 computed.
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"pred.bin", &[b'p'; 32]).local_compressed_size(u32::MAX),
            ScopedMember::new(b"target.bin", &[b't'; 16]).at(200),
        ]);
        // Taken literally the sentinel would place pred's span end 4 GiB past
        // the target.
        assert!(u64::from(u32::MAX) > 200);

        assert_eq!(
            scoped_indexed_read(&bytes, "target.bin").unwrap(),
            vec![b't'; 16]
        );
        assert_eq!(
            scoped_borrowed_read(&bytes, "target.bin").unwrap(),
            vec![b't'; 16]
        );
    }

    #[test]
    fn a_predecessor_local_uncompressed_size_does_not_move_the_payload() {
        // `uncompressed_size` never describes an on-disk region — the payload
        // that separates one local record from the next is `compressed_size`
        // bytes long whatever the member inflates to — so an inflated local
        // value cannot move where any reader places a payload.  Change 0582's
        // `neighbour-local-usize-smuggles` is an ordinary change-0580
        // local-versus-central relaxation, and is deliberately not given the
        // compressed-size treatment.
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"pred.bin", &[b'p'; 16]).local_uncompressed_size(100_000),
            ScopedMember::new(b"target.bin", &[b't'; 16]).at(200),
        ]);

        assert_eq!(
            scoped_indexed_read(&bytes, "target.bin").unwrap(),
            vec![b't'; 16]
        );
        assert_eq!(
            scoped_borrowed_read(&bytes, "target.bin").unwrap(),
            vec![b't'; 16]
        );
    }

    #[test]
    fn a_verdict_does_not_depend_on_what_the_reader_read_before() {
        let bytes = overlap_witness_archive();
        let expected = [("A.bin", false), ("B.bin", false), ("C.bin", true)];
        let orders: [[&str; 3]; 6] = [
            ["A.bin", "B.bin", "C.bin"],
            ["A.bin", "C.bin", "B.bin"],
            ["B.bin", "A.bin", "C.bin"],
            ["B.bin", "C.bin", "A.bin"],
            ["C.bin", "A.bin", "B.bin"],
            ["C.bin", "B.bin", "A.bin"],
        ];

        for order in orders {
            // One reader, reading the three members in this order.
            let length = bytes.len() as u64;
            let indexed = IndexedArchive::from_reader_with_limits(
                std::io::Cursor::new(bytes.clone()),
                length,
                ArchiveLimits::UNBOUNDED,
            )
            .unwrap();
            let reader = ArchiveReader::new(&bytes).unwrap();
            let mut seen = Vec::new();
            for name in order {
                let entry_id = indexed.entry_id(name).unwrap();
                let mut sink = Vec::new();
                let indexed_ok = indexed.read_entry_to(entry_id, &mut sink).is_ok();
                let mut sink = Vec::new();
                let borrowed_ok = reader.read_to(name, &mut sink).is_ok();
                assert_eq!(
                    indexed_ok, borrowed_ok,
                    "both readers share one acceptance contract for {name}",
                );
                seen.push((name, indexed_ok));
            }
            seen.sort_by_key(|(name, _)| *name);
            assert_eq!(seen, expected.to_vec(), "order {order:?} changed a verdict",);

            // And the same three members, each on a reader that has read
            // nothing else.
            for (name, accepted) in expected {
                assert_eq!(
                    scoped_indexed_read(&bytes, name).is_ok(),
                    accepted,
                    "a fresh reader disagrees about {name}",
                );
            }
        }
    }

    #[test]
    fn duplicate_local_header_offsets_are_refused_on_the_strict_path() {
        // Two central records that declare the same local-header offset.  The
        // archive-wide proof reported this as an overlap; it is a pure
        // central-directory property, decidable with no read at all, and it
        // keeps its own refusal.
        let mut bytes = fixture(&[
            FixtureEntry::stored(b"first.bin", b"first"),
            FixtureEntry::stored(b"second.bin", b"second"),
        ]);
        let archive = ZipArchive::from_slice(&bytes).unwrap();
        let central = archive.directory_offset() as usize;
        let second = central + central_record_len(&bytes, central);
        bytes[second + 42..second + 46].copy_from_slice(&0u32.to_le_bytes());

        for name in ["first.bin", "second.bin"] {
            let error = scoped_indexed_read(&bytes, name).unwrap_err();
            assert_eq!(
                scoped_invalid_input_message(&error),
                "strict streaming refuses duplicate ZIP local spans",
            );
            let error = scoped_borrowed_read(&bytes, name).unwrap_err();
            assert_eq!(
                scoped_invalid_input_message(&error),
                "strict streaming refuses duplicate ZIP local spans",
            );
        }
    }

    #[test]
    fn a_descriptor_bearing_predecessor_is_resolved_exactly_not_conservatively() {
        // A data descriptor's encoded width is 12, 16, 20 or 24 bytes and is
        // not decidable from a fixed local header.  In a gapless archive the
        // predecessor's payload ends exactly one descriptor before the target,
        // so a bound that simply reserved the widest descriptor would refuse
        // every gapless descriptor-bearing archive.  The proof resolves the
        // width exactly instead.
        for signed in [false, true] {
            let bytes = scoped_fixture(&[
                ScopedMember::new(b"first.bin", b"descriptor payload").descriptor(signed),
                ScopedMember::new(b"target.bin", b"target"),
                ScopedMember::new(b"last.bin", b"last").descriptor(signed),
            ]);
            assert_eq!(
                scoped_indexed_read(&bytes, "target.bin").unwrap(),
                b"target".to_vec(),
                "a descriptor-bearing predecessor must not be treated as 24 bytes wide",
            );
            assert_eq!(
                scoped_indexed_read(&bytes, "first.bin").unwrap(),
                b"descriptor payload".to_vec(),
            );
            assert_eq!(
                scoped_indexed_read(&bytes, "last.bin").unwrap(),
                b"last".to_vec(),
            );
            assert_eq!(
                scoped_borrowed_read(&bytes, "target.bin").unwrap(),
                b"target".to_vec(),
            );
            assert_eq!(
                scoped_borrowed_read(&bytes, "last.bin").unwrap(),
                b"last".to_vec(),
            );
        }
    }

    #[test]
    fn a_predecessor_outside_the_residual_window_costs_no_read() {
        // Six tiny members, then one whose payload is larger than the residual
        // window, then the target.  Every member before the large one is
        // pruned by resident central metadata alone.
        let filler = vec![0xA5u8; 140_000];
        let mut members: Vec<ScopedMember> = (0..6)
            .map(|index| ScopedMember::new(format!("tiny{index}.bin").as_bytes(), b"tiny"))
            .collect();
        members.push(ScopedMember::new(b"large.bin", &filler));
        members.push(ScopedMember::new(b"target.bin", b"target"));
        let bytes = scoped_fixture(&members);

        // One read for the target's own layout, one 30-byte probe for the
        // large member.  The six tiny members are outside the window.
        assert_eq!(scoped_strict_layout_reads(&bytes, "target.bin"), 2);
        assert_eq!(
            scoped_indexed_read(&bytes, "target.bin").unwrap(),
            b"target".to_vec(),
        );
        // Every member still reads, and the first one needs no predecessor at
        // all.
        assert_eq!(scoped_strict_layout_reads(&bytes, "tiny0.bin"), 1);
    }

    #[test]
    fn opening_and_listing_issues_no_strict_layout_read() {
        let bytes = scoped_fixture(&[
            ScopedMember::new(b"first.bin", b"first"),
            ScopedMember::new(b"second.bin", b"second"),
            ScopedMember::new(b"third.bin", b"third"),
        ]);
        let length = bytes.len() as u64;
        let indexed = IndexedArchive::from_reader_with_limits(
            ScopedCountingReaderAt::new(bytes),
            length,
            ArchiveLimits::UNBOUNDED,
        )
        .unwrap();
        indexed.archive.get_ref().clear();
        assert_eq!(indexed.file_names().count(), 3);
        assert_eq!(indexed.archive.get_ref().calls(), 0);
        // The materializing read family does not reach the proof either.
        let entry_id = indexed.entry_id("second.bin").unwrap();
        assert_eq!(indexed.read_entry(entry_id).unwrap(), b"second".to_vec());
        assert!(!indexed.strict_layout_cache.is_ready());
    }

    #[test]
    fn reading_every_member_converges_on_the_archive_wide_verdict() {
        // Convergence: for each of these archives, the target-scoped verdict
        // of every member equals the archive-wide verdict, and no member's
        // local header is read twice when members are read in physical order.
        let large = vec![0x5Au8; 140_000];
        let spread: Vec<ScopedMember> = vec![
            ScopedMember::new(b"head.bin", b"head"),
            ScopedMember::new(b"large.bin", &large),
            ScopedMember::new(b"tail.bin", b"tail"),
        ];

        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("stored.bin", b"stored payload")
            .unwrap();
        writer
            .write_deflated("deflated.bin", b"deflated payload")
            .unwrap();
        writer.write_stored("empty.bin", b"").unwrap();
        let mixed = writer.finish_to_bytes().unwrap();

        let archives: Vec<(&str, Vec<u8>)> = vec![
            (
                "gapless",
                scoped_fixture(&[
                    ScopedMember::new(b"a.bin", b"aaaa"),
                    ScopedMember::new(b"b.bin", b"bbbb"),
                    ScopedMember::new(b"c.bin", b"cccc"),
                ]),
            ),
            ("spread", scoped_fixture(&spread)),
            ("store_and_deflate", mixed),
            (
                "descriptor",
                scoped_fixture(&[
                    ScopedMember::new(b"first.bin", b"first payload").descriptor(false),
                    ScopedMember::new(b"second.bin", b"second payload").descriptor(false),
                    ScopedMember::new(b"third.bin", b"third payload"),
                ]),
            ),
            (
                "descriptor_signed",
                scoped_fixture(&[
                    ScopedMember::new(b"first.bin", b"first payload").descriptor(true),
                    ScopedMember::new(b"second.bin", b"second payload").descriptor(true),
                ]),
            ),
            ("overlap_witness", overlap_witness_archive()),
        ];

        for (label, bytes) in archives {
            let length = bytes.len() as u64;
            let indexed = IndexedArchive::from_reader_with_limits(
                std::io::Cursor::new(bytes.clone()),
                length,
                ArchiveLimits::UNBOUNDED,
            )
            .unwrap();
            let targets: Vec<_> = indexed.layout.iter().map(|entry| entry.wayfinder).collect();

            // The archive-wide verdict, computed here from the same per-entry
            // validator the proof uses: every record's layout must prove, and
            // adjacent spans must not overlap.
            let mut archive_wide_ok = true;
            let mut previous_end: Option<u64> = None;
            for position in 0..indexed.layout.len() {
                let central_name = indexed.strict_layout_central_name(position).unwrap();
                let wayfinder = indexed.layout[position].wayfinder;
                match indexed.archive.validate_strict_entry_layout(
                    wayfinder,
                    central_name,
                    indexed.strict_layout_read_bound(position),
                ) {
                    Ok(span) => {
                        if previous_end.is_some_and(|end| span.local_header_offset < end) {
                            archive_wide_ok = false;
                            break;
                        }
                        previous_end = Some(span.span_end);
                    },
                    Err(_) => {
                        archive_wide_ok = false;
                        break;
                    },
                }
            }

            let scoped: Vec<bool> = targets
                .iter()
                .map(|target| indexed.strict_layout_for(*target).is_ok())
                .collect();
            if archive_wide_ok {
                assert!(
                    scoped.iter().all(|accepted| *accepted),
                    "{label}: an archive-wide accept must stay accepted member by member",
                );
            } else {
                assert!(
                    scoped.iter().any(|accepted| !*accepted),
                    "{label}: an archive-wide refusal must refuse at least one member",
                );
            }

            // Reading every member in physical order costs exactly what the
            // archive-wide proof cost: no record's local header is read twice,
            // because a validated target is its own exact span bound for the
            // members that follow it.
            if archive_wide_ok {
                let scoped = {
                    let indexed = IndexedArchive::from_reader_with_limits(
                        ScopedCountingReaderAt::new(bytes.clone()),
                        length,
                        ArchiveLimits::UNBOUNDED,
                    )
                    .unwrap();
                    let targets: Vec<_> =
                        indexed.layout.iter().map(|entry| entry.wayfinder).collect();
                    indexed.archive.get_ref().clear();
                    for target in targets {
                        indexed.strict_layout_for(target).unwrap();
                    }
                    indexed.archive.get_ref().calls()
                };
                let archive_wide = {
                    let indexed = IndexedArchive::from_reader_with_limits(
                        ScopedCountingReaderAt::new(bytes.clone()),
                        length,
                        ArchiveLimits::UNBOUNDED,
                    )
                    .unwrap();
                    indexed.archive.get_ref().clear();
                    for position in 0..indexed.layout.len() {
                        let central_name = indexed.strict_layout_central_name(position).unwrap();
                        indexed
                            .archive
                            .validate_strict_entry_layout(
                                indexed.layout[position].wayfinder,
                                central_name,
                                indexed.strict_layout_read_bound(position),
                            )
                            .unwrap();
                    }
                    indexed.archive.get_ref().calls()
                };
                assert_eq!(
                    scoped,
                    archive_wide,
                    "{label}: reading every member must cost what one \
                     archive-wide proof cost ({} members)",
                    targets.len(),
                );
            }
        }
    }

    #[test]
    fn zip64_and_descriptor_members_keep_their_target_scoped_proof() {
        // A ZIP64 member with a data descriptor, written by the crate's own
        // writer, still proves its layout through the target-scoped path, on
        // both readers, for Store and Deflate.
        let payload = vec![0x42u8; 4096];
        for deflate in [false, true] {
            let mut writer = StreamingArchiveWriter::new();
            writer.write_stored("first.bin", b"first").unwrap();
            if deflate {
                writer.write_deflated("member.bin", &payload).unwrap();
            } else {
                writer.write_stored("member.bin", &payload).unwrap();
            }
            writer.write_stored("last.bin", b"last").unwrap();
            let bytes = writer.finish_to_bytes().unwrap();

            assert_eq!(scoped_indexed_read(&bytes, "member.bin").unwrap(), payload);
            assert_eq!(scoped_borrowed_read(&bytes, "member.bin").unwrap(), payload);
            assert_eq!(scoped_indexed_read(&bytes, "first.bin").unwrap(), b"first");
            assert_eq!(scoped_indexed_read(&bytes, "last.bin").unwrap(), b"last");
        }

        for signature in [false, true] {
            let bytes = stored_descriptor_fixture(b"descriptor payload", signature);
            assert_eq!(
                scoped_indexed_read(&bytes, "stored.bin").unwrap(),
                b"descriptor payload".to_vec(),
            );
            assert_eq!(
                scoped_borrowed_read(&bytes, "stored.bin").unwrap(),
                b"descriptor payload".to_vec(),
            );
        }
    }

    #[test]
    fn a_failed_proof_memoises_nothing_and_the_next_read_retries() {
        let bytes = overlap_witness_archive();
        let length = bytes.len() as u64;
        let indexed = IndexedArchive::from_reader_with_limits(
            std::io::Cursor::new(bytes.clone()),
            length,
            ArchiveLimits::UNBOUNDED,
        )
        .unwrap();

        let refused = indexed.entry_id("A.bin").unwrap();
        let mut sink = Vec::new();
        assert!(indexed.read_entry_to(refused, &mut sink).is_err());
        // A refusal for one member neither poisons nor primes the reader.
        let mut sink = Vec::new();
        assert!(indexed.read_entry_to(refused, &mut sink).is_err());
        let admitted = indexed.entry_id("C.bin").unwrap();
        let mut sink = Vec::new();
        indexed.read_entry_to(admitted, &mut sink).unwrap();
        assert_eq!(sink, vec![b'C'; 32]);
    }

    #[test]
    fn rejects_inflated_entry_count_before_reserving_across_central_gap() {
        let mut bytes = fixture(&[FixtureEntry::stored(b"a", b"data")]);
        let eocd_offset = bytes.len() - 22;
        let gap = 4096;
        bytes.splice(eocd_offset..eocd_offset, vec![0xa5; gap]);
        let eocd_offset = eocd_offset + gap;
        bytes[eocd_offset + 8..eocd_offset + 10].copy_from_slice(&1000u16.to_le_bytes());
        bytes[eocd_offset + 10..eocd_offset + 12].copy_from_slice(&1000u16.to_le_bytes());

        let error = ArchiveReader::new_with_limits(&bytes, ArchiveLimits::UNBOUNDED)
            .expect_err("inflated EOCD count must be rejected");
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));

        let error = indexed_archive_result(bytes, ArchiveLimits::UNBOUNDED)
            .expect_err("inflated EOCD count must be rejected");
        assert!(matches!(error.kind(), ErrorKind::InvalidInput { .. }));
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

    fn assert_zero_crc_checksum(error: Error, actual_crc: u32) {
        match error.kind() {
            ErrorKind::InvalidChecksum { expected, actual } => {
                assert_eq!(*expected, 0);
                assert_eq!(*actual, actual_crc);
            },
            other => panic!("expected zero-CRC checksum error, got {other:?}"),
        }
    }

    fn assert_materialized_size_error(error: Error) {
        match error.kind() {
            ErrorKind::InvalidSize { .. } => {},
            ErrorKind::IO(io_error) | ErrorKind::Io(io_error) => {
                let source = io_error
                    .get_ref()
                    .and_then(|source| source.downcast_ref::<Error>());
                match source {
                    Some(source) => {
                        assert!(matches!(source.kind(), ErrorKind::InvalidSize { .. }));
                    },
                    None => panic!("expected nested ZIP size error, got {io_error}"),
                }
            },
            other => panic!("expected materialized size error, got {other:?}"),
        }
    }

    fn assert_materialized_checksum_error(error: Error) {
        match error.kind() {
            ErrorKind::InvalidChecksum { .. } => {},
            ErrorKind::IO(io_error) | ErrorKind::Io(io_error) => {
                let source = io_error
                    .get_ref()
                    .and_then(|source| source.downcast_ref::<Error>());
                match source {
                    Some(source) => {
                        assert!(matches!(source.kind(), ErrorKind::InvalidChecksum { .. }));
                    },
                    None => panic!("expected nested ZIP checksum error, got {io_error}"),
                }
            },
            other => panic!("expected materialized checksum error, got {other:?}"),
        }
    }

    fn skipped_crc_fixture(payload: &[u8]) -> Vec<u8> {
        fn write_entry<W: Write>(
            archive: &mut ZipArchiveWriter<W>,
            name: &str,
            compression_method: CompressionMethod,
            payload: &[u8],
        ) {
            let (mut entry, config) = archive
                .new_file(name)
                .compression_method(compression_method)
                .crc32(Crc32Option::Skip)
                .start()
                .unwrap();
            match compression_method {
                CompressionMethod::Store => {
                    let mut writer = config.wrap(&mut entry);
                    writer.write_all(payload).unwrap();
                    let (_, descriptor) = writer.finish().unwrap();
                    entry.finish(descriptor).unwrap();
                },
                CompressionMethod::Deflate => {
                    let encoder = DeflateEncoder::new(&mut entry, Compression::default());
                    let mut writer = config.wrap(encoder);
                    writer.write_all(payload).unwrap();
                    let (encoder, descriptor) = writer.finish().unwrap();
                    encoder.finish().unwrap();
                    entry.finish(descriptor).unwrap();
                },
                _ => unreachable!("test fixture uses only Store and Deflate"),
            }
        }

        let mut output = Cursor::new(Vec::new());
        let mut archive = ZipArchiveWriter::new(&mut output);
        write_entry(
            &mut archive,
            "stored.bin",
            CompressionMethod::Store,
            payload,
        );
        write_entry(
            &mut archive,
            "deflated.bin",
            CompressionMethod::Deflate,
            payload,
        );
        write_entry(
            &mut archive,
            "empty-stored.bin",
            CompressionMethod::Store,
            b"",
        );
        write_entry(
            &mut archive,
            "empty-deflated.bin",
            CompressionMethod::Deflate,
            b"",
        );
        archive.finish().unwrap();
        output.into_inner()
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

    fn oversized_metadata_fixture() -> (Vec<u8>, u64) {
        let name = vec![b'n'; 4 * 1024];
        let extra = vec![b'e'; 40 * 1024];
        let comment = vec![b'c'; 40 * 1024];
        let metadata_bytes = (name.len() + extra.len() + comment.len()) as u64;
        let bytes = fixture(&[FixtureEntry {
            name: &name,
            extra: &extra,
            comment: &comment,
            compressed_size: 7,
            uncompressed_size: 7,
            data: b"payload",
        }]);
        (bytes, metadata_bytes)
    }

    fn oversized_and_ordinary_fixture(oversized_first: bool) -> (Vec<u8>, u64) {
        let name = vec![b'n'; 4 * 1024];
        let extra = vec![b'e'; 40 * 1024];
        let comment = vec![b'c'; 40 * 1024];
        let metadata_bytes = (name.len() + extra.len() + comment.len()) as u64;
        let oversized = FixtureEntry {
            name: &name,
            extra: &extra,
            comment: &comment,
            compressed_size: 7,
            uncompressed_size: 7,
            data: b"payload",
        };
        let ordinary = FixtureEntry::stored(b"ordinary", b"");
        let entries = if oversized_first {
            [oversized, ordinary]
        } else {
            [ordinary, oversized]
        };
        (fixture(&entries), metadata_bytes)
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

    fn indexed_strict_error(bytes: Vec<u8>) -> Error {
        match indexed_archive_result(bytes, ArchiveLimits::UNBOUNDED) {
            Err(error) => error,
            Ok(indexed) => {
                let entry_id = indexed.entry_id("zip64.bin").unwrap();
                let mut sink = vec![0xA5];
                let mut accounting = ZipOperationAccounting::default();
                let error = indexed
                    .read_entry_to_with_accounting(entry_id, &mut sink, &mut accounting)
                    .unwrap_err();
                assert_eq!(sink, vec![0xA5]);
                assert_eq!(accounting, ZipOperationAccounting::default());
                error
            },
        }
    }

    fn zip64_descriptor_fixture(payload: &[u8], signature: bool) -> Vec<u8> {
        let extra = zip64_sizes(
            u64::try_from(payload.len()).unwrap(),
            u64::try_from(payload.len()).unwrap(),
        );
        zip64_descriptor_fixture_with_central_sizes(payload, signature, &extra, u32::MAX, u32::MAX)
    }

    fn zip64_descriptor_fixture_with_central_sizes(
        payload: &[u8],
        signature: bool,
        central_extra: &[u8],
        central_compressed_size: u32,
        central_uncompressed_size: u32,
    ) -> Vec<u8> {
        let name = b"zip64.bin";
        let size = u64::try_from(payload.len()).unwrap();
        let crc = crate::crc32(payload);
        let local_extra = zip64_sizes(size, size);
        let mut archive = Vec::new();

        push_u32(&mut archive, 0x0403_4b50);
        push_u16(&mut archive, 45);
        push_u16(&mut archive, 0x08);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u32(&mut archive, 0);
        push_u32(&mut archive, u32::MAX);
        push_u32(&mut archive, u32::MAX);
        push_u16(&mut archive, u16::try_from(name.len()).unwrap());
        push_u16(&mut archive, u16::try_from(local_extra.len()).unwrap());
        archive.extend_from_slice(name);
        archive.extend_from_slice(&local_extra);
        archive.extend_from_slice(payload);
        if signature {
            push_u32(&mut archive, 0x0807_4b50);
        }
        push_u32(&mut archive, crc);
        archive.extend_from_slice(&size.to_le_bytes());
        archive.extend_from_slice(&size.to_le_bytes());

        let central_directory_offset = u32::try_from(archive.len()).unwrap();
        let mut central = Vec::new();
        push_u32(&mut central, 0x0201_4b50);
        push_u16(&mut central, 45);
        push_u16(&mut central, 45);
        push_u16(&mut central, 0x08);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, crc);
        push_u32(&mut central, central_compressed_size);
        push_u32(&mut central, central_uncompressed_size);
        push_u16(&mut central, u16::try_from(name.len()).unwrap());
        push_u16(&mut central, u16::try_from(central_extra.len()).unwrap());
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, 0);
        push_u32(&mut central, 0);
        central.extend_from_slice(name);
        central.extend_from_slice(central_extra);
        let central_size = u32::try_from(central.len()).unwrap();
        archive.extend_from_slice(&central);

        push_u32(&mut archive, 0x0605_4b50);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 0);
        push_u16(&mut archive, 1);
        push_u16(&mut archive, 1);
        push_u32(&mut archive, central_size);
        push_u32(&mut archive, central_directory_offset);
        push_u16(&mut archive, 0);
        archive
    }

    fn zip64_descriptor_start(archive: &[u8], payload_len: usize) -> usize {
        let local = local_header_offset_for_name(archive, b"zip64.bin");
        let name_len = usize::from(u16::from_le_bytes(
            archive[local + 26..local + 28].try_into().unwrap(),
        ));
        let extra_len = usize::from(u16::from_le_bytes(
            archive[local + 28..local + 30].try_into().unwrap(),
        ));
        local + 30 + name_len + extra_len + payload_len
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

    fn rewrite_uncompressed_size(archive: &mut [u8], wanted_name: &[u8], size: u32) {
        const CENTRAL_HEADER: [u8; 4] = 0x0201_4b50_u32.to_le_bytes();
        const LOCAL_HEADER: [u8; 4] = 0x0403_4b50_u32.to_le_bytes();

        let local_offset = archive
            .windows(4)
            .enumerate()
            .find_map(|(offset, signature)| {
                if signature != LOCAL_HEADER || offset.saturating_add(30) > archive.len() {
                    return None;
                }
                let name_len =
                    u16::from_le_bytes([archive[offset + 26], archive[offset + 27]]) as usize;
                let name_end = offset.saturating_add(30).saturating_add(name_len);
                (name_end <= archive.len() && &archive[offset + 30..name_end] == wanted_name)
                    .then_some(offset)
            })
            .expect("local member header");
        archive[local_offset + 22..local_offset + 26].copy_from_slice(&size.to_le_bytes());

        let central_offset = archive
            .windows(4)
            .enumerate()
            .find_map(|(offset, signature)| {
                if signature != CENTRAL_HEADER || offset.saturating_add(46) > archive.len() {
                    return None;
                }
                let name_len =
                    u16::from_le_bytes([archive[offset + 28], archive[offset + 29]]) as usize;
                let name_end = offset.saturating_add(46).saturating_add(name_len);
                (name_end <= archive.len() && &archive[offset + 46..name_end] == wanted_name)
                    .then_some(offset)
            })
            .expect("central member header");
        archive[central_offset + 24..central_offset + 28].copy_from_slice(&size.to_le_bytes());
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

    #[test]
    fn callback_scoped_reader_streams_store_and_deflate_with_fixed_reads() {
        let payload = vec![b'x'; VERIFIED_ENTRY_READER_BUFFER_SIZE * 2 + 37];
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", &payload).unwrap();
        writer.write_deflated("deflated.bin", &payload).unwrap();
        let indexed = indexed_archive(writer.finish_to_bytes().unwrap());

        for name in ["stored.bin", "deflated.bin"] {
            let entry_id = indexed.entry_id(name).unwrap();
            let mut accounting = ZipOperationAccounting::default();
            let decoded = indexed
                .with_verified_entry_reader_with_accounting(
                    entry_id,
                    |reader| {
                        let mut decoded = Vec::new();
                        reader.read_to_end(&mut decoded).map(|_| decoded)
                    },
                    &mut accounting,
                )
                .unwrap();
            assert_eq!(decoded, payload);
            if name == "stored.bin" {
                assert_eq!(accounting.stored_payload_bytes_read(), payload.len() as u64);
                assert_eq!(
                    accounting.stored_payload_bytes_accepted(),
                    payload.len() as u64
                );
            } else {
                assert_eq!(accounting.deflate_bytes_produced(), payload.len() as u64);
                assert_eq!(accounting.deflate_bytes_accepted(), payload.len() as u64);
                assert!(accounting.compressed_deflate_payload_bytes_read() > 0);
            }
        }
    }

    #[test]
    fn verified_reader_bound_matches_private_fixed_layouts_and_dispatches_both_methods() {
        let payload = b"fixed verified-reader bound evidence";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("stored.bin", payload).unwrap();
        writer.write_deflated("deflated.bin", payload).unwrap();
        let indexed = indexed_archive(writer.finish_to_bytes().unwrap());
        let stored_id = indexed.entry_id("stored.bin").unwrap();
        let deflated_id = indexed.entry_id("deflated.bin").unwrap();

        let fixed_reader_and_source = size_of::<VerifiedEntryBufReader<'static, &'static mut ()>>()
            .checked_add(size_of::<CountingReader<ZipReader<&'static ()>>>())
            .unwrap();
        let stored_bound = indexed
            .verified_entry_reader_memory_upper_bound(stored_id)
            .unwrap();
        assert_eq!(stored_bound, fixed_reader_and_source as u64);

        let deflated_fixed = fixed_reader_and_source
            .checked_add(FLATE2_DEFLATE_INPUT_BUFFER_SIZE)
            .and_then(|size| size.checked_add(size_of::<DeflateDecoder<&'static mut ()>>()))
            .and_then(|size| size.checked_add(ZLIB_RS_INFLATE_STATE_UPPER_BOUND_BYTES))
            .unwrap();
        let deflated_bound = indexed
            .verified_entry_reader_memory_upper_bound(deflated_id)
            .unwrap();
        assert_eq!(deflated_bound, deflated_fixed as u64);
        assert!(deflated_bound > stored_bound);

        for entry_id in [stored_id, deflated_id] {
            indexed
                .with_verified_entry_reader(entry_id, |reader| {
                    let mut decoded = Vec::new();
                    reader.read_to_end(&mut decoded).map(|_| decoded)
                })
                .unwrap();
        }
    }

    #[test]
    fn callback_scoped_reader_drains_after_prefix_and_callback_error() {
        let payload = b"callback prefix and drain payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("payload.bin", payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let indexed = indexed_archive(bytes.clone());
        let entry_id = indexed.entry_id("payload.bin").unwrap();
        let mut accounting = ZipOperationAccounting::default();
        let prefix = indexed
            .with_verified_entry_reader_with_accounting(
                entry_id,
                |reader| {
                    let mut prefix = [0; 7];
                    reader.read_exact(&mut prefix)?;
                    Ok::<_, io::Error>(prefix.to_vec())
                },
                &mut accounting,
            )
            .unwrap();
        assert_eq!(prefix, &payload[..7]);
        assert_eq!(accounting.deflate_bytes_produced(), payload.len() as u64);
        assert_eq!(accounting.deflate_bytes_accepted(), 7);

        let indexed = indexed_archive(bytes);
        let entry_id = indexed.entry_id("payload.bin").unwrap();
        let callback_error = io::Error::other("callback stopped");
        let error = indexed
            .with_verified_entry_reader(entry_id, |_reader| {
                Err::<(), _>(io::Error::other(callback_error.to_string()))
            })
            .unwrap_err();
        assert!(error.archive().is_none());
        assert!(error.transport().is_none());
        assert!(matches!(error, VerifiedEntryReaderError::Callback(_)));
        assert!(error.callback().is_some());
    }

    #[test]
    fn callback_scoped_abortable_reader_stops_short_source_after_callback_error() {
        let payload = vec![b'a'; 256 * 1024 + 11];
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("payload.bin", &payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let length = u64::try_from(bytes.len()).unwrap();
        let requests = Arc::new(Mutex::new(Vec::new()));
        let indexed = IndexedArchive::from_reader_with_limits(
            InstrumentedChunkedReaderAt {
                bytes,
                max_chunk: 7,
                requests: Arc::clone(&requests),
            },
            length,
            ArchiveLimits::default(),
        )
        .unwrap();
        let entry_id = indexed.entry_id("payload.bin").unwrap();

        requests.lock().unwrap().clear();
        let decoded = indexed
            .with_verified_entry_reader(entry_id, |reader| {
                let mut decoded = Vec::new();
                reader.read_to_end(&mut decoded)?;
                Ok::<_, io::Error>(decoded)
            })
            .unwrap();
        assert_eq!(decoded, payload);
        let full_read: usize = requests
            .lock()
            .unwrap()
            .iter()
            .map(|(_, actual)| *actual)
            .sum();
        assert!(full_read >= payload.len());

        requests.lock().unwrap().clear();
        let error = indexed
            .with_verified_entry_reader_abortable(entry_id, |reader| {
                let mut prefix = [0_u8; 1];
                reader.read_exact(&mut prefix)?;
                Err::<(), _>(io::Error::other("callback stopped"))
            })
            .unwrap_err();
        assert!(matches!(error, VerifiedEntryReaderError::Callback(_)));
        let abort_read: usize = requests
            .lock()
            .unwrap()
            .iter()
            .map(|(_, actual)| *actual)
            .sum();

        requests.lock().unwrap().clear();
        let error = indexed
            .with_verified_entry_reader(entry_id, |reader| {
                let mut prefix = [0_u8; 1];
                reader.read_exact(&mut prefix)?;
                Err::<(), _>(io::Error::other("callback stopped"))
            })
            .unwrap_err();
        assert!(matches!(error, VerifiedEntryReaderError::Callback(_)));
        let drained_read: usize = requests
            .lock()
            .unwrap()
            .iter()
            .map(|(_, actual)| *actual)
            .sum();
        assert!(drained_read > abort_read);
        assert!(drained_read >= payload.len());
    }

    #[test]
    fn callback_scoped_reader_retains_callback_error_on_archive_failure() {
        let payload = b"callback secondary error payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated_sized("payload.bin", payload).unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let local = local_header_offset_for_name(&bytes, b"payload.bin");
        let central = central_header_offset_for_name(&bytes, b"payload.bin");
        bytes[local + 14..local + 18].copy_from_slice(&0_u32.to_le_bytes());
        bytes[central + 16..central + 20].copy_from_slice(&0_u32.to_le_bytes());
        let indexed = indexed_archive(bytes);
        let entry_id = indexed.entry_id("payload.bin").unwrap();
        let error = indexed
            .with_verified_entry_reader(entry_id, |_reader| {
                Err::<(), _>(io::Error::other("callback secondary"))
            })
            .unwrap_err();
        assert!(
            matches!(error.archive(), Some(error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. }))
        );
        assert!(error.callback().is_some());
    }

    #[test]
    fn callback_scoped_reader_drains_descriptor_failures_but_abortable_reader_stops() {
        let payload = b"descriptor callback payload";
        for signature in [false, true] {
            let mut bytes = stored_descriptor_fixture(payload, signature);
            let payload_start = 30 + b"stored.bin".len();
            bytes[payload_start] ^= 0x80;

            let indexed = indexed_archive(bytes.clone());
            let entry_id = indexed.entry_id("stored.bin").unwrap();
            let error = indexed
                .with_verified_entry_reader(entry_id, |_reader| {
                    Err::<(), _>(io::Error::other("callback stopped"))
                })
                .unwrap_err();
            assert!(matches!(
                error.archive(),
                Some(error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. })
            ));
            assert!(error.callback().is_some());

            let indexed = indexed_archive(bytes);
            let entry_id = indexed.entry_id("stored.bin").unwrap();
            let error = indexed
                .with_verified_entry_reader_abortable(entry_id, |_reader| {
                    Err::<(), _>(io::Error::other("callback stopped"))
                })
                .unwrap_err();
            assert!(matches!(error, VerifiedEntryReaderError::Callback(_)));

            let mut bytes = stored_descriptor_fixture(payload, signature);
            let descriptor = descriptor_start(payload, signature);
            bytes[descriptor..descriptor + 4].copy_from_slice(&0_u32.to_le_bytes());
            for abortable in [false, true] {
                let indexed = indexed_archive(bytes.clone());
                let entry_id = indexed.entry_id("stored.bin").unwrap();
                let mut called = false;
                let error = if abortable {
                    indexed
                        .with_verified_entry_reader_abortable(entry_id, |_reader| {
                            called = true;
                            Err::<(), _>(io::Error::other("callback stopped"))
                        })
                        .unwrap_err()
                } else {
                    indexed
                        .with_verified_entry_reader(entry_id, |_reader| {
                            called = true;
                            Err::<(), _>(io::Error::other("callback stopped"))
                        })
                        .unwrap_err()
                };
                assert!(!called);
                assert!(matches!(
                    error.archive(),
                    Some(error) if matches!(error.kind(), ErrorKind::InvalidChecksum { .. })
                ));
                assert!(error.callback().is_none());
            }
        }
    }

    #[test]
    fn callback_scoped_reader_defers_invalid_consume_and_retries_interrupts() {
        let payload = b"deferred consume payload";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("payload.bin", payload).unwrap();
        let indexed = indexed_archive(writer.finish_to_bytes().unwrap());
        let entry_id = indexed.entry_id("payload.bin").unwrap();
        let error = indexed
            .with_verified_entry_reader(entry_id, |reader| {
                let _ = reader.fill_buf()?;
                reader.consume(usize::MAX);
                Ok::<_, io::Error>(())
            })
            .unwrap_err();
        assert!(matches!(
            error.archive(),
            Some(error) if matches!(error.kind(), ErrorKind::InvalidInput { .. })
        ));

        let mut source =
            ScriptedReader::new([ReadStep::Interrupted, ReadStep::Bytes(payload.to_vec())]);
        let mut reader = VerifiedEntryBufReader::new(
            &mut source,
            ZipVerification {
                crc: crate::crc32(payload),
                uncompressed_size: payload.len() as u64,
            },
            None,
            AccountingReadKind::Stored,
        );
        let mut output = Vec::new();
        reader.read_to_end(&mut output).unwrap();
        reader.finish().unwrap();
        assert_eq!(output, payload);
    }

    #[test]
    fn precompressed_capture_verifies_store_and_deflate_exact_payloads() {
        let payload = b"verified precompressed payload with repeated repeated bytes";
        for deflated in [false, true] {
            let mut writer = StreamingArchiveWriter::new();
            if deflated {
                writer.write_deflated("payload.bin", payload).unwrap();
            } else {
                writer.write_stored("payload.bin", payload).unwrap();
            }
            let bytes = writer.finish_to_bytes().unwrap();
            let raw = raw_payload_for_single_entry(&bytes);
            let archive = indexed_archive(bytes);
            let entry_id = archive.entry_id("payload.bin").unwrap();
            let mut progress_chunks = 0usize;
            let token = archive
                .read_entry_precompressed_with_progress(entry_id, payload, |progress| {
                    progress_chunks += 1;
                    assert!(matches!(
                        progress,
                        PrecompressedProgress::Compressed { .. }
                            | PrecompressedProgress::Decoded { .. }
                    ));
                    Ok::<(), io::Error>(())
                })
                .unwrap();

            assert!(progress_chunks > 0);
            assert_eq!(
                token.compression_method(),
                if deflated {
                    CompressionMethod::Deflate
                } else {
                    CompressionMethod::Store
                }
            );
            assert_eq!(token.compressed.as_slice(), raw.as_slice());
            assert_eq!(token.compressed_size(), raw.len() as u64);
            assert_eq!(token.uncompressed_size(), payload.len() as u64);
            assert_eq!(token.crc32(), crate::crc32(payload));
        }
    }

    #[test]
    fn precompressed_capture_records_actual_crc_for_compatibility_zero_crc() {
        let payload = b"zero CRC remains compatibility-readable";
        let bytes = skipped_crc_fixture(payload);
        let archive = indexed_archive(bytes);
        for name in ["stored.bin", "deflated.bin"] {
            let entry_id = archive.entry_id(name).unwrap();
            let token = archive
                .read_entry_precompressed_with_progress(entry_id, payload, |_| {
                    Ok::<(), io::Error>(())
                })
                .unwrap();
            assert_eq!(token.crc32(), crate::crc32(payload));
        }
    }

    #[test]
    fn precompressed_capture_compares_expected_bytes_and_retains_callback_errors() {
        let payload = b"expected logical bytes are part of the proof";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("payload.bin", payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let archive = indexed_archive(bytes);
        let entry_id = archive.entry_id("payload.bin").unwrap();
        let mut wrong_payload = payload.to_vec();
        wrong_payload[0] ^= 0x01;

        let error = archive
            .read_entry_precompressed_with_progress(entry_id, &wrong_payload, |_| {
                Ok::<(), io::Error>(())
            })
            .unwrap_err();
        assert!(matches!(
            error.archive(),
            Some(error) if matches!(error.kind(), ErrorKind::InvalidInput { .. })
        ));

        let error = archive
            .read_entry_precompressed_with_progress(entry_id, payload, |_| {
                Err::<(), _>(io::Error::other("execution cancelled"))
            })
            .unwrap_err();
        assert!(matches!(error, VerifiedPrecompressedError::Callback(_)));
        assert!(error.callback().is_some());

        let error = archive
            .read_entry_precompressed_with_progress(entry_id, payload, |progress| {
                if matches!(progress, PrecompressedProgress::Decoded { .. }) {
                    Err(io::Error::other("decode execution cancelled"))
                } else {
                    Ok(())
                }
            })
            .unwrap_err();
        assert!(matches!(error, VerifiedPrecompressedError::Callback(_)));
    }

    #[test]
    fn precompressed_capture_rejects_corrupt_compressed_payload() {
        let payload = b"compressed bytes must remain an exact verified source";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("payload.bin", payload).unwrap();
        let mut bytes = writer.finish_to_bytes().unwrap();
        let (start, end) = {
            let source = ZipArchive::from_slice(&bytes).unwrap();
            let record = source.entries().next_entry().unwrap().unwrap();
            let entry = source.get_entry(record.wayfinder()).unwrap();
            entry.compressed_data_range()
        };
        assert!(end > start);
        bytes[usize::try_from(start).unwrap()] ^= 0x80;

        let archive = indexed_archive(bytes);
        let entry_id = archive.entry_id("payload.bin").unwrap();
        let error = archive
            .read_entry_precompressed_with_progress(entry_id, payload, |_| Ok::<(), io::Error>(()))
            .unwrap_err();
        assert!(matches!(error, VerifiedPrecompressedError::Archive(_)));
    }

    #[test]
    fn precompressed_capture_rejects_deflate_without_stream_end() {
        let payload = b"a truncated Deflate stream cannot publish its expected prefix";
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("payload.bin", payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let compressed = raw_payload_for_single_entry(&bytes);
        assert!(compressed.len() > 1);

        let error = verify_captured_precompressed_payload(
            CompressionMethod::Deflate,
            &compressed[..compressed.len() - 1],
            &mut CapturedDecodedOutput::Compare(payload),
            ZipVerification {
                crc: crate::crc32(payload),
                uncompressed_size: payload.len() as u64,
            },
            &mut |_| Ok::<(), io::Error>(()),
        )
        .unwrap_err();
        assert!(matches!(error, VerifiedPrecompressedError::Archive(_)));
    }

    #[test]
    fn precompressed_capture_handles_short_reader_at_chunks() {
        let payload = vec![b'z'; VERIFIED_ENTRY_READER_BUFFER_SIZE * 128 + 19];
        let mut writer = StreamingArchiveWriter::new();
        writer.write_deflated("payload.bin", &payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let source_len = bytes.len() as u64;
        let archive = IndexedArchive::from_reader_with_limits(
            ChunkedReaderAt {
                bytes,
                max_chunk: 3,
            },
            source_len,
            ArchiveLimits::default(),
        )
        .unwrap();
        let entry_id = archive.entry_id("payload.bin").unwrap();
        let token = archive
            .read_entry_precompressed_with_progress(entry_id, &payload, |_| Ok::<(), io::Error>(()))
            .unwrap();
        assert_eq!(token.uncompressed_size(), payload.len() as u64);
        assert_eq!(token.crc32(), crate::crc32(&payload));
    }

    #[test]
    fn precompressed_capture_bounds_source_requests_and_cancels_after_short_read() {
        let payload = (0..(PRECOMPRESSED_CAPTURE_BUFFER_SIZE * 2 + 37))
            .map(|index| {
                let index = u8::try_from(index % 251).expect("payload pattern fits in u8");
                index.wrapping_mul(37).wrapping_add(11)
            })
            .collect::<Vec<_>>();
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("payload.bin", &payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let source_len = bytes.len() as u64;
        let requests = Arc::new(Mutex::new(Vec::new()));
        let archive = IndexedArchive::from_reader_with_limits(
            InstrumentedChunkedReaderAt {
                bytes,
                max_chunk: 4096,
                requests: Arc::clone(&requests),
            },
            source_len,
            ArchiveLimits::default(),
        )
        .unwrap();
        let entry_id = archive.entry_id("payload.bin").unwrap();

        requests.lock().unwrap().clear();
        let token = archive
            .read_entry_precompressed_with_progress(entry_id, &payload, |_| Ok::<(), io::Error>(()))
            .unwrap();
        assert_eq!(token.compressed_size(), payload.len() as u64);
        let successful_requests = requests.lock().unwrap().clone();
        assert!(!successful_requests.is_empty());
        assert!(
            successful_requests
                .iter()
                .all(
                    |(requested, returned)| *requested <= PRECOMPRESSED_CAPTURE_BUFFER_SIZE
                        && *returned <= *requested
                )
        );
        assert!(
            successful_requests
                .iter()
                .any(
                    |(requested, returned)| *requested == PRECOMPRESSED_CAPTURE_BUFFER_SIZE
                        && *returned < *requested
                )
        );

        requests.lock().unwrap().clear();
        let error = archive
            .read_entry_precompressed_with_progress(entry_id, &payload, |progress| {
                if matches!(progress, PrecompressedProgress::Compressed { .. }) {
                    Err(io::Error::other("transfer cancelled"))
                } else {
                    Ok(())
                }
            })
            .unwrap_err();
        assert!(matches!(error, VerifiedPrecompressedError::Callback(_)));
        let cancelled_requests = requests.lock().unwrap().clone();
        assert_eq!(cancelled_requests.len(), 1);
        assert_eq!(cancelled_requests[0].0, PRECOMPRESSED_CAPTURE_BUFFER_SIZE);
        assert!(cancelled_requests[0].1 < cancelled_requests[0].0);
    }

    #[test]
    fn precompressed_capture_handles_zip64_archive_member() {
        let payload = b"This small file is in ZIP64 format.\n";
        let bytes = include_bytes!("../assets/zip64.zip").to_vec();
        let archive = indexed_archive_result(bytes, ArchiveLimits::UNBOUNDED).unwrap();
        let entry_id = archive.entry_id("README").unwrap();
        let token = archive
            .read_entry_precompressed_with_progress(entry_id, payload, |_| Ok::<(), io::Error>(()))
            .unwrap();
        assert_eq!(token.uncompressed_size(), payload.len() as u64);
        assert_eq!(token.crc32(), crate::crc32(payload));
    }

    #[test]
    fn fused_precompressed_capture_removes_a_cold_source_pass() {
        for deflated in [false, true] {
            for size in [0, 1, 65536, 262181, 1048613] {
                let mut state = 0x5a17_932d_u32;
                let payload = (0..size)
                    .map(|_| {
                        state ^= state << 13;
                        state ^= state >> 17;
                        state ^= state << 5;
                        state.to_le_bytes()[0]
                    })
                    .collect::<Vec<_>>();
                let mut writer = StreamingArchiveWriter::new();
                if deflated {
                    writer.write_deflated("payload.bin", &payload).unwrap();
                } else {
                    writer.write_stored("payload.bin", &payload).unwrap();
                }
                let bytes = writer.finish_to_bytes().unwrap();
                let source_len = bytes.len() as u64;
                let fused_input = bytes.clone();
                let compressed = raw_payload_for_single_entry(&bytes);
                let requests = Arc::new(Mutex::new(Vec::new()));
                let archive = IndexedArchive::from_reader_with_limits(
                    InstrumentedChunkedReaderAt {
                        bytes,
                        max_chunk: 65536,
                        requests: Arc::clone(&requests),
                    },
                    source_len,
                    ArchiveLimits::default(),
                )
                .unwrap();
                let id = archive.entry_id("payload.bin").unwrap();
                requests.lock().unwrap().clear();
                let decoded = archive.read_entry(id).unwrap();
                let control = archive
                    .read_entry_precompressed_with_progress(id, &decoded, |_| {
                        Ok::<(), io::Error>(())
                    })
                    .unwrap();
                let control_reads = requests.lock().unwrap().clone();
                drop(archive);
                // Each alternative starts with an independently indexed archive;
                // the control cannot warm the candidate's strict-layout cache.
                let archive = IndexedArchive::from_reader_with_limits(
                    InstrumentedChunkedReaderAt {
                        bytes: fused_input,
                        max_chunk: 65536,
                        requests: Arc::clone(&requests),
                    },
                    source_len,
                    ArchiveLimits::default(),
                )
                .unwrap();
                let id = archive.entry_id("payload.bin").unwrap();
                requests.lock().unwrap().clear();
                let mut captured = 0;
                let mut produced = 0;
                let (fused, decoded) = archive
                    .read_entry_precompressed_and_decoded_with_progress(id, |p| {
                        match p {
                            PrecompressedProgress::Compressed { bytes } => {
                                assert!(bytes >= captured);
                                captured = bytes;
                            },
                            PrecompressedProgress::Decoded { bytes } => {
                                assert!(bytes >= produced);
                                produced = bytes;
                            },
                        }
                        Ok::<(), io::Error>(())
                    })
                    .unwrap();
                let fused_reads = requests.lock().unwrap().clone();
                let control_bytes: usize = control_reads.iter().map(|(_, n)| n).sum();
                let fused_bytes: usize = fused_reads.iter().map(|(_, n)| n).sum();
                assert_eq!(decoded, payload);
                assert_eq!(fused.compressed_payload(), compressed);
                assert_eq!(fused.compressed_payload(), control.compressed_payload());
                assert_eq!(
                    (
                        fused.compression_method(),
                        fused.crc32(),
                        fused.uncompressed_size()
                    ),
                    (
                        control.compression_method(),
                        control.crc32(),
                        control.uncompressed_size()
                    )
                );
                assert_eq!(captured, compressed.len() as u64);
                assert_eq!(produced, payload.len() as u64);
                assert!(control_bytes >= fused_bytes + compressed.len());
                assert!(fused_reads.iter().all(|(requested, returned)| *requested
                    <= PRECOMPRESSED_CAPTURE_BUFFER_SIZE
                    && returned <= requested));
                println!(
                    "FUSED_CAPTURE_IO {{\"deflated\":{deflated},\"decoded_bytes\":{size},\"compressed_bytes\":{},\"control_calls\":{},\"control_bytes\":{control_bytes},\"fused_calls\":{},\"fused_bytes\":{fused_bytes}}}",
                    compressed.len(),
                    control_reads.len(),
                    fused_reads.len()
                );
                drop(archive);
                let mut destination = StreamingArchiveWriter::new();
                destination
                    .write_stored("untouched.bin", b"original bytes")
                    .unwrap();
                let destination = destination.finish_to_bytes().unwrap();
                let destination = ZipArchive::from_slice(destination.as_slice())
                    .unwrap()
                    .into_zip_archive();
                let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
                let index = PreservationIndex::new(&destination, &mut scratch).unwrap();
                let mut plan = crate::PreservationPlan::copy_all(&index);
                plan.try_append(crate::RegeneratedEntry::new_precompressed_shared(
                    "copied.bin",
                    fused,
                ))
                .unwrap();
                let mut accounting = ZipOperationAccounting::default();
                let output = index
                    .write_to_with_accounting(&plan, Vec::new(), &mut accounting)
                    .unwrap();
                let reopened = ArchiveReader::new(&output).unwrap();
                assert_eq!(reopened.read("copied.bin").unwrap(), payload);
                assert_eq!(reopened.read("untouched.bin").unwrap(), b"original bytes");
                assert_eq!(
                    accounting.precompressed_payload_bytes_emitted(),
                    if deflated { compressed.len() as u64 } else { 0 }
                );
                assert_eq!(
                    accounting.stored_payload_bytes_emitted(),
                    if deflated { 0 } else { payload.len() as u64 }
                );
                assert_eq!(accounting.generated_deflate_payload_bytes_emitted(), 0);
            }
        }
    }

    #[test]
    fn fused_precompressed_capture_handles_short_reads_and_cancellation() {
        let payload = vec![b'x'; 131091];
        for deflated in [false, true] {
            let mut writer = StreamingArchiveWriter::new();
            if deflated {
                writer.write_deflated("payload.bin", &payload).unwrap();
            } else {
                writer.write_stored("payload.bin", &payload).unwrap();
            }
            let bytes = writer.finish_to_bytes().unwrap();
            let len = bytes.len() as u64;
            let requests = Arc::new(Mutex::new(Vec::new()));
            let archive = IndexedArchive::from_reader_with_limits(
                InstrumentedChunkedReaderAt {
                    bytes,
                    max_chunk: 3,
                    requests: Arc::clone(&requests),
                },
                len,
                ArchiveLimits::default(),
            )
            .unwrap();
            let id = archive.entry_id("payload.bin").unwrap();
            let (token, decoded) = archive
                .read_entry_precompressed_and_decoded_with_progress(id, |_| {
                    Ok::<(), &'static str>(())
                })
                .unwrap();
            assert_eq!(decoded, payload);
            assert_eq!(token.crc32(), crate::crc32(&payload));
            for cancel_decode in [false, true] {
                let mut callbacks = 0;
                let result = archive.read_entry_precompressed_and_decoded_with_progress(id, |p| {
                    if matches!(p, PrecompressedProgress::Decoded { .. }) == cancel_decode {
                        callbacks += 1;
                        Err("cancelled at exact phase")
                    } else {
                        Ok(())
                    }
                });
                assert!(matches!(
                    result,
                    Err(VerifiedPrecompressedError::Callback(
                        "cancelled at exact phase"
                    ))
                ));
                assert_eq!(callbacks, 1);
            }
        }
    }

    #[test]
    fn fused_precompressed_capture_checks_crc_size_and_deflate_end() {
        let payload = vec![b'z'; 131091];
        for deflated in [false, true] {
            let mut writer = StreamingArchiveWriter::new();
            if deflated {
                writer.write_deflated("payload.bin", &payload).unwrap();
            } else {
                writer.write_stored("payload.bin", &payload).unwrap();
            }
            let archive_bytes = writer.finish_to_bytes().unwrap();
            let method = if deflated {
                CompressionMethod::Deflate
            } else {
                CompressionMethod::Store
            };
            let compressed = raw_payload_for_single_entry(&archive_bytes);
            for (declared, crc) in [
                (payload.len() as u64, crate::crc32(&payload) ^ 1),
                (payload.len() as u64 - 1, crate::crc32(&payload)),
                (payload.len() as u64 + 1, crate::crc32(&payload)),
            ] {
                let mut writer = StreamingArchiveWriter::new();
                writer
                    .archive
                    .write_precompressed_file("payload.bin", method, crc, declared, &compressed)
                    .unwrap();
                let archive = indexed_archive(writer.finish_to_bytes().unwrap());
                let id = archive.entry_id("payload.bin").unwrap();
                assert!(matches!(
                    archive.read_entry_precompressed_and_decoded_with_progress(id, |_| Ok::<
                        (),
                        io::Error,
                    >(
                        ()
                    )),
                    Err(VerifiedPrecompressedError::Archive(_))
                ));
            }
            if deflated {
                for bad in [
                    compressed[..compressed.len() - 1].to_vec(),
                    [compressed.as_slice(), &[0, 1, 2]].concat(),
                ] {
                    let mut writer = StreamingArchiveWriter::new();
                    writer
                        .archive
                        .write_precompressed_file(
                            "payload.bin",
                            method,
                            crate::crc32(&payload),
                            payload.len() as u64,
                            &bad,
                        )
                        .unwrap();
                    let archive = indexed_archive(writer.finish_to_bytes().unwrap());
                    assert!(
                        archive
                            .read_entry_precompressed_and_decoded_with_progress(
                                archive.entry_id("payload.bin").unwrap(),
                                |_| Ok::<(), io::Error>(())
                            )
                            .is_err()
                    );
                }
            }
        }
    }

    #[test]
    fn fused_precompressed_capture_preserves_zero_crc_zip64_and_entry_bounds() {
        let payload = b"zero CRC remains compatibility-readable";
        let archive = indexed_archive(skipped_crc_fixture(payload));
        for name in ["stored.bin", "deflated.bin"] {
            let (token, decoded) = archive
                .read_entry_precompressed_and_decoded_with_progress(
                    archive.entry_id(name).unwrap(),
                    |_| Ok::<(), io::Error>(()),
                )
                .unwrap();
            assert_eq!(decoded, payload);
            assert_eq!(token.crc32(), crate::crc32(payload));
        }
        assert!(
            archive
                .read_entry_precompressed_and_decoded_with_progress(EntryId(usize::MAX), |_| Ok::<
                    (),
                    io::Error,
                >(
                    ()
                ))
                .is_err()
        );
        let archive = indexed_archive_result(
            include_bytes!("../assets/zip64.zip").to_vec(),
            ArchiveLimits::UNBOUNDED,
        )
        .unwrap();
        let (token, decoded) = archive
            .read_entry_precompressed_and_decoded_with_progress(
                archive.entry_id("README").unwrap(),
                |_| Ok::<(), io::Error>(()),
            )
            .unwrap();
        assert_eq!(decoded, b"This small file is in ZIP64 format.\n");
        assert_eq!(token.crc32(), crate::crc32(&decoded));
    }

    #[test]
    fn fused_precompressed_capture_preserves_transport_and_allocation_failures() {
        let payload = vec![b'a'; 256 * 1024];
        let mut writer = StreamingArchiveWriter::new();
        writer.write_stored("payload.bin", &payload).unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let start = {
            let archive = ZipArchive::from_slice(&bytes).unwrap();
            let record = archive.entries().next_entry().unwrap().unwrap();
            archive
                .get_entry(record.wayfinder())
                .unwrap()
                .compressed_data_range()
                .0
        };
        let length = bytes.len() as u64;
        let archive = IndexedArchive::from_reader_with_limits(
            FailOnRangeReaderAt {
                bytes,
                fail_start: start + 32,
                fail_end: start + 33,
            },
            length,
            ArchiveLimits::default(),
        )
        .unwrap();
        let result = archive.read_entry_precompressed_and_decoded_with_progress(
            archive.entry_id("payload.bin").unwrap(),
            |_| Ok::<(), io::Error>(()),
        );
        assert!(matches!(
            result,
            Err(VerifiedPrecompressedError::Transport(_))
        ));

        // A ZIP64 declaration larger than Vec's capacity must fail before
        // capture, rather than panic or publish a partially decoded token.
        let mut writer = StreamingArchiveWriter::new();
        writer
            .archive
            .write_precompressed_file(
                "huge.bin",
                CompressionMethod::Deflate,
                1,
                usize::MAX as u64,
                &[3, 0],
            )
            .unwrap();
        let archive =
            indexed_archive_result(writer.finish_to_bytes().unwrap(), ArchiveLimits::UNBOUNDED)
                .unwrap();
        let mut callbacks = 0;
        let result = archive.read_entry_precompressed_and_decoded_with_progress(
            archive.entry_id("huge.bin").unwrap(),
            |_| {
                callbacks += 1;
                Ok::<(), io::Error>(())
            },
        );
        assert!(
            matches!(result, Err(VerifiedPrecompressedError::Archive(error)) if matches!(error.kind(), ErrorKind::Allocation { .. }))
        );
        assert_eq!(callbacks, 0);
    }

    #[derive(Debug)]
    struct ChunkedReaderAt {
        bytes: Vec<u8>,
        max_chunk: usize,
    }

    impl ReaderAt for ChunkedReaderAt {
        fn read_at(&self, output: &mut [u8], offset: u64) -> io::Result<usize> {
            let start = usize::try_from(offset).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "chunked source offset overflow",
                )
            })?;
            if start >= self.bytes.len() {
                return Ok(0);
            }
            let count = output
                .len()
                .min(self.max_chunk)
                .min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        }
    }

    #[derive(Debug)]
    struct InstrumentedChunkedReaderAt {
        bytes: Vec<u8>,
        max_chunk: usize,
        requests: Arc<Mutex<Vec<(usize, usize)>>>,
    }

    impl ReaderAt for InstrumentedChunkedReaderAt {
        fn read_at(&self, output: &mut [u8], offset: u64) -> io::Result<usize> {
            let start = usize::try_from(offset).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "instrumented source offset overflow",
                )
            })?;
            if start >= self.bytes.len() {
                self.requests.lock().unwrap().push((output.len(), 0));
                return Ok(0);
            }
            let count = output
                .len()
                .min(self.max_chunk)
                .min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            self.requests.lock().unwrap().push((output.len(), count));
            Ok(count)
        }
    }

    fn raw_payload_for_single_entry(bytes: &[u8]) -> Vec<u8> {
        let archive = ZipArchive::from_slice(bytes).unwrap();
        let record = archive.entries().next_entry().unwrap().unwrap();
        let entry = archive.get_entry(record.wayfinder()).unwrap();
        let (start, end) = entry.compressed_data_range();
        bytes[usize::try_from(start).unwrap()..usize::try_from(end).unwrap()].to_vec()
    }
}
