//! DOCX source-provider lifecycle benchmark.
//!
//! This benchmark measures the typed
//! source-backed DOCX owner (`litchi_docx::source_backed::Package`) through an
//! owned byte source, a recently-written `FileSource`, an instrumented byte
//! source, or the existing bounded `PptxRangeSource` adapter.  The adapter is
//! generic `ReadAt` machinery despite its historical PPTX name.
//!
//! Provider construction, file staging, source hashing, and all correctness
//! checks are outside the operation clock.  The clock contains package open,
//! document materialization, full-text extraction, and package/document drop.
//! The returned `String` is retained until after the clock and compared with
//! the deterministic oracle then; text destruction and oracle hashing are
//! outside the clock.  This is a provider baseline and does not claim equal
//! scope with the high-level filesystem facade lifecycle.

#![allow(clippy::module_name_repetitions)]

use std::{
    error::Error,
    ffi::OsString,
    fmt::Write as _,
    fs::{self, OpenOptions},
    io::{self, Cursor, Write},
    num::NonZeroU64,
    ops::Range,
    path::{Path, PathBuf},
    sync::{
        Arc, Mutex,
        atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering},
    },
    time::{Duration, Instant},
};

use litchi_core::{FileSource, OwnedSource, ReadAt, SourceVersion};
use litchi_docx::{Package as OwnedPackage, source_backed::Package as SourcePackage};
use litchi_opc::{ReadLimits, SourceCacheLimits};
use serde::Serialize;
use sha2::{Digest, Sha256};

use crate::docx_read_ahead::{ReadAheadReadAt, ReadAheadSnapshot};
use crate::pptx_range_source::{
    PPTX_RANGE_REQUEST_SIZE_BUCKETS, PptxRangeSource, PptxRangeSourceConfig,
    PptxRangeSourceSnapshot, TransferDelayPolicy, request_size_bucket,
};

const SCHEMA: &str = "docx_provider_lifecycle_v1";
const TRACE_SCHEMA: &str = "docx_provider_lifecycle_v2";
const TRACE_TIMING_SCOPE: &str = "Package::from_read_at_with_limits_and_cache_limits + document + extract_text + two cache diagnostic snapshots + package/document drop; returned text remains live after the clock; text destruction, hashing, oracle comparison, range counters and traces are outside";
const TRACE_SETUP_SCOPE: &str = "all v1 setup plus fresh bounded read-ahead buffer and wrapper construction outside the clock; fixed window capacity is reported separately; unmanaged benchmark pilot only";
const CORPUS_GENERATOR: &str = "litchi-docx-source-edit-media-v1";
const CORPUS_VERSION: &str = "0188-media-v1";
const CORPUS_ARCHIVE_SHA256: &str =
    "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4";
const CORPUS_ARCHIVE_BYTES: usize = 16_793_036;
const PARAGRAPH_COUNT: usize = 200;
const MEDIA_COUNT: usize = 8;
const MEDIA_BYTES: usize = 2 * 1024 * 1024;
const ARCHIVE_MEMBER_COUNT: usize = 20;
const MAX_RANGE_BYTES: usize = 1_048_576;
const MAX_DELAY_US: u64 = 100_000;
const MAX_TRANSFER_BYTES_PER_SECOND: u64 = 1_099_511_627_776;
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 1_000;
const CACHE_MAX_BYTES: usize = 8 * 1024 * 1024;
const CACHE_MAX_ENTRIES: usize = 128;
const MAX_TRACKED_RANGES: usize = 4_096;

const PROVIDER_SCOPE: &str = "typed-owner source-backed DOCX provider baseline; direct comparison with the high-level filesystem facade is out of scope";
const TIMING_SCOPE: &str = "Package::from_read_at_with_limits_and_cache_limits + document + extract_text + package/document drop; returned text remains live after the clock; text destruction, hashing, oracle comparison, counters, and diagnostics are outside";
const SETUP_SCOPE: &str = "corpus construction, source hashing, file staging, FileSource::open, instrumented wrapper construction, range-adapter construction, and limits construction are outside the clock";
const FILE_SCOPE: &str = "FileSource is opened before each iteration and the staged file is recently written; this is a warm-cache/recent-file provider observation, not a filesystem-cold result";
const READ_SCOPE: &str = "logical ReadAt calls observed by the named wrapper; counters are not physical filesystem or network I/O observations";
const MEDIA_SCOPE: &str = "compressed ZIP data ranges for word/media members only; a zero overlap proves only that the observed logical wrapper did not request or return those ranges";
const ALLOCATION_SCOPE: &str = "optional operation-scoped global-system-allocator region begins immediately before the operation clock and finishes immediately after it; normal binary runs report null and the allocator binary emits the existing Sample record";

static NEXT_STAGE_ID: AtomicUsize = AtomicUsize::new(0);

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ProviderKind {
    Bytes,
    File,
    InstrumentedBytes,
    Range,
}

impl ProviderKind {
    const fn name(self) -> &'static str {
        match self {
            Self::Bytes => "bytes",
            Self::File => "file",
            Self::InstrumentedBytes => "instrumented-bytes",
            Self::Range => "range",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Config {
    provider: ProviderKind,
    read_ahead_window_bytes: Option<usize>,
    trace_ranges: bool,
    max_range_bytes: Option<usize>,
    delay_us: Option<u64>,
    transfer_bytes_per_second: Option<NonZeroU64>,
    transfer_delay_policy: TransferDelayPolicy,
    samples: usize,
    warmup: usize,
    source_revision: String,
    output: PathBuf,
}

#[derive(Clone, Debug, Serialize)]
struct CorpusManifest {
    version: &'static str,
    generator: &'static str,
    shape: &'static str,
    paragraph_count: usize,
    media_member_count: usize,
    media_member_bytes: usize,
    archive_member_count: usize,
    archive_bytes: usize,
    archive_sha256: String,
    expected_text_bytes: usize,
    expected_text_sha256: String,
}

struct Corpus {
    bytes: Arc<Vec<u8>>,
    expected_text: String,
    media_ranges: Vec<Range<u64>>,
    manifest: CorpusManifest,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct SourceVersionRecord {
    id: u64,
    revision: u64,
}

impl From<SourceVersion> for SourceVersionRecord {
    fn from(version: SourceVersion) -> Self {
        Self {
            id: version.id(),
            revision: version.revision(),
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct LimitsRecord {
    max_input_bytes: u64,
    max_archive_members: usize,
    max_archive_member_name_bytes: u64,
    max_archive_metadata_bytes: u64,
    max_archive_compressed_bytes: u64,
    max_archive_entry_bytes: u64,
    max_archive_total_bytes: u64,
    max_parts: usize,
    max_part_bytes: u64,
    max_total_part_bytes: u64,
    max_content_types_bytes: usize,
    max_content_type_mappings: usize,
    max_relationship_parts: usize,
    max_relationship_xml_bytes: usize,
    max_total_relationship_xml_bytes: usize,
    max_relationships_per_part: usize,
    max_total_relationships: usize,
    max_relationship_graph_nodes: usize,
    max_xml_events: usize,
    max_total_relationship_xml_events: usize,
    max_xml_depth: usize,
    max_xml_attribute_bytes: usize,
    max_relationship_target_bytes: usize,
    cache_max_bytes: usize,
    cache_max_entries: usize,
}

impl LimitsRecord {
    fn from_limits(limits: ReadLimits, cache: SourceCacheLimits) -> Self {
        Self {
            max_input_bytes: limits.max_input_bytes(),
            max_archive_members: limits.max_archive_members(),
            max_archive_member_name_bytes: limits.max_archive_member_name_bytes(),
            max_archive_metadata_bytes: limits.max_archive_metadata_bytes(),
            max_archive_compressed_bytes: limits.max_archive_compressed_bytes(),
            max_archive_entry_bytes: limits.max_archive_entry_bytes(),
            max_archive_total_bytes: limits.max_archive_total_bytes(),
            max_parts: limits.max_parts(),
            max_part_bytes: limits.max_part_bytes(),
            max_total_part_bytes: limits.max_total_part_bytes(),
            max_content_types_bytes: limits.max_content_types_bytes(),
            max_content_type_mappings: limits.max_content_type_mappings(),
            max_relationship_parts: limits.max_relationship_parts(),
            max_relationship_xml_bytes: limits.max_relationship_xml_bytes(),
            max_total_relationship_xml_bytes: limits.max_total_relationship_xml_bytes(),
            max_relationships_per_part: limits.max_relationships_per_part(),
            max_total_relationships: limits.max_total_relationships(),
            max_relationship_graph_nodes: limits.max_relationship_graph_nodes(),
            max_xml_events: limits.max_xml_events(),
            max_total_relationship_xml_events: limits.max_total_relationship_xml_events(),
            max_xml_depth: limits.max_xml_depth(),
            max_xml_attribute_bytes: limits.max_xml_attribute_bytes(),
            max_relationship_target_bytes: limits.max_relationship_target_bytes(),
            cache_max_bytes: cache.max_bytes(),
            cache_max_entries: cache.max_entries(),
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ProviderConfigRecord {
    #[serde(skip_serializing_if = "Option::is_none")]
    read_ahead_window_bytes: Option<usize>,
    provider: &'static str,
    max_range_bytes: Option<usize>,
    delay_us: Option<u64>,
    transfer_bytes_per_second: Option<NonZeroU64>,
    transfer_delay_policy: TransferDelayPolicy,
    source_construction: &'static str,
    file_scope: &'static str,
    read_counter_scope: &'static str,
}

impl Config {
    fn record(&self) -> ProviderConfigRecord {
        ProviderConfigRecord {
            read_ahead_window_bytes: self.read_ahead_window_bytes,
            provider: self.provider.name(),
            max_range_bytes: self.max_range_bytes,
            delay_us: self.delay_us,
            transfer_bytes_per_second: self.transfer_bytes_per_second,
            transfer_delay_policy: self.transfer_delay_policy,
            source_construction: "all provider adapters are constructed outside the operation clock",
            file_scope: FILE_SCOPE,
            read_counter_scope: READ_SCOPE,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct CounterRecord {
    availability: &'static str,
    scope: &'static str,
    logical_calls: Option<u64>,
    requested_bytes: Option<u64>,
    returned_bytes: Option<u64>,
    min_request_bytes: Option<u64>,
    max_request_bytes: Option<u64>,
    short_reads: Option<u64>,
    delayed_calls: Option<u64>,
    transfer_paced_calls: Option<u64>,
    transfer_delay_ns: Option<u64>,
    request_size_counts: Option<[u64; PPTX_RANGE_REQUEST_SIZE_BUCKETS]>,
}

impl CounterRecord {
    fn unavailable() -> Self {
        Self {
            availability: "unavailable",
            scope: READ_SCOPE,
            logical_calls: None,
            requested_bytes: None,
            returned_bytes: None,
            min_request_bytes: None,
            max_request_bytes: None,
            short_reads: None,
            delayed_calls: None,
            transfer_paced_calls: None,
            transfer_delay_ns: None,
            request_size_counts: None,
        }
    }

    fn from_counted(snapshot: &CountingSnapshot) -> Self {
        Self {
            availability: "available",
            scope: snapshot.scope,
            logical_calls: Some(snapshot.logical_calls),
            requested_bytes: Some(snapshot.requested_bytes),
            returned_bytes: Some(snapshot.returned_bytes),
            min_request_bytes: snapshot.min_request_bytes,
            max_request_bytes: snapshot.max_request_bytes,
            short_reads: Some(snapshot.short_reads),
            delayed_calls: Some(0),
            transfer_paced_calls: Some(0),
            transfer_delay_ns: Some(0),
            request_size_counts: Some(snapshot.request_size_counts),
        }
    }

    fn from_range(snapshot: PptxRangeSourceSnapshot) -> Self {
        Self {
            availability: "available",
            scope: "PptxRangeSource logical adapter counters; generic ReadAt behavior, not physical I/O",
            logical_calls: Some(snapshot.logical_calls),
            requested_bytes: Some(snapshot.requested_bytes),
            returned_bytes: Some(snapshot.returned_bytes),
            min_request_bytes: snapshot.min_request_bytes,
            max_request_bytes: snapshot.max_request_bytes,
            short_reads: Some(snapshot.short_reads),
            delayed_calls: Some(snapshot.delayed_calls),
            transfer_paced_calls: Some(snapshot.transfer_paced_calls),
            transfer_delay_ns: Some(snapshot.transfer_delay_ns),
            request_size_counts: Some(snapshot.request_size_counts),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct MediaRangeProof {
    availability: &'static str,
    scope: &'static str,
    media_range_count: usize,
    observed_call_count: Option<u64>,
    requested_overlap_bytes: Option<u64>,
    returned_overlap_bytes: Option<u64>,
    status: &'static str,
    reason: Option<&'static str>,
}

impl MediaRangeProof {
    fn unavailable(reason: &'static str, media_range_count: usize) -> Self {
        Self {
            availability: "unavailable",
            scope: MEDIA_SCOPE,
            media_range_count,
            observed_call_count: None,
            requested_overlap_bytes: None,
            returned_overlap_bytes: None,
            status: "unavailable",
            reason: Some(reason),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct ReadEvidence {
    #[serde(skip_serializing_if = "Option::is_none")]
    read_ahead: Option<ReadAheadSnapshot>,
    #[serde(skip_serializing_if = "Option::is_none")]
    logical_wrapper: Option<CounterRecord>,
    #[serde(skip_serializing_if = "Option::is_none")]
    logical_ranges: Option<Vec<ObservedRange>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    physical_ranges: Option<Vec<ObservedRange>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    logical_media_range_proof: Option<MediaRangeProof>,
    wrapper: CounterRecord,
    range_adapter: Option<CounterRecord>,
    media_range_proof: MediaRangeProof,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct CacheRecord {
    open_successful_loads: u64,
    successful_loads: u64,
    failed_loads: u64,
    retained_bytes: usize,
    retained_entries: usize,
    budget_managed: bool,
}

#[derive(Clone, Debug, Serialize)]
struct SampleRecord {
    #[serde(skip_serializing_if = "Option::is_none")]
    cache: Option<CacheRecord>,
    sample_index: usize,
    latency_ns: u64,
    actual_text_verified: bool,
    actual_text_bytes: usize,
    actual_text_sha256: String,
    source_version_before: SourceVersionRecord,
    source_version_after: SourceVersionRecord,
    source_version_unchanged: bool,
    reads: ReadEvidence,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation: Option<crate::allocation_metrics::Sample>,
}

#[derive(Debug, Serialize)]
struct Report {
    #[serde(skip_serializing_if = "Option::is_none")]
    media_ranges: Option<Vec<Range<u64>>>,
    schema: &'static str,
    provider_scope: &'static str,
    timing_scope: &'static str,
    setup_scope: &'static str,
    allocation_scope: &'static str,
    corpus: CorpusManifest,
    source_bytes: usize,
    source_sha256: String,
    requested_source_revision: String,
    limits: LimitsRecord,
    provider: ProviderConfigRecord,
    warmup: usize,
    samples: usize,
    rows: Vec<SampleRecord>,
}

#[derive(Clone, Debug, Serialize)]
struct ObservedRange {
    offset: u64,
    requested: u64,
    returned: u64,
}

#[derive(Debug)]
struct CountingState {
    logical_calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    min_request_bytes: AtomicU64,
    max_request_bytes: AtomicU64,
    short_reads: AtomicU64,
    request_size_counts: [AtomicU64; PPTX_RANGE_REQUEST_SIZE_BUCKETS],
    failed: AtomicBool,
    ranges: Mutex<Vec<ObservedRange>>,
    ranges_overflowed: AtomicBool,
}

impl Default for CountingState {
    fn default() -> Self {
        Self {
            logical_calls: AtomicU64::new(0),
            requested_bytes: AtomicU64::new(0),
            returned_bytes: AtomicU64::new(0),
            min_request_bytes: AtomicU64::new(u64::MAX),
            max_request_bytes: AtomicU64::new(0),
            short_reads: AtomicU64::new(0),
            request_size_counts: std::array::from_fn(|_| AtomicU64::new(0)),
            failed: AtomicBool::new(false),
            ranges: Mutex::new(Vec::new()),
            ranges_overflowed: AtomicBool::new(false),
        }
    }
}

#[derive(Clone, Debug)]
struct CountingSnapshot {
    scope: &'static str,
    logical_calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    min_request_bytes: Option<u64>,
    max_request_bytes: Option<u64>,
    short_reads: u64,
    request_size_counts: [u64; PPTX_RANGE_REQUEST_SIZE_BUCKETS],
    ranges: Vec<ObservedRange>,
}

impl CountingState {
    fn add(counter: &AtomicU64, value: u64, failed: &AtomicBool) -> io::Result<()> {
        if counter
            .fetch_update(Ordering::AcqRel, Ordering::Acquire, |current| {
                current.checked_add(value)
            })
            .is_err()
        {
            failed.store(true, Ordering::Release);
            return Err(io::Error::other("instrumented source counter overflow"));
        }
        Ok(())
    }

    fn record(&self, offset: u64, requested: usize, returned: usize) -> io::Result<()> {
        let requested = u64::try_from(requested)
            .map_err(|_| io::Error::other("read request length does not fit u64"))?;
        let returned = u64::try_from(returned)
            .map_err(|_| io::Error::other("read result length does not fit u64"))?;
        Self::add(&self.logical_calls, 1, &self.failed)?;
        Self::add(&self.requested_bytes, requested, &self.failed)?;
        Self::add(&self.returned_bytes, returned, &self.failed)?;
        if returned < requested {
            Self::add(&self.short_reads, 1, &self.failed)?;
        }
        let bucket = request_size_bucket(requested);
        Self::add(&self.request_size_counts[bucket], 1, &self.failed)?;
        self.min_request_bytes
            .fetch_min(requested, Ordering::AcqRel);
        self.max_request_bytes
            .fetch_max(requested, Ordering::AcqRel);
        let end = offset
            .checked_add(requested)
            .ok_or_else(|| io::Error::other("observed read range overflows u64"))?;
        let _ = end;
        let mut ranges = self
            .ranges
            .lock()
            .map_err(|_| io::Error::other("instrumented source range lock poisoned"))?;
        if ranges.len() < MAX_TRACKED_RANGES {
            ranges.push(ObservedRange {
                offset,
                requested,
                returned,
            });
        } else {
            self.ranges_overflowed.store(true, Ordering::Release);
        }
        Ok(())
    }

    fn snapshot(&self) -> io::Result<CountingSnapshot> {
        if self.failed.load(Ordering::Acquire) {
            return Err(io::Error::other("instrumented source counters unavailable"));
        }
        let ranges = self
            .ranges
            .lock()
            .map_err(|_| io::Error::other("instrumented source range lock poisoned"))?
            .clone();
        if self.ranges_overflowed.load(Ordering::Acquire) {
            return Err(io::Error::other(
                "instrumented source range sample bound exceeded",
            ));
        }
        let logical_calls = self.logical_calls.load(Ordering::Acquire);
        let request_size_counts =
            std::array::from_fn(|index| self.request_size_counts[index].load(Ordering::Acquire));
        let histogram_calls = request_size_counts.iter().try_fold(0_u64, |total, count| {
            total
                .checked_add(*count)
                .ok_or_else(|| io::Error::other("instrumented histogram overflow"))
        })?;
        if histogram_calls != logical_calls || ranges.len() as u64 != logical_calls {
            return Err(io::Error::other(
                "instrumented source counter snapshot is inconsistent",
            ));
        }
        Ok(CountingSnapshot {
            scope: "caller-visible CountingReadAt calls around the provider source",
            logical_calls,
            requested_bytes: self.requested_bytes.load(Ordering::Acquire),
            returned_bytes: self.returned_bytes.load(Ordering::Acquire),
            min_request_bytes: (logical_calls != 0)
                .then(|| self.min_request_bytes.load(Ordering::Acquire)),
            max_request_bytes: (logical_calls != 0)
                .then(|| self.max_request_bytes.load(Ordering::Acquire)),
            short_reads: self.short_reads.load(Ordering::Acquire),
            request_size_counts,
            ranges,
        })
    }
}

struct CountingReadAt {
    inner: Arc<dyn ReadAt>,
    state: Arc<CountingState>,
}

impl CountingReadAt {
    fn new(inner: Arc<dyn ReadAt>, state: Arc<CountingState>) -> Self {
        Self { inner, state }
    }
}

impl ReadAt for CountingReadAt {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let returned = self.inner.read_at(offset, output)?;
        if returned > output.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "wrapped source returned more bytes than requested",
            ));
        }
        self.state.record(offset, output.len(), returned)?;
        Ok(returned)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

struct Provider {
    read_ahead: Option<Arc<ReadAheadReadAt>>,
    logical_counting: Option<Arc<CountingState>>,
    trace_ranges: bool,
    source: Arc<dyn ReadAt>,
    counting: Option<Arc<CountingState>>,
    range: Option<Arc<PptxRangeSource>>,
    version: SourceVersion,
}

fn provider(
    config: &Config,
    bytes: &Arc<Vec<u8>>,
    file_path: Option<&Path>,
) -> Result<Provider, Box<dyn Error>> {
    let (source, counting, range): (
        Arc<dyn ReadAt>,
        Option<Arc<CountingState>>,
        Option<Arc<PptxRangeSource>>,
    ) = match config.provider {
        ProviderKind::Bytes => (
            Arc::new(OwnedSource::from_arc(Arc::clone(bytes))),
            None,
            None,
        ),
        ProviderKind::File => {
            let path = file_path.ok_or("file provider requires a staged path")?;
            let state = Arc::new(CountingState::default());
            let inner: Arc<dyn ReadAt> = Arc::new(FileSource::open(path)?);
            let wrapped: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(inner, Arc::clone(&state)));
            (wrapped, Some(state), None)
        },
        ProviderKind::InstrumentedBytes => {
            let state = Arc::new(CountingState::default());
            let inner: Arc<dyn ReadAt> = Arc::new(OwnedSource::from_arc(Arc::clone(bytes)));
            let wrapped: Arc<dyn ReadAt> =
                Arc::new(CountingReadAt::new(Arc::clone(&inner), Arc::clone(&state)));
            (wrapped, Some(state), None)
        },
        ProviderKind::Range => {
            let state = Arc::new(CountingState::default());
            let inner: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(
                Arc::new(OwnedSource::from_arc(Arc::clone(bytes))),
                Arc::clone(&state),
            ));
            let mut range_config = PptxRangeSourceConfig::new(
                config.max_range_bytes,
                config.delay_us.map(Duration::from_micros),
            );
            range_config.transfer_bytes_per_second = config.transfer_bytes_per_second;
            range_config.transfer_delay_policy = config.transfer_delay_policy;
            let adapter = Arc::new(PptxRangeSource::new(inner, range_config));
            let source: Arc<dyn ReadAt> = adapter.clone();
            (source, Some(state), Some(adapter))
        },
    };
    let (source, read_ahead, logical_counting) =
        if let Some(window_bytes) = config.read_ahead_window_bytes {
            let adapter = Arc::new(ReadAheadReadAt::new(source, window_bytes)?);
            let logical = Arc::new(CountingState::default());
            let source: Arc<dyn ReadAt> =
                Arc::new(CountingReadAt::new(adapter.clone(), Arc::clone(&logical)));
            (source, Some(adapter), Some(logical))
        } else {
            (source, None, None)
        };
    let version = source.version()?;
    Ok(Provider {
        read_ahead,
        logical_counting,
        trace_ranges: config.trace_ranges,
        source,
        counting,
        range,
        version,
    })
}

fn overlap(a: u64, b: u64, range: &Range<u64>) -> u64 {
    let start = a.max(range.start);
    let end = b.min(range.end);
    end.saturating_sub(start)
}

fn media_proof(
    snapshot: Option<&CountingSnapshot>,
    media_ranges: &[Range<u64>],
) -> MediaRangeProof {
    let Some(snapshot) = snapshot else {
        return MediaRangeProof::unavailable(
            "provider has no offset instrumentation",
            media_ranges.len(),
        );
    };
    let mut requested = 0_u64;
    let mut returned = 0_u64;
    for observed in &snapshot.ranges {
        let requested_end = match observed.offset.checked_add(observed.requested) {
            Some(end) => end,
            None => {
                return MediaRangeProof::unavailable(
                    "an observed request range overflowed",
                    media_ranges.len(),
                );
            },
        };
        let returned_end = match observed.offset.checked_add(observed.returned) {
            Some(end) => end,
            None => {
                return MediaRangeProof::unavailable(
                    "an observed returned range overflowed",
                    media_ranges.len(),
                );
            },
        };
        for media in media_ranges {
            requested = requested.saturating_add(overlap(observed.offset, requested_end, media));
            returned = returned.saturating_add(overlap(observed.offset, returned_end, media));
        }
    }
    let status = if requested == 0 && returned == 0 {
        "proved_no_media_overlap"
    } else {
        "media_overlap_observed"
    };
    MediaRangeProof {
        availability: "available",
        scope: MEDIA_SCOPE,
        media_range_count: media_ranges.len(),
        observed_call_count: Some(snapshot.logical_calls),
        requested_overlap_bytes: Some(requested),
        returned_overlap_bytes: Some(returned),
        status,
        reason: None,
    }
}

fn evidence(
    provider: &Provider,
    media_ranges: &[Range<u64>],
) -> Result<ReadEvidence, Box<dyn Error>> {
    let counted = provider
        .counting
        .as_ref()
        .map(|state| state.snapshot())
        .transpose()?;
    let mut wrapper = counted
        .as_ref()
        .map(CounterRecord::from_counted)
        .unwrap_or_else(CounterRecord::unavailable);
    let range_adapter = provider
        .range
        .as_ref()
        .map(|adapter| adapter.snapshot().map(CounterRecord::from_range))
        .transpose()?;
    let logical = provider
        .logical_counting
        .as_ref()
        .map(|state| state.snapshot())
        .transpose()?;
    let logical = logical.as_ref().or(counted.as_ref());
    if provider.trace_ranges {
        wrapper.scope = "physical adapter calls below read-ahead; synthetic transport observations, not disk/network I/O";
    }
    Ok(ReadEvidence {
        read_ahead: provider
            .read_ahead
            .as_ref()
            .map(|adapter| adapter.snapshot())
            .transpose()?,
        logical_wrapper: provider.trace_ranges.then(|| {
            let mut counter = logical
                .map(CounterRecord::from_counted)
                .unwrap_or_else(CounterRecord::unavailable);
            counter.scope = "package logical ReadAt calls above read-ahead";
            counter
        }),
        logical_ranges: provider
            .trace_ranges
            .then(|| logical.map(|s| s.ranges.clone()).unwrap_or_default()),
        physical_ranges: provider.trace_ranges.then(|| {
            counted
                .as_ref()
                .map(|s| s.ranges.clone())
                .unwrap_or_default()
        }),
        logical_media_range_proof: provider
            .trace_ranges
            .then(|| media_proof(logical, media_ranges)),
        wrapper,
        range_adapter,
        media_range_proof: media_proof(counted.as_ref(), media_ranges),
    })
}

fn run_sample(
    config: &Config,
    corpus: &Corpus,
    file_path: Option<&Path>,
    sample_index: usize,
    limits: ReadLimits,
    cache: SourceCacheLimits,
) -> Result<SampleRecord, Box<dyn Error>> {
    let provider = provider(config, &corpus.bytes, file_path)?;
    let before = provider.version;
    let allocation_region = crate::allocation_metrics::begin();
    let started = Instant::now();
    let (actual_text, cache_record) = {
        let package = SourcePackage::from_read_at_with_limits_and_cache_limits(
            Arc::clone(&provider.source),
            limits,
            cache,
        )?;
        let open_cache = config.trace_ranges.then(|| package.cache_diagnostics());
        let document = package.document()?;
        let text = document.extract_text()?;
        let cache_record = open_cache.map(|open| {
            let after = package.cache_diagnostics();
            CacheRecord {
                open_successful_loads: open.successful_loads,
                successful_loads: after.successful_loads,
                failed_loads: after.failed_loads,
                retained_bytes: after.retained_bytes,
                retained_entries: after.retained_entries,
                budget_managed: after.budget_managed,
            }
        });
        std::hint::black_box(&text);
        // `package` and `document` both drop before this block ends, so their
        // ownership cost remains inside the lifecycle clock. `text` escapes.
        (text, cache_record)
    };
    let elapsed = started.elapsed();
    let allocation = allocation_region.finish();
    let latency_ns = u64::try_from(elapsed.as_nanos())
        .map_err(|_| "operation duration does not fit u64 nanoseconds")?;
    let after = provider.source.version()?;
    if before != after {
        return Err(format!(
            "provider source changed during sample {sample_index}: before={before:?}, after={after:?}"
        )
        .into());
    }
    if actual_text != corpus.expected_text {
        return Err(format!("full-text oracle mismatch in sample {sample_index}").into());
    }
    let reads = evidence(&provider, &corpus.media_ranges)?;
    Ok(SampleRecord {
        cache: cache_record,
        sample_index,
        latency_ns,
        actual_text_verified: true,
        actual_text_bytes: actual_text.len(),
        actual_text_sha256: sha256_hex(actual_text.as_bytes()),
        source_version_before: before.into(),
        source_version_after: after.into(),
        source_version_unchanged: true,
        reads,
        allocation,
    })
}

fn stage_file(bytes: &[u8]) -> Result<StagedFile, Box<dyn Error>> {
    let stage_id = NEXT_STAGE_ID.fetch_add(1, Ordering::Relaxed);
    let root = std::env::temp_dir().join(format!(
        "litchi-docx-provider-{}-{stage_id}",
        std::process::id()
    ));
    fs::create_dir(&root)?;
    let path = root.join("source.docx");
    let mut file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&path)?;
    file.write_all(bytes)?;
    file.sync_all()?;
    Ok(StagedFile { root, path })
}

struct StagedFile {
    root: PathBuf,
    path: PathBuf,
}

impl Drop for StagedFile {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.path);
        let _ = fs::remove_dir(&self.root);
    }
}

fn paragraph_text(index: usize) -> String {
    format!("litchi-perf-baseline-docx-semantic-v1-source-{index:05}")
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn corpus_ranges(bytes: &[u8]) -> Result<(Vec<Range<u64>>, usize), Box<dyn Error>> {
    let archive = soapberry_zip::ZipArchive::from_slice(bytes)?;
    let mut media = Vec::new();
    let mut members = 0;
    for header in archive.entries() {
        let header = header?;
        let name = header.file_path().try_normalize()?.as_ref().to_owned();
        let entry = archive.get_entry(header.wayfinder())?;
        let (start, end) = entry.compressed_data_range();
        members += 1;
        if name.starts_with("word/media/") {
            media.push(start..end);
        }
    }
    Ok((media, members))
}

fn build_corpus() -> Result<Corpus, Box<dyn Error>> {
    // Reuse the accepted 0188 generator from the harness root. It owns the
    // narrow compatibility restoration that removes the producer's admitted
    // section namespace spelling and reserializes the exact historical ZIP.
    let bytes = crate::docx_source_edit_bytes()?;
    let archive_sha256 = sha256_hex(&bytes);
    if bytes.len() != CORPUS_ARCHIVE_BYTES || archive_sha256 != CORPUS_ARCHIVE_SHA256 {
        return Err(format!(
            "pinned 0188 corpus changed: expected {CORPUS_ARCHIVE_BYTES} bytes / {CORPUS_ARCHIVE_SHA256}, got {} bytes / {archive_sha256}; resolve an explicit corpus version before measuring",
            bytes.len()
        )
        .into());
    }
    let (media_ranges, archive_member_count) = corpus_ranges(&bytes)?;
    if media_ranges.len() != MEDIA_COUNT || archive_member_count != ARCHIVE_MEMBER_COUNT {
        return Err(format!(
            "pinned 0188 corpus shape changed: expected {MEDIA_COUNT} media members and {ARCHIVE_MEMBER_COUNT} archive members, got {} and {archive_member_count}",
            media_ranges.len()
        )
        .into());
    }
    let expected_text = (0..PARAGRAPH_COUNT).map(paragraph_text).collect::<String>();
    let semantic_package = OwnedPackage::from_reader(Cursor::new(bytes.clone()))?;
    let parsed_text = semantic_package.document()?.text()?;
    if parsed_text != expected_text {
        return Err("pinned DOCX corpus semantic oracle differs from the package parser".into());
    }
    let manifest = CorpusManifest {
        version: CORPUS_VERSION,
        generator: CORPUS_GENERATOR,
        shape: "200 paragraphs + eight 2 MiB PNG-signature media members",
        paragraph_count: PARAGRAPH_COUNT,
        media_member_count: MEDIA_COUNT,
        media_member_bytes: MEDIA_BYTES,
        archive_member_count,
        archive_bytes: bytes.len(),
        archive_sha256,
        expected_text_bytes: expected_text.len(),
        expected_text_sha256: sha256_hex(expected_text.as_bytes()),
    };
    Ok(Corpus {
        bytes: Arc::new(bytes),
        expected_text,
        media_ranges,
        manifest,
    })
}

fn parse_usize(value: &str, flag: &str) -> Result<usize, Box<dyn Error>> {
    value
        .parse::<usize>()
        .map_err(|_| format!("{flag} requires a finite unsigned integer").into())
}

fn parse_u64(value: &str, flag: &str) -> Result<u64, Box<dyn Error>> {
    value
        .parse::<u64>()
        .map_err(|_| format!("{flag} requires a finite unsigned integer").into())
}

fn need_value(args: &[OsString], index: &mut usize, flag: &str) -> Result<String, Box<dyn Error>> {
    *index += 1;
    args.get(*index)
        .ok_or_else(|| format!("{flag} requires a value").into())
        .and_then(|value| {
            value
                .to_str()
                .map(str::to_owned)
                .ok_or_else(|| format!("{flag} value must be valid UTF-8").into())
        })
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut provider = None;
    let mut read_ahead_window_bytes = None;
    let mut trace_ranges = false;
    let mut max_range_bytes = None;
    let mut delay_us = None;
    let mut transfer_bytes_per_second = None;
    let mut transfer_delay_policy = TransferDelayPolicy::SeparateSleeps;
    let mut samples = None;
    let mut warmup = 0;
    let mut source_revision = None;
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index].to_str().ok_or("argument must be valid UTF-8")?;
        match flag {
            "--trace-ranges" => {
                if trace_ranges {
                    return Err("--trace-ranges was specified more than once".into());
                }
                trace_ranges = true;
            },
            "--read-ahead" => {
                if read_ahead_window_bytes.is_some() {
                    return Err("--read-ahead was specified more than once".into());
                }
                let value = parse_usize(&need_value(args, &mut index, flag)?, flag)?;
                if !(1..=65_536).contains(&value) {
                    return Err("--read-ahead must be in 1..=65536".into());
                }
                read_ahead_window_bytes = Some(value);
            },
            "--provider" => {
                if provider.is_some() {
                    return Err("--provider was specified more than once".into());
                }
                provider = Some(match need_value(args, &mut index, flag)?.as_str() {
                    "bytes" => ProviderKind::Bytes,
                    "file" => ProviderKind::File,
                    "instrumented-bytes" => ProviderKind::InstrumentedBytes,
                    "range" => ProviderKind::Range,
                    _ => {
                        return Err(
                            "--provider must be bytes, file, instrumented-bytes, or range".into(),
                        );
                    },
                });
            },
            "--max-range" => {
                if max_range_bytes.is_some() {
                    return Err("--max-range was specified more than once".into());
                }
                let value = parse_usize(&need_value(args, &mut index, flag)?, flag)?;
                if value == 0 || value > MAX_RANGE_BYTES {
                    return Err(format!("--max-range must be in 1..={MAX_RANGE_BYTES}").into());
                }
                max_range_bytes = Some(value);
            },
            "--delay-us" => {
                if delay_us.is_some() {
                    return Err("--delay-us was specified more than once".into());
                }
                let value = parse_u64(&need_value(args, &mut index, flag)?, flag)?;
                if value > MAX_DELAY_US {
                    return Err(format!("--delay-us must be <= {MAX_DELAY_US}").into());
                }
                delay_us = Some(value);
            },
            "--transfer-bytes-per-second" => {
                if transfer_bytes_per_second.is_some() {
                    return Err("--transfer-bytes-per-second was specified more than once".into());
                }
                let value = parse_u64(&need_value(args, &mut index, flag)?, flag)?;
                if !(1_048_576..=MAX_TRANSFER_BYTES_PER_SECOND).contains(&value) {
                    return Err(
                        "--transfer-bytes-per-second must be in 1048576..=1099511627776".into(),
                    );
                }
                transfer_bytes_per_second = NonZeroU64::new(value);
            },
            "--transfer-delay-policy" => {
                let value = need_value(args, &mut index, flag)?;
                transfer_delay_policy =
                    match value.as_str() {
                        "separate-sleeps" => TransferDelayPolicy::SeparateSleeps,
                        "minimum-service" => TransferDelayPolicy::MinimumService,
                        _ => return Err(
                            "--transfer-delay-policy must be separate-sleeps or minimum-service"
                                .into(),
                        ),
                    };
            },
            "--samples" => {
                if samples.is_some() {
                    return Err("--samples was specified more than once".into());
                }
                let value = parse_usize(&need_value(args, &mut index, flag)?, flag)?;
                if value == 0 || value > MAX_SAMPLES {
                    return Err(format!("--samples must be in 1..={MAX_SAMPLES}").into());
                }
                samples = Some(value);
            },
            "--warmup" => {
                let value = parse_usize(&need_value(args, &mut index, flag)?, flag)?;
                if value > MAX_WARMUP {
                    return Err(format!("--warmup must be <= {MAX_WARMUP}").into());
                }
                warmup = value;
            },
            "--source-revision" => {
                if source_revision.is_some() {
                    return Err("--source-revision was specified more than once".into());
                }
                let value = need_value(args, &mut index, flag)?;
                if value.len() != 40 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                    return Err(
                        "--source-revision must be exactly 40 hexadecimal characters".into(),
                    );
                }
                source_revision = Some(value);
            },
            "--output" => {
                if output.is_some() {
                    return Err("--output was specified more than once".into());
                }
                let value = need_value(args, &mut index, flag)?;
                if value == "-" || value.is_empty() {
                    return Err("--output must be a new regular file path".into());
                }
                output = Some(PathBuf::from(value));
            },
            other => return Err(format!("unknown argument: {other}").into()),
        }
        index += 1;
    }
    let provider = provider.ok_or("--provider is required")?;
    let samples = samples.ok_or("--samples is required")?;
    let source_revision = source_revision.ok_or("--source-revision is required")?;
    let output = output.ok_or("--output is required")?;
    match provider {
        ProviderKind::Range => {
            if max_range_bytes.is_none() || delay_us.is_none() {
                return Err("range provider requires --max-range and --delay-us".into());
            }
        },
        ProviderKind::Bytes | ProviderKind::File | ProviderKind::InstrumentedBytes => {
            if max_range_bytes.is_some()
                || delay_us.is_some()
                || transfer_bytes_per_second.is_some()
            {
                return Err("range controls require --provider range".into());
            }
        },
    }
    if transfer_bytes_per_second.is_none()
        && transfer_delay_policy != TransferDelayPolicy::SeparateSleeps
    {
        return Err("--transfer-delay-policy requires --transfer-bytes-per-second".into());
    }
    if read_ahead_window_bytes.is_some() && (provider != ProviderKind::Range || !trace_ranges) {
        return Err(
            "--read-ahead requires --provider range and --trace-ranges (unmanaged pilot only)"
                .into(),
        );
    }
    if trace_ranges && provider == ProviderKind::Bytes {
        return Err("--trace-ranges requires an instrumented provider".into());
    }
    Ok(Config {
        read_ahead_window_bytes,
        trace_ranges,
        provider,
        max_range_bytes,
        delay_us,
        transfer_bytes_per_second,
        transfer_delay_policy,
        samples,
        warmup,
        source_revision,
        output,
    })
}

fn usage() -> &'static str {
    "docx-provider-lifecycle --provider <bytes|file|instrumented-bytes|range> [--max-range N --delay-us N [--transfer-bytes-per-second N --transfer-delay-policy separate-sleeps|minimum-service]] [--trace-ranges [--read-ahead N]] --samples N --warmup N --source-revision <40 hex chars> --output PATH"
}

/// Runs the benchmark from the arguments following the subcommand.
pub fn run_from_args(args: impl IntoIterator<Item = OsString>) -> Result<(), Box<dyn Error>> {
    let args: Vec<OsString> = args.into_iter().collect();
    if args
        .iter()
        .any(|argument| argument == "--help" || argument == "-h")
    {
        println!("{}", usage());
        return Ok(());
    }
    let config = parse_config(&args)?;
    let corpus = build_corpus()?;
    let staged = (config.provider == ProviderKind::File).then(|| stage_file(&corpus.bytes));
    let staged = staged.transpose()?;
    let limits = ReadLimits::default();
    let cache = SourceCacheLimits::new(CACHE_MAX_BYTES, CACHE_MAX_ENTRIES)?;
    let mut rows = Vec::with_capacity(config.samples);
    for iteration in 0..config.warmup.saturating_add(config.samples) {
        let row = run_sample(
            &config,
            &corpus,
            staged.as_ref().map(|file| file.path.as_path()),
            iteration,
            limits,
            cache,
        )?;
        if iteration >= config.warmup {
            rows.push(row);
        }
    }
    let report = Report {
        media_ranges: config.trace_ranges.then(|| corpus.media_ranges.clone()),
        schema: if config.trace_ranges {
            TRACE_SCHEMA
        } else {
            SCHEMA
        },
        provider_scope: PROVIDER_SCOPE,
        timing_scope: if config.trace_ranges {
            TRACE_TIMING_SCOPE
        } else {
            TIMING_SCOPE
        },
        setup_scope: if config.trace_ranges {
            TRACE_SETUP_SCOPE
        } else {
            SETUP_SCOPE
        },
        allocation_scope: if config.trace_ranges {
            "operation-scoped allocator region; allocation field omitted for normal binary; fresh read-ahead window allocated during setup and reported separately"
        } else {
            ALLOCATION_SCOPE
        },
        source_bytes: corpus.bytes.len(),
        source_sha256: sha256_hex(&corpus.bytes),
        requested_source_revision: config.source_revision.clone(),
        limits: LimitsRecord::from_limits(limits, cache),
        provider: config.record(),
        corpus: corpus.manifest,
        warmup: config.warmup,
        samples: config.samples,
        rows,
    };
    let mut output = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&config.output)?;
    serde_json::to_writer_pretty(&mut output, &report)?;
    output.write_all(b"\n")?;
    output.sync_all()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn source(bytes: &[u8]) -> Arc<dyn ReadAt> {
        Arc::new(OwnedSource::new(bytes.to_vec()))
    }

    struct ChangingSource {
        inner: OwnedSource,
        reads: AtomicUsize,
    }

    impl ReadAt for ChangingSource {
        fn len(&self) -> io::Result<u64> {
            self.inner.len()
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let result = self.inner.read_at(offset, output);
            self.reads.fetch_add(1, Ordering::Relaxed);
            result
        }

        fn version(&self) -> io::Result<SourceVersion> {
            let revision = if self.reads.load(Ordering::Relaxed) == 0 {
                0
            } else {
                1
            };
            Ok(SourceVersion::new(0x0188, revision))
        }
    }

    #[test]
    fn parser_requires_finite_range_configuration() {
        let args = [
            OsString::from("--provider"),
            OsString::from("range"),
            OsString::from("--max-range"),
            OsString::from("64"),
            OsString::from("--delay-us"),
            OsString::from("1"),
            OsString::from("--samples"),
            OsString::from("2"),
            OsString::from("--warmup"),
            OsString::from("1"),
            OsString::from("--source-revision"),
            OsString::from("0123456789012345678901234567890123456789"),
            OsString::from("--output"),
            OsString::from("result.json"),
        ];
        let config = parse_config(&args).expect("valid range configuration");
        assert_eq!(config.provider, ProviderKind::Range);
        assert_eq!(config.max_range_bytes, Some(64));
        assert_eq!(config.delay_us, Some(1));
        assert!(
            parse_config(&[
                OsString::from("--provider"),
                OsString::from("range"),
                OsString::from("--max-range"),
                OsString::from("0"),
            ])
            .is_err()
        );
    }

    fn trace_config(extra: &[&str]) -> Config {
        let mut args: Vec<OsString> = [
            "--provider",
            "range",
            "--max-range",
            "65536",
            "--delay-us",
            "0",
            "--trace-ranges",
            "--samples",
            "1",
            "--source-revision",
            "e44a23396146d504ffc738e0989de896635f02a3",
            "--output",
            "unused.json",
        ]
        .into_iter()
        .map(OsString::from)
        .collect();
        args.extend(extra.iter().map(OsString::from));
        parse_config(&args).expect("finite traced configuration")
    }

    #[test]
    fn read_ahead_requires_traced_range_provider_and_finite_window() {
        let base = trace_config(&["--read-ahead", "4096"]);
        assert_eq!(base.read_ahead_window_bytes, Some(4096));
        for window in ["0", "65537"] {
            let args: Vec<OsString> = ["--read-ahead", window]
                .into_iter()
                .map(OsString::from)
                .collect();
            assert!(parse_config(&args).is_err());
        }
        let args: Vec<OsString> = [
            "--provider",
            "instrumented-bytes",
            "--trace-ranges",
            "--read-ahead",
            "4096",
            "--samples",
            "1",
            "--source-revision",
            "e44a23396146d504ffc738e0989de896635f02a3",
            "--output",
            "unused.json",
        ]
        .into_iter()
        .map(OsString::from)
        .collect();
        assert!(parse_config(&args).is_err());
    }

    #[test]
    fn read_ahead_lifecycle_preserves_text_and_reports_physical_overfetch() {
        let corpus = build_corpus().expect("accepted corpus");
        let limits = ReadLimits::default();
        let cache = SourceCacheLimits::new(CACHE_MAX_BYTES, CACHE_MAX_ENTRIES).expect("cache");
        let baseline =
            run_sample(&trace_config(&[]), &corpus, None, 0, limits, cache).expect("baseline");
        let config = trace_config(&["--read-ahead", "4096"]);
        let candidate = run_sample(&config, &corpus, None, 0, limits, cache).expect("candidate");
        assert_eq!(baseline.actual_text_sha256, candidate.actual_text_sha256);
        assert_eq!(candidate.actual_text_bytes, 10_000);
        for row in [&baseline, &candidate] {
            let diagnostics = row.cache.expect("cache diagnostics");
            assert_eq!(diagnostics.open_successful_loads, 0);
            assert_eq!(diagnostics.successful_loads, 1);
            assert!(!diagnostics.budget_managed);
            assert_eq!(
                row.reads
                    .logical_media_range_proof
                    .as_ref()
                    .expect("logical media")
                    .returned_overlap_bytes,
                Some(0)
            );
        }
        let baseline_calls = baseline
            .reads
            .wrapper
            .logical_calls
            .expect("baseline calls");
        let candidate_calls = candidate
            .reads
            .wrapper
            .logical_calls
            .expect("candidate calls");
        assert!(candidate_calls < baseline_calls);
        assert!(candidate.reads.wrapper.returned_bytes > baseline.reads.wrapper.returned_bytes);
        assert!(
            candidate
                .reads
                .media_range_proof
                .returned_overlap_bytes
                .expect("physical media")
                > 0
        );
        // A transport cap smaller than the window must still reconstruct text.
        let mut short_config = config;
        short_config.max_range_bytes = Some(64);
        let short =
            run_sample(&short_config, &corpus, None, 0, limits, cache).expect("short fills");
        assert_eq!(short.actual_text_sha256, baseline.actual_text_sha256);
    }

    #[test]
    fn bounded_range_source_reconstructs_exact_bytes() {
        let bytes = (0_u8..=127).collect::<Vec<_>>();
        let adapter = PptxRangeSource::with_limits(source(&bytes), Some(3), Some(Duration::ZERO));
        let mut reconstructed = Vec::new();
        let mut offset = 0_u64;
        loop {
            let mut chunk = [0_u8; 32];
            let read = adapter.read_at(offset, &mut chunk).expect("bounded read");
            if read == 0 {
                break;
            }
            reconstructed.extend_from_slice(&chunk[..read]);
            offset += read as u64;
        }
        assert_eq!(reconstructed, bytes);
        let snapshot = adapter.snapshot().expect("range counters");
        assert!(snapshot.short_reads > 0);
        assert_eq!(snapshot.returned_bytes, bytes.len() as u64);
    }

    #[test]
    fn media_proof_is_available_only_for_instrumented_offsets() {
        let state = CountingState::default();
        state.record(0, 4, 4).expect("record range");
        let snapshot = state.snapshot().expect("counter snapshot");
        let range = 10..20;
        let proof = media_proof(Some(&snapshot), std::slice::from_ref(&range));
        assert_eq!(proof.status, "proved_no_media_overlap");
        let unavailable = media_proof(None, std::slice::from_ref(&range));
        assert_eq!(unavailable.availability, "unavailable");
    }

    #[test]
    fn source_change_is_rejected_by_the_typed_owner() {
        let mut package = OwnedPackage::new().expect("small DOCX package");
        package
            .document_mut()
            .expect("document")
            .add_paragraph_with_text("source-change");
        let mut bytes = Cursor::new(Vec::new());
        package.to_stream(&mut bytes).expect("serialize DOCX");
        let source: Arc<dyn ReadAt> = Arc::new(ChangingSource {
            inner: OwnedSource::new(bytes.into_inner()),
            reads: AtomicUsize::new(0),
        });
        let result = SourcePackage::from_read_at(source).and_then(|owner| owner.document());
        assert!(result.is_err(), "changed source must fail closed");
    }

    #[test]
    fn corpus_builder_keeps_the_historical_0188_gate() {
        let corpus = build_corpus().expect("accepted 0188 corpus");
        assert_eq!(corpus.manifest.archive_bytes, CORPUS_ARCHIVE_BYTES);
        assert_eq!(corpus.manifest.archive_sha256, CORPUS_ARCHIVE_SHA256);
        assert_eq!(corpus.manifest.media_member_count, MEDIA_COUNT);
        assert_eq!(corpus.manifest.archive_member_count, ARCHIVE_MEMBER_COUNT);
        assert_eq!(
            corpus.expected_text.len(),
            corpus.manifest.expected_text_bytes
        );
    }
}
