//! DOCX managed/unmanaged one-edit/save measurements across explicit providers.
//!
//! This module keeps provider setup and independent correctness work outside the
//! operation clock.  Each measured iteration creates a fresh source adapter,
//! opens the source-backed DOCX owner, stages and commits one paragraph edit,
//! publishes to a bounded sequential sink, and drops the commit before the
//! clock stops.  The output remains owned after the clock so exact byte,
//! semantic, media, replay, and inverse checks run outside the reported
//! lifecycle. Commit diagnostics and source/candidate XML identity checks are
//! timed; each entry point records its source-version fence boundaries.
//!
//! The `file` provider is deliberately a recently-written, warm-cache
//! observation. This 0495 entry point has no cold-filesystem mode.

#![allow(clippy::module_name_repetitions)]

use std::{
    error::Error,
    ffi::OsString,
    fmt::Write as _,
    fs::{self, OpenOptions},
    io::{self, Read, Write},
    num::{NonZeroU64, NonZeroUsize},
    ops::Range,
    path::{Path, PathBuf},
    sync::{
        Arc, Mutex,
        atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering},
    },
    time::{Duration, Instant},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, FileSource, Limits, OwnedSource,
    Position, ReadAt, Resource, SourceVersion,
};
use litchi_docx::{document::TransactionError, source_backed};
use litchi_opc::{ReadLimits, SourceCacheLimits};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::pptx_range_source::{
    PptxRangeSource, PptxRangeSourceConfig, PptxRangeSourceSnapshot, TransferDelayPolicy,
};

const SCHEMA: &str = "docx_edit_provider_managed_v1";
const REPORT_VERSION: u32 = 1;
const CASE_NAME: &str = "docx_opened_document_managed_vs_unmanaged_one_paragraph_edit_save";
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 1_000;
const MAX_RANGE_BYTES: usize = 1_048_576;
const MAX_DELAY_US: u64 = 100_000;
const MIN_TRANSFER_BYTES_PER_SECOND: u64 = 1_048_576;
const MAX_TRANSFER_BYTES_PER_SECOND: u64 = 1_099_511_627_776;
const DEFAULT_SHORT_RANGE_BYTES: usize = 256;
const DEFAULT_DELAY_RANGE_BYTES: usize = 65_536;
const DEFAULT_DELAY_US: u64 = 1_000;
const DEFAULT_TRANSFER_BYTES_PER_SECOND: u64 = 100 * 1024 * 1024;
const MAX_TRACKED_RANGES: usize = 131_072;
const SINK_MAX_WRITE_BYTES: u64 = 1024 * 1024;
const CORPUS_VERSION: &str = "0188-media-v1";
const CORPUS_GENERATOR: &str = "litchi-docx-source-edit-media-v1";
const EXPECTED_TEXT_BYTES: usize = 10_000;
const EXPECTED_TEXT_SHA256: &str =
    "ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af";
const EXPECTED_TEXT_SCOPE: &str = "original deterministic corpus text identity before the selected paragraph replacement; it is not the changed output text identity";
const ALLOCATION_SCOPE: &str = "operation-scoped global-system-allocator region surrounds the timed opened document edit, commit, publication, diagnostics, and package/document drop; normal binaries report unavailable and allocator binaries report the separate region sample";
const FILE_SCOPE: &str = "recently-written FileSource reopened for each iteration; warm-cache observation only, with no cold-filesystem claim";
const TIMING_SCOPE: &str = "fresh source-backed DOCX open + one paragraph edit staging/commit + sequential publication + commit/package/document drops; provider construction, sink reservation and retained output are outside; commit diagnostics and source/candidate XML identity comparisons are inside, while output hashing, semantic/media verification and preflight patch oracles are outside";
const SETUP_SCOPE: &str = "deterministic corpus construction, expected publication, source adapter construction, warm file staging, sink reservation, and oracle preparation are outside the clock";
const PROVIDER_SCOPE: &str = "explicit caller-owned positional provider matrix for one DOCX paragraph replacement; no ambient filesystem or network behavior";
const PHYSICAL_SCOPE: &str = "caller-visible ReadAt calls at the selected provider/source boundary; requested and returned ranges are transport observations, not filesystem or network I/O";
const RANGE_SCOPE: &str = "logical ReadAt calls and adapter calls observed by harness wrappers; zero-length caller calls are counted separately; counters are transport-model evidence, not physical network or filesystem observations";
const ZERO_LENGTH_SCOPE: &str = "zero-length caller ReadAt calls are delegated to the wrapped provider, preserve its return/error behavior, and are counted separately from nonempty range totals";
const MEDIA_SCOPE: &str = "source compressed ranges for word/media members; exact output and OPC semantic checks additionally prove unchanged media payloads";
const BUDGET_SCOPE: &str = "managed budget gauges are sampled before package open, after publication while the commit remains live, and after package/document/commit drops; cumulative input/output/work counters are not release gauges";
const PHASE_TIMING_SCOPE: &str = "opt-in wall-clock Instant intervals nested inside the full lifecycle clock: open, edit staging, commit, diagnostics/XML identity, publication, published snapshot drop, and commit drop; these are not CPU-time measurements";
const PHASE_RESIDUAL_SCOPE: &str = "full lifecycle time minus the listed phase intervals; includes phase-boundary arithmetic, budget evidence, result handling, and any package work not assigned to a named phase";
const PHASE_INSTRUMENTATION_SCOPE: &str = "phase-clock overhead is included in the full lifecycle and is not isolated by a second control clock; phase fields are absent unless --phase-diagnostics is enabled";
const PHASE_ALLOCATION_SCOPE: &str = "no nested phase allocation regions are reported: the shared allocator region is non-reentrant, so allocation remains one full-lifecycle sample";
const UNMANAGED_BUDGET_REASON: &str =
    "unmanaged-api uses the compatibility constructor without an ExecutionContext";
const BUDGET_MEMORY_BYTES: u64 = 64 * 1024 * 1024;
const BUDGET_INPUT_BYTES: u64 = 64 * 1024 * 1024;
const BUDGET_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
const BUDGET_OBJECTS: u64 = 1_000_000;
const BUDGET_DEPTH: u64 = 1_024;
const BUDGET_WORK: u64 = 2 * 1024 * 1024 * 1024;
const MAX_IN_FLIGHT_BYTES: u64 = 64 * 1024 * 1024;
const CACHE_MAX_BYTES: usize = 8 * 1024 * 1024;
const CACHE_MAX_ENTRIES: usize = 128;

static NEXT_STAGE_ID: AtomicUsize = AtomicUsize::new(0);

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ProviderKind {
    Owned,
    Instrumented,
    Short,
    Delayed,
    File,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
enum ApiMode {
    UnmanagedApi,
    ManagedApi,
}

impl ApiMode {
    const fn name(self) -> &'static str {
        match self {
            Self::UnmanagedApi => "unmanaged-api",
            Self::ManagedApi => "managed-api",
        }
    }

    const fn managed(self) -> bool {
        matches!(self, Self::ManagedApi)
    }
}

impl ProviderKind {
    const fn name(self) -> &'static str {
        match self {
            Self::Owned => "owned",
            Self::Instrumented => "instrumented",
            Self::Short => "short",
            Self::Delayed => "delayed",
            Self::File => "file",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Config {
    api: ApiMode,
    provider: ProviderKind,
    phase_diagnostics: bool,
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
struct ProviderConfig {
    name: &'static str,
    provider: &'static str,
    kind: &'static str,
    trace_ranges: bool,
    short_read_bytes: Option<usize>,
    zero_length_scope: &'static str,
    max_range_bytes: Option<usize>,
    delay_us: Option<u64>,
    transfer_bytes_per_second: Option<NonZeroU64>,
    transfer_delay_policy: TransferDelayPolicy,
    source_construction: &'static str,
    file_scope: &'static str,
    range_scope: &'static str,
}

impl Config {
    fn provider_config(&self) -> ProviderConfig {
        ProviderConfig {
            name: self.provider.name(),
            provider: self.provider.name(),
            kind: match self.provider {
                ProviderKind::Owned => "owned",
                ProviderKind::Instrumented => "instrumented",
                ProviderKind::Short => "short-read",
                ProviderKind::Delayed if self.delay_us == Some(0) => "range-control",
                ProviderKind::Delayed => "range-delay",
                ProviderKind::File => "file-warm",
            },
            trace_ranges: !matches!(self.provider, ProviderKind::Owned | ProviderKind::File),
            short_read_bytes: (self.provider == ProviderKind::Short)
                .then_some(self.max_range_bytes.unwrap_or(DEFAULT_SHORT_RANGE_BYTES)),
            zero_length_scope: ZERO_LENGTH_SCOPE,
            max_range_bytes: self.max_range_bytes,
            delay_us: self.delay_us,
            transfer_bytes_per_second: self.transfer_bytes_per_second,
            transfer_delay_policy: self.transfer_delay_policy,
            source_construction: "source adapters and FileSource handles are constructed before the operation clock",
            file_scope: FILE_SCOPE,
            range_scope: RANGE_SCOPE,
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ReadLimitsRecord {
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
}

impl ReadLimitsRecord {
    fn from_limits(limits: ReadLimits) -> Self {
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
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct BudgetPolicyRecord {
    managed: bool,
    mode_reason: &'static str,
    memory_bytes: Option<u64>,
    input_bytes: Option<u64>,
    output_bytes: Option<u64>,
    objects: Option<u64>,
    depth: Option<u64>,
    work: Option<u64>,
}

#[derive(Clone, Debug, Serialize)]
struct LimitsRecord {
    read_limits: ReadLimitsRecord,
    cache_max_bytes: u64,
    cache_max_entries: u64,
    resource_budget: BudgetPolicyRecord,
    max_tracked_ranges: u64,
    sink_max_write_bytes: u64,
}

fn limits_record(api: ApiMode) -> LimitsRecord {
    let read_limits = ReadLimits::default();
    let cache_limits = SourceCacheLimits::default();
    LimitsRecord {
        read_limits: ReadLimitsRecord::from_limits(read_limits),
        cache_max_bytes: cache_limits.max_bytes() as u64,
        cache_max_entries: cache_limits.max_entries() as u64,
        resource_budget: BudgetPolicyRecord {
            managed: api.managed(),
            mode_reason: if api.managed() {
                "managed-api uses the explicit finite ExecutionContext"
            } else {
                UNMANAGED_BUDGET_REASON
            },
            memory_bytes: api.managed().then_some(BUDGET_MEMORY_BYTES),
            input_bytes: api.managed().then_some(BUDGET_INPUT_BYTES),
            output_bytes: api.managed().then_some(BUDGET_OUTPUT_BYTES),
            objects: api.managed().then_some(BUDGET_OBJECTS),
            depth: api.managed().then_some(BUDGET_DEPTH),
            work: api.managed().then_some(BUDGET_WORK),
        },
        max_tracked_ranges: MAX_TRACKED_RANGES as u64,
        sink_max_write_bytes: SINK_MAX_WRITE_BYTES,
    }
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

#[derive(Clone, Debug, Serialize, Deserialize)]
struct RangeRecord {
    offset: u64,
    requested: u64,
    returned: u64,
}

#[derive(Clone, Debug)]
struct ReadStatsSnapshot {
    calls: u64,
    empty_calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    short_reads: u64,
    min_request_bytes: Option<u64>,
    max_request_bytes: Option<u64>,
    ranges: Vec<RangeRecord>,
}

#[derive(Debug)]
struct ReadStats {
    calls: AtomicU64,
    empty_calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    short_reads: AtomicU64,
    min_request_bytes: AtomicU64,
    max_request_bytes: AtomicU64,
    failed: AtomicBool,
    ranges_overflowed: AtomicBool,
    ranges: Mutex<Vec<RangeRecord>>,
}

impl ReadStats {
    fn new() -> io::Result<Self> {
        let mut ranges = Vec::new();
        ranges
            .try_reserve_exact(MAX_TRACKED_RANGES)
            .map_err(|error| {
                io::Error::other(format!("DOCX provider range reservation failed: {error}"))
            })?;
        Ok(Self {
            calls: AtomicU64::new(0),
            empty_calls: AtomicU64::new(0),
            requested_bytes: AtomicU64::new(0),
            returned_bytes: AtomicU64::new(0),
            short_reads: AtomicU64::new(0),
            min_request_bytes: AtomicU64::new(u64::MAX),
            max_request_bytes: AtomicU64::new(0),
            failed: AtomicBool::new(false),
            ranges_overflowed: AtomicBool::new(false),
            ranges: Mutex::new(ranges),
        })
    }

    fn add(counter: &AtomicU64, value: u64, failed: &AtomicBool) -> io::Result<()> {
        if counter
            .fetch_update(Ordering::AcqRel, Ordering::Acquire, |current| {
                current.checked_add(value)
            })
            .is_err()
        {
            failed.store(true, Ordering::Release);
            return Err(io::Error::other("DOCX provider counter overflow"));
        }
        Ok(())
    }

    fn record_empty(&self) -> io::Result<()> {
        Self::add(&self.empty_calls, 1, &self.failed)
    }

    fn record(&self, offset: u64, requested: usize, returned: usize) -> io::Result<()> {
        let requested = u64::try_from(requested)
            .map_err(|_| io::Error::other("DOCX provider request does not fit u64"))?;
        let returned = u64::try_from(returned)
            .map_err(|_| io::Error::other("DOCX provider result does not fit u64"))?;
        Self::add(&self.calls, 1, &self.failed)?;
        Self::add(&self.requested_bytes, requested, &self.failed)?;
        Self::add(&self.returned_bytes, returned, &self.failed)?;
        if returned < requested {
            Self::add(&self.short_reads, 1, &self.failed)?;
        }
        self.min_request_bytes
            .fetch_min(requested, Ordering::AcqRel);
        self.max_request_bytes
            .fetch_max(requested, Ordering::AcqRel);
        let _end = offset
            .checked_add(requested)
            .ok_or_else(|| io::Error::other("DOCX provider range overflows u64"))?;
        let mut ranges = self
            .ranges
            .lock()
            .map_err(|_| io::Error::other("DOCX provider range mutex poisoned"))?;
        if ranges.len() < MAX_TRACKED_RANGES {
            ranges.push(RangeRecord {
                offset,
                requested,
                returned,
            });
        } else {
            self.ranges_overflowed.store(true, Ordering::Release);
        }
        Ok(())
    }

    fn snapshot(&self) -> io::Result<ReadStatsSnapshot> {
        if self.failed.load(Ordering::Acquire) {
            return Err(io::Error::other("DOCX provider counters are unavailable"));
        }
        if self.ranges_overflowed.load(Ordering::Acquire) {
            return Err(io::Error::other(
                "DOCX provider range trace exceeded its bounded evidence capacity",
            ));
        }
        let calls = self.calls.load(Ordering::Acquire);
        let ranges = self
            .ranges
            .lock()
            .map_err(|_| io::Error::other("DOCX provider range mutex poisoned"))?
            .clone();
        if ranges.len() as u64 != calls {
            return Err(io::Error::other(
                "DOCX provider range trace count differs from call count",
            ));
        }
        Ok(ReadStatsSnapshot {
            calls,
            empty_calls: self.empty_calls.load(Ordering::Acquire),
            requested_bytes: self.requested_bytes.load(Ordering::Acquire),
            returned_bytes: self.returned_bytes.load(Ordering::Acquire),
            short_reads: self.short_reads.load(Ordering::Acquire),
            min_request_bytes: (calls != 0).then(|| self.min_request_bytes.load(Ordering::Acquire)),
            max_request_bytes: (calls != 0).then(|| self.max_request_bytes.load(Ordering::Acquire)),
            ranges,
        })
    }
}

struct CountingReadAt {
    inner: Arc<dyn ReadAt>,
    stats: Arc<ReadStats>,
}

impl CountingReadAt {
    fn new(inner: Arc<dyn ReadAt>, stats: Arc<ReadStats>) -> Self {
        Self { inner, stats }
    }
}

impl ReadAt for CountingReadAt {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let returned = self.inner.read_at(offset, output)?;
        if output.is_empty() {
            if returned != 0 {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "DOCX provider returned bytes for an empty request",
                ));
            }
            self.stats.record_empty()?;
            return Ok(0);
        }
        if returned > output.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "DOCX provider returned more bytes than requested",
            ));
        }
        self.stats.record(offset, output.len(), returned)?;
        Ok(returned)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

struct Provider {
    source: Arc<dyn ReadAt>,
    logical_stats: Option<Arc<ReadStats>>,
    physical_stats: Option<Arc<ReadStats>>,
    range_adapter: Option<Arc<PptxRangeSource>>,
}

fn provider(
    config: &Config,
    bytes: &[u8],
    file_path: Option<&Path>,
) -> Result<Provider, Box<dyn Error>> {
    let base: Arc<dyn ReadAt> = match config.provider {
        ProviderKind::Owned
        | ProviderKind::Instrumented
        | ProviderKind::Short
        | ProviderKind::Delayed => Arc::new(OwnedSource::new(bytes.to_vec())),
        ProviderKind::File => {
            let path = file_path.ok_or("file provider requires a staged source path")?;
            Arc::new(FileSource::open(path)?)
        },
    };

    let (source, logical_stats, physical_stats, range_adapter): (
        Arc<dyn ReadAt>,
        Option<Arc<ReadStats>>,
        Option<Arc<ReadStats>>,
        Option<Arc<PptxRangeSource>>,
    ) = match config.provider {
        ProviderKind::Owned => (base, None, None, None),
        ProviderKind::Instrumented | ProviderKind::File => {
            let stats = Arc::new(ReadStats::new()?);
            let source: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(base, Arc::clone(&stats)));
            (source, Some(stats), None, None)
        },
        ProviderKind::Short | ProviderKind::Delayed => {
            let physical_stats = Arc::new(ReadStats::new()?);
            let physical: Arc<dyn ReadAt> =
                Arc::new(CountingReadAt::new(base, Arc::clone(&physical_stats)));
            let max_range = config
                .max_range_bytes
                .ok_or("range provider is missing a maximum range")?;
            let delay = config.delay_us.map(Duration::from_micros);
            let mut range_config = PptxRangeSourceConfig::new(Some(max_range), delay);
            range_config.transfer_bytes_per_second = config.transfer_bytes_per_second;
            range_config.transfer_delay_policy = config.transfer_delay_policy;
            let adapter = Arc::new(PptxRangeSource::new(physical, range_config));
            let logical_stats = Arc::new(ReadStats::new()?);
            let source: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(
                adapter.clone(),
                Arc::clone(&logical_stats),
            ));
            (
                source,
                Some(logical_stats),
                Some(physical_stats),
                Some(adapter),
            )
        },
    };
    Ok(Provider {
        source,
        logical_stats,
        physical_stats,
        range_adapter,
    })
}

#[derive(Clone, Debug, Serialize)]
struct ReadCounter {
    availability: &'static str,
    scope: &'static str,
    calls: Option<u64>,
    empty_calls: Option<u64>,
    requested_bytes: Option<u64>,
    returned_bytes: Option<u64>,
    short_reads: Option<u64>,
    min_request_bytes: Option<u64>,
    max_request_bytes: Option<u64>,
    traced_ranges: Option<usize>,
    ranges: Option<Vec<RangeRecord>>,
    requested_media_overlap_bytes: Option<u64>,
    returned_media_overlap_bytes: Option<u64>,
}

impl ReadCounter {
    fn unavailable(scope: &'static str) -> Self {
        Self {
            availability: "unavailable",
            scope,
            calls: None,
            empty_calls: None,
            requested_bytes: None,
            returned_bytes: None,
            short_reads: None,
            min_request_bytes: None,
            max_request_bytes: None,
            traced_ranges: None,
            ranges: None,
            requested_media_overlap_bytes: None,
            returned_media_overlap_bytes: None,
        }
    }

    fn from_snapshot(
        snapshot: &ReadStatsSnapshot,
        scope: &'static str,
        media_ranges: &[Range<u64>],
    ) -> Self {
        let mut requested_media = 0_u64;
        let mut returned_media = 0_u64;
        for observed in &snapshot.ranges {
            let Some(requested_end) = observed.offset.checked_add(observed.requested) else {
                continue;
            };
            let Some(returned_end) = observed.offset.checked_add(observed.returned) else {
                continue;
            };
            for media in media_ranges {
                requested_media =
                    requested_media.saturating_add(overlap(observed.offset, requested_end, media));
                returned_media =
                    returned_media.saturating_add(overlap(observed.offset, returned_end, media));
            }
        }
        Self {
            availability: "available",
            scope,
            calls: Some(snapshot.calls),
            empty_calls: Some(snapshot.empty_calls),
            requested_bytes: Some(snapshot.requested_bytes),
            returned_bytes: Some(snapshot.returned_bytes),
            short_reads: Some(snapshot.short_reads),
            min_request_bytes: snapshot.min_request_bytes,
            max_request_bytes: snapshot.max_request_bytes,
            traced_ranges: Some(snapshot.ranges.len()),
            ranges: Some(snapshot.ranges.clone()),
            requested_media_overlap_bytes: Some(requested_media),
            returned_media_overlap_bytes: Some(returned_media),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct RangeAdapterRecord {
    availability: &'static str,
    logical_calls: Option<u64>,
    requested_bytes: Option<u64>,
    returned_bytes: Option<u64>,
    short_reads: Option<u64>,
    delayed_calls: Option<u64>,
    transfer_paced_calls: Option<u64>,
    transfer_delay_ns: Option<u64>,
}

impl RangeAdapterRecord {
    fn from_snapshot(snapshot: PptxRangeSourceSnapshot) -> Self {
        Self {
            availability: "available",
            logical_calls: Some(snapshot.logical_calls),
            requested_bytes: Some(snapshot.requested_bytes),
            returned_bytes: Some(snapshot.returned_bytes),
            short_reads: Some(snapshot.short_reads),
            delayed_calls: Some(snapshot.delayed_calls),
            transfer_paced_calls: Some(snapshot.transfer_paced_calls),
            transfer_delay_ns: Some(snapshot.transfer_delay_ns),
        }
    }

    const fn unavailable() -> Self {
        Self {
            availability: "unavailable",
            logical_calls: None,
            requested_bytes: None,
            returned_bytes: None,
            short_reads: None,
            delayed_calls: None,
            transfer_paced_calls: None,
            transfer_delay_ns: None,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct ReadEvidence {
    logical: ReadCounter,
    physical: ReadCounter,
    range_adapter: RangeAdapterRecord,
    media_scope: &'static str,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ResourceSnapshot {
    availability: &'static str,
    managed: bool,
    memory_used: Option<u64>,
    input_bytes_used: Option<u64>,
    output_bytes_used: Option<u64>,
    objects_used: Option<u64>,
    depth_used: Option<u64>,
    work_used: Option<u64>,
}

impl ResourceSnapshot {
    fn unavailable() -> Self {
        Self {
            availability: "unavailable",
            managed: false,
            memory_used: None,
            input_bytes_used: None,
            output_bytes_used: None,
            objects_used: None,
            depth_used: None,
            work_used: None,
        }
    }

    fn from_budget(budget: &Budget) -> Self {
        Self {
            availability: "available",
            managed: true,
            memory_used: Some(budget.used(Resource::Memory)),
            input_bytes_used: Some(budget.used(Resource::InputBytes)),
            output_bytes_used: Some(budget.used(Resource::OutputBytes)),
            objects_used: Some(budget.used(Resource::Objects)),
            depth_used: Some(budget.used(Resource::Depth)),
            work_used: Some(budget.used(Resource::Work)),
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct CacheBudgetSnapshot {
    availability: &'static str,
    managed: bool,
    budget_reservation_failures: Option<u64>,
    budget_memory_used: Option<u64>,
    budget_cache_reserved_bytes: Option<u64>,
    budget_input_bytes_used: Option<u64>,
    budget_output_bytes_used: Option<u64>,
    budget_work_used: Option<u64>,
    budget_objects_used: Option<u64>,
    budget_catalog_reserved_objects: Option<u64>,
    budget_cache_reserved_objects: Option<u64>,
    retained_entries: Option<usize>,
    retained_bytes: Option<usize>,
    in_flight_loads: Option<usize>,
}

impl CacheBudgetSnapshot {
    fn unavailable() -> Self {
        Self {
            availability: "unavailable",
            managed: false,
            budget_reservation_failures: None,
            budget_memory_used: None,
            budget_cache_reserved_bytes: None,
            budget_input_bytes_used: None,
            budget_output_bytes_used: None,
            budget_work_used: None,
            budget_objects_used: None,
            budget_catalog_reserved_objects: None,
            budget_cache_reserved_objects: None,
            retained_entries: None,
            retained_bytes: None,
            in_flight_loads: None,
        }
    }

    fn from_diagnostics(diagnostics: litchi_opc::SourceCacheDiagnostics) -> Self {
        Self {
            availability: "available",
            managed: diagnostics.budget_managed,
            budget_reservation_failures: Some(diagnostics.budget_reservation_failures),
            budget_memory_used: Some(diagnostics.budget_memory_used),
            budget_cache_reserved_bytes: Some(diagnostics.budget_cache_reserved_bytes),
            budget_input_bytes_used: Some(diagnostics.budget_input_bytes_used),
            budget_output_bytes_used: Some(diagnostics.budget_output_bytes_used),
            budget_work_used: Some(diagnostics.budget_work_used),
            budget_objects_used: Some(diagnostics.budget_objects_used),
            budget_catalog_reserved_objects: Some(diagnostics.budget_catalog_reserved_objects),
            budget_cache_reserved_objects: Some(diagnostics.budget_cache_reserved_objects),
            retained_entries: Some(diagnostics.retained_entries),
            retained_bytes: Some(diagnostics.retained_bytes),
            in_flight_loads: Some(diagnostics.in_flight_loads),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct BudgetEvidence {
    scope: &'static str,
    before: ResourceSnapshot,
    live: ResourceSnapshot,
    after_drop: ResourceSnapshot,
    cache_before: CacheBudgetSnapshot,
    cache_live: CacheBudgetSnapshot,
    memory_released_to_baseline: Option<bool>,
    objects_released_to_baseline: Option<bool>,
    reservation_failures: Option<u64>,
}

impl BudgetEvidence {
    fn unmanaged() -> Self {
        Self {
            scope: BUDGET_SCOPE,
            before: ResourceSnapshot::unavailable(),
            live: ResourceSnapshot::unavailable(),
            after_drop: ResourceSnapshot::unavailable(),
            cache_before: CacheBudgetSnapshot::unavailable(),
            cache_live: CacheBudgetSnapshot::unavailable(),
            memory_released_to_baseline: None,
            objects_released_to_baseline: None,
            reservation_failures: None,
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct SinkRecord {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
}

#[derive(Debug)]
struct SequentialSink {
    bytes: Vec<u8>,
    maximum: u64,
    max_write: u64,
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
}

impl SequentialSink {
    fn new(maximum: u64, max_write: u64) -> Result<Self, Box<dyn Error>> {
        let capacity = usize::try_from(maximum)
            .map_err(|_| "DOCX sequential sink ceiling does not fit usize")?;
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(capacity)?;
        Ok(Self {
            bytes,
            maximum,
            max_write,
            accepted_bytes: 0,
            write_calls: 0,
            largest_write: 0,
        })
    }

    const fn record(&self) -> SinkRecord {
        SinkRecord {
            accepted_bytes: self.accepted_bytes,
            write_calls: self.write_calls,
            largest_write: self.largest_write,
        }
    }
}

impl Write for SequentialSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let length = u64::try_from(bytes.len())
            .map_err(|_| io::Error::other("DOCX sink write does not fit u64"))?;
        if length > self.max_write {
            return Err(io::Error::other("DOCX sink write exceeds configured bound"));
        }
        let accepted = self
            .accepted_bytes
            .checked_add(length)
            .ok_or_else(|| io::Error::other("DOCX sink byte count overflow"))?;
        if accepted > self.maximum {
            return Err(io::Error::other("DOCX sink output ceiling exceeded"));
        }
        self.bytes.extend_from_slice(bytes);
        self.accepted_bytes = accepted;
        self.write_calls = self
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("DOCX sink write count overflow"))?;
        self.largest_write = self.largest_write.max(length);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Clone, Debug, Serialize)]
struct CacheEvidence {
    successful_loads: usize,
    expected_successful_loads: usize,
    exactly_one_main_part_materialization: bool,
}

#[derive(Clone, Debug, Serialize)]
struct OracleEvidence {
    commit_identity_verified: bool,
    output_exact_bytes: bool,
    semantic_reopen: bool,
    unchanged_media_preserved: bool,
    source_version_unchanged: bool,
    cache_load_count: bool,
    patch_oracles: PatchOracleEvidence,
}

#[derive(Clone, Debug, Serialize)]
struct PatchOracleEvidence {
    scope: &'static str,
    replay_forward: bool,
    inverse_restores_source: bool,
    stale_target_refused: bool,
    foreign_source_refused: bool,
}

#[derive(Clone, Debug, Serialize)]
struct SampleRecord {
    sample_index: usize,
    api: ApiMode,
    latency_ns: u64,
    output_bytes: usize,
    output_sha256: String,
    output_exact_bytes: bool,
    materializations: usize,
    commit_changed: bool,
    commit_operations: usize,
    source_version_before: SourceVersionRecord,
    source_version_after: SourceVersionRecord,
    source_version_unchanged: bool,
    reads: ReadEvidence,
    cache: CacheEvidence,
    sink: SinkRecord,
    oracles: OracleEvidence,
    budget: BudgetEvidence,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation: Option<crate::allocation_metrics::Sample>,
    #[serde(skip_serializing_if = "Option::is_none")]
    phase_diagnostics: Option<PhaseDiagnostics>,
}

#[derive(Clone, Copy, Debug)]
struct PhaseTimings {
    open_ns: u64,
    edit_staging_ns: u64,
    commit_ns: u64,
    diagnostics_xml_identity_ns: u64,
    publication_ns: u64,
    published_snapshot_drop_ns: u64,
    commit_drop_ns: u64,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PhaseDiagnostics {
    schema: &'static str,
    timing_scope: &'static str,
    open_ns: u64,
    edit_staging_ns: u64,
    commit_ns: u64,
    diagnostics_xml_identity_ns: u64,
    publication_ns: u64,
    published_snapshot_drop_ns: u64,
    commit_drop_ns: u64,
    phase_sum_ns: u64,
    lifecycle_residual_ns: u64,
    residual_scope: &'static str,
    instrumentation_overhead_ns: Option<u64>,
    instrumentation_scope: &'static str,
    allocation_scope: &'static str,
}

impl PhaseTimings {
    fn finish(self, lifecycle_ns: u64) -> Result<PhaseDiagnostics, Box<dyn Error>> {
        let phase_sum_ns = self
            .open_ns
            .checked_add(self.edit_staging_ns)
            .and_then(|value| value.checked_add(self.commit_ns))
            .and_then(|value| value.checked_add(self.diagnostics_xml_identity_ns))
            .and_then(|value| value.checked_add(self.publication_ns))
            .and_then(|value| value.checked_add(self.published_snapshot_drop_ns))
            .and_then(|value| value.checked_add(self.commit_drop_ns))
            .ok_or("DOCX phase timing sum overflows nanoseconds")?;
        let lifecycle_residual_ns = lifecycle_ns
            .checked_sub(phase_sum_ns)
            .ok_or("DOCX phase intervals exceed the full lifecycle clock")?;
        Ok(PhaseDiagnostics {
            schema: "docx_managed_edit_phase_diagnostics_v1",
            timing_scope: PHASE_TIMING_SCOPE,
            open_ns: self.open_ns,
            edit_staging_ns: self.edit_staging_ns,
            commit_ns: self.commit_ns,
            diagnostics_xml_identity_ns: self.diagnostics_xml_identity_ns,
            publication_ns: self.publication_ns,
            published_snapshot_drop_ns: self.published_snapshot_drop_ns,
            commit_drop_ns: self.commit_drop_ns,
            phase_sum_ns,
            lifecycle_residual_ns,
            residual_scope: PHASE_RESIDUAL_SCOPE,
            instrumentation_overhead_ns: None,
            instrumentation_scope: PHASE_INSTRUMENTATION_SCOPE,
            allocation_scope: PHASE_ALLOCATION_SCOPE,
        })
    }
}

#[derive(Clone, Debug, Serialize)]
struct PreflightEvidence {
    expected_materializations: usize,
    source_document_xml_sha256: String,
    candidate_document_xml_sha256: String,
    output_exact_source_changed: bool,
    semantic_reopen_verified: bool,
    unchanged_media_preserved: bool,
    replay_forward_verified: bool,
    inverse_restores_source_verified: bool,
    stale_target_refusal_verified: bool,
    foreign_source_refusal_verified: bool,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    version: u32,
    case_name: &'static str,
    benchmark: &'static str,
    api: ApiMode,
    api_name: &'static str,
    provider_scope: &'static str,
    timing_scope: &'static str,
    setup_scope: &'static str,
    allocation_scope: &'static str,
    physical_scope: &'static str,
    range_scope: &'static str,
    zero_length_scope: &'static str,
    file_scope: &'static str,
    budget_scope: &'static str,
    corpus_version: &'static str,
    corpus_generator: &'static str,
    source_archive_sha256: String,
    source_archive_bytes: usize,
    source_bytes: usize,
    source_sha256: String,
    expected_text_bytes: usize,
    expected_text_sha256: &'static str,
    expected_text_scope: &'static str,
    expected_archive_members: usize,
    expected_output_sha256: String,
    expected_output_bytes: usize,
    source_revision: String,
    requested_source_revision: String,
    limits: LimitsRecord,
    binary_sha256: String,
    binary_bytes: u64,
    current_exe: String,
    provider: ProviderConfig,
    corpus: crate::CorpusManifest,
    preflight: PreflightEvidence,
    warmup: usize,
    samples: usize,
    allocator: &'static str,
    instrumentation: &'static str,
    rows: Vec<SampleRecord>,
}

#[derive(Debug)]
struct Prepared {
    corpus: crate::Corpus,
    expected_output: Vec<u8>,
    source_document_xml: Vec<u8>,
    candidate_document_xml: Vec<u8>,
    expected_output_sha256: String,
    media_ranges: Vec<Range<u64>>,
    preflight: PreflightEvidence,
    target_position: usize,
    replacement_text: String,
}

fn overlap(a: u64, b: u64, range: &Range<u64>) -> u64 {
    b.min(range.end).saturating_sub(a.max(range.start))
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut result = String::with_capacity(64);
    for byte in digest {
        let _ = write!(result, "{byte:02x}");
    }
    result
}

fn media_ranges(bytes: &[u8]) -> Result<Vec<Range<u64>>, Box<dyn Error>> {
    Ok(crate::zip_member_ranges(bytes)?
        .into_iter()
        .filter_map(|(name, range)| name.starts_with("word/media/").then_some(range))
        .collect())
}

fn prepare() -> Result<Prepared, Box<dyn Error>> {
    let corpus = crate::build_docx_source_edit_corpus()?;
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected_output = Vec::new();
    let (expected_materializations, commit) =
        crate::publish_docx_source_edit(source, &mut expected_output)?;
    if expected_materializations != 1 {
        return Err(format!(
            "DOCX edit preflight materialized {expected_materializations} main Parts; expected exactly one"
        )
        .into());
    }
    if expected_output == corpus.archive {
        return Err("DOCX edit preflight produced an exact no-op".into());
    }
    crate::verify_docx_source_edit_output(&corpus, &expected_output)?;
    let source_document_xml = commit.patch().source().xml_bytes().to_vec();
    let candidate_document_xml = commit.snapshot().xml_bytes().to_vec();
    let replayed = commit.patch().apply(commit.patch().source())?;
    let replay_forward_verified = replayed.xml_bytes() == commit.snapshot().xml_bytes();
    let restored = commit.patch().inverse().apply(commit.snapshot())?;
    let inverse_restores_source_verified =
        restored.xml_bytes() == commit.patch().source().xml_bytes();
    let stale_target_refusal_verified = commit.patch().apply(commit.snapshot()).is_err();
    let foreign_package = litchi_docx::source_backed::Package::from_read_at(Arc::new(
        OwnedSource::new(expected_output.clone()),
    ))?;
    let foreign_snapshot = foreign_package.document_snapshot()?;
    let foreign_source_refusal_verified = commit.patch().apply(&foreign_snapshot).is_err();
    if !replay_forward_verified
        || !inverse_restores_source_verified
        || !stale_target_refusal_verified
        || !foreign_source_refusal_verified
    {
        return Err("DOCX edit patch preflight oracle failed".into());
    }
    let preflight = PreflightEvidence {
        expected_materializations,
        source_document_xml_sha256: sha256_hex(&source_document_xml),
        candidate_document_xml_sha256: sha256_hex(&candidate_document_xml),
        output_exact_source_changed: true,
        semantic_reopen_verified: true,
        unchanged_media_preserved: true,
        replay_forward_verified,
        inverse_restores_source_verified,
        stale_target_refusal_verified,
        foreign_source_refusal_verified,
    };
    drop(commit);
    let target_position = corpus
        .manifest
        .target_entry
        .strip_prefix("paragraph:")
        .ok_or("DOCX source-edit corpus target is not a paragraph")?
        .parse::<usize>()?;
    let replacement_text =
        format!("litchi-perf-baseline-docx-semantic-v1-updated-{target_position:05}");
    Ok(Prepared {
        media_ranges: media_ranges(&corpus.archive)?,
        expected_output_sha256: sha256_hex(&expected_output),
        expected_output,
        source_document_xml,
        candidate_document_xml,
        preflight,
        target_position,
        replacement_text,
        corpus,
    })
}

fn snapshot_record(
    stats: Option<&Arc<ReadStats>>,
    scope: &'static str,
    media_ranges: &[Range<u64>],
) -> Result<ReadCounter, Box<dyn Error>> {
    stats
        .map(|stats| {
            stats
                .snapshot()
                .map(|snapshot| ReadCounter::from_snapshot(&snapshot, scope, media_ranges))
                .map_err(Into::into)
        })
        .transpose()
        .map(|record| record.unwrap_or_else(|| ReadCounter::unavailable(scope)))
}

struct TimedOperation {
    materializations: usize,
    commit_changed: bool,
    commit_operations: usize,
    commit_identity_verified: bool,
    budget: BudgetEvidence,
    phase_timings: Option<PhaseTimings>,
}

fn finite_context() -> Result<(Budget, ExecutionContext), Box<dyn Error>> {
    let budget = Budget::root(
        "docx-managed-edit-benchmark",
        Limits::new(
            BUDGET_MEMORY_BYTES,
            BUDGET_INPUT_BYTES,
            BUDGET_OUTPUT_BYTES,
            BUDGET_OBJECTS,
            BUDGET_DEPTH,
            BUDGET_WORK,
        ),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).ok_or("one worker is not representable")?,
        NonZeroUsize::new(1).ok_or("one task slot is not representable")?,
        NonZeroU64::new(MAX_IN_FLIGHT_BYTES).ok_or("in-flight byte limit must be positive")?,
        0,
    )?;
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    Ok((budget, context))
}

fn is_managed_boundary_refusal(error: &(dyn Error + 'static)) -> bool {
    matches!(
        error.downcast_ref::<TransactionError>(),
        Some(TransactionError::Document(litchi_docx::Error::UnsafeEdit {
            operation: "document_snapshot",
            ..
        }))
    )
}

#[inline]
fn phase_start(enabled: bool) -> Option<Instant> {
    enabled.then(Instant::now)
}

#[inline]
fn phase_elapsed(started: Option<Instant>) -> Result<Option<u64>, Box<dyn Error>> {
    started
        .map(|started| {
            u64::try_from(started.elapsed().as_nanos())
                .map_err(|_| "DOCX phase duration does not fit u64 nanoseconds".into())
        })
        .transpose()
}

fn execute_api(
    config: &Config,
    source: Arc<dyn ReadAt>,
    sink: &mut SequentialSink,
    prepared: &Prepared,
    budget: Option<(&Budget, &ExecutionContext, ResourceSnapshot)>,
) -> Result<TimedOperation, Box<dyn Error>> {
    let open_started = phase_start(config.phase_diagnostics);
    let package = match budget {
        Some((_budget, context, _before)) => {
            source_backed::Package::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                ReadLimits::default(),
                SourceCacheLimits::new(CACHE_MAX_BYTES, CACHE_MAX_ENTRIES)?,
                context.clone(),
            )?
        },
        None => source_backed::Package::from_read_at(source)?,
    };
    let cache_before = package.cache_diagnostics();
    let open_ns = phase_elapsed(open_started)?;

    let edit_staging_started = phase_start(config.phase_diagnostics);
    let mut edit = package.edit_document()?;
    edit.replace_paragraph_text(
        Position::new(prepared.target_position),
        &prepared.replacement_text,
    )?;
    let edit_staging_ns = phase_elapsed(edit_staging_started)?;

    let commit_started = phase_start(config.phase_diagnostics);
    let commit = edit.commit()?;
    let commit_ns = phase_elapsed(commit_started)?;

    let diagnostics_xml_identity_started = phase_start(config.phase_diagnostics);
    let commit_changed = commit.patch().changed();
    let commit_operations = commit.diagnostics().operations();
    let commit_identity_verified = commit.patch().source().xml_bytes()
        == prepared.source_document_xml.as_slice()
        && commit.snapshot().xml_bytes() == prepared.candidate_document_xml.as_slice();
    let cache_live = package.cache_diagnostics();
    let materializations = cache_live
        .successful_loads
        .checked_sub(cache_before.successful_loads)
        .ok_or("DOCX source cache successful-load counter moved backwards")?;
    let diagnostics_xml_identity_ns = phase_elapsed(diagnostics_xml_identity_started)?;

    let publication_started = phase_start(config.phase_diagnostics);
    let published = package.publish_document_commit_to_stream(sink, &commit)?;
    let publication_ns = phase_elapsed(publication_started)?;

    let published_snapshot_drop_started = phase_start(config.phase_diagnostics);
    drop(published);
    let published_snapshot_drop_ns = phase_elapsed(published_snapshot_drop_started)?;

    let (budget_evidence, commit_drop_ns) =
        if let Some((budget, _context, resource_before)) = budget {
            let resource_live = ResourceSnapshot::from_budget(budget);

            let commit_drop_started = phase_start(config.phase_diagnostics);
            drop(commit);
            let commit_drop_ns = phase_elapsed(commit_drop_started)?;
            let resource_after_drop = ResourceSnapshot::from_budget(budget);
            let memory_released_to_baseline =
                resource_after_drop.memory_used == resource_before.memory_used;
            let objects_released_to_baseline =
                resource_after_drop.objects_used == resource_before.objects_used;
            let reservation_failures = Some(cache_live.budget_reservation_failures);
            (
                BudgetEvidence {
                    scope: BUDGET_SCOPE,
                    before: resource_before,
                    live: resource_live,
                    after_drop: resource_after_drop,
                    cache_before: CacheBudgetSnapshot::from_diagnostics(cache_before),
                    cache_live: CacheBudgetSnapshot::from_diagnostics(cache_live),
                    memory_released_to_baseline: Some(memory_released_to_baseline),
                    objects_released_to_baseline: Some(objects_released_to_baseline),
                    reservation_failures,
                },
                commit_drop_ns,
            )
        } else {
            let commit_drop_started = phase_start(config.phase_diagnostics);
            drop(commit);
            let commit_drop_ns = phase_elapsed(commit_drop_started)?;
            (BudgetEvidence::unmanaged(), commit_drop_ns)
        };
    if config.api.managed() && !budget_evidence.before.managed {
        return Err("managed-api did not construct a managed execution budget".into());
    }
    Ok(TimedOperation {
        materializations: usize::try_from(materializations)?,
        commit_changed,
        commit_operations,
        commit_identity_verified,
        budget: budget_evidence,
        phase_timings: if config.phase_diagnostics {
            Some(PhaseTimings {
                open_ns: open_ns.ok_or("enabled phase diagnostics did not record open")?,
                edit_staging_ns: edit_staging_ns
                    .ok_or("enabled phase diagnostics did not record edit staging")?,
                commit_ns: commit_ns.ok_or("enabled phase diagnostics did not record commit")?,
                diagnostics_xml_identity_ns: diagnostics_xml_identity_ns
                    .ok_or("enabled phase diagnostics did not record diagnostics")?,
                publication_ns: publication_ns
                    .ok_or("enabled phase diagnostics did not record publication")?,
                published_snapshot_drop_ns: published_snapshot_drop_ns
                    .ok_or("enabled phase diagnostics did not record snapshot drop")?,
                commit_drop_ns: commit_drop_ns
                    .ok_or("enabled phase diagnostics did not record commit drop")?,
            })
        } else {
            None
        },
    })
}

fn run_sample(
    config: &Config,
    prepared: &Prepared,
    file_path: Option<&Path>,
    sample_index: usize,
) -> Result<SampleRecord, Box<dyn Error>> {
    let provider = provider(config, &prepared.corpus.archive, file_path)?;
    let before = provider.source.version()?;
    let maximum = u64::try_from(prepared.expected_output.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("DOCX output ceiling overflows u64")?;
    let mut sink = SequentialSink::new(maximum, SINK_MAX_WRITE_BYTES)?;
    let managed = config.api.managed();
    let context_state = managed.then(finite_context).transpose()?;
    let resource_before = context_state
        .as_ref()
        .map(|(budget, _context)| ResourceSnapshot::from_budget(budget));
    let allocation_region = crate::allocation_metrics::begin();
    let started = Instant::now();
    let operation = execute_api(
        config,
        Arc::clone(&provider.source),
        &mut sink,
        prepared,
        context_state
            .as_ref()
            .zip(resource_before.as_ref())
            .map(|((budget, context), before)| (budget, context, *before)),
    );
    let elapsed = started.elapsed();
    let allocation = allocation_region.finish();
    let operation = match operation {
        Ok(operation) => operation,
        Err(error) if managed && is_managed_boundary_refusal(error.as_ref()) => {
            return Err(format!(
                "managed-api unsupported at current source-backed edit boundary; \
                 refusal is not a timed sample: {error}"
            )
            .into());
        },
        Err(error) => return Err(error),
    };
    let latency_ns = u64::try_from(elapsed.as_nanos())?;
    let phase_diagnostics = operation
        .phase_timings
        .map(|timings| timings.finish(latency_ns))
        .transpose()?;
    let after = provider.source.version()?;
    let source_version_unchanged = before == after;
    let logical = snapshot_record(
        provider.logical_stats.as_ref(),
        "caller-visible logical ReadAt calls",
        &prepared.media_ranges,
    )?;
    let physical = snapshot_record(
        provider.physical_stats.as_ref(),
        "underlying adapter ReadAt calls; transport model only",
        &prepared.media_ranges,
    )?;
    let range_adapter = provider
        .range_adapter
        .as_ref()
        .map(|adapter| adapter.snapshot().map(RangeAdapterRecord::from_snapshot))
        .transpose()?
        .unwrap_or_else(RangeAdapterRecord::unavailable);
    let reads = ReadEvidence {
        logical,
        physical,
        range_adapter,
        media_scope: MEDIA_SCOPE,
    };
    drop(provider);
    let output_exact_bytes = sink.bytes == prepared.expected_output;
    if !output_exact_bytes {
        return Err(format!("DOCX provider output mismatch in sample {sample_index}").into());
    }
    let output_sha256 = sha256_hex(&sink.bytes);
    crate::verify_docx_source_edit_output(&prepared.corpus, &sink.bytes)?;
    let semantic_reopen = true;
    let unchanged_media_preserved = true;
    let cache_load_count = operation.materializations == 1
        && operation.materializations == prepared.preflight.expected_materializations;
    if !source_version_unchanged
        || !cache_load_count
        || !operation.commit_changed
        || operation.commit_operations != 1
        || !operation.commit_identity_verified
    {
        return Err(
            format!("DOCX managed-edit lifecycle oracle failed in sample {sample_index}").into(),
        );
    }
    if managed
        && (operation.budget.reservation_failures != Some(0)
            || operation.budget.memory_released_to_baseline != Some(true)
            || operation.budget.objects_released_to_baseline != Some(true))
    {
        return Err(format!(
            "DOCX managed-edit budget ownership oracle failed in sample {sample_index}"
        )
        .into());
    }
    let oracle = OracleEvidence {
        commit_identity_verified: operation.commit_identity_verified,
        output_exact_bytes,
        semantic_reopen,
        unchanged_media_preserved,
        source_version_unchanged,
        cache_load_count,
        patch_oracles: PatchOracleEvidence {
            scope: "untimed_preflight_commit_patch_oracles",
            replay_forward: prepared.preflight.replay_forward_verified,
            inverse_restores_source: prepared.preflight.inverse_restores_source_verified,
            stale_target_refused: prepared.preflight.stale_target_refusal_verified,
            foreign_source_refused: prepared.preflight.foreign_source_refusal_verified,
        },
    };
    Ok(SampleRecord {
        sample_index,
        api: config.api,
        latency_ns,
        output_bytes: sink.bytes.len(),
        output_sha256,
        output_exact_bytes,
        materializations: operation.materializations,
        commit_changed: operation.commit_changed,
        commit_operations: operation.commit_operations,
        source_version_before: before.into(),
        source_version_after: after.into(),
        source_version_unchanged,
        reads,
        cache: CacheEvidence {
            successful_loads: operation.materializations,
            expected_successful_loads: prepared.preflight.expected_materializations,
            exactly_one_main_part_materialization: cache_load_count,
        },
        sink: sink.record(),
        oracles: oracle,
        budget: operation.budget,
        allocation,
        phase_diagnostics,
    })
}

#[derive(Debug)]
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

fn stage_file(bytes: &[u8]) -> Result<StagedFile, Box<dyn Error>> {
    let id = NEXT_STAGE_ID.fetch_add(1, Ordering::Relaxed);
    let root = std::env::temp_dir().join(format!(
        "litchi-docx-edit-provider-{}-{id}",
        std::process::id()
    ));
    fs::create_dir(&root)?;
    let path = root.join("source.docx");
    let result = (|| -> Result<(), Box<dyn Error>> {
        let mut file = OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(&path)?;
        file.write_all(bytes)?;
        file.sync_all()?;
        Ok(())
    })();
    if let Err(error) = result {
        let _ = fs::remove_file(&path);
        let _ = fs::remove_dir(&root);
        return Err(error);
    }
    Ok(StagedFile { root, path })
}

fn parse_usize(value: &str, flag: &str, maximum: usize) -> Result<usize, Box<dyn Error>> {
    let value = value
        .parse::<usize>()
        .map_err(|_| format!("{flag} requires an unsigned decimal integer"))?;
    if value == 0 || value > maximum {
        return Err(format!("{flag} must be in 1..={maximum}").into());
    }
    Ok(value)
}

fn parse_optional_u64(value: &str, flag: &str) -> Result<u64, Box<dyn Error>> {
    value
        .parse::<u64>()
        .map_err(|_| format!("{flag} requires an unsigned decimal integer").into())
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
    let mut api = None;
    let mut provider = None;
    let mut phase_diagnostics = false;
    let mut max_range_bytes = None;
    let mut delay_us = None;
    let mut transfer_bytes_per_second = None;
    let mut transfer_delay_policy = TransferDelayPolicy::SeparateSleeps;
    let mut samples = None;
    let mut warmup = None;
    let mut source_revision = None;
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("DOCX provider argument must be valid UTF-8")?;
        match flag {
            "--edit-api" | "--api" => {
                if api.is_some() {
                    return Err("--edit-api was specified more than once".into());
                }
                api = Some(match need_value(args, &mut index, flag)?.as_str() {
                    "unmanaged-api" | "unmanaged" => ApiMode::UnmanagedApi,
                    "managed-api" | "managed" => ApiMode::ManagedApi,
                    _ => return Err("--edit-api must be unmanaged-api or managed-api".into()),
                });
            },
            "--provider" => {
                if provider.is_some() {
                    return Err("--provider was specified more than once".into());
                }
                provider = Some(match need_value(args, &mut index, flag)?.as_str() {
                    "owned" | "bytes" => ProviderKind::Owned,
                    "instrumented" | "instrumented-bytes" => ProviderKind::Instrumented,
                    "short" | "short-read" => ProviderKind::Short,
                    "delayed" | "range" => ProviderKind::Delayed,
                    "file" | "file-warm" => ProviderKind::File,
                    _ => {
                        return Err(
                            "--provider must be owned, instrumented, short, delayed, or file"
                                .into(),
                        );
                    },
                });
            },
            "--phase-diagnostics" => {
                if phase_diagnostics {
                    return Err("--phase-diagnostics was specified more than once".into());
                }
                phase_diagnostics = true;
            },
            "--short-range" | "--short-read-bytes" | "--max-range" => {
                if max_range_bytes.is_some() {
                    return Err("range bound was specified more than once".into());
                }
                max_range_bytes = Some(parse_usize(
                    &need_value(args, &mut index, flag)?,
                    flag,
                    MAX_RANGE_BYTES,
                )?);
            },
            "--delay-us" => {
                if delay_us.is_some() {
                    return Err("--delay-us was specified more than once".into());
                }
                let value = parse_optional_u64(&need_value(args, &mut index, flag)?, flag)?;
                if value > MAX_DELAY_US {
                    return Err(format!("--delay-us must be <= {MAX_DELAY_US}").into());
                }
                delay_us = Some(value);
            },
            "--transfer-bytes-per-second" => {
                if transfer_bytes_per_second.is_some() {
                    return Err("--transfer-bytes-per-second was specified more than once".into());
                }
                let value = parse_optional_u64(&need_value(args, &mut index, flag)?, flag)?;
                if !(MIN_TRANSFER_BYTES_PER_SECOND..=MAX_TRANSFER_BYTES_PER_SECOND).contains(&value)
                {
                    return Err(format!(
                        "--transfer-bytes-per-second must be in {MIN_TRANSFER_BYTES_PER_SECOND}..={MAX_TRANSFER_BYTES_PER_SECOND}"
                    )
                    .into());
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
                samples = Some(parse_usize(
                    &need_value(args, &mut index, flag)?,
                    flag,
                    MAX_SAMPLES,
                )?);
            },
            "--warmup" | "--warmups" => {
                if warmup.is_some() {
                    return Err("--warmup was specified more than once".into());
                }
                let value = parse_optional_u64(&need_value(args, &mut index, flag)?, flag)?;
                warmup = Some(usize::try_from(value).map_err(|_| "--warmup overflows usize")?);
                if warmup.unwrap_or(0) > MAX_WARMUP {
                    return Err(format!("--warmup must be <= {MAX_WARMUP}").into());
                }
            },
            "--source-revision" => {
                if source_revision.is_some() {
                    return Err("--source-revision was specified more than once".into());
                }
                let value = need_value(args, &mut index, flag)?;
                if value.len() != 40
                    || !value
                        .bytes()
                        .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
                {
                    return Err(
                        "--source-revision must be exactly 40 hexadecimal characters".into(),
                    );
                }
                source_revision = Some(value);
            },
            "--output" | "--json" => {
                if output.is_some() {
                    return Err("--output was specified more than once".into());
                }
                let value = need_value(args, &mut index, flag)?;
                if value.is_empty() || value == "-" {
                    return Err("--output must be a new regular file path".into());
                }
                output = Some(PathBuf::from(value));
            },
            "--trace-ranges" => {},
            other => return Err(format!("unknown argument: {other}").into()),
        }
        index += 1;
    }
    let api = api.unwrap_or(ApiMode::UnmanagedApi);
    let provider = provider.ok_or("--provider is required")?;
    let samples = samples.ok_or("--samples is required")?;
    let source_revision = source_revision.ok_or("--source-revision is required")?;
    let output = output.ok_or("--output is required")?;
    let (max_range_bytes, delay_us, transfer_bytes_per_second, transfer_delay_policy) =
        match provider {
            ProviderKind::Short => {
                if delay_us.is_some()
                    || transfer_bytes_per_second.is_some()
                    || transfer_delay_policy != TransferDelayPolicy::SeparateSleeps
                {
                    return Err("delay and transfer controls require --provider delayed".into());
                }
                (
                    Some(max_range_bytes.unwrap_or(DEFAULT_SHORT_RANGE_BYTES)),
                    None,
                    None,
                    TransferDelayPolicy::SeparateSleeps,
                )
            },
            ProviderKind::Delayed => (
                Some(max_range_bytes.unwrap_or(DEFAULT_DELAY_RANGE_BYTES)),
                Some(delay_us.unwrap_or(DEFAULT_DELAY_US)),
                if delay_us == Some(0) && transfer_bytes_per_second.is_none() {
                    if transfer_delay_policy != TransferDelayPolicy::SeparateSleeps {
                        return Err("minimum-service requires --transfer-bytes-per-second".into());
                    }
                    None
                } else {
                    Some(transfer_bytes_per_second.unwrap_or_else(|| {
                        NonZeroU64::new(DEFAULT_TRANSFER_BYTES_PER_SECOND).unwrap()
                    }))
                },
                transfer_delay_policy,
            ),
            ProviderKind::Owned | ProviderKind::Instrumented | ProviderKind::File => {
                if max_range_bytes.is_some()
                    || delay_us.is_some()
                    || transfer_bytes_per_second.is_some()
                    || transfer_delay_policy != TransferDelayPolicy::SeparateSleeps
                {
                    return Err("range controls require --provider short or delayed".into());
                }
                (None, None, None, TransferDelayPolicy::SeparateSleeps)
            },
        };
    Ok(Config {
        api,
        provider,
        phase_diagnostics,
        max_range_bytes,
        delay_us,
        transfer_bytes_per_second,
        transfer_delay_policy,
        samples,
        warmup: warmup.unwrap_or(0),
        source_revision,
        output,
    })
}

fn current_executable_identity() -> Result<(String, u64), Box<dyn Error>> {
    let path = std::env::current_exe()?;
    let length = fs::metadata(&path)?.len();
    let mut file = OpenOptions::new().read(true).open(&path)?;
    let mut digest = Sha256::new();
    let mut buffer = [0_u8; 64 * 1024];
    let mut total = 0_u64;
    loop {
        let read = file.read(&mut buffer)?;
        if read == 0 {
            break;
        }
        digest.update(&buffer[..read]);
        total = total
            .checked_add(u64::try_from(read)?)
            .ok_or("executable byte count overflows u64")?;
    }
    if total != length {
        return Err("executable changed while its identity was being captured".into());
    }
    let mut hash = String::with_capacity(64);
    for byte in digest.finalize() {
        let _ = write!(hash, "{byte:02x}");
    }
    Ok((hash, length))
}

fn usage() -> &'static str {
    "docx-managed-edit [--edit-api <unmanaged-api|managed-api>] --provider <owned|instrumented|short|delayed|file> [--phase-diagnostics] [--short-range N|--short-read-bytes N|--max-range N] [--delay-us N --transfer-bytes-per-second N --transfer-delay-policy separate-sleeps|minimum-service] --samples N --warmup N --source-revision <40 hex chars> --output PATH"
}

/// Runs the benchmark from arguments following the `docx-edit-provider`
/// selector.
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
    let prepared = prepare()?;
    let staged = (config.provider == ProviderKind::File)
        .then(|| stage_file(&prepared.corpus.archive))
        .transpose()?;
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or("warmup plus samples overflows usize")?;
    let mut rows = Vec::with_capacity(config.samples);
    for iteration in 0..total {
        let row = run_sample(
            &config,
            &prepared,
            staged.as_ref().map(|file| file.path.as_path()),
            iteration,
        )?;
        if iteration >= config.warmup {
            rows.push(SampleRecord {
                sample_index: iteration - config.warmup,
                ..row
            });
        }
    }
    if rows.len() != config.samples
        || rows
            .iter()
            .any(|row| row.output_sha256 != prepared.expected_output_sha256)
    {
        return Err("DOCX provider measured output digest vector is not stable".into());
    }
    let (binary_sha256, binary_bytes) = current_executable_identity()?;
    let provider_config = config.provider_config();
    let source_revision = config.source_revision.clone();
    let report = Report {
        schema: SCHEMA,
        version: REPORT_VERSION,
        case_name: CASE_NAME,
        benchmark: "one DOCX paragraph replacement through managed or unmanaged source-backed open/edit/commit/sequential publication",
        api: config.api,
        api_name: config.api.name(),
        provider_scope: PROVIDER_SCOPE,
        timing_scope: TIMING_SCOPE,
        setup_scope: SETUP_SCOPE,
        allocation_scope: ALLOCATION_SCOPE,
        physical_scope: PHYSICAL_SCOPE,
        range_scope: RANGE_SCOPE,
        zero_length_scope: ZERO_LENGTH_SCOPE,
        file_scope: FILE_SCOPE,
        budget_scope: BUDGET_SCOPE,
        corpus_version: CORPUS_VERSION,
        corpus_generator: CORPUS_GENERATOR,
        source_archive_sha256: sha256_hex(&prepared.corpus.archive),
        source_archive_bytes: prepared.corpus.archive.len(),
        source_bytes: prepared.corpus.archive.len(),
        source_sha256: sha256_hex(&prepared.corpus.archive),
        expected_text_bytes: EXPECTED_TEXT_BYTES,
        expected_text_sha256: EXPECTED_TEXT_SHA256,
        expected_text_scope: EXPECTED_TEXT_SCOPE,
        expected_archive_members: prepared.corpus.manifest.archive_member_count,
        expected_output_sha256: prepared.expected_output_sha256.clone(),
        expected_output_bytes: prepared.expected_output.len(),
        source_revision: source_revision.clone(),
        requested_source_revision: source_revision,
        limits: limits_record(config.api),
        binary_sha256,
        binary_bytes,
        current_exe: std::env::current_exe()?.to_string_lossy().into_owned(),
        provider: provider_config,
        corpus: prepared.corpus.manifest.clone(),
        preflight: prepared.preflight,
        warmup: config.warmup,
        samples: config.samples,
        allocator: crate::allocation_metrics::allocator_identity(),
        instrumentation: crate::allocation_metrics::instrumentation_identity(),
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

    fn args(values: &[&str]) -> Vec<OsString> {
        values.iter().map(OsString::from).collect()
    }

    #[test]
    fn parser_accepts_each_provider_and_defaults_transport_bounds() {
        for provider in ["owned", "instrumented", "short", "delayed", "file"] {
            let config = parse_config(&args(&[
                "--provider",
                provider,
                "--samples",
                "1",
                "--warmup",
                "0",
                "--source-revision",
                "0123456789012345678901234567890123456789",
                "--output",
                "result.json",
            ]))
            .expect("valid provider configuration");
            assert_eq!(config.provider.name(), provider);
        }
        let default = parse_config(&args(&[
            "--provider",
            "owned",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789012345678901234567890123456789",
            "--output",
            "result.json",
        ]))
        .expect("default phase diagnostics configuration");
        assert!(!default.phase_diagnostics);
        let enabled = parse_config(&args(&[
            "--provider",
            "owned",
            "--phase-diagnostics",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789012345678901234567890123456789",
            "--output",
            "result.json",
        ]))
        .expect("opt-in phase diagnostics configuration");
        assert!(enabled.phase_diagnostics);
        let short = parse_config(&args(&[
            "--provider",
            "short",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789012345678901234567890123456789",
            "--output",
            "result.json",
        ]))
        .expect("short defaults");
        assert_eq!(short.max_range_bytes, Some(DEFAULT_SHORT_RANGE_BYTES));
        let delayed = parse_config(&args(&[
            "--provider",
            "delayed",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789012345678901234567890123456789",
            "--output",
            "result.json",
        ]))
        .expect("delayed defaults");
        assert_eq!(delayed.max_range_bytes, Some(DEFAULT_DELAY_RANGE_BYTES));
        assert_eq!(delayed.delay_us, Some(DEFAULT_DELAY_US));

        let range_zero = parse_config(&args(&[
            "--provider",
            "delayed",
            "--max-range",
            "65536",
            "--delay-us",
            "0",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789012345678901234567890123456789",
            "--output",
            "result.json",
        ]))
        .expect("range-zero controls");
        assert_eq!(range_zero.delay_us, Some(0));
        assert_eq!(range_zero.transfer_bytes_per_second, None);
        assert_eq!(range_zero.provider_config().kind, "range-control");
    }

    #[test]
    fn parser_rejects_range_controls_on_non_range_providers() {
        assert!(
            parse_config(&args(&[
                "--provider",
                "owned",
                "--max-range",
                "128",
                "--samples",
                "1",
                "--source-revision",
                "0123456789012345678901234567890123456789",
                "--output",
                "result.json",
            ]))
            .is_err()
        );
        assert!(
            parse_config(&args(&[
                "--provider",
                "short",
                "--delay-us",
                "1",
                "--samples",
                "1",
                "--source-revision",
                "0123456789012345678901234567890123456789",
                "--output",
                "result.json",
            ]))
            .is_err()
        );
    }

    #[test]
    fn parser_accepts_capture_aliases_and_preserves_short_bound() {
        let config = parse_config(&args(&[
            "--provider",
            "short-read",
            "--short-read-bytes",
            "4096",
            "--trace-ranges",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789012345678901234567890123456789",
            "--output",
            "result.json",
        ]))
        .expect("capture short-read aliases");
        assert_eq!(config.provider, ProviderKind::Short);
        assert_eq!(config.max_range_bytes, Some(4096));

        let file = parse_config(&args(&[
            "--provider",
            "file-warm",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789012345678901234567890123456789",
            "--output",
            "result.json",
        ]))
        .expect("capture file-warm alias");
        assert_eq!(file.provider, ProviderKind::File);
    }

    #[test]
    fn parser_rejects_malformed_source_revision() {
        assert!(
            parse_config(&args(&[
                "--provider",
                "owned",
                "--samples",
                "1",
                "--warmup",
                "0",
                "--source-revision",
                "not-a-revision",
                "--output",
                "result.json",
            ]))
            .is_err()
        );
    }

    #[test]
    fn phase_diagnostics_reports_named_intervals_and_rejects_nonconservation() {
        let timings = PhaseTimings {
            open_ns: 10,
            edit_staging_ns: 10,
            commit_ns: 10,
            diagnostics_xml_identity_ns: 10,
            publication_ns: 10,
            published_snapshot_drop_ns: 10,
            commit_drop_ns: 10,
        };
        let diagnostics = timings.finish(100).expect("conserved phase intervals");
        assert_eq!(diagnostics.phase_sum_ns, 70);
        assert_eq!(diagnostics.lifecycle_residual_ns, 30);
        assert!(timings.finish(69).is_err());
        assert!(diagnostics.instrumentation_overhead_ns.is_none());
    }

    #[test]
    fn counting_source_rejects_impossible_range_trace() {
        let stats = Arc::new(ReadStats::new().expect("range stats"));
        assert!(stats.record(u64::MAX, 2, 2).is_err());
        assert!(stats.snapshot().is_err());
    }

    #[test]
    fn limits_report_uses_package_defaults_without_claiming_managed_budget() {
        let report = limits_record(ApiMode::UnmanagedApi);
        let read_limits = ReadLimits::default();
        let cache_limits = SourceCacheLimits::default();
        assert_eq!(
            report.read_limits.max_input_bytes,
            read_limits.max_input_bytes()
        );
        assert_eq!(report.cache_max_bytes, cache_limits.max_bytes() as u64);
        assert_eq!(report.cache_max_entries, cache_limits.max_entries() as u64);
        assert!(!report.resource_budget.managed);
        assert!(report.resource_budget.memory_bytes.is_none());
        assert!(report.resource_budget.input_bytes.is_none());
        assert!(report.resource_budget.output_bytes.is_none());
        assert!(report.resource_budget.objects.is_none());
        assert!(report.resource_budget.depth.is_none());
        assert!(report.resource_budget.work.is_none());
        assert!(!report.resource_budget.mode_reason.is_empty());
    }

    #[test]
    fn counting_source_records_zero_length_calls_and_delegates() {
        struct EmptyProbe {
            observed: Arc<AtomicBool>,
        }

        impl ReadAt for EmptyProbe {
            fn len(&self) -> io::Result<u64> {
                Ok(3)
            }

            fn read_at(&self, _offset: u64, output: &mut [u8]) -> io::Result<usize> {
                assert!(output.is_empty());
                self.observed.store(true, Ordering::Release);
                Ok(0)
            }

            fn version(&self) -> io::Result<SourceVersion> {
                Ok(SourceVersion::new(1, 0))
            }
        }

        let stats = Arc::new(ReadStats::new().expect("range stats"));
        let observed = Arc::new(AtomicBool::new(false));
        let inner: Arc<dyn ReadAt> = Arc::new(EmptyProbe {
            observed: Arc::clone(&observed),
        });
        let source = CountingReadAt::new(inner, Arc::clone(&stats));
        assert_eq!(source.read_at(1, &mut []).expect("empty read"), 0);
        assert!(observed.load(Ordering::Acquire));
        let snapshot = stats.snapshot().expect("counter snapshot");
        assert_eq!(snapshot.calls, 0);
        assert_eq!(snapshot.empty_calls, 1);
    }

    #[test]
    fn counting_source_preserves_source_version_identity() {
        let stats = Arc::new(ReadStats::new().expect("range stats"));
        let inner: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(vec![1, 2, 3]));
        let expected = inner.version().expect("source version");
        let source = CountingReadAt::new(inner, Arc::clone(&stats));
        assert_eq!(source.version().expect("wrapped source version"), expected);
    }

    #[test]
    fn sequential_sink_preserves_all_writes_without_seek() {
        let mut sink = SequentialSink::new(8, 8).expect("sink");
        sink.write_all(b"abcd").expect("first write");
        sink.write_all(b"efgh").expect("second write");
        assert_eq!(sink.bytes, b"abcdefgh");
        assert_eq!(sink.record().write_calls, 2);
        assert!(sink.write(b"!").is_err());
    }

    #[test]
    fn real_owned_and_short_lifecycles_keep_the_edit_oracles() {
        let prepared = prepare().expect("deterministic DOCX preflight");
        let owned = Config {
            api: ApiMode::UnmanagedApi,
            provider: ProviderKind::Owned,
            phase_diagnostics: false,
            max_range_bytes: None,
            delay_us: None,
            transfer_bytes_per_second: None,
            transfer_delay_policy: TransferDelayPolicy::SeparateSleeps,
            samples: 1,
            warmup: 0,
            source_revision: "0123456789012345678901234567890123456789".to_owned(),
            output: PathBuf::from("unused-owned.json"),
        };
        let owned_row = run_sample(&owned, &prepared, None, 0).expect("owned lifecycle");
        assert!(owned_row.output_exact_bytes);
        assert_eq!(owned_row.materializations, 1);
        assert!(owned_row.commit_changed);
        assert_eq!(owned_row.commit_operations, 1);
        assert!(owned_row.source_version_unchanged);
        assert!(owned_row.cache.exactly_one_main_part_materialization);
        assert!(owned_row.oracles.commit_identity_verified);
        assert!(owned_row.oracles.semantic_reopen);
        assert!(owned_row.phase_diagnostics.is_none());
        let serialized_owned = serde_json::to_value(&owned_row).expect("serialize default row");
        assert!(serialized_owned.get("phase_diagnostics").is_none());

        let short = Config {
            api: ApiMode::UnmanagedApi,
            provider: ProviderKind::Short,
            phase_diagnostics: true,
            max_range_bytes: Some(4096),
            delay_us: None,
            transfer_bytes_per_second: None,
            transfer_delay_policy: TransferDelayPolicy::SeparateSleeps,
            samples: 1,
            warmup: 0,
            source_revision: "0123456789012345678901234567890123456789".to_owned(),
            output: PathBuf::from("unused-short.json"),
        };
        let short_row = run_sample(&short, &prepared, None, 0).expect("short lifecycle");
        assert!(short_row.output_exact_bytes);
        assert_eq!(short_row.materializations, 1);
        assert!(short_row.commit_changed);
        assert_eq!(short_row.commit_operations, 1);
        assert!(short_row.source_version_unchanged);
        assert_eq!(short_row.reads.logical.availability, "available");
        assert!(short_row.oracles.patch_oracles.foreign_source_refused);
        let phases = short_row
            .phase_diagnostics
            .expect("enabled lifecycle phase diagnostics");
        assert_eq!(
            phases
                .phase_sum_ns
                .checked_add(phases.lifecycle_residual_ns)
                .expect("phase interval sum"),
            short_row.latency_ns
        );
    }
}
