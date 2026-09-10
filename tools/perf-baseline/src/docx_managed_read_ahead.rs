//! Managed DOCX source-read policy measurements.
//!
//! This is the small production follow-up to the unmanaged 0492 pilot.  Each
//! row opens a fresh managed DOCX owner with an explicit finite
//! [`litchi_core::ExecutionContext`].  The control uses the exact source-read
//! policy and the candidate opts into one 4 KiB forward-start window.  The
//! simulated range source is retained only as a deterministic transport
//! model; `CountingReadAt` below it records the physical calls separately
//! from the package's own diagnostics.

#![allow(clippy::module_name_repetitions)]

use std::{
    error::Error,
    ffi::OsString,
    fmt::Write as _,
    fs::OpenOptions,
    io::{self, Write},
    num::{NonZeroU64, NonZeroUsize},
    ops::Range,
    path::PathBuf,
    sync::{
        Arc, Mutex,
        atomic::{AtomicBool, AtomicU64, Ordering},
    },
    time::{Duration, Instant},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    Resource, SourceVersion,
};
use litchi_docx::{ReadLimits, source_backed};
use litchi_opc::SourceCacheLimits;
use serde::Serialize;
use sha2::{Digest, Sha256};

use crate::{allocation_metrics, pptx_range_source};

const SCHEMA: &str = "docx_managed_read_ahead_v1";
const CASE: &str = "docx_managed_source_open_document_extract_text";
const CORPUS_VERSION: &str = "source-edit-media-v1";
const CORPUS_GENERATOR: &str = "litchi-docx-source-edit-media-v1";
const CORPUS_ARCHIVE_BYTES: usize = 16_793_036;
const CORPUS_ARCHIVE_SHA256: &str =
    "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4";
const EXPECTED_TEXT_BYTES: usize = 10_000;
const EXPECTED_TEXT_SHA256: &str =
    "ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af";
const EXPECTED_ARCHIVE_MEMBERS: usize = 20;
const EXPECTED_MEDIA_MEMBERS: usize = 8;
const MAX_TRACKED_RANGES: usize = 4_096;
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 1_000;
const MAX_RANGE_BYTES: usize = 1_048_576;
const MAX_DELAY_US: u64 = 100_000;
const MAX_TRANSFER_BYTES_PER_SECOND: u64 = 1_099_511_627_776;
const MAX_WINDOW_BYTES: usize = 65_536;
const WINDOW_BYTES: usize = 4_096;
const CACHE_MAX_BYTES: usize = 8 * 1024 * 1024;
const CACHE_MAX_ENTRIES: usize = 128;

// These limits are deliberately finite and shared by every measured arm.  A
// package source is 16 MiB, but the managed operation reads only bounded ZIP
// metadata and the 10 KiB document payload.  The budget is large enough for
// the explicitly bounded cache and parser while remaining a real admission
// policy rather than an unlimited sentinel.
const BUDGET_MEMORY_BYTES: u64 = 64 * 1024 * 1024;
const BUDGET_INPUT_BYTES: u64 = 64 * 1024 * 1024;
const BUDGET_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
const BUDGET_OBJECTS: u64 = 1_000_000;
const BUDGET_DEPTH: u64 = 1_024;
const BUDGET_WORK: u64 = 2 * 1024 * 1024 * 1024;
const MAX_IN_FLIGHT_BYTES: u64 = 64 * 1024 * 1024;

const TIMING_SCOPE: &str = "SourceBackedPackage managed open (including read-ahead window construction) + cache snapshot after open + document + extract_text + source-read and cache snapshots after text + package/document drop; returned text remains live after the clock; hashing, oracle comparison, source-version checks, physical traces, and post-drop budget checks are outside";
const SETUP_SCOPE: &str = "deterministic corpus generation, source/provider construction, bounded physical-trace capacity reservation, range transport construction, finite limits and ExecutionContext construction are outside the operation clock; the production read-ahead window is constructed inside the clock by the managed DOCX owner";
const ALLOCATION_SCOPE: &str = "operation-scoped global-system-allocator region surrounds the timing scope and includes managed package open, read-ahead window allocation, document materialization, text extraction, cache/source-read diagnostics, and package/document drop; normal binaries report unavailable and allocator binaries report the separate region sample";
const PROVIDER_SCOPE: &str = "managed litchi-docx source-backed owner through litchi-opc; the source adapter is a deterministic in-memory range transport and is not disk or network I/O evidence";
const PHYSICAL_SCOPE: &str = "CountingReadAt calls below PptxRangeSource and below the production SourceReader read-ahead policy; requested and returned ranges are physical transport observations for this synthetic adapter";
const RANGE_SCOPE: &str = "PptxRangeSource logical adapter counters; fixed delay and transfer pacing describe the configured synthetic transport, not a filesystem or network service-level measurement";

const EXPECTED_EXACT_RANGES: &[(u64, usize)] = &[
    (16_793_014, 22),
    (16_791_709, 46),
    (16_791_709, 1_305),
    (0, 30),
    (49, 363),
    (412, 16),
    (412, 16),
    (428, 30),
    (469, 234),
    (703, 16),
    (703, 16),
    (2_907, 30),
    (2_965, 324),
    (3_289, 16),
    (3_289, 16),
    (1_420, 30),
    (1_467, 1_424),
    (2_891, 16),
    (2_891, 16),
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Policy {
    Exact,
    ForwardStart(usize),
}

impl Policy {
    const fn name(self) -> &'static str {
        match self {
            Self::Exact => "exact",
            Self::ForwardStart(_) => "forward_start",
        }
    }

    const fn window_bytes(self) -> Option<usize> {
        match self {
            Self::Exact => None,
            Self::ForwardStart(window) => Some(window),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Config {
    policy: Policy,
    max_range_bytes: usize,
    delay_us: u64,
    transfer_bytes_per_second: Option<NonZeroU64>,
    transfer_delay_policy: pptx_range_source::TransferDelayPolicy,
    samples: usize,
    warmup: usize,
    source_revision: [u8; 40],
    output: PathBuf,
}

#[derive(Clone, Debug)]
struct Corpus {
    archive: Arc<Vec<u8>>,
    expected_text: String,
    media_ranges: Vec<Range<u64>>,
    manifest: crate::CorpusManifest,
}

#[derive(Clone, Debug, Serialize)]
struct PolicyRecord {
    name: &'static str,
    enabled: bool,
    configured_window_bytes: Option<usize>,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct SourceVersionRecord {
    id: u64,
    revision: u64,
}

impl From<SourceVersion> for SourceVersionRecord {
    fn from(value: SourceVersion) -> Self {
        Self {
            id: value.id(),
            revision: value.revision(),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct ObservedRange {
    offset: u64,
    requested: u64,
    returned: u64,
}

#[derive(Clone, Debug, Serialize)]
struct PhysicalSnapshot {
    scope: &'static str,
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    short_reads: u64,
    ranges: Vec<ObservedRange>,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct BudgetSnapshot {
    memory_before: u64,
    memory_live: u64,
    memory_after_drop: u64,
    memory_released: u64,
    input_bytes_before: u64,
    input_bytes_after_drop: u64,
    input_bytes_delta: u64,
    objects_before: u64,
    objects_live: u64,
    objects_after_drop: u64,
    objects_released: u64,
    managed: bool,
    memory_released_to_baseline: bool,
    input_bytes_match_physical_returned: bool,
}

#[derive(Clone, Debug, Serialize)]
struct SourceReadRecord {
    enabled: bool,
    configured_window_bytes: usize,
    retained_window_bytes: usize,
    requests: u64,
    hits: u64,
    misses: u64,
    fills: u64,
    requested_bytes: u64,
    returned_bytes: u64,
}

#[derive(Clone, Debug, Serialize)]
struct AllocationRecord {
    #[serde(skip_serializing_if = "Option::is_none")]
    sample: Option<allocation_metrics::Sample>,
    #[serde(skip_serializing_if = "Option::is_none")]
    region_peak_increment_bytes: Option<u64>,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct CacheRecord {
    /// Successful payload loads observed immediately after package open.
    open_successful_loads: u64,
    /// Cumulative successful payload loads after the document text read.
    successful_loads: u64,
    failed_loads: u64,
    retained_bytes: usize,
    retained_entries: usize,
    budget_managed: bool,
}

#[derive(Clone, Debug, Serialize)]
struct SampleRecord {
    sample_index: usize,
    elapsed_ns: u64,
    actual_text_verified: bool,
    actual_text_bytes: usize,
    actual_text_sha256: String,
    source_version_before: SourceVersionRecord,
    source_version_after: SourceVersionRecord,
    source_version_unchanged: bool,
    policy: PolicyRecord,
    source_read: Option<SourceReadRecord>,
    cache: CacheRecord,
    physical: PhysicalSnapshot,
    transport: pptx_range_source::PptxRangeSourceSnapshot,
    budget: BudgetSnapshot,
    allocation: AllocationRecord,
}

#[derive(Clone, Debug, Serialize)]
struct LimitRecord {
    read_limits: ReadLimitsRecord,
    cache_max_bytes: usize,
    cache_max_entries: usize,
    budget_memory_bytes: u64,
    budget_input_bytes: u64,
    budget_output_bytes: u64,
    budget_objects: u64,
    budget_depth: u64,
    budget_work: u64,
    execution_workers: usize,
    execution_max_in_flight_tasks: usize,
    execution_max_in_flight_bytes: u64,
    execution_min_parallel_bytes: u64,
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

#[derive(Clone, Debug, Serialize)]
struct ProviderRecord {
    name: &'static str,
    max_range_bytes: usize,
    delay_us: u64,
    transfer_bytes_per_second: Option<u64>,
    transfer_delay_policy: pptx_range_source::TransferDelayPolicy,
    physical_scope: &'static str,
    range_scope: &'static str,
}

#[derive(Clone, Debug, Serialize)]
struct Report {
    schema: &'static str,
    version: u32,
    case_name: &'static str,
    provider_scope: &'static str,
    timing_scope: &'static str,
    setup_scope: &'static str,
    allocation_scope: &'static str,
    corpus_version: &'static str,
    corpus_generator: &'static str,
    corpus: crate::CorpusManifest,
    source_bytes: usize,
    source_sha256: String,
    expected_text_bytes: usize,
    expected_text_sha256: &'static str,
    expected_archive_members: usize,
    media_ranges: Vec<Range<u64>>,
    expected_exact_ranges: Vec<ObservedRange>,
    expected_exact_physical_calls: u64,
    expected_exact_physical_bytes: u64,
    requested_source_revision: String,
    limits: LimitRecord,
    provider: ProviderRecord,
    policy: PolicyRecord,
    warmup: usize,
    samples: usize,
    rows: Vec<SampleRecord>,
}

#[derive(Debug)]
struct CountingState {
    calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    short_reads: AtomicU64,
    failed: AtomicBool,
    ranges_overflowed: AtomicBool,
    ranges: Mutex<Vec<ObservedRange>>,
}

impl Default for CountingState {
    fn default() -> Self {
        Self {
            calls: AtomicU64::new(0),
            requested_bytes: AtomicU64::new(0),
            returned_bytes: AtomicU64::new(0),
            short_reads: AtomicU64::new(0),
            failed: AtomicBool::new(false),
            ranges_overflowed: AtomicBool::new(false),
            // Reserve the bounded trace outside the measured operation so
            // policy arms do not pay observer Vec growth at different points
            // in the timed lifecycle.
            ranges: Mutex::new(Vec::with_capacity(MAX_TRACKED_RANGES)),
        }
    }
}

impl CountingState {
    fn add(counter: &AtomicU64, amount: u64, failed: &AtomicBool) -> io::Result<()> {
        counter
            .fetch_update(Ordering::AcqRel, Ordering::Acquire, |value| {
                value.checked_add(amount)
            })
            .map(|_| ())
            .map_err(|_| {
                failed.store(true, Ordering::Release);
                io::Error::other("managed benchmark physical counter overflow")
            })
    }

    fn record(&self, offset: u64, requested: usize, returned: usize) -> io::Result<()> {
        if returned > requested {
            self.failed.store(true, Ordering::Release);
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "physical source returned more bytes than requested",
            ));
        }
        let requested = u64::try_from(requested)
            .map_err(|_| io::Error::other("physical request length does not fit u64"))?;
        let returned = u64::try_from(returned)
            .map_err(|_| io::Error::other("physical returned length does not fit u64"))?;
        offset
            .checked_add(requested)
            .ok_or_else(|| io::Error::other("physical range overflows u64"))?;
        Self::add(&self.calls, 1, &self.failed)?;
        Self::add(&self.requested_bytes, requested, &self.failed)?;
        Self::add(&self.returned_bytes, returned, &self.failed)?;
        if returned < requested {
            Self::add(&self.short_reads, 1, &self.failed)?;
        }
        let mut ranges = self
            .ranges
            .lock()
            .map_err(|_| io::Error::other("physical trace lock poisoned"))?;
        if ranges.len() >= MAX_TRACKED_RANGES {
            self.ranges_overflowed.store(true, Ordering::Release);
            return Err(io::Error::other("physical trace sample bound exceeded"));
        }
        ranges.push(ObservedRange {
            offset,
            requested,
            returned,
        });
        Ok(())
    }

    fn snapshot(&self) -> io::Result<PhysicalSnapshot> {
        if self.failed.load(Ordering::Acquire) {
            return Err(io::Error::other("physical benchmark counters unavailable"));
        }
        if self.ranges_overflowed.load(Ordering::Acquire) {
            return Err(io::Error::other("physical benchmark trace overflowed"));
        }
        let ranges = self
            .ranges
            .lock()
            .map_err(|_| io::Error::other("physical trace lock poisoned"))?
            .clone();
        let calls = self.calls.load(Ordering::Acquire);
        if ranges.len() as u64 != calls {
            return Err(io::Error::other("physical trace and call counter disagree"));
        }
        Ok(PhysicalSnapshot {
            scope: PHYSICAL_SCOPE,
            calls,
            requested_bytes: self.requested_bytes.load(Ordering::Acquire),
            returned_bytes: self.returned_bytes.load(Ordering::Acquire),
            short_reads: self.short_reads.load(Ordering::Acquire),
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
        self.state.record(offset, output.len(), returned)?;
        Ok(returned)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

struct Provider {
    source: Arc<dyn ReadAt>,
    range: Arc<pptx_range_source::PptxRangeSource>,
    physical: Arc<CountingState>,
    version: SourceVersion,
}

fn provider(config: &Config, corpus: &Corpus) -> Result<Provider, Box<dyn Error>> {
    let physical = Arc::new(CountingState::default());
    let owned: Arc<dyn ReadAt> = Arc::new(OwnedSource::from_arc(Arc::clone(&corpus.archive)));
    let counted: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(owned, Arc::clone(&physical)));
    let mut range_config = pptx_range_source::PptxRangeSourceConfig::new(
        Some(config.max_range_bytes),
        Some(Duration::from_micros(config.delay_us)),
    );
    range_config.transfer_bytes_per_second = config.transfer_bytes_per_second;
    range_config.transfer_delay_policy = config.transfer_delay_policy;
    let range = Arc::new(pptx_range_source::PptxRangeSource::new(
        counted,
        range_config,
    ));
    let source: Arc<dyn ReadAt> = range.clone();
    let version = source.version()?;
    Ok(Provider {
        source,
        range,
        physical,
        version,
    })
}

fn finite_context() -> Result<(Budget, ExecutionContext), Box<dyn Error>> {
    let budget = Budget::root(
        "docx-managed-read-ahead-benchmark",
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

fn source_policy(policy: Policy) -> Result<source_backed::SourceReadPolicy, Box<dyn Error>> {
    Ok(match policy {
        Policy::Exact => source_backed::SourceReadPolicy::exact(),
        Policy::ForwardStart(window) => source_backed::SourceReadPolicy::forward_start(window)?,
    })
}

fn run_sample(
    config: &Config,
    corpus: &Corpus,
    sample_index: usize,
    limits: ReadLimits,
    cache_limits: SourceCacheLimits,
) -> Result<SampleRecord, Box<dyn Error>> {
    let provider = provider(config, corpus)?;
    let (budget, context) = finite_context()?;
    let memory_before = budget.used(Resource::Memory);
    let input_before = budget.used(Resource::InputBytes);
    let objects_before = budget.used(Resource::Objects);
    let before = provider.version;
    let allocation_region = allocation_metrics::begin();
    let started = Instant::now();
    let (actual_text, source_read, cache, memory_live, objects_live) = {
        let package = source_backed::Package::from_read_at_with_limits_and_cache_limits_and_source_read_policy_and_execution_context(
            Arc::clone(&provider.source),
            limits,
            cache_limits,
            source_policy(config.policy)?,
            context,
        )?;
        let open_cache = package.cache_diagnostics();
        let document = package.document()?;
        let text = document.extract_text()?;
        let source_read = package
            .source_read_diagnostics()?
            .map(SourceReadRecord::from_diagnostics);
        let loaded_cache = package.cache_diagnostics();
        let cache = CacheRecord {
            open_successful_loads: open_cache.successful_loads,
            successful_loads: loaded_cache.successful_loads,
            failed_loads: loaded_cache.failed_loads,
            retained_bytes: loaded_cache.retained_bytes,
            retained_entries: loaded_cache.retained_entries,
            budget_managed: loaded_cache.budget_managed,
        };
        let memory_live = budget.used(Resource::Memory);
        let objects_live = budget.used(Resource::Objects);
        std::hint::black_box(&text);
        drop(document);
        drop(package);
        (text, source_read, cache, memory_live, objects_live)
    };
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos())
        .map_err(|_| "managed DOCX operation duration does not fit u64 nanoseconds")?;
    let allocation = allocation_region.finish();

    let after = provider.source.version()?;
    let physical = provider.physical.snapshot()?;
    let transport = provider.range.snapshot()?;
    let memory_after_drop = budget.used(Resource::Memory);
    let input_after_drop = budget.used(Resource::InputBytes);
    let objects_after_drop = budget.used(Resource::Objects);
    let input_bytes_delta = input_after_drop
        .checked_sub(input_before)
        .ok_or("managed InputBytes usage moved backwards")?;
    let memory_released = memory_live
        .checked_sub(memory_after_drop)
        .ok_or("managed Memory usage moved backwards after package drop")?;
    let objects_released = objects_live
        .checked_sub(objects_after_drop)
        .ok_or("managed Objects usage moved backwards after package drop")?;
    if memory_after_drop != 0 || memory_after_drop != memory_before {
        return Err(format!(
            "managed Memory usage did not return to zero and its pre-sample value in sample {sample_index}: before={memory_before}, after_drop={memory_after_drop}"
        )
        .into());
    }
    if objects_after_drop != 0 || objects_after_drop != objects_before {
        return Err(format!(
            "managed Objects usage did not return to zero and its pre-sample value in sample {sample_index}: before={objects_before}, after_drop={objects_after_drop}"
        )
        .into());
    }
    if input_bytes_delta != physical.returned_bytes {
        return Err(format!(
            "managed InputBytes usage differs from accepted physical bytes in sample {sample_index}: delta={input_bytes_delta}, physical_returned={}",
            physical.returned_bytes
        )
        .into());
    }
    if cache.open_successful_loads != 0 {
        return Err(format!(
            "managed DOCX loaded payload data during package open in sample {sample_index}: loads={}",
            cache.open_successful_loads
        )
        .into());
    }
    if cache.successful_loads != 1 {
        return Err(format!(
            "managed DOCX expected one successful payload load after text extraction in sample {sample_index}: loads={}",
            cache.successful_loads
        )
        .into());
    }
    if cache.failed_loads != 0 || !cache.budget_managed {
        return Err(format!(
            "managed DOCX cache diagnostics invalid in sample {sample_index}: failed_loads={}, budget_managed={}",
            cache.failed_loads, cache.budget_managed
        )
        .into());
    }
    match (config.policy, source_read.as_ref()) {
        (Policy::Exact, None) => {},
        (Policy::Exact, Some(_)) => {
            return Err(format!(
                "exact managed DOCX policy unexpectedly published source-read diagnostics in sample {sample_index}"
            )
            .into());
        },
        (Policy::ForwardStart(window), Some(diagnostics)) => {
            if !diagnostics.enabled || diagnostics.configured_window_bytes != window {
                return Err(format!(
                    "managed DOCX source-read policy diagnostics disagree with configuration in sample {sample_index}: enabled={}, configured_window_bytes={}, expected_window_bytes={window}",
                    diagnostics.enabled, diagnostics.configured_window_bytes
                )
                .into());
            }
            if diagnostics.returned_bytes != physical.returned_bytes {
                return Err(format!(
                    "managed DOCX source-read diagnostics differ from physical returned bytes in sample {sample_index}: diagnostics={}, physical={}",
                    diagnostics.returned_bytes, physical.returned_bytes
                )
                .into());
            }
        },
        (Policy::ForwardStart(_), None) => {
            return Err(format!(
                "forward-start managed DOCX policy did not publish source-read diagnostics in sample {sample_index}"
            )
            .into());
        },
    }
    let actual_text_sha256 = sha256_hex(actual_text.as_bytes());
    let actual_text_verified = actual_text.as_bytes() == corpus.expected_text.as_bytes()
        && actual_text.len() == EXPECTED_TEXT_BYTES
        && actual_text_sha256 == EXPECTED_TEXT_SHA256;
    if !actual_text_verified {
        return Err(format!("managed DOCX text oracle mismatch in sample {sample_index}").into());
    }
    if before != after {
        return Err(format!(
            "managed DOCX source changed during sample {sample_index}: before={before:?}, after={after:?}"
        )
        .into());
    }
    let allocation = AllocationRecord {
        region_peak_increment_bytes: allocation.as_ref().and_then(|sample| {
            sample
                .region_peak_live_bytes
                .zip(sample.live_bytes_before)
                .and_then(|(peak, before)| peak.checked_sub(before))
        }),
        sample: allocation,
    };
    Ok(SampleRecord {
        sample_index,
        elapsed_ns,
        actual_text_verified,
        actual_text_bytes: actual_text.len(),
        actual_text_sha256,
        source_version_before: before.into(),
        source_version_after: after.into(),
        source_version_unchanged: true,
        policy: PolicyRecord::from_policy(config.policy),
        source_read,
        cache,
        physical: physical.clone(),
        transport,
        budget: BudgetSnapshot {
            memory_before,
            memory_live,
            memory_after_drop,
            memory_released,
            input_bytes_before: input_before,
            input_bytes_after_drop: input_after_drop,
            input_bytes_delta,
            objects_before,
            objects_live,
            objects_after_drop,
            objects_released,
            managed: true,
            memory_released_to_baseline: memory_after_drop == memory_before,
            input_bytes_match_physical_returned: input_bytes_delta == physical.returned_bytes,
        },
        allocation,
    })
}

impl PolicyRecord {
    fn from_policy(policy: Policy) -> Self {
        Self {
            name: policy.name(),
            enabled: !matches!(policy, Policy::Exact),
            configured_window_bytes: policy.window_bytes(),
        }
    }
}

impl SourceReadRecord {
    fn from_diagnostics(value: source_backed::SourceReadDiagnostics) -> Self {
        Self {
            enabled: value.enabled,
            configured_window_bytes: value.configured_window_bytes,
            retained_window_bytes: value.retained_window_bytes,
            requests: value.requests,
            hits: value.hits,
            misses: value.misses,
            fills: value.fills,
            requested_bytes: value.requested_bytes,
            returned_bytes: value.returned_bytes,
        }
    }
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn media_ranges(bytes: &[u8]) -> Result<(Vec<Range<u64>>, usize), Box<dyn Error>> {
    let archive = soapberry_zip::ZipArchive::from_slice(bytes)?;
    let mut ranges = Vec::new();
    let mut members = 0_usize;
    for header in archive.entries() {
        let header = header?;
        let name = header.file_path().try_normalize()?.as_ref().to_owned();
        let entry = archive.get_entry(header.wayfinder())?;
        let (start, end) = entry.compressed_data_range();
        members = members
            .checked_add(1)
            .ok_or("DOCX archive member count overflow")?;
        if name.starts_with("word/media/") {
            ranges.push(start..end);
        }
    }
    Ok((ranges, members))
}

fn build_corpus() -> Result<Corpus, Box<dyn Error>> {
    let corpus = crate::build_docx_source_edit_corpus()?;
    if corpus.manifest.generator != CORPUS_GENERATOR
        || corpus.manifest.archive_bytes != CORPUS_ARCHIVE_BYTES
        || corpus.manifest.archive_sha256 != CORPUS_ARCHIVE_SHA256
        || corpus.manifest.archive_member_count != EXPECTED_ARCHIVE_MEMBERS
    {
        return Err(
            "managed DOCX benchmark corpus no longer matches the pinned source-edit corpus".into(),
        );
    }
    let archive_sha256 = sha256_hex(&corpus.archive);
    if archive_sha256 != CORPUS_ARCHIVE_SHA256 {
        return Err("managed DOCX benchmark archive digest differs from its manifest".into());
    }
    let expected_text = (0..crate::SemanticShape::Medium.docx_paragraphs())
        .map(|index| crate::semantic_docx_text(index, false))
        .collect::<String>();
    if expected_text.len() != EXPECTED_TEXT_BYTES
        || sha256_hex(expected_text.as_bytes()) != EXPECTED_TEXT_SHA256
    {
        return Err("managed DOCX benchmark text oracle changed".into());
    }
    let (media_ranges, archive_members) = media_ranges(&corpus.archive)?;
    if media_ranges.len() != EXPECTED_MEDIA_MEMBERS || archive_members != EXPECTED_ARCHIVE_MEMBERS {
        return Err("managed DOCX benchmark ZIP shape changed".into());
    }
    Ok(Corpus {
        archive: Arc::new(corpus.archive),
        expected_text,
        media_ranges,
        manifest: corpus.manifest,
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

fn next_value(args: &[OsString], index: &mut usize, flag: &str) -> Result<String, Box<dyn Error>> {
    *index = index
        .checked_add(1)
        .ok_or("benchmark argument index overflow")?;
    args.get(*index)
        .ok_or_else(|| format!("{flag} requires a value").into())
        .and_then(|value| {
            value
                .to_str()
                .map(str::to_owned)
                .ok_or_else(|| format!("{flag} value must be valid UTF-8").into())
        })
}

fn parse_revision(value: &str) -> Result<[u8; 40], Box<dyn Error>> {
    if value.len() != 40 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err("--source-revision must be exactly 40 hexadecimal characters".into());
    }
    let mut revision = [0_u8; 40];
    revision.copy_from_slice(value.as_bytes());
    Ok(revision)
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut policy = None;
    let mut window_bytes = None;
    let mut max_range_bytes = None;
    let mut delay_us = None;
    let mut transfer_bytes_per_second = None;
    let mut transfer_delay_policy = pptx_range_source::TransferDelayPolicy::SeparateSleeps;
    let mut transfer_delay_policy_seen = false;
    let mut samples = None;
    let mut warmup = None;
    let mut source_revision = None;
    let mut output = None;
    let mut index = 0_usize;
    while index < args.len() {
        let flag = args[index].to_str().ok_or("argument must be valid UTF-8")?;
        match flag {
            "--policy" => {
                if policy.is_some() {
                    return Err("--policy was specified more than once".into());
                }
                policy = Some(next_value(args, &mut index, flag)?.to_lowercase());
            },
            "--window-bytes" | "--read-ahead" => {
                if window_bytes.is_some() {
                    return Err("the read-ahead window was specified more than once".into());
                }
                let value = parse_usize(&next_value(args, &mut index, flag)?, flag)?;
                if !(1..=MAX_WINDOW_BYTES).contains(&value) {
                    return Err(format!("{flag} must be in 1..={MAX_WINDOW_BYTES}").into());
                }
                window_bytes = Some(value);
            },
            "--max-range" => {
                if max_range_bytes.is_some() {
                    return Err("--max-range was specified more than once".into());
                }
                let value = parse_usize(&next_value(args, &mut index, flag)?, flag)?;
                if !(1..=MAX_RANGE_BYTES).contains(&value) {
                    return Err(format!("--max-range must be in 1..={MAX_RANGE_BYTES}").into());
                }
                max_range_bytes = Some(value);
            },
            "--delay-us" => {
                if delay_us.is_some() {
                    return Err("--delay-us was specified more than once".into());
                }
                let value = parse_u64(&next_value(args, &mut index, flag)?, flag)?;
                if value > MAX_DELAY_US {
                    return Err(format!("--delay-us must be <= {MAX_DELAY_US}").into());
                }
                delay_us = Some(value);
            },
            "--transfer-bytes-per-second" => {
                if transfer_bytes_per_second.is_some() {
                    return Err("--transfer-bytes-per-second was specified more than once".into());
                }
                let value = parse_u64(&next_value(args, &mut index, flag)?, flag)?;
                if !(1_048_576..=MAX_TRANSFER_BYTES_PER_SECOND).contains(&value) {
                    return Err(
                        "--transfer-bytes-per-second must be in 1048576..=1099511627776".into(),
                    );
                }
                transfer_bytes_per_second = NonZeroU64::new(value);
            },
            "--transfer-delay-policy" => {
                if transfer_delay_policy_seen {
                    return Err("--transfer-delay-policy was specified more than once".into());
                }
                transfer_delay_policy_seen = true;
                transfer_delay_policy =
                    match next_value(args, &mut index, flag)?.as_str() {
                        "separate-sleeps" => pptx_range_source::TransferDelayPolicy::SeparateSleeps,
                        "minimum-service" => pptx_range_source::TransferDelayPolicy::MinimumService,
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
                let value = parse_usize(&next_value(args, &mut index, flag)?, flag)?;
                if !(1..=MAX_SAMPLES).contains(&value) {
                    return Err(format!("--samples must be in 1..={MAX_SAMPLES}").into());
                }
                samples = Some(value);
            },
            "--warmup" => {
                if warmup.is_some() {
                    return Err("--warmup was specified more than once".into());
                }
                let value = parse_usize(&next_value(args, &mut index, flag)?, flag)?;
                if value > MAX_WARMUP {
                    return Err(format!("--warmup must be <= {MAX_WARMUP}").into());
                }
                warmup = Some(value);
            },
            "--source-revision" => {
                if source_revision.is_some() {
                    return Err("--source-revision was specified more than once".into());
                }
                source_revision = Some(parse_revision(&next_value(args, &mut index, flag)?)?);
            },
            "--output" => {
                if output.is_some() {
                    return Err("--output was specified more than once".into());
                }
                let value = next_value(args, &mut index, flag)?;
                if value.is_empty() || value == "-" {
                    return Err("--output must be a new regular file path".into());
                }
                output = Some(PathBuf::from(value));
            },
            other => return Err(format!("unknown managed-read-ahead argument: {other}").into()),
        }
        index = index
            .checked_add(1)
            .ok_or("benchmark argument index overflow")?;
    }
    let policy_name = policy.ok_or("--policy is required")?;
    let window_bytes = window_bytes;
    let policy = match policy_name.as_str() {
        "exact" => {
            if window_bytes.is_some() {
                return Err("exact policy cannot specify a read-ahead window".into());
            }
            Policy::Exact
        },
        "forward-start" | "forward_start" | "candidate" => {
            Policy::ForwardStart(window_bytes.unwrap_or(WINDOW_BYTES))
        },
        _ => return Err("--policy must be exact or forward-start".into()),
    };
    Ok(Config {
        policy,
        max_range_bytes: max_range_bytes.ok_or("--max-range is required")?,
        delay_us: delay_us.ok_or("--delay-us is required")?,
        transfer_bytes_per_second,
        transfer_delay_policy,
        samples: samples.ok_or("--samples is required")?,
        warmup: warmup.unwrap_or(0),
        source_revision: source_revision.ok_or("--source-revision is required")?,
        output: output.ok_or("--output is required")?,
    })
}

fn limits_record(limits: ReadLimits, cache: SourceCacheLimits) -> LimitRecord {
    let execution_max_in_flight_bytes = MAX_IN_FLIGHT_BYTES;
    LimitRecord {
        read_limits: ReadLimitsRecord::from_limits(limits),
        cache_max_bytes: cache.max_bytes(),
        cache_max_entries: cache.max_entries(),
        budget_memory_bytes: BUDGET_MEMORY_BYTES,
        budget_input_bytes: BUDGET_INPUT_BYTES,
        budget_output_bytes: BUDGET_OUTPUT_BYTES,
        budget_objects: BUDGET_OBJECTS,
        budget_depth: BUDGET_DEPTH,
        budget_work: BUDGET_WORK,
        execution_workers: 1,
        execution_max_in_flight_tasks: 1,
        execution_max_in_flight_bytes,
        execution_min_parallel_bytes: 0,
    }
}

fn policy_record(policy: Policy) -> PolicyRecord {
    PolicyRecord::from_policy(policy)
}

fn to_revision_string(revision: [u8; 40]) -> String {
    String::from_utf8_lossy(&revision).into_owned()
}

/// Runs the managed production benchmark from arguments following the
/// `docx-managed-read-ahead` selector.
pub fn run_from_args(args: impl IntoIterator<Item = OsString>) -> Result<(), Box<dyn Error>> {
    let args = args.into_iter().collect::<Vec<_>>();
    if args
        .iter()
        .any(|argument| argument == "--help" || argument == "-h")
    {
        println!(
            "docx-managed-read-ahead --policy <exact|forward-start> [--window-bytes N] --max-range N --delay-us N [--transfer-bytes-per-second N --transfer-delay-policy separate-sleeps|minimum-service] --samples N --warmup N --source-revision <40 hex chars> --output PATH"
        );
        return Ok(());
    }
    let config = parse_config(&args)?;
    let corpus = build_corpus()?;
    let limits = ReadLimits::default();
    let cache_limits = SourceCacheLimits::new(CACHE_MAX_BYTES, CACHE_MAX_ENTRIES)?;
    let mut rows = Vec::with_capacity(config.samples);
    for iteration in 0..config.warmup.saturating_add(config.samples) {
        let row = run_sample(&config, &corpus, iteration, limits, cache_limits)?;
        if iteration >= config.warmup {
            rows.push(row);
        }
    }
    let expected_exact_ranges = EXPECTED_EXACT_RANGES
        .iter()
        .map(|&(offset, requested)| ObservedRange {
            offset,
            requested: requested as u64,
            returned: requested as u64,
        })
        .collect();
    let report = Report {
        schema: SCHEMA,
        version: 1,
        case_name: CASE,
        provider_scope: PROVIDER_SCOPE,
        timing_scope: TIMING_SCOPE,
        setup_scope: SETUP_SCOPE,
        allocation_scope: ALLOCATION_SCOPE,
        corpus_version: CORPUS_VERSION,
        corpus_generator: CORPUS_GENERATOR,
        source_bytes: corpus.archive.len(),
        source_sha256: sha256_hex(&corpus.archive),
        expected_text_bytes: EXPECTED_TEXT_BYTES,
        expected_text_sha256: EXPECTED_TEXT_SHA256,
        expected_archive_members: EXPECTED_ARCHIVE_MEMBERS,
        media_ranges: corpus.media_ranges,
        expected_exact_ranges,
        expected_exact_physical_calls: EXPECTED_EXACT_RANGES.len() as u64,
        expected_exact_physical_bytes: EXPECTED_EXACT_RANGES
            .iter()
            .map(|&(_, length)| length as u64)
            .sum(),
        requested_source_revision: to_revision_string(config.source_revision),
        limits: limits_record(limits, cache_limits),
        provider: ProviderRecord {
            name: "PptxRangeSource",
            max_range_bytes: config.max_range_bytes,
            delay_us: config.delay_us,
            transfer_bytes_per_second: config.transfer_bytes_per_second.map(NonZeroU64::get),
            transfer_delay_policy: config.transfer_delay_policy,
            physical_scope: PHYSICAL_SCOPE,
            range_scope: RANGE_SCOPE,
        },
        policy: policy_record(config.policy),
        warmup: config.warmup,
        samples: config.samples,
        corpus: corpus.manifest,
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
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "benchmark parser and lifecycle assertions intentionally panic"
    )]

    use super::*;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};

    fn base_args(extra: &[&str]) -> Vec<OsString> {
        let mut args = vec![
            OsString::from("--policy"),
            OsString::from("exact"),
            OsString::from("--max-range"),
            OsString::from("65536"),
            OsString::from("--delay-us"),
            OsString::from("0"),
            OsString::from("--samples"),
            OsString::from("1"),
            OsString::from("--warmup"),
            OsString::from("0"),
            OsString::from("--source-revision"),
            OsString::from("0000000000000000000000000000000000000000"),
            OsString::from("--output"),
            OsString::from("managed-read-ahead-test.json"),
        ];
        args.extend(extra.iter().map(OsString::from));
        args
    }

    fn tiny_docx() -> Vec<u8> {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>managed</w:t></w:r></w:p></w:body></w:document>"#;
        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/document.xml").unwrap(),
                ct::WML_DOCUMENT_MAIN.to_owned(),
                xml.to_vec(),
            )))
            .unwrap();
        package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        PackageWriter::to_bytes(&package).unwrap()
    }

    fn lifecycle_package(policy: Policy, bytes: Vec<u8>) -> (Budget, source_backed::Package) {
        let budget = Budget::root(
            "managed-read-ahead-test",
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
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(MAX_IN_FLIGHT_BYTES).unwrap(),
            0,
        )
        .unwrap();
        let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
        let package = source_backed::Package::from_read_at_with_limits_and_cache_limits_and_source_read_policy_and_execution_context(
            source,
            ReadLimits::default(),
            SourceCacheLimits::new(CACHE_MAX_BYTES, CACHE_MAX_ENTRIES).unwrap(),
            source_policy(policy).unwrap(),
            context,
        )
        .unwrap();
        (budget, package)
    }

    #[test]
    fn parser_requires_a_finite_policy_and_range_configuration() {
        let exact = parse_config(&base_args(&[])).unwrap();
        assert_eq!(exact.policy, Policy::Exact);
        let mut candidate_args = base_args(&["--window-bytes", "4096"]);
        candidate_args[1] = OsString::from("forward-start");
        let candidate = parse_config(&candidate_args).unwrap();
        assert_eq!(candidate.policy, Policy::ForwardStart(4096));
        assert!(parse_config(&base_args(&["--window-bytes", "0"])).is_err());
        assert!(parse_config(&base_args(&["--window-bytes", "65537"])).is_err());
        assert!(
            parse_config(&base_args(&[
                "--transfer-delay-policy",
                "separate-sleeps",
                "--transfer-delay-policy",
                "minimum-service",
            ]))
            .is_err()
        );
        assert!(
            parse_config(&base_args(
                &["--policy", "exact", "--window-bytes", "4096",]
            ))
            .is_err()
        );
    }

    #[test]
    fn managed_exact_and_forward_lifecycle_have_distinct_policy_diagnostics() {
        let exact_bytes = tiny_docx();
        let (exact_budget, exact_package) = lifecycle_package(Policy::Exact, exact_bytes.clone());
        assert!(exact_package.source_read_diagnostics().unwrap().is_none());
        assert_eq!(
            exact_package.document().unwrap().extract_text().unwrap(),
            "managed"
        );
        drop(exact_package);
        assert_eq!(exact_budget.used(Resource::Memory), 0);

        let (candidate_budget, candidate_package) =
            lifecycle_package(Policy::ForwardStart(WINDOW_BYTES), exact_bytes);
        let diagnostics = candidate_package
            .source_read_diagnostics()
            .unwrap()
            .expect("forward policy must publish diagnostics");
        assert!(diagnostics.enabled);
        assert_eq!(diagnostics.configured_window_bytes, WINDOW_BYTES);
        assert_eq!(
            candidate_package
                .document()
                .unwrap()
                .extract_text()
                .unwrap(),
            "managed"
        );
        drop(candidate_package);
        assert_eq!(candidate_budget.used(Resource::Memory), 0);
    }
}
