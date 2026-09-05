//! Managed source-cache lifecycle evidence for the fixed PPTX corpora.
//!
//! This target records cache diagnostics, hierarchical budget gauges, and
//! positional source counters at ownership boundaries.  It deliberately does
//! not time an operation and makes no latency, allocator, RSS, or cache
//! efficiency claim.  Corpus construction, all semantic/package gates, and
//! the exact output oracle are completed before the first lifecycle sample.

use std::{error::Error, ffi::OsString, fs::OpenOptions, io::Write, path::PathBuf, sync::Arc};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, Resource,
};
use litchi_opc::{ReadLimits, SourceCacheCounterDelta, SourceCacheDiagnostics, SourceCacheLimits};
use serde::Serialize;

use crate::process_metrics;

const SCHEMA: &str = "pptx_cache_retention_v1";
pub(super) const CACHE_LIMIT_BYTES: usize = 64 * 1024 * 1024;
pub(super) const CACHE_LIMIT_ENTRIES: usize = 128;
pub(super) const MAX_SAMPLES: usize = 1_000;
pub(super) const MAX_WARMUP: usize = 1_000;
pub(super) const MAX_WRITE: u64 = 64 * 1024;
pub(super) const MIB: u64 = 1024 * 1024;
pub(super) const GIB: u64 = 1024 * MIB;
pub(super) const MEMORY_LIMIT: u64 = 512 * MIB;
pub(super) const IO_LIMIT: u64 = 8 * GIB;
pub(super) const OBJECT_LIMIT: u64 = 1_000_000;
pub(super) const DEPTH_LIMIT: u64 = 256;

const CACHE_SCOPE: &str = "content-free source-cache counters and gauges from the public PPTX diagnostics API; unavailable states are explicit; untimed evidence only with no latency, allocator, RSS, or cache-efficiency claim";
const BUDGET_SCOPE: &str = "independent source and destination managed Budget roots; Memory and Objects are releasable gauges while InputBytes, OutputBytes, and Work are cumulative charges";
const SOURCE_IO_SCOPE: &str = "InstrumentedSource positional read_calls/read_bytes atomics; logical source reads only, with no physical-storage or decompression attribution";
const PUBLICATION_SCOPE: &str = "source-backed PPTX cross-copy exact output and consuming destination-editor boundary; output is an oracle, not a latency measurement";
const RSS_SCOPE: &str = "optional process-wide VmRSS/VmHWM snapshots from procfs; probe overhead and unrelated process memory remain in scope";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Scenario {
    Lifecycle,
    ExactAdmission,
    OneUnder,
    PinnedEviction,
    OversizedBypass,
    RepeatedPublication,
}

impl Scenario {
    const fn name(self) -> &'static str {
        match self {
            Self::Lifecycle => "lifecycle",
            Self::ExactAdmission => "exact-admission",
            Self::OneUnder => "one-under",
            Self::PinnedEviction => "pinned-eviction",
            Self::OversizedBypass => "oversized-bypass",
            Self::RepeatedPublication => "repeated-publication",
        }
    }

    const fn requires_media_rich(self) -> bool {
        !matches!(self, Self::Lifecycle | Self::RepeatedPublication)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CorpusKind {
    Plain,
    MediaRich,
}

impl CorpusKind {
    const fn name(self) -> &'static str {
        match self {
            Self::Plain => "plain",
            Self::MediaRich => "media-rich",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Config {
    scenario: Scenario,
    corpus: CorpusKind,
    samples: usize,
    warmup: usize,
    source_revision: String,
    output: PathBuf,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct EventDelta {
    hits: u64,
    cold_loads: u64,
    waiter_joins: u64,
    successful_loads: u64,
    failed_loads: u64,
    evictions: u64,
    bypasses: u64,
    oversized_bypasses: u64,
    allocation_bypasses: u64,
    budget_reservation_failures: u64,
}

impl From<SourceCacheCounterDelta> for EventDelta {
    fn from(delta: SourceCacheCounterDelta) -> Self {
        Self {
            hits: delta.hits,
            cold_loads: delta.cold_loads,
            waiter_joins: delta.waiter_joins,
            successful_loads: delta.successful_loads,
            failed_loads: delta.failed_loads,
            evictions: delta.evictions,
            bypasses: delta.bypasses,
            oversized_bypasses: delta.oversized_bypasses,
            allocation_bypasses: delta.allocation_bypasses,
            budget_reservation_failures: delta.budget_reservation_failures,
        }
    }
}

/// A content-free cache snapshot.  Every diagnostic field is optional so an
/// unavailable owner is represented by `null` fields rather than fabricated
/// zero counters.  `availability` is the discriminator consumed by reports.
#[derive(Clone, Debug, Serialize)]
pub(super) struct CachePoint {
    availability: &'static str,
    unavailable_reason: Option<&'static str>,
    interval_label: Option<&'static str>,
    counter_delta_checked: Option<bool>,
    event_delta: Option<EventDelta>,
    hits: Option<u64>,
    cold_loads: Option<u64>,
    waiter_joins: Option<u64>,
    successful_loads: Option<u64>,
    failed_loads: Option<u64>,
    evictions: Option<u64>,
    bypasses: Option<u64>,
    oversized_bypasses: Option<u64>,
    allocation_bypasses: Option<u64>,
    retained_entries: Option<usize>,
    retained_bytes: Option<usize>,
    in_flight_loads: Option<usize>,
    budget_managed: Option<bool>,
    budget_reservation_failures: Option<u64>,
    budget_memory_used: Option<u64>,
    budget_cache_reserved_bytes: Option<u64>,
    budget_memory_limit: Option<u64>,
    budget_input_bytes_used: Option<u64>,
    budget_input_bytes_limit: Option<u64>,
    budget_output_bytes_used: Option<u64>,
    budget_output_bytes_limit: Option<u64>,
    budget_work_used: Option<u64>,
    budget_work_limit: Option<u64>,
    budget_objects_used: Option<u64>,
    budget_objects_limit: Option<u64>,
    budget_catalog_reserved_objects: Option<u64>,
    budget_cache_reserved_objects: Option<u64>,
}

impl CachePoint {
    pub(super) fn unavailable(reason: &'static str) -> Self {
        Self {
            availability: "unavailable",
            unavailable_reason: Some(reason),
            interval_label: None,
            counter_delta_checked: None,
            event_delta: None,
            hits: None,
            cold_loads: None,
            waiter_joins: None,
            successful_loads: None,
            failed_loads: None,
            evictions: None,
            bypasses: None,
            oversized_bypasses: None,
            allocation_bypasses: None,
            retained_entries: None,
            retained_bytes: None,
            in_flight_loads: None,
            budget_managed: None,
            budget_reservation_failures: None,
            budget_memory_used: None,
            budget_cache_reserved_bytes: None,
            budget_memory_limit: None,
            budget_input_bytes_used: None,
            budget_input_bytes_limit: None,
            budget_output_bytes_used: None,
            budget_output_bytes_limit: None,
            budget_work_used: None,
            budget_work_limit: None,
            budget_objects_used: None,
            budget_objects_limit: None,
            budget_catalog_reserved_objects: None,
            budget_cache_reserved_objects: None,
        }
    }

    fn available(
        diagnostics: SourceCacheDiagnostics,
        interval_label: Option<&'static str>,
        delta: Option<EventDelta>,
    ) -> Self {
        Self {
            availability: "available",
            unavailable_reason: None,
            interval_label,
            counter_delta_checked: delta.as_ref().map(|_| true),
            event_delta: delta,
            hits: Some(diagnostics.hits),
            cold_loads: Some(diagnostics.cold_loads),
            waiter_joins: Some(diagnostics.waiter_joins),
            successful_loads: Some(diagnostics.successful_loads),
            failed_loads: Some(diagnostics.failed_loads),
            evictions: Some(diagnostics.evictions),
            bypasses: Some(diagnostics.bypasses),
            oversized_bypasses: Some(diagnostics.oversized_bypasses),
            allocation_bypasses: Some(diagnostics.allocation_bypasses),
            retained_entries: Some(diagnostics.retained_entries),
            retained_bytes: Some(diagnostics.retained_bytes),
            in_flight_loads: Some(diagnostics.in_flight_loads),
            budget_managed: Some(diagnostics.budget_managed),
            budget_reservation_failures: Some(diagnostics.budget_reservation_failures),
            budget_memory_used: Some(diagnostics.budget_memory_used),
            budget_cache_reserved_bytes: Some(diagnostics.budget_cache_reserved_bytes),
            budget_memory_limit: diagnostics.budget_memory_limit,
            budget_input_bytes_used: Some(diagnostics.budget_input_bytes_used),
            budget_input_bytes_limit: diagnostics.budget_input_bytes_limit,
            budget_output_bytes_used: Some(diagnostics.budget_output_bytes_used),
            budget_output_bytes_limit: diagnostics.budget_output_bytes_limit,
            budget_work_used: Some(diagnostics.budget_work_used),
            budget_work_limit: diagnostics.budget_work_limit,
            budget_objects_used: Some(diagnostics.budget_objects_used),
            budget_objects_limit: diagnostics.budget_objects_limit,
            budget_catalog_reserved_objects: Some(diagnostics.budget_catalog_reserved_objects),
            budget_cache_reserved_objects: Some(diagnostics.budget_cache_reserved_objects),
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
pub(super) struct BudgetPoint {
    pub(super) memory_used: u64,
    pub(super) memory_limit: u64,
    input_bytes_used: u64,
    input_bytes_limit: u64,
    output_bytes_used: u64,
    output_bytes_limit: u64,
    work_used: u64,
    work_limit: u64,
    objects_used: u64,
    objects_limit: u64,
    depth_used: u64,
    depth_limit: u64,
}

pub(super) fn budget_point(budget: &Budget) -> BudgetPoint {
    BudgetPoint {
        memory_used: budget.used(Resource::Memory),
        memory_limit: budget.limit(Resource::Memory),
        input_bytes_used: budget.used(Resource::InputBytes),
        input_bytes_limit: budget.limit(Resource::InputBytes),
        output_bytes_used: budget.used(Resource::OutputBytes),
        output_bytes_limit: budget.limit(Resource::OutputBytes),
        work_used: budget.used(Resource::Work),
        work_limit: budget.limit(Resource::Work),
        objects_used: budget.used(Resource::Objects),
        objects_limit: budget.limit(Resource::Objects),
        depth_used: budget.used(Resource::Depth),
        depth_limit: budget.limit(Resource::Depth),
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ReadPoint {
    availability: &'static str,
    unavailable_reason: Option<&'static str>,
    read_calls: Option<u64>,
    read_bytes: Option<u64>,
}

impl ReadPoint {
    const fn unavailable(reason: &'static str) -> Self {
        Self {
            availability: "unavailable",
            unavailable_reason: Some(reason),
            read_calls: None,
            read_bytes: None,
        }
    }

    fn available(source: &crate::InstrumentedSource) -> Self {
        Self {
            availability: "available",
            unavailable_reason: None,
            read_calls: Some(source.read_calls.load(std::sync::atomic::Ordering::SeqCst)),
            read_bytes: Some(source.read_bytes.load(std::sync::atomic::Ordering::SeqCst)),
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
pub(super) struct RssPoint {
    availability: &'static str,
    unavailable_reason: Option<&'static str>,
    rss_bytes: Option<u64>,
    vm_hwm_bytes: Option<u64>,
}

pub(super) fn rss_point() -> RssPoint {
    match process_metrics::Snapshot::read() {
        Ok(snapshot) => RssPoint {
            availability: "available",
            unavailable_reason: None,
            rss_bytes: Some(snapshot.rss_bytes),
            vm_hwm_bytes: Some(snapshot.peak_rss_bytes),
        },
        Err(_) => RssPoint {
            availability: "unavailable",
            unavailable_reason: Some("procfs process metrics are unavailable"),
            rss_bytes: None,
            vm_hwm_bytes: None,
        },
    }
}

#[derive(Clone, Debug, Serialize)]
struct PhaseRecord {
    label: &'static str,
    source_cache: CachePoint,
    destination_cache: CachePoint,
    source_reads: ReadPoint,
    destination_reads: ReadPoint,
    source_budget: BudgetPoint,
    destination_budget: BudgetPoint,
    rss: RssPoint,
}

#[derive(Clone, Debug, Serialize)]
struct LifecycleRow {
    sample_index: usize,
    exact_output_verified: bool,
    output_sha256: String,
    output_bytes: usize,
    phases: Vec<PhaseRecord>,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PhaseDescription {
    label: &'static str,
    live_owners: &'static str,
}

const PHASES: [PhaseDescription; 9] = [
    PhaseDescription {
        label: "baseline",
        live_owners: "fixed corpus only; caller budgets exist, source and destination cache owners are not opened",
    },
    PhaseDescription {
        label: "opened",
        live_owners: "source-backed presentation view, destination editor, caller source Arcs, and reserved sink",
    },
    PhaseDescription {
        label: "planned",
        live_owners: "source-backed view, destination editor, source-retaining plan, caller source Arcs, and sink",
    },
    PhaseDescription {
        label: "published",
        live_owners: "source-backed view, source-retaining plan, publication result, caller source Arcs, and exact sink; destination editor consumed",
    },
    PhaseDescription {
        label: "drop_result",
        live_owners: "source-backed view, source-retaining plan, caller source Arcs, and exact sink",
    },
    PhaseDescription {
        label: "drop_plan",
        live_owners: "source-backed view, caller source Arcs, and exact sink",
    },
    PhaseDescription {
        label: "drop_view",
        live_owners: "caller source Arcs and exact sink; public source-view diagnostic owner released",
    },
    PhaseDescription {
        label: "drop_caller_sources",
        live_owners: "exact sink only; caller source Arcs released",
    },
    PhaseDescription {
        label: "drop_sink",
        live_owners: "no lifecycle-owned source, cache owner, plan, result, or sink handle",
    },
];

#[derive(Clone, Copy, Debug, Serialize)]
struct ConfiguredLimits {
    cache_max_bytes: usize,
    cache_max_entries: usize,
    memory_limit: u64,
    input_bytes_limit: u64,
    output_bytes_limit: u64,
    work_limit: u64,
    objects_limit: u64,
    depth_limit: u64,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    scenario: &'static str,
    corpus: &'static str,
    cache_scope: &'static str,
    budget_scope: &'static str,
    source_io_scope: &'static str,
    publication_scope: &'static str,
    rss_scope: &'static str,
    samples: usize,
    warmup: usize,
    checked_iteration_count: usize,
    source_revision: String,
    binary_sha256: String,
    binary_bytes: u64,
    current_exe: String,
    source_archive_sha256: String,
    source_archive_bytes: usize,
    destination_archive_sha256: String,
    destination_archive_bytes: usize,
    expected_output_sha256: String,
    expected_output_bytes: usize,
    corpus_manifest: crate::CorpusManifest,
    gates: crate::PptxSourceBackedCrossCopyLifecycleGateSummary,
    configured_limits: ConfiguredLimits,
    destination_configured_limits: ConfiguredLimits,
    destination_editor_consumed_during_publish: bool,
    phases: Vec<PhaseDescription>,
    samples_raw: Vec<LifecycleRow>,
}

#[derive(Clone, Debug, Serialize)]
struct NearRow {
    sample_index: usize,
    result: &'static str,
    phases: Vec<PhaseRecord>,
    exact_payload_verified: bool,
    payload_sha256: Option<String>,
    payload_bytes: Option<usize>,
    payload_read_calls: Option<u64>,
    payload_read_bytes: Option<u64>,
    reload_read_calls: Option<u64>,
    reload_read_bytes: Option<u64>,
    accepted_output_bytes: u64,
    refusal_accepted_output_bytes: u64,
    refusal_sink_accepted_bytes: u64,
    refusal_resource: Option<&'static str>,
    typed_memory_refusal: bool,
    typed_resource_refusal: bool,
    eviction_verified: bool,
    pinned_bypass_verified: bool,
    oversized_bypass_verified: bool,
    repeated_output_identities_verified: bool,
    publication_output_sha256: Vec<String>,
    root_memory_floor: Option<u64>,
    metadata_memory_used: Option<u64>,
    payload_memory_used: Option<u64>,
    drop_image_memory_used: Option<u64>,
    leaf_bytes: Option<u64>,
    memory_limit: u64,
    cache_max_bytes: usize,
    cache_max_entries: usize,
    configured_limits: ConfiguredLimits,
    destination_configured_limits: ConfiguredLimits,
}

#[derive(Debug, Serialize)]
struct NearReport {
    schema: &'static str,
    scenario: &'static str,
    corpus: &'static str,
    cache_scope: &'static str,
    budget_scope: &'static str,
    source_io_scope: &'static str,
    publication_scope: &'static str,
    rss_scope: &'static str,
    samples: usize,
    warmup: usize,
    checked_iteration_count: usize,
    source_revision: String,
    binary_sha256: String,
    binary_bytes: u64,
    current_exe: String,
    source_archive_sha256: String,
    source_archive_bytes: usize,
    destination_archive_sha256: String,
    destination_archive_bytes: usize,
    expected_output_sha256: String,
    expected_output_bytes: usize,
    expected_image_sha256: Option<String>,
    expected_image_bytes: Option<usize>,
    corpus_manifest: crate::CorpusManifest,
    gates: crate::PptxSourceBackedCrossCopyLifecycleGateSummary,
    configured_limits: ConfiguredLimits,
    destination_configured_limits: ConfiguredLimits,
    phases: Vec<PhaseDescription>,
    samples_raw: Vec<NearRow>,
}

#[derive(Clone)]
pub(super) struct ManagedContext {
    pub(super) budget: Budget,
    pub(super) context: ExecutionContext,
}

pub(super) fn managed_context(scope: &'static str) -> Result<ManagedContext, Box<dyn Error>> {
    managed_context_with_limits(scope, MEMORY_LIMIT, IO_LIMIT)
}

fn managed_context_with_limits(
    scope: &'static str,
    memory_limit: u64,
    output_limit: u64,
) -> Result<ManagedContext, Box<dyn Error>> {
    let budget = Budget::root(
        scope,
        Limits::new(
            memory_limit.max(1),
            IO_LIMIT,
            output_limit.max(1),
            OBJECT_LIMIT,
            DEPTH_LIMIT,
            IO_LIMIT,
        ),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let workers = std::num::NonZeroUsize::new(1).ok_or("one worker is invalid")?;
    let in_flight_bytes =
        std::num::NonZeroU64::new(memory_limit.max(1)).ok_or("memory limit cannot be zero")?;
    let execution_limits = ExecutionLimits::new(workers, workers, in_flight_bytes, 0)?;
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    Ok(ManagedContext { budget, context })
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut scenario = None;
    let mut corpus = None;
    let mut samples = None;
    let mut warmup = None;
    let mut source_revision = None;
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("cache-retention argument is not valid UTF-8")?;
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("missing value for {flag}"))?;
        let value_text = value
            .to_str()
            .ok_or("cache-retention argument value is not valid UTF-8")?;
        index += 2;
        match flag {
            "--scenario" => {
                if scenario.is_some() {
                    return Err("duplicate --scenario".into());
                }
                scenario = Some(match value_text {
                    "lifecycle" => Scenario::Lifecycle,
                    "exact-admission" => Scenario::ExactAdmission,
                    "one-under" => Scenario::OneUnder,
                    "pinned-eviction" => Scenario::PinnedEviction,
                    "oversized-bypass" => Scenario::OversizedBypass,
                    "repeated-publication" => Scenario::RepeatedPublication,
                    _ => return Err("unknown cache-retention scenario".into()),
                });
            },
            "--corpus" => {
                if corpus.is_some() {
                    return Err("duplicate --corpus".into());
                }
                corpus = Some(match value_text {
                    "plain" => CorpusKind::Plain,
                    "media-rich" => CorpusKind::MediaRich,
                    _ => return Err("--corpus must be plain or media-rich".into()),
                });
            },
            "--samples" => {
                if samples.is_some() {
                    return Err("duplicate --samples".into());
                }
                samples = Some(parse_bounded(value_text, "samples", 1, MAX_SAMPLES)?);
            },
            "--warmup" => {
                if warmup.is_some() {
                    return Err("duplicate --warmup".into());
                }
                warmup = Some(parse_bounded(value_text, "warmup", 0, MAX_WARMUP)?);
            },
            "--source-revision" => {
                if source_revision.is_some() {
                    return Err("duplicate --source-revision".into());
                }
                if value_text.len() != 40
                    || !value_text.bytes().all(|byte| byte.is_ascii_hexdigit())
                    || value_text.to_ascii_lowercase() != value_text
                {
                    return Err(
                        "--source-revision must contain exactly 40 lowercase hexadecimal characters"
                            .into(),
                    );
                }
                source_revision = Some(value_text.to_owned());
            },
            "--output" => {
                if output.is_some() {
                    return Err("duplicate --output".into());
                }
                if value_text.is_empty() || value_text == "-" {
                    return Err("--output must be a create_new file path".into());
                }
                output = Some(PathBuf::from(value));
            },
            _ => return Err(format!("unknown cache-retention argument: {flag}").into()),
        }
    }
    let scenario = scenario.ok_or("missing --scenario")?;
    let corpus = corpus.ok_or("missing --corpus")?;
    if scenario.requires_media_rich() && corpus != CorpusKind::MediaRich {
        return Err(format!("{} requires --corpus media-rich", scenario.name()).into());
    }
    Ok(Config {
        scenario,
        corpus,
        samples: samples.ok_or("missing --samples")?,
        warmup: warmup.ok_or("missing --warmup")?,
        source_revision: source_revision.ok_or("missing --source-revision")?,
        output: output.ok_or("missing --output")?,
    })
}

fn parse_bounded(
    value: &str,
    name: &str,
    minimum: usize,
    maximum: usize,
) -> Result<usize, Box<dyn Error>> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("--{name} must be an unsigned decimal integer"))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(format!("--{name} must be between {minimum} and {maximum}").into());
    }
    Ok(parsed)
}

fn validate_budget_gauges(
    diagnostics: SourceCacheDiagnostics,
    budget: &Budget,
    cache_limits: SourceCacheLimits,
) -> Result<(), Box<dyn Error>> {
    if !diagnostics.budget_managed {
        return Err("managed cache lifecycle returned an unmanaged diagnostic".into());
    }
    let expected = [
        (
            "memory",
            diagnostics.budget_memory_used,
            budget.used(Resource::Memory),
        ),
        (
            "input bytes",
            diagnostics.budget_input_bytes_used,
            budget.used(Resource::InputBytes),
        ),
        (
            "output bytes",
            diagnostics.budget_output_bytes_used,
            budget.used(Resource::OutputBytes),
        ),
        (
            "work",
            diagnostics.budget_work_used,
            budget.used(Resource::Work),
        ),
        (
            "objects",
            diagnostics.budget_objects_used,
            budget.used(Resource::Objects),
        ),
    ];
    for (name, observed, expected) in expected {
        if observed != expected {
            return Err(format!("managed {name} gauge differs from caller budget").into());
        }
    }
    let limits = [
        (
            "memory",
            diagnostics.budget_memory_limit,
            budget.limit(Resource::Memory),
        ),
        (
            "input bytes",
            diagnostics.budget_input_bytes_limit,
            budget.limit(Resource::InputBytes),
        ),
        (
            "output bytes",
            diagnostics.budget_output_bytes_limit,
            budget.limit(Resource::OutputBytes),
        ),
        (
            "work",
            diagnostics.budget_work_limit,
            budget.limit(Resource::Work),
        ),
        (
            "objects",
            diagnostics.budget_objects_limit,
            budget.limit(Resource::Objects),
        ),
    ];
    for (name, observed, expected) in limits {
        if observed != Some(expected) {
            return Err(format!("managed {name} limit differs from caller budget").into());
        }
    }
    if diagnostics.budget_cache_reserved_bytes > diagnostics.budget_memory_used
        || diagnostics.budget_cache_reserved_objects > diagnostics.budget_objects_used
        || diagnostics.budget_catalog_reserved_objects > diagnostics.budget_objects_used
        || diagnostics.retained_entries > cache_limits.max_entries()
        || diagnostics.retained_bytes > cache_limits.max_bytes()
    {
        return Err("managed cache diagnostic exceeds its configured gauge bounds".into());
    }
    Ok(())
}

pub(super) fn cache_point(
    diagnostics: SourceCacheDiagnostics,
    previous: Option<SourceCacheDiagnostics>,
    budget: &Budget,
    cache_limits: SourceCacheLimits,
    interval_label: Option<&'static str>,
) -> Result<(CachePoint, SourceCacheDiagnostics), Box<dyn Error>> {
    validate_budget_gauges(diagnostics, budget, cache_limits)?;
    let delta = previous
        .map(|before| SourceCacheDiagnostics::checked_counter_delta(before, diagnostics))
        .transpose()?
        .map(EventDelta::from);
    Ok((
        CachePoint::available(diagnostics, interval_label, delta),
        diagnostics,
    ))
}

fn phase(
    label: &'static str,
    source_cache: CachePoint,
    destination_cache: CachePoint,
    source: Option<&Arc<crate::InstrumentedSource>>,
    destination: Option<&Arc<crate::InstrumentedSource>>,
    source_budget: &Budget,
    destination_budget: &Budget,
) -> PhaseRecord {
    PhaseRecord {
        label,
        source_cache,
        destination_cache,
        source_reads: source.map_or(
            ReadPoint::unavailable("caller source has not been created or was dropped"),
            |source| ReadPoint::available(source.as_ref()),
        ),
        destination_reads: destination.map_or(
            ReadPoint::unavailable("caller destination source has not been created or was dropped"),
            |destination| ReadPoint::available(destination.as_ref()),
        ),
        source_budget: budget_point(source_budget),
        destination_budget: budget_point(destination_budget),
        rss: rss_point(),
    }
}

fn sink_ceiling(expected_bytes: usize) -> Result<u64, Box<dyn Error>> {
    u64::try_from(expected_bytes)?
        .checked_mul(2)
        .and_then(|value| value.checked_add(MAX_WRITE))
        .ok_or_else(|| "PPTX cache lifecycle sink ceiling overflows u64".into())
}

fn configured_limits(
    cache_max_bytes: usize,
    cache_max_entries: usize,
    memory_limit: u64,
    output_limit: u64,
) -> ConfiguredLimits {
    ConfiguredLimits {
        cache_max_bytes,
        cache_max_entries,
        memory_limit,
        input_bytes_limit: IO_LIMIT,
        output_bytes_limit: output_limit,
        work_limit: IO_LIMIT,
        objects_limit: OBJECT_LIMIT,
        depth_limit: DEPTH_LIMIT,
    }
}

const NEAR_PHASES: [PhaseDescription; 10] = [
    PhaseDescription {
        label: "baseline",
        live_owners: "fixed corpus and caller budget only; cache owner is not opened",
    },
    PhaseDescription {
        label: "metadata",
        live_owners: "source-backed presentation, selected slide metadata, caller source Arc, and budget",
    },
    PhaseDescription {
        label: "payload",
        live_owners: "source-backed presentation, selected image or publication owner, returned payload handles when admitted, and caller source Arc",
    },
    PhaseDescription {
        label: "drop_image",
        live_owners: "source-backed presentation and caller source Arc after returned image handle scope ends",
    },
    PhaseDescription {
        label: "drop_view",
        live_owners: "caller source Arc after the public source-view diagnostic owner ends",
    },
    PhaseDescription {
        label: "opened_1",
        live_owners: "first repeated-publication destination editor and shared caller sources",
    },
    PhaseDescription {
        label: "planned_1",
        live_owners: "first repeated-publication plan, source view, destination editor, and shared caller sources",
    },
    PhaseDescription {
        label: "published_1",
        live_owners: "first publication result, source view, plan, sink, and shared caller sources; editor consumed",
    },
    PhaseDescription {
        label: "drop_plan_1",
        live_owners: "first repeated-publication source view, shared caller sources, and destination sink scope",
    },
    PhaseDescription {
        label: "drop_sink_1",
        live_owners: "first repeated-publication source view and shared caller sources after sink drop",
    },
];

const REPEATED_OPENED_LABELS: [&str; 3] = ["opened_1", "opened_2", "opened_3"];
const REPEATED_PLANNED_LABELS: [&str; 3] = ["planned_1", "planned_2", "planned_3"];
const REPEATED_PUBLISHED_LABELS: [&str; 3] = ["published_1", "published_2", "published_3"];
const REPEATED_DROP_RESULT_LABELS: [&str; 3] = ["drop_result_1", "drop_result_2", "drop_result_3"];
const REPEATED_DROP_PLAN_LABELS: [&str; 3] = ["drop_plan_1", "drop_plan_2", "drop_plan_3"];
const REPEATED_DROP_SINK_LABELS: [&str; 3] = ["drop_sink_1", "drop_sink_2", "drop_sink_3"];

pub(super) fn validate_gates(
    corpus: &crate::PptxSourceBackedCrossCopyCorpus,
) -> Result<crate::PptxSourceBackedCrossCopyLifecycleGateSummary, Box<dyn Error>> {
    let gates = corpus
        .lifecycle_gates
        .as_ref()
        .ok_or("source-backed PPTX cache lifecycle corpus has no lifecycle gates")?;
    let complete = gates.matched_owned_corpus_verified
        && gates.semantic_output_verified
        && gates.package_topology_verified
        && gates.dependency_boundary_verified
        && gates.layout_reuse_verified
        && gates.untouched_destination_members_verified
        && gates.deterministic_output_verified
        && gates.source_version_stability_verified
        && gates.source_revision_refusal_verified
        && gates.destination_revision_refusal_verified
        && gates.foreign_destination_refusal_verified
        && gates.added_opc_parts_verified
        && gates.added_zip_members_verified
        && gates.media_leaf_payloads_verified
        && gates.media_leaf_content_types_verified
        && gates.media_relationships_verified;
    if !complete {
        return Err("source-backed PPTX cache lifecycle corpus has incomplete gates".into());
    }
    Ok(gates.clone())
}

struct NearCorpus {
    pinned_a_payload: Vec<u8>,
    images: crate::PptxSourceImageQueryCorpus,
    lifecycle: crate::PptxSourceBackedCrossCopyCorpus,
    gates: crate::PptxSourceBackedCrossCopyLifecycleGateSummary,
}

fn build_near_corpus() -> Result<NearCorpus, Box<dyn Error>> {
    let images = crate::build_pptx_source_image_query_corpus()?;
    let lifecycle = crate::build_pptx_source_backed_cross_copy_corpus(
        crate::Case::PptxSourceBackedCrossCopyMediaRichLifecycle,
    )?;
    let gates = validate_gates(&lifecycle)?;
    if crate::sha256_hex(&images.archive) != crate::sha256_hex(&lifecycle.source_archive) {
        return Err("PPTX image and lifecycle corpora do not share the source archive".into());
    }
    Ok(NearCorpus {
        pinned_a_payload: crate::pptx_cross_copy_media_payload(0),
        images,
        lifecycle,
        gates,
    })
}

fn validate_image_metadata(
    slide: &litchi_pptx::SourceSlide,
    corpus: &crate::PptxSourceImageQueryCorpus,
) -> Result<(), Box<dyn Error>> {
    let descriptors = slide.images()?;
    let observed = descriptors
        .iter()
        .map(crate::pptx_source_image_metadata)
        .collect::<Result<Vec<_>, _>>()?;
    if observed != corpus.expected_images
        || observed.get(corpus.selected_position) != Some(&corpus.selected_image)
    {
        return Err("PPTX image metadata differs from the prevalidated oracle".into());
    }
    Ok(())
}

fn calibrate_image_memory(
    corpus: &crate::PptxSourceImageQueryCorpus,
    cache_limits: SourceCacheLimits,
) -> Result<(u64, u64, u64), Box<dyn Error>> {
    let managed = managed_context_with_limits(
        "pptx-cache-retention-image-calibration",
        MEMORY_LIMIT,
        IO_LIMIT,
    )?;
    let source = Arc::new(crate::InstrumentedSource::new(
        corpus.archive.clone(),
        Vec::new(),
    ));
    let source_read: Arc<dyn ReadAt> = source.clone();
    let view = litchi_pptx::SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
        source_read,
        ReadLimits::default(),
        cache_limits,
        managed.context,
    )?;
    let root_memory = managed.budget.used(Resource::Memory);
    let slide = view
        .slide(corpus.source_slide)
        .ok_or("PPTX image calibration source slide is missing")?;
    validate_image_metadata(&slide, corpus)?;
    let metadata_memory = managed.budget.used(Resource::Memory);
    let leaf_bytes = u64::try_from(corpus.selected_payload.len())?;
    if root_memory == 0 || metadata_memory < root_memory || leaf_bytes == 0 {
        return Err("PPTX image calibration produced invalid managed memory boundaries".into());
    }
    drop(slide);
    drop(view);
    drop(source);
    Ok((root_memory, metadata_memory, leaf_bytes))
}

fn memory_refusal(error: &litchi_pptx::Error) -> bool {
    matches!(
        error,
        litchi_pptx::Error::Opc(litchi_opc::OpcError::Execution(
            litchi_core::ExecutionError::ResourceLimit(limit)
        )) if limit.resource == Resource::Memory
    )
}

fn output_refusal(error: &litchi_pptx::Error) -> bool {
    matches!(
        error,
        litchi_pptx::Error::Opc(litchi_opc::OpcError::Execution(
            litchi_core::ExecutionError::ResourceLimit(limit)
        )) if limit.resource == Resource::OutputBytes
    )
}

fn source_bytes_delta(before: &ReadPoint, after: &ReadPoint) -> Result<(u64, u64), Box<dyn Error>> {
    let calls = after
        .read_calls
        .zip(before.read_calls)
        .and_then(|(after, before)| after.checked_sub(before))
        .ok_or("source-read call counters are unavailable or regressed")?;
    let bytes = after
        .read_bytes
        .zip(before.read_bytes)
        .and_then(|(after, before)| after.checked_sub(before))
        .ok_or("source-read byte counters are unavailable or regressed")?;
    Ok((calls, bytes))
}

fn run_image_boundary_iteration(
    corpus: &crate::PptxSourceImageQueryCorpus,
    scenario: Scenario,
    sample_index: usize,
    root_memory: u64,
    _metadata_memory: u64,
    leaf_bytes: u64,
) -> Result<NearRow, Box<dyn Error>> {
    let cache_max_bytes = if scenario == Scenario::OversizedBypass {
        usize::try_from(leaf_bytes.checked_sub(1).ok_or("image leaf is empty")?)?
    } else {
        CACHE_LIMIT_BYTES
    };
    let cache_max_entries = if scenario == Scenario::PinnedEviction {
        2
    } else {
        CACHE_LIMIT_ENTRIES
    };
    let cache_limits = SourceCacheLimits::new(cache_max_bytes, cache_max_entries)?;
    let memory_limit = match scenario {
        Scenario::ExactAdmission => root_memory
            .checked_add(leaf_bytes)
            .ok_or("exact image memory ceiling overflows")?,
        Scenario::OneUnder => root_memory
            .checked_add(leaf_bytes)
            .and_then(|value| value.checked_sub(1))
            .ok_or("one-under image memory ceiling underflows")?,
        _ => MEMORY_LIMIT,
    };
    let source_context =
        managed_context_with_limits("pptx-cache-retention-image-source", memory_limit, IO_LIMIT)?;
    let destination_context = managed_context("pptx-cache-retention-image-destination")?;
    let configured = configured_limits(cache_max_bytes, cache_max_entries, memory_limit, IO_LIMIT);
    let destination_configured = configured_limits(
        CACHE_LIMIT_BYTES,
        CACHE_LIMIT_ENTRIES,
        MEMORY_LIMIT,
        IO_LIMIT,
    );
    let mut phases = Vec::with_capacity(7);
    phases.push(phase(
        "baseline",
        CachePoint::unavailable("source image owner has not been opened"),
        CachePoint::unavailable("image boundary has no destination editor owner"),
        None,
        None,
        &source_context.budget,
        &destination_context.budget,
    ));

    let source = Arc::new(crate::InstrumentedSource::new(
        corpus.archive.clone(),
        Vec::new(),
    ));
    let source_read: Arc<dyn ReadAt> = source.clone();
    let view = litchi_pptx::SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
        source_read,
        ReadLimits::default(),
        cache_limits,
        source_context.context.clone(),
    )?;
    let actual_root_memory = source_context.budget.used(Resource::Memory);
    if actual_root_memory != root_memory {
        return Err(format!(
            "image memory calibration root changed: calibrated {root_memory}, observed {actual_root_memory}"
        )
        .into());
    }
    let slide = view
        .slide(corpus.source_slide)
        .ok_or("PPTX image boundary source slide is missing")?;
    validate_image_metadata(&slide, corpus)?;
    let actual_metadata_memory = source_context.budget.used(Resource::Memory);
    let metadata_diagnostics = view.try_cache_diagnostics()?;
    let (metadata_cache, metadata_diagnostics) = cache_point(
        metadata_diagnostics,
        None,
        &source_context.budget,
        cache_limits,
        None,
    )?;
    phases.push(phase(
        "metadata",
        metadata_cache,
        CachePoint::unavailable("image boundary has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    if actual_metadata_memory < root_memory {
        return Err("image metadata memory fell below the calibrated root floor".into());
    }
    let reads_before = ReadPoint::available(&source);
    let read_result = slide.read_image(corpus.selected_position);
    let mut retained_image = None;
    let mut exact_payload_verified = false;
    let mut payload_sha256 = None;
    let mut typed_memory_refusal = false;
    let mut typed_resource_refusal = false;
    match read_result {
        Ok(image) => {
            if image.bytes() != corpus.selected_payload.as_slice()
                || crate::sha256_hex(image.bytes()) != corpus.selected_payload_sha256
            {
                return Err("selected image payload differs from the prevalidated oracle".into());
            }
            exact_payload_verified = true;
            payload_sha256 = Some(crate::sha256_hex(image.bytes()));
            retained_image = Some(image);
        },
        Err(error) => {
            typed_memory_refusal = memory_refusal(&error);
            typed_resource_refusal = typed_memory_refusal;
            if scenario != Scenario::OneUnder || !typed_memory_refusal {
                return Err(format!("selected image read failed unexpectedly: {error}").into());
            }
        },
    }
    let reads_after = ReadPoint::available(&source);
    let (payload_read_calls, payload_read_bytes) = source_bytes_delta(&reads_before, &reads_after)?;
    let payload_memory_used = source_context.budget.used(Resource::Memory);
    let after_diagnostics = view.try_cache_diagnostics()?;
    let (payload_cache, after_diagnostics) = cache_point(
        after_diagnostics,
        Some(metadata_diagnostics),
        &source_context.budget,
        cache_limits,
        Some("metadata_to_payload"),
    )?;
    if scenario == Scenario::OneUnder {
        if payload_read_calls != 0
            || payload_read_bytes != 0
            || source_context.budget.used(Resource::OutputBytes) != 0
            || payload_memory_used != root_memory
        {
            return Err("one-under image refusal crossed the payload or output boundary".into());
        }
    } else if payload_read_calls == 0
        || payload_read_bytes == 0
        || payload_memory_used
            != root_memory
                .checked_add(leaf_bytes)
                .ok_or("image payload memory ceiling overflows")?
    {
        return Err("successful image admission performed no selected payload read".into());
    }
    if scenario == Scenario::OversizedBypass
        && (after_diagnostics.oversized_bypasses == metadata_diagnostics.oversized_bypasses
            || after_diagnostics.bypasses == metadata_diagnostics.bypasses)
    {
        return Err("oversized image read did not record both bypass counters".into());
    }
    phases.push(phase(
        "payload",
        payload_cache,
        CachePoint::unavailable("image boundary has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    drop(retained_image);
    let drop_image_memory_used = source_context.budget.used(Resource::Memory);
    match scenario {
        Scenario::ExactAdmission if drop_image_memory_used != payload_memory_used => {
            return Err("exact admitted payload reservation was released before owner drop".into());
        },
        Scenario::OversizedBypass if drop_image_memory_used != root_memory => {
            return Err(
                "oversized bypass payload reservation did not release at owner drop".into(),
            );
        },
        Scenario::OneUnder if drop_image_memory_used != root_memory => {
            return Err("one-under refusal left a payload memory reservation".into());
        },
        _ => {},
    }
    let after_drop_diagnostics = view.try_cache_diagnostics()?;
    let (drop_cache, after_drop_diagnostics) = cache_point(
        after_drop_diagnostics,
        Some(after_diagnostics),
        &source_context.budget,
        cache_limits,
        Some("payload_to_drop_image"),
    )?;
    phases.push(phase(
        "drop_image",
        drop_cache,
        CachePoint::unavailable("image boundary has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    let (reload_read_calls, reload_read_bytes) = if scenario == Scenario::OversizedBypass {
        let reload_before = ReadPoint::available(&source);
        let reloaded = slide.read_image(corpus.selected_position)?;
        if reloaded.bytes() != corpus.selected_payload.as_slice()
            || crate::sha256_hex(reloaded.bytes()) != corpus.selected_payload_sha256
        {
            return Err("oversized image reload differs from the prevalidated oracle".into());
        }
        let reload_memory_used = source_context.budget.used(Resource::Memory);
        if reload_memory_used
            != root_memory
                .checked_add(leaf_bytes)
                .ok_or("oversized image reload memory ceiling overflows")?
        {
            return Err("oversized image reload did not reserve its payload memory".into());
        }
        let reload_after = ReadPoint::available(&source);
        let (reload_calls, reload_bytes) = source_bytes_delta(&reload_before, &reload_after)?;
        if reload_calls == 0 || reload_bytes == 0 {
            return Err("oversized image reload did not perform a cold source read".into());
        }
        let reload_diagnostics = view.try_cache_diagnostics()?;
        let (reload_cache, _reload_diagnostics) = cache_point(
            reload_diagnostics,
            Some(after_drop_diagnostics),
            &source_context.budget,
            cache_limits,
            Some("drop_image_to_oversized_reload"),
        )?;
        if reload_cache
            .event_delta
            .as_ref()
            .is_none_or(|delta| delta.oversized_bypasses == 0 || delta.bypasses == 0)
        {
            return Err("oversized image reload did not record a bypass interval".into());
        }
        phases.push(phase(
            "oversized_reload",
            reload_cache,
            CachePoint::unavailable("image boundary has no destination editor owner"),
            Some(&source),
            None,
            &source_context.budget,
            &destination_context.budget,
        ));
        drop(reloaded);
        if source_context.budget.used(Resource::Memory) != root_memory {
            return Err("oversized image reload reservation did not release at owner drop".into());
        }
        (Some(reload_calls), Some(reload_bytes))
    } else {
        (None, None)
    };
    drop(slide);
    drop(view);
    phases.push(phase(
        "drop_view",
        CachePoint::unavailable("source image view was dropped; no public cache owner remains"),
        CachePoint::unavailable("image boundary has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    drop(source);
    phases.push(phase(
        "drop_source",
        CachePoint::unavailable("caller source owner was dropped"),
        CachePoint::unavailable("image boundary has no destination editor owner"),
        None,
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    let result = match scenario {
        Scenario::ExactAdmission => "admitted_exact_memory_ceiling",
        Scenario::OneUnder => "typed_memory_refusal_before_payload_read",
        Scenario::OversizedBypass => "admitted_uncached_oversized_payload",
        _ => return Err("invalid image-boundary scenario".into()),
    };
    Ok(NearRow {
        sample_index,
        result,
        phases,
        exact_payload_verified,
        payload_sha256,
        payload_bytes: exact_payload_verified.then_some(usize::try_from(leaf_bytes)?),
        payload_read_calls: Some(payload_read_calls),
        payload_read_bytes: Some(payload_read_bytes),
        reload_read_calls,
        reload_read_bytes,
        accepted_output_bytes: 0,
        refusal_accepted_output_bytes: 0,
        refusal_sink_accepted_bytes: 0,
        refusal_resource: (scenario == Scenario::OneUnder).then_some("Memory"),
        typed_memory_refusal,
        typed_resource_refusal,
        eviction_verified: false,
        pinned_bypass_verified: false,
        oversized_bypass_verified: scenario == Scenario::OversizedBypass,
        repeated_output_identities_verified: false,
        publication_output_sha256: Vec::new(),
        root_memory_floor: Some(root_memory),
        metadata_memory_used: Some(actual_metadata_memory),
        payload_memory_used: Some(payload_memory_used),
        drop_image_memory_used: Some(drop_image_memory_used),
        leaf_bytes: Some(leaf_bytes),
        memory_limit,
        cache_max_bytes,
        cache_max_entries,
        configured_limits: configured,
        destination_configured_limits: destination_configured,
    })
}

fn run_pinned_eviction_iteration(
    corpus: &crate::PptxSourceImageQueryCorpus,
    sample_index: usize,
    expected_a: &[u8],
) -> Result<NearRow, Box<dyn Error>> {
    let cache_limits = SourceCacheLimits::new(CACHE_LIMIT_BYTES, 2)?;
    let source_context = managed_context("pptx-cache-retention-pinned-source")?;
    let destination_context = managed_context("pptx-cache-retention-pinned-destination")?;
    let configured = configured_limits(CACHE_LIMIT_BYTES, 2, MEMORY_LIMIT, IO_LIMIT);
    let destination_configured = configured_limits(
        CACHE_LIMIT_BYTES,
        CACHE_LIMIT_ENTRIES,
        MEMORY_LIMIT,
        IO_LIMIT,
    );
    let mut phases = Vec::with_capacity(9);
    phases.push(phase(
        "baseline",
        CachePoint::unavailable("source image owner has not been opened"),
        CachePoint::unavailable("pinned row has no destination editor owner"),
        None,
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    let source = Arc::new(crate::InstrumentedSource::new(
        corpus.archive.clone(),
        Vec::new(),
    ));
    let source_read: Arc<dyn ReadAt> = source.clone();
    let view = litchi_pptx::SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
        source_read,
        ReadLimits::default(),
        cache_limits,
        source_context.context.clone(),
    )?;
    let slide = view
        .slide(corpus.source_slide)
        .ok_or("pinned row source slide is missing")?;
    validate_image_metadata(&slide, corpus)?;
    let metadata_diagnostics = view.try_cache_diagnostics()?;
    let (metadata_cache, metadata_diagnostics) = cache_point(
        metadata_diagnostics,
        None,
        &source_context.budget,
        cache_limits,
        None,
    )?;
    phases.push(phase(
        "metadata",
        metadata_cache,
        CachePoint::unavailable("pinned row has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    let reads_before_pair = ReadPoint::available(&source);
    let image_a = slide.read_image(0)?;
    if image_a.bytes() != expected_a {
        return Err("pinned image A payload differs from its oracle".into());
    }
    let after_a = view.try_cache_diagnostics()?;
    let (cache_a, after_a) = cache_point(
        after_a,
        Some(metadata_diagnostics),
        &source_context.budget,
        cache_limits,
        Some("metadata_to_pinned_a"),
    )?;
    phases.push(phase(
        "pinned_a",
        cache_a,
        CachePoint::unavailable("pinned row has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    let image_b = slide.read_image(corpus.selected_position)?;
    let expected_b = corpus.selected_payload.as_slice();
    if image_b.bytes() != expected_b {
        return Err("pinned image B payload differs from its oracle".into());
    }
    let (payload_read_calls, payload_read_bytes) =
        source_bytes_delta(&reads_before_pair, &ReadPoint::available(&source))?;
    let after_b = view.try_cache_diagnostics()?;
    let (cache_b, after_b) = cache_point(
        after_b,
        Some(after_a),
        &source_context.budget,
        cache_limits,
        Some("pinned_a_to_pinned_b"),
    )?;
    if after_a.evictions <= metadata_diagnostics.evictions
        || after_b.evictions != after_a.evictions
        || after_b.bypasses <= after_a.bypasses
        || after_b.oversized_bypasses != metadata_diagnostics.oversized_bypasses
        || after_a.retained_entries != 2
        || after_b.retained_entries != after_a.retained_entries
        || after_b.retained_bytes != after_a.retained_bytes
        || after_b.budget_memory_used
            != after_a
                .budget_memory_used
                .checked_add(u64::try_from(expected_b.len())?)
                .ok_or("pinned image memory sum overflows")?
    {
        return Err("pinned row did not record clean eviction and normal pinned bypass".into());
    }
    phases.push(phase(
        "pinned_b",
        cache_b,
        CachePoint::unavailable("pinned row has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    drop(image_a);
    drop(image_b);
    let after_drop = view.try_cache_diagnostics()?;
    if after_drop.budget_memory_used != after_a.budget_memory_used {
        return Err("uncached pinned image reservation did not release on handle drop".into());
    }
    let (drop_cache, after_drop) = cache_point(
        after_drop,
        Some(after_b),
        &source_context.budget,
        cache_limits,
        Some("pinned_b_to_drop_images"),
    )?;
    phases.push(phase(
        "drop_image",
        drop_cache,
        CachePoint::unavailable("pinned row has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    let reads_before_reload = ReadPoint::available(&source);
    let reloaded_b = slide.read_image(corpus.selected_position)?;
    if reloaded_b.bytes() != expected_b {
        return Err("reloaded pinned image B payload differs from its oracle".into());
    }
    let (reload_read_calls, reload_read_bytes) =
        source_bytes_delta(&reads_before_reload, &ReadPoint::available(&source))?;
    let after_reload = view.try_cache_diagnostics()?;
    let (reload_cache, after_reload) = cache_point(
        after_reload,
        Some(after_drop),
        &source_context.budget,
        cache_limits,
        Some("drop_images_to_reload_b"),
    )?;
    if after_reload.successful_loads <= after_drop.successful_loads
        || after_reload.evictions <= after_drop.evictions
        || after_reload.bypasses != after_drop.bypasses
        || after_reload.retained_entries != 2
        || after_reload.budget_memory_used != after_a.budget_memory_used
    {
        return Err(
            "pinned row reload did not evict and retain within the released cache slot".into(),
        );
    }
    phases.push(phase(
        "reload_b",
        reload_cache,
        CachePoint::unavailable("pinned row has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    drop(reloaded_b);
    drop(slide);
    drop(view);
    phases.push(phase(
        "drop_view",
        CachePoint::unavailable("source image view was dropped; no public cache owner remains"),
        CachePoint::unavailable("pinned row has no destination editor owner"),
        Some(&source),
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    drop(source);
    phases.push(phase(
        "drop_source",
        CachePoint::unavailable("caller source owner was dropped"),
        CachePoint::unavailable("pinned row has no destination editor owner"),
        None,
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    Ok(NearRow {
        sample_index,
        result: "pinned_normal_bypass_then_clean_reload",
        phases,
        exact_payload_verified: true,
        payload_sha256: Some(crate::sha256_hex(expected_b)),
        payload_bytes: Some(expected_b.len()),
        payload_read_calls: Some(payload_read_calls),
        payload_read_bytes: Some(payload_read_bytes),
        reload_read_calls: Some(reload_read_calls),
        reload_read_bytes: Some(reload_read_bytes),
        accepted_output_bytes: 0,
        refusal_accepted_output_bytes: 0,
        refusal_sink_accepted_bytes: 0,
        refusal_resource: None,
        typed_memory_refusal: false,
        typed_resource_refusal: false,
        eviction_verified: true,
        pinned_bypass_verified: true,
        oversized_bypass_verified: false,
        repeated_output_identities_verified: false,
        publication_output_sha256: Vec::new(),
        root_memory_floor: None,
        metadata_memory_used: None,
        payload_memory_used: None,
        drop_image_memory_used: None,
        leaf_bytes: Some(u64::try_from(expected_b.len())?),
        memory_limit: MEMORY_LIMIT,
        cache_max_bytes: CACHE_LIMIT_BYTES,
        cache_max_entries: 2,
        configured_limits: configured,
        destination_configured_limits: destination_configured,
    })
}

fn run_repeated_publication_iteration(
    corpus: &crate::PptxSourceBackedCrossCopyCorpus,
    sample_index: usize,
) -> Result<NearRow, Box<dyn Error>> {
    let cache_limits = SourceCacheLimits::new(CACHE_LIMIT_BYTES, CACHE_LIMIT_ENTRIES)?;
    let expected = &corpus.source_backed_expected_output;
    let expected_digest = crate::sha256_hex(expected);
    let expected_bytes = u64::try_from(expected.len())?;
    let output_limit = expected_bytes
        .checked_mul(3)
        .ok_or("repeated publication output limit overflows")?;
    let source_context = managed_context("pptx-cache-retention-repeat-source")?;
    let destination_context = managed_context_with_limits(
        "pptx-cache-retention-repeat-destination",
        MEMORY_LIMIT,
        output_limit,
    )?;
    let configured = configured_limits(
        CACHE_LIMIT_BYTES,
        CACHE_LIMIT_ENTRIES,
        MEMORY_LIMIT,
        IO_LIMIT,
    );
    let destination_configured = configured_limits(
        CACHE_LIMIT_BYTES,
        CACHE_LIMIT_ENTRIES,
        MEMORY_LIMIT,
        output_limit,
    );
    let mut phases = Vec::with_capacity(32);
    phases.push(phase(
        "baseline",
        CachePoint::unavailable("repeated source view has not been opened"),
        CachePoint::unavailable("repeated destination editor has not been opened"),
        None,
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    let source = Arc::new(crate::InstrumentedSource::new(
        corpus.source_archive.clone(),
        Vec::new(),
    ));
    let destination = Arc::new(crate::InstrumentedSource::new(
        corpus.destination_archive.clone(),
        Vec::new(),
    ));
    let source_read: Arc<dyn ReadAt> = source.clone();
    let source_view = litchi_pptx::SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
        source_read,
        ReadLimits::default(),
        cache_limits,
        source_context.context.clone(),
    )?;
    if source_view.slide_count() != crate::PPTX_CROSS_COPY_SOURCE_SLIDE_COUNT {
        return Err("repeated publication source slide count changed".into());
    }
    let mut source_diagnostics = source_view.try_cache_diagnostics()?;
    let (source_opened_cache, source_opened_diagnostics) = cache_point(
        source_diagnostics,
        None,
        &source_context.budget,
        cache_limits,
        None,
    )?;
    source_diagnostics = source_opened_diagnostics;
    phases.push(phase(
        "source_opened",
        source_opened_cache,
        CachePoint::unavailable("repeated destination editor has not been opened"),
        Some(&source),
        Some(&destination),
        &source_context.budget,
        &destination_context.budget,
    ));
    let mut publication_hashes = Vec::with_capacity(3);
    for publication_index in 1..=3usize {
        let destination_read: Arc<dyn ReadAt> = destination.clone();
        let editor = litchi_pptx::SourceBackedPresentationEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
            destination_read,
            ReadLimits::default(),
            cache_limits,
            destination_context.context.clone(),
        )?;
        if editor.slide_count() != corpus.destination_slide_count {
            return Err("repeated publication destination slide count changed".into());
        }
        let destination_opened = editor.try_cache_diagnostics()?;
        let (destination_opened_cache, destination_opened) = cache_point(
            destination_opened,
            None,
            &destination_context.budget,
            cache_limits,
            None,
        )?;
        let source_before_plan = source_view.try_cache_diagnostics()?;
        let (source_before_plan_cache, source_before_plan) = cache_point(
            source_before_plan,
            Some(source_diagnostics),
            &source_context.budget,
            cache_limits,
            Some("between_publications_to_opened"),
        )?;
        source_diagnostics = source_before_plan;
        phases.push(phase(
            REPEATED_OPENED_LABELS[publication_index - 1],
            source_before_plan_cache,
            destination_opened_cache,
            Some(&source),
            Some(&destination),
            &source_context.budget,
            &destination_context.budget,
        ));
        let plan = editor.plan_cross_slide_copy(
            &source_view,
            corpus.source_slide,
            corpus.destination_slide,
            corpus.insertion_position,
        )?;
        if plan.source_position() != corpus.source_slide
            || plan.destination_slide_position() != corpus.destination_slide
            || plan.insertion_position() != corpus.insertion_position
            || plan.destination_slide_count() != corpus.destination_slide_count + 1
        {
            return Err("repeated publication plan metadata changed".into());
        }
        let destination_planned = editor.try_cache_diagnostics()?;
        let (destination_planned_cache, _destination_planned) = cache_point(
            destination_planned,
            Some(destination_opened),
            &destination_context.budget,
            cache_limits,
            Some("opened_to_planned"),
        )?;
        let source_planned = source_view.try_cache_diagnostics()?;
        let (source_planned_cache, source_planned) = cache_point(
            source_planned,
            Some(source_diagnostics),
            &source_context.budget,
            cache_limits,
            Some("opened_to_planned"),
        )?;
        source_diagnostics = source_planned;
        phases.push(phase(
            REPEATED_PLANNED_LABELS[publication_index - 1],
            source_planned_cache,
            destination_planned_cache,
            Some(&source),
            Some(&destination),
            &source_context.budget,
            &destination_context.budget,
        ));
        let mut sink = crate::CountingSink::bounded(sink_ceiling(expected.len())?, MAX_WRITE);
        sink.reserve_budget()?;
        let published = editor.publish_cross_slide_copy_to_stream(&mut sink, &plan)?;
        if published.destination_slide_count() != corpus.destination_slide_count + 1
            || published.insertion_position() != corpus.insertion_position
            || published.name() != corpus.source_slide_name
            || sink.bytes != *expected
            || sink.summary().accepted_bytes != expected_bytes
            || sink.summary().largest_write > MAX_WRITE
        {
            return Err("repeated publication differs from exact output oracle".into());
        }
        publication_hashes.push(crate::sha256_hex(&sink.bytes));
        let source_published = source_view.try_cache_diagnostics()?;
        let (source_published_cache, source_published) = cache_point(
            source_published,
            Some(source_diagnostics),
            &source_context.budget,
            cache_limits,
            Some("planned_to_published"),
        )?;
        source_diagnostics = source_published;
        phases.push(phase(
            REPEATED_PUBLISHED_LABELS[publication_index - 1],
            source_published_cache,
            CachePoint::unavailable("destination editor was consumed by publication"),
            Some(&source),
            Some(&destination),
            &source_context.budget,
            &destination_context.budget,
        ));
        drop(published);
        let source_after_result = source_view.try_cache_diagnostics()?;
        let (source_after_result_cache, source_after_result) = cache_point(
            source_after_result,
            Some(source_diagnostics),
            &source_context.budget,
            cache_limits,
            Some("published_to_drop_result"),
        )?;
        source_diagnostics = source_after_result;
        phases.push(phase(
            REPEATED_DROP_RESULT_LABELS[publication_index - 1],
            source_after_result_cache,
            CachePoint::unavailable("destination editor was consumed by publication"),
            Some(&source),
            Some(&destination),
            &source_context.budget,
            &destination_context.budget,
        ));
        drop(plan);
        let source_after_plan = source_view.try_cache_diagnostics()?;
        let (source_after_plan_cache, source_after_plan) = cache_point(
            source_after_plan,
            Some(source_diagnostics),
            &source_context.budget,
            cache_limits,
            Some("drop_result_to_drop_plan"),
        )?;
        source_diagnostics = source_after_plan;
        if destination_context.budget.used(Resource::Memory) != 0
            || destination_context.budget.used(Resource::Objects) != 0
        {
            return Err(
                "repeated publication destination releasable budget did not return to zero".into(),
            );
        }
        phases.push(phase(
            REPEATED_DROP_PLAN_LABELS[publication_index - 1],
            source_after_plan_cache,
            CachePoint::unavailable("destination editor was consumed by publication"),
            Some(&source),
            Some(&destination),
            &source_context.budget,
            &destination_context.budget,
        ));
        drop(sink);
        let source_after_sink = source_view.try_cache_diagnostics()?;
        let (source_after_sink_cache, source_after_sink) = cache_point(
            source_after_sink,
            Some(source_diagnostics),
            &source_context.budget,
            cache_limits,
            Some("drop_plan_to_drop_sink"),
        )?;
        source_diagnostics = source_after_sink;
        phases.push(phase(
            REPEATED_DROP_SINK_LABELS[publication_index - 1],
            source_after_sink_cache,
            CachePoint::unavailable("destination editor was consumed by publication"),
            Some(&source),
            Some(&destination),
            &source_context.budget,
            &destination_context.budget,
        ));
        let expected_output_used = expected_bytes
            .checked_mul(u64::try_from(publication_index)?)
            .ok_or("repeated publication output usage overflows")?;
        if destination_context.budget.used(Resource::OutputBytes) != expected_output_used
            || expected_output_used > output_limit
        {
            return Err(
                "repeated publication output budget does not match cumulative boundary".into(),
            );
        }
    }

    // The fourth operation is intentionally outside the three-success sample
    // count.  It proves the cumulative OutputBytes ceiling when the public
    // publisher performs its preflight reservation before writing.
    let destination_read: Arc<dyn ReadAt> = destination.clone();
    let editor = litchi_pptx::SourceBackedPresentationEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
        destination_read,
        ReadLimits::default(),
        cache_limits,
        destination_context.context.clone(),
    )?;
    let plan = editor.plan_cross_slide_copy(
        &source_view,
        corpus.source_slide,
        corpus.destination_slide,
        corpus.insertion_position,
    )?;
    let mut sink = crate::CountingSink::bounded(sink_ceiling(expected.len())?, MAX_WRITE);
    sink.reserve_budget()?;
    let fourth = editor.publish_cross_slide_copy_to_stream(&mut sink, &plan);
    let typed_output_refusal = match fourth {
        Ok(_) => false,
        Err(error) => output_refusal(&error),
    };
    if !typed_output_refusal || !sink.bytes.is_empty() {
        return Err(
            "repeated publication fourth probe did not preserve typed zero-output refusal".into(),
        );
    }
    drop(plan);
    drop(sink);
    if destination_context.budget.used(Resource::Memory) != 0
        || destination_context.budget.used(Resource::Objects) != 0
        || destination_context.budget.used(Resource::OutputBytes) != output_limit
    {
        return Err("repeated publication fourth refusal changed the cumulative destination budget unexpectedly".into());
    }
    let source_after_fourth = source_view.try_cache_diagnostics()?;
    let (source_after_fourth_cache, _source_after_fourth) = cache_point(
        source_after_fourth,
        Some(source_diagnostics),
        &source_context.budget,
        cache_limits,
        Some("drop_sink_3_to_fourth_refusal"),
    )?;
    phases.push(phase(
        "fourth_refusal",
        source_after_fourth_cache,
        CachePoint::unavailable("destination editor was consumed by the refused publication"),
        Some(&source),
        Some(&destination),
        &source_context.budget,
        &destination_context.budget,
    ));
    drop(source_view);
    phases.push(phase(
        "drop_view",
        CachePoint::unavailable("source view was dropped; no public source cache owner remains"),
        CachePoint::unavailable("destination editor was consumed by publication"),
        Some(&source),
        Some(&destination),
        &source_context.budget,
        &destination_context.budget,
    ));
    drop(source);
    drop(destination);
    phases.push(phase(
        "drop_sources",
        CachePoint::unavailable("caller source owners were dropped"),
        CachePoint::unavailable("destination editor was consumed by publication"),
        None,
        None,
        &source_context.budget,
        &destination_context.budget,
    ));
    if publication_hashes
        .iter()
        .any(|hash| hash != &expected_digest)
    {
        return Err("repeated publication output identity changed across iterations".into());
    }
    Ok(NearRow {
        sample_index,
        result: "three_exact_publications_and_fourth_output_refusal",
        phases,
        exact_payload_verified: true,
        payload_sha256: None,
        payload_bytes: None,
        payload_read_calls: None,
        payload_read_bytes: None,
        reload_read_calls: None,
        reload_read_bytes: None,
        accepted_output_bytes: output_limit,
        refusal_accepted_output_bytes: 0,
        refusal_sink_accepted_bytes: 0,
        refusal_resource: Some("OutputBytes"),
        typed_memory_refusal: false,
        typed_resource_refusal: typed_output_refusal,
        eviction_verified: false,
        pinned_bypass_verified: false,
        oversized_bypass_verified: false,
        repeated_output_identities_verified: true,
        publication_output_sha256: publication_hashes,
        root_memory_floor: None,
        metadata_memory_used: None,
        payload_memory_used: None,
        drop_image_memory_used: None,
        leaf_bytes: None,
        memory_limit: MEMORY_LIMIT,
        cache_max_bytes: CACHE_LIMIT_BYTES,
        cache_max_entries: CACHE_LIMIT_ENTRIES,
        configured_limits: configured,
        destination_configured_limits: destination_configured,
    })
}

enum NearPrepared {
    Image(Box<NearCorpus>),
    Repeated {
        lifecycle: Box<crate::PptxSourceBackedCrossCopyCorpus>,
        gates: crate::PptxSourceBackedCrossCopyLifecycleGateSummary,
    },
}

fn run_near_capture(config: &Config) -> Result<NearReport, Box<dyn Error>> {
    // Build and validate the immutable corpus before calibration and before
    // any warmup/retained iteration.  Rebuilding it inside the loop would
    // charge corpus generation and gate work to each observation boundary.
    let prepared = if config.scenario == Scenario::RepeatedPublication {
        let case = match config.corpus {
            CorpusKind::Plain => crate::Case::PptxSourceBackedCrossCopyPlainLifecycle,
            CorpusKind::MediaRich => crate::Case::PptxSourceBackedCrossCopyMediaRichLifecycle,
        };
        let lifecycle = crate::build_pptx_source_backed_cross_copy_corpus(case)?;
        let gates = validate_gates(&lifecycle)?;
        NearPrepared::Repeated {
            lifecycle: Box::new(lifecycle),
            gates,
        }
    } else {
        NearPrepared::Image(Box::new(build_near_corpus()?))
    };

    let (
        source_archive,
        destination_archive,
        expected_output,
        manifest,
        gates,
        expected_image_sha256,
        expected_image_bytes,
    ) = match &prepared {
        NearPrepared::Image(near) => (
            &near.images.archive,
            &near.lifecycle.destination_archive,
            &near.lifecycle.source_backed_expected_output,
            &near.images.manifest,
            &near.gates,
            Some(near.images.selected_payload_sha256.clone()),
            Some(near.images.selected_payload.len()),
        ),
        NearPrepared::Repeated { lifecycle, gates } => (
            &lifecycle.source_archive,
            &lifecycle.destination_archive,
            &lifecycle.source_backed_expected_output,
            &lifecycle.manifest,
            gates,
            None,
            None,
        ),
    };

    let (configured, destination_configured, root_memory, metadata_memory, leaf_bytes) =
        match &prepared {
            NearPrepared::Repeated { .. } => {
                let output_limit = u64::try_from(expected_output.len())?
                    .checked_mul(3)
                    .ok_or("repeated publication output limit overflows")?;
                (
                    configured_limits(
                        CACHE_LIMIT_BYTES,
                        CACHE_LIMIT_ENTRIES,
                        MEMORY_LIMIT,
                        IO_LIMIT,
                    ),
                    configured_limits(
                        CACHE_LIMIT_BYTES,
                        CACHE_LIMIT_ENTRIES,
                        MEMORY_LIMIT,
                        output_limit,
                    ),
                    0,
                    0,
                    0,
                )
            },
            NearPrepared::Image(near) => {
                let cache_limits = SourceCacheLimits::new(CACHE_LIMIT_BYTES, CACHE_LIMIT_ENTRIES)?;
                let (root_memory, metadata_memory, leaf_bytes) =
                    calibrate_image_memory(&near.images, cache_limits)?;
                let (cache_bytes, entries, memory) = match config.scenario {
                    Scenario::ExactAdmission => (
                        CACHE_LIMIT_BYTES,
                        CACHE_LIMIT_ENTRIES,
                        root_memory
                            .checked_add(leaf_bytes)
                            .ok_or("exact image memory limit overflows")?,
                    ),
                    Scenario::OneUnder => (
                        CACHE_LIMIT_BYTES,
                        CACHE_LIMIT_ENTRIES,
                        root_memory
                            .checked_add(leaf_bytes)
                            .and_then(|value| value.checked_sub(1))
                            .ok_or("one-under image memory limit underflows")?,
                    ),
                    Scenario::PinnedEviction => (CACHE_LIMIT_BYTES, 2, MEMORY_LIMIT),
                    Scenario::OversizedBypass => (
                        leaf_bytes
                            .checked_sub(1)
                            .and_then(|value| usize::try_from(value).ok())
                            .ok_or("oversized image cache limit underflows")?,
                        CACHE_LIMIT_ENTRIES,
                        MEMORY_LIMIT,
                    ),
                    _ => return Err("invalid near image scenario".into()),
                };
                (
                    configured_limits(cache_bytes, entries, memory, IO_LIMIT),
                    configured_limits(
                        CACHE_LIMIT_BYTES,
                        CACHE_LIMIT_ENTRIES,
                        MEMORY_LIMIT,
                        IO_LIMIT,
                    ),
                    root_memory,
                    metadata_memory,
                    leaf_bytes,
                )
            },
        };
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or("warmup and samples overflow usize")?;
    let mut rows = Vec::with_capacity(config.samples);
    for iteration in 0..total {
        let row = match config.scenario {
            Scenario::ExactAdmission | Scenario::OneUnder | Scenario::OversizedBypass => {
                let NearPrepared::Image(near) = &prepared else {
                    return Err("image boundary scenario has no image corpus".into());
                };
                run_image_boundary_iteration(
                    &near.images,
                    config.scenario,
                    iteration,
                    root_memory,
                    metadata_memory,
                    leaf_bytes,
                )?
            },
            Scenario::PinnedEviction => {
                let NearPrepared::Image(near) = &prepared else {
                    return Err("pinned eviction scenario has no image corpus".into());
                };
                run_pinned_eviction_iteration(&near.images, iteration, &near.pinned_a_payload)?
            },
            Scenario::RepeatedPublication => {
                let NearPrepared::Repeated { lifecycle, .. } = &prepared else {
                    return Err("repeated publication scenario has no lifecycle corpus".into());
                };
                run_repeated_publication_iteration(lifecycle, iteration)?
            },
            Scenario::Lifecycle => {
                return Err("lifecycle must use the lifecycle capture path".into());
            },
        };
        if iteration >= config.warmup {
            rows.push(NearRow {
                sample_index: iteration - config.warmup,
                ..row
            });
        }
    }
    let binary = crate::current_executable_identity()?;
    Ok(NearReport {
        schema: SCHEMA,
        scenario: config.scenario.name(),
        corpus: config.corpus.name(),
        cache_scope: CACHE_SCOPE,
        budget_scope: BUDGET_SCOPE,
        source_io_scope: SOURCE_IO_SCOPE,
        publication_scope: PUBLICATION_SCOPE,
        rss_scope: RSS_SCOPE,
        samples: config.samples,
        warmup: config.warmup,
        checked_iteration_count: total,
        source_revision: config.source_revision.clone(),
        binary_sha256: binary.binary_sha256.clone(),
        binary_bytes: binary.binary_bytes,
        current_exe: binary.path.clone(),
        source_archive_sha256: crate::sha256_hex(source_archive),
        source_archive_bytes: source_archive.len(),
        destination_archive_sha256: crate::sha256_hex(destination_archive),
        destination_archive_bytes: destination_archive.len(),
        expected_output_sha256: crate::sha256_hex(expected_output),
        expected_output_bytes: expected_output.len(),
        expected_image_sha256,
        expected_image_bytes,
        corpus_manifest: manifest.clone(),
        gates: gates.clone(),
        configured_limits: configured,
        destination_configured_limits: destination_configured,
        phases: NEAR_PHASES.to_vec(),
        samples_raw: rows,
    })
}

fn run_lifecycle_iteration(
    corpus: &crate::PptxSourceBackedCrossCopyCorpus,
    sample_index: usize,
) -> Result<LifecycleRow, Box<dyn Error>> {
    let source_context = managed_context("pptx-cache-retention-source")?;
    let destination_context = managed_context("pptx-cache-retention-destination")?;
    let cache_limits = SourceCacheLimits::new(CACHE_LIMIT_BYTES, CACHE_LIMIT_ENTRIES)?;
    let expected_digest = crate::sha256_hex(&corpus.source_backed_expected_output);
    let mut phases = Vec::with_capacity(9);

    phases.push(phase(
        "baseline",
        CachePoint::unavailable("source view has not been opened"),
        CachePoint::unavailable("destination editor has not been opened"),
        None,
        None,
        &source_context.budget,
        &destination_context.budget,
    ));

    let source = Arc::new(crate::InstrumentedSource::new(
        corpus.source_archive.clone(),
        Vec::new(),
    ));
    let destination = Arc::new(crate::InstrumentedSource::new(
        corpus.destination_archive.clone(),
        Vec::new(),
    ));
    let source_read: Arc<dyn ReadAt> = source.clone();
    let destination_read: Arc<dyn ReadAt> = destination.clone();
    let mut sink = crate::CountingSink::bounded(
        sink_ceiling(corpus.source_backed_expected_output.len())?,
        MAX_WRITE,
    );
    sink.reserve_budget()?;

    let source_view = litchi_pptx::SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
        source_read,
        ReadLimits::default(),
        cache_limits,
        source_context.context.clone(),
    )?;
    let editor = litchi_pptx::SourceBackedPresentationEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
        destination_read,
        ReadLimits::default(),
        cache_limits,
        destination_context.context.clone(),
    )?;
    if source_view.slide_count() != crate::PPTX_CROSS_COPY_SOURCE_SLIDE_COUNT
        || editor.slide_count() != corpus.destination_slide_count
    {
        return Err("managed PPTX lifecycle opening changed prevalidated slide counts".into());
    }
    let source_opened = source_view.try_cache_diagnostics()?;
    let destination_opened = editor.try_cache_diagnostics()?;
    let (source_opened_point, source_opened) = cache_point(
        source_opened,
        None,
        &source_context.budget,
        cache_limits,
        None,
    )?;
    let (destination_opened_point, destination_opened) = cache_point(
        destination_opened,
        None,
        &destination_context.budget,
        cache_limits,
        None,
    )?;
    phases.push(phase(
        "opened",
        source_opened_point,
        destination_opened_point,
        Some(&source),
        Some(&destination),
        &source_context.budget,
        &destination_context.budget,
    ));

    let plan = editor.plan_cross_slide_copy(
        &source_view,
        corpus.source_slide,
        corpus.destination_slide,
        corpus.insertion_position,
    )?;
    if plan.source_position() != corpus.source_slide
        || plan.destination_slide_position() != corpus.destination_slide
        || plan.insertion_position() != corpus.insertion_position
        || plan.destination_slide_count() != corpus.destination_slide_count + 1
    {
        return Err("managed PPTX lifecycle plan metadata is not deterministic".into());
    }
    let source_planned = source_view.try_cache_diagnostics()?;
    let destination_planned = editor.try_cache_diagnostics()?;
    let (source_planned_point, source_planned) = cache_point(
        source_planned,
        Some(source_opened),
        &source_context.budget,
        cache_limits,
        Some("opened_to_planned"),
    )?;
    let (destination_planned_point, _destination_planned) = cache_point(
        destination_planned,
        Some(destination_opened),
        &destination_context.budget,
        cache_limits,
        Some("opened_to_planned"),
    )?;
    phases.push(phase(
        "planned",
        source_planned_point,
        destination_planned_point,
        Some(&source),
        Some(&destination),
        &source_context.budget,
        &destination_context.budget,
    ));

    let published = editor.publish_cross_slide_copy_to_stream(&mut sink, &plan)?;
    if published.destination_slide_count() != corpus.destination_slide_count + 1
        || published.insertion_position() != corpus.insertion_position
        || published.name() != corpus.source_slide_name
        || sink.bytes != corpus.source_backed_expected_output
        || sink.summary().accepted_bytes
            != u64::try_from(corpus.source_backed_expected_output.len())?
        || sink.summary().largest_write > MAX_WRITE
    {
        return Err("managed PPTX lifecycle publication differs from exact output oracle".into());
    }
    let source_published = source_view.try_cache_diagnostics()?;
    let (source_published_point, source_published) = cache_point(
        source_published,
        Some(source_planned),
        &source_context.budget,
        cache_limits,
        Some("planned_to_published"),
    )?;
    phases.push(phase(
        "published",
        source_published_point,
        CachePoint::unavailable("destination editor was consumed by publication"),
        Some(&source),
        Some(&destination),
        &source_context.budget,
        &destination_context.budget,
    ));

    drop(published);
    let source_after_result = source_view.try_cache_diagnostics()?;
    let (source_after_result_point, source_after_result) = cache_point(
        source_after_result,
        Some(source_published),
        &source_context.budget,
        cache_limits,
        Some("published_to_drop_result"),
    )?;
    phases.push(phase(
        "drop_result",
        source_after_result_point,
        CachePoint::unavailable("destination editor was consumed by publication"),
        Some(&source),
        Some(&destination),
        &source_context.budget,
        &destination_context.budget,
    ));

    drop(plan);
    let source_after_plan = source_view.try_cache_diagnostics()?;
    let (source_after_plan_point, _source_after_plan) = cache_point(
        source_after_plan,
        Some(source_after_result),
        &source_context.budget,
        cache_limits,
        Some("drop_result_to_drop_plan"),
    )?;
    phases.push(phase(
        "drop_plan",
        source_after_plan_point,
        CachePoint::unavailable("destination editor was consumed by publication"),
        Some(&source),
        Some(&destination),
        &source_context.budget,
        &destination_context.budget,
    ));

    drop(source_view);
    phases.push(phase(
        "drop_view",
        CachePoint::unavailable("source view handle was dropped; no public cache owner remains"),
        CachePoint::unavailable("destination editor was consumed by publication"),
        Some(&source),
        Some(&destination),
        &source_context.budget,
        &destination_context.budget,
    ));

    drop(source);
    drop(destination);
    phases.push(phase(
        "drop_caller_sources",
        CachePoint::unavailable("caller source owners were dropped"),
        CachePoint::unavailable("destination editor was consumed by publication"),
        None,
        None,
        &source_context.budget,
        &destination_context.budget,
    ));

    drop(sink);
    phases.push(phase(
        "drop_sink",
        CachePoint::unavailable("caller source owners were dropped"),
        CachePoint::unavailable("destination editor was consumed by publication"),
        None,
        None,
        &source_context.budget,
        &destination_context.budget,
    ));

    Ok(LifecycleRow {
        sample_index,
        exact_output_verified: true,
        output_sha256: expected_digest,
        output_bytes: corpus.source_backed_expected_output.len(),
        phases,
    })
}

fn run_capture(config: &Config) -> Result<Report, Box<dyn Error>> {
    if config.scenario != Scenario::Lifecycle {
        return Err(format!(
            "cache-retention scenario {} is reserved for the near-image protocol tranche",
            config.scenario.name()
        )
        .into());
    }
    let case = match config.corpus {
        CorpusKind::Plain => crate::Case::PptxSourceBackedCrossCopyPlainLifecycle,
        CorpusKind::MediaRich => crate::Case::PptxSourceBackedCrossCopyMediaRichLifecycle,
    };
    let corpus = crate::build_pptx_source_backed_cross_copy_corpus(case)?;
    let gates = validate_gates(&corpus)?;
    let expected_digest = crate::sha256_hex(&corpus.source_backed_expected_output);
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or("warmup and samples overflow usize")?;
    let mut rows = Vec::with_capacity(config.samples);
    for iteration in 0..total {
        let row = run_lifecycle_iteration(&corpus, iteration)?;
        if iteration >= config.warmup {
            rows.push(LifecycleRow {
                sample_index: iteration - config.warmup,
                ..row
            });
        }
    }
    let binary = crate::current_executable_identity()?;
    Ok(Report {
        schema: SCHEMA,
        scenario: config.scenario.name(),
        corpus: config.corpus.name(),
        cache_scope: CACHE_SCOPE,
        budget_scope: BUDGET_SCOPE,
        source_io_scope: SOURCE_IO_SCOPE,
        publication_scope: PUBLICATION_SCOPE,
        rss_scope: RSS_SCOPE,
        samples: config.samples,
        warmup: config.warmup,
        checked_iteration_count: total,
        source_revision: config.source_revision.clone(),
        binary_sha256: binary.binary_sha256.clone(),
        binary_bytes: binary.binary_bytes,
        current_exe: binary.path.clone(),
        source_archive_sha256: crate::sha256_hex(&corpus.source_archive),
        source_archive_bytes: corpus.source_archive.len(),
        destination_archive_sha256: crate::sha256_hex(&corpus.destination_archive),
        destination_archive_bytes: corpus.destination_archive.len(),
        expected_output_sha256: expected_digest,
        expected_output_bytes: corpus.source_backed_expected_output.len(),
        corpus_manifest: corpus.manifest.clone(),
        gates,
        configured_limits: ConfiguredLimits {
            cache_max_bytes: CACHE_LIMIT_BYTES,
            cache_max_entries: CACHE_LIMIT_ENTRIES,
            memory_limit: MEMORY_LIMIT,
            input_bytes_limit: IO_LIMIT,
            output_bytes_limit: IO_LIMIT,
            work_limit: IO_LIMIT,
            objects_limit: OBJECT_LIMIT,
            depth_limit: DEPTH_LIMIT,
        },
        destination_configured_limits: ConfiguredLimits {
            cache_max_bytes: CACHE_LIMIT_BYTES,
            cache_max_entries: CACHE_LIMIT_ENTRIES,
            memory_limit: MEMORY_LIMIT,
            input_bytes_limit: IO_LIMIT,
            output_bytes_limit: IO_LIMIT,
            work_limit: IO_LIMIT,
            objects_limit: OBJECT_LIMIT,
            depth_limit: DEPTH_LIMIT,
        },
        destination_editor_consumed_during_publish: true,
        phases: PHASES.to_vec(),
        samples_raw: rows,
    })
}

/// Run the managed cache-retention target from arguments after `main` has
/// consumed the `cache-retention` selector.  Reports use `create_new` output
/// semantics so an existing evidence file can never be silently replaced.
pub fn run_from_args<I>(args: I) -> Result<(), Box<dyn Error>>
where
    I: IntoIterator<Item = OsString>,
{
    let args = args.into_iter().collect::<Vec<_>>();
    let config = parse_config(&args)?;
    let bytes = if config.scenario == Scenario::Lifecycle {
        serde_json::to_vec_pretty(&run_capture(&config)?)?
    } else {
        serde_json::to_vec_pretty(&run_near_capture(&config)?)?
    };
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&config.output)?;
    output.write_all(&bytes)?;
    output.write_all(b"\n")?;
    output.flush()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn args(values: &[&str]) -> Vec<OsString> {
        values.iter().map(OsString::from).collect()
    }

    #[test]
    fn config_accepts_lifecycle_both_corpora_and_rejects_near_plain() {
        let base = [
            "--scenario",
            "lifecycle",
            "--corpus",
            "plain",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789abcdef0123456789abcdef01234567",
            "--output",
            "capture.json",
        ];
        assert!(parse_config(&args(&base)).is_ok());
        let mut near = base.map(str::to_owned).into_iter().collect::<Vec<_>>();
        near[1] = "exact-admission".to_owned();
        assert!(parse_config(&near.into_iter().map(OsString::from).collect::<Vec<_>>()).is_err());
    }

    #[test]
    fn config_accepts_near_media_and_repeated_plain() {
        let base = [
            "--scenario",
            "exact-admission",
            "--corpus",
            "media-rich",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789abcdef0123456789abcdef01234567",
            "--output",
            "capture.json",
        ];
        assert!(parse_config(&args(&base)).is_ok());
        let mut repeated = base.map(str::to_owned).into_iter().collect::<Vec<_>>();
        repeated[1] = "repeated-publication".to_owned();
        repeated[3] = "plain".to_owned();
        assert!(
            parse_config(&repeated.into_iter().map(OsString::from).collect::<Vec<_>>()).is_ok()
        );
    }

    #[test]
    fn config_rejects_non_lowercase_revision_duplicates_and_bounds() {
        let valid = args(&[
            "--scenario",
            "lifecycle",
            "--corpus",
            "media-rich",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789abcdef0123456789abcdef01234567",
            "--output",
            "capture.json",
        ]);
        assert!(parse_config(&valid).is_ok());
        let mut uppercase = valid.clone();
        uppercase[9] = OsString::from("0123456789ABCDEF0123456789abcdef01234567");
        assert!(parse_config(&uppercase).is_err());
        let mut duplicate = valid.clone();
        duplicate.extend(args(&["--samples", "2"]));
        assert!(parse_config(&duplicate).is_err());
        let mut too_many = valid;
        too_many[5] = OsString::from("1001");
        assert!(parse_config(&too_many).is_err());
    }

    #[test]
    fn unavailable_cache_points_do_not_fabricate_counters() {
        let point = CachePoint::unavailable("owner dropped");
        assert_eq!(point.availability, "unavailable");
        assert_eq!(point.hits, None);
        assert_eq!(point.retained_bytes, None);
        assert!(point.event_delta.is_none());
    }

    #[test]
    fn event_interval_uses_checked_counter_delta() {
        let before = SourceCacheDiagnostics::default();
        let mut after = before;
        after.hits = 3;
        after.successful_loads = 2;
        let delta = SourceCacheDiagnostics::checked_counter_delta(before, after).unwrap();
        let event = EventDelta::from(delta);
        assert_eq!(event.hits, 3);
        assert_eq!(event.successful_loads, 2);
        let mut regressed = after;
        regressed.hits = 1;
        assert!(SourceCacheDiagnostics::checked_counter_delta(after, regressed).is_err());
    }
}
