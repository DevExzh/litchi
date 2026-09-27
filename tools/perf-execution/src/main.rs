//! Reproducible process benchmark for the three explicit ADR 0031 read routes.
//!
//! This binary intentionally owns its corpus and measurement lifecycle.  It
//! does not add a production dependency, install a runtime, or instrument the
//! normal source path with counters.

#![forbid(unsafe_code)]

use litchi_cfb::{OleWriter, SharedOleBulkRead, SharedOleFile, SharedOleFileLimits};
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, Resource,
    SourceVersion,
};
use litchi_opc::{OpenSession, PackURI, ReadLimits, SourceBackedPackage};
use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;
use std::fmt;
use std::io::Cursor;
use std::num::{NonZeroU64, NonZeroUsize};
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

const SCHEMA: &str = "litchi.execution-baseline.v1";
const MEMBER_COUNT: usize = 32;
const SMALL_BYTES: usize = 4 * 1024;
const LARGE_BYTES: usize = 256 * 1024;
const MIN_PARALLEL_BYTES: u64 = 64 * 1024;
const MAX_WORKERS: usize = MEMBER_COUNT;
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 10_000;
const CPU_TASK_LIMIT: u64 = 1_000_000;
const MAX_IN_FLIGHT_BYTES: u64 = 16 * 1024 * 1024;
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const MEMBER_REL_TYPE: &str = "https://litchi.invalid/performance/member";

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "lowercase")]
enum Route {
    Opc,
    Cfb,
    Parts,
}

impl std::str::FromStr for Route {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "opc" => Ok(Self::Opc),
            "cfb" => Ok(Self::Cfb),
            "parts" => Ok(Self::Parts),
            _ => Err(format!("--route must be opc, cfb, or parts; got {value:?}")),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "lowercase")]
enum Shape {
    Small,
    Large,
    Mixed,
}

impl Shape {
    const fn member_size(self, index: usize) -> usize {
        match self {
            Self::Small => SMALL_BYTES,
            Self::Large => LARGE_BYTES,
            // Keep the small member last.  The requested order is part of the
            // corpus manifest and makes mixed-shape verification unambiguous.
            Self::Mixed if index + 1 == MEMBER_COUNT => SMALL_BYTES,
            Self::Mixed => LARGE_BYTES,
        }
    }
}

impl std::str::FromStr for Shape {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "small" => Ok(Self::Small),
            "large" => Ok(Self::Large),
            "mixed" => Ok(Self::Mixed),
            _ => Err(format!(
                "--shape must be small, large, or mixed; got {value:?}"
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "lowercase")]
enum State {
    Fresh,
    Primed,
}

impl std::str::FromStr for State {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "fresh" => Ok(Self::Fresh),
            "primed" => Ok(Self::Primed),
            _ => Err(format!("--state must be fresh or primed; got {value:?}")),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Config {
    route: Route,
    shape: Shape,
    workers: usize,
    task_floor: u64,
    state: State,
    samples: usize,
    warmup: usize,
    output: PathBuf,
}

#[derive(Debug, Clone, Serialize)]
struct Corpus {
    shape: Shape,
    metadata_members: Vec<String>,
    members: Vec<MemberManifest>,
    selected_payload_member_count: usize,
    opc_sha256: String,
    cfb_sha256: String,
    #[serde(skip)]
    opc_bytes: Vec<u8>,
    #[serde(skip)]
    cfb_bytes: Vec<u8>,
    #[serde(skip)]
    payloads: Vec<Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
struct MemberManifest {
    index: usize,
    opc_name: String,
    opc_uri: String,
    cfb_name: String,
    bytes: usize,
    sha256: String,
}

#[derive(Debug, Clone, Serialize)]
struct Scope {
    ingress: String,
    source: String,
    source_metrics_claim: String,
    timing: String,
    cpu_time: String,
    claims_excluded: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
struct ReportConfig {
    route: Route,
    shape: Shape,
    workers: usize,
    task_floor: u64,
    aggregate_parallel_bytes: u64,
    state: State,
    samples: usize,
    warmup: usize,
    cpu_task_limit: u64,
}

#[derive(Debug, Clone, Copy, Serialize)]
struct ResourceSnapshot {
    workers: u64,
    io_concurrency: u64,
    cpu_tasks: u64,
}

#[derive(Debug, Clone, Copy, Serialize)]
struct ResourceLimits {
    memory: u64,
    input_bytes: u64,
    output_bytes: u64,
    objects: u64,
    depth: u64,
    work: u64,
    workers: u64,
    io_concurrency: u64,
    cpu_tasks: u64,
    max_in_flight_tasks: u64,
    max_in_flight_bytes: u64,
}

#[derive(Debug, Clone, Serialize)]
struct ResourceReport {
    limits: ResourceLimits,
    before_operation: ResourceSnapshot,
    after_operation: ResourceSnapshot,
    after_drop: ResourceSnapshot,
    worker_and_io_released: bool,
    cpu_tasks_within_limit: bool,
}

#[derive(Debug, Clone, Serialize)]
struct ByteVerification {
    ordered: bool,
    all_member_sha256_match: bool,
    members: usize,
    logical_bytes: u64,
    sequence_sha256: String,
}

#[derive(Debug, Clone, Serialize)]
struct SourceMetricsReport {
    availability: String,
    logical_calls: Option<u64>,
    requested_bytes: Option<u64>,
    returned_bytes: Option<u64>,
    short_reads: Option<u64>,
    active_reads_after_operation: Option<u64>,
    max_simultaneous_reads: Option<u64>,
    request_size_histogram: Option<Vec<u64>>,
}

#[derive(Debug, Clone, Serialize)]
struct SampleReport {
    sample: usize,
    wall_ns: u128,
    cpu_ns: Option<u64>,
    primed_cache_hit_control: bool,
    verification: ByteVerification,
    resources: ResourceReport,
    source_metrics: SourceMetricsReport,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    scope: Scope,
    config: ReportConfig,
    corpus: Corpus,
    metrics: MetricsAvailability,
    samples: Vec<SampleReport>,
}

#[derive(Debug, Serialize)]
struct MetricsAvailability {
    source_metrics_feature: bool,
    source_metrics_scope: String,
    opc_external_read_at: String,
    cpu_clock: String,
}

#[derive(Debug)]
struct BenchmarkError(String);

impl fmt::Display for BenchmarkError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl std::error::Error for BenchmarkError {}

type AnyResult<T> = Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn boxed_error(message: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    Box::new(BenchmarkError(message.into()))
}

fn parse_usize(name: &str, value: &str, maximum: usize) -> Result<usize, String> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("{name} must be an unsigned integer; got {value:?}"))?;
    if parsed == 0 || parsed > maximum {
        return Err(format!(
            "{name} must be in the inclusive range 1..={maximum}; got {parsed}"
        ));
    }
    Ok(parsed)
}

fn parse_task_floor(value: &str) -> Result<u64, String> {
    let parsed = value
        .parse::<u64>()
        .map_err(|_| format!("--task-floor must be an unsigned integer; got {value:?}"))?;
    if parsed > MAX_IN_FLIGHT_BYTES {
        return Err(format!(
            "--task-floor must not exceed {MAX_IN_FLIGHT_BYTES}; got {parsed}"
        ));
    }
    Ok(parsed)
}

fn usage() -> &'static str {
    "usage: litchi-perf-execution --route opc|cfb|parts --shape small|large|mixed --workers N --task-floor BYTES --state fresh|primed --samples N --warmup N --output PATH"
}

fn parse_args_from<I>(args: I) -> Result<Config, String>
where
    I: IntoIterator<Item = String>,
{
    let mut route = None;
    let mut shape = None;
    let mut workers = None;
    let mut task_floor = None;
    let mut state = None;
    let mut samples = None;
    let mut warmup = None;
    let mut output = None;
    let mut args = args.into_iter();
    while let Some(flag) = args.next() {
        if flag == "--help" || flag == "-h" {
            return Err(usage().to_string());
        }
        let value = args
            .next()
            .ok_or_else(|| format!("missing value after {flag}; {}", usage()))?;
        match flag.as_str() {
            "--route" if route.is_none() => route = Some(value.parse()?),
            "--shape" if shape.is_none() => shape = Some(value.parse()?),
            "--workers" if workers.is_none() => {
                workers = Some(parse_usize("--workers", &value, MAX_WORKERS)?)
            },
            "--task-floor" if task_floor.is_none() => task_floor = Some(parse_task_floor(&value)?),
            "--state" if state.is_none() => state = Some(value.parse()?),
            "--samples" if samples.is_none() => {
                samples = Some(parse_usize("--samples", &value, MAX_SAMPLES)?)
            },
            "--warmup" if warmup.is_none() => {
                let parsed = value
                    .parse::<usize>()
                    .map_err(|_| format!("--warmup must be an unsigned integer; got {value:?}"))?;
                if parsed > MAX_WARMUP {
                    return Err(format!(
                        "--warmup must not exceed {MAX_WARMUP}; got {parsed}"
                    ));
                }
                warmup = Some(parsed);
            },
            "--output" if output.is_none() => {
                if value.is_empty() {
                    return Err("--output must not be empty".to_string());
                }
                output = Some(PathBuf::from(value));
            },
            _ => return Err(format!("unknown or duplicate option {flag}; {}", usage())),
        }
    }
    Ok(Config {
        route: route.ok_or_else(|| "missing --route".to_string())?,
        shape: shape.ok_or_else(|| "missing --shape".to_string())?,
        workers: workers.ok_or_else(|| "missing --workers".to_string())?,
        task_floor: task_floor.ok_or_else(|| "missing --task-floor".to_string())?,
        state: state.ok_or_else(|| "missing --state".to_string())?,
        samples: samples.ok_or_else(|| "missing --samples".to_string())?,
        warmup: warmup.ok_or_else(|| "missing --warmup".to_string())?,
        output: output.ok_or_else(|| "missing --output".to_string())?,
    })
}

fn parse_args() -> Result<Config, String> {
    parse_args_from(std::env::args().skip(1))
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn payload_for(shape: Shape, index: usize) -> Vec<u8> {
    let size = shape.member_size(index);
    let mut payload = vec![0_u8; size];
    let label = format!("litchi-0786-member-{index:02}-");
    for (position, byte) in payload.iter_mut().enumerate() {
        let lane = position % 97;
        let label_byte = label.as_bytes()[position % label.len()];
        // Repeated, distinct lanes are deliberately compressible while the
        // member index changes every payload's byte sequence and digest.
        *byte = label_byte
            .wrapping_add(u8::try_from(lane).unwrap_or(0).wrapping_mul(3))
            .wrapping_add(u8::try_from(index * 11).unwrap_or(0));
    }
    payload
}

fn member_manifest(index: usize, shape: Shape, payload: &[u8]) -> MemberManifest {
    let opc_name = format!("custom/member{index:02}.bin");
    MemberManifest {
        index,
        opc_uri: format!("/{opc_name}"),
        cfb_name: format!("Member{index:02}"),
        bytes: shape.member_size(index),
        sha256: sha256_hex(payload),
        opc_name,
    }
}

fn build_opc(payloads: &[Vec<u8>], manifests: &[MemberManifest]) -> AnyResult<Vec<u8>> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#
    );
    let mut root_relationships = format!(r#"<Relationships xmlns="{RELATIONSHIPS_NS}">"#);
    for manifest in manifests {
        root_relationships.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{MEMBER_REL_TYPE}" Target="{}"/>"#,
            manifest.index, manifest.opc_name
        ));
    }
    root_relationships.push_str("</Relationships>");

    let mut writer = StreamingArchiveWriter::new();
    writer.write_deflated("[Content_Types].xml", content_types.as_bytes())?;
    writer.write_deflated("_rels/.rels", root_relationships.as_bytes())?;
    for (payload, manifest) in payloads.iter().zip(manifests) {
        writer.write_deflated(&manifest.opc_name, payload)?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn build_cfb(payloads: &[Vec<u8>], manifests: &[MemberManifest]) -> AnyResult<Vec<u8>> {
    let mut writer = OleWriter::new();
    for (payload, manifest) in payloads.iter().zip(manifests) {
        let path = [manifest.cfb_name.as_str()];
        writer.create_stream_owned(&path, payload.clone())?;
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn build_corpus(shape: Shape) -> AnyResult<Corpus> {
    let payloads: Vec<Vec<u8>> = (0..MEMBER_COUNT)
        .map(|index| payload_for(shape, index))
        .collect();
    let manifests: Vec<MemberManifest> = payloads
        .iter()
        .enumerate()
        .map(|(index, payload)| member_manifest(index, shape, payload))
        .collect();
    let opc_bytes = build_opc(&payloads, &manifests)?;
    let cfb_bytes = build_cfb(&payloads, &manifests)?;
    Ok(Corpus {
        shape,
        metadata_members: vec!["[Content_Types].xml".to_string(), "_rels/.rels".to_string()],
        members: manifests,
        selected_payload_member_count: MEMBER_COUNT,
        opc_sha256: sha256_hex(&opc_bytes),
        cfb_sha256: sha256_hex(&cfb_bytes),
        opc_bytes,
        cfb_bytes,
        payloads,
    })
}

#[cfg(feature = "source-metrics")]
const HISTOGRAM_BINS: usize = 8;

#[cfg(feature = "source-metrics")]
#[derive(Debug)]
struct ObserverCounters {
    logical_calls: std::sync::atomic::AtomicU64,
    requested_bytes: std::sync::atomic::AtomicU64,
    returned_bytes: std::sync::atomic::AtomicU64,
    short_reads: std::sync::atomic::AtomicU64,
    active_reads: std::sync::atomic::AtomicU64,
    max_simultaneous_reads: std::sync::atomic::AtomicU64,
    request_size_histogram: [std::sync::atomic::AtomicU64; HISTOGRAM_BINS],
}

#[cfg(feature = "source-metrics")]
impl ObserverCounters {
    fn new() -> Self {
        Self {
            logical_calls: std::sync::atomic::AtomicU64::new(0),
            requested_bytes: std::sync::atomic::AtomicU64::new(0),
            returned_bytes: std::sync::atomic::AtomicU64::new(0),
            short_reads: std::sync::atomic::AtomicU64::new(0),
            active_reads: std::sync::atomic::AtomicU64::new(0),
            max_simultaneous_reads: std::sync::atomic::AtomicU64::new(0),
            request_size_histogram: std::array::from_fn(|_| std::sync::atomic::AtomicU64::new(0)),
        }
    }

    fn reset(&self) {
        use std::sync::atomic::Ordering;
        self.logical_calls.store(0, Ordering::Relaxed);
        self.requested_bytes.store(0, Ordering::Relaxed);
        self.returned_bytes.store(0, Ordering::Relaxed);
        self.short_reads.store(0, Ordering::Relaxed);
        self.active_reads.store(0, Ordering::Relaxed);
        self.max_simultaneous_reads.store(0, Ordering::Relaxed);
        for bucket in &self.request_size_histogram {
            bucket.store(0, Ordering::Relaxed);
        }
    }

    fn bucket(requested: usize) -> usize {
        match requested {
            0..=64 => 0,
            65..=256 => 1,
            257..=1024 => 2,
            1025..=4096 => 3,
            4097..=16384 => 4,
            16385..=65536 => 5,
            65537..=262144 => 6,
            _ => 7,
        }
    }

    fn snapshot(&self) -> SourceMetricsReport {
        use std::sync::atomic::Ordering;
        SourceMetricsReport {
            availability: "source-metrics-feature".to_string(),
            logical_calls: Some(self.logical_calls.load(Ordering::Relaxed)),
            requested_bytes: Some(self.requested_bytes.load(Ordering::Relaxed)),
            returned_bytes: Some(self.returned_bytes.load(Ordering::Relaxed)),
            short_reads: Some(self.short_reads.load(Ordering::Relaxed)),
            active_reads_after_operation: Some(self.active_reads.load(Ordering::Relaxed)),
            max_simultaneous_reads: Some(self.max_simultaneous_reads.load(Ordering::Relaxed)),
            request_size_histogram: Some(
                self.request_size_histogram
                    .iter()
                    .map(|bucket| bucket.load(Ordering::Relaxed))
                    .collect(),
            ),
        }
    }
}

#[derive(Debug)]
struct InMemoryReadAt {
    bytes: Arc<Vec<u8>>,
    version: SourceVersion,
    #[cfg(feature = "source-metrics")]
    observer: ObserverCounters,
}

impl InMemoryReadAt {
    fn new(bytes: Vec<u8>, version: SourceVersion) -> Self {
        Self {
            bytes: Arc::new(bytes),
            version,
            #[cfg(feature = "source-metrics")]
            observer: ObserverCounters::new(),
        }
    }

    #[cfg(feature = "source-metrics")]
    fn reset_observer(&self) {
        self.observer.reset();
    }

    #[cfg(not(feature = "source-metrics"))]
    const fn reset_observer(&self) {}

    #[cfg(feature = "source-metrics")]
    fn observer_report(&self) -> SourceMetricsReport {
        self.observer.snapshot()
    }

    #[cfg(not(feature = "source-metrics"))]
    fn observer_report(&self) -> SourceMetricsReport {
        SourceMetricsReport {
            availability: "unavailable-normal-build".to_string(),
            logical_calls: None,
            requested_bytes: None,
            returned_bytes: None,
            short_reads: None,
            active_reads_after_operation: None,
            max_simultaneous_reads: None,
            request_size_histogram: None,
        }
    }
}

impl ReadAt for InMemoryReadAt {
    fn len(&self) -> std::io::Result<u64> {
        u64::try_from(self.bytes.len()).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidData, "source length exceeds u64")
        })
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(self.version)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        #[cfg(feature = "source-metrics")]
        let active = {
            use std::sync::atomic::Ordering;
            self.observer.logical_calls.fetch_add(1, Ordering::Relaxed);
            self.observer.requested_bytes.fetch_add(
                u64::try_from(output.len()).unwrap_or(u64::MAX),
                Ordering::Relaxed,
            );
            let active = self
                .observer
                .active_reads
                .fetch_add(1, Ordering::AcqRel)
                .saturating_add(1);
            self.observer
                .max_simultaneous_reads
                .fetch_max(active, Ordering::Relaxed);
            self.observer.request_size_histogram[ObserverCounters::bucket(output.len())]
                .fetch_add(1, Ordering::Relaxed);
            active
        };

        let result = (|| {
            let start = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "source offset exceeds usize",
                )
            })?;
            if start >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        })();

        #[cfg(feature = "source-metrics")]
        {
            use std::sync::atomic::Ordering;
            let returned = result.as_ref().copied().unwrap_or(0);
            self.observer.returned_bytes.fetch_add(
                u64::try_from(returned).unwrap_or(u64::MAX),
                Ordering::Relaxed,
            );
            if returned < output.len() {
                self.observer.short_reads.fetch_add(1, Ordering::Relaxed);
            }
            let _ = active;
            self.observer.active_reads.fetch_sub(1, Ordering::AcqRel);
        }

        result
    }
}

fn not_applicable_source_metrics() -> SourceMetricsReport {
    SourceMetricsReport {
        availability: "not-applicable-opc-from-bytes".to_string(),
        logical_calls: None,
        requested_bytes: None,
        returned_bytes: None,
        short_reads: None,
        active_reads_after_operation: None,
        max_simultaneous_reads: None,
        request_size_histogram: None,
    }
}

fn cpu_now_ns() -> Option<u64> {
    #[cfg(target_os = "linux")]
    {
        let value = rustix::time::clock_gettime(rustix::time::ClockId::ProcessCPUTime);
        let seconds = u64::try_from(value.tv_sec).ok()?;
        let nanos = u64::try_from(value.tv_nsec).ok()?;
        seconds.checked_mul(1_000_000_000)?.checked_add(nanos)
    }
    #[cfg(not(target_os = "linux"))]
    {
        None
    }
}

fn measure<T, F>(operation: F) -> AnyResult<(T, u128, Option<u64>)>
where
    F: FnOnce() -> AnyResult<T>,
{
    let cpu_start = cpu_now_ns();
    let wall_start = Instant::now();
    let value = operation()?;
    let wall_ns = wall_start.elapsed().as_nanos();
    let cpu_ns = cpu_start
        .zip(cpu_now_ns())
        .and_then(|(start, end)| end.checked_sub(start));
    Ok((value, wall_ns, cpu_ns))
}

fn make_context(workers: usize, task_floor: u64) -> AnyResult<(Budget, ExecutionContext)> {
    let root_limits = Limits::new(
        128 * 1024 * 1024,
        64 * 1024 * 1024,
        64 * 1024 * 1024,
        1_000_000,
        512,
        512 * 1024 * 1024,
    )
    .with_execution_io(workers as u64, workers as u64, CPU_TASK_LIMIT);
    let root = Budget::root("perf-execution-0786", root_limits);
    let (_source, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(workers).ok_or_else(|| boxed_error("workers must be nonzero"))?,
        NonZeroUsize::new(MEMBER_COUNT).ok_or_else(|| boxed_error("member count is zero"))?,
        NonZeroU64::new(MAX_IN_FLIGHT_BYTES)
            .ok_or_else(|| boxed_error("in-flight byte limit must be nonzero"))?,
        MIN_PARALLEL_BYTES,
    )?
    .with_min_task_bytes(task_floor)?;
    let context = ExecutionContext::new(root.clone(), token, limits);
    Ok((root, context))
}

fn resource_snapshot(root: &Budget) -> ResourceSnapshot {
    ResourceSnapshot {
        workers: root.used(Resource::Workers),
        io_concurrency: root.used(Resource::IoConcurrency),
        cpu_tasks: root.used(Resource::CpuTasks),
    }
}

fn make_resource_report(
    root: &Budget,
    workers: usize,
    before_operation: ResourceSnapshot,
    after_operation: ResourceSnapshot,
    after_drop: ResourceSnapshot,
) -> AnyResult<ResourceReport> {
    let limits = ResourceLimits {
        memory: 128 * 1024 * 1024,
        input_bytes: 64 * 1024 * 1024,
        output_bytes: 64 * 1024 * 1024,
        objects: 1_000_000,
        depth: 512,
        work: 512 * 1024 * 1024,
        workers: workers as u64,
        io_concurrency: workers as u64,
        cpu_tasks: CPU_TASK_LIMIT,
        max_in_flight_tasks: MEMBER_COUNT as u64,
        max_in_flight_bytes: MAX_IN_FLIGHT_BYTES,
    };
    let cpu_tasks_within_limit = [
        before_operation.cpu_tasks,
        after_operation.cpu_tasks,
        after_drop.cpu_tasks,
    ]
    .into_iter()
    .all(|used| used <= limits.cpu_tasks);
    let worker_and_io_released = after_drop.workers == 0 && after_drop.io_concurrency == 0;
    if !worker_and_io_released {
        return Err(boxed_error(format!(
            "execution permits leaked after drop: workers={}, io_concurrency={}",
            after_drop.workers, after_drop.io_concurrency
        )));
    }
    if !cpu_tasks_within_limit {
        return Err(boxed_error(format!(
            "CPU task budget exceeded: before={}, after={}, final={}, limit={}",
            before_operation.cpu_tasks,
            after_operation.cpu_tasks,
            after_drop.cpu_tasks,
            limits.cpu_tasks
        )));
    }
    // Keep the root in the function's contract so a future route cannot
    // accidentally report a snapshot from an unrelated budget instance.
    let _ = root.limit(Resource::Workers);
    Ok(ResourceReport {
        limits,
        before_operation,
        after_operation,
        after_drop,
        worker_and_io_released,
        cpu_tasks_within_limit,
    })
}

fn verify_sequence<'a, I>(items: I, corpus: &Corpus) -> AnyResult<ByteVerification>
where
    I: IntoIterator<Item = &'a [u8]>,
{
    let mut sequence = Sha256::new();
    let mut logical_bytes = 0_u64;
    let mut count = 0usize;
    for (index, bytes) in items.into_iter().enumerate() {
        let expected = corpus
            .payloads
            .get(index)
            .ok_or_else(|| boxed_error(format!("returned too many members: {index}")))?;
        let manifest = corpus
            .members
            .get(index)
            .ok_or_else(|| boxed_error(format!("missing member manifest: {index}")))?;
        if bytes != expected.as_slice() {
            return Err(boxed_error(format!(
                "member {} bytes differ from deterministic corpus",
                manifest.index
            )));
        }
        if sha256_hex(bytes) != manifest.sha256 {
            return Err(boxed_error(format!(
                "member {} SHA-256 differs",
                manifest.index
            )));
        }
        sequence.update((index as u64).to_le_bytes());
        sequence.update(bytes);
        logical_bytes = logical_bytes
            .checked_add(
                u64::try_from(bytes.len()).map_err(|_| boxed_error("byte count overflow"))?,
            )
            .ok_or_else(|| boxed_error("logical byte count overflow"))?;
        count += 1;
    }
    if count != MEMBER_COUNT {
        return Err(boxed_error(format!(
            "returned {} members, expected {}",
            count, MEMBER_COUNT
        )));
    }
    let sequence_digest = sequence.finalize();
    Ok(ByteVerification {
        ordered: true,
        all_member_sha256_match: true,
        members: count,
        logical_bytes,
        sequence_sha256: hex_digest(&sequence_digest),
    })
}

fn hex_digest(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn verify_buffers(buffers: &[Vec<u8>], corpus: &Corpus) -> AnyResult<ByteVerification> {
    verify_sequence(buffers.iter().map(Vec::as_slice), corpus)
}

fn part_uris(corpus: &Corpus) -> AnyResult<Vec<PackURI>> {
    corpus
        .members
        .iter()
        .map(|member| {
            PackURI::new(member.opc_uri.clone()).map_err(|error| boxed_error(error.to_string()))
        })
        .collect()
}

fn verify_opc(package: &litchi_opc::OpcPackage, corpus: &Corpus) -> AnyResult<ByteVerification> {
    if package.part_count() != MEMBER_COUNT {
        return Err(boxed_error(format!(
            "OPC package returned {} payload parts, expected {}",
            package.part_count(),
            MEMBER_COUNT
        )));
    }
    let mut sequence = Sha256::new();
    let mut logical_bytes = 0_u64;
    for (index, member) in corpus.members.iter().enumerate() {
        let uri =
            PackURI::new(member.opc_uri.clone()).map_err(|error| boxed_error(error.to_string()))?;
        let part = package.get_part(&uri)?;
        let bytes = part.blob();
        let expected = corpus
            .payloads
            .get(index)
            .ok_or_else(|| boxed_error(format!("missing payload {index}")))?;
        if bytes != expected.as_slice() {
            return Err(boxed_error(format!("OPC member {index} bytes differ")));
        }
        if sha256_hex(bytes) != member.sha256 {
            return Err(boxed_error(format!("OPC member {index} SHA-256 differs")));
        }
        sequence.update((index as u64).to_le_bytes());
        sequence.update(bytes);
        logical_bytes = logical_bytes
            .checked_add(
                u64::try_from(bytes.len()).map_err(|_| boxed_error("byte count overflow"))?,
            )
            .ok_or_else(|| boxed_error("logical byte count overflow"))?;
    }
    let sequence_digest = sequence.finalize();
    Ok(ByteVerification {
        ordered: true,
        all_member_sha256_match: true,
        members: MEMBER_COUNT,
        logical_bytes,
        sequence_sha256: hex_digest(&sequence_digest),
    })
}

fn verify_parts(batch: &litchi_opc::PartBatch, corpus: &Corpus) -> AnyResult<ByteVerification> {
    verify_sequence(batch.as_slice().iter().map(|part| part.as_bytes()), corpus)
}

fn run_opc_sample(config: &Config, corpus: &Corpus, sample: usize) -> AnyResult<SampleReport> {
    let (root, context) = make_context(config.workers, config.task_floor)?;
    let session = OpenSession::new(context.clone())?;
    if config.state == State::Primed {
        let preload = session.from_bytes(&corpus.opc_bytes, ReadLimits::default())?;
        let _ = verify_opc(&preload, corpus)?;
        drop(preload);
    }
    let before_operation = resource_snapshot(&root);
    let (package, wall_ns, cpu_ns) = measure(|| {
        session
            .from_bytes(&corpus.opc_bytes, ReadLimits::default())
            .map_err(|error| boxed_error(error.to_string()))
    })?;
    let after_operation = resource_snapshot(&root);
    let verification = verify_opc(&package, corpus)?;
    drop(package);
    drop(session);
    drop(context);
    let after_drop = resource_snapshot(&root);
    let resources = make_resource_report(
        &root,
        config.workers,
        before_operation,
        after_operation,
        after_drop,
    )?;
    Ok(SampleReport {
        sample,
        wall_ns,
        cpu_ns,
        primed_cache_hit_control: false,
        verification,
        resources,
        source_metrics: not_applicable_source_metrics(),
    })
}

fn run_cfb_sample(config: &Config, corpus: &Corpus, sample: usize) -> AnyResult<SampleReport> {
    let (root, context) = make_context(config.workers, config.task_floor)?;
    let source = Arc::new(InMemoryReadAt::new(
        corpus.cfb_bytes.clone(),
        SourceVersion::new(0x0786_0001, 1),
    ));
    let source_for_reader: Arc<dyn ReadAt> = source.clone();
    let file = SharedOleFile::open_with_limits(source_for_reader, SharedOleFileLimits::default())?;
    let session: SharedOleBulkRead<'_> = file.bulk_read(context.clone());
    let names: Vec<String> = corpus
        .members
        .iter()
        .map(|member| member.cfb_name.clone())
        .collect();
    let path_storage: Vec<Vec<&str>> = names.iter().map(|name| vec![name.as_str()]).collect();
    let paths: Vec<&[&str]> = path_storage.iter().map(Vec::as_slice).collect();
    if config.state == State::Primed {
        source.reset_observer();
        let preload = session.read_streams(&paths)?;
        let _ = verify_buffers(&preload, corpus)?;
        drop(preload);
        source.reset_observer();
    } else {
        source.reset_observer();
    }
    let before_operation = resource_snapshot(&root);
    let (buffers, wall_ns, cpu_ns) = measure(|| {
        session
            .read_streams(&paths)
            .map_err(|error| boxed_error(error.to_string()))
    })?;
    let after_operation = resource_snapshot(&root);
    let verification = verify_buffers(&buffers, corpus)?;
    let source_metrics = source.observer_report();
    drop(buffers);
    drop(session);
    drop(file);
    drop(context);
    let after_drop = resource_snapshot(&root);
    let resources = make_resource_report(
        &root,
        config.workers,
        before_operation,
        after_operation,
        after_drop,
    )?;
    Ok(SampleReport {
        sample,
        wall_ns,
        cpu_ns,
        primed_cache_hit_control: false,
        verification,
        resources,
        source_metrics,
    })
}

fn run_parts_sample(config: &Config, corpus: &Corpus, sample: usize) -> AnyResult<SampleReport> {
    let (root, context) = make_context(config.workers, config.task_floor)?;
    let source = Arc::new(InMemoryReadAt::new(
        corpus.opc_bytes.clone(),
        SourceVersion::new(0x0786_0002, 1),
    ));
    let source_for_reader: Arc<dyn ReadAt> = source.clone();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source_for_reader,
        ReadLimits::default(),
        context.clone(),
    )?;
    let uris = part_uris(corpus)?;
    if config.state == State::Primed {
        source.reset_observer();
        let preload = package.read_parts_ordered(&uris)?;
        let _ = verify_parts(&preload, corpus)?;
        drop(preload);
        source.reset_observer();
    } else {
        source.reset_observer();
    }
    let before_operation = resource_snapshot(&root);
    let (batch, wall_ns, cpu_ns) = measure(|| {
        package
            .read_parts_ordered(&uris)
            .map_err(|error| boxed_error(error.to_string()))
    })?;
    let after_operation = resource_snapshot(&root);
    let verification = verify_parts(&batch, corpus)?;
    let source_metrics = source.observer_report();
    drop(batch);
    drop(package);
    drop(context);
    let after_drop = resource_snapshot(&root);
    let resources = make_resource_report(
        &root,
        config.workers,
        before_operation,
        after_operation,
        after_drop,
    )?;
    Ok(SampleReport {
        sample,
        wall_ns,
        cpu_ns,
        primed_cache_hit_control: config.state == State::Primed,
        verification,
        resources,
        source_metrics,
    })
}

fn run_sample(config: &Config, corpus: &Corpus, sample: usize) -> AnyResult<SampleReport> {
    match config.route {
        Route::Opc => run_opc_sample(config, corpus, sample),
        Route::Cfb => run_cfb_sample(config, corpus, sample),
        Route::Parts => run_parts_sample(config, corpus, sample),
    }
}

fn report_scope(route: Route, state: State) -> Scope {
    Scope {
        ingress: match route {
            Route::Opc => "OpenSession::from_bytes borrowed in-memory ZIP".to_string(),
            Route::Cfb => "SharedOleFile owned immutable in-memory ReadAt".to_string(),
            Route::Parts => "SourceBackedPackage owned immutable in-memory ReadAt".to_string(),
        },
        source: "in-memory immutable source; no disk or network".to_string(),
        source_metrics_claim: if route == Route::Opc {
            "external ReadAt not applicable to from_bytes".to_string()
        } else {
            "logical ReadAt observer only when source-metrics is enabled".to_string()
        },
        timing: match state {
            State::Fresh => {
                "first public operation after session/package metadata setup".to_string()
            },
            State::Primed => {
                "next public operation after one verified preload outside the clock".to_string()
            },
        },
        cpu_time:
            "ProcessCPUTime interval is process-wide and may be slightly wider than wall timing"
                .to_string(),
        claims_excluded: vec![
            "physical cold-cache behavior".to_string(),
            "filesystem or network latency".to_string(),
            "remote/range-source request cost".to_string(),
            "production source instrumentation overhead".to_string(),
        ],
    }
}

fn metrics_availability(route: Route) -> MetricsAvailability {
    MetricsAvailability {
        source_metrics_feature: cfg!(feature = "source-metrics"),
        source_metrics_scope:
            "CFB and source-backed Parts timed operation only; counters reset after preload"
                .to_string(),
        opc_external_read_at: if route == Route::Opc {
            "not-applicable".to_string()
        } else {
            "not-used".to_string()
        },
        cpu_clock: if cfg!(target_os = "linux") {
            "rustix ProcessCPUTime".to_string()
        } else {
            "unavailable outside Linux build".to_string()
        },
    }
}

fn write_report(path: &PathBuf, report: &Report) -> AnyResult<()> {
    if path.exists() {
        return Err(boxed_error(format!(
            "refusing to overwrite existing output {}",
            path.display()
        )));
    }
    let bytes = serde_json::to_vec_pretty(report)?;
    let mut file = match std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)
    {
        Ok(file) => file,
        Err(error) => return Err(Box::new(error)),
    };
    use std::io::Write;
    if let Err(error) = file.write_all(&bytes) {
        drop(file);
        let _ = std::fs::remove_file(path);
        return Err(Box::new(error));
    }
    Ok(())
}

fn run(config: Config) -> AnyResult<()> {
    if config.output.exists() {
        return Err(boxed_error(format!(
            "output already exists: {}",
            config.output.display()
        )));
    }
    let corpus = build_corpus(config.shape)?;
    for _ in 0..config.warmup {
        let _ = run_sample(&config, &corpus, usize::MAX)?;
    }
    let mut samples = Vec::new();
    samples
        .try_reserve_exact(config.samples)
        .map_err(|error| boxed_error(format!("sample report allocation failed: {error}")))?;
    for sample in 0..config.samples {
        samples.push(run_sample(&config, &corpus, sample)?);
    }
    let report = Report {
        schema: SCHEMA,
        scope: report_scope(config.route, config.state),
        config: ReportConfig {
            route: config.route,
            shape: config.shape,
            workers: config.workers,
            task_floor: config.task_floor,
            aggregate_parallel_bytes: MIN_PARALLEL_BYTES,
            state: config.state,
            samples: config.samples,
            warmup: config.warmup,
            cpu_task_limit: CPU_TASK_LIMIT,
        },
        corpus,
        metrics: metrics_availability(config.route),
        samples,
    };
    write_report(&config.output, &report)
}

fn main() {
    let config = match parse_args() {
        Ok(config) => config,
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(2);
        },
    };
    if let Err(error) = run(config) {
        eprintln!("benchmark failed: {error}");
        std::process::exit(1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::{AtomicUsize, Ordering};

    fn args(values: &[&str]) -> Vec<String> {
        values.iter().map(|value| (*value).to_string()).collect()
    }

    fn valid_args() -> Vec<String> {
        args(&[
            "--route",
            "parts",
            "--shape",
            "small",
            "--workers",
            "4",
            "--task-floor",
            "65536",
            "--state",
            "fresh",
            "--samples",
            "2",
            "--warmup",
            "1",
            "--output",
            "result.json",
        ])
    }

    #[test]
    fn generator_identity_and_argument_bounds() {
        let first = build_corpus(Shape::Mixed).expect("first corpus");
        let second = build_corpus(Shape::Mixed).expect("second corpus");
        assert_eq!(first.members, second.members);
        assert_eq!(first.opc_bytes, second.opc_bytes);
        assert_eq!(first.cfb_bytes, second.cfb_bytes);
        assert_eq!(
            first.metadata_members,
            vec!["[Content_Types].xml".to_string(), "_rels/.rels".to_string()]
        );
        assert_eq!(first.selected_payload_member_count, MEMBER_COUNT);
        assert_eq!(
            first.members.last().expect("last member").bytes,
            SMALL_BYTES
        );
        assert!(parse_args_from(valid_args()).is_ok());

        let mut too_many_workers = valid_args();
        too_many_workers[5] = (MAX_WORKERS + 1).to_string();
        assert!(parse_args_from(too_many_workers).is_err());

        let mut too_large_floor = valid_args();
        too_large_floor[7] = (MAX_IN_FLIGHT_BYTES + 1).to_string();
        assert!(parse_args_from(too_large_floor).is_err());
    }

    #[test]
    fn public_routes_verify_order_and_release_resources() {
        let corpus = build_corpus(Shape::Small).expect("corpus");
        for route in [Route::Opc, Route::Cfb, Route::Parts] {
            let config = Config {
                route,
                shape: Shape::Small,
                workers: 1,
                task_floor: 0,
                state: State::Fresh,
                samples: 1,
                warmup: 0,
                output: PathBuf::from("unused.json"),
            };
            let sample = run_sample(&config, &corpus, 0).expect("route sample");
            assert!(sample.verification.ordered);
            assert!(sample.verification.all_member_sha256_match);
            assert!(sample.resources.worker_and_io_released);
            assert!(sample.resources.cpu_tasks_within_limit);
        }
    }

    #[derive(Debug)]
    struct ReadCountingSource {
        bytes: Arc<Vec<u8>>,
        reads: AtomicUsize,
    }

    impl ReadCountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: Arc::new(bytes),
                reads: AtomicUsize::new(0),
            }
        }
    }

    impl ReadAt for ReadCountingSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(0x0786_ffff, 1))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            self.reads.fetch_add(1, Ordering::Relaxed);
            let start = usize::try_from(offset).map_err(|_| {
                std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset overflow")
            })?;
            if start >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - start);
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            Ok(count)
        }
    }

    #[test]
    fn zero_io_budget_refuses_before_first_payload_read() {
        let corpus = build_corpus(Shape::Small).expect("corpus");
        let source = Arc::new(ReadCountingSource::new(corpus.cfb_bytes.clone()));
        let source_for_reader: Arc<dyn ReadAt> = source.clone();
        let file =
            SharedOleFile::open_with_limits(source_for_reader, SharedOleFileLimits::default())
                .expect("open CFB source");
        source.reads.store(0, Ordering::Relaxed);

        let root = Budget::root(
            "zero-io-test",
            Limits::new(
                128 * 1024 * 1024,
                64 * 1024 * 1024,
                64 * 1024 * 1024,
                1_000_000,
                512,
                512 * 1024 * 1024,
            )
            .with_execution_io(1, 0, CPU_TASK_LIMIT),
        );
        let (_cancel, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("worker"),
            NonZeroUsize::new(MEMBER_COUNT).expect("tasks"),
            NonZeroU64::new(MAX_IN_FLIGHT_BYTES).expect("bytes"),
            MIN_PARALLEL_BYTES,
        )
        .expect("limits");
        let context = ExecutionContext::new(root, token, limits);
        let bulk = file.bulk_read(context);
        let names: Vec<String> = corpus
            .members
            .iter()
            .map(|member| member.cfb_name.clone())
            .collect();
        let path_storage: Vec<Vec<&str>> = names.iter().map(|name| vec![name.as_str()]).collect();
        let paths: Vec<&[&str]> = path_storage.iter().map(Vec::as_slice).collect();
        assert!(bulk.read_streams(&paths).is_err());
        assert_eq!(source.reads.load(Ordering::Relaxed), 0);
    }
}
