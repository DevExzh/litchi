//! Generic, harness-only source-backed PPTX cross-copy lifecycle evidence.
//!
//! This target is deliberately separate from the pinned 0454 external pilot.
//! It accepts a custody-bound source/destination pair, keeps the two caller
//! owners and their counters separate, and times only the four public API
//! phases: source open, destination open, planning, and publication.  Input
//! loading, adapter construction, sink reservation, semantic/raw ZIP oracles,
//! artifact writes, and teardown are outside those clocks.

use std::{
    collections::{BTreeMap, BTreeSet},
    error::Error,
    ffi::OsString,
    fs::{self, File, OpenOptions},
    io::{Read, Write},
    num::{NonZeroU64, NonZeroUsize},
    path::{Path, PathBuf},
    sync::Arc,
    time::Duration,
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    Resource,
};
use litchi_opc::{ReadLimits, SourceCacheDiagnostics, SourceCacheLimits};
use serde::{Deserialize, Serialize};
use soapberry_zip::office::{ArchiveLimits, ArchiveReader};
use soapberry_zip::{PreservationIndex, ZipArchive};

use crate::{pptx_cache_retention, pptx_range_source};

const SCHEMA: &str = "pptx_pair_lifecycle_v1";
const MAX_MANIFEST_BYTES: u64 = 1024 * 1024;
const MAX_INPUT_BYTES: u64 = 512 * 1024 * 1024;
const MAX_OUTPUT_BYTES: u64 = 1024 * 1024 * 1024;
const DEFAULT_OUTPUT_BYTES: u64 = 1024 * 1024;
const DEFAULT_MEMORY_BYTES: u64 = 512 * 1024 * 1024;
const DEFAULT_MAX_WRITE_BYTES: u64 = 64 * 1024;
const DEFAULT_RANGE_BYTES: usize = 256;
const MAX_RANGE_BYTES: usize = 1024 * 1024;
const MAX_DELAY_US: u64 = 100_000;
const MAX_WRITE_BYTES: u64 = 1024 * 1024;
const CACHE_BYTES: usize = 64 * 1024 * 1024;
const CACHE_ENTRIES: usize = 128;
const OBJECT_LIMIT: u64 = 1_000_000;
const DEPTH_LIMIT: u64 = 256;
const WORK_LIMIT: u64 = 8 * 1024 * 1024 * 1024;
const MAX_WRITE: u64 = 64 * 1024;
const METADATA_MEMBERS: [&str; 3] = [
    "ppt/presentation.xml",
    "ppt/_rels/presentation.xml.rels",
    "[Content_Types].xml",
];

const PROVIDER_SCOPE: &str = "custody-bound distinct source/destination PPTX pair through caller-owned bytes or bounded range adapters";
const TIMING_SCOPE: &str = "open_source, open_destination, plan, and publication contain only their immediately surrounding public API calls; input loading, adapter construction, sink reservation, diagnostics, oracles, artifact writes, and drops are outside";
const API_SUM_SCOPE: &str = "api_sum_ns is the checked sum of open_source_ns, open_destination_ns, plan_ns, and publication_ns; it is not an outer end-to-end timer";
const RANGE_SCOPE: &str = "PptxRangeSource logical caller ReadAt counters; request lengths and returned bytes do not describe physical storage or network I/O";
const ALLOCATION_SCOPE: &str = "optional operation-scoped global allocator counters; one non-nested region surrounds each API phase and normal binaries omit allocation fields";
const ORACLE_SCOPE: &str = "semantic slide readback plus generic ZIP preservation and copied-payload checks; an independent external oracle remains required for full dependency-closure and XML relationship proof";

#[derive(Clone, Debug, Deserialize, Serialize)]
struct ArchiveSpec {
    #[serde(alias = "archive_path", alias = "input_path")]
    path: PathBuf,
    #[serde(alias = "sha")]
    sha256: String,
    #[serde(alias = "size", alias = "length")]
    bytes: u64,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
struct OperationSpec {
    #[serde(alias = "source_slide_position")]
    source_slide: Option<usize>,
    #[serde(alias = "destination_slide_position")]
    destination_slide: Option<usize>,
    #[serde(alias = "insert_position")]
    insertion_position: Option<usize>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
struct LimitSpec {
    #[serde(alias = "input_bytes_limit")]
    input_bytes: Option<u64>,
    #[serde(alias = "output_bytes_limit")]
    output_bytes: Option<u64>,
    #[serde(alias = "memory_bytes_limit")]
    memory_bytes: Option<u64>,
    #[serde(alias = "max_write_request_bytes")]
    max_write_bytes: Option<u64>,
    #[serde(alias = "max_returned_bytes")]
    max_range_bytes: Option<usize>,
    #[serde(alias = "fixed_delay_microseconds")]
    delay_us: Option<u64>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
struct ExpectedOutputSpec {
    #[serde(alias = "output_sha256")]
    sha256: Option<String>,
    #[serde(alias = "output_bytes")]
    bytes: Option<u64>,
}

/// The protocol is intentionally permissive about aliases so a checked
/// manifest can be reused by the Python custody wrapper without rewriting
/// pair identities.  `normalize_manifest` still requires every operational
/// identity and selector before a sample can start.
#[derive(Clone, Debug, Deserialize, Serialize)]
struct PairManifest {
    schema: String,
    pair_id: String,
    #[serde(default)]
    provenance: serde_json::Value,
    #[serde(alias = "source_archive", alias = "source_input")]
    source: ArchiveSpec,
    #[serde(alias = "destination_archive", alias = "destination_input")]
    destination: ArchiveSpec,
    #[serde(default)]
    operation: Option<OperationSpec>,
    #[serde(default)]
    source_slide: Option<usize>,
    #[serde(default)]
    destination_slide: Option<usize>,
    #[serde(default)]
    insertion_position: Option<usize>,
    #[serde(default)]
    limits: LimitSpec,
    #[serde(default)]
    expected_output: Option<ExpectedOutputSpec>,
    #[serde(default)]
    source_revision: Option<String>,
}

#[derive(Clone, Debug)]
struct NormalizedManifest {
    raw: PairManifest,
    manifest_sha256: String,
    source_slide: usize,
    destination_slide: usize,
    insertion_position: usize,
    input_limit: u64,
    output_limit: u64,
    memory_limit: u64,
    max_write: u64,
    max_range: usize,
    delay_us: u64,
    source_revision: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ProviderKind {
    Bytes,
    Range,
}

impl ProviderKind {
    const fn name(self) -> &'static str {
        match self {
            Self::Bytes => "bytes",
            Self::Range => "range",
        }
    }
}

#[derive(Clone, Debug)]
struct Config {
    manifest_path: PathBuf,
    manifest: NormalizedManifest,
    provider: ProviderKind,
    samples: usize,
    warmup: usize,
    repeat: String,
    output: PathBuf,
    output_pptx: PathBuf,
    source_revision: String,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ConfiguredLimits {
    input_bytes: u64,
    output_bytes: u64,
    memory_bytes: u64,
    max_write_bytes: u64,
    cache_bytes: usize,
    cache_entries: usize,
    max_range_bytes: usize,
    fixed_delay_us: u64,
    workers: usize,
    max_in_flight_tasks: usize,
}

#[derive(Clone, Debug, Serialize)]
struct ArchiveIdentity {
    path: String,
    sha256: String,
    bytes: u64,
}

#[derive(Clone, Debug, Serialize)]
struct OperationIdentity {
    source_slide: usize,
    destination_slide: usize,
    insertion_position: usize,
    destination_slide_count_before: usize,
    destination_slide_count_after: usize,
    source_slide_name_sha256: String,
    source_slide_text_sha256: String,
    source_direct_image_count: usize,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PhaseDescription {
    label: &'static str,
    live_owners: &'static str,
}

const PHASES: [PhaseDescription; 9] = [
    PhaseDescription {
        label: "baseline",
        live_owners: "caller budgets only; source view, destination editor, adapters, and sink are not opened",
    },
    PhaseDescription {
        label: "opened",
        live_owners: "source view, destination editor, caller source adapters, and reserved sink",
    },
    PhaseDescription {
        label: "planned",
        live_owners: "source view, destination editor, source-retaining plan, caller source adapters, and sink",
    },
    PhaseDescription {
        label: "published",
        live_owners: "source view, source-retaining plan, publication result, caller source adapters, and sink; destination editor consumed",
    },
    PhaseDescription {
        label: "drop_result",
        live_owners: "source view, source-retaining plan, caller source adapters, and sink",
    },
    PhaseDescription {
        label: "drop_plan",
        live_owners: "source view, caller source adapters, and sink",
    },
    PhaseDescription {
        label: "drop_view",
        live_owners: "caller source adapters and sink",
    },
    PhaseDescription {
        label: "drop_caller_sources",
        live_owners: "sink only; caller source adapters released",
    },
    PhaseDescription {
        label: "drop_sink",
        live_owners: "no lifecycle-owned source, plan, result, or sink handle",
    },
];

#[derive(Clone, Copy, Debug, Serialize)]
struct ReadPoint {
    availability: &'static str,
    unavailable_reason: Option<&'static str>,
    snapshot: Option<pptx_range_source::PptxRangeSourceSnapshot>,
    delta: Option<pptx_range_source::PptxRangeSourceDelta>,
    counter_delta_checked: Option<bool>,
}

impl ReadPoint {
    const fn unavailable(reason: &'static str) -> Self {
        Self {
            availability: "unavailable",
            unavailable_reason: Some(reason),
            snapshot: None,
            delta: None,
            counter_delta_checked: None,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct PhaseRecord {
    label: &'static str,
    source_cache: pptx_cache_retention::CachePoint,
    destination_cache: pptx_cache_retention::CachePoint,
    source_reads: ReadPoint,
    destination_reads: ReadPoint,
    source_budget: pptx_cache_retention::BudgetPoint,
    destination_budget: pptx_cache_retention::BudgetPoint,
    rss: pptx_cache_retention::RssPoint,
}

#[derive(Clone, Debug, Serialize)]
struct TimingRecord {
    open_source_ns: u64,
    open_destination_ns: u64,
    open_ns: u64,
    plan_ns: u64,
    publication_ns: u64,
    api_sum_ns: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    open_source_allocation_metrics: Option<crate::allocation_metrics::Sample>,
    #[serde(skip_serializing_if = "Option::is_none")]
    open_destination_allocation_metrics: Option<crate::allocation_metrics::Sample>,
    #[serde(skip_serializing_if = "Option::is_none")]
    plan_allocation_metrics: Option<crate::allocation_metrics::Sample>,
    #[serde(skip_serializing_if = "Option::is_none")]
    publication_allocation_metrics: Option<crate::allocation_metrics::Sample>,
}

#[derive(Clone, Debug, Serialize)]
struct PublicationRecord {
    source_position: usize,
    destination_slide_position: usize,
    insertion_position: usize,
    destination_slide_count: usize,
    source_name: String,
}

#[derive(Clone, Debug, Serialize)]
struct SemanticOracleRecord {
    input_semantics_derived_before_timing: bool,
    output_semantics_verified: bool,
    destination_slide_order_verified: bool,
    source_direct_images_verified: bool,
    raw_untouched_destination_records_verified: bool,
    copied_source_payloads_verified: bool,
    source_unchanged_after_publication: bool,
    independent_external_oracle_required: bool,
}

#[derive(Clone, Debug, Serialize)]
struct RawOracleRecord {
    destination_member_count: usize,
    output_member_count: usize,
    added_member_count: usize,
    copied_source_payload_member_count: usize,
    regenerated_relationship_member_count: usize,
    metadata_members_exempted: Vec<&'static str>,
    untouched_destination_records_verified: bool,
    copied_source_payloads_verified: bool,
}

#[derive(Clone, Debug, Serialize)]
struct LifecycleRow {
    sample_index: usize,
    output_sha256: String,
    output_bytes: u64,
    publication: PublicationRecord,
    timings: TimingRecord,
    publication_sink: crate::SinkSummary,
    semantic_oracle: SemanticOracleRecord,
    raw_oracle: RawOracleRecord,
    phases: Vec<PhaseRecord>,
}

#[derive(Debug, Serialize)]
struct OutputArtifact {
    path: String,
    bytes: u64,
    sha256: String,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    pair_id: String,
    provenance: serde_json::Value,
    manifest_sha256: String,
    provider: &'static str,
    repeat: String,
    provider_scope: &'static str,
    timing_scope: &'static str,
    api_sum_scope: &'static str,
    allocation_scope: &'static str,
    range_scope: &'static str,
    oracle_scope: &'static str,
    source_revision: String,
    binary_sha256: String,
    binary_bytes: u64,
    current_exe: String,
    allocator: &'static str,
    instrumentation: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocator_counter_revision: Option<&'static str>,
    source: ArchiveIdentity,
    destination: ArchiveIdentity,
    operation: OperationIdentity,
    configured_limits: ConfiguredLimits,
    samples: usize,
    warmup: usize,
    checked_iteration_count: usize,
    input_source_unchanged: bool,
    output_artifact: OutputArtifact,
    phases: Vec<PhaseDescription>,
    samples_raw: Vec<LifecycleRow>,
}

#[derive(Clone, Debug)]
struct OwnerContext {
    budget: Budget,
    context: ExecutionContext,
}

#[derive(Clone, Debug)]
struct PairInputs {
    manifest: NormalizedManifest,
    source_bytes: Vec<u8>,
    destination_bytes: Vec<u8>,
    source_path: PathBuf,
    destination_path: PathBuf,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SlideExpectation {
    name: String,
    text: String,
    images: Vec<Option<Vec<u8>>>,
}

#[derive(Clone, Debug)]
struct SemanticExpectation {
    source: SlideExpectation,
    destination: Vec<SlideExpectation>,
}

type ProviderSource = pptx_range_source::PptxRangeSource;

#[derive(Clone, Debug, PartialEq, Eq)]
struct RawMember {
    payload: Vec<u8>,
    local: Vec<u8>,
    central: Vec<u8>,
}

fn parse_usize(
    value: &str,
    flag: &str,
    minimum: usize,
    maximum: usize,
) -> Result<usize, Box<dyn Error>> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("{flag} must be an unsigned decimal integer"))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(format!("{flag} must be between {minimum} and {maximum}").into());
    }
    Ok(parsed)
}

fn parse_u64(value: &str, flag: &str, maximum: u64) -> Result<u64, Box<dyn Error>> {
    let parsed = value
        .parse::<u64>()
        .map_err(|_| format!("{flag} must be an unsigned decimal integer"))?;
    if parsed > maximum {
        return Err(format!("{flag} must be between 0 and {maximum}").into());
    }
    Ok(parsed)
}

fn validate_digest(value: &str, label: &str) -> Result<(), Box<dyn Error>> {
    if value.len() != 64
        || !value.bytes().all(|byte| byte.is_ascii_hexdigit())
        || value != value.to_ascii_lowercase()
    {
        return Err(
            format!("{label} must contain exactly 64 lowercase hexadecimal characters").into(),
        );
    }
    Ok(())
}

fn validate_revision(value: &str, label: &str) -> Result<(), Box<dyn Error>> {
    if value.len() != 40
        || !value.bytes().all(|byte| byte.is_ascii_hexdigit())
        || value != value.to_ascii_lowercase()
    {
        return Err(
            format!("{label} must contain exactly 40 lowercase hexadecimal characters").into(),
        );
    }
    Ok(())
}

fn read_bounded(path: &Path, limit: u64, label: &str) -> Result<Vec<u8>, Box<dyn Error>> {
    if limit == 0 {
        return Err(format!("{label} read limit must be nonzero").into());
    }
    let metadata = fs::metadata(path)?;
    if !metadata.is_file() {
        return Err(format!("{label} is not a regular file").into());
    }
    if metadata.len() > limit {
        return Err(format!("{label} exceeds bounded read limit {limit}").into());
    }
    let take_limit = limit
        .checked_add(1)
        .ok_or_else(|| format!("{label} read limit overflows u64"))?;
    let file = File::open(path)?;
    let mut bytes = Vec::new();
    file.take(take_limit).read_to_end(&mut bytes)?;
    if u64::try_from(bytes.len())? > limit {
        return Err(format!("{label} changed during bounded read").into());
    }
    Ok(bytes)
}

fn provenance_kind(value: &serde_json::Value) -> Option<&str> {
    match value {
        serde_json::Value::String(value) => Some(value.as_str()),
        serde_json::Value::Object(map) => map.get("kind").and_then(serde_json::Value::as_str),
        _ => None,
    }
}

fn normalize_manifest(
    manifest: PairManifest,
    raw: &[u8],
) -> Result<NormalizedManifest, Box<dyn Error>> {
    if manifest.schema != SCHEMA
        && manifest.schema != "pptx_pair_manifest_v1"
        && manifest.schema != "pptx_pair_lifecycle_manifest_v1"
    {
        return Err(format!("unsupported pair manifest schema: {}", manifest.schema).into());
    }
    if manifest.pair_id.trim().is_empty() {
        return Err("pair manifest pair_id is empty".into());
    }
    let provenance_kind =
        provenance_kind(&manifest.provenance).ok_or("pair manifest provenance.kind is required")?;
    if !matches!(
        provenance_kind,
        "self" | "derived" | "same-source-derived" | "independent"
    ) {
        return Err(format!("unsupported pair manifest provenance.kind: {provenance_kind}").into());
    }
    validate_digest(&manifest.source.sha256, "source.sha256")?;
    validate_digest(&manifest.destination.sha256, "destination.sha256")?;
    if manifest.source.bytes == 0 || manifest.destination.bytes == 0 {
        return Err("source and destination byte identities must be nonzero".into());
    }
    let operation = manifest.operation.clone().unwrap_or_default();
    let source_slide = manifest
        .source_slide
        .or(operation.source_slide)
        .ok_or("pair manifest must provide source_slide or operation.source_slide")?;
    let destination_slide = manifest
        .destination_slide
        .or(operation.destination_slide)
        .ok_or("pair manifest must provide destination_slide or operation.destination_slide")?;
    let insertion_position = manifest
        .insertion_position
        .or(operation.insertion_position)
        .ok_or("pair manifest must provide insertion_position or operation.insertion_position")?;
    let maximum_input = manifest.source.bytes.max(manifest.destination.bytes);
    let input_limit = manifest
        .limits
        .input_bytes
        .unwrap_or_else(|| maximum_input.max(512 * 1024));
    let output_default = maximum_input
        .checked_mul(4)
        .and_then(|value| value.checked_add(64 * 1024))
        .unwrap_or(DEFAULT_OUTPUT_BYTES)
        .max(DEFAULT_OUTPUT_BYTES);
    let output_limit = manifest.limits.output_bytes.unwrap_or(output_default);
    let memory_limit = manifest.limits.memory_bytes.unwrap_or(DEFAULT_MEMORY_BYTES);
    let max_write = manifest
        .limits
        .max_write_bytes
        .unwrap_or(DEFAULT_MAX_WRITE_BYTES);
    let max_range = manifest
        .limits
        .max_range_bytes
        .unwrap_or(DEFAULT_RANGE_BYTES);
    let delay_us = manifest.limits.delay_us.unwrap_or(0);
    if input_limit < maximum_input || input_limit > MAX_INPUT_BYTES {
        return Err(format!(
            "input byte limit {input_limit} does not cover the pair or exceeds the harness ceiling"
        )
        .into());
    }
    if output_limit == 0 || output_limit > MAX_OUTPUT_BYTES {
        return Err(format!("output byte limit must be between 1 and {MAX_OUTPUT_BYTES}").into());
    }
    if memory_limit == 0 || memory_limit > MAX_OUTPUT_BYTES {
        return Err(format!("memory byte limit must be between 1 and {MAX_OUTPUT_BYTES}").into());
    }
    if max_write == 0 || max_write > MAX_WRITE_BYTES {
        return Err(format!("max write size must be between 1 and {MAX_WRITE_BYTES}").into());
    }
    if max_range == 0 || max_range > MAX_RANGE_BYTES {
        return Err(format!("max range size must be between 1 and {MAX_RANGE_BYTES}").into());
    }
    if delay_us > MAX_DELAY_US {
        return Err(format!("fixed range delay exceeds {MAX_DELAY_US} microseconds").into());
    }
    if manifest.source.path == manifest.destination.path {
        return Err("source and destination paths must be distinct inputs".into());
    }
    if manifest.source.sha256 == manifest.destination.sha256 && provenance_kind != "self" {
        return Err("equal source and destination digests require provenance.kind=self".into());
    }
    let source_revision = manifest
        .source_revision
        .clone()
        .ok_or("source_revision is required and must bind a source revision")?;
    validate_revision(&source_revision, "source_revision")?;
    if let Some(expected) = manifest
        .expected_output
        .as_ref()
        .and_then(|value| value.sha256.as_deref())
    {
        validate_digest(expected, "expected_output.sha256")?;
    }
    Ok(NormalizedManifest {
        raw: manifest,
        manifest_sha256: crate::sha256_hex(raw),
        source_slide,
        destination_slide,
        insertion_position,
        input_limit,
        output_limit,
        memory_limit,
        max_write,
        max_range,
        delay_us,
        source_revision,
    })
}

fn resolve_input_path(manifest_path: &Path, path: &Path) -> PathBuf {
    if path.is_absolute() {
        return path.to_owned();
    }
    let cwd_path = PathBuf::from(path);
    if cwd_path.is_file() {
        cwd_path
    } else {
        manifest_path
            .parent()
            .unwrap_or_else(|| Path::new("."))
            .join(path)
    }
}

fn load_inputs(
    manifest_path: &Path,
    manifest: NormalizedManifest,
) -> Result<PairInputs, Box<dyn Error>> {
    let source_path = resolve_input_path(manifest_path, &manifest.raw.source.path);
    let destination_path = resolve_input_path(manifest_path, &manifest.raw.destination.path);
    let source_canonical = fs::canonicalize(&source_path)?;
    let destination_canonical = fs::canonicalize(&destination_path)?;
    if source_canonical == destination_canonical {
        return Err("source and destination resolve to the same file".into());
    }
    let source_bytes = read_bounded(&source_path, manifest.input_limit, "source input")?;
    let destination_bytes =
        read_bounded(&destination_path, manifest.input_limit, "destination input")?;
    if u64::try_from(source_bytes.len())? != manifest.raw.source.bytes {
        return Err("source byte identity differs from the pinned manifest".into());
    }
    if u64::try_from(destination_bytes.len())? != manifest.raw.destination.bytes {
        return Err("destination byte identity differs from the pinned manifest".into());
    }
    if crate::sha256_hex(&source_bytes) != manifest.raw.source.sha256 {
        return Err("source SHA-256 differs from the pinned manifest".into());
    }
    if crate::sha256_hex(&destination_bytes) != manifest.raw.destination.sha256 {
        return Err("destination SHA-256 differs from the pinned manifest".into());
    }
    Ok(PairInputs {
        manifest,
        source_bytes,
        destination_bytes,
        source_path,
        destination_path,
    })
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut manifest_path = None;
    let mut provider = None;
    let mut samples = None;
    let mut warmup = None;
    let mut repeat = None;
    let mut output = None;
    let mut output_pptx = None;
    let mut source_revision = None;
    let mut max_range = None;
    let mut delay_us = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("pptx-pair-lifecycle argument is not valid UTF-8")?;
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("missing value for {flag}"))?;
        let value_text = value
            .to_str()
            .ok_or("pptx-pair-lifecycle argument value is not valid UTF-8")?;
        index += 2;
        match flag {
            "--manifest" => {
                if manifest_path.is_some() {
                    return Err("duplicate --manifest".into());
                }
                manifest_path = Some(PathBuf::from(value));
            },
            "--provider" => {
                if provider.is_some() {
                    return Err("duplicate --provider".into());
                }
                provider = Some(match value_text {
                    "bytes" => ProviderKind::Bytes,
                    "range" => ProviderKind::Range,
                    _ => return Err("--provider must be bytes or range".into()),
                });
            },
            "--samples" => {
                if samples.is_some() {
                    return Err("duplicate --samples".into());
                }
                samples = Some(parse_usize(
                    value_text,
                    "--samples",
                    1,
                    pptx_cache_retention::MAX_SAMPLES,
                )?);
            },
            "--warmup" => {
                if warmup.is_some() {
                    return Err("duplicate --warmup".into());
                }
                warmup = Some(parse_usize(
                    value_text,
                    "--warmup",
                    0,
                    pptx_cache_retention::MAX_WARMUP,
                )?);
            },
            "--repeat" => {
                if repeat.is_some() {
                    return Err("duplicate --repeat".into());
                }
                if value_text.is_empty() || value_text.len() > 32 || !value_text.is_ascii() {
                    return Err(
                        "--repeat must be a nonempty ASCII label of at most 32 bytes".into(),
                    );
                }
                repeat = Some(value_text.to_owned());
            },
            "--output" => {
                if output.is_some() {
                    return Err("duplicate --output".into());
                }
                output = Some(PathBuf::from(value));
            },
            "--output-pptx" => {
                if output_pptx.is_some() {
                    return Err("duplicate --output-pptx".into());
                }
                output_pptx = Some(PathBuf::from(value));
            },
            "--source-revision" => {
                if source_revision.is_some() {
                    return Err("duplicate --source-revision".into());
                }
                validate_revision(value_text, "--source-revision")?;
                source_revision = Some(value_text.to_owned());
            },
            "--max-range" => {
                if max_range.is_some() {
                    return Err("duplicate --max-range".into());
                }
                max_range = Some(parse_usize(value_text, "--max-range", 1, MAX_RANGE_BYTES)?);
            },
            "--delay-us" => {
                if delay_us.is_some() {
                    return Err("duplicate --delay-us".into());
                }
                delay_us = Some(parse_u64(value_text, "--delay-us", MAX_DELAY_US)?);
            },
            _ => return Err(format!("unknown pptx-pair-lifecycle argument: {flag}").into()),
        }
    }
    let manifest_path = manifest_path.ok_or("--manifest is required")?;
    let raw = read_bounded(&manifest_path, MAX_MANIFEST_BYTES, "pair manifest")?;
    let manifest: PairManifest = serde_json::from_slice(&raw)?;
    let mut manifest = normalize_manifest(manifest, &raw)?;
    if let Some(value) = max_range {
        manifest.max_range = value;
    }
    if let Some(value) = delay_us {
        manifest.delay_us = value;
    }
    let output = output.ok_or("--output is required")?;
    let output_pptx = output_pptx.unwrap_or_else(|| output.with_extension("pptx"));
    if output == output_pptx {
        return Err("--output and --output-pptx must be different paths".into());
    }
    let source_revision = match source_revision {
        Some(value) => {
            if value != manifest.source_revision {
                return Err("--source-revision does not match manifest source_revision".into());
            }
            value
        },
        None => manifest.source_revision.clone(),
    };
    Ok(Config {
        manifest_path,
        manifest,
        provider: provider.ok_or("--provider is required")?,
        samples: samples.ok_or("--samples is required")?,
        warmup: warmup.ok_or("--warmup is required")?,
        repeat: repeat.ok_or("--repeat is required")?,
        output,
        output_pptx,
        source_revision,
    })
}

fn owner_context(
    label: &'static str,
    input_limit: u64,
    output_limit: u64,
    memory_limit: u64,
) -> Result<OwnerContext, Box<dyn Error>> {
    let budget = Budget::root(
        label,
        Limits::new(
            memory_limit,
            input_limit,
            output_limit,
            OBJECT_LIMIT,
            DEPTH_LIMIT,
            WORK_LIMIT,
        ),
    );
    let workers = NonZeroUsize::new(1).ok_or("one worker is invalid")?;
    let in_flight_bytes = NonZeroU64::new(memory_limit).ok_or("memory limit is zero")?;
    let execution_limits = ExecutionLimits::new(workers, workers, in_flight_bytes, 0)?;
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    Ok(OwnerContext { budget, context })
}

fn read_limits(input_limit: u64, expansion_limit: u64) -> Result<ReadLimits, Box<dyn Error>> {
    let expansion_limit = expansion_limit.max(1);
    let metadata_limit = (64 * 1024 * 1024).min(expansion_limit);
    let expansion_usize = usize::try_from(expansion_limit)?;
    let xml_limit = (8 * 1024 * 1024).min(expansion_usize).max(1);
    let attribute_limit = (64 * 1024).min(xml_limit).max(1);
    Ok(ReadLimits::builder()
        .max_input_bytes(input_limit)?
        .max_archive_member_name_bytes(4 * 1024)?
        .max_archive_metadata_bytes(metadata_limit)?
        .max_archive_compressed_bytes(input_limit)?
        .max_archive_entry_bytes(expansion_limit)?
        .max_archive_total_bytes(expansion_limit)?
        .max_part_bytes(expansion_limit)?
        .max_total_part_bytes(expansion_limit)?
        .max_content_types_bytes(xml_limit)?
        .max_relationship_xml_bytes(xml_limit)?
        .max_total_relationship_xml_bytes(xml_limit)?
        .max_xml_attribute_bytes(attribute_limit)?
        .max_relationship_target_bytes(attribute_limit)?
        .build()?)
}

fn make_provider(
    bytes: &[u8],
    provider: ProviderKind,
    max_range: usize,
    delay_us: u64,
) -> Arc<ProviderSource> {
    let inner: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes.to_vec()));
    let config = pptx_range_source::PptxRangeSourceConfig::new(
        matches!(provider, ProviderKind::Range).then_some(max_range),
        matches!(provider, ProviderKind::Range).then_some(Duration::from_micros(delay_us)),
    );
    Arc::new(ProviderSource::new(inner, config))
}

fn read_point(
    source: Option<&ProviderSource>,
    before: &mut Option<pptx_range_source::PptxRangeSourceSnapshot>,
) -> Result<ReadPoint, Box<dyn Error>> {
    let Some(source) = source else {
        *before = None;
        return Ok(ReadPoint::unavailable("caller source has been dropped"));
    };
    let snapshot = source.snapshot()?;
    let delta = snapshot.checked_delta(before.unwrap_or_default())?;
    *before = Some(snapshot);
    Ok(ReadPoint {
        availability: "available",
        unavailable_reason: None,
        snapshot: Some(snapshot),
        delta: Some(delta),
        counter_delta_checked: Some(true),
    })
}

fn phase(
    label: &'static str,
    source_view: Option<&litchi_pptx::SourceBackedPresentation>,
    destination_editor: Option<&litchi_pptx::SourceBackedPresentationEditor>,
    source_provider: Option<&ProviderSource>,
    destination_provider: Option<&ProviderSource>,
    source_before: &mut Option<pptx_range_source::PptxRangeSourceSnapshot>,
    destination_before: &mut Option<pptx_range_source::PptxRangeSourceSnapshot>,
    source_cache_before: &mut Option<SourceCacheDiagnostics>,
    destination_cache_before: &mut Option<SourceCacheDiagnostics>,
    source_context: &OwnerContext,
    destination_context: &OwnerContext,
    cache_limits: SourceCacheLimits,
) -> Result<PhaseRecord, Box<dyn Error>> {
    let (source_cache, source_cache_after) = if let Some(view) = source_view {
        let diagnostics = view.try_cache_diagnostics()?;
        let (point, after) = pptx_cache_retention::cache_point(
            diagnostics,
            *source_cache_before,
            &source_context.budget,
            cache_limits,
            None,
        )?;
        (point, Some(after))
    } else {
        (
            pptx_cache_retention::CachePoint::unavailable("source view is not retained"),
            None,
        )
    };
    let (destination_cache, destination_cache_after) = if let Some(editor) = destination_editor {
        let diagnostics = editor.try_cache_diagnostics()?;
        let (point, after) = pptx_cache_retention::cache_point(
            diagnostics,
            *destination_cache_before,
            &destination_context.budget,
            cache_limits,
            None,
        )?;
        (point, Some(after))
    } else {
        (
            pptx_cache_retention::CachePoint::unavailable("destination editor is not retained"),
            None,
        )
    };
    *source_cache_before = source_cache_after;
    *destination_cache_before = destination_cache_after;
    Ok(PhaseRecord {
        label,
        source_cache,
        destination_cache,
        source_reads: read_point(source_provider, source_before)?,
        destination_reads: read_point(destination_provider, destination_before)?,
        source_budget: pptx_cache_retention::budget_point(&source_context.budget),
        destination_budget: pptx_cache_retention::budget_point(&destination_context.budget),
        rss: pptx_cache_retention::rss_point(),
    })
}

fn checked_sum(values: &[u64], label: &str) -> Result<u64, Box<dyn Error>> {
    values.iter().try_fold(0_u64, |total, value| {
        total
            .checked_add(*value)
            .ok_or_else(|| format!("{label} overflows u64").into())
    })
}

fn image_expectations(
    slide: &litchi_pptx::SourceSlide,
) -> Result<Vec<Option<Vec<u8>>>, Box<dyn Error>> {
    let descriptors = slide.images()?;
    descriptors
        .iter()
        .map(|descriptor| {
            if descriptor.is_external() {
                Ok(None)
            } else {
                Ok(Some(
                    slide.read_image(descriptor.position())?.bytes().to_vec(),
                ))
            }
        })
        .collect()
}

fn semantic_expectation(
    bytes: &[u8],
    expansion_limit: u64,
) -> Result<Vec<SlideExpectation>, Box<dyn Error>> {
    let presentation = litchi_pptx::SourceBackedPresentation::from_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes.to_vec())),
        read_limits(u64::try_from(bytes.len())?, expansion_limit)?,
    )?;
    presentation
        .slides()
        .map(|slide| {
            let (text, name) = slide.text_and_name()?;
            Ok(SlideExpectation {
                name,
                text,
                images: image_expectations(&slide)?,
            })
        })
        .collect()
}

fn compare_semantics(
    expected: &SemanticExpectation,
    output: &[u8],
    insertion_position: usize,
    expansion_limit: u64,
) -> Result<(), Box<dyn Error>> {
    let mut expected_slides = expected.destination.clone();
    expected_slides.insert(insertion_position, expected.source.clone());
    let actual_slides = semantic_expectation(output, expansion_limit)?;
    if actual_slides != expected_slides {
        return Err("output semantic slide ordering, text, name, or direct-image projection differs from input-derived expectation".into());
    }
    Ok(())
}

fn archive_limits(compressed_limit: u64, expansion_limit: u64) -> ArchiveLimits {
    let expansion_limit = expansion_limit.max(1);
    ArchiveLimits {
        max_files: 100_000,
        max_member_name_bytes: 4 * 1024,
        max_metadata_bytes: (64 * 1024 * 1024).min(expansion_limit),
        max_compressed_size: compressed_limit.max(1),
        max_entry_size: expansion_limit,
        max_total_size: expansion_limit,
    }
}

fn raw_members(
    bytes: &[u8],
    expansion_limit: u64,
) -> Result<BTreeMap<String, RawMember>, Box<dyn Error>> {
    let compressed_limit = u64::try_from(bytes.len())?;
    let limits = archive_limits(compressed_limit, expansion_limit);
    let archive = ZipArchive::from_slice(bytes)?.into_zip_archive();
    let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new_with_limits(&archive, &mut scratch, limits)?;
    let reader = ArchiveReader::new_with_limits(bytes, limits)?;
    let mut members = BTreeMap::new();
    for entry in index.entries() {
        let name = std::str::from_utf8(entry.raw_name_bytes())?.to_owned();
        let local = entry.local_span();
        let central = entry.central_record();
        let local_start = usize::try_from(local.start)?;
        let local_end = usize::try_from(local.end)?;
        let central_start = usize::try_from(central.start)?;
        let central_end = usize::try_from(central.end)?;
        let mut central_bytes = bytes
            .get(central_start..central_end)
            .ok_or("central record lies outside the archive")?
            .to_vec();
        central_bytes
            .get_mut(42..46)
            .ok_or("central record is shorter than its relative offset")?
            .fill(0);
        let member = RawMember {
            payload: reader.read(&name)?,
            local: bytes
                .get(local_start..local_end)
                .ok_or("local record lies outside the archive")?
                .to_vec(),
            central: central_bytes,
        };
        if members.insert(name, member).is_some() {
            return Err("archive contains duplicate member names".into());
        }
    }
    Ok(members)
}

fn raw_oracle(
    source: &[u8],
    destination: &[u8],
    output: &[u8],
    expansion_limit: u64,
) -> Result<RawOracleRecord, Box<dyn Error>> {
    let source_members = raw_members(source, expansion_limit)?;
    let destination_members = raw_members(destination, expansion_limit)?;
    let output_members = raw_members(output, expansion_limit)?;
    if output_members.len() < destination_members.len() {
        return Err("publication dropped destination ZIP members".into());
    }
    let metadata: BTreeSet<&str> = METADATA_MEMBERS.into_iter().collect();
    for (name, member) in &destination_members {
        let actual = output_members
            .get(name)
            .ok_or_else(|| format!("publication dropped destination member {name}"))?;
        if !metadata.contains(name.as_str()) && actual != member {
            return Err(format!("untouched destination ZIP record changed: {name}").into());
        }
    }
    let added = output_members
        .keys()
        .filter(|name| !destination_members.contains_key(*name))
        .cloned()
        .collect::<Vec<_>>();
    if added.is_empty() {
        return Err("publication did not add a copied ZIP closure".into());
    }
    let mut copied_source_payload_member_count = 0;
    let mut regenerated_relationship_member_count = 0;
    for name in &added {
        let actual = &output_members[name];
        if name.ends_with(".rels") || name.contains("/_rels/") {
            regenerated_relationship_member_count += 1;
            continue;
        }
        if !source_members
            .values()
            .any(|source_member| source_member.payload == actual.payload)
        {
            return Err(format!("added member {name} has no source payload identity").into());
        }
        copied_source_payload_member_count += 1;
    }
    if copied_source_payload_member_count == 0 {
        return Err("publication did not add any non-relationship source payload".into());
    }
    Ok(RawOracleRecord {
        destination_member_count: destination_members.len(),
        output_member_count: output_members.len(),
        added_member_count: added.len(),
        copied_source_payload_member_count,
        regenerated_relationship_member_count,
        metadata_members_exempted: METADATA_MEMBERS.to_vec(),
        untouched_destination_records_verified: true,
        copied_source_payloads_verified: true,
    })
}

fn sink_ceiling(output_limit: u64) -> Result<u64, Box<dyn Error>> {
    output_limit
        .checked_add(MAX_WRITE)
        .ok_or_else(|| "pair lifecycle sink ceiling overflows u64".into())
}

fn cache_bytes(memory_limit: u64) -> Result<usize, Box<dyn Error>> {
    let half = memory_limit / 2;
    Ok(CACHE_BYTES.min(usize::try_from(half.max(1))?))
}

fn run_iteration(
    inputs: &PairInputs,
    semantics: &SemanticExpectation,
    config: &Config,
    sample_index: usize,
) -> Result<(LifecycleRow, Vec<u8>), Box<dyn Error>> {
    let source_context = owner_context(
        "pptx-pair-lifecycle-source",
        inputs.manifest.input_limit,
        inputs.manifest.output_limit,
        inputs.manifest.memory_limit,
    )?;
    let destination_context = owner_context(
        "pptx-pair-lifecycle-destination",
        inputs.manifest.input_limit,
        inputs.manifest.output_limit,
        inputs.manifest.memory_limit,
    )?;
    let read_limits = read_limits(inputs.manifest.input_limit, inputs.manifest.memory_limit)?;
    let cache_limit_bytes = cache_bytes(inputs.manifest.memory_limit)?;
    let cache_limits = SourceCacheLimits::new(cache_limit_bytes, CACHE_ENTRIES)?;
    let mut phases = Vec::with_capacity(PHASES.len());
    let mut source_reads_before = None;
    let mut destination_reads_before = None;
    let mut source_cache_before = None;
    let mut destination_cache_before = None;
    phases.push(phase(
        "baseline",
        None,
        None,
        None,
        None,
        &mut source_reads_before,
        &mut destination_reads_before,
        &mut source_cache_before,
        &mut destination_cache_before,
        &source_context,
        &destination_context,
        cache_limits,
    )?);

    let source_provider = make_provider(
        &inputs.source_bytes,
        config.provider,
        inputs.manifest.max_range,
        inputs.manifest.delay_us,
    );
    let destination_provider = make_provider(
        &inputs.destination_bytes,
        config.provider,
        inputs.manifest.max_range,
        inputs.manifest.delay_us,
    );
    let source_read: Arc<dyn ReadAt> = source_provider.clone();
    let destination_read: Arc<dyn ReadAt> = destination_provider.clone();
    let mut sink = crate::CountingSink::bounded(
        sink_ceiling(inputs.manifest.output_limit)?,
        inputs.manifest.max_write.min(MAX_WRITE),
    );
    sink.reserve_budget()?;

    let source_allocation_region = crate::allocation_metrics::begin();
    let source_started = std::time::Instant::now();
    let source_view = litchi_pptx::SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
        source_read,
        read_limits,
        cache_limits,
        source_context.context.clone(),
    )?;
    let open_source_ns = crate::elapsed_ns(source_started.elapsed())?;
    let open_source_allocation_metrics = source_allocation_region.finish();

    let destination_allocation_region = crate::allocation_metrics::begin();
    let destination_started = std::time::Instant::now();
    let mut destination_editor = Some(
        litchi_pptx::SourceBackedPresentationEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
            destination_read,
            read_limits,
            cache_limits,
            destination_context.context.clone(),
        )?,
    );
    let open_destination_ns = crate::elapsed_ns(destination_started.elapsed())?;
    let open_destination_allocation_metrics = destination_allocation_region.finish();
    if source_view.slide_count() <= inputs.manifest.source_slide
        || destination_editor
            .as_ref()
            .is_none_or(|editor| editor.slide_count() <= inputs.manifest.destination_slide)
        || inputs.manifest.insertion_position > destination_editor.as_ref().unwrap().slide_count()
    {
        return Err("pair operation slide selector is outside the opened presentations".into());
    }
    phases.push(phase(
        "opened",
        Some(&source_view),
        destination_editor.as_ref(),
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        &mut source_reads_before,
        &mut destination_reads_before,
        &mut source_cache_before,
        &mut destination_cache_before,
        &source_context,
        &destination_context,
        cache_limits,
    )?);

    let plan_allocation_region = crate::allocation_metrics::begin();
    let plan_started = std::time::Instant::now();
    let plan = destination_editor
        .as_ref()
        .ok_or("destination editor was consumed before planning")?
        .plan_cross_slide_copy(
            &source_view,
            inputs.manifest.source_slide,
            inputs.manifest.destination_slide,
            inputs.manifest.insertion_position,
        )?;
    let plan_ns = crate::elapsed_ns(plan_started.elapsed())?;
    let plan_allocation_metrics = plan_allocation_region.finish();
    if plan.source_position() != inputs.manifest.source_slide
        || plan.destination_slide_position() != inputs.manifest.destination_slide
        || plan.insertion_position() != inputs.manifest.insertion_position
        || plan.destination_slide_count() != destination_editor.as_ref().unwrap().slide_count() + 1
    {
        return Err("pair operation plan metadata is not deterministic".into());
    }
    phases.push(phase(
        "planned",
        Some(&source_view),
        destination_editor.as_ref(),
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        &mut source_reads_before,
        &mut destination_reads_before,
        &mut source_cache_before,
        &mut destination_cache_before,
        &source_context,
        &destination_context,
        cache_limits,
    )?);

    let publication_allocation_region = crate::allocation_metrics::begin();
    let publication_started = std::time::Instant::now();
    let published = destination_editor
        .take()
        .ok_or("destination editor was consumed before publication")?
        .publish_cross_slide_copy_to_stream(&mut sink, &plan)?;
    let publication_ns = crate::elapsed_ns(publication_started.elapsed())?;
    let publication_allocation_metrics = publication_allocation_region.finish();
    let publication_sink = sink.summary();
    let output = sink.bytes.clone();
    if published.destination_slide_count() != semantics.destination.len() + 1
        || published.insertion_position() != inputs.manifest.insertion_position
        || published.name() != semantics.source.name
        || publication_sink.accepted_bytes != u64::try_from(output.len())?
        || publication_sink.largest_write > inputs.manifest.max_write.min(MAX_WRITE)
    {
        return Err(
            "publication result or sink accounting differs from the input-derived contract".into(),
        );
    }
    source_view.check_source()?;
    compare_semantics(
        semantics,
        &output,
        inputs.manifest.insertion_position,
        inputs.manifest.memory_limit,
    )?;
    let raw_oracle = raw_oracle(
        &inputs.source_bytes,
        &inputs.destination_bytes,
        &output,
        inputs.manifest.memory_limit,
    )?;
    if let Some(expected) = inputs.manifest.raw.expected_output.as_ref() {
        if let Some(bytes) = expected.bytes
            && bytes != u64::try_from(output.len())?
        {
            return Err("output bytes differ from expected_output.bytes".into());
        }
        if let Some(sha256) = expected.sha256.as_deref()
            && sha256 != crate::sha256_hex(&output)
        {
            return Err("output SHA-256 differs from expected_output.sha256".into());
        }
    }
    phases.push(phase(
        "published",
        Some(&source_view),
        None,
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        &mut source_reads_before,
        &mut destination_reads_before,
        &mut source_cache_before,
        &mut destination_cache_before,
        &source_context,
        &destination_context,
        cache_limits,
    )?);

    drop(published);
    phases.push(phase(
        "drop_result",
        Some(&source_view),
        None,
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        &mut source_reads_before,
        &mut destination_reads_before,
        &mut source_cache_before,
        &mut destination_cache_before,
        &source_context,
        &destination_context,
        cache_limits,
    )?);
    drop(plan);
    phases.push(phase(
        "drop_plan",
        Some(&source_view),
        None,
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        &mut source_reads_before,
        &mut destination_reads_before,
        &mut source_cache_before,
        &mut destination_cache_before,
        &source_context,
        &destination_context,
        cache_limits,
    )?);
    drop(source_view);
    phases.push(phase(
        "drop_view",
        None,
        None,
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        &mut source_reads_before,
        &mut destination_reads_before,
        &mut source_cache_before,
        &mut destination_cache_before,
        &source_context,
        &destination_context,
        cache_limits,
    )?);
    drop(source_provider);
    drop(destination_provider);
    phases.push(phase(
        "drop_caller_sources",
        None,
        None,
        None,
        None,
        &mut source_reads_before,
        &mut destination_reads_before,
        &mut source_cache_before,
        &mut destination_cache_before,
        &source_context,
        &destination_context,
        cache_limits,
    )?);
    drop(sink);
    phases.push(phase(
        "drop_sink",
        None,
        None,
        None,
        None,
        &mut source_reads_before,
        &mut destination_reads_before,
        &mut source_cache_before,
        &mut destination_cache_before,
        &source_context,
        &destination_context,
        cache_limits,
    )?);
    for (label, budget) in [
        ("source", &source_context.budget),
        ("destination", &destination_context.budget),
    ] {
        for resource in [Resource::Memory, Resource::Objects, Resource::Depth] {
            if budget.used(resource) != 0 {
                return Err(format!(
                    "{label} budget {resource:?} remained nonzero after drop_sink"
                )
                .into());
            }
        }
    }
    let open_ns = checked_sum(&[open_source_ns, open_destination_ns], "open timing")?;
    let api_sum_ns = checked_sum(&[open_ns, plan_ns, publication_ns], "API timing")?;
    let output_sha256 = crate::sha256_hex(&output);
    let row = LifecycleRow {
        sample_index,
        output_sha256: output_sha256.clone(),
        output_bytes: u64::try_from(output.len())?,
        publication: PublicationRecord {
            source_position: inputs.manifest.source_slide,
            destination_slide_position: inputs.manifest.destination_slide,
            insertion_position: inputs.manifest.insertion_position,
            destination_slide_count: semantics.destination.len() + 1,
            source_name: semantics.source.name.clone(),
        },
        timings: TimingRecord {
            open_source_ns,
            open_destination_ns,
            open_ns,
            plan_ns,
            publication_ns,
            api_sum_ns,
            open_source_allocation_metrics,
            open_destination_allocation_metrics,
            plan_allocation_metrics,
            publication_allocation_metrics,
        },
        publication_sink,
        semantic_oracle: SemanticOracleRecord {
            input_semantics_derived_before_timing: true,
            output_semantics_verified: true,
            destination_slide_order_verified: true,
            source_direct_images_verified: true,
            raw_untouched_destination_records_verified: raw_oracle
                .untouched_destination_records_verified,
            copied_source_payloads_verified: raw_oracle.copied_source_payloads_verified,
            source_unchanged_after_publication: true,
            independent_external_oracle_required: true,
        },
        raw_oracle,
        phases,
    };
    Ok((row, output))
}

fn retain_output(path: &Path, bytes: &[u8]) -> Result<OutputArtifact, Box<dyn Error>> {
    let mut output = OpenOptions::new().write(true).create_new(true).open(path)?;
    output.write_all(bytes)?;
    output.flush()?;
    output.sync_all()?;
    Ok(OutputArtifact {
        path: path.to_string_lossy().into_owned(),
        bytes: u64::try_from(bytes.len())?,
        sha256: crate::sha256_hex(bytes),
    })
}

fn checked_file_identity(path: &Path, expected: &[u8], label: &str) -> Result<(), Box<dyn Error>> {
    let actual = read_bounded(path, u64::try_from(expected.len())?, label)?;
    if actual != expected {
        return Err(format!("{label} changed during capture").into());
    }
    Ok(())
}

fn run_capture(config: &Config) -> Result<Report, Box<dyn Error>> {
    let inputs = load_inputs(&config.manifest_path, config.manifest.clone())?;
    let source_semantics =
        semantic_expectation(&inputs.source_bytes, inputs.manifest.memory_limit)?;
    let destination_semantics =
        semantic_expectation(&inputs.destination_bytes, inputs.manifest.memory_limit)?;
    if source_semantics.len() <= inputs.manifest.source_slide
        || destination_semantics.len() <= inputs.manifest.destination_slide
        || inputs.manifest.insertion_position > destination_semantics.len()
    {
        return Err(
            "pair operation selector is outside the input-derived presentation catalogs".into(),
        );
    }
    let semantics = SemanticExpectation {
        source: source_semantics[inputs.manifest.source_slide].clone(),
        destination: destination_semantics,
    };
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or("warmup and samples overflow usize")?;
    let mut rows = Vec::with_capacity(config.samples);
    let mut first_output = None;
    let mut expected_identity = None;
    for iteration in 0..total {
        let (mut row, output) = run_iteration(&inputs, &semantics, config, iteration)?;
        if let Some((expected_sha, expected_bytes)) = expected_identity.as_ref()
            && (expected_sha != &row.output_sha256 || *expected_bytes != row.output_bytes)
        {
            return Err("pair publication output identity changed between iterations".into());
        }
        expected_identity = Some((row.output_sha256.clone(), row.output_bytes));
        if iteration >= config.warmup {
            row.sample_index = iteration - config.warmup;
            if first_output.is_none() {
                first_output = Some(retain_output(&config.output_pptx, &output)?);
            }
            rows.push(row);
        }
    }
    let output_artifact = first_output.ok_or("pair lifecycle retained no samples")?;
    checked_file_identity(&inputs.source_path, &inputs.source_bytes, "source input")?;
    checked_file_identity(
        &inputs.destination_path,
        &inputs.destination_bytes,
        "destination input",
    )?;
    let binary = crate::current_executable_identity()?;
    let configured_limits = ConfiguredLimits {
        input_bytes: inputs.manifest.input_limit,
        output_bytes: inputs.manifest.output_limit,
        memory_bytes: inputs.manifest.memory_limit,
        max_write_bytes: inputs.manifest.max_write.min(MAX_WRITE),
        cache_bytes: cache_bytes(inputs.manifest.memory_limit)?,
        cache_entries: CACHE_ENTRIES,
        max_range_bytes: inputs.manifest.max_range,
        fixed_delay_us: inputs.manifest.delay_us,
        workers: 1,
        max_in_flight_tasks: 1,
    };
    Ok(Report {
        schema: SCHEMA,
        pair_id: inputs.manifest.raw.pair_id.clone(),
        provenance: inputs.manifest.raw.provenance.clone(),
        manifest_sha256: inputs.manifest.manifest_sha256.clone(),
        provider: config.provider.name(),
        repeat: config.repeat.clone(),
        provider_scope: PROVIDER_SCOPE,
        timing_scope: TIMING_SCOPE,
        api_sum_scope: API_SUM_SCOPE,
        allocation_scope: ALLOCATION_SCOPE,
        range_scope: RANGE_SCOPE,
        oracle_scope: ORACLE_SCOPE,
        source_revision: config.source_revision.clone(),
        binary_sha256: binary.binary_sha256.clone(),
        binary_bytes: binary.binary_bytes,
        current_exe: binary.path.clone(),
        allocator: crate::allocation_metrics::allocator_identity(),
        instrumentation: crate::allocation_metrics::instrumentation_identity(),
        allocator_counter_revision: crate::allocation_metrics::counter_revision(),
        source: ArchiveIdentity {
            path: inputs.source_path.to_string_lossy().into_owned(),
            sha256: crate::sha256_hex(&inputs.source_bytes),
            bytes: u64::try_from(inputs.source_bytes.len())?,
        },
        destination: ArchiveIdentity {
            path: inputs.destination_path.to_string_lossy().into_owned(),
            sha256: crate::sha256_hex(&inputs.destination_bytes),
            bytes: u64::try_from(inputs.destination_bytes.len())?,
        },
        operation: OperationIdentity {
            source_slide: inputs.manifest.source_slide,
            destination_slide: inputs.manifest.destination_slide,
            insertion_position: inputs.manifest.insertion_position,
            destination_slide_count_before: semantics.destination.len(),
            destination_slide_count_after: semantics.destination.len() + 1,
            source_slide_name_sha256: crate::sha256_hex(semantics.source.name.as_bytes()),
            source_slide_text_sha256: crate::sha256_hex(semantics.source.text.as_bytes()),
            source_direct_image_count: semantics.source.images.len(),
        },
        configured_limits,
        samples: config.samples,
        warmup: config.warmup,
        checked_iteration_count: total,
        input_source_unchanged: true,
        output_artifact,
        phases: PHASES.to_vec(),
        samples_raw: rows,
    })
}

/// Runs the generic source/destination pair lifecycle after the selector has
/// been consumed by the normal or allocator benchmark binary.
pub fn run_from_args<I>(args: I) -> Result<(), Box<dyn Error>>
where
    I: IntoIterator<Item = OsString>,
{
    let args = args.into_iter().collect::<Vec<_>>();
    let config = parse_config(&args)?;
    let report = run_capture(&config)?;
    let bytes = serde_json::to_vec_pretty(&report)?;
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

    fn manifest(source_path: &str, destination_path: &str) -> PairManifest {
        PairManifest {
            schema: "pptx_pair_manifest_v1".to_owned(),
            pair_id: "test-pair".to_owned(),
            provenance: serde_json::json!({"kind":"derived"}),
            source: ArchiveSpec {
                path: PathBuf::from(source_path),
                sha256: "00".repeat(32),
                bytes: 1,
            },
            destination: ArchiveSpec {
                path: PathBuf::from(destination_path),
                sha256: "11".repeat(32),
                bytes: 1,
            },
            operation: Some(OperationSpec {
                source_slide: Some(0),
                destination_slide: Some(0),
                insertion_position: Some(1),
            }),
            source_slide: None,
            destination_slide: None,
            insertion_position: None,
            limits: LimitSpec::default(),
            expected_output: None,
            source_revision: Some("0123456789abcdef0123456789abcdef01234567".to_owned()),
        }
    }

    #[test]
    fn bad_manifest_digest_is_rejected_before_input_read() {
        let mut value = manifest("source.pptx", "destination.pptx");
        value.source.sha256 = "not-a-digest".to_owned();
        assert!(normalize_manifest(value, b"{}").is_err());
    }

    #[test]
    fn same_input_path_is_rejected_even_when_digests_differ() {
        let value = manifest("same.pptx", "same.pptx");
        assert!(normalize_manifest(value, b"{}").is_err());
    }

    #[test]
    fn operation_selectors_are_required_and_retained() {
        let value = manifest("source.pptx", "destination.pptx");
        let normalized = normalize_manifest(value, b"{}").unwrap();
        assert_eq!(normalized.source_slide, 0);
        assert_eq!(normalized.destination_slide, 0);
        assert_eq!(normalized.insertion_position, 1);
    }

    #[test]
    fn provenance_and_source_revision_are_closed_identities() {
        let mut unknown_provenance = manifest("source.pptx", "destination.pptx");
        unknown_provenance.provenance = serde_json::json!({"kind":"synthetic"});
        assert!(normalize_manifest(unknown_provenance, b"{}").is_err());

        let mut missing_revision = manifest("source.pptx", "destination.pptx");
        missing_revision.source_revision = None;
        assert!(normalize_manifest(missing_revision, b"{}").is_err());

        let mut malformed_revision = manifest("source.pptx", "destination.pptx");
        malformed_revision.source_revision = Some("unbound".to_owned());
        assert!(normalize_manifest(malformed_revision, b"{}").is_err());
    }
}
