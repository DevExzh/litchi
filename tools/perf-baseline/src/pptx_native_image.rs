//! Native-producer PPTX selected-image lifecycle evidence.
//!
//! This target deliberately stays separate from the synthetic cross-copy
//! fixtures.  It opens one fixed producer archive, queries one direct picture
//! on its fixed selected slide, and retains the returned image while each source-backed
//! owner is released.  The three duration fields cover only the public open,
//! metadata, and image-read calls.  Source construction, file hashing,
//! diagnostics, payload hashing, drops, and all correctness checks are
//! outside those clocks.

use std::{
    error::Error,
    ffi::OsString,
    fs::{self, OpenOptions},
    io::Write,
    path::PathBuf,
    sync::Arc,
    time::{Duration, Instant},
};

use litchi_core::{FileSource, OwnedSource, ReadAt, Resource};
use litchi_opc::{ReadLimits, SourceCacheLimits};
use serde::Serialize;
use serde_json::Value;

use crate::{pptx_cache_retention, pptx_range_source};

const SCHEMA: &str = "pptx_native_image_lifecycle_v1";
const MAX_SAMPLES: usize = 1_000;
const MAX_WARMUP: usize = 1_000;
const MAX_RANGE_BYTES: usize = 1_048_576;
const MAX_DELAY_US: u64 = 100_000;
const NATIVE_MEMORY_SCOPE: &str = "native producer PPTX selected-image lifecycle through the existing SourceBackedPresentation API; ownership and resource observations only";
const TIMING_SCOPE: &str = "open, metadata, and image-read durations contain only their immediately surrounding public API calls; source construction, hashing, diagnostics, checks, copies, and drops are outside the clocks";
const FILE_SCOPE: &str = "FileSource reads a task-selected local fixture after a setup hash read; this is a warm recently-hashed filesystem baseline and carries no cold-filesystem claim";
const RANGE_SCOPE: &str = "PptxRangeSource logical ReadAt adapter counters; caps and delay describe caller-visible calls, not physical network, filesystem, or device behavior";
const RSS_SCOPE: &str = "optional process-wide VmRSS/VmHWM snapshots from procfs; setup and unrelated process memory remain in scope and no comparative RSS claim is authorized";
const SOURCE_READ_SCOPE: &str = "checked PptxRangeSource snapshots at each lifecycle boundary while the caller adapter is retained; unavailable after the caller source is dropped";

/// The values below are independently derived by the task's Python ZIP/XML
/// oracle.  They are intentionally in source rather than loaded from the
/// evidence bundle, so a report cannot silently change when a sidecar file is
/// edited after the executable is built.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct NativeOracle {
    fixture: &'static str,
    path: &'static str,
    name: &'static str,
    resave_scope: &'static str,
    archive_sha256: &'static str,
    archive_bytes: usize,
    slide: usize,
    image: usize,
    image_count: usize,
    shape_position: usize,
    shape_id: u32,
    shape_name: &'static str,
    bounds: Option<BoundsOracle>,
    relationship_id: &'static str,
    part: &'static str,
    content_type: &'static str,
    payload_bytes: usize,
    payload_sha256: &'static str,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
struct BoundsOracle {
    x: i64,
    y: i64,
    width: i64,
    height: i64,
}

const POI_SLIDE: NativeOracle = NativeOracle {
    fixture: "poi-slide",
    path: "test-data/poi/test-data/slideshow/bug62513.pptx",
    name: "bug62513.pptx",
    resave_scope: "Apache POI producer fixture; retained as supplied",
    archive_sha256: "cd841112bd5b53f21e8434d080af8b2eeca78a07bc820e9ea0b97a0962c85c79",
    archive_bytes: 384_775,
    slide: 4,
    image: 0,
    image_count: 1,
    shape_position: 1,
    shape_id: 37_890,
    shape_name: "Picture 2",
    bounds: Some(BoundsOracle {
        x: 1_115_616,
        y: 2_276_872,
        width: 7_017_380,
        height: 3_262_858,
    }),
    relationship_id: "rId2",
    part: "/ppt/media/image2.jpeg",
    content_type: "image/jpeg",
    payload_bytes: 21_997,
    payload_sha256: "d5f10480ab75ce1175eea7432e684e9fac7319a7b9bd15512759385eae4842db",
};

const POI_VIDEO: NativeOracle = NativeOracle {
    fixture: "poi-video",
    path: "test-data/poi/test-data/slideshow/EmbeddedVideo.pptx",
    name: "EmbeddedVideo.pptx",
    resave_scope: "Apache POI producer fixture with embedded-video poster image; retained as supplied",
    archive_sha256: "7940e3b1a339db11f00b65399a2fe77e0e85a5da3a30ac8d6c8a0a77527b2ab2",
    archive_bytes: 201_418,
    slide: 0,
    image: 0,
    image_count: 1,
    shape_position: 0,
    shape_id: 2,
    shape_name: "file_example_MP4_480_1_5MG_Trim",
    bounds: Some(BoundsOracle {
        x: 3_810_000,
        y: 2_143_125,
        width: 4_572_000,
        height: 2_571_750,
    }),
    relationship_id: "rId4",
    part: "/ppt/media/image1.png",
    content_type: "image/png",
    payload_bytes: 65_215,
    payload_sha256: "f5516c6cae484df63ce03db77fb69b778660916b9207de5a4e04aa5e3b72908d",
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Fixture {
    PoiSlide,
    PoiVideo,
}

impl Fixture {
    const fn oracle(self) -> NativeOracle {
        match self {
            Self::PoiSlide => POI_SLIDE,
            Self::PoiVideo => POI_VIDEO,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Provider {
    Bytes,
    File,
    Range,
}

impl Provider {
    const fn name(self) -> &'static str {
        match self {
            Self::Bytes => "bytes",
            Self::File => "file",
            Self::Range => "range",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Config {
    fixture: Fixture,
    provider: Provider,
    max_range: Option<usize>,
    delay_us: Option<u64>,
    samples: usize,
    warmup: usize,
    source_revision: String,
    output: PathBuf,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct TimingRecord {
    open_ns: u64,
    metadata_ns: u64,
    read_ns: u64,
    api_sum_ns: u64,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ReadDelta {
    logical_calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    request_size_counts: [u64; pptx_range_source::PPTX_RANGE_REQUEST_SIZE_BUCKETS],
    short_reads: u64,
    delayed_calls: u64,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ReadPoint {
    availability: &'static str,
    unavailable_reason: Option<&'static str>,
    logical_calls: Option<u64>,
    requested_bytes: Option<u64>,
    returned_bytes: Option<u64>,
    request_size_counts: Option<[u64; pptx_range_source::PPTX_RANGE_REQUEST_SIZE_BUCKETS]>,
    min_request_bytes: Option<u64>,
    max_request_bytes: Option<u64>,
    short_reads: Option<u64>,
    delayed_calls: Option<u64>,
    counter_delta_checked: Option<bool>,
    delta: Option<ReadDelta>,
}

impl ReadPoint {
    const fn unavailable(reason: &'static str) -> Self {
        Self {
            availability: "unavailable",
            unavailable_reason: Some(reason),
            logical_calls: None,
            requested_bytes: None,
            returned_bytes: None,
            request_size_counts: None,
            min_request_bytes: None,
            max_request_bytes: None,
            short_reads: None,
            delayed_calls: None,
            counter_delta_checked: None,
            delta: None,
        }
    }

    fn available(
        snapshot: pptx_range_source::PptxRangeSourceSnapshot,
        delta: Option<pptx_range_source::PptxRangeSourceDelta>,
    ) -> Self {
        let delta = delta.map(|delta| ReadDelta {
            logical_calls: delta.logical_calls,
            requested_bytes: delta.requested_bytes,
            returned_bytes: delta.returned_bytes,
            request_size_counts: delta.request_size_counts,
            short_reads: delta.short_reads,
            delayed_calls: delta.delayed_calls,
        });
        Self {
            availability: "available",
            unavailable_reason: None,
            logical_calls: Some(snapshot.logical_calls),
            requested_bytes: Some(snapshot.requested_bytes),
            returned_bytes: Some(snapshot.returned_bytes),
            request_size_counts: Some(snapshot.request_size_counts),
            min_request_bytes: snapshot.min_request_bytes,
            max_request_bytes: snapshot.max_request_bytes,
            short_reads: Some(snapshot.short_reads),
            delayed_calls: Some(snapshot.delayed_calls),
            counter_delta_checked: delta.as_ref().map(|_| true),
            delta,
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PhaseDescription {
    label: &'static str,
    live_owners: &'static str,
}

const PHASES: [PhaseDescription; 8] = [
    PhaseDescription {
        label: "baseline",
        live_owners: "fixed native archive and caller budget only; no source adapter or package owner",
    },
    PhaseDescription {
        label: "opened",
        live_owners: "source-backed presentation view and caller source adapter",
    },
    PhaseDescription {
        label: "selected",
        live_owners: "source-backed presentation view, selected slide handle, and caller source adapter",
    },
    PhaseDescription {
        label: "loaded",
        live_owners: "source-backed presentation view, selected slide, returned image payload, and caller source adapter",
    },
    PhaseDescription {
        label: "drop_view",
        live_owners: "selected slide and returned image payload retain the package owner after the public view is dropped",
    },
    PhaseDescription {
        label: "drop_slide",
        live_owners: "returned image payload and its PartData remain after the selected slide handle is dropped",
    },
    PhaseDescription {
        label: "drop_image",
        live_owners: "caller source adapter only after the package owner and returned image are dropped",
    },
    PhaseDescription {
        label: "drop_source",
        live_owners: "no lifecycle-owned source or package owner; caller budget remains observable",
    },
];

#[derive(Clone, Debug, Serialize)]
struct PhaseRecord {
    label: &'static str,
    source_cache: Value,
    source_reads: ReadPoint,
    source_budget: Value,
    rss: Value,
}

#[derive(Clone, Debug, Serialize)]
struct PayloadCheck {
    bytes: usize,
    sha256: String,
    verified: bool,
    verified_after_slide_drop: bool,
}

#[derive(Clone, Debug, Serialize)]
struct LifecycleRow {
    sample_index: usize,
    timings: TimingRecord,
    phases: Vec<PhaseRecord>,
    descriptor_verified: bool,
    returned_descriptor_equal: bool,
    payload: PayloadCheck,
    final_memory_objects_depth_zero: bool,
    source_owner_release_verified: bool,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PayloadOracleRecord {
    slide: usize,
    image: usize,
    image_count: usize,
    shape_position: usize,
    shape_id: u32,
    shape_name: &'static str,
    bounds: Option<BoundsOracle>,
    relationship_id: &'static str,
    part: &'static str,
    content_type: &'static str,
    payload_bytes: usize,
    payload_sha256: &'static str,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct SourceSnapshotRecord {
    adapter: &'static str,
    scope: &'static str,
    checked_phase_delta: bool,
    unavailable_after_caller_source_drop: bool,
}

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

#[derive(Clone, Copy, Debug, Serialize)]
struct ProviderRecord {
    provider: &'static str,
    max_range_bytes: Option<usize>,
    delay_us: Option<u64>,
    delay_configured: bool,
    file_scope: &'static str,
    range_scope: &'static str,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    fixture: &'static str,
    provider: &'static str,
    native_fixture_sha256: String,
    native_fixture_bytes: usize,
    native_fixture_path: String,
    native_fixture_name: String,
    native_fixture_resave_scope: String,
    payload_oracle: PayloadOracleRecord,
    source_snapshot: SourceSnapshotRecord,
    provider_config: ProviderRecord,
    provider_scope: &'static str,
    timing_scope: &'static str,
    file_scope: &'static str,
    range_scope: &'static str,
    rss_scope: &'static str,
    samples: usize,
    warmup: usize,
    checked_iteration_count: usize,
    source_revision: String,
    binary_sha256: String,
    binary_bytes: u64,
    current_exe: String,
    configured_limits: ConfiguredLimits,
    phases: Vec<PhaseDescription>,
    samples_raw: Vec<LifecycleRow>,
}

#[derive(Debug)]
struct ProviderSource {
    adapter: Arc<pptx_range_source::PptxRangeSource>,
}

fn fixture_path(oracle: NativeOracle) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../")
        .join(oracle.path)
}

fn parse_bounded_usize(
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

fn parse_bounded_u64(
    value: &str,
    name: &str,
    minimum: u64,
    maximum: u64,
) -> Result<u64, Box<dyn Error>> {
    let parsed = value
        .parse::<u64>()
        .map_err(|_| format!("--{name} must be an unsigned decimal integer"))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(format!("--{name} must be between {minimum} and {maximum}").into());
    }
    Ok(parsed)
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut fixture = None;
    let mut provider = None;
    let mut max_range = None;
    let mut delay_us = None;
    let mut samples = None;
    let mut warmup = None;
    let mut source_revision = None;
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("native-image-lifecycle argument is not valid UTF-8")?;
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("missing value for {flag}"))?;
        let value_text = value
            .to_str()
            .ok_or("native-image-lifecycle argument value is not valid UTF-8")?;
        index += 2;
        match flag {
            "--fixture" => {
                if fixture.is_some() {
                    return Err("duplicate --fixture".into());
                }
                fixture = Some(match value_text {
                    "poi-slide" => Fixture::PoiSlide,
                    "poi-video" => Fixture::PoiVideo,
                    _ => return Err("--fixture must be poi-slide or poi-video".into()),
                });
            },
            "--provider" => {
                if provider.is_some() {
                    return Err("duplicate --provider".into());
                }
                provider = Some(match value_text {
                    "bytes" => Provider::Bytes,
                    "file" => Provider::File,
                    "range" => Provider::Range,
                    _ => return Err("--provider must be bytes, file, or range".into()),
                });
            },
            "--max-range" => {
                if max_range.is_some() {
                    return Err("duplicate --max-range".into());
                }
                max_range = Some(parse_bounded_usize(
                    value_text,
                    "max-range",
                    1,
                    MAX_RANGE_BYTES,
                )?);
            },
            "--delay-us" => {
                if delay_us.is_some() {
                    return Err("duplicate --delay-us".into());
                }
                delay_us = Some(parse_bounded_u64(value_text, "delay-us", 0, MAX_DELAY_US)?);
            },
            "--samples" => {
                if samples.is_some() {
                    return Err("duplicate --samples".into());
                }
                samples = Some(parse_bounded_usize(value_text, "samples", 1, MAX_SAMPLES)?);
            },
            "--warmup" => {
                if warmup.is_some() {
                    return Err("duplicate --warmup".into());
                }
                warmup = Some(parse_bounded_usize(value_text, "warmup", 0, MAX_WARMUP)?);
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
            _ => return Err(format!("unknown native-image-lifecycle argument: {flag}").into()),
        }
    }

    let provider = provider.ok_or("missing --provider")?;
    match provider {
        Provider::Range if max_range.is_none() || delay_us.is_none() => {
            return Err("--provider range requires --max-range and --delay-us".into());
        },
        Provider::Bytes | Provider::File if max_range.is_some() || delay_us.is_some() => {
            return Err("--max-range and --delay-us are valid only with --provider range".into());
        },
        Provider::Range | Provider::Bytes | Provider::File => {},
    }
    Ok(Config {
        fixture: fixture.ok_or("missing --fixture")?,
        provider,
        max_range,
        delay_us,
        samples: samples.ok_or("missing --samples")?,
        warmup: warmup.ok_or("missing --warmup")?,
        source_revision: source_revision.ok_or("missing --source-revision")?,
        output: output.ok_or("missing --output")?,
    })
}

fn verify_bytes(
    bytes: &[u8],
    oracle: NativeOracle,
    stage: &'static str,
) -> Result<(), Box<dyn Error>> {
    if bytes.len() != oracle.archive_bytes {
        return Err(format!(
            "native {} fixture length mismatch at {stage}: observed {}, expected {}",
            oracle.fixture,
            bytes.len(),
            oracle.archive_bytes
        )
        .into());
    }
    let digest = crate::sha256_hex(bytes);
    if digest != oracle.archive_sha256 {
        return Err(format!(
            "native {} fixture SHA-256 mismatch at {stage}: observed {digest}, expected {}",
            oracle.fixture, oracle.archive_sha256
        )
        .into());
    }
    Ok(())
}

fn read_and_verify_fixture(oracle: NativeOracle) -> Result<Vec<u8>, Box<dyn Error>> {
    let bytes = fs::read(fixture_path(oracle))?;
    verify_bytes(&bytes, oracle, "setup")?;
    Ok(bytes)
}

fn verify_fixture_again(oracle: NativeOracle) -> Result<(), Box<dyn Error>> {
    let path = fixture_path(oracle);
    let bytes = fs::read(path)?;
    verify_bytes(&bytes, oracle, "file-end")?;
    Ok(())
}

fn provider_source(
    config: &Config,
    oracle: NativeOracle,
    fixture_bytes: &[u8],
) -> Result<ProviderSource, Box<dyn Error>> {
    let inner: Arc<dyn ReadAt> = match config.provider {
        Provider::Bytes | Provider::Range => Arc::new(OwnedSource::new(fixture_bytes.to_vec())),
        Provider::File => Arc::new(FileSource::open(fixture_path(oracle))?),
    };
    let delay = config.delay_us.map(Duration::from_micros);
    let adapter = Arc::new(pptx_range_source::PptxRangeSource::with_limits(
        inner,
        config.max_range,
        delay,
    ));
    Ok(ProviderSource { adapter })
}

fn duration_ns(started: Instant) -> Result<u64, Box<dyn Error>> {
    u64::try_from(started.elapsed().as_nanos())
        .map_err(|_| "API duration does not fit in u64 nanoseconds".into())
}

fn sum_durations(timing: &TimingRecord) -> Result<u64, Box<dyn Error>> {
    timing
        .open_ns
        .checked_add(timing.metadata_ns)
        .and_then(|value| value.checked_add(timing.read_ns))
        .ok_or_else(|| "native API duration sum overflowed u64".into())
}

fn cache_unavailable(reason: &'static str) -> Result<Value, Box<dyn Error>> {
    Ok(serde_json::to_value(
        pptx_cache_retention::CachePoint::unavailable(reason),
    )?)
}

fn cache_available(
    view: &litchi_pptx::SourceBackedPresentation,
    budget: &litchi_core::Budget,
    cache_limits: SourceCacheLimits,
    previous: &mut Option<litchi_opc::SourceCacheDiagnostics>,
    label: &'static str,
) -> Result<Value, Box<dyn Error>> {
    let diagnostics = view.try_cache_diagnostics()?;
    let (point, after) = pptx_cache_retention::cache_point(
        diagnostics,
        *previous,
        budget,
        cache_limits,
        Some(label),
    )?;
    *previous = Some(after);
    Ok(serde_json::to_value(point)?)
}

fn budget_value(budget: &litchi_core::Budget) -> Result<Value, Box<dyn Error>> {
    Ok(serde_json::to_value(pptx_cache_retention::budget_point(
        budget,
    ))?)
}

fn rss_value() -> Result<Value, Box<dyn Error>> {
    Ok(serde_json::to_value(pptx_cache_retention::rss_point())?)
}

fn read_point(
    source: Option<&ProviderSource>,
    previous: &mut Option<pptx_range_source::PptxRangeSourceSnapshot>,
) -> Result<ReadPoint, Box<dyn Error>> {
    let Some(source) = source else {
        *previous = None;
        return Ok(ReadPoint::unavailable(
            "caller source adapter has been dropped",
        ));
    };
    let snapshot = source.adapter.snapshot()?;
    let delta = previous
        .map(|before| snapshot.checked_delta(before))
        .transpose()?;
    *previous = Some(snapshot);
    Ok(ReadPoint::available(snapshot, delta))
}

fn phase(
    label: &'static str,
    source_cache: Value,
    source: Option<&ProviderSource>,
    read_previous: &mut Option<pptx_range_source::PptxRangeSourceSnapshot>,
    budget: &litchi_core::Budget,
) -> Result<PhaseRecord, Box<dyn Error>> {
    Ok(PhaseRecord {
        label,
        source_cache,
        source_reads: read_point(source, read_previous)?,
        source_budget: budget_value(budget)?,
        rss: rss_value()?,
    })
}

fn payload_oracle_record(oracle: NativeOracle) -> PayloadOracleRecord {
    PayloadOracleRecord {
        slide: oracle.slide,
        image: oracle.image,
        image_count: oracle.image_count,
        shape_position: oracle.shape_position,
        shape_id: oracle.shape_id,
        shape_name: oracle.shape_name,
        bounds: oracle.bounds,
        relationship_id: oracle.relationship_id,
        part: oracle.part,
        content_type: oracle.content_type,
        payload_bytes: oracle.payload_bytes,
        payload_sha256: oracle.payload_sha256,
    }
}

fn verify_descriptor_fields(
    descriptor: &litchi_pptx::SourceImageDescriptor,
    oracle: NativeOracle,
) -> Result<(), Box<dyn Error>> {
    if descriptor.position() != oracle.image {
        return Err(format!(
            "native {} image position mismatch: observed {}, expected {}",
            oracle.fixture,
            descriptor.position(),
            oracle.image
        )
        .into());
    }
    if descriptor.shape_position() != oracle.shape_position {
        return Err(format!(
            "native {} shape position mismatch: observed {}, expected {}",
            oracle.fixture,
            descriptor.shape_position(),
            oracle.shape_position
        )
        .into());
    }
    if descriptor.id() != Some(oracle.shape_id) {
        return Err(format!(
            "native {} shape ID mismatch: observed {:?}, expected {}",
            oracle.fixture,
            descriptor.id(),
            oracle.shape_id
        )
        .into());
    }
    if descriptor.name() != Some(oracle.shape_name) {
        return Err(format!(
            "native {} shape name mismatch: observed {:?}, expected {:?}",
            oracle.fixture,
            descriptor.name(),
            oracle.shape_name
        )
        .into());
    }
    let bounds = descriptor.bounds().map(|bounds| BoundsOracle {
        x: bounds.x(),
        y: bounds.y(),
        width: bounds.width(),
        height: bounds.height(),
    });
    if bounds != oracle.bounds {
        return Err(format!(
            "native {} bounds mismatch: observed {:?}, expected {:?}",
            oracle.fixture, bounds, oracle.bounds
        )
        .into());
    }
    if descriptor.relationship_id() != oracle.relationship_id {
        return Err(format!(
            "native {} relationship mismatch: observed {}, expected {}",
            oracle.fixture,
            descriptor.relationship_id(),
            oracle.relationship_id
        )
        .into());
    }
    let part = descriptor
        .target()
        .part_uri()
        .ok_or("native image descriptor is not an internal image")?
        .as_str();
    if part != oracle.part {
        return Err(format!(
            "native {} image part mismatch: observed {part}, expected {}",
            oracle.fixture, oracle.part
        )
        .into());
    }
    if descriptor.target().content_type() != Some(oracle.content_type) {
        return Err(format!(
            "native {} content type mismatch: observed {:?}, expected {:?}",
            oracle.fixture,
            descriptor.target().content_type(),
            oracle.content_type
        )
        .into());
    }
    Ok(())
}

fn verify_descriptor(
    descriptor: &litchi_pptx::SourceImageDescriptor,
    descriptors: &[litchi_pptx::SourceImageDescriptor],
    oracle: NativeOracle,
) -> Result<(), Box<dyn Error>> {
    if descriptors.len() != oracle.image_count {
        return Err(format!(
            "native {} image count mismatch: observed {}, expected {}",
            oracle.fixture,
            descriptors.len(),
            oracle.image_count
        )
        .into());
    }
    verify_descriptor_fields(descriptor, oracle)
}

fn verify_payload(
    payload: &[u8],
    oracle: NativeOracle,
    stage: &'static str,
) -> Result<PayloadCheck, Box<dyn Error>> {
    let digest = crate::sha256_hex(payload);
    if payload.len() != oracle.payload_bytes || digest != oracle.payload_sha256 {
        return Err(format!(
            "native {} image payload mismatch at {stage}: observed {} bytes/{digest}, expected {} bytes/{}",
            oracle.fixture,
            payload.len(),
            oracle.payload_bytes,
            oracle.payload_sha256
        )
        .into());
    }
    Ok(PayloadCheck {
        bytes: payload.len(),
        sha256: digest,
        verified: true,
        verified_after_slide_drop: stage == "after-slide-drop",
    })
}

fn verify_zero_gauges(
    budget: &litchi_core::Budget,
    stage: &'static str,
) -> Result<(), Box<dyn Error>> {
    let gauges = [
        ("Memory", Resource::Memory),
        ("Objects", Resource::Objects),
        ("Depth", Resource::Depth),
    ];
    for (name, resource) in gauges {
        let used = budget.used(resource);
        if used != 0 {
            return Err(
                format!("native lifecycle {stage} retained managed {name} usage {used}").into(),
            );
        }
    }
    Ok(())
}

fn run_iteration(
    config: &Config,
    oracle: NativeOracle,
    fixture_bytes: &[u8],
    sample_index: usize,
) -> Result<LifecycleRow, Box<dyn Error>> {
    let managed = pptx_cache_retention::managed_context("native-pptx-image")?;
    let cache_limits = SourceCacheLimits::new(
        pptx_cache_retention::CACHE_LIMIT_BYTES,
        pptx_cache_retention::CACHE_LIMIT_ENTRIES,
    )?;
    let mut read_previous = None;
    let mut cache_previous = None;
    let mut phases = Vec::with_capacity(PHASES.len());
    phases.push(phase(
        "baseline",
        cache_unavailable("source-backed presentation has not been opened")?,
        None,
        &mut read_previous,
        &managed.budget,
    )?);

    let mut source = Some(provider_source(config, oracle, fixture_bytes)?);
    let source_weak = Arc::downgrade(
        &source
            .as_ref()
            .ok_or("native source unexpectedly absent before open")?
            .adapter,
    );
    read_previous = Some(pptx_range_source::PptxRangeSourceSnapshot::default());
    let source_for_view: Arc<dyn ReadAt> = source
        .as_ref()
        .ok_or("native source unexpectedly absent before open")?
        .adapter
        .clone();
    let opened = {
        let started = Instant::now();
        let view = litchi_pptx::SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source_for_view,
            ReadLimits::default(),
            cache_limits,
            managed.context.clone(),
        );
        let elapsed = duration_ns(started)?;
        (view, elapsed)
    };
    let view = opened.0?;
    let open_ns = opened.1;
    phases.push(phase(
        "opened",
        cache_available(
            &view,
            &managed.budget,
            cache_limits,
            &mut cache_previous,
            "opened",
        )?,
        source.as_ref(),
        &mut read_previous,
        &managed.budget,
    )?);

    let slide = view
        .slide(oracle.slide)
        .ok_or("native fixture selected slide is out of bounds")?;

    let (descriptors, metadata_ns) = {
        let started = Instant::now();
        let descriptors = slide.images();
        let elapsed = duration_ns(started)?;
        (descriptors?, elapsed)
    };
    let descriptor = descriptors
        .get(oracle.image)
        .cloned()
        .ok_or("native fixture selected image is out of bounds")?;
    verify_descriptor(&descriptor, &descriptors, oracle)?;
    drop(descriptors);
    phases.push(phase(
        "selected",
        cache_available(
            &view,
            &managed.budget,
            cache_limits,
            &mut cache_previous,
            "selected",
        )?,
        source.as_ref(),
        &mut read_previous,
        &managed.budget,
    )?);

    let (image, read_ns) = {
        let started = Instant::now();
        let image = slide.read_image(oracle.image);
        let elapsed = duration_ns(started)?;
        (image?, elapsed)
    };
    let returned_descriptor_equal = image.descriptor() == &descriptor;
    if !returned_descriptor_equal {
        return Err(format!(
            "native {} returned image descriptor differs from selected descriptor",
            oracle.fixture
        )
        .into());
    }
    verify_descriptor_fields(image.descriptor(), oracle)?;
    let payload = verify_payload(image.bytes(), oracle, "loaded")?;
    phases.push(phase(
        "loaded",
        cache_available(
            &view,
            &managed.budget,
            cache_limits,
            &mut cache_previous,
            "loaded",
        )?,
        source.as_ref(),
        &mut read_previous,
        &managed.budget,
    )?);

    drop(view);
    phases.push(phase(
        "drop_view",
        cache_unavailable("public source-backed presentation view was dropped")?,
        source.as_ref(),
        &mut read_previous,
        &managed.budget,
    )?);

    drop(slide);
    let payload_after_slide_drop = verify_payload(image.bytes(), oracle, "after-slide-drop")?;
    phases.push(phase(
        "drop_slide",
        cache_unavailable(
            "selected slide handle and public view were dropped; returned image retains only its payload owner",
        )?,
        source.as_ref(),
        &mut read_previous,
        &managed.budget,
    )?);

    drop(image);
    verify_zero_gauges(&managed.budget, "drop_image")?;
    phases.push(phase(
        "drop_image",
        cache_unavailable("all package-owned image and cache handles were dropped")?,
        source.as_ref(),
        &mut read_previous,
        &managed.budget,
    )?);

    if source_weak.strong_count() != 1 {
        return Err(format!(
            "native source adapter has an unexpected owner count before caller drop: {}",
            source_weak.strong_count()
        )
        .into());
    }
    drop(source.take());
    verify_zero_gauges(&managed.budget, "drop_source")?;
    phases.push(phase(
        "drop_source",
        cache_unavailable("caller source adapter was dropped")?,
        None,
        &mut read_previous,
        &managed.budget,
    )?);
    if source_weak.strong_count() != 0 {
        return Err(format!(
            "native source adapter retained a hidden package owner after caller drop: {}",
            source_weak.strong_count()
        )
        .into());
    }

    let mut timing = TimingRecord {
        open_ns,
        metadata_ns,
        read_ns,
        api_sum_ns: 0,
    };
    timing.api_sum_ns = sum_durations(&timing)?;
    let final_phase = phases.last().ok_or("native lifecycle has no final phase")?;
    let final_memory_objects_depth_zero = final_phase
        .source_budget
        .get("memory_used")
        .and_then(Value::as_u64)
        .is_some_and(|value| value == 0)
        && final_phase
            .source_budget
            .get("objects_used")
            .and_then(Value::as_u64)
            .is_some_and(|value| value == 0)
        && final_phase
            .source_budget
            .get("depth_used")
            .and_then(Value::as_u64)
            .is_some_and(|value| value == 0);
    if !final_memory_objects_depth_zero {
        return Err("native lifecycle final managed gauges were not zero".into());
    }
    Ok(LifecycleRow {
        sample_index,
        timings: timing,
        phases,
        descriptor_verified: true,
        returned_descriptor_equal,
        payload: PayloadCheck {
            bytes: payload.bytes,
            sha256: payload.sha256,
            verified: payload.verified,
            verified_after_slide_drop: payload_after_slide_drop.verified,
        },
        final_memory_objects_depth_zero,
        source_owner_release_verified: true,
    })
}

fn configured_limits() -> ConfiguredLimits {
    ConfiguredLimits {
        cache_max_bytes: pptx_cache_retention::CACHE_LIMIT_BYTES,
        cache_max_entries: pptx_cache_retention::CACHE_LIMIT_ENTRIES,
        memory_limit: pptx_cache_retention::MEMORY_LIMIT,
        input_bytes_limit: pptx_cache_retention::IO_LIMIT,
        output_bytes_limit: pptx_cache_retention::IO_LIMIT,
        work_limit: pptx_cache_retention::IO_LIMIT,
        objects_limit: pptx_cache_retention::OBJECT_LIMIT,
        depth_limit: pptx_cache_retention::DEPTH_LIMIT,
    }
}

fn run_capture(config: &Config) -> Result<Report, Box<dyn Error>> {
    let oracle = config.fixture.oracle();
    let fixture_path = fixture_path(oracle);
    if !fixture_path.is_file() {
        return Err(format!(
            "native fixture is not a regular file: {}",
            fixture_path.display()
        )
        .into());
    }
    let fixture_bytes = read_and_verify_fixture(oracle)?;
    let mut rows = Vec::with_capacity(config.samples);
    for index in 0..config.warmup {
        let _ = run_iteration(config, oracle, &fixture_bytes, index)?;
    }
    for index in 0..config.samples {
        rows.push(run_iteration(config, oracle, &fixture_bytes, index)?);
    }
    verify_fixture_again(oracle)?;
    let checked_iteration_count = config
        .warmup
        .checked_add(rows.len())
        .ok_or("native iteration count overflowed usize")?;
    let binary = crate::current_executable_identity()?;
    Ok(Report {
        schema: SCHEMA,
        fixture: oracle.fixture,
        provider: config.provider.name(),
        native_fixture_sha256: oracle.archive_sha256.to_owned(),
        native_fixture_bytes: oracle.archive_bytes,
        native_fixture_path: oracle.path.to_owned(),
        native_fixture_name: oracle.name.to_owned(),
        native_fixture_resave_scope: oracle.resave_scope.to_owned(),
        payload_oracle: payload_oracle_record(oracle),
        source_snapshot: SourceSnapshotRecord {
            adapter: "PptxRangeSource",
            scope: SOURCE_READ_SCOPE,
            checked_phase_delta: true,
            unavailable_after_caller_source_drop: true,
        },
        provider_config: ProviderRecord {
            provider: config.provider.name(),
            max_range_bytes: config.max_range,
            delay_us: config.delay_us,
            delay_configured: config.delay_us.is_some(),
            file_scope: FILE_SCOPE,
            range_scope: RANGE_SCOPE,
        },
        provider_scope: NATIVE_MEMORY_SCOPE,
        timing_scope: TIMING_SCOPE,
        file_scope: FILE_SCOPE,
        range_scope: RANGE_SCOPE,
        rss_scope: RSS_SCOPE,
        samples: config.samples,
        warmup: config.warmup,
        checked_iteration_count,
        source_revision: config.source_revision.clone(),
        binary_sha256: binary.binary_sha256.clone(),
        binary_bytes: binary.binary_bytes,
        current_exe: binary.path.clone(),
        configured_limits: configured_limits(),
        phases: PHASES.to_vec(),
        samples_raw: rows,
    })
}

/// Run the native selected-image lifecycle target after `main` has consumed
/// the `native-image-lifecycle` selector.  Reports use `create_new` output
/// semantics so an existing evidence file can never be silently replaced.
pub fn run_from_args<I>(args: I) -> Result<(), Box<dyn Error>>
where
    I: IntoIterator<Item = OsString>,
{
    let args = args.into_iter().collect::<Vec<_>>();
    let config = parse_config(&args)?;
    let report = serde_json::to_vec_pretty(&run_capture(&config)?)?;
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&config.output)?;
    output.write_all(&report)?;
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

    fn valid(provider: &str) -> Vec<OsString> {
        let mut values = args(&[
            "--fixture",
            "poi-slide",
            "--provider",
            provider,
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789abcdef0123456789abcdef01234567",
            "--output",
            "native.json",
        ]);
        if provider == "range" {
            values.extend(args(&["--max-range", "97", "--delay-us", "0"]));
        }
        values
    }

    #[test]
    fn config_accepts_each_provider_and_requires_range_controls() {
        assert!(parse_config(&valid("bytes")).is_ok());
        assert!(parse_config(&valid("file")).is_ok());
        assert!(parse_config(&valid("range")).is_ok());
        let mut missing = valid("range");
        missing.retain(|value| value != "--delay-us" && value != "0");
        assert!(parse_config(&missing).is_err());
    }

    #[test]
    fn config_rejects_range_controls_for_direct_providers() {
        let mut values = valid("bytes");
        values.extend(args(&["--max-range", "97"]));
        assert!(parse_config(&values).is_err());
        let mut values = valid("file");
        values.extend(args(&["--delay-us", "0"]));
        assert!(parse_config(&values).is_err());
    }

    #[test]
    fn config_rejects_duplicates_bounds_and_invalid_revision() {
        let mut duplicate = valid("bytes");
        duplicate.extend(args(&["--samples", "2"]));
        assert!(parse_config(&duplicate).is_err());
        let mut sample_zero = valid("bytes");
        sample_zero[5] = OsString::from("0");
        assert!(parse_config(&sample_zero).is_err());
        let mut range_large = valid("range");
        let max_range_index = range_large
            .iter()
            .position(|value| value == "97")
            .expect("range bound present");
        range_large[max_range_index] = OsString::from("1048577");
        assert!(parse_config(&range_large).is_err());
        let mut revision = valid("bytes");
        revision[9] = OsString::from("0123456789ABCDEF0123456789abcdef01234567");
        assert!(parse_config(&revision).is_err());
    }
    #[test]
    fn native_image_payload_survives_owner_drops_across_providers() -> Result<(), Box<dyn Error>> {
        for fixture in [Fixture::PoiSlide, Fixture::PoiVideo] {
            let oracle = fixture.oracle();
            let bytes = read_and_verify_fixture(oracle)?;
            for provider in ["bytes", "file", "range"] {
                let mut config = parse_config(&valid(provider))?;
                config.fixture = fixture;
                let row = run_iteration(&config, oracle, &bytes, 0)?;
                assert!(row.payload.verified && row.payload.verified_after_slide_drop);
                assert_eq!(row.payload.sha256, oracle.payload_sha256);
                assert!(row.final_memory_objects_depth_zero && row.source_owner_release_verified);
            }
        }
        Ok(())
    }

    #[test]
    fn shapes_original_and_resaved_keep_their_typed_refusals() -> Result<(), Box<dyn Error>> {
        for (relative, digest, original) in [
            (
                "test-data/ooxml/pptx/shapes.pptx",
                "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571",
                true,
            ),
            (
                "test-data/office-interop/libreoffice-resaved/shapes-litchi.pptx",
                "fcc8acffad88f5091316f67403c099a4c9eaa372e17927edfee29c35fb132034",
                false,
            ),
        ] {
            let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("../..")
                .join(relative);
            let bytes = fs::read(path)?;
            assert_eq!(crate::sha256_hex(&bytes), digest);
            let view = litchi_pptx::SourceBackedPresentation::from_read_at(Arc::new(
                OwnedSource::new(bytes),
            ))?;
            let error = view
                .slide(0)
                .ok_or("missing shapes slide")?
                .images()
                .expect_err("fixture must preserve its refusal");
            if original {
                assert!(
                    matches!(error, litchi_pptx::Error::Invalid(ref reason) if reason.contains("unsupported nested element")),
                    "{error:?}"
                );
            } else {
                assert!(
                    matches!(
                        error,
                        litchi_pptx::Error::UnsafeEdit {
                            operation: "source-backed picture inventory",
                            ..
                        }
                    ),
                    "{error:?}"
                );
            }
        }
        Ok(())
    }
}
