//! Matched PPTX source-provider lifecycle observations.
//!
//! This target runs the same source-backed cross-copy publication lifecycle
//! through an owned byte source, a freshly staged positional file source, or
//! the explicit range adapter.  It records source/cache/budget snapshots at
//! the nine ownership boundaries used by `pptx_cache_retention`, while the
//! three duration fields contain only the corresponding public API calls.
//! Corpus construction, staging, source-adapter construction, reservations,
//! diagnostics, and all correctness checks remain outside those clocks.

use std::{
    error::Error,
    ffi::OsString,
    fs::{self, OpenOptions},
    io::{self, Write},
    num::NonZeroU64,
    path::{Path, PathBuf},
    sync::Arc,
    time::{Duration, Instant},
};

use litchi_core::{Budget, FileSource, OwnedSource, ReadAt, Resource};
use litchi_opc::{ReadLimits, SourceCacheLimits};
use serde::Serialize;

use crate::{pptx_cache_retention, pptx_range_source};

const SCHEMA: &str = "pptx_provider_lifecycle_v1";
const MAX_RANGE_BYTES: usize = 1_048_576;
const MAX_DELAY_US: u64 = 100_000;
const PROVIDER_SCOPE: &str = "matched source-backed PPTX lifecycle through explicit bytes, warm recently-written file, or caller-owned bounded range adapter";
const TIMING_SCOPE: &str = "open, plan, and publication durations contain only their immediately surrounding public API calls; setup, source construction, reservations, diagnostics, drops, and correctness checks are outside the clocks";
const RANGE_SCOPE: &str = "PptxRangeSource logical adapter counters; request lengths and returned bytes describe caller-visible ReadAt calls, not physical network or filesystem I/O";
const FILE_SCOPE: &str = "fresh task-owned files are staged once before warmups and reopened through FileSource for every iteration; files are recently written and warm-cache, with no cold-cache claim";
const RSS_SCOPE: &str = "optional process-wide VmRSS/VmHWM snapshots from procfs; setup and unrelated process memory remain in scope and no comparative RSS claim is authorized";

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

    const fn case(self) -> crate::Case {
        match self {
            Self::Plain => crate::Case::PptxSourceBackedCrossCopyPlainLifecycle,
            Self::MediaRich => crate::Case::PptxSourceBackedCrossCopyMediaRichLifecycle,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ProviderKind {
    Bytes,
    File,
    Range,
}

impl ProviderKind {
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
    corpus: CorpusKind,
    provider: ProviderKind,
    max_range: Option<usize>,
    delay_us: Option<u64>,
    transfer_bytes_per_second: Option<NonZeroU64>,
    samples: usize,
    warmup: usize,
    source_revision: String,
    output: PathBuf,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ProviderConfigRecord {
    provider: &'static str,
    max_range_bytes: Option<usize>,
    delay_us: Option<u64>,
    transfer_bytes_per_second: Option<NonZeroU64>,
    adapter_max_range_bytes: Option<usize>,
    adapter_delay_us: Option<u64>,
    delay_configured: bool,
    bytes_file_unlimited_cap: bool,
    file_scope: &'static str,
    range_scope: &'static str,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct TimingRecord {
    open_source_ns: u64,
    open_destination_ns: u64,
    open_ns: u64,
    plan_ns: u64,
    publication_ns: u64,
    api_sum_ns: u64,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ReadDelta {
    logical_calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    min_request_bytes: Option<u64>,
    max_request_bytes: Option<u64>,
    short_reads: u64,
    delayed_calls: u64,
    transfer_paced_calls: u64,
    transfer_delay_ns: u64,
    request_size_counts: [u64; pptx_range_source::PPTX_RANGE_REQUEST_SIZE_BUCKETS],
}

#[derive(Clone, Copy, Debug, Serialize)]
struct ReadPoint {
    availability: &'static str,
    unavailable_reason: Option<&'static str>,
    logical_calls: Option<u64>,
    requested_bytes: Option<u64>,
    returned_bytes: Option<u64>,
    min_request_bytes: Option<u64>,
    max_request_bytes: Option<u64>,
    short_reads: Option<u64>,
    delayed_calls: Option<u64>,
    transfer_paced_calls: Option<u64>,
    transfer_delay_ns: Option<u64>,
    request_size_counts: Option<[u64; pptx_range_source::PPTX_RANGE_REQUEST_SIZE_BUCKETS]>,
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
            min_request_bytes: None,
            max_request_bytes: None,
            short_reads: None,
            delayed_calls: None,
            transfer_paced_calls: None,
            transfer_delay_ns: None,
            request_size_counts: None,
            counter_delta_checked: None,
            delta: None,
        }
    }

    fn available(snapshot: pptx_range_source::PptxRangeSourceSnapshot, delta: ReadDelta) -> Self {
        Self {
            availability: "available",
            unavailable_reason: None,
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
            counter_delta_checked: Some(true),
            delta: Some(delta),
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PhaseDescription {
    label: &'static str,
    live_owners: &'static str,
}

const PHASES: [PhaseDescription; 9] = [
    PhaseDescription {
        label: "baseline",
        live_owners: "fixed corpus only; independent caller budgets exist, source and destination cache owners are not opened",
    },
    PhaseDescription {
        label: "opened",
        live_owners: "source-backed presentation view, destination editor, caller source adapters, and reserved sink",
    },
    PhaseDescription {
        label: "planned",
        live_owners: "source-backed view, destination editor, source-retaining plan, caller source adapters, and sink",
    },
    PhaseDescription {
        label: "published",
        live_owners: "source-backed view, source-retaining plan, publication result, caller source adapters, and sink; destination editor consumed",
    },
    PhaseDescription {
        label: "drop_result",
        live_owners: "source-backed view, source-retaining plan, caller source adapters, and sink",
    },
    PhaseDescription {
        label: "drop_plan",
        live_owners: "source-backed view, caller source adapters, and sink",
    },
    PhaseDescription {
        label: "drop_view",
        live_owners: "caller source adapters and sink; public source-view diagnostic owner released",
    },
    PhaseDescription {
        label: "drop_caller_sources",
        live_owners: "sink only; caller source adapters released",
    },
    PhaseDescription {
        label: "drop_sink",
        live_owners: "no lifecycle-owned source, cache owner, plan, result, or sink handle",
    },
];

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
struct LifecycleRow {
    sample_index: usize,
    exact_output_verified: bool,
    output_sha256: String,
    output_bytes: usize,
    timings: TimingRecord,
    phases: Vec<PhaseRecord>,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    corpus: &'static str,
    provider: &'static str,
    provider_scope: &'static str,
    timing_scope: &'static str,
    range_scope: &'static str,
    file_scope: &'static str,
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
    provider_config: ProviderConfigRecord,
    configured_limits: ConfiguredLimits,
    destination_configured_limits: ConfiguredLimits,
    destination_editor_consumed_during_publish: bool,
    final_memory_objects_depth_zero_checked: bool,
    phases: Vec<PhaseDescription>,
    samples_raw: Vec<LifecycleRow>,
}

type ConfiguredLimits = ConfiguredLimitsLocal;

#[derive(Clone, Copy, Debug, Serialize)]
struct ConfiguredLimitsLocal {
    cache_max_bytes: usize,
    cache_max_entries: usize,
    memory_limit: u64,
    input_bytes_limit: u64,
    output_bytes_limit: u64,
    work_limit: u64,
    objects_limit: u64,
    depth_limit: u64,
}

type ProviderSource = pptx_range_source::PptxRangeSource;

#[derive(Debug)]
struct StagedFiles {
    root: PathBuf,
    source: PathBuf,
    destination: PathBuf,
}

impl Drop for StagedFiles {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.source);
        let _ = fs::remove_file(&self.destination);
        let _ = fs::remove_dir(&self.root);
    }
}

fn stage_files(
    source_bytes: &[u8],
    destination_bytes: &[u8],
    sample_index: usize,
) -> Result<StagedFiles, Box<dyn Error>> {
    let base = std::env::temp_dir();
    let mut root = None;
    for attempt in 0..128_u32 {
        let candidate = base.join(format!(
            "litchi-pptx-provider-{}-{}-{}",
            std::process::id(),
            sample_index,
            attempt
        ));
        match fs::create_dir(&candidate) {
            Ok(()) => {
                root = Some(candidate);
                break;
            },
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => {},
            Err(error) => return Err(error.into()),
        }
    }
    let root = root.ok_or("could not create a unique provider lifecycle task directory")?;
    let source = root.join("source.pptx");
    let destination = root.join("destination.pptx");
    let write_file = |path: &Path, bytes: &[u8]| -> Result<(), Box<dyn Error>> {
        let mut file = OpenOptions::new().write(true).create_new(true).open(path)?;
        file.write_all(bytes)?;
        file.flush()?;
        file.sync_all()?;
        Ok(())
    };
    if let Err(error) =
        write_file(&source, source_bytes).and_then(|()| write_file(&destination, destination_bytes))
    {
        let _ = fs::remove_file(&source);
        let _ = fs::remove_file(&destination);
        let _ = fs::remove_dir(&root);
        return Err(error);
    }
    Ok(StagedFiles {
        root,
        source,
        destination,
    })
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

fn parse_u64(value: &str, flag: &str) -> Result<u64, Box<dyn Error>> {
    let parsed = value
        .parse::<u64>()
        .map_err(|_| format!("{flag} must be an unsigned decimal integer"))?;
    if parsed > MAX_DELAY_US {
        return Err(format!("{flag} must be between 0 and {MAX_DELAY_US}").into());
    }
    Ok(parsed)
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut corpus = None;
    let mut provider = None;
    let mut max_range = None;
    let mut delay_us = None;
    let mut transfer_bytes_per_second = None;
    let mut samples = None;
    let mut warmup = None;
    let mut source_revision = None;
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("provider-lifecycle argument is not valid UTF-8")?;
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("missing value for {flag}"))?;
        let value_text = value
            .to_str()
            .ok_or("provider-lifecycle argument value is not valid UTF-8")?;
        index += 2;
        match flag {
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
            "--provider" => {
                if provider.is_some() {
                    return Err("duplicate --provider".into());
                }
                provider = Some(match value_text {
                    "bytes" => ProviderKind::Bytes,
                    "file" => ProviderKind::File,
                    "range" => ProviderKind::Range,
                    _ => return Err("--provider must be bytes, file, or range".into()),
                });
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
                delay_us = Some(parse_u64(value_text, "--delay-us")?);
            },
            "--transfer-bytes-per-second" => {
                if transfer_bytes_per_second.is_some() {
                    return Err("duplicate --transfer-bytes-per-second".into());
                }
                let rate = value_text.parse::<u64>()?;
                if !(1_048_576..=1_099_511_627_776).contains(&rate) {
                    return Err(
                        "--transfer-bytes-per-second must be between 1048576 and 1099511627776"
                            .into(),
                    );
                }
                transfer_bytes_per_second = NonZeroU64::new(rate);
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
            _ => return Err(format!("unknown provider-lifecycle argument: {flag}").into()),
        }
    }
    let corpus = corpus.ok_or("missing --corpus")?;
    let provider = provider.ok_or("missing --provider")?;
    let (max_range, delay_us) = match provider {
        ProviderKind::Range => (
            Some(max_range.ok_or("range provider requires --max-range")?),
            Some(delay_us.ok_or("range provider requires --delay-us")?),
        ),
        ProviderKind::Bytes | ProviderKind::File => {
            if transfer_bytes_per_second.is_some() {
                return Err(
                    "--transfer-bytes-per-second is valid only with --provider range".into(),
                );
            }
            if max_range.is_some() {
                return Err("--max-range is valid only with --provider range".into());
            }
            if delay_us.is_some() {
                return Err("--delay-us is valid only with --provider range".into());
            }
            (None, None)
        },
    };
    Ok(Config {
        corpus,
        provider,
        max_range,
        delay_us,
        transfer_bytes_per_second,
        samples: samples.ok_or("missing --samples")?,
        warmup: warmup.ok_or("missing --warmup")?,
        source_revision: source_revision.ok_or("missing --source-revision")?,
        output: output.ok_or("missing --output")?,
    })
}

fn provider_config(config: &Config) -> ProviderConfigRecord {
    ProviderConfigRecord {
        provider: config.provider.name(),
        max_range_bytes: config.max_range,
        delay_us: config.delay_us,
        transfer_bytes_per_second: config.transfer_bytes_per_second,
        adapter_max_range_bytes: config.max_range,
        adapter_delay_us: config.delay_us,
        delay_configured: config.delay_us.is_some(),
        bytes_file_unlimited_cap: matches!(
            config.provider,
            ProviderKind::Bytes | ProviderKind::File
        ),
        file_scope: FILE_SCOPE,
        range_scope: RANGE_SCOPE,
    }
}

fn configured_limits(output_limit: u64) -> ConfiguredLimits {
    ConfiguredLimitsLocal {
        cache_max_bytes: pptx_cache_retention::CACHE_LIMIT_BYTES,
        cache_max_entries: pptx_cache_retention::CACHE_LIMIT_ENTRIES,
        memory_limit: pptx_cache_retention::MEMORY_LIMIT,
        input_bytes_limit: pptx_cache_retention::IO_LIMIT,
        output_bytes_limit: output_limit,
        work_limit: pptx_cache_retention::IO_LIMIT,
        objects_limit: pptx_cache_retention::OBJECT_LIMIT,
        depth_limit: pptx_cache_retention::DEPTH_LIMIT,
    }
}

fn make_provider(
    config: &Config,
    bytes: &[u8],
    path: Option<&Path>,
) -> Result<Arc<ProviderSource>, Box<dyn Error>> {
    let inner: Arc<dyn ReadAt> = match config.provider {
        ProviderKind::Bytes | ProviderKind::Range => Arc::new(OwnedSource::new(bytes.to_vec())),
        ProviderKind::File => {
            let path = path.ok_or("file provider is missing its staged path")?;
            Arc::new(FileSource::open(path)?)
        },
    };
    let (max_range, delay_us) = match config.provider {
        ProviderKind::Bytes | ProviderKind::File => (None, None),
        ProviderKind::Range => (config.max_range, config.delay_us),
    };
    let mut adapter_config = pptx_range_source::PptxRangeSourceConfig::new(
        max_range,
        delay_us.map(Duration::from_micros),
    );
    adapter_config.transfer_bytes_per_second = config.transfer_bytes_per_second;
    Ok(Arc::new(ProviderSource::new(inner, adapter_config)))
}

fn read_point(
    source: Option<&ProviderSource>,
    before: Option<pptx_range_source::PptxRangeSourceSnapshot>,
) -> Result<
    (
        ReadPoint,
        Option<pptx_range_source::PptxRangeSourceSnapshot>,
    ),
    Box<dyn Error>,
> {
    let Some(source) = source else {
        return Ok((
            ReadPoint::unavailable("caller source has not been created or was dropped"),
            None,
        ));
    };
    let snapshot = source.snapshot()?;
    let delta = snapshot.checked_delta(before.unwrap_or_default())?;
    let delta = ReadDelta {
        logical_calls: delta.logical_calls,
        requested_bytes: delta.requested_bytes,
        returned_bytes: delta.returned_bytes,
        min_request_bytes: delta.min_request_bytes,
        max_request_bytes: delta.max_request_bytes,
        short_reads: delta.short_reads,
        delayed_calls: delta.delayed_calls,
        transfer_paced_calls: delta.transfer_paced_calls,
        transfer_delay_ns: delta.transfer_delay_ns,
        request_size_counts: delta.request_size_counts,
    };
    Ok((ReadPoint::available(snapshot, delta), Some(snapshot)))
}

fn phase(
    label: &'static str,
    source_cache: pptx_cache_retention::CachePoint,
    destination_cache: pptx_cache_retention::CachePoint,
    source: Option<&ProviderSource>,
    destination: Option<&ProviderSource>,
    source_before: Option<pptx_range_source::PptxRangeSourceSnapshot>,
    destination_before: Option<pptx_range_source::PptxRangeSourceSnapshot>,
    source_budget: &Budget,
    destination_budget: &Budget,
) -> Result<
    (
        PhaseRecord,
        Option<pptx_range_source::PptxRangeSourceSnapshot>,
        Option<pptx_range_source::PptxRangeSourceSnapshot>,
    ),
    Box<dyn Error>,
> {
    let (source_reads, source_after) = read_point(source, source_before)?;
    let (destination_reads, destination_after) = read_point(destination, destination_before)?;
    Ok((
        PhaseRecord {
            label,
            source_cache,
            destination_cache,
            source_reads,
            destination_reads,
            source_budget: pptx_cache_retention::budget_point(source_budget),
            destination_budget: pptx_cache_retention::budget_point(destination_budget),
            rss: pptx_cache_retention::rss_point(),
        },
        source_after,
        destination_after,
    ))
}

fn sink_ceiling(expected_bytes: usize) -> Result<u64, Box<dyn Error>> {
    u64::try_from(expected_bytes)?
        .checked_mul(2)
        .and_then(|value| value.checked_add(pptx_cache_retention::MAX_WRITE))
        .ok_or_else(|| "PPTX provider lifecycle sink ceiling overflows u64".into())
}

fn checked_sum(values: &[u64], label: &'static str) -> Result<u64, Box<dyn Error>> {
    values.iter().try_fold(0_u64, |total, value| {
        total
            .checked_add(*value)
            .ok_or_else(|| format!("{label} overflows u64").into())
    })
}

#[inline(never)]
fn run_lifecycle_iteration(
    config: &Config,
    corpus: &crate::PptxSourceBackedCrossCopyCorpus,
    sample_index: usize,
    staged: Option<&StagedFiles>,
) -> Result<LifecycleRow, Box<dyn Error>> {
    let source_context = pptx_cache_retention::managed_context("pptx-provider-lifecycle-source")?;
    let destination_context =
        pptx_cache_retention::managed_context("pptx-provider-lifecycle-destination")?;
    let cache_limits = SourceCacheLimits::new(
        pptx_cache_retention::CACHE_LIMIT_BYTES,
        pptx_cache_retention::CACHE_LIMIT_ENTRIES,
    )?;
    let mut phases = Vec::with_capacity(PHASES.len());
    let mut source_before = None;
    let mut destination_before = None;
    let (baseline, source_after, destination_after) = phase(
        "baseline",
        pptx_cache_retention::CachePoint::unavailable("source view has not been opened"),
        pptx_cache_retention::CachePoint::unavailable("destination editor has not been opened"),
        None,
        None,
        source_before,
        destination_before,
        &source_context.budget,
        &destination_context.budget,
    )?;
    phases.push(baseline);
    source_before = source_after;
    destination_before = destination_after;

    let source_provider = make_provider(
        config,
        &corpus.source_archive,
        staged.map(|paths| paths.source.as_path()),
    )?;
    let destination_provider = make_provider(
        config,
        &corpus.destination_archive,
        staged.map(|paths| paths.destination.as_path()),
    )?;
    let source_read: Arc<dyn ReadAt> = source_provider.clone();
    let destination_read: Arc<dyn ReadAt> = destination_provider.clone();
    let mut sink = crate::CountingSink::bounded(
        sink_ceiling(corpus.source_backed_expected_output.len())?,
        pptx_cache_retention::MAX_WRITE,
    );
    sink.reserve_budget()?;

    let source_started = Instant::now();
    let source_view =
        litchi_pptx::SourceBackedPresentation::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source_read,
            ReadLimits::default(),
            cache_limits,
            source_context.context.clone(),
        )?;
    let open_source_ns = crate::elapsed_ns(source_started.elapsed())?;
    let destination_started = Instant::now();
    let editor =
        litchi_pptx::SourceBackedPresentationEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
            destination_read,
            ReadLimits::default(),
            cache_limits,
            destination_context.context.clone(),
        )?;
    let open_destination_ns = crate::elapsed_ns(destination_started.elapsed())?;
    if source_view.slide_count() != crate::PPTX_CROSS_COPY_SOURCE_SLIDE_COUNT
        || editor.slide_count() != corpus.destination_slide_count
    {
        return Err("provider lifecycle opening changed prevalidated slide counts".into());
    }
    let source_opened = source_view.try_cache_diagnostics()?;
    let destination_opened = editor.try_cache_diagnostics()?;
    let (source_opened_point, source_opened) = pptx_cache_retention::cache_point(
        source_opened,
        None,
        &source_context.budget,
        cache_limits,
        None,
    )?;
    let (destination_opened_point, destination_opened) = pptx_cache_retention::cache_point(
        destination_opened,
        None,
        &destination_context.budget,
        cache_limits,
        None,
    )?;
    let (opened, source_after, destination_after) = phase(
        "opened",
        source_opened_point,
        destination_opened_point,
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        source_before,
        destination_before,
        &source_context.budget,
        &destination_context.budget,
    )?;
    phases.push(opened);
    source_before = source_after;
    destination_before = destination_after;

    let plan_started = Instant::now();
    let plan = editor.plan_cross_slide_copy(
        &source_view,
        corpus.source_slide,
        corpus.destination_slide,
        corpus.insertion_position,
    )?;
    let plan_ns = crate::elapsed_ns(plan_started.elapsed())?;
    if plan.source_position() != corpus.source_slide
        || plan.destination_slide_position() != corpus.destination_slide
        || plan.insertion_position() != corpus.insertion_position
        || plan.destination_slide_count() != corpus.destination_slide_count + 1
    {
        return Err("provider lifecycle plan metadata is not deterministic".into());
    }
    let source_planned = source_view.try_cache_diagnostics()?;
    let destination_planned = editor.try_cache_diagnostics()?;
    let (source_planned_point, source_planned) = pptx_cache_retention::cache_point(
        source_planned,
        Some(source_opened),
        &source_context.budget,
        cache_limits,
        Some("opened_to_planned"),
    )?;
    let (destination_planned_point, _destination_planned) = pptx_cache_retention::cache_point(
        destination_planned,
        Some(destination_opened),
        &destination_context.budget,
        cache_limits,
        Some("opened_to_planned"),
    )?;
    let (planned, source_after, destination_after) = phase(
        "planned",
        source_planned_point,
        destination_planned_point,
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        source_before,
        destination_before,
        &source_context.budget,
        &destination_context.budget,
    )?;
    phases.push(planned);
    source_before = source_after;
    destination_before = destination_after;

    let publication_started = Instant::now();
    let published = editor.publish_cross_slide_copy_to_stream(&mut sink, &plan)?;
    let publication_ns = crate::elapsed_ns(publication_started.elapsed())?;
    if published.destination_slide_count() != corpus.destination_slide_count + 1
        || published.insertion_position() != corpus.insertion_position
        || published.name() != corpus.source_slide_name
        || sink.bytes != corpus.source_backed_expected_output
        || sink.summary().accepted_bytes
            != u64::try_from(corpus.source_backed_expected_output.len())?
        || sink.summary().largest_write > pptx_cache_retention::MAX_WRITE
    {
        return Err("provider lifecycle publication differs from exact output oracle".into());
    }
    let source_published = source_view.try_cache_diagnostics()?;
    let (source_published_point, source_published) = pptx_cache_retention::cache_point(
        source_published,
        Some(source_planned),
        &source_context.budget,
        cache_limits,
        Some("planned_to_published"),
    )?;
    let (published_phase, source_after, destination_after) = phase(
        "published",
        source_published_point,
        pptx_cache_retention::CachePoint::unavailable(
            "destination editor was consumed by publication",
        ),
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        source_before,
        destination_before,
        &source_context.budget,
        &destination_context.budget,
    )?;
    phases.push(published_phase);
    source_before = source_after;
    destination_before = destination_after;

    drop(published);
    let source_after_result = source_view.try_cache_diagnostics()?;
    let (source_after_result_point, source_after_result) = pptx_cache_retention::cache_point(
        source_after_result,
        Some(source_published),
        &source_context.budget,
        cache_limits,
        Some("published_to_drop_result"),
    )?;
    let (drop_result, source_after, destination_after) = phase(
        "drop_result",
        source_after_result_point,
        pptx_cache_retention::CachePoint::unavailable(
            "destination editor was consumed by publication",
        ),
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        source_before,
        destination_before,
        &source_context.budget,
        &destination_context.budget,
    )?;
    phases.push(drop_result);
    source_before = source_after;
    destination_before = destination_after;

    drop(plan);
    let source_after_plan = source_view.try_cache_diagnostics()?;
    let (source_after_plan_point, _source_after_plan) = pptx_cache_retention::cache_point(
        source_after_plan,
        Some(source_after_result),
        &source_context.budget,
        cache_limits,
        Some("drop_result_to_drop_plan"),
    )?;
    let (drop_plan, source_after, destination_after) = phase(
        "drop_plan",
        source_after_plan_point,
        pptx_cache_retention::CachePoint::unavailable(
            "destination editor was consumed by publication",
        ),
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        source_before,
        destination_before,
        &source_context.budget,
        &destination_context.budget,
    )?;
    phases.push(drop_plan);
    source_before = source_after;
    destination_before = destination_after;

    drop(source_view);
    let (drop_view, source_after, destination_after) = phase(
        "drop_view",
        pptx_cache_retention::CachePoint::unavailable(
            "source view handle was dropped; no public source cache owner remains",
        ),
        pptx_cache_retention::CachePoint::unavailable(
            "destination editor was consumed by publication",
        ),
        Some(source_provider.as_ref()),
        Some(destination_provider.as_ref()),
        source_before,
        destination_before,
        &source_context.budget,
        &destination_context.budget,
    )?;
    phases.push(drop_view);
    source_before = source_after;
    destination_before = destination_after;

    drop(source_provider);
    drop(destination_provider);
    let (drop_sources, source_after, destination_after) = phase(
        "drop_caller_sources",
        pptx_cache_retention::CachePoint::unavailable("caller source owners were dropped"),
        pptx_cache_retention::CachePoint::unavailable(
            "destination editor was consumed by publication",
        ),
        None,
        None,
        source_before,
        destination_before,
        &source_context.budget,
        &destination_context.budget,
    )?;
    phases.push(drop_sources);
    source_before = source_after;
    destination_before = destination_after;

    drop(sink);
    let (drop_sink, _source_after, _destination_after) = phase(
        "drop_sink",
        pptx_cache_retention::CachePoint::unavailable("caller source owners were dropped"),
        pptx_cache_retention::CachePoint::unavailable(
            "destination editor was consumed by publication",
        ),
        None,
        None,
        source_before,
        destination_before,
        &source_context.budget,
        &destination_context.budget,
    )?;
    phases.push(drop_sink);

    for (name, budget) in [
        ("source", &source_context.budget),
        ("destination", &destination_context.budget),
    ] {
        for resource in [Resource::Memory, Resource::Objects, Resource::Depth] {
            if budget.used(resource) != 0 {
                return Err(format!(
                    "provider lifecycle {name} {resource:?} gauge remained nonzero after drop_sink"
                )
                .into());
            }
        }
    }
    let open_ns = checked_sum(&[open_source_ns, open_destination_ns], "open timing")?;
    let api_sum_ns = checked_sum(&[open_ns, plan_ns, publication_ns], "API timing")?;
    Ok(LifecycleRow {
        sample_index,
        exact_output_verified: true,
        output_sha256: crate::sha256_hex(&corpus.source_backed_expected_output),
        output_bytes: corpus.source_backed_expected_output.len(),
        timings: TimingRecord {
            open_source_ns,
            open_destination_ns,
            open_ns,
            plan_ns,
            publication_ns,
            api_sum_ns,
        },
        phases,
    })
}

fn run_capture(config: &Config) -> Result<Report, Box<dyn Error>> {
    let corpus = crate::build_pptx_source_backed_cross_copy_corpus(config.corpus.case())?;
    let gates = pptx_cache_retention::validate_gates(&corpus)?;
    let expected_digest = crate::sha256_hex(&corpus.source_backed_expected_output);
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or("warmup and samples overflow usize")?;
    let staged = if config.provider == ProviderKind::File {
        Some(stage_files(
            &corpus.source_archive,
            &corpus.destination_archive,
            0,
        )?)
    } else {
        None
    };
    let mut rows = Vec::with_capacity(config.samples);
    for iteration in 0..total {
        let row = run_lifecycle_iteration(config, &corpus, iteration, staged.as_ref())?;
        if iteration >= config.warmup {
            rows.push(LifecycleRow {
                sample_index: iteration - config.warmup,
                ..row
            });
        }
    }
    if rows.iter().any(|row| {
        !row.exact_output_verified
            || row.output_sha256 != expected_digest
            || row.output_bytes != corpus.source_backed_expected_output.len()
            || !row
                .phases
                .iter()
                .all(|phase| phase.source_budget.memory_used <= phase.source_budget.memory_limit)
    }) {
        return Err("provider lifecycle retained an invalid output or budget observation".into());
    }
    let binary = crate::current_executable_identity()?;
    Ok(Report {
        schema: SCHEMA,
        corpus: config.corpus.name(),
        provider: config.provider.name(),
        provider_scope: PROVIDER_SCOPE,
        timing_scope: TIMING_SCOPE,
        range_scope: RANGE_SCOPE,
        file_scope: FILE_SCOPE,
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
        provider_config: provider_config(config),
        configured_limits: configured_limits(pptx_cache_retention::IO_LIMIT),
        destination_configured_limits: configured_limits(pptx_cache_retention::IO_LIMIT),
        destination_editor_consumed_during_publish: true,
        final_memory_objects_depth_zero_checked: true,
        phases: PHASES.to_vec(),
        samples_raw: rows,
    })
}

/// Runs the provider lifecycle target after `main` has consumed the
/// `provider-lifecycle` selector.  The output path is opened with
/// `create_new`, so a completed evidence file can never be silently replaced.
pub fn run_from_args<I>(args: I) -> Result<(), Box<dyn Error>>
where
    I: IntoIterator<Item = OsString>,
{
    let args = args.into_iter().collect::<Vec<_>>();
    let config = parse_config(&args)?;
    let bytes = serde_json::to_vec_pretty(&run_capture(&config)?)?;
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

    const REVISION: &str = "0123456789abcdef0123456789abcdef01234567";

    #[test]
    fn parser_requires_range_controls_only_for_range_provider() {
        let base = [
            "--corpus",
            "plain",
            "--provider",
            "bytes",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            REVISION,
            "--output",
            "provider.json",
        ];
        assert_eq!(
            parse_config(&args(&base)).unwrap().provider,
            ProviderKind::Bytes
        );
        let mut range = base.to_vec();
        range[3] = "range";
        range.extend(["--max-range", "97", "--delay-us", "0"]);
        let config = parse_config(&args(&range)).unwrap();
        assert_eq!(config.provider, ProviderKind::Range);
        assert_eq!(config.max_range, Some(97));
        assert_eq!(config.delay_us, Some(0));
    }

    #[test]
    fn parser_rejects_unknown_duplicate_and_mismatched_arguments() {
        let base = [
            "--corpus",
            "plain",
            "--provider",
            "bytes",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            REVISION,
            "--output",
            "provider.json",
        ];
        let mut duplicate = base.to_vec();
        duplicate.extend(["--provider", "file"]);
        assert!(parse_config(&args(&duplicate)).is_err());
        let mut unknown = base.to_vec();
        unknown.extend(["--unknown", "value"]);
        assert!(parse_config(&args(&unknown)).is_err());
        let mut mismatched = base.to_vec();
        mismatched.extend(["--max-range", "97"]);
        assert!(parse_config(&args(&mismatched)).is_err());
    }

    #[test]
    fn provider_adapter_preserves_exact_bytes_and_records_zero_delay() {
        let inner: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(b"provider-test".to_vec()));
        let source = ProviderSource::with_limits(inner, None, Some(Duration::ZERO));
        let mut output = [0_u8; 8];
        let returned = source.read_at(1, &mut output).unwrap();
        assert_eq!(&output[..returned], b"rovider-");
        let snapshot = source.snapshot().unwrap();
        assert_eq!(snapshot.logical_calls, 1);
        assert_eq!(snapshot.returned_bytes, 8);
        assert_eq!(snapshot.delayed_calls, 1);
    }

    #[test]
    fn plain_lifecycle_output_identity_matches_all_providers() {
        let mut digests = Vec::new();
        for (provider, max_range, delay_us, transfer_bytes_per_second) in [
            (ProviderKind::Bytes, None, None, None),
            (ProviderKind::File, None, None, None),
            (ProviderKind::Range, Some(97), Some(0), None),
            (
                ProviderKind::Range,
                Some(65536),
                Some(0),
                NonZeroU64::new(26_214_400),
            ),
        ] {
            let report = run_capture(&Config {
                corpus: CorpusKind::Plain,
                provider,
                max_range,
                delay_us,
                transfer_bytes_per_second,
                samples: 1,
                warmup: 0,
                source_revision: REVISION.to_owned(),
                output: PathBuf::from("provider-lifecycle-test.json"),
            })
            .unwrap();
            assert_eq!(report.samples_raw.len(), 1);
            assert!(report.samples_raw[0].exact_output_verified);
            assert!(report.final_memory_objects_depth_zero_checked);
            digests.push(report.samples_raw[0].output_sha256.clone());
        }
        assert_eq!(digests[0], digests[1]);
        assert_eq!(digests[1], digests[2]);
        assert_eq!(digests[2], digests[3]);
    }
    #[test]
    fn transfer_rate_requires_range_provider_and_explicit_bounded_nonzero_value() {
        let base = [
            "--corpus",
            "plain",
            "--provider",
            "range",
            "--max-range",
            "65536",
            "--delay-us",
            "0",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            REVISION,
            "--output",
            "provider.json",
        ];
        for value in ["1048576", "26214400", "1099511627776"] {
            let mut values = base.to_vec();
            values.extend(["--transfer-bytes-per-second", value]);
            assert_eq!(
                parse_config(&args(&values))
                    .unwrap()
                    .transfer_bytes_per_second
                    .unwrap()
                    .get()
                    .to_string(),
                value
            );
            values.extend(["--transfer-bytes-per-second", value]);
            assert!(parse_config(&args(&values)).is_err());
        }
        for value in [
            "0",
            "1048575",
            "1099511627777",
            "-1",
            "18446744073709551616",
            "fast",
        ] {
            let mut values = base.to_vec();
            values.extend(["--transfer-bytes-per-second", value]);
            assert!(parse_config(&args(&values)).is_err());
        }
        for provider in ["bytes", "file"] {
            let values = [
                "--corpus",
                "plain",
                "--provider",
                provider,
                "--samples",
                "1",
                "--warmup",
                "0",
                "--source-revision",
                REVISION,
                "--output",
                "provider.json",
                "--transfer-bytes-per-second",
                "26214400",
            ];
            assert!(parse_config(&args(&values)).is_err());
        }
    }
}
