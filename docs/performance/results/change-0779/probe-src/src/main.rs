//! Standalone XLSX phase probe for change 0779.
//!
//! The probe deliberately lives in the evidence packet instead of the
//! production workspace or the ordinary-save harness. It exercises the same
//! public route as the retained ordinary-save XLSX case:
//! `Workbook::open`, a semantic `edit` setting `A1`, `commit`, and (for save
//! phases) `save_with_durability(..., Durability::NoSync)`.
//!
//! Allocation regions begin immediately before and finish immediately after
//! the selected clock interval. Reopen, hash, readback, and cleanup happen
//! after the region and clock have closed. The owner remains alive through the
//! region boundary, so `live_bytes_after` describes retained workbook state at
//! the end of the measured phase.

mod allocation_metrics;

#[cfg(feature = "allocator-metrics")]
mod counting_allocator;

use std::env;
use std::error::Error;
use std::fs;
use std::hint::black_box;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

use litchi_core::Durability;
use litchi_xlsx::{Cell, Value, Workbook};
use serde::Serialize;
use sha2::{Digest, Sha256};

const SCHEMA: &str = "litchi.xlsx.allocation-probe.v1";
const MARKER: &str = "litchi-perf-0638-ordinary-save";
const ADDRESS: &str = "A1";
const DEFAULT_SAMPLES: usize = 1;
const DEFAULT_WARMUP: usize = 0;
const MAX_SAMPLES: usize = 100_000;
const MAX_WARMUP: usize = 100_000;

static DESTINATION_COUNTER: AtomicU64 = AtomicU64::new(1);

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Phase {
    Open,
    Edit,
    Save,
    Lifecycle,
}

impl Phase {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "open" => Ok(Self::Open),
            "edit" => Ok(Self::Edit),
            "save" => Ok(Self::Save),
            "lifecycle" => Ok(Self::Lifecycle),
            _ => Err(
                format!("--phase must be one of open, edit, save, lifecycle (got {value:?})")
                    .into(),
            ),
        }
    }

    fn timing_scope(self) -> &'static str {
        match self {
            Self::Open => "Workbook::open(source) only",
            Self::Edit => {
                "Workbook::edit, sheet(sheet).set(A1, marker), and commit only; open is outside the clock"
            },
            Self::Save => {
                "Workbook::save_with_durability(destination, NoSync) only; open and edit are outside the clock"
            },
            Self::Lifecycle => {
                "Workbook::open, semantic A1 edit and commit, then save_with_durability(destination, NoSync)"
            },
        }
    }

    fn uses_destination(self) -> bool {
        matches!(self, Self::Save | Self::Lifecycle)
    }
}

#[derive(Debug)]
struct Config {
    source: PathBuf,
    sheet: String,
    phase: Phase,
    samples: usize,
    warmup: usize,
    output: PathBuf,
}

#[derive(Debug, Serialize)]
struct Identity {
    path: String,
    bytes: u64,
    sha256: String,
}

#[derive(Debug, Serialize)]
struct AllocatorIdentity {
    binary: &'static str,
    allocator: &'static str,
    instrumentation: &'static str,
    counter_revision: Option<&'static str>,
}

#[derive(Debug, Serialize)]
struct Verification {
    reopened: bool,
    sheet: String,
    address: &'static str,
    expected: &'static str,
    actual: Option<String>,
    marker_matches: bool,
}

#[derive(Debug, Serialize)]
struct SampleRecord {
    index: usize,
    elapsed_ns: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation: Option<allocation_metrics::Sample>,
    #[serde(skip_serializing_if = "Option::is_none")]
    published_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    published_sha256: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    verification: Option<Verification>,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    source: Identity,
    sheet: String,
    address: &'static str,
    marker: &'static str,
    phase: Phase,
    timing_scope: &'static str,
    durability: Option<&'static str>,
    warmup: usize,
    samples_requested: usize,
    samples: Vec<SampleRecord>,
    allocator: AllocatorIdentity,
}

fn main() -> Result<(), Box<dyn Error>> {
    #[cfg(feature = "allocator-metrics")]
    allocation_metrics::enable();

    let config = parse_args(env::args_os().skip(1))?;
    run(config)
}

fn run(config: Config) -> Result<(), Box<dyn Error>> {
    if config.source == config.output {
        return Err("--source and --output must name different files".into());
    }
    if !config.source.is_file() {
        return Err(format!("source is not a regular file: {}", config.source.display()).into());
    }
    if let Some(parent) = config.output.parent()
        && !parent.as_os_str().is_empty()
        && !parent.is_dir()
    {
        return Err(format!("output parent is not a directory: {}", parent.display()).into());
    }

    let source = source_identity(&config.source)?;
    let destination = config
        .phase
        .uses_destination()
        .then(|| destination_path(config.output.parent().unwrap_or_else(|| Path::new("."))));

    let mut records = Vec::with_capacity(config.samples);
    for _ in 0..config.warmup {
        let record = run_one(&config, destination.as_deref(), 0)?;
        black_box(record);
    }
    for index in 0..config.samples {
        records.push(run_one(&config, destination.as_deref(), index)?);
    }

    if let Some(destination) = destination.as_deref() {
        remove_destination(destination)?;
    }

    let report = Report {
        schema: SCHEMA,
        tool: "litchi-xlsx-allocation-attribution-probe-0779",
        source,
        sheet: config.sheet,
        address: ADDRESS,
        marker: MARKER,
        phase: config.phase,
        timing_scope: config.phase.timing_scope(),
        durability: config.phase.uses_destination().then_some("no-sync"),
        warmup: config.warmup,
        samples_requested: config.samples,
        samples: records,
        allocator: AllocatorIdentity {
            binary: allocation_metrics::binary_identity(),
            allocator: allocation_metrics::allocator_identity(),
            instrumentation: allocation_metrics::instrumentation_identity(),
            counter_revision: allocation_metrics::counter_revision(),
        },
    };
    let encoded = serde_json::to_vec_pretty(&report)?;
    fs::write(&config.output, encoded)?;
    Ok(())
}

fn run_one(
    config: &Config,
    destination: Option<&Path>,
    index: usize,
) -> Result<SampleRecord, Box<dyn Error>> {
    if let Some(destination) = destination {
        remove_destination(destination)?;
    }

    let (workbook, output, elapsed_ns, allocation) = match config.phase {
        Phase::Open => {
            let region = allocation_metrics::begin();
            let started = Instant::now();
            let workbook = Workbook::open(&config.source)?;
            let elapsed_ns = elapsed_ns(started.elapsed())?;
            let allocation = region.finish();
            (workbook, None, elapsed_ns, allocation)
        },
        Phase::Edit => {
            let workbook = Workbook::open(&config.source)?;
            let region = allocation_metrics::begin();
            let started = Instant::now();
            let edited = apply_edit(workbook, &config.sheet)?;
            let elapsed_ns = elapsed_ns(started.elapsed())?;
            let allocation = region.finish();
            (edited, None, elapsed_ns, allocation)
        },
        Phase::Save => {
            let workbook = Workbook::open(&config.source)?;
            let workbook = apply_edit(workbook, &config.sheet)?;
            let destination = destination.ok_or("save phase has no destination")?;
            let region = allocation_metrics::begin();
            let started = Instant::now();
            workbook.save_with_durability(destination, Durability::NoSync)?;
            let elapsed_ns = elapsed_ns(started.elapsed())?;
            let allocation = region.finish();
            (workbook, Some(destination), elapsed_ns, allocation)
        },
        Phase::Lifecycle => {
            let region = allocation_metrics::begin();
            let started = Instant::now();
            let workbook = Workbook::open(&config.source)?;
            let workbook = apply_edit(workbook, &config.sheet)?;
            let destination = destination.ok_or("lifecycle phase has no destination")?;
            workbook.save_with_durability(destination, Durability::NoSync)?;
            let elapsed_ns = elapsed_ns(started.elapsed())?;
            let allocation = region.finish();
            (workbook, Some(destination), elapsed_ns, allocation)
        },
    };
    // Keep the owner live across the region boundary and make its existence
    // visible to the optimizer before the post-clock verification begins.
    black_box(&workbook);

    let (published_bytes, published_sha256, verification) = match output {
        Some(path) => {
            let (bytes, sha256) = file_identity(path)?;
            let verification = verify_marker(path, &config.sheet)?;
            (Some(bytes), Some(sha256), Some(verification))
        },
        None if matches!(config.phase, Phase::Open) => {
            let reopened = Workbook::open(&config.source)?;
            verify_sheet_exists(&reopened, &config.sheet)?;
            (None, None, None)
        },
        None => {
            // A committed edit is serialized and independently reopened
            // outside the clock as a cheap guard against a benchmark that
            // silently measures a failed mutation.
            let verification = verify_committed_workbook(&workbook, &config.sheet)?;
            (None, None, Some(verification))
        },
    };
    drop(workbook);

    Ok(SampleRecord {
        index,
        elapsed_ns,
        allocation,
        published_bytes,
        published_sha256,
        verification,
    })
}

fn apply_edit(workbook: Workbook, sheet_name: &str) -> Result<Workbook, Box<dyn Error>> {
    let mut edit = workbook.edit()?;
    {
        let mut sheet = edit
            .sheet(sheet_name)?
            .ok_or_else(|| format!("selected worksheet is absent: {sheet_name}"))?;
        sheet.set(ADDRESS, MARKER)?;
    }
    let commit = edit.commit()?;
    if commit.patch().is_empty() {
        return Err("the edit produced an empty patch".into());
    }
    Ok(commit.into_workbook())
}

fn verify_marker(path: &Path, sheet_name: &str) -> Result<Verification, Box<dyn Error>> {
    let workbook = Workbook::open(path)?;
    let mut verification = verify_workbook_marker(&workbook, sheet_name)?;
    verification.reopened = true;
    Ok(verification)
}

fn verify_committed_workbook(
    workbook: &Workbook,
    sheet_name: &str,
) -> Result<Verification, Box<dyn Error>> {
    let reopened = Workbook::from_bytes(workbook.to_bytes()?)?;
    let mut verification = verify_workbook_marker(&reopened, sheet_name)?;
    verification.reopened = true;
    Ok(verification)
}

fn verify_sheet_exists(workbook: &Workbook, sheet_name: &str) -> Result<(), Box<dyn Error>> {
    if workbook.sheet(sheet_name)?.is_none() {
        return Err(
            format!("selected worksheet is absent during verification: {sheet_name}").into(),
        );
    }
    Ok(())
}

fn verify_workbook_marker(
    workbook: &Workbook,
    sheet_name: &str,
) -> Result<Verification, Box<dyn Error>> {
    let sheet = workbook
        .sheet(sheet_name)?
        .ok_or_else(|| format!("selected worksheet is absent during verification: {sheet_name}"))?;
    let actual = match sheet.cell(ADDRESS)?.stored() {
        Some(Cell::Value(Value::Text(value))) => Some(value.as_str().to_owned()),
        _ => None,
    };
    let marker_matches = actual.as_deref() == Some(MARKER);
    if !marker_matches {
        return Err(format!(
            "verification failed for {sheet_name}!{ADDRESS}: expected {MARKER:?}, got {actual:?}"
        )
        .into());
    }
    Ok(Verification {
        reopened: false,
        sheet: sheet_name.to_owned(),
        address: ADDRESS,
        expected: MARKER,
        actual,
        marker_matches,
    })
}

fn source_identity(path: &Path) -> Result<Identity, Box<dyn Error>> {
    let (bytes, sha256) = file_identity(path)?;
    Ok(Identity {
        path: path.display().to_string(),
        bytes,
        sha256,
    })
}

fn file_identity(path: &Path) -> Result<(u64, String), Box<dyn Error>> {
    let bytes = fs::read(path)?;
    let length = u64::try_from(bytes.len())?;
    Ok((length, sha256_hex(&bytes)))
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn elapsed_ns(duration: std::time::Duration) -> Result<u64, Box<dyn Error>> {
    u64::try_from(duration.as_nanos())
        .map_err(|_| "elapsed duration overflows u64 nanoseconds".into())
}

fn destination_path(parent: &Path) -> PathBuf {
    let serial = DESTINATION_COUNTER.fetch_add(1, Ordering::Relaxed);
    let pid = std::process::id();
    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_or(0, |duration| duration.as_nanos());
    parent.join(format!(
        ".litchi-xlsx-probe-{pid}-{timestamp}-{serial}.xlsx"
    ))
}

fn remove_destination(path: &Path) -> Result<(), Box<dyn Error>> {
    match fs::symlink_metadata(path) {
        Ok(metadata) if metadata.is_file() || metadata.file_type().is_symlink() => {
            fs::remove_file(path)?;
        },
        Ok(_) => return Err(format!("probe destination is not a file: {}", path.display()).into()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {},
        Err(error) => return Err(error.into()),
    }
    Ok(())
}

fn parse_args<I>(arguments: I) -> Result<Config, Box<dyn Error>>
where
    I: IntoIterator<Item = std::ffi::OsString>,
{
    let mut source = None;
    let mut sheet = None;
    let mut phase = None;
    let mut samples = DEFAULT_SAMPLES;
    let mut warmup = DEFAULT_WARMUP;
    let mut output = None;
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        let argument = argument.to_string_lossy();
        match argument.as_ref() {
            "--source" => source = Some(next_path(&mut arguments, "--source")?),
            "--sheet" => sheet = Some(next_string(&mut arguments, "--sheet")?),
            "--phase" => phase = Some(Phase::parse(&next_string(&mut arguments, "--phase")?)?),
            "--samples" => {
                samples = parse_count(
                    &next_string(&mut arguments, "--samples")?,
                    "samples",
                    MAX_SAMPLES,
                )?
            },
            "--warmup" => {
                warmup = parse_count(
                    &next_string(&mut arguments, "--warmup")?,
                    "warmup",
                    MAX_WARMUP,
                )?
            },
            "--output" => output = Some(next_path(&mut arguments, "--output")?),
            "--help" | "-h" => return Err(usage().into()),
            value => return Err(format!("unknown argument {value:?}\n{}", usage()).into()),
        }
    }
    let source = source.ok_or_else(|| format!("--source is required\n{}", usage()))?;
    let sheet = sheet.ok_or_else(|| format!("--sheet is required\n{}", usage()))?;
    let phase = phase.ok_or_else(|| format!("--phase is required\n{}", usage()))?;
    let output = output.ok_or_else(|| format!("--output is required\n{}", usage()))?;
    if samples == 0 {
        return Err("--samples must be at least 1".into());
    }
    Ok(Config {
        source,
        sheet,
        phase,
        samples,
        warmup,
        output,
    })
}

fn next_path<I>(arguments: &mut I, option: &str) -> Result<PathBuf, Box<dyn Error>>
where
    I: Iterator<Item = std::ffi::OsString>,
{
    arguments
        .next()
        .map(PathBuf::from)
        .ok_or_else(|| format!("{option} requires a path").into())
}

fn next_string<I>(arguments: &mut I, option: &str) -> Result<String, Box<dyn Error>>
where
    I: Iterator<Item = std::ffi::OsString>,
{
    arguments
        .next()
        .map(|value| value.to_string_lossy().into_owned())
        .ok_or_else(|| format!("{option} requires a value").into())
}

fn parse_count(value: &str, name: &str, maximum: usize) -> Result<usize, Box<dyn Error>> {
    let count = value
        .parse::<usize>()
        .map_err(|error| format!("--{name} must be a non-negative integer: {error}"))?;
    if count > maximum {
        return Err(format!("--{name} exceeds maximum {maximum}").into());
    }
    Ok(count)
}

fn usage() -> &'static str {
    "usage: xlsx-allocation-probe --source PATH --sheet NAME --phase open|edit|save|lifecycle --samples N --warmup N --output PATH"
}
