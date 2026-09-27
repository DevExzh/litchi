//! Standalone public-workflow probe for change 0784.
//!
//! The corpus and the public PPTX calls are copied from the deterministic
//! 0760 probe.  This executable is deliberately independent of the workspace
//! harness: the build driver materializes its manifest, and the optional
//! allocator target installs instrumentation only in this process.
//!
//! The timed regions are deliberately small and explicit:
//!
//! * `capture` measures `Package::opened_presentation`.
//! * `commit` stages one edit before the clock and measures `commit`.
//! * `lifecycle` measures capture, edit, commit, publication, and
//!   serialization as one public operation.
//! * `capabilities` is a constructor diagnostic.  It repeats the fixed
//!   `Capabilities::ooxml_baseline` profile, but it is not an end-to-end PPTX
//!   claim.
//!
//! Package serialization, reopening, marker checks, and destruction happen
//! after the timed region unless they are explicitly part of `lifecycle`.
//! Owners remain live until the allocation region has finished.
//!
//! The optional `phase-timing` feature records five consecutive `Instant`
//! boundaries inside the lifecycle clock.  Those boundaries are diagnostic
//! measurements and add clock-reading overhead; they do not establish a
//! causal historical regression claim.

mod allocation_metrics;

#[cfg(feature = "allocator-metrics")]
mod counting_allocator;

use std::collections::BTreeMap;
use std::env;
use std::error::Error;
use std::fs;
use std::hint::black_box;
use std::path::PathBuf;
use std::time::Instant;

use litchi_ooxml_common::mce::Capabilities;
use litchi_pptx::Package;
use serde::Serialize;
use sha2::{Digest, Sha256};

const SCHEMA: &str = "litchi.pptx.capture-profile-probe.v1";
const TOOL: &str = "pptx-capture-probe-0784";
const MARKER: &str = "litchi-perf-0780-static-mce-capabilities";
const DEFAULT_SAMPLES: usize = 1;
const DEFAULT_WARMUP: usize = 0;
const MAX_SAMPLES: usize = 100_000;
const MAX_WARMUP: usize = 100_000;
const CAPABILITY_REPETITIONS: usize = 1_000;

const BASELINE_NAMESPACES: &[&str] = &[
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://purl.oclc.org/ooxml/presentationml/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/main",
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://purl.oclc.org/ooxml/drawingml/chart",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://purl.oclc.org/ooxml/officeDocument/math",
    "urn:schemas-microsoft-com:vml",
    "urn:schemas-microsoft-com:office:office",
    "http://www.w3.org/XML/1998/namespace",
];

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Mode {
    Capture,
    Commit,
    Lifecycle,
    Capabilities,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "capture" => Ok(Self::Capture),
            "commit" => Ok(Self::Commit),
            "lifecycle" => Ok(Self::Lifecycle),
            "capabilities" => Ok(Self::Capabilities),
            _ => Err(format!(
                "--mode must be one of capture, commit, lifecycle, capabilities (got {value:?})"
            )
            .into()),
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::Capture => "Package::opened_presentation only",
            Self::Commit => {
                "Transaction::commit only; package capture and one set_shape_text staging are outside the clock"
            },
            Self::Lifecycle => {
                "Package::opened_presentation, edit, set_shape_text, commit, apply_opened_presentation_commit, and Package::to_bytes; phase-timing is a diagnostic clock-boundary mode with clock overhead and no causal historical-regression claim"
            },
            Self::Capabilities => {
                "1000 Capabilities::ooxml_baseline constructions and membership checks; constructor diagnostic only"
            },
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "kebab-case")]
enum Shape {
    Tiny,
    Medium,
    Large,
}

impl Shape {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "tiny" => Ok(Self::Tiny),
            "medium" => Ok(Self::Medium),
            "large" => Ok(Self::Large),
            _ => Err(format!("--shape must be one of tiny, medium, large (got {value:?})").into()),
        }
    }

    const fn dimensions(self) -> (usize, usize) {
        match self {
            Self::Tiny => (3, 4),
            Self::Medium => (12, 8),
            Self::Large => (100, 100),
        }
    }
}

#[derive(Debug)]
struct Config {
    mode: Mode,
    shape: Shape,
    samples: usize,
    warmup: usize,
    output: PathBuf,
}

#[derive(Debug, Serialize, Clone)]
struct Identity {
    bytes: u64,
    sha256: String,
}

#[derive(Debug, Serialize)]
struct AllocatorIdentity {
    binary: String,
    allocator: &'static str,
    instrumentation: &'static str,
    counter_revision: Option<&'static str>,
}

#[derive(Debug, Serialize)]
struct Verification {
    semantic_check: bool,
    reopened: bool,
    expected_text: Option<String>,
    actual_text: Option<String>,
    semantic_text_bytes: Option<u64>,
    semantic_text_sha256: Option<String>,
    readback_bytes: Option<u64>,
    readback_sha256: Option<String>,
    marker_matches: Option<bool>,
}

#[derive(Debug, Serialize)]
struct SampleRecord {
    index: usize,
    elapsed_ns: u64,
    metrics: BTreeMap<String, u64>,
    source_sha256: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    output: Option<Identity>,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation: Option<allocation_metrics::Sample>,
    verification: Verification,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    mode: Mode,
    shape: Shape,
    slides: usize,
    shapes_per_slide: usize,
    timing_scope: &'static str,
    marker: &'static str,
    source: Identity,
    warmup: usize,
    samples_requested: usize,
    samples: Vec<SampleRecord>,
    allocator: AllocatorIdentity,
}

#[derive(Debug)]
struct OperationResult {
    elapsed_ns: u64,
    output: Option<Vec<u8>>,
    verification: Verification,
    allocation: Option<allocation_metrics::Sample>,
    extra_metrics: BTreeMap<String, u64>,
}

fn main() -> Result<(), Box<dyn Error>> {
    #[cfg(feature = "allocator-metrics")]
    allocation_metrics::enable();

    let config = parse_args(env::args_os().skip(1))?;
    run(config)
}

fn run(config: Config) -> Result<(), Box<dyn Error>> {
    if let Some(parent) = config.output.parent()
        && !parent.as_os_str().is_empty()
        && !parent.is_dir()
    {
        return Err(format!("output parent is not a directory: {}", parent.display()).into());
    }
    if config.output.exists() {
        return Err(format!("output already exists: {}", config.output.display()).into());
    }

    let (slides, shapes_per_slide) = config.shape.dimensions();
    let source_bytes = build(slides, shapes_per_slide)?;
    let source = identity(&source_bytes)?;
    let source_sha256 = source.sha256.clone();

    let mut samples = Vec::with_capacity(config.samples);
    let mut deterministic_output = None;
    for _ in 0..config.warmup {
        let result = run_one(&config, &source_bytes, slides, shapes_per_slide, 0)?;
        observe_output(&mut deterministic_output, result.output.as_deref())?;
        black_box(result);
    }
    for index in 0..config.samples {
        let result = run_one(&config, &source_bytes, slides, shapes_per_slide, index)?;
        observe_output(&mut deterministic_output, result.output.as_deref())?;
        let output = result.output.as_deref().map(identity).transpose()?;
        let mut metrics = result.extra_metrics;
        metrics.insert("elapsed_ns".to_owned(), result.elapsed_ns);
        metrics.insert("slides".to_owned(), slides as u64);
        metrics.insert("shapes_per_slide".to_owned(), shapes_per_slide as u64);
        samples.push(SampleRecord {
            index,
            elapsed_ns: result.elapsed_ns,
            metrics,
            source_sha256: source_sha256.clone(),
            output,
            allocation: result.allocation,
            verification: result.verification,
        });
    }

    let report = Report {
        schema: SCHEMA,
        tool: TOOL,
        mode: config.mode,
        shape: config.shape,
        slides,
        shapes_per_slide,
        timing_scope: config.mode.timing_scope(),
        marker: MARKER,
        source,
        warmup: config.warmup,
        samples_requested: config.samples,
        samples,
        allocator: AllocatorIdentity {
            binary: executable_identity(),
            allocator: allocation_metrics::allocator_identity(),
            instrumentation: allocation_metrics::instrumentation_identity(),
            counter_revision: allocation_metrics::counter_revision(),
        },
    };
    fs::write(&config.output, serde_json::to_vec_pretty(&report)?)?;
    Ok(())
}

fn run_one(
    config: &Config,
    source_bytes: &[u8],
    slides: usize,
    shapes_per_slide: usize,
    index: usize,
) -> Result<OperationResult, Box<dyn Error>> {
    match config.mode {
        Mode::Capture => run_capture(source_bytes, slides, shapes_per_slide),
        Mode::Commit => run_commit(source_bytes, slides, shapes_per_slide),
        Mode::Lifecycle => run_lifecycle(source_bytes, slides, shapes_per_slide),
        Mode::Capabilities => run_capabilities(source_bytes, index),
    }
}

// Exactly one call per capture operation; no setup, verification or owner drops.
// The non-inlined boundary is a profiler diagnostic, not a production change.
#[cfg(feature = "capture-profile")]
#[inline(never)]
fn capture_region_0784(package: &Package) -> litchi_pptx::Result<litchi_pptx::opened::Snapshot> {
    package.opened_presentation()
}

fn run_capture(
    source_bytes: &[u8],
    slides: usize,
    shapes_per_slide: usize,
) -> Result<OperationResult, Box<dyn Error>> {
    let mut package = Package::from_bytes(source_bytes)?;
    let region = allocation_metrics::begin();
    let started = Instant::now();
    #[cfg(feature = "capture-profile")]
    let snapshot = capture_region_0784(&package)?;
    #[cfg(not(feature = "capture-profile"))]
    let snapshot = package.opened_presentation()?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    let allocation = region.finish();
    black_box(&snapshot);

    if snapshot.slides().len() != slides {
        return Err(format!(
            "capture readback found {} slides, expected {slides}",
            snapshot.slides().len()
        )
        .into());
    }
    let output = package.to_bytes()?;
    let verification = verify_output(&output, slides, shapes_per_slide, false)?;
    let mut extra_metrics = BTreeMap::new();
    extra_metrics.insert("captured_slides".to_owned(), snapshot.slides().len() as u64);
    extra_metrics.insert(
        "captured_shapes_per_slide".to_owned(),
        shapes_per_slide as u64,
    );
    Ok(OperationResult {
        elapsed_ns,
        output: Some(output),
        verification,
        allocation,
        extra_metrics,
    })
}

fn run_commit(
    source_bytes: &[u8],
    slides: usize,
    shapes_per_slide: usize,
) -> Result<OperationResult, Box<dyn Error>> {
    let mut package = Package::from_bytes(source_bytes)?;
    let snapshot = package.opened_presentation()?;
    let mut edit = snapshot.edit();
    if !edit.set_shape_text(0, 0, MARKER)? {
        return Err("commit setup did not stage a changed marker".into());
    }

    let region = allocation_metrics::begin();
    let started = Instant::now();
    let commit = edit.commit()?;
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    let allocation = region.finish();
    black_box(&commit);
    if !commit.is_changed() {
        return Err("commit produced an empty patch".into());
    }

    package.apply_opened_presentation_commit(commit)?;
    let output = package.to_bytes()?;
    let verification = verify_output(&output, slides, shapes_per_slide, true)?;
    Ok(OperationResult {
        elapsed_ns,
        output: Some(output),
        verification,
        allocation,
        extra_metrics: BTreeMap::new(),
    })
}

fn run_lifecycle(
    source_bytes: &[u8],
    slides: usize,
    shapes_per_slide: usize,
) -> Result<OperationResult, Box<dyn Error>> {
    // Package ownership and source ingress are intentionally outside this
    // clock.  The public capture/edit/publication route is the measured unit.
    let mut package = Package::from_bytes(source_bytes)?;
    let region = allocation_metrics::begin();
    let started = Instant::now();
    let snapshot = package.opened_presentation()?;
    #[cfg(feature = "phase-timing")]
    let captured_at = Instant::now();
    let mut edit = snapshot.edit();
    if !edit.set_shape_text(0, 0, MARKER)? {
        return Err("lifecycle setup did not stage a changed marker".into());
    }
    #[cfg(feature = "phase-timing")]
    let staged_at = Instant::now();
    let commit = edit.commit()?;
    #[cfg(feature = "phase-timing")]
    let committed_at = Instant::now();
    package.apply_opened_presentation_commit(commit)?;
    #[cfg(feature = "phase-timing")]
    let applied_at = Instant::now();
    let output = package.to_bytes()?;
    #[cfg(feature = "phase-timing")]
    let serialized_at = Instant::now();
    let elapsed_ns = {
        #[cfg(feature = "phase-timing")]
        {
            elapsed_ns(serialized_at.duration_since(started))?
        }
        #[cfg(not(feature = "phase-timing"))]
        {
            elapsed_ns(started.elapsed())?
        }
    };
    let allocation = region.finish();
    black_box(&package);

    let verification = verify_output(&output, slides, shapes_per_slide, true)?;
    #[cfg(feature = "phase-timing")]
    let extra_metrics = {
        let mut metrics = BTreeMap::new();
        let phase_capture_ns = crate::elapsed_ns(captured_at.duration_since(started))?;
        let phase_stage_ns = crate::elapsed_ns(staged_at.duration_since(captured_at))?;
        let phase_commit_ns = crate::elapsed_ns(committed_at.duration_since(staged_at))?;
        let phase_apply_ns = crate::elapsed_ns(applied_at.duration_since(committed_at))?;
        let phase_serialize_ns = crate::elapsed_ns(serialized_at.duration_since(applied_at))?;
        let phase_sum = phase_capture_ns
            .checked_add(phase_stage_ns)
            .and_then(|sum| sum.checked_add(phase_commit_ns))
            .and_then(|sum| sum.checked_add(phase_apply_ns))
            .and_then(|sum| sum.checked_add(phase_serialize_ns))
            .ok_or("lifecycle phase durations overflow u64")?;
        if phase_sum != elapsed_ns {
            return Err(format!(
                "lifecycle phase durations sum to {phase_sum} ns, elapsed clock reports {elapsed_ns} ns"
            )
            .into());
        }
        metrics.insert("phase_capture_ns".to_owned(), phase_capture_ns);
        metrics.insert("phase_stage_ns".to_owned(), phase_stage_ns);
        metrics.insert("phase_commit_ns".to_owned(), phase_commit_ns);
        metrics.insert("phase_apply_ns".to_owned(), phase_apply_ns);
        metrics.insert("phase_serialize_ns".to_owned(), phase_serialize_ns);
        metrics
    };
    #[cfg(not(feature = "phase-timing"))]
    let extra_metrics = BTreeMap::new();
    Ok(OperationResult {
        elapsed_ns,
        output: Some(output),
        verification,
        allocation,
        extra_metrics,
    })
}

fn run_capabilities(source_bytes: &[u8], _index: usize) -> Result<OperationResult, Box<dyn Error>> {
    let region = allocation_metrics::begin();
    let started = Instant::now();
    let mut recognized = 0usize;
    for _ in 0..CAPABILITY_REPETITIONS {
        let capabilities = Capabilities::ooxml_baseline();
        recognized += BASELINE_NAMESPACES
            .iter()
            .filter(|namespace| capabilities.understands(black_box(*namespace)))
            .count();
        black_box(&capabilities);
    }
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    let allocation = region.finish();
    let expected = CAPABILITY_REPETITIONS * BASELINE_NAMESPACES.len();
    if recognized != expected {
        return Err(format!(
            "capability diagnostic recognized {recognized} names, expected {expected}"
        )
        .into());
    }
    let mut extra_metrics = BTreeMap::new();
    extra_metrics.insert(
        "constructor_repetitions".to_owned(),
        CAPABILITY_REPETITIONS as u64,
    );
    extra_metrics.insert("recognized_memberships".to_owned(), recognized as u64);
    Ok(OperationResult {
        elapsed_ns,
        output: None,
        verification: Verification {
            semantic_check: true,
            reopened: false,
            expected_text: None,
            actual_text: None,
            semantic_text_bytes: None,
            semantic_text_sha256: None,
            readback_bytes: None,
            readback_sha256: None,
            marker_matches: None,
        },
        allocation,
        extra_metrics,
    })
}

fn verify_output(
    output: &[u8],
    slides: usize,
    shapes_per_slide: usize,
    marker_operation: bool,
) -> Result<Verification, Box<dyn Error>> {
    let reopened = Package::from_bytes(output)?;
    let actual_text = reopened.presentation()?.text()?;
    let expected_full_text = expected_text(slides, shapes_per_slide, marker_operation);
    let semantic_check = actual_text == expected_full_text;
    let first_expected = if marker_operation {
        MARKER.to_owned()
    } else {
        source_text(0, 0)
    };
    if !semantic_check {
        return Err(format!(
            "PPTX readback differed from the exact generated text; expected {} bytes, got {} bytes",
            expected_full_text.len(),
            actual_text.len()
        )
        .into());
    }
    Ok(Verification {
        semantic_check,
        reopened: true,
        expected_text: Some(first_expected.clone()),
        actual_text: Some(first_expected),
        semantic_text_bytes: Some(u64::try_from(actual_text.len())?),
        semantic_text_sha256: Some(sha256_hex(actual_text.as_bytes())),
        // The readback identity binds the semantic check to the exact
        // reopened package that the sample published.  The selected text is
        // retained separately in `actual_text`; the full semantic text is
        // deliberately represented by a digest rather than copied into every
        // large-sample report.
        readback_bytes: Some(u64::try_from(output.len())?),
        readback_sha256: Some(sha256_hex(output)),
        marker_matches: marker_operation.then_some(true),
    })
}

fn expected_text(slides: usize, shapes_per_slide: usize, marker_operation: bool) -> String {
    let mut values = Vec::with_capacity(slides.saturating_mul(shapes_per_slide));
    for slide in 0..slides {
        for shape in 0..shapes_per_slide {
            values.push(if marker_operation && slide == 0 && shape == 0 {
                MARKER.to_owned()
            } else {
                source_text(slide, shape)
            });
        }
    }
    values.join("\n")
}

fn executable_identity() -> String {
    env::current_exe()
        .ok()
        .and_then(|path| {
            path.file_name()
                .map(|name| name.to_string_lossy().into_owned())
        })
        .unwrap_or_else(|| allocation_metrics::binary_identity().to_owned())
}

fn observe_output(
    expected: &mut Option<Identity>,
    output: Option<&[u8]>,
) -> Result<(), Box<dyn Error>> {
    let Some(output) = output else {
        return Ok(());
    };
    let current = identity(output)?;
    match expected {
        None => *expected = Some(current),
        Some(previous) if previous.bytes == current.bytes && previous.sha256 == current.sha256 => {
        },
        Some(previous) => {
            return Err(format!(
                "non-deterministic publication: {} bytes / {} followed {} bytes / {}",
                previous.bytes, previous.sha256, current.bytes, current.sha256
            )
            .into());
        },
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Deterministic corpus copied from change 0760.
// ---------------------------------------------------------------------------

fn source_text(slide: usize, shape: usize) -> String {
    format!("litchi-perf-baseline-pptx-semantic-v1-source-{slide:03}-{shape:03}")
}

fn build(slides: usize, boxes: usize) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for slide_index in 0..slides {
            let slide = presentation.add_slide()?;
            for shape_index in 0..boxes {
                slide.add_text_box(
                    &source_text(slide_index, shape_index),
                    36 + i64::try_from(shape_index % 4)? * 180,
                    36 + i64::try_from(shape_index / 4)? * 90,
                    144,
                    54,
                );
            }
        }
    }
    Ok(package.to_bytes()?)
}

fn identity(bytes: &[u8]) -> Result<Identity, Box<dyn Error>> {
    Ok(Identity {
        bytes: u64::try_from(bytes.len())?,
        sha256: sha256_hex(bytes),
    })
}

fn sha256_hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn elapsed_ns(duration: std::time::Duration) -> Result<u64, Box<dyn Error>> {
    u64::try_from(duration.as_nanos())
        .map_err(|_| "elapsed duration overflows u64 nanoseconds".into())
}

fn parse_args<I>(arguments: I) -> Result<Config, Box<dyn Error>>
where
    I: IntoIterator<Item = std::ffi::OsString>,
{
    let mut mode = None;
    let mut shape = None;
    let mut samples = DEFAULT_SAMPLES;
    let mut warmup = DEFAULT_WARMUP;
    let mut output = None;
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        let argument = argument.to_string_lossy();
        match argument.as_ref() {
            "--mode" => mode = Some(Mode::parse(&next_string(&mut arguments, "--mode")?)?),
            "--shape" => shape = Some(Shape::parse(&next_string(&mut arguments, "--shape")?)?),
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
    let mode = mode.ok_or_else(|| format!("--mode is required\n{}", usage()))?;
    let shape = shape.ok_or_else(|| format!("--shape is required\n{}", usage()))?;
    let output = output.ok_or_else(|| format!("--output is required\n{}", usage()))?;
    if samples == 0 {
        return Err("--samples must be at least 1".into());
    }
    Ok(Config {
        mode,
        shape,
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
    "usage: pptx-capture-probe --mode capture|commit|lifecycle|capabilities --shape tiny|medium|large --samples N --warmup N --output PATH"
}
