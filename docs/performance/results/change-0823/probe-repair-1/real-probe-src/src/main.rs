//! Isolated real-file PPTX edit profile for change 0822.
//!
//! The probe measures only the opened-presentation edit transaction. Each
//! iteration opens the source path before the clock, invokes the same small
//! edit helper in either direct or wrapped form, and stops the clock as soon
//! as that helper returns. Serialization, hashing, package destruction, and
//! complete semantic readback are deliberately outside the measured region.
//!
//! The source and reference identities are pinned to the admitted 0821
//! real-file corpus. A reference is required so the full semantic oracle is
//! derived from a separately opened package rather than from the code under
//! test or from a marker-only check.
//!
//! The optional `allocator-metrics` feature installs a packet-local global
//! allocator wrapper and publishes operation-region allocation counters. The
//! normal binary does not install that wrapper and omits allocation samples.

mod allocation_metrics;

#[cfg(feature = "allocator-metrics")]
mod counting_allocator;

use std::{
    env,
    error::Error,
    fs,
    hint::black_box,
    io::{self, Read, Write},
    path::{Path, PathBuf},
    time::Instant,
};

use litchi_pptx::Package;
use serde::Serialize;
use sha2::{Digest, Sha256};

const SCHEMA: &str = "litchi.performance.0823.pptx-edit-trial.v1";
const TOOL: &str = "pptx-edit-profile-0822";
const BASE_REVISION: &str = "b76786208d";
const MARKER: &str = "litchi-perf-0638-ordinary-save";
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 100;
const MAX_FILE_BYTES: u64 = 32 * 1024 * 1024;
const EXPECTED_INPUT_BYTES: u64 = 68_822;
const EXPECTED_INPUT_SHA256: &str =
    "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571";
const EXPECTED_OUTPUT_BYTES: u64 = 68_284;
const EXPECTED_OUTPUT_SHA256: &str =
    "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf";
const TIMING_SCOPE: &str = "opened_presentation_transaction set_shape_text(0, 0), commit, and apply_opened_presentation_commit; Package::open, Package owner drop, serialization, hashing, semantic readback, and allocator-region observer reads are outside the clock; the returned commit Snapshot is dropped inside the helper";

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq)]
#[serde(rename_all = "kebab-case")]
enum Mode {
    Direct,
    Wrapped,
}

impl Mode {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "direct" => Ok(Self::Direct),
            "wrapped" => Ok(Self::Wrapped),
            _ => Err(format!("--mode must be direct or wrapped (got {value:?})").into()),
        }
    }
}

#[derive(Debug)]
struct Config {
    input: PathBuf,
    reference: PathBuf,
    samples: usize,
    warmup: usize,
    mode: Mode,
    output: Option<PathBuf>,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct Identity {
    bytes: u64,
    sha256: String,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct FileIdentity {
    path: String,
    bytes: u64,
    sha256: String,
}

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq)]
struct Target {
    slide: usize,
    shape: usize,
}

#[derive(Clone, Debug)]
struct Oracle {
    reference: Identity,
    expected_output: Vec<u8>,
    full_text: String,
    full_text_sha256: String,
    slide_count: usize,
    target_text: String,
    marker_present: bool,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct Verification {
    all_verified: bool,
    input_hash_verified: bool,
    reference_hash_verified: bool,
    output_hash_verified: bool,
    output_size_verified: bool,
    output_bytes_verified: bool,
    reopened: bool,
    marker_verified: bool,
    target_verified: bool,
    full_text_digest_verified: bool,
    slide_count_verified: bool,
}

#[derive(Debug, Serialize)]
struct SampleRecord {
    index: usize,
    elapsed_ns: u64,
    output: Identity,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation: Option<allocation_metrics::Sample>,
    verification: Verification,
}

#[derive(Debug, Serialize)]
struct ElapsedSamples {
    unit: &'static str,
    samples: Vec<u64>,
    sample_order: Vec<usize>,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    base_revision: &'static str,
    mode: Mode,
    timing_scope: &'static str,
    input: FileIdentity,
    reference: FileIdentity,
    output: Identity,
    marker: &'static str,
    target: Target,
    target_text: String,
    full_text_sha256: String,
    full_text_digest: String,
    slide_count: usize,
    warmup: usize,
    samples_requested: usize,
    warmup_verified: bool,
    all_verified: bool,
    elapsed_ns: ElapsedSamples,
    samples: Vec<SampleRecord>,
    allocator: AllocatorIdentity,
}

#[derive(Debug, Serialize)]
struct AllocatorIdentity {
    binary: String,
    allocator: &'static str,
    instrumentation: &'static str,
    counter_revision: Option<&'static str>,
}

#[derive(Debug)]
struct RunResult {
    elapsed_ns: u64,
    output: Vec<u8>,
    output_identity: Identity,
    allocation: Option<allocation_metrics::Sample>,
    verification: Verification,
}

fn main() -> Result<(), Box<dyn Error>> {
    #[cfg(feature = "allocator-metrics")]
    allocation_metrics::enable();

    let config = parse_args(env::args_os().skip(1))?;
    run(config)
}

fn run(config: Config) -> Result<(), Box<dyn Error>> {
    validate_output_destination(&config)?;
    if config.input == config.reference {
        return Err("--input and --reference must name different files".into());
    }

    let (input_bytes, input) = read_pinned_file(
        &config.input,
        "input",
        EXPECTED_INPUT_BYTES,
        EXPECTED_INPUT_SHA256,
    )?;
    let (reference_bytes, reference) = read_pinned_file(
        &config.reference,
        "reference",
        EXPECTED_OUTPUT_BYTES,
        EXPECTED_OUTPUT_SHA256,
    )?;
    let oracle = build_oracle(&reference_bytes, &reference)?;
    let input_identity = Identity {
        bytes: input.bytes,
        sha256: input.sha256.clone(),
    };
    let reference_identity = oracle.reference.clone();
    black_box(input_bytes);

    let mut warmup_verified = true;
    for index in 0..config.warmup {
        let result = execute_one(
            &config.input,
            config.mode,
            &oracle,
            &input_identity,
            &reference_identity,
        )?;
        let verified = result.verification.all_verified;
        warmup_verified &= verified;
        black_box(result);
        if !verified {
            return Err(format!("warmup {index} failed its verification oracle").into());
        }
    }

    let mut records = Vec::with_capacity(config.samples);
    let mut elapsed = Vec::with_capacity(config.samples);
    for index in 0..config.samples {
        let result = execute_one(
            &config.input,
            config.mode,
            &oracle,
            &input_identity,
            &reference_identity,
        )?;
        if !result.verification.all_verified {
            return Err(format!("sample {index} failed its verification oracle").into());
        }
        elapsed.push(result.elapsed_ns);
        records.push(SampleRecord {
            index,
            elapsed_ns: result.elapsed_ns,
            output: result.output_identity,
            allocation: result.allocation,
            verification: result.verification,
        });
        black_box(result.output);
    }

    let all_verified = warmup_verified
        && records
            .iter()
            .all(|record| record.verification.all_verified);
    let output = Identity {
        bytes: EXPECTED_OUTPUT_BYTES,
        sha256: EXPECTED_OUTPUT_SHA256.to_owned(),
    };
    let report = Report {
        schema: SCHEMA,
        tool: TOOL,
        base_revision: BASE_REVISION,
        mode: config.mode,
        timing_scope: TIMING_SCOPE,
        input,
        reference,
        output,
        marker: MARKER,
        target: Target { slide: 0, shape: 0 },
        target_text: oracle.target_text,
        full_text_sha256: oracle.full_text_sha256.clone(),
        full_text_digest: oracle.full_text_sha256,
        slide_count: oracle.slide_count,
        warmup: config.warmup,
        samples_requested: config.samples,
        warmup_verified,
        all_verified,
        elapsed_ns: ElapsedSamples {
            unit: "ns",
            sample_order: (0..config.samples).collect(),
            samples: elapsed,
        },
        samples: records,
        allocator: AllocatorIdentity {
            binary: executable_identity(),
            allocator: allocation_metrics::allocator_identity(),
            instrumentation: allocation_metrics::instrumentation_identity(),
            counter_revision: allocation_metrics::counter_revision(),
        },
    };
    write_report(config.output.as_deref(), &report)?;
    Ok(())
}

/// The one operation body used by both profile modes. Keeping this helper
/// inline in the direct arm makes the direct and wrapped variants differ only
/// by the explicit non-inlined call boundary below.
#[inline(always)]
fn edit_helper_0822(package: &mut Package) -> Result<(), Box<dyn Error>> {
    let mut transaction = package.opened_presentation_transaction()?;
    if !transaction.set_shape_text(0, 0, MARKER)? {
        return Err("set_shape_text(0, 0) reported no change".into());
    }
    let commit = transaction.commit()?;
    if !commit.is_changed() {
        return Err("opened-presentation commit reported no change".into());
    }
    // The returned Snapshot is intentionally discarded at this semicolon,
    // inside the timed helper. Owner::edit in ordinary_save.rs has the same
    // ownership and drop point.
    package.apply_opened_presentation_commit(commit)?;
    Ok(())
}

/// Keep the wrapper as a real call boundary and consume the result inside it
/// so LLVM cannot turn the wrapper into a tail call to the edit helper.
#[inline(never)]
fn edit_region_0822(package: &mut Package) -> Result<(), Box<dyn Error>> {
    let result = edit_helper_0822(package);
    black_box(&result);
    result
}

fn execute_one(
    input_path: &Path,
    mode: Mode,
    oracle: &Oracle,
    input_identity: &Identity,
    reference: &Identity,
) -> Result<RunResult, Box<dyn Error>> {
    // File ingress is deliberately outside the clock. Each iteration opens
    // a fresh owner so no opened snapshot or package cache is shared between
    // samples.
    let mut package = Package::open(input_path)?;
    // The region boundary calls only take observer snapshots. They are kept
    // outside the elapsed clock so the feature reports allocation evidence
    // for the same edit body without changing the timed operation boundary.
    let region = allocation_metrics::begin();
    let started = Instant::now();
    let edited = match mode {
        Mode::Direct => edit_helper_0822(&mut package),
        Mode::Wrapped => edit_region_0822(&mut package),
    };
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    let allocation = region.finish();
    edited?;

    // Serialization, owner destruction, digesting, and semantic readback are
    // all outside the measured region.
    let output = package.to_bytes()?;
    drop(package);
    let output_identity = identity(&output)?;
    let verification = verify_output(&output, oracle, input_identity, reference)?;
    Ok(RunResult {
        elapsed_ns,
        output,
        output_identity,
        allocation,
        verification,
    })
}

fn build_oracle(
    reference_bytes: &[u8],
    reference: &FileIdentity,
) -> Result<Oracle, Box<dyn Error>> {
    if reference.bytes != EXPECTED_OUTPUT_BYTES || reference.sha256 != EXPECTED_OUTPUT_SHA256 {
        return Err("reference identity is not the admitted 0821 PPTX output".into());
    }
    let package = Package::from_bytes(reference_bytes)?;
    let (full_text, slide_count, target_text) = {
        let presentation = package.presentation()?;
        let slide = presentation
            .slide(0)?
            .ok_or("reference presentation has no slide at target index 0")?;
        let scene = slide.shapes()?;
        let shape = scene.shape(0)?;
        let target_text = shape
            .text()
            .ok_or("reference target shape has no text body")?
            .to_owned();
        (
            presentation.text()?,
            presentation.slide_count()?,
            target_text,
        )
    };
    let marker_present = target_text == MARKER && full_text.contains(MARKER);
    if !marker_present {
        return Err("reference target shape does not contain the admitted edit marker".into());
    }
    Ok(Oracle {
        reference: Identity {
            bytes: reference.bytes,
            sha256: reference.sha256.clone(),
        },
        expected_output: reference_bytes.to_vec(),
        full_text_sha256: sha256_hex(full_text.as_bytes()),
        full_text,
        slide_count,
        target_text,
        marker_present,
    })
}

fn verify_output(
    output: &[u8],
    oracle: &Oracle,
    input: &Identity,
    reference: &Identity,
) -> Result<Verification, Box<dyn Error>> {
    let output_identity = identity(output)?;
    let output_hash_verified = output_identity.sha256 == EXPECTED_OUTPUT_SHA256
        && output_identity.sha256 == oracle.reference.sha256;
    let output_size_verified = output_identity.bytes == EXPECTED_OUTPUT_BYTES
        && output_identity.bytes == oracle.reference.bytes;
    let output_bytes_verified = output == oracle.expected_output.as_slice();
    let reopened_package = Package::from_bytes(output)?;
    let (actual_text, actual_slide_count, actual_target_text) = {
        let presentation = reopened_package.presentation()?;
        let slide = presentation
            .slide(0)?
            .ok_or("reopened presentation has no slide at target index 0")?;
        let scene = slide.shapes()?;
        let shape = scene.shape(0)?;
        let target_text = shape
            .text()
            .ok_or("reopened target shape has no text body")?
            .to_owned();
        (
            presentation.text()?,
            presentation.slide_count()?,
            target_text,
        )
    };
    let actual_text_sha256 = sha256_hex(actual_text.as_bytes());
    let full_text_digest_verified =
        actual_text == oracle.full_text && actual_text_sha256 == oracle.full_text_sha256;
    let marker_verified = oracle.marker_present && actual_text.contains(MARKER);
    let target_verified = actual_target_text == oracle.target_text && actual_target_text == MARKER;
    let slide_count_verified = actual_slide_count == oracle.slide_count;
    let input_hash_verified =
        input.sha256 == EXPECTED_INPUT_SHA256 && input.bytes == EXPECTED_INPUT_BYTES;
    let reference_hash_verified =
        reference.sha256 == EXPECTED_OUTPUT_SHA256 && reference.bytes == EXPECTED_OUTPUT_BYTES;
    let reopened = true;
    let all_verified = input_hash_verified
        && reference_hash_verified
        && output_hash_verified
        && output_size_verified
        && output_bytes_verified
        && reopened
        && marker_verified
        && target_verified
        && full_text_digest_verified
        && slide_count_verified;
    Ok(Verification {
        all_verified,
        input_hash_verified,
        reference_hash_verified,
        output_hash_verified,
        output_size_verified,
        output_bytes_verified,
        reopened,
        marker_verified,
        target_verified,
        full_text_digest_verified,
        slide_count_verified,
    })
}

fn read_pinned_file(
    path: &Path,
    label: &str,
    expected_bytes: u64,
    expected_sha256: &str,
) -> Result<(Vec<u8>, FileIdentity), Box<dyn Error>> {
    let metadata = fs::metadata(path)
        .map_err(|error| format!("{label} cannot be stat'ed ({}): {error}", path.display()))?;
    if !metadata.is_file() {
        return Err(format!("{label} is not a regular file: {}", path.display()).into());
    }
    if metadata.len() > MAX_FILE_BYTES {
        return Err(format!(
            "{label} exceeds the 32 MiB input bound: {} bytes",
            metadata.len()
        )
        .into());
    }
    let mut reader = fs::File::open(path)
        .map_err(|error| format!("{label} cannot be read ({}): {error}", path.display()))?
        .take(MAX_FILE_BYTES.saturating_add(1));
    let mut bytes = Vec::new();
    reader
        .read_to_end(&mut bytes)
        .map_err(|error| format!("{label} cannot be read ({}): {error}", path.display()))?;
    if u64::try_from(bytes.len())? > MAX_FILE_BYTES {
        return Err(format!("{label} exceeds the 32 MiB input bound").into());
    }
    let identity = identity(&bytes)?;
    if identity.bytes != metadata.len() {
        return Err(format!("{label} changed while it was being read").into());
    }
    if identity.bytes != expected_bytes || identity.sha256 != expected_sha256 {
        return Err(format!(
            "{label} hash/size changed: got {} bytes sha256 {}, expected {} bytes sha256 {}",
            identity.bytes, identity.sha256, expected_bytes, expected_sha256
        )
        .into());
    }
    Ok((
        bytes,
        FileIdentity {
            path: path.display().to_string(),
            bytes: identity.bytes,
            sha256: identity.sha256,
        },
    ))
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

fn executable_identity() -> String {
    env::current_exe()
        .ok()
        .and_then(|path| {
            path.file_name()
                .map(|name| name.to_string_lossy().into_owned())
        })
        .unwrap_or_else(|| allocation_metrics::binary_identity().to_owned())
}

fn validate_output_destination(config: &Config) -> Result<(), Box<dyn Error>> {
    let Some(output) = config.output.as_deref() else {
        return Ok(());
    };
    if output == Path::new("-") {
        return Ok(());
    }
    if output == config.input || output == config.reference {
        return Err("--output must name a file distinct from --input and --reference".into());
    }
    if output.exists() {
        return Err(format!("output already exists: {}", output.display()).into());
    }
    if let Some(parent) = output.parent()
        && !parent.as_os_str().is_empty()
        && !parent.is_dir()
    {
        return Err(format!("output parent is not a directory: {}", parent.display()).into());
    }
    Ok(())
}

fn write_report(output: Option<&Path>, report: &Report) -> Result<(), Box<dyn Error>> {
    let encoded = serde_json::to_vec_pretty(report)?;
    match output {
        None => {
            let mut stdout = io::BufWriter::new(io::stdout().lock());
            stdout.write_all(&encoded)?;
            stdout.write_all(b"\n")?;
            stdout.flush()?;
        },
        Some(path) if path == Path::new("-") => {
            let mut stdout = io::BufWriter::new(io::stdout().lock());
            stdout.write_all(&encoded)?;
            stdout.write_all(b"\n")?;
            stdout.flush()?;
        },
        Some(path) => fs::write(path, encoded)?,
    }
    Ok(())
}

fn parse_args<I>(arguments: I) -> Result<Config, Box<dyn Error>>
where
    I: IntoIterator<Item = std::ffi::OsString>,
{
    let mut input = None;
    let mut reference = None;
    let mut samples = None;
    let mut warmup = 0;
    let mut mode = None;
    let mut output = None;
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        let argument = argument.to_string_lossy();
        match argument.as_ref() {
            "--input" => assign_once(&mut input, next_path(&mut arguments, "--input")?, "--input")?,
            "--reference" => assign_once(
                &mut reference,
                next_path(&mut arguments, "--reference")?,
                "--reference",
            )?,
            "--samples" => {
                let value = parse_count(
                    &next_string(&mut arguments, "--samples")?,
                    "samples",
                    1,
                    MAX_SAMPLES,
                )?;
                assign_once(&mut samples, value, "--samples")?;
            },
            "--warmup" => {
                warmup = parse_count(
                    &next_string(&mut arguments, "--warmup")?,
                    "warmup",
                    0,
                    MAX_WARMUP,
                )?;
            },
            "--mode" => {
                let value = Mode::parse(&next_string(&mut arguments, "--mode")?)?;
                assign_once(&mut mode, value, "--mode")?;
            },
            "--output" => assign_once(
                &mut output,
                next_path(&mut arguments, "--output")?,
                "--output",
            )?,
            "--help" | "-h" => return Err(usage().into()),
            value => return Err(format!("unknown argument {value:?}\n{}", usage()).into()),
        }
    }
    Ok(Config {
        input: input.ok_or_else(|| format!("--input is required\n{}", usage()))?,
        reference: reference.ok_or_else(|| format!("--reference is required\n{}", usage()))?,
        samples: samples.ok_or_else(|| format!("--samples is required\n{}", usage()))?,
        warmup,
        mode: mode.ok_or_else(|| format!("--mode is required\n{}", usage()))?,
        output,
    })
}

fn assign_once<T>(slot: &mut Option<T>, value: T, option: &str) -> Result<(), Box<dyn Error>> {
    if slot.is_some() {
        return Err(format!("{option} was provided more than once").into());
    }
    *slot = Some(value);
    Ok(())
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

fn parse_count(
    value: &str,
    name: &str,
    minimum: usize,
    maximum: usize,
) -> Result<usize, Box<dyn Error>> {
    let count = value
        .parse::<usize>()
        .map_err(|error| format!("--{name} must be a non-negative integer: {error}"))?;
    if count < minimum {
        return Err(format!("--{name} must be at least {minimum}").into());
    }
    if count > maximum {
        return Err(format!("--{name} exceeds maximum {maximum}").into());
    }
    Ok(count)
}

fn usage() -> &'static str {
    "usage: pptx-edit-profile-0822 --input PATH --reference PATH --samples N --warmup N --mode direct|wrapped [--output PATH|-]"
}

#[cfg(test)]
mod tests {
    use super::*;

    fn pinned_input_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../../../../test-data/ooxml/pptx/shapes.pptx")
    }

    fn pinned_reference_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../change-0821/artifacts/real-002-pptx/default.pptx")
    }

    fn args(values: &[&str]) -> Vec<std::ffi::OsString> {
        values.iter().map(std::ffi::OsString::from).collect()
    }

    #[test]
    fn cli_rejects_unknown_and_out_of_bounds_values() {
        let base = [
            "--input",
            "input.pptx",
            "--reference",
            "reference.pptx",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--mode",
            "direct",
        ];
        let mut unknown = base.to_vec();
        unknown.push("--unknown");
        assert!(parse_args(args(&unknown)).is_err());

        for value in ["0", "10001"] {
            let mut candidate = base.to_vec();
            candidate[5] = value;
            assert!(parse_args(args(&candidate)).is_err(), "samples={value}");
        }
        for value in ["101", "-1"] {
            let mut candidate = base.to_vec();
            candidate[7] = value;
            assert!(parse_args(args(&candidate)).is_err(), "warmup={value}");
        }
    }

    #[test]
    fn direct_and_wrapped_match_the_pinned_real_file_bytes_and_semantics() {
        let input_path = pinned_input_path();
        let reference_path = pinned_reference_path();
        let (input_bytes, input_file) = read_pinned_file(
            &input_path,
            "input",
            EXPECTED_INPUT_BYTES,
            EXPECTED_INPUT_SHA256,
        )
        .expect("pinned PPTX input");
        let (reference_bytes, reference_file) = read_pinned_file(
            &reference_path,
            "reference",
            EXPECTED_OUTPUT_BYTES,
            EXPECTED_OUTPUT_SHA256,
        )
        .expect("sealed 0821 PPTX reference");
        let oracle = build_oracle(&reference_bytes, &reference_file).expect("reference oracle");
        let input_identity = Identity {
            bytes: input_file.bytes,
            sha256: input_file.sha256,
        };
        let reference_identity = Identity {
            bytes: reference_file.bytes,
            sha256: reference_file.sha256,
        };
        let direct = execute_one(
            &input_path,
            Mode::Direct,
            &oracle,
            &input_identity,
            &reference_identity,
        )
        .expect("direct edit");
        let wrapped = execute_one(
            &input_path,
            Mode::Wrapped,
            &oracle,
            &input_identity,
            &reference_identity,
        )
        .expect("wrapped edit");
        assert_eq!(direct.output, wrapped.output);
        assert_eq!(direct.output, reference_bytes);
        assert_eq!(direct.output_identity.sha256, EXPECTED_OUTPUT_SHA256);
        assert_eq!(direct.output_identity.bytes, EXPECTED_OUTPUT_BYTES);
        assert!(direct.verification.all_verified);
        assert!(wrapped.verification.all_verified);
        assert_eq!(oracle.slide_count, 6);
        assert_eq!(input_bytes.len() as u64, EXPECTED_INPUT_BYTES);
    }

    #[test]
    fn wrong_source_or_reference_identity_is_rejected() {
        let input_path = pinned_input_path();
        let reference_path = pinned_reference_path();
        assert!(
            read_pinned_file(
                &reference_path,
                "input",
                EXPECTED_INPUT_BYTES,
                EXPECTED_INPUT_SHA256,
            )
            .is_err()
        );
        assert!(
            read_pinned_file(
                &input_path,
                "reference",
                EXPECTED_OUTPUT_BYTES,
                EXPECTED_OUTPUT_SHA256,
            )
            .is_err()
        );
    }
}
