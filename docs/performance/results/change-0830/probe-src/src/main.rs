//! Isolated real-file XLSX edit profile for change 0830.
//!
//! The probe measures only the public semantic edit transaction. Each sample
//! opens the source workbook before the clock, edits `Munka1!A1`, commits the
//! transaction, and adopts the committed workbook before stopping the clock.
//! Serialization, owner destruction, hashing, and the complete stored-cell
//! readback are outside the measured interval.
//!
//! The direct and wrapped modes execute the same public operation. Wrapped
//! mode adds one no-inline call boundary for attribution; it does not add a
//! second edit or any pre-timing cache warmup.

use std::{
    env,
    error::Error,
    fmt::Write as FmtWrite,
    fs,
    hint::black_box,
    io::{self, Read, Write},
    path::{Path, PathBuf},
    time::Instant,
};

use litchi_xlsx::{Cell, Value, Workbook, WorksheetKind};
use serde::Serialize;
use sha2::{Digest, Sha256};

const SCHEMA: &str = "litchi.performance.0830.xlsx-edit-profile.v1";
const TOOL: &str = "xlsx-edit-profile-0830";
const BASE_REVISION: &str = "87eb57182be1622385dc3b28dfc3c7be868dca32";
const MARKER: &str = "litchi-perf-0638-ordinary-save";
const TARGET_SHEET: &str = "Munka1";
const TARGET_ADDRESS: &str = "A1";
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 100;
const MAX_FILE_BYTES: u64 = 32 * 1024 * 1024;
const EXPECTED_INPUT_BYTES: u64 = 8_435;
const EXPECTED_INPUT_SHA256: &str =
    "d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4";
const EXPECTED_OUTPUT_BYTES: u64 = 8_521;
const EXPECTED_OUTPUT_SHA256: &str =
    "0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68";
const EXPECTED_WORKSHEET_COUNT: usize = 1;
const EXPECTED_STORED_CELL_COUNT: usize = 8;
const TIMING_SCOPE: &str = "Workbook::open is outside the clock; Workbook::edit, sheet(\"Munka1\"), set(\"A1\", marker), commit, patch non-empty check, commit.into_workbook adoption, and the old Workbook drop at assignment are inside the clock; serialization, owner drop, hashing, and complete stored-cell semantic readback are outside the clock";

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
    sheet: &'static str,
    address: &'static str,
}

#[derive(Clone, Debug)]
struct SemanticSnapshot {
    projection: String,
    sha256: String,
    worksheet_count: usize,
    stored_cell_count: usize,
    target_value: String,
}

#[derive(Clone, Debug)]
struct Oracle {
    reference: Identity,
    expected_output: Vec<u8>,
    semantic_projection: String,
    semantic_sha256: String,
    worksheet_count: usize,
    stored_cell_count: usize,
    target_value: String,
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
    semantic_sha256_verified: bool,
    worksheet_count_verified: bool,
    stored_cell_count_verified: bool,
}

#[derive(Debug, Serialize)]
struct SampleRecord {
    index: usize,
    elapsed_ns: u64,
    output: Identity,
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
    target_value: String,
    worksheet_count: usize,
    stored_cell_count: usize,
    semantic_sha256: String,
    warmup: usize,
    samples_requested: usize,
    warmup_verified: bool,
    all_verified: bool,
    elapsed_ns: ElapsedSamples,
    samples: Vec<SampleRecord>,
}

#[derive(Debug)]
struct RunResult {
    elapsed_ns: u64,
    output: Vec<u8>,
    output_identity: Identity,
    verification: Verification,
}

fn main() -> Result<(), Box<dyn Error>> {
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
        target: Target {
            sheet: TARGET_SHEET,
            address: TARGET_ADDRESS,
        },
        target_value: oracle.target_value.clone(),
        worksheet_count: oracle.worksheet_count,
        stored_cell_count: oracle.stored_cell_count,
        semantic_sha256: oracle.semantic_sha256,
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
    };
    write_report(config.output.as_deref(), &report)?;
    Ok(())
}

/// The unwrapped public operation used by direct mode. The assignment adopts
/// the committed snapshot and drops the old workbook at the same point as
/// the ordinary-save `Owner::edit` path.
#[inline(always)]
fn edit_helper_0830(workbook: &mut Workbook) -> Result<(), Box<dyn Error>> {
    let mut edit = workbook.edit()?;
    {
        let mut sheet = edit
            .sheet(TARGET_SHEET)?
            .ok_or_else(|| format!("worksheet {TARGET_SHEET:?} is absent"))?;
        sheet.set(TARGET_ADDRESS, MARKER)?;
    }
    let commit = edit.commit()?;
    if commit.patch().is_empty() {
        return Err("the edit produced an empty patch".into());
    }
    *workbook = commit.into_workbook();
    Ok(())
}

/// Keep one real no-inline call boundary for wrapped-mode attribution. The
/// result is consumed inside this wrapper so LLVM cannot turn it into a tail
/// call or discard the operation.
#[inline(never)]
fn edit_region_0830(workbook: &mut Workbook) -> Result<(), Box<dyn Error>> {
    let result = edit_helper_0830(workbook);
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
    // Ingress is deliberately outside the clock. Every iteration owns a
    // fresh workbook, so no parsed snapshot or package cache is shared.
    let mut workbook = Workbook::open(input_path)?;
    let started = Instant::now();
    let edited = match mode {
        Mode::Direct => edit_helper_0830(&mut workbook),
        Mode::Wrapped => edit_region_0830(&mut workbook),
    };
    let elapsed_ns = elapsed_ns(started.elapsed())?;
    edited?;

    // Serialization, workbook destruction, identity calculation, and full
    // semantic readback are outside the measured region.
    let output = workbook.to_bytes()?;
    drop(workbook);
    let output_identity = identity(&output)?;
    let verification = verify_output(&output, oracle, input_identity, reference)?;
    Ok(RunResult {
        elapsed_ns,
        output,
        output_identity,
        verification,
    })
}

fn build_oracle(
    reference_bytes: &[u8],
    reference: &FileIdentity,
) -> Result<Oracle, Box<dyn Error>> {
    if reference.bytes != EXPECTED_OUTPUT_BYTES || reference.sha256 != EXPECTED_OUTPUT_SHA256 {
        return Err("reference identity is not the admitted 0821 XLSX output".into());
    }
    let workbook = Workbook::from_bytes(reference_bytes.to_vec())?;
    let semantic = semantic_snapshot(&workbook)?;
    if semantic.worksheet_count != EXPECTED_WORKSHEET_COUNT
        || semantic.stored_cell_count != EXPECTED_STORED_CELL_COUNT
    {
        return Err(format!(
            "reference workbook has {} worksheets and {} stored cells, expected {EXPECTED_WORKSHEET_COUNT} and {EXPECTED_STORED_CELL_COUNT}",
            semantic.worksheet_count, semantic.stored_cell_count
        )
        .into());
    }
    if semantic.target_value != MARKER {
        return Err(format!(
            "reference target {TARGET_SHEET}!{TARGET_ADDRESS} does not contain the edit marker"
        )
        .into());
    }
    Ok(Oracle {
        reference: Identity {
            bytes: reference.bytes,
            sha256: reference.sha256.clone(),
        },
        expected_output: reference_bytes.to_vec(),
        semantic_projection: semantic.projection,
        semantic_sha256: semantic.sha256,
        worksheet_count: semantic.worksheet_count,
        stored_cell_count: semantic.stored_cell_count,
        target_value: semantic.target_value,
    })
}

fn semantic_snapshot(workbook: &Workbook) -> Result<SemanticSnapshot, Box<dyn Error>> {
    let mut projection = String::new();
    let mut worksheet_count = 0;
    let mut stored_cell_count = 0;
    for sheet in workbook.sheets() {
        if sheet.kind() != WorksheetKind::Worksheet {
            continue;
        }
        worksheet_count += 1;
        projection.push_str("sheet=");
        projection.push_str(&format!("{:?}\n", sheet.name()));
        if let Some(extent) = sheet.stored_extent()? {
            for (address, cell) in sheet.cells(extent)? {
                stored_cell_count += 1;
                let _ = writeln!(&mut projection, "cell={:?}\tstate={cell:?}", address.a1());
            }
        }
    }
    let target_value = target_value(workbook)?;
    let sha256 = sha256_hex(projection.as_bytes());
    Ok(SemanticSnapshot {
        projection,
        sha256,
        worksheet_count,
        stored_cell_count,
        target_value,
    })
}

fn target_value(workbook: &Workbook) -> Result<String, Box<dyn Error>> {
    let sheet = workbook
        .sheet(TARGET_SHEET)?
        .ok_or_else(|| format!("worksheet {TARGET_SHEET:?} is absent"))?;
    let value = match sheet.cell(TARGET_ADDRESS)? {
        litchi_xlsx::cell::View::Stored(Cell::Value(Value::Text(value))) => {
            value.as_str().to_owned()
        },
        litchi_xlsx::cell::View::Stored(cell) => format!("{cell:?}"),
        view => format!("{view:?}"),
    };
    Ok(value)
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
    let reopened_workbook = Workbook::from_bytes(output.to_vec())?;
    let semantic = semantic_snapshot(&reopened_workbook)?;
    let semantic_sha256_verified = semantic.sha256 == oracle.semantic_sha256
        && semantic.projection == oracle.semantic_projection;
    let marker_verified = semantic.target_value == MARKER;
    let target_verified =
        semantic.target_value == oracle.target_value && semantic.target_value == MARKER;
    let worksheet_count_verified = semantic.worksheet_count == oracle.worksheet_count;
    let stored_cell_count_verified = semantic.stored_cell_count == oracle.stored_cell_count;
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
        && semantic_sha256_verified
        && worksheet_count_verified
        && stored_cell_count_verified;
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
        semantic_sha256_verified,
        worksheet_count_verified,
        stored_cell_count_verified,
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
        Some(path) => {
            let mut file = fs::OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(path)
                .map_err(|error| {
                    format!("cannot create report output {}: {error}", path.display())
                })?;
            file.write_all(&encoded)?;
            file.flush()?;
        },
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
    let mut warmup = None;
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
                let value = parse_count(
                    &next_string(&mut arguments, "--warmup")?,
                    "warmup",
                    0,
                    MAX_WARMUP,
                )?;
                assign_once(&mut warmup, value, "--warmup")?;
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
        warmup: warmup.unwrap_or(0),
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
    "usage: xlsx-edit-profile-0830 --input PATH --reference PATH --samples N --warmup N --mode direct|wrapped [--output PATH|-]"
}

#[cfg(test)]
mod tests {
    use super::*;

    fn pinned_input_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(
            "../../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx",
        )
    }

    fn pinned_reference_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../change-0821/artifacts/real-001-xlsx/default.xlsx")
    }

    fn args(values: &[&str]) -> Vec<std::ffi::OsString> {
        values.iter().map(std::ffi::OsString::from).collect()
    }

    #[test]
    fn cli_rejects_unknown_and_out_of_bounds_values() {
        let base = [
            "--input",
            "input.xlsx",
            "--reference",
            "reference.xlsx",
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
        .expect("pinned XLSX input");
        let (reference_bytes, reference_file) = read_pinned_file(
            &reference_path,
            "reference",
            EXPECTED_OUTPUT_BYTES,
            EXPECTED_OUTPUT_SHA256,
        )
        .expect("sealed 0821 XLSX reference");
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
        assert_eq!(oracle.worksheet_count, EXPECTED_WORKSHEET_COUNT);
        assert_eq!(oracle.stored_cell_count, EXPECTED_STORED_CELL_COUNT);
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
