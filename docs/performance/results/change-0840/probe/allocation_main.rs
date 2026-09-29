//! Operation-scoped allocation companion for the fresh CFB/DOC emission probe.
//!
//! Corpus generation and output verification are outside the allocation
//! region. The region contains the same public writer construction,
//! prepared-input registration, and write_to operation as the native binary.

#[allow(dead_code)]
mod allocation_metrics;
mod counting_allocator;

#[allow(dead_code)]
#[path = "main.rs"]
mod probe;

use serde::Serialize;
use std::error::Error;
use std::fs::OpenOptions;
use std::io::Write;
use std::path::{Path, PathBuf};

const SCHEMA: &str = "litchi.execution-allocation-observer.v1";
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 10_000;

type AnyResult<T> = Result<T, Box<dyn Error + Send + Sync>>;

#[derive(Debug)]
struct AllocationError(String);

impl std::fmt::Display for AllocationError {
    fn fmt(&self, output: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        output.write_str(&self.0)
    }
}

impl Error for AllocationError {}

fn error(message: impl Into<String>) -> Box<dyn Error + Send + Sync> {
    Box::new(AllocationError(message.into()))
}

struct Config {
    case: probe::Case,
    samples: usize,
    warmup: usize,
    output: PathBuf,
    artifact: Option<PathBuf>,
}

fn bounded(name: &str, value: &str, maximum: usize) -> Result<usize, String> {
    let value = value
        .parse::<usize>()
        .map_err(|_| format!("{name} must be an integer, got {value:?}"))?;
    if value > maximum {
        return Err(format!("{name} must be at most {maximum}, got {value}"));
    }
    Ok(value)
}

fn parse_args() -> Result<Config, String> {
    let mut args = std::env::args().skip(1);
    let mut case = None;
    let mut samples = None;
    let mut warmup = None;
    let mut output = None;
    let mut artifact = None;
    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--case" if case.is_none() => {
                case = Some(
                    args.next()
                        .ok_or("--case requires a value")?
                        .parse::<probe::Case>()?,
                );
            },
            "--samples" if samples.is_none() => {
                let value = args.next().ok_or("--samples requires a value")?;
                samples = Some(bounded("--samples", &value, MAX_SAMPLES)?);
            },
            "--warmup" if warmup.is_none() => {
                let value = args.next().ok_or("--warmup requires a value")?;
                warmup = Some(bounded("--warmup", &value, MAX_WARMUP)?);
            },
            "--output" if output.is_none() => {
                output = Some(PathBuf::from(
                    args.next().ok_or("--output requires a value")?,
                ));
            },
            "--artifact" if artifact.is_none() => {
                artifact = Some(PathBuf::from(
                    args.next().ok_or("--artifact requires a value")?,
                ));
            },
            "--help" | "-h" => {
                return Err(
                    "usage: cfb-emission-probe-alloc --case CASE --samples N --warmup N --output PATH [--artifact PATH]"
                        .to_string(),
                );
            },
            other => return Err(format!("unknown or duplicate argument {other:?}")),
        }
    }
    let samples = samples.ok_or("missing --samples")?;
    if samples == 0 {
        return Err("--samples must be at least 1".to_string());
    }
    Ok(Config {
        case: case.ok_or("missing --case")?,
        samples,
        warmup: warmup.unwrap_or(0),
        output: output.ok_or("missing --output")?,
        artifact,
    })
}

#[derive(Debug, Serialize)]
struct Metrics {
    allocator_identity: &'static str,
    counter_revision: Option<&'static str>,
    instrumentation_identity: &'static str,
    timing: &'static str,
}

#[derive(Debug, Serialize)]
struct AllocationSample {
    sample: usize,
    output_bytes: usize,
    output_sha256: String,
    allocation: allocation_metrics::Sample,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    scope: &'static str,
    config: ConfigReport,
    corpus: serde_json::Value,
    metrics: Metrics,
    samples: Vec<AllocationSample>,
    final_output_sha256: String,
    final_output_bytes: usize,
}

#[derive(Debug, Serialize)]
struct ConfigReport {
    case: String,
    samples: usize,
    warmup: usize,
    measured_operation: &'static str,
    verification: &'static str,
}

fn write_new(path: &Path, bytes: &[u8]) -> AnyResult<()> {
    let mut file = OpenOptions::new().write(true).create_new(true).open(path)?;
    file.write_all(bytes)?;
    file.flush()?;
    Ok(())
}

fn run() -> AnyResult<()> {
    allocation_metrics::enable();
    let config = parse_args().map_err(error)?;
    let input = probe::prepare_case(config.case)?;
    let identity = probe::identity_for_case(config.case, &input);
    let corpus = serde_json::to_value(identity)?;
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or_else(|| error("warmup plus samples overflowed"))?;
    let mut samples = Vec::with_capacity(config.samples);
    let mut final_bytes = None;
    let mut final_hash = None;

    for iteration in 0..total {
        let region = allocation_metrics::begin();
        let result = probe::write_case(&input);
        let allocation = region
            .finish()
            .ok_or_else(|| error("allocator region was unavailable"))?;
        let bytes = result?;
        probe::verify_case(&input, &bytes)?;
        let output_sha256 = sha256_hex(&bytes);
        if iteration >= config.warmup {
            samples.push(AllocationSample {
                sample: iteration - config.warmup,
                output_bytes: bytes.len(),
                output_sha256: output_sha256.clone(),
                allocation,
            });
        }
        if iteration + 1 == total {
            final_hash = Some(output_sha256);
            final_bytes = Some(bytes);
        }
    }

    let final_bytes = final_bytes.ok_or_else(|| error("no final output"))?;
    let final_hash = final_hash.ok_or_else(|| error("no final hash"))?;
    if let Some(path) = config.artifact.as_deref() {
        write_new(path, &final_bytes)?;
    }
    let report = Report {
        schema: SCHEMA,
        scope: "global system allocator callbacks around the same fresh public writer operation; no wall or CPU timing",
        config: ConfigReport {
            case: config.case.name().to_string(),
            samples: config.samples,
            warmup: config.warmup,
            measured_operation: "fresh public writer construction, prepared-input registration, and write_to",
            verification: "outside allocation region: reopen, semantic projection, hashes, and inventory",
        },
        corpus,
        metrics: Metrics {
            allocator_identity: allocation_metrics::allocator_identity(),
            counter_revision: allocation_metrics::counter_revision(),
            instrumentation_identity: allocation_metrics::instrumentation_identity(),
            timing: "not measured by this binary",
        },
        samples,
        final_output_sha256: final_hash,
        final_output_bytes: final_bytes.len(),
    };
    let mut json = serde_json::to_vec_pretty(&report)?;
    json.push(b'\n');
    write_new(&config.output, &json)?;
    Ok(())
}

fn sha256_hex(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    let mut output = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn main() {
    if let Err(error) = run() {
        eprintln!("cfb-emission-probe-alloc: {error}");
        std::process::exit(2);
    }
}
