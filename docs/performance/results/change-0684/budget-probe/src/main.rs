use litchi_core::{ReadAt, SourceVersion, sheet::CellValue};
use litchi_xls::{SourceBackedLimits, SourceBackedWorkbook};
use serde::Serialize;
use serde_json::to_writer_pretty;
use sha2::{Digest, Sha256};
use std::env;
use std::fs;
use std::io;
use std::path::PathBuf;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;

const DEFAULT_QUERIES: usize = 5;

#[derive(Debug, Default)]
struct Metrics {
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    version_calls: AtomicU64,
    len_calls: AtomicU64,
}

#[derive(Clone, Copy, Debug, Default, Serialize)]
struct MetricSnapshot {
    read_calls: u64,
    read_bytes: u64,
    version_calls: u64,
    len_calls: u64,
}

impl Metrics {
    fn snapshot(&self) -> MetricSnapshot {
        MetricSnapshot {
            read_calls: self.read_calls.load(Ordering::Relaxed),
            read_bytes: self.read_bytes.load(Ordering::Relaxed),
            version_calls: self.version_calls.load(Ordering::Relaxed),
            len_calls: self.len_calls.load(Ordering::Relaxed),
        }
    }

    fn delta(&self, before: MetricSnapshot) -> MetricSnapshot {
        let after = self.snapshot();
        MetricSnapshot {
            read_calls: after.read_calls.saturating_sub(before.read_calls),
            read_bytes: after.read_bytes.saturating_sub(before.read_bytes),
            version_calls: after.version_calls.saturating_sub(before.version_calls),
            len_calls: after.len_calls.saturating_sub(before.len_calls),
        }
    }
}

struct CountingOwned {
    bytes: Arc<Vec<u8>>,
    metrics: Arc<Metrics>,
    version: SourceVersion,
}

impl ReadAt for CountingOwned {
    fn len(&self) -> io::Result<u64> {
        self.metrics.len_calls.fetch_add(1, Ordering::Relaxed);
        u64::try_from(self.bytes.len()).map_err(|_| io::Error::other("source length overflow"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let result = if let Some(remaining) = self.bytes.get(start..) {
            let count = remaining.len().min(output.len());
            output[..count].copy_from_slice(&remaining[..count]);
            Ok(count)
        } else {
            Ok(0)
        };
        self.metrics.read_calls.fetch_add(1, Ordering::Relaxed);
        if let Ok(count) = &result {
            self.metrics
                .read_bytes
                .fetch_add(u64::try_from(*count).unwrap_or(u64::MAX), Ordering::Relaxed);
        }
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.metrics.version_calls.fetch_add(1, Ordering::Relaxed);
        Ok(self.version)
    }
}

#[derive(Clone, Debug, Serialize)]
struct Outcome {
    status: String,
    projection: Option<String>,
    error: Option<String>,
}

fn fnv_hex(bytes: &[u8]) -> String {
    let mut hash = 0xcbf29ce484222325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    format!("{hash:016x}")
}

fn projection(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "empty".to_owned(),
        CellValue::Bool(value) => format!("bool:{value}"),
        CellValue::Int(value) => format!("int:{value}"),
        CellValue::Float(value) => format!("float:{:016x}", value.to_bits()),
        CellValue::String(value) => format!("string:{}:{}", value.len(), fnv_hex(value.as_bytes())),
        CellValue::DateTime(value) => format!("datetime:{:016x}", value.to_bits()),
        CellValue::Error(value) => format!("error:{}:{}", value.len(), fnv_hex(value.as_bytes())),
        CellValue::Formula {
            formula,
            cached_value,
            is_array,
            array_range,
        } => format!(
            "formula:{}:{}:array={is_array}:range={}:cached={}",
            formula.len(),
            fnv_hex(formula.as_bytes()),
            array_range
                .as_deref()
                .map(|range| fnv_hex(range.as_bytes()))
                .unwrap_or_else(|| "none".to_owned()),
            cached_value
                .as_deref()
                .map(projection)
                .unwrap_or_else(|| "none".to_owned()),
        ),
    }
}

fn outcome(
    result: Result<Option<CellValue>, litchi_xls::SourceBackedError>,
) -> (Outcome, Option<CellValue>) {
    match result {
        Ok(Some(value)) => (
            Outcome {
                status: "value".to_owned(),
                projection: Some(projection(&value)),
                error: None,
            },
            Some(value),
        ),
        Ok(None) => (
            Outcome {
                status: "missing".to_owned(),
                projection: None,
                error: None,
            },
            None,
        ),
        Err(error) => (
            Outcome {
                status: "error".to_owned(),
                projection: None,
                error: Some(error.to_string()),
            },
            None,
        ),
    }
}

#[derive(Serialize)]
struct QueryRecord {
    ordinal: usize,
    elapsed_ns: u64,
    metrics: MetricSnapshot,
    outcome: Outcome,
    agrees_with_first: bool,
}

#[derive(Serialize)]
struct Report {
    schema_version: u32,
    probe: &'static str,
    input_path: String,
    input_bytes: u64,
    input_sha256: String,
    worksheet: usize,
    row: u32,
    column: u32,
    max_query_index_bytes: u64,
    queries_requested: usize,
    open_elapsed_ns: u64,
    open_metrics: MetricSnapshot,
    open_error: Option<String>,
    queries: Vec<QueryRecord>,
    all_queries_agree: bool,
    timing_scope: &'static str,
    source_scope: &'static str,
}

fn sha256_hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn parse_usize(
    arguments: &mut impl Iterator<Item = String>,
    option: &str,
) -> Result<usize, String> {
    arguments
        .next()
        .ok_or_else(|| format!("{option} requires a value"))?
        .parse()
        .map_err(|_| format!("{option} requires a non-negative integer"))
}

fn parse_u32(arguments: &mut impl Iterator<Item = String>, option: &str) -> Result<u32, String> {
    arguments
        .next()
        .ok_or_else(|| format!("{option} requires a value"))?
        .parse()
        .map_err(|_| format!("{option} requires a u32"))
}

fn usage() -> &'static str {
    "usage: xls-index-budget-probe-0684 --input FILE --worksheet N --row N --column N --budget BYTES [--queries N]"
}

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let mut arguments = env::args().skip(1);
    let mut input = None;
    let mut worksheet = 0_usize;
    let mut row = 0_u32;
    let mut column = 0_u32;
    let mut budget = None;
    let mut queries = DEFAULT_QUERIES;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--input" => {
                input = Some(PathBuf::from(
                    arguments.next().ok_or("--input requires a file path")?,
                ));
            },
            "--worksheet" => worksheet = parse_usize(&mut arguments, "--worksheet")?,
            "--row" => row = parse_u32(&mut arguments, "--row")?,
            "--column" => column = parse_u32(&mut arguments, "--column")?,
            "--budget" => budget = Some(parse_usize(&mut arguments, "--budget")? as u64),
            "--queries" => queries = parse_usize(&mut arguments, "--queries")?,
            "--help" | "-h" => {
                println!("{}", usage());
                return Ok(());
            },
            value => {
                return Err(
                    io::Error::other(format!("unknown option {value:?}\n{}", usage())).into(),
                );
            },
        }
    }
    let input = input.ok_or("--input is required")?;
    let budget = budget.ok_or("--budget is required")?;
    if queries == 0 {
        return Err(io::Error::other("--queries must be nonzero").into());
    }
    let path = fs::canonicalize(&input)?;
    let bytes = Arc::new(fs::read(&path)?);
    let input_sha256 = sha256_hex(&bytes);
    let metrics = Arc::new(Metrics::default());
    let source = CountingOwned {
        bytes: Arc::clone(&bytes),
        metrics: Arc::clone(&metrics),
        version: SourceVersion::new(684, 0),
    };
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(budget);
    let started = Instant::now();
    let owner = SourceBackedWorkbook::from_read_at_with_limits(Arc::new(source), limits)
        .map_err(|error| error.to_string());
    let open_elapsed_ns = started.elapsed().as_nanos().min(u128::from(u64::MAX)) as u64;
    let open_metrics = metrics.snapshot();
    let mut retained_values = Vec::new();
    let (queries_output, all_queries_agree, open_error) = match owner {
        Err(error) => (Vec::new(), false, Some(error)),
        Ok(owner) => {
            let mut records = Vec::with_capacity(queries);
            let mut first_value = None;
            let mut first_outcome = None;
            let mut all_agree = true;
            for ordinal in 0..queries {
                let before = metrics.snapshot();
                let query_started = Instant::now();
                let result = owner.cell_value_by_index(worksheet, row, column);
                let elapsed_ns =
                    query_started.elapsed().as_nanos().min(u128::from(u64::MAX)) as u64;
                let (current, value) = outcome(result);
                let value_for_compare = value.clone();
                if let Some(value) = value {
                    if first_value.is_none() {
                        first_value = Some(value.clone());
                    }
                    retained_values.push(value);
                }
                if first_outcome.is_none() {
                    first_outcome = Some(current.clone());
                }
                let agrees = first_outcome.as_ref().is_some_and(|first| {
                    first.status == current.status
                        && first.projection == current.projection
                        && first.error == current.error
                        && (first.status != "value"
                            || first_value.as_ref() == value_for_compare.as_ref())
                });
                all_agree &= agrees;
                records.push(QueryRecord {
                    ordinal,
                    elapsed_ns,
                    metrics: metrics.delta(before),
                    outcome: current,
                    agrees_with_first: agrees,
                });
            }
            (records, all_agree, None)
        },
    };
    std::hint::black_box(&retained_values);
    let report = Report {
        schema_version: 1,
        probe: "change-0684-xls-index-budget",
        input_path: path.display().to_string(),
        input_bytes: u64::try_from(bytes.len())?,
        input_sha256,
        worksheet,
        row,
        column,
        max_query_index_bytes: budget,
        queries_requested: queries,
        open_elapsed_ns,
        open_metrics,
        open_error,
        queries: queries_output,
        all_queries_agree,
        timing_scope: "elapsed_ns is a diagnostic around one selected query; source counters are the primary comparison and include only the counted owned ReadAt",
        source_scope: "counted immutable in-memory ReadAt with stable SourceVersion(684,0); no filesystem or physical-I/O result",
    };
    let stdout = io::stdout();
    let mut output = stdout.lock();
    to_writer_pretty(&mut output, &report)?;
    use std::io::Write;
    writeln!(output)?;
    Ok(())
}
