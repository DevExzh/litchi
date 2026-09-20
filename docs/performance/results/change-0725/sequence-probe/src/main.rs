//! Public-API route probe for one owner and an ordered worksheet query chain.
//!
//! The first two queries use the same late coordinate so the second query
//! publishes the worksheet index.  The following first/origin-late/missing
//! queries exercise the target-local checkpoint's forward and backward
//! selection rules.  This binary has no timers; stderr is reserved for the
//! production trace markers injected by the packet driver.

use litchi_core::{OwnedSource, ReadAt};
use litchi_xls::{SourceBackedLimits, SourceBackedWorkbook};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::error::Error;
use std::fs;
use std::io;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

const DEFAULT_BUDGET: u64 = 2_097_152;
const DEFAULT_WORKSHEET: usize = 0;
const DEFAULT_LATE_ROW: u32 = 5_095;
const DEFAULT_LATE_COLUMN: u32 = 106;
const DEFAULT_FIRST_ROW: u32 = 0;
const DEFAULT_FIRST_COLUMN: u32 = 0;
const DEFAULT_MISSING_ROW: u32 = 0;
const DEFAULT_MISSING_COLUMN: u32 = 108;

#[derive(Clone, Copy)]
struct Coordinate {
    row: u32,
    column: u32,
}

#[derive(Default)]
struct Config {
    input: Option<PathBuf>,
    budget: u64,
    worksheet: usize,
    late: Option<Coordinate>,
    first: Option<Coordinate>,
    missing: Option<Coordinate>,
    intermediate: Option<Coordinate>,
    later: Option<Coordinate>,
}

#[derive(Serialize)]
struct QueryRecord {
    label: String,
    row: u32,
    column: u32,
    outcome: Outcome,
    io: IoSnapshot,
}

#[derive(Serialize)]
struct Outcome {
    kind: &'static str,
    value_debug: Option<String>,
    error: Option<String>,
}

#[derive(Clone, Serialize)]
struct ReadRange {
    offset: u64,
    requested: usize,
    returned: usize,
    error: Option<String>,
}

#[derive(Clone, Serialize)]
struct IoSnapshot {
    len_calls: u64,
    version_calls: u64,
    reads: Vec<ReadRange>,
}

#[derive(Serialize)]
struct Report {
    schema_version: u32,
    probe: &'static str,
    input_path: String,
    input_sha256: String,
    worksheet: usize,
    budget: u64,
    open_io: IoSnapshot,
    queries: Vec<QueryRecord>,
}

struct TracedSource {
    source: OwnedSource,
    reads: Mutex<Vec<ReadRange>>,
    len_calls: AtomicU64,
    version_calls: AtomicU64,
}

impl TracedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            source: OwnedSource::new(bytes),
            reads: Mutex::new(Vec::new()),
            len_calls: AtomicU64::new(0),
            version_calls: AtomicU64::new(0),
        }
    }

    fn reset(&self) {
        self.reads.lock().unwrap().clear();
        self.len_calls.store(0, Ordering::Relaxed);
        self.version_calls.store(0, Ordering::Relaxed);
    }

    fn snapshot(&self) -> IoSnapshot {
        IoSnapshot {
            len_calls: self.len_calls.load(Ordering::Relaxed),
            version_calls: self.version_calls.load(Ordering::Relaxed),
            reads: self.reads.lock().unwrap().clone(),
        }
    }
}

impl ReadAt for TracedSource {
    fn len(&self) -> io::Result<u64> {
        self.len_calls.fetch_add(1, Ordering::Relaxed);
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let requested = output.len();
        let result = self.source.read_at(offset, output);
        let (returned, error) = match &result {
            Ok(returned) => (*returned, None),
            Err(error) => (0, Some(error.to_string())),
        };
        self.reads.lock().unwrap().push(ReadRange {
            offset,
            requested,
            returned,
            error,
        });
        result
    }

    fn version(&self) -> io::Result<litchi_core::SourceVersion> {
        self.version_calls.fetch_add(1, Ordering::Relaxed);
        self.source.version()
    }
}

fn parse_u32(value: Option<String>, option: &str) -> Result<u32, Box<dyn Error>> {
    value
        .ok_or_else(|| format!("{option} requires a value"))?
        .parse::<u32>()
        .map_err(|_| format!("{option} requires a u32").into())
}

fn parse_usize(value: Option<String>, option: &str) -> Result<usize, Box<dyn Error>> {
    value
        .ok_or_else(|| format!("{option} requires a value"))?
        .parse::<usize>()
        .map_err(|_| format!("{option} requires a non-negative integer").into())
}

fn take(args: &mut impl Iterator<Item = String>, option: &str) -> Result<String, Box<dyn Error>> {
    args.next()
        .ok_or_else(|| format!("{option} requires a value").into())
}

fn parse_args(args: impl IntoIterator<Item = String>) -> Result<Config, Box<dyn Error>> {
    let mut config = Config {
        budget: DEFAULT_BUDGET,
        worksheet: DEFAULT_WORKSHEET,
        late: Some(Coordinate {
            row: DEFAULT_LATE_ROW,
            column: DEFAULT_LATE_COLUMN,
        }),
        first: Some(Coordinate {
            row: DEFAULT_FIRST_ROW,
            column: DEFAULT_FIRST_COLUMN,
        }),
        missing: Some(Coordinate {
            row: DEFAULT_MISSING_ROW,
            column: DEFAULT_MISSING_COLUMN,
        }),
        ..Config::default()
    };
    let mut args = args.into_iter();
    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--input" => config.input = Some(PathBuf::from(take(&mut args, "--input")?)),
            "--budget" => config.budget = take(&mut args, "--budget")?.parse()?,
            "--worksheet" => {
                config.worksheet =
                    parse_usize(Some(take(&mut args, "--worksheet")?), "--worksheet")?
            },
            "--late-row" => {
                // Coordinates are an ordered row/column flag pair.
                // Read the row below, then validate the named column flag.
                let row = parse_u32(Some(take(&mut args, "--late-row")?), "--late-row")?;
                if take(&mut args, "--late-column")? != "--late-column" {
                    return Err("expected --late-column after --late-row".into());
                }
                let column = parse_u32(Some(take(&mut args, "--late-column")?), "--late-column")?;
                config.late = Some(Coordinate { row, column });
            },
            "--first-row" => {
                // Coordinates are an ordered row/column flag pair.
                // Read the row below, then validate the named column flag.
                let row = parse_u32(Some(take(&mut args, "--first-row")?), "--first-row")?;
                if take(&mut args, "--first-column")? != "--first-column" {
                    return Err("expected --first-column after --first-row".into());
                }
                let column = parse_u32(Some(take(&mut args, "--first-column")?), "--first-column")?;
                config.first = Some(Coordinate { row, column });
            },
            "--missing-row" => {
                // Coordinates are an ordered row/column flag pair.
                // Read the row below, then validate the named column flag.
                let row = parse_u32(Some(take(&mut args, "--missing-row")?), "--missing-row")?;
                if take(&mut args, "--missing-column")? != "--missing-column" {
                    return Err("expected --missing-column after --missing-row".into());
                }
                let column = parse_u32(
                    Some(take(&mut args, "--missing-column")?),
                    "--missing-column",
                )?;
                config.missing = Some(Coordinate { row, column });
            },
            "--intermediate-row" => {
                // Coordinates are an ordered row/column flag pair.
                // Read the row below, then validate the named column flag.
                let row = parse_u32(
                    Some(take(&mut args, "--intermediate-row")?),
                    "--intermediate-row",
                )?;
                if take(&mut args, "--intermediate-column")? != "--intermediate-column" {
                    return Err("expected --intermediate-column after --intermediate-row".into());
                }
                let column = parse_u32(
                    Some(take(&mut args, "--intermediate-column")?),
                    "--intermediate-column",
                )?;
                config.intermediate = Some(Coordinate { row, column });
            },
            "--later-row" => {
                // Coordinates are an ordered row/column flag pair.
                // Read the row below, then validate the named column flag.
                let row = parse_u32(Some(take(&mut args, "--later-row")?), "--later-row")?;
                if take(&mut args, "--later-column")? != "--later-column" {
                    return Err("expected --later-column after --later-row".into());
                }
                let column = parse_u32(Some(take(&mut args, "--later-column")?), "--later-column")?;
                config.later = Some(Coordinate { row, column });
            },
            "--help" | "-h" => {
                return Err(usage().to_owned().into());
            },
            value => return Err(format!("unknown option {value:?}\n{}", usage()).into()),
        }
    }
    if config.input.is_none() {
        return Err("--input is required".into());
    }
    Ok(config)
}

fn usage() -> &'static str {
    "usage: xls-query-chain-sequence-0725 --input FILE [--budget BYTES] [--worksheet N] [--late-row N --late-column N] [--first-row N --first-column N] [--missing-row N --missing-column N] [--intermediate-row N --intermediate-column N] [--later-row N --later-column N]"
}

fn query(
    sheet: &litchi_xls::SourceBackedWorksheet,
    source: &TracedSource,
    label: &str,
    coordinate: Coordinate,
) -> QueryRecord {
    source.reset();
    let outcome = match sheet.cell_value(coordinate.row, coordinate.column) {
        Ok(Some(cell)) => Outcome {
            kind: "value",
            value_debug: Some(format!("{:?}", cell)),
            error: None,
        },
        Ok(None) => Outcome {
            kind: "missing",
            value_debug: None,
            error: None,
        },
        Err(error) => Outcome {
            kind: "error",
            value_debug: None,
            error: Some(error.to_string()),
        },
    };
    QueryRecord {
        label: label.to_owned(),
        row: coordinate.row,
        column: coordinate.column,
        outcome,
        io: source.snapshot(),
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let config = parse_args(std::env::args().skip(1))?;
    let input = config.input.expect("validated input");
    let bytes = fs::read(&input)?;
    let input_sha256: String = Sha256::digest(&bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect();
    let traced = Arc::new(TracedSource::new(bytes));
    let source: Arc<dyn ReadAt> = traced.clone();
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(config.budget);
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source, limits)?;
    let sheet = owner
        .worksheet_by_index(config.worksheet)?
        .ok_or_else(|| format!("worksheet {} is absent", config.worksheet))?;
    let open_io = traced.snapshot();

    let late = config.late.expect("default late coordinate");
    let first = config.first.expect("default first coordinate");
    let missing = config.missing.expect("default missing coordinate");
    let mut queries = Vec::new();
    queries.push(query(&sheet, &traced, "late-build", late));
    queries.push(query(&sheet, &traced, "late-publish", late));
    queries.push(query(&sheet, &traced, "first-earlier", first));
    if let Some(intermediate) = config.intermediate {
        queries.push(query(&sheet, &traced, "intermediate", intermediate));
    }
    queries.push(query(&sheet, &traced, "origin-late", late));
    if let Some(later) = config.later {
        queries.push(query(&sheet, &traced, "later", later));
    }
    queries.push(query(&sheet, &traced, "missing", missing));

    serde_json::to_writer_pretty(
        std::io::stdout(),
        &Report {
            schema_version: 1,
            probe: "change-0725-xls-query-chain-sequence",
            input_path: input.display().to_string(),
            input_sha256,
            worksheet: config.worksheet,
            budget: config.budget,
            open_io,
            queries,
        },
    )?;
    println!();
    Ok(())
}
