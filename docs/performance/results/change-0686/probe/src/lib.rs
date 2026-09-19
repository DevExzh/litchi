//! Native retry probe for the bounded XLS worksheet index budget.
//!
//! This package deliberately uses the public source-backed API with the
//! production `OwnedSource` and `FileSource` adapters.  It has no counted
//! `ReadAt` wrapper and no global allocator, so query timings do not include
//! instrumentation atomics.  Allocation measurements for this change use the
//! separate 0686 allocation probe.

use litchi_core::{FileSource, OwnedSource, ReadAt, sheet::CellValue};
use litchi_xls::{SourceBackedLimits, SourceBackedWorkbook};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::error::Error;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Instant;

const DEFAULT_QUERIES: usize = 8;
const DEFAULT_WARMUPS: usize = 3;
const DEFAULT_SAMPLES: usize = 30;

pub type ProbeError = Box<dyn Error + Send + Sync>;
pub type ProbeResult<T> = Result<T, ProbeError>;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Mode {
    Owned,
    File,
}

impl Mode {
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Owned => "owned",
            Self::File => "file",
        }
    }
}

impl std::str::FromStr for Mode {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "owned" => Ok(Self::Owned),
            "file" => Ok(Self::File),
            _ => Err(format!(
                "unknown source mode {value:?}; expected owned or file"
            )),
        }
    }
}

#[derive(Clone, Debug)]
pub struct Config {
    pub input: PathBuf,
    pub budget: u64,
    pub queries: usize,
    pub samples: usize,
    pub warmups: usize,
    pub mode: Mode,
    pub worksheet: usize,
    pub row: u32,
    pub column: u32,
}

impl Default for Config {
    fn default() -> Self {
        Self {
            input: PathBuf::new(),
            budget: 0,
            queries: DEFAULT_QUERIES,
            samples: DEFAULT_SAMPLES,
            warmups: DEFAULT_WARMUPS,
            mode: Mode::Owned,
            worksheet: 0,
            row: 0,
            column: 0,
        }
    }
}

#[derive(Clone, Debug, Serialize, PartialEq)]
pub struct Outcome {
    pub status: &'static str,
    pub value: Option<SemanticValue>,
    pub error: Option<String>,
}

#[derive(Clone, Debug, Serialize, PartialEq)]
#[serde(tag = "kind")]
pub enum SemanticValue {
    Empty,
    Bool {
        value: bool,
    },
    Int {
        value: i64,
    },
    Float {
        bits: u64,
    },
    String {
        value: String,
    },
    DateTime {
        bits: u64,
    },
    Error {
        value: String,
    },
    Formula {
        formula: String,
        cached_value: Option<Box<SemanticValue>>,
        is_array: bool,
        array_range: Option<String>,
    },
}

impl SemanticValue {
    #[must_use]
    pub fn from_cell(value: &CellValue) -> Self {
        match value {
            CellValue::Empty => Self::Empty,
            CellValue::Bool(value) => Self::Bool { value: *value },
            CellValue::Int(value) => Self::Int { value: *value },
            CellValue::Float(value) => Self::Float {
                bits: value.to_bits(),
            },
            CellValue::String(value) => Self::String {
                value: value.clone(),
            },
            CellValue::DateTime(value) => Self::DateTime {
                bits: value.to_bits(),
            },
            CellValue::Error(value) => Self::Error {
                value: value.clone(),
            },
            CellValue::Formula {
                formula,
                cached_value,
                is_array,
                array_range,
            } => Self::Formula {
                formula: formula.clone(),
                cached_value: cached_value.as_deref().map(Self::from_cell).map(Box::new),
                is_array: *is_array,
                array_range: array_range.clone(),
            },
        }
    }
}

impl Outcome {
    #[must_use]
    pub fn from_result(result: &Result<Option<CellValue>, litchi_xls::SourceBackedError>) -> Self {
        match result {
            Ok(Some(value)) => Self {
                status: "value",
                value: Some(SemanticValue::from_cell(value)),
                error: None,
            },
            Ok(None) => Self {
                status: "missing",
                value: None,
                error: None,
            },
            Err(error) => Self {
                status: "error",
                value: None,
                error: Some(error.to_string()),
            },
        }
    }

    #[must_use]
    fn opened() -> Self {
        Self {
            status: "ok",
            value: None,
            error: None,
        }
    }

    #[must_use]
    fn open_error(error: impl Into<String>) -> Self {
        Self {
            status: "error",
            value: None,
            error: Some(error.into()),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
pub struct OpenRecord {
    pub elapsed_ns: u64,
    pub outcome: Outcome,
}

#[derive(Clone, Debug, Serialize)]
pub struct QueryRecord {
    pub ordinal: usize,
    pub elapsed_ns: u64,
    pub outcome: Outcome,
    pub agrees_with_first: bool,
}

#[derive(Clone, Debug, Serialize)]
pub struct SampleRecord {
    pub sample: usize,
    pub open: OpenRecord,
    pub queries: Vec<QueryRecord>,
    pub all_queries_agree: bool,
}

#[derive(Clone, Debug, Serialize)]
pub struct Report {
    pub schema_version: u32,
    pub probe: &'static str,
    pub input_path: String,
    pub input_bytes: u64,
    pub input_sha256: String,
    pub mode: &'static str,
    pub worksheet: usize,
    pub row: u32,
    pub column: u32,
    pub max_query_index_bytes: u64,
    pub queries: usize,
    pub warmups: usize,
    pub samples: usize,
    pub fresh_owner_per_sample: bool,
    pub source_construction_scope: &'static str,
    pub timing_scope: &'static str,
    pub allocator_scope: &'static str,
    pub records: Vec<SampleRecord>,
}

#[derive(Debug)]
struct Input {
    path: PathBuf,
    bytes: Arc<Vec<u8>>,
    sha256: String,
}

fn read_input(path: &Path) -> ProbeResult<Input> {
    let path = fs::canonicalize(path)?;
    let bytes = Arc::new(fs::read(&path)?);
    let sha256 = Sha256::digest(bytes.as_slice())
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect();
    Ok(Input {
        path,
        bytes,
        sha256,
    })
}

fn source_for_sample(input: &Input, mode: Mode) -> ProbeResult<Arc<dyn ReadAt>> {
    match mode {
        Mode::Owned => Ok(Arc::new(OwnedSource::from_arc(Arc::clone(&input.bytes)))),
        Mode::File => Ok(Arc::new(FileSource::open(&input.path)?)),
    }
}

fn elapsed_ns(started: Instant) -> u64 {
    started.elapsed().as_nanos().min(u128::from(u64::MAX)) as u64
}

fn execute_sample(config: &Config, input: &Input, sample: usize) -> ProbeResult<SampleRecord> {
    // Source adapter construction is intentionally outside the open timer.
    // This keeps owned-byte setup and pathname/open work out of the measured
    // source-backed owner opening, matching the 0684 native probe boundary.
    let source = match source_for_sample(input, config.mode) {
        Ok(source) => source,
        Err(error) => {
            return Ok(SampleRecord {
                sample,
                open: OpenRecord {
                    elapsed_ns: 0,
                    outcome: Outcome::open_error(error.to_string()),
                },
                queries: Vec::new(),
                all_queries_agree: false,
            });
        },
    };
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(config.budget);
    let started = Instant::now();
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source, limits);
    let open_elapsed_ns = elapsed_ns(started);
    let owner = match owner {
        Ok(owner) => owner,
        Err(error) => {
            return Ok(SampleRecord {
                sample,
                open: OpenRecord {
                    elapsed_ns: open_elapsed_ns,
                    outcome: Outcome::open_error(error.to_string()),
                },
                queries: Vec::new(),
                all_queries_agree: false,
            });
        },
    };

    let mut retained_values = Vec::with_capacity(config.queries);
    let mut records = Vec::with_capacity(config.queries);
    let mut first_outcome = None;
    let mut all_queries_agree = true;
    for ordinal in 0..config.queries {
        let query_started = Instant::now();
        let result = owner.cell_value_by_index(config.worksheet, config.row, config.column);
        let query_elapsed_ns = elapsed_ns(query_started);
        // Build the complete semantic outcome only after stopping the timer.
        // The result itself stays alive while the outcome is captured and
        // while a successful value is moved into the retained vector.
        let current = Outcome::from_result(&result);
        let agrees_with_first = match &first_outcome {
            None => true,
            Some(first) => first == &current,
        };
        all_queries_agree &= agrees_with_first;
        if first_outcome.is_none() {
            first_outcome = Some(current.clone());
        }
        if let Ok(Some(value)) = result {
            retained_values.push(value);
        }
        records.push(QueryRecord {
            ordinal,
            elapsed_ns: query_elapsed_ns,
            outcome: current,
            agrees_with_first,
        });
    }
    std::hint::black_box((&owner, &retained_values));
    Ok(SampleRecord {
        sample,
        open: OpenRecord {
            elapsed_ns: open_elapsed_ns,
            outcome: Outcome::opened(),
        },
        queries: records,
        all_queries_agree,
    })
}

pub fn run(config: Config) -> ProbeResult<Report> {
    if config.input.as_os_str().is_empty() {
        return Err(io::Error::new(io::ErrorKind::InvalidInput, "--input is required").into());
    }
    if config.samples == 0 {
        return Err(
            io::Error::new(io::ErrorKind::InvalidInput, "--samples must be nonzero").into(),
        );
    }
    if config.warmups == 0 {
        return Err(
            io::Error::new(io::ErrorKind::InvalidInput, "--warmups must be nonzero").into(),
        );
    }
    if config.queries == 0 {
        return Err(
            io::Error::new(io::ErrorKind::InvalidInput, "--queries must be nonzero").into(),
        );
    }
    let input = read_input(&config.input)?;
    let total = config
        .warmups
        .checked_add(config.samples)
        .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "sample count overflow"))?;
    let mut records = Vec::with_capacity(config.samples);
    for iteration in 0..total {
        let record = execute_sample(&config, &input, iteration)?;
        if iteration >= config.warmups {
            records.push(record);
        }
    }
    Ok(Report {
        schema_version: 1,
        probe: "change-0686-xls-index-budget-retry",
        input_path: input.path.display().to_string(),
        input_bytes: u64::try_from(input.bytes.len())?,
        input_sha256: input.sha256,
        mode: config.mode.as_str(),
        worksheet: config.worksheet,
        row: config.row,
        column: config.column,
        max_query_index_bytes: config.budget,
        queries: config.queries,
        warmups: config.warmups,
        samples: config.samples,
        fresh_owner_per_sample: true,
        source_construction_scope: "source adapter is constructed before each open timer; owned bytes are read once before all samples",
        timing_scope: "open.elapsed_ns measures SourceBackedWorkbook construction; query.elapsed_ns measures one cell_value_by_index call only; semantic projection and JSON serialization are outside timers",
        allocator_scope: "native binary uses the system allocator without allocation counters; allocation metrics come from the separate change-0686 allocation probe",
        records,
    })
}

pub fn parse_args(arguments: impl IntoIterator<Item = String>) -> Result<Config, String> {
    let mut arguments = arguments.into_iter();
    let mut config = Config::default();
    let mut budget_seen = false;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--input" => config.input = PathBuf::from(take_value(&mut arguments, "--input")?),
            "--budget" => {
                config.budget = take_value(&mut arguments, "--budget")?
                    .parse()
                    .map_err(|_| "--budget requires a u64".to_owned())?;
                budget_seen = true;
            },
            "--queries" => config.queries = parse_usize(&mut arguments, "--queries")?,
            "--samples" => config.samples = parse_usize(&mut arguments, "--samples")?,
            "--warmups" => config.warmups = parse_usize(&mut arguments, "--warmups")?,
            "--mode" => {
                config.mode = take_value(&mut arguments, "--mode")?.parse()?;
            },
            "--worksheet" => config.worksheet = parse_usize(&mut arguments, "--worksheet")?,
            "--row" => config.row = parse_u32(&mut arguments, "--row")?,
            "--column" => config.column = parse_u32(&mut arguments, "--column")?,
            "--help" | "-h" => return Err(usage().to_owned()),
            value => return Err(format!("unknown option {value:?}\n{}", usage())),
        }
    }
    if !budget_seen {
        return Err("--budget is required".to_owned());
    }
    Ok(config)
}

fn take_value(
    arguments: &mut impl Iterator<Item = String>,
    option: &str,
) -> Result<String, String> {
    arguments
        .next()
        .ok_or_else(|| format!("{option} requires a value"))
}

fn parse_usize(
    arguments: &mut impl Iterator<Item = String>,
    option: &str,
) -> Result<usize, String> {
    take_value(arguments, option)?
        .parse()
        .map_err(|_| format!("{option} requires a non-negative integer"))
}

fn parse_u32(arguments: &mut impl Iterator<Item = String>, option: &str) -> Result<u32, String> {
    take_value(arguments, option)?
        .parse()
        .map_err(|_| format!("{option} requires a u32"))
}

#[must_use]
pub const fn usage() -> &'static str {
    "usage: xls-index-retry-probe-0686 --input FILE --budget BYTES --mode owned|file --worksheet N --row N --column N [--queries N] [--samples N] [--warmups N]"
}

pub fn write_report(report: &Report) -> ProbeResult<()> {
    let stdout = io::stdout();
    let mut output = stdout.lock();
    serde_json::to_writer_pretty(&mut output, report)?;
    use std::io::Write;
    writeln!(output)?;
    Ok(())
}
