//! Standalone source-backed XLS query-cache probe for change 0684.
//!
//! The probe deliberately uses only the public `litchi-xls` source-backed
//! worksheet API.  That keeps one binary buildable from the pre-change
//! checkout and from a candidate checkout, while the counted `ReadAt`
//! adapters make source reads, bytes, versions, and logical file ranges
//! visible in the result.  The timing path keeps semantic projection and
//! digest work outside the measured query/visitor calls.

use litchi_core::{ReadAt, SourceVersion, sheet::CellValue};
use litchi_xls::SourceBackedWorkbook;
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::cmp::min;
use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::fs::{self, File};
use std::io;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use std::time::Instant;

#[cfg(unix)]
use std::os::unix::fs::FileExt;
#[cfg(windows)]
use std::os::windows::fs::FileExt;

pub type ProbeError = Box<dyn Error + Send + Sync>;
pub type ProbeResult<T> = Result<T, ProbeError>;

const DEFAULT_WARMUPS: usize = 2;
const DEFAULT_SAMPLES: usize = 7;
const DEFAULT_SAMPLE_COORDINATES: usize = 16;
const DEFAULT_MAX_QUERIES: usize = 512;

/// The source adapter used by one probe route.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ProbeMode {
    Owned,
    File,
    OwnedNative,
    FileNative,
}

impl ProbeMode {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Owned => "owned",
            Self::File => "file",
            Self::OwnedNative => "owned-native",
            Self::FileNative => "file-native",
        }
    }
}

impl std::str::FromStr for ProbeMode {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "owned" => Ok(Self::Owned),
            "file" => Ok(Self::File),
            "owned-native" => Ok(Self::OwnedNative),
            "file-native" => Ok(Self::FileNative),
            _ => Err(format!(
                "unknown source mode {value:?}; expected owned, file, owned-native, or file-native"
            )),
        }
    }
}

/// One fresh-query route.  `prepared` opens one owner and executes three
/// selected queries (`q1`, `q2`, and `q1` again) against that snapshot.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Route {
    Cold1,
    Cold2,
    Cold3,
    Prepared,
    Visit,
}

impl Route {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Cold1 => "cold1",
            Self::Cold2 => "cold2",
            Self::Cold3 => "cold3",
            Self::Prepared => "prepared",
            Self::Visit => "visit",
        }
    }
}

impl std::str::FromStr for Route {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "cold1" => Ok(Self::Cold1),
            "cold2" => Ok(Self::Cold2),
            "cold3" => Ok(Self::Cold3),
            "prepared" => Ok(Self::Prepared),
            "visit" => Ok(Self::Visit),
            _ => Err(format!(
                "unknown route {value:?}; expected cold1, cold2, cold3, prepared, or visit"
            )),
        }
    }
}

/// Inputs for the route probe.
#[derive(Clone, Debug)]
pub struct RouteConfig {
    pub input: PathBuf,
    pub route: Route,
    pub mode: ProbeMode,
    pub worksheet: usize,
    pub row: u32,
    pub column: u32,
    pub second_row: u32,
    pub second_column: u32,
    pub warmups: usize,
    pub samples: usize,
}

impl RouteConfig {
    pub fn defaults(input: PathBuf, route: Route, mode: ProbeMode) -> Self {
        Self {
            input,
            route,
            mode,
            worksheet: 0,
            row: 0,
            column: 0,
            second_row: 1,
            second_column: 0,
            warmups: DEFAULT_WARMUPS,
            samples: DEFAULT_SAMPLES,
        }
    }
}

/// Inputs for a bounded full-corpus differential.
#[derive(Clone, Debug)]
pub struct CorpusConfig {
    pub root: PathBuf,
    pub mode: ProbeMode,
    pub sample_coordinates: usize,
    pub max_queries: usize,
}

impl Default for CorpusConfig {
    fn default() -> Self {
        Self {
            root: PathBuf::from("test-data"),
            mode: ProbeMode::Owned,
            sample_coordinates: DEFAULT_SAMPLE_COORDINATES,
            max_queries: DEFAULT_MAX_QUERIES,
        }
    }
}

#[derive(Debug, Default)]
pub struct Metrics {
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    version_calls: AtomicU64,
    len_calls: AtomicU64,
    range_union_bytes: AtomicU64,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize)]
pub struct MetricsSnapshot {
    pub read_calls: u64,
    pub read_bytes: u64,
    pub version_calls: u64,
    pub len_calls: u64,
    pub range_union_bytes: u64,
}

impl Metrics {
    fn snapshot(&self) -> MetricsSnapshot {
        MetricsSnapshot {
            read_calls: self.read_calls.load(Ordering::Relaxed),
            read_bytes: self.read_bytes.load(Ordering::Relaxed),
            version_calls: self.version_calls.load(Ordering::Relaxed),
            len_calls: self.len_calls.load(Ordering::Relaxed),
            range_union_bytes: self.range_union_bytes.load(Ordering::Relaxed),
        }
    }

    fn record_read(&self, result: &io::Result<usize>) {
        self.read_calls.fetch_add(1, Ordering::Relaxed);
        if let Ok(count) = result {
            self.read_bytes
                .fetch_add(u64::try_from(*count).unwrap_or(u64::MAX), Ordering::Relaxed);
        }
    }

    fn delta(&self, before: MetricsSnapshot) -> MetricsSnapshot {
        let after = self.snapshot();
        MetricsSnapshot {
            read_calls: after.read_calls.saturating_sub(before.read_calls),
            read_bytes: after.read_bytes.saturating_sub(before.read_bytes),
            version_calls: after.version_calls.saturating_sub(before.version_calls),
            len_calls: after.len_calls.saturating_sub(before.len_calls),
            range_union_bytes: after
                .range_union_bytes
                .saturating_sub(before.range_union_bytes),
        }
    }
}

#[derive(Debug)]
struct RangeUnion {
    ranges: Vec<std::ops::Range<u64>>,
    bytes: u64,
}

impl RangeUnion {
    fn insert(&mut self, range: std::ops::Range<u64>) -> u64 {
        if range.start >= range.end {
            return self.bytes;
        }
        self.ranges.push(range);
        self.ranges.sort_unstable_by_key(|item| item.start);
        let mut merged = Vec::<std::ops::Range<u64>>::with_capacity(self.ranges.len());
        for item in self.ranges.drain(..) {
            if let Some(last) = merged.last_mut()
                && item.start <= last.end
            {
                last.end = last.end.max(item.end);
            } else {
                merged.push(item);
            }
        }
        self.bytes = merged
            .iter()
            .map(|item| item.end.saturating_sub(item.start))
            .sum();
        self.ranges = merged;
        self.bytes
    }
}

#[derive(Debug)]
struct OwnedReadAt {
    bytes: Arc<Vec<u8>>,
    metrics: Arc<Metrics>,
    version: SourceVersion,
}

impl ReadAt for OwnedReadAt {
    fn len(&self) -> io::Result<u64> {
        self.metrics.len_calls.fetch_add(1, Ordering::Relaxed);
        u64::try_from(self.bytes.len()).map_err(|_| io::Error::other("source length overflow"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let result = if let Some(remaining) = self.bytes.get(start..) {
            let count = min(remaining.len(), output.len());
            output[..count].copy_from_slice(&remaining[..count]);
            Ok(count)
        } else {
            Ok(0)
        };
        self.metrics.record_read(&result);
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.metrics.version_calls.fetch_add(1, Ordering::Relaxed);
        Ok(self.version)
    }
}

#[derive(Debug)]
struct FileReadAt {
    file: File,
    length: u64,
    metrics: Arc<Metrics>,
    ranges: Mutex<RangeUnion>,
    version: SourceVersion,
}

impl ReadAt for FileReadAt {
    fn len(&self) -> io::Result<u64> {
        self.metrics.len_calls.fetch_add(1, Ordering::Relaxed);
        Ok(self.length)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let result = positioned_read(&self.file, offset, output);
        self.metrics.record_read(&result);
        if let Ok(count) = &result {
            let mut ranges = self
                .ranges
                .lock()
                .map_err(|_| io::Error::other("range counter mutex poisoned"))?;
            let merged = ranges.insert(offset..offset.saturating_add(*count as u64));
            self.metrics
                .range_union_bytes
                .store(merged, Ordering::Relaxed);
        }
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.metrics.version_calls.fetch_add(1, Ordering::Relaxed);
        Ok(self.version)
    }
}

fn positioned_read(file: &File, offset: u64, output: &mut [u8]) -> io::Result<usize> {
    #[cfg(unix)]
    {
        file.read_at(output, offset)
    }
    #[cfg(windows)]
    {
        file.seek_read(output, offset)
    }
    #[cfg(not(any(unix, windows)))]
    {
        let mut copy = file.try_clone()?;
        use std::io::{Read, Seek, SeekFrom};
        copy.seek(SeekFrom::Start(offset))?;
        copy.read(output)
    }
}

#[derive(Clone, Debug)]
struct Input {
    path: PathBuf,
    bytes: Arc<Vec<u8>>,
    sha256: String,
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn read_input(path: &Path) -> ProbeResult<Input> {
    let path = fs::canonicalize(path)?;
    let bytes = Arc::new(fs::read(&path)?);
    let sha256 = sha256_hex(&bytes);
    Ok(Input {
        path,
        bytes,
        sha256,
    })
}

struct SourceRun {
    source: Arc<dyn ReadAt>,
    metrics: Arc<Metrics>,
}

fn make_source(input: &Input, mode: ProbeMode) -> ProbeResult<SourceRun> {
    let metrics = Arc::new(Metrics::default());
    let source: Arc<dyn ReadAt> = match mode {
        ProbeMode::Owned => Arc::new(OwnedReadAt {
            bytes: Arc::clone(&input.bytes),
            metrics: Arc::clone(&metrics),
            version: SourceVersion::new(1, 0),
        }),
        ProbeMode::File => Arc::new(FileReadAt {
            file: File::open(&input.path)?,
            length: u64::try_from(input.bytes.len())?,
            metrics: Arc::clone(&metrics),
            ranges: Mutex::new(RangeUnion {
                ranges: Vec::new(),
                bytes: 0,
            }),
            version: SourceVersion::new(2, 0),
        }),
        ProbeMode::OwnedNative => {
            Arc::new(litchi_core::OwnedSource::from_arc(Arc::clone(&input.bytes)))
        },
        ProbeMode::FileNative => {
            #[cfg(any(unix, windows))]
            {
                Arc::new(litchi_core::FileSource::open(&input.path)?)
            }
            #[cfg(not(any(unix, windows)))]
            {
                return Err(io::Error::new(
                    io::ErrorKind::Unsupported,
                    "file-native requires a positional filesystem platform",
                )
                .into());
            }
        },
    };
    Ok(SourceRun { source, metrics })
}

fn open_workbook(
    input: &Input,
    mode: ProbeMode,
) -> ProbeResult<(Result<SourceBackedWorkbook, String>, Arc<Metrics>, u64)> {
    let source_run = make_source(input, mode)?;
    let metrics = Arc::clone(&source_run.metrics);
    let started = Instant::now();
    let result =
        SourceBackedWorkbook::from_read_at(source_run.source).map_err(|error| error.to_string());
    let elapsed_ns = started.elapsed().as_nanos().min(u128::from(u64::MAX)) as u64;
    Ok((result, metrics, elapsed_ns))
}

fn compact_projection(value: &CellValue) -> String {
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
                .map(compact_projection)
                .unwrap_or_else(|| "none".to_owned()),
        ),
    }
}

fn fnv_hex(bytes: &[u8]) -> String {
    let mut hash = 0xcbf29ce484222325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    format!("{hash:016x}")
}

fn cells_digest(cells: &[(u32, u32, CellValue)]) -> String {
    let mut hasher = Sha256::new();
    for (row, column, value) in cells {
        hasher.update(row.to_le_bytes());
        hasher.update(column.to_le_bytes());
        hasher.update(compact_projection(value).as_bytes());
        hasher.update([b'\n']);
    }
    format!(
        "cells:{}:{}",
        cells.len(),
        hasher
            .finalize()
            .iter()
            .map(|byte| format!("{byte:02x}"))
            .collect::<String>()
    )
}

#[derive(Clone, Debug, Serialize)]
pub struct Outcome {
    pub status: String,
    pub row: Option<u32>,
    pub column: Option<u32>,
    pub projection: Option<String>,
    pub error: Option<String>,
}

impl Outcome {
    fn opened() -> Self {
        Self {
            status: "ok".to_owned(),
            row: None,
            column: None,
            projection: None,
            error: None,
        }
    }

    fn value(row: u32, column: u32, value: &CellValue) -> Self {
        Self {
            status: "value".to_owned(),
            row: Some(row),
            column: Some(column),
            projection: Some(compact_projection(value)),
            error: None,
        }
    }

    fn missing(row: u32, column: u32) -> Self {
        Self {
            status: "missing".to_owned(),
            row: Some(row),
            column: Some(column),
            projection: None,
            error: None,
        }
    }

    fn error(error: impl Into<String>) -> Self {
        Self {
            status: "error".to_owned(),
            row: None,
            column: None,
            projection: None,
            error: Some(error.into()),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
pub struct OpenPhase {
    pub elapsed_ns: u64,
    pub metrics: MetricsSnapshot,
    pub outcome: Outcome,
}

#[derive(Clone, Debug, Serialize)]
pub struct QueryPhase {
    pub label: String,
    pub elapsed_ns: u64,
    pub metrics: MetricsSnapshot,
    pub outcome: Outcome,
}

#[derive(Clone, Debug, Serialize)]
pub struct VisitPhase {
    pub elapsed_ns: u64,
    pub metrics: MetricsSnapshot,
    pub outcome: Outcome,
    pub actual_callbacks: u64,
    pub oracle_callbacks: u64,
    pub oracle_digest: Option<String>,
    pub oracle_metrics: MetricsSnapshot,
}

#[derive(Clone, Debug, Serialize)]
pub struct RouteSample {
    pub sample: usize,
    pub total_elapsed_ns: u64,
    pub open: OpenPhase,
    pub query_phases: Vec<QueryPhase>,
    pub visit: Option<VisitPhase>,
}

#[derive(Clone, Debug, Serialize)]
pub struct RouteReport {
    pub schema_version: u32,
    pub probe: &'static str,
    pub route: &'static str,
    pub mode: &'static str,
    pub input_path: String,
    pub input_bytes: u64,
    pub input_sha256: String,
    pub worksheet: usize,
    pub first_row: u32,
    pub first_column: u32,
    pub second_row: u32,
    pub second_column: u32,
    pub warmups: usize,
    pub samples: usize,
    pub source_scope: &'static str,
    pub timing_scope: &'static str,
    pub digest_scope: &'static str,
    pub fresh_snapshot_per_sample: bool,
    pub records: Vec<RouteSample>,
}

#[derive(Debug, Default)]
pub struct RetainedResult {
    values: Vec<CellValue>,
}

impl RetainedResult {
    /// Number of semantic values held alive after the measured operation.
    #[must_use]
    pub fn value_count(&self) -> usize {
        self.values.len()
    }
}

#[derive(Debug)]
pub struct RouteRun {
    pub report: RouteReport,
    retained: RetainedResult,
}

impl RouteRun {
    /// Keeps the measured operation's returned values alive for an allocator
    /// gauge, without exposing implementation-specific cell internals.
    #[must_use]
    pub fn retained_value_count(&self) -> usize {
        self.retained.value_count()
    }
}

fn query_phase(
    owner: &SourceBackedWorkbook,
    metrics: &Metrics,
    label: &str,
    worksheet: usize,
    row: u32,
    column: u32,
    retained: &mut RetainedResult,
) -> QueryPhase {
    let before = metrics.snapshot();
    let started = Instant::now();
    let result = owner.cell_value_by_index(worksheet, row, column);
    let elapsed_ns = started.elapsed().as_nanos().min(u128::from(u64::MAX)) as u64;
    let outcome = match result {
        Ok(Some(value)) => {
            let outcome = Outcome::value(row, column, &value);
            retained.values.push(value);
            outcome
        },
        Ok(None) => Outcome::missing(row, column),
        Err(error) => Outcome::error(error.to_string()),
    };
    QueryPhase {
        label: label.to_owned(),
        elapsed_ns,
        metrics: metrics.delta(before),
        outcome,
    }
}

fn visit_phase(
    owner: &SourceBackedWorkbook,
    metrics: &Metrics,
    worksheet: usize,
    retained: &mut RetainedResult,
) -> VisitPhase {
    let before = metrics.snapshot();
    let mut actual_callbacks = 0_u64;
    let started = Instant::now();
    let result = match owner.worksheet_by_index(worksheet) {
        Ok(Some(sheet)) => sheet.visit_cells(|_cell| {
            actual_callbacks = actual_callbacks.saturating_add(1);
            Ok(())
        }),
        Ok(None) => Err(litchi_xls::SourceBackedError::WorksheetNotFound(
            worksheet.to_string(),
        )),
        Err(error) => Err(error),
    };
    let elapsed_ns = started.elapsed().as_nanos().min(u128::from(u64::MAX)) as u64;
    let outcome = match result {
        Ok(()) => Outcome::opened(),
        Err(error) => Outcome::error(error.to_string()),
    };
    let metrics_delta = metrics.delta(before);

    // Digest and retention run as a separate oracle operation.  Its reads are
    // reported separately, so digest work never contaminates the native visit
    // timing or its source counters.
    let oracle_before = metrics.snapshot();
    let mut oracle_callbacks = 0_u64;
    let mut oracle_cells = Vec::new();
    let oracle_digest = if outcome.status == "ok" {
        match owner.worksheet_by_index(worksheet) {
            Ok(Some(sheet)) => match sheet.visit_cells(|cell| {
                oracle_callbacks = oracle_callbacks.saturating_add(1);
                oracle_cells.push((cell.row(), cell.column(), cell.into_value()));
                Ok(())
            }) {
                Ok(()) => {
                    for (_, _, value) in &oracle_cells {
                        retained.values.push(value.clone());
                    }
                    Some(cells_digest(&oracle_cells))
                },
                Err(error) => Some(format!("error:{}", error)),
            },
            Ok(None) => Some("error:worksheet not found".to_owned()),
            Err(error) => Some(format!("error:{error}")),
        }
    } else {
        None
    };
    let oracle_metrics = metrics.delta(oracle_before);
    VisitPhase {
        elapsed_ns,
        metrics: metrics_delta,
        outcome,
        actual_callbacks,
        oracle_callbacks,
        oracle_digest,
        oracle_metrics,
    }
}

fn execute_sample(
    config: &RouteConfig,
    input: &Input,
    sample: usize,
) -> ProbeResult<(RouteSample, RetainedResult)> {
    let sample_started = Instant::now();
    let (owner_result, metrics, open_elapsed_ns) = open_workbook(input, config.mode)?;
    let open_before = MetricsSnapshot::default();
    let open_outcome = match &owner_result {
        Ok(_) => Outcome::opened(),
        Err(error) => Outcome::error(error.clone()),
    };
    let open_metrics = metrics.delta(open_before);
    let open_phase = OpenPhase {
        elapsed_ns: open_elapsed_ns,
        metrics: open_metrics,
        outcome: open_outcome,
    };
    let mut retained = RetainedResult::default();
    let mut query_phases = Vec::new();
    let mut visit = None;
    if let Ok(owner) = owner_result {
        match config.route {
            Route::Cold1 | Route::Cold2 | Route::Cold3 => {
                query_phases.push(query_phase(
                    &owner,
                    &metrics,
                    "q1",
                    config.worksheet,
                    config.row,
                    config.column,
                    &mut retained,
                ));
            },
            Route::Prepared => {
                query_phases.push(query_phase(
                    &owner,
                    &metrics,
                    "q1",
                    config.worksheet,
                    config.row,
                    config.column,
                    &mut retained,
                ));
                query_phases.push(query_phase(
                    &owner,
                    &metrics,
                    "q2-build-trigger",
                    config.worksheet,
                    config.second_row,
                    config.second_column,
                    &mut retained,
                ));
                query_phases.push(query_phase(
                    &owner,
                    &metrics,
                    "q3-warm-repeat-q1",
                    config.worksheet,
                    config.row,
                    config.column,
                    &mut retained,
                ));
            },
            Route::Visit => {
                visit = Some(visit_phase(
                    &owner,
                    &metrics,
                    config.worksheet,
                    &mut retained,
                ));
            },
        }
    }
    let total_elapsed_ns = sample_started
        .elapsed()
        .as_nanos()
        .min(u128::from(u64::MAX)) as u64;
    Ok((
        RouteSample {
            sample,
            total_elapsed_ns,
            open: open_phase,
            query_phases,
            visit,
        },
        retained,
    ))
}

/// Runs a route with fresh source-backed owner state for every sample.
pub fn run_route(config: RouteConfig) -> ProbeResult<RouteRun> {
    if config.warmups == 0 || config.samples == 0 {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "warmups and samples must be nonzero",
        )
        .into());
    }
    let input = read_input(&config.input)?;
    let total = config
        .warmups
        .checked_add(config.samples)
        .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "sample count overflow"))?;
    let mut records = Vec::with_capacity(config.samples);
    let mut retained = RetainedResult::default();
    for iteration in 0..total {
        let (record, result) = execute_sample(&config, &input, iteration)?;
        if iteration >= config.warmups {
            if iteration == total - 1 {
                retained = result;
            }
            records.push(record);
        }
    }
    Ok(RouteRun {
        report: RouteReport {
            schema_version: 1,
            probe: "change-0684-xls-index",
            route: config.route.as_str(),
            mode: config.mode.as_str(),
            input_path: input.path.display().to_string(),
            input_bytes: u64::try_from(input.bytes.len())?,
            input_sha256: input.sha256,
            worksheet: config.worksheet,
            first_row: config.row,
            first_column: config.column,
            second_row: config.second_row,
            second_column: config.second_column,
            warmups: config.warmups,
            samples: config.samples,
            source_scope: match config.mode {
                ProbeMode::Owned => "counted immutable in-memory ReadAt; no filesystem I/O",
                ProbeMode::File => {
                    "counted positional file ReadAt; logical reads and unioned ranges; physical I/O not measured"
                },
                ProbeMode::OwnedNative => {
                    "native litchi-core OwnedSource; source counters unavailable; no filesystem I/O"
                },
                ProbeMode::FileNative => {
                    "native litchi-core FileSource; source counters unavailable; logical warm-cache file access"
                },
            },
            timing_scope: "query/visitor elapsed_ns excludes semantic projection and digest work; total_elapsed_ns includes route setup for diagnostics",
            digest_scope: "compact value projection and visitor digest are outside the measured query/visitor call",
            fresh_snapshot_per_sample: true,
            records,
        },
        retained,
    })
}

#[derive(Clone, Debug, Serialize)]
pub struct CorpusQuery {
    pub row: u32,
    pub column: u32,
    pub reasons: Vec<String>,
    pub expected: Option<String>,
    pub actual: Option<String>,
    pub error: Option<String>,
    pub agrees: bool,
}

#[derive(Clone, Debug, Serialize)]
pub struct CorpusSheet {
    pub worksheet: usize,
    pub name: Option<String>,
    pub status: String,
    pub visitor_callbacks: u64,
    pub visitor_digest: Option<String>,
    pub distinct_coordinates: usize,
    pub duplicate_coordinates: usize,
    pub candidate_coordinates: usize,
    pub sampled_coordinates: usize,
    pub queried_coordinates: usize,
    pub query_cap: usize,
    pub query_truncated: bool,
    pub queries: Vec<CorpusQuery>,
    pub repeat: Option<CorpusQuery>,
    pub error: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct CorpusFile {
    pub path: String,
    pub bytes: u64,
    pub sha256: String,
    pub status: String,
    pub open_error: Option<String>,
    pub worksheets: Vec<CorpusSheet>,
}

#[derive(Clone, Debug, Serialize)]
pub struct CorpusReport {
    pub schema_version: u32,
    pub probe: &'static str,
    pub root: String,
    pub mode: &'static str,
    pub sample_coordinates: usize,
    pub max_queries: usize,
    pub files_seen: usize,
    pub files_opened: usize,
    pub files_refused: usize,
    pub sheets_walked: usize,
    pub query_mismatches: usize,
    pub files: Vec<CorpusFile>,
}

#[derive(Debug)]
struct ObservedCell {
    row: u32,
    column: u32,
    value: CellValue,
}

fn corpus_paths(root: &Path) -> ProbeResult<Vec<PathBuf>> {
    fn visit(path: &Path, output: &mut Vec<PathBuf>) -> ProbeResult<()> {
        for entry in fs::read_dir(path)? {
            let entry = entry?;
            let path = entry.path();
            if path.is_dir() {
                visit(&path, output)?;
            } else if path.is_file()
                && matches!(
                    path.extension().and_then(|value| value.to_str()),
                    Some("xls" | "XLS" | "xlt" | "XLT")
                )
            {
                output.push(path);
            }
        }
        Ok(())
    }
    let mut paths = Vec::new();
    visit(root, &mut paths)?;
    paths.sort();
    Ok(paths)
}

fn add_reason(
    reasons: &mut BTreeMap<(u32, u32), BTreeSet<String>>,
    point: (u32, u32),
    reason: &str,
) {
    reasons.entry(point).or_default().insert(reason.to_owned());
}

fn expected_projection(value: Option<&CellValue>) -> Option<String> {
    value.map(compact_projection)
}

fn run_corpus_sheet(
    owner: &SourceBackedWorkbook,
    worksheet: usize,
    name: Option<String>,
    sample_coordinates: usize,
    max_queries: usize,
) -> ProbeResult<CorpusSheet> {
    let Some(sheet) = owner
        .worksheet_by_index(worksheet)
        .map_err(|error| io::Error::other(error.to_string()))?
    else {
        return Ok(CorpusSheet {
            worksheet,
            name,
            status: "refused".to_owned(),
            visitor_callbacks: 0,
            visitor_digest: None,
            distinct_coordinates: 0,
            duplicate_coordinates: 0,
            candidate_coordinates: 0,
            sampled_coordinates: 0,
            queried_coordinates: 0,
            query_cap: max_queries,
            query_truncated: false,
            queries: Vec::new(),
            repeat: None,
            error: Some("worksheet not found".to_owned()),
        });
    };
    let mut observed = Vec::new();
    let visitor_result = sheet.visit_cells(|cell| {
        observed.push(ObservedCell {
            row: cell.row(),
            column: cell.column(),
            value: cell.into_value(),
        });
        Ok(())
    });
    if let Err(error) = visitor_result {
        return Ok(CorpusSheet {
            worksheet,
            name,
            status: "refused".to_owned(),
            visitor_callbacks: observed.len() as u64,
            visitor_digest: None,
            distinct_coordinates: 0,
            duplicate_coordinates: 0,
            candidate_coordinates: 0,
            sampled_coordinates: 0,
            queried_coordinates: 0,
            query_cap: max_queries,
            query_truncated: false,
            queries: Vec::new(),
            repeat: None,
            error: Some(error.to_string()),
        });
    }

    let visitor_digest = {
        let all_occurrences: Vec<_> = observed
            .iter()
            .map(|cell| (cell.row, cell.column, cell.value.clone()))
            .collect();
        cells_digest(&all_occurrences)
    };

    let mut last_by_coordinate = BTreeMap::<(u32, u32), CellValue>::new();
    let mut occurrences = BTreeMap::<(u32, u32), usize>::new();
    let mut order = Vec::new();
    for cell in observed {
        let point = (cell.row, cell.column);
        if !occurrences.contains_key(&point) {
            order.push(point);
        }
        *occurrences.entry(point).or_default() += 1;
        last_by_coordinate.insert(point, cell.value);
    }
    let mut reasons = BTreeMap::<(u32, u32), BTreeSet<String>>::new();
    for (point, count) in &occurrences {
        if *count > 1 {
            add_reason(&mut reasons, *point, "duplicate-coordinate");
        }
    }
    if let Some(first) = order.first().copied() {
        add_reason(&mut reasons, first, "first-stored");
    }
    if let Some(last) = order.last().copied() {
        add_reason(&mut reasons, last, "last-stored");
    }
    let sample_count = min(sample_coordinates, order.len());
    if sample_count == 1 {
        add_reason(&mut reasons, order[0], "deterministic-sample");
    } else if sample_count > 1 {
        for index in 0..sample_count {
            let position = index.saturating_mul(order.len() - 1) / (sample_count - 1);
            add_reason(&mut reasons, order[position], "deterministic-sample");
        }
    }
    add_reason(&mut reasons, (0, 0), "missing-target");
    let max_row = order.iter().map(|point| point.0).max().unwrap_or(0);
    let max_column = order.iter().map(|point| point.1).max().unwrap_or(0);
    if max_row < u32::MAX {
        add_reason(&mut reasons, (max_row + 1, 0), "missing-target");
    }
    if max_column < u32::MAX {
        add_reason(&mut reasons, (0, max_column + 1), "missing-target");
    }

    // Required duplicate and edge coordinates are never dropped.  The cap
    // applies only to ordinary sampled coordinates, which keeps the corpus
    // run bounded while still proving every duplicate ordering case.
    let mut selected = Vec::new();
    let mut ordinary_count = 0;
    for (point, point_reasons) in &reasons {
        let required = point_reasons.iter().any(|reason| {
            matches!(
                reason.as_str(),
                "duplicate-coordinate" | "first-stored" | "last-stored" | "missing-target"
            )
        });
        if required || ordinary_count < max_queries {
            if !required {
                ordinary_count += 1;
            }
            selected.push(*point);
        }
    }
    let query_truncated = reasons.len() > selected.len();
    let mut queries = Vec::with_capacity(selected.len());
    for (row, column) in &selected {
        let expected_value = last_by_coordinate.get(&(*row, *column));
        let result = owner.cell_value_by_index(worksheet, *row, *column);
        let (actual, error, agrees) = match result {
            Ok(Some(value)) => {
                let actual = compact_projection(&value);
                let agrees = expected_value == Some(&value);
                (Some(actual), None, agrees)
            },
            Ok(None) => {
                let agrees = expected_value.is_none();
                (None, None, agrees)
            },
            Err(error) => (None, Some(error.to_string()), false),
        };
        queries.push(CorpusQuery {
            row: *row,
            column: *column,
            reasons: reasons
                .get(&(*row, *column))
                .map(|values| values.iter().cloned().collect())
                .unwrap_or_default(),
            expected: expected_projection(expected_value),
            actual,
            error,
            agrees,
        });
    }
    let repeat = selected.first().and_then(|(row, column)| {
        let expected_value = last_by_coordinate.get(&(*row, *column));
        match owner.cell_value_by_index(worksheet, *row, *column) {
            Ok(Some(value)) => Some(CorpusQuery {
                row: *row,
                column: *column,
                reasons: vec!["warm-repeat".to_owned()],
                expected: expected_projection(expected_value),
                actual: Some(compact_projection(&value)),
                error: None,
                agrees: expected_value == Some(&value),
            }),
            Ok(None) => Some(CorpusQuery {
                row: *row,
                column: *column,
                reasons: vec!["warm-repeat".to_owned()],
                expected: expected_projection(expected_value),
                actual: None,
                error: None,
                agrees: expected_value.is_none(),
            }),
            Err(error) => Some(CorpusQuery {
                row: *row,
                column: *column,
                reasons: vec!["warm-repeat".to_owned()],
                expected: expected_projection(expected_value),
                actual: None,
                error: Some(error.to_string()),
                agrees: false,
            }),
        }
    });
    let duplicate_coordinates = occurrences.values().filter(|count| **count > 1).count();
    Ok(CorpusSheet {
        worksheet,
        name,
        status: "ok".to_owned(),
        visitor_callbacks: occurrences.values().sum::<usize>() as u64,
        visitor_digest: Some(visitor_digest),
        distinct_coordinates: last_by_coordinate.len(),
        duplicate_coordinates,
        candidate_coordinates: reasons.len(),
        sampled_coordinates: selected.len(),
        queried_coordinates: queries.len(),
        query_cap: max_queries,
        query_truncated,
        queries,
        repeat,
        error: None,
    })
}

/// Runs the selected-coordinate differential over every `.xls`/`.xlt` file
/// beneath `config.root`.  Open and worksheet refusals are retained as typed
/// display strings; an eligible visitor is checked against selected queries
/// including all duplicate coordinates, stored edges, missing targets, and a
/// deterministic bounded sample.
pub fn run_corpus(config: CorpusConfig) -> ProbeResult<CorpusReport> {
    if config.sample_coordinates == 0 || config.max_queries == 0 {
        return Err(
            io::Error::new(io::ErrorKind::InvalidInput, "corpus bounds must be nonzero").into(),
        );
    }
    let root_path = fs::canonicalize(&config.root)?;
    let paths = corpus_paths(&root_path)?;
    let mut files = Vec::with_capacity(paths.len());
    let mut files_opened = 0;
    let mut files_refused = 0;
    let mut sheets_walked = 0;
    let mut query_mismatches = 0;
    for path in paths {
        let input = read_input(&path)?;
        let relative_path = input
            .path
            .strip_prefix(&root_path)
            .unwrap_or(&input.path)
            .display()
            .to_string();
        let (owner_result, _metrics, _open_elapsed_ns) = open_workbook(&input, config.mode)?;
        match owner_result {
            Err(error) => {
                files_refused += 1;
                files.push(CorpusFile {
                    path: relative_path,
                    bytes: u64::try_from(input.bytes.len())?,
                    sha256: input.sha256,
                    status: "refused".to_owned(),
                    open_error: Some(error),
                    worksheets: Vec::new(),
                });
            },
            Ok(owner) => {
                files_opened += 1;
                let descriptors = owner
                    .worksheets()
                    .map_err(|error| io::Error::other(error.to_string()))?;
                let mut worksheets = Vec::with_capacity(descriptors.len());
                for descriptor in descriptors {
                    let worksheet = descriptor
                        .index()
                        .map_err(|error| io::Error::other(error.to_string()))?;
                    let name = descriptor
                        .name()
                        .map_err(|error| io::Error::other(error.to_string()))?;
                    let sheet = run_corpus_sheet(
                        &owner,
                        worksheet,
                        Some(name),
                        config.sample_coordinates,
                        config.max_queries,
                    )?;
                    sheets_walked += 1;
                    query_mismatches += sheet.queries.iter().filter(|query| !query.agrees).count();
                    if let Some(repeat) = &sheet.repeat
                        && !repeat.agrees
                    {
                        query_mismatches += 1;
                    }
                    worksheets.push(sheet);
                }
                files.push(CorpusFile {
                    path: relative_path,
                    bytes: u64::try_from(input.bytes.len())?,
                    sha256: input.sha256,
                    status: "opened".to_owned(),
                    open_error: None,
                    worksheets,
                });
            },
        }
    }
    Ok(CorpusReport {
        schema_version: 1,
        probe: "change-0684-xls-index-corpus",
        root: "<corpus-root>".to_owned(),
        mode: config.mode.as_str(),
        sample_coordinates: config.sample_coordinates,
        max_queries: config.max_queries,
        files_seen: files.len(),
        files_opened,
        files_refused,
        sheets_walked,
        query_mismatches,
        files,
    })
}

pub fn default_warmups() -> usize {
    DEFAULT_WARMUPS
}

pub fn default_samples() -> usize {
    DEFAULT_SAMPLES
}

pub fn default_sample_coordinates() -> usize {
    DEFAULT_SAMPLE_COORDINATES
}

pub fn default_max_queries() -> usize {
    DEFAULT_MAX_QUERIES
}
