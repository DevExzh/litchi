use std::{
    env,
    error::Error,
    fs::{self, File, OpenOptions},
    io::{self, Cursor, Read, Seek, SeekFrom, Write},
    ops::Range,
    path::{Path, PathBuf},
    sync::{
        Arc, Mutex,
        atomic::{AtomicU64, Ordering},
    },
    time::Instant,
};

#[cfg(unix)]
use std::os::unix::fs::{FileExt, OpenOptionsExt};
#[cfg(windows)]
use std::os::windows::fs::FileExt;

use litchi_cfb::OleFile;
#[cfg(any(unix, windows))]
use litchi_core::FileSource;
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_core::{sheet::Cell as _, sheet::CellValue};
use serde::Serialize;
use sha2::{Digest, Sha256};

type AnyError = Box<dyn Error + Send + Sync>;

const DEFAULT_WARMUPS: usize = 3;
const DEFAULT_SAMPLES: usize = 15;
const DEFAULT_WORKSHEET_INDEX: usize = 1;
const DEFAULT_ROW: u32 = 1;
const DEFAULT_COLUMN: u32 = 0;

const PROCESS_SCOPE: &str =
    "samples share one benchmark process; invoke once per retained sample for fresh-child evidence";
const INPUT_SCOPE: &str = "the original input is copied once outside timing into a private read-only staged file; all path-backed samples reopen that snapshot; full-file staging and post-run SHA-256 verification are outside elapsed_ns, warm the page cache, and do not measure cold or physical I/O";
const ORACLE_SCOPE: &str = "implementation-local compatibility projections only; no independent source-parser oracle and no source/eager equivalence assertion";
static NEXT_STAGED_INPUT_ID: AtomicU64 = AtomicU64::new(0);

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    Open,
    List,
    OneCell,
}

impl Operation {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Open => "open",
            Self::List => "list",
            Self::OneCell => "one-cell",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Mode {
    OwnedReadAt,
    AtomicFile,
    TrackedFile,
    FileSource,
    EagerFile,
    FacadeFile,
}

impl Mode {
    const fn as_str(self) -> &'static str {
        match self {
            Self::OwnedReadAt => "owned-readat",
            Self::AtomicFile => "atomic-file",
            Self::TrackedFile => "tracked-file",
            Self::FileSource => "file-source",
            Self::EagerFile => "eager-file",
            Self::FacadeFile => "facade-file",
        }
    }

    const fn counter_scope(self) -> &'static str {
        match self {
            Self::OwnedReadAt => "custom immutable ReadAt over staged bytes; no filesystem I/O",
            Self::AtomicFile => {
                "custom positional FileExt ReadAt over a staged immutable file; logical reads only; warm page-cache and no physical-I/O counter"
            },
            Self::TrackedFile => {
                "custom positional ReadAt over a staged immutable file with mutex-protected File and range-union counters; logical reads only; warm page-cache and no physical-I/O counter"
            },
            Self::FileSource => {
                "FileSource ReadAt wrapper over a staged immutable file; len/version/read calls are counted; warm page-cache and no physical-I/O counter"
            },
            Self::EagerFile => {
                "counting Read/Seek wrapper around the eager raw CFB/XLS owner; logical reads only; warm page-cache and no physical-I/O counter"
            },
            Self::FacadeFile => {
                "facade file API; filesystem counters are not exposed; staged warm page-cache and no physical-I/O counter"
            },
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::OwnedReadAt | Self::AtomicFile | Self::TrackedFile | Self::FileSource => {
                "source-backed family: staged source construction is outside elapsed_ns; production CFB/XLS open/query and internal source probes are timed; compare elapsed_ns only within this family, never against eager-file or facade-file"
            },
            Self::EagerFile => {
                "eager-file family: staged file open, CFB open, XLS owner construction, and query are inside elapsed_ns; compare elapsed_ns only within eager-file, never against source-backed or facade-file"
            },
            Self::FacadeFile => {
                "facade-file family: public facade open/detection and query are inside elapsed_ns; compare elapsed_ns only within facade-file, never against source-backed or eager-file"
            },
        }
    }

    const fn limitation(self, operation: Operation) -> Option<&'static str> {
        match self {
            Self::FacadeFile if matches!(operation, Operation::OneCell) => Some(
                "unsupported: the unified facade exposes worksheet listing but not selected-cell access",
            ),
            Self::FacadeFile => Some(
                "facade-file uses the unified public Workbook API; filesystem counters are not exposed; the staged fixture is immutable and warm-cache, so physical I/O is not measured",
            ),
            Self::EagerFile => Some(
                "eager compatibility control; source construction and timing scope are not paired with positional modes; the staged fixture is immutable and warm-cache",
            ),
            Self::AtomicFile | Self::TrackedFile => Some(
                "synthetic fixed SourceVersion::new(1, 0) over an immutable staged fixture; freshness/version behavior is not representative",
            ),
            Self::FileSource => Some(
                "the staged fixture is immutable and warm-cache; this measures logical FileSource calls, not physical I/O",
            ),
            Self::OwnedReadAt => None,
        }
    }
}

#[derive(Clone, Debug)]
struct Config {
    input: PathBuf,
    operation: Operation,
    mode: Mode,
    warmups: usize,
    samples: usize,
    worksheet_index: usize,
    row: u32,
    column: u32,
}

#[derive(Clone, Copy, Debug)]
struct Coordinates {
    worksheet_index: usize,
    row: u32,
    column: u32,
}

#[derive(Debug)]
struct InputSnapshot {
    path: PathBuf,
    bytes: Arc<Vec<u8>>,
    sha256: String,
    staged_path: PathBuf,
}

impl Drop for InputSnapshot {
    fn drop(&mut self) {
        remove_staged_file(&self.staged_path);
    }
}

#[derive(Debug, Serialize)]
struct Identity {
    path: String,
    bytes: u64,
    sha256: String,
}

#[derive(Debug, Serialize)]
struct ToolIdentity {
    name: &'static str,
    version: &'static str,
    revision: String,
}

#[derive(Debug, Serialize)]
struct SemanticProjection {
    worksheet_count: usize,
    worksheet_names: Vec<String>,
    selected_cell: Option<String>,
}

#[derive(Debug, Serialize)]
struct SemanticOracle {
    source_implementation_projection: SemanticProjection,
    eager_implementation_projection: SemanticProjection,
    scope: &'static str,
}

#[derive(Debug, Clone, Serialize)]
struct ObservationJson {
    kind: &'static str,
    worksheet_count: Option<usize>,
    worksheet_names: Option<Vec<String>>,
    cell: Option<String>,
}

#[derive(Debug)]
enum Observation {
    Open { worksheet_count: usize },
    List { worksheet_names: Vec<String> },
    OneCell { cell: Option<CellValue> },
}

impl Observation {
    fn into_json(self) -> ObservationJson {
        match self {
            Self::Open { worksheet_count } => ObservationJson {
                kind: "open",
                worksheet_count: Some(worksheet_count),
                worksheet_names: None,
                cell: None,
            },
            Self::List { worksheet_names } => ObservationJson {
                kind: "list",
                worksheet_count: None,
                worksheet_names: Some(worksheet_names),
                cell: None,
            },
            Self::OneCell { cell } => ObservationJson {
                kind: "one-cell",
                worksheet_count: None,
                worksheet_names: None,
                cell: cell.as_ref().map(cell_projection),
            },
        }
    }
}

#[derive(Clone, Copy, Debug, Default, Serialize)]
struct MetricsSnapshot {
    read_calls: u64,
    read_bytes: u64,
    version_calls: u64,
    len_calls: u64,
    read_ns: u64,
    version_ns: u64,
    len_ns: u64,
    mutex_ns: u64,
    range_union_ns: u64,
    range_union_bytes: u64,
    seek_calls: u64,
}

impl MetricsSnapshot {
    fn delta(self, baseline: Self) -> Self {
        Self {
            read_calls: self.read_calls.saturating_sub(baseline.read_calls),
            read_bytes: self.read_bytes.saturating_sub(baseline.read_bytes),
            version_calls: self.version_calls.saturating_sub(baseline.version_calls),
            len_calls: self.len_calls.saturating_sub(baseline.len_calls),
            read_ns: self.read_ns.saturating_sub(baseline.read_ns),
            version_ns: self.version_ns.saturating_sub(baseline.version_ns),
            len_ns: self.len_ns.saturating_sub(baseline.len_ns),
            mutex_ns: self.mutex_ns.saturating_sub(baseline.mutex_ns),
            range_union_ns: self.range_union_ns.saturating_sub(baseline.range_union_ns),
            range_union_bytes: self
                .range_union_bytes
                .saturating_sub(baseline.range_union_bytes),
            seek_calls: self.seek_calls.saturating_sub(baseline.seek_calls),
        }
    }
}

#[derive(Debug, Default)]
struct Metrics {
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    version_calls: AtomicU64,
    len_calls: AtomicU64,
    read_ns: AtomicU64,
    version_ns: AtomicU64,
    len_ns: AtomicU64,
    mutex_ns: AtomicU64,
    range_union_ns: AtomicU64,
    range_union_bytes: AtomicU64,
    seek_calls: AtomicU64,
}

impl Metrics {
    fn snapshot(&self) -> MetricsSnapshot {
        MetricsSnapshot {
            read_calls: self.read_calls.load(Ordering::Relaxed),
            read_bytes: self.read_bytes.load(Ordering::Relaxed),
            version_calls: self.version_calls.load(Ordering::Relaxed),
            len_calls: self.len_calls.load(Ordering::Relaxed),
            read_ns: self.read_ns.load(Ordering::Relaxed),
            version_ns: self.version_ns.load(Ordering::Relaxed),
            len_ns: self.len_ns.load(Ordering::Relaxed),
            mutex_ns: self.mutex_ns.load(Ordering::Relaxed),
            range_union_ns: self.range_union_ns.load(Ordering::Relaxed),
            range_union_bytes: self.range_union_bytes.load(Ordering::Relaxed),
            seek_calls: self.seek_calls.load(Ordering::Relaxed),
        }
    }

    fn add_duration(counter: &AtomicU64, duration: std::time::Duration) {
        let nanos = duration.as_nanos().min(u128::from(u64::MAX)) as u64;
        counter.fetch_add(nanos, Ordering::Relaxed);
    }

    fn record_read(&self, result: &io::Result<usize>, duration: std::time::Duration) {
        self.read_calls.fetch_add(1, Ordering::Relaxed);
        if let Ok(count) = result {
            self.read_bytes
                .fetch_add(u64::try_from(*count).unwrap_or(u64::MAX), Ordering::Relaxed);
        }
        Self::add_duration(&self.read_ns, duration);
    }
}

#[derive(Debug, Default)]
struct RangeUnion {
    ranges: Vec<Range<u64>>,
    bytes: u64,
}

impl RangeUnion {
    fn insert(&mut self, range: Range<u64>) -> u64 {
        if range.start >= range.end {
            return self.bytes;
        }
        self.ranges.push(range);
        self.ranges.sort_unstable_by_key(|item| item.start);
        let mut merged: Vec<Range<u64>> = Vec::with_capacity(self.ranges.len());
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
        let started = Instant::now();
        self.metrics.len_calls.fetch_add(1, Ordering::Relaxed);
        let result = u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("owned source length exceeds u64"));
        Metrics::add_duration(&self.metrics.len_ns, started.elapsed());
        result
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let started = Instant::now();
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let result = if let Some(remaining) = self.bytes.get(start..) {
            let count = remaining.len().min(output.len());
            output[..count].copy_from_slice(&remaining[..count]);
            Ok(count)
        } else {
            Ok(0)
        };
        self.metrics.record_read(&result, started.elapsed());
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let started = Instant::now();
        self.metrics.version_calls.fetch_add(1, Ordering::Relaxed);
        let result = Ok(self.version);
        Metrics::add_duration(&self.metrics.version_ns, started.elapsed());
        result
    }
}

#[derive(Debug)]
struct AtomicFileReadAt {
    file: File,
    length: u64,
    metrics: Arc<Metrics>,
    version: SourceVersion,
}

impl ReadAt for AtomicFileReadAt {
    fn len(&self) -> io::Result<u64> {
        let started = Instant::now();
        self.metrics.len_calls.fetch_add(1, Ordering::Relaxed);
        let result = Ok(self.length);
        Metrics::add_duration(&self.metrics.len_ns, started.elapsed());
        result
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let started = Instant::now();
        let result = read_file_at(&self.file, offset, output);
        self.metrics.record_read(&result, started.elapsed());
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let started = Instant::now();
        self.metrics.version_calls.fetch_add(1, Ordering::Relaxed);
        let result = Ok(self.version);
        Metrics::add_duration(&self.metrics.version_ns, started.elapsed());
        result
    }
}

#[derive(Debug)]
struct TrackedFileReadAt {
    file: Mutex<File>,
    length: u64,
    metrics: Arc<Metrics>,
    ranges: Mutex<RangeUnion>,
    version: SourceVersion,
}

impl ReadAt for TrackedFileReadAt {
    fn len(&self) -> io::Result<u64> {
        let started = Instant::now();
        self.metrics.len_calls.fetch_add(1, Ordering::Relaxed);
        let result = Ok(self.length);
        Metrics::add_duration(&self.metrics.len_ns, started.elapsed());
        result
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let lock_started = Instant::now();
        let file = self
            .file
            .lock()
            .map_err(|_| io::Error::other("tracked file mutex is poisoned"))?;
        Metrics::add_duration(&self.metrics.mutex_ns, lock_started.elapsed());

        let started = Instant::now();
        let result = read_file_at(&file, offset, output);
        self.metrics.record_read(&result, started.elapsed());
        let count = result.as_ref().map_or(0, |value| *value);
        drop(file);

        let union_started = Instant::now();
        let mut ranges = self
            .ranges
            .lock()
            .map_err(|_| io::Error::other("tracked range mutex is poisoned"))?;
        let union_bytes = ranges.insert(offset..offset.saturating_add(count as u64));
        drop(ranges);
        Metrics::add_duration(&self.metrics.range_union_ns, union_started.elapsed());
        self.metrics
            .range_union_bytes
            .store(union_bytes, Ordering::Relaxed);
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let started = Instant::now();
        self.metrics.version_calls.fetch_add(1, Ordering::Relaxed);
        let result = Ok(self.version);
        Metrics::add_duration(&self.metrics.version_ns, started.elapsed());
        result
    }
}

fn read_file_at(file: &File, offset: u64, output: &mut [u8]) -> io::Result<usize> {
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
        copy.seek(SeekFrom::Start(offset))?;
        copy.read(output)
    }
}

#[cfg(any(unix, windows))]
#[derive(Debug)]
struct FileSourceReadAt {
    source: FileSource,
    metrics: Arc<Metrics>,
}

#[cfg(any(unix, windows))]
impl ReadAt for FileSourceReadAt {
    fn len(&self) -> io::Result<u64> {
        let started = Instant::now();
        self.metrics.len_calls.fetch_add(1, Ordering::Relaxed);
        let result = self.source.len();
        Metrics::add_duration(&self.metrics.len_ns, started.elapsed());
        result
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let started = Instant::now();
        let result = self.source.read_at(offset, output);
        self.metrics.record_read(&result, started.elapsed());
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let started = Instant::now();
        self.metrics.version_calls.fetch_add(1, Ordering::Relaxed);
        let result = self.source.version();
        Metrics::add_duration(&self.metrics.version_ns, started.elapsed());
        result
    }
}

#[derive(Debug)]
struct CountingFile {
    file: File,
    metrics: Arc<Metrics>,
}

impl Read for CountingFile {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let started = Instant::now();
        let result = self.file.read(output);
        self.metrics.record_read(&result, started.elapsed());
        result
    }
}

impl Seek for CountingFile {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.metrics.seek_calls.fetch_add(1, Ordering::Relaxed);
        self.file.seek(position)
    }
}

#[derive(Debug, Serialize)]
struct EagerPhaseTiming {
    cfb_open_ns: u64,
    xls_owner_ns: u64,
    selected_query_ns: u64,
    cfb_open_read_calls: u64,
    cfb_open_read_bytes: u64,
    xls_owner_read_calls: u64,
    xls_owner_read_bytes: u64,
    selected_query_read_calls: u64,
    selected_query_read_bytes: u64,
}

#[derive(Debug, Serialize)]
struct Sample {
    elapsed_ns: u64,
    observation: ObservationJson,
    metrics: MetricsSnapshot,
    source_version_stable: Option<bool>,
    eager_phases: Option<EagerPhaseTiming>,
}

#[derive(Debug, Serialize)]
struct Report {
    schema_version: u32,
    revision: String,
    tool: ToolIdentity,
    binary: Identity,
    input: Identity,
    mode: &'static str,
    operation: &'static str,
    counter_scope: &'static str,
    limitation: Option<&'static str>,
    timing_scope: &'static str,
    process_scope: &'static str,
    input_scope: &'static str,
    warmups: usize,
    samples: usize,
    worksheet_index: usize,
    row: u32,
    column: u32,
    semantic_oracle: SemanticOracle,
    elapsed_samples_ns: Vec<u64>,
    records: Vec<Sample>,
}

struct SourceRun {
    source: Arc<dyn ReadAt>,
    metrics: Arc<Metrics>,
}

fn make_source(mode: Mode, input: &InputSnapshot) -> Result<SourceRun, AnyError> {
    let metrics = Arc::new(Metrics::default());
    let version = SourceVersion::new(1, 0);
    match mode {
        Mode::OwnedReadAt => {
            let source = OwnedReadAt {
                bytes: Arc::clone(&input.bytes),
                metrics: Arc::clone(&metrics),
                version,
            };
            Ok(SourceRun {
                source: Arc::new(source),
                metrics,
            })
        },
        Mode::AtomicFile => {
            let source = AtomicFileReadAt {
                file: File::open(&input.staged_path)?,
                length: u64::try_from(input.bytes.len())?,
                metrics: Arc::clone(&metrics),
                version,
            };
            Ok(SourceRun {
                source: Arc::new(source),
                metrics,
            })
        },
        Mode::TrackedFile => {
            let source = TrackedFileReadAt {
                file: Mutex::new(File::open(&input.staged_path)?),
                length: u64::try_from(input.bytes.len())?,
                metrics: Arc::clone(&metrics),
                ranges: Mutex::new(RangeUnion::default()),
                version,
            };
            Ok(SourceRun {
                source: Arc::new(source),
                metrics,
            })
        },
        #[cfg(any(unix, windows))]
        Mode::FileSource => {
            let source = FileSourceReadAt {
                source: FileSource::open(&input.staged_path)?,
                metrics: Arc::clone(&metrics),
            };
            Ok(SourceRun {
                source: Arc::new(source),
                metrics,
            })
        },
        #[cfg(not(any(unix, windows)))]
        Mode::FileSource => Err(io::Error::new(
            io::ErrorKind::Unsupported,
            "file-source mode requires a positional filesystem platform",
        )
        .into()),
        Mode::EagerFile | Mode::FacadeFile => Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "file-owner mode does not construct a ReadAt source",
        )
        .into()),
    }
}

fn source_operation(
    workbook: &litchi_xls::SourceBackedWorkbook,
    operation: Operation,
    coordinates: Coordinates,
) -> Result<Observation, AnyError> {
    match operation {
        Operation::Open => Ok(Observation::Open {
            worksheet_count: std::hint::black_box(workbook.worksheet_count()?),
        }),
        Operation::List => Ok(Observation::List {
            worksheet_names: std::hint::black_box(workbook.worksheet_names()?),
        }),
        Operation::OneCell => Ok(Observation::OneCell {
            cell: std::hint::black_box(workbook.cell_value_by_index(
                coordinates.worksheet_index,
                coordinates.row,
                coordinates.column,
            )?),
        }),
    }
}

fn eager_operation<R: Read + Seek>(
    workbook: &litchi_xls::Workbook<R>,
    operation: Operation,
    coordinates: Coordinates,
) -> Result<Observation, AnyError> {
    match operation {
        Operation::Open => Ok(Observation::Open {
            worksheet_count: std::hint::black_box(
                workbook
                    .sheets()
                    .iter()
                    .filter(|sheet| sheet.parsed_worksheet_index().is_some())
                    .count(),
            ),
        }),
        Operation::List => Ok(Observation::List {
            worksheet_names: std::hint::black_box(
                workbook
                    .sheets()
                    .iter()
                    .filter(|sheet| sheet.parsed_worksheet_index().is_some())
                    .map(|sheet| sheet.name().to_owned())
                    .collect(),
            ),
        }),
        Operation::OneCell => Ok(Observation::OneCell {
            cell: std::hint::black_box(
                workbook
                    .xls_worksheet(coordinates.worksheet_index)?
                    .get_cell(coordinates.row, coordinates.column)
                    .map(|cell| cell.value().clone()),
            ),
        }),
    }
}

fn eager_names<R: Read + Seek>(workbook: &litchi_xls::Workbook<R>) -> Vec<String> {
    workbook
        .sheets()
        .iter()
        .filter(|sheet| sheet.parsed_worksheet_index().is_some())
        .map(|sheet| sheet.name().to_owned())
        .collect()
}

fn cell_projection(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "empty".to_owned(),
        CellValue::Bool(value) => format!("bool:{value}"),
        CellValue::Int(value) => format!("int:{value}"),
        CellValue::Float(value) => format!("float:{:016x}", value.to_bits()),
        CellValue::String(value) => format!("string:{}:{value}", value.len()),
        CellValue::DateTime(value) => format!("datetime:{value}"),
        CellValue::Error(value) => format!("error:{}:{value}", value.len()),
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => {
            let cached = cached_value
                .as_deref()
                .map(cell_projection)
                .unwrap_or_else(|| "none".to_owned());
            format!("formula:{}:{formula}:cached:{cached}", formula.len())
        },
    }
}

fn build_oracle(
    input: &[u8],
    operation: Operation,
    coordinates: Coordinates,
) -> Result<SemanticOracle, AnyError> {
    let source_workbook =
        litchi_xls::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(input.to_vec())))?;
    let source_worksheet_count = source_workbook.worksheet_count()?;
    let source_worksheet_names = source_workbook.worksheet_names()?;
    let source_selected_cell = if operation == Operation::OneCell {
        source_workbook
            .cell_value_by_index(
                coordinates.worksheet_index,
                coordinates.row,
                coordinates.column,
            )?
            .as_ref()
            .map(cell_projection)
    } else {
        None
    };

    let eager_workbook = litchi_xls::Workbook::new(Cursor::new(input.to_vec()))?;
    let eager_worksheet_names = eager_names(&eager_workbook);
    let eager_selected_cell = if operation == Operation::OneCell {
        eager_workbook
            .xls_worksheet(coordinates.worksheet_index)?
            .get_cell(coordinates.row, coordinates.column)
            .map(|cell| cell_projection(cell.value()))
    } else {
        None
    };
    Ok(SemanticOracle {
        source_implementation_projection: SemanticProjection {
            worksheet_count: source_worksheet_count,
            worksheet_names: source_worksheet_names,
            selected_cell: source_selected_cell,
        },
        eager_implementation_projection: SemanticProjection {
            worksheet_count: eager_worksheet_names.len(),
            worksheet_names: eager_worksheet_names,
            selected_cell: eager_selected_cell,
        },
        scope: ORACLE_SCOPE,
    })
}

fn validate_source_observation(
    observation: &Observation,
    oracle: &SemanticOracle,
) -> Result<(), AnyError> {
    let projection = &oracle.source_implementation_projection;
    match observation {
        Observation::Open { worksheet_count } => {
            if *worksheet_count != projection.worksheet_count {
                return Err(io::Error::other("source semantic oracle mismatch").into());
            }
        },
        Observation::List { worksheet_names } => {
            if worksheet_names != &projection.worksheet_names {
                return Err(io::Error::other("source worksheet-name oracle mismatch").into());
            }
        },
        Observation::OneCell { cell } => {
            let actual = cell.as_ref().map(cell_projection);
            if actual != projection.selected_cell {
                return Err(io::Error::other("source cell oracle mismatch").into());
            }
        },
    }
    Ok(())
}

fn validate_eager_observation<R: Read + Seek>(
    workbook: &litchi_xls::Workbook<R>,
    observation: &Observation,
    oracle: &SemanticOracle,
) -> Result<(), AnyError> {
    let projection = &oracle.eager_implementation_projection;
    match observation {
        Observation::Open { worksheet_count } => {
            if *worksheet_count != projection.worksheet_count
                || eager_names(workbook) != projection.worksheet_names
            {
                return Err(io::Error::other("eager semantic oracle mismatch").into());
            }
        },
        Observation::List { worksheet_names } => {
            if worksheet_names != &projection.worksheet_names {
                return Err(io::Error::other("eager worksheet-name oracle mismatch").into());
            }
        },
        Observation::OneCell { cell } => {
            let actual = cell.as_ref().map(cell_projection);
            if actual != projection.selected_cell {
                return Err(io::Error::other("eager cell oracle mismatch").into());
            }
        },
    }
    Ok(())
}

fn validate_facade_parity(
    observation: &Observation,
    source_projection: &SemanticProjection,
) -> Result<(), AnyError> {
    match observation {
        Observation::Open { worksheet_count }
            if *worksheet_count == source_projection.worksheet_count =>
        {
            Ok(())
        },
        Observation::List { worksheet_names }
            if worksheet_names == &source_projection.worksheet_names =>
        {
            Ok(())
        },
        Observation::OneCell { .. } => Err(io::Error::new(
            io::ErrorKind::Unsupported,
            "facade/source parity is defined only for open and list",
        )
        .into()),
        Observation::Open { .. } => {
            Err(io::Error::other("facade/source worksheet-count parity mismatch").into())
        },
        Observation::List { .. } => {
            Err(io::Error::other("facade/source worksheet-name parity mismatch").into())
        },
    }
}

fn duration_ns(duration: std::time::Duration) -> u64 {
    duration.as_nanos().min(u128::from(u64::MAX)) as u64
}

fn run_source_sample(
    mode: Mode,
    input: &InputSnapshot,
    operation: Operation,
    coordinates: Coordinates,
    oracle: &SemanticOracle,
) -> Result<Sample, AnyError> {
    let source_run = make_source(mode, input)?;
    let source_version_before = source_run.source.version()?;
    let baseline_metrics = source_run.metrics.snapshot();
    let started = Instant::now();
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(Arc::clone(&source_run.source))?;
    let observation = source_operation(&workbook, operation, coordinates)?;
    let elapsed = duration_ns(started.elapsed());
    let metrics = source_run.metrics.snapshot().delta(baseline_metrics);
    let source_version_after = source_run.source.version()?;
    if source_version_before != source_version_after {
        return Err(io::Error::other("source version changed during timed operation").into());
    }
    validate_source_observation(&observation, oracle)?;
    Ok(Sample {
        elapsed_ns: elapsed,
        observation: observation.into_json(),
        metrics,
        source_version_stable: Some(source_version_before == source_version_after),
        eager_phases: None,
    })
}

fn run_eager_sample(
    input: &InputSnapshot,
    operation: Operation,
    coordinates: Coordinates,
    oracle: &SemanticOracle,
) -> Result<Sample, AnyError> {
    let metrics = Arc::new(Metrics::default());
    let total_started = Instant::now();
    let cfb_started = Instant::now();
    let reader = CountingFile {
        file: File::open(&input.staged_path)?,
        metrics: Arc::clone(&metrics),
    };
    let ole = OleFile::open(reader)?;
    let cfb_open_ns = duration_ns(cfb_started.elapsed());
    let cfb_metrics = metrics.snapshot();

    let owner_started = Instant::now();
    let workbook = litchi_xls::Workbook::from_ole_file(ole)?;
    let xls_owner_ns = duration_ns(owner_started.elapsed());
    let owner_metrics = metrics.snapshot();

    let query_started = Instant::now();
    let observation = eager_operation(&workbook, operation, coordinates)?;
    let selected_query_ns = duration_ns(query_started.elapsed());
    let query_metrics = metrics.snapshot();
    let elapsed = duration_ns(total_started.elapsed());
    validate_eager_observation(&workbook, &observation, oracle)?;
    Ok(Sample {
        elapsed_ns: elapsed,
        observation: observation.into_json(),
        metrics: query_metrics,
        source_version_stable: None,
        eager_phases: Some(EagerPhaseTiming {
            cfb_open_ns,
            xls_owner_ns,
            selected_query_ns,
            cfb_open_read_calls: cfb_metrics.read_calls,
            cfb_open_read_bytes: cfb_metrics.read_bytes,
            xls_owner_read_calls: owner_metrics
                .read_calls
                .saturating_sub(cfb_metrics.read_calls),
            xls_owner_read_bytes: owner_metrics
                .read_bytes
                .saturating_sub(cfb_metrics.read_bytes),
            selected_query_read_calls: query_metrics
                .read_calls
                .saturating_sub(owner_metrics.read_calls),
            selected_query_read_bytes: query_metrics
                .read_bytes
                .saturating_sub(owner_metrics.read_bytes),
        }),
    })
}

fn run_facade_sample(
    input: &InputSnapshot,
    operation: Operation,
    _coordinates: Coordinates,
    oracle: &SemanticOracle,
) -> Result<Sample, AnyError> {
    let started = Instant::now();
    let workbook = litchi::sheet::Workbook::open(&input.staged_path)?;
    let observation = match operation {
        Operation::Open => Observation::Open {
            worksheet_count: std::hint::black_box(workbook.worksheet_count()?),
        },
        Operation::List => Observation::List {
            worksheet_names: std::hint::black_box(workbook.worksheet_names()?),
        },
        Operation::OneCell => {
            return Err(io::Error::new(
                io::ErrorKind::Unsupported,
                "facade-file does not expose selected-cell access",
            )
            .into());
        },
    };
    let elapsed = duration_ns(started.elapsed());
    validate_facade_parity(&observation, &oracle.source_implementation_projection)?;
    Ok(Sample {
        elapsed_ns: elapsed,
        observation: observation.into_json(),
        metrics: MetricsSnapshot::default(),
        source_version_stable: None,
        eager_phases: None,
    })
}

fn run_one_sample(
    config: &Config,
    input: &InputSnapshot,
    oracle: &SemanticOracle,
) -> Result<Sample, AnyError> {
    let coordinates = Coordinates {
        worksheet_index: config.worksheet_index,
        row: config.row,
        column: config.column,
    };
    match config.mode {
        Mode::OwnedReadAt | Mode::AtomicFile | Mode::TrackedFile | Mode::FileSource => {
            run_source_sample(config.mode, input, config.operation, coordinates, oracle)
        },
        Mode::EagerFile => run_eager_sample(input, config.operation, coordinates, oracle),
        Mode::FacadeFile => run_facade_sample(input, config.operation, coordinates, oracle),
    }
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn remove_staged_file(path: &Path) {
    if fs::remove_file(path).is_ok() {
        return;
    }
    #[cfg(windows)]
    if let Ok(metadata) = fs::metadata(path) {
        let mut permissions = metadata.permissions();
        permissions.set_readonly(false);
        let _ = fs::set_permissions(path, permissions);
        let _ = fs::remove_file(path);
    }
}

fn stage_input(bytes: &[u8]) -> Result<PathBuf, AnyError> {
    const ATTEMPTS: usize = 64;
    for _ in 0..ATTEMPTS {
        let id = NEXT_STAGED_INPUT_ID.fetch_add(1, Ordering::Relaxed);
        let path = env::temp_dir().join(format!(
            "litchi-xls-source-attribution-{}-{id}.xls",
            std::process::id()
        ));
        let mut options = OpenOptions::new();
        options.write(true).create_new(true);
        #[cfg(unix)]
        options.mode(0o600);
        let mut file = match options.open(&path) {
            Ok(file) => file,
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        };
        let staged = (|| -> io::Result<()> {
            file.write_all(bytes)?;
            file.flush()?;
            file.sync_data()?;
            let mut permissions = file.metadata()?.permissions();
            permissions.set_readonly(true);
            drop(file);
            fs::set_permissions(&path, permissions)
        })();
        if let Err(error) = staged {
            remove_staged_file(&path);
            return Err(error.into());
        }
        return Ok(path);
    }
    Err(io::Error::new(
        io::ErrorKind::AlreadyExists,
        "could not allocate a unique staged input path",
    )
    .into())
}

fn verify_staged_input(input: &InputSnapshot) -> Result<(), AnyError> {
    let bytes = fs::read(&input.staged_path)?;
    if bytes.len() != input.bytes.len() || sha256_hex(&bytes) != input.sha256 {
        return Err(io::Error::other("staged input identity changed during the run").into());
    }
    Ok(())
}

fn read_input(path: &Path) -> Result<InputSnapshot, AnyError> {
    let canonical = fs::canonicalize(path)?;
    let bytes = Arc::new(fs::read(&canonical)?);
    let sha256 = sha256_hex(&bytes);
    let staged_path = stage_input(&bytes)?;
    Ok(InputSnapshot {
        path: canonical,
        bytes,
        sha256,
        staged_path,
    })
}

fn identity(path: &Path, bytes: &[u8], sha256: String) -> Result<Identity, AnyError> {
    Ok(Identity {
        path: path.display().to_string(),
        bytes: u64::try_from(bytes.len())?,
        sha256,
    })
}

fn binary_identity() -> Result<Identity, AnyError> {
    let path = env::current_exe()?;
    let bytes = fs::read(&path)?;
    identity(&path, &bytes, sha256_hex(&bytes))
}

fn revision() -> String {
    env::var("LITCHI_REVISION").unwrap_or_else(|_| "unknown".to_owned())
}

fn usage() -> &'static str {
    "usage: xls_source_attribution --input PATH [--mode owned-readat|atomic-file|tracked-file|file-source|eager-file|facade-file] [--operation open|list|one-cell] [--warmups N] [--samples N] [--worksheet-index N] [--row N] [--column N]"
}

fn parse_usize(value: Option<&str>, option: &str) -> Result<usize, String> {
    value
        .ok_or_else(|| format!("{option} requires a value"))?
        .parse()
        .map_err(|_| format!("{option} requires a non-negative integer"))
}

fn parse_u32(value: Option<&str>, option: &str) -> Result<u32, String> {
    value
        .ok_or_else(|| format!("{option} requires a value"))?
        .parse()
        .map_err(|_| format!("{option} requires a u32"))
}

fn parse_args_from<I>(arguments: I) -> Result<Option<Config>, String>
where
    I: IntoIterator<Item = String>,
{
    let mut input = None;
    let mut operation = Operation::Open;
    let mut mode = Mode::OwnedReadAt;
    let mut warmups = DEFAULT_WARMUPS;
    let mut samples = DEFAULT_SAMPLES;
    let mut worksheet_index = DEFAULT_WORKSHEET_INDEX;
    let mut row = DEFAULT_ROW;
    let mut column = DEFAULT_COLUMN;
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => return Ok(None),
            "--input" => {
                input = Some(PathBuf::from(
                    arguments.next().ok_or("--input requires PATH")?,
                ))
            },
            "--operation" => {
                operation = match arguments
                    .next()
                    .ok_or("--operation requires a value")?
                    .as_str()
                {
                    "open" => Operation::Open,
                    "list" => Operation::List,
                    "one-cell" => Operation::OneCell,
                    value => return Err(format!("unknown operation {value:?}")),
                };
            },
            "--mode" => {
                mode = match arguments.next().ok_or("--mode requires a value")?.as_str() {
                    "owned-readat" => Mode::OwnedReadAt,
                    "atomic-file" => Mode::AtomicFile,
                    "tracked-file" => Mode::TrackedFile,
                    "file-source" => Mode::FileSource,
                    "eager-file" => Mode::EagerFile,
                    "facade-file" => Mode::FacadeFile,
                    value => return Err(format!("unknown mode {value:?}")),
                };
            },
            "--warmups" | "--warmup" => {
                warmups = parse_usize(arguments.next().as_deref(), "--warmups")?;
            },
            "--samples" => {
                samples = parse_usize(arguments.next().as_deref(), "--samples")?;
            },
            "--worksheet-index" => {
                worksheet_index = parse_usize(arguments.next().as_deref(), "--worksheet-index")?;
            },
            "--row" => row = parse_u32(arguments.next().as_deref(), "--row")?,
            "--column" => column = parse_u32(arguments.next().as_deref(), "--column")?,
            value if !value.starts_with('-') && input.is_none() => {
                input = Some(PathBuf::from(value))
            },
            value => return Err(format!("unrecognized argument {value:?}; use --help")),
        }
    }
    if warmups == 0 || samples == 0 {
        return Err("--warmups and --samples must be greater than zero".into());
    }
    Ok(Some(Config {
        input: input.ok_or("--input PATH is required")?,
        operation,
        mode,
        warmups,
        samples,
        worksheet_index,
        row,
        column,
    }))
}

fn parse_args() -> Result<Option<Config>, AnyError> {
    parse_args_from(env::args().skip(1))
        .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error).into())
}

fn allocate_records(warmups: usize, samples: usize) -> Result<(usize, Vec<Sample>), AnyError> {
    let total_iterations = warmups.checked_add(samples).ok_or_else(|| {
        io::Error::new(io::ErrorKind::InvalidInput, "warmup/sample count overflow")
    })?;
    let mut records = Vec::new();
    records.try_reserve_exact(samples).map_err(|error| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("sample count cannot be allocated: {error}"),
        )
    })?;
    Ok((total_iterations, records))
}

fn run(config: Config) -> Result<Report, AnyError> {
    let input = read_input(&config.input)?;
    let coordinates = Coordinates {
        worksheet_index: config.worksheet_index,
        row: config.row,
        column: config.column,
    };
    let oracle = build_oracle(&input.bytes, config.operation, coordinates)?;
    let (total_iterations, mut records) = allocate_records(config.warmups, config.samples)?;
    for iteration in 0..total_iterations {
        let sample = run_one_sample(&config, &input, &oracle)?;
        if iteration >= config.warmups {
            records.push(sample);
        }
    }
    verify_staged_input(&input)?;
    let elapsed_samples_ns = records.iter().map(|sample| sample.elapsed_ns).collect();
    let input_identity = identity(&input.path, &input.bytes, input.sha256.clone())?;
    Ok(Report {
        schema_version: 2,
        revision: revision(),
        tool: ToolIdentity {
            name: env!("CARGO_PKG_NAME"),
            version: env!("CARGO_PKG_VERSION"),
            revision: revision(),
        },
        binary: binary_identity()?,
        input: input_identity,
        mode: config.mode.as_str(),
        operation: config.operation.as_str(),
        counter_scope: config.mode.counter_scope(),
        limitation: config.mode.limitation(config.operation),
        timing_scope: config.mode.timing_scope(),
        process_scope: PROCESS_SCOPE,
        input_scope: INPUT_SCOPE,
        warmups: config.warmups,
        samples: config.samples,
        worksheet_index: config.worksheet_index,
        row: config.row,
        column: config.column,
        semantic_oracle: oracle,
        elapsed_samples_ns,
        records,
    })
}

fn main() -> Result<(), AnyError> {
    let Some(config) = parse_args()? else {
        println!("{}", usage());
        return Ok(());
    };
    let report = run(config)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_mode_operation_and_coordinates() {
        let config = parse_args_from([
            "--input".to_owned(),
            "book.xls".to_owned(),
            "--mode".to_owned(),
            "tracked-file".to_owned(),
            "--operation".to_owned(),
            "one-cell".to_owned(),
            "--warmups".to_owned(),
            "2".to_owned(),
            "--samples".to_owned(),
            "4".to_owned(),
            "--worksheet-index".to_owned(),
            "3".to_owned(),
            "--row".to_owned(),
            "7".to_owned(),
            "--column".to_owned(),
            "9".to_owned(),
        ])
        .unwrap()
        .unwrap();
        assert_eq!(config.mode, Mode::TrackedFile);
        assert_eq!(config.operation, Operation::OneCell);
        assert_eq!(config.warmups, 2);
        assert_eq!(config.samples, 4);
        assert_eq!(config.worksheet_index, 3);
        assert_eq!(config.row, 7);
        assert_eq!(config.column, 9);
    }

    #[test]
    fn range_union_reports_unique_bytes() {
        let mut union = RangeUnion::default();
        assert_eq!(union.insert(10..20), 10);
        assert_eq!(union.insert(15..25), 15);
        assert_eq!(union.insert(0..5), 20);
        assert_eq!(union.insert(5..10), 25);
    }

    #[test]
    fn cell_projection_and_mode_limits_are_explicit() {
        assert_eq!(
            cell_projection(&CellValue::Float(42.0)),
            "float:4045000000000000"
        );
        assert!(Mode::FacadeFile.limitation(Operation::OneCell).is_some());
        assert!(Mode::TrackedFile.counter_scope().contains("range-union"));
    }

    #[test]
    fn metrics_delta_excludes_untimed_oracle_work() {
        let baseline = MetricsSnapshot {
            version_calls: 2,
            version_ns: 11,
            ..MetricsSnapshot::default()
        };
        let observed = MetricsSnapshot {
            read_calls: 3,
            read_bytes: 9,
            version_calls: 7,
            version_ns: 31,
            ..MetricsSnapshot::default()
        };
        let delta = observed.delta(baseline);
        assert_eq!(delta.read_calls, 3);
        assert_eq!(delta.read_bytes, 9);
        assert_eq!(delta.version_calls, 5);
        assert_eq!(delta.version_ns, 20);
    }

    #[test]
    fn zero_warmups_or_samples_are_rejected() {
        for option in ["--warmups", "--samples"] {
            let result = parse_args_from([
                "--input".to_owned(),
                "book.xls".to_owned(),
                option.to_owned(),
                "0".to_owned(),
            ]);
            assert!(result.is_err());
        }
    }

    #[test]
    fn staged_input_is_identity_bound_read_only_and_removed_on_drop() {
        let id = NEXT_STAGED_INPUT_ID.fetch_add(1, Ordering::Relaxed);
        let original = env::temp_dir().join(format!(
            "litchi-xls-source-attribution-test-{}-{id}.xls",
            std::process::id()
        ));
        fs::write(&original, b"staged identity").unwrap();
        let snapshot = read_input(&original).unwrap();
        let staged_path = snapshot.staged_path.clone();
        assert_eq!(fs::read(&staged_path).unwrap(), b"staged identity");
        assert!(fs::metadata(&staged_path).unwrap().permissions().readonly());
        verify_staged_input(&snapshot).unwrap();
        drop(snapshot);
        assert!(!staged_path.exists());
        fs::remove_file(original).unwrap();
    }

    #[test]
    fn timing_scopes_for_incomparable_families_are_distinct() {
        assert_eq!(
            Mode::OwnedReadAt.timing_scope(),
            Mode::FileSource.timing_scope()
        );
        assert_ne!(
            Mode::OwnedReadAt.timing_scope(),
            Mode::EagerFile.timing_scope()
        );
        assert_ne!(
            Mode::EagerFile.timing_scope(),
            Mode::FacadeFile.timing_scope()
        );
    }

    #[test]
    fn facade_parity_is_limited_to_source_open_and_list() {
        let projection = SemanticProjection {
            worksheet_count: 2,
            worksheet_names: vec!["One".to_owned(), "Two".to_owned()],
            selected_cell: Some("string:4:Date".to_owned()),
        };
        assert!(
            validate_facade_parity(&Observation::Open { worksheet_count: 2 }, &projection,).is_ok()
        );
        assert!(
            validate_facade_parity(
                &Observation::List {
                    worksheet_names: vec!["One".to_owned(), "Two".to_owned()],
                },
                &projection,
            )
            .is_ok()
        );
        assert!(validate_facade_parity(&Observation::OneCell { cell: None }, &projection).is_err());
    }

    #[test]
    fn sample_iteration_overflow_is_rejected_before_allocation() {
        let error = allocate_records(1, usize::MAX).unwrap_err();
        assert!(error.to_string().contains("overflow"));
    }
}
