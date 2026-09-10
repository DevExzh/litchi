//! Profile source-backed XLSB indexed cached-value reads against full worksheet
//! materialization.
//!
//! This is intentionally a standalone example so it can be built with the
//! `litchi-xlsb` package without the umbrella facade.  One invocation measures
//! one backend/case in one process.  The final matrix script runs each case in
//! an independent process and retains this program's JSON as the raw sample
//! evidence.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    clippy::similar_names,
    reason = "this opt-in profile owns a process allocator wrapper and deliberately emits machine-readable stdout"
)]

use litchi_core::sheet::{
    Cell as SheetCell, CellValue, WorkbookTrait, Worksheet as WorksheetTrait,
};
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_opc::{SourceCacheCounterDelta, SourceCacheDiagnostics};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi_xlsb::{SourceBackedWorkbook, Workbook};
use std::alloc::{GlobalAlloc, Layout, System};
use std::collections::BTreeSet;
use std::env;
use std::error::Error;
use std::fmt::Write as FmtWrite;
use std::fs;
use std::io::{self, Cursor};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Instant;

type BoxError = Box<dyn Error + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const DEFAULT_FIXTURE: &str = "test-data/poi/test-data/spreadsheet/testVarious.xlsb";
const DEFAULT_WARMUP: usize = 3;
const DEFAULT_SAMPLES: usize = 30;
const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

// SAFETY: every method delegates allocation ownership and layout handling to
// `std::alloc::System`; the atomic counters only observe completed calls.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the pointer/layout pair is owned by the caller and is passed
        // unchanged to the system allocator.
        unsafe { System.dealloc(pointer, layout) };
        ALLOC_DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_DEALLOCATED_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
        subtract_live(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller supplies the valid pointer/layout contract for a
        // system reallocation.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            ALLOC_REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            ALLOC_ALLOCATED_BYTES.fetch_add(as_u64(new_size), Ordering::Relaxed);
            ALLOC_DEALLOCATED_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
            replace_live(layout.size(), new_size);
        }
        result
    }
}

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static ALLOC_ALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_DEALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_INVALID: AtomicBool = AtomicBool::new(false);

fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn observe_alloc(size: usize) {
    let size = as_u64(size);
    ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
    ALLOC_ALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
    let live = ALLOC_LIVE_BYTES
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    update_peak(live);
}

fn subtract_live(size: usize) {
    let size = as_u64(size);
    let before = ALLOC_LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
    if before < size {
        ALLOC_INVALID.store(true, Ordering::Release);
    }
}

fn replace_live(old_size: usize, new_size: usize) {
    subtract_live(old_size);
    let size = as_u64(new_size);
    let live = ALLOC_LIVE_BYTES
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    update_peak(live);
}

fn update_peak(live: u64) {
    let mut peak = ALLOC_PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    while live > peak {
        match ALLOC_PEAK_LIVE_BYTES.compare_exchange_weak(
            peak,
            live,
            Ordering::Relaxed,
            Ordering::Relaxed,
        ) {
            Ok(_) => break,
            Err(observed) => peak = observed,
        }
    }
}

/// Reset the requested-live highwater to the current live baseline.
///
/// The harness is single-threaded for each case. Resetting immediately before
/// an interval keeps eager corpus construction and warmup allocations out of
/// that interval's peak; the requested-live counter remains distinct from an
/// allocator-internal highwater or process RSS measurement.
fn reset_peak_to_live() {
    ALLOC_PEAK_LIVE_BYTES.store(ALLOC_LIVE_BYTES.load(Ordering::Acquire), Ordering::Release);
}

#[derive(Clone, Copy)]
struct AllocationSnapshot {
    calls: u64,
    dealloc_calls: u64,
    realloc_calls: u64,
    failed: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_bytes: u64,
    peak_live_bytes: u64,
    invalid: bool,
}

impl AllocationSnapshot {
    fn now() -> Self {
        Self {
            calls: ALLOC_CALLS.load(Ordering::Acquire),
            dealloc_calls: ALLOC_DEALLOC_CALLS.load(Ordering::Acquire),
            realloc_calls: ALLOC_REALLOC_CALLS.load(Ordering::Acquire),
            failed: ALLOC_FAILED.load(Ordering::Acquire),
            allocated_bytes: ALLOC_ALLOCATED_BYTES.load(Ordering::Acquire),
            deallocated_bytes: ALLOC_DEALLOCATED_BYTES.load(Ordering::Acquire),
            live_bytes: ALLOC_LIVE_BYTES.load(Ordering::Acquire),
            peak_live_bytes: ALLOC_PEAK_LIVE_BYTES.load(Ordering::Acquire),
            invalid: ALLOC_INVALID.load(Ordering::Acquire),
        }
    }

    fn delta(self, after: Self) -> AllocationDelta {
        AllocationDelta {
            calls: after.calls.saturating_sub(self.calls),
            dealloc_calls: after.dealloc_calls.saturating_sub(self.dealloc_calls),
            realloc_calls: after.realloc_calls.saturating_sub(self.realloc_calls),
            failed: after.failed.saturating_sub(self.failed),
            allocated_bytes: after.allocated_bytes.saturating_sub(self.allocated_bytes),
            deallocated_bytes: after
                .deallocated_bytes
                .saturating_sub(self.deallocated_bytes),
            live_before: self.live_bytes,
            live_after: after.live_bytes,
            peak_before: self.peak_live_bytes,
            peak_after: after.peak_live_bytes,
            invalid: self.invalid || after.invalid,
        }
    }
}

#[derive(Clone, Copy)]
struct AllocationDelta {
    calls: u64,
    dealloc_calls: u64,
    realloc_calls: u64,
    failed: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_before: u64,
    live_after: u64,
    peak_before: u64,
    peak_after: u64,
    invalid: bool,
}

#[derive(Default)]
struct ReadCounters {
    calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
}

struct CountingReadAt {
    inner: OwnedSource,
    counters: Arc<ReadCounters>,
}

impl CountingReadAt {
    fn new(bytes: Vec<u8>, counters: Arc<ReadCounters>) -> Self {
        Self {
            inner: OwnedSource::new(bytes),
            counters,
        }
    }
}

impl ReadAt for CountingReadAt {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.counters.calls.fetch_add(1, Ordering::Relaxed);
        self.counters
            .requested_bytes
            .fetch_add(as_u64(output.len()), Ordering::Relaxed);
        let read = self.inner.read_at(offset, output)?;
        self.counters
            .returned_bytes
            .fetch_add(as_u64(read), Ordering::Relaxed);
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
struct ReadSnapshot {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
}

impl ReadSnapshot {
    fn now(counters: &ReadCounters) -> Self {
        Self {
            calls: counters.calls.load(Ordering::Acquire),
            requested_bytes: counters.requested_bytes.load(Ordering::Acquire),
            returned_bytes: counters.returned_bytes.load(Ordering::Acquire),
        }
    }

    fn delta(self, after: Self) -> Self {
        Self {
            calls: after.calls.saturating_sub(self.calls),
            requested_bytes: after.requested_bytes.saturating_sub(self.requested_bytes),
            returned_bytes: after.returned_bytes.saturating_sub(self.returned_bytes),
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
struct CacheDelta {
    hits: u64,
    cold_loads: u64,
    waiter_joins: u64,
    successful_loads: u64,
    failed_loads: u64,
    evictions: u64,
    bypasses: u64,
    oversized_bypasses: u64,
    allocation_bypasses: u64,
    budget_reservation_failures: u64,
}

impl CacheDelta {
    fn from(before: SourceCacheDiagnostics, after: SourceCacheDiagnostics) -> Result<Self> {
        let delta = SourceCacheDiagnostics::checked_counter_delta(before, after)?;
        Ok(Self::from_counter_delta(delta))
    }

    fn from_counter_delta(delta: SourceCacheCounterDelta) -> Self {
        Self {
            hits: delta.hits,
            cold_loads: delta.cold_loads,
            waiter_joins: delta.waiter_joins,
            successful_loads: delta.successful_loads,
            failed_loads: delta.failed_loads,
            evictions: delta.evictions,
            bypasses: delta.bypasses,
            oversized_bypasses: delta.oversized_bypasses,
            allocation_bypasses: delta.allocation_bypasses,
            budget_reservation_failures: delta.budget_reservation_failures,
        }
    }
}

#[derive(Clone)]
struct Target {
    row: u32,
    column: u32,
    expected: Option<CellValue>,
}

#[derive(Clone)]
struct Corpus {
    path: PathBuf,
    bytes: Vec<u8>,
    sheet: usize,
    targets: Vec<Target>,
    cell_count: usize,
    dimensions: Option<(u32, u32, u32, u32)>,
}

#[derive(Clone, Copy)]
enum Case {
    IndexedCold,
    MaterializeCold,
    IndexedWarm,
    MaterializeWarm,
}

impl Case {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "indexed_cold" => Ok(Self::IndexedCold),
            "materialize_cold" => Ok(Self::MaterializeCold),
            "indexed_warm" => Ok(Self::IndexedWarm),
            "materialize_warm" => Ok(Self::MaterializeWarm),
            _ => Err(format!(
                "unknown case {value:?}; expected indexed_cold, materialize_cold, indexed_warm, or materialize_warm"
            )
            .into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::IndexedCold => "indexed_cold",
            Self::MaterializeCold => "materialize_cold",
            Self::IndexedWarm => "indexed_warm",
            Self::MaterializeWarm => "materialize_warm",
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::IndexedCold => {
                "source-backed open + indexed_values preparation + cached_value lookups + source-backed teardown"
            },
            Self::MaterializeCold => {
                "source-backed open + complete worksheet materialization + cell lookups + source-backed teardown"
            },
            Self::IndexedWarm => "cached_value lookups on one retained indexed handle",
            Self::MaterializeWarm => "cell lookups on one retained complete Worksheet",
        }
    }

    const fn indexed(self) -> bool {
        matches!(self, Self::IndexedCold | Self::IndexedWarm)
    }
}

struct Args {
    case: Case,
    fixture: PathBuf,
    sheet: usize,
    warmup: usize,
    samples: usize,
}

impl Args {
    fn parse() -> Result<Self> {
        let mut case = None;
        let mut fixture = PathBuf::from(DEFAULT_FIXTURE);
        let mut sheet = 0usize;
        let mut warmup = DEFAULT_WARMUP;
        let mut samples = DEFAULT_SAMPLES;
        let mut arguments = env::args().skip(1);
        while let Some(argument) = arguments.next() {
            match argument.as_str() {
                "--case" => case = Some(Case::parse(&next_arg(&mut arguments, "--case")?)?),
                "--fixture" => fixture = PathBuf::from(next_arg(&mut arguments, "--fixture")?),
                "--sheet" => {
                    sheet = parse_positive(&next_arg(&mut arguments, "--sheet")?, "--sheet")?
                },
                "--warmup" => {
                    warmup = parse_positive(&next_arg(&mut arguments, "--warmup")?, "--warmup")?
                },
                "--samples" => {
                    samples = parse_positive(&next_arg(&mut arguments, "--samples")?, "--samples")?
                },
                "--help" | "-h" => return Err(Usage.into()),
                other => return Err(format!("unknown argument {other:?}\n\n{Usage}").into()),
            }
        }
        Ok(Self {
            case: case.ok_or_else(|| format!("--case is required\n\n{Usage}"))?,
            fixture,
            sheet,
            warmup,
            samples,
        })
    }
}

fn next_arg<I>(arguments: &mut I, name: &str) -> Result<String>
where
    I: Iterator<Item = String>,
{
    arguments
        .next()
        .ok_or_else(|| format!("missing value for {name}").into())
}

fn parse_positive(value: &str, name: &str) -> Result<usize> {
    let parsed = value.parse::<usize>()?;
    if parsed == 0 {
        return Err(format!("{name} must be positive").into());
    }
    Ok(parsed)
}

struct Usage;

impl std::fmt::Display for Usage {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(
            "usage: binary_index_profile --case <indexed_cold|materialize_cold|indexed_warm|materialize_warm> [--fixture PATH] [--sheet N] [--warmup N] [--samples N]",
        )
    }
}

impl std::fmt::Debug for Usage {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        std::fmt::Display::fmt(self, formatter)
    }
}

impl Error for Usage {}

fn main() -> Result<()> {
    let args = Args::parse()?;
    let corpus = Corpus::load(&args.fixture, args.sheet)?;
    let report = match args.case {
        Case::IndexedCold | Case::MaterializeCold => {
            run_cold(&corpus, args.case, args.warmup, args.samples)?
        },
        Case::IndexedWarm | Case::MaterializeWarm => {
            run_warm(&corpus, args.case, args.warmup, args.samples)?
        },
    };
    println!("{}", report.to_json());
    Ok(())
}

impl Corpus {
    fn load(path: &Path, sheet_index: usize) -> Result<Self> {
        let bytes = if let Some(value) = path
            .to_str()
            .and_then(|value| value.strip_prefix("synthetic:"))
        {
            generate_fixture(value)?
        } else {
            fs::read(path)?
        };
        let workbook = Workbook::new(Cursor::new(bytes.clone()))?;
        if sheet_index >= workbook.worksheet_count() {
            return Err(format!(
                "worksheet index {sheet_index} is outside workbook count {}",
                workbook.worksheet_count()
            )
            .into());
        }
        let worksheet = workbook.worksheet(sheet_index)?;
        let dimensions = worksheet.dimensions();
        let mut cells = Vec::new();
        let mut iterator = WorksheetTrait::cells(&worksheet);
        while let Some(cell) = iterator.next() {
            let cell = cell?;
            cells.push((cell.row(), cell.column(), canonical_value(cell.value())));
        }
        cells.sort_by_key(|(row, column, _)| (*row, *column));
        let coordinates = cells
            .iter()
            .map(|(row, column, _)| (*row, *column))
            .collect::<BTreeSet<_>>();
        let mut targets = Vec::new();
        for index in [
            0,
            cells.len().saturating_div(2),
            cells.len().saturating_sub(1),
        ] {
            if let Some((row, column, expected)) = cells.get(index) {
                add_target(
                    &mut targets,
                    &coordinates,
                    *row,
                    *column,
                    Some(expected.clone()),
                );
            }
        }
        if let Some((row, _, _)) = cells.first() {
            let missing_column = (0..=MAX_COLUMN)
                .find(|column| !coordinates.contains(&(*row, *column)))
                .ok_or("could not find a missing column in the first populated row")?;
            add_target(&mut targets, &coordinates, *row, missing_column, None);
        }
        if let Some((min_row, _, max_row, _)) = dimensions {
            if let Some(row) = (min_row..=max_row)
                .find(|row| !coordinates.iter().any(|(cell_row, _)| cell_row == row))
            {
                add_target(&mut targets, &coordinates, row, 0, None);
            }
        }
        let missing_row = dimensions
            .map(|(_, _, max_row, _)| max_row.saturating_add(1))
            .filter(|row| *row <= MAX_ROW)
            .unwrap_or(0);
        add_target(&mut targets, &coordinates, missing_row, 0, None);
        if targets.is_empty() {
            targets.push(Target {
                row: 0,
                column: 0,
                expected: None,
            });
        }
        Ok(Self {
            path: path.to_path_buf(),
            bytes,
            sheet: sheet_index,
            targets,
            cell_count: cells.len(),
            dimensions,
        })
    }
}

fn generate_fixture(value: &str) -> Result<Vec<u8>> {
    let cell_count = value.parse::<usize>()?;
    if cell_count == 0 || cell_count > 1_000_000 {
        return Err("synthetic fixture cell count must be in 1..=1_000_000".into());
    }
    let columns_per_row = 32_u32;
    let mut worksheet = MutableWorksheet::new("Synthetic");
    for index in 0..cell_count {
        let row = u32::try_from(index / columns_per_row as usize)
            .map_err(|_| "synthetic fixture row exceeds XLSB bounds")?;
        let column = u32::try_from(index % columns_per_row as usize)
            .map_err(|_| "synthetic fixture column exceeds XLSB bounds")?;
        worksheet.set_cell(row, column, index as f64);
    }
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(worksheet);
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output)?;
    Ok(output.into_inner())
}

fn add_target(
    targets: &mut Vec<Target>,
    coordinates: &BTreeSet<(u32, u32)>,
    row: u32,
    column: u32,
    expected: Option<CellValue>,
) {
    if row > MAX_ROW
        || column > MAX_COLUMN
        || targets
            .iter()
            .any(|target| target.row == row && target.column == column)
    {
        return;
    }
    if expected.is_none() && coordinates.contains(&(row, column)) {
        return;
    }
    targets.push(Target {
        row,
        column,
        expected,
    });
}

fn canonical_value(value: &CellValue) -> CellValue {
    match value {
        CellValue::Formula { cached_value, .. } => cached_value
            .as_deref()
            .map_or(CellValue::Empty, canonical_value),
        CellValue::Int(value) => CellValue::Float(*value as f64),
        CellValue::DateTime(value) => CellValue::Float(*value),
        CellValue::Empty => CellValue::Empty,
        CellValue::Bool(value) => CellValue::Bool(*value),
        CellValue::Float(value) => CellValue::Float(*value),
        CellValue::String(value) => CellValue::String(value.clone()),
        CellValue::Error(value) => CellValue::Error(value.clone()),
    }
}

#[derive(Clone, Copy)]
struct Sample {
    elapsed_ns: u64,
    semantic_ok: bool,
    digest: u64,
    reads_total: ReadSnapshot,
    reads_before: ReadSnapshot,
    reads_operation: ReadSnapshot,
    cache_total: CacheDelta,
    cache_before: CacheDelta,
    cache_operation: CacheDelta,
    retained_entries: usize,
    retained_bytes: usize,
    allocation: AllocationDelta,
}

struct Report {
    case: Case,
    corpus: Corpus,
    warmup: usize,
    samples: Vec<Sample>,
    setup_reads: Option<ReadSnapshot>,
    setup_cache: Option<CacheDelta>,
    setup_allocation: Option<AllocationDelta>,
    setup_retained_entries: Option<usize>,
    setup_retained_bytes: Option<usize>,
}

fn run_cold(corpus: &Corpus, case: Case, warmup: usize, samples: usize) -> Result<Report> {
    for _ in 0..warmup {
        let sample = run_cold_sample(corpus, case)?;
        if !sample.semantic_ok {
            return Err("semantic oracle failed during warmup".into());
        }
    }
    let mut results = Vec::with_capacity(samples);
    for _ in 0..samples {
        let sample = run_cold_sample(corpus, case)?;
        if !sample.semantic_ok {
            return Err("semantic oracle failed during sample".into());
        }
        results.push(sample);
    }
    Ok(Report {
        case,
        corpus: corpus.clone(),
        warmup,
        samples: results,
        setup_reads: None,
        setup_cache: None,
        setup_allocation: None,
        setup_retained_entries: None,
        setup_retained_bytes: None,
    })
}

fn run_cold_sample(corpus: &Corpus, case: Case) -> Result<Sample> {
    let counters = Arc::new(ReadCounters::default());
    let source: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(
        corpus.bytes.clone(),
        Arc::clone(&counters),
    ));
    // Keep the caller-owned source allocation alive through the post-teardown
    // snapshot. The source wrapper is setup outside the timed API call; an
    // extra Arc owner avoids counting its release as a negative operation
    // live-byte delta.
    let source_guard = Arc::clone(&source);
    reset_peak_to_live();
    let allocation_before = AllocationSnapshot::now();
    let started = Instant::now();
    let workbook = SourceBackedWorkbook::from_read_at(source)?;
    let reads_open = ReadSnapshot::now(&counters);
    let cache_open_snapshot = workbook.cache_diagnostics();
    let (semantic_ok, digest) = run_lookup(&workbook, corpus, case)?;
    let reads_total = ReadSnapshot::now(&counters);
    let cache_total_snapshot = workbook.cache_diagnostics();
    drop(workbook);
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
    let allocation = allocation_before.delta(AllocationSnapshot::now());
    drop(source_guard);
    let reads_operation = reads_open.delta(reads_total);
    let cache_open = CacheDelta::from(SourceCacheDiagnostics::default(), cache_open_snapshot)?;
    let cache_total = CacheDelta::from(SourceCacheDiagnostics::default(), cache_total_snapshot)?;
    let cache_operation = CacheDelta::from(cache_open_snapshot, cache_total_snapshot)?;
    Ok(Sample {
        elapsed_ns,
        semantic_ok,
        digest,
        reads_total,
        reads_before: reads_open,
        reads_operation,
        cache_total,
        cache_before: cache_open,
        cache_operation,
        retained_entries: cache_total_snapshot.retained_entries,
        retained_bytes: cache_total_snapshot.retained_bytes,
        allocation,
    })
}

fn run_warm(corpus: &Corpus, case: Case, warmup: usize, samples: usize) -> Result<Report> {
    let counters = Arc::new(ReadCounters::default());
    let source: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(
        corpus.bytes.clone(),
        Arc::clone(&counters),
    ));
    let workbook = SourceBackedWorkbook::from_read_at(source)?;
    let worksheet = workbook
        .worksheet_by_index(corpus.sheet)?
        .ok_or("selected worksheet is absent")?;
    let setup_reads_before = ReadSnapshot::now(&counters);
    let setup_cache_before = workbook.cache_diagnostics();
    reset_peak_to_live();
    let setup_allocation_before = AllocationSnapshot::now();
    let indexed = if case.indexed() {
        Some(
            worksheet
                .indexed_values()?
                .ok_or("selected worksheet has no binary-index relationship")?,
        )
    } else {
        None
    };
    let materialized = if case.indexed() {
        None
    } else {
        Some(worksheet.materialize()?)
    };
    for _ in 0..warmup {
        let (semantic_ok, _) =
            run_prepared_lookup(indexed.as_ref(), materialized.as_ref(), corpus, case)?;
        if !semantic_ok {
            return Err("semantic oracle failed during warmup".into());
        }
    }
    let setup_reads = setup_reads_before.delta(ReadSnapshot::now(&counters));
    let setup_cache_after = workbook.cache_diagnostics();
    let setup_cache = CacheDelta::from(setup_cache_before, setup_cache_after)?;
    let setup_allocation = setup_allocation_before.delta(AllocationSnapshot::now());
    let setup_retained_entries = setup_cache_after.retained_entries;
    let setup_retained_bytes = setup_cache_after.retained_bytes;
    let mut results = Vec::with_capacity(samples);
    for _ in 0..samples {
        let reads_before = ReadSnapshot::now(&counters);
        let cache_before = workbook.cache_diagnostics();
        reset_peak_to_live();
        let allocation_before = AllocationSnapshot::now();
        let started = Instant::now();
        let (semantic_ok, digest) =
            run_prepared_lookup(indexed.as_ref(), materialized.as_ref(), corpus, case)?;
        let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
        let reads_after = ReadSnapshot::now(&counters);
        let cache_after = workbook.cache_diagnostics();
        let allocation = allocation_before.delta(AllocationSnapshot::now());
        results.push(Sample {
            elapsed_ns,
            semantic_ok,
            digest,
            reads_total: reads_after,
            reads_before,
            reads_operation: reads_before.delta(reads_after),
            cache_total: CacheDelta::from(SourceCacheDiagnostics::default(), cache_after)?,
            cache_before: CacheDelta::from(SourceCacheDiagnostics::default(), cache_before)?,
            cache_operation: CacheDelta::from(cache_before, cache_after)?,
            retained_entries: cache_after.retained_entries,
            retained_bytes: cache_after.retained_bytes,
            allocation,
        });
        if !semantic_ok {
            return Err("semantic oracle failed during sample".into());
        }
    }
    drop(indexed);
    drop(materialized);
    drop(worksheet);
    drop(workbook);
    Ok(Report {
        case,
        corpus: corpus.clone(),
        warmup,
        samples: results,
        setup_reads: Some(setup_reads),
        setup_cache: Some(setup_cache),
        setup_allocation: Some(setup_allocation),
        setup_retained_entries: Some(setup_retained_entries),
        setup_retained_bytes: Some(setup_retained_bytes),
    })
}

fn run_lookup(workbook: &SourceBackedWorkbook, corpus: &Corpus, case: Case) -> Result<(bool, u64)> {
    let worksheet = workbook
        .worksheet_by_index(corpus.sheet)?
        .ok_or("selected worksheet is absent")?;
    if case.indexed() {
        let indexed = worksheet
            .indexed_values()?
            .ok_or("selected worksheet has no binary-index relationship")?;
        run_prepared_lookup(Some(&indexed), None, corpus, case)
    } else {
        let materialized = worksheet.materialize()?;
        run_prepared_lookup(None, Some(&materialized), corpus, case)
    }
}

fn run_prepared_lookup(
    indexed: Option<&litchi_xlsb::SourceBackedIndexedWorksheet>,
    materialized: Option<&litchi_xlsb::Worksheet>,
    corpus: &Corpus,
    case: Case,
) -> Result<(bool, u64)> {
    let mut digest = 0xcbf29ce484222325_u64;
    let mut semantic_ok = true;
    for target in &corpus.targets {
        let actual = if case.indexed() {
            indexed
                .ok_or("indexed lookup handle missing")?
                .cached_value(target.row, target.column)?
        } else {
            materialized
                .ok_or("materialized worksheet missing")?
                .get_cell(target.row, target.column)
                .map(|cell| canonical_value(cell.value()))
        };
        if actual != target.expected {
            semantic_ok = false;
        }
        digest = digest_value(digest, target.row, target.column, actual.as_ref());
    }
    Ok((semantic_ok, digest))
}

fn digest_value(mut digest: u64, row: u32, column: u32, value: Option<&CellValue>) -> u64 {
    digest = digest_bytes(digest, &row.to_le_bytes());
    digest = digest_bytes(digest, &column.to_le_bytes());
    match value {
        None => digest_bytes(digest, &[0]),
        Some(CellValue::Empty) => digest_bytes(digest, &[1]),
        Some(CellValue::Bool(value)) => digest_bytes(digest, &[2, u8::from(*value)]),
        Some(CellValue::Int(value)) => digest_number(digest, 3, value.to_le_bytes()),
        Some(CellValue::Float(value)) | Some(CellValue::DateTime(value)) => {
            digest_number(digest, 4, value.to_bits().to_le_bytes())
        },
        Some(CellValue::String(value)) => {
            let digest = digest_bytes(digest, &[5]);
            digest_bytes(digest, value.as_bytes())
        },
        Some(CellValue::Error(value)) => {
            let digest = digest_bytes(digest, &[6]);
            digest_bytes(digest, value.as_bytes())
        },
        Some(CellValue::Formula { .. }) => digest_bytes(digest, &[7]),
    }
}

fn digest_bytes(mut digest: u64, bytes: &[u8]) -> u64 {
    for byte in bytes {
        digest ^= u64::from(*byte);
        digest = digest.wrapping_mul(0x100000001b3);
    }
    digest
}

fn digest_number<const N: usize>(digest: u64, tag: u8, bytes: [u8; N]) -> u64 {
    let digest = digest_bytes(digest, &[tag]);
    digest_bytes(digest, &bytes)
}

impl Report {
    fn to_json(&self) -> String {
        let mut output = String::new();
        output.push('{');
        json_field(&mut output, "schema", "xlsb-binary-index-profile-v1");
        json_field(&mut output, "case", self.case.name());
        json_field(&mut output, "timing_scope", self.case.timing_scope());
        json_field(
            &mut output,
            "source_scope",
            "logical ReadAt calls/requested/returned bytes over an in-memory OwnedSource; not physical I/O",
        );
        json_field(
            &mut output,
            "cache_scope",
            "SourceCacheDiagnostics event deltas plus retained cache gauges; warm setup includes preparation and warmups; not allocator or RSS",
        );
        json_field(
            &mut output,
            "limits_scope",
            "SourceBackedWorkbook and binary-index default finite limits; no custom limit override",
        );
        json_field(
            &mut output,
            "allocation_scope",
            "requested-live operation interval of a process-global std::alloc::System wrapper; peak is reset to the live baseline immediately before each interval; not production latency, allocator-internal highwater, or RSS",
        );
        json_field(
            &mut output,
            "fixture",
            &self.corpus.path.display().to_string(),
        );
        json_number(&mut output, "input_bytes", self.corpus.bytes.len());
        json_number(
            &mut output,
            "input_digest",
            digest_bytes(0xcbf29ce484222325, &self.corpus.bytes),
        );
        json_number(&mut output, "sheet", self.corpus.sheet);
        json_number(&mut output, "cell_count", self.corpus.cell_count);
        json_optional_dimensions(&mut output, "dimensions", self.corpus.dimensions);
        json_number(&mut output, "warmup", self.warmup);
        json_number(&mut output, "sample_count", self.samples.len());
        json_bool(
            &mut output,
            "semantic_ok",
            self.samples.iter().all(|sample| sample.semantic_ok),
        );
        json_bool(
            &mut output,
            "digest_stable",
            self.samples
                .windows(2)
                .all(|samples| samples[0].digest == samples[1].digest),
        );
        json_bool(
            &mut output,
            "source_observation_stable",
            self.samples.windows(2).all(|samples| {
                samples[0].reads_operation == samples[1].reads_operation
                    && samples[0].cache_operation == samples[1].cache_operation
                    && samples[0].retained_entries == samples[1].retained_entries
                    && samples[0].retained_bytes == samples[1].retained_bytes
            }),
        );
        output.push_str(",\"targets\":[");
        for (index, target) in self.corpus.targets.iter().enumerate() {
            if index != 0 {
                output.push(',');
            }
            output.push('{');
            json_number(&mut output, "row", target.row);
            json_number(&mut output, "column", target.column);
            json_bool(&mut output, "expected_present", target.expected.is_some());
            json_number(
                &mut output,
                "expected_digest",
                digest_value(
                    0xcbf29ce484222325,
                    target.row,
                    target.column,
                    target.expected.as_ref(),
                ),
            );
            output.push('}');
        }
        output.push(']');
        json_number(&mut output, "p50_ns", percentile(&self.samples, 50));
        json_number(&mut output, "p95_ns", percentile(&self.samples, 95));
        json_number(&mut output, "p99_ns", percentile(&self.samples, 99));
        let mean = self
            .samples
            .iter()
            .map(|sample| sample.elapsed_ns as f64)
            .sum::<f64>()
            / self.samples.len() as f64;
        json_float(&mut output, "mean_ns", mean);
        json_optional_read(&mut output, "setup_reads", self.setup_reads);
        json_optional_cache(&mut output, "setup_cache", self.setup_cache);
        json_optional_allocation(&mut output, "setup_allocation", self.setup_allocation);
        json_optional_usize(
            &mut output,
            "setup_retained_entries",
            self.setup_retained_entries,
        );
        json_optional_usize(
            &mut output,
            "setup_retained_bytes",
            self.setup_retained_bytes,
        );
        output.push_str(",\"samples\":[");
        for (index, sample) in self.samples.iter().enumerate() {
            if index != 0 {
                output.push(',');
            }
            sample_json(&mut output, sample);
        }
        output.push_str("]}");
        output
    }
}

fn percentile(samples: &[Sample], percentile: usize) -> u64 {
    let mut values = samples
        .iter()
        .map(|sample| sample.elapsed_ns)
        .collect::<Vec<_>>();
    values.sort_unstable();
    let rank = (values.len() * percentile).div_ceil(100).max(1);
    values[rank.saturating_sub(1).min(values.len().saturating_sub(1))]
}

fn sample_json(output: &mut String, sample: &Sample) {
    output.push('{');
    json_number(output, "elapsed_ns", sample.elapsed_ns);
    json_bool(output, "semantic_ok", sample.semantic_ok);
    json_number(output, "digest", sample.digest);
    json_read(output, "reads_total", sample.reads_total);
    json_read(output, "reads_before", sample.reads_before);
    json_read(output, "reads_operation", sample.reads_operation);
    json_cache(output, "cache_total", sample.cache_total);
    json_cache(output, "cache_before", sample.cache_before);
    json_cache(output, "cache_operation", sample.cache_operation);
    json_number(output, "retained_entries", sample.retained_entries);
    json_number(output, "retained_bytes", sample.retained_bytes);
    json_allocation(output, "allocation", sample.allocation);
    output.push('}');
}

fn json_optional_allocation(output: &mut String, name: &str, value: Option<AllocationDelta>) {
    match value {
        Some(value) => json_allocation(output, name, value),
        None => {
            comma(output);
            write!(output, "\"{name}\":null").expect("null write");
        },
    }
}

fn json_allocation(output: &mut String, name: &str, value: AllocationDelta) {
    comma(output);
    write!(output, "\"{name}\":{{").expect("allocation write");
    allocation_fields(output, value);
    output.push('}');
}

fn allocation_fields(output: &mut String, value: AllocationDelta) {
    json_number(output, "calls", value.calls);
    json_number(output, "dealloc_calls", value.dealloc_calls);
    json_number(output, "realloc_calls", value.realloc_calls);
    json_number(output, "failed", value.failed);
    json_number(output, "allocated_bytes", value.allocated_bytes);
    json_number(output, "deallocated_bytes", value.deallocated_bytes);
    json_number(output, "live_before", value.live_before);
    json_number(output, "live_after", value.live_after);
    json_number(output, "peak_before", value.peak_before);
    json_number(output, "peak_after", value.peak_after);
    json_bool(output, "invalid", value.invalid);
}

fn json_optional_usize(output: &mut String, name: &str, value: Option<usize>) {
    comma(output);
    match value {
        Some(value) => write!(output, "\"{name}\":{value}").expect("number write"),
        None => write!(output, "\"{name}\":null").expect("null write"),
    }
}

fn json_field(output: &mut String, name: &str, value: &str) {
    comma(output);
    write!(output, "\"{}\":\"{}\"", name, escape(value)).expect("string write");
}

fn json_number<T: std::fmt::Display>(output: &mut String, name: &str, value: T) {
    comma(output);
    write!(output, "\"{name}\":{value}").expect("number write");
}

fn json_float(output: &mut String, name: &str, value: f64) {
    comma(output);
    write!(output, "\"{name}\":{value:.3}").expect("float write");
}

fn json_bool(output: &mut String, name: &str, value: bool) {
    comma(output);
    write!(output, "\"{name}\":{value}").expect("bool write");
}

fn json_optional_dimensions(
    output: &mut String,
    name: &str,
    dimensions: Option<(u32, u32, u32, u32)>,
) {
    comma(output);
    match dimensions {
        Some((min_row, min_column, max_row, max_column)) => write!(
            output,
            "\"{name}\":[{min_row},{min_column},{max_row},{max_column}]"
        )
        .expect("dimensions write"),
        None => write!(output, "\"{name}\":null").expect("null write"),
    }
}

fn json_optional_read(output: &mut String, name: &str, value: Option<ReadSnapshot>) {
    comma(output);
    match value {
        Some(value) => {
            write!(output, "\"{name}\":{{").expect("read write");
            read_fields(output, value);
            output.push('}');
        },
        None => write!(output, "\"{name}\":null").expect("null write"),
    }
}

fn json_read(output: &mut String, name: &str, value: ReadSnapshot) {
    comma(output);
    write!(output, "\"{name}\":{{").expect("read write");
    read_fields(output, value);
    output.push('}');
}

fn read_fields(output: &mut String, value: ReadSnapshot) {
    json_number(output, "calls", value.calls);
    json_number(output, "requested_bytes", value.requested_bytes);
    json_number(output, "returned_bytes", value.returned_bytes);
}

fn json_optional_cache(output: &mut String, name: &str, value: Option<CacheDelta>) {
    comma(output);
    match value {
        Some(value) => {
            write!(output, "\"{name}\":{{").expect("cache write");
            cache_fields(output, value);
            output.push('}');
        },
        None => write!(output, "\"{name}\":null").expect("null write"),
    }
}

fn json_cache(output: &mut String, name: &str, value: CacheDelta) {
    comma(output);
    write!(output, "\"{name}\":{{").expect("cache write");
    cache_fields(output, value);
    output.push('}');
}

fn cache_fields(output: &mut String, value: CacheDelta) {
    json_number(output, "hits", value.hits);
    json_number(output, "cold_loads", value.cold_loads);
    json_number(output, "waiter_joins", value.waiter_joins);
    json_number(output, "successful_loads", value.successful_loads);
    json_number(output, "failed_loads", value.failed_loads);
    json_number(output, "evictions", value.evictions);
    json_number(output, "bypasses", value.bypasses);
    json_number(output, "oversized_bypasses", value.oversized_bypasses);
    json_number(output, "allocation_bypasses", value.allocation_bypasses);
    json_number(
        output,
        "budget_reservation_failures",
        value.budget_reservation_failures,
    );
}

fn comma(output: &mut String) {
    if !output.is_empty()
        && !output.ends_with('{')
        && !output.ends_with('[')
        && !output.ends_with(',')
    {
        output.push(',');
    }
}

fn escape(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '"' => output.push_str("\\\""),
            '\\' => output.push_str("\\\\"),
            '\n' => output.push_str("\\n"),
            '\r' => output.push_str("\\r"),
            '\t' => output.push_str("\\t"),
            character if character.is_control() => {
                write!(output, "\\u{:04x}", character as u32).expect("escape write");
            },
            character => output.push(character),
        }
    }
    output
}
