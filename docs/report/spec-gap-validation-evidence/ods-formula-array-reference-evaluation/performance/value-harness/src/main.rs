//! Candidate-only value-evaluation benchmark harness.
//!
//! This binary deliberately exercises the public value API with an immutable,
//! in-memory resolver.  It is a measurement fixture, rather than a workbook
//! adapter: resolver setup is reported by the `setup` phase, `parse` measures
//! expression construction, `evaluate` reuses one parsed expression, and
//! `parse-evaluate` includes parsing on every timed call.  The value workload
//! builds its execution context and resolver outside those timed loops.  The
//! provider returns borrowed text and records pointer identity, but its zero
//! copied-byte counter is only a property of this provider.  It is not
//! evidence that every evaluator allocation is zero-copy.
//!
//! The optional worksheet-adapter workload exercises the production
//! `worksheet::formula::Resolver`.  Its `construct` phase measures
//! `worksheet::formula::Resolver::new` separately; its `evaluate` and
//! `parse-evaluate` phases construct one index and reuse it for all timed
//! calls.  The adapter retains a bounded index reservation for the resolver
//! lifetime: compact physical row/cell runs represent many logical positions,
//! while logical repetition and borrowed worksheet payloads do not expand or
//! copy the index.  The background fixture uses ordinary physical rows to
//! exercise index growth while each evaluation still reads one cell.  Every
//! lookup uses the constructor's retained `ExecutionContext`, so work
//! accounting and cancellation checks continue on reused calls.
//! `adapter_index_reserved_bytes_*` is the exact `Resource::Memory`
//! reservation retained by that index; it is separate from evaluator result
//! reservations and is released when the resolver drops.  The worksheet
//! source cells remain borrowed and are not included in that index charge.
//!
//! `memory_retained_used_*` is the execution-budget reservation observed
//! after each result is produced (and therefore includes retained result
//! storage).  It is not a transient allocator peak; allocator counters and
//! `peak_live_delta_*` are reported separately.

use std::{
    alloc::{GlobalAlloc, Layout, System},
    collections::{HashMap, HashSet},
    env,
    error::Error,
    hint::black_box,
    num::{NonZeroU64, NonZeroUsize},
    sync::{
        Mutex,
        atomic::{AtomicU64, Ordering},
    },
    time::{Duration, Instant},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource, SourceVersion,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        ArrayView, CellRead, Context, Evaluated, Limits, Mode, Position, ReferenceView, Resolver,
        SheetExtent, Value, evaluate,
    },
    evaluation::{EvaluationFailure, ScalarError, UnsupportedKind},
    expression::Expression,
};
use litchi_ods::worksheet::formula::{
    Error as WorksheetResolverError, Resolver as WorksheetFormulaResolver,
};
use litchi_ods::worksheet::{
    Cell as WorksheetCell, CellValue as WorksheetCellValue, Row as WorksheetRow, Sheet,
};

type AnyResult<T> = Result<T, Box<dyn Error>>;

const BORROWED_TEXT: &str = "borrowed";

struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        let live =
            LIVE_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed) + layout.size() as u64;
        record_peak(live);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.dealloc(pointer, layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(size as u64, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        let old = layout.size() as u64;
        let new = size as u64;
        let live = if new >= old {
            LIVE_BYTES.fetch_add(new - old, Ordering::Relaxed) + (new - old)
        } else {
            LIVE_BYTES.fetch_sub(old - new, Ordering::Relaxed) - (old - new)
        };
        record_peak(live);
        unsafe { System.realloc(pointer, layout, size) }
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Revision {
    Candidate,
}

impl Revision {
    fn parse(value: &str) -> Self {
        match value {
            "candidate" => Self::Candidate,
            other => panic!("value harness is candidate-only, got revision {other:?}"),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Candidate => "candidate",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Workload {
    Value,
    WorksheetAdapter,
}

impl Workload {
    fn parse(value: &str) -> Self {
        match value {
            "value-evaluation" => Self::Value,
            "worksheet-adapter" => Self::WorksheetAdapter,
            other => panic!("unknown workload {other:?}"),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Value => "value-evaluation",
            Self::WorksheetAdapter => "worksheet-adapter",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Phase {
    Setup,
    Parse,
    Construct,
    Evaluate,
    ParseEvaluate,
}

impl Phase {
    fn parse(value: &str) -> Self {
        match value {
            "setup" => Self::Setup,
            "parse" => Self::Parse,
            "construct" => Self::Construct,
            "evaluate" => Self::Evaluate,
            "parse-evaluate" => Self::ParseEvaluate,
            other => panic!("unknown phase {other:?}"),
        }
    }

    const fn evaluates(self) -> bool {
        matches!(self, Self::Evaluate | Self::ParseEvaluate)
    }
}

#[derive(Clone, Copy, Debug)]
struct Config {
    workload: Workload,
    instrumented: bool,
    revision: Revision,
    group: &'static str,
    phase: Phase,
    case: &'static str,
    warmups: usize,
    iterations: usize,
    repeat: usize,
}

impl Config {
    fn from_args() -> Self {
        let mut workload = None;
        let mut instrumented = false;
        let mut revision = None;
        let mut group = "all";
        let mut phase = None;
        let mut case = None;
        let mut warmups = 3;
        let mut iterations = 15;
        let mut repeat = None;
        let mut args = env::args().skip(1);
        while let Some(argument) = args.next() {
            if argument == "--instrumented" {
                instrumented = true;
                continue;
            }
            let value = args
                .next()
                .unwrap_or_else(|| panic!("missing value for {argument}"));
            match argument.as_str() {
                "--workload" => workload = Some(Workload::parse(&value)),
                "--revision" => revision = Some(Revision::parse(&value)),
                "--group" => {
                    group = match value.as_str() {
                        "all" | "scalar" | "reference" | "array" | "limits" => {
                            Box::leak(value.into_boxed_str())
                        },
                        other => panic!("unknown group {other:?}"),
                    };
                },
                "--phase" => phase = Some(Phase::parse(&value)),
                "--case" => {
                    case = Some(Box::leak(value.into_boxed_str()) as &'static str);
                },
                "--warmups" => warmups = value.parse().expect("warmups must be an integer"),
                "--iterations" => {
                    iterations = value.parse().expect("iterations must be an integer");
                },
                "--repeat" => repeat = Some(value.parse().expect("repeat must be an integer")),
                other => panic!("unknown option {other}"),
            }
        }
        let revision = revision.unwrap_or_else(|| panic!("--revision is required"));
        let workload = workload.unwrap_or_else(|| panic!("--workload is required"));
        let phase = phase.unwrap_or_else(|| panic!("--phase is required"));
        let case = case.unwrap_or_else(|| panic!("--case is required"));
        assert!(iterations > 0, "iterations must be positive");
        let repeat = repeat.unwrap_or_else(|| default_repeat(case));
        Self {
            workload,
            instrumented,
            revision,
            group,
            phase,
            case,
            warmups,
            iterations,
            repeat,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
struct CellKey {
    sheet: u8,
    row: usize,
    column: usize,
}

#[derive(Clone, Copy, Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(&'static str),
    Error(ScalarError),
}

#[derive(Debug)]
struct ResolverStats {
    reads: AtomicU64,
    distinct: Mutex<HashSet<u64>>,
    extent_calls: AtomicU64,
    sheet_index_calls: AtomicU64,
    sheet_name_calls: AtomicU64,
    sheet_count_calls: AtomicU64,
    borrowed_text_bytes: AtomicU64,
    copied_bytes: AtomicU64,
    pointer_checks: AtomicU64,
    pointer_matches: AtomicU64,
}

#[derive(Clone, Copy, Debug, Default)]
struct StatsSnapshot {
    reads: u64,
    distinct_reads: u64,
    extent_calls: u64,
    sheet_index_calls: u64,
    sheet_name_calls: u64,
    sheet_count_calls: u64,
    borrowed_text_bytes: u64,
    copied_bytes: u64,
    pointer_checks: u64,
    pointer_matches: u64,
}

impl ResolverStats {
    fn new(distinct_capacity: usize) -> Self {
        Self {
            reads: AtomicU64::new(0),
            distinct: Mutex::new(HashSet::with_capacity(distinct_capacity)),
            extent_calls: AtomicU64::new(0),
            sheet_index_calls: AtomicU64::new(0),
            sheet_name_calls: AtomicU64::new(0),
            sheet_count_calls: AtomicU64::new(0),
            borrowed_text_bytes: AtomicU64::new(0),
            copied_bytes: AtomicU64::new(0),
            pointer_checks: AtomicU64::new(0),
            pointer_matches: AtomicU64::new(0),
        }
    }

    fn reset(&self) {
        self.reads.store(0, Ordering::Relaxed);
        self.extent_calls.store(0, Ordering::Relaxed);
        self.sheet_index_calls.store(0, Ordering::Relaxed);
        self.sheet_name_calls.store(0, Ordering::Relaxed);
        self.sheet_count_calls.store(0, Ordering::Relaxed);
        self.borrowed_text_bytes.store(0, Ordering::Relaxed);
        self.copied_bytes.store(0, Ordering::Relaxed);
        self.pointer_checks.store(0, Ordering::Relaxed);
        self.pointer_matches.store(0, Ordering::Relaxed);
        self.distinct.lock().expect("distinct resolver set").clear();
    }

    fn snapshot(&self) -> StatsSnapshot {
        StatsSnapshot {
            reads: self.reads.load(Ordering::Acquire),
            distinct_reads: self.distinct.lock().expect("distinct resolver set").len() as u64,
            extent_calls: self.extent_calls.load(Ordering::Acquire),
            sheet_index_calls: self.sheet_index_calls.load(Ordering::Acquire),
            sheet_name_calls: self.sheet_name_calls.load(Ordering::Acquire),
            sheet_count_calls: self.sheet_count_calls.load(Ordering::Acquire),
            borrowed_text_bytes: self.borrowed_text_bytes.load(Ordering::Acquire),
            copied_bytes: self.copied_bytes.load(Ordering::Acquire),
            pointer_checks: self.pointer_checks.load(Ordering::Acquire),
            pointer_matches: self.pointer_matches.load(Ordering::Acquire),
        }
    }
}

#[derive(Debug)]
struct FixtureResolver {
    extent: SheetExtent,
    cells: HashMap<CellKey, FixtureCell>,
    background: Vec<FixtureCell>,
    stats: ResolverStats,
}

impl FixtureResolver {
    fn for_case(case: &str) -> Self {
        let size = scale(case).unwrap_or(1);
        let (rows, columns) = if case.starts_with("reference-background-") {
            (1, size)
        } else if case.starts_with("reference-range-") {
            let side = square_side(size);
            (side, side)
        } else if case.starts_with("reference-distinct-") {
            (8, size.max(8))
        } else if case.starts_with("matrix-lazy-aggregate-") {
            // The aggregate branch scans one physical row.  The inline
            // condition still controls a square result with `size` cells.
            (1, size)
        } else {
            (8, 8)
        };
        let cell_capacity = size.saturating_add(32).max(64);
        let mut resolver = Self {
            extent: SheetExtent::new(rows, columns),
            cells: HashMap::with_capacity(cell_capacity),
            background: if case.starts_with("reference-background-") {
                vec![FixtureCell::Number(0.0); size]
            } else {
                Vec::new()
            },
            stats: ResolverStats::new(cell_capacity.saturating_mul(2)),
        };

        if case.starts_with("reference-range-") {
            let side = square_side(size);
            for row in 0..side {
                for column in 0..side {
                    resolver.set(
                        "Main",
                        row,
                        column,
                        FixtureCell::Number((row * side + column + 1) as f64),
                    );
                }
            }
        } else if case.starts_with("reference-distinct-") {
            for column in 0..size {
                resolver.set("Main", 0, column, FixtureCell::Number((column + 1) as f64));
            }
        } else if case.starts_with("matrix-lazy-aggregate-") {
            for column in 0..size {
                resolver.set("Main", 0, column, FixtureCell::Number(1.0));
            }
        } else {
            resolver.set("Main", 0, 0, FixtureCell::Number(7.0));
            resolver.set("Main", 0, 1, FixtureCell::Text(BORROWED_TEXT));
            resolver.set("Main", 0, 2, FixtureCell::Empty);
            resolver.set("Main", 0, 3, FixtureCell::Error(ScalarError::NotAvailable));
            resolver.set("Main", 0, 4, FixtureCell::Text("two"));
            resolver.set("Main", 0, 7, FixtureCell::Logical(true));
        }
        resolver
    }

    fn set(&mut self, sheet: &str, row: usize, column: usize, value: FixtureCell) {
        let sheet = sheet_id(sheet).unwrap_or(255);
        self.cells.insert(CellKey { sheet, row, column }, value);
    }

    fn reset_stats(&self) {
        self.stats.reset();
    }

    fn stats(&self) -> StatsSnapshot {
        self.stats.snapshot()
    }
}

impl Resolver for FixtureResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        execution.check()?;
        self.stats.extent_calls.fetch_add(1, Ordering::Relaxed);
        Ok(sheet_id(sheet).map(|_| self.extent))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        execution.check()?;
        self.stats.reads.fetch_add(1, Ordering::Relaxed);
        let sheet_id = sheet_id(sheet).unwrap_or(255);
        self.stats
            .distinct
            .lock()
            .expect("distinct resolver set")
            .insert(pack_key(sheet_id, row, column));

        let value = self
            .cells
            .get(&CellKey {
                sheet: sheet_id,
                row,
                column,
            })
            .copied()
            .or_else(|| {
                (sheet_id == 0 && row == 0)
                    .then(|| self.background.get(column).copied())
                    .flatten()
            })
            .unwrap_or(FixtureCell::Empty);
        Ok(match value {
            FixtureCell::Empty => CellRead::Empty,
            FixtureCell::Number(value) => CellRead::Number(value),
            FixtureCell::Logical(value) => CellRead::Logical(value),
            FixtureCell::Text(value) => {
                self.stats
                    .borrowed_text_bytes
                    .fetch_add(value.len() as u64, Ordering::Relaxed);
                self.stats.pointer_checks.fetch_add(1, Ordering::Relaxed);
                if value.as_ptr() == BORROWED_TEXT.as_ptr() {
                    self.stats.pointer_matches.fetch_add(1, Ordering::Relaxed);
                }
                // The resolver returns its immutable fixture slice directly.
                // `copied_bytes == 0` describes this provider only.
                CellRead::Text(value)
            },
            FixtureCell::Error(error) => CellRead::Error(error),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        execution.check()?;
        self.stats.sheet_index_calls.fetch_add(1, Ordering::Relaxed);
        Ok(sheet_id(sheet).map(usize::from))
    }

    fn sheet_name_at<'a>(
        &'a self,
        index: usize,
        execution: &ExecutionContext,
    ) -> Result<Option<&'a str>, EvaluationFailure> {
        execution.check()?;
        self.stats.sheet_name_calls.fetch_add(1, Ordering::Relaxed);
        Ok(match index {
            0 => Some("Main"),
            1 => Some("Data"),
            2 => Some("Archive"),
            _ => None,
        })
    }

    fn sheet_count(&self, execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        execution.check()?;
        self.stats.sheet_count_calls.fetch_add(1, Ordering::Relaxed);
        Ok(3)
    }

    fn source_version(
        &self,
        execution: &ExecutionContext,
    ) -> Result<Option<SourceVersion>, EvaluationFailure> {
        execution.check()?;
        Ok(None)
    }
}

/// A thin measurement wrapper around the production worksheet resolver.
/// Counters are outside the production adapter and only observe calls made by
/// the evaluator.  The wrapper deliberately retains the real resolver/index
/// and returns its borrowed `CellRead` values unchanged.
struct CountingWorksheetResolver<'source> {
    inner: WorksheetFormulaResolver<'source>,
    stats: ResolverStats,
    text_pointer: Option<usize>,
}

impl<'source> CountingWorksheetResolver<'source> {
    fn new(
        sheets: &'source [Sheet],
        extent: SheetExtent,
        execution: &ExecutionContext,
        text_pointer: Option<usize>,
    ) -> Result<Self, WorksheetResolverError> {
        // Keep this explicit call in the harness so construction is visibly
        // measured through the same public production adapter API.
        let inner = WorksheetFormulaResolver::new(sheets, extent, execution)?;
        Ok(Self {
            inner,
            stats: ResolverStats::new(128),
            text_pointer,
        })
    }

    fn reset_stats(&self) {
        self.stats.reset();
    }

    fn stats(&self) -> StatsSnapshot {
        self.stats.snapshot()
    }

    fn reserved_index_bytes(&self) -> u64 {
        self.inner.reserved_index_bytes()
    }
}

impl<'source> Resolver for CountingWorksheetResolver<'source> {
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        self.stats.extent_calls.fetch_add(1, Ordering::Relaxed);
        <WorksheetFormulaResolver<'source> as Resolver>::sheet_extent(&self.inner, sheet, execution)
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.stats.reads.fetch_add(1, Ordering::Relaxed);
        self.stats
            .distinct
            .lock()
            .expect("distinct adapter resolver set")
            .insert(pack_key(sheet_id(sheet).unwrap_or(255), row, column));
        let value = <WorksheetFormulaResolver<'source> as Resolver>::read_cell(
            &self.inner,
            sheet,
            row,
            column,
            execution,
        )?;
        if let CellRead::Text(text) = value {
            self.stats
                .borrowed_text_bytes
                .fetch_add(text.len() as u64, Ordering::Relaxed);
            self.stats.pointer_checks.fetch_add(1, Ordering::Relaxed);
            if self.text_pointer == Some(text.as_ptr() as usize) {
                self.stats.pointer_matches.fetch_add(1, Ordering::Relaxed);
            }
        }
        Ok(value)
    }

    fn sheet_index(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        self.stats.sheet_index_calls.fetch_add(1, Ordering::Relaxed);
        <WorksheetFormulaResolver<'source> as Resolver>::sheet_index(&self.inner, sheet, execution)
    }

    fn sheet_name_at<'a>(
        &'a self,
        index: usize,
        execution: &ExecutionContext,
    ) -> Result<Option<&'a str>, EvaluationFailure> {
        self.stats.sheet_name_calls.fetch_add(1, Ordering::Relaxed);
        <WorksheetFormulaResolver<'source> as Resolver>::sheet_name_at(
            &self.inner,
            index,
            execution,
        )
    }

    fn sheet_count(&self, execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        self.stats.sheet_count_calls.fetch_add(1, Ordering::Relaxed);
        <WorksheetFormulaResolver<'source> as Resolver>::sheet_count(&self.inner, execution)
    }

    fn source_version(
        &self,
        execution: &ExecutionContext,
    ) -> Result<Option<SourceVersion>, EvaluationFailure> {
        <WorksheetFormulaResolver<'source> as Resolver>::source_version(&self.inner, execution)
    }
}

fn worksheet_extent(case: &str) -> SheetExtent {
    match case {
        case if case.starts_with("adapter-repeated-row-") => {
            SheetExtent::new(scale(case).expect("adapter repeated row size"), 8)
        },
        case if case.starts_with("adapter-repeated-cell-") => {
            SheetExtent::new(8, scale(case).expect("adapter repeated cell size"))
        },
        case if case.starts_with("adapter-background-") => SheetExtent::new(
            scale(case)
                .expect("adapter background size")
                .saturating_add(1),
            8,
        ),
        _ => SheetExtent::new(128, 128),
    }
}

fn worksheet_number_cell(value: f64) -> WorksheetCell {
    WorksheetCell::new(WorksheetCellValue::Number(value), value.to_string())
}

fn worksheet_text_cell(value: &str) -> WorksheetCell {
    WorksheetCell::new(WorksheetCellValue::Text(value.to_owned()), value.to_owned())
}

fn worksheet_row_with(cell: WorksheetCell) -> WorksheetRow {
    let mut row = WorksheetRow::new();
    row.push_cell(cell)
        .expect("worksheet benchmark cell must validate");
    row
}

fn worksheet_sheets(case: &str) -> Vec<Sheet> {
    let mut data = Sheet::new("Data").expect("worksheet benchmark sheet must validate");
    let mut other = Sheet::new("Other").expect("worksheet benchmark sheet must validate");
    let archive = Sheet::new("Archive").expect("worksheet benchmark sheet must validate");
    match case {
        "adapter-cell" | "adapter-missing" => {
            data.rows
                .push(worksheet_row_with(worksheet_number_cell(7.0)));
        },
        "adapter-text" => {
            data.rows
                .push(worksheet_row_with(worksheet_text_cell(BORROWED_TEXT)));
        },
        "adapter-empty" => {
            data.rows.push(worksheet_row_with(WorksheetCell::empty()));
        },
        case if case.starts_with("adapter-repeated-row-") => {
            let mut row = WorksheetRow::repeated(scale(case).expect("adapter repeated row size"))
                .expect("worksheet row repetition must be positive");
            row.push_cell(worksheet_number_cell(7.0))
                .expect("worksheet benchmark cell must validate");
            data.rows.push(row);
        },
        case if case.starts_with("adapter-repeated-cell-") => {
            let size = scale(case).expect("adapter repeated cell size");
            let mut row = WorksheetRow::new();
            row.push_cell(
                WorksheetCell::repeated(WorksheetCellValue::Number(7.0), "7", size)
                    .expect("worksheet cell repetition must be positive"),
            )
            .expect("worksheet benchmark cell must validate");
            data.rows.push(row);
        },
        case if case.starts_with("adapter-background-") => {
            data.rows
                .push(worksheet_row_with(worksheet_number_cell(7.0)));
            for _ in 0..scale(case).expect("adapter background size") {
                // Empty physical rows are retained to scale the production
                // index while every timed evaluation still reads Data.A1.
                data.rows.push(WorksheetRow::new());
            }
        },
        other_case => panic!("unknown worksheet adapter case {other_case:?}"),
    }
    // Keep additional names in every fixture so name lookup has a real
    // ordered workbook and the adapter's sheet_count contract is exercised.
    other.rows.push(WorksheetRow::new());
    vec![data, other, archive]
}

fn worksheet_input(case: &str) -> String {
    match case {
        "adapter-cell"
        | "adapter-text"
        | "adapter-empty"
        | "adapter-background-1"
        | "adapter-background-4"
        | "adapter-background-16"
        | "adapter-background-256"
        | "adapter-background-8"
        | "adapter-background-64"
        | "adapter-background-1024"
        | "adapter-background-4096" => "=[Data.A1]".to_owned(),
        "adapter-missing" => "=[Data.Z99]".to_owned(),
        case if case.starts_with("adapter-repeated-row-") => {
            let row = scale(case).expect("adapter repeated row size");
            format!("=[Data.A{}]", row)
        },
        case if case.starts_with("adapter-repeated-cell-") => {
            let column = a1_column(scale(case).expect("adapter repeated cell size") - 1);
            format!("=[Data.{}1]", column)
        },
        other => panic!("unknown worksheet adapter case {other:?}"),
    }
}

fn worksheet_text_pointer(sheets: &[Sheet]) -> Option<usize> {
    sheets
        .first()
        .and_then(|sheet| sheet.rows.first())
        .and_then(|row| row.cells.first())
        .and_then(|cell| match &cell.value {
            WorksheetCellValue::Text(text) => Some(text.as_ptr() as usize),
            _ => None,
        })
}

fn validate_adapter_value(case: &str, value: Value<'_>) -> AnyResult<()> {
    match case {
        "adapter-text" => match value {
            Value::Text(text) if text == BORROWED_TEXT => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected borrowed text"
            ))),
        },
        "adapter-cell"
        | "adapter-background-1"
        | "adapter-background-4"
        | "adapter-background-16"
        | "adapter-background-256"
        | "adapter-background-8"
        | "adapter-background-64"
        | "adapter-background-1024"
        | "adapter-background-4096" => match value {
            Value::Number(7.0) => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected worksheet scalar"
            ))),
        },
        "adapter-empty" => match value {
            Value::Empty => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected Empty"
            ))),
        },
        case if case.starts_with("adapter-repeated-") => match value {
            Value::Number(7.0) => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected Number(7)"
            ))),
        },
        "adapter-missing" => match value {
            Value::Empty => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected Empty"
            ))),
        },
        other => Err(validation_error(format!(
            "no worksheet value expectation for {other}"
        ))),
    }
}

fn evaluate_adapter_once<'expression, R: Resolver + ?Sized>(
    expression: &'expression Expression,
    resolver: &'expression R,
    execution: &ExecutionContext,
) -> Result<Evaluated<'expression>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Data", 0, 0)).with_mode(Mode::Scalar);
    evaluate(expression, resolver, &context, &Limits::default())
}

fn sheet_id(sheet: &str) -> Option<u8> {
    match sheet {
        "Main" => Some(0),
        "Data" => Some(1),
        "Archive" => Some(2),
        _ => None,
    }
}

fn pack_key(sheet: u8, row: usize, column: usize) -> u64 {
    (u64::from(sheet) << 62) ^ ((row as u64 & 0x7fff_ffff) << 31) ^ (column as u64 & 0x7fff_ffff)
}

fn scale(case: &str) -> Option<usize> {
    case.rsplit_once('-')?.1.parse().ok()
}

fn square_side(cells: usize) -> usize {
    let mut side = 1usize;
    while side.saturating_mul(side) < cells {
        side = side.saturating_add(1);
    }
    assert_eq!(
        side.saturating_mul(side),
        cells,
        "corpus size must be square"
    );
    side
}

fn default_repeat(case: &str) -> usize {
    match scale(case) {
        Some(4096) => 1,
        Some(1024) => 2,
        Some(256) => 8,
        Some(16) => 32,
        Some(4) => 64,
        Some(1) => 128,
        _ => 128,
    }
}

fn a1_column(mut column: usize) -> String {
    let mut output = String::new();
    loop {
        let digit = (column % 26) as u8;
        output.insert(0, char::from(b'A' + digit));
        if column < 26 {
            break;
        }
        column = column / 26 - 1;
    }
    output
}

fn array_literal(size: usize, kind: &str) -> String {
    let side = square_side(size);
    let mut rows = Vec::with_capacity(side);
    for row in 0..side {
        let mut cells = Vec::with_capacity(side);
        for column in 0..side {
            let index = row * side + column;
            let value = match kind {
                "number" | "arithmetic" => (index + 1).to_string(),
                "logical" => {
                    if index % 2 == 0 {
                        "TRUE()".to_owned()
                    } else {
                        "FALSE()".to_owned()
                    }
                },
                "bitand" => "6".to_owned(),
                other => panic!("unknown array literal kind {other:?}"),
            };
            cells.push(value);
        }
        rows.push(cells.join(";"));
    }
    format!("={{{}}}", rows.join("|"))
}

fn offset_array_literal(size: usize, offset: usize) -> String {
    let side = square_side(size);
    let mut rows = Vec::with_capacity(side);
    for row in 0..side {
        let mut cells = Vec::with_capacity(side);
        for column in 0..side {
            let index = row * side + column;
            cells.push((index + offset + 1).to_string());
        }
        rows.push(cells.join(";"));
    }
    format!("={{{}}}", rows.join("|"))
}

fn logical_array_literal(size: usize, truth: bool) -> String {
    let side = square_side(size);
    let cell = if truth { "TRUE()" } else { "FALSE()" };
    let row = std::iter::repeat_n(cell, side)
        .collect::<Vec<_>>()
        .join(";");
    format!(
        "={{{}}}",
        std::iter::repeat_n(row, side).collect::<Vec<_>>().join("|")
    )
}

fn matrix_lazy_inline_input(size: usize) -> String {
    let condition = logical_array_literal(size, true);
    let selected = array_literal(size, "number");
    let skipped = offset_array_literal(size, size);
    format!(
        "=IF({};{};{})",
        condition.trim_start_matches('='),
        selected.trim_start_matches('='),
        skipped.trim_start_matches('=')
    )
}

fn matrix_lazy_aggregate_input(size: usize) -> String {
    let condition = logical_array_literal(size, true);
    let end_column = a1_column(size - 1);
    format!(
        "=IF({};AND([.A1:.{}1]);0)",
        condition.trim_start_matches('='),
        end_column
    )
}

fn logical_sequence(size: usize, disjunction: bool) -> String {
    let name = if disjunction { "OR" } else { "AND" };
    let mut source = String::with_capacity(size.saturating_mul(8) + name.len() + 3);
    source.push('=');
    source.push_str(name);
    source.push('(');
    for index in 0..size {
        if index != 0 {
            source.push(';');
        }
        let truth = if disjunction { index + 1 == size } else { true };
        source.push_str(if truth { "TRUE()" } else { "FALSE()" });
    }
    source.push(')');
    source
}

fn reference_chain(size: usize, distinct: bool) -> String {
    let mut source = String::with_capacity(size.saturating_mul(10) + 1);
    source.push('=');
    for index in 0..size {
        if index != 0 {
            source.push('+');
        }
        let column = if distinct {
            a1_column(index)
        } else {
            "A".to_owned()
        };
        source.push_str("[.");
        source.push_str(&column);
        source.push('1');
        source.push(']');
    }
    source
}

fn value_input(case: &str) -> String {
    if let Some(size) = scale(case) {
        if case.starts_with("reference-range-") {
            let side = square_side(size);
            // A bare matrix reference remains `Value::Reference`; use an
            // operator when this lane intentionally forces cell materialize.
            return format!("=0+[.A1:.{}{}]", a1_column(side - 1), side);
        }
        if case.starts_with("reference-background-") {
            return "=[.A1]".to_owned();
        }
        if case.starts_with("reference-repeat-") {
            return reference_chain(size, false);
        }
        if case.starts_with("reference-distinct-") {
            return reference_chain(size, true);
        }
        if case.starts_with("matrix-lazy-inline-") {
            return matrix_lazy_inline_input(size);
        }
        if case.starts_with("matrix-lazy-aggregate-") {
            return matrix_lazy_aggregate_input(size);
        }
        if case.starts_with("array-literal-") {
            return array_literal(size, "number");
        }
        if case.starts_with("array-arithmetic-") {
            return format!("{}+1", array_literal(size, "arithmetic"));
        }
        if case.starts_with("array-not-") {
            return format!(
                "=NOT({})",
                array_literal(size, "logical").trim_start_matches('=')
            );
        }
        if case.starts_with("array-bitand-") {
            return format!(
                "=BITAND({};3)",
                array_literal(size, "bitand").trim_start_matches('=')
            );
        }
        if case.starts_with("array-limit-") {
            return array_literal(size, "number");
        }
        if case.starts_with("sequence-and-") {
            return logical_sequence(size, false);
        }
        if case.starts_with("sequence-or-") {
            return logical_sequence(size, true);
        }
    }
    match case {
        "scalar-number" => "=42".to_owned(),
        "scalar-not" => "=NOT(FALSE())".to_owned(),
        "scalar-bitand" => "=BITAND(6;3)".to_owned(),
        "scalar-roman" => "=ROMAN(3888)".to_owned(),
        "scalar-and" => "=AND(TRUE();1)".to_owned(),
        "scalar-or" => "=OR(FALSE();1)".to_owned(),
        "reference-cell" => "=[.A1]".to_owned(),
        "reference-text" => "=[.B1]".to_owned(),
        "reference-empty" => "=[.C1]".to_owned(),
        "reference-logical" => "=[.H1]".to_owned(),
        "reference-text-arithmetic" => "=0+[.E1]".to_owned(),
        "reference-empty-arithmetic" => "=0+[.C1]".to_owned(),
        "reference-matrix" => "=[.A1:.D4]".to_owned(),
        "reference-error" => "=[.D1]".to_owned(),
        "reference-lazy" => "=IF(TRUE();42;[.A1])".to_owned(),
        "reference-limit-cells"
        | "reference-limit-work"
        | "reference-limit-memory"
        | "reference-cancelled" => "=[.A1:.D4]".to_owned(),
        "array-broadcast-mismatch" => "={1;2}+{3;4;5}".to_owned(),
        // These two lanes deliberately force a matrix through an operator;
        // the direct-reference behavior is covered by `reference-matrix`.
        "array-empty" => "=0+[.F1:.G2]".to_owned(),
        "array-error" => "=0+[.D1:.D2]".to_owned(),
        "array-iferror" => "=IFERROR({#N/A;1;#DIV/0!;2};{10;11;12;13})".to_owned(),
        "array-ifna" => "=IFNA({#N/A;#DIV/0!;1;2};{10;11;12;13})".to_owned(),
        other => panic!("unknown value-harness case {other:?}"),
    }
}

fn mode_for(case: &str) -> Mode {
    if case.starts_with("array-")
        || case.starts_with("matrix-lazy-")
        || case.starts_with("reference-range-")
        || case == "reference-matrix"
    {
        Mode::Matrix
    } else if case.starts_with("reference-limit-") || case == "reference-cancelled" {
        Mode::Matrix
    } else {
        Mode::Scalar
    }
}

fn shape_metadata(case: &str) -> (usize, usize, usize) {
    if let Some(size) = scale(case) {
        if case.starts_with("reference-range-")
            || case.starts_with("array-")
            || case.starts_with("matrix-lazy-")
        {
            let side = square_side(size);
            return (side, side, size);
        }
    }
    match case {
        "reference-limit-cells"
        | "reference-limit-work"
        | "reference-limit-memory"
        | "reference-cancelled"
        | "reference-matrix" => (4, 4, 16),
        "array-broadcast-mismatch" => (1, 3, 3),
        "array-empty" => (2, 2, 4),
        "array-error" => (2, 1, 2),
        "array-iferror" | "array-ifna" => (1, 4, 4),
        _ => (0, 0, 0),
    }
}

fn worksheet_shape_metadata(case: &str) -> (usize, usize, usize) {
    let extent = worksheet_extent(case);
    let physical = match case {
        case if case.starts_with("adapter-repeated-row-") => 1,
        case if case.starts_with("adapter-repeated-cell-") => 1,
        case if case.starts_with("adapter-background-") => scale(case)
            .expect("adapter background size")
            .saturating_add(1),
        _ => 1,
    };
    (extent.rows(), extent.columns(), physical)
}

fn expected_success(case: &str, phase: Phase) -> bool {
    if !phase.evaluates() {
        return true;
    }
    case != "reference-limit-cells"
        && case != "reference-limit-work"
        && case != "reference-limit-memory"
        && case != "reference-cancelled"
        && !case.starts_with("array-limit-")
}

fn expected_failure(case: &str) -> Option<&'static str> {
    if case == "reference-limit-cells" || case.starts_with("array-limit-") {
        Some("resource-objects")
    } else if case == "reference-limit-work" {
        Some("resource-work")
    } else if case == "reference-limit-memory" {
        Some("resource-memory")
    } else if case == "reference-cancelled" {
        Some("cancelled")
    } else {
        None
    }
}

fn limits(case: &str) -> Limits {
    let limits = Limits::default();
    if case == "reference-limit-cells" {
        limits.with_max_reference_cells(8)
    } else if case == "reference-limit-work" {
        limits.with_max_steps(1)
    } else if case == "reference-limit-memory" {
        limits.with_max_storage_bytes(0)
    } else if let Some(size) = case
        .strip_prefix("array-limit-")
        .and_then(|s| s.parse::<usize>().ok())
    {
        limits.with_max_array_cells(size.saturating_sub(1))
    } else {
        limits
    }
}

fn execution(cancelled: bool) -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "ods-formula-array-reference-evaluation-value-profile",
        CoreLimits::for_profile(Profile::Server),
    );
    let (source, token) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(1_u64 << 40).expect("finite in-flight bytes"),
        0,
    )
    .expect("valid execution limits");
    if cancelled {
        source.cancel();
    }
    (
        source,
        ExecutionContext::new(budget, token, execution_limits),
    )
}

fn evaluate_once<'a>(
    expression: &'a Expression,
    resolver: &'a FixtureResolver,
    execution: &ExecutionContext,
    case: &str,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode_for(case));
    evaluate(expression, resolver, &context, &limits(case))
}

fn validation_error(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(std::io::Error::new(
        std::io::ErrorKind::Other,
        message.into(),
    ))
}

fn scalar_error_code(error: ScalarError) -> u64 {
    match error {
        ScalarError::NotAvailable => 1,
        ScalarError::Name => 2,
        ScalarError::Value => 3,
        ScalarError::DivisionByZero => 4,
        ScalarError::Reference => 5,
        ScalarError::Number => 6,
        ScalarError::Null => 7,
        _ => 255,
    }
}

fn checksum_bytes(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

fn checksum_value(value: Value<'_>) -> u64 {
    match value {
        Value::Empty => 0x454d_5054_5900_0001,
        Value::Number(value) => value.to_bits().rotate_left(11) ^ 0x4e,
        Value::Logical(value) => u64::from(value) ^ 0x4c,
        Value::Text(value) => checksum_bytes(value.as_bytes()) ^ 0x54,
        Value::Error(error) => scalar_error_code(error) ^ 0x45,
        Value::Array(array) => {
            let mut checksum =
                ((array.shape().rows() as u64) << 32) ^ array.shape().columns() as u64 ^ 0xa2;
            for index in 0..array.len() {
                if let Some(cell) = array.get(index) {
                    checksum = checksum.rotate_left(5) ^ checksum_value(cell);
                }
            }
            checksum
        },
        Value::Reference(reference) => 0x5245_4645_5245_4e43 ^ reference.len() as u64,
        Value::ReferenceList(references) => 0x5245_4645_5245_4e4c ^ references.len() as u64,
        _ => 0x5641_4c55_4500_ffff,
    }
}

fn validate_number(value: Value<'_>, expected: f64, case: &str) -> AnyResult<()> {
    match value {
        Value::Number(actual) if actual == expected => Ok(()),
        other => Err(validation_error(format!(
            "{case} returned {other:?}, expected Number({expected})"
        ))),
    }
}

fn validate_array_shape(
    array: ArrayView<'_>,
    rows: usize,
    columns: usize,
    case: &str,
) -> AnyResult<()> {
    if array.shape().rows() != rows || array.shape().columns() != columns {
        return Err(validation_error(format!(
            "{case} returned shape {}x{}, expected {rows}x{columns}",
            array.shape().rows(),
            array.shape().columns()
        )));
    }
    Ok(())
}

fn validate_array_numbers(
    array: ArrayView<'_>,
    rows: usize,
    columns: usize,
    value: impl Fn(usize) -> f64,
    case: &str,
) -> AnyResult<()> {
    validate_array_shape(array, rows, columns, case)?;
    for index in 0..array.len() {
        match array.get(index) {
            Some(Value::Number(actual)) if actual == value(index) => {},
            other => {
                return Err(validation_error(format!(
                    "{case} cell {index} returned {other:?}, expected Number({})",
                    value(index)
                )));
            },
        }
    }
    Ok(())
}

fn validate_reference_extent(
    reference: ReferenceView<'_>,
    rows: usize,
    columns: usize,
    case: &str,
) -> AnyResult<()> {
    let areas = reference.areas();
    if areas.len() != 1 || areas[0].extent() != [1, rows, columns] {
        return Err(validation_error(format!(
            "{case} retained {:?}, expected one {rows}x{columns} area",
            areas
        )));
    }
    Ok(())
}

fn validate_value(case: &str, value: Value<'_>) -> AnyResult<()> {
    match case {
        "scalar-number"
        | "reference-cell"
        | "reference-background-8"
        | "reference-background-64"
        | "reference-background-1024"
        | "reference-background-4096" => validate_number(
            value,
            if case == "scalar-number" { 42.0 } else { 7.0 },
            case,
        ),
        "scalar-bitand" => validate_number(value, 2.0, case),
        _ if case.starts_with("sequence-and-") || case.starts_with("sequence-or-") => match value {
            Value::Logical(true) => Ok(()),
            other => Err(validation_error(format!("{case} returned {other:?}"))),
        },
        "scalar-not" => match value {
            Value::Logical(true) => Ok(()),
            other => Err(validation_error(format!("{case} returned {other:?}"))),
        },
        "scalar-and" | "scalar-or" => match value {
            Value::Logical(true) => Ok(()),
            other => Err(validation_error(format!("{case} returned {other:?}"))),
        },
        "scalar-roman" => match value {
            Value::Text(text) if text == "MMMDCCCLXXXVIII" => Ok(()),
            other => Err(validation_error(format!("{case} returned {other:?}"))),
        },
        "reference-text" => match value {
            Value::Text(text) if text == BORROWED_TEXT => Ok(()),
            other => Err(validation_error(format!("{case} returned {other:?}"))),
        },
        "reference-empty" => match value {
            Value::Empty => Ok(()),
            other => Err(validation_error(format!("{case} returned {other:?}"))),
        },
        "reference-logical" => match value {
            Value::Logical(true) => Ok(()),
            other => Err(validation_error(format!("{case} returned {other:?}"))),
        },
        "reference-text-arithmetic" => match value {
            Value::Error(ScalarError::Value) => Ok(()),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected #VALUE from 0+text"
            ))),
        },
        "reference-empty-arithmetic" => validate_number(value, 0.0, case),
        "reference-matrix" => match value {
            Value::Reference(reference) => validate_reference_extent(reference, 4, 4, case),
            other => Err(validation_error(format!(
                "{case} returned {other:?}, expected first-class matrix reference"
            ))),
        },
        "reference-error" => match value {
            Value::Error(ScalarError::NotAvailable) => Ok(()),
            other => Err(validation_error(format!("{case} returned {other:?}"))),
        },
        "reference-lazy" => validate_number(value, 42.0, case),
        _ if case.starts_with("matrix-lazy-inline-") => {
            let size = scale(case).expect("matrix lazy inline size");
            let side = square_side(size);
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_numbers(array, side, side, |index| (index + 1) as f64, case)
        },
        _ if case.starts_with("matrix-lazy-aggregate-") => {
            let size = scale(case).expect("matrix lazy aggregate size");
            let side = square_side(size);
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_shape(array, side, side, case)?;
            for index in 0..array.len() {
                if !matches!(array.get(index), Some(Value::Logical(true))) {
                    return Err(validation_error(format!(
                        "{case} cell {index} is not Logical(true) from AND range"
                    )));
                }
            }
            Ok(())
        },
        "array-broadcast-mismatch" => {
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_shape(array, 1, 3, case)?;
            match (array.get(0), array.get(1), array.get(2)) {
                (
                    Some(Value::Number(4.0)),
                    Some(Value::Number(6.0)),
                    Some(Value::Error(ScalarError::NotAvailable)),
                ) => Ok(()),
                other => Err(validation_error(format!("{case} returned {other:?}"))),
            }
        },
        "array-empty" => {
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_shape(array, 2, 2, case)?;
            for index in 0..array.len() {
                if !matches!(array.get(index), Some(Value::Number(0.0))) {
                    return Err(validation_error(format!(
                        "{case} cell {index} is not Number(0) from 0+Empty"
                    )));
                }
            }
            Ok(())
        },
        "array-error" => {
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_shape(array, 2, 1, case)?;
            if !matches!(array.get(0), Some(Value::Error(ScalarError::NotAvailable))) {
                return Err(validation_error(format!("{case} first cell is not #N/A")));
            }
            for index in 1..array.len() {
                if !matches!(array.get(index), Some(Value::Number(0.0))) {
                    return Err(validation_error(format!(
                        "{case} cell {index} is not Number(0) from 0+Empty"
                    )));
                }
            }
            Ok(())
        },
        "array-iferror" => {
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_shape(array, 1, 4, case)?;
            if !matches!(array.get(0), Some(Value::Number(10.0)))
                || !matches!(array.get(1), Some(Value::Number(1.0)))
                || !matches!(array.get(2), Some(Value::Number(12.0)))
                || !matches!(array.get(3), Some(Value::Number(2.0)))
            {
                return Err(validation_error(format!(
                    "{case} returned unexpected IFERROR array values"
                )));
            }
            Ok(())
        },
        "array-ifna" => {
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_shape(array, 1, 4, case)?;
            if !matches!(array.get(0), Some(Value::Number(10.0)))
                || !matches!(
                    array.get(1),
                    Some(Value::Error(ScalarError::DivisionByZero))
                )
                || !matches!(array.get(2), Some(Value::Number(1.0)))
                || !matches!(array.get(3), Some(Value::Number(2.0)))
            {
                return Err(validation_error(format!(
                    "{case} returned unexpected IFNA array values"
                )));
            }
            Ok(())
        },
        _ if case.starts_with("reference-range-") => {
            let size = scale(case).expect("reference range size");
            let side = square_side(size);
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_numbers(array, side, side, |index| (index + 1) as f64, case)
        },
        _ if case.starts_with("reference-repeat-") => {
            validate_number(value, 7.0 * scale(case).expect("repeat size") as f64, case)
        },
        _ if case.starts_with("reference-distinct-") => {
            let size = scale(case).expect("distinct size") as f64;
            validate_number(value, size * (size + 1.0) / 2.0, case)
        },
        _ if case.starts_with("array-literal-") => {
            let size = scale(case).expect("array literal size");
            let side = square_side(size);
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_numbers(array, side, side, |index| (index + 1) as f64, case)
        },
        _ if case.starts_with("array-arithmetic-") => {
            let size = scale(case).expect("array arithmetic size");
            let side = square_side(size);
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_numbers(array, side, side, |index| (index + 2) as f64, case)
        },
        _ if case.starts_with("array-not-") => {
            let size = scale(case).expect("array NOT size");
            let side = square_side(size);
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_shape(array, side, side, case)?;
            for index in 0..array.len() {
                let expected = index % 2 != 0;
                if !matches!(array.get(index), Some(Value::Logical(actual)) if actual == expected) {
                    return Err(validation_error(format!(
                        "{case} cell {index} has unexpected logical value"
                    )));
                }
            }
            Ok(())
        },
        _ if case.starts_with("array-bitand-") => {
            let size = scale(case).expect("array BITAND size");
            let side = square_side(size);
            let array = match value {
                Value::Array(array) => array,
                other => return Err(validation_error(format!("{case} returned {other:?}"))),
            };
            validate_array_numbers(array, side, side, |_| 2.0, case)
        },
        other => Err(validation_error(format!(
            "no value expectation for {other}"
        ))),
    }
}

fn failure_label(error: &EvaluationFailure) -> &'static str {
    match error {
        EvaluationFailure::Unsupported(kind) => match kind {
            UnsupportedKind::Reference => "unsupported-reference",
            UnsupportedKind::ReferenceOperator => "unsupported-reference-operator",
            UnsupportedKind::Array => "unsupported-array",
            UnsupportedKind::NamedExpression => "unsupported-name",
            UnsupportedKind::Label => "unsupported-label",
            UnsupportedKind::MissingArgument => "unsupported-missing-argument",
            UnsupportedKind::Function => "unsupported-function",
            _ => "unsupported-other",
        },
        EvaluationFailure::InvalidExpression(_) => "invalid-expression",
        EvaluationFailure::Cancelled => "cancelled",
        EvaluationFailure::ResourceLimit(limit) => resource_label(limit.resource),
        EvaluationFailure::Allocation { .. } => "allocation",
        EvaluationFailure::Execution(_) => "execution-error",
        EvaluationFailure::SourceChanged { .. } => "source-changed",
        EvaluationFailure::SourceVersionAvailabilityChanged => "source-version-unavailable",
        _ => "other-error",
    }
}

fn resource_label(resource: Resource) -> &'static str {
    match resource {
        Resource::Memory => "resource-memory",
        Resource::InputBytes => "resource-input-bytes",
        Resource::OutputBytes => "resource-output-bytes",
        Resource::Objects => "resource-objects",
        Resource::Depth => "resource-depth",
        Resource::Work => "resource-work",
        _ => "resource-other",
    }
}

fn validate_evaluation(
    case: &str,
    result: Result<Evaluated<'_>, EvaluationFailure>,
) -> AnyResult<()> {
    match result {
        Ok(value) if expected_success(case, Phase::Evaluate) => validate_value(case, value.value()),
        Ok(value) => Err(validation_error(format!(
            "{case} unexpectedly succeeded with {:?}",
            value.value()
        ))),
        Err(error) if !expected_success(case, Phase::Evaluate) => {
            let expected = expected_failure(case).expect("refusal case has expected failure");
            let actual = failure_label(&error);
            if actual == expected {
                Ok(())
            } else {
                Err(validation_error(format!(
                    "{case} expected {expected}, got {actual}: {error}"
                )))
            }
        },
        Err(error) => Err(validation_error(format!(
            "{case} unexpectedly refused: {error}"
        ))),
    }
}

fn validate_resolver_bounds(case: &str, resolver: &FixtureResolver) -> AnyResult<()> {
    let Some(size) = case
        .strip_prefix("matrix-lazy-aggregate-")
        .and_then(|suffix| suffix.parse::<usize>().ok())
    else {
        return Ok(());
    };
    let stats = resolver.stats();
    let expected = size as u64;
    if stats.reads != expected || stats.distinct_reads != expected {
        return Err(validation_error(format!(
            "{case} aggregate branch read {} cells ({} distinct), expected one {}-cell scan",
            stats.reads, stats.distinct_reads, size
        )));
    }
    Ok(())
}

fn preflight(case: &str, phase: Phase, source: &str) -> AnyResult<()> {
    match phase {
        Phase::Setup => {
            black_box(FixtureResolver::for_case(case));
            Ok(())
        },
        Phase::Parse => Expression::parse(source)
            .map(|_| ())
            .map_err(|error| validation_error(format!("{case} parse preflight failed: {error}"))),
        Phase::Construct => unreachable!("value evaluator has no resolver construct phase"),
        Phase::Evaluate | Phase::ParseEvaluate => {
            let expression = Expression::parse(source).map_err(|error| {
                validation_error(format!("{case} parse preflight failed: {error}"))
            })?;
            let resolver = FixtureResolver::for_case(case);
            let (_source, execution) = execution(case == "reference-cancelled");
            validate_evaluation(
                case,
                evaluate_once(&expression, &resolver, &execution, case),
            )?;
            validate_resolver_bounds(case, &resolver)
        },
    }
}

#[derive(Clone, Copy, Debug)]
struct Sample {
    elapsed: Duration,
    alloc_calls: u64,
    dealloc_calls: u64,
    requested_bytes: u64,
    released_bytes: u64,
    live_before: u64,
    live_after: u64,
    peak_live_delta: u64,
    work_used: u64,
    /// Execution-budget memory still reserved after a successful evaluation.
    /// This is a retained-storage observation, not a transient peak.
    memory_retained_used: u64,
    adapter_index_reserved_bytes: u64,
    resolver: StatsSnapshot,
    successes: u64,
    refusals: u64,
    checksum: u64,
    failure: &'static str,
}

fn note_failure(slot: &mut &'static str, failure: &'static str) {
    if *slot == "none" {
        *slot = failure;
    } else if *slot != failure {
        *slot = "mixed";
    }
}

fn consume_result(
    result: Result<Evaluated<'_>, EvaluationFailure>,
    execution: &ExecutionContext,
    successes: &mut u64,
    refusals: &mut u64,
    checksum: &mut u64,
    memory_peak: &mut u64,
    failure: &mut &'static str,
) {
    match result {
        Ok(value) => {
            *successes = successes.saturating_add(1);
            *checksum = checksum.wrapping_add(checksum_value(value.value()));
            black_box(*checksum);
            *memory_peak = (*memory_peak).max(execution.budget().used(Resource::Memory));
            drop(value);
        },
        Err(error) => {
            *refusals = refusals.saturating_add(1);
            note_failure(failure, failure_label(&error));
        },
    }
}

fn measure_setup(case: &str, repeat: usize) -> Sample {
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0;
    for _ in 0..repeat {
        let source = value_input(case);
        let resolver = FixtureResolver::for_case(case);
        black_box((&source, &resolver));
        successes += 1;
    }
    let elapsed = started.elapsed();
    Sample {
        elapsed,
        alloc_calls: ALLOC_CALLS.load(Ordering::Acquire),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
        requested_bytes: ALLOC_BYTES.load(Ordering::Acquire),
        released_bytes: DEALLOC_BYTES.load(Ordering::Acquire),
        live_before,
        live_after: LIVE_BYTES.load(Ordering::Acquire),
        peak_live_delta: PEAK_LIVE_BYTES
            .load(Ordering::Acquire)
            .saturating_sub(live_before),
        work_used: 0,
        memory_retained_used: 0,
        adapter_index_reserved_bytes: 0,
        resolver: StatsSnapshot::default(),
        successes,
        refusals: 0,
        checksum: 0,
        failure: "none",
    }
}

fn measure_parse(source: &str, repeat: usize) -> Sample {
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0;
    let mut refusals = 0;
    let mut failure = "none";
    for _ in 0..repeat {
        match Expression::parse(black_box(source)) {
            Ok(expression) => {
                successes += 1;
                black_box(&expression);
            },
            Err(error) => {
                refusals += 1;
                note_failure(&mut failure, "parse-error");
                black_box(error);
            },
        }
    }
    let elapsed = started.elapsed();
    Sample {
        elapsed,
        alloc_calls: ALLOC_CALLS.load(Ordering::Acquire),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
        requested_bytes: ALLOC_BYTES.load(Ordering::Acquire),
        released_bytes: DEALLOC_BYTES.load(Ordering::Acquire),
        live_before,
        live_after: LIVE_BYTES.load(Ordering::Acquire),
        peak_live_delta: PEAK_LIVE_BYTES
            .load(Ordering::Acquire)
            .saturating_sub(live_before),
        work_used: 0,
        memory_retained_used: 0,
        adapter_index_reserved_bytes: 0,
        resolver: StatsSnapshot::default(),
        successes,
        refusals,
        checksum: 0,
        failure,
    }
}

fn measure_evaluation(
    case: &str,
    phase: Phase,
    source: &str,
    parsed: Option<&Expression>,
    repeat: usize,
) -> Sample {
    let resolver = FixtureResolver::for_case(case);
    resolver.reset_stats();
    let (_cancellation, execution) = execution(case == "reference-cancelled");
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(mode_for(case));
    let value_limits = limits(case);
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0;
    let mut refusals = 0;
    let mut checksum = 0;
    let mut memory_peak = baseline_memory;
    let mut failure = "none";

    for _ in 0..repeat {
        match phase {
            Phase::Evaluate => {
                let expression = parsed.expect("evaluate phase requires parsed expression");
                consume_result(
                    evaluate(expression, &resolver, &context, &value_limits),
                    &execution,
                    &mut successes,
                    &mut refusals,
                    &mut checksum,
                    &mut memory_peak,
                    &mut failure,
                );
            },
            Phase::ParseEvaluate => match Expression::parse(black_box(source)) {
                Ok(expression) => consume_result(
                    evaluate(&expression, &resolver, &context, &value_limits),
                    &execution,
                    &mut successes,
                    &mut refusals,
                    &mut checksum,
                    &mut memory_peak,
                    &mut failure,
                ),
                Err(error) => {
                    refusals += 1;
                    note_failure(&mut failure, "parse-error");
                    black_box(error);
                },
            },
            Phase::Setup | Phase::Parse | Phase::Construct => {
                unreachable!("evaluation measurement phase")
            },
        }
    }
    let elapsed = started.elapsed();
    let stats = resolver.stats();
    Sample {
        elapsed,
        alloc_calls: ALLOC_CALLS.load(Ordering::Acquire),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
        requested_bytes: ALLOC_BYTES.load(Ordering::Acquire),
        released_bytes: DEALLOC_BYTES.load(Ordering::Acquire),
        live_before,
        live_after: LIVE_BYTES.load(Ordering::Acquire),
        peak_live_delta: PEAK_LIVE_BYTES
            .load(Ordering::Acquire)
            .saturating_sub(live_before),
        work_used: execution
            .budget()
            .used(Resource::Work)
            .saturating_sub(baseline_work),
        memory_retained_used: memory_peak.saturating_sub(baseline_memory),
        adapter_index_reserved_bytes: 0,
        resolver: stats,
        successes,
        refusals,
        checksum,
        failure,
    }
}

/// Metadata shared by the raw and instrumented worksheet-adapter lanes.
/// `worksheet-adapter` is raw by default so its elapsed time contains the
/// production resolver and evaluator only.  `--instrumented` wraps the same
/// resolver with atomics/HashSet counters for lookup and borrowing evidence;
/// those counters are reported separately and their timing is not used as the
/// uninstrumented adapter timing.
trait WorksheetAdapterView: Resolver {
    fn reset_adapter_stats(&self);

    fn adapter_stats(&self) -> StatsSnapshot;

    fn reserved_index_bytes(&self) -> u64;
}

impl<'source> WorksheetAdapterView for WorksheetFormulaResolver<'source> {
    fn reset_adapter_stats(&self) {}

    fn adapter_stats(&self) -> StatsSnapshot {
        StatsSnapshot::default()
    }

    fn reserved_index_bytes(&self) -> u64 {
        WorksheetFormulaResolver::reserved_index_bytes(self)
    }
}

impl<'source> WorksheetAdapterView for CountingWorksheetResolver<'source> {
    fn reset_adapter_stats(&self) {
        self.reset_stats();
    }

    fn adapter_stats(&self) -> StatsSnapshot {
        self.stats()
    }

    fn reserved_index_bytes(&self) -> u64 {
        self.reserved_index_bytes()
    }
}

fn adapter_construction_failure(_error: &WorksheetResolverError) -> &'static str {
    "adapter-construction"
}

fn validate_adapter_evaluation(
    case: &str,
    result: Result<Evaluated<'_>, EvaluationFailure>,
) -> AnyResult<()> {
    match result {
        Ok(value) => validate_adapter_value(case, value.value()),
        Err(error) => Err(validation_error(format!(
            "{case} worksheet adapter unexpectedly refused: {error}"
        ))),
    }
}

fn worksheet_preflight(case: &str, phase: Phase, source: &str) -> AnyResult<()> {
    let sheets = worksheet_sheets(case);
    let extent = worksheet_extent(case);
    match phase {
        Phase::Setup => Ok(()),
        Phase::Construct => {
            let (_cancellation, execution) = execution(false);
            litchi_ods::worksheet::formula::Resolver::new(&sheets, extent, &execution)
                .map(|resolver| {
                    black_box(resolver.reserved_index_bytes());
                })
                .map_err(|error| {
                    validation_error(format!("{case} adapter construction failed: {error}"))
                })
        },
        Phase::Evaluate | Phase::ParseEvaluate => {
            let expression = Expression::parse(source).map_err(|error| {
                validation_error(format!("{case} parse preflight failed: {error}"))
            })?;
            let (_cancellation, execution) = execution(false);
            let resolver = CountingWorksheetResolver::new(
                &sheets,
                extent,
                &execution,
                worksheet_text_pointer(&sheets),
            )
            .map_err(|error| {
                validation_error(format!("{case} adapter construction failed: {error}"))
            })?;
            validate_adapter_evaluation(
                case,
                evaluate_adapter_once(&expression, &resolver, &execution),
            )
        },
        Phase::Parse => unreachable!("worksheet adapter has no parse-only phase"),
    }
}

fn measure_worksheet_setup(case: &str, repeat: usize) -> Sample {
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0;
    for _ in 0..repeat {
        let source = worksheet_input(case);
        let sheets = worksheet_sheets(case);
        black_box((&source, &sheets));
        successes += 1;
    }
    let elapsed = started.elapsed();
    Sample {
        elapsed,
        alloc_calls: ALLOC_CALLS.load(Ordering::Acquire),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
        requested_bytes: ALLOC_BYTES.load(Ordering::Acquire),
        released_bytes: DEALLOC_BYTES.load(Ordering::Acquire),
        live_before,
        live_after: LIVE_BYTES.load(Ordering::Acquire),
        peak_live_delta: PEAK_LIVE_BYTES
            .load(Ordering::Acquire)
            .saturating_sub(live_before),
        work_used: 0,
        memory_retained_used: 0,
        adapter_index_reserved_bytes: 0,
        resolver: StatsSnapshot::default(),
        successes,
        refusals: 0,
        checksum: 0,
        failure: "none",
    }
}

fn measure_worksheet_construct(case: &str, repeat: usize) -> Sample {
    let sheets = worksheet_sheets(case);
    let extent = worksheet_extent(case);
    let (_cancellation, execution) = execution(false);
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0;
    let mut refusals = 0;
    let mut failure = "none";
    let mut index_reserved = 0;
    let mut memory_peak = baseline_memory;
    for _ in 0..repeat {
        match WorksheetFormulaResolver::new(&sheets, extent, &execution) {
            Ok(resolver) => {
                index_reserved = index_reserved.max(resolver.reserved_index_bytes());
                memory_peak = memory_peak.max(execution.budget().used(Resource::Memory));
                black_box(resolver.reserved_index_bytes());
                successes += 1;
            },
            Err(error) => {
                refusals += 1;
                note_failure(&mut failure, adapter_construction_failure(&error));
            },
        }
    }
    let elapsed = started.elapsed();
    Sample {
        elapsed,
        alloc_calls: ALLOC_CALLS.load(Ordering::Acquire),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
        requested_bytes: ALLOC_BYTES.load(Ordering::Acquire),
        released_bytes: DEALLOC_BYTES.load(Ordering::Acquire),
        live_before,
        live_after: LIVE_BYTES.load(Ordering::Acquire),
        peak_live_delta: PEAK_LIVE_BYTES
            .load(Ordering::Acquire)
            .saturating_sub(live_before),
        work_used: execution
            .budget()
            .used(Resource::Work)
            .saturating_sub(baseline_work),
        memory_retained_used: memory_peak.saturating_sub(baseline_memory),
        adapter_index_reserved_bytes: index_reserved,
        resolver: StatsSnapshot::default(),
        successes,
        refusals,
        checksum: index_reserved,
        failure,
    }
}

fn measure_worksheet_evaluation<R: WorksheetAdapterView + ?Sized>(
    _case: &str,
    phase: Phase,
    source: &str,
    parsed: Option<&Expression>,
    repeat: usize,
    resolver: &R,
    execution: &ExecutionContext,
) -> Sample {
    resolver.reset_adapter_stats();
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0;
    let mut refusals = 0;
    let mut checksum = 0;
    let mut memory_peak = baseline_memory;
    let mut failure = "none";
    for _ in 0..repeat {
        match phase {
            Phase::Evaluate => {
                let expression =
                    parsed.expect("worksheet evaluate phase requires parsed expression");
                consume_result(
                    evaluate_adapter_once(expression, resolver, execution),
                    execution,
                    &mut successes,
                    &mut refusals,
                    &mut checksum,
                    &mut memory_peak,
                    &mut failure,
                );
            },
            Phase::ParseEvaluate => match Expression::parse(black_box(source)) {
                Ok(expression) => consume_result(
                    evaluate_adapter_once(&expression, resolver, execution),
                    execution,
                    &mut successes,
                    &mut refusals,
                    &mut checksum,
                    &mut memory_peak,
                    &mut failure,
                ),
                Err(error) => {
                    refusals += 1;
                    note_failure(&mut failure, "parse-error");
                    black_box(error);
                },
            },
            Phase::Setup | Phase::Parse | Phase::Construct => {
                unreachable!("worksheet evaluation measurement phase")
            },
        }
    }
    let elapsed = started.elapsed();
    Sample {
        elapsed,
        alloc_calls: ALLOC_CALLS.load(Ordering::Acquire),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
        requested_bytes: ALLOC_BYTES.load(Ordering::Acquire),
        released_bytes: DEALLOC_BYTES.load(Ordering::Acquire),
        live_before,
        live_after: LIVE_BYTES.load(Ordering::Acquire),
        peak_live_delta: PEAK_LIVE_BYTES
            .load(Ordering::Acquire)
            .saturating_sub(live_before),
        work_used: execution
            .budget()
            .used(Resource::Work)
            .saturating_sub(baseline_work),
        memory_retained_used: memory_peak.saturating_sub(baseline_memory),
        adapter_index_reserved_bytes: resolver.reserved_index_bytes(),
        resolver: resolver.adapter_stats(),
        successes,
        refusals,
        checksum,
        failure,
    }
}

fn measure_worksheet_evaluation_with_index(
    case: &str,
    phase: Phase,
    source: &str,
    parsed: Option<&Expression>,
    repeat: usize,
    instrumented: bool,
) -> Sample {
    let sheets = worksheet_sheets(case);
    let extent = worksheet_extent(case);
    let (_cancellation, execution) = execution(false);
    if instrumented {
        let resolver = CountingWorksheetResolver::new(
            &sheets,
            extent,
            &execution,
            worksheet_text_pointer(&sheets),
        )
        .expect("worksheet adapter construction preflighted");
        measure_worksheet_evaluation(case, phase, source, parsed, repeat, &resolver, &execution)
    } else {
        let resolver = litchi_ods::worksheet::formula::Resolver::new(&sheets, extent, &execution)
            .expect("worksheet adapter construction preflighted");
        measure_worksheet_evaluation(case, phase, source, parsed, repeat, &resolver, &execution)
    }
}

fn measure(
    workload: Workload,
    instrumented: bool,
    case: &str,
    phase: Phase,
    source: &str,
    parsed: Option<&Expression>,
    repeat: usize,
) -> Sample {
    match workload {
        Workload::Value => match phase {
            Phase::Setup => measure_setup(case, repeat),
            Phase::Parse => measure_parse(source, repeat),
            Phase::Construct => unreachable!("value evaluator has no construct phase"),
            Phase::Evaluate | Phase::ParseEvaluate => {
                measure_evaluation(case, phase, source, parsed, repeat)
            },
        },
        Workload::WorksheetAdapter => match phase {
            Phase::Setup => measure_worksheet_setup(case, repeat),
            Phase::Construct => measure_worksheet_construct(case, repeat),
            Phase::Evaluate | Phase::ParseEvaluate => measure_worksheet_evaluation_with_index(
                case,
                phase,
                source,
                parsed,
                repeat,
                instrumented,
            ),
            Phase::Parse => unreachable!("worksheet adapter has no parse-only phase"),
        },
    }
}

fn record_peak(value: u64) {
    let mut current = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    while value > current {
        match PEAK_LIVE_BYTES.compare_exchange_weak(
            current,
            value,
            Ordering::Relaxed,
            Ordering::Relaxed,
        ) {
            Ok(_) => break,
            Err(observed) => current = observed,
        }
    }
}

fn reset_observer() -> u64 {
    ALLOC_CALLS.store(0, Ordering::Relaxed);
    DEALLOC_CALLS.store(0, Ordering::Relaxed);
    ALLOC_BYTES.store(0, Ordering::Relaxed);
    DEALLOC_BYTES.store(0, Ordering::Relaxed);
    let live = LIVE_BYTES.load(Ordering::Acquire);
    PEAK_LIVE_BYTES.store(live, Ordering::Relaxed);
    live
}

fn percentile(samples: &[Sample], metric: impl Fn(&Sample) -> u64, percentile: usize) -> u64 {
    let mut values: Vec<u64> = samples.iter().map(metric).collect();
    values.sort_unstable();
    let index = (values.len().saturating_sub(1) * percentile) / 100;
    values[index]
}

fn mean_ns(samples: &[Sample]) -> u64 {
    let total: u128 = samples.iter().map(|sample| sample.elapsed.as_nanos()).sum();
    (total / samples.len() as u128).min(u64::MAX as u128) as u64
}

fn metric_p50(samples: &[Sample], metric: impl Fn(&Sample) -> u64) -> u64 {
    percentile(samples, metric, 50)
}

fn metric_max(samples: &[Sample], metric: impl Fn(&Sample) -> u64) -> u64 {
    samples.iter().map(metric).max().unwrap_or(0)
}

fn failure_summary(samples: &[Sample]) -> &'static str {
    let first = samples.first().map_or("none", |sample| sample.failure);
    if samples.iter().all(|sample| sample.failure == first) {
        first
    } else {
        "mixed"
    }
}

fn main() -> AnyResult<()> {
    let config = Config::from_args();
    let source = match config.workload {
        Workload::Value => value_input(config.case),
        Workload::WorksheetAdapter => worksheet_input(config.case),
    };
    match config.workload {
        Workload::Value => preflight(config.case, config.phase, &source)?,
        Workload::WorksheetAdapter => worksheet_preflight(config.case, config.phase, &source)?,
    }
    let parsed =
        if config.phase == Phase::Evaluate {
            Some(Expression::parse(&source).map_err(|error| {
                validation_error(format!("{0} parse failed: {error}", config.case))
            })?)
        } else {
            None
        };

    let (rows, columns, elements) = match config.workload {
        Workload::Value => shape_metadata(config.case),
        Workload::WorksheetAdapter => worksheet_shape_metadata(config.case),
    };
    let expected = match config.workload {
        Workload::Value => expected_success(config.case, config.phase),
        Workload::WorksheetAdapter => true,
    };
    let mode = match config.workload {
        Workload::Value => match mode_for(config.case) {
            Mode::Matrix => "matrix",
            Mode::Scalar => "scalar",
            _ => "other",
        },
        Workload::WorksheetAdapter => "scalar",
    };
    println!(
        "config revision={} workload={} instrumented={} group={} phase={} case={} input_bytes={} repeat={} warmups={} iterations={} expected_success={} rows={} columns={} elements={} mode={}",
        config.revision.label(),
        config.workload.label(),
        config.instrumented,
        config.group,
        match config.phase {
            Phase::Setup => "setup",
            Phase::Parse => "parse",
            Phase::Construct => "construct",
            Phase::Evaluate => "evaluate",
            Phase::ParseEvaluate => "parse-evaluate",
        },
        config.case,
        source.len(),
        config.repeat,
        config.warmups,
        config.iterations,
        expected,
        rows,
        columns,
        elements,
        mode,
    );

    for _ in 0..config.warmups {
        black_box(measure(
            config.workload,
            config.instrumented,
            config.case,
            config.phase,
            &source,
            parsed.as_ref(),
            config.repeat,
        ));
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(measure(
            config.workload,
            config.instrumented,
            config.case,
            config.phase,
            &source,
            parsed.as_ref(),
            config.repeat,
        ));
    }
    let p50 = |metric: fn(&Sample) -> u64| metric_p50(&samples, metric);
    let maximum = |metric: fn(&Sample) -> u64| metric_max(&samples, metric);
    let resolver_p50 =
        |metric: fn(&StatsSnapshot) -> u64| metric_p50(&samples, |sample| metric(&sample.resolver));
    let resolver_max =
        |metric: fn(&StatsSnapshot) -> u64| metric_max(&samples, |sample| metric(&sample.resolver));
    println!(
        "result mean_ns={} p50_ns={} p95_ns={} p99_ns={} alloc_calls_p50={} alloc_calls_max={} dealloc_calls_p50={} dealloc_calls_max={} requested_bytes_p50={} requested_bytes_max={} released_bytes_p50={} released_bytes_max={} live_before_p50={} live_after_p50={} live_after_max={} peak_live_delta_p50={} peak_live_delta_max={} work_used_p50={} work_used_max={} memory_retained_used_p50={} memory_retained_used_max={} adapter_index_reserved_bytes_p50={} adapter_index_reserved_bytes_max={} resolver_reads_p50={} resolver_reads_max={} resolver_distinct_reads_p50={} resolver_distinct_reads_max={} resolver_extent_calls_p50={} resolver_extent_calls_max={} resolver_sheet_index_calls_p50={} resolver_sheet_index_calls_max={} resolver_sheet_name_calls_p50={} resolver_sheet_name_calls_max={} resolver_sheet_count_calls_p50={} resolver_sheet_count_calls_max={} resolver_borrowed_text_bytes_p50={} resolver_borrowed_text_bytes_max={} resolver_copied_bytes_p50={} resolver_copied_bytes_max={} resolver_pointer_checks_p50={} resolver_pointer_checks_max={} resolver_pointer_matches_p50={} resolver_pointer_matches_max={} successes_p50={} successes_max={} refusals_p50={} refusals_max={} checksum_p50={} checksum_max={} failure={}",
        mean_ns(&samples),
        p50(|sample| sample.elapsed.as_nanos().min(u64::MAX as u128) as u64),
        percentile(
            &samples,
            |sample| sample.elapsed.as_nanos().min(u64::MAX as u128) as u64,
            95,
        ),
        percentile(
            &samples,
            |sample| sample.elapsed.as_nanos().min(u64::MAX as u128) as u64,
            99,
        ),
        p50(|sample| sample.alloc_calls),
        maximum(|sample| sample.alloc_calls),
        p50(|sample| sample.dealloc_calls),
        maximum(|sample| sample.dealloc_calls),
        p50(|sample| sample.requested_bytes),
        maximum(|sample| sample.requested_bytes),
        p50(|sample| sample.released_bytes),
        maximum(|sample| sample.released_bytes),
        p50(|sample| sample.live_before),
        p50(|sample| sample.live_after),
        maximum(|sample| sample.live_after),
        p50(|sample| sample.peak_live_delta),
        maximum(|sample| sample.peak_live_delta),
        p50(|sample| sample.work_used),
        maximum(|sample| sample.work_used),
        p50(|sample| sample.memory_retained_used),
        maximum(|sample| sample.memory_retained_used),
        p50(|sample| sample.adapter_index_reserved_bytes),
        maximum(|sample| sample.adapter_index_reserved_bytes),
        resolver_p50(|stats| stats.reads),
        resolver_max(|stats| stats.reads),
        resolver_p50(|stats| stats.distinct_reads),
        resolver_max(|stats| stats.distinct_reads),
        resolver_p50(|stats| stats.extent_calls),
        resolver_max(|stats| stats.extent_calls),
        resolver_p50(|stats| stats.sheet_index_calls),
        resolver_max(|stats| stats.sheet_index_calls),
        resolver_p50(|stats| stats.sheet_name_calls),
        resolver_max(|stats| stats.sheet_name_calls),
        resolver_p50(|stats| stats.sheet_count_calls),
        resolver_max(|stats| stats.sheet_count_calls),
        resolver_p50(|stats| stats.borrowed_text_bytes),
        resolver_max(|stats| stats.borrowed_text_bytes),
        resolver_p50(|stats| stats.copied_bytes),
        resolver_max(|stats| stats.copied_bytes),
        resolver_p50(|stats| stats.pointer_checks),
        resolver_max(|stats| stats.pointer_checks),
        resolver_p50(|stats| stats.pointer_matches),
        resolver_max(|stats| stats.pointer_matches),
        p50(|sample| sample.successes),
        maximum(|sample| sample.successes),
        p50(|sample| sample.refusals),
        maximum(|sample| sample.refusals),
        p50(|sample| sample.checksum),
        maximum(|sample| sample.checksum),
        failure_summary(&samples),
    );
    Ok(())
}
