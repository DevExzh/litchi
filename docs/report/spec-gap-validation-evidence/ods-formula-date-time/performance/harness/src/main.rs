//! Bounded process-level profile for the ODS date/time evaluator.
//!
//! The fixture is deterministic and in memory. Expected date serials and
//! sequence results are computed by the local oracle helpers below, never by
//! invoking a second production evaluator. Reference cells are borrowed and
//! counted so timing rows preserve read/work/resource evidence.

#![allow(dead_code)]

use std::{
    alloc::{GlobalAlloc, Layout, System},
    env,
    error::Error,
    fmt::Write as _,
    hint::black_box,
    num::{NonZeroU64, NonZeroUsize},
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, EvaluationOptions, ScalarError,
        ScalarValue, evaluate_scalar,
    },
    expression::Expression,
};

#[cfg(feature = "date-time-candidate")]
use litchi_ods::codec::formula::evaluation::{CalculationTimestamp, UnsupportedKind};

type AnyResult<T> = Result<T, Box<dyn Error>>;

const TIMESTAMP_SERIAL: f64 = 45_382.5;

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
enum Shape {
    Scalar,
    Array {
        rows: usize,
        columns: usize,
        reference: bool,
    },
}

impl Shape {
    const fn is_scalar(self) -> bool {
        matches!(self, Self::Scalar)
    }

    const fn is_reference(self) -> bool {
        matches!(
            self,
            Self::Array {
                reference: true,
                ..
            }
        )
    }

    const fn dimensions(self) -> Option<(usize, usize)> {
        match self {
            Self::Scalar => None,
            Self::Array { rows, columns, .. } => Some((rows, columns)),
        }
    }

    const fn mode(self) -> Mode {
        Mode::Matrix
    }

    const fn elements(self) -> usize {
        match self {
            Self::Scalar => 1,
            Self::Array { rows, columns, .. } => rows.saturating_mul(columns),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Path {
    Scalar,
    Value,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FailureKind {
    ReferenceCells,
    Cancelled,
    Work,
    CalculationClock,
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum Expected {
    Number(f64),
    NumberAny,
    Logical(bool),
    LogicalAny,
    Text(&'static str),
    Error(ScalarError),
    Failure(FailureKind),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ControlOp {
    Arithmetic,
    Sin,
    ImSum,
    Average,
    CountA,
    Var,
    Stdev,
    DSum,
    DVar,
    DStdev,
    Median,
    Rank,
    PercentRank,
    Concatenate,
    ValueDate,
    ValueTime,
    ValueMixedFraction,
    LazyIsNumber,
    LazyIsBlank,
    LazyIsText,
}

impl ControlOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Arithmetic => "ARITHMETIC",
            Self::Sin => "SIN",
            Self::ImSum => "IMSUM",
            Self::Average => "AVERAGE",
            Self::CountA => "COUNTA",
            Self::Var => "VAR",
            Self::Stdev => "STDEV",
            Self::DSum => "DSUM",
            Self::DVar => "DVAR",
            Self::DStdev => "DSTDEV",
            Self::Median => "MEDIAN",
            Self::Rank => "RANK",
            Self::PercentRank => "PERCENTRANK",
            Self::Concatenate => "CONCATENATE",
            Self::ValueDate | Self::ValueTime | Self::ValueMixedFraction => "VALUE",
            Self::LazyIsNumber => "ISNUMBER",
            Self::LazyIsBlank => "ISBLANK",
            Self::LazyIsText => "ISTEXT",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DateOp {
    Date,
    Datedif,
    DateValue,
    Day,
    Days,
    Days360,
    EasterSunday,
    Edate,
    Eomonth,
    Hour,
    IsoWeeknum,
    Minute,
    Month,
    Networkdays,
    Now,
    Second,
    Time,
    TimeValue,
    Today,
    Weekday,
    Weeknum,
    Workday,
    Year,
    Yearfrac,
}

impl DateOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Date => "DATE",
            Self::Datedif => "DATEDIF",
            Self::DateValue => "DATEVALUE",
            Self::Day => "DAY",
            Self::Days => "DAYS",
            Self::Days360 => "DAYS360",
            Self::EasterSunday => "EASTERSUNDAY",
            Self::Edate => "EDATE",
            Self::Eomonth => "EOMONTH",
            Self::Hour => "HOUR",
            Self::IsoWeeknum => "ISOWEEKNUM",
            Self::Minute => "MINUTE",
            Self::Month => "MONTH",
            Self::Networkdays => "NETWORKDAYS",
            Self::Now => "NOW",
            Self::Second => "SECOND",
            Self::Time => "TIME",
            Self::TimeValue => "TIMEVALUE",
            Self::Today => "TODAY",
            Self::Weekday => "WEEKDAY",
            Self::Weeknum => "WEEKNUM",
            Self::Workday => "WORKDAY",
            Self::Year => "YEAR",
            Self::Yearfrac => "YEARFRAC",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DateLane {
    Core,
    Parser,
    Boundary,
    Reference,
    Projected,
    HolidayScale,
    WorkweekScale,
    IntervalScale,
    SequenceProjection,
    Lazy,
    Timestamp,
    MissingTimestamp,
    Refusal,
    Cancellation,
    Resource,
    WorkLimit,
    FormulaError,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ResolverMode {
    Control,
    Database,
    Conditional,
    Lazy,
    DateNormal,
    DateAllWorkdays,
    DateAllOff,
    DateMalformedWorkweek,
    DateErrorCell,
}

#[derive(Debug)]
struct Case {
    name: String,
    source: String,
    shape: Shape,
    path: Path,
    operation: Option<ControlOp>,
    date: Option<(DateOp, DateLane)>,
    expectation: Expected,
    read_bound: u64,
    timestamp: bool,
    max_reference_cells: Option<usize>,
    max_steps: Option<u64>,
    cancellation: bool,
    resolver_mode: ResolverMode,
}

#[derive(Debug)]
struct ResolverStats {
    reads: AtomicU64,
}

impl ResolverStats {
    fn reset(&self) {
        self.reads.store(0, Ordering::Release);
    }

    fn reads(&self) -> u64 {
        self.reads.load(Ordering::Acquire)
    }
}

#[derive(Debug)]
struct FixtureResolver {
    stats: ResolverStats,
    mode: ResolverMode,
    cancel_after_read: Option<CancellationSource>,
}

impl FixtureResolver {
    fn new(mode: ResolverMode, cancel_after_read: Option<CancellationSource>) -> Self {
        Self {
            stats: ResolverStats {
                reads: AtomicU64::new(0),
            },
            mode,
            cancel_after_read,
        }
    }
}

impl Resolver for FixtureResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        execution.check()?;
        Ok((matches!(sheet, "Main" | "Data" | "Archive")).then_some(SheetExtent::new(2048, 24)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        execution.check()?;
        let read = self.stats.reads.fetch_add(1, Ordering::Relaxed) + 1;
        if read == 1 {
            if let Some(cancellation) = &self.cancel_after_read {
                cancellation.cancel();
            }
        }
        if !matches!(sheet, "Main" | "Data" | "Archive") || row >= 2048 || column >= 24 {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        match self.mode {
            ResolverMode::DateNormal
            | ResolverMode::DateAllWorkdays
            | ResolverMode::DateAllOff
            | ResolverMode::DateMalformedWorkweek
            | ResolverMode::DateErrorCell => Ok(date_cell(self.mode, row, column)),
            ResolverMode::Database => Ok(database_cell(row, column)),
            ResolverMode::Conditional => Ok(conditional_cell(row, column)),
            ResolverMode::Lazy => Ok(lazy_cell(row, column)),
            ResolverMode::Control => Ok(CellRead::Number(aggregate_value(row, column))),
        }
    }

    fn sheet_index(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        execution.check()?;
        Ok(match sheet {
            "Main" => Some(0),
            "Data" => Some(1),
            "Archive" => Some(2),
            _ => None,
        })
    }

    fn sheet_name_at(
        &self,
        index: usize,
        execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        execution.check()?;
        Ok(match index {
            0 => Some("Main"),
            1 => Some("Data"),
            2 => Some("Archive"),
            _ => None,
        })
    }

    fn sheet_count(&self, execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        execution.check()?;
        Ok(3)
    }
}

fn aggregate_value(row: usize, column: usize) -> f64 {
    1.01 + ((row.saturating_mul(4) + column % 4) % 17) as f64 * 0.001
}

fn database_cell(row: usize, column: usize) -> CellRead<'static> {
    match (row, column) {
        (0, 0) => CellRead::Text("Name"),
        (0, 1) => CellRead::Text("Value"),
        (1, 0) => CellRead::Text("A"),
        (2, 0) => CellRead::Text("B"),
        (3, 0) => CellRead::Text("C"),
        (1, 1) => CellRead::Number(7.0),
        (2, 1) => CellRead::Number(1.0),
        (3, 1) => CellRead::Number(4.0),
        (0, 3) => CellRead::Text("Value"),
        (1, 3) => CellRead::Text(">0"),
        _ => CellRead::Empty,
    }
}

fn conditional_cell(row: usize, column: usize) -> CellRead<'static> {
    let index = row.saturating_mul(4) + column % 4;
    match column / 4 {
        0 => CellRead::Number((index % 4) as f64),
        1 => CellRead::Number((index % 2) as f64),
        2 => CellRead::Number(0.5 + index as f64 * 0.03125),
        3 => CellRead::Text(if row % 2 == 0 { "A" } else { "B" }),
        _ => CellRead::Empty,
    }
}

fn lazy_cell(row: usize, column: usize) -> CellRead<'static> {
    match (row, column) {
        (0, 0) => CellRead::Empty,
        (1, 0) => CellRead::Text(""),
        (2, 0) => CellRead::Number(42.5),
        (3, 0) => CellRead::Logical(true),
        _ => CellRead::Empty,
    }
}

fn date_cell(mode: ResolverMode, row: usize, column: usize) -> CellRead<'static> {
    match column {
        0 => CellRead::Number(holiday_serial(row)),
        1 => {
            if mode == ResolverMode::DateMalformedWorkweek && row == 0 {
                CellRead::Text("bad-workweek")
            } else if mode == ResolverMode::DateAllOff {
                CellRead::Logical(true)
            } else if mode == ResolverMode::DateAllWorkdays {
                CellRead::Logical(false)
            } else {
                CellRead::Logical(matches!(row, 0 | 6))
            }
        },
        2 => {
            if mode == ResolverMode::DateErrorCell && row == 0 {
                CellRead::Error(ScalarError::NotAvailable)
            } else {
                CellRead::Number(date_serial(2024, 3, 31) + row as f64)
            }
        },
        3 => CellRead::Text("2024-03-31T12:34:56"),
        4 => CellRead::Empty,
        _ => CellRead::Number(0.0),
    }
}

fn date_serial(year: i32, month: u32, day: u32) -> f64 {
    fn leap(year: i32) -> bool {
        year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
    }
    fn ordinal(year: i32, month: u32, day: u32) -> i64 {
        let lengths = [31_i64, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        let mut result = 0_i64;
        for value in 1..year {
            result += if leap(value) { 366 } else { 365 };
        }
        for value in 1..month {
            result += lengths[(value - 1) as usize];
            if value == 2 && leap(year) {
                result += 1;
            }
        }
        result + i64::from(day)
    }
    (ordinal(year, month, day) - ordinal(1899, 12, 30)) as f64
}

fn holiday_serial(index: usize) -> f64 {
    date_serial(2024, 1, 1) + (index as f64 * 17.0)
}

fn weekday(serial: i64) -> usize {
    (serial + 6).rem_euclid(7) as usize
}

fn default_workweek() -> [bool; 7] {
    [true, false, false, false, false, false, true]
}

fn networkdays(start: i64, end: i64, holiday_count: usize, workweek: [bool; 7]) -> f64 {
    let direction = if start <= end { 1_i64 } else { -1_i64 };
    let mut day = start;
    let mut total = 0_i64;
    loop {
        let is_holiday = (0..holiday_count)
            .map(|index| holiday_serial(index) as i64)
            .any(|holiday| holiday == day);
        if !workweek[weekday(day)] && !is_holiday {
            total += 1;
        }
        if day == end {
            break;
        }
        day += direction;
    }
    (total * direction) as f64
}

fn workday(start: i64, offset: i64, holiday_count: usize, workweek: [bool; 7]) -> Option<f64> {
    if offset == 0 {
        return Some(start as f64);
    }
    let direction = offset.signum();
    let mut remaining = offset.unsigned_abs();
    let mut day = start;
    while remaining != 0 {
        day += direction;
        let is_holiday = (0..holiday_count)
            .map(|index| holiday_serial(index) as i64)
            .any(|holiday| holiday == day);
        if !workweek[weekday(day)] && !is_holiday {
            remaining -= 1;
        }
        if day.abs() > 2_958_465 {
            return None;
        }
    }
    Some(day as f64)
}

fn expected_control_number(operation: ControlOp) -> f64 {
    match operation {
        ControlOp::Arithmetic => 1.42,
        ControlOp::Sin => 0.17_f64.sin(),
        ControlOp::ImSum => 3.0,
        ControlOp::Average => 2.5,
        ControlOp::CountA => 3.0,
        ControlOp::Var => 1.5625,
        ControlOp::Stdev => 1.25,
        ControlOp::DSum => 12.0,
        ControlOp::DVar => 9.0,
        ControlOp::DStdev => 3.0,
        ControlOp::Median => 3.0,
        ControlOp::Rank => 1.0,
        ControlOp::PercentRank => 1.0,
        ControlOp::ValueDate => 45_293.0,
        ControlOp::ValueTime => 45_296.0 / 86_400.0,
        ControlOp::ValueMixedFraction => 3.5,
        ControlOp::LazyIsNumber | ControlOp::LazyIsBlank | ControlOp::LazyIsText => 0.0,
        ControlOp::Concatenate => 0.0,
    }
}

fn date_expected(case: &Case, index: usize) -> Option<f64> {
    let name = case.name.as_str();
    let start = date_serial(2024, 1, 1) as i64;
    let end = date_serial(2024, 12, 31) as i64;
    if name == "date-core-date" {
        return Some(date_serial(2020, 13, 1));
    }
    if name == "date-core-datedif" {
        return Some(1.0);
    }
    if name == "date-core-datevalue" {
        return Some(date_serial(2024, 3, 31));
    }
    if name == "date-core-day" {
        return Some(31.0);
    }
    if name == "date-core-days" {
        return Some(1.0);
    }
    if name == "date-core-days360" {
        return Some(29.0);
    }
    if name == "date-core-eastersunday" {
        return Some(date_serial(2024, 3, 31));
    }
    if name == "date-core-edate" {
        return Some(date_serial(2024, 2, 29));
    }
    if name == "date-core-eomonth" {
        return Some(date_serial(2024, 2, 29));
    }
    if name == "date-core-hour" {
        return Some(18.0);
    }
    if name == "date-core-isoweeknum" {
        return Some(1.0);
    }
    if name == "date-core-minute" {
        return Some(54.0);
    }
    if name == "date-core-month" {
        return Some(3.0);
    }
    if name == "date-core-networkdays" {
        return Some(2.0);
    }
    if name == "date-core-now" {
        return Some(TIMESTAMP_SERIAL);
    }
    if name == "date-core-second" {
        return Some(22.0);
    }
    if name == "date-core-time" || name == "date-core-timevalue" {
        return Some(12.5 / 24.0);
    }
    if name == "date-core-today" {
        return Some(TIMESTAMP_SERIAL.floor());
    }
    if name == "date-core-weekday" {
        return Some(1.0);
    }
    if name == "date-core-weeknum" {
        return Some(1.0);
    }
    if name == "date-core-workday" {
        return Some(date_serial(2024, 4, 1));
    }
    if name == "date-core-year" {
        return Some(2024.0);
    }
    if name == "date-core-yearfrac" {
        return Some(366.0 / 365.0);
    }
    if name == "date-timestamp-now" {
        return Some(TIMESTAMP_SERIAL);
    }
    if name == "date-timestamp-today" {
        return Some(TIMESTAMP_SERIAL.floor());
    }
    if name == "date-timestamp-eastersunday" {
        return Some(date_serial(2025, 4, 20));
    }
    if name == "date-parser-datevalue-iso" {
        return Some(date_serial(2024, 3, 31));
    }
    if name == "date-parser-datevalue-datetime" {
        return Some(date_serial(2024, 3, 31));
    }
    if name == "date-parser-datevalue-enus" {
        return Some(date_serial(2024, 3, 31));
    }
    if name == "date-parser-datevalue-month" {
        return Some(date_serial(2024, 3, 31));
    }
    if name == "date-parser-timevalue-clock" {
        return Some(12.5 / 24.0);
    }
    if name == "date-parser-timevalue-datetime" {
        return Some((3.0 * 3_600.0 + 4.0 * 60.0 + 5.0) / 86_400.0);
    }
    if name == "date-parser-timevalue-fraction" {
        return Some((12.0 * 3600.0 + 30.0 * 60.0 + 0.5) / 86_400.0);
    }
    if name == "date-parser-datevalue-number" {
        return Some(123.0);
    }
    if name == "date-parser-timevalue-number" {
        return Some(0.5);
    }
    if name == "date-boundary-date-rollover" {
        return Some(date_serial(2021, 3, 1));
    }
    if name == "date-boundary-date-1900" {
        return Some(date_serial(1900, 3, 1));
    }
    if name == "date-boundary-days360-eu" {
        return Some(30.0);
    }
    if name == "date-boundary-edate-clamp" {
        return Some(date_serial(2024, 2, 29));
    }
    if name == "date-boundary-eomonth-clamp" {
        return Some(date_serial(2024, 2, 29));
    }
    if name == "date-boundary-yearfrac-basis1" {
        return Some(1.0);
    }
    if name == "date-boundary-weekday-17" {
        return Some(1.0);
    }
    if let Some(value) = name.strip_prefix("date-networkdays-holidays-") {
        let count = value.parse::<usize>().ok()?;
        return Some(networkdays(start, end, count, default_workweek()));
    }
    if let Some(value) = name.strip_prefix("date-workday-holidays-") {
        let count = value.parse::<usize>().ok()?;
        return workday(start, 64, count, default_workweek());
    }
    if name == "date-networkdays-workweek-default" {
        return Some(networkdays(start, start + 31, 0, default_workweek()));
    }
    if name == "date-networkdays-workweek-custom" {
        return Some(networkdays(
            start,
            start + 31,
            0,
            [false, false, false, false, false, false, false],
        ));
    }
    if name == "date-networkdays-workweek-alloff" {
        return Some(0.0);
    }
    if name == "date-workday-workweek-default" {
        return workday(start, 64, 0, default_workweek());
    }
    if name == "date-workday-workweek-custom" {
        return workday(
            start,
            64,
            0,
            [false, false, false, false, false, false, false],
        );
    }
    if name == "date-networkdays-interval-7" {
        return Some(networkdays(start, start + 6, 0, default_workweek()));
    }
    if name == "date-networkdays-interval-31" {
        return Some(networkdays(start, start + 30, 0, default_workweek()));
    }
    if name == "date-networkdays-interval-365" {
        return Some(networkdays(start, start + 364, 0, default_workweek()));
    }
    if name == "date-networkdays-interval-4096" {
        return Some(networkdays(start, start + 4095, 0, default_workweek()));
    }
    if let Some(value) = name.strip_prefix("date-workday-offset-") {
        let offset = value.parse::<i64>().ok()?;
        return workday(start, offset, 0, default_workweek());
    }
    if name == "date-sequence-networkdays-projection" {
        return Some(networkdays(start, start + 31, 16, default_workweek()));
    }
    if name == "date-sequence-workday-projection" {
        return workday(start, 64, 16, default_workweek());
    }
    if name == "date-projected-day" {
        return Some(if index == 0 { 29.0 } else { 1.0 });
    }
    if name == "date-lazy-unselected" {
        return Some(0.0);
    }
    None
}

fn date_source(name: &str) -> String {
    match name {
        "date-core-date" => "=DATE(2020;13;1)".to_owned(),
        "date-core-datedif" => "=DATEDIF(DATE(2020;1;31);DATE(2021;3;1);\"Y\")".to_owned(),
        "date-core-datevalue" => "=DATEVALUE(\"2024-03-31T12:34:56\")".to_owned(),
        "date-core-day" => "=DAY(DATE(2024;3;31)+0.75)".to_owned(),
        "date-core-days" => "=DAYS(DATE(2024;4;1)+0.5;DATE(2024;3;31)+0.5)".to_owned(),
        "date-core-days360" => "=DAYS360(DATE(2024;1;31);DATE(2024;2;29))".to_owned(),
        "date-core-eastersunday" => "=EASTERSUNDAY(2024)".to_owned(),
        "date-core-edate" => "=EDATE(DATE(2024;1;31);1)".to_owned(),
        "date-core-eomonth" => "=EOMONTH(DATE(2024;2;10);0)".to_owned(),
        "date-core-hour" => "=HOUR(0.75)".to_owned(),
        "date-core-isoweeknum" => "=ISOWEEKNUM(DATE(2024;1;1))".to_owned(),
        "date-core-minute" => "=MINUTE(-1/256)".to_owned(),
        "date-core-month" => "=MONTH(DATE(2024;3;31)+0.5)".to_owned(),
        "date-core-networkdays" => "=NETWORKDAYS(DATE(2024;3;29);DATE(2024;4;1))".to_owned(),
        "date-core-now" => "=NOW()".to_owned(),
        "date-core-second" => "=SECOND(-1/256)".to_owned(),
        "date-core-time" => "=TIME(12;30;0)".to_owned(),
        "date-core-timevalue" => "=TIMEVALUE(\"12:30:00\")".to_owned(),
        "date-core-today" => "=TODAY()".to_owned(),
        "date-core-weekday" => "=WEEKDAY(DATE(2024;1;1);2)".to_owned(),
        "date-core-weeknum" => "=WEEKNUM(DATE(2024;1;1);21)".to_owned(),
        "date-core-workday" => "=WORKDAY(DATE(2024;3;29);1)".to_owned(),
        "date-core-year" => "=YEAR(DATE(2024;3;31)+0.5)".to_owned(),
        "date-core-yearfrac" => "=YEARFRAC(DATE(2024;1;1);DATE(2025;1;1);3)".to_owned(),
        "date-parser-datevalue-iso" => "=DATEVALUE(\"2024-03-31\")".to_owned(),
        "date-parser-datevalue-datetime" => "=DATEVALUE(\"2024-03-31 12:34:56\")".to_owned(),
        "date-parser-datevalue-enus" => "=DATEVALUE(\"3/31/2024\")".to_owned(),
        "date-parser-datevalue-month" => "=DATEVALUE(\"Mar 31, 2024\")".to_owned(),
        "date-parser-timevalue-clock" => "=TIMEVALUE(\"12:30:00\")".to_owned(),
        "date-parser-timevalue-datetime" => {
            "=TIMEVALUE(\"2020-01-02T03:04:05\")".to_owned()
        },
        "date-parser-timevalue-fraction" => "=TIMEVALUE(\"12:30:00.5\")".to_owned(),
        "date-parser-datevalue-number" => "=DATEVALUE(\"123\")".to_owned(),
        "date-parser-timevalue-number" => "=TIMEVALUE(\"0.5\")".to_owned(),
        "date-boundary-date-rollover" => "=DATE(2020;2;30)".to_owned(),
        "date-boundary-date-1900" => "=DATE(1900;3;1)".to_owned(),
        "date-boundary-days360-eu" => "=DAYS360(DATE(2024;1;31);DATE(2024;2;29);TRUE())".to_owned(),
        "date-boundary-edate-clamp" => "=EDATE(DATE(2024;1;31);1)".to_owned(),
        "date-boundary-eomonth-clamp" => "=EOMONTH(DATE(2024;2;10);0)".to_owned(),
        "date-boundary-yearfrac-basis1" => "=YEARFRAC(DATE(2024;1;1);DATE(2025;1;1);1)".to_owned(),
        "date-boundary-weeknum-noninteger" => "=WEEKNUM(44199;21.5)".to_owned(),
        "date-boundary-weekday-17" => "=WEEKDAY(DATE(2024;1;7);17)".to_owned(),
        _ if name.starts_with("date-networkdays-holidays-") => {
            let count = name.rsplit('-').next().unwrap();
            if count == "0" {
                "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;12;31))".to_owned()
            } else {
                format!("=NETWORKDAYS(DATE(2024;1;1);DATE(2024;12;31);[.A1:.A{count}])")
            }
        },
        _ if name.starts_with("date-workday-holidays-") => {
            let count = name.rsplit('-').next().unwrap();
            if count == "0" {
                "=WORKDAY(DATE(2024;1;1);64)".to_owned()
            } else {
                format!("=WORKDAY(DATE(2024;1;1);64;[.A1:.A{count}])")
            }
        },
        "date-networkdays-workweek-default" => {
            "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);;[.B1:.B7])".to_owned()
        },
        "date-networkdays-workweek-custom" => {
            "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);;[.B1:.B7])".to_owned()
        },
        "date-networkdays-workweek-alloff" => {
            "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);;[.B1:.B7])".to_owned()
        },
        "date-networkdays-workweek-malformed" => {
            "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);;[.B1:.B7])".to_owned()
        },
        "date-workday-workweek-default" => {
            "=WORKDAY(DATE(2024;1;1);64;;[.B1:.B7])".to_owned()
        },
        "date-workday-workweek-custom" => {
            "=WORKDAY(DATE(2024;1;1);64;;[.B1:.B7])".to_owned()
        },
        "date-workday-workweek-alloff" => {
            "=WORKDAY(DATE(2024;1;1);1;;[.B1:.B7])".to_owned()
        },
        "date-workday-workweek-malformed" => {
            "=WORKDAY(DATE(2024;1;1);64;;[.B1:.B7])".to_owned()
        },
        "date-networkdays-interval-7" => "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;7))".to_owned(),
        "date-networkdays-interval-31" => "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;1;31))".to_owned(),
        "date-networkdays-interval-365" => "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;12;30))".to_owned(),
        "date-networkdays-interval-4096" => "=NETWORKDAYS(DATE(2024;1;1);DATE(2035;3;19))".to_owned(),
        _ if name.starts_with("date-workday-offset-") => {
            let offset = name.rsplit('-').next().unwrap();
            format!("=WORKDAY(DATE(2024;1;1);{offset})")
        },
        "date-sequence-networkdays-projection" => {
            "=IF({TRUE()|TRUE()};NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);[.A1:.A16];[.B1:.B7]);0)".to_owned()
        },
        "date-sequence-workday-projection" => {
            "=IF({TRUE()|TRUE()};WORKDAY(DATE(2024;1;1);64;[.A1:.A16];[.B1:.B7]);0)".to_owned()
        },
        "date-projected-day" => "=DAY({43890|43891})".to_owned(),
        "date-lazy-unselected" => {
            "=IF({FALSE()|FALSE()};NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);[.A1:.A256];[.B1:.B7]);0)".to_owned()
        },
        "date-list-refusal" => "=DAY([.A1:.A2]~[.A3:.A4])".to_owned(),
        "date-direct-sequence-refusal" => {
            "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);[.A1:.A2]~[.A3:.A4])".to_owned()
        },
        "date-cancel-networkdays" => {
            "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);[.A1:.A16])".to_owned()
        },
        "date-cancel-workday" => "=WORKDAY(DATE(2024;1;1);64;[.A1:.A16])".to_owned(),
        "date-resource-networkdays" => {
            "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);[.A1:.A16])".to_owned()
        },
        "date-work-limit-networkdays" => {
            "=NETWORKDAYS(DATE(2024;1;1);DATE(2035;3;19))".to_owned()
        },
        "date-formula-error-sequence" => {
            "=NETWORKDAYS(DATE(2024;1;1);DATE(2024;2;1);[.C1:.C16])".to_owned()
        },
        "date-dateparam-error" => "=DAY([.C1])".to_owned(),
        "date-missing-now" => "=NOW()".to_owned(),
        "date-missing-today" => "=TODAY()".to_owned(),
        "date-missing-eastersunday" => "=EASTERSUNDAY()".to_owned(),
        _ => panic!("unknown date case {name}"),
    }
}

fn controls() -> Vec<Case> {
    let mut cases = Vec::new();
    let mut push = |name: &str,
                    source: &str,
                    shape: Shape,
                    path: Path,
                    operation: ControlOp,
                    expectation: Expected,
                    read_bound: u64,
                    mode: ResolverMode| {
        cases.push(Case {
            name: name.to_owned(),
            source: source.to_owned(),
            shape,
            path,
            operation: Some(operation),
            date: None,
            expectation,
            read_bound,
            timestamp: false,
            max_reference_cells: None,
            max_steps: None,
            cancellation: false,
            resolver_mode: mode,
        });
    };
    push(
        "scalar-control-arithmetic",
        "=0.17+1.25",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Arithmetic,
        Expected::Number(1.42),
        0,
        ResolverMode::Control,
    );
    push(
        "scalar-control-sin",
        "=SIN(0.17)",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Sin,
        Expected::Number(0.17_f64.sin()),
        0,
        ResolverMode::Control,
    );
    push(
        "scalar-control-imsum",
        "=IMREAL(IMSUM(COMPLEX(2;3);COMPLEX(1;4)))",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::ImSum,
        Expected::Number(3.0),
        0,
        ResolverMode::Control,
    );
    push(
        "scalar-control-average",
        "=AVERAGE(1.25;2.5;3.75)",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Average,
        Expected::Number(2.5),
        0,
        ResolverMode::Control,
    );
    push(
        "scalar-control-counta",
        "=COUNTA(1.25;\"x\";TRUE())",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::CountA,
        Expected::Number(3.0),
        0,
        ResolverMode::Control,
    );
    push(
        "scalar-control-var",
        "=VAR(1.25;2.5;3.75)",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Var,
        Expected::Number(1.5625),
        0,
        ResolverMode::Control,
    );
    push(
        "scalar-control-stdev",
        "=STDEV(1.25;2.5;3.75)",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Stdev,
        Expected::Number(1.25),
        0,
        ResolverMode::Control,
    );
    for (name, source, op, value) in [
        (
            "database-control-dsum",
            "=DSUM([.A1:.B4];2;[.D1:.D2])",
            ControlOp::DSum,
            12.0,
        ),
        (
            "database-control-dvar",
            "=DVAR([.A1:.B4];2;[.D1:.D2])",
            ControlOp::DVar,
            9.0,
        ),
        (
            "database-control-dstdev",
            "=DSTDEV([.A1:.B4];2;[.D1:.D2])",
            ControlOp::DStdev,
            3.0,
        ),
    ] {
        push(
            name,
            source,
            Shape::Scalar,
            Path::Value,
            op,
            Expected::Number(value),
            7,
            ResolverMode::Database,
        );
    }
    let array_4 = literal_control_array(4, 4);
    push(
        "array-control-4x4-arithmetic",
        &format!("={array_4}+1.25"),
        Shape::Array {
            rows: 4,
            columns: 4,
            reference: false,
        },
        Path::Value,
        ControlOp::Arithmetic,
        Expected::NumberAny,
        0,
        ResolverMode::Control,
    );
    push(
        "array-control-4x4-sin",
        &format!("=SIN({array_4})"),
        Shape::Array {
            rows: 4,
            columns: 4,
            reference: false,
        },
        Path::Value,
        ControlOp::Sin,
        Expected::NumberAny,
        0,
        ResolverMode::Control,
    );
    let array_16 = literal_control_array(16, 16);
    push(
        "array-control-16x16-arithmetic",
        &format!("={array_16}+1.25"),
        Shape::Array {
            rows: 16,
            columns: 16,
            reference: false,
        },
        Path::Value,
        ControlOp::Arithmetic,
        Expected::NumberAny,
        0,
        ResolverMode::Control,
    );
    push(
        "array-control-16x16-sin",
        &format!("=SIN({array_16})"),
        Shape::Array {
            rows: 16,
            columns: 16,
            reference: false,
        },
        Path::Value,
        ControlOp::Sin,
        Expected::NumberAny,
        0,
        ResolverMode::Control,
    );
    push(
        "reference-array-16x4-arithmetic",
        "=[.A1:.D16]+1.25",
        Shape::Array {
            rows: 16,
            columns: 4,
            reference: true,
        },
        Path::Value,
        ControlOp::Arithmetic,
        Expected::NumberAny,
        64,
        ResolverMode::Control,
    );
    push(
        "scalar-aggregate-sum",
        "=SUM(1.25)",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Arithmetic,
        Expected::Number(1.25),
        0,
        ResolverMode::Control,
    );
    push(
        "literal-aggregate-4x1-sum",
        "=SUM({1|2|3|4})",
        Shape::Array {
            rows: 1,
            columns: 1,
            reference: false,
        },
        Path::Value,
        ControlOp::Arithmetic,
        Expected::Number(10.0),
        0,
        ResolverMode::Control,
    );
    push(
        "reference-aggregate-64x4-sum",
        "=SUM([.A1:.D64])",
        Shape::Scalar,
        Path::Value,
        ControlOp::Arithmetic,
        Expected::Number(expected_reference_sum(64, 4)),
        256,
        ResolverMode::Control,
    );
    push(
        "reference-conditional-256x4-sumifs",
        "=SUMIFS([.I1:.L256];[.A1:.D256];\">=2\";[.E1:.H256];1)",
        Shape::Scalar,
        Path::Value,
        ControlOp::Arithmetic,
        Expected::Number(conditional_sum(256, 4)),
        1792,
        ResolverMode::Conditional,
    );
    push(
        "reference-control-average",
        "=AVERAGE([.A1:.D64])",
        Shape::Scalar,
        Path::Value,
        ControlOp::Average,
        Expected::Number(expected_reference_sum(64, 4) / 256.0),
        256,
        ResolverMode::Control,
    );
    push(
        "reference-control-counta",
        "=COUNTA([.E1:.H64])",
        Shape::Scalar,
        Path::Value,
        ControlOp::CountA,
        Expected::Number(256.0),
        256,
        ResolverMode::Control,
    );
    push(
        "representative-median",
        "=MEDIAN(1;2;4;8)",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Median,
        Expected::Number(3.0),
        0,
        ResolverMode::Control,
    );
    push(
        "representative-rank",
        "=RANK(8;8;1)",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Rank,
        Expected::Number(1.0),
        0,
        ResolverMode::Control,
    );
    push(
        "representative-percentrank",
        "=PERCENTRANK(8;8;3)",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::PercentRank,
        Expected::Number(1.0),
        0,
        ResolverMode::Control,
    );
    push(
        "concat-borrowed-literals",
        "=\"left\"&\"right\"",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Concatenate,
        Expected::Text("leftright"),
        0,
        ResolverMode::Control,
    );
    push(
        "concat-owned-left",
        "=(\"left\"&\"mid\")&\"right\"",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Concatenate,
        Expected::Text("leftmidright"),
        0,
        ResolverMode::Control,
    );
    push(
        "concat-owned-right",
        "\"left\"&(\"mid\"&\"right\")",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Concatenate,
        Expected::Text("leftmidright"),
        0,
        ResolverMode::Control,
    );
    push(
        "concat-growth-chain",
        "=\"a\"&\"b\"&\"c\"&\"d\"&\"e\"&\"f\"&\"g\"&\"h\"",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::Concatenate,
        Expected::Text("abcdefgh"),
        0,
        ResolverMode::Control,
    );
    push(
        "value-date-fraction-value-date",
        "=VALUE(\"2024-01-02\")",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::ValueDate,
        Expected::Number(45_293.0),
        0,
        ResolverMode::Control,
    );
    push(
        "value-date-fraction-value-time",
        "=VALUE(\"12:34:56\")",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::ValueTime,
        Expected::Number(45_296.0 / 86_400.0),
        0,
        ResolverMode::Control,
    );
    push(
        "value-date-fraction-value-mixed-fraction",
        "=VALUE(\"3 1/2\")",
        Shape::Scalar,
        Path::Scalar,
        ControlOp::ValueMixedFraction,
        Expected::Number(3.5),
        0,
        ResolverMode::Control,
    );
    for (name, source, op, expected) in [
        (
            "lazy-if-cache-isnumber",
            "=IF({TRUE()|TRUE()};ISNUMBER([.A3:.A4]);FALSE())",
            ControlOp::LazyIsNumber,
            [true, false],
        ),
        (
            "lazy-if-cache-isblank",
            "=IF({TRUE()|TRUE()};ISBLANK([.A1:.A2]);FALSE())",
            ControlOp::LazyIsBlank,
            [true, false],
        ),
        (
            "lazy-if-cache-istext",
            "=IF({TRUE()|TRUE()};ISTEXT([.A1:.A2]);FALSE())",
            ControlOp::LazyIsText,
            [false, true],
        ),
    ] {
        let _ = expected;
        push(
            name,
            source,
            Shape::Array {
                rows: 2,
                columns: 1,
                reference: true,
            },
            Path::Value,
            op,
            Expected::LogicalAny,
            2,
            ResolverMode::Lazy,
        );
    }
    cases
}

fn expected_reference_sum(rows: usize, columns: usize) -> f64 {
    (0..rows)
        .flat_map(|row| (0..columns).map(move |column| aggregate_value(row, column)))
        .sum()
}

fn literal_control_array(rows: usize, columns: usize) -> String {
    let mut output = String::with_capacity(rows.saturating_mul(columns).saturating_mul(8) + 2);
    output.push('{');
    for row in 0..rows {
        if row != 0 {
            output.push('|');
        }
        for column in 0..columns {
            if column != 0 {
                output.push(';');
            }
            let value = literal_control_value(row.saturating_mul(columns).saturating_add(column));
            let _ = write!(output, "{value:.6}");
        }
    }
    output.push('}');
    output
}

fn conditional_sum(rows: usize, columns: usize) -> f64 {
    let mut total = 0.0;
    for row in 0..rows {
        for column in 0..columns {
            let index = row * columns + column;
            if index % 4 >= 2 && index % 2 == 1 {
                total += 0.5 + index as f64 * 0.03125;
            }
        }
    }
    total
}

fn date_case(
    name: &str,
    operation: DateOp,
    lane: DateLane,
    source: String,
    shape: Shape,
    path: Path,
    expectation: Expected,
    read_bound: u64,
    timestamp: bool,
    resolver_mode: ResolverMode,
    max_reference_cells: Option<usize>,
    max_steps: Option<u64>,
    cancellation: bool,
) -> Case {
    Case {
        name: name.to_owned(),
        source,
        shape,
        path,
        operation: None,
        date: Some((operation, lane)),
        expectation,
        read_bound,
        timestamp,
        max_reference_cells,
        max_steps,
        cancellation,
        resolver_mode,
    }
}

fn date_cases() -> Vec<Case> {
    let mut cases = Vec::new();
    let core = [
        ("date-core-date", DateOp::Date, "=DATE(2020;13;1)", false),
        (
            "date-core-datedif",
            DateOp::Datedif,
            "=DATEDIF(DATE(2020;1;31);DATE(2021;3;1);\"Y\")",
            false,
        ),
        (
            "date-core-datevalue",
            DateOp::DateValue,
            "=DATEVALUE(\"2024-03-31T12:34:56\")",
            false,
        ),
        (
            "date-core-day",
            DateOp::Day,
            "=DAY(DATE(2024;3;31)+0.75)",
            false,
        ),
        (
            "date-core-days",
            DateOp::Days,
            "=DAYS(DATE(2024;4;1)+0.5;DATE(2024;3;31)+0.5)",
            false,
        ),
        (
            "date-core-days360",
            DateOp::Days360,
            "=DAYS360(DATE(2024;1;31);DATE(2024;2;29))",
            false,
        ),
        (
            "date-core-eastersunday",
            DateOp::EasterSunday,
            "=EASTERSUNDAY(2024)",
            false,
        ),
        (
            "date-core-edate",
            DateOp::Edate,
            "=EDATE(DATE(2024;1;31);1)",
            false,
        ),
        (
            "date-core-eomonth",
            DateOp::Eomonth,
            "=EOMONTH(DATE(2024;2;10);0)",
            false,
        ),
        ("date-core-hour", DateOp::Hour, "=HOUR(0.75)", false),
        (
            "date-core-isoweeknum",
            DateOp::IsoWeeknum,
            "=ISOWEEKNUM(DATE(2024;1;1))",
            false,
        ),
        ("date-core-minute", DateOp::Minute, "=MINUTE(-1/256)", false),
        (
            "date-core-month",
            DateOp::Month,
            "=MONTH(DATE(2024;3;31)+0.5)",
            false,
        ),
        (
            "date-core-networkdays",
            DateOp::Networkdays,
            "=NETWORKDAYS(DATE(2024;3;29);DATE(2024;4;1))",
            false,
        ),
        ("date-core-now", DateOp::Now, "=NOW()", true),
        ("date-core-second", DateOp::Second, "=SECOND(-1/256)", false),
        ("date-core-time", DateOp::Time, "=TIME(12;30;0)", false),
        (
            "date-core-timevalue",
            DateOp::TimeValue,
            "=TIMEVALUE(\"12:30:00\")",
            false,
        ),
        ("date-core-today", DateOp::Today, "=TODAY()", true),
        (
            "date-core-weekday",
            DateOp::Weekday,
            "=WEEKDAY(DATE(2024;1;1);2)",
            false,
        ),
        (
            "date-core-weeknum",
            DateOp::Weeknum,
            "=WEEKNUM(DATE(2024;1;1);21)",
            false,
        ),
        (
            "date-core-workday",
            DateOp::Workday,
            "=WORKDAY(DATE(2024;3;29);1)",
            false,
        ),
        (
            "date-core-year",
            DateOp::Year,
            "=YEAR(DATE(2024;3;31)+0.5)",
            false,
        ),
        (
            "date-core-yearfrac",
            DateOp::Yearfrac,
            "=YEARFRAC(DATE(2024;1;1);DATE(2025;1;1);3)",
            false,
        ),
    ];
    for (name, operation, source, timestamp) in core {
        cases.push(date_case(
            name,
            operation,
            DateLane::Core,
            source.to_owned(),
            Shape::Scalar,
            Path::Scalar,
            Expected::Number(date_expected_name(name)),
            0,
            timestamp,
            ResolverMode::DateNormal,
            None,
            None,
            false,
        ));
    }

    for (name, operation, source) in [
        (
            "date-parser-datevalue-iso",
            DateOp::DateValue,
            "=DATEVALUE(\"2024-03-31\")",
        ),
        (
            "date-parser-datevalue-datetime",
            DateOp::DateValue,
            "=DATEVALUE(\"2024-03-31 12:34:56\")",
        ),
        (
            "date-parser-datevalue-enus",
            DateOp::DateValue,
            "=DATEVALUE(\"3/31/2024\")",
        ),
        (
            "date-parser-datevalue-month",
            DateOp::DateValue,
            "=DATEVALUE(\"Mar 31, 2024\")",
        ),
        (
            "date-parser-timevalue-clock",
            DateOp::TimeValue,
            "=TIMEVALUE(\"12:30:00\")",
        ),
        (
            "date-parser-timevalue-datetime",
            DateOp::TimeValue,
            "=TIMEVALUE(\"2020-01-02T03:04:05\")",
        ),
        (
            "date-parser-timevalue-fraction",
            DateOp::TimeValue,
            "=TIMEVALUE(\"12:30:00.5\")",
        ),
        (
            "date-parser-datevalue-number",
            DateOp::DateValue,
            "=DATEVALUE(\"123\")",
        ),
        (
            "date-parser-timevalue-number",
            DateOp::TimeValue,
            "=TIMEVALUE(\"0.5\")",
        ),
    ] {
        cases.push(date_case(
            name,
            operation,
            DateLane::Parser,
            source.to_owned(),
            Shape::Scalar,
            Path::Scalar,
            Expected::NumberAny,
            0,
            false,
            ResolverMode::DateNormal,
            None,
            None,
            false,
        ));
    }

    for (name, operation, source) in [
        (
            "date-boundary-date-rollover",
            DateOp::Date,
            "=DATE(2020;2;30)",
        ),
        ("date-boundary-date-1900", DateOp::Date, "=DATE(1900;3;1)"),
        (
            "date-boundary-days360-eu",
            DateOp::Days360,
            "=DAYS360(DATE(2024;1;31);DATE(2024;2;29);TRUE())",
        ),
        (
            "date-boundary-edate-clamp",
            DateOp::Edate,
            "=EDATE(DATE(2024;1;31);1)",
        ),
        (
            "date-boundary-eomonth-clamp",
            DateOp::Eomonth,
            "=EOMONTH(DATE(2024;2;10);0)",
        ),
        (
            "date-boundary-yearfrac-basis1",
            DateOp::Yearfrac,
            "=YEARFRAC(DATE(2024;1;1);DATE(2025;1;1);1)",
        ),
    ] {
        cases.push(date_case(
            name,
            operation,
            DateLane::Boundary,
            source.to_owned(),
            Shape::Scalar,
            Path::Scalar,
            Expected::NumberAny,
            0,
            false,
            ResolverMode::DateNormal,
            None,
            None,
            false,
        ));
    }

    cases.push(date_case(
        "date-boundary-weeknum-noninteger",
        DateOp::Weeknum,
        DateLane::Boundary,
        "=WEEKNUM(44199;21.5)".to_owned(),
        Shape::Scalar,
        Path::Scalar,
        Expected::Error(ScalarError::Number),
        0,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        false,
    ));
    cases.push(date_case(
        "date-boundary-weekday-17",
        DateOp::Weekday,
        DateLane::Boundary,
        "=WEEKDAY(DATE(2024;1;7);17)".to_owned(),
        Shape::Scalar,
        Path::Scalar,
        Expected::NumberAny,
        0,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        false,
    ));

    for count in [0_usize, 1, 16, 64, 256] {
        let name = format!("date-networkdays-holidays-{count}");
        cases.push(date_case(
            &name,
            DateOp::Networkdays,
            DateLane::HolidayScale,
            date_source(&name),
            Shape::Scalar,
            Path::Value,
            Expected::NumberAny,
            count as u64,
            false,
            ResolverMode::DateNormal,
            None,
            None,
            false,
        ));
        let name = format!("date-workday-holidays-{count}");
        cases.push(date_case(
            &name,
            DateOp::Workday,
            DateLane::HolidayScale,
            date_source(&name),
            Shape::Scalar,
            Path::Value,
            Expected::NumberAny,
            count as u64,
            false,
            ResolverMode::DateNormal,
            None,
            None,
            false,
        ));
    }

    for (name, operation, source, mode, expected, reads) in [
        (
            "date-networkdays-workweek-default",
            DateOp::Networkdays,
            date_source("date-networkdays-workweek-default"),
            ResolverMode::DateNormal,
            Expected::NumberAny,
            7,
        ),
        (
            "date-networkdays-workweek-custom",
            DateOp::Networkdays,
            date_source("date-networkdays-workweek-custom"),
            ResolverMode::DateAllWorkdays,
            Expected::NumberAny,
            7,
        ),
        (
            "date-networkdays-workweek-alloff",
            DateOp::Networkdays,
            date_source("date-networkdays-workweek-alloff"),
            ResolverMode::DateAllOff,
            Expected::Number(0.0),
            7,
        ),
        (
            "date-networkdays-workweek-malformed",
            DateOp::Networkdays,
            date_source("date-networkdays-workweek-malformed"),
            ResolverMode::DateMalformedWorkweek,
            Expected::Error(ScalarError::Value),
            7,
        ),
        (
            "date-workday-workweek-default",
            DateOp::Workday,
            date_source("date-workday-workweek-default"),
            ResolverMode::DateNormal,
            Expected::NumberAny,
            7,
        ),
        (
            "date-workday-workweek-custom",
            DateOp::Workday,
            date_source("date-workday-workweek-custom"),
            ResolverMode::DateAllWorkdays,
            Expected::NumberAny,
            7,
        ),
        (
            "date-workday-workweek-alloff",
            DateOp::Workday,
            date_source("date-workday-workweek-alloff"),
            ResolverMode::DateAllOff,
            Expected::Error(ScalarError::Number),
            7,
        ),
        (
            "date-workday-workweek-malformed",
            DateOp::Workday,
            date_source("date-workday-workweek-malformed"),
            ResolverMode::DateMalformedWorkweek,
            Expected::Error(ScalarError::Value),
            7,
        ),
    ] {
        cases.push(date_case(
            name,
            operation,
            DateLane::WorkweekScale,
            source,
            Shape::Scalar,
            Path::Value,
            expected,
            reads,
            false,
            mode,
            None,
            None,
            false,
        ));
    }

    for span in [7_usize, 31, 365, 4096] {
        let name = format!("date-networkdays-interval-{span}");
        cases.push(date_case(
            &name,
            DateOp::Networkdays,
            DateLane::IntervalScale,
            date_source(&name),
            Shape::Scalar,
            Path::Scalar,
            Expected::NumberAny,
            0,
            false,
            ResolverMode::DateNormal,
            None,
            None,
            false,
        ));
    }
    for offset in [0_i64, 1, 64, 256, 1024] {
        let name = format!("date-workday-offset-{offset}");
        cases.push(date_case(
            &name,
            DateOp::Workday,
            DateLane::IntervalScale,
            date_source(&name),
            Shape::Scalar,
            Path::Scalar,
            Expected::NumberAny,
            0,
            false,
            ResolverMode::DateNormal,
            None,
            None,
            false,
        ));
    }

    cases.push(date_case(
        "date-sequence-networkdays-projection",
        DateOp::Networkdays,
        DateLane::SequenceProjection,
        date_source("date-sequence-networkdays-projection"),
        Shape::Array {
            rows: 2,
            columns: 1,
            reference: true,
        },
        Path::Value,
        Expected::NumberAny,
        46,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        false,
    ));
    cases.push(date_case(
        "date-sequence-workday-projection",
        DateOp::Workday,
        DateLane::SequenceProjection,
        date_source("date-sequence-workday-projection"),
        Shape::Array {
            rows: 2,
            columns: 1,
            reference: true,
        },
        Path::Value,
        Expected::NumberAny,
        46,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        false,
    ));
    cases.push(date_case(
        "date-projected-day",
        DateOp::Day,
        DateLane::Projected,
        date_source("date-projected-day"),
        Shape::Array {
            rows: 2,
            columns: 1,
            reference: false,
        },
        Path::Value,
        Expected::NumberAny,
        0,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        false,
    ));
    cases.push(date_case(
        "date-lazy-unselected",
        DateOp::Networkdays,
        DateLane::Lazy,
        date_source("date-lazy-unselected"),
        Shape::Array {
            rows: 2,
            columns: 1,
            reference: true,
        },
        Path::Value,
        Expected::NumberAny,
        0,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        false,
    ));

    for (name, source, operation) in [
        ("date-timestamp-now", "=NOW()", DateOp::Now),
        ("date-timestamp-today", "=TODAY()", DateOp::Today),
        (
            "date-timestamp-eastersunday",
            "=EASTERSUNDAY()",
            DateOp::EasterSunday,
        ),
    ] {
        cases.push(date_case(
            name,
            operation,
            DateLane::Timestamp,
            source.to_owned(),
            Shape::Scalar,
            Path::Scalar,
            Expected::NumberAny,
            0,
            true,
            ResolverMode::DateNormal,
            None,
            None,
            false,
        ));
    }
    for (name, source, operation) in [
        ("date-missing-now", "=NOW()", DateOp::Now),
        ("date-missing-today", "=TODAY()", DateOp::Today),
        (
            "date-missing-eastersunday",
            "=EASTERSUNDAY()",
            DateOp::EasterSunday,
        ),
    ] {
        cases.push(date_case(
            name,
            operation,
            DateLane::MissingTimestamp,
            source.to_owned(),
            Shape::Scalar,
            Path::Scalar,
            Expected::Failure(FailureKind::CalculationClock),
            0,
            false,
            ResolverMode::DateNormal,
            None,
            None,
            false,
        ));
    }

    cases.push(date_case(
        "date-list-refusal",
        DateOp::Day,
        DateLane::Refusal,
        date_source("date-list-refusal"),
        Shape::Scalar,
        Path::Value,
        Expected::Error(ScalarError::Value),
        0,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        false,
    ));
    cases.push(date_case(
        "date-direct-sequence-refusal",
        DateOp::Networkdays,
        DateLane::Refusal,
        date_source("date-direct-sequence-refusal"),
        Shape::Scalar,
        Path::Value,
        Expected::Error(ScalarError::Value),
        0,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        false,
    ));
    cases.push(date_case(
        "date-cancel-networkdays",
        DateOp::Networkdays,
        DateLane::Cancellation,
        date_source("date-cancel-networkdays"),
        Shape::Scalar,
        Path::Value,
        Expected::Failure(FailureKind::Cancelled),
        1,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        true,
    ));
    cases.push(date_case(
        "date-cancel-workday",
        DateOp::Workday,
        DateLane::Cancellation,
        date_source("date-cancel-workday"),
        Shape::Scalar,
        Path::Value,
        Expected::Failure(FailureKind::Cancelled),
        1,
        false,
        ResolverMode::DateNormal,
        None,
        None,
        true,
    ));
    cases.push(date_case(
        "date-resource-networkdays",
        DateOp::Networkdays,
        DateLane::Resource,
        date_source("date-resource-networkdays"),
        Shape::Scalar,
        Path::Value,
        Expected::Failure(FailureKind::ReferenceCells),
        0,
        false,
        ResolverMode::DateNormal,
        Some(0),
        None,
        false,
    ));
    cases.push(date_case(
        "date-work-limit-networkdays",
        DateOp::Networkdays,
        DateLane::WorkLimit,
        date_source("date-work-limit-networkdays"),
        Shape::Scalar,
        Path::Scalar,
        Expected::Failure(FailureKind::Work),
        0,
        false,
        ResolverMode::DateNormal,
        None,
        Some(64),
        false,
    ));
    cases.push(date_case(
        "date-formula-error-sequence",
        DateOp::Networkdays,
        DateLane::FormulaError,
        date_source("date-formula-error-sequence"),
        Shape::Scalar,
        Path::Value,
        Expected::Error(ScalarError::NotAvailable),
        16,
        false,
        ResolverMode::DateErrorCell,
        None,
        None,
        false,
    ));
    cases.push(date_case(
        "date-dateparam-error",
        DateOp::Day,
        DateLane::FormulaError,
        date_source("date-dateparam-error"),
        Shape::Array {
            rows: 1,
            columns: 1,
            reference: false,
        },
        Path::Value,
        Expected::Error(ScalarError::NotAvailable),
        1,
        false,
        ResolverMode::DateErrorCell,
        None,
        None,
        false,
    ));
    cases
}

fn date_expected_name(name: &str) -> f64 {
    let case = Case {
        name: name.to_owned(),
        source: String::new(),
        shape: Shape::Scalar,
        path: Path::Scalar,
        operation: None,
        date: None,
        expectation: Expected::NumberAny,
        read_bound: 0,
        timestamp: false,
        max_reference_cells: None,
        max_steps: None,
        cancellation: false,
        resolver_mode: ResolverMode::DateNormal,
    };
    date_expected(&case, 0).unwrap()
}

fn all_cases() -> Vec<Case> {
    let mut cases = controls();
    cases.extend(date_cases());
    cases
}

fn date_options(case: &Case) -> EvaluationOptions {
    #[cfg(feature = "date-time-candidate")]
    if case.timestamp {
        let serial = if case.name == "date-timestamp-eastersunday" {
            // The independent oracle deliberately places the snapshot after
            // Easter 2024, exercising the omitted-year next-year rule.
            45_383.25
        } else {
            TIMESTAMP_SERIAL
        };
        let timestamp = CalculationTimestamp::from_serial(serial).expect("valid profile timestamp");
        return EvaluationOptions::default().with_calculation_timestamp(timestamp);
    }
    let _ = case;
    EvaluationOptions::default()
}

fn execution() -> (ExecutionContext, CancellationSource) {
    let budget = Budget::root(
        "ods-formula-date-time-performance",
        CoreLimits::for_profile(Profile::Server),
    );
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(1_u64 << 40).expect("finite in-flight bytes"),
        0,
    )
    .expect("valid execution limits");
    (ExecutionContext::new(budget, token, limits), cancellation)
}

fn value_limits(case: &Case) -> Limits {
    let mut limits = Limits::default();
    if let Some(value) = case.max_reference_cells {
        limits = limits.with_max_reference_cells(value);
    }
    if let Some(value) = case.max_steps {
        limits = limits.with_max_steps(value);
    }
    limits
}

fn scalar_limits(case: &Case) -> EvaluationLimits {
    let mut limits = EvaluationLimits::default();
    if let Some(value) = case.max_steps {
        limits = limits.with_max_steps(value);
    }
    limits
}

fn eval_scalar<'a>(
    case: &Case,
    expression: &'a Expression,
    execution: &ExecutionContext,
) -> Result<litchi_ods::codec::formula::evaluation::EvaluatedScalar<'a>, EvaluationFailure> {
    let context = EvaluationContext::with_options(execution, date_options(case));
    evaluate_scalar(expression, &context, &scalar_limits(case))
}

fn eval_value<'a>(
    case: &Case,
    expression: &'a Expression,
    resolver: &'a FixtureResolver,
    execution: &ExecutionContext,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0))
        .with_mode(case.shape.mode())
        .with_options(date_options(case));
    value::evaluate(expression, resolver, &context, &value_limits(case))
}

fn expected_control_array(case: &Case, index: usize) -> Option<f64> {
    match case.name.as_str() {
        "array-control-4x4-arithmetic" | "array-control-16x16-arithmetic" => {
            Some(literal_control_value(index) + 1.25)
        },
        "array-control-4x4-sin" | "array-control-16x16-sin" => {
            Some(literal_control_value(index).sin())
        },
        "reference-array-16x4-arithmetic" => {
            let row = index / 4;
            let column = index % 4;
            Some(aggregate_value(row, column) + 1.25)
        },
        _ => None,
    }
}

fn literal_control_value(index: usize) -> f64 {
    0.17 + (index % 9) as f64 * 0.047
}

fn expected_control_logical(case: &Case, index: usize) -> Option<bool> {
    match case.name.as_str() {
        "lazy-if-cache-isnumber" | "lazy-if-cache-isblank" => Some(index == 0),
        "lazy-if-cache-istext" => Some(index == 1),
        _ => None,
    }
}

fn approximately_equal(actual: f64, expected: f64) -> bool {
    (actual - expected).abs() <= expected.abs().max(1.0) * 1.0e-10
}

fn validate_failure(case: &Case, error: &EvaluationFailure) -> AnyResult<()> {
    match (case.expectation, error) {
        (
            Expected::Failure(FailureKind::ReferenceCells),
            EvaluationFailure::ResourceLimit(limit),
        ) if limit.resource == Resource::Objects => Ok(()),
        (Expected::Failure(FailureKind::Work), EvaluationFailure::ResourceLimit(limit))
            if limit.resource == Resource::Work =>
        {
            Ok(())
        },
        (Expected::Failure(FailureKind::Cancelled), EvaluationFailure::Cancelled) => Ok(()),
        #[cfg(feature = "date-time-candidate")]
        (
            Expected::Failure(FailureKind::CalculationClock),
            EvaluationFailure::Unsupported(UnsupportedKind::CalculationClock),
        ) => Ok(()),
        (Expected::Failure(expected), observed) => Err(format!(
            "{} returned {observed}, expected evaluator failure {expected:?}",
            case.name
        )
        .into()),
        (_, observed) => {
            Err(format!("{} returned unexpected failure {observed}", case.name).into())
        },
    }
}

fn failure_checksum(expectation: Expected) -> u64 {
    match expectation {
        Expected::Failure(FailureKind::ReferenceCells) => 0xf1f1_f1f1_f1f1_f1f1,
        Expected::Failure(FailureKind::Work) => 0xaaaa_aaaa_aaaa_aaaa,
        Expected::Failure(FailureKind::Cancelled) => 0xcaca_caca_caca_caca,
        Expected::Failure(FailureKind::CalculationClock) => 0xcccc_cccc_cccc_cccc,
        _ => 0,
    }
}

fn scalar_checksum(value: &ScalarValue<'_>) -> AnyResult<u64> {
    match value {
        ScalarValue::Number(value) => Ok(value.to_bits().rotate_left(11)),
        ScalarValue::Logical(value) => Ok(u64::from(*value).rotate_left(11)),
        ScalarValue::Text(value) => Ok(text_checksum(value.as_ref())),
        ScalarValue::Error(error) => Ok((*error as u64).rotate_left(11)),
        other => Err(format!("unexpected scalar {other:?}").into()),
    }
}

fn text_checksum(value: &str) -> u64 {
    value.bytes().fold(0xcbf29ce484222325, |hash, byte| {
        hash.wrapping_mul(0x100000001b3)
            .wrapping_add(u64::from(byte))
    })
}

fn value_checksum(value: Value<'_>) -> AnyResult<u64> {
    match value {
        Value::Number(value) => Ok(value.to_bits().rotate_left(11)),
        Value::Logical(value) => Ok(u64::from(value).rotate_left(11)),
        Value::Text(value) => Ok(text_checksum(value)),
        Value::Error(error) => Ok((error as u64).rotate_left(11)),
        other => Err(format!("unexpected value {other:?}").into()),
    }
}

fn array_checksum(array: value::ArrayView<'_>) -> AnyResult<u64> {
    let mut checksum = ((array.shape().rows() as u64) << 32) ^ array.shape().columns() as u64;
    for index in 0..array.len() {
        checksum = checksum.rotate_left(5)
            ^ value_checksum(array.get(index).ok_or("missing array cell")?)?;
    }
    Ok(checksum)
}

fn validate_scalar(case: &Case, value: &ScalarValue<'_>) -> AnyResult<()> {
    match (case.expectation, value) {
        (Expected::Number(expected), ScalarValue::Number(actual))
            if approximately_equal(*actual, expected) =>
        {
            Ok(())
        },
        (Expected::NumberAny, ScalarValue::Number(_)) => Ok(()),
        (Expected::Logical(expected), ScalarValue::Logical(actual)) if expected == *actual => {
            Ok(())
        },
        (Expected::LogicalAny, ScalarValue::Logical(_)) => Ok(()),
        (Expected::Text(expected), ScalarValue::Text(actual)) if expected == actual => Ok(()),
        (Expected::Error(expected), ScalarValue::Error(actual)) if expected == *actual => Ok(()),
        (Expected::Failure(expected), observed) => Err(format!(
            "{} returned {observed:?}, expected failure {expected:?}",
            case.name
        )
        .into()),
        (expected, observed) => {
            Err(format!("{} returned {observed:?}, expected {expected:?}", case.name).into())
        },
    }
}

fn validate_date_value(case: &Case, result: &Evaluated<'_>) -> AnyResult<()> {
    match case.expectation {
        Expected::Error(expected) => {
            if case.shape.is_scalar() {
                if matches!(result.value(), Value::Error(actual) if actual == expected) {
                    return Ok(());
                }
            } else {
                let array = result
                    .as_array()
                    .ok_or("date error result was not an array")?;
                let (rows, columns) = case.shape.dimensions().ok_or("missing error dimensions")?;
                if (array.shape().rows(), array.shape().columns()) != (rows, columns) {
                    return Err(format!("{} returned wrong error shape", case.name).into());
                }
                if array.len() == 1
                    && matches!(array.get(0), Some(Value::Error(actual)) if actual == expected)
                {
                    return Ok(());
                }
            }
            return Err(format!(
                "{} returned {:?}, expected {expected}",
                case.name,
                result.value()
            )
            .into());
        },
        Expected::Failure(expected) => {
            return Err(format!("{} produced a value, expected {expected:?}", case.name).into());
        },
        Expected::Number(expected) => {
            let actual = match result.value() {
                Value::Number(value) => value,
                other => return Err(format!("{} returned {other:?}", case.name).into()),
            };
            if approximately_equal(actual, expected) {
                return Ok(());
            }
            return Err(format!("{} returned {actual}, expected {expected}", case.name).into());
        },
        Expected::NumberAny => {},
        Expected::Logical(_) | Expected::LogicalAny | Expected::Text(_) => {
            return Err(format!("{} has invalid date expectation", case.name).into());
        },
    }
    if case.shape.is_scalar() {
        let actual = match result.value() {
            Value::Number(value) => value,
            other => return Err(format!("{} returned {other:?}", case.name).into()),
        };
        let expected = date_expected(case, 0).ok_or("missing date oracle value")?;
        if approximately_equal(actual, expected) {
            return Ok(());
        }
        return Err(format!("{} returned {actual}, expected {expected}", case.name).into());
    }
    let array = result.as_array().ok_or("date result was not an array")?;
    let (rows, columns) = case.shape.dimensions().ok_or("missing date dimensions")?;
    if (array.shape().rows(), array.shape().columns()) != (rows, columns) {
        return Err(format!("{} returned wrong shape", case.name).into());
    }
    for index in 0..array.len() {
        let actual = match array.get(index).ok_or("missing date result cell")? {
            Value::Number(value) => value,
            other => return Err(format!("{} cell {index}: {other:?}", case.name).into()),
        };
        let expected = date_expected(case, index).ok_or("missing array date oracle value")?;
        if !approximately_equal(actual, expected) {
            return Err(format!(
                "{} cell {index} returned {actual}, expected {expected}",
                case.name
            )
            .into());
        }
    }
    Ok(())
}

fn validate_control_value(case: &Case, result: &Evaluated<'_>) -> AnyResult<()> {
    match case.expectation {
        Expected::Number(expected) => {
            let actual = match result.value() {
                Value::Number(value) => value,
                other => return Err(format!("{} returned {other:?}", case.name).into()),
            };
            if approximately_equal(actual, expected) {
                Ok(())
            } else {
                Err(format!("{} returned {actual}, expected {expected}", case.name).into())
            }
        },
        Expected::NumberAny => {
            let array = result.as_array().ok_or("control expected array")?;
            for index in 0..array.len() {
                let actual = match array.get(index).ok_or("missing control cell")? {
                    Value::Number(value) => value,
                    other => return Err(format!("{} cell {index}: {other:?}", case.name).into()),
                };
                let expected =
                    expected_control_array(case, index).ok_or("missing control oracle")?;
                if !approximately_equal(actual, expected) {
                    return Err(format!(
                        "{} cell {index} returned {actual}, expected {expected}",
                        case.name
                    )
                    .into());
                }
            }
            Ok(())
        },
        Expected::LogicalAny => {
            let array = result.as_array().ok_or("control expected array")?;
            for index in 0..array.len() {
                let actual = match array.get(index).ok_or("missing control cell")? {
                    Value::Logical(value) => value,
                    other => return Err(format!("{} cell {index}: {other:?}", case.name).into()),
                };
                let expected = expected_control_logical(case, index)
                    .ok_or("missing logical control oracle")?;
                if actual != expected {
                    return Err(format!(
                        "{} cell {index} returned {actual}, expected {expected}",
                        case.name
                    )
                    .into());
                }
            }
            Ok(())
        },
        Expected::Text(expected) => match result.value() {
            Value::Text(actual) if actual == expected => Ok(()),
            other => Err(format!("{} returned {other:?}, expected {expected:?}", case.name).into()),
        },
        Expected::Error(expected) => match result.value() {
            Value::Error(actual) if actual == expected => Ok(()),
            other => Err(format!("{} returned {other:?}, expected {expected}", case.name).into()),
        },
        Expected::Logical(_) | Expected::Failure(_) => Err("invalid control expectation".into()),
    }
}

fn validate_value(case: &Case, result: &Evaluated<'_>) -> AnyResult<()> {
    if case.date.is_some() {
        validate_date_value(case, result)
    } else {
        validate_control_value(case, result)
    }
}

fn input_bytes(case: &Case) -> u64 {
    case.source.len() as u64
}

#[derive(Clone, Copy, Debug)]
struct Sample {
    elapsed_ns: u64,
    alloc_calls: u64,
    dealloc_calls: u64,
    requested_bytes: u64,
    released_bytes: u64,
    live_before: u64,
    live_after: u64,
    peak_live_delta: u64,
    work: u64,
    memory_retained: u64,
    reference_reads: u64,
    output_bytes: u64,
    checksum: u64,
}

#[derive(Clone, Copy, Debug)]
enum Phase {
    Evaluate,
    ParseEvaluate,
}

impl Phase {
    fn parse(value: &str) -> AnyResult<Self> {
        match value {
            "evaluate" => Ok(Self::Evaluate),
            "parse-evaluate" => Ok(Self::ParseEvaluate),
            other => Err(format!("unknown phase {other:?}").into()),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Evaluate => "evaluate",
            Self::ParseEvaluate => "parse-evaluate",
        }
    }
}

#[derive(Debug)]
struct Config {
    case: Option<String>,
    phase: Phase,
    warmups: usize,
    iterations: usize,
    repeat: Option<usize>,
    preflight_only: bool,
    describe: bool,
}

fn parse_config() -> AnyResult<Option<Config>> {
    let mut case = None;
    let mut phase = Phase::Evaluate;
    let mut warmups = 3;
    let mut iterations = 1;
    let mut repeat = None;
    let mut preflight_only = false;
    let mut describe = false;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                println!(
                    "usage: ods-formula-date-time-performance [--case NAME|all] [--phase evaluate|parse-evaluate] [--warmups N] [--iterations N] [--repeat N] [--preflight-only] [--describe]"
                );
                return Ok(None);
            },
            "--list" => {
                for case in all_cases() {
                    println!("{}", case.name);
                }
                return Ok(None);
            },
            "--case" => case = Some(arguments.next().ok_or("--case requires a name")?),
            "--phase" => {
                phase = Phase::parse(&arguments.next().ok_or("--phase requires a value")?)?
            },
            "--warmups" => {
                warmups = arguments
                    .next()
                    .ok_or("--warmups requires an integer")?
                    .parse()?
            },
            "--iterations" => {
                iterations = arguments
                    .next()
                    .ok_or("--iterations requires an integer")?
                    .parse()?
            },
            "--repeat" => {
                repeat = Some(
                    arguments
                        .next()
                        .ok_or("--repeat requires an integer")?
                        .parse()?,
                )
            },
            "--preflight-only" => preflight_only = true,
            "--describe" => describe = true,
            other => return Err(format!("unknown option {other:?}; use --help").into()),
        }
    }
    if iterations == 0 || repeat == Some(0) {
        return Err("iterations and repeat must be positive".into());
    }
    Ok(Some(Config {
        case,
        phase,
        warmups,
        iterations,
        repeat,
        preflight_only,
        describe,
    }))
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

fn consume_scalar(
    case: &Case,
    result: Result<litchi_ods::codec::formula::evaluation::EvaluatedScalar<'_>, EvaluationFailure>,
    execution: &ExecutionContext,
    checksum: &mut u64,
    output_bytes: &mut u64,
) -> AnyResult<u64> {
    let result = match result {
        Ok(result) => result,
        Err(error) => {
            validate_failure(case, &error)?;
            *checksum = checksum.wrapping_add(failure_checksum(case.expectation));
            black_box(*checksum);
            return Ok(execution.budget().used(Resource::Memory));
        },
    };
    if matches!(case.expectation, Expected::Failure(_)) {
        return Err(format!("{} unexpectedly produced a value", case.name).into());
    }
    validate_scalar(case, result.value())?;
    *checksum = checksum.wrapping_add(scalar_checksum(result.value())?);
    *output_bytes = output_bytes.saturating_add(scalar_output_bytes(result.value()));
    black_box(*checksum);
    let retained = execution.budget().used(Resource::Memory);
    drop(result);
    Ok(retained)
}

fn consume_value(
    result: Result<Evaluated<'_>, EvaluationFailure>,
    case: &Case,
    execution: &ExecutionContext,
    checksum: &mut u64,
    output_bytes: &mut u64,
) -> AnyResult<u64> {
    let result = match result {
        Ok(result) => result,
        Err(error) => {
            validate_failure(case, &error)?;
            *checksum = checksum.wrapping_add(failure_checksum(case.expectation));
            black_box(*checksum);
            return Ok(execution.budget().used(Resource::Memory));
        },
    };
    if matches!(case.expectation, Expected::Failure(_)) {
        return Err(format!("{} unexpectedly produced a value", case.name).into());
    }
    validate_value(case, &result)?;
    *checksum = checksum.wrapping_add(if case.shape.is_scalar() {
        value_checksum(result.value())?
    } else {
        array_checksum(result.as_array().ok_or("missing array checksum")?)?
    });
    *output_bytes = output_bytes.saturating_add(if case.shape.is_scalar() {
        value_output_bytes(result.value())
    } else {
        array_output_bytes(result.as_array().ok_or("missing array output")?)
    });
    black_box(*checksum);
    let retained = execution.budget().used(Resource::Memory);
    drop(result);
    Ok(retained)
}

fn scalar_output_bytes(value: &ScalarValue<'_>) -> u64 {
    match value {
        ScalarValue::Text(value) => value.len() as u64,
        _ => 0,
    }
}

fn value_output_bytes(value: Value<'_>) -> u64 {
    match value {
        Value::Text(value) => value.len() as u64,
        Value::Array(array) => array_output_bytes(array),
        _ => 0,
    }
}

fn array_output_bytes(array: value::ArrayView<'_>) -> u64 {
    (0..array.len())
        .filter_map(|index| array.get(index))
        .map(value_output_bytes)
        .sum()
}

fn measure_evaluate(case: &Case, expression: &Expression, repeat: usize) -> AnyResult<Sample> {
    let (execution, cancellation) = execution();
    let resolver = FixtureResolver::new(
        case.resolver_mode,
        case.cancellation.then_some(cancellation.clone()),
    );
    let scalar_context = EvaluationContext::with_options(&execution, date_options(case));
    let value_context = Context::new(&execution, Position::new("Main", 0, 0))
        .with_mode(case.shape.mode())
        .with_options(date_options(case));
    let scalar_limits = scalar_limits(case);
    let value_limits = value_limits(case);
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut checksum = 0_u64;
    let mut output_bytes = 0_u64;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        if case.path == Path::Scalar {
            match evaluate_scalar(expression, &scalar_context, &scalar_limits) {
                Ok(value) => {
                    if matches!(case.expectation, Expected::Failure(_)) {
                        return Err(format!("{} unexpectedly produced a value", case.name).into());
                    }
                    validate_scalar(case, value.value())?;
                    checksum = checksum.wrapping_add(scalar_checksum(value.value())?);
                    output_bytes = output_bytes.saturating_add(scalar_output_bytes(value.value()));
                    retained_peak = retained_peak.max(execution.budget().used(Resource::Memory));
                    drop(value);
                },
                Err(error) => {
                    validate_failure(case, &error)?;
                    checksum = checksum.wrapping_add(failure_checksum(case.expectation));
                },
            }
        } else {
            let value = value::evaluate(expression, &resolver, &value_context, &value_limits);
            retained_peak = retained_peak.max(consume_value(
                value,
                case,
                &execution,
                &mut checksum,
                &mut output_bytes,
            )?);
        }
    }
    Ok(sample_from(
        &execution,
        &resolver,
        baseline_work,
        baseline_memory,
        live_before,
        started,
        checksum,
        output_bytes,
        retained_peak,
    ))
}

fn measure_parse_evaluate(case: &Case, repeat: usize) -> AnyResult<Sample> {
    let (execution, cancellation) = execution();
    let resolver = FixtureResolver::new(
        case.resolver_mode,
        case.cancellation.then_some(cancellation.clone()),
    );
    let scalar_context = EvaluationContext::with_options(&execution, date_options(case));
    let value_context = Context::new(&execution, Position::new("Main", 0, 0))
        .with_mode(case.shape.mode())
        .with_options(date_options(case));
    let scalar_limits = scalar_limits(case);
    let value_limits = value_limits(case);
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut checksum = 0_u64;
    let mut output_bytes = 0_u64;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        let expression = Expression::parse(black_box(&case.source))
            .map_err(|error| format!("{} parse failed: {error}", case.name))?;
        if case.path == Path::Scalar {
            match evaluate_scalar(&expression, &scalar_context, &scalar_limits) {
                Ok(value) => {
                    if matches!(case.expectation, Expected::Failure(_)) {
                        return Err(format!("{} unexpectedly produced a value", case.name).into());
                    }
                    validate_scalar(case, value.value())?;
                    checksum = checksum.wrapping_add(scalar_checksum(value.value())?);
                    output_bytes = output_bytes.saturating_add(scalar_output_bytes(value.value()));
                    retained_peak = retained_peak.max(execution.budget().used(Resource::Memory));
                    drop(value);
                },
                Err(error) => {
                    validate_failure(case, &error)?;
                    checksum = checksum.wrapping_add(failure_checksum(case.expectation));
                },
            }
        } else {
            retained_peak = retained_peak.max(consume_value(
                value::evaluate(&expression, &resolver, &value_context, &value_limits),
                case,
                &execution,
                &mut checksum,
                &mut output_bytes,
            )?);
        }
        black_box(expression);
    }
    Ok(sample_from(
        &execution,
        &resolver,
        baseline_work,
        baseline_memory,
        live_before,
        started,
        checksum,
        output_bytes,
        retained_peak,
    ))
}

fn sample_from(
    execution: &ExecutionContext,
    resolver: &FixtureResolver,
    baseline_work: u64,
    baseline_memory: u64,
    live_before: u64,
    started: Instant,
    checksum: u64,
    output_bytes: u64,
    retained_peak: u64,
) -> Sample {
    Sample {
        elapsed_ns: started.elapsed().as_nanos().min(u64::MAX as u128) as u64,
        alloc_calls: ALLOC_CALLS.load(Ordering::Acquire),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
        requested_bytes: ALLOC_BYTES.load(Ordering::Acquire),
        released_bytes: DEALLOC_BYTES.load(Ordering::Acquire),
        live_before,
        live_after: LIVE_BYTES.load(Ordering::Acquire),
        peak_live_delta: PEAK_LIVE_BYTES
            .load(Ordering::Acquire)
            .saturating_sub(live_before),
        work: execution
            .budget()
            .used(Resource::Work)
            .saturating_sub(baseline_work),
        memory_retained: retained_peak.saturating_sub(baseline_memory),
        reference_reads: resolver.stats.reads(),
        output_bytes,
        checksum,
    }
}

fn preflight(case: &Case) -> AnyResult<u64> {
    let expression = Expression::parse(&case.source)
        .map_err(|error| format!("{} parse preflight failed: {error}", case.name))?;
    let (execution, cancellation) = execution();
    let resolver = FixtureResolver::new(
        case.resolver_mode,
        case.cancellation.then_some(cancellation.clone()),
    );
    match case.path {
        Path::Scalar => match evaluate_scalar(
            &expression,
            &EvaluationContext::with_options(&execution, date_options(case)),
            &scalar_limits(case),
        ) {
            Ok(result) => {
                if matches!(case.expectation, Expected::Failure(_)) {
                    return Err(format!("{} unexpectedly produced a value", case.name).into());
                }
                validate_scalar(case, result.value())?;
            },
            Err(error) => validate_failure(case, &error)?,
        },
        Path::Value => match value::evaluate(
            &expression,
            &resolver,
            &Context::new(&execution, Position::new("Main", 0, 0))
                .with_mode(case.shape.mode())
                .with_options(date_options(case)),
            &value_limits(case),
        ) {
            Ok(result) => {
                if matches!(case.expectation, Expected::Failure(_)) {
                    return Err(format!("{} unexpectedly produced a value", case.name).into());
                }
                validate_value(case, &result)?;
            },
            Err(error) => validate_failure(case, &error)?,
        },
    }
    Ok(resolver.stats.reads())
}

fn sample_json(sample: &Sample) -> String {
    format!(
        "{{\"elapsed_ns\":{},\"alloc_calls\":{},\"dealloc_calls\":{},\"requested_bytes\":{},\"released_bytes\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"work\":{},\"memory_retained\":{},\"reference_reads\":{},\"output_bytes\":{},\"checksum\":{}}}",
        sample.elapsed_ns,
        sample.alloc_calls,
        sample.dealloc_calls,
        sample.requested_bytes,
        sample.released_bytes,
        sample.live_before,
        sample.live_after,
        sample.peak_live_delta,
        sample.work,
        sample.memory_retained,
        sample.reference_reads,
        sample.output_bytes,
        sample.checksum
    )
}

fn default_repeat(case: &Case) -> usize {
    if matches!(case.expectation, Expected::Failure(FailureKind::Cancelled)) {
        return 4;
    }
    if case.name.contains("4096") || case.name.contains("256") || case.name.contains("1024") {
        return 1;
    }
    if case.shape.is_scalar() {
        1000
    } else if case.shape.elements() >= 256 {
        1
    } else {
        16
    }
}

fn json_escape(value: &str) -> String {
    value.replace('\\', "\\\\").replace('"', "\\\"")
}

fn evaluation_path(case: &Case) -> &'static str {
    match case.path {
        Path::Scalar => "scalar",
        Path::Value => "value",
    }
}

fn describe_json(case: &Case, reference_reads: Option<u64>) -> String {
    let reads = reference_reads.map_or_else(|| "null".to_owned(), |value| value.to_string());
    format!(
        "{{\"case\":\"{}\",\"source\":\"{}\",\"evaluation_path\":\"{}\",\"reference_reads\":{reads}}}",
        json_escape(&case.name),
        json_escape(&case.source),
        evaluation_path(case),
    )
}

fn emit(
    case: &Case,
    phase: Phase,
    repeat: usize,
    warmups: usize,
    iterations: usize,
    samples: &[Sample],
) {
    let (shape, rows, columns, elements) = match case.shape {
        Shape::Scalar => ("scalar", 0, 0, 1),
        Shape::Array { rows, columns, .. } => {
            ("array", rows, columns, rows.saturating_mul(columns))
        },
    };
    let raw = samples
        .iter()
        .map(sample_json)
        .collect::<Vec<_>>()
        .join(",");
    let p = |f: fn(&Sample) -> u64, q: usize| {
        let mut values = samples.iter().map(f).collect::<Vec<_>>();
        values.sort_unstable();
        values[(values.len().saturating_sub(1) * q) / 100]
    };
    let mean_elapsed = samples
        .iter()
        .map(|s| u128::from(s.elapsed_ns))
        .sum::<u128>()
        / samples.len() as u128;
    println!(
        "{{\"case\":\"{}\",\"operation\":\"{}\",\"phase\":\"{}\",\"input_bytes\":{},\"output_bytes_p50\":{},\"bytes_per_repeat_p50\":{},\"repeat\":{},\"warmups\":{},\"iterations\":{},\"shape\":\"{}\",\"rows\":{},\"columns\":{},\"elements\":{},\"supported\":true,\"expected\":\"{}\",\"elapsed_ns_p50\":{},\"elapsed_ns_mean\":{},\"elapsed_ns_p95\":{},\"elapsed_ns_p99\":{},\"elapsed_ns_per_repeat\":{},\"allocator_calls_p50\":{},\"requested_bytes_p50\":{},\"released_bytes_p50\":{},\"peak_live_delta_p50\":{},\"memory_retained_p50\":{},\"work_p50\":{},\"work_per_repeat\":{},\"reference_reads_p50\":{},\"reference_reads_per_repeat\":{},\"checksum_p50\":{},\"live_before_p50\":{},\"live_after_p50\":{},\"samples\":{},\"rss_kib\":null,\"rss_source\":\"external /usr/bin/time -v\",\"memory_retained_source\":\"execution budget while result is live\",\"reference_source\":\"instrumented borrowing resolver\",\"validation_scope\":\"one untimed independent oracle; timed evaluator/checksum/drop\"}}",
        json_escape(&case.name),
        case.date.map_or_else(
            || case.operation.map_or("control", ControlOp::name),
            |(op, _)| op.name()
        ),
        phase.label(),
        input_bytes(case),
        p(|s| s.output_bytes, 50),
        input_bytes(case).saturating_add(p(|s| s.output_bytes, 50) / repeat as u64),
        repeat,
        warmups,
        iterations,
        shape,
        rows,
        columns,
        elements,
        json_escape(&format!("{:?}", case.expectation)),
        p(|s| s.elapsed_ns, 50),
        mean_elapsed,
        p(|s| s.elapsed_ns, 95),
        p(|s| s.elapsed_ns, 99),
        p(|s| s.elapsed_ns, 50) / repeat as u64,
        p(|s| s.alloc_calls, 50),
        p(|s| s.requested_bytes, 50),
        p(|s| s.released_bytes, 50),
        p(|s| s.peak_live_delta, 50),
        p(|s| s.memory_retained, 50),
        p(|s| s.work, 50),
        p(|s| s.work, 50) / repeat as u64,
        p(|s| s.reference_reads, 50),
        p(|s| s.reference_reads, 50) / repeat as u64,
        p(|s| s.checksum, 50),
        p(|s| s.live_before, 50),
        p(|s| s.live_after, 50),
        raw
    );
}

fn selected_cases<'a>(all: &'a [Case], requested: Option<&str>) -> AnyResult<Vec<&'a Case>> {
    match requested {
        None | Some("all") => Ok(all.iter().collect()),
        Some(name) => all
            .iter()
            .find(|case| case.name == name)
            .map(|case| vec![case])
            .ok_or_else(|| format!("unknown case {name}").into()),
    }
}

fn run_case(case: &Case, config: &Config) -> AnyResult<()> {
    let reads = preflight(case)?;
    if config.preflight_only {
        println!("preflight-json {}", describe_json(case, Some(reads)));
        println!("preflight case={} reference_reads={reads}", case.name);
        return Ok(());
    }
    let repeat = config.repeat.unwrap_or_else(|| default_repeat(case));
    let expression = Expression::parse(&case.source)?;
    for _ in 0..config.warmups {
        match config.phase {
            Phase::Evaluate => {
                black_box(measure_evaluate(case, &expression, repeat)?);
            },
            Phase::ParseEvaluate => {
                black_box(measure_parse_evaluate(case, repeat)?);
            },
        }
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(match config.phase {
            Phase::Evaluate => measure_evaluate(case, &expression, repeat)?,
            Phase::ParseEvaluate => measure_parse_evaluate(case, repeat)?,
        });
    }
    emit(
        case,
        config.phase,
        repeat,
        config.warmups,
        config.iterations,
        &samples,
    );
    Ok(())
}

fn main() -> AnyResult<()> {
    let Some(config) = parse_config()? else {
        return Ok(());
    };
    let all = all_cases();
    let selected = selected_cases(&all, config.case.as_deref())?;
    if config.describe {
        for case in selected.iter().copied() {
            println!("{}", describe_json(case, None));
        }
        return Ok(());
    }
    for case in selected.iter().copied() {
        run_case(case, &config)?;
    }
    if config.preflight_only {
        println!(
            "preflight-ok cases={} phase={}",
            selected.len(),
            config.phase.label()
        );
    }
    Ok(())
}
