//! Bounded performance harness for the OpenFormula database-function family.
//!
//! The corpus uses a deterministic read-only resolver and covers all twelve
//! database functions, reference row scaling, unused-column locality, criteria
//! row/column preparation, lazy and projected queries, DGET cardinality, and
//! typed Work/Memory/cancellation refusals.  Database and criteria references
//! are prepared in the resolver; the evaluator parses each expression before
//! the evaluate timer.  A single untimed evaluation is checked against the
//! independent record/statistics oracle before timing begins.
//!
//! The evaluate timer includes evaluator execution, provider reads, checksum
//! folding, and result drop.  It excludes expression parsing, fixture setup,
//! and oracle validation.  `memory_retained` is the execution-budget
//! reservation observed while a result is live, not transient allocator peak;
//! allocator counters are process-wide and `rss_kib` is supplied by the
//! external `/usr/bin/time -v` wrapper used by the capture script.  Resolver
//! read counters are intentionally instrumented and therefore are diagnostic,
//! rather than a claim about an uninstrumented provider's latency.

use std::{
    alloc::{GlobalAlloc, Layout, System},
    env,
    error::Error,
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
        ArrayView, CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, SheetExtent,
        Value, evaluate,
    },
    evaluation::{EvaluationFailure, ScalarError},
    expression::Expression,
};

type AnyResult<T> = Result<T, Box<dyn Error>>;

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

#[derive(Clone, Copy, Debug)]
enum RegionKind {
    East,
    West,
    North,
    None,
    EastOrWest,
}

impl RegionKind {
    const fn text(self) -> &'static str {
        match self {
            Self::East => "East",
            Self::West => "West",
            Self::North => "North",
            Self::None => "NoSuchRegion",
            Self::EastOrWest => "East",
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum CriteriaSpec {
    Region { kind: RegionKind, body_rows: usize },
    NumericBounds { columns: usize },
}

impl CriteriaSpec {
    const fn rows(self) -> usize {
        match self {
            Self::Region { body_rows, .. } => body_rows.saturating_add(1),
            Self::NumericBounds { .. } => 2,
        }
    }

    const fn columns(self) -> usize {
        match self {
            Self::Region { .. } => 1,
            Self::NumericBounds { columns } => columns,
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct Fixture {
    data_rows: usize,
    database_columns: usize,
    criteria: CriteriaSpec,
    wide_unused_errors: bool,
}

impl Fixture {
    fn criteria_start(self) -> usize {
        // Leave one physical column between the database and criteria.  The
        // explicit dot on both ends of the generated references is deliberate.
        self.database_columns.saturating_add(1)
    }

    fn database_reference(self) -> String {
        let end = a1_column(self.database_columns.saturating_sub(1));
        format!("[.A1:.{end}{}]", self.data_rows.saturating_add(1))
    }

    fn criteria_reference(self) -> String {
        let start = a1_column(self.criteria_start());
        let end = a1_column(
            self.criteria_start()
                .saturating_add(self.criteria.columns().saturating_sub(1)),
        );
        format!("[.{start}1:.{end}{}]", self.criteria.rows())
    }

    fn extent(self) -> SheetExtent {
        let rows = self.data_rows.saturating_add(1).max(self.criteria.rows());
        let columns = self.database_columns.max(
            self.criteria_start()
                .saturating_add(self.criteria.columns()),
        );
        SheetExtent::new(rows, columns)
    }
}

#[derive(Clone, Copy, Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(&'static str),
    Error(ScalarError),
}

/// The first seven records make the small cases easy to audit.  Larger cases
/// continue with a periodic pattern so the oracle stays compact and the
/// provider does not allocate one object per physical cell.
fn record_region(row: usize) -> &'static str {
    match row {
        0 | 2 | 3 | 5 => "East",
        1 => "West",
        4 => "North",
        6 => "Eastward",
        _ => match row % 5 {
            0 => "East",
            1 => "West",
            2 | 3 => "East",
            _ => "Eastward",
        },
    }
}

fn record_amount(row: usize) -> FixtureCell {
    match row {
        0 => FixtureCell::Number(10.0),
        1 => FixtureCell::Number(20.0),
        2 => FixtureCell::Number(30.0),
        3 => FixtureCell::Text("n/a"),
        4 => FixtureCell::Number(40.0),
        5 => FixtureCell::Empty,
        6 => FixtureCell::Number(60.0),
        _ => match row % 5 {
            0 => FixtureCell::Number(10.0 + (row % 91) as f64),
            1 => FixtureCell::Number(20.0 + (row % 91) as f64),
            2 => FixtureCell::Number(30.0 + (row % 91) as f64),
            3 => FixtureCell::Text("n/a"),
            _ => FixtureCell::Number(60.0 + (row % 91) as f64),
        },
    }
}

fn database_cell(fixture: Fixture, row: usize, column: usize) -> FixtureCell {
    if row == 0 {
        return match column {
            0 => FixtureCell::Text("Name"),
            1 => FixtureCell::Text("Region"),
            2 => FixtureCell::Text("Amount"),
            3 => FixtureCell::Text("Score"),
            4 => FixtureCell::Text("Active"),
            _ => FixtureCell::Text("Unused"),
        };
    }

    let record = row - 1;
    match column {
        0 => FixtureCell::Text("record"),
        1 => FixtureCell::Text(record_region(record)),
        2 => record_amount(record),
        3 => FixtureCell::Number((record.saturating_add(1)) as f64),
        4 => FixtureCell::Logical(record % 2 == 0),
        _ if fixture.wide_unused_errors => FixtureCell::Error(ScalarError::NotAvailable),
        _ => FixtureCell::Text("unused"),
    }
}

fn criteria_cell(criteria: CriteriaSpec, row: usize, column: usize) -> FixtureCell {
    if row == 0 {
        return match criteria {
            CriteriaSpec::Region { .. } => FixtureCell::Text("rEgIoN"),
            CriteriaSpec::NumericBounds { .. } => FixtureCell::Number(3.0),
        };
    }
    match criteria {
        CriteriaSpec::Region { kind, .. } => {
            let text = match kind {
                RegionKind::EastOrWest if row % 2 == 0 => "West",
                RegionKind::EastOrWest => "East",
                _ => kind.text(),
            };
            FixtureCell::Text(text)
        },
        CriteriaSpec::NumericBounds { .. } => FixtureCell::Text(match column {
            0 => ">=10",
            1 => "<=100",
            2 => "<>30",
            _ => ">0",
        }),
    }
}

#[derive(Debug)]
struct ResolverStats {
    reads: AtomicU64,
    database_reads: AtomicU64,
    criteria_reads: AtomicU64,
    extent_calls: AtomicU64,
}

impl ResolverStats {
    fn new() -> Self {
        Self {
            reads: AtomicU64::new(0),
            database_reads: AtomicU64::new(0),
            criteria_reads: AtomicU64::new(0),
            extent_calls: AtomicU64::new(0),
        }
    }

    fn snapshot(&self) -> ResolverSnapshot {
        ResolverSnapshot {
            reads: self.reads.load(Ordering::Acquire),
            database_reads: self.database_reads.load(Ordering::Acquire),
            criteria_reads: self.criteria_reads.load(Ordering::Acquire),
            extent_calls: self.extent_calls.load(Ordering::Acquire),
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct ResolverSnapshot {
    reads: u64,
    database_reads: u64,
    criteria_reads: u64,
    extent_calls: u64,
}

/// Immutable, borrow-returning provider.  Large fixtures are generated from
/// row/column coordinates so setup does not itself become a database-table
/// allocation benchmark.  Unused wide columns deliberately return Error: a
/// successful query demonstrates that required-column admission avoids them.
#[derive(Debug)]
struct DatabaseResolver {
    fixture: Fixture,
    stats: ResolverStats,
}

impl DatabaseResolver {
    fn new(fixture: Fixture) -> Self {
        Self {
            fixture,
            stats: ResolverStats::new(),
        }
    }

    fn in_database(&self, row: usize, column: usize) -> bool {
        row <= self.fixture.data_rows && column < self.fixture.database_columns
    }

    fn in_criteria(&self, row: usize, column: usize) -> bool {
        row < self.fixture.criteria.rows()
            && column >= self.fixture.criteria_start()
            && column
                < self
                    .fixture
                    .criteria_start()
                    .saturating_add(self.fixture.criteria.columns())
    }
}

impl Resolver for DatabaseResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        execution.check()?;
        self.stats.extent_calls.fetch_add(1, Ordering::Relaxed);
        Ok((sheet == "Main").then_some(self.fixture.extent()))
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
        let value = if sheet != "Main" {
            FixtureCell::Error(ScalarError::Reference)
        } else if self.in_database(row, column) {
            self.stats.database_reads.fetch_add(1, Ordering::Relaxed);
            database_cell(self.fixture, row, column)
        } else if self.in_criteria(row, column) {
            self.stats.criteria_reads.fetch_add(1, Ordering::Relaxed);
            criteria_cell(
                self.fixture.criteria,
                row,
                column.saturating_sub(self.fixture.criteria_start()),
            )
        } else {
            FixtureCell::Empty
        };
        Ok(match value {
            FixtureCell::Empty => CellRead::Empty,
            FixtureCell::Number(value) => CellRead::Number(value),
            FixtureCell::Logical(value) => CellRead::Logical(value),
            FixtureCell::Text(value) => CellRead::Text(value),
            FixtureCell::Error(error) => CellRead::Error(error),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        execution.check()?;
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        execution.check()?;
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        execution.check()?;
        Ok(1)
    }

    fn source_version(
        &self,
        execution: &ExecutionContext,
    ) -> Result<Option<litchi_core::SourceVersion>, EvaluationFailure> {
        execution.check()?;
        Ok(None)
    }
}

#[derive(Clone, Copy, Debug)]
enum LimitProfile {
    Default,
    Work,
    Memory,
    Cancelled,
}

#[derive(Clone, Debug)]
enum Expected {
    Number(f64),
    Array {
        rows: usize,
        columns: usize,
        cells: Vec<f64>,
    },
    FormulaError,
    EvaluationFailure(&'static str),
}

#[derive(Debug)]
struct CaseSpec {
    name: String,
    function: &'static str,
    source: String,
    size: usize,
    fixture: Fixture,
    expected: Expected,
    limit: LimitProfile,
    zero_reads: bool,
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
    successes: u64,
    refusals: u64,
    checksum: u64,
    resolver: ResolverSnapshot,
}

#[derive(Clone, Copy, Debug)]
enum Phase {
    Evaluate,
    ParseEvaluate,
    Parse,
}

impl Phase {
    fn parse(value: &str) -> AnyResult<Self> {
        match value {
            "evaluate" => Ok(Self::Evaluate),
            "parse-evaluate" => Ok(Self::ParseEvaluate),
            "parse" => Ok(Self::Parse),
            other => Err(format!(
                "unknown phase {other:?}; expected evaluate, parse-evaluate, or parse"
            )
            .into()),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Evaluate => "evaluate",
            Self::ParseEvaluate => "parse-evaluate",
            Self::Parse => "parse",
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
}

fn parse_config() -> AnyResult<Option<Config>> {
    let mut case = None;
    let mut phase = Phase::Evaluate;
    let mut warmups = 3;
    let mut iterations = 15;
    let mut repeat = None;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                println!(
                    "usage: database-profile [--case NAME|all] [--phase evaluate|parse-evaluate|parse] [--warmups N] [--iterations N] [--repeat N]\n\
                     --list prints deterministic case names; references use the Main resolver and a bounded generated database.\n\
                     JSON lines are one validated aggregate row per case. Wrap the binary in /usr/bin/time -v for RSS."
                );
                return Ok(None);
            },
            "--list" => {
                for spec in corpus() {
                    println!("{}", spec.name);
                }
                return Ok(None);
            },
            "--case" => {
                case = Some(arguments.next().ok_or("--case requires a name or all")?);
            },
            "--phase" => {
                phase = Phase::parse(&arguments.next().ok_or("--phase requires a value")?)?;
            },
            "--warmups" => {
                warmups = arguments
                    .next()
                    .ok_or("--warmups requires a nonnegative integer")?
                    .parse()
                    .map_err(|_| "--warmups requires a nonnegative integer")?;
            },
            "--iterations" => {
                iterations = arguments
                    .next()
                    .ok_or("--iterations requires a positive integer")?
                    .parse()
                    .map_err(|_| "--iterations requires a positive integer")?;
            },
            "--repeat" => {
                repeat = Some(
                    arguments
                        .next()
                        .ok_or("--repeat requires a positive integer")?
                        .parse()
                        .map_err(|_| "--repeat requires a positive integer")?,
                );
            },
            other => return Err(format!("unknown option {other:?}; use --help").into()),
        }
    }
    if iterations == 0 {
        return Err("--iterations must be positive".into());
    }
    if repeat == Some(0) {
        return Err("--repeat must be positive".into());
    }
    Ok(Some(Config {
        case,
        phase,
        warmups,
        iterations,
        repeat,
    }))
}

fn default_repeat(size: usize) -> usize {
    match size {
        0..=8 => 8,
        16 => 4,
        256 => 2,
        _ => 1,
    }
}

fn a1_column(mut column: usize) -> String {
    let mut output = String::new();
    loop {
        let remainder = column % 26;
        output.insert(0, (b'A' + remainder as u8) as char);
        if column < 26 {
            break;
        }
        column = column / 26 - 1;
    }
    output
}

fn record_matches(criteria: CriteriaSpec, row: usize) -> bool {
    let region = record_region(row);
    match criteria {
        CriteriaSpec::Region { kind, .. } => match kind {
            RegionKind::East => region == "East",
            RegionKind::West => region == "West",
            RegionKind::North => region == "North",
            RegionKind::None => false,
            RegionKind::EastOrWest => region == "East" || region == "West",
        },
        CriteriaSpec::NumericBounds { .. } => match record_amount(row) {
            FixtureCell::Number(value) => (10.0..=100.0).contains(&value) && value != 30.0,
            FixtureCell::Empty
            | FixtureCell::Logical(_)
            | FixtureCell::Text(_)
            | FixtureCell::Error(_) => false,
        },
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct NumericSummary {
    matched: usize,
    numeric: usize,
    nonblank: usize,
    sum: f64,
    product: f64,
    minimum: Option<f64>,
    maximum: Option<f64>,
    mean: f64,
    m2: f64,
}

fn summary(fixture: Fixture) -> NumericSummary {
    let mut result = NumericSummary {
        product: 1.0,
        ..NumericSummary::default()
    };
    for row in 0..fixture.data_rows {
        if !record_matches(fixture.criteria, row) {
            continue;
        }
        result.matched = result.matched.saturating_add(1);
        let value = record_amount(row);
        if !matches!(value, FixtureCell::Empty) {
            result.nonblank = result.nonblank.saturating_add(1);
        }
        let FixtureCell::Number(value) = value else {
            continue;
        };
        result.numeric = result.numeric.saturating_add(1);
        result.sum += value;
        result.product *= value;
        result.minimum = Some(result.minimum.map_or(value, |old| old.min(value)));
        result.maximum = Some(result.maximum.map_or(value, |old| old.max(value)));
        let count = result.numeric as f64;
        let delta = value - result.mean;
        result.mean += delta / count;
        result.m2 += delta * (value - result.mean);
    }
    result
}

fn expected_database(function: &'static str, fixture: Fixture, omitted_field: bool) -> Expected {
    let stats = summary(fixture);
    match function {
        "DCOUNT" => Expected::Number(if omitted_field {
            stats.matched as f64
        } else {
            stats.numeric as f64
        }),
        "DCOUNTA" => Expected::Number(if omitted_field {
            stats.matched as f64
        } else {
            stats.nonblank as f64
        }),
        "DGET" => {
            if stats.matched != 1 {
                Expected::FormulaError
            } else {
                match (0..fixture.data_rows)
                    .filter(|row| record_matches(fixture.criteria, *row))
                    .find_map(|row| match record_amount(row) {
                        FixtureCell::Number(value) => Some(value),
                        FixtureCell::Empty => Some(0.0),
                        FixtureCell::Logical(value) => Some(if value { 1.0 } else { 0.0 }),
                        FixtureCell::Text(_) | FixtureCell::Error(_) => None,
                    }) {
                    Some(value) => Expected::Number(value),
                    None => Expected::FormulaError,
                }
            }
        },
        "DAVERAGE" => {
            if stats.numeric == 0 {
                Expected::FormulaError
            } else {
                Expected::Number(stats.sum / stats.numeric as f64)
            }
        },
        "DMAX" => Expected::Number(stats.maximum.unwrap_or(0.0)),
        "DMIN" => Expected::Number(stats.minimum.unwrap_or(0.0)),
        "DPRODUCT" => Expected::Number(if stats.numeric == 0 {
            1.0
        } else {
            stats.product
        }),
        "DSUM" => Expected::Number(stats.sum),
        "DSTDEV" | "DVAR" => {
            if stats.numeric < 2 {
                Expected::FormulaError
            } else {
                let value = stats.m2 / (stats.numeric as f64 - 1.0);
                if function == "DSTDEV" {
                    Expected::Number(value.sqrt())
                } else {
                    Expected::Number(value)
                }
            }
        },
        "DSTDEVP" | "DVARP" => {
            if stats.numeric == 0 {
                Expected::FormulaError
            } else {
                let value = stats.m2 / stats.numeric as f64;
                if function == "DSTDEVP" {
                    Expected::Number(value.sqrt())
                } else {
                    Expected::Number(value)
                }
            }
        },
        _ => panic!("unknown database function {function}"),
    }
}

fn query_source(fixture: Fixture, function: &str, field: Option<&str>) -> String {
    let database = fixture.database_reference();
    let criteria = fixture.criteria_reference();
    match field {
        Some(field) => format!("={function}({database};\"{field}\";{criteria})"),
        None => format!("={function}({database};;{criteria})"),
    }
}

fn query_case(
    name: &str,
    function: &'static str,
    fixture: Fixture,
    omitted_field: bool,
) -> CaseSpec {
    let source = query_source(fixture, function, (!omitted_field).then_some("Amount"));
    CaseSpec {
        name: name.to_owned(),
        function,
        source,
        size: fixture.data_rows,
        fixture,
        expected: expected_database(function, fixture, omitted_field),
        limit: LimitProfile::Default,
        zero_reads: false,
    }
}

fn failure_case(
    name: &str,
    fixture: Fixture,
    limit: LimitProfile,
    label: &'static str,
) -> CaseSpec {
    CaseSpec {
        name: name.to_owned(),
        function: "DSUM",
        source: query_source(fixture, "DSUM", Some("Amount")),
        size: fixture.data_rows,
        fixture,
        expected: Expected::EvaluationFailure(label),
        limit,
        zero_reads: false,
    }
}

fn corpus() -> Vec<CaseSpec> {
    let small_east = Fixture {
        data_rows: 7,
        database_columns: 6,
        criteria: CriteriaSpec::Region {
            kind: RegionKind::East,
            body_rows: 1,
        },
        wide_unused_errors: false,
    };
    let small_north = Fixture {
        criteria: CriteriaSpec::Region {
            kind: RegionKind::North,
            body_rows: 1,
        },
        ..small_east
    };
    let mut cases = Vec::new();

    for function in [
        "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP",
        "DSUM", "DVAR", "DVARP",
    ] {
        let fixture = if function == "DGET" {
            small_north
        } else {
            small_east
        };
        cases.push(query_case(
            &format!("database-{function}").to_ascii_lowercase(),
            function,
            fixture,
            false,
        ));
    }

    cases.push(query_case(
        "database-dcount-omitted-field",
        "DCOUNT",
        small_east,
        true,
    ));

    for rows in [16, 256, 4096] {
        let fixture = Fixture {
            data_rows: rows,
            ..small_east
        };
        cases.push(query_case(
            &format!("database-dsum-reference-{rows}"),
            "DSUM",
            fixture,
            false,
        ));
    }

    cases.push(query_case(
        "database-dsum-wide-unused-4096",
        "DSUM",
        Fixture {
            data_rows: 4096,
            database_columns: 32,
            wide_unused_errors: true,
            ..small_east
        },
        false,
    ));

    cases.push(query_case(
        "database-dsum-criteria-rows-16",
        "DSUM",
        Fixture {
            data_rows: 256,
            criteria: CriteriaSpec::Region {
                kind: RegionKind::EastOrWest,
                body_rows: 16,
            },
            ..small_east
        },
        false,
    ));
    cases.push(query_case(
        "database-dsum-criteria-columns-4-numeric-headers",
        "DSUM",
        Fixture {
            data_rows: 256,
            criteria: CriteriaSpec::NumericBounds { columns: 4 },
            ..small_east
        },
        false,
    ));

    let large = Fixture {
        data_rows: 4096,
        ..small_east
    };
    let lazy_source = format!(
        "=IF(FALSE();DSUM({};\"Amount\";{});0)",
        large.database_reference(),
        large.criteria_reference()
    );
    cases.push(CaseSpec {
        name: "database-lazy-unselected-4096".to_owned(),
        function: "DSUM",
        source: lazy_source,
        size: large.data_rows,
        fixture: large,
        expected: Expected::Number(0.0),
        limit: LimitProfile::Default,
        zero_reads: true,
    });

    let cached_sum = match expected_database(
        "DSUM",
        Fixture {
            data_rows: 256,
            ..small_east
        },
        false,
    ) {
        Expected::Number(value) => value,
        _ => unreachable!("DSUM cache oracle is numeric"),
    };
    let cached_fixture = Fixture {
        data_rows: 256,
        ..small_east
    };
    cases.push(CaseSpec {
        name: "database-projected-cache-256".to_owned(),
        function: "DSUM",
        source: format!(
            "=IF({{TRUE();FALSE();TRUE()}};DSUM({};\"Amount\";{});0)",
            cached_fixture.database_reference(),
            cached_fixture.criteria_reference()
        ),
        size: cached_fixture.data_rows,
        fixture: cached_fixture,
        expected: Expected::Array {
            rows: 1,
            columns: 3,
            cells: vec![cached_sum, 0.0, cached_sum],
        },
        limit: LimitProfile::Default,
        zero_reads: false,
    });

    cases.push(query_case(
        "database-dget-multiple-cardinality",
        "DGET",
        small_east,
        false,
    ));

    for (function, name) in [
        ("DSUM", "database-empty-dsum"),
        ("DMAX", "database-empty-dmax"),
        ("DMIN", "database-empty-dmin"),
        ("DPRODUCT", "database-empty-dproduct"),
    ] {
        let fixture = Fixture {
            criteria: CriteriaSpec::Region {
                kind: RegionKind::None,
                body_rows: 1,
            },
            ..small_east
        };
        cases.push(query_case(name, function, fixture, false));
    }

    let refusal_fixture = Fixture {
        data_rows: 256,
        ..small_east
    };
    cases.push(failure_case(
        "database-refusal-work",
        refusal_fixture,
        LimitProfile::Work,
        "work",
    ));
    cases.push(failure_case(
        "database-refusal-memory",
        refusal_fixture,
        LimitProfile::Memory,
        "memory",
    ));
    cases.push(failure_case(
        "database-refusal-cancelled",
        refusal_fixture,
        LimitProfile::Cancelled,
        "cancelled",
    ));
    cases
}

fn execution(cancelled: bool) -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "ods-formula-database-functions-performance",
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
    let context = ExecutionContext::new(budget, token, limits);
    if cancelled {
        cancellation.cancel();
    }
    (cancellation, context)
}

fn limits(case: &CaseSpec) -> Limits {
    match case.limit {
        LimitProfile::Work => Limits::default().with_max_steps(0),
        LimitProfile::Memory => Limits::default().with_max_storage_bytes(0),
        LimitProfile::Default | LimitProfile::Cancelled => Limits::default(),
    }
}

fn is_cancelled(case: &CaseSpec) -> bool {
    matches!(case.limit, LimitProfile::Cancelled)
}

fn approximately_equal(actual: f64, expected: f64) -> bool {
    let tolerance = 1e-9 * actual.abs().max(expected.abs()).max(1.0);
    (actual - expected).abs() <= tolerance
}

fn validate_number(value: Value<'_>, expected: f64, case: &str) -> AnyResult<()> {
    match value {
        Value::Number(actual) if approximately_equal(actual, expected) => Ok(()),
        other => Err(format!("{case} returned {other:?}, expected Number({expected})").into()),
    }
}

fn validate_array(
    value: ArrayView<'_>,
    rows: usize,
    columns: usize,
    cells: &[f64],
    case: &str,
) -> AnyResult<()> {
    if value.shape().rows() != rows || value.shape().columns() != columns {
        return Err(format!(
            "{case} returned shape {}x{}, expected {rows}x{columns}",
            value.shape().rows(),
            value.shape().columns()
        )
        .into());
    }
    if cells.len() != rows.saturating_mul(columns) {
        return Err(format!("{case} oracle shape does not match its cells").into());
    }
    for (index, expected) in cells.iter().copied().enumerate() {
        match value.get(index) {
            Some(Value::Number(actual)) if approximately_equal(actual, expected) => {},
            other => {
                return Err(format!(
                    "{case} cell {index} returned {other:?}, expected Number({expected})"
                )
                .into());
            },
        }
    }
    Ok(())
}

fn validate_value(case: &CaseSpec, value: Value<'_>) -> AnyResult<()> {
    match &case.expected {
        Expected::Number(expected) => validate_number(value, *expected, &case.name),
        Expected::Array {
            rows,
            columns,
            cells,
        } => match value {
            Value::Array(array) => validate_array(array, *rows, *columns, cells, &case.name),
            other => Err(format!("{} returned {other:?}, expected array", case.name).into()),
        },
        Expected::FormulaError => match value {
            Value::Error(_) => Ok(()),
            other => {
                Err(format!("{} returned {other:?}, expected formula error", case.name).into())
            },
        },
        Expected::EvaluationFailure(label) => Err(format!(
            "{} has evaluator failure expectation {label}, not a Value",
            case.name
        )
        .into()),
    }
}

fn failure_matches(label: &str, error: &EvaluationFailure) -> bool {
    match (label, error) {
        ("cancelled", EvaluationFailure::Cancelled) => true,
        ("work", EvaluationFailure::ResourceLimit(limit)) => limit.resource == Resource::Work,
        ("memory", EvaluationFailure::ResourceLimit(limit)) => limit.resource == Resource::Memory,
        _ => false,
    }
}

fn preflight_evaluate(case: &CaseSpec, expression: &Expression) -> AnyResult<()> {
    let resolver = DatabaseResolver::new(case.fixture);
    let (_cancellation, execution) = execution(is_cancelled(case));
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = evaluate(expression, &resolver, &context, &limits(case));
    match &case.expected {
        Expected::EvaluationFailure(label) => match result {
            Err(error) if failure_matches(label, &error) => Ok(()),
            Err(error) => {
                Err(format!("{} returned wrong {label} refusal: {error}", case.name).into())
            },
            Ok(value) => Err(format!(
                "{} unexpectedly returned {:?} for {label} refusal",
                case.name,
                value.value()
            )
            .into()),
        },
        expected => {
            let value =
                result.map_err(|error| format!("{} evaluator failure: {error}", case.name))?;
            validate_value(case, value.value())?;
            if case.zero_reads && resolver.stats.snapshot().reads != 0 {
                return Err(format!(
                    "{} lazy oracle read {} provider cells",
                    case.name,
                    resolver.stats.snapshot().reads
                )
                .into());
            }
            if matches!(expected, Expected::Array { .. }) && resolver.stats.snapshot().reads == 0 {
                return Err(format!("{} projected query read no provider cells", case.name).into());
            }
            Ok(())
        },
    }
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
        Value::Number(number) => number.to_bits().rotate_left(11) ^ 0x4e,
        Value::Logical(logical) => u64::from(logical) ^ 0x4c,
        Value::Text(text) => checksum_bytes(text.as_bytes()) ^ 0x54,
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

fn consume_result(
    case: &CaseSpec,
    result: Result<Evaluated<'_>, EvaluationFailure>,
    execution: &ExecutionContext,
    successes: &mut u64,
    refusals: &mut u64,
    checksum: &mut u64,
    retained_peak: &mut u64,
) -> AnyResult<()> {
    match (&case.expected, result) {
        (Expected::EvaluationFailure(label), Err(error)) if failure_matches(label, &error) => {
            *refusals = refusals.saturating_add(1);
            black_box(error);
            Ok(())
        },
        (Expected::EvaluationFailure(label), Err(error)) => {
            Err(format!("{} returned wrong {label} refusal: {error}", case.name).into())
        },
        (Expected::EvaluationFailure(label), Ok(value)) => Err(format!(
            "{} unexpectedly returned {:?} for {label} refusal",
            case.name,
            value.value()
        )
        .into()),
        (_, Err(error)) => Err(format!("{} evaluator failure: {error}", case.name).into()),
        (_, Ok(value)) => {
            *successes = successes.saturating_add(1);
            *checksum = checksum.wrapping_add(checksum_value(value.value()));
            *retained_peak = (*retained_peak).max(execution.budget().used(Resource::Memory));
            black_box(*checksum);
            drop(value);
            Ok(())
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

fn sample(
    elapsed_ns: u64,
    live_before: u64,
    baseline_work: u64,
    baseline_memory: u64,
    execution: &ExecutionContext,
    resolver: ResolverSnapshot,
    successes: u64,
    refusals: u64,
    checksum: u64,
    retained_peak: u64,
) -> Sample {
    Sample {
        elapsed_ns,
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
        successes,
        refusals,
        checksum,
        resolver,
    }
}

fn measure_evaluate(case: &CaseSpec, expression: &Expression, repeat: usize) -> AnyResult<Sample> {
    let resolver = DatabaseResolver::new(case.fixture);
    let (_cancellation, execution) = execution(is_cancelled(case));
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let limits = limits(case);
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0;
    let mut refusals = 0;
    let mut checksum = 0;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        consume_result(
            case,
            evaluate(expression, &resolver, &context, &limits),
            &execution,
            &mut successes,
            &mut refusals,
            &mut checksum,
            &mut retained_peak,
        )?;
    }
    let elapsed_ns = started.elapsed().as_nanos().min(u64::MAX as u128) as u64;
    Ok(sample(
        elapsed_ns,
        live_before,
        baseline_work,
        baseline_memory,
        &execution,
        resolver.stats.snapshot(),
        successes,
        refusals,
        checksum,
        retained_peak,
    ))
}

fn measure_parse(case: &CaseSpec, repeat: usize) -> AnyResult<Sample> {
    let live_before = reset_observer();
    let started = Instant::now();
    for _ in 0..repeat {
        let expression = Expression::parse(black_box(&case.source))
            .map_err(|error| format!("{} parse failure: {error}", case.name))?;
        black_box(expression);
    }
    Ok(Sample {
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
        work: 0,
        memory_retained: 0,
        successes: repeat as u64,
        refusals: 0,
        checksum: 0,
        resolver: ResolverSnapshot::default(),
    })
}

fn measure_parse_evaluate(case: &CaseSpec, repeat: usize) -> AnyResult<Sample> {
    let resolver = DatabaseResolver::new(case.fixture);
    let (_cancellation, execution) = execution(is_cancelled(case));
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let limits = limits(case);
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0;
    let mut refusals = 0;
    let mut checksum = 0;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        let expression = Expression::parse(black_box(&case.source))
            .map_err(|error| format!("{} parse failure: {error}", case.name))?;
        consume_result(
            case,
            evaluate(&expression, &resolver, &context, &limits),
            &execution,
            &mut successes,
            &mut refusals,
            &mut checksum,
            &mut retained_peak,
        )?;
    }
    let elapsed_ns = started.elapsed().as_nanos().min(u64::MAX as u128) as u64;
    Ok(sample(
        elapsed_ns,
        live_before,
        baseline_work,
        baseline_memory,
        &execution,
        resolver.stats.snapshot(),
        successes,
        refusals,
        checksum,
        retained_peak,
    ))
}

fn percentile(samples: &[Sample], metric: impl Fn(&Sample) -> u64, percentile: usize) -> u64 {
    let mut values: Vec<u64> = samples.iter().map(metric).collect();
    values.sort_unstable();
    let index = (values.len().saturating_sub(1) * percentile) / 100;
    values[index]
}

fn mean(samples: &[Sample], metric: impl Fn(&Sample) -> u64) -> u64 {
    let total: u128 = samples
        .iter()
        .map(|sample| u128::from(metric(sample)))
        .sum();
    (total / samples.len() as u128).min(u64::MAX as u128) as u64
}

fn expected_label(expected: &Expected) -> &'static str {
    match expected {
        Expected::Number(_) => "number",
        Expected::Array { .. } => "array",
        Expected::FormulaError => "formula-error",
        Expected::EvaluationFailure(_) => "evaluation-failure",
    }
}

fn expected_failure(expected: &Expected) -> &'static str {
    match expected {
        Expected::EvaluationFailure(label) => label,
        Expected::Number(_) | Expected::Array { .. } | Expected::FormulaError => "none",
    }
}

fn json_escape(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len() + 2);
    for character in value.chars() {
        match character {
            '"' => escaped.push_str("\\\""),
            '\\' => escaped.push_str("\\\\"),
            '\n' => escaped.push_str("\\n"),
            '\r' => escaped.push_str("\\r"),
            '\t' => escaped.push_str("\\t"),
            character if character.is_control() => {
                escaped.push_str(&format!("\\u{:04x}", character as u32));
            },
            character => escaped.push(character),
        }
    }
    escaped
}

fn emit(
    case: &CaseSpec,
    phase: Phase,
    warmups: usize,
    iterations: usize,
    repeat: usize,
    samples: &[Sample],
) {
    let elapsed_ns = percentile(samples, |sample| sample.elapsed_ns, 50);
    let work = percentile(samples, |sample| sample.work, 50);
    let expected = expected_label(&case.expected);
    let failure = expected_failure(&case.expected);
    let resolver_reads = percentile(samples, |sample| sample.resolver.reads, 50);
    let database_reads = percentile(samples, |sample| sample.resolver.database_reads, 50);
    let criteria_reads = percentile(samples, |sample| sample.resolver.criteria_reads, 50);
    let extent_calls = percentile(samples, |sample| sample.resolver.extent_calls, 50);
    println!(
        "{{\"case\":\"{}\",\"function\":\"{}\",\"size\":{},\"phase\":\"{}\",\"input_bytes\":{},\"repeat\":{},\"warmups\":{},\"iterations\":{},\"elapsed_ns\":{},\"elapsed_ns_mean\":{},\"elapsed_ns_p95\":{},\"elapsed_ns_p99\":{},\"elapsed_ns_per_repeat\":{},\"checksum\":{},\"shape\":\"{}\",\"rows\":{},\"columns\":{},\"elements\":{},\"work\":{},\"work_mean\":{},\"work_per_repeat\":{},\"allocator_calls\":{},\"allocator_calls_max\":{},\"deallocator_calls\":{},\"deallocator_calls_max\":{},\"requested_bytes\":{},\"requested_bytes_max\":{},\"released_bytes\":{},\"released_bytes_max\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"peak_live_delta_max\":{},\"memory_retained\":{},\"memory_retained_max\":{},\"successes\":{},\"refusals\":{},\"failure\":\"{}\",\"expected\":\"{}\",\"provider_reads\":{},\"provider_database_reads\":{},\"provider_criteria_reads\":{},\"provider_extent_calls\":{},\"rss_kib\":null,\"rss_source\":\"external /usr/bin/time -v\",\"memory_retained_source\":\"execution budget while result is live\",\"provider_counter_source\":\"instrumented in-memory resolver\",\"validation_scope\":\"one untimed oracle; timed evaluate/checksum/drop\"}}",
        json_escape(&case.name),
        case.function,
        case.size,
        phase.label(),
        case.source.len(),
        repeat,
        warmups,
        iterations,
        elapsed_ns,
        mean(samples, |sample| sample.elapsed_ns),
        percentile(samples, |sample| sample.elapsed_ns, 95),
        percentile(samples, |sample| sample.elapsed_ns, 99),
        elapsed_ns / repeat as u64,
        percentile(samples, |sample| sample.checksum, 50),
        if matches!(&case.expected, Expected::Array { .. }) {
            "array"
        } else if matches!(
            &case.expected,
            Expected::FormulaError | Expected::EvaluationFailure(_)
        ) {
            "error"
        } else {
            "scalar"
        },
        match &case.expected {
            Expected::Array { rows, .. } => *rows,
            _ => 0,
        },
        match &case.expected {
            Expected::Array { columns, .. } => *columns,
            _ => 0,
        },
        match &case.expected {
            Expected::Array { cells, .. } => cells.len(),
            _ => 1,
        },
        work,
        mean(samples, |sample| sample.work),
        work / repeat as u64,
        percentile(samples, |sample| sample.alloc_calls, 50),
        samples
            .iter()
            .map(|sample| sample.alloc_calls)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.dealloc_calls, 50),
        samples
            .iter()
            .map(|sample| sample.dealloc_calls)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.requested_bytes, 50),
        samples
            .iter()
            .map(|sample| sample.requested_bytes)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.released_bytes, 50),
        samples
            .iter()
            .map(|sample| sample.released_bytes)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.live_before, 50),
        percentile(samples, |sample| sample.live_after, 50),
        percentile(samples, |sample| sample.peak_live_delta, 50),
        samples
            .iter()
            .map(|sample| sample.peak_live_delta)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.memory_retained, 50),
        samples
            .iter()
            .map(|sample| sample.memory_retained)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.successes, 50),
        percentile(samples, |sample| sample.refusals, 50),
        failure,
        expected,
        resolver_reads,
        database_reads,
        criteria_reads,
        extent_calls,
    );
}

fn selected_cases<'a>(
    all: &'a [CaseSpec],
    requested: Option<&str>,
) -> AnyResult<Vec<&'a CaseSpec>> {
    match requested {
        None | Some("all") => Ok(all.iter().collect()),
        Some(name) => all
            .iter()
            .find(|case| case.name == name)
            .map(|case| vec![case])
            .ok_or_else(|| format!("unknown case {name:?}; use --list").into()),
    }
}

fn run_case(case: &CaseSpec, config: &Config) -> AnyResult<()> {
    let repeat = config.repeat.unwrap_or_else(|| default_repeat(case.size));
    let parsed = match config.phase {
        Phase::Evaluate => Some(
            Expression::parse(&case.source)
                .map_err(|error| format!("{} parse preflight failed: {error}", case.name))?,
        ),
        Phase::ParseEvaluate | Phase::Parse => None,
    };

    if let Some(expression) = parsed.as_ref() {
        preflight_evaluate(case, expression)?;
    } else if matches!(config.phase, Phase::ParseEvaluate) {
        let expression = Expression::parse(&case.source)
            .map_err(|error| format!("{} parse preflight failed: {error}", case.name))?;
        preflight_evaluate(case, &expression)?;
    } else {
        Expression::parse(&case.source)
            .map_err(|error| format!("{} parse preflight failed: {error}", case.name))?;
    }

    for _ in 0..config.warmups {
        let measured = match config.phase {
            Phase::Evaluate => {
                measure_evaluate(case, parsed.as_ref().expect("parsed evaluate"), repeat)?
            },
            Phase::ParseEvaluate => measure_parse_evaluate(case, repeat)?,
            Phase::Parse => measure_parse(case, repeat)?,
        };
        black_box(measured);
    }

    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(match config.phase {
            Phase::Evaluate => {
                measure_evaluate(case, parsed.as_ref().expect("parsed evaluate"), repeat)?
            },
            Phase::ParseEvaluate => measure_parse_evaluate(case, repeat)?,
            Phase::Parse => measure_parse(case, repeat)?,
        });
    }
    emit(
        case,
        config.phase,
        config.warmups,
        config.iterations,
        repeat,
        &samples,
    );
    Ok(())
}

fn main() -> AnyResult<()> {
    let Some(config) = parse_config()? else {
        return Ok(());
    };
    let all = corpus();
    for case in selected_cases(&all, config.case.as_deref())? {
        run_case(case, &config)?;
    }
    Ok(())
}
