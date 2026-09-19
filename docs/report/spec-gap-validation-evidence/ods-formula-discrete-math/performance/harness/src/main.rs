//! Bounded process-level profile for the ODS discrete-math evaluator.
//!
//! The fixture is deterministic and entirely in memory.  Every timed child
//! performs a correctness preflight before measuring one immutable expression
//! repeatedly. Reference cases use a borrowing resolver and count cell reads
//! so a profile row can distinguish streamed reducer paths from literal arrays.

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
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarValue, evaluate_scalar,
    },
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
        // SAFETY: the platform allocator receives the unchanged layout.
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(layout.size() as u64, Ordering::Relaxed);
        // SAFETY: pointer and layout are the pair previously returned by alloc.
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
        // SAFETY: pointer, old layout, and new size are forwarded unchanged.
        unsafe { System.realloc(pointer, layout, size) }
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum AggregateOp {
    Sum,
}

impl AggregateOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Sum => "SUM",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DiscreteOp {
    Combin,
    Combina,
    Fact,
    FactDouble,
    Gcd,
    Lcm,
    Multinomial,
    Even,
    Odd,
    Delta,
    GStep,
}

impl DiscreteOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Combin => "COMBIN",
            Self::Combina => "COMBINA",
            Self::Fact => "FACT",
            Self::FactDouble => "FACTDOUBLE",
            Self::Gcd => "GCD",
            Self::Lcm => "LCM",
            Self::Multinomial => "MULTINOMIAL",
            Self::Even => "EVEN",
            Self::Odd => "ODD",
            Self::Delta => "DELTA",
            Self::GStep => "GESTEP",
        }
    }

    const fn all() -> &'static [Self] {
        &[
            Self::Combin,
            Self::Combina,
            Self::Fact,
            Self::FactDouble,
            Self::Gcd,
            Self::Lcm,
            Self::Multinomial,
            Self::Even,
            Self::Odd,
            Self::Delta,
            Self::GStep,
        ]
    }

    const fn binary(self) -> bool {
        matches!(
            self,
            Self::Combin | Self::Combina | Self::Delta | Self::GStep
        )
    }

    const fn reducer(self) -> bool {
        matches!(self, Self::Gcd | Self::Lcm | Self::Multinomial)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ControlOp {
    Arithmetic,
    Sin,
    ImSum,
    DSum,
}

impl ControlOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Arithmetic => "ARITHMETIC",
            Self::Sin => "SIN",
            Self::ImSum => "IMSUM",
            Self::DSum => "DSUM",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    Control(ControlOp),
    Aggregate(AggregateOp),
    Discrete(DiscreteOp),
    DiscreteReducer(DiscreteOp),
    NestedDiscrete(DiscreteOp),
}

impl Operation {
    const fn name(self) -> &'static str {
        match self {
            Self::Control(operation) => operation.name(),
            Self::Aggregate(operation) => operation.name(),
            Self::Discrete(operation)
            | Self::DiscreteReducer(operation)
            | Self::NestedDiscrete(operation) => operation.name(),
        }
    }

    const fn is_aggregate(self) -> bool {
        matches!(
            self,
            Self::Aggregate(_) | Self::DiscreteReducer(_) | Self::NestedDiscrete(_)
        )
    }

    const fn is_discrete(self) -> bool {
        matches!(
            self,
            Self::Discrete(_) | Self::DiscreteReducer(_) | Self::NestedDiscrete(_)
        )
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Shape {
    Scalar,
    LiteralArray { rows: usize, columns: usize },
    ReferenceArray { rows: usize, columns: usize },
}

impl Shape {
    const fn is_scalar_input(self) -> bool {
        matches!(self, Self::Scalar)
    }

    const fn is_reference(self) -> bool {
        matches!(self, Self::ReferenceArray { .. })
    }

    const fn dimensions(self) -> Option<(usize, usize)> {
        match self {
            Self::LiteralArray { rows, columns } | Self::ReferenceArray { rows, columns } => {
                Some((rows, columns))
            },
            Self::Scalar => None,
        }
    }

    const fn mode(self) -> Mode {
        match self {
            Self::Scalar | Self::LiteralArray { .. } | Self::ReferenceArray { .. } => Mode::Matrix,
        }
    }
}

#[derive(Debug)]
struct Case {
    name: String,
    source: String,
    shape: Shape,
    operation: Operation,
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
    database: bool,
    aggregate: bool,
    discrete: bool,
}

impl FixtureResolver {
    fn new(database: bool, aggregate: bool, discrete: bool) -> Self {
        Self {
            stats: ResolverStats {
                reads: AtomicU64::new(0),
            },
            database,
            aggregate,
            discrete,
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
        Ok((sheet == "Main").then_some(SheetExtent::new(2048, 16)))
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
        if sheet != "Main" || row >= 2048 || column >= 16 {
            return Ok(CellRead::Error(
                litchi_ods::codec::formula::evaluation::ScalarError::Reference,
            ));
        }
        // The DSUM control uses a small database and criteria range in the
        // same immutable provider.  Aggregate reference cases use A:D/E:H;
        // the criteria range in D1:D2 is only used by this named control.
        if self.database && column == 0 && row == 0 {
            return Ok(CellRead::Text("Name"));
        }
        if self.database && column == 1 && row == 0 {
            return Ok(CellRead::Text("Value"));
        }
        if self.database && column == 0 && (1..=3).contains(&row) {
            return Ok(CellRead::Text(match row {
                1 => "A",
                2 => "B",
                _ => "C",
            }));
        }
        if self.database && column == 1 && (1..=3).contains(&row) {
            return Ok(CellRead::Number(match row {
                1 => 7.0,
                2 => 1.0,
                _ => 4.0,
            }));
        }
        if self.database && column == 3 && row == 0 {
            return Ok(CellRead::Text("Value"));
        }
        if self.database && column == 3 && row == 1 {
            return Ok(CellRead::Text(">0"));
        }
        Ok(CellRead::Number(if self.aggregate {
            reference_value(row, column)
        } else if self.discrete {
            discrete_reference_value(row, column)
        } else {
            fixture_value(row.saturating_mul(4).saturating_add(column))
        }))
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
}

fn aggregate_value(index: usize, lane: usize) -> f64 {
    1.01 + ((index.saturating_add(lane.saturating_mul(5))) % 17) as f64 * 0.001
}

fn fixture_value(index: usize) -> f64 {
    0.17 + (index % 9) as f64 * 0.047
}

fn reference_value(row: usize, column: usize) -> f64 {
    // Pair ranges use A:D and E:H.  Unary ranges use A:D, so their values are
    // lane zero as well.  The remaining columns are deliberately deterministic
    // for bounds and diagnostic cases, though this profile never references
    // them.
    let lane = column / 4;
    let local_column = column % 4;
    aggregate_value(row.saturating_mul(4).saturating_add(local_column), lane)
}

fn discrete_reference_value(row: usize, column: usize) -> f64 {
    // Keep streamed reducer fixtures integral and bounded.  The first four
    // columns deliberately share factors while E:H carries a different lane,
    // making GCD/LCM/MULTINOMIAL reads observable without making ordinary
    // reference rows expensive to validate.
    let index = row.saturating_mul(4).saturating_add(column % 4);
    let base = if row != 0 {
        0
    } else {
        match column / 4 {
            0 => [6_u64, 10, 0, 1][index % 4],
            1 => [9_u64, 15, 0, 1][index % 4],
            _ => [4_u64, 8, 0, 1][index % 4],
        }
    };
    base as f64
}

fn literal_array(rows: usize, columns: usize, lane: usize) -> String {
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
            let index = row.saturating_mul(columns).saturating_add(column);
            let _ = write!(output, "{:.6}", aggregate_value(index, lane));
        }
    }
    output.push('}');
    output
}

fn discrete_literal_value(operation: DiscreteOp, index: usize, lane: usize) -> f64 {
    let slot = index % 9;
    match operation {
        DiscreteOp::Combin => {
            if lane == 0 {
                (12 + slot) as f64
            } else {
                (2 + slot % 4) as f64
            }
        },
        DiscreteOp::Combina => {
            if lane == 0 {
                (5 + slot % 5) as f64
            } else {
                (2 + slot % 4) as f64
            }
        },
        DiscreteOp::Fact => (20 + slot) as f64,
        DiscreteOp::FactDouble => (24 + slot) as f64,
        DiscreteOp::Gcd => {
            if index >= 4 {
                0.0
            } else if lane == 0 {
                [6, 10, 0, 1][index % 4] as f64
            } else {
                [9, 15, 0, 1][index % 4] as f64
            }
        },
        DiscreteOp::Lcm => {
            if index >= 4 {
                0.0
            } else if lane == 0 {
                [6, 10, 0, 1][index % 4] as f64
            } else {
                [9, 15, 0, 1][index % 4] as f64
            }
        },
        DiscreteOp::Multinomial => {
            if index >= 4 {
                0.0
            } else if lane == 0 {
                [6, 10, 0, 1][index % 4] as f64
            } else {
                [9, 15, 0, 1][index % 4] as f64
            }
        },
        DiscreteOp::Even | DiscreteOp::Odd => -2.5 + (slot % 5) as f64 * 0.75,
        DiscreteOp::Delta => {
            if lane == 0 {
                (slot % 4) as f64
            } else {
                (slot % 3) as f64
            }
        },
        DiscreteOp::GStep => {
            if lane == 0 {
                -1.0 + slot as f64 * 0.5
            } else {
                0.5 + (slot % 3) as f64 * 0.5
            }
        },
    }
}

fn literal_discrete_array(
    rows: usize,
    columns: usize,
    operation: DiscreteOp,
    lane: usize,
) -> String {
    let mut output = String::with_capacity(rows.saturating_mul(columns).saturating_mul(10) + 2);
    output.push('{');
    for row in 0..rows {
        if row != 0 {
            output.push('|');
        }
        for column in 0..columns {
            if column != 0 {
                output.push(';');
            }
            let index = row.saturating_mul(columns).saturating_add(column);
            let _ = write!(
                output,
                "{:.6}",
                discrete_literal_value(operation, index, lane)
            );
        }
    }
    output.push('}');
    output
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
            let index = row.saturating_mul(columns).saturating_add(column);
            let _ = write!(output, "{:.6}", fixture_value(index));
        }
    }
    output.push('}');
    output
}

fn reference_range(rows: usize, columns: usize, lane: usize) -> String {
    let start_column = lane.saturating_mul(4);
    let end_column = start_column.saturating_add(columns.saturating_sub(1));
    format!(
        "[.{start}1: .{end}{rows}]",
        start = a1_column(start_column),
        end = a1_column(end_column),
    )
    .replace(' ', "")
}

fn a1_column(mut column: usize) -> String {
    let mut result = String::new();
    loop {
        let digit = (column % 26) as u8;
        result.push((b'A' + digit) as char);
        column /= 26;
        if column == 0 {
            break;
        }
        column -= 1;
    }
    result.chars().rev().collect()
}

fn discrete_source(operation: DiscreteOp, arguments: &[&str]) -> String {
    format!("={}({})", operation.name(), arguments.join(";"))
}

fn discrete_scalar_source(operation: DiscreteOp) -> String {
    match operation {
        DiscreteOp::Combin => "=COMBIN(1028;514)".to_owned(),
        // Keep the COMBINA timing lane finite; 1499 choose 500 exceeds
        // binary64 even though the inputs themselves are finite.
        DiscreteOp::Combina => "=COMBINA(1000;2)".to_owned(),
        DiscreteOp::Fact => "=FACT(100)".to_owned(),
        DiscreteOp::FactDouble => "=FACTDOUBLE(100)".to_owned(),
        DiscreteOp::Gcd => "=GCD(9007199254740992;9007199254740991;48)".to_owned(),
        // The first pair has an LCM above binary64's finite range.  Zero is
        // absorbing, so this specifically exercises overflow-then-zero.
        DiscreteOp::Lcm => "=LCM(1e308;3;0)".to_owned(),
        DiscreteOp::Multinomial => "=MULTINOMIAL(12;8;4)".to_owned(),
        DiscreteOp::Even => "=EVEN(-2.5)".to_owned(),
        DiscreteOp::Odd => "=ODD(-2.5)".to_owned(),
        DiscreteOp::Delta => "=DELTA(1;1)".to_owned(),
        DiscreteOp::GStep => "=GESTEP(1;0)".to_owned(),
    }
}

fn discrete_array_source(operation: DiscreteOp, rows: usize, columns: usize) -> String {
    let left = literal_discrete_array(rows, columns, operation, 0);
    if operation.binary() || operation.reducer() {
        let right = literal_discrete_array(rows, columns, operation, 1);
        discrete_source(operation, &[&left, &right])
    } else {
        discrete_source(operation, &[&left])
    }
}

fn scalar_source(operation: Operation) -> String {
    match operation {
        Operation::Control(ControlOp::Arithmetic) => "=0.17+1.25".to_owned(),
        Operation::Control(ControlOp::Sin) => "=SIN(0.17)".to_owned(),
        Operation::Control(ControlOp::ImSum) => {
            "=IMREAL(IMSUM(COMPLEX(2;3);COMPLEX(1;4)))".to_owned()
        },
        Operation::Control(ControlOp::DSum) => "=DSUM([.A1:.B4];2;[.D1:.D2])".to_owned(),
        Operation::Aggregate(AggregateOp::Sum) => "=SUM(1.25)".to_owned(),
        Operation::Discrete(operation)
        | Operation::DiscreteReducer(operation)
        | Operation::NestedDiscrete(operation) => discrete_scalar_source(operation),
    }
}

fn control_array_source(operation: ControlOp, argument: &str) -> String {
    match operation {
        ControlOp::Arithmetic => format!("={argument}+1.25"),
        ControlOp::Sin => format!("=SIN({argument})"),
        ControlOp::ImSum | ControlOp::DSum => unreachable!("scalar-only control"),
    }
}

fn all_cases() -> Vec<Case> {
    let mut cases = Vec::new();
    for operation in [
        ControlOp::Arithmetic,
        ControlOp::Sin,
        ControlOp::ImSum,
        ControlOp::DSum,
    ] {
        cases.push(Case {
            name: if operation == ControlOp::DSum {
                "database-control-dsum".to_owned()
            } else {
                format!("scalar-control-{}", operation.name().to_ascii_lowercase())
            },
            source: scalar_source(Operation::Control(operation)),
            shape: if operation == ControlOp::DSum {
                Shape::ReferenceArray {
                    rows: 4,
                    columns: 2,
                }
            } else {
                Shape::Scalar
            },
            operation: Operation::Control(operation),
        });
    }

    for (rows, columns) in [(4, 4)] {
        for operation in [ControlOp::Arithmetic, ControlOp::Sin] {
            let literal = literal_control_array(rows, columns);
            cases.push(Case {
                name: format!(
                    "array-control-{}x{}-{}",
                    rows,
                    columns,
                    operation.name().to_ascii_lowercase()
                ),
                source: control_array_source(operation, &literal),
                shape: Shape::LiteralArray { rows, columns },
                operation: Operation::Control(operation),
            });
        }
    }

    for operation in [ControlOp::Arithmetic, ControlOp::Sin] {
        let rows = 16;
        let columns = 4;
        let range = reference_range(rows, columns, 0);
        cases.push(Case {
            name: format!(
                "reference-array-{}x{}-{}",
                rows,
                columns,
                operation.name().to_ascii_lowercase()
            ),
            source: match operation {
                ControlOp::Arithmetic => format!("={range}+1.25"),
                ControlOp::Sin => format!("=SIN({range})"),
                _ => unreachable!(),
            },
            shape: Shape::ReferenceArray { rows, columns },
            operation: Operation::Control(operation),
        });
    }

    let operation = AggregateOp::Sum;
    cases.push(Case {
        name: format!("scalar-aggregate-{}", operation.name().to_ascii_lowercase()),
        source: scalar_source(Operation::Aggregate(operation)),
        shape: Shape::Scalar,
        operation: Operation::Aggregate(operation),
    });
    let (rows, columns) = (4, 4);
    let left = literal_array(rows, columns, 0);
    cases.push(Case {
        name: format!(
            "literal-aggregate-{}x{}-{}",
            rows,
            columns,
            operation.name().to_ascii_lowercase()
        ),
        source: format!("=SUM({left})"),
        shape: Shape::LiteralArray { rows, columns },
        operation: Operation::Aggregate(operation),
    });
    let (rows, columns) = (16, 4);
    let left = reference_range(rows, columns, 0);
    cases.push(Case {
        name: format!(
            "reference-aggregate-{}x{}-{}",
            rows,
            columns,
            operation.name().to_ascii_lowercase()
        ),
        source: format!("=SUM({left})"),
        shape: Shape::ReferenceArray { rows, columns },
        operation: Operation::Aggregate(operation),
    });

    // Every discrete function has a scalar and one bounded literal-array
    // rows.  Reducers additionally get streamed local-reference rows below;
    // those rows intentionally remain scalar results even though their input
    // shape is rectangular.
    for operation in DiscreteOp::all().iter().copied() {
        cases.push(Case {
            name: format!("scalar-discrete-{}", operation.name().to_ascii_lowercase()),
            source: discrete_scalar_source(operation),
            shape: Shape::Scalar,
            operation: Operation::Discrete(operation),
        });
        let (rows, columns) = (4, 4);
        cases.push(Case {
            name: format!(
                "literal-discrete-{}x{}-{}",
                rows,
                columns,
                operation.name().to_ascii_lowercase()
            ),
            source: discrete_array_source(operation, rows, columns),
            shape: Shape::LiteralArray { rows, columns },
            operation: if operation.reducer() {
                Operation::DiscreteReducer(operation)
            } else {
                Operation::Discrete(operation)
            },
        });
    }

    for operation in [DiscreteOp::Gcd, DiscreteOp::Lcm, DiscreteOp::Multinomial] {
        for rows in [64, 256, 1024] {
            let left = reference_range(rows, 4, 0);
            let right = reference_range(rows, 4, 1);
            cases.push(Case {
                name: format!(
                    "reference-discrete-{}x4-{}",
                    rows,
                    operation.name().to_ascii_lowercase()
                ),
                source: discrete_source(operation, &[&left, &right]),
                shape: Shape::ReferenceArray { rows, columns: 4 },
                operation: Operation::DiscreteReducer(operation),
            });
        }
        for rows in [64, 256, 1024] {
            let left = reference_range(rows, 4, 0);
            let right = reference_range(rows, 4, 1);
            let source = format!(
                "=SUM(IF({left}+1;{}({left}+0;{right});0))",
                operation.name()
            );
            cases.push(Case {
                name: format!(
                    "nested-discrete-{}-{}",
                    rows,
                    operation.name().to_ascii_lowercase()
                ),
                source,
                shape: Shape::ReferenceArray { rows, columns: 4 },
                operation: Operation::NestedDiscrete(operation),
            });
        }
    }

    cases
}

fn default_repeat(shape: Shape) -> usize {
    match shape {
        Shape::Scalar => 1_000,
        Shape::LiteralArray {
            rows: 4,
            columns: 4,
        }
        | Shape::ReferenceArray {
            rows: 4,
            columns: 4,
        }
        | Shape::ReferenceArray {
            rows: 16,
            columns: 4,
        } => 80,
        Shape::LiteralArray {
            rows: 16,
            columns: 16,
        }
        | Shape::ReferenceArray {
            rows: 64,
            columns: 4,
        } => 4,
        Shape::ReferenceArray {
            rows: 256,
            columns: 4,
        } => 2,
        Shape::ReferenceArray {
            rows: 1024,
            columns: 4,
        } => 1,
        Shape::LiteralArray { .. } | Shape::ReferenceArray { .. } => 1,
    }
}

fn execution() -> ExecutionContext {
    let budget = Budget::root(
        "ods-formula-discrete-math-performance",
        CoreLimits::for_profile(Profile::Server),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(1_u64 << 40).expect("finite in-flight bytes"),
        0,
    )
    .expect("valid execution limits");
    ExecutionContext::new(budget, token, limits)
}

fn approximately_equal(actual: f64, expected: f64) -> bool {
    (actual - expected).abs() <= expected.abs().max(1.0) * 1.0e-10
}

fn finite_expected(case: &Case, expected: f64) -> AnyResult<f64> {
    if expected.is_finite() {
        Ok(expected)
    } else {
        Err(format!(
            "{} produced a non-finite expected value: {expected}",
            case.name
        )
        .into())
    }
}

fn as_integer(value: f64) -> u128 {
    value.abs().round() as u128
}

fn gcd_integer(mut left: u128, mut right: u128) -> u128 {
    while right != 0 {
        let remainder = left % right;
        left = right;
        right = remainder;
    }
    left
}

fn lcm_integer(left: u128, right: u128) -> u128 {
    if left == 0 || right == 0 {
        0
    } else {
        left / gcd_integer(left, right) * right
    }
}

fn combin_f64(n: u128, k: u128) -> f64 {
    if k > n {
        return 0.0;
    }
    let k = k.min(n - k);
    let mut result = 1.0;
    for index in 1..=k {
        result *= (n - k + index) as f64;
        result /= index as f64;
    }
    result
}

// This is the rounded binary64 value of the exact integer in
// numeric-goldens.json for COMBIN(1028;514).  The ordinary multiply-then-
// divide recurrence can overflow an intermediate product even though this
// final result is finite, so the scalar hard case uses the retained exact
// golden instead of that recurrence.
const COMBIN_1028_514_GOLDEN: f64 = f64::from_bits(0x7fd9_79f4_8681_bf35);

fn discrete_apply(operation: DiscreteOp, arguments: &[f64]) -> f64 {
    let integers: Vec<u128> = arguments.iter().map(|value| as_integer(*value)).collect();
    match operation {
        DiscreteOp::Combin => combin_f64(integers[0], integers[1]),
        DiscreteOp::Combina => combin_f64(
            integers[0].saturating_add(integers[1]).saturating_sub(1),
            integers[1],
        ),
        DiscreteOp::Fact => (1..=integers[0]).fold(1.0, |result, value| result * value as f64),
        DiscreteOp::FactDouble => {
            let mut result = 1.0;
            let mut value = integers[0];
            while value > 1 {
                result *= value as f64;
                value = value.saturating_sub(2);
            }
            result
        },
        DiscreteOp::Gcd => integers.into_iter().fold(0, gcd_integer) as f64,
        DiscreteOp::Lcm => integers.into_iter().fold(1, lcm_integer) as f64,
        DiscreteOp::Multinomial => {
            let mut total = 0_u128;
            let mut result = 1.0;
            for value in integers {
                result *= combin_f64(total.saturating_add(value), value);
                total = total.saturating_add(value);
            }
            result
        },
        DiscreteOp::Even => {
            let mut magnitude = arguments[0].abs().ceil();
            if !(magnitude as u128).is_multiple_of(2) {
                magnitude += 1.0;
            }
            if arguments[0].is_sign_negative() {
                -magnitude
            } else {
                magnitude
            }
        },
        DiscreteOp::Odd => {
            let mut magnitude = arguments[0].abs().ceil();
            if (magnitude as u128).is_multiple_of(2) {
                magnitude += 1.0;
            }
            if arguments[0].is_sign_negative() {
                -magnitude
            } else {
                magnitude
            }
        },
        DiscreteOp::Delta => f64::from(arguments[0] == arguments[1]),
        DiscreteOp::GStep => f64::from(arguments[0] >= arguments[1]),
    }
}

fn discrete_arguments(operation: DiscreteOp, index: usize, reference: bool) -> Vec<f64> {
    if reference {
        let row = index / 4;
        let column = index % 4;
        vec![
            discrete_reference_value(row, column),
            discrete_reference_value(row, column + 4),
        ]
    } else if operation.binary() || operation.reducer() {
        vec![
            discrete_literal_value(operation, index, 0),
            discrete_literal_value(operation, index, 1),
        ]
    } else {
        vec![discrete_literal_value(operation, index, 0)]
    }
}

fn expected_discrete(case: &Case) -> f64 {
    let operation = match case.operation {
        Operation::Discrete(operation)
        | Operation::DiscreteReducer(operation)
        | Operation::NestedDiscrete(operation) => operation,
        _ => unreachable!("discrete oracle called for non-discrete operation"),
    };
    if matches!(case.operation, Operation::DiscreteReducer(_)) {
        let (rows, columns) = case.shape.dimensions().expect("reducer dimensions");
        let mut arguments = if operation.reducer() {
            vec![Vec::new(), Vec::new()]
        } else {
            vec![Vec::new()]
        };
        for index in 0..rows.saturating_mul(columns) {
            let values = discrete_arguments(operation, index, case.shape.is_reference());
            for (slot, value) in values.into_iter().enumerate() {
                if slot < arguments.len() {
                    arguments[slot].push(value);
                }
            }
        }
        let flattened: Vec<f64> = arguments.into_iter().flatten().collect();
        return discrete_apply(operation, &flattened);
    }
    if matches!(case.operation, Operation::NestedDiscrete(_)) {
        let (rows, columns) = case.shape.dimensions().expect("nested dimensions");
        // The branch reducer receives the two reference matrices.  `+0` keeps
        // the arithmetic projection explicit without changing the bounded
        // integer fixture, so each outer IF cell repeats the same scalar
        // result when projection is implemented by the evaluator.
        let mut arguments = vec![Vec::new(), Vec::new()];
        for index in 0..rows.saturating_mul(columns) {
            let values = discrete_arguments(operation, index, true);
            arguments[0].push(values[0]);
            arguments[1].push(values[1]);
        }
        let flattened = arguments.into_iter().flatten().collect::<Vec<_>>();
        let reducer = discrete_apply(operation, &flattened);
        return reducer * rows.saturating_mul(columns) as f64;
    }
    if case.shape == Shape::Scalar {
        return match operation {
            DiscreteOp::Combin => COMBIN_1028_514_GOLDEN,
            DiscreteOp::Combina => discrete_apply(operation, &[1000.0, 2.0]),
            DiscreteOp::Fact => discrete_apply(operation, &[100.0]),
            DiscreteOp::FactDouble => discrete_apply(operation, &[100.0]),
            DiscreteOp::Gcd => {
                discrete_apply(operation, &[9007199254740992.0, 9007199254740991.0, 48.0])
            },
            // Do not cast 1e308 through the integer helper in this oracle.
            // LCM is zero as soon as the final zero operand is observed.
            DiscreteOp::Lcm => 0.0,
            DiscreteOp::Multinomial => discrete_apply(operation, &[12.0, 8.0, 4.0]),
            DiscreteOp::Even => discrete_apply(operation, &[-2.5]),
            DiscreteOp::Odd => discrete_apply(operation, &[-2.5]),
            DiscreteOp::Delta => discrete_apply(operation, &[1.0, 1.0]),
            DiscreteOp::GStep => discrete_apply(operation, &[1.0, 0.0]),
        };
    }
    let (rows, columns) = case.shape.dimensions().expect("literal dimensions");
    let mut total = 0.0;
    for index in 0..rows.saturating_mul(columns) {
        total += discrete_apply(operation, &discrete_arguments(operation, index, false));
    }
    total
}

fn expected_aggregate(case: &Case) -> f64 {
    if !matches!(case.operation, Operation::Aggregate(AggregateOp::Sum)) {
        unreachable!("aggregate oracle called for non-SUM operation");
    }
    match case.shape {
        Shape::Scalar => 1.25,
        Shape::LiteralArray { rows, columns } | Shape::ReferenceArray { rows, columns } => (0
            ..rows.saturating_mul(columns))
            .map(|index| aggregate_value(index, 0))
            .sum(),
    }
}

fn expected_control_scalar(operation: ControlOp) -> f64 {
    match operation {
        ControlOp::Arithmetic => 1.42,
        ControlOp::Sin => 0.17_f64.sin(),
        ControlOp::ImSum => 3.0,
        ControlOp::DSum => 12.0,
    }
}

fn expected_control_cell(operation: ControlOp, index: usize) -> f64 {
    match operation {
        ControlOp::Arithmetic => fixture_value(index) + 1.25,
        ControlOp::Sin => fixture_value(index).sin(),
        ControlOp::ImSum => unreachable!("array-only control oracle"),
        ControlOp::DSum => 12.0,
    }
}

fn scalar_checksum(value: &ScalarValue<'_>) -> AnyResult<u64> {
    match value {
        ScalarValue::Number(value) => Ok(value.to_bits().rotate_left(11)),
        other => Err(format!("expected scalar Number, got {other:?}").into()),
    }
}

fn value_checksum(value: Value<'_>) -> AnyResult<u64> {
    match value {
        Value::Number(value) => Ok(value.to_bits().rotate_left(11)),
        other => Err(format!("expected Number, got {other:?}").into()),
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
    let ScalarValue::Number(actual) = value else {
        return Err(format!("{} returned {value:?}", case.name).into());
    };
    let expected = match case.operation {
        Operation::Control(operation) => expected_control_scalar(operation),
        Operation::Aggregate(_) => expected_aggregate(case),
        Operation::Discrete(_) | Operation::DiscreteReducer(_) | Operation::NestedDiscrete(_) => {
            expected_discrete(case)
        },
    };
    let expected = finite_expected(case, expected)?;
    if approximately_equal(*actual, expected) {
        Ok(())
    } else {
        Err(format!("{} returned {actual}, expected {expected}", case.name).into())
    }
}

fn validate_value(case: &Case, result: &Evaluated<'_>) -> AnyResult<()> {
    if case.operation.is_aggregate() {
        let actual = match result.value() {
            Value::Number(value) => value,
            value => return Err(format!("{} returned {value:?}", case.name).into()),
        };
        let expected = if case.operation.is_discrete() {
            expected_discrete(case)
        } else {
            expected_aggregate(case)
        };
        let expected = finite_expected(case, expected)?;
        if !approximately_equal(actual, expected) {
            return Err(format!("{} returned {actual}, expected {expected}", case.name).into());
        }
        return Ok(());
    }
    if case.operation.is_discrete() {
        let array = result.as_array().ok_or("discrete value was not an array")?;
        let (rows, columns) = case.shape.dimensions().ok_or("missing array dimensions")?;
        if (array.shape().rows(), array.shape().columns()) != (rows, columns) {
            return Err(format!("{} returned wrong shape", case.name).into());
        }
        let operation = match case.operation {
            Operation::Discrete(operation) => operation,
            _ => unreachable!("reducer should have been handled above"),
        };
        for index in 0..rows.saturating_mul(columns) {
            let actual = match array.get(index).ok_or("missing array cell")? {
                Value::Number(value) => value,
                value => return Err(format!("{} cell {index}: {value:?}", case.name).into()),
            };
            let expected = finite_expected(
                case,
                discrete_apply(operation, &discrete_arguments(operation, index, false)),
            )?;
            if !approximately_equal(actual, expected) {
                return Err(format!(
                    "{} cell {index} returned {actual}, expected {expected}",
                    case.name
                )
                .into());
            }
        }
        return Ok(());
    }
    if case.operation == Operation::Control(ControlOp::DSum) {
        let actual = match result.value() {
            Value::Number(value) => value,
            value => return Err(format!("{} returned {value:?}", case.name).into()),
        };
        if !approximately_equal(actual, expected_control_scalar(ControlOp::DSum)) {
            return Err(format!("{} returned {actual}", case.name).into());
        }
        return Ok(());
    }
    if case.shape.is_scalar_input() {
        let actual = match result.value() {
            Value::Number(value) => value,
            value => return Err(format!("{} returned {value:?}", case.name).into()),
        };
        let operation = match case.operation {
            Operation::Control(operation) => operation,
            _ => unreachable!(),
        };
        if !approximately_equal(actual, expected_control_cell(operation, 0)) {
            return Err(format!("{} returned {actual}", case.name).into());
        }
        return Ok(());
    }
    let array = result.as_array().ok_or("control value was not an array")?;
    let (rows, columns) = case.shape.dimensions().ok_or("missing array dimensions")?;
    if (array.shape().rows(), array.shape().columns()) != (rows, columns) {
        return Err(format!("{} returned wrong shape", case.name).into());
    }
    let operation = match case.operation {
        Operation::Control(operation) => operation,
        _ => unreachable!(),
    };
    for index in 0..rows.saturating_mul(columns) {
        let actual = match array.get(index).ok_or("missing array cell")? {
            Value::Number(value) => value,
            value => return Err(format!("{} cell {index}: {value:?}", case.name).into()),
        };
        let expected = expected_control_cell(operation, index);
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

fn eval_scalar<'a>(
    expression: &'a Expression,
    execution: &ExecutionContext,
) -> Result<litchi_ods::codec::formula::evaluation::EvaluatedScalar<'a>, EvaluationFailure> {
    evaluate_scalar(
        expression,
        &EvaluationContext::new(execution),
        &EvaluationLimits::default(),
    )
}

fn eval_value<'a>(
    case: &Case,
    expression: &'a Expression,
    resolver: &'a FixtureResolver,
    execution: &ExecutionContext,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(case.shape.mode());
    value::evaluate(expression, resolver, &context, &Limits::default())
}

fn preflight(case: &Case) -> AnyResult<()> {
    let expression = Expression::parse(&case.source).map_err(|error| {
        format!(
            "{} parse preflight failed for {:?}: {error}",
            case.name, case.source
        )
    })?;
    let execution = execution();
    if case.shape.is_reference() || !case.shape.is_scalar_input() {
        let resolver = FixtureResolver::new(
            case.operation == Operation::Control(ControlOp::DSum),
            case.operation.is_aggregate() && !case.operation.is_discrete(),
            case.operation.is_discrete(),
        );
        let result = eval_value(case, &expression, &resolver, &execution)
            .map_err(|error| format!("{} value preflight failed: {error}", case.name))?;
        validate_value(case, &result)?;
    } else {
        let result = eval_scalar(&expression, &execution)
            .map_err(|error| format!("{} scalar preflight failed: {error}", case.name))?;
        validate_scalar(case, result.value())?;
    }
    Ok(())
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
}

fn parse_config() -> AnyResult<Option<Config>> {
    let mut case = None;
    let mut phase = Phase::Evaluate;
    let mut warmups = 3;
    let mut iterations = 1;
    let mut repeat = None;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                println!(
                    "usage: ods-formula-discrete-math-performance [--case NAME|all] [--phase evaluate|parse-evaluate] [--warmups N] [--iterations N] [--repeat N]"
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
                    .parse()?;
            },
            "--iterations" => {
                iterations = arguments
                    .next()
                    .ok_or("--iterations requires an integer")?
                    .parse()?;
            },
            "--repeat" => {
                repeat = Some(
                    arguments
                        .next()
                        .ok_or("--repeat requires an integer")?
                        .parse()?,
                );
            },
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
    result: Result<litchi_ods::codec::formula::evaluation::EvaluatedScalar<'_>, EvaluationFailure>,
    execution: &ExecutionContext,
    checksum: &mut u64,
) -> AnyResult<u64> {
    let result = result.map_err(|error| format!("scalar evaluation failed: {error}"))?;
    *checksum = checksum.wrapping_add(scalar_checksum(result.value())?);
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
) -> AnyResult<u64> {
    let result = result.map_err(|error| format!("value evaluation failed: {error}"))?;
    if case.operation.is_aggregate()
        || case.shape.is_scalar_input()
        || case.operation == Operation::Control(ControlOp::DSum)
    {
        *checksum = checksum.wrapping_add(value_checksum(result.value())?);
    } else {
        *checksum =
            checksum.wrapping_add(array_checksum(result.as_array().ok_or("missing array")?)?);
    }
    black_box(*checksum);
    let retained = execution.budget().used(Resource::Memory);
    drop(result);
    Ok(retained)
}

fn sample_from(
    started: Instant,
    execution: &ExecutionContext,
    resolver: &FixtureResolver,
    baseline_work: u64,
    baseline_memory: u64,
    live_before: u64,
    checksum: u64,
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
        checksum,
    }
}

fn measure_evaluate(case: &Case, expression: &Expression, repeat: usize) -> AnyResult<Sample> {
    let execution = execution();
    let resolver = FixtureResolver::new(
        case.operation == Operation::Control(ControlOp::DSum),
        case.operation.is_aggregate() && !case.operation.is_discrete(),
        case.operation.is_discrete(),
    );
    resolver.stats.reset();
    let context =
        Context::new(&execution, Position::new("Main", 0, 0)).with_mode(case.shape.mode());
    let scalar_context = EvaluationContext::new(&execution);
    let scalar_limits = EvaluationLimits::default();
    let value_limits = Limits::default();
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut checksum = 0_u64;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        retained_peak = if case.shape.is_scalar_input() && !case.shape.is_reference() {
            retained_peak.max(consume_scalar(
                evaluate_scalar(expression, &scalar_context, &scalar_limits),
                &execution,
                &mut checksum,
            )?)
        } else {
            retained_peak.max(consume_value(
                value::evaluate(expression, &resolver, &context, &value_limits),
                case,
                &execution,
                &mut checksum,
            )?)
        };
    }
    Ok(sample_from(
        started,
        &execution,
        &resolver,
        baseline_work,
        baseline_memory,
        live_before,
        checksum,
        retained_peak,
    ))
}

fn measure_parse_evaluate(case: &Case, repeat: usize) -> AnyResult<Sample> {
    let execution = execution();
    let resolver = FixtureResolver::new(
        case.operation == Operation::Control(ControlOp::DSum),
        case.operation.is_aggregate() && !case.operation.is_discrete(),
        case.operation.is_discrete(),
    );
    let context =
        Context::new(&execution, Position::new("Main", 0, 0)).with_mode(case.shape.mode());
    let scalar_context = EvaluationContext::new(&execution);
    let scalar_limits = EvaluationLimits::default();
    let value_limits = Limits::default();
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut checksum = 0_u64;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        let expression = Expression::parse(black_box(&case.source))
            .map_err(|error| format!("{} parse failed: {error}", case.name))?;
        retained_peak = if case.shape.is_scalar_input() && !case.shape.is_reference() {
            retained_peak.max(consume_scalar(
                evaluate_scalar(&expression, &scalar_context, &scalar_limits),
                &execution,
                &mut checksum,
            )?)
        } else {
            retained_peak.max(consume_value(
                value::evaluate(&expression, &resolver, &context, &value_limits),
                case,
                &execution,
                &mut checksum,
            )?)
        };
        retained_peak = retained_peak.max(execution.budget().used(Resource::Memory));
        black_box(expression);
    }
    Ok(sample_from(
        started,
        &execution,
        &resolver,
        baseline_work,
        baseline_memory,
        live_before,
        checksum,
        retained_peak,
    ))
}

fn percentile(samples: &[Sample], metric: impl Fn(&Sample) -> u64, percentile: usize) -> u64 {
    let mut values: Vec<u64> = samples.iter().map(metric).collect();
    values.sort_unstable();
    values[(values.len().saturating_sub(1) * percentile) / 100]
}

fn mean(samples: &[Sample], metric: impl Fn(&Sample) -> u64) -> u64 {
    let total: u128 = samples
        .iter()
        .map(|sample| u128::from(metric(sample)))
        .sum();
    (total / samples.len() as u128).min(u64::MAX as u128) as u64
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
                let _ = write!(escaped, "\\u{:04x}", character as u32);
            },
            character => escaped.push(character),
        }
    }
    escaped
}

fn sample_json(sample: &Sample) -> String {
    format!(
        "{{\"elapsed_ns\":{},\"alloc_calls\":{},\"dealloc_calls\":{},\"requested_bytes\":{},\"released_bytes\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"work\":{},\"memory_retained\":{},\"reference_reads\":{},\"checksum\":{}}}",
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
        sample.checksum,
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
        Shape::LiteralArray { rows, columns } | Shape::ReferenceArray { rows, columns } => {
            ("array", rows, columns, rows.saturating_mul(columns))
        },
    };
    let mut raw_samples = String::from("[");
    for (index, sample) in samples.iter().enumerate() {
        if index != 0 {
            raw_samples.push(',');
        }
        raw_samples.push_str(&sample_json(sample));
    }
    raw_samples.push(']');
    println!(
        "{{\"case\":\"{}\",\"operation\":\"{}\",\"phase\":\"{}\",\"input_bytes\":{},\"repeat\":{},\"warmups\":{},\"iterations\":{},\"shape\":\"{}\",\"rows\":{},\"columns\":{},\"elements\":{},\"supported\":true,\"elapsed_ns_p50\":{},\"elapsed_ns_mean\":{},\"elapsed_ns_p95\":{},\"elapsed_ns_p99\":{},\"elapsed_ns_per_repeat\":{},\"allocator_calls_p50\":{},\"requested_bytes_p50\":{},\"released_bytes_p50\":{},\"peak_live_delta_p50\":{},\"memory_retained_p50\":{},\"work_p50\":{},\"work_per_repeat\":{},\"reference_reads_p50\":{},\"reference_reads_per_repeat\":{},\"checksum_p50\":{},\"live_before_p50\":{},\"live_after_p50\":{},\"samples\":{},\"rss_kib\":null,\"rss_source\":\"external /usr/bin/time -v\",\"memory_retained_source\":\"execution budget while result is live\",\"reference_source\":\"instrumented borrowing resolver\",\"validation_scope\":\"one untimed direct f64 fixture oracle; timed evaluator/checksum/drop\"}}",
        json_escape(&case.name),
        case.operation.name(),
        phase.label(),
        case.source.len(),
        repeat,
        warmups,
        iterations,
        shape,
        rows,
        columns,
        elements,
        percentile(samples, |sample| sample.elapsed_ns, 50),
        mean(samples, |sample| sample.elapsed_ns),
        percentile(samples, |sample| sample.elapsed_ns, 95),
        percentile(samples, |sample| sample.elapsed_ns, 99),
        percentile(samples, |sample| sample.elapsed_ns, 50) / repeat as u64,
        percentile(samples, |sample| sample.alloc_calls, 50),
        percentile(samples, |sample| sample.requested_bytes, 50),
        percentile(samples, |sample| sample.released_bytes, 50),
        percentile(samples, |sample| sample.peak_live_delta, 50),
        percentile(samples, |sample| sample.memory_retained, 50),
        percentile(samples, |sample| sample.work, 50),
        percentile(samples, |sample| sample.work, 50) / repeat as u64,
        percentile(samples, |sample| sample.reference_reads, 50),
        percentile(samples, |sample| sample.reference_reads, 50) / repeat as u64,
        percentile(samples, |sample| sample.checksum, 50),
        percentile(samples, |sample| sample.live_before, 50),
        percentile(samples, |sample| sample.live_after, 50),
        raw_samples,
    );
}

fn selected_cases<'a>(all: &'a [Case], requested: Option<&str>) -> AnyResult<Vec<&'a Case>> {
    match requested {
        None | Some("all") => Ok(all.iter().collect()),
        Some(name) => all
            .iter()
            .find(|case| case.name == name)
            .map(|case| vec![case])
            .ok_or_else(|| format!("unknown case {name:?}; use --list").into()),
    }
}

fn run_case(case: &Case, config: &Config) -> AnyResult<()> {
    preflight(case)?;
    let repeat = config.repeat.unwrap_or_else(|| default_repeat(case.shape));
    let expression = Expression::parse(&case.source).map_err(|error| {
        format!(
            "{} parse preflight failed for {:?}: {error}",
            case.name, case.source
        )
    })?;
    for _ in 0..config.warmups {
        let sample = match config.phase {
            Phase::Evaluate => measure_evaluate(case, &expression, repeat)?,
            Phase::ParseEvaluate => measure_parse_evaluate(case, repeat)?,
        };
        black_box(sample);
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
    for case in selected_cases(&all, config.case.as_deref())? {
        run_case(case, &config)?;
    }
    Ok(())
}
