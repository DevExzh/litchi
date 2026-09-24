//! Bounded process-level profile for the ODS numeric aggregate evaluator.
//!
//! The fixture is deterministic and entirely in memory.  Every timed child
//! performs a correctness preflight before measuring one immutable expression
//! repeatedly.  Reference cases use a borrowing resolver and count cell reads
//! so a profile row can distinguish the sequence path from literal arrays.

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
    Product,
    SumSq,
    SumProduct,
    SumX2MY2,
    SumX2PY2,
    SumXMY2,
}

impl AggregateOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Sum => "SUM",
            Self::Product => "PRODUCT",
            Self::SumSq => "SUMSQ",
            Self::SumProduct => "SUMPRODUCT",
            Self::SumX2MY2 => "SUMX2MY2",
            Self::SumX2PY2 => "SUMX2PY2",
            Self::SumXMY2 => "SUMXMY2",
        }
    }

    const fn all() -> &'static [Self] {
        &[
            Self::Sum,
            Self::Product,
            Self::SumSq,
            Self::SumProduct,
            Self::SumX2MY2,
            Self::SumX2PY2,
            Self::SumXMY2,
        ]
    }

    const fn binary(self) -> bool {
        matches!(
            self,
            Self::SumProduct | Self::SumX2MY2 | Self::SumX2PY2 | Self::SumXMY2
        )
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
    AggregateThreeProduct,
    Nested(AggregateOp),
    NestedIf(AggregateOp),
}

impl Operation {
    const fn name(self) -> &'static str {
        match self {
            Self::Control(operation) => operation.name(),
            Self::Aggregate(operation) | Self::Nested(operation) | Self::NestedIf(operation) => {
                operation.name()
            },
            Self::AggregateThreeProduct => "SUMPRODUCT",
        }
    }

    const fn is_aggregate(self) -> bool {
        matches!(
            self,
            Self::Aggregate(_) | Self::AggregateThreeProduct | Self::Nested(_) | Self::NestedIf(_)
        )
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Shape {
    Scalar,
    LiteralArray { rows: usize, columns: usize },
    ReferenceScalar,
    ReferenceArray { rows: usize, columns: usize },
}

impl Shape {
    const fn is_scalar_input(self) -> bool {
        matches!(self, Self::Scalar | Self::ReferenceScalar)
    }

    const fn is_reference(self) -> bool {
        matches!(self, Self::ReferenceScalar | Self::ReferenceArray { .. })
    }

    const fn dimensions(self) -> Option<(usize, usize)> {
        match self {
            Self::LiteralArray { rows, columns } | Self::ReferenceArray { rows, columns } => {
                Some((rows, columns))
            },
            Self::Scalar | Self::ReferenceScalar => None,
        }
    }

    const fn mode(self) -> Mode {
        match self {
            Self::ReferenceScalar => Mode::Scalar,
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
}

impl FixtureResolver {
    fn new(database: bool, aggregate: bool) -> Self {
        Self {
            stats: ResolverStats {
                reads: AtomicU64::new(0),
            },
            database,
            aggregate,
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

fn aggregate_source(operation: AggregateOp, left: &str, right: Option<&str>) -> String {
    match (operation, right) {
        (AggregateOp::Sum, None) => format!("=SUM({left})"),
        (AggregateOp::Product, None) => format!("=PRODUCT({left})"),
        (AggregateOp::SumSq, None) => format!("=SUMSQ({left})"),
        (AggregateOp::SumProduct, Some(right)) => format!("=SUMPRODUCT({left};{right})"),
        (AggregateOp::SumX2MY2, Some(right)) => format!("=SUMX2MY2({left};{right})"),
        (AggregateOp::SumX2PY2, Some(right)) => format!("=SUMX2PY2({left};{right})"),
        (AggregateOp::SumXMY2, Some(right)) => format!("=SUMXMY2({left};{right})"),
        (operation, Some(_)) => panic!("unary aggregate {} received two inputs", operation.name()),
        (operation, None) => panic!("binary aggregate {} received one input", operation.name()),
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
        Operation::Aggregate(operation) => {
            let left = "1.25";
            let right = if operation.binary() {
                Some("1.011")
            } else {
                None
            };
            aggregate_source(operation, left, right)
        },
        Operation::AggregateThreeProduct => "=SUMPRODUCT(1.25;1.011;1.013)".to_owned(),
        Operation::Nested(operation) | Operation::NestedIf(operation) => {
            aggregate_source(operation, "1.25", Some("1.011"))
        },
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

    for (rows, columns) in [(4, 4), (16, 16)] {
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
        cases.push(Case {
            name: format!("reference-scalar-{}", operation.name().to_ascii_lowercase()),
            source: match operation {
                ControlOp::Arithmetic => "=[.A1]+1.25".to_owned(),
                ControlOp::Sin => "=SIN([.A1])".to_owned(),
                _ => unreachable!(),
            },
            shape: Shape::ReferenceScalar,
            operation: Operation::Control(operation),
        });
        cases.push(Case {
            name: format!("reference-array-{}", operation.name().to_ascii_lowercase()),
            source: match operation {
                ControlOp::Arithmetic => "=[.A1:.D4]+1.25".to_owned(),
                ControlOp::Sin => "=SIN([.A1:.D4])".to_owned(),
                _ => unreachable!(),
            },
            shape: Shape::ReferenceArray {
                rows: 4,
                columns: 4,
            },
            operation: Operation::Control(operation),
        });
    }

    for operation in AggregateOp::all().iter().copied() {
        cases.push(Case {
            name: format!("scalar-aggregate-{}", operation.name().to_ascii_lowercase()),
            source: scalar_source(Operation::Aggregate(operation)),
            shape: Shape::Scalar,
            operation: Operation::Aggregate(operation),
        });
        for (rows, columns) in [(4, 4), (16, 16)] {
            let left = literal_array(rows, columns, 0);
            let right = operation.binary().then(|| literal_array(rows, columns, 1));
            cases.push(Case {
                name: format!(
                    "literal-aggregate-{}x{}-{}",
                    rows,
                    columns,
                    operation.name().to_ascii_lowercase()
                ),
                source: aggregate_source(operation, &left, right.as_deref()),
                shape: Shape::LiteralArray { rows, columns },
                operation: Operation::Aggregate(operation),
            });
        }
        for (rows, columns) in [(16, 4), (64, 4), (256, 4), (1024, 4)] {
            let left = reference_range(rows, columns, 0);
            let right = operation
                .binary()
                .then(|| reference_range(rows, columns, 1));
            cases.push(Case {
                name: format!(
                    "reference-aggregate-{}x{}-{}",
                    rows,
                    columns,
                    operation.name().to_ascii_lowercase()
                ),
                source: aggregate_source(operation, &left, right.as_deref()),
                shape: Shape::ReferenceArray { rows, columns },
                operation: Operation::Aggregate(operation),
            });
        }
    }

    for (shape, left, right, third) in [
        (
            Shape::LiteralArray {
                rows: 4,
                columns: 4,
            },
            literal_array(4, 4, 0),
            literal_array(4, 4, 1),
            literal_array(4, 4, 2),
        ),
        (
            Shape::ReferenceArray {
                rows: 1024,
                columns: 4,
            },
            reference_range(1024, 4, 0),
            reference_range(1024, 4, 1),
            reference_range(1024, 4, 2),
        ),
    ] {
        let (label, rows, columns) = match shape {
            Shape::LiteralArray { rows, columns } | Shape::ReferenceArray { rows, columns } => (
                match shape {
                    Shape::LiteralArray { .. } => "literal",
                    Shape::ReferenceArray { .. } => "reference",
                    _ => unreachable!(),
                },
                rows,
                columns,
            ),
            _ => unreachable!(),
        };
        cases.push(Case {
            name: format!("{label}-aggregate-{rows}x{columns}-sumproduct-k3"),
            source: format!("=SUMPRODUCT({left};{right};{third})"),
            shape,
            operation: Operation::AggregateThreeProduct,
        });
    }

    for rows in [64, 256, 1024] {
        let left = reference_range(rows, 4, 0);
        let right = reference_range(rows, 4, 1);
        cases.push(Case {
            name: format!("nested-projected-{rows}-sumproduct"),
            source: format!("=SUMPRODUCT({left}+SUM({right});{right})"),
            shape: Shape::ReferenceArray { rows, columns: 4 },
            operation: Operation::Nested(AggregateOp::SumProduct),
        });
        cases.push(Case {
            name: format!("nested-if-projected-{rows}-sumproduct"),
            source: format!("=SUM(IF({left};SUMPRODUCT({left}+1;{right});0))"),
            shape: Shape::ReferenceArray { rows, columns: 4 },
            operation: Operation::NestedIf(AggregateOp::SumProduct),
        });
    }
    cases
}

fn default_repeat(shape: Shape) -> usize {
    match shape {
        Shape::Scalar | Shape::ReferenceScalar => 1_000,
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
        "ods-formula-aggregates-performance",
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

fn expected_aggregate(case: &Case) -> f64 {
    let (operation, nested, k3) = match case.operation {
        Operation::Aggregate(operation) => (operation, false, false),
        Operation::AggregateThreeProduct => (AggregateOp::SumProduct, false, true),
        Operation::Nested(operation) => (operation, true, false),
        Operation::NestedIf(operation) => (operation, false, false),
        Operation::Control(_) => unreachable!("aggregate oracle called for control"),
    };
    if matches!(case.shape, Shape::Scalar) {
        let left = 1.25;
        let right = 1.011;
        let third = 1.013;
        let right_values = [right];
        let third_values = [third];
        if k3 {
            return aggregate_values_three(&[left], &[right], &third_values);
        }
        return aggregate_values(
            operation,
            &[left],
            if operation.binary() {
                Some(&right_values)
            } else {
                None
            },
        );
    }
    let (rows, columns) = case.shape.dimensions().expect("aggregate dimensions");
    if matches!(case.operation, Operation::NestedIf(_)) {
        let inner: f64 = (0..rows.saturating_mul(columns))
            .map(|index| (aggregate_value(index, 0) + 1.0) * aggregate_value(index, 1))
            .sum();
        return rows.saturating_mul(columns) as f64 * inner;
    }
    if k3 {
        let mut left = Vec::with_capacity(rows.saturating_mul(columns));
        let mut right = Vec::with_capacity(rows.saturating_mul(columns));
        let mut third = Vec::with_capacity(rows.saturating_mul(columns));
        for index in 0..rows.saturating_mul(columns) {
            left.push(aggregate_value(index, 0));
            right.push(aggregate_value(index, 1));
            third.push(aggregate_value(index, 2));
        }
        return aggregate_values_three(&left, &right, &third);
    }
    if nested {
        let mut inner = 0.0;
        for index in 0..rows.saturating_mul(columns) {
            inner += aggregate_value(index, 1);
        }
        return (0..rows.saturating_mul(columns))
            .map(|index| (aggregate_value(index, 0) + inner) * aggregate_value(index, 1))
            .sum();
    }
    let mut left = Vec::with_capacity(rows.saturating_mul(columns));
    let mut right = Vec::with_capacity(rows.saturating_mul(columns));
    for index in 0..rows.saturating_mul(columns) {
        left.push(aggregate_value(index, 0));
        if operation.binary() {
            right.push(aggregate_value(index, 1));
        }
    }
    aggregate_values(
        operation,
        &left,
        operation.binary().then_some(right.as_slice()),
    )
}

fn aggregate_values(operation: AggregateOp, left: &[f64], right: Option<&[f64]>) -> f64 {
    match operation {
        AggregateOp::Sum => left.iter().sum(),
        AggregateOp::Product => left.iter().product(),
        AggregateOp::SumSq => left.iter().map(|value| value * value).sum(),
        AggregateOp::SumProduct => left
            .iter()
            .zip(right.expect("binary aggregate right values"))
            .map(|(left, right)| left * right)
            .sum(),
        AggregateOp::SumX2MY2 => left
            .iter()
            .zip(right.expect("binary aggregate right values"))
            .map(|(left, right)| left * left - right * right)
            .sum(),
        AggregateOp::SumX2PY2 => left
            .iter()
            .zip(right.expect("binary aggregate right values"))
            .map(|(left, right)| left * left + right * right)
            .sum(),
        AggregateOp::SumXMY2 => left
            .iter()
            .zip(right.expect("binary aggregate right values"))
            .map(|(left, right)| (left - right) * (left - right))
            .sum(),
    }
}

fn aggregate_values_three(left: &[f64], right: &[f64], third: &[f64]) -> f64 {
    left.iter()
        .zip(right)
        .zip(third)
        .map(|((left, right), third)| left * right * third)
        .sum()
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
        Operation::Aggregate(_)
        | Operation::AggregateThreeProduct
        | Operation::Nested(_)
        | Operation::NestedIf(_) => expected_aggregate(case),
    };
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
        let expected = expected_aggregate(case);
        if !approximately_equal(actual, expected) {
            return Err(format!("{} returned {actual}, expected {expected}", case.name).into());
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
            case.operation.is_aggregate(),
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
                    "usage: ods-formula-aggregates-performance [--case NAME|all] [--phase evaluate|parse-evaluate] [--warmups N] [--iterations N] [--repeat N]"
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
        case.operation.is_aggregate(),
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
        case.operation.is_aggregate(),
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
        Shape::Scalar | Shape::ReferenceScalar => ("scalar", 0, 0, 1),
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
