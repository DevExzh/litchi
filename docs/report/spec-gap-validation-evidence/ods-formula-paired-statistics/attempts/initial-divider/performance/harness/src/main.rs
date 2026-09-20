//! Bounded process-level profile for the ODS paired-statistics evaluator.
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
}

impl AggregateOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Sum => "SUM",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum StatisticalOp {
    CountA,
    Average,
}

impl StatisticalOp {
    const fn name(self) -> &'static str {
        match self {
            Self::CountA => "COUNTA",
            Self::Average => "AVERAGE",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PairedOp {
    Correl,
    Covar,
    Pearson,
    Rsq,
    Slope,
    Intercept,
    Steyx,
    Forecast,
}

impl PairedOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Correl => "CORREL",
            Self::Covar => "COVAR",
            Self::Pearson => "PEARSON",
            Self::Rsq => "RSQ",
            Self::Slope => "SLOPE",
            Self::Intercept => "INTERCEPT",
            Self::Steyx => "STEYX",
            Self::Forecast => "FORECAST",
        }
    }

    const fn all() -> &'static [Self] {
        &[
            Self::Correl,
            Self::Covar,
            Self::Pearson,
            Self::Rsq,
            Self::Slope,
            Self::Intercept,
            Self::Steyx,
            Self::Forecast,
        ]
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DescriptiveOp {
    Avedev,
    Devsq,
    Kurt,
    Skew,
    Skewp,
}

impl DescriptiveOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Avedev => "AVEDEV",
            Self::Devsq => "DEVSQ",
            Self::Kurt => "KURT",
            Self::Skew => "SKEW",
            Self::Skewp => "SKEWP",
        }
    }

    const fn all() -> &'static [Self] {
        &[
            Self::Avedev,
            Self::Devsq,
            Self::Kurt,
            Self::Skew,
            Self::Skewp,
        ]
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ConditionalOp {
    SumIfs,
}

impl ConditionalOp {
    const fn name(self) -> &'static str {
        match self {
            Self::SumIfs => "SUMIFS",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum Expectation {
    Number(f64),
    Error(litchi_ods::codec::formula::evaluation::ScalarError),
    Failure(FailureExpectation),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FailureExpectation {
    ReferenceCells,
    Cancelled,
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
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RepresentativeOp {
    Median,
    Rank,
    PercentRank,
}

impl RepresentativeOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Median => "MEDIAN",
            Self::Rank => "RANK",
            Self::PercentRank => "PERCENTRANK",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    Control(ControlOp),
    Aggregate(AggregateOp),
    Conditional(ConditionalOp),
    Statistical(StatisticalOp),
    Representative(RepresentativeOp),
    Descriptive(DescriptiveOp),
    Paired(PairedOp),
    NestedPaired(PairedOp),
}

impl Operation {
    const fn name(self) -> &'static str {
        match self {
            Self::Control(operation) => operation.name(),
            Self::Aggregate(operation) => operation.name(),
            Self::Conditional(operation) => operation.name(),
            Self::Statistical(operation) => operation.name(),
            Self::Representative(operation) => operation.name(),
            Self::Descriptive(operation) => operation.name(),
            Self::Paired(operation) | Self::NestedPaired(operation) => operation.name(),
        }
    }

    const fn is_aggregate(self) -> bool {
        matches!(
            self,
            Self::Aggregate(_)
                | Self::Conditional(_)
                | Self::Statistical(_)
                | Self::Descriptive(_)
                | Self::Paired(_)
                | Self::NestedPaired(_)
        )
    }

    const fn is_conditional(self) -> bool {
        matches!(self, Self::Conditional(_))
    }

    const fn is_statistical(self) -> bool {
        matches!(
            self,
            Self::Statistical(_) | Self::Descriptive(_) | Self::Paired(_) | Self::NestedPaired(_)
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
    expectation: Expectation,
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
    conditional: bool,
    statistical: bool,
    descriptive_statistics: bool,
    paired_statistics: bool,
    cancel_after_read: Option<CancellationSource>,
}

impl FixtureResolver {
    fn new(
        database: bool,
        aggregate: bool,
        conditional: bool,
        statistical: bool,
        descriptive_statistics: bool,
        paired_statistics: bool,
        cancel_after_read: Option<CancellationSource>,
    ) -> Self {
        Self {
            stats: ResolverStats {
                reads: AtomicU64::new(0),
            },
            database,
            aggregate,
            conditional,
            statistical,
            descriptive_statistics,
            paired_statistics,
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
            return Ok(CellRead::Error(
                litchi_ods::codec::formula::evaluation::ScalarError::Reference,
            ));
        }
        // The DSUM control uses a small database and criteria range in the
        // same immutable provider. Conditional cases use four-cell lanes:
        // A:D is numeric criterion one, E:H is criterion two, I:L is the
        // selected numeric destination, and M:P is exact text criteria.
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
        if self.paired_statistics {
            let index = row.saturating_mul(4).saturating_add(column % 4);
            return Ok(paired_statistics_cell(column / 4, index, row));
        }
        if self.descriptive_statistics {
            let index = row.saturating_mul(4).saturating_add(column % 4);
            return Ok(descriptive_statistics_cell(column / 4, index, row));
        }
        if self.statistical {
            let index = row.saturating_mul(4).saturating_add(column % 4);
            let plane = match sheet {
                "Main" => 0,
                "Data" => 1,
                "Archive" => 2,
                _ => unreachable!("sheet was checked above"),
            };
            return Ok(statistical_cell(column / 4, index, plane));
        }
        if self.conditional {
            let index = row.saturating_mul(4).saturating_add(column % 4);
            let plane = match sheet {
                "Main" => 0,
                "Data" => 1,
                "Archive" => 2,
                _ => unreachable!("sheet was checked above"),
            };
            return Ok(match column / 4 {
                0 => CellRead::Number((index % 4 + plane) as f64),
                1 => CellRead::Number((index % 2 + plane) as f64),
                2 => CellRead::Number(0.5 + index as f64 * 0.03125 + plane as f64),
                3 => CellRead::Text(if (row + plane) % 2 == 0 { "A" } else { "B" }),
                _ => CellRead::Empty,
            });
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

fn statistical_cell(lane: usize, index: usize, plane: usize) -> CellRead<'static> {
    // A:D is a dense numeric lane; E:H deliberately mixes the types that the
    // A-suffixed reducers admit; I:L is an all-empty lane for identity and
    // COUNTBLANK rows; M:P carries one formula error followed by numbers.
    match lane {
        0 => CellRead::Number(((index % 17) as f64) - 8.0 + plane as f64 * 0.25),
        1 => match index % 4 {
            0 => CellRead::Number((index % 11) as f64 - 4.0),
            1 => CellRead::Text(if index % 8 == 1 { "" } else { "word" }),
            2 => CellRead::Logical(index % 2 == 0),
            _ => CellRead::Empty,
        },
        2 => CellRead::Empty,
        3 => {
            if index == 0 {
                CellRead::Error(litchi_ods::codec::formula::evaluation::ScalarError::NotAvailable)
            } else {
                CellRead::Number((index % 13) as f64 - 6.0)
            }
        },
        _ => CellRead::Empty,
    }
}

fn paired_statistics_number(lane: usize, index: usize, _row: usize) -> f64 {
    let x = (index % 257) as f64 + 1.0;
    match lane {
        0 | 1 => {
            if lane == 0 {
                2.0 * x + 1.0
            } else {
                x
            }
        },
        3 | 4 => {
            if lane == 3 {
                2.0 * x + 1.0
            } else {
                x
            }
        },
        _ => x,
    }
}

fn paired_statistics_cell(lane: usize, index: usize, _row: usize) -> CellRead<'static> {
    if lane == 3 && index % 4 == 1 {
        return CellRead::Text("skip");
    }
    if lane == 4 && index % 4 == 2 {
        return CellRead::Empty;
    }
    CellRead::Number(paired_statistics_number(lane, index, 0))
}

fn descriptive_statistics_number(_lane: usize, index: usize, _row: usize) -> f64 {
    1.0 + (index % 17) as f64
}

fn descriptive_statistics_cell(lane: usize, index: usize, _row: usize) -> CellRead<'static> {
    if lane == 4 && index == 0 {
        return CellRead::Error(litchi_ods::codec::formula::evaluation::ScalarError::NotAvailable);
    }
    CellRead::Number(descriptive_statistics_number(lane, index, 0))
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

fn aggregate_source(operation: AggregateOp, argument: &str) -> String {
    match operation {
        AggregateOp::Sum => format!("=SUM({argument})"),
    }
}

fn scalar_source(operation: Operation) -> String {
    match operation {
        Operation::Control(ControlOp::Arithmetic) => "=0.17+1.25".to_owned(),
        Operation::Control(ControlOp::Sin) => "=SIN(0.17)".to_owned(),
        Operation::Control(ControlOp::ImSum) => {
            "=IMREAL(IMSUM(COMPLEX(2;3);COMPLEX(1;4)))".to_owned()
        },
        Operation::Control(ControlOp::Average) => "=AVERAGE(1.25;2.5;3.75)".to_owned(),
        Operation::Control(ControlOp::CountA) => "=COUNTA(1.25;\"x\";TRUE())".to_owned(),
        Operation::Control(ControlOp::Var) => "=VAR(1.25;2.5;3.75)".to_owned(),
        Operation::Control(ControlOp::Stdev) => "=STDEV(1.25;2.5;3.75)".to_owned(),
        Operation::Control(ControlOp::DSum) => "=DSUM([.A1:.B4];2;[.D1:.D2])".to_owned(),
        Operation::Control(ControlOp::DVar) => "=DVAR([.A1:.B4];2;[.D1:.D2])".to_owned(),
        Operation::Control(ControlOp::DStdev) => "=DSTDEV([.A1:.B4];2;[.D1:.D2])".to_owned(),
        Operation::Aggregate(operation) => aggregate_source(operation, "1.25"),
        Operation::Representative(RepresentativeOp::Median) => "=MEDIAN(1;2;4;8)".to_owned(),
        Operation::Representative(RepresentativeOp::Rank) => "=RANK(8;8;1)".to_owned(),
        Operation::Representative(RepresentativeOp::PercentRank) => {
            "=PERCENTRANK(8;8;3)".to_owned()
        },
        Operation::Statistical(operation) => format!("={}(1.25)", operation.name()),
        Operation::Descriptive(operation) => descriptive_scalar_source(operation),
        Operation::Paired(operation) => paired_scalar_source(operation),
        Operation::Conditional(_) | Operation::NestedPaired(_) => {
            panic!("operation has no scalar profile entry")
        },
    }
}

fn descriptive_scalar_source(operation: DescriptiveOp) -> String {
    format!("={}({})", operation.name(), "1;2;4;8")
}

fn descriptive_reference_source(operation: DescriptiveOp, rows: usize) -> String {
    format!("={}([.A1:.D{rows}])", operation.name())
}

fn descriptive_sensitive_source(operation: DescriptiveOp) -> String {
    format!("={}({{1024;1025;2048;2049}})", operation.name())
}

fn descriptive_extreme_source(operation: DescriptiveOp) -> String {
    format!("={}({{-1048576;-1;1;1048576}})", operation.name())
}

fn control_array_source(operation: ControlOp, argument: &str) -> String {
    match operation {
        ControlOp::Arithmetic => format!("={argument}+1.25"),
        ControlOp::Sin => format!("=SIN({argument})"),
        ControlOp::ImSum
        | ControlOp::Average
        | ControlOp::CountA
        | ControlOp::Var
        | ControlOp::Stdev
        | ControlOp::DSum
        | ControlOp::DVar
        | ControlOp::DStdev => unreachable!("scalar-only control"),
    }
}

fn descriptive_result(operation: DescriptiveOp, values: &[f64]) -> f64 {
    let n = values.len() as f64;
    let mean = values.iter().sum::<f64>() / n;
    match operation {
        DescriptiveOp::Avedev => values.iter().map(|value| (value - mean).abs()).sum::<f64>() / n,
        DescriptiveOp::Devsq => values.iter().map(|value| (value - mean).powi(2)).sum(),
        DescriptiveOp::Kurt => {
            let variance = values
                .iter()
                .map(|value| (value - mean).powi(2))
                .sum::<f64>()
                / (n - 1.0);
            let standard = variance.sqrt();
            let standardized = values
                .iter()
                .map(|value| ((value - mean) / standard).powi(4))
                .sum::<f64>();
            n * (n + 1.0) / ((n - 1.0) * (n - 2.0) * (n - 3.0)) * standardized
                - 3.0 * (n - 1.0).powi(2) / ((n - 2.0) * (n - 3.0))
        },
        DescriptiveOp::Skew => {
            let variance = values
                .iter()
                .map(|value| (value - mean).powi(2))
                .sum::<f64>()
                / (n - 1.0);
            let standard = variance.sqrt();
            let standardized = values
                .iter()
                .map(|value| ((value - mean) / standard).powi(3))
                .sum::<f64>();
            n / ((n - 1.0) * (n - 2.0)) * standardized
        },
        DescriptiveOp::Skewp => {
            let variance = values
                .iter()
                .map(|value| (value - mean).powi(2))
                .sum::<f64>()
                / n;
            let standard = variance.sqrt();
            values
                .iter()
                .map(|value| ((value - mean) / standard).powi(3))
                .sum::<f64>()
                / n
        },
    }
}

fn descriptive_expected(
    operation: DescriptiveOp,
    lane: usize,
    start_row: usize,
    rows: usize,
    columns: usize,
) -> f64 {
    let mut values = Vec::with_capacity(rows.saturating_mul(columns));
    for row in start_row..start_row.saturating_add(rows) {
        for column in 0..columns {
            values.push(descriptive_statistics_number(
                lane,
                row.saturating_mul(4).saturating_add(column),
                row,
            ));
        }
    }
    descriptive_result(operation, &values)
}

fn paired_scalar_source(operation: PairedOp) -> String {
    match operation {
        PairedOp::Forecast => "=FORECAST(1;{3;5;7;9};{1;2;3;4})".to_owned(),
        _ => format!("={}({{3;5;7;9}};{{1;2;3;4}})", operation.name()),
    }
}

fn paired_reference_source(operation: PairedOp, rows: usize) -> String {
    let end = rows;
    match operation {
        PairedOp::Forecast => format!("=FORECAST(1;[.A1:.D{end}];[.E1:.H{end}])"),
        _ => format!("={}([.A1:.D{end}];[.E1:.H{end}])", operation.name()),
    }
}

fn paired_offset_source(operation: PairedOp) -> String {
    match operation {
        PairedOp::Forecast => "=FORECAST(1;[.A3:.D18];[.E3:.H18])".to_owned(),
        _ => format!("={}([.A3:.D18];[.E3:.H18])", operation.name()),
    }
}

fn paired_skip_source(operation: PairedOp) -> String {
    match operation {
        PairedOp::Forecast => "=FORECAST(1;[.M1:.P16];[.Q1:.T16])".to_owned(),
        _ => format!("={}([.M1:.P16];[.Q1:.T16])", operation.name()),
    }
}

fn paired_extreme_source(operation: PairedOp) -> String {
    match operation {
        PairedOp::Forecast => "=FORECAST(1;{3;5;9;17};{1;2;4;8})".to_owned(),
        _ => format!("={}({{3;5;9;17}};{{1;2;4;8}})", operation.name()),
    }
}

fn paired_shape_source(operation: PairedOp) -> String {
    if operation == PairedOp::Rsq {
        format!("={}([.A1:.D16];[.E1:.H15])", operation.name())
    } else {
        format!("={}([.A1:.D16]~[.E1:.H16];[.A1:.D16])", operation.name())
    }
}

fn paired_query_array_source(cached: bool) -> String {
    if cached {
        "=IF({TRUE()|TRUE()};FORECAST({1|2};[.A1:.D16];[.E1:.H16]);0)".to_owned()
    } else {
        "=FORECAST({1|2};[.A1:.D16];[.E1:.H16])".to_owned()
    }
}

fn paired_sequence(lane: usize, start_row: usize, rows: usize, columns: usize) -> Vec<Option<f64>> {
    let mut values = Vec::with_capacity(rows.saturating_mul(columns));
    for row in start_row..start_row.saturating_add(rows) {
        for column in 0..columns {
            let index = row.saturating_mul(4).saturating_add(column);
            if lane == 3 && index % 4 == 1 || lane == 4 && index % 4 == 2 {
                values.push(None);
            } else {
                values.push(Some(paired_statistics_number(lane, index, row)));
            }
        }
    }
    values
}

fn paired_result(operation: PairedOp, x: &[f64], y: &[f64], query: Option<f64>) -> f64 {
    let n = x.len() as f64;
    let x_mean = x.iter().sum::<f64>() / n;
    let y_mean = y.iter().sum::<f64>() / n;
    let sxx = x.iter().map(|value| (value - x_mean).powi(2)).sum::<f64>();
    let syy = y.iter().map(|value| (value - y_mean).powi(2)).sum::<f64>();
    let sxy = x
        .iter()
        .zip(y)
        .map(|(x, y)| (x - x_mean) * (y - y_mean))
        .sum::<f64>();
    let slope = sxy / sxx;
    match operation {
        PairedOp::Correl | PairedOp::Pearson => sxy / (sxx * syy).sqrt(),
        PairedOp::Covar => sxy / n,
        PairedOp::Rsq => (sxy / (sxx * syy).sqrt()).powi(2),
        PairedOp::Slope => slope,
        PairedOp::Intercept => y_mean - slope * x_mean,
        PairedOp::Forecast => y_mean + slope * (query.unwrap_or(1.0) - x_mean),
        PairedOp::Steyx => {
            let residual = syy - sxy * sxy / sxx;
            (residual / (n - 2.0)).max(0.0).sqrt()
        },
    }
}

fn paired_expected(
    operation: PairedOp,
    lane_x: usize,
    lane_y: usize,
    start_row: usize,
    rows: usize,
    columns: usize,
    query: Option<f64>,
) -> f64 {
    let left = paired_sequence(lane_x, start_row, rows, columns);
    let right = paired_sequence(lane_y, start_row, rows, columns);
    let mut left_values = Vec::new();
    let mut right_values = Vec::new();
    for (left, right) in left.into_iter().zip(right) {
        if let (Some(left), Some(right)) = (left, right) {
            left_values.push(left);
            right_values.push(right);
        }
    }
    let (x, y) = match operation {
        PairedOp::Slope | PairedOp::Intercept | PairedOp::Steyx | PairedOp::Forecast => {
            (&right_values, &left_values)
        },
        PairedOp::Correl | PairedOp::Covar | PairedOp::Pearson | PairedOp::Rsq => {
            (&left_values, &right_values)
        },
    };
    paired_result(operation, x, y, query)
}

fn paired_extreme_expected(operation: PairedOp, query: Option<f64>) -> f64 {
    let x = [1.0, 2.0, 4.0, 8.0];
    let y = [3.0, 5.0, 9.0, 17.0];
    paired_result(operation, &x, &y, query)
}

fn paired_query_expected(query: f64) -> f64 {
    paired_extreme_expected(PairedOp::Forecast, Some(query))
}

fn statistical_number(index: usize, plane: usize) -> f64 {
    (index % 17) as f64 - 8.0 + plane as f64 * 0.25
}

fn statistical_expected(
    operation: StatisticalOp,
    lane: usize,
    start_row: usize,
    rows: usize,
    columns: usize,
    plane: usize,
) -> f64 {
    let mut numbers = Vec::new();
    let mut count_nonblank = 0.0;
    for row in start_row..start_row.saturating_add(rows) {
        for column in 0..columns {
            let index = row.saturating_mul(4).saturating_add(column);
            match lane {
                0 => numbers.push(statistical_number(index, plane)),
                1 => match index % 4 {
                    0 => numbers.push((index % 11) as f64 - 4.0),
                    1 | 2 => {},
                    _ => {},
                },
                _ => {},
            }
            if lane != 2 && !(lane == 1 && index % 4 == 3) {
                count_nonblank += 1.0;
            }
        }
    }
    match operation {
        StatisticalOp::CountA => count_nonblank,
        StatisticalOp::Average => {
            if numbers.is_empty() {
                f64::NAN
            } else {
                numbers.iter().sum::<f64>() / numbers.len() as f64
            }
        },
    }
}

fn conditional_value(index: usize) -> f64 {
    0.5 + index as f64 * 0.03125
}

fn conditional_selected(index: usize, second: bool) -> bool {
    index % 4 >= 2 && (!second || index % 2 == 1)
}

fn conditional_sum(
    rows: usize,
    columns: usize,
    lane: usize,
    _threshold: usize,
    second: bool,
) -> f64 {
    let mut result = 0.0;
    for index in 0..rows.saturating_mul(columns) {
        if conditional_selected(index, second) {
            result += if lane == 0 {
                (index % 4) as f64
            } else {
                conditional_value(index)
            };
        }
    }
    result
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
            expectation: Expectation::Number(expected_control_scalar(operation)),
        });
    }
    for operation in [
        ControlOp::Average,
        ControlOp::CountA,
        ControlOp::Var,
        ControlOp::Stdev,
    ] {
        cases.push(Case {
            name: format!("scalar-control-{}", operation.name().to_ascii_lowercase()),
            source: scalar_source(Operation::Control(operation)),
            shape: Shape::Scalar,
            operation: Operation::Control(operation),
            expectation: Expectation::Number(expected_control_scalar(operation)),
        });
    }
    for operation in [ControlOp::DVar, ControlOp::DStdev] {
        cases.push(Case {
            name: format!("database-control-{}", operation.name().to_ascii_lowercase()),
            source: scalar_source(Operation::Control(operation)),
            shape: Shape::ReferenceArray {
                rows: 4,
                columns: 2,
            },
            operation: Operation::Control(operation),
            expectation: Expectation::Number(expected_control_scalar(operation)),
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
                expectation: Expectation::Number(expected_control_scalar(operation)),
            });
        }
    }
    cases.push(Case {
        name: "reference-array-16x4-arithmetic".to_owned(),
        source: "=[.A1:.D16]+1.25".to_owned(),
        shape: Shape::ReferenceArray {
            rows: 16,
            columns: 4,
        },
        operation: Operation::Control(ControlOp::Arithmetic),
        expectation: Expectation::Number(expected_control_scalar(ControlOp::Arithmetic)),
    });
    for (name, source, shape, expectation) in [
        (
            "scalar-aggregate-sum",
            "=SUM(1.25)",
            Shape::Scalar,
            Expectation::Number(1.25),
        ),
        (
            "literal-aggregate-4x1-sum",
            "=SUM({1|2|3|4})",
            Shape::LiteralArray {
                rows: 4,
                columns: 1,
            },
            Expectation::Number(10.0),
        ),
        (
            "reference-aggregate-64x4-sum",
            "=SUM([.A1:.D64])",
            Shape::ReferenceArray {
                rows: 64,
                columns: 4,
            },
            Expectation::Number(expected_reference_sum(64, 4)),
        ),
    ] {
        cases.push(Case {
            name: name.to_owned(),
            source: source.to_owned(),
            shape,
            operation: Operation::Aggregate(AggregateOp::Sum),
            expectation,
        });
    }
    cases.push(Case {
        name: "reference-conditional-256x4-sumifs".to_owned(),
        source: "=SUMIFS([.I1:.L256];[.A1:.D256];\">=2\";[.E1:.H256];1)".to_owned(),
        shape: Shape::ReferenceArray {
            rows: 256,
            columns: 4,
        },
        operation: Operation::Conditional(ConditionalOp::SumIfs),
        expectation: Expectation::Number(conditional_sum(256, 4, 2, 2, true)),
    });
    for (name, source, expectation, operation) in [
        (
            "reference-control-average",
            "=AVERAGE([.A1:.D64])",
            statistical_expected(StatisticalOp::Average, 0, 0, 64, 4, 0),
            StatisticalOp::Average,
        ),
        (
            "reference-control-counta",
            "=COUNTA([.E1:.H64])",
            statistical_expected(StatisticalOp::CountA, 1, 0, 64, 4, 0),
            StatisticalOp::CountA,
        ),
    ] {
        cases.push(Case {
            name: name.to_owned(),
            source: source.to_owned(),
            shape: Shape::ReferenceArray {
                rows: 64,
                columns: 4,
            },
            operation: Operation::Statistical(operation),
            expectation: Expectation::Number(expectation),
        });
    }
    for (name, source, expected) in [
        ("representative-median", "=MEDIAN(1;2;4;8)", 3.0),
        ("representative-rank", "=RANK(8;8;1)", 1.0),
        ("representative-percentrank", "=PERCENTRANK(8;8;3)", 1.0),
    ] {
        let operation = match name {
            "representative-median" => RepresentativeOp::Median,
            "representative-rank" => RepresentativeOp::Rank,
            _ => RepresentativeOp::PercentRank,
        };
        cases.push(Case {
            name: name.to_owned(),
            source: source.to_owned(),
            shape: Shape::Scalar,
            operation: Operation::Representative(operation),
            expectation: Expectation::Number(expected),
        });
    }

    // Existing descriptive reducers are matched in this profile because the
    // paired implementation shares their dyadic centered-moment machinery.
    for operation in DescriptiveOp::all().iter().copied() {
        let lower = operation.name().to_ascii_lowercase();
        let normal = descriptive_expected(operation, 0, 0, 64, 4);
        cases.push(Case {
            name: format!("matched-descriptive-reference-{lower}"),
            source: descriptive_reference_source(operation, 64),
            shape: Shape::ReferenceArray {
                rows: 64,
                columns: 4,
            },
            operation: Operation::Descriptive(operation),
            expectation: Expectation::Number(normal),
        });
        let sensitive = [1024.0, 1025.0, 2048.0, 2049.0];
        cases.push(Case {
            name: format!("matched-descriptive-sensitive-inline-{lower}"),
            source: descriptive_sensitive_source(operation),
            shape: Shape::LiteralArray {
                rows: 4,
                columns: 1,
            },
            operation: Operation::Descriptive(operation),
            expectation: Expectation::Number(descriptive_result(operation, &sensitive)),
        });
        let extreme = [-1_048_576.0, -1.0, 1.0, 1_048_576.0];
        cases.push(Case {
            name: format!("matched-descriptive-extreme-inline-{lower}"),
            source: descriptive_extreme_source(operation),
            shape: Shape::LiteralArray {
                rows: 4,
                columns: 1,
            },
            operation: Operation::Descriptive(operation),
            expectation: Expectation::Number(descriptive_result(operation, &extreme)),
        });
    }

    for operation in PairedOp::all().iter().copied() {
        let lower = operation.name().to_ascii_lowercase();
        for (lane, rows, source, shape) in [
            (
                "small",
                16,
                paired_reference_source(operation, 16),
                Shape::ReferenceArray {
                    rows: 16,
                    columns: 4,
                },
            ),
            (
                "large",
                256,
                paired_reference_source(operation, 256),
                Shape::ReferenceArray {
                    rows: 256,
                    columns: 4,
                },
            ),
            (
                "extreme",
                4,
                paired_extreme_source(operation),
                Shape::LiteralArray {
                    rows: 4,
                    columns: 1,
                },
            ),
            (
                "offset",
                16,
                paired_offset_source(operation),
                Shape::ReferenceArray {
                    rows: 16,
                    columns: 4,
                },
            ),
            (
                "pairwise-skip",
                16,
                paired_skip_source(operation),
                Shape::ReferenceArray {
                    rows: 16,
                    columns: 4,
                },
            ),
        ] {
            let query = if operation == PairedOp::Forecast {
                Some(1.0)
            } else {
                None
            };
            let expected = if lane == "extreme" {
                paired_extreme_expected(operation, query)
            } else {
                paired_expected(
                    operation,
                    if lane == "pairwise-skip" { 3 } else { 0 },
                    if lane == "pairwise-skip" { 4 } else { 1 },
                    if lane == "offset" { 2 } else { 0 },
                    rows,
                    4,
                    query,
                )
            };
            cases.push(Case {
                name: format!("{lane}-paired-{lower}"),
                source,
                shape,
                operation: Operation::Paired(operation),
                expectation: Expectation::Number(expected),
            });
        }
        let shape_error = if operation == PairedOp::Rsq {
            litchi_ods::codec::formula::evaluation::ScalarError::NotAvailable
        } else {
            litchi_ods::codec::formula::evaluation::ScalarError::Value
        };
        cases.push(Case {
            name: format!("shape-reject-paired-{lower}"),
            source: paired_shape_source(operation),
            shape: Shape::ReferenceArray {
                rows: 16,
                columns: 4,
            },
            operation: Operation::Paired(operation),
            expectation: Expectation::Error(shape_error),
        });
        cases.push(Case {
            name: format!("cancellation-paired-{lower}"),
            source: paired_reference_source(operation, 64),
            shape: Shape::ReferenceArray {
                rows: 64,
                columns: 4,
            },
            operation: Operation::Paired(operation),
            expectation: Expectation::Failure(FailureExpectation::Cancelled),
        });
        cases.push(Case {
            name: format!("resource-paired-{lower}"),
            source: paired_reference_source(operation, 64),
            shape: Shape::ReferenceArray {
                rows: 64,
                columns: 4,
            },
            operation: Operation::Paired(operation),
            expectation: Expectation::Failure(FailureExpectation::ReferenceCells),
        });
    }
    cases.push(Case {
        name: "forecast-query-array".to_owned(),
        source: paired_query_array_source(false),
        shape: Shape::LiteralArray {
            rows: 2,
            columns: 1,
        },
        operation: Operation::Paired(PairedOp::Forecast),
        expectation: Expectation::Number(0.0),
    });
    cases.push(Case {
        name: "forecast-query-array-cache".to_owned(),
        source: paired_query_array_source(true),
        shape: Shape::LiteralArray {
            rows: 2,
            columns: 1,
        },
        operation: Operation::NestedPaired(PairedOp::Forecast),
        expectation: Expectation::Number(0.0),
    });
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

fn execution() -> (ExecutionContext, CancellationSource) {
    let budget = Budget::root(
        "ods-formula-paired-statistics-performance",
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

fn approximately_equal(actual: f64, expected: f64) -> bool {
    (actual - expected).abs() <= expected.abs().max(1.0) * 1.0e-10
}

fn expected_reference_sum(rows: usize, columns: usize) -> f64 {
    (0..rows.saturating_mul(columns))
        .map(|index| aggregate_value(index, 0))
        .sum()
}

fn expected_control_scalar(operation: ControlOp) -> f64 {
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
    }
}

fn expected_control_cell(operation: ControlOp, index: usize) -> f64 {
    match operation {
        ControlOp::Arithmetic => fixture_value(index) + 1.25,
        ControlOp::Sin => fixture_value(index).sin(),
        ControlOp::ImSum => unreachable!("array-only control oracle"),
        ControlOp::Average | ControlOp::CountA | ControlOp::Var | ControlOp::Stdev => {
            let _ = index;
            expected_control_scalar(operation)
        },
        ControlOp::DSum => 12.0,
        ControlOp::DVar | ControlOp::DStdev => {
            let _ = index;
            expected_control_scalar(operation)
        },
    }
}

fn value_limits(case: &Case) -> Limits {
    match case.expectation {
        Expectation::Failure(FailureExpectation::ReferenceCells) => {
            Limits::default().with_max_reference_cells(0)
        },
        Expectation::Number(_)
        | Expectation::Error(_)
        | Expectation::Failure(FailureExpectation::Cancelled) => Limits::default(),
    }
}

fn validate_failure(case: &Case, error: &EvaluationFailure) -> AnyResult<()> {
    match (case.expectation, error) {
        (
            Expectation::Failure(FailureExpectation::ReferenceCells),
            EvaluationFailure::ResourceLimit(limit),
        ) if limit.resource == Resource::Objects => Ok(()),
        (Expectation::Failure(FailureExpectation::Cancelled), EvaluationFailure::Cancelled) => {
            Ok(())
        },
        (Expectation::Failure(expected), observed) => Err(format!(
            "{} returned {observed}, expected evaluator failure {expected:?}",
            case.name
        )
        .into()),
        (_, observed) => {
            Err(format!("{} returned unexpected failure {observed}", case.name).into())
        },
    }
}

fn failure_checksum(expectation: Expectation) -> u64 {
    match expectation {
        Expectation::Failure(FailureExpectation::ReferenceCells) => 0xf1f1_f1f1_f1f1_f1f1,
        Expectation::Failure(FailureExpectation::Cancelled) => 0xcaca_caca_caca_caca,
        Expectation::Number(_) | Expectation::Error(_) => 0,
    }
}

fn scalar_checksum(value: &ScalarValue<'_>) -> AnyResult<u64> {
    match value {
        ScalarValue::Number(value) => Ok(value.to_bits().rotate_left(11)),
        ScalarValue::Error(error) => Ok((*error as u64).rotate_left(11)),
        other => Err(format!("expected scalar Number, got {other:?}").into()),
    }
}

fn value_checksum(value: Value<'_>) -> AnyResult<u64> {
    match value {
        Value::Number(value) => Ok(value.to_bits().rotate_left(11)),
        Value::Error(error) => Ok((error as u64).rotate_left(11)),
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
    match (case.expectation, value) {
        (Expectation::Number(expected), ScalarValue::Number(actual)) => {
            if approximately_equal(*actual, expected) {
                Ok(())
            } else {
                Err(format!("{} returned {actual}, expected {expected}", case.name).into())
            }
        },
        (Expectation::Error(expected), ScalarValue::Error(actual)) if *actual == expected => Ok(()),
        (Expectation::Error(expected), observed) => Err(format!(
            "{} returned {observed:?}, expected scalar error {expected}",
            case.name
        )
        .into()),
        (Expectation::Number(expected), observed) => Err(format!(
            "{} returned {observed:?}, expected scalar number {expected}",
            case.name
        )
        .into()),
        (Expectation::Failure(expected), observed) => Err(format!(
            "{} returned {observed:?}, expected evaluator failure {expected:?}",
            case.name
        )
        .into()),
    }
}

fn validate_value(case: &Case, result: &Evaluated<'_>) -> AnyResult<()> {
    let is_forecast_array = matches!(
        case.operation,
        Operation::Paired(PairedOp::Forecast) | Operation::NestedPaired(PairedOp::Forecast)
    ) && case.name.starts_with("forecast-query-array");
    if is_forecast_array {
        let array = result
            .as_array()
            .ok_or("FORECAST query result was not an array")?;
        let (rows, columns) = case.shape.dimensions().ok_or("missing query dimensions")?;
        if (array.shape().rows(), array.shape().columns()) != (rows, columns) {
            return Err(format!("{} returned wrong query shape", case.name).into());
        }
        for (index, expected) in [paired_query_expected(1.0), paired_query_expected(2.0)]
            .into_iter()
            .enumerate()
        {
            let actual = match array.get(index).ok_or("missing query result")? {
                Value::Number(value) => value,
                value => return Err(format!("{} query cell {index}: {value:?}", case.name).into()),
            };
            if !approximately_equal(actual, expected) {
                return Err(format!(
                    "{} query cell {index} returned {actual}, expected {expected}",
                    case.name
                )
                .into());
            }
        }
        return Ok(());
    }
    if matches!(
        case.operation,
        Operation::Descriptive(_) | Operation::Paired(_) | Operation::NestedPaired(_)
    ) {
        match case.expectation {
            Expectation::Number(expected) => {
                let actual = match result.value() {
                    Value::Number(value) => value,
                    value => return Err(format!("{} returned {value:?}", case.name).into()),
                };
                if !approximately_equal(actual, expected) {
                    return Err(
                        format!("{} returned {actual}, expected {expected}", case.name).into(),
                    );
                }
            },
            Expectation::Error(expected) => {
                if !matches!(result.value(), Value::Error(actual) if actual == expected) {
                    return Err(format!(
                        "{} returned {:?}, expected {expected}",
                        case.name,
                        result.value()
                    )
                    .into());
                }
            },
            Expectation::Failure(expected) => {
                return Err(format!(
                    "{} returned {:?}, expected typed failure {expected:?}",
                    case.name,
                    result.value()
                )
                .into());
            },
        }
        return Ok(());
    }
    if case.operation.is_aggregate() {
        match case.expectation {
            Expectation::Number(expected) => {
                let actual = match result.value() {
                    Value::Number(value) => value,
                    value => return Err(format!("{} returned {value:?}", case.name).into()),
                };
                if !approximately_equal(actual, expected) {
                    return Err(
                        format!("{} returned {actual}, expected {expected}", case.name).into(),
                    );
                }
            },
            Expectation::Error(expected) => {
                if !matches!(result.value(), Value::Error(actual) if actual == expected) {
                    return Err(format!(
                        "{} returned {:?}, expected {expected}",
                        case.name,
                        result.value()
                    )
                    .into());
                }
            },
            Expectation::Failure(expected) => {
                return Err(format!(
                    "{} returned {:?}, expected evaluator failure {expected:?}",
                    case.name,
                    result.value()
                )
                .into());
            },
        }
        return Ok(());
    }
    if matches!(
        case.operation,
        Operation::Control(ControlOp::DSum | ControlOp::DVar | ControlOp::DStdev)
    ) {
        let actual = match result.value() {
            Value::Number(value) => value,
            value => return Err(format!("{} returned {value:?}", case.name).into()),
        };
        let expected = match case.operation {
            Operation::Control(operation) => expected_control_scalar(operation),
            _ => unreachable!(),
        };
        if !approximately_equal(actual, expected) {
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
    let limits = value_limits(case);
    value::evaluate(expression, resolver, &context, &limits)
}

fn preflight(case: &Case) -> AnyResult<u64> {
    let expression = Expression::parse(&case.source).map_err(|error| {
        format!(
            "{} parse preflight failed for {:?}: {error}",
            case.name, case.source
        )
    })?;
    let (execution, cancellation) = execution();
    if case.shape.is_reference() || !case.shape.is_scalar_input() {
        let resolver = FixtureResolver::new(
            matches!(
                case.operation,
                Operation::Control(ControlOp::DSum | ControlOp::DVar | ControlOp::DStdev)
            ),
            matches!(case.operation, Operation::Aggregate(_)),
            case.operation.is_conditional(),
            case.operation.is_statistical(),
            matches!(case.operation, Operation::Descriptive(_)),
            matches!(
                case.operation,
                Operation::Paired(_) | Operation::NestedPaired(_)
            ),
            matches!(
                case.expectation,
                Expectation::Failure(FailureExpectation::Cancelled)
            )
            .then_some(cancellation.clone()),
        );
        match eval_value(case, &expression, &resolver, &execution) {
            Ok(result) => {
                if matches!(case.expectation, Expectation::Failure(_)) {
                    return Err(format!("{} unexpectedly produced a value", case.name).into());
                }
                validate_value(case, &result)?;
            },
            Err(error) if matches!(case.expectation, Expectation::Failure(_)) => {
                validate_failure(case, &error)?;
            },
            Err(error) => {
                return Err(format!("{} value preflight failed: {error}", case.name).into());
            },
        }
        Ok(resolver.stats.reads())
    } else {
        let result = eval_scalar(&expression, &execution)
            .map_err(|error| format!("{} scalar preflight failed: {error}", case.name))?;
        validate_scalar(case, result.value())?;
        Ok(0)
    }
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
    preflight_only: bool,
}

fn parse_config() -> AnyResult<Option<Config>> {
    let mut case = None;
    let mut phase = Phase::Evaluate;
    let mut warmups = 3;
    let mut iterations = 1;
    let mut repeat = None;
    let mut preflight_only = false;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                println!(
                    "usage: ods-formula-paired-statistics-performance [--case NAME|all] [--phase evaluate|parse-evaluate] [--warmups N] [--iterations N] [--repeat N] [--preflight-only]"
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
            "--preflight-only" => preflight_only = true,
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
) -> AnyResult<u64> {
    let result = result.map_err(|error| format!("scalar evaluation failed: {error}"))?;
    validate_scalar(case, result.value())?;
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
    let result = match result {
        Ok(result) => result,
        Err(error) => {
            validate_failure(case, &error)?;
            *checksum = checksum.wrapping_add(failure_checksum(case.expectation));
            black_box(*checksum);
            return Ok(execution.budget().used(Resource::Memory));
        },
    };
    if matches!(case.expectation, Expectation::Failure(_)) {
        return Err(format!("{} unexpectedly produced a value", case.name).into());
    }
    let is_forecast_array = matches!(
        case.operation,
        Operation::Paired(PairedOp::Forecast) | Operation::NestedPaired(PairedOp::Forecast)
    ) && case.name.starts_with("forecast-query-array");
    if (case.operation.is_aggregate() && !is_forecast_array)
        || case.shape.is_scalar_input()
        || matches!(
            case.operation,
            Operation::Control(ControlOp::DSum | ControlOp::DVar | ControlOp::DStdev)
        )
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
    let (execution, cancellation) = execution();
    let resolver = FixtureResolver::new(
        matches!(
            case.operation,
            Operation::Control(ControlOp::DSum | ControlOp::DVar | ControlOp::DStdev)
        ),
        matches!(case.operation, Operation::Aggregate(_)),
        case.operation.is_conditional(),
        case.operation.is_statistical(),
        matches!(case.operation, Operation::Descriptive(_)),
        matches!(
            case.operation,
            Operation::Paired(_) | Operation::NestedPaired(_)
        ),
        matches!(
            case.expectation,
            Expectation::Failure(FailureExpectation::Cancelled)
        )
        .then_some(cancellation.clone()),
    );
    resolver.stats.reset();
    let context =
        Context::new(&execution, Position::new("Main", 0, 0)).with_mode(case.shape.mode());
    let scalar_context = EvaluationContext::new(&execution);
    let scalar_limits = EvaluationLimits::default();
    let value_limits = value_limits(case);
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut checksum = 0_u64;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        retained_peak = if case.shape.is_scalar_input() && !case.shape.is_reference() {
            retained_peak.max(consume_scalar(
                case,
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
    let (execution, cancellation) = execution();
    let resolver = FixtureResolver::new(
        matches!(
            case.operation,
            Operation::Control(ControlOp::DSum | ControlOp::DVar | ControlOp::DStdev)
        ),
        matches!(case.operation, Operation::Aggregate(_)),
        case.operation.is_conditional(),
        case.operation.is_statistical(),
        matches!(case.operation, Operation::Descriptive(_)),
        matches!(
            case.operation,
            Operation::Paired(_) | Operation::NestedPaired(_)
        ),
        matches!(
            case.expectation,
            Expectation::Failure(FailureExpectation::Cancelled)
        )
        .then_some(cancellation.clone()),
    );
    let context =
        Context::new(&execution, Position::new("Main", 0, 0)).with_mode(case.shape.mode());
    let scalar_context = EvaluationContext::new(&execution);
    let scalar_limits = EvaluationLimits::default();
    let value_limits = value_limits(case);
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
                case,
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

fn expectation_label(expectation: Expectation) -> String {
    match expectation {
        Expectation::Number(_) => "finite-number".to_owned(),
        Expectation::Error(error) => format!("error:{error}"),
        Expectation::Failure(FailureExpectation::ReferenceCells) => {
            "failure:reference-cells".to_owned()
        },
        Expectation::Failure(FailureExpectation::Cancelled) => "failure:cancelled".to_owned(),
    }
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
        "{{\"case\":\"{}\",\"operation\":\"{}\",\"phase\":\"{}\",\"input_bytes\":{},\"repeat\":{},\"warmups\":{},\"iterations\":{},\"shape\":\"{}\",\"rows\":{},\"columns\":{},\"elements\":{},\"supported\":true,\"expected\":\"{}\",\"elapsed_ns_p50\":{},\"elapsed_ns_mean\":{},\"elapsed_ns_p95\":{},\"elapsed_ns_p99\":{},\"elapsed_ns_per_repeat\":{},\"allocator_calls_p50\":{},\"requested_bytes_p50\":{},\"released_bytes_p50\":{},\"peak_live_delta_p50\":{},\"memory_retained_p50\":{},\"work_p50\":{},\"work_per_repeat\":{},\"reference_reads_p50\":{},\"reference_reads_per_repeat\":{},\"checksum_p50\":{},\"live_before_p50\":{},\"live_after_p50\":{},\"samples\":{},\"rss_kib\":null,\"rss_source\":\"external /usr/bin/time -v\",\"memory_retained_source\":\"execution budget while result is live\",\"reference_source\":\"instrumented borrowing resolver\",\"validation_scope\":\"one untimed direct f64 fixture oracle; timed evaluator/checksum/drop\"}}",
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
        json_escape(&expectation_label(case.expectation)),
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
    let preflight_reads = preflight(case)?;
    if config.preflight_only {
        println!(
            "preflight case={} reference_reads={preflight_reads}",
            case.name
        );
        return Ok(());
    }
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
    let selected = selected_cases(&all, config.case.as_deref())?;
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
