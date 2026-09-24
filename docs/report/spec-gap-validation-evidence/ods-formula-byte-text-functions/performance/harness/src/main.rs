//! Bounded process-level profile for the ODS byte-position text evaluator.
//!
//! The fixture is deterministic and entirely in memory.  Every timed child
//! performs a correctness preflight before measuring one immutable expression
//! repeatedly. Reference cases use a borrowing resolver and count cell reads
//! so a profile row can distinguish streaming descriptors from literal arrays.

// The fixture primitives for the earlier aggregate profiles are retained for
// matched-control compatibility while this package keeps its case matrix
// bounded to text workloads. They are intentionally not emitted by
// `all_cases` unless a future control is explicitly added.
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
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarValue, evaluate_scalar,
    },
    expression::Expression,
};

type AnyResult<T> = Result<T, Box<dyn Error>>;

// The six-position fraction lane intentionally admits the formatter's
// bounded 999,999-denominator search. Its caller limit is raised only for
// that lane so the profile measures the maximum admitted kernel; every other
// scalar case retains EvaluationLimits::default().
const FRACTION_PROFILE_MAX_STEPS: u64 = 2_000_000;

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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TextOp {
    Asc,
    Char,
    Clean,
    Code,
    Concatenate,
    Dollar,
    Exact,
    Find,
    Fixed,
    Jis,
    Left,
    Len,
    Lower,
    Mid,
    Proper,
    Replace,
    Rept,
    Right,
    Search,
    Substitute,
    T,
    Text,
    Trim,
    Unichar,
    Unicode,
    Upper,
}

impl TextOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Asc => "ASC",
            Self::Char => "CHAR",
            Self::Clean => "CLEAN",
            Self::Code => "CODE",
            Self::Concatenate => "CONCATENATE",
            Self::Dollar => "DOLLAR",
            Self::Exact => "EXACT",
            Self::Find => "FIND",
            Self::Fixed => "FIXED",
            Self::Jis => "JIS",
            Self::Left => "LEFT",
            Self::Len => "LEN",
            Self::Lower => "LOWER",
            Self::Mid => "MID",
            Self::Proper => "PROPER",
            Self::Replace => "REPLACE",
            Self::Rept => "REPT",
            Self::Right => "RIGHT",
            Self::Search => "SEARCH",
            Self::Substitute => "SUBSTITUTE",
            Self::T => "T",
            Self::Text => "TEXT",
            Self::Trim => "TRIM",
            Self::Unichar => "UNICHAR",
            Self::Unicode => "UNICODE",
            Self::Upper => "UPPER",
        }
    }

    const fn all() -> &'static [Self] {
        &[
            Self::Asc,
            Self::Char,
            Self::Clean,
            Self::Code,
            Self::Concatenate,
            Self::Dollar,
            Self::Exact,
            Self::Find,
            Self::Fixed,
            Self::Jis,
            Self::Left,
            Self::Len,
            Self::Lower,
            Self::Mid,
            Self::Proper,
            Self::Replace,
            Self::Rept,
            Self::Right,
            Self::Search,
            Self::Substitute,
            Self::T,
            Self::Text,
            Self::Trim,
            Self::Unichar,
            Self::Unicode,
            Self::Upper,
        ]
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ByteOp {
    FindB,
    LeftB,
    LenB,
    MidB,
    ReplaceB,
    RightB,
    SearchB,
}

impl ByteOp {
    const fn name(self) -> &'static str {
        match self {
            Self::FindB => "FINDB",
            Self::LeftB => "LEFTB",
            Self::LenB => "LENB",
            Self::MidB => "MIDB",
            Self::ReplaceB => "REPLACEB",
            Self::RightB => "RIGHTB",
            Self::SearchB => "SEARCHB",
        }
    }

    const fn all() -> &'static [Self] {
        &[
            Self::FindB,
            Self::LeftB,
            Self::LenB,
            Self::MidB,
            Self::ReplaceB,
            Self::RightB,
            Self::SearchB,
        ]
    }

    const fn returns_text(self) -> bool {
        matches!(
            self,
            Self::LeftB | Self::MidB | Self::ReplaceB | Self::RightB
        )
    }

    const fn is_search(self) -> bool {
        matches!(self, Self::FindB | Self::SearchB)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ByteLane {
    Tiny,
    LargeUnicode,
    LargeAscii,
    Reference64,
    MatrixBroadcast,
    Refusal,
    Cancellation,
    Resource,
    SearchWorstCase,
    ReplacementGrowth,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ProjectionOp {
    AverageLenB,
    SumLenB,
    AverageLen,
}

impl ProjectionOp {
    const fn all() -> &'static [Self] {
        &[Self::AverageLenB, Self::SumLenB, Self::AverageLen]
    }

    const fn name(self) -> &'static str {
        match self {
            Self::AverageLenB => "AVERAGE(LENB)",
            Self::SumLenB => "SUM(LENB)",
            Self::AverageLen => "AVERAGE(LEN)",
        }
    }

    const fn source(self) -> &'static str {
        match self {
            Self::AverageLenB => "=IF({TRUE()|TRUE()};AVERAGE(LENB([.A1:.A2]));0)",
            Self::SumLenB => "=IF({TRUE()|TRUE()};SUM(LENB([.A1:.A2]));0)",
            Self::AverageLen => "=IF({TRUE()|TRUE()};AVERAGE(LEN([.A1:.A2]));0)",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TextLane {
    Tiny,
    LargeUnicode,
    Reference64,
    MatrixBroadcast,
    Refusal,
    Cancellation,
    Resource,
    SearchWorstCase,
    ReptGrowth,
    AscJisExpansion,
    FormatterFraction,
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
    NumberAny,
    LogicalAny,
    Error(litchi_ods::codec::formula::evaluation::ScalarError),
    Text,
    TextExact(&'static str),
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
    Text(TextOp, TextLane),
    Byte(ByteOp, ByteLane),
    Projection(ProjectionOp),
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
            Self::Text(operation, _) => operation.name(),
            Self::Byte(operation, _) => operation.name(),
            Self::Projection(operation) => operation.name(),
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
                | Self::Projection(_)
        )
    }

    const fn is_conditional(self) -> bool {
        matches!(self, Self::Conditional(_))
    }

    const fn is_statistical(self) -> bool {
        matches!(
            self,
            Self::Statistical(_)
                | Self::Descriptive(_)
                | Self::Paired(_)
                | Self::NestedPaired(_)
                | Self::Projection(_)
        )
    }

    const fn is_text(self) -> bool {
        matches!(self, Self::Text(_, _) | Self::Byte(_, _))
    }

    const fn uses_text_inputs(self) -> bool {
        self.is_text() || matches!(self, Self::Projection(_))
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
    text: bool,
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
        text: bool,
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
            text,
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
        if self.text {
            let index = row.saturating_mul(4).saturating_add(column % 4);
            return Ok(text_cell(column / 4, index));
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

const TEXT_FIXTURE: [&str; 8] = [
    "Alpha βeta",
    "東京 ｶﾀｶﾅa",
    "wide ＡＢＣa",
    "mixed café",
    "search needle",
    "repeat xx a",
    "Trim   words a",
    "Ωmega 42",
];

fn text_cell(lane: usize, index: usize) -> CellRead<'static> {
    // A:D is borrowed UTF-8 text. E:E is a compact numeric companion used by
    // MID/REPLACE matrix broadcast cases; additional lanes stay deterministic
    // so a contract fixture can add a second descriptor without changing the
    // resolver interface.
    if lane == 0 {
        CellRead::Text(TEXT_FIXTURE[index % TEXT_FIXTURE.len()])
    } else if lane == 1 {
        CellRead::Number((index % 3 + 1) as f64)
    } else {
        CellRead::Text(TEXT_FIXTURE[(index + lane) % TEXT_FIXTURE.len()])
    }
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
        Operation::Text(operation, _) => text_source(operation, TextLane::Tiny),
        Operation::Byte(operation, _) => byte_source(operation, ByteLane::Tiny),
        Operation::Projection(operation) => operation.source().to_owned(),
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

fn quote_text(value: &str) -> String {
    format!("\"{}\"", value.replace('"', "\"\""))
}

fn large_unicode_text() -> String {
    let seed = "Alpha βeta 東京 ｶﾀｶﾅ café Ωmega ";
    seed.repeat(64)
}

fn text_argument(operation: TextOp, lane: TextLane) -> String {
    let numeric = matches!(
        operation,
        TextOp::Char | TextOp::Dollar | TextOp::Fixed | TextOp::Text | TextOp::Unichar
    );
    match lane {
        TextLane::Reference64 | TextLane::MatrixBroadcast => {
            if numeric {
                "[.E1:.H16]".to_owned()
            } else {
                "[.A1:.D16]".to_owned()
            }
        },
        TextLane::LargeUnicode => {
            if numeric {
                "65".to_owned()
            } else {
                quote_text(&large_unicode_text())
            }
        },
        _ if numeric => "65".to_owned(),
        _ => quote_text("Alpha βeta"),
    }
}

fn text_source(operation: TextOp, lane: TextLane) -> String {
    if matches!(lane, TextLane::FormatterFraction) {
        return "=TEXT(0.3333333333333333;\"######/######\")".to_owned();
    }
    if matches!(lane, TextLane::Refusal) {
        return match operation {
            TextOp::Char | TextOp::Unichar => format!("={}(-1)", operation.name()),
            TextOp::Dollar | TextOp::Fixed | TextOp::Text => {
                format!("={}({};\"0\")", operation.name(), quote_text("bad"))
            },
            TextOp::Concatenate
            | TextOp::Exact
            | TextOp::Find
            | TextOp::Left
            | TextOp::Len
            | TextOp::Lower
            | TextOp::Mid
            | TextOp::Proper
            | TextOp::Replace
            | TextOp::Rept
            | TextOp::Right
            | TextOp::Search
            | TextOp::Substitute
            | TextOp::T
            | TextOp::Trim
            | TextOp::Asc
            | TextOp::Clean
            | TextOp::Code
            | TextOp::Jis
            | TextOp::Unicode
            | TextOp::Upper => format!("={}([.A1:.D16]~[.E1:.H16])", operation.name()),
        };
    }
    if matches!(lane, TextLane::SearchWorstCase) {
        let haystack = "a".repeat(4096) + "needle";
        let needle = "a".repeat(127) + "needle";
        return match operation {
            TextOp::Find | TextOp::Search => {
                format!(
                    "={}({};{})",
                    operation.name(),
                    quote_text(&needle),
                    quote_text(&haystack)
                )
            },
            TextOp::Substitute => format!(
                "=SUBSTITUTE({};{};\"x\")",
                quote_text(&haystack),
                quote_text(&needle)
            ),
            TextOp::Exact => format!(
                "=EXACT({};{})",
                quote_text(&haystack),
                quote_text(&haystack)
            ),
            _ => format!("=LEN({})", quote_text(&haystack)),
        };
    }
    if matches!(lane, TextLane::ReptGrowth) {
        return "=REPT(\"x\";8192)".to_owned();
    }
    if matches!(lane, TextLane::AscJisExpansion) {
        return match operation {
            TextOp::Asc => "=ASC(\"ＡＢＣガ\")".to_owned(),
            TextOp::Jis => "=JIS(\"ABCｶﾞ\")".to_owned(),
            _ => format!("=LEN({})", quote_text("width")),
        };
    }
    let text = text_argument(operation, lane);
    let matrix_companion = if matches!(lane, TextLane::MatrixBroadcast) {
        "[.E1:.E16]"
    } else {
        "2"
    };
    match operation {
        TextOp::Asc
        | TextOp::Jis
        | TextOp::Clean
        | TextOp::Lower
        | TextOp::Proper
        | TextOp::Trim
        | TextOp::Upper => format!("={}( {})", operation.name(), text).replace("( ", "("),
        TextOp::Char | TextOp::Unichar => format!("={}({})", operation.name(), text),
        TextOp::Code | TextOp::Unicode => format!("={}({})", operation.name(), text),
        TextOp::Concatenate => {
            if matches!(lane, TextLane::MatrixBroadcast) {
                format!("=CONCATENATE({};{})", text, matrix_companion)
            } else {
                format!("=CONCATENATE({};\"-tail\")", text)
            }
        },
        TextOp::Dollar => format!("=DOLLAR({};2)", text),
        TextOp::Fixed => format!("=FIXED({};2;TRUE())", text),
        TextOp::Exact => {
            if matches!(lane, TextLane::MatrixBroadcast) {
                format!("=EXACT({};{})", text, matrix_companion)
            } else {
                format!("=EXACT({};{})", text, text)
            }
        },
        TextOp::Find | TextOp::Search => {
            if matches!(lane, TextLane::MatrixBroadcast) {
                format!("={}({};{})", operation.name(), quote_text(" "), text)
            } else {
                format!("={}({};{})", operation.name(), quote_text(" "), text)
            }
        },
        TextOp::Left | TextOp::Right => {
            format!("={}({};{})", operation.name(), text, matrix_companion)
        },
        TextOp::Len => format!("=LEN({})", text),
        TextOp::Mid => format!("=MID({};{};2)", text, matrix_companion),
        TextOp::Replace => format!("=REPLACE({};{};1;\"Z\")", text, matrix_companion),
        TextOp::Rept => format!("=REPT({};2)", text),
        TextOp::Substitute => format!("=SUBSTITUTE({};\"a\";\"x\")", text),
        TextOp::T => format!("=T({})", text),
        TextOp::Text => format!("=TEXT({};\"0.00\")", text),
    }
}

fn large_ascii_text() -> String {
    "The quick brown fox jumps over the lazy dog. ".repeat(256)
}

fn byte_text_argument(lane: ByteLane) -> String {
    match lane {
        ByteLane::Reference64 | ByteLane::MatrixBroadcast => "[.A1:.D16]".to_owned(),
        ByteLane::LargeUnicode => quote_text(&large_unicode_text()),
        ByteLane::LargeAscii => quote_text(&large_ascii_text()),
        ByteLane::SearchWorstCase => quote_text(&("a".repeat(4096) + "needle")),
        ByteLane::ReplacementGrowth => quote_text("Aé界🙂"),
        _ => quote_text("Aé界🙂"),
    }
}

fn byte_source(operation: ByteOp, lane: ByteLane) -> String {
    if matches!(lane, ByteLane::Refusal) {
        return format!("={}([.A1:.D16]~[.E1:.H16])", operation.name());
    }
    if matches!(lane, ByteLane::SearchWorstCase) {
        let haystack = "a".repeat(4096) + "needle";
        let needle = "a".repeat(127) + "needle";
        return format!(
            "={}({};{};1)",
            operation.name(),
            quote_text(&needle),
            quote_text(&haystack)
        );
    }
    if matches!(lane, ByteLane::ReplacementGrowth) {
        return format!(
            "=REPLACEB({};1;0;{})",
            quote_text(&large_ascii_text()),
            quote_text(&"x".repeat(4096))
        );
    }
    let text = byte_text_argument(lane);
    if matches!(lane, ByteLane::MatrixBroadcast) {
        let companion = "[.E1:.E16]";
        return match operation {
            ByteOp::FindB | ByteOp::SearchB => {
                format!("={}({};{};1)", operation.name(), quote_text("a"), text)
            },
            ByteOp::LenB => format!("=LENB({text})"),
            ByteOp::LeftB | ByteOp::RightB => {
                format!("={}({};{})", operation.name(), text, companion)
            },
            ByteOp::MidB => format!("=MIDB({};{};2)", text, companion),
            ByteOp::ReplaceB => format!("=REPLACEB({};{};2;\"Z\")", text, companion),
        };
    }
    match operation {
        ByteOp::FindB | ByteOp::SearchB => format!(
            "={}({};{};1)",
            operation.name(),
            quote_text(&match lane {
                ByteLane::LargeUnicode => "β",
                ByteLane::LargeAscii => "fox",
                ByteLane::Reference64 | ByteLane::MatrixBroadcast => "a",
                _ => "界",
            }),
            text
        ),
        ByteOp::LenB => format!("=LENB({text})"),
        ByteOp::LeftB | ByteOp::RightB => format!("={}({};7)", operation.name(), text),
        ByteOp::MidB => format!("=MIDB({text};3;5)"),
        ByteOp::ReplaceB => format!("=REPLACEB({text};3;2;\"XYZ\")"),
    }
}

fn byte_expectation(operation: ByteOp) -> Expectation {
    if operation.returns_text() {
        Expectation::Text
    } else {
        Expectation::NumberAny
    }
}

fn byte_cases() -> Vec<Case> {
    let mut cases = Vec::new();
    for operation in ByteOp::all().iter().copied() {
        let lower = operation.name().to_ascii_lowercase();
        for (label, lane, shape) in [
            ("tiny", ByteLane::Tiny, Shape::Scalar),
            ("large-unicode", ByteLane::LargeUnicode, Shape::Scalar),
            ("large-ascii", ByteLane::LargeAscii, Shape::Scalar),
            (
                "reference-64",
                ByteLane::Reference64,
                Shape::ReferenceArray {
                    rows: 16,
                    columns: 4,
                },
            ),
            (
                "refusal",
                ByteLane::Refusal,
                Shape::ReferenceArray {
                    rows: 16,
                    columns: 4,
                },
            ),
        ] {
            let expectation = if matches!(lane, ByteLane::Refusal) {
                Expectation::Error(litchi_ods::codec::formula::evaluation::ScalarError::Value)
            } else {
                byte_expectation(operation)
            };
            cases.push(Case {
                name: format!("{label}-byte-{lower}"),
                source: byte_source(operation, lane),
                shape,
                operation: Operation::Byte(operation, lane),
                expectation,
            });
        }
        cases.push(Case {
            name: format!("matrix-broadcast-byte-{lower}"),
            source: byte_source(operation, ByteLane::MatrixBroadcast),
            shape: Shape::ReferenceArray {
                rows: 16,
                columns: 4,
            },
            operation: Operation::Byte(operation, ByteLane::MatrixBroadcast),
            expectation: byte_expectation(operation),
        });
    }
    for operation in [
        ByteOp::LenB,
        ByteOp::MidB,
        ByteOp::ReplaceB,
        ByteOp::SearchB,
    ] {
        let lower = operation.name().to_ascii_lowercase();
        cases.push(Case {
            name: format!("cancellation-byte-{lower}"),
            source: byte_source(operation, ByteLane::Reference64),
            shape: Shape::ReferenceArray {
                rows: 16,
                columns: 4,
            },
            operation: Operation::Byte(operation, ByteLane::Cancellation),
            expectation: Expectation::Failure(FailureExpectation::Cancelled),
        });
        cases.push(Case {
            name: format!("resource-byte-{lower}"),
            source: byte_source(operation, ByteLane::Reference64),
            shape: Shape::ReferenceArray {
                rows: 16,
                columns: 4,
            },
            operation: Operation::Byte(operation, ByteLane::Resource),
            expectation: Expectation::Failure(FailureExpectation::ReferenceCells),
        });
    }
    for operation in [ByteOp::FindB, ByteOp::SearchB] {
        cases.push(Case {
            name: format!(
                "search-worstcase-byte-{}",
                operation.name().to_ascii_lowercase()
            ),
            source: byte_source(operation, ByteLane::SearchWorstCase),
            shape: Shape::Scalar,
            operation: Operation::Byte(operation, ByteLane::SearchWorstCase),
            expectation: Expectation::NumberAny,
        });
    }
    cases.push(Case {
        name: "replaceb-growth".to_owned(),
        source: byte_source(ByteOp::ReplaceB, ByteLane::ReplacementGrowth),
        shape: Shape::Scalar,
        operation: Operation::Byte(ByteOp::ReplaceB, ByteLane::ReplacementGrowth),
        expectation: Expectation::Text,
    });
    cases
}

fn projection_expected(operation: ProjectionOp, index: usize) -> f64 {
    let first = TEXT_FIXTURE[0];
    let second = TEXT_FIXTURE[4];
    match (operation, index) {
        (ProjectionOp::AverageLenB, 0) => first.len() as f64,
        (ProjectionOp::AverageLenB, _) => second.len() as f64,
        (ProjectionOp::SumLenB, _) => (first.len() + second.len()) as f64,
        (ProjectionOp::AverageLen, 0) => first.chars().count() as f64,
        (ProjectionOp::AverageLen, _) => second.chars().count() as f64,
    }
}

fn projection_cases() -> Vec<Case> {
    ProjectionOp::all()
        .iter()
        .copied()
        .map(|operation| Case {
            name: match operation {
                ProjectionOp::AverageLenB => "projected-statistical-average-lenb".to_owned(),
                ProjectionOp::SumLenB => "projected-statistical-sum-lenb".to_owned(),
                ProjectionOp::AverageLen => "projected-statistical-average-len".to_owned(),
            },
            source: operation.source().to_owned(),
            shape: Shape::ReferenceArray {
                rows: 2,
                columns: 1,
            },
            operation: Operation::Projection(operation),
            expectation: Expectation::NumberAny,
        })
        .collect()
}

fn text_expectation(operation: TextOp) -> Expectation {
    match operation {
        TextOp::Code | TextOp::Find | TextOp::Len | TextOp::Search | TextOp::Unicode => {
            Expectation::NumberAny
        },
        TextOp::Exact => Expectation::LogicalAny,
        _ => Expectation::Text,
    }
}

fn text_control_source(name: &str) -> String {
    match name {
        "concat-borrowed-literals" => "=\"left\"&\"right\"".to_owned(),
        "concat-owned-left" => "=(\"left\"&\"mid\")&\"right\"".to_owned(),
        "concat-owned-right" => "\"left\"&(\"mid\"&\"right\")".to_owned(),
        "concat-growth-chain" => "=\"a\"&\"b\"&\"c\"&\"d\"&\"e\"&\"f\"&\"g\"&\"h\"".to_owned(),
        _ => unreachable!("unknown text control {name}"),
    }
}

fn text_cases() -> Vec<Case> {
    let mut cases = Vec::new();
    for operation in TextOp::all().iter().copied() {
        let lower = operation.name().to_ascii_lowercase();
        for (label, lane, shape, expectation) in [
            (
                "tiny",
                TextLane::Tiny,
                Shape::Scalar,
                text_expectation(operation),
            ),
            (
                "large-unicode",
                TextLane::LargeUnicode,
                Shape::Scalar,
                text_expectation(operation),
            ),
            (
                "reference-64",
                TextLane::Reference64,
                Shape::ReferenceArray {
                    rows: 16,
                    columns: 4,
                },
                text_expectation(operation),
            ),
            (
                "refusal",
                TextLane::Refusal,
                Shape::ReferenceArray {
                    rows: 16,
                    columns: 4,
                },
                Expectation::Error(litchi_ods::codec::formula::evaluation::ScalarError::Value),
            ),
        ] {
            cases.push(Case {
                name: format!("{label}-text-{lower}"),
                source: text_source(operation, lane),
                shape,
                operation: Operation::Text(operation, lane),
                expectation,
            });
        }
    }
    for operation in [
        TextOp::Concatenate,
        TextOp::Exact,
        TextOp::Find,
        TextOp::Left,
        TextOp::Len,
        TextOp::Lower,
        TextOp::Mid,
        TextOp::Proper,
        TextOp::Replace,
        TextOp::Right,
        TextOp::Search,
        TextOp::Substitute,
        TextOp::Trim,
    ] {
        cases.push(Case {
            name: format!(
                "matrix-broadcast-text-{}",
                operation.name().to_ascii_lowercase()
            ),
            source: text_source(operation, TextLane::MatrixBroadcast),
            shape: Shape::ReferenceArray {
                rows: 16,
                columns: 4,
            },
            operation: Operation::Text(operation, TextLane::MatrixBroadcast),
            expectation: text_expectation(operation),
        });
    }
    for operation in [
        TextOp::Concatenate,
        TextOp::Len,
        TextOp::Substitute,
        TextOp::Search,
    ] {
        let lower = operation.name().to_ascii_lowercase();
        cases.push(Case {
            name: format!("cancellation-text-{lower}"),
            source: text_source(operation, TextLane::Reference64),
            shape: Shape::ReferenceArray {
                rows: 16,
                columns: 4,
            },
            operation: Operation::Text(operation, TextLane::Cancellation),
            expectation: Expectation::Failure(FailureExpectation::Cancelled),
        });
        cases.push(Case {
            name: format!("resource-text-{lower}"),
            source: text_source(operation, TextLane::Reference64),
            shape: Shape::ReferenceArray {
                rows: 16,
                columns: 4,
            },
            operation: Operation::Text(operation, TextLane::Resource),
            expectation: Expectation::Failure(FailureExpectation::ReferenceCells),
        });
    }
    for operation in [
        TextOp::Find,
        TextOp::Search,
        TextOp::Substitute,
        TextOp::Exact,
    ] {
        cases.push(Case {
            name: format!("search-worstcase-{}", operation.name().to_ascii_lowercase()),
            source: text_source(operation, TextLane::SearchWorstCase),
            shape: Shape::Scalar,
            operation: Operation::Text(operation, TextLane::SearchWorstCase),
            expectation: text_expectation(operation),
        });
    }
    cases.push(Case {
        name: "rept-growth".to_owned(),
        source: text_source(TextOp::Rept, TextLane::ReptGrowth),
        shape: Shape::Scalar,
        operation: Operation::Text(TextOp::Rept, TextLane::ReptGrowth),
        expectation: text_expectation(TextOp::Rept),
    });
    cases.push(Case {
        name: "text-format-fraction-six".to_owned(),
        source: text_source(TextOp::Text, TextLane::FormatterFraction),
        shape: Shape::Scalar,
        operation: Operation::Text(TextOp::Text, TextLane::FormatterFraction),
        expectation: Expectation::TextExact("1/3"),
    });
    for operation in [TextOp::Asc, TextOp::Jis] {
        cases.push(Case {
            name: format!(
                "asc-jis-expansion-{}",
                operation.name().to_ascii_lowercase()
            ),
            source: text_source(operation, TextLane::AscJisExpansion),
            shape: Shape::Scalar,
            operation: Operation::Text(operation, TextLane::AscJisExpansion),
            expectation: text_expectation(operation),
        });
    }
    cases
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

    for name in [
        "concat-borrowed-literals",
        "concat-owned-left",
        "concat-owned-right",
        "concat-growth-chain",
    ] {
        cases.push(Case {
            name: name.to_owned(),
            source: text_control_source(name),
            shape: Shape::Scalar,
            operation: Operation::Text(TextOp::Concatenate, TextLane::Tiny),
            expectation: Expectation::Text,
        });
    }
    cases.extend(byte_cases());
    cases.extend(projection_cases());
    cases
}

fn default_repeat(case: &Case) -> usize {
    if matches!(
        case.expectation,
        Expectation::Failure(FailureExpectation::Cancelled)
    ) {
        return 4;
    }
    if case.name == "text-format-fraction-six" {
        return 1;
    }
    match case.shape {
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
        "ods-formula-text-functions-performance",
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
        | Expectation::NumberAny
        | Expectation::LogicalAny
        | Expectation::Error(_)
        | Expectation::Text
        | Expectation::TextExact(_)
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
        Expectation::Number(_)
        | Expectation::NumberAny
        | Expectation::LogicalAny
        | Expectation::Error(_)
        | Expectation::Text
        | Expectation::TextExact(_) => 0,
    }
}

fn scalar_checksum(value: &ScalarValue<'_>) -> AnyResult<u64> {
    match value {
        ScalarValue::Number(value) => Ok(value.to_bits().rotate_left(11)),
        ScalarValue::Logical(value) => Ok(u64::from(*value).rotate_left(11)),
        ScalarValue::Text(value) => Ok(text_checksum(value.as_ref())),
        ScalarValue::Error(error) => Ok((*error as u64).rotate_left(11)),
        other => Err(format!("expected scalar Number or Text, got {other:?}").into()),
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
        other => Err(format!("expected Number or Text, got {other:?}").into()),
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

#[derive(Debug)]
enum ByteExpected {
    Number(f64),
    Text(String),
    Error(litchi_ods::codec::formula::evaluation::ScalarError),
}

fn byte_text_for(lane: ByteLane, index: usize) -> String {
    match lane {
        ByteLane::Tiny => "Aé界🙂".to_owned(),
        ByteLane::LargeUnicode => large_unicode_text(),
        ByteLane::LargeAscii => large_ascii_text(),
        ByteLane::Reference64 | ByteLane::MatrixBroadcast => {
            TEXT_FIXTURE[index % TEXT_FIXTURE.len()].to_owned()
        },
        ByteLane::SearchWorstCase => "a".repeat(4096) + "needle",
        ByteLane::ReplacementGrowth => large_ascii_text(),
        ByteLane::Refusal | ByteLane::Cancellation | ByteLane::Resource => String::new(),
    }
}

fn byte_search_needle(lane: ByteLane) -> String {
    match lane {
        ByteLane::LargeUnicode => "β".to_owned(),
        ByteLane::LargeAscii => "fox".to_owned(),
        ByteLane::Reference64 | ByteLane::MatrixBroadcast => "a".to_owned(),
        ByteLane::SearchWorstCase => "a".repeat(127) + "needle",
        _ => "界".to_owned(),
    }
}

fn byte_start(text: &str, position: usize) -> Option<usize> {
    let offset = position.checked_sub(1)?;
    if offset > text.len() {
        return None;
    }
    Some(
        (0..=offset)
            .rev()
            .find(|candidate| text.is_char_boundary(*candidate))
            .unwrap_or(0),
    )
}

fn byte_span_end(text: &str, start: usize, length: usize) -> usize {
    let mut end = start;
    let mut used = 0usize;
    for character in text[start..].chars() {
        let width = character.len_utf8();
        if used.saturating_add(width) > length {
            break;
        }
        used = used.saturating_add(width);
        end = end.saturating_add(width);
    }
    end
}

fn byte_search(text: &str, needle: &str, start: usize, insensitive: bool) -> Option<f64> {
    let start = byte_start(text, start)?;
    if needle.is_empty() {
        return Some((start + 1) as f64);
    }
    if !insensitive {
        return text[start..]
            .find(needle)
            .map(|offset| (start + offset + 1) as f64);
    }
    let folded_needle = needle.to_lowercase();
    text[start..].char_indices().find_map(|(offset, _)| {
        let candidate = start + offset;
        text.get(candidate..)
            .filter(|suffix| suffix.to_lowercase().starts_with(&folded_needle))
            .map(|_| (candidate + 1) as f64)
    })
}

fn byte_expected(operation: ByteOp, lane: ByteLane, index: usize) -> ByteExpected {
    if matches!(lane, ByteLane::Refusal) {
        return ByteExpected::Error(litchi_ods::codec::formula::evaluation::ScalarError::Value);
    }
    let text = byte_text_for(lane, index);
    let row = index / 4;
    let companion = row.saturating_mul(4) % 3 + 1;
    let (start, length) = match lane {
        ByteLane::MatrixBroadcast => (companion, companion),
        ByteLane::Tiny
        | ByteLane::LargeUnicode
        | ByteLane::LargeAscii
        | ByteLane::Reference64
        | ByteLane::SearchWorstCase
        | ByteLane::ReplacementGrowth
        | ByteLane::Cancellation
        | ByteLane::Resource => (3, 5),
        ByteLane::Refusal => (1, 1),
    };
    match operation {
        ByteOp::LenB => ByteExpected::Number(text.len() as f64),
        ByteOp::LeftB => {
            let end = byte_span_end(
                &text,
                0,
                if matches!(lane, ByteLane::MatrixBroadcast) {
                    length
                } else {
                    7
                },
            );
            ByteExpected::Text(text[..end].to_owned())
        },
        ByteOp::RightB => {
            let target = if matches!(lane, ByteLane::MatrixBroadcast) {
                length
            } else {
                7
            };
            let mut start_offset = text.len().saturating_sub(target);
            while start_offset > 0 && !text.is_char_boundary(start_offset) {
                start_offset += 1;
            }
            ByteExpected::Text(text[start_offset..].to_owned())
        },
        ByteOp::MidB => {
            let start_position = if matches!(lane, ByteLane::MatrixBroadcast) {
                start
            } else {
                3
            };
            let Some(start_offset) = byte_start(&text, start_position) else {
                return ByteExpected::Text(String::new());
            };
            let end = byte_span_end(
                &text,
                start_offset,
                if matches!(lane, ByteLane::MatrixBroadcast) {
                    2
                } else {
                    5
                },
            );
            ByteExpected::Text(text[start_offset..end].to_owned())
        },
        ByteOp::ReplaceB => {
            if matches!(lane, ByteLane::ReplacementGrowth) {
                return ByteExpected::Text(large_ascii_text().replacen("", &"x".repeat(4096), 1));
            }
            let start_position = if matches!(lane, ByteLane::MatrixBroadcast) {
                start
            } else {
                3
            };
            let remove_length = 2;
            let start_offset = byte_start(&text, start_position).unwrap_or(text.len());
            let end = byte_span_end(&text, start_offset, remove_length);
            let replacement = if matches!(lane, ByteLane::MatrixBroadcast) {
                "Z"
            } else {
                "XYZ"
            };
            ByteExpected::Text(format!(
                "{}{}{}",
                &text[..start_offset],
                replacement,
                &text[end..]
            ))
        },
        ByteOp::FindB | ByteOp::SearchB => {
            let needle = byte_search_needle(lane);
            let insensitive = matches!(operation, ByteOp::SearchB);
            match byte_search(&text, &needle, 1, insensitive) {
                Some(value) => ByteExpected::Number(value),
                None => {
                    ByteExpected::Error(litchi_ods::codec::formula::evaluation::ScalarError::Value)
                },
            }
        },
    }
}

fn validate_byte_scalar_exact(case: &Case, value: &ScalarValue<'_>) -> AnyResult<()> {
    let Operation::Byte(operation, lane) = case.operation else {
        return Ok(());
    };
    let expected = byte_expected(operation, lane, 0);
    match (expected, value) {
        (ByteExpected::Number(expected), ScalarValue::Number(actual))
            if approximately_equal(*actual, expected) =>
        {
            Ok(())
        },
        (ByteExpected::Text(expected), ScalarValue::Text(actual))
            if actual.as_ref() == expected =>
        {
            Ok(())
        },
        (ByteExpected::Error(expected), ScalarValue::Error(actual)) if *actual == expected => {
            Ok(())
        },
        (expected, observed) => Err(format!(
            "{} exact byte oracle mismatch: expected {expected:?}, observed {observed:?}",
            case.name
        )
        .into()),
    }
}

fn validate_byte_value_exact(case: &Case, result: &Evaluated<'_>) -> AnyResult<()> {
    let Operation::Byte(operation, lane) = case.operation else {
        return Ok(());
    };
    if matches!(case.expectation, Expectation::Error(_)) {
        return Ok(());
    }
    let array = result
        .as_array()
        .ok_or_else(|| format!("{} exact byte result was not an array", case.name))?;
    for index in 0..array.len() {
        let observed = array.get(index).ok_or("missing byte oracle cell")?;
        let expected = byte_expected(operation, lane, index);
        match (expected, observed) {
            (ByteExpected::Number(expected), Value::Number(actual)) if approximately_equal(actual, expected) => {},
            (ByteExpected::Text(expected), Value::Text(actual)) if actual == expected => {},
            (ByteExpected::Error(expected), Value::Error(actual)) if actual == expected => {},
            (expected, observed) => return Err(format!(
                "{} exact byte oracle mismatch at {index}: expected {expected:?}, observed {observed:?}",
                case.name
            ).into()),
        }
    }
    Ok(())
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
        (Expectation::Text, ScalarValue::Text(_)) => Ok(()),
        (Expectation::TextExact(expected), ScalarValue::Text(actual))
            if actual.as_ref() == expected =>
        {
            Ok(())
        },
        (Expectation::TextExact(expected), observed) => Err(format!(
            "{} returned {observed:?}, expected scalar text {expected:?}",
            case.name
        )
        .into()),
        (Expectation::Text, observed) => {
            Err(format!("{} returned {observed:?}, expected scalar text", case.name).into())
        },
        (Expectation::Number(expected), observed) => Err(format!(
            "{} returned {observed:?}, expected scalar number {expected}",
            case.name
        )
        .into()),
        (Expectation::NumberAny, ScalarValue::Number(_)) => Ok(()),
        (Expectation::NumberAny, observed) => Err(format!(
            "{} returned {observed:?}, expected scalar number",
            case.name
        )
        .into()),
        (Expectation::LogicalAny, ScalarValue::Logical(_)) => Ok(()),
        (Expectation::LogicalAny, observed) => Err(format!(
            "{} returned {observed:?}, expected scalar logical",
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
    if case.operation.is_text() {
        match case.expectation {
            Expectation::Text => {
                if case.shape.is_scalar_input() {
                    if !matches!(result.value(), Value::Text(_)) {
                        return Err(format!(
                            "{} returned {:?}, expected text",
                            case.name,
                            result.value()
                        )
                        .into());
                    }
                } else {
                    let array = result.as_array().ok_or_else(|| {
                        format!(
                            "{} returned {:?}, expected text array",
                            case.name,
                            result.value()
                        )
                    })?;
                    let (rows, columns) = case.shape.dimensions().ok_or("missing text shape")?;
                    if (array.shape().rows(), array.shape().columns()) != (rows, columns) {
                        return Err(format!("{} returned wrong text shape", case.name).into());
                    }
                    for index in 0..array.len() {
                        if !matches!(array.get(index).ok_or("missing text cell")?, Value::Text(_)) {
                            return Err(format!("{} cell {index} was not text", case.name).into());
                        }
                    }
                }
            },
            Expectation::TextExact(expected) => {
                if case.shape.is_scalar_input() {
                    if !matches!(result.value(), Value::Text(actual) if actual == expected) {
                        return Err(format!(
                            "{} returned {:?}, expected text {expected:?}",
                            case.name,
                            result.value()
                        )
                        .into());
                    }
                } else {
                    let array = result.as_array().ok_or_else(|| {
                        format!(
                            "{} returned {:?}, expected text array",
                            case.name,
                            result.value()
                        )
                    })?;
                    for index in 0..array.len() {
                        if !matches!(
                            array.get(index).ok_or("missing text cell")?,
                            Value::Text(actual) if actual == expected
                        ) {
                            return Err(format!(
                                "{} cell {index} did not equal expected text {expected:?}",
                                case.name
                            )
                            .into());
                        }
                    }
                }
            },
            Expectation::NumberAny => {
                if case.shape.is_scalar_input() {
                    if !matches!(result.value(), Value::Number(_)) {
                        return Err(format!(
                            "{} returned {:?}, expected number",
                            case.name,
                            result.value()
                        )
                        .into());
                    }
                } else {
                    let array = result.as_array().ok_or("expected number array")?;
                    for index in 0..array.len() {
                        if !matches!(
                            array.get(index).ok_or("missing number cell")?,
                            Value::Number(_)
                        ) {
                            return Err(format!(
                                "{} cell {index} was not number: {:?}",
                                case.name,
                                array.get(index)
                            )
                            .into());
                        }
                    }
                }
            },
            Expectation::LogicalAny => {
                if case.shape.is_scalar_input() {
                    if !matches!(result.value(), Value::Logical(_)) {
                        return Err(format!(
                            "{} returned {:?}, expected logical",
                            case.name,
                            result.value()
                        )
                        .into());
                    }
                } else {
                    let array = result.as_array().ok_or("expected logical array")?;
                    for index in 0..array.len() {
                        if !matches!(
                            array.get(index).ok_or("missing logical cell")?,
                            Value::Logical(_)
                        ) {
                            return Err(
                                format!("{} cell {index} was not logical", case.name).into()
                            );
                        }
                    }
                }
            },
            Expectation::Error(expected) => {
                if !matches!(result.value(), Value::Error(actual) if actual == expected) {
                    return Err(format!(
                        "{} returned {:?}, expected text error {expected}",
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
            Expectation::Number(_) => {
                return Err(format!("{} has invalid text expectation", case.name).into());
            },
        }
        return Ok(());
    }
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
            Expectation::NumberAny | Expectation::LogicalAny => {
                return Err(format!("{} has invalid non-text expectation", case.name).into());
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
            Expectation::Text | Expectation::TextExact(_) => {
                return Err(format!("{} has invalid text expectation", case.name).into());
            },
        }
        return Ok(());
    }
    if let Operation::Projection(operation) = case.operation {
        let array = result
            .as_array()
            .ok_or_else(|| format!("{} returned a scalar, expected projected array", case.name))?;
        let (rows, columns) = case
            .shape
            .dimensions()
            .ok_or("missing projected dimensions")?;
        if (array.shape().rows(), array.shape().columns()) != (rows, columns) {
            return Err(format!("{} returned wrong projected shape", case.name).into());
        }
        for index in 0..array.len() {
            let actual = match array.get(index).ok_or("missing projected cell")? {
                Value::Number(value) => value,
                value => {
                    return Err(format!("{} projected cell {index}: {value:?}", case.name).into());
                },
            };
            let expected = projection_expected(operation, index);
            if !approximately_equal(actual, expected) {
                return Err(format!(
                    "{} projected cell {index} returned {actual}, expected {expected}",
                    case.name
                )
                .into());
            }
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
            Expectation::NumberAny | Expectation::LogicalAny => {
                return Err(format!("{} has invalid non-text expectation", case.name).into());
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
            Expectation::Text | Expectation::TextExact(_) => {
                return Err(format!("{} has invalid text expectation", case.name).into());
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

fn scalar_evaluation_limits(case: &Case) -> EvaluationLimits {
    if case.name == "text-format-fraction-six" {
        EvaluationLimits::default().with_max_steps(FRACTION_PROFILE_MAX_STEPS)
    } else {
        EvaluationLimits::default()
    }
}

fn eval_scalar<'a>(
    case: &Case,
    expression: &'a Expression,
    execution: &ExecutionContext,
) -> Result<litchi_ods::codec::formula::evaluation::EvaluatedScalar<'a>, EvaluationFailure> {
    evaluate_scalar(
        expression,
        &EvaluationContext::new(execution),
        &scalar_evaluation_limits(case),
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
            case.operation.uses_text_inputs(),
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
                validate_byte_value_exact(case, &result)?;
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
        let result = eval_scalar(case, &expression, &execution)
            .map_err(|error| format!("{} scalar preflight failed: {error}", case.name))?;
        validate_scalar(case, result.value())?;
        validate_byte_scalar_exact(case, result.value())?;
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
                    "usage: ods-formula-text-functions-performance [--case NAME|all] [--phase evaluate|parse-evaluate] [--warmups N] [--iterations N] [--repeat N] [--preflight-only]"
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
    output_bytes: &mut u64,
) -> AnyResult<u64> {
    let result = result.map_err(|error| format!("scalar evaluation failed: {error}"))?;
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
    if matches!(case.expectation, Expectation::Failure(_)) {
        return Err(format!("{} unexpectedly produced a value", case.name).into());
    }
    if case.operation.is_text() && matches!(case.expectation, Expectation::Error(_)) {
        validate_value(case, &result)?;
        *checksum = checksum.wrapping_add(value_checksum(result.value())?);
        black_box(*checksum);
        let retained = execution.budget().used(Resource::Memory);
        drop(result);
        return Ok(retained);
    }
    let is_forecast_array = matches!(
        case.operation,
        Operation::Paired(PairedOp::Forecast) | Operation::NestedPaired(PairedOp::Forecast)
    ) && case.name.starts_with("forecast-query-array");
    let is_projected_array = matches!(case.operation, Operation::Projection(_));
    if (case.operation.is_aggregate() && !is_forecast_array && !is_projected_array)
        || case.shape.is_scalar_input()
        || matches!(
            case.operation,
            Operation::Control(ControlOp::DSum | ControlOp::DVar | ControlOp::DStdev)
        )
    {
        *checksum = checksum.wrapping_add(value_checksum(result.value())?);
        *output_bytes = output_bytes.saturating_add(value_output_bytes(result.value()));
    } else {
        let array = result.as_array().ok_or("missing array")?;
        *checksum = checksum.wrapping_add(array_checksum(array)?);
        *output_bytes = output_bytes.saturating_add(array_output_bytes(array));
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
        .map(|value| value_output_bytes(value))
        .sum()
}

fn input_bytes(case: &Case) -> u64 {
    match case.operation {
        Operation::Text(_, lane) => match lane {
            TextLane::Tiny => 12,
            TextLane::LargeUnicode => large_unicode_text().len() as u64,
            TextLane::Reference64 | TextLane::MatrixBroadcast => {
                let fixture_bytes = (64
                    * TEXT_FIXTURE.iter().map(|value| value.len()).sum::<usize>()
                    / TEXT_FIXTURE.len()) as u64;
                if matches!(lane, TextLane::Reference64)
                    && matches!(case.operation, Operation::Text(TextOp::Exact, _))
                {
                    fixture_bytes.saturating_mul(2)
                } else {
                    fixture_bytes
                }
            },
            TextLane::SearchWorstCase => 4096 + 133,
            TextLane::ReptGrowth => 8192,
            TextLane::AscJisExpansion => 16,
            TextLane::FormatterFraction => {
                ("0.3333333333333333".len() + "######/######".len()) as u64
            },
            TextLane::Refusal | TextLane::Cancellation | TextLane::Resource => 0,
        },
        Operation::Byte(operation, lane) => byte_input_bytes(operation, lane),
        _ => case.source.len() as u64,
    }
}

fn byte_input_bytes(operation: ByteOp, lane: ByteLane) -> u64 {
    let fixture_bytes = (64 * TEXT_FIXTURE.iter().map(|value| value.len()).sum::<usize>()
        / TEXT_FIXTURE.len()) as u64;
    match lane {
        ByteLane::Tiny => 10,
        ByteLane::LargeUnicode => large_unicode_text().len() as u64,
        ByteLane::LargeAscii => large_ascii_text().len() as u64,
        ByteLane::Reference64 | ByteLane::MatrixBroadcast => {
            if operation.is_search() {
                fixture_bytes.saturating_add(1)
            } else {
                fixture_bytes
            }
        },
        ByteLane::SearchWorstCase => 4096 + 6 + 127,
        ByteLane::ReplacementGrowth => large_ascii_text().len() as u64 + 4096,
        ByteLane::Refusal | ByteLane::Cancellation | ByteLane::Resource => 0,
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
        case.operation.uses_text_inputs(),
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
    let scalar_limits = scalar_evaluation_limits(case);
    let value_limits = value_limits(case);
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut checksum = 0_u64;
    let mut output_bytes = 0_u64;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        retained_peak = if case.shape.is_scalar_input() && !case.shape.is_reference() {
            retained_peak.max(consume_scalar(
                case,
                evaluate_scalar(expression, &scalar_context, &scalar_limits),
                &execution,
                &mut checksum,
                &mut output_bytes,
            )?)
        } else {
            retained_peak.max(consume_value(
                value::evaluate(expression, &resolver, &context, &value_limits),
                case,
                &execution,
                &mut checksum,
                &mut output_bytes,
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
        output_bytes,
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
        case.operation.uses_text_inputs(),
        matches!(
            case.expectation,
            Expectation::Failure(FailureExpectation::Cancelled)
        )
        .then_some(cancellation.clone()),
    );
    let context =
        Context::new(&execution, Position::new("Main", 0, 0)).with_mode(case.shape.mode());
    let scalar_context = EvaluationContext::new(&execution);
    let scalar_limits = scalar_evaluation_limits(case);
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
        retained_peak = if case.shape.is_scalar_input() && !case.shape.is_reference() {
            retained_peak.max(consume_scalar(
                case,
                evaluate_scalar(&expression, &scalar_context, &scalar_limits),
                &execution,
                &mut checksum,
                &mut output_bytes,
            )?)
        } else {
            retained_peak.max(consume_value(
                value::evaluate(&expression, &resolver, &context, &value_limits),
                case,
                &execution,
                &mut checksum,
                &mut output_bytes,
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
        output_bytes,
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
        sample.checksum,
    )
}

fn expectation_label(expectation: Expectation) -> String {
    match expectation {
        Expectation::Number(_) => "finite-number".to_owned(),
        Expectation::NumberAny => "number".to_owned(),
        Expectation::LogicalAny => "logical".to_owned(),
        Expectation::Error(error) => format!("error:{error}"),
        Expectation::Text => "text".to_owned(),
        Expectation::TextExact(expected) => format!("text:{expected}"),
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
        "{{\"case\":\"{}\",\"operation\":\"{}\",\"phase\":\"{}\",\"input_bytes\":{},\"output_bytes_p50\":{},\"bytes_per_repeat_p50\":{},\"repeat\":{},\"warmups\":{},\"iterations\":{},\"shape\":\"{}\",\"rows\":{},\"columns\":{},\"elements\":{},\"supported\":true,\"expected\":\"{}\",\"elapsed_ns_p50\":{},\"elapsed_ns_mean\":{},\"elapsed_ns_p95\":{},\"elapsed_ns_p99\":{},\"elapsed_ns_per_repeat\":{},\"allocator_calls_p50\":{},\"requested_bytes_p50\":{},\"released_bytes_p50\":{},\"peak_live_delta_p50\":{},\"memory_retained_p50\":{},\"work_p50\":{},\"work_per_repeat\":{},\"reference_reads_p50\":{},\"reference_reads_per_repeat\":{},\"checksum_p50\":{},\"live_before_p50\":{},\"live_after_p50\":{},\"samples\":{},\"rss_kib\":null,\"rss_source\":\"external /usr/bin/time -v\",\"memory_retained_source\":\"execution budget while result is live\",\"reference_source\":\"instrumented borrowing resolver\",\"validation_scope\":\"one untimed contract fixture oracle; timed evaluator/checksum/drop\"}}",
        json_escape(&case.name),
        case.operation.name(),
        phase.label(),
        input_bytes(case),
        percentile(samples, |sample| sample.output_bytes, 50),
        input_bytes(case)
            .saturating_add(percentile(samples, |sample| sample.output_bytes, 50) / repeat as u64,),
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
    let repeat = config.repeat.unwrap_or_else(|| default_repeat(case));
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
