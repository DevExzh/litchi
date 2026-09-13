//! Candidate-only matrix-function performance harness.
//!
//! The corpus exercises the five matrix functions with deterministic inline
//! numeric arrays. Parsing and the no-op resolver are prepared before the
//! evaluate timer; parse-evaluate is provided as a separate caller-cost lane.
//! One untimed preflight result per case is checked against an independent
//! mathematical oracle before timing begins. Formula-level error values are
//! expected for the bounded invalid-shape, singular, and MUNIT(0) cases;
//! evaluator/resource failures remain harness failures.
//!
//! JSON-lines output reports median batch values over the requested iterations.
//! memory_retained is the execution-budget reservation observed while a result
//! is live, not transient allocator peak. RSS is left to the external
//! /usr/bin/time -v wrapper used by capture scripts. The evaluate timer includes
//! evaluator execution, checksum folding, and result drop; it excludes parser,
//! fixture, and oracle work. This binary has no library statistics or benchmark
//! dependencies.

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

/// The matrix corpus contains no references. The resolver still supplies a
/// finite immutable workbook contract because the value evaluator requires a
/// resolver for all expressions, including inline arrays.
#[derive(Debug, Default)]
struct EmptyResolver;

impl Resolver for EmptyResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        execution.check()?;
        Ok((sheet == "Main").then_some(SheetExtent::new(1, 1)))
    }

    fn read_cell<'a>(
        &'a self,
        _sheet: &str,
        _row: usize,
        _column: usize,
        execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        execution.check()?;
        Ok(CellRead::Unsupported)
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
enum Oracle {
    Determinant {
        size: usize,
    },
    Inverse {
        size: usize,
    },
    Multiply {
        left_rows: usize,
        left_columns: usize,
        right_columns: usize,
    },
    Identity {
        size: usize,
    },
    Transpose {
        rows: usize,
        columns: usize,
    },
    Scalar(f64),
    FormulaError,
}

#[derive(Debug)]
struct CaseSpec {
    name: String,
    function: &'static str,
    source: String,
    size: usize,
    oracle: Oracle,
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
                    "usage: matrix-profile [--case NAME|all] [--phase evaluate|parse-evaluate|parse] [--warmups N] [--iterations N] [--repeat N]\n\
                     --list prints deterministic case names; evaluate reuses parsed expressions and a bounded no-op resolver.\n\
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
        0..=2 => 16,
        8 => 4,
        16 => 2,
        _ => 1,
    }
}

fn array_literal(rows: usize, columns: usize, cell: impl Fn(usize, usize) -> f64) -> String {
    let mut output = String::from("{");
    for row in 0..rows {
        if row != 0 {
            output.push('|');
        }
        for column in 0..columns {
            if column != 0 {
                output.push(';');
            }
            output.push_str(&format_number(cell(row, column)));
        }
    }
    output.push('}');
    output
}

fn format_number(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < i64::MAX as f64 {
        format!("{}", value as i64)
    } else {
        format!("{value:.15}")
    }
}

fn dense_value(row: usize, column: usize) -> f64 {
    if row == column { 2.0 } else { 1.0 }
}

fn sequential_value(row: usize, column: usize, columns: usize) -> f64 {
    (row * columns + column + 1) as f64
}

fn matrix_function_source(function: &str, size: usize) -> String {
    match function {
        "MDETERM" | "MINVERSE" => {
            format!("={function}({})", array_literal(size, size, dense_value))
        },
        "MMULT" => {
            let left = array_literal(size, size, dense_value);
            let right = array_literal(size, size, |row, column| {
                sequential_value(row, column, size)
            });
            format!("=MMULT({left};{right})")
        },
        "MUNIT" => format!("=MUNIT({size})"),
        "TRANSPOSE" => format!(
            "=TRANSPOSE({})",
            array_literal(size, size, |row, column| {
                sequential_value(row, column, size)
            })
        ),
        _ => unreachable!("unknown matrix function"),
    }
}

fn expectation_for(function: &'static str, size: usize) -> Oracle {
    match function {
        "MDETERM" => Oracle::Determinant { size },
        "MINVERSE" => Oracle::Inverse { size },
        "MMULT" => Oracle::Multiply {
            left_rows: size,
            left_columns: size,
            right_columns: size,
        },
        "MUNIT" => Oracle::Identity { size },
        "TRANSPOSE" => Oracle::Transpose {
            rows: size,
            columns: size,
        },
        _ => unreachable!("unknown matrix function"),
    }
}

fn base_case(function: &'static str, size: usize) -> CaseSpec {
    CaseSpec {
        name: format!("{}-{size}", function.to_ascii_lowercase()),
        function,
        source: matrix_function_source(function, size),
        size,
        oracle: expectation_for(function, size),
    }
}

fn corpus() -> Vec<CaseSpec> {
    let mut cases = Vec::new();
    for size in [2, 8, 16] {
        for function in ["MDETERM", "MINVERSE", "MMULT", "MUNIT", "TRANSPOSE"] {
            cases.push(base_case(function, size));
        }
    }

    cases.push(base_case("TRANSPOSE", 64));
    cases.push(base_case("MUNIT", 64));

    let left = array_literal(4, 8, dense_value);
    let right = array_literal(8, 2, |row, column| sequential_value(row, column, 2));
    cases.push(CaseSpec {
        name: "mmult-rect-4x8x2".to_owned(),
        function: "MMULT",
        source: format!("=MMULT({left};{right})"),
        size: 8,
        oracle: Oracle::Multiply {
            left_rows: 4,
            left_columns: 8,
            right_columns: 2,
        },
    });
    cases.push(CaseSpec {
        name: "transpose-rect-2x8".to_owned(),
        function: "TRANSPOSE",
        source: format!(
            "=TRANSPOSE({})",
            array_literal(2, 8, |row, column| sequential_value(row, column, 8))
        ),
        size: 8,
        oracle: Oracle::Transpose {
            rows: 2,
            columns: 8,
        },
    });

    cases.push(CaseSpec {
        name: "minverse-singular-2".to_owned(),
        function: "MINVERSE",
        source: "=MINVERSE({1;2|2;4})".to_owned(),
        size: 2,
        oracle: Oracle::FormulaError,
    });
    cases.push(CaseSpec {
        name: "mdeterm-nonsquare-2x1".to_owned(),
        function: "MDETERM",
        source: "=MDETERM({1|2})".to_owned(),
        size: 2,
        oracle: Oracle::FormulaError,
    });
    cases.push(CaseSpec {
        name: "mmult-incompatible-2x3-2x2".to_owned(),
        function: "MMULT",
        source: "=MMULT({1;2;3|4;5;6};{1;2|3;4})".to_owned(),
        size: 3,
        oracle: Oracle::FormulaError,
    });
    cases.push(CaseSpec {
        name: "munit-zero".to_owned(),
        function: "MUNIT",
        source: "=MUNIT(0)".to_owned(),
        size: 0,
        oracle: Oracle::FormulaError,
    });

    for function in ["MDETERM", "MINVERSE", "MMULT", "MUNIT", "TRANSPOSE"] {
        let base = base_case(function, 8);
        let branch = base
            .source
            .strip_prefix('=')
            .expect("matrix source starts with =");
        cases.push(CaseSpec {
            name: format!("if-selected-{}-8", function.to_ascii_lowercase()),
            function,
            source: format!("=IF(TRUE();{branch};0)"),
            size: 8,
            oracle: base.oracle,
        });
        cases.push(CaseSpec {
            name: format!("if-unselected-{}-8", function.to_ascii_lowercase()),
            function,
            source: format!("=IF(FALSE();{branch};0)"),
            size: 8,
            oracle: Oracle::Scalar(0.0),
        });
    }
    cases
}

fn execution() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "ods-formula-matrix-functions-performance",
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
    (cancellation, ExecutionContext::new(budget, token, limits))
}

fn validate_number(value: Value<'_>, expected: f64, case: &str) -> AnyResult<()> {
    match value {
        Value::Number(actual) if approximately_equal(actual, expected) => Ok(()),
        other => Err(format!("{case} returned {other:?}, expected Number({expected})").into()),
    }
}

fn approximately_equal(actual: f64, expected: f64) -> bool {
    let tolerance = 1e-9 * expected.abs().max(1.0);
    (actual - expected).abs() <= tolerance
}

fn validate_array_shape(
    array: ArrayView<'_>,
    rows: usize,
    columns: usize,
    case: &str,
) -> AnyResult<()> {
    if array.shape().rows() != rows || array.shape().columns() != columns {
        return Err(format!(
            "{case} returned shape {}x{}, expected {rows}x{columns}",
            array.shape().rows(),
            array.shape().columns()
        )
        .into());
    }
    Ok(())
}

fn expected_matrix_number(oracle: Oracle, row: usize, column: usize) -> f64 {
    match oracle {
        Oracle::Inverse { size } => {
            // The fixture is I + J, where every entry of J is one.
            // Since J² = nJ, its inverse is I - J/(n+1).
            let denominator = (size + 1) as f64;
            if row == column {
                size as f64 / denominator
            } else {
                -1.0 / denominator
            }
        },
        Oracle::Multiply {
            left_rows: _,
            left_columns,
            right_columns,
        } => {
            let mut value = 0.0;
            for inner in 0..left_columns {
                value += dense_value(row, inner) * sequential_value(inner, column, right_columns);
            }
            value
        },
        Oracle::Identity { .. } => f64::from((row == column) as u8),
        Oracle::Transpose { rows: _, columns } => sequential_value(column, row, columns),
        _ => unreachable!("scalar oracle has no matrix elements"),
    }
}

fn validate_array(array: ArrayView<'_>, oracle: Oracle, case: &str) -> AnyResult<()> {
    let (rows, columns) = match oracle {
        Oracle::Inverse { size } | Oracle::Identity { size } | Oracle::Determinant { size } => {
            (size, size)
        },
        Oracle::Multiply {
            left_rows,
            right_columns,
            ..
        } => (left_rows, right_columns),
        Oracle::Transpose { rows, columns } => (columns, rows),
        Oracle::Scalar(_) | Oracle::FormulaError => {
            return Err(format!("{case} has no array oracle").into());
        },
    };
    validate_array_shape(array, rows, columns, case)?;
    for row in 0..rows {
        for column in 0..columns {
            let index = row * columns + column;
            let actual = match array.get(index) {
                Some(Value::Number(number)) => number,
                other => {
                    return Err(format!("{case} cell {index} returned {other:?}").into());
                },
            };
            let expected = expected_matrix_number(oracle, row, column);
            if !approximately_equal(actual, expected) {
                return Err(
                    format!("{case} cell {index} returned {actual}, expected {expected}").into(),
                );
            }
        }
    }
    Ok(())
}

fn validate_value(case: &CaseSpec, value: Value<'_>) -> AnyResult<()> {
    match case.oracle {
        Oracle::Determinant { size } => validate_number(value, (size + 1) as f64, &case.name),
        Oracle::Inverse { .. }
        | Oracle::Multiply { .. }
        | Oracle::Identity { .. }
        | Oracle::Transpose { .. } => {
            let array = match value {
                Value::Array(array) => array,
                other => {
                    return Err(format!("{} returned {other:?}, expected array", case.name).into());
                },
            };
            validate_array(array, case.oracle, &case.name)
        },
        Oracle::Scalar(expected) => validate_number(value, expected, &case.name),
        Oracle::FormulaError => match value {
            Value::Error(_) => Ok(()),
            other => {
                Err(format!("{} returned {other:?}, expected formula error", case.name).into())
            },
        },
    }
}

fn preflight_evaluate(case: &CaseSpec, expression: &Expression) -> AnyResult<()> {
    let resolver = EmptyResolver;
    let (_cancellation, execution) = execution();
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = evaluate(expression, &resolver, &context, &Limits::default())
        .map_err(|error| format!("{} evaluation preflight failed: {error}", case.name))?;
    validate_value(case, result.value())
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
    case_name: &str,
    result: Result<Evaluated<'_>, EvaluationFailure>,
    execution: &ExecutionContext,
    successes: &mut u64,
    refusals: &mut u64,
    checksum: &mut u64,
    retained_peak: &mut u64,
) -> AnyResult<()> {
    match result {
        Ok(value) => {
            *successes = successes.saturating_add(1);
            *checksum = checksum.wrapping_add(checksum_value(value.value()));
            *retained_peak = (*retained_peak).max(execution.budget().used(Resource::Memory));
            black_box(*checksum);
            drop(value);
            Ok(())
        },
        Err(error) => {
            *refusals = refusals.saturating_add(1);
            Err(format!("{case_name} evaluator failure: {error}").into())
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

fn measure_evaluate(case: &CaseSpec, expression: &Expression, repeat: usize) -> AnyResult<Sample> {
    let resolver = EmptyResolver;
    let (_cancellation, execution) = execution();
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let limits = Limits::default();
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
            &case.name,
            evaluate(expression, &resolver, &context, &limits),
            &execution,
            &mut successes,
            &mut refusals,
            &mut checksum,
            &mut retained_peak,
        )?;
    }
    let elapsed_ns = started.elapsed().as_nanos().min(u64::MAX as u128) as u64;
    Ok(Sample {
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
    })
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
    })
}

fn measure_parse_evaluate(case: &CaseSpec, repeat: usize) -> AnyResult<Sample> {
    let resolver = EmptyResolver;
    let (_cancellation, execution) = execution();
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let limits = Limits::default();
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
            &case.name,
            evaluate(&expression, &resolver, &context, &limits),
            &execution,
            &mut successes,
            &mut refusals,
            &mut checksum,
            &mut retained_peak,
        )?;
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
        work: execution
            .budget()
            .used(Resource::Work)
            .saturating_sub(baseline_work),
        memory_retained: retained_peak.saturating_sub(baseline_memory),
        successes,
        refusals,
        checksum,
    })
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

fn expectation_shape(oracle: Oracle) -> (&'static str, usize, usize, usize) {
    match oracle {
        Oracle::Determinant { .. } | Oracle::Scalar(_) => ("scalar", 0, 0, 1),
        Oracle::FormulaError => ("error", 0, 0, 0),
        Oracle::Inverse { size } | Oracle::Identity { size } => {
            ("array", size, size, size.saturating_mul(size))
        },
        Oracle::Multiply {
            left_rows,
            right_columns,
            ..
        } => (
            "array",
            left_rows,
            right_columns,
            left_rows.saturating_mul(right_columns),
        ),
        Oracle::Transpose { rows, columns } => {
            ("array", columns, rows, rows.saturating_mul(columns))
        },
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
    let (shape, rows, columns, elements) = expectation_shape(case.oracle);
    let elapsed_ns = percentile(samples, |sample| sample.elapsed_ns, 50);
    let work = percentile(samples, |sample| sample.work, 50);
    println!(
        "{{\"case\":\"{}\",\"function\":\"{}\",\"size\":{},\"phase\":\"{}\",\"input_bytes\":{},\"repeat\":{},\"warmups\":{},\"iterations\":{},\"elapsed_ns\":{},\"elapsed_ns_mean\":{},\"elapsed_ns_p95\":{},\"elapsed_ns_p99\":{},\"elapsed_ns_per_repeat\":{},\"checksum\":{},\"shape\":\"{}\",\"rows\":{},\"columns\":{},\"elements\":{},\"work\":{},\"work_mean\":{},\"work_per_repeat\":{},\"allocator_calls\":{},\"allocator_calls_max\":{},\"deallocator_calls\":{},\"deallocator_calls_max\":{},\"requested_bytes\":{},\"requested_bytes_max\":{},\"released_bytes\":{},\"released_bytes_max\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"peak_live_delta_max\":{},\"memory_retained\":{},\"memory_retained_max\":{},\"successes\":{},\"refusals\":{},\"failure\":\"none\",\"rss_kib\":null,\"rss_source\":\"external /usr/bin/time -v\",\"memory_retained_source\":\"execution budget while result is live\",\"validation_scope\":\"one untimed preflight; timed evaluate/checksum/drop\"}}",
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
        shape,
        rows,
        columns,
        elements,
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

    // Validate one result before resetting the allocation observer. This keeps
    // fixture/parser setup and the O(n^3) oracle outside the timed lane.
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
        let sample = match config.phase {
            Phase::Evaluate => {
                measure_evaluate(case, parsed.as_ref().expect("parsed evaluate"), repeat)?
            },
            Phase::ParseEvaluate => measure_parse_evaluate(case, repeat)?,
            Phase::Parse => measure_parse(case, repeat)?,
        };
        black_box(sample);
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
