//! Small, candidate-only OpenFormula rounding profile.
//!
//! The source expressions and expected values are prepared before each timed
//! lane. `evaluate` reuses a parsed expression, while `parse-evaluate` keeps
//! parsing in the caller-visible path. Scalar cases use the resolver-free
//! evaluator; array cases use the matrix evaluator with an immutable no-op
//! resolver. One untimed result is checked before the allocation counters are
//! reset. This binary reports absolute observations only; it contains no
//! baseline implementation and makes no speedup claim.

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

/// The rounding corpus uses literal arrays only. The value evaluator still
/// requires a finite, coherent resolver contract, so this provider refuses all
/// reads and performs no allocation or I/O.
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
enum Kind {
    Scalar,
    Array { rows: usize, columns: usize },
}

#[derive(Clone, Copy, Debug)]
enum ArrayOp {
    AddOne,
    Int,
    Floor,
    Round,
    RoundDown,
    Ceiling,
    Mround,
    RoundUp,
    Trunc,
}

#[derive(Debug)]
struct Case {
    name: &'static str,
    source: String,
    kind: Kind,
    array_op: Option<ArrayOp>,
    scalar_expected: Option<f64>,
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
            other => {
                Err(format!("unknown phase {other:?}; expected evaluate or parse-evaluate").into())
            },
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
    let mut iterations = 20;
    let mut repeat = None;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                println!(
                    "usage: ods-formula-rounding-performance [--case NAME|all] [--phase evaluate|parse-evaluate] [--warmups N] [--iterations N] [--repeat N]\n\
                     --list prints the fixed case names. Wrap one invocation in /usr/bin/time -v for RSS."
                );
                return Ok(None);
            },
            "--list" => {
                for case in corpus() {
                    println!("{}", case.name);
                }
                return Ok(None);
            },
            "--case" => case = Some(arguments.next().ok_or("--case requires a name or all")?),
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

fn scalar_case(name: &'static str, source: &'static str, expected: f64) -> Case {
    Case {
        name,
        source: source.to_owned(),
        kind: Kind::Scalar,
        array_op: None,
        scalar_expected: Some(expected),
    }
}

fn array_value(row: usize, column: usize, columns: usize) -> f64 {
    let index = row.saturating_mul(columns).saturating_add(column);
    (index % 37) as f64 / 10.0 + 0.1
}

fn array_literal(rows: usize, columns: usize) -> String {
    let mut output = String::with_capacity(rows.saturating_mul(columns).saturating_mul(5) + 2);
    output.push('{');
    for row in 0..rows {
        if row != 0 {
            output.push('|');
        }
        for column in 0..columns {
            if column != 0 {
                output.push(';');
            }
            let _ = write!(output, "{:.1}", array_value(row, column, columns));
        }
    }
    output.push('}');
    output
}

fn array_case(name: &'static str, rows: usize, columns: usize, op: ArrayOp) -> Case {
    let literal = array_literal(rows, columns);
    let source = match op {
        ArrayOp::AddOne => format!("={literal}+1"),
        ArrayOp::Int => format!("=INT({literal})"),
        ArrayOp::Floor => format!("=FLOOR({literal};0.5)"),
        ArrayOp::Round => format!("=ROUND({literal};1)"),
        ArrayOp::RoundDown => format!("=ROUNDDOWN({literal};0)"),
        ArrayOp::Ceiling => format!("=CEILING({literal};0.5)"),
        ArrayOp::Mround => format!("=MROUND({literal};0.5)"),
        // Digits zero keeps the oracle away from decimal-binary ties while
        // still exercising the directed rounding kernel over a 16x16 array.
        ArrayOp::RoundUp => format!("=ROUNDUP({literal};0)"),
        ArrayOp::Trunc => format!("=TRUNC({literal};0)"),
    };
    debug_assert!(source.starts_with('='));
    Case {
        name,
        source,
        kind: Kind::Array { rows, columns },
        array_op: Some(op),
        scalar_expected: None,
    }
}

fn corpus() -> Vec<Case> {
    vec![
        scalar_case("scalar-control", "=1+2*3", 7.0),
        scalar_case("scalar-int", "=INT(12345.6789)", 12345.0),
        scalar_case("scalar-floor", "=FLOOR(12345.6789;0.25)", 12345.5),
        scalar_case("scalar-round", "=ROUND(12345.6789;2)", 12345.68),
        scalar_case("scalar-rounddown", "=ROUNDDOWN(12345.6789;2)", 12345.67),
        scalar_case("scalar-ceiling", "=CEILING(12345.6789;0.25)", 12345.75),
        scalar_case("scalar-mround", "=MROUND(12345;2)", 12346.0),
        scalar_case("scalar-roundup", "=ROUNDUP(12345.601;1)", 12345.7),
        scalar_case("scalar-trunc", "=TRUNC(12345.6789;2)", 12345.67),
        array_case("array-control-4x4", 4, 4, ArrayOp::AddOne),
        array_case("array-int-4x4", 4, 4, ArrayOp::Int),
        array_case("array-floor-4x4", 4, 4, ArrayOp::Floor),
        array_case("array-round-4x4", 4, 4, ArrayOp::Round),
        array_case("array-rounddown-4x4", 4, 4, ArrayOp::RoundDown),
        array_case("array-ceiling-4x4", 4, 4, ArrayOp::Ceiling),
        array_case("array-mround-4x4", 4, 4, ArrayOp::Mround),
        array_case("array-roundup-4x4", 4, 4, ArrayOp::RoundUp),
        array_case("array-trunc-4x4", 4, 4, ArrayOp::Trunc),
        array_case("array-roundup-16x16", 16, 16, ArrayOp::RoundUp),
    ]
}

fn default_repeat(kind: Kind) -> usize {
    match kind {
        Kind::Scalar => 1_000,
        Kind::Array {
            rows: 4,
            columns: 4,
        } => 100,
        Kind::Array {
            rows: 16,
            columns: 16,
        } => 8,
        Kind::Array { .. } => 1,
    }
}

fn execution() -> ExecutionContext {
    let budget = Budget::root(
        "ods-formula-rounding-performance",
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
    (actual - expected).abs() <= expected.abs().max(1.0) * 1.0e-12
}

fn scalar_checksum(value: &ScalarValue<'_>) -> AnyResult<u64> {
    match value {
        ScalarValue::Number(value) => Ok(value.to_bits().rotate_left(11) ^ 0x4e),
        other => Err(format!("expected a scalar Number, got {other:?}").into()),
    }
}

fn scalar_preflight(case: &Case, expression: &Expression) -> AnyResult<()> {
    let execution = execution();
    let result = evaluate_scalar(
        expression,
        &EvaluationContext::new(&execution),
        &EvaluationLimits::default(),
    )
    .map_err(|error| format!("{} preflight failed: {error}", case.name))?;
    let expected = case.scalar_expected.expect("scalar case expected value");
    match result.value() {
        ScalarValue::Number(actual) if approximately_equal(*actual, expected) => Ok(()),
        other => Err(format!(
            "{} returned {other:?}, expected Number({expected})",
            case.name
        )
        .into()),
    }
}

fn array_expected(op: ArrayOp, value: f64) -> f64 {
    match op {
        ArrayOp::AddOne => value + 1.0,
        ArrayOp::Int => value.floor(),
        ArrayOp::Floor => (value / 0.5).floor() * 0.5,
        ArrayOp::Round => (value * 10.0).round() / 10.0,
        ArrayOp::RoundDown => value.trunc(),
        ArrayOp::Ceiling => (value / 0.5).ceil() * 0.5,
        ArrayOp::Mround => (value / 0.5).round() * 0.5,
        ArrayOp::RoundUp => {
            let scaled = value;
            if scaled < 0.0 {
                scaled.floor()
            } else {
                scaled.ceil()
            }
        },
        ArrayOp::Trunc => value.trunc(),
    }
}

fn array_checksum(array: value::ArrayView<'_>) -> AnyResult<u64> {
    let mut checksum = ((array.shape().rows() as u64) << 32) ^ array.shape().columns() as u64;
    for index in 0..array.len() {
        let cell = array
            .get(index)
            .ok_or_else(|| format!("missing array cell {index}"))?;
        let Value::Number(number) = cell else {
            return Err(format!("array cell {index} was not numeric: {cell:?}").into());
        };
        checksum = checksum.rotate_left(5) ^ number.to_bits();
    }
    Ok(checksum)
}

fn array_preflight(case: &Case, expression: &Expression) -> AnyResult<()> {
    let Kind::Array { rows, columns } = case.kind else {
        unreachable!("array preflight called for scalar case")
    };
    let execution = execution();
    let resolver = EmptyResolver;
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = value::evaluate(expression, &resolver, &context, &Limits::default())
        .map_err(|error| format!("{} preflight failed: {error}", case.name))?;
    let array = result
        .as_array()
        .ok_or_else(|| format!("{} did not return an array", case.name))?;
    if array.shape().rows() != rows || array.shape().columns() != columns {
        return Err(format!(
            "{} returned {}x{}, expected {rows}x{columns}",
            case.name,
            array.shape().rows(),
            array.shape().columns()
        )
        .into());
    }
    let op = case.array_op.expect("array operation");
    for row in 0..rows {
        for column in 0..columns {
            let index = row * columns + column;
            let Value::Number(actual) = array
                .get(index)
                .ok_or_else(|| format!("{} missing cell {index}", case.name))?
            else {
                return Err(format!("{} cell {index} was not numeric", case.name).into());
            };
            let expected = array_expected(op, array_value(row, column, columns));
            if !approximately_equal(actual, expected) {
                return Err(format!("{} cell {index}: {actual} != {expected}", case.name).into());
            }
        }
    }
    Ok(())
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
    retained_peak: &mut u64,
) -> AnyResult<()> {
    let result = result.map_err(|error| format!("scalar evaluation failed: {error}"))?;
    *checksum = checksum.wrapping_add(scalar_checksum(result.value())?);
    *retained_peak = (*retained_peak).max(execution.budget().used(Resource::Memory));
    black_box(*checksum);
    drop(result);
    Ok(())
}

fn consume_array(
    result: Result<Evaluated<'_>, EvaluationFailure>,
    execution: &ExecutionContext,
    checksum: &mut u64,
    retained_peak: &mut u64,
) -> AnyResult<()> {
    let result = result.map_err(|error| format!("array evaluation failed: {error}"))?;
    let array = result
        .as_array()
        .ok_or("array evaluation did not return an array")?;
    *checksum = checksum.wrapping_add(array_checksum(array)?);
    *retained_peak = (*retained_peak).max(execution.budget().used(Resource::Memory));
    black_box(*checksum);
    drop(result);
    Ok(())
}

fn measure_evaluate(case: &Case, expression: &Expression, repeat: usize) -> AnyResult<Sample> {
    let execution = execution();
    let resolver = EmptyResolver;
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let value_limits = Limits::default();
    let scalar_context = EvaluationContext::new(&execution);
    let scalar_limits = EvaluationLimits::default();
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut checksum = 0_u64;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        match case.kind {
            Kind::Scalar => consume_scalar(
                evaluate_scalar(expression, &scalar_context, &scalar_limits),
                &execution,
                &mut checksum,
                &mut retained_peak,
            )?,
            Kind::Array { .. } => consume_array(
                value::evaluate(expression, &resolver, &context, &value_limits),
                &execution,
                &mut checksum,
                &mut retained_peak,
            )?,
        }
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
        checksum,
    })
}

fn measure_parse_evaluate(case: &Case, repeat: usize) -> AnyResult<Sample> {
    let execution = execution();
    let resolver = EmptyResolver;
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let value_limits = Limits::default();
    let scalar_context = EvaluationContext::new(&execution);
    let scalar_limits = EvaluationLimits::default();
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let live_before = reset_observer();
    let started = Instant::now();
    let mut checksum = 0_u64;
    let mut retained_peak = baseline_memory;
    for _ in 0..repeat {
        let expression = Expression::parse(black_box(&case.source))
            .map_err(|error| format!("{} parse failed: {error}", case.name))?;
        match case.kind {
            Kind::Scalar => consume_scalar(
                evaluate_scalar(&expression, &scalar_context, &scalar_limits),
                &execution,
                &mut checksum,
                &mut retained_peak,
            )?,
            Kind::Array { .. } => consume_array(
                value::evaluate(&expression, &resolver, &context, &value_limits),
                &execution,
                &mut checksum,
                &mut retained_peak,
            )?,
        }
        black_box(expression);
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
        checksum,
    })
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
        "{{\"elapsed_ns\":{},\"alloc_calls\":{},\"dealloc_calls\":{},\"requested_bytes\":{},\"released_bytes\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"work\":{},\"memory_retained\":{},\"checksum\":{}}}",
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
        sample.checksum,
    )
}

fn emit(
    case: &Case,
    phase: Phase,
    warmups: usize,
    iterations: usize,
    repeat: usize,
    samples: &[Sample],
) {
    let (shape, rows, columns, elements) = match case.kind {
        Kind::Scalar => ("scalar", 0, 0, 1),
        Kind::Array { rows, columns } => ("array", rows, columns, rows.saturating_mul(columns)),
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
        "{{\"case\":\"{}\",\"phase\":\"{}\",\"input_bytes\":{},\"repeat\":{},\"warmups\":{},\"iterations\":{},\"shape\":\"{}\",\"rows\":{},\"columns\":{},\"elements\":{},\"elapsed_ns_p50\":{},\"elapsed_ns_mean\":{},\"elapsed_ns_p95\":{},\"elapsed_ns_p99\":{},\"elapsed_ns_per_repeat\":{},\"allocator_calls_p50\":{},\"allocator_calls_max\":{},\"deallocator_calls_p50\":{},\"requested_bytes_p50\":{},\"requested_bytes_max\":{},\"released_bytes_p50\":{},\"peak_live_delta_p50\":{},\"peak_live_delta_max\":{},\"memory_retained_p50\":{},\"memory_retained_max\":{},\"work_p50\":{},\"work_per_repeat\":{},\"checksum_p50\":{},\"live_before_p50\":{},\"live_after_p50\":{},\"samples\":{},\"rss_kib\":null,\"rss_source\":\"external /usr/bin/time -v\",\"memory_retained_source\":\"execution budget while result is live\",\"validation_scope\":\"one untimed oracle preflight; timed evaluator/checksum/drop\"}}",
        json_escape(case.name),
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
        samples
            .iter()
            .map(|sample| sample.alloc_calls)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.dealloc_calls, 50),
        percentile(samples, |sample| sample.requested_bytes, 50),
        samples
            .iter()
            .map(|sample| sample.requested_bytes)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.released_bytes, 50),
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
        percentile(samples, |sample| sample.work, 50),
        percentile(samples, |sample| sample.work, 50) / repeat as u64,
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
    let repeat = config.repeat.unwrap_or_else(|| default_repeat(case.kind));
    let expression = Expression::parse(&case.source)
        .map_err(|error| format!("{} parse preflight failed: {error}", case.name))?;
    match case.kind {
        Kind::Scalar => scalar_preflight(case, &expression)?,
        Kind::Array { .. } => array_preflight(case, &expression)?,
    }

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
