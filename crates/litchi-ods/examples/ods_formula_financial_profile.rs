//! Process-level financial evaluator profile.
//!
//! This example is deliberately self-contained so the performance runner can
//! copy it into the immutable baseline checkout.  It exercises only public
//! ODS formula APIs, keeps the resolver deterministic, and reports the same
//! prepared outcome/checksum/read contract in preflight and timed paths.

#![allow(clippy::cast_precision_loss)]
#![allow(clippy::missing_panics_doc)]

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
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        UnsupportedKind, evaluate_scalar,
        value::{self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value},
    },
    expression::Expression,
};
use serde_json::{Value as JsonValue, json};

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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Path {
    Scalar,
    Value,
}

impl Path {
    const fn label(self) -> &'static str {
        match self {
            Self::Scalar => "scalar",
            Self::Value => "value",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Shape {
    Scalar,
    Array { rows: usize, columns: usize },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FailureKind {
    ReferenceCells,
    Work,
    Cancelled,
    Unsupported,
}

#[derive(Clone, Debug, PartialEq)]
enum Expected {
    Number(f64),
    Array {
        rows: usize,
        columns: usize,
        values: Vec<f64>,
    },
    Error(ScalarError),
    Failure(FailureKind),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CancelMode {
    None,
    Before,
    AfterRead,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ResolverMode {
    Normal,
    FormulaErrorAtFourthCell,
    FailAfterFirstRead,
    FailOnDateAfterValues,
}

#[derive(Clone, Debug)]
struct Case {
    name: String,
    source: String,
    path: Path,
    shape: Shape,
    expected: Expected,
    reference_reads: u64,
    max_reference_cells: Option<usize>,
    max_steps: Option<u64>,
    cancel: CancelMode,
    resolver: ResolverMode,
    repeat: usize,
    class: String,
}

#[derive(Debug)]
struct ResolverStats {
    reads: AtomicU64,
}

impl ResolverStats {
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
        Ok((matches!(sheet, "Main" | "Data" | "Archive")).then_some(SheetExtent::new(8192, 8)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        execution.check()?;
        let read = self.stats.reads.load(Ordering::Relaxed) + 1;
        if self.mode == ResolverMode::FailAfterFirstRead && read > 1 {
            return Err(EvaluationFailure::Unsupported(UnsupportedKind::CellValue));
        }
        if self.mode == ResolverMode::FailOnDateAfterValues && column == 1 && row == 0 {
            return Err(EvaluationFailure::Unsupported(UnsupportedKind::CellValue));
        }
        self.stats.reads.fetch_add(1, Ordering::Relaxed);
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        if !matches!(sheet, "Main" | "Data" | "Archive") || row >= 8192 || column >= 8 {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if self.mode == ResolverMode::FormulaErrorAtFourthCell && row == 3 && column == 0 {
            return Ok(CellRead::Error(ScalarError::NotAvailable));
        }
        Ok(match column {
            // Cash flows: a canonical two-period investment followed by zeroes.
            0 if row == 0 => CellRead::Number(-100.0),
            0 if row == 1 => CellRead::Number(110.0),
            0 => CellRead::Number(0.0),
            // 365-day serials exercise XIRR/XNPV source ordering.
            1 => CellRead::Number(43_831.0 + row as f64 * 365.0),
            // Small schedule rates keep the large product finite.
            2 => CellRead::Number(0.0001 + row as f64 * 0.000001),
            // Unchanged aggregate control data.
            3 => CellRead::Number(1.0 + (row % 17) as f64 * 0.25),
            _ => CellRead::Empty,
        })
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

fn reset_allocator() -> u64 {
    ALLOC_CALLS.store(0, Ordering::Relaxed);
    DEALLOC_CALLS.store(0, Ordering::Relaxed);
    ALLOC_BYTES.store(0, Ordering::Relaxed);
    DEALLOC_BYTES.store(0, Ordering::Relaxed);
    let live = LIVE_BYTES.load(Ordering::Acquire);
    PEAK_LIVE_BYTES.store(live, Ordering::Relaxed);
    live
}

fn execution() -> (ExecutionContext, CancellationSource) {
    let budget = Budget::root(
        "ods-formula-financial-performance",
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

fn approx(actual: f64, expected: f64) -> bool {
    (actual - expected).abs() <= expected.abs().max(1.0) * 1.0e-10
}

fn check_scalar(expected: &Expected, value: &ScalarValue<'_>, name: &str) -> AnyResult<()> {
    match (expected, value) {
        (Expected::Number(want), ScalarValue::Number(got)) if approx(*got, *want) => Ok(()),
        (Expected::Error(want), ScalarValue::Error(got)) if want == got => Ok(()),
        (Expected::Number(want), got) => {
            Err(format!("{name}: expected Number {want:?}, got {got:?}").into())
        },
        (Expected::Error(want), got) => {
            Err(format!("{name}: expected formula error {want:?}, got {got:?}").into())
        },
        (Expected::Array { .. } | Expected::Failure(_), got) => {
            Err(format!("{name}: scalar result does not match {got:?}").into())
        },
    }
}

fn check_value(expected: &Expected, value: Value<'_>, name: &str) -> AnyResult<()> {
    match (expected, value) {
        (Expected::Number(want), Value::Number(got)) if approx(got, *want) => Ok(()),
        (Expected::Number(want), got) => {
            Err(format!("{name}: expected scalar Number {want:?}, got {got:?}").into())
        },
        (Expected::Error(want), Value::Error(got)) if want == &got => Ok(()),
        (Expected::Error(want), got) => {
            Err(format!("{name}: expected formula error {want:?}, got {got:?}").into())
        },
        (
            Expected::Array {
                rows,
                columns,
                values,
            },
            Value::Array(array),
        ) => {
            let shape = array.shape();
            if shape.rows() != *rows || shape.columns() != *columns || array.len() != values.len() {
                return Err(format!("{name}: array shape mismatch").into());
            }
            for (index, want) in values.iter().enumerate() {
                let Some(Value::Number(got)) = array.get(index) else {
                    return Err(format!("{name}: array cell {index} is not Number").into());
                };
                if !approx(got, *want) {
                    return Err(format!("{name}: array cell {index}: {got} != {want}").into());
                }
            }
            Ok(())
        },
        (Expected::Array { .. }, got) => Err(format!("{name}: expected array, got {got:?}").into()),
        (Expected::Failure(_), got) => Err(format!("{name}: failure expected, got {got:?}").into()),
    }
}

fn check_failure(expected: &Expected, error: &EvaluationFailure, name: &str) -> AnyResult<()> {
    let Expected::Failure(kind) = expected else {
        return Err(format!("{name}: unexpected evaluator failure {error}").into());
    };
    let matches = match kind {
        FailureKind::ReferenceCells => {
            matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects)
        },
        FailureKind::Work => {
            matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work)
        },
        FailureKind::Cancelled => matches!(error, EvaluationFailure::Cancelled),
        FailureKind::Unsupported => matches!(error, EvaluationFailure::Unsupported(_)),
    };
    if matches {
        Ok(())
    } else {
        Err(format!("{name}: expected {kind:?}, got {error}").into())
    }
}

fn checksum_scalar(value: &ScalarValue<'_>) -> u64 {
    match value {
        ScalarValue::Number(value) => value.to_bits().rotate_left(11),
        ScalarValue::Logical(value) => u64::from(*value).rotate_left(11),
        ScalarValue::Text(value) => text_checksum(value.as_ref()),
        ScalarValue::Error(error) => text_checksum(&error.to_string()),
        ScalarValue::Complex(value) => text_checksum(&format!("{value:?}")),
        _ => 0,
    }
}

fn checksum_value(value: Value<'_>) -> AnyResult<u64> {
    Ok(match value {
        Value::Number(value) => value.to_bits().rotate_left(11),
        Value::Logical(value) => u64::from(value).rotate_left(11),
        Value::Text(value) => text_checksum(value),
        Value::Error(error) => text_checksum(&error.to_string()),
        Value::Array(array) => {
            let mut checksum =
                ((array.shape().rows() as u64) << 32) ^ array.shape().columns() as u64;
            for index in 0..array.len() {
                checksum = checksum.rotate_left(5)
                    ^ checksum_value(array.get(index).ok_or("missing array cell")?)?;
            }
            checksum
        },
        other => return Err(format!("unexpected value {other:?}").into()),
    })
}

fn text_checksum(value: &str) -> u64 {
    value.bytes().fold(0xcbf29ce484222325, |hash, byte| {
        hash.wrapping_mul(0x100000001b3)
            .wrapping_add(u64::from(byte))
    })
}

#[derive(Clone, Debug)]
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

fn expected_json(expected: &Expected) -> JsonValue {
    match expected {
        Expected::Number(value) => json!(value),
        Expected::Array {
            rows,
            columns,
            values,
        } => json!({ "kind": "array", "rows": rows, "columns": columns, "values": values }),
        Expected::Error(error) => json!({ "kind": "formula_error", "value": error.to_string() }),
        Expected::Failure(kind) => json!({ "kind": "failure", "value": format!("{kind:?}") }),
    }
}

fn description(case: &Case, reads: Option<u64>, supported: bool) -> JsonValue {
    let shape = match case.shape {
        Shape::Scalar => json!("scalar"),
        Shape::Array { rows, columns } => json!(format!("{rows}x{columns}")),
    };
    json!({
        "case": case.name.as_str(),
        "source": case.source.as_str(),
        "evaluation_path": case.path.label(),
        "shape": shape,
        "expected": expected_json(&case.expected),
        "reference_reads": reads,
        "max_steps": case.max_steps,
        "repeat": case.repeat,
        "class": case.class.as_str(),
        "supported": supported,
    })
}

fn sample_json(sample: &Sample) -> JsonValue {
    json!({
        "elapsed_ns": sample.elapsed_ns,
        "alloc_calls": sample.alloc_calls,
        "dealloc_calls": sample.dealloc_calls,
        "requested_bytes": sample.requested_bytes,
        "released_bytes": sample.released_bytes,
        "live_before": sample.live_before,
        "live_after": sample.live_after,
        "peak_live_delta": sample.peak_live_delta,
        "work": sample.work,
        "memory_retained": sample.memory_retained,
        "reference_reads": sample.reference_reads,
        "output_bytes": sample.output_bytes,
        "checksum": sample.checksum,
    })
}

fn evaluate_once(
    case: &Case,
    expression: &Expression,
    execution: &ExecutionContext,
    cancellation: &CancellationSource,
    resolver: &FixtureResolver,
) -> AnyResult<(u64, u64, u64)> {
    if case.cancel == CancelMode::Before {
        cancellation.cancel();
    }
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let (checksum, output_bytes) = match case.path {
        Path::Scalar => match evaluate_scalar(
            expression,
            &EvaluationContext::new(execution),
            &scalar_limits(case),
        ) {
            Ok(result) => {
                if matches!(&case.expected, Expected::Failure(_)) {
                    return Err(format!("{} unexpectedly succeeded", case.name).into());
                }
                check_scalar(&case.expected, result.value(), &case.name)?;
                (checksum_scalar(result.value()), 0)
            },
            Err(error) => {
                check_failure(&case.expected, &error, &case.name)?;
                (failure_checksum(&case.expected), 0)
            },
        },
        Path::Value => {
            let context =
                Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
            match value::evaluate(expression, resolver, &context, &value_limits(case)) {
                Ok(result) => {
                    if matches!(&case.expected, Expected::Failure(_)) {
                        return Err(format!("{} unexpectedly succeeded", case.name).into());
                    }
                    let value = result.value();
                    check_value(&case.expected, value, &case.name)?;
                    let output = output_bytes(value);
                    (checksum_value(value)?, output)
                },
                Err(error) => {
                    check_failure(&case.expected, &error, &case.name)?;
                    (failure_checksum(&case.expected), 0)
                },
            }
        },
    };
    let work = execution
        .budget()
        .used(Resource::Work)
        .saturating_sub(baseline_work);
    let retained = execution
        .budget()
        .used(Resource::Memory)
        .saturating_sub(baseline_memory);
    Ok((checksum, output_bytes, work.saturating_add(retained)))
}

fn failure_checksum(expected: &Expected) -> u64 {
    let text = format!("{expected:?}");
    text_checksum(&text)
}

fn output_bytes(value: Value<'_>) -> u64 {
    match value {
        Value::Text(text) => text.len() as u64,
        Value::Array(array) => (0..array.len())
            .filter_map(|index| array.get(index))
            .map(output_bytes)
            .sum(),
        _ => 0,
    }
}

fn measure(
    case: &Case,
    expression: Option<&Expression>,
    phase: Phase,
    repeat: usize,
) -> AnyResult<Sample> {
    let (execution, cancellation) = execution();
    let resolver = FixtureResolver::new(
        case.resolver,
        (case.cancel == CancelMode::AfterRead).then_some(cancellation.clone()),
    );
    let live_before = reset_allocator();
    let baseline_work = execution.budget().used(Resource::Work);
    let baseline_memory = execution.budget().used(Resource::Memory);
    let started = Instant::now();
    let mut checksum = 0_u64;
    let mut output = 0_u64;
    for _ in 0..repeat {
        let parsed;
        let expression = match phase {
            Phase::Evaluate => expression.expect("pre-parsed expression"),
            Phase::ParseEvaluate => {
                parsed = Expression::parse(&case.source)
                    .map_err(|error| format!("{} parse failed: {error}", case.name))?;
                &parsed
            },
        };
        let (part_checksum, part_output, _) =
            evaluate_once(case, expression, &execution, &cancellation, &resolver)?;
        checksum = checksum.wrapping_add(part_checksum);
        output = output.saturating_add(part_output);
        black_box(checksum);
    }
    let retained = execution
        .budget()
        .used(Resource::Memory)
        .saturating_sub(baseline_memory);
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
        memory_retained: retained,
        reference_reads: resolver.stats.reads(),
        output_bytes: output,
        checksum,
    })
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
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
    allow_unsupported: bool,
}

fn parse_config() -> AnyResult<Option<Config>> {
    let mut case = None;
    let mut phase = Phase::Evaluate;
    let mut warmups = 3;
    let mut iterations = 1;
    let mut repeat = None;
    let mut preflight_only = false;
    let mut describe = false;
    let mut allow_unsupported = false;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                println!(
                    "usage: ods_formula_financial_profile [--case NAME|all] [--phase evaluate|parse-evaluate] [--warmups N] [--iterations N] [--repeat N] [--preflight-only] [--describe] [--allow-unsupported]"
                );
                return Ok(None);
            },
            "--list" => {
                for row in all_cases() {
                    println!("{}", row.name);
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
            "--allow-unsupported" => allow_unsupported = true,
            other => return Err(format!("unknown option {other:?}").into()),
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
        allow_unsupported,
    }))
}

fn case(
    name: impl Into<String>,
    source: impl Into<String>,
    path: Path,
    shape: Shape,
    expected: Expected,
    reads: u64,
    class: impl Into<String>,
) -> Case {
    Case {
        name: name.into(),
        source: source.into(),
        path,
        shape,
        expected,
        reference_reads: reads,
        max_reference_cells: None,
        max_steps: None,
        cancel: CancelMode::None,
        resolver: ResolverMode::Normal,
        repeat: 32,
        class: class.into(),
    }
}

fn scalar(name: &str, source: &str, expected: f64) -> Case {
    case(
        name,
        source,
        Path::Scalar,
        Shape::Scalar,
        Expected::Number(expected),
        0,
        "financial",
    )
}

fn scalar_error(name: &str, source: &str, error: ScalarError, class: &str) -> Case {
    case(
        name,
        source,
        Path::Scalar,
        Shape::Scalar,
        Expected::Error(error),
        0,
        class,
    )
}

fn array_case(name: &str, source: &str, values: Vec<f64>, rows: usize, columns: usize) -> Case {
    case(
        name,
        source,
        Path::Value,
        Shape::Array { rows, columns },
        Expected::Array {
            rows,
            columns,
            values,
        },
        0,
        "matched-control",
    )
}

fn schedule_product(n: usize) -> f64 {
    (0..n).fold(100.0, |value, row| {
        value * (1.0 + 0.0001 + row as f64 * 0.000001)
    })
}

fn npv_value(n: usize) -> f64 {
    (0..n)
        .map(|row| {
            let cash = match row {
                0 => -100.0,
                1 => 110.0,
                _ => 0.0,
            };
            cash / 1.1_f64.powi((row + 1) as i32)
        })
        .sum()
}

fn conditional_sum_value(n: usize) -> f64 {
    (0..n)
        .filter_map(|row| {
            let value = 1.0 + (row % 17) as f64 * 0.25;
            (value >= 2.0).then_some(value)
        })
        .sum()
}

fn annuity_payment(rate: f64, periods: usize, principal: f64) -> f64 {
    -principal * rate / (1.0 - (1.0 + rate).powi(-(periods as i32)))
}

fn cumulative_interest(rate: f64, periods: usize, end: usize, principal: f64) -> f64 {
    let payment = annuity_payment(rate, periods, principal);
    (1..=end)
        .map(|period| {
            let remaining = periods - period + 1;
            let balance = -payment * (1.0 - (1.0 + rate).powi(-(remaining as i32))) / rate;
            -balance * rate
        })
        .sum()
}

fn cumulative_principal(rate: f64, periods: usize, end: usize, principal: f64) -> f64 {
    let payment = annuity_payment(rate, periods, principal);
    end as f64 * payment - cumulative_interest(rate, periods, end, principal)
}

fn mirr_value(n: usize) -> f64 {
    let positive_npv = 110.0 / 1.1_f64.powi(n as i32);
    let negative_npv = -100.0 / 1.1;
    ((-positive_npv * 1.1_f64.powi(n as i32)) / (negative_npv * 1.1)).powf(1.0 / (n as f64 - 1.0))
        - 1.0
}

fn inline_mirr(n: usize) -> String {
    let mut values = Vec::with_capacity(n);
    values.push("-100".to_owned());
    for _ in 2..n {
        values.push("0".to_owned());
    }
    values.push("110".to_owned());
    format!("=MIRR({{{}}};0.1;0.1)", values.join("|"))
}

fn all_cases() -> Vec<Case> {
    let mut rows = vec![
        scalar("control-scalar-arithmetic", "=0.17+1.25", 1.42),
        scalar("control-scalar-sin", "=SIN(0.17)", 0.17_f64.sin()),
        array_case(
            "control-array-arithmetic",
            "={1|2|3|4}+1",
            vec![2.0, 3.0, 4.0, 5.0],
            4,
            1,
        ),
        array_case(
            "control-array-sin",
            "=SIN({0.1|0.2|0.3|0.4})",
            vec![0.1_f64.sin(), 0.2_f64.sin(), 0.3_f64.sin(), 0.4_f64.sin()],
            4,
            1,
        ),
        case(
            "control-scalar-aggregate",
            "=SUM(1.25)",
            Path::Scalar,
            Shape::Scalar,
            Expected::Number(1.25),
            0,
            "matched-control",
        ),
        case(
            "control-reference-sum-64",
            "=SUM([.D1:.D64])",
            Path::Value,
            Shape::Scalar,
            Expected::Number((0..64).map(|row| 1.0 + (row % 17) as f64 * 0.25).sum()),
            64,
            "matched-control",
        ),
        case(
            "control-reference-average-64",
            "=AVERAGE([.D1:.D64])",
            Path::Value,
            Shape::Scalar,
            Expected::Number(
                (0..64)
                    .map(|row| 1.0 + (row % 17) as f64 * 0.25)
                    .sum::<f64>()
                    / 64.0,
            ),
            64,
            "matched-control",
        ),
        case(
            "control-reference-conditional-sumif-64",
            "=SUMIF([.D1:.D64];\">=2\";[.D1:.D64])",
            Path::Value,
            Shape::Scalar,
            Expected::Number(conditional_sum_value(64)),
            112,
            "matched-control",
        ),
        array_case(
            "control-lazy-if",
            "=IF({FALSE()|FALSE()};SUM([.D1:.D64]);0)",
            vec![0.0, 0.0],
            2,
            1,
        ),
        case(
            "control-iferror",
            "=IFERROR(1/0;42)",
            Path::Scalar,
            Shape::Scalar,
            Expected::Number(42.0),
            0,
            "matched-control",
        ),
        case(
            "control-ifna",
            "=IFNA(#N/A;42)",
            Path::Scalar,
            Shape::Scalar,
            Expected::Number(42.0),
            0,
            "matched-control",
        ),
        scalar("financial-cumipmt-core", "=CUMIPMT(0.1;2;100;1;1;0)", -10.0),
        scalar(
            "financial-cumprinc-core",
            "=CUMPRINC(0.1;2;100;1;1;0)",
            -47.61904761904762,
        ),
        scalar("financial-effect-core", "=EFFECT(0.1;2)", 0.1025),
        scalar("financial-fv-core", "=FV(0.1;1;-110)", 110.0),
        case(
            "financial-fvschedule-core",
            "=FVSCHEDULE(100;{0.1|0.2})",
            Path::Value,
            Shape::Scalar,
            Expected::Number(132.0),
            0,
            "financial",
        ),
        scalar("financial-ipmt-core", "=IPMT(0.1;1;2;100)", -10.0),
        case(
            "financial-irr-core",
            "=IRR({-100|110})",
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.1),
            0,
            "financial",
        ),
        scalar("financial-ispmt-core", "=ISPMT(0.1;1;2;100)", -5.0),
        case(
            "financial-mirr-core",
            "=MIRR({-100|110};0.1;0.1)",
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.1),
            0,
            "financial",
        ),
        scalar("financial-nominal-core", "=NOMINAL(0.1025;2)", 0.1),
        scalar("financial-nper-core", "=NPER(0.1;-110;100)", 1.0),
        case(
            "financial-npv-core",
            "=NPV(0.1;{-100|110})",
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.0),
            0,
            "financial",
        ),
        scalar("financial-pduration-core", "=PDURATION(0.1;100;121)", 2.0),
        scalar("financial-pmt-core", "=PMT(0.1;1;100)", -110.0),
        scalar(
            "financial-ppmt-core",
            "=PPMT(0.1;1;2;100)",
            -47.61904761904762,
        ),
        scalar("financial-pv-core", "=PV(0.1;1;-110)", 100.0),
        scalar("financial-rate-core", "=RATE(1;-110;100)", 0.1),
        scalar("financial-rri-core", "=RRI(2;100;121)", 0.1),
        case(
            "financial-xirr-core",
            "=XIRR({-100|110};{43831|44196})",
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.1),
            0,
            "financial",
        ),
        case(
            "financial-xnpv-core",
            "=XNPV(0.1;{-100|110};{43831|44196})",
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.0),
            0,
            "financial",
        ),
    ];

    for periods in [4, 64, 512, 2048] {
        let suffix = periods.to_string();
        rows.push(scalar(
            &format!("financial-cumipmt-span-{suffix}"),
            &format!("=CUMIPMT(0.1;{periods};100;1;{};0)", periods - 1),
            cumulative_interest(0.1, periods, periods - 1, 100.0),
        ));
        rows.push(scalar(
            &format!("financial-cumprinc-span-{suffix}"),
            &format!("=CUMPRINC(0.1;{periods};100;1;{};0)", periods - 1),
            cumulative_principal(0.1, periods, periods - 1, 100.0),
        ));
    }

    rows.extend([
        scalar("financial-effect-zero", "=EFFECT(0;2)", 0.0),
        scalar("financial-fv-zero-rate", "=FV(0;2;-10)", 20.0),
        scalar("financial-pmt-zero-rate", "=PMT(0;2;100)", -50.0),
        scalar("financial-pv-zero-rate", "=PV(0;2;-50)", 100.0),
        scalar("financial-nper-zero-rate", "=NPER(0;-50;100)", 2.0),
        scalar("financial-rri-negative-inputs", "=RRI(2;-100;-121)", 0.1),
        case(
            "financial-fvschedule-zero-factor",
            "=FVSCHEDULE(100;{0|-1|0.1})",
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.0),
            0,
            "financial-sequence",
        ),
        case(
            "financial-fvschedule-formula-error",
            "=FVSCHEDULE(100;{0.1|#N/A|0.2})",
            Path::Value,
            Shape::Scalar,
            Expected::Error(ScalarError::NotAvailable),
            0,
            "financial-error",
        ),
        case(
            "financial-npv-split-arguments",
            "=NPV(0.1;{-100};{110})",
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.0),
            0,
            "financial-sequence",
        ),
        case(
            "financial-npv-reference-list",
            "=NPV(0.1;[.A1:.A1]~[.A2:.A2])",
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.0),
            2,
            "financial-sequence",
        ),
        case(
            "financial-npv-projected-if",
            "=IF({TRUE()|FALSE()};NPV(0.1;[.A1:.A4]);0)",
            Path::Value,
            Shape::Array {
                rows: 2,
                columns: 1,
            },
            Expected::Array {
                rows: 2,
                columns: 1,
                values: vec![npv_value(4), 0.0],
            },
            4,
            "financial-projected",
        ),
        case(
            "financial-irr-negative-branch",
            "=IRR({-100|0|100};-2)",
            Path::Value,
            Shape::Scalar,
            Expected::Number(-2.0),
            0,
            "financial-solver",
        ),
        case(
            "financial-irr-no-sign-variation",
            "=IRR({1|2})",
            Path::Value,
            Shape::Scalar,
            Expected::Error(ScalarError::Number),
            0,
            "financial-error",
        ),
        case(
            "financial-xirr-nonannual",
            "=XIRR({-100|110};{43831|44011})",
            Path::Value,
            Shape::Scalar,
            Expected::Number(1.1_f64.powf(365.0 / 180.0) - 1.0),
            0,
            "financial-solver",
        ),
        case(
            "financial-xnpv-nonannual",
            "=XNPV(0.1;{-100|110};{43831|44011})",
            Path::Value,
            Shape::Scalar,
            Expected::Number(-100.0 + 110.0 / 1.1_f64.powf(180.0 / 365.0)),
            0,
            "financial-stream",
        ),
        case(
            "financial-fvschedule-reference-list-refusal",
            "=FVSCHEDULE(100;[.C1:.C2]~[.C3:.C4])",
            Path::Value,
            Shape::Scalar,
            Expected::Error(ScalarError::Value),
            0,
            "financial-refusal",
        ),
        case(
            "financial-xirr-reference-list-refusal",
            "=XIRR([.A1:.A2]~[.A3:.A4];{43831|44196})",
            Path::Value,
            Shape::Scalar,
            Expected::Error(ScalarError::Value),
            0,
            "financial-refusal",
        ),
        case(
            "financial-xnpv-reference-list-refusal",
            "=XNPV(0.1;[.A1:.A2]~[.A3:.A4];[.B1:.B4])",
            Path::Value,
            Shape::Scalar,
            Expected::Error(ScalarError::Value),
            0,
            "financial-refusal",
        ),
        scalar_error(
            "financial-effect-invalid-rate",
            "=EFFECT(-0.1;2)",
            ScalarError::Number,
            "financial-error",
        ),
    ]);

    rows.extend([
        scalar_error(
            "financial-nominal-invalid-rate",
            "=NOMINAL(0;2)",
            ScalarError::Number,
            "financial-error",
        ),
        scalar_error(
            "financial-pduration-invalid-rate",
            "=PDURATION(0;100;121)",
            ScalarError::Number,
            "financial-error",
        ),
        scalar_error(
            "financial-ppmt-invalid-period",
            "=PPMT(0.1;2;2;100)",
            ScalarError::Number,
            "financial-error",
        ),
        scalar_error(
            "financial-rri-nonintegral-negative",
            "=RRI(3;100;-800)",
            ScalarError::Number,
            "financial-error",
        ),
    ]);

    for n in [4, 64, 1024, 4096] {
        let suffix = n.to_string();
        let mut schedule = case(
            format!("financial-fvschedule-{suffix}"),
            format!("=FVSCHEDULE(100;[.C1:.C{suffix}])"),
            Path::Value,
            Shape::Scalar,
            Expected::Number(schedule_product(n)),
            n as u64,
            "financial-stream",
        );
        schedule.repeat = if n >= 1024 { 1 } else { 8 };
        rows.push(schedule);

        let mut npv = case(
            format!("financial-npv-{suffix}"),
            format!("=NPV(0.1;[.A1:.A{suffix}])"),
            Path::Value,
            Shape::Scalar,
            Expected::Number(npv_value(n)),
            n as u64,
            "financial-stream",
        );
        if n >= 1024 {
            npv.max_steps = Some(8_000_000);
        }
        npv.repeat = if n >= 1024 { 1 } else { 8 };
        rows.push(npv);
    }
    for n in [4, 64, 1024] {
        let suffix = n.to_string();
        let mut irr = case(
            format!("financial-irr-{suffix}"),
            format!("=IRR([.A1:.A{suffix}])"),
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.1),
            n as u64,
            "financial-solver",
        );
        irr.repeat = if n >= 1024 { 1 } else { 4 };
        rows.push(irr);
        let mut xirr = case(
            format!("financial-xirr-{suffix}"),
            format!("=XIRR([.A1:.A{suffix}];[.B1:.B{suffix}])"),
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.1),
            (n * 2) as u64,
            "financial-solver",
        );
        xirr.repeat = if n >= 1024 { 1 } else { 4 };
        rows.push(xirr);
        let mut xnpv = case(
            format!("financial-xnpv-{suffix}"),
            format!("=XNPV(0.1;[.A1:.A{suffix}];[.B1:.B{suffix}])"),
            Path::Value,
            Shape::Scalar,
            Expected::Number(0.0),
            (n * 2) as u64,
            "financial-stream",
        );
        if n >= 1024 {
            xnpv.max_steps = Some(8_000_000);
        }
        xnpv.repeat = if n >= 1024 { 1 } else { 4 };
        rows.push(xnpv);
        let mut mirr = case(
            format!("financial-mirr-{suffix}"),
            inline_mirr(n),
            Path::Value,
            Shape::Scalar,
            Expected::Number(mirr_value(n)),
            0,
            "financial-array",
        );
        if n >= 1024 {
            mirr.max_steps = Some(8_000_000);
        }
        mirr.repeat = if n >= 1024 { 1 } else { 4 };
        rows.push(mirr);
    }

    let mut refusal = case(
        "financial-irr-reference-list-refusal",
        "=IRR([.A1:.A2]~[.A3:.A4])",
        Path::Value,
        Shape::Scalar,
        Expected::Error(ScalarError::Value),
        0,
        "financial-refusal",
    );
    refusal.repeat = 1;
    rows.push(refusal);

    let mut reference_limit = case(
        "financial-npv-reference-limit",
        "=NPV(0.1;[.A1:.A4])",
        Path::Value,
        Shape::Scalar,
        Expected::Failure(FailureKind::ReferenceCells),
        0,
        "financial-resource",
    );
    reference_limit.max_reference_cells = Some(3);
    reference_limit.repeat = 1;
    rows.push(reference_limit);

    let mut work_limit = case(
        "financial-irr-work-limit",
        "=IRR({-100|110})",
        Path::Value,
        Shape::Scalar,
        Expected::Failure(FailureKind::Work),
        0,
        "financial-resource",
    );
    work_limit.max_steps = Some(2);
    work_limit.repeat = 1;
    rows.push(work_limit);

    let mut cancelled = case(
        "financial-npv-cancel-after-read",
        "=NPV(0.1;[.A1:.A4])",
        Path::Value,
        Shape::Scalar,
        Expected::Failure(FailureKind::Cancelled),
        1,
        "financial-cancellation",
    );
    cancelled.cancel = CancelMode::AfterRead;
    cancelled.repeat = 4;
    rows.push(cancelled);

    let mut cancel_before = case(
        "financial-npv-cancel-before-read",
        "=NPV(0.1;[.A1:.A4])",
        Path::Value,
        Shape::Scalar,
        Expected::Failure(FailureKind::Cancelled),
        0,
        "financial-cancellation",
    );
    cancel_before.cancel = CancelMode::Before;
    cancel_before.repeat = 4;
    rows.push(cancel_before);

    let mut formula_error = case(
        "financial-npv-retained-formula-error",
        "=NPV(0.1;[.A1:.A4])",
        Path::Value,
        Shape::Scalar,
        Expected::Error(ScalarError::NotAvailable),
        4,
        "financial-error",
    );
    formula_error.resolver = ResolverMode::FormulaErrorAtFourthCell;
    formula_error.repeat = 1;
    rows.push(formula_error);

    let mut provider_failure = case(
        "financial-npv-late-provider-failure",
        "=NPV(0.1;[.A1:.A4])",
        Path::Value,
        Shape::Scalar,
        Expected::Failure(FailureKind::Unsupported),
        1,
        "financial-error",
    );
    provider_failure.resolver = ResolverMode::FailAfterFirstRead;
    provider_failure.repeat = 1;
    rows.push(provider_failure);

    let mut root_cancelled = case(
        "financial-irr-cancel-during-scan",
        "=IRR([.A1:.A4])",
        Path::Value,
        Shape::Scalar,
        Expected::Failure(FailureKind::Cancelled),
        1,
        "financial-cancellation",
    );
    root_cancelled.cancel = CancelMode::AfterRead;
    root_cancelled.repeat = 1;
    rows.push(root_cancelled);

    let mut date_provider_failure = case(
        "financial-xnpv-late-date-provider-failure",
        "=XNPV(0.1;[.A1:.A4];[.B1:.B4])",
        Path::Value,
        Shape::Scalar,
        Expected::Failure(FailureKind::Unsupported),
        4,
        "financial-error",
    );
    date_provider_failure.resolver = ResolverMode::FailOnDateAfterValues;
    date_provider_failure.repeat = 1;
    rows.push(date_provider_failure);

    rows
}

fn selected_cases<'a>(
    all: &'a mut [Case],
    requested: Option<&str>,
) -> AnyResult<Vec<&'a mut Case>> {
    match requested {
        None | Some("all") => Ok(all.iter_mut().collect()),
        Some(name) => all
            .iter_mut()
            .find(|row| row.name == name)
            .map(|row| vec![row])
            .ok_or_else(|| format!("unknown case {name}").into()),
    }
}

fn preflight(case: &Case, allow_unsupported: bool) -> AnyResult<(bool, u64, JsonValue)> {
    let expression = Expression::parse(&case.source)
        .map_err(|error| format!("{} parse preflight failed: {error}", case.name))?;
    let (execution, cancellation) = execution();
    let resolver = FixtureResolver::new(
        case.resolver,
        (case.cancel == CancelMode::AfterRead).then_some(cancellation.clone()),
    );
    if case.cancel == CancelMode::Before {
        cancellation.cancel();
    }
    let result = match case.path {
        Path::Scalar => evaluate_scalar(
            &expression,
            &EvaluationContext::new(&execution),
            &scalar_limits(case),
        )
        .map(|result| {
            check_scalar(&case.expected, result.value(), &case.name)?;
            Ok::<_, Box<dyn Error>>(())
        }),
        Path::Value => value::evaluate(
            &expression,
            &resolver,
            &Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix),
            &value_limits(case),
        )
        .map(|result| {
            check_value(&case.expected, result.value(), &case.name)?;
            Ok::<_, Box<dyn Error>>(())
        }),
    };
    match result {
        Ok(Ok(())) => {
            let reads = resolver.stats.reads();
            if reads != case.reference_reads {
                return Err(format!(
                    "{}: expected {} reads, observed {}",
                    case.name, case.reference_reads, reads
                )
                .into());
            }
            Ok((true, reads, description(case, Some(reads), true)))
        },
        Ok(Err(error)) => Err(error),
        Err(error) if allow_unsupported && matches!(error, EvaluationFailure::Unsupported(_)) => {
            let reads = resolver.stats.reads();
            Ok((false, reads, description(case, Some(reads), false)))
        },
        Err(error) => {
            if matches!(&case.expected, Expected::Failure(_)) {
                check_failure(&case.expected, &error, &case.name)?;
                let reads = resolver.stats.reads();
                if reads != case.reference_reads {
                    return Err(format!(
                        "{}: expected {} reads, observed {}",
                        case.name, case.reference_reads, reads
                    )
                    .into());
                }
                Ok((true, reads, description(case, Some(reads), true)))
            } else {
                Err(error.into())
            }
        },
    }
}

fn run_case(case: &Case, config: &Config) -> AnyResult<()> {
    let (supported, reads, description) = preflight(case, config.allow_unsupported)?;
    if config.describe {
        println!("{}", description);
        return Ok(());
    }
    println!("preflight-json {description}");
    println!(
        "preflight case={} reference_reads={reads} supported={supported}",
        case.name
    );
    if config.preflight_only || !supported {
        return Ok(());
    }
    let repeat = config.repeat.unwrap_or(case.repeat);
    let expression = Expression::parse(&case.source)?;
    for _ in 0..config.warmups {
        let _ = measure(case, Some(&expression), config.phase, repeat)?;
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(measure(case, Some(&expression), config.phase, repeat)?);
    }
    // `--repeat` is a capture override, so keep the value in the emitted
    // record even though the immutable matrix retains the default repeat.
    let records = samples.iter().map(sample_json).collect::<Vec<_>>();
    println!(
        "{}",
        json!({
            "case": case.name,
            "phase": config.phase.label(),
            "warmups": config.warmups,
            "iterations": config.iterations,
            "repeat": repeat,
            "shape": match case.shape { Shape::Scalar => "scalar".to_owned(), Shape::Array { rows, columns } => format!("{rows}x{columns}") },
            "expected": expected_json(&case.expected),
            "supported": true,
            "samples": records,
        })
    );
    Ok(())
}

fn main() -> AnyResult<()> {
    let Some(config) = parse_config()? else {
        return Ok(());
    };
    let mut all = all_cases();
    let selected = selected_cases(&mut all, config.case.as_deref())?;
    for row in selected {
        run_case(row, &config)?;
    }
    Ok(())
}
