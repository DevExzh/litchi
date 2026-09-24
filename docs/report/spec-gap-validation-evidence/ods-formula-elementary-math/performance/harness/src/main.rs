//! Bounded ODS OpenFormula elementary-math profile.
//!
//! The binary is deliberately separate from production crates.  It runs one
//! immutable in-memory fixture through the scalar and value evaluator entry
//! points, checks a numerical oracle before timing, and reports process-local
//! allocation observations.  A baseline may opt into recording the expected
//! typed refusal for the newly implemented elementary-math functions; those
//! rows are capability evidence and are not compared as numeric timings.

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
        // SAFETY: forwarded unchanged to the platform allocator.
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
        // SAFETY: forwarded unchanged to the platform allocator.
        unsafe { System.realloc(pointer, layout, size) }
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[allow(dead_code)]
enum TrigOp {
    Acos,
    Acosh,
    Acot,
    Acoth,
    Asin,
    Asinh,
    Atan,
    Atan2,
    Atanh,
    Cos,
    Cosh,
    Cot,
    Coth,
    Csc,
    Csch,
    Degrees,
    Pi,
    Radians,
    Sec,
    Sech,
    Sin,
    Sinh,
    Tan,
    Tanh,
}

impl TrigOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Acos => "ACOS",
            Self::Acosh => "ACOSH",
            Self::Acot => "ACOT",
            Self::Acoth => "ACOTH",
            Self::Asin => "ASIN",
            Self::Asinh => "ASINH",
            Self::Atan => "ATAN",
            Self::Atan2 => "ATAN2",
            Self::Atanh => "ATANH",
            Self::Cos => "COS",
            Self::Cosh => "COSH",
            Self::Cot => "COT",
            Self::Coth => "COTH",
            Self::Csc => "CSC",
            Self::Csch => "CSCH",
            Self::Degrees => "DEGREES",
            Self::Pi => "PI",
            Self::Radians => "RADIANS",
            Self::Sec => "SEC",
            Self::Sech => "SECH",
            Self::Sin => "SIN",
            Self::Sinh => "SINH",
            Self::Tan => "TAN",
            Self::Tanh => "TANH",
        }
    }

    fn input(self, index: usize) -> f64 {
        match self {
            Self::Acos
            | Self::Acot
            | Self::Asin
            | Self::Atan
            | Self::Cos
            | Self::Cot
            | Self::Sin
            | Self::Tan => 0.17 + (index % 9) as f64 * 0.047,
            Self::Atan2 => 0.17 + (index % 9) as f64 * 0.047,
            Self::Acosh | Self::Acoth => 1.35 + (index % 9) as f64 * 0.11,
            Self::Asinh
            | Self::Atanh
            | Self::Cosh
            | Self::Coth
            | Self::Csch
            | Self::Sech
            | Self::Sinh
            | Self::Tanh => 0.17 + (index % 9) as f64 * 0.047,
            Self::Csc | Self::Sec => 0.17 + (index % 9) as f64 * 0.047,
            Self::Degrees => 0.17 + (index % 9) as f64 * 0.047,
            Self::Pi => 0.0,
            Self::Radians => 17.0 + (index % 9) as f64 * 4.7,
        }
    }

    fn expected(self, input: f64) -> f64 {
        match self {
            Self::Acos => input.acos(),
            Self::Acosh => input.acosh(),
            Self::Acot => 1.0_f64.atan2(input),
            Self::Acoth => 0.5 * ((input + 1.0) / (input - 1.0)).ln(),
            Self::Asin => input.asin(),
            Self::Asinh => input.asinh(),
            Self::Atan => input.atan(),
            Self::Atan2 => (input + 0.47).atan2(input),
            Self::Atanh => input.atanh(),
            Self::Cos => input.cos(),
            Self::Cosh => input.cosh(),
            Self::Cot => 1.0 / input.tan(),
            Self::Coth => 1.0 / input.tanh(),
            Self::Csc => 1.0 / input.sin(),
            Self::Csch => 1.0 / input.sinh(),
            Self::Degrees => input.to_degrees(),
            Self::Pi => std::f64::consts::PI,
            Self::Radians => input.to_radians(),
            Self::Sec => 1.0 / input.cos(),
            Self::Sech => 1.0 / input.cosh(),
            Self::Sin => input.sin(),
            Self::Sinh => input.sinh(),
            Self::Tan => input.tan(),
            Self::Tanh => input.tanh(),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ElementaryOp {
    Abs,
    Exp,
    Ln,
    Log,
    Log10,
    Power,
    Sqrt,
    SqrtPi,
    Sign,
    Mod,
    Quotient,
}

impl ElementaryOp {
    const fn name(self) -> &'static str {
        match self {
            Self::Abs => "ABS",
            Self::Exp => "EXP",
            Self::Ln => "LN",
            Self::Log => "LOG",
            Self::Log10 => "LOG10",
            Self::Power => "POWER",
            Self::Sqrt => "SQRT",
            Self::SqrtPi => "SQRTPI",
            Self::Sign => "SIGN",
            Self::Mod => "MOD",
            Self::Quotient => "QUOTIENT",
        }
    }

    const fn all() -> &'static [Self] {
        &[
            Self::Abs,
            Self::Exp,
            Self::Ln,
            Self::Log,
            Self::Log10,
            Self::Power,
            Self::Sqrt,
            Self::SqrtPi,
            Self::Sign,
            Self::Mod,
            Self::Quotient,
        ]
    }

    fn input(self, index: usize) -> f64 {
        let magnitude = 0.17 + (index % 9) as f64 * 0.047;
        match self {
            Self::Abs | Self::Sign => -magnitude,
            Self::Exp => magnitude,
            Self::Ln | Self::Log10 | Self::Sqrt | Self::SqrtPi => 1.17 + magnitude,
            Self::Log => 10.17 + (index % 9) as f64 * 0.47,
            Self::Power => 1.17 + magnitude,
            Self::Mod | Self::Quotient => 7.17 + (index % 9) as f64 * 0.47,
        }
    }

    const fn second(self) -> f64 {
        match self {
            Self::Log | Self::Log10 => 10.0,
            Self::Power => 2.25,
            Self::Mod | Self::Quotient => 2.3,
            Self::Abs | Self::Exp | Self::Ln | Self::Sqrt | Self::SqrtPi | Self::Sign => 0.0,
        }
    }

    fn expected(self, input: f64) -> f64 {
        match self {
            Self::Abs => input.abs(),
            Self::Exp => input.exp(),
            Self::Ln => input.ln(),
            Self::Log => input.log(self.second()),
            Self::Log10 => input.log10(),
            Self::Power => input.powf(self.second()),
            Self::Sqrt => input.sqrt(),
            Self::SqrtPi => (input * std::f64::consts::PI).sqrt(),
            Self::Sign => input.signum(),
            Self::Mod => input % self.second(),
            Self::Quotient => (input / self.second()).trunc(),
        }
    }

    fn source(self, argument: &str) -> String {
        match self {
            Self::Log => format!("=LOG({argument};10)"),
            Self::Power => format!("=POWER({argument};2.25)"),
            Self::Mod => format!("=MOD({argument};2.3)"),
            Self::Quotient => format!("=QUOTIENT({argument};2.3)"),
            _ => format!("={}({argument})", self.name()),
        }
    }

    const fn needs_positive(self) -> bool {
        matches!(
            self,
            Self::Ln | Self::Log | Self::Log10 | Self::Sqrt | Self::SqrtPi | Self::Power
        )
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Shape {
    Scalar,
    Array { rows: usize, columns: usize },
    ReferenceScalar,
    ReferenceArray { rows: usize, columns: usize },
}

impl Shape {
    const fn is_scalar(self) -> bool {
        matches!(self, Self::Scalar | Self::ReferenceScalar)
    }

    const fn dimensions(self) -> Option<(usize, usize)> {
        match self {
            Self::Array { rows, columns } | Self::ReferenceArray { rows, columns } => {
                Some((rows, columns))
            },
            Self::Scalar | Self::ReferenceScalar => None,
        }
    }

    const fn mode(self) -> Mode {
        match self {
            Self::ReferenceScalar => Mode::Scalar,
            Self::Scalar | Self::Array { .. } | Self::ReferenceArray { .. } => Mode::Matrix,
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum Operation {
    Arithmetic,
    Round,
    Trig(TrigOp),
    Elementary(ElementaryOp),
}

impl Operation {
    #[allow(dead_code)]
    const fn is_trig(self) -> bool {
        matches!(self, Self::Trig(_))
    }

    const fn is_elementary(self) -> bool {
        matches!(self, Self::Elementary(_))
    }

    fn name(self) -> &'static str {
        match self {
            Self::Arithmetic => "ARITHMETIC",
            Self::Round => "ROUND",
            Self::Trig(op) => op.name(),
            Self::Elementary(op) => op.name(),
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
struct FixtureResolver;

impl Resolver for FixtureResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        execution.check()?;
        Ok((sheet == "Main").then_some(SheetExtent::new(32, 32)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        execution.check()?;
        if sheet != "Main" || row >= 32 || column >= 32 {
            return Ok(CellRead::Error(
                litchi_ods::codec::formula::evaluation::ScalarError::Reference,
            ));
        }
        Ok(CellRead::Number(reference_value(row, column)))
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

fn fixture_value(index: usize) -> f64 {
    0.17 + (index % 9) as f64 * 0.047
}

fn reference_value(row: usize, column: usize) -> f64 {
    if column >= 4 {
        reference_domain_value(row.saturating_mul(4).saturating_add(column - 4))
    } else {
        fixture_value(row.saturating_mul(4).saturating_add(column))
    }
}

fn reference_domain_value(index: usize) -> f64 {
    1.35 + ((index.saturating_add(4)) % 9) as f64 * 0.11
}

fn array_literal(rows: usize, columns: usize, operation: Operation) -> String {
    let mut output = String::with_capacity(rows.saturating_mul(columns).saturating_mul(7) + 2);
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
            let value = match operation {
                Operation::Trig(op) => op.input(index),
                Operation::Elementary(op) => op.input(index),
                Operation::Arithmetic | Operation::Round => fixture_value(index),
            };
            let _ = write!(output, "{value:.9}");
        }
    }
    output.push('}');
    output
}

fn function_source(operation: Operation, argument: &str) -> String {
    match operation {
        Operation::Arithmetic => format!("={argument}+1.25"),
        Operation::Round => format!("=ROUND({argument};1)"),
        Operation::Trig(TrigOp::Atan2) => format!("=ATAN2({argument};{argument})"),
        Operation::Trig(TrigOp::Pi) => "=PI()".to_owned(),
        Operation::Trig(op) => format!("={}( {})", op.name(), argument).replace("( ", "("),
        Operation::Elementary(op) => op.source(argument),
    }
}

fn reference_argument(operation: Operation, array: bool) -> &'static str {
    if matches!(operation, Operation::Trig(TrigOp::Acosh | TrigOp::Acoth))
        || matches!(operation, Operation::Elementary(op) if op.needs_positive())
    {
        if array { "[.E1:.H4]" } else { "[.E1]" }
    } else if array {
        "[.A1:.D4]"
    } else {
        "[.A1]"
    }
}

fn case_input(case: &Case, index: usize) -> f64 {
    if matches!(
        case.shape,
        Shape::ReferenceScalar | Shape::ReferenceArray { .. }
    ) && matches!(
        case.operation,
        Operation::Trig(TrigOp::Acosh | TrigOp::Acoth)
            | Operation::Elementary(
                ElementaryOp::Ln
                    | ElementaryOp::Log
                    | ElementaryOp::Log10
                    | ElementaryOp::Power
                    | ElementaryOp::Sqrt
                    | ElementaryOp::SqrtPi
            )
    ) {
        reference_domain_value(index)
    } else {
        match case.operation {
            Operation::Arithmetic | Operation::Round => fixture_value(index),
            Operation::Trig(op) => op.input(index),
            Operation::Elementary(op) => op.input(index),
        }
    }
}

fn expected(operation: Operation, input: f64) -> f64 {
    match operation {
        Operation::Arithmetic => input + 1.25,
        Operation::Round => (input * 10.0).round() / 10.0,
        Operation::Trig(op) => op.expected(input),
        Operation::Elementary(op) => op.expected(input),
    }
}

fn scalar_source(operation: Operation) -> String {
    match operation {
        Operation::Arithmetic => "=0.17+1.25".to_owned(),
        Operation::Round => "=ROUND(0.234;1)".to_owned(),
        Operation::Trig(TrigOp::Atan2) => "=ATAN2(0.17;0.64)".to_owned(),
        Operation::Trig(TrigOp::Pi) => "=PI()".to_owned(),
        Operation::Trig(op) => {
            let value = op.input(0);
            format!("={}( {value:.9})", op.name()).replace("( ", "(")
        },
        Operation::Elementary(op) => {
            let value = op.input(0);
            op.source(&format!("{value:.9}"))
        },
    }
}

fn all_cases() -> Vec<Case> {
    let mut cases = Vec::new();
    for operation in [
        Operation::Arithmetic,
        Operation::Round,
        Operation::Trig(TrigOp::Sin),
    ] {
        cases.push(Case {
            name: format!("scalar-control-{}", operation.name().to_ascii_lowercase()),
            source: scalar_source(operation),
            shape: Shape::Scalar,
            operation,
        });
    }
    for operation in ElementaryOp::all()
        .iter()
        .copied()
        .map(Operation::Elementary)
    {
        cases.push(Case {
            name: format!(
                "scalar-elementary-{}",
                operation.name().to_ascii_lowercase()
            ),
            source: scalar_source(operation),
            shape: Shape::Scalar,
            operation,
        });
    }

    for (rows, columns) in [(4, 4), (16, 16)] {
        let array_operations = [
            Operation::Arithmetic,
            Operation::Round,
            Operation::Trig(TrigOp::Sin),
            Operation::Elementary(ElementaryOp::Abs),
            Operation::Elementary(ElementaryOp::Exp),
            Operation::Elementary(ElementaryOp::Ln),
            Operation::Elementary(ElementaryOp::Sqrt),
            Operation::Elementary(ElementaryOp::Sign),
        ];
        for operation in array_operations {
            let literal = array_literal(rows, columns, operation);
            let prefix = if operation.is_elementary() {
                "array-elementary"
            } else {
                "array-control"
            };
            cases.push(Case {
                name: format!(
                    "{prefix}-{}x{}-{}",
                    rows,
                    columns,
                    operation.name().to_ascii_lowercase()
                ),
                source: function_source(operation, &literal),
                shape: Shape::Array { rows, columns },
                operation,
            });
        }
    }

    let reference_scalar_operations = [
        Operation::Arithmetic,
        Operation::Round,
        Operation::Trig(TrigOp::Sin),
        Operation::Elementary(ElementaryOp::Abs),
        Operation::Elementary(ElementaryOp::Ln),
        Operation::Elementary(ElementaryOp::Sqrt),
    ];
    for operation in reference_scalar_operations {
        let prefix = if operation.is_elementary() {
            "reference-elementary"
        } else {
            "reference-scalar"
        };
        cases.push(Case {
            name: format!("{prefix}-{}", operation.name().to_ascii_lowercase()),
            source: function_source(operation, reference_argument(operation, false)),
            shape: Shape::ReferenceScalar,
            operation,
        });
    }
    let reference_array_operations = [
        Operation::Arithmetic,
        Operation::Round,
        Operation::Trig(TrigOp::Sin),
        Operation::Elementary(ElementaryOp::Abs),
        Operation::Elementary(ElementaryOp::Ln),
        Operation::Elementary(ElementaryOp::Sqrt),
    ];
    for operation in reference_array_operations {
        let prefix = if operation.is_elementary() {
            "reference-elementary-array"
        } else {
            "reference-array"
        };
        cases.push(Case {
            name: format!("{prefix}-{}", operation.name().to_ascii_lowercase()),
            source: function_source(operation, reference_argument(operation, true)),
            shape: Shape::ReferenceArray {
                rows: 4,
                columns: 4,
            },
            operation,
        });
    }
    cases
}

fn default_repeat(shape: Shape) -> usize {
    match shape {
        Shape::Scalar | Shape::ReferenceScalar => 1_000,
        Shape::Array {
            rows: 4,
            columns: 4,
        }
        | Shape::ReferenceArray {
            rows: 4,
            columns: 4,
        } => 80,
        Shape::Array {
            rows: 16,
            columns: 16,
        } => 4,
        Shape::ReferenceArray { .. } | Shape::Array { .. } => 1,
    }
}

fn execution() -> ExecutionContext {
    let budget = Budget::root(
        "ods-formula-elementary-math-performance",
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
        other => Err(format!("expected scalar Number, got {other:?}").into()),
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

fn validate_scalar(case: &Case, value: &ScalarValue<'_>) -> AnyResult<()> {
    let ScalarValue::Number(actual) = value else {
        return Err(format!("{} returned {value:?}", case.name).into());
    };
    let input = case_input(case, 0);
    let expected = expected(case.operation, input);
    if approximately_equal(*actual, expected) {
        Ok(())
    } else {
        Err(format!("{} returned {actual}, expected {expected}", case.name).into())
    }
}

fn validate_array(case: &Case, array: value::ArrayView<'_>) -> AnyResult<()> {
    let (rows, columns) = case
        .shape
        .dimensions()
        .ok_or("array validation requires dimensions")?;
    if (array.shape().rows(), array.shape().columns()) != (rows, columns) {
        return Err(format!(
            "{} returned {}x{}, expected {rows}x{columns}",
            case.name,
            array.shape().rows(),
            array.shape().columns()
        )
        .into());
    }
    for index in 0..rows.saturating_mul(columns) {
        let Value::Number(actual) = array
            .get(index)
            .ok_or_else(|| format!("{} missing array cell {index}", case.name))?
        else {
            return Err(format!("{} array cell {index} was not numeric", case.name).into());
        };
        let input_index = match case.shape {
            Shape::Array { rows: _, columns } | Shape::ReferenceArray { rows: _, columns } => {
                (index / columns)
                    .saturating_mul(columns)
                    .saturating_add(index % columns)
            },
            _ => index,
        };
        let input = case_input(case, input_index);
        let expected = expected(case.operation, input);
        if !approximately_equal(actual, expected) {
            return Err(format!(
                "{} array cell {index}: {actual}, expected {expected}",
                case.name
            )
            .into());
        }
    }
    Ok(())
}

fn eval_scalar<'a>(
    case: &Case,
    expression: &'a Expression,
    execution: &ExecutionContext,
) -> Result<litchi_ods::codec::formula::evaluation::EvaluatedScalar<'a>, EvaluationFailure> {
    evaluate_scalar(
        expression,
        &EvaluationContext::new(execution),
        &EvaluationLimits::default(),
    )
    .map_err(|error| {
        // The argument is retained to keep this helper's call sites explicit
        // about the case whose operation is being evaluated.
        let _ = case;
        error
    })
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

fn is_unsupported(error: &EvaluationFailure) -> bool {
    matches!(error, EvaluationFailure::Unsupported(_))
}

fn preflight(case: &Case, allow_unsupported_elementary: bool) -> AnyResult<bool> {
    let expression = Expression::parse(&case.source)
        .map_err(|error| format!("{} parse preflight failed: {error}", case.name))?;
    let execution = execution();
    match case.shape {
        Shape::Scalar => match eval_scalar(case, &expression, &execution) {
            Ok(result) => {
                validate_scalar(case, result.value())?;
                Ok(true)
            },
            Err(error)
                if case.operation.is_elementary()
                    && allow_unsupported_elementary
                    && is_unsupported(&error) =>
            {
                Ok(false)
            },
            Err(error) => Err(format!("{} preflight failed: {error}", case.name).into()),
        },
        Shape::Array { .. } | Shape::ReferenceScalar | Shape::ReferenceArray { .. } => {
            let resolver = FixtureResolver;
            match eval_value(case, &expression, &resolver, &execution) {
                Ok(result) => {
                    if case.shape.is_scalar() {
                        let Value::Number(actual) = result.value() else {
                            return Err(
                                format!("{} returned {:?}", case.name, result.value()).into()
                            );
                        };
                        let input = case_input(case, 0);
                        let expected = expected(case.operation, input);
                        if !approximately_equal(actual, expected) {
                            return Err(format!(
                                "{} returned {actual}, expected {expected}",
                                case.name
                            )
                            .into());
                        }
                    } else {
                        validate_array(
                            case,
                            result.as_array().ok_or("value result was not an array")?,
                        )?;
                    }
                    Ok(true)
                },
                Err(error)
                    if case.operation.is_elementary()
                        && allow_unsupported_elementary
                        && is_unsupported(&error) =>
                {
                    Ok(false)
                },
                Err(error) => Err(format!("{} preflight failed: {error}", case.name).into()),
            }
        },
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
    allow_unsupported_elementary: bool,
}

fn parse_config() -> AnyResult<Option<Config>> {
    let mut case = None;
    let mut phase = Phase::Evaluate;
    let mut warmups = 3;
    let mut iterations = 1;
    let mut repeat = None;
    let mut allow_unsupported_elementary = false;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                println!(
                    "usage: ods-formula-elementary-math-performance [--case NAME|all] [--phase evaluate|parse-evaluate] [--warmups N] [--iterations N] [--repeat N] [--allow-unsupported-elementary]"
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
                    .parse()
                    .map_err(|_| "--warmups requires a nonnegative integer")?;
            },
            "--iterations" => {
                iterations = arguments
                    .next()
                    .ok_or("--iterations requires an integer")?
                    .parse()
                    .map_err(|_| "--iterations requires a positive integer")?;
            },
            "--repeat" => {
                repeat = Some(
                    arguments
                        .next()
                        .ok_or("--repeat requires an integer")?
                        .parse()
                        .map_err(|_| "--repeat requires a positive integer")?,
                );
            },
            "--allow-unsupported-elementary" => allow_unsupported_elementary = true,
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
        allow_unsupported_elementary,
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
    unsupported: bool,
) -> AnyResult<u64> {
    match result {
        Ok(result) => {
            if unsupported {
                return Err("baseline unexpectedly evaluated an elementary case".into());
            }
            *checksum = checksum.wrapping_add(scalar_checksum(result.value())?);
            black_box(*checksum);
            let retained = execution.budget().used(Resource::Memory);
            drop(result);
            Ok(retained)
        },
        Err(error) if unsupported && is_unsupported(&error) => {
            *checksum = checksum.wrapping_add(0x5355_4e53_5550);
            black_box(*checksum);
            Ok(execution.budget().used(Resource::Memory))
        },
        Err(error) => Err(format!("scalar evaluation failed: {error}").into()),
    }
}

fn consume_value(
    result: Result<Evaluated<'_>, EvaluationFailure>,
    case: &Case,
    execution: &ExecutionContext,
    checksum: &mut u64,
    unsupported: bool,
) -> AnyResult<u64> {
    match result {
        Ok(result) => {
            if unsupported {
                return Err("baseline unexpectedly evaluated an elementary case".into());
            }
            if case.shape.is_scalar() {
                let Value::Number(number) = result.value() else {
                    return Err(format!("{} returned {:?}", case.name, result.value()).into());
                };
                *checksum = checksum.wrapping_add(number.to_bits().rotate_left(11));
            } else {
                *checksum = checksum.wrapping_add(array_checksum(
                    result.as_array().ok_or("value result was not an array")?,
                )?);
            }
            black_box(*checksum);
            let retained = execution.budget().used(Resource::Memory);
            drop(result);
            Ok(retained)
        },
        Err(error) if unsupported && is_unsupported(&error) => {
            *checksum = checksum.wrapping_add(0x5355_4e53_5550);
            black_box(*checksum);
            Ok(execution.budget().used(Resource::Memory))
        },
        Err(error) => Err(format!("value evaluation failed: {error}").into()),
    }
}

fn sample_from(
    started: Instant,
    execution: &ExecutionContext,
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
        checksum,
    }
}

fn measure_evaluate(
    case: &Case,
    expression: &Expression,
    repeat: usize,
    unsupported: bool,
) -> AnyResult<Sample> {
    let execution = execution();
    let resolver = FixtureResolver;
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
        match case.shape {
            Shape::Scalar => {
                retained_peak = retained_peak.max(consume_scalar(
                    evaluate_scalar(expression, &scalar_context, &scalar_limits),
                    &execution,
                    &mut checksum,
                    unsupported,
                )?);
            },
            Shape::Array { .. } | Shape::ReferenceScalar | Shape::ReferenceArray { .. } => {
                retained_peak = retained_peak.max(consume_value(
                    value::evaluate(expression, &resolver, &context, &value_limits),
                    case,
                    &execution,
                    &mut checksum,
                    unsupported,
                )?);
            },
        }
    }
    Ok(sample_from(
        started,
        &execution,
        baseline_work,
        baseline_memory,
        live_before,
        checksum,
        retained_peak,
    ))
}

fn measure_parse_evaluate(case: &Case, repeat: usize, unsupported: bool) -> AnyResult<Sample> {
    let execution = execution();
    let resolver = FixtureResolver;
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
        match case.shape {
            Shape::Scalar => {
                retained_peak = retained_peak.max(consume_scalar(
                    evaluate_scalar(&expression, &scalar_context, &scalar_limits),
                    &execution,
                    &mut checksum,
                    unsupported,
                )?);
            },
            Shape::Array { .. } | Shape::ReferenceScalar | Shape::ReferenceArray { .. } => {
                retained_peak = retained_peak.max(consume_value(
                    value::evaluate(&expression, &resolver, &context, &value_limits),
                    case,
                    &execution,
                    &mut checksum,
                    unsupported,
                )?);
            },
        }
        retained_peak = retained_peak.max(execution.budget().used(Resource::Memory));
        black_box(expression);
    }
    Ok(sample_from(
        started,
        &execution,
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
    repeat: usize,
    warmups: usize,
    iterations: usize,
    supported: bool,
    samples: &[Sample],
) {
    let (shape, rows, columns, elements) = match case.shape {
        Shape::Scalar | Shape::ReferenceScalar => ("scalar", 0, 0, 1),
        Shape::Array { rows, columns } | Shape::ReferenceArray { rows, columns } => {
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
        "{{\"case\":\"{}\",\"operation\":\"{}\",\"phase\":\"{}\",\"input_bytes\":{},\"repeat\":{},\"warmups\":{},\"iterations\":{},\"shape\":\"{}\",\"rows\":{},\"columns\":{},\"elements\":{},\"supported\":{},\"elapsed_ns_p50\":{},\"elapsed_ns_mean\":{},\"elapsed_ns_p95\":{},\"elapsed_ns_p99\":{},\"elapsed_ns_per_repeat\":{},\"allocator_calls_p50\":{},\"allocator_calls_max\":{},\"deallocator_calls_p50\":{},\"requested_bytes_p50\":{},\"requested_bytes_max\":{},\"released_bytes_p50\":{},\"peak_live_delta_p50\":{},\"peak_live_delta_max\":{},\"memory_retained_p50\":{},\"memory_retained_max\":{},\"work_p50\":{},\"work_per_repeat\":{},\"checksum_p50\":{},\"live_before_p50\":{},\"live_after_p50\":{},\"samples\":{},\"rss_kib\":null,\"rss_source\":\"external /usr/bin/time -v\",\"memory_retained_source\":\"execution budget while result is live\",\"validation_scope\":\"one untimed direct f64 numeric/shape oracle independent of evaluator dispatch; timed evaluator/checksum/drop\"}}",
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
        supported,
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
    let supported = preflight(case, config.allow_unsupported_elementary)?;
    let repeat = config.repeat.unwrap_or_else(|| default_repeat(case.shape));
    let expression = Expression::parse(&case.source)
        .map_err(|error| format!("{} parse preflight failed: {error}", case.name))?;
    for _ in 0..config.warmups {
        let sample = match config.phase {
            Phase::Evaluate => measure_evaluate(case, &expression, repeat, !supported)?,
            Phase::ParseEvaluate => measure_parse_evaluate(case, repeat, !supported)?,
        };
        black_box(sample);
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(match config.phase {
            Phase::Evaluate => measure_evaluate(case, &expression, repeat, !supported)?,
            Phase::ParseEvaluate => measure_parse_evaluate(case, repeat, !supported)?,
        });
    }
    emit(
        case,
        config.phase,
        repeat,
        config.warmups,
        config.iterations,
        supported,
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
