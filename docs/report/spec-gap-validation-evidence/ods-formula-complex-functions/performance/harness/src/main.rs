//! Bounded performance harness for the OpenFormula complex-number family.
//!
//! The corpus covers every §6.8 function, representative text and reference
//! conversions, variadic aggregation, lazy branches, nested calls, and typed
//! error/resource boundaries.  The evaluator result is timed directly.  A
//! successful complex result is checked outside that timer against the public
//! `Value::Complex` components and suffix, then independently through
//! `IMREAL` and `IMAGINARY` wrappers.  The second check keeps the numerical
//! oracle independent of the evaluator's internal complex kernels.
//!
//! Parsing, resolver construction, and the untimed oracle are outside the
//! evaluate timer.  The evaluate timer includes evaluator execution,
//! checksum folding, and result drop.  Allocation counters describe the
//! timed process and the execution budget's retained-memory observation is
//! sampled while a result is live.  RSS is supplied by an external
//! `/usr/bin/time -v` wrapper.  The harness has no library statistics or
//! benchmark dependencies.

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
        CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, Value, evaluate,
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

/// A fixed-size independent oracle value.  It carries the representation
/// suffix because the public complex value preserves that metadata even
/// though it does not affect arithmetic.
#[derive(Clone, Copy, Debug)]
struct ComplexPair {
    real: f64,
    imaginary: f64,
    suffix: char,
}

impl ComplexPair {
    const fn new(real: f64, imaginary: f64) -> Self {
        Self::with_suffix(real, imaginary, 'i')
    }

    const fn with_suffix(real: f64, imaginary: f64, suffix: char) -> Self {
        Self {
            real,
            imaginary,
            suffix,
        }
    }

    fn abs(self) -> f64 {
        self.real.hypot(self.imaginary)
    }

    fn argument(self) -> f64 {
        self.imaginary.atan2(self.real)
    }

    fn add(self, other: Self) -> Self {
        Self::with_suffix(
            self.real + other.real,
            self.imaginary + other.imaginary,
            choose_suffix(self.suffix, other.suffix),
        )
    }

    fn sub(self, other: Self) -> Self {
        Self::with_suffix(
            self.real - other.real,
            self.imaginary - other.imaginary,
            choose_suffix(self.suffix, other.suffix),
        )
    }

    fn mul(self, other: Self) -> Self {
        Self::with_suffix(
            self.real * other.real - self.imaginary * other.imaginary,
            self.real * other.imaginary + self.imaginary * other.real,
            choose_suffix(self.suffix, other.suffix),
        )
    }

    fn div(self, other: Self) -> Self {
        let denominator_scale = other.real.abs().max(other.imaginary.abs());
        let numerator_scale = self.real.abs().max(self.imaginary.abs());
        let suffix = choose_suffix(self.suffix, other.suffix);
        if denominator_scale == 0.0 {
            return Self::with_suffix(f64::NAN, f64::NAN, suffix);
        }
        if numerator_scale == 0.0 {
            return Self::with_suffix(0.0, 0.0, suffix);
        }
        let real_denominator = other.real / denominator_scale;
        let imaginary_denominator = other.imaginary / denominator_scale;
        let denominator =
            real_denominator * real_denominator + imaginary_denominator * imaginary_denominator;
        let real_numerator = self.real / numerator_scale;
        let imaginary_numerator = self.imaginary / numerator_scale;
        let real_factor = (real_numerator * real_denominator
            + imaginary_numerator * imaginary_denominator)
            / denominator;
        let imaginary_factor = (imaginary_numerator * real_denominator
            - real_numerator * imaginary_denominator)
            / denominator;
        Self::with_suffix(
            numerator_scale * real_factor / denominator_scale,
            numerator_scale * imaginary_factor / denominator_scale,
            suffix,
        )
    }

    fn exp(self) -> Self {
        let scale = self.real.exp();
        Self::with_suffix(
            scale * self.imaginary.cos(),
            scale * self.imaginary.sin(),
            self.suffix,
        )
    }

    fn ln(self) -> Self {
        Self::with_suffix(self.abs().ln(), self.argument(), self.suffix)
    }

    fn sin(self) -> Self {
        Self::with_suffix(
            self.real.sin() * self.imaginary.cosh(),
            self.real.cos() * self.imaginary.sinh(),
            self.suffix,
        )
    }

    fn cos(self) -> Self {
        Self::with_suffix(
            self.real.cos() * self.imaginary.cosh(),
            -self.real.sin() * self.imaginary.sinh(),
            self.suffix,
        )
    }

    fn sinh(self) -> Self {
        Self::with_suffix(
            self.real.sinh() * self.imaginary.cos(),
            self.real.cosh() * self.imaginary.sin(),
            self.suffix,
        )
    }

    fn cosh(self) -> Self {
        Self::with_suffix(
            self.real.cosh() * self.imaginary.cos(),
            self.real.sinh() * self.imaginary.sin(),
            self.suffix,
        )
    }

    fn reciprocal(self) -> Self {
        Self::with_suffix(1.0, 0.0, self.suffix).div(self)
    }

    fn powi(self, exponent: usize) -> Self {
        let mut result = Self::with_suffix(1.0, 0.0, self.suffix);
        for _ in 0..exponent {
            result = result.mul(self);
        }
        result
    }

    fn sqrt(self) -> Self {
        let magnitude = self.abs();
        let real = ((magnitude + self.real) / 2.0).max(0.0).sqrt();
        let sign = if self.imaginary.is_sign_negative() {
            -1.0
        } else {
            1.0
        };
        let imaginary = sign * ((magnitude - self.real) / 2.0).max(0.0).sqrt();
        Self::with_suffix(real, imaginary, self.suffix)
    }
}

fn choose_suffix(left: char, right: char) -> char {
    if left == 'j' || right == 'j' {
        'j'
    } else {
        'i'
    }
}

#[derive(Clone, Copy, Debug)]
enum Expected {
    Complex(ComplexPair),
    Number(f64),
    FormulaError,
    EvaluationFailure(&'static str),
}

#[derive(Debug)]
struct CaseSpec {
    name: String,
    function: &'static str,
    source: String,
    size: usize,
    group: &'static str,
    expected: Expected,
}

#[derive(Clone, Copy, Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(&'static str),
    Error(ScalarError),
}

#[derive(Debug)]
struct ResolverStats {
    reads: AtomicU64,
    extent_calls: AtomicU64,
    sheet_index_calls: AtomicU64,
    sheet_name_calls: AtomicU64,
    sheet_count_calls: AtomicU64,
}

impl ResolverStats {
    fn new() -> Self {
        Self {
            reads: AtomicU64::new(0),
            extent_calls: AtomicU64::new(0),
            sheet_index_calls: AtomicU64::new(0),
            sheet_name_calls: AtomicU64::new(0),
            sheet_count_calls: AtomicU64::new(0),
        }
    }

    fn reset(&self) {
        self.reads.store(0, Ordering::Relaxed);
        self.extent_calls.store(0, Ordering::Relaxed);
        self.sheet_index_calls.store(0, Ordering::Relaxed);
        self.sheet_name_calls.store(0, Ordering::Relaxed);
        self.sheet_count_calls.store(0, Ordering::Relaxed);
    }

    fn snapshot(self: &Self) -> ResolverSnapshot {
        ResolverSnapshot {
            reads: self.reads.load(Ordering::Acquire),
            extent_calls: self.extent_calls.load(Ordering::Acquire),
            sheet_index_calls: self.sheet_index_calls.load(Ordering::Acquire),
            sheet_name_calls: self.sheet_name_calls.load(Ordering::Acquire),
            sheet_count_calls: self.sheet_count_calls.load(Ordering::Acquire),
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct ResolverSnapshot {
    reads: u64,
    extent_calls: u64,
    sheet_index_calls: u64,
    sheet_name_calls: u64,
    sheet_count_calls: u64,
}

/// Deterministic in-memory provider for reference and lazy cases.  Text is
/// static and returned directly, so the provider itself does not copy complex
/// text while the evaluator is timed.
#[derive(Debug)]
struct ComplexResolver {
    extent: litchi_ods::codec::formula::evaluation::value::SheetExtent,
    cells: Vec<FixtureCell>,
    stats: ResolverStats,
}

impl ComplexResolver {
    fn for_case(case: &str) -> Self {
        let reference_size = case
            .rsplit_once('-')
            .and_then(|(_, suffix)| suffix.parse::<usize>().ok())
            .unwrap_or(1);
        let row_count = if case.starts_with("reference-") {
            reference_size.max(8).saturating_add(2)
        } else {
            8
        };
        let mut cells = vec![FixtureCell::Empty; row_count];
        cells[0] = FixtureCell::Number(7.0);
        cells[1] = FixtureCell::Text("3+4i");
        cells[2] = FixtureCell::Empty;
        cells[3] = FixtureCell::Error(ScalarError::NotAvailable);
        cells[4] = FixtureCell::Logical(true);

        if let Some(size) = case
            .strip_prefix("reference-complex-sum-")
            .and_then(|suffix| suffix.parse::<usize>().ok())
        {
            for cell in cells.iter_mut().take(size) {
                *cell = FixtureCell::Text("1+1i");
            }
        } else if let Some(size) = case
            .strip_prefix("reference-complex-product-")
            .and_then(|suffix| suffix.parse::<usize>().ok())
        {
            for cell in cells.iter_mut().take(size) {
                *cell = FixtureCell::Text("1+0i");
            }
        } else if case == "reference-sequence-specials" {
            cells[0] = FixtureCell::Empty;
            cells[1] = FixtureCell::Logical(true);
            cells[2] = FixtureCell::Text("1+1i");
        } else if case == "reference-sequence-error" {
            cells[0] = FixtureCell::Error(ScalarError::NotAvailable);
            cells[1] = FixtureCell::Text("1+1i");
        }

        Self {
            extent: litchi_ods::codec::formula::evaluation::value::SheetExtent::new(row_count, 1),
            cells,
            stats: ResolverStats::new(),
        }
    }

    fn reset_stats(&self) {
        self.stats.reset();
    }

    fn stats(&self) -> ResolverSnapshot {
        self.stats.snapshot()
    }
}

impl Resolver for ComplexResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<litchi_ods::codec::formula::evaluation::value::SheetExtent>, EvaluationFailure>
    {
        execution.check()?;
        self.stats.extent_calls.fetch_add(1, Ordering::Relaxed);
        Ok((sheet == "Main").then_some(self.extent))
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
        let value = if sheet == "Main" && column == 0 {
            self.cells.get(row).copied().unwrap_or(FixtureCell::Empty)
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
        self.stats.sheet_index_calls.fetch_add(1, Ordering::Relaxed);
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        execution.check()?;
        self.stats.sheet_name_calls.fetch_add(1, Ordering::Relaxed);
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        execution.check()?;
        self.stats.sheet_count_calls.fetch_add(1, Ordering::Relaxed);
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
    group: Option<String>,
    phase: Phase,
    warmups: usize,
    iterations: usize,
    repeat: Option<usize>,
}

fn parse_config() -> AnyResult<Option<Config>> {
    let mut case = None;
    let mut group = None;
    let mut phase = Phase::Evaluate;
    let mut warmups = 3;
    let mut iterations = 15;
    let mut repeat = None;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                println!(
                    "usage: complex-profile [--case NAME|all] [--group NAME] [--phase evaluate|parse-evaluate|parse] [--warmups N] [--iterations N] [--repeat N]\n\
                     --list prints deterministic case names.  Evaluate reuses parsed expressions and a finite in-memory resolver.\n\
                     Complex outputs are checked against Value::Complex components/suffix and untimed IMREAL/IMAGINARY wrappers.  Wrap the binary in /usr/bin/time -v for RSS."
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
            "--group" => {
                group = Some(arguments.next().ok_or("--group requires a name")?);
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
        group,
        phase,
        warmups,
        iterations,
        repeat,
    }))
}

fn default_repeat(size: usize) -> usize {
    match size {
        0..=2 => 16,
        4 | 8 => 4,
        16 => 2,
        _ => 1,
    }
}

fn case(
    name: &str,
    function: &'static str,
    source: impl Into<String>,
    size: usize,
    group: &'static str,
    expected: Expected,
) -> CaseSpec {
    CaseSpec {
        name: name.to_owned(),
        function,
        source: source.into(),
        size,
        group,
        expected,
    }
}

fn variadic(function: &str, count: usize, operand: &str) -> String {
    let mut source =
        String::with_capacity(count.saturating_mul(operand.len() + 1) + function.len() + 4);
    source.push('=');
    source.push_str(function);
    source.push('(');
    for index in 0..count {
        if index != 0 {
            source.push(';');
        }
        source.push_str(operand);
    }
    source.push(')');
    source
}

fn reference_range(function: &str, rows: usize) -> String {
    format!("={function}([.A1:.A{rows}])")
}

fn expected_components(expected: Expected) -> Option<ComplexPair> {
    match expected {
        Expected::Complex(value) => Some(value),
        Expected::Number(value) => Some(ComplexPair::new(value, 0.0)),
        Expected::FormulaError | Expected::EvaluationFailure(_) => None,
    }
}

fn expected_suffix(expected: Expected) -> Option<char> {
    match expected {
        Expected::Complex(value) => Some(value.suffix),
        Expected::Number(_) | Expected::FormulaError | Expected::EvaluationFailure(_) => None,
    }
}

fn expected_label(expected: Expected) -> &'static str {
    match expected {
        Expected::Complex(_) => "complex",
        Expected::Number(_) => "number",
        Expected::FormulaError => "formula-error",
        Expected::EvaluationFailure(_) => "evaluation-failure",
    }
}

fn expected_failure_label(expected: Expected) -> Option<&'static str> {
    match expected {
        Expected::EvaluationFailure(label) => Some(label),
        Expected::Complex(_) | Expected::Number(_) | Expected::FormulaError => None,
    }
}

fn component_source(source: &str, function: &str) -> String {
    let body = source
        .strip_prefix('=')
        .expect("complex corpus source starts with '='");
    format!("={function}({body})")
}

fn approximately_equal(actual: f64, expected: f64) -> bool {
    let tolerance = 1e-9 * actual.abs().max(expected.abs()).max(1.0);
    (actual - expected).abs() <= tolerance
}

fn validate_number(value: Value<'_>, expected: f64, label: &str) -> AnyResult<()> {
    match value {
        Value::Number(actual) if approximately_equal(actual, expected) => Ok(()),
        other => Err(format!("{label} returned {other:?}, expected Number({expected})").into()),
    }
}

fn validate_complex(value: Value<'_>, expected: ComplexPair, label: &str) -> AnyResult<()> {
    match value {
        Value::Complex(actual)
            if approximately_equal(actual.real(), expected.real)
                && approximately_equal(actual.imaginary(), expected.imaginary)
                && actual.suffix() == expected.suffix =>
        {
            Ok(())
        },
        Value::Complex(actual) => Err(format!(
            "{label} returned Complex({}, {} {:?}), expected Complex({}, {} {:?})",
            actual.real(),
            actual.imaginary(),
            actual.suffix(),
            expected.real,
            expected.imaginary,
            expected.suffix
        )
        .into()),
        other => {
            Err(format!("{label} returned {other:?}, expected a public Value::Complex").into())
        },
    }
}

fn execution(cancelled: bool) -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "ods-formula-complex-functions-performance",
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
    match case.expected {
        Expected::EvaluationFailure("work") => Limits::default().with_max_steps(1),
        Expected::EvaluationFailure("memory") => Limits::default().with_max_storage_bytes(0),
        Expected::Complex(_)
        | Expected::Number(_)
        | Expected::FormulaError
        | Expected::EvaluationFailure(_) => Limits::default(),
    }
}

fn is_cancelled(case: &CaseSpec) -> bool {
    matches!(case.expected, Expected::EvaluationFailure("cancelled"))
}

fn preflight_component(
    case: &CaseSpec,
    source: &str,
    expected: f64,
    component: &str,
) -> AnyResult<()> {
    let expression = Expression::parse(source)
        .map_err(|error| format!("{} {component} oracle parse failed: {error}", case.name))?;
    let resolver = ComplexResolver::for_case(&case.name);
    let (_cancellation, execution) = execution(false);
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = evaluate(&expression, &resolver, &context, &limits(case)).map_err(|error| {
        format!(
            "{} {component} oracle evaluation failed: {error}",
            case.name
        )
    })?;
    validate_number(
        result.value(),
        expected,
        &format!("{} {component}", case.name),
    )
}

fn matches_failure(label: &str, error: &EvaluationFailure) -> bool {
    match (label, error) {
        ("cancelled", EvaluationFailure::Cancelled) => true,
        ("work", EvaluationFailure::ResourceLimit(limit)) => limit.resource == Resource::Work,
        ("memory", EvaluationFailure::ResourceLimit(limit)) => limit.resource == Resource::Memory,
        ("text", EvaluationFailure::ResourceLimit(limit)) => {
            limit.resource == Resource::Memory && limit.observed == 32_768 && limit.limit == 32_767
        },
        _ => false,
    }
}

fn preflight_evaluate(case: &CaseSpec, expression: &Expression) -> AnyResult<()> {
    let resolver = ComplexResolver::for_case(&case.name);
    let (_cancellation, execution) = execution(is_cancelled(case));
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = evaluate(expression, &resolver, &context, &limits(case));
    match case.expected {
        Expected::EvaluationFailure(label) => match result {
            Err(error) if matches_failure(label, &error) => Ok(()),
            Err(error) => Err(format!("{} wrong {label} refusal: {error}", case.name).into()),
            Ok(value) => Err(format!(
                "{} unexpectedly returned a value {:?} for an evaluation failure",
                case.name,
                value.value()
            )
            .into()),
        },
        Expected::FormulaError => match result {
            Ok(value) if matches!(value.value(), Value::Error(_)) => Ok(()),
            Ok(value) => Err(format!(
                "{} returned {:?}, expected a formula error",
                case.name,
                value.value()
            )
            .into()),
            Err(error) => Err(format!("{} evaluator failure: {error}", case.name).into()),
        },
        Expected::Number(expected) => {
            let value =
                result.map_err(|error| format!("{} evaluator failure: {error}", case.name))?;
            validate_number(value.value(), expected, &case.name)
        },
        Expected::Complex(expected) => {
            let value =
                result.map_err(|error| format!("{} evaluator failure: {error}", case.name))?;
            validate_complex(value.value(), expected, &case.name)?;
            drop(value);
            preflight_component(
                case,
                &component_source(&case.source, "IMREAL"),
                expected.real,
                "real",
            )?;
            preflight_component(
                case,
                &component_source(&case.source, "IMAGINARY"),
                expected.imaginary,
                "imaginary",
            )
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
        Value::Complex(value) => {
            // Fold every public component and the suffix into the timed
            // checksum so a changed complex result cannot be optimized away
            // or collide with another complex value in the measurement.
            value.real().to_bits().rotate_left(11)
                ^ value.imaginary().to_bits().rotate_left(23)
                ^ (value.suffix() as u64).rotate_left(37)
                ^ 0x434f_4d50_4c45_5800
        },
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
        // Keep a fallback for future non-exhaustive value variants.  Known
        // complex values are handled explicitly above.
        _ => 0x434f_4d50_4c45_5800,
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

fn consume_result(
    case: &CaseSpec,
    result: Result<Evaluated<'_>, EvaluationFailure>,
    execution: &ExecutionContext,
    successes: &mut u64,
    refusals: &mut u64,
    checksum: &mut u64,
    retained_peak: &mut u64,
) -> AnyResult<()> {
    match (case.expected, result) {
        (Expected::EvaluationFailure(label), Err(error)) if matches_failure(label, &error) => {
            *refusals = refusals.saturating_add(1);
            *checksum = checksum.wrapping_add(0x4641_494c);
            black_box(*checksum);
            Ok(())
        },
        (Expected::EvaluationFailure(label), Ok(_value)) => Err(format!(
            "{} unexpectedly succeeded in {label} refusal lane",
            case.name
        )
        .into()),
        (_, Ok(value)) => {
            *successes = successes.saturating_add(1);
            let observed = black_box(value.value());
            *checksum = checksum.wrapping_add(checksum_value(observed));
            *retained_peak = (*retained_peak).max(execution.budget().used(Resource::Memory));
            black_box(*checksum);
            drop(value);
            Ok(())
        },
        (_, Err(error)) => Err(format!("{} evaluator failure: {error}", case.name).into()),
    }
}

fn measure_evaluate(case: &CaseSpec, expression: &Expression, repeat: usize) -> AnyResult<Sample> {
    let resolver = ComplexResolver::for_case(&case.name);
    resolver.reset_stats();
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
        resolver: resolver.stats(),
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
        resolver: ResolverSnapshot::default(),
    })
}

fn measure_parse_evaluate(case: &CaseSpec, repeat: usize) -> AnyResult<Sample> {
    let resolver = ComplexResolver::for_case(&case.name);
    resolver.reset_stats();
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
        resolver: resolver.stats(),
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
    let expected = expected_components(case.expected);
    let expected_real = expected
        .map(|value| format!("{:.17}", value.real))
        .unwrap_or_else(|| "null".to_owned());
    let expected_imaginary = expected
        .map(|value| format!("{:.17}", value.imaginary))
        .unwrap_or_else(|| "null".to_owned());
    let expected_suffix = expected_suffix(case.expected)
        .map(|value| format!("\"{}\"", value))
        .unwrap_or_else(|| "null".to_owned());
    let expected_failure = expected_failure_label(case.expected)
        .map(|value| format!("\"{}\"", json_escape(value)))
        .unwrap_or_else(|| "null".to_owned());
    let elapsed_ns = percentile(samples, |sample| sample.elapsed_ns, 50);
    let work = percentile(samples, |sample| sample.work, 50);
    println!(
        "{{\"case\":\"{}\",\"function\":\"{}\",\"group\":\"{}\",\"size\":{},\"phase\":\"{}\",\"input_bytes\":{},\"repeat\":{},\"warmups\":{},\"iterations\":{},\"elapsed_ns\":{},\"elapsed_ns_mean\":{},\"elapsed_ns_p95\":{},\"elapsed_ns_p99\":{},\"elapsed_ns_per_repeat\":{},\"checksum\":{},\"expected_kind\":\"{}\",\"expected_real\":{},\"expected_imaginary\":{},\"expected_suffix\":{},\"expected_failure\":{},\"allocator_calls\":{},\"allocator_calls_max\":{},\"deallocator_calls\":{},\"deallocator_calls_max\":{},\"requested_bytes\":{},\"requested_bytes_max\":{},\"released_bytes\":{},\"released_bytes_max\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"peak_live_delta_max\":{},\"memory_retained\":{},\"memory_retained_max\":{},\"work\":{},\"work_mean\":{},\"work_per_repeat\":{},\"successes\":{},\"refusals\":{},\"resolver_reads\":{},\"resolver_reads_max\":{},\"resolver_extent_calls\":{},\"resolver_extent_calls_max\":{},\"resolver_sheet_index_calls\":{},\"resolver_sheet_index_calls_max\":{},\"resolver_sheet_name_calls\":{},\"resolver_sheet_name_calls_max\":{},\"resolver_sheet_count_calls\":{},\"resolver_sheet_count_calls_max\":{},\"failure\":\"none\",\"rss_kib\":null,\"rss_source\":\"external /usr/bin/time -v\",\"memory_retained_source\":\"execution budget while result is live\",\"validation_scope\":\"one untimed public Value::Complex component/suffix check plus IMREAL/IMAGINARY or error oracle; timed evaluate/checksum/drop\"}}",
        json_escape(&case.name),
        case.function,
        case.group,
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
        expected_label(case.expected),
        expected_real,
        expected_imaginary,
        expected_suffix,
        expected_failure,
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
        work,
        mean(samples, |sample| sample.work),
        work / repeat as u64,
        percentile(samples, |sample| sample.successes, 50),
        percentile(samples, |sample| sample.refusals, 50),
        percentile(samples, |sample| sample.resolver.reads, 50),
        samples
            .iter()
            .map(|sample| sample.resolver.reads)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.resolver.extent_calls, 50),
        samples
            .iter()
            .map(|sample| sample.resolver.extent_calls)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.resolver.sheet_index_calls, 50),
        samples
            .iter()
            .map(|sample| sample.resolver.sheet_index_calls)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.resolver.sheet_name_calls, 50),
        samples
            .iter()
            .map(|sample| sample.resolver.sheet_name_calls)
            .max()
            .unwrap_or(0),
        percentile(samples, |sample| sample.resolver.sheet_count_calls, 50),
        samples
            .iter()
            .map(|sample| sample.resolver.sheet_count_calls)
            .max()
            .unwrap_or(0),
    );
}

fn selected_cases<'a>(
    all: &'a [CaseSpec],
    requested: Option<&str>,
    group: Option<&str>,
) -> AnyResult<Vec<&'a CaseSpec>> {
    let grouped: Vec<&CaseSpec> = all
        .iter()
        .filter(|case| group.is_none_or(|wanted| wanted == "all" || wanted == case.group))
        .collect();
    match requested {
        None | Some("all") => Ok(grouped),
        Some(name) => grouped
            .into_iter()
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

    // The raw result and its mathematical component checks are deliberately
    // outside the allocation/timing samples.
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

fn corpus() -> Vec<CaseSpec> {
    let z34 = ComplexPair::new(3.0, 4.0);
    let z12 = ComplexPair::new(1.0, 2.0);
    let zero = ComplexPair::new(0.0, 0.0);
    let mut cases = vec![
        case(
            "complex",
            "COMPLEX",
            "=COMPLEX(3;4)",
            1,
            "base",
            Expected::Complex(z34),
        ),
        case(
            "imabs",
            "IMABS",
            "=IMABS(COMPLEX(3;4))",
            1,
            "base",
            Expected::Number(z34.abs()),
        ),
        case(
            "imaginary",
            "IMAGINARY",
            "=IMAGINARY(COMPLEX(3;4))",
            1,
            "base",
            Expected::Number(4.0),
        ),
        case(
            "imargument",
            "IMARGUMENT",
            "=IMARGUMENT(COMPLEX(3;4))",
            1,
            "base",
            Expected::Number(z34.argument()),
        ),
        case(
            "imconjugate",
            "IMCONJUGATE",
            "=IMCONJUGATE(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(ComplexPair::new(3.0, -4.0)),
        ),
        case(
            "imcos",
            "IMCOS",
            "=IMCOS(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.cos()),
        ),
        case(
            "imcosh",
            "IMCOSH",
            "=IMCOSH(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.cosh()),
        ),
        case(
            "imcot",
            "IMCOT",
            "=IMCOT(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.cos().div(z34.sin())),
        ),
        case(
            "imcsc",
            "IMCSC",
            "=IMCSC(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.sin().reciprocal()),
        ),
        case(
            "imcsch",
            "IMCSCH",
            "=IMCSCH(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.sinh().reciprocal()),
        ),
        case(
            "imdiv",
            "IMDIV",
            "=IMDIV(COMPLEX(3;4);COMPLEX(1;2))",
            1,
            "base",
            Expected::Complex(z34.div(z12)),
        ),
        case(
            "imdiv-scaled-finite",
            "IMDIV",
            "=IMDIV(COMPLEX(1e308;1e308);COMPLEX(1;1))",
            2,
            "scaled",
            Expected::Complex(ComplexPair::new(1e308, 0.0)),
        ),
        case(
            "imexp",
            "IMEXP",
            "=IMEXP(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.exp()),
        ),
        case(
            "imln",
            "IMLN",
            "=IMLN(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.ln()),
        ),
        case(
            "imlog10",
            "IMLOG10",
            "=IMLOG10(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.ln().div(ComplexPair::new(10.0_f64.ln(), 0.0))),
        ),
        case(
            "imlog2",
            "IMLOG2",
            "=IMLOG2(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.ln().div(ComplexPair::new(2.0_f64.ln(), 0.0))),
        ),
        case(
            "impower",
            "IMPOWER",
            "=IMPOWER(COMPLEX(1;2);3)",
            3,
            "base",
            Expected::Complex(z12.powi(3)),
        ),
        case(
            "improduct",
            "IMPRODUCT",
            "=IMPRODUCT(COMPLEX(1;2);COMPLEX(3;4))",
            2,
            "base",
            Expected::Complex(z12.mul(z34)),
        ),
        case(
            "imreal",
            "IMREAL",
            "=IMREAL(COMPLEX(3;4))",
            1,
            "base",
            Expected::Number(3.0),
        ),
        case(
            "imsin",
            "IMSIN",
            "=IMSIN(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.sin()),
        ),
        case(
            "imsinh",
            "IMSINH",
            "=IMSINH(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.sinh()),
        ),
        case(
            "imsec",
            "IMSEC",
            "=IMSEC(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.cos().reciprocal()),
        ),
        case(
            "imsech",
            "IMSECH",
            "=IMSECH(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.cosh().reciprocal()),
        ),
        case(
            "imsqrt",
            "IMSQRT",
            "=IMSQRT(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.sqrt()),
        ),
        case(
            "imsub",
            "IMSUB",
            "=IMSUB(COMPLEX(3;4);COMPLEX(1;2))",
            1,
            "base",
            Expected::Complex(z34.sub(z12)),
        ),
        case(
            "imsum",
            "IMSUM",
            "=IMSUM(COMPLEX(3;4);COMPLEX(1;2))",
            2,
            "base",
            Expected::Complex(z34.add(z12)),
        ),
        case(
            "imtan",
            "IMTAN",
            "=IMTAN(COMPLEX(3;4))",
            1,
            "base",
            Expected::Complex(z34.sin().div(z34.cos())),
        ),
    ];
    assert_eq!(
        cases
            .iter()
            .filter(|candidate| candidate.group == "base")
            .count(),
        26,
        "the base corpus must cover every §6.8 function"
    );

    cases.extend([
        case(
            "complex-suffix-j",
            "COMPLEX",
            "=COMPLEX(3;4;\"j\")",
            1,
            "text",
            Expected::Complex(ComplexPair::with_suffix(3.0, 4.0, 'j')),
        ),
        case(
            "complex-text-j",
            "IMREAL",
            "=IMREAL(\"3+4j\")",
            1,
            "text",
            Expected::Number(3.0),
        ),
        case(
            "complex-imaginary-text",
            "IMAGINARY",
            "=IMAGINARY(\"4i\")",
            1,
            "text",
            Expected::Number(4.0),
        ),
        case(
            "complex-text-real",
            "IMREAL",
            "=IMREAL(\"3\")",
            1,
            "text",
            Expected::Number(3.0),
        ),
        case(
            "complex-nested-chain",
            "IMDIV",
            "=IMDIV(IMCONJUGATE(COMPLEX(3;4));COMPLEX(1;2))",
            2,
            "nested",
            Expected::Complex(ComplexPair::new(-1.0, -2.0)),
        ),
        case(
            "complex-nested-trig",
            "IMCOS",
            "=IMCOS(IMSUM(COMPLEX(1;2);COMPLEX(3;4)))",
            2,
            "nested",
            Expected::Complex(ComplexPair::new(4.0, 6.0).cos()),
        ),
        case(
            "lazy-selected-complex",
            "IF",
            "=IF(TRUE();COMPLEX(3;4);IMLN(0))",
            1,
            "lazy",
            Expected::Complex(z34),
        ),
        case(
            "lazy-unselected-complex",
            "IF",
            "=IF(FALSE();IMSUM([.A1:.A256]);COMPLEX(0;0))",
            256,
            "lazy",
            Expected::Complex(zero),
        ),
        case(
            "lazy-iferror-complex",
            "IFERROR",
            "=IFERROR(IMLN(0);COMPLEX(7;8))",
            1,
            "lazy",
            Expected::Complex(ComplexPair::new(7.0, 8.0)),
        ),
        case(
            "lazy-ifna-complex",
            "IFNA",
            "=IFNA(#N/A;COMPLEX(7;8))",
            1,
            "lazy",
            Expected::Complex(ComplexPair::new(7.0, 8.0)),
        ),
        case(
            "imsum-zero",
            "IMSUM",
            "=IMSUM()",
            0,
            "aggregate",
            Expected::Number(0.0),
        ),
        case(
            "improduct-zero",
            "IMPRODUCT",
            "=IMPRODUCT()",
            0,
            "error",
            Expected::FormulaError,
        ),
        case(
            "imsum-ignore-text",
            "IMSUM",
            "=IMSUM(\"bad\";COMPLEX(1;2))",
            2,
            "text",
            Expected::Complex(ComplexPair::new(1.0, 2.0)),
        ),
        case(
            "complex-error-invalid-text",
            "IMREAL",
            "=IMREAL(\"not-complex\")",
            1,
            "error",
            Expected::FormulaError,
        ),
        case(
            "complex-error-invalid-suffix",
            "COMPLEX",
            "=COMPLEX(1;2;\"k\")",
            1,
            "error",
            Expected::FormulaError,
        ),
        case(
            "complex-error-arity",
            "COMPLEX",
            "=COMPLEX(1)",
            1,
            "error",
            Expected::FormulaError,
        ),
        case(
            "complex-error-div-zero",
            "IMDIV",
            "=IMDIV(COMPLEX(1;2);0)",
            1,
            "error",
            Expected::FormulaError,
        ),
        case(
            "complex-error-ln-zero",
            "IMLN",
            "=IMLN(0)",
            0,
            "error",
            Expected::FormulaError,
        ),
        case(
            "complex-error-log10-zero",
            "IMLOG10",
            "=IMLOG10(0)",
            0,
            "error",
            Expected::FormulaError,
        ),
        case(
            "complex-error-log2-zero",
            "IMLOG2",
            "=IMLOG2(0)",
            0,
            "error",
            Expected::FormulaError,
        ),
        case(
            "complex-error-power-zero",
            "IMPOWER",
            "=IMPOWER(0;0)",
            0,
            "error",
            Expected::FormulaError,
        ),
        case(
            "complex-error-overflow",
            "IMEXP",
            "=IMEXP(1e308)",
            1,
            "error",
            Expected::FormulaError,
        ),
        case(
            "complex-error-product-text",
            "IMPRODUCT",
            "=IMPRODUCT(\"bad\";COMPLEX(1;0))",
            2,
            "error",
            Expected::FormulaError,
        ),
    ]);

    for size in [1, 4, 16, 64, 256, 1024] {
        let sum_expected = ComplexPair::new(size as f64, size as f64);
        cases.push(case(
            &format!("aggregate-sum-{size}"),
            "IMSUM",
            variadic("IMSUM", size, "COMPLEX(1;1)"),
            size,
            "aggregate",
            Expected::Complex(sum_expected),
        ));
        cases.push(case(
            &format!("aggregate-product-{size}"),
            "IMPRODUCT",
            variadic("IMPRODUCT", size, "COMPLEX(1;0)"),
            size,
            "aggregate",
            Expected::Complex(ComplexPair::new(1.0, 0.0)),
        ));
    }

    for size in [16, 256, 1024] {
        cases.push(case(
            &format!("reference-complex-sum-{size}"),
            "IMSUM",
            reference_range("IMSUM", size),
            size,
            "reference",
            Expected::Complex(ComplexPair::new(size as f64, size as f64)),
        ));
        cases.push(case(
            &format!("reference-complex-product-{size}"),
            "IMPRODUCT",
            reference_range("IMPRODUCT", size),
            size,
            "reference",
            Expected::Complex(ComplexPair::new(1.0, 0.0)),
        ));
    }
    cases.extend([
        case(
            "reference-sequence-specials",
            "IMSUM",
            "=IMSUM([.A1:.A3])",
            3,
            "reference",
            Expected::Complex(ComplexPair::new(1.0, 1.0)),
        ),
        case(
            "reference-sequence-error",
            "IMSUM",
            "=IMSUM([.A1:.A2])",
            2,
            "reference",
            Expected::FormulaError,
        ),
    ]);

    let long_digits = "9".repeat(4096);
    cases.push(case(
        "complex-long-text-4096",
        "IMREAL",
        format!("=IMREAL(\"{long_digits}i\")"),
        4096,
        "text",
        Expected::FormulaError,
    ));
    let over_limit = "x".repeat(32_768);
    cases.push(case(
        "complex-long-text-over-limit",
        "IMREAL",
        format!("=IMREAL(\"{over_limit}\")"),
        32_768,
        "text",
        Expected::EvaluationFailure("text"),
    ));

    cases.extend([
        case(
            "complex-refusal-work",
            "IMSUM",
            variadic("IMSUM", 256, "COMPLEX(1;1)"),
            256,
            "resource",
            Expected::EvaluationFailure("work"),
        ),
        case(
            "complex-refusal-memory",
            "IMREAL",
            "=IMREAL(\"3+4i\")",
            1,
            "resource",
            Expected::EvaluationFailure("memory"),
        ),
        case(
            "complex-refusal-cancelled",
            "IMSUM",
            variadic("IMSUM", 256, "COMPLEX(1;1)"),
            256,
            "resource",
            Expected::EvaluationFailure("cancelled"),
        ),
    ]);
    cases
}

fn main() -> AnyResult<()> {
    let Some(config) = parse_config()? else {
        return Ok(());
    };
    let all = corpus();
    for candidate in selected_cases(&all, config.case.as_deref(), config.group.as_deref())? {
        run_case(candidate, &config)?;
    }
    Ok(())
}
