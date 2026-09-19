//! Resource and capability-boundary coverage for the descriptive reducers.
//!
//! The resolver makes every cell read observable.  The tests consequently
//! lock charge-before-read order, cumulative variadic-array reservations,
//! two-pass centered reducers, source/cancellation fences, typed provider
//! failures, and the zero-read `NumberSequence` shape refusal.

use std::{
    cell::{Cell, RefCell},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource, SourceVersion,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits, Mode, OwnedValueView, Position, Resolver,
        SheetExtent, Value,
    },
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        UnsupportedKind, evaluate_scalar,
    },
    expression::Expression,
};

const RANGE: &str = "[.A1:.A8]";

#[derive(Debug, Clone)]
enum FixtureCell {
    Number(f64),
    Text(String),
    Error(ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct LimitsResolver {
    rows: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_coordinates: RefCell<Vec<(usize, usize)>>,
    cancel_after_read: Option<CancellationSource>,
    failure_at_read: Option<usize>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl LimitsResolver {
    fn standard(rows: usize) -> Self {
        Self {
            rows,
            cells: (0..rows)
                .map(|row| FixtureCell::Number((row + 1) as f64))
                .collect(),
            reads: Cell::new(0),
            read_coordinates: RefCell::new(Vec::new()),
            cancel_after_read: None,
            failure_at_read: None,
            source_versions: None,
            source_version_calls: Cell::new(0),
        }
    }

    fn set(&mut self, row: usize, value: FixtureCell) {
        self.cells[row] = value;
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn with_cancel_after_read(mut self, cancellation: &CancellationSource) -> Self {
        self.cancel_after_read = Some(cancellation.clone());
        self
    }

    fn with_failure_at_read(mut self, read: usize) -> Self {
        self.failure_at_read = Some(read);
        self
    }

    fn set_source_versions(&mut self, expected: SourceVersion, observed: SourceVersion) {
        self.source_versions = Some((expected, observed));
        self.source_version_calls.set(0);
    }

    fn harmonic_adversarial(with_residual: bool) -> Self {
        Self::harmonic_adversarial_with_residual(with_residual.then_some(1.0))
    }

    fn harmonic_adversarial_with_residual(residual: Option<f64>) -> Self {
        const PAIRS: usize = 100;
        let rows = if residual.is_some() {
            PAIRS * 2 + 1
        } else {
            PAIRS * 2
        };
        let mut resolver = Self::standard(rows);
        let base = (1_u64 << 52) as f64;
        for index in 0..PAIRS {
            // 2^52 minus an odd offset is an exactly represented distinct odd
            // integer. The positive and negative halves cancel exactly in the
            // reciprocal domain, while their odd denominators have an LCM
            // beyond the fast binary interval.
            let value = base - (2 * index + 1) as f64;
            resolver.set(index, FixtureCell::Number(value));
            resolver.set(index + PAIRS, FixtureCell::Number(-value));
        }
        if let Some(residual) = residual {
            resolver.set(rows - 1, FixtureCell::Number(residual));
        }
        resolver
    }

    fn harmonic_repeating(rows: usize) -> Self {
        let mut resolver = Self::standard(rows);
        for row in 0..rows {
            resolver.set(
                row,
                FixtureCell::Number(match row % 3 {
                    0 => 3.0,
                    1 => 6.0,
                    _ => -2.0,
                }),
            );
        }
        resolver
    }
}

impl Resolver for LimitsResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(SheetExtent::new(self.rows, 4)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        let read = self.reads.get().saturating_add(1);
        self.reads.set(read);
        self.read_coordinates.borrow_mut().push((row, column));
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        if self.failure_at_read == Some(read) {
            return Err(EvaluationFailure::Unsupported(UnsupportedKind::CellValue));
        }
        if sheet != "Main" || column != 0 || row >= self.rows {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        Ok(match &self.cells[row] {
            FixtureCell::Number(value) => CellRead::Number(*value),
            FixtureCell::Text(value) => CellRead::Text(value.as_str()),
            FixtureCell::Error(error) => CellRead::Error(*error),
            FixtureCell::Unsupported => CellRead::Unsupported,
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(1)
    }

    fn source_version(
        &self,
        _execution: &ExecutionContext,
    ) -> Result<Option<SourceVersion>, EvaluationFailure> {
        let Some((expected, observed)) = self.source_versions else {
            return Ok(None);
        };
        let call = self.source_version_calls.get();
        self.source_version_calls.set(call.saturating_add(1));
        Ok(Some(if call == 0 { expected } else { observed }))
    }
}

fn execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1024 * 1024).expect("one in-flight MiB"),
        0,
    )
    .expect("valid execution limits");
    (
        budget.clone(),
        cancellation,
        ExecutionContext::new(budget, token, limits),
    )
}

fn parse(source: &str) -> Expression {
    Expression::parse(source).unwrap_or_else(|error| panic!("{source:?} should parse: {error}"))
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum LimitResult {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn evaluate_source(
    source: &str,
    resolver: &LimitsResolver,
    execution: &ExecutionContext,
    limits: &Limits,
) -> Result<LimitResult, EvaluationFailure> {
    evaluate_source_mode(source, resolver, execution, Mode::Matrix, limits)
}

fn evaluate_source_mode(
    source: &str,
    resolver: &LimitsResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<LimitResult, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(match result.value() {
        Value::Number(value) => LimitResult::Number(value),
        Value::Error(error) => LimitResult::Error(error),
        _ => LimitResult::Other,
    })
}

fn evaluate_scalar_source(
    source: &str,
    execution: &ExecutionContext,
    limits: &EvaluationLimits,
) -> Result<LimitResult, EvaluationFailure> {
    let expression = parse(source);
    let result = evaluate_scalar(&expression, &EvaluationContext::new(execution), limits)?;
    Ok(match result.value() {
        ScalarValue::Number(value) => LimitResult::Number(*value),
        ScalarValue::Error(error) => LimitResult::Error(*error),
        _ => LimitResult::Other,
    })
}

fn evaluate<'a>(
    expression: &'a Expression,
    resolver: &'a LimitsResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn run_harmonic(
    source: &str,
    resolver: &LimitsResolver,
    execution_name: &str,
    storage: usize,
    max_reference_cells: usize,
) -> (Result<LimitResult, EvaluationFailure>, usize, u64) {
    let (budget, _cancellation, execution) = execution(execution_name);
    let result = evaluate_source(
        source,
        resolver,
        &execution,
        &Limits::default()
            .with_max_reference_cells(max_reference_cells)
            .with_max_storage_bytes(storage),
    );
    (result, resolver.reads(), budget.used(Resource::Memory))
}

fn harmonic_storage_threshold<F>(mut succeeds: F) -> usize
where
    F: FnMut(usize) -> bool,
{
    if succeeds(0) {
        return 0;
    }
    let mut high = 1usize;
    while !succeeds(high) {
        high = high
            .checked_mul(2)
            .expect("harmonic storage threshold should fit usize");
    }
    let mut low = 0usize;
    while high - low > 1 {
        let middle = low + (high - low) / 2;
        if succeeds(middle) {
            high = middle;
        } else {
            low = middle;
        }
    }
    high
}

fn harmonic_literal_cancellation_source() -> String {
    const PAIRS: usize = 100;
    let base = (1_u64 << 52) as f64;
    let mut values = Vec::with_capacity(PAIRS * 2);
    for index in 0..PAIRS {
        let value = base - (2 * index + 1) as f64;
        values.push(value.to_string());
    }
    for index in 0..PAIRS {
        let value = base - (2 * index + 1) as f64;
        values.push((-value).to_string());
    }
    format!("=HARMEAN({})", values.join(";"))
}

fn run_scalar_harmonic(
    source: &str,
    execution_name: &str,
    storage: usize,
) -> (Result<LimitResult, EvaluationFailure>, u64) {
    let (budget, _cancellation, execution) = execution(execution_name);
    let result = evaluate_scalar_source(
        source,
        &execution,
        &EvaluationLimits::default().with_max_storage_bytes(storage),
    );
    (result, budget.used(Resource::Memory))
}

fn assert_number(result: LimitResult, expected: f64, source: &str) {
    match result {
        LimitResult::Number(actual) => assert_eq!(actual, expected, "{source:?}"),
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

fn assert_close(result: LimitResult, expected: f64, source: &str) {
    let LimitResult::Number(actual) = result else {
        panic!("{source:?}: expected Number({expected}), got {result:?}");
    };
    let tolerance = expected.abs().max(actual.abs()).max(1.0) * 128.0 * f64::EPSILON;
    assert!(
        actual.is_finite() && (actual - expected).abs() <= tolerance,
        "{source:?}: {actual} != {expected} (tolerance {tolerance})"
    );
}

fn assert_finite_number(result: LimitResult, source: &str) {
    let LimitResult::Number(actual) = result else {
        panic!("{source:?}: expected a finite Number, got {result:?}");
    };
    assert!(actual.is_finite(), "{source:?}: non-finite result {actual}");
}

#[test]
fn reference_cell_limit_is_charged_before_descriptive_reads() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-descriptive-cell-limit");
    let error = evaluate_source(
        &format!("=AVEDEV({RANGE})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("zero reference-cell budget must reject descriptive ranges");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference-cell failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn work_array_and_storage_limits_are_atomic() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, work_execution) = execution("ods-formula-descriptive-work-limit");
    let error = evaluate_source(
        &format!("=DEVSQ({RANGE})"),
        &resolver,
        &work_execution,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero Work must refuse before descriptive scanning");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, array_execution) =
        execution("ods-formula-descriptive-array-limit");
    let error = evaluate_source(
        "=AVEDEV({1;2|3;4})",
        &resolver,
        &array_execution,
        &Limits::default().with_max_array_cells(3),
    )
    .expect_err("array-cell limit must reject before reduction");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);

    let error = evaluate_source(
        "=DEVSQ({1;2|3;4})",
        &resolver,
        &array_execution,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero evaluator storage must reject array admission");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn variadic_descriptive_arrays_use_cumulative_admitted_capacity() {
    let resolver = LimitsResolver::standard(1);
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-variadic-arrays");
    let limits = Limits::default()
        .with_max_array_cells(2)
        .with_max_reference_cells(1)
        .with_max_storage_bytes(64 * 1024);

    let result = evaluate_source("=AVEDEV({1;3};{5;7})", &resolver, &execution, &limits)
        .expect("AVEDEV should admit both two-cell sequence arrays");
    assert_number(result, 2.0, "variadic AVEDEV arrays");

    let result = evaluate_source("=DEVSQ({1;3};{5;7})", &resolver, &execution, &limits)
        .expect("DEVSQ should admit both two-cell sequence arrays");
    assert_close(result, 20.0, "variadic DEVSQ arrays");
}

#[test]
fn referenced_text_is_omitted_by_number_sequence_conversion() {
    let mut resolver = LimitsResolver::standard(2);
    resolver.set(0, FixtureCell::Text("not-a-number".to_owned()));
    resolver.set(1, FixtureCell::Number(3.0));
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-text");
    let result = evaluate_source(
        "=DEVSQ([.A1:.A2])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("referenced Text should be omitted by NumberSequence");
    assert_number(result, 0.0, "referenced Text omission");
    assert_eq!(resolver.reads(), 2);
}

#[test]
fn descriptive_reducers_stream_large_references_with_function_specific_reads() {
    let rows = 20_000;
    let limits = Limits::default()
        .with_max_reference_cells(rows * 2)
        .with_max_array_cells(1)
        .with_max_storage_bytes(64 * 1024);

    let resolver = LimitsResolver::standard(rows);
    let (_budget, _cancellation, avedev_execution) =
        execution("ods-formula-descriptive-avedev-stream");
    let result = evaluate_source_mode(
        "=AVEDEV([.A1:.A20000])",
        &resolver,
        &avedev_execution,
        Mode::Scalar,
        &limits,
    )
    .expect("AVEDEV should stream a large reference");
    assert_close(result, 5000.0, "large AVEDEV reference");
    assert_eq!(
        resolver.reads(),
        rows * 2,
        "AVEDEV should perform two scans"
    );

    let resolver = LimitsResolver::standard(rows);
    let (_budget, _cancellation, devsq_execution) =
        execution("ods-formula-descriptive-devsq-stream");
    let result = evaluate_source_mode(
        "=DEVSQ([.A1:.A20000])",
        &resolver,
        &devsq_execution,
        Mode::Scalar,
        &limits,
    )
    .expect("DEVSQ should stream a large reference");
    assert_close(result, 666_666_665_000.0, "large DEVSQ reference");
    assert_eq!(
        resolver.reads(),
        rows,
        "DEVSQ exact moments should complete in one scan"
    );

    for function in ["KURT", "SKEW", "SKEWP"] {
        let resolver = LimitsResolver::standard(rows);
        let scope = format!("ods-formula-descriptive-{}-stream", function.to_lowercase());
        let (_budget, _cancellation, function_execution) = execution(&scope);
        let source = format!("={function}([.A1:.A20000])");
        let result = evaluate_source_mode(
            &source,
            &resolver,
            &function_execution,
            Mode::Scalar,
            &limits,
        )
        .unwrap_or_else(|error| panic!("{source} should stream: {error:?}"));
        assert_finite_number(result, &source);
        assert_eq!(
            resolver.reads(),
            rows,
            "{function} exact moments should complete in one scan"
        );
    }
}

#[test]
fn harmonic_exact_replay_admits_only_bounded_limb_storage() {
    const PAIRS: usize = 100;
    const CANCEL_ROWS: usize = PAIRS * 2;
    const RESIDUAL_ROWS: usize = CANCEL_ROWS + 1;
    let cancel_source = format!("=HARMEAN([.A1:.A{CANCEL_ROWS}])");
    let residual_source = format!("=HARMEAN([.A1:.A{RESIDUAL_ROWS}])");
    let max_reference_cells = RESIDUAL_ROWS * 2;

    // The residual-one vector is a metadata-admitting baseline: its fast
    // interval proves 201, so it does not need the exact odd-denominator
    // replay.  Find that baseline rather than baking in a descriptor size.
    let metadata_storage = harmonic_storage_threshold(|storage| {
        let resolver = LimitsResolver::harmonic_adversarial(true);
        let (result, _reads, _memory) = run_harmonic(
            &residual_source,
            &resolver,
            "ods-formula-descriptive-harmonic-metadata-threshold",
            storage,
            max_reference_cells,
        );
        matches!(result, Ok(LimitResult::Number(value)) if (value - 201.0).abs() <= 201.0 * 1e-12)
    });

    // The all-pairs cancellation forces the exact replay.  Determine its
    // actual admission threshold from the implementation, then test one byte
    // below it while retaining enough storage for the reference metadata.
    let exact_storage = harmonic_storage_threshold(|storage| {
        let resolver = LimitsResolver::harmonic_adversarial(false);
        let (result, _reads, _memory) = run_harmonic(
            &cancel_source,
            &resolver,
            "ods-formula-descriptive-harmonic-exact-threshold",
            storage,
            max_reference_cells,
        );
        matches!(result, Ok(LimitResult::Error(ScalarError::DivisionByZero)))
    });
    assert!(
        exact_storage > metadata_storage,
        "exact replay must require additional bounded storage"
    );
    let tight_storage = exact_storage - 1;
    assert!(
        tight_storage >= metadata_storage,
        "tight exact-replay cap must still admit reference metadata"
    );

    // A residual 1e100 has reciprocal 1e-100, below the first-pass
    // uncertainty of the cancellation-heavy prefix. It therefore exercises
    // the exact replay while still finishing with a positive finite result.
    let forced_residual_source = format!("=HARMEAN([.A1:.A{RESIDUAL_ROWS}])");
    let forced_exact_storage = harmonic_storage_threshold(|storage| {
        let resolver = LimitsResolver::harmonic_adversarial_with_residual(Some(1e100));
        let (result, _reads, _memory) = run_harmonic(
            &forced_residual_source,
            &resolver,
            "ods-formula-descriptive-harmonic-forced-threshold",
            storage,
            max_reference_cells,
        );
        matches!(
            result,
            Ok(LimitResult::Number(value))
                if (value - 201e100).abs() <= 201e100 * 1e-12
        )
    });
    let resolver = LimitsResolver::harmonic_adversarial_with_residual(Some(1e100));
    let (result, reads, memory) = run_harmonic(
        &forced_residual_source,
        &resolver,
        "ods-formula-descriptive-harmonic-forced",
        forced_exact_storage,
        max_reference_cells,
    );
    assert_close(
        result.expect("adaptive exact replay should finish positively"),
        201e100,
        &forced_residual_source,
    );
    assert_eq!(
        reads,
        RESIDUAL_ROWS * 2,
        "forced tiny reciprocal should replay the reference"
    );
    assert_eq!(memory, 0, "positive exact replay must drop temporary state");

    let resolver = LimitsResolver::harmonic_adversarial(false);
    let (result, _reads, memory) = run_harmonic(
        &cancel_source,
        &resolver,
        "ods-formula-descriptive-harmonic-generous",
        exact_storage,
        max_reference_cells,
    );
    assert!(
        matches!(result, Ok(LimitResult::Error(ScalarError::DivisionByZero))),
        "generous exact replay should publish the exact zero reciprocal sum"
    );
    assert_eq!(
        memory, 0,
        "successful replay must drop temporary limb state"
    );

    let resolver = LimitsResolver::harmonic_adversarial(true);
    let (result, _reads, memory) = run_harmonic(
        &residual_source,
        &resolver,
        "ods-formula-descriptive-harmonic-residual",
        tight_storage,
        max_reference_cells,
    );
    assert_number(
        result.expect("residual-one harmonic case should remain numeric"),
        201.0,
        &residual_source,
    );
    assert_eq!(memory, 0, "residual baseline must drop temporary state");

    let resolver = LimitsResolver::harmonic_adversarial(false);
    let (result, _reads, memory) = run_harmonic(
        &cancel_source,
        &resolver,
        "ods-formula-descriptive-harmonic-tight",
        tight_storage,
        max_reference_cells,
    );
    assert!(matches!(
        result,
        Err(EvaluationFailure::ResourceLimit(limit)) if limit.resource == Resource::Memory
    ));
    assert_eq!(
        memory, 0,
        "failed exact replay must refund all limb storage"
    );

    let resolver = LimitsResolver::harmonic_adversarial(false);
    let (result, _reads, memory) = run_harmonic(
        &format!("=IFERROR(HARMEAN([.A1:.A{CANCEL_ROWS}]);7)"),
        &resolver,
        "ods-formula-descriptive-harmonic-iferror",
        tight_storage,
        max_reference_cells,
    );
    assert!(matches!(
        result,
        Err(EvaluationFailure::ResourceLimit(limit)) if limit.resource == Resource::Memory
    ));
    assert_eq!(memory, 0, "IFERROR must not retain refused limb storage");
}

#[test]
fn harmonic_replay_precision_is_independent_of_repeated_reference_count() {
    const ROWS: usize = 20_001;
    let resolver = LimitsResolver::harmonic_repeating(ROWS);
    let (budget, _cancellation, execution) =
        execution("ods-formula-descriptive-harmonic-repeating");
    let limits = Limits::default()
        .with_max_reference_cells(ROWS * 2)
        .with_max_array_cells(1)
        .with_max_storage_bytes(64 * 1024)
        .with_max_steps(100_000_000);
    let result = evaluate_source("=HARMEAN([.A1:.A20001])", &resolver, &execution, &limits)
        .expect("repeated reciprocal cancellation should fit bounded precision storage");
    assert_eq!(
        result,
        LimitResult::Error(ScalarError::DivisionByZero),
        "repeated exact reciprocal cancellation should publish #DIV/0!"
    );
    assert_eq!(
        resolver.reads(),
        ROWS * 2,
        "exact harmonic fallback should replay the streamed reference"
    );
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "repeated harmonic fallback must refund its temporary precision state"
    );
}

#[test]
fn scalar_harmonic_exact_replay_resource_refusal_is_typed_and_uncatchable() {
    let source = harmonic_literal_cancellation_source();
    let exact_storage = harmonic_storage_threshold(|storage| {
        let (result, _memory) = run_scalar_harmonic(
            &source,
            "ods-formula-descriptive-scalar-harmonic-threshold",
            storage,
        );
        matches!(result, Ok(LimitResult::Error(ScalarError::DivisionByZero)))
    });
    assert!(
        exact_storage > 1,
        "exact scalar replay should allocate limbs"
    );
    let tight_storage = exact_storage - 1;

    let (result, memory) = run_scalar_harmonic(
        &source,
        "ods-formula-descriptive-scalar-harmonic-generous",
        exact_storage,
    );
    assert!(
        matches!(result, Ok(LimitResult::Error(ScalarError::DivisionByZero))),
        "generous scalar replay should publish the exact zero reciprocal sum"
    );
    assert_eq!(memory, 0, "scalar replay must drop temporary limb state");

    let (result, memory) = run_scalar_harmonic(
        &source,
        "ods-formula-descriptive-scalar-harmonic-tight",
        tight_storage,
    );
    assert!(matches!(
        result,
        Err(EvaluationFailure::ResourceLimit(limit)) if limit.resource == Resource::Memory
    ));
    assert_eq!(memory, 0, "scalar refusal must refund exact replay storage");

    let iferror_source = format!(
        "=IFERROR({};7)",
        source.strip_prefix('=').unwrap_or(source.as_str())
    );
    let (result, memory) = run_scalar_harmonic(
        &iferror_source,
        "ods-formula-descriptive-scalar-harmonic-iferror",
        tight_storage,
    );
    assert!(matches!(
        result,
        Err(EvaluationFailure::ResourceLimit(limit)) if limit.resource == Resource::Memory
    ));
    assert_eq!(
        memory, 0,
        "scalar IFERROR refusal must refund exact replay storage"
    );
}

#[test]
fn later_typed_failure_is_visible_from_a_centered_second_pass() {
    let rows = 8;
    let resolver = LimitsResolver::standard(rows).with_failure_at_read(rows + 1);
    let (_budget, _cancellation, execution) =
        execution("ods-formula-descriptive-second-pass-failure");
    let error = evaluate_source(
        &format!("=AVEDEV({RANGE})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(rows * 2),
    )
    .expect_err("the second-pass provider failure must remain typed");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), rows + 1);
}

#[test]
fn cumulative_reference_budget_refuses_a_centered_replay_after_first_pass() {
    let rows = 8;
    let resolver = LimitsResolver::standard(rows);
    let (_budget, _cancellation, execution) =
        execution("ods-formula-descriptive-cumulative-reference-limit");
    let error = evaluate_source(
        &format!("=AVEDEV({RANGE})"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(rows),
    )
    .expect_err("the second centered pass must consume the cumulative reference budget");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), rows);
}

#[test]
fn typed_provider_failures_escape_all_descriptive_variants() {
    let mut resolver = LimitsResolver::standard(8);
    resolver.set(0, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-unsupported");
    for source in [
        "=AVEDEV([.A1:.A8])",
        "=DEVSQ([.A1:.A8])",
        "=GEOMEAN([.A1:.A8])",
        "=HARMEAN([.A1:.A8])",
        "=KURT([.A1:.A8])",
        "=SKEW([.A1:.A8])",
        "=SKEWP([.A1:.A8])",
    ] {
        let error = evaluate_source(source, &resolver, &execution, &Limits::default())
            .expect_err("unsupported provider cell must remain typed");
        assert!(
            matches!(
                error,
                EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
            ),
            "{source}: wrong provider failure {error:?}"
        );
    }
}

#[test]
fn later_typed_failure_supersedes_a_retained_formula_error() {
    let mut resolver = LimitsResolver::standard(8);
    resolver.set(0, FixtureCell::Error(ScalarError::NotAvailable));
    resolver.set(1, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-error-precedence");
    for source in [
        "=AVEDEV([.A1:.A8])",
        "=DEVSQ([.A1:.A8])",
        "=GEOMEAN([.A1:.A8])",
        "=HARMEAN([.A1:.A8])",
        "=KURT([.A1:.A8])",
        "=SKEW([.A1:.A8])",
        "=SKEWP([.A1:.A8])",
    ] {
        let error = evaluate_source(source, &resolver, &execution, &Limits::default())
            .expect_err("typed failure must supersede retained formula error");
        assert!(
            matches!(
                error,
                EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
            ),
            "{source}: wrong precedence result {error:?}"
        );
    }
}

#[test]
fn cancellation_after_first_descriptive_read_is_atomic() {
    let (budget, cancellation, execution) = execution("ods-formula-descriptive-cancel");
    let resolver = LimitsResolver::standard(64).with_cancel_after_read(&cancellation);
    let error = evaluate_source(
        "=DEVSQ([.A1:.A64])",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("provider cancellation must fence descriptive publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn source_version_change_after_descriptive_reads_is_not_published() {
    let mut resolver = LimitsResolver::standard(8);
    let expected = SourceVersion::new(0x4453_4352, 0);
    let observed = SourceVersion::new(0x4453_4352, 1);
    resolver.set_source_versions(expected, observed);
    let (budget, _cancellation, execution) = execution("ods-formula-descriptive-source-change");
    let error = evaluate_source(
        &format!("=DEVSQ({RANGE})"),
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect_err("source change must fence descriptive publication");
    assert!(matches!(
        error,
        EvaluationFailure::SourceChanged {
            expected: got_expected,
            observed: got_observed
        } if got_expected == expected && got_observed == observed
    ));
    assert!(resolver.reads() > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn unselected_descriptive_branch_does_not_probe_provider() {
    let resolver = LimitsResolver::standard(8);
    let (budget, _cancellation, execution) = execution("ods-formula-descriptive-lazy");
    let result = evaluate_source(
        "=IF(FALSE();DEVSQ([Missing.A1:.Z100]);0)",
        &resolver,
        &execution,
        &Limits::default(),
    )
    .expect("unselected descriptive branch");
    assert_number(result, 0.0, "lazy descriptive branch");
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn iferror_cannot_catch_a_descriptive_resource_refusal() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-iferror-limit");
    let error = evaluate_source(
        &format!("=IFERROR(DEVSQ({RANGE});7)"),
        &resolver,
        &execution,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("IFERROR must not catch a descriptive resource refusal");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn number_sequence_reference_list_refusal_has_zero_reads() {
    let resolver = LimitsResolver::standard(8);
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-list-refusal");
    for function in ["DEVSQ", "SKEWP"] {
        let source = format!("={function}([.A1:.A2]~[.A3:.A4])");
        let result = evaluate_source(&source, &resolver, &execution, &Limits::default())
            .expect("reference-list shape refusal should be a formula value");
        assert_eq!(result, LimitResult::Error(ScalarError::Value), "{source}");
        assert_eq!(resolver.reads(), 0, "{source} must refuse before reads");
    }
}

#[test]
fn successful_descriptive_reference_results_can_be_owned() {
    let resolver = LimitsResolver::standard(4);
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-owned");
    let expression = parse("=AVEDEV([.A1:.A4])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("descriptive reference result");
    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("descriptive result should be ownable");
    assert!(matches!(owned.value(), OwnedValueView::Number(_)));
}
