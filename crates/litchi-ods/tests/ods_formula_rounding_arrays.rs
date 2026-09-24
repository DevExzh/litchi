//! Matrix-mode integration coverage for the OpenFormula 1.4 §6.17 rounding
//! family.
//!
//! Numeric vectors are independently derived from the section's
//! directed-rounding rules.  The file also checks shape/broadcasting,
//! per-element coercion and errors, lazy selection, and caller-owned limits.

use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationFailure, EvaluationResult, ScalarError,
        value::{
            self, CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, SheetExtent,
            Value,
        },
    },
    expression::Expression,
};

#[derive(Debug)]
struct RoundingResolver {
    reads: std::cell::Cell<usize>,
}

impl RoundingResolver {
    fn new() -> Self {
        Self {
            reads: std::cell::Cell::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }
}

impl Resolver for RoundingResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<SheetExtent>> {
        Ok((sheet == "Main").then_some(SheetExtent::new(1, 2)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<CellRead<'a>> {
        self.reads.set(self.reads.get().saturating_add(1));
        if sheet != "Main" || row != 0 || column >= 2 {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        // A blank cell is an evaluated Empty value.  It must be converted to
        // significance zero, unlike CEILING/FLOOR's syntactic `;;` default.
        Ok(if column == 0 {
            CellRead::Empty
        } else {
            CellRead::Number(2.0)
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<usize>> {
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<&str>> {
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> EvaluationResult<usize> {
        Ok(1)
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

fn evaluate<'a>(
    expression: &'a Expression,
    resolver: &'a RoundingResolver,
    execution: &ExecutionContext,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_array_close(result: &Evaluated<'_>, rows: usize, columns: usize, expected: &[f64]) {
    let array = result
        .as_array()
        .expect("rounding function should return an array");
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (rows, columns)
    );
    assert_eq!(array.len(), expected.len());
    for (index, expected) in expected.iter().copied().enumerate() {
        match array.get(index).expect("array cell should be present") {
            Value::Number(actual) if expected == 0.0 => {
                assert_eq!(actual, 0.0, "array cell {index}: {actual} != {expected}");
            },
            Value::Number(actual) => {
                assert!(
                    actual.is_finite(),
                    "array cell {index}: non-finite result {actual}"
                );
                let tolerance = expected.abs().max(actual.abs()) * (16.0 * f64::EPSILON);
                assert!(
                    (actual - expected).abs() <= tolerance,
                    "array cell {index}: {actual} != {expected}"
                );
            },
            other => panic!("expected Number({expected}) at cell {index}, got {other:?}"),
        }
    }
}

fn assert_array_with_error(
    result: &Evaluated<'_>,
    rows: usize,
    columns: usize,
    expected: &[Result<f64, ScalarError>],
) {
    let array = result
        .as_array()
        .expect("rounding function should return an array");
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (rows, columns)
    );
    assert_eq!(array.len(), expected.len());
    for (index, expected) in expected.iter().enumerate() {
        match (
            array.get(index).expect("array cell should be present"),
            expected,
        ) {
            (Value::Number(actual), Ok(expected)) if *expected == 0.0 => {
                assert_eq!(actual, 0.0, "array cell {index}: {actual} != {expected}");
            },
            (Value::Number(actual), Ok(expected)) => {
                assert!(
                    actual.is_finite(),
                    "array cell {index}: non-finite result {actual}"
                );
                let tolerance = expected.abs().max(actual.abs()) * (16.0 * f64::EPSILON);
                assert!(
                    (actual - expected).abs() <= tolerance,
                    "array cell {index}: {actual} != {expected}"
                );
            },
            (Value::Error(actual), Err(expected)) => assert_eq!(actual, *expected),
            (actual, expected) => {
                panic!("array cell {index}: got {actual:?}, expected {expected:?}")
            },
        }
    }
}

#[test]
fn every_rounding_function_broadcasts_over_a_numeric_matrix() {
    let resolver = RoundingResolver::new();
    let (_budget, _cancellation, execution) = execution("ods-formula-rounding-array-family");
    let cases = [
        ("=CEILING({1.2;2.1|3.9;4.0};2)", [2.0, 4.0, 4.0, 4.0]),
        ("=INT({1.9;2.1|-1.1;0})", [1.0, 2.0, -2.0, 0.0]),
        ("=FLOOR({1.2;2.1|3.9;4.0};2)", [0.0, 2.0, 2.0, 4.0]),
        ("=MROUND({1.2;2.9|4.1;5.0};2)", [2.0, 2.0, 4.0, 6.0]),
        ("=ROUND({1.25;2.35|-1.25;-2.35};1)", [1.3, 2.4, -1.3, -2.4]),
        (
            "=ROUNDDOWN({1.29;2.35|-1.29;-2.35};1)",
            [1.2, 2.3, -1.2, -2.3],
        ),
        (
            "=ROUNDUP({1.21;2.35|-1.21;-2.35};1)",
            [1.3, 2.4, -1.3, -2.4],
        ),
        (
            "=TRUNC({123.45;234.56|-123.45;-234.56};-1)",
            [120.0, 230.0, -120.0, -230.0],
        ),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(&expression, &resolver, &execution, &Limits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_array_close(&result, 2, 2, &expected);
    }
    assert_eq!(
        resolver.reads(),
        0,
        "literal arrays need no worksheet reads"
    );
}

#[test]
fn array_rounding_coerces_text_and_logicals_per_element() {
    let resolver = RoundingResolver::new();
    let (_budget, _cancellation, execution) = execution("ods-formula-rounding-array-coercion");

    let expression = parse("=ROUND({\"2.5\";TRUE()|FALSE();\"bad\"};0)");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_with_error(
        &result,
        2,
        2,
        &[Ok(3.0), Ok(1.0), Ok(0.0), Err(ScalarError::Value)],
    );

    // Optional arguments use the same element-wise coercion and preserve the
    // input rectangle while significance zero remains a defined zero result.
    let expression = parse("=CEILING({\"1.2\";2.1|3.9;4.0};{\"1\";TRUE()|FALSE();\"2\"})");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_close(&result, 2, 2, &[2.0, 3.0, 0.0, 4.0]);
}

#[test]
fn significance_digits_and_modes_broadcast_elementwise_without_losing_shape() {
    let resolver = RoundingResolver::new();
    let (_budget, _cancellation, execution) = execution("ods-formula-rounding-array-broadcast");

    let expression = parse("=CEILING({1.2;2.2|3.2;4.2};{1;2|1;2})");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_close(&result, 2, 2, &[2.0, 4.0, 4.0, 6.0]);

    let expression = parse("=ROUND({1.25;2.25|3.25;4.25};{0;1|0;1})");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_close(&result, 2, 2, &[1.0, 2.3, 3.0, 4.3]);

    let expression = parse("=CEILING({1.2;-1.2};;0.5)");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_close(&result, 1, 2, &[2.0, -2.0]);
}

#[test]
fn syntactic_empty_significance_differs_from_a_blank_cell() {
    let resolver = RoundingResolver::new();
    let (_budget, _cancellation, execution) = execution("ods-formula-rounding-empty");

    let expression = parse("=CEILING({4.2;4.2};;)");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_close(&result, 1, 2, &[5.0, 5.0]);

    let expression = parse("=CEILING({4.2;4.2};[.A1])");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_close(&result, 1, 2, &[0.0, 0.0]);
    assert_eq!(
        resolver.reads(),
        1,
        "the scalar blank-cell significance is read once"
    );

    let expression = parse("=ROUND({4.2;4.8};[.A1])");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_close(&result, 1, 2, &[4.0, 5.0]);
    assert_eq!(
        resolver.reads(),
        2,
        "the blank-cell digits argument is read once"
    );
}

#[test]
fn scalar_rounding_uses_the_current_cell_for_reference_intersection() {
    let resolver = RoundingResolver::new();
    let (_budget, _cancellation, execution) = execution("ods-formula-rounding-scalar-intersection");
    let expression = parse("=ROUND([.A1:.B1];0)");

    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Scalar);
    let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
        .expect("scalar mode should project the range at the current cell");
    match result.value() {
        Value::Number(value) => assert_eq!(value.to_bits(), 0),
        value => panic!("expected projected blank as +0, got {value:?}"),
    }

    let context = Context::new(&execution, Position::new("Main", 0, 1)).with_mode(Mode::Scalar);
    let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
        .expect("scalar mode should project the range at the current cell");
    match result.value() {
        Value::Number(value) => assert_eq!(value, 2.0),
        value => panic!("expected projected B1 value, got {value:?}"),
    }
    assert_eq!(resolver.reads(), 2);
}

#[test]
fn matrix_if_visits_only_the_selected_rounding_branch() {
    let resolver = RoundingResolver::new();
    let (_budget, _cancellation, execution) = execution("ods-formula-rounding-lazy");
    let expression = parse("=IF({TRUE();TRUE()};ROUND({1.5;2.5};0);ROUND([Missing.A1:.Z100];0))");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default())
        .expect("the unselected missing-sheet branch must not be resolved");
    assert_array_close(&result, 1, 2, &[2.0, 3.0]);
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn matrix_rounding_preserves_element_errors_and_near_overflow_numbers() {
    let resolver = RoundingResolver::new();
    let (_budget, _cancellation, execution) = execution("ods-formula-rounding-array-errors");

    let expression = parse("=ROUND({1.5;#N/A|2.5;3.5};0)");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_with_error(
        &result,
        2,
        2,
        &[Ok(2.0), Err(ScalarError::NotAvailable), Ok(3.0), Ok(4.0)],
    );

    let expression = parse("=TRUNC({1.7976931348623155e308;0};0)");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_close(&result, 1, 2, &[1.7976931348623155e308, 0.0]);
}

#[test]
fn matrix_rounding_preserves_first_formula_error_and_refuses_zero_work() {
    let resolver = RoundingResolver::new();
    let (_budget, _cancellation, execution) = execution("ods-formula-rounding-array-error-order");

    let expression = parse("=CEILING({#N/A;#DIV/0!};2)");
    let result = evaluate(&expression, &resolver, &execution, &Limits::default()).unwrap();
    assert_array_with_error(
        &result,
        1,
        2,
        &[
            Err(ScalarError::NotAvailable),
            Err(ScalarError::DivisionByZero),
        ],
    );

    let expression = parse("=ROUND({1.5;2.5};0)");
    let limits = Limits::default().with_max_steps(0);
    let error = evaluate(&expression, &resolver, &execution, &limits)
        .expect_err("zero work must refuse before publishing a matrix result");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "unexpected zero-work failure: {error}"
    );
}

#[test]
fn matrix_rounding_honors_cancellation_before_evaluating_cells() {
    let resolver = RoundingResolver::new();
    let (_budget, cancellation, execution) = execution("ods-formula-rounding-array-cancel");
    cancellation.cancel();
    let expression = parse("=ROUND({1.5;2.5};0)");
    let error = evaluate(&expression, &resolver, &execution, &Limits::default())
        .expect_err("pre-cancelled rounding must refuse");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 0);
}
