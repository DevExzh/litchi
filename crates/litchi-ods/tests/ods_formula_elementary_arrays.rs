//! Independent matrix/reference integration coverage for the OpenFormula 1.4
//! elementary mathematical functions.
//!
//! This file checks element-wise definitions, scalar broadcasting, coercion,
//! references, implicit intersection, lazy branches, ownership, and the
//! value evaluator's shape/work/reference/storage/cancellation limits.

use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        CellRead, Context, Evaluated, Limits, Mode, OwnedValueView, Position, Resolver,
        SheetExtent, Value, evaluate,
    },
    evaluation::{EvaluationFailure, ScalarError},
    expression::Expression,
};

const PI: f64 = std::f64::consts::PI;

#[derive(Clone, Copy, Debug)]
enum FixtureCell {
    Number(f64),
}

#[derive(Debug)]
struct ElementaryResolver {
    cells: [FixtureCell; 4],
    reads: AtomicUsize,
}

impl ElementaryResolver {
    fn new(cells: [FixtureCell; 4]) -> Self {
        Self {
            cells,
            reads: AtomicUsize::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Acquire)
    }
}

impl Resolver for ElementaryResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(SheetExtent::new(2, 2)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.fetch_add(1, Ordering::AcqRel);
        if sheet != "Main" || row >= 2 || column >= 2 {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        Ok(match self.cells[row * 2 + column] {
            FixtureCell::Number(value) => CellRead::Number(value),
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
}

fn execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).expect("one worker"),
        std::num::NonZeroUsize::new(1).expect("one in-flight task"),
        std::num::NonZeroU64::new(1024 * 1024).expect("one in-flight MiB"),
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

fn evaluate_matrix<'a>(
    expression: &'a Expression,
    resolver: &'a ElementaryResolver,
    execution: &ExecutionContext,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    evaluate(expression, resolver, &context, limits)
}

fn evaluate_at<'a>(
    expression: &'a Expression,
    resolver: &'a ElementaryResolver,
    execution: &ExecutionContext,
    position: Position<'a>,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    evaluate(expression, resolver, &context, limits)
}

fn assert_close(actual: f64, expected: f64, label: &str) {
    if expected == 0.0 {
        assert_eq!(actual.to_bits(), expected.to_bits(), "{label}: {actual:?}");
        return;
    }
    assert!(actual.is_finite(), "{label} returned non-finite {actual:?}");
    let tolerance = expected.abs().max(actual.abs()) * (32.0 * f64::EPSILON);
    assert!(
        (actual - expected).abs() <= tolerance,
        "{label}: {actual:?} != {expected:?} (tol {tolerance:e})"
    );
}

fn assert_array(
    result: &Evaluated<'_>,
    rows: usize,
    columns: usize,
    expected: &[Result<f64, ScalarError>],
    label: &str,
) {
    let array = result
        .as_array()
        .unwrap_or_else(|| panic!("{label} should return an array"));
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (rows, columns)
    );
    assert_eq!(array.len(), expected.len(), "{label} cell count");
    for (index, expected) in expected.iter().enumerate() {
        let value = array
            .get(index)
            .unwrap_or_else(|| panic!("{label} missing cell {index}"));
        match (value, expected) {
            (Value::Number(actual), Ok(expected)) => {
                assert_close(actual, *expected, &format!("{label}[{index}]"));
            },
            (Value::Error(actual), Err(expected)) => {
                assert_eq!(actual, *expected, "{label}[{index}]");
            },
            (actual, expected) => {
                panic!("{label}[{index}] returned {actual:?}, expected {expected:?}");
            },
        }
    }
}

#[test]
fn every_elementary_function_broadcasts_over_a_literal_matrix() {
    let resolver = ElementaryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-elementary-array-family");
    let cases: [(&str, [Result<f64, ScalarError>; 4]); 11] = [
        ("=ABS({-1;0|2;-3})", [Ok(1.0), Ok(0.0), Ok(2.0), Ok(3.0)]),
        (
            "=EXP({0;1|2;3})",
            [
                Ok(1.0),
                Ok(1.0_f64.exp()),
                Ok(2.0_f64.exp()),
                Ok(3.0_f64.exp()),
            ],
        ),
        (
            "=LN({1;EXP(1)|10;EXP(2)})",
            [Ok(0.0), Ok(1.0), Ok(10.0_f64.ln()), Ok(2.0)],
        ),
        (
            "=LOG({1;10|100;1000})",
            [Ok(0.0), Ok(1.0), Ok(2.0), Ok(3.0)],
        ),
        (
            "=LOG10({1;10|100;1000})",
            [Ok(0.0), Ok(1.0), Ok(2.0), Ok(3.0)],
        ),
        (
            "=POWER({2;3|4;5};2)",
            [Ok(4.0), Ok(9.0), Ok(16.0), Ok(25.0)],
        ),
        ("=SQRT({0;1|4;9})", [Ok(0.0), Ok(1.0), Ok(2.0), Ok(3.0)]),
        (
            "=SQRTPI({0;1|4;9})",
            [
                Ok(0.0),
                Ok(PI.sqrt()),
                Ok((4.0 * PI).sqrt()),
                Ok((9.0 * PI).sqrt()),
            ],
        ),
        ("=SIGN({-2;0|3;-4})", [Ok(-1.0), Ok(0.0), Ok(1.0), Ok(-1.0)]),
        ("=MOD({5;-5|7;-7};3)", [Ok(2.0), Ok(1.0), Ok(1.0), Ok(2.0)]),
        (
            "=QUOTIENT({5;-5|7;-7};3)",
            [Ok(1.0), Ok(-1.0), Ok(2.0), Ok(-2.0)],
        ),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_array(&result, 2, 2, &expected, source);
    }
    assert_eq!(
        resolver.reads(),
        0,
        "literal arrays need no worksheet reads"
    );
}

#[test]
fn binary_arguments_broadcast_by_shape_and_optional_log_base_is_elementwise() {
    let resolver = ElementaryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-elementary-broadcast");

    let expression = parse("=POWER({1;2|3;4};{2;3})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("a 1-by-2 exponent should broadcast over a 2-by-2 base");
    assert_array(
        &result,
        2,
        2,
        &[Ok(1.0), Ok(8.0), Ok(9.0), Ok(64.0)],
        "POWER broadcast",
    );

    let expression = parse("=MOD({5;6|7;8};{2;3|4;5})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("same-shaped MOD arguments should remain rectangular");
    assert_array(
        &result,
        2,
        2,
        &[Ok(1.0), Ok(0.0), Ok(3.0), Ok(3.0)],
        "MOD broadcast",
    );

    let expression = parse("=LOG({10;100|1000;10000};{10;10|10;10})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("LOG base should broadcast elementwise");
    assert_array(
        &result,
        2,
        2,
        &[Ok(1.0), Ok(2.0), Ok(3.0), Ok(4.0)],
        "LOG base broadcast",
    );
}

#[test]
fn array_coercion_and_domain_errors_stay_local_to_each_cell() {
    let resolver = ElementaryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-elementary-array-errors");

    let expression = parse("=ABS({\"-2\";TRUE()|FALSE();\"bad\"})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("number coercion should be element-wise");
    assert_array(
        &result,
        2,
        2,
        &[Ok(2.0), Ok(1.0), Ok(0.0), Err(ScalarError::Value)],
        "ABS coercion",
    );

    let expression = parse("=LOG({10;0|100;-1};10)");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("LOG domain errors should be formula cells");
    assert_array(
        &result,
        2,
        2,
        &[
            Ok(1.0),
            Err(ScalarError::Number),
            Ok(2.0),
            Err(ScalarError::Number),
        ],
        "LOG domain cells",
    );

    let expression = parse("=MOD({1;2|3;4};{0;1|2;3})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("MOD division errors should be formula cells");
    assert_array(
        &result,
        2,
        2,
        &[Err(ScalarError::DivisionByZero), Ok(0.0), Ok(1.0), Ok(1.0)],
        "MOD division cells",
    );
}

#[test]
fn matrix_references_broadcast_and_scalar_mode_uses_implicit_intersection() {
    let resolver = ElementaryResolver::new([
        FixtureCell::Number(-2.0),
        FixtureCell::Number(0.0),
        FixtureCell::Number(3.0),
        FixtureCell::Number(4.0),
    ]);
    let (_budget, _cancellation, execution) = execution("ods-formula-elementary-references");

    let expression = parse("=ABS([.A1:.B2])");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("a matrix reference should be admitted");
    assert_array(
        &result,
        2,
        2,
        &[Ok(2.0), Ok(0.0), Ok(3.0), Ok(4.0)],
        "ABS reference",
    );
    assert_eq!(resolver.reads(), 4, "one reference should read four cells");

    let expression = parse("=MOD([.A1:.B2];2)");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("a scalar divisor should broadcast over a reference");
    assert_array(
        &result,
        2,
        2,
        &[Ok(0.0), Ok(0.0), Ok(1.0), Ok(0.0)],
        "MOD reference",
    );

    // Implicit intersection is specified for a one-row reference at the
    // caller's column. Use a separate fixture so the selected B1 value is
    // positive while the matrix checks above retain their four-cell inputs.
    let scalar_resolver = ElementaryResolver::new([
        FixtureCell::Number(-2.0),
        FixtureCell::Number(4.0),
        FixtureCell::Number(3.0),
        FixtureCell::Number(4.0),
    ]);
    let expression = parse("=SQRT([.A1:.B1])");
    let result = evaluate_at(
        &expression,
        &scalar_resolver,
        &execution,
        Position::new("Main", 0, 1),
        Mode::Scalar,
        &Limits::default(),
    )
    .expect("scalar mode should project the caller's B1 cell");
    match result.value() {
        Value::Number(value) => assert_close(value, 2.0, "SQRT scalar B1"),
        value => panic!("expected B1's projected number, got {value:?}"),
    }
    assert_eq!(
        scalar_resolver.reads(),
        1,
        "implicit intersection should read only the selected cell"
    );
}

#[test]
fn lazy_elementary_branches_skip_missing_references_and_catch_formula_errors() {
    let resolver = ElementaryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-elementary-lazy");
    for (source, expected) in [
        ("=IF(TRUE();ABS(-2);SQRT([Missing.A1]))", 2.0),
        ("=IF(FALSE();SQRT([Missing.A1]);ABS(-2))", 2.0),
        ("=IFERROR(SQRT(-1);42)", 42.0),
        ("=IFERROR(LOG(10;1);17)", 17.0),
    ] {
        let expression = parse(source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 1, 1),
            Mode::Scalar,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate lazily: {error}"));
        match result.value() {
            Value::Number(value) => assert_close(value, expected, source),
            value => panic!("{source:?} returned {value:?}, expected Number({expected})"),
        }
    }
    assert_eq!(
        resolver.reads(),
        0,
        "unselected references must not be resolved"
    );
}

#[test]
fn formula_error_cells_are_values_and_owned_results_keep_their_shape() {
    let resolver = ElementaryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-elementary-owned");
    let expression = parse("=SQRT({0;#N/A|4;#DIV/0!})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("matrix formula errors should remain cells");
    assert_array(
        &result,
        2,
        2,
        &[
            Ok(0.0),
            Err(ScalarError::NotAvailable),
            Ok(2.0),
            Err(ScalarError::DivisionByZero),
        ],
        "SQRT formula errors",
    );

    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("numeric/error array should be ownable");
    match owned.value() {
        OwnedValueView::Array(array) => {
            assert_eq!((array.shape().rows(), array.shape().columns()), (2, 2));
            assert!(matches!(array.get(0), Some(OwnedValueView::Number(0.0))));
            assert!(matches!(
                array.get(1),
                Some(OwnedValueView::Error(ScalarError::NotAvailable))
            ));
            assert!(matches!(array.get(2), Some(OwnedValueView::Number(2.0))));
            assert!(matches!(
                array.get(3),
                Some(OwnedValueView::Error(ScalarError::DivisionByZero))
            ));
        },
        value => panic!("expected owned array, got {value:?}"),
    }
}

#[test]
fn matrix_shape_work_reference_storage_and_cancellation_limits_are_atomic() {
    let expression = parse("=SQRT({0;1|4;9})");
    let resolver = ElementaryResolver::new([FixtureCell::Number(0.0); 4]);
    let (budget, _cancellation, exec) = execution("ods-formula-elementary-shape-limit");
    let error = evaluate_matrix(
        &expression,
        &resolver,
        &exec,
        &Limits::default().with_max_array_cells(3),
    )
    .expect_err("a four-cell result must exceed a three-cell shape cap");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong shape refusal: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, exec) = execution("ods-formula-elementary-work-limit");
    let error = evaluate_matrix(
        &expression,
        &resolver,
        &exec,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero work must refuse before publishing an array");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong work refusal: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let reference_expression = parse("=ABS([.A1:.B2])");
    let resolver = ElementaryResolver::new([FixtureCell::Number(0.0); 4]);
    let (budget, _cancellation, exec) = execution("ods-formula-elementary-reference-limit");
    let error = evaluate_matrix(
        &reference_expression,
        &resolver,
        &exec,
        &Limits::default().with_max_reference_cells(3),
    )
    .expect_err("a four-cell reference must exceed the admission cap");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference refusal: {error:?}"
    );
    assert_eq!(resolver.reads(), 0, "geometry must be refused before reads");
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation, exec) = execution("ods-formula-elementary-cancel");
    cancellation.cancel();
    let resolver = ElementaryResolver::new([FixtureCell::Number(0.0); 4]);
    let error = evaluate_matrix(&expression, &resolver, &exec, &Limits::default())
        .expect_err("pre-cancelled matrix evaluation must refuse");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, exec) = execution("ods-formula-elementary-storage-limit");
    let resolver = ElementaryResolver::new([FixtureCell::Number(0.0); 4]);
    let error = evaluate_matrix(
        &expression,
        &resolver,
        &exec,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero matrix storage must refuse atomically");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong storage refusal: {error:?}"
    );
    assert_eq!(
        resolver.reads(),
        0,
        "literal storage refusal reads no cells"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}
