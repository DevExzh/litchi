//! Independent matrix/reference integration coverage for the OpenFormula 1.4
//! trigonometric family.
//!
//! The expected values are derived from the Part 4 §6.16 definitions.  This
//! file checks element-wise function evaluation, scalar broadcasting, input
//! coercion, references, formula errors, ownership, and the value evaluator's
//! separate work/shape/cancellation limits.

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
    Error(ScalarError),
}

#[derive(Debug)]
struct TrigonometryResolver {
    cells: [FixtureCell; 4],
    reads: AtomicUsize,
}

impl TrigonometryResolver {
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

impl Resolver for TrigonometryResolver {
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
        let cell = self.cells[row * 2 + column];
        Ok(match cell {
            FixtureCell::Number(value) => CellRead::Number(value),
            FixtureCell::Error(error) => CellRead::Error(error),
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
    resolver: &'a TrigonometryResolver,
    execution: &ExecutionContext,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    evaluate(expression, resolver, &context, limits)
}

fn evaluate_at<'a>(
    expression: &'a Expression,
    resolver: &'a TrigonometryResolver,
    execution: &ExecutionContext,
    position: Position<'a>,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    evaluate(expression, resolver, &context, limits)
}

fn acot(value: f64) -> f64 {
    if value == 0.0 {
        PI / 2.0
    } else if value > 0.0 {
        (1.0 / value).atan()
    } else {
        (1.0 / value).atan() + PI
    }
}

fn acoth(value: f64) -> f64 {
    0.5 * ((value + 1.0) / (value - 1.0)).ln()
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
    assert_eq!(array.shape().rows(), rows, "{label} row shape");
    assert_eq!(array.shape().columns(), columns, "{label} column shape");
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
                assert_eq!(actual, *expected, "{label}[{index}]")
            },
            (actual, expected) => {
                panic!("{label}[{index}] returned {actual:?}, expected {expected:?}")
            },
        }
    }
}

#[test]
fn every_trigonometric_function_broadcasts_over_a_literal_matrix() {
    let resolver = TrigonometryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-trigonometry-array-family");
    let cases: [(&str, [Result<f64, ScalarError>; 4]); 23] = [
        (
            "=ACOS({1;0|-1;0.5})",
            [Ok(0.0), Ok(PI / 2.0), Ok(PI), Ok(PI / 3.0)],
        ),
        (
            "=ACOSH({1;2|10;4})",
            [
                Ok(1.0_f64.acosh()),
                Ok(2.0_f64.acosh()),
                Ok(10.0_f64.acosh()),
                Ok(4.0_f64.acosh()),
            ],
        ),
        (
            "=ACOT({-1;0|1;2})",
            [Ok(acot(-1.0)), Ok(acot(0.0)), Ok(acot(1.0)), Ok(acot(2.0))],
        ),
        (
            "=ACOTH({-2;-4|2;10})",
            [
                Ok(acoth(-2.0)),
                Ok(acoth(-4.0)),
                Ok(acoth(2.0)),
                Ok(acoth(10.0)),
            ],
        ),
        (
            "=ASIN({-1;0|1;0.5})",
            [Ok(-PI / 2.0), Ok(0.0), Ok(PI / 2.0), Ok(PI / 6.0)],
        ),
        (
            "=ASINH({-1;0|1;2})",
            [
                Ok((-1.0_f64).asinh()),
                Ok(0.0),
                Ok(1.0_f64.asinh()),
                Ok(2.0_f64.asinh()),
            ],
        ),
        (
            "=ATAN({-1;0|1;2})",
            [Ok(-PI / 4.0), Ok(0.0), Ok(PI / 4.0), Ok(2.0_f64.atan())],
        ),
        (
            "=ATAN2({1;-1|1;-1};{1;1|-1;-1})",
            [
                Ok(PI / 4.0),
                Ok(3.0 * PI / 4.0),
                Ok(-PI / 4.0),
                Ok(-3.0 * PI / 4.0),
            ],
        ),
        (
            "=ATANH({-0.5;0|0.5;0.25})",
            [
                Ok((-0.5_f64).atanh()),
                Ok(0.0),
                Ok(0.5_f64.atanh()),
                Ok(0.25_f64.atanh()),
            ],
        ),
        (
            "=COS({0;PI()/2|PI();-PI()/2})",
            [
                Ok(1.0),
                Ok((PI / 2.0).cos()),
                Ok(PI.cos()),
                Ok((-PI / 2.0).cos()),
            ],
        ),
        (
            "=COSH({0;1|-1;2})",
            [
                Ok(1.0),
                Ok(1.0_f64.cosh()),
                Ok(1.0_f64.cosh()),
                Ok(2.0_f64.cosh()),
            ],
        ),
        (
            "=COT({PI()/4;-PI()/4|PI()/6;-PI()/6})",
            [Ok(1.0), Ok(-1.0), Ok(3.0_f64.sqrt()), Ok(-3.0_f64.sqrt())],
        ),
        (
            "=COTH({1;-1|2;-2})",
            [
                Ok(1.0_f64.tanh().recip()),
                Ok(-1.0_f64.tanh().recip()),
                Ok(2.0_f64.tanh().recip()),
                Ok(-2.0_f64.tanh().recip()),
            ],
        ),
        (
            "=CSC({PI()/2;-PI()/2|PI()/6;-PI()/6})",
            [Ok(1.0), Ok(-1.0), Ok(2.0), Ok(-2.0)],
        ),
        (
            "=CSCH({1;-1|2;-2})",
            [
                Ok(1.0_f64.sinh().recip()),
                Ok(-1.0_f64.sinh().recip()),
                Ok(2.0_f64.sinh().recip()),
                Ok(-2.0_f64.sinh().recip()),
            ],
        ),
        (
            "=SEC({0;PI()/3|PI();-PI()/3})",
            [Ok(1.0), Ok(2.0), Ok(-1.0), Ok(2.0)],
        ),
        (
            "=SECH({0;1|-1;2})",
            [
                Ok(1.0),
                Ok(1.0_f64.cosh().recip()),
                Ok(1.0_f64.cosh().recip()),
                Ok(2.0_f64.cosh().recip()),
            ],
        ),
        (
            "=SIN({0;PI()/2|PI();-PI()/2})",
            [
                Ok(0.0),
                Ok((PI / 2.0).sin()),
                Ok(PI.sin()),
                Ok((-PI / 2.0).sin()),
            ],
        ),
        (
            "=SINH({0;1|-1;2})",
            [
                Ok(0.0),
                Ok(1.0_f64.sinh()),
                Ok(-1.0_f64.sinh()),
                Ok(2.0_f64.sinh()),
            ],
        ),
        (
            "=TAN({0;PI()/4|PI()/6;-PI()/6})",
            [
                Ok(0.0),
                Ok(1.0),
                Ok(1.0 / 3.0_f64.sqrt()),
                Ok(-1.0 / 3.0_f64.sqrt()),
            ],
        ),
        (
            "=TANH({0;1|-1;2})",
            [
                Ok(0.0),
                Ok(1.0_f64.tanh()),
                Ok(-1.0_f64.tanh()),
                Ok(2.0_f64.tanh()),
            ],
        ),
        (
            "=DEGREES({0;PI()|PI()/2;-PI()/2})",
            [Ok(0.0), Ok(180.0), Ok(90.0), Ok(-90.0)],
        ),
        (
            "=RADIANS({0;180|90;-90})",
            [Ok(0.0), Ok(PI), Ok(PI / 2.0), Ok(-PI / 2.0)],
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
        "literal arrays must not read a worksheet"
    );
}

#[test]
fn scalar_arguments_broadcast_and_atan2_keeps_x_then_y_order() {
    let resolver = TrigonometryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-trigonometry-broadcast");
    let expression = parse("=ATAN2({1;-1|0;2};1)");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("scalar y should broadcast over the x matrix");
    assert_array(
        &result,
        2,
        2,
        &[
            Ok(PI / 4.0),
            Ok(3.0 * PI / 4.0),
            Ok(PI / 2.0),
            Ok(1.0_f64.atan2(2.0)),
        ],
        "ATAN2 broadcast",
    );

    let expression = parse("=SIN({0;PI()/2|PI();-PI()/2})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("a non-square literal array should retain its shape");
    assert_array(
        &result,
        2,
        2,
        &[
            Ok(0.0),
            Ok((PI / 2.0).sin()),
            Ok(PI.sin()),
            Ok((-PI / 2.0).sin()),
        ],
        "SIN shape",
    );
}

#[test]
fn scalar_mode_projects_a_non_origin_reference_and_keeps_trig_branches_lazy() {
    let resolver = TrigonometryResolver::new([
        FixtureCell::Number(0.0),
        FixtureCell::Number(PI / 2.0),
        FixtureCell::Error(ScalarError::NotAvailable),
        FixtureCell::Number(-PI / 2.0),
    ]);
    let (_budget, _cancellation, exec) = execution("ods-formula-trigonometry-scalar-intersection");
    let expression = parse("=SIN([.A1:.B1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &exec,
        Position::new("Main", 0, 1),
        Mode::Scalar,
        &Limits::default(),
    )
    .expect("scalar mode should project the reference at the current cell");
    match result.value() {
        Value::Number(value) => assert_close(value, 1.0, "SIN scalar B1"),
        value => panic!("expected B1's projected numeric value, got {value:?}"),
    }
    assert_eq!(
        resolver.reads(),
        1,
        "implicit intersection should read only B1"
    );

    let resolver = TrigonometryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-trigonometry-lazy-branches");
    for (source, expected) in [
        ("=IF(TRUE();SIN(0);COS([Missing.A1]))", 0.0),
        ("=IF(FALSE();SIN([Missing.A1]);COS(0))", 1.0),
        ("=IFERROR(ACOS(2);SIN(0))", 0.0),
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
        "unselected references must never be resolved"
    );
}

#[test]
fn references_broadcast_through_trigonometry_and_preserve_element_errors() {
    let resolver = TrigonometryResolver::new([
        FixtureCell::Number(0.0),
        FixtureCell::Number(PI / 2.0),
        FixtureCell::Error(ScalarError::NotAvailable),
        FixtureCell::Number(-PI / 2.0),
    ]);
    let (_budget, _cancellation, execution) = execution("ods-formula-trigonometry-references");

    let expression = parse("=SIN([.A1:.B2])");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("a matrix reference should be admitted");
    assert_array(
        &result,
        2,
        2,
        &[Ok(0.0), Ok(1.0), Err(ScalarError::NotAvailable), Ok(-1.0)],
        "SIN reference",
    );

    let expression = parse("=ATAN2([.A1:.B2];1)");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("the scalar y coordinate should broadcast over a reference");
    assert_array(
        &result,
        2,
        2,
        &[
            Ok(PI / 2.0),
            Ok(1.0_f64.atan2(PI / 2.0)),
            Err(ScalarError::NotAvailable),
            Ok(1.0_f64.atan2(-PI / 2.0)),
        ],
        "ATAN2 reference",
    );
    assert_eq!(resolver.reads(), 8, "each reference matrix is read once");
}

#[test]
fn array_elements_use_number_conversion_and_retain_individual_errors() {
    let resolver = TrigonometryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-trigonometry-array-coercion");
    let expression = parse("=SIN({\"0\";TRUE()|FALSE();\"bad\"})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("element-wise number conversion should succeed");
    assert_array(
        &result,
        2,
        2,
        &[Ok(0.0), Ok(1.0_f64.sin()), Ok(0.0), Err(ScalarError::Value)],
        "SIN coercion",
    );

    let expression = parse("=ACOS({\"1\";TRUE()|FALSE();\"bad\"})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("ACOS should retain per-cell conversion/domain values");
    assert_array(
        &result,
        2,
        2,
        &[Ok(0.0), Ok(0.0), Ok(PI / 2.0), Err(ScalarError::Value)],
        "ACOS coercion",
    );
}

#[test]
fn matrix_formula_errors_are_values_and_owned_numeric_results_survive_source_drop() {
    let resolver = TrigonometryResolver::new([FixtureCell::Number(0.0); 4]);
    let (_budget, _cancellation, execution) = execution("ods-formula-trigonometry-array-owned");
    let expression = parse("=SIN({0;#N/A|PI()/2;#DIV/0!})");
    let result = evaluate_matrix(&expression, &resolver, &execution, &Limits::default())
        .expect("matrix formula errors should remain cells");
    assert_array(
        &result,
        2,
        2,
        &[
            Ok(0.0),
            Err(ScalarError::NotAvailable),
            Ok(1.0),
            Err(ScalarError::DivisionByZero),
        ],
        "SIN formula errors",
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
            assert!(
                matches!(array.get(2), Some(OwnedValueView::Number(value)) if (value - 1.0).abs() <= 32.0 * f64::EPSILON)
            );
            assert!(matches!(
                array.get(3),
                Some(OwnedValueView::Error(ScalarError::DivisionByZero))
            ));
        },
        value => panic!("expected owned numeric array, got {value:?}"),
    }
}

#[test]
fn array_shape_work_reference_and_cancellation_limits_are_exact_and_atomic() {
    let resolver = TrigonometryResolver::new([FixtureCell::Number(0.0); 4]);
    let expression = parse("=SIN({0;PI()/2|PI();-PI()/2})");
    let (budget, _cancellation, exec) = execution("ods-formula-trigonometry-array-shape-limit");
    let error = evaluate_matrix(
        &expression,
        &resolver,
        &exec,
        &Limits::default().with_max_array_cells(3),
    )
    .expect_err("a four-cell trigonometric result must exceed a three-cell cap");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong array-shape refusal: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, exec) = execution("ods-formula-trigonometry-array-work-limit");
    let error = evaluate_matrix(
        &expression,
        &resolver,
        &exec,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero work must refuse before publishing an array");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong array-work refusal: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let reference_expression = parse("=SIN([.A1:.B2])");
    let resolver = TrigonometryResolver::new([FixtureCell::Number(0.0); 4]);
    let (budget, _cancellation, exec) = execution("ods-formula-trigonometry-reference-limit");
    let error = evaluate_matrix(
        &reference_expression,
        &resolver,
        &exec,
        &Limits::default().with_max_reference_cells(3),
    )
    .expect_err("a four-cell reference must exceed a three-cell admission cap");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference refusal: {error:?}"
    );
    assert_eq!(resolver.reads(), 0, "geometry must be refused before reads");
    assert_eq!(budget.used(Resource::Memory), 0);

    let (_budget, cancellation, exec) = execution("ods-formula-trigonometry-array-cancel");
    cancellation.cancel();
    let error = evaluate_matrix(&expression, &resolver, &exec, &Limits::default())
        .expect_err("pre-cancelled array trigonometry must refuse");
    assert!(matches!(error, EvaluationFailure::Cancelled));
}

#[test]
fn matrix_storage_cap_refusal_is_atomic_and_refunds_the_caller_budget() {
    let resolver = TrigonometryResolver::new([FixtureCell::Number(0.0); 4]);
    let expression = parse("=SIN({0;1|2;3})");
    let (budget, _cancellation, execution) =
        execution("ods-formula-trigonometry-array-storage-limit");
    let error = evaluate_matrix(
        &expression,
        &resolver,
        &execution,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("a matrix result must refuse with zero evaluator storage");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong storage refusal: {error:?}"
    );
    assert_eq!(
        resolver.reads(),
        0,
        "literal storage refusal must read no cells"
    );
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "failed evaluation leaked storage"
    );
}
