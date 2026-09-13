//! Integration coverage for the OpenFormula 1.4 matrix-function family.
//!
//! The expected values in this file follow the profile recorded in
//! `docs/report/spec-gap-validation-evidence/ods-formula-matrix-functions/`:
//! numeric matrix elements are finite `f64` values, matrix errors are formula
//! values, and `MUNIT` truncates its scalar input toward zero.  The fixtures
//! use only integer matrix elements so no locale-dependent element coercion is
//! involved.

use std::{
    cell::Cell,
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{EvaluationFailure, EvaluationResult, ScalarError, UnsupportedKind},
    expression::Expression,
};

/// A tiny immutable provider used only for the ForceArray and lazy-branch
/// tests.  The matrix-function value cases do not read a worksheet at all.
#[derive(Debug)]
struct MatrixResolver {
    cells: [[f64; 2]; 2],
    reads: Cell<usize>,
    missing_metadata_calls: Cell<usize>,
    reject_missing: bool,
}

impl MatrixResolver {
    fn empty() -> Self {
        Self {
            cells: [[0.0; 2]; 2],
            reads: Cell::new(0),
            missing_metadata_calls: Cell::new(0),
            reject_missing: false,
        }
    }

    fn numeric() -> Self {
        Self {
            cells: [[1.0, 2.0], [3.0, 4.0]],
            ..Self::empty()
        }
    }

    fn rejecting_missing() -> Self {
        Self {
            reject_missing: true,
            ..Self::empty()
        }
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn missing_metadata_calls(&self) -> usize {
        self.missing_metadata_calls.get()
    }

    fn missing_sheet(&self, sheet: &str) -> Result<Option<SheetExtent>, EvaluationFailure> {
        if sheet != "Missing" {
            return Ok(None);
        }
        self.missing_metadata_calls
            .set(self.missing_metadata_calls.get().saturating_add(1));
        if self.reject_missing {
            return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
        }
        Ok(None)
    }
}

impl Resolver for MatrixResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<SheetExtent>> {
        if sheet == "Missing" {
            return self.missing_sheet(sheet);
        }
        Ok((sheet == "Main").then_some(SheetExtent::new(2, 2)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<CellRead<'a>> {
        self.reads.set(self.reads.get().saturating_add(1));
        if sheet != "Main" || row >= 2 || column >= 2 {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        Ok(CellRead::Number(self.cells[row][column]))
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<usize>> {
        if sheet == "Missing" {
            if self.reject_missing {
                self.missing_metadata_calls
                    .set(self.missing_metadata_calls.get().saturating_add(1));
                return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
            }
            return Ok(None);
        }
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
        NonZeroU64::new(1_024).expect("one in-flight KiB"),
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
    resolver: &'a MatrixResolver,
    execution: &ExecutionContext,
    mode: Mode,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    value::evaluate(expression, resolver, &context, &Limits::default())
}

fn assert_number(result: &Evaluated<'_>, expected: f64) {
    match result.value() {
        Value::Number(actual) => assert_eq!(actual, expected),
        other => panic!("expected Number({expected}), got {other:?}"),
    }
}

fn assert_array(result: &Evaluated<'_>, rows: usize, columns: usize, expected: &[f64]) {
    let array = result
        .as_array()
        .expect("matrix function should return an array");
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (rows, columns)
    );
    assert_eq!(array.len(), expected.len());
    for (index, expected) in expected.iter().copied().enumerate() {
        match array.get(index).expect("array cell should be present") {
            Value::Number(actual) => assert_eq!(actual, expected, "array cell {index}"),
            other => panic!("expected Number({expected}) at cell {index}, got {other:?}"),
        }
    }
}

fn assert_array_close(result: &Evaluated<'_>, rows: usize, columns: usize, expected: &[f64]) {
    let array = result
        .as_array()
        .expect("matrix function should return an array");
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (rows, columns)
    );
    assert_eq!(array.len(), expected.len());
    for (index, expected) in expected.iter().copied().enumerate() {
        match array.get(index).expect("array cell should be present") {
            Value::Number(actual) => {
                assert!(
                    (actual - expected).abs() <= 1.0e-12,
                    "array cell {index}: {actual} != {expected}"
                );
            },
            other => panic!("expected Number({expected}) at cell {index}, got {other:?}"),
        }
    }
}

fn assert_formula_error(result: &Evaluated<'_>, expected: ScalarError) {
    match result.value() {
        Value::Error(actual) => assert_eq!(actual, expected),
        other => panic!("expected formula error {expected}, got {other:?}"),
    }
}

fn assert_any_formula_error(result: &Evaluated<'_>) {
    assert!(
        matches!(result.value(), Value::Error(_)),
        "expected a formula error value, got {:?}",
        result.value()
    );
}

#[test]
fn mdeterm_returns_the_determinant_of_a_square_numeric_array() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-mdeterm");
    let expression = parse("=MDETERM({1;2|3;4})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("MDETERM should evaluate a square numeric array");
    assert_number(&result, -2.0);
}

#[test]
fn minverse_returns_a_square_inverse_with_finite_numeric_results() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-minverse");
    let expression = parse("=MINVERSE({1;2|3;4})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("MINVERSE should evaluate a nonsingular square array");
    assert_array_close(&result, 2, 2, &[-2.0, 1.0, 1.5, -0.5]);
}

#[test]
fn mmult_returns_the_rectangular_product_in_row_major_order() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-mmult");
    let expression = parse("=MMULT({1;2;3|4;5;6};{7;8|9;10|11;12})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("MMULT should evaluate compatible rectangular arrays");
    assert_array(&result, 2, 2, &[58.0, 64.0, 139.0, 154.0]);
}

#[test]
fn munit_returns_an_identity_array_and_uses_the_first_scalar_array_element() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-munit");

    let expression = parse("=MUNIT(3)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("MUNIT should construct a positive identity matrix");
    assert_array(
        &result,
        3,
        3,
        &[1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0],
    );

    // MUNIT's parameter is scalar, so an array-valued argument supplies its
    // [0,0] element instead of producing one identity matrix per element.
    let expression = parse("=MUNIT({2;3})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("MUNIT should project its scalar parameter from an array");
    assert_array(&result, 2, 2, &[1.0, 0.0, 0.0, 1.0]);
}

#[test]
fn transpose_exchanges_rectangular_axes_and_can_feed_matrix_broadcasting() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-transpose");

    let expression = parse("=TRANSPOSE({1;2;3|4;5;6})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("TRANSPOSE should exchange rectangular dimensions");
    assert_array(&result, 3, 2, &[1.0, 4.0, 2.0, 5.0, 3.0, 6.0]);

    let expression = parse("=TRANSPOSE({1;2|3;4})+1");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("a transposed array should participate in matrix broadcasting");
    assert_array(&result, 2, 2, &[2.0, 4.0, 3.0, 5.0]);
}

#[test]
fn matrix_function_names_are_case_insensitive() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-case");

    let expression = parse("=mdeterm({1;2|3;4})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_number(&result, -2.0);

    let expression = parse("=minverse({1;2|3;4})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_array_close(&result, 2, 2, &[-2.0, 1.0, 1.5, -0.5]);

    let expression = parse("=mmult({1;2|3;4};{5;6|7;8})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_array(&result, 2, 2, &[19.0, 22.0, 43.0, 50.0]);

    let expression = parse("=munit(2)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_array(&result, 2, 2, &[1.0, 0.0, 0.0, 1.0]);

    let expression = parse("=transpose({1;2|3;4})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_array(&result, 2, 2, &[1.0, 3.0, 2.0, 4.0]);
}

#[test]
fn matrix_functions_return_value_errors_for_invalid_arity() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-arity");
    for source in [
        "=MDETERM({1};{2})",
        "=MINVERSE()",
        "=MMULT({1})",
        "=MUNIT(2;3)",
        "=TRANSPOSE()",
    ] {
        let expression = parse(source);
        let result =
            evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap_or_else(|error| {
                panic!("{source} should return a formula arity error: {error}")
            });
        assert_formula_error(&result, ScalarError::Value);
    }
}

#[test]
fn matrix_functions_reject_invalid_shapes_singular_inverse_and_nonpositive_units() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-errors");

    for source in [
        "=MDETERM({1;2;3|4;5;6})",
        "=MINVERSE({1;2;3|4;5;6})",
        "=MMULT({1;2;3|4;5;6};{1;2|3;4})",
    ] {
        let expression = parse(source);
        let result =
            evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap_or_else(|error| {
                panic!("{source} should return a formula shape error: {error}")
            });
        assert_formula_error(&result, ScalarError::Value);
    }

    let expression = parse("=MINVERSE({1;2|2;4})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("a singular inverse should remain a formula error");
    assert_formula_error(&result, ScalarError::Number);

    for source in ["=MUNIT(0)", "=MUNIT(-1)"] {
        let expression = parse(source);
        let result =
            evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap_or_else(|error| {
                panic!("{source} should return a formula domain error: {error}")
            });
        assert_any_formula_error(&result);
    }
}

#[test]
fn matrix_numeric_operations_propagate_element_errors_as_formula_values() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-error-propagation");
    let expression = parse("=MDETERM({1;#N/A|3;4})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("an input formula error should be propagated as a value");
    assert_formula_error(&result, ScalarError::NotAvailable);
}

#[test]
fn force_array_arguments_are_fully_materialized_in_scalar_context() {
    let resolver = MatrixResolver::numeric();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-force-array");

    // MDETERM's ForceArray parameter must retain the whole expression result
    // even when the caller requests a scalar result.
    let expression = parse("=MDETERM({1;2|3;4}+1)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Scalar)
        .expect("ForceArray should evaluate a scalar-context array expression");
    assert_number(&result, -2.0);

    let expression = parse("=MDETERM([.A1:.B2])");
    let result = evaluate(&expression, &resolver, &execution, Mode::Scalar)
        .expect("ForceArray should materialize a multi-cell reference");
    assert_number(&result, -2.0);
    assert_eq!(resolver.reads(), 4, "the full 2x2 reference must be read");
}

#[test]
fn scalar_if_evaluates_a_selected_matrix_function_without_touching_unselected_reference() {
    let resolver = MatrixResolver::rejecting_missing();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-lazy-true");
    let expression = parse("=IF(TRUE();MDETERM({1;2|3;4});[Missing.A1:.Z100])");
    let result = evaluate(&expression, &resolver, &execution, Mode::Scalar)
        .expect("the selected matrix function should evaluate");
    assert_number(&result, -2.0);
    assert_eq!(resolver.reads(), 0);
    assert_eq!(resolver.missing_metadata_calls(), 0);
}

#[test]
fn scalar_if_selects_the_matrix_function_in_the_else_branch_lazily() {
    let resolver = MatrixResolver::rejecting_missing();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-lazy-false");
    let expression = parse("=IF(FALSE();[Missing.A1:.Z100];MDETERM({1;2|3;4}))");
    let result = evaluate(&expression, &resolver, &execution, Mode::Scalar)
        .expect("the selected else matrix function should evaluate");
    assert_number(&result, -2.0);
    assert_eq!(resolver.reads(), 0);
    assert_eq!(resolver.missing_metadata_calls(), 0);
}

#[test]
fn matrix_if_broadcasts_a_scalar_branch_against_a_transposed_array() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-if-transpose");
    let expression = parse("=IF({TRUE();FALSE()};TRANSPOSE({1;2;3|4;5;6});0)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("matrix IF should retain the selected array shape");
    assert_array(&result, 3, 2, &[1.0, 0.0, 2.0, 0.0, 3.0, 0.0]);
}

#[test]
fn nested_scalar_if_and_error_handler_keep_the_selected_array_shape() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-nested-handler");
    // The singular inverse is a formula Number error.  IFERROR selects the
    // transposed array, and the outer scalar IF must retain that 2x2 result
    // while discarding its scalar else branch.
    let expression = parse("=IF(TRUE();IFERROR(MINVERSE({1;2|2;4});TRANSPOSE({1;2|3;4}));0)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("nested scalar IF/error handler should evaluate its selected array");
    assert_array(&result, 2, 2, &[1.0, 3.0, 2.0, 4.0]);
}

#[test]
fn force_array_matrix_functions_consume_lazy_if_arrays_and_nested_transpose() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-force-lazy-arrays");

    let expression = parse("=MDETERM(TRANSPOSE(IF(TRUE();{1;3|2;4};{9;9|9;9})))");
    let result = evaluate(&expression, &resolver, &execution, Mode::Scalar)
        .expect("MDETERM should consume the selected ForceArray expression");
    assert_number(&result, -2.0);

    let expression = parse("=MINVERSE(TRANSPOSE(IF(FALSE();{1;0|0;1};{2;0|0;2})))");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("MINVERSE should consume the selected nested array");
    assert_array(&result, 2, 2, &[0.5, 0.0, 0.0, 0.5]);

    let expression =
        parse("=MMULT(TRANSPOSE(IF(TRUE();{1;3|2;4};{0;0|0;0}));IF(FALSE();{5;6|7;8};{2;0|0;2}))");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("MMULT should consume both selected ForceArray expressions");
    assert_array(&result, 2, 2, &[2.0, 4.0, 6.0, 8.0]);
}

#[test]
fn numeric_matrix_functions_reject_text_and_logical_elements_while_transpose_preserves_types() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-element-profile");

    for source in [
        "=MDETERM({\"1\";2|3;4})",
        "=MINVERSE({TRUE();2|3;4})",
        "=MMULT({1;\"2\"};{3|4})",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source} should return a formula type error: {error}"));
        assert_formula_error(&result, ScalarError::Value);
    }

    let expression = parse("=TRANSPOSE({\"x\";TRUE()|2;#N/A})");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("TRANSPOSE should preserve nonnumeric element types");
    let array = result.as_array().expect("TRANSPOSE should return an array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (2, 2));
    assert!(matches!(array.get(0), Some(Value::Text("x"))));
    assert!(matches!(array.get(1), Some(Value::Number(2.0))));
    assert!(matches!(array.get(2), Some(Value::Logical(true))));
    assert!(matches!(
        array.get(3),
        Some(Value::Error(ScalarError::NotAvailable))
    ));
}

#[test]
fn matrix_if_restores_outer_state_after_force_array_inner_if() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-formula-matrix-state-restoration");
    // The inner IF is evaluated as MDETERM's ForceArray argument while the
    // outer matrix IF has a pending 1x2 continuation.  Its selected matrix is
    // [[2, 0], [0, 3]], whose determinant is 6; the false outer position uses
    // the scalar else branch and must remain zero.
    let expression = parse(
        "=IF({TRUE();FALSE()};MDETERM(IF({TRUE();FALSE()|FALSE();TRUE()};{2;9|9;3};{9;0|0;9}));0)",
    );
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix)
        .expect("nested ForceArray evaluation should restore the outer matrix state");
    assert_array(&result, 1, 2, &[6.0, 0.0]);
}

#[test]
fn munit_converts_logical_parameters_and_truncates_fractional_sizes() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-munit-integer-profile");
    for source in ["=MUNIT(2.9)", "=MUNIT({2.9;3.9})"] {
        let expression = parse(source);
        let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
        assert_array(&result, 2, 2, &[1.0, 0.0, 0.0, 1.0]);
    }
    let expression = parse("=MUNIT(TRUE())");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_array(&result, 1, 1, &[1.0]);
}

#[test]
fn explicit_matrix_errors_precede_generated_element_type_errors() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-matrix-error-order");
    for source in [
        "=MDETERM({\"x\";#N/A|1;2})",
        "=MINVERSE({\"x\";#N/A|1;2})",
        "=MMULT({\"x\";1};{#N/A|2})",
        "=MMULT({#N/A};#DIV/0!)",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
        assert_formula_error(&result, ScalarError::NotAvailable);
    }
    let expression = parse("=TRANSPOSE(#N/A)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_formula_error(&result, ScalarError::NotAvailable);
}

#[test]
fn matrix_if_plans_computed_identity_dimensions() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-munit-computed-shape");
    let expression = parse("=IF({TRUE();FALSE()};MUNIT(1+1);0)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_array(&result, 2, 2, &[1.0, 0.0, 0.0, 0.0]);
}

#[test]
fn reference_operators_distinguish_successful_inverse_from_error_fallback() {
    let resolver = MatrixResolver::numeric();
    let (_budget, _cancellation, execution) = execution("ods-matrix-reference-fallback");
    let expression = parse("=IFERROR(MINVERSE({1;2|2;4});[.A1:.B1]):[.A1]");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert!(result.as_reference().is_some());
    assert_eq!(resolver.reads(), 0);

    let expression = parse("=IFERROR(MINVERSE({1;0|0;1});[.A1:.B1]):[.A1]");
    let error = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap_err();
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator)
    ));
}

#[test]
fn computed_identity_size_does_not_replace_its_callers_shape_continuations() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-matrix-nested-shape-probe");
    // Inner IF gives the row [1,4]; its dot product with [1,0] is one.
    // MUNIT(1) plus [1,2] is the row [2,3], whose two columns must survive
    // the nested planner used to obtain MUNIT's scalar argument.
    let expression =
        parse("=IF({TRUE()};MUNIT(MMULT(IF({TRUE();FALSE()};{1;2};{3;4});{1|0}))+{1;2};0)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_array(&result, 1, 2, &[2.0, 3.0]);
}

#[test]
fn transpose_argument_expression_keeps_matrix_context_in_a_lazy_branch() {
    let resolver = MatrixResolver::empty();
    let (_budget, _cancellation, execution) = execution("ods-transpose-lazy-context");
    let expression = parse("=IF({TRUE();TRUE()};TRANSPOSE({1;2|3;4}+1);0)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_array(&result, 2, 2, &[2.0, 4.0, 3.0, 5.0]);
}

#[test]
fn munit_uses_the_first_parameter_cell_in_matrix_mode() {
    let resolver = MatrixResolver::numeric();
    let (_budget, _cancellation, execution) = execution("ods-munit-first-parameter");
    let expression = parse("=MUNIT([.A1:.B2])");
    let context = Context::new(&execution, Position::new("Main", 1, 1));
    let result = value::evaluate(&expression, &resolver, &context, &Limits::default()).unwrap();
    assert_array(&result, 1, 1, &[1.0]);
    assert_eq!(resolver.reads(), 1);

    let expression = parse("=IF({TRUE()};MUNIT([.A1:.B2]);0)");
    let result = value::evaluate(&expression, &resolver, &context, &Limits::default()).unwrap();
    assert_array(&result, 1, 1, &[1.0]);

    let expression = parse("=IF({FALSE();TRUE()};MUNIT({1;2});0)");
    let result = evaluate(&expression, &resolver, &execution, Mode::Matrix).unwrap();
    assert_array(&result, 1, 2, &[0.0, 1.0]);
}
