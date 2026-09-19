//! Independent scalar coverage for the OpenFormula 1.4 numeric aggregate
//! family in §6.16: SUM, PRODUCT, SUMSQ, SUMPRODUCT, SUMX2MY2, SUMX2PY2,
//! and SUMXMY2.
//!
//! These tests intentionally keep the scalar profile separate from the value
//! evaluator's reference/array bridge.  The expected values are calculated in
//! this file from the definitions, so a test does not merely repeat the
//! implementation's helper arithmetic.

use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        UnsupportedKind, evaluate_scalar,
    },
    expression::Expression,
};

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
enum ScalarAggregateResult {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn evaluate(
    source: &str,
    execution: &ExecutionContext,
    limits: EvaluationLimits,
) -> Result<ScalarAggregateResult, EvaluationFailure> {
    let expression = parse(source);
    let result = evaluate_scalar(&expression, &EvaluationContext::new(execution), &limits)?;
    Ok(match result.value() {
        ScalarValue::Number(value) => ScalarAggregateResult::Number(*value),
        ScalarValue::Error(error) => ScalarAggregateResult::Error(*error),
        _ => ScalarAggregateResult::Other,
    })
}

fn number(result: &ScalarAggregateResult) -> f64 {
    match result {
        ScalarAggregateResult::Number(value) => *value,
        value => panic!("expected Number, got {value:?}"),
    }
}

fn formula_error(result: &ScalarAggregateResult) -> ScalarError {
    match result {
        ScalarAggregateResult::Error(error) => *error,
        value => panic!("expected formula Error, got {value:?}"),
    }
}

fn assert_close(source: &str, expected: f64) {
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-scalar");
    let result = evaluate(source, &execution, EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    let actual = number(&result);
    if expected == 0.0 {
        assert_eq!(
            actual.to_bits(),
            expected.to_bits(),
            "{source:?}: {actual:?}"
        );
    } else {
        assert!(
            actual.is_finite(),
            "{source:?} returned non-finite {actual:?}"
        );
        let tolerance = expected.abs().max(actual.abs()) * (32.0 * f64::EPSILON);
        assert!(
            (actual - expected).abs() <= tolerance,
            "{source:?}: {actual:?} != {expected:?} (tol {tolerance:e})"
        );
    }
}

fn assert_formula_error(source: &str, expected: ScalarError) {
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-error");
    let result = evaluate(source, &execution, EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should produce a formula error: {error}"));
    assert_eq!(formula_error(&result), expected, "{source:?}");
}

fn assert_unsupported(source: &str, expected: UnsupportedKind) {
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-unsupported");
    let error = evaluate(source, &execution, EvaluationLimits::default())
        .expect_err("reference-dependent expression should be refused by scalar evaluation");
    assert!(
        matches!(&error, EvaluationFailure::Unsupported(kind) if *kind == expected),
        "{source:?} returned the wrong capability error: {error:?}"
    );
}

#[test]
fn every_aggregate_function_has_a_scalar_definition_vector() {
    for (source, expected) in [
        ("=SUM(1;2;3;4)", 10.0),
        ("=PRODUCT(2;3;4)", 24.0),
        ("=SUMSQ(1;2;3)", 14.0),
        ("=SUMPRODUCT(1;2;3)", 6.0),
        ("=SUMX2MY2(3;4)", -7.0),
        ("=SUMX2PY2(3;4)", 25.0),
        ("=SUMXMY2(3;4)", 1.0),
    ] {
        assert_close(source, expected);
    }
}

#[test]
fn aggregate_identities_and_nested_results_follow_the_sequence_contract() {
    // This profile exposes SUM/SUMSQ's additive identity for an omitted
    // sequence while keeping PRODUCT's strict variadic arity; nested
    // reductions still compose as scalar values.
    for (source, expected) in [
        ("=SUM()", 0.0),
        ("=SUMSQ()", 0.0),
        ("=SUM(1;PRODUCT(2;3);SUMSQ(2;3))", 20.0),
        ("=PRODUCT(SUM(1;2);SUMSQ(2))", 12.0),
        ("=SUMSQ(SUM(1;2);PRODUCT(2;3))", 45.0),
        ("=SUMPRODUCT(SUM(1;2);PRODUCT(2;3))", 18.0),
        ("=SUMX2MY2(SUM(1;2);PRODUCT(2;3))", -27.0),
        ("=SUMX2PY2(SUM(1;2);PRODUCT(2;3))", 45.0),
        ("=SUMXMY2(SUM(1;2);PRODUCT(2;3))", 9.0),
    ] {
        assert_close(source, expected);
    }
    assert_formula_error("=PRODUCT()", ScalarError::Value);
}

#[test]
fn scalar_number_conversion_is_distinct_from_reference_sequence_filtering() {
    // Text and logical scalar arguments are converted to Numbers by the
    // selected locale-independent scalar profile.  Reference sequences use
    // the §6.3.7 rule: text and empty cells are omitted and distinguished
    // logical cells are omitted; the companion value tests cover that bridge.
    for (source, expected) in [
        ("=SUM(\"2.5\";TRUE();FALSE())", 3.5),
        ("=PRODUCT(\"2\";TRUE();FALSE())", 0.0),
        ("=SUMSQ(\"2\";TRUE())", 5.0),
        ("=SUMPRODUCT(\"2\";TRUE())", 2.0),
        ("=SUMX2MY2(\"2\";TRUE())", 3.0),
        ("=SUMX2PY2(\"2\";TRUE())", 5.0),
        ("=SUMXMY2(\"2\";TRUE())", 1.0),
    ] {
        assert_close(source, expected);
    }
    assert_formula_error("=SUM(\"not-a-number\")", ScalarError::Value);
    assert_formula_error("=PRODUCT(\"not-a-number\")", ScalarError::Value);
    assert_formula_error("=SUMSQ(\"not-a-number\")", ScalarError::Value);
    assert_formula_error("=SUMPRODUCT(\"not-a-number\";1)", ScalarError::Value);
}

#[test]
fn aggregate_arity_and_pair_shape_errors_are_formula_values() {
    for source in [
        "=SUMPRODUCT()",
        "=SUMX2MY2(1)",
        "=SUMX2MY2(1;2;3)",
        "=SUMX2PY2(1)",
        "=SUMXMY2(1;2;3)",
        "=SUM(;)",
    ] {
        assert_formula_error(source, ScalarError::Value);
    }
}

#[test]
fn sequence_aggregates_accept_more_than_the_minimum_number_of_arguments() {
    assert_close(
        "=SUM(1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30;31;32;33)",
        561.0,
    );
    assert_close(
        "=PRODUCT(1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1)",
        1.0,
    );
    assert_close(
        "=SUMSQ(1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30;31;32;33)",
        12529.0,
    );
}

#[test]
fn aggregate_formula_errors_are_propagated_in_argument_order() {
    for (source, expected) in [
        ("=SUM(#N/A;1)", ScalarError::NotAvailable),
        ("=SUM(1;1/0)", ScalarError::DivisionByZero),
        ("=PRODUCT(#N/A;1)", ScalarError::NotAvailable),
        ("=SUMSQ(1;#N/A)", ScalarError::NotAvailable),
        ("=SUMPRODUCT(1;#N/A)", ScalarError::NotAvailable),
        ("=SUMX2MY2(#N/A;1)", ScalarError::NotAvailable),
        ("=SUMX2PY2(1;#N/A)", ScalarError::NotAvailable),
        ("=SUMXMY2(1;#N/A)", ScalarError::NotAvailable),
    ] {
        assert_formula_error(source, expected);
    }
    assert_close("=SUM(1;IFERROR(1/0;7))", 8.0);
}

#[test]
fn scalar_aggregate_arithmetic_uses_an_independent_cancellation_oracle() {
    // Pairwise terms are chosen so the exact mathematical sums are small even
    // though individual products/squares are large.  The expected values are
    // written from the definitions, rather than obtained by calling the
    // evaluator's arithmetic helpers.
    for (source, expected) in [
        ("=SUM(1e16;1;-1e16)", 1.0),
        ("=SUMSQ(1e-200;2e-200)", 0.0),
        ("=SUMSQ(1.5e-162;1.5e-162)", f64::from_bits(1)),
        ("=SUMPRODUCT(1e154;1e-154)", 1.0),
        ("=SUMX2MY2(1e154;1e154)", 0.0),
        ("=SUMXMY2(1e154;1e154)", 0.0),
    ] {
        assert_close(source, expected);
    }
    assert_formula_error("=SUMX2PY2(1e154;1e154)", ScalarError::Number);
    assert_formula_error(
        "=SUM(1.7976931348623157e308;1.7976931348623157e308)",
        ScalarError::Number,
    );
    assert_close(
        "=PRODUCT(1e308;1e308;1e-308;1e-308)",
        f64::from_bits(0x3fefffffffffffff),
    );
    assert_close(
        "=PRODUCT(1e-308;1e-308;1e308;1e308)",
        f64::from_bits(0x3fefffffffffffff),
    );
    assert_close("=SUM(5e-324;5e-324)", 1e-323);
}

#[test]
fn scalar_aggregate_limits_and_reference_capability_are_typed() {
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-limits");
    let error = evaluate(
        "=SUM(1;2;3;4;5;6;7;8;9;10)",
        &execution,
        EvaluationLimits::default().with_max_steps(0),
    )
    .expect_err("zero Work must refuse before reduction");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong zero-work error: {error:?}"
    );
    assert_unsupported("=SUM([.A1:.A2])", UnsupportedKind::Reference);
    assert_unsupported("=SUMPRODUCT({1;2|3;4})", UnsupportedKind::Array);
}
