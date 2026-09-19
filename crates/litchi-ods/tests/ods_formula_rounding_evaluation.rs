//! Independent scalar coverage for the OpenFormula 1.4 rounding family.
//!
//! The vectors in this file follow Part 4 section 6.17.  They deliberately
//! keep the profile finite and locale-independent: CEILING and FLOOR use the
//! specified sign-compatible significance rules, MROUND chooses the greater
//! multiple on a tie, ROUND is halfway-away-from-zero, and the other three
//! digit functions use their specified directed rounding.

use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluatedScalar, EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError,
        ScalarValue, UnsupportedKind, evaluate_scalar,
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

fn evaluate<'a>(
    expression: &'a Expression,
    context: &ExecutionContext,
    limits: &EvaluationLimits,
) -> Result<EvaluatedScalar<'a>, EvaluationFailure> {
    evaluate_scalar(expression, &EvaluationContext::new(context), limits)
}

fn number(result: &EvaluatedScalar<'_>) -> f64 {
    match result.value() {
        ScalarValue::Number(value) => *value,
        value => panic!("expected Number, got {value:?}"),
    }
}

fn formula_error(result: &EvaluatedScalar<'_>) -> ScalarError {
    match result.value() {
        ScalarValue::Error(error) => *error,
        value => panic!("expected formula Error, got {value:?}"),
    }
}

fn assert_number(source: &str, expected: f64) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-rounding-number");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_eq!(number(&result), expected, "{source:?}");
}

fn assert_number_close(source: &str, expected: f64) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-rounding-number-close");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    let actual = number(&result);
    // A fixed absolute tolerance would turn every subnormal result into a
    // passing zero.  Keep zero exact and compare nonzero values at a small
    // multiple of their magnitude instead.
    if expected == 0.0 {
        assert_eq!(actual, 0.0, "{source:?}: {actual} != {expected}");
        return;
    }
    assert!(actual.is_finite(), "{source:?}: non-finite result {actual}");
    let tolerance = expected.abs().max(actual.abs()) * (16.0 * f64::EPSILON);
    assert!(
        (actual - expected).abs() <= tolerance,
        "{source:?}: {actual} != {expected}"
    );
}

fn assert_positive_zero(source: &str) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-rounding-positive-zero");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    let actual = number(&result);
    assert_eq!(actual.to_bits(), 0, "{source:?} returned {actual:?}");
}

fn assert_formula_error(source: &str, expected: ScalarError) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-rounding-error");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
    assert_eq!(formula_error(&result), expected, "{source:?}");
}

fn assert_unsupported(source: &str, expected: UnsupportedKind) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-rounding-unsupported");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("reference-dependent rounding should be refused without a resolver");
    assert!(
        matches!(&error, EvaluationFailure::Unsupported(kind) if *kind == expected),
        "{source:?} returned the wrong capability error: {error}"
    );
}

#[test]
fn all_eight_functions_have_independent_scalar_vectors() {
    for (source, expected) in [
        ("=CEILING(4.2;2)", 6.0),
        ("=INT(4.9)", 4.0),
        ("=FLOOR(4.2;2)", 4.0),
        ("=MROUND(4.2;2)", 4.0),
        ("=ROUND(4.25;1)", 4.3),
        ("=ROUNDDOWN(4.29;1)", 4.2),
        ("=ROUNDUP(4.21;1)", 4.3),
        ("=TRUNC(4.29;1)", 4.2),
    ] {
        assert_number_close(source, expected);
    }
}

#[test]
fn negative_values_and_ceiling_floor_modes_follow_section_617() {
    for (source, expected) in [
        ("=INT(-4.2)", -5.0),
        ("=CEILING(-4.2)", -4.0),
        ("=CEILING(-4.2;;1)", -5.0),
        ("=CEILING(-4.2;-2)", -4.0),
        ("=CEILING(-4.2;-2;1)", -6.0),
        ("=CEILING(-5.3;-2;0)", -4.0),
        ("=CEILING(-5.3;-2;0.5)", -6.0),
        ("=FLOOR(-4.2)", -5.0),
        ("=FLOOR(-4.2;;1)", -4.0),
        ("=FLOOR(-4.2;-2)", -6.0),
        ("=FLOOR(-4.2;-2;1)", -4.0),
        ("=FLOOR(-5.3;-2;0)", -6.0),
        ("=FLOOR(-5.3;-2;0.5)", -4.0),
        ("=CEILING(4.2;2;1)", 6.0),
        ("=FLOOR(4.2;2;1)", 4.0),
        ("=CEILING(4.2;1.5)", 4.5),
        ("=FLOOR(4.2;1.5)", 3.0),
    ] {
        assert_number_close(source, expected);
    }
}

#[test]
fn fn_zero_and_omitted_or_empty_optional_arguments_are_distinct_inputs() {
    for (source, expected) in [
        ("=CEILING(0;2)", 0.0),
        ("=FLOOR(0;2)", 0.0),
        ("=CEILING(4.2;0)", 0.0),
        ("=FLOOR(-4.2;0)", 0.0),
        ("=ROUND(4.2)", 4.0),
        ("=ROUND(4.2;)", 4.0),
        ("=ROUNDDOWN(-4.2)", -4.0),
        ("=ROUNDDOWN(-4.2;)", -4.0),
        ("=ROUNDUP(-4.2)", -5.0),
        ("=TRUNC(-4.2;)", -4.0),
        ("=TRUNC(-4.2)", -4.0),
        ("=CEILING(4.2;;)", 5.0),
        ("=FLOOR(-4.2;;)", -5.0),
    ] {
        assert_number_close(source, expected);
    }
    assert_formula_error("=MROUND(4.2;0)", ScalarError::DivisionByZero);
    assert_formula_error("=MROUND(0;0)", ScalarError::DivisionByZero);
}

#[test]
fn halfway_rules_are_away_from_zero_or_choose_the_greater_multiple() {
    for (source, expected) in [
        ("=ROUND(2.5)", 3.0),
        ("=ROUND(-2.5)", -3.0),
        ("=MROUND(5;2)", 6.0),
        ("=MROUND(-5;2)", -4.0),
        ("=MROUND(5;-2)", 6.0),
    ] {
        assert_number(source, expected);
    }
}

#[test]
fn digit_arguments_control_decimal_and_left_of_decimal_rounding() {
    for (source, expected) in [
        ("=ROUND(12.345;2)", 12.35),
        ("=ROUND(-12.345;2)", -12.35),
        // ROUND declares Digits as Number, so 2.5 is an exponent rather than
        // an Integer conversion.  This is the independent 10^-2.5 result.
        ("=ROUND(12.345;2.5)", 12.345531985297352),
        ("=ROUND(1234;-2)", 1200.0),
        ("=ROUNDDOWN(-1234.9;-2)", -1200.0),
        ("=ROUNDUP(1200.1;-2)", 1300.0),
        ("=TRUNC(-1234.9;-2)", -1200.0),
        ("=ROUNDDOWN(12.99;2.9)", 12.99),
        ("=ROUNDUP(-12.01;2.9)", -12.01),
        ("=TRUNC(12.99;2.9)", 12.99),
        // Integer powers must preserve the exact decimal quantum at the
        // largest finite exponent used by f64.
        ("=ROUNDDOWN(1e308;-308)", 1e308),
        ("=TRUNC(1e23;-23)", 1e23),
        ("=ROUNDUP(1e23;-23)", 1e23),
    ] {
        assert_number_close(source, expected);
    }
}

#[test]
fn fractional_decimal_exponents_keep_their_finite_quantum() {
    // Digits is a Number in ROUND, so fractional exponents use a fractional
    // power rather than truncating to an integer exponent.
    assert_number_close("=ROUND(1e308;-308.1)", 1.258_925_411_794_233e308);

    // 10^-323.1 is still representable as two minimum subnormal quanta.
    // This exercises the boundary just below the point where decimal powers
    // underflow to zero.
    assert_number("=ROUND(5e-324;323.1)", f64::from_bits(2));
    assert_number("=ROUNDUP(5e-324;323.1)", f64::from_bits(2));
}

#[test]
fn directed_decimal_rounding_preserves_exact_decimal_multiples() {
    assert_number("=ROUNDDOWN(0.3;1)", 0.3);
    assert_number("=TRUNC(0.3;1)", 0.3);
    assert_number("=ROUNDUP(1.2;1)", 1.2);
    assert_number("=ROUNDUP(1e-20;0)", 1.0);
    assert_number("=ROUNDUP(-1e-20;0)", -1.0);
    assert_number("=ROUNDUP(0.07;2)", 0.07);
    assert_number("=ROUNDUP(0.07000000000000002;2)", 0.08);
    assert_number("=ROUNDDOWN(1.15;2)", 1.15);
    assert_number("=ROUNDDOWN(1.1499999999999997;2)", 1.14);
}

#[test]
fn zero_results_are_published_with_a_positive_sign_bit() {
    for source in [
        "=ROUND(-0.1;0)",
        "=ROUNDDOWN(-0.1;0)",
        "=TRUNC(-0.1;0)",
        "=MROUND(-5e-324;1e-323)",
        "=CEILING(-5e-324;-1e-323)",
    ] {
        assert_positive_zero(source);
    }
}

#[test]
fn numeric_text_and_logical_arguments_follow_number_coercion() {
    for (source, expected) in [
        ("=ROUND(\"2.5\";0)", 3.0),
        ("=ROUND(TRUE();0)", 1.0),
        ("=ROUND(FALSE();0)", 0.0),
        ("=CEILING(\"4.2\";\"2\")", 6.0),
        ("=FLOOR(TRUE();TRUE())", 1.0),
        ("=ROUNDDOWN(\"12.99\";\"2.9\")", 12.99),
        ("=ROUNDUP(\"-12.01\";\"2.9\")", -12.01),
        ("=TRUNC(\"-12.99\";\"-1.9\")", -10.0),
    ] {
        assert_number_close(source, expected);
    }
}

#[test]
fn nonnumeric_text_is_a_formula_value_error_after_argument_coercion() {
    for source in [
        "=ROUND(\"not-a-number\";0)",
        "=ROUND(1;\"not-a-number\")",
        "=CEILING(\"not-a-number\";1)",
        "=CEILING(1;\"not-a-number\")",
        "=MROUND(\"not-a-number\";1)",
        "=ROUNDDOWN(1;\"not-a-number\")",
        "=ROUNDUP(1;\"not-a-number\")",
        "=TRUNC(1;\"not-a-number\")",
    ] {
        assert_formula_error(source, ScalarError::Value);
    }
}

#[test]
fn sign_constraints_and_unrepresentable_digit_parameters_are_typed_errors() {
    for source in [
        "=CEILING(4.2;-2)",
        "=FLOOR(4.2;-2)",
        "=CEILING(-4.2;2)",
        "=FLOOR(-4.2;2)",
    ] {
        assert_formula_error(source, ScalarError::Value);
    }

    // Directed decimal rounding remains defined for exponents beyond the
    // finite f64 power-of-ten range: toward zero reaches zero, while a
    // positive input rounded away from zero would be non-finite.  ROUNDUP's
    // very fine positive precision leaves a representable input unchanged.
    assert_number("=ROUNDDOWN(1;-1e100)", 0.0);
    assert_number("=TRUNC(1;-1e100)", 0.0);
    assert_number("=ROUNDUP(1;1e100)", 1.0);
}

#[test]
fn finite_near_overflow_values_remain_finite_when_the_result_is_representable() {
    assert_number("=TRUNC(1.7976931348623155e308;0)", 1.7976931348623155e308);
    assert_number_close("=ROUND(1e307;-2)", 1e307);
    assert_number_close("=ROUND(1e-307;307)", 1e-307);
    assert_number_close("=ROUND(1;1e100)", 1.0);
    assert_formula_error("=ROUND(1e309)", ScalarError::Number);
}

#[test]
fn subnormal_rounding_keeps_nonzero_ties_and_directed_zero_results_distinct() {
    let minimum = f64::from_bits(1);
    let two_minimum = f64::from_bits(2);

    // 5e-324 is the minimum positive f64 and is exactly halfway between zero
    // and the representable 10^-323 quantum.  ROUND and ROUNDUP are away from
    // zero on that tie; ROUNDDOWN and TRUNC reach zero.
    assert_number("=ROUND(5e-324;323)", two_minimum);
    assert_number("=ROUND(-5e-324;323)", -two_minimum);
    assert_number("=ROUNDDOWN(5e-324;323)", 0.0);
    assert_number("=ROUNDUP(5e-324;323)", two_minimum);
    assert_number("=TRUNC(-5e-324;323)", 0.0);

    // A finer-than-representable decimal place leaves the input untouched.
    assert_number("=ROUND(5e-324;324)", minimum);
    assert_number("=ROUNDUP(5e-324;324)", minimum);
}

#[test]
fn multiple_rounding_handles_subnormal_ties_without_collapsing_to_zero() {
    let two_minimum = f64::from_bits(2);

    assert_number("=CEILING(5e-324;1e-323)", two_minimum);
    assert_number("=FLOOR(5e-324;1e-323)", 0.0);
    assert_number("=MROUND(5e-324;1e-323)", two_minimum);
    // For a negative tie the greater numerical multiple is zero.
    assert_number("=MROUND(-5e-324;1e-323)", 0.0);
    assert_number("=CEILING(-5e-324;-1e-323)", 0.0);
    assert_number("=FLOOR(-5e-324;-1e-323)", -two_minimum);
}

#[test]
fn rounding_storage_refusal_is_atomic_and_temporary_text_reservations_refund() {
    let expression = parse("=ROUND(\"1\"&\"2\";0)");

    let (budget, _cancellation, context) = execution("ods-formula-rounding-storage-refusal");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_storage_bytes(1),
    )
    .expect_err("the two-byte concatenated argument must exceed one byte");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "unexpected storage refusal: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("the same rounding expression should fit the default storage limit");
    assert_eq!(number(&result), 12.0);
    assert_eq!(result.reserved_output_bytes(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn mround_reports_overflow_only_when_the_nearest_multiple_is_unrepresentable() {
    assert_number_close("=MROUND(1.1e308;1e308)", 1e308);
    assert_formula_error("=MROUND(1.7e308;1e308)", ScalarError::Number);
}

#[test]
fn original_error_precedence_is_preserved_across_the_family() {
    for (source, expected) in [
        ("=INT(#N/A)", ScalarError::NotAvailable),
        ("=INT(#DIV/0!)", ScalarError::DivisionByZero),
        ("=CEILING(#N/A;#DIV/0!;#VALUE!)", ScalarError::NotAvailable),
        ("=FLOOR(#DIV/0!;#N/A;#VALUE!)", ScalarError::DivisionByZero),
        ("=MROUND(#N/A;#DIV/0!)", ScalarError::NotAvailable),
        ("=ROUND(#N/A;#DIV/0!)", ScalarError::NotAvailable),
        ("=ROUNDDOWN(#DIV/0!;#N/A)", ScalarError::DivisionByZero),
        ("=ROUNDUP(#N/A;#DIV/0!)", ScalarError::NotAvailable),
        ("=TRUNC(#DIV/0!;#N/A)", ScalarError::DivisionByZero),
    ] {
        assert_formula_error(source, expected);
    }
}

#[test]
fn rounding_can_be_lazy_and_keeps_capability_failures_out_of_formula_values() {
    let (_budget, _cancellation, context) = execution("ods-formula-rounding-lazy");
    let selected = parse("=IF(TRUE();ROUND(2.5);ROUND([.A1];0))");
    let result = evaluate(&selected, &context, &EvaluationLimits::default())
        .expect("the unselected reference branch must not be visited");
    assert_eq!(number(&result), 3.0);

    assert_unsupported("=ROUND([.A1];0)", UnsupportedKind::Reference);
    assert_unsupported("=CEILING(1;[.A1])", UnsupportedKind::Reference);
}

#[test]
fn rounding_work_and_cancellation_fail_before_publishing_a_result() {
    let expression = parse("=ROUND(123.45;2)");
    let (budget, _cancellation, context) = execution("ods-formula-rounding-work");
    let limits = EvaluationLimits::default().with_max_steps(0);
    let error = evaluate(&expression, &context, &limits).expect_err("zero work must refuse");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "unexpected zero-work error: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation, context) = execution("ods-formula-rounding-cancel");
    cancellation.cancel();
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("pre-cancelled rounding must refuse");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Memory), 0);
}
