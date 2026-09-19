//! Independent scalar integration coverage for the OpenFormula 1.4
//! mathematical functions added in §6.16.
//!
//! The expected values are derived from the function definitions and from
//! exact binary64/high-precision oracles retained with the validation
//! evidence.  Formula-level domain failures stay scalar formula values;
//! capability, resource, and cancellation failures remain typed evaluator
//! failures.

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

const PI: f64 = std::f64::consts::PI;

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

fn assert_close(source: &str, expected: f64) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-elementary-close");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    let actual = number(&result);
    if expected == 0.0 {
        assert_eq!(
            actual.to_bits(),
            expected.to_bits(),
            "{source:?}: {actual:?}"
        );
        return;
    }
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

fn assert_exact(source: &str, expected: f64) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-elementary-exact");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_eq!(number(&result).to_bits(), expected.to_bits(), "{source:?}");
}

fn assert_golden(source: &str, expected_bits: u64) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-elementary-golden");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    let actual = number(&result);
    let expected = f64::from_bits(expected_bits);
    assert!(
        actual.is_finite() && actual.is_sign_negative() == expected.is_sign_negative(),
        "{source:?}: sign or finiteness differs ({actual:?} versus {expected:?})"
    );
    assert!(
        actual.to_bits().abs_diff(expected_bits) <= 4,
        "{source:?}: {actual:e} versus high-precision golden {expected:e}"
    );
}

fn assert_bits(source: &str, expected_bits: u64) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-elementary-bits");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_eq!(number(&result).to_bits(), expected_bits, "{source:?}");
}

fn assert_formula_error(source: &str, expected: ScalarError) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-elementary-error");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
    assert_eq!(formula_error(&result), expected, "{source:?}");
}

fn assert_unsupported(source: &str, expected: UnsupportedKind) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-elementary-unsupported");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("reference-dependent elementary expression should be refused");
    assert!(
        matches!(&error, EvaluationFailure::Unsupported(kind) if *kind == expected),
        "{source:?} returned the wrong capability error: {error}"
    );
}

#[test]
fn every_elementary_function_has_a_scalar_vector() {
    for (source, expected) in [
        ("=ABS(-3.5)", 3.5),
        ("=EXP(1)", 1.0_f64.exp()),
        ("=LN(EXP(1))", 1.0),
        ("=LOG(1000)", 3.0),
        ("=LOG10(1000)", 3.0),
        ("=POWER(2;10)", 1024.0),
        ("=SQRT(2)", 2.0_f64.sqrt()),
        ("=SQRTPI(2)", (2.0 * PI).sqrt()),
        ("=SIGN(-3.5)", -1.0),
        ("=MOD(22;3)", 1.0),
        ("=QUOTIENT(11;3)", 3.0),
    ] {
        assert_close(source, expected);
    }
}

#[test]
fn log_default_base_and_function_arity_follow_the_catalog() {
    for (source, expected) in [("=LOG(1000)", 3.0), ("=LOG(1000;10)", 3.0)] {
        assert_close(source, expected);
    }
    // An explicitly empty second slot is a supplied Missing argument. The
    // OpenFormula default applies only when the optional parameter is absent.
    assert_formula_error("=LOG(1000;)", ScalarError::Value);

    for source in [
        "=ABS()",
        "=ABS(1;2)",
        "=EXP()",
        "=LN(1;2)",
        "=LOG()",
        "=LOG(1;10;100)",
        "=LOG10()",
        "=POWER(2)",
        "=POWER(2;3;4)",
        "=SQRT()",
        "=SQRTPI(1;2)",
        "=SIGN()",
        "=MOD(1)",
        "=MOD(1;2;3)",
        "=QUOTIENT(1)",
        "=QUOTIENT(1;2;3)",
    ] {
        assert_formula_error(source, ScalarError::Value);
    }
}

#[test]
fn elementary_number_coercion_and_formula_error_precedence_are_explicit() {
    for (source, expected) in [
        ("=ABS(\"-2.5\")", 2.5),
        ("=EXP(\"0\")", 1.0),
        ("=LOG(\"100\";\"10\")", 2.0),
        ("=POWER(TRUE();3)", 1.0),
        ("=SIGN(FALSE())", 0.0),
        ("=MOD(\"5\";2)", 1.0),
        ("=QUOTIENT(\"5\";2)", 2.0),
    ] {
        assert_close(source, expected);
    }
    for source in [
        "=ABS(\"not-a-number\")",
        "=EXP(\"not-a-number\")",
        "=LOG(\"not-a-number\")",
        "=POWER(\"not-a-number\";2)",
        "=MOD(\"not-a-number\";2)",
        "=QUOTIENT(\"not-a-number\";2)",
    ] {
        assert_formula_error(source, ScalarError::Value);
    }

    for (source, expected) in [
        ("=ABS(#N/A)", ScalarError::NotAvailable),
        ("=MOD(#N/A;0)", ScalarError::NotAvailable),
        ("=MOD(1;#N/A)", ScalarError::NotAvailable),
        ("=LOG(\"bad\";#N/A)", ScalarError::NotAvailable),
        ("=LOG(#N/A;\"bad\")", ScalarError::NotAvailable),
        ("=POWER(#DIV/0!;2)", ScalarError::DivisionByZero),
        ("=QUOTIENT(1;#DIV/0!)", ScalarError::DivisionByZero),
    ] {
        assert_formula_error(source, expected);
    }
}

#[test]
fn logarithm_and_root_domains_return_number_errors() {
    for source in [
        "=LN(0)",
        "=LN(-1)",
        "=LOG(0)",
        "=LOG(-1)",
        "=LOG(10;0)",
        "=LOG(10;-2)",
        "=LOG(10;1)",
        "=LOG10(0)",
        "=LOG10(-1)",
        "=SQRT(-1)",
        "=SQRTPI(-1)",
        "=POWER(0;-1)",
        "=POWER(-4;0.5)",
    ] {
        assert_formula_error(source, ScalarError::Number);
    }

    // The selected finite profile accepts the implementation-defined 0^0
    // choice as one, as permitted by Part 4 §6.16.46.
    assert_exact("=POWER(0;0)", 1.0);
}

#[test]
fn powers_preserve_integer_sign_parity_and_match_the_caret_operator() {
    for (source, expected) in [
        ("=POWER(-2;3)", -8.0),
        ("=POWER(-2;4)", 16.0),
        ("=POWER(-2;-3)", -0.125),
        ("=(-2)^3", -8.0),
        ("=(-2)^4", 16.0),
        ("=(-2)^-3", -0.125),
        ("=POWER(4;1.5)", 8.0),
        ("=POWER(4;-1.5)", 0.125),
    ] {
        assert_close(source, expected);
    }

    for source in ["=POWER(2;1024)", "=POWER(10;400)", "=2^1024"] {
        assert_formula_error(source, ScalarError::Number);
    }
}

#[test]
fn mod_and_quotient_follow_divisor_sign_and_truncate_toward_zero() {
    for (source, expected) in [
        ("=MOD(3;2)", 1.0),
        ("=MOD(-3;2)", 1.0),
        ("=MOD(3;-2)", -1.0),
        ("=MOD(-3;-2)", -1.0),
        ("=MOD(1;-3)", -2.0),
        ("=MOD(-1;-3)", -1.0),
        ("=MOD(11.25;2.5)", 1.25),
        ("=MOD(7;2.1)", 0.7),
        ("=QUOTIENT(10;3)", 3.0),
        ("=QUOTIENT(-10;3)", -3.0),
        ("=QUOTIENT(10;-3)", -3.0),
        ("=QUOTIENT(-10;-3)", 3.0),
    ] {
        assert_close(source, expected);
    }
    assert_exact("=QUOTIENT(-5;10)", -0.0);
    assert_exact("=QUOTIENT(5;-10)", -0.0);
    assert_exact("=MOD(-0;2)", 0.0);
    assert_exact("=MOD(0;-2)", 0.0);
    for source in ["=MOD(1;0)", "=QUOTIENT(1;0)"] {
        assert_formula_error(source, ScalarError::DivisionByZero);
    }
}

#[test]
fn finite_extremes_subnormals_and_intermediate_root_products_remain_stable() {
    // These bits come from the retained 600-digit mpmath oracle. Four ULPs
    // allow the host libm and represented PI multiplication to round.
    for (source, expected) in [
        ("=SQRTPI(1.7976931348623157e308)", 0x5ffc5bf891b4ef6a),
        ("=SQRTPI(5e-324)", 0x1e6c5bf891b4ef6b),
        ("=LN(5e-324)", 0xc0874385446d71c3),
        ("=LN(1.0000000000000002)", 0x3cafffffffffffff),
        ("=LOG10(5e-324)", 0xc07434e6420f4374),
        ("=LOG(5e-324;1.0000000000000002)", 0xc3c74385446d71c4),
        (
            "=LOG(1.0000000000000002;0.9999999999999999)",
            0xbfffffffffffffff,
        ),
        ("=EXP(709)", 0x7fdd422d2be5dc9b),
    ] {
        assert_golden(source, expected);
    }
    assert_bits("=EXP(-745)", 0x0000000000000001);
    assert_exact("=ABS(-5e-324)", f64::from_bits(1));
    assert_exact("=SIGN(-5e-324)", -1.0);
    assert_close("=EXP(-1000)", 0.0);
}

#[test]
fn large_quotient_remainders_use_the_finite_binary64_definition() {
    for (source, expected) in [
        ("=MOD(1e308;3)", 0x4000000000000000),
        ("=MOD(-1e308;3)", 0x3ff0000000000000),
        ("=MOD(1e308;-3)", 0xbff0000000000000),
        ("=MOD(1.677259342285726e21;77)", 0x4022000000000000),
        ("=MOD(-5e-324;1.7976931348623157e308)", 0x7fefffffffffffff),
        ("=MOD(1.7976931348623157e308;1e-300)", 0x0151c210546b28c0),
    ] {
        assert_bits(source, expected);
    }
}

#[test]
fn elementary_functions_remain_lazy_and_keep_references_out_of_formula_values() {
    let (_budget, _cancellation, context) = execution("ods-formula-elementary-lazy");
    for (source, expected) in [
        ("=IF(TRUE();ABS(-2);SQRT([.A1]))", 2.0),
        ("=IF(FALSE();SQRT([.A1]);ABS(-2))", 2.0),
        ("=IFERROR(SQRT(-1);42)", 42.0),
        ("=IFERROR(LOG(10;1);17)", 17.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate lazily: {error}"));
        assert_close_value(source, number(&result), expected);
    }
    assert_unsupported("=ABS([.A1])", UnsupportedKind::Reference);
    assert_unsupported("=MOD(1;[.A1])", UnsupportedKind::Reference);
}

fn assert_close_value(source: &str, actual: f64, expected: f64) {
    if expected == 0.0 {
        assert_eq!(actual.to_bits(), expected.to_bits(), "{source:?}");
    } else {
        let tolerance = expected.abs().max(actual.abs()) * (32.0 * f64::EPSILON);
        assert!(
            (actual - expected).abs() <= tolerance,
            "{source:?}: {actual} != {expected}"
        );
    }
}

#[test]
fn scalar_work_storage_and_cancellation_fail_atomically_and_refund_text() {
    let expression = parse("=SQRT(4)");
    let (budget, _cancellation, context) = execution("ods-formula-elementary-work");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(0),
    )
    .expect_err("zero work must refuse before publishing an elementary result");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "unexpected work refusal: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation, context) = execution("ods-formula-elementary-cancel");
    cancellation.cancel();
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("pre-cancelled elementary evaluation must refuse");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Memory), 0);

    let expression = parse("=ABS(\"1\"&\"2\")");
    let (budget, _cancellation, context) = execution("ods-formula-elementary-storage");
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
        .expect("the same expression should fit the default storage limit");
    assert_eq!(number(&result), 12.0);
    assert_eq!(result.reserved_output_bytes(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}
