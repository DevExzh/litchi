//! Independent scalar integration coverage for the OpenFormula 1.4
//! trigonometric and angle-conversion functions in Part 4 §6.16.
//!
//! The expected values below come from the mathematical definitions in the
//! specification and the host's correctly rounded elementary functions.  The
//! tests deliberately exercise the finite-`f64` profile at its domain edges:
//! formula-domain failures remain formula values, while evaluator limits,
//! cancellation, and unsupported references remain typed Rust failures.

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
    let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-close");
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
    let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-exact");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_eq!(number(&result).to_bits(), expected.to_bits(), "{source:?}");
}

fn assert_golden(source: &str, expected: f64) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-golden");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    let actual = number(&result);
    // The high-precision reference is fixed, while the implementation's libm
    // operations may differ in their last bits across supported platforms.
    // Equal signs make bit-distance the ULP distance, including subnormals.
    assert!(actual.is_finite() && actual.is_sign_negative() == expected.is_sign_negative());
    assert!(
        actual.to_bits().abs_diff(expected.to_bits()) <= 4,
        "{source:?}: {actual:e} versus high-precision golden {expected:e}"
    );
}

fn assert_formula_error(source: &str, expected: ScalarError) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-error");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
    assert_eq!(formula_error(&result), expected, "{source:?}");
}

fn assert_unsupported(source: &str, expected: UnsupportedKind) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-unsupported");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("reference-dependent trigonometry should be refused");
    assert!(
        matches!(&error, EvaluationFailure::Unsupported(kind) if *kind == expected),
        "{source:?} returned the wrong capability error: {error}"
    );
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

#[test]
fn every_section_616_trigonometric_function_has_a_scalar_vector() {
    let cases = [
        ("=ACOS(0.5)", PI / 3.0),
        ("=ACOSH(2)", 2.0_f64.acosh()),
        ("=ACOT(-1)", acot(-1.0)),
        ("=ACOTH(-2)", acoth(-2.0)),
        ("=ASIN(0.5)", PI / 6.0),
        ("=ASINH(-1)", (-1.0_f64).asinh()),
        ("=ATAN(1)", PI / 4.0),
        ("=ATAN2(1;1)", PI / 4.0),
        ("=ATANH(0.5)", 0.5_f64.atanh()),
        ("=COS(PI()/3)", 0.5),
        ("=COSH(1)", 1.0_f64.cosh()),
        ("=COT(PI()/4)", 1.0),
        ("=COTH(1)", 1.0_f64.tanh().recip()),
        ("=CSC(PI()/2)", 1.0),
        ("=CSCH(1)", 1.0_f64.sinh().recip()),
        ("=SEC(PI()/3)", 2.0),
        ("=SECH(0)", 1.0),
        ("=SIN(PI()/6)", 0.5),
        ("=SINH(1)", 1.0_f64.sinh()),
        ("=TAN(PI()/4)", 1.0),
        ("=TANH(1)", 1.0_f64.tanh()),
        ("=DEGREES(PI())", 180.0),
        ("=RADIANS(180)", PI),
        ("=PI()", PI),
    ];
    for (source, expected) in cases {
        assert_close(source, expected);
    }
}

#[test]
fn inverse_functions_and_reciprocals_follow_specification_boundaries() {
    for (source, expected) in [
        ("=ACOS(1)", 0.0),
        ("=ACOS(-1)", PI),
        ("=ASIN(1)", PI / 2.0),
        ("=ASIN(-1)", -PI / 2.0),
        ("=ATAN(0)", 0.0),
        ("=ATAN(-0)", -0.0),
        ("=ACOT(0)", PI / 2.0),
        ("=ACOT(1e308)", acot(1e308)),
        ("=ACOTH(2)", acoth(2.0)),
        ("=ACOTH(-2)", acoth(-2.0)),
        // Both near-domain-edge values are valid because ABS(N) > 1.  The
        // negative case also exercises the odd extension without evaluating
        // a mathematically valid result as an infinity.
        ("=ACOTH(1.0000000000000002)", acoth(1.0000000000000002)),
        ("=ACOTH(-1.0000000000000002)", acoth(-1.0000000000000002)),
        ("=ATANH(0)", 0.0),
        ("=ATANH(-0)", -0.0),
    ] {
        assert_close(source, expected);
    }

    for (source, expected) in [
        ("=SIN(-PI()/2)", -1.0),
        ("=COS(-PI()/3)", 0.5),
        ("=TAN(-PI()/4)", -1.0),
        ("=COT(-PI()/4)", -1.0),
        ("=CSC(-PI()/2)", -1.0),
        ("=SEC(-PI()/3)", 2.0),
        ("=SINH(-1)", -1.0_f64.sinh()),
        ("=COSH(-1)", 1.0_f64.cosh()),
        ("=TANH(-1)", -1.0_f64.tanh()),
        ("=COTH(-1)", -1.0_f64.tanh().recip()),
        ("=CSCH(-1)", -1.0_f64.sinh().recip()),
        ("=SECH(-1)", 1.0_f64.cosh().recip()),
    ] {
        assert_close(source, expected);
    }
}

#[test]
fn atan2_uses_the_specified_x_then_y_order_and_all_quadrants() {
    for (source, expected) in [
        ("=ATAN2(1;1)", PI / 4.0),
        ("=ATAN2(-1;1)", 3.0 * PI / 4.0),
        ("=ATAN2(-1;-1)", -3.0 * PI / 4.0),
        ("=ATAN2(1;-1)", -PI / 4.0),
        ("=ATAN2(0;1)", PI / 2.0),
        ("=ATAN2(1;0)", 0.0),
        ("=ATAN2(0;-1)", -PI / 2.0),
        ("=ATAN2(-1;0)", PI),
        // The principal range excludes -PI, so the negative-zero y-axis
        // branch is normalized to +PI.
        ("=ATAN2(-1;-0)", PI),
        // A nonzero y-coordinate may round to -PI at binary64 precision.
        // It is still distinct from the exact negative-zero axis and must not
        // be shifted by 2*PI merely because the rounded value equals -PI.
        ("=ATAN2(-1;-1e-20)", -PI),
        ("=ATAN2(-1;-5e-324)", -PI),
    ] {
        assert_close(source, expected);
    }

    // Part 4 leaves ATAN2(0;0) implementation-defined.  This profile makes
    // the documented finite-domain choice explicit as a Number error.
    assert_formula_error("=ATAN2(0;0)", ScalarError::Number);
}

#[test]
fn negative_zero_signs_are_preserved_where_the_mathematics_is_odd() {
    for source in [
        "=ASIN(-0)",
        "=ASINH(-0)",
        "=ATAN(-0)",
        "=ATANH(-0)",
        "=SIN(-0)",
        "=SINH(-0)",
        "=TAN(-0)",
        "=TANH(-0)",
        "=DEGREES(-0)",
        "=RADIANS(-0)",
    ] {
        assert_exact(source, -0.0);
    }

    for source in ["=ACOS(-0)", "=COS(-0)", "=COSH(-0)", "=SECH(-0)"] {
        let expression = parse(source);
        let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-even-zero");
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert!(
            !number(&result).is_sign_negative(),
            "{source:?} returned -0"
        );
    }
}

#[test]
fn finite_extremes_and_subnormals_remain_finite_or_report_number() {
    for (source, expected) in [
        ("=SIN(1e308)", (1e308_f64).sin()),
        ("=COS(1e308)", (1e308_f64).cos()),
        ("=TAN(1e308)", (1e308_f64).tan()),
        ("=ACOSH(1e308)", (1e308_f64).acosh()),
        ("=ASINH(-1e308)", (-1e308_f64).asinh()),
        ("=COSH(700)", 700.0_f64.cosh()),
        ("=SINH(-700)", (-700.0_f64).sinh()),
        ("=TANH(1e308)", 1.0),
        ("=COTH(1e308)", 1.0),
        ("=SECH(700)", 700.0_f64.cosh().recip()),
        ("=CSCH(700)", 700.0_f64.sinh().recip()),
    ] {
        assert_close(source, expected);
    }

    // Fixed binary64 goldens were rounded from 600-decimal-place evaluations
    // of the Part 4 definitions.  In particular, (x + 1) / (x - 1) loses the
    // entire signal for ACOTH(±1e308) when it is formed at ordinary f64
    // precision, even though the result is a representable subnormal.
    assert_golden("=ACOTH(1e308)", f64::from_bits(0x0007_30d6_7819_e8d2));
    assert_golden("=ACOTH(-1e308)", -f64::from_bits(0x0007_30d6_7819_e8d2));
    assert_golden("=ACOSH(1e308)", f64::from_bits(0x4086_2f1d_6695_e8ec));
    assert_golden("=ASINH(1e308)", f64::from_bits(0x4086_2f1d_6695_e8ec));
    assert_golden("=ASINH(-1e308)", -f64::from_bits(0x4086_2f1d_6695_e8ec));

    // Hyperbolic reciprocals cross from ordinary subnormals into the
    // minimum-subnormal tail while the original hyperbolic values remain
    // finite.  Keeping these values nonzero catches premature exp underflow.
    assert_golden("=SECH(710)", f64::from_bits(0x0006_7005_fa51_6786));
    assert_golden("=CSCH(710)", f64::from_bits(0x0006_7005_fa51_6786));
    assert_golden("=SECH(720)", f64::from_bits(0x0000_0013_2769_b92a));
    assert_golden("=CSCH(720)", f64::from_bits(0x0000_0013_2769_b92a));

    let minimum = f64::from_bits(1);
    for (source, expected) in [
        ("=SIN(5e-324)", minimum),
        ("=ASIN(5e-324)", minimum),
        ("=ATAN(5e-324)", minimum),
        ("=SINH(5e-324)", minimum),
        ("=TAN(5e-324)", minimum),
        ("=TANH(5e-324)", minimum),
    ] {
        assert_exact(source, expected);
    }
    for source in [
        "=COS(5e-324)",
        "=COSH(5e-324)",
        "=SEC(5e-324)",
        "=SECH(5e-324)",
    ] {
        assert_exact(source, 1.0);
    }

    // These reciprocals have a representable zero result at the far end of
    // the finite domain, rather than an evaluator failure.
    assert_exact("=CSCH(1e308)", 0.0);
    assert_exact("=SECH(1e308)", 0.0);

    // 2*exp(-x) is still one minimum subnormal at this point.  An
    // implementation that computes the exponential in an underflowing
    // intermediate loses a valid finite result too early.
    let minimum = f64::from_bits(1);
    assert_exact("=CSCH(745.5)", minimum);
    assert_exact("=SECH(745.5)", minimum);
    assert_exact("=CSCH(745.8)", minimum);
    assert_exact("=SECH(745.8)", minimum);

    // The input is finite, but the result is outside the finite Number profile.
    for source in ["=COSH(711)", "=SINH(711)"] {
        assert_formula_error(source, ScalarError::Number);
    }
}

#[test]
fn number_conversion_accepts_numeric_text_and_logicals_but_rejects_bad_text() {
    for (source, expected) in [
        ("=SIN(\"0.5\")", 0.5_f64.sin()),
        ("=COS(TRUE())", 1.0_f64.cos()),
        ("=TAN(FALSE())", 0.0),
        ("=ACOS(\"1\")", 0.0),
        ("=DEGREES(\"3.141592653589793\")", 180.0),
    ] {
        assert_close(source, expected);
    }
    assert_formula_error("=SIN(\"not-a-number\")", ScalarError::Value);
    assert_formula_error("=ATAN2(\"bad\";1)", ScalarError::Value);
}

#[test]
fn formula_errors_and_arity_are_values_with_leftmost_error_precedence() {
    for (source, expected) in [
        ("=SIN(#N/A)", ScalarError::NotAvailable),
        ("=COS(#DIV/0!)", ScalarError::DivisionByZero),
        ("=ATAN2(#DIV/0!;#N/A)", ScalarError::DivisionByZero),
        ("=PI(1)", ScalarError::Value),
        ("=PI(1;2)", ScalarError::Value),
    ] {
        assert_formula_error(source, expected);
    }

    for name in [
        "ACOS", "ACOSH", "ACOT", "ACOTH", "ASIN", "ASINH", "ATAN", "ATANH", "COS", "COSH", "COT",
        "COTH", "CSC", "CSCH", "SEC", "SECH", "SIN", "SINH", "TAN", "TANH", "DEGREES", "RADIANS",
    ] {
        assert_formula_error(&format!("={name}()"), ScalarError::Value);
        assert_formula_error(&format!("={name}(1;2)"), ScalarError::Value);
    }
    assert_formula_error("=ATAN2(1)", ScalarError::Value);
    assert_formula_error("=ATAN2(1;2;3)", ScalarError::Value);
}

#[test]
fn constrained_domains_return_number_errors_without_panicking() {
    for source in [
        "=ACOS(1.0000000000000002)",
        "=ACOS(-1.0000000000000002)",
        "=ASIN(1.0000000000000002)",
        "=ASIN(-1.0000000000000002)",
        "=ACOSH(0.999)",
        "=ACOSH(-1)",
        "=ACOTH(1)",
        "=ACOTH(-1)",
        "=ATANH(1)",
        "=ATANH(-1)",
    ] {
        assert_formula_error(source, ScalarError::Number);
    }
    for source in ["=COTH(0)", "=CSC(0)", "=CSCH(0)", "=COT(0)"] {
        assert_formula_error(source, ScalarError::DivisionByZero);
    }
}

#[test]
fn conditional_evaluation_is_lazy_for_unsupported_references() {
    let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-lazy");
    for (source, expected) in [
        ("=IF(TRUE();SIN(PI()/2);COS([.A1]))", 1.0),
        ("=IF(FALSE();SIN([.A1]);COS(0))", 1.0),
        ("=IFERROR(SIN(\"bad\");2)", 2.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate lazily: {error}"));
        assert_close_value(source, number(&result), expected);
    }
    assert_unsupported("=SIN([.A1])", UnsupportedKind::Reference);
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
fn work_text_and_cancellation_limits_refuse_before_publishing() {
    let expression = parse("=SIN(PI())");
    let (budget, _cancellation, context) = execution("ods-formula-trigonometry-work");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(0),
    )
    .expect_err("zero work must refuse before publishing a trigonometric result");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "unexpected work refusal: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let expression = parse("=SIN(\"0.5\")");
    let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-text-limit");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_text_bytes(2),
    )
    .expect_err("the three-byte numeric text must exceed the two-byte limit");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "unexpected text refusal: {error}"
    );

    let (budget, cancellation, context) = execution("ods-formula-trigonometry-cancel");
    cancellation.cancel();
    let error = evaluate(&parse("=SIN(PI())"), &context, &EvaluationLimits::default())
        .expect_err("pre-cancelled trigonometry must refuse");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn scalar_work_limit_accepts_the_exact_required_budget_and_rejects_one_less() {
    let expression = parse("=SIN(PI())");
    let (budget, _cancellation, context) = execution("ods-formula-trigonometry-work-measure");
    evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("the unbounded reference run should evaluate");
    let required = budget.used(Resource::Work);
    assert!(required > 0, "the formula must consume measurable work");

    let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-work-below");
    let below = EvaluationLimits::default().with_max_steps(required - 1);
    let error = evaluate(&expression, &context, &below)
        .expect_err("one work unit below the measured requirement must refuse");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "unexpected below-boundary refusal: {error}"
    );

    let (_budget, _cancellation, context) = execution("ods-formula-trigonometry-work-exact");
    let exact = EvaluationLimits::default().with_max_steps(required);
    evaluate(&expression, &context, &exact)
        .expect("the exact measured work budget must be sufficient");
}
