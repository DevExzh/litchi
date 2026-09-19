//! Independent scalar coverage for the OpenFormula 1.4 §6.16 discrete
//! mathematical functions.
//!
//! The value/reference bridge lives in `ods_formula_discrete_math_arrays.rs`.
//! This file keeps the scalar contract visible: defaults are only applied to
//! omitted optional arguments, formula errors win over conversion/domain
//! errors in source order, integer parameters are converted before the
//! discrete operation, and evaluator resource failures remain typed failures.

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
    execution: &ExecutionContext,
    limits: &EvaluationLimits,
) -> Result<EvaluatedScalar<'a>, EvaluationFailure> {
    evaluate_scalar(expression, &EvaluationContext::new(execution), limits)
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
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-number");
    let result = evaluate(&expression, &execution, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    let actual = number(&result);
    if expected == 0.0 {
        assert_eq!(actual.to_bits(), expected.to_bits(), "{source:?}");
    } else {
        assert!(
            actual.is_finite(),
            "{source:?} returned non-finite {actual:?}"
        );
        let tolerance = expected.abs().max(actual.abs()) * (16.0 * f64::EPSILON);
        assert!(
            (actual - expected).abs() <= tolerance,
            "{source:?}: {actual:?} != {expected:?} (tol {tolerance:e})"
        );
    }
}

fn assert_bits(source: &str, expected_bits: u64) {
    let expression = parse(source);
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-bits");
    let result = evaluate(&expression, &execution, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_eq!(number(&result).to_bits(), expected_bits, "{source:?}");
}

fn assert_formula_error(source: &str, expected: ScalarError) {
    let expression = parse(source);
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-error");
    let result = evaluate(&expression, &execution, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
    assert_eq!(formula_error(&result), expected, "{source:?}");
}

fn assert_unsupported(source: &str, expected: UnsupportedKind) {
    let expression = parse(source);
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-unsupported");
    let error = evaluate(&expression, &execution, &EvaluationLimits::default())
        .expect_err("a reference cannot be consumed by the scalar-only profile");
    assert!(
        matches!(&error, EvaluationFailure::Unsupported(kind) if *kind == expected),
        "{source:?} returned the wrong capability error: {error:?}"
    );
}

#[test]
fn every_discrete_function_has_a_scalar_definition_example() {
    for (source, expected) in [
        ("=COMBIN(5;2)", 10.0),
        ("=COMBINA(3;2)", 6.0),
        ("=FACT(5)", 120.0),
        ("=FACTDOUBLE(8)", 384.0),
        ("=GCD(48;18;30)", 6.0),
        ("=LCM(4;6;8)", 24.0),
        ("=MULTINOMIAL(2;3;4)", 1260.0),
        ("=EVEN(-2.5)", -4.0),
        ("=ODD(-2.5)", -3.0),
        ("=DELTA(3;3)", 1.0),
        ("=GESTEP(3;2)", 1.0),
    ] {
        assert_number(source, expected);
    }
}

#[test]
fn discrete_arity_and_optional_defaults_are_explicit() {
    for source in [
        "=COMBIN()",
        "=COMBIN(5)",
        "=COMBIN(5;2;1)",
        "=COMBINA()",
        "=COMBINA(5)",
        "=COMBINA(5;2;1)",
        "=FACT()",
        "=FACT(5;1)",
        "=FACTDOUBLE()",
        "=FACTDOUBLE(5;1)",
        "=GCD()",
        "=LCM()",
        "=MULTINOMIAL()",
        "=EVEN()",
        "=EVEN(1;2)",
        "=ODD()",
        "=ODD(1;2)",
        "=DELTA()",
        "=DELTA(1;2;3)",
        "=GESTEP()",
        "=GESTEP(1;2;3)",
    ] {
        assert_formula_error(source, ScalarError::Value);
    }

    // Omitted optional slots use the specified zero default.  A syntactic
    // missing value is a supplied argument and therefore is not a default.
    assert_number("=DELTA(0)", 1.0);
    assert_number("=DELTA(1)", 0.0);
    assert_number("=GESTEP(0)", 1.0);
    assert_number("=GESTEP(-1)", 0.0);
    assert_formula_error("=DELTA(0;)", ScalarError::Value);
    assert_formula_error("=GESTEP(0;)", ScalarError::Value);
}

#[test]
fn scalar_number_coercion_and_error_precedence_follow_sequence_rules() {
    // Scalar Number parameters use Conversion to Number.  Integer-valued
    // functions apply INT (floor toward negative infinity) before their
    // domain checks; LCM is the deliberate exact-integer exception.
    for (source, expected) in [
        ("=COMBIN(5.9;2.9)", 10.0),
        ("=COMBINA(3.9;2.9)", 6.0),
        ("=FACT(5.9)", 120.0),
        ("=FACTDOUBLE(8.9)", 384.0),
        ("=GCD(48.9;18.1;30.9)", 6.0),
        ("=EVEN(2.1)", 4.0),
        ("=ODD(2.1)", 3.0),
        ("=DELTA(\"1\";TRUE())", 1.0),
        ("=GESTEP(TRUE();\"1\")", 1.0),
        ("=MULTINOMIAL(1.5;1.5)", 6.0),
        ("=MULTINOMIAL(1.4;0.6)", 1.0),
        ("=MULTINOMIAL(1.4;0.6;1)", 2.0),
    ] {
        assert_number(source, expected);
    }

    assert_formula_error("=LCM(4.9;6.1)", ScalarError::Number);
    for source in [
        "=FACT(-0.5)",
        "=FACTDOUBLE(-0.5)",
        "=GCD(-0.5;2)",
        "=LCM(-0.5;2)",
        "=MULTINOMIAL(-0.5;2)",
    ] {
        assert_formula_error(source, ScalarError::Number);
    }

    for (source, expected) in [
        ("=COMBIN(\"bad\";#N/A)", ScalarError::NotAvailable),
        ("=COMBINA(\"bad\";#DIV/0!)", ScalarError::DivisionByZero),
        ("=FACT(\"bad\")", ScalarError::Value),
        ("=FACTDOUBLE(\"bad\")", ScalarError::Value),
        ("=GCD(\"bad\";#N/A)", ScalarError::NotAvailable),
        ("=LCM(\"bad\";#DIV/0!)", ScalarError::DivisionByZero),
        ("=MULTINOMIAL(\"bad\";#N/A)", ScalarError::NotAvailable),
        ("=DELTA(\"bad\";#N/A)", ScalarError::NotAvailable),
        ("=GESTEP(\"bad\";#N/A)", ScalarError::NotAvailable),
        ("=COMBIN(-1;#N/A)", ScalarError::NotAvailable),
        ("=GCD(-1;#DIV/0!)", ScalarError::DivisionByZero),
        ("=MULTINOMIAL(-1;#N/A)", ScalarError::NotAvailable),
    ] {
        assert_formula_error(source, expected);
    }

    // Generated conversion/domain errors retain source order when no formula
    // Error is present: the first generated failure wins.
    for (source, expected) in [
        ("=LCM(0.5;\"bad\")", ScalarError::Number),
        ("=LCM(\"bad\";0.5)", ScalarError::Value),
        ("=GCD(-1;\"bad\")", ScalarError::Number),
        ("=GCD(\"bad\";-1)", ScalarError::Value),
        ("=MULTINOMIAL(-1;\"bad\")", ScalarError::Number),
        ("=MULTINOMIAL(\"bad\";-1)", ScalarError::Value),
    ] {
        assert_formula_error(source, expected);
    }

    // An empty text value is still Text in the scalar profile; it is not the
    // omitted/default argument and cannot be parsed as a Number.
    for source in ["=FACT(\"\")", "=GCD(\"\")", "=LCM(\"\")"] {
        assert_formula_error(source, ScalarError::Value);
    }
}

#[test]
fn discrete_constraints_report_number_errors() {
    for source in [
        "=COMBIN(-1;0)",
        "=COMBIN(5;-1)",
        "=COMBIN(4;5)",
        "=COMBINA(-1;0)",
        "=COMBINA(2;-1)",
        "=FACT(-1)",
        "=FACTDOUBLE(-1)",
        "=GCD(-1;2)",
        "=LCM(-1;2)",
        "=MULTINOMIAL(-1;2)",
    ] {
        assert_formula_error(source, ScalarError::Number);
    }

    assert_number("=GCD(0.5;0)", 0.0);
    assert_number("=COMBINA(0;0)", 1.0);
    assert_number("=COMBINA(1;0)", 1.0);
    assert_formula_error("=COMBINA(0;1)", ScalarError::Number);

    // The standard leaves the all-zero GCD result implementation-defined;
    // this bounded profile chooses the useful zero identity.
    assert_number("=GCD(0;0)", 0.0);
    assert_number("=LCM(0;42)", 0.0);
}

#[test]
fn represented_integer_boundaries_and_finite_overflow_are_stable() {
    // The hexadecimal expectations are independently rounded binary64
    // values of the exact integer results.  These are the boundaries where a
    // factorial/binomial result is still finite versus an evaluator Number
    // error.
    assert_bits("=FACT(170)", 0x7fa4_ab78_6441_8639);
    assert_formula_error("=FACT(171)", ScalarError::Number);
    assert_bits("=FACTDOUBLE(299)", 0x7f95_611d_abe3_7e61);
    assert_bits("=FACTDOUBLE(300)", 0x7fdd_07da_7ecb_62cc);
    assert_formula_error("=FACTDOUBLE(301)", ScalarError::Number);
    assert_formula_error("=FACTDOUBLE(302)", ScalarError::Number);

    // 2^53 is an exactly represented integer input, even though its decimal
    // spelling is beyond the exact-integer range of most f64 values.
    assert_bits("=COMBIN(9007199254740992;1)", 0x4340_0000_0000_0000);
    assert_bits("=COMBIN(9007199254740992;2)", 0x467f_ffff_ffff_ffff);
    assert_bits("=COMBINA(9007199254740992;3)", 0x49b5_5555_5555_5557);
    assert_number("=EVEN(9007199254740992)", 9007199254740992.0);
    assert_formula_error("=ODD(9007199254740992)", ScalarError::Number);
    assert_bits("=GCD(1e308;1e308)", 0x7fe1_ccf3_85eb_c8a0);
    assert_formula_error("=LCM(1e308;3)", ScalarError::Number);
    assert_formula_error("=MULTINOMIAL(1e308;1;1)", ScalarError::Number);

    // A variadic NumberSequence argument must not inherit a small fixed
    // parameter cap from the parser or dispatch table.
    let forty_ones = std::iter::repeat_n("1", 40).collect::<Vec<_>>().join(";");
    for function in ["GCD", "LCM"] {
        let source = format!("={function}({forty_ones})");
        assert_number(&source, 1.0);
    }
    let source = format!("=MULTINOMIAL({forty_ones})");
    assert_bits(&source, 0x49e1_dd5d_0370_98fe);
}

#[test]
fn scalar_discrete_functions_stay_lazy_and_keep_references_typed() {
    for (source, expected) in [
        ("=IF(TRUE();COMBIN(5;2);FACT([Missing.A1]))", 10.0),
        ("=IF(FALSE();FACT([Missing.A1]);FACT(5))", 120.0),
        ("=IFERROR(COMBIN(-1;2);42)", 42.0),
        ("=IFERROR(GCD(-1;2);43)", 43.0),
        ("=IFERROR(DELTA(\"bad\");44)", 44.0),
    ] {
        let expression = parse(source);
        let (_budget, _cancellation, execution) = execution("ods-formula-discrete-lazy");
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate lazily: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }
    assert_unsupported("=FACT([.A1])", UnsupportedKind::Reference);
    assert_unsupported("=GCD([.A1:.A2])", UnsupportedKind::Reference);
}

#[test]
fn scalar_budget_failures_are_typed_uncatchable_and_refund_storage() {
    let expression = parse("=MULTINOMIAL(1;1;1;1;1;1;1;1;1;1)");
    let (budget, _cancellation, execution) = execution("ods-formula-discrete-work");
    let error = evaluate(
        &expression,
        &execution,
        &EvaluationLimits::default().with_max_steps(0),
    )
    .expect_err("zero work must refuse before publishing a discrete result");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "unexpected work refusal: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    // The same typed refusal must cross IFERROR unchanged.  Formula-level
    // #NUM! and #VALUE! are catchable; evaluator budget failures are not.
    let expression = parse("=IFERROR(MULTINOMIAL(1;1;1;1;1;1;1;1;1;1);7)");
    let error = evaluate(
        &expression,
        &execution,
        &EvaluationLimits::default().with_max_steps(0),
    )
    .expect_err("IFERROR must not catch a scalar Work refusal");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "IFERROR changed the typed outcome: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}
