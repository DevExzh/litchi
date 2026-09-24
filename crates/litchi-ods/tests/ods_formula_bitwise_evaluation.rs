//! Independent integration coverage for the OpenFormula 1.4 bit-operation
//! functions in section 6.6.
//!
//! These tests exercise the scalar profile through its public evaluator.  The
//! expected bit patterns are computed from an independent `u64` expression and
//! are kept inside the 48-bit unsigned domain required by the specification.
//! Formula errors remain scalar values; capability, resource, allocation, and
//! cancellation failures remain typed evaluator results.

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

const MAX_U48: u64 = (1 << 48) - 1;

fn execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1024 * 1024).expect("one in-flight MiB"),
        0,
    )
    .expect("valid execution limits");
    (
        budget.clone(),
        cancellation,
        ExecutionContext::new(budget, token, execution_limits),
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

fn assert_unsupported(source: &str, expected: UnsupportedKind) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-bitwise-unsupported");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("capability-dependent bitwise expression should be refused");
    assert!(
        matches!(&error, EvaluationFailure::Unsupported(kind) if *kind == expected),
        "{source:?} returned the wrong capability error: {error}"
    );
}

#[test]
fn bitwise_boolean_operations_match_independent_u48_truth_tables() {
    let pairs = [
        (0_u64, 0_u64),
        (1, 0),
        (0, 1),
        (0x55aa, 0x0f0f),
        (0x1234_5678_9abc, 0x0fed_cba9_8765),
        (1 << 47, MAX_U48),
        (MAX_U48, MAX_U48 - 1),
        (MAX_U48, MAX_U48),
    ];
    let (_budget, _cancellation, context) = execution("ods-formula-bitwise-truth-table");

    for name in ["BITAND", "BITOR", "BITXOR"] {
        for (left, right) in pairs {
            let expected = match name {
                "BITAND" => left & right,
                "BITOR" => left | right,
                "BITXOR" => left ^ right,
                _ => unreachable!("case table only contains bitwise binary functions"),
            };
            let source = format!("={name}({left};{right})");
            let expression = parse(&source);
            let result = evaluate(&expression, &context, &EvaluationLimits::default())
                .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
            assert_eq!(number(&result), expected as f64, "{source:?}");
            assert!(number(&result).is_finite());
            assert!(number(&result) >= 0.0 && number(&result) <= MAX_U48 as f64);
        }
    }
}

#[test]
fn shifts_follow_odf_direction_zero_and_48_bit_boundaries() {
    let cases = [
        ("=BITLSHIFT(5;3)", 40.0),
        ("=BITRSHIFT(40;3)", 5.0),
        ("=BITLSHIFT(40;-3)", 5.0),
        ("=BITRSHIFT(5;-3)", 40.0),
        ("=BITLSHIFT(9;0)", 9.0),
        ("=BITLSHIFT(9;-0.9)", 9.0),
        ("=BITLSHIFT(9;-1e100)", 0.0),
        ("=BITRSHIFT(0;-1e100)", 0.0),
        ("=BITRSHIFT(9;0)", 9.0),
        ("=BITLSHIFT(3.9;2.9)", 12.0),
        ("=BITRSHIFT(40.9;2.9)", 10.0),
        ("=BITLSHIFT(1;47)", (1_u64 << 47) as f64),
        ("=BITRSHIFT(281474976710655;47)", 1.0),
        ("=BITRSHIFT(281474976710655;48)", 0.0),
        ("=BITLSHIFT(0;1e100)", 0.0),
        ("=BITRSHIFT(281474976710655;1e100)", 0.0),
    ];
    let (_budget, _cancellation, context) = execution("ods-formula-bitwise-shifts");
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    // The selected scalar profile reports a Number error when a left shift
    // would produce a result outside its required unsigned 48-bit domain.
    for source in [
        "=BITLSHIFT(1;48)",
        "=BITLSHIFT(1;1e100)",
        "=BITLSHIFT(281474976710655;47)",
        "=BITLSHIFT(140737488355328;17)",
        "=BITRSHIFT(1;-48)",
        "=BITRSHIFT(1;-1e100)",
        "=BITRSHIFT(140737488355328;-17)",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
        assert_eq!(formula_error(&result), ScalarError::Number, "{source:?}");
    }
}

#[test]
fn integer_parameters_use_profile_truncation_text_and_logical_conversion() {
    let cases = [
        ("=BITAND(7.9;3.1)", 3.0),
        ("=BITOR(4.9;2.1)", 6.0),
        ("=BITXOR(6.9;3.1)", 5.0),
        ("=BITXOR(\"6.9\";\"3.1\")", 5.0),
        ("=BITLSHIFT(\"3.9\";\"2.9\")", 12.0),
        ("=BITAND(TRUE();1)", 1.0),
        ("=BITOR(FALSE();1)", 1.0),
        // Truncation happens before the non-negative constraint: -0.9
        // truncates to signed zero and is therefore accepted as zero.
        ("=BITAND(-0.9;1)", 0.0),
        ("=BITAND(281474976710655.5;281474976710655)", MAX_U48 as f64),
        (
            "=BITAND(281474976710654.9;281474976710655)",
            (MAX_U48 - 1) as f64,
        ),
    ];
    let (_budget, _cancellation, context) = execution("ods-formula-bitwise-conversion");
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    for (source, expected) in [
        ("=BITAND(-1.9;1)", ScalarError::Number),
        ("=BITOR(\"-1\";1)", ScalarError::Number),
        ("=BITXOR(\"not-a-number\";1)", ScalarError::Value),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
        assert_eq!(formula_error(&result), expected, "{source:?}");
    }
}

#[test]
fn bitwise_constraints_and_results_stay_within_the_unsigned_u48_profile() {
    let (_budget, _cancellation, context) = execution("ods-formula-bitwise-u48");
    for (source, expected) in [
        ("=BITAND(281474976710655;281474976710655)", MAX_U48 as f64),
        ("=BITOR(140737488355328;140737488355327)", MAX_U48 as f64),
        ("=BITXOR(281474976710655;0)", MAX_U48 as f64),
        ("=BITLSHIFT(140737488355328;0)", (1_u64 << 47) as f64),
        ("=BITRSHIFT(281474976710655;0)", MAX_U48 as f64),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        let value = number(&result);
        assert_eq!(value, expected, "{source:?}");
        assert!((0.0..=MAX_U48 as f64).contains(&value));
    }

    for source in [
        "=BITAND(-1;0)",
        "=BITOR(-1;0)",
        "=BITXOR(-1;0)",
        "=BITLSHIFT(-1;0)",
        "=BITRSHIFT(-1;0)",
        "=BITAND(281474976710656;0)",
        "=BITOR(281474976710656;0)",
        "=BITXOR(281474976710656;0)",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
        assert_eq!(formula_error(&result), ScalarError::Number, "{source:?}");
    }
}

#[test]
fn formula_errors_propagate_in_source_order_as_values() {
    let (_budget, _cancellation, context) = execution("ods-formula-bitwise-errors");
    for (source, expected) in [
        ("=BITAND(#N/A;#DIV/0!)", ScalarError::NotAvailable),
        ("=BITOR(#DIV/0!;#N/A)", ScalarError::DivisionByZero),
        ("=BITXOR(#REF!;#VALUE!)", ScalarError::Reference),
        ("=BITLSHIFT(#NUM!;1)", ScalarError::Number),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should produce a formula value: {error}"));
        assert_eq!(formula_error(&result), expected, "{source:?}");
    }
}

#[test]
fn bitwise_calls_visit_capability_failures_even_with_wrong_arity() {
    for (source, expected) in [
        ("=BITAND([.A1])", UnsupportedKind::Reference),
        ("=BITOR(1;[.A1];2)", UnsupportedKind::Reference),
        ("=BITLSHIFT([.A1];0;1)", UnsupportedKind::Reference),
        ("=BITXOR({1;2};1)", UnsupportedKind::Array),
    ] {
        assert_unsupported(source, expected);
    }
}

#[test]
fn conditional_handlers_skip_unselected_bitwise_work_but_do_not_catch_refusals() {
    let (_budget, _cancellation, context) = execution("ods-formula-bitwise-lazy");
    for (source, expected) in [
        ("=IF(TRUE();BITAND(1;2);[.A1])", 0.0),
        ("=IF(FALSE();[.A1];BITOR(1;2))", 3.0),
        ("=IF(TRUE();BITAND(1;1);BITLSHIFT(1;1e100))", 1.0),
        ("=IF(FALSE();BITLSHIFT(1;1e100);7)", 7.0),
        ("=IFERROR(BITAND(#N/A;1);7)", 7.0),
        ("=IFNA(BITOR(#N/A;1);9)", 9.0),
        ("=IFERROR(BITLSHIFT(1;48);9)", 9.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    for source in ["=IFERROR(BITAND([.A1];1);7)", "=IFNA(BITXOR([.A1];1);9)"] {
        let expression = parse(source);
        let error = evaluate(&expression, &context, &EvaluationLimits::default())
            .expect_err("a capability refusal is not a formula error handler input");
        assert!(
            matches!(
                error,
                EvaluationFailure::Unsupported(UnsupportedKind::Reference)
            ),
            "{source:?} returned the wrong handler result: {error}"
        );
    }

    let expression = parse("=IFNA(BITAND(#DIV/0!;1);9)");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("IFNA should leave a non-NA formula error as a value");
    assert_eq!(formula_error(&result), ScalarError::DivisionByZero);
}

#[test]
fn bitwise_limits_cancellation_and_result_lifetime_are_typed_and_atomic() {
    let expression = parse("=BITAND(1;2)");

    let (budget, cancellation, context) = execution("ods-formula-bitwise-cancel");
    cancellation.cancel();
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("cancelled bitwise evaluation should refuse before work");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Work), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = execution("ods-formula-bitwise-work-limit");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(0),
    )
    .expect_err("zero Work allowance should refuse before the first evaluator step");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work)
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = execution("ods-formula-bitwise-memory-limit");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero evaluator storage allowance should refuse atomically");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory)
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = execution("ods-formula-bitwise-result-lifetime");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("ordinary bitwise result should evaluate");
    assert_eq!(number(&result), 0.0);
    assert_eq!(result.reserved_output_bytes(), 0);
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}

fn nested_bitxor(depth: usize) -> String {
    assert!(depth > 0);
    let mut body = String::from("0");
    for _ in 0..depth {
        body = format!("BITXOR({body};1)");
    }
    format!("={body}")
}

#[test]
fn nested_bitwise_calls_use_the_explicit_stack_under_bounded_work() {
    const DEPTH: usize = 128;
    let source = nested_bitxor(DEPTH);
    let expression = parse(&source);
    let (_budget, _cancellation, context) = execution("ods-formula-bitwise-nested");
    let result = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default()
            .with_max_steps(100_000)
            .with_max_stack_entries(2_048),
    )
    .unwrap_or_else(|error| panic!("nested bitwise calls should evaluate: {error}"));
    assert_eq!(number(&result), if DEPTH % 2 == 0 { 0.0 } else { 1.0 });

    let (_budget, _cancellation, context) = execution("ods-formula-bitwise-nested-limit");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(16),
    )
    .expect_err("nested bitwise work should honor the local step limit");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work)
    );
}
