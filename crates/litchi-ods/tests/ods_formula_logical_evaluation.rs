//! Independent integration coverage for the scalar logical-function profile.
//!
//! The expected values in this file come from OpenDocument v1.4 Part 4
//! sections 6.1, 6.3, and 6.15. Formula errors remain values; evaluator
//! capability, cancellation, and resource failures remain Rust results.
//! Conditional functions are tested for their specified one-branch behavior,
//! while aggregate Boolean functions are tested for eager child evaluation.

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
    execution_with_core_limits(scope, CoreLimits::for_profile(Profile::Server))
}

fn execution_with_memory(
    scope: &str,
    memory_bytes: usize,
) -> (Budget, CancellationSource, ExecutionContext) {
    execution_with_core_limits(
        scope,
        CoreLimits::new(
            u64::try_from(memory_bytes).expect("test memory limit fits u64"),
            1 << 30,
            1 << 30,
            10_000_000,
            256,
            1_000_000_000,
        ),
    )
}

fn execution_with_core_limits(
    scope: &str,
    core_limits: CoreLimits,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), core_limits);
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
    let evaluation_context = EvaluationContext::new(context);
    evaluate_scalar(expression, &evaluation_context, limits)
}

fn number(result: &EvaluatedScalar<'_>) -> f64 {
    match result.value() {
        ScalarValue::Number(value) => *value,
        value => panic!("expected Number, got {value:?}"),
    }
}

fn logical(result: &EvaluatedScalar<'_>) -> bool {
    match result.value() {
        ScalarValue::Logical(value) => *value,
        value => panic!("expected Logical, got {value:?}"),
    }
}

fn text<'r>(result: &'r EvaluatedScalar<'_>) -> &'r str {
    match result.value() {
        ScalarValue::Text(value) => value.as_ref(),
        value => panic!("expected Text, got {value:?}"),
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
    let (_budget, _cancellation, context) = execution("ods-formula-logical-unsupported");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("capability-dependent logical expression should be refused");
    assert!(
        matches!(&error, EvaluationFailure::Unsupported(kind) if *kind == expected),
        "{source:?} returned the wrong capability error: {error}"
    );
}

fn long_addition_body(terms: usize) -> String {
    assert!(terms > 0);
    let mut source = String::from("1");
    for _ in 1..terms {
        source.push_str("+1");
    }
    source
}

fn escaped_literal(decoded_bytes: usize) -> String {
    assert!(decoded_bytes >= 1);
    let mut literal = String::from("\"");
    literal.push_str(&"x".repeat(decoded_bytes - 1));
    literal.push_str("\"\"");
    literal.push('"');
    literal
}

#[test]
fn all_standard_logical_functions_return_their_scalar_profile_values() {
    let (_budget, _cancellation, context) = execution("ods-formula-logical-values");
    let cases = [
        ("=TRUE()", true),
        ("=FALSE()", false),
        ("=NOT(TRUE())", false),
        ("=NOT(FALSE())", true),
        ("=AND(TRUE();TRUE())", true),
        ("=AND(TRUE();FALSE())", false),
        ("=OR(FALSE();FALSE())", false),
        ("=OR(FALSE();TRUE())", true),
        ("=XOR(FALSE();TRUE())", true),
        ("=XOR(TRUE();TRUE())", false),
        ("=XOR(1;2;3;4)", false),
        ("=IF(TRUE();1;2)", true),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        if source == "=IF(TRUE();1;2)" {
            assert_eq!(number(&result), 1.0, "{source:?}");
        } else {
            assert_eq!(logical(&result), expected, "{source:?}");
        }
    }

    for (source, expected) in [("=IFERROR(#N/A;2)", 2.0), ("=IFNA(#N/A;3)", 3.0)] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }
}

#[test]
fn logical_functions_apply_number_coercion_and_and_or_numeric_text_profile() {
    let (_budget, _cancellation, context) = execution("ods-formula-logical-coercion");
    let logical_cases = [
        ("=NOT(0)", true),
        ("=NOT(-2)", false),
        ("=AND(1;2)", true),
        ("=AND(1;0)", false),
        ("=OR(0;0)", false),
        ("=OR(0;2)", true),
        ("=XOR(1;0;1)", false),
    ];
    for (source, expected) in logical_cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(logical(&result), expected, "{source:?}");
    }

    let number_cases = [("=IF(0;1;2)", 2.0), ("=IF(-1;1;2)", 1.0)];
    for (source, expected) in number_cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    // This profile uses locale-independent decimal text-to-number conversion
    // for NumberSequenceList, which is the parameter type of AND and OR.
    for (source, expected) in [("=AND(\"1\")", true), ("=OR(\"0\")", false)] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(logical(&result), expected, "{source:?}");
    }

    // The profile deliberately does not guess a locale-specific textual
    // Boolean spelling for the Logical parameter family.
    for source in ["=IF(\"1\";1;2)", "=NOT(\"TRUE\")", "=XOR(\"1\")"] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula Error: {error}"));
        assert_eq!(formula_error(&result), ScalarError::Value, "{source:?}");
    }
}

#[test]
fn conditional_functions_honor_all_optional_argument_forms() {
    let (_budget, _cancellation, context) = execution("ods-formula-if-forms");
    let number_cases = [
        ("=IF(TRUE();)", 0.0),
        ("=IF(TRUE();1)", 1.0),
        ("=IF(TRUE();;)", 0.0),
        ("=IF(FALSE();;)", 0.0),
        ("=IF(TRUE();;2)", 0.0),
        ("=IF(FALSE();;2)", 2.0),
        ("=IF(TRUE();1;)", 1.0),
        ("=IF(FALSE();1;)", 0.0),
        ("=IF(TRUE();1;2)", 1.0),
        ("=IF(FALSE();1;2)", 2.0),
    ];
    for (source, expected) in number_cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    let logical_cases = [
        ("=IF(TRUE())", true),
        ("=IF(FALSE())", false),
        ("=IF(FALSE();1)", false),
        ("=IF(FALSE();)", false),
    ];
    for (source, expected) in logical_cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(logical(&result), expected, "{source:?}");
    }
}

#[test]
fn required_arity_errors_are_formula_values() {
    let cases = [
        ("=TRUE(1)", ScalarError::Value),
        ("=FALSE(1)", ScalarError::Value),
        ("=NOT()", ScalarError::Value),
        ("=XOR()", ScalarError::Value),
        ("=NOT(1;2)", ScalarError::Value),
        ("=IF()", ScalarError::Value),
        ("=IFERROR(1)", ScalarError::Value),
        ("=IFNA(1;2;3)", ScalarError::Value),
        ("=AND(;TRUE())", ScalarError::Value),
        ("=OR(FALSE();)", ScalarError::Value),
        ("=XOR(;TRUE())", ScalarError::Value),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let (budget, _cancellation, context) = execution("ods-formula-logical-arity");
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
        assert_eq!(formula_error(&result), expected, "{source:?}");
        assert_eq!(
            budget.used(Resource::Memory),
            0,
            "{source:?} leaked storage"
        );
    }

    // The standard permits either a Logical identity or an Error for zero
    // arguments to AND and OR. Whichever choice the profile makes remains a
    // formula result rather than a capability/refusal failure.
    for source in ["=AND()", "=OR()"] {
        let expression = parse(source);
        let (_budget, _cancellation, context) = execution("ods-formula-logical-zero");
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a scalar result: {error}"));
        assert!(
            matches!(
                result.value(),
                ScalarValue::Logical(_) | ScalarValue::Error(_)
            ),
            "{source:?} returned a non-logical formula value: {:?}",
            result.value()
        );
    }
}

#[test]
fn eager_wrong_arity_still_evaluates_present_children() {
    let no_arguments = parse("=NOT()");
    let (empty_budget, _cancellation, empty_context) = execution("ods-formula-arity-empty");
    let empty_result = evaluate(&no_arguments, &empty_context, &EvaluationLimits::default())
        .expect("NOT() should return a formula arity error");
    assert_eq!(formula_error(&empty_result), ScalarError::Value);
    let empty_work = empty_budget.used(Resource::Work);

    let with_children = parse("=NOT(1+2;3+4)");
    let (child_budget, _cancellation, child_context) = execution("ods-formula-arity-children");
    let child_result = evaluate(&with_children, &child_context, &EvaluationLimits::default())
        .expect("wrong-arity NOT should return a formula arity error");
    assert_eq!(formula_error(&child_result), ScalarError::Value);
    assert!(
        child_budget.used(Resource::Work) > empty_work,
        "present arguments must be visited before a wrong-arity result"
    );
}

#[test]
fn and_or_evaluate_all_children_and_propagate_source_errors() {
    let (_budget, _cancellation, context) = execution("ods-formula-logical-errors");
    for (source, expected) in [
        ("=AND(#N/A;TRUE())", ScalarError::NotAvailable),
        ("=OR(#N/A;FALSE())", ScalarError::NotAvailable),
        ("=XOR(#DIV/0!;TRUE())", ScalarError::DivisionByZero),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should retain formula Error: {error}"));
        assert_eq!(formula_error(&result), expected, "{source:?}");
    }

    // A decisive Boolean value does not short-circuit aggregate functions.
    // The reference is a capability refusal, so this distinguishes eager
    // child scheduling from merely checking the final Boolean result.
    assert_unsupported("=AND(FALSE();[.A1])", UnsupportedKind::Reference);
    assert_unsupported("=OR(TRUE();[.A1])", UnsupportedKind::Reference);
}

#[test]
fn if_and_error_handlers_visit_only_the_selected_or_needed_value() {
    let (_budget, _cancellation, context) = execution("ods-formula-logical-lazy");
    let selected = [
        ("=IF(TRUE();1;#DIV/0!)", 1.0),
        ("=IF(FALSE();#N/A;2)", 2.0),
        ("=IFERROR(1;#DIV/0!)", 1.0),
        ("=IFNA(1;#DIV/0!)", 1.0),
    ];
    for (source, expected) in selected {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should skip its unused branch: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    for source in [
        "=IF(TRUE();1;[.A1])",
        "=IF(FALSE();[.A1];2)",
        "=IFERROR(1;[.A1])",
        "=IFNA(1;[.A1])",
    ] {
        let expression = parse(source);
        let result =
            evaluate(&expression, &context, &EvaluationLimits::default()).unwrap_or_else(|error| {
                panic!("{source:?} should skip its unused capability branch: {error}")
            });
        assert_eq!(
            number(&result),
            if source.contains("FALSE") { 2.0 } else { 1.0 }
        );
    }
}

#[test]
fn iferror_and_ifna_handle_formula_errors_but_not_evaluator_failures() {
    let (_budget, _cancellation, context) = execution("ods-formula-logical-error-handlers");
    for (source, expected) in [
        ("=IFERROR(#N/A;2)", 2.0),
        ("=IFERROR(1/0;2)", 2.0),
        ("=IFNA(#N/A;3)", 3.0),
        // Missing required slots produce a Value error when evaluated. A
        // handler can catch that formula error, and an unused slot stays lazy.
        ("=IFERROR(;2)", 2.0),
        ("=IFERROR(1;)", 1.0),
        ("=IFNA(1;)", 1.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should catch formula Error: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    let expression = parse("=IFNA(#DIV/0!;3)");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("IFNA should return the non-NA formula Error");
    assert_eq!(formula_error(&result), ScalarError::DivisionByZero);

    for source in ["=IFNA(;2)", "=IFERROR(#N/A;)", "=IFNA(#N/A;)"] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .expect("an evaluated missing slot should remain a formula error");
        assert_eq!(formula_error(&result), ScalarError::Value, "{source:?}");
    }

    assert_unsupported("=IFERROR([.A1];2)", UnsupportedKind::Reference);
    assert_unsupported("=IFNA([.A1];2)", UnsupportedKind::Reference);
    assert_unsupported("=IFERROR(#N/A;[.A1])", UnsupportedKind::Reference);
    assert_unsupported("=IFNA(#N/A;[.A1])", UnsupportedKind::Reference);

    let expression = parse("=IFERROR(1+2;3)");
    let (budget, cancellation, context) = execution("ods-formula-logical-handler-cancel");
    cancellation.cancel();
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("cancellation must not be converted into Alternative");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Memory), 0);

    let source = format!("=IFERROR({};2)", long_addition_body(512));
    let expression = parse(&source);
    let (_budget, _cancellation, context) = execution("ods-formula-logical-handler-limit");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(32),
    )
    .expect_err("a Work refusal must not be converted into Alternative");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong handler failure: {error}"
    );
}

#[test]
fn error_handlers_compute_a_successful_input_once() {
    let input_body = long_addition_body(256);
    let input_source = format!("={input_body}");
    let input_expression = parse(&input_source);
    let (direct_budget, _cancellation, direct_context) =
        execution("ods-formula-handler-direct-input");
    let direct = evaluate(
        &input_expression,
        &direct_context,
        &EvaluationLimits::default(),
    )
    .expect("the handler input should evaluate directly");
    assert_eq!(number(&direct), 256.0);
    let direct_work = direct_budget.used(Resource::Work);
    assert!(direct_work > 0);

    for name in ["IFERROR", "IFNA"] {
        let source = format!("={name}({input_body};2)");
        let expression = parse(&source);
        let (budget, _cancellation, context) = execution("ods-formula-handler-single-evaluation");
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), 256.0, "{source:?}");
        assert!(
            budget.used(Resource::Work) < direct_work.saturating_mul(2),
            "{source:?} appears to evaluate X more than once: direct={direct_work}, handler={}",
            budget.used(Resource::Work)
        );
    }
}

#[test]
fn unused_conditional_work_and_storage_are_not_admitted() {
    let body = long_addition_body(512);
    for source in [
        format!("=IF(TRUE();1;{})", body),
        format!("=IF(FALSE();{};2)", body),
    ] {
        let expression = parse(&source);
        let (_budget, _cancellation, context) = execution("ods-formula-logical-unused-work");
        let result = evaluate(
            &expression,
            &context,
            &EvaluationLimits::default().with_max_steps(64),
        )
        .unwrap_or_else(|error| panic!("{source:?} should skip its Work-heavy branch: {error}"));
        assert_eq!(
            number(&result),
            if source.contains("TRUE") { 1.0 } else { 2.0 }
        );
    }

    let large = escaped_literal(4_096);
    let source = format!("=IF(TRUE();\"plain\";{})", large);
    let expression = parse(&source);
    let (budget, _cancellation, context) =
        execution_with_memory("ods-formula-logical-unused-memory", 256);
    let result = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default()
            .with_max_storage_bytes(256)
            .with_max_text_bytes(4_096),
    )
    .expect("the unselected owned text must not consume storage");
    assert_eq!(text(&result), "plain");
    assert_eq!(result.reserved_output_bytes(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn selected_owned_text_keeps_and_releases_its_result_reservation() {
    let expression = parse("=IF(TRUE();\"a\"\"b\";\"unused\")");
    let (budget, _cancellation, context) = execution("ods-formula-logical-result");
    let result = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_storage_bytes(256),
    )
    .expect("selected escaped text should evaluate");
    assert_eq!(text(&result), "a\"b");
    assert!(result.reserved_output_bytes() >= 3);
    assert!(budget.used(Resource::Memory) > 0);
    drop(context);
    assert!(budget.used(Resource::Memory) > 0);
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn eager_boolean_children_charge_work_and_limits() {
    let body = long_addition_body(512);
    for source in [
        format!("=AND(FALSE();{})", body),
        format!("=OR(TRUE();{})", body),
    ] {
        let expression = parse(&source);
        let (budget, _cancellation, context) = execution("ods-formula-logical-eager-work");
        let error = evaluate(
            &expression,
            &context,
            &EvaluationLimits::default().with_max_steps(64),
        )
        .expect_err("eager AND/OR must visit a Work-heavy child");
        assert!(
            matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
            "{source:?} returned the wrong eager failure: {error}"
        );
        assert!(
            budget.used(Resource::Work) > 0,
            "{source:?} charged no work"
        );
        assert_eq!(
            budget.used(Resource::Memory),
            0,
            "{source:?} leaked storage"
        );
    }
}

#[test]
fn cancellation_is_checked_before_logical_work_and_is_not_caught() {
    let expression = parse("=IFERROR(#N/A;2)");
    let (budget, cancellation, context) = execution("ods-formula-logical-cancel");
    cancellation.cancel();
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("cancelled logical evaluation should refuse before handler logic");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Work), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn nested_conditionals_evaluate_iteratively_with_bounded_stack() {
    const DEPTH: usize = 128;
    let mut source = String::from("=");
    for _ in 0..DEPTH {
        source.push_str("IF(TRUE();");
    }
    source.push('1');
    for _ in 0..DEPTH {
        source.push_str(";0)");
    }
    let expression = parse(&source);
    let (_budget, _cancellation, context) = execution("ods-formula-logical-deep");
    let result = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default()
            .with_max_steps(100_000)
            .with_max_stack_entries(2_048),
    )
    .expect("deep selected branches should not recurse in the evaluator");
    assert_eq!(number(&result), 1.0);
}

#[test]
fn logical_source_spelling_and_function_case_are_preserved() {
    let source = "of:== if(TRUE();1;2) ";
    let expression = parse(source);
    assert_eq!(expression.source(), source);
    assert!(expression.is_force_recalculate());
    let (_budget, _cancellation, context) = execution("ods-formula-logical-source");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("function names are case-insensitive");
    assert_eq!(number(&result), 1.0);
    assert_eq!(expression.source(), source);
    assert!(expression.is_force_recalculate());
}
