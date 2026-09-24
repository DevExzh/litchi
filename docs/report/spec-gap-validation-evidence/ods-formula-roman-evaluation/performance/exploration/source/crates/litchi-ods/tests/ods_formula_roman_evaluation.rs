//! Independent integration coverage for the OpenFormula 1.4 ROMAN and ARABIC
//! conversion functions in section 6.19.
//!
//! The classic representation and the Arabic decoder below are deliberately
//! independent of the evaluator implementation.  Format 4 is checked with a
//! breadth-first oracle over the Roman sign rule (a symbol is negative when a
//! larger symbol occurs to its right), rather than by reproducing the
//! evaluator's construction algorithm.

use std::{
    collections::HashSet,
    num::{NonZeroU64, NonZeroUsize},
};

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

fn assert_formula_error(source: &str, expected: ScalarError) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-roman-error");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
    assert_eq!(formula_error(&result), expected, "{source:?}");
}

fn assert_unsupported(source: &str, expected: UnsupportedKind) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-roman-unsupported");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("capability-dependent Roman expression should be refused");
    assert!(
        matches!(&error, EvaluationFailure::Unsupported(kind) if *kind == expected),
        "{source:?} returned the wrong capability error: {error}"
    );
}

fn classical_roman(mut value: usize) -> String {
    const DIGITS: &[(usize, &str)] = &[
        (1_000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut result = String::new();
    for &(digit, spelling) in DIGITS {
        while value >= digit {
            result.push_str(spelling);
            value -= digit;
        }
    }
    result
}

fn arabic_reference(value: &str) -> Option<i64> {
    fn roman_value(character: char) -> Option<i64> {
        Some(match character.to_ascii_uppercase() {
            'I' => 1,
            'V' => 5,
            'X' => 10,
            'L' => 50,
            'C' => 100,
            'D' => 500,
            'M' => 1_000,
            _ => return None,
        })
    }

    let values = value.chars().map(roman_value).collect::<Option<Vec<_>>>()?;
    let mut total = 0_i64;
    for (index, &current) in values.iter().enumerate() {
        let subtract = values[index + 1..].iter().any(|&right| right > current);
        total += if subtract { -current } else { current };
    }
    Some(total)
}

/// Find the shortest string satisfying the direct/indirect Roman sign rule.
///
/// The search runs right-to-left.  Its state is the accumulated value and the
/// largest symbol already placed to the right, so every string accepted by the
/// Arabic semantics is represented.  Fifteen symbols cover the classic
/// spelling of every value below 4000; states are intentionally allowed to
/// leave the requested output interval so overshoot-and-subtract spellings are
/// not accidentally omitted from the oracle.
fn shortest_roman_lengths(max_value: usize) -> Vec<usize> {
    const SYMBOLS: [i32; 7] = [1, 5, 10, 50, 100, 500, 1_000];
    const MAX_DEPTH: usize = 15;
    let max_value = i32::try_from(max_value).expect("test bound fits i32");
    let mut shortest = vec![usize::MAX; usize::try_from(max_value).unwrap() + 1];
    shortest[0] = 0;
    let mut frontier = HashSet::from([(0_i32, 0_i32)]);

    for depth in 1..=MAX_DEPTH {
        let mut next = HashSet::new();
        for &(value, rightmost) in &frontier {
            for symbol in SYMBOLS {
                let signed = if symbol < rightmost { -symbol } else { symbol };
                let value = value + signed;
                let rightmost = rightmost.max(symbol);
                if (0..=max_value).contains(&value) {
                    let index = usize::try_from(value).expect("non-negative value");
                    if shortest[index] == usize::MAX {
                        shortest[index] = depth;
                    }
                }
                next.insert((value, rightmost));
            }
        }
        frontier = next;
        if shortest.iter().all(|length| *length != usize::MAX) {
            break;
        }
    }
    assert!(
        shortest.iter().all(|length| *length != usize::MAX),
        "the independent Roman search did not cover all values through {max_value}"
    );
    shortest
}

#[test]
fn classic_roman_matches_an_independent_digit_oracle() {
    let (_budget, _cancellation, context) = execution("ods-formula-roman-classic");
    for value in 0..4_000 {
        let source = format!("=ROMAN({value})");
        let expression = parse(&source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), classical_roman(value), "{source:?}");
    }
}

#[test]
fn simplified_roman_matches_independent_shortest_lengths_and_arabic() {
    let shortest = shortest_roman_lengths(3_999);
    let (_budget, _cancellation, context) = execution("ods-formula-roman-simplified");
    for (value, &expected_length) in shortest.iter().enumerate() {
        let source = format!("=ROMAN({value};4)");
        let expression = parse(&source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        let spelling = text(&result);
        assert_eq!(arabic_reference(spelling), Some(value as i64), "{source:?}");
        assert_eq!(
            spelling.len(),
            expected_length,
            "{source:?} produced {spelling:?}"
        );
        assert!(
            spelling
                .chars()
                .all(|character| "IVXLCDM".contains(character)),
            "{source:?} emitted a non-Roman character: {spelling:?}"
        );
    }
    assert_eq!(shortest[8], 3);
    assert_eq!(shortest[998], 3);
    assert_eq!(shortest[3_999], 5);
}

#[test]
fn arabic_accepts_ascii_case_and_indirect_subtraction() {
    let (_budget, _cancellation, context) = execution("ods-formula-roman-arabic");
    for (source, expected) in [
        ("=ARABIC(\"\")", 0.0),
        ("=ARABIC(\"I\")", 1.0),
        ("=ARABIC(\"iv\")", 4.0),
        ("=ARABIC(\"mCmXcIx\")", 1_999.0),
        ("=ARABIC(\"IIX\")", 8.0),
        ("=ARABIC(\"IXX\")", 19.0),
        ("=ARABIC(\"XLIX\")", 49.0),
        ("=ARABIC(\"IC\")", 99.0),
        ("=ARABIC(\"MIM\")", 1_999.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    for source in [
        "=ARABIC(\"A\")",
        "=ARABIC(\"Ⅰ\")",
        "=ARABIC(\"IV \")",
        "=ARABIC(\"-I\")",
        "=ARABIC(\"0\")",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
        assert_eq!(formula_error(&result), ScalarError::Value, "{source:?}");
    }
}

#[test]
fn roman_formats_follow_specified_subtractive_levels_and_logical_compatibility() {
    let (_budget, _cancellation, context) = execution("ods-formula-roman-formats");
    for (source, expected) in [
        ("=ROMAN(499;0)", "CDXCIX"),
        ("=ROMAN(499;1)", "LDVLIV"),
        ("=ROMAN(499;2)", "ID"),
        ("=ROMAN(499;3)", "ID"),
        ("=ROMAN(499;4)", "ID"),
        ("=ROMAN(45;1)", "VL"),
        ("=ROMAN(45;2)", "XLV"),
        ("=ROMAN(9)", "IX"),
        ("=ROMAN(9;TRUE())", "IX"),
        ("=ROMAN(499;FALSE())", "ID"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }
}

#[test]
fn roman_and_arabic_round_trip_every_value_in_each_format() {
    for format in 0..=4 {
        let (_budget, _cancellation, context) = execution("ods-formula-roman-round-trip");
        for value in 0..4_000 {
            let source = format!("=ARABIC(ROMAN({value};{format}))");
            let expression = parse(&source);
            let result = evaluate(&expression, &context, &EvaluationLimits::default())
                .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
            assert_eq!(number(&result), value as f64, "{source:?}");
        }
    }
}

#[test]
fn roman_integer_bounds_fractional_conversion_and_arity_are_formula_values() {
    let (_budget, _cancellation, context) = execution("ods-formula-roman-bounds");
    for (source, expected) in [
        ("=ROMAN(4.9)", "IV"),
        ("=ROMAN(3999.9)", "MMMCMXCIX"),
        ("=ROMAN(-0.9)", ""),
        ("=ROMAN(499;1.9)", "LDVLIV"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    for (source, expected) in [
        ("=ROMAN(-1)", ScalarError::Number),
        ("=ROMAN(4000)", ScalarError::Number),
        ("=ROMAN(1;-1)", ScalarError::Number),
        ("=ROMAN(1;5)", ScalarError::Number),
        ("=ROMAN()", ScalarError::Value),
        ("=ROMAN(1;0;2)", ScalarError::Value),
        ("=ARABIC()", ScalarError::Value),
        ("=ARABIC(\"I\";\"V\")", ScalarError::Value),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
        assert_eq!(formula_error(&result), expected, "{source:?}");
    }
}

#[test]
fn roman_eagerly_propagates_children_and_preserves_leftmost_errors() {
    for (source, expected) in [
        ("=ROMAN(#N/A;#DIV/0!)", ScalarError::NotAvailable),
        ("=ARABIC(#N/A)", ScalarError::NotAvailable),
        ("=ROMAN(#REF!;#DIV/0!)", ScalarError::Reference),
        ("=ARABIC(#NUM!;#VALUE!)", ScalarError::Number),
    ] {
        assert_formula_error(source, expected);
    }

    for source in ["=ROMAN(1;[.A1];3)", "=ARABIC([.A1];1)", "=ROMAN(1;[.A1])"] {
        assert_unsupported(source, UnsupportedKind::Reference);
    }
}

#[test]
fn conditional_handlers_skip_roman_work_but_do_not_catch_refusals() {
    let (_budget, _cancellation, context) = execution("ods-formula-roman-lazy");
    for (source, expected) in [
        ("=IF(TRUE();ROMAN(4);ARABIC([.A1]))", "IV"),
        ("=IF(FALSE();ARABIC([.A1]);ROMAN(9))", "IX"),
        ("=IF(TRUE();ROMAN(4);ROMAN(4000))", "IV"),
        ("=IFERROR(ROMAN(4000);\"fallback\")", "fallback"),
        ("=IFNA(ARABIC(#N/A);\"fallback\")", "fallback"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate lazily: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    for source in [
        "=IFERROR(ARABIC([.A1]);\"fallback\")",
        "=IFNA(ROMAN([.A1]);\"fallback\")",
    ] {
        assert_unsupported(source, UnsupportedKind::Reference);
    }

    let expression = parse("=IF(TRUE();\"ok\";ROMAN(3999))");
    let result = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default()
            .with_max_steps(64)
            .with_max_text_bytes(32_767),
    )
    .expect("the unused Roman branch should not consume Work");
    assert_eq!(text(&result), "ok");
}

#[test]
fn roman_limits_cancellation_and_owned_result_reservations_are_atomic() {
    let expression = parse("=ROMAN(3999)");

    let (budget, cancellation, context) = execution("ods-formula-roman-cancel");
    cancellation.cancel();
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("cancelled Roman evaluation should refuse before work");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Work), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = execution("ods-formula-roman-work-limit");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(0),
    )
    .expect_err("zero Work allowance should refuse atomically");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong Work failure: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = execution("ods-formula-roman-memory-limit");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero storage allowance should refuse atomically");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong Memory failure: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = execution("ods-formula-roman-text-limit");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_text_bytes(8),
    )
    .expect_err("a nine-byte Roman result should exceed an eight-byte text limit");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong text limit failure: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    // Format 4 performs a bounded search after the ordinary evaluator frames
    // have been admitted.  Sweep the local Work limit through that search so
    // a refusal after partial progress is tested separately from max_steps=0.
    let search_expression = parse("=ROMAN(3888;4)");
    let shortest = shortest_roman_lengths(3_888);
    let mut saw_mid_search_refusal = false;
    let mut saw_success = false;
    for max_steps in 0..=600 {
        let scope = format!("ods-formula-roman-mid-search-{max_steps}");
        let (budget, _cancellation, context) = execution(&scope);
        let result = evaluate(
            &search_expression,
            &context,
            &EvaluationLimits::default().with_max_steps(max_steps),
        );
        match result {
            Err(EvaluationFailure::ResourceLimit(limit)) => {
                assert_eq!(limit.resource, Resource::Work);
                assert_eq!(budget.used(Resource::Memory), 0);
                if max_steps >= 32 {
                    saw_mid_search_refusal = true;
                }
            },
            Ok(result) => {
                let spelling = text(&result).to_owned();
                assert_eq!(arabic_reference(&spelling), Some(3_888));
                assert_eq!(spelling.len(), shortest[3_888]);
                drop(result);
                assert_eq!(budget.used(Resource::Memory), 0);
                saw_success = true;
                break;
            },
            Err(error) => panic!("mid-search Roman evaluation returned {error}"),
        }
    }
    assert!(
        saw_mid_search_refusal,
        "the Work sweep never reached a partial Roman search"
    );
    assert!(
        saw_success,
        "ROMAN(3888;4) did not fit within 600 Work units"
    );

    let (budget, _cancellation, context) = execution("ods-formula-roman-result-lifetime");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("ordinary Roman output should evaluate");
    assert_eq!(text(&result), "MMMCMXCIX");
    assert!(result.reserved_output_bytes() >= 9);
    assert!(budget.used(Resource::Memory) > 0);
    drop(context);
    assert!(budget.used(Resource::Memory) > 0);
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = execution_with_memory("ods-formula-roman-tight", 256);
    let expression = parse("=IF(TRUE();\"ok\";ROMAN(3999))");
    let result = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default()
            .with_max_storage_bytes(256)
            .with_max_text_bytes(32_767),
    )
    .expect("an unused output-heavy Roman branch should not consume storage");
    assert_eq!(text(&result), "ok");
    assert_eq!(result.reserved_output_bytes(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn roman_source_spelling_and_force_marker_survive_evaluation() {
    let source = "of:== roman(499;0) ";
    let expression = parse(source);
    assert_eq!(expression.source(), source);
    assert!(expression.is_force_recalculate());
    let (_budget, _cancellation, context) = execution("ods-formula-roman-source");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("function names are case-insensitive in the scalar profile");
    assert_eq!(text(&result), "CDXCIX");
    assert_eq!(expression.source(), source);
    assert!(expression.is_force_recalculate());
}
