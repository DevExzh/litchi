//! Independent integration coverage for the bounded scalar OpenFormula
//! evaluation profile.
//!
//! The cases in this file exercise an already parsed expression with an
//! explicit caller context.  They keep formula errors as values, distinguish
//! evaluator capability refusals from formula results, and check that local
//! limits and cancellation release temporary budget reservations.

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
    let memory_bytes = u64::try_from(memory_bytes).expect("test memory limit fits u64");
    execution_with_core_limits(
        scope,
        CoreLimits::new(
            memory_bytes,
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

fn escaped_literal(decoded_bytes: usize) -> String {
    assert!(decoded_bytes >= 1);
    let mut literal = String::from("\"");
    literal.push_str(&"a".repeat(decoded_bytes - 1));
    literal.push_str("\"\"");
    literal.push('"');
    literal
}

fn evaluate<'a>(
    expression: &'a Expression,
    context: &ExecutionContext,
    limits: &EvaluationLimits,
) -> Result<EvaluatedScalar<'a>, EvaluationFailure> {
    // Keep the context explicit at the integration boundary.  This makes it
    // impossible for a caller to accidentally bypass cancellation or budget
    // charging by using an implicit evaluator context.
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
    let (_budget, _cancellation, context) = execution("ods-formula-unsupported");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("capability-dependent expression should be refused");
    assert!(
        matches!(error, EvaluationFailure::Unsupported(kind) if kind == expected),
        "{source:?} returned the wrong capability error: {error}"
    );
}

#[test]
fn arithmetic_uses_odf_precedence_and_left_associativity() {
    let (_budget, _cancellation, context) = execution("ods-formula-arithmetic");
    let cases = [
        ("=2+3", 5.0),
        ("=7-2", 5.0),
        ("=6*7", 42.0),
        ("=8/2", 4.0),
        ("=2^3", 8.0),
        ("=2^3^2", 64.0),
        ("=2^(-1)", 0.5),
        ("=+2", 2.0),
        ("=-2", -2.0),
        ("=-2^2", 4.0),
        ("=-(2^2)", -4.0),
        ("=50%", 0.5),
        ("=(2+3)%", 0.05),
        ("=2+3*4", 14.0),
        ("=8-3-2", 3.0),
        ("=8/2*2", 8.0),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }
}

#[test]
fn text_and_logical_values_follow_the_declared_scalar_profile() {
    let (_budget, _cancellation, context) = execution("ods-formula-values");
    let cases = [
        ("=\"a\"&\"b\"", "ab"),
        ("=\"\"&\"x\"", "x"),
        ("=1&2", "12"),
        ("=TRUE()&FALSE()", "TRUEFALSE"),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    let expression = parse("=TRUE()");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert!(logical(&result));

    let expression = parse("=FALSE()");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert!(!logical(&result));

    for (source, expected) in [("=TRUE()+2", 3.0), ("=FALSE()+2", 2.0)] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }
}

#[test]
fn comparisons_keep_scalar_types_and_use_exact_text_ordering() {
    let (_budget, _cancellation, context) = execution("ods-formula-comparisons");
    let cases = [
        ("=2=2", true),
        ("=2=3", false),
        ("=2<3", true),
        ("=3<=3", true),
        ("=4>3", true),
        ("=3>=3", true),
        ("=\"a\"=\"a\"", true),
        ("=\"a\"=\"b\"", false),
        ("=1=\"1\"", false),
        ("=2<>3", true),
        ("=2<>2", false),
        ("=TRUE()=TRUE()", true),
        ("=TRUE()=FALSE()", false),
        ("=TRUE()<>FALSE()", true),
        ("=TRUE()>FALSE()", true),
        ("=FALSE()<=TRUE()", true),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(logical(&result), expected, "{source:?}");
    }

    // The published profile compares Unicode scalar values without case
    // folding or normalization.  The decomposed and precomposed spellings
    // therefore remain distinct, and the insensitive host policy is deferred.
    let expression = parse("=\"\u{00e9}\"=\"e\u{0301}\"");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert!(!logical(&result));
    let expression = parse("=\"\u{00e9}\">\"e\u{0301}\"");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert!(logical(&result));
}

#[test]
fn formula_errors_are_values_and_propagate_without_becoming_rust_failures() {
    let (_budget, _cancellation, context) = execution("ods-formula-errors");
    let cases = [
        ("=#N/A", ScalarError::NotAvailable),
        ("=#N/A+1", ScalarError::NotAvailable),
        ("=1-#DIV/0!", ScalarError::DivisionByZero),
        ("=#N/A*2", ScalarError::NotAvailable),
        ("=#N/A^2", ScalarError::NotAvailable),
        ("=#N/A<>1", ScalarError::NotAvailable),
        ("=#N/A<1", ScalarError::NotAvailable),
        ("=#N/A&\"x\"", ScalarError::NotAvailable),
        ("=#N/A=#N/A", ScalarError::NotAvailable),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should produce a formula value: {error}"));
        assert_eq!(formula_error(&result), expected, "{source:?}");
    }
}

#[test]
fn declared_text_conversion_and_mixed_ordering_profile_are_typed() {
    let (_budget, _cancellation, context) = execution("ods-formula-conversion");
    let expression = parse("=\"2\"+1");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert_eq!(number(&result), 3.0);

    let expression = parse("=\"not-a-number\"+1");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert_eq!(formula_error(&result), ScalarError::Value);

    let expression = parse("=1<\"2\"");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert_eq!(formula_error(&result), ScalarError::Value);
}

#[test]
fn numeric_domain_and_nonfinite_results_become_formula_number_errors() {
    let (_budget, _cancellation, context) = execution("ods-formula-number-errors");
    for source in ["=1e9999", "=1e308*1e308", "=0^(-1)"] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula result: {error}"));
        assert_eq!(formula_error(&result), ScalarError::Number, "{source:?}");
    }

    // §6.16.46 permits 0^0 to be 0, 1, or an Error.  This profile documents
    // and selects 1, so this is a profile assertion rather than a universal
    // OpenFormula claim.
    let expression = parse("=0^0");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert_eq!(number(&result), 1.0);
}

#[test]
fn source_and_forced_recalculation_marker_survive_evaluation() {
    let source = "of:== 2 + 3 ";
    let expression = parse(source);
    assert_eq!(expression.source(), source);
    assert!(expression.is_force_recalculate());
    let (_budget, _cancellation, context) = execution("ods-formula-source");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert_eq!(number(&result), 5.0);
    assert_eq!(expression.source(), source);
    assert!(expression.is_force_recalculate());
}

#[test]
fn context_dependent_features_are_explicit_capability_refusals() {
    assert_unsupported("=[.A1]", UnsupportedKind::Reference);
    assert_unsupported("=[.A1:.B2]", UnsupportedKind::Reference);
    assert_unsupported("=SUMIF([.A1];1)", UnsupportedKind::Reference);
    assert_unsupported("={1;2|3}", UnsupportedKind::Array);
    assert_unsupported("=Named", UnsupportedKind::NamedExpression);
    assert_unsupported("='Column Label'", UnsupportedKind::Label);
    let expression = parse("=SUMIF(1;1)");
    let (_budget, _cancellation, context) = execution("ods-formula-conditional-constant");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("a conditional aggregate with a constant range is a formula value");
    assert_eq!(formula_error(&result), ScalarError::Value);
}

#[test]
fn supported_boolean_functions_eagerly_evaluate_children_before_arity_errors() {
    let expression = parse("=TRUE([.A1])");
    let (_budget, _cancellation, context) = execution("ods-formula-eager-reference");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("TRUE(reference) must visit the reference child");
    assert!(
        matches!(
            &error,
            EvaluationFailure::Unsupported(UnsupportedKind::Reference)
        ),
        "reference child was bypassed: {error}"
    );

    let expression = parse("=FALSE(1/0)");
    let (_budget, _cancellation, context) = execution("ods-formula-eager-division");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("FALSE(1/0) should retain the child formula error");
    assert_eq!(formula_error(&result), ScalarError::DivisionByZero);

    let expression = parse("=TRUE(#N/A;#DIV/0!)");
    let (_budget, _cancellation, context) = execution("ods-formula-eager-errors");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("boolean arity handling should retain child errors");
    assert_eq!(
        formula_error(&result),
        ScalarError::NotAvailable,
        "the first source-ordered child error must win"
    );

    let (budget, _cancellation, context) = execution("ods-formula-eager-arity");
    let expression = parse("=TRUE(1+2)");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("a wrong-arity scalar call should produce a formula value");
    assert_eq!(formula_error(&result), ScalarError::Value);
    assert!(
        budget.used(Resource::Work) > 0,
        "argument evaluation must consume work before the arity value"
    );

    let (budget, _cancellation, context) = execution("ods-formula-eager-work");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(8),
    )
    .expect_err("a child that exceeds Work must not be bypassed by arity handling");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "child work failure was bypassed: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn cancellation_is_checked_before_work_and_leaves_no_reservation() {
    let (budget, cancellation, context) = execution("ods-formula-cancel");
    cancellation.cancel();
    let expression = parse("=1+2");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("cancelled evaluation should refuse");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Work), 0);
}

#[test]
fn zero_step_and_zero_stack_limits_are_typed_and_atomic() {
    let (budget, _cancellation, context) = execution("ods-formula-step-limit");
    let expression = parse("=1+2");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(0),
    )
    .expect_err("zero step allowance should refuse before the first step");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work)
    );
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Work), 0);

    let (budget, _cancellation, context) = execution("ods-formula-stack-limit");
    let expression = parse("=1");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_stack_entries(0),
    )
    .expect_err("zero stack allowance should refuse before adoption");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects)
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn text_limit_failure_is_typed_and_releases_temporary_storage() {
    let (budget, _cancellation, context) = execution("ods-formula-text-limit");
    let expression = parse("=\"abcdefgh\"&\"ijkl\"");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_text_bytes(11),
    )
    .expect_err("the twelve-byte result should exceed the eleven-byte limit");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory)
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn decoded_quoted_text_obeys_the_decoded_byte_boundary() {
    let expression = parse("=\"a\"\"b\"");
    let (_budget, _cancellation, context) = execution("ods-formula-decoded-text");
    let result = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_text_bytes(3),
    )
    .expect("decoded three-byte text should fit exactly");
    assert_eq!(text(&result), "a\"b");

    let (budget, _cancellation, context) = execution("ods-formula-decoded-text-short");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_text_bytes(2),
    )
    .expect_err("decoded three-byte text should exceed a two-byte limit");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory)
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn escaped_literal_copy_has_its_own_work_boundary_after_source_scan() {
    let expression = parse("=\"aaaa\"\"bbbb\"");
    let root_text = expression.root().text();
    let body = root_text
        .strip_prefix('"')
        .and_then(|value| value.strip_suffix('"'))
        .expect("string root delimiters");
    let escaped_quotes = body.matches("\"\"").count();
    let decoded_bytes = body.len() - escaped_quotes;
    assert_eq!(body.len(), 10);
    assert_eq!(decoded_bytes, 9);

    let (budget, _cancellation, context) = execution("ods-formula-string-copy-work");
    // Evaluation charges the initial source scan and the second source scan
    // used to locate copy segments before charging decoded output bytes.
    let source_scan_only = 1 + 2 * body.len() as u64;
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(source_scan_only),
    )
    .expect_err("source scan should fit while decoded copy exceeds Work");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "decoded copy returned the wrong result: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (_budget, _cancellation, context) = execution("ods-formula-string-copy-exact");
    let exact = EvaluationLimits::default().with_max_steps(source_scan_only + decoded_bytes as u64);
    let result = evaluate(&expression, &context, &exact)
        .expect("source scan plus decoded copy should fit exactly");
    assert_eq!(text(&result), "aaaa\"bbbb");
}

#[test]
fn cancellation_before_a_long_escaped_quote_scan_is_atomic() {
    let mut source = String::from("=\"");
    for _ in 0..4_097 {
        source.push_str("\"\"");
    }
    source.push('"');
    let expression = parse(&source);
    let (budget, cancellation, context) = execution("ods-formula-string-cancel");
    cancellation.cancel();
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("cancelled long escaped literal should refuse before scanning");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Work), 0);
}

#[test]
fn borrowed_plain_text_has_no_output_reservation() {
    let (budget, _cancellation, context) = execution("ods-formula-borrowed-text");
    let expression = parse("=\"plain\"");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert_eq!(text(&result), "plain");
    assert_eq!(result.reserved_output_bytes(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn owned_text_result_keeps_its_output_reservation_until_drop() {
    let (budget, _cancellation, context) = execution("ods-formula-text-result");
    let expression = parse("=\"a\"\"b\"");
    let result = evaluate(&expression, &context, &EvaluationLimits::default()).unwrap();
    assert_eq!(text(&result), "a\"b");
    assert!(result.reserved_output_bytes() > 0);
    assert!(budget.used(Resource::Memory) > 0);
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn storage_admission_accounts_for_live_operands_and_ancestor_budget() {
    const DECODED_BYTES: usize = 2_048;
    const STORAGE_CAP: usize = 3_000;
    const TEXT_CAP: usize = 4_096;
    let left = escaped_literal(DECODED_BYTES);
    let right = escaped_literal(DECODED_BYTES);
    let single_source = format!("={left}");
    let combined_source = format!("={left}&{right}");
    let single = parse(&single_source);
    let combined = parse(&combined_source);
    let limits = EvaluationLimits::default()
        .with_max_storage_bytes(STORAGE_CAP)
        .with_max_text_bytes(TEXT_CAP);

    // Each operand fits by itself under the cap, including its escaped quote
    // decode and evaluator stack.  This prevents a test that only proves the
    // cap is smaller than one allocation.
    let (single_budget, _cancellation, single_context) = execution("ods-formula-storage-single");
    let single_result = evaluate(&single, &single_context, &limits)
        .expect("one 2048-byte decoded operand should fit under 3000 bytes");
    assert_eq!(text(&single_result).len(), DECODED_BYTES);
    assert!(single_result.reserved_output_bytes() >= DECODED_BYTES);
    drop(single_result);
    assert_eq!(single_budget.used(Resource::Memory), 0);

    // During the binary operation both decoded operands are live.  The
    // aggregate admission must fail before adopting the second operand or
    // silently growing an uncharged result buffer.
    let (budget, _cancellation, context) = execution("ods-formula-storage-local");
    let error = evaluate(&combined, &context, &limits)
        .expect_err("two live 2048-byte operands should exceed the 3000-byte cap");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory)
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    // The same local operation must also respect a tighter ancestor budget.
    let (budget, _cancellation, context) =
        execution_with_memory("ods-formula-storage-ancestor-short", STORAGE_CAP);
    let error = evaluate(
        &combined,
        &context,
        &EvaluationLimits::default()
            .with_max_storage_bytes(TEXT_CAP)
            .with_max_text_bytes(TEXT_CAP),
    )
    .expect_err("the ancestor budget must cap aggregate evaluator storage");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory)
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    // A successful output reservation remains live independently of the
    // caller's context handle and is released exactly when the result drops.
    let (budget, _cancellation, context) =
        execution_with_memory("ods-formula-storage-ancestor-output", STORAGE_CAP);
    let result = evaluate(
        &single,
        &context,
        &EvaluationLimits::default()
            .with_max_storage_bytes(TEXT_CAP)
            .with_max_text_bytes(TEXT_CAP),
    )
    .expect("one decoded operand should fit the ancestor budget");
    assert_eq!(text(&result).len(), DECODED_BYTES);
    assert!(budget.used(Resource::Memory) > 0);
    drop(context);
    assert!(
        budget.used(Resource::Memory) > 0,
        "dropping the execution context must not release a live output"
    );
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn storage_limit_refusal_is_atomic_for_owned_text() {
    let (budget, _cancellation, context) = execution("ods-formula-storage-limit");
    let expression = parse("=\"a\"\"b\"");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_storage_bytes(2),
    )
    .expect_err("decoded three-byte text should exceed storage limit two");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory)
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn number_to_text_reservation_covers_tiny_and_large_finite_values() {
    let (_budget, _cancellation, context) = execution("ods-formula-number-text");
    for source in [
        "=0.000000000000000000000000000000000000000001&\"\"",
        "=10000000000000000000000000000000000000000&\"\"",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        let value = text(&result);
        assert!(!value.is_empty(), "{source:?} produced empty text");
        assert!(
            value.parse::<f64>().is_ok(),
            "{source:?} produced nonnumeric text {value:?}"
        );
        assert!(
            result.reserved_output_bytes() >= value.len(),
            "{source:?} reserved {} bytes for {} output bytes",
            result.reserved_output_bytes(),
            value.len()
        );
        if source.starts_with("=100") {
            assert!(
                value.len() > 32,
                "large finite value should exercise growth"
            );
        }
    }
}

#[test]
fn repeated_escaped_literals_and_left_associative_concat_keep_full_output_budgeted() {
    const FRAGMENTS: usize = 128;
    let mut source = String::from("=");
    for index in 0..FRAGMENTS {
        if index != 0 {
            source.push('&');
        }
        source.push_str("\"x\"\"y\"");
    }
    let expression = parse(&source);
    let (_budget, _cancellation, context) = execution("ods-formula-concat-growth");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("escaped literal concatenation should evaluate");
    let expected = "x\"y".repeat(FRAGMENTS);
    assert_eq!(text(&result), expected);
    assert!(result.reserved_output_bytes() >= expected.len());
}

#[test]
fn long_comparison_chain_honors_the_work_limit() {
    let mut source = String::from("=1");
    for value in 2..=256 {
        source.push('<');
        source.push_str(&value.to_string());
    }
    let expression = parse(&source);
    let (budget, _cancellation, context) = execution("ods-formula-comparison-work");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_steps(8),
    )
    .expect_err("a long comparison chain should exceed eight evaluator steps");
    assert!(
        matches!(error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work)
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn byte_work_limit_rejects_literal_number_and_comparison_scans() {
    for source in [
        "=\"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\"",
        "=12345678901234567890123456789012",
        "=1<2",
    ] {
        let expression = parse(source);
        let (budget, _cancellation, context) = execution("ods-formula-byte-work");
        let error = evaluate(
            &expression,
            &context,
            &EvaluationLimits::default().with_max_steps(2),
        )
        .expect_err("byte accounting should stop the scan under two work units");
        assert!(
            matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
            "{source:?} returned the wrong limit: {error}"
        );
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}

#[test]
fn long_left_chain_evaluates_iteratively_under_explicit_work_limit() {
    const TERMS: usize = 8_192;
    let mut source = String::from("=");
    for index in 0..TERMS {
        if index != 0 {
            source.push('+');
        }
        source.push('1');
    }
    let expression = parse(&source);
    let (_budget, _cancellation, context) = execution("ods-formula-long-chain");
    let limits = EvaluationLimits::default()
        .with_max_steps(100_000)
        .with_max_stack_entries(16_384);
    let result = evaluate(&expression, &context, &limits)
        .unwrap_or_else(|error| panic!("long scalar chain should evaluate: {error}"));
    assert_eq!(number(&result), TERMS as f64);
}
