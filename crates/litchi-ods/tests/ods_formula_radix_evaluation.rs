//! Independent integration coverage for the OpenFormula 1.4 radix conversion
//! functions in section 6.19.
//!
//! The conversion vectors in this file are taken from the specified digit
//! alphabets and two's-complement widths, rather than from the evaluator's
//! implementation.  Formula errors remain scalar values; capability,
//! resource, and cancellation failures remain typed evaluator results.

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

const MAX_BINARY_MAGNITUDE: i64 = 511;
const MIN_BINARY_SIGNED: i64 = -512;
const MAX_HEX_MAGNITUDE: i64 = (1_i64 << 39) - 1;
const MAX_OCTAL_MAGNITUDE: i64 = (1_i64 << 29) - 1;
const F64_MAX_BASE10: &str = "179769313486231570814527423731704356798070567525844996598917476803157260780028538760589558632766878171540458953514382464234321326889464182768467546703537516986049910576551282076245490090389328944075868508455133942304583236903222948165808559332123348274797826204144723168738177180919299881250404026184124858368";
const F64_MAX_BASE36: &str = "1A1E4VNGAIKU6SCYIL2A1VCBG6QVBZJFSEU5NTY6QYR6FT0FMXYR3NMWLM21AXDQ6ED914EDAR7ZMC0M6NPHL75RAN22ULSB7GK2X0W8EH76J4MR2DVCV9TLPR9QO3AP6MY00O4K4HHS2393945UO1RSPBZ2QHHHVHWP0K4Z956E50710Y4RP9WVBY29LPSVD8XLURK";

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
    let (_budget, _cancellation, context) = execution("ods-formula-radix-error");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
    assert_eq!(formula_error(&result), expected, "{source:?}");
}

fn assert_unsupported(source: &str, expected: UnsupportedKind) {
    let expression = parse(source);
    let (_budget, _cancellation, context) = execution("ods-formula-radix-unsupported");
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("capability-dependent radix expression should be refused");
    assert!(
        matches!(&error, EvaluationFailure::Unsupported(kind) if *kind == expected),
        "{source:?} returned the wrong capability error: {error}"
    );
}

#[test]
fn base_and_decimal_use_the_specified_digit_alphabet() {
    let (_budget, _cancellation, context) = execution("ods-formula-radix-base-decimal");
    let text_cases = [
        ("=BASE(0;2)", "0"),
        ("=BASE(1;2)", "1"),
        ("=BASE(15;16)", "F"),
        ("=BASE(35;36)", "Z"),
        ("=BASE(255;16)", "FF"),
        ("=BASE(45745;36)", "ZAP"),
        ("=BASE(511;2)", "111111111"),
        ("=BASE(65535;16)", "FFFF"),
    ];
    for (source, expected) in text_cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    let number_cases = [
        ("=DECIMAL(\"0\";2)", 0.0),
        ("=DECIMAL(\"11111111\";2)", 255.0),
        ("=DECIMAL(\"FF\";16)", 255.0),
        ("=DECIMAL(\"7f\";16)", 127.0),
        ("=DECIMAL(\"ZAP\";36)", 45745.0),
        ("=DECIMAL(\"zap\";36)", 45745.0),
    ];
    for (source, expected) in number_cases {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    // These are independent composition checks: the textual result is then
    // consumed by a different parser, not merely compared with itself.
    for (source, expected) in [
        ("=DECIMAL(BASE(45745;36);36)", 45745.0),
        ("=HEX2DEC(BASE(45745;16))", 45745.0),
        ("=OCT2DEC(DEC2OCT(45745))", 45745.0),
        ("=DECIMAL(DEC2HEX(45745);16)", 45745.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }
}

#[test]
fn base_padding_and_large_finite_numbers_are_exact() {
    let (_budget, _cancellation, context) = execution("ods-formula-radix-base-limits");
    for (source, expected) in [
        ("=BASE(0;2;4)", "0000"),
        ("=BASE(5;2;4)", "0101"),
        ("=BASE(5;16;4)", "0005"),
        ("=BASE(65535;16;2)", "FFFF"),
        ("=BASE(45745;36;6)", "000ZAP"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    // Both values are exactly representable in f64.  The expected spellings
    // are hand-checked powers-of-two boundaries, including all 53 low bits.
    for (source, expected) in [
        ("=BASE(9007199254740991;16)", "1FFFFFFFFFFFFF"),
        (
            "=BASE(9007199254740992;2)",
            "100000000000000000000000000000000000000000000000000000",
        ),
        (
            "=BASE(281474976710655;2)",
            "111111111111111111111111111111111111111111111111",
        ),
        ("=BASE(281474976710655;16)", "FFFFFFFFFFFF"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    // f64::MAX is itself an exact integer: (2^53 - 1) * 2^971.  These
    // independent arbitrary-integer spellings ensure the implementation does
    // not first narrow a finite Number to u64.  The binary oracle is assembled
    // from that exact representation rather than by reusing BASE.
    for (source, expected) in [
        ("=BASE(1.7976931348623157e308;10)", F64_MAX_BASE10),
        ("=BASE(1.7976931348623157e308;36)", F64_MAX_BASE36),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }
    let mut expected_binary = "1".repeat(53);
    expected_binary.push_str(&"0".repeat(971));
    let expression = parse("=BASE(1.7976931348623157e308;2)");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("BASE should support every finite f64 integer");
    assert_eq!(text(&result), expected_binary);

    let expression = parse("=DECIMAL(\"1FFFFFFFFFFFFF\";16)");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("the 53-bit hexadecimal integer should convert");
    assert_eq!(number(&result).to_bits(), (9007199254740991_f64).to_bits());

    // 9007199254740993 is exactly halfway between two representable f64
    // values; the selected scalar profile uses IEEE nearest-even conversion.
    let expression = parse("=DECIMAL(\"9007199254740993\";10)");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("the large decimal should remain a finite Number");
    assert_eq!(number(&result).to_bits(), (9007199254740992_f64).to_bits());
}

#[test]
fn decimal_ignores_only_documented_prefix_suffix_and_leading_tabs() {
    let (_budget, _cancellation, context) = execution("ods-formula-radix-decimal-spelling");
    for (source, expected) in [
        ("=DECIMAL(\" \t FF\";16)", 255.0),
        ("=DECIMAL(\"0xFF\";16)", 255.0),
        ("=DECIMAL(\"xFF\";16)", 255.0),
        ("=DECIMAL(\"0XffH\";16)", 255.0),
        ("=DECIMAL(\"\t101b\";2)", 5.0),
        ("=DECIMAL(\"  \t 101\";10)", 101.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    for source in [
        "=DECIMAL(\"FF \";16)",
        "=DECIMAL(\"0xFG\";16)",
        "=DECIMAL(\"ＦＦ\";16)",
        "=DECIMAL(\"101b\";10)",
        "=DECIMAL(\"0x10\";10)",
    ] {
        assert_formula_error(source, ScalarError::Value);
    }
}

#[test]
fn radix_functions_distinguish_text_digits_from_numeric_values() {
    let (_budget, _cancellation, context) = execution("ods-formula-radix-text-number");
    // TextOrNumber conversion functions interpret a numeric argument by its
    // decimal digit spelling.  A quoted value is never silently interpreted
    // in base ten first.
    for (source, expected) in [
        ("=BIN2DEC(\"10\")", 2.0),
        ("=BIN2DEC(10)", 2.0),
        ("=HEX2DEC(\"10\")", 16.0),
        ("=HEX2DEC(10)", 16.0),
        ("=OCT2DEC(\"10\")", 8.0),
        ("=OCT2DEC(10)", 8.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    let expression = parse("=BASE(\"15\";16)");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("a decimal text integer should be accepted by Integer conversion");
    assert_eq!(text(&result), "F");

    // Conversion functions accepting an integer use the profile's truncation
    // toward zero for numeric arguments.  Text remains strict: a decimal
    // point is not an accepted digit in a base conversion input.
    let expression = parse("=BASE(15.9;16)");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("BASE uses the selected profile's truncating Integer conversion");
    assert_eq!(text(&result), "F");

    // The twelve direct converters require INT(X)=X.  A fractional Number X
    // therefore produces a formula Value error; this does not apply to BASE,
    // whose generic Integer argument uses the profile conversion above.
    for source in [
        "=BIN2DEC(1.5)",
        "=BIN2HEX(1.5)",
        "=BIN2OCT(1.5)",
        "=DEC2BIN(5.9)",
        "=DEC2HEX(255.9)",
        "=DEC2OCT(8.9)",
        "=HEX2BIN(1.5)",
        "=HEX2DEC(1.5)",
        "=HEX2OCT(1.5)",
        "=OCT2BIN(1.5)",
        "=OCT2DEC(1.5)",
        "=OCT2HEX(1.5)",
    ] {
        assert_formula_error(source, ScalarError::Value);
    }
    for source in ["=DEC2BIN(\"5.9\")", "=BIN2DEC(\"10.1\")"] {
        assert_formula_error(source, ScalarError::Value);
    }
}

#[test]
fn binary_conversion_uses_ten_bit_twos_complement_and_known_vectors() {
    let (_budget, _cancellation, context) = execution("ods-formula-radix-binary");
    for (source, expected) in [
        ("=BIN2DEC(\"0000000000\")", 0.0),
        ("=BIN2DEC(\"0000000001\")", 1.0),
        ("=BIN2DEC(\"0111111111\")", MAX_BINARY_MAGNITUDE as f64),
        ("=BIN2DEC(\"1000000000\")", MIN_BINARY_SIGNED as f64),
        ("=BIN2DEC(\"1111111111\")", -1.0),
        ("=BIN2DEC(-0)", 0.0),
        ("=BIN2DEC(101)", 5.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    for (source, expected) in [
        ("=BIN2HEX(\"0000000010\")", "2"),
        ("=BIN2HEX(\"0000000010\";4)", "0002"),
        ("=BIN2HEX(\"1111111111\")", "FFFFFFFFFF"),
        ("=BIN2HEX(\"1000000000\")", "FFFFFFFE00"),
        ("=BIN2OCT(\"0000000010\")", "2"),
        ("=BIN2OCT(\"0000000010\";4)", "0002"),
        ("=BIN2OCT(\"1111111111\")", "7777777777"),
        ("=BIN2OCT(\"1000000000\")", "7777777000"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }
}

#[test]
fn decimal_hex_and_octal_conversions_cover_signed_width_boundaries() {
    let (_budget, _cancellation, context) = execution("ods-formula-radix-signed");
    for (source, expected) in [
        ("=DEC2BIN(-0)", "0"),
        ("=DEC2BIN(-512)", "1000000000"),
        ("=DEC2BIN(-1)", "1111111111"),
        ("=DEC2BIN(511)", "111111111"),
        ("=DEC2HEX(-0)", "0"),
        ("=DEC2HEX(-549755813888)", "8000000000"),
        ("=DEC2HEX(549755813887)", "7FFFFFFFFF"),
        ("=DEC2HEX(-1)", "FFFFFFFFFF"),
        ("=DEC2OCT(-0)", "0"),
        ("=DEC2OCT(-536870912)", "4000000000"),
        ("=DEC2OCT(536870911)", "3777777777"),
        ("=DEC2OCT(-1)", "7777777777"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    for (source, expected) in [
        ("=HEX2DEC(\"0000000000\")", 0.0),
        ("=HEX2DEC(\"7FFFFFFFFF\")", MAX_HEX_MAGNITUDE as f64),
        ("=HEX2DEC(\"FFFFFFFFFF\")", -1.0),
        ("=HEX2DEC(-0)", 0.0),
        ("=OCT2DEC(\"0000000000\")", 0.0),
        ("=OCT2DEC(\"3777777777\")", MAX_OCTAL_MAGNITUDE as f64),
        ("=OCT2DEC(\"7777777777\")", -1.0),
        ("=OCT2DEC(-0)", 0.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(number(&result), expected, "{source:?}");
    }

    for (source, expected) in [
        ("=HEX2BIN(\"0000000001\")", "1"),
        ("=HEX2BIN(\"0000000001\";10)", "0000000001"),
        ("=HEX2BIN(\"FFFFFFFFFF\")", "1111111111"),
        ("=HEX2OCT(\"0000000001\")", "1"),
        ("=HEX2OCT(\"FFFFFFFFFF\")", "7777777777"),
        ("=OCT2BIN(\"0000000001\")", "1"),
        ("=OCT2BIN(\"7777777777\")", "1111111111"),
        ("=OCT2HEX(\"0000000001\")", "1"),
        ("=OCT2HEX(\"7777777777\")", "FFFFFFFFFF"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }
}

#[test]
fn positive_results_pad_and_negative_results_ignore_digits() {
    let (_budget, _cancellation, context) = execution("ods-formula-radix-padding");
    for (source, expected) in [
        ("=BIN2HEX(\"101\";4)", "0005"),
        ("=BIN2OCT(\"101\";4)", "0005"),
        ("=DEC2BIN(5;5)", "00101"),
        ("=DEC2HEX(5;4)", "0005"),
        ("=DEC2OCT(5;4)", "0005"),
        ("=HEX2BIN(\"5\";5)", "00101"),
        ("=HEX2OCT(\"5\";4)", "0005"),
        ("=OCT2BIN(\"5\";5)", "00101"),
        ("=OCT2HEX(\"5\";4)", "0005"),
        ("=DEC2BIN(5;4.9)", "0101"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    // Negative conversion results have a fixed sign-extension width, so a
    // requested positive-width padding value cannot truncate them.
    for (source, expected) in [
        ("=BIN2HEX(\"1111111111\";1)", "FFFFFFFFFF"),
        ("=BIN2OCT(\"1111111111\";1)", "7777777777"),
        ("=DEC2BIN(-1;1)", "1111111111"),
        ("=DEC2HEX(-1;1)", "FFFFFFFFFF"),
        ("=DEC2OCT(-1;1)", "7777777777"),
        ("=HEX2BIN(\"FFFFFFFFFF\";1)", "1111111111"),
        ("=HEX2OCT(\"FFFFFFFFFF\";1)", "7777777777"),
        ("=OCT2BIN(\"7777777777\";1)", "1111111111"),
        ("=OCT2HEX(\"7777777777\";1)", "FFFFFFFFFF"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    for source in [
        "=BIN2HEX(\"1\";-1)",
        "=BIN2OCT(\"1\";-1)",
        "=DEC2BIN(1;-1)",
        "=DEC2HEX(1;-1)",
        "=DEC2OCT(1;-1)",
    ] {
        assert_formula_error(source, ScalarError::Number);
    }
}

#[test]
fn invalid_digits_ranges_and_arity_are_formula_errors() {
    for (source, expected) in [
        ("=BIN2DEC(\"102\")", ScalarError::Value),
        ("=BIN2HEX(\"102\")", ScalarError::Value),
        ("=BIN2OCT(\"108\")", ScalarError::Value),
        ("=DEC2BIN(\"5x\")", ScalarError::Value),
        ("=HEX2DEC(\"GG\")", ScalarError::Value),
        ("=OCT2DEC(\"8\")", ScalarError::Value),
        ("=DECIMAL(\"G\";16)", ScalarError::Value),
        ("=BASE(-1;10)", ScalarError::Number),
        ("=BASE(1;1)", ScalarError::Number),
        ("=BASE(1;37)", ScalarError::Number),
        ("=DECIMAL(\"1\";1)", ScalarError::Number),
        ("=DECIMAL(\"1\";37)", ScalarError::Number),
        ("=DEC2BIN(512)", ScalarError::Number),
        ("=DEC2HEX(549755813888)", ScalarError::Number),
        ("=DEC2OCT(536870912)", ScalarError::Number),
    ] {
        assert_formula_error(source, expected);
    }

    for source in [
        "=BIN2DEC()",
        "=DECIMAL(\"1\")",
        "=BASE(1;2;3;4)",
        "=DEC2BIN(1;2;3)",
        "=HEX2DEC(\"1\";16)",
    ] {
        assert_formula_error(source, ScalarError::Value);
    }
}

#[test]
fn radix_calls_are_eager_for_present_children_even_when_arity_is_wrong() {
    for source in [
        "=BIN2DEC(1;[.A1])",
        "=DECIMAL(1;2;[.A1])",
        "=BASE(1;2;3;[.A1])",
        "=DEC2HEX(1;[.A1];3)",
        "=HEX2BIN(1;[.A1];3)",
    ] {
        assert_unsupported(source, UnsupportedKind::Reference);
    }

    for (source, expected) in [
        ("=BASE(#N/A;#DIV/0!)", ScalarError::NotAvailable),
        ("=DECIMAL(#N/A;#DIV/0!)", ScalarError::NotAvailable),
        ("=DEC2HEX(#REF!;#DIV/0!)", ScalarError::Reference),
        ("=BIN2DEC(#NUM!;#VALUE!)", ScalarError::Number),
    ] {
        assert_formula_error(source, expected);
    }
}

#[test]
fn conditional_handlers_skip_radix_work_but_do_not_catch_refusals() {
    let (_budget, _cancellation, context) = execution("ods-formula-radix-lazy");
    for (source, expected) in [
        ("=IF(TRUE();BASE(5;2);DECIMAL(\"GG\";16))", "101"),
        ("=IF(FALSE();DECIMAL(\"GG\";16);DEC2HEX(5))", "5"),
        ("=IF(TRUE();DEC2HEX(255);BASE(1;1))", "FF"),
    ] {
        let expression = parse(source);
        let result =
            evaluate(&expression, &context, &EvaluationLimits::default()).unwrap_or_else(|error| {
                panic!("{source:?} should skip the unused radix branch: {error}")
            });
        assert_eq!(text(&result), expected, "{source:?}");
    }

    for (source, expected) in [
        ("=IFERROR(DECIMAL(\"GG\";16);\"fallback\")", "fallback"),
        ("=IFERROR(DEC2BIN(512);\"fallback\")", "fallback"),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &context, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should catch the formula error: {error}"));
        assert_eq!(text(&result), expected, "{source:?}");
    }

    for source in [
        "=IFERROR(DECIMAL([.A1];16);\"fallback\")",
        "=IFNA(BASE([.A1];16);\"fallback\")",
    ] {
        let error = evaluate(&parse(source), &context, &EvaluationLimits::default())
            .expect_err("capability failures must not become formula alternatives");
        assert!(
            matches!(
                error,
                EvaluationFailure::Unsupported(UnsupportedKind::Reference)
            ),
            "{source:?} returned the wrong handler result: {error}"
        );
    }

    // A decisive IF branch does not admit the very large output of the
    // unselected conversion.  The text limit stays finite to make accidental
    // eager evaluation visible without a timing assertion.
    let expression = parse("=IF(TRUE();\"ok\";BASE(1;2;32767))");
    let result = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default()
            .with_max_steps(64)
            .with_max_text_bytes(32_767),
    )
    .expect("the unused output-heavy radix branch should remain lazy");
    assert_eq!(text(&result), "ok");
}

#[test]
fn limits_cancellation_and_owned_radix_results_are_atomic() {
    let expression = parse("=BASE(45745;36;8)");

    let (budget, cancellation, context) = execution("ods-formula-radix-cancel");
    cancellation.cancel();
    let error = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect_err("cancelled radix evaluation should refuse before work");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Work), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = execution("ods-formula-radix-work-limit");
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

    let (budget, _cancellation, context) = execution("ods-formula-radix-memory-limit");
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

    let (budget, _cancellation, context) = execution("ods-formula-radix-text-limit");
    let error = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default().with_max_text_bytes(7),
    )
    .expect_err("an eight-byte BASE result should exceed a seven-byte text limit");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong text limit failure: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = execution("ods-formula-radix-result-lifetime");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("ordinary BASE output should evaluate");
    assert_eq!(text(&result), "00000ZAP");
    assert!(result.reserved_output_bytes() >= 5);
    assert!(budget.used(Resource::Memory) > 0);
    drop(context);
    assert!(budget.used(Resource::Memory) > 0);
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) =
        execution_with_memory("ods-formula-radix-memory-profile", 1024);
    let expression = parse("=IF(TRUE();\"ok\";BASE(1;2;32767))");
    let result = evaluate(
        &expression,
        &context,
        &EvaluationLimits::default()
            .with_max_storage_bytes(1024)
            .with_max_text_bytes(32_767),
    )
    .expect("the unselected large radix result should not consume storage");
    assert_eq!(text(&result), "ok");
    assert_eq!(result.reserved_output_bytes(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let selected_expression = parse("=BASE(1;2;32767)");
    let error = evaluate(
        &selected_expression,
        &context,
        &EvaluationLimits::default()
            .with_max_storage_bytes(1024)
            .with_max_text_bytes(32_767),
    )
    .expect_err("the selected large radix result should exceed the storage limit");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong selected-output failure: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn radix_source_spelling_and_force_marker_survive_evaluation() {
    let source = "of:== dec2hex(255;4) ";
    let expression = parse(source);
    assert_eq!(expression.source(), source);
    assert!(expression.is_force_recalculate());
    let (_budget, _cancellation, context) = execution("ods-formula-radix-source");
    let result = evaluate(&expression, &context, &EvaluationLimits::default())
        .expect("function names are case-insensitive in the scalar profile");
    assert_eq!(text(&result), "00FF");
    assert_eq!(expression.source(), source);
    assert!(expression.is_force_recalculate());
}
