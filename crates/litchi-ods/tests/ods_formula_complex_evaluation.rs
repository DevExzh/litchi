//! Independent integration coverage for the OpenFormula 1.4 complex-number
//! functions.
//!
//! The contract for this file is recorded in
//! `docs/report/spec-gap-validation-evidence/ods-formula-complex-functions/contract.md`.
//! Complex results are observed through `IMREAL` and `IMAGINARY`, which keeps
//! these tests independent of the public representation chosen for a complex
//! scalar.  Formula errors remain values; evaluator cancellation and resource
//! failures remain typed Rust results.

use std::{cell::Cell, f64::consts::PI};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::{
        self, EvaluatedScalar, EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError,
        ScalarValue,
    },
    expression::Expression,
};

fn new_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(1 << 20).expect("nonzero in-flight byte limit"),
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
    evaluation::evaluate_scalar(expression, &EvaluationContext::new(execution), limits)
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

fn assert_close(actual: f64, expected: f64, source: &str) {
    let tolerance = if expected == 0.0 {
        2e-12
    } else {
        (expected.abs() * 2e-12).max(f64::from_bits(1) * 16.0)
    };
    assert!(
        (actual - expected).abs() <= tolerance,
        "{source:?}: expected {expected:.17e}, got {actual:.17e} (tol {tolerance:.3e})"
    );
}

fn assert_real(
    source: &str,
    execution: &ExecutionContext,
    expected: f64,
) -> Result<(), EvaluationFailure> {
    let expression = parse(source);
    let result = evaluate(&expression, execution, &EvaluationLimits::default())?;
    assert_close(number(&result), expected, source);
    Ok(())
}

#[test]
fn every_complex_function_has_a_numeric_projection_vector() {
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-basic");

    // Representation, components, and the arithmetic family.
    for (source, expected) in [
        ("=IMREAL(COMPLEX(3;4))", 3.0),
        ("=IMAGINARY(COMPLEX(3;4;\"j\"))", 4.0),
        ("=IMABS(COMPLEX(3;4))", 5.0),
        ("=IMARGUMENT(COMPLEX(0;1))", PI / 2.0),
        ("=IMREAL(IMCONJUGATE(COMPLEX(3;4)))", 3.0),
        ("=IMAGINARY(IMCONJUGATE(COMPLEX(3;4)))", -4.0),
        ("=IMREAL(IMDIV(COMPLEX(3;4);COMPLEX(1;2)))", 2.2),
        ("=IMAGINARY(IMDIV(COMPLEX(3;4);COMPLEX(1;2)))", -0.4),
        ("=IMREAL(IMEXP(COMPLEX(1;0)))", 1_f64.exp()),
        ("=IMAGINARY(IMEXP(COMPLEX(1;0)))", 0.0),
        ("=IMREAL(IMLN(COMPLEX(1;0)))", 0.0),
        ("=IMAGINARY(IMLN(COMPLEX(0;1)))", PI / 2.0),
        ("=IMREAL(IMLOG10(COMPLEX(10;0)))", 1.0),
        ("=IMREAL(IMLOG2(COMPLEX(2;0)))", 1.0),
        ("=IMREAL(IMPOWER(COMPLEX(2;0);2))", 4.0),
        ("=IMREAL(IMPRODUCT(COMPLEX(2;3);COMPLEX(4;5)))", -7.0),
        ("=IMAGINARY(IMPRODUCT(COMPLEX(2;3);COMPLEX(4;5)))", 22.0),
        ("=IMREAL(IMSUB(COMPLEX(3;4);COMPLEX(1;2)))", 2.0),
        ("=IMAGINARY(IMSUB(COMPLEX(3;4);COMPLEX(1;2)))", 2.0),
        ("=IMREAL(IMSUM(COMPLEX(3;4);COMPLEX(1;2)))", 4.0),
        ("=IMAGINARY(IMSUM(COMPLEX(3;4);COMPLEX(1;2)))", 6.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error:?}"));
        assert_close(number(&result), expected, source);
    }

    // The trigonometric family uses real inputs here, so the exact real
    // projection is independently calculable from the standard functions.
    for (source, expected) in [
        ("=IMREAL(IMCOS(COMPLEX(1;0)))", 1_f64.cos()),
        ("=IMREAL(IMCOSH(COMPLEX(1;0)))", 1_f64.cosh()),
        ("=IMREAL(IMCOT(COMPLEX(1;0)))", 1_f64.tan().recip()),
        ("=IMREAL(IMCSC(COMPLEX(1;0)))", 1_f64.sin().recip()),
        ("=IMREAL(IMCSCH(COMPLEX(1;0)))", 1_f64.sinh().recip()),
        ("=IMREAL(IMSEC(COMPLEX(1;0)))", 1_f64.cos().recip()),
        ("=IMREAL(IMSIN(COMPLEX(1;0)))", 1_f64.sin()),
        ("=IMREAL(IMSINH(COMPLEX(1;0)))", 1_f64.sinh()),
        ("=IMREAL(IMTAN(COMPLEX(1;0)))", 1_f64.tan()),
        ("=IMREAL(IMSQRT(COMPLEX(1;0)))", 1.0),
        ("=IMAGINARY(IMSQRT(COMPLEX(-1;0)))", 1.0),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error:?}"));
        assert_close(number(&result), expected, source);
    }

    // IMSECH is deliberately checked at a non-real point: the selected
    // profile returns the complex reciprocal of IMCOSH, not only its real
    // projection.
    let x = 1_f64.cosh() * 1_f64.cos();
    let y = 1_f64.sinh() * 1_f64.sin();
    let denominator = x.mul_add(x, y * y);
    assert_close(
        number(
            &evaluate(
                &parse("=IMREAL(IMSECH(COMPLEX(1;1)))"),
                &execution,
                &EvaluationLimits::default(),
            )
            .unwrap_or_else(|error| panic!("IMSECH real projection should evaluate: {error:?}")),
        ),
        x / denominator,
        "IMSECH real projection",
    );
    assert_close(
        number(
            &evaluate(
                &parse("=IMAGINARY(IMSECH(COMPLEX(1;1)))"),
                &execution,
                &EvaluationLimits::default(),
            )
            .unwrap_or_else(|error| {
                panic!("IMSECH imaginary projection should evaluate: {error:?}")
            }),
        ),
        -y / denominator,
        "IMSECH imaginary projection",
    );
}

#[test]
fn complex_function_names_and_suffixes_follow_the_case_profile() {
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-spelling");

    for (source, expected) in [
        ("=imreal(complex(2;3;\"j\"))", 2.0),
        ("=imaginary(complex(2;3;\"i\"))", 3.0),
        ("=IMREAL(COMPLEX(2;3;\"j\"))", 2.0),
        ("=IMAGINARY(IMSUM(COMPLEX(1;2);COMPLEX(3;4)))", 6.0),
    ] {
        assert_real(source, &execution, expected)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error:?}"));
    }

    for source in ["=IMREAL(COMPLEX(2;3;\"I\"))", "=IMREAL(COMPLEX(2;3;\"J\"))"] {
        let expression = parse(source);
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error:?}"));
        assert_eq!(formula_error(&result), ScalarError::Value, "{source:?}");
    }
}

#[test]
fn complex_text_forms_accept_exponents_and_both_imaginary_suffixes() {
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-text");
    for (source, expected) in [
        ("=IMREAL(\"1e2+2e-1i\")", 100.0),
        ("=IMAGINARY(\"1e2+2e-1i\")", 0.2),
        ("=IMREAL(\"-3.5j\")", 0.0),
        ("=IMAGINARY(\"-3.5j\")", -3.5),
        ("=IMAGINARY(\"4i\")", 4.0),
        ("=IMREAL(\"4i\")", 0.0),
    ] {
        assert_real(source, &execution, expected)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error:?}"));
    }

    for source in [
        "=IMREAL(\"1+2k\")",
        "=IMREAL(\"1e+\")",
        "=IMREAL(\"1+2\")",
        "=IMREAL(\"1+2ii\")",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error:?}"));
        assert_eq!(formula_error(&result), ScalarError::Value, "{source:?}");
    }
}

#[test]
fn complex_domains_nonfinite_values_and_extreme_finite_values_are_typed() {
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-domains");
    for source in [
        "=IMREAL(IMLN(COMPLEX(0;0)))",
        "=IMREAL(IMLOG10(COMPLEX(0;0)))",
        "=IMREAL(IMLOG2(COMPLEX(0;0)))",
        "=IMREAL(IMPOWER(COMPLEX(0;0);0))",
        "=IMREAL(IMDIV(COMPLEX(1;0);COMPLEX(0;0)))",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error:?}"));
        assert!(
            matches!(result.value(), ScalarValue::Error(_)),
            "{source:?} returned {:?} instead of a formula error",
            result.value()
        );
    }

    for source in [
        "=IMREAL(IMEXP(COMPLEX(1e308;0)))",
        "=IMAGINARY(COMPLEX(1e309;0))",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error:?}"));
        assert_eq!(formula_error(&result), ScalarError::Number, "{source:?}");
    }

    // These vectors guard scaled complex arithmetic around f64 overflow and
    // underflow boundaries.  The expected values are from the independent
    // decimal oracle in the contract, with a relative tolerance suitable for
    // binary f64 evaluation.
    for (source, expected) in [
        ("=IMREAL(IMLN(COMPLEX(1e308;1e308)))", 709.542_782_232_446),
        (
            "=IMREAL(IMEXP(COMPLEX(710;0.7853981633974483)))",
            1.5796728482882014e308,
        ),
        (
            "=IMAGINARY(IMEXP(COMPLEX(710;0.7853981633974483)))",
            1.5796728482882014e308,
        ),
        (
            "=IMREAL(IMSIN(COMPLEX(0.7853981633974483;710.6)))",
            1.439175797666178e308,
        ),
        (
            "=IMAGINARY(IMCOS(COMPLEX(0.7853981633974483;710.6)))",
            -1.439175797666178e308,
        ),
        (
            "=IMREAL(IMSINH(COMPLEX(710.6;0.7853981633974483)))",
            1.439175797666178e308,
        ),
        (
            "=IMAGINARY(IMCOSH(COMPLEX(710.6;0.7853981633974483)))",
            1.439175797666178e308,
        ),
        (
            "=IMREAL(IMPRODUCT(COMPLEX(1.4e154;5.6e153);COMPLEX(1.4e154;5.6e153)))",
            1.6464e308,
        ),
        (
            "=IMAGINARY(IMPRODUCT(COMPLEX(1.4e154;5.6e153);COMPLEX(1.4e154;5.6e153)))",
            1.568e308,
        ),
        ("=IMAGINARY(IMPRODUCT(COMPLEX(1e308;1e-308);1))", 1e-308),
        ("=IMAGINARY(IMDIV(COMPLEX(1e308;1e-308);1))", 1e-308),
        (
            "=IMAGINARY(IMDIV(COMPLEX(1e308;0);COMPLEX(2;5e-324)))",
            -1.235_164_114_603_116_4e-16,
        ),
        ("=IMAGINARY(IMTAN(COMPLEX(0;1000)))", 1.0),
        ("=IMAGINARY(IMCOT(COMPLEX(0;1000)))", -1.0),
        ("=IMREAL(IMSEC(COMPLEX(0;1000)))", 0.0),
        ("=IMREAL(IMCSCH(COMPLEX(1000;0)))", 0.0),
        (
            "=IMREAL(IMLN(COMPLEX(1.7976931348623157e308;1.7976931348623157e308)))",
            710.1292864836639,
        ),
        (
            "=IMREAL(IMSQRT(COMPLEX(1.7976931348623157e308;1.7976931348623157e308)))",
            1.4730945569055652e154,
        ),
        (
            "=IMAGINARY(IMSQRT(COMPLEX(1.7976931348623157e308;1.7976931348623157e308)))",
            6.1017574412827024e153,
        ),
        ("=IMAGINARY(IMSQRT(COMPLEX(1;1e-300)))", 5e-301),
        ("=IMAGINARY(IMSQRT(COMPLEX(-1;-0)))", 1.0),
        ("=IMREAL(IMDIV(COMPLEX(1e308;1e308);COMPLEX(1;1)))", 1e308),
        ("=IMABS(COMPLEX(1.2e308;1.2e308))", 1.697056274847714e308),
        (
            "=IMREAL(IMDIV(COMPLEX(1e308;1e308);COMPLEX(1e308;1e308)))",
            1.0,
        ),
        (
            "=IMREAL(IMDIV(COMPLEX(1e-308;1e-308);COMPLEX(1e-308;1e-308)))",
            1.0,
        ),
    ] {
        assert_real(source, &execution, expected)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error:?}"));
    }

    for (source, expected) in [
        (
            "=IMREAL(IMSQRT(COMPLEX(1e308;1e308)))",
            1.098_684_113_467_81e154,
        ),
        (
            "=IMAGINARY(IMSQRT(COMPLEX(1e308;1e308)))",
            4.550898605622273e153,
        ),
    ] {
        assert_real(source, &execution, expected)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error:?}"));
    }
}

#[test]
fn complex_arity_empty_products_and_scalar_logicals_follow_the_profile() {
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-arity");
    let expression = parse("=IMREAL(IMSUM())");
    let result = evaluate(&expression, &execution, &EvaluationLimits::default())
        .expect("IMSUM() is the selected additive identity");
    assert_eq!(number(&result), 0.0);

    let expression = parse("=IMPRODUCT()");
    let result = evaluate(&expression, &execution, &EvaluationLimits::default())
        .expect("IMPRODUCT() should produce a formula-level arity error");
    assert_eq!(formula_error(&result), ScalarError::Value);

    for (source, expected) in [
        ("=IMREAL(TRUE())", 1.0),
        ("=IMREAL(FALSE())", 0.0),
        ("=IMARGUMENT(COMPLEX(0;0))", 0.0),
    ] {
        assert_real(source, &execution, expected)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error:?}"));
    }

    for source in [
        "=IMABS()",
        "=IMDIV(COMPLEX(1;0))",
        "=IMSUB(COMPLEX(1;0))",
        "=COMPLEX(1)",
        "=COMPLEX(1;2;\"i\";\"j\")",
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return an arity value: {error:?}"));
        assert_eq!(formula_error(&result), ScalarError::Value, "{source:?}");
    }
}

#[test]
fn complex_errors_are_values_and_if_keeps_unselected_branches_lazy() {
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-errors");
    for (source, expected) in [
        ("=IMREAL(IMSUM(#N/A;#DIV/0!))", ScalarError::NotAvailable),
        ("=IMREAL(IMSUM(#DIV/0!;#N/A))", ScalarError::DivisionByZero),
        (
            "=IMREAL(IMDIV(#N/A;COMPLEX(0;0)))",
            ScalarError::NotAvailable,
        ),
    ] {
        let expression = parse(source);
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error:?}"));
        assert_eq!(formula_error(&result), expected, "{source:?}");
    }

    // The unselected Missing reference must not be resolved by scalar IF.
    let expression = parse("=IF(TRUE();IMREAL(IMSUM(COMPLEX(2;3);COMPLEX(1;0)));[Missing.A1])");
    let result = evaluate(&expression, &execution, &EvaluationLimits::default())
        .expect("the selected complex branch should evaluate without a resolver");
    assert_eq!(number(&result), 3.0);
}

#[derive(Debug)]
enum FixtureCell {
    Empty,
    Logical(bool),
    Number(f64),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct ComplexResolver {
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
}

impl ComplexResolver {
    fn sequence() -> Self {
        Self {
            cells: vec![
                FixtureCell::Number(2.0),
                FixtureCell::Empty,
                FixtureCell::Logical(true),
                FixtureCell::Text("3+4i".to_owned()),
                FixtureCell::Error(ScalarError::NotAvailable),
            ],
            reads: Cell::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }
}

impl evaluation::value::Resolver for ComplexResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> evaluation::EvaluationResult<Option<evaluation::value::SheetExtent>> {
        Ok((sheet == "Main").then_some(evaluation::value::SheetExtent::new(1, self.cells.len())))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> evaluation::EvaluationResult<evaluation::value::CellRead<'a>> {
        self.reads.set(self.reads.get().saturating_add(1));
        if sheet != "Main" || row != 0 {
            return Ok(evaluation::value::CellRead::Empty);
        }
        Ok(match self.cells.get(column) {
            Some(FixtureCell::Empty) | None => evaluation::value::CellRead::Empty,
            Some(FixtureCell::Logical(value)) => evaluation::value::CellRead::Logical(*value),
            Some(FixtureCell::Number(value)) => evaluation::value::CellRead::Number(*value),
            Some(FixtureCell::Text(value)) => evaluation::value::CellRead::Text(value.as_str()),
            Some(FixtureCell::Error(error)) => evaluation::value::CellRead::Error(*error),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> evaluation::EvaluationResult<Option<usize>> {
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> evaluation::EvaluationResult<Option<&str>> {
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> evaluation::EvaluationResult<usize> {
        Ok(1)
    }
}

#[test]
fn complex_reference_sequences_omit_empty_and_logical_but_keep_text_and_errors() {
    let resolver = ComplexResolver::sequence();
    let (_budget, _cancellation, execution) =
        new_execution("ods-formula-complex-reference-sequence");

    let expression = parse("=IMREAL(IMSUM([.A1:.D1]))");
    let result = evaluation::value::evaluate(
        &expression,
        &resolver,
        &evaluation::value::Context::new(
            &execution,
            evaluation::value::Position::new("Main", 0, 0),
        )
        .with_mode(evaluation::value::Mode::Scalar),
        &evaluation::value::Limits::default(),
    )
    .expect("IMSUM should evaluate the reference sequence");
    match result.value() {
        evaluation::value::Value::Number(value) => {
            assert_close(value, 5.0, "reference sequence sum")
        },
        value => panic!("expected Number, got {value:?}"),
    }
    assert_eq!(resolver.reads(), 4);

    let expression = parse("=IMAGINARY(IMSUM([.A1:.D1]~[.A1]))");
    let result = evaluation::value::evaluate(
        &expression,
        &resolver,
        &evaluation::value::Context::new(
            &execution,
            evaluation::value::Position::new("Main", 0, 0),
        )
        .with_mode(evaluation::value::Mode::Scalar),
        &evaluation::value::Limits::default(),
    )
    .expect("IMSUM should preserve duplicate list entries");
    match result.value() {
        evaluation::value::Value::Number(value) => {
            assert_close(value, 4.0, "reference sequence imaginary sum")
        },
        value => panic!("expected Number, got {value:?}"),
    }
    assert_eq!(resolver.reads(), 9);

    let expression = parse("=IMREAL(IMSUM(TRUE()))");
    let result = evaluation::value::evaluate(
        &expression,
        &resolver,
        &evaluation::value::Context::new(
            &execution,
            evaluation::value::Position::new("Main", 0, 0),
        )
        .with_mode(evaluation::value::Mode::Scalar),
        &evaluation::value::Limits::default(),
    )
    .expect("a direct logical scalar converts through Number");
    match result.value() {
        evaluation::value::Value::Number(value) => {
            assert_close(value, 1.0, "direct logical scalar")
        },
        value => panic!("expected Number, got {value:?}"),
    }

    let expression = parse("=IMREAL(IMSUM([.A1:.E1]))");
    let result = evaluation::value::evaluate(
        &expression,
        &resolver,
        &evaluation::value::Context::new(
            &execution,
            evaluation::value::Position::new("Main", 0, 0),
        )
        .with_mode(evaluation::value::Mode::Scalar),
        &evaluation::value::Limits::default(),
    )
    .expect("formula errors should remain values in a sequence");
    match result.value() {
        evaluation::value::Value::Error(error) => {
            assert_eq!(error, ScalarError::NotAvailable)
        },
        value => panic!("expected formula Error, got {value:?}"),
    }
}

#[test]
fn complex_evaluation_honors_cancellation_and_local_text_limits() {
    let expression = parse("=IMREAL(IMSUM(COMPLEX(2;3);COMPLEX(4;5)))");
    let (budget, cancellation, execution) = new_execution("ods-formula-complex-cancel");
    cancellation.cancel();
    let error = evaluate(&expression, &execution, &EvaluationLimits::default())
        .expect_err("cancelled complex evaluation must not publish a result");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Memory), 0);

    let expression = parse("=IMREAL(\"123456789+1i\")");
    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-text-limit");
    let error = evaluate(
        &expression,
        &execution,
        &EvaluationLimits::default().with_max_text_bytes(4),
    )
    .expect_err("complex text should honor the scalar text-byte limit");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong complex text limit error: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn complex_constructor_errors_and_numeric_equality_are_explicit() {
    use evaluation::complex::{Complex, Error};
    assert_eq!(Complex::new(f64::INFINITY, 0.0, 'i'), Err(Error::NonFinite));
    assert_eq!(Complex::new(0.0, f64::NAN, 'i'), Err(Error::NonFinite));
    assert_eq!(Complex::new(1.0, 2.0, 'I'), Err(Error::InvalidSuffix));
    let i = Complex::new(1.0, 2.0, 'i').expect("finite complex");
    let j = Complex::new(1.0, 2.0, 'j').expect("finite complex");
    assert_eq!(i, j, "suffix is formatting metadata");
    assert_eq!(i.to_string(), "1+2i");
    assert_eq!(j.to_string(), "1+2j");
    assert_eq!(
        Complex::new(3.0, 0.0, 'j').expect("real only").to_string(),
        "3"
    );
}

#[test]
fn complex_zero_axis_components_preserve_sign() {
    for (source, negative) in [
        ("=IMAGINARY(IMSIN(COMPLEX(0;-0)))", true),
        ("=IMREAL(IMSINH(COMPLEX(-0;0)))", true),
        ("=IMAGINARY(IMCOS(COMPLEX(0;-0)))", false),
    ] {
        let expression = parse(source);
        let (_budget, _cancellation, execution) = new_execution("complex-signed-zero");
        let result = evaluate(&expression, &execution, &EvaluationLimits::default())
            .expect("finite zero-axis result");
        let actual = number(&result);
        assert_eq!(actual, 0.0, "{source}");
        assert_eq!(actual.is_sign_negative(), negative, "{source}");
    }
}
