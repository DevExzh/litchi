//! Bounded resource, cancellation, and ownership coverage for OpenFormula
//! complex-number functions.
//!
//! Complex values use the public formula representation selected by the ODS
//! evaluator.  The tests project through `IMREAL` and `IMAGINARY` wherever a
//! numeric assertion is sufficient, so the resource checks do not depend on
//! a particular textual spelling of a complex result.

use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluatedScalar, EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError,
        ScalarValue, evaluate_scalar,
        value::{
            CellRead, Context, Evaluated, Limits as ValueLimits, Mode, OwnedValueView, Position,
            Resolver, SheetExtent, Value, evaluate as evaluate_value,
        },
    },
    expression::Expression,
};

#[derive(Debug)]
enum FixtureCell {
    Empty,
    Logical(bool),
    Text(String),
}

#[derive(Debug)]
struct ComplexResolver {
    cells: Vec<FixtureCell>,
    columns: usize,
    reads: AtomicUsize,
    cancel_after_read: Option<CancellationSource>,
}

impl ComplexResolver {
    fn empty() -> Self {
        Self::with_cells(Vec::new())
    }

    fn with_cells(cells: Vec<FixtureCell>) -> Self {
        Self {
            columns: cells.len().max(1),
            cells,
            reads: AtomicUsize::new(0),
            cancel_after_read: None,
        }
    }

    fn sequence() -> Self {
        Self::with_cells(vec![
            FixtureCell::Text("1+2i".to_owned()),
            FixtureCell::Empty,
            FixtureCell::Logical(true),
            FixtureCell::Text("malformed-complex".to_owned()),
            FixtureCell::Text("3-4i".to_owned()),
        ])
    }

    fn with_cancel_after_read(mut self, cancellation: &CancellationSource) -> Self {
        self.cancel_after_read = Some(cancellation.clone());
        self
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Acquire)
    }
}

impl Resolver for ComplexResolver {
    fn sheet_extent(
        &self,
        _sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok(Some(SheetExtent::new(8, self.columns)))
    }

    fn read_cell<'a>(
        &'a self,
        _sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.fetch_add(1, Ordering::AcqRel);
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        let index = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column));
        Ok(match index.and_then(|index| self.cells.get(index)) {
            Some(FixtureCell::Empty) | None => CellRead::Empty,
            Some(FixtureCell::Logical(value)) => CellRead::Logical(*value),
            Some(FixtureCell::Text(value)) => CellRead::Text(value.as_str()),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(1)
    }
}

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

fn evaluate_scalar_at<'a>(
    expression: &'a Expression,
    execution: &ExecutionContext,
    limits: &EvaluationLimits,
) -> Result<EvaluatedScalar<'a>, EvaluationFailure> {
    evaluate_scalar(expression, &EvaluationContext::new(execution), limits)
}

fn evaluate_value_at<'a>(
    expression: &'a Expression,
    resolver: &'a ComplexResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &ValueLimits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    evaluate_value(expression, resolver, &context, limits)
}

fn scalar_number(result: &EvaluatedScalar<'_>) -> f64 {
    match result.value() {
        ScalarValue::Number(value) => *value,
        value => panic!("expected Number, got {value:?}"),
    }
}

fn value_number(result: &Evaluated<'_>) -> f64 {
    match result.value() {
        Value::Number(value) => value,
        value => panic!("expected value Number, got {value:?}"),
    }
}

#[test]
fn complex_real_and_imaginary_projections_are_finite_numbers() {
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-projections");
    for (source, expected) in [
        ("=IMREAL(COMPLEX(3;4))", 3.0),
        ("=IMAGINARY(COMPLEX(3;4))", 4.0),
        ("=IMREAL(IMSUM(COMPLEX(1;2);COMPLEX(3;4)))", 4.0),
        ("=IMAGINARY(IMSUM(COMPLEX(1;2);COMPLEX(3;4)))", 6.0),
    ] {
        let expression = parse(source);
        let result = evaluate_scalar_at(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_eq!(scalar_number(&result), expected, "{source:?}");
        drop(result);
    }
}

#[test]
fn long_malformed_complex_text_charges_work_before_returning_a_formula_error() {
    let malformed = "x".repeat(8_192);
    let source = format!("=IMREAL(\"{malformed}\")");
    let expression = parse(&source);
    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-long-text");
    let error = evaluate_scalar_at(
        &expression,
        &execution,
        &EvaluationLimits::default().with_max_steps(128),
    )
    .expect_err("long malformed complex text must exceed the small Work limit");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong long-text failure: {error:?}"
    );
    assert!(budget.used(Resource::Work) > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn scalar_complex_zero_work_and_storage_limits_are_typed_and_atomic() {
    let expression = parse("=IMREAL(COMPLEX(1;2))");

    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-scalar-work");
    let error = evaluate_scalar_at(
        &expression,
        &execution,
        &EvaluationLimits::default().with_max_steps(0),
    )
    .expect_err("zero scalar Work must refuse complex evaluation");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong scalar Work failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-scalar-storage");
    let error = evaluate_scalar_at(
        &expression,
        &execution,
        &EvaluationLimits::default().with_max_storage_bytes(0),
    )
    .expect_err("zero scalar storage must refuse complex evaluation");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong scalar storage failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn value_complex_zero_work_and_storage_limits_are_typed_and_atomic() {
    let expression = parse("=IMAGINARY(COMPLEX(1;2))");
    let resolver = ComplexResolver::empty();

    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-value-work");
    let limits = ValueLimits::default().with_max_steps(0);
    let error = evaluate_value_at(&expression, &resolver, &execution, Mode::Scalar, &limits)
        .expect_err("zero value Work must refuse complex evaluation");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong value Work failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-value-storage");
    let limits = ValueLimits::default().with_max_storage_bytes(0);
    let error = evaluate_value_at(&expression, &resolver, &execution, Mode::Scalar, &limits)
        .expect_err("zero value storage must refuse complex evaluation");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong value storage failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn complex_sequence_references_omit_empty_logical_and_malformed_text() {
    let resolver = ComplexResolver::sequence();
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-sequence");
    let expression = parse("=IMREAL(IMSUM([.A1:.E1]))");
    let result = evaluate_value_at(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default(),
    )
    .expect("IMSUM should ignore non-complex sequence cells");
    assert_eq!(value_number(&result), 4.0);
    assert_eq!(resolver.reads(), 5, "the complete sequence must be read");
    drop(result);

    let expression = parse("=IMAGINARY(IMSUM([.A1:.E1]))");
    let result = evaluate_value_at(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default(),
    )
    .expect("IMSUM should retain the imaginary component of valid cells");
    assert_eq!(value_number(&result), -2.0);
    assert_eq!(
        resolver.reads(),
        10,
        "both projections must charge every read"
    );
    drop(result);
}

#[test]
fn complex_sequence_reads_charge_work_and_mid_read_cancellation_is_atomic() {
    let expression = parse("=IMAGINARY(IMSUM([.A1:.E1]))");

    let (budget, cancellation, execution) = new_execution("ods-formula-complex-sequence-cancel");
    let resolver = ComplexResolver::sequence().with_cancel_after_read(&cancellation);
    let error = evaluate_value_at(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default(),
    )
    .expect_err("provider cancellation must fence complex sequence publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert!(budget.used(Resource::Work) > 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-sequence-work");
    let resolver = ComplexResolver::sequence();
    let error = evaluate_value_at(
        &expression,
        &resolver,
        &execution,
        Mode::Scalar,
        &ValueLimits::default().with_max_steps(8),
    )
    .expect_err("a short Work allowance must stop sequence traversal");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong sequence Work failure: {error:?}"
    );
    assert!(resolver.reads() < 5);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn owned_complex_text_survives_source_and_resolver_drop_until_result_drop() {
    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-owned-result");
    let owned = {
        let source = String::from("=IMSUM(\"1+2i\";\"3+4i\")&\"\"");
        let expression = parse(&source);
        let resolver = ComplexResolver::empty();
        let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Scalar);
        let borrowed = evaluate_value(&expression, &resolver, &context, &ValueLimits::default())
            .expect("complex text result should evaluate");
        match borrowed.value() {
            Value::Text(text) => assert!(!text.is_empty()),
            value => panic!("expected borrowed complex text, got {value:?}"),
        }
        let owned = borrowed
            .to_owned(&execution, &ValueLimits::default())
            .expect("complex text should be ownable");
        match owned.value() {
            OwnedValueView::Text(text) => assert!(!text.is_empty()),
            value => panic!("expected owned complex text, got {value:?}"),
        }
        assert!(owned.reserved_storage_bytes() > 0);
        owned
    };

    drop(execution);
    assert!(budget.used(Resource::Memory) > 0);
    match owned.value() {
        OwnedValueView::Text(text) => assert!(!text.is_empty()),
        value => panic!("owned result changed after source drop: {value:?}"),
    }
    drop(owned);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn non_finite_complex_inputs_remain_formula_number_errors() {
    for source in ["=IMREAL(COMPLEX(1e309;0))", "=IMAGINARY(COMPLEX(0;1e309))"] {
        let expression = parse(source);
        let (_budget, _cancellation, execution) =
            new_execution("ods-formula-complex-number-domain");
        let result = evaluate_scalar_at(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
        assert!(
            matches!(result.value(), ScalarValue::Error(ScalarError::Number)),
            "{source:?} returned {:?}",
            result.value()
        );
    }
}

#[test]
fn owned_numeric_complex_retains_components_without_text_storage() {
    let (budget, _cancellation, execution) = new_execution("ods-complex-owned-numeric");
    let owned = {
        let expression = parse("=COMPLEX(3;4;\"j\")");
        let resolver = ComplexResolver::empty();
        let result = evaluate_value_at(
            &expression,
            &resolver,
            &execution,
            Mode::Scalar,
            &ValueLimits::default(),
        )
        .expect("numeric complex result");
        match result.value() {
            Value::Complex(value) => {
                assert_eq!(value.real(), 3.0);
                assert_eq!(value.imaginary(), 4.0);
                assert_eq!(value.suffix(), 'j');
            },
            other => panic!("expected Complex: {other:?}"),
        }
        result
            .to_owned(&execution, &ValueLimits::default())
            .expect("own numeric complex")
    };
    assert_eq!(owned.reserved_storage_bytes(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    drop(execution);
    match owned.value() {
        OwnedValueView::Complex(value) => {
            assert_eq!(value.real(), 3.0);
            assert_eq!(value.imaginary(), 4.0);
            assert_eq!(value.suffix(), 'j');
        },
        other => panic!("expected owned Complex: {other:?}"),
    }
    drop(owned);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn scalar_product_preserves_an_earlier_numeric_failure_over_later_bad_text() {
    let cases = [
        (
            "=IMPRODUCT(COMPLEX(1e308;0);COMPLEX(1e308;0);\"bad\")",
            ScalarError::Number,
        ),
        (
            "=IMPRODUCT(#N/A;COMPLEX(1e308;0);\"bad\")",
            ScalarError::NotAvailable,
        ),
        (
            "=IMPRODUCT(COMPLEX(1e308;0);COMPLEX(1e308;0);#DIV/0!;\"bad\")",
            ScalarError::DivisionByZero,
        ),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let (_budget, _cancellation, execution) = new_execution("ods-complex-product-errors");
        let result = evaluate_scalar_at(&expression, &execution, &EvaluationLimits::default())
            .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
        match result.value() {
            ScalarValue::Error(actual) => assert_eq!(*actual, expected, "{source:?}"),
            value => panic!("{source:?} returned {value:?}, expected {expected:?}"),
        }
    }
}

#[test]
fn value_product_preserves_numeric_failure_and_original_formula_error_precedence() {
    let resolver = ComplexResolver::empty();
    let (_budget, _cancellation, execution) = new_execution("ods-complex-value-product-errors");

    let expression = parse("=IMPRODUCT({1e308;1e308};\"bad\")");
    let result = evaluate_value_at(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("value product should return its formula error as a value");
    match result.value() {
        Value::Error(error) => assert_eq!(error, ScalarError::Number),
        value => panic!("expected Number formula error, got {value:?}"),
    }

    let expression = parse("=IMPRODUCT({#N/A;1e308};\"bad\")");
    let result = evaluate_value_at(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("original formula errors must remain values");
    match result.value() {
        Value::Error(error) => assert_eq!(error, ScalarError::NotAvailable),
        value => panic!("expected original formula error, got {value:?}"),
    }
}

#[test]
fn scalar_product_has_a_bounded_work_cost_for_each_aggregate_argument() {
    const ARGUMENTS: usize = 96;
    let mut source = String::from("=IMREAL(IMPRODUCT(");
    for index in 0..ARGUMENTS {
        if index != 0 {
            source.push(';');
        }
        source.push_str("COMPLEX(1;0)");
    }
    source.push_str("))");
    let expression = parse(&source);

    let (baseline_budget, _cancellation, baseline_execution) =
        new_execution("ods-complex-product-work-baseline");
    let result = evaluate_scalar_at(
        &expression,
        &baseline_execution,
        &EvaluationLimits::default(),
    )
    .expect("the bounded aggregate should evaluate with default limits");
    match result.value() {
        ScalarValue::Number(value) => assert_eq!(*value, 1.0),
        value => panic!("expected product identity projection, got {value:?}"),
    }
    let full_work = baseline_budget.used(Resource::Work);
    assert!(
        full_work > ARGUMENTS as u64,
        "aggregate work must grow with its argument count: {full_work}"
    );

    let (budget, _cancellation, execution) = new_execution("ods-complex-product-work-limit");
    let error = evaluate_scalar_at(
        &expression,
        &execution,
        &EvaluationLimits::default().with_max_steps(full_work.saturating_sub(1)),
    )
    .expect_err("one less than the measured aggregate work must refuse");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong aggregate work failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}
