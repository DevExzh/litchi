//! Integration coverage for complex values in the worksheet-aware evaluator.
//!
//! The scalar complex family is covered separately.  These cases exercise the
//! public `Value::Complex` representation when arrays are materialized,
//! transformed, retained through lazy branches, folded as reference
//! sequences, and copied into lifetime-free owned storage.

use std::{cell::Cell, num::NonZeroU64, num::NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationFailure, EvaluationResult, ScalarError, UnsupportedKind,
        complex::Complex,
        value::{
            self, CellRead, Context, Evaluated, Limits as ValueLimits, Mode, OwnedValueView,
            Position, Resolver, SheetExtent, Value,
        },
    },
    expression::Expression,
};

fn new_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(1 << 20).expect("nonzero in-flight byte limit"),
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
    resolver: &'a ComplexArrayResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &ValueLimits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_complex(actual: Complex, real: f64, imaginary: f64, suffix: char, label: &str) {
    assert_eq!(actual.real(), real, "{label} real component");
    assert_eq!(actual.imaginary(), imaginary, "{label} imaginary component");
    assert_eq!(actual.suffix(), suffix, "{label} suffix");
}

fn assert_complex_array(
    result: &Evaluated<'_>,
    rows: usize,
    columns: usize,
    expected: &[(f64, f64, char)],
) {
    let array = result.as_array().expect("expected a materialized array");
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (rows, columns)
    );
    assert_eq!(array.len(), expected.len());
    for (index, &(real, imaginary, suffix)) in expected.iter().enumerate() {
        match array.get(index).expect("array cell should be present") {
            Value::Complex(value) => assert_complex(
                value,
                real,
                imaginary,
                suffix,
                &format!("array cell {index}"),
            ),
            other => panic!("expected Complex at array cell {index}, got {other:?}"),
        }
    }
}

#[derive(Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct ComplexArrayResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    missing_metadata_calls: Cell<usize>,
    reject_missing: bool,
}

impl ComplexArrayResolver {
    fn empty() -> Self {
        Self::with_shape(8, 8, false)
    }

    fn with_shape(rows: usize, columns: usize, reject_missing: bool) -> Self {
        let size = rows.checked_mul(columns).expect("fixture shape fits");
        let mut cells = Vec::with_capacity(size);
        cells.resize_with(size, || FixtureCell::Empty);
        Self {
            rows,
            columns,
            cells,
            reads: Cell::new(0),
            missing_metadata_calls: Cell::new(0),
            reject_missing,
        }
    }

    fn sequence() -> Self {
        let mut resolver = Self::with_shape(4, 4, false);
        resolver.set(0, 0, FixtureCell::Text("1+2i".to_owned()));
        resolver.set(0, 1, FixtureCell::Text("3+4i".to_owned()));
        resolver.set(0, 2, FixtureCell::Error(ScalarError::NotAvailable));
        resolver.set(0, 3, FixtureCell::Text("ignored".to_owned()));
        resolver
    }

    fn rejecting_missing() -> Self {
        Self::with_shape(8, 8, true)
    }

    fn set(&mut self, row: usize, column: usize, value: FixtureCell) {
        let index = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column))
            .expect("fixture coordinate fits");
        self.cells[index] = value;
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn missing_metadata_calls(&self) -> usize {
        self.missing_metadata_calls.get()
    }
}

impl Resolver for ComplexArrayResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<SheetExtent>> {
        if sheet == "Missing" {
            self.missing_metadata_calls
                .set(self.missing_metadata_calls.get().saturating_add(1));
            if self.reject_missing {
                return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
            }
            return Ok(None);
        }
        Ok((sheet == "Main").then_some(SheetExtent::new(self.rows, self.columns)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<CellRead<'a>> {
        self.reads.set(self.reads.get().saturating_add(1));
        if sheet != "Main" || row >= self.rows || column >= self.columns {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        let index = row * self.columns + column;
        Ok(match self.cells.get(index) {
            Some(FixtureCell::Empty) | None => CellRead::Empty,
            Some(FixtureCell::Number(value)) => CellRead::Number(*value),
            Some(FixtureCell::Logical(value)) => CellRead::Logical(*value),
            Some(FixtureCell::Text(value)) => CellRead::Text(value.as_str()),
            Some(FixtureCell::Error(error)) => CellRead::Error(*error),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<usize>> {
        if sheet == "Missing" {
            self.missing_metadata_calls
                .set(self.missing_metadata_calls.get().saturating_add(1));
            if self.reject_missing {
                return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
            }
            return Ok(None);
        }
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<&str>> {
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> EvaluationResult<usize> {
        Ok(1)
    }
}

#[test]
fn matrix_complex_arrays_retain_components_and_suffixes() {
    let resolver = ComplexArrayResolver::empty();
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-array-elements");
    let expression = parse("=COMPLEX({1;2|3;4};{5;6|7;8};\"j\")");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("COMPLEX should broadcast numeric arrays elementwise");
    assert_complex_array(
        &result,
        2,
        2,
        &[
            (1.0, 5.0, 'j'),
            (2.0, 6.0, 'j'),
            (3.0, 7.0, 'j'),
            (4.0, 8.0, 'j'),
        ],
    );
}

#[test]
fn nested_lazy_if_and_transpose_retain_complex_elements_without_reading_unused_refs() {
    let resolver = ComplexArrayResolver::rejecting_missing();
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-array-lazy");
    let expression = parse(
        "=IF({TRUE()};IF({TRUE()};TRANSPOSE(COMPLEX({1;2};{3;4};\"j\"));[Missing.A1]);[Missing.B1])",
    );
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("selected nested branch should evaluate");
    assert_complex_array(&result, 2, 1, &[(1.0, 3.0, 'j'), (2.0, 4.0, 'j')]);
    assert_eq!(resolver.reads(), 0);
    assert_eq!(resolver.missing_metadata_calls(), 0);
}

#[test]
fn owned_complex_arrays_survive_input_drop_and_release_storage_on_drop() {
    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-array-owned");
    let owned = {
        let source = String::from("=TRANSPOSE(COMPLEX({1;2};{3;4};\"j\"))");
        let expression = parse(&source);
        let resolver = ComplexArrayResolver::empty();
        let borrowed = evaluate(
            &expression,
            &resolver,
            &execution,
            Mode::Matrix,
            &ValueLimits::default(),
        )
        .expect("complex array should evaluate");
        assert_complex_array(&borrowed, 2, 1, &[(1.0, 3.0, 'j'), (2.0, 4.0, 'j')]);
        let owned = borrowed
            .to_owned(&execution, &ValueLimits::default())
            .expect("complex array should be ownable");
        assert!(owned.reserved_storage_bytes() > 0);
        owned
    };

    drop(execution);
    match owned.value() {
        OwnedValueView::Array(array) => {
            assert_eq!((array.shape().rows(), array.shape().columns()), (2, 1));
            match array.get(0).expect("owned cell 0") {
                OwnedValueView::Complex(value) => assert_complex(value, 1.0, 3.0, 'j', "owned 0"),
                other => panic!("expected owned Complex cell 0, got {other:?}"),
            }
            match array.get(1).expect("owned cell 1") {
                OwnedValueView::Complex(value) => assert_complex(value, 2.0, 4.0, 'j', "owned 1"),
                other => panic!("expected owned Complex cell 1, got {other:?}"),
            }
        },
        other => panic!("expected owned Complex array, got {other:?}"),
    }
    assert!(budget.used(Resource::Memory) > 0);
    drop(owned);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn complex_sequence_folds_reference_lists_in_order_with_duplicates() {
    let resolver = ComplexArrayResolver::sequence();
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-array-sequence");

    let expression = parse("=[.A1]~[.B1]~[.A1]");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("a union should remain a first-class reference list");
    let list = match result.value() {
        Value::ReferenceList(list) => list,
        other => panic!("expected ReferenceList, got {other:?}"),
    };
    assert_eq!(list.len(), 3);
    for (index, column) in [0, 1, 0].into_iter().enumerate() {
        let reference = list.get(index).expect("reference record");
        assert_eq!(reference.areas().len(), 1);
        assert_eq!(reference.areas()[0].starts(), [0, 0, column]);
        assert_eq!(reference.areas()[0].ends(), [1, 1, column + 1]);
    }
    assert_eq!(
        resolver.reads(),
        0,
        "retaining a reference list reads no cells"
    );

    let expression = parse("=IMSUM([.A1]~[.B1]~[.A1])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("IMSUM should consume a reference list");
    match result.value() {
        Value::Complex(value) => assert_complex(value, 5.0, 8.0, 'i', "reference sum"),
        other => panic!("expected complex reference sum, got {other:?}"),
    }

    let expression = parse("=IMPRODUCT([.A1]~[.B1]~[.A1])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("IMPRODUCT should consume duplicate reference records");
    match result.value() {
        Value::Complex(value) => assert_complex(value, -25.0, 0.0, 'i', "reference product"),
        other => panic!("expected complex reference product, got {other:?}"),
    }
    assert_eq!(resolver.reads(), 6);
}

#[test]
fn complex_sequence_ignores_unconvertible_text_but_preserves_first_formula_error() {
    let resolver = ComplexArrayResolver::sequence();
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-array-errors");

    let expression = parse("=IMSUM([.A1]~[.D1]~[.B1])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("IMSUM should ignore unconvertible text");
    match result.value() {
        Value::Complex(value) => assert_complex(value, 4.0, 6.0, 'i', "ignored text sum"),
        other => panic!("expected complex sum, got {other:?}"),
    }

    let expression = parse("=IMSUM([.C1]~[.D1]~[.B1])");
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("formula errors should remain values");
    match result.value() {
        Value::Error(error) => assert_eq!(error, ScalarError::NotAvailable),
        other => panic!("expected first formula error, got {other:?}"),
    }
}

#[test]
fn complex_array_values_have_structural_equality_and_component_safety() {
    let resolver = ComplexArrayResolver::empty();
    let (_budget, _cancellation, execution) = new_execution("ods-formula-complex-array-equality");
    let first_expression = parse("=COMPLEX({1;2};{3;4};\"j\")");
    let second_expression = parse("=COMPLEX({1;2};{3;4};\"j\")");
    let different_expression = parse("=COMPLEX({1;2};{3;5};\"j\")");
    let first = evaluate(
        &first_expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("first complex array");
    let second = evaluate(
        &second_expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("second complex array");
    let different = evaluate(
        &different_expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect("different complex array");
    assert_eq!(first.value(), second.value());
    assert_ne!(first.value(), different.value());

    let first_owned = first
        .to_owned(&execution, &ValueLimits::default())
        .expect("first array ownership");
    let second_owned = second
        .to_owned(&execution, &ValueLimits::default())
        .expect("second array ownership");
    assert_eq!(first_owned.value(), second_owned.value());
}

#[test]
fn complex_array_work_and_cancellation_fail_before_publishing_storage() {
    let expression = parse("=COMPLEX({1;2|3;4};{5;6|7;8};\"j\")");
    let resolver = ComplexArrayResolver::empty();
    let (budget, _cancellation, execution) = new_execution("ods-formula-complex-array-work");
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default().with_max_steps(0),
    )
    .expect_err("zero Work must refuse before array publication");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work
    ));
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation, execution) = new_execution("ods-formula-complex-array-cancel");
    cancellation.cancel();
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Mode::Matrix,
        &ValueLimits::default(),
    )
    .expect_err("pre-cancelled evaluation must refuse");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn single_cell_sequences_omit_logicals_and_empty_but_keep_numbers() {
    let mut resolver = ComplexArrayResolver::with_shape(1, 3, false);
    resolver.set(0, 0, FixtureCell::Logical(true));
    resolver.set(0, 1, FixtureCell::Number(2.0));
    let (_budget, _cancellation, execution) = new_execution("ods-complex-single-cell-sequences");
    for (source, expected) in [
        ("=IMREAL(IMSUM([.A1]))", 0.0),
        ("=IMREAL(IMSUM([.B1]))", 2.0),
        ("=IMREAL(IMPRODUCT([.C1];2))", 2.0),
        ("=IMREAL(IMSUM([.A1:.C1];TRUE()))", 3.0),
    ] {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Mode::Scalar,
            &ValueLimits::default(),
        )
        .expect("sequence evaluation");
        assert_eq!(result.value(), Value::Number(expected), "{source}");
    }
}
