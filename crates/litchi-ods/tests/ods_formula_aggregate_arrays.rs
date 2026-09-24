//! Independent array/reference coverage for the OpenFormula 1.4 numeric
//! reduction functions.  These tests exercise the ForceArray and
//! NumberSequence conversion rules separately from scalar evaluation.

use std::sync::{
    Mutex,
    atomic::{AtomicUsize, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits, Mode, OwnedValueView, Position, Resolver,
        SheetExtent, Value,
    },
    evaluation::{EvaluationFailure, ScalarError, UnsupportedKind},
    expression::Expression,
};

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct AggregateResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: AtomicUsize,
    read_order: Mutex<Vec<(String, usize, usize)>>,
    cancel_after_read: Option<CancellationSource>,
}

impl AggregateResolver {
    fn blank(rows: usize, columns: usize) -> Self {
        Self {
            rows,
            columns,
            cells: (0..rows * columns).map(|_| FixtureCell::Empty).collect(),
            reads: AtomicUsize::new(0),
            read_order: Mutex::new(Vec::new()),
            cancel_after_read: None,
        }
    }

    fn numeric_fixture() -> Self {
        let mut resolver = Self::blank(16, 8);
        resolver.set(0, 0, FixtureCell::Number(1.0));
        resolver.set(1, 0, FixtureCell::Number(2.0));
        resolver.set(2, 0, FixtureCell::Text("ignored".to_owned()));
        resolver.set(3, 0, FixtureCell::Empty);
        resolver.set(4, 0, FixtureCell::Logical(true));
        resolver.set(5, 0, FixtureCell::Error(ScalarError::NotAvailable));

        resolver.set(0, 1, FixtureCell::Number(5.0));
        resolver.set(1, 1, FixtureCell::Number(6.0));
        resolver.set(2, 1, FixtureCell::Number(7.0));
        resolver.set(3, 1, FixtureCell::Number(8.0));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: FixtureCell) {
        let index = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column))
            .expect("fixture coordinate fits");
        self.cells[index] = value;
    }

    fn fill_column(&mut self, column: usize, value: f64) {
        for row in 0..self.rows {
            self.set(row, column, FixtureCell::Number(value));
        }
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Acquire)
    }

    fn read_order(&self) -> Vec<(String, usize, usize)> {
        self.read_order
            .lock()
            .expect("fixture read-order lock")
            .clone()
    }

    fn cancel_after_read(&mut self, cancellation: &CancellationSource) {
        self.cancel_after_read = Some(cancellation.clone());
    }
}

impl Resolver for AggregateResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((matches!(sheet, "Main" | "Data" | "Archive"))
            .then_some(SheetExtent::new(self.rows, self.columns)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.fetch_add(1, Ordering::AcqRel);
        self.read_order
            .lock()
            .expect("fixture read-order lock")
            .push((sheet.to_owned(), row, column));
        if let Some(cancellation) = &self.cancel_after_read {
            cancellation.cancel();
        }
        if !matches!(sheet, "Main" | "Data" | "Archive")
            || row >= self.rows
            || column >= self.columns
        {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        // The three planes make a single 3-D Reference observable without
        // changing the ordinary Main-plane fixtures.  Only A1 is populated
        // on the additional planes; all other cells remain Empty.
        if sheet == "Data" && row == 0 && column == 0 {
            return Ok(CellRead::Number(2.0));
        }
        if sheet == "Archive" && row == 0 && column == 0 {
            return Ok(CellRead::Number(3.0));
        }
        if sheet != "Main" {
            return Ok(CellRead::Empty);
        }
        Ok(match &self.cells[row * self.columns + column] {
            FixtureCell::Empty => CellRead::Empty,
            FixtureCell::Number(value) => CellRead::Number(*value),
            FixtureCell::Logical(value) => CellRead::Logical(*value),
            FixtureCell::Text(value) => CellRead::Text(value.as_str()),
            FixtureCell::Error(error) => CellRead::Error(*error),
            FixtureCell::Unsupported => CellRead::Unsupported,
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok(match sheet {
            "Main" => Some(0),
            "Data" => Some(1),
            "Archive" => Some(2),
            _ => None,
        })
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok(match index {
            0 => Some("Main"),
            1 => Some("Data"),
            2 => Some("Archive"),
            _ => None,
        })
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(3)
    }
}

fn execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).expect("one worker"),
        std::num::NonZeroUsize::new(1).expect("one in-flight task"),
        std::num::NonZeroU64::new(1024 * 1024).expect("one in-flight MiB"),
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
    resolver: &'a AggregateResolver,
    execution: &ExecutionContext,
    position: Position<'a>,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn evaluate_source<'a>(
    source: &'a str,
    resolver: &'a AggregateResolver,
    execution: &ExecutionContext,
    mode: Mode,
    limits: &Limits,
) -> Result<AggregateResult, EvaluationFailure> {
    let expression = parse(source);
    let result = evaluate(
        &expression,
        resolver,
        execution,
        Position::new("Main", 0, 0),
        mode,
        limits,
    )?;
    Ok(match result.value() {
        Value::Number(value) => AggregateResult::Number(value),
        Value::Error(error) => AggregateResult::Error(error),
        _ => AggregateResult::Other,
    })
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum AggregateResult {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn aggregate_result(result: &Evaluated<'_>) -> AggregateResult {
    match result.value() {
        Value::Number(value) => AggregateResult::Number(value),
        Value::Error(error) => AggregateResult::Error(error),
        _ => AggregateResult::Other,
    }
}

fn assert_number(result: &AggregateResult, expected: f64, source: &str) {
    let AggregateResult::Number(actual) = *result else {
        panic!("{source:?}: expected Number({expected}), got {result:?}");
    };
    if expected == 0.0 {
        assert_eq!(actual.to_bits(), expected.to_bits(), "{source:?}");
    } else {
        let tolerance = expected.abs().max(actual.abs()) * (64.0 * f64::EPSILON);
        assert!(
            (actual - expected).abs() <= tolerance,
            "{source:?}: {actual:?} != {expected:?} (tol {tolerance:e})"
        );
    }
}

fn assert_bits(result: &AggregateResult, expected: u64, source: &str) {
    let AggregateResult::Number(actual) = *result else {
        panic!("{source:?}: expected a numeric golden, got {result:?}");
    };
    assert_eq!(actual.to_bits(), expected, "{source:?}: {actual:?}");
}

fn assert_bits_within_ulp(result: &AggregateResult, expected: u64, max_ulps: u64, source: &str) {
    let AggregateResult::Number(actual) = *result else {
        panic!("{source:?}: expected a numeric golden, got {result:?}");
    };
    let actual = actual.to_bits();
    let distance = actual.abs_diff(expected);
    assert!(
        distance <= max_ulps,
        "{source:?}: bits {actual:016x} differ from {expected:016x} by {distance} ulps"
    );
}

fn assert_formula_error(result: &AggregateResult, expected: ScalarError, source: &str) {
    match result {
        AggregateResult::Error(actual) => assert_eq!(*actual, expected, "{source:?}"),
        value => panic!("{source:?}: expected Error({expected:?}), got {value:?}"),
    }
}

#[test]
fn every_aggregate_function_reduces_literal_matrices_from_the_definition() {
    let resolver = AggregateResolver::blank(4, 4);
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-array-family");
    let limits = Limits::default();
    for (source, expected) in [
        ("=SUM({1;2|3;4})", 10.0),
        ("=PRODUCT({1;2|3;4})", 24.0),
        ("=SUMSQ({1;2|3;4})", 30.0),
        ("=SUMPRODUCT({1;2|3;4})", 10.0),
        ("=SUMPRODUCT({1;2|3;4};{5;6|7;8})", 70.0),
        ("=SUMX2MY2({1;2|3;4};{5;6|7;8})", -144.0),
        ("=SUMX2PY2({1;2|3;4};{5;6|7;8})", 204.0),
        ("=SUMXMY2({1;2|3;4};{5;6|7;8})", 64.0),
    ] {
        let result = evaluate_source(source, &resolver, &execution, Mode::Matrix, &limits)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, source);
    }
    assert_eq!(
        resolver.reads(),
        0,
        "literal arrays must not read a worksheet"
    );

    for (source, expected) in [("=SUM()", 0.0), ("=SUMSQ()", 0.0)] {
        let result = evaluate_source(source, &resolver, &execution, Mode::Matrix, &limits)
            .unwrap_or_else(|error| panic!("{source:?} should use its additive identity: {error}"));
        assert_number(&result, expected, source);
    }
    let result = evaluate_source("=PRODUCT()", &resolver, &execution, Mode::Matrix, &limits)
        .expect("PRODUCT's strict variadic arity is a formula value");
    assert_formula_error(&result, ScalarError::Value, "=PRODUCT()");
    let result = evaluate_source(
        "=SUMPRODUCT()",
        &resolver,
        &execution,
        Mode::Matrix,
        &limits,
    )
    .expect("SUMPRODUCT's strict variadic arity is a formula value");
    assert_formula_error(&result, ScalarError::Value, "=SUMPRODUCT()");
}

#[test]
fn paired_matrix_functions_require_equal_nonempty_shapes() {
    let resolver = AggregateResolver::blank(2, 2);
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-shape-errors");
    for function in ["SUMPRODUCT", "SUMX2MY2", "SUMX2PY2", "SUMXMY2"] {
        for source in [
            format!("={function}({{1;2|3;4}};{{1;2}})"),
            format!("={function}({{1;2}};{{1;2|3;4}})"),
        ] {
            let result = evaluate_source(
                &source,
                &resolver,
                &execution,
                Mode::Matrix,
                &Limits::default(),
            )
            .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
            assert_formula_error(&result, ScalarError::Value, &source);
        }
    }

    // A direct formula Error has precedence over a later shape violation.
    for source in [
        "=SUMX2MY2({#N/A};{1;2})",
        "=SUMX2PY2({1;2};{#N/A})",
        "=SUMXMY2({#N/A};{1;2})",
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
        assert_formula_error(&result, ScalarError::NotAvailable, source);
    }

    // Direct formula Errors are observed before fixed-arity validation.  The
    // error is therefore retained even when the pair call has one argument,
    // or has an extra argument after the error-bearing pair.
    for (source, expected) in [
        ("=SUMX2MY2(#N/A)", ScalarError::NotAvailable),
        ("=SUMX2PY2(#DIV/0!;1;2)", ScalarError::DivisionByZero),
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
        assert_formula_error(&result, expected, source);
    }
}

#[test]
fn matrix_formula_error_precedence_follows_argument_then_row_order() {
    let resolver = AggregateResolver::blank(2, 2);
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-matrix-error-order");
    for function in ["SUMPRODUCT", "SUMX2MY2", "SUMX2PY2", "SUMXMY2"] {
        let source = format!("={function}({{1;#DIV/0!}};{{#N/A;1}})");
        let result = evaluate_source(
            &source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
        assert_formula_error(&result, ScalarError::DivisionByZero, &source);
    }
}

#[test]
fn array_cells_convert_numbers_and_retain_element_errors() {
    let resolver = AggregateResolver::blank(2, 2);
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-array-coercion");
    for (source, expected) in [
        ("=SUM({\"1\";TRUE()|FALSE();0})", 2.0),
        ("=PRODUCT({\"2\";TRUE()|3;4})", 24.0),
        ("=SUMSQ({\"2\";TRUE()|3;4})", 30.0),
        ("=SUMPRODUCT({\"2\";TRUE()|3;4};{1;2|3;4})", 29.0),
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, source);
    }
    for source in [
        "=SUM({1;#N/A|3;4})",
        "=PRODUCT({1;#N/A|3;4})",
        "=SUMSQ({1;#N/A|3;4})",
        "=SUMPRODUCT({1;#N/A|3;4};{1;2|3;4})",
        "=SUM({\"bad\";#N/A})",
        "=PRODUCT({\"bad\";#N/A})",
        "=SUMSQ({\"bad\";#N/A})",
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should produce a formula error: {error}"));
        assert_formula_error(&result, ScalarError::NotAvailable, source);
    }
    let malformed = evaluate_source(
        "=SUM({\"bad\"})",
        &resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("malformed inline array text should be a formula value");
    assert_formula_error(&malformed, ScalarError::Value, "=SUM({\"bad\"})");
}

#[test]
fn references_filter_text_empty_and_distinguished_logical_cells_for_sequences() {
    let resolver = AggregateResolver::numeric_fixture();
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-reference-sequence");
    for (source, expected) in [
        ("=SUM([.A1:.A5])", 3.0),
        ("=PRODUCT([.A1:.A5])", 2.0),
        ("=SUMSQ([.A1:.A5])", 5.0),
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, source);
    }
    for source in [
        "=SUM([.A1:.A6])",
        "=PRODUCT([.A1:.A6])",
        "=SUMSQ([.A1:.A6])",
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should propagate an error: {error}"));
        assert_formula_error(&result, ScalarError::NotAvailable, source);
    }

    // A syntactically present NumberSequence may contain no Number cells at
    // all after the range conversion.  The sequence aggregate identities are
    // still defined for that selected-empty case.
    for (source, expected) in [
        ("=SUM([.A3:.A4])", 0.0),
        ("=PRODUCT([.A3:.A4])", 1.0),
        ("=SUMSQ([.A3:.A4])", 0.0),
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should use its empty identity: {error}"));
        assert_number(&result, expected, source);
    }

    // A single 3-D Reference is still one NumberSequence, even though the
    // resolver expands its sheet axis into several physical areas.  SUMSQ's
    // ReferenceList rejection must inspect the `is_list` marker, not merely
    // count the expanded areas.
    let three_d_resolver = AggregateResolver::numeric_fixture();
    let result = evaluate_source(
        "=SUMSQ([Main.A1:Archive.A1])",
        &three_d_resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a single 3-D Reference is a legal SUMSQ NumberSequence");
    assert_number(&result, 1.0 + 4.0 + 9.0, "=SUMSQ([Main.A1:Archive.A1])");
    assert_eq!(three_d_resolver.reads(), 3);

    let mut unsupported_resolver = AggregateResolver::blank(1, 1);
    unsupported_resolver.set(0, 0, FixtureCell::Unsupported);
    let error = evaluate_source(
        "=SUM([.A1])",
        &unsupported_resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("unsupported provider cells must remain typed evaluator failures");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
}

#[test]
fn paired_references_reduce_in_corresponding_cell_order() {
    let resolver = AggregateResolver::numeric_fixture();
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-reference-pairs");
    for (source, expected) in [
        ("=SUMPRODUCT([.A1:.A2];[.B1:.B2])", 17.0),
        ("=SUMX2MY2([.A1:.A2];[.B1:.B2])", -56.0),
        ("=SUMX2PY2([.A1:.A2];[.B1:.B2])", 66.0),
        ("=SUMXMY2([.A1:.A2];[.B1:.B2])", 32.0),
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, source);
    }
    assert_eq!(resolver.reads(), 16, "each paired reference is read once");

    let ordered_resolver = AggregateResolver::numeric_fixture();
    let ordered = evaluate_source(
        "=SUMPRODUCT([.A1:.A2];[.B1:.B2])",
        &ordered_resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("paired references should be traversable in lockstep");
    assert_number(&ordered, 17.0, "=SUMPRODUCT([.A1:.A2];[.B1:.B2])");
    assert_eq!(
        ordered_resolver.read_order(),
        vec![
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 0, 1),
            ("Main".to_owned(), 1, 0),
            ("Main".to_owned(), 1, 1),
        ]
    );

    for function in ["SUMPRODUCT", "SUMX2MY2", "SUMX2PY2", "SUMXMY2"] {
        let source = format!("={function}([.A1]~[.B1];[.A1])");
        let result = evaluate_source(
            &source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
        assert_formula_error(&result, ScalarError::Value, &source);
    }

    let list_resolver = AggregateResolver::numeric_fixture();
    let source = "=SUMPRODUCT([.A1]~[.B1];[.A1])";
    let result = evaluate_source(
        source,
        &list_resolver,
        &execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a ForceArray ReferenceList rejection is a formula value");
    assert_formula_error(&result, ScalarError::Value, source);
    assert_eq!(
        list_resolver.reads(),
        0,
        "matrix list rejection is read-free"
    );
}

#[test]
fn reference_lists_preserve_area_order_duplicates_and_stream_without_materializing() {
    let resolver = AggregateResolver::numeric_fixture();
    let (_budget, _cancellation, list_execution) =
        execution("ods-formula-aggregate-reference-list");
    let source = "=SUM([.A1:.A2]~[.B1:.B2]~[.A1])";
    let result = evaluate_source(
        source,
        &resolver,
        &list_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a NumberSequenceList should consume ordered reference areas");
    assert_number(&result, 1.0 + 2.0 + 5.0 + 6.0 + 1.0, source);
    assert_eq!(
        resolver.read_order(),
        vec![
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 1, 0),
            ("Main".to_owned(), 0, 1),
            ("Main".to_owned(), 1, 1),
            ("Main".to_owned(), 0, 0),
        ],
        "reference-list areas and duplicate cells retain source order"
    );

    let product_resolver = AggregateResolver::numeric_fixture();
    let product_result = evaluate_source(
        "=PRODUCT([.A1:.A2]~[.B1:.B2]~[.A1])",
        &product_resolver,
        &list_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("PRODUCT should consume the same ordered ReferenceList");
    assert_number(&product_result, 1.0 * 2.0 * 5.0 * 6.0 * 1.0, "PRODUCT list");
    assert_eq!(product_resolver.reads(), 5);

    // SUMSQ is a NumberSequence function.  A true ReferenceList is not a
    // rectangular sequence argument, so it must become a formula Value error
    // without probing any of its member areas.
    let before = resolver.reads();
    let result = evaluate_source(
        "=SUMSQ([.A1]~[.B1])",
        &resolver,
        &list_execution,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a ReferenceList rejection is represented as a formula value");
    assert_formula_error(&result, ScalarError::Value, "=SUMSQ([.A1]~[.B1])");
    assert_eq!(
        resolver.reads(),
        before,
        "ReferenceList rejection is read-free"
    );

    let mut large = AggregateResolver::blank(2_048, 1);
    large.fill_column(0, 1.0);
    let (budget, _cancellation, stream_execution) =
        execution("ods-formula-aggregate-reference-stream");
    let source = "=SUM([.A1:.A2048])";
    let result = evaluate_source(
        source,
        &large,
        &stream_execution,
        Mode::Matrix,
        &Limits::default().with_max_reference_cells(2_048),
    )
    .expect("the exact reference-cell boundary should stream");
    assert_number(&result, 2_048.0, source);
    assert_eq!(large.reads(), 2_048);
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "scalar reduction retains no array"
    );

    let mut refused = AggregateResolver::blank(2_048, 1);
    refused.fill_column(0, 1.0);
    let (_budget, _cancellation, under_execution) =
        execution("ods-formula-aggregate-reference-under");
    let error = evaluate_source(
        source,
        &refused,
        &under_execution,
        Mode::Matrix,
        &Limits::default().with_max_reference_cells(2_047),
    )
    .expect_err("one cell below the reference admission boundary must refuse");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference-cap error: {error:?}"
    );
    assert_eq!(
        refused.reads(),
        0,
        "geometry is rejected before provider reads"
    );
}

#[test]
fn aggregate_reductions_are_invariant_under_scalar_projection_and_lazy_branches() {
    let resolver = AggregateResolver::numeric_fixture();
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-projection");
    for source in [
        "=SUM([.A1:.A2])",
        "=PRODUCT([.A1:.A2])",
        "=SUMSQ([.A1:.A2])",
    ] {
        let expression = parse(source);
        let matrix = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} matrix evaluation: {error}"));
        let scalar = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 1, 0),
            Mode::Scalar,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} scalar evaluation: {error}"));
        assert_eq!(matrix.value(), scalar.value(), "{source:?}");
    }

    let lazy_resolver = AggregateResolver::numeric_fixture();
    for (source, expected) in [
        ("=IF(TRUE();SUM({1;2});SUM([Missing.A1:.Z100]))", 3.0),
        ("=IF(FALSE();SUM([Missing.A1:.Z100]);SUM({1;2}))", 3.0),
        ("=IFERROR(SUM({1;#N/A});SUM({3;4}))", 7.0),
        ("=SUM(IF(TRUE();{1;2};[Missing.A1:.A2]))", 3.0),
        ("=SUM(IF(FALSE();[Missing.A1:.A2];{1;2}))", 3.0),
    ] {
        let result = evaluate_source(
            source,
            &lazy_resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate lazily: {error}"));
        assert_number(&result, expected, source);
    }
    assert_eq!(
        lazy_resolver.reads(),
        0,
        "all unselected missing branches stay inert"
    );
}

#[test]
fn force_array_pair_arguments_survive_scalar_mode_projection() {
    let mut resolver = AggregateResolver::blank(4, 2);
    resolver.set(0, 0, FixtureCell::Number(1.0));
    resolver.set(1, 0, FixtureCell::Number(2.0));
    resolver.set(2, 0, FixtureCell::Number(3.0));
    resolver.set(0, 1, FixtureCell::Number(5.0));
    resolver.set(1, 1, FixtureCell::Number(6.0));
    resolver.set(2, 1, FixtureCell::Number(7.0));
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-force-array");

    // The arithmetic expression on the first reference is array-valued.  The
    // ForceArray parameters of the pair reductions must retain all three
    // cells even when the enclosing formula is evaluated in scalar mode.  A
    // projected first cell would produce 10, -21, 29, and 9 respectively.
    for (source, expected) in [
        ("=SUMPRODUCT([.A1:.A3]+1;[.B1:.B3])", 56.0),
        ("=SUMX2MY2([.A1:.A3]+1;[.B1:.B3])", -81.0),
        ("=SUMX2PY2([.A1:.A3]+1;[.B1:.B3])", 139.0),
        ("=SUMXMY2([.A1:.A3]+1;[.B1:.B3])", 27.0),
    ] {
        let expression = parse(source);
        let matrix = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} matrix evaluation: {error}"));
        for row in [0, 2] {
            let scalar = evaluate(
                &expression,
                &resolver,
                &execution,
                Position::new("Main", row, 0),
                Mode::Scalar,
                &Limits::default(),
            )
            .unwrap_or_else(|error| panic!("{source:?} scalar evaluation at row {row}: {error}"));
            let scalar = aggregate_result(&scalar);
            assert_number(&scalar, expected, source);
            assert_eq!(matrix, scalar, "{source:?} mode invariance at row {row}");
        }
    }
}

#[test]
fn deeply_nested_aggregate_projection_obeys_bounded_work() {
    // A projected lazy branch classifies nested aggregates before its first
    // demanded cell is cached.  Keep the AST deliberately deep and cap work
    // well below the depth: the public result must be a typed Work refusal,
    // rather than an unbounded recursive walk or a process-stack failure.
    let mut nested = "1".to_owned();
    for _ in 0..192 {
        nested = format!("SUM({nested})");
    }
    let source = format!("=IF({{TRUE();TRUE()}};{nested};0)");
    let expression = parse(&source);
    let resolver = AggregateResolver::blank(2, 2);
    let (budget, _cancellation, execution) = execution("ods-formula-aggregate-deep-projection");
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default().with_max_steps(256),
    )
    .expect_err("deep aggregate classification must be bounded by Work");
    assert!(
        matches!(
            &error,
            EvaluationFailure::ResourceLimit(limit)
                if matches!(limit.resource, Resource::Work | Resource::Depth)
        ),
        "wrong deep-projection failure: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn aggregate_formula_limits_cancellation_storage_and_owned_results_are_atomic() {
    let resolver = AggregateResolver::blank(2, 2);
    let expression = parse("=SUMPRODUCT({1;2|3;4};{5;6|7;8})");
    let (budget, _cancellation, limits_execution) = execution("ods-formula-aggregate-array-limits");
    let error = evaluate(
        &expression,
        &resolver,
        &limits_execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero work must refuse before reducing arrays");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "wrong zero-work error: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation, storage_execution) =
        execution("ods-formula-aggregate-array-storage");
    let error = evaluate(
        &expression,
        &resolver,
        &storage_execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("array reduction must refuse when its admitted storage is zero");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "wrong zero-storage error: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation, cancel_execution) = execution("ods-formula-aggregate-cancel");
    let mut cancelling = AggregateResolver::blank(64, 1);
    cancelling.fill_column(0, 1.0);
    cancelling.cancel_after_read(&cancellation);
    let expression = parse("=SUM([.A1:.A64])");
    let error = evaluate(
        &expression,
        &cancelling,
        &cancel_execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("cancellation during a streaming reduction must be typed");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert!(cancelling.reads() < 64);
    assert_eq!(budget.used(Resource::Memory), 0);

    let resolver = AggregateResolver::blank(2, 2);
    let (_budget, _cancellation, owned_execution) = execution("ods-formula-aggregate-owned");
    let expression = parse("=SUMPRODUCT({1;2|3;4};{5;6|7;8})");
    let result = evaluate(
        &expression,
        &resolver,
        &owned_execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("numeric aggregate result should evaluate");
    let owned = result
        .to_owned(&owned_execution, &Limits::default())
        .expect("numeric aggregate scalar should be ownable");
    assert!(
        matches!(owned.value(), OwnedValueView::Number(value) if (value - 70.0).abs() <= 64.0 * f64::EPSILON)
    );
}

#[test]
fn aggregate_numeric_goldens_cover_cancellation_overflow_and_subnormal_results() {
    let resolver = AggregateResolver::blank(2, 2);
    let (_budget, _cancellation, execution) = execution("ods-formula-aggregate-goldens");
    let exact_cases = [
        ("=SUM(1e16;1;-1e16)", 0x3ff0_0000_0000_0000),
        ("=SUM(5e-324;1e308;-1e308)", 0x0000_0000_0000_0001),
        (
            "=SUM(1.7976931348623157e308;1.7976931348623157e308;-1.7976931348623157e308)",
            0x7fefffffffffffff,
        ),
        (
            "=SUMX2MY2({1.7976931348623157e308;1};{1.7976931348623157e308;0})",
            0x3ff0_0000_0000_0000,
        ),
        ("=SUMX2MY2({1e308};{-1e308})", 0),
        (
            "=SUMPRODUCT({1.7976931348623157e308;1.7976931348623157e308};{2;-2})",
            0,
        ),
        ("=SUMSQ({1.5e-162;1.5e-162})", 1),
        ("=SUMX2PY2({1e-162;1e-162};{1e-162;1e-162})", 1),
        (
            "=SUMXMY2({1e154};{9.999999999999999e+153})",
            0x7950_0000_0000_0000,
        ),
        ("=SUMSQ({1e-162;1e-162;1e-162})", 1),
        (
            "=SUMPRODUCT({1.7976931348623157e308;1.7976931348623157e308;1};{1.7976931348623157e308;-1.7976931348623157e308;1})",
            0x3ff0_0000_0000_0000,
        ),
        ("=SUMPRODUCT({1.5e-162;1.5e-162};{1.5e-162;1.5e-162})", 1),
    ];
    for (source, expected) in exact_cases {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_bits(&result, expected, source);
    }

    // Product accumulation and a scaled three-factor product are finite
    // representable results whose exact Fraction oracle may round by a few
    // ulps when the implementation combines mantissa/exponent buckets.
    for (source, expected) in [
        ("=PRODUCT(1e308;1e308;1e-308;1e-308)", 0x3fefffffffffffff),
        ("=PRODUCT(1e-308;1e-308;1e308;1e308)", 0x3fefffffffffffff),
        (
            "=SUMPRODUCT({1.7976931348623157e308};{1.7976931348623157e308};{1e-309})",
            0x7fc7_02ae_4d1f_b5df,
        ),
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_bits_within_ulp(&result, expected, 8, source);
    }

    for source in [
        "=PRODUCT(1e308;1e308)",
        "=SUMPRODUCT({1e308};{1e308})",
        "=SUMX2PY2({1e308};{1e308})",
        "=SUMX2MY2({1e308};{0})",
    ] {
        let result = evaluate_source(
            source,
            &resolver,
            &execution,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
        assert_formula_error(&result, ScalarError::Number, source);
    }
}

#[test]
fn k_factor_sumproduct_refuses_exponent_window_and_refunds_memory() {
    // Five factors with one huge and one tiny cell force the K>2 scaled
    // product accumulator to retain terms separated by more than its 8128
    // input-bit window.  The refusal is a resource outcome: an ordinary
    // #NUM! value would let IFERROR hide the bounded-precision decision.
    let source =
        "=SUMPRODUCT({1e308;1e-308};{1e308;1e-308};{1e308;1e-308};{1e308;1e-308};{1e308;1e-308})";
    let resolver = AggregateResolver::blank(2, 2);
    let expression = parse(source);
    let (budget, _cancellation, execution) = execution("ods-formula-aggregate-exponent-window");
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("a K>2 exponent-window refusal must be typed");
    match &error {
        EvaluationFailure::ResourceLimit(limit) => {
            assert_eq!(limit.resource, Resource::Memory, "{source}: {error:?}");
            assert_eq!(
                limit.limit, 2_032,
                "the 8128-bit two-magnitude window is reported in bytes: {limit:?}"
            );
            assert!(
                limit.observed > limit.limit,
                "window refusal must report bytes above its limit: {limit:?}"
            );
        },
        _ => panic!("{source}: expected Memory ResourceLimit, got {error:?}"),
    }
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "failed K-factor reduction must refund all temporary storage"
    );

    let wrapped_source = format!("=IFERROR({};7)", source.trim_start_matches('='));
    let wrapped = parse(&wrapped_source);
    let wrapped_error = evaluate(
        &wrapped,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("IFERROR must not catch an evaluator ResourceLimit");
    match &wrapped_error {
        EvaluationFailure::ResourceLimit(limit) => {
            assert_eq!(limit.resource, Resource::Memory);
            assert_eq!(limit.limit, 2_032);
            assert!(limit.observed > limit.limit);
        },
        _ => panic!("IFERROR changed the resource outcome: {wrapped_error:?}"),
    }
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "resource refusal under IFERROR must also refund temporary storage"
    );
}
