//! Independent value and scalar coverage for the first OpenFormula statistical
//! reducer slice: COUNT, COUNTA, COUNTBLANK, AVERAGE, AVERAGEA, MIN, MAX,
//! MINA, and MAXA.
//!
//! The fixture deliberately keeps Number, Logical, Text, empty, and Error
//! cells distinct. Reference sequence conversion and direct scalar
//! conversion have different rules in Part 4, so both evaluator entry points
//! are exercised here. The resolver also exposes three ordered sheets and
//! read coordinates so 3-D traversal, ordered reference lists, lazy branches,
//! and cache decisions remain observable.

use std::{
    cell::{Cell, RefCell},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    SourceVersion,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits, Mode, OwnedValueView, Position, Resolver,
        SheetExtent, Value,
    },
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        UnsupportedKind, evaluate_scalar,
    },
    expression::Expression,
};

const MIXED: &str = "[.A1:.A8]";
const NUMBERS: &str = "[.B1:.B4]";

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
struct StatisticalResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl StatisticalResolver {
    fn standard() -> Self {
        let mut resolver = Self {
            rows: 8,
            columns: 12,
            cells: (0..96).map(|_| FixtureCell::Empty).collect(),
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            source_versions: None,
            source_version_calls: Cell::new(0),
        };

        // A is intentionally heterogeneous. In a reference NumberSequence
        // only the first two Numbers and the Error participate; Logical and
        // Text cells are skipped, while COUNTBLANK observes the empty cell
        // and this profile's empty Text value.
        for (row, value) in [
            FixtureCell::Number(2.0),
            FixtureCell::Number(-4.0),
            FixtureCell::Logical(true),
            FixtureCell::Logical(false),
            FixtureCell::Text("3".to_owned()),
            FixtureCell::Text(String::new()),
            FixtureCell::Empty,
            FixtureCell::Error(ScalarError::NotAvailable),
        ]
        .into_iter()
        .enumerate()
        {
            resolver.set(row, 0, value);
        }
        for (row, value) in [1.0, 3.0, 5.0, 7.0].into_iter().enumerate() {
            resolver.set(row, 1, FixtureCell::Number(value));
        }
        resolver.set(0, 6, FixtureCell::Number(1.0));
        resolver.set(1, 6, FixtureCell::Number(2.0));
        resolver.set(0, 7, FixtureCell::Number(1.0));
        resolver.set(1, 7, FixtureCell::Number(2.0));
        resolver
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

    fn clear_reads(&self) {
        self.reads.set(0);
        self.read_order.borrow_mut().clear();
    }
}

impl Resolver for StatisticalResolver {
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
        self.reads.set(self.reads.get().saturating_add(1));
        self.read_order
            .borrow_mut()
            .push((sheet.to_owned(), row, column));

        if !matches!(sheet, "Main" | "Data" | "Archive") {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if row >= self.rows || column >= self.columns {
            return Ok(CellRead::Error(ScalarError::Reference));
        }

        // One-cell 3-D probes make the sheet axis visible without changing
        // the Main-plane fixture. The two criterion cells used by the nested
        // MUNIT regression remain the same on every plane.
        if row == 0 && column == 0 {
            return Ok(match sheet {
                "Main" => CellRead::Number(2.0),
                "Data" => CellRead::Number(4.0),
                "Archive" => CellRead::Number(6.0),
                _ => unreachable!(),
            });
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

    fn source_version(
        &self,
        _execution: &ExecutionContext,
    ) -> Result<Option<SourceVersion>, EvaluationFailure> {
        let Some((expected, observed)) = self.source_versions else {
            return Ok(None);
        };
        let call = self.source_version_calls.get();
        self.source_version_calls.set(call.saturating_add(1));
        Ok(Some(if call == 0 { expected } else { observed }))
    }
}

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

#[derive(Clone, Copy, Debug, PartialEq)]
enum StatisticalResult {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn result_value(result: &Evaluated<'_>) -> StatisticalResult {
    match result.value() {
        Value::Number(value) => StatisticalResult::Number(value),
        Value::Error(error) => StatisticalResult::Error(error),
        _ => StatisticalResult::Other,
    }
}

fn evaluate_value_source(
    source: &str,
    resolver: &StatisticalResolver,
    execution: &ExecutionContext,
    mode: Mode,
) -> Result<StatisticalResult, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, &Limits::default())?;
    Ok(result_value(&result))
}

fn evaluate_value<'a>(
    expression: &'a Expression,
    resolver: &'a StatisticalResolver,
    execution: &ExecutionContext,
    position: Position<'a>,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn evaluate_scalar_source(
    source: &str,
    execution: &ExecutionContext,
) -> Result<StatisticalResult, EvaluationFailure> {
    let expression = parse(source);
    let result = evaluate_scalar(
        &expression,
        &EvaluationContext::new(execution),
        &EvaluationLimits::default(),
    )?;
    Ok(match result.value() {
        ScalarValue::Number(value) => StatisticalResult::Number(*value),
        ScalarValue::Error(error) => StatisticalResult::Error(*error),
        _ => StatisticalResult::Other,
    })
}

fn assert_number(result: StatisticalResult, expected: f64, source: &str) {
    match result {
        StatisticalResult::Number(actual) => {
            if expected == 0.0 {
                assert_eq!(actual.to_bits(), expected.to_bits(), "{source:?}");
            } else {
                assert_eq!(actual, expected, "{source:?}");
            }
        },
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

fn assert_error(result: StatisticalResult, expected: ScalarError, source: &str) {
    match result {
        StatisticalResult::Error(actual) => assert_eq!(actual, expected, "{source:?}"),
        other => panic!("{source:?}: expected Error({expected:?}), got {other:?}"),
    }
}

fn assert_array_numbers(result: &Evaluated<'_>, expected: &[f64], source: &str) {
    let array = result
        .as_array()
        .unwrap_or_else(|| panic!("{source:?}: expected an array result"));
    assert_eq!(array.len(), expected.len(), "{source:?} array length");
    for (index, expected) in expected.iter().copied().enumerate() {
        match array.get(index) {
            Some(Value::Number(actual)) => {
                assert_eq!(actual, expected, "{source:?}[{index}]")
            },
            other => panic!("{source:?}[{index}]: expected Number({expected}), got {other:?}"),
        }
    }
}

#[test]
fn scalar_statistical_functions_follow_direct_value_and_error_rules() {
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-scalar");
    for (source, expected) in [
        ("=COUNT(1;2;TRUE();FALSE())", 4.0),
        ("=COUNT(\"3\")", 1.0),
        ("=COUNTA(1;TRUE();FALSE();\"text\";\"\")", 5.0),
        ("=AVERAGE(\"3\";1;TRUE();FALSE())", 1.25),
        ("=AVERAGEA(1;3;TRUE();FALSE();\"text\";\"\")", 5.0 / 6.0),
        ("=MIN(4;2;9)", 2.0),
        ("=MAX(4;2;9)", 9.0),
        ("=MINA(4;TRUE();\"text\")", 0.0),
        ("=MAXA(4;TRUE();\"text\")", 4.0),
    ] {
        let result = evaluate_scalar_source(source, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(result, expected, source);
    }

    for (source, expected) in [
        ("=AVERAGE(\"not-a-number\")", ScalarError::Value),
        ("=AVERAGEA(#N/A;1)", ScalarError::NotAvailable),
        ("=MIN(#N/A;1)", ScalarError::NotAvailable),
        ("=MAXA(#N/A;1)", ScalarError::NotAvailable),
    ] {
        let result = evaluate_scalar_source(source, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should produce a formula value: {error}"));
        assert_error(result, expected, source);
    }

    // COUNT deliberately suppresses formula Error values, while COUNTA
    // counts an Error as one non-blank value.
    assert_number(
        evaluate_scalar_source("=COUNT(#N/A;2)", &execution).unwrap(),
        1.0,
        "=COUNT(#N/A;2)",
    );
    assert_number(
        evaluate_scalar_source("=COUNT(\"not-a-number\")", &execution).unwrap(),
        0.0,
        "=COUNT(\"not-a-number\")",
    );
    assert_number(
        evaluate_scalar_source("=COUNTA(#N/A;2)", &execution).unwrap(),
        2.0,
        "=COUNTA(#N/A;2)",
    );
}

#[test]
fn scalar_zero_arity_is_typed_and_average_has_no_number_error() {
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-zero-arity");
    for source in ["=AVERAGE()", "=AVERAGEA()"] {
        let result = evaluate_scalar_source(source, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
        assert_error(result, ScalarError::DivisionByZero, source);
    }
    for source in [
        "=COUNT()",
        "=COUNTA()",
        "=MIN()",
        "=MAX()",
        "=MINA()",
        "=MAXA()",
    ] {
        let result = evaluate_scalar_source(source, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
        assert_number(result, 0.0, source);
    }

    let resolver = StatisticalResolver::standard();
    let result = evaluate_value_source("=COUNTBLANK()", &resolver, &execution, Mode::Matrix)
        .expect("COUNTBLANK zero arity should return a formula value");
    assert_error(result, ScalarError::Value, "=COUNTBLANK()");
}

#[test]
fn omitted_slots_have_the_same_typed_profile_in_scalar_and_value_modes() {
    let resolver = StatisticalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-missing");
    for (source, expected) in [
        ("=COUNT(;)", StatisticalResult::Number(0.0)),
        ("=COUNTA(;)", StatisticalResult::Number(2.0)),
        ("=AVERAGE(;)", StatisticalResult::Error(ScalarError::Value)),
        ("=MAXA(;)", StatisticalResult::Error(ScalarError::Value)),
    ] {
        let scalar = evaluate_scalar_source(source, &execution)
            .unwrap_or_else(|error| panic!("{source:?} scalar: {error}"));
        assert_eq!(scalar, expected, "{source:?} scalar profile");
        let matrix = evaluate_value_source(source, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source:?} matrix: {error}"));
        assert_eq!(matrix, expected, "{source:?} matrix profile");
    }
}

#[test]
fn scalar_average_uses_exact_cancellation_and_preserves_signed_subnormal_zero() {
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-numeric");
    for (source, expected_bits) in [
        ("=AVERAGE(1e16;1;-1e16)", 0x3fd5_5555_5555_5555),
        ("=AVERAGE(5e-324;5e-324)", 0x0000_0000_0000_0001),
        ("=AVERAGE(-5e-324;0)", 0x8000_0000_0000_0000),
        (
            "=AVERAGE(1.7976931348623157e308;3;-1.7976931348623157e308)",
            0x3ff0_0000_0000_0000,
        ),
    ] {
        let result = evaluate_scalar_source(source, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        let StatisticalResult::Number(actual) = result else {
            panic!("{source:?}: expected a Number, got {result:?}");
        };
        assert_eq!(actual.to_bits(), expected_bits, "{source:?}");
    }
}

#[test]
fn references_filter_types_and_distinguish_count_average_and_extrema() {
    let resolver = StatisticalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-reference");
    for (source, expected) in [
        (format!("=COUNT({})", MIXED), 2.0),
        (format!("=COUNTA({})", MIXED), 7.0),
        (format!("=COUNTBLANK({})", MIXED), 2.0),
        ("=AVERAGE([.A1:.A7])".to_owned(), -1.0),
        ("=AVERAGEA([.A1:.A7])".to_owned(), -1.0 / 6.0),
        ("=MIN([.A1:.A7])".to_owned(), -4.0),
        ("=MAX([.A1:.A7])".to_owned(), 2.0),
        ("=MINA([.A1:.A7])".to_owned(), -4.0),
        ("=MAXA([.A1:.A7])".to_owned(), 2.0),
    ] {
        let result = evaluate_value_source(&source, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(result, expected, &source);
    }

    // COUNT is the only family member here with an explicit non-propagating
    // Error rule. Number-sequence reducers retain the first cell Error.
    assert_number(
        evaluate_value_source("=COUNT([.A8])", &resolver, &execution, Mode::Matrix).unwrap(),
        0.0,
        "=COUNT([.A8])",
    );
    assert_number(
        evaluate_value_source("=COUNTA([.A8])", &resolver, &execution, Mode::Matrix).unwrap(),
        1.0,
        "=COUNTA([.A8])",
    );
    for function in ["AVERAGE", "AVERAGEA", "MIN", "MAX", "MINA", "MAXA"] {
        let source = format!("={function}([.A8])");
        let result = evaluate_value_source(&source, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
        assert_error(result, ScalarError::NotAvailable, &source);
    }
}

#[test]
fn value_modes_arrays_reference_lists_and_three_dimensional_references_follow_shapes() {
    let resolver = StatisticalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-shapes");

    // A reference reducer is invariant under scalar projection: the complete
    // reference is its NumberSequence argument in either parent mode.
    for source in [
        format!("=COUNT({})", NUMBERS),
        format!("=AVERAGE({})", NUMBERS),
        format!("=MIN({})", NUMBERS),
        format!("=MAX({})", NUMBERS),
    ] {
        let matrix = evaluate_value_source(&source, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source:?} matrix: {error}"));
        let scalar = evaluate_value_source(&source, &resolver, &execution, Mode::Scalar)
            .unwrap_or_else(|error| panic!("{source:?} scalar: {error}"));
        assert_eq!(matrix, scalar, "{source:?} mode result");
    }

    for (source, expected) in [
        ("=COUNT({1;2|3;4})", 4.0),
        ("=COUNT({1;TRUE();#N/A|\"3\";\"not-a-number\";0})", 4.0),
        ("=COUNTA({1;TRUE()|\"text\";#N/A})", 4.0),
        ("=AVERAGE({1;2|3;4})", 2.5),
        ("=AVERAGEA({1;TRUE()|\"text\";\"\"})", 0.5),
        ("=MINA({1;2|3;4})", 1.0),
        ("=MAXA({1;2|3;4})", 4.0),
    ] {
        let result = evaluate_value_source(source, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(result, expected, source);
    }

    // NumberSequenceList functions preserve list occurrence order. AVERAGE's
    // singular NumberSequence intentionally rejects a true union list.
    let source = "=COUNT([.B1:.B2]~[.B3])";
    assert_number(
        evaluate_value_source(source, &resolver, &execution, Mode::Matrix).unwrap(),
        3.0,
        source,
    );
    for (source, expected) in [
        ("=COUNTA([.B1:.B2]~[.B3])", 3.0),
        ("=AVERAGEA([.B1:.B2]~[.B3])", 3.0),
        ("=MIN([.B1:.B2]~[.B3])", 1.0),
        ("=MAX([.B1:.B2]~[.B3])", 5.0),
        ("=MINA([.B1:.B2]~[.B3])", 1.0),
        ("=MAXA([.B1:.B2]~[.B3])", 5.0),
    ] {
        let result = evaluate_value_source(source, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(result, expected, source);
    }
    let source = "=AVERAGE([.B1:.B2]~[.B3])";
    assert_error(
        evaluate_value_source(source, &resolver, &execution, Mode::Matrix).unwrap(),
        ScalarError::Value,
        source,
    );

    for (source, expected) in [
        ("=COUNT([Main.A1:Archive.A1])", 3.0),
        ("=AVERAGE([Main.A1:Archive.A1])", 4.0),
        ("=MIN([Main.A1:Archive.A1])", 2.0),
        ("=MAX([Main.A1:Archive.A1])", 6.0),
    ] {
        let result = evaluate_value_source(source, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(result, expected, source);
    }
}

#[test]
fn value_scalar_and_matrix_modes_share_direct_typed_argument_rules() {
    let resolver = StatisticalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-direct-modes");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=COUNT(1;TRUE();\"not-a-number\";#N/A)", 2.0),
            ("=COUNTA(1;TRUE();\"not-a-number\";\"\";#N/A)", 5.0),
            ("=AVERAGE(1;3;TRUE();FALSE())", 1.25),
            (
                "=AVERAGEA(1;3;TRUE();FALSE();\"not-a-number\";\"\")",
                5.0 / 6.0,
            ),
            ("=MIN(4;2;9)", 2.0),
            ("=MAX(4;2;9)", 9.0),
            ("=MINA(4;TRUE();\"not-a-number\")", 0.0),
            ("=MAXA(4;TRUE();\"not-a-number\")", 4.0),
        ] {
            let result = evaluate_value_source(source, &resolver, &execution, mode)
                .unwrap_or_else(|error| panic!("{source:?} mode {mode:?}: {error}"));
            assert_number(result, expected, source);
        }
    }
}

#[test]
fn countblank_requires_a_reference_and_uses_the_profile_empty_text_policy() {
    let resolver = StatisticalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-countblank");
    for source in ["=COUNTBLANK(1)", "=COUNTBLANK({1;2})"] {
        let result = evaluate_value_source(source, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error}"));
        assert_error(result, ScalarError::Value, source);
    }

    let source = "=COUNTBLANK([.A1:.B2])";
    assert_number(
        evaluate_value_source(source, &resolver, &execution, Mode::Matrix).unwrap(),
        0.0,
        source,
    );

    // A missing sheet is a resolved Reference error, not an empty range.
    let source = "=COUNTBLANK([Missing.A1])";
    assert_error(
        evaluate_value_source(source, &resolver, &execution, Mode::Matrix).unwrap(),
        ScalarError::Reference,
        source,
    );

    // A direct Error is an invalid ReferenceList argument and propagates its
    // formula error; an Error-valued cell inside an admitted reference is
    // simply non-blank and therefore contributes zero.
    assert_error(
        evaluate_value_source("=COUNTBLANK(#N/A)", &resolver, &execution, Mode::Matrix).unwrap(),
        ScalarError::NotAvailable,
        "=COUNTBLANK(#N/A)",
    );
    assert_number(
        evaluate_value_source("=COUNTBLANK([.A8])", &resolver, &execution, Mode::Matrix).unwrap(),
        0.0,
        "=COUNTBLANK([.A8])",
    );

    let source = "=COUNTBLANK([.A1]~[.A7])";
    assert_number(
        evaluate_value_source(source, &resolver, &execution, Mode::Matrix).unwrap(),
        1.0,
        source,
    );
}

#[test]
fn lazy_if_and_iferror_preserve_statistical_error_and_capability_boundaries() {
    let resolver = StatisticalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-lazy");
    let source = "=IF(FALSE();AVERAGE([Missing.A1:.A8]);3)";
    assert_number(
        evaluate_value_source(source, &resolver, &execution, Mode::Matrix).unwrap(),
        3.0,
        source,
    );
    assert_eq!(resolver.reads(), 0, "the unselected branch stays lazy");

    let source = "=IFERROR(AVERAGE([.A8]);7)";
    assert_number(
        evaluate_value_source(source, &resolver, &execution, Mode::Matrix).unwrap(),
        7.0,
        source,
    );
    let source = "=IFERROR(COUNT(#N/A);7)";
    assert_number(
        evaluate_value_source(source, &resolver, &execution, Mode::Matrix).unwrap(),
        0.0,
        source,
    );

    // Formula errors consumed by COUNT/COUNTA do not authorize swallowing a
    // later provider capability failure.
    let mut failing = StatisticalResolver::standard();
    failing.set(0, 0, FixtureCell::Error(ScalarError::NotAvailable));
    failing.set(1, 0, FixtureCell::Unsupported);
    for function in ["COUNT", "COUNTA"] {
        let source = format!("={function}([.A1:.A2])");
        let error = evaluate_value_source(&source, &failing, &execution, Mode::Matrix)
            .expect_err("typed provider failure must escape a counting reducer");
        assert!(
            matches!(
                error,
                EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
            ),
            "{source:?}: wrong typed failure {error:?}"
        );
    }
}

#[test]
fn nested_scalar_projection_caches_invariant_count_but_recomputes_munit_criterion() {
    let mut resolver = StatisticalResolver::standard();
    // Main.A1 is reserved by the fixture's 3-D probe as 2. The resulting
    // criteria range is [2, 1, 1, 1, .5, .25]. AVERAGE(MUNIT(G1:G2)) is 1
    // at G1=1 and .5 at G2=2, so each projected IF cell must count a
    // different criterion result.
    for (row, value) in [0.0, 1.0, 1.0, 1.0, 0.5, 0.25].into_iter().enumerate() {
        resolver.set(row, 0, FixtureCell::Number(value));
    }
    resolver.set(0, 6, FixtureCell::Number(1.0));
    resolver.set(1, 6, FixtureCell::Number(2.0));
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-projection");

    let expression = parse("=IF({TRUE()|TRUE()};COUNTIF([.A1:.A6];AVERAGE(MUNIT([.G1:.G2])));0)");
    let result = evaluate_value(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected nested scalar criterion");
    assert_array_numbers(&result, &[3.0, 1.0], "nested MUNIT criterion");

    resolver.clear_reads();
    let expression = parse("=IF({TRUE()|TRUE()};COUNT([.A1:.A6]);0)");
    let result = evaluate_value(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected invariant COUNT");
    assert_array_numbers(&result, &[6.0, 6.0], "cached COUNT projection");
    assert_eq!(
        resolver.reads(),
        6,
        "position-independent COUNT should be evaluated once and reused"
    );
}

#[test]
fn extrema_canonicalize_signed_zero_while_average_retains_underflow_sign() {
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-zero-sign");
    for source in ["=MIN(-0;0)", "=MAX(-0;0)", "=MINA(-0;0)", "=MAXA(-0;0)"] {
        let result = evaluate_scalar_source(source, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(result, 0.0, source);
    }
    let result = evaluate_scalar_source("=AVERAGE(-5e-324;0)", &execution).unwrap();
    assert_eq!(
        result,
        StatisticalResult::Number(f64::from_bits(0x8000_0000_0000_0000))
    );
}

#[test]
fn statistical_results_can_be_owned_without_the_expression_or_resolver() {
    let resolver = StatisticalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-statistical-owned");
    let expression = parse("=AVERAGE({1;3|5;7})");
    let result = evaluate_value(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("average array");
    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("numeric statistical result should be ownable");
    assert!(matches!(owned.value(), OwnedValueView::Number(value) if value == 4.0));
}
