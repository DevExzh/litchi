//! Independent value-evaluator coverage for the OpenFormula conditional
//! aggregate family in Part 4 §§4.11.8, 6.13.9–6.13.10, 6.16.62–6.16.63,
//! and 6.18.5–6.18.6.
//!
//! The six functions consume worksheet references.  These tests therefore use
//! a deliberately observable read-only resolver instead of the scalar bridge.
//! Numeric criteria, comparator criteria, text criteria, empty criterion
//! references, aligned alternate ranges, reference lists, and no-number
//! average cases are checked independently of the implementation's helpers.

use std::{
    cell::{Cell, RefCell},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        UnsupportedKind, evaluate_scalar,
    },
    expression::Expression,
};

const CRITERIA: &str = "[.A1:.A6]";
const LABELS: &str = "[.B1:.B6]";
const VALUES: &str = "[.C1:.C6]";
const ALTERNATE: &str = "[.D1:.D6]";
const FLAGS: &str = "[.F1:.F6]";

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
struct ConditionalResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(usize, usize)>>,
}

impl ConditionalResolver {
    fn standard() -> Self {
        // A:F is the data plane.  G is reserved for a physically empty
        // criterion reference, and H is a second shape used by mismatch tests.
        let mut resolver = Self {
            rows: 8,
            columns: 10,
            cells: (0..80).map(|_| FixtureCell::Empty).collect(),
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
        };

        // A: numeric criteria; B: text criteria; C: source values; D:
        // alternate source values; E: text-only alternate values; F: logical
        // criteria.  The fourth row deliberately contains text in C so SUMIF
        // and AVERAGEIF must apply their Number-only admission rule.
        let rows = [
            (1.0, "East", FixtureCell::Number(10.0), 100.0, "oak", true),
            (2.0, "West", FixtureCell::Number(20.0), 200.0, "pine", false),
            (1.0, "East", FixtureCell::Number(30.0), 300.0, "oak", true),
            (
                3.0,
                "North",
                FixtureCell::Text("ignored".to_owned()),
                400.0,
                "oak",
                false,
            ),
            (0.0, "", FixtureCell::Number(50.0), 500.0, "", true),
            (
                1.0,
                "West",
                FixtureCell::Number(60.0),
                600.0,
                "oakwood",
                false,
            ),
        ];
        for (row, (criterion, label, value, alternate, text, flag)) in rows.into_iter().enumerate()
        {
            resolver.set(row, 0, FixtureCell::Number(criterion));
            resolver.set(row, 1, FixtureCell::Text(label.to_owned()));
            resolver.set(row, 2, value);
            resolver.set(row, 3, FixtureCell::Number(alternate));
            resolver.set(row, 4, FixtureCell::Text(text.to_owned()));
            resolver.set(row, 5, FixtureCell::Logical(flag));
        }
        resolver.set(4, 1, FixtureCell::Empty);
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: FixtureCell) {
        let index = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column))
            .expect("fixture coordinate fits");
        self.cells[index] = value;
    }

    fn set_cell(&mut self, row: usize, column: usize, value: FixtureCell) {
        self.set(row, column, value);
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn read_order(&self) -> Vec<(usize, usize)> {
        self.read_order.borrow().clone()
    }

    fn clear_reads(&self) {
        self.reads.set(0);
        self.read_order.borrow_mut().clear();
    }
}

impl Resolver for ConditionalResolver {
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
        self.read_order.borrow_mut().push((row, column));
        if !matches!(sheet, "Main" | "Data" | "Archive") {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if sheet == "Data" && row == 0 && column == 0 {
            return Ok(CellRead::Number(2.0));
        }
        if sheet == "Archive" && row == 0 && column == 0 {
            return Ok(CellRead::Number(3.0));
        }
        if sheet == "Data" && row == 0 && column == 3 {
            return Ok(CellRead::Number(200.0));
        }
        if sheet == "Archive" && row == 0 && column == 3 {
            return Ok(CellRead::Number(300.0));
        }
        if sheet != "Main" {
            return Ok(CellRead::Empty);
        }
        let Some(index) = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column))
        else {
            return Ok(CellRead::Empty);
        };
        Ok(match self.cells.get(index) {
            None | Some(FixtureCell::Empty) => CellRead::Empty,
            Some(FixtureCell::Number(value)) => CellRead::Number(*value),
            Some(FixtureCell::Logical(value)) => CellRead::Logical(*value),
            Some(FixtureCell::Text(value)) => CellRead::Text(value.as_str()),
            Some(FixtureCell::Error(error)) => CellRead::Error(*error),
            Some(FixtureCell::Unsupported) => CellRead::Unsupported,
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
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1_024).expect("one in-flight KiB"),
        0,
    )
    .expect("valid execution limits");
    (
        budget.clone(),
        cancellation,
        ExecutionContext::new(budget, token, limits),
    )
}

fn scalar_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1_024).expect("one in-flight KiB"),
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

fn evaluate(
    source: &str,
    resolver: &ConditionalResolver,
    execution: &ExecutionContext,
) -> Result<ConditionalResult, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, resolver, &context, &Limits::default())?;
    Ok(conditional_result(&result))
}

fn evaluate_expression<'a>(
    expression: &'a Expression,
    resolver: &'a ConditionalResolver,
    execution: &ExecutionContext,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    value::evaluate(expression, resolver, &context, &Limits::default())
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum ConditionalResult {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn conditional_result(result: &Evaluated<'_>) -> ConditionalResult {
    match result.value() {
        Value::Number(value) => ConditionalResult::Number(value),
        Value::Error(error) => ConditionalResult::Error(error),
        _ => ConditionalResult::Other,
    }
}

fn assert_number(result: &ConditionalResult, expected: f64, source: &str) {
    match result {
        ConditionalResult::Number(actual) => {
            let tolerance = expected.abs().max(1.0) * 64.0 * f64::EPSILON;
            assert!(
                (*actual - expected).abs() <= tolerance,
                "{source:?}: expected {expected}, got {actual} (tol {tolerance:e})"
            );
        },
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

fn assert_formula_error(result: &ConditionalResult, source: &str) {
    assert!(
        matches!(result, ConditionalResult::Error(_)),
        "{source:?}: expected a formula Error value, got {:?}",
        result
    );
}

fn assert_formula_error_kind(result: &ConditionalResult, expected: ScalarError, source: &str) {
    match result {
        ConditionalResult::Error(actual) => assert_eq!(*actual, expected, "{source:?}"),
        other => panic!("{source:?}: expected formula Error({expected}), got {other:?}"),
    }
}

#[test]
fn all_six_conditional_aggregates_follow_the_reference_sequence_contract() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-family");

    for (source, expected) in [
        (format!("=SUMIF({CRITERIA};1;{VALUES})"), 100.0),
        (
            format!("=SUMIFS({VALUES};{CRITERIA};1;{LABELS};\"East\")"),
            40.0,
        ),
        (format!("=COUNTIF({CRITERIA};1)"), 3.0),
        (format!("=COUNTIFS({CRITERIA};1;{LABELS};\"East\")"), 2.0),
        (format!("=AVERAGEIF({CRITERIA};1)"), 1.0),
        (
            format!("=AVERAGEIFS({ALTERNATE};{CRITERIA};1;{LABELS};\"East\")"),
            200.0,
        ),
    ] {
        let result = evaluate(&source, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, &source);
    }
    assert!(
        resolver.reads() > 0,
        "conditional aggregates must read references"
    );
}

#[test]
fn numeric_logical_text_and_comparator_criteria_are_typed() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-criteria");

    for (source, expected) in [
        (format!("=COUNTIF({CRITERIA};\">=2\")"), 2.0),
        (format!("=COUNTIF({CRITERIA};\"<2\")"), 4.0),
        (format!("=COUNTIF({CRITERIA};\"<=1\")"), 4.0),
        (format!("=COUNTIF({CRITERIA};\">2\")"), 1.0),
        (format!("=COUNTIF({CRITERIA};\"<>1\")"), 3.0),
        (format!("=COUNTIF({LABELS};\"East\")"), 2.0),
        (format!("=COUNTIF({LABELS};\"east\")"), 0.0),
        (format!("=COUNTIF({LABELS};\" East \")"), 0.0),
        (format!("=COUNTIF({LABELS};\"<>\")"), 5.0),
        (format!("=COUNTIF({FLAGS};TRUE())"), 3.0),
        (format!("=SUMIF({LABELS};\"East\";{ALTERNATE})"), 400.0),
    ] {
        let result = evaluate(&source, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, &source);
    }

    // An explicit empty equality criterion matches the empty label cell.  A
    // comparator containing zero is distinct and only matches the numeric
    // zero in A5.
    let source = format!("=COUNTIF({LABELS};\"=\")");
    let result = evaluate(&source, &resolver, &execution).expect("empty equality criterion");
    assert_number(&result, 1.0, &source);
    let source = format!("=COUNTIF({CRITERIA};\"=0\")");
    let result = evaluate(&source, &resolver, &execution).expect("explicit zero criterion");
    assert_number(&result, 1.0, &source);

    // A reference to the physically empty G1 criterion cell is interpreted as
    // numeric zero by §4.11.8, so it selects A5 but does not match an empty
    // criterion-range cell as text.
    let source = "=COUNTIF([.A1:.A6];[.G1])";
    let result = evaluate(source, &resolver, &execution).expect("empty criterion reference");
    assert_number(&result, 1.0, source);
}

#[test]
fn optional_sum_and_average_ranges_use_the_criteria_geometry_from_the_top_left() {
    let mut resolver = ConditionalResolver::standard();
    resolver.set_cell(6, 3, FixtureCell::Number(700.0));
    let (_budget, _cancellation, execution) =
        execution("ods-formula-conditional-aligned-alternate-range");

    // D2:D7 is intentionally shifted by one row.  Only its first six cells
    // participate, and rows matching A1:A6 are selected by corresponding
    // position: D2 + D4 + D7 = 200 + 400 + 700.
    let source = format!("=SUMIF({CRITERIA};1;[.D2])");
    let result = evaluate(&source, &resolver, &execution).expect("aligned SUMIF");
    assert_number(&result, 1_300.0, &source);

    // A wider actual S range still uses only its top-left anchor and the
    // dimensions of R.  E is ignored even though it is present in S.
    let source = format!("=AVERAGEIF({CRITERIA};1;[.D2:.E7])");
    let result = evaluate(&source, &resolver, &execution).expect("aligned AVERAGEIF");
    assert_number(&result, 1_300.0 / 3.0, &source);
}

#[test]
fn reference_lists_are_accepted_by_sumif_and_countif_with_single_record_rule() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-reference-list");

    // Intersecting one member of a two-record list leaves a one-record
    // ReferenceList. SUMIF's optional S admits this exact list shape.
    let list_expression = parse("=([.A1]~[.B1])![.A1]");
    let list_context =
        Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let list_value = value::evaluate(
        &list_expression,
        &resolver,
        &list_context,
        &Limits::default(),
    )
    .expect("single-record list expression");
    assert_eq!(
        list_value.as_reference_list().map(|list| list.len()),
        Some(1),
        "intersection must preserve the one-entry ReferenceList kind"
    );

    resolver.clear_reads();
    let source = "=SUMIF((([.A1]~[.B1])![.A1]);1;[.D1])";
    let result = evaluate(source, &resolver, &execution).expect("single-record list SUMIF");
    assert_number(&result, 100.0, source);
    assert_eq!(
        resolver.read_order(),
        vec![(0, 0), (0, 3)],
        "single-record SUMIF reads criterion before the selected destination"
    );

    resolver.clear_reads();
    let source = "=SUMIF(([.A1:.A3]~[.A4:.A6]);1)";
    let result = evaluate(source, &resolver, &execution).expect("reference-list SUMIF");
    // With no S, SUMIF sums the matching Number values in R itself.
    assert_number(&result, 3.0, source);

    resolver.clear_reads();
    let source = "=COUNTIF(([.A1]~[.A1]~[.A2]);1)";
    let result = evaluate(source, &resolver, &execution).expect("ordered duplicate COUNTIF");
    assert_number(&result, 2.0, source);
    assert_eq!(
        resolver.read_order(),
        vec![(0, 0), (0, 0), (1, 0)],
        "COUNTIF must preserve ReferenceList order and duplicate occurrences"
    );

    resolver.clear_reads();
    let source = "=SUMIF(([.A1:.A3]~[.A4:.A6]);1;[.D1:.D6])";
    let result = evaluate(source, &resolver, &execution).expect("reference-list alternate error");
    assert_formula_error(&result, source);
    assert_eq!(resolver.reads(), 0, "list rejection must be read-free");
}

#[test]
fn a_single_three_dimensional_reference_is_one_conditional_geometry() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-conditional-three-dimensional");

    // A1 on Main/Data/Archive is [1, 2, 3], while D1 is [100, 200, 300].
    // The one 3-D reference therefore aligns its three physical planes rather
    // than being mistaken for a ReferenceList of independent arguments.
    let source = "=SUMIF([Main.A1:Archive.A1];\">=2\";[Main.D1:Archive.D1])";
    let result = evaluate(source, &resolver, &execution).expect("3-D conditional reference");
    assert_number(&result, 500.0, source);
    assert!(
        resolver.reads() >= 5,
        "three criteria and both selected values must be read: {}",
        resolver.reads()
    );
}

#[test]
fn one_plane_three_dimensional_anchor_advances_for_sum_and_average() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-conditional-three-dimensional-anchor");
    let sources = [
        ("=SUMIF([Main.A1:Archive.A1];\">=1\";[Main.D1])", 600.0),
        ("=AVERAGEIF([Main.A1:Archive.A1];\">=1\";[Main.D1])", 200.0),
    ];
    for (source, expected) in sources {
        let result = evaluate(source, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, source);
    }
}

#[test]
fn generated_missing_three_dimensional_anchor_plane_is_reference_error() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-conditional-three-dimensional-missing-anchor-plane");
    for source in [
        "=SUMIF([Main.A1:Archive.A1];\">=1\";[Archive.D1])",
        "=AVERAGEIF([Main.A1:Archive.A1];\">=1\";[Archive.D1])",
    ] {
        let result = evaluate(source, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should produce a formula error: {error}"));
        assert_formula_error_kind(&result, ScalarError::Reference, source);
        assert_eq!(
            resolver.reads(),
            0,
            "{source:?} scanned before plane validation"
        );
    }
}

#[test]
fn sumif_clips_generated_rows_and_columns_but_averageif_validates_first() {
    let mut resolver = ConditionalResolver::standard();
    resolver.set_cell(7, 3, FixtureCell::Number(900.0));
    resolver.set_cell(0, 9, FixtureCell::Number(900.0));
    let (_budget, _cancellation, execution) =
        execution("ods-formula-conditional-destination-clipping");

    // D8 is in bounds for the first selected row; the later selected rows
    // generate D10 and D13, which SUMIF silently omits.
    let row_source = "=SUMIF([.A1:.A6];1;[.D8])";
    let row_result = evaluate(row_source, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{row_source:?} should evaluate: {error}"));
    assert_number(&row_result, 900.0, row_source);

    // J1 is in bounds for the first selected column; K1 and L1 are clipped
    // for the other numeric matches in A1:F1.
    resolver.clear_reads();
    let column_source = "=SUMIF([.A1:.F1];\">=0\";[.J1])";
    let column_result = evaluate(column_source, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{column_source:?} should evaluate: {error}"));
    assert_number(&column_result, 900.0, column_source);
    assert!(resolver.read_order().contains(&(0, 9)));

    // AVERAGEIF has no clipping permission. Its complete generated geometry
    // is rejected before the no-match scan or any provider cell read.
    resolver.clear_reads();
    let average_source = "=AVERAGEIF([.A1:.F6];99;[.J8])";
    let average_result = evaluate(average_source, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{average_source:?} should evaluate: {error}"));
    assert_formula_error_kind(&average_result, ScalarError::Reference, average_source);
    assert_eq!(
        resolver.reads(),
        0,
        "AVERAGEIF scanned before geometry refusal"
    );
}

#[test]
fn criterion_range_must_project_to_one_cell() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-criterion-shape");
    let source = "=COUNTIF([.A1:.A6];[.G1:.G2])";
    let result = evaluate(source, &resolver, &execution).expect("criterion shape formula value");
    assert_formula_error(&result, source);
}

#[test]
fn criterion_arrays_multicell_references_lists_and_missing_slots_stay_value_errors() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-conditional-criterion-pseudotypes");
    let source_for = |function: &str, criterion: &str| match function {
        "SUMIF" => format!("=SUMIF({CRITERIA};{criterion})"),
        "SUMIFS" => format!("=SUMIFS({VALUES};{CRITERIA};{criterion})"),
        "COUNTIF" => format!("=COUNTIF({CRITERIA};{criterion})"),
        "COUNTIFS" => format!("=COUNTIFS({CRITERIA};{criterion})"),
        "AVERAGEIF" => format!("=AVERAGEIF({CRITERIA};{criterion})"),
        "AVERAGEIFS" => format!("=AVERAGEIFS({VALUES};{CRITERIA};{criterion})"),
        _ => unreachable!("unknown conditional function"),
    };
    let functions = [
        "SUMIF",
        "SUMIFS",
        "COUNTIF",
        "COUNTIFS",
        "AVERAGEIF",
        "AVERAGEIFS",
    ];
    for function in functions {
        for criterion in ["{1;2}", "[.G1:.G2]", "([.G1]~[.G2])", ""] {
            let source = source_for(function, criterion);
            let result = evaluate(&source, &resolver, &execution).unwrap_or_else(|error| {
                panic!("{source:?} should return a formula value: {error}")
            });
            assert_formula_error_kind(&result, ScalarError::Value, &source);
        }
        for criterion in [
            "IF(TRUE();{1;2};0)",
            "IF(TRUE();[.G1:.G2];0)",
            "IF(TRUE();([.G1]~[.G2]);0)",
        ] {
            let source = source_for(function, criterion);
            let result = evaluate(&source, &resolver, &execution).unwrap_or_else(|error| {
                panic!("{source:?} should return a formula value: {error}")
            });
            assert_formula_error_kind(&result, ScalarError::Value, &source);
        }
    }

    // A missing required criterion remains Missing when the aggregate is
    // evaluated in a projected IF branch. The branch's other position stays
    // lazy and returns the ordinary IF fallback value.
    for function in functions {
        let branch = source_for(function, "");
        let branch = branch
            .strip_prefix('=')
            .expect("conditional source starts with an equals sign");
        let source = format!("=IF({{TRUE();FALSE()}};{branch};0)");
        let expression = parse(&source);
        let evaluated = evaluate_expression(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        let array = evaluated
            .as_array()
            .unwrap_or_else(|| panic!("{source:?} should retain its projected shape"));
        assert!(matches!(
            array.get(0),
            Some(Value::Error(ScalarError::Value))
        ));
        assert!(matches!(array.get(1), Some(Value::Number(0.0))));
    }
}

#[test]
fn conditional_ifs_require_equal_reference_geometry_and_and_the_positions() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-ifs-shape");

    for source in [
        "=SUMIFS([.C1:.C6];[.A1:.A5];1)",
        "=COUNTIFS([.A1:.A6];1;[.B1:.B5];\"East\")",
        "=AVERAGEIFS([.D1:.D6];[.A1:.A5];1)",
    ] {
        let result = evaluate(source, &resolver, &execution).expect("shape mismatch formula value");
        assert_formula_error(&result, source);
    }

    let source = "=SUMIFS([.D1:.D6];[.A1:.A6];1;[.B1:.B6];\"East\")";
    let result = evaluate(source, &resolver, &execution).expect("AND criteria");
    assert_number(&result, 400.0, source);
}

#[test]
fn averages_reject_no_matches_and_no_numeric_selected_values() {
    let mut resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-average-errors");

    for source in [
        "=AVERAGEIF([.A1:.A6];99)",
        "=AVERAGEIFS([.E1:.E6];[.A1:.A6];1)",
        "=AVERAGEIFS([.E1:.E6];[.A1:.A6];1;[.B1:.B6];\"East\")",
    ] {
        let result = evaluate(source, &resolver, &execution).expect("average error value");
        assert_formula_error_kind(&result, ScalarError::DivisionByZero, source);
    }

    // A selected error in a criteria range remains a formula error.  This is
    // separate from the no-number average case above, which is an ordinary
    // aggregate error after all matching cells have been examined.
    resolver.set_cell(0, 0, FixtureCell::Error(ScalarError::NotAvailable));
    let source = "=COUNTIF([.A1:.A6];1)";
    let result = evaluate(source, &resolver, &execution).expect("criteria error result");
    assert_formula_error_kind(&result, ScalarError::NotAvailable, source);
}

#[test]
fn conditional_functions_reject_constant_range_arguments_as_formula_values() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-reference-only");

    for source in [
        "=SUMIF({1;2};1)",
        "=SUMIFS({1;2};{1;2};1)",
        "=COUNTIF({1;2};1)",
        "=COUNTIFS({1;2};{1;2};1)",
        "=AVERAGEIF({1;2};1)",
        "=AVERAGEIFS({1;2};{1;2};1)",
    ] {
        let result = evaluate(source, &resolver, &execution).expect("constant range value error");
        assert_formula_error(&result, source);
    }
    assert_eq!(resolver.reads(), 0, "constant range refusal is read-free");
}

#[test]
fn conditional_aggregate_arity_is_checked_before_range_scanning() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-arity");

    for source in [
        "=COUNTIF([.A1:.A6];1;[.C1:.C6])",
        "=COUNTIFS([.A1:.A6];1;[.B1:.B6])",
        "=AVERAGEIF([.A1:.A6];1;[.C1:.C6];[.D1:.D6])",
    ] {
        let result = evaluate(source, &resolver, &execution).expect("arity formula value");
        assert_formula_error_kind(&result, ScalarError::Value, source);
    }

    for (source, expected) in [
        ("=COUNTIFS([.A1:.A6];1)", 3.0),
        ("=COUNTIFS([.A1:.A6];1;[.B1:.B6];\"East\")", 2.0),
    ] {
        let result = evaluate(source, &resolver, &execution).expect("valid COUNTIFS arity");
        assert_number(&result, expected, source);
    }
}

#[test]
fn error_criteria_is_preserved_before_reference_scanning() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-error-order");

    for source in [
        "=SUMIF([.A1:.A6];#N/A)",
        "=SUMIFS([.C1:.C6];[.A1:.A6];#N/A)",
        "=COUNTIF([.A1:.A6];#N/A)",
        "=COUNTIFS([.A1:.A6];#N/A)",
        "=AVERAGEIF([.A1:.A6];#N/A)",
        "=AVERAGEIFS([.C1:.C6];[.A1:.A6];#N/A)",
    ] {
        let result = evaluate(source, &resolver, &execution).expect("error criterion result");
        assert_formula_error_kind(&result, ScalarError::NotAvailable, source);
    }
}

#[test]
fn criterion_cell_errors_and_formula_error_precedence_are_observable() {
    let mut resolver = ConditionalResolver::standard();
    resolver.set_cell(0, 7, FixtureCell::Error(ScalarError::NotAvailable));
    let (_budget, _cancellation, execution) =
        execution("ods-formula-conditional-single-cell-criterion-error");
    let source = "=COUNTIF([.A1:.A6];[.H1])";
    let result = evaluate(source, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
    assert_formula_error_kind(&result, ScalarError::NotAvailable, source);
    assert_eq!(resolver.read_order(), vec![(0, 7)]);

    // A formula Error observed first is retained only until a later provider
    // operation fails with a typed capability error. The incomplete scan must
    // publish that typed failure instead of a catchable formula value.
    let mut resolver = ConditionalResolver::standard();
    resolver.set_cell(0, 1, FixtureCell::Error(ScalarError::NotAvailable));
    resolver.set_cell(1, 1, FixtureCell::Unsupported);
    resolver.clear_reads();
    let source = "=COUNTIF([.B1:.B6];\"East\")";
    let error = evaluate(source, &resolver, &execution)
        .expect_err("later typed provider failure must supersede a retained formula error");
    assert!(
        matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
        ),
        "wrong later typed failure: {error:?}"
    );
    assert_eq!(resolver.read_order(), vec![(0, 1), (1, 1)]);

    // A selected destination Error at an earlier position wins over a later
    // criterion-range Error, because positions are scanned in order and the
    // destination is observed after that position's criterion.
    let mut resolver = ConditionalResolver::standard();
    resolver.set_cell(0, 3, FixtureCell::Error(ScalarError::DivisionByZero));
    resolver.set_cell(2, 0, FixtureCell::Error(ScalarError::NotAvailable));
    resolver.clear_reads();
    let source = "=SUMIF([.A1:.A6];1;[.D1])";
    let result = evaluate(source, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should return a formula error: {error}"));
    assert_formula_error_kind(&result, ScalarError::DivisionByZero, source);
}

#[test]
fn averageifs_reads_average_values_only_after_all_criteria_match() {
    let mut resolver = ConditionalResolver::standard();
    // Put unsupported cells in D2 and D4.  Both fail A=1, so AVERAGEIFS must
    // never read them while evaluating D1:D6.  The selected D1 and D3 values
    // still produce the average 200.
    resolver.set_cell(1, 3, FixtureCell::Unsupported);
    resolver.set_cell(3, 3, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-lazy-average");
    let source = "=AVERAGEIFS([.D1:.D6];[.A1:.A6];1)";
    let result = evaluate(source, &resolver, &execution)
        .expect("unselected average values must remain lazy");
    assert_number(&result, 1_000.0 / 3.0, source);
    assert!(
        !resolver
            .read_order()
            .iter()
            .any(|&(row, column)| column == 3 && (row == 1 || row == 3)),
        "AVERAGEIFS read an average value whose criteria did not match: {:?}",
        resolver.read_order()
    );
}

#[test]
fn projected_conditional_scalar_is_cached_across_matrix_broadcast_positions() {
    let resolver = ConditionalResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-conditional-projected-cache");
    let baseline = evaluate("=SUMIF([.A1:.A6];1;[.C1:.C6])", &resolver, &execution)
        .expect("conditional baseline");
    assert_number(&baseline, 100.0, "conditional baseline");
    let baseline_reads = resolver.reads();
    resolver.clear_reads();

    let source = "=IF({TRUE();FALSE();TRUE()};SUMIF([.A1:.A6];1;[.C1:.C6]);0)";
    let expression = parse(source);
    let result =
        evaluate_expression(&expression, &resolver, &execution).expect("broadcast conditional");
    let array = result.as_array().expect("IF keeps its condition shape");
    // In this parser profile `;` separates columns, so the three-cell
    // condition is a 1x3 row.  The values remain in row-major order.
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 3));
    assert!(matches!(array.get(0), Some(Value::Number(100.0))));
    assert!(matches!(array.get(1), Some(Value::Number(0.0))));
    assert!(matches!(array.get(2), Some(Value::Number(100.0))));
    assert_eq!(resolver.reads(), baseline_reads);
}

#[test]
fn projected_conditional_recomputes_position_dependent_criteria() {
    let mut resolver = ConditionalResolver::standard();
    // The vertical reference is projected at the IF output row: the first
    // branch cell uses G1 (0), while the second uses G2 (2). Their criteria
    // select different results and must not share a cache slot.
    resolver.set_cell(0, 6, FixtureCell::Number(0.0));
    resolver.set_cell(1, 6, FixtureCell::Number(2.0));
    let (_budget, _cancellation, execution) =
        execution("ods-formula-conditional-position-dependent-criterion");
    let source = "=IF({TRUE()|TRUE()};SUMIF([.A1:.A6];\">\"&[.G1:.G2]);0)";
    let expression = parse(source);
    let result = evaluate_expression(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    let array = result
        .as_array()
        .unwrap_or_else(|| panic!("{source:?} should retain its projected shape"));
    assert_eq!((array.shape().rows(), array.shape().columns()), (2, 1));
    assert!(matches!(array.get(0), Some(Value::Number(8.0))));
    assert!(matches!(array.get(1), Some(Value::Number(3.0))));
}

#[test]
fn scalar_conditional_api_keeps_reference_capability_and_constant_value_boundaries() {
    for function in [
        "SUMIF",
        "SUMIFS",
        "COUNTIF",
        "COUNTIFS",
        "AVERAGEIF",
        "AVERAGEIFS",
    ] {
        let source = match function {
            "SUMIF" | "COUNTIF" | "AVERAGEIF" => {
                format!("={function}([.A1:.A2];1)")
            },
            _ => format!("={function}([.A1:.A2];[.A1:.A2];1)"),
        };
        let expression = parse(&source);
        let (_budget, _cancellation, execution) =
            scalar_execution("ods-formula-conditional-scalar-reference");
        let error = evaluate_scalar(
            &expression,
            &EvaluationContext::new(&execution),
            &EvaluationLimits::default(),
        )
        .expect_err("scalar conditional references require the value resolver");
        assert!(
            matches!(
                error,
                EvaluationFailure::Unsupported(UnsupportedKind::Reference)
            ),
            "{source:?} returned the wrong scalar capability error: {error:?}"
        );
    }

    // The scalar bridge admits the function names so an invalid constant range
    // is a formula #VALUE! result.  Reference-dependent valid calls still hit
    // the typed capability refusal above.
    for source in [
        "=SUMIF(1;1)",
        "=SUMIFS(1;1;1)",
        "=COUNTIF(1;1)",
        "=COUNTIFS(1;1;1)",
        "=AVERAGEIF(1;1)",
        "=AVERAGEIFS(1;1;1)",
    ] {
        let expression = parse(source);
        let (_budget, _cancellation, execution) =
            scalar_execution("ods-formula-conditional-scalar-constant");
        let result = evaluate_scalar(
            &expression,
            &EvaluationContext::new(&execution),
            &EvaluationLimits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should return a formula value: {error:?}"));
        assert!(
            matches!(result.value(), ScalarValue::Error(ScalarError::Value)),
            "{source:?} returned {:?} instead of #VALUE!",
            result.value()
        );
    }
}

#[test]
fn projected_criterion_keeps_nested_matrix_scalar_parameters_position_dependent() {
    let mut resolver = ConditionalResolver::standard();
    resolver.set_cell(0, 6, FixtureCell::Number(1.0));
    resolver.set_cell(1, 6, FixtureCell::Number(2.0));
    let (_budget, _cancellation, execution) =
        execution("ods-conditional-matrix-scalar-criterion-cache");
    // SUM consumes the complete identity matrix, but MUNIT's size still
    // intersects G1:G2 at the current row. Its sums are 1 and 2 respectively.
    let expression = parse("=IF({TRUE()|TRUE()};COUNTIF([.A1:.A6];SUM(MUNIT([.G1:.G2])));0)");
    let result = evaluate_expression(&expression, &resolver, &execution)
        .expect("position-dependent matrix size criterion");
    let array = result.as_array().expect("two projected rows");
    assert_eq!((array.shape().rows(), array.shape().columns()), (2, 1));
    assert!(matches!(array.get(0), Some(Value::Number(3.0))));
    assert!(matches!(array.get(1), Some(Value::Number(1.0))));
}

#[test]
fn selected_formula_error_precedes_final_sum_overflow() {
    let mut resolver = ConditionalResolver::standard();
    for row in 0..3 {
        resolver.set_cell(row, 0, FixtureCell::Number(1.0));
    }
    resolver.set_cell(0, 2, FixtureCell::Number(f64::MAX));
    resolver.set_cell(1, 2, FixtureCell::Number(f64::MAX));
    resolver.set_cell(2, 2, FixtureCell::Error(ScalarError::NotAvailable));
    let (_budget, _cancellation, execution) = execution("ods-conditional-overflow-error-order");
    for source in [
        "=SUMIF([.A1:.A3];1;[.C1])",
        "=SUMIFS([.C1:.C3];[.A1:.A3];1)",
    ] {
        let result = evaluate(source, &resolver, &execution).expect("selected formula error");
        assert_formula_error_kind(&result, ScalarError::NotAvailable, source);
    }
}
