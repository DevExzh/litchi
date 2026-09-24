//! Independent integration coverage for the OpenFormula 1.4 database family.
//!
//! The expected records and statistics follow Part 4 §§4.11.8–4.11.11 and
//! §§6.9.1–6.9.13 in the local `OpenDocument-v1.4-os.zip`.  A database range
//! has a header row, criteria rows are OR'ed, and fields in one criteria row
//! are AND'ed.  Numeric aggregators use the corresponding scalar sequence
//! rule, so text and empty cells are not counted as numbers.  Textual
//! wildcard/substring criteria are intentionally absent: their result depends
//! on the host properties listed in §6.9.1 and §3.4.

use std::{
    cell::Cell,
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits as ValueLimits, Mode, Position, Resolver,
        SheetExtent, Value,
    },
    evaluation::{EvaluationFailure, EvaluationResult, ScalarError, UnsupportedKind},
    expression::Expression,
};

const DATABASE: &str = "[.A1:.E8]";
const DATABASE_WITHOUT_FAULT: &str = "[.A1:.E7]";
const CRITERIA_OR: &str = "[.G1:.H3]";
const CRITERIA_AND: &str = "[.G1:.H2]";
const CRITERIA_UNIQUE: &str = "[.J1:.K2]";
const CRITERIA_NONE: &str = "[.M1:.N2]";
const CRITERIA_ERROR: &str = "[.P1:.Q2]";
const CRITERIA_EMPTY: &str = "[.S1:.S2]";
const CRITERIA_ZERO: &str = "[.S4:.S5]";
const CRITERIA_GREATER_THAN_15: &str = "[.U1:.U2]";
const PROFILE_DATABASE: &str = "[.A1:.E4]";
const PROFILE_CRITERIA: &str = "[.G1:.G2]";
const CRITERIA_BOUNDS: &str = "[.E1:.F2]";
const EMPTY_HEADER_DATABASE: &str = "[.A1:.D3]";
const DGET_DATABASE: &str = "[.A1:.B5]";
const DGET_EMPTY_CRITERIA: &str = "[.D1:.D2]";
const DGET_LOGICAL_CRITERIA: &str = "[.D3:.D4]";
const DGET_NUMERIC_TEXT_CRITERIA: &str = "[.D5:.D6]";
const DGET_BAD_TEXT_CRITERIA: &str = "[.D7:.D8]";
const WHITESPACE_DATABASE: &str = "[.A1:.B4]";
const WHITESPACE_LITERAL_CRITERIA: &str = "[.D1:.D2]";
const WHITESPACE_EQUALS_CRITERIA: &str = "[.D3:.D4]";
const ERROR_HEADER_DATABASE: &str = "[.A1:.B4]";
const ERROR_HEADER_CRITERIA: &str = "[.D1:.D2]";
const BLANK_CRITERIA_DATABASE: &str = "[.A1:.B3]";
const BLANK_CRITERIA: &str = "[.D1:.D2]";

#[derive(Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct DatabaseResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
}

impl DatabaseResolver {
    fn blank(rows: usize, columns: usize) -> Self {
        Self {
            rows,
            columns,
            cells: (0..rows * columns).map(|_| FixtureCell::Empty).collect(),
            reads: Cell::new(0),
        }
    }

    fn standard() -> Self {
        let mut resolver = Self::blank(8, 21);

        for (column, header) in ["Name", "Region", "Amount", "Score", "Active"]
            .into_iter()
            .enumerate()
        {
            resolver.set(0, column, FixtureCell::Text(header.to_owned()));
        }

        // The first six records drive all ordinary aggregate cases.  Dave's
        // text and EmptyAmount's empty Amount distinguish DCOUNT/DCOUNTA from
        // the number-producing functions without changing the numeric sample
        // [10, 20, 30].  Fault is selected only by CRITERIA_ERROR.
        resolver.record(1, "Alice", "East", FixtureCell::Number(10.0), 1.0, true);
        resolver.record(2, "Bob", "West", FixtureCell::Number(20.0), 2.0, false);
        resolver.record(3, "Carol", "East", FixtureCell::Number(30.0), 3.0, true);
        resolver.record(
            4,
            "Dave",
            "East",
            FixtureCell::Text("n/a".to_owned()),
            4.0,
            true,
        );
        resolver.record(5, "Erin", "North", FixtureCell::Number(40.0), 5.0, true);
        resolver.record(6, "EmptyAmount", "East", FixtureCell::Empty, 7.0, true);
        resolver.record(
            7,
            "Fault",
            "Error",
            FixtureCell::Error(ScalarError::NotAvailable),
            8.0,
            true,
        );

        // G:H: one East/TRUE row (AND), then a West/FALSE row (OR).
        resolver.criteria_pair(6, 0, "Region", "Active");
        resolver.criteria_pair(6, 1, "East", "TRUE");
        resolver.set(2, 6, FixtureCell::Text("West".to_owned()));
        resolver.set(2, 7, FixtureCell::Logical(false));

        // A unique record, a criterion matching none, and a criterion
        // selecting the Error-valued record.
        resolver.criteria_pair(9, 0, "Region", "Active");
        resolver.criteria_pair(9, 1, "West", "FALSE");
        resolver.criteria_pair(12, 0, "Region", "Active");
        resolver.criteria_pair(12, 1, "Nowhere", "TRUE");
        resolver.criteria_pair(15, 0, "Region", "Active");
        resolver.criteria_pair(15, 1, "Error", "TRUE");

        // An empty criterion matches the empty Amount cell; =0 deliberately
        // does not, as required by §4.11.8.
        resolver.set(0, 18, FixtureCell::Text("Amount".to_owned()));
        resolver.set(1, 18, FixtureCell::Text("=".to_owned()));
        resolver.set(3, 18, FixtureCell::Text("Amount".to_owned()));
        resolver.set(4, 18, FixtureCell::Text("=0".to_owned()));

        // Numeric comparator criteria are independent of host wildcard and
        // whole-cell text-search policies.
        resolver.set(0, 20, FixtureCell::Text("Amount".to_owned()));
        resolver.set(1, 20, FixtureCell::Text(">15".to_owned()));
        resolver
    }

    fn duplicate_and_non_text_headers() -> Self {
        let mut resolver = Self::blank(4, 8);
        resolver.set(0, 0, FixtureCell::Text("Amount".to_owned()));
        resolver.set(0, 1, FixtureCell::Text("Amount".to_owned()));
        resolver.set(0, 2, FixtureCell::Number(7.0));
        resolver.set(0, 3, FixtureCell::Empty);
        resolver.set(0, 4, FixtureCell::Text("Key".to_owned()));
        for (row, values) in [
            [10.0, 20.0, 30.0, 40.0],
            [1.0, 2.0, 3.0, 4.0],
            [100.0, 200.0, 300.0, 400.0],
        ]
        .into_iter()
        .enumerate()
        {
            let row = row + 1;
            for (column, value) in values.into_iter().enumerate() {
                resolver.set(row, column, FixtureCell::Number(value));
            }
            resolver.set(row, 4, FixtureCell::Text("all".to_owned()));
        }
        resolver.set(0, 6, FixtureCell::Text("Key".to_owned()));
        resolver.set(1, 6, FixtureCell::Text("all".to_owned()));
        resolver
    }

    fn numeric_criteria_headers() -> Self {
        let mut resolver = Self::blank(4, 8);
        resolver.set(0, 0, FixtureCell::Text("ID".to_owned()));
        resolver.set(0, 1, FixtureCell::Text("Amount".to_owned()));
        for (row, (id, amount)) in [(1.0, 10.0), (2.0, 20.0), (3.0, 30.0)]
            .into_iter()
            .enumerate()
        {
            resolver.set(row + 1, 0, FixtureCell::Number(id));
            resolver.set(row + 1, 1, FixtureCell::Number(amount));
        }
        // Two numeric criteria headers both select field 2.  Their bounds
        // are AND'ed in one criteria row, selecting Amount 10 and 20.
        resolver.set(0, 4, FixtureCell::Number(2.0));
        resolver.set(0, 5, FixtureCell::Number(2.0));
        resolver.set(1, 4, FixtureCell::Text(">=10".to_owned()));
        resolver.set(1, 5, FixtureCell::Text("<=20".to_owned()));
        resolver
    }

    fn unique_empty_text_header() -> Self {
        let mut resolver = Self::blank(3, 6);
        resolver.set(0, 0, FixtureCell::Text(String::new()));
        resolver.set(0, 1, FixtureCell::Empty);
        resolver.set(0, 2, FixtureCell::Text("Value".to_owned()));
        resolver.set(0, 3, FixtureCell::Text("Key".to_owned()));
        resolver.set(1, 0, FixtureCell::Number(5.0));
        resolver.set(1, 1, FixtureCell::Number(50.0));
        resolver.set(1, 2, FixtureCell::Number(1.0));
        resolver.set(1, 3, FixtureCell::Text("only".to_owned()));
        resolver.set(2, 0, FixtureCell::Number(7.0));
        resolver.set(2, 1, FixtureCell::Number(70.0));
        resolver.set(2, 2, FixtureCell::Number(2.0));
        resolver.set(2, 3, FixtureCell::Text("other".to_owned()));
        resolver.set(0, 5, FixtureCell::Text("Key".to_owned()));
        resolver.set(1, 5, FixtureCell::Text("only".to_owned()));
        resolver
    }

    fn dget_conversions() -> Self {
        let mut resolver = Self::blank(8, 6);
        resolver.set(0, 0, FixtureCell::Text("Value".to_owned()));
        resolver.set(0, 1, FixtureCell::Text("Key".to_owned()));
        resolver.set(1, 0, FixtureCell::Empty);
        resolver.set(1, 1, FixtureCell::Text("empty".to_owned()));
        resolver.set(2, 0, FixtureCell::Logical(true));
        resolver.set(2, 1, FixtureCell::Text("logical".to_owned()));
        resolver.set(3, 0, FixtureCell::Text("12.5".to_owned()));
        resolver.set(3, 1, FixtureCell::Text("numeric-text".to_owned()));
        resolver.set(4, 0, FixtureCell::Text("bad".to_owned()));
        resolver.set(4, 1, FixtureCell::Text("bad-text".to_owned()));
        for (row, key) in ["empty", "logical", "numeric-text", "bad-text"]
            .into_iter()
            .enumerate()
        {
            let row = row * 2;
            resolver.set(row, 3, FixtureCell::Text("Key".to_owned()));
            resolver.set(row + 1, 3, FixtureCell::Text(key.to_owned()));
        }
        resolver
    }

    fn whitespace_criteria() -> Self {
        let mut resolver = Self::blank(4, 6);
        resolver.set(0, 0, FixtureCell::Text("Label".to_owned()));
        resolver.set(0, 1, FixtureCell::Text("Amount".to_owned()));
        for (row, (label, amount)) in [(" oak ", 1.0), ("oak", 2.0), (" oak", 3.0)]
            .into_iter()
            .enumerate()
        {
            resolver.set(row + 1, 0, FixtureCell::Text(label.to_owned()));
            resolver.set(row + 1, 1, FixtureCell::Number(amount));
        }
        resolver.set(0, 3, FixtureCell::Text("Label".to_owned()));
        resolver.set(1, 3, FixtureCell::Text(" oak ".to_owned()));
        resolver.set(2, 3, FixtureCell::Text("Label".to_owned()));
        resolver.set(3, 3, FixtureCell::Text("= oak ".to_owned()));
        resolver
    }

    fn error_header_database() -> Self {
        let mut resolver = Self::blank(4, 6);
        resolver.set(0, 0, FixtureCell::Text("Key".to_owned()));
        resolver.set(0, 1, FixtureCell::Error(ScalarError::NotAvailable));
        resolver.set(1, 0, FixtureCell::Text("wanted".to_owned()));
        resolver.set(1, 1, FixtureCell::Number(7.0));
        resolver.set(2, 0, FixtureCell::Text("other".to_owned()));
        resolver.set(2, 1, FixtureCell::Number(11.0));
        resolver.set(3, 0, FixtureCell::Text("ignored".to_owned()));
        resolver.set(3, 1, FixtureCell::Number(13.0));
        resolver.set(0, 3, FixtureCell::Text("Key".to_owned()));
        resolver.set(1, 3, FixtureCell::Text("wanted".to_owned()));
        resolver
    }

    fn blank_criteria_header() -> Self {
        let mut resolver = Self::blank(3, 6);
        resolver.set(0, 0, FixtureCell::Text("Key".to_owned()));
        resolver.set(0, 1, FixtureCell::Text("Amount".to_owned()));
        resolver.set(1, 0, FixtureCell::Text("wanted".to_owned()));
        resolver.set(1, 1, FixtureCell::Number(7.0));
        resolver.set(2, 0, FixtureCell::Text("other".to_owned()));
        resolver.set(2, 1, FixtureCell::Number(11.0));
        // A nonempty body expression without a Field selector must be a
        // typed Value error rather than an implicit match-all criterion.
        resolver.set(1, 3, FixtureCell::Text("wanted".to_owned()));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: FixtureCell) {
        let index = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column))
            .expect("fixture coordinate fits");
        self.cells[index] = value;
    }

    fn record(
        &mut self,
        row: usize,
        name: &str,
        region: &str,
        amount: FixtureCell,
        score: f64,
        active: bool,
    ) {
        self.set(row, 0, FixtureCell::Text(name.to_owned()));
        self.set(row, 1, FixtureCell::Text(region.to_owned()));
        self.set(row, 2, amount);
        self.set(row, 3, FixtureCell::Number(score));
        self.set(row, 4, FixtureCell::Logical(active));
    }

    fn criteria_pair(&mut self, first_column: usize, row: usize, left: &str, right: &str) {
        self.set(row, first_column, FixtureCell::Text(left.to_owned()));
        if right == "TRUE" {
            self.set(row, first_column + 1, FixtureCell::Logical(true));
        } else if right == "FALSE" {
            self.set(row, first_column + 1, FixtureCell::Logical(false));
        } else {
            self.set(row, first_column + 1, FixtureCell::Text(right.to_owned()));
        }
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }
}

impl Resolver for DatabaseResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<SheetExtent>> {
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
        if sheet != "Main" {
            return Ok(CellRead::Error(ScalarError::Reference));
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
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> EvaluationResult<Option<usize>> {
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

fn parse(source: &str) -> Expression {
    Expression::parse(source).unwrap_or_else(|error| panic!("{source:?} should parse: {error}"))
}

fn evaluate<'a>(
    expression: &'a Expression,
    resolver: &'a DatabaseResolver,
    execution: &ExecutionContext,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    value::evaluate(expression, resolver, &context, &ValueLimits::default())
}

fn assert_number(result: &Evaluated<'_>, expected: f64, source: &str) {
    match result.value() {
        Value::Number(actual) => {
            let tolerance = expected.abs().max(1.0) * 1.0e-12;
            assert!(
                (actual - expected).abs() <= tolerance,
                "{source:?}: expected {expected}, got {actual}"
            );
        },
        other => panic!("{source:?}: expected Number({expected}), got {other:?}"),
    }
}

fn assert_formula_error(result: &Evaluated<'_>, source: &str) {
    assert!(
        matches!(result.value(), Value::Error(_)),
        "{source:?}: expected a formula Error value, got {:?}",
        result.value()
    );
}

fn assert_formula_error_kind(result: &Evaluated<'_>, expected: ScalarError, source: &str) {
    match result.value() {
        Value::Error(actual) => assert_eq!(actual, expected, "{source:?}"),
        other => panic!("{source:?}: expected formula Error({expected}), got {other:?}"),
    }
}

#[test]
fn reference_database_runs_all_twelve_functions() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-reference-all");

    // East/TRUE OR West/FALSE selects Alice, Bob, Carol, Dave, and
    // EmptyAmount.  Numeric Amount values are 10, 20, and 30; Dave is text
    // and EmptyAmount is empty.  DGET intentionally has multiple matches.
    let cases = [
        ("DAVERAGE", Some(20.0)),
        ("DCOUNT", Some(3.0)),
        ("DCOUNTA", Some(4.0)),
        ("DGET", None),
        ("DMAX", Some(30.0)),
        ("DMIN", Some(10.0)),
        ("DPRODUCT", Some(6_000.0)),
        ("DSTDEV", Some(10.0)),
        ("DSTDEVP", Some((200.0_f64 / 3.0).sqrt())),
        ("DSUM", Some(60.0)),
        ("DVAR", Some(100.0)),
        ("DVARP", Some(200.0_f64 / 3.0)),
    ];
    for (function, expected) in cases {
        let source = format!("={function}({DATABASE};\"Amount\";{CRITERIA_OR})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        match expected {
            Some(expected) => assert_number(&result, expected, &source),
            None => assert_formula_error(&result, &source),
        }
    }
    assert!(
        resolver.reads() > 0,
        "reference database must read its source"
    );
}

#[test]
fn rectangular_inline_database_and_criteria_cover_the_same_numeric_family() {
    // Inline arrays are a rectangular Database/Criteria representation in
    // this value profile.  The source range path above remains the normative
    // required ODF range path; this case catches array-only implementations
    // that accidentally bypass the database pseudotype.
    const INLINE_DATABASE: &str = r#"{"Name";"Region";"Amount";"Score";"Active"|"Alice";"East";10;1;TRUE()|"Bob";"West";20;2;FALSE()|"Carol";"East";30;3;TRUE()|"Dave";"East";"n/a";4;TRUE()}"#;
    const INLINE_CRITERIA: &str = r#"{"Region";"Active"|"East";TRUE()}"#;
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-inline");
    let cases = [
        ("DAVERAGE", Some(20.0)),
        ("DCOUNT", Some(2.0)),
        ("DCOUNTA", Some(3.0)),
        ("DGET", None),
        ("DMAX", Some(30.0)),
        ("DMIN", Some(10.0)),
        ("DPRODUCT", Some(300.0)),
        ("DSTDEV", Some(200.0_f64.sqrt())),
        ("DSTDEVP", Some(10.0)),
        ("DSUM", Some(40.0)),
        ("DVAR", Some(200.0)),
        ("DVARP", Some(100.0)),
    ];
    for (function, expected) in cases {
        let source = format!("={function}({INLINE_DATABASE};\"Amount\";{INLINE_CRITERIA})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        match expected {
            Some(expected) => assert_number(&result, expected, &source),
            None => assert_formula_error(&result, &source),
        }
    }

    let source = format!("=DGET({INLINE_DATABASE};\"Amount\";{{\"Region\"|\"West\"}})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 20.0, &source);
    assert_eq!(
        resolver.reads(),
        0,
        "inline database must not read a worksheet"
    );
}

#[test]
fn field_selection_is_case_insensitive_one_based_and_type_aware() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-field");

    for field in ["\"amount\"", "3"] {
        let source = format!("=DSUM({DATABASE};{field};{CRITERIA_OR})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, 60.0, &source);
    }

    let source = format!("=DCOUNT({DATABASE};\"Name\";{CRITERIA_OR})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 0.0, &source);

    let source = format!("=DCOUNTA({DATABASE};\"Name\";{CRITERIA_OR})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 5.0, &source);

    let source = format!("=DSUM({DATABASE};\"NoSuchField\";{CRITERIA_OR})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_formula_error(&result, &source);
}

#[test]
fn numeric_field_selector_ignores_an_error_header_when_selecting_by_ordinal() {
    let resolver = DatabaseResolver::error_header_database();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-database-error-header-ordinal");

    // The second header is an Error value, but the numeric Field selector is
    // positional and therefore still selects the second data column.
    for (function, expected) in [("DSUM", 7.0), ("DCOUNT", 1.0)] {
        let source = format!("={function}({ERROR_HEADER_DATABASE};2;{ERROR_HEADER_CRITERIA})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, &source);
    }
}

#[test]
fn criteria_rows_are_or_and_fields_in_each_row_are_and() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-criteria");

    let cases = [
        ("DCOUNT", CRITERIA_AND, 2.0),
        ("DCOUNTA", CRITERIA_AND, 3.0),
        ("DCOUNT", CRITERIA_OR, 3.0),
        ("DCOUNTA", CRITERIA_OR, 4.0),
        ("DSUM", CRITERIA_OR, 60.0),
    ];
    for (function, criteria, expected) in cases {
        let source = format!("={function}({DATABASE};\"Amount\";{criteria})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, &source);
    }

    // A numeric comparator is a fixed Criterion rule and does not use the
    // host-dependent wildcard/substring matching branch.  It selects Bob,
    // Carol, and Erin, whose Amount sum is 20 + 30 + 40.
    let source = format!("=DSUM({DATABASE_WITHOUT_FAULT};\"Amount\";{CRITERIA_GREATER_THAN_15})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 90.0, &source);
}

#[test]
fn dcount_and_dcounta_support_omitted_fields() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-omitted-field");

    // With no Field argument both functions count matching records rather
    // than cells in a selected column.  The five records selected by OR are
    // therefore counted even when Amount is text or empty.
    for function in ["DCOUNT", "DCOUNTA"] {
        let source = format!("={function}({DATABASE};;{CRITERIA_OR})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, 5.0, &source);
    }
}

#[test]
fn dcount_and_dcounta_accept_the_two_argument_omitted_field_form() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-database-two-argument-omitted-field");

    // The two-argument spelling omits Field entirely.  It counts matching
    // records, just like the explicit missing-middle form above.
    for function in ["DCOUNT", "DCOUNTA"] {
        let source = format!("={function}({DATABASE};{CRITERIA_OR})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, 5.0, &source);
    }
}

#[test]
fn nonempty_expression_under_a_blank_criteria_header_is_value_error() {
    let resolver = DatabaseResolver::blank_criteria_header();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-database-blank-criteria-header");
    let source = format!("=DSUM({BLANK_CRITERIA_DATABASE};2;{BLANK_CRITERIA})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_formula_error_kind(&result, ScalarError::Value, &source);
}

#[test]
fn dget_requires_one_match_and_empty_criterion_is_distinct_from_equals_zero() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-dget");

    let source = format!("=DGET({DATABASE};\"Amount\";{CRITERIA_UNIQUE})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 20.0, &source);

    for criteria in [CRITERIA_NONE, CRITERIA_OR] {
        let source = format!("=DGET({DATABASE};\"Amount\";{criteria})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_formula_error(&result, &source);
    }

    // The empty criterion selects EmptyAmount and permits a unique DGET of
    // its Score.  The explicit =0 criterion selects no record because §4.11.8
    // expressly says that =0 does not match an empty cell.
    let source = format!("=DGET({DATABASE_WITHOUT_FAULT};\"Score\";{CRITERIA_EMPTY})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 7.0, &source);

    let source = format!("=DCOUNT({DATABASE_WITHOUT_FAULT};\"Amount\";{CRITERIA_ZERO})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 0.0, &source);
}

#[test]
fn database_formula_errors_propagate_and_selected_error_is_not_a_number() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-errors");

    let source = format!("=DSUM(#N/A;\"Amount\";{CRITERIA_AND})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_formula_error_kind(&result, ScalarError::NotAvailable, &source);

    let source = format!("=DSUM({DATABASE};\"Amount\";#DIV/0!)");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_formula_error_kind(&result, ScalarError::DivisionByZero, &source);

    // The Error-valued Amount is a selected record, so DGET returns that
    // formula error.  DCOUNT follows COUNT's explicit no-error-propagation
    // rule and counts no numeric value in that record.
    let source = format!("=DGET({DATABASE};\"Amount\";{CRITERIA_ERROR})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_formula_error_kind(&result, ScalarError::NotAvailable, &source);

    let source = format!("=DCOUNT({DATABASE};\"Amount\";{CRITERIA_ERROR})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 0.0, &source);
}

#[test]
fn numeric_field_selectors_use_position_even_with_duplicate_non_text_and_empty_headers() {
    let resolver = DatabaseResolver::duplicate_and_non_text_headers();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-database-header-field-profile");

    // The four selectors are positional 1..4.  Their header cells are,
    // respectively, duplicate Text, duplicate Text, Number, and Empty.  A
    // numeric selector must still address each selected column directly.
    for (field, expected) in [("1", 111.0), ("2", 222.0), ("3", 333.0), ("4", 444.0)] {
        let source = format!("=DSUM({PROFILE_DATABASE};{field};{PROFILE_CRITERIA})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, &source);
    }

    // Text lookup is deliberately different: two equal headers are an
    // ambiguous field selector and do not silently select the first column.
    let source = format!("=DSUM({PROFILE_DATABASE};\"Amount\";{PROFILE_CRITERIA})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_formula_error(&result, &source);
}

#[test]
fn numeric_criteria_headers_and_duplicate_criteria_columns_form_an_and_bound() {
    let resolver = DatabaseResolver::numeric_criteria_headers();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-database-numeric-criteria-headers");

    // Both criteria headers are Number(2), so both select Amount.  The two
    // expressions in one criteria row are AND'ed: >=10 and <=20 select the
    // two records with amounts 10 and 20.
    let source = format!("=DCOUNT([.A1:.B4];2;{CRITERIA_BOUNDS})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 2.0, &source);

    let source = format!("=DSUM([.A1:.B4];2;{CRITERIA_BOUNDS})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 30.0, &source);
}

#[test]
fn a_unique_empty_text_header_is_selectable_without_treating_an_empty_cell_header_as_text() {
    let resolver = DatabaseResolver::unique_empty_text_header();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-empty-text-header");

    // Header A1 is Text(""), while B1 is a genuinely Empty cell.  The empty
    // Text field selector resolves A1 uniquely and selects the first record.
    let source = format!("=DSUM({EMPTY_HEADER_DATABASE};\"\";[.F1:.F2])");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_number(&result, 5.0, &source);
}

#[test]
fn dget_converts_empty_logical_numeric_text_and_rejects_bad_text() {
    let resolver = DatabaseResolver::dget_conversions();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-dget-conversions");

    let cases = [
        (DGET_EMPTY_CRITERIA, Some(0.0)),
        (DGET_LOGICAL_CRITERIA, Some(1.0)),
        (DGET_NUMERIC_TEXT_CRITERIA, Some(12.5)),
        (DGET_BAD_TEXT_CRITERIA, None),
    ];
    for (criteria, expected) in cases {
        let source = format!("=DGET({DGET_DATABASE};\"Value\";{criteria})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        match expected {
            Some(expected) => assert_number(&result, expected, &source),
            None => assert_formula_error_kind(&result, ScalarError::Value, &source),
        }
    }
}

#[test]
fn empty_database_selection_uses_profiled_zero_sum_extrema_and_unit_product() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-database-empty-aggregate");

    for (function, expected) in [
        ("DSUM", 0.0),
        ("DMAX", 0.0),
        ("DMIN", 0.0),
        ("DPRODUCT", 1.0),
    ] {
        let source = format!("={function}({DATABASE};\"Amount\";{CRITERIA_NONE})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, &source);
    }
}

#[test]
fn iferror_can_return_a_reference_union_when_dget_fails() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-database-reference-valued-iferror");
    let source = format!("=IFERROR(DGET({DATABASE};\"Amount\";{CRITERIA_NONE});([.A1]~[.B1]))");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    let Value::ReferenceList(list) = result.value() else {
        panic!(
            "failed DGET should select the reference-union fallback, got {:?}",
            result.value()
        );
    };
    assert_eq!(list.len(), 2);
    assert_eq!(
        list.get(0).expect("first union record").areas()[0].starts(),
        [0, 0, 0]
    );
    assert_eq!(
        list.get(1).expect("second union record").areas()[0].starts(),
        [0, 0, 1]
    );
}

#[test]
fn successful_dget_number_cannot_be_used_as_a_reference_operand() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-database-dget-reference-boundary");
    let source = format!("=DGET({DATABASE};\"Amount\";{CRITERIA_UNIQUE}):[.A1]");
    let expression = parse(&source);
    let error = evaluate(&expression, &resolver, &execution)
        .expect_err("a successful numeric DGET is not a reference operand");
    assert!(
        matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator)
        ),
        "wrong DGET/reference boundary error: {error:?}"
    );
}

#[test]
fn criteria_text_preserves_leading_and_trailing_whitespace() {
    // Both forms must compare the complete criterion operand.  In the second
    // form `=` is the comparison operator and the spaces around `oak` remain
    // part of the text value; neither form may silently select the unpadded
    // or one-sidedly padded records.
    const INLINE_DATABASE: &str = r#"{"Label";"Amount"|" oak ";1|"oak";2|" oak";3}"#;
    const INLINE_LITERAL_CRITERIA: &str = r#"{"Label"|" oak "}"#;
    const INLINE_EQUALS_CRITERIA: &str = r#"{"Label"|"= oak "}"#;

    let resolver = DatabaseResolver::whitespace_criteria();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-database-criterion-whitespace");

    for (database, literal_criteria, equals_criteria) in [
        (
            WHITESPACE_DATABASE,
            WHITESPACE_LITERAL_CRITERIA,
            WHITESPACE_EQUALS_CRITERIA,
        ),
        (
            INLINE_DATABASE,
            INLINE_LITERAL_CRITERIA,
            INLINE_EQUALS_CRITERIA,
        ),
    ] {
        let source = format!("=DSUM({database};\"Amount\";{literal_criteria})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, 1.0, &source);

        let source = format!("=DSUM({database};\"Amount\";{equals_criteria})");
        let expression = parse(&source);
        let result = evaluate(&expression, &resolver, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, 1.0, &source);
    }
}

#[test]
fn computed_database_and_criteria_missing_broadcast_cells_preserve_not_available() {
    let resolver = DatabaseResolver::standard();
    let (_budget, _cancellation, execution) =
        execution("ods-formula-database-computed-missing-cells");

    // The three-row, two-column condition selects a two-row database branch.
    // Its third row is therefore an out-of-shape broadcast position.  This
    // is produced by the value VM's IF/broadcast machinery, rather than by
    // an explicit #N/A literal, and must remain the formula error #N/A when
    // the database scan reaches that row.
    let source =
        r#"=DSUM(IF({TRUE();TRUE()|TRUE();TRUE()|TRUE();TRUE()};{"Value";"Take"|7;1};0);1;{2|1})"#;
    let expression = parse(source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_formula_error_kind(&result, ScalarError::NotAvailable, source);

    // The same shape mismatch in the criteria argument must also be an
    // out-of-shape #N/A.  A conversion to Empty would instead make the query
    // silently produce an ordinary aggregate result.
    let source =
        r#"=DSUM({"Value";"Take"|7;1};1;IF({TRUE();TRUE()|TRUE();TRUE()|TRUE();TRUE()};{2|1};0))"#;
    let expression = parse(source);
    let result = evaluate(&expression, &resolver, &execution)
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
    assert_formula_error_kind(&result, ScalarError::NotAvailable, source);
}

#[test]
fn empty_criterion_reference_converts_to_zero_without_matching_empty_records() {
    let mut resolver = DatabaseResolver::standard();
    resolver.set(1, 2, FixtureCell::Number(0.0));
    resolver.set(4, 18, FixtureCell::Empty);
    let (_budget, _cancellation, execution) = execution("ods-database-empty-criterion-zero");
    let source = format!("=DCOUNT({DATABASE_WITHOUT_FAULT};3;{CRITERIA_ZERO})");
    let expression = parse(&source);
    let result = evaluate(&expression, &resolver, &execution).expect("empty criterion reference");
    assert_number(&result, 1.0, &source);
}
