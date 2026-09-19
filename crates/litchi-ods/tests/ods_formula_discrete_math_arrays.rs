//! Independent value-VM coverage for the OpenFormula 1.4 §6.16 discrete
//! mathematical functions.
//!
//! Scalar kernels are covered separately.  The cases here exercise the
//! observable type signatures: scalar functions broadcast in matrix mode,
//! GCD/LCM consume `NumberSequenceList`, MULTINOMIAL consumes
//! `NumberSequence`, reference cells omit Text/Logical/Empty values while
//! inline arrays convert them, and evaluator limits remain typed failures.

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
        self, CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{EvaluationFailure, ScalarError},
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
struct CellEntry {
    sheet: String,
    row: usize,
    column: usize,
    value: FixtureCell,
}

#[derive(Debug)]
struct DiscreteResolver {
    rows: usize,
    columns: usize,
    cells: Vec<CellEntry>,
    reads: AtomicUsize,
    read_order: Mutex<Vec<(String, usize, usize)>>,
}

impl DiscreteResolver {
    fn blank(rows: usize, columns: usize) -> Self {
        Self {
            rows,
            columns,
            cells: Vec::new(),
            reads: AtomicUsize::new(0),
            read_order: Mutex::new(Vec::new()),
        }
    }

    fn set(&mut self, sheet: &str, row: usize, column: usize, value: FixtureCell) {
        self.cells.push(CellEntry {
            sheet: sheet.to_owned(),
            row,
            column,
            value,
        });
    }

    fn set_main(&mut self, row: usize, column: usize, value: FixtureCell) {
        self.set("Main", row, column, value);
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Acquire)
    }

    fn read_order(&self) -> Vec<(String, usize, usize)> {
        self.read_order.lock().expect("read-order lock").clone()
    }
}

impl Resolver for DiscreteResolver {
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
            .expect("read-order lock")
            .push((sheet.to_owned(), row, column));
        if !matches!(sheet, "Main" | "Data" | "Archive")
            || row >= self.rows
            || column >= self.columns
        {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        let Some(entry) = self
            .cells
            .iter()
            .find(|entry| entry.sheet == sheet && entry.row == row && entry.column == column)
        else {
            return Ok(CellRead::Empty);
        };
        Ok(match &entry.value {
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

    fn sheet_name_at<'a>(
        &'a self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&'a str>, EvaluationFailure> {
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
    resolver: &'a DiscreteResolver,
    execution: &ExecutionContext,
    position: Position<'a>,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_array(
    result: &Evaluated<'_>,
    rows: usize,
    columns: usize,
    expected: &[Result<f64, ScalarError>],
    source: &str,
) {
    let array = result
        .as_array()
        .unwrap_or_else(|| panic!("{source:?} should return an array"));
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (rows, columns),
        "{source:?} shape"
    );
    assert_eq!(array.len(), expected.len(), "{source:?} length");
    for (index, expected) in expected.iter().enumerate() {
        let actual = array
            .get(index)
            .unwrap_or_else(|| panic!("{source:?} missing cell {index}"));
        match (actual, expected) {
            (Value::Number(actual), Ok(expected)) => {
                if *expected == 0.0 {
                    assert_eq!(actual.to_bits(), expected.to_bits(), "{source:?}[{index}]");
                } else {
                    let tolerance = actual.abs().max(expected.abs()) * (16.0 * f64::EPSILON);
                    assert!(
                        (actual - expected).abs() <= tolerance,
                        "{source:?}[{index}]: {actual:?} != {expected:?}"
                    );
                }
            },
            (Value::Error(actual), Err(expected)) => {
                assert_eq!(actual, *expected, "{source:?}[{index}]");
            },
            (actual, expected) => {
                panic!("{source:?}[{index}] returned {actual:?}, expected {expected:?}");
            },
        }
    }
}

fn assert_number(result: &Evaluated<'_>, expected: f64, source: &str) {
    match result.value() {
        Value::Number(actual) => {
            let tolerance = expected.abs().max(actual.abs()) * (16.0 * f64::EPSILON);
            assert!(
                (actual - expected).abs() <= tolerance,
                "{source:?}: {actual:?} != {expected:?}"
            );
        },
        value => panic!("{source:?}: expected Number({expected}), got {value:?}"),
    }
}

fn assert_error(result: &Evaluated<'_>, expected: ScalarError, source: &str) {
    match result.value() {
        Value::Error(actual) => assert_eq!(actual, expected, "{source:?}"),
        value => panic!("{source:?}: expected Error({expected:?}), got {value:?}"),
    }
}

#[test]
fn scalar_discrete_functions_broadcast_over_literal_matrices() {
    let resolver = DiscreteResolver::blank(8, 8);
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-array-broadcast");
    let cases: [(&str, [Result<f64, ScalarError>; 4]); 11] = [
        (
            "=COMBIN({5;6|7;8};{2;2|3;3})",
            [Ok(10.0), Ok(15.0), Ok(35.0), Ok(56.0)],
        ),
        (
            "=COMBINA({2;3|4;5};2)",
            [Ok(3.0), Ok(6.0), Ok(10.0), Ok(15.0)],
        ),
        ("=FACT({0;1|2;3})", [Ok(1.0), Ok(1.0), Ok(2.0), Ok(6.0)]),
        (
            "=FACTDOUBLE({0;1|2;3})",
            [Ok(1.0), Ok(1.0), Ok(2.0), Ok(3.0)],
        ),
        (
            "=EVEN({-2.5;-1|0.5;2.5})",
            [Ok(-4.0), Ok(-2.0), Ok(2.0), Ok(4.0)],
        ),
        (
            "=ODD({-2.5;-1|0.5;2.5})",
            [Ok(-3.0), Ok(-1.0), Ok(1.0), Ok(3.0)],
        ),
        (
            "=DELTA({1;2|3;4};{1;1|4;3})",
            [Ok(1.0), Ok(0.0), Ok(0.0), Ok(0.0)],
        ),
        (
            "=GESTEP({-1;0|1;2};{0;0|2;1})",
            [Ok(0.0), Ok(1.0), Ok(0.0), Ok(1.0)],
        ),
        ("=DELTA({0;1|2;3})", [Ok(1.0), Ok(0.0), Ok(0.0), Ok(0.0)]),
        ("=GESTEP({-1;0|1;2})", [Ok(0.0), Ok(1.0), Ok(1.0), Ok(1.0)]),
        (
            "=COMBIN({5;#N/A|7;8};2)",
            [Ok(10.0), Err(ScalarError::NotAvailable), Ok(21.0), Ok(28.0)],
        ),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_array(&result, 2, 2, &expected, source);
    }
    assert_eq!(
        resolver.reads(),
        0,
        "literal arrays need no worksheet reads"
    );
}

#[test]
fn matrix_scalar_functions_convert_inline_text_logical_and_formula_errors_per_cell() {
    let resolver = DiscreteResolver::blank(2, 2);
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-array-coercion");
    let cases: [(&str, [Result<f64, ScalarError>; 4]); 2] = [
        (
            "=FACT({\"5\";TRUE()|\"\";#N/A})",
            [
                Ok(120.0),
                Ok(1.0),
                Err(ScalarError::Value),
                Err(ScalarError::NotAvailable),
            ],
        ),
        (
            "=EVEN({\"2.5\";TRUE()|\"\";#DIV/0!})",
            [
                Ok(4.0),
                Ok(2.0),
                Err(ScalarError::Value),
                Err(ScalarError::DivisionByZero),
            ],
        ),
    ];
    for (source, expected) in cases {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_array(&result, 2, 2, &expected, source);
    }
}

#[test]
fn reference_sequences_omit_non_numeric_cells_but_inline_arrays_convert_them() {
    let mut resolver = DiscreteResolver::blank(8, 8);
    resolver.set_main(0, 0, FixtureCell::Number(48.0));
    resolver.set_main(1, 0, FixtureCell::Number(18.0));
    resolver.set_main(2, 0, FixtureCell::Text("7".to_owned()));
    resolver.set_main(3, 0, FixtureCell::Empty);
    resolver.set_main(4, 0, FixtureCell::Logical(true));
    resolver.set_main(0, 1, FixtureCell::Number(2.0));
    resolver.set_main(1, 1, FixtureCell::Number(3.0));
    resolver.set_main(2, 1, FixtureCell::Text("2".to_owned()));
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-reference-kinds");

    for (source, expected) in [
        ("=GCD([.A1:.A5])", 6.0),
        ("=LCM([.A1:.A5])", 144.0),
        ("=GCD({48;18;\"7\";TRUE()})", 1.0),
        ("=LCM({48;18;\"7\";TRUE()})", 1008.0),
        ("=MULTINOMIAL([.B1:.B3])", 10.0),
        ("=MULTINOMIAL({2;3;\"2\"})", 210.0),
    ] {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, source);
    }
    resolver.set_main(5, 0, FixtureCell::Error(ScalarError::NotAvailable));
    let source = "=GCD([.A1:.A6])";
    let expression = parse(source);
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("{source:?} should publish its formula error: {error}"));
    assert_error(&result, ScalarError::NotAvailable, source);
    assert_eq!(
        resolver.reads(),
        19,
        "each reference sequence cell is read once"
    );
}

#[test]
fn value_generated_errors_follow_sequence_source_order() {
    let resolver = DiscreteResolver::blank(2, 2);
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-generated-order");
    for (source, expected) in [
        ("=LCM(0.5;\"bad\")", ScalarError::Number),
        ("=LCM(\"bad\";0.5)", ScalarError::Value),
        ("=GCD(-1;\"bad\")", ScalarError::Number),
        ("=GCD(\"bad\";-1)", ScalarError::Value),
        ("=MULTINOMIAL(-1;\"bad\")", ScalarError::Number),
        ("=MULTINOMIAL(\"bad\";-1)", ScalarError::Value),
    ] {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should publish a formula error: {error}"));
        assert_error(&result, expected, source);
    }
    assert_eq!(
        resolver.reads(),
        0,
        "scalar sequence arguments need no reads"
    );
}

#[test]
fn empty_reference_sequences_use_the_selected_identities() {
    let resolver = DiscreteResolver::blank(8, 8);
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-empty-sequences");
    for (source, expected) in [
        ("=GCD([.A1:.A2])", Err(ScalarError::Number)),
        ("=LCM([.A1:.A2])", Err(ScalarError::Number)),
        ("=MULTINOMIAL([.A1:.A2])", Ok(1.0)),
    ] {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        match expected {
            Ok(expected) => assert_number(&result, expected, source),
            Err(expected) => assert_error(&result, expected, source),
        }
    }
    assert_eq!(
        resolver.reads(),
        6,
        "each empty reference is admitted and read"
    );
}

#[test]
fn scalar_reference_empty_cells_use_the_number_conversion_bridge() {
    let resolver = DiscreteResolver::blank(2, 2);
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-empty-cell");
    for (source, expected) in [
        ("=FACT([.A1])", [Ok(1.0)]),
        ("=EVEN([.A1])", [Ok(0.0)]),
        ("=DELTA([.A1])", [Ok(1.0)]),
        ("=GESTEP([.A1])", [Ok(1.0)]),
    ] {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_array(&result, 1, 1, &expected, source);
    }
    assert_eq!(resolver.reads(), 4, "each scalar reference is read once");
}

#[test]
fn reference_lists_and_three_dimensional_cuboids_follow_exact_signatures() {
    let mut resolver = DiscreteResolver::blank(8, 8);
    resolver.set_main(0, 0, FixtureCell::Number(48.0));
    resolver.set_main(1, 0, FixtureCell::Number(18.0));
    resolver.set_main(0, 1, FixtureCell::Number(30.0));
    resolver.set_main(1, 1, FixtureCell::Number(6.0));
    resolver.set_main(0, 2, FixtureCell::Number(2.0));
    resolver.set_main(0, 3, FixtureCell::Number(3.0));
    resolver.set("Data", 0, 0, FixtureCell::Number(18.0));
    resolver.set("Data", 0, 2, FixtureCell::Number(3.0));
    resolver.set("Data", 0, 3, FixtureCell::Number(3.0));
    resolver.set("Archive", 0, 0, FixtureCell::Number(30.0));
    resolver.set("Archive", 0, 2, FixtureCell::Number(4.0));
    resolver.set("Archive", 0, 3, FixtureCell::Number(4.0));
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-reference-shapes");

    for (source, expected) in [
        ("=GCD([.A1:.A2]~[.B1:.B2]~[.A1])", 6.0),
        ("=LCM([.A1:.A2]~[.B1:.B2])", 720.0),
        ("=GCD([Main.A1:Archive.A1])", 6.0),
        ("=LCM([Main.A1:Archive.A1])", 720.0),
        ("=MULTINOMIAL([Main.C1:Archive.C1])", 1260.0),
    ] {
        let expression = parse(source);
        let result = evaluate(
            &expression,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_number(&result, expected, source);
    }

    // MULTINOMIAL is NumberSequence, so an explicit ReferenceList is a
    // formula-value shape error and must be rejected before any provider read.
    let before = resolver.reads();
    let source = "=MULTINOMIAL([.C1]~[.D1])";
    let expression = parse(source);
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("ReferenceList rejection is a formula value");
    assert_error(&result, ScalarError::Value, source);
    assert_eq!(
        resolver.reads(),
        before,
        "shape rejection must be read-free"
    );

    let before = resolver.reads();
    let source = "=MULTINOMIAL(([.A1]~[.A2]);#N/A)";
    let expression = parse(source);
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a later direct formula error must remain a formula value");
    assert_error(&result, ScalarError::NotAvailable, source);
    assert_eq!(
        resolver.reads(),
        before,
        "inadmitted ReferenceList must be rejected before provider reads"
    );

    let mut ordered = DiscreteResolver::blank(8, 8);
    ordered.set_main(0, 0, FixtureCell::Number(1.0));
    ordered.set_main(1, 0, FixtureCell::Number(2.0));
    ordered.set_main(0, 1, FixtureCell::Number(5.0));
    ordered.set_main(1, 1, FixtureCell::Number(6.0));
    let source = "=GCD([.A1:.A2]~[.B1:.B2]~[.A1])";
    let expression = parse(source);
    let result = evaluate(
        &expression,
        &ordered,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("{source:?} should preserve list order: {error}"));
    assert_number(&result, 1.0, source);
    assert_eq!(
        ordered.read_order(),
        vec![
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 1, 0),
            ("Main".to_owned(), 0, 1),
            ("Main".to_owned(), 1, 1),
            ("Main".to_owned(), 0, 0),
        ],
        "ReferenceList areas and duplicate cells retain source order"
    );

    let mut cuboid = DiscreteResolver::blank(8, 8);
    cuboid.set("Main", 0, 0, FixtureCell::Number(1.0));
    cuboid.set("Data", 0, 0, FixtureCell::Number(2.0));
    cuboid.set("Archive", 0, 0, FixtureCell::Number(3.0));
    let source = "=GCD([Main.A1:Archive.A1])";
    let expression = parse(source);
    let result = evaluate(
        &expression,
        &cuboid,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("{source:?} should preserve 3-D order: {error}"));
    assert_number(&result, 1.0, source);
    assert_eq!(
        cuboid.read_order(),
        vec![
            ("Main".to_owned(), 0, 0),
            ("Data".to_owned(), 0, 0),
            ("Archive".to_owned(), 0, 0),
        ],
        "3-D sheets retain provider order"
    );
}

#[test]
fn reducers_cache_projection_and_keep_unselected_provider_branches_inert() {
    let mut resolver = DiscreteResolver::blank(8, 8);
    resolver.set_main(0, 0, FixtureCell::Number(6.0));
    resolver.set_main(1, 0, FixtureCell::Number(10.0));
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-projection");

    let source = "=IF({TRUE();TRUE()};GCD([.A1:.A2]);GCD([Missing.A1:.Z100]))";
    let expression = parse(source);
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("{source:?} should evaluate lazily: {error}"));
    assert_array(&result, 1, 2, &[Ok(2.0), Ok(2.0)], source);
    assert_eq!(
        resolver.reads(),
        2,
        "the selected reducer range is cached once"
    );

    let source = "=COMBIN([.A1:.A2];2)";
    let expression = parse(source);
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("{source:?} should broadcast: {error}"));
    assert_array(&result, 2, 1, &[Ok(15.0), Ok(45.0)], source);
    assert_eq!(
        resolver.reads(),
        4,
        "broadcast projection reads each source cell once"
    );

    let mut scalar_resolver = DiscreteResolver::blank(8, 8);
    scalar_resolver.set_main(0, 0, FixtureCell::Number(6.0));
    scalar_resolver.set_main(1, 0, FixtureCell::Number(10.0));
    let source = "=GCD([.A1:.A2])";
    let expression = parse(source);
    let result = evaluate(
        &expression,
        &scalar_resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Scalar,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("{source:?} should reduce in scalar mode: {error}"));
    assert_number(&result, 2.0, source);
    assert_eq!(
        scalar_resolver.reads(),
        2,
        "a reducer must not implicitly intersect its sequence"
    );
}

#[test]
fn value_budget_failures_are_typed_uncatchable_and_refund_storage() {
    let resolver = DiscreteResolver::blank(8, 8);
    let expression = parse("=GCD({1;2|3;4})");
    let (budget, _cancellation, execution) = execution("ods-formula-discrete-array-work");
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default().with_max_steps(0),
    )
    .expect_err("zero work must refuse before publishing a reducer result");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Work),
        "unexpected work refusal: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let expression = parse("=IFERROR(GCD({1;2|3;4});7)");
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default().with_max_storage_bytes(0),
    )
    .expect_err("IFERROR must not catch a value-VM storage refusal");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Memory),
        "IFERROR changed the typed outcome: {error:?}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);

    let mut large = DiscreteResolver::blank(64, 1);
    for row in 0..64 {
        large.set_main(row, 0, FixtureCell::Number(1.0));
    }
    let expression = parse("=GCD([.A1:.A64])");
    let error = evaluate(
        &expression,
        &large,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default().with_max_reference_cells(63),
    )
    .expect_err("reference admission must fail before the 64th cell");
    assert!(
        matches!(&error, EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects),
        "wrong reference-cap failure: {error:?}"
    );
    assert_eq!(
        large.reads(),
        0,
        "geometry refusal must precede provider reads"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn unsupported_provider_cells_remain_typed_capability_failures() {
    let mut resolver = DiscreteResolver::blank(2, 2);
    resolver.set_main(0, 0, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) = execution("ods-formula-discrete-provider-error");
    let source = "=GCD([.A1])";
    let expression = parse(source);
    let error = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("unsupported provider cells must remain typed failures");
    assert!(
        matches!(
            error,
            EvaluationFailure::Unsupported(
                litchi_ods::codec::formula::evaluation::UnsupportedKind::CellValue
            )
        ),
        "wrong provider capability failure: {error:?}"
    );
    assert_eq!(resolver.reads(), 1);

    // A missing-sheet reference is a formula #REF! value and remains
    // catchable by IFERROR; provider capability failures are a different
    // typed channel.
    let source = "=IFERROR(GCD([Missing.A1]);9)";
    let expression = parse(source);
    let result = evaluate(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("{source:?} should catch its formula reference error: {error}"));
    assert_number(&result, 9.0, source);
    assert_eq!(
        resolver.reads(),
        1,
        "the missing sheet requires no provider read"
    );
}
