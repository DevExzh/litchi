//! Independent semantic coverage for the OpenFormula descriptive reducers:
//! `AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`, `KURT`, `SKEW`, and `SKEWP`.
//!
//! The fixture keeps references typed and ordered.  It therefore covers the
//! difference between `NumberSequence` and `NumberSequenceList`, 3-D ranges,
//! projected conditional criteria, and both value-evaluator modes without
//! depending on a host workbook or cached formula cells.

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
    },
    expression::Expression,
};

const NUMBERS: &str = "[.B1:.B4]";

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct DescriptiveResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl DescriptiveResolver {
    fn standard() -> Self {
        let mut resolver = Self {
            rows: 16,
            columns: 12,
            cells: (0..192).map(|_| FixtureCell::Empty).collect(),
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            source_versions: None,
            source_version_calls: Cell::new(0),
        };

        // A1:A2 provide the Main-plane cells of the six-value 3-D vector;
        // Data and Archive supply 3..6 in read_cell below.
        resolver.set(0, 0, FixtureCell::Number(1.0));
        resolver.set(1, 0, FixtureCell::Number(2.0));
        // B1:B4 is the canonical positive vector {1, 2, 3, 4}.
        for (row, value) in [1.0, 2.0, 3.0, 4.0].into_iter().enumerate() {
            resolver.set(row, 1, FixtureCell::Number(value));
        }
        // D1:E2 form two areas whose concatenation is the same vector.
        for (row, value) in [1.0, 2.0].into_iter().enumerate() {
            resolver.set(row, 3, FixtureCell::Number(value));
        }
        for (row, value) in [3.0, 4.0].into_iter().enumerate() {
            resolver.set(row, 4, FixtureCell::Number(value));
        }
        // C is used for error/omission and projected-criterion cases.
        for (row, value) in [1.0, 2.0, 3.0, 4.0].into_iter().enumerate() {
            resolver.set(row, 2, FixtureCell::Number(value));
        }
        resolver.set(4, 2, FixtureCell::Logical(true));
        resolver.set(5, 2, FixtureCell::Text("5".to_owned()));
        resolver.set(6, 2, FixtureCell::Empty);
        resolver.set(7, 2, FixtureCell::Error(ScalarError::NotAvailable));
        // G1:G2 are projected as scalar MUNIT arguments in the conditional
        // criterion test below.
        resolver.set(0, 6, FixtureCell::Number(1.0));
        resolver.set(1, 6, FixtureCell::Number(2.0));
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

impl Resolver for DescriptiveResolver {
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

        if sheet != "Main" {
            // The 3-D fixture is six ordered values, two on each sheet.
            // Other cells on Data and Archive are intentionally empty.
            if column == 0 && row < 2 {
                return Ok(match sheet {
                    "Data" => CellRead::Number((row + 3) as f64),
                    "Archive" => CellRead::Number((row + 5) as f64),
                    _ => unreachable!(),
                });
            }
            return Ok(CellRead::Empty);
        }
        Ok(match &self.cells[row * self.columns + column] {
            FixtureCell::Empty => CellRead::Empty,
            FixtureCell::Number(value) => CellRead::Number(*value),
            FixtureCell::Logical(value) => CellRead::Logical(*value),
            FixtureCell::Text(value) => CellRead::Text(value.as_str()),
            FixtureCell::Error(error) => CellRead::Error(*error),
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
enum ResultValue {
    Number(f64),
    Error(ScalarError),
    Other,
}

fn evaluate_scalar_source(
    source: &str,
    execution: &ExecutionContext,
) -> Result<ResultValue, EvaluationFailure> {
    let expression = parse(source);
    let result = litchi_ods::codec::formula::evaluation::evaluate_scalar(
        &expression,
        &EvaluationContext::new(execution),
        &EvaluationLimits::default(),
    )?;
    Ok(match result.value() {
        ScalarValue::Number(value) => ResultValue::Number(*value),
        ScalarValue::Error(error) => ResultValue::Error(*error),
        _ => ResultValue::Other,
    })
}

fn evaluate_value_source(
    source: &str,
    resolver: &DescriptiveResolver,
    execution: &ExecutionContext,
    mode: Mode,
) -> Result<ResultValue, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, &Limits::default())?;
    Ok(match result.value() {
        Value::Number(value) => ResultValue::Number(value),
        Value::Error(error) => ResultValue::Error(error),
        _ => ResultValue::Other,
    })
}

fn evaluate_value<'a>(
    expression: &'a Expression,
    resolver: &'a DescriptiveResolver,
    execution: &ExecutionContext,
    position: Position<'a>,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'a>, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    value::evaluate(expression, resolver, &context, limits)
}

fn assert_close(result: ResultValue, expected: f64, source: &str) {
    let ResultValue::Number(actual) = result else {
        panic!("{source:?}: expected Number({expected}), got {result:?}");
    };
    assert!(actual.is_finite(), "{source:?}: non-finite result {actual}");
    let tolerance = expected.abs().max(actual.abs()).max(1.0) * 128.0 * f64::EPSILON;
    assert!(
        (actual - expected).abs() <= tolerance,
        "{source:?}: {actual} != {expected} (tolerance {tolerance})"
    );
}

fn assert_error(result: ResultValue, expected: ScalarError, source: &str) {
    assert_eq!(result, ResultValue::Error(expected), "{source:?}");
}

fn assert_any_error(result: ResultValue, source: &str) {
    assert!(
        matches!(result, ResultValue::Error(_)),
        "{source:?}: expected a formula error, got {result:?}"
    );
}

fn assert_array_numbers(result: &Evaluated<'_>, expected: &[f64], source: &str) {
    let array = result
        .as_array()
        .unwrap_or_else(|| panic!("{source:?}: expected an array result"));
    assert_eq!(array.len(), expected.len(), "{source:?} array length");
    for (index, expected) in expected.iter().copied().enumerate() {
        match array.get(index) {
            Some(Value::Number(actual)) => assert_close(
                ResultValue::Number(actual),
                expected,
                &format!("{source:?}[{index}]"),
            ),
            other => panic!("{source:?}[{index}]: expected Number, got {other:?}"),
        }
    }
}

#[test]
fn scalar_descriptive_reducers_cover_centering_domains_and_variants() {
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-scalar");
    for (source, expected) in [
        ("=AVEDEV(1;2;3;4)", 1.0),
        ("=DEVSQ(1;2;3;4)", 5.0),
        ("=GEOMEAN(1;2;3;4)", 24.0_f64.powf(0.25)),
        ("=HARMEAN(1;2;3;4)", 1.92),
        ("=KURT(1;2;3;4)", -1.2),
        ("=SKEW(1;2;3;4)", 0.0),
        ("=SKEWP(1;2;3;4)", 0.0),
        ("=SKEW(1;3;4;5;9)", 0.8848873194830047),
        ("=SKEWP(1;3;4;5;9)", 0.5936004596374717),
        ("=AVEDEV(\"3\";1;TRUE();FALSE())", 0.875),
        ("=DEVSQ(\"3\";1;TRUE();FALSE())", 4.75),
    ] {
        assert_close(
            evaluate_scalar_source(source, &execution)
                .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}")),
            expected,
            source,
        );
    }
}

#[test]
fn scalar_descriptive_reducers_enforce_arity_and_variance_constraints() {
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-invalid");
    for source in [
        "=AVEDEV()",
        "=DEVSQ()",
        "=GEOMEAN()",
        "=HARMEAN()",
        "=KURT()",
        "=SKEW()",
        "=SKEWP()",
        "=KURT(1;2;3)",
        "=KURT(1;1;1;1)",
        "=SKEW(1;2)",
        "=SKEWP(1;2)",
        "=SKEW(1;1;1)",
        "=SKEWP(1;1;1)",
    ] {
        assert_any_error(
            evaluate_scalar_source(source, &execution).unwrap_or_else(|error| {
                panic!("{source:?} should produce a formula value: {error}")
            }),
            source,
        );
    }
}

#[test]
fn scalar_descriptive_reducers_retain_formula_error_precedence() {
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-errors");
    for source in [
        "=AVEDEV(#N/A;1)",
        "=DEVSQ(#N/A;1)",
        "=GEOMEAN(#N/A;1)",
        "=HARMEAN(#N/A;1)",
        "=KURT(#N/A;1;2;3)",
        "=SKEW(#N/A;1;2)",
        "=SKEWP(#N/A;1;2)",
    ] {
        assert_error(
            evaluate_scalar_source(source, &execution)
                .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}")),
            ScalarError::NotAvailable,
            source,
        );
    }
}

#[test]
fn geomean_and_harmean_follow_the_signed_and_zero_profile() {
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-domain");
    for (source, expected) in [
        ("=GEOMEAN(-2;-8)", 4.0),
        ("=GEOMEAN(-8)", -8.0),
        ("=GEOMEAN(-1;-8;-27)", -6.0),
        ("=GEOMEAN(0;2;8)", 0.0),
        ("=HARMEAN(-2;-4)", -8.0 / 3.0),
        ("=HARMEAN(-2;4)", -8.0),
        ("=HARMEAN(3;6;-2)", f64::NAN),
    ] {
        if expected.is_nan() {
            assert_error(
                evaluate_scalar_source(source, &execution).unwrap(),
                ScalarError::DivisionByZero,
                source,
            );
        } else {
            assert_close(
                evaluate_scalar_source(source, &execution).unwrap(),
                expected,
                source,
            );
        }
    }
    assert_error(
        evaluate_scalar_source("=GEOMEAN(-1;2)", &execution).unwrap(),
        ScalarError::Number,
        "=GEOMEAN(-1;2)",
    );
    for source in ["=HARMEAN(0;2)", "=HARMEAN(2;-2)", "=HARMEAN(3;6;-2)"] {
        assert_error(
            evaluate_scalar_source(source, &execution).unwrap(),
            ScalarError::DivisionByZero,
            source,
        );
    }
    // The cancellation case deliberately uses a very small reciprocal pair;
    // it must not be replaced with a tolerance-based zero test.
    assert_close(
        evaluate_scalar_source("=HARMEAN(1e-100;-1e-100;1e100)", &execution).unwrap(),
        3e100,
        "=HARMEAN(1e-100;-1e-100;1e100)",
    );
}

#[test]
fn signed_and_zero_domains_are_preserved_for_value_arrays_in_both_modes() {
    let resolver = DescriptiveResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-domain-arrays");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=GEOMEAN({-2;-8})", 4.0),
            ("=GEOMEAN({-8})", -8.0),
            ("=GEOMEAN({-1;-8;-27})", -6.0),
            ("=GEOMEAN({0;2;8})", 0.0),
            ("=HARMEAN({-2;-4})", -8.0 / 3.0),
            ("=HARMEAN({-2;4})", -8.0),
        ] {
            assert_close(
                evaluate_value_source(source, &resolver, &execution, mode)
                    .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
                expected,
                source,
            );
        }
        for (source, expected) in [
            ("=GEOMEAN({-1;2})", ScalarError::Number),
            ("=HARMEAN({0;2})", ScalarError::DivisionByZero),
            ("=HARMEAN({2;-2})", ScalarError::DivisionByZero),
        ] {
            assert_error(
                evaluate_value_source(source, &resolver, &execution, mode).unwrap(),
                expected,
                source,
            );
        }
    }
}

#[test]
fn value_arrays_are_reduced_in_both_modes() {
    let resolver = DescriptiveResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-arrays");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=AVEDEV({1;2;4;8})", 2.25),
            ("=DEVSQ({1;2;4;8})", 28.75),
            ("=GEOMEAN({1;2;4;8})", 2.8284271247461903),
            ("=HARMEAN({1;2;4;8})", 2.1333333333333333),
            ("=KURT({1;2;4;8})", 0.7576559546313799),
            ("=SKEW({1;2;4;8})", 1.1376243669576884),
            ("=SKEWP({1;2;4;8})", 0.6568077344996993),
        ] {
            assert_close(
                evaluate_value_source(source, &resolver, &execution, mode)
                    .unwrap_or_else(|error| panic!("{source:?} in {mode:?}: {error}")),
                expected,
                source,
            );
        }
        let result =
            evaluate_value_source("=AVEDEV({1;2|4;8})", &resolver, &execution, mode).unwrap();
        assert_close(result, 2.25, "=AVEDEV({1;2|4;8})");
    }
}

#[test]
fn references_lists_and_three_dimensional_ranges_follow_signatures() {
    let resolver = DescriptiveResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-references");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (function, expected) in [
            ("AVEDEV", 1.0),
            ("DEVSQ", 5.0),
            ("GEOMEAN", 24.0_f64.powf(0.25)),
            ("HARMEAN", 1.92),
            ("KURT", -1.2),
            ("SKEW", 0.0),
            ("SKEWP", 0.0),
        ] {
            let source = format!("={function}({NUMBERS})");
            assert_close(
                evaluate_value_source(&source, &resolver, &execution, mode)
                    .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
                expected,
                &source,
            );
        }
    }

    // NumberSequenceList functions admit ordered ReferenceList areas in
    // either evaluator mode.
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (function, expected) in [
            ("AVEDEV", 1.0),
            ("GEOMEAN", 24.0_f64.powf(0.25)),
            ("HARMEAN", 1.92),
            ("KURT", -1.2),
            ("SKEW", 0.0),
        ] {
            let source = format!("={function}([.D1:.D2]~[.E1:.E2])");
            assert_close(
                evaluate_value_source(&source, &resolver, &execution, mode)
                    .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
                expected,
                &source,
            );
        }

        // A 3-D cuboid is one Reference, so both NumberSequence and
        // NumberSequenceList functions admit its six ordered numeric cells.
        for (function, expected) in [
            ("AVEDEV", 1.5),
            ("DEVSQ", 17.5),
            ("GEOMEAN", 2.993795165523909),
            ("HARMEAN", 2.4489795918367343),
            ("KURT", -1.2),
            ("SKEW", 0.0),
            ("SKEWP", 0.0),
        ] {
            let source = format!("={function}([Main.A1:Archive.A2])");
            assert_close(
                evaluate_value_source(&source, &resolver, &execution, mode)
                    .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
                expected,
                &source,
            );
        }
    }
}

#[test]
fn number_sequence_reducers_refuse_reference_lists_before_reads() {
    let resolver = DescriptiveResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-list-refusal");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for function in ["DEVSQ", "SKEWP"] {
            resolver.clear_reads();
            let source = format!("={function}([.D1:.D2]~[.E1:.E2])");
            assert_error(
                evaluate_value_source(&source, &resolver, &execution, mode)
                    .unwrap_or_else(|error| panic!("{source}: {error}")),
                ScalarError::Value,
                &source,
            );
            assert_eq!(resolver.reads(), 0, "{source} must refuse before reads");
        }
    }
}

#[test]
fn referenced_formula_errors_are_not_omitted_by_descriptive_reducers() {
    let resolver = DescriptiveResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-reference-errors");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for function in [
            "AVEDEV", "DEVSQ", "GEOMEAN", "HARMEAN", "KURT", "SKEW", "SKEWP",
        ] {
            let source = format!("={function}([.C1:.C8])");
            assert_error(
                evaluate_value_source(&source, &resolver, &execution, mode).unwrap(),
                ScalarError::NotAvailable,
                &source,
            );
        }
    }
}

#[test]
fn empty_references_have_no_admitted_descriptive_values() {
    let resolver = DescriptiveResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-empty");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for function in [
            "AVEDEV", "DEVSQ", "GEOMEAN", "HARMEAN", "KURT", "SKEW", "SKEWP",
        ] {
            let source = format!("={function}([.H1:.H2])");
            assert_any_error(
                evaluate_value_source(&source, &resolver, &execution, mode).unwrap(),
                &source,
            );
        }
    }
}

#[test]
fn nested_descriptive_criterion_is_position_sensitive_and_complete_refs_cache() {
    let mut resolver = DescriptiveResolver::standard();
    // AVEDEV(MUNIT(1);0) = .5 and AVEDEV(MUNIT(2);0) = .48.  COUNTIF's
    // criterion context projects the MUNIT input by output position.
    for (row, value) in [99.0, 0.5, 0.48, 1.0, 2.0, 3.0].into_iter().enumerate() {
        resolver.set(row, 2, FixtureCell::Number(value));
    }
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-projection");
    let expression = parse("=IF({TRUE()|TRUE()};COUNTIF([.C1:.C6];AVEDEV(MUNIT([.G1:.G2]);0));0)");
    let result = evaluate_value(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected nested descriptive criterion");
    assert_array_numbers(&result, &[1.0, 1.0], "nested descriptive criterion");

    resolver.clear_reads();
    for (row, value) in [1.0, 2.0, 3.0, 4.0, 5.0, 6.0].into_iter().enumerate() {
        resolver.set(row, 2, FixtureCell::Number(value));
    }
    let expression = parse("=IF({TRUE()|TRUE()};AVEDEV([.C1:.C6]);0)");
    let result = evaluate_value(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected complete descriptive reference");
    assert_array_numbers(&result, &[1.5, 1.5], "cached descriptive projection");
    assert_eq!(
        resolver.reads(),
        12,
        "AVEDEV complete reference should be scanned in both passes once"
    );
}

#[test]
fn owned_descriptive_results_release_borrowed_reference_state() {
    let resolver = DescriptiveResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-descriptive-owned");
    let expression = parse("=DEVSQ([.B1:.B4])");
    let result = evaluate_value(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("descriptive reference result");
    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("descriptive result should be ownable");
    assert!(matches!(owned.value(), OwnedValueView::Number(value) if value == 5.0));
}
