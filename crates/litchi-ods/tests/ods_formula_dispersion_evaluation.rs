//! Independent semantic coverage for the OpenFormula dispersion reducers:
//! `VAR`, `VARA`, `VARP`, `VARPA`, `STDEV`, `STDEVA`, `STDEVP`, and
//! `STDEVPA`.
//!
//! The fixture keeps reference cells typed so the NumberSequence and
//! NumberSequenceA admission rules remain observable.  It also exposes three
//! ordered sheets and a small projected matrix, which makes list, 3-D, and
//! nested criterion evaluation testable without relying on a host workbook.

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

const MIXED: &str = "[.A1:.A7]";

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct DispersionResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: Cell<usize>,
}

impl DispersionResolver {
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

impl Resolver for DispersionResolver {
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
        // The one-cell 3-D probe gives each sheet a distinct number while
        // preserving the heterogeneous Main-plane fixture for ordinary
        // references.
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
    resolver: &DispersionResolver,
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
    resolver: &'a DispersionResolver,
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
    let tolerance = expected.abs().max(actual.abs()).max(1.0) * 64.0 * f64::EPSILON;
    assert!(
        (actual - expected).abs() <= tolerance,
        "{source:?}: {actual} != {expected} (tolerance {tolerance})"
    );
}

fn assert_error(result: ResultValue, expected: ScalarError, source: &str) {
    assert_eq!(result, ResultValue::Error(expected), "{source:?}");
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
fn scalar_sample_population_and_a_variant_vectors() {
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-scalar");
    for (source, expected) in [
        ("=VAR(1;2;3;4)", 5.0 / 3.0),
        ("=VARP(1;2;3;4)", 5.0 / 4.0),
        ("=STDEV(1;2;3;4)", (5.0_f64 / 3.0).sqrt()),
        ("=STDEVP(1;2;3;4)", (5.0_f64 / 4.0).sqrt()),
        ("=VAR(\"3\";1;TRUE();FALSE())", 19.0 / 12.0),
        ("=VARP(\"3\";1;TRUE();FALSE())", 19.0 / 16.0),
        ("=VARA(1;3;TRUE();FALSE();\"text\";\"\")", 41.0 / 30.0),
        ("=VARPA(1;3;TRUE();FALSE();\"text\";\"\")", 41.0 / 36.0),
        (
            "=STDEVA(1;3;TRUE();FALSE();\"text\";\"\")",
            (41.0_f64 / 30.0).sqrt(),
        ),
        (
            "=STDEVPA(1;3;TRUE();FALSE();\"text\";\"\")",
            (41.0_f64 / 36.0).sqrt(),
        ),
    ] {
        let result = evaluate_scalar_source(source, &execution)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_close(result, expected, source);
    }
}

#[test]
fn scalar_dispersion_empty_and_singleton_denominators_are_typed() {
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-empty");
    for source in [
        "=VAR()",
        "=VARA()",
        "=VARP()",
        "=VARPA()",
        "=STDEV()",
        "=STDEVA()",
        "=STDEVP()",
        "=STDEVPA()",
    ] {
        assert_error(
            evaluate_scalar_source(source, &execution).unwrap(),
            ScalarError::Value,
            source,
        );
    }
    for source in ["=VAR(7)", "=VARA(7)", "=STDEV(7)", "=STDEVA(7)"] {
        assert_error(
            evaluate_scalar_source(source, &execution).unwrap(),
            ScalarError::Value,
            source,
        );
    }
    for source in ["=VARP(7)", "=VARPA(7)", "=STDEVP(7)", "=STDEVPA(7)"] {
        assert_close(
            evaluate_scalar_source(source, &execution).unwrap(),
            0.0,
            source,
        );
    }
}

#[test]
fn scalar_dispersion_errors_are_retained_and_counting_is_not_reused() {
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-errors");
    for function in [
        "VAR", "VARA", "VARP", "VARPA", "STDEV", "STDEVA", "STDEVP", "STDEVPA",
    ] {
        let source = format!("={function}(#N/A;1)");
        assert_error(
            evaluate_scalar_source(&source, &execution).unwrap(),
            ScalarError::NotAvailable,
            &source,
        );
    }
    assert_error(
        evaluate_scalar_source("=VAR(\"not-a-number\")", &execution).unwrap(),
        ScalarError::Value,
        "=VAR(\"not-a-number\")",
    );
    assert_close(
        evaluate_scalar_source("=VARPA(\"not-a-number\")", &execution).unwrap(),
        0.0,
        "=VARPA(\"not-a-number\")",
    );
    for source in [
        "=VARP(;)",
        "=VARPA(;)",
        "=STDEVP(;)",
        "=STDEVPA(;)",
        "=VARP(COMPLEX(1;2))",
    ] {
        assert_error(
            evaluate_scalar_source(source, &execution).unwrap(),
            ScalarError::Value,
            source,
        );
    }
}

#[test]
fn value_references_apply_number_sequence_and_a_admission() {
    let resolver = DispersionResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-reference");
    let expected = [
        ("VAR", 18.0),
        ("VARP", 9.0),
        ("STDEV", 18.0_f64.sqrt()),
        ("STDEVP", 3.0),
        ("VARA", 25.0 / 6.0),
        ("VARPA", 125.0 / 36.0),
        ("STDEVA", (25.0_f64 / 6.0).sqrt()),
        ("STDEVPA", (125.0_f64 / 36.0).sqrt()),
    ];
    for (function, expected) in expected {
        let source = format!("={function}({MIXED})");
        let result = evaluate_value_source(&source, &resolver, &execution, Mode::Matrix)
            .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert_close(result, expected, &source);
    }
    for function in [
        "VAR", "VARA", "VARP", "VARPA", "STDEV", "STDEVA", "STDEVP", "STDEVPA",
    ] {
        let source = format!("={function}([.A8])");
        assert_error(
            evaluate_value_source(&source, &resolver, &execution, Mode::Matrix).unwrap(),
            ScalarError::NotAvailable,
            &source,
        );
    }

    // A completely empty reference has no admitted member.  Empty cells are
    // omitted from both NumberSequence and Any; they are not coerced to zero.
    for function in [
        "VAR", "VARA", "VARP", "VARPA", "STDEV", "STDEVA", "STDEVP", "STDEVPA",
    ] {
        let source = format!("={function}([.A7])");
        assert_error(
            evaluate_value_source(&source, &resolver, &execution, Mode::Matrix).unwrap(),
            ScalarError::Value,
            &source,
        );
    }
}

#[test]
fn value_arrays_lists_three_dimensional_refs_and_modes_are_consistent() {
    let resolver = DispersionResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-shapes");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=VAR({1;2|TRUE();FALSE()})", 2.0 / 3.0),
            ("=VARP({1;2|TRUE();FALSE()})", 0.5),
            ("=VARA({1;2|TRUE();FALSE()})", 2.0 / 3.0),
            ("=VARPA({1;2|TRUE();FALSE()})", 0.5),
            ("=STDEV({1;2|TRUE();FALSE()})", (2.0_f64 / 3.0).sqrt()),
            ("=STDEVP({1;2|TRUE();FALSE()})", 0.5_f64.sqrt()),
        ] {
            let result = evaluate_value_source(source, &resolver, &execution, mode)
                .unwrap_or_else(|error| panic!("{source:?} mode {mode:?}: {error}"));
            assert_close(result, expected, source);
        }
    }

    for (function, expected) in [
        ("STDEV", 2.0),
        ("VARA", 4.0),
        ("VARPA", 8.0 / 3.0),
        ("STDEVA", 2.0),
        ("STDEVPA", (8.0_f64 / 3.0).sqrt()),
    ] {
        let list_source = format!("={function}([.B1:.B2]~[.B3])");
        assert_close(
            evaluate_value_source(&list_source, &resolver, &execution, Mode::Matrix).unwrap(),
            expected,
            &list_source,
        );
        let three_d = format!("={function}([Main.A1:Archive.A1])");
        assert_close(
            evaluate_value_source(&three_d, &resolver, &execution, Mode::Matrix).unwrap(),
            expected,
            &three_d,
        );
    }

    // NumberSequence rejects an explicit ReferenceList before any provider
    // cell is read.  STDEV uses NumberSequenceList and the A variants use
    // Any, so those functions are covered by the successful list vectors.
    for function in ["VAR", "VARP", "STDEVP"] {
        resolver.clear_reads();
        let source = format!("={function}([.B1:.B2]~[.B3])");
        assert_error(
            evaluate_value_source(&source, &resolver, &execution, Mode::Matrix).unwrap(),
            ScalarError::Value,
            &source,
        );
        assert_eq!(
            resolver.reads(),
            0,
            "{source} must refuse before resolver reads"
        );
    }
}

#[test]
fn direct_value_modes_match_scalar_dispersion_results() {
    let resolver = DispersionResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-direct");
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=VAR(\"3\";1;TRUE();FALSE())", 19.0 / 12.0),
            ("=VARP(\"3\";1;TRUE();FALSE())", 19.0 / 16.0),
            ("=VARA(1;3;TRUE();FALSE();\"text\";\"\")", 41.0 / 30.0),
            ("=VARPA(1;3;TRUE();FALSE();\"text\";\"\")", 41.0 / 36.0),
            ("=STDEV(\"3\";1;TRUE();FALSE())", (19.0_f64 / 12.0).sqrt()),
            ("=STDEVP(\"3\";1;TRUE();FALSE())", (19.0_f64 / 16.0).sqrt()),
        ] {
            let result = evaluate_value_source(source, &resolver, &execution, mode)
                .unwrap_or_else(|error| panic!("{source:?} mode {mode:?}: {error}"));
            assert_close(result, expected, source);
        }
    }
}

#[test]
fn nested_dispersion_criterion_recomputes_munit_but_caches_complete_refs() {
    let mut resolver = DispersionResolver::standard();
    // The projected MUNIT scalar is 1 at G1 and the 2x2 identity at G2.
    // VARP(MUNIT(...); 0) therefore yields .25 and .24 respectively, making a position-sensitive
    // criterion whose surrounding IF may still cache the complete A range.
    // Main.A1 is the resolver's 3-D probe and is always 2.0.  A2/A3 hold the
    // two projected criterion values; the remaining cells are distractors.
    for (row, value) in [99.0, 0.25, 0.24, 1.0, 0.5, 0.75].into_iter().enumerate() {
        resolver.set(row, 0, FixtureCell::Number(value));
    }
    resolver.set(0, 6, FixtureCell::Number(1.0));
    resolver.set(1, 6, FixtureCell::Number(2.0));
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-projection");
    let expression = parse("=IF({TRUE()|TRUE()};COUNTIF([.A1:.A6];VARP(MUNIT([.G1:.G2]);0));0)");
    let result = evaluate_value(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected nested dispersion criterion");
    assert_array_numbers(&result, &[1.0, 1.0], "nested dispersion criterion");

    resolver.clear_reads();
    let expression = parse("=IF({TRUE()|TRUE()};VARP([.A1:.A6]);0)");
    let result = evaluate_value(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("projected invariant dispersion");
    assert_array_numbers(
        &result,
        &[2.188 / 6.0, 2.188 / 6.0],
        "cached dispersion projection",
    );
    assert_eq!(
        resolver.reads(),
        6,
        "complete reference should be scanned once"
    );
}

#[test]
fn dispersion_state_handles_extreme_scale_without_spurious_infinity() {
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-extreme");
    let result = evaluate_scalar_source("=STDEVP(1e154;-1e154)", &execution).unwrap();
    assert_close(result, 1e154, "=STDEVP(1e154;-1e154)");
    let result = evaluate_scalar_source("=VARP(1e154;-1e154)", &execution).unwrap();
    assert_close(result, 1e308, "=VARP(1e154;-1e154)");
}

#[test]
fn owned_dispersion_results_release_borrowed_reference_state() {
    let resolver = DispersionResolver::standard();
    let (_budget, _cancellation, execution) = execution("ods-formula-dispersion-owned");
    let expression = parse("=VARP([.B1:.B2])");
    let result = evaluate_value(
        &expression,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("dispersion reference result");
    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("dispersion result should be ownable");
    assert!(matches!(owned.value(), OwnedValueView::Number(value) if value == 1.0));
}
