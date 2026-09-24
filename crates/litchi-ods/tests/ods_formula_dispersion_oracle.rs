//! Independent exact-rational observations for VAR/VARA/VARP/VARPA and the
//! four standard-deviation reducers.
//!
//! The retained JSON is generated without consulting the Rust evaluator.  It
//! includes typed resolver cells, ReferenceList admission, order-sensitive
//! cancellation probes, and high-precision Decimal square-root references.

use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationFailure, ScalarError,
        value::{
            CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value, evaluate,
        },
    },
    expression::Expression,
};
use serde_json::Value as JsonValue;
use std::num::{NonZeroU64, NonZeroUsize};

#[derive(Clone, Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Text(String),
    Logical(bool),
    Error(ScalarError),
}

struct Cells {
    cells: Vec<FixtureCell>,
    reads: AtomicUsize,
}

impl Cells {
    fn from_fixture(fixture: &JsonValue) -> Self {
        let encoding = fixture["encoding"].as_str().expect("fixture encoding");
        let mut cells = Vec::new();
        match encoding {
            "cells" => {
                cells.extend(parse_cells(
                    fixture["cells"].as_array().expect("fixture cells"),
                ));
            },
            "prefix_repeat_suffix" => {
                cells.extend(parse_cells(
                    fixture["prefix"].as_array().expect("fixture prefix"),
                ));
                let repeat = fixture["repeat"]
                    .as_u64()
                    .and_then(|value| usize::try_from(value).ok())
                    .expect("fixture repeat fits usize");
                let repeated = parse_cell(&fixture["repeated"]);
                cells.extend(std::iter::repeat_n(repeated, repeat));
                cells.extend(parse_cells(
                    fixture["suffix"].as_array().expect("fixture suffix"),
                ));
            },
            other => panic!("unknown fixture encoding {other}"),
        }
        Self {
            cells,
            reads: AtomicUsize::new(0),
        }
    }

    fn read_count(&self) -> usize {
        self.reads.load(Ordering::Acquire)
    }
}

impl Resolver for Cells {
    fn sheet_extent(
        &self,
        sheet: &str,
        _: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(SheetExtent::new(self.cells.len(), 1)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        assert_eq!(sheet, "Main");
        assert_eq!(column, 0);
        self.reads.fetch_add(1, Ordering::AcqRel);
        let cell = self.cells.get(row).expect("bounded fixture row");
        Ok(match cell {
            FixtureCell::Empty => CellRead::Empty,
            FixtureCell::Number(value) => CellRead::Number(*value),
            FixtureCell::Text(value) => CellRead::Text(value.as_str()),
            FixtureCell::Logical(value) => CellRead::Logical(*value),
            FixtureCell::Error(error) => CellRead::Error(*error),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, _: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(1)
    }
}

fn parse_cells(values: &[JsonValue]) -> Vec<FixtureCell> {
    values.iter().map(parse_cell).collect()
}

fn parse_cell(cell: &JsonValue) -> FixtureCell {
    match cell["kind"].as_str().expect("cell kind") {
        "Empty" => FixtureCell::Empty,
        "Number" => FixtureCell::Number(f64::from_bits(bits(&cell["bits"]))),
        "Text" => FixtureCell::Text(cell["value"].as_str().expect("text").to_owned()),
        "Logical" => FixtureCell::Logical(cell["value"].as_bool().expect("logical")),
        "Error" => FixtureCell::Error(formula_error(cell["value"].as_str().expect("error"))),
        other => panic!("unknown fixture cell kind {other}"),
    }
}

fn bits(value: &JsonValue) -> u64 {
    u64::from_str_radix(value.as_str().expect("binary64 hexadecimal bits"), 16)
        .expect("valid binary64 bits")
}

fn formula_error(name: &str) -> ScalarError {
    match name {
        "Value" => ScalarError::Value,
        "Number" => ScalarError::Number,
        "DivisionByZero" => ScalarError::DivisionByZero,
        "NotAvailable" => ScalarError::NotAvailable,
        other => panic!("unknown oracle error {other}"),
    }
}

fn execution(scope: &str) -> ExecutionContext {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (_cancellation, token) = CancellationSource::pair();
    ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("execution limits"),
    )
}

fn compare_result(row: &JsonValue, result: Value<'_>) {
    let case = row["fixture"].as_str().expect("fixture name");
    if let Some(error) = row["error"].as_str() {
        let expected = formula_error(error);
        assert!(
            matches!(result, Value::Error(actual) if actual == expected),
            "{case}: expected {expected:?}, got {result:?}"
        );
        return;
    }

    let Value::Number(actual) = result else {
        panic!("{case}: expected Number, got {result:?}");
    };
    assert!(actual.is_finite(), "{case}: non-finite result {actual:?}");
    let expected_bits = bits(&row["expected_bits"]);
    let expected = f64::from_bits(expected_bits);
    if expected == 0.0 {
        assert_eq!(
            actual.to_bits(),
            0,
            "{case}: expected canonical positive zero"
        );
        return;
    }
    assert!(
        !actual.is_sign_negative(),
        "{case}: negative dispersion {actual}"
    );

    match row["comparison"].as_str().expect("comparison policy") {
        "ulps" => {
            let maximum = row["max_ulps"].as_u64().expect("ULP bound");
            assert!(
                actual.to_bits().abs_diff(expected_bits) <= maximum,
                "{case}: {actual:e} versus {expected:e}, allowed {maximum} ULP"
            );
        },
        "relative" => {
            let maximum = row["max_relative_error"].as_f64().expect("relative bound");
            let relative = (actual - expected).abs() / expected.abs();
            assert!(
                relative <= maximum,
                "{case}: {actual:e} versus {expected:e}, relative error {relative:e} > {maximum:e}"
            );
        },
        other => panic!("{case}: unknown comparison policy {other}"),
    }
}

fn compare_function(function: &str) {
    let document: JsonValue = serde_json::from_str(include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-dispersion/numeric-goldens.json"
    ))
    .expect("retained independent dispersion oracle");
    let rows = document["observations"].as_array().expect("observations");
    assert_eq!(rows.len(), 512);
    assert_eq!(
        rows.iter()
            .filter(|row| row["function"] == function)
            .count(),
        64
    );
    let fixtures = document["fixtures"].as_object().expect("fixtures");

    for row in rows.iter().filter(|row| row["function"] == function) {
        let fixture_name = row["fixture"].as_str().expect("fixture");
        let fixture = fixtures.get(fixture_name).expect("fixture payload");
        let expression = Expression::parse(row["formula"].as_str().expect("formula"))
            .expect("dispersion reference parses");
        for mode in [Mode::Scalar, Mode::Matrix] {
            let cells = Cells::from_fixture(fixture);
            let execution = execution(&format!("dispersion-oracle-{function}-{fixture_name}"));
            let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(mode);
            let result = evaluate(&expression, &cells, &context, &Limits::default())
                .unwrap_or_else(|error| panic!("{fixture_name}: {error}"));
            compare_result(row, result.value());
            assert_eq!(
                cells.read_count(),
                row["expected_reads"].as_u64().expect("read count") as usize,
                "{fixture_name}: resolver read count in {mode:?} mode"
            );
        }
    }
}

macro_rules! oracle_test {
    ($test:ident, $name:literal) => {
        #[test]
        fn $test() {
            compare_function($name);
        }
    };
}

oracle_test!(var_matches_independent_oracle, "VAR");
oracle_test!(vara_matches_independent_oracle, "VARA");
oracle_test!(varp_matches_independent_oracle, "VARP");
oracle_test!(varpa_matches_independent_oracle, "VARPA");
oracle_test!(stdev_matches_independent_oracle, "STDEV");
oracle_test!(stdeva_matches_independent_oracle, "STDEVA");
oracle_test!(stdevp_matches_independent_oracle, "STDEVP");
oracle_test!(stdevpa_matches_independent_oracle, "STDEVPA");
