//! Independent exact observations for the OpenFormula 1.4 §6.20 text family.
//!
//! The expected values come from `text_oracle.py`; this target only decodes
//! the retained typed corpus and supplies a deliberately small borrowed-text
//! resolver.  It exercises both scalar and matrix value contexts and checks
//! the resolver read receipts carried by the corpus.

use std::{
    collections::HashMap,
    num::{NonZeroU64, NonZeroUsize},
    sync::atomic::{AtomicUsize, Ordering},
};

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

#[derive(Clone, Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Text(String),
    Logical(bool),
    Error(ScalarError),
}

struct Cells {
    values: HashMap<(usize, usize), FixtureCell>,
    reads: AtomicUsize,
}

impl Cells {
    fn from_row(row: &JsonValue) -> Self {
        let mut values = HashMap::new();
        let cells = row["cells"].as_object().expect("fixture cells");
        for (address, value) in cells {
            values.insert(coordinate(address), parse_cell(value));
        }
        Self {
            values,
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
        Ok((sheet == "Main").then_some(SheetExtent::new(4, 4)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        assert_eq!(sheet, "Main");
        self.reads.fetch_add(1, Ordering::AcqRel);
        Ok(match self.values.get(&(row, column)) {
            None | Some(FixtureCell::Empty) => CellRead::Empty,
            Some(FixtureCell::Number(value)) => CellRead::Number(*value),
            Some(FixtureCell::Text(value)) => CellRead::Text(value.as_str()),
            Some(FixtureCell::Logical(value)) => CellRead::Logical(*value),
            Some(FixtureCell::Error(error)) => CellRead::Error(*error),
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

fn coordinate(address: &str) -> (usize, usize) {
    let mut letters = String::new();
    let mut digits = String::new();
    for character in address.chars() {
        if character.is_ascii_alphabetic() {
            letters.push(character.to_ascii_uppercase());
        } else {
            digits.push(character);
        }
    }
    assert!(
        !letters.is_empty() && !digits.is_empty(),
        "bad A1 address {address}"
    );
    let mut column = 0usize;
    for character in letters.bytes() {
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add(usize::from(character - b'A' + 1)))
            .expect("column fits")
            - 1;
    }
    let row = digits.parse::<usize>().expect("row digits") - 1;
    (row, column)
}

fn parse_cell(cell: &JsonValue) -> FixtureCell {
    match cell["kind"].as_str().expect("fixture kind") {
        "Empty" => FixtureCell::Empty,
        "Number" => FixtureCell::Number(f64::from_bits(bits(&cell["bits"]))),
        "Text" => FixtureCell::Text(cell["value"].as_str().expect("text value").to_owned()),
        "Logical" => FixtureCell::Logical(cell["value"].as_bool().expect("logical value")),
        "Error" => FixtureCell::Error(formula_error(cell["value"].as_str().expect("error"))),
        other => panic!("unsupported fixture cell kind {other}"),
    }
}

fn bits(value: &JsonValue) -> u64 {
    u64::from_str_radix(value.as_str().expect("binary64 bits"), 16).expect("valid bits")
}

fn formula_error(name: &str) -> ScalarError {
    match name {
        "NotAvailable" => ScalarError::NotAvailable,
        "Name" => ScalarError::Name,
        "Value" => ScalarError::Value,
        "DivisionByZero" => ScalarError::DivisionByZero,
        "Reference" => ScalarError::Reference,
        "Number" => ScalarError::Number,
        "Null" => ScalarError::Null,
        other => panic!("unsupported formula error {other}"),
    }
}

fn execution(scope: &str) -> ExecutionContext {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (_cancellation, token) = CancellationSource::pair();
    ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one worker"),
            NonZeroUsize::new(1).expect("one task"),
            NonZeroU64::new(1024 * 1024).expect("one MiB"),
            0,
        )
        .expect("execution limits"),
    )
}

fn compare_expected(row: &JsonValue, result: Value<'_>) {
    let case = format!("{} / {}", row["id"], row["origin"]);
    let expected = &row["expected"];
    match expected["kind"].as_str().expect("expected kind") {
        "Text" => {
            let Value::Text(actual) = result else {
                panic!("{case}: expected text, got {result:?}");
            };
            assert_eq!(
                actual,
                expected["value"].as_str().expect("expected text"),
                "{case}"
            );
        },
        "Number" => {
            let Value::Number(actual) = result else {
                panic!("{case}: expected number, got {result:?}");
            };
            assert_eq!(actual.to_bits(), bits(&expected["bits"]), "{case}");
        },
        "Logical" => {
            let Value::Logical(actual) = result else {
                panic!("{case}: expected logical, got {result:?}");
            };
            assert_eq!(
                actual,
                expected["value"].as_bool().expect("expected logical"),
                "{case}"
            );
        },
        "Error" => {
            let Value::Error(actual) = result else {
                panic!("{case}: expected formula error, got {result:?}");
            };
            assert_eq!(
                actual,
                formula_error(expected["value"].as_str().expect("expected error")),
                "{case}"
            );
        },
        other => panic!("{case}: unsupported expected kind {other}"),
    }
}

fn compare(row: &JsonValue, result: Value<'_>, mode: Mode) {
    // A reference is a non-scalar matrix operand.  Matrix lifting therefore
    // returns its rectangular result shape (a one-cell array in this corpus),
    // while scalar demand projects that same reference to one value.
    if mode == Mode::Matrix && row["expected_reads"].as_u64() == Some(1) {
        let Value::Array(array) = result else {
            panic!(
                "{} / Matrix: expected one-cell array, got {result:?}",
                row["id"]
            );
        };
        assert_eq!(array.shape().rows(), 1, "{} / Matrix rows", row["id"]);
        assert_eq!(array.shape().columns(), 1, "{} / Matrix columns", row["id"]);
        compare_expected(row, array.get(0).expect("one-cell matrix result"));
    } else {
        compare_expected(row, result);
    }
}

fn compare_function(function: &str) {
    let document: JsonValue = serde_json::from_str(include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-text-functions/text-goldens.json"
    ))
    .expect("retained text oracle");
    assert_eq!(document["schema"], "litchi-ods-text-oracle-v1");
    let rows = document["observations"].as_array().expect("observations");
    assert_eq!(rows.len(), 219);
    assert_eq!(
        rows.iter()
            .filter(|row| row["function"] == function)
            .count(),
        document["function_counts"][function]
    );

    for row in rows.iter().filter(|row| row["function"] == function) {
        let expression = Expression::parse(row["formula"].as_str().expect("formula"))
            .unwrap_or_else(|error| panic!("{}: formula parse failed: {error}", row["id"]));
        for mode in [Mode::Scalar, Mode::Matrix] {
            let cells = Cells::from_row(row);
            let execution = execution(&format!("text-oracle-{function}-{}", row["id"]));
            let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(mode);
            let result = evaluate(&expression, &cells, &context, &Limits::default())
                .unwrap_or_else(|error| panic!("{} / {mode:?}: {error}", row["id"]));
            compare(row, result.value(), mode);
            assert_eq!(
                cells.read_count(),
                row["expected_reads"].as_u64().expect("read receipt") as usize,
                "{} / {mode:?}: resolver reads",
                row["id"]
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

oracle_test!(asc_matches_independent_oracle, "ASC");
oracle_test!(char_matches_independent_oracle, "CHAR");
oracle_test!(clean_matches_independent_oracle, "CLEAN");
oracle_test!(code_matches_independent_oracle, "CODE");
oracle_test!(concatenate_matches_independent_oracle, "CONCATENATE");
oracle_test!(dollar_matches_independent_oracle, "DOLLAR");
oracle_test!(exact_matches_independent_oracle, "EXACT");
oracle_test!(find_matches_independent_oracle, "FIND");
oracle_test!(fixed_matches_independent_oracle, "FIXED");
oracle_test!(jis_matches_independent_oracle, "JIS");
oracle_test!(left_matches_independent_oracle, "LEFT");
oracle_test!(len_matches_independent_oracle, "LEN");
oracle_test!(lower_matches_independent_oracle, "LOWER");
oracle_test!(mid_matches_independent_oracle, "MID");
oracle_test!(proper_matches_independent_oracle, "PROPER");
oracle_test!(replace_matches_independent_oracle, "REPLACE");
oracle_test!(rept_matches_independent_oracle, "REPT");
oracle_test!(right_matches_independent_oracle, "RIGHT");
oracle_test!(search_matches_independent_oracle, "SEARCH");
oracle_test!(substitute_matches_independent_oracle, "SUBSTITUTE");
oracle_test!(t_matches_independent_oracle, "T");
oracle_test!(text_matches_independent_oracle, "TEXT");
oracle_test!(trim_matches_independent_oracle, "TRIM");
oracle_test!(unichar_matches_independent_oracle, "UNICHAR");
oracle_test!(unicode_matches_independent_oracle, "UNICODE");
oracle_test!(upper_matches_independent_oracle, "UPPER");
