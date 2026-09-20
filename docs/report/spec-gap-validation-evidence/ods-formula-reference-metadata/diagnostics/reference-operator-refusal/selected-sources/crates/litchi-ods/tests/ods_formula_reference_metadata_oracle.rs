//! Independent oracle consumer for the reference-metadata evidence bundle.
//!
//! The expected rows are authored by the small Python geometry model in the
//! evidence directory.  This test supplies only workbook order and finite
//! extents to the value evaluator; its cell provider panics through a counted
//! empty result if metadata evaluation attempts a cell read.  The oracle thus
//! checks typed values, array shape, and the retained cell-read count.  The
//! metadata-only rows retain an expected zero; projected computed-reference
//! rows record the selected cell reads that precede their pseudotype refusal.

use std::{
    cell::Cell,
    collections::BTreeSet,
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationFailure, ScalarError,
        value::{self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value},
    },
    expression::Expression,
};
use serde_json::Value as JsonValue;

#[derive(Debug)]
struct MetadataOracleResolver {
    sheets: [&'static str; 4],
    extent: SheetExtent,
    reads: Cell<usize>,
}

impl MetadataOracleResolver {
    fn new() -> Self {
        Self {
            sheets: ["Main", "Data", "Archive", "Hidden"],
            extent: SheetExtent::new(16, 12),
            reads: Cell::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }
}

impl Resolver for MetadataOracleResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok(self.sheets.contains(&sheet).then_some(self.extent))
    }

    fn read_cell<'a>(
        &'a self,
        _sheet: &str,
        _row: usize,
        _column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.set(self.reads.get().saturating_add(1));
        Ok(CellRead::Empty)
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok(self.sheets.iter().position(|candidate| *candidate == sheet))
    }

    fn sheet_name_at<'a>(
        &'a self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&'a str>, EvaluationFailure> {
        Ok(self.sheets.get(index).copied())
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(self.sheets.len())
    }
}

fn execution() -> ExecutionContext {
    let budget = Budget::root(
        "reference-metadata-independent-oracle",
        CoreLimits::for_profile(Profile::Server),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1024 * 1024).expect("one in-flight MiB"),
        0,
    )
    .expect("valid execution limits");
    ExecutionContext::new(budget, token, limits)
}

fn formula_error(value: &str) -> ScalarError {
    match value {
        "#N/A" => ScalarError::NotAvailable,
        "#VALUE!" => ScalarError::Value,
        "#REF!" => ScalarError::Reference,
        "#NUM!" => ScalarError::Number,
        "#DIV/0!" => ScalarError::DivisionByZero,
        other => panic!("unsupported oracle formula error {other:?}"),
    }
}

fn assert_scalar(expected: &JsonValue, actual: Value<'_>, case: &str) {
    if let Some(expected) = expected.as_f64() {
        let Value::Number(actual) = actual else {
            panic!("{case}: expected Number({expected}), got {actual:?}");
        };
        assert_eq!(actual, expected, "{case}: number mismatch");
        return;
    }
    let kind = expected["type"].as_str().expect("expected scalar type");
    match kind {
        "number" => {
            let expected = expected["value"].as_f64().expect("expected number");
            let Value::Number(actual) = actual else {
                panic!("{case}: expected Number({expected}), got {actual:?}");
            };
            assert_eq!(actual, expected, "{case}: number mismatch");
        },
        "logical" => {
            let expected = expected["value"].as_bool().expect("expected logical");
            let Value::Logical(actual) = actual else {
                panic!("{case}: expected Logical({expected}), got {actual:?}");
            };
            assert_eq!(actual, expected, "{case}: logical mismatch");
        },
        "text" => {
            let expected = expected["value"].as_str().expect("expected text");
            let Value::Text(actual) = actual else {
                panic!("{case}: expected Text({expected:?}), got {actual:?}");
            };
            assert_eq!(actual, expected, "{case}: text mismatch");
        },
        "error" => {
            let expected = formula_error(expected["value"].as_str().expect("expected error"));
            let Value::Error(actual) = actual else {
                panic!("{case}: expected Error({expected:?}), got {actual:?}");
            };
            assert_eq!(actual, expected, "{case}: formula error mismatch");
        },
        other => panic!("{case}: unsupported scalar oracle type {other:?}"),
    }
}

fn assert_result(expected: &JsonValue, result: &value::Evaluated<'_>, case: &str) {
    if expected["type"].as_str() != Some("array") {
        assert_scalar(expected, result.value(), case);
        return;
    }

    let Value::Array(array) = result.value() else {
        panic!("{case}: expected an array, got {:?}", result.value());
    };
    let rows = expected["rows"].as_u64().expect("array rows") as usize;
    let columns = expected["columns"].as_u64().expect("array columns") as usize;
    assert_eq!(array.shape().rows(), rows, "{case}: array rows");
    assert_eq!(array.shape().columns(), columns, "{case}: array columns");
    let expected_values = expected["values"].as_array().expect("array values");
    assert_eq!(array.len(), expected_values.len(), "{case}: array length");
    for (index, expected) in expected_values.iter().enumerate() {
        let actual = array
            .get(index)
            .unwrap_or_else(|| panic!("{case}: missing array cell {index}"));
        assert_scalar(expected, actual, &format!("{case}[{index}]"));
    }
}

#[test]
fn reference_metadata_independent_oracle_matches_all_retained_rows() {
    let document: JsonValue = serde_json::from_str(include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/reference-metadata-goldens.json"
    ))
    .expect("retained reference-metadata oracle JSON");
    let rows = document["observations"]
        .as_array()
        .expect("oracle observations");
    assert!(!rows.is_empty(), "oracle observation set is empty");
    let expected_functions: BTreeSet<&str> = [
        "AREAS", "COLUMN", "COLUMNS", "ISREF", "ROW", "ROWS", "SHEET", "SHEETS",
    ]
    .into_iter()
    .collect();
    let observed_functions: BTreeSet<&str> = rows
        .iter()
        .map(|row| row["function"].as_str().expect("oracle function"))
        .collect();
    assert_eq!(observed_functions, expected_functions);

    let execution = execution();
    for row in rows {
        let case = row["case"].as_str().expect("oracle case");
        let formula = row["formula"].as_str().expect("oracle formula");
        let mode = match row["mode"].as_str().expect("oracle mode") {
            "scalar" => Mode::Scalar,
            "matrix" => Mode::Matrix,
            other => panic!("{case}: unsupported mode {other:?}"),
        };
        let position = &row["position"];
        let sheet = position["sheet"].as_str().expect("oracle sheet");
        let row_number = position["row"].as_u64().expect("oracle row");
        let column_number = position["column"].as_u64().expect("oracle column");
        let position = Position::new(
            sheet,
            row_number.checked_sub(1).expect("one-based oracle row") as usize,
            column_number
                .checked_sub(1)
                .expect("one-based oracle column") as usize,
        );
        let expression = Expression::parse(formula)
            .unwrap_or_else(|error| panic!("{case}: {formula} should parse: {error}"));
        let resolver = MetadataOracleResolver::new();
        let context = Context::new(&execution, position).with_mode(mode);
        let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
            .unwrap_or_else(|error| panic!("{case}: {formula} evaluation failed: {error:?}"));
        assert_result(&row["expected"], &result, case);
        let expected_reads = row["expected_reads"].as_u64().expect("oracle read count") as usize;
        assert_eq!(
            resolver.reads(),
            expected_reads,
            "{case}: resolver cell reads"
        );
    }
}
