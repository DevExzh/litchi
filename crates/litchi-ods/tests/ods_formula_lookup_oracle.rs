//! Independent consumer for the ODS lookup/reference oracle.
//!
//! The Python model owns the expected values and the retained corpus is
//! intentionally loaded at runtime.  The consumer asserts the corpus's
//! declared implementation-contract identity before replaying its rows.  The
//! root evidence verifier independently hashes the contract and corpus bytes;
//! this test does not duplicate that cryptographic source binding.

use std::{
    cell::Cell,
    num::{NonZeroU64, NonZeroUsize},
    path::PathBuf,
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
struct LookupOracleResolver {
    sheets: [&'static str; 4],
    extent: SheetExtent,
    reads: Cell<usize>,
}

impl LookupOracleResolver {
    fn new() -> Self {
        Self {
            sheets: ["Main", "Data", "Hidden", "Archive"],
            extent: SheetExtent::new(16, 24),
            reads: Cell::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }
}

impl Resolver for LookupOracleResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok(self.sheets.contains(&sheet).then_some(self.extent))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.set(self.reads.get().saturating_add(1));
        // This is the finite workbook profile used by the independent Python
        // model: duplicate numeric tables, empty/error seams, and one explicit
        // Data-sheet value.  It has no fallback workbook or ambient access.
        Ok(match (sheet, row, column) {
            ("Main", 0, 0) => CellRead::Number(1.0),
            ("Main", 1, 0) => CellRead::Number(2.0),
            ("Main", 2, 0) => CellRead::Number(2.0),
            ("Main", 3, 0) => CellRead::Number(4.0),
            ("Main", 0, 1) => CellRead::Text("one"),
            ("Main", 1, 1) => CellRead::Text("two-first"),
            ("Main", 2, 1) => CellRead::Text("two-last"),
            ("Main", 3, 1) => CellRead::Text("four"),
            ("Main", 0, 2) => CellRead::Number(10.0),
            ("Main", 1, 2) => CellRead::Number(20.0),
            ("Main", 2, 2) => CellRead::Number(21.0),
            ("Main", 3, 2) => CellRead::Number(40.0),
            ("Main", 0, 4) => CellRead::Number(1.0),
            ("Main", 0, 5) | ("Main", 0, 6) => CellRead::Number(2.0),
            ("Main", 0, 7) => CellRead::Number(4.0),
            ("Main", 1, 4) => CellRead::Text("one"),
            ("Main", 1, 5) => CellRead::Text("two-first"),
            ("Main", 1, 6) => CellRead::Text("two-last"),
            ("Main", 1, 7) => CellRead::Text("four"),
            ("Main", 2, 4) => CellRead::Number(10.0),
            ("Main", 2, 5) => CellRead::Number(20.0),
            ("Main", 2, 6) => CellRead::Number(21.0),
            ("Main", 2, 7) => CellRead::Number(40.0),
            ("Main", 1, 9) => CellRead::Number(1.0),
            ("Main", 2, 9) => CellRead::Number(2.0),
            ("Main", 3, 9) => CellRead::Number(4.0),
            ("Main", 1, 10) => CellRead::Text("one"),
            ("Main", 2, 10) => CellRead::Text("two"),
            ("Main", 3, 10) => CellRead::Text("four"),
            ("Main", 0, 11) => CellRead::Error(ScalarError::NotAvailable),
            ("Main", 1, 11) => CellRead::Number(2.0),
            ("Main", 0, 12) => CellRead::Number(2.0),
            ("Main", 0, 13) => CellRead::Error(ScalarError::NotAvailable),
            ("Main", 1, 13) => CellRead::Number(20.0),
            ("Data", 2, 2) => CellRead::Number(303.0),
            _ => CellRead::Empty,
        })
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
        "ods-formula-lookup-independent-oracle",
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
        "#NULL!" => ScalarError::Null,
        "#DIV/0!" => ScalarError::DivisionByZero,
        "#VALUE!" => ScalarError::Value,
        "#REF!" => ScalarError::Reference,
        "#NAME?" => ScalarError::Name,
        "#NUM!" => ScalarError::Number,
        "#N/A" => ScalarError::NotAvailable,
        other => panic!("unsupported lookup-oracle error {other:?}"),
    }
}

fn assert_scalar(expected: &JsonValue, actual: Value<'_>, case: &str) {
    let kind = expected["type"].as_str().expect("expected scalar type");
    match kind {
        "number" => {
            let wanted = expected["value"].as_f64().expect("expected number");
            let Value::Number(observed) = actual else {
                panic!("{case}: expected Number({wanted}), got {actual:?}");
            };
            assert_eq!(
                observed.to_bits(),
                wanted.to_bits(),
                "{case}: number mismatch"
            );
        },
        "logical" => {
            let wanted = expected["value"].as_bool().expect("expected logical");
            let Value::Logical(observed) = actual else {
                panic!("{case}: expected Logical({wanted}), got {actual:?}");
            };
            assert_eq!(observed, wanted, "{case}: logical mismatch");
        },
        "text" => {
            let wanted = expected["value"].as_str().expect("expected text");
            let Value::Text(observed) = actual else {
                panic!("{case}: expected Text({wanted:?}), got {actual:?}");
            };
            assert_eq!(observed, wanted, "{case}: text mismatch");
        },
        "empty" => {
            assert!(
                matches!(actual, Value::Empty),
                "{case}: expected Empty, got {actual:?}"
            );
        },
        "error" => {
            let wanted = formula_error(expected["value"].as_str().expect("expected error"));
            let Value::Error(observed) = actual else {
                panic!("{case}: expected Error({wanted:?}), got {actual:?}");
            };
            assert_eq!(observed, wanted, "{case}: formula error mismatch");
        },
        other => panic!("{case}: scalar assertion received {other:?}"),
    }
}

fn assert_value(
    expected: &JsonValue,
    actual: Value<'_>,
    case: &str,
    resolver: &LookupOracleResolver,
) {
    match expected["type"].as_str().expect("expected value type") {
        "array" => {
            let Value::Array(array) = actual else {
                panic!("{case}: expected an array, got {actual:?}");
            };
            let rows = expected["rows"].as_u64().expect("array rows") as usize;
            let columns = expected["columns"].as_u64().expect("array columns") as usize;
            assert_eq!(array.shape().rows(), rows, "{case}: array rows");
            assert_eq!(array.shape().columns(), columns, "{case}: array columns");
            let values = expected["values"].as_array().expect("array values");
            assert_eq!(array.len(), values.len(), "{case}: array length");
            for (index, wanted) in values.iter().enumerate() {
                let observed = array
                    .get(index)
                    .unwrap_or_else(|| panic!("{case}: missing array cell {index}"));
                assert_scalar(wanted, observed, &format!("{case}[{index}]"));
            }
        },
        "reference" => {
            let Value::Reference(reference) = actual else {
                panic!("{case}: expected a reference, got {actual:?}");
            };
            let expected_areas = expected["areas"].as_array().expect("reference areas");
            assert_eq!(
                reference.areas().len(),
                expected_areas.len(),
                "{case}: area count"
            );
            match expected["owner"].as_str().unwrap_or("derived") {
                "direct" => assert!(
                    reference.reference().is_some(),
                    "{case}: direct descriptor lost its lexical owner"
                ),
                "derived" => assert!(
                    reference.reference().is_none(),
                    "{case}: derived descriptor retained a lexical owner"
                ),
                owner => panic!("{case}: unsupported descriptor owner {owner:?}"),
            }
            for (area, wanted) in reference.areas().iter().zip(expected_areas) {
                let sheet = wanted["sheet"].as_str().expect("reference sheet");
                let sheet_index = resolver
                    .sheets
                    .iter()
                    .position(|candidate| *candidate == sheet)
                    .unwrap_or_else(|| panic!("{case}: unknown expected sheet {sheet:?}"));
                let sheet_end = wanted["sheet_end"].as_str().unwrap_or(sheet);
                let sheet_end_index = resolver
                    .sheets
                    .iter()
                    .position(|candidate| *candidate == sheet_end)
                    .unwrap_or_else(|| panic!("{case}: unknown expected end sheet {sheet_end:?}"));
                let row = wanted["row"].as_u64().expect("reference row") as usize;
                let column = wanted["column"].as_u64().expect("reference column") as usize;
                let rows = wanted["rows"].as_u64().expect("reference rows") as usize;
                let columns = wanted["columns"].as_u64().expect("reference columns") as usize;
                assert_eq!(
                    area.starts(),
                    [sheet_index, row - 1, column - 1],
                    "{case}: starts"
                );
                assert_eq!(
                    area.extent(),
                    [sheet_end_index - sheet_index + 1, rows, columns],
                    "{case}: extent"
                );
            }
        },
        "reference_list" => {
            let Value::ReferenceList(list) = actual else {
                panic!("{case}: expected a reference list, got {actual:?}");
            };
            let expected_records = expected["records"].as_array().expect("reference records");
            assert_eq!(list.len(), expected_records.len(), "{case}: record count");
            for (record, wanted) in list.iter().zip(expected_records) {
                let areas = record.areas();
                assert_eq!(areas.len(), 1, "{case}: each fixture record has one area");
                let area = areas[0];
                match wanted["owner"].as_str().unwrap_or("derived") {
                    "direct" => assert!(
                        record.reference().is_some(),
                        "{case}: direct list descriptor lost its lexical owner"
                    ),
                    "derived" => assert!(
                        record.reference().is_none(),
                        "{case}: derived list descriptor retained a lexical owner"
                    ),
                    owner => panic!("{case}: unsupported descriptor owner {owner:?}"),
                }
                let sheet = wanted["sheet"].as_str().expect("reference sheet");
                let sheet_index = resolver
                    .sheets
                    .iter()
                    .position(|candidate| *candidate == sheet)
                    .unwrap_or_else(|| panic!("{case}: unknown expected sheet {sheet:?}"));
                let sheet_end = wanted["sheet_end"].as_str().unwrap_or(sheet);
                let sheet_end_index = resolver
                    .sheets
                    .iter()
                    .position(|candidate| *candidate == sheet_end)
                    .unwrap_or_else(|| panic!("{case}: unknown expected end sheet {sheet_end:?}"));
                let row = wanted["row"].as_u64().expect("reference row") as usize;
                let column = wanted["column"].as_u64().expect("reference column") as usize;
                let rows = wanted["rows"].as_u64().expect("reference rows") as usize;
                let columns = wanted["columns"].as_u64().expect("reference columns") as usize;
                assert_eq!(
                    area.starts(),
                    [sheet_index, row - 1, column - 1],
                    "{case}: starts"
                );
                assert_eq!(
                    area.extent(),
                    [sheet_end_index - sheet_index + 1, rows, columns],
                    "{case}: extent"
                );
            }
        },
        _ => assert_scalar(expected, actual, case),
    }
}

#[test]
fn lookup_oracle_matches_all_retained_rows() {
    let path: PathBuf = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(
        "../../docs/report/spec-gap-validation-evidence/ods-formula-lookups/lookup-goldens.json",
    );
    let bytes = std::fs::read(&path).unwrap_or_else(|error| {
        panic!(
            "retained lookup oracle is unavailable at {}: {error}",
            path.display()
        )
    });
    let document: JsonValue = serde_json::from_slice(&bytes).expect("lookup oracle JSON");
    let contract = document["contract_sha256"].as_str().expect("contract hash");
    assert_eq!(
        contract, "b112d66d687337912333f932c199f6e0ee5241fedc335aeaa95de572889e6aaf",
        "lookup corpus is bound to the reviewed implementation contract"
    );
    let rows = document["observations"].as_array().expect("observations");
    assert!(!rows.is_empty(), "lookup oracle observations are empty");
    let expected_functions: std::collections::BTreeSet<&str> = [
        "ADDRESS", "CHOOSE", "HLOOKUP", "INDEX", "INDIRECT", "LOOKUP", "MATCH", "OFFSET", "VLOOKUP",
    ]
    .into_iter()
    .collect();

    let execution = execution();
    let mut observed_functions = std::collections::BTreeSet::new();
    for row in rows {
        let case = row["case"].as_str().expect("oracle case");
        let function = row["function"].as_str().expect("oracle function");
        observed_functions.insert(function);
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
            row_number.checked_sub(1).expect("one-based row") as usize,
            column_number.checked_sub(1).expect("one-based column") as usize,
        );
        let expression = Expression::parse(formula)
            .unwrap_or_else(|error| panic!("{case}: {formula} should parse: {error}"));
        let resolver = LookupOracleResolver::new();
        let context = Context::new(&execution, position).with_mode(mode);
        let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
            .unwrap_or_else(|error| panic!("{case}: {formula} evaluation failed: {error:?}"));
        assert_value(&row["expected"], result.value(), case, &resolver);
        if let Some(expected_reads) = row["expected_reads"].as_u64() {
            assert_eq!(
                resolver.reads(),
                expected_reads as usize,
                "{case}: resolver reads"
            );
        }
        if let Some(minimum) = row["expected_reads_min"].as_u64() {
            assert!(
                resolver.reads() >= minimum as usize,
                "{case}: reads below bound"
            );
        }
        if let Some(maximum) = row["expected_reads_max"].as_u64() {
            assert!(
                resolver.reads() <= maximum as usize,
                "{case}: reads above bound"
            );
        }
    }
    assert_eq!(
        observed_functions, expected_functions,
        "oracle must cover all nine functions"
    );
}
