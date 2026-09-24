//! Selected numeric caches from the pinned LibreOffice paired-statistics fixtures.
//!
//! The retained source closure contains only literal cells.  This test never
//! evaluates an upstream formula cell or treats a cached intermediate as an
//! input; it reconstructs the literal cells and runs the library evaluator in
//! the value modes recorded by the receipt.

use std::{
    collections::{BTreeMap, BTreeSet, HashMap},
    num::{NonZeroU64, NonZeroUsize},
    sync::atomic::{AtomicUsize, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{EvaluationFailure, ScalarError},
    expression::Expression,
};

#[derive(Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
}

#[derive(Debug)]
struct SheetFixture {
    rows: usize,
    columns: usize,
    cells: HashMap<(usize, usize), FixtureCell>,
}

#[derive(Debug)]
struct PairedStatisticsResolver {
    sheets: BTreeMap<String, SheetFixture>,
    reads: AtomicUsize,
}

impl PairedStatisticsResolver {
    fn new(
        formula_sheet: String,
        formula_row: usize,
        formula_column: usize,
        cells: &[serde_json::Value],
    ) -> Self {
        let mut sheets = BTreeMap::new();
        sheets.insert(
            formula_sheet,
            SheetFixture {
                rows: formula_row.max(1),
                columns: formula_column.max(1),
                cells: HashMap::new(),
            },
        );
        for cell in cells {
            let sheet = cell["sheet"].as_str().expect("cell sheet").to_owned();
            let row = cell["row"].as_u64().expect("cell row") as usize;
            let column = cell["column"].as_u64().expect("cell column") as usize;
            let value = match cell["type"].as_str().expect("cell type") {
                "empty" => FixtureCell::Empty,
                "number" => FixtureCell::Number(
                    cell["value"]
                        .as_str()
                        .expect("serialized number")
                        .parse()
                        .expect("finite fixture number"),
                ),
                "logical" => {
                    FixtureCell::Logical(cell["value"].as_str().expect("logical value") == "true")
                },
                "text" => FixtureCell::Text(cell["value"].as_str().expect("text value").to_owned()),
                other => panic!("unsupported fixture cell type {other:?}"),
            };
            let entry = sheets.entry(sheet).or_insert_with(|| SheetFixture {
                rows: 0,
                columns: 0,
                cells: HashMap::new(),
            });
            entry.rows = entry.rows.max(row);
            entry.columns = entry.columns.max(column);
            entry.cells.insert((row - 1, column - 1), value);
        }
        Self {
            sheets,
            reads: AtomicUsize::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Acquire)
    }
}

impl Resolver for PairedStatisticsResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok(self
            .sheets
            .get(sheet)
            .map(|fixture| SheetExtent::new(fixture.rows, fixture.columns)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.fetch_add(1, Ordering::AcqRel);
        let Some(fixture) = self.sheets.get(sheet) else {
            return Ok(CellRead::Error(ScalarError::Reference));
        };
        if row >= fixture.rows || column >= fixture.columns {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        Ok(match fixture.cells.get(&(row, column)) {
            None | Some(FixtureCell::Empty) => CellRead::Empty,
            Some(FixtureCell::Number(value)) => CellRead::Number(*value),
            Some(FixtureCell::Logical(value)) => CellRead::Logical(*value),
            Some(FixtureCell::Text(value)) => CellRead::Text(value),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok(self.sheets.keys().position(|name| name == sheet))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok(self.sheets.keys().nth(index).map(String::as_str))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(self.sheets.len())
    }
}

fn assert_close(actual: f64, expected: f64, formula: &str, mode: &str) {
    if expected == 0.0 {
        assert_eq!(
            actual.to_bits(),
            expected.to_bits(),
            "{mode} {formula}: {actual:?}"
        );
        return;
    }
    let tolerance = expected.abs() * 1e-13;
    assert!(
        actual.is_finite() && (actual - expected).abs() <= tolerance,
        "{mode} {formula}: {actual} versus {expected} (tolerance {tolerance})"
    );
}

fn assert_steyx_host_deviation(
    actual: f64,
    expected: f64,
    formula: &str,
    mode: &str,
    proof: &serde_json::Value,
) {
    assert_eq!(
        proof["formula"].as_str(),
        Some(formula),
        "{mode} STEYX proof formula"
    );
    let reference_bits = u64::from_str_radix(
        proof["reference_bits"]
            .as_str()
            .expect("STEYX reference bits"),
        16,
    )
    .expect("STEYX reference bits are hexadecimal");
    let native_bits = u64::from_str_radix(
        proof["native_bits"].as_str().expect("STEYX native bits"),
        16,
    )
    .expect("STEYX native bits are hexadecimal");
    assert_eq!(
        actual.to_bits(),
        reference_bits,
        "{mode} {formula}: evaluator differs from independent STEYX reference"
    );
    assert_eq!(
        expected.to_bits(),
        native_bits,
        "{mode} {formula}: retained native cache changed"
    );
    assert_eq!(
        reference_bits - native_bits,
        proof["ulp_difference"]
            .as_u64()
            .expect("STEYX ULP difference"),
        "{mode} {formula}: host difference changed"
    );
    assert_eq!(
        reference_bits - native_bits,
        2005,
        "{mode} {formula}: expected declared host difference"
    );
}

#[test]
fn selected_libreoffice_paired_statistics_caches_match_declared_host_deviation() {
    let source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/cached-results.json"
    );
    let observations: serde_json::Value =
        serde_json::from_str(source).expect("cached observations");
    let rows = observations.as_array().expect("observation array");
    assert_eq!(rows.len(), 42);
    let proof_source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/steyx-deviation-proof.json"
    );
    let proof: serde_json::Value = serde_json::from_str(proof_source).expect("STEYX proof");
    assert_eq!(proof["function"], "STEYX");
    assert_eq!(proof["pair_count"], 99);
    assert_eq!(proof["decimal_precision"], 300);

    let budget = Budget::root(
        "paired-statistics-native-caches",
        CoreLimits::for_profile(Profile::Server),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("execution limits"),
    );

    let expected_functions = BTreeSet::from([
        "CORREL",
        "COVAR",
        "FORECAST",
        "INTERCEPT",
        "PEARSON",
        "RSQ",
        "SLOPE",
        "STEYX",
    ]);
    let mut functions = BTreeSet::new();
    let mut saw_empty = false;
    let mut saw_text = false;
    let mut ordinary_observations = 0;
    let mut host_deviation_observations = 0;
    for row in rows {
        let function = row["function"].as_str().expect("function");
        functions.insert(function);
        let formula = row["formula"].as_str().expect("formula");
        let sheet = row["sheet"].as_str().expect("formula sheet").to_owned();
        let formula_row = row["row"].as_u64().expect("formula row") as usize;
        let formula_column = row["column"].as_u64().expect("formula column") as usize;
        let host_deviation = function == "STEYX" && formula_row == 3 && formula_column == 1;
        if host_deviation {
            host_deviation_observations += 1;
        } else {
            ordinary_observations += 1;
        }
        let cells = row["cells"].as_array().expect("literal cell closure");
        saw_empty |= cells.iter().any(|cell| cell["type"] == "empty");
        saw_text |= cells.iter().any(|cell| cell["type"] == "text");
        assert!(
            cells.iter().all(|cell| cell["type"] != "formula"),
            "formula-cell input leaked into {formula}"
        );
        let expected: f64 = row["cached"]
            .as_str()
            .expect("cached decimal")
            .parse()
            .expect("finite native cache");
        let valid_modes = row["valid_modes"].as_array().expect("valid modes");
        for mode in valid_modes {
            let mode_name = mode.as_str().expect("mode name");
            let mode = match mode_name {
                "scalar" => Mode::Scalar,
                "matrix" => Mode::Matrix,
                other => panic!("unsupported value mode {other:?}"),
            };
            let resolver =
                PairedStatisticsResolver::new(sheet.clone(), formula_row, formula_column, cells);
            let expression = Expression::parse(formula).expect("upstream formula parses");
            let context = Context::new(
                &execution,
                Position::new(&sheet, formula_row - 1, formula_column - 1),
            )
            .with_mode(mode);
            let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
                .unwrap_or_else(|error| panic!("{mode_name} {formula} should evaluate: {error:?}"));
            let Value::Number(actual) = result.value() else {
                panic!("{mode_name} {formula} produced {:?}", result.value());
            };
            if host_deviation {
                assert_steyx_host_deviation(actual, expected, formula, mode_name, &proof);
            } else {
                assert_close(actual, expected, formula, mode_name);
            }
            if cells.is_empty() {
                assert_eq!(
                    resolver.reads(),
                    0,
                    "literal formula unexpectedly read cells: {formula}"
                );
            } else {
                assert!(
                    resolver.reads() > 0,
                    "reference formula did not read cells: {formula}"
                );
            }
        }
    }
    assert_eq!(functions, expected_functions);
    assert_eq!(ordinary_observations, 41);
    assert_eq!(host_deviation_observations, 1);
    assert!(
        saw_empty,
        "native cases must retain a literal empty source cell"
    );
    assert!(
        saw_text,
        "native cases must retain a literal text source cell"
    );
}

#[test]
fn paired_statistics_native_receipt_pins_sources_and_exclusions() {
    let source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/provenance.json"
    );
    let provenance: serde_json::Value = serde_json::from_str(source).expect("provenance receipt");
    assert_eq!(
        provenance["commit"].as_str(),
        Some("d804d6aff49054bad1719ec3c2d136b545bbc7e7")
    );
    assert_eq!(provenance["inputs"].as_object().expect("inputs").len(), 8);
    assert_eq!(provenance["selected_observations"], 42);
    let host_deviations = provenance["host_numeric_deviations"]
        .as_array()
        .expect("host numeric deviations");
    assert_eq!(host_deviations.len(), 1);
    assert_eq!(host_deviations[0]["function"], "STEYX");
    assert_eq!(host_deviations[0]["row"], 3);
    assert_eq!(host_deviations[0]["kind"], "pinned_host_numeric_deviation");
    assert_eq!(host_deviations[0]["ulp_difference"], 2005);
    assert_eq!(
        provenance["normative_source"]["archive_sha256"],
        "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4"
    );
    assert_eq!(
        provenance["normative_source"]["member_sha256"],
        "ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1"
    );
    assert_eq!(
        provenance["normative_source"]["sections"],
        serde_json::json!(["6.18.38 INTERCEPT", "6.18.41 LINEST"])
    );
    assert!(
        provenance["normative_source"]["linest_const_text"]
            .as_str()
            .expect("LINEST Const text")
            .contains("equal to 0")
    );
    assert_eq!(
        provenance["selected_modes"],
        serde_json::json!(["scalar", "matrix"])
    );
    assert!(
        provenance["coverage_notes"]["source_format"]
            .as_str()
            .expect("source format note")
            .contains("FODS")
    );
    let exclusions = provenance["excluded_profile_variances"]
        .as_array()
        .expect("profile exclusions");
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "CORREL" && entry["row"] == 4 && entry["kind"] == "formula_closure"
    }));
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "INTERCEPT"
            && entry["row"] == 6
            && entry["kind"] == "named_range_unmaterialized"
    }));
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "FORECAST"
            && entry["row"] == 4
            && entry["column"] == 10
            && entry["kind"] == "invalid_forecast"
    }));
    assert!(
        provenance["coverage_notes"]["intercept_normative_ambiguity"]
            .as_str()
            .expect("INTERCEPT note")
            .contains("LINEST")
    );
    assert!(
        provenance["source_staging"]
            .as_str()
            .expect("source staging receipt")
            .contains("temporary tree")
    );
    assert!(
        provenance["byte_comparison"]
            .as_str()
            .expect("byte comparison receipt")
            .contains("SHA-256")
    );

    let selected_source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/cached-results.json"
    );
    let selected: serde_json::Value =
        serde_json::from_str(selected_source).expect("selected cache");
    let proof_source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/steyx-deviation-proof.json"
    );
    let proof: serde_json::Value = serde_json::from_str(proof_source).expect("STEYX proof");
    assert_eq!(proof["reference_hex"], "0x1.9a26885914f04p+0");
    assert_eq!(proof["native_hex"], "0x1.9a2688591472fp+0");
    assert_eq!(proof["ulp_difference"], 2005);
    assert!(
        selected
            .as_array()
            .expect("selected rows")
            .iter()
            .any(|row| {
                row["function"] == "INTERCEPT"
                    && row["row"] == 3
                    && row["cached"] == "6.75000000513889"
                    && row["cells"]
                        .as_array()
                        .is_some_and(|cells| cells.iter().all(|cell| cell["type"] == "number"))
            })
    );
}
