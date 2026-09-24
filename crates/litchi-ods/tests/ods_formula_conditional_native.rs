//! Selected numeric caches from the upstream LibreOffice conditional
//! aggregate fixtures.
//!
//! The extractor stores only literal worksheet cells reachable from each
//! selected formula.  This resolver has no formula evaluator of its own: an
//! upstream formula cell is rejected during extraction and cannot become a
//! hidden cached input to this test.

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
struct ConditionalResolver {
    sheets: BTreeMap<String, SheetFixture>,
    reads: AtomicUsize,
}

impl ConditionalResolver {
    fn new(
        formula_sheet: String,
        formula_row: usize,
        formula_column: usize,
        cells: &[serde_json::Value],
    ) -> Self {
        let mut sheets: BTreeMap<String, SheetFixture> = BTreeMap::new();
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

impl Resolver for ConditionalResolver {
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

fn assert_close(actual: f64, expected: f64, formula: &str) {
    if expected == 0.0 {
        assert_eq!(
            actual.to_bits(),
            expected.to_bits(),
            "{formula}: {actual:?}"
        );
        return;
    }
    let tolerance = expected.abs() * 1e-13;
    assert!(
        actual.is_finite() && (actual - expected).abs() <= tolerance,
        "{formula}: {actual} versus {expected} (tolerance {tolerance})"
    );
}

#[test]
fn selected_libreoffice_conditional_caches_match_literal_resolved_sources() {
    let source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/native/cached-results.json"
    );
    let observations: serde_json::Value =
        serde_json::from_str(source).expect("cached observations");
    let rows = observations.as_array().expect("observation array");
    assert_eq!(rows.len(), 64);
    assert!(rows.iter().any(|entry| {
        entry["function"] == "COUNTIFS" && entry["row"] == 38 && entry["cached"] == "1"
    }));

    let budget = Budget::root(
        "conditional-native-caches",
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
        "AVERAGEIF",
        "AVERAGEIFS",
        "COUNTIF",
        "COUNTIFS",
        "SUMIF",
        "SUMIFS",
    ]);
    let mut functions = BTreeSet::new();
    let mut saw_empty = false;
    let mut saw_reference = false;
    let mut saw_multi_sheet = false;
    for row in rows {
        let function = row["function"].as_str().expect("function");
        functions.insert(function);
        let formula = row["formula"].as_str().expect("formula");
        let sheet = row["sheet"].as_str().expect("formula sheet").to_owned();
        let formula_row = row["row"].as_u64().expect("formula row") as usize;
        let formula_column = row["column"].as_u64().expect("formula column") as usize;
        let cells = row["cells"].as_array().expect("literal cell closure");
        saw_empty |= cells.iter().any(|cell| cell["type"] == "empty");
        saw_reference |= !cells.is_empty();
        saw_multi_sheet |= cells
            .iter()
            .map(|cell| cell["sheet"].as_str().expect("cell sheet"))
            .any(|cell_sheet| cell_sheet != sheet);
        assert!(
            cells.iter().all(|cell| cell["type"] != "formula"),
            "formula-cell input leaked into {formula}"
        );
        let resolver = ConditionalResolver::new(sheet.clone(), formula_row, formula_column, cells);
        let expression = Expression::parse(formula).expect("upstream formula parses");
        let context = Context::new(
            &execution,
            Position::new(&sheet, formula_row - 1, formula_column - 1),
        )
        .with_mode(Mode::Matrix);
        let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
            .unwrap_or_else(|error| panic!("{formula} should evaluate: {error:?}"));
        let Value::Number(actual) = result.value() else {
            panic!("{formula} produced {:?}", result.value());
        };
        let expected: f64 = row["cached"]
            .as_str()
            .expect("cached decimal")
            .parse()
            .expect("finite number");
        assert_close(actual, expected, formula);
        assert!(
            resolver.reads() > 0,
            "reference formula did not read cells: {formula}"
        );
    }
    assert_eq!(functions, expected_functions);
    assert!(
        saw_empty,
        "native cases must retain a literal empty source cell"
    );
    assert!(
        saw_reference,
        "native cases must exercise a literal reference closure"
    );
    assert!(
        saw_multi_sheet,
        "SUMIFS must exercise its retained Data/Main closure"
    );
}

#[test]
fn conditional_native_receipt_keeps_profile_variances_and_converter_explicit() {
    let selected_source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/native/cached-results.json"
    );
    let selected: serde_json::Value =
        serde_json::from_str(selected_source).expect("selected cached observations");
    let rows = selected.as_array().expect("selected observation array");
    let source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/native/provenance.json"
    );
    let provenance: serde_json::Value = serde_json::from_str(source).expect("provenance receipt");
    assert_eq!(
        provenance["commit"].as_str(),
        Some("d804d6aff49054bad1719ec3c2d136b545bbc7e7")
    );
    assert_eq!(provenance["selected_observations"], 64);
    assert_eq!(
        provenance["inputs"]["sc/qa/unit/data/xls/opencl/math/sumifs.xls"],
        "813c493fd001683f3b7ae488689fdced7df053a5586fbb006fb76fc9295ff4b0"
    );
    assert_eq!(
        provenance["converter"]["version"],
        "LibreOffice 26.2.5.2 620(Build:2)"
    );
    assert!(
        provenance["sumifs_materialization"]
            .as_str()
            .expect("SUMIFS materialization receipt")
            .contains("temporary")
    );
    let exclusions = provenance["excluded_profile_variances"]
        .as_array()
        .expect("profile exclusions");
    for exclusion in exclusions {
        let source = exclusion["source"].as_str().expect("exclusion source");
        let sheet = exclusion["sheet"].as_str().expect("exclusion sheet");
        let excluded_rows = if let Some(row) = exclusion["row"].as_u64() {
            vec![row]
        } else {
            exclusion["rows"]
                .as_array()
                .expect("exclusion row or rows")
                .iter()
                .map(|row| row.as_u64().expect("exclusion row number"))
                .collect()
        };
        let function = exclusion["function"].as_str().expect("exclusion function");
        assert!(
            rows.iter().all(|accepted| {
                accepted["source"] != source
                    || accepted["sheet"] != sheet
                    || accepted["function"] != function
                    || accepted["row"]
                        .as_u64()
                        .is_none_or(|row| !excluded_rows.contains(&row))
            }),
            "excluded native observation leaked into selected corpus: {function} {sheet:?}"
        );
    }
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "SUMIF"
            && entry["row"] == 31
            && entry["kind"] == "reference_profile_variance"
    }));
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "AVERAGEIFS"
            && entry["row"] == 5
            && entry["kind"] == "host_regex_variance"
    }));
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "COUNTIFS"
            && entry["row"] == 8
            && entry["kind"] == "host_regex_variance"
    }));
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "SUMIF" && entry["kind"] == "empty_relational_criterion_variance"
    }));
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "COUNTIFS"
            && entry["row"] == 39
            && entry["kind"] == "stale_cached_result_variance"
            && entry["formula_cell_attributes"]["office_value"] == "1"
            && entry["formula_cell_attributes"]["display"] == "0"
            && entry["expected_literal"]["office_value"] == "0"
    }));
    assert!(
        provenance["source_staging"]
            .as_str()
            .expect("source staging receipt")
            .contains("temporary")
    );
    assert!(
        provenance["byte_comparison"]
            .as_str()
            .expect("byte comparison receipt")
            .contains("SHA-256")
    );
    assert!(
        provenance["coverage_notes"]["text_materialization"]
            .as_str()
            .expect("text materialization receipt")
            .contains("rejects nested")
    );
}
