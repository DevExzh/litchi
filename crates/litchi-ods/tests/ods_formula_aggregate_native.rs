//! Selected numeric caches from the upstream LibreOffice aggregate fixtures.
//!
//! The extractor stores only literal worksheet cells reachable from each
//! selected formula.  This resolver deliberately has no formula evaluator:
//! an upstream formula cell is rejected during extraction and cannot become a
//! hidden cached input to this test.

use std::{
    collections::HashMap,
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
struct AggregateResolver {
    sheet: String,
    rows: usize,
    columns: usize,
    cells: HashMap<(usize, usize), FixtureCell>,
    reads: AtomicUsize,
}

impl AggregateResolver {
    fn new(
        sheet: String,
        formula_row: usize,
        formula_column: usize,
        cells: &[serde_json::Value],
    ) -> Self {
        let mut map = HashMap::new();
        let mut rows = formula_row.max(1);
        let mut columns = formula_column.max(1);
        for cell in cells {
            let row = cell["row"].as_u64().expect("cell row") as usize;
            let column = cell["column"].as_u64().expect("cell column") as usize;
            rows = rows.max(row);
            columns = columns.max(column);
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
            map.insert((row - 1, column - 1), value);
        }
        Self {
            sheet,
            rows,
            columns,
            cells: map,
            reads: AtomicUsize::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Acquire)
    }
}

impl Resolver for AggregateResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((sheet == self.sheet).then_some(SheetExtent::new(self.rows, self.columns)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.fetch_add(1, Ordering::AcqRel);
        if sheet != self.sheet || row >= self.rows || column >= self.columns {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        Ok(match self.cells.get(&(row, column)) {
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
        Ok((sheet == self.sheet).then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok((index == 0).then_some(self.sheet.as_str()))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(1)
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
fn selected_libreoffice_aggregate_caches_match_literal_resolved_sources() {
    let source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-aggregates/native/cached-results.json"
    );
    let observations: serde_json::Value =
        serde_json::from_str(source).expect("cached observations");
    let rows = observations.as_array().expect("observation array");
    assert_eq!(rows.len(), 48);

    let budget = Budget::root(
        "aggregate-native-caches",
        CoreLimits::for_profile(Profile::Server),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("execution limits"),
    );

    let mut functions = std::collections::BTreeSet::new();
    let mut saw_empty = false;
    for row in rows {
        let function = row["function"].as_str().expect("function");
        functions.insert(function);
        let formula = row["formula"].as_str().expect("formula");
        let sheet = row["sheet"].as_str().expect("formula sheet").to_owned();
        let formula_row = row["row"].as_u64().expect("formula row") as usize;
        let formula_column = row["column"].as_u64().expect("formula column") as usize;
        let cells = row["cells"].as_array().expect("literal cell closure");
        saw_empty |= cells.iter().any(|cell| cell["type"] == "empty");
        let resolver = AggregateResolver::new(sheet.clone(), formula_row, formula_column, cells);
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
    assert_eq!(functions.len(), 7);
    assert!(
        saw_empty,
        "native cases must retain a literal empty source cell"
    );
}

#[test]
fn aggregate_native_receipt_keeps_formula_cells_out_of_the_resolver_contract() {
    let source = include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-aggregates/native/provenance.json"
    );
    let provenance: serde_json::Value = serde_json::from_str(source).expect("provenance receipt");
    assert_eq!(
        provenance["commit"].as_str(),
        Some("d804d6aff49054bad1719ec3c2d136b545bbc7e7")
    );
    let exclusions = provenance["excluded_profile_variances"]
        .as_array()
        .expect("profile exclusions");
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "PRODUCT"
            && entry["row"] == 12
            && entry["reason"]
                .as_str()
                .unwrap_or("")
                .contains("zero arguments")
    }));
    assert!(exclusions.iter().any(|entry| {
        entry["function"] == "SUMPRODUCT"
            && entry["row"] == 15
            && entry["reason"]
                .as_str()
                .unwrap_or("")
                .contains("formula cell")
    }));
    let text_variance = exclusions
        .iter()
        .find(|entry| entry["function"] == "SUMPRODUCT" && entry["row"] == 16)
        .expect("text conversion variance");
    assert_eq!(
        text_variance["source"],
        "sc/qa/unit/data/functions/array/fods/sumproduct.fods"
    );
    assert_eq!(text_variance["sheet"], "Sheet2");
    assert_eq!(text_variance["column"], 1);
    assert_eq!(
        text_variance["formula"],
        "of:=SUMPRODUCT([.J16:.J18];[.N16:.N18])"
    );
    assert_eq!(text_variance["cached"], "18");
    assert_eq!(text_variance["reference_cell"]["row"], 17);
    assert_eq!(text_variance["reference_cell"]["column"], 14);
    assert_eq!(text_variance["reference_cell"]["type"], "text");
    assert_eq!(text_variance["reference_cell"]["value"], "Unknown");
    assert!(
        text_variance["reason"]
            .as_str()
            .expect("text variance reason")
            .contains("converts malformed text")
    );
    assert_eq!(provenance["selected_observations"], 48);
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
            .contains("SHA-256 checked")
    );
}
