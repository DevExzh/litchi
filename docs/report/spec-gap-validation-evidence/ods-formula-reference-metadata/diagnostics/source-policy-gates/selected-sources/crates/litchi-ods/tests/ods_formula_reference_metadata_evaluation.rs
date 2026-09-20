//! Semantic coverage for the ODF reference and worksheet-metadata functions.
//!
//! The fixture deliberately keeps reference geometry separate from cell
//! contents.  AREAS, COLUMN(S), ISREF, ROW(S), and SHEET(S) must be able to
//! inspect a retained descriptor without reading its cells.  The same
//! resolver is used for scalar, matrix, projected, and three-dimensional
//! cases so that a test cannot accidentally pass through an array-only path.

use std::{
    cell::Cell,
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Limits, Mode, Position, Resolver, SheetExtent, Value,
    },
    evaluation::{EvaluationFailure, ScalarError, ScalarValue, UnsupportedKind, evaluate_scalar},
    expression::Expression,
};
use litchi_ods::{Builder, Spreadsheet};

#[derive(Debug, Clone)]
enum CellValue {
    Empty,
    Number(f64),
    Error(ScalarError),
}

#[derive(Debug)]
struct MetadataResolver {
    sheets: Vec<String>,
    extent: SheetExtent,
    cells: Vec<CellValue>,
    reads: Cell<usize>,
    metadata_calls: Cell<usize>,
}

impl MetadataResolver {
    fn standard() -> Self {
        let extent = SheetExtent::new(16, 12);
        let mut resolver = Self {
            sheets: ["Main", "Data", "Hidden", "Archive"]
                .into_iter()
                .map(str::to_owned)
                .collect(),
            extent,
            cells: vec![CellValue::Empty; extent.rows() * extent.columns()],
            reads: Cell::new(0),
            metadata_calls: Cell::new(0),
        };
        resolver.set(0, 0, CellValue::Number(10.0));
        resolver.set(1, 0, CellValue::Number(20.0));
        resolver.set(0, 1, CellValue::Number(30.0));
        resolver.set(1, 1, CellValue::Error(ScalarError::NotAvailable));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: CellValue) {
        let index = row * self.extent.columns() + column;
        self.cells[index] = value;
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn clear_reads(&self) {
        self.reads.set(0);
    }

    fn metadata_calls(&self) -> usize {
        self.metadata_calls.get()
    }

    fn clear_metadata_calls(&self) {
        self.metadata_calls.set(0);
    }
}

impl Resolver for MetadataResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        Ok(self
            .sheets
            .iter()
            .any(|name| name == sheet)
            .then_some(self.extent))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        self.reads.set(self.reads.get().saturating_add(1));
        if !self.sheets.iter().any(|name| name == sheet) {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        if row >= self.extent.rows() || column >= self.extent.columns() {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        Ok(match &self.cells[row * self.extent.columns() + column] {
            CellValue::Empty => CellRead::Empty,
            CellValue::Number(value) => CellRead::Number(*value),
            CellValue::Error(error) => CellRead::Error(*error),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        Ok(self.sheets.iter().position(|name| name == sheet))
    }

    fn sheet_name_at<'a>(
        &'a self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&'a str>, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        Ok(self.sheets.get(index).map(String::as_str))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        Ok(self.sheets.len())
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

#[derive(Debug, PartialEq)]
enum ScalarObserved {
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Other,
}

fn scalar(source: &str, execution: &ExecutionContext) -> Result<ScalarObserved, EvaluationFailure> {
    let expression = parse(source);
    let result = evaluate_scalar(
        &expression,
        &litchi_ods::codec::formula::evaluation::EvaluationContext::new(execution),
        &litchi_ods::codec::formula::evaluation::EvaluationLimits::default(),
    )?;
    Ok(match result.value() {
        ScalarValue::Number(value) => ScalarObserved::Number(*value),
        ScalarValue::Logical(value) => ScalarObserved::Logical(*value),
        ScalarValue::Text(value) => ScalarObserved::Text(value.to_string()),
        ScalarValue::Error(error) => ScalarObserved::Error(*error),
        ScalarValue::Complex(_) => ScalarObserved::Other,
        _ => ScalarObserved::Other,
    })
}

#[derive(Debug, PartialEq)]
enum CellObserved {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Other,
}

#[derive(Debug, PartialEq)]
enum Observed {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Array {
        rows: usize,
        columns: usize,
        cells: Vec<CellObserved>,
    },
    Other,
}

fn observe_cell(value: Option<Value<'_>>) -> CellObserved {
    match value {
        Some(Value::Empty) => CellObserved::Empty,
        Some(Value::Number(value)) => CellObserved::Number(value),
        Some(Value::Logical(value)) => CellObserved::Logical(value),
        Some(Value::Text(value)) => CellObserved::Text(value.to_owned()),
        Some(Value::Error(error)) => CellObserved::Error(error),
        Some(_) | None => CellObserved::Other,
    }
}

fn observe(result: &value::Evaluated<'_>) -> Observed {
    match result.value() {
        Value::Empty => Observed::Empty,
        Value::Number(value) => Observed::Number(value),
        Value::Logical(value) => Observed::Logical(value),
        Value::Text(value) => Observed::Text(value.to_owned()),
        Value::Error(error) => Observed::Error(error),
        Value::Array(array) => Observed::Array {
            rows: array.shape().rows(),
            columns: array.shape().columns(),
            cells: (0..array.len())
                .map(|index| observe_cell(array.get(index)))
                .collect(),
        },
        Value::Complex(_) | Value::Reference(_) | Value::ReferenceList(_) => Observed::Other,
        _ => Observed::Other,
    }
}

fn value_at(
    source: &str,
    resolver: &MetadataResolver,
    execution: &ExecutionContext,
    position: Position<'_>,
    mode: Mode,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    let context = Context::new(execution, position).with_mode(mode);
    let expression = parse(source);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(observe(&result))
}

fn value(
    source: &str,
    resolver: &MetadataResolver,
    execution: &ExecutionContext,
    mode: Mode,
) -> Result<Observed, EvaluationFailure> {
    value_at(
        source,
        resolver,
        execution,
        Position::new("Main", 0, 0),
        mode,
        &Limits::default(),
    )
}

fn assert_value(
    source: &str,
    expected: Observed,
    resolver: &MetadataResolver,
    execution: &ExecutionContext,
    mode: Mode,
) {
    assert_eq!(
        value(source, resolver, execution, mode)
            .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
        expected,
        "{source} in {mode:?}"
    );
}

fn assert_value_at(
    source: &str,
    expected: Observed,
    resolver: &MetadataResolver,
    execution: &ExecutionContext,
    position: Position<'_>,
    mode: Mode,
) {
    assert_eq!(
        value_at(
            source,
            resolver,
            execution,
            position,
            mode,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
        expected,
        "{source} in {mode:?}"
    );
}

#[test]
fn scalar_context_keeps_context_free_metadata_subset_typed() {
    let (_budget, _cancellation, execution) = execution("reference-metadata-scalar");

    assert_eq!(
        scalar("=ISREF(42)", &execution).expect("ISREF scalar"),
        ScalarObserved::Logical(false)
    );
    assert_eq!(
        scalar("=ISREF(#N/A)", &execution).expect("ISREF formula error"),
        ScalarObserved::Logical(false)
    );
    for source in ["=AREAS(42)", "=COLUMNS(42)", "=ROWS(42)"] {
        assert_eq!(
            scalar(source, &execution).expect("wrong scalar pseudotype produces formula value"),
            ScalarObserved::Error(ScalarError::Value),
            "{source}"
        );
    }
    for source in ["=SHEET(42)", "=SHEET(TRUE())", "=SHEET(\"Main\")"] {
        assert!(
            matches!(
                scalar(source, &execution),
                Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference))
            ),
            "{source} requires a value-evaluator resolver"
        );
    }
    for source in ["=ROW()", "=COLUMN()", "=SHEET()", "=SHEETS()"] {
        assert!(
            matches!(
                scalar(source, &execution),
                Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference))
            ),
            "{source} must not invent worksheet context"
        );
    }
}

#[test]
fn descriptor_functions_inspect_geometry_without_cell_reads() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-descriptors");

    for (source, expected) in [
        ("=AREAS([.A1:.B2])", Observed::Number(1.0)),
        ("=COLUMNS([.A1:.C2])", Observed::Number(3.0)),
        ("=ROWS([.A1:.C2])", Observed::Number(2.0)),
        ("=COLUMNS([.D4:.B2])", Observed::Number(3.0)),
        ("=ROWS([.D4:.B2])", Observed::Number(3.0)),
        ("=COLUMNS([.$B:.$D])", Observed::Number(3.0)),
        ("=ROWS([.$2:.$4])", Observed::Number(3.0)),
        ("=SHEET([.A1])", Observed::Number(1.0)),
        ("=SHEET([Data.A1:Archive.A1])", Observed::Number(2.0)),
        ("=SHEET(\"Hidden\")", Observed::Number(3.0)),
        (
            "=SHEET(\"Missing\")",
            Observed::Error(ScalarError::Reference),
        ),
        ("=SHEET(1)", Observed::Error(ScalarError::Reference)),
        ("=SHEET(TRUE())", Observed::Error(ScalarError::Reference)),
        ("=SHEETS([Main.A1:Archive.A1])", Observed::Number(4.0)),
        ("=ISREF([.A1])", Observed::Logical(true)),
        ("=ISREF(\"text\")", Observed::Logical(false)),
        ("=ISREF(TRUE())", Observed::Logical(false)),
        (r#"=ISREF("")"#, Observed::Logical(false)),
        ("=COLUMNS({1|#N/A})", Observed::Number(1.0)),
        ("=ROWS({1|#N/A})", Observed::Number(2.0)),
    ] {
        resolver.clear_reads();
        assert_value(source, expected, &resolver, &execution, Mode::Scalar);
        assert_eq!(resolver.reads(), 0, "{source} is descriptor-only");
    }
    assert_value(
        "=ISREF({1|2})",
        Observed::Logical(false),
        &resolver,
        &execution,
        Mode::Matrix,
    );
}

#[test]
fn computed_metadata_arguments_keep_scalar_intersection_and_matrix_shape() {
    let mut resolver = MetadataResolver::standard();
    resolver
        .sheets
        .extend(["1", "2"].into_iter().map(str::to_owned));
    resolver.set(0, 0, CellValue::Number(1.0));
    resolver.set(1, 0, CellValue::Number(2.0));
    let (_budget, _cancellation, execution) = execution("reference-metadata-computed-shape");

    assert_value_at(
        "=SHEET(ABS([.A1:.A2]))",
        Observed::Number(6.0),
        &resolver,
        &execution,
        Position::new("Main", 1, 0),
        Mode::Scalar,
    );
    assert_value(
        "=SHEET(ABS([.A1:.A2]))",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(5.0), CellObserved::Number(6.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=IF(TRUE();SHEET(ABS([.A1:.A2]));0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(5.0), CellObserved::Number(6.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value_at(
        "=COLUMNS(ABS([.A1:.A2]))",
        Observed::Error(ScalarError::Value),
        &resolver,
        &execution,
        Position::new("Main", 1, 0),
        Mode::Scalar,
    );
    assert_value(
        "=COLUMNS(ABS([.A1:.A2]))",
        Observed::Number(1.0),
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=COLUMNS({1;2})",
        Observed::Number(2.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
}

#[test]
fn catalog_dispatch_is_case_insensitive_and_rejects_nearby_names() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-catalog");
    for (source, expected) in [
        ("=aReAs([.A1])", Observed::Number(1.0)),
        ("=cOlUmN([.B1:.D1])", Observed::Number(2.0)),
        ("=cOlUmNs([.A1:.C2])", Observed::Number(3.0)),
        ("=iSrEf([.A1])", Observed::Logical(true)),
        ("=rOw([.A1:.A2])", Observed::Number(1.0)),
        ("=rOwS([.A1:.C2])", Observed::Number(2.0)),
        ("=sHeEt(\"Hidden\")", Observed::Number(3.0)),
        ("=sHeEtS()", Observed::Number(4.0)),
    ] {
        assert_value(source, expected, &resolver, &execution, Mode::Scalar);
    }
    assert!(matches!(
        scalar("=AREASX(1)", &execution),
        Err(EvaluationFailure::Unsupported(UnsupportedKind::Function))
    ));
}

#[test]
fn lists_and_three_dimensional_references_keep_area_and_sheet_order() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-3d");

    assert_value(
        "=AREAS(([.A1]~[.B1]))",
        Observed::Number(2.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=AREAS([Main.A1:Archive.C2])",
        Observed::Number(1.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=COLUMNS([Main.A1:Archive.C2])",
        Observed::Number(3.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=ROWS([Main.A1:Archive.C2])",
        Observed::Number(2.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=SHEETS([Main.A1:Archive.C2])",
        Observed::Number(4.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 0, "3-D/list metadata never reads cells");
}

#[test]
fn arbitrary_decomposition_single_record_lists_keep_their_list_pseudotype() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-single-record-list");
    let list = "(([.A1]~[.B2])![.A1])";

    for function in ["COLUMN", "COLUMNS", "ROW", "ROWS", "SHEET", "SHEETS"] {
        let source = format!("={function}({list})");
        resolver.clear_reads();
        resolver.clear_metadata_calls();
        assert_value(
            &source,
            Observed::Error(ScalarError::Value),
            &resolver,
            &execution,
            Mode::Scalar,
        );
        assert_eq!(resolver.reads(), 0, "{source} must refuse before reads");
    }

    resolver.clear_reads();
    resolver.clear_metadata_calls();
    assert_value(
        &format!("=AREAS({list})"),
        Observed::Number(1.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 0);

    assert_value(
        &format!("=ISREF({list})"),
        Observed::Logical(true),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn current_position_and_projected_matrix_metadata_are_stable() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-shape");

    assert_eq!(
        value_at(
            "=ROW()",
            &resolver,
            &execution,
            Position::new("Main", 4, 5),
            Mode::Scalar,
            &Limits::default(),
        )
        .expect("ROW uses the explicit value-evaluator position"),
        Observed::Number(5.0)
    );
    assert_eq!(
        value_at(
            "=COLUMN()",
            &resolver,
            &execution,
            Position::new("Main", 4, 5),
            Mode::Scalar,
            &Limits::default(),
        )
        .expect("COLUMN uses the explicit value-evaluator position"),
        Observed::Number(6.0)
    );
    assert_value(
        "=COLUMN([.B2:.D2])",
        Observed::Number(2.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=ROW([.B2:.D4])",
        Observed::Number(2.0),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(
        value_at(
            "=SHEET()",
            &resolver,
            &execution,
            Position::new("Main", 4, 5),
            Mode::Scalar,
            &Limits::default(),
        )
        .expect("SHEET has an explicit value-evaluator workbook context"),
        Observed::Number(1.0)
    );
    assert_eq!(
        value_at(
            "=SHEETS()",
            &resolver,
            &execution,
            Position::new("Main", 4, 5),
            Mode::Scalar,
            &Limits::default(),
        )
        .expect("SHEETS has an explicit value-evaluator workbook context"),
        Observed::Number(4.0)
    );
    for (source, expected) in [
        (
            "=IF({TRUE()|TRUE()};COLUMN();0)",
            vec![CellObserved::Number(6.0), CellObserved::Number(6.0)],
        ),
        (
            "=IF({TRUE()|TRUE()};ROW();0)",
            vec![CellObserved::Number(5.0), CellObserved::Number(5.0)],
        ),
    ] {
        assert_eq!(
            value_at(
                source,
                &resolver,
                &execution,
                Position::new("Main", 4, 5),
                Mode::Matrix,
                &Limits::default(),
            )
            .expect("position-sensitive metadata branch"),
            Observed::Array {
                rows: 2,
                columns: 1,
                cells: expected,
            },
            "{source}"
        );
    }
    assert_value(
        "=COLUMN([.A1:.C2])",
        Observed::Array {
            rows: 1,
            columns: 3,
            cells: vec![
                CellObserved::Number(1.0),
                CellObserved::Number(2.0),
                CellObserved::Number(3.0),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=ROW([.A1:.C2])",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=COLUMN([Main.A1:Archive.C2])",
        Observed::Array {
            rows: 1,
            columns: 3,
            cells: vec![
                CellObserved::Number(1.0),
                CellObserved::Number(2.0),
                CellObserved::Number(3.0),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=ROW([Main.A1:Archive.C2])",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=SHEET({\"Main\";\"Archive\"})",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(4.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );

    assert_value(
        "=IF({TRUE()|TRUE()};COLUMNS([.A1:.C2]);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(3.0), CellObserved::Number(3.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=IF({TRUE()|TRUE()};SHEETS([Main.A1:Archive.A1]);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(4.0), CellObserved::Number(4.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_value(
        "=IF({TRUE()|TRUE()};ISREF([.A1]);FALSE())",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Logical(true), CellObserved::Logical(true)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0, "projected metadata keeps descriptors");
}

#[test]
fn metadata_formula_errors_remain_values_and_do_not_trigger_reads() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-errors");

    assert_value(
        "=ISREF(#N/A)",
        Observed::Logical(false),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    for (source, error) in [
        ("=AREAS(#N/A)", ScalarError::NotAvailable),
        ("=COLUMN(#N/A)", ScalarError::NotAvailable),
        ("=ROW(#N/A)", ScalarError::NotAvailable),
        ("=SHEET(#N/A)", ScalarError::NotAvailable),
    ] {
        assert_value(
            source,
            Observed::Error(error),
            &resolver,
            &execution,
            Mode::Scalar,
        );
    }
    assert_value(
        "=COLUMNS(#N/A)",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=SHEETS(#REF!)",
        Observed::Error(ScalarError::Reference),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=ISREF(['file:///book.ods'#.A1])",
        Observed::Logical(true),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=SHEET(['file:///book.ods'#.A1])",
        Observed::Error(ScalarError::Value),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_value(
        "=SHEETS(['file:///book.ods'#.A1])",
        Observed::Error(ScalarError::Value),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn isref_preserves_computed_reference_identity_before_intersection() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-computed-kind");

    for source in [
        "=ISREF(IF(TRUE();[.A1];[.B1]))",
        "=ISREF(IF(TRUE();['file:///book.ods'#.A1];[.B1]))",
    ] {
        assert_value(
            source,
            Observed::Logical(true),
            &resolver,
            &execution,
            Mode::Scalar,
        );
    }
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn typed_external_descriptor_failures_are_not_caught_as_formula_values() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-external-typed");
    let mut unexpected = Vec::new();
    for source in [
        "=SHEET(IFERROR(['file:///book.ods'#.A1]+1;\"Main\"))",
        "=ISREF(['file:///book.ods'#.A1]+1)",
    ] {
        let result = value_at(
            source,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Scalar,
            &Limits::default(),
        );
        if !matches!(
            result,
            Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference))
        ) {
            unexpected.push(format!("{source}: {result:?}"));
        }
    }
    assert!(
        unexpected.is_empty(),
        "external value capability refusal must remain typed: {unexpected:?}"
    );
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn source_markers_survive_error_fallback_until_metadata_consumes_them() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-source-fallback");

    for source in [
        "=SHEET(IFERROR(['file:///book.ods'#.A1];\"Main\"))",
        "=SHEET(IFNA(['file:///book.ods'#.A1];\"Main\"))",
    ] {
        resolver.clear_reads();
        assert_value(
            source,
            Observed::Error(ScalarError::Value),
            &resolver,
            &execution,
            Mode::Scalar,
        );
        assert_eq!(resolver.reads(), 0, "{source} must remain descriptor-only");
    }

    for source in [
        "=ISREF(IFERROR(['file:///book.ods'#.A1];\"Main\"))",
        "=ISREF(IFNA(['file:///book.ods'#.A1];\"Main\"))",
    ] {
        assert_value(
            source,
            Observed::Logical(true),
            &resolver,
            &execution,
            Mode::Scalar,
        );
    }

    for source in [
        "=SHEET(IFERROR(['file:///book.ods'#.A1]+1;\"Main\"))",
        "=SHEET(IFNA(['file:///book.ods'#.A1]+1;\"Main\"))",
    ] {
        let result = value_at(
            source,
            &resolver,
            &execution,
            Position::new("Main", 0, 0),
            Mode::Scalar,
            &Limits::default(),
        );
        assert!(
            matches!(
                result,
                Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference))
            ),
            "{source}: {result:?}"
        );
    }
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn projected_lazy_scalar_condition_preserves_sheet_text_array_shape() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-sheet-array-shape");

    assert_value(
        "=IF(TRUE();SHEET({\"Main\";\"Data\"});0)",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn projected_sheet_distinguishes_computed_arrays_from_reference_descriptors() {
    let mut resolver = MetadataResolver::standard();
    resolver
        .sheets
        .extend(["1", "2"].into_iter().map(str::to_owned));
    resolver.set(0, 0, CellValue::Number(1.0));
    resolver.set(1, 0, CellValue::Number(2.0));
    let (_budget, _cancellation, execution) = execution("reference-metadata-sheet-projection");
    assert_value(
        "=IF({TRUE()};SHEET(ABS([.A1:.A2]));0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(5.0), CellObserved::Number(6.0)],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    for source in [
        "=IF({TRUE()};SHEET([.A1]:[.A2]);0)",
        "=IF({TRUE()};SHEET(IF(TRUE();[.A1:.A2];[.B1]));0)",
    ] {
        resolver.clear_reads();
        assert_value(
            source,
            Observed::Array {
                rows: 1,
                columns: 1,
                cells: vec![CellObserved::Number(1.0)],
            },
            &resolver,
            &execution,
            Mode::Matrix,
        );
        assert_eq!(resolver.reads(), 0, "{source}");
    }
}

#[test]
fn invalid_arity_and_reference_list_shape_refuse_before_reads() {
    let resolver = MetadataResolver::standard();
    let (_budget, _cancellation, execution) = execution("reference-metadata-arity");

    for source in [
        "=AREAS()",
        "=COLUMNS()",
        "=ROWS()",
        "=ISREF()",
        "=AREAS([.A1];[.B1])",
        "=COLUMNS([.A1];[.B1])",
        "=COLUMN(;) ",
        "=ROW(;) ",
        "=SHEET(;) ",
        "=SHEETS(;) ",
    ] {
        let source = source.trim_end();
        resolver.clear_reads();
        assert_value(
            source,
            Observed::Error(ScalarError::Value),
            &resolver,
            &execution,
            Mode::Scalar,
        );
        assert_eq!(resolver.reads(), 0, "{source} is rejected before reads");
        assert_eq!(
            resolver.metadata_calls(),
            0,
            "{source} is rejected before metadata calls"
        );
    }
    for source in [
        "=COLUMNS(([.A1]~[.B1]))",
        "=ROWS(([.A1]~[.B1]))",
        "=SHEET(([.A1]~[.B1]))",
        "=SHEETS(([.A1]~[.B1]))",
    ] {
        resolver.clear_reads();
        resolver.clear_metadata_calls();
        assert_value(
            source,
            Observed::Error(ScalarError::Value),
            &resolver,
            &execution,
            Mode::Scalar,
        );
        assert_eq!(resolver.reads(), 0, "{source} is rejected before reads");
    }
    assert_value(
        "=ISREF(([.A1]~[.B1]))",
        Observed::Logical(true),
        &resolver,
        &execution,
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_reads();
    resolver.clear_metadata_calls();
    assert_value(
        "=IF({TRUE()|TRUE()};COLUMN([.A1];[.B1]);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Error(ScalarError::Value),
            ],
        },
        &resolver,
        &execution,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);
    assert_eq!(resolver.metadata_calls(), 0);
}

#[test]
fn worksheet_resolver_counts_the_full_ordered_source_including_hidden_named_sheet() {
    use litchi_ods::worksheet::formula::Resolver as WorksheetResolver;
    const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.3"><office:automatic-styles><style:style style:name="hidden-sheet" style:family="table"><style:table-properties table:display="false"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Main"><table:table-row><table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Hidden" table:style-name="hidden-sheet"><table:table-row><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Archive"><table:table-row><table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
    let spreadsheet = Spreadsheet::from_bytes(
        Builder::new()
            .content_xml(CONTENT)
            .build()
            .expect("hidden worksheet ODS fixture should build"),
    )
    .expect("hidden worksheet ODS fixture should parse");
    assert!(
        spreadsheet
            .content_xml()
            .contains("table:display=\"false\"")
    );
    assert_eq!(spreadsheet.sheets()[1].name, "Hidden");
    assert_eq!(
        spreadsheet.sheets()[1].style_name.as_deref(),
        Some("hidden-sheet")
    );
    let sheets = spreadsheet.sheets();
    let (_budget, _cancellation, execution) = execution("reference-metadata-worksheet");
    let resolver = WorksheetResolver::new(sheets, SheetExtent::new(4, 4), &execution)
        .expect("worksheet resolver accepts the complete ordered source");
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Scalar);
    for (source, expected) in [
        ("=SHEETS()", 3.0),
        ("=SHEETS([Main.A1:Archive.A1])", 3.0),
        ("=SHEET(\"Hidden\")", 2.0),
    ] {
        let expression = parse(source);
        let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
            .expect("worksheet metadata evaluation");
        assert!(
            matches!(result.value(), Value::Number(value) if value == expected),
            "{source}"
        );
    }
}
