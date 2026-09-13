//! Integration coverage for the non-evaluating expression tree's worksheet
//! value path.
//!
//! The fixtures are deliberately small and immutable.  They exercise the
//! OpenFormula rules in Part 4 §§3.3, 4.7–4.9, 5.8–5.13, and 6.3.2–6.4.13:
//! empty cells remain distinct from scalar zero, references preserve their
//! rectangular order, matrix operators broadcast by shape, and scalar
//! references use the caller's position for implicit intersection.  The
//! resolver never performs I/O or formula recalculation.

use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Mutex,
    atomic::{AtomicUsize, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
    Resource, SourceVersion,
};
use litchi_ods::codec::formula::evaluation::{EvaluationFailure, ScalarError, UnsupportedKind};
use litchi_ods::codec::formula::{
    evaluation::value::{
        CellRead, Context, Evaluated, Limits, Mode, Position, Resolver, SheetExtent, Value,
        evaluate,
    },
    expression::Expression,
    reference::{Address, EndpointValue, Reference, SheetSelector},
};

#[derive(Debug)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
    Unsupported,
}

#[derive(Debug)]
struct CellEntry {
    sheet: String,
    row: usize,
    column: usize,
    value: FixtureCell,
}

#[derive(Debug)]
struct FixtureResolver {
    sheets: Vec<(String, SheetExtent)>,
    cells: Vec<CellEntry>,
    reads: AtomicUsize,
    read_order: Mutex<Vec<(String, usize, usize)>>,
    fail_after: Option<usize>,
    fail_missing_metadata: bool,
    missing_metadata_calls: AtomicUsize,
    cancel_after_read: Option<CancellationSource>,
    cancel_on_final_source_version: Option<CancellationSource>,
    source_versions: Option<(SourceVersion, SourceVersion)>,
    source_version_calls: AtomicUsize,
}

impl FixtureResolver {
    fn new() -> Self {
        Self {
            sheets: vec![
                ("Main".to_owned(), SheetExtent::new(8, 8)),
                ("Data".to_owned(), SheetExtent::new(8, 8)),
                ("Archive".to_owned(), SheetExtent::new(8, 8)),
            ],
            cells: Vec::new(),
            reads: AtomicUsize::new(0),
            read_order: Mutex::new(Vec::new()),
            fail_after: None,
            fail_missing_metadata: false,
            missing_metadata_calls: AtomicUsize::new(0),
            cancel_after_read: None,
            cancel_on_final_source_version: None,
            source_versions: None,
            source_version_calls: AtomicUsize::new(0),
        }
    }

    fn set(&mut self, sheet: &str, row: usize, column: usize, value: FixtureCell) {
        self.cells.push(CellEntry {
            sheet: sheet.to_owned(),
            row,
            column,
            value,
        });
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::Acquire)
    }

    fn fail_after(&mut self, reads: usize) {
        self.fail_after = Some(reads);
    }

    fn fail_missing_metadata(&mut self) {
        self.fail_missing_metadata = true;
    }

    fn missing_metadata_calls(&self) -> usize {
        self.missing_metadata_calls.load(Ordering::Acquire)
    }

    fn cancel_after_read(&mut self, cancellation: &CancellationSource) {
        self.cancel_after_read = Some(cancellation.clone());
    }

    fn read_order(&self) -> Vec<(String, usize, usize)> {
        self.read_order
            .lock()
            .expect("fixture read-order lock")
            .clone()
    }

    fn source_changes(&mut self, expected: SourceVersion, observed: SourceVersion) {
        self.source_versions = Some((expected, observed));
        self.source_version_calls.store(0, Ordering::Release);
    }

    fn cancel_on_final_source_version(&mut self, cancellation: &CancellationSource) {
        self.cancel_on_final_source_version = Some(cancellation.clone());
    }
}

impl Resolver for FixtureResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        if sheet == "Missing" {
            self.missing_metadata_calls.fetch_add(1, Ordering::AcqRel);
            if self.fail_missing_metadata {
                return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
            }
        }
        Ok(self
            .sheets
            .iter()
            .find(|(name, _)| name == sheet)
            .map(|(_, extent)| *extent))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        let read_number = self.reads.fetch_add(1, Ordering::AcqRel);
        self.read_order
            .lock()
            .expect("fixture read-order lock")
            .push((sheet.to_owned(), row, column));

        if self.fail_after.is_some_and(|limit| read_number >= limit) {
            return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
        }

        let cancel_after_read = self.cancel_after_read.clone();
        let read = match self
            .cells
            .iter()
            .find(|entry| entry.sheet == sheet && entry.row == row && entry.column == column)
        {
            None => CellRead::Empty,
            Some(entry) => match &entry.value {
                FixtureCell::Empty => CellRead::Empty,
                FixtureCell::Number(value) => CellRead::Number(*value),
                FixtureCell::Logical(value) => CellRead::Logical(*value),
                FixtureCell::Text(value) => CellRead::Text(value.as_str()),
                FixtureCell::Error(error) => CellRead::Error(*error),
                FixtureCell::Unsupported => CellRead::Unsupported,
            },
        };
        if let Some(cancellation) = cancel_after_read {
            cancellation.cancel();
        }
        Ok(read)
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        if sheet == "Missing" {
            self.missing_metadata_calls.fetch_add(1, Ordering::AcqRel);
            if self.fail_missing_metadata {
                return Err(EvaluationFailure::Unsupported(UnsupportedKind::Reference));
            }
        }
        Ok(self.sheets.iter().position(|(name, _)| name == sheet))
    }

    fn sheet_name_at<'a>(
        &'a self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&'a str>, EvaluationFailure> {
        Ok(self.sheets.get(index).map(|(name, _)| name.as_str()))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(self.sheets.len())
    }

    fn source_version(
        &self,
        _execution: &ExecutionContext,
    ) -> Result<Option<SourceVersion>, EvaluationFailure> {
        let Some((expected, observed)) = self.source_versions else {
            return Ok(None);
        };
        let call = self.source_version_calls.fetch_add(1, Ordering::AcqRel);
        if call > 0 {
            if let Some(cancellation) = &self.cancel_on_final_source_version {
                cancellation.cancel();
            }
        }
        Ok(Some(if call == 0 { expected } else { observed }))
    }
}

fn make_execution(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), CoreLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1_024).expect("one in-flight KiB"),
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

fn evaluate_at<'expr>(
    expression: &'expr Expression,
    resolver: &'expr FixtureResolver,
    execution: &ExecutionContext,
    row: usize,
    column: usize,
    mode: Mode,
    limits: &Limits,
) -> Result<Evaluated<'expr>, EvaluationFailure> {
    let position = Position::new("Main", row, column);
    let context = Context::new(execution, position).with_mode(mode);
    evaluate(expression, resolver, &context, limits)
}

fn assert_number(value: Value<'_>, expected: f64) {
    match value {
        Value::Number(actual) => assert_eq!(actual, expected),
        other => panic!("expected Number({expected}), got {other:?}"),
    }
}

fn assert_error(value: Value<'_>, expected: ScalarError) {
    match value {
        Value::Error(actual) => assert_eq!(actual, expected),
        other => panic!("expected formula error {expected}, got {other:?}"),
    }
}

fn assert_array_numbers(
    array: litchi_ods::codec::formula::evaluation::value::ArrayView<'_>,
    expected: &[f64],
) {
    assert_eq!(array.len(), expected.len());
    for (index, expected) in expected.iter().copied().enumerate() {
        assert_number(
            array
                .get(index)
                .unwrap_or_else(|| panic!("array cell {index} is missing")),
            expected,
        );
    }
}

fn column_label(mut index: usize) -> String {
    let mut label = Vec::new();
    loop {
        label.push(char::from(b'A' + u8::try_from(index % 26).unwrap()));
        index /= 26;
        if index == 0 {
            break;
        }
        index -= 1;
    }
    label.into_iter().rev().collect()
}

fn repeated_union_source(pattern: &[&str], count: usize) -> String {
    assert!(!pattern.is_empty());
    let mut source = String::from("=(");
    for index in 0..count {
        if index != 0 {
            source.push('~');
        }
        source.push_str(pattern[index % pattern.len()]);
    }
    source.push(')');
    source
}

fn assert_large_duplicate_union(pattern: &[(&str, [usize; 3], [usize; 3])], scope: &str) {
    const COUNT: usize = 4_096;
    const LINEAR_WORK_PER_ENTRY: u64 = 256;
    const LINEAR_WORK_FIXED_COST: u64 = 16_384;

    let reference_sources: Vec<_> = pattern.iter().map(|(source, _, _)| *source).collect();
    let source = repeated_union_source(&reference_sources, COUNT);
    let expected: Vec<_> = pattern
        .iter()
        .map(|(source, _, _)| Reference::parse(source).expect("valid expected reference"))
        .collect();
    let resolver = FixtureResolver::new();
    let (budget, _cancellation, execution) = make_execution(scope);
    let expression = parse(&source);
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("4096-entry union should evaluate: {error}"));
    let list = result
        .as_reference_list()
        .expect("a union chain should remain a reference list");
    assert_eq!(list.len(), COUNT);
    for index in 0..COUNT {
        let (source, starts, ends) = pattern[index % pattern.len()];
        let reference = list
            .get(index)
            .unwrap_or_else(|| panic!("missing duplicate union record {index}"));
        assert_eq!(
            reference.reference(),
            Some(&expected[index % expected.len()])
        );
        assert_eq!(reference.areas().len(), 1, "{source} at index {index}");
        assert_eq!(
            reference.areas()[0].starts(),
            starts,
            "{source} at index {index}"
        );
        assert_eq!(
            reference.areas()[0].ends(),
            ends,
            "{source} at index {index}"
        );
    }
    assert_eq!(
        resolver.reads(),
        0,
        "reference union construction must not read provider cells"
    );

    let work = budget.used(Resource::Work);
    let linear_bound = (COUNT as u64)
        .saturating_mul(LINEAR_WORK_PER_ENTRY)
        .saturating_add(LINEAR_WORK_FIXED_COST);
    assert!(
        work <= linear_bound,
        "4096-entry union used {work} work units, above linear bound {linear_bound}"
    );
}

#[test]
fn left_associated_duplicate_unions_keep_order_and_use_linear_work() {
    // Repeated coordinates make duplicate retention observable while the
    // interspersed B/C records make an accidental sort or deduplication
    // visible in the public list order.
    let local_pattern = [
        ("[.A1]", [0, 0, 0], [1, 1, 1]),
        ("[.B1]", [0, 0, 1], [1, 1, 2]),
        ("[.A1]", [0, 0, 0], [1, 1, 1]),
        ("[.C1]", [0, 0, 2], [1, 1, 3]),
    ];
    assert_large_duplicate_union(&local_pattern, "ods-formula-value-large-local-union");

    // The same left-associated chain is exercised across the ordered Main,
    // Data, Archive sheets. Each reference is one retained 3-D cuboid and
    // duplicates must remain separate records in source order.
    let three_dimensional_pattern = [
        ("[Main.A1:Archive.A1]", [0, 0, 0], [3, 1, 1]),
        ("[Main.B1:Archive.B1]", [0, 0, 1], [3, 1, 2]),
        ("[Main.A1:Archive.A1]", [0, 0, 0], [3, 1, 1]),
        ("[Main.C1:Archive.C1]", [0, 0, 2], [3, 1, 3]),
    ];
    assert_large_duplicate_union(
        &three_dimensional_pattern,
        "ods-formula-value-large-3d-union",
    );
}

#[test]
fn large_union_geometry_planning_stays_within_default_work_budget() {
    const COUNT: usize = 4_096;
    let union = repeated_union_source(&["[.A1]"], COUNT);
    let union_body = union
        .strip_prefix('=')
        .expect("repeated union source starts with an equals sign");
    let source = format!("=IF({{TRUE()}};({union_body}:[.A1]);0)");
    let expression = parse(&source);
    let resolver = FixtureResolver::new();
    let (budget, _cancellation, execution) = make_execution("ods-formula-value-large-union-shape");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("large union geometry should stay bounded: {error}"));
    let array = result
        .as_array()
        .expect("the collapsed one-cell range should materialize as an array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 1));
    assert!(matches!(array.get(0), Some(Value::Empty)));
    assert_eq!(resolver.reads(), 1);
    let work = budget.used(Resource::Work);
    assert!(
        work <= (COUNT as u64).saturating_mul(128).saturating_add(65_536),
        "large union shape planning used {work} work units"
    );
}

#[test]
fn reference_operand_kind_planning_keeps_intersections_linear() {
    const COUNT: usize = 4_096;
    let intersections = repeated_union_source(&["[.A1]"], COUNT).replace('~', "!");
    let body = intersections.strip_prefix('=').expect("formula prefix");
    for source in [
        format!("=IF({{TRUE()}};([.A1]:IF(AND({body});[.C1:.D1];0));0)"),
        format!("=IF({{TRUE()}};([.A1]:IFERROR(({body}![.B1]);[.C1:.D1]));0)"),
    ] {
        let mut resolver = FixtureResolver::new();
        for (column, value) in [10.0, 11.0, 12.0, 13.0].into_iter().enumerate() {
            resolver.set("Main", 0, column, FixtureCell::Number(value));
        }
        let (_budget, _cancellation, execution) =
            make_execution("ods-formula-value-intersection-kind-work");
        let expression = parse(&source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("linear intersection operand should evaluate: {error:?}"));
        let array = result.as_array().expect("selected range array");
        assert_eq!((array.shape().rows(), array.shape().columns()), (1, 4));
        assert_array_numbers(array, &[10.0, 11.0, 12.0, 13.0]);
    }
}

#[test]
fn union_cell_limit_includes_intersection_left_operand_before_provider_reads() {
    let resolver = FixtureResolver::new();
    let (budget, _cancellation, execution) =
        make_execution("ods-formula-value-union-intersection-cell-limit");
    let expression = parse("=(([.A1:.B2]![.B1:.C3])~[.D1:.E2])");

    // The intersection contributes B1:B2 (two logical cells) and the right
    // range contributes D1:E2 (four). A limit of five must therefore reject
    // the union while it is still a reference-only operation.
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default().with_max_reference_cells(5),
    )
    .expect_err("the union must include its intersection operand in the cell budget");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn sibling_shape_growth_reuses_broadcast_condition_cells() {
    const ROWS: usize = 32;
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Logical(true));
    // The outer FALSE branch prevents B1 from being selected. If the nested
    // XOR condition is re-evaluated for every widened row, this unreadable
    // sibling would also expose a wrong provider read.
    resolver.set("Main", 0, 1, FixtureCell::Unsupported);

    let mut false_branch = String::from("{");
    for row in 0..ROWS {
        if row != 0 {
            false_branch.push('|');
        }
        false_branch.push_str(&row.to_string());
    }
    false_branch.push('}');
    let source = format!("=IF({{TRUE();FALSE()}};IF(XOR([.A1:.B1]);{{1}};0);{false_branch})");
    let expression = parse(&source);
    let (budget, _cancellation, execution) = make_execution("ods-formula-value-sibling-growth");
    let limits = Limits::default().with_max_reference_cells(ROWS - 1);
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .unwrap_or_else(|error| panic!("broadcast condition should evaluate once: {error}"));
    let array = result
        .as_array()
        .expect("the widened sibling branch should produce an array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (ROWS, 2));
    for row in 0..ROWS {
        assert_number(array.get(row * 2).expect("selected branch cell"), 1.0);
        assert_number(
            array.get(row * 2 + 1).expect("outer false branch cell"),
            row as f64,
        );
    }
    assert_eq!(
        resolver.reads(),
        1,
        "the broadcast A1 condition is read once"
    );
    assert_eq!(
        resolver.read_order(),
        vec![("Main".to_owned(), 0, 0)],
        "the unreadable B1 sibling must remain unselected"
    );
    assert!(
        budget.used(Resource::Work) < 100_000,
        "condition-cache reuse should keep sibling growth bounded"
    );
}

#[test]
fn nested_range_conditions_contribute_their_matrix_shape() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Logical(true));
    resolver.set("Main", 0, 1, FixtureCell::Logical(false));
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-range-condition");
    let expression = parse("=IF({TRUE()};IF(([.A1]:[.B1]);1;0);0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a bounded range condition should drive the nested shape");
    let array = result
        .as_array()
        .expect("the range condition should produce a matrix");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[1.0, 0.0]);
    assert_eq!(
        resolver.read_order(),
        vec![("Main".to_owned(), 0, 0), ("Main".to_owned(), 0, 1),],
        "both selected range-condition cells should be read in order"
    );
}

#[test]
fn nested_range_branches_contribute_their_matrix_shape() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(10.0));
    resolver.set("Main", 0, 1, FixtureCell::Number(11.0));
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-range-branch");
    let expression = parse("=IF({TRUE()};([.A1]:[.B1]);0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a selected bounded range branch should retain its shape");
    let array = result
        .as_array()
        .expect("the selected range branch should produce a matrix");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[10.0, 11.0]);
    assert_eq!(
        resolver.read_order(),
        vec![("Main".to_owned(), 0, 0), ("Main".to_owned(), 0, 1),],
        "selected range cells should be materialized in row-major order"
    );
}

#[test]
fn nested_composed_range_condition_drives_selected_shape() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Logical(true));
    resolver.set("Main", 0, 1, FixtureCell::Logical(false));
    resolver.set("Main", 0, 2, FixtureCell::Logical(true));
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-composed-range-condition");
    let expression = parse("=IF({TRUE()};IF((([.A1]:[.B1]):[.C1]);1;0);[Missing.A1:.Z100])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a composed range condition should retain its full shape");
    let array = result
        .as_array()
        .expect("the composed condition should produce a matrix");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 3));
    assert_array_numbers(array, &[1.0, 0.0, 1.0]);
    assert_eq!(
        resolver.read_order(),
        vec![
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 0, 1),
            ("Main".to_owned(), 0, 2),
        ],
        "each cell of the selected composed condition should be read once"
    );
    assert_eq!(resolver.missing_metadata_calls(), 0);
}

#[test]
fn nested_range_branch_collapses_disjoint_union_to_bounding_shape() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(10.0));
    resolver.set("Main", 0, 1, FixtureCell::Number(11.0));
    resolver.set("Main", 0, 2, FixtureCell::Number(12.0));
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-union-range-branch");
    let expression = parse("=IF({TRUE()};([.A1]~[.C1]):[.B1];[Missing.A1:.Z100])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a range over a disjoint union should collapse its bounding shape");
    let array = result
        .as_array()
        .expect("the selected range branch should produce a matrix");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 3));
    assert_array_numbers(array, &[10.0, 11.0, 12.0]);
    assert_eq!(
        resolver.read_order(),
        vec![
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 0, 1),
            ("Main".to_owned(), 0, 2),
        ],
        "the bounding range should materialize cells in row-major order"
    );
    assert_eq!(resolver.missing_metadata_calls(), 0);
}

#[test]
fn scalar_condition_broadcast_over_reference_shape_does_not_consume_alias_stack() {
    const ROWS: usize = 64;
    let mut resolver = FixtureResolver::new();
    resolver.sheets[0].1 = SheetExtent::new(ROWS, 8);
    for row in 0..ROWS {
        resolver.set("Main", row, 0, FixtureCell::Number((row + 1) as f64));
    }
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-condition-alias-stack");
    let expression = parse("=IF({TRUE()};IF(TRUE();[.A1:.A64];0);0)");
    let limits = Limits::default()
        .with_max_array_cells(ROWS)
        .with_max_stack_entries(32);
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("a scalar condition should broadcast without per-output aliases");
    let array = result
        .as_array()
        .expect("the selected reference should produce a vertical array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (ROWS, 1));
    let expected: Vec<_> = (1..=ROWS).map(|value| value as f64).collect();
    assert_array_numbers(array, &expected);
    assert_eq!(
        resolver.reads(),
        ROWS,
        "each selected reference cell should be read once"
    );
    let expected_order: Vec<_> = (0..ROWS).map(|row| ("Main".to_owned(), row, 0)).collect();
    assert_eq!(resolver.read_order(), expected_order);
}

#[test]
fn if_without_a_branch_keeps_reference_operand_refusal_typed() {
    for source in ["=([.A1]:IF(TRUE()))", "=IF({TRUE()};([.A1]:IF(TRUE()));0)"] {
        let mut resolver = FixtureResolver::new();
        resolver.set("Main", 0, 0, FixtureCell::Number(10.0));
        let (_budget, _cancellation, execution) =
            make_execution("ods-formula-value-if-single-condition-reference");
        let expression = parse(source);
        let error = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &Limits::default(),
        )
        .expect_err("a scalar IF result is a typed reference-operator refusal");
        assert!(
            matches!(
                error,
                EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator)
            ),
            "{source:?} returned {error:?}"
        );
    }
}

#[test]
fn matrix_reference_operator_operands_keep_their_typed_refusal() {
    let mut failures = Vec::new();
    for operand in [
        "IF({TRUE();FALSE()};[.C1:.D1];[.E1:.F1])",
        "IFERROR({#DIV/0!;1};[.C1:.D1])",
        "IFNA({#N/A;1};[.C1:.D1])",
        "IF({TRUE()};[.C1:.D1];0)",
        "IFERROR([.C1:.D1];0)",
        "IFNA([.C1:.D1];0)",
        "IF(IF(TRUE();{TRUE()};FALSE());[.C1:.D1];0)",
        "IF(({TRUE()}+0);[.C1:.D1];0)",
        "IF([.C1];[.C1:.D1];[.C1:.D1])",
        "IFERROR(IF(TRUE();{#DIV/0!};1);[.C1:.D1])",
        "IFNA(IF(TRUE();{#N/A};1);[.C1:.D1])",
        "IFERROR((([.A1]![.B1]):[.C1]);[.C1:.D1])",
        "IFERROR((([.A1]~[.B1])+{1});[.C1:.D1])",
        "IFERROR(XOR(([.A1]~[.B1]);{TRUE()});[.C1:.D1])",
        "IFERROR(XOR({TRUE()};([.A1]~[.B1]));[.C1:.D1])",
        "IFERROR((([.A1]~[.B1]):[.C1]);[.C1:.D1])",
        "IFERROR(IFNA((#N/A+([.A1]~[.B1]));{1});[.C1:.D1])",
    ] {
        for source in [
            format!("=([.A1]:{operand})"),
            format!("=IF({{TRUE()}};([.A1]:{operand});[Missing.A1:.Z100])"),
        ] {
            let mut resolver = FixtureResolver::new();
            resolver.fail_missing_metadata();
            let (_budget, _cancellation, execution) =
                make_execution("ods-formula-value-matrix-reference-operand");
            let expression = parse(&source);
            let result = evaluate_at(
                &expression,
                &resolver,
                &execution,
                0,
                0,
                Mode::Matrix,
                &Limits::default(),
            );
            match result {
                Err(EvaluationFailure::Unsupported(UnsupportedKind::ReferenceOperator)) => {},
                Err(error) => failures.push(format!("{source:?} returned {error:?}")),
                Ok(result) => failures.push(format!(
                    "{source:?} unexpectedly succeeded with {:?}",
                    result.value()
                )),
            }
            assert_eq!(
                resolver.missing_metadata_calls(),
                0,
                "unselected Missing branch was inspected for {source:?}"
            );
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn finite_inner_condition_does_not_cache_out_of_shape_broadcast_positions() {
    const ROWS: usize = 33;
    let mut resolver = FixtureResolver::new();
    resolver.sheets[0].1 = SheetExtent::new(ROWS, 8);
    for row in 0..ROWS {
        resolver.set("Main", row, 1, FixtureCell::Number((100 + row) as f64));
    }
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-finite-condition-alias");
    let expression = parse("=IF({TRUE();FALSE()};IF(IF({TRUE()|TRUE()};1;0);1;2);[.B1:.B33])");
    let limits = Limits::default()
        .with_max_array_cells(66)
        .with_max_stack_entries(32);
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("finite condition shape should not require per-output aliases");
    let array = result.as_array().expect("broadcast result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (ROWS, 2));
    for row in 0..ROWS {
        if row < 2 {
            assert_number(array.get(row * 2).expect("true-column value"), 1.0);
        } else {
            assert_error(
                array.get(row * 2).expect("out-of-shape true-column value"),
                ScalarError::NotAvailable,
            );
        }
        assert_number(
            array.get(row * 2 + 1).expect("reference-column value"),
            (100 + row) as f64,
        );
    }
    assert_eq!(resolver.reads(), ROWS);
    let expected_order: Vec<_> = (0..ROWS).map(|row| ("Main".to_owned(), row, 1)).collect();
    assert_eq!(resolver.read_order(), expected_order);
}

#[test]
fn reference_operand_error_handlers_use_resolved_reference_errors() {
    for operand in [
        "IFERROR([Missing.A1];[.C1:.D1])",
        "IFERROR(([Missing.A1]:[.A1]);[.C1:.D1])",
        "IFERROR(([.A1]![.B1]);[.C1:.D1])",
        "IFERROR(XOR(([.A1]~[.B1])![.A1]);[.C1:.D1])",
        "IFERROR(([.A1]~[.B1]);[.C1:.D1])",
        "IFERROR(-([.A1]~[.B1]);[.C1:.D1])",
        "IFNA((#N/A+([.A1]~[.B1]));[.C1:.D1])",
    ] {
        let source = format!("=IF({{TRUE()}};([.A1]:{operand});0)");
        let mut resolver = FixtureResolver::new();
        for (column, value) in [10.0, 11.0, 12.0, 13.0].into_iter().enumerate() {
            resolver.set("Main", 0, column, FixtureCell::Number(value));
        }
        let (_budget, _cancellation, execution) =
            make_execution("ods-formula-value-resolved-reference-error-operand");
        let expression = parse(&source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should catch the reference error: {error:?}"));
        let array = result.as_array().expect("selected reference range array");
        assert_eq!((array.shape().rows(), array.shape().columns()), (1, 4));
        assert_array_numbers(array, &[10.0, 11.0, 12.0, 13.0]);
        assert_eq!(resolver.reads(), 4);
    }
}

#[test]
fn lazy_reference_operands_provide_full_range_shape_and_skip_missing_alternatives() {
    for source in [
        "=IF({TRUE()};([.A1]:IF(TRUE();[.C1:.D1];[Missing.A1:.Z100]));0)",
        "=IF({TRUE()};([.A1]:IFERROR(1/0;[.C1:.D1]));0)",
        "=IF({TRUE()};([.A1]:IFNA(#N/A;[.C1:.D1]));0)",
        "=IF({TRUE()};([.A1]:IF(IF(TRUE();TRUE();{TRUE();FALSE()});[.C1:.D1];[Missing.A1]));0)",
    ] {
        let mut resolver = FixtureResolver::new();
        for (column, value) in [10.0, 11.0, 12.0, 13.0].into_iter().enumerate() {
            resolver.set("Main", 0, column, FixtureCell::Number(value));
        }
        resolver.fail_missing_metadata();
        let (_budget, _cancellation, execution) =
            make_execution("ods-formula-value-lazy-reference-operand");
        let expression = parse(source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| {
            panic!("{source:?} should preserve evaluated reference shape: {error}")
        });
        let array = result
            .as_array()
            .expect("range over selected reference should produce an array");
        assert_eq!((array.shape().rows(), array.shape().columns()), (1, 4));
        assert_array_numbers(array, &[10.0, 11.0, 12.0, 13.0]);
        assert_eq!(
            resolver.read_order(),
            vec![
                ("Main".to_owned(), 0, 0),
                ("Main".to_owned(), 0, 1),
                ("Main".to_owned(), 0, 2),
                ("Main".to_owned(), 0, 3),
            ],
            "only the selected reference range should be materialized: {source:?}"
        );
        assert_eq!(resolver.missing_metadata_calls(), 0, "{source:?}");
    }
}

#[test]
fn lazy_reference_lists_and_multiplane_references_are_not_flattened() {
    for source in [
        "=IF({TRUE()};IF(TRUE();([.A1]~[.C1]);[Missing.A1:.Z100]);0)",
        "=IF({TRUE()};IFERROR(1/0;[Main.A1:Archive.A1]);0)",
    ] {
        let mut resolver = FixtureResolver::new();
        resolver.fail_missing_metadata();
        let (_budget, _cancellation, execution) =
            make_execution("ods-formula-value-lazy-reference-refusal");
        let expression = parse(source);
        let error = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &Limits::default(),
        )
        .expect_err("non-rectangular or multi-plane reference results must stay typed");
        assert!(matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::Reference)
        ));
        assert_eq!(
            resolver.reads(),
            0,
            "reference refusal must precede reads: {source:?}"
        );
        assert_eq!(resolver.missing_metadata_calls(), 0, "{source:?}");
    }
}

#[test]
fn local_reference_values_keep_empty_text_logical_and_formula_errors_distinct() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Empty);
    resolver.set("Main", 1, 0, FixtureCell::Number(0.0));
    resolver.set("Main", 2, 0, FixtureCell::Text(String::new()));
    resolver.set("Main", 3, 0, FixtureCell::Logical(false));
    resolver.set("Main", 4, 0, FixtureCell::Error(ScalarError::NotAvailable));
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-kinds");
    let limits = Limits::default();

    let expression = parse("=[.A1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("empty reference should evaluate");
    assert!(matches!(result.value(), Value::Empty));

    let expression = parse("=0+[.A1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("empty reference should coerce to zero");
    assert_number(result.value(), 0.0);

    let expression = parse("=\"\"&[.A1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("empty reference should concatenate as empty text");
    assert!(matches!(result.value(), Value::Text(text) if text.is_empty()));

    let expression = parse("=[.A2]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("numeric reference should evaluate");
    assert_number(result.value(), 0.0);

    let expression = parse("=[.A3]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("text reference should evaluate");
    assert!(matches!(result.value(), Value::Text(text) if text.is_empty()));

    let expression = parse("=[.A4]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("logical reference should evaluate");
    assert!(matches!(result.value(), Value::Logical(false)));

    let expression = parse("=[.A5]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("formula errors remain values");
    assert_error(result.value(), ScalarError::NotAvailable);
}

#[test]
fn matrix_reference_materialization_preserves_row_major_cell_order() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(1.0));
    resolver.set("Main", 0, 1, FixtureCell::Text("two".to_owned()));
    resolver.set("Main", 1, 0, FixtureCell::Logical(true));
    resolver.set("Main", 1, 1, FixtureCell::Empty);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-matrix");
    // Arithmetic forces a reference into a rectangular value while retaining
    // the resolver's row-major read order.  A bare reference is tested below
    // as a first-class reference view instead of silently materializing it.
    let expression = parse("=0+[.A1:.B2]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("a finite local range should materialize");
    let shape = result.shape().expect("range result shape");
    assert_eq!((shape.rows(), shape.columns()), (2, 2));
    let array = result.as_array().expect("range result array");
    assert_eq!(array.len(), 4);
    assert_number(array.get(0).unwrap(), 1.0);
    assert_error(array.get(1).unwrap(), ScalarError::Value);
    assert_number(array.get(2).unwrap(), 1.0);
    assert_number(array.get(3).unwrap(), 0.0);
    assert_eq!(resolver.reads(), 4);
    assert_eq!(
        resolver.read_order(),
        vec![
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 0, 1),
            ("Main".to_owned(), 1, 0),
            ("Main".to_owned(), 1, 1),
        ]
    );
}

#[test]
fn scalar_implicit_intersection_uses_current_row_or_column_and_rejects_ambiguous_2d() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(10.0));
    resolver.set("Main", 1, 0, FixtureCell::Number(20.0));
    resolver.set("Main", 2, 0, FixtureCell::Number(30.0));
    resolver.set("Main", 0, 1, FixtureCell::Number(40.0));
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-intersection");
    let limits = Limits::default();

    // B2 is row 1 and column 1 in zero-based coordinates.  A1:A3 selects
    // A2 by row, while A1:C1 selects B1 by column.
    let vertical = parse("=0+[.A1:.A3]");
    let result = evaluate_at(
        &vertical,
        &resolver,
        &execution,
        1,
        1,
        Mode::Scalar,
        &limits,
    )
    .expect("vertical range should have one row intersection");
    assert_number(result.value(), 20.0);

    let horizontal = parse("=0+[.A1:.C1]");
    let result = evaluate_at(
        &horizontal,
        &resolver,
        &execution,
        1,
        1,
        Mode::Scalar,
        &limits,
    )
    .expect("horizontal range should have one column intersection");
    assert_number(result.value(), 40.0);

    let rectangle = parse("=0+[.A1:.B2]");
    let result = evaluate_at(
        &rectangle,
        &resolver,
        &execution,
        1,
        1,
        Mode::Scalar,
        &limits,
    )
    .expect("ambiguous two-dimensional intersection should be a formula value");
    assert_error(result.value(), ScalarError::NotAvailable);
}

#[test]
fn matrix_scalar_and_vector_broadcasting_follows_odf_shapes() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-broadcast");
    let limits = Limits::default();

    let expression = parse("={1;2|3;4}");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("rectangular inline array should evaluate");
    let shape = result.shape().expect("array shape");
    assert_eq!((shape.rows(), shape.columns()), (2, 2));
    assert_array_numbers(result.as_array().unwrap(), &[1.0, 2.0, 3.0, 4.0]);

    let expression = parse("={1}+{10;20|30;40}");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("scalar should broadcast over a matrix");
    assert_array_numbers(result.as_array().unwrap(), &[11.0, 21.0, 31.0, 41.0]);

    let expression = parse("={1;2}+{10|20}");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("one-row and one-column arrays should use outer broadcasting");
    assert_array_numbers(result.as_array().unwrap(), &[11.0, 12.0, 21.0, 22.0]);

    // Two row vectors have incompatible widths.  The profile retains the
    // larger result shape and represents the missing tail as #N/A.
    let expression = parse("={1;2}+{3;4;5}");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("incompatible broadcast positions should remain formula values");
    let array = result.as_array().expect("broadcast result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 3));
    assert_number(array.get(0).unwrap(), 4.0);
    assert_number(array.get(1).unwrap(), 6.0);
    assert_error(array.get(2).unwrap(), ScalarError::NotAvailable);
}

#[test]
fn ragged_inline_arrays_are_retained_by_parser_but_refused_by_rectangular_value_profile() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-ragged");
    let expression = parse("={1;2|3}");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("the non-evaluating parser's ragged array needs a rectangular evaluator");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Array)
    ));
}

#[test]
fn range_intersection_retains_one_area_and_empty_intersection_is_null() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 1, 1, FixtureCell::Number(22.0));
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-range-ops");

    let expression = parse("=([.A1:.B2]![.B2:.C3])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("overlapping ranges should intersect");
    let reference = result
        .as_reference()
        .expect("intersection should remain a first-class reference");
    assert_eq!(reference.areas().len(), 1);
    assert_eq!(reference.areas()[0].starts(), [0, 1, 1]);
    assert_eq!(reference.areas()[0].extent(), [1, 1, 1]);

    let expression = parse("=([.A1:.A1]![.B1:.B1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect("empty reference intersection should remain a formula value");
    assert_error(result.value(), ScalarError::Null);
}

#[test]
fn reference_lists_keep_source_order_duplicates_and_three_dimensional_sheet_order() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-reference-list");
    let limits = Limits::default();

    let expression = parse("=([.A1]~[.B1]~[.A1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("a union should retain its ordered reference records");
    let list = result
        .as_reference_list()
        .expect("union should expose a first-class reference list");
    assert_eq!(list.len(), 3);
    for (index, expected_column) in [0, 1, 0].into_iter().enumerate() {
        let reference = list.get(index).expect("reference-list entry");
        assert_eq!(reference.len(), 1);
        assert_eq!(reference.areas()[0].starts(), [0, 0, expected_column]);
        assert_eq!(reference.areas()[0].extent(), [1, 1, 1]);
        assert!(reference.reference().is_some());
    }

    // The local range crosses two ordered worksheet names.  The resolver's
    // sheet order supplies the first (sheet) axis of the retained cuboid.
    let expression = parse("=[Data.A1:Archive.A1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("ordered local 3-D range should resolve");
    let reference = result.as_reference().expect("3-D range reference");
    assert_eq!(reference.areas().len(), 1);
    assert_eq!(reference.areas()[0].starts(), [1, 0, 0]);
    assert_eq!(reference.areas()[0].extent(), [2, 1, 1]);
}

#[test]
fn intersections_keep_cross_product_records_and_drop_empty_phantoms() {
    let resolver = FixtureResolver::new();
    let limits = Limits::default();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-intersection-list");
    let expression = parse("=([.A1]~[.B1])!([.A1]~[.B1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("surviving intersections should remain a reference list");
    let list = result
        .as_reference_list()
        .expect("intersection of lists should expose list records");
    assert_eq!(list.len(), 2, "one record is required per surviving pair");
    for (index, expected_column) in [0, 1].into_iter().enumerate() {
        let reference = list.get(index).expect("cross-product list entry");
        assert_eq!(reference.len(), 1);
        assert_eq!(reference.areas()[0].starts(), [0, 0, expected_column]);
        assert_eq!(reference.areas()[0].extent(), [1, 1, 1]);
    }

    let expression = parse("=([.A1]~[.B1])!([.C1]~[.D1])");
    for mode in [Mode::Matrix, Mode::Scalar] {
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            mode,
            &Limits::default(),
        )
        .unwrap_or_else(|error| {
            panic!("all-empty intersection should be a formula error: {error}")
        });
        assert_error(result.value(), ScalarError::Null);
    }

    // Duplicate union records remain distinct after intersection.  Each
    // duplicate must retain one plane; merging them into one record or
    // appending both planes to either record changes the reference-list
    // identity even though the geometry is equal.
    let expression = parse("=([.A1]~[.A1])![.A1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("duplicate union records should survive intersection");
    let list = result
        .as_reference_list()
        .expect("duplicate intersections should remain a reference list");
    assert_eq!(list.len(), 2);
    for index in 0..list.len() {
        let reference = list.get(index).expect("duplicate intersection record");
        assert_eq!(reference.areas().len(), 1);
        assert_eq!(reference.areas()[0].starts(), [0, 0, 0]);
        assert_eq!(reference.areas()[0].extent(), [1, 1, 1]);
    }
    assert_eq!(
        resolver.reads(),
        0,
        "reference operators must not fetch cells"
    );

    // A 3-D left record intersects the middle sheet of the range, while a
    // second left record already names that same physical plane.  Both
    // record pairs survive in source order and each owns exactly one plane;
    // the 3-D plane must not be attributed to the neighboring record.
    let expression = parse("=([Main.A1:Archive.A1]~[Data.A1])![Data.A1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("nested 3-D intersections should retain pair identity");
    let list = result
        .as_reference_list()
        .expect("nested 3-D intersections should remain a reference list");
    assert_eq!(list.len(), 2);
    for index in 0..list.len() {
        let reference = list.get(index).expect("nested 3-D intersection record");
        assert_eq!(reference.areas().len(), 1);
        assert_eq!(reference.areas()[0].starts(), [1, 0, 0]);
        assert_eq!(reference.areas()[0].extent(), [1, 1, 1]);
    }
    assert_eq!(
        resolver.reads(),
        0,
        "3-D reference operators must not fetch cells"
    );
}

#[test]
fn three_dimensional_intersection_lists_preserve_pair_order_and_cuboids() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-intersection-3d-list");
    let expression = parse(
        "=([Main.A1:Archive.C3]~[Main.E5:Archive.G7])!([Main.B2:Archive.D4]~[Main.F6:Archive.H8])",
    );
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("overlapping 3-D list references should intersect");
    let list = result
        .as_reference_list()
        .expect("3-D intersection should expose ordered list records");
    assert_eq!(
        list.len(),
        2,
        "only the two same-order pairs overlap across all three sheets"
    );

    // Each pair remains one record with one public area whose sheet extent is
    // three planes.  Splitting the planes into separate records would lose
    // the identity of the original left/right pair.
    for (index, (row, column)) in [(1, 1), (5, 5)].into_iter().enumerate() {
        let reference = list.get(index).expect("ordered 3-D intersection record");
        assert_eq!(reference.len(), 1);
        assert_eq!(reference.areas()[0].starts(), [0, row, column]);
        assert_eq!(reference.areas()[0].extent(), [3, 2, 2]);
    }
}

#[test]
fn range_over_reference_lists_collapses_to_one_bounding_three_dimensional_cuboid() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-range-list-cuboid");
    let expression = parse("=([Data.A1]~[Archive.C3]):[Main.B2]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("range over list operands should resolve a cuboid");
    let reference = result
        .as_reference()
        .expect("the range operator should collapse list operands");
    assert_eq!(reference.areas().len(), 1);
    assert_eq!(reference.areas()[0].starts(), [0, 0, 0]);
    assert_eq!(reference.areas()[0].extent(), [3, 3, 3]);
}

#[test]
fn reference_cell_limits_cover_direct_and_cumulative_fetches() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(1.0));
    let (budget, _cancellation, execution) = make_execution("ods-formula-value-cell-limit-zero");
    let expression = parse("=[.A1]");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default().with_max_reference_cells(0),
    )
    .expect_err("a zero reference-cell allowance must refuse before reading");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(resolver.reads(), 0);
    assert_eq!(budget.used(Resource::Memory), 0);

    let mut resolver = FixtureResolver::new();
    for row in 0..8 {
        resolver.set("Main", row, 0, FixtureCell::Number(1.0));
        resolver.set("Main", row, 1, FixtureCell::Number(1.0));
    }
    let (budget, _cancellation, execution) =
        make_execution("ods-formula-value-cell-limit-cumulative");
    let expression = parse("=AND([.A1:.A8];[.B1:.B8])");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default().with_max_reference_cells(8),
    )
    .expect_err("two eight-cell references must exceed a cumulative limit of eight");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert!(
        resolver.reads() <= 8,
        "the cumulative limit must prevent a ninth provider read"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn provider_cancellation_after_read_or_final_source_version_is_not_published() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(7.0));
    let (budget, cancellation, execution) = make_execution("ods-formula-value-cancel-after-read");
    resolver.cancel_after_read(&cancellation);
    let expression = parse("=[.A1]");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("cancellation requested by the final provider read must fence publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);

    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(8.0));
    let version = SourceVersion::new(0x4f44_5302, 0);
    resolver.source_changes(version, version);
    let (budget, cancellation, execution) =
        make_execution("ods-formula-value-cancel-final-source-version");
    resolver.cancel_on_final_source_version(&cancellation);
    let expression = parse("=[.A1]");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("cancellation during the final source sample must fence publication");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn explicit_cells_outside_the_finite_extent_refuse_before_provider_read() {
    let mut resolver = FixtureResolver::new();
    resolver.fail_after(0);
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-explicit-cell-extent");
    let expression = parse("=[.Z100]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect("an out-of-extent explicit cell should become #REF! without a read");
    assert_error(result.value(), ScalarError::Reference);
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn sheet_qualified_references_keep_local_metadata_and_external_sources_stay_inert() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-sheets");
    let limits = Limits::default();

    let expression = parse("=['Data'.A1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("quoted local sheet reference should remain resolvable");
    let reference = result.as_reference().expect("local sheet reference");
    assert_eq!(reference.areas()[0].starts(), [1, 0, 0]);
    assert_eq!(reference.areas()[0].extent(), [1, 1, 1]);
    match reference.reference() {
        Some(Reference::Local(Address::Cell(endpoint))) => {
            let SheetSelector::Explicit(locator) = &endpoint.sheet else {
                panic!("quoted local reference lost its explicit sheet locator");
            };
            assert_eq!(locator.sheet.name, "Data");
            assert!(locator.sheet.quoted);
            assert!(matches!(&endpoint.value, EndpointValue::Cell(_)));
        },
        other => panic!("local reference lost its parsed owner: {other:?}"),
    }

    let expression = parse("=['https://example.invalid/book.ods'#.A1]");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect_err("external references must not trigger resolver/network access");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Reference)
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn missing_sheet_is_a_formula_reference_error_without_a_resolver_read() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-missing-sheet");
    let expression = parse("=[Missing.A1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect("a missing sheet is a formula #REF! value");
    assert_error(result.value(), ScalarError::Reference);
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn logical_aggregates_filter_empty_and_text_while_xor_iterates_matrix_cells() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Empty);
    resolver.set("Main", 0, 1, FixtureCell::Text("1".to_owned()));
    resolver.set("Main", 0, 2, FixtureCell::Logical(true));
    resolver.set("Main", 0, 3, FixtureCell::Number(1.0));
    resolver.set("Main", 1, 0, FixtureCell::Number(0.0));
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-logical");
    let limits = Limits::default();

    // NumberSequenceList ignores Empty/Text/distinguished Logical and keeps
    // the numeric member.  The nonzero numeric member makes both identities
    // true in this fixture.
    for source in ["=AND([.A1:.D1])", "=OR([.A1:.D1])"] {
        let expression = parse(source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Scalar,
            &limits,
        )
        .unwrap_or_else(|error| panic!("{source:?} should evaluate: {error}"));
        assert!(matches!(result.value(), Value::Logical(true)), "{source:?}");
    }

    let expression = parse("=AND([.A1:.A1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("all-filtered AND should use its profile identity");
    assert!(matches!(result.value(), Value::Logical(true)));

    let expression = parse("=OR([.A1:.A1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("all-filtered OR should use its profile identity");
    assert!(matches!(result.value(), Value::Logical(false)));

    let expression = parse("=AND([.A2:.A2])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect("numeric false should be retained by AND");
    assert!(matches!(result.value(), Value::Logical(false)));

    let expression = parse("=XOR({TRUE();FALSE()};{FALSE();TRUE()})");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("XOR should map one logical result per matrix position");
    let array = result.as_array().expect("XOR matrix result");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert!(matches!(array.get(0), Some(Value::Logical(true))));
    assert!(matches!(array.get(1), Some(Value::Logical(true))));

    let expression = parse("=XOR({TRUE();TRUE()};TRUE())");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("scalar logical should broadcast over XOR matrix cells");
    let array = result.as_array().expect("XOR broadcast result");
    assert!(matches!(array.get(0), Some(Value::Logical(false))));
    assert!(matches!(array.get(1), Some(Value::Logical(false))));
}

#[test]
fn logical_aggregates_flatten_inline_arrays_and_keep_reference_sequence_rules() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Empty);
    resolver.set("Main", 0, 1, FixtureCell::Logical(true));
    resolver.set("Main", 0, 2, FixtureCell::Text("TRUE".to_owned()));
    resolver.set(
        "Main",
        0,
        3,
        FixtureCell::Error(ScalarError::DivisionByZero),
    );
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-logical-arrays");
    let limits = Limits::default();

    // Inline arrays are a NumberSequenceList only after the value evaluator
    // has selected the profile's row-major array aggregation.  Both rows and
    // columns participate in the same logical fold.
    let expression = parse("=AND({TRUE();TRUE()|TRUE();FALSE()})");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("AND should aggregate an inline matrix");
    assert!(matches!(result.value(), Value::Logical(false)));

    let expression = parse("=OR({FALSE();FALSE()|FALSE();TRUE()})");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("OR should aggregate an inline matrix");
    assert!(matches!(result.value(), Value::Logical(true)));

    // Unary plus materializes a reference into an array while retaining its
    // Empty/Text/Logical element kinds.  Empty is false, distinguished
    // Logical values are converted directly, and Text follows the existing
    // scalar to_logical profile (#VALUE! for this non-boolean spelling).
    let expression = parse("=AND(+[.A1:.B1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("an Empty element should participate as false");
    assert!(matches!(result.value(), Value::Logical(false)));

    let expression = parse("=OR(+[.A1:.B1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("an Empty element should participate as false");
    assert!(matches!(result.value(), Value::Logical(true)));

    let expression = parse("=AND(+[.C1:.C1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("text conversion should remain a formula result");
    assert_error(result.value(), ScalarError::Value);

    let expression = parse("=OR(+[.D1:.D1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("array formula errors should remain values");
    assert_error(result.value(), ScalarError::DivisionByZero);
}

#[test]
fn matrix_if_and_error_handlers_select_only_the_matching_array_elements() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-matrix-lazy");
    let limits = Limits::default();

    let expression = parse("=IF({TRUE();FALSE()};{1;2};{3;4})");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("IF should select one branch per matrix position");
    let array = result.as_array().expect("IF matrix result");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[1.0, 4.0]);

    let expression = parse("=IFERROR({#N/A;1};{2;3})");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("IFERROR should handle each array error independently");
    assert_array_numbers(
        result.as_array().expect("IFERROR matrix result"),
        &[2.0, 1.0],
    );

    let expression = parse("=IFNA({#N/A;1};{2;3})");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("IFNA should handle only #N/A array errors");
    assert_array_numbers(
        result.as_array().expect("IFNA matching matrix result"),
        &[2.0, 1.0],
    );

    let expression = parse("=IFNA({#DIV/0!;1};{2;3})");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("IFNA should retain non-#N/A array errors");
    let array = result.as_array().expect("IFNA nonmatching matrix result");
    assert_error(
        array.get(0).expect("first IFNA element"),
        ScalarError::DivisionByZero,
    );
    assert_number(array.get(1).expect("second IFNA element"), 1.0);

    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(10.0));
    resolver.set("Main", 0, 1, FixtureCell::Number(11.0));
    resolver.set("Main", 1, 0, FixtureCell::Number(20.0));
    resolver.set("Main", 1, 1, FixtureCell::Number(21.0));
    let expression = parse("=IF({TRUE();FALSE()};[.A1:.B1];[.A2:.B2])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("IF should lazily materialize selected reference elements");
    assert_array_numbers(
        result.as_array().expect("lazy IF reference matrix result"),
        &[10.0, 21.0],
    );
    assert_eq!(
        resolver.read_order(),
        vec![("Main".to_owned(), 0, 0), ("Main".to_owned(), 1, 1)],
        "unselected B1/A2 cells must not be read"
    );
}

#[test]
fn matrix_lazy_branches_infer_reference_and_nested_shapes() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(10.0));
    resolver.set("Main", 0, 1, FixtureCell::Number(11.0));
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-shape");
    let limits = Limits::default();

    // A one-cell condition does not limit the selected branch to one cell.
    // The direct reference, an infix around that reference, and a nested IF
    // must all carry the branch's 1x2 geometry into the outer result.
    for source in [
        "=IF({TRUE()};[.A1:.B1];0)",
        "=IF({TRUE()};0+[.A1:.B1];0)",
        "=IF({TRUE()};IF({TRUE()};[.A1:.B1];0);0)",
    ] {
        let expression = parse(source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &limits,
        )
        .unwrap_or_else(|error| panic!("{source:?} should preserve branch shape: {error}"));
        let array = result.as_array().expect("lazy branch should be an array");
        assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
        assert_array_numbers(array, &[10.0, 11.0]);
    }
    assert_eq!(
        resolver.read_order(),
        vec![
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 0, 1),
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 0, 1),
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 0, 1),
        ]
    );
}

#[test]
fn unselected_array_branches_do_not_contribute_shape_or_reference_metadata() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-lazy-shape");
    let expression = parse("=IF({TRUE()};{1};{2|3})");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("the unselected larger array must not widen the result");
    let array = result.as_array().expect("IF result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 1));
    assert_number(array.get(0).unwrap(), 1.0);

    // Shape discovery must not probe an unselected reference.  The resolver
    // turns any Missing-sheet metadata request into an operational failure,
    // making a successful result evidence that no extent/index lookup was
    // attempted.
    let mut resolver = FixtureResolver::new();
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-lazy-reference-shape");
    let expression = parse("=IF({TRUE()};1;[Missing.A1:.Z100])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("an unselected missing-sheet reference must remain untouched");
    let array = result.as_array().expect("scalar IF result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 1));
    assert_number(array.get(0).unwrap(), 1.0);
    assert_eq!(resolver.reads(), 0);
    assert_eq!(resolver.missing_metadata_calls(), 0);

    let mut resolver = FixtureResolver::new();
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-lazy-handler-shape");
    let expression = parse("=IFERROR({1};[Missing.A1:.Z100])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("an unused IFERROR alternative must remain untouched");
    let array = result.as_array().expect("IFERROR result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 1));
    assert_number(array.get(0).unwrap(), 1.0);
    assert_eq!(resolver.reads(), 0);
    assert_eq!(resolver.missing_metadata_calls(), 0);
}

#[test]
fn nested_matrix_handlers_preserve_selected_shapes_and_skip_unused_references() {
    // The first two cases put IFERROR/IFNA inside the outer IF condition.  The
    // last two put the matrix IF inside the error handler.  Every Missing
    // reference is unreachable, so even shape planning must avoid its sheet
    // metadata; only the selected two-cell Main range may be read.
    for source in [
        "=IF(IFERROR({TRUE();TRUE()};[Missing.A1:.Z100]);[.A1:.B1];[Missing.C1:.Z100])",
        "=IF(IFNA({TRUE();TRUE()};[Missing.A1:.Z100]);[.A1:.B1];[Missing.C1:.Z100])",
        "=IFERROR(IF({TRUE();TRUE()};[.A1:.B1];[Missing.C1:.Z100]);[Missing.D1:.Z100])",
        "=IFNA(IF({TRUE();TRUE()};[.A1:.B1];[Missing.C1:.Z100]);[Missing.D1:.Z100])",
    ] {
        let mut resolver = FixtureResolver::new();
        resolver.set("Main", 0, 0, FixtureCell::Number(10.0));
        resolver.set("Main", 0, 1, FixtureCell::Number(11.0));
        resolver.fail_missing_metadata();
        let (_budget, _cancellation, execution) =
            make_execution("ods-formula-value-nested-matrix-handlers");
        let expression = parse(source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should keep nested laziness: {error}"));
        let array = result
            .as_array()
            .expect("nested matrix handler result array");
        assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
        assert_array_numbers(array, &[10.0, 11.0]);
        assert_eq!(resolver.reads(), 2, "only selected Main cells may be read");
        assert_eq!(resolver.missing_metadata_calls(), 0);
        assert_eq!(
            resolver.read_order(),
            vec![("Main".to_owned(), 0, 0), ("Main".to_owned(), 0, 1)]
        );
    }
}

#[test]
fn nested_matrix_shape_planning_stays_selected_through_handlers_and_wrappers() {
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-nested-matrix-shape-planning");

    // The inner IF has a wider false branch, but its true branch is the only
    // branch selected by the outer and inner conditions.  Its unselected
    // geometry must not widen the outer result.
    let resolver = FixtureResolver::new();
    let expression = parse("=IF({TRUE()};IF({TRUE()};{1};{2|3});0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("nested IF should retain only the selected branch shape");
    let array = result.as_array().expect("nested IF result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 1));
    assert_number(array.get(0).expect("nested IF result element"), 1.0);

    // The same selective shape rule applies when the nested handler is the
    // selected branch.  A metadata failure on the unreachable reference
    // proves that planning does not descend into that branch.
    let mut resolver = FixtureResolver::new();
    resolver.fail_missing_metadata();
    let expression = parse("=IF({TRUE()};IF({TRUE()};1;[Missing.A1:.Z100]);0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("nested IF should not inspect its unreachable reference branch");
    let array = result.as_array().expect("nested reference IF result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 1));
    assert_number(array.get(0).expect("nested reference IF result"), 1.0);
    assert_eq!(resolver.reads(), 0);
    assert_eq!(resolver.missing_metadata_calls(), 0);

    // NOT carries the matrix shape of its child.  If it were planned as a
    // scalar wrapper, this two-cell selected branch would be truncated to a
    // one-cell result.
    let resolver = FixtureResolver::new();
    let expression = parse("=IF({TRUE();FALSE()};NOT({TRUE();FALSE()});0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("NOT should preserve its selected matrix shape");
    let array = result.as_array().expect("NOT matrix IF result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert!(matches!(array.get(0), Some(Value::Logical(false))));
    assert_number(array.get(1).expect("second NOT matrix result"), 0.0);
}

#[test]
fn provider_driven_nested_handlers_keep_selected_shape_and_skip_unselected_metadata() {
    // The inner condition is decided by a provider cell.  Once it selects the
    // two-cell branch, the inner IF's unreachable Missing reference must not
    // participate in the outer shape plan or metadata lookup.
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Logical(true));
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-provider-nested-if");
    let expression = parse("=IF({TRUE()};IF([.A1];{1;2};[Missing.A1:.Z100]);0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("provider-backed nested IF should select its true array");
    let array = result
        .as_array()
        .expect("provider-backed nested IF result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[1.0, 2.0]);
    assert_eq!(
        resolver.reads(),
        1,
        "only the provider condition may be read"
    );
    assert_eq!(resolver.missing_metadata_calls(), 0);
    assert_eq!(resolver.read_order(), vec![("Main".to_owned(), 0, 0)]);

    // A provider-backed formula error selects IFERROR's array fallback while
    // the outer IF keeps its Missing alternative unreachable.  This exercises
    // the opposite nested continuation: the handler decision is dynamic, but
    // its selected fallback shape must still reach the outer result.
    let mut resolver = FixtureResolver::new();
    resolver.set(
        "Main",
        0,
        0,
        FixtureCell::Error(ScalarError::DivisionByZero),
    );
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-provider-nested-iferror");
    let expression = parse("=IF({TRUE()};IFERROR([.A1];{3;4});[Missing.A1:.Z100])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("provider-backed IFERROR should select its fallback array");
    let array = result
        .as_array()
        .expect("provider-backed IFERROR result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[3.0, 4.0]);
    assert_eq!(resolver.reads(), 1, "only the provider error may be read");
    assert_eq!(resolver.missing_metadata_calls(), 0);
    assert_eq!(resolver.read_order(), vec![("Main".to_owned(), 0, 0)]);
}

#[test]
fn nested_iferror_static_division_error_preserves_fallback_shape() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-nested-static-iferror");
    let expression = parse("=IF({TRUE()};IFERROR(1/0;{3;4});0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("IFERROR should catch the static division error");
    let array = result
        .as_array()
        .expect("nested static IFERROR result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[3.0, 4.0]);
}

#[test]
fn nested_ifna_static_not_available_error_preserves_fallback_shape() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-nested-static-ifna");
    let expression = parse("=IF({TRUE()};IFNA(#N/A;{3;4});0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("IFNA should catch the static not-available error");
    let array = result.as_array().expect("nested static IFNA result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[3.0, 4.0]);
}

#[test]
fn nested_iferror_function_error_preserves_fallback_shape() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-nested-function-iferror");
    let expression = parse("=IF({TRUE()};IFERROR(NOT(\"bad\");{3;4});0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("IFERROR should catch NOT's scalar value error");
    let array = result
        .as_array()
        .expect("nested function IFERROR result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[3.0, 4.0]);
}

#[test]
fn nested_if_scalar_branch_uses_selected_condition_shape_only() {
    let mut resolver = FixtureResolver::new();
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-nested-scalar-shape");
    let expression = parse("=IF({TRUE()};IF({TRUE();TRUE()};1;[Missing.A1:.Z100]);0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("selected nested scalar branch should determine the result shape");
    let array = result
        .as_array()
        .expect("nested scalar branch result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[1.0, 1.0]);
    assert_eq!(resolver.reads(), 0);
    assert_eq!(resolver.missing_metadata_calls(), 0);
}

#[test]
fn nested_if_selected_branches_broadcast_to_their_combined_shape() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-nested-broadcast-shape");
    let expression = parse("=IF({TRUE()};IF({TRUE();FALSE()};{1;2};{3|4|5});0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("nested IF branches should broadcast by singleton dimensions");
    let array = result
        .as_array()
        .expect("nested broadcast branch result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (3, 2));
    // The 1x2 true branch repeats down rows and the 3x1 false branch repeats
    // across columns; the inner condition selects the first column from the
    // former and the second column from the latter.
    assert_array_numbers(array, &[1.0, 3.0, 1.0, 4.0, 1.0, 5.0]);
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn independently_evaluated_views_compare_structurally_in_source_order() {
    for (left, right, equal) in [
        ("={1;2}", "={1;2}", true),
        ("={1;2}", "={1;3}", false),
        ("={1;2}", "={1|2}", false),
        ("=([.A1]~[.B1])", "=([.A1]~[.B1])", true),
        ("=([.A1]~[.B1])", "=([.B1]~[.A1])", false),
        ("=([.A1]~[.A1])", "=([.A1]~[.B1])", false),
    ] {
        let resolver = FixtureResolver::new();
        let (_budget, _cancellation, execution) = make_execution("ods-value-view-equality");
        let left_expression = parse(left);
        let right_expression = parse(right);
        let limits = Limits::default();
        let left_value = evaluate_at(
            &left_expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &limits,
        )
        .expect("left evaluation");
        let right_value = evaluate_at(
            &right_expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &limits,
        )
        .expect("right evaluation");
        assert_eq!(
            left_value.value() == right_value.value(),
            equal,
            "{left} vs {right}"
        );
    }
}

#[test]
fn reference_geometry_admission_precedes_matrix_and_scalar_cell_reads() {
    for mode in [Mode::Matrix, Mode::Scalar] {
        let mut resolver = FixtureResolver::new();
        resolver.set("Main", 3, 0, FixtureCell::Number(7.0));
        let (budget, _cancellation, execution) = make_execution("ods-reference-admission");
        let expression = parse("=[.A1:.A8]");
        let error = evaluate_at(
            &expression,
            &resolver,
            &execution,
            3,
            0,
            mode,
            &Limits::default().with_max_reference_cells(1),
        )
        .expect_err("full reference geometry exceeds admission limit");
        assert!(matches!(error, EvaluationFailure::ResourceLimit(limit)
            if limit.resource == Resource::Objects && limit.observed == 8 && limit.limit == 1));
        assert_eq!(resolver.reads(), 0);
        assert_eq!(budget.used(Resource::Memory), 0);

        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            3,
            0,
            mode,
            &Limits::default()
                .with_max_reference_cells(8)
                .with_max_array_cells(1),
        )
        .expect("admitted reference needs no array materialization");
        match mode {
            Mode::Matrix => {
                assert!(matches!(result.value(), Value::Reference(_)));
                assert_eq!(resolver.reads(), 0);
            },
            Mode::Scalar => {
                assert_eq!(result.value(), Value::Number(7.0));
                assert_eq!(resolver.reads(), 1);
            },
            _ => unreachable!("the fixture enumerates only matrix and scalar modes"),
        }
    }
}

#[test]
fn composed_array_conditions_keep_shape_through_operators_and_lazy_handlers() {
    let cases: [(&str, &[f64]); 4] = [
        (
            "=IF({TRUE()};IF(({TRUE();FALSE()}+0)=1;7;9);0)",
            &[7.0, 9.0],
        ),
        (
            "=IF({TRUE()};IF(IF({TRUE();TRUE()};TRUE();[Missing.A1:.Z100]);7;9);0)",
            &[7.0, 7.0],
        ),
        (
            "=IF({TRUE()};IF(IFERROR({TRUE();#DIV/0!};FALSE());7;9);0)",
            &[7.0, 9.0],
        ),
        (
            "=IF({TRUE()};IF(IFNA({TRUE();#N/A};FALSE());7;9);0)",
            &[7.0, 9.0],
        ),
    ];
    for (formula, expected) in cases {
        let mut resolver = FixtureResolver::new();
        resolver.fail_missing_metadata();
        let (_budget, _cancellation, execution) =
            make_execution("ods-formula-value-composed-array-condition");
        let expression = parse(formula);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{formula}: {error}"));
        let array = result.as_array().expect("composed condition result array");
        assert_eq!(
            (array.shape().rows(), array.shape().columns()),
            (1, 2),
            "{formula}",
        );
        assert_array_numbers(array, expected);
        assert_eq!(resolver.reads(), 0, "{formula}");
        assert_eq!(resolver.missing_metadata_calls(), 0, "{formula}");
    }
}

#[test]
fn computed_nested_conditions_preserve_selected_shape_without_metadata_probe() {
    // A computed scalar comparison must still let the selected nested array
    // contribute its shape.  The unreachable reference is configured to fail
    // at metadata lookup, so successful evaluation proves it was not planned.
    let mut resolver = FixtureResolver::new();
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-computed-nested-condition");
    let expression = parse("=IF({TRUE()};IF(1=1;{1;2};[Missing.A1:.Z100]);0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("computed nested IF should select its true array");
    let array = result
        .as_array()
        .expect("computed nested comparison result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[1.0, 2.0]);
    assert_eq!(resolver.reads(), 0);
    assert_eq!(resolver.missing_metadata_calls(), 0);

    // A provider-backed aggregate condition exercises the same shape path
    // after a real cell read rather than a literal/comparison shortcut.
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(1.0));
    resolver.fail_missing_metadata();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-aggregate-nested-condition");
    let expression = parse("=IF({TRUE()};IF(AND([.A1]);{1;2};[Missing.A1:.Z100]);0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("provider-backed aggregate condition should select its true array");
    let array = result
        .as_array()
        .expect("aggregate nested condition result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[1.0, 2.0]);
    assert_eq!(resolver.reads(), 1, "the aggregate condition reads A1 once");
    assert_eq!(resolver.missing_metadata_calls(), 0);
    assert_eq!(resolver.read_order(), vec![("Main".to_owned(), 0, 0)]);
}

#[test]
fn nested_if_keeps_provider_range_condition_lazy_per_output_position() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Logical(true));
    resolver.set("Main", 0, 1, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-nested-provider-condition-lazy");
    let expression = parse("=IF({TRUE();FALSE()};IF([.A1:.B1];{1;2};{3;4});0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("nested IF should read only the selected condition position");
    let array = result.as_array().expect("nested provider IF result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[1.0, 0.0]);
    assert_eq!(resolver.reads(), 1, "the unsupported B1 must remain unread");
    assert_eq!(resolver.read_order(), vec![("Main".to_owned(), 0, 0)]);
}

#[test]
fn nested_iferror_keeps_provider_range_lazy_per_output_position() {
    let mut resolver = FixtureResolver::new();
    resolver.set(
        "Main",
        0,
        0,
        FixtureCell::Error(ScalarError::DivisionByZero),
    );
    resolver.set("Main", 0, 1, FixtureCell::Unsupported);
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-nested-provider-iferror-lazy");
    let expression = parse("=IF({TRUE();FALSE()};IFERROR([.A1:.B1];{5;6});0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect("nested IFERROR should read only the selected range position");
    let array = result
        .as_array()
        .expect("nested provider IFERROR result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
    assert_array_numbers(array, &[5.0, 0.0]);
    assert_eq!(resolver.reads(), 1, "the unsupported B1 must remain unread");
    assert_eq!(resolver.read_order(), vec![("Main".to_owned(), 0, 0)]);
}

#[test]
fn out_of_shape_positions_are_not_available_across_broadcast_and_lazy_handlers() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-missing-array");
    let limits = Limits::default();

    // The larger row vector supplies a third output position, but the first
    // operand has no corresponding cell.  That position is #N/A, rather than
    // a generic #VALUE! conversion failure.
    let expression = parse("={1;2}+{3;4;5}");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("binary broadcast should retain an out-of-shape position");
    let array = result.as_array().expect("binary broadcast array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 3));
    assert_number(array.get(0).unwrap(), 4.0);
    assert_number(array.get(1).unwrap(), 6.0);
    assert_error(array.get(2).unwrap(), ScalarError::NotAvailable);

    // IF selects a two-cell true branch at a three-cell condition shape.  The
    // third selected branch position is an out-of-shape #N/A.
    let expression = parse("=IF({TRUE();TRUE();TRUE()};{1;2};0)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("IF should retain a missing selected branch position");
    let array = result.as_array().expect("IF broadcast array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, 3));
    assert_number(array.get(0).unwrap(), 1.0);
    assert_number(array.get(1).unwrap(), 2.0);
    assert_error(array.get(2).unwrap(), ScalarError::NotAvailable);

    // The handler alternative is shorter than the input array.  Its missing
    // replacement at the final #N/A input position must stay #N/A as well.
    for name in ["IFERROR", "IFNA"] {
        let source = format!("={name}({{#N/A;2;#N/A}};{{7;8}})");
        let expression = parse(&source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &limits,
        )
        .unwrap_or_else(|error| panic!("{source:?} should project its handler: {error}"));
        let array = result.as_array().expect("handler broadcast array");
        assert_eq!((array.shape().rows(), array.shape().columns()), (1, 3));
        assert_number(array.get(0).unwrap(), 7.0);
        assert_number(array.get(1).unwrap(), 2.0);
        assert_error(array.get(2).unwrap(), ScalarError::NotAvailable);
    }
}

#[test]
fn matrix_nested_lazy_work_obeys_one_aggregate_step_limit() {
    let resolver = FixtureResolver::new();
    let simple = parse("=IF({TRUE();TRUE()};1;0)");
    let one_cell = parse("=IF({TRUE()};1+2+3+4+5+6+7+8+9+10;0)");
    let two_cells = parse("=IF({TRUE();TRUE()};1+2+3+4+5+6+7+8+9+10;0)");

    // Search only the small local limit dimension, using a fresh high-budget
    // execution context for every attempt.  At one limit the setup and one
    // expensive branch fit, while the same branch repeated for two selected
    // cells exceeds the complete evaluation's aggregate local Work allowance.
    // The simple two-cell control ensures the witness is not just an outer
    // matrix-shape admission failure.
    let mut witness = None;
    for max_steps in 0..=2_048_u64 {
        let evaluate_with_limit = |expression: &Expression| {
            let (_budget, _cancellation, execution) =
                make_execution("ods-formula-value-aggregate-work");
            evaluate_at(
                expression,
                &resolver,
                &execution,
                0,
                0,
                Mode::Matrix,
                &Limits::default().with_max_steps(max_steps),
            )
            .map(|_| ())
        };
        let simple_result = evaluate_with_limit(&simple);
        let one_result = evaluate_with_limit(&one_cell);
        let two_result = evaluate_with_limit(&two_cells);
        if simple_result.is_ok()
            && one_result.is_ok()
            && matches!(
                two_result,
                Err(EvaluationFailure::ResourceLimit(limit))
                    if limit.resource == Resource::Work
            )
        {
            witness = Some(max_steps);
            break;
        }
    }
    assert!(
        witness.is_some(),
        "no local Work limit distinguished one selected branch from two"
    );
}

#[test]
fn growing_lazy_branches_read_only_the_selected_cells() {
    const WIDTH: usize = 32;
    let mut resolver = FixtureResolver::new();
    resolver.sheets[0].1 = SheetExtent::new(8, WIDTH);

    let mut conditions = String::new();
    let mut expected_reads = Vec::with_capacity(WIDTH);
    for column in 0..WIDTH {
        if column != 0 {
            conditions.push(';');
        }
        conditions.push_str(if column % 2 == 0 { "TRUE()" } else { "FALSE()" });
        if column % 2 == 0 {
            resolver.set(
                "Main",
                0,
                column,
                FixtureCell::Number(1_000.0 + column as f64),
            );
            resolver.set(
                "Main",
                1,
                column,
                FixtureCell::Error(ScalarError::DivisionByZero),
            );
            expected_reads.push(("Main".to_owned(), 0, column));
        } else {
            resolver.set(
                "Main",
                0,
                column,
                FixtureCell::Error(ScalarError::DivisionByZero),
            );
            resolver.set(
                "Main",
                1,
                column,
                FixtureCell::Number(2_000.0 + column as f64),
            );
            expected_reads.push(("Main".to_owned(), 1, column));
        }
    }
    let end_label = column_label(WIDTH - 1);
    let source = format!("=IF({{{conditions}}};[.A1:.{end_label}1];[.A2:.{end_label}2])");
    let expression = parse(&source);
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-growing-lazy");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("growing lazy matrix should evaluate: {error}"));
    let array = result.as_array().expect("growing IF result array");
    assert_eq!((array.shape().rows(), array.shape().columns()), (1, WIDTH));
    for column in 0..WIDTH {
        let expected = if column % 2 == 0 {
            1_000.0 + column as f64
        } else {
            2_000.0 + column as f64
        };
        assert_number(
            array
                .get(column)
                .unwrap_or_else(|| panic!("missing selected cell {column}")),
            expected,
        );
    }
    assert_eq!(resolver.reads(), WIDTH);
    assert_eq!(resolver.read_order(), expected_reads);
}

#[test]
fn growing_lazy_aggregate_branch_is_evaluated_once_for_selected_cells() {
    const CONDITION_WIDTH: usize = 32;
    const RANGE_WIDTH: usize = 16;
    let mut resolver = FixtureResolver::new();
    resolver.sheets[0].1 = SheetExtent::new(8, CONDITION_WIDTH);
    for column in 0..RANGE_WIDTH {
        resolver.set("Main", 1, column, FixtureCell::Number(1.0));
    }

    let conditions = (0..CONDITION_WIDTH)
        .map(|column| if column % 2 == 0 { "TRUE()" } else { "FALSE()" })
        .collect::<Vec<_>>()
        .join(";");
    let source = format!("=IF({{{conditions}}};AND([.$A$2:.$P$2]);0)");
    let expression = parse(&source);
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-growing-aggregate-branch");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("growing aggregate branch should evaluate: {error}"));
    let array = result.as_array().expect("aggregate branch result array");
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (1, CONDITION_WIDTH)
    );
    for index in 0..CONDITION_WIDTH {
        if index % 2 == 0 {
            assert!(matches!(array.get(index), Some(Value::Logical(true))));
        } else {
            assert_number(array.get(index).unwrap(), 0.0);
        }
    }

    // There are sixteen selected output positions, but the selected AND
    // range must be read once and reused.  Re-running its sixteen-cell scan
    // for each selected position would produce 256 resolver reads.
    assert_eq!(resolver.reads(), RANGE_WIDTH);
    assert_eq!(
        resolver.read_order(),
        (0..RANGE_WIDTH)
            .map(|column| ("Main".to_owned(), 1, column))
            .collect::<Vec<_>>()
    );

    let mut nested_resolver = FixtureResolver::new();
    nested_resolver.sheets[0].1 = SheetExtent::new(8, CONDITION_WIDTH);
    for column in 0..RANGE_WIDTH {
        nested_resolver.set("Main", 1, column, FixtureCell::Number(1.0));
    }
    let conditions = (0..CONDITION_WIDTH)
        .map(|column| if column % 2 == 0 { "TRUE()" } else { "FALSE()" })
        .collect::<Vec<_>>()
        .join(";");
    let numbers = (1..=CONDITION_WIDTH)
        .map(|value| value.to_string())
        .collect::<Vec<_>>()
        .join(";");
    let source = format!("=IF({{{conditions}}};{{{numbers}}}+AND([.$A$2:.$P$2]);0)");
    let expression = parse(&source);
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-growing-nested-aggregate");
    let result = evaluate_at(
        &expression,
        &nested_resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .unwrap_or_else(|error| panic!("nested aggregate branch should evaluate: {error}"));
    let array = result
        .as_array()
        .expect("nested aggregate branch result array");
    assert_eq!(
        (array.shape().rows(), array.shape().columns()),
        (1, CONDITION_WIDTH)
    );
    for index in 0..CONDITION_WIDTH {
        let expected = if index % 2 == 0 {
            (index + 2) as f64
        } else {
            0.0
        };
        assert_number(
            array
                .get(index)
                .unwrap_or_else(|| panic!("missing nested selected cell {index}")),
            expected,
        );
    }
    // The array-producing branch contains the same selected AND range.  Its
    // sixteen cells must be materialized once for all sixteen true outputs;
    // evaluating that nested branch independently would read 256 cells.
    assert_eq!(nested_resolver.reads(), RANGE_WIDTH);
    assert_eq!(
        nested_resolver.read_order(),
        (0..RANGE_WIDTH)
            .map(|column| ("Main".to_owned(), 1, column))
            .collect::<Vec<_>>()
    );
}

#[test]
fn absolute_and_relative_matrix_reference_branches_project_each_cell() {
    for source in [
        "=IF({TRUE();TRUE()};[.A1:.B1];0)",
        "=IF({TRUE();TRUE()};[.$A$1:.$B$1];0)",
    ] {
        let mut resolver = FixtureResolver::new();
        resolver.set("Main", 0, 0, FixtureCell::Number(10.0));
        resolver.set("Main", 0, 1, FixtureCell::Number(11.0));
        let (_budget, _cancellation, execution) =
            make_execution("ods-formula-value-reference-branch-projection");
        let expression = parse(source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Matrix,
            &Limits::default(),
        )
        .unwrap_or_else(|error| panic!("{source:?} should project its reference branch: {error}"));
        let array = result.as_array().expect("reference branch result array");
        assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
        assert_array_numbers(array, &[10.0, 11.0]);
        assert_eq!(resolver.reads(), 2);
        assert_eq!(
            resolver.read_order(),
            vec![("Main".to_owned(), 0, 0), ("Main".to_owned(), 0, 1)]
        );
    }
}

#[test]
fn iferror_catches_missing_sheet_values_but_not_resolver_failures() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-iferror-reference");
    let expression = parse("=IFERROR([Missing.A1];42)");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect("missing sheet should be a catchable #REF! value");
    assert_number(result.value(), 42.0);
    assert_eq!(resolver.reads(), 0);

    let mut resolver = FixtureResolver::new();
    resolver.fail_after(0);
    let expression = parse("=IFERROR([.A1];42)");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("operational resolver failures must not be caught by IFERROR");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Reference)
    ));
    assert_eq!(resolver.reads(), 1);
}

#[test]
fn source_version_changes_refuse_a_borrowed_evaluation() {
    let mut resolver = FixtureResolver::new();
    resolver.source_changes(
        SourceVersion::new(0x4f44_5301, 0),
        SourceVersion::new(0x4f44_5301, 1),
    );
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-source-version");
    let expression = parse("=42");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("a changed resolver source must invalidate borrowed output");
    assert!(matches!(
        error,
        EvaluationFailure::SourceChanged { expected, observed }
            if expected == SourceVersion::new(0x4f44_5301, 0)
                && observed == SourceVersion::new(0x4f44_5301, 1)
    ));
}

#[test]
fn reference_list_kind_survives_single_intersection_and_range_collapses_it() {
    let resolver = FixtureResolver::new();
    let (_budget, _cancellation, execution) =
        make_execution("ods-formula-value-reference-list-kind");
    let limits = Limits::default();

    let expression = parse("=([.A1]~[.B1])![.A1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("one surviving reference-list record should remain first class");
    let Value::ReferenceList(list) = result.value() else {
        panic!("a one-entry reference list must not be inferred as a single reference")
    };
    assert_eq!(list.len(), 1);
    let reference = list.get(0).expect("surviving reference-list entry");
    assert_eq!(reference.areas()[0].starts(), [0, 0, 0]);
    assert_eq!(reference.areas()[0].extent(), [1, 1, 1]);

    let expression = parse("=NOT(([.A1]~[.B1])![.A1])");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("reference-list logical conversion should be a formula value");
    assert_error(result.value(), ScalarError::Value);

    let expression = parse("=(([.A1]~[.B1])![.A1]):[.B1]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &limits,
    )
    .expect("range over a one-entry list should produce a bounding reference");
    let Value::Reference(reference) = result.value() else {
        panic!("range application should collapse to one bounding reference")
    };
    assert_eq!(reference.areas().len(), 1);
    assert_eq!(reference.areas()[0].starts(), [0, 0, 0]);
    assert_eq!(reference.areas()[0].extent(), [1, 1, 2]);
}

#[test]
fn lazy_branches_do_not_resolve_unused_references_and_cancellation_is_atomic() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Number(99.0));
    let (_budget, _cancellation, execution) = make_execution("ods-formula-value-lazy");
    let limits = Limits::default();

    for source in ["=IF(TRUE();42;[.A1])", "=IF(FALSE();[.A1];42)"] {
        let expression = parse(source);
        let result = evaluate_at(
            &expression,
            &resolver,
            &execution,
            0,
            0,
            Mode::Scalar,
            &limits,
        )
        .unwrap_or_else(|error| panic!("{source:?} should skip its reference: {error}"));
        assert_number(result.value(), 42.0);
    }
    assert_eq!(resolver.reads(), 0, "unused lazy branch was resolved");

    let expression = parse("=IF(FALSE();[.A1];42)");
    let (_budget, cancellation, execution) = make_execution("ods-formula-value-cancel");
    cancellation.cancel();
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Scalar,
        &limits,
    )
    .expect_err("pre-cancelled value evaluation should stop before resolver access");
    assert!(matches!(error, EvaluationFailure::Cancelled));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn reference_limits_and_failed_reads_leave_budget_and_partial_results_clean() {
    let mut resolver = FixtureResolver::new();
    for row in 0..4 {
        for column in 0..4 {
            resolver.set(
                "Main",
                row,
                column,
                FixtureCell::Number((row * 4 + column) as f64),
            );
        }
    }
    let (budget, _cancellation, execution) = make_execution("ods-formula-value-limits");
    let expression = parse("=0+[.A1:.D4]");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default().with_max_reference_cells(8),
    )
    .expect_err("reference cell limit should refuse before partial adoption");
    assert!(matches!(
        error,
        EvaluationFailure::ResourceLimit(limit) if limit.resource == Resource::Objects
    ));
    assert_eq!(budget.used(Resource::Memory), 0);

    let expression = parse("=0+[.A1:.D4]");
    let result = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default()
            .with_max_reference_cells(16)
            .with_max_array_cells(16),
    )
    .expect("exact reference-cell capacity should be admitted");
    assert_eq!(result.as_array().unwrap().len(), 16);
    drop(result);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn resolver_failure_does_not_become_a_partial_array_or_leak_cell_storage() {
    let mut resolver = FixtureResolver::new();
    for row in 0..4 {
        for column in 0..4 {
            resolver.set(
                "Main",
                row,
                column,
                FixtureCell::Number((row * 4 + column) as f64),
            );
        }
    }
    resolver.fail_after(3);
    let (budget, _cancellation, execution) = make_execution("ods-formula-value-read-failure");
    let expression = parse("=0+[.A1:.D4]");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("resolver failure must remain an evaluator failure");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Reference)
    ));
    assert!(resolver.reads() >= 3);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn unsupported_resolver_cell_is_a_typed_capability_failure() {
    let mut resolver = FixtureResolver::new();
    resolver.set("Main", 0, 0, FixtureCell::Unsupported);
    let (budget, _cancellation, execution) = make_execution("ods-formula-value-unsupported-cell");
    let expression = parse("=0+[.A1]");
    let error = evaluate_at(
        &expression,
        &resolver,
        &execution,
        0,
        0,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("unsupported cell types must not be silently treated as empty");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::CellValue)
    ));
    assert_eq!(resolver.reads(), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}
