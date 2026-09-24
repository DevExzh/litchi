//! Semantic coverage for the OpenFormula 1.4 lookup family (§6.14).
//!
//! The fixture keeps references and cell values separate.  Functions that
//! return a reference (`INDEX`, `OFFSET`, and `INDIRECT`) are therefore tested
//! in matrix mode as first-class descriptors before a separate `SUM` consumer
//! is used to prove that a cell read happens only when a caller asks for one.
//! The lookup tables contain duplicate numeric keys so exact and approximate
//! search order remain observable.

use std::{
    cell::{Cell, RefCell},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::value::{
        self, CellRead, Context, Evaluated, Limits, Mode, OwnedValueView, Position, Resolver,
        SheetExtent, Value,
    },
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, UnsupportedKind,
        evaluate_scalar,
    },
    expression::Expression,
};

#[derive(Debug, Clone)]
enum FixtureCell {
    Empty,
    Number(f64),
    Logical(bool),
    Text(String),
    Error(ScalarError),
}

#[derive(Debug)]
struct LookupResolver {
    sheets: Vec<String>,
    extent: SheetExtent,
    cells: Vec<FixtureCell>,
    reads: Cell<usize>,
    read_order: RefCell<Vec<(String, usize, usize)>>,
    metadata_calls: Cell<usize>,
    metadata_sheets: RefCell<Vec<String>>,
}

impl LookupResolver {
    fn standard() -> Self {
        let extent = SheetExtent::new(16, 16);
        let mut resolver = Self {
            sheets: ["Main", "Data", "Hidden", "Archive"]
                .into_iter()
                .map(str::to_owned)
                .collect(),
            extent,
            cells: vec![FixtureCell::Empty; extent.rows() * extent.columns()],
            reads: Cell::new(0),
            read_order: RefCell::new(Vec::new()),
            metadata_calls: Cell::new(0),
            metadata_sheets: RefCell::new(Vec::new()),
        };

        // A1:C4 is a vertical lookup table.  The duplicate key at rows two
        // and three distinguishes exact-first from approximate-last search.
        for (row, (key, label, amount)) in [
            (1.0, "one", 10.0),
            (2.0, "two-first", 20.0),
            (2.0, "two-last", 21.0),
            (4.0, "four", 40.0),
        ]
        .into_iter()
        .enumerate()
        {
            resolver.set(row, 0, FixtureCell::Number(key));
            resolver.set(row, 1, FixtureCell::Text(label.to_owned()));
            resolver.set(row, 2, FixtureCell::Number(amount));
        }

        // E1:H3 is the corresponding horizontal table.  The first row is
        // the search vector; the second and third rows are result vectors.
        for (column, key) in [1.0, 2.0, 2.0, 4.0].into_iter().enumerate() {
            resolver.set(0, 4 + column, FixtureCell::Number(key));
        }
        for (column, label) in ["one", "two-first", "two-last", "four"]
            .into_iter()
            .enumerate()
        {
            resolver.set(1, 4 + column, FixtureCell::Text(label.to_owned()));
        }
        for (column, value) in [10.0, 20.0, 21.0, 40.0].into_iter().enumerate() {
            resolver.set(2, 4 + column, FixtureCell::Number(value));
        }

        // D1:D6 exercises the comparison owner's full Unicode case folding,
        // no-normalization rule, and literal wildcard handling.  The lookup
        // tables above stay numeric so approximate duplicate behavior remains
        // independently observable.
        for (row, text) in ["Straße", "STRASSE", "é", "e\u{301}", "*", "literal"]
            .into_iter()
            .enumerate()
        {
            resolver.set(row, 3, FixtureCell::Text(text.to_owned()));
        }

        // A5 has an Empty result in column B, and A16 is used for LOOKUP's
        // near-edge short-reference extension cases. D7 is a visited formula
        // error that must retain its identity through a search.
        resolver.set(4, 0, FixtureCell::Number(5.0));
        resolver.set(15, 0, FixtureCell::Number(111.0));
        resolver.set(6, 3, FixtureCell::Error(ScalarError::NotAvailable));
        resolver.set(7, 3, FixtureCell::Logical(false));
        resolver.set(8, 3, FixtureCell::Logical(true));
        resolver.set(4, 2, FixtureCell::Error(ScalarError::NotAvailable));
        resolver.set(5, 0, FixtureCell::Error(ScalarError::NotAvailable));

        // The explicit-sheet INDIRECT case is supplied by the Data-sheet
        // branch in `read_cell`; keeping Main.C3 in the lookup table avoids
        // changing the duplicate-key fixture.
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: FixtureCell) {
        let index = row
            .checked_mul(self.extent.columns())
            .and_then(|index| index.checked_add(column))
            .expect("fixture coordinate fits");
        self.cells[index] = value;
    }

    fn reads(&self) -> usize {
        self.reads.get()
    }

    fn clear_observations(&self) {
        self.reads.set(0);
        self.read_order.borrow_mut().clear();
        self.metadata_calls.set(0);
        self.metadata_sheets.borrow_mut().clear();
    }

    fn metadata_sheets(&self) -> Vec<String> {
        self.metadata_sheets.borrow().clone()
    }

    fn read_order(&self) -> Vec<(String, usize, usize)> {
        self.read_order.borrow().clone()
    }
}

impl Resolver for LookupResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        self.metadata_sheets.borrow_mut().push(sheet.to_owned());
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
        self.read_order
            .borrow_mut()
            .push((sheet.to_owned(), row, column));
        let Some(sheet_index) = self.sheets.iter().position(|name| name == sheet) else {
            return Ok(CellRead::Error(ScalarError::Reference));
        };
        if row >= self.extent.rows() || column >= self.extent.columns() {
            return Ok(CellRead::Error(ScalarError::Reference));
        }

        // Keep non-Main sheets empty except for Data.C3, which is used by the
        // explicit-sheet INDIRECT descriptor/consumer pair.
        if sheet_index != 0 && !(sheet == "Data" && row == 2 && column == 2) {
            return Ok(CellRead::Empty);
        }
        if sheet == "Data" && row == 2 && column == 2 {
            return Ok(CellRead::Number(303.0));
        }
        Ok(match &self.cells[row * self.extent.columns() + column] {
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
        self.metadata_calls
            .set(self.metadata_calls.get().saturating_add(1));
        self.metadata_sheets.borrow_mut().push(sheet.to_owned());
        Ok(self.sheets.iter().position(|name| name == sheet))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
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
    Reference {
        starts: [usize; 3],
        ends: [usize; 3],
        has_owner: bool,
    },
    ReferenceList(usize),
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

fn observe(result: &Evaluated<'_>) -> Observed {
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
        Value::Reference(reference) => Observed::Reference {
            starts: reference.areas()[0].starts(),
            ends: reference.areas()[0].ends(),
            has_owner: reference.reference().is_some(),
        },
        Value::ReferenceList(list) => Observed::ReferenceList(list.len()),
        Value::Complex(_) => Observed::Other,
        _ => Observed::Other,
    }
}

fn value(
    source: &str,
    resolver: &LookupResolver,
    execution: &ExecutionContext,
    position: Position<'_>,
    mode: Mode,
    limits: &Limits,
) -> Result<Observed, EvaluationFailure> {
    let expression = parse(source);
    let context = Context::new(execution, position).with_mode(mode);
    let result = value::evaluate(&expression, resolver, &context, limits)?;
    Ok(observe(&result))
}

fn assert_value(
    source: &str,
    expected: Observed,
    resolver: &LookupResolver,
    execution: &ExecutionContext,
    position: Position<'_>,
    mode: Mode,
) {
    assert_eq!(
        value(
            source,
            resolver,
            execution,
            position,
            mode,
            &Limits::default()
        )
        .unwrap_or_else(|error| panic!("{source} in {mode:?}: {error}")),
        expected,
        "{source} in {mode:?}"
    );
}

#[test]
fn address_formats_all_absolute_modes_and_r1c1_without_origin_shift() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-address");
    let position = Position::new("Main", 9, 9);

    for (source, expected) in [
        ("=ADDRESS(1;1)", "$A$1"),
        ("=ADDRESS(4;5;1;TRUE())", "$E$4"),
        ("=ADDRESS(4;5;2;TRUE())", "E$4"),
        ("=ADDRESS(4;5;3;TRUE())", "$E4"),
        ("=ADDRESS(4;5;4;TRUE())", "E4"),
        ("=ADDRESS(4;5;1;FALSE())", "R4C5"),
        ("=ADDRESS(4;5;2;FALSE())", "R4C[5]"),
        ("=ADDRESS(4;5;3;FALSE())", "R[4]C5"),
        ("=ADDRESS(4;5;4;FALSE())", "R[4]C[5]"),
        (r#"=ADDRESS(1;1;1;TRUE();"O'Brien")"#, "'O''Brien'.$A$1"),
        ("=ADDRESS(1;1;1;TRUE();17)", "17.$A$1"),
        ("=ADDRESS(1;1;1;TRUE();TRUE())", "TRUE.$A$1"),
        (r#"=ADDRESS(1;1;1;TRUE();"Da"&"ta")"#, "Data.$A$1"),
        (r#"=ADDRESS("2";"3")"#, "$C$2"),
        ("=ADDRESS(2.9;3.9)", "$C$2"),
        ("=ADDRESS(1;1;1.9)", "$A$1"),
    ] {
        assert_value(
            source,
            Observed::Text(expected.to_owned()),
            &resolver,
            &execution,
            position,
            Mode::Scalar,
        );
    }

    assert_value(
        "=ADDRESS({1|2};1;4;TRUE())",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![
                CellObserved::Text("A1".to_owned()),
                CellObserved::Text("A2".to_owned()),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
}

#[test]
fn lookup_domain_and_arity_refusals_are_formula_values() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-malformed");
    let position = Position::new("Main", 0, 0);

    for source in [
        "=ADDRESS(0;1)",
        "=ADDRESS(1;1;5)",
        "=ADDRESS(1e100;1)",
        "=CHOOSE(0;1)",
        "=CHOOSE(1)",
        "=MATCH(2;[.A1:.A4];2)",
        "=OFFSET([.A1];;0)",
        "=OFFSET([.A1];0;)",
        "=OFFSET([.A1:.B2];0;0;0;1)",
        "=OFFSET([.A1:.B2];0;0;1;0)",
    ] {
        assert_value(
            source,
            Observed::Error(ScalarError::Value),
            &resolver,
            &execution,
            position,
            Mode::Scalar,
        );
        assert_eq!(resolver.reads(), 0, "{source} has no cell-consuming path");
        resolver.clear_observations();
    }
    assert_value(
        "=OFFSET([.A1:.B2];-1;0)",
        Observed::Error(ScalarError::Reference),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
}

#[test]
fn choose_evaluates_only_the_selected_value() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-choose-lazy");

    assert_value(
        "=CHOOSE(2;10;20;30)",
        Observed::Number(20.0),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Scalar,
    );
    assert_value(
        "=CHOOSE(1.9;10;20)",
        Observed::Number(10.0),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Scalar,
    );

    resolver.clear_observations();
    assert_value(
        "=CHOOSE(1;42;[Missing.A1])",
        Observed::Number(42.0),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 0, "unselected CHOOSE reference is unread");
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing"),
        "unselected CHOOSE reference is not probed"
    );

    resolver.clear_observations();
    assert_value(
        "=CHOOSE(2;42;[.A1])",
        Observed::Reference {
            starts: [0, 0, 0],
            ends: [1, 1, 1],
            has_owner: true,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "selected reference remains first-class"
    );
}

#[test]
fn choose_lifts_matrix_indices_and_preserves_selected_identity() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-choose-matrix");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=CHOOSE({1|1};{10;11};[Missing.A1])",
        Observed::Array {
            rows: 2,
            columns: 2,
            cells: vec![
                CellObserved::Number(10.0),
                CellObserved::Number(11.0),
                CellObserved::Number(10.0),
                CellObserved::Number(11.0),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing"),
        "unselected matrix CHOOSE branches do not participate in shape probing"
    );

    resolver.clear_observations();
    assert_value(
        "=CHOOSE({1|1};[.C1:.C2];[Missing.A1])",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(10.0), CellObserved::Number(20.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 2);
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing"),
        "unselected CHOOSE reference branches do not participate in shape probing"
    );

    resolver.clear_observations();
    assert_value(
        "=IF({TRUE()|TRUE()};CHOOSE(IF({TRUE()|FALSE()};1;2);10;20);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(10.0), CellObserved::Number(20.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "a computed CHOOSE index keeps its complete projected shape"
    );

    resolver.clear_observations();
    assert_value(
        "=SUM(IF({TRUE()|TRUE()};CHOOSE(IF({TRUE()|TRUE()};1;1);[.C1:.C2];[Missing.A1]);0))",
        Observed::Number(30.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "only the selected CHOOSE reference branch is consumed"
    );
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing"),
        "the unselected computed-index branch is not probed"
    );

    resolver.clear_observations();
    assert_value(
        "=CHOOSE({1|3};10;20)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![
                CellObserved::Number(10.0),
                CellObserved::Error(ScalarError::Value),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );

    resolver.clear_observations();
    assert_value(
        "=CHOOSE(2;0;([.A1]~[.B1]))",
        Observed::ReferenceList(2),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=SUM(IF({TRUE()};CHOOSE(1;INDIRECT("C1:C2"));0))"#,
        Observed::Number(30.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "nested CHOOSE keeps the selected INDIRECT range geometry"
    );

    assert_value(
        "=CHOOSE(2;0;#N/A)",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
}

#[test]
fn choose_computed_index_reads_scalar_input_once() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-choose-computed-index-read");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=SUM(IF({TRUE()};CHOOSE(SUM([.A1]);10;20);0))",
        Observed::Number(10.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        1,
        "the computed CHOOSE index must not discard and reread its A1 input"
    );
    assert_eq!(
        resolver.read_order(),
        vec![("Main".to_owned(), 0, 0)],
        "CHOOSE index evaluation should consume A1 once"
    );

    resolver.clear_observations();
    assert_value(
        "=SUM(IF({TRUE()|TRUE()};CHOOSE(SUM([.A1]);10;20);0))",
        Observed::Number(20.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        1,
        "a two-position projected index still consumes A1 once"
    );
}

#[test]
fn choose_computed_index_respects_outer_false_mask() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-choose-computed-index-mask");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=IF({TRUE();FALSE()};CHOOSE(IF({TRUE();FALSE()};1;[Missing.A1]);10;20);0)",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![CellObserved::Number(10.0), CellObserved::Number(0.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "an outer false mask must avoid consuming the computed CHOOSE index"
    );
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing"),
        "an outer false mask must avoid probing the computed index's Missing branch"
    );
}

#[test]
fn index_selects_array_cells_and_missing_row_or_column_returns_vectors() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-index-array");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=INDEX({1;2|3;4};2;1)",
        Observed::Number(3.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=INDEX({1;2|3;4};2;1)",
        Observed::Number(3.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_value(
        "=INDEX({1;2|3;4};;2)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(2.0), CellObserved::Number(4.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_value(
        "=INDEX({1;2|3;4};2;)",
        Observed::Array {
            rows: 1,
            columns: 2,
            cells: vec![CellObserved::Number(3.0), CellObserved::Number(4.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_value(
        "=INDEX({1;2|3;4};;)",
        Observed::Array {
            rows: 2,
            columns: 2,
            cells: vec![
                CellObserved::Number(1.0),
                CellObserved::Number(2.0),
                CellObserved::Number(3.0),
                CellObserved::Number(4.0),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_value(
        "=INDEX({1;2|3;4};{2|1};1)",
        Observed::Number(3.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    for (source, expected) in [
        ("=INDEX({1};1e100;1)", ScalarError::Reference),
        ("=INDEX({1;2|3;4};3;1)", ScalarError::Reference),
        ("=INDEX({1;2|3;4};1;1;2)", ScalarError::Reference),
        ("=INDEX([.A1:.C4];5;1)", ScalarError::Reference),
        ("=INDEX([.A1:.C4];1;4)", ScalarError::Reference),
        ("=INDEX(([.A1]~[.B1]);1;1;3)", ScalarError::Reference),
        ("=INDEX([.A1:.C4];1;1;0)", ScalarError::Reference),
        ("=INDEX([.A1:.C4];1;1;-1)", ScalarError::Reference),
        ("=INDEX([.A1:.C4];1e100;1)", ScalarError::Reference),
        ("=INDEX([.A1:.C4];-1;1)", ScalarError::Value),
        ("=INDEX([.A1:.C4];-1e100;1)", ScalarError::Value),
        ("=INDEX([.A1:.C4];1;1;1e100)", ScalarError::Reference),
        ("=INDEX([.A1:.C4];1;1;-1e100)", ScalarError::Reference),
        ("=INDEX([.A1:.C4];#N/A;1)", ScalarError::NotAvailable),
        ("=INDEX([.A1:.C4];#NUM!;1)", ScalarError::Number),
        ("=INDEX(#N/A;1;1)", ScalarError::NotAvailable),
    ] {
        assert_value(
            source,
            Observed::Error(expected),
            &resolver,
            &execution,
            position,
            Mode::Matrix,
        );
        assert_eq!(resolver.reads(), 0, "{source} is a descriptor-only refusal");
        resolver.clear_observations();
    }
}

#[test]
fn index_and_offset_return_derived_reference_descriptors_without_reads() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-derived-references");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=INDEX([.A1:.C4];2;3)",
        Observed::Reference {
            starts: [0, 1, 2],
            ends: [1, 2, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=INDEX(([.A1]~[.B1]~[.A1]);1;1;2)",
        Observed::Reference {
            starts: [0, 0, 1],
            ends: [1, 1, 2],
            has_owner: false,
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=INDEX([Main.A1:Archive.C3];2;3)",
        Observed::Reference {
            starts: [0, 1, 2],
            ends: [4, 2, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=OFFSET([.A1:.B2];1;1)",
        Observed::Reference {
            starts: [0, 1, 1],
            ends: [1, 3, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=OFFSET([.A1:.B2];0;0)",
        Observed::Reference {
            starts: [0, 0, 0],
            ends: [1, 2, 2],
            has_owner: false,
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=OFFSET([.A1:.B2];0;0;;1)",
        Observed::Reference {
            starts: [0, 0, 0],
            ends: [1, 2, 1],
            has_owner: false,
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=OFFSET([.A1:.B2];0;0;[.E4])",
        Observed::Error(ScalarError::Value),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        1,
        "an explicit Empty dimension is read as zero"
    );

    resolver.clear_observations();
    assert_value(
        "=OFFSET([Main.A1:Archive.B2];1;1)",
        Observed::Reference {
            starts: [0, 1, 1],
            ends: [4, 3, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=INDEX([.A1:.C4];{2|1};3)",
        Observed::Reference {
            starts: [0, 1, 2],
            ends: [1, 2, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=OFFSET([.A1:.B2];{1|0};{1|0})",
        Observed::Reference {
            starts: [0, 1, 1],
            ends: [1, 3, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=SUM(INDEX([.A1:.C4];2;3))",
        Observed::Number(20.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 1, "consumer reads selected INDEX cell");
    assert_eq!(resolver.read_order(), vec![("Main".to_owned(), 1, 2)]);
}

#[test]
fn offset_preserves_formula_errors_and_rejects_source_references_as_typed_capability() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-offset-errors");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=OFFSET(#N/A;0;0)",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    for (source, expected) in [
        ("=OFFSET([.A1];1e100;0)", ScalarError::Reference),
        ("=OFFSET([.A1];-1e100;0)", ScalarError::Reference),
        ("=OFFSET([.A1];0;0;1e100;1)", ScalarError::Reference),
        ("=OFFSET([.A1];0;0;1;1e100)", ScalarError::Reference),
        ("=OFFSET([.A1];#NUM!;0)", ScalarError::Number),
    ] {
        resolver.clear_observations();
        assert_value(
            source,
            Observed::Error(expected),
            &resolver,
            &execution,
            position,
            Mode::Matrix,
        );
        assert_eq!(resolver.reads(), 0, "{source} is a descriptor-only refusal");
    }

    let error = value(
        "=OFFSET(['file:///book.ods'#.A1];0;0)",
        &resolver,
        &execution,
        position,
        Mode::Matrix,
        &Limits::default(),
    )
    .expect_err("OFFSET cannot resolve a source-qualified reference");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Reference)
    ));
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn selected_derived_references_keep_matrix_shape_without_unsliced_branch_probes() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-derived-lazy-shape");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=SUM(IF({TRUE()};INDEX([.A1:.C4];2;3);[Missing.A1:.Z100]))",
        Observed::Number(20.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 1);
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing"),
        "an unselected branch must not be inspected during shape planning"
    );

    resolver.clear_observations();
    assert_value(
        "=SUM(CHOOSE({1};OFFSET([.A1:.C4];1;2;1;1);[Missing.A1:.Z100]))",
        Observed::Number(20.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 1);
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing"),
        "CHOOSE shape planning must retain only selected derived geometry"
    );
}

#[test]
fn projected_if_keeps_dynamic_reference_shapes_and_reads_only_selected_cells() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-projected-derived-shape");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=SUM(IF({TRUE()};INDEX([.A1:.C4];0;3);[Missing.A1:.Z100]))",
        Observed::Number(91.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        4,
        "INDEX's selected column geometry is read only after projection"
    );
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing"),
        "shape planning must not inspect an unselected IF branch"
    );

    resolver.clear_observations();
    assert_value(
        "=SUM(IF({TRUE()};OFFSET([.A1:.C4];1;2;2;1);[Missing.A1:.Z100]))",
        Observed::Number(41.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "OFFSET's selected two-cell geometry is planned without source probing"
    );
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing")
    );

    resolver.clear_observations();
    assert_value(
        r#"=SUM(IF({TRUE()};INDIRECT("C1:C2");[Missing.A1:.Z100]))"#,
        Observed::Number(30.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "INDIRECT's selected range is consumed without shape-probe reads"
    );
    assert!(
        !resolver
            .metadata_sheets()
            .iter()
            .any(|name| name == "Missing")
    );
}

#[test]
fn projected_indirect_growth_does_not_replay_its_text_selector() {
    let mut resolver = LookupResolver::standard();
    // H1 is outside the lookup tables for this case.  Its text selects a
    // two-cell range whose shape is discovered only after the projected IF
    // branch is retained.
    resolver.set(0, 7, FixtureCell::Text("C1:C2".to_owned()));
    let (_budget, _cancellation, execution) = execution("lookup-projected-indirect-growth");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=SUM(IF({TRUE()};INDIRECT([.H1]);0))",
        Observed::Number(30.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );

    let order = resolver.read_order();
    let selector_reads = order
        .iter()
        .filter(|(sheet, row, column)| sheet == "Main" && *row == 0 && *column == 7)
        .count();
    let result_reads = order
        .iter()
        .filter(|(sheet, row, column)| sheet == "Main" && *column == 2 && *row <= 1)
        .count();
    assert!(
        selector_reads <= 2,
        "the H1 selector must not be replayed during shape refinement: {order:?}"
    );
    assert_eq!(
        result_reads, 2,
        "both selected C1:C2 cells are consumed exactly once: {order:?}"
    );
    assert!(
        order.len() <= 4,
        "projected INDIRECT should need at most two selector and two result reads: {order:?}"
    );
    assert!(
        order.iter().all(|(sheet, row, column)| {
            sheet == "Main" && ((*column == 2 && *row <= 1) || (*row == 0 && *column == 7))
        }),
        "shape refinement must not probe unrelated cells: {order:?}"
    );
}

#[test]
fn projected_indirect_width_growth_tracks_read_text_selector() {
    let mut resolver = LookupResolver::standard();
    // Start with a 2x1 IF demand and let the selected INDIRECT range
    // widen it to a genuine 2x2 result. D1:D2 are made numeric so every
    // selected target cell contributes an observable value.
    resolver.set(0, 7, FixtureCell::Text("C1:D2".to_owned()));
    resolver.set(0, 3, FixtureCell::Number(11.0));
    resolver.set(1, 3, FixtureCell::Number(12.0));
    let (_budget, _cancellation, execution) = execution("lookup-projected-indirect-width-growth");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=SUM(IF({TRUE()|TRUE()};INDIRECT([.H1]);0))",
        Observed::Number(53.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );

    let order = resolver.read_order();
    assert_eq!(
        order.len(),
        8,
        "width growth should read four selector coordinates and four target cells: {order:?}"
    );
    assert_eq!(
        order
            .iter()
            .filter(|(sheet, row, column)| { sheet == "Main" && *row == 0 && *column == 7 })
            .count(),
        4,
        "the read-text selector must be visited once at each demanded coordinate: {order:?}"
    );
    for (row, column) in [(0, 2), (0, 3), (1, 2), (1, 3)] {
        assert_eq!(
            order
                .iter()
                .filter(|(sheet, read_row, read_column)| {
                    sheet == "Main" && *read_row == row && *read_column == column
                })
                .count(),
            1,
            "selected C1:D2 target must be read once at ({row},{column}): {order:?}"
        );
    }
}

#[test]
fn projected_scalar_selectors_use_first_element_but_munit_stays_position_sensitive() {
    let mut resolver = LookupResolver::standard();
    resolver.set(0, 6, FixtureCell::Number(1.0));
    resolver.set(1, 6, FixtureCell::Number(2.0));
    let (_budget, _cancellation, execution) = execution("lookup-projected-selectors");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=SUM(IF({TRUE()|TRUE()};INDEX([.A1:.C4];{2|1};3);[Missing.A1:.Z100]))",
        Observed::Number(40.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "INDEX matrix selector uses element (0,0) for each output"
    );

    resolver.clear_observations();
    assert_value(
        "=SUM(IF({TRUE()|TRUE()};OFFSET([.A1];{1|0};{2|0});[Missing.A1:.Z100]))",
        Observed::Number(40.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "OFFSET matrix selectors use element (0,0) without shape probes"
    );

    resolver.clear_observations();
    assert_value(
        r#"=SUM(IF({TRUE()|TRUE()};INDIRECT({"C2"|"C1"});[Missing.A1:.Z100]))"#,
        Observed::Number(40.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "INDIRECT matrix text selector uses element (0,0)"
    );

    resolver.clear_observations();
    assert_value(
        "=SUM(IF({TRUE()|TRUE()};INDEX([.A1:.C4];SUM(MUNIT([.G1:.G2]));3);[Missing.A1:.Z100]))",
        Observed::Number(30.0),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        4,
        "computed MUNIT selectors read G1/G2 and select C1/C2 independently"
    );
}

#[test]
fn indirect_keeps_a1_and_explicit_r1c1_references_first_class() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-indirect");

    assert_value(
        r#"=INDIRECT("B2")"#,
        Observed::Reference {
            starts: [0, 1, 1],
            ends: [1, 2, 2],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=ISREF(INDIRECT("'file:///book.ods'#.A1"))"#,
        Observed::Logical(true),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "ISREF classifies a source-qualified INDIRECT descriptor without reads"
    );

    let error = value(
        r#"=SUM(INDIRECT("'file:///book.ods'#.A1"))"#,
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Scalar,
        &Limits::default(),
    )
    .expect_err("a source-qualified INDIRECT consumer needs an external provider");
    assert!(matches!(
        error,
        EvaluationFailure::Unsupported(UnsupportedKind::Reference)
    ));
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=INDIRECT("Data.C3")"#,
        Observed::Reference {
            starts: [1, 2, 2],
            ends: [2, 3, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=INDIRECT("R[1]C[1]";FALSE())"#,
        Observed::Reference {
            starts: [0, 5, 6],
            ends: [1, 6, 7],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 4, 5),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=INDIRECT("R[-1]C";FALSE())"#,
        Observed::Reference {
            starts: [0, 3, 5],
            ends: [1, 4, 6],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 4, 5),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=INDIRECT("A1:C2")"#,
        Observed::Reference {
            starts: [0, 0, 0],
            ends: [1, 2, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=INDIRECT("A:C")"#,
        Observed::Reference {
            starts: [0, 0, 0],
            ends: [1, 16, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=INDIRECT("1:4")"#,
        Observed::Reference {
            starts: [0, 0, 0],
            ends: [1, 4, 16],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=INDIRECT("Main.A1:Archive.C3")"#,
        Observed::Reference {
            starts: [0, 0, 0],
            ends: [4, 3, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=INDIRECT("Data!C3")"#,
        Observed::Reference {
            starts: [1, 2, 2],
            ends: [2, 3, 3],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=INDIRECT({"A1"|"A2"})"#,
        Observed::Reference {
            starts: [0, 0, 0],
            ends: [1, 1, 1],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        r#"=SUM(INDIRECT({"A1"|"A2"}))"#,
        Observed::Number(1.0),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 1);

    resolver.clear_observations();
    assert_value(
        r#"=ISREF(INDIRECT("bad text"))"#,
        Observed::Logical(false),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 0);

    resolver.clear_observations();
    assert_value(
        "=INDIRECT(1)",
        Observed::Error(ScalarError::Reference),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0, "numeric text conversion is read-free");

    resolver.clear_observations();
    assert_value(
        "=INDIRECT(TRUE())",
        Observed::Error(ScalarError::Reference),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0, "logical text conversion is read-free");

    resolver.clear_observations();
    assert_value(
        "=INDIRECT([.E4])",
        Observed::Error(ScalarError::Reference),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        1,
        "an actual Empty text argument is read once before parsing"
    );

    resolver.clear_observations();
    assert_value(
        "=INDIRECT(\"R1C1\";[.E4])",
        Observed::Reference {
            starts: [0, 0, 0],
            ends: [1, 1, 1],
            has_owner: false,
        },
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        1,
        "an actual Empty A1 flag is read once and selects R1C1"
    );

    resolver.clear_observations();
    assert_value(
        r#"=SUM(INDIRECT("Data.C3"))"#,
        Observed::Number(303.0),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 1, "INDIRECT consumer reads its target");
}

#[test]
fn scalar_indirect_relative_reference_reports_capability_refusal() {
    let (_budget, _cancellation, execution) = execution("lookup-indirect-scalar");
    let expression = parse(r#"=INDIRECT("R[-1]C";FALSE())"#);
    let error = evaluate_scalar(
        &expression,
        &EvaluationContext::new(&execution),
        &EvaluationLimits::default(),
    )
    .expect_err("a scalar evaluator cannot resolve a worksheet reference");
    assert!(
        matches!(
            error,
            EvaluationFailure::Unsupported(UnsupportedKind::Reference)
        ),
        "a syntactically valid relative reference must reach the typed capability boundary"
    );
}

#[test]
fn match_exact_returns_first_duplicate_and_approximate_returns_last() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-match");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=MATCH(2;[.A1:.A4];0)",
        Observed::Number(2.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=MATCH(3;[.A1:.A4];1)",
        Observed::Number(3.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
}

#[test]
fn search_any_keys_and_scalar_selectors_lift_by_matrix_shape() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-search-matrix");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=MATCH({1|4};[.A1:.A4];0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(4.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    resolver.clear_observations();
    assert_value(
        "=MATCH({1|2}+0;[.A1:.A4];0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    resolver.clear_observations();
    assert_value(
        "=IF({TRUE()|TRUE()};MATCH(MATCH({1|2};{1|2};0);{1|2};0);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "nested lifted lookup keys must not cross the resolver boundary"
    );
    assert_value(
        "=LOOKUP({1|4};[.A1:.A4];[.B1:.B4])",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![
                CellObserved::Text("one".to_owned()),
                CellObserved::Text("four".to_owned()),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_value(
        "=VLOOKUP({1|4};[.A1:.C4];2;FALSE())",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![
                CellObserved::Text("one".to_owned()),
                CellObserved::Text("four".to_owned()),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_value(
        "=MATCH(2;[.A1:.A4];{0|1})",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(2.0), CellObserved::Number(3.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_value(
        "=VLOOKUP(2;[.A1:.C4];2;{FALSE()|TRUE()})",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![
                CellObserved::Text("two-first".to_owned()),
                CellObserved::Text("two-last".to_owned()),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );

    // Unequal, well-formed selector shapes use the same rectangular
    // broadcasting contract as the other matrix functions.  A source cell
    // that has no coordinate in the larger output shape publishes #N/A for
    // that coordinate instead of becoming an evaluator invariant failure.
    assert_value(
        "=MATCH({1;2|3;4};[.A1:.A4];{0;1;0})",
        Observed::Array {
            rows: 2,
            columns: 3,
            cells: vec![
                CellObserved::Number(1.0),
                CellObserved::Number(3.0),
                CellObserved::Error(ScalarError::NotAvailable),
                CellObserved::Error(ScalarError::NotAvailable),
                CellObserved::Number(4.0),
                CellObserved::Error(ScalarError::NotAvailable),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
}

#[test]
fn projected_lookup_cache_reuses_direct_search_but_keeps_munit_position_sensitive() {
    let mut resolver = LookupResolver::standard();
    resolver.set(0, 6, FixtureCell::Number(1.0));
    resolver.set(1, 6, FixtureCell::Number(2.0));
    let (_budget, _cancellation, execution) = execution("lookup-projected-cache");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=IF({TRUE()|TRUE()};MATCH(2;[.A1:.A4];0);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(2.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "a direct lookup branch is evaluated once and reused across projections"
    );

    resolver.clear_observations();
    assert_value(
        "=IF({TRUE()|TRUE()};MATCH(SUM(MUNIT([.G1:.G2]));[.A1:.A4];0);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        5,
        "a computed MUNIT selector stays position-sensitive while each lookup searches"
    );

    resolver.clear_observations();
    assert_value(
        "=IF({TRUE()|TRUE()};MATCH(SUM(MUNIT([.G1:.G2]))+0;[.A1:.A4];0);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        5,
        "an arithmetic wrapper must preserve the position-sensitive MUNIT key"
    );

    resolver.clear_observations();
    assert_value(
        "=IF({TRUE()|TRUE()};MATCH(ABS(SUM(MUNIT([.G1:.G2])));[.A1:.A4];0);0)",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![CellObserved::Number(1.0), CellObserved::Number(2.0)],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        5,
        "ABS must preserve the position-sensitive MUNIT key"
    );
}

#[test]
fn search_comparison_is_type_preserving_casefolded_and_not_normalized() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-search-comparison");
    let position = Position::new("Main", 0, 0);

    assert_value(
        r#"=MATCH("strasse";[.D1:.D2];0)"#,
        Observed::Number(1.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        r#"=MATCH("é";[.D3:.D4];0)"#,
        Observed::Number(2.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        r#"=MATCH("é";[.D3:.D4];0)"#,
        Observed::Number(1.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        r#"=MATCH("*";[.D5:.D6];0)"#,
        Observed::Number(1.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        r#"=MATCH("2";[.A1:.A4];0)"#,
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        r#"=MATCH("3";[.A1:.A4];1)"#,
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=MATCH(FALSE();[.D8:.D9];0)",
        Observed::Number(1.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=MATCH(0;[.D8:.D9];0)",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
}

#[test]
fn search_distinguishes_missing_empty_and_visited_formula_errors() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-search-values");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=MATCH(2;[.A1:.A4])",
        Observed::Number(3.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=MATCH(3;[.A1:.A4];1.9)",
        Observed::Number(3.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=MATCH(2;[.A1:.A4];[.E4])",
        Observed::Number(2.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=MATCH(0;[.E4];0)",
        Observed::Number(1.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=HLOOKUP(2;[.E1:.H3];2)",
        Observed::Text("two-last".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=HLOOKUP(2;[.E1:.H3];2;[.E4])",
        Observed::Text("two-first".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        r#"=HLOOKUP(2;[.E1:.H3];2;"1")"#,
        Observed::Error(ScalarError::Value),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=VLOOKUP(5;[.A1:.C5];2;FALSE())",
        Observed::Empty,
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=VLOOKUP(2;[.A1:.C4];2.9;FALSE())",
        Observed::Text("two-first".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=VLOOKUP(5;[.A1:.C5];3;FALSE())",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=MATCH(1;[.D7];0)",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    resolver.clear_observations();
    assert_value(
        "=MATCH(1;[.A1:.A6];0)",
        Observed::Number(1.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        1,
        "an exact match may stop before an unvisited formula-error cell"
    );
    resolver.clear_observations();
    assert_value(
        "=MATCH(TRUE();[.D7:.D9];0)",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        3,
        "a formula error in the first search cell does not stop later exact probes"
    );
    resolver.clear_observations();
    assert_value(
        "=MATCH(([.A1]~[.A2]);[.A1:.A4];0)",
        Observed::Error(ScalarError::Value),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "ReferenceList keys are refused before searching"
    );
    resolver.clear_observations();
    assert_value(
        "=MATCH(#N/A;[.A1:.A4];0)",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "formula-error keys do not probe the search region"
    );
    resolver.clear_observations();
    assert_value(
        "=MATCH(1;#N/A;0)",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "formula-error data propagates before cell reads"
    );
    resolver.clear_observations();
    assert_value(
        "=MATCH(2;[.A1:.A4];#N/A)",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        0,
        "formula-error MatchType propagates before reads"
    );
}

#[test]
fn horizontal_and_vertical_lookup_exact_and_approximate_modes() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-hv");
    let position = Position::new("Main", 0, 0);

    for (source, expected) in [
        ("=HLOOKUP(2;[.E1:.H3];2;FALSE())", "two-first"),
        ("=HLOOKUP(3;[.E1:.H3];2;TRUE())", "two-last"),
        ("=VLOOKUP(2;[.A1:.C4];2;FALSE())", "two-first"),
        ("=VLOOKUP(3;[.A1:.C4];2;TRUE())", "two-last"),
    ] {
        assert_value(
            source,
            Observed::Text(expected.to_owned()),
            &resolver,
            &execution,
            position,
            Mode::Scalar,
        );
    }
}

#[test]
fn table_lookup_reads_search_prefix_and_selected_result_only() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-selected-read");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=VLOOKUP(2;[.A1:.C4];2;FALSE())",
        Observed::Text("two-first".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.read_order(),
        vec![
            ("Main".to_owned(), 0, 0),
            ("Main".to_owned(), 1, 0),
            ("Main".to_owned(), 1, 1),
        ]
    );

    resolver.clear_observations();
    assert_value(
        "=VLOOKUP(3;[.A1:.C4];2;TRUE())",
        Observed::Text("two-last".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert!(
        resolver.reads() <= 4,
        "binary approximate search and selected result must stay bounded"
    );
    assert_eq!(
        resolver.read_order().last(),
        Some(&("Main".to_owned(), 2, 1)),
        "the approximate lookup's final read is its selected result"
    );
    assert!(
        resolver
            .read_order()
            .iter()
            .take(resolver.read_order().len().saturating_sub(1))
            .all(|(_, _, column)| *column == 0),
        "unselected VLOOKUP result cells remain unread"
    );

    resolver.clear_observations();
    assert_value(
        "=HLOOKUP(2;[.E1:.H3];2;FALSE())",
        Observed::Text("two-first".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.read_order(),
        vec![
            ("Main".to_owned(), 0, 4),
            ("Main".to_owned(), 0, 5),
            ("Main".to_owned(), 1, 5),
        ]
    );
}

#[test]
fn lookup_uses_tall_and_wide_array_orientation_and_last_approximate_duplicate() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-orientation");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=LOOKUP(3;[.A1:.B4])",
        Observed::Text("two-last".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=LOOKUP(3;[.E1:.H2])",
        Observed::Text("two-last".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=LOOKUP(3;[.A1:.A4];[.B1:.B4])",
        Observed::Text("two-last".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        "=LOOKUP(3;[.A1:.B4];[.B1:.B4])",
        Observed::Text("two-last".to_owned()),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
}

#[test]
fn approximate_search_crosses_numeric_and_text_midpoints() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-mixed-approximate");
    let position = Position::new("Main", 0, 0);

    assert_value(
        r#"=MATCH("m";{1|2|3|4|"a"|"m"|"z"};1)"#,
        Observed::Number(6.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        r#"=MATCH("m";{"z"|"m"|"a"|4|3|2|1};-1)"#,
        Observed::Number(2.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
}

#[test]
fn lookup_extends_short_references_only_after_an_out_of_range_match() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-result-extension");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=LOOKUP(1;[.A1:.A4];[.A16])",
        Observed::Number(111.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.read_order().last(),
        Some(&("Main".to_owned(), 15, 0)),
        "an in-range short result reads its existing cell without hypothetical extension"
    );

    resolver.clear_observations();
    assert_value(
        "=LOOKUP(4;[.A1:.A4];[.A16])",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert!(
        !resolver
            .read_order()
            .iter()
            .any(|(_, row, column)| *row == 15 && *column == 0),
        "an out-of-bounds extension refuses before reading the edge result cell"
    );

    resolver.clear_observations();
    assert_value(
        "=LOOKUP(4;[.E1:.H1];{111})",
        Observed::Error(ScalarError::NotAvailable),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(
        resolver.reads(),
        2,
        "a short Array result returns #N/A after the search vector is scanned"
    );
}

#[test]
fn lookup_search_data_rejects_scalars_lists_and_three_dimensional_references() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-shape-refusals");
    let position = Position::new("Main", 0, 0);

    for source in [
        "=MATCH(2;7;0)",
        "=LOOKUP(2;7)",
        "=HLOOKUP(2;7;1;FALSE())",
        "=VLOOKUP(2;7;1;FALSE())",
        "=MATCH(2;([.A1]~[.A2]);0)",
        "=LOOKUP(2;([.A1]~[.A2]))",
        "=HLOOKUP(2;([.A1]~[.A2]);1;FALSE())",
        "=VLOOKUP(2;([.A1]~[.A2]);1;FALSE())",
        "=MATCH(2;[Main.A1:Archive.A1];0)",
        "=LOOKUP(2;[Main.A1:Archive.A1])",
        "=HLOOKUP(2;[Main.A1:Archive.A1];1;FALSE())",
        "=VLOOKUP(2;[Main.A1:Archive.A1];1;FALSE())",
    ] {
        resolver.clear_observations();
        assert_value(
            source,
            Observed::Error(ScalarError::Value),
            &resolver,
            &execution,
            position,
            Mode::Scalar,
        );
        assert_eq!(
            resolver.reads(),
            0,
            "{source} must refuse before cell reads"
        );
    }

    for source in ["=MATCH(2;[.A1:.B2];0)", "=LOOKUP(2;[.A1:.A4];[.B1:.C2])"] {
        resolver.clear_observations();
        assert_value(
            source,
            Observed::Error(ScalarError::Value),
            &resolver,
            &execution,
            position,
            Mode::Scalar,
        );
        assert_eq!(
            resolver.reads(),
            0,
            "{source} shape refusal must precede reads"
        );
    }

    for (source, expected) in [
        ("=HLOOKUP(2;[.E1:.H3];4;FALSE())", ScalarError::Reference),
        ("=VLOOKUP(2;[.A1:.C4];4;FALSE())", ScalarError::Reference),
    ] {
        resolver.clear_observations();
        assert_value(
            source,
            Observed::Error(expected),
            &resolver,
            &execution,
            position,
            Mode::Scalar,
        );
        assert_eq!(
            resolver.reads(),
            0,
            "{source} bounds refusal must precede reads"
        );
    }

    resolver.clear_observations();
    assert_value(
        "=OFFSET(([.A1]~[.B1]);0;0)",
        Observed::Error(ScalarError::Value),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn lookup_reference_keys_refuse_invalid_shapes_before_reading_cells() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-reference-key-shape-refusals");
    let position = Position::new("Main", 0, 0);

    for mode in [Mode::Scalar, Mode::Matrix] {
        for source in [
            "=MATCH([.B1];[Main.A1:Archive.A1];0)",
            "=MATCH([.B1];[.A1:.B2];0)",
            "=VLOOKUP([.B1];[Main.A1:Archive.A1];1;FALSE())",
            "=LOOKUP([.B1];[.A1:.A4];[Main.A1:Archive.A1])",
        ] {
            resolver.clear_observations();
            assert_value(
                source,
                Observed::Error(ScalarError::Value),
                &resolver,
                &execution,
                position,
                mode,
            );
            assert_eq!(
                resolver.reads(),
                0,
                "{source} in {mode:?} must refuse before key or data reads"
            );
        }
    }
}

#[test]
fn lookup_invalid_selectors_precede_reference_key_reads() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-invalid-selector-key-refusals");
    let position = Position::new("Main", 0, 0);

    // Selector and mode refusals are known from the complete argument
    // descriptors.  They must win before projecting or reading the
    // reference-valued key, in both evaluator modes.
    for mode in [Mode::Scalar, Mode::Matrix] {
        for (source, expected) in [
            ("=VLOOKUP([.B1];[.A1:.C4];0;FALSE())", ScalarError::Value),
            ("=VLOOKUP([.B1];[.A1:.C4];0;[.D1])", ScalarError::Value),
            (
                "=VLOOKUP([.B1];[.A1:.C4];[.D1];\"bad\")",
                ScalarError::Value,
            ),
            ("=VLOOKUP([.B1];[.A1:.C4];2;\"bad\")", ScalarError::Value),
            (
                "=HLOOKUP([.B1];[.E1:.H3];4;FALSE())",
                ScalarError::Reference,
            ),
            ("=MATCH([.B1];[.A1:.A4];\"bad\")", ScalarError::Value),
        ] {
            resolver.clear_observations();
            assert_value(
                source,
                Observed::Error(expected),
                &resolver,
                &execution,
                position,
                mode,
            );
            assert_eq!(
                resolver.reads(),
                0,
                "{source} in {mode:?} must refuse before reading its reference key"
            );
        }
    }

    // Array-valued keys still publish the selector error at their complete
    // broadcast shape.  The search data and every key element remain
    // unread because the invalid selector is rejected first.
    for (source, expected) in [
        (
            "=VLOOKUP({1|4};[.A1:.C4];0;FALSE())",
            vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Error(ScalarError::Value),
            ],
        ),
        (
            "=VLOOKUP({1|4};[.A1:.C4];2;\"bad\")",
            vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Error(ScalarError::Value),
            ],
        ),
        (
            "=HLOOKUP({1|4};[.E1:.H3];4;FALSE())",
            vec![
                CellObserved::Error(ScalarError::Reference),
                CellObserved::Error(ScalarError::Reference),
            ],
        ),
        (
            "=MATCH({1|2};[.A1:.A4];\"bad\")",
            vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Error(ScalarError::Value),
            ],
        ),
    ] {
        resolver.clear_observations();
        assert_value(
            source,
            Observed::Array {
                rows: 2,
                columns: 1,
                cells: expected,
            },
            &resolver,
            &execution,
            position,
            Mode::Matrix,
        );
        assert_eq!(
            resolver.reads(),
            0,
            "{source} must publish its key-shaped refusal without reads"
        );
    }

    // A reference key can itself be lifted.  Its output shape must be
    // preserved while the invalid table selector still prevents every key
    // cell from being consumed.
    for (source, rows, columns, expected) in [
        (
            "=VLOOKUP([.B1:.B2];[.A1:.C4];0;FALSE())",
            2,
            1,
            vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Error(ScalarError::Value),
            ],
        ),
        (
            "=HLOOKUP([.E1:.F1];[.E1:.H3];4;FALSE())",
            1,
            2,
            vec![
                CellObserved::Error(ScalarError::Reference),
                CellObserved::Error(ScalarError::Reference),
            ],
        ),
        (
            "=VLOOKUP([.B1];[.A1:.C4];{0|0};FALSE())",
            2,
            1,
            vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Error(ScalarError::Value),
            ],
        ),
        (
            "=VLOOKUP([.B1];[.A1:.C4];[.D1];{\"bad\"|\"bad\"})",
            2,
            1,
            vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Error(ScalarError::Value),
            ],
        ),
    ] {
        resolver.clear_observations();
        assert_value(
            source,
            Observed::Array {
                rows,
                columns,
                cells: expected,
            },
            &resolver,
            &execution,
            position,
            Mode::Matrix,
        );
        assert_eq!(
            resolver.reads(),
            0,
            "{source} must preserve its lifted shape without key reads"
        );
    }
}

#[test]
fn lookup_mixed_lifted_controls_keep_valid_coordinates() {
    let mut resolver = LookupResolver::standard();
    // The standard semantic fixture uses B1 as the text label "one".  This
    // focused selector test uses B1 as the numeric result so the successful
    // branch is directly visible as the requested 10.0 value.
    resolver.set(0, 1, FixtureCell::Number(10.0));
    let (_budget, _cancellation, execution) = execution("lookup-mixed-lifted-controls");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=VLOOKUP([.A1];[.A1:.C4];{0|2};FALSE())",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Number(10.0),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        3,
        "only the valid VLOOKUP coordinate may read key, search, and result"
    );

    resolver.clear_observations();
    assert_value(
        "=VLOOKUP([.A1];[.A1:.C4];[.A2];{\"bad\"|FALSE()})",
        Observed::Array {
            rows: 2,
            columns: 1,
            cells: vec![
                CellObserved::Error(ScalarError::Value),
                CellObserved::Number(10.0),
            ],
        },
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(
        resolver.reads(),
        4,
        "invalid range-mode coordinate must be read-free; the valid coordinate reads key, index, search, and result"
    );
}

#[test]
fn lookup_formula_error_keys_win_before_shape_refusals() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-formula-error-key-precedence");
    let position = Position::new("Main", 0, 0);

    for mode in [Mode::Scalar, Mode::Matrix] {
        for source in [
            "=VLOOKUP(1/0;1;1;FALSE())",
            "=HLOOKUP(1/0;([.A1]~[.A2]);1;FALSE())",
            "=MATCH(1/0;[.A1:.B2];0)",
            "=LOOKUP(1/0;[.A1:.A4];[.B1:.C2])",
        ] {
            resolver.clear_observations();
            assert_value(
                source,
                Observed::Error(ScalarError::DivisionByZero),
                &resolver,
                &execution,
                position,
                mode,
            );
            assert_eq!(
                resolver.reads(),
                0,
                "{source} in {mode:?} must preserve the key error without reads"
            );
        }
    }

    // A scalar formula error in the data argument must not displace the
    // earlier key error when both arguments are already values.
    for mode in [Mode::Scalar, Mode::Matrix] {
        resolver.clear_observations();
        assert_value(
            "=VLOOKUP(1/0;#N/A;1;FALSE())",
            Observed::Error(ScalarError::DivisionByZero),
            &resolver,
            &execution,
            position,
            mode,
        );
        assert_eq!(
            resolver.reads(),
            0,
            "a scalar data error must not replace the earlier key error"
        );
    }
}

#[test]
fn index_selects_one_union_record_and_clears_list_identity() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-index-list-identity");
    let position = Position::new("Main", 0, 0);

    assert_value(
        "=AREAS(INDEX(([.A1]~[.B1]~[.A1]);1;1;2))",
        Observed::Number(1.0),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_eq!(resolver.reads(), 0);
    assert_value(
        "=ISREF(INDEX(([.A1]~[.B1]~[.A1]);1;1;2))",
        Observed::Logical(true),
        &resolver,
        &execution,
        position,
        Mode::Scalar,
    );
    assert_value(
        r#"=ISREF(INDEX(INDIRECT("B2");1;1))"#,
        Observed::Logical(true),
        &resolver,
        &execution,
        position,
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 0);
}

#[test]
fn derived_reference_owned_roundtrip_drops_synthetic_lexical_owner() {
    let resolver = LookupResolver::standard();
    let (_budget, _cancellation, execution) = execution("lookup-owned-derived-reference");
    let expression = parse("=OFFSET([.A1:.B2];1;1)");
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, &resolver, &context, &Limits::default())
        .expect("OFFSET should produce a descriptor");
    assert!(
        matches!(result.value(), Value::Reference(reference) if reference.reference().is_none())
    );
    assert_eq!(resolver.reads(), 0);

    let owned = result
        .to_owned(&execution, &Limits::default())
        .expect("derived reference should be ownable");
    match owned.value() {
        OwnedValueView::Reference(reference) => {
            assert_eq!(reference.reference(), None);
            assert_eq!(reference.areas()[0].starts(), [0, 1, 1]);
            assert_eq!(reference.areas()[0].ends(), [1, 3, 3]);
        },
        other => panic!("expected owned derived reference, got {other:?}"),
    }
}

#[test]
fn sequential_choose_shape_probes_release_depth() {
    let resolver = LookupResolver::standard();
    let (budget, _cancellation, execution) = execution("lookup-sequential-choose-probes");
    let branches = std::iter::repeat_n("CHOOSE(SUM([.A1]);1;0)", 40)
        .collect::<Vec<_>>()
        .join("+");
    let formula = format!("=SUM(IF({{TRUE()}};{branches};0))");
    assert_value(
        &formula,
        Observed::Number(40.0),
        &resolver,
        &execution,
        Position::new("Main", 0, 0),
        Mode::Matrix,
    );
    assert_eq!(resolver.reads(), 40);
    assert_eq!(budget.used(litchi_core::Resource::Depth), 0);
    assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
}
