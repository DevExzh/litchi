//! Focused coverage for the shared OpenFormula criterion text profile.
//!
//! A bare numeric-looking Text criterion remains Text. An explicit comparator
//! such as `=3` uses numeric comparison and therefore selects Number cells.
//! Both database and conditional aggregate callers use the same private
//! matcher, so this test keeps the two paths aligned.

use std::num::{NonZeroU64, NonZeroUsize};

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
    Text(String),
}

#[derive(Debug)]
struct CriteriaResolver {
    rows: usize,
    columns: usize,
    cells: Vec<FixtureCell>,
}

impl CriteriaResolver {
    fn new() -> Self {
        let mut resolver = Self {
            rows: 4,
            columns: 4,
            cells: (0..16).map(|_| FixtureCell::Empty).collect(),
        };
        // Database: A1:B4.  The criterion range is D1:D2.
        resolver.set(0, 0, FixtureCell::Text("Value".to_owned()));
        resolver.set(0, 1, FixtureCell::Text("Key".to_owned()));
        resolver.set(1, 0, FixtureCell::Number(10.0));
        resolver.set(1, 1, FixtureCell::Text("3".to_owned()));
        resolver.set(2, 0, FixtureCell::Number(20.0));
        resolver.set(2, 1, FixtureCell::Text("1".to_owned()));
        resolver.set(3, 0, FixtureCell::Number(40.0));
        resolver.set(3, 1, FixtureCell::Number(3.0));
        resolver.set(0, 3, FixtureCell::Text("Key".to_owned()));
        resolver.set(1, 3, FixtureCell::Text("3".to_owned()));
        resolver.set(2, 3, FixtureCell::Text("Key".to_owned()));
        resolver.set(3, 3, FixtureCell::Text("=3".to_owned()));
        resolver
    }

    fn set(&mut self, row: usize, column: usize, value: FixtureCell) {
        let index = row * self.columns + column;
        self.cells[index] = value;
    }
}

impl Resolver for CriteriaResolver {
    fn sheet_extent(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(SheetExtent::new(self.rows, self.columns)))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _execution: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        if sheet != "Main" {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        let Some(index) = row
            .checked_mul(self.columns)
            .and_then(|index| index.checked_add(column))
        else {
            return Ok(CellRead::Empty);
        };
        Ok(match self.cells.get(index) {
            None | Some(FixtureCell::Empty) => CellRead::Empty,
            Some(FixtureCell::Number(value)) => CellRead::Number(*value),
            Some(FixtureCell::Text(value)) => CellRead::Text(value),
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        _execution: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(0))
    }

    fn sheet_name_at(
        &self,
        index: usize,
        _execution: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok((index == 0).then_some("Main"))
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(1)
    }
}

fn execution() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "ods-formula-criterion-text-profile",
        CoreLimits::for_profile(Profile::Server),
    );
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(1_024).expect("one in-flight KiB"),
        0,
    )
    .expect("valid execution limits");
    (cancellation, ExecutionContext::new(budget, token, limits))
}

fn evaluate_number(source: &str, resolver: &CriteriaResolver, execution: &ExecutionContext) -> f64 {
    let expression = Expression::parse(source).expect("formula parses");
    let context = Context::new(execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, resolver, &context, &Limits::default())
        .expect("formula evaluates");
    match result.value() {
        Value::Number(value) => value,
        other => panic!("{source} produced {other:?}"),
    }
}

#[test]
fn bare_numeric_text_stays_text_in_database_and_conditional_matchers() {
    let resolver = CriteriaResolver::new();
    let (_cancellation, execution) = execution();

    assert_eq!(
        evaluate_number(
            r#"=DSUM([.A1:.B4];"Value";[.D1:.D2])"#,
            &resolver,
            &execution,
        ),
        10.0
    );
    assert_eq!(
        evaluate_number(
            r#"=DSUM([.A1:.B4];"Value";[.D3:.D4])"#,
            &resolver,
            &execution,
        ),
        40.0
    );
    assert_eq!(
        evaluate_number(
            r#"=SUMIF([.B2:.B4];[.D2];[.A2:.A4])"#,
            &resolver,
            &execution,
        ),
        10.0
    );
    assert_eq!(
        evaluate_number(r#"=SUMIF([.B2:.B4];"=3";[.A2:.A4])"#, &resolver, &execution,),
        40.0
    );
}
