//! Independent exact-rational and selection-mask observations over cell references.

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
use std::num::{NonZeroU64, NonZeroUsize};

struct Cells(Vec<[f64; 3]>);
impl Resolver for Cells {
    fn sheet_extent(
        &self,
        sheet: &str,
        _: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok((sheet == "Main").then_some(SheetExtent::new(self.0.len(), 3)))
    }
    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        _: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        assert_eq!(sheet, "Main");
        Ok(CellRead::Number(self.0[row][column]))
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

fn bits(value: &serde_json::Value) -> u64 {
    u64::from_str_radix(value.as_str().expect("binary64 hexadecimal bits"), 16)
        .expect("valid binary64 bits")
}

fn compare(function: &str) {
    let document: serde_json::Value = serde_json::from_str(include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/numeric-goldens.json"
    )).expect("retained independent oracle");
    let rows = document["observations"].as_array().expect("observations");
    assert_eq!(rows.len(), 288);
    let budget = Budget::root(
        "conditional-oracle",
        CoreLimits::for_profile(Profile::Server),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("execution limits"),
    );
    for row in rows.iter().filter(|row| row["function"] == function) {
        let formula = row["formula"].as_str().expect("formula");
        let expression = Expression::parse(formula).expect("reference expression parses");
        let cells = Cells(
            row["cells_bits"]
                .as_array()
                .expect("cell rows")
                .iter()
                .map(|row| std::array::from_fn(|column| f64::from_bits(bits(&row[column]))))
                .collect(),
        );
        for mode in [Mode::Scalar, Mode::Matrix] {
            let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(mode);
            let result = evaluate(&expression, &cells, &context, &Limits::default())
                .unwrap_or_else(|error| panic!("{formula}: {error}"));
            if let Some(error) = row["error"].as_str() {
                let expected = match error {
                    "Number" => ScalarError::Number,
                    "DivisionByZero" => ScalarError::DivisionByZero,
                    other => panic!("unknown oracle error {other}"),
                };
                assert!(
                    matches!(result.value(), Value::Error(actual) if actual == expected),
                    "case {} {formula}: {:?}",
                    row["case"],
                    result.value()
                );
            } else {
                let Value::Number(actual) = result.value() else {
                    panic!("case {} {formula}: {:?}", row["case"], result.value());
                };
                assert_eq!(
                    actual.to_bits(),
                    bits(&row["expected_bits"]),
                    "case {} {formula}: exact nearest-even reduction",
                    row["case"]
                );
            }
        }
    }
}

#[test]
fn sumif_matches_exact_selection_oracle() {
    compare("SUMIF");
}
#[test]
fn sumifs_matches_exact_selection_oracle() {
    compare("SUMIFS");
}
#[test]
fn countif_matches_exact_selection_oracle() {
    compare("COUNTIF");
}
#[test]
fn countifs_matches_exact_selection_oracle() {
    compare("COUNTIFS");
}
#[test]
fn averageif_matches_exact_selection_oracle() {
    compare("AVERAGEIF");
}
#[test]
fn averageifs_matches_exact_selection_oracle() {
    compare("AVERAGEIFS");
}
