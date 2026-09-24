//! Independent bigint numerical observations over represented binary64 operands.

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Profile,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        evaluate_scalar,
        value::{CellRead, Context, Limits, Position, Resolver, SheetExtent, Value, evaluate},
    },
    expression::Expression,
};
use std::num::{NonZeroU64, NonZeroUsize};

struct NoCells;

impl Resolver for NoCells {
    fn sheet_extent(
        &self,
        _: &str,
        _: &ExecutionContext,
    ) -> Result<Option<SheetExtent>, EvaluationFailure> {
        Ok(None)
    }

    fn read_cell<'a>(
        &'a self,
        _: &str,
        _: usize,
        _: usize,
        _: &ExecutionContext,
    ) -> Result<CellRead<'a>, EvaluationFailure> {
        panic!("closed numerical oracle must not read cells")
    }

    fn sheet_index(
        &self,
        _: &str,
        _: &ExecutionContext,
    ) -> Result<Option<usize>, EvaluationFailure> {
        Ok(None)
    }

    fn sheet_name_at(
        &self,
        _: usize,
        _: &ExecutionContext,
    ) -> Result<Option<&str>, EvaluationFailure> {
        Ok(None)
    }

    fn sheet_count(&self, _: &ExecutionContext) -> Result<usize, EvaluationFailure> {
        Ok(0)
    }
}

fn check(formula: &str, row: &serde_json::Value, actual: Result<f64, ScalarError>) {
    let Some(expected_hex) = row["expected_bits"].as_str() else {
        assert_eq!(actual, Err(ScalarError::Number), "{formula}");
        return;
    };
    let bits = u64::from_str_radix(expected_hex, 16).expect("binary64 bits");
    let expected = f64::from_bits(bits);
    let actual =
        actual.unwrap_or_else(|error| panic!("{formula}: expected {expected:e}, got {error:?}"));
    assert!(actual.is_finite(), "{formula}: non-finite result");
    if expected == 0.0 {
        assert_eq!(actual.to_bits(), 0, "{formula}: canonical zero");
        return;
    }
    assert_eq!(
        actual.is_sign_negative(),
        expected.is_sign_negative(),
        "{formula}"
    );
    // The retained reference and fixed-width kernel both preserve the exact
    // mathematical integer until a single nearest-even Number conversion.
    let max_ulps = 0;
    assert!(
        actual.to_bits().abs_diff(bits) <= max_ulps,
        "{formula}: {actual:e} versus {expected:e}, allowed {max_ulps} ULP"
    );
}

fn compare_function(function: &str) {
    let document: serde_json::Value = serde_json::from_str(include_str!(
        "../../../docs/report/spec-gap-validation-evidence/ods-formula-discrete-math/numeric-goldens.json"
    )).expect("retained oracle JSON");
    let rows = document["observations"].as_array().expect("oracle rows");
    assert_eq!(rows.len(), 361);
    let budget = Budget::root("discrete-oracle", CoreLimits::for_profile(Profile::Server));
    let (_cancellation, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        budget,
        token,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MIN, 0)
            .expect("execution limits"),
    );
    let context = Context::new(&execution, Position::new("Main", 0, 0));
    for row in rows.iter().filter(|row| row["function"] == function) {
        let formula = row["formula"].as_str().expect("formula");
        let expression = Expression::parse(formula).expect("closed expression parses");
        let result = evaluate(&expression, &NoCells, &context, &Limits::default())
            .unwrap_or_else(|error| panic!("{formula}: {error}"));
        check(
            formula,
            row,
            match result.value() {
                Value::Number(value) => Ok(value),
                Value::Error(error) => Err(error),
                other => panic!("{formula}: unexpected value {other:?}"),
            },
        );
        let scalar = evaluate_scalar(
            &expression,
            &EvaluationContext::new(&execution),
            &EvaluationLimits::default(),
        )
        .unwrap_or_else(|error| panic!("{formula}: {error}"));
        check(
            formula,
            row,
            match scalar.value() {
                ScalarValue::Number(value) => Ok(*value),
                ScalarValue::Error(error) => Err(*error),
                other => panic!("{formula}: unexpected scalar {other:?}"),
            },
        );
    }
}

macro_rules! oracle_test {
    ($test:ident, $name:literal) => {
        #[test]
        fn $test() {
            compare_function($name);
        }
    };
}

oracle_test!(combin_matches_integer_oracle, "COMBIN");
oracle_test!(combina_matches_integer_oracle, "COMBINA");
oracle_test!(fact_matches_integer_oracle, "FACT");
oracle_test!(factdouble_matches_integer_oracle, "FACTDOUBLE");
oracle_test!(gcd_matches_integer_oracle, "GCD");
oracle_test!(lcm_matches_integer_oracle, "LCM");
oracle_test!(multinomial_matches_integer_oracle, "MULTINOMIAL");
oracle_test!(even_matches_integer_oracle, "EVEN");
oracle_test!(odd_matches_integer_oracle, "ODD");
oracle_test!(delta_matches_integer_oracle, "DELTA");
oracle_test!(gestep_matches_integer_oracle, "GESTEP");
