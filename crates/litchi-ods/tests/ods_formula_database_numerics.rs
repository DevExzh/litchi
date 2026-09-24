//! Independent extreme-value database aggregation vectors.

use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Profile};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationFailure, ScalarError,
        value::{self, CellRead, Context, Mode, Position, Resolver, SheetExtent, Value},
    },
    expression::Expression,
};
use std::num::{NonZeroU64, NonZeroUsize};

struct NoReads;
impl Resolver for NoReads {
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
        panic!("inline database must not read a provider")
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

fn execution() -> ExecutionContext {
    let (_, token) = CancellationSource::pair();
    ExecutionContext::new(
        Budget::root("database-extremes", Limits::for_profile(Profile::Server)),
        token,
        ExecutionLimits::new(
            NonZeroUsize::MIN,
            NonZeroUsize::MIN,
            NonZeroU64::new(1 << 20).expect("positive"),
            0,
        )
        .expect("valid limits"),
    )
}

#[derive(Clone, Copy)]
enum Expected {
    Exact(f64),
    Near(f64),
    NumberError,
}

fn check(function: &str, values: &[f64], expected: Expected) {
    let mut source = format!("={function}({{\"Value\";\"Take\"");
    for value in values {
        source.push_str(&format!("|{value:e};1"));
    }
    source.push_str("};1;{2|1})");
    let expression = Expression::parse(&source).expect("numeric inline fixture");
    let execution = execution();
    let context = Context::new(&execution, Position::new("Main", 0, 0)).with_mode(Mode::Matrix);
    let result = value::evaluate(&expression, &NoReads, &context, &value::Limits::default())
        .unwrap_or_else(|error| panic!("{source}: {error}"));
    match (result.value(), expected) {
        (Value::Number(actual), Expected::Exact(expected)) => {
            assert_eq!(actual.to_bits(), expected.to_bits(), "{source}")
        },
        (Value::Number(actual), Expected::Near(expected)) => {
            assert!(
                actual.is_finite() && (actual - expected).abs() <= expected.abs() * 2e-12,
                "{source}: {actual} != {expected}"
            );
        },
        (Value::Error(ScalarError::Number), Expected::NumberError) => {},
        (actual, _) => panic!("unexpected result for {source}: {actual:?}"),
    }
}

#[test]
fn database_sum_and_mean_preserve_finite_cancellation_and_extreme_scale() {
    check("DSUM", &[f64::MAX, 3.0, -f64::MAX], Expected::Exact(3.0));
    check(
        "DSUM",
        &[f64::from_bits(1), f64::MAX, -f64::MAX],
        Expected::Exact(f64::from_bits(1)),
    );
    check("DAVERAGE", &[f64::MAX, f64::MAX], Expected::Exact(f64::MAX));
    check(
        "DAVERAGE",
        &[f64::from_bits(1), f64::from_bits(1)],
        Expected::Exact(f64::from_bits(1)),
    );
    check("DSUM", &[f64::MAX, f64::MAX], Expected::NumberError);
}

#[test]
fn database_product_retains_exponent_until_final_result() {
    check(
        "DPRODUCT",
        &[1e308, 1e308, 1e-308, 1e-308],
        Expected::Near(1.0),
    );
    check("DPRODUCT", &[1e308, 1e308, 0.0], Expected::Exact(0.0));
    check(
        "DPRODUCT",
        &[-1e308, 1e308, 1e-308, 1e-308],
        Expected::Near(-1.0),
    );
}

#[test]
fn database_statistics_keep_finite_standard_deviation_when_variance_overflows() {
    check("DSTDEVP", &[1e154, -1e154], Expected::Near(1e154));
    check("DVARP", &[1e154, -1e154], Expected::Near(1e308));
    check(
        "DSTDEV",
        &[1e154, -1e154],
        Expected::Near(std::f64::consts::SQRT_2 * 1e154),
    );
    check("DVAR", &[1e154, -1e154], Expected::NumberError);
    check(
        "DSTDEVP",
        &[1e154, -1e154, 0.0],
        Expected::Near((2.0_f64 / 3.0).sqrt() * 1e154),
    );
    let first = 1e150_f64;
    let next = f64::from_bits(first.to_bits() + 1);
    check(
        "DSTDEVP",
        &[first, next],
        Expected::Near((next - first) * 0.5),
    );
    check(
        "DSTDEV",
        &[first, next],
        Expected::Near((next - first) / std::f64::consts::SQRT_2),
    );
}
