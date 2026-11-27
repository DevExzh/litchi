//! Date/time parser adapter around the shared VALUE grammar.
//!
//! Date/time parsing remains implemented by the inspection owner. This layer
//! charges borrowed input and, when grouped normalization is long, reserves
//! caller-owned scratch with the evaluator budget before delegating.

use super::super::{EvaluationFailure, EvaluationResult, Evaluator, ScalarError};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum TextKind {
    Date,
    Time,
}

pub(super) fn parse<'a>(
    evaluator: &mut Evaluator<'a, '_, '_>,
    text: &str,
    kind: TextKind,
) -> EvaluationResult<Result<f64, ScalarError>> {
    evaluator.charge_bytes(text.len())?;
    evaluator
        .context
        .execution
        .check()
        .map_err(super::super::map_execution_error)?;

    let parser = match kind {
        TextKind::Date => parse_date,
        TextKind::Time => parse_time,
    };
    let result = if super::super::inspection::parse_value::needs_scratch(text) {
        let bytes = text.len();
        let reservation = evaluator.reserve_storage(bytes, "formula date/time parser scratch")?;
        let mut scratch = Vec::new();
        scratch
            .try_reserve_exact(bytes)
            .map_err(|source| EvaluationFailure::Allocation {
                resource: "formula date/time parser scratch",
                source,
            })?;
        scratch.resize(bytes, 0);
        evaluator
            .context
            .execution
            .check()
            .map_err(super::super::map_execution_error)?;
        let result = parser(text, &mut scratch);
        drop(scratch);
        drop(reservation);
        evaluator
            .context
            .execution
            .check()
            .map_err(super::super::map_execution_error)?;
        result
    } else {
        parser(text, &mut [])
    };
    Ok(result)
}

fn parse_date(text: &str, scratch: &mut [u8]) -> Result<f64, ScalarError> {
    if scratch.is_empty() {
        super::super::inspection::parse_value::parse_date_value(text)
    } else {
        super::super::inspection::parse_value::parse_date_value_with_scratch(text, scratch)
    }
}

fn parse_time(text: &str, scratch: &mut [u8]) -> Result<f64, ScalarError> {
    if scratch.is_empty() {
        super::super::inspection::parse_value::parse_time_value(text)
    } else {
        super::super::inspection::parse_value::parse_time_value_with_scratch(text, scratch)
    }
}
