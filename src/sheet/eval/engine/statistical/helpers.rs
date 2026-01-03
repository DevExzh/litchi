use std::cmp::Ordering;
use std::result::Result as StdResult;

use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

pub(super) const EPS: f64 = 1e-12;

pub(super) async fn number_arg(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Option<f64>> {
    let v = evaluate_expression(ctx, current_sheet, expr).await?;
    Ok(to_number(&v))
}

pub(super) fn collect_numeric_values(
    values: Vec<CellValue>,
    func_name: &str,
) -> StdResult<Vec<f64>, CellValue> {
    let mut numbers = Vec::new();
    for value in values {
        match value {
            CellValue::Error(msg) => return Err(CellValue::Error(msg)),
            other => {
                if let Some(n) = to_number(&other) {
                    if n.is_nan() {
                        return Err(CellValue::Error(format!(
                            "{func_name} encountered an invalid numeric value"
                        )));
                    }
                    numbers.push(n);
                }
            },
        }
    }

    if numbers.is_empty() {
        return Err(CellValue::Error(format!(
            "{func_name} requires at least one numeric value in the array"
        )));
    }

    numbers.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    Ok(numbers)
}

pub(super) fn collect_numeric_values_unsorted(
    values: Vec<CellValue>,
    func_name: &str,
) -> StdResult<Vec<f64>, CellValue> {
    let mut numbers = Vec::new();
    for value in values {
        match value {
            CellValue::Error(msg) => return Err(CellValue::Error(msg)),
            other => {
                if let Some(n) = to_number(&other) {
                    if n.is_nan() {
                        return Err(CellValue::Error(format!(
                            "{func_name} encountered an invalid numeric value"
                        )));
                    }
                    numbers.push(n);
                }
            },
        }
    }

    if numbers.is_empty() {
        return Err(CellValue::Error(format!(
            "{func_name} requires at least one numeric value in the reference"
        )));
    }

    Ok(numbers)
}

pub(super) fn to_positive_index(
    value: f64,
    func_name: &str,
    arg_name: &str,
) -> StdResult<usize, CellValue> {
    if !value.is_finite() {
        return Err(CellValue::Error(format!(
            "{func_name} {arg_name} must be a finite positive integer"
        )));
    }
    let rounded = value.round();
    if (value - rounded).abs() > EPS {
        return Err(CellValue::Error(format!(
            "{func_name} {arg_name} must be an integer"
        )));
    }
    if rounded < 1.0 {
        return Err(CellValue::Error(format!(
            "{func_name} {arg_name} must be greater than or equal to 1"
        )));
    }
    Ok(rounded as usize)
}
