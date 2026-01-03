use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, for_each_value_in_expr, to_number};

pub(crate) async fn eval_average(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut total = 0.0f64;
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                total += n;
                count += 1;
            }
            Ok(())
        })
        .await?;
    }
    if count == 0 {
        return Ok(CellValue::Error("AVERAGE of empty set".to_string()));
    }
    Ok(CellValue::Float(total / count as f64))
}

pub(crate) async fn eval_averagea(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut total = 0.0f64;
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            match v {
                CellValue::Empty => {},
                CellValue::Bool(true) => {
                    total += 1.0;
                    count += 1;
                },
                CellValue::Bool(false) => {
                    count += 1;
                },
                CellValue::Int(i) => {
                    total += *i as f64;
                    count += 1;
                },
                CellValue::Float(f) => {
                    total += *f;
                    count += 1;
                },
                CellValue::DateTime(d) => {
                    total += *d;
                    count += 1;
                },
                CellValue::String(_) => {
                    count += 1;
                },
                CellValue::Error(_) => {},
                CellValue::Formula {
                    cached_value: Some(value),
                    ..
                } => match &**value {
                    CellValue::Bool(true) => {
                        total += 1.0;
                        count += 1;
                    },
                    CellValue::Bool(false) => {
                        count += 1;
                    },
                    CellValue::Int(i) => {
                        total += *i as f64;
                        count += 1;
                    },
                    CellValue::Float(f) => {
                        total += *f;
                        count += 1;
                    },
                    CellValue::DateTime(d) => {
                        total += *d;
                        count += 1;
                    },
                    CellValue::String(_) => {
                        count += 1;
                    },
                    _ => {},
                },
                CellValue::Formula { .. } => {
                    count += 1;
                },
            }
            Ok(())
        })
        .await?;
    }
    if count == 0 {
        return Ok(CellValue::Error("AVERAGEA of empty set".to_string()));
    }
    Ok(CellValue::Float(total / count as f64))
}

pub(crate) async fn eval_avedev(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut values = Vec::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                values.push(n);
            }
            Ok(())
        })
        .await?;
    }

    if values.is_empty() {
        return Ok(CellValue::Error(
            "AVEDEV requires at least one numeric value".to_string(),
        ));
    }

    let mean = values.iter().sum::<f64>() / values.len() as f64;
    let total_dev = values.iter().map(|v| (v - mean).abs()).sum::<f64>();
    Ok(CellValue::Float(total_dev / values.len() as f64))
}
