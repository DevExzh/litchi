use std::result::Result as StdResult;

use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::super::{EvalCtx, for_each_value_in_expr, to_number};

pub(crate) async fn eval_min(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut min_val: Option<f64> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                min_val = Some(match min_val {
                    Some(cur) => cur.min(n),
                    None => n,
                });
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Float(min_val.unwrap_or(0.0)))
}

pub(crate) async fn eval_max(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut max_val: Option<f64> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                max_val = Some(match max_val {
                    Some(cur) => cur.max(n),
                    None => n,
                });
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Float(max_val.unwrap_or(0.0)))
}

pub(crate) async fn eval_mina(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut min_val: Option<f64> = None;
    let mut error: Option<CellValue> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if error.is_some() {
                return Ok(());
            }
            match coerce_for_mina_maxa(v) {
                Ok(Some(value)) => {
                    min_val = Some(match min_val {
                        Some(cur) => cur.min(value),
                        None => value,
                    });
                },
                Ok(None) => {},
                Err(msg) => {
                    error = Some(CellValue::Error(msg));
                },
            }
            Ok(())
        })
        .await?;
        if error.is_some() {
            break;
        }
    }

    if let Some(err) = error {
        return Ok(err);
    }

    match min_val {
        Some(value) => Ok(CellValue::Float(value)),
        None => Ok(CellValue::Float(0.0)),
    }
}

pub(crate) async fn eval_maxa(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut max_val: Option<f64> = None;
    let mut error: Option<CellValue> = None;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if error.is_some() {
                return Ok(());
            }
            match coerce_for_mina_maxa(v) {
                Ok(Some(value)) => {
                    max_val = Some(match max_val {
                        Some(cur) => cur.max(value),
                        None => value,
                    });
                },
                Ok(None) => {},
                Err(msg) => {
                    error = Some(CellValue::Error(msg));
                },
            }
            Ok(())
        })
        .await?;
        if error.is_some() {
            break;
        }
    }

    if let Some(err) = error {
        return Ok(err);
    }

    match max_val {
        Some(value) => Ok(CellValue::Float(value)),
        None => Ok(CellValue::Float(0.0)),
    }
}

fn coerce_for_mina_maxa(value: &CellValue) -> StdResult<Option<f64>, String> {
    match value {
        CellValue::Empty => Ok(None),
        CellValue::Bool(true) => Ok(Some(1.0)),
        CellValue::Bool(false) => Ok(Some(0.0)),
        CellValue::Int(i) => Ok(Some(*i as f64)),
        CellValue::Float(f) => Ok(Some(*f)),
        CellValue::DateTime(d) => Ok(Some(*d)),
        CellValue::String(_) => Ok(Some(0.0)),
        CellValue::Error(msg) => Err(msg.clone()),
        CellValue::Formula {
            cached_value: Some(inner),
            ..
        } => coerce_for_mina_maxa(inner),
        CellValue::Formula { .. } => Ok(None),
    }
}
