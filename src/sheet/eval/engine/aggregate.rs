use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::to_number;
use super::{EngineCtx, for_each_value_in_expr};

pub(crate) fn eval_sum<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut total = 0.0f64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                total += n;
            }
            Ok(())
        })?;
    }
    Ok(CellValue::Float(total))
}

pub(crate) fn eval_product<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut product = 1.0f64;
    let mut found_numeric = false;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if let Some(n) = to_number(v) {
                product *= n;
                found_numeric = true;
            }
            Ok(())
        })?;
    }
    if !found_numeric {
        Ok(CellValue::Float(0.0))
    } else {
        Ok(CellValue::Float(product))
    }
}

pub(crate) fn eval_min<C: EngineCtx + ?Sized>(
    ctx: &C,
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
        })?;
    }
    Ok(CellValue::Float(min_val.unwrap_or(0.0)))
}

pub(crate) fn eval_max<C: EngineCtx + ?Sized>(
    ctx: &C,
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
        })?;
    }
    Ok(CellValue::Float(max_val.unwrap_or(0.0)))
}

pub(crate) fn eval_average<C: EngineCtx + ?Sized>(
    ctx: &C,
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
        })?;
    }
    if count == 0 {
        return Ok(CellValue::Error("AVERAGE of empty set".to_string()));
    }
    Ok(CellValue::Float(total / count as f64))
}

pub(crate) fn eval_count<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if matches!(
                v,
                CellValue::Int(_) | CellValue::Float(_) | CellValue::DateTime(_)
            ) {
                count += 1;
            }
            Ok(())
        })?;
    }
    Ok(CellValue::Int(count as i64))
}

pub(crate) fn eval_counta<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut count = 0u64;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if !matches!(v, CellValue::Empty) {
                count += 1;
            }
            Ok(())
        })?;
    }
    Ok(CellValue::Int(count as i64))
}
