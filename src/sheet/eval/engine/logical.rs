use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::{EngineCtx, evaluate_expression, for_each_value_in_expr, to_bool};

pub(crate) fn eval_if<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error("IF expects 2 or 3 arguments".to_string()));
    }

    // Condition
    let cond_val = evaluate_expression(ctx, current_sheet, &args[0])?;
    let cond = to_bool(&cond_val);

    if cond {
        evaluate_expression(ctx, current_sheet, &args[1])
    } else if args.len() == 3 {
        evaluate_expression(ctx, current_sheet, &args[2])
    } else {
        Ok(CellValue::Bool(false))
    }
}

pub(crate) fn eval_and<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Bool(true));
    }
    for arg in args {
        let mut all_true = true;
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if !to_bool(v) {
                all_true = false;
            }
            Ok(())
        })?;
        if !all_true {
            return Ok(CellValue::Bool(false));
        }
    }
    Ok(CellValue::Bool(true))
}

pub(crate) fn eval_or<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Bool(false));
    }
    let mut any_true = false;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if to_bool(v) {
                any_true = true;
            }
            Ok(())
        })?;
        if any_true {
            return Ok(CellValue::Bool(true));
        }
    }
    Ok(CellValue::Bool(false))
}

pub(crate) fn eval_not<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("NOT expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0])?;
    Ok(CellValue::Bool(!to_bool(&v)))
}
