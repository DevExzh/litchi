use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::{EvalCtx, evaluate_expression, for_each_value_in_expr, to_bool};

pub(crate) async fn eval_if(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error("IF expects 2 or 3 arguments".to_string()));
    }

    // Condition
    let cond_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let cond = to_bool(&cond_val);

    if cond {
        evaluate_expression(ctx, current_sheet, &args[1]).await
    } else if args.len() == 3 {
        evaluate_expression(ctx, current_sheet, &args[2]).await
    } else {
        Ok(CellValue::Bool(false))
    }
}

pub(crate) async fn eval_ifs(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        return Ok(CellValue::Error(
            "IFS expects an even number of arguments (condition/result pairs)".to_string(),
        ));
    }
    for pair in args.chunks(2) {
        let condition = evaluate_expression(ctx, current_sheet, &pair[0]).await?;
        if let CellValue::Error(_) = condition {
            return Ok(condition);
        }
        if to_bool(&condition) {
            return evaluate_expression(ctx, current_sheet, &pair[1]).await;
        }
    }
    Ok(CellValue::Error("#N/A".to_string()))
}

pub(crate) async fn eval_and(
    ctx: EvalCtx<'_>,
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
        })
        .await?;
        if !all_true {
            return Ok(CellValue::Bool(false));
        }
    }
    Ok(CellValue::Bool(true))
}

pub(crate) async fn eval_or(
    ctx: EvalCtx<'_>,
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
        })
        .await?;
        if any_true {
            return Ok(CellValue::Bool(true));
        }
    }
    Ok(CellValue::Bool(false))
}

pub(crate) async fn eval_xor(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() {
        return Ok(CellValue::Error(
            "XOR expects at least 1 argument".to_string(),
        ));
    }
    let mut parity = false;
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if to_bool(v) {
                parity = !parity;
            }
            Ok(())
        })
        .await?;
    }
    Ok(CellValue::Bool(parity))
}

pub(crate) async fn eval_not(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("NOT expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    Ok(CellValue::Bool(!to_bool(&v)))
}

pub(crate) async fn eval_true(_: EvalCtx<'_>, _: &str, args: &[Expr]) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("TRUE expects no arguments".to_string()));
    }
    Ok(CellValue::Bool(true))
}

pub(crate) async fn eval_false(_: EvalCtx<'_>, _: &str, args: &[Expr]) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("FALSE expects no arguments".to_string()));
    }
    Ok(CellValue::Bool(false))
}

pub(crate) async fn eval_switch(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 {
        return Ok(CellValue::Error(
            "SWITCH expects at least 3 arguments (expression, value1, result1, ...)".to_string(),
        ));
    }

    let target = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    if let CellValue::Error(_) = target {
        return Ok(target);
    }

    let mut i = 1;
    while i + 1 < args.len() {
        let val = evaluate_expression(ctx, current_sheet, &args[i]).await?;
        if val == target {
            return evaluate_expression(ctx, current_sheet, &args[i + 1]).await;
        }
        i += 2;
    }

    // Default value if exists (odd number of arguments total)
    if args.len().is_multiple_of(2) {
        evaluate_expression(ctx, current_sheet, &args[args.len() - 1]).await
    } else {
        Ok(CellValue::Error("#N/A".to_string()))
    }
}
