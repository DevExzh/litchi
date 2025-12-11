use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::{EngineCtx, evaluate_expression, for_each_value_in_expr, is_blank, to_text};

pub(crate) fn eval_len<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("LEN expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0])?;
    let s = to_text(&v);
    Ok(CellValue::Int(s.chars().count() as i64))
}

pub(crate) fn eval_lower<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("LOWER expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0])?;
    let s = to_text(&v).to_lowercase();
    Ok(CellValue::String(s))
}

pub(crate) fn eval_upper<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("UPPER expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0])?;
    let s = to_text(&v).to_uppercase();
    Ok(CellValue::String(s))
}

pub(crate) fn eval_trim<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("TRIM expects 1 argument".to_string()));
    }
    let v = evaluate_expression(ctx, current_sheet, &args[0])?;
    let s = to_text(&v);
    Ok(CellValue::String(s.trim().to_string()))
}

pub(crate) fn eval_concat<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    let mut out = String::new();
    for arg in args {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            out.push_str(&to_text(v));
            Ok(())
        })?;
    }
    Ok(CellValue::String(out))
}

pub(crate) fn eval_textjoin<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 3 {
        return Ok(CellValue::Error(
            "TEXTJOIN expects at least 3 arguments (delimiter, ignore_empty, text1, ...)"
                .to_string(),
        ));
    }

    let delimiter_val = evaluate_expression(ctx, current_sheet, &args[0])?;
    let delimiter = to_text(&delimiter_val);

    let ignore_empty_val = evaluate_expression(ctx, current_sheet, &args[1])?;
    let ignore_empty = super::to_bool(&ignore_empty_val);

    let mut parts: Vec<String> = Vec::new();
    for arg in &args[2..] {
        for_each_value_in_expr(ctx, current_sheet, arg, |v| {
            if ignore_empty && is_blank(v) {
                return Ok(());
            }
            parts.push(to_text(v));
            Ok(())
        })?;
    }

    Ok(CellValue::String(parts.join(&delimiter)))
}
