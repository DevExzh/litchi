use crate::sheet::eval::engine::{EvalCtx, evaluate_expression};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use chrono::Datelike;

use super::helpers::{SECONDS_PER_DAY, coerce_date_value, coerce_time_fraction};

pub(crate) async fn eval_year(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("YEAR expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_date_value(&value) {
        Some(date) => Ok(CellValue::Int(date.year() as i64)),
        None => Ok(CellValue::Error(
            "YEAR expects a valid date serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_month(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("MONTH expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_date_value(&value) {
        Some(date) => Ok(CellValue::Int(date.month() as i64)),
        None => Ok(CellValue::Error(
            "MONTH expects a valid date serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_day(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("DAY expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_date_value(&value) {
        Some(date) => Ok(CellValue::Int(date.day() as i64)),
        None => Ok(CellValue::Error(
            "DAY expects a valid date serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_hour(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("HOUR expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_time_fraction(&value) {
        Some(frac) => {
            let seconds = (frac * SECONDS_PER_DAY).rem_euclid(SECONDS_PER_DAY);
            Ok(CellValue::Int((seconds / 3600.0).floor() as i64))
        },
        None => Ok(CellValue::Error(
            "HOUR expects a valid time serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_minute(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("MINUTE expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_time_fraction(&value) {
        Some(frac) => {
            let seconds = (frac * SECONDS_PER_DAY).rem_euclid(SECONDS_PER_DAY);
            let minutes = (seconds / 60.0).floor() as i64 % 60;
            Ok(CellValue::Int(minutes))
        },
        None => Ok(CellValue::Error(
            "MINUTE expects a valid time serial or text".to_string(),
        )),
    }
}

pub(crate) async fn eval_second(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("SECOND expects 1 argument".to_string()));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    match coerce_time_fraction(&value) {
        Some(frac) => {
            let seconds = (frac * SECONDS_PER_DAY).rem_euclid(SECONDS_PER_DAY);
            Ok(CellValue::Int((seconds % 60.0).round() as i64))
        },
        None => Ok(CellValue::Error(
            "SECOND expects a valid time serial or text".to_string(),
        )),
    }
}
