use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use chrono::Datelike;

use super::helpers::{coerce_date_value, weekday_number, weeknum_value};

pub(crate) async fn eval_weekday(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "WEEKDAY expects 1 or 2 arguments (serial_number, [return_type])".to_string(),
        ));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let date = match coerce_date_value(&value) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WEEKDAY expects a valid date serial or text".to_string(),
            ));
        },
    };
    let return_type = if args.len() == 2 {
        let rt_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&rt_val) {
            Some(n) if matches!(n as i32, 1..=3) => n as i32,
            _ => {
                return Ok(CellValue::Error(
                    "WEEKDAY return_type must be 1, 2, or 3".to_string(),
                ));
            },
        }
    } else {
        1
    };
    let number = match weekday_number(date.weekday(), return_type) {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "WEEKDAY return_type not supported".to_string(),
            ));
        },
    };
    Ok(CellValue::Int(number))
}

pub(crate) async fn eval_weeknum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "WEEKNUM expects 1 or 2 arguments (serial_number, [return_type])".to_string(),
        ));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let date = match coerce_date_value(&value) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WEEKNUM expects a valid date serial or text".to_string(),
            ));
        },
    };
    let return_type = if args.len() == 2 {
        let rt_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
        match to_number(&rt_val) {
            Some(n) if matches!(n as i32, 1 | 2) => n as i32,
            _ => {
                return Ok(CellValue::Error(
                    "WEEKNUM return_type must be 1 or 2".to_string(),
                ));
            },
        }
    } else {
        1
    };
    match weeknum_value(date, return_type) {
        Some(week) => Ok(CellValue::Int(week)),
        None => Ok(CellValue::Error(
            "WEEKNUM return_type not supported".to_string(),
        )),
    }
}

pub(crate) async fn eval_isoweeknum(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error(
            "ISOWEEKNUM expects 1 argument".to_string(),
        ));
    }
    let value = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let date = match coerce_date_value(&value) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "ISOWEEKNUM expects a valid date serial or text".to_string(),
            ));
        },
    };
    Ok(CellValue::Int(date.iso_week().week() as i64))
}
