use crate::sheet::eval::engine::EvalCtx;
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use super::helpers::{
    add_months, date_to_excel_serial_1900, last_day_of_month, number_arg, serial_to_excel_date_1900,
};

pub(crate) async fn eval_edate(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "EDATE expects 2 arguments (start_date, months)".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "EDATE start_date is not numeric".to_string(),
            ));
        },
    };
    let months = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("EDATE months is not numeric".to_string()));
        },
    };

    let date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "EDATE start_date is not a valid date".to_string(),
            ));
        },
    };

    let shifted = match add_months(date, months.trunc() as i32) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error("EDATE result is out of range".to_string()));
        },
    };

    let serial = date_to_excel_serial_1900(shifted);
    Ok(CellValue::DateTime(serial))
}

pub(crate) async fn eval_eomonth(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "EOMONTH expects 2 arguments (start_date, months)".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH start_date is not numeric".to_string(),
            ));
        },
    };
    let months = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH months is not numeric".to_string(),
            ));
        },
    };

    let date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH start_date is not a valid date".to_string(),
            ));
        },
    };

    let shifted = match add_months(date, months.trunc() as i32) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH result is out of range".to_string(),
            ));
        },
    };

    let last_day = match last_day_of_month(shifted) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH result is out of range".to_string(),
            ));
        },
    };

    let serial = date_to_excel_serial_1900(last_day);
    Ok(CellValue::DateTime(serial))
}
