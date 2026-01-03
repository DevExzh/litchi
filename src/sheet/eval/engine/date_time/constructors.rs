use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use chrono::Timelike;

use super::helpers::{
    SECONDS_PER_DAY, date_to_excel_serial_1900, make_date_serial_1900, number_arg,
    parse_date_string, parse_time_string,
};

pub(crate) async fn eval_date(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "DATE expects 3 arguments (year, month, day)".to_string(),
        ));
    }

    let y = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("DATE year is not numeric".to_string()));
        },
    };
    let m = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("DATE month is not numeric".to_string()));
        },
    };
    let d = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("DATE day is not numeric".to_string()));
        },
    };

    let serial = match make_date_serial_1900(y, m, d) {
        Some(s) => s,
        None => {
            return Ok(CellValue::Error("DATE arguments out of range".to_string()));
        },
    };

    Ok(CellValue::DateTime(serial))
}

pub(crate) async fn eval_time(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "TIME expects 3 arguments (hour, minute, second)".to_string(),
        ));
    }

    let h = match number_arg(ctx, current_sheet, &args[0]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("TIME hour is not numeric".to_string()));
        },
    };
    let m = match number_arg(ctx, current_sheet, &args[1]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("TIME minute is not numeric".to_string()));
        },
    };
    let s = match number_arg(ctx, current_sheet, &args[2]).await? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("TIME second is not numeric".to_string()));
        },
    };

    let (hour, min, sec) = (h.trunc() as i64, m.trunc() as i64, s.trunc() as i64);

    if !(0..24).contains(&hour) || !(0..60).contains(&min) || !(0..60).contains(&sec) {
        return Ok(CellValue::Error("TIME arguments out of range".to_string()));
    }

    let total_seconds = (hour * 3600 + min * 60 + sec) as f64;
    let fraction = total_seconds / SECONDS_PER_DAY;
    Ok(CellValue::DateTime(fraction))
}

pub(crate) async fn eval_datevalue(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("DATEVALUE expects 1 argument".to_string()));
    }

    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;

    if let Some(n) = to_number(&v) {
        // Numeric input: return date part of serial.
        return Ok(CellValue::DateTime(n.floor()));
    }

    let s = match v {
        CellValue::String(ref s) => s.trim(),
        _ => {
            return Ok(CellValue::Error(
                "DATEVALUE expects a date text or serial number".to_string(),
            ));
        },
    };

    let date = match parse_date_string(s) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "DATEVALUE: unsupported date format".to_string(),
            ));
        },
    };
    let serial = date_to_excel_serial_1900(date);
    Ok(CellValue::DateTime(serial))
}

pub(crate) async fn eval_timevalue(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("TIMEVALUE expects 1 argument".to_string()));
    }

    let v = evaluate_expression(ctx, current_sheet, &args[0]).await?;

    if let Some(n) = to_number(&v) {
        // Numeric input: use fractional part as time-of-day.
        let frac = n.fract();
        return Ok(CellValue::DateTime(if frac < 0.0 { 0.0 } else { frac }));
    }

    let s = match v {
        CellValue::String(ref s) => s.trim(),
        _ => {
            return Ok(CellValue::Error(
                "TIMEVALUE expects a time text or serial number".to_string(),
            ));
        },
    };

    let time = match parse_time_string(s) {
        Some(t) => t,
        None => {
            return Ok(CellValue::Error(
                "TIMEVALUE: unsupported time format".to_string(),
            ));
        },
    };

    let total_seconds = time.num_seconds_from_midnight() as f64;
    let fraction = total_seconds / SECONDS_PER_DAY;
    Ok(CellValue::DateTime(fraction))
}
