use crate::sheet::{CellValue, Result};

use super::super::parser::Expr;
use super::{EngineCtx, evaluate_expression, flatten_range_expr, to_number};

use chrono::{Datelike, Duration, NaiveDate, NaiveDateTime, NaiveTime, Timelike, Utc, Weekday};

const SECONDS_PER_DAY: f64 = 86_400.0;

pub(crate) fn eval_today<C: EngineCtx + ?Sized>(
    _ctx: &C,
    _current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("TODAY expects 0 arguments".to_string()));
    }

    let now = Utc::now().date_naive();
    let serial = date_to_excel_serial_1900(now);
    Ok(CellValue::DateTime(serial))
}

pub(crate) fn eval_now<C: EngineCtx + ?Sized>(
    _ctx: &C,
    _current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if !args.is_empty() {
        return Ok(CellValue::Error("NOW expects 0 arguments".to_string()));
    }

    let now = Utc::now().naive_utc();
    let serial = datetime_to_excel_serial_1900(now);
    Ok(CellValue::DateTime(serial))
}

pub(crate) fn eval_date<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "DATE expects 3 arguments (year, month, day)".to_string(),
        ));
    }

    let y = match number_arg(ctx, current_sheet, &args[0])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("DATE year is not numeric".to_string()));
        },
    };
    let m = match number_arg(ctx, current_sheet, &args[1])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("DATE month is not numeric".to_string()));
        },
    };
    let d = match number_arg(ctx, current_sheet, &args[2])? {
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

pub(crate) fn eval_time<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "TIME expects 3 arguments (hour, minute, second)".to_string(),
        ));
    }

    let h = match number_arg(ctx, current_sheet, &args[0])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("TIME hour is not numeric".to_string()));
        },
    };
    let m = match number_arg(ctx, current_sheet, &args[1])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("TIME minute is not numeric".to_string()));
        },
    };
    let s = match number_arg(ctx, current_sheet, &args[2])? {
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

pub(crate) fn eval_datevalue<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("DATEVALUE expects 1 argument".to_string()));
    }

    let v = evaluate_expression(ctx, current_sheet, &args[0])?;

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

pub(crate) fn eval_timevalue<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("TIMEVALUE expects 1 argument".to_string()));
    }

    let v = evaluate_expression(ctx, current_sheet, &args[0])?;

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

pub(crate) fn eval_edate<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "EDATE expects 2 arguments (start_date, months)".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "EDATE start_date is not numeric".to_string(),
            ));
        },
    };
    let months = match number_arg(ctx, current_sheet, &args[1])? {
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

pub(crate) fn eval_eomonth<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "EOMONTH expects 2 arguments (start_date, months)".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "EOMONTH start_date is not numeric".to_string(),
            ));
        },
    };
    let months = match number_arg(ctx, current_sheet, &args[1])? {
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

pub(crate) fn eval_workday<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "WORKDAY expects 2 or 3 arguments (start_date, days, [holidays])".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY start_date is not numeric".to_string(),
            ));
        },
    };
    let days = match number_arg(ctx, current_sheet, &args[1])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error("WORKDAY days is not numeric".to_string()));
        },
    };

    let start_date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY start_date is not a valid date".to_string(),
            ));
        },
    };

    let holidays = if args.len() == 3 {
        collect_holiday_dates(ctx, current_sheet, &args[2])?
    } else {
        Vec::new()
    };

    let result = match workday_core(start_date, days.trunc() as i64, &holidays) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY result is out of range".to_string(),
            ));
        },
    };

    let serial = date_to_excel_serial_1900(result);
    Ok(CellValue::DateTime(serial))
}

pub(crate) fn eval_workday_intl<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 4 {
        return Ok(CellValue::Error(
            "WORKDAY.INTL expects 2 to 4 arguments (start_date, days, [weekend], [holidays])"
                .to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY.INTL start_date is not numeric".to_string(),
            ));
        },
    };
    let days = match number_arg(ctx, current_sheet, &args[1])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY.INTL days is not numeric".to_string(),
            ));
        },
    };

    // For now, only support default weekend (Saturday/Sunday).
    if args.len() >= 3 {
        let weekend_val = evaluate_expression(ctx, current_sheet, &args[2])?;
        if let Some(n) = to_number(&weekend_val) {
            if n != 1.0 {
                return Ok(CellValue::Error(
                    "WORKDAY.INTL currently only supports weekend code 1 (Saturday/Sunday)"
                        .to_string(),
                ));
            }
        } else if let CellValue::String(ref s) = weekend_val
            && !s.is_empty()
            && s != "0000011"
        {
            return Ok(CellValue::Error(
                "WORKDAY.INTL currently only supports default weekend pattern".to_string(),
            ));
        }
    }

    let start_date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY.INTL start_date is not a valid date".to_string(),
            ));
        },
    };

    let holidays = if args.len() == 4 {
        collect_holiday_dates(ctx, current_sheet, &args[3])?
    } else if args.len() == 3 {
        // Third arg used for weekend, no holidays.
        Vec::new()
    } else {
        Vec::new()
    };

    let result = match workday_core(start_date, days.trunc() as i64, &holidays) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "WORKDAY.INTL result is out of range".to_string(),
            ));
        },
    };

    let serial = date_to_excel_serial_1900(result);
    Ok(CellValue::DateTime(serial))
}

pub(crate) fn eval_networkdays<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "NETWORKDAYS expects 2 or 3 arguments (start_date, end_date, [holidays])".to_string(),
        ));
    }

    let start = match number_arg(ctx, current_sheet, &args[0])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS start_date is not numeric".to_string(),
            ));
        },
    };
    let end = match number_arg(ctx, current_sheet, &args[1])? {
        Some(n) => n,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS end_date is not numeric".to_string(),
            ));
        },
    };

    let start_date = match serial_to_excel_date_1900(start) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS start_date is not a valid date".to_string(),
            ));
        },
    };
    let end_date = match serial_to_excel_date_1900(end) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "NETWORKDAYS end_date is not a valid date".to_string(),
            ));
        },
    };

    let holidays = if args.len() == 3 {
        collect_holiday_dates(ctx, current_sheet, &args[2])?
    } else {
        Vec::new()
    };

    let count = networkdays_core(start_date, end_date, &holidays);
    Ok(CellValue::Int(count))
}

fn number_arg<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Option<f64>> {
    let v = evaluate_expression(ctx, current_sheet, expr)?;
    Ok(to_number(&v))
}

fn date_to_excel_serial_1900(date: NaiveDate) -> f64 {
    let base = NaiveDate::from_ymd_opt(1899, 12, 30).expect("Invalid Excel 1900 base date");
    let days = (date - base).num_days();
    days as f64
}

fn datetime_to_excel_serial_1900(dt: NaiveDateTime) -> f64 {
    let date_serial = date_to_excel_serial_1900(dt.date());
    let seconds = dt.time().num_seconds_from_midnight() as f64;
    date_serial + seconds / SECONDS_PER_DAY
}

fn serial_to_excel_date_1900(serial: f64) -> Option<NaiveDate> {
    let base = NaiveDate::from_ymd_opt(1899, 12, 30)?;
    let days = serial.floor() as i64;
    base.checked_add_signed(Duration::days(days))
}

fn make_date_serial_1900(year: f64, month: f64, day: f64) -> Option<f64> {
    let y = year.trunc() as i32;
    let m = month.trunc() as i32;
    let d = day.trunc() as i64;

    // Excel-style month normalization: months can be outside 1..=12.
    let mut ym = (y, m - 1); // month 0-based
    ym.0 += ym.1.div_euclid(12);
    ym.1 = ym.1.rem_euclid(12);
    let norm_year = ym.0;
    let norm_month = (ym.1 + 1) as u32;

    let first_of_month = NaiveDate::from_ymd_opt(norm_year, norm_month, 1)?;
    let target = first_of_month.checked_add_signed(Duration::days(d - 1))?;
    Some(date_to_excel_serial_1900(target))
}

fn parse_date_string(s: &str) -> Option<NaiveDate> {
    if let Ok(d) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return Some(d);
    }
    if let Ok(d) = NaiveDate::parse_from_str(s, "%m/%d/%Y") {
        return Some(d);
    }
    None
}

fn parse_time_string(s: &str) -> Option<NaiveTime> {
    if let Ok(t) = NaiveTime::parse_from_str(s, "%H:%M:%S") {
        return Some(t);
    }
    if let Ok(t) = NaiveTime::parse_from_str(s, "%H:%M") {
        return Some(t);
    }
    None
}

fn last_day_of_month(date: NaiveDate) -> Option<NaiveDate> {
    let year = date.year();
    let month = date.month();
    let first_next_month = if month == 12 {
        NaiveDate::from_ymd_opt(year + 1, 1, 1)?
    } else {
        NaiveDate::from_ymd_opt(year, month + 1, 1)?
    };
    first_next_month.checked_sub_signed(Duration::days(1))
}

fn add_months(date: NaiveDate, months: i32) -> Option<NaiveDate> {
    let mut year = date.year();
    let mut month_index = date.month0() as i32 + months; // 0-based

    year += month_index.div_euclid(12);
    month_index = month_index.rem_euclid(12);
    let month = (month_index + 1) as u32;

    let day = date.day();
    let mut result = NaiveDate::from_ymd_opt(year, month, day);
    if result.is_none() {
        // Clamp to last valid day of target month.
        let last = last_day_of_month(NaiveDate::from_ymd_opt(year, month, 1)?)?;
        result = Some(last);
    }
    result
}

fn is_business_day(date: NaiveDate, holidays: &[NaiveDate]) -> bool {
    match date.weekday() {
        Weekday::Sat | Weekday::Sun => false,
        _ => !holidays.contains(&date),
    }
}

fn collect_holiday_dates<C: EngineCtx + ?Sized>(
    ctx: &C,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Vec<NaiveDate>> {
    let range = flatten_range_expr(ctx, current_sheet, expr)?;
    let mut out = Vec::new();
    for v in &range.values {
        if let Some(n) = to_number(v)
            && let Some(d) = serial_to_excel_date_1900(n)
        {
            out.push(d);
        }
    }
    Ok(out)
}

fn workday_core(start: NaiveDate, days: i64, holidays: &[NaiveDate]) -> Option<NaiveDate> {
    if days == 0 {
        return Some(start);
    }

    let step = if days > 0 { 1 } else { -1 };
    let mut remaining = days.abs();
    let mut date = start;

    while remaining > 0 {
        date = date.checked_add_signed(Duration::days(step))?;
        if is_business_day(date, holidays) {
            remaining -= 1;
        }
    }

    Some(date)
}

fn networkdays_core(start: NaiveDate, end: NaiveDate, holidays: &[NaiveDate]) -> i64 {
    let (mut from, to, sign) = if start <= end {
        (start, end, 1)
    } else {
        (end, start, -1)
    };

    let mut count = 0i64;
    while from <= to {
        if is_business_day(from, holidays) {
            count += 1;
        }
        from = match from.checked_add_signed(Duration::days(1)) {
            Some(d) => d,
            None => break,
        };
    }

    count * sign
}
