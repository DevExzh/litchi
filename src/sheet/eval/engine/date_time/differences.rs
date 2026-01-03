use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use chrono::{Datelike, NaiveDate};

use super::helpers::{add_months, coerce_date_value, last_day_of_month};

pub(crate) async fn eval_days(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "DAYS expects 2 arguments (end_date, start_date)".to_string(),
        ));
    }

    let end_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let start_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;

    let end_date = match coerce_date_value(&end_val) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "DAYS end_date must be a valid date serial or text".to_string(),
            ));
        },
    };
    let start_date = match coerce_date_value(&start_val) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "DAYS start_date must be a valid date serial or text".to_string(),
            ));
        },
    };

    let diff = (end_date - start_date).num_days();
    Ok(CellValue::Float(diff as f64))
}

pub(crate) async fn eval_days360(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "DAYS360 expects 2 or 3 arguments (start_date, end_date, [method])".to_string(),
        ));
    }

    let start_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let end_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;

    let start_date = match coerce_date_value(&start_val) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "DAYS360 start_date must be a valid date serial or text".to_string(),
            ));
        },
    };
    let end_date = match coerce_date_value(&end_val) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "DAYS360 end_date must be a valid date serial or text".to_string(),
            ));
        },
    };

    let european_method = if args.len() == 3 {
        let method_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match method_val {
            CellValue::Bool(b) => b,
            _ => match to_number(&method_val) {
                Some(n) => n != 0.0,
                None => {
                    return Ok(CellValue::Error(
                        "DAYS360 method must be a boolean or numeric value".to_string(),
                    ));
                },
            },
        }
    } else {
        false
    };

    let days = if european_method {
        days360_european(start_date, end_date)
    } else {
        days360_us(start_date, end_date)
    };

    Ok(CellValue::Float(days as f64))
}

pub(crate) async fn eval_datedif(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error(
            "DATEDIF expects 3 arguments (start_date, end_date, unit)".to_string(),
        ));
    }

    let start_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let end_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let unit_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;

    let mut start_date = match coerce_date_value(&start_val) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "DATEDIF start_date must be a valid date serial or text".to_string(),
            ));
        },
    };
    let mut end_date = match coerce_date_value(&end_val) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "DATEDIF end_date must be a valid date serial or text".to_string(),
            ));
        },
    };

    if end_date < start_date {
        std::mem::swap(&mut start_date, &mut end_date);
    }

    let unit = match unit_val {
        CellValue::String(ref s) if !s.trim().is_empty() => s.trim().to_uppercase(),
        _ => {
            return Ok(CellValue::Error(
                "DATEDIF unit must be a non-empty text value".to_string(),
            ));
        },
    };

    let value = match unit.as_str() {
        "Y" => match difference_in_complete_years(start_date, end_date) {
            Some(v) => v as f64,
            None => {
                return Ok(CellValue::Error(
                    "DATEDIF encountered an invalid date when computing years".to_string(),
                ));
            },
        },
        "M" => match difference_in_complete_months(start_date, end_date) {
            Some(v) => v as f64,
            None => {
                return Ok(CellValue::Error(
                    "DATEDIF encountered an invalid date when computing months".to_string(),
                ));
            },
        },
        "D" => (end_date - start_date).num_days() as f64,
        "MD" => match difference_md(start_date, end_date) {
            Some(v) => v as f64,
            None => {
                return Ok(CellValue::Error(
                    "DATEDIF encountered an invalid date when computing MD".to_string(),
                ));
            },
        },
        "YM" => match difference_ym(start_date, end_date) {
            Some(v) => v as f64,
            None => {
                return Ok(CellValue::Error(
                    "DATEDIF encountered an invalid date when computing YM".to_string(),
                ));
            },
        },
        "YD" => match difference_yd(start_date, end_date) {
            Some(v) => v as f64,
            None => {
                return Ok(CellValue::Error(
                    "DATEDIF encountered an invalid date when computing YD".to_string(),
                ));
            },
        },
        _ => {
            return Ok(CellValue::Error(
                "DATEDIF unit must be one of \"Y\",\"M\",\"D\",\"MD\",\"YM\",\"YD\"".to_string(),
            ));
        },
    };

    Ok(CellValue::Float(value))
}

fn days360_us(start: chrono::NaiveDate, end: chrono::NaiveDate) -> i64 {
    let sy = start.year();
    let sm = start.month() as i32;
    let mut sd = adjust_day_us(start);

    let mut ey = end.year();
    let mut em = end.month() as i32;
    let mut ed = adjust_day_us(end);

    let start_is_eom = is_last_day_of_month(start);
    let end_is_eom = is_last_day_of_month(end);

    if start_is_eom {
        sd = 30;
    }

    if end_is_eom {
        if start_is_eom {
            ed = 30;
        } else if sd < 30 {
            ed = 1;
            em += 1;
            if em > 12 {
                em = 1;
                ey += 1;
            }
        } else {
            ed = 30;
        }
    }

    ((ey - sy) * 360 + (em - sm) * 30 + (ed - sd)) as i64
}

fn days360_european(start: chrono::NaiveDate, end: chrono::NaiveDate) -> i64 {
    let mut sd = start.day() as i32;
    let mut ed = end.day() as i32;

    if sd == 31 {
        sd = 30;
    }
    if ed == 31 {
        ed = 30;
    }

    ((end.year() - start.year()) * 360
        + (end.month() as i32 - start.month() as i32) * 30
        + (ed - sd)) as i64
}

fn adjust_day_us(date: chrono::NaiveDate) -> i32 {
    let day = date.day() as i32;
    if is_last_day_of_month(date) { 30 } else { day }
}

fn is_last_day_of_month(date: chrono::NaiveDate) -> bool {
    if let Some(last) = last_day_of_month(date) {
        last.day() == date.day()
    } else {
        false
    }
}

fn difference_in_complete_years(start: NaiveDate, end: NaiveDate) -> Option<i64> {
    let mut years = end.year() - start.year();
    if years <= 0 {
        return Some(0);
    }
    let total_months = years * 12;
    if let Some(anniversary) = add_months(start, total_months)
        && end < anniversary
    {
        years -= 1;
    }
    Some(years as i64)
}

fn difference_in_complete_months(start: NaiveDate, end: NaiveDate) -> Option<i64> {
    let mut months = (end.year() - start.year()) * 12 + (end.month() as i32 - start.month() as i32);
    if months <= 0 {
        return Some(0);
    }
    if let Some(candidate) = add_months(start, months)
        && end < candidate
    {
        months -= 1;
    }
    Some(months as i64)
}

fn difference_md(start: NaiveDate, end: NaiveDate) -> Option<i64> {
    if end < start {
        return Some(0);
    }

    let sd = start.day() as i32;
    let mut ed = end.day() as i32;

    if ed >= sd {
        return Some((ed - sd) as i64);
    }

    let prev_month_first = if end.month() == 1 {
        NaiveDate::from_ymd_opt(end.year() - 1, 12, 1)?
    } else {
        NaiveDate::from_ymd_opt(end.year(), end.month() - 1, 1)?
    };
    let prev_month_last = last_day_of_month(prev_month_first)?;
    ed += prev_month_last.day() as i32;

    Some((ed - sd) as i64)
}

fn difference_ym(start: NaiveDate, end: NaiveDate) -> Option<i64> {
    if end < start {
        return Some(0);
    }

    let total_months =
        (end.year() - start.year()) * 12 + (end.month() as i32 - start.month() as i32);
    let mut months = total_months.rem_euclid(12);

    if end.day() < start.day() {
        months = (months - 1 + 12) % 12;
    }

    Some(months as i64)
}

fn difference_yd(start: NaiveDate, end: NaiveDate) -> Option<i64> {
    if end < start {
        return Some(0);
    }

    let mut end_same_year = adjust_year(end, start.year())?;
    if end_same_year < start {
        end_same_year = adjust_year(end, start.year() + 1)?;
    }

    Some((end_same_year - start).num_days())
}

fn adjust_year(date: NaiveDate, year: i32) -> Option<NaiveDate> {
    NaiveDate::from_ymd_opt(year, date.month(), date.day()).or_else(|| {
        let first_of_month = NaiveDate::from_ymd_opt(year, date.month(), 1)?;
        last_day_of_month(first_of_month)
    })
}

pub(crate) async fn eval_yearfrac(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(CellValue::Error(
            "YEARFRAC expects 2 or 3 arguments (start_date, end_date, [basis])".to_string(),
        ));
    }

    let start_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let end_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;

    let start_date = match coerce_date_value(&start_val) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "YEARFRAC start_date must be a valid date serial or text".to_string(),
            ));
        },
    };
    let end_date = match coerce_date_value(&end_val) {
        Some(d) => d,
        None => {
            return Ok(CellValue::Error(
                "YEARFRAC end_date must be a valid date serial or text".to_string(),
            ));
        },
    };

    let basis = if args.len() == 3 {
        let basis_val = evaluate_expression(ctx, current_sheet, &args[2]).await?;
        match to_number(&basis_val) {
            Some(n) => n.trunc() as i32,
            None => {
                return Ok(CellValue::Error(
                    "YEARFRAC basis must be a numeric value".to_string(),
                ));
            },
        }
    } else {
        0
    };

    if !(0..=4).contains(&basis) {
        return Ok(CellValue::Error("#NUM!".to_string()));
    }

    let (d1, d2) = if start_date <= end_date {
        (start_date, end_date)
    } else {
        (end_date, start_date)
    };

    let frac = match basis {
        0 => {
            // US (NASD) 30/360
            days360_us(d1, d2) as f64 / 360.0
        },
        1 => {
            // Actual/actual
            let diff_days = (d2 - d1).num_days() as f64;
            let year1 = d1.year();
            let year2 = d2.year();
            let mut days_in_year = 0.0;
            if year1 == year2 {
                days_in_year = if is_leap_year(year1) { 366.0 } else { 365.0 };
            } else {
                let n_years = (year2 - year1 + 1) as f64;
                for y in year1..=year2 {
                    days_in_year += if is_leap_year(y) { 366.0 } else { 365.0 };
                }
                days_in_year /= n_years;
            }
            diff_days / days_in_year
        },
        2 => {
            // Actual/360
            (d2 - d1).num_days() as f64 / 360.0
        },
        3 => {
            // Actual/365
            (d2 - d1).num_days() as f64 / 365.0
        },
        4 => {
            // European 30/360
            days360_european(d1, d2) as f64 / 360.0
        },
        _ => unreachable!(),
    };

    Ok(CellValue::Float(frac))
}

fn is_leap_year(year: i32) -> bool {
    year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
}
