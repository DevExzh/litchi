use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, flatten_range_expr, to_number};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};

use chrono::{Datelike, Duration, NaiveDate, NaiveDateTime, NaiveTime, Timelike, Weekday};

pub(crate) const SECONDS_PER_DAY: f64 = 86_400.0;

pub(super) async fn number_arg(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Option<f64>> {
    let v = evaluate_expression(ctx, current_sheet, expr).await?;
    Ok(to_number(&v))
}

pub(super) fn date_to_excel_serial_1900(date: NaiveDate) -> f64 {
    let base = NaiveDate::from_ymd_opt(1899, 12, 30).expect("Invalid Excel 1900 base date");
    let days = (date - base).num_days();
    days as f64
}

pub(super) fn datetime_to_excel_serial_1900(dt: NaiveDateTime) -> f64 {
    let date_serial = date_to_excel_serial_1900(dt.date());
    let seconds = dt.time().num_seconds_from_midnight() as f64;
    date_serial + seconds / SECONDS_PER_DAY
}

pub(super) fn serial_to_excel_date_1900(serial: f64) -> Option<NaiveDate> {
    let base = NaiveDate::from_ymd_opt(1899, 12, 30)?;
    let days = serial.floor() as i64;
    base.checked_add_signed(Duration::days(days))
}

pub(super) fn make_date_serial_1900(year: f64, month: f64, day: f64) -> Option<f64> {
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

pub(super) fn parse_date_string(s: &str) -> Option<NaiveDate> {
    if let Ok(d) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return Some(d);
    }
    if let Ok(d) = NaiveDate::parse_from_str(s, "%m/%d/%Y") {
        return Some(d);
    }
    None
}

pub(super) fn parse_time_string(s: &str) -> Option<NaiveTime> {
    if let Ok(t) = NaiveTime::parse_from_str(s, "%H:%M:%S") {
        return Some(t);
    }
    if let Ok(t) = NaiveTime::parse_from_str(s, "%H:%M") {
        return Some(t);
    }
    None
}

pub(super) fn parse_datetime_string(s: &str) -> Option<NaiveDateTime> {
    const FORMATS: &[&str] = &[
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
    ];
    for fmt in FORMATS {
        if let Ok(dt) = NaiveDateTime::parse_from_str(s, fmt) {
            return Some(dt);
        }
    }
    None
}

pub(super) fn coerce_date_value(value: &CellValue) -> Option<NaiveDate> {
    if let Some(n) = to_number(value)
        && let Some(date) = serial_to_excel_date_1900(n)
    {
        return Some(date);
    }
    if let CellValue::String(s) = value {
        let trimmed = s.trim();
        if let Some(dt) = parse_datetime_string(trimmed) {
            return Some(dt.date());
        }
        if let Some(d) = parse_date_string(trimmed) {
            return Some(d);
        }
    }
    None
}

pub(super) fn coerce_time_fraction(value: &CellValue) -> Option<f64> {
    if let Some(n) = to_number(value) {
        let frac = n.fract();
        return Some(if frac < 0.0 { (frac + 1.0) % 1.0 } else { frac });
    }
    if let CellValue::String(s) = value {
        let trimmed = s.trim();
        if let Some(dt) = parse_datetime_string(trimmed) {
            return Some(dt.time().num_seconds_from_midnight() as f64 / SECONDS_PER_DAY);
        }
        if let Some(t) = parse_time_string(trimmed) {
            return Some(t.num_seconds_from_midnight() as f64 / SECONDS_PER_DAY);
        }
    }
    None
}

pub(super) fn weekday_number(weekday: Weekday, return_type: i32) -> Option<i64> {
    let num = match return_type {
        1 => weekday.num_days_from_sunday() as i64 + 1,
        2 => weekday.num_days_from_monday() as i64 + 1,
        3 => weekday.num_days_from_monday() as i64,
        _ => return None,
    };
    Some(num)
}

pub(super) fn weeknum_value(date: NaiveDate, return_type: i32) -> Option<i64> {
    let start_weekday = match return_type {
        1 => Weekday::Sun,
        2 => Weekday::Mon,
        _ => return None,
    };
    let first_day = NaiveDate::from_ymd_opt(date.year(), 1, 1)?;
    let mut week_start = first_day;
    while week_start.weekday() != start_weekday {
        week_start = week_start.checked_sub_signed(Duration::days(1))?;
    }
    let days = (date - week_start).num_days();
    Some(days / 7 + 1)
}

pub(super) fn last_day_of_month(date: NaiveDate) -> Option<NaiveDate> {
    let year = date.year();
    let month = date.month();
    let first_next_month = if month == 12 {
        NaiveDate::from_ymd_opt(year + 1, 1, 1)?
    } else {
        NaiveDate::from_ymd_opt(year, month + 1, 1)?
    };
    first_next_month.checked_sub_signed(Duration::days(1))
}

pub(super) fn add_months(date: NaiveDate, months: i32) -> Option<NaiveDate> {
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

pub(super) async fn collect_holiday_dates(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    expr: &Expr,
) -> Result<Vec<NaiveDate>> {
    let range = flatten_range_expr(ctx, current_sheet, expr).await?;
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

pub(super) fn is_business_day(
    date: NaiveDate,
    holidays: &[NaiveDate],
    weekend: &WeekendConfig,
) -> bool {
    if weekend.is_weekend(date.weekday()) {
        return false;
    }
    !holidays.contains(&date)
}

#[derive(Debug, Clone)]
pub(super) enum WeekendConfig {
    Code(i32),
    Pattern(String),
}

impl Default for WeekendConfig {
    fn default() -> Self {
        WeekendConfig::Code(1)
    }
}

impl WeekendConfig {
    pub(super) fn is_weekend(&self, weekday: Weekday) -> bool {
        match self {
            WeekendConfig::Code(c) => match c {
                1 => matches!(weekday, Weekday::Sat | Weekday::Sun),
                2 => matches!(weekday, Weekday::Sun | Weekday::Mon),
                3 => matches!(weekday, Weekday::Mon | Weekday::Tue),
                4 => matches!(weekday, Weekday::Tue | Weekday::Wed),
                5 => matches!(weekday, Weekday::Wed | Weekday::Thu),
                6 => matches!(weekday, Weekday::Thu | Weekday::Fri),
                7 => matches!(weekday, Weekday::Fri | Weekday::Sat),
                11 => matches!(weekday, Weekday::Sun),
                12 => matches!(weekday, Weekday::Mon),
                13 => matches!(weekday, Weekday::Tue),
                14 => matches!(weekday, Weekday::Wed),
                15 => matches!(weekday, Weekday::Thu),
                16 => matches!(weekday, Weekday::Fri),
                17 => matches!(weekday, Weekday::Sat),
                _ => matches!(weekday, Weekday::Sat | Weekday::Sun), // Default
            },
            WeekendConfig::Pattern(p) => {
                if p.len() != 7 {
                    return matches!(weekday, Weekday::Sat | Weekday::Sun);
                }
                let idx = match weekday {
                    Weekday::Mon => 0,
                    Weekday::Tue => 1,
                    Weekday::Wed => 2,
                    Weekday::Thu => 3,
                    Weekday::Fri => 4,
                    Weekday::Sat => 5,
                    Weekday::Sun => 6,
                };
                p.as_bytes().get(idx) == Some(&b'1')
            },
        }
    }
}

pub(super) fn workday_core(
    start: NaiveDate,
    days: i64,
    holidays: &[NaiveDate],
    weekend: &WeekendConfig,
) -> Option<NaiveDate> {
    if days == 0 {
        return Some(start);
    }

    let step = if days > 0 { 1 } else { -1 };
    let mut remaining = days.abs();
    let mut date = start;

    while remaining > 0 {
        date = date.checked_add_signed(Duration::days(step))?;
        if is_business_day(date, holidays, weekend) {
            remaining -= 1;
        }
    }

    Some(date)
}

pub(super) fn networkdays_core(
    start: NaiveDate,
    end: NaiveDate,
    holidays: &[NaiveDate],
    weekend: &WeekendConfig,
) -> i64 {
    let (mut from, to, sign) = if start <= end {
        (start, end, 1)
    } else {
        (end, start, -1)
    };

    let mut count = 0i64;
    while from <= to {
        if is_business_day(from, holidays, weekend) {
            count += 1;
        }
        from = match from.checked_add_signed(Duration::days(1)) {
            Some(d) => d,
            None => break,
        };
    }

    count * sign
}
