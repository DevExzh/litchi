//! Pure, bounded civil-date kernels shared by the scalar date/time façade and
//! the resolver-backed value evaluator.
//!
//! The civil representation and Gregorian conversion live in the shared
//! evaluation::calendar foundation. This module owns date/time-family
//! semantics, so the value VM and scalar VM cannot acquire different epochs.

use super::super::ScalarError;
pub(in super::super) use super::super::calendar::CivilDate;
use super::super::calendar::{
    self, MAX_DATE_SERIAL_EXCLUSIVE as CALENDAR_MAX_DATETIME_SERIAL,
    MIN_DATE_SERIAL as CALENDAR_MIN_DATE_SERIAL, SERIAL_EPOCH_TO_UNIX_DAYS,
};

pub(in super::super) const MIN_SERIAL: f64 = CALENDAR_MIN_DATE_SERIAL as f64;
pub(in super::super) const MAX_DATETIME_SERIAL: f64 = CALENDAR_MAX_DATETIME_SERIAL as f64;

#[derive(Clone, Copy, Debug, PartialEq)]
pub(in super::super) struct DateTimeParts {
    pub(in super::super) date: CivilDate,
    pub(in super::super) fraction: f64,
}

#[inline]
pub(in super::super) const fn is_leap_year(year: i32) -> bool {
    calendar::is_leap_year(year)
}

#[inline]
pub(in super::super) fn days_in_month(year: i32, month: u32) -> u32 {
    calendar::days_in_month(year, month).unwrap_or(0)
}

#[inline]
fn date_key(date: CivilDate) -> (i32, u32, u32) {
    (date.year, date.month, date.day)
}

#[inline]
pub(in super::super) fn date_lt(left: CivilDate, right: CivilDate) -> bool {
    date_key(left) < date_key(right)
}

#[inline]
pub(in super::super) fn date_le(left: CivilDate, right: CivilDate) -> bool {
    date_key(left) <= date_key(right)
}

pub(in super::super) fn days_from_civil(date: CivilDate) -> Result<i64, ScalarError> {
    calendar::days_from_civil(date).ok_or(ScalarError::Number)
}

pub(in super::super) fn date_to_serial(date: CivilDate) -> Result<f64, ScalarError> {
    let serial = calendar::serial_from_civil(date).ok_or(ScalarError::Number)?;
    Ok(serial as f64)
}

pub(in super::super) fn serial_to_parts(serial: f64) -> Result<DateTimeParts, ScalarError> {
    let (date, fraction) = calendar::civil_from_serial(serial).ok_or(ScalarError::Number)?;
    Ok(DateTimeParts { date, fraction })
}

pub(in super::super) fn serial_to_date(serial: f64) -> Result<CivilDate, ScalarError> {
    Ok(serial_to_parts(serial)?.date)
}

/// Normalize positive DATE month/day values with checked Gregorian rollover.
pub(in super::super) fn normalize_ymd(
    year: i32,
    month: i64,
    day: i64,
) -> Result<CivilDate, ScalarError> {
    if month <= 0 || day <= 0 {
        return Err(ScalarError::Number);
    }
    let month_zero = month - 1;
    let year_delta = month_zero.div_euclid(12);
    let normalized_month = month_zero.rem_euclid(12) + 1;
    let normalized_year = i64::from(year)
        .checked_add(year_delta)
        .ok_or(ScalarError::Number)?;
    let normalized_year = i32::try_from(normalized_year).map_err(|_| ScalarError::Number)?;
    let first =
        CivilDate::new(normalized_year, normalized_month as u32, 1).ok_or(ScalarError::Number)?;
    let first_days = days_from_civil(first)?;
    let target_days = first_days.checked_add(day - 1).ok_or(ScalarError::Number)?;
    calendar::civil_from_days_checked(target_days).ok_or(ScalarError::Number)
}

pub(in super::super) fn add_months(date: CivilDate, months: i64) -> Result<CivilDate, ScalarError> {
    let month_index = i64::from(date.year)
        .checked_mul(12)
        .and_then(|value| value.checked_add(i64::from(date.month) - 1))
        .and_then(|value| value.checked_add(months))
        .ok_or(ScalarError::Number)?;
    let year = month_index.div_euclid(12);
    let month = month_index.rem_euclid(12) + 1;
    let year = i32::try_from(year).map_err(|_| ScalarError::Number)?;
    let day = date.day.min(days_in_month(year, month as u32));
    CivilDate::new(year, month as u32, day).ok_or(ScalarError::Number)
}

pub(in super::super) fn weekday_sunday_zero(date: CivilDate) -> Result<u32, ScalarError> {
    let days = days_from_civil(date)?
        .checked_add(SERIAL_EPOCH_TO_UNIX_DAYS)
        .ok_or(ScalarError::Number)?;
    // 1899-12-30 was a Saturday (6 when Sunday is zero).
    Ok(days
        .checked_add(6)
        .ok_or(ScalarError::Number)?
        .rem_euclid(7) as u32)
}

pub(in super::super) fn iso_week(date: CivilDate) -> Result<(i32, u32), ScalarError> {
    let weekday = weekday_sunday_zero(date)?;
    let monday_zero = (weekday + 6) % 7;
    let thursday_offset = 3_i64 - i64::from(monday_zero);
    let thursday = days_from_civil(date)?
        .checked_add(thursday_offset)
        .ok_or(ScalarError::Number)?;
    let week_year = calendar::civil_from_days_checked(thursday)
        .ok_or(ScalarError::Number)?
        .year;
    let jan_four_date = CivilDate::new(week_year, 1, 4).ok_or(ScalarError::Number)?;
    let jan_four = days_from_civil(jan_four_date)?;
    let jan_four_weekday = weekday_sunday_zero(jan_four_date)?;
    let jan_four_monday = (jan_four_weekday + 6) % 7;
    let week_one = jan_four
        .checked_sub(i64::from(jan_four_monday))
        .ok_or(ScalarError::Number)?;
    let current = days_from_civil(date)?;
    let week = current
        .checked_sub(week_one)
        .and_then(|value| value.checked_div(7))
        .and_then(|value| value.checked_add(1))
        .ok_or(ScalarError::Number)?;
    Ok((
        week_year,
        u32::try_from(week).map_err(|_| ScalarError::Number)?,
    ))
}

pub(in super::super) fn easter_sunday(year: i32) -> Result<CivilDate, ScalarError> {
    if !(1583..=9956).contains(&year) {
        return Err(ScalarError::Number);
    }
    let a = year % 19;
    let b = year / 100;
    let c = year % 100;
    let d = b / 4;
    let e = b % 4;
    let f = (b + 8) / 25;
    let g = (b - f + 1) / 3;
    let h = (19 * a + b - d - g + 15) % 30;
    let i = c / 4;
    let k = c % 4;
    let l = (32 + 2 * e + 2 * i - h - k) % 7;
    let m = (a + 11 * h + 22 * l) / 451;
    let month = (h + l - 7 * m + 114) / 31;
    let day = (h + l - 7 * m + 114) % 31 + 1;
    CivilDate::new(year, month as u32, day as u32).ok_or(ScalarError::Number)
}

pub(in super::super) fn days360_us(start: CivilDate, end: CivilDate) -> Result<i64, ScalarError> {
    let start_last_feb = start.month == 2 && start.day == days_in_month(start.year, 2);
    let start_day = if start.day == 31 || start_last_feb {
        30
    } else {
        start.day
    };
    let end_day = if end.day == 31 && start_day == 30 {
        30
    } else {
        end.day
    };
    let start_value = i64::from(start.year)
        .checked_mul(360)
        .and_then(|value| {
            i64::from(start.month)
                .checked_mul(30)
                .and_then(|month| value.checked_add(month))
        })
        .and_then(|value| value.checked_add(i64::from(start_day)))
        .ok_or(ScalarError::Number)?;
    let end_value = i64::from(end.year)
        .checked_mul(360)
        .and_then(|value| {
            i64::from(end.month)
                .checked_mul(30)
                .and_then(|month| value.checked_add(month))
        })
        .and_then(|value| value.checked_add(i64::from(end_day)))
        .ok_or(ScalarError::Number)?;
    end_value
        .checked_sub(start_value)
        .ok_or(ScalarError::Number)
}

pub(in super::super) fn days360_european(
    start: CivilDate,
    end: CivilDate,
) -> Result<i64, ScalarError> {
    let (start, end, sign) = if date_lt(start, end) || start == end {
        (start, end, 1_i64)
    } else {
        (end, start, -1_i64)
    };
    let start_day = start.day.min(30);
    let end_day = end.day.min(30);
    let end_value = i64::from(end.year)
        .checked_mul(360)
        .and_then(|value| {
            i64::from(end.month)
                .checked_mul(30)
                .and_then(|month| value.checked_add(month))
        })
        .and_then(|value| value.checked_add(i64::from(end_day)))
        .ok_or(ScalarError::Number)?;
    let start_value = i64::from(start.year)
        .checked_mul(360)
        .and_then(|value| {
            i64::from(start.month)
                .checked_mul(30)
                .and_then(|month| value.checked_add(month))
        })
        .and_then(|value| value.checked_add(i64::from(start_day)))
        .ok_or(ScalarError::Number)?;
    end_value
        .checked_sub(start_value)
        .and_then(|value| value.checked_mul(sign))
        .ok_or(ScalarError::Number)
}

/// Basis 0's Procedure A. This is intentionally separate from the `DAYS360`
/// US procedure: YEARFRAC orders reversed dates and applies the two
/// February-end adjustments after the 31st rules, including the
/// counterfactual February 30th endpoint when both dates are February ends.
pub(in super::super) fn days360_basis_a(
    start: CivilDate,
    end: CivilDate,
) -> Result<i64, ScalarError> {
    if start == end {
        return Ok(0);
    }
    let (start, end) = if date_lt(start, end) {
        (start, end)
    } else {
        (end, start)
    };

    let start_last_feb = start.month == 2 && start.day == days_in_month(start.year, 2);
    let end_last_feb = end.month == 2 && end.day == days_in_month(end.year, 2);
    let mut start_day = start.day;
    let mut end_day = end.day;

    if start_day == 31 {
        start_day = 30;
    }
    if start_day == 30 && end_day == 31 {
        end_day = 30;
    }
    if start_last_feb && end_last_feb {
        end_day = 30;
    }
    if start_last_feb {
        start_day = 30;
    }

    let start_value = i64::from(start.year)
        .checked_mul(360)
        .and_then(|value| value.checked_add(i64::from(start.month) * 30))
        .and_then(|value| value.checked_add(i64::from(start_day)))
        .ok_or(ScalarError::Number)?;
    let end_value = i64::from(end.year)
        .checked_mul(360)
        .and_then(|value| value.checked_add(i64::from(end.month) * 30))
        .and_then(|value| value.checked_add(i64::from(end_day)))
        .ok_or(ScalarError::Number)?;
    end_value
        .checked_sub(start_value)
        .ok_or(ScalarError::Number)
}

pub(in super::super) fn date_difference(
    start: CivilDate,
    end: CivilDate,
) -> Result<i64, ScalarError> {
    days_from_civil(end)?
        .checked_sub(days_from_civil(start)?)
        .ok_or(ScalarError::Number)
}

pub(in super::super) fn datedif(
    start: CivilDate,
    end: CivilDate,
    format: &str,
) -> Result<i64, ScalarError> {
    if date_lt(end, start) {
        return Err(ScalarError::Number);
    }
    match format {
        "D" => date_difference(start, end),
        "Y" => {
            let mut years = end.year - start.year;
            let anniversary = anniversary(start, start.year + years)?;
            if date_lt(end, anniversary) {
                years -= 1;
            }
            Ok(i64::from(years))
        },
        "M" => {
            let mut months = (end.year - start.year) * 12 + end.month as i32 - start.month as i32;
            let anniversary = add_months(start, i64::from(months))?;
            if date_lt(end, anniversary) {
                months -= 1;
            }
            Ok(i64::from(months))
        },
        "MD" => {
            let mut day = i64::from(end.day) - i64::from(start.day);
            if day < 0 {
                let previous_month = add_months(
                    CivilDate::new(end.year, end.month, 1).ok_or(ScalarError::Number)?,
                    -1,
                )?;
                day += i64::from(days_in_month(previous_month.year, previous_month.month));
            }
            Ok(day)
        },
        "YM" => {
            let mut month = i64::from(end.month) - i64::from(start.month);
            if month < 0 {
                month += 12;
            }
            if end.day < start.day {
                month = (month + 11) % 12;
            }
            Ok(month)
        },
        "YD" => {
            let mut anniversary = CivilDate::new(
                end.year,
                start.month,
                start.day.min(days_in_month(end.year, start.month)),
            )
            .ok_or(ScalarError::Number)?;
            if date_lt(end, anniversary) {
                anniversary = CivilDate::new(
                    end.year - 1,
                    start.month,
                    start.day.min(days_in_month(end.year - 1, start.month)),
                )
                .ok_or(ScalarError::Number)?;
            }
            date_difference(anniversary, end)
        },
        _ => Err(ScalarError::Value),
    }
}

fn anniversary(start: CivilDate, year: i32) -> Result<CivilDate, ScalarError> {
    CivilDate::new(
        year,
        start.month,
        start.day.min(days_in_month(year, start.month)),
    )
    .ok_or(ScalarError::Number)
}

/// YEARFRAC's actual/actual denominator follows the ODF Procedure E rule.
pub(in super::super) fn yearfrac(
    start: CivilDate,
    end: CivilDate,
    basis: i64,
) -> Result<f64, ScalarError> {
    if !(0..=4).contains(&basis) {
        return Err(ScalarError::Number);
    }
    let (start, end) = if date_le(start, end) {
        (start, end)
    } else {
        (end, start)
    };
    let actual = date_difference(start, end)? as f64;
    let value = match basis {
        0 => days360_basis_a(start, end)? as f64 / 360.0,
        1 => actual_actual(start, end, actual)?,
        2 => actual / 360.0,
        3 => actual / 365.0,
        4 => days360_european(start, end)? as f64 / 360.0,
        _ => return Err(ScalarError::Number),
    };
    if value.is_finite() {
        Ok(value)
    } else {
        Err(ScalarError::Number)
    }
}

fn actual_actual(start: CivilDate, end: CivilDate, actual: f64) -> Result<f64, ScalarError> {
    let y1 = start.year;
    let y2 = end.year;
    if y1 == y2 {
        let denominator = if is_leap_year(y1) { 366.0 } else { 365.0 };
        return Ok(actual / denominator);
    }
    let a = y1 != y2;
    let b = y2 != y1 + 1;
    let c = start.month < end.month;
    let d = start.month == end.month;
    let e = start.day < end.day;
    let f = a && (b || c || (d && e));
    let denominator = if f {
        let years = i64::from(y2)
            .checked_sub(i64::from(y1))
            .and_then(|value| value.checked_add(1))
            .ok_or(ScalarError::Number)?;
        let leap_years = leap_years_inclusive(y1, y2)?;
        let total_days = years
            .checked_mul(365)
            .and_then(|value| value.checked_add(leap_years))
            .ok_or(ScalarError::Number)?;
        total_days as f64 / years as f64
    } else if is_leap_year(y1) {
        // Procedure E deliberately uses the first year whenever the dates
        // span years, even when the first date is after that year's leap day.
        366.0
    } else if february_29_between(start, end)? || (end.month == 2 && end.day == 29) {
        366.0
    } else {
        365.0
    };
    Ok(actual / denominator)
}

fn leap_years_through(year: i32) -> i64 {
    let year = i64::from(year);
    year.div_euclid(4) - year.div_euclid(100) + year.div_euclid(400)
}

fn leap_years_inclusive(first: i32, last: i32) -> Result<i64, ScalarError> {
    if first > last {
        return Ok(0);
    }
    let before = first.checked_sub(1).ok_or(ScalarError::Number)?;
    leap_years_through(last)
        .checked_sub(leap_years_through(before))
        .ok_or(ScalarError::Number)
}

/// Return whether a leap day lies strictly after the first date and before
/// the second date. Procedure E handles the first and second years with its
/// separate rules, so only intervening years and a post-February end year are
/// considered here.
fn february_29_between(start: CivilDate, end: CivilDate) -> Result<bool, ScalarError> {
    let interior_first = start.year.checked_add(1).ok_or(ScalarError::Number)?;
    let interior_last = end.year.checked_sub(1).ok_or(ScalarError::Number)?;
    if leap_years_inclusive(interior_first, interior_last)? > 0 {
        return Ok(true);
    }
    Ok(is_leap_year(end.year) && end.month > 2)
}

pub(in super::super) fn is_workday(
    date: CivilDate,
    workdays: &[bool; 7],
    holidays: &[i64],
) -> Result<bool, ScalarError> {
    let weekday = weekday_sunday_zero(date)? as usize;
    if workdays[weekday] {
        return Ok(false);
    }
    let serial = date_to_serial(date)? as i64;
    Ok(holidays.binary_search(&serial).is_err())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn date(year: i32, month: u32, day: u32) -> CivilDate {
        CivilDate::new(year, month, day).expect("test date is valid")
    }

    #[test]
    fn yearfrac_basis_zero_uses_procedure_a_february_ordering() {
        let leap_end = days360_basis_a(date(2020, 2, 29), date(2021, 2, 28)).unwrap();
        assert_eq!(leap_end, 360);
        assert_eq!(
            yearfrac(date(2020, 2, 29), date(2021, 2, 28), 0).unwrap(),
            1.0
        );

        let non_leap_end = days360_basis_a(date(2021, 2, 28), date(2022, 2, 28)).unwrap();
        assert_eq!(non_leap_end, 360);
    }

    #[test]
    fn yearfrac_basis_one_uses_procedure_e_start_year_rule() {
        let value = yearfrac(date(2020, 3, 1), date(2021, 3, 1), 1).unwrap();
        assert!((value - (365.0 / 366.0)).abs() < 1.0e-12);

        let multi_year = yearfrac(date(2019, 1, 1), date(2021, 1, 1), 1).unwrap();
        let expected = 731.0 / (1_096.0 / 3.0);
        assert!((multi_year - expected).abs() < 1.0e-12);
    }

    #[test]
    fn iso_week_uses_signed_thursday_offset() {
        assert_eq!(iso_week(date(2021, 1, 1)).unwrap(), (2020, 53));
        assert_eq!(iso_week(date(2021, 1, 4)).unwrap(), (2021, 1));
        assert_eq!(iso_week(date(2020, 12, 31)).unwrap(), (2020, 53));
    }
}
