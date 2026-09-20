//! Shared deterministic Gregorian calendar data and serial-date helpers.
//! The supported evaluation profile uses 1899-12-30 without a fictitious
//! 1900 leap day. Callers validate their accepted serial/date domain.

pub(super) const SECONDS_PER_DAY: f64 = 86_400.0;
pub(super) const SERIAL_EPOCH_TO_UNIX_DAYS: i64 = 25_569;

/// The date/time family uses the proleptic Gregorian profile from year one
/// through year 9999.  Keep these bounds as integers as well as floating
/// point constants so civil-date arithmetic never has to round a boundary.
pub(super) const MIN_CALENDAR_YEAR: i32 = 1;
pub(super) const MAX_CALENDAR_YEAR: i32 = 9_999;
pub(super) const MIN_DATE_SERIAL: i64 = -693_593;
pub(super) const MAX_DATE_SERIAL: i64 = 2_958_465;
pub(super) const MAX_DATE_SERIAL_EXCLUSIVE: i64 = 2_958_466;

/// A validated proleptic-Gregorian civil date.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct CivilDate {
    pub(super) year: i32,
    pub(super) month: u32,
    pub(super) day: u32,
}

impl CivilDate {
    /// Construct a date in the selected profile.
    #[must_use]
    pub(super) fn new(year: i32, month: u32, day: u32) -> Option<Self> {
        if !(MIN_CALENDAR_YEAR..=MAX_CALENDAR_YEAR).contains(&year) {
            return None;
        }
        let limit = days_in_month(year, month)?;
        (1..=limit)
            .contains(&day)
            .then_some(Self { year, month, day })
    }
}

// Howard Hinnant's proleptic-Gregorian civil-date conversion, expressed in
// terms of days since 1970-01-01 and kept integer-only for deterministic evaluation.
pub(super) fn civil_from_days(days: i64) -> (i32, u32, u32) {
    let z = days + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = mp + if mp < 10 { 3 } else { -9 };
    let year = y + if month <= 2 { 1 } else { 0 };
    (
        i32::try_from(year).unwrap_or(i32::MAX),
        u32::try_from(month).unwrap_or(u32::MAX),
        u32::try_from(day).unwrap_or(u32::MAX),
    )
}

/// Convert a profile civil date to days since 1970-01-01.
///
/// The checked entry point keeps invalid month/day combinations out of the
/// integer conversion.  The resulting range is small enough for all
/// supported dates, but the intermediate arithmetic remains checked at the
/// public boundary.
#[must_use]
pub(super) fn days_from_civil(date: CivilDate) -> Option<i64> {
    if CivilDate::new(date.year, date.month, date.day) != Some(date) {
        return None;
    }
    let year = i64::from(date.year) - i64::from(date.month <= 2);
    let era = if year >= 0 {
        year / 400
    } else {
        (year - 399) / 400
    };
    let year_of_era = year.checked_sub(era.checked_mul(400)?)?;
    let month = i64::from(date.month);
    let month_term = month.checked_add(if date.month > 2 { -3 } else { 9 })?;
    let day_of_year = (153_i64.checked_mul(month_term)?.checked_add(2)? / 5)
        .checked_add(i64::from(date.day))?
        .checked_sub(1)?;
    let day_of_era = year_of_era
        .checked_mul(365)?
        .checked_add(year_of_era / 4)?
        .checked_sub(year_of_era / 100)?
        .checked_add(day_of_year)?;
    era.checked_mul(146_097)?
        .checked_add(day_of_era)?
        .checked_sub(719_468)
}

/// Convert days since 1970-01-01 to a validated profile date.
#[must_use]
pub(super) fn civil_from_days_checked(days: i64) -> Option<CivilDate> {
    let min_days = MIN_DATE_SERIAL.checked_sub(SERIAL_EPOCH_TO_UNIX_DAYS)?;
    let max_days = MAX_DATE_SERIAL.checked_sub(SERIAL_EPOCH_TO_UNIX_DAYS)?;
    if !(min_days..=max_days).contains(&days) {
        return None;
    }
    let (year, month, day) = civil_from_days(days);
    CivilDate::new(year, month, day)
}

/// Convert a validated profile date to its integer serial.
#[must_use]
pub(super) fn serial_from_civil(date: CivilDate) -> Option<i64> {
    let serial = days_from_civil(date)?.checked_add(SERIAL_EPOCH_TO_UNIX_DAYS)?;
    (MIN_DATE_SERIAL..=MAX_DATE_SERIAL)
        .contains(&serial)
        .then_some(serial)
}

/// Split a finite date/time serial into its civil date and nonnegative day
/// fraction.  The half-open upper bound admits every time on the final date
/// while excluding the first instant after 9999-12-31.
#[must_use]
pub(super) fn civil_from_serial(serial: f64) -> Option<(CivilDate, f64)> {
    let min = MIN_DATE_SERIAL as f64;
    let max = MAX_DATE_SERIAL_EXCLUSIVE as f64;
    if !serial.is_finite() || serial < min || serial >= max {
        return None;
    }
    let whole = serial.floor();
    if !whole.is_finite() || whole < i64::MIN as f64 || whole > i64::MAX as f64 {
        return None;
    }
    let whole = whole as i64;
    let mut fraction = serial - whole as f64;
    // For a negative subnormal serial, subtraction from floor(-tiny) can
    // round the mathematical fraction to exactly one.  Preserve the date
    // decomposition without manufacturing a next-day serial; the profile
    // only promises the finite f64 value itself, so the nearest representable
    // fraction below one is the deterministic fallback.
    if fraction == 1.0 {
        fraction = 1.0 - f64::EPSILON;
    } else if fraction == 0.0 {
        fraction = 0.0;
    }
    if !fraction.is_finite() || !(0.0..1.0).contains(&fraction) {
        return None;
    }
    let days = whole.checked_sub(SERIAL_EPOCH_TO_UNIX_DAYS)?;
    Some((civil_from_days_checked(days)?, fraction))
}

/// Return the number of days in one profile month.
#[must_use]
pub(super) const fn days_in_month(year: i32, month: u32) -> Option<u32> {
    if month == 0 || month > 12 {
        return None;
    }
    Some(match month {
        2 if is_leap_year(year) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    })
}

/// Return whether a profile year is a Gregorian leap year.
#[must_use]
pub(super) const fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

pub(super) const MONTH_SHORT: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];
pub(super) const MONTH_LONG: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_profile_date_serial_round_trips() {
        for serial in MIN_DATE_SERIAL..=MAX_DATE_SERIAL {
            let (date, fraction) = civil_from_serial(serial as f64)
                .expect("every admitted integer serial has a civil date");
            assert_eq!(fraction, 0.0);
            assert_eq!(serial_from_civil(date), Some(serial));
        }
    }

    #[test]
    fn checked_civil_conversion_rejects_extreme_day_counts() {
        assert!(civil_from_days_checked(i64::MIN).is_none());
        assert!(civil_from_days_checked(i64::MAX).is_none());
    }
}
