//! Shared deterministic Gregorian calendar data and serial-date helpers.
//! The supported evaluation profile uses 1899-12-30 without a fictitious
//! 1900 leap day. Callers validate their accepted serial/date domain.

pub(super) const SECONDS_PER_DAY: f64 = 86_400.0;
pub(super) const SERIAL_EPOCH_TO_UNIX_DAYS: i64 = 25_569;

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
