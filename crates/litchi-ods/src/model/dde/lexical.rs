//! Calendar and lexical checks for ODF cached date values.
//!
//! ODF `dateOrDateTime` uses XML Schema 1.0 `date` and `dateTime`. Keep the
//! original lexical value and avoid narrowing years or fractional seconds to
//! a machine date/time representation. Callers enforce their text-size limit.

use litchi_core::{Error, Result};

pub(super) fn validate_date_value(value: &str) -> Result<()> {
    let value = value.trim_matches([' ', '\t', '\r', '\n']);
    let valid = value.is_ascii()
        && match value.split_once('T') {
            Some((date, time)) => {
                date_tail(date.as_bytes()).is_some_and(|tail| tail.is_empty())
                    && valid_time(time.as_bytes())
            },
            None => date_tail(value.as_bytes()).is_some_and(valid_timezone),
        };
    if valid {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "DDE cached date must be an XML Schema date or dateTime".to_owned(),
        ))
    }
}

fn date_tail(value: &[u8]) -> Option<&[u8]> {
    let digits = value.strip_prefix(b"-").unwrap_or(value);
    let year_end = digits.iter().position(|byte| *byte == b'-')?;
    let year = digits.get(..year_end)?;
    if year.len() < 4
        || (year.len() > 4 && year.first() == Some(&b'0'))
        || !year.iter().all(u8::is_ascii_digit)
        || year.iter().all(|byte| *byte == b'0')
    {
        return None;
    }
    // Divisibility determines leap years even when the year cannot fit u64.
    let year_mod_400 = year.iter().fold(0u16, |remainder, byte| {
        (remainder * 10 + u16::from(*byte - b'0')) % 400
    });
    let after_year = digits.get(year_end + 1..)?;
    if after_year.get(2) != Some(&b'-') {
        return None;
    }
    let month = two_digits(after_year.get(..2)?)?;
    let day = two_digits(after_year.get(3..5)?)?;
    let leap = year_mod_400 % 4 == 0 && (year_mod_400 % 100 != 0 || year_mod_400 == 0);
    let maximum = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => return None,
    };
    if day == 0 || day > maximum {
        return None;
    }
    after_year.get(5..)
}

fn valid_time(value: &[u8]) -> bool {
    if value.get(2) != Some(&b':') || value.get(5) != Some(&b':') {
        return false;
    }
    let Some(hour) = value.get(..2).and_then(two_digits) else {
        return false;
    };
    let Some(minute) = value.get(3..5).and_then(two_digits) else {
        return false;
    };
    let Some(second) = value.get(6..8).and_then(two_digits) else {
        return false;
    };
    if hour > 24 || minute > 59 || second > 59 {
        return false;
    }
    let mut tail = value.get(8..).unwrap_or_default();
    let mut fractional_zero = true;
    if let Some(fraction) = tail.strip_prefix(b".") {
        let digits = fraction
            .iter()
            .take_while(|byte| byte.is_ascii_digit())
            .count();
        if digits == 0 {
            return false;
        }
        fractional_zero = fraction[..digits].iter().all(|byte| *byte == b'0');
        tail = &fraction[digits..];
    }
    (hour < 24 || (minute == 0 && second == 0 && fractional_zero)) && valid_timezone(tail)
}

fn valid_timezone(value: &[u8]) -> bool {
    if value.is_empty() || value == b"Z" {
        return true;
    }
    if value.len() != 6 || !matches!(value.first(), Some(b'+' | b'-')) || value[3] != b':' {
        return false;
    }
    let Some(hour) = two_digits(&value[1..3]) else {
        return false;
    };
    let Some(minute) = two_digits(&value[4..6]) else {
        return false;
    };
    hour <= 14 && minute <= 59 && (hour < 14 || minute == 0)
}

fn two_digits(value: &[u8]) -> Option<u8> {
    if value.len() == 2 && value.iter().all(u8::is_ascii_digit) {
        Some((value[0] - b'0') * 10 + value[1] - b'0')
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::validate_date_value;

    #[test]
    fn accepts_calendar_timezone_and_precision_boundaries() {
        for value in [
            "2026-09-13",
            "2024-02-29Z",
            "2000-02-29+14:00",
            "1900-02-28-14:00",
            "-0001-01-01",
            "-0004-02-29",
            "12345678912345678800-02-29",
            "2026-09-13T24:00:00",
            "2026-09-13T24:00:00.000Z",
            "2026-09-13T23:59:59.12345678901234567890-13:59",
            "\t2026-09-13Z\n",
        ] {
            assert!(validate_date_value(value).is_ok(), "{value}");
        }
    }

    #[test]
    fn rejects_invalid_calendar_and_time_fields() {
        for value in [
            "",
            "not-a-date",
            "0000-01-01",
            "-0000-01-01",
            "+2026-01-01",
            "02026-01-01",
            "2026-2-01",
            "2026-02-29",
            "1900-02-29",
            "2026-04-31",
            "2026-00-01",
            "2026-01-00",
            "2026-01-01+14:01",
            "2026-01-01+15:00",
            "2026-01-01+01:60",
            "2026-01-01z",
            "2026-01-01T25:00:00",
            "2026-01-01T24:00:01",
            "2026-01-01T24:00:00.1",
            "2026-01-01T23:59:60",
            "2026-01-01T12:60:00",
            "2026-01-01T12:00:00.",
            "2026-01-01T12:00:00.Z",
            "2026-01-01T12:00",
            "2026-01-01ZT12:00:00",
            "2026-01-01 T12:00:00",
        ] {
            assert!(validate_date_value(value).is_err(), "{value}");
        }
    }
}
