//! Strict XML Schema `date` and `dateTime` lexical validation.
//!
//! The OTH structure codec retains temporal values exactly as they appeared in
//! the source XML.  This module therefore only validates the lexical grammar:
//! it does not parse into a calendar type, normalize a timezone, or allocate a
//! canonical copy.  The grammar follows the strict ODS tracked-change and ODT
//! field validators, including signed/extended years and XML Schema's
//! `24:00:00` midnight spelling.
//!
//! Integration contract: `structure.rs` should call
//! [`is_xsd_date_or_datetime`] and retain its existing field-specific error
//! diagnostic when the predicate returns `false`.

/// The ODS and ODT temporal-value owners bound one retained lexical value to
/// 64 KiB.  Keep the same bound here so validation cannot spend unbounded time
/// walking a fractional-second or extended-year token.
pub(crate) const MAX_TEMPORAL_BYTES: usize = 65_536;

/// Returns whether `value` is an XML Schema `date` or `dateTime` lexical form.
///
/// A `T` selects the `dateTime` grammar; otherwise the `date` grammar is used.
/// The input is borrowed throughout, and no timezone or calendar
/// normalization is performed.
pub(crate) fn is_xsd_date_or_datetime(value: &str) -> bool {
    if value.len() > MAX_TEMPORAL_BYTES || !value.is_ascii() {
        return false;
    }
    if value.contains('T') {
        is_xsd_datetime(value)
    } else {
        is_xsd_date(value)
    }
}

/// Returns whether `value` is an XML Schema `date` lexical form.
pub(crate) fn is_xsd_date(value: &str) -> bool {
    if value.len() > MAX_TEMPORAL_BYTES || !value.is_ascii() {
        return false;
    }
    let Some((date, timezone)) = split_timezone(value) else {
        return false;
    };
    validate_date_core(date) && validate_timezone(timezone)
}

/// Returns whether `value` is an XML Schema `dateTime` lexical form.
pub(crate) fn is_xsd_datetime(value: &str) -> bool {
    if value.len() > MAX_TEMPORAL_BYTES || !value.is_ascii() {
        return false;
    }
    let Some((body, timezone)) = split_timezone(value) else {
        return false;
    };
    let Some((date, time)) = body.split_once('T') else {
        return false;
    };
    !time.contains('T')
        && validate_date_core(date)
        && validate_time(time)
        && validate_timezone(timezone)
}

/// Separates an optional XML Schema timezone without copying the input.
///
/// A final six-byte `±hh:mm` candidate is returned even when its numeric range
/// is invalid; [`validate_timezone`] then rejects it.  This distinction keeps
/// malformed offsets from being mistaken for part of the date or time body.
fn split_timezone(value: &str) -> Option<(&str, Option<&str>)> {
    if let Some(body) = value.strip_suffix('Z') {
        return (!body.is_empty()).then_some((body, Some("Z")));
    }

    let bytes = value.as_bytes();
    if bytes.len() >= 6
        && matches!(bytes[bytes.len() - 6], b'+' | b'-')
        && bytes[bytes.len() - 3] == b':'
    {
        let split = bytes.len() - 6;
        return Some((&value[..split], Some(&value[split..])));
    }
    Some((value, None))
}

fn validate_timezone(timezone: Option<&str>) -> bool {
    let Some(timezone) = timezone else {
        return true;
    };
    if timezone == "Z" {
        return true;
    }
    let bytes = timezone.as_bytes();
    if bytes.len() != 6
        || !matches!(bytes[0], b'+' | b'-')
        || bytes[3] != b':'
        || !bytes[1..3].iter().all(u8::is_ascii_digit)
        || !bytes[4..6].iter().all(u8::is_ascii_digit)
    {
        return false;
    }
    let hours = decimal2(&bytes[1..3]);
    let minutes = decimal2(&bytes[4..6]);
    minutes < 60 && (hours < 14 || (hours == 14 && minutes == 0))
}

fn validate_date_core(value: &str) -> bool {
    let unsigned = value.strip_prefix('-').unwrap_or(value);
    let mut parts = unsigned.split('-');
    let year = parts.next().unwrap_or_default();
    let Some(month) = parts.next() else {
        return false;
    };
    let Some(day) = parts.next() else {
        return false;
    };
    if parts.next().is_some()
        || year.len() < 4
        || (year.len() > 4 && year.starts_with('0'))
        || !year.bytes().all(|byte| byte.is_ascii_digit())
        || year.bytes().all(|byte| byte == b'0')
        || month.len() != 2
        || day.len() != 2
        || !month.bytes().all(|byte| byte.is_ascii_digit())
        || !day.bytes().all(|byte| byte.is_ascii_digit())
    {
        return false;
    }

    let month = decimal2(month.as_bytes());
    let day = decimal2(day.as_bytes());
    let max_day = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year_mod(year.as_bytes(), 400) == 0
            || (year_mod(year.as_bytes(), 4) == 0 && year_mod(year.as_bytes(), 100) != 0) =>
        {
            29
        },
        2 => 28,
        _ => return false,
    };
    (1..=max_day).contains(&day)
}

fn validate_time(value: &str) -> bool {
    let mut parts = value.split(':');
    let (Some(hour), Some(minute), Some(second), None) =
        (parts.next(), parts.next(), parts.next(), parts.next())
    else {
        return false;
    };
    let (second, fraction) = second
        .split_once('.')
        .map_or((second, None), |(whole, fraction)| (whole, Some(fraction)));
    if hour.len() != 2
        || minute.len() != 2
        || second.len() != 2
        || !hour.bytes().all(|byte| byte.is_ascii_digit())
        || !minute.bytes().all(|byte| byte.is_ascii_digit())
        || !second.bytes().all(|byte| byte.is_ascii_digit())
        || fraction.is_some_and(|value| {
            value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit())
        })
    {
        return false;
    }

    let hour = decimal2(hour.as_bytes());
    let minute = decimal2(minute.as_bytes());
    let second = decimal2(second.as_bytes());
    minute < 60
        && second < 60
        && (hour < 24
            || (hour == 24
                && minute == 0
                && second == 0
                && fraction.is_none_or(|value| value.bytes().all(|byte| byte == b'0'))))
}

fn decimal2(value: &[u8]) -> u8 {
    (value[0] - b'0') * 10 + (value[1] - b'0')
}

fn year_mod(value: &[u8], modulus: u16) -> u16 {
    value.iter().fold(0u16, |remainder, byte| {
        (remainder * 10 + u16::from(byte - b'0')) % modulus
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_schema_date_edges_without_normalizing_them() {
        for value in [
            "2024-01-01Z",
            "2024-02-29",
            "-0001-03-01-14:00",
            "12345-12-31+14:00",
            "2024-01-01-00:00",
        ] {
            assert!(is_xsd_date(value), "rejected {value}");
        }
        for value in [
            "0000-01-01",
            "-0000-01-01",
            "02024-01-01",
            "2023-02-29",
            "2024-04-31",
            "2024-01-01+14:01",
            "2024-01-01+15:00",
            "2024-01-01+00:60",
            "2024-01-01+0000",
            "2024-01-01z",
            "+2024-01-01",
            "2024-01-01 ",
        ] {
            assert!(!is_xsd_date(value), "accepted {value}");
        }
    }

    #[test]
    fn accepts_schema_datetime_edges_and_rejects_malformed_time() {
        for value in [
            "2024-01-01T00:00:00Z",
            "2024-02-29T23:59:59.123456789+14:00",
            "-0001-01-01T00:00:00-08:30",
            "12345-12-31T24:00:00.000000Z",
            "2024-01-01T12:34:56.12345678901234567890",
            "2024-01-01T24:00:00.0-14:00",
        ] {
            assert!(is_xsd_datetime(value), "rejected {value}");
            assert!(is_xsd_date_or_datetime(value));
        }
        for value in [
            "0000-01-01T00:00:00",
            "2023-02-29T00:00:00",
            "2024-01-01T24:00:00.1",
            "2024-01-01T24:00:01",
            "2024-01-01T00:00:60",
            "2024-01-01T00:00:00.",
            "2024-01-01T00:00:00+14:01",
            "2024-01-01T00:00:00+15:00",
            "2024-01-01T00:00:00+0000",
            "2024-01-01T00:00:00z",
            "2024-01-01T00:00:00TT00:00:00",
            "2024-01-01T00:00:00 ",
        ] {
            assert!(!is_xsd_datetime(value), "accepted {value}");
            assert!(!is_xsd_date_or_datetime(value));
        }
    }

    #[test]
    fn enforces_the_bounded_lexical_input_without_canonicalizing_fractional_digits() {
        let fraction = "1".repeat(MAX_TEMPORAL_BYTES);
        let value = format!("2024-01-01T00:00:00.{fraction}");
        assert!(!is_xsd_datetime(&value));

        let prefix = "2024-01-01T24:00:00.";
        let fraction = "0".repeat(MAX_TEMPORAL_BYTES - prefix.len() - 1);
        let value = format!("{prefix}{fraction}Z");
        assert!(is_xsd_datetime(&value));
        assert_eq!(
            value
                .strip_prefix(prefix)
                .and_then(|value| value.strip_suffix('Z')),
            Some(fraction.as_str())
        );
    }
}
