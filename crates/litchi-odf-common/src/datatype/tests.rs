//! Behavioral tests for the public ODF scalar facade.

#![allow(
    clippy::unwrap_used,
    reason = "Test assertions intentionally unwrap known-valid scalar fixture construction failures."
)]

use super::{Boolean, Date, DateTime, Duration};
use chrono::{Datelike, Duration as ChronoDuration, NaiveDate, TimeZone, Utc};

#[test]
fn test_boolean_decode() {
    assert!(Boolean::decode("true").unwrap());
    assert!(!Boolean::decode("false").unwrap());
    assert!(Boolean::decode("invalid").is_err());
    assert!(Boolean::decode("TRUE").is_err());
    assert!(Boolean::decode("1").is_err());
}

#[test]
fn test_boolean_encode() {
    assert_eq!(Boolean::encode(true), "true");
    assert_eq!(Boolean::encode(false), "false");
}

#[test]
fn test_date_decode() {
    let date = Date::decode("2024-01-31").unwrap();
    assert_eq!(date, NaiveDate::from_ymd_opt(2024, 1, 31).unwrap());

    assert!(Date::decode("invalid").is_err());
    assert!(Date::decode("2024-13-01").is_err());
}

#[test]
fn test_date_encode() {
    let date = NaiveDate::from_ymd_opt(2024, 1, 31).unwrap();
    assert_eq!(Date::encode(&date), "2024-01-31");
}

#[test]
fn test_datetime_decode() {
    let dt = DateTime::decode("2024-01-31T15:30:00").unwrap();
    assert_eq!(dt.year(), 2024);
    assert_eq!(dt.month(), 1);
    assert_eq!(dt.day(), 31);

    let offset_datetime = DateTime::decode("2024-01-31T15:30:00+01:00").unwrap();
    assert_eq!(offset_datetime.year(), 2024);

    let utc_datetime = DateTime::decode("2024-01-31T15:30:00Z").unwrap();
    assert_eq!(utc_datetime.year(), 2024);
}

#[test]
fn test_datetime_encode() {
    let dt = Utc
        .with_ymd_and_hms(2024, 1, 31, 15, 30, 0)
        .unwrap()
        .fixed_offset();
    let encoded = DateTime::encode(&dt);
    assert!(encoded.ends_with('Z'));
    assert!(encoded.starts_with("2024-01-31"));
}

#[test]
fn test_duration_decode() {
    let dur = Duration::decode("PT1H30M").unwrap();
    assert_eq!(dur, ChronoDuration::minutes(90));

    let day_duration = Duration::decode("P1D").unwrap();
    assert_eq!(day_duration, ChronoDuration::days(1));

    let negative_duration = Duration::decode("-PT5M").unwrap();
    assert_eq!(negative_duration, ChronoDuration::minutes(-5));

    let compound_duration = Duration::decode("P1DT2H30M15S").unwrap();
    assert_eq!(
        compound_duration,
        ChronoDuration::days(1)
            + ChronoDuration::hours(2)
            + ChronoDuration::minutes(30)
            + ChronoDuration::seconds(15)
    );
}

#[test]
fn test_duration_encode() {
    let dur = ChronoDuration::minutes(90);
    assert_eq!(Duration::encode(&dur), "PT1H30M0S");

    let negative_duration = ChronoDuration::minutes(-5);
    assert_eq!(Duration::encode(&negative_duration), "-PT0H5M0S");

    let compound_duration =
        ChronoDuration::days(1) + ChronoDuration::hours(2) + ChronoDuration::minutes(30);
    assert_eq!(Duration::encode(&compound_duration), "PT26H30M0S");
}

#[test]
fn exact_duration_preserves_calendar_and_arbitrary_width_components() {
    let lexical = "-P123456789012345678901234567890Y11M30DT23H59M59.123456789012S";
    let duration = Duration::decode_exact(lexical).unwrap();

    assert_eq!(duration.as_str(), lexical);
    assert_eq!(duration.to_string(), lexical);
    assert!(duration.is_negative());
    assert_eq!(duration.years(), Some("123456789012345678901234567890"));
    assert_eq!(duration.months(), Some("11"));
    assert_eq!(duration.days(), Some("30"));
    assert_eq!(duration.hours(), Some("23"));
    assert_eq!(duration.minutes(), Some("59"));
    assert_eq!(duration.seconds(), Some("59.123456789012"));
    assert!(duration.to_chrono().is_err());
}

#[test]
fn duration_fractional_seconds_convert_and_encode_without_truncation() {
    assert_eq!(
        Duration::decode("PT1.125S").unwrap(),
        ChronoDuration::milliseconds(1125)
    );
    assert_eq!(
        Duration::decode("-PT0.000000001S").unwrap(),
        ChronoDuration::nanoseconds(-1)
    );
    assert_eq!(
        Duration::encode(&ChronoDuration::milliseconds(1125)),
        "PT0H0M1.125S"
    );
    assert_eq!(
        Duration::encode(&ChronoDuration::nanoseconds(-1)),
        "-PT0H0M0.000000001S"
    );
}

#[test]
fn exact_duration_retains_precision_beyond_chrono() {
    let exact = Duration::decode_exact("PT0.123456789012300S").unwrap();
    assert_eq!(exact.seconds(), Some("0.123456789012300"));
    assert!(exact.to_chrono().is_err());

    let exact_zero_tail = Duration::decode_exact("PT0.123456789000S").unwrap();
    assert_eq!(
        exact_zero_tail.to_chrono().unwrap(),
        ChronoDuration::nanoseconds(123_456_789)
    );
}

#[test]
fn duration_parser_rejects_malformed_component_grammar() {
    for value in [
        "",
        "P",
        "PT",
        "+P1D",
        "P1H",
        "PT1D",
        "P1DT2D",
        "P1D2Y",
        "P1Y2Y",
        "PT1M2H",
        "PT1S2M",
        "PT.5S",
        "PT1.S",
        "P1.5D",
        "P1DT2H3M4S5S",
    ] {
        assert!(
            Duration::decode_exact(value).is_err(),
            "accepted malformed duration {value}"
        );
    }
}

#[test]
fn duration_chrono_conversion_reports_range_errors() {
    let huge = Duration::decode_exact("P999999999999999999999999999D").unwrap();
    assert!(huge.to_chrono().is_err());
    assert!(Duration::decode("P999999999999999999999999999D").is_err());
}
