//! ODF data type conversions (Boolean, Date, DateTime, Duration).
//!
//! This module provides conversion utilities between ODF format strings and Rust native types.
//! Based on reference implementations from odfdo library.
//!
//! # Implementation Status
//!
//! ✅ COMPLETED: Boolean conversion (ODF "true"/"false" ↔ Rust bool)
//! ✅ COMPLETED: Date conversion (ISO 8601 date ↔ chrono::NaiveDate)
//! ✅ COMPLETED: DateTime conversion (ISO 8601 datetime ↔ chrono::DateTime)
//! ✅ COMPLETED: Exact XML Schema duration parsing plus checked chrono conversion
//!
//! # References
//!
//! - odfdo: `3rdparty/odfdo/src/odfdo/datatype.py`

use chrono::{
    DateTime as ChronoDateTime, Duration as ChronoDuration, FixedOffset, NaiveDate, NaiveDateTime,
    Utc,
};
use litchi_core::Result;
use std::fmt;

// ============================================================================
// BOOLEAN CONVERSION
// ============================================================================
// Reference: odfdo/datatype.py lines 39-58

/// Boolean data type conversion utilities
///
/// Converts between ODF boolean format ("true"/"false") and Rust bool.
pub struct Boolean;

impl Boolean {
    /// Decode ODF boolean string to Rust bool
    ///
    /// # Arguments
    ///
    /// * `data` - ODF boolean string ("true" or "false")
    ///
    /// # Returns
    ///
    /// `Ok(bool)` on success, `Err` if the string is not "true" or "false"
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf_common::datatype::Boolean;
    ///
    /// assert_eq!(Boolean::decode("true").unwrap(), true);
    /// assert_eq!(Boolean::decode("false").unwrap(), false);
    /// assert!(Boolean::decode("invalid").is_err());
    /// ```
    pub fn decode(data: &str) -> Result<bool> {
        match data {
            "true" => Ok(true),
            "false" => Ok(false),
            _ => Err(litchi_core::Error::Other(format!(
                "boolean '{}' is invalid, expected 'true' or 'false'",
                data
            ))),
        }
    }

    /// Encode Rust bool to ODF boolean string
    ///
    /// # Arguments
    ///
    /// * `value` - Rust bool value
    ///
    /// # Returns
    ///
    /// ODF boolean string ("true" or "false")
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf_common::datatype::Boolean;
    ///
    /// assert_eq!(Boolean::encode(true), "true");
    /// assert_eq!(Boolean::encode(false), "false");
    /// ```
    #[inline]
    pub fn encode(value: bool) -> &'static str {
        if value { "true" } else { "false" }
    }
}

// ============================================================================
// DATE CONVERSION
// ============================================================================
// Reference: odfdo/datatype.py lines 61-74

/// Date data type conversion utilities
///
/// Converts between ODF date format (ISO 8601: "YYYY-MM-DD") and chrono::NaiveDate.
pub struct Date;

impl Date {
    /// Decode ODF date string to chrono::NaiveDate
    ///
    /// # Arguments
    ///
    /// * `data` - ISO 8601 date string (e.g., "2024-01-31")
    ///
    /// # Returns
    ///
    /// `Ok(NaiveDate)` on success, `Err` on parse error
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf_common::datatype::Date;
    /// use chrono::NaiveDate;
    ///
    /// let date = Date::decode("2024-01-31").unwrap();
    /// assert_eq!(date, NaiveDate::from_ymd_opt(2024, 1, 31).unwrap());
    /// ```
    pub fn decode(data: &str) -> Result<NaiveDate> {
        NaiveDate::parse_from_str(data, "%Y-%m-%d").map_err(|e| {
            litchi_core::Error::Other(format!("Failed to parse ODF date '{}': {}", data, e))
        })
    }

    /// Encode chrono::NaiveDate to ODF date string
    ///
    /// # Arguments
    ///
    /// * `value` - chrono::NaiveDate value
    ///
    /// # Returns
    ///
    /// ISO 8601 date string (format: "YYYY-MM-DD")
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf_common::datatype::Date;
    /// use chrono::NaiveDate;
    ///
    /// let date = NaiveDate::from_ymd_opt(2024, 1, 31).unwrap();
    /// assert_eq!(Date::encode(&date), "2024-01-31");
    /// ```
    #[inline]
    pub fn encode(value: &NaiveDate) -> String {
        value.format("%Y-%m-%d").to_string()
    }
}

// ============================================================================
// DATETIME CONVERSION
// ============================================================================
// Reference: odfdo/datatype.py lines 77-111

/// DateTime data type conversion utilities
///
/// Converts between ODF datetime format (ISO 8601) and chrono::DateTime.
pub struct DateTime;

impl DateTime {
    /// Decode ODF datetime string to chrono::DateTime
    ///
    /// Supports various ISO 8601 formats including timezone information.
    ///
    /// # Arguments
    ///
    /// * `data` - ISO 8601 datetime string
    ///
    /// # Returns
    ///
    /// `Ok(chrono::DateTime<FixedOffset>)` on success, `Err` on parse error
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf_common::datatype::DateTime;
    ///
    /// let dt = DateTime::decode("2024-01-31T15:30:00").unwrap();
    /// let dt_with_tz = DateTime::decode("2024-01-31T15:30:00+01:00").unwrap();
    /// let dt_utc = DateTime::decode("2024-01-31T15:30:00Z").unwrap();
    /// ```
    pub fn decode(data: &str) -> Result<ChronoDateTime<FixedOffset>> {
        // Handle 'Z' suffix (UTC timezone)
        let normalized = if data.ends_with('Z') {
            data.replacen('Z', "+00:00", 1)
        } else {
            data.to_string()
        };

        // Try parsing with timezone
        if let Ok(dt) = ChronoDateTime::parse_from_rfc3339(&normalized) {
            return Ok(dt);
        }

        // Try parsing without timezone (assume UTC)
        if let Ok(naive_dt) = NaiveDateTime::parse_from_str(&normalized, "%Y-%m-%dT%H:%M:%S") {
            return Ok(
                ChronoDateTime::<Utc>::from_naive_utc_and_offset(naive_dt, Utc).fixed_offset(),
            );
        }

        // Try with microseconds
        if let Ok(naive_dt) = NaiveDateTime::parse_from_str(&normalized, "%Y-%m-%dT%H:%M:%S%.f") {
            return Ok(
                ChronoDateTime::<Utc>::from_naive_utc_and_offset(naive_dt, Utc).fixed_offset(),
            );
        }

        Err(litchi_core::Error::Other(format!(
            "Failed to parse ODF datetime '{}'",
            data
        )))
    }

    /// Encode chrono::DateTime to ODF datetime string
    ///
    /// # Arguments
    ///
    /// * `value` - chrono::DateTime value
    ///
    /// # Returns
    ///
    /// ISO 8601 datetime string (UTC times end with 'Z')
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf_common::datatype::DateTime;
    /// use chrono::{TimeZone, Utc};
    ///
    /// let dt = Utc.with_ymd_and_hms(2024, 1, 31, 15, 30, 0).unwrap();
    /// let encoded = DateTime::encode(&dt.fixed_offset());
    /// assert!(encoded.ends_with("Z"));
    /// ```
    pub fn encode(value: &ChronoDateTime<FixedOffset>) -> String {
        let formatted = value.to_rfc3339();
        // Convert +00:00 to Z for canonical representation
        if formatted.ends_with("+00:00") {
            formatted.replacen("+00:00", "Z", 1)
        } else {
            formatted
        }
    }
}

// ============================================================================
// DURATION CONVERSION
// ============================================================================
// Reference: odfdo/datatype.py lines 114-165

/// Duration data type conversion utilities
///
/// Converts between ODF duration format (ISO 8601: "PT1H30M") and chrono::Duration.
pub struct Duration;

/// Exact XML Schema duration value used by ODF.
///
/// Calendar years and months cannot be represented by [`chrono::Duration`]
/// without a reference date. This type retains every component and its exact
/// lexical representation, including arbitrary-width integers and fractional
/// seconds.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DurationValue {
    lexical: String,
    negative: bool,
    years: Option<String>,
    months: Option<String>,
    days: Option<String>,
    hours: Option<String>,
    minutes: Option<String>,
    seconds: Option<String>,
}

impl DurationValue {
    /// Return the exact validated ODF lexical representation.
    pub fn as_str(&self) -> &str {
        &self.lexical
    }

    /// Whether the duration carries a negative sign.
    pub fn is_negative(&self) -> bool {
        self.negative
    }

    /// Calendar-year component, if present.
    pub fn years(&self) -> Option<&str> {
        self.years.as_deref()
    }

    /// Calendar-month component, if present.
    pub fn months(&self) -> Option<&str> {
        self.months.as_deref()
    }

    /// Day component, if present.
    pub fn days(&self) -> Option<&str> {
        self.days.as_deref()
    }

    /// Hour component, if present.
    pub fn hours(&self) -> Option<&str> {
        self.hours.as_deref()
    }

    /// Minute component, if present.
    pub fn minutes(&self) -> Option<&str> {
        self.minutes.as_deref()
    }

    /// Seconds component, including its fractional part, if present.
    pub fn seconds(&self) -> Option<&str> {
        self.seconds.as_deref()
    }

    /// Convert a day/time-only value to a checked [`chrono::Duration`].
    ///
    /// Non-zero calendar years or months require a reference date and are
    /// rejected. Fractional precision beyond nanoseconds is accepted only when
    /// the additional digits are zero.
    pub fn to_chrono(&self) -> Result<ChronoDuration> {
        if self.years.as_deref().is_some_and(component_is_nonzero)
            || self.months.as_deref().is_some_and(component_is_nonzero)
        {
            return Err(litchi_core::Error::InvalidFormat(
                "calendar years and months require a reference date".to_string(),
            ));
        }

        let mut value = ChronoDuration::zero();
        add_duration_component(
            &mut value,
            self.days.as_deref(),
            "days",
            ChronoDuration::try_days,
        )?;
        add_duration_component(
            &mut value,
            self.hours.as_deref(),
            "hours",
            ChronoDuration::try_hours,
        )?;
        add_duration_component(
            &mut value,
            self.minutes.as_deref(),
            "minutes",
            ChronoDuration::try_minutes,
        )?;

        if let Some(seconds) = &self.seconds {
            let (whole, fraction) = seconds.split_once('.').unwrap_or((seconds, ""));
            let whole = parse_duration_i64(whole, "seconds")?;
            value = value
                .checked_add(
                    &ChronoDuration::try_seconds(whole)
                        .ok_or_else(|| duration_range_error("seconds"))?,
                )
                .ok_or_else(|| duration_range_error("total"))?;
            if !fraction.is_empty() {
                let significant = fraction.trim_end_matches('0');
                if !significant.is_empty() {
                    if significant.len() > 9 {
                        return Err(litchi_core::Error::InvalidFormat(
                            "duration fractional seconds exceed nanosecond precision".to_string(),
                        ));
                    }
                    let mut nanoseconds = significant.parse::<i64>().map_err(|_| {
                        litchi_core::Error::InvalidFormat(
                            "duration fractional seconds are out of range".to_string(),
                        )
                    })?;
                    for _ in significant.len()..9 {
                        nanoseconds *= 10;
                    }
                    value = value
                        .checked_add(&ChronoDuration::nanoseconds(nanoseconds))
                        .ok_or_else(|| duration_range_error("total"))?;
                }
            }
        }

        if self.negative {
            ChronoDuration::zero()
                .checked_sub(&value)
                .ok_or_else(|| duration_range_error("total"))
        } else {
            Ok(value)
        }
    }
}

fn add_duration_component(
    value: &mut ChronoDuration,
    component: Option<&str>,
    description: &str,
    unit: fn(i64) -> Option<ChronoDuration>,
) -> Result<()> {
    let Some(component) = component else {
        return Ok(());
    };
    let amount = parse_duration_i64(component, description)?;
    let part = unit(amount).ok_or_else(|| duration_range_error(description))?;
    *value = value
        .checked_add(&part)
        .ok_or_else(|| duration_range_error("total"))?;
    Ok(())
}

impl fmt::Display for DurationValue {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.lexical)
    }
}

fn component_is_nonzero(component: &str) -> bool {
    component.bytes().any(|byte| byte != b'0')
}

fn parse_duration_i64(component: &str, description: &str) -> Result<i64> {
    component.parse::<i64>().map_err(|_| {
        litchi_core::Error::InvalidFormat(format!("duration {description} are out of range"))
    })
}

fn duration_range_error(component: &str) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(format!("duration {component} is out of range"))
}

impl Duration {
    /// Decode ODF duration string to chrono::Duration
    ///
    /// Supports ISO 8601 duration format (e.g., "PT1H30M", "P1DT2H", "-PT5M").
    ///
    /// # Arguments
    ///
    /// * `data` - ISO 8601 duration string
    ///
    /// # Returns
    ///
    /// `Ok(chrono::Duration)` on success, `Err` on parse error
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf_common::datatype::Duration;
    /// use chrono::Duration as ChronoDuration;
    ///
    /// let dur = Duration::decode("PT1H30M").unwrap();
    /// assert_eq!(dur, ChronoDuration::minutes(90));
    ///
    /// let dur_neg = Duration::decode("-PT5M").unwrap();
    /// assert_eq!(dur_neg, ChronoDuration::minutes(-5));
    /// ```
    pub fn decode(data: &str) -> Result<ChronoDuration> {
        Self::decode_exact(data)?.to_chrono()
    }

    /// Parse and retain a complete XML Schema duration without narrowing it.
    pub fn decode_exact(data: &str) -> Result<DurationValue> {
        parse_exact_duration(data)
    }

    /// Encode chrono::Duration to ODF duration string
    ///
    /// # Arguments
    ///
    /// * `value` - chrono::Duration value
    ///
    /// # Returns
    ///
    /// ISO 8601 duration string (format: "PT#H#M#S")
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf_common::datatype::Duration;
    /// use chrono::Duration as ChronoDuration;
    ///
    /// let dur = ChronoDuration::minutes(90);
    /// assert_eq!(Duration::encode(&dur), "PT1H30M0S");
    ///
    /// let dur_neg = ChronoDuration::minutes(-5);
    /// assert_eq!(Duration::encode(&dur_neg), "-PT0H5M0S");
    /// ```
    pub fn encode(value: &ChronoDuration) -> String {
        let total_seconds = value.num_seconds();
        let subsecond_nanoseconds = value.subsec_nanos();
        let negative = total_seconds < 0 || subsecond_nanoseconds < 0;
        let sign = if negative { "-" } else { "" };
        let abs_seconds = total_seconds.unsigned_abs();
        let abs_nanoseconds = subsecond_nanoseconds.unsigned_abs();

        let hours = abs_seconds / 3600;
        let minutes = (abs_seconds % 3600) / 60;
        let seconds = abs_seconds % 60;
        if abs_nanoseconds == 0 {
            format!("{sign}PT{hours}H{minutes}M{seconds}S")
        } else {
            let fraction = format!("{abs_nanoseconds:09}");
            format!(
                "{sign}PT{hours}H{minutes}M{seconds}.{}S",
                fraction.trim_end_matches('0')
            )
        }
    }
}

fn parse_exact_duration(data: &str) -> Result<DurationValue> {
    if data.len() > 1_048_576 {
        return Err(litchi_core::Error::InvalidFormat(
            "duration exceeds 1 MiB".to_string(),
        ));
    }
    let (negative, body) = data
        .strip_prefix('-')
        .map_or((false, data), |body| (true, body));
    let body = body.strip_prefix('P').ok_or_else(|| {
        litchi_core::Error::InvalidFormat(format!(
            "invalid duration '{data}': expected a 'P' designator"
        ))
    })?;

    let mut value = DurationValue {
        lexical: data.to_string(),
        negative,
        years: None,
        months: None,
        days: None,
        hours: None,
        minutes: None,
        seconds: None,
    };
    let bytes = body.as_bytes();
    let mut position = 0usize;
    let mut in_time = false;
    let mut last_rank = 0u8;
    let mut component_count = 0usize;
    let mut time_component_count = 0usize;

    while position < bytes.len() {
        if bytes[position] == b'T' {
            if in_time {
                return Err(invalid_duration(data, "duplicate 'T' designator"));
            }
            in_time = true;
            last_rank = 0;
            position += 1;
            continue;
        }

        let start = position;
        while position < bytes.len() && bytes[position].is_ascii_digit() {
            position += 1;
        }
        if position == start {
            return Err(invalid_duration(data, "expected a numeric component"));
        }
        if position < bytes.len() && bytes[position] == b'.' {
            position += 1;
            let fraction_start = position;
            while position < bytes.len() && bytes[position].is_ascii_digit() {
                position += 1;
            }
            if position == fraction_start {
                return Err(invalid_duration(data, "empty fractional seconds"));
            }
        }
        if position == bytes.len() {
            return Err(invalid_duration(data, "component has no designator"));
        }

        let component = &body[start..position];
        let designator = bytes[position];
        position += 1;
        let (rank, slot): (u8, &mut Option<String>) = match (in_time, designator) {
            (false, b'Y') => (1, &mut value.years),
            (false, b'M') => (2, &mut value.months),
            (false, b'D') => (3, &mut value.days),
            (true, b'H') => (1, &mut value.hours),
            (true, b'M') => (2, &mut value.minutes),
            (true, b'S') => (3, &mut value.seconds),
            _ => return Err(invalid_duration(data, "invalid or misplaced designator")),
        };
        if component.contains('.') && designator != b'S' {
            return Err(invalid_duration(
                data,
                "only seconds may contain a fraction",
            ));
        }
        if rank <= last_rank || slot.is_some() {
            return Err(invalid_duration(
                data,
                "components are duplicated or unordered",
            ));
        }
        last_rank = rank;
        *slot = Some(component.to_string());
        component_count += 1;
        if in_time {
            time_component_count += 1;
        }
    }

    if component_count == 0 {
        return Err(invalid_duration(data, "duration has no components"));
    }
    if in_time && time_component_count == 0 {
        return Err(invalid_duration(data, "'T' has no time components"));
    }
    Ok(value)
}

fn invalid_duration(data: &str, description: &str) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(format!("invalid duration '{data}': {description}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::{Datelike, TimeZone, Utc};

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
        assert!(Date::decode("2024-13-01").is_err()); // Invalid month
    }

    #[test]
    fn test_date_encode() {
        let date = NaiveDate::from_ymd_opt(2024, 1, 31).unwrap();
        assert_eq!(Date::encode(&date), "2024-01-31");
    }

    #[test]
    fn test_datetime_decode() {
        // Without timezone
        let dt = DateTime::decode("2024-01-31T15:30:00").unwrap();
        assert_eq!(dt.year(), 2024);
        assert_eq!(dt.month(), 1);
        assert_eq!(dt.day(), 31);

        // With timezone
        let dt = DateTime::decode("2024-01-31T15:30:00+01:00").unwrap();
        assert_eq!(dt.year(), 2024);

        // UTC (Z suffix)
        let dt = DateTime::decode("2024-01-31T15:30:00Z").unwrap();
        assert_eq!(dt.year(), 2024);
    }

    #[test]
    fn test_datetime_encode() {
        let dt = Utc
            .with_ymd_and_hms(2024, 1, 31, 15, 30, 0)
            .unwrap()
            .fixed_offset();
        let encoded = DateTime::encode(&dt);
        assert!(encoded.ends_with("Z"));
        assert!(encoded.starts_with("2024-01-31"));
    }

    #[test]
    fn test_duration_decode() {
        // Hours and minutes
        let dur = Duration::decode("PT1H30M").unwrap();
        assert_eq!(dur, ChronoDuration::minutes(90));

        // Days
        let dur = Duration::decode("P1D").unwrap();
        assert_eq!(dur, ChronoDuration::days(1));

        // Negative
        let dur = Duration::decode("-PT5M").unwrap();
        assert_eq!(dur, ChronoDuration::minutes(-5));

        // Complex
        let dur = Duration::decode("P1DT2H30M15S").unwrap();
        assert_eq!(
            dur,
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

        let dur = ChronoDuration::minutes(-5);
        assert_eq!(Duration::encode(&dur), "-PT0H5M0S");

        let dur = ChronoDuration::days(1) + ChronoDuration::hours(2) + ChronoDuration::minutes(30);
        assert_eq!(Duration::encode(&dur), "PT26H30M0S"); // 24+2 hours
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
}
