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
//! ✅ COMPLETED: Duration conversion (ISO 8601 duration ↔ chrono::Duration)
//!
//! # References
//!
//! - odfdo: `3rdparty/odfdo/src/odfdo/datatype.py`

use crate::Result;
use chrono::{DateTime, Duration, FixedOffset, NaiveDate, NaiveDateTime, Utc};

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
    /// use litchi::odf::datatype::Boolean;
    ///
    /// assert_eq!(Boolean::decode("true").unwrap(), true);
    /// assert_eq!(Boolean::decode("false").unwrap(), false);
    /// assert!(Boolean::decode("invalid").is_err());
    /// ```
    pub fn decode(data: &str) -> Result<bool> {
        match data {
            "true" => Ok(true),
            "false" => Ok(false),
            _ => Err(crate::Error::Other(format!(
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
    /// use litchi::odf::datatype::Boolean;
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
    /// use litchi::odf::datatype::Date;
    /// use chrono::NaiveDate;
    ///
    /// let date = Date::decode("2024-01-31").unwrap();
    /// assert_eq!(date, NaiveDate::from_ymd_opt(2024, 1, 31).unwrap());
    /// ```
    pub fn decode(data: &str) -> Result<NaiveDate> {
        NaiveDate::parse_from_str(data, "%Y-%m-%d")
            .map_err(|e| crate::Error::Other(format!("Failed to parse ODF date '{}': {}", data, e)))
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
    /// use litchi::odf::datatype::Date;
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
pub struct DateTimeOdf;

impl DateTimeOdf {
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
    /// `Ok(DateTime<FixedOffset>)` on success, `Err` on parse error
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::datatype::DateTimeOdf;
    ///
    /// let dt = DateTimeOdf::decode("2024-01-31T15:30:00").unwrap();
    /// let dt_with_tz = DateTimeOdf::decode("2024-01-31T15:30:00+01:00").unwrap();
    /// let dt_utc = DateTimeOdf::decode("2024-01-31T15:30:00Z").unwrap();
    /// ```
    pub fn decode(data: &str) -> Result<DateTime<FixedOffset>> {
        // Handle 'Z' suffix (UTC timezone)
        let normalized = if data.ends_with('Z') {
            data.replacen('Z', "+00:00", 1)
        } else {
            data.to_string()
        };

        // Try parsing with timezone
        if let Ok(dt) = DateTime::parse_from_rfc3339(&normalized) {
            return Ok(dt);
        }

        // Try parsing without timezone (assume UTC)
        if let Ok(naive_dt) = NaiveDateTime::parse_from_str(&normalized, "%Y-%m-%dT%H:%M:%S") {
            return Ok(DateTime::<Utc>::from_naive_utc_and_offset(naive_dt, Utc).fixed_offset());
        }

        // Try with microseconds
        if let Ok(naive_dt) = NaiveDateTime::parse_from_str(&normalized, "%Y-%m-%dT%H:%M:%S%.f") {
            return Ok(DateTime::<Utc>::from_naive_utc_and_offset(naive_dt, Utc).fixed_offset());
        }

        Err(crate::Error::Other(format!(
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
    /// use litchi::odf::datatype::DateTimeOdf;
    /// use chrono::{DateTime, Utc, TimeZone};
    ///
    /// let dt = Utc.with_ymd_and_hms(2024, 1, 31, 15, 30, 0).unwrap();
    /// let encoded = DateTimeOdf::encode(&dt.fixed_offset());
    /// assert!(encoded.ends_with("Z"));
    /// ```
    pub fn encode(value: &DateTime<FixedOffset>) -> String {
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
pub struct DurationOdf;

impl DurationOdf {
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
    /// `Ok(Duration)` on success, `Err` on parse error
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::datatype::DurationOdf;
    /// use chrono::Duration;
    ///
    /// let dur = DurationOdf::decode("PT1H30M").unwrap();
    /// assert_eq!(dur, Duration::minutes(90));
    ///
    /// let dur_neg = DurationOdf::decode("-PT5M").unwrap();
    /// assert_eq!(dur_neg, Duration::minutes(-5));
    /// ```
    pub fn decode(data: &str) -> Result<Duration> {
        let (sign, data) = if let Some(rest) = data.strip_prefix('-') {
            (-1, rest)
        } else {
            (1, data)
        };

        if !data.starts_with('P') {
            return Err(crate::Error::Other(format!(
                "Invalid duration format '{}', must start with 'P'",
                data
            )));
        }

        let mut days = 0i64;
        let mut hours = 0i64;
        let mut minutes = 0i64;
        let mut seconds = 0i64;

        let mut buffer = String::new();
        let mut in_time = false;

        for c in data.chars().skip(1) {
            // Skip 'P'
            match c {
                '0'..='9' => buffer.push(c),
                'D' => {
                    days = buffer
                        .parse()
                        .map_err(|_| crate::Error::Other("Invalid days in duration".to_string()))?;
                    buffer.clear();
                },
                'T' => {
                    in_time = true;
                },
                'H' => {
                    if !in_time {
                        return Err(crate::Error::Other(
                            "Hours must come after 'T' in duration".to_string(),
                        ));
                    }
                    hours = buffer.parse().map_err(|_| {
                        crate::Error::Other("Invalid hours in duration".to_string())
                    })?;
                    buffer.clear();
                },
                'M' => {
                    if in_time {
                        minutes = buffer.parse().map_err(|_| {
                            crate::Error::Other("Invalid minutes in duration".to_string())
                        })?;
                    } else {
                        // Months not supported in chrono::Duration
                        return Err(crate::Error::Other(
                            "Months in duration not supported".to_string(),
                        ));
                    }
                    buffer.clear();
                },
                'S' => {
                    if !in_time {
                        return Err(crate::Error::Other(
                            "Seconds must come after 'T' in duration".to_string(),
                        ));
                    }
                    seconds = buffer.parse().map_err(|_| {
                        crate::Error::Other("Invalid seconds in duration".to_string())
                    })?;
                    buffer.clear();
                },
                _ => {
                    return Err(crate::Error::Other(format!(
                        "Invalid character '{}' in duration",
                        c
                    )));
                },
            }
        }

        let total_seconds = days * 86400 + hours * 3600 + minutes * 60 + seconds;
        Ok(Duration::seconds(total_seconds * sign))
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
    /// use litchi::odf::datatype::DurationOdf;
    /// use chrono::Duration;
    ///
    /// let dur = Duration::minutes(90);
    /// assert_eq!(DurationOdf::encode(&dur), "PT1H30M0S");
    ///
    /// let dur_neg = Duration::minutes(-5);
    /// assert_eq!(DurationOdf::encode(&dur_neg), "-PT0H5M0S");
    /// ```
    pub fn encode(value: &Duration) -> String {
        let total_seconds = value.num_seconds();
        let (sign, abs_seconds) = if total_seconds < 0 {
            ("-", -total_seconds)
        } else {
            ("", total_seconds)
        };

        let hours = abs_seconds / 3600;
        let minutes = (abs_seconds % 3600) / 60;
        let seconds = abs_seconds % 60;

        format!("{}PT{}H{}M{}S", sign, hours, minutes, seconds)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::{Datelike, TimeZone, Utc};

    #[test]
    fn test_boolean_decode() {
        assert_eq!(Boolean::decode("true").unwrap(), true);
        assert_eq!(Boolean::decode("false").unwrap(), false);
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
        let dt = DateTimeOdf::decode("2024-01-31T15:30:00").unwrap();
        assert_eq!(dt.year(), 2024);
        assert_eq!(dt.month(), 1);
        assert_eq!(dt.day(), 31);

        // With timezone
        let dt = DateTimeOdf::decode("2024-01-31T15:30:00+01:00").unwrap();
        assert_eq!(dt.year(), 2024);

        // UTC (Z suffix)
        let dt = DateTimeOdf::decode("2024-01-31T15:30:00Z").unwrap();
        assert_eq!(dt.year(), 2024);
    }

    #[test]
    fn test_datetime_encode() {
        let dt = Utc
            .with_ymd_and_hms(2024, 1, 31, 15, 30, 0)
            .unwrap()
            .fixed_offset();
        let encoded = DateTimeOdf::encode(&dt);
        assert!(encoded.ends_with("Z"));
        assert!(encoded.starts_with("2024-01-31"));
    }

    #[test]
    fn test_duration_decode() {
        // Hours and minutes
        let dur = DurationOdf::decode("PT1H30M").unwrap();
        assert_eq!(dur, Duration::minutes(90));

        // Days
        let dur = DurationOdf::decode("P1D").unwrap();
        assert_eq!(dur, Duration::days(1));

        // Negative
        let dur = DurationOdf::decode("-PT5M").unwrap();
        assert_eq!(dur, Duration::minutes(-5));

        // Complex
        let dur = DurationOdf::decode("P1DT2H30M15S").unwrap();
        assert_eq!(
            dur,
            Duration::days(1) + Duration::hours(2) + Duration::minutes(30) + Duration::seconds(15)
        );
    }

    #[test]
    fn test_duration_encode() {
        let dur = Duration::minutes(90);
        assert_eq!(DurationOdf::encode(&dur), "PT1H30M0S");

        let dur = Duration::minutes(-5);
        assert_eq!(DurationOdf::encode(&dur), "-PT0H5M0S");

        let dur = Duration::days(1) + Duration::hours(2) + Duration::minutes(30);
        assert_eq!(DurationOdf::encode(&dur), "PT26H30M0S"); // 24+2 hours
    }
}
