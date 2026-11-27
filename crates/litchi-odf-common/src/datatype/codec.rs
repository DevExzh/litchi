//! Checked ODF scalar parsing and canonical encoding.

use chrono::{
    DateTime as ChronoDateTime, Duration as ChronoDuration, FixedOffset, NaiveDate, NaiveDateTime,
    Utc,
};
use litchi_core::Result;

use super::model::{Boolean, Date, DateTime, Duration, DurationValue};

// ============================================================================
// BOOLEAN CONVERSION
// ============================================================================
// Reference: odfdo/datatype.py lines 39-58

impl Boolean {
    /// Decode an ODF boolean string to a Rust `bool`.
    ///
    /// Accepts only the schema literals `"true"` and `"false"`.
    ///
    /// # Errors
    ///
    /// Returns an error when `data` is not an `ODF` boolean literal.
    pub fn decode(data: &str) -> Result<bool> {
        match data {
            "true" => Ok(true),
            "false" => Ok(false),
            _ => Err(litchi_core::Error::Other(format!(
                "boolean '{data}' is invalid, expected 'true' or 'false'"
            ))),
        }
    }

    /// Encode a Rust `bool` as an ODF boolean string.
    #[inline]
    #[must_use]
    pub fn encode(value: bool) -> &'static str {
        if value { "true" } else { "false" }
    }
}

// ============================================================================
// DATE CONVERSION
// ============================================================================
// Reference: odfdo/datatype.py lines 61-74

impl Date {
    /// Decode an ISO 8601 ODF date string to `chrono::NaiveDate`.
    ///
    /// # Errors
    ///
    /// Returns an error when `data` is not an ISO 8601 calendar date.
    pub fn decode(data: &str) -> Result<NaiveDate> {
        NaiveDate::parse_from_str(data, "%Y-%m-%d").map_err(|error| {
            litchi_core::Error::Other(format!("Failed to parse ODF date '{data}': {error}"))
        })
    }

    /// Encode a `chrono::NaiveDate` as an ISO 8601 ODF date string.
    #[inline]
    #[must_use]
    pub fn encode(value: &NaiveDate) -> String {
        value.format("%Y-%m-%d").to_string()
    }
}

// ============================================================================
// DATETIME CONVERSION
// ============================================================================
// Reference: odfdo/datatype.py lines 77-111

impl DateTime {
    /// Decode an ODF datetime string to `chrono::DateTime<FixedOffset>`.
    ///
    /// RFC 3339 timezone forms are preserved. A timezone-free value is
    /// interpreted as UTC, matching the former scalar behavior.
    ///
    /// # Errors
    ///
    /// Returns an error when `data` cannot be parsed as an `ODF` date-time.
    pub fn decode(data: &str) -> Result<ChronoDateTime<FixedOffset>> {
        // Handle the `Z` suffix (UTC timezone).
        let normalized = if data.ends_with('Z') {
            data.replacen('Z', "+00:00", 1)
        } else {
            data.to_string()
        };

        // Try parsing with timezone.
        if let Ok(dt) = ChronoDateTime::parse_from_rfc3339(&normalized) {
            return Ok(dt);
        }

        // Try parsing without timezone (assume UTC).
        if let Ok(naive_dt) = NaiveDateTime::parse_from_str(&normalized, "%Y-%m-%dT%H:%M:%S") {
            return Ok(
                ChronoDateTime::<Utc>::from_naive_utc_and_offset(naive_dt, Utc).fixed_offset(),
            );
        }

        // Try with fractional seconds.
        if let Ok(naive_dt) = NaiveDateTime::parse_from_str(&normalized, "%Y-%m-%dT%H:%M:%S%.f") {
            return Ok(
                ChronoDateTime::<Utc>::from_naive_utc_and_offset(naive_dt, Utc).fixed_offset(),
            );
        }

        Err(litchi_core::Error::Other(format!(
            "Failed to parse ODF datetime '{data}'"
        )))
    }

    /// Encode a `chrono::DateTime<FixedOffset>` as an ODF datetime string.
    ///
    /// UTC values use the canonical `Z` suffix.
    #[must_use]
    pub fn encode(value: &ChronoDateTime<FixedOffset>) -> String {
        let formatted = value.to_rfc3339();
        // Convert +00:00 to Z for canonical representation.
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

impl DurationValue {
    /// Convert a day/time-only value to a checked [`chrono::Duration`].
    ///
    /// Non-zero calendar years or months require a reference date and are
    /// rejected. Fractional precision beyond nanoseconds is accepted only when
    /// the additional digits are zero.
    ///
    /// # Errors
    ///
    /// Returns an error when calendar units are nonzero or the day/time value
    /// cannot fit in `chrono::Duration`.
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
            let (whole_seconds_text, fraction) = seconds.split_once('.').unwrap_or((seconds, ""));
            let whole_seconds = parse_duration_i64(whole_seconds_text, "seconds")?;
            value = value
                .checked_add(
                    &ChronoDuration::try_seconds(whole_seconds)
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
                    let mut nanoseconds = significant.parse::<i64>().map_err(|_parse_error| {
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

impl Duration {
    /// Decode an ODF duration string to `chrono::Duration`.
    ///
    /// Supports ISO 8601 duration values such as `PT1H30M`, `P1DT2H`, and
    /// `-PT5M`.
    ///
    /// # Errors
    ///
    /// Returns an error when `data` is invalid, needs calendar arithmetic, or
    /// does not fit in `chrono::Duration`.
    pub fn decode(data: &str) -> Result<ChronoDuration> {
        Self::decode_exact(data)?.to_chrono()
    }

    /// Parse and retain a complete XML Schema duration without narrowing it.
    ///
    /// # Errors
    ///
    /// Returns an error when `data` is not a valid XML Schema duration.
    pub fn decode_exact(data: &str) -> Result<DurationValue> {
        parse_exact_duration(data)
    }

    /// Encode a `chrono::Duration` as an ODF duration string.
    #[inline]
    #[must_use]
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

fn add_duration_component(
    value: &mut ChronoDuration,
    component: Option<&str>,
    description: &str,
    unit: fn(i64) -> Option<ChronoDuration>,
) -> Result<()> {
    let Some(component_text) = component else {
        return Ok(());
    };
    let amount = parse_duration_i64(component_text, description)?;
    let part = unit(amount).ok_or_else(|| duration_range_error(description))?;
    *value = value
        .checked_add(&part)
        .ok_or_else(|| duration_range_error("total"))?;
    Ok(())
}

fn component_is_nonzero(component: &str) -> bool {
    component.bytes().any(|byte| byte != b'0')
}

fn parse_duration_i64(component: &str, description: &str) -> Result<i64> {
    component.parse::<i64>().map_err(|_parse_error| {
        litchi_core::Error::InvalidFormat(format!("duration {description} are out of range"))
    })
}

fn duration_range_error(component: &str) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(format!("duration {component} is out of range"))
}

fn parse_exact_duration(data: &str) -> Result<DurationValue> {
    if data.len() > 1_048_576 {
        return Err(litchi_core::Error::InvalidFormat(
            "duration exceeds 1 MiB".to_string(),
        ));
    }
    let (negative, unsigned_body) = data
        .strip_prefix('-')
        .map_or((false, data), |body| (true, body));
    let duration_body = unsigned_body.strip_prefix('P').ok_or_else(|| {
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
    let bytes = duration_body.as_bytes();
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

        let component = &duration_body[start..position];
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
