//! Lossless validated XML Schema time values used by core properties.

use super::{MAX_PROPERTY_TEXT, keyword::collapse_xml_whitespace};
use crate::{Error, Result};
use chrono::{DateTime as ChronoDateTime, NaiveDateTime, SecondsFormat, Utc};
use std::fmt;
use std::str::FromStr;

/// A lossless W3CDTF lexical value.
///
/// W3CDTF is the union of `xsd:gYear`, `xsd:gYearMonth`, `xsd:date`, and
/// `xsd:dateTime`. Time zones are optional for every member.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct W3c(String);

impl W3c {
    /// Validates and retains a W3CDTF lexical value.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        reject_oversized(&value)?;
        if !value.is_ascii() {
            return Err(invalid("W3CDTF", &value));
        }
        validate_w3c(&collapse_xml_whitespace(&value))?;
        Ok(Self(value))
    }

    /// Returns the retained lexical value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Moves out the retained lexical value.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0
    }

    pub(crate) fn utc(&self) -> Option<ChronoDateTime<Utc>> {
        parse_utc(self.as_str())
    }

    pub(crate) fn local(&self) -> Option<NaiveDateTime> {
        parse_local(self.as_str())
    }
}

impl AsRef<str> for W3c {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for W3c {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for W3c {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<String> for W3c {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::new(value)
    }
}

impl From<ChronoDateTime<Utc>> for W3c {
    fn from(value: ChronoDateTime<Utc>) -> Self {
        Self(value.to_rfc3339_opts(SecondsFormat::AutoSi, true))
    }
}

/// A lossless validated `xsd:dateTime` lexical value.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct DateTime(String);

impl DateTime {
    /// Validates and retains an `xsd:dateTime` lexical value.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        reject_oversized(&value)?;
        if !value.is_ascii() {
            return Err(invalid("xsd:dateTime", &value));
        }
        validate_date_time(&collapse_xml_whitespace(&value))?;
        Ok(Self(value))
    }

    /// Returns the retained lexical value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Moves out the retained lexical value.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0
    }

    pub(crate) fn utc(&self) -> Option<ChronoDateTime<Utc>> {
        parse_utc(self.as_str())
    }

    pub(crate) fn local(&self) -> Option<NaiveDateTime> {
        parse_local(self.as_str())
    }
}

impl AsRef<str> for DateTime {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for DateTime {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl FromStr for DateTime {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<String> for DateTime {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::new(value)
    }
}

impl From<ChronoDateTime<Utc>> for DateTime {
    fn from(value: ChronoDateTime<Utc>) -> Self {
        Self(value.to_rfc3339_opts(SecondsFormat::AutoSi, true))
    }
}

fn validate_w3c(value: &str) -> Result<()> {
    let (value, _) = split_timezone(value)?;
    if let Some((date, time)) = value.split_once('T') {
        if time.contains('T') {
            return Err(invalid("W3CDTF", value));
        }
        validate_date(date, DatePrecision::Day)?;
        validate_time(time)?;
        return Ok(());
    }
    let (_, precision) = validate_date(value, DatePrecision::Year)?;
    if matches!(
        precision,
        DatePrecision::Year | DatePrecision::Month | DatePrecision::Day
    ) {
        Ok(())
    } else {
        Err(invalid("W3CDTF", value))
    }
}

fn validate_date_time(value: &str) -> Result<()> {
    let (value, _) = split_timezone(value)?;
    let (date, time) = value
        .split_once('T')
        .ok_or_else(|| invalid("xsd:dateTime", value))?;
    if time.contains('T') {
        return Err(invalid("xsd:dateTime", value));
    }
    validate_date(date, DatePrecision::Day)?;
    validate_time(time)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DatePrecision {
    Year,
    Month,
    Day,
}

#[derive(Clone, Copy)]
struct Year<'a> {
    negative: bool,
    digits: &'a str,
}

fn validate_date<'a>(value: &'a str, minimum: DatePrecision) -> Result<(Year<'a>, DatePrecision)> {
    let negative = value.starts_with('-');
    let start = usize::from(negative);
    let separator = value[start..].find('-').map(|index| start + index);
    let year_end = separator.unwrap_or(value.len());
    let digits = &value[start..year_end];
    if digits.len() < 4
        || !digits.bytes().all(|byte| byte.is_ascii_digit())
        || digits.bytes().all(|byte| byte == b'0')
        || (digits.len() > 4 && digits.starts_with('0'))
    {
        return Err(invalid("XML Schema date", value));
    }
    let year = Year { negative, digits };
    let Some(_) = separator else {
        if minimum == DatePrecision::Year {
            return Ok((year, DatePrecision::Year));
        }
        return Err(invalid("XML Schema date", value));
    };
    let remainder = &value[year_end + 1..];
    if remainder.len() < 2 || !remainder[..2].bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid("XML Schema date", value));
    }
    let month = remainder[..2]
        .parse::<u8>()
        .map_err(|_| invalid("XML Schema date", value))?;
    if !(1..=12).contains(&month) {
        return Err(invalid("XML Schema date", value));
    }
    if remainder.len() == 2 {
        if minimum == DatePrecision::Year {
            return Ok((year, DatePrecision::Month));
        }
        return Err(invalid("XML Schema date", value));
    }
    if remainder.as_bytes().get(2) != Some(&b'-') || remainder.len() != 5 {
        return Err(invalid("XML Schema date", value));
    }
    let day = remainder[3..]
        .parse::<u8>()
        .map_err(|_| invalid("XML Schema date", value))?;
    if day == 0 || day > days_in_month(year, month) {
        return Err(invalid("XML Schema date", value));
    }
    Ok((year, DatePrecision::Day))
}

fn validate_time(value: &str) -> Result<()> {
    if value.len() < 8
        || value.as_bytes().get(2) != Some(&b':')
        || value.as_bytes().get(5) != Some(&b':')
        || !value[..2].bytes().all(|byte| byte.is_ascii_digit())
        || !value[3..5].bytes().all(|byte| byte.is_ascii_digit())
        || !value[6..8].bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(invalid("XML Schema time", value));
    }
    let fraction = if value.len() == 8 {
        ""
    } else if value.as_bytes().get(8) == Some(&b'.')
        && value.len() > 9
        && value[9..].bytes().all(|byte| byte.is_ascii_digit())
    {
        &value[9..]
    } else {
        return Err(invalid("XML Schema time", value));
    };
    let hour = value[..2]
        .parse::<u8>()
        .map_err(|_| invalid("XML Schema time", value))?;
    let minute = value[3..5]
        .parse::<u8>()
        .map_err(|_| invalid("XML Schema time", value))?;
    let second = value[6..8]
        .parse::<u8>()
        .map_err(|_| invalid("XML Schema time", value))?;
    if hour > 24
        || minute > 59
        || second > 59
        || (hour == 24 && (minute != 0 || second != 0 || fraction.bytes().any(|byte| byte != b'0')))
    {
        return Err(invalid("XML Schema time", value));
    }
    Ok(())
}

fn split_timezone(value: &str) -> Result<(&str, Option<&str>)> {
    if let Some(value) = value.strip_suffix('Z') {
        if value.is_empty() {
            return Err(invalid("XML Schema timezone", value));
        }
        return Ok((value, Some("Z")));
    }
    if value.len() >= 6 {
        let index = value.len() - 6;
        let zone = &value[index..];
        if matches!(zone.as_bytes().first(), Some(b'+' | b'-')) && zone.as_bytes()[3] == b':' {
            if !zone[1..3].bytes().all(|byte| byte.is_ascii_digit())
                || !zone[4..].bytes().all(|byte| byte.is_ascii_digit())
            {
                return Err(invalid("XML Schema timezone", zone));
            }
            let hours = zone[1..3]
                .parse::<u8>()
                .map_err(|_| invalid("XML Schema timezone", zone))?;
            let minutes = zone[4..]
                .parse::<u8>()
                .map_err(|_| invalid("XML Schema timezone", zone))?;
            if hours > 14 || minutes > 59 || (hours == 14 && minutes != 0) {
                return Err(invalid("XML Schema timezone", zone));
            }
            return Ok((&value[..index], Some(zone)));
        }
    }
    Ok((value, None))
}

fn days_in_month(year: Year<'_>, month: u8) -> u8 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_leap_year(year) => 29,
        2 => 28,
        _ => 0,
    }
}

fn is_leap_year(year: Year<'_>) -> bool {
    let magnitude = year.digits.bytes().fold(0_u16, |rest, byte| {
        (rest * 10 + u16::from(byte - b'0')) % 400
    });
    let astronomical = if year.negative {
        (401 - magnitude) % 400
    } else {
        magnitude
    };
    astronomical.is_multiple_of(4) && (!astronomical.is_multiple_of(100) || astronomical == 0)
}

fn parse_utc(value: &str) -> Option<ChronoDateTime<Utc>> {
    let value = collapse_xml_whitespace(value);
    if !value.contains('T') {
        return None;
    }
    ChronoDateTime::parse_from_rfc3339(&value)
        .ok()
        .map(|value| value.with_timezone(&Utc))
}

fn parse_local(value: &str) -> Option<NaiveDateTime> {
    let value = collapse_xml_whitespace(value);
    let (_, zone) = split_timezone(&value).ok()?;
    if zone.is_some() || !value.contains('T') {
        return None;
    }
    NaiveDateTime::parse_from_str(&value, "%Y-%m-%dT%H:%M:%S%.f").ok()
}

fn invalid(kind: &str, value: &str) -> Error {
    Error::Invalid(format!("invalid {kind} value '{value}'"))
}

fn reject_oversized(value: &str) -> Result<()> {
    if value.len() > MAX_PROPERTY_TEXT {
        return Err(Error::Limit {
            resource: "core property text bytes",
            max: MAX_PROPERTY_TEXT,
            actual: value.len(),
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_every_w3cdtf_member_and_retains_lexical_form() {
        for value in [
            "2026",
            "2026Z",
            "2026-08",
            "2026-08+05:30",
            "2026-08-03",
            "-0001-02-29Z",
            "2026-08-03T04:05:06",
            " 2026-08-03T04:05:06.007-00:00 ",
            "2026-08-03T24:00:00Z",
        ] {
            let parsed = W3c::new(value).expect(value);
            assert_eq!(parsed.as_str(), value);
        }
    }

    #[test]
    fn datetime_requires_a_full_valid_date_and_time() {
        assert!(DateTime::new("2026-08-03T04:05:06").is_ok());
        assert!(DateTime::new("2026-08-03T04:05:06Z").is_ok());
        for value in [
            "2026",
            "0000-01-01T00:00:00Z",
            "2026-02-29T00:00:00Z",
            "2026-08-03T24:00:01Z",
            "2026-08-03T04:05:06+14:01",
            "2026-13-03T04:05:06",
            "2026-08-03 04:05:06",
        ] {
            assert!(DateTime::new(value).is_err(), "{value}");
        }
    }

    #[test]
    fn chrono_conversion_is_ergonomic_without_inventing_unzoned_utc() {
        let chrono = ChronoDateTime::parse_from_rfc3339("2026-08-03T04:05:06Z")
            .unwrap()
            .with_timezone(&Utc);
        assert_eq!(W3c::from(chrono).as_str(), "2026-08-03T04:05:06Z");
        let unzoned = W3c::new("2026-08-03T04:05:06").unwrap();
        assert!(unzoned.utc().is_none());
        assert!(unzoned.local().is_some());
    }

    #[test]
    fn non_ascii_lexical_values_fail_without_panicking() {
        for value in [
            "é026-08-03T04:05:06Z",
            "2026-é8-03T04:05:06Z",
            "2026-08-03Té4:05:06Z",
            "2026-08-03T04:é5:06+05:30",
            "2026-08-03T04:05:06+é5:30",
        ] {
            let w3c = std::panic::catch_unwind(|| W3c::new(value));
            assert!(w3c.is_ok(), "W3c panicked for {value}");
            assert!(w3c.unwrap().is_err());
            let date_time = std::panic::catch_unwind(|| DateTime::new(value));
            assert!(date_time.is_ok(), "DateTime panicked for {value}");
            assert!(date_time.unwrap().is_err());
        }
    }

    #[test]
    fn rejects_oversized_times_before_whitespace_collapse() {
        let oversized = "2".repeat(MAX_PROPERTY_TEXT + 1);
        for result in [
            W3c::new(oversized.clone()).map(|_| ()),
            DateTime::new(oversized).map(|_| ()),
        ] {
            assert!(matches!(
                result,
                Err(Error::Limit {
                    resource: "core property text bytes",
                    max: MAX_PROPERTY_TEXT,
                    actual,
                }) if actual == MAX_PROPERTY_TEXT + 1
            ));
        }
    }
}
