//! Package-independent values for PresentationML slide synchronization.

use chrono::NaiveDate;
use litchi_opc::PackURI;

use crate::{Error, Result};

const MIN_YEAR: u32 = 1;
const MAX_YEAR: u32 = 9999;
const MAX_FRACTION_DIGITS: usize = 16;
const MAX_TIMEZONE_HOURS: u8 = 14;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// UTC offset recorded by an `xsd:dateTime` value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Offset {
    /// No time-zone designator was recorded.
    Unspecified,
    /// The `Z` designator (UTC).
    Utc,
    /// An explicit `+hh:mm` or `-hh:mm` offset from UTC.
    Explicit {
        /// True when the offset is behind UTC.
        negative: bool,
        /// Whole-hour component (`0..=14`).
        hours: u8,
        /// Minute component (`0..=59`, and `0` when `hours == 14`).
        minutes: u8,
    },
}

/// Validated `xsd:dateTime` value used by slide synchronization metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateTime {
    /// Gregorian year (`1..=9999`).
    pub year: u32,
    /// Month (`1..=12`).
    pub month: u8,
    /// Day of the month.
    pub day: u8,
    /// Hour (`0..=23`).
    pub hour: u8,
    /// Minute (`0..=59`).
    pub minute: u8,
    /// Second (`0..=59`).
    pub second: u8,
    /// Fractional-second digits, when recorded.
    pub fraction_digits: Option<String>,
    /// Recorded UTC offset.
    pub offset: Offset,
}

impl DateTime {
    /// Parse and validate an `xsd:dateTime` lexical form.
    pub fn parse(value: &str) -> Result<Self> {
        let digits = |text: &str, label: &str| -> Result<u32> {
            if text.is_empty() || !text.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid(format!(
                    "slide synchronization {label} is not numeric"
                )));
            }
            text.parse::<u32>()
                .map_err(|_| invalid(format!("slide synchronization {label} is out of range")))
        };
        let (date, rest) = value.split_once('T').ok_or_else(|| {
            invalid("slide synchronization timestamp is missing the 'T' separator")
        })?;
        let mut date_parts = date.split('-');
        let year_text = date_parts.next().unwrap_or_default();
        if year_text.len() != 4 {
            return Err(invalid("slide synchronization year is not four digits"));
        }
        let year = digits(year_text, "year")?;
        if !(MIN_YEAR..=MAX_YEAR).contains(&year) {
            return Err(invalid("slide synchronization year is out of range"));
        }
        let month = digits(date_parts.next().unwrap_or_default(), "month")?;
        let day = digits(date_parts.next().unwrap_or_default(), "day")?;
        if date_parts.next().is_some() {
            return Err(invalid("slide synchronization date has trailing fields"));
        }
        let month = u8::try_from(month)
            .map_err(|_| invalid("slide synchronization month is out of range"))?;
        let day =
            u8::try_from(day).map_err(|_| invalid("slide synchronization day is out of range"))?;
        if NaiveDate::from_ymd_opt(year as i32, month.into(), day.into()).is_none() {
            return Err(invalid("slide synchronization date is not a calendar day"));
        }

        let (time, offset) = match rest.find(['Z', '+', '-']) {
            Some(index) => {
                let (time, zone) = rest.split_at(index);
                let offset = match zone.as_bytes()[0] {
                    b'Z' => {
                        if zone.len() != 1 {
                            return Err(invalid(
                                "slide synchronization UTC designator has trailing data",
                            ));
                        }
                        Offset::Utc
                    },
                    marker => {
                        let mut zone_parts = zone[1..].split(':');
                        let hours =
                            digits(zone_parts.next().unwrap_or_default(), "time-zone hours")?;
                        let minutes =
                            digits(zone_parts.next().unwrap_or_default(), "time-zone minutes")?;
                        if zone_parts.next().is_some() {
                            return Err(invalid(
                                "slide synchronization time zone has trailing fields",
                            ));
                        }
                        let hours = u8::try_from(hours).map_err(|_| {
                            invalid("slide synchronization time-zone hours are out of range")
                        })?;
                        let minutes = u8::try_from(minutes).map_err(|_| {
                            invalid("slide synchronization time-zone minutes are out of range")
                        })?;
                        if hours > MAX_TIMEZONE_HOURS
                            || (hours == MAX_TIMEZONE_HOURS && minutes != 0)
                            || minutes > 59
                        {
                            return Err(invalid("slide synchronization time zone exceeds 14:00"));
                        }
                        Offset::Explicit {
                            negative: marker == b'-',
                            hours,
                            minutes,
                        }
                    },
                };
                (time, offset)
            },
            None => (rest, Offset::Unspecified),
        };

        let (hms, fraction_digits) = match time.split_once('.') {
            Some((hms, fraction)) => {
                if fraction.is_empty()
                    || fraction.len() > MAX_FRACTION_DIGITS
                    || !fraction.bytes().all(|byte| byte.is_ascii_digit())
                {
                    return Err(invalid(
                        "slide synchronization fractional seconds are invalid",
                    ));
                }
                (hms, Some(fraction.to_owned()))
            },
            None => (time, None),
        };
        let mut time_parts = hms.split(':');
        let hour = digits(time_parts.next().unwrap_or_default(), "hour")?;
        let minute = digits(time_parts.next().unwrap_or_default(), "minute")?;
        let second = digits(time_parts.next().unwrap_or_default(), "second")?;
        if time_parts.next().is_some() {
            return Err(invalid("slide synchronization time has trailing fields"));
        }
        let hour = u8::try_from(hour)
            .map_err(|_| invalid("slide synchronization hour is out of range"))?;
        let minute = u8::try_from(minute)
            .map_err(|_| invalid("slide synchronization minute is out of range"))?;
        let second = u8::try_from(second)
            .map_err(|_| invalid("slide synchronization second is out of range"))?;
        if hour > 23 || minute > 59 || second > 59 {
            return Err(invalid("slide synchronization time of day is out of range"));
        }
        Ok(Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
            fraction_digits,
            offset,
        })
    }

    /// Serialize in canonical PresentationML `xsd:dateTime` form.
    pub fn to_lexical(&self) -> String {
        let mut out = format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            self.year, self.month, self.day, self.hour, self.minute, self.second
        );
        if let Some(fraction) = &self.fraction_digits {
            out.push('.');
            out.push_str(fraction);
        }
        match self.offset {
            Offset::Unspecified => {},
            Offset::Utc => out.push('Z'),
            Offset::Explicit {
                negative,
                hours,
                minutes,
            } => {
                out.push(if negative { '-' } else { '+' });
                out.push_str(&format!("{hours:02}:{minutes:02}"));
            },
        }
        out
    }
}

/// Inert slide-library synchronization metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Properties {
    /// Server-side slide file identifier.
    pub server_slide_id: String,
    /// Last modification time of the server-side slide.
    pub server_modified_time: DateTime,
    /// Time the slide was inserted into this presentation.
    pub client_inserted_time: DateTime,
    /// Whether the source part carried an extension list.
    pub has_extension_list: bool,
}

impl Properties {
    /// Construct synchronization metadata without an extension list.
    pub fn new(
        server_slide_id: impl Into<String>,
        server_modified_time: DateTime,
        client_inserted_time: DateTime,
    ) -> Self {
        Self {
            server_slide_id: server_slide_id.into(),
            server_modified_time,
            client_inserted_time,
            has_extension_list: false,
        }
    }

    /// Retain an opaque extension-list marker when authoring a replacement.
    pub const fn with_extension_list(mut self, present: bool) -> Self {
        self.has_extension_list = present;
        self
    }
}

/// Synchronization metadata bound to the slide relationship that owns it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    /// Relationship ID from the source slide.
    pub relationship_id: String,
    /// Source slide part name.
    pub slide_part_name: PackURI,
    /// Synchronization data part name.
    pub part_name: PackURI,
    /// Parsed synchronization metadata.
    pub properties: Properties,
}

impl Part {
    /// Bind properties to a slide and a new synchronization part identity.
    pub fn new(
        relationship_id: impl Into<String>,
        slide_part_name: PackURI,
        part_name: PackURI,
        properties: Properties,
    ) -> Self {
        Self {
            relationship_id: relationship_id.into(),
            slide_part_name,
            part_name,
            properties,
        }
    }
}
