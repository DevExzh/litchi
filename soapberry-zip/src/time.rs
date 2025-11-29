//! ZIP file timestamp handling
//!
//! Datetimes for ZIP files come in two flavors: UTC and local time. It is not
//! possible for the local time zone to be encoded in the ZIP format, so
//! converting between the two requires assuming that UTC is the local time.
//!
//! When reading a ZIP file, [`ZipDateTimeKind`] will provide information about
//! the timestamp's original time zone (UTC and local time)
//!
//! However, when writing a ZIP file, only a [`UtcDateTime`] is supported.
//!
//! # Example: Copying Modification Times
//!
//! This example shows how to read a ZIP file and create a new one while
//! preserving modification times:
//!
//! ```
//! use rawzip::{ZipArchive, ZipArchiveWriter, ZipDataWriter};
//! use rawzip::time::{ZipDateTimeKind, UtcDateTime};
//! use std::io::Write;
//!
//! // Read a test ZIP file with timestamps
//! let input_data = include_bytes!("../assets/time-go.zip");
//! let input_archive = ZipArchive::from_slice(input_data).unwrap();
//!
//! // Create output archive
//! let mut output_data = Vec::new();
//! let mut output_archive = ZipArchiveWriter::new(&mut output_data);
//!
//! // Copy each entry with its modification time
//! let mut entries = input_archive.entries();
//! while let Ok(Some(entry)) = entries.next_entry() {
//!     let name = entry.file_path().try_normalize().unwrap().as_ref().to_string();
//!     let modification_time = entry.last_modified();
//!     
//!     let utc_time = match modification_time {
//!         ZipDateTimeKind::Utc(utc_time) => utc_time,
//!         ZipDateTimeKind::Local(local_time) => {
//!             // Convert local time to UTC by reinterpreting the components
//!             // This treats the local time as if it were UTC
//!             UtcDateTime::from_components(
//!                 local_time.year(),
//!                 local_time.month(),
//!                 local_time.day(),
//!                 local_time.hour(),
//!                 local_time.minute(),
//!                 local_time.second(),
//!                 local_time.nanosecond()
//!             ).unwrap()
//!         }
//!     };
//!
//!     if !entry.is_dir() {
//!         // Copy file with preserved modification time
//!         let (mut entry, config) = output_archive.new_file(&name)
//!             .last_modified(utc_time)
//!             .start()
//!             .unwrap();
//!         let mut writer = config.wrap(&mut entry);
//!         writer.write_all(b"example data").unwrap();
//!         let (_, descriptor) = writer.finish().unwrap();
//!         entry.finish(descriptor).unwrap();
//!     } else {
//!         // Copy directory with preserved modification time
//!         output_archive.new_dir(&name)
//!             .last_modified(utc_time)
//!             .create()
//!             .unwrap();
//!     }
//! }
//!
//! output_archive.finish().unwrap();
//!
//! // Verify the output archive preserves timestamps
//! let output_archive = ZipArchive::from_slice(&output_data).unwrap();
//!
//! assert!(output_archive.entries_hint() > 0, "Output should contain entries");
//!
//! // Verify at least one entry has a UTC timestamp
//! let mut output_entries = output_archive.entries();
//! let mut has_utc_timestamp = false;
//! while let Ok(Some(entry)) = output_entries.next_entry() {
//!     if matches!(entry.last_modified(), ZipDateTimeKind::Utc(_)) {
//!         has_utc_timestamp = true;
//!         break;
//!     }
//! }
//! assert!(has_utc_timestamp, "Output should contain UTC timestamps");
//! ```

use crate::{
    extra_fields::{ExtraFieldId, ExtraFields},
    utils::{le_u16, le_u32, le_u64},
};

/// Represents the time zone of a timestamp.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeZone {
    /// UTC (Coordinated Universal Time)
    Utc,
    /// Local time (timezone unknown)
    Local,
}

/// Marker type for UTC timezone
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Utc;

/// Marker type for Local timezone
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Local;

/// Trait for timezone markers
pub trait TimeZoneMarker {
    fn timezone() -> TimeZone;
}

impl TimeZoneMarker for Utc {
    fn timezone() -> TimeZone {
        TimeZone::Utc
    }
}

impl TimeZoneMarker for Local {
    fn timezone() -> TimeZone {
        TimeZone::Local
    }
}

/// Represents a timestamp found in a ZIP file
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct ZipDateTime<TZ = Utc> {
    year: u16,
    month: u8,       // 1-12
    day: u8,         // 1-31
    hour: u8,        // 0-23
    minute: u8,      // 0-59
    second: u8,      // 0-59
    nanosecond: u32, // 0-999,999,999
    _timezone: std::marker::PhantomData<TZ>,
}

/// Type alias for UTC timestamps
pub type UtcDateTime = ZipDateTime<Utc>;

/// Type alias for Local timestamps  
pub type LocalDateTime = ZipDateTime<Local>;

/// Enum for timestamp parsing results that can be either UTC or Local
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ZipDateTimeKind {
    Utc(UtcDateTime),
    Local(LocalDateTime),
}

impl ZipDateTimeKind {
    /// Returns the timezone of this timestamp
    #[must_use]
    pub const fn timezone(&self) -> TimeZone {
        match self {
            ZipDateTimeKind::Utc(_) => TimeZone::Utc,
            ZipDateTimeKind::Local(_) => TimeZone::Local,
        }
    }

    /// Returns the year component of the timestamp
    #[must_use]
    pub fn year(&self) -> u16 {
        match self {
            ZipDateTimeKind::Utc(dt) => dt.year(),
            ZipDateTimeKind::Local(dt) => dt.year(),
        }
    }

    /// Returns the month component (1-12) of the timestamp
    #[must_use]
    pub fn month(&self) -> u8 {
        match self {
            ZipDateTimeKind::Utc(dt) => dt.month(),
            ZipDateTimeKind::Local(dt) => dt.month(),
        }
    }

    /// Returns the day component (1-31) of the timestamp
    #[must_use]
    pub fn day(&self) -> u8 {
        match self {
            ZipDateTimeKind::Utc(dt) => dt.day(),
            ZipDateTimeKind::Local(dt) => dt.day(),
        }
    }

    /// Returns the hour component (0-23) of the timestamp
    #[must_use]
    pub fn hour(&self) -> u8 {
        match self {
            ZipDateTimeKind::Utc(dt) => dt.hour(),
            ZipDateTimeKind::Local(dt) => dt.hour(),
        }
    }

    /// Returns the minute component (0-59) of the timestamp
    #[must_use]
    pub fn minute(&self) -> u8 {
        match self {
            ZipDateTimeKind::Utc(dt) => dt.minute(),
            ZipDateTimeKind::Local(dt) => dt.minute(),
        }
    }

    /// Returns the second component (0-59) of the timestamp
    #[must_use]
    pub fn second(&self) -> u8 {
        match self {
            ZipDateTimeKind::Utc(dt) => dt.second(),
            ZipDateTimeKind::Local(dt) => dt.second(),
        }
    }

    /// Returns the nanosecond component (0-999,999,999) of the timestamp
    #[must_use]
    pub fn nanosecond(&self) -> u32 {
        match self {
            ZipDateTimeKind::Utc(dt) => dt.nanosecond(),
            ZipDateTimeKind::Local(dt) => dt.nanosecond(),
        }
    }
}

impl std::fmt::Display for ZipDateTimeKind {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ZipDateTimeKind::Utc(dt) => dt.fmt(f),
            ZipDateTimeKind::Local(dt) => dt.fmt(f),
        }
    }
}

impl<TZ: TimeZoneMarker> std::fmt::Display for ZipDateTime<TZ> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Write out the date and time in ISO 8601 format. RFC 3339 requires a
        // time zone, which we won't have for local times.
        write!(
            f,
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            self.year, self.month, self.day, self.hour, self.minute, self.second
        )?;
        if self.nanosecond != 0 {
            write!(f, ".{:09}", self.nanosecond)?;
        }
        match TZ::timezone() {
            TimeZone::Utc => write!(f, "Z"),
            TimeZone::Local => Ok(()),
        }
    }
}

impl<TZ: TimeZoneMarker> ZipDateTime<TZ> {
    /// Creates a ZipDateTime from date/time components with validation.
    ///
    /// # Arguments
    ///
    /// * `year` - Year (1-65535)
    /// * `month` - Month (1-12)
    /// * `day` - Day of month (1-31, validated against month)
    /// * `hour` - Hour (0-23)
    /// * `minute` - Minute (0-59)
    /// * `second` - Second (0-59)
    /// * `nanosecond` - Nanosecond (0-999,999,999), defaults to 0
    ///
    /// # Errors
    ///
    /// Returns `None` if any component is invalid or the date doesn't exist
    /// (e.g. February 30th, April 31st).
    ///
    /// # Examples
    ///
    /// ```
    /// # use rawzip::time::{UtcDateTime, LocalDateTime};
    /// let utc_datetime = UtcDateTime::from_components(
    ///     2023, 6, 15, 14, 30, 45, 500_000_000
    /// ).unwrap();
    /// assert_eq!(utc_datetime.year(), 2023);
    /// assert_eq!(utc_datetime.nanosecond(), 500_000_000);
    ///
    /// // Invalid date returns None
    /// assert!(UtcDateTime::from_components(2023, 2, 30, 0, 0, 0, 0).is_none());
    /// ```
    pub fn from_components(
        year: u16,
        month: u8,
        day: u8,
        hour: u8,
        minute: u8,
        second: u8,
        nanosecond: u32,
    ) -> Option<Self> {
        // Validate components
        if year == 0
            || month == 0
            || month > 12
            || day == 0
            || hour > 23
            || minute > 59
            || second > 59
            || nanosecond > 999_999_999
        {
            return None;
        }

        let max_day = last_day_of_month(year, month);
        if day > max_day {
            return None;
        }

        Some(Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
            nanosecond,
            _timezone: std::marker::PhantomData,
        })
    }

    /// Returns the year component of the timestamp.
    #[must_use]
    pub const fn year(&self) -> u16 {
        self.year
    }

    /// Returns the month component (1-12) of the timestamp.
    #[must_use]
    pub const fn month(&self) -> u8 {
        self.month
    }

    /// Returns the day component (1-31) of the timestamp.
    #[must_use]
    pub const fn day(&self) -> u8 {
        self.day
    }

    /// Returns the hour component (0-23) of the timestamp.
    #[must_use]
    pub const fn hour(&self) -> u8 {
        self.hour
    }

    /// Returns the minute component (0-59) of the timestamp.
    #[must_use]
    pub const fn minute(&self) -> u8 {
        self.minute
    }

    /// Returns the second component (0-59) of the timestamp.
    #[must_use]
    pub const fn second(&self) -> u8 {
        self.second
    }

    /// Returns the nanosecond component (0-999,999,999) of the timestamp.
    /// For timestamps that don't support nanosecond precision, this returns 0.
    #[must_use]
    pub const fn nanosecond(&self) -> u32 {
        self.nanosecond
    }

    /// Returns the timezone of this timestamp.
    #[must_use]
    pub fn timezone(&self) -> TimeZone {
        TZ::timezone()
    }

    /// Calculate days since Unix epoch (1970-01-01) for this date.
    ///
    /// Based on Howard Hinnant's `days_from_civil` algorithm:
    /// <https://howardhinnant.github.io/date_algorithms.html#days_from_civil>
    ///
    /// Negative values indicate dates prior to 1970-01-01.
    const fn days_from_civil(&self) -> i32 {
        let (y, m) = if self.month <= 2 {
            (self.year as i32 - 1, self.month as i32 + 9)
        } else {
            (self.year as i32, self.month as i32 - 3)
        };

        // Calculate era (400-year cycles)
        let era = y / 400;
        let yoe = y - era * 400; // year of era [0, 399]

        // Calculate day of year
        let doy = (153 * m + 2) / 5 + self.day as i32 - 1; // day of year [0, 365]

        // Calculate day of era
        let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy; // day of era [0, 146096]

        // Calculate days since epoch (era 0 starts at year 0, not 1970)
        era * 146097 + doe - 719468
    }
}

impl ZipDateTime<Utc> {
    /// Creates a ZipDateTime from a Unix timestamp (seconds since epoch)
    pub fn from_unix(seconds: i64) -> UtcDateTime {
        let (year, month, day, hour, minute, second) = unix_timestamp_to_components(seconds);
        ZipDateTime {
            year,
            month,
            day,
            hour,
            minute,
            second,
            nanosecond: 0,
            _timezone: std::marker::PhantomData,
        }
    }

    /// Creates a ZipDateTime from an NTFS timestamp (100ns ticks since 1601)
    pub(crate) fn from_ntfs(ticks: u64) -> UtcDateTime {
        let unix_seconds = (ticks / 10_000_000).saturating_sub(NTFS_EPOCH_OFFSET) as i64;
        let (year, month, day, hour, minute, second) = unix_timestamp_to_components(unix_seconds);
        let nanosecond = ((ticks % 10_000_000) * 100) as u32;
        ZipDateTime {
            year,
            month,
            day,
            hour,
            minute,
            second,
            nanosecond,
            _timezone: std::marker::PhantomData,
        }
    }

    /// Convert to Unix timestamp (seconds since epoch).
    ///
    /// Returns the number of seconds since the Unix epoch (1970-01-01 00:00:00 UTC).
    /// Negative values represent dates before 1970.
    #[must_use]
    pub fn to_unix(&self) -> i64 {
        let days_since_epoch = self.days_from_civil();

        (i64::from(days_since_epoch)) * 86400
            + (i64::from(self.hour)) * 3600
            + (i64::from(self.minute)) * 60
            + (i64::from(self.second))
    }
}

impl ZipDateTime<Local> {
    /// Creates a ZipDateTime from a DosDateTime
    pub(crate) fn from_dos(dos: DosDateTime) -> LocalDateTime {
        // Note: DOS timestamps with month=0 and day=0 are a gray area. Some
        // seem to normalize to 1980-01-01 while others normalize to 1979-11-30.
        ZipDateTime {
            year: dos.year(),
            month: dos.month(),
            day: dos.day(),
            hour: dos.hour(),
            minute: dos.minute(),
            second: dos.second(),
            nanosecond: 0,
            _timezone: std::marker::PhantomData,
        }
    }
}

/// Represents an MS-DOS timestamp with 2-second precision.
///
/// MS-DOS timestamps are stored as packed 16-bit values for date and time,
/// with a limited range from 1980 to 2107 and 2-second precision for seconds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DosDateTime {
    time: u16,
    date: u16,
}

impl DosDateTime {
    /// Creates a new MS-DOS datetime from packed date and time values.
    #[must_use]
    pub(crate) const fn new(time: u16, date: u16) -> Self {
        Self { time, date }
    }

    /// Returns the year (1980-2107).
    #[must_use]
    pub fn year(&self) -> u16 {
        ((self.date >> 9) & 0x7f) + 1980
    }

    /// Returns the month (1-12).
    #[must_use]
    pub fn month(&self) -> u8 {
        let raw_month = ((self.date >> 5) & 0x0f) as u8;
        raw_month.clamp(1, 12)
    }

    /// Returns the day of the month (1-31).
    #[must_use]
    pub fn day(&self) -> u8 {
        let raw_day = (self.date & 0x1f) as u8;
        raw_day.clamp(1, last_day_of_month(self.year(), self.month()))
    }

    /// Returns the hour (0-23).
    #[must_use]
    pub fn hour(&self) -> u8 {
        let raw_hour = ((self.time >> 11) & 0x1f) as u8;
        raw_hour.min(23)
    }

    /// Returns the minute (0-59).
    #[must_use]
    pub fn minute(&self) -> u8 {
        let raw_minute = ((self.time >> 5) & 0x3f) as u8;
        raw_minute.min(59)
    }

    /// Returns the second (0-58, always even due to 2-second precision).
    #[must_use]
    pub fn second(&self) -> u8 {
        let raw_second = ((self.time & 0x1f) * 2) as u8;
        raw_second.min(58)
    }

    /// Returns the packed time and date components as (time, date).
    #[must_use]
    pub(crate) const fn into_parts(self) -> (u16, u16) {
        (self.time, self.date)
    }
}

impl From<&ZipDateTime> for DosDateTime {
    fn from(zip_dt: &ZipDateTime) -> Self {
        // Saturate year to DOS range (1980-2107)
        let dos_year = zip_dt.year.clamp(1980, 2107);

        // Pack the date: bits 15-9: year-1980, bits 8-5: month, bits 4-0: day
        let packed_date =
            ((dos_year - 1980) << 9) | ((zip_dt.month as u16) << 5) | (zip_dt.day as u16);

        // Pack the time: bits 15-11: hour, bits 10-5: minute, bits 4-0: second/2
        let packed_time = ((zip_dt.hour as u16) << 11)
            | ((zip_dt.minute as u16) << 5)
            | ((zip_dt.second as u16) / 2);

        Self {
            time: packed_time,
            date: packed_date,
        }
    }
}

/// Extracts timestamp from the extra field using "last wins" strategy.
/// Returns the last valid timestamp found, or falls back to MS-DOS if none found.
/// This matches Go's zip reader behavior.
pub(crate) fn extract_best_timestamp(
    extra_fields: ExtraFields<'_>,
    dos_time: u16,
    dos_date: u16,
) -> ZipDateTimeKind {
    let mut last_timestamp = None;

    for (field_id, field_data) in extra_fields {
        match field_id {
            ExtraFieldId::NTFS => {
                if let Some(timestamp) = parse_ntfs_timestamp(field_data) {
                    last_timestamp = Some(ZipDateTimeKind::Utc(timestamp));
                }
            },
            ExtraFieldId::EXTENDED_TIMESTAMP => {
                if let Some(timestamp) = parse_extended_timestamp(field_data) {
                    last_timestamp = Some(ZipDateTimeKind::Utc(timestamp));
                }
            },
            ExtraFieldId::INFO_ZIP_UNIX_ORIGINAL => {
                if let Some(timestamp) = parse_unix_timestamp(field_data) {
                    last_timestamp = Some(ZipDateTimeKind::Utc(timestamp));
                }
            },
            _ => {},
        }
    }

    // Return the last timestamp found, or fall back to MS-DOS
    last_timestamp.unwrap_or_else(|| {
        ZipDateTimeKind::Local(LocalDateTime::from_dos(DosDateTime::new(
            dos_time, dos_date,
        )))
    })
}

/// Parses NTFS timestamp extra field (0x000a)
fn parse_ntfs_timestamp(data: &[u8]) -> Option<UtcDateTime> {
    if data.len() < 32 {
        return None;
    }

    // NTFS extra field format:
    // 4 bytes: reserved (usually 0)
    // 2 bytes: attribute tag (0x0001 for timestamps)
    // 2 bytes: attribute size (24 bytes for 3 timestamps)
    // 8 bytes: modification time
    // 8 bytes: access time
    // 8 bytes: creation time

    let tag = le_u16(&data[4..6]);
    if tag != 0x0001 {
        return None;
    }

    let size = le_u16(&data[6..8]) as usize;
    if size < 24 || data.len() < 8 + size {
        return None;
    }

    // Extract modification time (first 8 bytes of timestamp data)
    let mtime_ticks = le_u64(&data[8..16]);
    Some(UtcDateTime::from_ntfs(mtime_ticks))
}

/// Parses Extended Timestamp extra field (0x5455)
fn parse_extended_timestamp(data: &[u8]) -> Option<UtcDateTime> {
    if data.len() < 5 {
        return None;
    }

    let flags = data[0];
    let pos = 1;

    // Check if modification time is present (bit 0)
    if flags & 0x01 != 0 && pos + 4 <= data.len() {
        let mtime_seconds = le_u32(&data[pos..pos + 4]);
        return Some(UtcDateTime::from_unix(i64::from(mtime_seconds)));
    }

    None
}

/// Parses Unix timestamp extra field (0x5855) - obsolete format
fn parse_unix_timestamp(data: &[u8]) -> Option<UtcDateTime> {
    if data.len() < 8 {
        return None;
    }

    // Unix format has access time first, then modification time
    let mtime_seconds = le_u32(&data[4..8]);
    Some(UtcDateTime::from_unix(i64::from(mtime_seconds)))
}

/// Convert Unix timestamp to broken down date/time components
///
/// Based on Howard Hinnant's date library algorithm `civil_from_days`:
///
/// <https://howardhinnant.github.io/date_algorithms.html#civil_from_days>
fn unix_timestamp_to_components(timestamp: i64) -> (u16, u8, u8, u8, u8, u8) {
    const SECONDS_PER_DAY: i64 = 86400;

    // Break timestamp into days and seconds within day
    let total_days = timestamp / SECONDS_PER_DAY;
    let mut seconds_in_day = timestamp % SECONDS_PER_DAY;

    // Handle negative remainder for negative timestamps
    if seconds_in_day < 0 {
        seconds_in_day += SECONDS_PER_DAY;
    }

    // Convert seconds within day to H:M:S
    let hour = (seconds_in_day / 3600) as u8;
    let minute = ((seconds_in_day % 3600) / 60) as u8;
    let second = (seconds_in_day % 60) as u8;

    let days_since_epoch = total_days;

    // Shift epoch from 1970-01-01 to 0000-03-01 for easier leap year handling
    // This makes March 1st, year 0 our epoch (which aligns with leap year cycle)
    let days_since_shifted_epoch = days_since_epoch + 719468; // Days from 0000-03-01 to 1970-01-01

    // Calculate the era (400-year period)
    let era = days_since_shifted_epoch / 146097;
    let days_of_era = days_since_shifted_epoch % 146097;

    // Calculate year within the era (0-399)
    let year_of_era =
        (days_of_era - days_of_era / 1460 + days_of_era / 36524 - days_of_era / 146096) / 365;

    // Calculate the actual year
    let year = era * 400 + year_of_era;

    // Calculate day of year
    let days_before_year = year_of_era * 365 + year_of_era / 4 - year_of_era / 100;
    let day_of_year = days_of_era - days_before_year;

    // Calculate month and day
    // Months are shifted: Mar=0, Apr=1, ..., Dec=9, Jan=10, Feb=11
    let month_shifted = (5 * day_of_year + 2) / 153;
    let day_of_month = day_of_year - (153 * month_shifted + 2) / 5 + 1;

    // Convert back to normal calendar
    let (final_year, final_month) = if month_shifted < 10 {
        (year, month_shifted + 3)
    } else {
        (year + 1, month_shifted - 9)
    };

    (
        final_year as u16,
        final_month as u8,
        day_of_month as u8,
        hour,
        minute,
        second,
    )
}

// NTFS timestamp is 100-nanosecond intervals since 1601-01-01 00:00:00 UTC
const NTFS_EPOCH_OFFSET: u64 = 11644473600; // Seconds between 1601-01-01 and 1970-01-01

/// Returns true if the given year is a leap year.
const fn is_leap(year: u16) -> bool {
    year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
}

/// Returns the last valid day of the given month in the given year.
const fn last_day_of_month(year: u16, month: u8) -> u8 {
    if month != 2 || !is_leap(year) {
        last_day_of_month_common_year(month as usize)
    } else {
        29
    }
}

const fn last_day_of_month_common_year(m: usize) -> u8 {
    [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
}

#[cfg(test)]
mod tests {
    use super::*;

    fn utc_from_components(
        year: u16,
        month: u8,
        day: u8,
        hour: u8,
        minute: u8,
        second: u8,
        nanosecond: u32,
    ) -> UtcDateTime {
        UtcDateTime::from_components(year, month, day, hour, minute, second, nanosecond).unwrap()
    }

    fn local_from_components(
        year: u16,
        month: u8,
        day: u8,
        hour: u8,
        minute: u8,
        second: u8,
        nanosecond: u32,
    ) -> LocalDateTime {
        LocalDateTime::from_components(year, month, day, hour, minute, second, nanosecond).unwrap()
    }

    #[test]
    fn test_zip_to_dos_conversion() {
        // Test normal conversion
        let zip_dt = utc_from_components(2023, 6, 15, 14, 30, 45, 0);
        let dos_dt: DosDateTime = (&zip_dt).into();
        let (dos_time, dos_date) = dos_dt.into_parts();
        let dos_dt_check = DosDateTime::new(dos_time, dos_date);

        assert_eq!(dos_dt_check.year(), 2023);
        assert_eq!(dos_dt_check.month(), 6);
        assert_eq!(dos_dt_check.day(), 15);
        assert_eq!(dos_dt_check.hour(), 14);
        assert_eq!(dos_dt_check.minute(), 30);
        assert_eq!(dos_dt_check.second(), 44); // Rounded down to even second
    }

    #[test]
    fn test_zip_to_dos_year_saturation() {
        // Test year before DOS range (should saturate to 1980)
        let zip_dt_before = utc_from_components(1979, 6, 15, 14, 30, 45, 0);
        let dos_dt: DosDateTime = (&zip_dt_before).into();
        let (dos_time, dos_date) = dos_dt.into_parts();
        let dos_dt_check = DosDateTime::new(dos_time, dos_date);
        assert_eq!(dos_dt_check.year(), 1980); // Saturated to minimum
        assert_eq!(dos_dt_check.month(), 6);
        assert_eq!(dos_dt_check.day(), 15);

        // Test year way before DOS range
        let zip_dt_way_before = utc_from_components(1800, 1, 1, 0, 0, 0, 0);
        let dos_dt2: DosDateTime = (&zip_dt_way_before).into();
        let (dos_time2, dos_date2) = dos_dt2.into_parts();
        let dos_dt2_check = DosDateTime::new(dos_time2, dos_date2);
        assert_eq!(dos_dt2_check.year(), 1980); // Saturated to minimum

        // Test year after DOS range (should saturate to 2107)
        let zip_dt_after = utc_from_components(2108, 6, 15, 14, 30, 45, 0);
        let dos_dt3: DosDateTime = (&zip_dt_after).into();
        let (dos_time3, dos_date3) = dos_dt3.into_parts();
        let dos_dt3_check = DosDateTime::new(dos_time3, dos_date3);
        assert_eq!(dos_dt3_check.year(), 2107); // Saturated to maximum
        assert_eq!(dos_dt3_check.month(), 6);
        assert_eq!(dos_dt3_check.day(), 15);

        // Test year way after DOS range
        let zip_dt_way_after = utc_from_components(3000, 12, 31, 23, 59, 59, 0);
        let dos_dt4: DosDateTime = (&zip_dt_way_after).into();
        let (dos_time4, dos_date4) = dos_dt4.into_parts();
        let dos_dt4_check = DosDateTime::new(dos_time4, dos_date4);
        assert_eq!(dos_dt4_check.year(), 2107); // Saturated to maximum
    }

    #[test]
    fn test_dos_datetime() {
        // Test using the From trait
        let zip_dt = utc_from_components(2023, 6, 15, 14, 30, 45, 0);
        let dos_dt: DosDateTime = (&zip_dt).into();
        assert_eq!(dos_dt.year(), 2023);
        assert_eq!(dos_dt.month(), 6);
        assert_eq!(dos_dt.day(), 15);
        assert_eq!(dos_dt.hour(), 14);
        assert_eq!(dos_dt.minute(), 30);
        assert_eq!(dos_dt.second(), 44); // Rounded down to even second
    }

    #[test]
    fn test_dos_datetime_odd_seconds() {
        // Test that odd seconds are rounded down using the From trait
        let zip_dt_odd = utc_from_components(2020, 1, 1, 12, 30, 45, 0);
        let dos_dt_odd: DosDateTime = (&zip_dt_odd).into();
        assert_eq!(dos_dt_odd.second(), 44); // 45 rounded down to 44

        let zip_dt_even = utc_from_components(2020, 1, 1, 12, 30, 46, 0);
        let dos_dt_even: DosDateTime = (&zip_dt_even).into();
        assert_eq!(dos_dt_even.second(), 46); // 46 stays 46
    }

    #[test]
    fn test_dos_datetime_edge_cases() {
        // Test minimum date using From trait
        let zip_dt_min = utc_from_components(1980, 1, 1, 0, 0, 0, 0);
        let dos_dt_min: DosDateTime = (&zip_dt_min).into();
        assert_eq!(dos_dt_min.year(), 1980);
        assert_eq!(dos_dt_min.month(), 1);
        assert_eq!(dos_dt_min.day(), 1);

        // Test maximum date using From trait
        let zip_dt_max = utc_from_components(2107, 12, 31, 23, 59, 58, 0);
        let dos_dt_max: DosDateTime = (&zip_dt_max).into();
        assert_eq!(dos_dt_max.year(), 2107);
        assert_eq!(dos_dt_max.month(), 12);
        assert_eq!(dos_dt_max.day(), 31);
        assert_eq!(dos_dt_max.hour(), 23);
        assert_eq!(dos_dt_max.minute(), 59);
        assert_eq!(dos_dt_max.second(), 58);
    }

    #[test]
    fn test_dos_datetime_zero_normalization() {
        // Test that zero DOS timestamp (0x0000 0x0000) is normalized to 1980-01-01 00:00:00
        let datetime = DosDateTime::new(0x0000, 0x0000);
        assert_eq!(datetime.year(), 1980);
        assert_eq!(datetime.month(), 1); // month 0 normalized to 1
        assert_eq!(datetime.day(), 1); // day 0 normalized to 1
        assert_eq!(datetime.hour(), 0);
        assert_eq!(datetime.minute(), 0);
        assert_eq!(datetime.second(), 0);

        // Test partial zero normalization - only month is zero
        let datetime = DosDateTime::new(0x0000, 0x0001); // day=1, month=0, year=1980
        assert_eq!(datetime.year(), 1980);
        assert_eq!(datetime.month(), 1); // month 0 normalized to 1
        assert_eq!(datetime.day(), 1);
        assert_eq!(datetime.hour(), 0);
        assert_eq!(datetime.minute(), 0);
        assert_eq!(datetime.second(), 0);

        // Test partial zero normalization - only day is zero
        let datetime = DosDateTime::new(0x0000, 0x0020); // day=0, month=1, year=1980
        assert_eq!(datetime.year(), 1980);
        assert_eq!(datetime.month(), 1);
        assert_eq!(datetime.day(), 1); // day 0 normalized to 1
        assert_eq!(datetime.hour(), 0);
        assert_eq!(datetime.minute(), 0);
        assert_eq!(datetime.second(), 0);
    }

    #[test]
    fn test_zip_datetime_dos() {
        let datetime = local_from_components(2020, 6, 15, 14, 30, 44, 0);

        assert_eq!(datetime.year(), 2020);
        assert_eq!(datetime.month(), 6);
        assert_eq!(datetime.day(), 15);
        assert_eq!(datetime.hour(), 14);
        assert_eq!(datetime.minute(), 30);
        assert_eq!(datetime.second(), 44);
        assert_eq!(datetime.nanosecond(), 0);
        assert_eq!(datetime.timezone(), TimeZone::Local);
    }

    #[test]
    fn test_zip_datetime_unix() {
        // Unix timestamp for 2010-09-05 02:12:01 UTC
        let datetime = utc_from_components(2010, 9, 5, 2, 12, 1, 0);

        assert_eq!(datetime.year(), 2010);
        assert_eq!(datetime.month(), 9);
        assert_eq!(datetime.day(), 5);
        assert_eq!(datetime.hour(), 2);
        assert_eq!(datetime.minute(), 12);
        assert_eq!(datetime.second(), 1);
        assert_eq!(datetime.nanosecond(), 0);
        assert_eq!(datetime.timezone(), TimeZone::Utc);
    }

    #[test]
    fn test_zip_datetime_ntfs() {
        // NTFS timestamp for roughly 2010-09-05 02:12:01 UTC with 500ms precision
        let datetime = utc_from_components(2010, 9, 5, 2, 12, 1, 500000000);

        assert_eq!(datetime.year(), 2010);
        assert_eq!(datetime.month(), 9);
        assert_eq!(datetime.day(), 5);
        assert_eq!(datetime.hour(), 2);
        assert_eq!(datetime.minute(), 12);
        assert_eq!(datetime.second(), 1);
        assert_eq!(datetime.nanosecond(), 500000000);
        assert_eq!(datetime.timezone(), TimeZone::Utc);
    }

    #[test]
    fn test_to_unix_comprehensive() {
        // Test comprehensive cases including edge cases and leap years

        // Test first day of each month in a leap year (2020)
        let jan_1_2020 = utc_from_components(2020, 1, 1, 0, 0, 0, 0);
        assert_eq!(jan_1_2020.to_unix(), 1577836800);

        let feb_29_2020 = utc_from_components(2020, 2, 29, 0, 0, 0, 0);
        assert_eq!(feb_29_2020.to_unix(), 1582934400);

        let mar_1_2020 = utc_from_components(2020, 3, 1, 0, 0, 0, 0);
        assert_eq!(mar_1_2020.to_unix(), 1583020800);

        // Test non-leap year (2021)
        let feb_28_2021 = utc_from_components(2021, 2, 28, 0, 0, 0, 0);
        assert_eq!(feb_28_2021.to_unix(), 1614470400);

        // Test century boundary (non-leap year despite being divisible by 4)
        let mar_1_1900 = utc_from_components(1900, 3, 1, 0, 0, 0, 0);
        // This is before Unix epoch, so returns negative value
        let result = mar_1_1900.to_unix();
        assert!(result < 0); // Dates before epoch return negative values

        // Test year 2038 boundary (close to u32::MAX seconds)
        let early_2038 = utc_from_components(2038, 1, 1, 0, 0, 0, 0);
        let timestamp_2038 = early_2038.to_unix();
        assert!(timestamp_2038 > 0); // Should have a valid positive timestamp

        // Test far future dates (beyond u32 range but handled by i64)
        let far_future = utc_from_components(2200, 1, 1, 0, 0, 0, 0);
        let result = far_future.to_unix();
        // Should return a valid i64 timestamp for far future dates
        assert!(result > u32::MAX as i64); // Should exceed u32 range
    }

    #[test]
    fn test_to_unix_accuracy() {
        // Test known dates against their Unix timestamps (verified with Python datetime)

        // Unix epoch: 1970-01-01 00:00:00 UTC = 0
        let epoch = utc_from_components(1970, 1, 1, 0, 0, 0, 0);
        assert_eq!(epoch.to_unix(), 0);

        // 2000-01-01 00:00:00 UTC = 946684800
        let y2k = utc_from_components(2000, 1, 1, 0, 0, 0, 0);
        assert_eq!(y2k.to_unix(), 946684800);

        // 2023-06-15 14:30:45 UTC = 1686839445
        let test_date = utc_from_components(2023, 6, 15, 14, 30, 45, 0);
        assert_eq!(test_date.to_unix(), 1686839445);

        // Leap year test: 2020-02-29 12:00:00 UTC = 1582977600
        let leap_day = utc_from_components(2020, 2, 29, 12, 0, 0, 0);
        assert_eq!(leap_day.to_unix(), 1582977600);

        // Test dates before Unix epoch return negative values
        let before_epoch = utc_from_components(1969, 12, 31, 23, 59, 59, 0);
        let result = before_epoch.to_unix();
        // One second before epoch should be -1
        assert_eq!(result, -1);
    }

    #[test]
    fn test_negative_unix_timestamps() {
        // Test that negative timestamps (before 1970) work correctly
        let negative_timestamp = -86400; // One day before epoch (1969-12-31)
        let datetime = UtcDateTime::from_unix(negative_timestamp);

        assert_eq!(datetime.year(), 1969);
        assert_eq!(datetime.month(), 12);
        assert_eq!(datetime.day(), 31);
        assert_eq!(datetime.hour(), 0);
        assert_eq!(datetime.minute(), 0);
        assert_eq!(datetime.second(), 0);

        // Round trip test
        assert_eq!(datetime.to_unix(), negative_timestamp);
    }

    #[test]
    fn test_days_from_civil() {
        // Test Unix epoch
        let epoch = utc_from_components(1970, 1, 1, 0, 0, 0, 0);
        assert_eq!(epoch.days_from_civil(), 0);

        // Test Y2K (verified with Python)
        let y2k = utc_from_components(2000, 1, 1, 0, 0, 0, 0);
        assert_eq!(y2k.days_from_civil(), 10957);

        // Test leap year boundary (verified with Python)
        let leap_day = utc_from_components(2020, 2, 29, 0, 0, 0, 0);
        assert_eq!(leap_day.days_from_civil(), 18321);

        // Test before epoch (negative value)
        let before_epoch = utc_from_components(1969, 12, 31, 0, 0, 0, 0);
        assert_eq!(before_epoch.days_from_civil(), -1);
    }

    #[test]
    fn test_zip_datetime_display() {
        // Test with zero nanoseconds - should omit the nanosecond part
        let datetime_no_nanos = utc_from_components(2023, 6, 15, 14, 30, 42, 0);
        assert_eq!(format!("{}", datetime_no_nanos), "2023-06-15T14:30:42Z");

        // Test with non-zero nanoseconds - should include the nanosecond part
        let datetime_with_nanos = utc_from_components(2023, 6, 15, 14, 30, 42, 500000000);
        assert_eq!(
            format!("{}", datetime_with_nanos),
            "2023-06-15T14:30:42.500000000Z"
        );

        // Test local time with zero nanoseconds
        let datetime_local = local_from_components(2023, 6, 15, 14, 30, 42, 0);
        assert_eq!(format!("{}", datetime_local), "2023-06-15T14:30:42");

        // Test local time with nanoseconds
        let datetime_local_nanos = local_from_components(2023, 6, 15, 14, 30, 42, 123456789);
        assert_eq!(
            format!("{}", datetime_local_nanos),
            "2023-06-15T14:30:42.123456789"
        );
    }

    #[test]
    fn test_parse_extended_timestamp() {
        // Extended timestamp with modification time flag and Unix timestamp
        let mut data = vec![0x01]; // Flags: modification time present
        data.extend_from_slice(&1283652721u32.to_le_bytes()); // Unix timestamp

        let result = parse_extended_timestamp(&data).unwrap();
        // Check that it's a Unix timestamp with the right components
        assert_eq!(result.year(), 2010);
        assert_eq!(result.month(), 9);
        assert_eq!(result.day(), 5);
        assert_eq!(result.hour(), 2);
        assert_eq!(result.minute(), 12);
        assert_eq!(result.second(), 1);
        assert_eq!(result.timezone(), TimeZone::Utc);
    }

    #[test]
    fn test_parse_unix_timestamp() {
        // Unix timestamp format: access time (4 bytes) + modification time (4 bytes)
        let mut data = vec![];
        data.extend_from_slice(&0u32.to_le_bytes()); // Access time (ignored)
        data.extend_from_slice(&1283652721u32.to_le_bytes()); // Modification time

        let result = parse_unix_timestamp(&data).unwrap();
        // Check that it's a Unix timestamp with the right components
        assert_eq!(result.year(), 2010);
        assert_eq!(result.month(), 9);
        assert_eq!(result.day(), 5);
        assert_eq!(result.hour(), 2);
        assert_eq!(result.minute(), 12);
        assert_eq!(result.second(), 1);
        assert_eq!(result.timezone(), TimeZone::Utc);
    }

    #[test]
    fn test_parse_ntfs_timestamp() {
        // NTFS timestamp format
        let mut data = vec![0; 4]; // Reserved
        data.extend_from_slice(&0x0001u16.to_le_bytes()); // Tag
        data.extend_from_slice(&24u16.to_le_bytes()); // Size

        // NTFS timestamp (100-nanosecond ticks since 1601-01-01)
        let ticks = (1283652721 + NTFS_EPOCH_OFFSET) * 10_000_000;
        data.extend_from_slice(&ticks.to_le_bytes()); // Modification time
        data.extend_from_slice(&0u64.to_le_bytes()); // Access time
        data.extend_from_slice(&0u64.to_le_bytes()); // Creation time

        let result = parse_ntfs_timestamp(&data).unwrap();
        // Check that it's an NTFS timestamp with the right components
        assert_eq!(result.year(), 2010);
        assert_eq!(result.month(), 9);
        assert_eq!(result.day(), 5);
        assert_eq!(result.hour(), 2);
        assert_eq!(result.minute(), 12);
        assert_eq!(result.second(), 1);
        assert_eq!(result.timezone(), TimeZone::Utc);
    }

    #[test]
    fn test_zip_datetime_ordering() {
        let dt1 = UtcDateTime::from_components(2020, 1, 1, 0, 0, 0, 0).unwrap();
        let dt2 = UtcDateTime::from_components(2020, 1, 1, 0, 0, 0, 500_000_000).unwrap(); // Same time, more nanoseconds
        let dt3 = UtcDateTime::from_components(2020, 1, 1, 0, 0, 1, 0).unwrap(); // One second later
        let dt4 = UtcDateTime::from_components(2020, 1, 1, 0, 1, 0, 0).unwrap(); // One minute later
        let dt5 = UtcDateTime::from_components(2020, 1, 1, 1, 0, 0, 0).unwrap(); // One hour later
        let dt6 = UtcDateTime::from_components(2020, 1, 2, 0, 0, 0, 0).unwrap(); // One day later
        let dt7 = UtcDateTime::from_components(2020, 2, 1, 0, 0, 0, 0).unwrap(); // One month later
        let dt8 = UtcDateTime::from_components(2021, 1, 1, 0, 0, 0, 0).unwrap(); // One year later

        let mut timestamps = vec![dt8, dt3, dt1, dt6, dt4, dt2, dt7, dt5];
        timestamps.sort_unstable();
        let expected = vec![dt1, dt2, dt3, dt4, dt5, dt6, dt7, dt8];
        assert_eq!(
            timestamps, expected,
            "sorting should produce chronological order"
        );
    }
}

#[cfg(test)]
mod property_tests {
    //! Property-based tests to verify timestamp conversion accuracy against jiff.

    use super::*;
    use quickcheck_macros::quickcheck;

    #[quickcheck]
    fn prop_unix_timestamp_conversion(unix_seconds: u32) {
        let zip_datetime = UtcDateTime::from_unix(i64::from(unix_seconds));

        let Ok(timestamp) = jiff::Timestamp::from_second(unix_seconds as i64) else {
            return;
        };

        let dt = timestamp.to_zoned(jiff::tz::TimeZone::UTC);

        assert_eq!(zip_datetime.year(), dt.year() as u16, "year");
        assert_eq!(zip_datetime.month(), dt.month() as u8, "month");
        assert_eq!(zip_datetime.day(), dt.day() as u8, "day");
        assert_eq!(zip_datetime.hour(), dt.hour() as u8, "hour");
        assert_eq!(zip_datetime.minute(), dt.minute() as u8, "minute");
        assert_eq!(zip_datetime.second(), dt.second() as u8, "second");
        assert_eq!(zip_datetime.timezone(), TimeZone::Utc);
        assert_eq!(zip_datetime.nanosecond(), 0, "nanosecond");

        assert_eq!(
            zip_datetime.to_unix(),
            i64::from(unix_seconds),
            "to_unix should match input"
        );
    }

    /// Property test: NTFS timestamp conversion should match jiff's conversion
    #[quickcheck]
    fn prop_ntfs_timestamp_conversion(ntfs_ticks: u64) {
        let zip_datetime = UtcDateTime::from_ntfs(ntfs_ticks);

        // Convert NTFS ticks to Unix timestamp for jiff
        // NTFS ticks are 100-nanosecond intervals since 1601-01-01
        let unix_seconds = (ntfs_ticks / 10_000_000).saturating_sub(NTFS_EPOCH_OFFSET);
        let nanoseconds = ((ntfs_ticks % 10_000_000) * 100) as u32;

        if unix_seconds > u32::MAX as u64 {
            return;
        }

        let Ok(jiff_timestamp) = jiff::Timestamp::new(unix_seconds as i64, nanoseconds as i32)
        else {
            return;
        };

        let dt = jiff_timestamp.to_zoned(jiff::tz::TimeZone::UTC);

        assert_eq!(zip_datetime.year(), dt.year() as u16, "year");
        assert_eq!(zip_datetime.month(), dt.month() as u8, "month");
        assert_eq!(zip_datetime.day(), dt.day() as u8, "day");
        assert_eq!(zip_datetime.hour(), dt.hour() as u8, "hour");
        assert_eq!(zip_datetime.minute(), dt.minute() as u8, "minute");
        assert_eq!(zip_datetime.second(), dt.second() as u8, "second");
        assert_eq!(zip_datetime.timezone(), TimeZone::Utc);
        assert_eq!(zip_datetime.nanosecond(), nanoseconds, "nanosecond");
    }

    /// Property test: DOS timestamp conversion should always produce valid jiff datetimes
    #[quickcheck]
    fn prop_dos_timestamp_always_valid(dos_time: u16, dos_date: u16) {
        let dos_datetime = DosDateTime::new(dos_time, dos_date);
        let zip_datetime = LocalDateTime::from_dos(dos_datetime);

        // Create jiff datetime - this should never fail with our normalization
        let dt = jiff::civil::DateTime::new(
            zip_datetime.year() as i16,
            zip_datetime.month() as i8,
            zip_datetime.day() as i8,
            zip_datetime.hour() as i8,
            zip_datetime.minute() as i8,
            zip_datetime.second() as i8,
            0, // nanosecond
        )
        .unwrap();

        // Verify the components match what we expect
        assert_eq!(zip_datetime.year(), dt.year() as u16, "year");
        assert_eq!(zip_datetime.month(), dt.month() as u8, "month");
        assert_eq!(zip_datetime.day(), dt.day() as u8, "day");
        assert_eq!(zip_datetime.hour(), dt.hour() as u8, "hour");
        assert_eq!(zip_datetime.minute(), dt.minute() as u8, "minute");
        assert_eq!(zip_datetime.second(), dt.second() as u8, "second");
    }
}
