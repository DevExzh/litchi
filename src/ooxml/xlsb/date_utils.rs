//! Date and time utilities for XLSB format
//!
//! This module provides utilities for converting between Rust date/time representations
//! and Excel's date serial number format.
//!
//! # Excel Date System
//!
//! Excel stores dates as floating-point numbers representing the number of days since
//! a base date:
//! - **1900 system**: Days since January 1, 1900 (default, but has a leap year bug)
//! - **1904 system**: Days since January 1, 1904 (used on Mac, correct)
//!
//! The fractional part represents the time of day (0.0 = midnight, 0.5 = noon).
//!
//! # Excel 1900 Leap Year Bug
//!
//! Excel incorrectly treats 1900 as a leap year for compatibility with Lotus 1-2-3.
//! Dates before March 1, 1900 are off by one day. This implementation maintains
//! compatibility with Excel's behavior.

/// Days between December 30, 1899 (Excel epoch) and January 1, 1970 (Unix epoch)
/// Excel serial 1 = December 31, 1899 (but displayed as January 1, 1900 due to the bug)
/// Excel serial 2 = January 1, 1900
const EXCEL_1900_TO_UNIX_EPOCH_DAYS: i64 = 25569;

/// Days between January 1, 1904 and January 1, 1970 (Unix epoch)
const EXCEL_1904_TO_UNIX_EPOCH_DAYS: i64 = 24107;

/// Seconds per day
const SECONDS_PER_DAY: f64 = 86400.0;

/// Convert Excel serial number to Unix timestamp (seconds since Unix epoch)
///
/// # Arguments
///
/// * `serial` - Excel serial number (days since epoch)
/// * `is_1904` - Whether to use the 1904 date system (Mac)
///
/// # Returns
///
/// Unix timestamp as f64 (seconds since January 1, 1970)
///
/// # Examples
///
/// ```
/// use litchi::ooxml::xlsb::date_utils::excel_serial_to_unix;
///
/// // January 1, 2000 at noon (1900 system)
/// let serial = 36526.5;
/// let timestamp = excel_serial_to_unix(serial, false);
/// assert!((timestamp - 946728000.0).abs() < 1.0);
/// ```
#[inline]
pub fn excel_serial_to_unix(serial: f64, is_1904: bool) -> f64 {
    let epoch_days = if is_1904 {
        EXCEL_1904_TO_UNIX_EPOCH_DAYS
    } else {
        EXCEL_1900_TO_UNIX_EPOCH_DAYS
    };

    // Convert serial to days since Unix epoch
    let days_since_unix = serial - epoch_days as f64;

    // Convert to seconds
    days_since_unix * SECONDS_PER_DAY
}

/// Convert Unix timestamp to Excel serial number
///
/// # Arguments
///
/// * `unix_timestamp` - Unix timestamp (seconds since January 1, 1970)
/// * `is_1904` - Whether to use the 1904 date system (Mac)
///
/// # Returns
///
/// Excel serial number (days since epoch)
///
/// # Examples
///
/// ```
/// use litchi::ooxml::xlsb::date_utils::unix_to_excel_serial;
///
/// // January 1, 2000 at noon
/// let timestamp = 946728000.0;
/// let serial = unix_to_excel_serial(timestamp, false);
/// assert!((serial - 36526.5).abs() < 0.001);
/// ```
#[inline]
pub fn unix_to_excel_serial(unix_timestamp: f64, is_1904: bool) -> f64 {
    let epoch_days = if is_1904 {
        EXCEL_1904_TO_UNIX_EPOCH_DAYS
    } else {
        EXCEL_1900_TO_UNIX_EPOCH_DAYS
    };

    // Convert to days since Unix epoch
    let days_since_unix = unix_timestamp / SECONDS_PER_DAY;

    // Convert to Excel serial
    days_since_unix + epoch_days as f64
}

/// Check if an Excel serial number represents a valid date
///
/// Valid dates in Excel:
/// - 1900 system: >= 1 (January 1, 1900) and < 2958466 (December 31, 9999)
/// - 1904 system: >= 0 (January 1, 1904) and < 2957003 (December 31, 9999)
#[inline]
pub fn is_valid_excel_date(serial: f64, is_1904: bool) -> bool {
    if is_1904 {
        (0.0..2_957_003.0).contains(&serial)
    } else {
        (1.0..2_958_466.0).contains(&serial)
    }
}

/// Extract the date part (integer) from an Excel serial number
#[inline]
pub fn date_part(serial: f64) -> i64 {
    serial.floor() as i64
}

/// Extract the time part (fractional) from an Excel serial number
#[inline]
pub fn time_part(serial: f64) -> f64 {
    serial.fract()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_excel_1900_epoch() {
        // Excel serial 1 = January 1, 1900 (with the leap year bug)
        let serial = 1.0;
        let unix = excel_serial_to_unix(serial, false);
        // January 1, 1900 00:00:00 = -2209161600 seconds
        assert!((unix - (-2_209_075_200.0)).abs() < 1.0);
    }

    #[test]
    fn test_excel_1904_epoch() {
        // Excel serial 0 = January 1, 1904
        let serial = 0.0;
        let unix = excel_serial_to_unix(serial, true);
        // January 1, 1904 00:00:00 = -2082844800 seconds
        assert!((unix - (-2_082_844_800.0)).abs() < 1.0);
    }

    #[test]
    fn test_unix_epoch() {
        // Unix epoch = January 1, 1970 00:00:00
        let unix = 0.0;
        let serial_1900 = unix_to_excel_serial(unix, false);
        let serial_1904 = unix_to_excel_serial(unix, true);

        assert!((serial_1900 - 25569.0).abs() < 0.001);
        assert!((serial_1904 - 24107.0).abs() < 0.001);
    }

    #[test]
    fn test_roundtrip_1900() {
        let original = 44562.5; // Some date with time
        let unix = excel_serial_to_unix(original, false);
        let converted = unix_to_excel_serial(unix, false);
        assert!((original - converted).abs() < 0.001);
    }

    #[test]
    fn test_roundtrip_1904() {
        let original = 43100.25;
        let unix = excel_serial_to_unix(original, true);
        let converted = unix_to_excel_serial(unix, true);
        assert!((original - converted).abs() < 0.001);
    }

    #[test]
    fn test_date_time_parts() {
        let serial = 44562.75; // Date with 18:00 (0.75 of day)
        assert_eq!(date_part(serial), 44562);
        assert!((time_part(serial) - 0.75).abs() < 0.001);
    }

    #[test]
    fn test_valid_dates() {
        assert!(is_valid_excel_date(1.0, false)); // Valid 1900 date
        assert!(is_valid_excel_date(0.0, true)); // Valid 1904 date
        assert!(!is_valid_excel_date(0.0, false)); // Invalid 1900 date
        assert!(!is_valid_excel_date(-1.0, true)); // Invalid 1904 date
    }
}
