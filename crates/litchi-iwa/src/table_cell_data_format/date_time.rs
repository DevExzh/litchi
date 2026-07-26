//! Strongly typed settings for native Date & Time table cells.

use super::TableCellDataFormat;
use crate::{Error, Result};

const MAXIMUM_PATTERN_BYTES: usize = 4 * 1_024;
const ISO_DATE_PATTERN: &str = "yyyy-MM-dd";
const TIME_24_HOUR_SECONDS_PATTERN: &str = "H:mm:ss";
const ISO_DATE_TIME_24_HOUR_SECONDS_PATTERN: &str = "yyyy-MM-dd H:mm:ss";

/// Validated ICU-style pattern used by iWork's Date & Time cell format.
///
/// Native files may contain preset or custom ICU patterns, so this type
/// preserves the exact pattern instead of narrowing imported files to one
/// locale's inspector menu.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TableCellDateTimeFormat(Box<str>);

impl TableCellDateTimeFormat {
    /// Validate and preserve an ICU-style date/time pattern.
    pub fn new(pattern: impl Into<Box<str>>) -> Result<Self> {
        let pattern = pattern.into();
        if pattern.trim().is_empty() {
            return Err(Error::InvalidFormat(
                "table-cell Date & Time pattern cannot be empty".to_owned(),
            ));
        }
        if pattern.len() > MAXIMUM_PATTERN_BYTES {
            return Err(Error::InvalidFormat(format!(
                "table-cell Date & Time pattern exceeds {MAXIMUM_PATTERN_BYTES} bytes"
            )));
        }
        if pattern.contains('\0') {
            return Err(Error::InvalidFormat(
                "table-cell Date & Time pattern cannot contain NUL".to_owned(),
            ));
        }
        Ok(Self(pattern))
    }

    /// Construct the native `yyyy-MM-dd` date-only preset.
    pub fn iso_date() -> Self {
        Self(ISO_DATE_PATTERN.into())
    }

    /// Construct the native `H:mm:ss` time-only preset.
    pub fn time_24_hour_with_seconds() -> Self {
        Self(TIME_24_HOUR_SECONDS_PATTERN.into())
    }

    /// Construct the native `yyyy-MM-dd H:mm:ss` date-and-time preset.
    pub fn iso_date_time_24_hour_with_seconds() -> Self {
        Self(ISO_DATE_TIME_24_HOUR_SECONDS_PATTERN.into())
    }

    /// Borrow the exact ICU-style pattern without allocation.
    pub fn pattern(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for TableCellDateTimeFormat {
    fn as_ref(&self) -> &str {
        self.pattern()
    }
}

impl From<TableCellDateTimeFormat> for TableCellDataFormat {
    fn from(value: TableCellDateTimeFormat) -> Self {
        Self::DateTime(value)
    }
}
