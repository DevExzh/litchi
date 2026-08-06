//! Archive-free Numbers table header and footer semantics.
//!
//! Native presence, protobuf fields, table bounds, and transactional
//! publication remain owned by the concrete IWA adapter. This module only
//! models the compact values that callers use to describe table sections.

use std::fmt;
use std::num::NonZeroU8;

const MAX_COUNT: u8 = 5;

/// A non-zero table header, footer, or repeating-axis count.
///
/// A missing field in [`Settings`] represents zero. Keeping zero out of this
/// type makes accidental native-sentinel writes impossible and lets the
/// semantic value remain one byte wide.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Count(NonZeroU8);

impl Count {
    /// One header, footer, or repeating-axis row or column.
    pub const ONE: Self = Self(NonZeroU8::MIN);
    /// Two header, footer, or repeating-axis rows or columns.
    pub const TWO: Self = Self(NonZeroU8::new(2).expect("two is non-zero"));
    /// Three header, footer, or repeating-axis rows or columns.
    pub const THREE: Self = Self(NonZeroU8::new(3).expect("three is non-zero"));
    /// Four header, footer, or repeating-axis rows or columns.
    pub const FOUR: Self = Self(NonZeroU8::new(4).expect("four is non-zero"));
    /// Five header, footer, or repeating-axis rows or columns.
    pub const FIVE: Self = Self(NonZeroU8::new(5).expect("five is non-zero"));

    /// Validates and constructs a non-zero native count.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCount`] when `count` is zero or exceeds the
    /// native maximum of five.
    pub fn new(count: usize) -> Result<Self, Error> {
        let Ok(raw) = u8::try_from(count) else {
            return Err(Error::InvalidCount { count });
        };
        let Some(non_zero) = NonZeroU8::new(raw) else {
            return Err(Error::InvalidCount { count: 0 });
        };
        if non_zero.get() > MAX_COUNT {
            return Err(Error::InvalidCount {
                count: usize::from(non_zero.get()),
            });
        }
        Ok(Self(non_zero))
    }

    /// Returns the non-zero count as a platform-sized integer.
    #[must_use]
    pub const fn get(self) -> usize {
        self.0.get() as usize
    }
}

impl TryFrom<usize> for Count {
    type Error = Error;

    fn try_from(count: usize) -> Result<Self, Self::Error> {
        Self::new(count)
    }
}

impl From<Count> for usize {
    fn from(count: Count) -> Self {
        count.get()
    }
}

/// Lossless optional header, footer, and repeating-axis settings.
///
/// `None` preserves native field absence. In particular, it is not silently
/// normalized to an explicit false or zero during a read-modify-write cycle.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Settings {
    /// Number of leading header rows, if explicitly stored.
    pub header_rows: Option<Count>,
    /// Number of leading header columns, if explicitly stored.
    pub header_columns: Option<Count>,
    /// Number of trailing footer rows, if explicitly stored.
    pub footer_rows: Option<Count>,
    /// Whether header rows are explicitly frozen.
    pub header_rows_frozen: Option<bool>,
    /// Whether header columns are explicitly frozen.
    pub header_columns_frozen: Option<bool>,
    /// Whether header rows are explicitly repeated when printing.
    pub repeating_header_rows_enabled: Option<bool>,
    /// Whether header columns are explicitly repeated when printing.
    pub repeating_header_columns_enabled: Option<bool>,
}

impl Settings {
    /// Returns the effective number of header rows, treating absence as zero.
    #[must_use]
    pub fn header_row_count(self) -> usize {
        self.header_rows.map_or(0, Count::get)
    }

    /// Returns the effective number of header columns, treating absence as
    /// zero.
    #[must_use]
    pub fn header_column_count(self) -> usize {
        self.header_columns.map_or(0, Count::get)
    }

    /// Returns the effective number of footer rows, treating absence as zero.
    #[must_use]
    pub fn footer_row_count(self) -> usize {
        self.footer_rows.map_or(0, Count::get)
    }

    /// Returns whether header rows are effectively frozen.
    #[must_use]
    pub fn header_rows_are_frozen(self) -> bool {
        self.header_rows_frozen.unwrap_or(false)
    }

    /// Returns whether header columns are effectively frozen.
    #[must_use]
    pub fn header_columns_are_frozen(self) -> bool {
        self.header_columns_frozen.unwrap_or(false)
    }

    /// Returns whether header rows effectively repeat when printing.
    #[must_use]
    pub fn repeats_header_rows(self) -> bool {
        self.repeating_header_rows_enabled.unwrap_or(false)
    }

    /// Returns whether header columns effectively repeat when printing.
    #[must_use]
    pub fn repeats_header_columns(self) -> bool {
        self.repeating_header_columns_enabled.unwrap_or(false)
    }
}

/// Errors returned while constructing header and footer values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A count was outside the native range `1..=5`.
    InvalidCount {
        /// Rejected count.
        count: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidCount { count } => {
                write!(
                    formatter,
                    "table header or footer count {count} is outside 1..=5"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    #[test]
    fn count_is_compact_and_checked() {
        assert_eq!(size_of::<Count>(), 1);
        assert_eq!(size_of::<Option<Count>>(), 1);
        assert!(Count::new(0).is_err());
        for count in 1..=5 {
            assert_eq!(Count::new(count).unwrap().get(), count);
        }
        assert!(Count::new(6).is_err());
        assert!(Count::new(usize::MAX).is_err());
    }

    #[test]
    fn settings_preserve_presence_and_expose_effective_values() {
        let settings = Settings {
            header_rows: Some(Count::TWO),
            header_columns: None,
            footer_rows: Some(Count::ONE),
            header_rows_frozen: Some(true),
            ..Settings::default()
        };
        assert_eq!(settings.header_row_count(), 2);
        assert_eq!(settings.header_column_count(), 0);
        assert_eq!(settings.footer_row_count(), 1);
        assert!(settings.header_rows_are_frozen());
        assert!(!settings.header_columns_are_frozen());
        assert_eq!(settings.header_columns, None);
    }
}
