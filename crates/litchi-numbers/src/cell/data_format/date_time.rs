//! Validated date-and-time pattern values.

use std::fmt;

/// Maximum UTF-8 size of one date-and-time pattern.
pub const MAX_PATTERN_BYTES: usize = 4 * 1_024;

const ISO_DATE_PATTERN: &str = "yyyy-MM-dd";
const TIME_24_HOUR_SECONDS_PATTERN: &str = "H:mm:ss";
const ISO_DATE_TIME_24_HOUR_SECONDS_PATTERN: &str = "yyyy-MM-dd H:mm:ss";

/// Errors returned by date-and-time pattern construction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// The pattern is empty or contains only whitespace.
    Empty,
    /// The pattern exceeds [`MAX_PATTERN_BYTES`].
    TooLong { length: usize, maximum: usize },
    /// The pattern contains a NUL byte.
    ContainsNul,
}
impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => formatter.write_str("date-and-time pattern cannot be empty"),
            Self::TooLong { length, maximum } => {
                write!(
                    formatter,
                    "date-and-time pattern is {length} bytes; maximum is {maximum}"
                )
            },
            Self::ContainsNul => formatter.write_str("date-and-time pattern cannot contain NUL"),
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked date-and-time constructors.
pub type Result<T> = std::result::Result<T, Error>;

/// A bounded date-and-time pattern.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct DateTime(Box<str>);

impl DateTime {
    /// Validates and stores a borrowed pattern.
    ///
    /// Validation happens before the semantic value allocates.
    ///
    /// # Errors
    ///
    /// Returns a typed error for empty, oversized, or NUL-containing input.
    pub fn new(pattern: &str) -> Result<Self> {
        validate(pattern)?;
        Ok(Self(pattern.into()))
    }

    /// Validates and stores an already-owned pattern without an extra copy.
    ///
    /// # Errors
    ///
    /// Returns a typed error for empty, oversized, or NUL-containing input.
    pub fn from_owned(pattern: String) -> Result<Self> {
        validate(&pattern)?;
        Ok(Self(pattern.into_boxed_str()))
    }

    /// Validates and stores an already-boxed pattern without an extra copy.
    ///
    /// # Errors
    ///
    /// Returns a typed error for empty, oversized, or NUL-containing input.
    pub fn from_boxed(pattern: Box<str>) -> Result<Self> {
        validate(&pattern)?;
        Ok(Self(pattern))
    }

    /// The ISO date-only preset.
    #[must_use]
    pub fn iso_date() -> Self {
        Self(ISO_DATE_PATTERN.into())
    }

    /// The 24-hour time-only preset with seconds.
    #[must_use]
    pub fn time_24_hour_with_seconds() -> Self {
        Self(TIME_24_HOUR_SECONDS_PATTERN.into())
    }

    /// The ISO date-and-time preset with 24-hour seconds.
    #[must_use]
    pub fn iso_date_time_24_hour_with_seconds() -> Self {
        Self(ISO_DATE_TIME_24_HOUR_SECONDS_PATTERN.into())
    }

    /// Borrows the exact pattern without allocating.
    #[must_use]
    pub fn pattern(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for DateTime {
    fn as_ref(&self) -> &str {
        self.pattern()
    }
}

impl TryFrom<&str> for DateTime {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<String> for DateTime {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::from_owned(value)
    }
}

impl TryFrom<Box<str>> for DateTime {
    type Error = Error;

    fn try_from(value: Box<str>) -> Result<Self> {
        Self::from_boxed(value)
    }
}

fn validate(pattern: &str) -> Result<()> {
    if pattern.trim().is_empty() {
        return Err(Error::Empty);
    }
    if pattern.len() > MAX_PATTERN_BYTES {
        return Err(Error::TooLong {
            length: pattern.len(),
            maximum: MAX_PATTERN_BYTES,
        });
    }
    if pattern.contains('\0') {
        return Err(Error::ContainsNul);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pattern_validation_rejects_malformed_values() {
        assert_eq!(DateTime::new("   "), Err(Error::Empty));
        assert_eq!(DateTime::new("yyyy\0MM"), Err(Error::ContainsNul));
        assert!(matches!(
            DateTime::new(&"x".repeat(MAX_PATTERN_BYTES + 1)),
            Err(Error::TooLong { .. })
        ));
    }

    #[test]
    fn presets_and_owned_values_round_trip_without_normalization() {
        let preset = DateTime::iso_date_time_24_hour_with_seconds();
        assert_eq!(preset.pattern(), "yyyy-MM-dd H:mm:ss");
        let Ok(value) = DateTime::from_owned("EEEE, MMMM d, y".to_owned()) else {
            panic!("valid date-and-time pattern should construct");
        };
        assert_eq!(value.pattern(), "EEEE, MMMM d, y");
        let Ok(round_trip) = DateTime::try_from(value.pattern()) else {
            panic!("borrowed pattern should reconstruct");
        };
        assert_eq!(round_trip, value);
    }
}
